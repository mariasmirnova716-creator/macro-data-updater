from __future__ import annotations

import io
import re
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin

import pandas as pd
import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook

import urllib3

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ============================================================
# PROJECT PATHS
# ============================================================

# Скрипт:
#   scripts/sheet_2_ipp_py/update_russia_ipp.py
#
# Данные:
#   data/sheet_2_ipp_data/

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parents[1]

DATA_DIR = PROJECT_ROOT / "data" / "sheet_2_ipp_data"

HISTORY_FILE = DATA_DIR / "russia_ipp_history_2020_2023.csv"
OLD_YOY_FILE = DATA_DIR / "ind_sub_2018_12-2025.xlsx"

OUTPUT_FILE = DATA_DIR / "russia_ipp_monthly.csv"
SOURCE_INFO_FILE = DATA_DIR / "russia_ipp_source.txt"

ROSSTAT_PAGE = "https://rosstat.gov.ru/enterprise_industrial"
EXPECTED_BASE_YEAR = 2023

MONTHS = {
    "январь": 1,
    "февраль": 2,
    "март": 3,
    "апрель": 4,
    "май": 5,
    "июнь": 6,
    "июль": 7,
    "август": 8,
    "сентябрь": 9,
    "октябрь": 10,
    "ноябрь": 11,
    "декабрь": 12,
}


# ============================================================
# HELPERS
# ============================================================

def clean_text(value) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").strip().lower()
    text = re.sub(r"\s+", " ", text)
    return text


def as_number(value):
    if value is None or value == "":
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip().replace(",", ".")
    text = re.sub(r"[^\d.\-]", "", text)

    if text in {"", "-", ".", "-."}:
        return None

    try:
        return float(text)
    except ValueError:
        return None


def detect_year(value) -> int | None:
    text = clean_text(value)
    match = re.search(r"(20\d{2})", text)
    return int(match.group(1)) if match else None


def month_end(year: int, month: int) -> pd.Timestamp:
    return pd.Timestamp(year=year, month=month, day=1) + pd.offsets.MonthEnd(0)


def find_rf_row(ws) -> int:
    for row in range(1, ws.max_row + 1):
        if clean_text(ws.cell(row, 1).value) == "российская федерация":
            return row
    raise RuntimeError(
        f"На листе '{ws.title}' не найдена строка 'Российская Федерация'."
    )

def resolve_history_file() -> Path:
    if HISTORY_FILE.exists():
        return HISTORY_FILE

    raise FileNotFoundError(
        "Не найден russia_ipp_history_2020_2023.csv.\n"
        f"Ожидалось:\n  {HISTORY_FILE}"
    )


# ============================================================
# ROSSTAT: FIND CURRENT 2023-BASE FILES
# ============================================================

def period_score(url: str) -> tuple[int, int]:
    """
    Для имен вроде:
      sezon_2023_06-2026.xlsx
      ind_sub_2023-06-2026.xlsx
    возвращает (2026, 6).
    """
    name = url.lower().split("/")[-1]
    match = re.search(r"(\d{1,2})-(20\d{2})\.xlsx(?:\?.*)?$", name)

    if match:
        month = int(match.group(1))
        year = int(match.group(2))
        return year, month

    return 0, 0


def find_latest_rosstat_files() -> tuple[str, str]:
    print("Opening Rosstat industrial production page:")
    print(ROSSTAT_PAGE)

    response = requests.get(
        ROSSTAT_PAGE,
        timeout=60,
        verify=False,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 Chrome/126 Safari/537.36"
            )
        },
    )
    response.raise_for_status()

    soup = BeautifulSoup(response.text, "html.parser")

    season_candidates = []
    ind_candidates = []

    for tag in soup.find_all("a", href=True):
        href = urljoin(ROSSTAT_PAGE, tag["href"])
        href_lower = href.lower()

        if not href_lower.endswith(".xlsx"):
            continue

        filename = href_lower.split("/")[-1]

        # Нам принципиально нужен новый блок с базисным 2023 годом.
        if re.search(r"sezon[_-]2023", filename):
            season_candidates.append(href)

        if re.search(r"ind_sub[_-]2023", filename):
            ind_candidates.append(href)

    if not season_candidates:
        raise RuntimeError(
            "На странице Росстата не найден XLSX вида sezon_2023_*.xlsx."
        )

    if not ind_candidates:
        raise RuntimeError(
            "На странице Росстата не найден XLSX вида ind_sub_2023-*.xlsx."
        )

    season_url = max(season_candidates, key=period_score)
    ind_url = max(ind_candidates, key=period_score)

    print("Latest seasonal file:")
    print(season_url)
    print("Latest indices file:")
    print(ind_url)

    return season_url, ind_url


def download_excel(url: str) -> bytes:
    response = requests.get(
        url,
        timeout=90,
        verify=False,
        headers={"User-Agent": "Mozilla/5.0"},
    )
    
    response.raise_for_status()

    content_type = response.headers.get("Content-Type", "").lower()
    if "html" in content_type and not response.content.startswith(b"PK"):
        raise RuntimeError(
            f"Вместо XLSX Росстат вернул HTML:\n{url}"
        )

    if not response.content.startswith(b"PK"):
        raise RuntimeError(
            f"Файл не похож на XLSX:\n{url}"
        )

    return response.content


# ============================================================
# PARSER 1:
# CURRENT SEASONAL FILE, BASE 2023
# ============================================================

def parse_current_seasonal(content: bytes) -> pd.DataFrame:
    """
    Файл вида sezon_2023_06-2026.xlsx, лист "1".

    Для общего индекса промышленного производства:
      B = к предыдущему периоду, факт
      C = к предыдущему периоду, SA
      D = к среднемесячному значению 2023, факт
      E = к среднемесячному значению 2023, SA
    """

    wb = load_workbook(
        io.BytesIO(content),
        data_only=True,
        read_only=True,
    )

    if "1" not in wb.sheetnames:
        raise RuntimeError(
            f"В seasonal-файле нет листа '1'. Есть: {wb.sheetnames}"
        )

    ws = wb["1"]

    # Проверка, что мы действительно получили новый базис 2023.
    header_text = " ".join(
        clean_text(ws.cell(r, c).value)
        for r in range(1, min(ws.max_row, 7) + 1)
        for c in range(1, min(ws.max_column, 6) + 1)
    )

    if (
        "среднемесячному значению 2023" not in header_text
        and "среднемесячному значению 2023г" not in header_text
    ):
        raise RuntimeError(
            "Seasonal-файл больше не содержит шкалу "
            f"'к среднемесячному значению {EXPECTED_BASE_YEAR} г.'. "
            "Возможно, Росстат сменил базисный год или структуру файла."
        )

    rows = []
    current_year = None

    for r in range(1, ws.max_row + 1):
        first = clean_text(ws.cell(r, 1).value)

        year = detect_year(ws.cell(r, 1).value)
        if year is not None and first not in MONTHS:
            current_year = year
            continue

        if first not in MONTHS or current_year is None:
            continue

        rows.append(
            {
                "date": month_end(current_year, MONTHS[first]),
                "ipp_mom": as_number(ws.cell(r, 2).value),
                "ipp_mom_sa": as_number(ws.cell(r, 3).value),
                "ipp_level_2023": as_number(ws.cell(r, 4).value),
                "ipp_level_2023_sa": as_number(ws.cell(r, 5).value),
            }
        )

    df = pd.DataFrame(rows)

    if df.empty:
        raise RuntimeError(
            "Не удалось извлечь месячные данные из seasonal-файла."
        )

    df = df.sort_values("date").reset_index(drop=True)

    # Новый официальный блок используем с 2024 года.
    df = df[df["date"].dt.year >= 2024].copy()

    if df.empty:
        raise RuntimeError(
            "Seasonal-файл не содержит данных начиная с 2024 года."
        )

    return df


# ============================================================
# PARSER 2:
# YoY FROM IND_SUB
# ============================================================

def parse_ind_sub_yoy(source, years: set[int] | None = None) -> pd.DataFrame:
    """
    Читает лист "1" файла ind_sub.

    Лист "1":
      "в % к соответствующему месяцу предыдущего года"

    Поддерживает:
      - старый большой ind_sub_2018_12-2025.xlsx
        с несколькими годами по горизонтали;
      - новый ind_sub_2023-06-2026.xlsx.
    """

    wb = load_workbook(
        source,
        data_only=True,
        read_only=True,
    )

    if "1" not in wb.sheetnames:
        raise RuntimeError(
            f"В ind_sub нет листа '1'. Есть: {wb.sheetnames}"
        )

    ws = wb["1"]

    title = clean_text(ws.cell(3, 1).value)

    if "соответствующему месяцу предыдущего года" not in title:
        raise RuntimeError(
            "Лист '1' ind_sub больше не является месячным YoY-рядом. "
            f"Заголовок: {ws.cell(3, 1).value!r}"
        )

    rf_row = find_rf_row(ws)

    rows = []
    current_year = None

    # Годы находятся в строке 4, месяцы в строке 5.
    for col in range(2, ws.max_column + 1):
        year_here = detect_year(ws.cell(4, col).value)
        if year_here is not None:
            current_year = year_here

        month_name = re.sub(r"\d+$", "", clean_text(ws.cell(5, col).value))

        if month_name not in MONTHS or current_year is None:
            continue

        if years is not None and current_year not in years:
            continue

        value = as_number(ws.cell(rf_row, col).value)

        if value is None:
            continue

        rows.append(
            {
                "date": month_end(current_year, MONTHS[month_name]),
                "ipp_yoy": value,
            }
        )

    df = pd.DataFrame(rows)

    if df.empty:
        raise RuntimeError(
            f"Не удалось извлечь YoY из файла {getattr(source, 'name', source)}."
        )

    return (
        df.sort_values("date")
        .drop_duplicates("date", keep="last")
        .reset_index(drop=True)
    )


# ============================================================
# VALIDATION
# ============================================================

def validate_result(df: pd.DataFrame):
    required = [
        "date",
        "ipp_yoy",
        "ipp_mom",
        "ipp_mom_sa",
        "ipp_level_2023",
        "ipp_level_2023_sa",
    ]

    missing_columns = [c for c in required if c not in df.columns]
    if missing_columns:
        raise RuntimeError(
            f"Нет обязательных колонок: {missing_columns}"
        )

    if df["date"].duplicated().any():
        duplicates = df.loc[
            df["date"].duplicated(keep=False), "date"
        ].dt.strftime("%Y-%m-%d").tolist()

        raise RuntimeError(
            f"Обнаружены повторяющиеся даты: {duplicates}"
        )

    df = df.sort_values("date").reset_index(drop=True)

    # История должна идти с января 2020.
    if df.iloc[0]["date"] != pd.Timestamp("2020-01-31"):
        raise RuntimeError(
            f"Первая дата неожиданная: {df.iloc[0]['date']}"
        )

    # Для официальной новой части никаких пропусков в пяти показателях.
    official = df[df["date"].dt.year >= 2024].copy()

    for col in required[1:]:
        missing = official[official[col].isna()]
        if not missing.empty:
            dates = missing["date"].dt.strftime("%Y-%m-%d").tolist()
            raise RuntimeError(
                f"В официальной части есть пропуски {col}: {dates}"
            )

    # Диапазоны — грубая защита от съехавшей структуры XLSX.
    for col in required[1:]:
        values = df[col].dropna()

        if ((values < 40) | (values > 180)).any():
            bad = df.loc[
                df[col].notna() & ((df[col] < 40) | (df[col] > 180)),
                ["date", col],
            ]
            raise RuntimeError(
                f"Подозрительные значения в {col}:\n{bad.to_string(index=False)}"
            )

    # Проверка базисного года для всей выходной таблицы.
    if "base_year" in df.columns:
        bad_base = df.loc[
            df["base_year"].notna()
            & (df["base_year"].astype(int) != EXPECTED_BASE_YEAR)
        ]
        if not bad_base.empty:
            raise RuntimeError(
                "В таблице обнаружен неожиданный base_year."
            )

    # Проверка непрерывности месяцев.
    expected_dates = pd.date_range(
        start=df["date"].min(),
        end=df["date"].max(),
        freq="ME",
    )

    actual_dates = pd.DatetimeIndex(df["date"])

    missing_dates = expected_dates.difference(actual_dates)
    if len(missing_dates):
        raise RuntimeError(
            "В ряду пропущены месяцы: "
            + ", ".join(d.strftime("%Y-%m-%d") for d in missing_dates)
        )

    # Последняя строка не должна быть слишком старой.
    latest = df["date"].max()
    today = pd.Timestamp.today().normalize()

    age_days = (today - latest).days
    if age_days > 120:
        raise RuntimeError(
            f"Последняя дата ИПП слишком старая: {latest.date()} "
            f"({age_days} дней назад)."
        )


# ============================================================
# MAIN
# ============================================================

def main():
    print("=" * 68)
    print("Russia industrial production updater — Sheet 2")
    print("=" * 68)

    history_path = resolve_history_file()

    if not OLD_YOY_FILE.exists():
        raise FileNotFoundError(
            "Для YoY 2024–2025 нужен зафиксированный файл:\n"
            f"  {OLD_YOY_FILE}\n"
            "Положи ind_sub_2018_12-2025.xlsx в data/sheet_2_ipp/."
        )

    print("\nLoading fixed history 2020–2023...")
    history = pd.read_csv(history_path, parse_dates=["date"])
    history = history[history["date"].dt.year <= 2023].copy()

    if len(history) != 48:
        raise RuntimeError(
            f"В seed 2020–2023 ожидалось 48 строк, получено {len(history)}."
        )

    print("Loading YoY 2024–2025 from fixed Rosstat workbook...")
    yoy_2024_2025 = parse_ind_sub_yoy(
        OLD_YOY_FILE,
        years={2024, 2025},
    )

    print("\nFinding current 2023-base Rosstat files...")
    season_url, ind_url = find_latest_rosstat_files()

    print("\nDownloading current seasonal XLSX...")
    season_content = download_excel(season_url)

    print("Downloading current ind_sub XLSX...")
    ind_content = download_excel(ind_url)

    print("\nParsing current seasonal data...")
    current = parse_current_seasonal(season_content)

    print("Parsing current YoY...")
    yoy_current = parse_ind_sub_yoy(
        io.BytesIO(ind_content),
        years=None,
    )

    # YoY: старый официальный 2024–2025 + текущий файл 2023-base.
    # Если когда-нибудь новый файл начнет включать 2024/2025,
    # он автоматически получит приоритет.
    yoy_all = pd.concat(
        [yoy_2024_2025, yoy_current],
        ignore_index=True,
    )

    yoy_all = (
        yoy_all
        .sort_values("date")
        .drop_duplicates("date", keep="last")
        .reset_index(drop=True)
    )

    current = current.merge(
        yoy_all,
        on="date",
        how="left",
        validate="one_to_one",
    )

    current["level_reconstructed"] = False
    current["base_year"] = EXPECTED_BASE_YEAR
    current["source_yoy"] = current["date"].apply(
        lambda d: (
            "ind_sub_2018_12-2025"
            if d.year in {2024, 2025}
            and d not in set(yoy_current["date"])
            else "ind_sub_2023_current"
        )
    )
    current["source_mom"] = "sezon_2023_current"
    current["source_mom_sa"] = "sezon_2023_current"
    current["source_level"] = "sezon_2023_current"
    current["source_level_sa"] = "sezon_2023_current"

    # У seed нет двух последних source-полей — добавим.
    if "source_level" not in history.columns:
        history["source_level"] = "reconstructed_from_mom"
    if "source_level_sa" not in history.columns:
        history["source_level_sa"] = "reconstructed_from_mom_sa"

    result = pd.concat(
        [history, current],
        ignore_index=True,
        sort=False,
    )

    result = (
        result
        .sort_values("date")
        .drop_duplicates("date", keep="last")
        .reset_index(drop=True)
    )

    # Округление: официальные исходники идут с 1 знаком,
    # reconstructed levels оставляем точнее.
    for col in ["ipp_yoy", "ipp_mom", "ipp_mom_sa"]:
        result[col] = result[col].round(1)

    official_mask = result["level_reconstructed"].eq(False)
    result.loc[official_mask, "ipp_level_2023"] = (
        result.loc[official_mask, "ipp_level_2023"].round(1)
    )
    result.loc[official_mask, "ipp_level_2023_sa"] = (
        result.loc[official_mask, "ipp_level_2023_sa"].round(1)
    )

    reconstructed_mask = result["level_reconstructed"].eq(True)
    result.loc[reconstructed_mask, "ipp_level_2023"] = (
        result.loc[reconstructed_mask, "ipp_level_2023"].round(3)
    )
    result.loc[reconstructed_mask, "ipp_level_2023_sa"] = (
        result.loc[reconstructed_mask, "ipp_level_2023_sa"].round(3)
    )

    validate_result(result)

    output_columns = [
        "date",
        "ipp_yoy",
        "ipp_mom",
        "ipp_mom_sa",
        "ipp_level_2023",
        "ipp_level_2023_sa",
        "level_reconstructed",
        "base_year",
        "source_yoy",
        "source_mom",
        "source_mom_sa",
        "source_level",
        "source_level_sa",
    ]

    result[output_columns].to_csv(
        OUTPUT_FILE,
        index=False,
        date_format="%Y-%m-%d",
        encoding="utf-8",
    )

    retrieved_at = datetime.now(timezone.utc).isoformat()

    SOURCE_INFO_FILE.write_text(
        "\n".join(
            [
                "Russia industrial production — Sheet 2",
                f"updated_utc={retrieved_at}",
                f"rosstat_page={ROSSTAT_PAGE}",
                f"current_seasonal_xlsx={season_url}",
                f"current_ind_sub_xlsx={ind_url}",
                f"fixed_yoy_2024_2025={OLD_YOY_FILE.name}",
                f"history_seed={history_path.name}",
                f"base_year={EXPECTED_BASE_YEAR}",
                "",
            ]
        ),
        encoding="utf-8",
    )

    print("\n" + "=" * 68)
    print("DONE")
    print("=" * 68)
    print(f"Output: {OUTPUT_FILE}")
    print(f"Source info: {SOURCE_INFO_FILE}")
    print(f"Rows: {len(result)}")
    print(f"First month: {result['date'].min().date()}")
    print(f"Latest month: {result['date'].max().date()}")
    print()
    print("Last 18 rows:")
    print(result[output_columns].tail(18).to_string(index=False))


if __name__ == "__main__":
    main()
