from __future__ import annotations

import io
import re
import ssl
import urllib.request
from datetime import date
from pathlib import Path

import pandas as pd


BASE_URL = "https://rosstat.gov.ru/storage/mediabank"
FIRST_YEAR = 2020

PROJECT_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = PROJECT_DIR / "data"
OUTPUT_FILE = DATA_DIR / "russia_cpi_monthly.csv"
SOURCE_INFO_FILE = DATA_DIR / "russia_cpi_source.txt"

MONTHS_RU = {
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


def month_candidates(months_back: int = 18) -> list[tuple[int, int]]:
    """Возвращает текущий месяц и предыдущие месяцы."""
    today = date.today()
    year = today.year
    month = today.month

    result: list[tuple[int, int]] = []

    for _ in range(months_back):
        result.append((year, month))

        month -= 1
        if month == 0:
            month = 12
            year -= 1

    return result


def download_file(url: str) -> bytes:
    """
    Скачивает файл.

    Контекст без проверки сертификата нужен из-за проблем
    с SSL-сертификатом сервера Росстата.
    """
    ssl_context = ssl.create_default_context()
    ssl_context.check_hostname = False
    ssl_context.verify_mode = ssl.CERT_NONE

    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 Chrome/126 Safari/537.36"
            ),
            "Accept": (
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet,application/octet-stream,*/*"
            ),
        },
    )

    with urllib.request.urlopen(
        request,
        context=ssl_context,
        timeout=60,
    ) as response:
        content = response.read()

    # XLSX является ZIP-архивом и обычно начинается с PK.
    if not content.startswith(b"PK"):
        raise ValueError("Сервер вернул не XLSX-файл")

    return content


def find_latest_rosstat_file() -> tuple[str, bytes]:
    """Перебирает последние месяцы и находит свежий существующий файл."""
    errors: list[str] = []

    for year, month in month_candidates():
        filename = f"ipc_mes_{month:02d}-{year}.xlsx"
        url = f"{BASE_URL}/{filename}"

        print(f"Проверяю: {url}")

        try:
            content = download_file(url)
            print(f"Найден файл: {filename}")
            return url, content
        except Exception as exc:
            errors.append(f"{filename}: {exc}")

    raise RuntimeError(
        "Не удалось найти актуальный файл Росстата.\n"
        + "\n".join(errors)
    )


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""

    text = str(value).strip().lower()
    text = text.replace("ё", "е")
    text = re.sub(r"\s+", " ", text)

    return text


def parse_year(value: object) -> int | None:
    """Извлекает год из ячейки."""
    if pd.isna(value):
        return None

    if isinstance(value, (int, float)):
        year = int(value)
        if 1990 <= year <= 2100:
            return year

    text = str(value).strip()
    match = re.search(r"\b(19|20)\d{2}\b", text)

    if match:
        return int(match.group(0))

    return None


def parse_number(value: object) -> float | None:
    """Преобразует 100,67 или 100.67 в число."""
    if pd.isna(value):
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    text = text.replace("\xa0", "")
    text = text.replace(" ", "")
    text = text.replace(",", ".")

    match = re.search(r"-?\d+(?:\.\d+)?", text)

    if not match:
        return None

    return float(match.group(0))


def find_target_sheet(
    workbook: dict[str, pd.DataFrame],
) -> tuple[str, pd.DataFrame]:
    """
    Ищет лист, на котором одновременно есть:
    - названия месяцев;
    - достаточно много годов.
    """
    best_sheet: str | None = None
    best_frame: pd.DataFrame | None = None
    best_score = -1

    for sheet_name, frame in workbook.items():
        years_count = 0
        months_count = 0

        for value in frame.to_numpy().ravel():
            if parse_year(value) is not None:
                years_count += 1

            if normalize_text(value) in MONTHS_RU:
                months_count += 1

        score = years_count + months_count * 10

        print(
            f"Лист {sheet_name!r}: "
            f"годы={years_count}, месяцы={months_count}"
        )

        if months_count >= 6 and score > best_score:
            best_sheet = sheet_name
            best_frame = frame
            best_score = score

    if best_sheet is None or best_frame is None:
        raise ValueError(
            "Не найден лист с месячными данными ИПЦ"
        )

    return best_sheet, best_frame


def find_year_header_row(frame: pd.DataFrame) -> int:
    """
    Находит строку, в которой расположено максимальное число годов.
    """
    best_row = -1
    best_year_count = 0

    for row_index in range(len(frame)):
        year_count = sum(
            parse_year(value) is not None
            for value in frame.iloc[row_index].tolist()
        )

        if year_count > best_year_count:
            best_year_count = year_count
            best_row = row_index

    if best_row < 0 or best_year_count < 3:
        raise ValueError("Не удалось найти строку с годами")

    print(
        f"Строка заголовков годов: {best_row}; "
        f"найдено годов: {best_year_count}"
    )

    return best_row


def find_month_column(
    frame: pd.DataFrame,
    start_row: int,
) -> int:
    """Находит столбец с январём–декабрём."""
    best_column = -1
    best_month_count = 0

    for column_index in range(frame.shape[1]):
        month_count = 0

        for value in frame.iloc[start_row:, column_index]:
            if normalize_text(value) in MONTHS_RU:
                month_count += 1

        if month_count > best_month_count:
            best_month_count = month_count
            best_column = column_index

    if best_column < 0 or best_month_count < 6:
        raise ValueError("Не найден столбец с названиями месяцев")

    print(
        f"Столбец месяцев: {best_column}; "
        f"найдено месяцев: {best_month_count}"
    )

    return best_column


def extract_cpi(frame: pd.DataFrame) -> pd.DataFrame:
    header_row = find_year_header_row(frame)
    month_column = find_month_column(frame, header_row + 1)

    year_columns: dict[int, int] = {}

    for column_index, value in enumerate(frame.iloc[header_row]):
        year = parse_year(value)

        if year is not None and year >= FIRST_YEAR:
            year_columns[year] = column_index

    if not year_columns:
        raise ValueError(
            f"В таблице не найдены годы начиная с {FIRST_YEAR}"
        )

    records: list[dict[str, object]] = []

    for row_index in range(header_row + 1, len(frame)):
        month_name = normalize_text(
            frame.iloc[row_index, month_column]
        )

        month_number = MONTHS_RU.get(month_name)

        if month_number is None:
            continue

        for year, column_index in year_columns.items():
            raw_value = frame.iloc[row_index, column_index]
            cpi_value = parse_number(raw_value)

            if cpi_value is None:
                continue

            records.append(
                {
                    "date": pd.Timestamp(
                        year=year,
                        month=month_number,
                        day=1,
                    )
                    + pd.offsets.MonthEnd(0),
                    "russia_cpi_mom_index": cpi_value,
                }
            )

    if not records:
        raise ValueError("Не удалось извлечь значения ИПЦ")

    result = pd.DataFrame(records)

    result = (
        result
        .drop_duplicates(subset=["date"], keep="last")
        .sort_values("date")
        .reset_index(drop=True)
    )

    # Не оставляем будущие месяцы и пустые периоды.
    current_month_end = (
        pd.Timestamp.today().normalize()
        + pd.offsets.MonthEnd(0)
    )

    result = result[result["date"] <= current_month_end].copy()

    result["date"] = result["date"].dt.strftime("%Y-%m-%d")
    result["russia_cpi_mom_index"] = (
        result["russia_cpi_mom_index"].round(2)
    )

    return result


def main() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    source_url, file_content = find_latest_rosstat_file()

    workbook = pd.read_excel(
        io.BytesIO(file_content),
        sheet_name=None,
        header=None,
        engine="openpyxl",
    )

    print("Листы в книге:", list(workbook.keys()))

    sheet_name, target_frame = find_target_sheet(workbook)
    print(f"Выбран лист: {sheet_name}")

    result = extract_cpi(target_frame)

    result.to_csv(
        OUTPUT_FILE,
        index=False,
        encoding="utf-8-sig",
    )

    SOURCE_INFO_FILE.write_text(
        f"source_url={source_url}\n"
        f"updated_at={pd.Timestamp.now().isoformat()}\n"
        f"rows={len(result)}\n",
        encoding="utf-8",
    )

    print()
    print("Готово.")
    print(f"Источник: {source_url}")
    print(f"Строк сохранено: {len(result)}")
    print(f"Последняя дата: {result.iloc[-1]['date']}")
    print(f"CSV: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()