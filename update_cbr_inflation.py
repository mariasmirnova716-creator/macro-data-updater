from __future__ import annotations

import io
import re
import ssl
import urllib.parse
import urllib.request
from datetime import date
from pathlib import Path

import certifi
import pandas as pd


PAGE_URL = "https://www.cbr.ru/hd_base/infl/"
CBR_HOST = "https://www.cbr.ru"

FIRST_DATE = date(2020, 1, 1)

PROJECT_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = PROJECT_DIR / "data"

OUTPUT_FILE = DATA_DIR / "cbr_inflation_monthly.csv"
SOURCE_INFO_FILE = DATA_DIR / "cbr_inflation_source.txt"


def create_ssl_context() -> ssl.SSLContext:
    """Создаёт SSL-контекст с сертификатами certifi."""
    return ssl.create_default_context(cafile=certifi.where())


def format_date_ru(value: date) -> str:
    """Преобразует дату в формат ДД.ММ.ГГГГ для сайта ЦБ."""
    return value.strftime("%d.%m.%Y")


def build_page_url(start_date: date, end_date: date) -> str:
    """Создаёт URL страницы таблицы ЦБ с нужным периодом."""
    query = urllib.parse.urlencode(
        {
            "UniDbQuery.Posted": "True",
            "UniDbQuery.From": format_date_ru(start_date),
            "UniDbQuery.To": format_date_ru(end_date),
        }
    )

    return f"{PAGE_URL}?{query}"


def download_bytes(url: str) -> tuple[bytes, str]:
    """
    Загружает содержимое URL.

    Возвращает:
    - байты ответа;
    - итоговый URL после перенаправлений.
    """
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 Chrome/126 Safari/537.36"
            ),
            "Accept": "*/*",
            "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
        },
    )

    with urllib.request.urlopen(
        request,
        timeout=90,
        context=create_ssl_context(),
    ) as response:
        return response.read(), response.geturl()


def find_excel_url(page_html: str, page_url: str) -> str:
    """
    Находит в HTML страницы ссылку DownloadExcel.

    Ссылка на выгрузку может содержать меняющиеся параметры,
    поэтому не прописываем её полностью вручную.
    """
    normalized_html = (
        page_html
        .replace("&amp;", "&")
        .replace("&#x2F;", "/")
        .replace("\\u0026", "&")
    )

    patterns = [
        r'href=["\']([^"\']*?/Queries/UniDbQuery/DownloadExcel/[^"\']+)["\']',
        r'["\']([^"\']*?/Queries/UniDbQuery/DownloadExcel/[^"\']+)["\']',
        r'href=["\']([^"\']*?DownloadExcel[^"\']+)["\']',
    ]

    for pattern in patterns:
        match = re.search(
            pattern,
            normalized_html,
            flags=re.IGNORECASE,
        )

        if match:
            raw_url = match.group(1)

            return urllib.parse.urljoin(
                page_url,
                raw_url,
            )

    raise RuntimeError(
        "На странице ЦБ не найдена ссылка на скачивание Excel."
    )


def download_cbr_excel(
    start_date: date,
    end_date: date,
) -> tuple[bytes, str]:
    """Открывает страницу ЦБ и скачивает Excel-таблицу."""
    page_url = build_page_url(start_date, end_date)

    print(f"Открываю страницу ЦБ: {page_url}")

    page_bytes, final_page_url = download_bytes(page_url)

    page_html = page_bytes.decode(
        "utf-8",
        errors="replace",
    )

    try:
        excel_url = find_excel_url(
            page_html=page_html,
            page_url=final_page_url,
        )
    except RuntimeError:
        # Резервный вариант для текущего идентификатора таблицы ЦБ.
        # Основной способ выше предпочтительнее, поскольку автоматически
        # извлекает ссылку со страницы.
        fallback_query = urllib.parse.urlencode(
            {
                "From": format_date_ru(start_date),
                "FromDate": start_date.strftime("%m/%d/%Y"),
                "Posted": "True",
                "To": format_date_ru(end_date),
                "ToDate": end_date.strftime("%m/%d/%Y"),
            }
        )

        excel_url = (
            f"{CBR_HOST}/Queries/UniDbQuery/DownloadExcel/132934?"
            f"{fallback_query}"
        )

        print(
            "Ссылка не найдена в HTML. "
            "Использую резервный адрес выгрузки."
        )

    print(f"Скачиваю Excel: {excel_url}")

    excel_bytes, final_excel_url = download_bytes(excel_url)

    if not excel_bytes.startswith(b"PK"):
        beginning = excel_bytes[:200].decode(
            "utf-8",
            errors="replace",
        )

        raise RuntimeError(
            "ЦБ вернул не XLSX-файл. "
            f"Начало ответа: {beginning!r}"
        )

    return excel_bytes, final_excel_url


def normalize_text(value: object) -> str:
    """Нормализует текст для поиска заголовков."""
    if pd.isna(value):
        return ""

    text = str(value).strip().lower()
    text = text.replace("ё", "е")
    text = re.sub(r"\s+", " ", text)

    return text


def find_data_sheet(
    workbook: dict[str, pd.DataFrame],
) -> tuple[str, pd.DataFrame]:
    """Находит лист с таблицей инфляции и ключевой ставки."""
    best_sheet: str | None = None
    best_frame: pd.DataFrame | None = None
    best_score = -1

    required_fragments = (
        "ключевая ставка",
        "инфляция",
        "цель по инфляции",
    )

    for sheet_name, frame in workbook.items():
        sample_values = [
            normalize_text(value)
            for value in frame.head(20).to_numpy().ravel()
        ]

        joined = " | ".join(sample_values)

        score = sum(
            fragment in joined
            for fragment in required_fragments
        )

        print(
            f"Лист {sheet_name!r}: "
            f"совпадений заголовков={score}"
        )

        if score > best_score:
            best_score = score
            best_sheet = sheet_name
            best_frame = frame

    if (
        best_sheet is None
        or best_frame is None
        or best_score < 2
    ):
        raise RuntimeError(
            "Не найден лист с инфляцией и ключевой ставкой."
        )

    return best_sheet, best_frame


def find_header_row(frame: pd.DataFrame) -> int:
    """Находит строку с названиями колонок."""
    for row_index in range(min(len(frame), 30)):
        row_texts = [
            normalize_text(value)
            for value in frame.iloc[row_index].tolist()
        ]

        joined = " | ".join(row_texts)

        if (
            "дата" in joined
            and "ключевая ставка" in joined
            and "инфляция" in joined
        ):
            return row_index

    raise RuntimeError(
        "Не найдена строка заголовков таблицы ЦБ."
    )


def find_column(
    columns: list[str],
    required_fragments: tuple[str, ...],
) -> str:
    """Находит колонку по одному из вариантов названия."""
    for column in columns:
        normalized = normalize_text(column)

        if any(
            fragment in normalized
            for fragment in required_fragments
        ):
            return column

    raise RuntimeError(
        "Не найдена колонка. "
        f"Искались варианты: {required_fragments}. "
        f"Доступные колонки: {columns}"
    )


def parse_cbr_excel(excel_bytes: bytes) -> pd.DataFrame:
    """Преобразует Excel ЦБ в стандартизированный DataFrame."""
    workbook = pd.read_excel(
        io.BytesIO(excel_bytes),
        sheet_name=None,
        header=None,
        engine="openpyxl",
    )

    print("Листы в книге:", list(workbook.keys()))

    sheet_name, raw_frame = find_data_sheet(workbook)
    print(f"Выбран лист: {sheet_name}")

    header_row = find_header_row(raw_frame)
    print(f"Строка заголовков: {header_row}")

    frame = raw_frame.iloc[header_row:].copy()
    frame.columns = [
        str(value).strip()
        for value in frame.iloc[0].tolist()
    ]
    frame = frame.iloc[1:].reset_index(drop=True)

    columns = list(frame.columns)

    date_column = find_column(
        columns,
        ("дата",),
    )

    key_rate_column = find_column(
        columns,
        ("ключевая ставка",),
    )

    inflation_column = find_column(
        columns,
        (
            "инфляция, % г/г",
            "инфляция, %",
        ),
    )

    target_column = find_column(
        columns,
        ("цель по инфляции",),
    )

    result = frame[
        [
            date_column,
            key_rate_column,
            inflation_column,
            target_column,
        ]
    ].copy()

    result.columns = [
        "date",
        "key_rate",
        "russia_inflation_yoy",
        "inflation_target",
    ]

    # Дата ЦБ обычно записана как 06.2026.
    # Преобразуем её в последний день соответствующего месяца.
    date_text = (
        result["date"]
        .astype(str)
        .str.strip()
        .str.replace(",", ".", regex=False)
    )

    parsed_month = pd.to_datetime(
        date_text,
        format="%m.%Y",
        errors="coerce",
    )

    # Запасной вариант на случай полноценной даты в исходнике.
    parsed_full_date = pd.to_datetime(
        result["date"],
        dayfirst=True,
        errors="coerce",
    )

    result["date"] = (
        parsed_month
        .fillna(parsed_full_date)
        + pd.offsets.MonthEnd(0)
    )

    for column in (
        "key_rate",
        "russia_inflation_yoy",
        "inflation_target",
    ):
        result[column] = pd.to_numeric(
            result[column]
            .astype(str)
            .str.replace("\xa0", "", regex=False)
            .str.replace(" ", "", regex=False)
            .str.replace(",", ".", regex=False),
            errors="coerce",
        )

    result = result.dropna(
        subset=[
            "date",
            "key_rate",
            "russia_inflation_yoy",
        ]
    )

    result = result[
        result["date"].dt.date >= FIRST_DATE
    ].copy()

    today_month_end = (
        pd.Timestamp.today().normalize()
        + pd.offsets.MonthEnd(0)
    )

    result = result[
        result["date"] <= today_month_end
    ].copy()

    result = (
        result
        .drop_duplicates(subset=["date"], keep="last")
        .sort_values("date")
        .reset_index(drop=True)
    )

    for column in (
        "key_rate",
        "russia_inflation_yoy",
        "inflation_target",
    ):
        result[column] = result[column].round(2)

    result["date"] = result["date"].dt.strftime(
        "%Y-%m-%d"
    )

    if result.empty:
        raise RuntimeError(
            "После обработки таблица ЦБ оказалась пустой."
        )

    return result


def main() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    end_date = date.today()

    excel_bytes, source_url = download_cbr_excel(
        start_date=FIRST_DATE,
        end_date=end_date,
    )

    result = parse_cbr_excel(excel_bytes)

    result.to_csv(
        OUTPUT_FILE,
        index=False,
        encoding="utf-8-sig",
    )

    SOURCE_INFO_FILE.write_text(
        f"source_url={source_url}\n"
        f"period_from={FIRST_DATE.isoformat()}\n"
        f"period_to={end_date.isoformat()}\n"
        f"updated_at={pd.Timestamp.now().isoformat()}\n"
        f"rows={len(result)}\n"
        f"last_date={result.iloc[-1]['date']}\n",
        encoding="utf-8",
    )

    print()
    print("Готово.")
    print(f"Строк сохранено: {len(result)}")
    print(f"Последняя дата: {result.iloc[-1]['date']}")
    print(
        "Последняя ключевая ставка: "
        f"{result.iloc[-1]['key_rate']}"
    )
    print(
        "Последняя инфляция: "
        f"{result.iloc[-1]['russia_inflation_yoy']}"
    )
    print(f"CSV: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()