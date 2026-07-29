from __future__ import annotations

import io
import re
import ssl
import urllib.request
from datetime import date
from pathlib import Path

import certifi
import pandas as pd


BASE_URL = "https://rosstat.gov.ru/storage/mediabank"
FIRST_YEAR = 2019

PROJECT_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = PROJECT_DIR / "data"

OUTPUT_FILE = DATA_DIR / "russia_cpi_quarterly.csv"
SOURCE_INFO_FILE = DATA_DIR / "russia_cpi_quarterly_source.txt"


def create_ssl_context() -> ssl.SSLContext:
    return ssl.create_default_context(cafile=certifi.where())


def download_file(url: str) -> bytes:
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

    try:
        with urllib.request.urlopen(
            request,
            timeout=90,
            context=create_ssl_context(),
        ) as response:
            content = response.read()

    except Exception:
        # Reserve option for Rosstat certificate problems.
        unsafe_context = ssl.create_default_context()
        unsafe_context.check_hostname = False
        unsafe_context.verify_mode = ssl.CERT_NONE

        with urllib.request.urlopen(
            request,
            timeout=90,
            context=unsafe_context,
        ) as response:
            content = response.read()

    if not content.startswith(b"PK"):
        raise ValueError("Rosstat returned a response that is not an XLSX file")

    return content


def find_latest_file() -> tuple[str, bytes]:
    """
    Searches for the newest file such as:
    ipc_kv2_2005-2026.xlsx
    """

    current_year = date.today().year
    errors: list[str] = []

    for end_year in range(current_year, current_year - 4, -1):
        filename = f"ipc_kv2_2005-{end_year}.xlsx"
        url = f"{BASE_URL}/{filename}"

        print(f"Checking: {url}")

        try:
            content = download_file(url)
            print(f"Found: {filename}")
            return url, content
        except Exception as exc:
            errors.append(f"{filename}: {exc}")

    raise RuntimeError(
        "Could not find the current quarterly CPI file.\n"
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
    if pd.isna(value):
        return None

    if isinstance(value, (int, float)):
        year = int(value)

        if 1990 <= year <= 2100:
            return year

    match = re.search(r"\b(19|20)\d{2}\b", str(value))

    if match:
        return int(match.group(0))

    return None


def parse_quarter(value: object) -> int | None:
    text = normalize_text(value)

    roman_mapping = {
        "i квартал": 1,
        "ii квартал": 2,
        "iii квартал": 3,
        "iv квартал": 4,
    }

    if text in roman_mapping:
        return roman_mapping[text]

    match = re.search(r"\b([1-4])\s*квартал\b", text)

    if match:
        return int(match.group(1))

    return None


def parse_number(value: object) -> float | None:
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
    Finds the sheet containing years, quarters and CPI values.
    """

    best_sheet: str | None = None
    best_frame: pd.DataFrame | None = None
    best_score = -1

    for sheet_name, frame in workbook.items():
        years_count = 0
        quarters_count = 0

        sample = frame.iloc[:, :8]

        for value in sample.to_numpy().ravel():
            if parse_year(value) is not None:
                years_count += 1

            if parse_quarter(value) is not None:
                quarters_count += 1

        score = years_count + quarters_count * 5

        print(
            f"Sheet {sheet_name!r}: "
            f"years={years_count}, quarters={quarters_count}"
        )

        if quarters_count >= 8 and score > best_score:
            best_sheet = sheet_name
            best_frame = frame
            best_score = score

    if best_sheet is None or best_frame is None:
        raise RuntimeError("Quarterly CPI sheet was not found")

    return best_sheet, best_frame


def extract_quarterly_cpi(frame: pd.DataFrame) -> pd.DataFrame:
    """
    Reads the first four columns:
    A — year / quarter
    B — quarter end to previous quarter end
    C — quarter to previous quarter
    D — quarter to same quarter of previous year
    """

    records: list[dict[str, object]] = []
    current_year: int | None = None

    for row_index in range(len(frame)):
        first_cell = frame.iloc[row_index, 0]

        year = parse_year(first_cell)

        if year is not None:
            current_year = year
            continue

        quarter = parse_quarter(first_cell)

        if quarter is None or current_year is None:
            continue

        if current_year < FIRST_YEAR:
            continue

        end_to_previous_end = parse_number(
            frame.iloc[row_index, 1]
            if frame.shape[1] > 1
            else None
        )

        quarter_to_previous = parse_number(
            frame.iloc[row_index, 2]
            if frame.shape[1] > 2
            else None
        )

        quarter_to_previous_year = parse_number(
            frame.iloc[row_index, 3]
            if frame.shape[1] > 3
            else None
        )

        if (
            end_to_previous_end is None
            and quarter_to_previous is None
            and quarter_to_previous_year is None
        ):
            continue

        quarter_end_month = quarter * 3

        observation_date = (
            pd.Timestamp(
                year=current_year,
                month=quarter_end_month,
                day=1,
            )
            + pd.offsets.MonthEnd(0)
        )

        records.append(
            {
                "date": observation_date,
                "year": current_year,
                "quarter": quarter,
                "cpi_end_to_previous_quarter_end": end_to_previous_end,
                "cpi_quarter_to_previous_quarter": quarter_to_previous,
                "cpi_quarter_to_previous_year": quarter_to_previous_year,
            }
        )

    if not records:
        raise RuntimeError("No quarterly CPI observations were extracted")

    result = (
        pd.DataFrame(records)
        .drop_duplicates(subset=["date"], keep="last")
        .sort_values("date")
        .reset_index(drop=True)
    )

    today = pd.Timestamp.today().normalize()

    result = result[result["date"] <= today].copy()

    numeric_columns = [
        "cpi_end_to_previous_quarter_end",
        "cpi_quarter_to_previous_quarter",
        "cpi_quarter_to_previous_year",
    ]

    for column in numeric_columns:
        result[column] = result[column].round(2)

    result["date"] = result["date"].dt.strftime("%Y-%m-%d")

    return result[
        [
            "date",
            "year",
            "quarter",
            "cpi_end_to_previous_quarter_end",
            "cpi_quarter_to_previous_quarter",
            "cpi_quarter_to_previous_year",
        ]
    ]


def main() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    source_url, file_content = find_latest_file()

    workbook = pd.read_excel(
        io.BytesIO(file_content),
        sheet_name=None,
        header=None,
        engine="openpyxl",
    )

    print("Workbook sheets:", list(workbook.keys()))

    sheet_name, target_frame = find_target_sheet(workbook)
    print(f"Selected sheet: {sheet_name}")

    result = extract_quarterly_cpi(target_frame)

    result.to_csv(
        OUTPUT_FILE,
        index=False,
        encoding="utf-8-sig",
    )

    SOURCE_INFO_FILE.write_text(
        f"source_url={source_url}\n"
        f"sheet={sheet_name}\n"
        f"updated_at={pd.Timestamp.now().isoformat()}\n"
        f"rows={len(result)}\n"
        f"last_date={result.iloc[-1]['date']}\n",
        encoding="utf-8",
    )

    print()
    print("Done.")
    print(f"Rows saved: {len(result)}")
    print(f"Last date: {result.iloc[-1]['date']}")
    print(
        "Last values: "
        f"{result.iloc[-1]['cpi_end_to_previous_quarter_end']}, "
        f"{result.iloc[-1]['cpi_quarter_to_previous_quarter']}, "
        f"{result.iloc[-1]['cpi_quarter_to_previous_year']}"
    )
    print(f"CSV: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
