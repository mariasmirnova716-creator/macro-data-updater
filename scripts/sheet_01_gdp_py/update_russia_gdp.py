from __future__ import annotations

import io
import re
import ssl
import urllib.request
from datetime import date
from pathlib import Path
from typing import Any

import certifi
import pandas as pd


BASE_URL = "https://rosstat.gov.ru/storage/mediabank"
FIRST_YEAR = 2011

PROJECT_DIR = Path(__file__).resolve().parents[2]
DATA_DIR = PROJECT_DIR / "data" / "sheet_01_gdp_data"

OUTPUT_FILE = DATA_DIR / "russia_gdp_quarterly.csv"
SOURCE_INFO_FILE = DATA_DIR / "russia_gdp_quarterly_source.txt"


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
            timeout=120,
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
            timeout=120,
            context=unsafe_context,
        ) as response:
            content = response.read()

    if not content.startswith(b"PK"):
        raise RuntimeError(
            "Rosstat returned a response that is not an XLSX file."
        )

    return content


def find_latest_file() -> tuple[str, bytes]:
    """
    Searches for a current workbook such as:
    VVP_kvartal_s-1995-2026.xlsx
    """

    current_year = date.today().year
    attempts: list[str] = []

    for end_year in range(current_year, current_year - 4, -1):
        filename = f"VVP_kvartal_s-1995-{end_year}.xlsx"
        url = f"{BASE_URL}/{filename}"

        print(f"Checking: {url}")

        try:
            content = download_file(url)
            print(f"Found: {filename}")
            return url, content

        except Exception as exc:
            attempts.append(f"{filename}: {exc}")

    raise RuntimeError(
        "Could not find the current Rosstat GDP workbook.\n"
        + "\n".join(attempts)
    )


def normalize_text(value: Any) -> str:
    if pd.isna(value):
        return ""

    text = str(value).lower().replace("ё", "е")
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text)

    return text.strip()


def parse_number(value: Any) -> float | None:
    if pd.isna(value):
        return None

    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    text = text.replace("\xa0", "")
    text = text.replace(" ", "")
    text = text.replace(",", ".")

    # Remove footnote characters while retaining the number.
    match = re.search(r"-?\d+(?:\.\d+)?", text)

    if not match:
        return None

    return float(match.group(0))


def parse_year(value: Any) -> int | None:
    if pd.isna(value):
        return None

    if isinstance(value, (int, float)):
        candidate = int(value)

        if 1990 <= candidate <= 2100:
            return candidate

    match = re.search(r"\b(19|20)\d{2}\b", str(value))

    if match:
        return int(match.group(0))

    return None


def parse_quarter(value: Any) -> int | None:
    text = normalize_text(value)

    patterns = {
        1: (
            r"\b1\s*квартал\b",
            r"\bi\s*квартал\b",
            r"\b1q\b",
            r"\bq1\b",
        ),
        2: (
            r"\b2\s*квартал\b",
            r"\bii\s*квартал\b",
            r"\b2q\b",
            r"\bq2\b",
        ),
        3: (
            r"\b3\s*квартал\b",
            r"\biii\s*квартал\b",
            r"\b3q\b",
            r"\bq3\b",
        ),
        4: (
            r"\b4\s*квартал\b",
            r"\biv\s*квартал\b",
            r"\b4q\b",
            r"\bq4\b",
        ),
    }

    for quarter, quarter_patterns in patterns.items():
        if any(re.search(pattern, text) for pattern in quarter_patterns):
            return quarter

    return None


def score_sheet(frame: pd.DataFrame) -> int:
    """
    Gives preference to sheets containing GDP,
    constant prices and a 2021 price-base reference.
    """

    sample = frame.iloc[:80, :30]
    text = " ".join(
        normalize_text(value)
        for value in sample.to_numpy().ravel()
    )

    score = 0

    if "валов" in text and "внутрен" in text and "продукт" in text:
        score += 10

    if "ввп" in text:
        score += 10

    if "постоян" in text:
        score += 10

    if "2021" in text:
        score += 8

    if "квартал" in text:
        score += 5

    return score


def find_target_sheet(
    workbook: dict[str, pd.DataFrame],
) -> tuple[str, pd.DataFrame]:
    """
    Finds the GDP sheet at constant 2021 prices
    without seasonal adjustment.
    """

    candidates: list[tuple[int, str, pd.DataFrame, str]] = []

    for sheet_name, frame in workbook.items():
        sample = frame.iloc[:15, :15]

        sheet_text = " ".join(
            normalize_text(value)
            for value in sample.to_numpy().ravel()
        )

        score = 0

        if "валовой внутренний продукт" in sheet_text:
            score += 20

        if "в ценах 2021" in sheet_text:
            score += 20

        if "млрд" in sheet_text and "руб" in sheet_text:
            score += 5

        # This is the key exclusion:
        # sheet 10 contains seasonally adjusted GDP.
        if "с исключением сезонного" in sheet_text:
            score -= 100

        if "сезонного фактора" in sheet_text:
            score -= 100

        print(
            f"Sheet {sheet_name!r}: "
            f"score={score}, "
            f"title={sheet_text[:180]}"
        )

        if score > 0:
            candidates.append(
                (
                    score,
                    str(sheet_name),
                    frame,
                    sheet_text,
                )
            )

    if not candidates:
        raise RuntimeError(
            "Could not find the non-seasonally-adjusted GDP sheet."
        )

    candidates.sort(
        key=lambda item: item[0],
        reverse=True,
    )

    best_score, best_name, best_frame, best_text = candidates[0]



    if best_score < 20:
        raise RuntimeError(
            "GDP sheet identification is uncertain. "
            f"Best candidate: {best_name!r}, text={best_text[:250]}"
        )

    return best_name, best_frame


def find_target_row(frame: pd.DataFrame) -> int:
    """
    Finds the actual data row for GDP at constant 2021 prices.
    Footnotes and explanatory rows are penalized.
    """

    candidates: list[tuple[int, int, int, str]] = []

    for row_index in range(len(frame)):
        row_values = frame.iloc[row_index].tolist()

        row_text = " ".join(
            normalize_text(value)
            for value in row_values[:15]
        )

        score = 0

        if "ввп" in row_text:
            score += 10

        if (
            "валов" in row_text
            and "внутрен" in row_text
            and "продукт" in row_text
        ):
            score += 10

        if "постоян" in row_text:
            score += 10

        if "2021" in row_text:
            score += 10

        if "цен" in row_text:
            score += 3

        # Explanatory and methodological rows must not be selected.
        if "оценка данных" in row_text:
            score -= 30

        if "осуществляется" in row_text:
            score -= 20

        if "программном продукте" in row_text:
            score -= 20

        if "методолог" in row_text:
            score -= 20

        # A real GDP row must contain many numeric observations.
        numeric_count = sum(
            parse_number(value) is not None
            for value in row_values[1:]
        )

        if numeric_count >= 20:
            score += 20
        elif numeric_count >= 10:
            score += 10
        elif numeric_count < 4:
            score -= 20

        if score > 0:
            candidates.append(
                (
                    score,
                    numeric_count,
                    row_index,
                    row_text,
                )
            )

    if not candidates:
        raise RuntimeError(
            "No rows resembling GDP at constant prices were found."
        )

    # First prefer the score, then the row with more actual numbers.
    candidates.sort(
        key=lambda item: (item[0], item[1]),
        reverse=True,
    )

    print()
    print("Best candidate rows:")

    for score, numeric_count, row_index, text in candidates[:10]:
        print(
            f"row={row_index}, "
            f"score={score}, "
            f"numeric_count={numeric_count}, "
            f"text={text[:250]}"
        )

    best_score, best_numeric_count, best_row, _ = candidates[0]

    if best_score < 20 or best_numeric_count < 10:
        raise RuntimeError(
            "GDP row identification is uncertain. "
            "Review the printed candidate rows."
        )

    return best_row


def find_header_rows(
    frame: pd.DataFrame,
    target_row: int,
) -> tuple[int, int]:
    """
    Finds nearby rows containing years and quarter labels.
    """

    search_start = max(0, target_row - 20)
    search_end = target_row

    best_year_row = -1
    best_year_count = 0

    best_period_row = -1
    best_period_count = 0

    for row_index in range(search_start, search_end):
        values = frame.iloc[row_index].tolist()

        year_count = sum(
            parse_year(value) is not None
            for value in values
        )

        period_count = sum(
            (
                parse_quarter(value) is not None
                or "год" in normalize_text(value)
            )
            for value in values
        )

        if year_count > best_year_count:
            best_year_count = year_count
            best_year_row = row_index

        if period_count > best_period_count:
            best_period_count = period_count
            best_period_row = row_index

    if best_year_row < 0 or best_year_count < 2:
        raise RuntimeError("Could not identify the GDP year header row.")

    if best_period_row < 0 or best_period_count < 4:
        raise RuntimeError("Could not identify the GDP period header row.")

    return best_year_row, best_period_row


def build_column_periods(
    frame: pd.DataFrame,
    year_row: int,
    period_row: int,
) -> dict[int, tuple[int, int, str]]:
    """
    Builds a column-to-period mapping for quarterly GDP.

    Uses explicit year headers where available.
    If merged or malformed year headers are not read correctly,
    derives the next year from the quarter sequence:
    Q4 -> Q1 means that the year increased by one.
    """

    periods: dict[int, tuple[int, int, str]] = {}

    current_year: int | None = None
    previous_quarter: int | None = None

    detected_years: list[tuple[int, int]] = []

    # First collect all explicitly readable year headers.
    for column in range(frame.shape[1]):
        year_candidate = parse_year(
            frame.iat[year_row, column]
        )

        if year_candidate is not None:
            detected_years.append(
                (column, year_candidate)
            )

    print()
    print("Explicit year headers found:")
    print(detected_years)

    for column in range(frame.shape[1]):
        year_candidate = parse_year(
            frame.iat[year_row, column]
        )

        quarter = parse_quarter(
            frame.iat[period_row, column]
        )

        if quarter is None:
            continue

        # Use an explicit year whenever Rosstat provides one.
        if year_candidate is not None:
            current_year = year_candidate

        # If the next Q1 follows Q4, increment the year even when
        # the merged year header was not read correctly.
        elif (
            current_year is not None
            and previous_quarter == 4
            and quarter == 1
        ):
            current_year += 1

        if current_year is None:
            continue

        periods[column] = (
            current_year,
            quarter,
            "quarter",
        )

        previous_quarter = quarter

    if not periods:
        raise RuntimeError(
            "No quarterly period columns were identified."
        )

    print()
    print("Detected quarterly columns:")

    for column, period in periods.items():
        print(
            f"column={column}, "
            f"year={period[0]}, "
            f"quarter={period[1]}"
        )

    return periods

def extract_gdp(frame: pd.DataFrame) -> pd.DataFrame:
    """
    Extracts quarterly GDP from the selected Rosstat sheet.

    Annual values are calculated as the sum of four quarters.
    An incomplete current year receives no annual value.
    """

    target_row = find_target_row(frame)

    year_row, period_row = find_header_rows(
        frame,
        target_row,
    )

    print()
    print(f"Selected data row: {target_row}")
    print(f"Selected year header row: {year_row}")
    print(f"Selected period header row: {period_row}")

    periods = build_column_periods(
        frame,
        year_row,
        period_row,
    )

    detected_quarters: list[dict[str, Any]] = []

    for column, (
        observation_year,
        quarter,
        period_type,
    ) in periods.items():

        if observation_year < FIRST_YEAR:
            continue

        value = parse_number(
            frame.iat[target_row, column]
        )

        if value is None:
            continue

        detected_quarters.append(
            {
                "year": observation_year,
                "quarter": quarter,
                "gdp_constant_2021_prices_bln_rub": value,
            }
        )

    if not detected_quarters:
        raise RuntimeError(
            "No quarterly GDP observations were extracted."
        )

    quarterly_frame = (
        pd.DataFrame(detected_quarters)
        .drop_duplicates(
            subset=["year", "quarter"],
            keep="last",
        )
        .sort_values(["year", "quarter"])
        .reset_index(drop=True)
    )

    print()
    print("Detected years and quarters:")
    print(
        quarterly_frame
        .groupby("year")["quarter"]
        .apply(list)
        .to_string()
    )

    records: list[dict[str, Any]] = []

    for row in quarterly_frame.itertuples(index=False):
        observation_date = (
            pd.Timestamp(
                year=int(row.year),
                month=int(row.quarter) * 3,
                day=1,
            )
            + pd.offsets.MonthEnd(0)
        )

        records.append(
            {
                "date": observation_date,
                "year": int(row.year),
                "quarter": int(row.quarter),
                "period_type": "quarter",
                "gdp_constant_2021_prices_bln_rub": float(
                    row.gdp_constant_2021_prices_bln_rub
                ),
            }
        )

    annual_frame = (
        quarterly_frame
        .groupby("year", as_index=False)
        .agg(
            quarter_count=("quarter", "nunique"),
            gdp_constant_2021_prices_bln_rub=(
                "gdp_constant_2021_prices_bln_rub",
                "sum",
            ),
        )
    )

    annual_frame = annual_frame[
        annual_frame["quarter_count"] == 4
    ].copy()

    for row in annual_frame.itertuples(index=False):
        records.append(
            {
                "date": pd.Timestamp(
                    year=int(row.year),
                    month=12,
                    day=31,
                ),
                "year": int(row.year),
                "quarter": None,
                "period_type": "year",
                "gdp_constant_2021_prices_bln_rub": float(
                    row.gdp_constant_2021_prices_bln_rub
                ),
            }
        )

    result = pd.DataFrame(records)

    result["period_order"] = result["period_type"].map(
        {
            "quarter": 1,
            "year": 2,
        }
    )

    result = (
        result
        .sort_values(
            ["year", "period_order", "quarter"],
            na_position="last",
        )
        .drop(columns="period_order")
        .reset_index(drop=True)
    )

    result["gdp_constant_2021_prices_bln_rub"] = (
        pd.to_numeric(
            result["gdp_constant_2021_prices_bln_rub"],
            errors="coerce",
        )
        .round(1)
    )

    result["date"] = (
        pd.to_datetime(result["date"])
        .dt.strftime("%Y-%m-%d")
    )

    return result

def validate_result(result: pd.DataFrame) -> None:
    required_columns = {
        "date",
        "year",
        "quarter",
        "period_type",
        "gdp_constant_2021_prices_bln_rub",
    }

    missing = required_columns.difference(result.columns)

    if missing:
        raise RuntimeError(
            f"Missing output columns: {sorted(missing)}"
        )

    if result.empty:
        raise RuntimeError("GDP result is empty.")

    values = pd.to_numeric(
        result["gdp_constant_2021_prices_bln_rub"],
        errors="coerce",
    )

    if values.isna().any():
        raise RuntimeError(
            "GDP output contains non-numeric values."
        )

    if not values.between(1_000, 1_000_000).all():
        bad_values = values[
            ~values.between(1_000, 1_000_000)
        ].tolist()

        raise RuntimeError(
            f"GDP output contains implausible values: {bad_values[:10]}"
        )

    quarter_rows = result[
        result["period_type"] == "quarter"
    ]

    if len(quarter_rows) < 40:
        raise RuntimeError(
            f"Too few quarterly observations: {len(quarter_rows)}"
        )

    latest_quarter_year = int(
        pd.to_numeric(
            quarter_rows["year"],
            errors="coerce",
        ).max()
    )

    minimum_expected_year = date.today().year - 1

    if latest_quarter_year < minimum_expected_year:
        raise RuntimeError(
            "GDP data unexpectedly stops at "
            f"{latest_quarter_year}. "
            f"Expected observations through at least "
            f"{minimum_expected_year}. "
            "The Rosstat table structure may have changed."
        )

    duplicate_quarters = quarter_rows.duplicated(
        subset=["year", "quarter"],
        keep=False,
    )

    if duplicate_quarters.any():
        duplicates = quarter_rows.loc[
            duplicate_quarters,
            ["year", "quarter"],
        ].to_dict("records")

        raise RuntimeError(
            "Duplicate GDP quarters found: "
            f"{duplicates[:10]}"
        )


def main() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    source_url, file_content = find_latest_file()

    workbook = pd.read_excel(
        io.BytesIO(file_content),
        sheet_name=None,
        header=None,
        engine="openpyxl",
    )

    print()
    print("Workbook sheets:")
    print(list(workbook.keys()))
    print()

    sheet_name, target_frame = find_target_sheet(workbook)

    print()
    print(f"Selected sheet: {sheet_name}")

    result = extract_gdp(target_frame)
    validate_result(result)

    result.to_csv(
        OUTPUT_FILE,
        index=False,
        encoding="utf-8-sig",
    )

    latest_row = result.iloc[-1]

    SOURCE_INFO_FILE.write_text(
        f"source_url={source_url}\n"
        f"sheet={sheet_name}\n"
        f"updated_at={pd.Timestamp.now().isoformat()}\n"
        f"rows={len(result)}\n"
        f"last_date={latest_row['date']}\n",
        encoding="utf-8",
    )

    print()
    print("Done.")
    print(f"Rows saved: {len(result)}")
    print(f"Last date: {latest_row['date']}")
    print(
        "Last value: "
        f"{latest_row['gdp_constant_2021_prices_bln_rub']}"
    )
    print(f"CSV: {OUTPUT_FILE}")

    print()
    print("Last 12 observations:")
    print(result.tail(12).to_string(index=False))


if __name__ == "__main__":
    main()