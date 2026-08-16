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


# ============================================================
# SETTINGS
# ============================================================

BASE_URL = "https://rosstat.gov.ru/storage/mediabank"
FIRST_YEAR = 2019

PROJECT_DIR = Path(__file__).resolve().parents[2]
DATA_DIR = PROJECT_DIR / "data" / "sheet_01_gdp_data"

OUTPUT_FILE = DATA_DIR / "russia_gdp_nominal_quarterly.csv"
SOURCE_INFO_FILE = DATA_DIR / "russia_gdp_nominal_quarterly_source.txt"


# ============================================================
# SSL / DOWNLOAD
# ============================================================

def create_ssl_context() -> ssl.SSLContext:
    return ssl.create_default_context(
        cafile=certifi.where()
    )


def download_file(url: str) -> bytes:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 "
                "(Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 "
                "Chrome/126 Safari/537.36"
            ),
            "Accept": (
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet,"
                "application/octet-stream,"
                "*/*"
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


# ============================================================
# FIND LATEST ROSSTAT FILE
# ============================================================

def find_latest_file() -> tuple[str, bytes]:
    """
    Looks for the newest workbook like:
    VVP_kvartal_s-1995-2026.xlsx
    """

    current_year = date.today().year
    attempts: list[str] = []

    for end_year in range(
        current_year,
        current_year - 4,
        -1,
    ):
        filename = f"VVP_kvartal_s-1995-{end_year}.xlsx"
        url = f"{BASE_URL}/{filename}"

        print(f"Checking: {url}")

        try:
            content = download_file(url)

            print(f"Found: {filename}")

            return url, content

        except Exception as exc:
            attempts.append(
                f"{filename}: {exc}"
            )

    raise RuntimeError(
        "Could not find the current Rosstat GDP workbook.\n"
        + "\n".join(attempts)
    )


# ============================================================
# HELPERS
# ============================================================

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

    text = (
        text
        .replace("\xa0", "")
        .replace(" ", "")
        .replace(",", ".")
    )

    match = re.search(
        r"-?\d+(?:\.\d+)?",
        text,
    )

    if not match:
        return None

    return float(
        match.group(0)
    )


def parse_year(value: Any) -> int | None:
    """
    Extracts a calendar year from a YEAR HEADER cell.

    Handles:
        2026
        "2026"
        "2026 2)"
        "2026²)"
        "2026¹"
        numeric 20262  -> 2026, if this is a header cell

    IMPORTANT:
    This function should be used for header rows,
    not arbitrary GDP value cells.
    """

    if pd.isna(value):
        return None

    # Normal numeric year.
    if isinstance(value, (int, float)):
        number = int(value)

        if 1990 <= number <= 2100:
            return number

        # Protection against footnote accidentally attached
        # to the year as an extra final digit:
        #
        # 20262 -> 2026
        # 20261 -> 2026
        text = str(abs(number))

        if len(text) == 5:
            first_four = int(text[:4])

            if 1990 <= first_four <= 2100:
                return first_four

        return None

    text = str(value)

    # Convert superscript digits to ordinary digits only
    # so that strings like 2026² can still be understood.
    superscripts = str.maketrans({
        "⁰": "0",
        "¹": "1",
        "²": "2",
        "³": "3",
        "⁴": "4",
        "⁵": "5",
        "⁶": "6",
        "⁷": "7",
        "⁸": "8",
        "⁹": "9",
    })

    text = text.translate(superscripts)

    # Prefer a plausible 4-digit year appearing at the
    # beginning of the cell.
    match = re.match(
        r"\s*((?:19|20)\d{2})",
        text,
    )

    if match:
        year = int(match.group(1))

        if 1990 <= year <= 2100:
            return year

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
        if any(
            re.search(
                pattern,
                text,
            )
            for pattern in quarter_patterns
        ):
            return quarter

    return None

def detect_latest_header_year(
    frame: pd.DataFrame,
) -> int:

    search_rows = min(
        15,
        len(frame),
    )

    # --------------------------------------------------------
    # 1. Find the row containing quarter headers.
    # --------------------------------------------------------

    quarter_row = None
    best_quarter_count = 0

    for row_index in range(search_rows):

        values = frame.iloc[
            row_index
        ].tolist()

        quarter_count = sum(
            parse_quarter(value) is not None
            for value in values
        )

        if quarter_count > best_quarter_count:
            best_quarter_count = quarter_count
            quarter_row = row_index

    if (
        quarter_row is None
        or best_quarter_count < 4
    ):
        return 0

    # --------------------------------------------------------
    # 2. Years must be in the rows immediately ABOVE quarters.
    #
    # Usually:
    # year row
    # quarter row
    # value row
    #
    # Search only a few rows above.
    # --------------------------------------------------------

    start_row = max(
        0,
        quarter_row - 3,
    )

    latest_year = 0

    for row_index in range(
        start_row,
        quarter_row,
    ):

        for value in frame.iloc[
            row_index
        ].tolist():

            year = parse_year(
                value
            )

            if (
                year is not None
                and year > latest_year
            ):
                latest_year = year

    return latest_year

# ============================================================
# FIND NOMINAL GDP SHEET
# ============================================================

def find_nominal_gdp_sheet(
    workbook: dict[str, pd.DataFrame],
) -> tuple[str, pd.DataFrame]:

    candidates: list[
        tuple[
            int,
            int,
            str,
            pd.DataFrame,
            str,
        ]
    ] = []

    for sheet_name, frame in workbook.items():

        sample = frame.iloc[
            :15,
            :30,
        ]

        sheet_text = " ".join(
            normalize_text(value)
            for value in sample.to_numpy().ravel()
        )

        score = 0

        # ----------------------------------------------------
        # We need GDP
        # ----------------------------------------------------

        if (
            "валовой внутренний продукт"
            in sheet_text
        ):
            score += 20

        # ----------------------------------------------------
        # We specifically need nominal GDP:
        # current prices
        # ----------------------------------------------------

        if (
            "в текущих ценах"
            in sheet_text
        ):
            score += 30

        if (
            "млрд" in sheet_text
            and "руб" in sheet_text
        ):
            score += 10

        # ----------------------------------------------------
        # Exclude real GDP / constant-price sheets
        # ----------------------------------------------------

        if (
            "в ценах 2021"
            in sheet_text
        ):
            score -= 100

        if (
            "постоянных ценах"
            in sheet_text
        ):
            score -= 100

        if (
            "с исключением сезонного"
            in sheet_text
        ):
            score -= 50

        # ----------------------------------------------------
        # Find the latest year actually visible on this sheet.
        #
        # This separates:
        #
        # Sheet 1 -> old nominal GDP history ending around 2011
        # Sheet 2 -> current nominal GDP history through 2026
        # ----------------------------------------------------

        latest_year = detect_latest_header_year(
            frame
        )

        print(
            f"Sheet {sheet_name!r}: "
            f"score={score}, "
            f"latest_year={latest_year}, "
            f"title={sheet_text[:200]}"
        )

        if score > 0:

            candidates.append(
                (
                    score,
                    latest_year,
                    str(sheet_name),
                    frame,
                    sheet_text,
                )
            )

    if not candidates:

        raise RuntimeError(
            "Could not find nominal GDP sheet."
        )

    # --------------------------------------------------------
    # First:
    # choose the correct type of table by score.
    #
    # If several sheets have the same score,
    # choose the one with the newest data.
    # --------------------------------------------------------

    candidates.sort(
        key=lambda item: (
            item[0],
            item[1],
        ),
        reverse=True,
    )

    (
        best_score,
        best_latest_year,
        best_name,
        best_frame,
        best_text,
    ) = candidates[0]

    if best_score < 40:

        raise RuntimeError(
            "Nominal GDP sheet identification is uncertain. "
            f"Best candidate: {best_name!r}, "
            f"score={best_score}, "
            f"latest_year={best_latest_year}, "
            f"text={best_text[:250]}"
        )

    if best_latest_year < FIRST_YEAR:

        raise RuntimeError(
            "Selected nominal GDP sheet is too old. "
            f"Sheet={best_name!r}, "
            f"latest_year={best_latest_year}"
        )

    print()
    print(
        "Nominal GDP sheet selected:"
    )

    print(
        f"sheet={best_name!r}, "
        f"latest_year={best_latest_year}"
    )

    return (
        best_name,
        best_frame,
    )


# ============================================================
# FIND GDP VALUE ROW
# ============================================================

def find_nominal_gdp_row(
    frame: pd.DataFrame,
) -> int:

    candidates: list[
        tuple[int, int, int]
    ] = []

    for row_index in range(
        len(frame)
    ):
        row_values = (
            frame
            .iloc[row_index]
            .tolist()
        )

        row_text = " ".join(
            normalize_text(value)
            for value in row_values[:15]
        )

        numeric_count = sum(
            parse_number(value) is not None
            for value in row_values
        )

        score = 0

        if numeric_count >= 20:
            score += 30

        elif numeric_count >= 10:
            score += 15

        elif numeric_count < 4:
            score -= 20

        if "валовой внутренний продукт" in row_text:
            score -= 10

        if "текущих ценах" in row_text:
            score -= 10

        if "квартал" in row_text:
            score -= 10

        if "данные содержат" in row_text:
            score -= 30

        if "методолог" in row_text:
            score -= 30

        if score > 0:
            candidates.append(
                (
                    score,
                    numeric_count,
                    row_index,
                )
            )

    if not candidates:
        raise RuntimeError(
            "Could not find nominal GDP value row."
        )

    candidates.sort(
        key=lambda item: (
            item[0],
            item[1],
        ),
        reverse=True,
    )

    print()
    print("Best nominal GDP row candidates:")

    for (
        score,
        numeric_count,
        row_index,
    ) in candidates[:10]:
        print(
            f"row={row_index}, "
            f"score={score}, "
            f"numeric_count={numeric_count}"
        )

    (
        best_score,
        best_numeric_count,
        best_row,
    ) = candidates[0]

    if (
        best_score < 15
        or best_numeric_count < 10
    ):
        raise RuntimeError(
            "Nominal GDP row identification is uncertain."
        )

    return best_row


# ============================================================
# FIND HEADER ROWS
# ============================================================

def find_header_rows(
    frame: pd.DataFrame,
    target_row: int,
) -> tuple[int, int]:

    search_start = max(
        0,
        target_row - 20,
    )

    search_end = target_row

    best_year_row = -1
    best_year_count = 0

    best_period_row = -1
    best_period_count = 0

    for row_index in range(
        search_start,
        search_end,
    ):
        values = (
            frame
            .iloc[row_index]
            .tolist()
        )

        year_count = sum(
            parse_year(value) is not None
            for value in values
        )

        period_count = sum(
            parse_quarter(value) is not None
            for value in values
        )

        if year_count > best_year_count:
            best_year_count = year_count
            best_year_row = row_index

        if period_count > best_period_count:
            best_period_count = period_count
            best_period_row = row_index

    if (
        best_year_row < 0
        or best_year_count < 2
    ):
        raise RuntimeError(
            "Could not identify nominal GDP year header row."
        )

    if (
        best_period_row < 0
        or best_period_count < 4
    ):
        raise RuntimeError(
            "Could not identify nominal GDP quarter header row."
        )

    return (
        best_year_row,
        best_period_row,
    )


# ============================================================
# BUILD COLUMN -> YEAR / QUARTER MAP
# ============================================================

def build_column_periods(
    frame: pd.DataFrame,
    year_row: int,
    period_row: int,
) -> dict[int, tuple[int, int]]:

    periods: dict[
        int,
        tuple[int, int],
    ] = {}

    current_year: int | None = None
    previous_quarter: int | None = None

    detected_years: list[
        tuple[int, int]
    ] = []

    for column in range(
        frame.shape[1]
    ):
        year_candidate = parse_year(
            frame.iat[
                year_row,
                column,
            ]
        )

        if year_candidate is not None:
            detected_years.append(
                (
                    column,
                    year_candidate,
                )
            )

    print()
    print(
        "Explicit year headers found:"
    )

    print(
        detected_years
    )

    for column in range(
        frame.shape[1]
    ):
        year_candidate = parse_year(
            frame.iat[
                year_row,
                column,
            ]
        )

        quarter = parse_quarter(
            frame.iat[
                period_row,
                column,
            ]
        )

        if quarter is None:
            continue

        if year_candidate is not None:
            current_year = year_candidate

        elif (
            current_year is not None
            and previous_quarter == 4
            and quarter == 1
        ):
            current_year += 1

        if current_year is None:
            continue

        periods[
            column
        ] = (
            current_year,
            quarter,
        )

        previous_quarter = quarter

    if not periods:
        raise RuntimeError(
            "No nominal GDP quarterly columns were identified."
        )

    print()
    print(
        "Detected nominal GDP columns:"
    )

    for column, (
        observation_year,
        quarter,
    ) in periods.items():
        print(
            f"column={column}, "
            f"year={observation_year}, "
            f"quarter={quarter}"
        )

    return periods


# ============================================================
# EXTRACT NOMINAL GDP
# ============================================================

def extract_nominal_gdp(
    frame: pd.DataFrame,
) -> pd.DataFrame:

    target_row = find_nominal_gdp_row(
        frame
    )

    (
        year_row,
        period_row,
    ) = find_header_rows(
        frame,
        target_row,
    )

    print()
    print(
        f"Selected data row: {target_row}"
    )

    print(
        f"Selected year header row: {year_row}"
    )

    print(
        f"Selected quarter header row: {period_row}"
    )

    periods = build_column_periods(
        frame,
        year_row,
        period_row,
    )

    records: list[
        dict[str, Any]
    ] = []

    for column, (
        observation_year,
        quarter,
    ) in periods.items():

        if observation_year < FIRST_YEAR:
            continue

        value = parse_number(
            frame.iat[
                target_row,
                column,
            ]
        )

        if value is None:
            continue

        observation_date = (
            pd.Timestamp(
                year=observation_year,
                month=quarter * 3,
                day=1,
            )
            + pd.offsets.MonthEnd(0)
        )

        records.append(
            {
                "date": observation_date,
                "year": observation_year,
                "quarter": quarter,
                "gdp_nominal_bln_rub": value,
            }
        )

    result = pd.DataFrame(
        records
    )

    if result.empty:
        raise RuntimeError(
            "No nominal GDP observations were extracted."
        )

    result = (
        result
        .drop_duplicates(
            subset=[
                "year",
                "quarter",
            ],
            keep="last",
        )
        .sort_values(
            [
                "year",
                "quarter",
            ]
        )
        .reset_index(
            drop=True
        )
    )

    result[
        "gdp_nominal_bln_rub"
    ] = (
        pd.to_numeric(
            result[
                "gdp_nominal_bln_rub"
            ],
            errors="coerce",
        )
        .round(1)
    )

    result[
        "date"
    ] = (
        pd.to_datetime(
            result["date"]
        )
        .dt.strftime(
            "%Y-%m-%d"
        )
    )

    return result


# ============================================================
# VALIDATION
# ============================================================

def validate_result(
    result: pd.DataFrame,
) -> None:

    required_columns = {
        "date",
        "year",
        "quarter",
        "gdp_nominal_bln_rub",
    }

    missing = (
        required_columns
        .difference(
            result.columns
        )
    )

    if missing:
        raise RuntimeError(
            "Missing output columns: "
            f"{sorted(missing)}"
        )

    if result.empty:
        raise RuntimeError(
            "Nominal GDP result is empty."
        )
    
    if int(result["year"].min()) > FIRST_YEAR:
        raise RuntimeError(
            "Nominal GDP history does not contain "
            f"the required base year {FIRST_YEAR}."
        )

    values = pd.to_numeric(
        result[
            "gdp_nominal_bln_rub"
        ],
        errors="coerce",
    )

    if values.isna().any():
        raise RuntimeError(
            "Nominal GDP output contains non-numeric values."
        )

    if not values.between(
        1_000,
        1_000_000,
    ).all():
        bad_values = values[
            ~values.between(
                1_000,
                1_000_000,
            )
        ].tolist()

        raise RuntimeError(
            "Nominal GDP output contains implausible values: "
            f"{bad_values[:10]}"
        )

    if len(result) < 20:
        raise RuntimeError(
            "Too few nominal GDP observations: "
            f"{len(result)}"
        )

    duplicate_quarters = (
        result
        .duplicated(
            subset=[
                "year",
                "quarter",
            ],
            keep=False,
        )
    )

    if duplicate_quarters.any():
        duplicates = (
            result.loc[
                duplicate_quarters,
                [
                    "year",
                    "quarter",
                ],
            ]
            .to_dict(
                "records"
            )
        )

        raise RuntimeError(
            "Duplicate nominal GDP quarters found: "
            f"{duplicates[:10]}"
        )

    latest_year = int(
        pd.to_numeric(
            result["year"],
            errors="coerce",
        ).max()
    )

    minimum_expected_year = (
        date.today().year - 1
    )

    if latest_year < minimum_expected_year:
        raise RuntimeError(
            "Nominal GDP data unexpectedly stops at "
            f"{latest_year}. "
            "The Rosstat table structure may have changed."
        )


# ============================================================
# MAIN
# ============================================================

def main() -> None:

    DATA_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    source_url, file_content = (
        find_latest_file()
    )

    workbook = pd.read_excel(
        io.BytesIO(
            file_content
        ),
        sheet_name=None,
        header=None,
        engine="openpyxl",
    )

    print()
    print(
        "Workbook sheets:"
    )

    print(
        list(
            workbook.keys()
        )
    )

    (
        sheet_name,
        target_frame,
    ) = find_nominal_gdp_sheet(
        workbook
    )

    print()
    print(
        "Selected nominal GDP sheet:"
    )

    print(
        sheet_name
    )

    result = extract_nominal_gdp(
        target_frame
    )

    validate_result(
        result
    )

    result.to_csv(
        OUTPUT_FILE,
        index=False,
        encoding="utf-8-sig",
    )

    latest_row = (
        result.iloc[-1]
    )

    SOURCE_INFO_FILE.write_text(
        "\n".join(
            [
                "Russia nominal quarterly GDP",
                "source=Rosstat",
                f"source_url={source_url}",
                f"sheet={sheet_name}",
                (
                    "indicator="
                    "Валовой внутренний продукт "
                    "в текущих ценах"
                ),
                "unit=bln_rub",
                "prices=current",
                f"first_year={FIRST_YEAR}",
                (
                    "last_date="
                    f"{latest_row['date']}"
                ),
                (
                    "updated_at="
                    f"{pd.Timestamp.now().isoformat()}"
                ),
                "",
            ]
        ),
        encoding="utf-8",
    )

    print()
    print(
        "=" * 72
    )

    print(
        "DONE"
    )

    print(
        "=" * 72
    )

    print()
    print(
        "Rows saved:",
        len(result),
    )

    print(
        "Range:",
        result["date"].min(),
        "->",
        result["date"].max(),
    )

    print(
        "Latest nominal GDP:",
        latest_row[
            "gdp_nominal_bln_rub"
        ],
        "bln RUB",
    )

    print(
        "CSV:",
        OUTPUT_FILE,
    )

    print()
    print(
        "Last 12 observations:"
    )

    print(
        result
        .tail(12)
        .to_string(
            index=False
        )
    )


if __name__ == "__main__":
    main()