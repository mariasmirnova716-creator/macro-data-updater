from __future__ import annotations
from datetime import datetime

import io
import re
import warnings
from pathlib import Path


import certifi
import pandas as pd
import requests
from bs4 import BeautifulSoup
from urllib3.exceptions import InsecureRequestWarning



# ============================================================
# PATHS
# ============================================================

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parents[1]

DATA_DIR = (
    PROJECT_ROOT
    / "data"
    / "sheet_01_gdp_data"
)

OUTPUT_FILE = (
    DATA_DIR
    / "russia_loss_making_companies.csv"
)

SOURCE_INFO_FILE = (
    DATA_DIR
    / "russia_loss_making_companies_source.txt"
)


# ============================================================
# EMISS / FEDSTAT
# ============================================================

INDICATOR_ID = 57746

INDICATOR_URL = (
    f"https://www.fedstat.ru/indicator/{INDICATOR_ID}"
)

DOWNLOAD_URL = (
    "https://www.fedstat.ru/indicator/"
    "downloadData.do?format=excel"
)

FIRST_YEAR = 2023


# ------------------------------------------------------------
# These IDs came from the REAL export request captured
# in browser DevTools -> Network -> downloadData.do -> Payload.
# ------------------------------------------------------------

LINE_OBJECT_IDS = [
    "57940",
    "57831",
]

COLUMN_OBJECT_IDS = [
    "3",
    "33560",
]

FILTER_OBJECT_IDS = [
    "0",
    "30611",
]


# ------------------------------------------------------------
# Fixed selected values
# ------------------------------------------------------------

INDICATOR_FILTER = "0_57746"

# Additional filter sent by Fedstat in the real request.
EXTRA_FILTER = "30611_950473"

# Territory:
# Российская Федерация без учета новых субъектов
# (с 01.01.2023)
RF_WITHOUT_NEW_REGIONS_FILTER = "57831_1849012"

# Activity:
# Всего по обследуемым видам экономической деятельности
ALL_ACTIVITIES_FILTER = "57940_1692933"


# ------------------------------------------------------------
# Cumulative periods:
#
# January
# January-February
# ...
# January-December
#
# These 12 IDs were also taken from the real browser request.
# ------------------------------------------------------------

PERIOD_FILTERS = [
    "33560_1540283",
    "33560_1540284",
    "33560_1540285",
    "33560_1540286",
    "33560_1540287",
    "33560_1540288",
    "33560_1540289",
    "33560_1540290",
    "33560_1540291",
    "33560_1540292",
    "33560_1540293",
    "33560_1540294",
]


HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 "
        "(KHTML, like Gecko) "
        "Chrome/148 Safari/537.36"
    ),
    "Accept-Language": "ru,en;q=0.9",
}


# ============================================================
# HTTP
# ============================================================

def request(
    session: requests.Session,
    method: str,
    url: str,
    **kwargs,
) -> requests.Response:

    kwargs.setdefault(
        "timeout",
        90,
    )

    kwargs.setdefault(
        "headers",
        HEADERS,
    )

    try:

        response = session.request(
            method,
            url,
            verify=certifi.where(),
            **kwargs,
        )

    except requests.exceptions.SSLError:

        print(
            "WARNING: SSL verification failed."
        )

        print(
            "Retrying without certificate verification."
        )

        warnings.simplefilter(
            "ignore",
            InsecureRequestWarning,
        )

        response = session.request(
            method,
            url,
            verify=False,
            **kwargs,
        )

    response.raise_for_status()

    return response


# ============================================================
# TEXT HELPERS
# ============================================================

def clean_text(
    value,
) -> str:

    if value is None:
        return ""

    text = str(value)

    text = text.replace(
        "\xa0",
        " ",
    )

    text = text.replace(
        "–",
        "-",
    )

    text = text.replace(
        "—",
        "-",
    )

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.strip()


def normalize_text(
    value,
) -> str:

    return clean_text(
        value
    ).lower()


def as_number(
    value,
) -> float | None:

    if value is None:
        return None

    if isinstance(
        value,
        bool,
    ):
        return None

    if isinstance(
        value,
        (int, float),
    ):

        if pd.isna(value):
            return None

        return float(value)

    text = clean_text(
        value
    )

    text = (
        text
        .replace(" ", "")
        .replace(",", ".")
    )

    if text in {
        "",
        "-",
        "…",
        "...",
    }:
        return None

    match = re.search(
        r"-?\d+(?:\.\d+)?",
        text,
    )

    if match is None:
        return None

    try:

        return float(
            match.group(0)
        )

    except ValueError:

        return None


# ============================================================
# FEDSTAT PAGE
# ============================================================

def extract_struts_token(
    html: str,
) -> tuple[str, str]:

    soup = BeautifulSoup(
        html,
        "html.parser",
    )

    token_name_input = soup.find(
        "input",
        attrs={
            "name": "struts.token.name",
        },
    )

    token_name = "token"

    if token_name_input is not None:

        value = token_name_input.get(
            "value"
        )

        if value:
            token_name = value

    token_input = soup.find(
        "input",
        attrs={
            "name": token_name,
        },
    )

    if token_input is not None:

        token_value = token_input.get(
            "value"
        )

        if token_value:

            return (
                token_name,
                token_value,
            )

    match = re.search(
        r'name=["\']token["\']'
        r'[^>]*value=["\']([^"\']+)["\']',
        html,
        flags=re.IGNORECASE,
    )

    if match:

        return (
            "token",
            match.group(1),
        )

    raise RuntimeError(
        "Could not find Fedstat Struts token."
    )


def extract_available_years(
    html: str,
) -> list[int]:

    current_year = datetime.now().year

    years = {
        int(x)
        for x in re.findall(
            r"\b(20\d{2})\b",
            html,
        )
    }

    years = sorted(
        year
        for year in years
        if FIRST_YEAR <= year <= current_year
    )

    if not years:
        raise RuntimeError(
            "Could not detect years >= 2023 "
            "on Fedstat indicator page."
        )

    return years


# ============================================================
# DOWNLOAD FEDSTAT XLS
# ============================================================

def download_loss_making_xls() -> tuple[bytes, list[int]]:

    print(
        "=" * 72
    )

    print(
        "Fedstat loss-making companies updater"
    )

    print(
        "=" * 72
    )

    session = requests.Session()

    session.headers.update(
        HEADERS
    )

    print(
        "\nOpening indicator:"
    )

    print(
        INDICATOR_URL
    )

    response = request(
        session,
        "GET",
        INDICATOR_URL,
    )

    html = response.text

    print(
        "Session cookies:",
        list(
            session.cookies.keys()
        ),
    )

    token_name, token_value = (
        extract_struts_token(
            html
        )
    )

    print(
        "Token detected."
    )

    years = extract_available_years(
        html
    )

    print(
        "Years selected:",
        years,
    )

    # ========================================================
    # SELECTED FILTERS
    # ========================================================

    selected_filters = [
        INDICATOR_FILTER,
    ]

    # Dynamic years:
    # 2023 -> maximum year currently available on Fedstat.
    selected_filters.extend(
        f"3_{year}"
        for year in years
    )

    selected_filters.append(
        EXTRA_FILTER
    )

    selected_filters.extend(
        PERIOD_FILTERS
    )

    selected_filters.append(
        RF_WITHOUT_NEW_REGIONS_FILTER
    )

    selected_filters.append(
        ALL_ACTIVITIES_FILTER
    )

    # ========================================================
    # BUILD EXACT POST PAYLOAD
    # ========================================================

    payload = [
        (
            "title",
            (
                "Удельный вес убыточных "
                "организаций с 2017 г. (процент)"
            ),
        ),
        (
            "struts.token.name",
            token_name,
        ),
        (
            token_name,
            token_value,
        ),
        (
            "id",
            str(INDICATOR_ID),
        ),
    ]

    for object_id in LINE_OBJECT_IDS:

        payload.append(
            (
                "lineObjectIds",
                object_id,
            )
        )

    for object_id in COLUMN_OBJECT_IDS:

        payload.append(
            (
                "columnObjectIds",
                object_id,
            )
        )

    for filter_id in selected_filters:

        payload.append(
            (
                "selectedFilterIds",
                filter_id,
            )
        )

    for object_id in FILTER_OBJECT_IDS:

        payload.append(
            (
                "filterObjectIds",
                object_id,
            )
        )

    export_headers = {
        **HEADERS,
        "Accept": (
            "text/html,"
            "application/xhtml+xml,"
            "application/xml;q=0.9,"
            "*/*;q=0.8"
        ),
        "Content-Type": (
            "application/x-www-form-urlencoded"
        ),
        "Origin": (
            "https://www.fedstat.ru"
        ),
        "Referer": INDICATOR_URL,
    }

    print(
        "\nRequesting Excel export..."
    )

    export_response = request(
        session,
        "POST",
        DOWNLOAD_URL,
        headers=export_headers,
        data=payload,
    )

    content = export_response.content

    print(
        "Downloaded:",
        len(content),
        "bytes",
    )

    if len(content) < 1000:

        preview = content[
            :500
        ].decode(
            "utf-8",
            errors="ignore",
        )

        raise RuntimeError(
            "Fedstat returned an unexpectedly "
            "small response.\n"
            f"Preview:\n{preview}"
        )

    return (
        content,
        years,
    )


# ============================================================
# PERIOD -> MONTH
# ============================================================

RUSSIAN_MONTHS = {
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


def period_to_month(
    value,
) -> int | None:

    text = normalize_text(
        value
    )

    if not text:
        return None

    # Examples:
    #
    # январь
    # январь-февраль
    # январь-март
    # ...
    # январь-декабрь
    #
    # We need the LAST month of the cumulative period.

    found_months = []

    for month_name, month_number in (
        RUSSIAN_MONTHS.items()
    ):

        if month_name in text:

            found_months.append(
                month_number
            )

    if not found_months:
        return None

    return max(
        found_months
    )


# ============================================================
# PARSE FEDSTAT XLS
# ============================================================

def parse_loss_making_xls(
    content: bytes,
) -> pd.DataFrame:

    # Fedstat currently sends old-style XLS.
    # Pandas uses xlrd for it.

    excel = pd.ExcelFile(
        io.BytesIO(content),
        engine="xlrd",
    )

    if not excel.sheet_names:

        raise RuntimeError(
            "Downloaded Fedstat XLS contains no sheets."
        )

    sheet_name = (
        "Данные"
        if "Данные" in excel.sheet_names
        else excel.sheet_names[0]
    )

    print(
        "\nParsing sheet:",
        sheet_name,
    )

    raw = pd.read_excel(
        io.BytesIO(content),
        sheet_name=sheet_name,
        header=None,
        engine="xlrd",
    )

    if raw.empty:

        raise RuntimeError(
            "Fedstat XLS is empty."
        )

    print(
        "Raw size:",
        raw.shape[0],
        "rows x",
        raw.shape[1],
        "columns",
    )

    # ========================================================
    # 1. FIND DATA ROW
    # ========================================================

    data_row = None

    territory_phrase = (
        "российская федерация без учета новых субъектов"
    )

    activity_phrase = (
        "всего по обследуемым видам экономической деятельности"
    )

    for idx in raw.index:

        row_text = " | ".join(
            normalize_text(value)
            for value in raw.iloc[idx].tolist()
            if clean_text(value)
        )

        if (
            territory_phrase in row_text
            and
            activity_phrase in row_text
        ):

            data_row = idx
            break

    # Some Fedstat layouts merge labels vertically and only
    # one of the two descriptions may remain on the value row.
    if data_row is None:

        for idx in raw.index:

            row_text = " | ".join(
                normalize_text(value)
                for value in raw.iloc[idx].tolist()
                if clean_text(value)
            )

            if (
                territory_phrase in row_text
                or
                activity_phrase in row_text
            ):

                numeric_count = sum(
                    as_number(value) is not None
                    for value in raw.iloc[idx].tolist()
                )

                if numeric_count >= 3:

                    data_row = idx
                    break

    if data_row is None:

        print(
            "\nFirst rows of downloaded file:"
        )

        print(
            raw.head(15).to_string(
                header=False
            )
        )

        raise RuntimeError(
            "Could not find the selected "
            "Russian Federation / total activity row."
        )

    print(
        "Detected data row:",
        data_row + 1,
    )

    # ========================================================
    # 2. FIND PERIOD HEADER ROW
    # ========================================================

    period_row = None
    best_period_count = 0

    for idx in range(
        0,
        data_row,
    ):

        count = 0

        for value in raw.iloc[idx]:

            if period_to_month(
                value
            ) is not None:

                count += 1

        if count > best_period_count:

            best_period_count = count
            period_row = idx

    if (
        period_row is None
        or best_period_count < 2
    ):

        raise RuntimeError(
            "Could not detect cumulative-period header row."
        )

    print(
        "Detected period row:",
        period_row + 1,
    )

    print(
        "Period cells detected:",
        best_period_count,
    )

    # ========================================================
    # 3. FIND YEAR HEADER ROW
    # ========================================================

    year_row = None
    best_year_count = 0

    for idx in range(
        0,
        period_row + 1,
    ):

        count = 0

        for value in raw.iloc[idx]:

            number = as_number(
                value
            )

            if (
                number is not None
                and
                FIRST_YEAR <= int(number) <= 2035
            ):

                count += 1

        if count > best_year_count:

            best_year_count = count
            year_row = idx

    if year_row is None:

        raise RuntimeError(
            "Could not detect year header row."
        )

    print(
        "Detected year row:",
        year_row + 1,
    )

    # ========================================================
    # 4. FORWARD-FILL YEAR HEADINGS
    #
    # Fedstat often merges a year heading over several columns,
    # so only the first cell contains 2023 and the next cells
    # are blank.
    # ========================================================

    years_by_column: dict[int, int] = {}

    current_year = None

    for col in range(
        raw.shape[1]
    ):

        value = raw.iloc[
            year_row,
            col,
        ]

        number = as_number(
            value
        )

        if (
            number is not None
            and
            FIRST_YEAR <= int(number) <= 2035
        ):

            current_year = int(
                number
            )

        if current_year is not None:

            years_by_column[
                col
            ] = current_year

    # ========================================================
    # 5. BUILD TIME SERIES
    # ========================================================

    rows = []

    for col in range(
        raw.shape[1]
    ):

        period_text = raw.iloc[
            period_row,
            col,
        ]

        month = period_to_month(
            period_text
        )

        if month is None:
            continue

        year = years_by_column.get(
            col
        )

        if year is None:
            continue

        if year < FIRST_YEAR:
            continue

        value = as_number(
            raw.iloc[
                data_row,
                col,
            ]
        )

        if value is None:
            continue

        date = pd.Timestamp(
            year=year,
            month=month,
            day=1,
        )

        rows.append(
            {
                "date": date,
                "loss_making_companies_share": value,
                "period": clean_text(period_text),
                "source": "fedstat_57746",
            }
        )

    df = pd.DataFrame(
        rows
    )

    if df.empty:

        raise RuntimeError(
            "No loss-making-company observations "
            "were extracted from Fedstat XLS."
        )

    df[
        "loss_making_companies_share"
    ] = (
        pd.to_numeric(
            df[
                "loss_making_companies_share"
            ],
            errors="coerce",
        )
        .round(2)
    )

    df = (
        df
        .sort_values(
            "date"
        )
        .drop_duplicates(
            "date",
            keep="last",
        )
        .reset_index(
            drop=True
        )
    )

    df["date"] = (
        df["date"]
        .dt.strftime(
            "%Y-%m-%d"
        )
    )

    return df


# ============================================================
# VALIDATION
# ============================================================

def validate(
    df: pd.DataFrame,
    page_years: list[int],
) -> None:

    if df.empty:

        raise RuntimeError(
            "Output is empty."
        )

    if df["date"].duplicated().any():

        raise RuntimeError(
            "Duplicate dates detected."
        )

    if df[
        "loss_making_companies_share"
    ].isna().any():

        raise RuntimeError(
            "Missing loss-making-company share values."
        )

    bad = df[
        ~df[
            "loss_making_companies_share"
        ].between(
            0,
            100,
        )
    ]

    if not bad.empty:

        raise RuntimeError(
            "Values outside 0-100% range:\n"
            + bad.to_string(
                index=False
            )
        )

    dates = pd.to_datetime(
        df["date"]
    )

    print(
        "\nFirst available observation:",
        dates.min().strftime("%Y-%m"),
    )

    latest_page_year = max(
        page_years
    )

    latest_data_year = int(
        dates.max().year
    )

    if latest_data_year < latest_page_year:

        print(
            "\nWARNING:"
        )

        print(
            "Fedstat page contains year",
            latest_page_year,
            "but the selected series ends in",
            latest_data_year,
        )

    # --------------------------------------------------------
    # We do NOT require December in the latest year.
    #
    # If the newest available observation is Jan-May,
    # Jan-June, etc., that is valid.
    # --------------------------------------------------------

    latest_date = dates.max()

    print(
        "\nLatest available observation:",
        latest_date.strftime(
            "%Y-%m"
        ),
    )


# ============================================================
# PRESERVE EXISTING HISTORY
# ============================================================

def merge_existing_history(
    current: pd.DataFrame,
) -> pd.DataFrame:

    if not OUTPUT_FILE.exists():

        return current

    print(
        "\nLoading existing history..."
    )

    existing = pd.read_csv(
        OUTPUT_FILE
    )

    required = {
        "date",
        "loss_making_companies_share",
    }

    if not required.issubset(
        existing.columns
    ):

        raise RuntimeError(
            "Existing CSV has unexpected structure."
        )

    current_dates = set(
        current["date"]
        .astype(str)
    )

    history_only = existing[
        ~existing["date"]
        .astype(str)
        .isin(
            current_dates
        )
    ].copy()

    result = pd.concat(
        [
            history_only,
            current,
        ],
        ignore_index=True,
        sort=False,
    )

    result["_date"] = pd.to_datetime(
        result["date"],
        errors="coerce",
    )

    result = (
        result
        .sort_values(
            "_date"
        )
        .drop_duplicates(
            "date",
            keep="last",
        )
        .drop(
            columns="_date"
        )
        .reset_index(
            drop=True
        )
    )

    return result


# ============================================================
# MAIN
# ============================================================

def load_existing_history_for_fallback() -> pd.DataFrame:
    """
    Used only when Fedstat blocks GitHub Actions with HTTP 403.

    The existing CSV is treated as preserved official history.
    We validate it before allowing the script to finish
    successfully.

    No files are rewritten in fallback mode.
    """

    print(
        "\nFedstat is unavailable from this environment."
    )

    print(
        "Checking preserved CSV history..."
    )

    if not OUTPUT_FILE.exists():

        raise RuntimeError(
            "Fedstat returned HTTP 403 and no existing "
            "loss-making companies CSV is available."
        )

    existing = pd.read_csv(
        OUTPUT_FILE
    )

    required_columns = {
        "date",
        "loss_making_companies_share",
    }

    missing_columns = (
        required_columns
        - set(existing.columns)
    )

    if missing_columns:

        raise RuntimeError(
            "Existing loss-making companies CSV "
            "has unexpected structure. Missing columns: "
            + ", ".join(
                sorted(missing_columns)
            )
        )

    if existing.empty:

        raise RuntimeError(
            "Existing loss-making companies CSV is empty."
        )

    dates = pd.to_datetime(
        existing["date"],
        errors="coerce",
    )

    if dates.isna().any():

        raise RuntimeError(
            "Existing loss-making companies CSV "
            "contains invalid dates."
        )

    if existing["date"].duplicated().any():

        raise RuntimeError(
            "Existing loss-making companies CSV "
            "contains duplicate dates."
        )

    values = pd.to_numeric(
        existing[
            "loss_making_companies_share"
        ],
        errors="coerce",
    )

    if values.isna().any():

        raise RuntimeError(
            "Existing loss-making companies CSV "
            "contains missing or non-numeric values."
        )

    bad_values = ~values.between(
        0,
        100,
    )

    if bad_values.any():

        raise RuntimeError(
            "Existing loss-making companies CSV "
            "contains values outside 0-100%."
        )

    latest_date = dates.max()

    print(
        "Preserved history is valid."
    )

    print(
        "Latest stored observation:",
        latest_date.strftime(
            "%Y-%m"
        ),
    )

    print(
        "No files will be changed."
    )

    print(
        "The workflow may continue using "
        "the last successfully downloaded Fedstat data."
    )

    return existing

def main():

    DATA_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    # ========================================================
    # DOWNLOAD CURRENT FEDSTAT DATA
    #
    # GitHub-hosted runners are sometimes blocked by Fedstat
    # with HTTP 403, while the same request works locally.
    #
    # ONLY HTTP 403 gets a fallback.
    # Any other error must still fail the workflow.
    # ========================================================

    try:

        content, page_years = (
            download_loss_making_xls()
        )

    except requests.exceptions.HTTPError as exc:

        status_code = None

        if exc.response is not None:

            status_code = (
                exc.response.status_code
            )

        if status_code != 403:

            raise

        print(
            "\n" + "=" * 72
        )

        print(
            "FEDSTAT HTTP 403"
        )

        print(
            "=" * 72
        )

        print(
            "\nFedstat blocked this request."
        )

        print(
            "This is allowed only as a 403 fallback."
        )

        existing = (
            load_existing_history_for_fallback()
        )

        print(
            "\n" + "=" * 72
        )

        print(
            "DONE WITH PRESERVED HISTORY"
        )

        print(
            "=" * 72
        )

        print(
            "\nOutput remains unchanged:"
        )

        print(
            OUTPUT_FILE
        )

        print(
            "\nRange:"
        )

        print(
            existing["date"].min(),
            "->",
            existing["date"].max(),
        )

        return

    # ========================================================
    # NORMAL FEDSTAT UPDATE
    # ========================================================

    print(
        "\nParsing Fedstat export..."
    )

    current = parse_loss_making_xls(
        content
    )

    validate(
        current,
        page_years,
    )

    print(
        "\nCurrent Fedstat data:"
    )

    print(
        current
        .tail(15)
        .to_string(
            index=False
        )
    )

    # ========================================================
    # PRESERVE OLD HISTORY + ADD CURRENT DATA
    # ========================================================

    result = merge_existing_history(
        current
    )

    # Validate merged output as well.
    validate(
        result,
        page_years,
    )

    result.to_csv(
        OUTPUT_FILE,
        index=False,
        encoding="utf-8",
    )

    latest_date = (
        pd.to_datetime(
            result["date"]
        )
        .max()
        .strftime(
            "%Y-%m-%d"
        )
    )

    SOURCE_INFO_FILE.write_text(
        "\n".join(
            [
                "Russia loss-making companies share",
                "source=EMISS / Fedstat",
                f"indicator_id={INDICATOR_ID}",
                f"indicator_url={INDICATOR_URL}",
                (
                    "indicator="
                    "Удельный вес убыточных организаций "
                    "с 2017 г."
                ),
                "unit=percent",
                (
                    "territory="
                    "Российская Федерация без учета "
                    "новых субъектов (с 01.01.2023)"
                ),
                (
                    "activity="
                    "Всего по обследуемым видам "
                    "экономической деятельности"
                ),
                (
                    "periodicity="
                    "cumulative from the beginning of year"
                ),
                (
                    "date_definition="
                    "date is assigned to the final month "
                    "of each cumulative period; "
                    "for example 2026-05-01 means "
                    "January-May 2026"
                ),
                f"first_year={FIRST_YEAR}",
                f"latest_date={latest_date}",
                "",
            ]
        ),
        encoding="utf-8",
    )

    print(
        "\n" + "=" * 72
    )

    print(
        "DONE"
    )

    print(
        "=" * 72
    )

    print(
        "\nOutput:"
    )

    print(
        OUTPUT_FILE
    )

    print(
        "\nSource info:"
    )

    print(
        SOURCE_INFO_FILE
    )

    print(
        "\nRange:"
    )

    print(
        result["date"].min(),
        "->",
        result["date"].max(),
    )

    print(
        "\nLast rows:"
    )

    print(
        result
        .tail(15)
        .to_string(
            index=False
        )
    )


if __name__ == "__main__":
    main()