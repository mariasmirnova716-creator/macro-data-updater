from __future__ import annotations

import io
import re
import ssl
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path

import certifi
import pandas as pd


# ============================================================
# PATHS
# ============================================================

PROJECT_DIR = Path(__file__).resolve().parents[2]

DATA_DIR = PROJECT_DIR / "data"
OUTPUT_DIR = DATA_DIR / "sheet_03a_keyrate_inflation"

ACTUAL_FILE = DATA_DIR / "cbr_inflation_monthly.csv"

PARAMETERS_OUTPUT_FILE = (
    OUTPUT_DIR / "cbr_medium_term_forecast_parameters.csv"
)

MONTHLY_OUTPUT_FILE = (
    OUTPUT_DIR / "cbr_keyrate_inflation_monthly_forecast.csv"
)

SOURCE_INFO_FILE = (
    OUTPUT_DIR / "cbr_medium_term_forecast_source.txt"
)


# ============================================================
# CBR
# ============================================================

CBR_HOST = "https://www.cbr.ru"

DECISION_MATERIALS_URL = (
    "https://www.cbr.ru/dkp/mp_dec/decision_key_rate/"
)


# ============================================================
# SETTINGS
# ============================================================

# Warning if our constructed monthly YoY path has an annual
# average that differs from the midpoint of the CBR's
# annual-average inflation forecast by more than this value.
INFLATION_WARNING_THRESHOLD_PP = 0.50

ROUND_DIGITS = 4


# ============================================================
# DATA TYPES
# ============================================================

@dataclass
class ForecastRange:
    low: float
    high: float

    @property
    def midpoint(self) -> float:
        return (self.low + self.high) / 2


# ============================================================
# DOWNLOAD HELPERS
# ============================================================

def create_ssl_context() -> ssl.SSLContext:
    return ssl.create_default_context(
        cafile=certifi.where()
    )


def download_bytes(url: str) -> tuple[bytes, str]:

    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 "
                "(Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 "
                "Chrome/126 Safari/537.36"
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


def download_text(url: str) -> tuple[str, str]:

    raw_bytes, final_url = download_bytes(url)

    return (
        raw_bytes.decode(
            "utf-8",
            errors="replace",
        ),
        final_url,
    )


# ============================================================
# TEXT HELPERS
# ============================================================

def clean_text(value: object) -> str:

    if pd.isna(value):
        return ""

    text = str(value)

    text = (
        text
        .replace("\xa0", " ")
        .replace("\u202f", " ")
        .replace("ё", "е")
    )

    return re.sub(
        r"\s+",
        " ",
        text,
    ).strip()


def normalize_text(value: object) -> str:
    return clean_text(value).lower()


def normalize_html(html: str) -> str:

    return (
        html
        .replace("&amp;", "&")
        .replace("&#x2F;", "/")
        .replace("&#47;", "/")
        .replace("\\u0026", "&")
    )


# ============================================================
# FIND LATEST CBR COMMENT PAGE
# ============================================================

def comment_date_from_url(url: str) -> pd.Timestamp:

    match = re.search(
        r"comment_(\d{2})(\d{2})(\d{4})",
        url,
        flags=re.IGNORECASE,
    )

    if not match:
        return pd.Timestamp.min

    day, month, year = map(
        int,
        match.groups(),
    )

    return pd.Timestamp(
        year=year,
        month=month,
        day=day,
    )


def find_latest_comment_page() -> str:

    print("Opening CBR decision materials page:")
    print(DECISION_MATERIALS_URL)

    html, final_url = download_text(
        DECISION_MATERIALS_URL
    )

    html = normalize_html(html)

    raw_links = re.findall(
        r'href=["\']([^"\']*'
        r'comment_\d{8}/?[^"\']*)["\']',
        html,
        flags=re.IGNORECASE,
    )

    links: list[str] = []

    for raw_link in raw_links:

        full_url = urllib.parse.urljoin(
            final_url,
            raw_link,
        )

        if full_url not in links:
            links.append(full_url)

    if not links:
        raise RuntimeError(
            "No 'Комментарий к среднесрочному прогнозу' "
            "pages were found on the CBR materials page."
        )

    links.sort(
        key=comment_date_from_url,
        reverse=True,
    )

    latest_url = links[0]

    print()
    print("Latest available CBR forecast comment:")
    print(latest_url)

    return latest_url


# ============================================================
# FIND XLSX ON COMMENT PAGE
# ============================================================

def find_xlsx_url(comment_page_url: str) -> str:

    html, final_url = download_text(
        comment_page_url
    )

    html = normalize_html(html)

    # First preference: any link containing comment_graph.
    matches = re.findall(
        r'href=["\']([^"\']*comment_graph[^"\']*'
        r'\.xlsx(?:\?[^"\']*)?)["\']',
        html,
        flags=re.IGNORECASE,
    )

    if matches:
        return urllib.parse.urljoin(
            final_url,
            matches[0],
        )

    # Fallback: any XLSX on the page.
    matches = re.findall(
        r'href=["\']([^"\']+\.xlsx(?:\?[^"\']*)?)["\']',
        html,
        flags=re.IGNORECASE,
    )

    if matches:
        return urllib.parse.urljoin(
            final_url,
            matches[0],
        )

    raise RuntimeError(
        "The latest CBR forecast comment exists, "
        "but its XLSX 'Графики и таблицы' file "
        "was not found."
    )


def download_forecast_xlsx(
    xlsx_url: str,
) -> tuple[bytes, str]:

    print()
    print("Downloading CBR XLSX:")
    print(xlsx_url)

    raw_bytes, final_url = download_bytes(
        xlsx_url
    )

    if not raw_bytes.startswith(b"PK"):

        beginning = raw_bytes[:200].decode(
            "utf-8",
            errors="replace",
        )

        raise RuntimeError(
            "CBR returned something other than XLSX. "
            f"Response begins with: {beginning!r}"
        )

    return raw_bytes, final_url


# ============================================================
# FORECAST TABLE PARSING
# ============================================================

def parse_forecast_range(
    value: object,
) -> ForecastRange | None:

    if pd.isna(value):
        return None

    if isinstance(value, (int, float)):
        number = float(value)

        return ForecastRange(
            low=number,
            high=number,
        )

    text = clean_text(value)

    if not text:
        return None

    text = (
        text
        .replace(",", ".")
        .replace("%", "")
    )

    # Remove superscript footnote marks.
    text = re.sub(
        r"[¹²³⁴⁵⁶⁷⁸⁹⁰]+",
        "",
        text,
    )

    # First try to parse a forecast range.
    # Dash here is a separator, not a minus sign.
    range_match = re.search(
        r"(-?\d+(?:\.\d+)?)\s*[–—−-]\s*(\d+(?:\.\d+)?)",
        text,
    )

    if range_match:

        low = float(
            range_match.group(1)
        )

        high = float(
            range_match.group(2)
        )

        return ForecastRange(
            low=low,
            high=high,
        )

    # Otherwise parse a single value.
    single_match = re.search(
        r"-?\d+(?:\.\d+)?",
        text,
    )

    if not single_match:
        return None

    number = float(
        single_match.group(0)
    )

    return ForecastRange(
        low=number,
        high=number,
    )

def normalize_key_rate_range(
    value: ForecastRange | None,
) -> ForecastRange | None:

    if value is None:
        return None

    def normalize_number(number: float) -> float:
        """
        CBR medium-term key-rate forecast values
        are published with one decimal place.

        This also protects against Excel footnotes
        being read as part of the number, e.g.
        14.6¹ -> 14.61.
        """
        return round(number, 1)

    low = normalize_number(value.low)
    high = normalize_number(value.high)

    return ForecastRange(
        low=low,
        high=high,
    )


def find_table_1(
    excel_bytes: bytes,
) -> pd.DataFrame:

    workbook = pd.read_excel(
        io.BytesIO(excel_bytes),
        sheet_name=None,
        header=None,
        engine="openpyxl",
    )

    # Prefer actual sheet name.
    for sheet_name, frame in workbook.items():

        normalized_name = normalize_text(
            sheet_name
        )

        if (
            normalized_name in {
                "табл 1",
                "табл. 1",
                "табл1",
            }
        ):
            print(
                f"Using sheet: {sheet_name}"
            )
            return frame

    # Fallback by content.
    for sheet_name, frame in workbook.items():

        sample = " | ".join(
            normalize_text(value)
            for value
            in frame.head(25)
            .to_numpy()
            .ravel()
        )

        if (
            "основные параметры прогноза" in sample
            and "ключевая ставка" in sample
            and "инфляция" in sample
        ):
            print(
                f"Table 1 found on sheet: {sheet_name}"
            )
            return frame

    raise RuntimeError(
        "Could not find 'Табл 1' in the CBR XLSX."
    )


def find_year_columns(
    frame: pd.DataFrame,
) -> dict[int, int]:

    result: dict[int, int] = {}

    for row_index in range(
        min(len(frame), 20)
    ):

        for column_index in range(
            frame.shape[1]
        ):

            text = clean_text(
                frame.iat[
                    row_index,
                    column_index,
                ]
            )

            match = re.search(
                r"\b(20\d{2})\b",
                text,
            )

            if match:

                year = int(
                    match.group(1)
                )

                if year not in result:
                    result[year] = column_index

    if not result:
        raise RuntimeError(
            "No year columns found in CBR Table 1."
        )

    return dict(
        sorted(result.items())
    )


def find_indicator_row(
    frame: pd.DataFrame,
    required_fragments: tuple[str, ...],
) -> int:

    for row_index in range(
        min(len(frame), 40)
    ):

        joined = " | ".join(
            normalize_text(value)
            for value
            in frame.iloc[row_index].tolist()
        )

        if all(
            fragment in joined
            for fragment
            in required_fragments
        ):
            return row_index

    raise RuntimeError(
        "Indicator row not found. "
        f"Required fragments: {required_fragments}"
    )


def parse_forecast_parameters(
    excel_bytes: bytes,
) -> pd.DataFrame:

    frame = find_table_1(
        excel_bytes
    )

    year_columns = find_year_columns(
        frame
    )

    inflation_dec_row = find_indicator_row(
        frame,
        (
            "инфляция",
            "декабр",
        ),
    )

    inflation_avg_row = find_indicator_row(
        frame,
        (
            "инфляция",
            "среднем за год",
        ),
    )

    key_rate_avg_row = find_indicator_row(
        frame,
        (
            "ключевая ставка",
            "среднем за год",
        ),
    )

    records: list[dict[str, object]] = []

    for year, column_index in (
        year_columns.items()
    ):

        inflation_dec = parse_forecast_range(
            frame.iat[
                inflation_dec_row,
                column_index,
            ]
        )

        inflation_avg = parse_forecast_range(
            frame.iat[
                inflation_avg_row,
                column_index,
            ]
        )

        raw_key_rate_avg = parse_forecast_range(
            frame.iat[
                key_rate_avg_row,
                column_index,
            ]
        )

        if raw_key_rate_avg is not None:
            if (
                    abs(
                        raw_key_rate_avg.low
                        - round(raw_key_rate_avg.low, 1)
                    ) > 1e-9
                    or
                    abs(
                        raw_key_rate_avg.high
                        - round(raw_key_rate_avg.high, 1)
                    ) > 1e-9
            ):
                print(
                    f"WARNING: suspicious key-rate precision "
                    f"for {year}: "
                    f"{raw_key_rate_avg.low}-"
                    f"{raw_key_rate_avg.high}. "
                    "Possible Excel footnote detected."
                )

        key_rate_avg = normalize_key_rate_range(
            raw_key_rate_avg
        )

        record: dict[str, object] = {
            "year": year,
        }

        if inflation_dec is not None:

            record.update(
                {
                    "inflation_dec_low":
                        inflation_dec.low,
                    "inflation_dec_high":
                        inflation_dec.high,
                    "inflation_dec_mid":
                        inflation_dec.midpoint,
                }
            )

        if inflation_avg is not None:

            record.update(
                {
                    "inflation_avg_low":
                        inflation_avg.low,
                    "inflation_avg_high":
                        inflation_avg.high,
                    "inflation_avg_mid":
                        inflation_avg.midpoint,
                }
            )

        if key_rate_avg is not None:

            record.update(
                {
                    "key_rate_avg_low":
                        key_rate_avg.low,
                    "key_rate_avg_high":
                        key_rate_avg.high,
                    "key_rate_avg_mid":
                        key_rate_avg.midpoint,
                }
            )

        records.append(record)

    result = pd.DataFrame(records)

    if result.empty:
        raise RuntimeError(
            "No forecast parameters were parsed "
            "from CBR Table 1."
        )

    return (
        result
        .sort_values("year")
        .reset_index(drop=True)
    )


# ============================================================
# ACTUAL DATA
# ============================================================

def load_actual_data() -> pd.DataFrame:

    if not ACTUAL_FILE.exists():

        raise FileNotFoundError(
            f"Actual data file not found: {ACTUAL_FILE}"
        )

    frame = pd.read_csv(
        ACTUAL_FILE
    )

    required_columns = {
        "date",
        "key_rate",
        "russia_inflation_yoy",
        "inflation_target",
    }

    missing = (
        required_columns
        - set(frame.columns)
    )

    if missing:

        raise RuntimeError(
            "Actual CBR file is missing columns: "
            f"{sorted(missing)}"
        )

    frame["date"] = pd.to_datetime(
        frame["date"],
        errors="coerce",
    )

    frame = frame.dropna(
        subset=["date"]
    )

    for column in (
        "key_rate",
        "russia_inflation_yoy",
        "inflation_target",
    ):

        frame[column] = pd.to_numeric(
            frame[column],
            errors="coerce",
        )

    return (
        frame
        .sort_values("date")
        .drop_duplicates(
            subset=["date"],
            keep="last",
        )
        .reset_index(drop=True)
    )


# ============================================================
# FORECAST HORIZON
# ============================================================

def forecast_start_year(
    latest_actual_date: pd.Timestamp,
) -> int:
    """
    Forecast begins in:
    - the same year if that year is incomplete;
    - the next year if December is already actual.
    """

    if latest_actual_date.month < 12:
        return int(latest_actual_date.year)

    return int(latest_actual_date.year + 1)


def select_two_forecast_years(
    parameters: pd.DataFrame,
    start_year: int,
    indicator_name: str,
) -> pd.DataFrame:

    end_year = start_year + 1

    result = (
        parameters[
            (parameters["year"] >= start_year)
            & (parameters["year"] <= end_year)
        ]
        .copy()
        .sort_values("year")
    )

    years_found = set(
        result["year"].astype(int)
    )

    required = {
        start_year,
        end_year,
    }

    missing = required - years_found

    if missing:

        raise RuntimeError(
            f"CBR forecast does not contain "
            f"the required {indicator_name} years: "
            f"{sorted(missing)}"
        )

    return result


# ============================================================
# DATE HELPERS
# ============================================================

def month_end(
    year: int,
    month: int,
) -> pd.Timestamp:

    return pd.Timestamp(
        year=year,
        month=month,
        day=1,
    ) + pd.offsets.MonthEnd(0)


def year_month_ends(
    year: int,
) -> pd.DatetimeIndex:

    return pd.date_range(
        start=f"{year}-01-31",
        periods=12,
        freq="ME",
    )


# ============================================================
# KEY RATE FORECAST
# ============================================================

def validate_current_year_actuals(
    actual_year: pd.DataFrame,
    value_column: str,
    latest_month: int,
    indicator_name: str,
) -> None:

    months_present = set(
        actual_year.loc[
            actual_year[value_column].notna(),
            "date",
        ].dt.month.astype(int)
    )

    expected = set(
        range(1, latest_month + 1)
    )

    missing = expected - months_present

    if missing:

        raise RuntimeError(
            f"{indicator_name}: missing actual months "
            f"before the latest observation: "
            f"{sorted(missing)}"
        )


def forecast_key_rate_current_year(
    actual: pd.DataFrame,
    year: int,
    target_annual_average: float,
) -> dict[pd.Timestamp, float]:

    year_data = (
        actual[
            actual["date"].dt.year == year
        ]
        .dropna(
            subset=["key_rate"]
        )
        .copy()
        .sort_values("date")
    )

    if year_data.empty:

        raise RuntimeError(
            f"No actual key-rate data found for {year}."
        )

    latest_date = year_data["date"].max()
    latest_month = int(latest_date.month)

    validate_current_year_actuals(
        actual_year=year_data,
        value_column="key_rate",
        latest_month=latest_month,
        indicator_name="Key rate",
    )

    if latest_month >= 12:
        return {}

    actual_values = (
        year_data[
            year_data["date"].dt.month
            <= latest_month
        ]["key_rate"]
        .astype(float)
        .tolist()
    )

    last_actual = float(
        year_data.loc[
            year_data["date"] == latest_date,
            "key_rate",
        ].iloc[-1]
    )

    months_remaining = 12 - latest_month

    required_forecast_sum = (
        target_annual_average * 12
        - sum(actual_values)
    )

    # Forecast path:
    # next month = last_actual + d
    # following   = last_actual + 2d
    # ...
    # December    = last_actual + n*d

    weight_sum = (
        months_remaining
        * (months_remaining + 1)
        / 2
    )

    step = (
        required_forecast_sum
        - months_remaining * last_actual
    ) / weight_sum

    result: dict[pd.Timestamp, float] = {}

    for i in range(
        1,
        months_remaining + 1,
    ):

        month_number = (
            latest_month + i
        )

        result[
            month_end(
                year,
                month_number,
            )
        ] = (
            last_actual
            + step * i
        )

    return result


def forecast_key_rate_future_year(
    year: int,
    previous_december: float,
    target_annual_average: float,
) -> dict[pd.Timestamp, float]:

    # Our methodological assumption:
    # January equals previous December.
    january = previous_december

    # Arithmetic sequence:
    # annual average = (January + December) / 2
    december = (
        2 * target_annual_average
        - january
    )

    step = (
        december - january
    ) / 11

    result: dict[pd.Timestamp, float] = {}

    for month_index, date_value in enumerate(
        year_month_ends(year)
    ):

        result[
            pd.Timestamp(date_value)
        ] = (
            january
            + step * month_index
        )

    return result


# ============================================================
# INFLATION FORECAST
# ============================================================

def forecast_inflation_current_year(
    actual: pd.DataFrame,
    year: int,
    december_target: float,
) -> dict[pd.Timestamp, float]:

    year_data = (
        actual[
            actual["date"].dt.year == year
        ]
        .dropna(
            subset=["russia_inflation_yoy"]
        )
        .copy()
        .sort_values("date")
    )

    if year_data.empty:

        raise RuntimeError(
            f"No actual inflation data found for {year}."
        )

    latest_date = year_data["date"].max()
    latest_month = int(latest_date.month)

    validate_current_year_actuals(
        actual_year=year_data,
        value_column="russia_inflation_yoy",
        latest_month=latest_month,
        indicator_name="Inflation",
    )

    if latest_month >= 12:
        return {}

    last_actual = float(
        year_data.loc[
            year_data["date"] == latest_date,
            "russia_inflation_yoy",
        ].iloc[-1]
    )

    months_remaining = (
        12 - latest_month
    )

    step = (
        december_target
        - last_actual
    ) / months_remaining

    result: dict[pd.Timestamp, float] = {}

    for i in range(
        1,
        months_remaining + 1,
    ):

        month_number = (
            latest_month + i
        )

        result[
            month_end(
                year,
                month_number,
            )
        ] = (
            last_actual
            + step * i
        )

    return result


def forecast_inflation_future_year(
    year: int,
    previous_december: float,
    december_target: float,
) -> dict[pd.Timestamp, float]:

    # Our methodological assumption:
    # January equals previous December.
    january = previous_december

    step = (
        december_target
        - january
    ) / 11

    result: dict[pd.Timestamp, float] = {}

    for month_index, date_value in enumerate(
        year_month_ends(year)
    ):

        result[
            pd.Timestamp(date_value)
        ] = (
            january
            + step * month_index
        )

    return result


# ============================================================
# BUILD KEY-RATE PATH
# ============================================================

def build_key_rate_forecasts(
    actual: pd.DataFrame,
    parameters: pd.DataFrame,
) -> tuple[
    dict[pd.Timestamp, float],
    int,
    int,
]:

    actual_key = (
        actual
        .dropna(subset=["key_rate"])
        .copy()
        .sort_values("date")
    )

    if actual_key.empty:
        raise RuntimeError(
            "No actual key-rate observations found."
        )

    latest_actual_date = (
        actual_key["date"].max()
    )

    start_year = forecast_start_year(
        latest_actual_date
    )

    end_year = start_year + 1

    selected = select_two_forecast_years(
        parameters,
        start_year,
        "key-rate forecast",
    )

    result: dict[pd.Timestamp, float] = {}

    previous_december: float | None = None

    for _, parameter_row in (
        selected.iterrows()
    ):

        year = int(
            parameter_row["year"]
        )

        target_average = (
            parameter_row.get(
                "key_rate_avg_mid"
            )
        )

        if pd.isna(target_average):

            raise RuntimeError(
                f"No key-rate annual-average "
                f"forecast found for {year}."
            )

        target_average = float(
            target_average
        )

        current_incomplete_year = (
            year == latest_actual_date.year
            and latest_actual_date.month < 12
        )

        if current_incomplete_year:

            year_forecast = (
                forecast_key_rate_current_year(
                    actual=actual,
                    year=year,
                    target_annual_average=(
                        target_average
                    ),
                )
            )

            result.update(
                year_forecast
            )

            december_date = (
                month_end(year, 12)
            )

            previous_december = (
                result[december_date]
            )

        else:

            if previous_december is None:

                previous_december_row = (
                    actual_key[
                        actual_key["date"]
                        == month_end(
                            year - 1,
                            12,
                        )
                    ]
                )

                if previous_december_row.empty:

                    raise RuntimeError(
                        "Cannot start future key-rate "
                        f"forecast for {year}: "
                        "previous December is unavailable."
                    )

                previous_december = float(
                    previous_december_row[
                        "key_rate"
                    ].iloc[-1]
                )

            year_forecast = (
                forecast_key_rate_future_year(
                    year=year,
                    previous_december=(
                        previous_december
                    ),
                    target_annual_average=(
                        target_average
                    ),
                )
            )

            result.update(
                year_forecast
            )

            previous_december = result[
                month_end(
                    year,
                    12,
                )
            ]

    return (
        result,
        start_year,
        end_year,
    )


# ============================================================
# BUILD INFLATION PATH
# ============================================================

def build_inflation_forecasts(
    actual: pd.DataFrame,
    parameters: pd.DataFrame,
) -> tuple[
    dict[pd.Timestamp, float],
    int,
    int,
]:

    actual_inflation = (
        actual
        .dropna(
            subset=["russia_inflation_yoy"]
        )
        .copy()
        .sort_values("date")
    )

    if actual_inflation.empty:

        raise RuntimeError(
            "No actual inflation observations found."
        )

    latest_actual_date = (
        actual_inflation["date"].max()
    )

    start_year = forecast_start_year(
        latest_actual_date
    )

    end_year = start_year + 1

    selected = select_two_forecast_years(
        parameters,
        start_year,
        "inflation forecast",
    )

    result: dict[pd.Timestamp, float] = {}

    previous_december: float | None = None

    for _, parameter_row in (
        selected.iterrows()
    ):

        year = int(
            parameter_row["year"]
        )

        december_target = (
            parameter_row.get(
                "inflation_dec_mid"
            )
        )

        if pd.isna(december_target):

            raise RuntimeError(
                "No December-to-December inflation "
                f"forecast found for {year}."
            )

        december_target = float(
            december_target
        )

        current_incomplete_year = (
            year == latest_actual_date.year
            and latest_actual_date.month < 12
        )

        if current_incomplete_year:

            year_forecast = (
                forecast_inflation_current_year(
                    actual=actual,
                    year=year,
                    december_target=(
                        december_target
                    ),
                )
            )

            result.update(
                year_forecast
            )

            previous_december = result[
                month_end(
                    year,
                    12,
                )
            ]

        else:

            if previous_december is None:

                previous_december_row = (
                    actual_inflation[
                        actual_inflation["date"]
                        == month_end(
                            year - 1,
                            12,
                        )
                    ]
                )

                if previous_december_row.empty:

                    raise RuntimeError(
                        "Cannot start future inflation "
                        f"forecast for {year}: "
                        "previous December is unavailable."
                    )

                previous_december = float(
                    previous_december_row[
                        "russia_inflation_yoy"
                    ].iloc[-1]
                )

            year_forecast = (
                forecast_inflation_future_year(
                    year=year,
                    previous_december=(
                        previous_december
                    ),
                    december_target=(
                        december_target
                    ),
                )
            )

            result.update(
                year_forecast
            )

            previous_december = result[
                month_end(
                    year,
                    12,
                )
            ]

    return (
        result,
        start_year,
        end_year,
    )


# ============================================================
# FINAL MONTHLY TABLE
# ============================================================

def build_monthly_output(
    actual: pd.DataFrame,
    parameters: pd.DataFrame,
) -> pd.DataFrame:

    (
        key_forecast,
        key_start_year,
        key_end_year,
    ) = build_key_rate_forecasts(
        actual,
        parameters,
    )

    (
        inflation_forecast,
        inflation_start_year,
        inflation_end_year,
    ) = build_inflation_forecasts(
        actual,
        parameters,
    )

    latest_output_year = max(
        key_end_year,
        inflation_end_year,
    )

    earliest_actual_date = (
        actual["date"].min()
    )

    final_date = month_end(
        latest_output_year,
        12,
    )

    dates = pd.date_range(
        start=earliest_actual_date,
        end=final_date,
        freq="ME",
    )

    result = pd.DataFrame(
        {
            "date": dates,
        }
    )

    result = result.merge(
        actual[
            [
                "date",
                "key_rate",
                "russia_inflation_yoy",
                "inflation_target",
            ]
        ],
        on="date",
        how="left",
    )

    result["year"] = (
        result["date"].dt.year
    )

    result["month"] = (
        result["date"].dt.month
    )

    # --------------------------------------------------------
    # Separate actual and forecast values
    # --------------------------------------------------------

    result[
        "key_rate_actual"
    ] = result["key_rate"]

    result[
        "inflation_actual_yoy"
    ] = result["russia_inflation_yoy"]

    result[
        "key_rate_forecast"
    ] = result["date"].map(
        key_forecast
    )

    result[
        "inflation_forecast_yoy"
    ] = result["date"].map(
        inflation_forecast
    )

    # Complete technical series:
    # actual takes priority; forecast fills future months.
    result[
        "key_rate_series"
    ] = result[
        "key_rate_actual"
    ].combine_first(
        result["key_rate_forecast"]
    )

    result[
        "inflation_series_yoy"
    ] = result[
        "inflation_actual_yoy"
    ].combine_first(
        result["inflation_forecast_yoy"]
    )

    result[
        "key_rate_status"
    ] = ""

    result.loc[
        result["key_rate_actual"].notna(),
        "key_rate_status",
    ] = "actual"

    result.loc[
        result["key_rate_actual"].isna()
        & result["key_rate_forecast"].notna(),
        "key_rate_status",
    ] = "forecast"

    result[
        "inflation_status"
    ] = ""

    result.loc[
        result["inflation_actual_yoy"].notna(),
        "inflation_status",
    ] = "actual"

    result.loc[
        result["inflation_actual_yoy"].isna()
        & result["inflation_forecast_yoy"].notna(),
        "inflation_status",
    ] = "forecast"

    # Inflation target.
    result[
        "inflation_target"
    ] = (
        result[
            "inflation_target"
        ]
        .ffill()
    )

    # If future rows still have no target,
    # use the current 4% target.
    result[
        "inflation_target"
    ] = result[
        "inflation_target"
    ].fillna(4.0)

    # --------------------------------------------------------
    # Forecast corridors
    # --------------------------------------------------------

    result[
        "key_rate_forecast_low"
    ] = pd.NA

    result[
        "key_rate_forecast_high"
    ] = pd.NA

    result[
        "key_rate_corridor_width"
    ] = pd.NA

    result[
        "inflation_forecast_low"
    ] = pd.NA

    result[
        "inflation_forecast_high"
    ] = pd.NA

    result[
        "inflation_corridor_width"
    ] = pd.NA

    for _, row in parameters.iterrows():

        year = int(
            row["year"]
        )

        # KEY RATE
        if (
            key_start_year
            <= year
            <= key_end_year
        ):

            low = row.get(
                "key_rate_avg_low"
            )

            high = row.get(
                "key_rate_avg_high"
            )

            if (
                not pd.isna(low)
                and not pd.isna(high)
            ):

                half_width = (
                    float(high)
                    - float(low)
                ) / 2

                mask = (
                    (result["year"] == year)
                    & (
                        result[
                            "key_rate_forecast"
                        ].notna()
                    )
                )

                result.loc[
                    mask,
                    "key_rate_forecast_low",
                ] = (
                    result.loc[
                        mask,
                        "key_rate_forecast",
                    ]
                    - half_width
                )

                result.loc[
                    mask,
                    "key_rate_forecast_high",
                ] = (
                    result.loc[
                        mask,
                        "key_rate_forecast",
                    ]
                    + half_width
                )

                result.loc[
                    mask,
                    "key_rate_corridor_width",
                ] = (
                    float(high)
                    - float(low)
                )

        # INFLATION
        if (
            inflation_start_year
            <= year
            <= inflation_end_year
        ):

            low = row.get(
                "inflation_dec_low"
            )

            high = row.get(
                "inflation_dec_high"
            )

            if (
                not pd.isna(low)
                and not pd.isna(high)
            ):

                half_width = (
                    float(high)
                    - float(low)
                ) / 2

                mask = (
                    (result["year"] == year)
                    & (
                        result[
                            "inflation_forecast_yoy"
                        ].notna()
                    )
                )

                result.loc[
                    mask,
                    "inflation_forecast_low",
                ] = (
                    result.loc[
                        mask,
                        "inflation_forecast_yoy",
                    ]
                    - half_width
                )

                result.loc[
                    mask,
                    "inflation_forecast_high",
                ] = (
                    result.loc[
                        mask,
                        "inflation_forecast_yoy",
                    ]
                    + half_width
                )

                result.loc[
                    mask,
                    "inflation_corridor_width",
                ] = (
                    float(high)
                    - float(low)
                )

    # --------------------------------------------------------
    # Validation columns
    # --------------------------------------------------------

    result[
        "key_rate_target_annual_avg"
    ] = pd.NA

    result[
        "key_rate_calculated_annual_avg"
    ] = pd.NA

    result[
        "key_rate_avg_difference_pp"
    ] = pd.NA

    result[
        "inflation_target_annual_avg"
    ] = pd.NA

    result[
        "inflation_calculated_annual_avg"
    ] = pd.NA

    result[
        "inflation_avg_difference_pp"
    ] = pd.NA

    result[
        "inflation_warning"
    ] = ""

    for _, parameter_row in (
        parameters.iterrows()
    ):

        year = int(
            parameter_row["year"]
        )

        year_mask = (
            result["year"] == year
        )

        # ---------------- KEY RATE ----------------

        if (
            key_start_year
            <= year
            <= key_end_year
        ):

            target = (
                parameter_row.get(
                    "key_rate_avg_mid"
                )
            )

            values = (
                result.loc[
                    year_mask,
                    "key_rate_series",
                ]
                .dropna()
                .astype(float)
            )

            if (
                not pd.isna(target)
                and len(values) == 12
            ):

                calculated = float(
                    values.mean()
                )

                difference = (
                    calculated
                    - float(target)
                )

                result.loc[
                    year_mask,
                    "key_rate_target_annual_avg",
                ] = float(target)

                result.loc[
                    year_mask,
                    "key_rate_calculated_annual_avg",
                ] = calculated

                result.loc[
                    year_mask,
                    "key_rate_avg_difference_pp",
                ] = difference

        # ---------------- INFLATION ----------------

        if (
            inflation_start_year
            <= year
            <= inflation_end_year
        ):

            target = (
                parameter_row.get(
                    "inflation_avg_mid"
                )
            )

            values = (
                result.loc[
                    year_mask,
                    "inflation_series_yoy",
                ]
                .dropna()
                .astype(float)
            )

            if (
                not pd.isna(target)
                and len(values) == 12
            ):

                calculated = float(
                    values.mean()
                )

                difference = (
                    calculated
                    - float(target)
                )

                warning = ""

                if (
                    abs(difference)
                    >
                    INFLATION_WARNING_THRESHOLD_PP
                ):

                    warning = (
                        "WARNING: monthly inflation "
                        "trajectory annual average differs "
                        "from CBR annual-average midpoint "
                        f"by {difference:+.2f} pp"
                    )

                result.loc[
                    year_mask,
                    "inflation_target_annual_avg",
                ] = float(target)

                result.loc[
                    year_mask,
                    "inflation_calculated_annual_avg",
                ] = calculated

                result.loc[
                    year_mask,
                    "inflation_avg_difference_pp",
                ] = difference

                result.loc[
                    year_mask,
                    "inflation_warning",
                ] = warning

    # --------------------------------------------------------
    # Horizon metadata
    # --------------------------------------------------------

    result[
        "key_rate_forecast_start_year"
    ] = key_start_year

    result[
        "key_rate_forecast_end_year"
    ] = key_end_year

    result[
        "inflation_forecast_start_year"
    ] = inflation_start_year

    result[
        "inflation_forecast_end_year"
    ] = inflation_end_year

    # --------------------------------------------------------
    # Final cleanup
    # --------------------------------------------------------

    numeric_columns = [
        "key_rate",
        "russia_inflation_yoy",
        "inflation_target",
        "key_rate_actual",
        "inflation_actual_yoy",
        "key_rate_forecast",
        "inflation_forecast_yoy",
        "key_rate_series",
        "inflation_series_yoy",
        "key_rate_forecast_low",
        "key_rate_forecast_high",
        "key_rate_corridor_width",
        "inflation_forecast_low",
        "inflation_forecast_high",
        "inflation_corridor_width",
        "key_rate_target_annual_avg",
        "key_rate_calculated_annual_avg",
        "key_rate_avg_difference_pp",
        "inflation_target_annual_avg",
        "inflation_calculated_annual_avg",
        "inflation_avg_difference_pp",
    ]

    for column in numeric_columns:

        result[column] = pd.to_numeric(
            result[column],
            errors="coerce",
        ).round(
            ROUND_DIGITS
        )

    result["date"] = (
        result["date"]
        .dt.strftime("%Y-%m-%d")
    )

    return result


# ============================================================
# SAVE OUTPUTS
# ============================================================

def save_outputs(
    parameters: pd.DataFrame,
    monthly: pd.DataFrame,
    comment_url: str,
    xlsx_url: str,
) -> None:

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    parameters.to_csv(
        PARAMETERS_OUTPUT_FILE,
        index=False,
        encoding="utf-8-sig",
    )

    monthly.to_csv(
        MONTHLY_OUTPUT_FILE,
        index=False,
        encoding="utf-8-sig",
    )

    key_forecast_rows = monthly[
        monthly["key_rate_status"]
        == "forecast"
    ]

    inflation_forecast_rows = monthly[
        monthly["inflation_status"]
        == "forecast"
    ]

    source_lines = [
        f"comment_page={comment_url}",
        f"xlsx_url={xlsx_url}",
        f"updated_at={pd.Timestamp.now().isoformat()}",
        (
            "actual_file="
            f"{ACTUAL_FILE.relative_to(PROJECT_DIR)}"
        ),
        (
            "parameters_file="
            f"{PARAMETERS_OUTPUT_FILE.relative_to(PROJECT_DIR)}"
        ),
        (
            "monthly_file="
            f"{MONTHLY_OUTPUT_FILE.relative_to(PROJECT_DIR)}"
        ),
        (
            "key_rate_forecast_from="
            + (
                key_forecast_rows["date"].min()
                if not key_forecast_rows.empty
                else ""
            )
        ),
        (
            "key_rate_forecast_to="
            + (
                key_forecast_rows["date"].max()
                if not key_forecast_rows.empty
                else ""
            )
        ),
        (
            "inflation_forecast_from="
            + (
                inflation_forecast_rows["date"].min()
                if not inflation_forecast_rows.empty
                else ""
            )
        ),
        (
            "inflation_forecast_to="
            + (
                inflation_forecast_rows["date"].max()
                if not inflation_forecast_rows.empty
                else ""
            )
        ),
        (
            "method_key_rate_current_year="
            "keep all actual months; calculate a linear "
            "remaining-month path so Jan-Dec average equals "
            "the midpoint of the CBR annual-average range"
        ),
        (
            "method_key_rate_future_year="
            "January equals previous December; calculate a "
            "linear Jan-Dec path whose annual average equals "
            "the midpoint of the CBR annual-average range"
        ),
        (
            "method_inflation_current_year="
            "keep actual YoY inflation; linearly interpolate "
            "from the last actual month to the midpoint of "
            "the CBR December-to-December forecast"
        ),
        (
            "method_inflation_future_year="
            "January equals previous December; linearly "
            "interpolate to the midpoint of the CBR "
            "December-to-December forecast"
        ),
        (
            "forecast_horizon="
            "first incomplete/next forecast year plus one year"
        ),
        (
            "inflation_warning_threshold_pp="
            f"{INFLATION_WARNING_THRESHOLD_PP}"
        ),
    ]

    SOURCE_INFO_FILE.write_text(
        "\n".join(source_lines) + "\n",
        encoding="utf-8",
    )


# ============================================================
# MAIN
# ============================================================

def main() -> None:

    print("=" * 60)
    print("CBR medium-term forecast updater — Sheet 3A")
    print("=" * 60)

    comment_url = (
        find_latest_comment_page()
    )

    xlsx_url = find_xlsx_url(
        comment_url
    )

    excel_bytes, final_xlsx_url = (
        download_forecast_xlsx(
            xlsx_url
        )
    )

    print()
    print("Parsing CBR Table 1...")

    parameters = (
        parse_forecast_parameters(
            excel_bytes
        )
    )

    print()
    print(
        parameters.to_string(
            index=False
        )
    )

    print()
    print("Loading actual CBR monthly data...")

    actual = load_actual_data()

    latest_key_rate_date = (
        actual
        .dropna(subset=["key_rate"])
        ["date"]
        .max()
    )

    latest_inflation_date = (
        actual
        .dropna(
            subset=[
                "russia_inflation_yoy"
            ]
        )
        ["date"]
        .max()
    )

    print(
        "Latest actual key-rate month:",
        latest_key_rate_date.date(),
    )

    print(
        "Latest actual inflation month:",
        latest_inflation_date.date(),
    )

    print()
    print("Building monthly trajectories...")

    monthly = build_monthly_output(
        actual=actual,
        parameters=parameters,
    )

    save_outputs(
        parameters=parameters,
        monthly=monthly,
        comment_url=comment_url,
        xlsx_url=final_xlsx_url,
    )

    print()
    print("=" * 60)
    print("DONE")
    print("=" * 60)

    print()
    print(
        "Parameters:",
        PARAMETERS_OUTPUT_FILE,
    )

    print(
        "Monthly output:",
        MONTHLY_OUTPUT_FILE,
    )

    print(
        "Source info:",
        SOURCE_INFO_FILE,
    )

    print()
    print("Last 30 monthly rows:")
    print()

    print(
        monthly.tail(30).to_string(
            index=False
        )
    )


if __name__ == "__main__":
    main()