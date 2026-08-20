from __future__ import annotations

import io
import re
import warnings
from datetime import date
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

DATA_DIR = PROJECT_ROOT / "data" / "sheet_04_rates_oil_data"
DATA_DIR.mkdir(parents=True, exist_ok=True)

OUTPUT_FILE = DATA_DIR / "sheet_04_rates_oil.csv"
INVESTING_API_URL = (
    "https://api.investing.com/api/financialdata/historical/8833"
)

# ============================================================
# SETTINGS
# ============================================================

START_DATE = "01.01.2020"

CBR_URL = "https://www.cbr.ru/currency_base/dynamics/"

USD_CODE = "R01235"
CNY_CODE = "R01375"

INVESTING_URL = (
    "https://www.investing.com/commodities/brent-oil-historical-data"
)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/139.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

TIMEOUT = 60


# ============================================================
# HTTP
# ============================================================

def get_session() -> requests.Session:
    session = requests.Session()
    session.headers.update(HEADERS)
    return session


def request_get(
    session: requests.Session,
    url: str,
    *,
    params: dict | None = None,
) -> requests.Response:

    try:
        response = session.get(
            url,
            params=params,
            timeout=TIMEOUT,
            verify=certifi.where(),
        )
        response.raise_for_status()
        return response

    except requests.exceptions.SSLError:
        warnings.simplefilter("ignore", InsecureRequestWarning)

        response = session.get(
            url,
            params=params,
            timeout=TIMEOUT,
            verify=False,
        )
        response.raise_for_status()
        return response


# ============================================================
# HELPERS
# ============================================================

def first_day_of_month(value: pd.Timestamp) -> pd.Timestamp:
    return pd.Timestamp(
        year=value.year,
        month=value.month,
        day=1,
    )


def parse_number(value: object) -> float:
    text = str(value).strip()

    text = text.replace("\xa0", "")
    text = text.replace(" ", "")
    text = text.replace(",", ".")

    return float(text)


# ============================================================
# CBR
# ============================================================

def download_cbr_currency(
    session: requests.Session,
    currency_code: str,
    column_name: str,
) -> pd.DataFrame:

    today = date.today().strftime("%d.%m.%Y")

    params = {
        "UniDbQuery.Posted": "True",
        "UniDbQuery.mode": "1",
        "UniDbQuery.date_req1": "",
        "UniDbQuery.date_req2": "",
        "UniDbQuery.VAL_NM_RQ": currency_code,
        "UniDbQuery.From": START_DATE,
        "UniDbQuery.To": today,
    }

    response = request_get(
        session,
        CBR_URL,
        params=params,
    )

    response.encoding = "utf-8"

    tables = pd.read_html(
        io.StringIO(response.text),
        decimal=",",
        thousands=" ",
    )

    target = None

    for table in tables:

        if table.shape[1] < 3:
            continue

        test = table.astype(str)

        # Look for a table containing real dates.
        date_hits = (
            test.iloc[:, 0]
            .str.match(r"\d{2}\.\d{2}\.\d{4}", na=False)
            .sum()
        )

        if date_hits >= 2:
            target = table.copy()
            break

    if target is None:
        print(f"Found {len(tables)} HTML tables on CBR page")

        for i, table in enumerate(tables):
            print()
            print(f"TABLE {i}:")
            print(table.head())

        raise RuntimeError(
            f"CBR currency data table not found for {currency_code}"
        )

    target = target.iloc[:, :3].copy()

    target.columns = [
        "date",
        "nominal",
        "rate",
    ]

    target["date"] = pd.to_datetime(
        target["date"],
        dayfirst=True,
        errors="coerce",
    )

    target["nominal"] = (
        target["nominal"]
        .astype(str)
        .str.replace("\xa0", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False)
    )

    target["rate"] = (
        target["rate"]
        .astype(str)
        .str.replace("\xa0", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", ".", regex=False)
    )

    target["nominal"] = pd.to_numeric(
        target["nominal"],
        errors="coerce",
    )

    target["rate"] = pd.to_numeric(
        target["rate"],
        errors="coerce",
    )

    target = target.dropna(
        subset=[
            "date",
            "nominal",
            "rate",
        ]
    )

    # Convert to the units used in our Excel sheet.
    if currency_code == USD_CODE:

        # RUB per 1 USD
        target[column_name] = (
            target["rate"]
            / target["nominal"]
        )

    elif currency_code == CNY_CODE:

        # RUB per 10 CNY
        target[column_name] = (
            target["rate"]
            / target["nominal"]
            * 10
        )

    else:

        target[column_name] = (
            target["rate"]
            / target["nominal"]
        )

    # One observation per month.
    # Use the first official CBR rate available in that month.
    target["month"] = (
        target["date"]
        .dt.to_period("M")
    )

    target = (
        target
        .sort_values("date")
        .groupby(
            "month",
            as_index=False,
        )
        .first()
    )

    target["date"] = (
        target["month"]
        .dt.to_timestamp()
    )

    return target[
        [
            "date",
            column_name,
        ]
    ]


# ============================================================
# BRENT / INVESTING
# ============================================================

def download_brent(
    session: requests.Session,
) -> pd.DataFrame:

    today = date.today().strftime("%Y-%m-%d")

    params = {
        "start-date": "2020-01-01",
        "end-date": today,
        "time-frame": "Monthly",
        "add-missing-rows": "false",
    }

    headers = {
        **HEADERS,
        "domain-id": "www",
        "Accept": "application/json, text/plain, */*",
        "Referer": INVESTING_URL,
    }

    response = session.get(
        INVESTING_API_URL,
        params=params,
        headers=headers,
        timeout=TIMEOUT,
        verify=certifi.where(),
    )

    response.raise_for_status()

    data = response.json()

    if "data" not in data:
        print("Investing response:")
        print(data)
        raise RuntimeError(
            "Investing historical data field not found"
        )

    rows = data["data"]

    print("Latest Investing rows:")
    for row in rows[:3]:
        print(
            row.get("rowDate"),
            row.get("last_close"),
        )

    if not rows:
        raise RuntimeError(
            "Investing returned no Brent data"
        )

    df = pd.DataFrame(rows)

    print("Investing columns:")
    print(df.columns.tolist())

    # Expected Investing fields normally include:
    # direction_color, rowDate, rowDateRaw, last_close,
    # last_open, last_max, last_min, volume, change_precent

    if "rowDate" not in df.columns:
        raise RuntimeError(
            f"Investing date column not found. Columns: {df.columns.tolist()}"
        )

    if "last_close" not in df.columns:
        raise RuntimeError(
            f"Investing price column not found. Columns: {df.columns.tolist()}"
        )

    df["date"] = pd.to_datetime(
        df["rowDate"],
        errors="coerce",
    )

    df["brent_usd"] = (
        df["last_close"]
        .astype(str)
        .str.replace(",", "", regex=False)
        .str.strip()
    )

    df["brent_usd"] = pd.to_numeric(
        df["brent_usd"],
        errors="coerce",
    )

    df = df.dropna(
        subset=["date", "brent_usd"]
    )

    # Normalize every monthly observation to first day of month
    df["date"] = (
        df["date"]
        .dt.to_period("M")
        .dt.to_timestamp()
    )

    df = (
        df[
            [
                "date",
                "brent_usd",
            ]
        ]
        .sort_values("date")
        .drop_duplicates(
            subset=["date"],
            keep="last",
        )
        .reset_index(drop=True)
    )

    if df.empty:
        raise RuntimeError(
            "Brent dataframe is empty after parsing"
        )

    return df

# ============================================================
# VALIDATION
# ============================================================

def validate_output(df: pd.DataFrame) -> None:

    if df.empty:
        raise RuntimeError(
            "Output dataset is empty"
        )

    required = {
        "date",
        "brent_usd",
        "usd_rub",
        "cny_10_rub",
    }

    missing = required - set(df.columns)

    if missing:
        raise RuntimeError(
            f"Missing columns: {sorted(missing)}"
        )

    if df["date"].duplicated().any():
        raise RuntimeError(
            "Duplicate dates detected"
        )

    # Basic sanity limits.
    valid_brent = df["brent_usd"].dropna()

    if not valid_brent.empty:
        if not valid_brent.between(10, 250).all():
            raise RuntimeError(
                "Brent values outside expected range"
            )

    valid_usd = df["usd_rub"].dropna()

    if not valid_usd.empty:
        if not valid_usd.between(20, 250).all():
            raise RuntimeError(
                "USD/RUB values outside expected range"
            )

    valid_cny = df["cny_10_rub"].dropna()

    if not valid_cny.empty:
        if not valid_cny.between(20, 400).all():
            raise RuntimeError(
                "CNY/10 values outside expected range"
            )


# ============================================================
# MAIN
# ============================================================

def main() -> None:

    session = get_session()

    print("Downloading USD/RUB from CBR...")

    usd = download_cbr_currency(
        session=session,
        currency_code=USD_CODE,
        column_name="usd_rub",
    )

    print(
        f"USD/RUB: {len(usd)} months, "
        f"{usd['date'].min().date()} -> "
        f"{usd['date'].max().date()}"
    )

    print("Downloading CNY/RUB from CBR...")

    cny = download_cbr_currency(
        session=session,
        currency_code=CNY_CODE,
        column_name="cny_10_rub",
    )

    print(
        f"CNY/10: {len(cny)} months, "
        f"{cny['date'].min().date()} -> "
        f"{cny['date'].max().date()}"
    )

    print("Downloading Brent from Investing.com...")

    brent = download_brent(
        session=session,
    )

    print(
        f"Brent: {len(brent)} rows, "
        f"{brent['date'].min().date()} -> "
        f"{brent['date'].max().date()}"
    )

    result = (
        usd
        .merge(
            cny,
            on="date",
            how="outer",
        )
        .merge(
            brent,
            on="date",
            how="outer",
        )
    )

    result = (
        result
        .sort_values("date")
        .reset_index(drop=True)
    )

    start = pd.Timestamp("2020-01-01")

    result = result[
        result["date"] >= start
    ].copy()

    validate_output(result)

    result["date"] = (
        result["date"]
        .dt.strftime("%Y-%m-%d")
    )

    result.to_csv(
        OUTPUT_FILE,
        index=False,
        encoding="utf-8-sig",
    )

    print()
    print("Saved:")
    print(OUTPUT_FILE)
    print()
    print(result.tail(12).to_string(index=False))


if __name__ == "__main__":
    main()