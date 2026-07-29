from __future__ import annotations

import json
import ssl
import urllib.parse
import urllib.request
from datetime import date
from pathlib import Path

import certifi
import pandas as pd


API_URL = (
    "https://ec.europa.eu/eurostat/api/dissemination/"
    "statistics/1.0/data/prc_hicp_manr"
)

FIRST_YEAR = 2020

PROJECT_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = PROJECT_DIR / "data"

OUTPUT_FILE = DATA_DIR / "eurozone_inflation_monthly.csv"
SOURCE_INFO_FILE = DATA_DIR / "eurozone_inflation_source.txt"


def create_ssl_context() -> ssl.SSLContext:
    return ssl.create_default_context(cafile=certifi.where())


def build_url() -> str:
    today = date.today()

    params = {
        "lang": "en",
        "freq": "M",
        "unit": "RCH_A",
        "coicop": "CP00",
        "geo": "EA20",
        "sinceTimePeriod": f"{FIRST_YEAR}-01",
        "untilTimePeriod": f"{today.year}-{today.month:02d}",
    }

    return f"{API_URL}?{urllib.parse.urlencode(params)}"


def fetch_data() -> dict:
    url = build_url()

    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": "macro-data-updater/1.0",
            "Accept": "application/json",
        },
    )

    with urllib.request.urlopen(
        request,
        timeout=90,
        context=create_ssl_context(),
    ) as response:
        payload = json.loads(response.read().decode("utf-8"))

    return payload


def build_dataframe(payload: dict) -> pd.DataFrame:
    try:
        time_index = (
            payload["dimension"]["time"]["category"]["index"]
        )
        values = payload["value"]
    except KeyError as exc:
        raise RuntimeError(
            "Eurostat response has an unexpected structure"
        ) from exc

    records: list[dict[str, object]] = []

    for period, position in time_index.items():
        raw_value = values.get(str(position))

        if raw_value is None:
            continue

        try:
            year_number, month_number = map(
                int,
                period.split("-"),
            )
        except ValueError as exc:
            raise ValueError(
                f"Invalid Eurostat period: {period}"
            ) from exc

        observation_date = (
            pd.Timestamp(
                year=year_number,
                month=month_number,
                day=1,
            )
            + pd.offsets.MonthEnd(0)
        )

        records.append(
            {
                "date": observation_date,
                "eurozone_inflation_yoy": float(raw_value),
            }
        )

    if not records:
        raise RuntimeError(
            "Eurostat returned no valid inflation observations"
        )

    result = (
        pd.DataFrame(records)
        .drop_duplicates(subset=["date"], keep="last")
        .sort_values("date")
        .reset_index(drop=True)
    )

    result = result[
        result["date"].dt.year >= FIRST_YEAR
    ].copy()

    result["eurozone_inflation_yoy"] = (
        result["eurozone_inflation_yoy"].round(2)
    )

    result["date"] = result["date"].dt.strftime(
        "%Y-%m-%d"
    )

    return result[
        [
            "date",
            "eurozone_inflation_yoy",
        ]
    ]


def main() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    print("Downloading Eurostat HICP data...")

    payload = fetch_data()
    result = build_dataframe(payload)

    result.to_csv(
        OUTPUT_FILE,
        index=False,
        encoding="utf-8-sig",
    )

    SOURCE_INFO_FILE.write_text(
        f"source={build_url()}\n"
        f"dataset=prc_hicp_manr\n"
        f"geo=EA20\n"
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
        "Last value: "
        f"{result.iloc[-1]['eurozone_inflation_yoy']}"
    )
    print(f"CSV: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()