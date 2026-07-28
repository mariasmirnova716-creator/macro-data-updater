from __future__ import annotations

import ssl
import certifi
import json
import urllib.request
from datetime import date
from pathlib import Path

import pandas as pd


API_URL = "https://api.bls.gov/publicAPI/v2/timeseries/data/"
SERIES_ID = "CUUR0000SA0"

FIRST_OUTPUT_YEAR = 2020
FIRST_SOURCE_YEAR = FIRST_OUTPUT_YEAR - 1

PROJECT_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = PROJECT_DIR / "data"

OUTPUT_FILE = DATA_DIR / "us_inflation_monthly.csv"
SOURCE_INFO_FILE = DATA_DIR / "us_inflation_source.txt"


def fetch_bls_data() -> list[dict]:
    """Получает месячные значения CPI-U из официального API BLS."""

    payload = {
        "seriesid": [SERIES_ID],
        "startyear": str(FIRST_SOURCE_YEAR),
        "endyear": str(date.today().year),
    }

    request = urllib.request.Request(
        API_URL,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "User-Agent": "macro-data-updater/1.0",
        },
        method="POST",
    )

    ssl_context = ssl.create_default_context(
        cafile=certifi.where()
    )

    with urllib.request.urlopen(
            request,
            timeout=60,
            context=ssl_context,
    ) as response:
        result = json.loads(response.read().decode("utf-8"))

    status = result.get("status")

    if status != "REQUEST_SUCCEEDED":
        messages = result.get("message", [])
        raise RuntimeError(
            f"BLS API returned status {status!r}: {messages}"
        )

    series = result.get("Results", {}).get("series", [])

    if not series:
        raise RuntimeError("BLS API returned no series")

    observations = series[0].get("data", [])

    if not observations:
        raise RuntimeError("BLS API returned no observations")

    return observations


def build_dataframe(observations: list[dict]) -> pd.DataFrame:
    """Преобразует CPI в инфляцию к тому же месяцу прошлого года."""

    records: list[dict[str, object]] = []

    for item in observations:
        period = str(item.get("period", ""))
        raw_value = item.get("value")

        if not period.startswith("M") or period == "M13":
            continue

        if raw_value in (None, "", "-"):
            continue

        try:
            year_number = int(item["year"])
            month_number = int(period[1:])
            cpi_value = float(raw_value)
        except (KeyError, TypeError, ValueError) as exc:
            raise ValueError(
                f"Invalid BLS observation: {item}"
            ) from exc

        records.append(
            {
                "year": year_number,
                "month": month_number,
                "date": (
                    pd.Timestamp(
                        year=year_number,
                        month=month_number,
                        day=1,
                    )
                    + pd.offsets.MonthEnd(0)
                ),
                "cpi_index": cpi_value,
            }
        )

    if not records:
        raise RuntimeError("No valid monthly CPI observations found")

    frame = (
        pd.DataFrame(records)
        .drop_duplicates(
            subset=["year", "month"],
            keep="last",
        )
        .sort_values(["year", "month"])
        .reset_index(drop=True)
    )

    previous_year = frame[
        ["year", "month", "cpi_index"]
    ].copy()

    previous_year["year"] = previous_year["year"] + 1

    previous_year = previous_year.rename(
        columns={"cpi_index": "cpi_previous_year"}
    )

    frame = frame.merge(
        previous_year,
        on=["year", "month"],
        how="left",
    )

    frame["us_inflation_yoy"] = (
        frame["cpi_index"]
        .div(frame["cpi_previous_year"])
        .sub(1)
        .mul(100)
    )

    frame = frame[
        frame["year"] >= FIRST_OUTPUT_YEAR
    ].copy()

    frame = frame.dropna(
        subset=["us_inflation_yoy"]
    ).reset_index(drop=True)

    frame["us_inflation_yoy"] = (
        frame["us_inflation_yoy"].round(2)
    )

    frame["date"] = frame["date"].dt.strftime(
        "%Y-%m-%d"
    )

    return frame[
        [
            "date",
            "us_inflation_yoy",
        ]
    ]


def main() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    print("Downloading BLS CPI data...")

    observations = fetch_bls_data()
    result = build_dataframe(observations)

    result.to_csv(
        OUTPUT_FILE,
        index=False,
        encoding="utf-8-sig",
    )

    SOURCE_INFO_FILE.write_text(
        f"source={API_URL}\n"
        f"series_id={SERIES_ID}\n"
        f"updated_at={pd.Timestamp.now().isoformat()}\n"
        f"rows={len(result)}\n"
        f"last_date={result.iloc[-1]['date']}\n",
        encoding="utf-8",
    )

    print()
    print("Done.")
    print(f"Series: {SERIES_ID}")
    print(f"Rows saved: {len(result)}")
    print(f"Last date: {result.iloc[-1]['date']}")
    print(f"Last value: {result.iloc[-1]['us_inflation_yoy']}")
    print(f"CSV: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
