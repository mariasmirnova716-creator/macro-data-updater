from __future__ import annotations

import json
import ssl
import urllib.parse
import urllib.request
from datetime import date
from pathlib import Path
from typing import Any

import certifi
import pandas as pd


COUNTRY_CODE = "RU"
FIRST_YEAR = 2011

INDICATORS = {
    "gdp_per_capita_constant_2015_usd": "NY.GDP.PCAP.KD",
    "gdp_per_capita_current_usd": "NY.GDP.PCAP.CD",
}

API_BASE_URL = "https://api.worldbank.org/v2"

PROJECT_DIR = Path(__file__).resolve().parents[2]
DATA_DIR = PROJECT_DIR / "data" / "sheet_01_gdp_data"

OUTPUT_FILE = DATA_DIR / "world_bank_gdp_per_capita.csv"
SOURCE_INFO_FILE = DATA_DIR / "world_bank_gdp_per_capita_source.txt"


def create_ssl_context() -> ssl.SSLContext:
    return ssl.create_default_context(cafile=certifi.where())


def download_json(url: str) -> Any:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
                "AppleWebKit/537.36 Chrome/126 Safari/537.36"
            ),
            "Accept": "application/json",
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
        # Резервный вариант для проблем сертификатов на Mac.
        unsafe_context = ssl.create_default_context()
        unsafe_context.check_hostname = False
        unsafe_context.verify_mode = ssl.CERT_NONE

        with urllib.request.urlopen(
            request,
            timeout=120,
            context=unsafe_context,
        ) as response:
            content = response.read()

    try:
        return json.loads(content.decode("utf-8"))

    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise RuntimeError(
            "World Bank returned a response that is not valid JSON."
        ) from exc


def build_api_url(indicator_code: str) -> str:
    current_year = date.today().year

    query = urllib.parse.urlencode(
        {
            "format": "json",
            "per_page": 1000,
            "date": f"{FIRST_YEAR}:{current_year}",
        }
    )

    return (
        f"{API_BASE_URL}/country/{COUNTRY_CODE}"
        f"/indicator/{indicator_code}?{query}"
    )


def fetch_indicator(
    output_column: str,
    indicator_code: str,
) -> tuple[pd.DataFrame, str]:

    url = build_api_url(indicator_code)

    print()
    print(f"Downloading {indicator_code}...")
    print(url)

    response = download_json(url)

    if not isinstance(response, list) or len(response) < 2:
        raise RuntimeError(
            f"{indicator_code}: unexpected World Bank API response."
        )

    metadata = response[0]
    observations = response[1]

    if not isinstance(metadata, dict):
        raise RuntimeError(
            f"{indicator_code}: API metadata is missing."
        )

    if not isinstance(observations, list):
        raise RuntimeError(
            f"{indicator_code}: API observations are missing."
        )

    records: list[dict[str, Any]] = []

    for observation in observations:
        if not isinstance(observation, dict):
            continue

        observation_year = observation.get("date")
        value = observation.get("value")

        try:
            observation_year_int = int(observation_year)
        except (TypeError, ValueError):
            continue

        if observation_year_int < FIRST_YEAR:
            continue

        if value is None:
            numeric_value = None
        else:
            try:
                numeric_value = float(value)
            except (TypeError, ValueError) as exc:
                raise RuntimeError(
                    f"{indicator_code}: non-numeric value "
                    f"for {observation_year_int}: {value!r}"
                ) from exc

        records.append(
            {
                "year": observation_year_int,
                output_column: numeric_value,
            }
        )

    if not records:
        raise RuntimeError(
            f"{indicator_code}: no observations were returned."
        )

    frame = (
        pd.DataFrame(records)
        .drop_duplicates(subset=["year"], keep="last")
        .sort_values("year")
        .reset_index(drop=True)
    )

    non_null_count = int(frame[output_column].notna().sum())

    if non_null_count < 10:
        raise RuntimeError(
            f"{indicator_code}: only {non_null_count} "
            "non-empty observations were returned."
        )

    print(
        f"Downloaded: rows={len(frame)}, "
        f"non-empty={non_null_count}, "
        f"latest year with value="
        f"{int(frame.loc[frame[output_column].notna(), 'year'].max())}"
    )

    return frame, url


def combine_indicators() -> tuple[pd.DataFrame, list[str]]:
    result: pd.DataFrame | None = None
    source_urls: list[str] = []

    for output_column, indicator_code in INDICATORS.items():
        indicator_frame, source_url = fetch_indicator(
            output_column=output_column,
            indicator_code=indicator_code,
        )

        source_urls.append(source_url)

        if result is None:
            result = indicator_frame
        else:
            result = result.merge(
                indicator_frame,
                on="year",
                how="outer",
                validate="one_to_one",
            )

    if result is None:
        raise RuntimeError(
            "No World Bank indicators were downloaded."
        )

    result = (
        result
        .sort_values("year")
        .reset_index(drop=True)
    )

    result["date"] = pd.to_datetime(
        result["year"].astype(str) + "-12-31"
    ).dt.strftime("%Y-%m-%d")

    result = result[
        [
            "date",
            "year",
            "gdp_per_capita_constant_2015_usd",
            "gdp_per_capita_current_usd",
        ]
    ]

    for column in (
        "gdp_per_capita_constant_2015_usd",
        "gdp_per_capita_current_usd",
    ):
        result[column] = (
            pd.to_numeric(result[column], errors="coerce")
            .round(1)
        )

    # Удаляем годы, в которых нет ни одного показателя.
    result = result.dropna(
        subset=[
            "gdp_per_capita_constant_2015_usd",
            "gdp_per_capita_current_usd",
        ],
        how="all",
    ).reset_index(drop=True)

    return result, source_urls


def validate_result(result: pd.DataFrame) -> None:
    required_columns = {
        "date",
        "year",
        "gdp_per_capita_constant_2015_usd",
        "gdp_per_capita_current_usd",
    }

    missing_columns = required_columns.difference(result.columns)

    if missing_columns:
        raise RuntimeError(
            f"Missing output columns: {sorted(missing_columns)}"
        )

    if result.empty:
        raise RuntimeError(
            "World Bank GDP per capita result is empty."
        )

    if result["year"].duplicated().any():
        duplicate_years = (
            result.loc[result["year"].duplicated(keep=False), "year"]
            .tolist()
        )

        raise RuntimeError(
            f"Duplicate years found: {duplicate_years}"
        )

    if not result["year"].is_monotonic_increasing:
        raise RuntimeError(
            "Years are not sorted in ascending order."
        )

    if int(result["year"].min()) > FIRST_YEAR:
        raise RuntimeError(
            f"First returned year is {int(result['year'].min())}, "
            f"but expected data from {FIRST_YEAR}."
        )

    for column in (
        "gdp_per_capita_constant_2015_usd",
        "gdp_per_capita_current_usd",
    ):
        values = pd.to_numeric(
            result[column],
            errors="coerce",
        ).dropna()

        if values.empty:
            raise RuntimeError(
                f"{column}: no numeric observations found."
            )

        if not values.between(100, 100_000).all():
            bad_values = values[
                ~values.between(100, 100_000)
            ].tolist()

            raise RuntimeError(
                f"{column}: implausible values found: "
                f"{bad_values[:10]}"
            )

        latest_year = int(
            result.loc[result[column].notna(), "year"].max()
        )

        # Всемирный банк нередко публикует годовые данные с лагом.
        if latest_year < date.today().year - 3:
            raise RuntimeError(
                f"{column}: data unexpectedly stops at "
                f"{latest_year}."
            )


def main() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)

    result, source_urls = combine_indicators()
    validate_result(result)

    result.to_csv(
        OUTPUT_FILE,
        index=False,
        encoding="utf-8-sig",
    )

    latest_constant_year = int(
        result.loc[
            result["gdp_per_capita_constant_2015_usd"].notna(),
            "year",
        ].max()
    )

    latest_current_year = int(
        result.loc[
            result["gdp_per_capita_current_usd"].notna(),
            "year",
        ].max()
    )

    SOURCE_INFO_FILE.write_text(
        "\n".join(
            [
                f"constant_indicator={INDICATORS['gdp_per_capita_constant_2015_usd']}",
                f"current_indicator={INDICATORS['gdp_per_capita_current_usd']}",
                f"constant_url={source_urls[0]}",
                f"current_url={source_urls[1]}",
                f"updated_at={pd.Timestamp.now().isoformat()}",
                f"rows={len(result)}",
                f"latest_constant_year={latest_constant_year}",
                f"latest_current_year={latest_current_year}",
                "",
            ]
        ),
        encoding="utf-8",
    )

    print()
    print("Done.")
    print(f"Rows saved: {len(result)}")
    print(
        "Latest constant-price year: "
        f"{latest_constant_year}"
    )
    print(
        "Latest current-price year: "
        f"{latest_current_year}"
    )
    print(f"CSV: {OUTPUT_FILE}")

    print()
    print("Last 10 observations:")
    print(result.tail(10).to_string(index=False))


if __name__ == "__main__":
    main()