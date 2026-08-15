from __future__ import annotations

from datetime import date
from pathlib import Path

import pandas as pd
import requests


# ============================================================
# PATHS
# ============================================================

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parents[1]

DATA_DIR = (
    PROJECT_ROOT
    / "data"
    / "sheet_2_ipp_data"
)

OUTPUT_FILE = (
    DATA_DIR
    / "russia_business_activity_expectations.csv"
)

SOURCE_INFO_FILE = (
    DATA_DIR
    / "russia_business_activity_expectations_source.txt"
)


# ============================================================
# CBR API
# ============================================================

API_URL = "https://cbr.ru/dataservice/data"

PUBLICATION_ID = 25
DATASET_ID = 58
MEASURE_ID = 119

START_YEAR = 2002

CURRENT_ACTIVITY_ID = 44
EXPECTED_ACTIVITY_ID = 45

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 "
        "(KHTML, like Gecko) "
        "Chrome/148 Safari/537.36"
    ),
    "Accept": "application/json",
}


# ============================================================
# DOWNLOAD
# ============================================================

def download_cbr_data() -> dict:

    end_year = date.today().year

    params = {
        "y1": START_YEAR,
        "y2": end_year,
        "publicationId": PUBLICATION_ID,
        "datasetId": DATASET_ID,
        "measureId": MEASURE_ID,
    }

    print("=" * 72)
    print("CBR business activity / expectations updater")
    print("=" * 72)

    print("\nRequesting CBR API:")
    print(API_URL)
    print(params)

    response = requests.get(
        API_URL,
        params=params,
        headers=HEADERS,
        timeout=90,
    )

    response.raise_for_status()

    data = response.json()

    if "RawData" not in data:
        raise RuntimeError(
            "CBR API response does not contain RawData."
        )

    if "headerData" not in data:
        raise RuntimeError(
            "CBR API response does not contain headerData."
        )

    return data


# ============================================================
# VALIDATE HEADER IDS
# ============================================================

def validate_header_ids(
    data: dict,
) -> None:

    headers = data["headerData"]

    header_map = {}

    for item in headers:

        element_id = item.get("id")

        name = (
            item.get("elname")
            or item.get("name")
            or ""
        )

        if element_id is not None:
            header_map[int(element_id)] = str(name)

    if CURRENT_ACTIVITY_ID not in header_map:
        raise RuntimeError(
            f"Current activity element_id "
            f"{CURRENT_ACTIVITY_ID} "
            f"not found in headerData."
        )

    if EXPECTED_ACTIVITY_ID not in header_map:
        raise RuntimeError(
            f"Expected activity element_id "
            f"{EXPECTED_ACTIVITY_ID} "
            f"not found in headerData."
        )

    print("\nSelected CBR indicators:")

    print(
        CURRENT_ACTIVITY_ID,
        "->",
        header_map[CURRENT_ACTIVITY_ID],
    )

    print(
        EXPECTED_ACTIVITY_ID,
        "->",
        header_map[EXPECTED_ACTIVITY_ID],
    )


# ============================================================
# PARSE
# ============================================================

def parse_business_activity(
    data: dict,
) -> pd.DataFrame:

    raw = data["RawData"]

    rows = []

    for item in raw:

        element_id = item.get(
            "element_id"
        )

        if element_id not in {
            CURRENT_ACTIVITY_ID,
            EXPECTED_ACTIVITY_ID,
        }:
            continue

        raw_date = item.get(
            "date"
        )

        value = item.get(
            "obs_val"
        )

        if raw_date is None:
            continue

        parsed_date = pd.to_datetime(
            raw_date,
            errors="coerce",
        )

        if pd.isna(parsed_date):
            continue

        numeric_value = pd.to_numeric(
            value,
            errors="coerce",
        )

        rows.append(
            {
                "date": parsed_date,
                "element_id": int(
                    element_id
                ),
                "value": numeric_value,
            }
        )

    df = pd.DataFrame(
        rows
    )

    if df.empty:
        raise RuntimeError(
            "No CBR business activity observations extracted."
        )

    # --------------------------------------------------------
    # One row per date, two selected indicators as columns
    # --------------------------------------------------------

    wide = (
        df
        .pivot_table(
            index="date",
            columns="element_id",
            values="value",
            aggfunc="last",
        )
        .reset_index()
    )

    wide = wide.rename(
        columns={
            CURRENT_ACTIVITY_ID:
                "current_activity",
            EXPECTED_ACTIVITY_ID:
                "expected_activity_3m",
        }
    )

    if "current_activity" not in wide.columns:
        raise RuntimeError(
            "Current activity series was not found "
            "after pivot."
        )

    if "expected_activity_3m" not in wide.columns:
        raise RuntimeError(
            "Expected activity series was not found "
            "after pivot."
        )

    wide[
        "expectations_gap"
    ] = (
        wide["expected_activity_3m"]
        - wide["current_activity"]
    )

    wide[
        "source"
    ] = "cbr_business_monitoring"

    # Use date only.
    wide[
        "date"
    ] = pd.to_datetime(
        wide["date"]
    ).dt.date

    numeric_columns = [
        "current_activity",
        "expected_activity_3m",
        "expectations_gap",
    ]

    for column in numeric_columns:

        wide[column] = (
            pd.to_numeric(
                wide[column],
                errors="coerce",
            )
            .round(2)
        )

    wide = (
        wide
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

    return wide


# ============================================================
# PRESERVE HISTORY
# ============================================================

def merge_existing_history(
    current: pd.DataFrame,
) -> pd.DataFrame:

    if not OUTPUT_FILE.exists():
        return current

    print(
        "\nLoading existing business-monitoring history..."
    )

    existing = pd.read_csv(
        OUTPUT_FILE
    )

    required = {
        "date",
        "current_activity",
        "expected_activity_3m",
        "expectations_gap",
    }

    if not required.issubset(
        existing.columns
    ):
        raise RuntimeError(
            "Existing business activity CSV "
            "has unexpected columns."
        )

    existing[
        "date"
    ] = pd.to_datetime(
        existing["date"],
        errors="coerce",
    ).dt.date

    if "source" not in existing.columns:
        existing[
            "source"
        ] = "existing_history"

    current_dates = set(
        current["date"]
    )

    # Preserve observations that disappeared from
    # the latest CBR response.
    history_only = existing[
        ~existing["date"].isin(
            current_dates
        )
    ].copy()

    history_only[
        "source"
    ] = "existing_history"

    result = pd.concat(
        [
            history_only,
            current,
        ],
        ignore_index=True,
        sort=False,
    )

    return (
        result
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


# ============================================================
# VALIDATION
# ============================================================

def validate(
    df: pd.DataFrame,
) -> None:

    if df.empty:
        raise RuntimeError(
            "Business activity output is empty."
        )

    if df["date"].duplicated().any():
        raise RuntimeError(
            "Duplicate dates in business activity output."
        )

    for column in [
        "current_activity",
        "expected_activity_3m",
    ]:

        bad = df[
            df[column].notna()
            &
            ~df[column].between(
                -100,
                100,
            )
        ]

        if not bad.empty:

            raise RuntimeError(
                f"Values outside expected range "
                f"in {column}:\n"
                + bad.to_string(
                    index=False
                )
            )

    latest_date = pd.to_datetime(
        df["date"]
    ).max()

    if latest_date.year < date.today().year - 1:
        raise RuntimeError(
            "CBR business-monitoring data "
            f"look unexpectedly old: {latest_date.date()}"
        )


# ============================================================
# MAIN
# ============================================================

def main():

    DATA_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    data = download_cbr_data()

    validate_header_ids(
        data
    )

    current = parse_business_activity(
        data
    )

    print(
        "\nCurrent CBR data:"
    )

    print(
        current
        .tail(15)
        .to_string(
            index=False
        )
    )

    result = merge_existing_history(
        current
    )

    validate(
        result
    )

    result.to_csv(
        OUTPUT_FILE,
        index=False,
        encoding="utf-8",
    )

    SOURCE_INFO_FILE.write_text(
        "\n".join(
            [
                "Russia business activity and expectations",
                "source=Bank of Russia",
                f"api={API_URL}",
                f"publication_id={PUBLICATION_ID}",
                f"dataset_id={DATASET_ID}",
                f"measure_id={MEASURE_ID}",
                f"current_activity_element_id={CURRENT_ACTIVITY_ID}",
                f"expected_activity_element_id={EXPECTED_ACTIVITY_ID}",
                "",
                (
                    "current_activity="
                    "How did production volume, contracted work, "
                    "turnover and services change?"
                ),
                (
                    "expected_activity_3m="
                    "How will production volume, contracted work, "
                    "turnover and services change in the next "
                    "three months?"
                ),
                (
                    "expectations_gap="
                    "expected_activity_3m-current_activity"
                ),
                (
                    "unit="
                    "balance of responses, points"
                ),
                (
                    "adjustment="
                    "seasonally adjusted CBR data"
                ),
                (
                    f"latest_date="
                    f"{result['date'].max()}"
                ),
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