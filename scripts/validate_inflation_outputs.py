from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd


PROJECT_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = PROJECT_DIR / "data"


@dataclass(frozen=True)
class DatasetRule:
    filename: str
    required_columns: tuple[str, ...]
    numeric_columns: tuple[str, ...]
    minimum_rows: int
    minimum_date: str
    maximum_age_days: int
    value_min: float
    value_max: float


RULES = (
    DatasetRule(
        filename="cbr_inflation_monthly.csv",
        required_columns=(
            "date",
            "key_rate",
            "russia_inflation_yoy",
            "inflation_target",
        ),
        numeric_columns=(
            "key_rate",
            "russia_inflation_yoy",
            "inflation_target",
        ),
        minimum_rows=70,
        minimum_date="2020-01-31",
        maximum_age_days=90,
        value_min=-20,
        value_max=100,
    ),
    DatasetRule(
        filename="us_inflation_monthly.csv",
        required_columns=(
            "date",
            "us_inflation_yoy",
        ),
        numeric_columns=(
            "us_inflation_yoy",
        ),
        minimum_rows=65,
        minimum_date="2020-01-31",
        maximum_age_days=90,
        value_min=-20,
        value_max=100,
    ),
    DatasetRule(
        filename="eurozone_inflation_monthly.csv",
        required_columns=(
            "date",
            "eurozone_inflation_yoy",
        ),
        numeric_columns=(
            "eurozone_inflation_yoy",
        ),
        minimum_rows=60,
        minimum_date="2020-01-31",
        # Пока Еврозона отстаёт. Это будет предупреждение,
        # а не причина остановки workflow.
        maximum_age_days=240,
        value_min=-20,
        value_max=100,
    ),
    DatasetRule(
        filename="russia_cpi_monthly.csv",
        required_columns=(
            "date",
            "russia_cpi_mom_index",
        ),
        numeric_columns=(
            "russia_cpi_mom_index",
        ),
        minimum_rows=70,
        minimum_date="2020-01-"
                     ""
                     "31",
        maximum_age_days=90,
        value_min=50,
        value_max=150,
    ),
    DatasetRule(
        filename="russia_cpi_quarterly.csv",
        required_columns=(
            "date",
            "year",
            "quarter",
            "cpi_end_to_previous_quarter_end",
            "cpi_quarter_to_previous_quarter",
            "cpi_quarter_to_previous_year",
        ),
        numeric_columns=(
            "year",
            "quarter",
            "cpi_end_to_previous_quarter_end",
            "cpi_quarter_to_previous_quarter",
            "cpi_quarter_to_previous_year",
        ),
        minimum_rows=25,
        minimum_date="2019-03-31",
        maximum_age_days=200,
        value_min=50,
        value_max=150,
    ),
)


def validate_dataset(rule: DatasetRule) -> list[str]:
    path = DATA_DIR / rule.filename
    warnings: list[str] = []

    if not path.exists():
        raise RuntimeError(f"Required file does not exist: {path}")

    frame = pd.read_csv(path, encoding="utf-8-sig")

    if frame.empty:
        raise RuntimeError(f"{rule.filename}: file is empty")

    missing_columns = [
        column
        for column in rule.required_columns
        if column not in frame.columns
    ]

    if missing_columns:
        raise RuntimeError(
            f"{rule.filename}: missing columns: {missing_columns}. "
            f"Actual columns: {list(frame.columns)}"
        )

    if len(frame) < rule.minimum_rows:
        raise RuntimeError(
            f"{rule.filename}: only {len(frame)} rows; "
            f"expected at least {rule.minimum_rows}"
        )

    dates = pd.to_datetime(
        frame["date"],
        errors="coerce",
    )

    if dates.isna().any():
        bad_rows = frame.loc[dates.isna(), "date"].tolist()

        raise RuntimeError(
            f"{rule.filename}: invalid dates found: {bad_rows[:10]}"
        )

    if dates.duplicated().any():
        duplicates = (
            dates[dates.duplicated(keep=False)]
            .dt.strftime("%Y-%m-%d")
            .tolist()
        )

        raise RuntimeError(
            f"{rule.filename}: duplicate dates found: {duplicates[:10]}"
        )

    if not dates.is_monotonic_increasing:
        raise RuntimeError(
            f"{rule.filename}: dates are not sorted in ascending order"
        )

    minimum_allowed_date = pd.Timestamp(rule.minimum_date)

    if dates.min() > minimum_allowed_date:
        raise RuntimeError(
            f"{rule.filename}: first date is {dates.min().date()}, "
            f"but data should begin no later than "
            f"{minimum_allowed_date.date()}"
        )

    today = pd.Timestamp.today().normalize()
    latest_date = dates.max()

    if latest_date > today + pd.Timedelta(days=31):
        raise RuntimeError(
            f"{rule.filename}: latest date {latest_date.date()} "
            f"is unexpectedly far in the future"
        )

    age_days = (today - latest_date).days

    if age_days > rule.maximum_age_days:
        warnings.append(
            f"{rule.filename}: latest observation is "
            f"{latest_date.date()} ({age_days} days old)"
        )

    for column in rule.numeric_columns:
        numeric_values = pd.to_numeric(
            frame[column],
            errors="coerce",
        )

        invalid_mask = (
            frame[column].notna()
            & numeric_values.isna()
        )

        if invalid_mask.any():
            bad_values = (
                frame.loc[invalid_mask, column]
                .astype(str)
                .tolist()
            )

            raise RuntimeError(
                f"{rule.filename}: non-numeric values in "
                f"{column}: {bad_values[:10]}"
            )

        non_null_values = numeric_values.dropna()

        if non_null_values.empty:
            raise RuntimeError(
                f"{rule.filename}: column {column} has no values"
            )

        # Year and quarter use separate logical checks.
        if column == "year":
            if not non_null_values.between(1990, 2100).all():
                raise RuntimeError(
                    f"{rule.filename}: invalid years found"
                )
            continue

        if column == "quarter":
            if not non_null_values.isin([1, 2, 3, 4]).all():
                raise RuntimeError(
                    f"{rule.filename}: quarter must be 1, 2, 3 or 4"
                )
            continue

        if not non_null_values.between(
            rule.value_min,
            rule.value_max,
        ).all():
            bad_values = non_null_values[
                ~non_null_values.between(
                    rule.value_min,
                    rule.value_max,
                )
            ].tolist()

            raise RuntimeError(
                f"{rule.filename}: implausible values in "
                f"{column}: {bad_values[:10]}"
            )

    print(
        f"OK: {rule.filename} | "
        f"rows={len(frame)} | "
        f"period={dates.min().date()} to {latest_date.date()}"
    )

    return warnings


def main() -> None:
    print("Validating inflation datasets...")
    print()

    all_warnings: list[str] = []

    for rule in RULES:
        warnings = validate_dataset(rule)
        all_warnings.extend(warnings)

    print()

    if all_warnings:
        print("Warnings:")

        for warning in all_warnings:
            print(f"- {warning}")

        print()

    print("All inflation datasets passed validation.")


if __name__ == "__main__":
    main()