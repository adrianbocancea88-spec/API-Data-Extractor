"""
Fetch data from a public REST API, extract useful fields, display them, and
export the result to Excel.

This example uses JSONPlaceholder's public /users endpoint:
https://jsonplaceholder.typicode.com/users
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Any, Iterable

try:
    import pandas as pd
    import requests
    from requests.exceptions import (
        ConnectionError as RequestsConnectionError,
        HTTPError,
        RequestException,
        Timeout,
    )
except ImportError as exc:  # pragma: no cover - runs only when dependencies are missing.
    missing_package = exc.name or "a required package"
    raise SystemExit(
        f"Missing dependency: {missing_package}. "
        "Install dependencies with: python -m pip install -r requirements.txt"
    ) from exc


API_URL = "https://jsonplaceholder.typicode.com/users"
DEFAULT_OUTPUT_FILE = Path(__file__).with_name("api_users.xlsx")
DEFAULT_TIMEOUT_SECONDS = 10


class ApiRequestError(RuntimeError):
    """Raised when the API request or response cannot be processed."""


def parse_arguments(argv: Iterable[str] | None = None) -> argparse.Namespace:
    """Parse command-line options for API URL, timeout, and Excel output path."""
    parser = argparse.ArgumentParser(
        description="Fetch public REST API data and export selected fields to Excel."
    )
    parser.add_argument(
        "--url",
        default=API_URL,
        help=f"REST API endpoint to call. Default: {API_URL}",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=DEFAULT_OUTPUT_FILE,
        help=f"Excel file path to create. Default: {DEFAULT_OUTPUT_FILE}",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        default=DEFAULT_TIMEOUT_SECONDS,
        help=f"Request timeout in seconds. Default: {DEFAULT_TIMEOUT_SECONDS}",
    )
    return parser.parse_args(argv)


def fetch_json(url: str, timeout_seconds: float) -> list[dict[str, Any]]:
    """
    Call the REST API and return its JSON payload.

    The function converts lower-level requests exceptions into one application
    error type so callers can handle failures consistently.
    """
    try:
        response = requests.get(url, timeout=timeout_seconds)
        response.raise_for_status()
    except RequestsConnectionError as exc:
        raise ApiRequestError(
            "Connection error: could not reach the API. Check your internet connection."
        ) from exc
    except Timeout as exc:
        raise ApiRequestError(
            f"Timeout error: the API did not respond within {timeout_seconds} seconds."
        ) from exc
    except HTTPError as exc:
        status_code = exc.response.status_code if exc.response is not None else "unknown"
        raise ApiRequestError(f"HTTP error: the API returned status {status_code}.") from exc
    except RequestException as exc:
        raise ApiRequestError(f"Request error: {exc}") from exc

    try:
        payload = response.json()
    except ValueError as exc:
        raise ApiRequestError("Invalid response: the API did not return valid JSON.") from exc

    if not isinstance(payload, list):
        raise ApiRequestError("Unexpected response format: expected a list of records.")

    return payload


def as_text(value: Any) -> str:
    """Convert API values to strings while keeping missing values readable."""
    if value is None:
        return ""
    return str(value)


def extract_user_information(records: list[dict[str, Any]]) -> list[dict[str, str]]:
    """
    Extract business-friendly user fields from the raw API records.

    JSONPlaceholder user objects include nested address and company values, so
    this function flattens those nested fields into a tabular structure.
    """
    users: list[dict[str, str]] = []

    for record in records:
        if not isinstance(record, dict):
            continue

        address = record.get("address") or {}
        company = record.get("company") or {}

        users.append(
            {
                "id": as_text(record.get("id")),
                "name": as_text(record.get("name")),
                "username": as_text(record.get("username")),
                "email": as_text(record.get("email")),
                "phone": as_text(record.get("phone")),
                "website": as_text(record.get("website")),
                "city": as_text(address.get("city") if isinstance(address, dict) else ""),
                "company": as_text(company.get("name") if isinstance(company, dict) else ""),
            }
        )

    return users


def build_dataframe(users: list[dict[str, str]]) -> pd.DataFrame:
    """Build a pandas DataFrame with a stable, user-friendly column order."""
    columns = ["id", "name", "username", "email", "phone", "website", "city", "company"]
    return pd.DataFrame(users, columns=columns)


def save_to_excel(dataframe: pd.DataFrame, output_file: Path) -> Path:
    """Save the extracted data to an Excel workbook and return its absolute path."""
    output_file.parent.mkdir(parents=True, exist_ok=True)

    # pandas uses openpyxl from requirements.txt to create .xlsx workbooks.
    dataframe.to_excel(output_file, index=False, sheet_name="Users")
    return output_file.resolve()


def display_table(dataframe: pd.DataFrame) -> None:
    """Print the extracted information as a clean table in the terminal."""
    if dataframe.empty:
        print("No records found.")
        return

    display_columns = ["id", "name", "username", "email", "city", "company"]
    table = dataframe[display_columns].rename(
        columns={
            "id": "ID",
            "name": "Name",
            "username": "Username",
            "email": "Email",
            "city": "City",
            "company": "Company",
        }
    )

    with pd.option_context(
        "display.max_columns",
        None,
        "display.max_colwidth",
        32,
        "display.width",
        140,
    ):
        print(table.to_string(index=False, justify="left"))


def run(url: str, output_file: Path, timeout_seconds: float) -> Path:
    """Run the complete workflow and return the generated Excel file path."""
    raw_records = fetch_json(url, timeout_seconds)
    users = extract_user_information(raw_records)
    dataframe = build_dataframe(users)

    print(f"\nExtracted {len(dataframe)} user records from {url}\n")
    display_table(dataframe)

    return save_to_excel(dataframe, output_file)


def main(argv: Iterable[str] | None = None) -> int:
    """Command-line entry point."""
    args = parse_arguments(argv)

    try:
        saved_file = run(args.url, args.output, args.timeout)
    except ApiRequestError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1
    except OSError as exc:
        print(f"ERROR: could not write the Excel file: {exc}", file=sys.stderr)
        return 1
    except ImportError as exc:
        print(
            "ERROR: pandas could not write the Excel file. "
            "Install openpyxl with: python -m pip install openpyxl",
            file=sys.stderr,
        )
        print(f"Details: {exc}", file=sys.stderr)
        return 1

    print(f"\nExcel file saved to: {saved_file}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
