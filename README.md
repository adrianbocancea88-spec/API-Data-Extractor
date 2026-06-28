# REST API to Excel Python App

This application fetches user data from the public JSONPlaceholder REST API,
extracts useful fields, prints them in a formatted terminal table, and saves the
result to an Excel workbook with pandas.

## Setup

Use Python 3.10 or newer.

```powershell
python -m pip install -r requirements.txt
```

## Run

```powershell
python rest_api_excel_app.py
```

By default, the script creates:

```text
api_users.xlsx
```

## Optional Arguments

```powershell
python rest_api_excel_app.py --output users.xlsx
python rest_api_excel_app.py --timeout 5
python rest_api_excel_app.py --url https://jsonplaceholder.typicode.com/users
```

## What It Demonstrates

- Calling a public REST API with `requests`.
- Handling connection, timeout, HTTP, and JSON parsing errors.
- Extracting useful nested JSON fields into flat records.
- Saving tabular data into an Excel file with `pandas`.
- Displaying extracted data in a formatted terminal table.
- Organizing logic into reusable functions.
