# Snowflake to Google Sheets Automation

This Python script extracts data from a Snowflake database and updates a Google Sheet automatically. It ensures that only non-null data is pulled, orders it by date, and clears specific columns before inserting new data.

## Features
- Connects to Snowflake and retrieves data from the `TBL_GLS_CREATED_SETTLEMENTS_SUMMARY` table.
- Authenticates with Google Sheets using a Service Account.
- Clears specific columns (`A` to `F`) before updating data.
- Converts date fields to an Excel-compatible format.
- Formats columns in Google Sheets (date formatting for columns `A:B`, numeric formatting for column `C`).

## Requirements
### Python Packages
Ensure you have the following dependencies installed:
```bash
pip install snowflake-connector-python pandas gspread google-auth
```

### Configuration Files
- **Snowflake Credentials File**: JSON file containing Snowflake login details (`snflk_ROBINSONC_creds.json`).
- **Google Drive Credentials File**: JSON file for Google Sheets authentication (`rs-python-scripts-001-b279ccb0d2c0.json`).

## Setup
1. Update the paths in the script to point to the correct credentials files:
   ```python
   SNOWFLAKE_CREDENTIALS_FILE = "path/to/snowflake_credentials.json"
   GDRIVE_CREDENTIALS_FILE = "path/to/gdrive_credentials.json"
   ```
2. Replace `SHEET_ID` with the actual Google Sheets ID.
3. Modify `WORKSHEET_NAME` if necessary.

## Usage
Run the script using:
```bash
python script.py
```

This will fetch data from Snowflake and update the specified Google Sheet.

## Customization
- Modify the `query` in `fetch_snowflake_data()` to extract different data.
- Adjust column formatting in `push_data_to_gsheet()` as needed.

## License
This project is licensed under the MIT License. Feel free to modify and use it as needed.

## Author
[Robinson Cruz]

