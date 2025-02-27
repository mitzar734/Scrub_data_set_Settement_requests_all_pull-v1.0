import json
import snowflake.connector
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import datetime

# Configurations
SNOWFLAKE_CREDENTIALS_FILE = r"C:\Users\rcruz\Documents\GLG\Python\Credentials\snflk_ROBINSONC_creds.json"  
GDRIVE_CREDENTIALS_FILE = r"C:\Users\rcruz\Documents\GLG\Python\Credentials\rs-python-scripts-001-b279ccb0d2c0.json"  # Google Service Account JSON
SHEET_ID = "17C1bESM6p0aD_27QM9hoKhGtJ9gTZc-bTyYcjnSUAxc"  # Replace with actual Sheet ID
WORKSHEET_NAME = "TBL_GLS_CREATED_SETTLEMENTS_SUMMARY"  # Change as needed
#COLUMNS_TO_CLEAR = ["A", "B", "C", "D", "E", "F"]  # Specify columns to clear before inserting new data


def load_snowflake_credentials():
    """Load Snowflake credentials from a JSON or TXT file."""
    with open(SNOWFLAKE_CREDENTIALS_FILE, "r") as file:
        creds = json.load(file)
    return creds

def datetime_to_excel(date_obj):
    """Converts a datetime.date object to an Excel serial date number."""
    if isinstance(date_obj, (datetime.date, pd.Timestamp)):
        excel_start_date = datetime.date(1899, 12, 30)
        return (date_obj - excel_start_date).days
    return date_obj  # Return unchanged if not a date

def fetch_snowflake_data():
    """Connects to Snowflake and fetches data from the specified table."""
    creds = load_snowflake_credentials()
    conn = snowflake.connector.connect(
        user=creds["user"],
        password=creds["password"],
        account=creds["account"],
        warehouse=creds["warehouse"],
        database=creds["database"],
        schema=creds["schema"]
    )
    
    query = "SELECT * FROM GLG_ICON_NEGOTIATIONS.EXPORT.TBL_GLS_CREATED_SETTLEMENTS_SUMMARY WHERE CREATED_BY_NAME IS NOT NULL ORDER BY CAL_DATE"  # Replace with actual table
    df = pd.read_sql(query, conn)
    conn.close()
    
    return df


def authenticate_gsheets():
    """Authenticates with Google Sheets using a Service Account."""
    creds = Credentials.from_service_account_file(GDRIVE_CREDENTIALS_FILE, scopes=["https://www.googleapis.com/auth/spreadsheets"])
    client = gspread.authorize(creds)
    return client


def clear_specific_columns(sheet):
    """Clears only specific columns in the Google Sheet."""
    #for col in COLUMNS_TO_CLEAR:
    #    sheet.batch_clear([f"{col}:{col}"])
    sheet.batch_clear([f"{'A'}:{'F'}"])


def push_data_to_gsheet(df):
    """Pushes the fetched Snowflake data into the Google Sheet."""
    client = authenticate_gsheets()
    sheet = client.open_by_key(SHEET_ID).worksheet(WORKSHEET_NAME)

    # Clear specific columns
    clear_specific_columns(sheet)

    # Convert DataFrame to list and convert only datetime.date values to Excel serial numbers
    values = [df.columns.tolist()] + df.applymap(
        lambda x:   x.strftime('%Y-%m-%d %H:%M:%S') if isinstance(x, pd.Timestamp) else
                    datetime_to_excel(x) if isinstance(x, (datetime.date)) else x
    ).values.tolist()

    # Write data to the Google Sheet
    sheet.update("A1", values)  # Writes data starting from A1

    # Apply formatting for date columns (A and B)
    sheet.format("A:B", {"numberFormat": {"type": "DATE", "pattern": "mm/dd/yyyy"}})

    # Apply numeric formatting if necessary (e.g., for column C with numeric data)
    sheet.format("C:C", {"numberFormat": {"type": "NUMBER", "pattern": "##"}})


def main():
    """Main execution function."""
    df = fetch_snowflake_data()
    push_data_to_gsheet(df)
    print("Data successfully updated in Google Sheets!")


if __name__ == "__main__":
    main()