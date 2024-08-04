import subprocess
import sys

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

try:
    import pandas
except:
    install("pandas")

try:
    import persiantools
except:
    install("persiantools")

try:
    import openpyxl
except:
    install("openpyxl")

import pandas as pd
from persiantools.jdatetime import JalaliDate

def get_persian_date():
    now = JalaliDate.today()
    persian_date = now.strftime("%y%m%d")
    return persian_date

def create_group_deposit_file(deposits):
    if not deposits:
        print("No data found in the Excel file.")
        return

    # Check if 'AMOUNT' key exists in each dictionary
    if any("AMOUNT" not in deposit for deposit in deposits):
        print("Invalid data format in the Excel file. 'AMOUNT' column is missing.")
        return
    # Calculate the sums for the first line
    num_deposits = str(len(deposits)).zfill(10)
    total_amount = str(int(sum(float(deposit.get("AMOUNT", 0)) for deposit in deposits))).zfill(15)

    # Prepare the first line according to specifications
    first_line = num_deposits + total_amount

    # Prepare the content for the deposit file
    deposit_data = []
    for deposit in deposits:
        account_number = str(deposit.get("ACCOUNT_NUMBER", "")).zfill(10)
        amount = str(int(float(deposit.get("AMOUNT", 0)))).zfill(15)
        transaction_number = str(deposit.get("TRANSACTION_NUMBER", "")).zfill(17)
        note = str(deposit.get("NOTE", "")).rjust(30, " ")  # Fill with spaces before text
        depositor_name = str(deposit.get("DEPOSITOR_NAME", "")).rjust(30, " ")  # Fill with spaces before text

        deposit_data.append(account_number + amount + transaction_number + note + depositor_name)

    # Combine all data into a single string
    deposit_file_content = first_line + "\n" + "\n".join(deposit_data) + "\n"  # Add an empty line at the end

    # Get the Persian date
    persian_date = get_persian_date()

    # Create the file name with the .PAY extension in capital letters
    output_file_name = "FL" + persian_date + ".PAY"

    # Write the content to the file with Windows 1256 encoding and the specified file name
    with open(output_file_name, "w", encoding="windows-1256", errors="replace") as file:
        file.write(deposit_file_content)

if __name__ == "__main__":
    # Read data from Excel file
    excel_file_path = "./payment.xlsx"
    try:
        df = pd.read_excel(excel_file_path)
    except FileNotFoundError:
        print("Excel file not found. Please provide the correct path.")
        exit(1)
    except Exception as e:
        print("Error reading Excel file:", e)
        exit(1)

    # Fill missing values (NaN) with empty strings
    df.astype(str).fillna("", inplace=True)

    # Replace Persian character '\u06cc' (ی) with its Unicode equivalent '\u064a' (ی) because Windows 1296 does not support '\u06cc'
    df.replace('\u06cc', '\u064a', regex=True, inplace=True)

    # Convert DataFrame to a list of dictionaries
    deposits_list = df.to_dict(orient="records")

    create_group_deposit_file(deposits_list)
