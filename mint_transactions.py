import mintapi
from datetime import datetime, timedelta
import pandas as pd
from mintapi.filters import DateFilter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl import load_workbook
import os

from dotenv import load_dotenv

load_dotenv()

email = os.getenv("MINT_EMAIL")
password = os.getenv("MINT_PASSWORD")

# print(email, password)

mint = mintapi.Mint(email=email, password=password)

end_date = datetime.today().date()
start_date = end_date - timedelta(days=7)

start_date_str = start_date.strftime('%Y-%m-%d')
# start_date_str = "" # ! Uncomment this line and set custom start date
end_date_str = end_date.strftime('%Y-%m-%d')

print('start_date: ' + start_date_str)
print('end_date: ' + end_date_str)

checking_account_id = os.getenv("CHECKING_ACCOUNT_ID")

savings_account_id = os.getenv("SAVINGS_ACCOUNT_ID")

# print(checking_account_id, savings_account_id)

# ! Custom date filter
transactions = mint.get_transaction_data(DateFilter.Options(11), account_ids=[checking_account_id], start_date=start_date_str, end_date=end_date_str)

# ! Last 7 days
# transactions = mint.get_transaction_data(DateFilter.Options(1), account_ids=[checking_account_id])

transactions.reverse() # API returns mostgetenv first, we want them last

# # replace with the path to your custom Excel sheet
excel_path = os.getenv("EXCEL_PATH")

print(excel_path)

transactions = pd.DataFrame(transactions)

# =============================================================================



# ! Create a new Excel file or overrite an existing ones
def create_transactions_excel_file(transactions, excel_path):
    transactions.insert(0, 'row_id', range(1, len(transactions)+1))
    writer = pd.ExcelWriter(excel_path)
    transactions.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()
    


# =============================================================================


def update_transactions_excel_file(transactions, excel_path):
    # ! Update existing Excel file

    # Add a column with unique IDs for each row
    existing_df = pd.read_excel(excel_path, sheet_name='Sheet1', engine='openpyxl')
    last_id = existing_df['row_id'].max()
    if pd.isna(last_id):
        last_id = 0
    transactions.insert(0, 'row_id', range(last_id+1, last_id+len(transactions)+1))


    # Read the existing sheet into a DataFrame
    if excel_path.endswith('.xls'):
        existing_df = pd.read_excel(excel_path, sheet_name='Sheet1', engine='xlrd')
    else:
        existing_df = pd.read_excel(excel_path, sheet_name='Sheet1', engine='openpyxl')

    # Append the existing data to the new data
    updated_df = pd.concat([existing_df, transactions])
    # updated_df = existing_df.append(transactions)

    # Save the updated data to the Excel file
    updated_df.to_excel(excel_path, index=False)
    
# =============================================================================

# * Uncomment the function you want to use
# create_transactions_excel_file(transactions, excel_path)
update_transactions_excel_file(transactions, excel_path)
