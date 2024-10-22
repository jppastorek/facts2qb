import pandas as pd
import os
from dotenv import load_dotenv
import gspread
from gspread_dataframe import set_with_dataframe
# make sure you install openpyxl, pandas, python-dotenv, gspread, gspread-dataframe

output_file = 'qb_import.csv'

# load the env file that has sensitive data like file and path, maybe this is overkill
load_dotenv()

# get the facts excel file from the env and the google credentials
FACTS_EXCEL_PATH = os.environ.get('FACTS_EXCEL_PATH')
GOOGLE_CREDENTIALS = os.environ.get('GOOGLE_CREDENTIALS')

# convert the excel file to a data frame for manipulation
df = pd.read_excel(FACTS_EXCEL_PATH)

# select the columns you need
selected_columns = df[['Payments', 'Name', 'Reference ID', 'Account', 'Payment Method']]

# rename those columns for the import file
selected_columns.rename(columns={
    'Payments': 'Amount',
    'Name': 'Received From',
    'Reference ID': 'Ref NO.'
}, inplace=True)

# add the class column and fill with Canal Street Campus
selected_columns['Class'] = 'Canal Street Campus'

###############################################################
# SHOULD THE PAYMENT METHOD BE CHANGED TO SOMETHING SPECIFIC? #
###############################################################

# convert all the names to all caps
selected_columns['Received From'] = selected_columns['Received From'].str.upper()

# write to a new file
# selected_columns.to_csv(output_file, index=False)

# connect and authenticate to google
gc = gspread.service_account(filename=GOOGLE_CREDENTIALS)
spreadsheet = gc.open('QB_Depo_Map')
worksheet = spreadsheet.worksheet('Sheet1')

set_with_dataframe(worksheet, selected_columns)