from dotenv import load_dotenv
import pandas as pd
import os

load_dotenv()

 
vendors = pd.read_excel('vendors.xlsx',sheet_name=os.getenv('SHEET_NAME'), dtype={'Employee Bank Account Number': str})

for index,row in vendors.iterrows():
    print(row["Bank Name"])