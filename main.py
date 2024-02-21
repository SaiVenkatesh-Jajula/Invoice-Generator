import pandas as pd
import glob

filepaths=glob.glob("Invoices/*xlsx")
# print(filepaths)

for filepath in filepaths:
    df=pd.read_excel(filepath)
    # print(df)
