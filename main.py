import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("Invoices/*xlsx")
# print(filepaths)

for filepath in filepaths:
    df = pd.read_excel(filepath)
    # print(df)
    pdf = FPDF(orientation='p', unit='mm', format='a4')
    pdf.add_page()
    # Design your invoice first how it looks!
    filename=Path(filepath).stem
    invoiceno, date=filename.split('-')
    pdf.set_font(family='Times', style='B', size=12)
    pdf.cell(w=0, h=12, txt=f"Invoice No: {invoiceno}", align='L', ln=1)
    pdf.cell(w=0, h=12, txt=f"Date : {date}", align='L', ln=1)

    pdf.output(f"GInvoices/{filename}.pdf")
