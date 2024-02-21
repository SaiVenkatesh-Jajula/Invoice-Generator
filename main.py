import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("Invoices/*xlsx")

for filepath in filepaths:
    pdf = FPDF(orientation='p', unit='mm', format='a4')
    pdf.add_page()
    # Design your invoice first how it looks!
    # Date & Invoice Number from filename
    filename=Path(filepath).stem
    invoiceno, date=filename.split('-')

    pdf.set_font(family='Times', style='B', size=12)
    pdf.cell(w=0, h=12, txt=f"Invoice No: {invoiceno}", align='R', ln=1)
    pdf.cell(w=0, h=12, txt=f"Date : {date}", align='R', ln=1)

    #Empty space
    pdf.ln(50)

    #Company Logo
    pdf.image("img.png", w=40,x=10,y=4)

    #Main Header
    pdf.set_font(family="Times",style='B',size=24)
    pdf.cell(w=0,h=20, txt="Invoice",ln=1,border=1,align="C")

    # Reading Excel
    excelfile = Path(filepath)
    df = pd.read_excel(excelfile,sheet_name='Sheet 1')

    # Code for Table Headers
    columns = df.columns
    tableheaders = [i.replace('_'," ").title() for i in columns]
    pdf.set_font(family='Times',style='B',size=10)
    pdf.cell(w=30, h=12, txt=tableheaders[0],align='L',border=1,ln=0)
    pdf.cell(w=70, h=12, txt=tableheaders[1],align='L',border=1,ln=0)
    pdf.cell(w=40, h=12, txt=tableheaders[2], align='L', border=1, ln=0)
    pdf.cell(w=30, h=12, txt=tableheaders[3], align='L', border=1, ln=0)
    pdf.cell(w=20, h=12, txt=tableheaders[4], align='L', border=1, ln=1)

    # Table items
    # total_price = 0
    for index,item in df.iterrows():
        pdf.set_font(family='Times', size=9)
        pdf.cell(w=30, h=12, txt=str(item['product_id']), align='L', border=1, ln=0)
        pdf.cell(w=70, h=12, txt=str(item['product_name']), align='L', border=1, ln=0)
        pdf.cell(w=40, h=12, txt=str(item['amount_purchased']), align='L', border=1, ln=0)
        pdf.cell(w=30, h=12, txt=str(item['price_per_unit']), align='L', border=1, ln=0)
        pdf.cell(w=20, h=12, txt=str(item['total_price']), align='L', border=1, ln=1)
        # total_price+=int(item['total_price'])

    # Total Price
    total_price=df["total_price"].sum()
    pdf.set_font(family='Times', size=9)
    pdf.cell(w=30, h=12, txt="", align='L', border=1, ln=0)
    pdf.cell(w=70, h=12, txt="", align='L', border=1, ln=0)
    pdf.cell(w=40, h=12, txt="", align='L', border=1, ln=0)
    pdf.cell(w=30, h=12, txt="", align='L', border=1, ln=0)
    pdf.cell(w=20,h=12,txt=str(total_price),align='L',border=1,ln=1)

    pdf.ln(5)

    pdf.set_font(family='Times',style='B',size=13)
    pdf.cell(w=20, h=12, txt=f"The total due : {total_price}$ ", align='L', ln=1)



    pdf.output(f"GInvoices/{filename}.pdf")
