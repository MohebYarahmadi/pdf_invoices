import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    # Generate a PDF
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Extract Invoice Information
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    # Add Invoice Information
    pdf.set_font(family="Arial", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)
    pdf.cell(w=50, h=8, txt=f'Date {date}', ln=1)
    pdf.ln(5)

    df = pd.read_excel(filepath, sheet_name='Sheet 1')

    # Add Table Header
    columns = [item.replace("_", " ").title() for item in df.columns]
    pdf.set_font(family="Arial", size=8, style="B")
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=70, h=8, txt=columns[1], border=1)
    pdf.cell(w=30, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    # Add Table Rows
    for index, row in df.iterrows():
        pdf.set_font(family="Arial", size=10)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=70, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1, ln=1)
        

    pdf.output(f"PDFs/{filename}.pdf")
