import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)

for filepath in filepaths:
    pdf = FPDF(orientation="P", unit="mm", format="a4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_num, date = filename.split("-")

    pdf.set_font("Times", "B", 24)
    pdf.cell(w=0, h=10, txt=f"Invoice #: {invoice_num}", ln=1)

    pdf.set_font("Times", "B", 24)
    pdf.cell(w=0, h=10, txt=f"Date: {date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    print(df)
    pdf.ln(10)
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    # Add header
    pdf.set_font("Times", "B", 10)
    pdf.set_text_color(30, 30, 30)
    pdf.cell(w=30, h=8, txt=columns[0], border=1, align="C")
    pdf.cell(w=70, h=8, txt=columns[1], border=1, align="C")
    pdf.cell(w=30, h=8, txt=columns[2], border=1, align="C")
    pdf.cell(w=30, h=8, txt=columns[3], border=1, align="C")
    pdf.cell(w=30, h=8, txt=columns[4], border=1, align="C", ln=1)

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font("Times", size=12)
        pdf.set_text_color(30, 30, 30)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1, align="C")
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1, align="C")
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1, align="C")
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1, align="C")
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, align="C", ln=1)

    # Add total price
    total_sum = df["total_price"].sum()
    pdf.set_font("Times", size=12)
    pdf.set_text_color(30, 30, 30)
    pdf.cell(w=30, h=8, txt="", border=1, align="C")
    pdf.cell(w=70, h=8, txt="", border=1, align="C")
    pdf.cell(w=30, h=8, txt="", border=1, align="C")
    pdf.cell(w=30, h=8, txt="", border=1, align="C")
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, align="C", ln=1)

    pdf.ln(10)

    # Add total sum sentence
    pdf.set_font("Times", "B", 17)
    pdf.cell(w=35, h=10, txt=f"The total price is: {total_sum}", ln=1)

    # Add company name and logo
    pdf.set_font("Times", "B", 30)
    pdf.cell(w=54, h=10, txt="PythonHow")
    pdf.image("pythonhow.png", w=11)

    pdf.output(f"PDFs/{filename}.pdf")