import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="a4")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_num = filename.split("-")[0]
    pdf.set_font("Times", "B", 24)
    pdf.cell(w=0, h=16, txt=f"Invoice #: {invoice_num}", ln=1, )
    pdf.output(f"PDFs/{filename}.pdf")