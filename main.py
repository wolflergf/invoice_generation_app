import glob
import pandas as pd
from fpdf import FPDF
from pathlib import Path

# Get all Excel files in the "invoices" directory
filepaths = glob.glob("./invoices/*.xlsx")

# Iterate over each file
for filepath in filepaths:
    
    # Create a new PDF document
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    # Extract invoice number and date from filename
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    # Add invoice number to PDF
    pdf.set_font("Times", size=16, style='B')
    pdf.cell(w=50, h=8, txt="Invoice nr: {}".format(invoice_nr), ln=1)

    # Add date to PDF
    pdf.set_font("Times", size=16, style='B')
    pdf.cell(w=50, h=8, txt="Date: {}".format(date), ln=1)

    # Read Excel file into pandas DataFrame
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Add a header to the table in the PDF
    columns = df.columns
    columns = [item.replace("_", " ").title() for item in columns]
    pdf.set_font("Times", size=10, style='B')
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=columns[0], border=1)
    pdf.cell(w=65, h=8, txt=columns[1], border=1)
    pdf.cell(w=35, h=8, txt=columns[2], border=1)
    pdf.cell(w=30, h=8, txt=columns[3], border=1)
    pdf.cell(w=30, h=8, txt=columns[4], border=1, ln=1)

    
    
    # Add rows to the table in the PDF
    for index, row in df.iterrows():
        pdf.set_font("Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=65, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)


    # Calculate total price
    total_price = df["total_price"].sum()
    # Add total price to PDF
    pdf.set_font("Times", size=12, style='B')
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=str(" "), border=1)
    pdf.cell(w=65, h=8, txt=str(" "), border=1)
    pdf.cell(w=35, h=8, txt=str(" "), border=1)
    pdf.cell(w=30, h=8, txt=str(" "), border=1)
    pdf.cell(w=30, h=8, txt=str(total_price), border=1, ln=1)

    # Add total price to PDF
    pdf.set_font("Times", size=12, style='B')
    pdf.cell(w=30, h=8, txt="The total price is: {}".format(total_price), ln=1)

    # Add company information to PDF
    pdf.set_font("Times", size=12, style='B')
    pdf.cell(w=30, h=8, txt="WolflerPython")
    pdf.image("logo.png", w=10)

    # Save the PDF to the "PDFs" directory
    pdf.output("PDFs/{}.pdf".format(filename))