from fpdf import FPDF
import pandas as pd
import glob
import pathlib


filepaths = glob.glob("invoices/*.xlsx")

for filepath in filepaths:
    file_name = pathlib.Path(filepath).stem
    invoce_number, date = file_name.split('-')

    # Create a PDF object
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    # Invoice Number
    pdf.set_font(family="Arial", style="B", size=24)
    pdf.cell(w=0,h=12,txt=f'Invoice nr. {invoce_number}',border=0,ln=1,align='',fill=0,link='')

    # Date
    pdf.set_font(family="Arial", style="B", size=24)
    pdf.cell(w=0,h=12,txt=f'Date: {date}',border=0,ln=1,align='',fill=0,link='')

    pdf.ln()
    # Getting the data from Excel and create a table in pdf
    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # setup
    cell_font_size = 8

    #  Creating the header
    for col in [i.replace('_', ' ').title() for i in df.columns]:
        pdf.set_font(family="Arial", style='B', size=cell_font_size)
        if col == "Product Name":
            pdf.cell(w=70, h=12, txt=f'{col}', border=1, ln=0, align='', fill=0, link='')
        else:
            pdf.cell(w=30, h=12, txt=f'{col}', border=1, ln=0, align='', fill=0, link='')
    pdf.ln()
    for _, row in df.iterrows():
        for col in df.columns:
            pdf.set_font(family="Arial", size=cell_font_size)
            if col == "product_name":
                pdf.cell(w=70, h=12, txt=f'{row[col]}', border=1, ln=0, align='', fill=0, link='')
            else:
                pdf.cell(w=30, h=12, txt=f'{row[col]}', border=1, ln=0, align='', fill=0, link='')
        pdf.ln()

    # Creating the Total Price row
    total_price = df[col].sum()
    for col in df.columns:
        pdf.set_font(family="Arial", size=cell_font_size)
        if col == "product_name":
            pdf.cell(w=70, h=12, txt="", border=1, ln=0, align='', fill=0, link='')
        elif col == "total_price":
            pdf.cell(w=30, h=12, txt=f"{total_price}", border=1, ln=0, align='', fill=0, link='')
        else:
            pdf.cell(w=30, h=12, txt="", border=1, ln=0, align='', fill=0, link='')
    pdf.ln()

    # Adding a line
    pdf.set_font(family="Arial", size=cell_font_size + (cell_font_size / 2))
    pdf.cell(w=0, h=12, txt=f"The Total Price is {total_price}", border=0, ln=1, align='', fill=0, link='')

    # Creating the PDF file
    pdf.output(f"PDFs/{file_name}.pdf")




