from fpdf import FPDF

def format_text(text, font="Arial", size=12):
    pdf = FPDF()
    pdf.set_font(font, size=size)
    pdf.add_page()
    pdf.multi_cell(200, 10, txt=text, align="L")
    return pdf
