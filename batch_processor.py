def apply_font_style(pdf, font="Arial", size=12):
    pdf.set_font(font, size=size)
    print(f"Applied font style: {font} with size {size}")

def set_page_layout(pdf, layout="P"):
    if layout == "P":
        pdf.add_page()
    elif layout == "L":
        pdf.add_page(orientation="L")
    print(f"Page layout set to {layout}")

def apply_custom_formatting(pdf, font="Arial", size=12, layout="P"):
    set_page_layout(pdf, layout)
    apply_font_style(pdf, font, size)
