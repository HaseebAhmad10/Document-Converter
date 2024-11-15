import tkinter as tk
from tkinter import filedialog, messagebox
from converter import (
    docx_to_pdf, pdf_to_text, markdown_to_html, txt_to_pdf,
    pdf_to_docx, image_to_pdf, audio_to_text, pptx_to_pdf, 
    compress_pdf, compress_text_file  
)

#  conversion 
conversion_map = {
    'DOCX to PDF': docx_to_pdf,
    'PDF to Text': pdf_to_text,
    'Markdown to HTML': markdown_to_html,
    'TXT to PDF': txt_to_pdf,
    'PDF to DOCX': pdf_to_docx,
    'Image to PDF': image_to_pdf,
    'Audio to Text': audio_to_text,
    'PPTX to PDF': pptx_to_pdf,
    'Compress PDF': compress_pdf,       
    'Compress Text File': compress_text_file  
}

def convert_file():
    conversion_type = conversion_var.get()
    convert_function = conversion_map.get(conversion_type)

    if not file_path.get():
        messagebox.showerror("Error", "Please select a file to convert.")
        return
    
    if not convert_function:
        messagebox.showerror("Error", "Please select a conversion type.")
        return

    input_file = file_path.get()
    output_extension = (
        ".pdf" if "to PDF" in conversion_type or "Compress PDF" == conversion_type else
        ".txt" if "to Text" in conversion_type or "Compress Text File" == conversion_type else
        ".html" if "to HTML" in conversion_type else
        ".docx"
    )
    
    output_file = filedialog.asksaveasfilename(
        defaultextension=output_extension, 
        filetypes=[(f"{output_extension.upper()} files", f"*{output_extension}"), ("All files", "*.*")]
    )
    
    if not output_file:
        return

    try:
        if conversion_type == "Image to PDF":
            image_files = filedialog.askopenfilenames(filetypes=[("Image files", "*.jpg *.jpeg *.png")])
            if not image_files:
                return
            convert_function(image_files, output_file)
        
        elif conversion_type == "Compress PDF":
            compress_pdf(input_file, output_file, method="pymupdf", quality=20)
        
        elif conversion_type == "Compress Text File":
            compress_text_file(input_file, output_file, method="gzip")
        
        else:
            convert_function(input_file, output_file)
        
        messagebox.showinfo("Success", f"File converted and saved as {output_file}")
    
    except Exception as e:
        messagebox.showerror("Conversion Error", str(e))

def browse_file():
    filetypes = [
        ("All files", "*.*"),
        ("DOCX files", "*.docx"),
        ("PDF files", "*.pdf"),
        ("Markdown files", "*.md"),
        ("TXT files", "*.txt"),
        ("Image files", "*.jpg *.jpeg *.png"),
        ("Audio files", "*.mp3 *.wav"),
        ("PPTX files", "*.pptx")
    ]
    filename = filedialog.askopenfilename(filetypes=filetypes)
    file_path.set(filename)

root = tk.Tk()
root.title("Document Converter")
root.geometry("500x300")

frame = tk.Frame(root)
frame.pack(pady=10)

file_path = tk.StringVar()

browse_button = tk.Button(frame, text="Browse File", command=browse_file)
browse_button.grid(row=0, column=0, padx=5)

file_entry = tk.Entry(frame, textvariable=file_path, width=40)
file_entry.grid(row=0, column=1, padx=5)

conversion_var = tk.StringVar()
conversion_var.set("Select Conversion Type")

conversion_options = list(conversion_map.keys())
conversion_menu = tk.OptionMenu(root, conversion_var, *conversion_options)
conversion_menu.pack(pady=10)

convert_button = tk.Button(root, text="Convert", command=convert_file)
convert_button.pack(pady=20)

root.mainloop()
