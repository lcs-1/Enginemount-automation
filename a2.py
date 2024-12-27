import pandas as pd
from docxtpl import DocxTemplate
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
import pdfplumber

# Function to handle saving the document
def save_document():
    co_name = co_name_entry.get()
    part = part_entry.get()
    d_no = d_no_entry.get()
    depth = depth_entry.get()

    if not co_name or not part or not d_no or not depth:
        messagebox.showerror("Error", "All fields must be filled.")
        return

    context = {'co_name': co_name, 'part': part, 'd_no': d_no, 'depth': depth}

    doc = DocxTemplate(r"C:\L.C.S\auto\Template.docx")
    doc.render(context)
    
    output_filename = f"Impact_{co_name}.docx"
    output_path = fr"C:\L.C.S\auto\{output_filename}"
    doc.save(output_path)
    
    messagebox.showinfo("Success", "Document saved successfully.")

# Function to open the PDF viewer
def open_pdf_viewer():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if pdf_path:
        with pdfplumber.open(pdf_path) as pdf:
            first_page = pdf.pages[0]
            extracted_text = first_page.extract_text()
            text_dialog = tk.Toplevel(root)
            text_dialog.title("Selected Text from PDF")
            text_widget = tk.Text(text_dialog)
            text_widget.insert(tk.END, extracted_text)
            text_widget.pack()

# Create the main UI window using tkinter
root = tk.Tk()
root.title("Document Generator")

# Create labels and entry fields for each variable
co_name_label = tk.Label(root, text="CO Name:")
co_name_label.pack()
co_name_entry = tk.Entry(root)
co_name_entry.pack()

part_label = tk.Label(root, text="Part:")
part_label.pack()
part_entry = tk.Entry(root)
part_entry.pack()

d_no_label = tk.Label(root, text="D No:")
d_no_label.pack()
d_no_entry = tk.Entry(root)
d_no_entry.pack()

depth_label = tk.Label(root, text="Depth:")
depth_label.pack()
depth_entry = tk.Entry(root)
depth_entry.pack()

# Create buttons for PDF viewer and document generation
pdf_viewer_button = tk.Button(root, text="Open PDF Viewer", command=open_pdf_viewer)
pdf_viewer_button.pack()

save_button = tk.Button(root, text="Save Document", command=save_document)
save_button.pack()

# Start the GUI event loop
root.mainloop()
