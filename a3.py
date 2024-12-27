import pandas as pd
from docxtpl import DocxTemplate
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
import pdfplumber

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

def on_resize(event):
    label_width = max(len("CO Name:"), len("Part:"), len("D No:"), len("Depth:"))
    entry_width = 20
    padding_x = 10
    padding_y = 5

    co_name_label.config(width=label_width)
    part_label.config(width=label_width)
    d_no_label.config(width=label_width)
    depth_label.config(width=label_width)

    co_name_entry.config(width=entry_width)
    part_entry.config(width=entry_width)
    d_no_entry.config(width=entry_width)
    depth_entry.config(width=entry_width)

    pdf_viewer_button.config(width=label_width + entry_width + padding_x)
    save_button.config(width=label_width + entry_width + padding_x)

root = tk.Tk()
root.title("Document Generator")

# Create labels and entry fields for each variable
co_name_label = tk.Label(root, text="CO Name:")
co_name_entry = tk.Entry(root)

part_label = tk.Label(root, text="Part:")
part_entry = tk.Entry(root)

d_no_label = tk.Label(root, text="D No:")
d_no_entry = tk.Entry(root)

depth_label = tk.Label(root, text="Depth:")
depth_entry = tk.Entry(root)

# Create buttons for PDF viewer and document generation
pdf_viewer_button = tk.Button(root, text="Open PDF Viewer", command=open_pdf_viewer)
save_button = tk.Button(root, text="Save Document", command=save_document)

# Use the grid geometry manager to arrange widgets
co_name_label.grid(row=0, column=0, padx=10, pady=5, sticky='e')
co_name_entry.grid(row=0, column=1, padx=10, pady=5)

part_label.grid(row=1, column=0, padx=10, pady=5, sticky='e')
part_entry.grid(row=1, column=1, padx=10, pady=5)

d_no_label.grid(row=2, column=0, padx=10, pady=5, sticky='e')
d_no_entry.grid(row=2, column=1, padx=10, pady=5)

depth_label.grid(row=3, column=0, padx=10, pady=5, sticky='e')
depth_entry.grid(row=3, column=1, padx=10, pady=5)

pdf_viewer_button.grid(row=4, column=0, columnspan=2, padx=10, pady=10)
save_button.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

# Configure grid column weights to make the entry fields expand with window resize
root.grid_columnconfigure(1, weight=1)

# Bind the window resize event to the on_resize function
root.bind("<Configure>", on_resize)

root.mainloop()
