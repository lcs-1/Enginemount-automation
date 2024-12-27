import pandas as pd
from docxtpl import DocxTemplate
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
import pdfplumber

def save_document():
    co_name = co_name_entry.get()
    part = part_entry.get()
    p_no = p_no_entry.get()
    p_sn = p_sn_entry.get()
    engine = engine_entry.get()
    date = date_entry.get()
    ac_type = ac_type_entry.get()
    structure = structure_entry.get()

    if not co_name or not part or not p_no or not p_sn or not engine or not date or not ac_type or not structure:
        messagebox.showerror("Error", "All fields must be filled.")
        return

    context = {'co_name': co_name, 'part': part, 'p_no': p_no, 'p_sn': p_sn, 'engine':engine,'date':date,'ac_type':ac_type,'structure':structure}

    doc = DocxTemplate(r"C:\L.C.S\auto\Template-COS.docx")
    doc.render(context)
    
    output_filename = f"COS-{co_name}_Design.docx"
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
    label_width = max(len("CO Name:"), len("Part:"), len("P No:"), len("P SN:"), len("Engine:"), len("Date:"), len("AC Type:"), len("Structure:"))
    entry_width = 20
    padding_x = 10
    padding_y = 5

    co_name_label.config(width=label_width)
    part_label.config(width=label_width)
    p_no_label.config(width=label_width)
    p_sn_label.config(width=label_width)
    engine_label.config(width=label_width)
    date_label.config(width=label_width)
    ac_type_label.config(width=label_width)
    structure_label.config(width=label_width)

    co_name_entry.config(width=entry_width)
    part_entry.config(width=entry_width)
    p_no_entry.config(width=entry_width)
    p_sn_entry.config(width=entry_width)
    engine_entry.config(width=entry_width)
    date_entry.config(width=entry_width)
    ac_type_entry.config(width=entry_width)
    structure_entry.config(width=entry_width)

    pdf_viewer_button.config(width=label_width + entry_width + padding_x)
    save_button.config(width=label_width + entry_width + padding_x)

root = tk.Tk()
root.title("Document Generator")

# Create labels and entry fields for each variable
co_name_label = tk.Label(root, text="CO Name:")
co_name_entry = tk.Entry(root)

part_label = tk.Label(root, text="Part:")
part_entry = tk.Entry(root)

p_no_label = tk.Label(root, text="P No:")
p_no_entry = tk.Entry(root)

p_sn_label = tk.Label(root, text="P SN:")
p_sn_entry = tk.Entry(root)

engine_label = tk.Label(root, text="Engine:")
engine_entry = tk.Entry(root)

date_label = tk.Label(root, text="Date:")
date_entry = tk.Entry(root)

ac_type_label = tk.Label(root, text="AC Type:")
ac_type_entry = tk.Entry(root)

structure_label = tk.Label(root, text="Structure:")
structure_entry = tk.Entry(root)

# Create buttons for PDF viewer and document generation
pdf_viewer_button = tk.Button(root, text="Open PDF Viewer", command=open_pdf_viewer)
save_button = tk.Button(root, text="Save Document", command=save_document)

# Use the grid geometry manager to arrange widgets
co_name_label.grid(row=0, column=0, padx=10, pady=5, sticky='e')
co_name_entry.grid(row=0, column=1, padx=10, pady=5)

part_label.grid(row=1, column=0, padx=10, pady=5, sticky='e')
part_entry.grid(row=1, column=1, padx=10, pady=5)

p_no_label.grid(row=2, column=0, padx=10, pady=5, sticky='e')
p_no_entry.grid(row=2, column=1, padx=10, pady=5)

p_sn_label.grid(row=3, column=0, padx=10, pady=5, sticky='e')
p_sn_entry.grid(row=3, column=1, padx=10, pady=5)

engine_label.grid(row=4, column=0, padx=10, pady=5, sticky='e')
engine_entry.grid(row=4, column=1, padx=10, pady=5)

date_label.grid(row=5, column=0, padx=10, pady=5, sticky='e')
date_entry.grid(row=5, column=1, padx=10, pady=5)

ac_type_label.grid(row=6, column=0, padx=10, pady=5, sticky='e')
ac_type_entry.grid(row=6, column=1, padx=10, pady=5)

structure_label.grid(row=7, column=0, padx=10, pady=5, sticky='e')
structure_entry.grid(row=7, column=1, padx=10, pady=5)

pdf_viewer_button.grid(row=8, column=0, columnspan=2, padx=10, pady=10)
save_button.grid(row=9, column=0, columnspan=2, padx=10, pady=10)

# Configure grid column weights to make the entry fields expand with window resize
root.grid_columnconfigure(1, weight=1)

# Bind the window resize event to the on_resize function
root.bind("<Configure>", on_resize)

root.mainloop()
