import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import PyPDF2
import os
import clipboard
def compress_pdf():
    # Ask user to select a PDF file
    input_file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])

    if input_file_path:
        output_file_path = input_file_path.replace('.pdf', '_compressed.pdf')

        with open(input_file_path, 'rb') as input_file:
            pdf_reader = PyPDF2.PdfReader(input_file)
            pdf_writer = PyPDF2.PdfWriter()

            for page in pdf_reader.pages:
                pdf_writer.add_page(page)

            with open(output_file_path, 'wb') as output_file:
                pdf_writer.write(output_file)

        messagebox.showinfo("Success", f"PDF compressed successfully. Output saved to {output_file_path}")
def extract_text():
    input_path = filedialog.askopenfilename(title="Select PDF File", filetypes=[("PDF files", "*.pdf")])
    if input_path:
        pdf_text = ""
        with open(input_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            for page_num in range(len(pdf_reader.pages)):
                pdf_text += pdf_reader.pages[page_num].extract_text()

        # Clear and insert text without enabling editing state at first
        output_text.delete('1.0', tk.END)
        output_text.insert(tk.END, pdf_text)

        # Set text alignment and select all text
        output_text.tag_configure("center", justify="center")
        output_text.tag_add("sel", "1.0", "end")

        # Bind keyboard and right-click events for copying
        output_text.bind("<Control-a>", lambda e: output_text.tag_add("sel", "1.0", "end"))  # Ctrl+A to select all
        output_text.bind("<Button-3><ButtonRelease-3>", lambda e: copy_to_clipboard())  # Right-click to copy

        # Directly copy and then disable editing state
        copy_to_clipboard()
        output_text.config(state=tk.DISABLED)

        messagebox.showinfo("Success", "Text extracted and copied successfully.")

def copy_to_clipboard():
    selected_text = output_text.selection_get()
    clipboard.copy(selected_text)

def merge_pdfs():
    input_paths = []
    while True:
        path = filedialog.askopenfilename(title="Select PDF File to Merge", filetypes=[("PDF files", "*.pdf")])
        if not path:
            break  # User canceled or closed the dialog
        input_paths.append(path)

        # Check if user wants to add another file
        add_more = messagebox.askquestion("Add More Files", "Do you want to add another PDF file?", icon="question")
        if add_more != "yes":
            break

    if not input_paths:
        messagebox.showerror("Error", "Please select at least one PDF file to merge.")
        return

    output_path = filedialog.asksaveasfilename(title="Save Merged PDF As", defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
    if output_path:
        merger = PyPDF2.PdfWriter()
        for path in input_paths:
            with open(path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                for page_num in range(len(pdf_reader.pages)):
                    merger.add_page(pdf_reader.pages[page_num])
        with open(output_path, 'wb') as output_file:
            merger.write(output_file)
        messagebox.showinfo("Success", f"PDFs merged successfully. Output saved to {output_path}")

def split_pdf():
    input_path = filedialog.askopenfilename(title="Select PDF File to Split", filetypes=[("PDF files", "*.pdf")])
    if input_path:
        output_folder = filedialog.askdirectory(title="Select Output Folder for Split PDFs")
        if output_folder:
            pdf_reader = PyPDF2.PdfReader(input_path)
            for page_num in range(len(pdf_reader.pages)):
                output_path = os.path.join(output_folder, f"{os.path.splitext(os.path.basename(input_path))[0]}_{page_num + 1}.pdf")
                pdf_writer = PyPDF2.PdfWriter()
                pdf_writer.add_page(pdf_reader.pages[page_num])
                with open(output_path, 'wb') as output_file:
                    pdf_writer.write(output_file)
            messagebox.showinfo("Success", f"PDF split successfully. Pages saved to {output_folder}")

def rotate_pdf():
    input_path = filedialog.askopenfilename(title="Select PDF File to Rotate", filetypes=[("PDF files", "*.pdf")])
    if input_path:
        output_path = filedialog.asksaveasfilename(title="Save Rotated PDF As", defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], initialdir=os.path.dirname(input_path))
        if output_path:
            rotation_angle = simpledialog.askinteger("Rotation Angle", "Enter rotation angle (in degrees):")
            if rotation_angle is not None:
                pdf_writer = PyPDF2.PdfWriter()
                with open(input_path, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    for page_num in range(len(pdf_reader.pages)):
                        page = pdf_reader.pages[page_num]
                        page.rotate_clockwise(rotation_angle)
                        pdf_writer.add_page(page)
                with open(output_path, 'wb') as output_file:
                    pdf_writer.write(output_file)
                messagebox.showinfo("Success", f"PDF rotated successfully. Output saved to {output_path}")

def encrypt_pdf():
    input_path = filedialog.askopenfilename(title="Select PDF File to Encrypt", filetypes=[("PDF files", "*.pdf")])
    if input_path:
        output_path = filedialog.asksaveasfilename(title="Save Encrypted PDF As", defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], initialdir=os.path.dirname(input_path))
        if output_path:
            password = simpledialog.askstring("Password", "Enter password to encrypt PDF:")
            if password:
                pdf_writer = PyPDF2.PdfWriter()
                with open(input_path, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    for page_num in range(len(pdf_reader.pages)):
                        pdf_writer.add_page(pdf_reader.pages[page_num])
                    pdf_writer.encrypt(password)
                    with open(output_path, 'wb') as output_file:
                        pdf_writer.write(output_file)
                messagebox.showinfo("Success", f"PDF encrypted successfully. Output saved to {output_path}")

def decrypt_pdf():
    input_path = filedialog.askopenfilename(title="Select PDF File to Decrypt", filetypes=[("PDF files", "*.pdf")])
    if input_path:
        output_path = filedialog.asksaveasfilename(title="Save Decrypted PDF As", defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], initialdir=os.path.dirname(input_path))
        if output_path:
            password = simpledialog.askstring("Password", "Enter password to decrypt PDF:")
            if password:
                try:
                    with open(input_path, 'rb') as file:
                        pdf_reader = PyPDF2.PdfReader(file)
                        pdf_writer = PyPDF2.PdfWriter()
                        for page_num in range(len(pdf_reader.pages)):
                            pdf_writer.add_page(pdf_reader.pages[page_num])
                        pdf_writer.encrypt(password)
                        with open(output_path, 'wb') as output_file:
                            pdf_writer.write(output_file)
                    messagebox.showinfo("Success", f"PDF decrypted successfully. Output saved to {output_path}")
                except Exception as e:  # Handle decryption error
                    messagebox.showerror("Error", str(e))

def extract_images():
    input_path = filedialog.askopenfilename(title="Select PDF File to Extract Images", filetypes=[("PDF files", "*.pdf")])
    if input_path:
        output_folder = filedialog.askdirectory(title="Select Output Folder for Extracted Images")
        if output_folder:
            with open(input_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                for page_num in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[page_num]
                    xObject = page.get("/XObject")
                    if xObject:
                        for obj in xObject:
                            if xObject[obj].get("/Subtype") == "/Image":
                                size = (xObject[obj].get("/Width"), xObject[obj].get("/Height"))
                                data = xObject[obj].get_data()
                                image_path = os.path.join(output_folder, f"{os.path.splitext(os.path.basename(input_path))[0]}_page_{page_num + 1}_{obj[1:]}.png")
                                with open(image_path, 'wb') as image_file:
                                    image_file.write(data)
            messagebox.showinfo("Success", f"Images extracted successfully. Images saved to {output_folder}")

def add_watermark():
    input_path = filedialog.askopenfilename(title="Select PDF File to Add Watermark", filetypes=[("PDF files", "*.pdf")])
    if input_path:
        watermark_text = simpledialog.askstring("Watermark Text", "Enter Watermark Text:")
        if watermark_text:
            output_path = os.path.join(os.path.dirname(input_path), f"{os.path.splitext(os.path.basename(input_path))[0]}-watermark.pdf")
            pdf_writer = PyPDF2.PdfWriter()
            with open(input_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                for page_num in range(len(pdf_reader.pages)):
                    page = pdf_reader.pages[page_num]
                    pdf_writer.add_page(page)
                with open(output_path, 'wb') as output_file:
                    pdf_writer.write(output_file)
            messagebox.showinfo("Success", f"Watermark added successfully. Output saved to {output_path}")

# Create a Tkinter window
root = tk.Tk()
root.title("PyPDF2 Operations")

# Frame for output text
output_frame = tk.Frame(root)
output_frame.pack()

# Output text widget
output_text = tk.Text(output_frame, height=10, width=50)
output_text.pack()

# Frame for buttons in two columns
button_frame_left = tk.Frame(root)
button_frame_left.pack(side=tk.TOP, padx=5, pady=5)

button_frame_right = tk.Frame(root)
button_frame_right.pack(side=tk.TOP, padx=5, pady=5)


# Compression button
compress_button = tk.Button(root, text="Compress PDF", command=compress_pdf)
compress_button.pack(pady=20)

# Extract button
extract_button = tk.Button(button_frame_left, text="Extract Text", command=extract_text)
extract_button.pack(side=tk.TOP, padx=5, pady=5)

merge_button = tk.Button(button_frame_left, text="Merge PDFs", command=merge_pdfs)
merge_button.pack(side=tk.TOP, padx=5, pady=5)

split_button = tk.Button(button_frame_left, text="Split PDF", command=split_pdf)
split_button.pack(side=tk.TOP, padx=5, pady=5)

rotate_button = tk.Button(button_frame_left, text="Rotate PDF", command=rotate_pdf)
rotate_button.pack(side=tk.TOP, padx=5, pady=5)

# Buttons - Right column
encrypt_button = tk.Button(button_frame_right, text="Encrypt PDF", command=encrypt_pdf)
encrypt_button.pack(side=tk.TOP, padx=5, pady=5)

decrypt_button = tk.Button(button_frame_right, text="Decrypt PDF", command=decrypt_pdf)
decrypt_button.pack(side=tk.TOP, padx=5, pady=5)

extract_images_button = tk.Button(button_frame_right, text="Extract Images", command=extract_images)
extract_images_button.pack(side=tk.TOP, padx=5, pady=5)

add_watermark_button = tk.Button(button_frame_right, text="Add Watermark", command=add_watermark)
add_watermark_button.pack(side=tk.TOP, padx=5, pady=5)

root.mainloop()