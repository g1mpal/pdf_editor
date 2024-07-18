import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from tkinter import ttk
from pdf_operations import merge_pdfs, delete_page, split_pdf, rotate_page, pdf_to_word

def select_files():
    files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
    return root.tk.splitlist(files)

def save_file(extension=".pdf"):
    save_path = filedialog.asksaveasfilename(defaultextension=extension, filetypes=[("PDF files", f"*{extension}")])
    return save_path

def merge_action():
    file_paths = select_files()
    if file_paths:
        output_path = save_file()
        if output_path:
            merge_pdfs(file_paths, output_path)
            messagebox.showinfo("Success", "PDFs merged successfully!")

def delete_page_action():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        page_number = simpledialog.askinteger("Input", "Enter the page number to delete:")
        if page_number is not None:
            output_path = save_file()
            if output_path:
                try:
                    delete_page(file_path, page_number, output_path)
                    messagebox.showinfo("Success", "Page deleted successfully!")
                except ValueError as e:
                    messagebox.showerror("Error", str(e))

def split_pdf_action():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        split_page = simpledialog.askinteger("Input", "Enter the page number to split at:")
        if split_page is not None:
            output_path1 = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], title="Save first part as")
            output_path2 = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], title="Save second part as")
            if output_path1 and output_path2:
                try:
                    split_pdf(file_path, split_page, output_path1, output_path2)
                    messagebox.showinfo("Success", "PDF split successfully!")
                except ValueError as e:
                    messagebox.showerror("Error", str(e))

def rotate_page_action():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        page_number = simpledialog.askinteger("Input", "Enter the page number to rotate:")
        if page_number is not None:
            rotation_angle = simpledialog.askinteger("Input", "Enter the rotation angle (90, 180, 270):")
            if rotation_angle in [90, 180, 270]:
                output_path = save_file()
                if output_path:
                    try:
                        rotate_page(file_path, page_number, rotation_angle, output_path)
                        messagebox.showinfo("Success", "Page rotated successfully!")
                    except ValueError as e:
                        messagebox.showerror("Error", str(e))
            else:
                messagebox.showerror("Error", "Invalid rotation angle. Must be 90, 180, or 270 degrees.")

def pdf_to_word_action():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        output_path = save_file(extension=".docx")
        if output_path:
            pdf_to_word(file_path, output_path)
            messagebox.showinfo("Success", "PDF converted to Word successfully!")

def create_gui():
    global root
    root = tk.Tk()
    root.title("PDF Editor")
    root.geometry("500x500")
    root.style = ttk.Style()
    root.style.theme_use('clam')  # Use the clam theme

    # Set unified background color
    bg_color = '#d9d9d9'
    root.configure(bg=bg_color)

    # Custom styles
    root.style.configure('TButton', font=('Ubuntu', 12), padding=10)
    root.style.map('TButton', background=[('active', '#9B194B'), ('!active', '#B1B1B1')])

    # Create frames for better organization
    frame = ttk.Frame(root, padding="20 20 20 20", style="TFrame")
    frame.grid(row=0, column=0, sticky=(tk.W, tk.E))
    root.grid_columnconfigure(0, weight=1)
    root.grid_rowconfigure(0, weight=1)

    # Set the same background color for the frame
    frame.configure(style="TFrame")
    root.style.configure("TFrame", background=bg_color)

    # Center content inside frame
    frame.grid_columnconfigure(0, weight=1)
    frame.grid_rowconfigure(0, weight=1)

    # Add buttons with grid layout
    button_width = 25

    ttk.Button(frame, text="Merge PDFs", command=merge_action, width=button_width).grid(row=0, column=0, pady=5, padx=10)
    ttk.Button(frame, text="Delete Page from PDF", command=delete_page_action, width=button_width).grid(row=1, column=0, pady=5, padx=10)
    ttk.Button(frame, text="Split PDF", command=split_pdf_action, width=button_width).grid(row=2, column=0, pady=5, padx=10)
    ttk.Button(frame, text="Rotate Page in PDF", command=rotate_page_action, width=button_width).grid(row=3, column=0, pady=5, padx=10)
    ttk.Button(frame, text="Convert PDF to Word", command=pdf_to_word_action, width=button_width).grid(row=4, column=0, pady=5, padx=10)

    root.mainloop()

if __name__ == "__main__":
    create_gui()
