# Wormhole File Converter
# This program provides a simple GUI for converting files between formats.
# Requires additional libraries: pip install pillow reportlab
# Pillow for image handling, ReportLab for PDF generation.

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

# Distinctly labeled defs for each converter

def convert_jpg_to_png():
    """
    Converter for JPG to PNG.
    Opens a file dialog to select a JPG file and converts it to PNG.
    """
    file_path = filedialog.askopenfilename(title="Select JPG File", filetypes=[("JPG files", "*.jpg;*.jpeg")])
    if file_path:
        try:
            img = Image.open(file_path)
            new_file_path = file_path.rsplit('.', 1)[0] + '.png'
            img.save(new_file_path, 'PNG')
            messagebox.showinfo("Success", f"File converted to: {new_file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")

def convert_png_to_jpg():
    """
    Converter for PNG to JPG.
    Opens a file dialog to select a PNG file and converts it to JPG.
    """
    file_path = filedialog.askopenfilename(title="Select PNG File", filetypes=[("PNG files", "*.png")])
    if file_path:
        try:
            img = Image.open(file_path)
            new_file_path = file_path.rsplit('.', 1)[0] + '.jpg'
            # Convert to RGB mode if necessary (JPG doesn't support alpha channel)
            if img.mode in ('RGBA', 'LA'):
                img = img.convert('RGB')
            img.save(new_file_path, 'JPEG')
            messagebox.showinfo("Success", f"File converted to: {new_file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")

def convert_txt_to_pdf():
    """
    Converter for TXT to PDF.
    Opens a file dialog to select a TXT file and converts it to PDF.
    """
    file_path = filedialog.askopenfilename(title="Select TXT File", filetypes=[("Text files", "*.txt")])
    if file_path:
        try:
            new_file_path = file_path.rsplit('.', 1)[0] + '.pdf'
            c = canvas.Canvas(new_file_path, pagesize=letter)
            width, height = letter
            y = height - 50  # Start from top with margin
            with open(file_path, 'r') as f:
                for line in f:
                    c.drawString(50, y, line.strip())
                    y -= 15  # Line spacing
                    if y < 50:  # Simple page break handling
                        c.showPage()
                        y = height - 50
            c.save()
            messagebox.showinfo("Success", f"File converted to: {new_file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")

# Functions to open subwindows for each category

def open_docs_window():
    docs_win = tk.Toplevel()
    docs_win.title("Docs Conversions")
    docs_win.geometry("300x200")
    docs_win.configure(bg="#f0f0f0")

    label = tk.Label(docs_win, text="Select Docs Conversion:", bg="#f0f0f0", font=("Arial", 12))
    label.pack(pady=10)

    btn_txt_to_pdf = tk.Button(docs_win, text="Convert TXT to PDF", command=convert_txt_to_pdf, width=25, bg="#FFC107", fg="black", font=("Arial", 10))
    btn_txt_to_pdf.pack(pady=5)

    # Add more doc converters here if needed

def open_presentations_window():
    pres_win = tk.Toplevel()
    pres_win.title("Presentations Conversions")
    pres_win.geometry("300x200")
    pres_win.configure(bg="#f0f0f0")

    label = tk.Label(pres_win, text="Presentations conversions coming soon!", bg="#f0f0f0", font=("Arial", 12))
    label.pack(pady=10)

    # Placeholder for future converters

def open_images_window():
    img_win = tk.Toplevel()
    img_win.title("Images Conversions")
    img_win.geometry("300x200")
    img_win.configure(bg="#f0f0f0")

    label = tk.Label(img_win, text="Select Images Conversion:", bg="#f0f0f0", font=("Arial", 12))
    label.pack(pady=10)

    btn_jpg_to_png = tk.Button(img_win, text="Convert JPG to PNG", command=convert_jpg_to_png, width=25, bg="#4CAF50", fg="white", font=("Arial", 10))
    btn_jpg_to_png.pack(pady=5)

    btn_png_to_jpg = tk.Button(img_win, text="Convert PNG to JPG", command=convert_png_to_jpg, width=25, bg="#2196F3", fg="white", font=("Arial", 10))
    btn_png_to_jpg.pack(pady=5)

    # Add more image converters here if needed

def open_videos_window():
    vid_win = tk.Toplevel()
    vid_win.title("Videos Conversions")
    vid_win.geometry("300x200")
    vid_win.configure(bg="#f0f0f0")

    label = tk.Label(vid_win, text="Videos conversions coming soon!", bg="#f0f0f0", font=("Arial", 12))
    label.pack(pady=10)

    # Placeholder for future converters

def open_audio_window():
    aud_win = tk.Toplevel()
    aud_win.title("Audio Conversions")
    aud_win.geometry("300x200")
    aud_win.configure(bg="#f0f0f0")

    label = tk.Label(aud_win, text="Audio conversions coming soon!", bg="#f0f0f0", font=("Arial", 12))
    label.pack(pady=10)

    # Placeholder for future converters

def open_archive_window():
    arch_win = tk.Toplevel()
    arch_win.title("Archive Conversions")
    arch_win.geometry("300x200")
    arch_win.configure(bg="#f0f0f0")

    label = tk.Label(arch_win, text="Archive conversions coming soon!", bg="#f0f0f0", font=("Arial", 12))
    label.pack(pady=10)

    # Placeholder for future converters

def open_other_window():
    other_win = tk.Toplevel()
    other_win.title("Other Conversions")
    other_win.geometry("300x200")
    other_win.configure(bg="#f0f0f0")

    label = tk.Label(other_win, text="Other conversions coming soon!", bg="#f0f0f0", font=("Arial", 12))
    label.pack(pady=10)

    # Placeholder for future converters

# Main UI section
def main_ui():
    root = tk.Tk()
    root.title("Wormhole File Converter")
    root.geometry("400x400")  # Adjusted window size
    root.configure(bg="#f0f0f0")  # Custom background color

    # Custom label for instructions
    label = tk.Label(root, text="Select a file type category:", bg="#f0f0f0", font=("Arial", 12))
    label.pack(pady=20)

    # Buttons for each category with custom styling
    btn_docs = tk.Button(root, text="Docs", command=open_docs_window, width=30, bg="#FFC107", fg="black", font=("Arial", 10))
    btn_docs.pack(pady=5)

    btn_presentations = tk.Button(root, text="Presentations", command=open_presentations_window, width=30, bg="#FF5722", fg="white", font=("Arial", 10))
    btn_presentations.pack(pady=5)

    btn_images = tk.Button(root, text="Images", command=open_images_window, width=30, bg="#4CAF50", fg="white", font=("Arial", 10))
    btn_images.pack(pady=5)

    btn_videos = tk.Button(root, text="Videos", command=open_videos_window, width=30, bg="#2196F3", fg="white", font=("Arial", 10))
    btn_videos.pack(pady=5)

    btn_audio = tk.Button(root, text="Audio", command=open_audio_window, width=30, bg="#9C27B0", fg="white", font=("Arial", 10))
    btn_audio.pack(pady=5)

    btn_archive = tk.Button(root, text="Archive", command=open_archive_window, width=30, bg="#607D8B", fg="white", font=("Arial", 10))
    btn_archive.pack(pady=5)

    btn_other = tk.Button(root, text="Other", command=open_other_window, width=30, bg="#795548", fg="white", font=("Arial", 10))
    btn_other.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    main_ui()