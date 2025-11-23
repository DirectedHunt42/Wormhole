# Wormhole File Converter
# This program provides a simple GUI for converting files between formats.
# Requires additional libraries: pip install pillow reportlab customtkinter
# Pillow for image handling, ReportLab for PDF generation.

import os
import sys
import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

# Theme colors from Black Hole style
BG = "#05050a"
CARD = "#0b0b0f"
CARD_HOVER = "#111327"
ACCENT = "#47a3ff"
ACCENT_DIM = "#2b6f9f"
TEXT = "#e6eef8"

# Set up customtkinter
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

class WormholeApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Wormhole File Converter")
        self.geometry("400x400")
        self.configure(fg_color=BG)
        self._build_ui()

    def _build_ui(self):
        # Custom label for instructions
        label = ctk.CTkLabel(self, text="Select a file type category:", fg_color=BG, text_color=TEXT, font=("Arial", 12))
        label.pack(pady=20)

        # Buttons for each category
        btn_docs = ctk.CTkButton(self, text="Docs", command=self.open_docs_window, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=300, font=("Arial", 10))
        btn_docs.pack(pady=5)

        btn_presentations = ctk.CTkButton(self, text="Presentations", command=self.open_presentations_window, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=300, font=("Arial", 10))
        btn_presentations.pack(pady=5)

        btn_images = ctk.CTkButton(self, text="Images", command=self.open_images_window, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=300, font=("Arial", 10))
        btn_images.pack(pady=5)

        btn_videos = ctk.CTkButton(self, text="Videos", command=self.open_videos_window, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=300, font=("Arial", 10))
        btn_videos.pack(pady=5)

        btn_audio = ctk.CTkButton(self, text="Audio", command=self.open_audio_window, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=300, font=("Arial", 10))
        btn_audio.pack(pady=5)

        btn_archive = ctk.CTkButton(self, text="Archive", command=self.open_archive_window, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=300, font=("Arial", 10))
        btn_archive.pack(pady=5)

        btn_other = ctk.CTkButton(self, text="Other", command=self.open_other_window, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=300, font=("Arial", 10))
        btn_other.pack(pady=5)

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

def open_docs_window(master):
    docs_win = ctk.CTkToplevel(master)
    docs_win.title("Docs Conversions")
    docs_win.geometry("300x200")
    docs_win.configure(fg_color=BG)

    label = ctk.CTkLabel(docs_win, text="Select Docs Conversion:", fg_color=BG, text_color=TEXT, font=("Arial", 12))
    label.pack(pady=10)

    btn_txt_to_pdf = ctk.CTkButton(docs_win, text="Convert TXT to PDF", command=convert_txt_to_pdf, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=250, font=("Arial", 10))
    btn_txt_to_pdf.pack(pady=5)

    # Add more doc converters here if needed

def open_presentations_window(master):
    pres_win = ctk.CTkToplevel(master)
    pres_win.title("Presentations Conversions")
    pres_win.geometry("300x200")
    pres_win.configure(fg_color=BG)

    label = ctk.CTkLabel(pres_win, text="Presentations conversions coming soon!", fg_color=BG, text_color=TEXT, font=("Arial", 12))
    label.pack(pady=10)

    # Placeholder for future converters

def open_images_window(master):
    img_win = ctk.CTkToplevel(master)
    img_win.title("Images Conversions")
    img_win.geometry("300x200")
    img_win.configure(fg_color=BG)

    label = ctk.CTkLabel(img_win, text="Select Images Conversion:", fg_color=BG, text_color=TEXT, font=("Arial", 12))
    label.pack(pady=10)

    btn_jpg_to_png = ctk.CTkButton(img_win, text="Convert JPG to PNG", command=convert_jpg_to_png, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=250, font=("Arial", 10))
    btn_jpg_to_png.pack(pady=5)

    btn_png_to_jpg = ctk.CTkButton(img_win, text="Convert PNG to JPG", command=convert_png_to_jpg, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=250, font=("Arial", 10))
    btn_png_to_jpg.pack(pady=5)

    # Add more image converters here if needed

def open_videos_window(master):
    vid_win = ctk.CTkToplevel(master)
    vid_win.title("Videos Conversions")
    vid_win.geometry("300x200")
    vid_win.configure(fg_color=BG)

    label = ctk.CTkLabel(vid_win, text="Videos conversions coming soon!", fg_color=BG, text_color=TEXT, font=("Arial", 12))
    label.pack(pady=10)

    # Placeholder for future converters

def open_audio_window(master):
    aud_win = ctk.CTkToplevel(master)
    aud_win.title("Audio Conversions")
    aud_win.geometry("300x200")
    aud_win.configure(fg_color=BG)

    label = ctk.CTkLabel(aud_win, text="Audio conversions coming soon!", fg_color=BG, text_color=TEXT, font=("Arial", 12))
    label.pack(pady=10)

    # Placeholder for future converters

def open_archive_window(master):
    arch_win = ctk.CTkToplevel(master)
    arch_win.title("Archive Conversions")
    arch_win.geometry("300x200")
    arch_win.configure(fg_color=BG)

    label = ctk.CTkLabel(arch_win, text="Archive conversions coming soon!", fg_color=BG, text_color=TEXT, font=("Arial", 12))
    label.pack(pady=10)

    # Placeholder for future converters

def open_other_window(master):
    other_win = ctk.CTkToplevel(master)
    other_win.title("Other Conversions")
    other_win.geometry("300x200")
    other_win.configure(fg_color=BG)

    label = ctk.CTkLabel(other_win, text="Other conversions coming soon!", fg_color=BG, text_color=TEXT, font=("Arial", 12))
    label.pack(pady=10)

    # Placeholder for future converters

# Extend the app class with open methods
class WormholeApp(WormholeApp):
    def open_docs_window(self):
        open_docs_window(self)

    def open_presentations_window(self):
        open_presentations_window(self)

    def open_images_window(self):
        open_images_window(self)

    def open_videos_window(self):
        open_videos_window(self)

    def open_audio_window(self):
        open_audio_window(self)

    def open_archive_window(self):
        open_archive_window(self)

    def open_other_window(self):
        open_other_window(self)

if __name__ == "__main__":
    app = WormholeApp()
    app.mainloop()