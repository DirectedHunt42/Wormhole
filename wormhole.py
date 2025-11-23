import os
import sys
import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from pypdf import PdfReader
import py7zr
import zipfile
import tarfile
import tempfile
import shutil
import threading
from docx import Document
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Inches
import openpyxl
import csv
try:
    from striprtf.striprtf import rtf_to_text
    RTF_SUPPORT = True
except ImportError:
    RTF_SUPPORT = False
import ezodf

BG = "#0a0812"
CARD = "#120f1e"
CARD_HOVER = "#19172b"
ACCENT = "#7aa3ff"
ACCENT_DIM = "#4d6bbc"
TEXT = "#e8e6f5"

WORMHOLE_IMAGE_PATH = os.path.join("Icons", "wormhole_Transparent_Light.png")
try:
    WORMHOLE_PIL_IMAGE = Image.open(WORMHOLE_IMAGE_PATH)
except Exception as e:
    print(f"Could not load wormhole image: {e}")
    WORMHOLE_PIL_IMAGE = Image.new("RGBA", (100, 100), (100, 100, 100, 255))

APP_ICON_PATH = os.path.join("Icons", "Wormhole_Icon.ico")

# Paths for custom fonts (adjust family and file names if needed; assumes Roboto as example)
FONTS_DIR = "fonts"
FONT_FAMILY_REGULAR = "Pathway Extreme 36pt Regular"
FONT_FAMILY_SEMIBOLD = "Pathway Extreme 36pt SemiBold"
FONT_FAMILY_ITALIC = "Pathway Extreme 36pt Italic"
FONT_FAMILY_THIN = "Pathway Extreme 36pt Thin"
FONT_FAMILY_BLACK = "Pathway Extreme 36pt Black"
FONT_FILES = [
    "PathwayExtreme_36pt-Black.ttf",
    "PathwayExtreme_36pt-Italic.ttf",
    "PathwayExtreme_36pt-Regular.ttf",
    "PathwayExtreme_36pt-SemiBold.ttf",
    "PathwayExtreme_36pt-Thin.ttf"
]

# Set up customtkinter
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

# Load custom fonts
for font_file in FONT_FILES:
    font_path = os.path.join(FONTS_DIR, font_file)
    if os.path.exists(font_path):
        ctk.FontManager.load_font(font_path)
    else:
        print(f"Custom font file not found: {font_path}; falling back to default for this variant.")

class WormholeApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Wormhole File Converter")
        self.geometry("400x650")
        self.configure(fg_color=BG)
        # Center the main window
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width // 2) - (400 // 2)
        y = (screen_height // 2) - (650 // 2)
        self.geometry(f"400x650+{x}+{y}")
        self._build_ui()

    def _build_ui(self):
        if os.path.exists(APP_ICON_PATH):
            try:
                if sys.platform.startswith('win'):
                    import ctypes
                    ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("wormhole.file.converter")
                self.iconbitmap(APP_ICON_PATH)
            except Exception as e:
                print(f"Could not set application icon: {e}")
                
        # Custom label for instructions
        image = ctk.CTkImage(light_image=WORMHOLE_PIL_IMAGE, dark_image=WORMHOLE_PIL_IMAGE, size=(306, 204))
        img_label = ctk.CTkLabel(self, image=image, text="", fg_color=BG)
        img_label.pack(pady=10)

        label = ctk.CTkLabel(self, text="Select a file type category:", fg_color=BG, text_color=TEXT, font=(FONT_FAMILY_REGULAR, 12))
        label.pack(pady=20)

        # Buttons for each category (using semibold for buttons if desired; otherwise keep normal)
        btn_docs = ctk.CTkButton(self, text="Docs", command=self.open_docs_window, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=300, font=(FONT_FAMILY_SEMIBOLD, 20))
        btn_docs.pack(pady=5)

        btn_presentations = ctk.CTkButton(self, text="Presentations", command=self.open_presentations_window, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=300, font=(FONT_FAMILY_SEMIBOLD, 20))
        btn_presentations.pack(pady=5)

        btn_images = ctk.CTkButton(self, text="Images", command=self.open_images_window, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=300, font=(FONT_FAMILY_SEMIBOLD, 20))
        btn_images.pack(pady=5)

        btn_archive = ctk.CTkButton(self, text="Archive", command=self.open_archive_window, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=300, font=(FONT_FAMILY_SEMIBOLD, 20))
        btn_archive.pack(pady=5)

        btn_spreadsheets = ctk.CTkButton(self, text="Spreadsheets", command=self.open_spreadsheets_window, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=300, font=(FONT_FAMILY_SEMIBOLD, 20))
        btn_spreadsheets.pack(pady=5)

# Functions to open subwindows for each category

def open_docs_window(master):
    docs_win = ctk.CTkToplevel(master)
    docs_win.title("Docs Conversions")
    docs_win.geometry("300x300")
    docs_win.configure(fg_color=BG)
    # Center the window
    docs_win.update_idletasks()
    screen_width = docs_win.winfo_screenwidth()
    screen_height = docs_win.winfo_screenheight()
    x = (screen_width // 2) - (300 // 2)
    y = (screen_height // 2) - (300 // 2)
    docs_win.geometry(f"300x300+{x}+{y}")
    # Set icon
    if os.path.exists(APP_ICON_PATH):
        try:
            docs_win.after(250, lambda: docs_win.iconbitmap(APP_ICON_PATH))
        except Exception as e:
            print(f"Could not set icon for docs window: {e}")
    # Make it transient and grab set to stay on top
    docs_win.transient(master)
    docs_win.grab_set()

    label = ctk.CTkLabel(docs_win, text="Docs Converter", fg_color=BG, text_color=TEXT, font=(FONT_FAMILY_REGULAR, 12))
    label.pack(pady=10)

    file_path_var = ctk.StringVar(value="")

    filetypes = [("Docs files", "*.txt;*.pdf;*.docx;*.html;*.md;*.odt")]
    if RTF_SUPPORT:
        filetypes[0] = ("Docs files", "*.txt;*.pdf;*.docx;*.html;*.md;*.odt;*.rtf")

    def select_file():
        fp = filedialog.askopenfilename(title="Select Docs File", filetypes=filetypes)
        if fp:
            file_path_var.set(fp)
            file_label.configure(text=os.path.basename(fp))

    btn_select = ctk.CTkButton(docs_win, text="Select File", command=select_file, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=250, font=(FONT_FAMILY_SEMIBOLD, 10))
    btn_select.pack(pady=5)

    file_label = ctk.CTkLabel(docs_win, text="No file selected", fg_color=BG, text_color=TEXT, font=(FONT_FAMILY_REGULAR, 10))
    file_label.pack(pady=5)

    target_var = ctk.StringVar(value="PDF")
    combo_values = ["PDF", "TXT", "DOCX", "HTML", "MD", "ODT"]
    if RTF_SUPPORT:
        combo_values.append("RTF")
    combo = ctk.CTkComboBox(docs_win, values=combo_values, variable=target_var, font=(FONT_FAMILY_REGULAR, 10), width=250)
    combo.pack(pady=5)

    progress_bar = ctk.CTkProgressBar(docs_win, width=250, mode="indeterminate")
    # Initially not packed

    def do_convert():
        fp = file_path_var.get()
        if not fp:
            messagebox.showerror("Error", "No file selected")
            return
        target = target_var.get().upper()
        input_ext = os.path.splitext(fp)[1].lower()[1:]
        if target.lower() == input_ext:
            messagebox.showwarning("Warning", "Input and output formats are the same")
            return
        new_file_path = os.path.splitext(fp)[0] + '.' + target.lower()

        def conversion_thread():
            try:
                text = ""
                if input_ext in ["txt", "md"]:
                    with open(fp, 'r') as f:
                        text = f.read()
                elif input_ext == "pdf":
                    reader = PdfReader(fp)
                    text = ''
                    for page in reader.pages:
                        text += page.extract_text() + '\n'
                elif input_ext == "docx":
                    doc = Document(fp)
                    text = '\n'.join([para.text for para in doc.paragraphs])
                elif input_ext == "html":
                    with open(fp, 'r') as f:
                        soup = BeautifulSoup(f.read(), 'html.parser')
                        text = soup.get_text()
                elif input_ext == "odt":
                    doc = ezodf.opendoc(fp)
                    text = '\n'.join(obj.text for obj in doc.body if obj.kind == 'Paragraph')
                elif input_ext == "rtf" and RTF_SUPPORT:
                    with open(fp, 'r') as f:
                        rtf = f.read()
                    text = rtf_to_text(rtf)
                else:
                    docs_win.after(0, lambda: messagebox.showerror("Error", "Unsupported input format"))
                    return

                if target in ["TXT", "MD"]:
                    with open(new_file_path, 'w') as f:
                        f.write(text)
                elif target == "PDF":
                    c = canvas.Canvas(new_file_path, pagesize=letter)
                    width, height = letter
                    y = height - 50  # Start from top with margin
                    for line in text.splitlines():
                        c.drawString(50, y, line.strip())
                        y -= 15  # Line spacing
                        if y < 50:  # Simple page break handling
                            c.showPage()
                            y = height - 50
                    c.save()
                elif target == "DOCX":
                    doc = Document()
                    doc.add_paragraph(text)
                    doc.save(new_file_path)
                elif target == "HTML":
                    with open(new_file_path, 'w') as f:
                        escaped_text = text.replace('<', '&lt;').replace('>', '&gt;')
                        f.write(f"<html><body><pre>{escaped_text}</pre></body></html>")
                elif target == "ODT":
                    doc = ezodf.newdoc(doctype='odt', filename=new_file_path)
                    doc.body.append(ezodf.Paragraph(text))
                    doc.save()
                elif target == "RTF" and RTF_SUPPORT:
                    docs_win.after(0, lambda: messagebox.showerror("Error", "RTF output not supported"))
                    return
                else:
                    docs_win.after(0, lambda: messagebox.showerror("Error", "Unsupported target format"))
                    return
                docs_win.after(0, lambda: messagebox.showinfo("Success", f"File converted to: {new_file_path}"))
            except Exception as e:
                docs_win.after(0, lambda: messagebox.showerror("Error", f"Conversion failed: {str(e)}"))
            finally:
                docs_win.after(0, progress_bar.stop)
                docs_win.after(0, progress_bar.pack_forget)
                docs_win.after(0, lambda: btn_convert.configure(state="normal"))

        progress_bar.pack(pady=5)
        progress_bar.start()
        btn_convert.configure(state="disabled")
        thread = threading.Thread(target=conversion_thread)
        thread.start()

    btn_convert = ctk.CTkButton(docs_win, text="Convert", command=do_convert, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=250, font=(FONT_FAMILY_SEMIBOLD, 10))
    btn_convert.pack(pady=5)

def open_presentations_window(master):
    pres_win = ctk.CTkToplevel(master)
    pres_win.title("Presentations Conversions")
    pres_win.geometry("300x300")
    pres_win.configure(fg_color=BG)
    # Center the window
    pres_win.update_idletasks()
    screen_width = pres_win.winfo_screenwidth()
    screen_height = pres_win.winfo_screenheight()
    x = (screen_width // 2) - (300 // 2)
    y = (screen_height // 2) - (300 // 2)
    pres_win.geometry(f"300x300+{x}+{y}")
    # Set icon
    if os.path.exists(APP_ICON_PATH):
        try:
            pres_win.after(250, lambda: pres_win.iconbitmap(APP_ICON_PATH))
        except Exception as e:
            print(f"Could not set icon for presentations window: {e}")
    # Make it transient and grab set to stay on top
    pres_win.transient(master)
    pres_win.grab_set()

    label = ctk.CTkLabel(pres_win, text="Presentations Converter", fg_color=BG, text_color=TEXT, font=(FONT_FAMILY_REGULAR, 12))
    label.pack(pady=10)

    file_path_var = ctk.StringVar(value="")

    def select_file():
        fp = filedialog.askopenfilename(title="Select Presentation File", filetypes=[("Presentation files", "*.pptx;*.odp")])
        if fp:
            file_path_var.set(fp)
            file_label.configure(text=os.path.basename(fp))

    btn_select = ctk.CTkButton(pres_win, text="Select File", command=select_file, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=250, font=(FONT_FAMILY_SEMIBOLD, 10))
    btn_select.pack(pady=5)

    file_label = ctk.CTkLabel(pres_win, text="No file selected", fg_color=BG, text_color=TEXT, font=(FONT_FAMILY_REGULAR, 10))
    file_label.pack(pady=5)

    target_var = ctk.StringVar(value="PDF")
    combo = ctk.CTkComboBox(pres_win, values=["PPTX", "PDF", "TXT", "DOCX", "ODP"], variable=target_var, font=(FONT_FAMILY_REGULAR, 10), width=250)
    combo.pack(pady=5)

    progress_bar = ctk.CTkProgressBar(pres_win, width=250, mode="indeterminate")
    # Initially not packed

    def do_convert():
        fp = file_path_var.get()
        if not fp:
            messagebox.showerror("Error", "No file selected")
            return
        target = target_var.get().upper()
        input_ext = os.path.splitext(fp)[1].lower()[1:]
        if target.lower() == input_ext:
            messagebox.showwarning("Warning", "Input and output formats are the same")
            return
        new_file_path = os.path.splitext(fp)[0] + '.' + target.lower()

        def conversion_thread():
            try:
                text = ""
                if input_ext == "pptx":
                    pres = Presentation(fp)
                    text = ''
                    for slide in pres.slides:
                        for shape in slide.shapes:
                            if shape.has_text_frame:
                                for paragraph in shape.text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        text += run.text
                                    text += '\n'
                elif input_ext == "odp":
                    doc = ezodf.opendoc(fp)
                    text = '\n'.join(obj.text for obj in doc.body if obj.kind == 'Paragraph')
                else:
                    pres_win.after(0, lambda: messagebox.showerror("Error", "Unsupported input format"))
                    return

                if target == "TXT":
                    with open(new_file_path, 'w') as f:
                        f.write(text)
                elif target == "PDF":
                    c = canvas.Canvas(new_file_path, pagesize=letter)
                    width, height = letter
                    y = height - 50  # Start from top with margin
                    for line in text.splitlines():
                        c.drawString(50, y, line.strip())
                        y -= 15  # Line spacing
                        if y < 50:  # Simple page break handling
                            c.showPage()
                            y = height - 50
                    c.save()
                elif target == "DOCX":
                    doc = Document()
                    doc.add_paragraph(text)
                    doc.save(new_file_path)
                elif target == "PPTX":
                    pres = Presentation()
                    slide_layout = pres.slide_layouts[0]
                    slide = pres.slides.add_slide(slide_layout)
                    left = top = Inches(1)
                    width = height = Inches(6)
                    txBox = slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    tf.text = text
                    pres.save(new_file_path)
                elif target == "ODP":
                    doc = ezodf.newdoc(doctype='odp', filename=new_file_path)
                    doc.body.append(ezodf.Paragraph(text))
                    doc.save()
                else:
                    pres_win.after(0, lambda: messagebox.showerror("Error", "Unsupported target format"))
                    return
                pres_win.after(0, lambda: messagebox.showinfo("Success", f"File converted to: {new_file_path}"))
            except Exception as e:
                pres_win.after(0, lambda: messagebox.showerror("Error", f"Conversion failed: {str(e)}"))
            finally:
                pres_win.after(0, progress_bar.stop)
                pres_win.after(0, progress_bar.pack_forget)
                pres_win.after(0, lambda: btn_convert.configure(state="normal"))

        progress_bar.pack(pady=5)
        progress_bar.start()
        btn_convert.configure(state="disabled")
        thread = threading.Thread(target=conversion_thread)
        thread.start()

    btn_convert = ctk.CTkButton(pres_win, text="Convert", command=do_convert, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=250, font=(FONT_FAMILY_SEMIBOLD, 10))
    btn_convert.pack(pady=5)

def open_images_window(master):
    img_win = ctk.CTkToplevel(master)
    img_win.title("Images Conversions")
    img_win.geometry("300x450")
    img_win.configure(fg_color=BG)
    # Center the window
    img_win.update_idletasks()
    screen_width = img_win.winfo_screenwidth()
    screen_height = img_win.winfo_screenheight()
    x = (screen_width // 2) - (300 // 2)
    y = (screen_height // 2) - (450 // 2)
    img_win.geometry(f"300x450+{x}+{y}")
    # Set icon
    if os.path.exists(APP_ICON_PATH):
        try:
            img_win.after(250, lambda: img_win.iconbitmap(APP_ICON_PATH))
        except Exception as e:
            print(f"Could not set icon for images window: {e}")
    # Make it transient and grab set to stay on top
    img_win.transient(master)
    img_win.grab_set()

    label = ctk.CTkLabel(img_win, text="Images Converter", fg_color=BG, text_color=TEXT, font=(FONT_FAMILY_REGULAR, 12))
    label.pack(pady=10)

    file_path_var = ctk.StringVar(value="")

    def select_file():
        fp = filedialog.askopenfilename(title="Select Image File", filetypes=[("Image files", "*.jpg;*.jpeg;*.png;*.webp;*.avif;*.ico;*.bmp;*.gif;*.tiff")])
        if fp:
            file_path_var.set(fp)
            file_label.configure(text=os.path.basename(fp))

    btn_select = ctk.CTkButton(img_win, text="Select File", command=select_file, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=250, font=(FONT_FAMILY_SEMIBOLD, 10))
    btn_select.pack(pady=5)

    file_label = ctk.CTkLabel(img_win, text="No file selected", fg_color=BG, text_color=TEXT, font=(FONT_FAMILY_REGULAR, 10))
    file_label.pack(pady=5)

    target_var = ctk.StringVar(value="PNG")
    combo = ctk.CTkComboBox(img_win, values=["PNG", "JPG", "WEBP", "AVIF", "ICO", "BMP", "GIF", "TIFF"], variable=target_var, font=(FONT_FAMILY_REGULAR, 10), width=250)
    combo.pack(pady=5)

    ico_frame = ctk.CTkFrame(img_win, fg_color=BG)
    ico_label = ctk.CTkLabel(ico_frame, text="Select ICO sizes:", fg_color=BG, text_color=TEXT, font=(FONT_FAMILY_REGULAR, 10))
    ico_label.pack(pady=5)
    sizes = [16, 32, 48, 64, 128, 256]
    check_vars = [ctk.BooleanVar(value=True) for _ in sizes]  # Default checked
    for i, s in enumerate(sizes):
        cb = ctk.CTkCheckBox(ico_frame, text=f"{s}x{s}", variable=check_vars[i])
        cb.pack(anchor="w")

    def update_ico_frame(event=None):
        if target_var.get() == "ICO":
            ico_frame.pack(pady=5)
        else:
            ico_frame.pack_forget()

    combo.bind("<<ComboboxSelected>>", update_ico_frame)

    progress_bar = ctk.CTkProgressBar(img_win, width=250, mode="indeterminate")
    # Initially not packed

    def do_convert():
        fp = file_path_var.get()
        if not fp:
            messagebox.showerror("Error", "No file selected")
            return
        target = target_var.get().upper()
        input_ext = os.path.splitext(fp)[1].lower()[1:]
        if input_ext == "jpeg":
            input_ext = "jpg"
        if target.lower() == input_ext or (target == "JPG" and input_ext in ["jpg", "jpeg"]):
            messagebox.showwarning("Warning", "Input and output formats are the same")
            return
        new_file_path = os.path.splitext(fp)[0] + '.' + target.lower()

        def conversion_thread():
            try:
                img = Image.open(fp)
                if target == "JPG":
                    if img.mode in ('RGBA', 'LA', 'P'):
                        img = img.convert('RGB')
                    img.save(new_file_path, 'JPEG')
                elif target == "PNG":
                    img.save(new_file_path, 'PNG')
                elif target == "WEBP":
                    img.save(new_file_path, 'WEBP')
                elif target == "AVIF":
                    img.save(new_file_path, 'AVIF')
                elif target == "ICO":
                    selected_sizes = [sizes[i] for i, v in enumerate(check_vars) if v.get()]
                    if not selected_sizes:
                        img_win.after(0, lambda: messagebox.showerror("Error", "Select at least one size for ICO"))
                        return
                    icon_images = [img.resize((s, s), Image.LANCZOS) for s in selected_sizes]
                    icon_images[0].save(new_file_path, format='ICO', append_images=icon_images[1:] if len(icon_images) > 1 else [])
                elif target == "BMP":
                    img.save(new_file_path, 'BMP')
                elif target == "GIF":
                    img.save(new_file_path, 'GIF')
                elif target == "TIFF":
                    img.save(new_file_path, 'TIFF')
                else:
                    raise ValueError("Unsupported target format")
                img_win.after(0, lambda: messagebox.showinfo("Success", f"File converted to: {new_file_path}"))
            except Exception as e:
                img_win.after(0, lambda: messagebox.showerror("Error", f"Conversion failed: {str(e)}"))
            finally:
                img_win.after(0, progress_bar.stop)
                img_win.after(0, progress_bar.pack_forget)
                img_win.after(0, lambda: btn_convert.configure(state="normal"))

        progress_bar.pack(pady=5)
        progress_bar.start()
        btn_convert.configure(state="disabled")
        thread = threading.Thread(target=conversion_thread)
        thread.start()

    btn_convert = ctk.CTkButton(img_win, text="Convert", command=do_convert, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=250, font=(FONT_FAMILY_SEMIBOLD, 10))
    btn_convert.pack(pady=5)

    # Initially hide ico_frame if not ICO
    update_ico_frame()

def open_archive_window(master):
    arch_win = ctk.CTkToplevel(master)
    arch_win.title("Archive Conversions")
    arch_win.geometry("300x300")
    arch_win.configure(fg_color=BG)
    # Center the window
    arch_win.update_idletasks()
    screen_width = arch_win.winfo_screenwidth()
    screen_height = arch_win.winfo_screenheight()
    x = (screen_width // 2) - (300 // 2)
    y = (screen_height // 2) - (300 // 2)
    arch_win.geometry(f"300x300+{x}+{y}")
    # Set icon
    if os.path.exists(APP_ICON_PATH):
        try:
            arch_win.after(250, lambda: arch_win.iconbitmap(APP_ICON_PATH))
        except Exception as e:
            print(f"Could not set icon for archive window: {e}")
    # Make it transient and grab set to stay on top
    arch_win.transient(master)
    arch_win.grab_set()

    label = ctk.CTkLabel(arch_win, text="Archive Converter", fg_color=BG, text_color=TEXT, font=(FONT_FAMILY_REGULAR, 12))
    label.pack(pady=10)

    file_path_var = ctk.StringVar(value="")

    def select_file():
        fp = filedialog.askopenfilename(title="Select Archive File", filetypes=[("Archive files", "*.zip;*.7z;*.tar;*.tar.gz;*.tgz;*.tar.bz2;*.tbz2")])
        if fp:
            file_path_var.set(fp)
            file_label.configure(text=os.path.basename(fp))

    btn_select = ctk.CTkButton(arch_win, text="Select File", command=select_file, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=250, font=(FONT_FAMILY_SEMIBOLD, 10))
    btn_select.pack(pady=5)

    file_label = ctk.CTkLabel(arch_win, text="No file selected", fg_color=BG, text_color=TEXT, font=(FONT_FAMILY_REGULAR, 10))
    file_label.pack(pady=5)

    target_var = ctk.StringVar(value="ZIP")
    combo = ctk.CTkComboBox(arch_win, values=["ZIP", "7Z", "TAR", "TGZ", "TBZ2"], variable=target_var, font=(FONT_FAMILY_REGULAR, 10), width=250)
    combo.pack(pady=5)

    progress_bar = ctk.CTkProgressBar(arch_win, width=250, mode="indeterminate")
    # Initially not packed

    def get_archive_type(fp):
        lfp = fp.lower()
        if lfp.endswith('.zip'):
            return 'zip'
        if lfp.endswith('.7z'):
            return '7z'
        if lfp.endswith('.tar'):
            return 'tar'
        if lfp.endswith('.tar.gz') or lfp.endswith('.tgz'):
            return 'tgz'
        if lfp.endswith('.tar.bz2') or lfp.endswith('.tbz2'):
            return 'tbz2'
        raise ValueError("Unsupported archive type")

    def do_convert():
        fp = file_path_var.get()
        if not fp:
            messagebox.showerror("Error", "No file selected")
            return
        target = target_var.get().upper()
        try:
            input_type = get_archive_type(fp)
        except ValueError:
            messagebox.showerror("Error", "Unsupported input format")
            return
        if target.lower() == input_type:
            messagebox.showwarning("Warning", "Input and output formats are the same")
            return
        ext_map = {
            'ZIP': '.zip',
            '7Z': '.7z',
            'TAR': '.tar',
            'TGZ': '.tar.gz',
            'TBZ2': '.tar.bz2'
        }
        new_file_path = os.path.splitext(fp)[0] + ext_map[target]
        temp_dir = tempfile.mkdtemp()

        def conversion_thread():
            try:
                # Extract
                if input_type == 'zip':
                    with zipfile.ZipFile(fp, 'r') as z:
                        z.extractall(temp_dir)
                elif input_type == '7z':
                    with py7zr.SevenZipFile(fp, 'r') as z:
                        z.extractall(temp_dir)
                elif input_type == 'tar':
                    with tarfile.open(fp, 'r') as t:
                        t.extractall(temp_dir)
                elif input_type == 'tgz':
                    with tarfile.open(fp, 'r:gz') as t:
                        t.extractall(temp_dir)
                elif input_type == 'tbz2':
                    with tarfile.open(fp, 'r:bz2') as t:
                        t.extractall(temp_dir)
                
                # Create new archive
                if target == 'ZIP':
                    with zipfile.ZipFile(new_file_path, 'w', zipfile.ZIP_DEFLATED) as z:
                        for root, dirs, files in os.walk(temp_dir):
                            for file in files:
                                z.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), temp_dir))
                elif target == '7Z':
                    with py7zr.SevenZipFile(new_file_path, 'w') as z:
                        for root, dirs, files in os.walk(temp_dir):
                            for file in files:
                                z.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), temp_dir))
                elif target == 'TAR':
                    with tarfile.open(new_file_path, 'w') as t:
                        for root, dirs, files in os.walk(temp_dir):
                            for file in files:
                                t.add(os.path.join(root, file), os.path.relpath(os.path.join(root, file), temp_dir))
                elif target == 'TGZ':
                    with tarfile.open(new_file_path, 'w:gz') as t:
                        for root, dirs, files in os.walk(temp_dir):
                            for file in files:
                                t.add(os.path.join(root, file), os.path.relpath(os.path.join(root, file), temp_dir))
                elif target == 'TBZ2':
                    with tarfile.open(new_file_path, 'w:bz2') as t:
                        for root, dirs, files in os.walk(temp_dir):
                            for file in files:
                                t.add(os.path.join(root, file), os.path.relpath(os.path.join(root, file), temp_dir))
                arch_win.after(0, lambda: messagebox.showinfo("Success", f"File converted to: {new_file_path}"))
            except Exception as e:
                arch_win.after(0, lambda: messagebox.showerror("Error", f"Conversion failed: {str(e)}"))
            finally:
                shutil.rmtree(temp_dir)
                arch_win.after(0, progress_bar.stop)
                arch_win.after(0, progress_bar.pack_forget)
                arch_win.after(0, lambda: btn_convert.configure(state="normal"))

        progress_bar.pack(pady=5)
        progress_bar.start()
        btn_convert.configure(state="disabled")
        thread = threading.Thread(target=conversion_thread)
        thread.start()

    btn_convert = ctk.CTkButton(arch_win, text="Convert", command=do_convert, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=250, font=(FONT_FAMILY_SEMIBOLD, 10))
    btn_convert.pack(pady=5)

def open_spreadsheets_window(master):
    spreadsheets_win = ctk.CTkToplevel(master)
    spreadsheets_win.title("Spreadsheets Conversions")
    spreadsheets_win.geometry("300x300")
    spreadsheets_win.configure(fg_color=BG)
    # Center the window
    spreadsheets_win.update_idletasks()
    screen_width = spreadsheets_win.winfo_screenwidth()
    screen_height = spreadsheets_win.winfo_screenheight()
    x = (screen_width // 2) - (300 // 2)
    y = (screen_height // 2) - (300 // 2)
    spreadsheets_win.geometry(f"300x300+{x}+{y}")
    # Set icon
    if os.path.exists(APP_ICON_PATH):
        try:
            spreadsheets_win.after(250, lambda: spreadsheets_win.iconbitmap(APP_ICON_PATH))
        except Exception as e:
            print(f"Could not set icon for spreadsheets window: {e}")
    # Make it transient and grab set to stay on top
    spreadsheets_win.transient(master)
    spreadsheets_win.grab_set()

    label = ctk.CTkLabel(spreadsheets_win, text="Spreadsheets Converter", fg_color=BG, text_color=TEXT, font=(FONT_FAMILY_REGULAR, 12))
    label.pack(pady=10)

    file_path_var = ctk.StringVar(value="")

    def select_file():
        fp = filedialog.askopenfilename(title="Select Spreadsheet File", filetypes=[("Spreadsheet files", "*.xlsx;*.csv;*.ods")])
        if fp:
            file_path_var.set(fp)
            file_label.configure(text=os.path.basename(fp))

    btn_select = ctk.CTkButton(spreadsheets_win, text="Select File", command=select_file, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=250, font=(FONT_FAMILY_SEMIBOLD, 10))
    btn_select.pack(pady=5)

    file_label = ctk.CTkLabel(spreadsheets_win, text="No file selected", fg_color=BG, text_color=TEXT, font=(FONT_FAMILY_REGULAR, 10))
    file_label.pack(pady=5)

    target_var = ctk.StringVar(value="XLSX")
    combo = ctk.CTkComboBox(spreadsheets_win, values=["XLSX", "CSV", "ODS"], variable=target_var, font=(FONT_FAMILY_REGULAR, 10), width=250)
    combo.pack(pady=5)

    progress_bar = ctk.CTkProgressBar(spreadsheets_win, width=250, mode="indeterminate")
    # Initially not packed

    def do_convert():
        fp = file_path_var.get()
        if not fp:
            messagebox.showerror("Error", "No file selected")
            return
        target = target_var.get().upper()
        input_ext = os.path.splitext(fp)[1].lower()[1:]
        if target.lower() == input_ext:
            messagebox.showwarning("Warning", "Input and output formats are the same")
            return
        new_file_path = os.path.splitext(fp)[0] + '.' + target.lower()

        def conversion_thread():
            try:
                data = []
                if input_ext == "xlsx":
                    wb = openpyxl.load_workbook(fp)
                    sheet = wb.active
                    data = [[cell.value for cell in row] for row in sheet.rows]
                elif input_ext == "csv":
                    with open(fp, 'r', newline='') as f:
                        reader = csv.reader(f)
                        data = list(reader)
                elif input_ext == "ods":
                    doc = ezodf.opendoc(fp)
                    sheet = doc.sheets[0]
                    data = [[cell.value for cell in row] for row in sheet.rows()]
                else:
                    spreadsheets_win.after(0, lambda: messagebox.showerror("Error", "Unsupported input format"))
                    return

                if target == "XLSX":
                    wb = openpyxl.Workbook()
                    sheet = wb.active
                    for row in data:
                        sheet.append(row)
                    wb.save(new_file_path)
                elif target == "CSV":
                    with open(new_file_path, 'w', newline='') as f:
                        writer = csv.writer(f)
                        writer.writerows(data)
                elif target == "ODS":
                    doc = ezodf.newdoc(doctype='ods', filename=new_file_path)
                    max_cols = max(len(row) for row in data) if data else 1
                    sht = ezodf.Sheet('Sheet1', size=(len(data), max_cols))
                    doc.sheets.append(sht)
                    for r, row in enumerate(data):
                        for c, val in enumerate(row):
                            sht[r, c].set_value(val)
                    doc.save()
                else:
                    spreadsheets_win.after(0, lambda: messagebox.showerror("Error", "Unsupported target format"))
                    return
                spreadsheets_win.after(0, lambda: messagebox.showinfo("Success", f"File converted to: {new_file_path}"))
            except Exception as e:
                spreadsheets_win.after(0, lambda: messagebox.showerror("Error", f"Conversion failed: {str(e)}"))
            finally:
                spreadsheets_win.after(0, progress_bar.stop)
                spreadsheets_win.after(0, progress_bar.pack_forget)
                spreadsheets_win.after(0, lambda: btn_convert.configure(state="normal"))

        progress_bar.pack(pady=5)
        progress_bar.start()
        btn_convert.configure(state="disabled")
        thread = threading.Thread(target=conversion_thread)
        thread.start()

    btn_convert = ctk.CTkButton(spreadsheets_win, text="Convert", command=do_convert, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=250, font=(FONT_FAMILY_SEMIBOLD, 10))
    btn_convert.pack(pady=5)

# Extend the app class with open methods
class WormholeApp(WormholeApp):
    def open_docs_window(self):
        open_docs_window(self)

    def open_presentations_window(self):
        open_presentations_window(self)

    def open_images_window(self):
        open_images_window(self)

    def open_archive_window(self):
        open_archive_window(self)

    def open_spreadsheets_window(self):
        open_spreadsheets_window(self)

if __name__ == "__main__":
    app = WormholeApp()
    app.mainloop()