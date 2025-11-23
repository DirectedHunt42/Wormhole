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

        btn_other = ctk.CTkButton(self, text="Other", command=self.open_other_window, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=300, font=(FONT_FAMILY_SEMIBOLD, 20))
        btn_other.pack(pady=5)

# Functions to open subwindows for each category

def open_docs_window(master):
    docs_win = ctk.CTkToplevel(master)
    docs_win.title("Docs Conversions")
    docs_win.geometry("300x250")
    docs_win.configure(fg_color=BG)
    # Center the window
    docs_win.update_idletasks()
    screen_width = docs_win.winfo_screenwidth()
    screen_height = docs_win.winfo_screenheight()
    x = (screen_width // 2) - (300 // 2)
    y = (screen_height // 2) - (250 // 2)
    docs_win.geometry(f"300x250+{x}+{y}")
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

    def select_file():
        fp = filedialog.askopenfilename(title="Select Docs File", filetypes=[("Docs files", "*.txt;*.pdf")])
        if fp:
            file_path_var.set(fp)
            file_label.configure(text=os.path.basename(fp))

    btn_select = ctk.CTkButton(docs_win, text="Select File", command=select_file, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=250, font=(FONT_FAMILY_SEMIBOLD, 10))
    btn_select.pack(pady=5)

    file_label = ctk.CTkLabel(docs_win, text="No file selected", fg_color=BG, text_color=TEXT, font=(FONT_FAMILY_REGULAR, 10))
    file_label.pack(pady=5)

    target_var = ctk.StringVar(value="PDF")
    combo = ctk.CTkComboBox(docs_win, values=["PDF", "TXT"], variable=target_var, font=(FONT_FAMILY_REGULAR, 10), width=250)
    combo.pack(pady=5)

    def do_convert():
        fp = file_path_var.get()
        if not fp:
            messagebox.showerror("Error", "No file selected")
            return
        target = target_var.get().lower()
        input_ext = os.path.splitext(fp)[1].lower()[1:]
        if target == input_ext:
            messagebox.showwarning("Warning", "Input and output formats are the same")
            return
        new_file_path = os.path.splitext(fp)[0] + '.' + target
        try:
            if input_ext == "txt" and target == "pdf":
                c = canvas.Canvas(new_file_path, pagesize=letter)
                width, height = letter
                y = height - 50  # Start from top with margin
                with open(fp, 'r') as f:
                    for line in f:
                        c.drawString(50, y, line.strip())
                        y -= 15  # Line spacing
                        if y < 50:  # Simple page break handling
                            c.showPage()
                            y = height - 50
                c.save()
            elif input_ext == "pdf" and target == "txt":
                reader = PdfReader(fp)
                text = ''
                for page in reader.pages:
                    text += page.extract_text() + '\n'
                with open(new_file_path, 'w') as f:
                    f.write(text)
            else:
                messagebox.showerror("Error", "Unsupported conversion")
                return
            messagebox.showinfo("Success", f"File converted to: {new_file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")

    btn_convert = ctk.CTkButton(docs_win, text="Convert", command=do_convert, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=250, font=(FONT_FAMILY_SEMIBOLD, 10))
    btn_convert.pack(pady=5)

def open_presentations_window(master):
    pres_win = ctk.CTkToplevel(master)
    pres_win.title("Presentations Conversions")
    pres_win.geometry("300x200")
    pres_win.configure(fg_color=BG)
    # Center the window
    pres_win.update_idletasks()
    screen_width = pres_win.winfo_screenwidth()
    screen_height = pres_win.winfo_screenheight()
    x = (screen_width // 2) - (300 // 2)
    y = (screen_height // 2) - (200 // 2)
    pres_win.geometry(f"300x200+{x}+{y}")
    # Set icon
    if os.path.exists(APP_ICON_PATH):
        try:
            pres_win.after(250, lambda: pres_win.iconbitmap(APP_ICON_PATH))
        except Exception as e:
            print(f"Could not set icon for presentations window: {e}")
    # Make it transient and grab set to stay on top
    pres_win.transient(master)
    pres_win.grab_set()

    label = ctk.CTkLabel(pres_win, text="Presentations conversions coming soon!", fg_color=BG, text_color=TEXT, font=(FONT_FAMILY_REGULAR, 12))
    label.pack(pady=10)

def open_images_window(master):
    img_win = ctk.CTkToplevel(master)
    img_win.title("Images Conversions")
    img_win.geometry("300x400")
    img_win.configure(fg_color=BG)
    # Center the window
    img_win.update_idletasks()
    screen_width = img_win.winfo_screenwidth()
    screen_height = img_win.winfo_screenheight()
    x = (screen_width // 2) - (300 // 2)
    y = (screen_height // 2) - (400 // 2)
    img_win.geometry(f"300x400+{x}+{y}")
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
        fp = filedialog.askopenfilename(title="Select Image File", filetypes=[("Image files", "*.jpg;*.jpeg;*.png;*.webp;*.avif;*.ico")])
        if fp:
            file_path_var.set(fp)
            file_label.configure(text=os.path.basename(fp))

    btn_select = ctk.CTkButton(img_win, text="Select File", command=select_file, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=250, font=(FONT_FAMILY_SEMIBOLD, 10))
    btn_select.pack(pady=5)

    file_label = ctk.CTkLabel(img_win, text="No file selected", fg_color=BG, text_color=TEXT, font=(FONT_FAMILY_REGULAR, 10))
    file_label.pack(pady=5)

    target_var = ctk.StringVar(value="PNG")
    combo = ctk.CTkComboBox(img_win, values=["PNG", "JPG", "WEBP", "AVIF", "ICO"], variable=target_var, font=(FONT_FAMILY_REGULAR, 10), width=250)
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
                    messagebox.showerror("Error", "Select at least one size for ICO")
                    return
                icon_images = [img.resize((s, s), Image.LANCZOS) for s in selected_sizes]
                icon_images[0].save(new_file_path, format='ICO', append_images=icon_images[1:] if len(icon_images) > 1 else [])
            else:
                raise ValueError("Unsupported target format")
            messagebox.showinfo("Success", f"File converted to: {new_file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")

    btn_convert = ctk.CTkButton(img_win, text="Convert", command=do_convert, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=250, font=(FONT_FAMILY_SEMIBOLD, 10))
    btn_convert.pack(pady=5)

    # Initially hide ico_frame
    ico_frame.pack_forget()

def open_archive_window(master):
    arch_win = ctk.CTkToplevel(master)
    arch_win.title("Archive Conversions")
    arch_win.geometry("300x250")
    arch_win.configure(fg_color=BG)
    # Center the window
    arch_win.update_idletasks()
    screen_width = arch_win.winfo_screenwidth()
    screen_height = arch_win.winfo_screenheight()
    x = (screen_width // 2) - (300 // 2)
    y = (screen_height // 2) - (250 // 2)
    arch_win.geometry(f"300x250+{x}+{y}")
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
        fp = filedialog.askopenfilename(title="Select Archive File", filetypes=[("Archive files", "*.zip;*.7z;*.tar")])
        if fp:
            file_path_var.set(fp)
            file_label.configure(text=os.path.basename(fp))

    btn_select = ctk.CTkButton(arch_win, text="Select File", command=select_file, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=250, font=(FONT_FAMILY_SEMIBOLD, 10))
    btn_select.pack(pady=5)

    file_label = ctk.CTkLabel(arch_win, text="No file selected", fg_color=BG, text_color=TEXT, font=(FONT_FAMILY_REGULAR, 10))
    file_label.pack(pady=5)

    target_var = ctk.StringVar(value="ZIP")
    combo = ctk.CTkComboBox(arch_win, values=["ZIP", "7Z", "TAR"], variable=target_var, font=(FONT_FAMILY_REGULAR, 10), width=250)
    combo.pack(pady=5)

    def do_convert():
        fp = file_path_var.get()
        if not fp:
            messagebox.showerror("Error", "No file selected")
            return
        target = target_var.get().lower()
        input_ext = os.path.splitext(fp)[1].lower()[1:]
        if target == input_ext:
            messagebox.showwarning("Warning", "Input and output formats are the same")
            return
        new_file_path = os.path.splitext(fp)[0] + '.' + target
        temp_dir = tempfile.mkdtemp()
        try:
            # Extract
            if input_ext == 'zip':
                with zipfile.ZipFile(fp, 'r') as z:
                    z.extractall(temp_dir)
            elif input_ext == '7z':
                with py7zr.SevenZipFile(fp, 'r') as z:
                    z.extractall(temp_dir)
            elif input_ext == 'tar':
                with tarfile.open(fp, 'r') as t:
                    t.extractall(temp_dir)
            else:
                raise ValueError("Unsupported input format")
            
            # Create new archive
            if target == 'zip':
                with zipfile.ZipFile(new_file_path, 'w', zipfile.ZIP_DEFLATED) as z:
                    for root, dirs, files in os.walk(temp_dir):
                        for file in files:
                            z.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), temp_dir))
            elif target == '7z':
                with py7zr.SevenZipFile(new_file_path, 'w') as z:
                    for root, dirs, files in os.walk(temp_dir):
                        for file in files:
                            z.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), temp_dir))
            elif target == 'tar':
                with tarfile.open(new_file_path, 'w') as t:
                    for root, dirs, files in os.walk(temp_dir):
                        for file in files:
                            t.add(os.path.join(root, file), os.path.relpath(os.path.join(root, file), temp_dir))
            else:
                raise ValueError("Unsupported target format")
            messagebox.showinfo("Success", f"File converted to: {new_file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Conversion failed: {str(e)}")
        finally:
            shutil.rmtree(temp_dir)

    btn_convert = ctk.CTkButton(arch_win, text="Convert", command=do_convert, fg_color=ACCENT, text_color=BG, hover_color=ACCENT_DIM, corner_radius=20, width=250, font=(FONT_FAMILY_SEMIBOLD, 10))
    btn_convert.pack(pady=5)

def open_other_window(master):
    other_win = ctk.CTkToplevel(master)
    other_win.title("Other Conversions")
    other_win.geometry("300x200")
    other_win.configure(fg_color=BG)
    # Center the window
    other_win.update_idletasks()
    screen_width = other_win.winfo_screenwidth()
    screen_height = other_win.winfo_screenheight()
    x = (screen_width // 2) - (300 // 2)
    y = (screen_height // 2) - (200 // 2)
    other_win.geometry(f"300x200+{x}+{y}")
    # Set icon
    if os.path.exists(APP_ICON_PATH):
        try:
            other_win.after(250, lambda: other_win.iconbitmap(APP_ICON_PATH))
        except Exception as e:
            print(f"Could not set icon for other window: {e}")
    # Make it transient and grab set to stay on top
    other_win.transient(master)
    other_win.grab_set()

    label = ctk.CTkLabel(other_win, text="Other conversions coming soon!", fg_color=BG, text_color=TEXT, font=(FONT_FAMILY_REGULAR, 12))
    label.pack(pady=10)

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

    def open_other_window(self):
        open_other_window(self)

if __name__ == "__main__":
    app = WormholeApp()
    app.mainloop()