import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
import subprocess
from PIL import Image

import docx2pdf
import pdf2docx

try:
    import comtypes.client
except ImportError:
    pass

class FileConverter:
    """Standalone class utilizing classmethods to handle conversion logic securely"""
    
    @classmethod
    def convert_docx2pdf(cls, input_path, output_path):
        docx2pdf.convert(input_path, output_path)

    @classmethod
    def convert_ppt2pdf(cls, input_path, output_path):
        abs_input = os.path.abspath(input_path)
        abs_output = os.path.abspath(output_path)
        
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1
        deck = powerpoint.Presentations.Open(abs_input)
        deck.SaveAs(abs_output, 32)
        deck.Close()
        powerpoint.Quit()

    @classmethod
    def convert_jpg2png(cls, input_path, output_path):
        img = Image.open(input_path)
        img.save(output_path, "PNG")

    @classmethod
    def convert_pdf2docx(cls, input_path, output_path):
        cv = pdf2docx.Converter(input_path)
        cv.convert(output_path)
        cv.close()

    @classmethod
    def convert_pdf2ppt(cls, input_path, output_path):
        code_snippet = f'''
import sys
from pdf2pptx.cli import main
# Mock commandline arguments
sys.argv = ['pdf2pptx', r"""{input_path}""", '-o', r"""{output_path}"""]
main()
'''
        res = subprocess.run([sys.executable, "-c", code_snippet], capture_output=True, text=True)
        if res.returncode != 0:
            raise Exception(res.stderr or res.stdout or "An unknown pdf2pptx error occurred.")

    @classmethod
    def convert_png2jpg(cls, input_path, output_path):
        img = Image.open(input_path)
        if img.mode in ('RGBA', 'LA'):
            background = Image.new('RGB', img.size, (255, 255, 255))
            background.paste(img, mask=img.split()[3])
            background.save(output_path, "JPEG")
        else:
            img.convert('RGB').save(output_path, "JPEG")


class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Multi-Format Quick Converter")
        self.root.geometry("650x500")
        self.root.resizable(False, False)
        
        self.input_file = tk.StringVar()
        self.conversion_type = tk.StringVar(value="docx2pdf")
        
        self.setup_ui()

    def setup_ui(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        bg_color = "#2b2d42"
        frame_bg = "#8d99ae"
        text_color = "#edf2f4"
        btn_bg = "#ef233c"
        btn_active = "#d90429"
        
        self.root.configure(bg=bg_color)
        
        style.configure("TFrame", background=bg_color)
        style.configure("TLabel", background=bg_color, foreground=text_color, font=("Helvetica", 12, "bold"))
        style.configure("Title.TLabel", font=("Helvetica", 22, "bold"), background=bg_color, foreground="#edf2f4")
        
        style.configure("TLabelframe", background=bg_color, foreground=text_color, bordercolor=frame_bg)
        style.configure("TLabelframe.Label", font=("Helvetica", 14, "bold"), background=bg_color, foreground="#ef233c")
        
        style.configure("TRadiobutton", background=bg_color, foreground=text_color, font=("Helvetica", 12, "bold"))
        style.map("TRadiobutton", background=[('active', bg_color)])
        
        style.configure("TButton", padding=10, relief="flat", font=("Helvetica", 12, "bold"), background=btn_bg, foreground="white")
        style.map("TButton", background=[('active', btn_active)])
        
        style.configure("TProgressbar", background="#2ecc71", troughcolor=bg_color, bordercolor="#2ecc71", lightcolor="#2ecc71", darkcolor="#2ecc71")
        
        main_frame = ttk.Frame(self.root, padding="25")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        title_label = ttk.Label(main_frame, text="Universal File Converter", style="Title.TLabel")
        title_label.pack(pady=(0, 25))
        
        type_frame = ttk.LabelFrame(main_frame, text="1. Select Conversion Type", padding="15")
        type_frame.pack(fill=tk.X, pady=(0, 20))
        
        types = ["docx2pdf", "ppt2pdf", "jpg2png", "pdf2docx", "pdf2ppt", "png2jpg"]
        for i, t in enumerate(types):
            # We add `command=self.on_radio_select` to trigger the browse popup automatically!
            rb = ttk.Radiobutton(type_frame, text=t.upper(), variable=self.conversion_type, value=t, command=self.on_radio_select)
            rb.grid(row=i//3, column=i%3, sticky=tk.W, padx=25, pady=8)
            
        file_frame = ttk.LabelFrame(main_frame, text="2. Select Input File", padding="15")
        file_frame.pack(fill=tk.X, pady=(0, 25))
        
        file_btn = ttk.Button(file_frame, text="Browse...", command=self.browse_file)
        file_btn.pack(side=tk.LEFT, padx=(0, 15))
        
        self.file_label = ttk.Label(file_frame, text="Waiting for file...", foreground=frame_bg)
        self.file_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.convert_btn = ttk.Button(btn_frame, text="Start Conversion", command=self.start_conversion)
        self.convert_btn.pack(pady=10, ipadx=20)
        
        self.progress_bar = ttk.Progressbar(main_frame, mode='determinate', length=400, maximum=100)
        self.progress_bar.pack(pady=(10, 0))
        
        self.status_label = ttk.Label(main_frame, text="", font=("Helvetica", 12, "bold italic"))
        self.status_label.pack(side=tk.BOTTOM, pady=5)

    def on_radio_select(self):
        # Automatically launch the file browser when they select a format
        self.browse_file()

    def browse_file(self):
        # Dynamically set the expected file extension to filter out irrelevant files
        conv = self.conversion_type.get()
        expected_ext = conv.split("2")[0]
        
        filename = filedialog.askopenfilename(
            initialdir=os.getcwd(),  # Forces it to open where your python files are, rather than Pictures!
            title=f"Select a {expected_ext.upper()} file to convert",
            filetypes=[(f"{expected_ext.upper()} Files", f"*.{expected_ext}"), ("All Files", "*.*")]
        )
        if filename:
            self.input_file.set(filename)
            self.file_label.config(text=os.path.basename(filename), foreground="#edf2f4")
            self.status_label.config(text="")

    def start_conversion(self):
        input_path = self.input_file.get()
        if not input_path:
            messagebox.showwarning("No File Found", "Please browse and select an input file first!")
            return
            
        conv_type = self.conversion_type.get()
        
        # Determine output path next to input
        base, _ = os.path.splitext(input_path)
        ext = conv_type.split("2")[1]
        output_path = f"{base}.{ext}"
        
        if not os.path.exists(input_path):
            messagebox.showerror("File Error", "The file specified does not exist.")
            return

        self.convert_btn.state(['disabled'])
        self.status_label.config(text=f"Processing {conv_type} conversion, please wait...", foreground="#8d99ae")
        self.progress_bar['value'] = 100
        
        thread = threading.Thread(target=self.run_conversion, args=(conv_type, input_path, output_path))
        thread.daemon = True
        thread.start()

    def run_conversion(self, conv_type, input_path, output_path):
        success = False
        error_msg = ""
        try:
            # Dynamically route the call to our standalone FileConverter class
            method_name = f"convert_{conv_type}"
            if hasattr(FileConverter, method_name):
                method = getattr(FileConverter, method_name)
                method(input_path, output_path)
                success = True
            else:
                raise NotImplementedError(f"Conversion {conv_type} is not setup.")
        except Exception as e:
            error_msg = str(e)
            
        self.root.after(0, lambda: self.conversion_done(success, output_path, error_msg))

    def conversion_done(self, success, output_path, error_msg):
        self.convert_btn.state(['!disabled'])
        self.progress_bar['value'] = 0
        if success:
            self.status_label.config(text=f"Success! file saved as {os.path.basename(output_path)}", foreground="#2ecc71")
            messagebox.showinfo("Conversion Success", f"Your file was converted correctly and saved as:\n\n{output_path}")
        else:
            self.status_label.config(text="Conversion Failed", foreground="#ef233c")
            messagebox.showerror("Error", f"Failed to convert file:\n\n{error_msg}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ConverterApp(root)
    root.mainloop()
