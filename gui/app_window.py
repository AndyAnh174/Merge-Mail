import tkinter as tk
from tkinter import ttk
from tkinter.font import Font
from .frames.file_frame import FileSelectionFrame
from .frames.email_frame import EmailConfigFrame
from .frames.progress_frame import ProgressFrame
from .frames.button_frame import ButtonFrame
import webbrowser

class AppWindow:
    def __init__(self, controller):
        self.window = tk.Tk()
        self.window.title("Merge Mail Application")
        self.window.geometry("800x600")
        self.window.configure(bg='#f0f0f0')
        
        self.controller = controller
        self.setup_fonts()
        self.setup_styles()
        self.setup_ui()

    def setup_fonts(self):
        self.header_font = Font(family="Helvetica", size=12, weight="bold")
        self.normal_font = Font(family="Helvetica", size=10)
        self.copyright_font = Font(family="Helvetica", size=9)

    def setup_styles(self):
        style = ttk.Style()
        style.configure(
            'Accent.TButton',
            background='#2980b9',
            foreground='black',
            padding=10
        )
        style.configure(
            'Info.TButton',
            background='#27ae60',
            foreground='black',
            padding=10
        )
        style.configure(
            'TLabel',
            foreground='black'
        )
        style.configure(
            'TLabelframe.Label',
            foreground='black'
        )
        style.configure(
            'TEntry',
            foreground='black'
        )

    def setup_ui(self):
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(
            main_frame, 
            text="Email Merge System of HCMUTE - Developer Student Club",
            font=("Helvetica", 16, "bold"),
            foreground="#2c3e50"
        )
        title_label.pack(pady=(0, 20))

        # File Selection Frame
        self.file_frame = FileSelectionFrame(main_frame, self.controller)
        
        # Email Configuration Frame
        self.email_frame = EmailConfigFrame(main_frame, self.controller)
        
        # Progress Frame
        self.progress_frame = ProgressFrame(main_frame)
        
        # Button Frame
        self.button_frame = ButtonFrame(main_frame, self.controller)

        # Copyright Frame
        copyright_frame = ttk.Frame(main_frame)
        copyright_frame.pack(fill=tk.X, pady=(20, 0))

        copyright_text = "This product is created by "
        andy_link = "AndyAnh"
        copyright_text2 = " and sponsored by "
        dsc_link = "HCMUTE - Developer Student Club"

        # Create copyright label with clickable links
        copyright_label = ttk.Label(
            copyright_frame,
            text=copyright_text,
            font=self.copyright_font
        )
        copyright_label.pack(side=tk.LEFT)

        andy_label = ttk.Label(
            copyright_frame,
            text=andy_link,
            font=self.copyright_font,
            foreground='blue',
            cursor='hand2'
        )
        andy_label.pack(side=tk.LEFT)
        andy_label.bind('<Button-1>', lambda e: webbrowser.open('https://github.com/AndyAnh174'))

        mid_label = ttk.Label(
            copyright_frame,
            text=copyright_text2,
            font=self.copyright_font
        )
        mid_label.pack(side=tk.LEFT)

        dsc_label = ttk.Label(
            copyright_frame,
            text=dsc_link,
            font=self.copyright_font,
            foreground='blue',
            cursor='hand2'
        )
        dsc_label.pack(side=tk.LEFT)
        dsc_label.bind('<Button-1>', lambda e: webbrowser.open('https://hcmute-dsc.vercel.app/'))

    def run(self):
        # Center window
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f'{width}x{height}+{x}+{y}')
        
        self.window.mainloop() 

    def get_email_frame(self):
        return self.email_frame 