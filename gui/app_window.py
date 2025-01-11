import tkinter as tk
from tkinter import ttk
from tkinter.font import Font
from .frames.file_frame import FileSelectionFrame
from .frames.email_frame import EmailConfigFrame
from .frames.progress_frame import ProgressFrame
from .frames.button_frame import ButtonFrame

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