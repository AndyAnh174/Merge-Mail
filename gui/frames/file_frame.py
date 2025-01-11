import tkinter as tk
from tkinter import ttk

class FileSelectionFrame:
    def __init__(self, parent, controller):
        self.frame = ttk.LabelFrame(
            parent,
            text="File Selection",
            padding="10"
        )
        self.frame.pack(fill=tk.X, pady=(0, 10))
        
        self.controller = controller
        self.setup_ui()

    def setup_ui(self):
        # CSV File Selection
        csv_container = ttk.Frame(self.frame)
        csv_container.pack(fill=tk.X, pady=5)
        
        ttk.Label(csv_container, text="CSV File:").pack(side=tk.LEFT)
        self.csv_entry = ttk.Entry(
            csv_container,
            textvariable=self.controller.csv_path,
            width=50,
            state='readonly'  # Chỉ đọc, không cho phép sửa trực tiếp
        )
        self.csv_entry.pack(side=tk.LEFT, padx=5)
        
        browse_csv_btn = ttk.Button(
            csv_container,
            text="Browse",
            command=self.controller.browse_csv,
            style='Accent.TButton'
        )
        browse_csv_btn.pack(side=tk.LEFT)

        # Template File Selection
        template_container = ttk.Frame(self.frame)
        template_container.pack(fill=tk.X, pady=5)
        
        ttk.Label(template_container, text="Template:").pack(side=tk.LEFT)
        self.template_entry = ttk.Entry(
            template_container,
            textvariable=self.controller.template_path,
            width=50,
            state='readonly'  # Chỉ đọc, không cho phép sửa trực tiếp
        )
        self.template_entry.pack(side=tk.LEFT, padx=5)
        
        browse_template_btn = ttk.Button(
            template_container,
            text="Browse",
            command=self.controller.browse_template,
            style='Accent.TButton'
        )
        browse_template_btn.pack(side=tk.LEFT)

    def clear_paths(self):
        """Xóa đường dẫn file đã chọn"""
        self.controller.csv_path.set("")
        self.controller.template_path.set("") 