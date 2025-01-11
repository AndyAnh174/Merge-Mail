import tkinter as tk
from tkinter import ttk

class ButtonFrame:
    def __init__(self, parent, controller):
        self.frame = ttk.Frame(parent)
        self.frame.pack(fill=tk.X, pady=10)
        
        self.controller = controller
        self.setup_ui()

    def setup_ui(self):
        # Start Button
        self.start_button = ttk.Button(
            self.frame,
            text="Start Merge and Send",
            command=self.controller.start_merge_mail,
            style='Accent.TButton'
        )
        self.start_button.pack(side=tk.LEFT, padx=10)
        
        # Guide Button
        self.guide_button = ttk.Button(
            self.frame,
            text="Hướng dẫn sử dụng",
            command=self.controller.show_guide,
            style='Info.TButton'
        )
        self.guide_button.pack(side=tk.LEFT, padx=10)
        
        # Sample Files Button
        self.sample_button = ttk.Button(
            self.frame,
            text="Tạo file mẫu",
            command=self.controller.create_sample_files,
            style='Info.TButton'
        )
        self.sample_button.pack(side=tk.LEFT, padx=10) 