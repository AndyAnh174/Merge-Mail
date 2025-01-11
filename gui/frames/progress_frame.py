import tkinter as tk
from tkinter import ttk

class ProgressFrame:
    def __init__(self, parent):
        self.frame = ttk.Frame(parent)
        self.frame.pack(fill=tk.X, pady=10)
        self.setup_ui()

    def setup_ui(self):
        self.progress = ttk.Progressbar(
            self.frame,
            orient=tk.HORIZONTAL,
            length=300,
            mode='determinate'
        )
        self.progress.pack(pady=5)
        
        self.status_label = ttk.Label(
            self.frame,
            text="Ready to send emails"
        )
        self.status_label.pack()

    def update_progress(self, value):
        self.progress['value'] = value
        self.frame.update()

    def update_status(self, message):
        self.status_label.config(text=message)
        self.frame.update() 