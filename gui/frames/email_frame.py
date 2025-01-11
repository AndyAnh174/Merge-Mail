import tkinter as tk
from tkinter import ttk

class EmailConfigFrame:
    def __init__(self, parent, controller):
        self.frame = ttk.LabelFrame(
            parent,
            text="Email Configuration",
            padding="10"
        )
        self.frame.pack(fill=tk.X, pady=10)
        
        self.controller = controller
        self.setup_ui()

    def setup_ui(self):
        # Gmail
        gmail_container = ttk.Frame(self.frame)
        gmail_container.pack(fill=tk.X, pady=5)
        
        ttk.Label(gmail_container, text="Gmail:").pack(side=tk.LEFT)
        self.email_entry = ttk.Entry(gmail_container, width=40)
        self.email_entry.pack(side=tk.LEFT, padx=5)

        # App Password
        password_container = ttk.Frame(self.frame)
        password_container.pack(fill=tk.X, pady=5)
        
        ttk.Label(password_container, text="App Password:").pack(side=tk.LEFT)
        self.password_entry = ttk.Entry(password_container, show="*", width=40)
        self.password_entry.pack(side=tk.LEFT, padx=5)

        # Subject (Tiêu đề)
        subject_container = ttk.Frame(self.frame)
        subject_container.pack(fill=tk.X, pady=5)
        
        ttk.Label(subject_container, text="Tiêu đề:").pack(side=tk.LEFT)
        self.subject_entry = ttk.Entry(subject_container, width=40)
        self.subject_entry.pack(side=tk.LEFT, padx=5)
        self.subject_entry.insert(0, "Thư gửi {{ name }}")  # Giá trị mặc định

        # Help Text
        help_text = "Để lấy App Password: Google Account > Security > 2-Step Verification > App Passwords"
        help_label = ttk.Label(
            self.frame,
            text=help_text,
            foreground='#0000EE',
            cursor="hand2",
            wraplength=600,
            font=('Helvetica', 9, 'underline')
        )
        help_label.pack(pady=5)
        help_label.bind("<Button-1>", lambda e: self.controller.open_help())

        # Attachment Frame
        attachment_container = ttk.Frame(self.frame)
        attachment_container.pack(fill=tk.X, pady=5)
        
        ttk.Label(attachment_container, text="File đính kèm:").pack(side=tk.LEFT)
        self.attachment_label = ttk.Label(
            attachment_container,
            text="Chưa có file nào được chọn",
            wraplength=300
        )
        self.attachment_label.pack(side=tk.LEFT, padx=5)
        
        self.browse_btn = ttk.Button(
            attachment_container,
            text="Chọn file",
            command=self.controller.browse_attachment,
            style='Info.TButton'
        )
        self.browse_btn.pack(side=tk.LEFT, padx=5)
        
        self.clear_btn = ttk.Button(
            attachment_container,
            text="Xóa file",
            command=self.controller.clear_attachments,
            style='Info.TButton'
        )
        self.clear_btn.pack(side=tk.LEFT)

    def get_email(self):
        return self.email_entry.get()

    def get_password(self):
        return self.password_entry.get()

    def get_subject(self):
        return self.subject_entry.get() 

    def update_attachment_label(self, text):
        """Cập nhật label hiển thị tên file đính kèm"""
        self.attachment_label.config(text=text) 