import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from email.mime.application import MIMEApplication
import time
import logging
import os
from bs4 import BeautifulSoup

class EmailSender:
    def __init__(self):
        self.email = None
        self.password = None
        self.max_retries = 3
        self.attachments = []
        
        logging.basicConfig(
            filename='error_log.txt',
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )

    def configure(self, email, password):
        """Cấu hình thông tin đăng nhập email"""
        if not email or not password:
            raise ValueError("Email và App Password không được để trống")
        
        if not email.endswith('@gmail.com'):
            raise ValueError("Chỉ hỗ trợ địa chỉ Gmail")
            
        if len(password) != 16:
            raise ValueError("App Password phải có 16 ký tự")
            
        self.email = email
        self.password = password
        logging.info(f"Đã cấu hình email: {email}")

    def add_attachment(self, filepath):
        """Thêm file đính kèm"""
        if os.path.exists(filepath):
            self.attachments.append(filepath)
            logging.info(f"Đã thêm file đính kèm: {filepath}")
        else:
            raise ValueError(f"Không tìm thấy file: {filepath}")

    def clear_attachments(self):
        """Xóa tất cả file đính kèm"""
        self.attachments = []
        logging.info("Đã xóa tất cả file đính kèm")

    def send_email(self, to_email, subject, content):
        """Gửi email với định dạng HTML và file đính kèm"""
        if not self.email or not self.password:
            raise Exception("Chưa cấu hình thông tin email")

        for attempt in range(self.max_retries):
            try:
                logging.info(f"Đang gửi email đến {to_email} (lần thử {attempt + 1})")
                
                # Tạo message với phần related để nhúng ảnh
                msg = MIMEMultipart('related')
                msg['From'] = self.email
                msg['To'] = to_email
                msg['Subject'] = subject

                # Tạo phần alternative cho text và html
                msgAlternative = MIMEMultipart('alternative')
                msg.attach(msgAlternative)

                # Xử lý nội dung email
                if isinstance(content, dict):
                    email_content = content.get('content', '')
                    is_html = content.get('is_html', False)
                    
                    if is_html:
                        # Tạo phiên bản text từ HTML
                        soup = BeautifulSoup(email_content, 'html.parser')
                        text_content = soup.get_text()
                        
                        # Thêm cả text và html version
                        msgAlternative.attach(MIMEText(text_content, 'plain', 'utf-8'))
                        msgAlternative.attach(MIMEText(email_content, 'html', 'utf-8'))
                    else:
                        msgAlternative.attach(MIMEText(email_content, 'plain', 'utf-8'))

                    # Xử lý ảnh từ template
                    if 'images' in content and content['images']:
                        for img in content['images']:
                            try:
                                image = MIMEImage(img['stream'].getvalue())
                                image.add_header('Content-ID', f'<{img["filename"]}>')
                                msg.attach(image)
                                logging.info(f"Đã đính kèm ảnh từ template: {img['filename']}")
                            except Exception as img_err:
                                logging.error(f"Lỗi đính kèm ảnh từ template: {str(img_err)}")
                else:
                    msgAlternative.attach(MIMEText(str(content), 'plain', 'utf-8'))

                # Xử lý file đính kèm thêm
                for filepath in self.attachments:
                    try:
                        filename = os.path.basename(filepath)
                        with open(filepath, 'rb') as f:
                            if filename.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp')):
                                attachment = MIMEImage(f.read())
                            else:
                                attachment = MIMEApplication(f.read())
                            attachment.add_header('Content-Disposition', 'attachment', filename=filename)
                            msg.attach(attachment)
                            logging.info(f"Đã đính kèm file: {filename}")
                    except Exception as att_err:
                        logging.error(f"Lỗi đính kèm file {filepath}: {str(att_err)}")

                # Kết nối và gửi email
                logging.info("Bắt đầu kết nối SMTP...")
                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.set_debuglevel(1)
                
                logging.info("Thiết lập TLS...")
                server.starttls()
                
                logging.info("Đang đăng nhập...")
                server.login(self.email, self.password)
                
                logging.info("Đang gửi email...")
                server.send_message(msg)
                
                server.quit()
                logging.info(f"Đã gửi email thành công đến {to_email}")
                return True

            except smtplib.SMTPAuthenticationError as auth_err:
                logging.error(f"Lỗi xác thực Gmail: {str(auth_err)}")
                raise Exception("Sai email hoặc App Password. Vui lòng kiểm tra lại.")
            
            except smtplib.SMTPException as smtp_err:
                logging.error(f"Lỗi SMTP: {str(smtp_err)}")
                if attempt < self.max_retries - 1:
                    logging.info(f"Đợi 2 giây trước khi thử lại...")
                    time.sleep(2)
                    continue
                raise Exception(f"Lỗi gửi email: {str(smtp_err)}")
            
            except Exception as e:
                logging.error(f"Lỗi không xác định: {str(e)}")
                if attempt < self.max_retries - 1:
                    logging.info(f"Đợi 2 giây trước khi thử lại...")
                    time.sleep(2)
                    continue
                raise Exception(f"Lỗi không xác định: {str(e)}")

        raise Exception(f"Không thể gửi email sau {self.max_retries} lần thử") 