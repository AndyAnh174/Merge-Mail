import re
import logging

class RuleManager:
    def __init__(self):
        logging.basicConfig(
            filename='error_log.txt',
            level=logging.DEBUG,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )

    def check_rules(self, row):
        """Kiểm tra các điều kiện của dữ liệu"""
        try:
            # Kiểm tra email có tồn tại không
            if 'email' not in row or not row['email']:
                logging.warning(f"Không tìm thấy email trong dữ liệu: {row}")
                return False

            # Kiểm tra email có đúng định dạng không
            email = str(row['email']).strip()
            
            # Pattern mới chấp nhận subdomain và domain dài hơn
            # Ví dụ: user@student.hcmute.edu.vn
            email_pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+(\.[a-zA-Z]{2,})+$'
            
            if not re.match(email_pattern, email):
                logging.warning(f"Email không đúng định dạng: {email}")
                return False

            logging.info(f"Email hợp lệ: {email}")
            return True

        except Exception as e:
            logging.error(f"Lỗi kiểm tra rules: {str(e)}")
            return False 