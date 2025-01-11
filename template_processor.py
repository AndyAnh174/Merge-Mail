import pandas as pd
from docx import Document
from jinja2 import Template
import logging
import os
from io import BytesIO
from bs4 import BeautifulSoup

class TemplateProcessor:
    def read_csv(self, csv_path):
        """Đọc file CSV và trả về DataFrame"""
        try:
            df = pd.read_csv(csv_path)
            logging.info(f"Đọc được {len(df)} dòng từ file CSV")
            logging.info(f"Các cột trong CSV: {df.columns.tolist()}")
            return df
        except Exception as e:
            logging.error(f"Lỗi đọc file CSV: {str(e)}")
            raise Exception(f"Error reading CSV file: {str(e)}")

    def _get_paragraph_style(self, paragraph):
        """Chuyển đổi định dạng đoạn văn sang HTML"""
        html = ""
        for run in paragraph.runs:
            text = run.text
            if text.strip():
                style = []
                if run.bold:
                    style.append("font-weight: bold")
                if run.italic:
                    style.append("font-style: italic")
                if run.underline:
                    style.append("text-decoration: underline")
                
                # Xử lý màu sắc
                if hasattr(run._element.rPr, 'color') and run._element.rPr.color is not None:
                    # Lấy mã màu trực tiếp từ XML
                    color_elem = run._element.rPr.color
                    if color_elem.val:
                        color = color_elem.val
                        # Nếu là auto thì dùng màu đen
                        if color == "auto":
                            style.append("color: #000000")
                        else:
                            style.append(f"color: #{color}")

                if hasattr(run.font, 'size') and run.font.size:
                    size = run.font.size.pt
                    style.append(f"font-size: {size}pt")
                
                if run.font.name:
                    style.append(f"font-family: '{run.font.name}'")
                
                if style:
                    # Thêm !important để đảm bảo styles được áp dụng
                    style = [f"{s} !important" for s in style]
                    html += f'<span style="{"; ".join(style)}">{text}</span>'
                else:
                    html += text
            else:
                html += text
        return html

    def read_template(self, template_path):
        """Đọc file template Word và trả về nội dung HTML và ảnh"""
        try:
            doc = Document(template_path)
            content = []
            images = []
            
            # Thêm CSS cơ bản
            content.append("""
                <style>
                    .email-content {
                        font-family: Arial, sans-serif !important;
                        line-height: 1.6 !important;
                        width: 100% !important;
                        max-width: 800px !important;
                        margin: 0 auto !important;
                    }
                    .paragraph {
                        margin: 10px 0 !important;
                        padding: 0 !important;
                        width: 100% !important;
                        display: block !important;
                    }
                </style>
            """)
            
            content.append('<div class="email-content">')
            
            # Chuyển đổi từng đoạn văn sang HTML
            for paragraph in doc.paragraphs:
                if paragraph.text.strip():
                    html = self._get_paragraph_style(paragraph)
                    
                    # Xử lý căn lề
                    try:
                        # Lấy style trực tiếp từ paragraph
                        style = []
                        
                        # Xử lý căn lề
                        if paragraph.alignment is not None:
                            # Mapping từ WD_ALIGN_PARAGRAPH sang CSS
                            alignment_map = {
                                0: 'left',     # WD_ALIGN_PARAGRAPH.LEFT
                                1: 'center',   # WD_ALIGN_PARAGRAPH.CENTER
                                2: 'right',    # WD_ALIGN_PARAGRAPH.RIGHT
                                3: 'justify'   # WD_ALIGN_PARAGRAPH.JUSTIFY
                            }
                            # Lấy giá trị số của alignment
                            alignment_value = int(paragraph.alignment)
                            text_align = alignment_map.get(alignment_value, 'left')
                            style.append(f"text-align: {text_align}")

                        # Thêm các style khác nếu cần
                        if style:
                            style_str = "; ".join(f"{s} !important" for s in style)
                            html = f'<div class="paragraph" style="{style_str}">{html}</div>'
                        else:
                            html = f'<div class="paragraph">{html}</div>'
                            
                    except Exception as align_err:
                        logging.warning(f"Lỗi xử lý căn lề: {str(align_err)}")
                        # Nếu có lỗi, sử dụng div mặc định
                        html = f'<div class="paragraph">{html}</div>'
                    
                    content.append(html)

            content.append("</div>")

            # Xử lý ảnh (giữ nguyên phần code cũ)
            for rel in doc.part.rels.values():
                if "image" in rel.reltype:
                    try:
                        image_part = rel.target_part
                        image_stream = BytesIO(image_part.blob)
                        ext = image_part.content_type.split('/')[-1]
                        if '.' in rel.target_ref:
                            filename = os.path.basename(rel.target_ref)
                        else:
                            filename = f'image_{len(images)}.{ext}'
                        
                        images.append({
                            'stream': image_stream,
                            'filename': filename,
                            'content_type': image_part.content_type
                        })
                        logging.info(f"Đã đọc ảnh: {filename}")
                    except Exception as img_err:
                        logging.warning(f"Không thể đọc ảnh {rel.target_ref}: {str(img_err)}")

            template_data = {
                'content': '\n'.join(content),
                'images': images,
                'is_html': True
            }
            
            logging.info(f"Đã đọc template: {len(content)} đoạn văn, {len(images)} ảnh")
            return template_data

        except Exception as e:
            logging.error(f"Lỗi đọc file template: {str(e)}")
            raise Exception(f"Error reading template file: {str(e)}")

    def merge_template(self, template_data, data):
        """Merge dữ liệu vào template"""
        try:
            template = Template(template_data['content'])
            merged_content = template.render(**data.to_dict())
            
            return {
                'content': merged_content,
                'images': template_data['images'],
                'is_html': template_data.get('is_html', False)
            }
        except Exception as e:
            logging.error(f"Lỗi merge template: {str(e)}")
            raise Exception(f"Error merging template: {str(e)}") 