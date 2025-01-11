import pandas as pd
from docx import Document
from jinja2 import Template
import logging
import os
from io import BytesIO
from bs4 import BeautifulSoup
from docx.shared import Pt

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
            # Xử lý đặc biệt cho các ký tự xuống dòng
            text = text.replace('\n', '<br class="line-break"/>')
            text = text.replace('\r', '<br class="line-break"/>')
            # Giữ nguyên khoảng trắng liên tiếp
            text = text.replace(' ', '&nbsp;')
            
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
                    color_elem = run._element.rPr.color
                    if color_elem.val:
                        color = color_elem.val
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
                    style = [f"{s} !important" for s in style]
                    html += f'<span style="{"; ".join(style)}">{text}</span>'
                else:
                    html += text
            else:
                html += text
        return html

    def _get_paragraph_spacing(self, paragraph):
        """Lấy thông tin khoảng cách đoạn văn"""
        try:
            style = []
            
            # Khoảng cách trước đoạn văn
            if paragraph.paragraph_format.space_before:
                space_before = paragraph.paragraph_format.space_before.pt
                style.append(f"margin-top: {space_before}pt")
            else:
                style.append("margin-top: 0pt")
            
            # Khoảng cách sau đoạn văn
            if paragraph.paragraph_format.space_after:
                space_after = paragraph.paragraph_format.space_after.pt
                style.append(f"margin-bottom: {space_after}pt")
            else:
                style.append("margin-bottom: 0pt")
            
            # Khoảng cách giữa các dòng
            if paragraph.paragraph_format.line_spacing:
                # Xử lý các loại line spacing khác nhau
                if paragraph.paragraph_format.line_spacing_rule == 1:  # Single
                    style.append("line-height: 1.15")  # Tăng lên 1.15 để dễ đọc hơn
                elif paragraph.paragraph_format.line_spacing_rule == 2:  # 1.5 lines
                    style.append("line-height: 1.5")
                elif paragraph.paragraph_format.line_spacing_rule == 3:  # Double
                    style.append("line-height: 2")
                elif paragraph.paragraph_format.line_spacing_rule == 4:  # At least
                    min_height = paragraph.paragraph_format.line_spacing.pt
                    style.append(f"line-height: {min_height}pt")
                    style.append("min-height: {min_height}pt")
                elif paragraph.paragraph_format.line_spacing_rule == 5:  # Exactly
                    exact_height = paragraph.paragraph_format.line_spacing.pt
                    style.append(f"line-height: {exact_height}pt")
                else:  # Multiple
                    if paragraph.paragraph_format.line_spacing:
                        multiplier = float(paragraph.paragraph_format.line_spacing)
                        style.append(f"line-height: {multiplier}")
            else:
                style.append("line-height: 1.15")  # Mặc định 1.15
            
            # Thêm white-space để giữ nguyên khoảng trắng
            style.append("white-space: pre-wrap")
            
            return style
        except Exception as e:
            logging.warning(f"Lỗi khi lấy thông tin khoảng cách: {str(e)}")
            return ["margin-top: 0pt", "margin-bottom: 0pt", "line-height: 1.15", "white-space: pre-wrap"]

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
                        width: 100% !important;
                        max-width: 800px !important;
                        margin: 0 auto !important;
                        white-space: pre-wrap !important;
                    }
                    .paragraph {
                        margin: 0 !important;
                        padding: 0 !important;
                        width: 100% !important;
                        display: block !important;
                        word-wrap: break-word !important;
                        word-break: break-word !important;
                    }
                    p {
                        margin: 0 !important;
                        padding: 0 !important;
                        white-space: pre-wrap !important;
                    }
                    br {
                        display: block !important;
                        content: "" !important;
                        margin-top: 0 !important;
                    }
                    br.line-break {
                        display: block !important;
                        margin: 12pt 0 !important;
                        content: "" !important;
                        height: 1em !important;
                    }
                    br.line-break + br.line-break {
                        display: none !important;
                    }
                    .paragraph:empty {
                        height: 1em !important;
                        margin: 12pt 0 !important;
                    }
                </style>
            """)
            
            content.append('<div class="email-content">')
            
            # Chuyển đổi từng đoạn văn sang HTML
            prev_was_empty = False
            for paragraph in doc.paragraphs:
                if paragraph.text.strip():
                    html = self._get_paragraph_style(paragraph)
                    
                    try:
                        style = []
                        
                        # Thêm căn lề
                        if paragraph.alignment is not None:
                            alignment_map = {
                                0: 'left',
                                1: 'center',
                                2: 'right',
                                3: 'justify'
                            }
                            alignment_value = int(paragraph.alignment)
                            text_align = alignment_map.get(alignment_value, 'left')
                            style.append(f"text-align: {text_align}")
                        
                        # Thêm khoảng cách
                        spacing_styles = self._get_paragraph_spacing(paragraph)
                        style.extend(spacing_styles)
                        
                        # Nếu đoạn trước là trống, thêm margin-top
                        if prev_was_empty:
                            style.append("margin-top: 12pt")
                        
                        # Tạo HTML với đầy đủ style
                        if style:
                            style_str = "; ".join(f"{s} !important" for s in style)
                            html = f'<div class="paragraph" style="{style_str}">{html}</div>'
                        else:
                            html = f'<div class="paragraph">{html}</div>'
                        
                        prev_was_empty = False
                            
                    except Exception as style_err:
                        logging.warning(f"Lỗi xử lý style: {str(style_err)}")
                        html = f'<div class="paragraph">{html}</div>'
                    
                    content.append(html)
                else:
                    # Thêm đoạn văn trống để giữ khoảng cách
                    if not prev_was_empty:  # Chỉ thêm nếu đoạn trước không trống
                        content.append('<div class="paragraph" style="height: 1em !important; margin: 12pt 0 !important;"></div>')
                        prev_was_empty = True

            content.append("</div>")

            # Xử lý ảnh
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