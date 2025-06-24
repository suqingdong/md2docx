import re
from docx import Document


def process_list_item(doc: Document, line: str):
    """处理列表项"""
    # 移除列表标记
    text = re.sub(r'^[\s]*[-*]\s*', '', line)
    doc.add_paragraph(text, style='List Bullet')
