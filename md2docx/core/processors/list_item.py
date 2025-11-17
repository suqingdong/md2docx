import re
from docx import Document

from .paragraph import process_inline_styles


def process_list_item(
        doc: Document,
        line: str,
        inline_color: str = 'C72E96',
        inline_background: str = 'F2F2F2',    
    ):
    """处理列表项"""
    # 移除列表标记
    text = re.sub(r'^[\s]*[-*]\s*', '', line)
    p = doc.add_paragraph(style='List Bullet')

    process_inline_styles(p, text, inline_color, inline_background)
