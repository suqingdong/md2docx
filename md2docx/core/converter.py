from pathlib import Path
from docx import Document
import loguru

from md2docx.utils import DocStyle
from . import processors, line_type


class MD2DOCX(object):

    def get_content(self, md):
        try:
            if Path(md).is_file():
                with open(md, encoding='utf-8') as f:
                    md = f.read()
        except Exception as e:
            pass
        return md

    def convert(
            self,
            md,
            output_file='output.docx',
            code_as_image=False,
            default_font='Times New Roman',
            chinese_font='宋体',
            default_font_size=12,
            heading_color='17a2b8',
            render_mermaid=True,
            
        ):
        doc = Document()

        md = self.get_content(md)
        md = processors.preprocess_checkboxes(md)
        if code_as_image:
            md = processors.preprocess_code_blocks(md)

        lines = md.strip().split('\n')

        i = 0
        while i < len(lines):
            line = lines[i]

            # 处理代码块
            if line_type.is_code_block(line):
                i = processors.process_code_block(doc, lines, i, render_mermaid=render_mermaid)
            # 处理表格
            elif line_type.is_table_row(line):
                i = processors.process_table(doc, lines, i, alignment='center')
            # 处理图像
            elif line_type.is_image_line(line):
                processors.process_image(doc, line)
                i += 1
            # 处理标题
            elif line_type.is_heading(line):
                processors.process_heading(doc, line)
                i += 1
            # 处理列表
            elif line_type.is_list_item(line):
                processors.process_list_item(doc, line, inline_color='FF0000', inline_background='F2F2F2')
                i += 1
            # 处理引用
            elif line_type.is_quote(line):
                processors.process_quote(doc, line, left_indent=24, italic=True, color='F2F2F2')
                i += 1
            # 处理普通段落
            elif line.strip():
                processors.process_paragraph(doc, line, inline_color='FF0000', inline_background='F2F2F2')
                i += 1
            else:
                # 空行
                i += 1

        DocStyle(
            doc,
            default_font=default_font,
            chinease_font=chinese_font,
            font_size=default_font_size,
            heading_color=heading_color,
        ).set_styles()

        if output_file:
            doc.save(output_file)
            loguru.logger.debug(f'save docx to: {output_file}')

        return doc
