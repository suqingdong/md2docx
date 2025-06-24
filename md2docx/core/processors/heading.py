from docx import Document


def process_heading(doc: Document, line: str):
    """处理标题"""
    level = 0
    while level < len(line) and line[level] == '#':
        level += 1
    
    title_text = line[level:].strip()
    doc.add_heading(title_text, level=min(level, 3))
