import re


def is_table_row(line: str) -> bool:
    """判断是否为表格行"""
    line = line.strip()
    return bool(line and '|' in line and line.count('|') >= 2)


def is_table_separator(line: str) -> bool:
    """判断是否为表格分隔符行"""
    line = line.strip()
    # 更精确的分隔符识别：包含|和-，可能包含:用于对齐
    if not line or '|' not in line:
        return False
    # 移除首尾的|，检查内容
    content = line.strip('|').strip()
    # 分隔符行应该主要由-、:、空格和|组成
    return bool(re.match(r'^[\s\-:|]+', content))


def is_image_line(line: str) -> bool:
    """判断是否为图像行"""
    line = line.strip()
    return bool(re.match(r'!\[.*?\]\(.*?\)', line))


def is_base64_image(path: str) -> bool:
    """判断是否为Base64编码的图像"""
    return path.startswith('data:image/')


def is_url_image(path: str) -> bool:
    """判断是否为URL"""
    return path.startswith(('http://', 'https://'))


def is_code_block(line: str) -> bool:
    """判断是否为代码块"""
    return line.strip().startswith('```')


def is_quote(line: str) -> bool:
    """判断是否为引用"""
    return line.strip().startswith('>')


def is_heading(line: str) -> bool:
    """判断是否为标题"""
    return bool(re.match(r'^#+ ', line))


def is_list_item(line: str) -> bool:
    """判断是否为列表项"""
    return bool(re.match(r'^\s*[-\*\+]+ ', line))