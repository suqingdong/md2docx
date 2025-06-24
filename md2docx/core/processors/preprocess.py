import io
import re
import base64

from PIL import Image
from pygments import highlight
from pygments.lexers import get_lexer_by_name
from pygments.formatters import ImageFormatter


def upscale_image(data: bytes, scale=2) -> bytes:
    img = Image.open(io.BytesIO(data))
    width, height = img.size
    img = img.resize((width * scale, height * scale), Image.LANCZOS)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def render_code_to_base64(code: str, lang: str, with_linenos=True) -> str:
    try:
        lexer = get_lexer_by_name(lang, stripall=True)
    except Exception:
        lexer = get_lexer_by_name("text", stripall=True)

    formatter = ImageFormatter(
        style="monokai",
        # font_name="DejaVu Sans Mono",
        # font_name="Courier New",
        # font_name='SimHei',
        font_name='',
        font_size=24,
        line_numbers=with_linenos,
        line_pad=2,
        image_pad=10,
        line_number_bg="#272822",
        line_number_fg="#888888",
    )

    buffer = io.BytesIO()
    highlight(code, lexer, formatter, buffer)

    img_data = buffer.getvalue()

    # 放大缩小处理
    img_data = upscale_image(img_data, scale=2)

    base64_data = base64.b64encode(img_data).decode("utf-8")

    return f"data:image/png;base64,{base64_data}"


def preprocess_code_blocks(md_text: str) -> str:
    CODE_BLOCK_RE = re.compile(r'```(\w+)\n(.*?)```', re.DOTALL)

    def replacer(match):
        lang = match.group(1)
        code = match.group(2).strip()
        b64_uri = render_code_to_base64(code, lang, with_linenos=True)
        return f"![]({b64_uri})"

    return CODE_BLOCK_RE.sub(replacer, md_text)


def preprocess_checkboxes(md_text: str, checked_char: str = '☑ ', unchecked_char: str = '☐ ') -> str:
    md_text = re.sub(r'- \[ \] ', unchecked_char, md_text)
    md_text = re.sub(r'- \[x\] ', checked_char, md_text)
    return md_text
