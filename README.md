# A user-friendly tool for converting Markdown to DOCX

## Installation

```bash
pip install python-md2docx
```

## Usage

### Use in CMD

```bash
md2docx --help

md2docx tests/demo.md -o tests/demo-default.docx
md2docx tests/demo.md -o tests/demo-code-as-image.docx --code-as-image
md2docx tests/demo.md -o tests/demo-render-mermaid.docx --render-mermaid 
md2docx tests/demo.md -o tests/demo-styles.docx --heading-color FF00FF --default-font Arial --chinese-font 微软雅黑
```

### Use in Python

```python
from md2docx.core import MD2DOCX

MD2DOCX().convert('tests/demo.md', output_file='tests/demo.docx')

MD2DOCX().convert(
    'tests/demo.md',
    output_file='tests/demo.docx',
    heading_color='FF00FF',
    default_font='Arial',
    chinese_font='微软雅黑',
    code_as_image=True,
)
```

## Support Features
- [x] Heading
- [x] Link
- [x] Image
- [x] Table
- [x] Inline Code
- [x] Code Block
- [x] Code Block as Image (pygments)
- [x] Quote
- [x] Checkbox
- [x] Mermaid
- [x] Latex
