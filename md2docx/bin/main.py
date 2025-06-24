import time
import json

import click
import loguru
from md2docx import version_info
from md2docx.core import MD2DOCX


CONTEXT_SETTINGS = dict(help_option_names=['-?', '-h', '--help'])


epilog = click.style('''
\n\b
examples:
    {prog} --help
    {prog} --version
                     
    {prog} tests/demo.md -o tests/demo-default.docx
    {prog} tests/demo.md -o tests/demo-code-as-image.docx --code-as-image
    {prog} tests/demo.md -o tests/demo-render-mermaid.docx --render-mermaid 
    {prog} tests/demo.md -o tests/demo-styles.docx --heading-color FF00FF --default-font Arial --chinese-font 微软雅黑

''')

@click.command(
    name=version_info['prog'],
    help=click.style(version_info['desc'], italic=True, fg='cyan', bold=True),
    context_settings=CONTEXT_SETTINGS,
    no_args_is_help=True,
)
@click.argument('input_file')
@click.option('-o', '--output-file', help='the output file name', default='output.docx', show_default=True)
@click.option('--code-as-image', help='render code blocks as images', is_flag=True, default=False)
@click.option('--default-font', help='default font for the document', default='Times New Roman', show_default=True)
@click.option('--chinese-font', help='default Chinese font for the document', default='宋体', show_default=True)
@click.option('--default-font-size', help='default font size for the document', type=int, default=12, show_default=True)
@click.option('--heading-color', help='default heading color in hex format', default='17a2b8', show_default=True)
@click.option('--render-mermaid', help='render mermaid diagrams', is_flag=True, default=False)
@click.version_option(version=version_info['version'], prog_name=version_info['prog'])
def cli(**kwargs):
    start_time = time.time()
    loguru.logger.debug(f'input arguments:\n{json.dumps(kwargs, indent=4, ensure_ascii=False)}')

    MD2DOCX().convert(
        kwargs['input_file'],
        output_file=kwargs['output_file'],
        code_as_image=kwargs['code_as_image'],
        default_font=kwargs['default_font'],
        chinese_font=kwargs['chinese_font'],
        default_font_size=kwargs['default_font_size'],
        heading_color=kwargs['heading_color'],
        render_mermaid=kwargs['render_mermaid'],
    )

    loguru.logger.info(f'elapsed time: {time.time() - start_time:.2f} seconds')


def main():
    cli()


if __name__ == '__main__':
    main()
