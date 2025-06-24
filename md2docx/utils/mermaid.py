import os
import tempfile
import subprocess as sp
import shutil

import loguru


def mermaid_to_png(mermaid_code: str):
    if not shutil.which('mmdc'):
        loguru.logger.warning('Mermaid CLI (mmdc) is not installed or not found in PATH.')
        return None

    with tempfile.NamedTemporaryFile(suffix='.mmd', delete=False) as mmd_file:
        mmd_file.write(mermaid_code.encode('utf-8'))
        mmd_file_path = mmd_file.name
        png_file_path = mmd_file_path.replace('.mmd', '.png')

    cmd = f"mmdc -i {mmd_file_path} -o {png_file_path}"
    res = sp.run(cmd, shell=True, capture_output=True, text=True)
    if res.returncode != 0:
        loguru.logger.warning(f'Mermaid rendering failed: {res.stderr}')
        return None

    os.unlink(mmd_file_path)
    return png_file_path


if __name__ == '__main__':
    mermaid_code = '''
flowchart TD
  A[Requirement Analysis] --> B[Product Design]
  B --> C[Technical Solution Design]
  C --> D[Frontend Development]
  C --> E[Backend Development]
  D --> F[Integration Testing]
  E --> F
  F --> G[Testing]
  G --> H[Deployment]
  H --> I[Operations and Iteration]
'''

    image = mermaid_to_png(mermaid_code)
    if image:
        image.show()