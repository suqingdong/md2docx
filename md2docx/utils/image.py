import io
import base64
from typing import Optional, Tuple, Union

import requests
from PIL import Image

from docx.shared import Inches


from md2docx.core import line_type


class ImageLoader(object):

    @classmethod
    def _image_from_url(cls, url: str):
        try:
            response = requests.get(url, timeout=10)
            image = Image.open(io.BytesIO(response.content))
            return image
        except Exception as e:
            print(f'Error loading image from url: {url}')
            print(f'Error: {e}')

    @classmethod
    def _image_from_base64(cls, b64_str: str):
        try:
            b64_str = b64_str.split(',')[1]
            image_data = base64.b64decode(b64_str)
            image = Image.open(io.BytesIO(image_data))
            return image
        except Exception as e:
            print(f'Error loading image from base64: {b64_str[:50]}')
            print(f'Error: {e}')
        
    @classmethod
    def _image_from_file(cls, file: str):
        try:
            return Image.open(file)
        except Exception as e:
            print(f'Error loading image from file: {file}')
            print(f'Error: {e}')
    
    @classmethod
    def load_image(cls, image_path: str, ):
        image = None
        if line_type.is_url_image(image_path):
            image = cls._image_from_url(image_path)
        elif line_type.is_base64_image(image_path):
            image = cls._image_from_base64(image_path)
        else:
            image = cls._image_from_file(image_path)
        return image



def calculate_image_size(image: Union[Image.Image, str]) -> Tuple[Inches, Optional[Inches]]:
    """计算图像在文档中的尺寸
    """
    if isinstance(image, str):
        image = Image.open(image)

    width, height = image.size
    
    # 设置最大宽度为6英寸
    max_width = Inches(6)
    
    # 计算比例
    ratio = width / height
    
    if width > height:
        # 横向图像
        new_width = max_width
        new_height = max_width / ratio
    else:
        # 纵向图像
        max_height = Inches(8)
        if height > width:
            new_height = min(max_height, max_width / ratio)
            new_width = new_height * ratio
        else:
            new_width = max_width
            new_height = max_width / ratio

    return new_width, new_height