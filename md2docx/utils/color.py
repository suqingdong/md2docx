from docx.shared import RGBColor


def color_to_rgb(color_str: str) -> RGBColor:
    """Convert color string to RGBColor object.

    Args:
        color_str (str): Color string in the format of "#RRGGBB" or 'RRGGBB'.

    Returns:
        RGBColor: RGBColor object.
    """
    return RGBColor.from_string(color_str.replace('#', '').strip())
