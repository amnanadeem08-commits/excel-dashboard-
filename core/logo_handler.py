"""Logo upload helpers for workbook generation."""

from io import BytesIO

from openpyxl.drawing.image import Image as OpenPyXLImage
from PIL import Image as PILImage

from core.styles import placeholder_style

SUPPORTED_LOGO_TYPES = {"png", "jpg", "jpeg"}
DEFAULT_LOGO_SIZE = (220, 90)


def prepare_logo(uploaded_logo, max_size=DEFAULT_LOGO_SIZE):
    """Return resized logo bytes that openpyxl can insert, or None when unavailable."""
    if uploaded_logo is None:
        return None

    file_type = _detect_file_type(uploaded_logo)
    if file_type not in SUPPORTED_LOGO_TYPES:
        return None

    uploaded_logo.seek(0)
    image_bytes = uploaded_logo.read()

    with PILImage.open(BytesIO(image_bytes)) as image:
        image.thumbnail(max_size)
        output = BytesIO()
        image_format = "PNG" if file_type == "png" else "JPEG"
        if image_format == "JPEG" and image.mode in {"RGBA", "P"}:
            image = image.convert("RGB")
        image.save(output, format=image_format)
        output.seek(0)
        return output.read()


def add_logo_or_placeholder(worksheet, uploaded_logo, palette, anchor="F2", placeholder_range="F2:G5"):
    """Add a resized logo when uploaded, otherwise add a visible placeholder."""
    logo_bytes = prepare_logo(uploaded_logo)
    if logo_bytes:
        _add_logo_to_sheet(worksheet, logo_bytes, anchor)
        return "logo"

    start_cell, end_cell = placeholder_range.split(":")
    worksheet.merge_cells(placeholder_range)
    cell = worksheet[start_cell]
    cell.value = "Your Logo Here"
    placeholder_style(cell, palette)
    return "placeholder"


def _add_logo_to_sheet(worksheet, logo_bytes, anchor):
    logo_stream = BytesIO(logo_bytes)
    logo = OpenPyXLImage(logo_stream)
    worksheet.add_image(logo, anchor)


def _detect_file_type(uploaded_logo):
    content_type = getattr(uploaded_logo, "type", "") or ""
    if "/" in content_type:
        return content_type.split("/")[-1].lower()

    name = getattr(uploaded_logo, "name", "") or ""
    if "." in name:
        return name.rsplit(".", 1)[-1].lower()

    return ""
