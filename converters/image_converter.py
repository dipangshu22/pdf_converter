"""
Image to PDF Converter
Supports: PNG, JPG, JPEG, BMP, GIF, TIFF, WEBP
"""

import img2pdf
from PIL import Image
import io
import os


SUPPORTED_FORMATS = {'.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff', '.tif', '.webp'}


def convert_image_to_pdf(input_path: str, output_path: str) -> dict:
    """
    Convert an image file to PDF.

    Args:
        input_path: Path to the input image file
        output_path: Path to save the output PDF

    Returns:
        dict with 'success' bool and 'message' string
    """
    ext = os.path.splitext(input_path)[1].lower()
    if ext not in SUPPORTED_FORMATS:
        return {
            'success': False,
            'message': f"Unsupported image format '{ext}'. Supported: {', '.join(SUPPORTED_FORMATS)}"
        }

    try:
        # Normalize image: convert to RGB (handles RGBA, palette, etc.)
        with Image.open(input_path) as img:
            # Convert to RGB if necessary (img2pdf needs JPEG-compatible mode)
            if img.mode in ('RGBA', 'P', 'LA'):
                background = Image.new('RGB', img.size, (255, 255, 255))
                if img.mode == 'P':
                    img = img.convert('RGBA')
                background.paste(img, mask=img.split()[-1] if img.mode in ('RGBA', 'LA') else None)
                img = background
            elif img.mode != 'RGB':
                img = img.convert('RGB')

            # Save normalized image to buffer
            buf = io.BytesIO()
            img.save(buf, format='JPEG', quality=95)
            buf.seek(0)

        # Convert to PDF using img2pdf
        pdf_bytes = img2pdf.convert(buf)

        with open(output_path, 'wb') as f:
            f.write(pdf_bytes)

        size_kb = os.path.getsize(output_path) / 1024
        return {
            'success': True,
            'message': f"Image converted successfully. Output size: {size_kb:.1f} KB"
        }

    except Exception as e:
        return {'success': False, 'message': f"Image conversion failed: {str(e)}"}