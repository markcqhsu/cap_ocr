from PIL import Image, ImageEnhance, ImageFilter, ImageOps


def enhance(image_path: str) -> str:
    """
    Enhance image before OCR to improve recognition of unusual handwriting styles:
    - Grayscale conversion removes colour noise
    - Autocontrast stretches tonal range to full 0-255
    - Contrast boost makes ink stand out from paper
    - UnsharpMask sharpens stroke edges without amplifying grain
    """
    img = Image.open(image_path)

    img = img.convert("L")
    img = ImageOps.autocontrast(img, cutoff=2)
    img = ImageEnhance.Contrast(img).enhance(1.5)
    img = img.filter(ImageFilter.UnsharpMask(radius=1, percent=150, threshold=3))

    # Save as RGB so downstream APIs (Google Vision / Claude) handle any format
    img.convert("RGB").save(image_path)
    return image_path
