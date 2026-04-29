import os

_backend = os.environ.get("OCR_BACKEND", "google").lower()

if _backend == "claude":
    from ocr_parser_claude import parse_form   # noqa: F401
else:
    from ocr_parser_google import parse_form   # noqa: F401
