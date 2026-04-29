from paddleocr import PaddleOCR
import sys

ocr = PaddleOCR(lang="chinese_cht")

def run(image_path):
    result = ocr.ocr(image_path)
    for line in result[0]:
        text, confidence = line[1]
        print(f"{text}  ({confidence:.2f})")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("用法: python3 ocr.py <圖片路徑>")
        sys.exit(1)
    run(sys.argv[1])
