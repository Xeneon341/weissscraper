try:
    from PIL import Image
except ImportError:
    import Image
import pytesseract

pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

print(pytesseract.image_to_string(Image.open('static\INVOICE - 460292743 - BARDEVE0001 (20-Apr-21).PDF')))
