from pathlib import Path
from PyPDF2 import PdfReader

ROOT_FOLDER = Path(__file__).parent
WORKBOOK_PATH = ROOT_FOLDER / '191342.pdf'

reader = PdfReader(WORKBOOK_PATH)

print(len(reader.pages))

# for page in reader.pages:
#     print(page)
#     print()

page0 = reader.pages[0]

print(page0.extract_text())