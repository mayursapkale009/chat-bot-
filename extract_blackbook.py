import json
from docx import Document
import sys

doc_path = r'e:\MAAM\chatbot code\docs\Black book\Black_Book_Final.docx'
output_path = r'e:\MAAM\chatbot code\docs\Black_book_extracted.txt'

doc = Document(doc_path)
text = '\n'.join(p.text for p in doc.paragraphs)
with open(output_path, 'w', encoding='utf-8') as f:
    f.write(text)
print('Extraction complete')
