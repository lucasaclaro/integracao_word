import pandas as pd
import openpyxl
import docx

doc = docx.Document()
doc.save('arquivo.docx')

from docx import document

for paragrafo in doc.paragraphs:
    print(paragrafo.text)


# https://www.youtube.com/watch?v=N01MPYL3UVY
# 6m58s