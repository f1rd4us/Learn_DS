from docxtpl import DocxTemplate
import pandas as pd

data = pd.read_excel('Book1.xlsx')
isi = dict(zip(data['var'], data['value']))

doc = DocxTemplate('OTOWEB ADMINISTRASI.docx')
doc.render(isi)
doc.save('coba.docx')
