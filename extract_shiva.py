from docx import Document
doc = Document(r'D:/Project/thesis/project-shiva/18-02-26/Impact of AI (23-02-26).docx')
paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
print('--- FIRST 30 PARAGRAPHS ---')
for i, p in enumerate(paragraphs[:30]):
    print(f'{i+1}: {p}')
