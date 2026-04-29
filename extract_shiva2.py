from docx import Document
doc = Document(r'D:/Project/thesis/project-shiva/18-02-26/Impact of AI (23-02-26).docx')
paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
for i, p in enumerate(paragraphs[:50]):
    print(f'{i+1}: {p}')
