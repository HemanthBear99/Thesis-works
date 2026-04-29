from docx import Document
doc = Document(r'D:/Project/thesis/project-james/Thesis James Chintalha (1) (1).docx')
paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
for i, p in enumerate(paragraphs[:80]):
    print(f'{i+1}. {p}')
