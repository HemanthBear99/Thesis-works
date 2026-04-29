from docx import Document
doc = Document(r'D:/Project/thesis/project-james/Thesis James Chintalha (1) (1).docx')
paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
# Find word MSc/degree lines and title
for i, p in enumerate(paragraphs[:20]):
    print(f'{i+1}: {p}')
print('\n... looking for ABSTRACT ...')
# Search for abstract in the document
idx = -1
for i, p in enumerate(paragraphs):
    if 'abstract' in p.lower():
        idx = i
        break
if idx >= 0:
    for j in range(idx, min(idx+20, len(paragraphs))):
        print(f'{j}: {paragraphs[j]}')
