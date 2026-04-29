from docx import Document
doc = Document(r'D:/Project/thesis/project-james/Thesis James Chintalha (1) (1).docx')
paragraphs = [p.text for p in doc.paragraphs if p.text.strip()]
# Find abstract section
for i, p in enumerate(paragraphs):
    if 'ABSTRACT' in p.upper() or 'Abstract' in p:
        print('--- ABSTRACT SECTION ---')
        for j in range(i, min(i+15, len(paragraphs))):
            print(f'{j}: {paragraphs[j]}')
        break
# Also print intro first paragraphs
print('\n--- FIRST 30 PARAGRAPHS ---')
for i, p in enumerate(paragraphs[:30]):
    print(f'{i+1}: {p}')
