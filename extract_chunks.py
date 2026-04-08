import zipfile
import xml.etree.ElementTree as ET
import csv
import re
import os

filepath = r'e:\English\toeic_20_articles_bilingual.docx'

with zipfile.ZipFile(filepath, 'r') as z:
    xml_content = z.read('word/document.xml')

tree = ET.fromstring(xml_content)
ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

paragraphs = tree.findall('.//w:p', ns)
all_text = []
for p in paragraphs:
    texts = p.findall('.//w:t', ns)
    line = ''.join(t.text for t in texts if t.text)
    all_text.append(line)

chunks = []
current_article = None
in_chunk_section = False
chunk_en = None

for i, line in enumerate(all_text):
    article_match = re.match(r'第\s*(\d+)\s*篇', line)
    if article_match:
        current_article = int(article_match.group(1))
        in_chunk_section = False

    if '核心词块' in line:
        in_chunk_section = True
        continue

    if '精背提示' in line:
        in_chunk_section = False
        continue

    if in_chunk_section and current_article:
        stripped = line.strip()
        if not stripped:
            continue
        if stripped in ['English Chunk', '中文释义']:
            continue
        if chunk_en is None:
            chunk_en = stripped
        else:
            chunks.append({
                'article': current_article,
                'english': chunk_en,
                'chinese': stripped
            })
            chunk_en = None

output_path = r'e:\English\04_Vocabulary_Anki\toeic_core_chunks.csv'
with open(output_path, 'w', newline='', encoding='utf-8-sig') as f:
    writer = csv.writer(f)
    writer.writerow(['Article', 'English Chunk', 'Chinese'])
    for c in chunks:
        writer.writerow([c['article'], c['english'], c['chinese']])

print(f'Total chunks extracted: {len(chunks)}')
print(f'Saved to: {output_path}')
print()
for c in chunks[:10]:
    print(f"  Art.{c['article']}: {c['english']} -> {c['chinese']}")
print('...')
