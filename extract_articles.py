import zipfile
import xml.etree.ElementTree as ET
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

articles = []
current_article = None
current_title = None
in_english = False
english_lines = []

for line in all_text:
    article_match = re.match(r'第\s*(\d+)\s*篇\s+(.+)', line)
    if article_match:
        if current_article and english_lines:
            articles.append({
                'number': current_article,
                'title': current_title,
                'english': '\n'.join(english_lines).strip()
            })
        current_article = int(article_match.group(1))
        current_title = article_match.group(2).strip()
        in_english = False
        english_lines = []
        continue

    if line.strip() == 'English':
        in_english = True
        continue

    if line.strip() == '中文':
        in_english = False
        if current_article and english_lines:
            articles.append({
                'number': current_article,
                'title': current_title,
                'english': '\n'.join(english_lines).strip()
            })
            english_lines = []
        continue

    if in_english and line.strip():
        english_lines.append(line.strip())

if current_article and english_lines:
    articles.append({
        'number': current_article,
        'title': current_title,
        'english': '\n'.join(english_lines).strip()
    })

output_dir = r'e:\English\02_Listening_P3P4\articles_text'
os.makedirs(output_dir, exist_ok=True)

for art in articles:
    filename = f"article_{art['number']:02d}.txt"
    filepath_out = os.path.join(output_dir, filename)
    with open(filepath_out, 'w', encoding='utf-8') as f:
        f.write(art['english'])
    print(f"Saved: {filename} ({len(art['english'])} chars)")

print(f"\nTotal articles extracted: {len(articles)}")
