import zipfile
import os
import docx
from docx.oxml import parse_xml

os.makedirs('extracted_figures', exist_ok=True)

# Extract all images
with zipfile.ZipFile('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx', 'r') as z:
    for name in z.namelist():
        if name.startswith('word/media/'):
            basename = os.path.basename(name)
            target = os.path.join('extracted_figures', basename)
            with open(target, 'wb') as f:
                f.write(z.read(name))

print("Extracted all images into extracted_figures/")

# Now find where images are referenced in docx paragraphs
doc = docx.Document('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx')

print(f"Total paragraphs: {len(doc.paragraphs)}")

for i, p in enumerate(doc.paragraphs):
    # Check if paragraph has drawing / blip
    xml = p._p.xml
    if 'blip' in xml or 'graphic' in xml or 'Figura' in p.text:
        # Check blip r:embed
        import re
        embeds = re.findall(r'r:embed="(rId\d+)"', xml)
        print(f"P{i:03d} [{p.style.name}]: text='{p.text[:80]}' embeds={embeds}")
