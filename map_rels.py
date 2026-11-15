import zipfile
import xml.etree.ElementTree as ET

with zipfile.ZipFile('ANEXO 004 - BORRADOR DE TESIS EPG UNA - CUANTITATIVO020125.docx', 'r') as z:
    rels_xml = z.read('word/_rels/document.xml.rels')
    root = ET.fromstring(rels_xml)
    for elem in root:
        rId = elem.attrib.get('Id')
        target = elem.attrib.get('Target')
        if 'media' in target:
            print(f"{rId} -> {target}")
