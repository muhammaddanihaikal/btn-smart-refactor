import zipfile, lxml.etree as etree

def inspect_cover(docx_path):
    print("=== COVER of", docx_path, "===")
    with zipfile.ZipFile(docx_path, 'r') as z:
        doc_xml = etree.fromstring(z.read('word/document.xml'))
        rels_xml = etree.fromstring(z.read('word/_rels/document.xml.rels'))
        
        rels = {rel.attrib['Id']: rel.attrib['Target'] for rel in rels_xml}
        ns = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        }
        r_ns = ns['r']
        for i, p in enumerate(doc_xml.findall('.//w:p', ns)[:10]):
            for b in p.findall('.//a:blip', ns):
                embed_id = b.attrib.get(f'{{{r_ns}}}embed')
                target = rels.get(embed_id, 'NOT FOUND')
                media_path = 'word/' + target if not target.startswith('word/') else target
                exists = media_path in z.namelist()
                size = len(z.read(media_path)) if exists else 0
                print(f"P {i}: Drawing rId={embed_id} -> {target} (exists={exists}, size={size})")

inspect_cover(r'H:\My Drive\Zegen\BTN Smart\Refactor\Hasil Uji\Dokumen_Hasil_Uji_Web.docx')
inspect_cover(r'D:\Project\BTN Smart\Refactor\Hasil Uji\Template Dokumen Hasil Uji.docx')
