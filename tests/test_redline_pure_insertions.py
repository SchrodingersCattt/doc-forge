from docx import Document
from lxml import etree

from docforge.docxdiff import create_tracked_docx
from docforge.docxdiff.package import Package


def test_insert_before_closing_punctuation_does_not_delete_original(tmp_path):
    base = tmp_path / 'base.docx'
    current = tmp_path / 'current.docx'
    output = tmp_path / 'tracked.docx'
    originals = ['A model (Fig. 1b).', 'Conditions (Tables S4–S7).', 'Transfer learning.']
    edited = ['A model (Fig. 1b; Supplementary Note 1).',
              'Conditions (Tables S4–S7; Supplementary Note 3).',
              'Transfer learning (Supplementary Note 1).']
    for path, paragraphs in ((base, originals), (current, edited)):
        doc = Document()
        for text in paragraphs:
            doc.add_paragraph(text)
        doc.save(path)
    create_tracked_docx(base, current, output, workers=1)
    root = Package.load(output).xml('word/document.xml')
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    assert not root.xpath('//w:del', namespaces=ns)
    inserted = [''.join(n.xpath('.//w:t/text()', namespaces=ns)) for n in root.xpath('//w:ins', namespaces=ns)]
    assert inserted == ['; Supplementary Note 1', '; Supplementary Note 3', ' (Supplementary Note 1)']
