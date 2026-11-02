from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.shared import Inches
from lxml import etree
from PIL import Image

from docforge.docxdiff import create_tracked_docx
from docforge.docxdiff.package import Package


def test_inline_drawing_from_baseline_uses_output_relationship(tmp_path):
    image = tmp_path / 'sample.png'
    Image.new('RGB', (4, 4), 'red').save(image)
    base = tmp_path / 'base.docx'
    current = tmp_path / 'current.docx'
    output = tmp_path / 'tracked.docx'
    for path, text, extra in ((base, 'Initial text ', False), (current, 'Updated text ', True)):
        document = Document()
        document.add_paragraph('Shared context stays the same.')
        if extra:
            document.part.relate_to('https://example.org/', RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
        paragraph = document.add_paragraph(text)
        paragraph.add_run().add_picture(str(image), width=Inches(2 if extra else 1))
        document.add_paragraph('More shared context stays the same.')
        document.save(path)
    create_tracked_docx(base, current, output, workers=1)
    package = Package.load(output)
    root = package.xml('word/document.xml')
    rels = package.xml('word/_rels/document.xml.rels')
    relationships = {item.get('Id'): item for item in rels}
    ns = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
          'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
    ids = root.xpath('//a:blip/@r:embed', namespaces=ns)
    assert ids
    assert all(relationships[rid].get('Type') == RELATIONSHIP_TYPE.IMAGE for rid in ids)


def test_matching_image_relationship_reuses_original_media(tmp_path):
    image = tmp_path / 'sample.png'
    Image.new('RGB', (4, 4), 'red').save(image)
    base, current, output = (tmp_path / name for name in ('base.docx', 'current.docx', 'tracked.docx'))
    for path, text in ((base, 'Old text '), (current, 'New text ')):
        document = Document()
        document.add_paragraph('Shared context remains.')
        paragraph = document.add_paragraph(text)
        paragraph.add_run().add_picture(str(image))
        document.add_paragraph('Additional shared context.')
        document.save(path)
    create_tracked_docx(base, current, output, workers=1)
    package = Package.load(output)
    assert [name for name in package.parts if name.startswith('word/media/')] == ['word/media/image1.png']
