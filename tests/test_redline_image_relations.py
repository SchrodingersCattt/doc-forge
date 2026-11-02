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


NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}
REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'


def _assert_drawing_relationships(package: Package) -> None:
    root = package.xml('word/document.xml')
    relationships = {item.get('Id'): item for item in package.xml('word/_rels/document.xml.rels')}
    embeds = root.xpath('//a:blip/@r:embed', namespaces=NS)
    links = root.xpath('//a:blip/@r:link', namespaces=NS)
    assert embeds
    for rid in [*embeds, *links]:
        relationship = relationships[rid]
        assert relationship.get('Type') == RELATIONSHIP_TYPE.IMAGE
        if relationship.get('TargetMode') == 'External':
            continue
        target = relationship.get('Target')
        assert f'word/{target}' in package.parts


def _picture_paragraph(document: Document, text: str, image, width):
    paragraph = document.add_paragraph(text)
    paragraph.add_run().add_picture(str(image), width=width)
    return paragraph


def test_table_cell_drawing_uses_output_relationship(tmp_path):
    image = tmp_path / 'sample.png'
    Image.new('RGB', (4, 4), 'red').save(image)
    base, current, output = (tmp_path / name for name in ('base.docx', 'current.docx', 'tracked.docx'))
    for path, text, extra, width in (
        (base, 'Initial ', False, 1),
        (current, 'Updated ', True, 2),
    ):
        document = Document()
        document.add_paragraph('Shared context stays the same.')
        if extra:
            document.part.relate_to('https://example.org/', RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
        cell = document.add_table(rows=1, cols=1).cell(0, 0)
        cell.text = ''
        paragraph = cell.paragraphs[0]
        paragraph.add_run(text)
        paragraph.add_run().add_picture(str(image), width=Inches(width))
        document.add_paragraph('More shared context stays the same.')
        document.save(path)
    create_tracked_docx(base, current, output, workers=1)
    _assert_drawing_relationships(Package.load(output))


def test_deleted_table_drawing_uses_output_relationship(tmp_path):
    image = tmp_path / 'sample.png'
    Image.new('RGB', (4, 4), 'red').save(image)
    base, current, output = (tmp_path / name for name in ('base.docx', 'current.docx', 'tracked.docx'))
    document = Document()
    document.add_paragraph('Intro stays put for alignment.')
    cell = document.add_table(rows=1, cols=1).cell(0, 0)
    cell.paragraphs[0].add_run('Gone ')
    cell.paragraphs[0].add_run().add_picture(str(image))
    document.add_paragraph('Outro stays put for alignment.')
    document.save(base)
    document = Document()
    document.add_paragraph('Intro stays put for alignment.')
    document.part.relate_to('https://example.org/', RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    document.add_paragraph('Replacement sentence without the table.')
    document.add_paragraph('Outro stays put for alignment.')
    document.save(current)
    create_tracked_docx(base, current, output, workers=1)
    _assert_drawing_relationships(Package.load(output))


def test_deleted_table_row_drawing_uses_output_relationship(tmp_path):
    image = tmp_path / 'sample.png'
    Image.new('RGB', (4, 4), 'red').save(image)
    base, current, output = (tmp_path / name for name in ('base.docx', 'current.docx', 'tracked.docx'))
    document = Document()
    document.add_paragraph('Intro stays put for alignment.')
    table = document.add_table(rows=2, cols=1)
    table.cell(0, 0).paragraphs[0].add_run('Keep row')
    gone = table.cell(1, 0).paragraphs[0]
    gone.add_run('Gone ')
    gone.add_run().add_picture(str(image))
    document.add_paragraph('Outro stays put for alignment.')
    document.save(base)
    document = Document()
    document.add_paragraph('Intro stays put for alignment.')
    document.part.relate_to('https://example.org/', RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    document.add_table(rows=1, cols=1).cell(0, 0).paragraphs[0].add_run('Keep row')
    document.add_paragraph('Outro stays put for alignment.')
    document.save(current)
    create_tracked_docx(base, current, output, workers=1)
    package = Package.load(output)
    _assert_drawing_relationships(package)
    root = package.xml('word/document.xml')
    assert root.xpath('//w:trPr/w:del', namespaces=NS)


def test_unreferenced_baseline_media_is_not_copied(tmp_path):
    image = tmp_path / 'sample.png'
    Image.new('RGB', (4, 4), 'red').save(image)
    base, current, output = (tmp_path / name for name in ('base.docx', 'current.docx', 'tracked.docx'))
    for path, text, extra in ((base, 'Initial ', False), (current, 'Updated ', True)):
        document = Document()
        document.add_paragraph('Shared context stays the same here now.')
        if extra:
            document.part.relate_to('https://example.org/', RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
        _picture_paragraph(document, text, image, Inches(1))
        document.add_paragraph('More shared context stays the same here now.')
        document.save(path)
    create_tracked_docx(base, current, output, workers=1)
    package = Package.load(output)
    _assert_drawing_relationships(package)
    media = [name for name in package.parts if name.startswith('word/media/')]
    assert media == ['word/media/image1.png']


def test_drawing_link_is_remapped_with_embed(tmp_path):
    image = tmp_path / 'sample.png'
    Image.new('RGB', (4, 4), 'red').save(image)
    base, current, output = (tmp_path / name for name in ('base.docx', 'current.docx', 'tracked.docx'))
    document = Document()
    document.add_paragraph('Context alpha is stable enough.')
    _picture_paragraph(document, 'Before ', image, Inches(1))
    document.add_paragraph('Context omega is stable enough.')
    document.save(base)
    package = Package.load(base)
    root = package.xml('word/document.xml')
    rels = package.xml('word/_rels/document.xml.rels')
    link = etree.SubElement(rels, f'{{{REL_NS}}}Relationship')
    link.set('Id', 'rIdLink9')
    link.set('Type', RELATIONSHIP_TYPE.IMAGE)
    link.set('Target', 'https://example.org/linked.png')
    link.set('TargetMode', 'External')
    blip = root.xpath('//a:blip', namespaces=NS)[0]
    blip.set(f'{{{NS["r"]}}}link', 'rIdLink9')
    package.set_xml('word/document.xml', root)
    package.set_xml('word/_rels/document.xml.rels', rels)
    package.write(base, overwrite=True)
    document = Document()
    document.add_paragraph('Context alpha is stable enough.')
    document.part.relate_to('https://example.org/', RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    _picture_paragraph(document, 'After ', image, Inches(2))
    document.add_paragraph('Context omega is stable enough.')
    document.save(current)
    create_tracked_docx(base, current, output, workers=1)
    package = Package.load(output)
    _assert_drawing_relationships(package)
    links = package.xml('word/document.xml').xpath('//a:blip/@r:link', namespaces=NS)
    assert links
    assert 'rIdLink9' not in links


def test_insertion_beside_a_drawing_keeps_closing_punctuation(tmp_path):
    image = tmp_path / 'sample.png'
    Image.new('RGB', (8, 8), 'red').save(image)
    base, current, output = (tmp_path / name for name in ('base.docx', 'current.docx', 'tracked.docx'))
    for path, text in ((base, 'A model (Fig. 1b).'), (current, 'A model (Fig. 1b; Note 1).')):
        document = Document()
        document.add_paragraph('Shared lead stays.')
        paragraph = document.add_paragraph()
        paragraph.add_run(text)
        paragraph.add_run().add_picture(str(image), width=Inches(1))
        document.add_paragraph('Shared tail stays.')
        document.save(path)
    create_tracked_docx(base, current, output, workers=1)
    package = Package.load(output)
    root = package.xml('word/document.xml')
    deleted = ''.join(root.xpath('//w:delText/text()', namespaces=NS))
    inserted = [''.join(node.xpath('.//w:t/text()', namespaces=NS)) for node in root.xpath('//w:ins', namespaces=NS)]
    assert deleted == ''
    assert inserted == ['; Note 1']
    assert len(root.xpath('//w:drawing', namespaces=NS)) == 1
    _assert_drawing_relationships(package)


def test_unchanged_drawing_with_shifted_relationship_stays_valid(tmp_path):
    image = tmp_path / 'sample.png'
    Image.new('RGB', (8, 8), 'red').save(image)
    base, current, output = (tmp_path / name for name in ('base.docx', 'current.docx', 'tracked.docx'))
    for path, extra in ((base, False), (current, True)):
        document = Document()
        document.add_paragraph('Intro paragraph that is identical.')
        if extra:
            document.part.relate_to('https://example.org/page', RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
        paragraph = document.add_paragraph('Caption stays')
        paragraph.add_run().add_picture(str(image), width=Inches(1))
        document.add_paragraph('Outro paragraph that is identical.')
        document.save(path)
    create_tracked_docx(base, current, output, workers=1)
    package = Package.load(output)
    root = package.xml('word/document.xml')
    assert len(root.xpath('//w:drawing', namespaces=NS)) == 1
    assert root.xpath('//w:ins', namespaces=NS) == []
    assert root.xpath('//w:del', namespaces=NS) == []
    _assert_drawing_relationships(package)
