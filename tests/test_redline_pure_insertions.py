from docx import Document

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


def _revisions(path):
    root = Package.load(path).xml('word/document.xml')
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    deleted = ''.join(root.xpath('//w:delText/text()', namespaces=ns))
    inserted = [''.join(node.xpath('.//w:t/text()', namespaces=ns)) for node in root.xpath('//w:ins', namespaces=ns)]
    return deleted, inserted


def _track(tmp_path, original: str, edited: str):
    base, current, output = (tmp_path / name for name in ('base.docx', 'current.docx', 'tracked.docx'))
    for path, text in ((base, original), (current, edited)):
        document = Document()
        document.add_paragraph(text)
        document.save(path)
    create_tracked_docx(base, current, output, workers=1)
    return _revisions(output)


def test_several_insertions_keep_each_closing_mark(tmp_path):
    deleted, inserted = _track(
        tmp_path,
        'See (Fig. 1) and (Fig. 2).',
        'See (Fig. 1; Note A) and (Fig. 2; Note B).',
    )
    assert deleted == ''
    assert inserted == ['; Note A', '; Note B']


def test_other_edits_do_not_revise_punctuation_around_an_insertion(tmp_path):
    deleted, inserted = _track(
        tmp_path,
        'The model (Fig. 1b).',
        'The updated model (Fig. 1b; Note 1).',
    )
    assert ')' not in deleted
    assert '1b' not in deleted
    assert 'updated' in ''.join(inserted)
    assert '; Note 1' in ''.join(inserted)


def test_fullwidth_punctuation_stays_outside_an_insertion(tmp_path):
    deleted, inserted = _track(tmp_path, '已完成（实验）。', '已完成（实验；补充）。')
    assert deleted == ''
    assert inserted == ['；补充']


def test_insertion_inside_a_word_stays_word_level(tmp_path):
    deleted, inserted = _track(tmp_path, 'The color is red.', 'The colour is red.')
    assert deleted == 'color'
    assert inserted == ['colour']
