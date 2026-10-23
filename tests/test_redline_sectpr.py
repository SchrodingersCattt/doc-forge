"""Regression: redline must not emit sectPr header/footer refs without r:id."""

from __future__ import annotations

from lxml import etree

from docforge.docxdiff.package import Package
from docforge.docxdiff.redline import (
    R,
    W,
    _base_relationships,
    _strip_incomplete_header_footer_refs,
)

CONTENT_TYPES = "http://schemas.openxmlformats.org/package/2006/content-types"
EMPTY_CONTENT_TYPES = (
    f'<Types xmlns="{CONTENT_TYPES}"/>'.encode()
)


def test_redline_omits_baseline_header_refs_without_rid() -> None:
    """Deleted baseline sectPr must not emit header/footer refs without r:id.

    AML baselines may cite word/header*.xml. Those non-media relationships are
    skipped when carrying deleted paragraphs; the reference elements themselves
    must be removed, or Word refuses to open the tracked package.
    """
    header_type = (
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    )
    paragraph = etree.fromstring(
        f"""
        <w:p xmlns:w="{W}" xmlns:r="{R}">
          <w:pPr>
            <w:sectPr>
              <w:headerReference w:type="even" r:id="rId99"/>
              <w:footerReference w:type="default" r:id="rId99"/>
            </w:sectPr>
          </w:pPr>
          <w:r><w:t>baseline-only section break</w:t></w:r>
        </w:p>
        """
    )
    base_rels = etree.fromstring(
        f"""
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
          <Relationship Id="rId99" Type="{header_type}" Target="header2.xml"/>
        </Relationships>
        """
    )
    current_rels = etree.fromstring(
        """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>
        """
    )
    base = Package(
        {
            "[Content_Types].xml": EMPTY_CONTENT_TYPES,
            "word/_rels/document.xml.rels": etree.tostring(base_rels),
            "word/header2.xml": b'<w:hdr xmlns:w="%s"/>' % W.encode(),
        }
    )
    current = Package(
        {
            "[Content_Types].xml": EMPTY_CONTENT_TYPES,
            "word/_rels/document.xml.rels": etree.tostring(current_rels),
        }
    )

    carry, _save = _base_relationships(base, current)
    carried = carry(paragraph)

    refs = carried.findall(".//w:headerReference", {"w": W}) + carried.findall(
        ".//w:footerReference", {"w": W}
    )
    assert refs == [], f"expected no header/footer refs, got {refs!r}"
    assert (
        current.xml("word/_rels/document.xml.rels").findall(
            "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
        )
        == []
    )


def test_strip_incomplete_header_footer_refs_removes_missing_rid() -> None:
    root = etree.fromstring(
        f"""
        <w:document xmlns:w="{W}" xmlns:r="{R}">
          <w:body>
            <w:sectPr>
              <w:headerReference w:type="even"/>
              <w:headerReference w:type="default" r:id="rId1"/>
              <w:footerReference w:type="primary"/>
            </w:sectPr>
          </w:body>
        </w:document>
        """
    )
    _strip_incomplete_header_footer_refs(root)
    headers = root.findall(".//w:headerReference", {"w": W})
    footers = root.findall(".//w:footerReference", {"w": W})
    assert len(headers) == 1
    assert headers[0].get(f"{{{R}}}id") == "rId1"
    assert footers == []
