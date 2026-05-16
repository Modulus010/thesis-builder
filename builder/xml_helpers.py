#!/usr/bin/env python3
"""Word OOXML helper functions for thesis builder."""

import copy
from docx.oxml import OxmlElement
from docx.oxml.ns import qn


def ensure_trPr(row):
    tr = row._tr
    trPr = tr.find(qn('w:trPr'))
    if trPr is None:
        trPr = OxmlElement('w:trPr')
        tr.insert(0, trPr)
    return trPr


def set_row_cant_split(row):
    ensure_trPr(row).append(OxmlElement('w:cantSplit'))


def set_row_table_header(row):
    ensure_trPr(row).append(OxmlElement('w:tblHeader'))


def set_spacing_lines(para, before_lines: int, after_lines: int):
    pPr = para._element.get_or_add_pPr()
    spacing = pPr.find(qn('w:spacing'))
    if spacing is None:
        spacing = OxmlElement('w:spacing')
        pPr.append(spacing)
    for attr in (qn('w:before'), qn('w:after'), qn('w:beforeAutospacing'), qn('w:afterAutospacing')):
        if spacing.get(attr) is not None:
            del spacing.attrib[attr]
    spacing.set(qn('w:beforeLines'), str(before_lines))
    spacing.set(qn('w:afterLines'), str(after_lines))


_THEME_FONT_ATTRS = (
    qn('w:asciiTheme'), qn('w:eastAsiaTheme'),
    qn('w:hAnsiTheme'), qn('w:cstheme'),
)


def remove_theme_fonts(rFonts):
    for attr in _THEME_FONT_ATTRS:
        if rFonts.get(attr) is not None:
            del rFonts.attrib[attr]


def make_border(tag, val="single", sz="12"):
    el = OxmlElement(f"w:{tag}")
    el.set(qn("w:val"), val)
    el.set(qn("w:sz"), sz)
    el.set(qn("w:space"), "0")
    el.set(qn("w:color"), "auto")
    return el


def find_drawing_and_inline(run_element):
    drawing_elem = run_element.find(qn('w:drawing'))
    if drawing_elem is None:
        return None, None
    inline = drawing_elem.find(qn('wp:inline'))
    return drawing_elem, inline


def build_anchor_from_inline(inline, *, behind_doc=True, allow_overlap=True,
                              include_effect_extent=True):
    anchor = OxmlElement('wp:anchor')
    anchor.set('behindDoc', '1' if behind_doc else '0')
    anchor.set('distT', '0')
    anchor.set('distB', '0')
    anchor.set('distL', '0')
    anchor.set('distR', '0')
    anchor.set('simplePos', '0')
    anchor.set('locked', '0')
    anchor.set('layoutInCell', '1')
    anchor.set('allowOverlap', '1' if allow_overlap else '0')

    simplePos = OxmlElement('wp:simplePos')
    simplePos.set('x', '0')
    simplePos.set('y', '0')
    anchor.append(simplePos)

    posH = OxmlElement('wp:positionH')
    posH.set('relativeFrom', 'page')
    posOffH = OxmlElement('wp:posOffset')
    posOffH.text = '0'
    posH.append(posOffH)
    anchor.append(posH)

    posV = OxmlElement('wp:positionV')
    posV.set('relativeFrom', 'page')
    posOffV = OxmlElement('wp:posOffset')
    posOffV.text = '0'
    posV.append(posOffV)
    anchor.append(posV)

    extent = inline.find(qn('wp:extent'))
    if extent is not None:
        anchor.append(copy.deepcopy(extent))
    if include_effect_extent:
        effectExt = inline.find(qn('wp:effectExtent'))
        if effectExt is not None:
            anchor.append(copy.deepcopy(effectExt))

    anchor.append(OxmlElement('wp:wrapNone'))

    docPr = inline.find(qn('wp:docPr'))
    if docPr is not None:
        anchor.append(copy.deepcopy(docPr))
    graphic = inline.find(qn('a:graphic'))
    if graphic is not None:
        anchor.append(copy.deepcopy(graphic))

    return anchor
