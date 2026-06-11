"""
Conservation Report Professional Formatter
===========================================
Reformats existing Word (.docx) documents to a publication-quality standard
suitable for donor reports, technical conservation reports, and project proposals.
"""

import sys
import os
import re
import shutil
import argparse
import zipfile
from pathlib import Path
from lxml import etree

# ─────────────────────────────────────────────────────────────────────────────
# THEME DEFINITIONS
# ─────────────────────────────────────────────────────────────────────────────
THEMES = {
    "forest": {
        "name": "Forest Green",
        "font": "Arial",
        "base_size": "22",
        "primary":       "1E4D2B",
        "accent":        "3D6B4F",
        "accent_light":  "6B9E7A",
        "cover_sub":     "2B5F45",
        "table_header":  "D5E8D4",
        "table_alt":     "F2F8F0",
        "callout_bg":    "1E4D2B",
        "callout_text":  "FFFFFF",
        "caption":       "3D6B4F",
        "footer_rule":   "3D6B4F",
        "body":          "1A1A1A",
        "italic_quote":  "404040",
    },
    "ocean": {
        "name": "Ocean Teal",
        "font": "Arial",
        "base_size": "22",
        "primary":       "0D3D56",
        "accent":        "1A6B8A",
        "accent_light":  "2E9BBF",
        "cover_sub":     "155B75",
        "table_header":  "C8E6F0",
        "table_alt":     "F0F8FC",
        "callout_bg":    "0D3D56",
        "callout_text":  "FFFFFF",
        "caption":       "1A6B8A",
        "footer_rule":   "1A6B8A",
        "body":          "1A1A1A",
        "italic_quote":  "404040",
    },
    "earth": {
        "name": "Earth Ochre",
        "font": "Arial",
        "base_size": "22",
        "primary":       "5C3A1E",
        "accent":        "8B5E3C",
        "accent_light":  "B8845A",
        "cover_sub":     "7A4F2E",
        "table_header":  "F0E0C8",
        "table_alt":     "FBF5EF",
        "callout_bg":    "5C3A1E",
        "callout_text":  "FFFFFF",
        "caption":       "8B5E3C",
        "footer_rule":   "8B5E3C",
        "body":          "1A1A1A",
        "italic_quote":  "404040",
    },
    "slate": {
        "name": "Institutional Slate",
        "font": "Arial",
        "base_size": "22",
        "primary":       "1C2B3A",
        "accent":        "2E5077",
        "accent_light":  "4A7FAD",
        "cover_sub":     "243447",
        "table_header":  "D0DCE8",
        "table_alt":     "F2F5F8",
        "callout_bg":    "1C2B3A",
        "callout_text":  "FFFFFF",
        "caption":       "2E5077",
        "footer_rule":   "2E5077",
        "body":          "1A1A1A",
        "italic_quote":  "404040",
    },
}

# ─────────────────────────────────────────────────────────────────────────────
# XML NAMESPACES
# ─────────────────────────────────────────────────────────────────────────────
NS = {
    "w":   "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r":   "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp":  "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a":   "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "mc":  "http://schemas.openxmlformats.org/markup-compatibility/2006",
}

def W(tag):
    return f"{{{NS['w']}}}{tag}"

def xml_el(tag, attribs=None, parent=None):
    el = etree.SubElement(parent, W(tag)) if parent is not None else etree.Element(W(tag))
    if attribs:
        for k, v in attribs.items():
            el.set(W(k), v)
    return el

# ─────────────────────────────────────────────────────────────────────────────
# STYLE DEFINITIONS
# ─────────────────────────────────────────────────────────────────────────────
def build_styles_xml(theme):
    T = THEMES[theme]
    xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
          mc:Ignorable="w14 w15"
          xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
          xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">

  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="{T['font']}" w:hAnsi="{T['font']}" w:cs="{T['font']}"/>
        <w:sz w:val="{T['base_size']}"/>
        <w:szCs w:val="{T['base_size']}"/>
        <w:color w:val="{T['body']}"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="120" w:line="276" w:lineRule="auto"/>
        <w:jc w:val="both"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>

  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:after="120" w:line="276" w:lineRule="auto"/>
      <w:jc w:val="both"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{T['font']}" w:hAnsi="{T['font']}" w:cs="{T['font']}"/>
      <w:sz w:val="{T['base_size']}"/>
      <w:szCs w:val="{T['base_size']}"/>
      <w:color w:val="{T['body']}"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="CoverTitle">
    <w:name w:val="CoverTitle"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:shd w:val="clear" w:color="auto" w:fill="{T['primary']}"/>
      <w:spacing w:before="80" w:after="80" w:line="240" w:lineRule="auto"/>
      <w:jc w:val="left"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{T['font']}" w:hAnsi="{T['font']}" w:cs="{T['font']}"/>
      <w:b/>
      <w:bCs/>
      <w:color w:val="FFFFFF"/>
      <w:sz w:val="40"/>
      <w:szCs w:val="40"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="CoverSubtitle">
    <w:name w:val="CoverSubtitle"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:shd w:val="clear" w:color="auto" w:fill="{T['cover_sub']}"/>
      <w:spacing w:before="0" w:after="0" w:line="276" w:lineRule="auto"/>
      <w:jc w:val="left"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{T['font']}" w:hAnsi="{T['font']}" w:cs="{T['font']}"/>
      <w:color w:val="FFFFFF"/>
      <w:sz w:val="{T['base_size']}"/>
      <w:szCs w:val="{T['base_size']}"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:pBdr>
        <w:bottom w:val="single" w:sz="8" w:space="4" w:color="{T['accent']}"/>
      </w:pBdr>
      <w:spacing w:before="360" w:after="100"/>
      <w:jc w:val="left"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{T['font']}" w:hAnsi="{T['font']}" w:cs="{T['font']}"/>
      <w:b/>
      <w:color w:val="{T['accent']}"/>
      <w:sz w:val="28"/>
      <w:szCs w:val="28"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:spacing w:before="240" w:after="80"/>
      <w:jc w:val="left"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{T['font']}" w:hAnsi="{T['font']}" w:cs="{T['font']}"/>
      <w:b/>
      <w:color w:val="{T['accent_light']}"/>
      <w:sz w:val="24"/>
      <w:szCs w:val="24"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="Caption">
    <w:name w:val="caption"/>
    <w:basedOn w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="60" w:after="160"/>
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{T['font']}" w:hAnsi="{T['font']}" w:cs="{T['font']}"/>
      <w:b/><w:i/>
      <w:color w:val="{T['caption']}"/>
      <w:sz w:val="18"/>
      <w:szCs w:val="18"/>
    </w:rPr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="PullQuote">
    <w:name w:val="PullQuote"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:pBdr>
        <w:left w:val="single" w:sz="24" w:space="12" w:color="{T['accent']}"/>
      </w:pBdr>
      <w:spacing w:before="160" w:after="160"/>
      <w:ind w:left="360"/>
      <w:jc w:val="both"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{T['font']}" w:hAnsi="{T['font']}" w:cs="{T['font']}"/>
      <w:i/>
      <w:color w:val="{T['italic_quote']}"/>
      <w:sz w:val="{T['base_size']}"/>
      <w:szCs w:val="{T['base_size']}"/>
    </w:rPr>
  </w:style>

  <w:style w:type="table" w:styleId="ConservationTable">
    <w:name w:val="ConservationTable"/>
    <w:uiPriority w:val="40"/>
    <w:pPr>
      <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:tblPr>
      <w:tblBorders>
        <w:top    w:val="single" w:sz="8" w:space="0" w:color="{T['accent']}"/>
        <w:left   w:val="none"   w:sz="0" w:space="0" w:color="auto"/>
        <w:bottom w:val="single" w:sz="8" w:space="0" w:color="{T['accent']}"/>
        <w:right  w:val="none"   w:sz="0" w:space="0" w:color="auto"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="CCCCCC"/>
        <w:insideV w:val="none"   w:sz="0" w:space="0" w:color="auto"/>
      </w:tblBorders>
      <w:tblCellMar>
        <w:top w:w="100" w:type="dxa"/><w:left w:w="140" w:type="dxa"/>
        <w:bottom w:w="100" w:type="dxa"/><w:right w:w="140" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
    <w:tblStylePr w:type="firstRow">
      <w:rPr>
        <w:b/><w:color w:val="{T['body']}"/>
      </w:rPr>
      <w:tcPr>
        <w:shd w:val="clear" w:color="auto" w:fill="{T['table_header']}"/>
      </w:tcPr>
    </w:tblStylePr>
    <w:tblStylePr w:type="band1Horz">
      <w:tcPr>
        <w:shd w:val="clear" w:color="auto" w:fill="{T['table_alt']}"/>
      </w:tcPr>
    </w:tblStylePr>
  </w:style>

  <w:style w:type="paragraph" w:styleId="Footer">
    <w:name w:val="footer"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs>
      <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="{T['font']}" w:hAnsi="{T['font']}" w:cs="{T['font']}"/>
      <w:color w:val="555555"/>
      <w:sz w:val="16"/>
      <w:szCs w:val="16"/>
    </w:rPr>
  </w:style>
</w:styles>
"""
    return xml

def build_footer_xml(theme, report_label="Report"):
    T = THEMES[theme]
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Footer"/>
      <w:pBdr>
        <w:top w:val="single" w:sz="6" w:space="4" w:color="{T['footer_rule']}"/>
      </w:pBdr>
      <w:tabs><w:tab w:val="right" w:pos="9360"/></w:tabs>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="{T['font']}" w:hAnsi="{T['font']}" w:cs="{T['font']}"/>
        <w:color w:val="555555"/>
        <w:sz w:val="16"/><w:szCs w:val="16"/>
      </w:rPr>
      <w:t xml:space="preserve">| {report_label}  </w:t>
    </w:r>
    <w:r>
      <w:rPr>
         <w:rFonts w:ascii="{T['font']}" w:hAnsi="{T['font']}" w:cs="{T['font']}"/>
         <w:color w:val="555555"/>
         <w:sz w:val="16"/><w:szCs w:val="16"/>
      </w:rPr>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
      <w:rPr>
         <w:rFonts w:ascii="{T['font']}" w:hAnsi="{T['font']}" w:cs="{T['font']}"/>
         <w:color w:val="555555"/>
         <w:sz w:val="16"/><w:szCs w:val="16"/>
      </w:rPr>
      <w:instrText xml:space="preserve"> PAGE \\* MERGEFORMAT </w:instrText>
    </w:r>
    <w:r>
      <w:rPr>
         <w:rFonts w:ascii="{T['font']}" w:hAnsi="{T['font']}" w:cs="{T['font']}"/>
         <w:color w:val="555555"/>
         <w:sz w:val="16"/><w:szCs w:val="16"/>
      </w:rPr>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
  </w:p>
</w:ftr>
"""

# ─────────────────────────────────────────────────────────────────────────────
# CLASSIFICATION & HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────────────────────
def classify_paragraph(pStyle, text, idx, total):
    text_clean = text.strip()
    text_lower = text_clean.lower()
    n_words = len(text_clean.split())

    if pStyle in ("Heading1", "heading 1"): return ("Heading1", {})
    if pStyle in ("Heading2", "heading 2"): return ("Heading2", {})
    if pStyle in ("Title",): return ("CoverTitle", {})
    if pStyle in ("Subtitle",): return ("CoverSubtitle", {})
    if pStyle in ("Caption", "caption"): return ("Caption", {})
    if re.match(r'^(fig(ure)?|table|photo|map|chart|plate)\s*[\d\.IVX]+', text_clean, re.I): return ("Caption", {})
    if pStyle in ("Quote", "IntenseQuote", "Intense Quote"): return ("PullQuote", {})

    callout_triggers = ("management recommendation", "key recommendation", "recommendation", "key concern", "key finding", "action required", "proposed action")
    if any(text_lower.startswith(t) for t in callout_triggers):
        return ("_CALLOUT_HEADER", {})

    return ("Normal", {})

PPR_ORDER = ["pStyle","numPr","keepNext","keepLines","pageBreakBefore","framePr","suppressLineNumbers","pBdr","shd","tabs","suppressAutoHyphens","kinsoku","wordWrap","adjustRightInd","snapToGrid","spacing","ind","contextualSpacing","mirrorIndents","suppressOverlap","jc","textDirection","textAlignment","textboxTightWrap","outlineLvl","divId","cnfStyle","rPr","sectPr","pPrChange"]

def _reorder_pPr(pPr):
    children = list(pPr)
    if not children: return
    def sort_key(el):
        local = el.tag.split("}")[-1] if "}" in el.tag else el.tag
        try: return PPR_ORDER.index(local)
        except ValueError: return 999
    children.sort(key=sort_key)
    for child in children: pPr.remove(child)
    for child in children: pPr.append(child)

def _insert_before_rPr(pPr, el):
    ns = NS["w"]
    rPr = pPr.find(f"{{{ns}}}rPr")
    if rPr is not None:
        idx = list(pPr).index(rPr)
        pPr.insert(idx, el)
    else:
        pPr.append(el)

def get_pPr(p):
    ns = NS["w"]
    pPr = p.find(f"{{{ns}}}pPr")
    if pPr is None:
        pPr = etree.SubElement(p, f"{{{ns}}}pPr")
        p.insert(0, pPr)
    return pPr

def set_pStyle(pPr, style_id):
    ns = NS["w"]
    ps = pPr.find(f"{{{ns}}}pStyle")
    if ps is None:
        ps = etree.Element(f"{{{ns}}}pStyle")
        pPr.insert(0, ps)
    ps.set(f"{{{ns}}}val", style_id)

def set_shd_pPr(pPr, fill, val="clear", color="auto"):
    ns = NS["w"]
    shd = pPr.find(f"{{{ns}}}shd")
    if shd is None:
        shd = etree.Element(f"{{{ns}}}shd")
        _insert_before_rPr(pPr, shd)
    shd.set(f"{{{ns}}}val", val)
    shd.set(f"{{{ns}}}color", color)
    shd.set(f"{{{ns}}}fill", fill)
    _reorder_pPr(pPr)

def set_spacing(pPr, before=None, after=None, line=None, lineRule="auto"):
    ns = NS["w"]
    sp = pPr.find(f"{{{ns}}}spacing")
    if sp is None:
        sp = etree.Element(f"{{{ns}}}spacing")
        _insert_before_rPr(pPr, sp)
    if before is not None: sp.set(f"{{{ns}}}before", str(before))
    if after is not None: sp.set(f"{{{ns}}}after", str(after))
    if line is not None:
        sp.set(f"{{{ns}}}line", str(line))
        sp.set(f"{{{ns}}}lineRule", lineRule)
    _reorder_pPr(pPr)

def ensure_jc(pPr, val="both"):
    ns = NS["w"]
    jc = pPr.find(f"{{{ns}}}jc")
    if jc is None:
        jc = etree.Element(f"{{{ns}}}jc")
        _insert_before_rPr(pPr, jc)
    jc.set(f"{{{ns}}}val", val)
    _reorder_pPr(pPr)

def set_pBdr_bottom(pPr, color, sz=8, space=4):
    ns = NS["w"]
    pBdr = pPr.find(f"{{{ns}}}pBdr")
    if pBdr is None:
        pBdr = etree.Element(f"{{{ns}}}pBdr")
        _insert_before_rPr(pPr, pBdr)
    bottom = pBdr.find(f"{{{ns}}}bottom")
    if bottom is None:
        bottom = etree.SubElement(pBdr, f"{{{ns}}}bottom")
    bottom.set(f"{{{ns}}}val", "single")
    bottom.set(f"{{{ns}}}sz", str(sz))
    bottom.set(f"{{{ns}}}space", str(space))
    bottom.set(f"{{{ns}}}color", color)
    _reorder_pPr(pPr)

def set_pBdr_left(pPr, color, sz=24, space=12):
    ns = NS["w"]
    pBdr = pPr.find(f"{{{ns}}}pBdr")
    if pBdr is None:
        pBdr = etree.Element(f"{{{ns}}}pBdr")
        _insert_before_rPr(pPr, pBdr)
    left = pBdr.find(f"{{{ns}}}left")
    if left is None:
        left = etree.SubElement(pBdr, f"{{{ns}}}left")
    left.set(f"{{{ns}}}val", "single")
    left.set(f"{{{ns}}}sz", str(sz))
    left.set(f"{{{ns}}}space", str(space))
    left.set(f"{{{ns}}}color", color)
    _reorder_pPr(pPr)

def get_para_text(p):
    ns = NS["w"]
    return "".join([t.text or "" for t in p.iter(f"{{{ns}}}t")])

def get_para_style(p):
    ns = NS["w"]
    pPr = p.find(f"{{{ns}}}pPr")
    if pPr is None: return ""
    ps = pPr.find(f"{{{ns}}}pStyle")
    if ps is None: return ""
    return ps.get(f"{{{ns}}}val", "")

def set_run_color(r, color):
    ns = NS["w"]
    rPr = r.find(f"{{{ns}}}rPr")
    if rPr is None:
        rPr = etree.SubElement(r, f"{{{ns}}}rPr")
        r.insert(0, rPr)
    col = rPr.find(f"{{{ns}}}color")
    if col is None:
        col = etree.SubElement(rPr, f"{{{ns}}}color")
    col.set(f"{{{ns}}}val", color)

def set_run_font(r, font="Arial"):
    ns = NS["w"]
    rPr = r.find(f"{{{ns}}}rPr")
    if rPr is None:
        rPr = etree.SubElement(r, f"{{{ns}}}rPr")
        r.insert(0, rPr)
    rf = rPr.find(f"{{{ns}}}rFonts")
    if rf is None:
        rf = etree.SubElement(rPr, f"{{{ns}}}rFonts")
    rf.set(f"{{{ns}}}ascii", font)
    rf.set(f"{{{ns}}}hAnsi", font)
    rf.set(f"{{{ns}}}cs", font)

def format_table(tbl, theme):
    T = THEMES[theme]
    ns = NS["w"]
    tblPr = tbl.find(f"{{{ns}}}tblPr")
    if tblPr is None:
        tblPr = etree.SubElement(tbl, f"{{{ns}}}tblPr")
        tbl.insert(0, tblPr)

    tblStyle = tblPr.find(f"{{{ns}}}tblStyle")
    if tblStyle is None:
        tblStyle = etree.SubElement(tblPr, f"{{{ns}}}tblStyle")
    tblStyle.set(f"{{{ns}}}val", "TableGrid")

    tblBorders = tblPr.find(f"{{{ns}}}tblBorders")
    if tblBorders is None:
        tblBorders = etree.SubElement(tblPr, f"{{{ns}}}tblBorders")

    def set_border(parent, side, val, sz, color):
        el = parent.find(f"{{{ns}}}{side}")
        if el is None: el = etree.SubElement(parent, f"{{{ns}}}{side}")
        el.set(f"{{{ns}}}val", val)
        el.set(f"{{{ns}}}sz", str(sz))
        el.set(f"{{{ns}}}space", "0")
        el.set(f"{{{ns}}}color", color)

    set_border(tblBorders, "top", "single", 8, T["accent"])
    set_border(tblBorders, "left", "none", 0, "auto")
    set_border(tblBorders, "bottom", "single", 8, T["accent"])
    set_border(tblBorders, "right", "none", 0, "auto")
    set_border(tblBorders, "insideH", "single", 4, "CCCCCC")
    set_border(tblBorders, "insideV", "none", 0, "auto")

    rows = tbl.findall(f"{{{ns}}}tr")
    for r_idx, tr in enumerate(rows):
        cells = tr.findall(f"{{{ns}}}tc")
        is_callout = False
        if len(cells) == 1:
            row_text = get_para_text(cells[0]).lower()
            if any(kw in row_text for kw in ["management recommendation", "recommendation", "key finding", "key concern"]):
                is_callout = True

        for tc in cells:
            tcPr = tc.find(f"{{{ns}}}tcPr")
            if tcPr is None:
                tcPr = etree.SubElement(tc, f"{{{ns}}}tcPr")
                tc.insert(0, tcPr)

            shd = tcPr.find(f"{{{ns}}}shd")
            if shd is None:
                shd = etree.SubElement(tcPr, f"{{{ns}}}shd")

            if is_callout:
                shd.set(f"{{{ns}}}val", "clear")
                shd.set(f"{{{ns}}}color", "auto")
                shd.set(f"{{{ns}}}fill", T["callout_bg"])
                for p in tc.findall(f"{{{ns}}}p"):
                    for run in p.findall(f"{{{ns}}}r"):
                        set_run_color(run, T["callout_text"])
                        set_run_font(run, T["font"])
            elif r_idx == 0:
                shd.set(f"{{{ns}}}val", "clear")
                shd.set(f"{{{ns}}}color", "auto")
                shd.set(f"{{{ns}}}fill", T["table_header"])
                for p in tc.findall(f"{{{ns}}}p"):
                    for run in p.findall(f"{{{ns}}}r"):
                        rPr = run.find(f"{{{ns}}}rPr")
                        if rPr is None:
                            rPr = etree.SubElement(run, f"{{{ns}}}rPr")
                            run.insert(0, rPr)
                        if rPr.find(f"{{{ns}}}b") is None: etree.SubElement(rPr, f"{{{ns}}}b")
                        set_run_font(run, T["font"])
            elif r_idx % 2 == 0:
                shd.set(f"{{{ns}}}val", "clear")
                shd.set(f"{{{ns}}}color", "auto")
                shd.set(f"{{{ns}}}fill", T["table_alt"])
                for p in tc.findall(f"{{{ns}}}p"):
                    for run in p.findall(f"{{{ns}}}r"): set_run_font(run, T["font"])
            else:
                shd.set(f"{{{ns}}}val", "clear")
                shd.set(f"{{{ns}}}color", "auto")
                shd.set(f"{{{ns}}}fill", "FFFFFF")
                for p in tc.findall(f"{{{ns}}}p"):
                    for run in p.findall(f"{{{ns}}}r"): set_run_font(run, T["font"])

def format_paragraph(p, idx, total, theme, section_counter):
    T = THEMES[theme]
    ns = NS["w"]
    text = get_para_text(p)
    pStyle = get_para_style(p)
    new_style, extra = classify_paragraph(pStyle, text, idx, total)
    pPr = get_pPr(p)

    if new_style != "_CALLOUT_HEADER":
        set_pStyle(pPr, new_style)

    if new_style == "Heading1":
        section_counter[0] += 1
        set_spacing(pPr, before=360, after=100, line=240, lineRule="auto")
        ensure_jc(pPr, "left")
        set_pBdr_bottom(pPr, T["accent"], sz=8, space=4)
        for r in p.findall(f"{{{ns}}}r"):
            set_run_color(r, T["accent"])
            set_run_font(r, T["font"])
    elif new_style == "Heading2":
        set_spacing(pPr, before=240, after=80, line=240, lineRule="auto")
        ensure_jc(pPr, "left")
        for r in p.findall(f"{{{ns}}}r"):
            set_run_color(r, T["accent_light"])
            set_run_font(r, T["font"])
    elif new_style in ("CoverTitle", "CoverSubtitle", "CoverLabel"):
        fill = T["primary"] if new_style in ("CoverTitle", "CoverLabel") else T["cover_sub"]
        set_shd_pPr(pPr, fill)
        for r in p.findall(f"{{{ns}}}r"):
            set_run_color(r, "FFFFFF")
            set_run_font(r, T["font"])
        ensure_jc(pPr, "left")
    elif new_style == "Caption":
        set_spacing(pPr, before=60, after=160, line=240, lineRule="auto")
        ensure_jc(pPr, "center")
        for r in p.findall(f"{{{ns}}}r"):
            set_run_color(r, T["caption"])
            set_run_font(r, T["font"])
    elif new_style == "PullQuote":
        set_spacing(pPr, before=160, after=160, line=276, lineRule="auto")
        set_pBdr_left(pPr, T["accent"])
        ensure_jc(pPr, "both")
        for r in p.findall(f"{{{ns}}}r"):
            set_run_color(r, T["italic_quote"])
            set_run_font(r, T["font"])
    elif new_style == "_CALLOUT_HEADER":
        set_pStyle(pPr, "Normal")
        set_shd_pPr(pPr, T["callout_bg"])
        set_spacing(pPr, before=120, after=60, line=240, lineRule="auto")
        ensure_jc(pPr, "left")
        for r in p.findall(f"{{{ns}}}r"):
            set_run_color(r, T["callout_text"])
            set_run_font(r, T["font"])
    else:
        set_spacing(pPr, before=0, after=120, line=276, lineRule="auto")
        ensure_jc(pPr, "both")
        for r in p.findall(f"{{{ns}}}r"):
            set_run_font(r, T["font"])
            rPr = r.find(f"{{{ns}}}rPr")
            if rPr is not None:
                col = rPr.find(f"{{{ns}}}color")
                if col is None:
                    col = etree.SubElement(rPr, f"{{{ns}}}color")
                    col.set(f"{{{ns}}}val", T["body"])

    for r in p.findall(f"{{{ns}}}r"):
        set_run_font(r, T["font"])

    return section_counter

def format_document(input_path, output_path, theme="forest", report_label=None):
    T = THEMES[theme]
    ns = NS["w"]
    ns_r = NS["r"]
    if report_label is None:
        report_label = Path(input_path).stem.replace("_", " ")

    work_dir = Path(output_path).parent / "_fmt_work"
    if work_dir.exists(): shutil.rmtree(work_dir)
    with zipfile.ZipFile(input_path, "r") as z: z.extractall(work_dir)

    doc_xml_path = work_dir / "word" / "document.xml"
    tree = etree.parse(str(doc_xml_path))
    root = tree.getroot()
    body = root.find(f"{{{ns}}}body")

    children = list(body)
    paragraphs = [c for c in children if c.tag == f"{{{ns}}}p"]
    tables     = [c for c in children if c.tag == f"{{{ns}}}tbl"]
    total_p    = len(paragraphs)

    section_counter = [0]
    for idx, p in enumerate(paragraphs):
        section_counter = format_paragraph(p, idx, total_p, theme, section_counter)

    for tbl in tables:
        format_table(tbl, theme)

    tree.write(str(doc_xml_path), xml_declaration=True, encoding="UTF-8", standalone=True)

    styles_path = work_dir / "word" / "styles.xml"
    with open(styles_path, "w", encoding="utf-8") as f:
        f.write(build_styles_xml(theme))

    footer_path = work_dir / "word" / "footer1.xml"
    with open(footer_path, "w", encoding="utf-8") as f:
        f.write(build_footer_xml(theme, report_label))

    _wire_footer(work_dir, ns, ns_r)

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for fpath in sorted(work_dir.rglob("*")):
            if fpath.is_file():
                arcname = fpath.relative_to(work_dir)
                zout.write(fpath, arcname)

    shutil.rmtree(work_dir)

def _wire_footer(work_dir, ns, ns_r):
    rels_path = work_dir / "word" / "_rels" / "document.xml.rels"
    if not rels_path.exists(): return
    rels_tree = etree.parse(str(rels_path))
    rels_root = rels_tree.getroot()
    
    footer_rel_id = None
    for rel in rels_root:
        if "footer1" in rel.get("Target", ""):
            footer_rel_id = rel.get("Id")
            break

    if footer_rel_id is None:
        existing_ids = [int(r.get("Id", "rId0").replace("rId", "0")) for r in rels_root if r.get("Id", "").startswith("rId")]
        new_id_num = max(existing_ids, default=0) + 1
        footer_rel_id = f"rId{new_id_num}"
        new_rel = etree.SubElement(rels_root, "Relationship")
        new_rel.set("Id", footer_rel_id)
        new_rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer")
        new_rel.set("Target", "footer1.xml")
        rels_tree.write(str(rels_path), xml_declaration=True, encoding="UTF-8", standalone=True)

    doc_xml_path = work_dir / "word" / "document.xml"
    doc_tree = etree.parse(str(doc_xml_path))
    doc_root = doc_tree.getroot()
    body = doc_root.find(f"{{{ns}}}body")
    if body is None: return

    sectPr = body.find(f"{{{ns}}}sectPr")
    if sectPr is None: sectPr = etree.SubElement(body, f"{{{ns}}}sectPr")

    has_footer_ref = False
    for child in sectPr:
        if child.tag == f"{{{ns}}}footerReference" and child.get(f"{{{ns}}}type") == "default":
            child.set(f"{{{ns_r}}}id", footer_rel_id)
            has_footer_ref = True
            break

    if not has_footer_ref:
        footer_ref = etree.SubElement(sectPr, f"{{{ns}}}footerReference")
        footer_ref.set(f"{{{ns}}}type", "default")
        footer_ref.set(f"{{{ns_r}}}id", footer_rel_id)

    doc_tree.write(str(doc_xml_path), xml_declaration=True, encoding="UTF-8", standalone=True)

    ct_path = work_dir / "[Content_Types].xml"
    if ct_path.exists():
        ct_tree = etree.parse(str(ct_path))
        ct_root = ct_tree.getroot()
        ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        existing = [el.get("PartName", "") for el in ct_root]
        if "/word/footer1.xml" not in existing:
            new_ct = etree.SubElement(ct_root, f"{{{ct_ns}}}Override")
            new_ct.set("PartName", "/word/footer1.xml")
            new_ct.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
        ct_tree.write(str(ct_path), xml_declaration=True, encoding="UTF-8", standalone=True)

if __name__ == "__main__":
    pass
