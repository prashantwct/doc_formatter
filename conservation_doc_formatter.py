"""
Conservation Report Professional Formatter
===========================================
Reformats existing Word (.docx) documents to a publication-quality standard
suitable for donor reports, technical conservation reports, and project proposals.

Usage:
    python conservation_doc_formatter.py input.docx output.docx [--theme forest|ocean|earth|slate]
    python conservation_doc_formatter.py input.docx output.docx --theme forest --logo logo.png

Themes:
    forest  – Deep greens (default; suits tiger/elephant/forest projects)
    ocean   – Teal blues (marine conservation, river projects)
    earth   – Warm terracotta + brown (dryland, savanna, arid ecosystems)
    slate   – Neutral grey-navy (multi-donor, neutral institutional)

What it does:
    - Teal/green cover block with white title text
    - Numbered section headings with bottom-rule accent
    - Subheadings in accent colour
    - Body text: justified, 11pt Arial, 1.15 line spacing
    - Pull-quote / callout box style for key findings
    - Data tables: alternating row shading, bold headers, accent border
    - Management Recommendation / Key Concern boxes: dark teal bg, white text
    - Figure captions: italic, accent colour, centred
    - Footer: report label left, page number right, accent top border
    - Header: optional logo + document title
    - Consistent paragraph spacing throughout
"""

import sys
import os
import re
import copy
import shutil
import argparse
import zipfile
from pathlib import Path
from lxml import etree

# ─────────────────────────────────────────────────────────────────────────────
# THEME DEFINITIONS
# Each theme carries: primary (dark), accent (medium), light (header shading),
# table_header (cell bg), table_alt (alternating row), text_on_primary (always white),
# callout_bg, callout_text, body_text, caption_colour
# ─────────────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
# PAGE LAYOUT PROFILES
# All measurements in DXA (1440 DXA = 1 inch).
# "content_width" = page_width - margin_left - margin_right (pre-computed)
# ─────────────────────────────────────────────────────────────────────────────
PAGE_LAYOUTS = {
    "a4_standard": {
        # A4 with generous 1.18" side margins — suits donor / technical reports
        "page_width":     11906,
        "page_height":    16838,
        "margin_top":     1440,   # 1.0"
        "margin_bottom":  1440,   # 1.0"
        "margin_left":    1701,   # 1.18"
        "margin_right":   1701,
        "margin_header":   720,   # 0.5" from edge
        "margin_footer":   720,
        "content_width":  8504,   # 11906 - 2*1701
        "gutter":            0,
    },
    "a4_wide": {
        # A4 with tighter 0.75" margins — maximises content area
        "page_width":     11906,
        "page_height":    16838,
        "margin_top":     1080,
        "margin_bottom":  1080,
        "margin_left":    1080,
        "margin_right":   1080,
        "margin_header":   540,
        "margin_footer":   540,
        "content_width":  9746,
        "gutter":            0,
    },
    "us_letter": {
        # US Letter 8.5×11 — 1" margins all round
        "page_width":     12240,
        "page_height":    15840,
        "margin_top":     1440,
        "margin_bottom":  1440,
        "margin_left":    1440,
        "margin_right":   1440,
        "margin_header":   720,
        "margin_footer":   720,
        "content_width":  9360,   # 12240 - 2*1440
        "gutter":            0,
    },
}

# ─────────────────────────────────────────────────────────────────────────────
# INDENTATION PROFILES
# left/right/firstLine/hanging all in DXA; None = do not set.
# Used by both styles.xml generation and runtime paragraph patching.
# ─────────────────────────────────────────────────────────────────────────────
INDENT_PROFILES = {
    # Conservative: suits formal institutional reports
    "formal": {
        "body_first_line":    0,     # no first-line indent (block paragraphs)
        "body_left":          0,
        "body_right":         0,
        "pullquote_left":   504,     # 0.35"
        "pullquote_right":  504,
        "callout_left":     180,     # 0.125"
        "callout_right":    180,
        "bullet_left":      504,     # 0.35"
        "bullet_hanging":   252,
        "number_left":      504,
        "number_hanging":   252,
        "caption_left":       0,
        "caption_right":      0,
        "heading1_left":      0,
        "heading2_left":      0,
        "heading3_left":      0,
    },
    # Editorial: paragraph first-line indent, wider pull-quotes — suits
    # narrative-style annual / progress reports
    "editorial": {
        "body_first_line":  360,     # 0.25" first-line indent
        "body_left":          0,
        "body_right":         0,
        "pullquote_left":   720,     # 0.5"
        "pullquote_right":  720,
        "callout_left":     216,
        "callout_right":    216,
        "bullet_left":      576,
        "bullet_hanging":   288,
        "number_left":      576,
        "number_hanging":   288,
        "caption_left":       0,
        "caption_right":      0,
        "heading1_left":      0,
        "heading2_left":    288,     # subtle step-in for H2
        "heading3_left":    432,
    },
    # Compact: no first-line, minimal margins — suits field reports, data tables
    "compact": {
        "body_first_line":    0,
        "body_left":          0,
        "body_right":         0,
        "pullquote_left":   360,
        "pullquote_right":  360,
        "callout_left":     144,
        "callout_right":    144,
        "bullet_left":      432,
        "bullet_hanging":   216,
        "number_left":      432,
        "number_hanging":   216,
        "caption_left":       0,
        "caption_right":      0,
        "heading1_left":      0,
        "heading2_left":      0,
        "heading3_left":    216,
    },
}

THEMES = {
    "forest": {
        "name": "Forest Green",
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
        # ── Layout / indentation defaults for this theme ──────────────────
        "page_layout":   "a4_standard",
        "indent_profile": "formal",
    },
    "ocean": {
        "name": "Ocean Teal",
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
        "page_layout":   "a4_standard",
        "indent_profile": "editorial",
    },
    "earth": {
        "name": "Earth Ochre",
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
        "page_layout":   "a4_standard",
        "indent_profile": "editorial",
    },
    "slate": {
        "name": "Institutional Slate",
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
        "page_layout":   "us_letter",
        "indent_profile": "formal",
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
    """Create a w: namespace element."""
    el = etree.SubElement(parent, W(tag)) if parent is not None else etree.Element(W(tag))
    if attribs:
        for k, v in attribs.items():
            el.set(W(k), v)
    return el

# ─────────────────────────────────────────────────────────────────────────────
# STYLE DEFINITIONS
# ─────────────────────────────────────────────────────────────────────────────
def build_styles_xml(theme):
    """Return a complete styles.xml string for the given theme."""
    T = THEMES[theme]
    IP = INDENT_PROFILES[T.get("indent_profile", "formal")]

    # Helper to emit an <w:ind> element string (only non-zero/non-None attrs)
    def _ind(left=None, right=None, firstLine=None, hanging=None):
        attrs = []
        if left      is not None and left      != 0: attrs.append(f'w:left="{left}"')
        if right     is not None and right     != 0: attrs.append(f'w:right="{right}"')
        if firstLine is not None and firstLine != 0: attrs.append(f'w:firstLine="{firstLine}"')
        if hanging   is not None and hanging   != 0: attrs.append(f'w:hanging="{hanging}"')
        return f'<w:ind {" ".join(attrs)}/>' if attrs else ''

    body_ind   = _ind(left=IP["body_left"],    right=IP["body_right"],
                      firstLine=IP["body_first_line"])
    h1_ind     = _ind(left=IP["heading1_left"])
    h2_ind     = _ind(left=IP["heading2_left"])
    h3_ind     = _ind(left=IP["heading3_left"])
    pq_ind     = _ind(left=IP["pullquote_left"],  right=IP["pullquote_right"])
    co_ind     = _ind(left=IP["callout_left"],    right=IP["callout_right"])
    cap_ind    = _ind(left=IP["caption_left"],    right=IP["caption_right"])
    bul_ind    = _ind(left=IP["bullet_left"],     hanging=IP["bullet_hanging"])
    num_ind    = _ind(left=IP["number_left"],     hanging=IP["number_hanging"])

    xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
          mc:Ignorable="w14 w15"
          xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
          xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">

  <!-- ── Document defaults ── -->
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:color w:val="{T['body']}"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="120" w:line="276" w:lineRule="auto"/>
        <w:jc w:val="both"/>
        {body_ind}
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>

  <!-- ── Normal ── -->
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:after="120" w:line="276" w:lineRule="auto"/>
      <w:jc w:val="both"/>
      {body_ind}
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="22"/>
      <w:color w:val="{T['body']}"/>
    </w:rPr>
  </w:style>

  <!-- ── Cover Title ── -->
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
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:b/>
      <w:bCs/>
      <w:color w:val="FFFFFF"/>
      <w:sz w:val="40"/>
      <w:szCs w:val="40"/>
    </w:rPr>
  </w:style>

  <!-- ── Cover Subtitle / Meta ── -->
  <w:style w:type="paragraph" w:styleId="CoverSubtitle">
    <w:name w:val="CoverSubtitle"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:shd w:val="clear" w:color="auto" w:fill="{T['cover_sub']}"/>
      <w:spacing w:before="0" w:after="0" w:line="276" w:lineRule="auto"/>
      <w:jc w:val="left"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:color w:val="FFFFFF"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="22"/>
    </w:rPr>
  </w:style>

  <!-- ── Cover Label (e.g. "INTERIM REPORT") ── -->
  <w:style w:type="paragraph" w:styleId="CoverLabel">
    <w:name w:val="CoverLabel"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:shd w:val="clear" w:color="auto" w:fill="{T['primary']}"/>
      <w:spacing w:before="200" w:after="60" w:line="240" w:lineRule="auto"/>
      <w:jc w:val="left"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:b/>
      <w:bCs/>
      <w:color w:val="FFFFFF"/>
      <w:sz w:val="18"/>
      <w:szCs w:val="18"/>
      <w:spacing w:val="100"/>
    </w:rPr>
  </w:style>

  <!-- ── Heading 1: numbered section ── -->
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:link w:val="Heading1Char"/>
    <w:uiPriority w:val="9"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:pBdr>
        <w:bottom w:val="single" w:sz="8" w:space="4" w:color="{T['accent']}"/>
      </w:pBdr>
      <w:spacing w:before="360" w:after="100"/>
      <w:jc w:val="left"/>
      {h1_ind}
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:b/>
      <w:bCs/>
      <w:color w:val="{T['accent']}"/>
      <w:sz w:val="28"/>
      <w:szCs w:val="28"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:customStyle="1" w:styleId="Heading1Char">
    <w:name w:val="Heading 1 Char"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:link w:val="Heading1"/>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:b/>
      <w:bCs/>
      <w:color w:val="{T['accent']}"/>
      <w:sz w:val="28"/>
      <w:szCs w:val="28"/>
    </w:rPr>
  </w:style>

  <!-- ── Heading 2 ── -->
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:link w:val="Heading2Char"/>
    <w:uiPriority w:val="9"/>
    <w:unhideWhenUsed/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:spacing w:before="240" w:after="80"/>
      <w:jc w:val="left"/>
      {h2_ind}
      <w:outlineLvl w:val="1"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:b/>
      <w:bCs/>
      <w:color w:val="{T['accent_light']}"/>
      <w:sz w:val="24"/>
      <w:szCs w:val="24"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:customStyle="1" w:styleId="Heading2Char">
    <w:name w:val="Heading 2 Char"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:link w:val="Heading2"/>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:b/>
      <w:bCs/>
      <w:color w:val="{T['accent_light']}"/>
      <w:sz w:val="24"/>
      <w:szCs w:val="24"/>
    </w:rPr>
  </w:style>

  <!-- ── Heading 3 ── -->
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:uiPriority w:val="9"/>
    <w:unhideWhenUsed/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:spacing w:before="200" w:after="60"/>
      <w:jc w:val="left"/>
      {h3_ind}
      <w:outlineLvl w:val="2"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:b/>
      <w:bCs/>
      <w:i/>
      <w:iCs/>
      <w:color w:val="{T['accent_light']}"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="22"/>
    </w:rPr>
  </w:style>

  <!-- ── Caption ── -->
  <w:style w:type="paragraph" w:styleId="Caption">
    <w:name w:val="caption"/>
    <w:basedOn w:val="Normal"/>
    <w:uiPriority w:val="35"/>
    <w:unhideWhenUsed/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="60" w:after="160"/>
      <w:jc w:val="center"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:b/>
      <w:bCs/>
      <w:i/>
      <w:iCs/>
      <w:color w:val="{T['caption']}"/>
      <w:sz w:val="18"/>
      <w:szCs w:val="18"/>
    </w:rPr>
  </w:style>

  <!-- ── Pull Quote / Key Finding ── -->
  <w:style w:type="paragraph" w:styleId="PullQuote">
    <w:name w:val="PullQuote"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:pBdr>
        <w:left w:val="single" w:sz="24" w:space="12" w:color="{T['accent']}"/>
      </w:pBdr>
      <w:spacing w:before="160" w:after="160"/>
      {pq_ind}
      <w:jc w:val="both"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:i/>
      <w:iCs/>
      <w:color w:val="{T['italic_quote']}"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="22"/>
    </w:rPr>
  </w:style>

  <!-- ── Callout (Management Rec / Key Concern) ── -->
  <w:style w:type="paragraph" w:styleId="Callout">
    <w:name w:val="Callout"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr>
      <w:spacing w:before="120" w:after="60" w:line="240" w:lineRule="auto"/>
      {co_ind}
      <w:jc w:val="left"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:color w:val="{T['callout_text']}"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="22"/>
    </w:rPr>
  </w:style>

  <!-- ── List Bullet ── -->
  <w:style w:type="paragraph" w:styleId="ListBullet">
    <w:name w:val="List Bullet"/>
    <w:basedOn w:val="Normal"/>
    <w:uiPriority w:val="34"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="0" w:after="80" w:line="276" w:lineRule="auto"/>
      {bul_ind}
      <w:jc w:val="both"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:color w:val="{T['body']}"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="22"/>
    </w:rPr>
  </w:style>

  <!-- ── List Number ── -->
  <w:style w:type="paragraph" w:styleId="ListNumber">
    <w:name w:val="List Number"/>
    <w:basedOn w:val="Normal"/>
    <w:uiPriority w:val="34"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="0" w:after="80" w:line="276" w:lineRule="auto"/>
      {num_ind}
      <w:jc w:val="both"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:color w:val="{T['body']}"/>
      <w:sz w:val="22"/>
      <w:szCs w:val="22"/>
    </w:rPr>
  </w:style>

  <!-- ── Table Normal ── -->
  <w:style w:type="table" w:default="1" w:styleId="TableNormal">
    <w:name w:val="Normal Table"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:tblPr>
      <w:tblInd w:w="0" w:type="dxa"/>
      <w:tblCellMar>
        <w:top w:w="80" w:type="dxa"/>
        <w:left w:w="120" w:type="dxa"/>
        <w:bottom w:w="80" w:type="dxa"/>
        <w:right w:w="120" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>

  <!-- ── ConservationTable: professional styled table ── -->
  <w:style w:type="table" w:styleId="ConservationTable">
    <w:name w:val="ConservationTable"/>
    <w:basedOn w:val="TableNormal"/>
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
        <w:top    w:w="100" w:type="dxa"/>
        <w:left   w:w="140" w:type="dxa"/>
        <w:bottom w:w="100" w:type="dxa"/>
        <w:right  w:w="140" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
    <!-- First row: header shading + bold white -->
    <w:tblStylePr w:type="firstRow">
      <w:pPr><w:jc w:val="left"/></w:pPr>
      <w:rPr>
        <w:b/><w:bCs/>
        <w:color w:val="{T['body']}"/>
      </w:rPr>
      <w:tcPr>
        <w:shd w:val="clear" w:color="auto" w:fill="{T['table_header']}"/>
      </w:tcPr>
    </w:tblStylePr>
    <!-- Even rows: alternating light shading -->
    <w:tblStylePr w:type="band1Horz">
      <w:tcPr>
        <w:shd w:val="clear" w:color="auto" w:fill="{T['table_alt']}"/>
      </w:tcPr>
    </w:tblStylePr>
  </w:style>

  <!-- ── TableGrid (fallback) ── -->
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
    <w:basedOn w:val="TableNormal"/>
    <w:uiPriority w:val="39"/>
    <w:pPr>
      <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:tblPr>
      <w:tblBorders>
        <w:top    w:val="single" w:sz="4" w:space="0" w:color="BBBBBB"/>
        <w:left   w:val="single" w:sz="4" w:space="0" w:color="BBBBBB"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="BBBBBB"/>
        <w:right  w:val="single" w:sz="4" w:space="0" w:color="BBBBBB"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="BBBBBB"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="BBBBBB"/>
      </w:tblBorders>
      <w:tblCellMar>
        <w:top    w:w="80"  w:type="dxa"/>
        <w:left   w:w="120" w:type="dxa"/>
        <w:bottom w:w="80"  w:type="dxa"/>
        <w:right  w:w="120" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>

  <!-- ── Header style ── -->
  <w:style w:type="paragraph" w:styleId="Header">
    <w:name w:val="header"/>
    <w:basedOn w:val="Normal"/>
    <w:link w:val="HeaderChar"/>
    <w:uiPriority w:val="99"/>
    <w:unhideWhenUsed/>
    <w:pPr>
      <w:tabs>
        <w:tab w:val="center" w:pos="4680"/>
        <w:tab w:val="right"  w:pos="9360"/>
      </w:tabs>
      <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
      <w:jc w:val="left"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:color w:val="666666"/>
      <w:sz w:val="16"/>
      <w:szCs w:val="16"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:customStyle="1" w:styleId="HeaderChar">
    <w:name w:val="Header Char"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:link w:val="Header"/>
  </w:style>

  <!-- ── Footer style ── -->
  <w:style w:type="paragraph" w:styleId="Footer">
    <w:name w:val="footer"/>
    <w:basedOn w:val="Normal"/>
    <w:link w:val="FooterChar"/>
    <w:uiPriority w:val="99"/>
    <w:unhideWhenUsed/>
    <w:pPr>
      <w:tabs>
        <w:tab w:val="right" w:pos="9360"/>
      </w:tabs>
      <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
      <w:jc w:val="left"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
      <w:color w:val="555555"/>
      <w:sz w:val="16"/>
      <w:szCs w:val="16"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:customStyle="1" w:styleId="FooterChar">
    <w:name w:val="Footer Char"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:link w:val="Footer"/>
  </w:style>

  <!-- ── Default Paragraph Font ── -->
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
    <w:name w:val="Default Paragraph Font"/>
    <w:uiPriority w:val="1"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
  </w:style>

  <!-- ── No List ── -->
  <w:style w:type="numbering" w:default="1" w:styleId="NoList">
    <w:name w:val="No List"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
  </w:style>

</w:styles>
"""
    return xml


# ─────────────────────────────────────────────────────────────────────────────
# FOOTER XML
# ─────────────────────────────────────────────────────────────────────────────
def build_footer_xml(theme, report_label="Report"):
    T = THEMES[theme]
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:p>
    <w:pPr>
      <w:pStyle w:val="Footer"/>
      <w:pBdr>
        <w:top w:val="single" w:sz="6" w:space="4" w:color="{T['footer_rule']}"/>
      </w:pBdr>
      <w:tabs>
        <w:tab w:val="right" w:pos="9360"/>
      </w:tabs>
    </w:pPr>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:color w:val="555555"/>
        <w:sz w:val="16"/><w:szCs w:val="16"/>
      </w:rPr>
      <w:t xml:space="preserve">| {report_label}</w:t>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:color w:val="555555"/>
        <w:sz w:val="16"/><w:szCs w:val="16"/>
      </w:rPr>
      <w:tab/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:color w:val="555555"/>
        <w:sz w:val="16"/><w:szCs w:val="16"/>
      </w:rPr>
      <w:t xml:space="preserve">Page </w:t>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:color w:val="555555"/>
        <w:sz w:val="16"/><w:szCs w:val="16"/>
      </w:rPr>
      <w:fldChar w:fldCharType="begin"/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:sz w:val="16"/><w:szCs w:val="16"/>
      </w:rPr>
      <w:instrText xml:space="preserve"> PAGE \\* MERGEFORMAT </w:instrText>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:sz w:val="16"/><w:szCs w:val="16"/>
      </w:rPr>
      <w:fldChar w:fldCharType="separate"/>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:noProof/>
        <w:sz w:val="16"/><w:szCs w:val="16"/>
      </w:rPr>
      <w:t>1</w:t>
    </w:r>
    <w:r>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:sz w:val="16"/><w:szCs w:val="16"/>
      </w:rPr>
      <w:fldChar w:fldCharType="end"/>
    </w:r>
  </w:p>
</w:ftr>
"""



# ─────────────────────────────────────────────────────────────────────────────
# NUMBERING XML  (bullet + numbered list definitions)
# ─────────────────────────────────────────────────────────────────────────────
def build_numbering_xml(theme):
    """Return a numbering.xml with bullet (abstractNumId=1) and numbered (abstractNumId=2)."""
    T = THEMES[theme]
    IP = INDENT_PROFILES[T.get("indent_profile", "formal")]
    bl = IP["bullet_left"]
    bh = IP["bullet_hanging"]
    nl = IP["number_left"]
    nh = IP["number_hanging"]
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">

  <!-- ── Bullet list ── -->
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#x2022;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="{bl}" w:hanging="{bh}"/>
        <w:spacing w:after="60"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:cs="Arial"/>
        <w:sz w:val="20"/>
      </w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#x25E6;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="{bl + bh}" w:hanging="{bh}"/>
        <w:spacing w:after="40"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:cs="Arial"/>
        <w:sz w:val="20"/>
      </w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="2">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="&#x25AA;"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="{bl + 2*bh}" w:hanging="{bh}"/>
        <w:spacing w:after="40"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:cs="Arial"/>
        <w:sz w:val="20"/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>

  <!-- ── Numbered list ── -->
  <w:abstractNum w:abstractNumId="2">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="{nl}" w:hanging="{nh}"/>
        <w:spacing w:after="60"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:sz w:val="22"/>
      </w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%2)"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
        <w:ind w:left="{nl + nh}" w:hanging="{nh}"/>
        <w:spacing w:after="40"/>
      </w:pPr>
      <w:rPr>
        <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/>
        <w:sz w:val="22"/>
      </w:rPr>
    </w:lvl>
  </w:abstractNum>

  <w:num w:numId="1">
    <w:abstractNumId w:val="1"/>
  </w:num>
  <w:num w:numId="2">
    <w:abstractNumId w:val="2"/>
  </w:num>

</w:numbering>
"""


# ─────────────────────────────────────────────────────────────────────────────
# SETTINGS XML  (page layout embedded in document settings)
# ─────────────────────────────────────────────────────────────────────────────
def build_settings_xml():
    """Minimal settings.xml with sensible defaults."""
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:defaultTabStop w:val="720"/>
  <w:autoHyphenation w:val="0"/>
  <w:compat>
    <w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
  </w:compat>
</w:settings>
"""


# Decides what kind of content a paragraph is, so we can apply the right style.
# ─────────────────────────────────────────────────────────────────────────────
def classify_paragraph(pStyle, text, idx, total):
    """
    Returns a tuple (new_style_id, extra_props_dict).
    Heuristics based on position, current style, text content.
    """
    text_clean = text.strip()
    text_lower = text_clean.lower()
    n_words = len(text_clean.split())

    # ── Detect existing heading styles ──────────────────────────────────────
    if pStyle in ("Heading1", "heading 1"):
        return ("Heading1", {})
    if pStyle in ("Heading2", "heading 2"):
        return ("Heading2", {})
    if pStyle in ("Heading3", "heading 3"):
        return ("Heading3", {})

    # ── Cover block (first ~6 paras if Title/Subtitle style) ─────────────
    if pStyle in ("Title",):
        return ("CoverTitle", {})
    if pStyle in ("Subtitle",):
        return ("CoverSubtitle", {})

    # ── Captions (Caption style or starts with Figure/Table/Fig) ─────────
    if pStyle in ("Caption", "caption"):
        return ("Caption", {})
    if re.match(r'^(fig(ure)?|table|photo|map|chart|plate)\s*[\d\.IVX]+', text_clean, re.I):
        return ("Caption", {})

    # ── Pull quote: italic-looking short quote para ───────────────────────
    if pStyle in ("Quote", "IntenseQuote", "Intense Quote"):
        return ("PullQuote", {})

    # ── Very short all-caps para = label/heading ─────────────────────────
    if text_clean and text_clean == text_clean.upper() and n_words <= 6 and len(text_clean) > 3:
        # Could be a cover label like "INTERIM REPORT"
        if idx < 8:
            return ("CoverLabel", {})
        return ("Heading1", {})

    # ── Short bold paras that look like headings ─────────────────────────
    # (handled by checking run-level bold below in apply_paragraph_formatting)

    # ── Management Recommendation / Callout boxes ─────────────────────────
    callout_triggers = (
        "management recommendation",
        "key recommendation",
        "recommendation",
        "key concern",
        "key finding",
        "action required",
        "immediate action",
        "conservation alert",
        "next steps",
        "proposed action",
    )
    if any(text_lower.startswith(t) for t in callout_triggers):
        return ("_CALLOUT_HEADER", {})

    # ── Normal body ───────────────────────────────────────────────────────
    return ("Normal", {})


# ─────────────────────────────────────────────────────────────────────────────
# LOW-LEVEL XML HELPERS
# ─────────────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
# pPr CHILD ORDER (per OOXML schema) — rPr must always be last
# ─────────────────────────────────────────────────────────────────────────────
PPR_ORDER = [
    "pStyle","numPr","keepNext","keepLines","pageBreakBefore","framePr",
    "suppressLineNumbers","pBdr","shd","tabs","suppressAutoHyphens","kinsoku",
    "wordWrap","adjustRightInd","snapToGrid","spacing","ind","contextualSpacing",
    "mirrorIndents","suppressOverlap","jc","textDirection","textAlignment",
    "textboxTightWrap","outlineLvl","divId","cnfStyle","rPr","sectPr","pPrChange"
]

def _reorder_pPr(pPr):
    """Re-sort children of pPr into OOXML schema order so rPr is always last."""
    ns = NS["w"]
    children = list(pPr)
    if not children:
        return
    def sort_key(el):
        local = el.tag.split("}")[-1] if "}" in el.tag else el.tag
        try:
            return PPR_ORDER.index(local)
        except ValueError:
            return 999
    children.sort(key=sort_key)
    for child in children:
        pPr.remove(child)
    for child in children:
        pPr.append(child)

def _insert_before_rPr(pPr, el):
    """Insert element into pPr before rPr (keeps rPr last)."""
    ns = NS["w"]
    rPr = pPr.find(f"{{{ns}}}rPr")
    if rPr is not None:
        idx = list(pPr).index(rPr)
        pPr.insert(idx, el)
    else:
        pPr.append(el)

def get_pPr(p):
    """Get or create w:pPr element (always first child of p)."""
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

def get_or_create(parent, tag_local):
    ns = NS["w"]
    el = parent.find(f"{{{ns}}}{tag_local}")
    if el is None:
        el = etree.SubElement(parent, f"{{{ns}}}{tag_local}")
    return el

def set_shd_pPr(pPr, fill, val="clear", color="auto"):
    """Set shading on a pPr, respecting element order (shd before spacing)."""
    ns = NS["w"]
    shd = pPr.find(f"{{{ns}}}shd")
    if shd is None:
        shd = etree.Element(f"{{{ns}}}shd")
        _insert_before_rPr(pPr, shd)
    shd.set(f"{{{ns}}}val", val)
    shd.set(f"{{{ns}}}color", color)
    shd.set(f"{{{ns}}}fill", fill)
    _reorder_pPr(pPr)

def set_shd(pPr_or_tcPr, fill, val="clear", color="auto"):
    """Generic shd setter (for tcPr compatibility)."""
    ns = NS["w"]
    shd = pPr_or_tcPr.find(f"{{{ns}}}shd")
    if shd is None:
        shd = etree.SubElement(pPr_or_tcPr, f"{{{ns}}}shd")
    shd.set(f"{{{ns}}}val", val)
    shd.set(f"{{{ns}}}color", color)
    shd.set(f"{{{ns}}}fill", fill)

def set_spacing(pPr, before=None, after=None, line=None, lineRule="auto"):
    ns = NS["w"]
    sp = pPr.find(f"{{{ns}}}spacing")
    if sp is None:
        sp = etree.Element(f"{{{ns}}}spacing")
        _insert_before_rPr(pPr, sp)
    if before is not None:
        sp.set(f"{{{ns}}}before", str(before))
    if after is not None:
        sp.set(f"{{{ns}}}after", str(after))
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
    parts = []
    for t in p.iter(f"{{{ns}}}t"):
        parts.append(t.text or "")
    return "".join(parts)

def get_para_style(p):
    ns = NS["w"]
    pPr = p.find(f"{{{ns}}}pPr")
    if pPr is None:
        return ""
    ps = pPr.find(f"{{{ns}}}pStyle")
    if ps is None:
        return ""
    return ps.get(f"{{{ns}}}val", "")

def is_para_bold(p):
    """Returns True if all non-empty runs in paragraph are bold."""
    ns = NS["w"]
    runs = p.findall(f"{{{ns}}}r")
    if not runs:
        return False
    for r in runs:
        t = r.find(f"{{{ns}}}t")
        if t is None or not (t.text or "").strip():
            continue
        rPr = r.find(f"{{{ns}}}rPr")
        if rPr is None:
            return False
        b = rPr.find(f"{{{ns}}}b")
        if b is None:
            return False
    return True

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

def ensure_run_size(r, half_pts):
    ns = NS["w"]
    rPr = r.find(f"{{{ns}}}rPr")
    if rPr is None:
        rPr = etree.SubElement(r, f"{{{ns}}}rPr")
        r.insert(0, rPr)
    sz = rPr.find(f"{{{ns}}}sz")
    szCs = rPr.find(f"{{{ns}}}szCs")
    val = str(half_pts)
    if sz is None:
        sz = etree.SubElement(rPr, f"{{{ns}}}sz")
    sz.set(f"{{{ns}}}val", val)
    if szCs is None:
        szCs = etree.SubElement(rPr, f"{{{ns}}}szCs")
    szCs.set(f"{{{ns}}}val", val)


# ─────────────────────────────────────────────────────────────────────────────
# TABLE FORMATTER
# ─────────────────────────────────────────────────────────────────────────────
def set_ind(pPr, left=None, right=None, firstLine=None, hanging=None):
    """Set indentation on a pPr element."""
    ns = NS["w"]
    ind = pPr.find(f"{{{ns}}}ind")
    if ind is None:
        ind = etree.Element(f"{{{ns}}}ind")
        _insert_before_rPr(pPr, ind)
    # Only set non-None, non-zero values (zero = reset to none)
    for attr, val in [("left", left), ("right", right), ("firstLine", firstLine), ("hanging", hanging)]:
        if val is not None:
            if val == 0:
                ind.attrib.pop(f"{{{ns}}}{attr}", None)
            else:
                ind.set(f"{{{ns}}}{attr}", str(val))
    # Remove empty ind element
    if not ind.attrib:
        pPr.remove(ind)
    else:
        _reorder_pPr(pPr)


def is_list_paragraph(p):
    """Returns True if paragraph has numPr (is a list item)."""
    ns = NS["w"]
    pPr = p.find(f"{{{ns}}}pPr")
    if pPr is None:
        return False
    return pPr.find(f"{{{ns}}}numPr") is not None


def detect_list_level(p):
    """Returns (is_list, ilvl) - indent level 0-based."""
    ns = NS["w"]
    pPr = p.find(f"{{{ns}}}pPr")
    if pPr is None:
        return False, 0
    numPr = pPr.find(f"{{{ns}}}numPr")
    if numPr is None:
        return False, 0
    ilvl_el = numPr.find(f"{{{ns}}}ilvl")
    ilvl = int(ilvl_el.get(f"{{{ns}}}val", "0")) if ilvl_el is not None else 0
    return True, ilvl


    """Apply professional formatting to a table element."""
    T = THEMES[theme]
    ns = NS["w"]

    # Switch table style
    tblPr = tbl.find(f"{{{ns}}}tblPr")
    if tblPr is None:
        tblPr = etree.SubElement(tbl, f"{{{ns}}}tblPr")
        tbl.insert(0, tblPr)

    # Set style
    tblStyle = tblPr.find(f"{{{ns}}}tblStyle")
    if tblStyle is None:
        tblStyle = etree.SubElement(tblPr, f"{{{ns}}}tblStyle")
    tblStyle.set(f"{{{ns}}}val", "TableGrid")

    # Set borders: top+bottom = accent, sides = none, insideH = light grey, insideV = none
    tblBorders = tblPr.find(f"{{{ns}}}tblBorders")
    if tblBorders is None:
        tblBorders = etree.SubElement(tblPr, f"{{{ns}}}tblBorders")

    def set_border(parent, side, val, sz, color):
        el = parent.find(f"{{{ns}}}{side}")
        if el is None:
            el = etree.SubElement(parent, f"{{{ns}}}{side}")
        el.set(f"{{{ns}}}val", val)
        el.set(f"{{{ns}}}sz", str(sz))
        el.set(f"{{{ns}}}space", "0")
        el.set(f"{{{ns}}}color", color)

    set_border(tblBorders, "top",     "single", 8, T["accent"])
    set_border(tblBorders, "left",    "none",   0, "auto")
    set_border(tblBorders, "bottom",  "single", 8, T["accent"])
    set_border(tblBorders, "right",   "none",   0, "auto")
    set_border(tblBorders, "insideH", "single", 4, "CCCCCC")
    set_border(tblBorders, "insideV", "none",   0, "auto")

    # Cell margins
    tblCellMar = tblPr.find(f"{{{ns}}}tblCellMar")
    if tblCellMar is None:
        tblCellMar = etree.SubElement(tblPr, f"{{{ns}}}tblCellMar")
    for side, val in [("top", "100"), ("left", "140"), ("bottom", "100"), ("right", "140")]:
        el = tblCellMar.find(f"{{{ns}}}{side}")
        if el is None:
            el = etree.SubElement(tblCellMar, f"{{{ns}}}{side}")
        el.set(f"{{{ns}}}w", val)
        el.set(f"{{{ns}}}type", "dxa")

    # Format rows
    rows = tbl.findall(f"{{{ns}}}tr")
    for r_idx, tr in enumerate(rows):
        cells = tr.findall(f"{{{ns}}}tc")

        # Detect if this looks like a callout row (single merged cell, key phrases)
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

            # Cell shading
            shd = tcPr.find(f"{{{ns}}}shd")
            if shd is None:
                shd = etree.SubElement(tcPr, f"{{{ns}}}shd")

            if is_callout:
                # Dark teal callout
                shd.set(f"{{{ns}}}val", "clear")
                shd.set(f"{{{ns}}}color", "auto")
                shd.set(f"{{{ns}}}fill", T["callout_bg"])
                # Make text white
                for p in tc.findall(f"{{{ns}}}p"):
                    for run in p.findall(f"{{{ns}}}r"):
                        set_run_color(run, T["callout_text"])
                        set_run_font(run)
                    # Also rPr in pPr
                    pPr = p.find(f"{{{ns}}}pPr")
                    if pPr is not None:
                        rPr_in_pPr = pPr.find(f"{{{ns}}}rPr")
                        if rPr_in_pPr is not None:
                            col = rPr_in_pPr.find(f"{{{ns}}}color")
                            if col is None:
                                col = etree.SubElement(rPr_in_pPr, f"{{{ns}}}color")
                            col.set(f"{{{ns}}}val", T["callout_text"])

            elif r_idx == 0:
                # Header row
                shd.set(f"{{{ns}}}val", "clear")
                shd.set(f"{{{ns}}}color", "auto")
                shd.set(f"{{{ns}}}fill", T["table_header"])
                # Make text bold
                for p in tc.findall(f"{{{ns}}}p"):
                    for run in p.findall(f"{{{ns}}}r"):
                        rPr = run.find(f"{{{ns}}}rPr")
                        if rPr is None:
                            rPr = etree.SubElement(run, f"{{{ns}}}rPr")
                            run.insert(0, rPr)
                        if rPr.find(f"{{{ns}}}b") is None:
                            etree.SubElement(rPr, f"{{{ns}}}b")
                        if rPr.find(f"{{{ns}}}bCs") is None:
                            etree.SubElement(rPr, f"{{{ns}}}bCs")
                        set_run_font(run)

            elif r_idx % 2 == 0:
                # Even data rows: alt shading
                shd.set(f"{{{ns}}}val", "clear")
                shd.set(f"{{{ns}}}color", "auto")
                shd.set(f"{{{ns}}}fill", T["table_alt"])
            else:
                # Odd data rows: white
                shd.set(f"{{{ns}}}val", "clear")
                shd.set(f"{{{ns}}}color", "auto")
                shd.set(f"{{{ns}}}fill", "FFFFFF")

            # Ensure all runs use Arial
            for p in tc.findall(f"{{{ns}}}p"):
                for run in p.findall(f"{{{ns}}}r"):
                    set_run_font(run)


# ─────────────────────────────────────────────────────────────────────────────
# MAIN PARAGRAPH FORMATTER
# ─────────────────────────────────────────────────────────────────────────────
def format_paragraph(p, idx, total, theme, section_counter):
    """Apply formatting to a single paragraph element. Returns updated section_counter."""
    T = THEMES[theme]
    IP = INDENT_PROFILES[T.get("indent_profile", "formal")]
    ns = NS["w"]

    text = get_para_text(p)
    pStyle = get_para_style(p)
    new_style, extra = classify_paragraph(pStyle, text, idx, total)

    pPr = get_pPr(p)

    # Detect list items early — they get list-specific indentation
    is_list, ilvl = detect_list_level(p)
    if is_list:
        # Apply bullet/numbered indent based on level
        level_mult = ilvl + 1
        bl = IP["bullet_left"] + (ilvl * IP["bullet_hanging"])
        bh = IP["bullet_hanging"]
        set_ind(pPr, left=bl, hanging=bh)
        set_spacing(pPr, before=0, after=80, line=276, lineRule="auto")
        ensure_jc(pPr, "both")
        for r in p.findall(f"{{{ns}}}r"):
            set_run_font(r)
            rPr = r.find(f"{{{ns}}}rPr")
            if rPr is not None:
                col = rPr.find(f"{{{ns}}}color")
                if col is None:
                    col = etree.SubElement(rPr, f"{{{ns}}}color")
                    col.set(f"{{{ns}}}val", T["body"])
        return section_counter

    # ── Apply new style ───────────────────────────────────────────────────
    if new_style == "_CALLOUT_HEADER":
        pass
    else:
        set_pStyle(pPr, new_style)

    # ── Per-style custom handling ─────────────────────────────────────────

    if new_style == "Heading1":
        section_counter[0] += 1
        set_spacing(pPr, before=360, after=100, line=240, lineRule="auto")
        ensure_jc(pPr, "left")
        set_pBdr_bottom(pPr, T["accent"], sz=8, space=4)
        if IP["heading1_left"]:
            set_ind(pPr, left=IP["heading1_left"])
        for r in p.findall(f"{{{ns}}}r"):
            set_run_color(r, T["accent"])
            set_run_font(r)

    elif new_style == "Heading2":
        set_spacing(pPr, before=240, after=80, line=240, lineRule="auto")
        ensure_jc(pPr, "left")
        if IP["heading2_left"]:
            set_ind(pPr, left=IP["heading2_left"])
        for r in p.findall(f"{{{ns}}}r"):
            set_run_color(r, T["accent_light"])
            set_run_font(r)

    elif new_style == "Heading3":
        set_spacing(pPr, before=200, after=60, line=240, lineRule="auto")
        ensure_jc(pPr, "left")
        if IP["heading3_left"]:
            set_ind(pPr, left=IP["heading3_left"])
        for r in p.findall(f"{{{ns}}}r"):
            set_run_color(r, T["accent_light"])
            set_run_font(r)

    elif new_style in ("CoverTitle", "CoverSubtitle", "CoverLabel"):
        fill = T["primary"] if new_style in ("CoverTitle", "CoverLabel") else T["cover_sub"]
        set_shd_pPr(pPr, fill)
        # No extra indent on cover block
        set_ind(pPr, left=0, right=0, firstLine=0)
        for r in p.findall(f"{{{ns}}}r"):
            set_run_color(r, "FFFFFF")
            set_run_font(r)
        ensure_jc(pPr, "left")

    elif new_style == "Caption":
        set_spacing(pPr, before=60, after=160, line=240, lineRule="auto")
        ensure_jc(pPr, "center")
        if IP["caption_left"] or IP["caption_right"]:
            set_ind(pPr, left=IP["caption_left"], right=IP["caption_right"])
        for r in p.findall(f"{{{ns}}}r"):
            set_run_color(r, T["caption"])
            set_run_font(r)

    elif new_style == "PullQuote":
        set_spacing(pPr, before=160, after=160, line=276, lineRule="auto")
        set_pBdr_left(pPr, T["accent"])
        set_ind(pPr, left=IP["pullquote_left"], right=IP["pullquote_right"])
        ensure_jc(pPr, "both")
        for r in p.findall(f"{{{ns}}}r"):
            set_run_color(r, T["italic_quote"])
            set_run_font(r)

    elif new_style == "_CALLOUT_HEADER":
        set_pStyle(pPr, "Normal")
        set_shd_pPr(pPr, T["callout_bg"])
        set_spacing(pPr, before=120, after=60, line=240, lineRule="auto")
        set_ind(pPr, left=IP["callout_left"], right=IP["callout_right"])
        ensure_jc(pPr, "left")
        for r in p.findall(f"{{{ns}}}r"):
            set_run_color(r, T["callout_text"])
            set_run_font(r)

    else:
        # Normal body text
        set_spacing(pPr, before=0, after=120, line=276, lineRule="auto")
        # Apply first-line indent if the profile requests it
        set_ind(pPr,
                left=IP["body_left"] if IP["body_left"] else None,
                right=IP["body_right"] if IP["body_right"] else None,
                firstLine=IP["body_first_line"] if IP["body_first_line"] else None)
        ensure_jc(pPr, "both")
        for r in p.findall(f"{{{ns}}}r"):
            set_run_font(r)
            rPr = r.find(f"{{{ns}}}rPr")
            if rPr is not None:
                col = rPr.find(f"{{{ns}}}color")
                if col is None:
                    col = etree.SubElement(rPr, f"{{{ns}}}color")
                    col.set(f"{{{ns}}}val", T["body"])

    # Ensure font on all runs (catch stragglers)
    for r in p.findall(f"{{{ns}}}r"):
        set_run_font(r)

    return section_counter


# ─────────────────────────────────────────────────────────────────────────────
# HEURISTIC: detect short all-bold paras that look like untagged headings
# ─────────────────────────────────────────────────────────────────────────────
HEADING_WORDS_BLACKLIST = {"the", "a", "an", "and", "of", "in", "to", "for", "with", "at", "by", "from", "on"}

def looks_like_heading(p, ns):
    text = get_para_text(p).strip()
    if not text:
        return False
    words = text.split()
    if len(words) > 12:
        return False
    if text.endswith(".") and len(words) > 6:
        return False
    if is_para_bold(p) and len(words) <= 10:
        return True
    return False


# ─────────────────────────────────────────────────────────────────────────────
# SECTION RELATIONSHIP UPDATER
# ─────────────────────────────────────────────────────────────────────────────
def update_section_references(doc_zip, theme, report_label):
    """Patch sectPr to reference our footer, write footer1.xml."""
    import zipfile
    ns_w = NS["w"]
    ns_r = NS["r"]

    # Build footer XML
    footer_xml = build_footer_xml(theme, report_label)

    return footer_xml


# ─────────────────────────────────────────────────────────────────────────────
# CORE FORMATTER
# ─────────────────────────────────────────────────────────────────────────────
def format_document(input_path, output_path, theme="forest", report_label=None):
    """
    Main entry point. Reads input_path, applies professional formatting,
    writes output_path.
    """
    T = THEMES[theme]
    ns = NS["w"]
    ns_r = NS["r"]

    if report_label is None:
        report_label = Path(input_path).stem.replace("_", " ")

    print(f"  Theme: {T['name']}")
    print(f"  Input:  {input_path}")
    print(f"  Output: {output_path}")

    # ── Unpack ────────────────────────────────────────────────────────────
    work_dir = Path(output_path).parent / "_fmt_work"
    if work_dir.exists():
        shutil.rmtree(work_dir)
    with zipfile.ZipFile(input_path, "r") as z:
        z.extractall(work_dir)

    # ── Parse document.xml ────────────────────────────────────────────────
    doc_xml_path = work_dir / "word" / "document.xml"
    tree = etree.parse(str(doc_xml_path))
    root = tree.getroot()
    body = root.find(f"{{{ns}}}body")

    # Collect all block-level elements
    children = list(body)
    paragraphs = [c for c in children if c.tag == f"{{{ns}}}p"]
    tables     = [c for c in children if c.tag == f"{{{ns}}}tbl"]
    total_p    = len(paragraphs)

    print(f"  Found: {total_p} paragraphs, {len(tables)} tables")

    # ── Promote untagged bold short paras to headings ─────────────────────
    for p in paragraphs:
        pStyle = get_para_style(p)
        if pStyle == "" or pStyle == "Normal":
            if looks_like_heading(p, ns):
                pPr = get_pPr(p)
                set_pStyle(pPr, "Heading2")

    # ── Format each paragraph ─────────────────────────────────────────────
    section_counter = [0]
    for idx, p in enumerate(paragraphs):
        section_counter = format_paragraph(p, idx, total_p, theme, section_counter)

    # ── Format tables ─────────────────────────────────────────────────────
    for tbl in tables:
        format_table(tbl, theme)

    # ── Write updated document.xml ────────────────────────────────────────
    tree.write(str(doc_xml_path), xml_declaration=True, encoding="UTF-8", standalone=True)

    # ── Write styles.xml ──────────────────────────────────────────────────
    styles_path = work_dir / "word" / "styles.xml"
    with open(styles_path, "w", encoding="utf-8") as f:
        f.write(build_styles_xml(theme))

    # ── Write numbering.xml ───────────────────────────────────────────────
    numbering_path = work_dir / "word" / "numbering.xml"
    with open(numbering_path, "w", encoding="utf-8") as f:
        f.write(build_numbering_xml(theme))
    _wire_numbering(work_dir, ns_r)

    # ── Write footer1.xml ─────────────────────────────────────────────────
    footer_path = work_dir / "word" / "footer1.xml"
    with open(footer_path, "w", encoding="utf-8") as f:
        f.write(build_footer_xml(theme, report_label))

    # ── Write settings.xml ────────────────────────────────────────────────
    settings_path = work_dir / "word" / "settings.xml"
    with open(settings_path, "w", encoding="utf-8") as f:
        f.write(build_settings_xml())

    # ── Ensure footer is wired into sectPr and relationships ─────────────
    _wire_footer(work_dir, ns, ns_r)

    # ── Apply page layout (size + margins) ────────────────────────────────
    _apply_page_layout(work_dir, theme, ns)

    # ── Repack ────────────────────────────────────────────────────────────
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zout:
        for fpath in sorted(work_dir.rglob("*")):
            if fpath.is_file():
                arcname = fpath.relative_to(work_dir)
                zout.write(fpath, arcname)

    shutil.rmtree(work_dir)
    print(f"  ✓ Done → {output_path}")


def _apply_page_layout(work_dir, theme, ns):
    """
    Patch (or create) the sectPr in document.xml with page size and margins
    from the theme's PAGE_LAYOUT profile.
    """
    T  = THEMES[theme]
    PL = PAGE_LAYOUTS[T.get("page_layout", "a4_standard")]

    doc_xml_path = work_dir / "word" / "document.xml"
    tree = etree.parse(str(doc_xml_path))
    root = tree.getroot()
    body = root.find(f"{{{ns}}}body")
    if body is None:
        return

    # sectPr must be the LAST element in body
    sectPr = body.find(f"{{{ns}}}sectPr")
    if sectPr is None:
        sectPr = etree.SubElement(body, f"{{{ns}}}sectPr")

    def _set_or_create(parent, tag_local, attrs):
        el = parent.find(f"{{{ns}}}{tag_local}")
        if el is None:
            el = etree.SubElement(parent, f"{{{ns}}}{tag_local}")
        for k, v in attrs.items():
            el.set(f"{{{ns}}}{k}", str(v))
        return el

    _set_or_create(sectPr, "pgSz", {
        "w": PL["page_width"],
        "h": PL["page_height"],
    })
    _set_or_create(sectPr, "pgMar", {
        "top":    PL["margin_top"],
        "right":  PL["margin_right"],
        "bottom": PL["margin_bottom"],
        "left":   PL["margin_left"],
        "header": PL["margin_header"],
        "footer": PL["margin_footer"],
        "gutter": PL["gutter"],
    })
    # Turn on widow/orphan control for the whole section
    _set_or_create(sectPr, "widowControl", {})

    tree.write(str(doc_xml_path), xml_declaration=True, encoding="UTF-8", standalone=True)



def _wire_footer(work_dir, ns, ns_r):
    """Ensure footer1.xml is referenced in document.xml.rels and sectPr."""
    rels_path = work_dir / "word" / "_rels" / "document.xml.rels"
    if not rels_path.exists():
        return

    rels_tree = etree.parse(str(rels_path))
    rels_root = rels_tree.getroot()

    # Check if footer1 relationship already exists
    footer_rel_id = None
    for rel in rels_root:
        if "footer1" in rel.get("Target", ""):
            footer_rel_id = rel.get("Id")
            break

    if footer_rel_id is None:
        # Add it
        existing_ids = [int(r.get("Id", "rId0").replace("rId", "0")) for r in rels_root if r.get("Id", "").startswith("rId")]
        new_id_num = max(existing_ids, default=0) + 1
        footer_rel_id = f"rId{new_id_num}"
        rel_ns = "http://schemas.openxmlformats.org/package/2006/relationships"
        new_rel = etree.SubElement(rels_root, "Relationship")
        new_rel.set("Id", footer_rel_id)
        new_rel.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer")
        new_rel.set("Target", "footer1.xml")
        rels_tree.write(str(rels_path), xml_declaration=True, encoding="UTF-8", standalone=True)

    # Ensure sectPr has footerReference
    doc_xml_path = work_dir / "word" / "document.xml"
    doc_tree = etree.parse(str(doc_xml_path))
    doc_root = doc_tree.getroot()
    body = doc_root.find(f"{{{ns}}}body")
    if body is None:
        return

    sectPr = body.find(f"{{{ns}}}sectPr")
    if sectPr is None:
        sectPr = etree.SubElement(body, f"{{{ns}}}sectPr")

    # Check for existing footerReference
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

    # Suppress header/footer on first page if desired (optional)
    doc_tree.write(str(doc_xml_path), xml_declaration=True, encoding="UTF-8", standalone=True)

    # Add footer1.xml to [Content_Types].xml if not already there
    ct_path = work_dir / "[Content_Types].xml"
    if ct_path.exists():
        ct_tree = etree.parse(str(ct_path))
        ct_root = ct_tree.getroot()
        ct_ns = "http://schemas.openxmlformats.org/package/2006/content-types"
        footer_ct = f"{{{ct_ns}}}Override"
        existing = [el.get("PartName", "") for el in ct_root]
        if "/word/footer1.xml" not in existing:
            new_ct = etree.SubElement(ct_root, f"{{{ct_ns}}}Override")
            new_ct.set("PartName", "/word/footer1.xml")
            new_ct.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml")
        ct_tree.write(str(ct_path), xml_declaration=True, encoding="UTF-8", standalone=True)


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="Professional formatter for conservation/donor Word documents",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument("input",  help="Input .docx file")
    parser.add_argument("output", help="Output .docx file")
    parser.add_argument(
        "--theme",
        choices=list(THEMES.keys()),
        default="forest",
        help="Colour theme (default: forest)"
    )
    parser.add_argument(
        "--label",
        default=None,
        help="Report label for footer (default: input filename)"
    )
    parser.add_argument(
        "--list-themes",
        action="store_true",
        help="Print available themes and exit"
    )

    args = parser.parse_args()

    if args.list_themes:
        print("\nAvailable themes:")
        for key, t in THEMES.items():
            print(f"  {key:10s}  {t['name']}")
            print(f"             Primary: #{t['primary']}  Accent: #{t['accent']}")
        print()
        sys.exit(0)

    if not os.path.exists(args.input):
        print(f"Error: input file not found: {args.input}")
        sys.exit(1)

    print(f"\n{'─'*60}")
    print("  Conservation Report Formatter")
    print(f"{'─'*60}")
    format_document(args.input, args.output, theme=args.theme, report_label=args.label)
    print(f"{'─'*60}\n")


if __name__ == "__main__":
    main()
