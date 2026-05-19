"""Generate a professional .pptx for the Inter Branch English Grammar Quiz.

This script writes the Open XML (PresentationML) structure directly using
only Python's standard library (zipfile + string templating). It produces
a clean, classroom-ready 16:9 deck with a dark-blue / white / gold theme
and one question per slide.
"""

from __future__ import annotations

import os
import zipfile
from xml.sax.saxutils import escape as xml_escape

from quiz_data import ROUNDS, TEAM_NAMES

# ---------------------------------------------------------------------------
# Theme constants
# ---------------------------------------------------------------------------
SLIDE_W = 12192000  # 13.333 in (widescreen)
SLIDE_H = 6858000   # 7.5 in
EMU_PER_INCH = 914400

NAVY = "0F2A5C"        # primary dark blue
NAVY_DEEP = "081A3C"   # darker accent for gradients
GOLD = "D4AF37"        # accent gold
GOLD_LIGHT = "F0D27A"  # lighter gold for hover/details
WHITE = "FFFFFF"
SOFT_WHITE = "F5F7FB"
INK = "0B1B3A"         # near-black for body text on light backgrounds


def emu_inches(inches: float) -> int:
    return int(round(inches * EMU_PER_INCH))


# ---------------------------------------------------------------------------
# Static package parts (Content Types, rels, theme, masters, layouts, props)
# ---------------------------------------------------------------------------

def content_types_xml(num_slides: int) -> str:
    overrides = []
    for i in range(1, num_slides + 1):
        overrides.append(
            f'<Override PartName="/ppt/slides/slide{i}.xml" '
            f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        )
    overrides_xml = "".join(overrides)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/ppt/presentation.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
        '<Override PartName="/ppt/theme/theme1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        '<Override PartName="/ppt/presProps.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>'
        '<Override PartName="/ppt/viewProps.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>'
        '<Override PartName="/ppt/tableStyles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>'
        '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        f'{overrides_xml}'
        '</Types>'
    )


ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
    'Target="ppt/presentation.xml"/>'
    '<Relationship Id="rId2" '
    'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" '
    'Target="docProps/core.xml"/>'
    '<Relationship Id="rId3" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" '
    'Target="docProps/app.xml"/>'
    '</Relationships>'
)


def docprops_core() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<cp:coreProperties '
        'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
        'xmlns:dc="http://purl.org/dc/elements/1.1/" '
        'xmlns:dcterms="http://purl.org/dc/terms/" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
        '<dc:title>Inter Branch English Grammar Quiz</dc:title>'
        '<dc:creator>Quiz Master</dc:creator>'
        '<cp:lastModifiedBy>Quiz Master</cp:lastModifiedBy>'
        '<dcterms:created xsi:type="dcterms:W3CDTF">2026-05-01T09:00:00Z</dcterms:created>'
        '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-05-01T09:00:00Z</dcterms:modified>'
        '</cp:coreProperties>'
    )


def docprops_app(num_slides: int) -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties '
        'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
        'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
        '<Application>Kiro PPTX Builder</Application>'
        f'<Slides>{num_slides}</Slides>'
        '<ScaleCrop>false</ScaleCrop>'
        '<LinksUpToDate>false</LinksUpToDate>'
        '<SharedDoc>false</SharedDoc>'
        '<HyperlinksChanged>false</HyperlinksChanged>'
        '<AppVersion>16.0000</AppVersion>'
        '</Properties>'
    )


def theme_xml() -> str:
    # Custom theme using our brand colors.
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Quiz Show">
  <a:themeElements>
    <a:clrScheme name="Quiz Show">
      <a:dk1><a:srgbClr val="0B1B3A"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="{NAVY}"/></a:dk2>
      <a:lt2><a:srgbClr val="{SOFT_WHITE}"/></a:lt2>
      <a:accent1><a:srgbClr val="{GOLD}"/></a:accent1>
      <a:accent2><a:srgbClr val="{NAVY_DEEP}"/></a:accent2>
      <a:accent3><a:srgbClr val="{GOLD_LIGHT}"/></a:accent3>
      <a:accent4><a:srgbClr val="2C4A7E"/></a:accent4>
      <a:accent5><a:srgbClr val="B0883E"/></a:accent5>
      <a:accent6><a:srgbClr val="DCE3F2"/></a:accent6>
      <a:hlink><a:srgbClr val="{GOLD}"/></a:hlink>
      <a:folHlink><a:srgbClr val="B0883E"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Quiz Show">
      <a:majorFont>
        <a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Quiz Show">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
        <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults/>
  <a:extraClrSchemeLst/>
</a:theme>'''


def slide_master_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="{WHITE}"/></a:solidFill>
        <a:effectLst/>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"
            accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6"
            hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
  <p:transition spd="med">
    <p:fade/>
  </p:transition>
  <p:txStyles>
    <p:titleStyle>
      <a:lvl1pPr algn="l"><a:defRPr sz="4400" b="1"><a:solidFill><a:srgbClr val="{NAVY}"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:lvl1pPr>
    </p:titleStyle>
    <p:bodyStyle>
      <a:lvl1pPr><a:defRPr sz="2400"><a:solidFill><a:srgbClr val="{INK}"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:lvl1pPr>
      <a:lvl2pPr><a:defRPr sz="2000"><a:solidFill><a:srgbClr val="{INK}"/></a:solidFill><a:latin typeface="Calibri"/></a:defRPr></a:lvl2pPr>
    </p:bodyStyle>
    <p:otherStyle/>
  </p:txStyles>
</p:sldMaster>'''


SLIDE_MASTER_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" '
    'Target="../slideLayouts/slideLayout1.xml"/>'
    '<Relationship Id="rId2" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" '
    'Target="../theme/theme1.xml"/>'
    '</Relationships>'
)


def slide_layout_xml() -> str:
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             type="blank" preserve="1">
  <p:cSld name="Blank">
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/><a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:spTree>
  </p:cSld>
  <p:transition spd="med"><p:fade/></p:transition>
</p:sldLayout>'''


SLIDE_LAYOUT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" '
    'Target="../slideMasters/slideMaster1.xml"/>'
    '</Relationships>'
)


PRES_PROPS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
)

VIEW_PROPS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
    '<p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr>'
    '<p:slideViewPr><p:cSldViewPr><p:cViewPr varScale="1"><p:scale><a:sx n="68" d="100"/><a:sy n="68" d="100"/></p:scale><p:origin x="-1500" y="-90"/></p:cViewPr></p:cSldViewPr></p:slideViewPr>'
    '<p:gridSpacing cx="76200" cy="76200"/></p:viewPr>'
)

TABLE_STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>'
)


def presentation_xml(num_slides: int) -> str:
    sld_ids = []
    for i in range(num_slides):
        # Relationship id for slide i is rId(i+2) because rId1 is the slideMaster
        sld_ids.append(
            f'<p:sldId id="{256 + i}" r:id="rId{i + 2}"/>'
        )
    sld_id_lst = "".join(sld_ids)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:presentation '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        f'<p:sldIdLst>{sld_id_lst}</p:sldIdLst>'
        f'<p:sldSz cx="{SLIDE_W}" cy="{SLIDE_H}" type="screen16x9"/>'
        '<p:notesSz cx="6858000" cy="9144000"/>'
        '<p:defaultTextStyle>'
        '<a:defPPr><a:defRPr lang="en-US"/></a:defPPr>'
        '</p:defaultTextStyle>'
        '</p:presentation>'
    )


def presentation_rels(num_slides: int) -> str:
    rels = [
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" '
        'Target="slideMasters/slideMaster1.xml"/>',
    ]
    for i in range(1, num_slides + 1):
        rels.append(
            f'<Relationship Id="rId{i + 1}" '
            f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" '
            f'Target="slides/slide{i}.xml"/>'
        )
    next_id = num_slides + 2
    rels.append(
        f'<Relationship Id="rId{next_id}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" '
        f'Target="theme/theme1.xml"/>'
    )
    rels.append(
        f'<Relationship Id="rId{next_id + 1}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" '
        f'Target="presProps.xml"/>'
    )
    rels.append(
        f'<Relationship Id="rId{next_id + 2}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" '
        f'Target="viewProps.xml"/>'
    )
    rels.append(
        f'<Relationship Id="rId{next_id + 3}" '
        f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" '
        f'Target="tableStyles.xml"/>'
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(rels)
        + '</Relationships>'
    )


def slide_rels_xml() -> str:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" '
        'Target="../slideLayouts/slideLayout1.xml"/>'
        '</Relationships>'
    )


# ---------------------------------------------------------------------------
# Slide-level shape helpers
# ---------------------------------------------------------------------------

def _xfrm(x_emu: int, y_emu: int, cx_emu: int, cy_emu: int) -> str:
    return (
        f'<a:xfrm><a:off x="{x_emu}" y="{y_emu}"/>'
        f'<a:ext cx="{cx_emu}" cy="{cy_emu}"/></a:xfrm>'
    )


def _runs_from_lines(
    lines,
    *,
    size: int,
    bold: bool = False,
    color: str = INK,
    italic: bool = False,
    align: str = "l",
    font: str = "Calibri",
    line_spacing: int = 100,
) -> str:
    """Build text body paragraphs from a list of strings."""
    paragraphs = []
    b = "1" if bold else "0"
    i = "1" if italic else "0"
    for ln in lines:
        text = xml_escape(ln) if ln else ""
        if not text:
            paragraphs.append(
                f'<a:p><a:pPr algn="{align}"><a:lnSpc><a:spcPct val="{line_spacing * 1000}"/></a:lnSpc>'
                f'</a:pPr><a:endParaRPr lang="en-US" sz="{size}"/></a:p>'
            )
            continue
        paragraphs.append(
            f'<a:p>'
            f'<a:pPr algn="{align}"><a:lnSpc><a:spcPct val="{line_spacing * 1000}"/></a:lnSpc></a:pPr>'
            f'<a:r>'
            f'<a:rPr lang="en-US" sz="{size}" b="{b}" i="{i}" dirty="0">'
            f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            f'<a:latin typeface="{font}"/><a:cs typeface="{font}"/>'
            f'</a:rPr>'
            f'<a:t>{text}</a:t>'
            f'</a:r>'
            f'</a:p>'
        )
    return "".join(paragraphs)


def textbox_shape(
    sp_id: int,
    name: str,
    x_in: float,
    y_in: float,
    w_in: float,
    h_in: float,
    text_xml: str,
    *,
    fill_color: str | None = None,
    anchor: str = "t",
    margin_pts: float = 0.1,
) -> str:
    fill = (
        f'<a:solidFill><a:srgbClr val="{fill_color}"/></a:solidFill>'
        if fill_color else '<a:noFill/>'
    )
    margin_emu = int(margin_pts * EMU_PER_INCH)
    return (
        f'<p:sp>'
        f'<p:nvSpPr><p:cNvPr id="{sp_id}" name="{name}"/>'
        f'<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr>'
        f'{_xfrm(emu_inches(x_in), emu_inches(y_in), emu_inches(w_in), emu_inches(h_in))}'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'{fill}'
        f'</p:spPr>'
        f'<p:txBody>'
        f'<a:bodyPr wrap="square" lIns="{margin_emu}" tIns="{margin_emu}" '
        f'rIns="{margin_emu}" bIns="{margin_emu}" anchor="{anchor}"><a:normAutofit/></a:bodyPr>'
        f'<a:lstStyle/>'
        f'{text_xml}'
        f'</p:txBody>'
        f'</p:sp>'
    )


def rect_shape(
    sp_id: int,
    name: str,
    x_in: float,
    y_in: float,
    w_in: float,
    h_in: float,
    fill_color: str,
    *,
    line_color: str | None = None,
    line_w: int = 0,
) -> str:
    line = '<a:ln><a:noFill/></a:ln>' if not line_color else (
        f'<a:ln w="{line_w}"><a:solidFill><a:srgbClr val="{line_color}"/></a:solidFill></a:ln>'
    )
    return (
        f'<p:sp>'
        f'<p:nvSpPr><p:cNvPr id="{sp_id}" name="{name}"/>'
        f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr>'
        f'{_xfrm(emu_inches(x_in), emu_inches(y_in), emu_inches(w_in), emu_inches(h_in))}'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{fill_color}"/></a:solidFill>'
        f'{line}'
        f'</p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        f'</p:sp>'
    )


def round_rect_shape(
    sp_id: int,
    name: str,
    x_in: float,
    y_in: float,
    w_in: float,
    h_in: float,
    fill_color: str,
    *,
    line_color: str | None = None,
    line_w: int = 0,
) -> str:
    line = '<a:ln><a:noFill/></a:ln>' if not line_color else (
        f'<a:ln w="{line_w}"><a:solidFill><a:srgbClr val="{line_color}"/></a:solidFill></a:ln>'
    )
    return (
        f'<p:sp>'
        f'<p:nvSpPr><p:cNvPr id="{sp_id}" name="{name}"/>'
        f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
        f'<p:spPr>'
        f'{_xfrm(emu_inches(x_in), emu_inches(y_in), emu_inches(w_in), emu_inches(h_in))}'
        f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{fill_color}"/></a:solidFill>'
        f'{line}'
        f'</p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        f'</p:sp>'
    )


def gradient_bg_xml(top_color: str, bottom_color: str) -> str:
    """A gradient slide background fill."""
    return (
        '<p:bg><p:bgPr>'
        '<a:gradFill rotWithShape="1">'
        '<a:gsLst>'
        f'<a:gs pos="0"><a:srgbClr val="{top_color}"/></a:gs>'
        f'<a:gs pos="100000"><a:srgbClr val="{bottom_color}"/></a:gs>'
        '</a:gsLst>'
        '<a:lin ang="5400000" scaled="1"/>'
        '</a:gradFill>'
        '<a:effectLst/>'
        '</p:bgPr></p:bg>'
    )


def solid_bg_xml(color: str) -> str:
    return (
        f'<p:bg><p:bgPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        '<a:effectLst/></p:bgPr></p:bg>'
    )


# ---------------------------------------------------------------------------
# Slide builders
# ---------------------------------------------------------------------------

SLIDE_OPEN = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
)

# Subtle but classy fade transition for every slide.
SLIDE_TRANSITION = '<p:transition spd="med" advClick="1"><p:fade/></p:transition>'


def _wrap_slide(body_xml: str, bg_xml: str) -> str:
    return (
        f'{SLIDE_OPEN}'
        '<p:cSld>'
        f'{bg_xml}'
        '<p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f'{body_xml}'
        '</p:spTree>'
        '</p:cSld>'
        f'{SLIDE_TRANSITION}'
        '</p:sld>'
    )


def build_title_slide() -> str:
    """Cover slide for the entire deck."""
    bg = gradient_bg_xml(NAVY, NAVY_DEEP)

    shapes = []
    # Top-left gold accent bar
    shapes.append(rect_shape(10, "TopBar", 0, 0, 13.333, 0.18, GOLD))
    # Bottom-right gold accent bar
    shapes.append(rect_shape(11, "BottomBar", 0, 7.32, 13.333, 0.18, GOLD))

    # Decorative thin gold lines around the title block
    shapes.append(rect_shape(12, "LineTop", 1.4, 2.05, 10.5, 0.04, GOLD))
    shapes.append(rect_shape(13, "LineBot", 1.4, 5.55, 10.5, 0.04, GOLD))

    # Tagline above main title
    shapes.append(textbox_shape(
        14, "Tagline", 1.4, 1.55, 10.5, 0.6,
        _runs_from_lines(["INTER BRANCH \u2022 GRAMMAR QUIZ \u2022 2026"],
                         size=2200, bold=True, color=GOLD, align="ctr"),
        anchor="ctr",
    ))

    # Main title
    shapes.append(textbox_shape(
        15, "MainTitle", 0.5, 2.25, 12.333, 1.7,
        _runs_from_lines(["Inter Branch English Grammar Quiz"],
                         size=6000, bold=True, color=WHITE, align="ctr",
                         line_spacing=100),
        anchor="ctr",
    ))

    # Subtitle: Classes XI & XII
    shapes.append(textbox_shape(
        16, "Subtitle", 0.5, 4.0, 12.333, 0.9,
        _runs_from_lines(["Classes XI & XII"],
                         size=4000, bold=False, color=GOLD_LIGHT, align="ctr"),
        anchor="ctr",
    ))

    # Date band
    shapes.append(round_rect_shape(17, "DateBand", 4.667, 4.95, 4.0, 0.6, GOLD))
    shapes.append(textbox_shape(
        18, "DateLabel", 4.667, 4.95, 4.0, 0.6,
        _runs_from_lines(["May 2026"], size=2400, bold=True, color=NAVY_DEEP, align="ctr"),
        anchor="ctr",
    ))

    # Footer
    shapes.append(textbox_shape(
        19, "Footer", 0.5, 6.7, 12.333, 0.5,
        _runs_from_lines(["Duration: 1 hour  \u2022  5 Rounds  \u2022  6 Teams per Class"],
                         size=1600, color=SOFT_WHITE, align="ctr"),
        anchor="ctr",
    ))

    return _wrap_slide("".join(shapes), bg)


def build_round_intro_slide(round_num: int, round_title: str) -> str:
    bg = gradient_bg_xml(NAVY, NAVY_DEEP)

    shapes = []
    # Side gold band on the left
    shapes.append(rect_shape(10, "SideBand", 0, 0, 0.35, 7.5, GOLD))

    # Faint "ROUND" label
    shapes.append(textbox_shape(
        11, "RoundLabel", 1.2, 2.1, 11.0, 0.7,
        _runs_from_lines([f"ROUND {round_num}"], size=3200, bold=True,
                         color=GOLD, align="l"),
        anchor="ctr",
    ))

    # Big round title
    shapes.append(textbox_shape(
        12, "RoundTitle", 1.2, 2.85, 11.0, 1.6,
        _runs_from_lines([round_title], size=7200, bold=True,
                         color=WHITE, align="l", line_spacing=100),
        anchor="ctr",
    ))

    # Underline accent under the round title
    shapes.append(rect_shape(13, "Underline", 1.2, 4.55, 3.5, 0.06, GOLD))

    # Helper sentence
    shapes.append(textbox_shape(
        14, "Helper", 1.2, 4.85, 11.0, 0.8,
        _runs_from_lines([
            "Class XI \u2192 Teams A to F",
            "Class XII \u2192 Teams A to F",
        ], size=2000, color=SOFT_WHITE, align="l", line_spacing=120),
        anchor="t",
    ))

    # Footer chip with round number
    shapes.append(round_rect_shape(15, "Chip", 11.0, 6.6, 1.83, 0.5, GOLD))
    shapes.append(textbox_shape(
        16, "ChipText", 11.0, 6.6, 1.83, 0.5,
        _runs_from_lines([f"Round {round_num} / 5"], size=1600, bold=True,
                         color=NAVY_DEEP, align="ctr"),
        anchor="ctr",
    ))

    return _wrap_slide("".join(shapes), bg)


def build_question_slide(
    round_num: int,
    round_title: str,
    class_name: str,
    team_name: str,
    question_number: int,
    question_text: str,
    hint: str,
    options,
) -> str:
    bg = solid_bg_xml(SOFT_WHITE)
    shapes = []

    # ---- Top header band (navy) ----
    shapes.append(rect_shape(10, "HeaderBand", 0, 0, 13.333, 1.05, NAVY))
    # Gold accent strip below header
    shapes.append(rect_shape(11, "HeaderAccent", 0, 1.05, 13.333, 0.08, GOLD))

    # Round name on left of header
    shapes.append(textbox_shape(
        12, "HeaderLeft", 0.5, 0.15, 8.5, 0.8,
        _runs_from_lines([f"Round {round_num} \u2014 {round_title}"],
                         size=2400, bold=True, color=WHITE, align="l"),
        anchor="ctr",
    ))

    # Class chip on right side of header
    shapes.append(round_rect_shape(13, "ClassChip", 9.5, 0.22, 3.4, 0.65, GOLD))
    shapes.append(textbox_shape(
        14, "ClassChipText", 9.5, 0.22, 3.4, 0.65,
        _runs_from_lines([f"Class {class_name}"], size=2000, bold=True,
                         color=NAVY_DEEP, align="ctr"),
        anchor="ctr",
    ))

    # ---- Sub-header strip with team and question number ----
    shapes.append(rect_shape(15, "SubHeader", 0, 1.13, 13.333, 0.7, WHITE))
    shapes.append(textbox_shape(
        16, "TeamLabel", 0.5, 1.18, 6.5, 0.6,
        _runs_from_lines([team_name], size=2400, bold=True, color=NAVY, align="l"),
        anchor="ctr",
    ))
    shapes.append(textbox_shape(
        17, "QnumLabel", 6.8, 1.18, 6.0, 0.6,
        _runs_from_lines([f"Question {question_number}"],
                         size=2000, bold=True, color=GOLD, align="r"),
        anchor="ctr",
    ))

    # ---- Question card ----
    card_x, card_y, card_w, card_h = 0.7, 2.05, 11.93, 3.5
    shapes.append(round_rect_shape(
        18, "QCard", card_x, card_y, card_w, card_h, WHITE,
        line_color=GOLD, line_w=19050,
    ))

    # Question icon - small gold square in the top-left of the card
    shapes.append(round_rect_shape(19, "QBadge", card_x + 0.3, card_y + 0.3, 0.7, 0.7, GOLD))
    shapes.append(textbox_shape(
        20, "QBadgeText", card_x + 0.3, card_y + 0.3, 0.7, 0.7,
        _runs_from_lines(["Q"], size=3200, bold=True, color=NAVY_DEEP, align="ctr"),
        anchor="ctr",
    ))

    # Question text - split on existing newlines so we keep the original layout
    q_lines = question_text.split("\n")
    shapes.append(textbox_shape(
        21, "QuestionText",
        card_x + 1.2, card_y + 0.25, card_w - 1.6, 1.9,
        _runs_from_lines(q_lines, size=2800, bold=False, color=INK,
                         align="l", line_spacing=120),
        anchor="t",
    ))

    # Hint (e.g. "(prepare)") in italics, just under the question
    if hint:
        shapes.append(textbox_shape(
            22, "QHint",
            card_x + 1.2, card_y + 2.1, card_w - 1.6, 0.5,
            _runs_from_lines([hint], size=2000, italic=True, color=NAVY, align="l"),
            anchor="t",
        ))

    # Options area
    if options:
        # Heading: "Options:"
        shapes.append(textbox_shape(
            23, "OptHeading",
            card_x + 1.2, card_y + 2.5, card_w - 1.6, 0.4,
            _runs_from_lines(["Options"], size=1800, bold=True, color=GOLD, align="l"),
            anchor="t",
        ))
        # Render options as bullet-style list
        opt_lines = [f"\u25C6  {opt}" for opt in options]
        shapes.append(textbox_shape(
            24, "OptList",
            card_x + 1.2, card_y + 2.85, card_w - 1.6, card_h - 2.95,
            _runs_from_lines(opt_lines, size=2000, color=INK,
                             align="l", line_spacing=130),
            anchor="t",
        ))

    # ---- Answer band at bottom ----
    ans_y = 5.85
    shapes.append(round_rect_shape(
        25, "AnswerBand", 0.7, ans_y, 11.93, 1.1, NAVY,
    ))
    shapes.append(rect_shape(26, "AnswerAccent", 0.7, ans_y, 0.18, 1.1, GOLD))
    shapes.append(textbox_shape(
        27, "AnswerLabel", 1.0, ans_y, 3.0, 1.1,
        _runs_from_lines(["Answer:"], size=2800, bold=True, color=GOLD, align="l"),
        anchor="ctr",
    ))
    shapes.append(textbox_shape(
        28, "AnswerLine", 4.0, ans_y, 8.6, 1.1,
        _runs_from_lines(["__________________________________"],
                         size=2400, color=SOFT_WHITE, align="l"),
        anchor="ctr",
    ))

    # Footer
    shapes.append(textbox_shape(
        29, "Footer", 0.5, 7.05, 12.333, 0.4,
        _runs_from_lines([
            f"Inter Branch English Grammar Quiz  \u2022  Class {class_name}  "
            f"\u2022  {team_name}  \u2022  Round {round_num}"
        ], size=1200, color=NAVY, align="ctr"),
        anchor="ctr",
    ))

    return _wrap_slide("".join(shapes), bg)


def build_thank_you_slide() -> str:
    bg = gradient_bg_xml(NAVY, NAVY_DEEP)
    shapes = []

    # Top and bottom gold bars
    shapes.append(rect_shape(10, "TopBar", 0, 0, 13.333, 0.18, GOLD))
    shapes.append(rect_shape(11, "BottomBar", 0, 7.32, 13.333, 0.18, GOLD))

    # Decorative diamond row
    shapes.append(textbox_shape(
        12, "Diamonds", 0.5, 1.6, 12.333, 0.6,
        _runs_from_lines(["\u25C6   \u25C6   \u25C6"],
                         size=2400, color=GOLD, align="ctr"),
        anchor="ctr",
    ))

    # Main "Thank You" text
    shapes.append(textbox_shape(
        13, "ThankYou", 0.5, 2.4, 12.333, 2.0,
        _runs_from_lines(["Thank You"], size=11000, bold=True,
                         color=WHITE, align="ctr"),
        anchor="ctr",
    ))

    # Sub-line
    shapes.append(textbox_shape(
        14, "SubLine", 0.5, 4.6, 12.333, 0.7,
        _runs_from_lines(["For your enthusiastic participation"],
                         size=2800, color=GOLD_LIGHT, align="ctr"),
        anchor="ctr",
    ))

    # Bottom small line
    shapes.append(textbox_shape(
        15, "BottomLine", 0.5, 5.6, 12.333, 0.5,
        _runs_from_lines(["Inter Branch English Grammar Quiz \u2022 May 2026"],
                         size=1800, color=SOFT_WHITE, align="ctr"),
        anchor="ctr",
    ))

    return _wrap_slide("".join(shapes), bg)


# ---------------------------------------------------------------------------
# Slide-list assembly
# ---------------------------------------------------------------------------

def build_all_slides():
    """Generate the ordered list of slide XML payloads."""
    slides = []

    # 1. Title slide
    slides.append(build_title_slide())

    # 2. Per round
    for r in ROUNDS:
        slides.append(build_round_intro_slide(r["number"], r["title"]))

        # Class XI then Class XII
        for class_name in ("XI", "XII"):
            team_questions = r["classes"][class_name]
            for team_idx, q in enumerate(team_questions):
                slides.append(build_question_slide(
                    round_num=r["number"],
                    round_title=r["title"],
                    class_name=class_name,
                    team_name=TEAM_NAMES[team_idx],
                    question_number=1,
                    question_text=q["question"],
                    hint=q["hint"],
                    options=q["options"],
                ))

    # 3. Thank-you slide
    slides.append(build_thank_you_slide())
    return slides


# ---------------------------------------------------------------------------
# Package writer
# ---------------------------------------------------------------------------

def write_pptx(out_path: str) -> None:
    slides = build_all_slides()
    n = len(slides)

    parts = {
        "[Content_Types].xml": content_types_xml(n),
        "_rels/.rels": ROOT_RELS,
        "docProps/core.xml": docprops_core(),
        "docProps/app.xml": docprops_app(n),
        "ppt/presentation.xml": presentation_xml(n),
        "ppt/_rels/presentation.xml.rels": presentation_rels(n),
        "ppt/presProps.xml": PRES_PROPS,
        "ppt/viewProps.xml": VIEW_PROPS,
        "ppt/tableStyles.xml": TABLE_STYLES,
        "ppt/theme/theme1.xml": theme_xml(),
        "ppt/slideMasters/slideMaster1.xml": slide_master_xml(),
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": SLIDE_MASTER_RELS,
        "ppt/slideLayouts/slideLayout1.xml": slide_layout_xml(),
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels": SLIDE_LAYOUT_RELS,
    }

    for i, slide_xml in enumerate(slides, start=1):
        parts[f"ppt/slides/slide{i}.xml"] = slide_xml
        parts[f"ppt/slides/_rels/slide{i}.xml.rels"] = slide_rels_xml()

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)

    print(f"Wrote {out_path} with {n} slides.")


if __name__ == "__main__":
    out = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "Inter_Branch_English_Grammar_Quiz.pptx",
    )
    write_pptx(out)
