"""Generate a professional .pptx for the Inter Branch English Grammar Quiz.

Uses only Python's standard library (zipfile + XML strings).
Produces a 16:9 deck: #052F61 background, #C7EA94 question card,
headings 36pt, instruction 40pt bold, question+options 48pt bold.
"""

from __future__ import annotations
import os
import zipfile
from xml.sax.saxutils import escape as xml_escape
from quiz_data import ROUNDS, TEAM_NAMES

# ---------------------------------------------------------------------------
# Theme constants
# ---------------------------------------------------------------------------
SLIDE_W = 12192000
SLIDE_H = 6858000
EMU_PER_INCH = 914400

NAVY = "0F2A5C"
NAVY_DEEP = "081A3C"
GOLD = "D4AF37"
GOLD_LIGHT = "F0D27A"
WHITE = "FFFFFF"
SOFT_WHITE = "F5F7FB"
INK = "0B1B3A"
QSLIDE_BG = "052F61"
QCARD_BG = "C7EA94"

def emu(inches: float) -> int:
    return int(round(inches * EMU_PER_INCH))



# ---------------------------------------------------------------------------
# XML helper functions
# ---------------------------------------------------------------------------

def _xfrm(x, y, cx, cy):
    return f'<a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'


def _run(text, *, size, bold=False, italic=False, color=INK, font="Calibri"):
    b = "1" if bold else "0"
    i = "1" if italic else "0"
    t = xml_escape(text)
    return (
        f'<a:r><a:rPr lang="en-US" sz="{size}" b="{b}" i="{i}" dirty="0">'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:latin typeface="{font}"/><a:cs typeface="{font}"/>'
        f'</a:rPr><a:t>{t}</a:t></a:r>'
    )


def _para(runs_xml, *, align="l", spc=110):
    return (
        f'<a:p><a:pPr algn="{align}">'
        f'<a:lnSpc><a:spcPct val="{spc * 1000}"/></a:lnSpc>'
        f'</a:pPr>{runs_xml}</a:p>'
    )


def _empty_para(size=1600):
    return f'<a:p><a:endParaRPr lang="en-US" sz="{size}"/></a:p>'



def textbox(sp_id, name, x_in, y_in, w_in, h_in, body_xml, *,
            fill=None, anchor="t", margin=0.1):
    f = f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>' if fill else '<a:noFill/>'
    m = int(margin * EMU_PER_INCH)
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{sp_id}" name="{name}"/>'
        f'<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>'
        f'{_xfrm(emu(x_in), emu(y_in), emu(w_in), emu(h_in))}'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{f}</p:spPr>'
        f'<p:txBody><a:bodyPr wrap="square" lIns="{m}" tIns="{m}" rIns="{m}" bIns="{m}" '
        f'anchor="{anchor}"><a:normAutofit/></a:bodyPr><a:lstStyle/>{body_xml}</p:txBody></p:sp>'
    )


def rect(sp_id, name, x_in, y_in, w_in, h_in, color, *, line_color=None, line_w=0):
    ln = ('<a:ln><a:noFill/></a:ln>' if not line_color else
          f'<a:ln w="{line_w}"><a:solidFill><a:srgbClr val="{line_color}"/></a:solidFill></a:ln>')
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{sp_id}" name="{name}"/>'
        f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>'
        f'{_xfrm(emu(x_in), emu(y_in), emu(w_in), emu(h_in))}'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>{ln}</p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
    )


def rrect(sp_id, name, x_in, y_in, w_in, h_in, color, *, line_color=None, line_w=0):
    ln = ('<a:ln><a:noFill/></a:ln>' if not line_color else
          f'<a:ln w="{line_w}"><a:solidFill><a:srgbClr val="{line_color}"/></a:solidFill></a:ln>')
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{sp_id}" name="{name}"/>'
        f'<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>'
        f'{_xfrm(emu(x_in), emu(y_in), emu(w_in), emu(h_in))}'
        f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>{ln}</p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
    )



def solid_bg(color):
    return (f'<p:bg><p:bgPr><a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
            '<a:effectLst/></p:bgPr></p:bg>')

def gradient_bg(top, bot):
    return (
        '<p:bg><p:bgPr><a:gradFill rotWithShape="1"><a:gsLst>'
        f'<a:gs pos="0"><a:srgbClr val="{top}"/></a:gs>'
        f'<a:gs pos="100000"><a:srgbClr val="{bot}"/></a:gs>'
        '</a:gsLst><a:lin ang="5400000" scaled="1"/></a:gradFill>'
        '<a:effectLst/></p:bgPr></p:bg>'
    )

SLIDE_OPEN = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
)
TRANSITION = '<p:transition spd="med" advClick="1"><p:fade/></p:transition>'

def wrap_slide(shapes_xml, bg_xml):
    return (
        f'{SLIDE_OPEN}<p:cSld>{bg_xml}<p:spTree>'
        '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
        '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        f'{shapes_xml}</p:spTree></p:cSld>{TRANSITION}</p:sld>'
    )



# ---------------------------------------------------------------------------
# Slide builders
# ---------------------------------------------------------------------------

def build_title_slide():
    bg = gradient_bg(NAVY, NAVY_DEEP)
    s = []
    s.append(rect(10, "TopBar", 0, 0, 13.333, 0.18, GOLD))
    s.append(rect(11, "BotBar", 0, 7.32, 13.333, 0.18, GOLD))
    s.append(rect(12, "LineT", 1.4, 2.05, 10.5, 0.04, GOLD))
    s.append(rect(13, "LineB", 1.4, 5.55, 10.5, 0.04, GOLD))
    s.append(textbox(14, "Tag", 1.4, 1.55, 10.5, 0.6,
        _para(_run("INTER BRANCH \u2022 GRAMMAR QUIZ \u2022 2026", size=2200, bold=True, color=GOLD), align="ctr"),
        anchor="ctr"))
    s.append(textbox(15, "Title", 0.5, 2.25, 12.333, 1.7,
        _para(_run("Inter Branch English Grammar Quiz", size=6000, bold=True, color=WHITE), align="ctr"),
        anchor="ctr"))
    s.append(textbox(16, "Sub", 0.5, 4.0, 12.333, 0.9,
        _para(_run("Classes XI & XII", size=4000, color=GOLD_LIGHT), align="ctr"),
        anchor="ctr"))
    s.append(rrect(17, "DateBand", 4.667, 4.95, 4.0, 0.6, GOLD))
    s.append(textbox(18, "Date", 4.667, 4.95, 4.0, 0.6,
        _para(_run("May 2026", size=2400, bold=True, color=NAVY_DEEP), align="ctr"),
        anchor="ctr"))
    s.append(textbox(19, "Foot", 0.5, 6.7, 12.333, 0.5,
        _para(_run("Duration: 1 hour  \u2022  5 Rounds  \u2022  6 Teams per Class", size=1600, color=SOFT_WHITE), align="ctr"),
        anchor="ctr"))
    return wrap_slide("".join(s), bg)



def build_round_intro(round_num, round_title):
    bg = gradient_bg(NAVY, NAVY_DEEP)
    s = []
    s.append(rect(10, "Side", 0, 0, 0.35, 7.5, GOLD))
    s.append(textbox(11, "RLbl", 1.2, 2.1, 11.0, 0.7,
        _para(_run(f"ROUND {round_num}", size=3200, bold=True, color=GOLD), align="l"),
        anchor="ctr"))
    s.append(textbox(12, "RTitle", 1.2, 2.85, 11.0, 1.6,
        _para(_run(round_title, size=7200, bold=True, color=WHITE), align="l"),
        anchor="ctr"))
    s.append(rect(13, "UL", 1.2, 4.55, 3.5, 0.06, GOLD))
    s.append(textbox(14, "Help", 1.2, 4.85, 11.0, 0.8,
        _para(_run("Class XI \u2192 Teams A to F", size=2000, color=SOFT_WHITE), align="l")
        + _para(_run("Class XII \u2192 Teams A to F", size=2000, color=SOFT_WHITE), align="l"),
        anchor="t"))
    s.append(rrect(15, "Chip", 11.0, 6.6, 1.83, 0.5, GOLD))
    s.append(textbox(16, "ChipT", 11.0, 6.6, 1.83, 0.5,
        _para(_run(f"Round {round_num} / 5", size=1600, bold=True, color=NAVY_DEEP), align="ctr"),
        anchor="ctr"))
    return wrap_slide("".join(s), bg)



def build_question_slide(round_num, round_title, class_name, team_name,
                         q_num, instruction, question, hint, options, options_inline):
    """Build a question slide matching the reference screenshots.

    Font sizes:
    - Header (round title, class chip, team, question#): 36pt bold
    - Instruction line: 40pt bold
    - Question text: 48pt bold
    - Hint: 48pt bold italic
    - "Options" word: 40pt bold
    - Option items: 48pt bold
    """
    bg = solid_bg(QSLIDE_BG)
    s = []

    # ---- Header band (dark navy) ----
    s.append(rect(10, "HBand", 0, 0, 13.333, 1.0, NAVY_DEEP))
    s.append(rect(11, "HAcc", 0, 1.0, 13.333, 0.07, GOLD))

    # Round title in header - 36pt
    s.append(textbox(12, "HLeft", 0.4, 0.12, 8.5, 0.8,
        _para(_run(f"Round {round_num} \u2014 {round_title}", size=3600, bold=True, color=WHITE), align="l"),
        anchor="ctr"))

    # Class chip - 36pt
    s.append(rrect(13, "CChip", 9.5, 0.18, 3.4, 0.65, GOLD))
    s.append(textbox(14, "CChipT", 9.5, 0.18, 3.4, 0.65,
        _para(_run(f"Class {class_name}", size=3600, bold=True, color=NAVY_DEEP), align="ctr"),
        anchor="ctr"))

    # ---- Sub-header (team + question#) on dark bg ----
    s.append(rect(15, "SubH", 0, 1.07, 13.333, 0.7, QSLIDE_BG))
    s.append(textbox(16, "Team", 0.4, 1.12, 6.5, 0.6,
        _para(_run(team_name, size=3600, bold=True, color=WHITE), align="l"),
        anchor="ctr"))
    s.append(textbox(17, "QNum", 6.8, 1.12, 6.1, 0.6,
        _para(_run(f"Question {q_num}", size=3600, bold=True, color=GOLD), align="r"),
        anchor="ctr"))

    # ---- Question card (light green, rounded rect) ----
    cx, cy, cw, ch = 0.35, 1.85, 12.62, 5.55
    s.append(rrect(18, "QCard", cx, cy, cw, ch, QCARD_BG, line_color=GOLD, line_w=19050))

    # Q badge
    s.append(rrect(19, "QBdg", cx + 0.25, cy + 0.25, 0.8, 0.8, GOLD))
    s.append(textbox(20, "QBdgT", cx + 0.25, cy + 0.25, 0.8, 0.8,
        _para(_run("Q", size=3600, bold=True, color=NAVY_DEEP), align="ctr"),
        anchor="ctr"))

    # Content area starts after badge
    tx = cx + 1.3
    tw = cw - 1.6
    cur_y = cy + 0.25

    # Build the body paragraphs for the content area
    body_paras = []

    # Instruction line(s) - 40pt bold
    if instruction:
        for iline in instruction.split("\n"):
            body_paras.append(
                _para(_run(iline, size=4000, bold=True, color=INK), align="l", spc=115)
            )

    # Question text - 48pt bold
    if question:
        for qline in question.split("\n"):
            body_paras.append(
                _para(_run(qline, size=4800, bold=True, color=INK), align="l", spc=115)
            )

    # Hint - 48pt bold italic
    if hint:
        body_paras.append(
            _para(_run(hint, size=4800, bold=True, italic=True, color=NAVY_DEEP), align="l", spc=115)
        )

    # Options
    if options:
        if options_inline:
            # Single line: "Options: (opt1 / opt2 / opt3)"
            opts_text = "Options: (" + " / ".join(options) + ")"
            body_paras.append(
                _para(_run(opts_text, size=4000, bold=True, color=NAVY_DEEP), align="l", spc=115)
            )
        else:
            # "Options" heading at 40pt
            body_paras.append(
                _para(_run("Options", size=4000, bold=True, color=NAVY_DEEP), align="l", spc=115)
            )
            # Each option as bullet at 48pt bold
            for opt in options:
                body_paras.append(
                    _para(_run(f"\u25C6  {opt}", size=4800, bold=True, color=INK), align="l", spc=120)
                )

    content_h = ch - 0.4
    s.append(textbox(21, "QContent", tx, cur_y, tw, content_h,
        "".join(body_paras), anchor="t", margin=0.15))

    return wrap_slide("".join(s), bg)



def build_thank_you():
    bg = gradient_bg(NAVY, NAVY_DEEP)
    s = []
    s.append(rect(10, "TopBar", 0, 0, 13.333, 0.18, GOLD))
    s.append(rect(11, "BotBar", 0, 7.32, 13.333, 0.18, GOLD))
    s.append(textbox(12, "Dia", 0.5, 1.6, 12.333, 0.6,
        _para(_run("\u25C6   \u25C6   \u25C6", size=2400, color=GOLD), align="ctr"),
        anchor="ctr"))
    s.append(textbox(13, "TY", 0.5, 2.4, 12.333, 2.0,
        _para(_run("Thank You", size=11000, bold=True, color=WHITE), align="ctr"),
        anchor="ctr"))
    s.append(textbox(14, "Sub", 0.5, 4.6, 12.333, 0.7,
        _para(_run("For your enthusiastic participation", size=2800, color=GOLD_LIGHT), align="ctr"),
        anchor="ctr"))
    s.append(textbox(15, "Bot", 0.5, 5.6, 12.333, 0.5,
        _para(_run("Inter Branch English Grammar Quiz \u2022 May 2026", size=1800, color=SOFT_WHITE), align="ctr"),
        anchor="ctr"))
    return wrap_slide("".join(s), bg)



# ---------------------------------------------------------------------------
# Slide assembly
# ---------------------------------------------------------------------------

def build_all_slides():
    slides = [build_title_slide()]
    for r in ROUNDS:
        slides.append(build_round_intro(r["number"], r["title"]))
        for cls in ("XI", "XII"):
            for tidx, q in enumerate(r["classes"][cls]):
                slides.append(build_question_slide(
                    round_num=r["number"],
                    round_title=r["title"],
                    class_name=cls,
                    team_name=TEAM_NAMES[tidx],
                    q_num=1,
                    instruction=q["instruction"],
                    question=q["question"],
                    hint=q["hint"],
                    options=q["options"],
                    options_inline=q.get("options_inline", False),
                ))
    slides.append(build_thank_you())
    return slides



# ---------------------------------------------------------------------------
# Package XML parts
# ---------------------------------------------------------------------------

def content_types(n):
    ov = "".join(
        f'<Override PartName="/ppt/slides/slide{i}.xml" '
        f'ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'
        for i in range(1, n+1))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>'
        '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>'
        '<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>'
        '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>'
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        f'{ov}</Types>'
    )



ROOT_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
    '</Relationships>'
)

CORE = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" '
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    '<dc:title>Inter Branch English Grammar Quiz</dc:title>'
    '<dc:creator>Quiz Master</dc:creator>'
    '<dcterms:created xsi:type="dcterms:W3CDTF">2026-05-01T09:00:00Z</dcterms:created>'
    '<dcterms:modified xsi:type="dcterms:W3CDTF">2026-05-01T09:00:00Z</dcterms:modified>'
    '</cp:coreProperties>'
)

def app_xml(n):
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
        f'<Application>Kiro</Application><Slides>{n}</Slides></Properties>'
    )



def theme():
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Quiz">
<a:themeElements>
<a:clrScheme name="Q"><a:dk1><a:srgbClr val="{INK}"/></a:dk1><a:lt1><a:srgbClr val="{WHITE}"/></a:lt1>
<a:dk2><a:srgbClr val="{NAVY}"/></a:dk2><a:lt2><a:srgbClr val="{SOFT_WHITE}"/></a:lt2>
<a:accent1><a:srgbClr val="{GOLD}"/></a:accent1><a:accent2><a:srgbClr val="{NAVY_DEEP}"/></a:accent2>
<a:accent3><a:srgbClr val="{GOLD_LIGHT}"/></a:accent3><a:accent4><a:srgbClr val="2C4A7E"/></a:accent4>
<a:accent5><a:srgbClr val="B0883E"/></a:accent5><a:accent6><a:srgbClr val="DCE3F2"/></a:accent6>
<a:hlink><a:srgbClr val="{GOLD}"/></a:hlink><a:folHlink><a:srgbClr val="B0883E"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Q"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>
<a:fmtScheme name="Q"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln>
<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>
</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'''



def slide_master():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
        '<p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
        '<a:effectLst/></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/>'
        '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/>'
        '<a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm>'
        '</p:grpSpPr></p:spTree></p:cSld>'
        '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" '
        'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
        '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>'
        '<p:transition spd="med"><p:fade/></p:transition>'
        '<p:txStyles><p:titleStyle><a:lvl1pPr algn="l"><a:defRPr sz="4400" b="1"/></a:lvl1pPr></p:titleStyle>'
        '<p:bodyStyle><a:lvl1pPr><a:defRPr sz="2400"/></a:lvl1pPr></p:bodyStyle><p:otherStyle/></p:txStyles>'
        '</p:sldMaster>'
    )

SM_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
    '</Relationships>'
)

def slide_layout():
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">'
        '<p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
        '</p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>'
        '<p:transition spd="med"><p:fade/></p:transition></p:sldLayout>'
    )

SL_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
    '</Relationships>'
)

SLIDE_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
    '</Relationships>'
)



PRES_PROPS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>'
VIEW_PROPS = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr><p:cViewPr varScale="1"><p:scale><a:sx n="68" d="100"/><a:sy n="68" d="100"/></p:scale><p:origin x="-1500" y="-90"/></p:cViewPr></p:cSldViewPr></p:slideViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>'
TABLE_STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>'


def presentation_xml(n):
    ids = "".join(f'<p:sldId id="{256+i}" r:id="rId{i+2}"/>' for i in range(n))
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" saveSubsetFonts="1">'
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
        f'<p:sldIdLst>{ids}</p:sldIdLst>'
        f'<p:sldSz cx="{SLIDE_W}" cy="{SLIDE_H}" type="screen16x9"/>'
        '<p:notesSz cx="6858000" cy="9144000"/>'
        '<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></p:defaultTextStyle>'
        '</p:presentation>'
    )


def pres_rels(n):
    r = ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>']
    for i in range(1, n+1):
        r.append(f'<Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i}.xml"/>')
    nxt = n + 2
    r.append(f'<Relationship Id="rId{nxt}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>')
    r.append(f'<Relationship Id="rId{nxt+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>')
    r.append(f'<Relationship Id="rId{nxt+2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>')
    r.append(f'<Relationship Id="rId{nxt+3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            + "".join(r) + '</Relationships>')



# ---------------------------------------------------------------------------
# Write the .pptx
# ---------------------------------------------------------------------------

def write_pptx(out_path):
    slides = build_all_slides()
    n = len(slides)
    parts = {
        "[Content_Types].xml": content_types(n),
        "_rels/.rels": ROOT_RELS,
        "docProps/core.xml": CORE,
        "docProps/app.xml": app_xml(n),
        "ppt/presentation.xml": presentation_xml(n),
        "ppt/_rels/presentation.xml.rels": pres_rels(n),
        "ppt/presProps.xml": PRES_PROPS,
        "ppt/viewProps.xml": VIEW_PROPS,
        "ppt/tableStyles.xml": TABLE_STYLES,
        "ppt/theme/theme1.xml": theme(),
        "ppt/slideMasters/slideMaster1.xml": slide_master(),
        "ppt/slideMasters/_rels/slideMaster1.xml.rels": SM_RELS,
        "ppt/slideLayouts/slideLayout1.xml": slide_layout(),
        "ppt/slideLayouts/_rels/slideLayout1.xml.rels": SL_RELS,
    }
    for i, sxml in enumerate(slides, 1):
        parts[f"ppt/slides/slide{i}.xml"] = sxml
        parts[f"ppt/slides/_rels/slide{i}.xml.rels"] = SLIDE_RELS

    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
    print(f"Wrote {out_path} with {n} slides.")


if __name__ == "__main__":
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                       "Inter_Branch_English_Grammar_Quiz.pptx")
    write_pptx(out)
