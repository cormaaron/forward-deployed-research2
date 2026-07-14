#!/usr/bin/env python3
"""Builds the forward-deployed engineering executive deck (.pptx)."""

import copy

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Emu, Inches, Pt

# ---------------------------------------------------------------- palette
NAVY = RGBColor(0x10, 0x2A, 0x43)
BLUE = RGBColor(0x2E, 0x74, 0xB5)
LIGHT_BLUE = RGBColor(0xE9, 0xF1, 0xF8)
INK = RGBColor(0x33, 0x38, 0x3D)
GREY = RGBColor(0x63, 0x6A, 0x71)
MID_GREY = RGBColor(0x9A, 0xA0, 0xA6)
HAIRLINE = RGBColor(0xC9, 0xCE, 0xD3)
CARD = RGBColor(0xF4, 0xF6, 0xF8)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)

FONT = "Arial"

PAGE_W = Inches(13.333)
PAGE_H = Inches(7.5)
MARGIN = Inches(0.55)
CONTENT_W = Inches(12.233)

prs = Presentation()
prs.slide_width = PAGE_W
prs.slide_height = PAGE_H
BLANK = prs.slide_layouts[6]


# ---------------------------------------------------------------- helpers
def _style_run(run, size, color, bold=False, italic=False):
    f = run.font
    f.name = FONT
    f.size = Pt(size)
    f.color.rgb = color
    f.bold = bold
    f.italic = italic


def add_text(slide, x, y, w, h, runs, size=10, color=INK, bold=False,
             align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, space_after=0,
             line_spacing=1.0, wrap=True):
    """runs: str, or list of paragraphs; each paragraph is a list of
    (text, {overrides}) tuples."""
    box = slide.shapes.add_textbox(x, y, w, h)
    tf = box.text_frame
    tf.word_wrap = wrap
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    if isinstance(runs, str):
        runs = [[(runs, {})]]
    first = True
    for para in runs:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.line_spacing = line_spacing
        p.space_after = Pt(space_after)
        for text, ov in para:
            r = p.add_run()
            r.text = text
            _style_run(r, ov.get("size", size), ov.get("color", color),
                       ov.get("bold", bold), ov.get("italic", False))
    return box


def add_rect(slide, x, y, w, h, fill=None, line=None, line_w=0.75,
             shape=MSO_SHAPE.RECTANGLE):
    sp = slide.shapes.add_shape(shape, x, y, w, h)
    sp.shadow.inherit = False
    style_el = sp._element.find(qn("p:style"))
    if style_el is not None:
        sp._element.remove(style_el)
    if fill is None:
        sp.fill.background()
    else:
        sp.fill.solid()
        sp.fill.fore_color.rgb = fill
    if line is None:
        sp.line.fill.background()
    else:
        sp.line.color.rgb = line
        sp.line.width = Pt(line_w)
    sp.text_frame.margin_left = sp.text_frame.margin_right = Inches(0.06)
    sp.text_frame.margin_top = sp.text_frame.margin_bottom = Inches(0.02)
    return sp


def shape_text(sp, runs, size=9.5, color=WHITE, bold=False,
               align=PP_ALIGN.CENTER, line_spacing=1.0):
    tf = sp.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    if isinstance(runs, str):
        runs = [[(runs, {})]]
    first = True
    for para in runs:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.line_spacing = line_spacing
        for text, ov in para:
            r = p.add_run()
            r.text = text
            _style_run(r, ov.get("size", size), ov.get("color", color),
                       ov.get("bold", bold))


def add_line(slide, x, y, w, color=HAIRLINE, weight=0.75):
    ln = slide.shapes.add_connector(1, x, y, x + w, y)
    ln.line.color.rgb = color
    ln.line.width = Pt(weight)
    ln.shadow.inherit = False
    style_el = ln._element.find(qn("p:style"))
    if style_el is not None:
        ln._element.remove(style_el)
    return ln


def header(slide, title, subtitle=None):
    add_text(slide, MARGIN, Inches(0.34), CONTENT_W, Inches(0.9),
             title, size=20, color=NAVY, bold=True, line_spacing=1.05)
    y = Inches(1.22)
    if subtitle:
        add_text(slide, MARGIN, Inches(1.06), CONTENT_W, Inches(0.5),
                 subtitle, size=11, color=GREY, line_spacing=1.15)
        y = Inches(1.56)
    add_line(slide, MARGIN, y, CONTENT_W, color=NAVY, weight=1.4)
    return y


def footer(slide, page_no, source=None):
    add_line(slide, MARGIN, Inches(7.12), CONTENT_W, color=HAIRLINE, weight=0.5)
    if source:
        add_text(slide, MARGIN, Inches(7.2), Inches(10.5), Inches(0.25),
                 source, size=7.5, color=MID_GREY)
    add_text(slide, Inches(12.28), Inches(7.2), Inches(0.5), Inches(0.25),
             str(page_no), size=8, color=MID_GREY, align=PP_ALIGN.RIGHT)


def section_head(slide, x, y, w, text):
    add_text(slide, x, y, w, Inches(0.3), text.upper(),
             size=10.5, color=NAVY, bold=True)
    add_line(slide, x, y + Inches(0.28), Inches(0.42), color=BLUE, weight=2.25)


def takeaway_box(slide, x, y, w, h, lead, body, body_size=9.5):
    add_rect(slide, x, y, w, h, fill=LIGHT_BLUE)
    add_rect(slide, x, y, Inches(0.045), h, fill=NAVY)
    add_text(slide, x + Inches(0.2), y, w - Inches(0.4), h,
             [[(lead + "  ", {"bold": True, "color": NAVY}),
               (body, {"color": INK})]],
             size=body_size, anchor=MSO_ANCHOR.MIDDLE, line_spacing=1.12)


def set_cell(cell, runs, size=8.5, color=INK, bold=False,
             align=PP_ALIGN.LEFT, fill=None):
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    else:
        cell.fill.solid()
        cell.fill.fore_color.rgb = WHITE
    cell.margin_left = Inches(0.08)
    cell.margin_right = Inches(0.08)
    cell.margin_top = Inches(0.045)
    cell.margin_bottom = Inches(0.045)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = cell.text_frame
    tf.word_wrap = True
    if isinstance(runs, str):
        runs = [[(runs, {})]]
    first = True
    for para in runs:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.line_spacing = 1.05
        for text, ov in para:
            r = p.add_run()
            r.text = text
            _style_run(r, ov.get("size", size), ov.get("color", color),
                       ov.get("bold", bold))


def plain_table_style(table):
    tbl = table._tbl
    tbl_pr = tbl.tblPr
    tbl_pr.set("firstRow", "0")
    tbl_pr.set("bandRow", "0")
    for el in tbl_pr.findall(qn("a:tableStyleId")):
        tbl_pr.remove(el)
    style_el = tbl_pr.makeelement(qn("a:tableStyleId"), {})
    style_el.text = "{5940675A-B579-460E-94D1-54222C63F5DA}"  # no style, no grid
    tbl_pr.append(style_el)


def cell_bottom_border(cell, color=HAIRLINE, weight=0.5):
    tc_pr = cell._tc.get_or_add_tcPr()
    for tag in ("a:lnB",):
        for el in tc_pr.findall(qn(tag)):
            tc_pr.remove(el)
    ln = tc_pr.makeelement(qn("a:lnB"), {
        "w": str(Emu(Pt(weight)).emu if hasattr(Emu(0), 'emu') else int(Pt(weight))),
        "cap": "flat"})
    ln.set("w", str(int(Pt(weight))))
    fill = tc_pr.makeelement(qn("a:solidFill"), {})
    clr = tc_pr.makeelement(qn("a:srgbClr"), {"val": "%02X%02X%02X" % (color[0], color[1], color[2])})
    fill.append(clr)
    ln.append(fill)
    tc_pr.append(ln)


def bullet(slide, x, y, w, text, size=10.5, color=INK, gap=0.115):
    add_rect(slide, x, y + Inches(0.055), Inches(0.075), Inches(0.075), fill=BLUE)
    add_text(slide, x + Inches(0.2), y, w - Inches(0.2), Inches(0.4),
             text, size=size, color=color, line_spacing=1.1)


# ================================================================ SLIDE 1
s = prs.slides.add_slide(BLANK)
header(
    s,
    "Forward-deployed engineering is emerging as the leading operating model "
    "for scaling enterprise AI",
    "Leading AI-native organizations are embedding engineers directly within "
    "business teams to accelerate workflow automation, increase adoption and "
    "continuously improve enterprise platforms.")

LX, LW = MARGIN, Inches(4.35)
RX, RW = Inches(5.3), Inches(7.48)

section_head(s, LX, Inches(1.78), LW, "What differentiates the model")

diff_points = [
    ("Embedded in the business", "Engineers sit inside business teams, not central IT"),
    ("Problem-led, not spec-led", "Work starts from operational problems, not requirements"),
    ("Rapid iteration", "Solutions iterated live with users in days, not quarters"),
    ("Productization", "Proven solutions become reusable platform capabilities"),
]
py = Inches(2.3)
for i, (lead, body) in enumerate(diff_points):
    add_text(s, LX, py, Inches(0.5), Inches(0.4), "0%d" % (i + 1),
             size=15, color=BLUE, bold=True)
    add_text(s, LX + Inches(0.52), py - Inches(0.01), LW - Inches(0.52), Inches(0.6),
             [[(lead, {"bold": True, "color": NAVY})],
              [(body, {"color": GREY, "size": 9.5})]],
             size=10.5, line_spacing=1.12)
    if i < 3:
        add_line(s, LX, py + Inches(0.72), LW, color=HAIRLINE, weight=0.5)
    py += Inches(0.92)

section_head(s, RX, Inches(1.78), RW, "Evidence from leading organizations")

rows = [
    ("Palantir",
     "Originator of the FDE model; engineers deploy on-site to build on Foundry and AIP",
     "Credits FDE-led AIP bootcamps for accelerating US commercial growth on earnings calls"),
    ("Uber",
     "Engineers embedded with city operations teams; model now offered externally via Uber AI Solutions",
     "Embedded ops–engineering pairing publicly cited as central to city-by-city scaling"),
    ("OpenAI",
     "FDE teams embed with enterprise customers to build agentic deployments",
     "Positions FDEs as core to converting frontier models into enterprise adoption"),
    ("ElevenLabs",
     "FDEs co-build production voice agents with customer engineering teams",
     "Attributes enterprise expansion to its deployment-led engineering model"),
    ("Stripe",
     "Engineers work directly with users; solution engineers embedded in key accounts",
     "User-proximate engineering credited for product velocity and enterprise wins"),
    ("Brex",
     "AI engineers embedded in internal teams to redesign core workflows",
     "Leadership reports company-wide daily AI usage from the embedded rollout"),
    ("Plaid",
     "FDEs work alongside customer engineers during onboarding and expansion",
     "Publicly cites materially faster enterprise integrations"),
]

tbl_y = Inches(2.24)
tbl_h = Inches(3.72)
gframe = s.shapes.add_table(len(rows) + 1, 3, RX, tbl_y, RW, tbl_h)
table = gframe.table
plain_table_style(table)
table.columns[0].width = Inches(1.08)
table.columns[1].width = Inches(3.1)
table.columns[2].width = Inches(3.3)
table.rows[0].height = Inches(0.3)
for i in range(1, len(rows) + 1):
    table.rows[i].height = Inches(0.487)

for c, label in enumerate(("Company", "Operating model", "Publicly reported outcomes")):
    set_cell(table.cell(0, c), label, size=9, color=WHITE, bold=True, fill=NAVY)
for r, (name, model, outcome) in enumerate(rows, start=1):
    fill = WHITE if r % 2 else CARD
    set_cell(table.cell(r, 0), name, size=8.5, color=NAVY, bold=True, fill=fill)
    set_cell(table.cell(r, 1), model, size=8, color=INK, fill=fill)
    set_cell(table.cell(r, 2), outcome, size=8, color=GREY, fill=fill)

takeaway_box(
    s, RX, Inches(6.14), RW, Inches(0.76),
    "Common pattern:",
    "every organization positions engineers at the point of value creation "
    "and relies on a platform team to convert local wins into reusable "
    "enterprise assets.")

footer(s, 1, "Source: Company earnings calls, engineering blogs and public statements. Outcomes as publicly reported; not independently verified.")

# ================================================================ SLIDE 2
s = prs.slides.add_slide(BLANK)
header(
    s,
    "Leading organizations consistently deploy the same operating model "
    "despite differences in industry")

LX, LW = MARGIN, Inches(4.5)
RX = Inches(5.45)

section_head(s, LX, Inches(1.42), LW, "Common design principles")
principles = [
    "Embed engineers alongside business operators",
    "Observe workflows first-hand before building",
    "Deploy small, multidisciplinary pods",
    "Run rapid build–measure–learn cycles",
    "Productize successful solutions",
    "Continuously strengthen the enterprise platform",
]
cy = Inches(1.94)
for i, ptext in enumerate(principles):
    card = add_rect(s, LX, cy, LW, Inches(0.66), fill=CARD)
    add_rect(s, LX, cy, Inches(0.04), Inches(0.66), fill=BLUE)
    add_text(s, LX + Inches(0.18), cy, Inches(0.5), Inches(0.66),
             str(i + 1), size=17, color=NAVY, bold=True,
             anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, LX + Inches(0.68), cy, LW - Inches(0.85), Inches(0.66),
             ptext, size=10, color=INK, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.1)
    cy += Inches(0.79)

section_head(s, RX, Inches(1.42), Inches(7.3), "How each organization applies it")

companies = [
    ("Palantir",
     "On-site FDEs building on Foundry and AIP",
     "Two-role split — domain specialists and product engineers",
     "Land and expand enterprise accounts"),
    ("Uber",
     "Engineers inside city operations teams",
     "Ops data loop tuned market by market",
     "Speed of launch and scale per city"),
    ("OpenAI",
     "FDE pods within enterprise accounts",
     "Proximity to frontier research",
     "Production agent deployments"),
    ("ElevenLabs",
     "FDEs co-build with client engineers",
     "Deep voice and audio specialization",
     "Fastest path to live voice agents"),
    ("Stripe",
     "Engineers paired directly with users",
     "Developer empathy and API craft",
     "Product velocity from user contact"),
    ("Brex",
     "AI engineers embedded in internal teams",
     "Internal-first automation of own workflows",
     "Company-wide productivity gains"),
    ("Plaid",
     "FDEs inside customer integration teams",
     "Deep bank and fintech data expertise",
     "Faster, stickier integrations"),
]

CW, CH = Inches(3.63), Inches(1.12)
GX, GY = Inches(0.17), Inches(0.115)
labels = ("Model", "Edge", "Goal")
for i, (name, model, edge, goal) in enumerate(companies):
    col, row = i % 2, i // 2
    x = RX + col * (CW + GX)
    y = Inches(1.94) + row * (CH + GY)
    add_rect(s, x, y, CW, CH, fill=WHITE, line=HAIRLINE, line_w=0.75)
    add_rect(s, x, y, CW, Inches(0.02), fill=NAVY)
    add_text(s, x + Inches(0.14), y + Inches(0.08), CW - Inches(0.28), Inches(0.22),
             name, size=9.5, color=NAVY, bold=True)
    paras = []
    for label, val in zip(labels, (model, edge, goal)):
        paras.append([(label + "   ", {"bold": True, "color": MID_GREY, "size": 7}),
                      (val, {"color": INK})])
    add_text(s, x + Inches(0.14), y + Inches(0.33), CW - Inches(0.28), Inches(0.75),
             paras, size=8, line_spacing=1.0, space_after=3)

# eighth grid slot: mini takeaway
x = RX + 1 * (CW + GX)
y = Inches(1.94) + 3 * (CH + GY)
takeaway_box(s, x, y, CW, CH,
             "Same blueprint:",
             "proximity to operations, speed of iteration and platform "
             "leverage — applied to different industries.",
             body_size=8.5)

footer(s, 2, "Source: Company engineering blogs, job postings and public statements")

# ================================================================ SLIDE 3
s = prs.slides.add_slide(BLANK)
header(
    s,
    "Forward-deployed engineering fundamentally changes how technology "
    "organizations create value")

PL_X, PL_W = MARGIN, Inches(5.85)
PR_X, PR_W = Inches(6.93), Inches(5.85)
PANEL_Y = Inches(1.42)

# panel headers
h1 = add_rect(s, PL_X, PANEL_Y, PL_W, Inches(0.34), fill=CARD)
shape_text(h1, "TRADITIONAL OPERATING MODEL", size=10, color=GREY, bold=True)
h2 = add_rect(s, PR_X, PANEL_Y, PR_W, Inches(0.34), fill=NAVY)
shape_text(h2, "FORWARD-DEPLOYED MODEL", size=10, color=WHITE, bold=True)

# left: sequential flow
steps = ["Business", "Requirements", "Engineering", "Testing", "Deployment", "Business"]
BW, BH = Inches(2.5), Inches(0.31)
bx = PL_X + (PL_W - BW) / 2
by = Inches(1.98)
for i, step in enumerate(steps):
    box = add_rect(s, bx, by, BW, BH, fill=WHITE, line=MID_GREY, line_w=0.75)
    shape_text(box, step, size=9.5, color=GREY)
    if i < len(steps) - 1:
        ar = add_rect(s, PL_X + PL_W / 2 - Inches(0.05), by + BH + Inches(0.025),
                      Inches(0.1), Inches(0.1), fill=MID_GREY,
                      shape=MSO_SHAPE.DOWN_ARROW)
    by += BH + Inches(0.15)

wy = Inches(4.95)
add_text(s, PL_X, wy, PL_W, Inches(0.25), "STRUCTURAL WEAKNESSES",
         size=9, color=GREY, bold=True)
weaknesses = [
    "Intent degrades at every handoff",
    "Feedback arrives only after deployment",
    "Cycle times measured in quarters",
    "Business and IT optimize different goals",
]
wy += Inches(0.3)
for wtext in weaknesses:
    bullet(s, PL_X, wy, PL_W, wtext, size=9.5, color=GREY)
    wy += Inches(0.275)

# right: bidirectional flow
nodes = ["Business Team", "Embedded Engineer", "AI Builder",
         "Platform Team", "Reusable Enterprise Capability"]
BW2, BH2 = Inches(3.1), Inches(0.36)
bx = PR_X + (PR_W - BW2) / 2
by = Inches(1.9)
for i, node in enumerate(nodes):
    last = i == len(nodes) - 1
    box = add_rect(s, bx, by, BW2, BH2,
                   fill=LIGHT_BLUE if last else NAVY,
                   line=BLUE if last else None, line_w=1.0)
    shape_text(box, node, size=9.5, color=NAVY if last else WHITE,
               bold=last)
    if not last:
        ar = add_rect(s, PR_X + PR_W / 2 - Inches(0.055),
                      by + BH2 + Inches(0.03),
                      Inches(0.11), Inches(0.15), fill=BLUE,
                      shape=MSO_SHAPE.UP_DOWN_ARROW)
    by += BH2 + Inches(0.21)

wy = Inches(4.95)
add_text(s, PR_X, wy, PR_W, Inches(0.25), "STRUCTURAL BENEFITS",
         size=9, color=NAVY, bold=True)
benefits = [
    "Problems observed first-hand, not translated",
    "Working software in days; feedback continuous",
    "Every deployment compounds into a shared platform",
    "One team accountable for the business outcome",
]
wy += Inches(0.3)
for btext in benefits:
    bullet(s, PR_X, wy, PR_W, btext, size=9.5, color=INK)
    wy += Inches(0.275)

takeaway_box(
    s, MARGIN, Inches(6.38), CONTENT_W, Inches(0.58),
    "Bottom line:",
    "the traditional model manages requirements across handoffs; the "
    "forward-deployed model removes the handoffs — compressing learning "
    "cycles from quarters to days and converting each win into enterprise "
    "capability.")

footer(s, 3)

# ================================================================ SLIDE 4
s = prs.slides.add_slide(BLANK)
header(
    s,
    "A forward-deployed engineering capability requires changes across "
    "operating model, governance and talent")

cols = [
    ("Operating Model", [
        "Domain-aligned pods",
        "Dedicated embedded engineers",
        "Central platform enablement team",
        "Rotation between pod and platform",
        "Executive sponsor per domain",
    ]),
    ("Talent", [
        "Product-minded engineers",
        "Business-domain fluency",
        "Applied AI engineering depth",
        "Distinct FDE career track",
        "Incentives tied to business outcomes",
    ]),
    ("Technology", [
        "Shared enterprise AI platform",
        "Reusable components and agents",
        "Governed data access",
        "Observability and evaluation",
        "Fast sandbox-to-production path",
    ]),
    ("Governance", [
        "Clear domain ownership",
        "Portfolio-based funding model",
        "Security and data protection",
        "AI risk management",
        "Stage-gates for productization",
    ]),
    ("Success Metrics", [
        "Time from problem to deployment",
        "Business-user adoption",
        "Productivity and cost impact",
        "Platform reuse rate",
        "Customer and P&L outcomes",
    ]),
]

COL_W = Inches(2.29)
GAP = Inches(0.196)
TOP = Inches(1.5)
HEAD_H = Inches(0.4)
BODY_H = Inches(3.95)
for i, (title, items) in enumerate(cols):
    x = MARGIN + i * (COL_W + GAP)
    hd = add_rect(s, x, TOP, COL_W, HEAD_H, fill=NAVY)
    shape_text(hd, title, size=10, color=WHITE, bold=True)
    add_rect(s, x, TOP + HEAD_H + Inches(0.02), COL_W, BODY_H, fill=CARD)
    iy = TOP + HEAD_H + Inches(0.22)
    for item in items:
        add_rect(s, x + Inches(0.14), iy + Inches(0.05),
                 Inches(0.07), Inches(0.07), fill=BLUE)
        add_text(s, x + Inches(0.32), iy, COL_W - Inches(0.46), Inches(0.6),
                 item, size=9, color=INK, line_spacing=1.05)
        iy += Inches(0.72)

takeaway_box(
    s, MARGIN, Inches(6.12), CONTENT_W, Inches(0.82),
    "Recommendation:",
    "stand up two to three forward-deployed pods in high-value domains "
    "within 90 days, fund a shared platform team from day one, and scale "
    "only what demonstrates measured business impact.",
    body_size=10.5)

footer(s, 4)

OUT = "Forward-Deployed-Engineering_Executive-Deck.pptx"
prs.save(OUT)
print("saved", OUT)
