#!/usr/bin/env python3
"""Builds the forward-deployed engineering executive deck (.pptx).

Design rules:
- Black and white only (black text, greys for secondary text and fills).
- Title = narrative (18pt bold), black rule, subtitle = objective
  description of the slide (16pt bold).
- Body organized in two halves with sentence-case headers.
- Plain text bullets/numbered lists; shapes only for diagrams and tables.
- No all caps, normal spacing.

All statistics and quotes are drawn from public sources verified in
July 2026 (earnings calls, SEC filings, company blogs and annual letters,
and reputable press); source tags appear inline and per slide.
"""

import os

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

# ---------------------------------------------------------------- palette
BLACK = RGBColor(0x00, 0x00, 0x00)
GREY = RGBColor(0x59, 0x59, 0x59)
MID_GREY = RGBColor(0x8C, 0x8C, 0x8C)
HAIRLINE = RGBColor(0xBF, 0xBF, 0xBF)
LIGHT = RGBColor(0xF2, 0xF2, 0xF2)
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
def _style_run(run, size, color, bold=False):
    f = run.font
    f.name = FONT
    f.size = Pt(size)
    f.color.rgb = color
    f.bold = bold


def add_text(slide, x, y, w, h, runs, size=10, color=BLACK, bold=False,
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
                       ov.get("bold", bold))
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


def header(slide, title, subtitle, logo=None):
    """Narrative title (18pt bold), black rule, objective subtitle
    (16pt bold). Optional logo PNG at top right. Returns the y where
    body content may start."""
    title_w = CONTENT_W - Inches(0.85) if logo else CONTENT_W
    add_text(slide, MARGIN, Inches(0.32), title_w, Inches(0.75),
             title, size=18, color=BLACK, bold=True, line_spacing=1.05)
    if logo:
        size = Inches(0.5)
        slide.shapes.add_picture(logo, PAGE_W - MARGIN - size, Inches(0.3),
                                 height=size, width=size)
    add_line(slide, MARGIN, Inches(0.98), CONTENT_W, color=BLACK, weight=1.5)
    add_text(slide, MARGIN, Inches(1.1), CONTENT_W, Inches(0.35),
             subtitle, size=16, color=BLACK, bold=True, line_spacing=1.05)
    return Inches(1.62)


def footer(slide, page_no, source=None):
    add_line(slide, MARGIN, Inches(7.12), CONTENT_W, color=HAIRLINE, weight=0.5)
    if source:
        add_text(slide, MARGIN, Inches(7.18), Inches(11.3), Inches(0.3),
                 source, size=7, color=MID_GREY, line_spacing=1.1)
    add_text(slide, Inches(12.28), Inches(7.2), Inches(0.5), Inches(0.25),
             str(page_no), size=8, color=MID_GREY, align=PP_ALIGN.RIGHT)


def section_head(slide, x, y, w, text):
    add_text(slide, x, y, w, Inches(0.28), text,
             size=11, color=BLACK, bold=True)


def bullets(slide, x, y, w, items, size=10, color=BLACK, space_after=6,
            line_spacing=1.1, marker="•  "):
    """Plain text bullet list in a single text box."""
    paras = [[(marker, {"color": color}), (t, {})] for t in items]
    return add_text(slide, x, y, w, Inches(0.3 * len(items) + 0.3), paras,
                    size=size, color=color, space_after=space_after,
                    line_spacing=line_spacing)


def takeaway(slide, x, y, w, h, lead, body, body_size=10):
    box = add_rect(slide, x, y, w, h, fill=LIGHT)
    add_rect(slide, x, y, Inches(0.04), h, fill=BLACK)
    add_text(slide, x + Inches(0.2), y, w - Inches(0.4), h,
             [[(lead + "  ", {"bold": True}), (body, {})]],
             size=body_size, color=BLACK, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.12)
    return box


def quote_strip(slide, x, y, w, h, quote, attribution):
    box = add_rect(slide, x, y, w, h, fill=LIGHT)
    add_rect(slide, x, y, Inches(0.04), h, fill=BLACK)
    add_text(slide, x + Inches(0.22), y, w - Inches(0.44), h,
             [[("“" + quote + "”", {})],
              [("— " + attribution, {"size": 8.5, "color": GREY,
                                     "bold": True})]],
             size=10, color=BLACK, anchor=MSO_ANCHOR.MIDDLE,
             line_spacing=1.15, space_after=4)
    return box


def set_cell(cell, runs, size=8.5, color=BLACK, bold=False,
             align=PP_ALIGN.LEFT, fill=WHITE):
    cell.fill.solid()
    cell.fill.fore_color.rgb = fill
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


# ================================================================ Slide 1
s = prs.slides.add_slide(BLANK)
header(
    s,
    "Forward-deployed engineering is emerging as the leading operating model "
    "for scaling enterprise AI",
    "Model differentiators and publicly reported evidence from seven "
    "leading organizations")

LX, LW = MARGIN, Inches(4.35)
RX, RW = Inches(5.3), Inches(7.48)

section_head(s, LX, Inches(1.84), LW, "What differentiates the model")

diff_points = [
    ("Embedded in the business",
     "Engineers sit inside business teams, not central IT"),
    ("Problem-led, not spec-led",
     "Work starts from operational problems, not requirements"),
    ("Rapid iteration",
     "Solutions iterated live with users in days, not quarters"),
    ("Productization",
     "Proven solutions become reusable platform capabilities"),
]
py = Inches(2.32)
for i, (lead, body) in enumerate(diff_points):
    add_text(s, LX, py, LW, Inches(0.7),
             [[("%d.  " % (i + 1), {"bold": True}),
               (lead, {"bold": True})],
              [("     " + body, {"color": GREY, "size": 9.5})]],
             size=10.5, color=BLACK, line_spacing=1.15, space_after=2)
    py += Inches(0.82)

section_head(s, RX, Inches(1.84), RW, "Evidence from leading organizations")

rows = [
    ("Palantir",
     "Originator of the FDE model; 'Delta' engineers and 'Echo' strategists "
     "deploy on-site on Foundry and AIP",
     "US commercial revenue growth accelerated from +64% to +133% YoY over "
     "five quarters; 134% net retention"),
    ("Uber",
     "City teams with local P&L scaled via a central launch playbook; "
     "'Agentic Pods' now embed AI engineers",
     "~400 markets launched; Uber AI Solutions unit now in 30 countries "
     "with 50+ enterprise customers"),
    ("OpenAI",
     "FDE function (since Jan 2025) embeds with enterprises to take "
     "frontier models to production",
     "Enterprise >40% of revenue; 9M business users; $4B deployment "
     "venture launched May 2026"),
    ("ElevenLabs",
     "FDEs co-build production voice agents inside customer infrastructure",
     "$0 to $500M ARR in ~3.5 years; enterprise agent deployments live "
     "in 4–8 weeks"),
    ("Stripe",
     "'Users first' engineering culture; solutions architects embedded "
     "in enterprise accounts",
     "$1.9T volume in 2025 (+34% YoY); 90% of Dow Jones constituents "
     "run on Stripe"),
    ("Brex",
     "'Forward Deployed Agent Builders' embed with internal teams to "
     "automate core workflows",
     "Onboarding auto-approval 0→40% in weeks; manual identity reviews "
     "down 70%; responses 90% faster"),
    ("Plaid",
     "Embedded customer engineering: solutions engineers and technical "
     "account managers own integrations",
     ">$500M ARR in 2025 (~40% YoY), profitable; 1,000+ enterprise "
     "customers; >1M new connections daily"),
]

tbl_y = Inches(2.28)
tbl_h = Inches(3.66)
gframe = s.shapes.add_table(len(rows) + 1, 3, RX, tbl_y, RW, tbl_h)
table = gframe.table
plain_table_style(table)
table.columns[0].width = Inches(1.08)
table.columns[1].width = Inches(3.1)
table.columns[2].width = Inches(3.3)
table.rows[0].height = Inches(0.3)
for i in range(1, len(rows) + 1):
    table.rows[i].height = Inches(0.48)

for c, label in enumerate(("Company", "Operating model", "Publicly reported outcomes")):
    set_cell(table.cell(0, c), label, size=9, color=WHITE, bold=True, fill=BLACK)
for r, (name, model, outcome) in enumerate(rows, start=1):
    fill = WHITE if r % 2 else LIGHT
    set_cell(table.cell(r, 0), name, size=8.5, bold=True, fill=fill)
    set_cell(table.cell(r, 1), model, size=8, fill=fill)
    set_cell(table.cell(r, 2), outcome, size=8, color=GREY, fill=fill)

takeaway(
    s, RX, Inches(6.12), RW, Inches(0.76),
    "Common pattern:",
    "every organization positions engineers at the point of value creation "
    "and relies on a platform team to convert local wins into reusable "
    "enterprise assets.", body_size=9.5)

footer(s, 1, "Source: company SEC filings, earnings calls, annual letters, blogs and "
             "press reporting, 2014–2026; company-by-company detail and citations in backup")

# ================================================================ Slide 2
s = prs.slides.add_slide(BLANK)
header(
    s,
    "Leading organizations consistently deploy the same operating model "
    "despite differences in industry",
    "Six shared design principles and how each organization applies them")

LX, LW = MARGIN, Inches(4.5)
RX = Inches(5.45)

section_head(s, LX, Inches(1.84), LW, "Common design principles")
principles = [
    "Embed engineers alongside business operators",
    "Observe workflows first-hand before building",
    "Deploy small, multidisciplinary pods",
    "Run rapid build–measure–learn cycles",
    "Productize successful solutions",
    "Continuously strengthen the enterprise platform",
]
py = Inches(2.32)
for i, ptext in enumerate(principles):
    add_text(s, LX, py, LW, Inches(0.45),
             [[("%d.  " % (i + 1), {"bold": True}), (ptext, {})]],
             size=10.5, color=BLACK, line_spacing=1.15)
    py += Inches(0.62)

section_head(s, RX, Inches(1.84), Inches(7.3), "How each organization applies it")

companies = [
    ("Palantir",
     "On-site 'Deltas' build; 'Echoes' own the domain",
     "AIP bootcamps as the primary sales motion",
     "Land and expand enterprise accounts"),
    ("Uber",
     "City GMs with P&L; central launch playbook",
     "Playbook refined across ~400 market launches",
     "Market entry speed; now sold externally"),
    ("OpenAI",
     "FDE pods in enterprise accounts since 2025",
     "Frontier research proximity; $4B deployment arm",
     "Production agent deployments"),
    ("ElevenLabs",
     "FDEs co-build in customer infrastructure",
     "Voice specialization; live in 4–8 weeks",
     "Fastest path to production voice agents"),
    ("Stripe",
     "Engineers work backwards from users",
     "Customers inside the exec cadence; friction logs",
     "Product velocity from user contact"),
    ("Brex",
     "Forward-deployed agent builders in internal teams",
     "Adoption tied to reviews, hiring and bonuses",
     "AI-native operations"),
    ("Plaid",
     "Solutions engineers embedded in integrations",
     "Product pods co-develop with customers",
     "Faster, stickier network growth"),
]

CW = Inches(3.63)
CH = Inches(1.14)
GX = Inches(0.17)
labels = ("Model", "Edge", "Goal")
for i, (name, model, edge, goal) in enumerate(companies):
    col, row = i % 2, i // 2
    x = RX + col * (CW + GX)
    y = Inches(2.32) + row * CH
    add_line(s, x, y, CW, color=HAIRLINE, weight=0.5)
    add_text(s, x, y + Inches(0.08), CW, Inches(0.22),
             name, size=9.5, color=BLACK, bold=True)
    paras = []
    for label, val in zip(labels, (model, edge, goal)):
        paras.append([(label + "   ", {"bold": True, "color": MID_GREY, "size": 7}),
                      (val, {})])
    add_text(s, x, y + Inches(0.32), CW, Inches(0.75),
             paras, size=8, color=BLACK, line_spacing=1.0, space_after=3)

# eighth grid slot: takeaway
x = RX + 1 * (CW + GX)
y = Inches(2.32) + 3 * CH
takeaway(s, x, y + Inches(0.06), CW, CH - Inches(0.12),
         "Same blueprint:",
         "proximity to operations, speed of iteration and platform "
         "leverage — applied to different industries.", body_size=8.5)

footer(s, 2, "Source: company engineering blogs, careers pages and public statements; "
             "detail and citations in backup")

# ================================================================ Slide 3
s = prs.slides.add_slide(BLANK)
header(
    s,
    "Forward-deployed engineering fundamentally changes how technology "
    "organizations create value",
    "Comparison of the traditional and forward-deployed delivery models")

PL_X, PL_W = MARGIN, Inches(5.85)
PR_X, PR_W = Inches(6.93), Inches(5.85)

section_head(s, PL_X, Inches(1.84), PL_W, "Traditional operating model")
section_head(s, PR_X, Inches(1.84), PR_W, "Forward-deployed model")

# left: sequential flow
steps = ["Business", "Requirements", "Engineering", "Testing", "Deployment", "Business"]
BW, BH = Inches(2.5), Inches(0.3)
bx = PL_X + (PL_W - BW) / 2
by = Inches(2.26)
for i, step in enumerate(steps):
    box = add_rect(s, bx, by, BW, BH, fill=WHITE, line=BLACK, line_w=0.75)
    shape_text(box, step, size=9.5, color=BLACK)
    if i < len(steps) - 1:
        add_rect(s, PL_X + PL_W / 2 - Inches(0.05), by + BH + Inches(0.02),
                 Inches(0.1), Inches(0.1), fill=GREY,
                 shape=MSO_SHAPE.DOWN_ARROW)
    by += BH + Inches(0.14)

wy = Inches(5.0)
add_text(s, PL_X, wy, PL_W, Inches(0.25), "Structural weaknesses",
         size=10, color=BLACK, bold=True)
bullets(s, PL_X, wy + Inches(0.3), PL_W, [
    "Intent degrades at every handoff",
    "Feedback arrives only after deployment",
    "Cycle times measured in quarters",
    "Business and IT optimize different goals",
], size=9.5, color=GREY, space_after=4)

# right: bidirectional flow
nodes = ["Business team", "Embedded engineer", "AI builder",
         "Platform team", "Reusable enterprise capability"]
BW2, BH2 = Inches(3.1), Inches(0.35)
bx = PR_X + (PR_W - BW2) / 2
by = Inches(2.2)
for i, node in enumerate(nodes):
    last = i == len(nodes) - 1
    box = add_rect(s, bx, by, BW2, BH2,
                   fill=WHITE if last else BLACK,
                   line=BLACK if last else None, line_w=1.0)
    shape_text(box, node, size=9.5, color=BLACK if last else WHITE,
               bold=last)
    if not last:
        add_rect(s, PR_X + PR_W / 2 - Inches(0.05),
                 by + BH2 + Inches(0.03),
                 Inches(0.1), Inches(0.14), fill=GREY,
                 shape=MSO_SHAPE.UP_DOWN_ARROW)
    by += BH2 + Inches(0.2)

wy = Inches(5.0)
add_text(s, PR_X, wy, PR_W, Inches(0.25), "Structural benefits",
         size=10, color=BLACK, bold=True)
bullets(s, PR_X, wy + Inches(0.3), PR_W, [
    "Problems observed first-hand, not translated",
    "Working software in days; feedback continuous",
    "Every deployment compounds into a shared platform",
    "One team accountable for the business outcome",
], size=9.5, color=BLACK, space_after=4)

takeaway(
    s, MARGIN, Inches(6.4), CONTENT_W, Inches(0.56),
    "Bottom line:",
    "the traditional model manages requirements across handoffs; the "
    "forward-deployed model removes the handoffs — compressing learning "
    "cycles from quarters to days and converting each win into enterprise "
    "capability.", body_size=9.5)

footer(s, 3)

# ================================================================ Slide 4
s = prs.slides.add_slide(BLANK)
header(
    s,
    "A forward-deployed engineering capability requires changes across "
    "operating model, governance and talent",
    "Design requirements across five dimensions of the operating model")

cols = [
    ("Operating model", [
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
    ("Success metrics", [
        "Time from problem to deployment",
        "Business-user adoption",
        "Productivity and cost impact",
        "Platform reuse rate",
        "Customer and P&L outcomes",
    ]),
]

COL_W = Inches(2.29)
GAP = Inches(0.196)
TOP = Inches(1.92)
for i, (title, items) in enumerate(cols):
    x = MARGIN + i * (COL_W + GAP)
    add_text(s, x, TOP, COL_W, Inches(0.28), title,
             size=10.5, color=BLACK, bold=True)
    add_line(s, x, TOP + Inches(0.32), COL_W, color=BLACK, weight=1.0)
    iy = TOP + Inches(0.48)
    for item in items:
        add_text(s, x, iy, COL_W, Inches(0.55),
                 [[("•  ", {}), (item, {})]],
                 size=9, color=BLACK, line_spacing=1.1)
        iy += Inches(0.52)

takeaway(
    s, MARGIN, Inches(6.12), CONTENT_W, Inches(0.82),
    "Recommendation:",
    "stand up two to three forward-deployed pods in high-value domains "
    "within 90 days, fund a shared platform team from day one, and scale "
    "only what demonstrates measured business impact.", body_size=10.5)

footer(s, 4)

# ================================================================ Backup divider
s = prs.slides.add_slide(BLANK)
add_text(s, MARGIN, Inches(3.15), CONTENT_W, Inches(0.5),
         "Backup", size=24, color=BLACK, bold=True)
add_line(s, MARGIN, Inches(3.72), Inches(2.2), color=BLACK, weight=1.5)
add_text(s, MARGIN, Inches(3.88), CONTENT_W, Inches(0.35),
         "One-pager per exemplar: model, motivation, differentiation and "
         "impact — with published data, quotes and citations",
         size=12, color=GREY)
footer(s, 5)

# ================================================================ Backup one-pagers
QUESTION_HEADS = (
    "What the model looks like",
    "Why they moved to it",
    "What sets it apart",
    "Impact to date",
)

# Each bullet: (text, source_tag or None). Quotes verbatim from cited venues.
EXEMPLARS = [
    ("Palantir",
     "Palantir created the forward-deployed model and built its commercial "
     "engine around it",
     [
         [("'Delta' engineers write production code on site; 'Echo' "
           "strategists own the domain problem", "Palantir blog"),
          ("Product teams build the platforms; Deltas deploy them at the "
           "customer — field work feeds the roadmap", "Palantir blog"),
          ("AIP bootcamps: 1–5-day builds on the customer's own data, "
           "launched Sept 2023", "Palantir blog")],
         [("Government and industrial clients could not specify "
           "requirements upfront; FDE role created ~2006 by Shyam Sankar, "
           "employee #13, now CTO", None),
          ("Months-long pilots failed to prove value; bootcamps compress "
           "“what used to take three months” into days",
           "Q1 2024 earnings call")],
         [("Bootcamps are the primary go-to-market, not a services "
           "add-on: 560+ run across 465 organizations in the first four "
           "months", "Q4 2023 earnings call"),
          ("Described by Bloomberg as Palantir's “AI sales secret "
           "weapon”; cited 21 times on one earnings call",
           "Bloomberg, Apr 2024"),
          ("Model now copied by OpenAI, Anthropic and defense tech",
           "Forbes, Jul 2026")],
         [("US commercial revenue growth accelerated from +64% to +133% "
           "YoY across five quarters", "SEC filings, Q4 2024–Q1 2026"),
          ("Seven-figure deals signed 5 and 16 days after bootcamps",
           "2024 earnings calls"),
          ("Net dollar retention 134%; $2.8B total contract value closed "
           "in a single quarter", "Q3 2025 earnings")],
     ],
     ("In October, we set a goal of executing 500 AIP bootcamps within one "
      "year. We have already blown that goal out of the water, having "
      "completed more than 560 bootcamps across 465 organizations to-date.",
      "Ryan Taylor, Chief Revenue Officer, Q4 2023 earnings call, Feb 2024"),
     "Sources: Palantir SEC 8-K filings and earnings-call transcripts, Q4 2023–Q1 2026; "
     "Palantir blog ('Dev versus Delta'; AIP bootcamps); Bloomberg (Apr 23, 2024); Forbes (Jul 10, 2026)"),

    ("Uber",
     "Uber scaled city by city on embedded operating teams — and now sells "
     "the deployment model it built",
     [
         [("City teams ran with local P&L: a general manager — “the "
           "CEO of the city” — plus operations and marketing leads",
           "Fortune, 2015"),
          ("A central launcher team parachuted in with a continuously "
           "updated playbook, then handed off", "Bloomberg, 2014"),
          ("2026: 'Agentic Pods' embed Uber's most AI-proficient engineers "
           "in finance, legal and HR", "Pragmatic Engineer, 2026")],
         [("Every market behaved differently; central roadmaps were too "
           "slow for winner-take-most competition", None),
          ("Launch speed was existential — a new city opened every other "
           "day at the late-2014 peak", "Bloomberg, Nov 2014")],
         [("Ops-led with local P&L ownership; playbook codified learnings "
           "across ~400 market launches", "Fortune, 2015"),
          ("The capability became a product: Uber AI Solutions sells data, "
           "evaluation and deployment services externally",
           "Uber, Jun 2025")],
         [("66 to 266 cities in 2014 alone; 10,000+ cities today",
           "Forbes, 2014; Uber Newsroom"),
          ("Uber AI Solutions reached 30 countries and 50+ corporate "
           "customers by mid-2025", "Uber IR; Forbes, Jun 2025"),
          ("16 internal AI pods completed in the program's first two "
           "months", "2026")],
     ],
     ("We're bringing together Uber's platform, people, and AI systems to "
      "help other organizations build smarter AI more quickly.",
      "Megha Yethadka, GM, Uber AI Solutions, press release, Jun 2025"),
     "Sources: Uber IR press release (Jun 20, 2025); Forbes (Dec 2014; Jun 2025); Bloomberg "
     "Businessweek (Nov 2014); Fortune (Sept 2015); Uber Newsroom; Pragmatic Engineer (2026)"),

    ("OpenAI",
     "OpenAI built a forward-deployed function to turn frontier models into "
     "deployed enterprise systems",
     [
         [("FDE function created Jan 2025 under Colin Jarvis; pods own "
           "discovery through production, with ~50% travel",
           "OpenAI careers"),
          ("Engagements priced from $10M; clients include Morgan Stanley, "
           "T-Mobile and the US Department of Defense",
           "The Information, Jul 2025"),
          ("May 2026: model spun into the OpenAI Deployment Company with "
           "$4B committed and ~150 FDEs via the Tomoro acquisition",
           "OpenAI, May 2026")],
         [("Model capability outpaced enterprises' ability to absorb it; "
           "API access alone did not convert to production", None),
          ("Playbook imported deliberately — many early FDE hires came "
           "from Palantir", "The Information, Jul 2025")],
         [("Frontier research proximity inside customer accounts", None),
          ("Scaled beyond internal hiring: a majority-owned deployment "
           "venture plus alliances with McKinsey, BCG, Accenture and "
           "Capgemini", "OpenAI, May 2026")],
         [("Enterprise is now >40% of revenue, tracking to parity with "
           "consumer by end-2026", "OpenAI, 2026"),
          ("9M paying business users in Feb 2026, up from 3M in Jun 2025",
           "OpenAI; CNBC"),
          ("Go-to-market organization grew from ~50 to 700+ people in 18 "
           "months", "CNBC, Aug 2025")],
     ],
     ("I see our responsibility as both building the tools and being, in "
      "some ways, the most knowledgeable people in the world on how to "
      "deploy them.",
      "Brad Lightcap, COO, OpenAI, CNBC, Aug 2025"),
     "Sources: OpenAI announcements and careers pages (2025–26); CNBC (Jun 2025; Aug 2025; "
     "May 2026); The Information via The Decoder and Business Standard (Jul 2025)"),

    ("ElevenLabs",
     "ElevenLabs uses forward-deployed engineers to compress "
     "time-to-production for enterprise voice agents",
     [
         [("FDEs work “shoulder-to-shoulder with customers, from "
           "pre-sales evaluation through post-sales implementation”, "
           "building inside customer infrastructure",
           "ElevenLabs careers"),
          ("Engagements go from scoping to live deployment in 4–8 weeks",
           "ElevenLabs"),
          ("Multi-track FDE organization — engineers, strategists and "
           "graduates — hiring across three continents",
           "ElevenLabs careers")],
         [("CEO Mati Staniszewski is an ex-Palantir deployment "
           "strategist; the model was imported deliberately",
           "Sifted, Mar 2025"),
          ("Enterprise buyers demanded production voice agents, not "
           "demos; four ex-Palantir FDEs hired through 2024",
           "Sifted, Mar 2025")],
         [("Deep voice specialization with compliance built in: SOC 2, "
           "GDPR, HIPAA, zero-retention options", "ElevenLabs"),
          ("FDEs “have helped hundreds of enterprises launch AI "
           "agents”", "ElevenLabs")],
         [("ARR: $100M in 20 months, $200M in 10, $330M in 5; $500M by "
           "May 2026", "TechCrunch, Jan 2026; ElevenLabs"),
          ("$500M Series D at an $11B valuation, Feb 2026 (Sequoia-led)",
           "TechCrunch, Feb 2026"),
          ("Revolut: agents serve 4M+ customers in 30+ languages; ticket "
           "resolution cut by more than 8x", "ElevenLabs, Jan 2026")],
     ],
     ("It took us 20 months to reach $100 million in ARR, 10 months to "
      "reach $200 million, and five months to reach the current number.",
      "Mati Staniszewski, co-founder and CEO, Bloomberg interview, Jan 2026"),
     "Sources: ElevenLabs blog and careers pages (2025–26); TechCrunch (Jan 30, 2025; Jan 13 "
     "and Feb 4, 2026); CNBC (Feb and May 2026); Sifted (Mar 2025)"),

    ("Stripe",
     "Stripe institutionalized user-proximate engineering a decade before "
     "it had a name",
     [
         [("“Users first” is a codified operating principle: "
           "work backwards from user needs", "stripe.com"),
          ("A customer joins the first 30 minutes of the executive team "
           "meeting every other week, before ~40 leaders",
           "P. Collison, Apr 2025"),
          ("Friction logging: engineers and leaders walk product flows as "
           "users and log every snag", "Stripe DevRel"),
          ("Solutions architects and professional services embedded in "
           "enterprise accounts", "stripe.com")],
         [("A developer-first product required first-hand user "
           "understanding", None),
          ("Stripe ran for years without product managers — engineers own "
           "scoping and user contact", "Pragmatic Engineer, 2023")],
         [("Cultural rather than organizational: proximity is expected of "
           "every engineer and screened in interviews", None),
          ("Onboarding puts every hire in front of users: build an "
           "integration, answer real support tickets", None)],
         [("$1.9T processed in 2025, +34% YoY — roughly 1.6% of global "
           "GDP", "2025 annual letter"),
          ("90% of Dow Jones constituents and 80% of the Nasdaq 100 run "
           "on Stripe", "2025 annual letter"),
          ("Businesses on Stripe grew ~7x faster than S&P 500 revenue in "
           "aggregate", "2024 annual letter")],
     ],
     ("Every other week, we have a customer join for the first 30 minutes "
      "of our management team meeting: they share their candid feedback, "
      "and ~40 leaders from across Stripe listen.",
      "Patrick Collison, co-founder and CEO, Apr 2025"),
     "Sources: Stripe annual letters (Feb 2025; Feb 2026); Stripe operating principles "
     "(stripe.com/jobs/culture); TechCrunch (Apr 2025); The Pragmatic Engineer (Dec 2023)"),

    ("Brex",
     "Brex points its forward-deployed engineers inward, rebuilding its own "
     "operations around AI",
     [
         [("A formal role: 'Forward Deployed Agent Builder' engineers "
           "“embed with partner teams”, shadow the work, then "
           "ship agents that take over workflows", "Brex careers"),
          ("Three-pillar strategy: corporate AI, operational AI (owned by "
           "the COO) and product AI", "Latent.Space, Feb 2026"),
          ("~1,000 approved AI tools; each engineer holds a $50/month "
           "self-serve budget", "TechCrunch, Jul 2025")],
         [("The founding question: “If we were starting Brex today, "
           "how would we build?”", "Brex Journal, Aug 2025"),
          ("Fintech margin pressure demanded step-change productivity; "
           "BPO strategy shifted to agents", None)],
         [("Adoption is engineered, not assumed: four AI-fluency levels "
           "feed performance reviews; coding interviews require AI",
           "Semafor, Sept 2025"),
          ("225+ spot bonuses paid for employee AI projects",
           "Dec 2025")],
         [("Onboarding rebuilt around agents: auto-approval from 0% to "
           "40% in weeks; manual identity reviews down 70%",
           "Brex Journal, Jan 2026"),
          ("Customer responses 90% faster; ~15,000 customer hours saved "
           "per year", "First Round, 2025"),
          ("Onboarding time cut from days to minutes",
           "Brex Journal, Jan 2026")],
     ],
     ("I knew we needed to be AI-native on all our manual tasks, which "
      "meant shifting our BPO strategy, which meant agents doing the work, "
      "which meant managers overseeing agents and people.",
      "Camilla Matias, COO, Brex, First Round Review, 2025"),
     "Sources: Brex Journal (Aug 2025; Jan 2026); Brex careers ('Forward Deployed Agent "
     "Builder'); TechCrunch (Jul 2025); First Round Review (2025); Semafor (Sept 2025); Latent.Space (Feb 2026)"),

    ("Plaid",
     "Plaid runs forward deployment in all but name: embedded engineers "
     "drive integration and expansion",
     [
         [("A 'Customer Engineering & Solutions' organization: solutions "
           "engineers design and implement customer integrations",
           "Plaid careers"),
          ("Technical account managers own post-sales technical strategy "
           "and track integration health and adoption", "Plaid careers"),
          ("Product engineers work in pods that iterate with selected "
           "customers before building anything",
           "Plaid blog, 2022")],
         [("Bank-data integrations gate customer revenue; complexity "
           "exceeded documentation-led support", None),
          ("Expansion is driven by embedded technical staff, not by "
           "sales alone", "Plaid careers")],
         [("Forward-deployed outcomes without the title — the embedded "
           "roles sit in solutions engineering", None),
          ("Developer experience as strategy: “a few lines of "
           "code” to reach thousands of institutions", "Plaid docs")],
         [(">$500M ARR in 2025, ~40% YoY, profitable",
           "2025 shareholder letter"),
          ("1,000+ enterprise customers of ~8,000 total by mid-2024",
           "TechCrunch, Jun 2024"),
          (">1M new data connections daily; 12,000+ institutions; 1 in 2 "
           "US adults have connected via Plaid", "Plaid, 2026")],
     ],
     ("We're trying to identify the lowest level engineer, [and] build a "
      "fantastic experience for them… Someone's able to quickly build "
      "something in our sandbox.",
      "Zach Perret, co-founder and CEO, 2020 interview (via Contrary Research)"),
     "Sources: Plaid 2025 shareholder letter (Jan 2026); Plaid careers and blog; TechCrunch "
     "(Jun 2024; Apr 2025); Bloomberg (Jan 2025); Contrary Research"),
]

BK_COL_W = Inches(2.87)
BK_GAP = Inches(0.25)
for page, (name, narrative, columns, quote, sources) in enumerate(EXEMPLARS, start=6):
    s = prs.slides.add_slide(BLANK)
    logo = os.path.join("assets", "logos", name.lower() + ".png")
    header(s, narrative,
           "Backup: " + name + " — model, motivation, differentiation and impact",
           logo=logo if os.path.exists(logo) else None)
    for i, (head, items) in enumerate(zip(QUESTION_HEADS, columns)):
        x = MARGIN + i * (BK_COL_W + BK_GAP)
        add_text(s, x, Inches(1.92), BK_COL_W, Inches(0.28), head,
                 size=10.5, color=BLACK, bold=True)
        add_line(s, x, Inches(2.24), BK_COL_W, color=BLACK, weight=1.0)
        paras = []
        for text, src in items:
            runs = [("•  ", {}), (text, {})]
            if src:
                runs.append(("  (" + src + ")",
                             {"color": MID_GREY, "size": 7}))
            paras.append(runs)
        add_text(s, x, Inches(2.4), BK_COL_W, Inches(3.1), paras,
                 size=9, color=BLACK, line_spacing=1.15, space_after=9)
    quote_strip(s, MARGIN, Inches(5.72), CONTENT_W, Inches(1.08),
                quote[0], quote[1])
    footer(s, page, sources)

OUT = "Forward-Deployed-Engineering_Executive-Deck.pptx"
prs.save(OUT)
print("saved", OUT)
