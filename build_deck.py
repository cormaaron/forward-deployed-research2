#!/usr/bin/env python3
"""Builds the forward-deployed engineering executive deck (.pptx).

Design rules:
- Black and white only (black text, greys for secondary text and fills).
- Title = narrative (18pt bold), black rule, subtitle = objective
  description of the slide (16pt bold).
- Body organized in two halves with sentence-case headers.
- Plain text bullets/numbered lists; shapes only for diagrams and tables.
- No all caps, normal spacing.
"""

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


def header(slide, title, subtitle):
    """Narrative title (18pt bold), black rule, objective subtitle
    (16pt bold). Returns the y where body content may start."""
    add_text(slide, MARGIN, Inches(0.32), CONTENT_W, Inches(0.75),
             title, size=18, color=BLACK, bold=True, line_spacing=1.05)
    add_line(slide, MARGIN, Inches(0.98), CONTENT_W, color=BLACK, weight=1.5)
    add_text(slide, MARGIN, Inches(1.1), CONTENT_W, Inches(0.35),
             subtitle, size=16, color=BLACK, bold=True, line_spacing=1.05)
    return Inches(1.62)


def footer(slide, page_no, source=None):
    add_line(slide, MARGIN, Inches(7.12), CONTENT_W, color=HAIRLINE, weight=0.5)
    if source:
        add_text(slide, MARGIN, Inches(7.2), Inches(10.5), Inches(0.25),
                 source, size=7.5, color=MID_GREY)
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

footer(s, 1, "Source: Company earnings calls, engineering blogs and public statements. Outcomes as publicly reported; not independently verified.")

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

footer(s, 2, "Source: Company engineering blogs, job postings and public statements")

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
         "One-pager per exemplar: model, motivation, differentiation and impact",
         size=12, color=GREY)
footer(s, 5)

# ================================================================ Backup one-pagers
QUESTION_HEADS = (
    "What the model looks like",
    "Why they moved to it",
    "What sets it apart",
    "Impact to date",
)

EXEMPLARS = [
    ("Palantir",
     "Palantir created the forward-deployed model and built its commercial "
     "engine around it",
     [
         ["FDEs deploy on-site within customer operations",
          "Two roles: domain-facing specialists and product-focused engineers",
          "Build directly on Foundry and AIP against live customer data",
          "Multi-week bootcamps deliver working software before contracts scale"],
         ["Government and industrial clients could not specify requirements upfront",
          "Complex, sensitive data environments demanded on-site iteration",
          "Conventional software sales failed to demonstrate value in these settings"],
         ["FDE is the core commercial motion, not a services add-on",
          "Field learnings flow directly into the product roadmap",
          "Engineers ship product code — not consulting deliverables"],
         ["Credits FDE-led AIP bootcamps for accelerating US commercial growth "
          "on earnings calls",
          "Sustained expansion within existing accounts publicly highlighted",
          "Model now widely emulated across the industry"],
     ],
     "Source: Palantir earnings calls, S-1 and public statements"),

    ("Uber",
     "Uber's embedded ops–engineering pairing powered city-by-city scaling "
     "and is now a product it sells",
     [
         ["Engineers embedded with city operations teams through the scaling era",
          "City teams owned the P&L; engineers tuned marketplace levers locally",
          "Model now offered externally to enterprises via Uber AI Solutions"],
         ["Every city market behaved differently; central roadmaps were too slow",
          "Launch speed was existential in winner-take-most markets",
          "Operational data lived in the field, not at headquarters"],
         ["Ops-led and engineer-supported — the business owned the outcome",
          "City learnings codified into playbooks and reused globally",
          "Matured from internal practice into an external commercial offering"],
         ["Embedded pairing publicly cited as central to rapid multi-market "
          "expansion",
          "Enabled local marketplace tuning at global scale",
          "Basis for the Uber AI Solutions business line"],
     ],
     "Source: Uber public statements, engineering blog and press coverage"),

    ("OpenAI",
     "OpenAI built a forward-deployed function to turn frontier models into "
     "deployed enterprise systems",
     [
         ["FDE pods embed within enterprise customer accounts",
          "Build production agent and workflow deployments on OpenAI models",
          "Scope spans integration, evaluation and adoption support"],
         ["Model capability outpaced enterprise ability to absorb it",
          "API access alone did not convert into production use cases",
          "Competition for enterprise adoption intensified"],
         ["Proximity to frontier research within customer engagements",
          "Field signal shapes model and product priorities",
          "Deployment patterns feed reusable product features"],
         ["Positions FDEs as core to enterprise adoption",
          "Enterprise agent deployments built with embedded teams publicly "
          "showcased"],
     ],
     "Source: OpenAI public statements, hiring posts and press coverage"),

    ("ElevenLabs",
     "ElevenLabs uses forward-deployed engineers to compress time-to-production "
     "for enterprise voice agents",
     [
         ["FDEs co-build production voice agents with customer engineering teams",
          "Engagements run end-to-end: design, integration, evaluation, launch"],
         ["Voice agents demand domain-specific tuning customers could not do alone",
          "Enterprise buyers required proof in production, not demonstrations"],
         ["Deep voice and audio specialization applied inside customer stacks",
          "Embedded talent lets a small company deliver enterprise-grade "
          "deployments"],
         ["Attributes enterprise expansion to its deployment-led model",
          "FDE hiring scaled publicly alongside enterprise growth"],
     ],
     "Source: ElevenLabs public statements, job postings and press coverage"),

    ("Stripe",
     "Stripe's user-proximate engineering culture is a long-standing form of "
     "forward deployment",
     [
         ["Engineers work directly with users — a founding practice",
          "Solution engineers embedded in key enterprise accounts",
          "Direct user contact expected of product engineers"],
         ["A developer-first product required first-hand user understanding",
          "Enterprise payments complexity could not be specified remotely"],
         ["Cultural rather than organizational — proximity is the default",
          "API design shaped by continuous developer feedback"],
         ["User-proximate engineering credited for product velocity and "
          "enterprise wins",
          "Developer experience widely treated as the industry benchmark"],
     ],
     "Source: Stripe public statements, founder interviews and press coverage"),

    ("Brex",
     "Brex points its forward-deployed engineers inward, rebuilding its own "
     "operations around AI",
     [
         ["AI engineers embedded within internal business teams",
          "Core workflows redesigned around AI-native tooling",
          "Adoption managed as an explicit leadership mandate"],
         ["Fintech margin pressure demanded step-change productivity",
          "Leadership prioritized internal AI adoption ahead of external "
          "products"],
         ["Internal-first: its own operations serve as the proving ground",
          "Embedded engineers accountable for adoption, not just delivery"],
         ["Leadership reports company-wide daily AI usage",
          "Workflow redesign publicly discussed as a productivity driver"],
     ],
     "Source: Brex leadership statements and press coverage"),

    ("Plaid",
     "Plaid deploys engineers into customer integration teams to make "
     "integrations faster and stickier",
     [
         ["FDEs work alongside customer engineers during onboarding and "
          "expansion",
          "Hands-on integration work against customer systems"],
         ["Integration timelines gated revenue for Plaid and its customers",
          "Bank and fintech data complexity exceeded documentation-led support"],
         ["Deep bank and fintech data expertise applied in customer codebases",
          "Embedded position surfaces product gaps early"],
         ["Publicly cites materially faster enterprise integrations",
          "Deeper, harder-to-displace customer relationships"],
     ],
     "Source: Plaid public statements, job postings and press coverage"),
]

BK_COL_W = Inches(2.87)
BK_GAP = Inches(0.25)
for page, (name, narrative, columns, source) in enumerate(EXEMPLARS, start=6):
    s = prs.slides.add_slide(BLANK)
    header(s, narrative,
           "Backup: " + name + " — model, motivation, differentiation and impact")
    for i, (head, items) in enumerate(zip(QUESTION_HEADS, columns)):
        x = MARGIN + i * (BK_COL_W + BK_GAP)
        add_text(s, x, Inches(1.92), BK_COL_W, Inches(0.28), head,
                 size=10.5, color=BLACK, bold=True)
        add_line(s, x, Inches(2.24), BK_COL_W, color=BLACK, weight=1.0)
        paras = [[("•  ", {}), (t, {})] for t in items]
        add_text(s, x, Inches(2.44), BK_COL_W, Inches(4.4), paras,
                 size=10.5, color=BLACK, line_spacing=1.2, space_after=16)
    footer(s, page, source)

OUT = "Forward-Deployed-Engineering_Executive-Deck.pptx"
prs.save(OUT)
print("saved", OUT)
