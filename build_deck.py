#!/usr/bin/env python3
"""Generate the agent-native accounting investor deck."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# ---------------------------------------------------------------- design tokens

INK    = RGBColor(0x11, 0x1A, 0x27)
BODY   = RGBColor(0x33, 0x3E, 0x4F)
MUTED  = RGBColor(0x6B, 0x76, 0x88)
ACCENT = RGBColor(0x0E, 0x74, 0x90)   # teal — structure, substrate
WARM   = RGBColor(0xB4, 0x53, 0x09)   # rust — the defensible / the risk
PANEL  = RGBColor(0xF3, 0xF6, 0xF8)
PANEL2 = RGBColor(0xE8, 0xEE, 0xF2)
LINE   = RGBColor(0xD5, 0xDC, 0xE4)
WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
TINT_A = RGBColor(0xE4, 0xF0, 0xF4)   # accent tint
TINT_W = RGBColor(0xFA, 0xEE, 0xE3)   # warm tint

FONT = "Helvetica Neue"

SW, SH = 13.333, 7.5
ML, MR = 0.72, 0.72
CW = SW - ML - MR          # content width = 11.893

prs = Presentation()
prs.slide_width = Inches(SW)
prs.slide_height = Inches(SH)

_page = {"n": 0}


# ------------------------------------------------------------------- primitives

def new_slide(numbered=True):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    if numbered:
        _page["n"] += 1
        tb = s.shapes.add_textbox(Inches(SW - MR - 0.8), Inches(SH - 0.36),
                                  Inches(0.8), Inches(0.26))
        p = tb.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.RIGHT
        r = p.add_run()
        r.text = str(_page["n"])
        r.font.size, r.font.name, r.font.color.rgb = Pt(10), FONT, LINE
    return s


def text(slide, s, x, y, w, h, size=15, color=BODY, bold=False, align=PP_ALIGN.LEFT,
         space=6, line=1.25, caps=False, anchor=MSO_ANCHOR.TOP):
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    p.alignment = align
    p.space_after = Pt(space)
    p.line_spacing = line
    r = p.add_run()
    r.text = s.upper() if caps else s
    r.font.size, r.font.name, r.font.bold = Pt(size), FONT, bold
    r.font.color.rgb = color
    return tb


def eyebrow(slide, label):
    text(slide, label, ML, 0.46, CW, 0.28, size=10.5, color=ACCENT,
         bold=True, caps=True)


def heading(slide, s, size=27, y=0.82, w=None):
    return text(slide, s, ML, y, w or CW, 1.0, size=size, color=INK,
                bold=True, line=1.1)


def rule(slide, y, x=ML, w=None, color=LINE, thick=1.0):
    ln = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y),
                                Inches(w or CW), Pt(thick))
    ln.fill.solid(); ln.fill.fore_color.rgb = color
    ln.line.fill.background(); ln.shadow.inherit = False
    return ln


def box(slide, x, y, w, h, fill=PANEL, edge=None, radius=0.04):
    sh = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y),
                                Inches(w), Inches(h))
    sh.fill.solid(); sh.fill.fore_color.rgb = fill
    if edge:
        sh.line.color.rgb = edge; sh.line.width = Pt(1)
    else:
        sh.line.fill.background()
    sh.shadow.inherit = False
    try:
        sh.adjustments[0] = radius
    except Exception:
        pass
    sh.text_frame.text = ""
    return sh


def bullets(slide, items, x, y, w, h, size=15, gap=9, line=1.28):
    """items: list of (text, kind) where kind in {'h','b','sub','note'}"""
    tb = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    first = True
    for body, kind in items:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.line_spacing = line
        if kind == "h":
            p.space_before = Pt(gap + 5); p.space_after = Pt(2)
            r = p.add_run(); r.text = body
            r.font.size, r.font.bold, r.font.color.rgb = Pt(size), True, INK
        elif kind == "b":
            p.space_after = Pt(gap)
            r = p.add_run(); r.text = "— " + body
            r.font.size, r.font.color.rgb = Pt(size), BODY
        elif kind == "sub":
            p.space_after = Pt(gap - 3); p.level = 1
            r = p.add_run(); r.text = body
            r.font.size, r.font.color.rgb = Pt(size - 2), MUTED
        else:  # note
            p.space_before = Pt(gap); p.space_after = Pt(2)
            r = p.add_run(); r.text = body
            r.font.size, r.font.italic, r.font.color.rgb = Pt(size - 2), True, WARM
        r.font.name = FONT
    return tb


def table(slide, data, x, y, w, col_w, row_h=0.42, head_h=0.44, size=11.5,
          head_fill=INK, emphasize_col=None):
    rows, cols = len(data), len(data[0])
    gt = slide.shapes.add_table(rows, cols, Inches(x), Inches(y), Inches(w),
                                Inches(head_h + row_h * (rows - 1))).table
    gt.first_row = False
    gt.horz_banding = False
    for i, cwi in enumerate(col_w):
        gt.columns[i].width = Inches(cwi)
    gt.rows[0].height = Inches(head_h)
    for i in range(1, rows):
        gt.rows[i].height = Inches(row_h)

    for ri, row in enumerate(data):
        for ci, val in enumerate(row):
            cell = gt.cell(ri, ci)
            cell.margin_left = Inches(0.1)
            cell.margin_right = Inches(0.08)
            cell.margin_top = Inches(0.04)
            cell.margin_bottom = Inches(0.04)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.fill.solid()
            if ri == 0:
                cell.fill.fore_color.rgb = head_fill
            elif emphasize_col is not None and ci == emphasize_col:
                cell.fill.fore_color.rgb = TINT_W
            else:
                cell.fill.fore_color.rgb = WHITE if ri % 2 else PANEL
            tf = cell.text_frame
            tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if ci == 0 else PP_ALIGN.LEFT
            r = p.add_run()
            r.text = val
            r.font.size = Pt(size)
            r.font.name = FONT
            if ri == 0:
                r.font.bold = True; r.font.color.rgb = WHITE
            elif ci == 0:
                r.font.bold = True; r.font.color.rgb = INK
            elif emphasize_col is not None and ci == emphasize_col:
                r.font.color.rgb = RGBColor(0x8A, 0x3F, 0x07)
            else:
                r.font.color.rgb = BODY
    return gt


def notes(slide, body):
    slide.notes_slide.notes_text_frame.text = body.strip()


def stat(slide, x, y, w, big, label, color=ACCENT):
    text(slide, big, x, y, w, 0.62, size=34, color=color, bold=True, line=1.0)
    text(slide, label, x, y + 0.66, w, 0.9, size=12, color=MUTED, line=1.25)


# ============================================================== 1. TITLE

s = new_slide(numbered=False)
box(s, 0, 0, SW, SH, fill=INK, radius=0.0)
band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(SW), Pt(6))
band.fill.solid(); band.fill.fore_color.rgb = WARM
band.line.fill.background(); band.shadow.inherit = False

text(s, "Company strategy · Series [X] · [Month] 2026", ML, 1.5, CW, 0.3,
     size=11.5, color=RGBColor(0x7A, 0xB8, 0xC8), bold=True, caps=True)
text(s, "The Accounting Department,\nDelivered as Software",
     ML, 2.05, 10.6, 2.1, size=42, color=WHITE, bold=True, line=1.08)
text(s, "Traditional SaaS sold accountants a better tool. We sell the finished work — "
        "on a substrate built so a machine's output can be proven correct.",
     ML, 4.35, 9.6, 1.2, size=16.5, color=RGBColor(0xB9, 0xC4, 0xD2), line=1.4)
rule(s, 5.85, x=ML, w=2.2, color=WARM, thick=2.5)
text(s, "[Company]  ·  [presenter, title]  ·  [contact]", ML, 6.15, CW, 0.4,
     size=12, color=RGBColor(0x6E, 0x7C, 0x8E))

notes(s, """
Framing sentence before slide 2: "There are two ways to pitch AI in accounting. One is
'we added a copilot to bookkeeping software.' The other is 'we replaced the bookkeeper.'
Only the second one is a venture-scale business, and it requires a different architecture
from the ground up. That's what this deck is about."

Set expectations on structure: 5 minutes of why-this-is-structurally-different, then the
technical core, then roadmap and resourcing. Tell them the technical section is the real
content and to interrupt there.

[FILL BEFORE USE] round, date, presenter, contact.
""")


# ============================================================== 2. THESIS

s = new_slide()
eyebrow(s, "The thesis")
heading(s, "Software has always stopped at the edge of the work.\nThat edge just moved.")

box(s, ML, 2.35, 3.72, 3.5, fill=PANEL)
text(s, "1969", ML + 0.28, 2.62, 3.2, 0.4, size=13, color=MUTED, bold=True)
text(s, "Software unbundles\nfrom hardware", ML + 0.28, 3.02, 3.2, 0.9, size=17,
     color=INK, bold=True, line=1.2)
text(s, "IBM separates the program from the machine. A software industry becomes possible.",
     ML + 0.28, 4.25, 3.2, 1.3, size=12.5, color=BODY, line=1.35)

box(s, ML + 4.08, 2.35, 3.72, 3.5, fill=PANEL)
text(s, "1999", ML + 4.36, 2.62, 3.2, 0.4, size=13, color=MUTED, bold=True)
text(s, "Software unbundles\nfrom ownership", ML + 4.36, 3.02, 3.2, 0.9, size=17,
     color=INK, bold=True, line=1.2)
text(s, "Browser, broadband and multi-tenancy solve distribution. The seat becomes the meter.",
     ML + 4.36, 4.25, 3.2, 1.3, size=12.5, color=BODY, line=1.35)

box(s, ML + 8.16, 2.35, 3.72, 3.5, fill=TINT_W)
text(s, "NOW", ML + 8.44, 2.62, 3.2, 0.4, size=13, color=WARM, bold=True)
text(s, "The outcome unbundles\nfrom the operator", ML + 8.44, 3.02, 3.2, 0.9, size=17,
     color=INK, bold=True, line=1.2)
text(s, "The vendor can do the work, not just host the place where a human does it. "
        "The meter has to change with it.",
     ML + 8.44, 4.25, 3.2, 1.3, size=12.5, color=RGBColor(0x7A, 0x3B, 0x08), line=1.35)

text(s, "Accounting is the best first market for this, for one technical reason: it is a domain "
        "where correctness is machine-checkable.",
     ML, 6.25, CW, 0.6, size=15.5, color=INK, bold=True, line=1.3)

notes(s, """
This is the whole pitch in one slide. Deliver it, pause, then say you'll spend the rest
of the time defending the third box and the bottom line.

The historical framing does real work with investors: it positions this as the next step in
a pattern they already believe in, rather than as AI hype. Each prior unbundling created
larger companies than the one before, because each moved closer to the customer's actual
spend.

The bottom line is the slide's most important sentence and the one that separates us from
every other "AI for X" company. Most agent startups operate in domains where you cannot
tell whether the output is right — so they cap out at "assistive." Accounting has
double-entry, trial-balance ties, and bank reconciliation: a machine-checkable definition
of correct. That is what makes autonomy shippable here first.

Likely question: "Isn't this just what Intuit will ship?" Answer on slide 8 and 16 — hold it.
""")


# ============================================================== 3. WHY SAAS EXISTS

s = new_slide()
eyebrow(s, "Background · 1 of 3")
heading(s, "Why SaaS exists: software's problem was never production")

text(s, "Zero marginal cost to copy, high fixed cost to create. So the business problem of "
        "software is only ever distribution and metering. Every era is an answer to those two.",
     ML, 1.92, 11.0, 0.7, size=15, color=BODY, line=1.35)

rule(s, 2.82)

cols = [
    ("Distribution", "No install, no procurement, self-serve trial, land-and-expand. "
     "Recurring revenue made CAC financeable — which made the entire GTM machine possible."),
    ("Marginal cost", "Multi-tenancy: one codebase, one deployment, N customers. "
     "This single fact — not anything else — is why SaaS gross margins are 75–85%."),
    ("Metering", "The seat. Nobody ever wanted a login. The seat was chosen because it was "
     "cheap to count, hard to game, and correlated with usage. It was a proxy."),
    ("Alignment", "Renewal disciplines the vendor; continuous delivery replaces the "
     "big-bang upgrade. Risk for infra, uptime and upgrades transfers to the vendor."),
]
cx = ML
for i, (h, b) in enumerate(cols):
    w = 2.79
    text(s, h, cx, 3.05, w, 0.4, size=17, color=ACCENT, bold=True)
    text(s, b, cx, 3.58, w, 2.2, size=13, color=BODY, line=1.4)
    cx += w + 0.25

box(s, ML, 5.95, CW, 0.95, fill=TINT_W)
text(s, "Hold on to the third column. The seat was a measurement convenience that assumed a "
        "human doing work per seat. Everything in the next section follows from that assumption breaking.",
     ML + 0.3, 6.16, CW - 0.6, 0.6, size=14.5, color=RGBColor(0x7A, 0x3B, 0x08),
     bold=True, line=1.3)

notes(s, """
Purpose of this slide: establish that you understand SaaS as an economic machine, not just
as "the thing before AI." Investors with technical backgrounds have usually never heard the
seat described as a proxy metric, and it reframes the rest of the conversation.

Say out loud: "SaaS didn't win because it was in the cloud. It won because multi-tenancy
made marginal cost near zero and the seat made revenue predictable. Both of those are
about metering and distribution — not about the software being better."

Land the last box hard. It's the pivot for the whole deck: we are not claiming SaaS was
badly built. We're claiming its meter was calibrated to an assumption that no longer holds.

If someone asks why history matters: because the failure mode of this category is companies
that build agent features and keep seat pricing. That combination is structurally
unprofitable and we want to explain why we avoided it.
""")


# ============================================================== 4. WHAT SAAS DIDN'T SOLVE

s = new_slide()
eyebrow(s, "Background · 2 of 3")
heading(s, "What SaaS deliberately did not solve")

text(s, "SaaS sells the arena, not the win. The customer still supplies intent, judgment, "
        "attention and keystrokes.",
     ML, 1.9, 11.0, 0.4, size=15.5, color=INK, bold=True)

items = [
    ("Value delivered is capped by the customer's own labor capacity.", "h"),
    ("A CRM doesn't sell you pipeline; it sells a place to put pipeline. This is why "
     "implementation, adoption, change management and shelfware are permanent features "
     "of the industry rather than execution failures.", "b"),
    ("The interface became the moat — and the cost centre.", "h"),
    ("Because humans do the work, switching cost is muscle memory, admin config and "
     "integrations. Feature bloat isn't incompetence: it is customer heterogeneity "
     "projected onto deterministic code.", "b"),
    ("The human was the error-handling layer.", "h"),
    ("SaaS validation is syntactic — is the date a date, do debits equal credits. It is "
     "almost never semantic — does this expense belong in that account. Semantic "
     "correctness was silently outsourced to the user.", "b"),
]
bullets(s, items, ML, 2.5, 7.3, 4.0, size=14)

box(s, ML + 7.75, 2.5, 4.12, 3.62, fill=PANEL)
text(s, "Consequence", ML + 8.05, 2.75, 3.5, 0.3, size=11, color=MUTED, bold=True, caps=True)
text(s, "Software self-selected\ninto the small slice.", ML + 8.05, 3.1, 3.5, 0.8,
     size=18, color=INK, bold=True, line=1.15)
text(s, "For any business function, software captures single-digit percentages of that "
        "function's total cost.\n\nPayroll software vs. the payroll team. Legal tech vs. "
        "legal spend. Accounting software vs. accountants.\n\nNot an oversight — software "
        "could not do the other part.",
     ML + 8.05, 4.05, 3.5, 2.0, size=12.5, color=BODY, line=1.4)

text(s, "[SOURCE NEEDED — put the accounting-specific software-spend vs. labour-spend ratio here; "
        "it carries the entire TAM argument]",
     ML, 6.45, CW, 0.5, size=11.5, color=WARM, bold=True)

notes(s, """
This is the most important background slide. The "human was the error-handling layer" point
is the one technical investors respond to, because it reframes 25 years of software as
having a much cheaper correctness contract than people realise: prevent invalid states and
show a clear error. Not: produce the right answer.

Say it plainly: "Every form you have ever filled in is a device for converting fuzzy human
intent into a structure a program can execute. Forms are the scar tissue of software's
inability to handle ambiguity."

The right-hand panel sets up the TAM slide. Do NOT overclaim the number here — put the
sourced figure in and cite it. If you assert a ratio you can't defend, a sharp investor
will spend the rest of the meeting on it and you lose the technical section.

ACTION: replace the red placeholder line before this deck leaves the room.
""")


# ============================================================== 5. THREE CONSTRAINTS

s = new_slide()
eyebrow(s, "Background · 3 of 3")
heading(s, "Three constraints forced that design. All three just lifted.")

hdr = [["The constraint that made forms necessary", "What it forced", "What changes now"]]
rows = [
    ["No tolerance for ambiguity",
     "Every path pre-specified. Anything needing interpretation of unstructured input was "
     "handed to a human behind a form.",
     "Models absorb messy documents, bank memos, contracts, email threads — the exact "
     "inputs SaaS pushed onto people."],
    ["No economics for the long tail",
     "The 1,000th edge case costs as much as the first. Vendors served the mode and left "
     "the tail to consultants and spreadsheets.",
     "Marginal cost of covering an edge case approaches the cost of describing it. The tail "
     "becomes addressable."],
    ["Every capability needed a pre-built interface",
     "Customers could only direct software through affordances a PM imagined and a roadmap "
     "funded.",
     "Natural language is a universal interface. No vendor has to anticipate the request "
     "in advance."],
]
table(s, hdr + rows, ML, 2.05, CW, [3.3, 4.29, 4.3], row_h=1.20, head_h=0.46,
      size=12.5, emphasize_col=2)

box(s, ML, 6.28, CW, 0.82, fill=TINT_A)
text(s, "The claim under this deck is not \"AI is smart.\" It is three specific engineering "
        "properties: ambiguity tolerance, near-zero marginal cost on the long tail, and an "
        "interface nobody has to build in advance.",
     ML + 0.3, 6.48, CW - 0.6, 0.5, size=14, color=RGBColor(0x0A, 0x53, 0x66),
     bold=True, line=1.3)

notes(s, """
This slide earns credibility with a technical audience because it refuses the hype framing.
Any investor who has sat through forty AI pitches has heard "AI changes everything." Almost
none have heard a precise statement of which constraints lifted.

Walk the rows left to right, one sentence each. Then deliver the bottom box verbatim — it's
the most quotable line in the deck and it's what they'll repeat to their partnership.

Anticipated pushback: "the long tail claim is doing a lot of work." Concede partially — the
marginal cost of the tail falls a lot but not to zero, and reliability on tail cases is
worse than on the mode. That's precisely why the architecture has a verification layer and
a human review queue rather than assumed autonomy. Point forward to slide 12.
""")


# ============================================================== 6. THREE ERAS TABLE

s = new_slide()
eyebrow(s, "The shift")
heading(s, "Three eras, side by side")

hdr = [["", "Perpetual licence", "SaaS", "Agent-native"]]
rows = [
    ["Unit of sale", "a copy", "access", "completed work"],
    ["Meter", "licence + maintenance", "seat / month", "unit of output"],
    ["Who does the work", "customer", "customer", "vendor"],
    ["Marginal cost", "media, channel", "hosting — low, fixed", "inference + review — variable, falling"],
    ["Gross margin", "high", "75–85%, flat from day one", "40–60% rising toward 70–80%"],
    ["Value ceiling", "customer's labour capacity", "customer's labour capacity", "the work itself"],
    ["Moat", "install base, data gravity", "UI habit, integrations, config", "verification, context, write permission, evals"],
    ["Buyer / budget", "IT · capex", "function owner · software opex", "P&L owner · labour opex"],
    ["Benchmarked against", "other licences", "other software", "a salary"],
    ["Interface", "GUI", "web GUI", "conversation + exception queue"],
]
table(s, hdr + rows, ML, 1.95, CW, [2.3, 2.75, 3.2, 3.68], row_h=0.42, head_h=0.44,
      size=11.5, emphasize_col=3)

notea = ("Read the last four rows first — they are the investment case: the value ceiling "
         "lifts, the price benchmark moves from a software budget to a salary, and the moat "
         "relocates to things that compound.")
text(s, notea, ML, 6.62, CW, 0.5, size=12, color=WARM, bold=True, line=1.3)

notes(s, """
Hand this one over rather than reading it. Say: "read the last four rows first." Then be
quiet for five seconds — technical investors want to scan a table like this themselves.

The two rows to speak to:

GROSS MARGIN. This is the row that will generate the hardest question, so pre-empt it.
Classic SaaS had 80% margins on day one because marginal cost was hosting. We start at
services-like margins because inference and human review are real COGS, and we climb as
automation rate rises and token prices fall. Our financial story is not "we have SaaS
margins" — it is "we have a margin slope." Anyone who claims 80% gross margin in year one
in this category is either subsidising with venture money or not measuring properly.

BENCHMARKED AGAINST. Our competitor is not Xero. It's the decision to hire a bookkeeper.
That changes the price ceiling by an order of magnitude, and it changes the sale: it's
bought from a labour budget by a P&L owner, so the scrutiny is on reliability and liability
rather than on features.
""")


# ============================================================== 7. WHAT ACTUALLY SHIFTS

s = new_slide()
eyebrow(s, "The shift")
heading(s, "Five shifts, and one thing that does not change")

left = [
    ("The unit of sale becomes completed work.", "h"),
    ("SaaS ate services firms' margin by selling them better tools. Agents take services "
     "firms' revenue by doing the work.", "b"),
    ("The meter breaks and must be rebuilt.", "h"),
    ("Price on a unit the customer already counts in their own P&L: entities, transactions, "
     "closes. Not seats.", "b"),
    ("Gross margin becomes earned, not structural.", "h"),
    ("Automation rate is simultaneously product quality, the GM driver, and the moat. It is "
     "our single north-star metric.", "b"),
]
bullets(s, left, ML, 2.0, 5.6, 3.4, size=13.5)

right = [
    ("The moat relocates — to four things that compound.", "h"),
    ("Verification (anyone can prompt a model; almost nobody can prove the output). "
     "Context (client-specific policy and history). Permission (write access to the ledger, "
     "the bank, the filing). Evals (a corpus of verified outcomes).", "b"),
    ("Buyer, budget and benchmark all move.", "h"),
    ("From software budget vs. other software, to labour budget vs. a salary. Higher "
     "ceiling, longer trust cycle, different scrutiny.", "b"),
]
bullets(s, right, ML + 6.05, 2.0, 5.85, 2.6, size=13.5)

box(s, ML + 6.05, 4.72, 5.85, 2.05, fill=PANEL)
text(s, "What does not change", ML + 6.35, 4.94, 5.3, 0.3, size=11, color=MUTED,
     bold=True, caps=True)
text(s, "Systems of record matter more, not less — agents need a substrate with truth, "
        "permissions and history. Agents make interfaces obsolete, not databases. "
        "Trust, security and compliance still gate enterprise adoption. And distribution "
        "still decides who wins: better technology has never once been sufficient.",
     ML + 6.35, 5.28, 5.3, 1.4, size=12.5, color=BODY, line=1.4)

notes(s, """
Pace: 30 seconds per left-hand item, then hand the right side over.

The "what does not change" box is deliberate and it does more for credibility than anything
else on the slide. Every investor in this category has heard someone claim agents make
databases or SaaS obsolete. Saying plainly that the system of record becomes MORE important
signals that we've actually built something.

On the moat: draw out permission specifically. UI habit is a preference, and preferences
get overridden. Write access to a company's ledger and its bank is a governance decision
with a liability trail — reversing it requires a board-level conversation, not a
preference change. That is a materially stronger switching cost than anything SaaS had.

If asked "which of the four moats do you have today?" — be honest. Verification is built,
evals are in progress, context accrues per customer per month, permission comes with
each deployment. Don't claim all four are mature.
""")


# ============================================================== 8. WHY ACCOUNTING

s = new_slide()
eyebrow(s, "Market selection")
heading(s, "Why accounting first: the domain hands us an oracle")

text(s, "Most agent products fail because nobody can tell whether the output is right — so "
        "they cap out at \"assistive\". Accounting is one of very few domains with a "
        "machine-checkable definition of correct.",
     ML, 1.9, 11.2, 0.7, size=15, color=BODY, line=1.35)

box(s, ML, 2.72, 5.75, 3.05, fill=TINT_A)
text(s, "Deterministic checks we get for free", ML + 0.3, 2.95, 5.15, 0.3, size=11,
     color=RGBColor(0x0A, 0x53, 0x66), bold=True, caps=True)
chk = [
    ("Debits must equal credits on every entry", "b"),
    ("Trial balance must tie", "b"),
    ("Subledgers must tie to control accounts", "b"),
    ("Bank reconciliation must net to zero", "b"),
    ("Tax and filing rules are codified, not inferred", "b"),
]
bullets(s, chk, ML + 0.3, 3.35, 5.15, 2.2, size=13, gap=7)

box(s, ML + 6.13, 2.72, 5.75, 3.05, fill=PANEL)
text(s, "What that buys us architecturally", ML + 6.43, 2.95, 5.15, 0.3, size=11,
     color=MUTED, bold=True, caps=True)
text(s, "We can verify before we commit.\n\nThat single capability is the difference between "
        "a demo and a product. It lets us auto-post with a provable error bound, route only "
        "genuine ambiguity to a human, and put a number on our own accuracy that a "
        "customer's auditor will accept.\n\nIt is also why our margin slope is credible: the "
        "review queue shrinks against a measurable ceiling, not a hoped-for one.",
     ML + 6.43, 3.35, 5.15, 2.3, size=13, color=BODY, line=1.42)

text(s, "Secondary reasons: high-volume repetitive work, messy unstructured inputs where "
        "classic software loses, and a buyer (the firm) whose gross margin we roughly double.",
     ML, 6.1, CW, 0.6, size=13.5, color=INK, line=1.35)
text(s, "[SOURCE NEEDED — firm gross-margin benchmark before/after, from design-partner data]",
     ML, 6.72, CW, 0.4, size=11.5, color=WARM, bold=True)

notes(s, """
This is the slide that answers "why won't you get killed by a horizontal agent platform"
and "why this vertical." Both answers are the same: the oracle.

Make the general point first — in law, in consulting, in most knowledge work, you cannot
mechanically check the answer, so the product ceiling is assistive and the human never
leaves the loop. In accounting the invariants are the check. We didn't choose this market
because accounting is exciting. We chose it because it is verifiable.

Then the consequence, which is the commercially important half: verification lets us
auto-post with a provable error bound. That is what makes autonomy sellable to a CFO and
defensible to an auditor.

Have the design-partner margin data ready as a backup slide if asked. If we don't have it
yet, say so and give the theoretical basis — don't invent a number.
""")


# ============================================================== 9. ARCHITECTURE OVERVIEW

s = new_slide()
eyebrow(s, "Technical architecture")
heading(s, "The agent decides. Deterministic code executes.\nNothing commits unverified.", size=24)

LX, LW = ML, 9.85
XX, XW = ML + 10.05, 1.83
top = 2.42
bh, bg = 0.515, 0.075

layers = [
    ("Client surfaces", "exception queue (the real UI) · chat & email · reports", PANEL, INK, False),
    ("Verification & guardrails", "hard invariants · anomaly and variance · confidence routing", TINT_W, INK, True),
    ("Agent layer", "workflow supervisors · extract · categorise · match · chase · analyse", PANEL, INK, False),
    ("Durable orchestration", "the close checklist as a resumable state machine, not free-form agency", PANEL, INK, False),
    ("Deterministic tool layer", "typed, invariant-enforcing actions: post_entry · reconcile · close_period", TINT_W, INK, True),
    ("Ledger core", "append-only · bitemporal · double-entry · full provenance", TINT_W, INK, True),
    ("Ingestion", "banks · cards · payroll · POS · AP inbox · documents · prior-year books", PANEL, INK, False),
]
y = top
for name, detail, fill, col, star in layers:
    box(s, LX, y, LW, bh, fill=fill, edge=LINE if fill is PANEL else WARM)
    text(s, name, LX + 0.22, y + 0.085, 3.0, 0.36, size=13.5, color=col, bold=True)
    text(s, detail, LX + 3.3, y + 0.105, LW - 3.55, 0.36, size=11.5, color=BODY)
    if star:
        text(s, "◆", LX + LW - 0.34, y + 0.1, 0.3, 0.3, size=12, color=WARM, bold=True)
    y += bh + bg

xh = y - top - bg
box(s, XX, top, XW, xh, fill=PANEL2, edge=LINE)
text(s, "Cross-cutting", XX + 0.16, top + 0.14, XW - 0.32, 0.3, size=10, color=MUTED,
     bold=True, caps=True)
cc = [("Audit trail\n& lineage", 0.62), ("Evals &\nreplay harness", 1.35),
      ("Cost & token\ntelemetry", 2.08), ("Agent identity,\nleast privilege", 2.81)]
for label, dy in cc:
    text(s, label, XX + 0.16, top + dy, XW - 0.32, 0.6, size=11, color=BODY, line=1.25)

text(s, "◆  Proprietary and defensible. The agent layer is the part we will rewrite most "
        "often and the part we depend on least.",
     ML, y + 0.12, CW, 0.4, size=12.5, color=WARM, bold=True)

notes(s, """
Spend the most time here. Read the title as a sentence — it is the design philosophy in
nine words.

Walk it bottom-up, not top-down: ingestion, ledger, tools, orchestration, agents,
verification, UI. Bottom-up is the build order and it makes the dependency structure
obvious.

The point to hammer: the three diamonds are where the value is, and the agent layer is not
one of them. Prompts are the cheapest and most replaceable component in the system. Teams
in this category consistently over-invest in the agent layer because it is the visible,
fun part, then discover they have no verification and no evals and cannot tell whether
they are improving. We inverted that.

Second point: note that verification sits ABOVE the agent and BELOW the customer. Nothing
the agent produces reaches the books without passing deterministic checks first. That is
an architectural guarantee, not a policy.

Expect: "why durable orchestration rather than letting the agent plan?" Answer: a month-end
close is a known checklist, spanning days, waiting on a client's email, needing compensation
logic on failure. We model the known process deterministically and use agents only inside
the judgment-bearing steps. Free-roaming agency over a general ledger is how you get an
unexplainable restatement.
""")


# ============================================================== 10. LEDGER SUBSTRATE

s = new_slide()
eyebrow(s, "Key component · 1 of 4")
heading(s, "Ledger substrate: append-only, bitemporal, fully attributed")

text(s, "SaaS updates a row and moves on. Three requirements SaaS never had make that "
        "impossible for us — and all three are schema decisions, so all three are day-one "
        "decisions. Retrofitting them is close to impossible.",
     ML, 1.9, 11.3, 0.7, size=14.5, color=BODY, line=1.35)

items = [
    ("Bitemporality — booking date vs. effective date", "h"),
    ("We must answer \"what did the books say as of 31 March?\" for audit, and — critically — "
     "to replay an agent against history and score it. Our eval capability only exists "
     "because the ledger was built this way.", "b"),
    ("Immutability — corrections are reversing entries, never updates", "h"),
    ("A mutable ledger cannot be explained after the fact, and an agent that can silently "
     "overwrite is not auditable at any price.", "b"),
    ("Provenance as a first-class, queryable chain", "h"),
    ("source document → extraction → reasoning trace → tool call → policy applied → "
     "approver → posting. In SaaS an audit log was a compliance checkbox bolted on the side. "
     "For us it is load-bearing product: it is how the output gets trusted and sold.", "b"),
    ("Every actor is typed: human, agent, or system — with the agent's identity, model "
     "version and prompt revision recorded on the entry.", "note"),
]
bullets(s, items, ML, 2.7, 7.4, 4.0, size=13.5)

box(s, ML + 7.85, 2.7, 4.02, 3.6, fill=PANEL)
text(s, "Why an investor should care", ML + 8.13, 2.94, 3.45, 0.3, size=11, color=MUTED,
     bold=True, caps=True)
text(s, "This is the layer a competitor cannot copy by copying our prompts.\n\n"
        "It is also the layer that decides whether we can ever sell to a regulated or "
        "audited customer. Deals in this market die on \"show me how this entry was "
        "produced,\" and that question is answerable only if it was designed for at the "
        "schema level.\n\nIt is roughly [15–20]% of engineering effort and almost none of "
        "the demo.",
     ML + 8.13, 3.32, 3.45, 2.8, size=12.5, color=BODY, line=1.42)

notes(s, """
The line to land: "this is the layer a competitor cannot copy by copying our prompts."

Bitemporality usually needs an example. Use this one: a customer's March books were closed
in April. In June, an auditor asks what we believed on 31 March and why. A normal SaaS
database cannot answer that — it only holds current state. We can, because we store both
when a fact became true and when we learned it. The same mechanism is what lets us take
last year's real closed books and replay our agents against them to measure accuracy —
which is our entire eval strategy, and it is only possible because of this schema choice.

If a technical investor pushes on cost: yes, append-only bitemporal ledgers are more
expensive to build and query than mutable tables, and we accepted that deliberately in
month one because it is unrecoverable later.

Note the last line — this layer is 15-20% of effort and almost none of the demo. That
asymmetry is worth stating; it tells them where our engineering judgment is.
""")


# ============================================================== 11. TOOL LAYER

s = new_slide()
eyebrow(s, "Key component · 2 of 4")
heading(s, "Tool layer: an interface for an unknown reasoner, not a known client")

text(s, "This is not our REST API with a different name. A REST API assumes a client that "
        "already knows what it wants. A tool interface assumes a capable, confident, "
        "occasionally wrong caller. Five things differ:",
     ML, 1.9, 11.3, 0.7, size=14.5, color=BODY, line=1.35)

items = [
    ("Errors must be instructive, because the agent reads them and retries.", "h"),
    ("\"400 invalid\" is worthless. \"Entry unbalanced: debits 1,200.00, credits 1,150.00, "
     "difference 50.00\" produces a retry that succeeds. Error-message quality is a "
     "functional requirement here, not a nicety.", "b"),
    ("The tool owns the invariant — not the caller.", "h"),
    ("In SaaS you trusted your own frontend to send sane data. We invert that: every tool "
     "assumes a wrong caller and enforces balance, period locks and permissions itself.", "b"),
    ("Idempotency is mandatory, because agents retry.", "h"),
    ("Granularity is a real tradeoff.", "h"),
    ("Too fine and the agent burns forty calls and loses the thread; too coarse and it "
     "cannot express intent. Expect to iterate — we have.", "b"),
    ("The surface must stay small and discoverable. Context is a scarce resource; you "
     "cannot expose four hundred endpoints.", "h"),
]
bullets(s, items, ML, 2.7, 7.4, 4.1, size=13.5, gap=7)

box(s, ML + 7.85, 2.7, 4.02, 2.35, fill=TINT_A)
text(s, "The rule", ML + 8.13, 2.92, 3.45, 0.3, size=11,
     color=RGBColor(0x0A, 0x53, 0x66), bold=True, caps=True)
text(s, "The model never does arithmetic or applies a rule in its head.\n\nEvery state "
        "change is a typed tool call that validates itself and rejects bad input. The "
        "agent's job is choosing which tool with which arguments. The tool's job is being "
        "correct.",
     ML + 8.13, 3.3, 3.45, 1.6, size=12.5, color=BODY, line=1.42)

box(s, ML + 7.85, 5.2, 4.02, 1.1, fill=PANEL)
text(s, "This is the single highest-leverage design decision in the system.",
     ML + 8.13, 5.45, 3.45, 0.7, size=13, color=INK, bold=True, line=1.3)

notes(s, """
The error-message point is the one that convinces engineers we've actually shipped this.
Nobody who hasn't built an agent system thinks about error strings as a functional
requirement. Give the concrete example verbatim — the unbalanced-entry one.

The inverted trust relationship is the second thing to emphasise. Every SaaS engineer has
written validation in the frontend and trusted it. You cannot do that when the caller is a
model, so the invariant has to live in the tool. It's a small idea with large consequences
for how the codebase is organised.

The right-hand rule box is the thing to repeat if you only get one point across: the model
never does arithmetic. Every wrong number a competitor's demo produces comes from letting
the model compute instead of calling a tool that computes.

If asked about MCP or protocol standardisation: we're compatible with it but it's an
integration detail, not the hard part. The hard part is designing the right ~30 tools at
the right granularity with the right invariants, and that took domain experts, not
protocol work.
""")


# ============================================================== 12. VERIFICATION

s = new_slide()
eyebrow(s, "Key component · 3 of 4")
heading(s, "Verification and confidence routing: where the business model lives")

text(s, "The single biggest change from SaaS: we move from syntactic validation (is this "
        "well-formed?) to semantic verification (is this well-formed answer correct?).",
     ML, 1.88, 11.3, 0.5, size=14.5, color=BODY, line=1.35)

cw3 = 3.79
xs = ML
tiers = [
    ("Tier 1 · Hard invariants", TINT_W,
     "Deterministic, non-negotiable, run before anything commits. Balance, trial-balance "
     "ties, subledger-to-control ties, reconciliation delta, duplicate detection, period "
     "locks. A failure here is a hard block, never a warning."),
    ("Tier 2 · Statistical checks", PANEL,
     "Variance against prior period, unusual vendor/account pairings, distributional "
     "anomalies. Catches the well-formed-but-wrong class that invariants pass. Produces "
     "a score, not a verdict."),
    ("Tier 3 · Routing", PANEL,
     "Above threshold: auto-post. Below: into the review queue with full context. "
     "Confidence is derived from verification signals and agreement between independent "
     "passes — never from asking the model how sure it is."),
]
for name, fill, body in tiers:
    box(s, xs, 2.55, cw3, 2.65, fill=fill, edge=WARM if fill is TINT_W else LINE)
    text(s, name, xs + 0.26, 2.78, cw3 - 0.5, 0.32, size=13.5, color=INK, bold=True)
    text(s, body, xs + 0.26, 3.2, cw3 - 0.5, 1.9, size=12.5, color=BODY, line=1.4)
    xs += cw3 + 0.28

box(s, ML, 5.42, CW, 1.35, fill=PANEL2)
text(s, "Models are badly calibrated and reliably overconfident. Asking one for a "
        "confidence score is not a control.", ML + 0.32, 5.62, 6.6, 0.6, size=13.5,
     color=INK, bold=True, line=1.3)
text(s, "Every correction a reviewer makes becomes an explicit, auditable rule first and a "
        "retrieval example second — because rules can be shown to an auditor and embeddings "
        "cannot. The queue is our service delivery, our data pipeline and our product spec "
        "at the same time.",
     ML + 7.2, 5.62, 4.4, 1.0, size=12, color=BODY, line=1.38)

notes(s, """
This is the technical heart of the deck. If they remember one architecture slide, it should
be this one.

Open with the syntactic-to-semantic line. Then the three tiers, briefly. Then spend your
remaining time on the bottom-left box, because it is the most common failure in the
category and stating it clearly is a credibility signal: models are overconfident, so
self-reported confidence is not a control. We derive confidence from verification signals
and from agreement between independent passes.

Then the commercial translation: the routing threshold IS the business model. Move it up
and margin improves but error rate rises; move it down and we're a services company. The
whole engineering programme is about raising the threshold safely, and that is measurable
week over week.

Say explicitly: this is why our margin slope is a forecast rather than a hope. We can see
the auto-post rate per transaction category and we know what each point is worth.

Likely question: "what's your current auto-post rate?" Have the real number by category,
with the denominator defined. If it's early, give the number and the trend. Never give a
blended number without the denominator — a technical investor will assume you're hiding
the mix.
""")


# ============================================================== 13. EVALS

s = new_slide()
eyebrow(s, "Key component · 4 of 4")
heading(s, "Evals and replay: the artefact competitors cannot copy")

text(s, "A SaaS feature is correct or it is a bug. An agent capability has a distribution. "
        "You stop asking \"does it work\" and start asking \"at what rate, on what input "
        "distribution, with what failure modes.\" That changes engineering practice more "
        "than any other single thing.",
     ML, 1.9, 11.3, 0.75, size=14.5, color=BODY, line=1.35)

items = [
    ("CI now contains statistical tests.", "h"),
    ("A 2% regression may be noise or a catastrophe, and only eval volume tells you which.", "b"),
    ("Bug reports are not reproducible.", "h"),
    ("Full trace capture on every run — prompt, context, tool calls, model version — or we "
     "are blind in production.", "b"),
    ("Model upgrades are breaking changes to behaviour we never specified.", "h"),
    ("Version-pin, shadow-eval every candidate model against the corpus before any swap, "
     "keep a fallback. This is a standing pipeline, not a project.", "b"),
    ("Ground truth comes from real closed books.", "h"),
    ("We replay agents against historical months and score against what the accountant "
     "actually did. Slow, unglamorous, requires domain experts — which is exactly why it "
     "is defensible.", "b"),
]
bullets(s, items, ML, 2.78, 7.4, 4.0, size=13.5, gap=7)

box(s, ML + 7.85, 2.78, 4.02, 3.55, fill=TINT_W, edge=WARM)
text(s, "The roadmap unit changes", ML + 8.13, 3.0, 3.45, 0.3, size=11, color=WARM,
     bold=True, caps=True)
text(s, "SaaS roadmaps listed features.\n\nOurs lists accuracy on a category:\n"
        "\"raise auto-post rate on the ambiguous-vendor bucket from 71% to 88%.\"\n\n"
        "Different planning, different staffing, different definition of done — and a "
        "roadmap whose progress is measured rather than asserted.",
     ML + 8.13, 3.4, 3.45, 2.8, size=12.5, color=BODY, line=1.42)

notes(s, """
Two things to get across.

First, that evals are harder than the product and we know it. Teams underinvest here, lose
the ability to tell whether they're improving, and then optimise on vibes. Our eval corpus
is the most valuable engineering artefact we will build, and it needs a data pipeline plus
domain experts — not a prompt engineer.

Second, the roadmap-unit change on the right. This is genuinely useful to an investor
because it tells them what our board reporting will look like. We will not come to you with
"we shipped feature X." We will come with "auto-post rate on this category moved from 71 to
88, here is what that is worth in gross margin." That's a much better governance
relationship and you should say so.

The competitive point: prompts are copyable in an afternoon. A corpus of thousands of
verified closes, with the reviewer corrections attached, is years of accumulated work tied
to customer relationships. When someone asks what stops a well-funded competitor, this is
the honest answer — this and permission.
""")


# ============================================================== 14. HARD PROBLEMS

s = new_slide()
eyebrow(s, "Technical risk")
heading(s, "The two hard problems, stated plainly")

box(s, ML, 2.0, 5.79, 4.35, fill=PANEL, edge=LINE)
text(s, "1 · Reliability compounds badly", ML + 0.3, 2.24, 5.2, 0.35, size=17,
     color=INK, bold=True)
text(s, "95% per step across 20 steps is 36% end-to-end.", ML + 0.3, 2.68, 5.2, 0.55,
     size=15, color=WARM, bold=True, line=1.25)
text(s, "This is the central engineering problem of the category, and there is no prompt "
        "that fixes it. The only real answer is architectural:",
     ML + 0.3, 3.28, 5.2, 0.62, size=13, color=BODY, line=1.35)
r1 = [
    ("Shorten agentic spans; keep deterministic orchestration between them", "b"),
    ("Make every step independently verifiable", "b"),
    ("Gate before commit, always", "b"),
    ("Prefer many short verified hops to one long autonomous chain", "b"),
]
bullets(s, r1, ML + 0.3, 4.02, 5.2, 1.9, size=12.5, gap=6)
text(s, "Any plan that assumes long autonomous chains without this will fail in production. "
        "Treat it as a filter when you look at our competitors.",
     ML + 0.3, 5.62, 5.2, 0.6, size=12, color=WARM, bold=True, line=1.35)

box(s, ML + 6.11, 2.0, 5.79, 4.35, fill=PANEL, edge=LINE)
text(s, "2 · The security model inverts", ML + 6.41, 2.24, 5.2, 0.35, size=17,
     color=INK, bold=True)
text(s, "An actor with authority, taking instructions from untrusted content.",
     ML + 6.41, 2.68, 5.2, 0.55, size=15, color=WARM, bold=True, line=1.25)
text(s, "In SaaS, code was trusted and users were untrusted. Our agent is neither — it acts "
        "with write authority while reading a vendor's PDF or an inbound email. That is a "
        "genuinely new vulnerability class.",
     ML + 6.41, 3.28, 5.2, 0.9, size=13, color=BODY, line=1.35)
r2 = [
    ("Scoped, least-privilege credentials per task", "b"),
    ("No single tool both reads untrusted content and takes irreversible action", "b"),
    ("Separation of duties in code: an agent cannot approve its own entry", "b"),
    ("Hard period locks; all document-derived text treated as tainted", "b"),
]
bullets(s, r2, ML + 6.41, 4.32, 5.2, 1.9, size=12.5, gap=6)

text(s, "Two more we track and do not minimise. The 60/40 trap: the easy majority of "
        "transactions automates fast and creates false confidence, while the residual holds "
        "most of the labour cost — so we instrument where hours go, not where transactions "
        "are. And migration onto our substrate: large, correctness-critical, and the first "
        "thing every customer needs.",
     ML, 6.48, CW, 0.6, size=12, color=BODY, line=1.32)

notes(s, """
Do not skip this slide and do not soften it. Volunteering the hard problems, with the
mitigations already built, is the strongest credibility move available in an AI pitch —
because every sophisticated investor already knows about compounding error rates and is
waiting to see whether you do.

On reliability: the arithmetic is the argument. 0.95^20 = 0.36. Say the number. Then make
clear the fix is architectural, not a better prompt, and that our architecture (slide 9)
was designed around exactly this — short spans, verification between them, deterministic
orchestration.

Offer the filter explicitly: "when you evaluate others in this space, ask how long their
autonomous chains are and what verifies each step. It's a fast way to tell who has shipped."

On security: prompt injection into an agent with ledger write access and bank access is the
scenario that should worry a CFO, and we should be the ones to raise it. Walk the four
mitigations. The separation-of-duties-in-code point resonates with anyone who knows
financial controls — we applied a human control framework to a software actor.

The 60/40 trap: this is why we measure hours, not transaction counts. A competitor
reporting "we automate 80% of transactions" may have automated 30% of the cost.
""")


# ============================================================== 15. EFFORT ALLOCATION

s = new_slide()
eyebrow(s, "Where the work goes")
heading(s, "Building this is a different engineering shape")

text(s, "In SaaS the system of record was the smallest, most stable, most valuable part of "
        "the codebase, and the interface was the largest and most churn-prone. The ledger "
        "under a major accounting product is a few thousand lines that barely changed in a "
        "decade; the UI around it is hundreds of thousands that changed every sprint.",
     ML, 1.88, 11.4, 0.75, size=14, color=BODY, line=1.35)

hdr = [["Layer", "SaaS", "Agent-native", "Nature of the change"]]
rows = [
    ["System of record", "10–15%", "15–20%", "Same concept, far harder requirements"],
    ["Tool layer (typed actions)", "—", "15–20%", "New — and not the same as a REST API"],
    ["Verification & guardrails", "—", "~15%", "New — no SaaS analogue at all"],
    ["Evals & ground-truth corpus", "—", "10–15%", "New — is our test suite and our spec"],
    ["Durable orchestration", "~2%", "~10%", "Qualitatively different from cron and queues"],
    ["Agent layer (prompts, memory)", "—", "~10%", "New, and smaller than people expect"],
    ["UI", "30–40%", "15–20%", "Shrinks, and changes character entirely"],
    ["Integrations", "10–15%", "~10%", "Similar, plus agent-driven fallback"],
    ["Platform", "~15%", "~15%", "Plus agent identity and cost telemetry"],
]
table(s, hdr + rows, ML, 2.78, CW, [3.3, 1.5, 1.85, 5.24], row_h=0.35, head_h=0.42,
      size=11.5, emphasize_col=2)

box(s, ML, 6.42, CW, 0.85, fill=INK)
text(s, "SaaS spent its engineering budget making a human effective at the interface. "
        "We spend ours making a machine's output provable at the substrate.",
     ML + 0.32, 6.62, CW - 0.64, 0.5, size=14.5, color=WHITE, bold=True, line=1.3)

notes(s, """
Percentages are directional — the shape of such codebases rather than measured data. Say
that out loud; a technical investor will respect the caveat and stop probing the digits.

Two observations to draw out:

The interface absorbed all the heterogeneity in SaaS. Every new segment, edge case and
regulatory variation arrived as a new screen, field or report, while the domain model
absorbed almost none of it. Interface complexity is customer diversity projected onto
deterministic code. Agents absorb that diversity in the model instead, which is why the UI
share roughly halves.

The agent layer is one of the smallest lines. Repeat it here — it's the second time they'll
hear it and it's the most counterintuitive thing in the deck.

The closing line in the dark box is the one-sentence summary of the entire technical
section. Deliver it and stop; it is a good place to take questions before the roadmap.

If asked "so is your headcount lower than a comparable SaaS company?" — no, it's differently
shaped: fewer frontend engineers, more data and infra engineers, plus domain experts inside
the loop. Slide 17.
""")


# ============================================================== 16. ROADMAP

s = new_slide()
eyebrow(s, "Roadmap")
heading(s, "Wrap, then own the record, then own the filing")

pw = 3.79
px = ML
phases = [
    ("Phase 1 · [0–9 mo]", "Wrap the incumbent", TINT_A, RGBColor(0x0A, 0x53, 0x66), [
        "Agent layer over QBO / Xero via API — low trust barrier, fast distribution",
        "Build the substrate in parallel, shadow-mode: our ledger mirrors theirs",
        "Ingestion, tool layer, tier-1 invariants, exception console",
        "5–10 design-partner firms; every close manually reviewed",
        "Goal: eval corpus exists and auto-post rate is measurable by category",
    ]),
    ("Phase 2 · [9–24 mo]", "Own the system of record", PANEL, INK, [
        "Our bitemporal ledger becomes primary; we write out to QBO for the accountant's comfort",
        "Migration and backfill tooling — correctness-critical, needed by every customer",
        "Tier-2 statistical checks; confidence routing live; the threshold starts moving",
        "Durable orchestration of the full close checklist",
        "Goal: hours per close down [X]%, gross margin slope visible per cohort",
    ]),
    ("Phase 3 · [24 mo +]", "Own the output", PANEL, INK, [
        "Filings, sales tax, advisory and reporting on top of the same substrate",
        "Outcome-based pricing at scale; per-entity economics proven",
        "Audit-package export as a product; attestation partnerships",
        "Self-serve for the long-tail SMB once auto-post rate clears the bar",
        "Goal: margin profile converging on software, not services",
    ]),
]
for label, title_, fill, tcol, pts in phases:
    box(s, px, 1.95, pw, 4.35, fill=fill, edge=ACCENT if fill is TINT_A else LINE)
    text(s, label, px + 0.26, 2.18, pw - 0.5, 0.3, size=11, color=tcol, bold=True, caps=True)
    text(s, title_, px + 0.26, 2.52, pw - 0.5, 0.4, size=16.5, color=INK, bold=True)
    bullets(s, [(p, "b") for p in pts], px + 0.26, 3.06, pw - 0.5, 3.1, size=12, gap=8)
    px += pw + 0.28

text(s, "Sequencing logic: we do not have to beat an incumbent ledger on day one. Wrapping "
        "buys distribution and, more importantly, buys the ground-truth data that makes "
        "phase 2 possible. The substrate is built from month one regardless — it cannot be "
        "retrofitted.",
     ML, 6.5, CW, 0.7, size=13, color=BODY, line=1.35)

notes(s, """
The sequencing logic at the bottom is the point of the slide; the bullets are supporting
detail. Two reasons to wrap first: distribution and trust are easier, and — the real reason
— wrapping generates the labelled ground truth we need before we can justify owning the
record.

Be explicit about what we do NOT defer: the substrate. The ledger, provenance and tool
layer are built from month one even while we're wrapping, because they are schema decisions
that cannot be retrofitted. A competitor who ships a QBO wrapper without a substrate looks
faster for a year and then cannot make phase 2 at all. That is the strategic bet in one
sentence.

Flag migration honestly as the most underestimated line item in phase 2. Bringing a
customer's historical books onto our substrate is large, correctness-critical and the first
thing every single customer asks for.

[FILL BEFORE USE] the bracketed durations and the phase-2 hours-per-close target. Use your
real design-partner data; do not present illustrative timings as commitments.
""")


# ============================================================== 17. RESOURCES

s = new_slide()
eyebrow(s, "Resources")
heading(s, "What it takes: a differently-shaped team")

text(s, "Illustrative for the phase-1 to early-phase-2 build. Fewer frontend engineers than "
        "a comparable SaaS company, more data and infrastructure engineering, and domain "
        "experts inside the system rather than beside it.",
     ML, 1.86, 11.3, 0.55, size=13.5, color=BODY, line=1.35)

hdr = [["Function", "Ph. 1", "Ph. 2", "Why this is not a normal SaaS org"]]
rows = [
    ["Ledger / backend platform", "3", "5", "Bitemporal append-only substrate, provenance, migration tooling"],
    ["Tool layer & orchestration", "2", "3", "Typed invariant-enforcing actions; durable long-running workflows"],
    ["Agent / applied ML", "2", "3", "Smaller than expected — prompts are the replaceable part"],
    ["Evals & data engineering", "2", "4", "Ground-truth corpus, replay harness, ingestion quality, entity resolution"],
    ["Verification & controls", "1", "2", "Invariants, anomaly detection, confidence calibration"],
    ["Product & exception-console UI", "2", "3", "One surface, done extremely well — not forty screens"],
    ["Accountants in the loop", "3", "6–10", "Service delivery, labelling function and spec authors, simultaneously"],
    ["Security / compliance / SOC 2", "0.5", "1.5", "Agent identity, least privilege, audit posture as a sales asset"],
    ["GTM (firm-led)", "1", "4", "Sold to firms on gross-margin uplift, not to SMBs on features"],
]
table(s, hdr + rows, ML, 2.44, CW, [3.25, 0.9, 0.9, 6.84], row_h=0.35, head_h=0.4,
      size=11.5)

box(s, ML, 6.08, 5.79, 0.94, fill=PANEL)
text(s, "Non-headcount costs to underwrite", ML + 0.28, 6.24, 5.25, 0.28, size=10.5,
     color=MUTED, bold=True, caps=True)
text(s, "Inference (a real COGS line, not R&D) · data acquisition for the corpus · "
        "SOC 2 and audit readiness · E&O insurance · [licensed CPA of record]",
     ML + 0.28, 6.54, 5.25, 0.42, size=11, color=BODY, line=1.28)

box(s, ML + 6.11, 6.08, 5.79, 0.94, fill=TINT_W)
text(s, "Capital ask", ML + 6.39, 6.24, 5.25, 0.28, size=10.5, color=WARM, bold=True, caps=True)
text(s, "[$X] for [N] months to reach [auto-post rate / entities / margin milestone]. "
        "Fill from your own model — do not present illustrative figures as a plan.",
     ML + 6.39, 6.54, 5.25, 0.42, size=11, color=RGBColor(0x7A, 0x3B, 0x08), line=1.28)

notes(s, """
Headcount figures are illustrative shape, not a hiring plan — say so, then talk about the
shape rather than the digits, because the shape is the insight.

Three things to draw out:

"Accountants in the loop" is the largest single line in phase 2 and it is deliberate. They
are simultaneously service delivery, our labelling function, and our spec authors. An
investor may read this as services-y headcount that hurts the multiple. The answer: it
grows sublinearly with revenue and it is what produces the eval corpus, which is the moat.
Show the ratio of entities-per-reviewer improving over cohorts — that ratio is the whole
argument, and it's the number to put on a board slide every month.

Evals and data engineering is larger than applied ML. That is intentional and it is the
opposite of how most AI startups staff. Prompts are cheap; ground truth is not.

Inference sits in COGS, not R&D. Most decks in this category get this wrong. Putting it in
COGS is what makes the gross-margin slope on slide 6 honest.

[FILL BEFORE USE] the capital ask, the milestone it buys, and the non-headcount figures.
Never present the bracketed placeholders live.
""")


# ============================================================== 18. METRICS

s = new_slide()
eyebrow(s, "What to hold us to")
heading(s, "The metrics that decide whether this works")

xs = ML
mets = [
    ("Auto-post rate", "% of transactions posted with zero human touch, reported by "
     "category with the denominator stated. Product quality, margin driver and moat, "
     "in one number."),
    ("Hours per close,\nper entity", "The honest automation measure. Transaction counts "
     "flatter us; hours do not, because the residual tail holds most of the labour cost."),
    ("Gross margin\nby cohort", "Must slope up as cohorts age. A flat slope means we are "
     "a services company with a good story."),
    ("Restatement /\nerror rate", "The metric that can end the company. Tracked against "
     "an absolute ceiling, not a trend."),
]
for name, body in mets:
    w = 2.79
    box(s, xs, 2.0, w, 2.9, fill=PANEL)
    text(s, name, xs + 0.24, 2.24, w - 0.45, 0.75, size=16, color=ACCENT, bold=True, line=1.15)
    text(s, body, xs + 0.24, 3.14, w - 0.45, 1.6, size=12.5, color=BODY, line=1.4)
    xs += w + 0.25

text(s, "Also standing on the board deck: inference cost per entity-month, days-to-close, "
        "entities per reviewer, and auto-post rate on the hardest category rather than the "
        "blended average.",
     ML, 5.1, CW, 0.5, size=13, color=BODY, line=1.35)

box(s, ML, 5.72, CW, 1.4, fill=INK)
text(s, "What we are actually building", ML + 0.32, 5.92, 5.4, 0.3, size=10.5,
     color=RGBColor(0x7A, 0xB8, 0xC8), bold=True, caps=True)
text(s, "Not AI features on top of accounting software. An accounting department, with "
        "software as its body — on a substrate designed so that every number it produces "
        "can be proven, replayed and defended.",
     ML + 0.32, 6.24, CW - 0.64, 0.7, size=15.5, color=WHITE, bold=True, line=1.3)

notes(s, """
Close by handing them the scorecard. Offering the metrics you can be judged on — including
the one that can kill the company — is a stronger close than a projection, and it sets up
the board relationship you actually want.

On auto-post rate: always with the denominator and the category mix. A blended number
without a denominator is the standard way companies in this category flatter themselves,
and a technical investor will assume that's what's happening unless you pre-empt it.

On restatement rate: note deliberately that this one is governed against an absolute
ceiling, not a trend. Everything else we optimise; this one we bound. That distinction is
what a CFO needs to hear before granting write access, and it's what an investor needs to
hear to believe we understand the liability.

Final line is the positioning statement. Say it, stop, and take questions.

Backup slides worth having ready: unit economics / cost-per-close model, the
software-spend-vs-labour-spend sourcing, design-partner results, competitive landscape
(incumbents vs. agent-native), and the liability and attestation structure.
""")


prs.save("/Users/cllu/Projects/aimi/agent-native-accounting-deck.pptx")
print(f"saved · {len(prs.slides.__iter__.__self__._sldIdLst)} slides")
