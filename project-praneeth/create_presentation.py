"""
SEVM Thesis Presentation Generator
===================================
Creates a professional 16:9 PowerPoint presentation for MBA Thesis Defense.
Design: Dark navy + gold accent theme, clean modern layout.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image
import os

# ============================================================
# CONFIGURATION
# ============================================================
SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)

# Color Palette
DARK_BG = RGBColor(0x0D, 0x1B, 0x2A)
NAVY = RGBColor(0x1B, 0x3A, 0x5C)
BLUE = RGBColor(0x2E, 0x75, 0xB6)
GOLD = RGBColor(0xD4, 0xA0, 0x2E)
LIGHT_GOLD = RGBColor(0xE8, 0xBE, 0x5A)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_BG = RGBColor(0xF5, 0xF7, 0xFA)
MED_GRAY = RGBColor(0xA0, 0xA0, 0xA0)
TEXT_DARK = RGBColor(0x1A, 0x1A, 0x2E)
TEXT_GRAY = RGBColor(0x5A, 0x5A, 0x6E)
RED = RGBColor(0xC0, 0x30, 0x30)
GREEN = RGBColor(0x2D, 0x7D, 0x46)
TEAL = RGBColor(0x17, 0xA2, 0xB8)
CARD_BG = RGBColor(0xF0, 0xF4, 0xF8)
INSIGHT_BG = RGBColor(0xE8, 0xEE, 0xF5)
GREEN_CARD = RGBColor(0xE8, 0xF5, 0xE9)
RED_CARD = RGBColor(0xFD, 0xED, 0xED)

CHART_DIR = r"D:\project-praneeth\simulation_results"
OUTPUT = r"D:\project-praneeth\SEVM_Thesis_Presentation.pptx"


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def bg(slide, color):
    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = color


def rect(slide, l, t, w, h, color):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    return s


def circle(slide, l, t, size, color):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, l, t, size, size)
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    return s


def arrow_right(slide, l, t, w, h, color):
    s = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW, l, t, w, h)
    s.fill.solid()
    s.fill.fore_color.rgb = color
    s.line.fill.background()
    return s


def text(slide, l, t, w, h, txt, size=18, color=TEXT_DARK, bold=False,
         font="Calibri", align=PP_ALIGN.LEFT, spacing=1.15):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = txt
    p.font.size = Pt(size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font
    p.alignment = align
    p.space_after = Pt(0)
    return tb


def multitext(slide, l, t, w, h, lines, size=16, color=TEXT_DARK,
              font="Calibri", align=PP_ALIGN.LEFT, spacing=Pt(6)):
    tb = slide.shapes.add_textbox(l, t, w, h)
    tf = tb.text_frame
    tf.word_wrap = True
    for i, item in enumerate(lines):
        if isinstance(item, dict):
            txt = item.get("t", "")
            sz = item.get("s", size)
            clr = item.get("c", color)
            bld = item.get("b", False)
            ital = item.get("i", False)
            sp = item.get("sp", spacing)
        else:
            txt, sz, clr, bld, ital, sp = item, size, color, False, False, spacing

        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = txt
        p.font.size = Pt(sz)
        p.font.color.rgb = clr
        p.font.bold = bld
        p.font.italic = ital
        p.font.name = font
        p.alignment = align
        p.space_after = sp
    return tb


def accent_bar(slide):
    rect(slide, Inches(0), Inches(0), Inches(0.3), SLIDE_HEIGHT, DARK_BG)
    rect(slide, Inches(0.3), Inches(0), Inches(0.05), SLIDE_HEIGHT, GOLD)


def title_bar(slide, title):
    rect(slide, Inches(0.35), Inches(0), SLIDE_WIDTH, Inches(0.85), LIGHT_BG)
    text(slide, Inches(0.7), Inches(0.12), Inches(12), Inches(0.6),
         title, size=24, color=DARK_BG, bold=True)


def slide_number(slide, num, total):
    text(slide, Inches(12.3), Inches(7.1), Inches(0.9), Inches(0.3),
         f"{num}/{total}", size=9, color=MED_GRAY, align=PP_ALIGN.RIGHT)


def add_centered_chart(slide, path, top, max_h=Inches(4.6)):
    if not os.path.exists(path):
        text(slide, Inches(2), top, Inches(8), Inches(1),
             f"[Chart not found: {os.path.basename(path)}]", size=14, color=RED)
        return
    with Image.open(path) as img:
        iw, ih = img.size
    aspect = iw / ih
    pic_h = max_h
    pic_w = int(pic_h * aspect)
    if pic_w > Inches(11.5):
        pic_w = Inches(11.5)
        pic_h = int(pic_w / aspect)
    left = (SLIDE_WIDTH - pic_w) // 2
    slide.shapes.add_picture(path, left, top, pic_w, pic_h)


def takeaway_bar(slide, headline, detail=None, y=Inches(6.0)):
    rect(slide, Inches(0.35), y, Inches(13.0), Inches(1.4), INSIGHT_BG)
    rect(slide, Inches(0.35), y, Inches(0.05), Inches(1.4), GOLD)
    text(slide, Inches(0.65), y + Inches(0.08), Inches(2), Inches(0.25),
         "KEY TAKEAWAY", size=9, color=GOLD, bold=True)
    text(slide, Inches(0.65), y + Inches(0.38), Inches(12.3), Inches(0.5),
         headline, size=14, color=TEXT_DARK, bold=True)
    if detail:
        text(slide, Inches(0.65), y + Inches(0.75), Inches(12.3), Inches(0.5),
             detail, size=12, color=TEXT_GRAY)


# ============================================================
# BUILD PRESENTATION
# ============================================================
prs = Presentation()
prs.slide_width = SLIDE_WIDTH
prs.slide_height = SLIDE_HEIGHT
BLANK = prs.slide_layouts[6]

TOTAL_SLIDES = 22  # for numbering

# ----------------------------------------------------------
# SLIDE 1: TITLE
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, DARK_BG)
rect(sl, Inches(0), Inches(0), Inches(0.1), SLIDE_HEIGHT, GOLD)
# Decorative lines top-right
rect(sl, Inches(9.5), Inches(0.6), Inches(3.5), Inches(0.025), BLUE)
rect(sl, Inches(10.5), Inches(0.85), Inches(2.5), Inches(0.025), GOLD)
rect(sl, Inches(11.3), Inches(1.1), Inches(1.7), Inches(0.025), RGBColor(0x1B, 0x3A, 0x5C))
# Tag
text(sl, Inches(1.0), Inches(1.2), Inches(4), Inches(0.4),
     "MBA THESIS DEFENSE", size=13, color=GOLD, bold=True, font="Calibri")
# Title
text(sl, Inches(1.0), Inches(2.0), Inches(11), Inches(1.5),
     "Stochastic Earned Value Management", size=46, color=WHITE, bold=True,
     font="Calibri Light")
# Gold line
rect(sl, Inches(1.0), Inches(3.7), Inches(3), Inches(0.04), GOLD)
# Subtitle
multitext(sl, Inches(1.0), Inches(4.1), Inches(11), Inches(1.5), [
    {"t": "Integrating Monte Carlo Simulation for Enhanced Probabilistic", "s": 20, "c": MED_GRAY},
    {"t": "Cost and Schedule Forecasting in EPC Projects", "s": 20, "c": MED_GRAY},
], font="Calibri")
# Author
text(sl, Inches(1.0), Inches(6.0), Inches(10), Inches(0.4),
     "Praneeth   |   ESLSCA School of Business", size=15, color=RGBColor(0x80, 0x80, 0x80))
# Bottom accent
rect(sl, Inches(0), SLIDE_HEIGHT - Inches(0.05), SLIDE_WIDTH, Inches(0.05), GOLD)


# ----------------------------------------------------------
# SLIDE 2: AGENDA
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, WHITE)
accent_bar(sl)
title_bar(sl, "Presentation Overview")

agenda = [
    ("01", "The Problem", "Why EPC projects fail at cost forecasting"),
    ("02", "The Solution", "Stochastic EVM with Monte Carlo simulation"),
    ("03", "Methodology", "Research design and simulation framework"),
    ("04", "Key Findings", "Empirical results from 960,000 simulations"),
    ("05", "Recommendations", "Practical implications for the EPC industry"),
]
for i, (num, title_txt, desc) in enumerate(agenda):
    y = Inches(1.3) + Inches(1.1) * i
    rect(sl, Inches(0.8), y, Inches(11.3), Inches(0.9), RGBColor(0xFA, 0xFA, 0xFC))
    c = circle(sl, Inches(1.1), y + Inches(0.15), Inches(0.6), DARK_BG)
    text(sl, Inches(1.1), y + Inches(0.2), Inches(0.6), Inches(0.5),
         num, size=15, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    text(sl, Inches(2.0), y + Inches(0.08), Inches(5), Inches(0.4),
         title_txt, size=19, color=DARK_BG, bold=True)
    text(sl, Inches(2.0), y + Inches(0.48), Inches(9), Inches(0.35),
         desc, size=13, color=TEXT_GRAY)
slide_number(sl, 2, TOTAL_SLIDES)


# ----------------------------------------------------------
# SLIDE 3: SECTION — THE PROBLEM
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, DARK_BG)
rect(sl, Inches(1.5), Inches(3.1), Inches(2), Inches(0.035), GOLD)
text(sl, Inches(1.5), Inches(1.8), Inches(4), Inches(1.2),
     "01", size=60, color=GOLD, bold=True, font="Calibri Light")
text(sl, Inches(1.5), Inches(3.4), Inches(10), Inches(1.0),
     "THE PROBLEM", size=38, color=WHITE, bold=True, font="Calibri Light")
text(sl, Inches(1.5), Inches(4.8), Inches(10), Inches(0.6),
     "Why the EPC industry needs better forecasting tools", size=18, color=MED_GRAY)
rect(sl, Inches(0), SLIDE_HEIGHT - Inches(0.05), SLIDE_WIDTH, Inches(0.05), GOLD)


# ----------------------------------------------------------
# SLIDE 4: EPC CHALLENGE
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, WHITE)
accent_bar(sl)
title_bar(sl, "The EPC Industry Challenge")

stats = [
    ("65%", "of large EPC projects\nexceed budget by 25%+", RED),
    ("80%", "schedule delays are\nshockingly common", RGBColor(0xE0, 0x6C, 0x00)),
    ("$942M", "total portfolio value\nstudied in this research", BLUE),
]
for i, (val, desc, clr) in enumerate(stats):
    x = Inches(0.8) + Inches(3.9) * i
    rect(sl, x, Inches(1.3), Inches(3.5), Inches(2.4), CARD_BG)
    rect(sl, x, Inches(1.3), Inches(3.5), Inches(0.06), clr)
    text(sl, x, Inches(1.65), Inches(3.5), Inches(0.9),
         val, size=48, color=clr, bold=True, font="Calibri Light", align=PP_ALIGN.CENTER)
    text(sl, x + Inches(0.2), Inches(2.7), Inches(3.1), Inches(0.8),
         desc, size=14, color=TEXT_GRAY, align=PP_ALIGN.CENTER)

multitext(sl, Inches(0.8), Inches(4.2), Inches(11.5), Inches(3.0), [
    {"t": "Traditional EVM uses deterministic formulas producing a single-point cost estimate.", "s": 16, "c": TEXT_DARK, "b": False},
    {"t": "It assumes past performance will linearly dictate the future.", "s": 16, "c": TEXT_DARK},
    {"t": "", "s": 8, "c": TEXT_DARK},
    {"t": "But EPC projects are volatile \u2014 subject to supply chain disruptions, labor shortages,", "s": 16, "c": TEXT_DARK},
    {"t": "commodity price spikes, and engineering rework. A single number creates a false sense of certainty.", "s": 16, "c": TEXT_DARK},
    {"t": "", "s": 8, "c": TEXT_DARK},
    {"t": "\"Why do we keep getting it wrong, and how can we fix it?\"", "s": 18, "c": DARK_BG, "b": True, "i": True},
])
slide_number(sl, 4, TOTAL_SLIDES)


# ----------------------------------------------------------
# SLIDE 5: DETERMINISTIC vs STOCHASTIC
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, WHITE)
accent_bar(sl)
title_bar(sl, "Deterministic EVM vs. Stochastic EVM")

# Left card - Deterministic
rect(sl, Inches(0.7), Inches(1.2), Inches(5.6), Inches(5.8), RED_CARD)
rect(sl, Inches(0.7), Inches(1.2), Inches(5.6), Inches(0.06), RED)
text(sl, Inches(0.9), Inches(1.45), Inches(5.2), Inches(0.4),
     "TRADITIONAL EVM", size=19, color=RED, bold=True)
text(sl, Inches(0.9), Inches(1.95), Inches(5.2), Inches(0.4),
     "EAC = BAC / CPI", size=22, color=TEXT_DARK, bold=True, font="Consolas")

det_items = [
    "Single-point estimate only",
    "Assumes linear future performance",
    "No risk quantification",
    "No confidence intervals",
    "\"The project will cost $150M\"",
]
for j, item in enumerate(det_items):
    y = Inches(2.7) + Inches(0.6) * j
    circle(sl, Inches(1.1), y + Inches(0.1), Inches(0.14), RED)
    text(sl, Inches(1.5), y, Inches(4.5), Inches(0.5),
         item, size=14, color=TEXT_DARK)

# Right card - Stochastic
rect(sl, Inches(7.0), Inches(1.2), Inches(5.6), Inches(5.8), GREEN_CARD)
rect(sl, Inches(7.0), Inches(1.2), Inches(5.6), Inches(0.06), GREEN)
text(sl, Inches(7.2), Inches(1.45), Inches(5.2), Inches(0.4),
     "STOCHASTIC EVM (SEVM)", size=19, color=GREEN, bold=True)
text(sl, Inches(7.2), Inches(1.95), Inches(5.2), Inches(0.4),
     "Monte Carlo Simulation", size=22, color=TEXT_DARK, bold=True, font="Consolas")

sevm_items = [
    "Probability distribution of outcomes",
    "Models uncertainty with distributions",
    "Risk-aware decision support",
    "P10 / P50 / P90 confidence levels",
    "\"80% chance cost is $140M\u2013$160M\"",
]
for j, item in enumerate(sevm_items):
    y = Inches(2.7) + Inches(0.6) * j
    circle(sl, Inches(7.4), y + Inches(0.1), Inches(0.14), GREEN)
    text(sl, Inches(7.8), y, Inches(4.5), Inches(0.5),
         item, size=14, color=TEXT_DARK)

# VS circle
circle(sl, Inches(6.05), Inches(3.4), Inches(1.1), DARK_BG)
text(sl, Inches(6.05), Inches(3.6), Inches(1.1), Inches(0.7),
     "VS", size=22, color=WHITE, bold=True, align=PP_ALIGN.CENTER)

slide_number(sl, 5, TOTAL_SLIDES)


# ----------------------------------------------------------
# SLIDE 6: SECTION — THE SOLUTION
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, DARK_BG)
rect(sl, Inches(1.5), Inches(3.1), Inches(2), Inches(0.035), GOLD)
text(sl, Inches(1.5), Inches(1.8), Inches(4), Inches(1.2),
     "02", size=60, color=GOLD, bold=True, font="Calibri Light")
text(sl, Inches(1.5), Inches(3.4), Inches(10), Inches(1.0),
     "THE SOLUTION", size=38, color=WHITE, bold=True, font="Calibri Light")
text(sl, Inches(1.5), Inches(4.8), Inches(10), Inches(0.6),
     "Stochastic EVM with Monte Carlo Simulation", size=18, color=MED_GRAY)
rect(sl, Inches(0), SLIDE_HEIGHT - Inches(0.05), SLIDE_WIDTH, Inches(0.05), GOLD)


# ----------------------------------------------------------
# SLIDE 7: HOW SEVM WORKS (4-step process)
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, WHITE)
accent_bar(sl)
title_bar(sl, "How SEVM Works")

steps = [
    ("1", "COLLECT", "Historical EVM data\n(PV, EV, AC) over\nproject lifecycle"),
    ("2", "FIT", "Fit lognormal distribution\nto CPI & normal\ndistribution to SPI"),
    ("3", "SIMULATE", "Run 10,000 Monte Carlo\niterations sampling from\nfitted distributions"),
    ("4", "FORECAST", "Generate probability\ndistribution of EAC\nwith confidence intervals"),
]
for i, (num, title_txt, desc) in enumerate(steps):
    x = Inches(0.7) + Inches(3.1) * i
    rect(sl, x, Inches(1.2), Inches(2.8), Inches(3.6), CARD_BG)
    circle(sl, x + Inches(1.05), Inches(1.4), Inches(0.7), DARK_BG)
    text(sl, x + Inches(1.05), Inches(1.45), Inches(0.7), Inches(0.6),
         num, size=22, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    text(sl, x + Inches(0.1), Inches(2.3), Inches(2.6), Inches(0.4),
         title_txt, size=17, color=DARK_BG, bold=True, align=PP_ALIGN.CENTER)
    text(sl, x + Inches(0.1), Inches(2.8), Inches(2.6), Inches(1.6),
         desc, size=12, color=TEXT_GRAY, align=PP_ALIGN.CENTER)
    if i < 3:
        arrow_right(sl, x + Inches(2.85), Inches(2.7), Inches(0.22), Inches(0.35), GOLD)

# Insight box
rect(sl, Inches(0.7), Inches(5.2), Inches(11.9), Inches(2.0), INSIGHT_BG)
rect(sl, Inches(0.7), Inches(5.2), Inches(0.05), Inches(2.0), GOLD)
text(sl, Inches(1.0), Inches(5.3), Inches(5), Inches(0.3),
     "WHY LOGNORMAL FOR CPI?", size=11, color=GOLD, bold=True)
multitext(sl, Inches(1.0), Inches(5.65), Inches(11.3), Inches(1.4), [
    {"t": "Cost variations in EPC projects are right-skewed: small overruns are common, but extreme", "s": 14, "c": TEXT_DARK},
    {"t": "cost blowouts (the 'fat tail') are statistically possible. A lognormal distribution captures", "s": 14, "c": TEXT_DARK},
    {"t": "this reality. Costs cannot go below zero, ruling out the symmetrical normal distribution.", "s": 14, "c": TEXT_DARK},
])
slide_number(sl, 7, TOTAL_SLIDES)


# ----------------------------------------------------------
# SLIDE 8: SECTION — METHODOLOGY
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, DARK_BG)
rect(sl, Inches(1.5), Inches(3.1), Inches(2), Inches(0.035), GOLD)
text(sl, Inches(1.5), Inches(1.8), Inches(4), Inches(1.2),
     "03", size=60, color=GOLD, bold=True, font="Calibri Light")
text(sl, Inches(1.5), Inches(3.4), Inches(10), Inches(1.0),
     "RESEARCH METHODOLOGY", size=38, color=WHITE, bold=True, font="Calibri Light")
text(sl, Inches(1.5), Inches(4.8), Inches(10), Inches(0.6),
     "Quantitative analysis with Monte Carlo simulation", size=18, color=MED_GRAY)
rect(sl, Inches(0), SLIDE_HEIGHT - Inches(0.05), SLIDE_WIDTH, Inches(0.05), GOLD)


# ----------------------------------------------------------
# SLIDE 9: RESEARCH DESIGN
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, WHITE)
accent_bar(sl)
title_bar(sl, "Research Design & Dataset")

metrics = [
    ("12", "EPC Projects", DARK_BG),
    ("5", "Industry Sectors", NAVY),
    ("$942M", "Total Portfolio", BLUE),
    ("960K", "MC Iterations", TEAL),
]
for i, (val, label, clr) in enumerate(metrics):
    x = Inches(0.7) + Inches(3.1) * i
    rect(sl, x, Inches(1.15), Inches(2.8), Inches(1.5), clr)
    text(sl, x, Inches(1.25), Inches(2.8), Inches(0.8),
         val, size=36, color=GOLD, bold=True, align=PP_ALIGN.CENTER, font="Calibri Light")
    text(sl, x, Inches(1.95), Inches(2.8), Inches(0.45),
         label, size=13, color=WHITE, align=PP_ALIGN.CENTER)

details = [
    "Sectors: Oil & Gas, Power Generation, Water Treatment, Industrial, Renewable Energy",
    "Budget range: $15M to $200M   |   Duration range: 12 to 42 months",
    "11 of 12 projects experienced actual cost overruns (realistic dataset)",
    "Forecast points evaluated: 20%, 30%, 40%, 50%, 60%, 70%, 80%, 90% completion",
    "10,000 Monte Carlo iterations per forecast point for convergence stability",
    "Statistical validation: Wilcoxon signed-rank test (non-parametric, paired)",
]
for j, line in enumerate(details):
    y = Inches(3.1) + Inches(0.58) * j
    circle(sl, Inches(0.9), y + Inches(0.12), Inches(0.12), BLUE)
    text(sl, Inches(1.3), y, Inches(11), Inches(0.45),
         line, size=14, color=TEXT_DARK)

slide_number(sl, 9, TOTAL_SLIDES)


# ----------------------------------------------------------
# SLIDE 10: PROJECT PORTFOLIO TABLE
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, WHITE)
accent_bar(sl)
title_bar(sl, "Project Portfolio Overview")

data = [
    ["ID", "BAC ($M)", "Actual ($M)", "Overrun", "CPI", "Complexity"],
    ["P01", "$45.0", "$51.2", "+13.9%", "0.878", "High"],
    ["P02", "$120.0", "$143.4", "+19.5%", "0.837", "High"],
    ["P03", "$28.0", "$30.8", "+9.8%", "0.910", "Medium"],
    ["P04", "$85.0", "$104.5", "+22.9%", "0.814", "High"],
    ["P05", "$15.0", "$17.1", "+14.3%", "0.875", "Low"],
    ["P06", "$62.0", "$60.8", "-1.9%", "1.019", "Medium"],
    ["P07", "$200.0", "$223.1", "+11.6%", "0.896", "High"],
    ["P08", "$35.0", "$35.8", "+2.2%", "0.978", "Medium"],
    ["P09", "$95.0", "$122.4", "+28.8%", "0.776", "High"],
    ["P10", "$22.0", "$24.2", "+10.1%", "0.909", "Low"],
    ["P11", "$55.0", "$59.8", "+8.6%", "0.920", "Medium"],
    ["P12", "$180.0", "$199.1", "+10.6%", "0.904", "High"],
]
rows, cols = len(data), len(data[0])
tbl = sl.shapes.add_table(rows, cols, Inches(0.7), Inches(1.05), Inches(11.8), Inches(6.1)).table

for r in range(rows):
    for c in range(cols):
        cell = tbl.cell(r, c)
        cell.text = data[r][c]
        for p in cell.text_frame.paragraphs:
            p.font.size = Pt(12)
            p.font.name = "Calibri"
            p.alignment = PP_ALIGN.CENTER
            if r == 0:
                p.font.bold = True
                p.font.color.rgb = WHITE
            else:
                p.font.color.rgb = TEXT_DARK
        if r == 0:
            cell.fill.solid()
            cell.fill.fore_color.rgb = DARK_BG
        elif r % 2 == 0:
            cell.fill.solid()
            cell.fill.fore_color.rgb = LIGHT_BG
        else:
            cell.fill.solid()
            cell.fill.fore_color.rgb = WHITE

slide_number(sl, 10, TOTAL_SLIDES)


# ----------------------------------------------------------
# SLIDE 11: SECTION — KEY FINDINGS
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, DARK_BG)
rect(sl, Inches(1.5), Inches(3.1), Inches(2), Inches(0.035), GOLD)
text(sl, Inches(1.5), Inches(1.8), Inches(4), Inches(1.2),
     "04", size=60, color=GOLD, bold=True, font="Calibri Light")
text(sl, Inches(1.5), Inches(3.4), Inches(10), Inches(1.0),
     "KEY FINDINGS", size=38, color=WHITE, bold=True, font="Calibri Light")
text(sl, Inches(1.5), Inches(4.8), Inches(10), Inches(0.6),
     "Empirical results from the SEVM simulation", size=18, color=MED_GRAY)
rect(sl, Inches(0), SLIDE_HEIGHT - Inches(0.05), SLIDE_WIDTH, Inches(0.05), GOLD)


# ----------------------------------------------------------
# SLIDE 12: FINDING 1 — ACCURACY COMPARISON (fig2)
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, WHITE)
accent_bar(sl)
title_bar(sl, "Finding 1: Forecast Accuracy Comparison")
add_centered_chart(sl, os.path.join(CHART_DIR, "fig2_accuracy_comparison.png"),
                   Inches(1.0), Inches(4.6))
takeaway_bar(sl,
    "SEVM (P50) achieves 3.66% MAPE vs. Deterministic Composite's 4.21% \u2014 a 0.55pp improvement.",
    "SEVM consistently outperforms the composite deterministic formula across all project completion stages.")
slide_number(sl, 12, TOTAL_SLIDES)


# ----------------------------------------------------------
# SLIDE 13: FINDING 2 — S-CURVE (fig3)
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, WHITE)
accent_bar(sl)
title_bar(sl, "Finding 2: S-Curves with SEVM Probabilistic Forecast")
add_centered_chart(sl, os.path.join(CHART_DIR, "fig3_scurve_sevm.png"),
                   Inches(1.0), Inches(4.6))
takeaway_bar(sl,
    "SEVM provides a probabilistic confidence band (P10\u2013P90) instead of a single-point forecast line.",
    "The 80% confidence band captures the range of likely outcomes, enabling better contingency planning.")
slide_number(sl, 13, TOTAL_SLIDES)


# ----------------------------------------------------------
# SLIDE 14: FINDING 3 — MC DISTRIBUTION (fig4)
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, WHITE)
accent_bar(sl)
title_bar(sl, "Finding 3: Monte Carlo EAC Distribution (10,000 Iterations)")
add_centered_chart(sl, os.path.join(CHART_DIR, "fig4_eac_distribution.png"),
                   Inches(1.0), Inches(4.6))
takeaway_bar(sl,
    "The right-skewed distribution shows extreme cost overruns are possible but unlikely \u2014 captured by SEVM.",
    "Det. EAC ($102.5M) sits near P50, but misses the tail risk SEVM reveals up to P90 ($116.4M).")
slide_number(sl, 14, TOTAL_SLIDES)


# ----------------------------------------------------------
# SLIDE 15: FINDING 4 — CALIBRATION (big numbers)
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, WHITE)
accent_bar(sl)
title_bar(sl, "Finding 4: SEVM Calibration \u2014 Can We Trust the Ranges?")

# 80% card
rect(sl, Inches(1.0), Inches(1.3), Inches(5.2), Inches(4.0), GREEN_CARD)
rect(sl, Inches(1.0), Inches(1.3), Inches(5.2), Inches(0.07), GREEN)
text(sl, Inches(1.0), Inches(1.7), Inches(5.2), Inches(1.3),
     "77.1%", size=72, color=GREEN, bold=True, align=PP_ALIGN.CENTER, font="Calibri Light")
text(sl, Inches(1.0), Inches(3.0), Inches(5.2), Inches(0.4),
     "80% Confidence Interval Calibration", size=17, color=DARK_BG, bold=True, align=PP_ALIGN.CENTER)
multitext(sl, Inches(1.3), Inches(3.6), Inches(4.6), Inches(1.4), [
    {"t": "Target: 80%  |  Achieved: 77.1%", "s": 13, "c": TEXT_GRAY},
    {"t": "", "s": 6},
    {"t": "The actual project cost fell within the SEVM", "s": 12, "c": TEXT_GRAY},
    {"t": "80% confidence band 77.1% of the time.", "s": 12, "c": TEXT_GRAY},
], align=PP_ALIGN.CENTER)

# 50% card
rect(sl, Inches(7.0), Inches(1.3), Inches(5.2), Inches(4.0), CARD_BG)
rect(sl, Inches(7.0), Inches(1.3), Inches(5.2), Inches(0.07), BLUE)
text(sl, Inches(7.0), Inches(1.7), Inches(5.2), Inches(1.3),
     "53.1%", size=72, color=BLUE, bold=True, align=PP_ALIGN.CENTER, font="Calibri Light")
text(sl, Inches(7.0), Inches(3.0), Inches(5.2), Inches(0.4),
     "50% Confidence Interval Calibration", size=17, color=DARK_BG, bold=True, align=PP_ALIGN.CENTER)
multitext(sl, Inches(7.3), Inches(3.6), Inches(4.6), Inches(1.4), [
    {"t": "Target: 50%  |  Achieved: 53.1%", "s": 13, "c": TEXT_GRAY},
    {"t": "", "s": 6},
    {"t": "The actual cost fell within the narrower", "s": 12, "c": TEXT_GRAY},
    {"t": "P25\u2013P75 band 53.1% of the time.", "s": 12, "c": TEXT_GRAY},
], align=PP_ALIGN.CENTER)

# Insight bar
rect(sl, Inches(0.7), Inches(5.7), Inches(11.9), Inches(1.5), INSIGHT_BG)
rect(sl, Inches(0.7), Inches(5.7), Inches(0.05), Inches(1.5), GOLD)
text(sl, Inches(1.0), Inches(5.8), Inches(5), Inches(0.25),
     "WHAT THIS MEANS", size=10, color=GOLD, bold=True)
multitext(sl, Inches(1.0), Inches(6.1), Inches(11.3), Inches(1.0), [
    {"t": "SEVM's confidence intervals are remarkably well-calibrated. When the model says", "s": 14, "c": TEXT_DARK},
    {"t": "\"80% chance the cost falls in this range,\" it is correct ~77% of the time.", "s": 14, "c": TEXT_DARK},
    {"t": "Decision-makers can trust these ranges for contingency planning.", "s": 14, "c": TEXT_DARK, "b": True},
])
slide_number(sl, 15, TOTAL_SLIDES)


# ----------------------------------------------------------
# SLIDE 16: FINDING 5 — COMPLEXITY (fig5)
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, WHITE)
accent_bar(sl)
title_bar(sl, "Finding 5: SEVM Accuracy Advantage by Project Complexity")
add_centered_chart(sl, os.path.join(CHART_DIR, "fig5_complexity_effect.png"),
                   Inches(1.0), Inches(4.6))
takeaway_bar(sl,
    "SEVM's value scales with project complexity \u2014 high complexity projects show the widest advantage range.",
    "For high-complexity projects (offshore, chemical plants), SEVM confidence intervals achieved 96% calibration.")
slide_number(sl, 16, TOTAL_SLIDES)


# ----------------------------------------------------------
# SLIDE 17: FINDING 6 — CPI VARIABILITY (fig7)
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, WHITE)
accent_bar(sl)
title_bar(sl, "Finding 6: Evidence of Stochastic Behavior in CPI")
add_centered_chart(sl, os.path.join(CHART_DIR, "fig7_cpi_variability.png"),
                   Inches(1.0), Inches(4.6))
takeaway_bar(sl,
    "Period-by-period CPI shows clear stochastic behavior \u2014 not the constant pattern traditional EVM assumes.",
    "This volatile, non-linear CPI behavior is exactly why deterministic single-point forecasts are structurally flawed.")
slide_number(sl, 17, TOTAL_SLIDES)


# ----------------------------------------------------------
# SLIDE 18: FINDING 7 — PHASE EFFECT (fig8)
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, WHITE)
accent_bar(sl)
title_bar(sl, "Finding 7: SEVM Accuracy Advantage by Project Phase")
add_centered_chart(sl, os.path.join(CHART_DIR, "fig8_phase_effect.png"),
                   Inches(1.0), Inches(4.6))
takeaway_bar(sl,
    "SEVM provides its maximum accuracy advantage vs. composite EVM at the 50% completion mark (+1.05pp).",
    "The 'sweet spot': enough historical data to fit distributions, yet enough remaining uncertainty where SEVM excels.")
slide_number(sl, 18, TOTAL_SLIDES)


# ----------------------------------------------------------
# SLIDE 19: FINDING 8 — CDF (fig6)
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, WHITE)
accent_bar(sl)
title_bar(sl, "Finding 8: Cumulative Distribution Function of SEVM Forecast")
add_centered_chart(sl, os.path.join(CHART_DIR, "fig6_cdf.png"),
                   Inches(1.0), Inches(4.6))
takeaway_bar(sl,
    "The CDF lets managers read exact probabilities: \"What is the chance of staying under $110M?\"",
    "This transforms project reporting from a single number to actionable probability-based decision support.")
slide_number(sl, 19, TOTAL_SLIDES)


# ----------------------------------------------------------
# SLIDE 20: HYPOTHESIS SUMMARY
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, WHITE)
accent_bar(sl)
title_bar(sl, "Hypothesis Results Summary")

hyps = [
    ("H1", "NOT SUPPORTED", "SEVM improves cost accuracy by -1.0pp vs BAC/CPI formula",
     RED, "However, SEVM outperforms the more complex Composite formula by +0.55pp"),
    ("H2", "SUPPORTED", "Schedule uncertainty linked to composite formula improvement of +0.5pp",
     GREEN, ""),
    ("H3", "SUPPORTED", "SEVM advantage scales with complexity (High > Medium > Low)",
     GREEN, "High complexity projects: 96% calibration accuracy"),
    ("H4", "PARTIALLY\nSUPPORTED", "Maximum SEVM advantage at 50% completion mark",
     GOLD, "Early stages have insufficient data; late stages have low uncertainty"),
    ("H5", "SUPPORTED", "Calibration accuracy improves with more data points",
     GREEN, ""),
]
for i, (hid, status, desc, clr, note) in enumerate(hyps):
    y = Inches(1.15) + Inches(1.1) * i
    rect(sl, Inches(0.7), y, Inches(11.8), Inches(0.95), RGBColor(0xFA, 0xFA, 0xFC))
    # H label
    rect(sl, Inches(0.7), y, Inches(0.75), Inches(0.95), DARK_BG)
    text(sl, Inches(0.7), y + Inches(0.25), Inches(0.75), Inches(0.45),
         hid, size=18, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    # Status
    rect(sl, Inches(1.65), y + Inches(0.15), Inches(2.0), Inches(0.65), clr)
    text(sl, Inches(1.65), y + Inches(0.18), Inches(2.0), Inches(0.6),
         status, size=10, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    # Description
    text(sl, Inches(3.9), y + Inches(0.1), Inches(8.3), Inches(0.4),
         desc, size=13, color=TEXT_DARK, bold=True)
    if note:
        text(sl, Inches(3.9), y + Inches(0.52), Inches(8.3), Inches(0.35),
             note, size=11, color=TEXT_GRAY)

# Stats footer
rect(sl, Inches(0.7), Inches(6.65), Inches(11.8), Inches(0.6), INSIGHT_BG)
text(sl, Inches(0.9), Inches(6.68), Inches(6), Inches(0.25),
     "STATISTICAL SIGNIFICANCE (Wilcoxon Signed-Rank Test)", size=10, color=GOLD, bold=True)
text(sl, Inches(0.9), Inches(6.9), Inches(5.5), Inches(0.3),
     "SEVM vs BAC/CPI:  W=1073, p=0.000005 ***", size=11, color=TEXT_DARK, font="Consolas")
text(sl, Inches(6.5), Inches(6.9), Inches(5.5), Inches(0.3),
     "SEVM vs Composite:  W=1915, p=0.131  (ns)", size=11, color=TEXT_DARK, font="Consolas")
slide_number(sl, 20, TOTAL_SLIDES)


# ----------------------------------------------------------
# SLIDE 21: SECTION — RECOMMENDATIONS
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, DARK_BG)
rect(sl, Inches(1.5), Inches(3.1), Inches(2), Inches(0.035), GOLD)
text(sl, Inches(1.5), Inches(1.8), Inches(4), Inches(1.2),
     "05", size=60, color=GOLD, bold=True, font="Calibri Light")
text(sl, Inches(1.5), Inches(3.4), Inches(10), Inches(1.0),
     "RECOMMENDATIONS & CONCLUSION", size=36, color=WHITE, bold=True, font="Calibri Light")
text(sl, Inches(1.5), Inches(4.8), Inches(10), Inches(0.6),
     "Practical implications for EPC organizations", size=18, color=MED_GRAY)
rect(sl, Inches(0), SLIDE_HEIGHT - Inches(0.05), SLIDE_WIDTH, Inches(0.05), GOLD)


# ----------------------------------------------------------
# SLIDE 22: RECOMMENDATIONS
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, WHITE)
accent_bar(sl)
title_bar(sl, "Recommendations for EPC Organizations")

recs = [
    ("COMPLEMENT, NOT REPLACE",
     "Adopt SEVM as a layer on top of traditional EVM. Use EVM for baseline measurement, SEVM for risk-aware forecasting."),
    ("TARGET DEPLOYMENT",
     "Deploy Monte Carlo simulation on high-complexity, high-value projects where calibrated confidence intervals provide the most value."),
    ("INVEST IN DATA HYGIENE",
     "SEVM requires at least 10 periods of high-quality, unmanipulated EVM data to fit reliable probability distributions."),
    ("COMMUNICATE IN RANGES",
     "Train executives to ask 'What is our P50 likely cost and P90 contingency ceiling?' instead of 'What is the final cost?'"),
    ("FOCUS AT MID-PROJECT",
     "If only one major risk assessment is possible, conduct it at the 50% completion mark where SEVM's advantage is greatest."),
]
for i, (title_txt, desc) in enumerate(recs):
    y = Inches(1.15) + Inches(1.15) * i
    circle(sl, Inches(0.8), y + Inches(0.15), Inches(0.55), DARK_BG)
    text(sl, Inches(0.8), y + Inches(0.2), Inches(0.55), Inches(0.45),
         str(i + 1), size=18, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    text(sl, Inches(1.6), y + Inches(0.03), Inches(10.5), Inches(0.35),
         title_txt, size=15, color=DARK_BG, bold=True)
    text(sl, Inches(1.6), y + Inches(0.4), Inches(10.5), Inches(0.65),
         desc, size=13, color=TEXT_GRAY)

slide_number(sl, 22, TOTAL_SLIDES)


# ----------------------------------------------------------
# SLIDE 23: CONCLUSION
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, DARK_BG)
rect(sl, Inches(1.5), Inches(0.7), Inches(2.5), Inches(0.035), GOLD)
text(sl, Inches(1.5), Inches(0.9), Inches(10), Inches(0.8),
     "CONCLUSION", size=36, color=WHITE, bold=True, font="Calibri Light")

conclusions = [
    "EPC projects are fundamentally stochastic \u2014 yet we manage them with deterministic tools.",
    "SEVM with Monte Carlo simulation bridges the gap between cost control and risk management.",
    "The 80% CI calibration (77.1%) proves SEVM ranges are trustworthy for executive decisions.",
    "SEVM provides maximum value for high-complexity projects at the critical mid-project phase.",
    "Transitioning to probabilistic ranges enables evidence-based contingency planning.",
]
for i, point in enumerate(conclusions):
    y = Inches(2.0) + Inches(0.9) * i
    circle(sl, Inches(1.7), y + Inches(0.1), Inches(0.17), GOLD)
    text(sl, Inches(2.2), y, Inches(10), Inches(0.7),
         point, size=18, color=WHITE)

rect(sl, Inches(0), SLIDE_HEIGHT - Inches(0.05), SLIDE_WIDTH, Inches(0.05), GOLD)


# ----------------------------------------------------------
# SLIDE 24: THANK YOU
# ----------------------------------------------------------
sl = prs.slides.add_slide(BLANK)
bg(sl, DARK_BG)
rect(sl, Inches(0), Inches(0), Inches(0.1), SLIDE_HEIGHT, GOLD)
rect(sl, Inches(8.5), Inches(0.8), Inches(4.5), Inches(0.025), BLUE)
rect(sl, Inches(9.5), Inches(1.05), Inches(3.5), Inches(0.025), GOLD)
rect(sl, Inches(10.3), Inches(1.3), Inches(2.7), Inches(0.025), NAVY)

text(sl, Inches(1.5), Inches(2.2), Inches(10.5), Inches(1.5),
     "Thank You", size=60, color=WHITE, bold=True,
     font="Calibri Light", align=PP_ALIGN.CENTER)
rect(sl, Inches(5.5), Inches(3.8), Inches(2.5), Inches(0.035), GOLD)
text(sl, Inches(1.5), Inches(4.2), Inches(10.5), Inches(0.8),
     "Questions & Discussion", size=24, color=MED_GRAY,
     font="Calibri", align=PP_ALIGN.CENTER)
text(sl, Inches(1.5), Inches(5.5), Inches(10.5), Inches(0.4),
     "Praneeth   |   ESLSCA School of Business   |   MBA Thesis",
     size=15, color=RGBColor(0x66, 0x66, 0x66), align=PP_ALIGN.CENTER)
rect(sl, Inches(0), SLIDE_HEIGHT - Inches(0.05), SLIDE_WIDTH, Inches(0.05), GOLD)


# ============================================================
# SAVE
# ============================================================
prs.save(OUTPUT)
print(f"\nPresentation saved: {OUTPUT}")
print(f"Total slides: {len(prs.slides)}")
print("Done!")
