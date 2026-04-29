#!/usr/bin/env python3
"""
Condensed 14-slide PowerPoint presentation for the thesis:
"Enhancing Risk Management Effectiveness in Hybrid Projects
through AI-Driven Strategic Governance: A SEM Study"
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

# ── Color Palette ──────────────────────────────────────────
NAVY      = RGBColor(0x0B, 0x1D, 0x3A)
DARK_BLUE = RGBColor(0x13, 0x2E, 0x5B)
MID_BLUE  = RGBColor(0x1B, 0x4F, 0x8A)
LIGHT_BLUE= RGBColor(0x3A, 0x7C, 0xBD)
SKY_BLUE  = RGBColor(0x5B, 0xA0, 0xD9)
GOLD      = RGBColor(0xE8, 0xA8, 0x38)
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
GREEN_OK  = RGBColor(0x27, 0xAE, 0x60)
ORANGE    = RGBColor(0xF3, 0x9C, 0x12)
TEAL      = RGBColor(0x16, 0xA0, 0x85)
SOFT_TEXT = RGBColor(0xCC, 0xDD, 0xEE)
MUTED     = RGBColor(0x99, 0xAA, 0xBB)

OUT_DIR = r"D:\project-shiva\18-02-26"
CHART_DIR = os.path.join(OUT_DIR, "charts_v2")
os.makedirs(CHART_DIR, exist_ok=True)

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)

TOTAL = 14

# ── Helpers ────────────────────────────────────────────────

def add_bg(slide, color=NAVY):
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = color

def bar(slide, l, t, w, h, c):
    s = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = c; s.line.fill.background()
    return s

def gold_bar(slide):
    bar(slide, 0, 7.42, 13.333, 0.08, GOLD)

def accent_line(slide, l=0.8, t=0.7, w=2.5):
    bar(slide, l, t, w, 0.055, GOLD)

def slide_num(slide, n):
    tb = slide.shapes.add_textbox(Inches(12.3), Inches(7.05), Inches(0.8), Inches(0.35))
    p = tb.text_frame.paragraphs[0]
    p.text = f"{n}/{TOTAL}"; p.font.size = Pt(10); p.font.color.rgb = MUTED; p.alignment = PP_ALIGN.RIGHT

def tb(slide, l, t, w, h, text, sz=16, c=WHITE, bold=False, align=PP_ALIGN.LEFT, ls=1.2):
    bx = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = bx.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]; p.text = text
    p.font.size = Pt(sz); p.font.color.rgb = c; p.font.bold = bold
    p.font.name = "Calibri"; p.alignment = align; p.line_spacing = Pt(sz * ls)
    return bx

def ml(slide, l, t, w, h, lines, sz=13, c=SOFT_TEXT, bullet=False, ls=1.3):
    bx = slide.shapes.add_textbox(Inches(l), Inches(t), Inches(w), Inches(h))
    tf = bx.text_frame; tf.word_wrap = True
    for i, line in enumerate(lines):
        if isinstance(line, str): txt, bld, clr = line, False, c
        else: txt, bld, clr = line[0], line[1] if len(line) > 1 else False, line[2] if len(line) > 2 else c
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.text = ("\u2022  " if bullet else "") + txt
        p.font.size = Pt(sz); p.font.color.rgb = clr; p.font.bold = bld
        p.font.name = "Calibri"; p.alignment = PP_ALIGN.LEFT
        p.line_spacing = Pt(sz * ls); p.space_after = Pt(3)
    return bx

def rrect(slide, l, t, w, h, fc, text="", sz=12, tc=WHITE, bold=False, align=PP_ALIGN.CENTER):
    s = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(l), Inches(t), Inches(w), Inches(h))
    s.fill.solid(); s.fill.fore_color.rgb = fc; s.line.fill.background(); s.shadow.inherit = False
    if text:
        tf = s.text_frame; tf.word_wrap = True; p = tf.paragraphs[0]
        p.text = text; p.font.size = Pt(sz); p.font.color.rgb = tc; p.font.bold = bold
        p.font.name = "Calibri"; p.alignment = align
    return s

def circ(slide, l, t, sz_in, fc, text="", fsz=18, tc=WHITE):
    s = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(l), Inches(t), Inches(sz_in), Inches(sz_in))
    s.fill.solid(); s.fill.fore_color.rgb = fc; s.line.fill.background()
    if text:
        tf = s.text_frame; tf.word_wrap = True; p = tf.paragraphs[0]
        p.text = text; p.font.size = Pt(fsz); p.font.color.rgb = tc; p.font.bold = True
        p.font.name = "Calibri"; p.alignment = PP_ALIGN.CENTER
    return s

def section(slide, title, sub=""):
    accent_line(slide)
    tb(slide, 0.8, 0.85, 10, 0.7, title, 30, GOLD, True)
    if sub: tb(slide, 0.8, 1.45, 10, 0.4, sub, 13, MUTED)

def save_fig(fig, name, dpi=200):
    p = os.path.join(CHART_DIR, f"{name}.png")
    fig.savefig(p, dpi=dpi, bbox_inches='tight', transparent=True, pad_inches=0.1)
    plt.close(fig); return p

def img(slide, path, l, t, w=None, h=None):
    if w and h: slide.shapes.add_picture(path, Inches(l), Inches(t), Inches(w), Inches(h))
    elif w: slide.shapes.add_picture(path, Inches(l), Inches(t), Inches(w))
    else: slide.shapes.add_picture(path, Inches(l), Inches(t))


# ═══════════════════════════════════════════════════════════
# GENERATE CHARTS
# ═══════════════════════════════════════════════════════════
print("Generating charts...")

# ── Demographics combo chart ───────────────────────────────
fig, axes = plt.subplots(2, 2, figsize=(10, 6))
fig.patch.set_alpha(0)
for ax in axes.flat: ax.set_facecolor('none')

# Gender donut
ax = axes[0, 0]
wedges, _, autotexts = ax.pie([46.1, 53.9], labels=None, autopct='%1.1f%%', startangle=90,
    colors=['#1B4F8A', '#E8A838'], pctdistance=0.78, wedgeprops=dict(width=0.45, edgecolor='white', linewidth=2))
for t in autotexts: t.set_fontsize(9); t.set_color('white'); t.set_fontweight('bold')
ax.legend(['Male (46.1%)', 'Female (53.9%)'], loc='lower center', ncol=2, fontsize=7,
          bbox_to_anchor=(0.5, -0.08), frameon=False, labelcolor='white')
ax.set_title('Gender', color='white', fontsize=11, fontweight='bold', pad=8)

# Age bar
ax = axes[0, 1]
ages = ['20-25', '26-35', '36-45', '46-55', '56+']
age_v = [17.7, 22.7, 19.5, 15.6, 24.5]
bars_a = ax.bar(ages, age_v, color=['#1B4F8A','#3A7CBD','#5BA0D9','#E8A838','#16A085'], width=0.6, edgecolor='white', linewidth=0.5)
for b, v in zip(bars_a, age_v): ax.text(b.get_x()+b.get_width()/2, b.get_height()+0.5, f'{v}%', ha='center', color='white', fontsize=8, fontweight='bold')
ax.set_title('Age Group (%)', color='white', fontsize=11, fontweight='bold', pad=8)
ax.tick_params(colors='white', labelsize=8); ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
ax.spines['bottom'].set_color('#555'); ax.spines['left'].set_color('#555')

# AI Usage donut
ax = axes[1, 0]
wedges2, _, autotexts2 = ax.pie([35.4, 32.0, 32.6], labels=None, autopct='%1.1f%%', startangle=90,
    colors=['#E74C3C', '#F39C12', '#27AE60'], pctdistance=0.78, wedgeprops=dict(width=0.45, edgecolor='white', linewidth=2))
for t in autotexts2: t.set_fontsize(9); t.set_color('white'); t.set_fontweight('bold')
ax.legend(['Not Used', 'Limited', 'Moderate'], loc='lower center', ncol=3, fontsize=7,
          bbox_to_anchor=(0.5, -0.08), frameon=False, labelcolor='white')
ax.set_title('AI Usage in Risk Mgmt', color='white', fontsize=11, fontweight='bold', pad=8)

# Methodology donut
ax = axes[1, 1]
wedges3, _, autotexts3 = ax.pie([32.3, 35.4, 32.3], labels=None, autopct='%1.1f%%', startangle=90,
    colors=['#3A7CBD', '#1B4F8A', '#E8A838'], pctdistance=0.78, wedgeprops=dict(width=0.45, edgecolor='white', linewidth=2))
for t in autotexts3: t.set_fontsize(9); t.set_color('white'); t.set_fontweight('bold')
ax.legend(['Agile', 'Waterfall', 'Hybrid'], loc='lower center', ncol=3, fontsize=7,
          bbox_to_anchor=(0.5, -0.08), frameon=False, labelcolor='white')
ax.set_title('Primary Methodology', color='white', fontsize=11, fontweight='bold', pad=8)

fig.subplots_adjust(hspace=0.35, wspace=0.3)
chart_demo = save_fig(fig, 'demographics_combo')

# ── Descriptive stats ──────────────────────────────────────
fig, ax = plt.subplots(figsize=(7, 3.5))
fig.patch.set_alpha(0); ax.set_facecolor('none')
cats = ['Risk Process\nAlignment', 'Risk Visibility\nIntegration', 'Timely Risk\nEscalation',
        'Governance\nConsistency', 'AI-Driven Risk\nAnalytics', 'Risk Mgmt\nEffectiveness']
vals = [3.6073, 3.5500, 3.6646, 3.6962, 3.6703, 3.4573]
cols = ['#1B4F8A', '#3A7CBD', '#5BA0D9', '#1B4F8A', '#27AE60', '#E8A838']
x = np.arange(len(cats))
brs = ax.bar(x, vals, color=cols, width=0.55, edgecolor='white', linewidth=0.5)
for b, v in zip(brs, vals): ax.text(b.get_x()+b.get_width()/2, b.get_height()+0.02, f'{v:.2f}', ha='center', color='white', fontsize=9, fontweight='bold')
ax.set_xticks(x); ax.set_xticklabels(cats, color='white', fontsize=7.5)
ax.set_ylabel('Mean (5-point Likert)', color='white', fontsize=9)
ax.set_title('Mean Scores of Study Variables', color='white', fontsize=12, fontweight='bold', pad=12)
ax.tick_params(colors='white'); ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
ax.spines['bottom'].set_color('#555'); ax.spines['left'].set_color('#555'); ax.set_ylim(0, 4.2)
chart_desc = save_fig(fig, 'descriptive_stats')

# ── Conceptual Framework ──────────────────────────────────
fig, ax = plt.subplots(figsize=(10, 5.5))
fig.patch.set_alpha(0); ax.set_facecolor('none')
ax.set_xlim(0, 10); ax.set_ylim(0, 6); ax.axis('off')
iv_labels = ['Risk Process\nAlignment', 'Risk Visibility\nIntegration', 'Timely Risk\nEscalation',
             'Governance\nConsistency', 'AI-Driven Risk\nAnalytics Capability']
iv_colors = ['#1B4F8A', '#3A7CBD', '#5BA0D9', '#1B4F8A', '#27AE60']
betas = ['\u03B2=0.464***', '\u03B2=0.121*', '\u03B2=0.129*', '\u03B2=0.397***', '\u03B2=0.449***']
h_labels = ['H1', 'H2', 'H3', 'H4', 'H5']
y_pos = [5.0, 4.0, 3.0, 2.0, 1.0]

for i, (label, y, col, h) in enumerate(zip(iv_labels, y_pos, iv_colors, h_labels)):
    box = mpatches.FancyBboxPatch((0.2, y - 0.35), 2.6, 0.7, boxstyle="round,pad=0.1",
                                   facecolor=col, edgecolor='white', linewidth=1.5)
    ax.add_patch(box)
    ax.text(1.5, y, label, ha='center', va='center', color='white', fontsize=9, fontweight='bold')
    # H label
    ax.text(0.05, y, h, ha='center', va='center', color='#E8A838', fontsize=10, fontweight='bold')

dv_box = mpatches.FancyBboxPatch((6.5, 2.15), 3.2, 1.7, boxstyle="round,pad=0.1",
                                  facecolor='#E8A838', edgecolor='white', linewidth=2)
ax.add_patch(dv_box)
ax.text(8.1, 3.0, 'Risk Management\nEffectiveness\nin Hybrid Projects', ha='center', va='center',
        color='#0B1D3A', fontsize=11, fontweight='bold')

for i, (y, beta) in enumerate(zip(y_pos, betas)):
    ac = '#27AE60' if '***' in beta else '#F39C12'
    ax.annotate('', xy=(6.5, 3.0), xytext=(2.8, y), arrowprops=dict(arrowstyle='->', color=ac, lw=2))
    mx, my = 4.65, (y + 3.0) / 2
    ax.text(mx, my + 0.12, beta, ha='center', va='center', color='white', fontsize=8.5, fontweight='bold',
            bbox=dict(boxstyle='round,pad=0.2', facecolor='#0B1D3A', edgecolor=ac, alpha=0.9))

ax.text(5.0, 0.15, '\u2588  Strong significance (p<0.001)      \u2588  Moderate significance (p<0.05)',
        color='white', fontsize=8)
chart_fw = save_fig(fig, 'framework')

# ── Hypothesis comparison bar ─────────────────────────────
fig, ax = plt.subplots(figsize=(8, 4))
fig.patch.set_alpha(0); ax.set_facecolor('none')
hyps = ['H1: Risk Process\nAlignment', 'H2: Risk Visibility\nIntegration', 'H3: Timely Risk\nEscalation',
        'H4: Governance\nConsistency', 'H5: AI-Driven Risk\nAnalytics']
betas_v = [0.464, 0.121, 0.129, 0.397, 0.449]
p_vals = ['p<0.001', 'p=0.034', 'p=0.022', 'p<0.001', 'p<0.001']
colors_h = ['#3A7CBD', '#F39C12', '#F39C12', '#3A7CBD', '#27AE60']
x = np.arange(len(hyps))
brs = ax.bar(x, betas_v, color=colors_h, width=0.55, edgecolor='white', linewidth=1)
for b, v, p in zip(brs, betas_v, p_vals):
    ax.text(b.get_x()+b.get_width()/2, b.get_height()+0.015,
            f'\u03B2={v:.3f}\n{p}', ha='center', color='white', fontsize=9, fontweight='bold')
ax.set_xticks(x); ax.set_xticklabels(hyps, color='white', fontsize=8)
ax.set_ylabel('Standardized Path Coefficient (\u03B2)', color='white', fontsize=10)
ax.set_title('SEM Path Coefficients \u2013 All Hypotheses', color='white', fontsize=12, fontweight='bold', pad=15)
ax.tick_params(colors='white'); ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
ax.spines['bottom'].set_color('#555'); ax.spines['left'].set_color('#555'); ax.set_ylim(0, 0.6)
chart_hyp = save_fig(fig, 'hypothesis_comparison')

# ── Model Fit comparison ──────────────────────────────────
fig, ax = plt.subplots(figsize=(7, 3.5))
fig.patch.set_alpha(0); ax.set_facecolor('none')
groups = ['H1', 'H2', 'H3', 'H4', 'H5']
cfi = [0.976, 0.994, 0.995, 0.985, 0.982]
rmsea = [0.060, 0.031, 0.029, 0.043, 0.051]
x = np.arange(len(groups)); w = 0.3
b1 = ax.bar(x - w/2, cfi, w, label='CFI', color='#3A7CBD', edgecolor='white', linewidth=0.5)
b2 = ax.bar(x + w/2, rmsea, w, label='RMSEA', color='#E8A838', edgecolor='white', linewidth=0.5)
for b, v in zip(b1, cfi): ax.text(b.get_x()+b.get_width()/2, b.get_height()+0.01, f'{v:.3f}', ha='center', color='white', fontsize=8, fontweight='bold')
for b, v in zip(b2, rmsea): ax.text(b.get_x()+b.get_width()/2, b.get_height()+0.01, f'{v:.3f}', ha='center', color='white', fontsize=8, fontweight='bold')
ax.axhline(y=0.90, color='#27AE60', linewidth=1, linestyle='--', alpha=0.7, label='CFI Threshold (0.90)')
ax.axhline(y=0.08, color='#E74C3C', linewidth=1, linestyle='--', alpha=0.7, label='RMSEA Threshold (0.08)')
ax.set_xticks(x); ax.set_xticklabels(groups, color='white', fontsize=10)
ax.set_title('Model Fit Indices Across Hypotheses', color='white', fontsize=12, fontweight='bold', pad=12)
ax.legend(loc='upper right', fontsize=7, framealpha=0.3, labelcolor='white')
ax.tick_params(colors='white'); ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
ax.spines['bottom'].set_color('#555'); ax.spines['left'].set_color('#555'); ax.set_ylim(0, 1.12)
chart_fit = save_fig(fig, 'model_fit')

# ── AVE / CR grouped bar ─────────────────────────────────
fig, ax = plt.subplots(figsize=(7, 3.5))
fig.patch.set_alpha(0); ax.set_facecolor('none')
constructs = ['RPA', 'RVI', 'TRE', 'GC', 'ADRAC', 'RMEHP']
ave = [0.761, 0.811, 0.824, 0.757, 0.787, 0.743]
cr = [0.858, 0.873, 0.877, 0.880, 0.866, 0.852]
x = np.arange(len(constructs)); w = 0.3
b1 = ax.bar(x - w/2, ave, w, label='AVE', color='#3A7CBD', edgecolor='white', linewidth=0.5)
b2 = ax.bar(x + w/2, cr, w, label='CR', color='#E8A838', edgecolor='white', linewidth=0.5)
for b, v in zip(b1, ave): ax.text(b.get_x()+b.get_width()/2, b.get_height()+0.008, f'{v:.3f}', ha='center', color='white', fontsize=7.5, fontweight='bold')
for b, v in zip(b2, cr): ax.text(b.get_x()+b.get_width()/2, b.get_height()+0.008, f'{v:.3f}', ha='center', color='white', fontsize=7.5, fontweight='bold')
ax.axhline(y=0.5, color='#27AE60', linewidth=1, linestyle='--', alpha=0.6, label='AVE Threshold (0.50)')
ax.axhline(y=0.7, color='#16A085', linewidth=1, linestyle='--', alpha=0.6, label='CR Threshold (0.70)')
ax.set_xticks(x); ax.set_xticklabels(constructs, color='white', fontsize=9)
ax.set_title('Convergent Validity: AVE & Composite Reliability', color='white', fontsize=11, fontweight='bold', pad=12)
ax.legend(loc='lower right', fontsize=7, framealpha=0.3, labelcolor='white')
ax.tick_params(colors='white'); ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
ax.spines['bottom'].set_color('#555'); ax.spines['left'].set_color('#555'); ax.set_ylim(0, 1.05)
chart_val = save_fig(fig, 'validity')

print("All charts generated.")


# ═══════════════════════════════════════════════════════════
# BUILD SLIDES
# ═══════════════════════════════════════════════════════════
print("Building slides...")

# ─── SLIDE 1: TITLE ───────────────────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s)
bar(s, 0, 0, 0.12, 7.5, GOLD)  # left accent

tb(s, 1.0, 1.0, 11, 0.8, "Enhancing Risk Management Effectiveness in", 34, WHITE, True)
tb(s, 1.0, 1.75, 11, 0.8, "Hybrid Projects through AI-Driven", 34, WHITE, True)
tb(s, 1.0, 2.5, 11, 0.8, "Strategic Governance", 34, GOLD, True)
tb(s, 1.0, 3.6, 8, 0.5, "A Structural Equation Modelling Study", 18, SKY_BLUE)
bar(s, 1.0, 4.3, 4, 0.03, GOLD)
tb(s, 1.0, 4.6, 8, 0.35, "Based on primary data from 384 project professionals", 13, MUTED)
tb(s, 1.0, 4.95, 8, 0.35, "Analysis: Structural Equation Modelling (SEM) | SPSS & AMOS", 13, MUTED)
gold_bar(s)


# ─── SLIDE 2: INTRODUCTION ───────────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s); section(s, "Introduction & Research Context"); gold_bar(s); slide_num(s, 2)

# Three problem cards
cards = [
    ("The Problem", "Traditional risk management relies on static registers, periodic assessments, and subjective expert judgment \u2014 insufficient for today's complex, data-intensive projects.", MID_BLUE),
    ("Hybrid Complexity", "Agile + Waterfall coexistence creates fragmented governance, inconsistent risk documentation, and ambiguous escalation authority.", DARK_BLUE),
    ("The AI Opportunity", "AI enables predictive analytics, real-time monitoring, continuous learning, and data-driven decision support for risk governance.", TEAL),
]
for i, (title, desc, col) in enumerate(cards):
    left = 0.8 + i * 4.0
    rrect(s, left, 2.2, 3.7, 4.5, col)
    tb(s, left + 0.25, 2.4, 3.2, 0.5, title, 16, GOLD, True, PP_ALIGN.CENTER)
    bar(s, left + 0.5, 3.1, 2.7, 0.02, GOLD)
    tb(s, left + 0.25, 3.3, 3.2, 3.2, desc, 12, SOFT_TEXT, ls=1.4)

# Research gap callout
rrect(s, 0.8, 6.85, 11.7, 0.45, DARK_BLUE,
      "Research Gap: No integrated SEM model testing AI capability alongside governance mechanisms in hybrid project risk management",
      11, GOLD, True)


# ─── SLIDE 3: OBJECTIVES & HYPOTHESES ────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s); section(s, "Research Objectives & Hypotheses"); gold_bar(s); slide_num(s, 3)

items = [
    ("H1", "Risk Process Alignment \u2192 Risk Management Effectiveness", MID_BLUE),
    ("H2", "Risk Visibility Integration \u2192 Risk Management Effectiveness", LIGHT_BLUE),
    ("H3", "Timely Risk Escalation \u2192 Risk Management Effectiveness", SKY_BLUE),
    ("H4", "Governance Consistency \u2192 Risk Management Effectiveness", MID_BLUE),
    ("H5", "AI-Driven Risk Analytics Capability \u2192 Risk Management Effectiveness", TEAL),
]
for i, (h, desc, col) in enumerate(items):
    top = 2.2 + i * 0.95
    rrect(s, 0.8, top, 0.7, 0.7, col, h, 16, WHITE, True)
    tb(s, 1.7, top + 0.15, 10, 0.5, desc, 14, WHITE)

tb(s, 0.8, 7.0, 11, 0.35, "All hypotheses posit a significant positive impact on Risk Management Effectiveness in Hybrid Project Environments.",
   11, MUTED, False, PP_ALIGN.LEFT)


# ─── SLIDE 4: CONCEPTUAL FRAMEWORK ──────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s); section(s, "Conceptual Framework", "Integrated Model of AI-Driven Strategic Risk Management"); gold_bar(s); slide_num(s, 4)
img(s, chart_fw, 1.5, 1.7, 10.5, 5.5)


# ─── SLIDE 5: METHODOLOGY ───────────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s); section(s, "Research Methodology"); gold_bar(s); slide_num(s, 5)

# Method cards
mcards = [
    ("\U0001F4CA", "Quantitative\nResearch Design", "Cross-sectional survey\nwith 5-point Likert scale", MID_BLUE),
    ("\U0001F465", "384 Professionals\nSurveyed", "Project managers, risk managers,\nPMO, analysts, scrum masters", TEAL),
    ("\U0001F527", "SPSS + AMOS\nAnalysis Tools", "CFA for measurement model\nSEM for hypothesis testing", DARK_BLUE),
    ("\u2705", "Rigorous\nValidation", "KMO = 0.927 | All AVE > 0.5\nAll CR > 0.7 | Fornell-Larcker", MID_BLUE),
]
for i, (icon, title, desc, col) in enumerate(mcards):
    left = 0.6 + i * 3.15
    rrect(s, left, 2.2, 2.9, 2.8, col)
    tb(s, left + 0.2, 2.35, 2.5, 0.5, title, 14, GOLD, True, PP_ALIGN.CENTER)
    bar(s, left + 0.4, 3.15, 2.1, 0.02, GOLD)
    tb(s, left + 0.2, 3.3, 2.5, 1.2, desc, 11, SOFT_TEXT, False, PP_ALIGN.CENTER, 1.4)

# Techniques list
tb(s, 0.8, 5.3, 12, 0.4, "Statistical Techniques", 15, GOLD, True)
ml(s, 0.8, 5.8, 11, 1.5, [
    "Confirmatory Factor Analysis (CFA) for measurement model validation",
    "Structural Equation Modelling (SEM) for testing causal relationships",
    "Discriminant Validity via Fornell-Larcker Criterion | Cronbach's Alpha for reliability",
], 12, SOFT_TEXT, bullet=True, ls=1.35)


# ─── SLIDE 6: DEMOGRAPHICS ──────────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s); section(s, "Sample Demographics", "N = 384 Project Professionals across Diverse Industries"); gold_bar(s); slide_num(s, 6)
img(s, chart_demo, 0.3, 1.9, 8.5, 5.2)

# Key stat badges
stats = [
    ("77%+", "Over 5 years\nexperience", MID_BLUE),
    ("53.9%", "Female\nrespondents", TEAL),
    ("27.6%", "Hold professional\ncertifications", DARK_BLUE),
    ("32.3%", "Hybrid project\nexperience", MID_BLUE),
]
for i, (num, label, col) in enumerate(stats):
    top = 2.1 + i * 1.25
    rrect(s, 9.3, top, 3.6, 1.05, col)
    tb(s, 9.5, top + 0.05, 1.1, 0.5, num, 20, GOLD, True, PP_ALIGN.CENTER)
    tb(s, 10.6, top + 0.1, 2.2, 0.7, label, 11, WHITE)


# ─── SLIDE 7: MEASUREMENT MODEL & VALIDITY ──────────────
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s); section(s, "Measurement Model & Validity"); gold_bar(s); slide_num(s, 7)

img(s, chart_val, 0.3, 1.9, 7.0, 4.0)
img(s, chart_desc, 0.3, 4.6, 7.0, 2.7)

# Validity summary
rrect(s, 7.8, 2.0, 5.0, 5.2, DARK_BLUE)
tb(s, 8.0, 2.2, 4.6, 0.4, "Measurement Validation Summary", 14, GOLD, True, PP_ALIGN.CENTER)

ml(s, 8.1, 2.8, 4.4, 4.2, [
    ("KMO = 0.927", True, GREEN_OK),
    "Excellent sampling adequacy",
    "",
    ("All AVE > 0.50", True, GREEN_OK),
    "Convergent validity confirmed",
    "",
    ("All CR > 0.70", True, GREEN_OK),
    "Internal consistency established",
    "",
    ("Fornell-Larcker Criterion Met", True, GREEN_OK),
    "Discriminant validity verified",
    "",
    ("Bartlett's Test: p < 0.001", True, GREEN_OK),
    "\u03C7\u00B2 = 7618.445, df = 465",
    "",
    ("All factor loadings > 0.60", True, GREEN_OK),
    "Indicator reliability confirmed",
], 10.5, SOFT_TEXT, ls=1.2)


# ─── SLIDE 8: H1 & H2 RESULTS ───────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s); section(s, "Hypothesis Results: H1 & H2"); gold_bar(s); slide_num(s, 8)

# H1 card
rrect(s, 0.8, 2.1, 5.8, 5.0, DARK_BLUE)
rrect(s, 0.8, 2.1, 5.8, 0.55, MID_BLUE, "H1: Risk Process Alignment \u2192 RME", 13, WHITE, True)
ml(s, 1.1, 2.85, 5.2, 4.0, [
    ("\u2705  SUPPORTED", True, GREEN_OK),
    "",
    ("\u03B2 = 0.464  |  C.R. = 7.346  |  p < 0.001", True, WHITE),
    "",
    ("Model Fit:", True, GOLD),
    "CMIN/DF = 2.366 | GFI = 0.958 | CFI = 0.976",
    "RMSEA = 0.060",
    "",
    ("Interpretation:", True, GOLD),
    "Moderate-to-strong positive effect. Aligning risk identification, assessment, and monitoring across Agile & Waterfall components significantly enhances coordination and reduces ambiguity.",
    "",
    "Process alignment functions as a foundational enabler of effective risk management.",
], 11, SOFT_TEXT, ls=1.2)

# H2 card
rrect(s, 6.9, 2.1, 5.8, 5.0, DARK_BLUE)
rrect(s, 6.9, 2.1, 5.8, 0.55, LIGHT_BLUE, "H2: Risk Visibility Integration \u2192 RME", 13, WHITE, True)
ml(s, 7.2, 2.85, 5.2, 4.0, [
    ("\u26A0\uFE0F  SUPPORTED (Marginal Effect)", True, ORANGE),
    "",
    ("\u03B2 = 0.121  |  C.R. = 2.116  |  p = 0.034", True, WHITE),
    "",
    ("Model Fit:", True, GOLD),
    "CMIN/DF = 1.376 | CFI = 0.994",
    "RMSEA = 0.031",
    "",
    ("Interpretation:", True, GOLD),
    "Statistically significant but modest effect. Visibility alone is insufficient \u2014 it functions as a supportive mechanism rather than an independent driver.",
    "",
    "Without strong governance and interpretive capacity, information availability does not translate into better outcomes.",
], 11, SOFT_TEXT, ls=1.2)


# ─── SLIDE 9: H3 & H4 RESULTS ───────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s); section(s, "Hypothesis Results: H3 & H4"); gold_bar(s); slide_num(s, 9)

# H3 card
rrect(s, 0.8, 2.1, 5.8, 5.0, DARK_BLUE)
rrect(s, 0.8, 2.1, 5.8, 0.55, SKY_BLUE, "H3: Timely Risk Escalation \u2192 RME", 13, NAVY, True)
ml(s, 1.1, 2.85, 5.2, 4.0, [
    ("\u26A0\uFE0F  SUPPORTED (Modest Effect)", True, ORANGE),
    "",
    ("\u03B2 = 0.129  |  C.R. = 2.283  |  p = 0.022", True, WHITE),
    "",
    ("Model Fit:", True, GOLD),
    "CMIN/DF = 1.326 | CFI = 0.995",
    "RMSEA = 0.029",
    "",
    ("Interpretation:", True, GOLD),
    "Positive but modest direct effect. In hybrid settings, Agile teams manage risks locally, reducing the need for frequent formal escalation.",
    "",
    "Escalation effectiveness depends on appropriateness and quality rather than on speed alone.",
], 11, SOFT_TEXT, ls=1.2)

# H4 card
rrect(s, 6.9, 2.1, 5.8, 5.0, DARK_BLUE)
rrect(s, 6.9, 2.1, 5.8, 0.55, MID_BLUE, "H4: Governance Consistency \u2192 RME", 13, WHITE, True)
ml(s, 7.2, 2.85, 5.2, 4.0, [
    ("\u2705  SUPPORTED", True, GREEN_OK),
    "",
    ("\u03B2 = 0.397  |  C.R. = 6.869  |  p < 0.001", True, WHITE),
    "",
    ("Model Fit:", True, GOLD),
    "CMIN/DF = 1.716 | CFI = 0.985",
    "RMSEA = 0.043",
    "",
    ("Interpretation:", True, GOLD),
    "Governance Consistency is the structural backbone that stabilizes risk management across diverse execution approaches.",
    "",
    "Uniform application of policies, decision rights, and accountability mechanisms reduces ambiguity and enhances coordination.",
], 11, SOFT_TEXT, ls=1.2)


# ─── SLIDE 10: H5 RESULT (AI - STAR FINDING) ────────────
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s); section(s, "Hypothesis 5 \u2013 Key Finding",
                   "AI-Driven Risk Analytics Capability \u2192 Risk Management Effectiveness"); gold_bar(s); slide_num(s, 10)

# Central highlight
rrect(s, 0.8, 2.1, 7.5, 5.0, DARK_BLUE)

# Big beta
tb(s, 1.2, 2.3, 6.5, 0.6, "\u2705  H5 STRONGLY SUPPORTED", 22, GREEN_OK, True, PP_ALIGN.CENTER)
bar(s, 1.5, 3.05, 6.0, 0.02, GREEN_OK)

tb(s, 1.2, 3.2, 6.5, 0.6, "\u03B2 = 0.449   |   C.R. = 7.697   |   p < 0.001", 18, WHITE, True, PP_ALIGN.CENTER)

ml(s, 1.2, 4.0, 6.5, 3.0, [
    ("Model Fit:", True, GOLD),
    "CMIN/DF = 1.995  |  NFI = 0.965  |  CFI = 0.982  |  RMSEA = 0.051",
    "",
    ("Key Insight:", True, GOLD),
    "AI-Driven Risk Analytics is one of the STRONGEST predictors of Risk Management Effectiveness among all variables tested.",
    "",
    "AI functions as an organizational capability that enhances predictive foresight, anomaly detection, real-time monitoring, and automated risk prioritization.",
    "",
    "It complements and amplifies governance structures rather than replacing them.",
], 12, SOFT_TEXT, ls=1.25)

# Right side summary badges
rrect(s, 8.8, 2.3, 3.9, 1.3, TEAL,
      "Strongest Predictor\nAmong All 5 Variables", 14, WHITE, True)
rrect(s, 8.8, 3.9, 3.9, 1.3, MID_BLUE,
      "Strategic Capability\nNot Just a Tech Tool", 14, WHITE, True)
rrect(s, 8.8, 5.5, 3.9, 1.3, DARK_BLUE,
      "Amplifies Governance\n& Decision Quality", 14, WHITE, True)


# ─── SLIDE 11: HYPOTHESIS SUMMARY ───────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s); section(s, "Summary of All Hypothesis Results"); gold_bar(s); slide_num(s, 11)

img(s, chart_hyp, 0.2, 1.8, 8.0, 4.8)

# Result table
rrect(s, 8.5, 2.0, 4.4, 4.6, DARK_BLUE)
tb(s, 8.7, 2.15, 4.0, 0.35, "Outcome Summary", 14, GOLD, True, PP_ALIGN.CENTER)

# Header row
tb(s, 8.7, 2.6, 0.5, 0.3, "H", 10, GOLD, True)
tb(s, 9.2, 2.6, 1.8, 0.3, "Path", 10, GOLD, True)
tb(s, 11.0, 2.6, 0.7, 0.3, "\u03B2", 10, GOLD, True, PP_ALIGN.CENTER)
tb(s, 11.7, 2.6, 1.1, 0.3, "Result", 10, GOLD, True, PP_ALIGN.CENTER)

rows = [
    ("H1", "RPA \u2192 RME", "0.464", "\u2705 Supported", GREEN_OK),
    ("H2", "RVI \u2192 RME", "0.121", "\u26A0 Marginal", ORANGE),
    ("H3", "TRE \u2192 RME", "0.129", "\u26A0 Modest", ORANGE),
    ("H4", "GC \u2192 RME", "0.397", "\u2705 Supported", GREEN_OK),
    ("H5", "ADRAC \u2192 RME", "0.449", "\u2705 Supported", GREEN_OK),
]
for i, (h, path, beta, result, rc) in enumerate(rows):
    t = 3.05 + i * 0.58
    tb(s, 8.7, t, 0.5, 0.3, h, 10, WHITE, True)
    tb(s, 9.2, t, 1.8, 0.3, path, 9, SOFT_TEXT)
    tb(s, 11.0, t, 0.7, 0.3, beta, 10, WHITE, True, PP_ALIGN.CENTER)
    tb(s, 11.7, t, 1.1, 0.3, result, 9, rc, True, PP_ALIGN.CENTER)

# Ranking
rrect(s, 8.5, 6.1, 4.4, 0.45, MID_BLUE,
      "Rank: H5 (AI) > H1 (RPA) > H4 (GC) >> H3 > H2", 10, GOLD, True)

# Model fit below chart
img(s, chart_fit, 0.2, 4.6, 5.0, 2.5)
tb(s, 5.3, 5.0, 3.0, 1.5, "All 5 models\nexceed fit\nthresholds\n\n\u2713 CFI > 0.90\n\u2713 RMSEA < 0.08\n\u2713 CMIN/DF < 3.0",
   10, GREEN_OK, True, PP_ALIGN.LEFT)


# ─── SLIDE 12: DISCUSSION ───────────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s); section(s, "Discussion & Key Insights"); gold_bar(s); slide_num(s, 12)

findings = [
    ("Structural Coherence\nDrives Effectiveness",
     "Governance Consistency and Risk Process Alignment are foundational enablers. Risk management effectiveness requires deliberate integration across Agile and Waterfall governance \u2014 not parallel or fragmented systems.",
     MID_BLUE),
    ("AI as a Strategic\nOrganizational Capability",
     "AI analytics is not merely a technical tool but an organizational capability that enhances predictive foresight and decision quality. It amplifies governance structures and reduces cognitive limitations in complex environments.",
     TEAL),
    ("Visibility & Speed\nAre Necessary but Insufficient",
     "Risk visibility and timely escalation are supportive but not dominant drivers. Without governance clarity and interpretive capacity, information availability alone does not improve risk outcomes. Quality of decisions matters more than speed.",
     DARK_BLUE),
]
for i, (title, desc, col) in enumerate(findings):
    left = 0.6 + i * 4.1
    rrect(s, left, 2.1, 3.8, 5.0, col)
    tb(s, left + 0.2, 2.3, 3.4, 0.7, title, 14, GOLD, True, PP_ALIGN.CENTER)
    bar(s, left + 0.4, 3.2, 3.0, 0.02, GOLD)
    tb(s, left + 0.2, 3.4, 3.4, 3.5, desc, 11, SOFT_TEXT, ls=1.4)


# ─── SLIDE 13: IMPLICATIONS ─────────────────────────────
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s); section(s, "Theoretical & Practical Implications"); gold_bar(s); slide_num(s, 13)

# Theoretical column
rrect(s, 0.8, 2.1, 5.5, 5.0, DARK_BLUE)
tb(s, 1.0, 2.25, 5.0, 0.4, "Theoretical Contributions", 16, GOLD, True, PP_ALIGN.CENTER)
bar(s, 1.5, 2.8, 4.0, 0.02, GOLD)
ml(s, 1.1, 2.95, 4.9, 3.8, [
    "Positions AI as a structural determinant of risk management effectiveness, not just a peripheral tool",
    "",
    "Advances hybrid project management research through integrated SEM-based empirical validation",
    "",
    "Bridges governance theory and dynamic capability theory in the hybrid project domain",
    "",
    "Demonstrates AI capability as complementary to governance structures, not a replacement for them",
], 11, SOFT_TEXT, bullet=True, ls=1.25)

# Practical column
rrect(s, 6.9, 2.1, 5.5, 5.0, MID_BLUE)
tb(s, 7.1, 2.25, 5.0, 0.4, "Practical Recommendations", 16, GOLD, True, PP_ALIGN.CENTER)
bar(s, 7.5, 2.8, 4.0, 0.02, GOLD)
ml(s, 7.2, 2.95, 4.9, 3.8, [
    "Prioritize harmonizing risk processes across Agile and Waterfall components",
    "",
    "Establish consistent governance: clear decision rights, accountability, and escalation pathways",
    "",
    "Invest in AI analytics embedded within governance frameworks \u2014 not deployed in isolation",
    "",
    "Train managers to interpret and act on AI-generated insights for strategic decisions",
], 11, RGBColor(0xDD, 0xEE, 0xFF), bullet=True, ls=1.25)


# ─── SLIDE 14: CONCLUSION + LIMITATIONS + THANK YOU ─────
s = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(s); section(s, "Conclusion, Limitations & Future Research"); gold_bar(s); slide_num(s, 14)

# Central conclusion
rrect(s, 0.8, 2.0, 11.7, 1.2, MID_BLUE)
tb(s, 1.0, 2.1, 11.3, 0.9,
   "Effective risk management in hybrid projects requires integrated structural alignment,\n"
   "governance stability, and AI-driven analytical intelligence \u2014 not isolated practices or technologies.",
   14, WHITE, True, PP_ALIGN.CENTER)

# Three conclusion pillars
pillars = [
    ("\u2460", "Governance + Process\nAlignment", "Structural coherence is the foundation for effective risk governance."),
    ("\u2461", "AI as Strategic\nEnabler", "Strongest predictor \u2014 enhances foresight and amplifies governance."),
    ("\u2462", "Beyond Visibility\n& Speed", "Decision quality and governance clarity matter more than information volume."),
]
for i, (num, title, desc) in enumerate(pillars):
    left = 0.8 + i * 4.0
    circ(s, left + 1.15, 3.45, 0.55, GOLD, num, 16, NAVY)
    tb(s, left, 4.15, 3.7, 0.5, title, 12, GOLD, True, PP_ALIGN.CENTER)
    tb(s, left, 4.65, 3.7, 0.6, desc, 10, SOFT_TEXT, False, PP_ALIGN.CENTER)

# Limitations & Future (side by side, compact)
rrect(s, 0.8, 5.45, 5.5, 1.5, DARK_BLUE)
tb(s, 1.0, 5.5, 3.0, 0.3, "Limitations", 12, GOLD, True)
ml(s, 1.0, 5.85, 5.0, 1.0, [
    "Cross-sectional design limits causal inference",
    "India-focused sample \u2014 may limit generalizability",
    "Self-reported data \u2014 potential response bias",
], 9.5, SOFT_TEXT, bullet=True, ls=1.2)

rrect(s, 6.9, 5.45, 5.5, 1.5, DARK_BLUE)
tb(s, 7.1, 5.5, 3.5, 0.3, "Future Research", 12, GOLD, True)
ml(s, 7.1, 5.85, 5.0, 1.0, [
    "Longitudinal studies across project phases",
    "Cross-cultural replication in other regions",
    "Mediating roles of culture & project complexity",
], 9.5, SOFT_TEXT, bullet=True, ls=1.2)

# Thank you
rrect(s, 3.5, 7.0, 6.3, 0.35, GOLD, "Thank You  \u2014  Questions & Discussion Welcome", 13, NAVY, True)


# ═══════════════════════════════════════════════════════════
# SAVE
# ═══════════════════════════════════════════════════════════
output = os.path.join(OUT_DIR, "AI_Risk_Management_Presentation_14slides.pptx")
prs.save(output)
print(f"\nPresentation saved: {output}")
print(f"Total slides: {TOTAL}")
