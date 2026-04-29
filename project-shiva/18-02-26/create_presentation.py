#!/usr/bin/env python3
"""
Creates a professional PowerPoint presentation for the thesis:
"Enhancing Risk Management Effectiveness in Hybrid Projects
through AI-Driven Strategic Governance: A SEM Study"
"""

import os
import io
import math
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_LABEL_POSITION
from pptx.chart.data import CategoryChartData
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

# ── Color Palette ──────────────────────────────────────────
NAVY      = RGBColor(0x0B, 0x1D, 0x3A)   # Deep navy
DARK_BLUE = RGBColor(0x13, 0x2E, 0x5B)   # Section headers
MID_BLUE  = RGBColor(0x1B, 0x4F, 0x8A)   # Accents
LIGHT_BLUE= RGBColor(0x3A, 0x7C, 0xBD)   # Charts
SKY_BLUE  = RGBColor(0x5B, 0xA0, 0xD9)   # Lighter accent
GOLD      = RGBColor(0xE8, 0xA8, 0x38)   # Gold accent
AMBER     = RGBColor(0xF0, 0xC0, 0x4A)   # Lighter gold
WHITE     = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY= RGBColor(0xF2, 0xF4, 0xF7)
DARK_GRAY = RGBColor(0x4A, 0x4A, 0x4A)
MED_GRAY  = RGBColor(0x6B, 0x6B, 0x6B)
GREEN_OK  = RGBColor(0x27, 0xAE, 0x60)
RED_NO    = RGBColor(0xE7, 0x4C, 0x3C)
ORANGE    = RGBColor(0xF3, 0x9C, 0x12)
TEAL      = RGBColor(0x16, 0xA0, 0x85)

OUT_DIR = r"D:\project-shiva\18-02-26"
CHART_DIR = os.path.join(OUT_DIR, "charts")
os.makedirs(CHART_DIR, exist_ok=True)

prs = Presentation()
prs.slide_width  = Inches(13.333)
prs.slide_height = Inches(7.5)

# ── Helper functions ───────────────────────────────────────

def add_bg(slide, color=NAVY):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color

def add_gradient_bg(slide, c1=NAVY, c2=DARK_BLUE):
    """Simple solid bg (gradient needs XML hacking, using solid for reliability)"""
    add_bg(slide, c1)

def add_bottom_bar(slide, color=GOLD, height=0.08):
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(7.5 - height),
        Inches(13.333), Inches(height)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = color
    bar.line.fill.background()

def add_top_accent(slide, color=GOLD, width=3.5, height=0.06):
    bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0.8), Inches(0.7),
        Inches(width), Inches(height)
    )
    bar.fill.solid()
    bar.fill.fore_color.rgb = color
    bar.line.fill.background()

def add_slide_number(slide, num, total):
    txBox = slide.shapes.add_textbox(
        Inches(12.3), Inches(7.05), Inches(0.8), Inches(0.35)
    )
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = f"{num}/{total}"
    p.font.size = Pt(10)
    p.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
    p.alignment = PP_ALIGN.RIGHT

def add_textbox(slide, left, top, width, height, text, font_size=16,
                color=WHITE, bold=False, alignment=PP_ALIGN.LEFT,
                font_name="Calibri", line_spacing=1.2):
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment
    p.line_spacing = Pt(font_size * line_spacing)
    return txBox

def add_multiline_textbox(slide, left, top, width, height, lines,
                          font_size=14, color=WHITE, font_name="Calibri",
                          line_spacing=1.3, bullet=False, alignment=PP_ALIGN.LEFT):
    """lines: list of (text, bold, color_override) tuples or just strings"""
    txBox = slide.shapes.add_textbox(
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, line in enumerate(lines):
        if isinstance(line, str):
            txt, bld, clr = line, False, color
        else:
            txt = line[0]
            bld = line[1] if len(line) > 1 else False
            clr = line[2] if len(line) > 2 else color

        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        prefix = "\u2022  " if bullet else ""
        p.text = prefix + txt
        p.font.size = Pt(font_size)
        p.font.color.rgb = clr
        p.font.bold = bld
        p.font.name = font_name
        p.alignment = alignment
        p.line_spacing = Pt(font_size * line_spacing)
        p.space_after = Pt(4)
    return txBox

def add_rounded_rect(slide, left, top, width, height, fill_color, text="",
                     font_size=12, font_color=WHITE, bold=False, alignment=PP_ALIGN.CENTER):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    shape.shadow.inherit = False
    if text:
        tf = shape.text_frame
        tf.word_wrap = True
        tf.paragraphs[0].alignment = alignment
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(font_size)
        p.font.color.rgb = font_color
        p.font.bold = bold
        p.font.name = "Calibri"
    return shape

def add_circle_icon(slide, left, top, size, fill_color, text="", font_size=20, font_color=WHITE):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.OVAL,
        Inches(left), Inches(top), Inches(size), Inches(size)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()
    if text:
        tf = shape.text_frame
        tf.word_wrap = True
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = Pt(font_size)
        p.font.color.rgb = font_color
        p.font.bold = True
        p.font.name = "Calibri"
        tf.paragraphs[0].space_before = Pt(0)
    return shape

def add_section_header(slide, section_title, subtitle=""):
    add_top_accent(slide, GOLD, 2.5)
    add_textbox(slide, 0.8, 0.85, 10, 0.7, section_title, 32, GOLD, True)
    if subtitle:
        add_textbox(slide, 0.8, 1.5, 10, 0.5, subtitle, 14, RGBColor(0xAA, 0xBB, 0xCC))

def save_chart_image(fig, name, dpi=200):
    path = os.path.join(CHART_DIR, f"{name}.png")
    fig.savefig(path, dpi=dpi, bbox_inches='tight', transparent=True, pad_inches=0.1)
    plt.close(fig)
    return path

def add_image_centered(slide, img_path, left, top, width=None, height=None):
    if width and height:
        slide.shapes.add_picture(img_path, Inches(left), Inches(top), Inches(width), Inches(height))
    elif width:
        slide.shapes.add_picture(img_path, Inches(left), Inches(top), Inches(width))
    else:
        slide.shapes.add_picture(img_path, Inches(left), Inches(top))


# ── Chart generation functions ─────────────────────────────

def make_demographics_pie(data, labels, title, colors, name):
    fig, ax = plt.subplots(figsize=(5, 4))
    fig.patch.set_alpha(0)
    wedges, texts, autotexts = ax.pie(
        data, labels=None, autopct='%1.1f%%', startangle=90,
        colors=colors, pctdistance=0.78,
        wedgeprops=dict(width=0.45, edgecolor='white', linewidth=2)
    )
    for t in autotexts:
        t.set_fontsize(9)
        t.set_color('white')
        t.set_fontweight('bold')
    ax.legend(labels, loc='lower center', ncol=2, fontsize=8,
              bbox_to_anchor=(0.5, -0.12), frameon=False,
              labelcolor='white')
    ax.set_title(title, color='white', fontsize=12, fontweight='bold', pad=10)
    return save_chart_image(fig, name)

def make_bar_chart(categories, values, title, name, color_list=None, ylabel="", horizontal=False):
    fig, ax = plt.subplots(figsize=(6, 3.8))
    fig.patch.set_alpha(0)
    ax.set_facecolor('none')

    if color_list is None:
        color_list = ['#1B4F8A', '#3A7CBD', '#5BA0D9', '#E8A838', '#F0C04A', '#16A085']

    x = np.arange(len(categories))
    if horizontal:
        bars = ax.barh(x, values, color=color_list[:len(values)], height=0.55, edgecolor='white', linewidth=0.5)
        ax.set_yticks(x)
        ax.set_yticklabels(categories, color='white', fontsize=9)
        ax.set_xlabel(ylabel, color='white', fontsize=10)
        for bar, val in zip(bars, values):
            ax.text(bar.get_width() + 0.02, bar.get_y() + bar.get_height()/2,
                    f'{val:.3f}' if val < 10 else f'{val}',
                    va='center', ha='left', color='white', fontsize=9, fontweight='bold')
        ax.invert_yaxis()
    else:
        bars = ax.bar(x, values, color=color_list[:len(values)], width=0.55, edgecolor='white', linewidth=0.5)
        ax.set_xticks(x)
        ax.set_xticklabels(categories, color='white', fontsize=8, rotation=15, ha='right')
        ax.set_ylabel(ylabel, color='white', fontsize=10)
        for bar, val in zip(bars, values):
            ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.02,
                    f'{val:.4f}' if val < 1 else (f'{val:.2f}' if val < 10 else f'{val}'),
                    ha='center', va='bottom', color='white', fontsize=9, fontweight='bold')

    ax.set_title(title, color='white', fontsize=11, fontweight='bold', pad=12)
    ax.tick_params(colors='white')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_color('#555555')
    ax.spines['left'].set_color('#555555')
    return save_chart_image(fig, name)

def make_hypothesis_chart(name):
    """Bar chart comparing path coefficients for all 5 hypotheses"""
    fig, ax = plt.subplots(figsize=(7, 4))
    fig.patch.set_alpha(0)
    ax.set_facecolor('none')

    hypotheses = ['H1: Risk\nProcess\nAlignment', 'H2: Risk\nVisibility\nIntegration',
                  'H3: Timely\nRisk\nEscalation', 'H4: Governance\nConsistency',
                  'H5: AI-Driven\nRisk Analytics\nCapability']
    betas = [0.464, 0.121, 0.129, 0.397, 0.449]
    p_vals = ['<0.001', '0.034', '0.022', '<0.001', '<0.001']
    colors_h = ['#3A7CBD', '#F39C12', '#F39C12', '#3A7CBD', '#27AE60']

    x = np.arange(len(hypotheses))
    bars = ax.bar(x, betas, color=colors_h, width=0.55, edgecolor='white', linewidth=1)

    for bar, val, p in zip(bars, betas, p_vals):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.012,
                f'\u03B2 = {val:.3f}\np = {p}',
                ha='center', va='bottom', color='white', fontsize=9, fontweight='bold')

    # Significance line
    ax.axhline(y=0, color='#555555', linewidth=0.5)

    ax.set_xticks(x)
    ax.set_xticklabels(hypotheses, color='white', fontsize=8)
    ax.set_ylabel('Standardized Path Coefficient (\u03B2)', color='white', fontsize=10)
    ax.set_title('SEM Path Coefficients \u2013 All Hypotheses', color='white', fontsize=12, fontweight='bold', pad=15)
    ax.tick_params(colors='white')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_color('#555555')
    ax.spines['left'].set_color('#555555')
    ax.set_ylim(0, 0.58)
    return save_chart_image(fig, name)

def make_model_fit_chart(name):
    """Grouped bar chart comparing model fit indices across hypotheses"""
    fig, ax = plt.subplots(figsize=(7, 4))
    fig.patch.set_alpha(0)
    ax.set_facecolor('none')

    groups = ['H1', 'H2', 'H3', 'H4', 'H5']
    cfi_vals = [0.976, 0.994, 0.995, 0.985, 0.982]
    rmsea_vals = [0.060, 0.031, 0.029, 0.043, 0.051]

    x = np.arange(len(groups))
    w = 0.3
    b1 = ax.bar(x - w/2, cfi_vals, w, label='CFI', color='#3A7CBD', edgecolor='white', linewidth=0.5)
    b2 = ax.bar(x + w/2, rmsea_vals, w, label='RMSEA', color='#E8A838', edgecolor='white', linewidth=0.5)

    for bar, val in zip(b1, cfi_vals):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.01,
                f'{val:.3f}', ha='center', va='bottom', color='white', fontsize=8, fontweight='bold')
    for bar, val in zip(b2, rmsea_vals):
        ax.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 0.01,
                f'{val:.3f}', ha='center', va='bottom', color='white', fontsize=8, fontweight='bold')

    ax.axhline(y=0.90, color='#27AE60', linewidth=1, linestyle='--', alpha=0.7, label='CFI Threshold (0.90)')
    ax.axhline(y=0.08, color='#E74C3C', linewidth=1, linestyle='--', alpha=0.7, label='RMSEA Threshold (0.08)')

    ax.set_xticks(x)
    ax.set_xticklabels(groups, color='white', fontsize=10)
    ax.set_title('Model Fit Indices Across Hypotheses', color='white', fontsize=12, fontweight='bold', pad=15)
    ax.legend(loc='upper right', fontsize=8, framealpha=0.3, labelcolor='white')
    ax.tick_params(colors='white')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['bottom'].set_color('#555555')
    ax.spines['left'].set_color('#555555')
    ax.set_ylim(0, 1.12)
    return save_chart_image(fig, name)

def make_conceptual_framework(name):
    """Create the conceptual framework diagram"""
    fig, ax = plt.subplots(figsize=(9, 5))
    fig.patch.set_alpha(0)
    ax.set_facecolor('none')
    ax.set_xlim(0, 10)
    ax.set_ylim(0, 6)
    ax.axis('off')

    iv_labels = [
        'Risk Process\nAlignment',
        'Risk Visibility\nIntegration',
        'Timely Risk\nEscalation',
        'Governance\nConsistency',
        'AI-Driven Risk\nAnalytics Capability'
    ]
    iv_colors = ['#1B4F8A', '#3A7CBD', '#5BA0D9', '#1B4F8A', '#27AE60']
    betas = ['\u03B2=0.464***', '\u03B2=0.121*', '\u03B2=0.129*', '\u03B2=0.397***', '\u03B2=0.449***']
    results = ['Supported', 'Marginal', 'Modest', 'Supported', 'Supported']

    y_positions = [5.0, 4.0, 3.0, 2.0, 1.0]

    for i, (label, y, col) in enumerate(zip(iv_labels, y_positions, iv_colors)):
        box = mpatches.FancyBboxPatch((0.3, y - 0.35), 2.8, 0.7,
                                       boxstyle="round,pad=0.1",
                                       facecolor=col, edgecolor='white', linewidth=1.5)
        ax.add_patch(box)
        ax.text(1.7, y, label, ha='center', va='center', color='white',
                fontsize=9, fontweight='bold')

    # DV box
    dv_box = mpatches.FancyBboxPatch((6.5, 2.15), 3.2, 1.7,
                                      boxstyle="round,pad=0.1",
                                      facecolor='#E8A838', edgecolor='white', linewidth=2)
    ax.add_patch(dv_box)
    ax.text(8.1, 3.0, 'Risk Management\nEffectiveness\nin Hybrid Projects', ha='center', va='center',
            color='#0B1D3A', fontsize=11, fontweight='bold')

    # Arrows
    for i, (y, beta) in enumerate(zip(y_positions, betas)):
        arrow_color = '#27AE60' if '***' in beta else '#F39C12'
        ax.annotate('', xy=(6.5, 3.0), xytext=(3.1, y),
                    arrowprops=dict(arrowstyle='->', color=arrow_color, lw=2))
        mid_x = 4.8
        mid_y = (y + 3.0) / 2
        ax.text(mid_x, mid_y + 0.1, beta, ha='center', va='center',
                color='white', fontsize=8, fontweight='bold',
                bbox=dict(boxstyle='round,pad=0.2', facecolor='#0B1D3A', edgecolor=arrow_color, alpha=0.9))

    # Legend
    ax.text(0.3, 0.2, '\u2588 Strong significance (p<0.001)     \u2588 Modest significance (p<0.05)',
            color='white', fontsize=8)

    return save_chart_image(fig, name)

def make_single_hypothesis_chart(h_num, iv_name, beta, cr, p_val, supported, color, name):
    """Create a focused chart for a single hypothesis"""
    fig, ax = plt.subplots(figsize=(5.5, 3.5))
    fig.patch.set_alpha(0)
    ax.set_facecolor('none')
    ax.axis('off')

    # IV box
    iv_box = mpatches.FancyBboxPatch((0.05, 0.35), 0.35, 0.3,
                                      boxstyle="round,pad=0.02",
                                      facecolor=color, edgecolor='white', linewidth=2,
                                      transform=ax.transAxes)
    ax.add_patch(iv_box)
    ax.text(0.225, 0.5, iv_name, ha='center', va='center', color='white',
            fontsize=11, fontweight='bold', transform=ax.transAxes)

    # DV box
    dv_box = mpatches.FancyBboxPatch((0.6, 0.35), 0.35, 0.3,
                                      boxstyle="round,pad=0.02",
                                      facecolor='#E8A838', edgecolor='white', linewidth=2,
                                      transform=ax.transAxes)
    ax.add_patch(dv_box)
    ax.text(0.775, 0.5, 'Risk Management\nEffectiveness', ha='center', va='center',
            color='#0B1D3A', fontsize=10, fontweight='bold', transform=ax.transAxes)

    # Arrow
    ax.annotate('', xy=(0.6, 0.5), xytext=(0.4, 0.5),
                arrowprops=dict(arrowstyle='->', color='white', lw=3),
                transform=ax.transAxes)

    # Beta label on arrow
    ax.text(0.5, 0.56, f'\u03B2 = {beta}', ha='center', va='bottom',
            color='white', fontsize=13, fontweight='bold', transform=ax.transAxes)

    # Stats below
    stats_text = f'C.R. = {cr}  |  p = {p_val}'
    ax.text(0.5, 0.18, stats_text, ha='center', va='center',
            color='white', fontsize=11, transform=ax.transAxes)

    result_color = '#27AE60' if 'Supported' in supported else '#F39C12'
    ax.text(0.5, 0.08, supported, ha='center', va='center',
            color=result_color, fontsize=14, fontweight='bold', transform=ax.transAxes,
            bbox=dict(boxstyle='round,pad=0.3', facecolor='#0B1D3A', edgecolor=result_color, linewidth=2))

    ax.set_title(f'Hypothesis {h_num}', color='white', fontsize=14, fontweight='bold', pad=20)
    return save_chart_image(fig, name)


# ═══════════════════════════════════════════════════════════
# GENERATE ALL CHARTS
# ═══════════════════════════════════════════════════════════
print("Generating charts...")

# 1. Gender pie
chart_gender = make_demographics_pie(
    [46.1, 53.9], ['Male (46.1%)', 'Female (53.9%)'],
    'Gender Distribution', ['#1B4F8A', '#E8A838'], 'gender_pie'
)

# 2. Age bar
chart_age = make_bar_chart(
    ['20-25', '26-35', '36-45', '46-55', '56+'],
    [17.7, 22.7, 19.5, 15.6, 24.5],
    'Age Distribution (%)', 'age_bar', ylabel='Percentage (%)'
)

# 3. Experience bar
chart_exp = make_bar_chart(
    ['< 5 yrs', '5-10 yrs', '11-15 yrs', '> 15 yrs'],
    [22.4, 30.2, 23.2, 24.2],
    'Project Management Experience (%)', 'exp_bar', ylabel='Percentage (%)'
)

# 4. AI Usage pie
chart_ai = make_demographics_pie(
    [35.4, 32.0, 32.6], ['Not Used (35.4%)', 'Limited Use (32.0%)', 'Moderate Use (32.6%)'],
    'Extent of AI Usage in Risk Management', ['#E74C3C', '#F39C12', '#27AE60'], 'ai_usage_pie'
)

# 5. Descriptive stats
chart_desc = make_bar_chart(
    ['RPA', 'RVI', 'TRE', 'GC', 'ADRAC', 'RMEHP'],
    [3.6073, 3.5500, 3.6646, 3.6962, 3.6703, 3.4573],
    'Mean Scores of Study Variables', 'descriptive_stats',
    color_list=['#1B4F8A', '#3A7CBD', '#5BA0D9', '#1B4F8A', '#27AE60', '#E8A838'],
    ylabel='Mean (5-point Likert Scale)'
)

# 6. Convergent validity
chart_validity = make_bar_chart(
    ['RPA', 'RVI', 'TRE', 'GC', 'ADRAC', 'RMEHP'],
    [0.761, 0.811, 0.824, 0.757, 0.787, 0.743],
    'Average Variance Extracted (AVE) by Construct', 'ave_chart',
    color_list=['#1B4F8A', '#3A7CBD', '#5BA0D9', '#1B4F8A', '#27AE60', '#E8A838'],
    ylabel='AVE'
)

# 7. Conceptual framework
chart_framework = make_conceptual_framework('conceptual_framework')

# 8. Hypothesis comparison
chart_hyp_compare = make_hypothesis_chart('hypothesis_comparison')

# 9. Model fit
chart_model_fit = make_model_fit_chart('model_fit')

# 10. Individual hypothesis charts
chart_h1 = make_single_hypothesis_chart(1, 'Risk Process\nAlignment', '0.464', '7.346', '<0.001',
                                         'H1 Supported', '#1B4F8A', 'h1_result')
chart_h2 = make_single_hypothesis_chart(2, 'Risk Visibility\nIntegration', '0.121', '2.116', '0.034',
                                         'H2 Supported (Marginal)', '#3A7CBD', 'h2_result')
chart_h3 = make_single_hypothesis_chart(3, 'Timely Risk\nEscalation', '0.129', '2.283', '0.022',
                                         'H3 Supported (Modest)', '#5BA0D9', 'h3_result')
chart_h4 = make_single_hypothesis_chart(4, 'Governance\nConsistency', '0.397', '6.869', '<0.001',
                                         'H4 Supported', '#1B4F8A', 'h4_result')
chart_h5 = make_single_hypothesis_chart(5, 'AI-Driven Risk\nAnalytics', '0.449', '7.697', '<0.001',
                                         'H5 Supported', '#27AE60', 'h5_result')

# 11. Methodology pie
chart_method = make_demographics_pie(
    [32.3, 35.4, 32.3], ['Agile (32.3%)', 'Waterfall (35.4%)', 'Hybrid (32.3%)'],
    'Primary Project Methodology', ['#3A7CBD', '#1B4F8A', '#E8A838'], 'methodology_pie'
)

# 12. Role bar
chart_role = make_bar_chart(
    ['Project\nManager', 'Program/\nPortfolio Mgr', 'Risk\nManager', 'Business\nAnalyst', 'Team Lead/\nScrum Master'],
    [22.4, 19.0, 18.2, 21.4, 19.0],
    'Respondent Role Distribution (%)', 'role_bar', ylabel='Percentage (%)'
)

print("All charts generated.")

# ═══════════════════════════════════════════════════════════
# BUILD SLIDES
# ═══════════════════════════════════════════════════════════
TOTAL_SLIDES = 22

# ─── SLIDE 1: Title ───────────────────────────────────────
print("Building slides...")
slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
add_bg(slide, NAVY)

# Left decorative accent bar
add_rounded_rect(slide, 0, 0, 0.12, 7.5, GOLD)

# Title block
add_textbox(slide, 1.0, 1.2, 11, 1.0,
            "Enhancing Risk Management Effectiveness in",
            36, WHITE, True, PP_ALIGN.LEFT, "Calibri Light")
add_textbox(slide, 1.0, 2.0, 11, 1.0,
            "Hybrid Projects through AI-Driven",
            36, WHITE, True, PP_ALIGN.LEFT, "Calibri Light")
add_textbox(slide, 1.0, 2.8, 11, 1.0,
            "Strategic Governance",
            36, GOLD, True, PP_ALIGN.LEFT, "Calibri Light")

# Subtitle
add_textbox(slide, 1.0, 4.0, 8, 0.5,
            "A Structural Equation Modelling Study",
            20, SKY_BLUE, False, PP_ALIGN.LEFT, "Calibri")

# Separator line
sep = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
    Inches(1.0), Inches(4.7), Inches(4), Inches(0.03))
sep.fill.solid()
sep.fill.fore_color.rgb = GOLD
sep.line.fill.background()

# Meta info
add_textbox(slide, 1.0, 5.0, 6, 0.4,
            "Based on primary data from 384 project professionals",
            14, RGBColor(0x99, 0xAA, 0xBB))
add_textbox(slide, 1.0, 5.4, 6, 0.4,
            "Methodology: Structural Equation Modelling (SEM) | SPSS & AMOS",
            14, RGBColor(0x99, 0xAA, 0xBB))

add_bottom_bar(slide)

# ─── SLIDE 2: Agenda ─────────────────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Presentation Outline")
add_bottom_bar(slide)

agenda_items = [
    ("01", "Introduction & Background"),
    ("02", "Research Objectives & Questions"),
    ("03", "Conceptual Framework & Hypotheses"),
    ("04", "Methodology"),
    ("05", "Demographics & Sample Profile"),
    ("06", "Measurement Model & Validity"),
    ("07", "SEM Results (H1 \u2013 H5)"),
    ("08", "Discussion & Implications"),
    ("09", "Conclusion & Future Research"),
]

for i, (num, title) in enumerate(agenda_items):
    row = i // 3
    col = i % 3
    left = 0.8 + col * 4.0
    top = 2.3 + row * 1.6

    add_circle_icon(slide, left, top, 0.5, GOLD if i < 6 else MID_BLUE, num, 14, NAVY)
    add_textbox(slide, left + 0.65, top + 0.05, 3.2, 0.5, title, 13, WHITE, False)

add_slide_number(slide, 2, TOTAL_SLIDES)

# ─── SLIDE 3: Introduction ───────────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Introduction & Background")
add_bottom_bar(slide)

# Left column
add_multiline_textbox(slide, 0.8, 2.2, 5.5, 4.5, [
    ("The Problem", True, GOLD),
    "Traditional risk management relies on static registers,",
    "periodic assessments, and subjective expert judgment.",
    "",
    ("Hybrid Project Complexity", True, GOLD),
    "Agile + Waterfall integration creates fragmented governance,",
    "inconsistent risk documentation, and ambiguous escalation.",
    "",
    ("The AI Opportunity", True, GOLD),
    "AI enables predictive analytics, real-time monitoring,",
    "and data-driven decision support for risk governance.",
], 13, RGBColor(0xCC, 0xDD, 0xEE), line_spacing=1.35)

# Right column - key stats
add_rounded_rect(slide, 7.2, 2.3, 5.3, 1.2, DARK_BLUE,
                 "Most IT projects fail to meet original objectives\n(Standish Group, 2020)",
                 13, WHITE)
add_rounded_rect(slide, 7.2, 3.8, 5.3, 1.2, MID_BLUE,
                 "Traditional methods struggle with dynamic,\ndata-intensive environments",
                 13, WHITE)
add_rounded_rect(slide, 7.2, 5.3, 5.3, 1.2, TEAL,
                 "AI: Machine Learning + NLP + Deep Learning\nenable continuous risk sensing",
                 13, WHITE)

add_slide_number(slide, 3, TOTAL_SLIDES)

# ─── SLIDE 4: Research Objectives ────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Research Objectives & Questions")
add_bottom_bar(slide)

objectives = [
    ("RO1", "Examine the impact of Risk Process Alignment on\nRisk Management Effectiveness"),
    ("RO2", "Analyze the influence of Risk Visibility Integration\non Risk Management Effectiveness"),
    ("RO3", "Assess the effect of Timely Risk Escalation on\nRisk Management Effectiveness"),
    ("RO4", "Evaluate the role of Governance Consistency in\nenhancing Risk Management Effectiveness"),
    ("RO5", "Examine the impact of AI-Driven Risk Analytics\nCapability on Risk Management Effectiveness"),
]

for i, (num, text) in enumerate(objectives):
    top = 2.2 + i * 1.0
    add_circle_icon(slide, 0.8, top, 0.55, GOLD if i % 2 == 0 else MID_BLUE, num, 11, NAVY if i % 2 == 0 else WHITE)
    add_textbox(slide, 1.55, top + 0.07, 11, 0.6, text, 13, WHITE)

add_slide_number(slide, 4, TOTAL_SLIDES)

# ─── SLIDE 5: Conceptual Framework ──────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Conceptual Framework", "Integrated Model of AI-Driven Strategic Risk Management")
add_bottom_bar(slide)
add_image_centered(slide, chart_framework, 2.0, 1.9, 9.5, 5.2)
add_slide_number(slide, 5, TOTAL_SLIDES)

# ─── SLIDE 6: Hypotheses ────────────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Research Hypotheses")
add_bottom_bar(slide)

hyps = [
    ("H1", "Risk Process Alignment has a significant positive impact on\nRisk Management Effectiveness in Hybrid Project Environments."),
    ("H2", "Risk Visibility Integration has a significant positive impact on\nRisk Management Effectiveness in Hybrid Project Environments."),
    ("H3", "Timely Risk Escalation has a significant positive impact on\nRisk Management Effectiveness in Hybrid Project Environments."),
    ("H4", "Governance Consistency has a significant positive impact on\nRisk Management Effectiveness in Hybrid Project Environments."),
    ("H5", "AI-Driven Risk Analytics Capability has a significant positive\nimpact on Risk Management Effectiveness in Hybrid Project Environments."),
]

for i, (hnum, htext) in enumerate(hyps):
    top = 2.2 + i * 0.95
    colors_h = [MID_BLUE, LIGHT_BLUE, SKY_BLUE, MID_BLUE, TEAL]
    add_rounded_rect(slide, 0.8, top, 0.65, 0.7, colors_h[i], hnum, 15, WHITE, True)
    add_textbox(slide, 1.65, top + 0.08, 10.5, 0.6, htext, 12, RGBColor(0xCC, 0xDD, 0xEE))

add_slide_number(slide, 6, TOTAL_SLIDES)

# ─── SLIDE 7: Methodology ───────────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Research Methodology")
add_bottom_bar(slide)

# Method cards
method_data = [
    ("Research Design", "Quantitative,\ncross-sectional\nsurvey design", MID_BLUE),
    ("Sample Size", "384 project\nprofessionals\nacross industries", TEAL),
    ("Data Collection", "Structured\nquestionnaire\n(5-point Likert)", DARK_BLUE),
    ("Analysis Tools", "SPSS + AMOS\nfor CFA and\nSEM analysis", MID_BLUE),
]

for i, (title, desc, col) in enumerate(method_data):
    left = 0.8 + i * 3.1
    add_rounded_rect(slide, left, 2.3, 2.7, 1.8, col)
    add_textbox(slide, left + 0.2, 2.45, 2.3, 0.4, title, 14, GOLD, True, PP_ALIGN.CENTER)
    add_textbox(slide, left + 0.2, 2.95, 2.3, 1.0, desc, 12, WHITE, False, PP_ALIGN.CENTER)

# Analysis techniques
add_textbox(slide, 0.8, 4.5, 12, 0.4, "Statistical Techniques Applied", 16, GOLD, True)

techniques = [
    "Confirmatory Factor Analysis (CFA) for measurement model validation",
    "Structural Equation Modelling (SEM) for hypothesis testing",
    "Cronbach's Alpha, AVE, and Composite Reliability for construct validity",
    "KMO and Bartlett's Test for sampling adequacy (KMO = 0.927)",
    "Discriminant Validity via Fornell-Larcker Criterion",
]
add_multiline_textbox(slide, 0.8, 5.0, 11, 2.0, techniques, 12,
                      RGBColor(0xCC, 0xDD, 0xEE), bullet=True, line_spacing=1.4)
add_slide_number(slide, 7, TOTAL_SLIDES)

# ─── SLIDE 8: Demographics 1 ────────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Sample Demographics", "N = 384 Project Professionals")
add_bottom_bar(slide)
add_image_centered(slide, chart_gender, 0.3, 2.0, 4.2, 3.4)
add_image_centered(slide, chart_age, 4.5, 2.0, 4.5, 3.2)
add_image_centered(slide, chart_exp, 0.3, 4.7, 4.5, 2.5)
add_image_centered(slide, chart_role, 4.8, 4.7, 4.5, 2.5)

# Key stat boxes on the right
add_rounded_rect(slide, 9.5, 2.3, 3.3, 1.0, MID_BLUE,
                 "77%+ have over\n5 years experience", 12, WHITE, True)
add_rounded_rect(slide, 9.5, 3.6, 3.3, 1.0, TEAL,
                 "53.9% Female\n46.1% Male", 12, WHITE, True)
add_rounded_rect(slide, 9.5, 4.9, 3.3, 1.0, DARK_BLUE,
                 "27.6% Hold\nProfessional Certifications", 12, WHITE, True)
add_rounded_rect(slide, 9.5, 6.2, 3.3, 0.8, MID_BLUE,
                 "Balanced role distribution", 12, WHITE, True)

add_slide_number(slide, 8, TOTAL_SLIDES)

# ─── SLIDE 9: Demographics 2 ────────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Methodology & AI Adoption Profile")
add_bottom_bar(slide)
add_image_centered(slide, chart_method, 0.5, 2.0, 5.5, 4.5)
add_image_centered(slide, chart_ai, 6.5, 2.0, 5.5, 4.5)

add_textbox(slide, 0.8, 6.5, 11, 0.5,
            "Even distribution across methodologies ensures balanced insights | AI adoption is progressing but not yet universal",
            12, RGBColor(0x99, 0xAA, 0xBB), False, PP_ALIGN.CENTER)
add_slide_number(slide, 9, TOTAL_SLIDES)

# ─── SLIDE 10: Descriptive Stats ────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Descriptive Statistics", "Mean Scores on 5-point Likert Scale")
add_bottom_bar(slide)
add_image_centered(slide, chart_desc, 0.5, 2.0, 7, 4.5)

# Interpretation box
add_rounded_rect(slide, 8.0, 2.3, 4.8, 4.5, DARK_BLUE)
add_multiline_textbox(slide, 8.3, 2.5, 4.2, 4.0, [
    ("Key Observations", True, GOLD),
    "",
    "Governance Consistency and AI-Driven Risk Analytics score highest (3.70, 3.67)",
    "",
    "Risk Management Effectiveness has the lowest mean (3.46), suggesting uneven translation of practices into outcomes",
    "",
    "All skewness values within \u00b12 and kurtosis within \u00b13 \u2192 data is suitable for SEM",
    "",
    ("All constructs show moderately high agreement", True, RGBColor(0x88, 0xDD, 0x88)),
], 11, RGBColor(0xBB, 0xCC, 0xDD), line_spacing=1.3)

add_slide_number(slide, 10, TOTAL_SLIDES)

# ─── SLIDE 11: Measurement Model ────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Measurement Model & Validity")
add_bottom_bar(slide)

add_image_centered(slide, chart_validity, 0.3, 2.0, 6.5, 4.0)

# Validity metrics table
add_rounded_rect(slide, 7.3, 2.0, 5.5, 5.0, DARK_BLUE)
add_textbox(slide, 7.5, 2.2, 5, 0.4, "Convergent Validity", 15, GOLD, True, PP_ALIGN.CENTER)

validity_rows = [
    ("Construct", "AVE", "CR"),
    ("Risk Process Alignment", "0.761", "0.858"),
    ("Risk Visibility Integration", "0.811", "0.873"),
    ("Timely Risk Escalation", "0.824", "0.877"),
    ("Governance Consistency", "0.757", "0.880"),
    ("AI-Driven Risk Analytics", "0.787", "0.866"),
    ("Risk Mgmt Effectiveness", "0.743", "0.852"),
]

for i, (c, a, cr) in enumerate(validity_rows):
    top_v = 2.7 + i * 0.45
    clr = GOLD if i == 0 else WHITE
    bld = i == 0
    fsz = 10 if i == 0 else 10
    add_textbox(slide, 7.5, top_v, 2.7, 0.35, c, fsz, clr, bld)
    add_textbox(slide, 10.2, top_v, 0.8, 0.35, a, fsz, clr, bld, PP_ALIGN.CENTER)
    add_textbox(slide, 11.0, top_v, 0.8, 0.35, cr, fsz, clr, bld, PP_ALIGN.CENTER)

# KMO stat
add_rounded_rect(slide, 7.5, 6.0, 5.0, 0.7, MID_BLUE,
                 "KMO = 0.927 (Excellent Sampling Adequacy)\nBartlett's Test: \u03C7\u00B2 = 7618.445, p < 0.001",
                 11, WHITE, True)

add_textbox(slide, 0.3, 6.2, 6.5, 0.5,
            "All AVE > 0.5 and CR > 0.7 confirm convergent validity",
            12, RGBColor(0x88, 0xDD, 0x88), True, PP_ALIGN.CENTER)

add_slide_number(slide, 11, TOTAL_SLIDES)

# ─── SLIDE 12: H1 Result ────────────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Hypothesis 1 \u2013 Result",
                   "Risk Process Alignment \u2192 Risk Management Effectiveness")
add_bottom_bar(slide)
add_image_centered(slide, chart_h1, 0.5, 2.0, 6.0, 3.8)

add_rounded_rect(slide, 7.0, 2.2, 5.8, 4.8, DARK_BLUE)
add_multiline_textbox(slide, 7.3, 2.4, 5.2, 4.4, [
    ("Result: H1 Supported", True, GREEN_OK),
    "",
    ("\u03B2 = 0.464 | C.R. = 7.346 | p < 0.001", True, WHITE),
    "",
    ("Model Fit:", True, GOLD),
    "CMIN/DF = 2.366 | GFI = 0.958",
    "CFI = 0.976 | RMSEA = 0.060",
    "",
    ("Interpretation:", True, GOLD),
    "Risk Process Alignment shows a moderate and meaningful positive relationship with RME.",
    "",
    "When risk processes are aligned across Agile & Waterfall, organizations experience improved coordination and reduced ambiguity.",
], 11, RGBColor(0xCC, 0xDD, 0xEE), line_spacing=1.25)
add_slide_number(slide, 12, TOTAL_SLIDES)

# ─── SLIDE 13: H2 Result ────────────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Hypothesis 2 \u2013 Result",
                   "Risk Visibility Integration \u2192 Risk Management Effectiveness")
add_bottom_bar(slide)
add_image_centered(slide, chart_h2, 0.5, 2.0, 6.0, 3.8)

add_rounded_rect(slide, 7.0, 2.2, 5.8, 4.8, DARK_BLUE)
add_multiline_textbox(slide, 7.3, 2.4, 5.2, 4.4, [
    ("Result: H2 Supported (Marginal Effect)", True, ORANGE),
    "",
    ("\u03B2 = 0.121 | C.R. = 2.116 | p = 0.034", True, WHITE),
    "",
    ("Model Fit:", True, GOLD),
    "CMIN/DF = 1.376 | CFI = 0.994",
    "RMSEA = 0.031",
    "",
    ("Interpretation:", True, GOLD),
    "While statistically significant, the effect size is modest.",
    "",
    "Risk visibility alone is insufficient \u2014 it functions as a supportive mechanism rather than a dominant driver when governance structures are weak.",
], 11, RGBColor(0xCC, 0xDD, 0xEE), line_spacing=1.25)
add_slide_number(slide, 13, TOTAL_SLIDES)

# ─── SLIDE 14: H3 Result ────────────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Hypothesis 3 \u2013 Result",
                   "Timely Risk Escalation \u2192 Risk Management Effectiveness")
add_bottom_bar(slide)
add_image_centered(slide, chart_h3, 0.5, 2.0, 6.0, 3.8)

add_rounded_rect(slide, 7.0, 2.2, 5.8, 4.8, DARK_BLUE)
add_multiline_textbox(slide, 7.3, 2.4, 5.2, 4.4, [
    ("Result: H3 Supported (Modest Effect)", True, ORANGE),
    "",
    ("\u03B2 = 0.129 | C.R. = 2.283 | p = 0.022", True, WHITE),
    "",
    ("Model Fit:", True, GOLD),
    "CMIN/DF = 1.326 | CFI = 0.995",
    "RMSEA = 0.029",
    "",
    ("Interpretation:", True, GOLD),
    "Statistically significant but modest direct effect.",
    "",
    "Agile teams manage risks locally, reducing the need for frequent formal escalation. Effectiveness depends on appropriateness, not speed.",
], 11, RGBColor(0xCC, 0xDD, 0xEE), line_spacing=1.25)
add_slide_number(slide, 14, TOTAL_SLIDES)

# ─── SLIDE 15: H4 Result ────────────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Hypothesis 4 \u2013 Result",
                   "Governance Consistency \u2192 Risk Management Effectiveness")
add_bottom_bar(slide)
add_image_centered(slide, chart_h4, 0.5, 2.0, 6.0, 3.8)

add_rounded_rect(slide, 7.0, 2.2, 5.8, 4.8, DARK_BLUE)
add_multiline_textbox(slide, 7.3, 2.4, 5.2, 4.4, [
    ("Result: H4 Supported", True, GREEN_OK),
    "",
    ("\u03B2 = 0.397 | C.R. = 6.869 | p < 0.001", True, WHITE),
    "",
    ("Model Fit:", True, GOLD),
    "CMIN/DF = 1.716 | CFI = 0.985",
    "RMSEA = 0.043",
    "",
    ("Interpretation:", True, GOLD),
    "Governance Consistency is a structural backbone that stabilizes risk management across diverse execution approaches.",
    "",
    "Uniform governance reduces ambiguity in authority, accountability, and decision-making in hybrid settings.",
], 11, RGBColor(0xCC, 0xDD, 0xEE), line_spacing=1.25)
add_slide_number(slide, 15, TOTAL_SLIDES)

# ─── SLIDE 16: H5 Result ────────────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Hypothesis 5 \u2013 Result",
                   "AI-Driven Risk Analytics Capability \u2192 Risk Management Effectiveness")
add_bottom_bar(slide)
add_image_centered(slide, chart_h5, 0.5, 2.0, 6.0, 3.8)

add_rounded_rect(slide, 7.0, 2.2, 5.8, 4.8, DARK_BLUE)
add_multiline_textbox(slide, 7.3, 2.4, 5.2, 4.4, [
    ("Result: H5 Supported", True, GREEN_OK),
    "",
    ("\u03B2 = 0.449 | C.R. = 7.697 | p < 0.001", True, WHITE),
    "",
    ("Model Fit:", True, GOLD),
    "CMIN/DF = 1.995 | CFI = 0.982",
    "RMSEA = 0.051",
    "",
    ("Interpretation:", True, GOLD),
    "AI-Driven Risk Analytics is one of the STRONGEST predictors of RME.",
    "",
    "AI enhances predictive foresight, anomaly detection, real-time monitoring, and automated risk prioritization in hybrid environments.",
], 11, RGBColor(0xCC, 0xDD, 0xEE), line_spacing=1.25)
add_slide_number(slide, 16, TOTAL_SLIDES)

# ─── SLIDE 17: Hypothesis Summary ───────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Summary of Hypothesis Testing Results")
add_bottom_bar(slide)
add_image_centered(slide, chart_hyp_compare, 0.3, 2.0, 7.5, 4.5)

# Summary table
add_rounded_rect(slide, 8.2, 2.2, 4.6, 5.0, DARK_BLUE)
add_textbox(slide, 8.4, 2.4, 4.2, 0.4, "Hypothesis Outcomes", 14, GOLD, True, PP_ALIGN.CENTER)

hyp_summary = [
    ("H1", "RPA \u2192 RME", "\u03B2=0.464", "Supported", GREEN_OK),
    ("H2", "RVI \u2192 RME", "\u03B2=0.121", "Marginal", ORANGE),
    ("H3", "TRE \u2192 RME", "\u03B2=0.129", "Modest", ORANGE),
    ("H4", "GC \u2192 RME", "\u03B2=0.397", "Supported", GREEN_OK),
    ("H5", "ADRAC \u2192 RME", "\u03B2=0.449", "Supported", GREEN_OK),
]

for i, (h, path, beta, result, rcolor) in enumerate(hyp_summary):
    top_h = 3.0 + i * 0.75
    add_textbox(slide, 8.5, top_h, 0.5, 0.3, h, 11, WHITE, True)
    add_textbox(slide, 9.1, top_h, 1.6, 0.3, path, 10, RGBColor(0xBB, 0xCC, 0xDD))
    add_textbox(slide, 10.7, top_h, 0.8, 0.3, beta, 10, WHITE, True)
    add_textbox(slide, 11.5, top_h, 1.2, 0.3, result, 10, rcolor, True)

# Ranking
add_textbox(slide, 8.4, 6.5, 4.2, 0.4,
            "Strongest: H5 (AI) > H1 (RPA) > H4 (GC)", 10, GOLD, True, PP_ALIGN.CENTER)

add_slide_number(slide, 17, TOTAL_SLIDES)

# ─── SLIDE 18: Model Fit ────────────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Model Fit Indices", "All models demonstrate good-to-excellent fit")
add_bottom_bar(slide)
add_image_centered(slide, chart_model_fit, 0.5, 1.8, 8, 5.0)

# Fit criteria box
add_rounded_rect(slide, 8.8, 2.2, 4.0, 4.8, DARK_BLUE)
add_textbox(slide, 9.0, 2.4, 3.6, 0.4, "Acceptance Criteria", 13, GOLD, True, PP_ALIGN.CENTER)

fit_criteria = [
    "CMIN/DF < 3.00",
    "GFI \u2265 0.90",
    "NFI \u2265 0.90",
    "CFI \u2265 0.90",
    "IFI \u2265 0.90",
    "RMSEA \u2264 0.08",
    "RMR \u2264 0.05",
]
add_multiline_textbox(slide, 9.0, 3.0, 3.6, 3.5, [
    (f"\u2713  {c}", False, GREEN_OK) for c in fit_criteria
], 11, GREEN_OK, line_spacing=1.5)

add_textbox(slide, 9.0, 6.3, 3.6, 0.5,
            "All models meet or\nexceed all thresholds", 11, GREEN_OK, True, PP_ALIGN.CENTER)

add_slide_number(slide, 18, TOTAL_SLIDES)

# ─── SLIDE 19: Discussion ───────────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Discussion & Key Insights")
add_bottom_bar(slide)

# Three key findings
findings = [
    ("Structural Coherence Matters Most",
     "Governance Consistency and Risk Process Alignment are foundational. Effective risk management requires deliberate integration across Agile and Waterfall governance systems.",
     MID_BLUE, "\u26A1"),
    ("AI as Strategic Capability",
     "AI analytics is not merely a technical tool but an organizational capability that enhances predictive foresight and decision quality. It complements and amplifies governance structures.",
     TEAL, "\U0001F916"),
    ("Visibility & Speed Are Insufficient Alone",
     "Risk visibility and timely escalation are supportive but not dominant drivers. Without strong governance and interpretive capacity, information availability does not improve outcomes.",
     DARK_BLUE, "\U0001F50D"),
]

for i, (title, desc, col, icon) in enumerate(findings):
    left = 0.8 + i * 4.0
    add_rounded_rect(slide, left, 2.3, 3.7, 4.5, col)
    add_textbox(slide, left + 0.3, 2.5, 3.1, 0.6, title, 14, GOLD, True, PP_ALIGN.CENTER)
    sep = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
        Inches(left + 0.5), Inches(3.3), Inches(2.7), Inches(0.02))
    sep.fill.solid()
    sep.fill.fore_color.rgb = GOLD
    sep.line.fill.background()
    add_textbox(slide, left + 0.3, 3.5, 3.1, 3.0, desc, 11, RGBColor(0xDD, 0xEE, 0xFF), line_spacing=1.4)

add_slide_number(slide, 19, TOTAL_SLIDES)

# ─── SLIDE 20: Implications ─────────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Theoretical & Practical Implications")
add_bottom_bar(slide)

# Two columns
add_rounded_rect(slide, 0.8, 2.3, 5.5, 4.5, DARK_BLUE)
add_textbox(slide, 1.0, 2.5, 5.0, 0.4, "Theoretical Contributions", 16, GOLD, True, PP_ALIGN.CENTER)

theo_items = [
    "Extends project risk management theory by positioning AI as a structural determinant, not just a tool",
    "Advances hybrid project management research through SEM-based empirical validation",
    "Bridges governance theory and dynamic capability theory in hybrid project domain",
    "Demonstrates AI capability as complementary to governance, not a replacement",
]
add_multiline_textbox(slide, 1.1, 3.1, 4.9, 3.5, theo_items, 11,
                      RGBColor(0xCC, 0xDD, 0xEE), bullet=True, line_spacing=1.4)

add_rounded_rect(slide, 7.0, 2.3, 5.5, 4.5, MID_BLUE)
add_textbox(slide, 7.2, 2.5, 5.0, 0.4, "Practical Recommendations", 16, GOLD, True, PP_ALIGN.CENTER)

prac_items = [
    "Prioritize harmonizing risk processes across Agile and Waterfall components",
    "Establish consistent governance standards: accountability, decision rights, and escalation pathways",
    "Invest in AI analytics embedded within governance frameworks, not standalone",
    "Train managers to interpret and act on AI-generated insights for strategic decisions",
]
add_multiline_textbox(slide, 7.3, 3.1, 4.9, 3.5, prac_items, 11,
                      RGBColor(0xDD, 0xEE, 0xFF), bullet=True, line_spacing=1.4)

add_slide_number(slide, 20, TOTAL_SLIDES)

# ─── SLIDE 21: Conclusion ───────────────────────────────
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Conclusion")
add_bottom_bar(slide)

# Central takeaway
add_rounded_rect(slide, 1.5, 2.3, 10.3, 1.4, MID_BLUE)
add_textbox(slide, 1.8, 2.5, 9.7, 1.0,
            "Effective risk management in hybrid projects requires integrated structural alignment,\n"
            "governance stability, and AI-driven analytical intelligence \u2014 not isolated practices.",
            14, WHITE, True, PP_ALIGN.CENTER)

# Three key conclusions
conclusions = [
    ("Governance +\nProcess Alignment",
     "Structural coherence is the foundation \u2014 consistent governance and aligned risk processes drive effectiveness."),
    ("AI as Strategic\nEnabler",
     "AI capability is one of the strongest predictors \u2014 it enhances predictive foresight and amplifies governance."),
    ("Beyond Visibility\n& Speed",
     "Organizations must focus on decision quality and governance clarity, not just information availability."),
]

for i, (title, desc) in enumerate(conclusions):
    left = 0.8 + i * 4.0
    add_circle_icon(slide, left + 1.3, 4.1, 0.7, GOLD, str(i+1), 18, NAVY)
    add_textbox(slide, left + 0.15, 4.95, 3.6, 0.6, title, 13, GOLD, True, PP_ALIGN.CENTER)
    add_textbox(slide, left + 0.15, 5.6, 3.6, 1.2, desc, 11, RGBColor(0xCC, 0xDD, 0xEE), False, PP_ALIGN.CENTER)

add_slide_number(slide, 21, TOTAL_SLIDES)

# ─── SLIDE 22: Limitations & Future Research + Thank You ─
slide = prs.slides.add_slide(prs.slide_layouts[6])
add_bg(slide, NAVY)
add_section_header(slide, "Limitations & Future Research")
add_bottom_bar(slide)

# Limitations
add_rounded_rect(slide, 0.8, 2.2, 5.5, 3.0, DARK_BLUE)
add_textbox(slide, 1.0, 2.4, 5.0, 0.4, "Limitations", 15, GOLD, True)
lim_items = [
    "Cross-sectional design limits causal inference",
    "Sample drawn from India may limit generalizability",
    "Self-reported perceptions may introduce response bias",
    "Focuses on direct relationships only (no mediation/moderation)",
]
add_multiline_textbox(slide, 1.0, 3.0, 5.0, 2.0, lim_items, 11,
                      RGBColor(0xCC, 0xDD, 0xEE), bullet=True, line_spacing=1.4)

# Future research
add_rounded_rect(slide, 7.0, 2.2, 5.5, 3.0, MID_BLUE)
add_textbox(slide, 7.2, 2.4, 5.0, 0.4, "Future Research Directions", 15, GOLD, True)
future_items = [
    "Longitudinal studies across project phases",
    "Cross-cultural replication in different regions",
    "Explore mediating roles of organizational culture and project complexity",
    "Examine AI interaction with specific Agile practices (sprints, retrospectives)",
]
add_multiline_textbox(slide, 7.2, 3.0, 5.0, 2.0, future_items, 11,
                      RGBColor(0xDD, 0xEE, 0xFF), bullet=True, line_spacing=1.4)

# Thank you section
add_rounded_rect(slide, 3.0, 5.8, 7.3, 1.2, GOLD)
add_textbox(slide, 3.2, 5.85, 6.9, 0.6, "Thank You", 32, NAVY, True, PP_ALIGN.CENTER, "Calibri Light")
add_textbox(slide, 3.2, 6.45, 6.9, 0.4, "Questions & Discussion Welcome", 14, NAVY, False, PP_ALIGN.CENTER)

add_slide_number(slide, 22, TOTAL_SLIDES)


# ═══════════════════════════════════════════════════════════
# SAVE
# ═══════════════════════════════════════════════════════════
output_path = os.path.join(OUT_DIR, "AI_Risk_Management_Thesis_Presentation.pptx")
prs.save(output_path)
print(f"\nPresentation saved to: {output_path}")
print(f"Total slides: {TOTAL_SLIDES}")
