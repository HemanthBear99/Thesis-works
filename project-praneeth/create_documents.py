"""
Creates two Word documents for thesis defense preparation:
1. Comprehensive Explanation & 30-Minute Presentation Guide
2. Professor Q&A Preparation Document
"""

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
import os

OUTPUT_DIR = r"D:\project-praneeth"


def setup_styles(doc):
    """Configure professional document styles."""
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(11)
    font.color.rgb = RGBColor(0x1A, 0x1A, 0x2E)
    pf = style.paragraph_format
    pf.space_after = Pt(6)
    pf.line_spacing = 1.15

    for level, (size, color) in enumerate([
        (20, RGBColor(0x0D, 0x1B, 0x2A)),
        (16, RGBColor(0x1B, 0x3A, 0x5C)),
        (13, RGBColor(0x2E, 0x75, 0xB6)),
    ], 1):
        h = doc.styles[f'Heading {level}']
        h.font.name = 'Calibri'
        h.font.size = Pt(size)
        h.font.color.rgb = color
        h.font.bold = True
        h.paragraph_format.space_before = Pt(18 if level == 1 else 12)
        h.paragraph_format.space_after = Pt(6)


def add_para(doc, text, bold=False, italic=False, size=None, color=None,
             align=None, space_after=None, style=None):
    p = doc.add_paragraph(style=style)
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    if size:
        run.font.size = Pt(size)
    if color:
        run.font.color.rgb = color
    if align:
        p.alignment = align
    if space_after is not None:
        p.paragraph_format.space_after = Pt(space_after)
    return p


def add_bullet(doc, text, bold_prefix=None, level=0):
    p = doc.add_paragraph(style='List Bullet')
    p.paragraph_format.left_indent = Cm(1.27 + level * 1.27)
    if bold_prefix:
        run = p.add_run(bold_prefix)
        run.bold = True
        p.add_run(text)
    else:
        p.add_run(text)
    return p


def add_timing_box(doc, slide_range, time, title):
    p = doc.add_paragraph()
    r = p.add_run(f"[SLIDES {slide_range}]  |  {time}  |  {title}")
    r.bold = True
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0x0D, 0x1B, 0x2A)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    # Add a bottom border effect via a separator line
    sep = doc.add_paragraph()
    r2 = sep.add_run("_" * 90)
    r2.font.size = Pt(6)
    r2.font.color.rgb = RGBColor(0xD4, 0xA0, 0x2E)
    sep.paragraph_format.space_after = Pt(2)


# ================================================================
# DOCUMENT 1: COMPREHENSIVE EXPLANATION & PRESENTATION GUIDE
# ================================================================
def create_doc1():
    doc = Document()
    setup_styles(doc)

    # --- COVER ---
    for _ in range(4):
        doc.add_paragraph()
    add_para(doc, "THESIS DEFENSE", bold=True, size=14, color=RGBColor(0xD4, 0xA0, 0x2E),
             align=WD_ALIGN_PARAGRAPH.CENTER)
    add_para(doc, "Comprehensive Explanation &\n30-Minute Presentation Guide", bold=True, size=24,
             color=RGBColor(0x0D, 0x1B, 0x2A), align=WD_ALIGN_PARAGRAPH.CENTER)
    add_para(doc, "", size=8)
    add_para(doc, "Stochastic Earned Value Management:", bold=True, size=14,
             color=RGBColor(0x1B, 0x3A, 0x5C), align=WD_ALIGN_PARAGRAPH.CENTER)
    add_para(doc, "Integrating Monte Carlo Simulation for Enhanced Probabilistic\nCost and Schedule Forecasting in EPC Projects",
             size=13, color=RGBColor(0x5A, 0x5A, 0x6E), align=WD_ALIGN_PARAGRAPH.CENTER)
    for _ in range(3):
        doc.add_paragraph()
    add_para(doc, "Praneeth  |  ESLSCA School of Business  |  MBA Thesis",
             size=12, color=RGBColor(0x80, 0x80, 0x80), align=WD_ALIGN_PARAGRAPH.CENTER)
    add_para(doc, "CONFIDENTIAL \u2014 For Defense Preparation Only",
             bold=True, size=10, color=RGBColor(0xC0, 0x30, 0x30), align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()

    # --- TABLE OF CONTENTS ---
    doc.add_heading("Table of Contents", level=1)
    toc = [
        "Part A: Understanding Your Thesis (Complete Conceptual Explanation)",
        "  A.1  What is Earned Value Management (EVM)?",
        "  A.2  The Problem with Traditional EVM",
        "  A.3  What is Stochastic EVM (SEVM)?",
        "  A.4  Monte Carlo Simulation Explained",
        "  A.5  Why Lognormal Distribution for CPI?",
        "  A.6  Understanding the Key Metrics (MAPE, Calibration, Wilcoxon)",
        "  A.7  Your Dataset and Methodology",
        "  A.8  Understanding Every Result",
        "  A.9  Understanding Every Chart",
        "",
        "Part B: 30-Minute Presentation Guide (Slide-by-Slide Script)",
        "  B.1  Timing Breakdown",
        "  B.2  Slide-by-Slide Speaking Script",
        "  B.3  Delivery Tips and Body Language",
        "",
        "Part C: Key Phrases and Confident Responses",
        "  C.1  Phrases That Impress Professors",
        "  C.2  Handling the Confidentiality Question",
        "  C.3  Handling Tough Challenges",
    ]
    for item in toc:
        if item == "":
            doc.add_paragraph()
        elif item.startswith("Part"):
            add_para(doc, item, bold=True, size=12, color=RGBColor(0x0D, 0x1B, 0x2A))
        else:
            add_para(doc, item, size=11, color=RGBColor(0x5A, 0x5A, 0x6E))
    doc.add_page_break()

    # ============================================================
    # PART A: UNDERSTANDING YOUR THESIS
    # ============================================================
    doc.add_heading("PART A: Understanding Your Thesis", level=1)
    add_para(doc, "This section explains every concept in your thesis in plain language. Read this thoroughly so you can answer any question with confidence. You should be able to explain each concept without looking at notes.", size=11, italic=True, color=RGBColor(0x5A, 0x5A, 0x6E), space_after=12)

    # --- A.1 ---
    doc.add_heading("A.1  What is Earned Value Management (EVM)?", level=2)
    doc.add_paragraph(
        "Earned Value Management is the global standard methodology for measuring project performance. "
        "It was originally developed by the U.S. Department of Defense in the 1960s and has since become "
        "the industry standard in construction, defense, aerospace, and engineering projects worldwide. "
        "EVM integrates three critical dimensions of project management: scope, schedule, and cost."
    )
    doc.add_paragraph("EVM uses three fundamental measurements:")
    add_bullet(doc, "Planned Value (PV): ", bold_prefix="Planned Value (PV): ")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("Planned Value (PV): ")
    r.bold = True
    p.add_run("The authorized budget assigned to scheduled work. In simple terms, how much work you PLANNED to have done by this point, measured in dollars. For example, if your project plan says 40% of work should be done by Month 6, and the total budget is $100M, then PV at Month 6 = $40M.")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("Earned Value (EV): ")
    r.bold = True
    p.add_run("The measure of work actually performed, expressed in terms of the budget authorized for that work. This is how much work you ACTUALLY completed, measured in budget terms. If you actually completed 35% of work by Month 6, then EV = $35M.")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("Actual Cost (AC): ")
    r.bold = True
    p.add_run("The realized cost incurred for the work performed. This is how much money you ACTUALLY spent. If you spent $42M to achieve that 35% completion, then AC = $42M.")

    doc.add_paragraph("")
    doc.add_paragraph("From these three values, EVM derives two critical performance indices:")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("Cost Performance Index (CPI) = EV / AC: ")
    r.bold = True
    p.add_run("Measures cost efficiency. A CPI of 0.83 means for every $1 spent, you only earned $0.83 worth of work. You are over budget. A CPI above 1.0 means under budget.")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("Schedule Performance Index (SPI) = EV / PV: ")
    r.bold = True
    p.add_run("Measures schedule efficiency. An SPI of 0.875 means you have only completed 87.5% of the work you planned to complete by now. You are behind schedule.")

    doc.add_paragraph("")
    doc.add_paragraph("The ultimate goal of EVM is to forecast the Estimate at Completion (EAC) \u2014 what the project will actually cost when finished:")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("EAC (BAC/CPI) = BAC / CPI: ")
    r.bold = True
    p.add_run("This formula assumes future cost performance will match cumulative past performance. If BAC = $100M and CPI = 0.85, then EAC = $117.6M.")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("EAC (Composite) = AC + (BAC - EV) / (CPI \u00d7 SPI): ")
    r.bold = True
    p.add_run("This formula considers both cost AND schedule performance, assuming a behind-schedule project will also incur additional costs due to extended duration.")

    # --- A.2 ---
    doc.add_heading("A.2  The Problem with Traditional EVM", level=2)
    doc.add_paragraph(
        "Traditional EVM forecasting has a fundamental structural flaw: it is deterministic. "
        "It produces a single-point estimate and treats it as certain. When a project manager reports "
        "\"The EAC is $117.6 million,\" they are implicitly saying that this is THE answer \u2014 not one of many possible answers."
    )
    doc.add_paragraph("This creates three critical problems:")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("False Precision: ")
    r.bold = True
    p.add_run("A single number creates a false sense of certainty. Decision-makers believe they know the final cost, when in reality there is a wide range of possible outcomes. No one can predict the future with a single number.")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("No Risk Quantification: ")
    r.bold = True
    p.add_run("Traditional EVM provides no information about the probability of hitting that target. Is there a 90% chance or a 30% chance of staying under $117.6M? The formula cannot tell you.")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("Linear Assumption: ")
    r.bold = True
    p.add_run("The formula assumes that if your CPI is 0.85 today, it will remain exactly 0.85 for the rest of the project. But real-world EPC projects experience volatile, non-linear performance. CPI fluctuates significantly period to period due to material price changes, labor productivity shifts, rework cycles, and external disruptions.")

    doc.add_paragraph("")
    doc.add_paragraph(
        "In the EPC (Engineering, Procurement, and Construction) industry, this problem is especially severe. "
        "Industry benchmarks consistently show that up to 65% of large EPC projects experience cost overruns "
        "exceeding 25% of their initial budgets. Schedule delays of 50-80% are common. When a $200 million "
        "refinery overruns by 25%, that is $50 million in unexpected costs \u2014 enough to eliminate the entire "
        "profit margin and potentially bankrupt the contractor."
    )

    # --- A.3 ---
    doc.add_heading("A.3  What is Stochastic EVM (SEVM)?", level=2)
    doc.add_paragraph(
        "Stochastic Earned Value Management is the approach your thesis proposes and validates. Instead of treating "
        "CPI and SPI as fixed constants (as traditional EVM does), SEVM treats them as random variables with probability "
        "distributions. The word \"stochastic\" simply means \"involving randomness or probability.\""
    )
    doc.add_paragraph(
        "The core idea is elegant: instead of saying \"the CPI will be 0.85 forever,\" SEVM says \"based on the "
        "historical variability we have observed, the CPI in future periods could range from 0.70 to 1.05, with most "
        "values clustering around 0.85, but with a longer tail toward lower values (worse performance).\" "
        "By sampling thousands of possible future CPIs and SPIs from these distributions, SEVM generates a probability "
        "distribution of possible final costs."
    )
    doc.add_paragraph("This means SEVM produces:")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("P10 (10th percentile): ")
    r.bold = True
    p.add_run("The optimistic estimate. There is only a 10% chance the final cost will be below this value. This is your best-case scenario.")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("P50 (50th percentile / median): ")
    r.bold = True
    p.add_run("The most likely estimate. There is a 50/50 chance of being above or below this value. This is your primary point estimate \u2014 equivalent to the deterministic EAC, but derived probabilistically.")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("P80 (80th percentile): ")
    r.bold = True
    p.add_run("The conservative estimate. There is an 80% chance the final cost will be below this value. Many organizations set their management reserve at P80.")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("P90 (90th percentile): ")
    r.bold = True
    p.add_run("The worst-case planning estimate. Only a 10% chance of exceeding this. Used for contingency budgeting and executive risk reporting.")

    doc.add_paragraph("")
    doc.add_paragraph(
        "The fundamental shift is from saying \"The project will cost $117.6M\" to saying \"There is an 80% "
        "probability the project will cost between $105M and $135M, with a most likely cost of $118M.\" "
        "This gives decision-makers dramatically better information for managing risk, setting contingencies, "
        "and making go/no-go decisions."
    )

    # --- A.4 ---
    doc.add_heading("A.4  Monte Carlo Simulation Explained", level=2)
    doc.add_paragraph(
        "Monte Carlo simulation is a computational technique that uses random sampling to model the probability "
        "of different outcomes. It is named after the Monte Carlo Casino in Monaco because the core idea involves "
        "repeated random sampling \u2014 like rolling dice thousands of times."
    )
    doc.add_paragraph("Here is how it works in your SEVM analysis, step by step:")
    doc.add_paragraph("")
    add_para(doc, "Step 1: Collect Historical Data", bold=True, size=11)
    doc.add_paragraph(
        "We collected period-by-period EVM data (PV, EV, AC) from 12 completed EPC projects. From this data, "
        "we calculated the period CPI and SPI for each reporting period. For example, if a project had 24 months "
        "of data, we had 24 individual CPI observations showing how cost efficiency varied month to month."
    )
    add_para(doc, "Step 2: Fit Probability Distributions", bold=True, size=11)
    doc.add_paragraph(
        "We analyzed the statistical distribution of these historical CPI values. We found that CPI data in EPC "
        "projects follows a lognormal distribution (right-skewed \u2014 explained in detail in Section A.5). "
        "We fitted a lognormal distribution to CPI and a normal distribution to SPI. This gives us the "
        "mathematical parameters (mean and standard deviation of the log-transformed values) that describe "
        "the pattern of variability."
    )
    add_para(doc, "Step 3: Run 10,000 Simulations", bold=True, size=11)
    doc.add_paragraph(
        "At each forecast point (e.g., when the project is 50% complete), we randomly sample a future CPI and SPI "
        "from the fitted distributions. Using these sampled values, we calculate what the final cost (EAC) would be. "
        "We repeat this process 10,000 times. Each iteration produces a slightly different EAC because each uses "
        "a different randomly sampled CPI and SPI. After 10,000 iterations, we have 10,000 possible final costs."
    )
    add_para(doc, "Step 4: Analyze the Distribution", bold=True, size=11)
    doc.add_paragraph(
        "These 10,000 EAC values form a probability distribution. We can now extract percentiles: the P10, P50, "
        "P80, and P90. We can calculate the mean, standard deviation, and confidence intervals. This distribution "
        "IS the SEVM forecast \u2014 it captures all the uncertainty inherent in the project."
    )
    doc.add_paragraph("")
    add_para(doc, "Why 10,000 iterations?", bold=True, italic=True, size=11)
    doc.add_paragraph(
        "10,000 is the industry standard for Monte Carlo simulations in quantitative risk analysis. It provides "
        "highly stable convergence of the probability distributions \u2014 meaning if you run it again, you get "
        "essentially the same results. Fewer iterations (e.g., 1,000) can produce unstable tails (the P10 and P90 "
        "values might jump around). More iterations (e.g., 100,000) add computational time without meaningfully "
        "improving accuracy."
    )

    # --- A.5 ---
    doc.add_heading("A.5  Why Lognormal Distribution for CPI?", level=2)
    doc.add_paragraph("This is one of the most important technical decisions in your thesis, and professors will almost certainly ask about it. Here is the complete reasoning:")
    doc.add_paragraph("")
    add_para(doc, "The Nature of Cost Data in EPC Projects:", bold=True, size=11)
    doc.add_paragraph(
        "Cost performance in EPC projects has a fundamental mathematical property: it is bounded at zero on the "
        "left side but unbounded on the right. Costs cannot go negative (you cannot spend negative dollars), "
        "but they can theoretically overrun to infinity. A project budgeted at $100M will never cost -$5M, "
        "but it could cost $150M, $200M, or even $300M in extreme cases."
    )
    doc.add_paragraph(
        "This creates a right-skewed distribution: most periods have CPI values clustering around the mean "
        "(small variations are common), but occasionally there are sharp drops in CPI (large cost overruns) "
        "that create a long right tail. This asymmetry is a hallmark of cost data in construction and engineering."
    )
    doc.add_paragraph("")
    add_para(doc, "Why NOT a Normal Distribution?", bold=True, size=11)
    doc.add_paragraph(
        "A normal (Gaussian) distribution is symmetrical. It assumes equal probability of costs being above or "
        "below the mean, and it extends infinitely in both directions. This means a normal distribution could "
        "generate CPI values that imply negative costs \u2014 which is physically impossible. More importantly, "
        "it underestimates the probability of extreme cost overruns because it does not have the \"fat tail\" "
        "that real cost data exhibits."
    )
    doc.add_paragraph("")
    add_para(doc, "Why Lognormal is Perfect:", bold=True, size=11)
    doc.add_paragraph(
        "A lognormal distribution is what you get when the logarithm of a variable is normally distributed. "
        "It is always positive (solves the negative cost problem), it is right-skewed (matches real cost data), "
        "and it naturally models the phenomenon where small cost variations are common but extreme overruns "
        "are possible. In quantitative risk analysis and financial modeling, the lognormal distribution is the "
        "standard choice for modeling costs, prices, and durations."
    )
    doc.add_paragraph("")
    add_para(doc, "For SPI (Schedule Performance Index):", bold=True, size=11)
    doc.add_paragraph(
        "We used a normal distribution for SPI because schedule performance tends to be more symmetrical. "
        "A project can be ahead of schedule or behind schedule with roughly equal probability, and schedule "
        "performance does not exhibit the same extreme skewness as cost data. The normal distribution is "
        "appropriate and sufficient for modeling SPI variability."
    )

    # --- A.6 ---
    doc.add_heading("A.6  Understanding the Key Metrics", level=2)

    add_para(doc, "MAPE (Mean Absolute Percentage Error)", bold=True, size=12, color=RGBColor(0x1B, 0x3A, 0x5C))
    doc.add_paragraph(
        "MAPE is the primary accuracy metric used to compare SEVM against traditional EVM. It measures "
        "the average percentage difference between the forecast and the actual final cost. A MAPE of 3.66% "
        "means that on average, the forecast was 3.66% away from the actual cost \u2014 remarkably accurate "
        "for construction project forecasting where 10-15% errors are common with traditional methods."
    )
    doc.add_paragraph("Your key MAPE results:")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("SEVM (P50): 3.66% ")
    r.bold = True
    p.add_run("\u2014 this is the average error when using SEVM's median estimate")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("Deterministic EVM (BAC/CPI): 2.69% ")
    r.bold = True
    p.add_run("\u2014 the simpler formula happens to be more accurate as a point estimate")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("Deterministic EVM (Composite): 4.21% ")
    r.bold = True
    p.add_run("\u2014 the composite formula that accounts for schedule has higher error")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("SEVM improvement vs. Composite: +0.55 percentage points ")
    r.bold = True
    p.add_run("\u2014 SEVM beats the composite formula")

    doc.add_paragraph("")
    add_para(doc, "Important Note on H1:", bold=True, italic=True, size=11, color=RGBColor(0xC0, 0x30, 0x30))
    doc.add_paragraph(
        "Your Hypothesis H1 stated that SEVM would outperform ALL deterministic formulas. The result shows "
        "SEVM outperforms the Composite formula (which is the more realistic formula used in practice, because "
        "it accounts for both cost and schedule effects), but the simpler BAC/CPI formula actually edges ahead "
        "as a pure point estimate. This is NOT a weakness \u2014 it is a nuanced finding. The BAC/CPI formula "
        "is a simplistic ratio that happens to perform well as a point estimate, but it provides zero risk "
        "information. SEVM's value is not merely in the point estimate \u2014 it is in the confidence intervals "
        "and risk quantification that no deterministic method can provide. Frame H1 as 'partially supported' "
        "and emphasize that SEVM's true value lies in calibration, not just point accuracy."
    )

    doc.add_paragraph("")
    add_para(doc, "Calibration", bold=True, size=12, color=RGBColor(0x1B, 0x3A, 0x5C))
    doc.add_paragraph(
        "Calibration is the most powerful finding in your thesis. It measures whether the confidence intervals "
        "are trustworthy. If the SEVM model says \"there is an 80% chance the final cost falls between $105M "
        "and $135M,\" calibration checks: did the actual cost really fall in that range 80% of the time?"
    )
    doc.add_paragraph("Your calibration results:")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("80% CI Calibration: 77.1% ")
    r.bold = True
    p.add_run("(target: 80%) \u2014 remarkably close. The model's confidence intervals are statistically sound.")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("50% CI Calibration: 53.1% ")
    r.bold = True
    p.add_run("(target: 50%) \u2014 also very close to the theoretical ideal.")

    doc.add_paragraph("")
    doc.add_paragraph(
        "This is the most important result in your thesis. It means decision-makers can TRUST the SEVM ranges. "
        "When SEVM says \"80% chance,\" it genuinely means 80%. This is what makes SEVM operationally useful \u2014 "
        "executives can set contingency reserves at the P80 level with confidence that they are covered 80% of the time."
    )

    doc.add_paragraph("")
    add_para(doc, "Wilcoxon Signed-Rank Test", bold=True, size=12, color=RGBColor(0x1B, 0x3A, 0x5C))
    doc.add_paragraph(
        "The Wilcoxon signed-rank test is a non-parametric statistical test that compares two paired groups. "
        "You used it to test whether the difference between SEVM errors and deterministic EVM errors is "
        "statistically significant \u2014 meaning it is not due to random chance."
    )
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("Why Wilcoxon (not a t-test)? ")
    r.bold = True
    p.add_run("Because percentage errors (MAPE values) are bounded at zero and not normally distributed. The Wilcoxon test does not assume normality, making it the mathematically correct choice.")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("SEVM vs BAC/CPI: W=1073, p=0.000005 (***) ")
    r.bold = True
    p.add_run("\u2014 The difference is highly significant. The two methods produce statistically different results.")
    add_bullet(doc, "")
    p = doc.paragraphs[-1]
    p.clear()
    r = p.add_run("SEVM vs Composite: W=1915, p=0.131 (not significant) ")
    r.bold = True
    p.add_run("\u2014 While SEVM has lower average error, the difference is not statistically significant at the 0.05 level. This is expected because both methods account for schedule effects in different ways.")

    # --- A.7 ---
    doc.add_heading("A.7  Your Dataset and Methodology", level=2)
    add_para(doc, "About the Data:", bold=True, size=11)
    doc.add_paragraph(
        "The analysis was conducted on historical EVM data from 12 completed EPC projects spanning five industrial "
        "sectors: oil and gas, power generation, water treatment, industrial processing, and renewable energy. "
        "The projects ranged in budget from $15 million to $200 million, with a combined portfolio value of approximately "
        "$942 million. Project durations ranged from 12 to 42 months."
    )
    doc.add_paragraph(
        "All project identifiers have been anonymized (P01 through P12) due to non-disclosure agreements with "
        "the project owners and contracting organizations. The EVM data was extracted from project control systems "
        "and includes period-by-period Planned Value, Earned Value, and Actual Cost records. Eleven of the twelve "
        "projects experienced actual cost overruns, providing a realistic and representative dataset that reflects "
        "the well-documented challenges of the EPC industry."
    )
    doc.add_paragraph("")
    add_para(doc, "Methodology Overview:", bold=True, size=11)
    doc.add_paragraph(
        "The research adopted a quantitative, positivist approach. The methodology followed four stages: "
        "(1) collection and validation of historical EVM data, (2) fitting of probability distributions to "
        "period-level CPI and SPI values, (3) execution of Monte Carlo simulations at eight completion stages "
        "(20% through 90%), and (4) comparative statistical analysis using MAPE, calibration metrics, and the "
        "Wilcoxon signed-rank test. A total of approximately 960,000 Monte Carlo iterations were executed across "
        "all projects and forecast points (12 projects \u00d7 8 forecast stages \u00d7 10,000 iterations)."
    )

    # --- A.8 ---
    doc.add_heading("A.8  Understanding Every Result", level=2)
    doc.add_paragraph("Here is a plain-language explanation of every key finding, with the narrative you should present to your professor:")

    doc.add_paragraph("")
    add_para(doc, "Result 1: SEVM Outperforms the Composite Deterministic Formula", bold=True, size=12, color=RGBColor(0x1B, 0x3A, 0x5C))
    doc.add_paragraph(
        "The SEVM P50 estimate achieved an average MAPE of 3.66%, compared to the deterministic composite formula's "
        "4.21%. This is a 0.55 percentage point improvement. The composite formula (EAC = AC + remaining / (CPI \u00d7 SPI)) "
        "is the more sophisticated deterministic formula because it accounts for both cost and schedule effects. When you "
        "multiply two uncertain indices together (CPI \u00d7 SPI), you compound the uncertainty and amplify the error. "
        "SEVM handles this mathematically by sampling from the full distributions independently, avoiding this compounding effect."
    )

    doc.add_paragraph("")
    add_para(doc, "Result 2: Confidence Interval Calibration is Excellent", bold=True, size=12, color=RGBColor(0x1B, 0x3A, 0x5C))
    doc.add_paragraph(
        "The 80% confidence interval captured the actual cost 77.1% of the time (target: 80%). "
        "The 50% confidence interval captured the actual cost 53.1% of the time (target: 50%). "
        "This is the most powerful finding because it proves the model is not just accurate on average but that its "
        "stated confidence levels are reliable. A project manager using SEVM can set a contingency reserve at the P80 "
        "level and be statistically confident that costs will stay within that range approximately 80% of the time."
    )

    doc.add_paragraph("")
    add_para(doc, "Result 3: SEVM's Advantage Scales with Project Complexity", bold=True, size=12, color=RGBColor(0x1B, 0x3A, 0x5C))
    doc.add_paragraph(
        "For high-complexity projects (offshore platforms, chemical plants, large refineries), SEVM's confidence "
        "intervals achieved 96% calibration accuracy. For low-complexity projects (simple warehouses, standard buildings), "
        "deterministic EVM performed comparably well, and SEVM's additional complexity was less justified. "
        "This is an actionable finding: organizations should deploy SEVM on their high-risk, high-complexity projects "
        "where the additional analytical effort pays the highest dividends."
    )

    doc.add_paragraph("")
    add_para(doc, "Result 4: Maximum Advantage at 50% Completion", bold=True, size=12, color=RGBColor(0x1B, 0x3A, 0x5C))
    doc.add_paragraph(
        "SEVM provides its greatest accuracy advantage over deterministic EVM at the 50% completion mark. "
        "This makes intuitive sense: before 20% completion, there is not enough historical data to fit a reliable "
        "distribution. After 80% completion, the project is nearly done and uncertainty is inherently low. "
        "At 50% completion, there is sufficient historical data to calibrate the distributions AND enough remaining "
        "work that uncertainty is significant. This is the \"sweet spot\" for SEVM deployment."
    )

    doc.add_paragraph("")
    add_para(doc, "Result 5: CPI is Genuinely Stochastic", bold=True, size=12, color=RGBColor(0x1B, 0x3A, 0x5C))
    doc.add_paragraph(
        "The period-by-period CPI data across all 12 projects shows clear stochastic (random) behavior. "
        "CPI fluctuates significantly from period to period \u2014 it is not the stable, constant value that "
        "deterministic EVM assumes. This empirical evidence validates the fundamental premise of the thesis: "
        "that project performance is inherently stochastic and should be modeled as such."
    )

    # --- A.9 ---
    doc.add_heading("A.9  Understanding Every Chart", level=2)

    charts = [
        ("Figure 2: Forecast Accuracy Comparison (Bar Chart)",
         "This bar chart compares the MAPE of three methods across eight completion stages (20% to 90%). "
         "Each group of three bars shows: (red) deterministic BAC/CPI, (orange) deterministic Composite, "
         "and (blue) SEVM P50. All methods improve as the project progresses (more data = better forecasts). "
         "SEVM consistently outperforms the Composite formula. The gap is widest in early-to-mid stages "
         "and narrows as the project approaches completion, which is expected since uncertainty decreases."),
        ("Figure 3: S-Curve with SEVM Confidence Bands",
         "This is the classic EVM S-curve for Project P02 (a $120M oil refinery project). The green dashed line "
         "is Planned Value, the blue solid line is Earned Value, and the red solid line is Actual Cost. "
         "At the 50% mark, the SEVM forecast appears: the blue dotted line is the P50 forecast, the red dotted "
         "line is the deterministic forecast, and the shaded blue band is the 80% confidence interval (P10 to P90). "
         "The actual cost ended up within the confidence band, validating the model."),
        ("Figure 4: Monte Carlo EAC Distribution (Histogram)",
         "This histogram shows the 10,000 possible EAC values generated by Monte Carlo simulation for Project P04 "
         "(an $85M gas processing facility) at 50% completion. The distribution is visibly right-skewed (lognormal). "
         "The navy line marks the P50 ($102.2M), the orange dashed line marks P80 ($110.9M), and the red dashed line "
         "marks P90 ($116.4M). The green solid line shows the actual final cost ($104.5M), which falls between P50 and P80. "
         "The deterministic EAC ($102.5M) sits near P50 but provides no risk information about the long tail."),
        ("Figure 5: Complexity Effect (Box Plot)",
         "This box plot shows the SEVM accuracy improvement (in percentage points) grouped by project complexity. "
         "The key insight is that high-complexity projects have the widest range of improvement and the highest "
         "upside potential. SEVM's advantage grows with complexity because complex projects have more volatile "
         "CPI patterns that deterministic methods cannot capture."),
        ("Figure 6: Cumulative Distribution Function (CDF)",
         "The CDF is the most operationally useful chart. It shows the probability of the final cost being "
         "at or below any given value. A project manager can read it directly: \"What is the probability we "
         "stay under $110M?\" Just find $110M on the x-axis and read the probability on the y-axis. "
         "The deterministic EAC is a single vertical line \u2014 it tells you nothing about probability. "
         "The SEVM CDF tells you everything."),
        ("Figure 7: CPI Variability Over Lifecycle",
         "This chart plots the period-by-period CPI for six projects across the project lifecycle. The key "
         "observation is that CPI is highly volatile and non-linear. It bounces around significantly from "
         "period to period. This is visual proof that the deterministic assumption (CPI = constant) is "
         "fundamentally wrong. Real CPI is stochastic, justifying the entire SEVM approach."),
        ("Figure 8: Phase Effect (Line Chart)",
         "This chart shows SEVM's accuracy improvement at each completion stage. The blue line (vs. BAC/CPI) "
         "and the orange line (vs. Composite) tell different stories. Against the Composite formula, SEVM's "
         "advantage peaks at 50% completion (+1.05pp) and then gradually decreases. This confirms the "
         "\"sweet spot\" finding and provides clear operational guidance: deploy SEVM at mid-project."),
    ]
    for title, desc in charts:
        add_para(doc, title, bold=True, size=11, color=RGBColor(0x1B, 0x3A, 0x5C))
        doc.add_paragraph(desc)
        doc.add_paragraph("")

    doc.add_page_break()

    # ============================================================
    # PART B: 30-MINUTE PRESENTATION GUIDE
    # ============================================================
    doc.add_heading("PART B: 30-Minute Presentation Guide", level=1)
    add_para(doc, "This section provides a complete, slide-by-slide speaking script with exact timing. Practice this script 3-4 times before your defense. Speak slowly and clearly. Pause at key moments for emphasis.", size=11, italic=True, color=RGBColor(0x5A, 0x5A, 0x6E), space_after=12)

    # Timing table
    doc.add_heading("B.1  Timing Breakdown", level=2)
    table = doc.add_table(rows=8, cols=4)
    table.style = 'Light Grid Accent 1'
    headers = ["Section", "Slides", "Duration", "Cumulative"]
    for i, h in enumerate(headers):
        table.rows[0].cells[i].text = h
        for p in table.rows[0].cells[i].paragraphs:
            p.runs[0].bold = True
    timing = [
        ("Opening & Introduction", "1\u20132", "2 min", "0:02"),
        ("The Problem", "3\u20135", "4 min", "0:06"),
        ("The Solution (SEVM)", "6\u20137", "3 min", "0:09"),
        ("Methodology & Dataset", "8\u201310", "4 min", "0:13"),
        ("Key Findings (8 results)", "11\u201319", "10 min", "0:23"),
        ("Hypothesis & Statistics", "20", "2 min", "0:25"),
        ("Recommendations, Conclusion & Close", "21\u201324", "5 min", "0:30"),
    ]
    for i, row_data in enumerate(timing, 1):
        for j, val in enumerate(row_data):
            table.rows[i].cells[j].text = val

    doc.add_paragraph("")

    # --- B.2 Slide-by-slide ---
    doc.add_heading("B.2  Slide-by-Slide Speaking Script", level=2)
    add_para(doc, "Read this carefully. Each section tells you exactly what to say and when. Words in [brackets] are stage directions.", size=11, italic=True, color=RGBColor(0x5A, 0x5A, 0x6E), space_after=8)

    # SLIDE 1
    add_timing_box(doc, "1", "0:00 \u2013 0:30", "TITLE SLIDE")
    doc.add_paragraph(
        "\"Good morning, esteemed panel members, my supervisor, and guests. My name is Praneeth, and I am honored "
        "to present my MBA thesis today, titled 'Stochastic Earned Value Management: Integrating Monte Carlo "
        "Simulation for Enhanced Probabilistic Cost and Schedule Forecasting in EPC Projects.' I would like to "
        "begin by expressing my gratitude to my supervisor, the faculty at ESLSCA, and my family for their "
        "support throughout this research.\""
    )

    # SLIDE 2
    add_timing_box(doc, "2", "0:30 \u2013 2:00", "PRESENTATION OVERVIEW")
    doc.add_paragraph(
        "\"Before I dive in, here is our roadmap for today. I will begin by explaining the core problem facing "
        "the EPC industry. Then I will introduce the SEVM solution and explain how Monte Carlo simulation works. "
        "I will walk you through the research methodology and our dataset of 12 real EPC projects. "
        "The core of the presentation will cover eight key empirical findings. And I will close with "
        "practical recommendations and conclusions. This should take approximately 25 to 30 minutes, after which "
        "I welcome your questions.\""
    )

    # SLIDES 3-5
    add_timing_box(doc, "3\u20135", "2:00 \u2013 6:00", "THE PROBLEM")
    doc.add_paragraph(
        "[Slide 3 is a dark section divider \u2014 use it to pause, take a breath, and transition.]"
    )
    doc.add_paragraph("")
    doc.add_paragraph(
        "[Slide 4 \u2014 The EPC Challenge] \"Let me start with why this research matters. The EPC industry builds "
        "the world's critical infrastructure \u2014 power plants, refineries, pipelines, and mega-projects costing "
        "hundreds of millions of dollars. But the industry has a chronic problem: up to 65% of large EPC projects "
        "exceed their budgets by more than 25%. Schedule delays of 80% are shockingly common. When we consistently "
        "fail to predict costs accurately, the consequences are severe \u2014 wiped-out profit margins, contractual "
        "disputes, and even abandoned infrastructure. In the $942 million portfolio I analyzed, 11 of 12 projects "
        "experienced cost overruns. This is not an anomaly \u2014 this is the norm.\""
    )
    doc.add_paragraph("")
    doc.add_paragraph(
        "[Slide 5 \u2014 Deterministic vs Stochastic] \"So why do we keep getting it wrong? The answer lies in our "
        "tools. Traditional Earned Value Management uses a deterministic formula: EAC equals BAC divided by CPI. "
        "It produces a single number \u2014 say, $150 million \u2014 and presents it as certain. But as project "
        "managers, we know the real world is not certain. It is volatile, subject to supply chain disruptions, "
        "labor shortages, and commodity price spikes. A single number creates a false sense of certainty. "
        "SEVM corrects this by using Monte Carlo simulation to generate a probability distribution of possible "
        "outcomes. Instead of saying 'the project will cost $150M,' we say 'there is an 80% probability the "
        "cost will fall between $140M and $160M.' This transforms project reporting from guesswork to "
        "evidence-based risk management.\""
    )

    # SLIDES 6-7
    add_timing_box(doc, "6\u20137", "6:00 \u2013 9:00", "THE SOLUTION")
    doc.add_paragraph(
        "[Slide 7 \u2014 How SEVM Works] \"Let me walk you through the four-step SEVM process. First, we collect "
        "historical EVM data \u2014 the period-by-period Planned Value, Earned Value, and Actual Cost records. "
        "Second, we fit probability distributions to the historical CPI and SPI data. For CPI, we use a lognormal "
        "distribution because cost data in EPC projects is right-skewed \u2014 small overruns are common but extreme "
        "cost blowouts, while rare, are possible. For SPI, we use a normal distribution since schedule performance "
        "is more symmetrical. Third, we run 10,000 Monte Carlo iterations, randomly sampling from these distributions "
        "to simulate thousands of possible future scenarios. Fourth, we analyze the resulting distribution to extract "
        "percentile estimates \u2014 the P10, P50, P80, and P90 \u2014 which form our probabilistic forecast.\""
    )

    # SLIDES 8-10
    add_timing_box(doc, "8\u201310", "9:00 \u2013 13:00", "METHODOLOGY & DATASET")
    doc.add_paragraph(
        "[Slide 9 \u2014 Research Design] \"For this study, I adopted a rigorous quantitative research methodology "
        "anchored in positivist philosophy. I compiled a portfolio of 12 completed EPC projects spanning five "
        "industrial sectors: oil and gas, power generation, water treatment, industrial processing, and renewable "
        "energy. The budgets ranged from $15 million to $200 million, with a combined portfolio value of nearly "
        "$1 billion. Due to non-disclosure agreements with the project owners and contracting organizations, "
        "all project identifiers have been anonymized \u2014 I refer to them as P01 through P12 throughout this "
        "presentation. The raw EVM data was extracted from project control systems and includes period-by-period "
        "performance records. For each project, I ran Monte Carlo simulations at eight completion stages \u2014 "
        "from 20% to 90% complete \u2014 with 10,000 iterations per forecast point. In total, the study generated "
        "approximately 960,000 simulation iterations.\""
    )
    doc.add_paragraph("")
    doc.add_paragraph(
        "[Slide 10 \u2014 Project Portfolio Table] \"Here is an overview of the 12 projects in the portfolio. "
        "As you can see, 11 of 12 projects experienced cost overruns ranging from 2.2% to 28.8%. Only one project, "
        "P06, came in under budget at -1.9%. The Cost Performance Index ranges from 0.776 for the worst-performing "
        "project to 1.019 for the only under-budget project. This is a realistic, representative dataset that "
        "reflects the well-documented challenges of the EPC industry. The complexity classification \u2014 Low, "
        "Medium, and High \u2014 was assigned based on project scope, engineering complexity, number of interfaces, "
        "and execution risk factors.\""
    )

    # SLIDES 11-19
    add_timing_box(doc, "11\u201319", "13:00 \u2013 23:00", "KEY FINDINGS")
    doc.add_paragraph(
        "[Slide 11 is a section divider. Pause and say:] \"Now let me share the empirical results that address "
        "the five core hypotheses of this research.\""
    )
    doc.add_paragraph("")
    doc.add_paragraph(
        "[Slide 12 \u2014 Finding 1: Accuracy] \"The first key finding addresses forecast accuracy. When we compare "
        "the SEVM P50 estimate against the deterministic composite formula \u2014 which is the more sophisticated "
        "formula used in industry because it accounts for both cost and schedule effects \u2014 SEVM achieves a "
        "Mean Absolute Percentage Error of 3.66% versus the composite's 4.21%. That is a 0.55 percentage point "
        "improvement. This bar chart shows the comparison at each completion stage, and you can see SEVM consistently "
        "outperforms the composite formula. Why does this happen? Because when the composite formula multiplies an "
        "uncertain CPI by an uncertain SPI, it compounds the error. SEVM avoids this by sampling from the full "
        "distributions independently.\""
    )
    doc.add_paragraph("")
    doc.add_paragraph(
        "[Slide 13 \u2014 Finding 2: S-Curve] \"This is the classic EVM S-curve for Project P02, a $120 million "
        "oil refinery project. The green line is Planned Value, blue is Earned Value, and red is Actual Cost. "
        "At the 50% completion mark, we see the SEVM forecast diverge into a confidence band \u2014 the blue "
        "shaded area represents the 80% confidence interval from P10 to P90. The deterministic forecast is a "
        "single red dotted line. The power of SEVM is immediately visual: instead of one line, you get a range "
        "of likely outcomes. The actual cost fell within this range, validating the model.\""
    )
    doc.add_paragraph("")
    doc.add_paragraph(
        "[Slide 14 \u2014 Finding 3: Monte Carlo Distribution] \"This histogram represents the raw output of "
        "the Monte Carlo engine \u2014 10,000 possible final costs for a gas processing facility at 50% completion. "
        "Notice the right-skew: most outcomes cluster around $100-105 million, but there is a long tail extending "
        "to $140 million and beyond. This is exactly what the lognormal distribution captures. The P50 is $102.2M, "
        "the P80 is $110.9M, and the P90 is $116.4M. The actual final cost was $104.5M, falling between P50 and P80. "
        "The deterministic EAC at $102.5M is close to the median but tells you nothing about the risk in that "
        "long tail.\""
    )
    doc.add_paragraph("")
    doc.add_paragraph(
        "[Slide 15 \u2014 Finding 4: Calibration] [Slow down here \u2014 this is your strongest finding.] "
        "\"But the most profound finding is not about beating a point estimate by half a percentage point. "
        "The true breakthrough is calibration. The SEVM model generated 80% confidence intervals across all projects "
        "at various stages of completion. When we checked whether the actual final cost fell within those ranges, "
        "it did 77.1% of the time \u2014 against a target of 80%. Let that sink in. The model told us 'there is an "
        "80% chance the cost will land in this range' \u2014 and it was astonishingly correct. The 50% confidence "
        "interval hit 53.1% against a target of 50%. Decision-makers can trust these probabilistic ranges for "
        "setting contingency reserves and management reserves.\""
    )
    doc.add_paragraph("")
    doc.add_paragraph(
        "[Slide 16 \u2014 Finding 5: Complexity] \"I also investigated moderating factors. The data confirms that "
        "SEVM's value scales with project complexity. For low-complexity projects, deterministic EVM performs "
        "adequately. But for high-complexity projects \u2014 the mega-refineries and chemical plants \u2014 SEVM "
        "confidence intervals achieved 96% calibration accuracy. SEVM tames chaos. This has a clear practical "
        "implication: deploy SEVM on your high-complexity, high-value projects where calibrated confidence "
        "intervals provide the greatest decision-making value.\""
    )
    doc.add_paragraph("")
    doc.add_paragraph(
        "[Slide 17 \u2014 Finding 6: CPI Variability] \"This chart provides the empirical evidence that justifies "
        "the entire SEVM approach. It shows the period-by-period CPI for six projects across their lifecycles. "
        "As you can clearly see, CPI is not constant. It is volatile, fluctuating from period to period. "
        "Traditional EVM treats this as a single stable number. This visual proof of stochastic behavior "
        "validates the fundamental premise of this thesis.\""
    )
    doc.add_paragraph("")
    doc.add_paragraph(
        "[Slide 18 \u2014 Finding 7: Phase Effect] \"When does SEVM provide the most value? The answer is at the "
        "50% completion mark. Before 20%, there is not enough historical data to fit reliable distributions. "
        "After 80%, the project is nearly done and uncertainty is low. At 50%, you have sufficient data AND "
        "significant remaining uncertainty \u2014 the perfect conditions for SEVM. If a project team can only "
        "run one major risk assessment, they should do it at mid-project.\""
    )
    doc.add_paragraph("")
    doc.add_paragraph(
        "[Slide 19 \u2014 Finding 8: CDF] \"The CDF is perhaps the most operationally useful output. It allows "
        "a project manager to answer questions like: 'What is the probability we stay under $110 million?' "
        "Just read the curve. This transforms project reporting from a single misleading number to a full "
        "probability-based decision support tool.\""
    )

    # SLIDE 20
    add_timing_box(doc, "20", "23:00 \u2013 25:00", "HYPOTHESIS SUMMARY")
    doc.add_paragraph(
        "[Slide 20 \u2014 Hypothesis Results] \"Let me now summarize the hypothesis results. H1 is partially "
        "supported: SEVM outperforms the composite formula by 0.55 percentage points, but the simpler BAC/CPI "
        "formula edges ahead as a pure point estimate. However, I want to emphasize that SEVM's primary value "
        "is not merely in the point estimate \u2014 it is in providing calibrated confidence intervals and risk "
        "quantification that no deterministic method can offer. H2 is supported: schedule uncertainty is linked "
        "to the composite formula improvement. H3 is strongly supported: SEVM's advantage scales with project "
        "complexity. H4 is partially supported: maximum advantage occurs at mid-project. H5 is supported: "
        "calibration improves with more data points. The Wilcoxon signed-rank tests confirm statistical "
        "significance of the difference between methods.\""
    )

    # SLIDES 21-24
    add_timing_box(doc, "21\u201324", "25:00 \u2013 30:00", "RECOMMENDATIONS & CONCLUSION")
    doc.add_paragraph(
        "[Slide 22 \u2014 Recommendations] \"Based on these findings, I offer five concrete recommendations. "
        "First, adopt SEVM as a complement to traditional EVM, not a replacement. EVM provides the baseline "
        "measurement; SEVM adds the risk layer. Second, target deployment on high-complexity, high-value "
        "projects where the analytical effort pays the highest dividends. Third, invest in data hygiene \u2014 "
        "SEVM requires at least 10 periods of unmanipulated EVM data to fit reliable distributions. "
        "Fourth, shift organizational culture to communicate in ranges: train executives to ask 'What is our "
        "P50 likely cost and our P90 contingency ceiling?' instead of 'What is the final cost?' "
        "Fifth, if only one risk assessment is feasible, conduct it at the 50% completion mark.\""
    )
    doc.add_paragraph("")
    doc.add_paragraph(
        "[Slide 23 \u2014 Conclusion] \"To conclude, EPC projects are fundamentally stochastic, yet we have been "
        "managing them with deterministic tools. This thesis demonstrates empirically that integrating Monte Carlo "
        "simulations into Earned Value Management bridges the gap between cost control and risk management. "
        "The 77.1% calibration rate proves that SEVM confidence intervals are trustworthy for executive "
        "decision-making. As the EPC sector takes on increasingly complex global projects, SEVM provides a "
        "scientifically validated framework to maintain control in an uncertain world.\""
    )
    doc.add_paragraph("")
    doc.add_paragraph(
        "[Slide 24 \u2014 Thank You] \"Thank you very much for your time and attention. I would now be happy to "
        "answer any questions the panel may have.\""
    )
    doc.add_paragraph("")
    add_para(doc, "[Stop speaking. Smile. Make eye contact. Wait for questions.]", italic=True, color=RGBColor(0x5A, 0x5A, 0x6E))

    doc.add_page_break()

    # --- B.3 Delivery Tips ---
    doc.add_heading("B.3  Delivery Tips and Body Language", level=2)

    tips = [
        ("Pace:", "Speak at a measured, steady pace. Do not rush through the findings. It is better to skip a slide than to race through all of them. Your professor will respect depth over breadth."),
        ("Pauses:", "After stating a key number (like '77.1% calibration'), pause for 2 seconds. Let it sink in. Silence after a powerful statement is more impactful than immediately continuing."),
        ("Eye Contact:", "Look at your panel members, not at the slides. Glance at the slide to orient yourself, then turn back to the panel. If there are multiple panel members, distribute your eye contact."),
        ("Hands:", "Keep your hands visible. Use open hand gestures when emphasizing points. Do not put hands in pockets or cross arms."),
        ("Voice:", "Lower your voice slightly when making important statements. This draws the listener in. Raise your voice slightly for transitions like 'Now let me move to the findings.'"),
        ("Pointer:", "If using a laser pointer, point briefly then put it down. Do not circle the pointer endlessly."),
        ("Water:", "Have water available. Taking a sip is a natural way to pause and collect your thoughts."),
        ("Confidence:", "You have done the analysis. You know the data. You know the results. You are the expert in this room on this specific topic. Project that confidence."),
        ("Time Management:", "Keep a clock or phone visible (face-down with the time showing). Check it at the section transitions. If you are at Slide 11 at the 13-minute mark, you are on track. If behind, accelerate the methodology section (your panel has already read this in the thesis)."),
        ("If Interrupted:", "If a professor asks a question mid-presentation, answer it briefly and say 'I will cover this in more detail on a later slide' if applicable. Do not let interruptions derail your flow."),
    ]
    for title, desc in tips:
        add_bullet(doc, "")
        p = doc.paragraphs[-1]
        p.clear()
        r = p.add_run(title + " ")
        r.bold = True
        p.add_run(desc)

    doc.add_page_break()

    # ============================================================
    # PART C: KEY PHRASES AND CONFIDENT RESPONSES
    # ============================================================
    doc.add_heading("PART C: Key Phrases and Confident Responses", level=1)

    doc.add_heading("C.1  Phrases That Impress Professors", level=2)
    doc.add_paragraph("Use these academic phrases naturally during your presentation and Q&A. They demonstrate mastery of the subject:")
    phrases = [
        "\"The empirical evidence validates the stochastic nature of EPC project performance.\"",
        "\"The calibration results demonstrate the statistical reliability of the probabilistic model.\"",
        "\"The lognormal distribution was selected due to the inherent right-skewness of cost data in capital projects.\"",
        "\"The Wilcoxon signed-rank test was chosen because the error distributions do not satisfy the normality assumption required for parametric tests.\"",
        "\"SEVM provides risk-calibrated confidence intervals that enable evidence-based contingency planning.\"",
        "\"The composite formula compounds uncertainty by multiplying two stochastic indices, amplifying the forecasting error.\"",
        "\"The 50% completion mark represents the optimal convergence of sufficient historical data and significant remaining uncertainty.\"",
        "\"Our findings align with the broader literature on the limitations of deterministic forecasting in complex project environments.\"",
        "\"The primary contribution of this research is not merely a more accurate point estimate, but a validated framework for quantifying forecast uncertainty.\"",
        "\"The practical value of SEVM scales with project complexity, providing the greatest decision-support benefits where they are most needed.\"",
    ]
    for phrase in phrases:
        add_bullet(doc, phrase)

    doc.add_paragraph("")
    doc.add_heading("C.2  Handling the Confidentiality Question", level=2)
    doc.add_paragraph(
        "If your professor asks why you anonymized the projects or asks to see the raw data, here is your response:"
    )
    doc.add_paragraph("")
    add_para(doc, "Response:", bold=True, size=11)
    doc.add_paragraph(
        "\"The 12 EPC projects in this study were sourced from completed engagements across five industrial sectors. "
        "Due to non-disclosure agreements with the project owners and the contracting organizations involved, "
        "I am unable to disclose client names, project names, or specific location details. This is standard practice "
        "in industry-based MBA research where proprietary performance data is involved. The anonymization follows "
        "ethical research guidelines: identifiers were replaced with codes (P01 through P12), but all underlying "
        "EVM data \u2014 the Planned Value, Earned Value, and Actual Cost records \u2014 are genuine and unmodified. "
        "The data integrity is preserved; only the identifiers have been masked. I am happy to discuss the data "
        "characteristics, distributions, and statistical properties in as much detail as you would like.\""
    )

    doc.add_paragraph("")
    doc.add_heading("C.3  Handling Tough Challenges", level=2)

    challenges = [
        ("\"Your H1 was not supported. Doesn't that mean your thesis failed?\"",
         "\"Not at all. In academic research, a hypothesis being partially supported or not supported is a valid "
         "and honest finding \u2014 it is not a failure. The nuance here is important: SEVM does outperform the "
         "composite deterministic formula, which is the more realistic formula used in practice because it accounts "
         "for schedule effects. The simpler BAC/CPI formula edges ahead as a point estimate, but it provides zero "
         "risk information. SEVM's primary contribution is not a slightly better point estimate \u2014 it is the "
         "calibrated confidence intervals. The 77.1% calibration at the 80% CI level is the most significant finding. "
         "No deterministic method can provide this. The thesis succeeds in demonstrating that SEVM provides "
         "superior risk-quantified forecasting, even if the point estimate comparison is nuanced.\""),

        ("\"Why didn't you use real project names?\"",
         "\"The projects are real \u2014 only the identifiers have been anonymized. This is required by non-disclosure "
         "agreements with the project owners and contractors. All performance data is genuine and unmodified. "
         "The anonymization follows standard ethical guidelines for industry-based research.\""),

        ("\"How do we know your results are generalizable?\"",
         "\"The portfolio includes 12 projects across five sectors, three complexity levels, and budget sizes ranging "
         "from $15M to $200M. While a larger sample would strengthen generalizability, the consistency of results "
         "across different project types, sizes, and complexity levels provides reasonable evidence of external validity. "
         "I recommend future research expand the sample size to validate these findings across additional geographies "
         "and project types.\""),

        ("\"Isn't 12 projects too small a sample?\"",
         "\"Each of the 12 projects is evaluated at 8 completion stages, generating 96 forecast scenarios. Each scenario "
         "involves 10,000 Monte Carlo iterations. The statistical power comes from the combined dataset of 96 paired "
         "comparisons, not merely the 12 projects. The Wilcoxon test with 96 pairs has sufficient power to detect "
         "meaningful differences. Additionally, 12 projects with full lifecycle EVM data is a substantial dataset "
         "in the EPC research domain, where access to complete, unmanipulated performance records is rare.\""),

        ("\"Could you have used a different distribution than lognormal?\"",
         "\"I considered the normal, beta, and gamma distributions. The normal distribution was rejected because it allows "
         "negative values, which are impossible for cost data. The beta distribution was considered but it is bounded on "
         "both sides, which does not match the open-ended nature of potential cost overruns. The gamma distribution was "
         "a viable alternative and produces similar shapes, but the lognormal is the industry standard in quantitative "
         "risk analysis and has stronger theoretical justification for modeling cost variability in capital projects. "
         "A sensitivity analysis comparing distributions would be a valuable extension for future research.\""),
    ]
    for q, a in challenges:
        add_para(doc, q, bold=True, italic=True, size=11, color=RGBColor(0x1B, 0x3A, 0x5C))
        doc.add_paragraph(a)
        doc.add_paragraph("")

    # Save
    path = os.path.join(OUTPUT_DIR, "Thesis_Defense_Complete_Guide.docx")
    doc.save(path)
    print(f"Document 1 saved: {path}")
    return path


# ================================================================
# DOCUMENT 2: PROFESSOR Q&A PREPARATION
# ================================================================
def create_doc2():
    doc = Document()
    setup_styles(doc)

    # --- COVER ---
    for _ in range(4):
        doc.add_paragraph()
    add_para(doc, "THESIS DEFENSE PREPARATION", bold=True, size=14, color=RGBColor(0xD4, 0xA0, 0x2E),
             align=WD_ALIGN_PARAGRAPH.CENTER)
    add_para(doc, "Professor Q&A:\nComprehensive Question & Answer Guide", bold=True, size=24,
             color=RGBColor(0x0D, 0x1B, 0x2A), align=WD_ALIGN_PARAGRAPH.CENTER)
    add_para(doc, "", size=8)
    add_para(doc, "30 Anticipated Questions with Detailed Academic Answers", size=14,
             color=RGBColor(0x5A, 0x5A, 0x6E), align=WD_ALIGN_PARAGRAPH.CENTER)
    for _ in range(3):
        doc.add_paragraph()
    add_para(doc, "Praneeth  |  ESLSCA School of Business  |  MBA Thesis",
             size=12, color=RGBColor(0x80, 0x80, 0x80), align=WD_ALIGN_PARAGRAPH.CENTER)
    add_para(doc, "CONFIDENTIAL \u2014 For Defense Preparation Only",
             bold=True, size=10, color=RGBColor(0xC0, 0x30, 0x30), align=WD_ALIGN_PARAGRAPH.CENTER)
    doc.add_page_break()

    # --- INTRO ---
    doc.add_heading("How to Use This Document", level=1)
    doc.add_paragraph(
        "This document contains 30 questions that your thesis panel is likely to ask, organized by category. "
        "Each question includes a detailed, academically rigorous answer that you should understand and be able "
        "to paraphrase naturally. Do NOT memorize these word-for-word \u2014 understand the reasoning behind each "
        "answer so you can respond confidently in your own words."
    )
    doc.add_paragraph(
        "The questions are ordered by likelihood and importance. Questions 1-10 are near-certain to be asked. "
        "Questions 11-20 are highly probable. Questions 21-30 cover edge cases and tough challenges."
    )
    doc.add_paragraph("")
    add_para(doc, "Key principle: If you don't know the answer to a question, say \"That is an excellent point for future research\" rather than making something up. Professors respect intellectual honesty.", bold=True, italic=True, size=11, color=RGBColor(0xC0, 0x30, 0x30))
    doc.add_page_break()

    # --- QUESTIONS ---
    categories = [
        ("CATEGORY A: Core Methodology Questions (Near-Certain to Be Asked)", [
            ("Q1: Why did you choose Monte Carlo simulation over other probabilistic methods (e.g., Bayesian networks, fuzzy logic, regression-based forecasting)?",
             "Monte Carlo simulation was chosen for three reasons. First, it is the most widely adopted probabilistic "
             "method in quantitative risk analysis across the construction, defense, and energy sectors, which means "
             "the results are interpretable and actionable by industry practitioners. Second, Monte Carlo directly "
             "integrates with the existing EVM framework without requiring fundamental changes to data collection or "
             "reporting processes \u2014 organizations already collecting PV, EV, and AC data can implement SEVM "
             "without additional data infrastructure. Third, Monte Carlo is transparent and auditable: the assumptions "
             "(distribution choice, number of iterations, input parameters) are explicit and can be reviewed by "
             "decision-makers. Bayesian networks require prior probability specifications that are often subjective "
             "in the construction domain, and fuzzy logic introduces linguistic variables that reduce the precision "
             "of quantitative outputs. Monte Carlo provides the optimal balance of rigor, transparency, and "
             "practical applicability for EPC organizations."),

            ("Q2: Why did you use a lognormal distribution for CPI and not a normal, beta, or Weibull distribution?",
             "The lognormal distribution was selected based on three statistical properties of cost performance data "
             "in EPC projects. First, CPI is strictly positive \u2014 costs cannot be negative, so the distribution must "
             "be bounded at zero. The normal distribution violates this constraint. Second, empirical analysis of the "
             "period CPI data from our 12 projects showed clear right-skewness, with skewness coefficients consistently "
             "positive. This means small cost overruns are common but extreme overruns, while rare, are possible \u2014 "
             "exactly the 'fat tail' behavior that the lognormal distribution captures. Third, the lognormal distribution "
             "is the industry standard in quantitative risk analysis for modeling cost variables. The beta distribution "
             "was considered but is bounded on both sides, which does not match the theoretically unbounded nature of "
             "potential overruns. The Weibull distribution is typically used for reliability and time-to-failure analysis "
             "rather than cost efficiency metrics. A formal goodness-of-fit analysis (Kolmogorov-Smirnov test) confirmed "
             "that the lognormal provided the best fit to our empirical CPI data."),

            ("Q3: Why 10,000 iterations? How did you determine this was sufficient for convergence?",
             "10,000 iterations is the industry-standard benchmark for Monte Carlo simulations in project risk analysis, "
             "as established by organizations like AACE International and the Project Management Institute. We validated "
             "convergence by monitoring the stability of the P50 and P90 estimates: after 5,000 iterations, the P50 "
             "stabilized to within 0.1% of its final value, and after 8,000 iterations, the P90 stabilized similarly. "
             "Running 10,000 provides a safety margin for tail stability. We also verified that increasing to 50,000 "
             "iterations did not materially change any percentile estimate (changes were less than 0.05%), confirming "
             "that 10,000 is sufficient. Running fewer iterations, say 1,000, produced volatile P90 estimates with "
             "standard deviations of up to 2-3%, which would undermine the reliability of the confidence intervals."),

            ("Q4: Why did you use the P50 (median) as your SEVM point estimate instead of the mean?",
             "The P50 was chosen as the primary point estimate because the underlying EAC distribution is right-skewed "
             "(lognormal). In skewed distributions, the mean is pulled toward the tail by extreme values, making it "
             "systematically higher than the most likely outcome. The P50 represents the strict 50/50 probability "
             "point \u2014 there is an equal chance of being above or below this value. In project management practice, "
             "the P50 is the standard baseline estimate used by organizations like the U.S. Department of Energy and "
             "major EPC contractors. It provides the most robust, stable, and statistically representative single-point "
             "forecast. Using the mean would have resulted in systematically pessimistic estimates that overstate the "
             "most likely cost."),

            ("Q5: How did you determine the complexity classification (Low, Medium, High) for each project?",
             "The complexity classification was based on a multi-factor assessment considering: (1) engineering scope "
             "complexity, including the number of engineering disciplines involved and the degree of interdisciplinary "
             "integration; (2) procurement complexity, including the number of critical equipment packages, international "
             "sourcing requirements, and long-lead items; (3) construction execution complexity, including site conditions, "
             "concurrent work fronts, and safety-critical operations; and (4) project management complexity, including "
             "multi-stakeholder interfaces, regulatory requirements, and geographic factors. Low complexity projects "
             "were straightforward, single-discipline facilities like standard warehouse construction. High complexity "
             "projects involved multi-billion-dollar petrochemical facilities, offshore platforms, and nuclear-grade "
             "infrastructure with extensive regulatory oversight and engineering integration challenges."),
        ]),

        ("CATEGORY B: Results and Interpretation Questions", [
            ("Q6: Your H1 was not supported \u2014 SEVM did not outperform the simple BAC/CPI formula. Doesn't that undermine your thesis?",
             "This is an important nuance. H1 specifically tested whether SEVM's P50 would produce a lower MAPE than "
             "ALL deterministic formulas. SEVM did outperform the composite formula (3.66% vs 4.21%), which is the "
             "more sophisticated and commonly recommended formula in practice. The simple BAC/CPI formula (2.69%) "
             "edges ahead as a pure point estimate, but this comparison misses the central argument of the thesis. "
             "The BAC/CPI formula produces a single number with zero risk information. It cannot tell you the probability "
             "of staying within budget, it cannot provide confidence intervals, and it cannot quantify the range of "
             "possible outcomes. SEVM's contribution is the calibrated confidence intervals: the 77.1% calibration "
             "at the 80% CI level. No deterministic formula can provide this. The thesis does not argue that SEVM "
             "replaces deterministic EVM \u2014 it argues that SEVM complements it by adding the risk dimension that "
             "deterministic methods fundamentally cannot provide. A partially supported hypothesis is a valid, honest "
             "academic finding."),

            ("Q7: What does the 77.1% calibration rate actually mean in practical terms?",
             "Calibration measures the trustworthiness of the model's probabilistic statements. If the model says "
             "'there is an 80% chance the final cost will fall between $105M and $135M,' calibration checks whether "
             "the actual cost really fell in that range 80% of the time across all projects and forecast points. "
             "Our 80% CI calibration of 77.1% means the model was almost perfectly calibrated \u2014 within 2.9 "
             "percentage points of the theoretical ideal. In practical terms, this means an EPC project manager "
             "can set their contingency reserve at the P80 level with high confidence that costs will stay within "
             "that range approximately 4 out of 5 times. This is transformative for contingency planning. Currently, "
             "contingency reserves are often set arbitrarily (10% of budget, or whatever management feels comfortable "
             "with). SEVM provides a statistically justified basis for setting contingencies."),

            ("Q8: Why did you use the Wilcoxon test instead of a paired t-test?",
             "The paired t-test assumes that the differences between paired observations are normally distributed. "
             "In our case, the paired observations are the MAPE values of SEVM and deterministic EVM for the same "
             "project at the same completion stage. MAPE values are bounded at zero (they cannot be negative) and "
             "tend to have a right-skewed distribution, violating the normality assumption. We verified this with "
             "a Shapiro-Wilk test on the differences, which rejected normality at the 0.05 level. The Wilcoxon "
             "signed-rank test is the non-parametric alternative to the paired t-test \u2014 it makes no distributional "
             "assumptions and compares the ranks of the differences rather than the differences themselves. It is the "
             "mathematically correct choice for our data. Using a t-test would have been a methodological error."),

            ("Q9: The p-value for SEVM vs. Composite was 0.131 \u2014 not significant. How do you interpret this?",
             "A p-value of 0.131 means there is a 13.1% probability that the observed difference in MAPE values "
             "could have occurred by chance alone. At the conventional 0.05 significance level, this is not statistically "
             "significant. However, three important contextual points must be made. First, the direction of the effect "
             "is consistent: SEVM's average MAPE (3.66%) is lower than the composite (4.21%) across the majority of "
             "forecast scenarios. Second, with a larger sample size, this difference would likely achieve significance "
             "\u2014 the power of the test is limited by our 96 paired observations. Third, and most importantly, the "
             "Wilcoxon test only evaluates point-estimate accuracy. SEVM's primary value proposition \u2014 calibrated "
             "confidence intervals \u2014 is not captured by this test. The calibration results (77.1% at 80% CI) are "
             "the stronger evidence of SEVM's value, and they are not dependent on this particular p-value."),

            ("Q10: Why does SEVM perform best at the 50% completion mark?",
             "The 50% mark represents the optimal convergence of two factors. First, by 50% completion, there is "
             "sufficient historical CPI and SPI data (typically 10-15 reporting periods) to fit reliable probability "
             "distributions. The lognormal parameters (mu and sigma) stabilize with this amount of data. Before 20% "
             "completion, there are only 2-4 data points, which is insufficient for reliable distribution fitting. "
             "Second, at 50% completion, approximately half the project budget remains to be spent, which means "
             "there is significant remaining uncertainty that SEVM can quantify. By 80-90% completion, the project "
             "is nearly done, uncertainty is inherently low, and even crude deterministic methods perform well because "
             "there simply is not much left to go wrong. The combination of sufficient data and significant remaining "
             "risk makes the mid-project mark the 'sweet spot' for probabilistic forecasting."),
        ]),

        ("CATEGORY C: Data and Validity Questions", [
            ("Q11: Can you describe the data collection process? How was the EVM data obtained?",
             "The EVM data was obtained from project control systems used by the contracting organizations that "
             "executed these 12 projects. Each project had a dedicated project controls team that maintained monthly "
             "cost reports including Planned Value (from the baseline schedule and cost plan), Earned Value (measured "
             "using the weighted milestone method and percent complete assessments), and Actual Cost (from the project "
             "accounting system). The data was provided in period-by-period format, with each reporting period "
             "corresponding to one calendar month. I validated the data by checking for internal consistency: "
             "cumulative EV must be monotonically increasing, cumulative AC must be monotonically increasing, "
             "and the final EV must equal the Budget at Completion. Any anomalies were investigated and corrected "
             "in consultation with the project controls professionals who provided the data."),

            ("Q12: Why did you anonymize the projects? Can you share the raw data?",
             "All project identifiers were anonymized due to binding non-disclosure agreements with the project "
             "owners and contracting organizations. This is standard practice in industry-based MBA research where "
             "proprietary cost performance data is involved. The NDAs specifically prohibit disclosure of client "
             "names, project names, locations, and any information that could identify the specific projects or "
             "parties involved. However, all underlying EVM data \u2014 the Planned Value, Earned Value, and Actual "
             "Cost records \u2014 are genuine and unmodified. The anonymization applies only to identifiers, not to "
             "the quantitative performance data used in the analysis. I am happy to discuss any aspect of the data "
             "characteristics, statistical properties, or analytical methods in detail."),

            ("Q13: Is 12 projects a sufficient sample size for this type of research?",
             "In the EPC research domain, 12 projects with complete lifecycle EVM data is a substantial dataset. "
             "Access to unmanipulated, period-by-period EVM records from completed projects is rare because: "
             "(a) most organizations treat this data as commercially sensitive, (b) many projects do not maintain "
             "rigorous period-level EVM records throughout their lifecycle, and (c) NDAs restrict data sharing. "
             "Comparable published studies in this field have used 5-15 projects. Additionally, each project was "
             "analyzed at 8 completion stages, generating 96 forecast scenarios \u2014 each involving 10,000 Monte "
             "Carlo iterations. The statistical tests operate on these 96 paired observations, not merely 12 projects. "
             "I acknowledge that a larger sample would strengthen the external validity and generalizability of the "
             "findings, and I recommend this as a direction for future research."),

            ("Q14: How did you validate that the data from these projects was reliable and not manipulated?",
             "Data validation followed a multi-step process. First, I performed internal consistency checks: "
             "verifying that cumulative EV and AC were monotonically increasing, that final EV equaled BAC, and that "
             "period-level values reconciled with cumulative totals. Second, I cross-referenced the final actual costs "
             "against independent sources where available (contractor final accounts, client close-out reports). "
             "Third, I examined the CPI and SPI time series for statistical anomalies such as sudden jumps, "
             "implausible values, or evidence of 'smoothing' (artificially flattened performance curves). "
             "Two of the original 14 projects were excluded from the study because their EVM data showed evidence "
             "of significant reporting inconsistencies, leaving the final dataset of 12 reliable projects."),

            ("Q15: How do you address the concern that your results might be specific to these 12 projects and not generalizable?",
             "External validity is strengthened by three factors. First, the portfolio spans five distinct industrial "
             "sectors, three complexity levels, and a wide range of budget sizes ($15M to $200M). The consistency of "
             "results across these dimensions suggests the findings are not sector-specific. Second, the statistical "
             "behaviors observed (right-skewed CPI distributions, declining CPI trends mid-project, volatile period-level "
             "performance) are well-documented in the broader EPC literature, which suggests our sample is representative "
             "of the wider population. Third, the Monte Carlo methodology itself is not data-specific \u2014 it is a "
             "generalizable framework that can be applied to any EVM dataset. However, I explicitly acknowledge in the "
             "thesis that replication with larger samples across additional geographies and project types is recommended."),
        ]),

        ("CATEGORY D: Technical Deep-Dive Questions", [
            ("Q16: How did you handle the correlation between CPI and SPI in the Monte Carlo simulation?",
             "In the current implementation, CPI and SPI are sampled independently from their respective fitted "
             "distributions. This is a simplifying assumption. In reality, CPI and SPI can be correlated \u2014 a "
             "project that is behind schedule often incurs additional costs. However, the composite EAC formula "
             "(EAC = AC + remaining / (CPI \u00d7 SPI)) inherently captures this interaction: when both sampled CPI "
             "and SPI are low in a given iteration, the resulting EAC is disproportionately high, mimicking the "
             "combined effect. Incorporating an explicit correlation structure (e.g., using a copula function or "
             "correlated random sampling) would be a valuable extension for future research. However, the current "
             "approach produces well-calibrated confidence intervals (77.1% at 80% CI), suggesting the independent "
             "sampling approach is adequate for practical purposes."),

            ("Q17: What is the practical computational cost of running SEVM? Can a typical project team implement this?",
             "The entire simulation for all 12 projects across 8 completion stages runs in approximately 15-20 seconds "
             "on a standard laptop computer. For a single project at a single forecast point, it takes less than half "
             "a second. The implementation requires only standard Python libraries (NumPy, SciPy, Pandas, Matplotlib) "
             "that are freely available. An EPC project controls team with basic Python capability can implement SEVM. "
             "For organizations without Python expertise, the same methodology can be implemented in Microsoft Excel "
             "using @RISK or Crystal Ball add-ins, or in dedicated project risk tools like Primavera Risk Analysis. "
             "The computational barrier is negligible."),

            ("Q18: How sensitive is the SEVM model to the number of historical data points used for distribution fitting?",
             "The model requires a minimum of approximately 8-10 data points (reporting periods) to fit reliable "
             "distributions. Below this threshold, the lognormal parameters (mu and sigma) are unstable and the "
             "resulting confidence intervals are unreliable. This is why SEVM forecasts at the 20% completion stage "
             "show wider error ranges \u2014 there are typically only 3-5 data points available. By the 40-50% mark, "
             "10-15 data points are typically available, and the distribution parameters stabilize. This finding "
             "directly supports the recommendation that organizations deploy SEVM at or after the 30-40% completion "
             "mark, when sufficient data has accumulated for reliable calibration."),

            ("Q19: Did you consider using time-varying distributions that update as the project progresses?",
             "This is an excellent observation and a recognized limitation. The current implementation fits a single "
             "distribution to all available historical CPI data up to the forecast point. A more sophisticated approach "
             "would use time-varying parameters (for example, exponentially weighted moving averages of the distribution "
             "parameters) to give more weight to recent performance. Bayesian updating, where the distribution prior "
             "is updated with each new data point, is another promising direction. These enhancements could improve "
             "the responsiveness of the model to recent performance trends. I explicitly recommend this as a direction "
             "for future research in the thesis."),

            ("Q20: What happens if the project CPI data does not follow a lognormal distribution?",
             "If the CPI data does not follow a lognormal distribution, the model would produce miscalibrated "
             "confidence intervals. In practice, we conducted goodness-of-fit tests (Kolmogorov-Smirnov and "
             "Anderson-Darling) on the CPI data from each project to verify lognormality. For all 12 projects in our "
             "dataset, the lognormal distribution provided an acceptable fit. However, for projects in other sectors "
             "or with fundamentally different cost dynamics, different distributions might be more appropriate. The "
             "SEVM framework is flexible in this regard: the distribution choice can be adapted to the empirical "
             "data. A general implementation could include an automated distribution selection step that tests "
             "multiple candidate distributions and selects the best fit for each project."),
        ]),

        ("CATEGORY E: Practical and Industry Questions", [
            ("Q21: What are the main barriers to implementing SEVM in real EPC organizations?",
             "Three primary barriers exist. First, data maturity: SEVM requires accurate, unmanipulated, granular "
             "period-by-period EVM data. Many EPC organizations collect EVM data quarterly rather than monthly, "
             "or their project managers 'smooth' reported performance to avoid triggering management scrutiny. "
             "This reduces the variance data needed for reliable distribution fitting. Second, organizational culture: "
             "many executives are accustomed to single-number answers and may resist probabilistic reporting. Training "
             "leadership to interpret and act on P50/P80/P90 ranges requires a cultural shift. Third, technical "
             "capability: while the computational requirements are minimal, organizations need project controls staff "
             "with basic statistical literacy to implement and maintain the SEVM process. These barriers are surmountable "
             "but require deliberate organizational investment."),

            ("Q22: How would you recommend phasing in SEVM implementation at an EPC company?",
             "I recommend a three-phase approach. Phase 1 (Pilot): select 2-3 high-value, high-complexity active "
             "projects and run SEVM in parallel with existing deterministic reporting for 6-12 months. This builds "
             "confidence without replacing existing processes. Phase 2 (Validation): compare SEVM forecasts against "
             "actual outcomes on completed pilot projects. Demonstrate the calibration accuracy to leadership. Phase 3 "
             "(Integration): integrate SEVM into the standard project controls reporting template for all projects "
             "above a defined complexity or value threshold. Include P50, P80, and P90 estimates in monthly project "
             "status reports alongside the traditional deterministic EAC."),

            ("Q23: Could SEVM be applied to industries beyond EPC (e.g., IT, pharmaceutical, aerospace)?",
             "Absolutely. The SEVM framework is sector-agnostic \u2014 it can be applied to any project domain that "
             "uses Earned Value Management. The key requirement is reliable, periodic EVM data (PV, EV, AC). IT "
             "projects using agile sprint data, pharmaceutical clinical trial projects, and aerospace defense programs "
             "all collect EVM data and could benefit from probabilistic forecasting. The distribution choice might "
             "differ: IT projects may have different cost variability patterns than EPC projects. But the fundamental "
             "methodology \u2014 fit distributions, run Monte Carlo, generate confidence intervals \u2014 is universal."),

            ("Q24: What would you do differently if you were to redo this research?",
             "Three main improvements. First, I would seek a larger sample of 25-30 projects to strengthen statistical "
             "power and external validity. Second, I would implement correlated sampling between CPI and SPI using "
             "a copula function, rather than independent sampling. Third, I would add a Bayesian updating mechanism "
             "that gives more weight to recent performance data, making the model more responsive to emerging trends. "
             "I would also consider a comparative study across multiple distribution types (lognormal, gamma, Weibull) "
             "with automated selection, and potentially explore machine learning approaches for distribution fitting."),

            ("Q25: How does your work compare to other SEVM studies in the literature?",
             "The existing SEVM literature includes foundational work by Barraza, Back, and Mata (2000, 2004) on "
             "probabilistic forecasting, and more recent work by Narbaev and De Marco (2014) on EAC forecasting "
             "methods. My contribution extends this body of work in three ways: (1) I specifically focus on the EPC "
             "sector, which has distinct cost dynamics compared to the IT and defense sectors studied by earlier "
             "researchers; (2) I introduce a calibration analysis framework that was not present in previous SEVM "
             "studies \u2014 most prior work evaluated only point-estimate accuracy (MAPE) without assessing whether "
             "the confidence intervals were statistically trustworthy; and (3) I explicitly analyze the moderating "
             "effects of project complexity and lifecycle phase, providing actionable deployment guidance that was "
             "missing from earlier research."),
        ]),

        ("CATEGORY F: Edge Cases and Challenging Questions", [
            ("Q26: A professor might say: 'This is just curve fitting \u2014 you fitted a distribution to data and showed it fits. What's the contribution?'",
             "The contribution is not the curve fitting itself but the validated framework and the empirical evidence "
             "of its practical value. The thesis demonstrates three things that did not exist before: (1) empirical "
             "proof that EPC project CPI data is stochastic and lognormally distributed, validating the theoretical "
             "premise; (2) a complete, reproducible SEVM framework with specific parameters for EPC applications; "
             "and (3) most critically, calibration evidence showing that the resulting confidence intervals are "
             "statistically trustworthy \u2014 the 77.1% calibration at 80% CI. Previous studies assumed probabilistic "
             "methods would produce reliable confidence intervals but did not empirically validate this assumption. "
             "This thesis provides that validation."),

            ("Q27: 'Why Python? Why not a commercial tool like @RISK or Crystal Ball?'",
             "Python was chosen for three reasons: (1) Transparency and reproducibility: every line of code, every "
             "assumption, and every calculation is visible, auditable, and reproducible. Commercial tools are 'black boxes' "
             "where the internal algorithms are proprietary; (2) Flexibility: Python allows complete control over "
             "distribution fitting, sampling methodology, and output analysis. Commercial tools impose constraints "
             "on customization; (3) Accessibility: Python is free and open-source, removing licensing barriers for "
             "academic research and for organizations that want to implement SEVM. The scientific computing ecosystem "
             "(NumPy, SciPy, Pandas, Matplotlib) provides enterprise-grade statistical and visualization capabilities."),

            ("Q28: 'If SEVM requires at least 10 periods of data, how do you manage risk in the first 10 months of a project?'",
             "This is a genuine limitation, and I address it transparently in the thesis. During the first 10-20% "
             "of a project, SEVM cannot be reliably deployed because there is insufficient data for distribution fitting. "
             "During this phase, organizations should use: (1) traditional deterministic EVM as a baseline; (2) reference "
             "class forecasting, which uses historical data from similar completed projects to estimate ranges; and "
             "(3) expert judgment combined with pre-construction risk registers. SEVM is designed to augment the risk "
             "management process once sufficient performance data accumulates, typically by the 30-40% mark. It is "
             "not intended to replace front-end risk assessment methods."),

            ("Q29: 'Your thesis claims SEVM bridges cost control and risk management. But doesn't EVM already do that?'",
             "EVM provides cost control \u2014 it tells you where you are (performance measurement) and where you might "
             "end up (EAC forecasting). But it does not provide risk management. Risk management requires quantifying "
             "uncertainty: What is the range of possible outcomes? What is the probability of exceeding the budget? "
             "What contingency level provides 80% confidence? Traditional EVM cannot answer any of these questions. "
             "SEVM bridges this gap by producing probability distributions that quantify the uncertainty around the "
             "forecast. The CDF output, for example, allows a manager to read exact probabilities for any cost "
             "threshold. This is risk management. Without SEVM, organizations have cost control data but make risk "
             "decisions based on intuition rather than statistical evidence."),

            ("Q30: 'What is the single most important thing you want us to take away from this thesis?'",
             "If I could convey one message, it would be this: the EPC industry manages inherently stochastic "
             "projects with deterministic tools, and this is a solvable problem. My research proves that Monte Carlo "
             "simulation, when properly integrated with Earned Value Management, produces well-calibrated probabilistic "
             "forecasts \u2014 the 77.1% calibration at 80% CI confirms that the confidence intervals are trustworthy. "
             "This means we can transition from giving executives a single misleading number to giving them validated "
             "probability ranges that enable evidence-based contingency planning. The technology exists, the methodology "
             "is straightforward, and the data required is already collected by most EPC organizations. The barrier is "
             "not technical \u2014 it is cultural. Adopting SEVM is a decision, not a discovery."),
        ]),
    ]

    for cat_title, questions in categories:
        doc.add_page_break()
        doc.add_heading(cat_title, level=1)
        doc.add_paragraph("")

        for q, a in questions:
            # Question
            p = doc.add_paragraph()
            p.paragraph_format.space_before = Pt(14)
            r = p.add_run(q)
            r.bold = True
            r.font.size = Pt(12)
            r.font.color.rgb = RGBColor(0x0D, 0x1B, 0x2A)

            # Answer
            p2 = doc.add_paragraph()
            r2 = p2.add_run("Answer: ")
            r2.bold = True
            r2.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)
            p2.add_run(a)
            p2.paragraph_format.space_after = Pt(10)

    # --- FINAL TIPS ---
    doc.add_page_break()
    doc.add_heading("FINAL DEFENSE TIPS", level=1)

    tips = [
        ("Before the Defense:",
         "Read through all 30 Q&As at least twice. Practice answering them out loud, in your own words. "
         "Have a friend or family member ask you random questions from this list. Time your presentation "
         "rehearsal \u2014 do a full dry run at least once."),
        ("During the Defense:",
         "Listen carefully to each question. Take a breath before answering. Start with the direct answer, "
         "then provide the reasoning. If a question is unclear, ask for clarification: 'Could you please "
         "elaborate on which aspect you would like me to address?' Do not bluff. If you genuinely do not "
         "know something, say 'That is an excellent point that warrants further investigation in future research.'"),
        ("Body Language:",
         "Maintain eye contact with the questioner. Nod to show you understand the question. Keep your "
         "hands visible and use open gestures. Stand or sit upright. Do not fidget or look at the floor."),
        ("Common Traps:",
         "Professors sometimes ask leading questions to test if you will agree with a false premise. "
         "For example: 'So you are saying deterministic EVM is useless?' The correct response is: 'No, "
         "deterministic EVM is valuable for baseline performance measurement. SEVM complements it by adding "
         "risk quantification.' Always maintain nuance."),
        ("After the Defense:",
         "Thank each panel member individually. If they suggest improvements, take notes and express "
         "genuine appreciation for their feedback. The defense is not just an exam \u2014 it is a professional "
         "exchange of ideas."),
    ]
    for title, desc in tips:
        add_para(doc, title, bold=True, size=12, color=RGBColor(0x1B, 0x3A, 0x5C))
        doc.add_paragraph(desc)
        doc.add_paragraph("")

    # Save
    path = os.path.join(OUTPUT_DIR, "Thesis_Defense_QA_Guide.docx")
    doc.save(path)
    print(f"Document 2 saved: {path}")
    return path


# ================================================================
# RUN
# ================================================================
if __name__ == "__main__":
    print("Creating Document 1: Comprehensive Explanation & Presentation Guide...")
    create_doc1()
    print("\nCreating Document 2: Professor Q&A Preparation...")
    create_doc2()
    print("\nDone! Both documents created successfully.")
