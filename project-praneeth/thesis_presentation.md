# Thesis Presentation Guide

## Project Summary: What is this project?
You are presenting your **MBA Thesis** submitted to the ESLSCA School of Business. 
**Title:** "Stochastic Earned Value Management: Integrating Monte Carlo Simulation for Enhanced Probabilistic Cost and Schedule Forecasting in EPC Projects."

**Core Concept:** The Engineering, Procurement, and Construction (EPC) sector frequently suffers from cost overruns and delays because traditional Earned Value Management (EVM) relies on "deterministic" (single-point) estimates that ignore real-world risks. Your project proposes using **Stochastic EVM (SEVM)**—which uses Monte Carlo Simulations to predict a range of probable outcomes (Confidence Intervals) rather than a single fixed number. You validated this by running 960,000 Monte Carlo simulation iterations on data from 12 real-world EPC projects (totaling $942 million in budget). 

***

## 20-Minute Presentation Speech Script

*(Note for Praneeth: Speak at a steady, engaging pace. Pause for emphasis at bolded points. A 20-minute speech should be roughly 2,500 words, but this script is structured efficiently to allow for slide transitions, breathing room, and natural cadence.)*

### 1. Introduction (Approx. 2 minutes)

**[Slide: Title Slide]**
"Good morning, esteemed panel members, my supervisor, and guests. My name is Praneeth, and I am honored to present my MBA thesis today, titled *'Stochastic Earned Value Management: Integrating Monte Carlo Simulation for Enhanced Probabilistic Cost and Schedule Forecasting in EPC Projects.'* 

Before I begin, I want to express my deepest gratitude to my supervisor, the faculty at ESLSCA School of Business, and my family for their unwavering support throughout this research journey."

**[Slide: The Reality of the EPC Industry]**
"To understand why this research matters, we simply need to look at the current state of the Engineering, Procurement, and Construction—or EPC—industry. EPC organizations build our world's critical infrastructure: power plants, refineries, and mega-projects. These are vast, capital-intensive undertakings often costing hundreds of millions of dollars. 

But there is a glaring problem: The industry is plagued by failure. Industry benchmarks show that up to 65% of large EPC projects experience cost overruns exceeding 25% of their initial budgets. Schedule delays of up to 80% are shockingly common. When we fail to predict costs accurately, the consequences are severe—wiped-out profit margins, contractual disputes, and even abandoned infrastructure.

My thesis asks a foundational question: *Why do we keep getting it so wrong, and how can we fix it?*"

### 2. The Problem Statement & Objectives (Approx. 4 minutes)

**[Slide: Deterministic EVM vs. The Real World]**
"Currently, the global standard for project control is Earned Value Management, or EVM. EVM is fantastic for measuring past performance. However, when it comes to forecasting the future, traditional EVM is tragically flawed. 

Traditional EVM forecasting uses deterministic formulas. It assumes that past performance will linearly dictate future performance. If your Cost Performance Index (CPI) to date is 0.85, traditional EVM assumes it will be exactly 0.85 until the project ends. It produces a **single-point estimate**—say, an Estimate at Completion of $150 million. 

But as project managers, we know the real world is not linear. It is volatile. It is subject to supply chain bottlenecks, labor shortages, and commodity price spikes. Giving a decision-maker a single, static number creates a false sense of certainty. It provides no information about the *likelihood* of hitting that target, or the worst-case scenario. 

**[Slide: Research Aims]**
"This brings us to the core aim of my research: to develop and validate a **Stochastic Earned Value Management (SEVM)** framework. 

Instead of treating project performance as a fixed constant, SEVM treats it as a *stochastic variable*. By integrating Monte Carlo Simulation, we generate a probability distribution of possible outcomes. We move from saying, *'The project will cost $150 million,'* to saying, *'There is an 80% probability the project will cost between $140 and $160 million.'*

My specific objectives were to:
1. Critically evaluate traditional EVM's limitations.
2. Develop a Monte Carlo-based EVM simulation framework.
3. Compare the forecasting accuracy of traditional EVM against SEVM using real-world EPC data.
4. Provide concrete, actionable recommendations for industry adoption."

### 3. Research Methodology (Approx. 4 minutes)

**[Slide: Methodology Overview]**
"To conduct this study, I adopted a rigorous quantitative research approach anchored in positivist philosophy. I wanted empirical proof, not just theoretical models.

I compiled a portfolio of **12 completed EPC projects** spanning five industrial sectors—including oil and gas, renewable energy, and water treatment. The budgets ranged from $15 million to $200 million, with a combined portfolio value of nearly $1 billion. Eleven of these twelve projects experienced actual cost overruns, providing a highly realistic dataset.

**[Slide: The Simulation Engine]**
"For the analysis, I built an SEVM engine using Python. The procedure followed four stages:
1. First, I took the historical EVM data (Planned Value, Earned Value, Actual Cost) over the lifecycle of these 12 projects.
2. Second, I fitted probability distributions—primarily lognormal distributions—to the Cost Performance Index (CPI), because cost surprises represent right-skewed data.
3. Third, I ran Monte Carlo Simulations at different completion phases—from 10% to 90% complete. For every single forecast scenario, I ran 10,000 iterations to ensure convergence. In total, the study generated nearly 1 million simulation iterations.
4. Finally, I compared the accuracy of the SEVM forecasts (specifically looking at the P50 median estimate and the confidence intervals) against traditional deterministic EVM formulas."

### 4. Key Findings (Approx. 6 minutes)

**[Slide: Core Finding 1: The Accuracy Matrix]**
"Let me share the empirical results, which address my five core hypotheses.

The first major finding is that SEVM significantly outperforms the complex deterministic EVM formula. When we look at the composite EVM formula—which factors in both schedule and cost—the traditional method had an average error, or MAPE, of 4.21%. Our SEVM approach brought that error down to 3.66%. 

Why does this happen? Because when traditional EVM multiplies an uncertain schedule index by an uncertain cost index, it amplifies the error. SEVM, by sampling from the full distributions, produces forecasts that are mathematically proven by Wilcoxon signed-rank tests to be more robust to combined uncertainty."

**[Slide: Core Finding 2: The Power of Calibration]**
"But the most profound value of SEVM is not merely beating a point estimate by a fraction of a percent. The true breakthrough is **Calibration**.

The SEVM model generated an 80% confidence interval for these projects at varying stages of completion. When looking at the final, actual project costs, the actual cost fell within our 80% confidence interval exactly 77.1% of the time. 

Let that sink in. The model told us 'there is an 80% chance the cost will land in this range'—and it was astonishingly correct. Decision-makers can trust these probabilistic ranges to accurately dictate their contingency buffers and management reserves."

**[Slide: Core Finding 3: Complexity and Phasing (H3 & H4)]**
"I also investigated moderating factors: Project Complexity and Project Phase. 

The data solidly confirms Hypothesis 3: SEVM's value scales with project complexity. For 'Low Complexity' projects, deterministic EVM actually performed quite well, and SEVM’s confidence intervals were less reliable. But for 'High Complexity' projects—the mega-refineries and chemical plants—SEVM confidence intervals exhibited a massive 96% calibration accuracy. SEVM tames chaos.

Regarding project phases, I found that SEVM provides its maximum accuracy advantage right at the 50% completion mark. Why? Because before 20%, you don't have enough historical data to fit a reliable statistical distribution. By 80%, the project is almost over, and uncertainty is inherently low. It is precisely in the murky middle—at 50%—where sufficient data meets massive remaining uncertainty, and where SEVM provides a critical, reliable compass."

### 5. Managerial Implications & Recommendations (Approx. 3 minutes)

**[Slide: Practical Recommendations for the Industry]**
"Based on these empirical findings, my thesis offers five concrete recommendations for EPC organizations transitioning to the 2025 digital era:

1. **Adopt SEVM as a Complement, Not a Replacement.** Traditional EVM provides a great baseline, but SEVM should be layered on top for risk management.
2. **Target Deployment.** Don’t use Monte Carlo on a simple $2 million warehouse. Deploy it on your high-complexity, high-risk, high-value projects where calibration pays dividends.
3. **Invest in Data Hygiene.** A Monte Carlo simulation is only as good as its inputs. My sensitivity analysis proved that SEVM requires at least 10 periods (e.g., 10 months) of high-quality EVM data to calibrate properly. 
4. **Communicate in Ranges.** We must train executives to stop asking 'what is the final cost?' and start asking, 'what is our P50 likely cost, and what is our P90 contingency ceiling?'.
5. **Focus Resources at Mid-Project.** If you only run one major risk assessment, do it at the 50% mark."

### 6. Conclusion (Approx. 1 minute)

**[Slide: Conclusion & Q&A]**
"To conclude, EPC projects are fundamentally stochastic, yet we have been managing them with deterministic tools. 

This thesis demonstrates empirically that integrating Monte Carlo simulations into Earned Value Management bridges the fatal gap between cost control and risk management. It transitions project reporting from a misleading single-point estimate to a validated, probabilistic confidence level. As the EPC sector takes on increasingly complex global projects, SEVM provides a scientifically proven framework to maintain control in an uncertain world.

Thank you very much for your time and attention. I would now be happy to answer any questions the panel may have."
