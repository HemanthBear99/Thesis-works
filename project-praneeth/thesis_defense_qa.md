# Thesis Defense: MBA Professor Q&A Guide

When you submit your thesis and codebase, your professor/panel will likely grill you on **why** you made specific mathematical and programmatic choices. They want to ensure you didn't just write a script, but understand the statistical reasoning behind it. 

Here are 10 highly likely questions they will ask about your `sevm_simulation.py` script, along with the correct academic answers.

***

### Q1: Why did you use a lognormal distribution for the Cost Performance Index (CPI) instead of a normal or beta distribution?
**Answer:** In EPC (Engineering, Procurement, and Construction) projects, cost variations are strongly right-skewed. Costs can theoretically overrun infinitely (due to severe delays or material price spikes), but they cannot go below zero. A lognormal distribution perfectly captures this reality: small variations around the mean are common, but extreme cost overruns (the "fat tail") are statistically possible. A normal distribution is symmetrical and could yield mathematically impossible negative costs in a simulation.

### Q2: How did you determine the variability parameters (low, medium, high complexity) for your generated projects?
**Answer:** The variability parameters (`0.04`, `0.08`, `0.14`) in the code represent the standard deviation of the log-transformed CPI. I assigned these to simulate realistic, distinct risk profiles for EPC projects. Low complexity projects have tight controls and predictable resources, while high complexity projects (like offshore platforms or nuclear builds) have much wider variance due to unforeseen engineering and execution risks. 

### Q3: Why did you choose 10,000 iterations for your Monte Carlo simulation?
**Answer:** 10,000 iterations (`N_ITERATIONS = 10000`) is the industry standard for Monte Carlo simulations in quantitative risk analysis. It provides highly stable convergence of the probability distributions. Running fewer (e.g., 1,000) might result in volatile or inaccurate 'tails' (P10 and P90 confidence intervals). Running more (e.g., 100,000) would increase computational time unnecessarily without significantly improving the accuracy of the P50 or the confidence intervals.

### Q4: Your simulation shows that SEVM outperforms Deterministic EVM. Why does Deterministic EVM mathematically fail in these scenarios?
**Answer:** Deterministic EVM (using basic formulas like `EAC = BAC / CPI`) assumes that past performance is a perfectly static indicator of the future. It provides a single point estimate and completely ignores risk and uncertainty. SEVM corrects this structural flaw by treating future performance as a probability distribution based on the *volatility* of past performance, thereby capturing the compounding effect of project uncertainty.

### Q5: Why did you use the P50 (median) instead of the mean (expected value) as your primary SEVM point estimate for calculating Mean Absolute Percentage Error (MAPE)?
**Answer:** Because the cost distribution in complex projects is right-skewed (lognormal), the "mean" is pulled disproportionately upwards by extreme worst-case scenarios in the simulation's long tail. The P50 (median) is a more robust, stable, and statistically likely outcome. It represents a strict 50/50 probability of being over or under the budget estimate, which is the standard baseline used by project managers.

### Q6: How does your script model the correlation between schedule and cost (the "composite" EAC)?
**Answer:** The script uses the formula `eac_sevm = ac_now + remaining / (sim_cpi * sim_spi)`. In the Monte Carlo engine, it samples both a simulated CPI (lognormal) and a simulated SPI (normal) for thousands of simulated futures. This mathematically models the reality that if a project is behind schedule (low SPI), the remaining work will take longer, which directly drags down the cost efficiency (burn rate) and drives the final simulated cost higher.

### Q7: In your `generate_epc_project` function, you force "declining" or "volatile" CPI trends. Why not just use completely random data?
**Answer:** I wanted the data to reflect real-world, empirical EPC behavior rather than pure white noise. EPC projects rarely have perfectly flat or purely random CPIs. Literature demonstrates they typically degrade mid-project due to integration complexities, design rework, or the "planning fallacy" fading away. By forcing these realistic trends, the test dataset tests the SEVM model against the actual kinds of data project managers deal with daily.

### Q8: What do the "80% CI Calibration" and "50% CI Calibration" metrics signify in your comparative analysis?
**Answer:** Calibration measures the *reliability and trustworthiness* of the probabilistic model. An 80% CI calibration checks if the *actual* final cost of the project actually fell within the P10 and P90 boundaries predicted by the simulation mid-flight. If the calibration score is close to 80%, it proves the model is not just accurate on average (MAPE), but that its stated confidence intervals are statistically sound for executive decision-making.

### Q9: You used a Wilcoxon test (`scipy.stats.wilcoxon`). Why this specific test, and what does the p-value prove for your thesis?
**Answer:** The Wilcoxon signed-rank test is a non-parametric test comparing two paired groups (SEVM's forecasting error against Deterministic EVM's forecasting error on the exact same projects). Because percentage errors (MAPE) are absolute metrics bounded at zero and not perfectly normally distributed, a standard T-test is inappropriate. Passing the Wilcoxon test (a p-value < 0.05) mathematically proves that SEVM's improved accuracy is statistically significant, validating my primary hypothesis (H1).

### Q10: What are the practical limitations of implementing this SEVM Python script in real-world EPC companies today?
**Answer:** The primary limitation is data maturity and hygiene. The mathematical model requires accurate, unmanipulated, and granular period-by-period EV, PV, and AC data to fit the distributions properly. If a company's data is heavily "smoothed" by project managers to hide bad news, or only updated quarterly instead of monthly, the simulation won't have enough variance data to calculate accurate lognormal distributions and confidence intervals.
