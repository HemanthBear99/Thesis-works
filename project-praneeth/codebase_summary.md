# SEVM Simulation Codebase Summary

The codebase primarily consists of a single Python script (`simulation_results/sevm_simulation.py`) that acts as a comprehensive engine for generating, simulating, and analyzing data for an MBA Thesis on **Stochastic Earned Value Management (SEVM)**.

## 🎯 Primary Purpose

The code generates simulated Engineering, Procurement, and Construction (EPC) project data, runs Monte Carlo simulations (10,000 iterations) on that data, and proves that SEVM (a probabilistic approach) is far more statistically accurate for predicting total project cost (Estimate at Completion, or EAC) than traditional Deterministic EVM (using basic CPI and composite SPI formulas).

## ⚙️ Code Structure & Architecture (6 Key Steps)

### 1. Generate Realistic EPC Project Data (`generate_epc_project` function)
- The script simulates exactly 12 projects (`P01` through `P12`) with varied characteristics such as Budget at Completion (BAC), planned durations, complexities, and cost trends.
- It models "realistic" performance issues typically found in EPC projects: S-Curve Planned Value (PV), declining or volatile Cost Performance Index (CPI) trends, right-skewed lognormal distributions for cost spikes, and normal distribution for SPI.

### 2. Compute Baseline EVM Metrics (`compute_evm_metrics` function)
- The code iterates through the project timeline and calculates standard period-by-period Earned Value metrics. 
- These include Percent Complete, cumulative CPI, cumulative SPI, Deterministic EAC (using CPI), and Deterministic EAC (Composite considering both cost & schedule).

### 3. SEVM Monte Carlo Simulation (`run_sevm` function)
- The simulation takes "actual" historical period CPI/SPIs for a project up to a specific forecast point (e.g., 20% to 90% complete).
- It fits a lognormal distribution for CPI (to better predict cost "surprises") and a normal distribution for SPI. 
- It simulates 10,000 future paths to project completion, returning a probability distribution for the EAC. It extracts key metrics like the P10, P50 (the SEVM median estimate), and P90 probabilities, as well as checking the MAPE (Mean Absolute Percentage Error) against the true simulation outcome.

### 4. Comparative Statistical Analysis
- It aggregates data across all projects at various lifecycle stages (20%, 30%, ... 90%).
- It tracks the percentage point (pp) improvement of the SEVM approach vs. Deterministic EVM.
- To mathematically prove the superiority of the SEVM approach, it runs **Wilcoxon statistical tests**.

### 5. Chart Generation
- `matplotlib` is used to output highly detailed PNG graphs to the `simulation_results/` folder:
  - **Fig 2**: Mean Absolute Percentage Error comparison bar chart.
  - **Fig 3**: Traditional S-curve but with 50% completion SEVM confidence bands (P10 to P90).
  - **Fig 4**: Histogram representation of the EAC Monte Carlo Probability Density.
  - **Fig 5**: Boxplot showing the effect of complexity on SEVM accuracy advantage.
  - **Fig 6**: Cumulative Distribution Function (CDF) for EAC.
  - **Fig 7**: CPI variability over lifecycle.
  - **Fig 8**: Accuracy advantage tracked by project phase.

### 6. Final Results Summary Output
- Finally, it writes out all these findings into `results_summary.txt`.
- It explicitly verifies five hypotheses (`H1` to `H5`) explicitly formatted to be copy-pasted directly into **Chapter 4 of your thesis**.
