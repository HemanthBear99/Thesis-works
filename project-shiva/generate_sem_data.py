"""
Generate Simulated SEM Survey Data - EXACT Google Forms Format
===============================================================
Column headers match EXACTLY what Google Forms outputs to Sheets.
"""

import numpy as np
import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta

np.random.seed(42)

N_RESPONSES = 150

# Exact column headers from Google Forms
COLUMNS = {
    'timestamp': 'Timestamp',
    'role': 'Current Role',
    'experience': 'Total Professional Experience',
    'industry': 'Industry Sector',
    'hybrid': 'Experience in Hybrid Project Environments (Agile + Waterfall)',
    'ai_tools': 'Use of AI-Based Tools in Project or Risk Management',
    
    # RPA
    'RPA1': 'RISK PROCESS ALIGNMENT (RPA) [Agile and Waterfall teams follow coordinated risk identification practices.]',
    'RPA2': 'RISK PROCESS ALIGNMENT (RPA) [Risk assessment criteria are consistent across Agile and Waterfall methodologies.]',
    'RPA3': 'RISK PROCESS ALIGNMENT (RPA) [Risk review cycles are aligned between Agile and Waterfall teams.]',
    'RPA4': 'RISK PROCESS ALIGNMENT (RPA) [Risks identified in Agile projects are effectively integrated into overall project risk registers.]',
    'RPA5': 'RISK PROCESS ALIGNMENT (RPA) [Risk management processes are harmonized across project methodologies.]',
    
    # RVI
    'RVI1': 'RISK VISIBILITY INTEGRATION (RVI) [Risk information is easily accessible to all project stakeholders.]',
    'RVI2': 'RISK VISIBILITY INTEGRATION (RVI) [Risks from Agile and Waterfall teams are consolidated in a unified system.]',
    'RVI3': 'RISK VISIBILITY INTEGRATION (RVI) [Real-time visibility of project risks is available.]',
    'RVI4': 'RISK VISIBILITY INTEGRATION (RVI) [Risk data from different tools and teams is well integrated.]',
    'RVI5': 'RISK VISIBILITY INTEGRATION (RVI) [Risk reporting provides a holistic view of the project.]',
    
    # TRE
    'TRE1': 'TIMELY RISK ESCALATION (TRE) [Project risks are escalated promptly when identified.]',
    'TRE2': 'TIMELY RISK ESCALATION (TRE) [Early warning signs of risks are addressed without delay.]',
    'TRE3': 'TIMELY RISK ESCALATION (TRE) [Escalation procedures for risks are clearly defined.]',
    'TRE4': 'TIMELY RISK ESCALATION (TRE) [Cross-team risks are communicated efficiently.]',
    'TRE5': 'TIMELY RISK ESCALATION (TRE) [Delays in risk escalation rarely occur.]',
    
    # GC
    'GC1': 'GOVERNANCE CONSISTENCY (GC) [Roles and responsibilities for risk management are clearly defined.]',
    'GC2': 'GOVERNANCE CONSISTENCY (GC) [Risk governance structures are consistent across Agile and Waterfall projects.]',
    'GC3': 'GOVERNANCE CONSISTENCY (GC) [Risk ownership is clearly assigned and understood.]',
    'GC4': 'GOVERNANCE CONSISTENCY (GC) [Decision-making authority related to risks is well established.]',
    'GC5': 'GOVERNANCE CONSISTENCY (GC) [Governance practices support effective risk control.]',
    
    # AIC
    'AIC1': 'AI-BASED RISK ANALYTICS CAPABILITY (AIC) [AI-based tools help identify project risks early.]',
    'AIC2': 'AI-BASED RISK ANALYTICS CAPABILITY (AIC) [AI improves the accuracy of risk assessments.]',
    'AIC3': 'AI-BASED RISK ANALYTICS CAPABILITY (AIC) [AI supports analysis of unstructured project data (e.g., reports, logs).]',
    'AIC4': 'AI-BASED RISK ANALYTICS CAPABILITY (AIC) [AI-driven dashboards enhance risk monitoring.]',
    'AIC5': 'AI-BASED RISK ANALYTICS CAPABILITY (AIC) [Predictive analytics supports proactive risk mitigation.]',
    
    # RME
    'RME1': 'RISK MANAGEMENT EFFECTIVENESS (RME) [Project risks are effectively managed in hybrid project environments.]',
    'RME2': 'RISK MANAGEMENT EFFECTIVENESS (RME) [Risk mitigation actions are timely and appropriate.]',
    'RME3': 'RISK MANAGEMENT EFFECTIVENESS (RME) [Risk-related decisions improve overall project performance.]',
    'RME4': 'RISK MANAGEMENT EFFECTIVENESS (RME) [Project uncertainty is reduced through effective risk management.]',
    'RME5': 'RISK MANAGEMENT EFFECTIVENESS (RME) [Overall project success is enhanced by current risk management practices.]',
}

# Demographic options (exact Google Forms values)
ROLES = [
    "Project Manager",
    "Scrum Master / Agile Coach",
    "PMO Member",
    "Risk Manager",
    "Team Lead",
    "Consultant",
    "Other"
]

EXPERIENCE_LEVELS = [
    "Less than 2 years",
    "2-5 years",
    "6-10 years",
    "More than 10 years"
]

INDUSTRY_SECTORS = [
    "IT / Software",
    "Manufacturing",
    "Finance / Banking",
    "Consulting",
    "Telecommunications",
    "Other"
]

# Likert scale responses (exact Google Forms values)
LIKERT_OPTIONS = [
    "Strongly Disagree",
    "Disagree",
    "Neutral",
    "Agree",
    "Strongly Agree"
]


def generate_timestamps(n: int) -> list:
    """Generate realistic timestamps spread over ~2 weeks."""
    base_date = datetime(2026, 1, 10, 9, 0, 0)
    timestamps = []
    for i in range(n):
        # Random offset: 0-14 days, random hours
        days_offset = np.random.randint(0, 15)
        hours_offset = np.random.randint(8, 22)  # 8 AM - 10 PM
        minutes_offset = np.random.randint(0, 60)
        seconds_offset = np.random.randint(0, 60)
        
        ts = base_date + timedelta(
            days=days_offset,
            hours=hours_offset,
            minutes=minutes_offset,
            seconds=seconds_offset
        )
        timestamps.append(ts.strftime("%m/%d/%Y %H:%M:%S"))
    
    # Sort by timestamp
    return sorted(timestamps)


def generate_demographics(n: int) -> dict:
    """Generate demographic responses."""
    role_weights = [0.25, 0.15, 0.12, 0.10, 0.18, 0.12, 0.08]
    exp_weights = [0.12, 0.30, 0.35, 0.23]
    industry_weights = [0.35, 0.12, 0.20, 0.15, 0.10, 0.08]
    
    return {
        'role': np.random.choice(ROLES, n, p=role_weights),
        'experience': np.random.choice(EXPERIENCE_LEVELS, n, p=exp_weights),
        'industry': np.random.choice(INDUSTRY_SECTORS, n, p=industry_weights),
        'hybrid': np.random.choice(["Yes", "No"], n, p=[0.82, 0.18]),
        'ai_tools': np.random.choice(["Yes", "No"], n, p=[0.68, 0.32]),
    }


def generate_correlated_latents(n: int, demographics: dict) -> dict:
    """Generate correlated latent variables."""
    exo_corr = np.array([
        [1.00, 0.45, 0.40, 0.35, 0.50],
        [0.45, 1.00, 0.42, 0.38, 0.48],
        [0.40, 0.42, 1.00, 0.44, 0.46],
        [0.35, 0.38, 0.44, 1.00, 0.40],
        [0.50, 0.48, 0.46, 0.40, 1.00],
    ])
    
    exo_latents = np.random.multivariate_normal(np.zeros(5), exo_corr, n)
    
    # AI users score higher on AIC
    ai_boost = np.array([0.4 if x == "Yes" else 0 for x in demographics['ai_tools']])
    exo_latents[:, 4] += ai_boost
    
    # Experience boost for RPA and GC
    exp_boost = np.array([
        0.0 if x == "Less than 2 years" else
        0.1 if x == "2-5 years" else
        0.2 if x == "6-10 years" else
        0.3 for x in demographics['experience']
    ])
    exo_latents[:, 0] += exp_boost
    exo_latents[:, 3] += exp_boost
    
    latents = {
        'RPA': exo_latents[:, 0],
        'RVI': exo_latents[:, 1],
        'TRE': exo_latents[:, 2],
        'GC': exo_latents[:, 3],
        'AIC': exo_latents[:, 4],
    }
    
    # RME depends on all predictors
    residual = np.random.normal(0, np.sqrt(0.35), n)
    latents['RME'] = (
        0.15 * latents['RPA'] +
        0.12 * latents['RVI'] +
        0.14 * latents['TRE'] +
        0.10 * latents['GC'] +
        0.28 * latents['AIC'] +
        residual
    )
    
    return latents


def generate_indicators(latent: np.ndarray, n_indicators: int) -> np.ndarray:
    """Generate indicators from latent variable."""
    n = len(latent)
    indicators = np.zeros((n, n_indicators))
    
    for i in range(n_indicators):
        loading = np.random.uniform(0.75, 0.90)
        error = np.random.normal(0, np.sqrt(1 - loading**2), n)
        indicators[:, i] = loading * latent + error
    
    return indicators


def transform_to_likert_text(values: np.ndarray) -> np.ndarray:
    """Transform continuous values to Likert scale TEXT responses."""
    z = (values - np.mean(values)) / np.std(values)
    
    # Map to 0-4 index
    result = np.zeros(len(z), dtype=int)
    result[z > -0.84] = 1
    result[z > -0.25] = 2
    result[z > 0.25] = 3
    result[z > 0.84] = 4
    
    # Slight positive skew
    bump_mask = np.random.random(len(result)) < 0.15
    result = np.where(bump_mask & (result < 4), result + 1, result)
    
    result = np.clip(result, 0, 4)
    
    # Convert to text
    return np.array([LIKERT_OPTIONS[i] for i in result])


def generate_dataset() -> pd.DataFrame:
    """Generate complete dataset with exact Google Forms format."""
    print(f"Generating {N_RESPONSES} survey responses (Google Forms format)...")
    
    demographics = generate_demographics(N_RESPONSES)
    latents = generate_correlated_latents(N_RESPONSES, demographics)
    
    # Build data with exact column names
    data = {
        COLUMNS['timestamp']: generate_timestamps(N_RESPONSES),
        COLUMNS['role']: demographics['role'],
        COLUMNS['experience']: demographics['experience'],
        COLUMNS['industry']: demographics['industry'],
        COLUMNS['hybrid']: demographics['hybrid'],
        COLUMNS['ai_tools']: demographics['ai_tools'],
    }
    
    # Generate Likert items
    constructs = ['RPA', 'RVI', 'TRE', 'GC', 'AIC', 'RME']
    
    for construct in constructs:
        indicators = generate_indicators(latents[construct], 5)
        for i in range(5):
            key = f"{construct}{i+1}"
            data[COLUMNS[key]] = transform_to_likert_text(indicators[:, i])
    
    return pd.DataFrame(data)


def main():
    df = generate_dataset()
    
    output_path = Path("d:/project-shiva/SEM_Survey_GoogleForms_Format.xlsx")
    
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Form Responses 1', index=False)
    
    print(f"\n✅ Excel file saved: {output_path}")
    print(f"   Rows: {len(df)}")
    print(f"   Columns: {len(df.columns)}")
    print(f"\n📋 Column headers match Google Forms EXACTLY!")
    print("\n🎯 Ready to use!")


if __name__ == "__main__":
    main()
