# Employee Turnover Prediction

**Which employees are most likely to leave, and why?** Logistic regression on the IBM HR Attrition dataset (1,470 employees), built in Python and reproduced in Excel.

[Data source: IBM HR Analytics — Kaggle](https://www.kaggle.com/datasets/pavansubhasht/ibm-hr-analytics-attrition-dataset)

## What the model shows

20 factors selected from 44 candidates using backward elimination (p < 0.01 cutoff). On a held-out test set:

| Metric | Result |
|---|---|
| AUC | 0.85 |
| Leavers caught (at 30% cutoff) | 70% |
| Overall attrition rate | 16.1% (237 of 1,470) |

## Key findings

- **Overtime is the biggest warning sign** — overtime workers leave at 3x the rate of non-overtime staff (30.5% vs 10.4%), and still multiply the odds of leaving by 6.7x after accounting for everything else
- **Sales Representatives have the highest turnover of any role** at 39.8%
- **Single employees leave at roughly twice the rate of married staff** — 25.5% vs ~12%
- **Satisfaction matters more than pay** — job involvement, environment satisfaction, job satisfaction, and work-life balance each cut the odds of leaving by 30–43% per point on their 1–4 scales
- **Income alone doesn't explain it** — leavers earned less on average ($4,787/mo vs $6,833), but pay tracks with job level and experience, so once those are in the model, income adds little new information

## Top factors

| Factor | Odds multiplier |
|---|---|
| Works overtime | 6.70x |
| Travels frequently (vs never) | 6.62x |
| Sales Representative | 5.40x |
| HR role | 3.56x |
| Single (vs divorced) | 3.12x |
| Laboratory Technician | 2.99x |
| Travels sometimes (vs never) | 2.86x |
| Technical degree | 2.38x |
| Sales Executive | 2.25x |
| Years since last promotion | +20% per year |
| Job involvement (1–4) | -43% per point |
| Environment satisfaction (1–4) | -35% per point |
| Job satisfaction (1–4) | -33% per point |

Full table with coefficients and p-values in `model_results.csv`.

## Tools used

- **Python** — pandas, statsmodels, scikit-learn; backward elimination logistic regression, model evaluation, Excel output
- **Excel** — formula-based model reproduction (SUMPRODUCT + EXP), Solver verification

## Files in this repo

| File | What it is |
|---|---|
| `analysis.py` | Python script: cleans data, runs the model, outputs all files |
| `attrition_clean.csv` | Cleaned dataset (1,470 rows, 33 columns) |
| `model_results.csv` | 20 factors with coefficients, odds multipliers, and p-values |
| `turnover_analysis.xlsx` | Excel workbook: Clean Data, Summary Stats, Charts, Regression Results, Model + Predictions |

Raw dataset not included — download from Kaggle using the link above.
