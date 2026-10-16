"""Employee turnover analysis (IBM HR Attrition dataset).

Run:  python analysis.py WA_Fn-UseC_-HR-Employee-Attrition.csv
Makes: attrition_clean.csv, model_results.csv, turnover_analysis.xlsx
"""
import sys
import numpy as np
import pandas as pd
import statsmodels.api as sm
from sklearn.metrics import roc_auc_score, confusion_matrix
from sklearn.model_selection import train_test_split
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L

SRC = sys.argv[1]
OUT = sys.argv[2] if len(sys.argv) > 2 else "."

# ---------- Step 1: load and check ----------
raw = pd.read_csv(SRC, encoding="utf-8-sig")  # file starts with an invisible BOM character
print("Raw shape:", raw.shape)
print("Missing values:", int(raw.isna().sum().sum()), "| Duplicate rows:", int(raw.duplicated().sum()))

# ---------- Step 2: clean ----------
constant = [c for c in raw if raw[c].nunique() == 1]  # EmployeeCount, Over18, StandardHours
df = raw.drop(columns=constant)
df["Left"] = (df.Attrition == "Yes").astype(int)  # 1 = left, 0 = stayed (needed for averages and regression)
print("Dropped constant columns:", constant, "-> clean shape:", df.shape)
df.to_csv(f"{OUT}/attrition_clean.csv", index=False)

# ---------- Step 3: quick facts ----------
print("Attrition rate: {:.1%}".format(df.Left.mean()))
print(df.groupby("OverTime").Left.mean().round(3).to_dict())
print(df.groupby("JobRole").Left.mean().round(3).sort_values().to_dict())
print(df.groupby("Attrition").MonthlyIncome.mean().round(0).to_dict())

# ---------- Step 4: regression ----------
cat = ["BusinessTravel", "Department", "EducationField", "Gender", "JobRole", "MaritalStatus", "OverTime"]
X = pd.get_dummies(df.drop(columns=["Attrition", "Left", "EmployeeNumber"]), columns=cat, drop_first=True).astype(float)
y = df.Left

# 4a. Honest test: learn on 70%, test on the 30% the model never saw.
Xtr, Xte, ytr, yte = train_test_split(X, y, test_size=0.3, random_state=1, stratify=y)


def backward(Xa, ya, limit=0.01):
    """Start with every factor, drop the weakest (highest p-value) until all are below `limit`."""
    cols = list(Xa.columns)
    while True:
        m = sm.Logit(ya, sm.add_constant(Xa[cols])).fit(disp=0)
        pv = m.pvalues.drop("const")
        if pv.max() < limit:
            return cols
        cols.remove(pv.idxmax())


test_cols = backward(Xtr, ytr)
mt = sm.Logit(ytr, sm.add_constant(Xtr[test_cols])).fit(disp=0)
pt = mt.predict(sm.add_constant(Xte[test_cols]))
test = {"AUC": roc_auc_score(yte, pt)}
for t in (0.5, 0.3):
    tn, fp, fn, tp = confusion_matrix(yte, pt >= t).ravel()
    test[t] = dict(acc=(tn + tp) / len(yte), caught=tp / (tp + fn), precision=tp / (tp + fp))
print("Holdout test:", test)

# 4b. Final model: same selection on all 1,470 employees (most data = most reliable numbers).
cols = backward(X, y)
model = sm.Logit(y, sm.add_constant(X[cols])).fit(disp=0)
order = model.params[cols].sort_values(ascending=False).index.tolist()
cols = order

PLAIN = {
    "OverTime_Yes": "Works overtime", "BusinessTravel_Travel_Frequently": "Travels often (vs never)",
    "BusinessTravel_Travel_Rarely": "Travels sometimes (vs never)", "JobRole_Sales Representative": "Sales Representative",
    "JobRole_Human Resources": "HR job role", "MaritalStatus_Single": "Single (vs married)",
    "JobRole_Laboratory Technician": "Laboratory Technician", "EducationField_Technical Degree": "Technical degree",
    "JobRole_Sales Executive": "Sales Executive", "YearsSinceLastPromotion": "Years since last promotion",
    "NumCompaniesWorked": "Number of past employers", "DistanceFromHome": "Distance from home (miles)",
    "YearsInCurrentRole": "Years in current role", "TotalWorkingYears": "Total years working",
    "TrainingTimesLastYear": "Trainings last year", "RelationshipSatisfaction": "Relationship satisfaction (1-4)",
    "JobSatisfaction": "Job satisfaction (1-4)", "EnvironmentSatisfaction": "Environment satisfaction (1-4)",
    "WorkLifeBalance": "Work-life balance (1-4)", "JobInvolvement": "Job involvement (1-4)",
}
res = pd.DataFrame({
    "Factor": [PLAIN.get(c, c) for c in cols], "Column": cols,
    "Coefficient": model.params[cols].values, "OddsMultiplier": np.exp(model.params[cols].values),
    "p_value": model.pvalues[cols].values,
})
res.round(4).to_csv(f"{OUT}/model_results.csv", index=False)
print(res.round(3).to_string())

# ---------- Step 5: Excel workbook ----------
FONT = "Arial"
B = Font(name=FONT, bold=True)
TITLE = Font(name=FONT, bold=True, size=14)
HEAD = PatternFill("solid", fgColor="DDEBF7")
INPUT = PatternFill("solid", fgColor="FFFF00")
BAR = "2A78D6"
wb = Workbook()


def header(ws, row, values, col=1):
    for j, v in enumerate(values):
        c = ws.cell(row, col + j, v)
        c.font = B
        c.fill = HEAD
        c.alignment = Alignment(wrap_text=True, vertical="center")


def widths(ws, w):
    for k, v in w.items():
        ws.column_dimensions[k].width = v


# Read Me
rd = wb.active
rd.title = "Read Me"
lines = [
    ("Employee Turnover Analysis (IBM HR Attrition data)", TITLE),
    ("Question: which employees leave, and what predicts it?", None), ("", None),
    ("Sheets in this file", B),
    ("Clean Data - all 1,470 employees after cleaning (32 columns). Left = 1 if they left, 0 if they stayed.", None),
    ("Summary Stats - attrition rates by group and averages for leavers vs stayers. All live formulas on Clean Data.", None),
    ("Charts - the key pictures, built from Summary Stats.", None),
    ("Regression Results - which factors predict leaving, how strongly, and how accurate the model is.", None),
    ("Model Data + Predictions - the regression rebuilt with Excel formulas. Change the yellow cutoff on Predictions to play with it.", None),
    ("", None),
    ("Colour key", B), ("Yellow cells = things you can change. Everything else is a formula or the original data.", None),
    ("", None),
    ("Cleaning in one line", B),
    ("Removed EmployeeCount, Over18 and StandardHours (same value for every person, so they cannot explain anything). "
     "Added Left (1/0). No missing values, no duplicate rows, nothing else changed.", None),
]
for i, (t, f) in enumerate(lines, 1):
    c = rd.cell(i, 1, t)
    c.font = f or Font(name=FONT)
rd.column_dimensions["A"].width = 120

# Clean Data
cd = wb.create_sheet("Clean Data")
data_cols = list(df.columns)
header(cd, 1, data_cols)
for r, row in enumerate(df.itertuples(index=False), 2):
    for j, v in enumerate(row, 1):
        cd.cell(r, j, v.item() if hasattr(v, "item") else v)
cd.freeze_panes = "B2"
N = len(df) + 1
COL = {c: L(i + 1) for i, c in enumerate(data_cols)}


def rng(c):
    return f"'Clean Data'!${COL[c]}$2:${COL[c]}${N}"


LEFT = rng("Left")

# Summary Stats
ss = wb.create_sheet("Summary Stats")
ss["A1"] = "Summary statistics (all formulas read from Clean Data)"
ss["A1"].font = TITLE
ss["A3"], ss["B3"] = "Total employees", f"=COUNT({LEFT})"
ss["A4"], ss["B4"] = "Employees who left", f"=SUM({LEFT})"
ss["A5"], ss["B5"] = "Attrition rate", "=B4/B3"
ss["B5"].number_format = "0.0%"
for a in ("A3", "A4", "A5"):
    ss[a].font = B

blocks = {}
row = 7
for g in ["OverTime", "JobRole", "MaritalStatus", "BusinessTravel", "Department", "JobLevel"]:
    ss.cell(row, 1, f"Attrition by {g}").font = Font(name=FONT, bold=True, size=12)
    header(ss, row + 1, [g, "Employees", "Left", "Attrition rate"])
    levels = sorted(df[g].unique()) if g == "JobLevel" else sorted(df[g].unique(), key=lambda v: -df.loc[df[g] == v, "Left"].mean())
    start = row + 2
    for k, lv in enumerate(levels):
        r = start + k
        ss.cell(r, 1, lv.item() if hasattr(lv, "item") else lv)
        ss.cell(r, 2, f"=COUNTIF({rng(g)},A{r})")
        ss.cell(r, 3, f"=COUNTIFS({rng(g)},A{r},{LEFT},1)")
        ss.cell(r, 4, f"=C{r}/B{r}").number_format = "0.0%"
    blocks[g] = (start, start + len(levels) - 1)
    row = start + len(levels) + 1

# Leavers vs stayers averages (to the right)
NUMS = ["MonthlyIncome", "Age", "TotalWorkingYears", "YearsAtCompany", "YearsInCurrentRole", "YearsWithCurrManager",
        "DistanceFromHome", "NumCompaniesWorked", "JobLevel", "StockOptionLevel", "JobSatisfaction",
        "EnvironmentSatisfaction", "WorkLifeBalance", "JobInvolvement", "PercentSalaryHike", "TrainingTimesLastYear"]
ss["F7"] = "Average values: stayed vs left"
ss["F7"].font = Font(name=FONT, bold=True, size=12)
header(ss, 8, ["Measure", "Stayed (avg)", "Left (avg)", "Difference (left - stayed)"], col=6)
for k, c in enumerate(NUMS):
    r = 9 + k
    ss.cell(r, 6, c)
    ss.cell(r, 7, f"=AVERAGEIFS({rng(c)},{LEFT},0)").number_format = "#,##0.0"
    ss.cell(r, 8, f"=AVEQAGEMFS({rng(c)},{LEFT},1)").number_format = "#,##0.0"
    ss.cell(r, 9, f"=H{r}-G{r}").number_format = "#,##0.0;-#,##0.0"
inc_row = 9
ss.cell(9 + len(NUMS) + 1, 6, "Satisfaction and balance scores run 1 (low) to 4 (high). MonthlyIncome is in US dollars.").font = Font(name=FONT, italic=True)
# small 2-row table for the income chart
ss["F28"], ss["G28"] = "Group", "Avg monthly income ($)"
header(ss, 28, ["Group", "Avg monthly income ($)"], col=6)
ss["F29"], ss["G29"] = "Stayed", f"=G{inc_row}"
ss["F30"], ss["G30"] = "Left", f"=H{inc_row}"
ss["G29"].number_format = ss["G30"].number_format = "$#,##0"
widths(ss, {"A": 28, "B": 12, "C": 10, "D": 15, "F": 26, "G": 22, "H": 14, "I": 24})

# Charts
ch = wb.create_sheet("Charts")
ch["A1"] = "Charts (they update automatically from Summary Stats)"
ch["A1"].font = TITLE


def bar(title, ws, col, r1, r2, cat_col=1, horizontal=False, fmt="0%", ytitle="Attrition rate"):
    c = BarChart()
    c.type = "bar" if horizontal else "col"
    c.title = title
    c.legend = None
    c.y_axis.title = ytitle
    c.y_axis.numFmt = fmt
    c.y_axis.majorGridlines = None
    c.y_axis.scaling.min = 0  # bars always start at zero
    c.add_data(Reference(ws, min_col=col, min_row=r1, max_row=r2), titles_from_data=False)
    c.set_categories(Reference(ws, min_col=cat_col, min_row=r1, max_row=r2))
    s = c.series[0]
    s.graphicalProperties.solidFill = BAR
    s.graphicalProperties.line.solidFill = BAR
    c.dataLabels = DataLabelList()
    c.dataLabels.showVal = True
    c.dataLabels.showSerName = c.dataLabels.showCatName = c.dataLabels.showLegendKey = False
    c.dataLabels.numFmt = fmt
    c.gapWidth = 60
    c.width, c.height = 16, 8.5
    if horizontal:
        c.x_axis.scaling.orientation = "maxMin"  # highest bar on top
    return c


s0, s1 = blocks["OverTime"]
ch.add_chart(bar("Overtime workers leave about 3x as often", ss, 4, s0, s1), "A3")
s0, s1 = blocks["JobRole"]
ch.add_chart(bar("Attrition rate by job role", ss, 4, s0, s1, horizontal=True), "K3")
s0, s1 = blocks["MaritalStatus"]
ch.add_chart(bar("Single employees leave the most", ss, 4, s0, s1), "A21")
s0, s1 = blocks["BusinessTravel"]
ch.add_chart(bar("More travel, more leaving", ss, 4, s0, s1), "K21")
s0, s1 = blocks["JobLevel"]
ch.add_chart(bar("Entry-level (job level 1) staff leave the most", ss, 4, s0, s1), "A39")
ch.add_chart(bar("Leavers earned about $2,000 less a month", ss, 7, 29, 30, cat_col=6, fmt="$#,##0",
                 ytitle="Avg monthly income"), "K39")

# Regression Results (+ the coefficient row that drives the Excel model)
rr = wb.create_sheet("Regression Results")
rr["A1"] = "Logistic regression: what predicts an employee leaving?"
rr["A1"].font = TITLE
names = ["Intercept"] + cols
header(rr, 3, ["Coefficient row used by Model Data (yellow = Solver can change these) ->"] + names)
rr.cell(4, 1, "Coefficient").font = B
for j, nm in enumerate(names):
    c = rr.cell(4, 2 + j, round(float(model.params["const" if j == 0 else nm]), 6))
    c.fill = INPUT
LASTC = L(1 + len(names))
rr["A5"] = "Total log-likelihood (Solver: maximise this cell)"
rr["A5"].font = B

header(rr, 7, ["Factor", "Coefficient", "Odds multiplier =EXP(coef)", "What it means", "p-value", "Matters? (p < 0.05)"])
for i, r in res.iterrows():
    row = 8 + i
    rr.cell(row, 1, r.Factor)
    rr.cell(row, 2, f"={L(3 + i)}4").number_format = "0.000"
    rr.cell(row, 3, f"=EXP(B{row})").number_format = "0.00"
    om = r.OddsMultiplier
    dummy = "_" in r.Column
    if dummy:
        txt = f"Odds of leaving {om:.1f}x higher than the comparison group"
    elif om > 1:
        txt = f"Each +1 raises the odds of leaving by {(om - 1) * 100:.0f}%"
    else:
        txt = f"Each +1 lowers the odds of leaving by {(1 - om) * 100:.0f}%"
    rr.cell(row, 4, txt)
    rr.cell(row, 5, float(r.p_value)).number_format = '[<0.0001]"<0.0001";0.0000'
    rr.cell(row, 6, f'=IF(E{row}<0.05,"Yes","No")')
end = 8 + len(res)
rr.cell(end, 1, "p-values come from the Python statsmodels fit (see analysis.py); Excel cannot compute them directly.").font = Font(name=FONT, italic=True)

rr.cell(end + 2, 1, "How good is the model?").font = Font(name=FONT, bold=True, size=12)
header(rr, end + 3, ["Measure", "Value", "Plain English"])
fit = [
    ("Employees used", int(model.nobs), "", "0"),
    ("Factors kept (of 44 possible)", len(cols), "Weak factors removed one at a time until every one left had p < 0.01", "0"),
    ("McFadden pseudo R-squared", model.prsquared, "0.2 to 0.4 is considered a good fit for this kind of model", "0.000"),
    ("Likelihood-ratio test p-value", model.llr_pvalue, "Tiny = the model is far better than guessing", "0.0E+00"),
    ("AUC on unseen 30% test set", test["AUC"], "0.5 = coin flip, 1.0 = perfect", "0.00"),
    ("Test accuracy, cutoff 0.5", test[0.5]["acc"], "Share of all guesses that were right", "0.0%"),
    ("Test leavers caught, cutoff 0.5", test[0.5]["caught"], "Of people who really left, share we flagged", "0.0%"),
    ("Test accuracy, cutoff 0.3", test[0.3]["acc"], "", "0.0%"),
    ("Test leavers caught, cutoff 0.3", test[0.3]["caught"], "Lower cutoff catches more leavers, with more false alarms", "0.0%"),
    ("'Nobody leaves' lazy-guess accuracy", 1 - y.mean(), "Any model has to beat this, so accuracy alone is misleading", "0.0%"),
]
for k, (a, v, t, f) in enumerate(fit):
    row = end + 4 + k
    rr.cell(row, 1, a)
    rr.cell(row, 2, float(v)).number_format = f
    rr.cell(row, 3, t)
rr.cell(end + 4 + len(fit), 1, "Source: analysis.py (Python statsmodels + scikit-learn), random_state=1.").font = Font(name=FONT, italic=True)
widths(rr, {"A": 44, "B": 13, "C": 26, "D": 58, "E": 10, "F": 18})
for j in range(len(names)):
    rr.column_dimensions[L(2 + j)].width = max(rr.column_dimensions[L(2 + j)].width or 0, 13)
rr.freeze_panes = "A5"

# Odds chart on Regression Results
oc = BarChart()
oc.type = "bar"
oc.title = "Odds multiplier per factor (above 1 = more likely to leave)"
oc.legend = None
oc.y_axis.numFmt = "0.0"
oc.y_axis.majorGridlines = None
oc.add_data(Reference(rr, min_col=3, min_row=8, max_row=end - 1), titles_from_data=False)
oc.set_categories(Reference(rr, min_col=1, min_row=8, max_row=end - 1))
oc.series[0].graphicalProperties.solidFill = BAR
oc.series[0].graphicalProperties.line.solidFill = BAR
oc.x_axis.scaling.orientation = "maxMin"
oc.gapWidth = 40
oc.width, oc.height = 18, 12
rr.add_chart(oc, f"H7")

# Model Data
md = wb.create_sheet("Model Data")
hdr = ["EmployeeNumber", "Left (1=yes)"] + cols + ["Score", "Probability of leaving", "Predicted to leave? (1=yes)", "Log-likelihood piece"]
header(md, 1, hdr)
k = len(cols)
fc, lc = L(3), L(2 + k)
sc, pc, prc, llc = (L(3 + k + i) for i in range(4))
for i in range(len(df)):
    r = i + 2
    md.cell(r, 1, int(df.EmployeeNumber.iat[i]))
    md.cell(r, 2, int(df.Left.iat[i]))
    for j, cn in enumerate(cols):
        v = float(X[cn].iat[i])
        md.cell(r, 3 + j, int(v) if v.is_integer() else v)
    md[f"{sc}{r}"] = f"='Regression Results'!$B$4+SUMPRODUCT('Regression Results'!$C$4:${LASTC}$4,{fc}{r}:{lc}{r})"
    md[f"{pc}{r}"] = f"=1/(1+EXP(-{sc}{r}))"
    md[f"{pc}{r}"].number_format = "0.0%"
    md[f"{prc}{r}"] = f"=IF({pc}{r}>=Predictions!$B$2,1,0)"
    md[f"{llc}{r}"] = f"=B{r}*LN(MAX({pc}{r},1E-12))+(1-B{r})*LN(MAX(1-{pc}{r},1E-12))"
md.freeze_panes = "C2"
rr["B5"] = f"=SUM('Model Data'!{llc}2:{llc}{N})"

# Predictions
pr = wb.create_sheet("Predictions")
Dr, Pr = f"'Model Data'!$B$2:$B${N}", f"'Model Data'!${prc}$2:${prc}${N}"
cells = {
    "A1": "TRY IT: change the yellow cutoff and watch the numbers move",
    "A2": "Cutoff: flag someone as a likely leaver if their probability is at least", "B2": 0.3,
    "B4": "Actually STAYED", "C4": "Actually LEFT",
    "A5": "Model said STAY", "B5": f"=COUNTIFS({Dr},0,{Pr},0)", "C5": f"=COUNTIFS({Dr},1,{Pr},0)",
    "A6": "Model said LEAVE", "B6": f"=COUNTIFS({Dr},0,{Pr},1)", "C6": f"=COUNTIFS({Dr},1,{Pr},1)",
    "A8": "Accuracy (share of all guesses that were right)", "B8": "=(B5+C6)/SUM(B5:C6)",
    "A9": "Leavers caught (of 
