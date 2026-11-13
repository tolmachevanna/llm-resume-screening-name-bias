"""
Analysis script for the bachelor thesis on name-based bias in LLM resume screening.

Input : Excel file with one row per model evaluation (default: alle_daten_final.xlsx).
        Required columns (names as in the data file):
        Name, Gruppe, Geschlecht, Stelle (Prompt), Lebenslauf-Beruf, Bedingung,
        Modell, Durchgang, Score (1-10), Ausgeschlossen
Output: analysis_results.xlsx (one sheet per result table) and a printed summary.

Usage : python analysis.py [input.xlsx] [output.xlsx] [--keep-excluded]
        With --keep-excluded the two excluded observations are retained (sensitivity
        check reported in the Results chapter); without it they are dropped.

Packages: numpy, pandas, scipy, openpyxl (for reading/writing .xlsx).

What is computed (see Chapter "Methodology", Section "Statistical Analysis"):
  1. Descriptive statistics per model, condition and name group.
  2. OLS with fixed effects for the resume x target occupation cell; omnibus F-test
     for the name-group factor (3 df); Holm correction across all testable omnibus tests.
  3. Differences of each non-German group to the German group (score points, 95% CI,
     Cohen's d = difference / residual SD).
  4. Gender analysis (female = reference), per model/condition and pooled, plus a
     supplementary Holm correction over the 13 testable model-specific and the 2 pooled
     gender tests (not part of the decision rule).
  5. Pooled analysis across models (fixed effects per model x cell).
  6. Range between the two names within each name group (descriptive).
  7. Sensitivity check: permutation tests (within-cell row permutation, 5,000 draws;
     exact permutation of group labels across the 8 names, 2,520 assignments).
"""
import sys
import itertools

import numpy as np
import pandas as pd
from scipy import stats

KEEP_EXCLUDED = "--keep-excluded" in sys.argv               # sensitivity check: keep the 2 excluded rows
ARGS = [a for a in sys.argv[1:] if not a.startswith("--")]
INPUT = ARGS[0] if len(ARGS) > 0 else "alle_daten_final.xlsx"
OUTPUT = ARGS[1] if len(ARGS) > 1 else "analysis_results.xlsx"

SCORE = "Score (1-10)"
GROUPS = ["Deutsch", "Russisch", "Spanisch", "Türkisch"]   # reference group first
GENDERS = ["weiblich", "männlich"]                          # reference group first
N_PERMUTATIONS = 5000                                       # row-level permutation tests
SEED = 1                                                    # fixed random seed


# --------------------------------------------------------------------------- data
def load_data(path):
    """Read the data, drop the two excluded observations (unless --keep-excluded is given),
    define the resume x target cell."""
    df = pd.read_excel(path)
    n_all = len(df)
    if not KEEP_EXCLUDED:
        df = df[~df["Ausgeschlossen"]].copy()
    print(f"{n_all} evaluations read, {n_all - len(df)} excluded, {len(df)} analyzed.")
    # cell = combination of target occupation (prompt) and resume occupation;
    # in the matched condition this is simply the occupation
    df["cell"] = df["Stelle (Prompt)"] + "|" + df["Lebenslauf-Beruf"]
    return df


# ------------------------------------------------------------------------ models
def ols(y, X):
    """OLS via least squares; returns coefficients, residuals and residual df."""
    beta, *_ = np.linalg.lstsq(X, y, rcond=None)
    resid = y - X @ beta
    return beta, resid, X.shape[0] - np.linalg.matrix_rank(X)


def analyse(g, fe_col, factor, levels, ref):
    """
    Regression  y = fixed effects (fe_col) + dummies for `factor` (reference level `ref`).
    Returns the omnibus F-test of the factor, the coefficients with 95% CIs
    (t-distribution) and the residual standard deviation.
    """
    y = g[SCORE].values.astype(float)
    fixed = pd.get_dummies(g[fe_col], dtype=float).values
    non_ref = [lv for lv in levels if lv != ref]
    dummies = pd.get_dummies(pd.Categorical(g[factor], categories=levels))[non_ref].astype(float).values
    X = np.hstack([fixed, dummies])
    k = dummies.shape[1]

    _, r0, _ = ols(y, fixed)           # restricted model: cell fixed effects only
    beta, r1, df1 = ols(y, X)          # full model: plus the factor
    s2 = (r1 @ r1) / df1
    F = ((r0 @ r0 - r1 @ r1) / k) / s2
    p = 1 - stats.f.cdf(F, k, df1)
    se = np.sqrt(np.diag(s2 * np.linalg.pinv(X.T @ X)))[-k:]
    est = beta[-k:]
    tcrit = stats.t.ppf(0.975, df1)
    ci = [(e - tcrit * s, e + tcrit * s) for e, s in zip(est, se)]
    return dict(F=F, p=p, est=est, ci=ci, sd_resid=np.sqrt(s2), n=len(y), df_resid=df1)


def holm(pvalues):
    """Holm-adjusted p-values; NaN entries (untestable) are ignored."""
    p = pd.Series(pvalues, dtype=float)
    adj = pd.Series(np.nan, index=p.index)
    valid = p.dropna().sort_values()
    m = len(valid)
    running = 0.0
    for rank, (idx, val) in enumerate(valid.items()):
        running = max(running, min(1.0, (m - rank) * val))
        adj[idx] = running
    return adj


# ------------------------------------------------------------ main analysis tables
def main_tables(df):
    omnibus, diffs, gender = [], [], []
    for (model, cond), g in df.groupby(["Modell", "Bedingung"]):
        a = analyse(g, "cell", "Gruppe", GROUPS, "Deutsch")
        testable = a["sd_resid"] > 1e-9    # zero variance -> F-test not defined
        omnibus.append(dict(
            Modell=model, Bedingung=cond, n=a["n"],
            F=a["F"] if testable else np.nan, p=a["p"] if testable else np.nan,
            SDres=a["sd_resid"], Gesamtmittel=g[SCORE].mean(),
            Anteil_max=(g[SCORE] == g[SCORE].max()).mean()))
        for lv, e, ci in zip(GROUPS[1:], a["est"], a["ci"]):
            diffs.append(dict(Modell=model, Bedingung=cond, Gruppe=lv, Diff_zu_Deutsch=e,
                              CI_lo=ci[0], CI_hi=ci[1], d=e / a["sd_resid"] if testable else np.nan))
        ge = analyse(g, "cell", "Geschlecht", GENDERS, "weiblich")
        gender.append(dict(Modell=model, Bedingung=cond, Diff_mw=ge["est"][0],
                           CI_lo=ge["ci"][0][0], CI_hi=ge["ci"][0][1], p=ge["p"]))
    omnibus = pd.DataFrame(omnibus)
    omnibus["p_Holm"] = holm(omnibus["p"])      # family: all testable omnibus tests, both conditions
    return omnibus, pd.DataFrame(diffs), pd.DataFrame(gender)


def pooled_table(df):
    """Pooled over all models; fixed effects per model x cell."""
    rows = []
    for cond, g in df.groupby("Bedingung"):
        g = g.copy()
        g["fe"] = g["Modell"] + "|" + g["cell"]
        a = analyse(g, "fe", "Gruppe", GROUPS, "Deutsch")
        row = dict(Bedingung=cond, n=a["n"], F=a["F"], p=a["p"])
        for lv, e, c in zip(GROUPS[1:], a["est"], a["ci"]):
            row[f"{lv}_diff"] = e
            row[f"{lv}_CI"] = f"[{c[0]:.2f}, {c[1]:.2f}]"
        ge = analyse(g, "fe", "Geschlecht", GENDERS, "weiblich")
        row.update(Gender_diff=ge["est"][0], Gender_CI=f"[{ge['ci'][0][0]:.2f}, {ge['ci'][0][1]:.2f}]",
                   Gender_p=ge["p"])
        rows.append(row)
    return pd.DataFrame(rows)


def gender_holm_table(gender, omnibus, pooled):
    """
    Supplementary Holm correction for the gender tests: the 13 testable model-condition
    tests (Llama 4 Maverick in the matched condition has zero variance and is left out)
    plus the 2 pooled tests (one per condition) form one family of 15 tests.
    Not used in the decision rule for the research questions.
    """
    g = gender.merge(omnibus[["Modell", "Bedingung", "SDres"]], on=["Modell", "Bedingung"])
    g["p_test"] = g["p"].where(g["SDres"] > 1e-9)          # NaN = untestable
    model_part = pd.DataFrame({"Test": g["Modell"] + " | " + g["Bedingung"],
                               "Diff_mw": g["Diff_mw"], "p": g["p_test"]})
    pooled_part = pd.DataFrame({"Test": "Pooled | " + pooled["Bedingung"],
                                "Diff_mw": pooled["Gender_diff"], "p": pooled["Gender_p"]})
    out = pd.concat([model_part, pooled_part], ignore_index=True)
    out["p_Holm"] = holm(out["p"])
    return out


def descriptive_tables(df):
    desc = df.groupby(["Modell", "Bedingung", "Gruppe"])[SCORE].agg(["mean", "std", "count"]).reset_index()
    names = df.groupby(["Modell", "Bedingung", "Gruppe", "Name"])[SCORE].mean().reset_index()
    # range between the two names of a group, averaged over the four groups
    rng = names.groupby(["Modell", "Bedingung", "Gruppe"])[SCORE].agg(lambda s: s.max() - s.min())
    spread = rng.groupby(level=["Modell", "Bedingung"]).mean().rename("mittlere_Spanne").reset_index()
    return desc, names, spread


# ------------------------------------------------------------ permutation tests
def f_stat(y, cell_codes, group_codes):
    """
    F statistic of the group factor given cell fixed effects (Frisch-Waugh-Lovell:
    demean the outcome and the group dummies within cells, then compare residual sums
    of squares). Equivalent to the omnibus F-test of analyse().
    """
    n = len(y)
    y_dm = y - pd.Series(y).groupby(cell_codes).transform("mean").values
    D = np.zeros((n, 4))
    D[np.arange(n), group_codes] = 1
    D_dm = D - pd.DataFrame(D).groupby(cell_codes).transform("mean").values
    D_dm = D_dm[:, 1:]                                   # drop the reference group
    beta, *_ = np.linalg.lstsq(D_dm, y_dm, rcond=None)
    ss1 = ((y_dm - D_dm @ beta) ** 2).sum()
    ss0 = (y_dm ** 2).sum()
    k = np.linalg.matrix_rank(D_dm)
    df_res = n - len(np.unique(cell_codes)) - k
    if ss1 < 1e-12:
        return np.nan
    return ((ss0 - ss1) / k) / (ss1 / df_res)


def permutation_tests(df):
    rng = np.random.default_rng(SEED)
    names = sorted(df["Name"].unique())
    name_index = {n: i for i, n in enumerate(names)}
    # all distinct assignments of the 4 groups (2 names each) to the 8 names: 8!/(2!)^4 = 2520
    assignments = sorted(set(itertools.permutations([0, 0, 1, 1, 2, 2, 3, 3])))
    rows = []
    for (model, cond), g in df.groupby(["Modell", "Bedingung"]):
        y = g[SCORE].values.astype(float)
        cell = pd.factorize(g["cell"])[0]
        grp = pd.Categorical(g["Gruppe"], categories=GROUPS).codes
        F = f_stat(y, cell, grp)
        if np.isnan(F):                                  # zero variance: not testable
            rows.append(dict(Modell=model, Bedingung=cond, F=np.nan, p_perm_Zeilen=np.nan,
                             p_perm_Namen=np.nan))
            continue
        # (a) permute group labels within each cell (row level)
        count = 0
        for _ in range(N_PERMUTATIONS):
            perm = grp.copy()
            for c in np.unique(cell):
                idx = np.where(cell == c)[0]
                perm[idx] = rng.permutation(grp[idx])
            if f_stat(y, cell, perm) >= F - 1e-12:
                count += 1
        p_rows = (count + 1) / (N_PERMUTATIONS + 1)
        # (b) permute group labels across the 8 names (name level, exact)
        name_idx = np.array([name_index[n] for n in g["Name"]])
        count = sum(f_stat(y, cell, np.array(a)[name_idx]) >= F - 1e-12 for a in assignments)
        p_names = count / len(assignments)
        rows.append(dict(Modell=model, Bedingung=cond, F=F, p_perm_Zeilen=p_rows, p_perm_Namen=p_names))
    return pd.DataFrame(rows)


# ------------------------------------------------------------------------- main
if __name__ == "__main__":
    data = load_data(INPUT)
    omnibus, diffs, gender = main_tables(data)
    pooled = pooled_table(data)
    gender_holm = gender_holm_table(gender, omnibus, pooled)
    desc, names, spread = descriptive_tables(data)
    perm = permutation_tests(data)
    omnibus = omnibus.merge(perm[["Modell", "Bedingung", "p_perm_Zeilen", "p_perm_Namen"]],
                            on=["Modell", "Bedingung"], how="left")

    pd.set_option("display.width", 250, "display.max_columns", 30, "display.max_rows", 200)
    print("\nOMNIBUS TESTS (p_Holm over all testable tests)")
    print(omnibus.round(3).to_string())
    print("\nDIFFERENCES TO THE GERMAN GROUP (score points)")
    print(diffs.round(2).to_string())
    print("\nGENDER (male minus female)")
    print(gender.round(3).to_string())
    print("\nGENDER, SUPPLEMENTARY HOLM CORRECTION (13 model-specific + 2 pooled tests)")
    print(gender_holm.round(3).to_string())
    print("\nPOOLED")
    print(pooled.round(3).T.to_string())
    print("\nMEAN RANGE BETWEEN THE TWO NAMES OF A GROUP")
    print(spread.round(2).to_string())

    with pd.ExcelWriter(OUTPUT) as w:
        desc.round(2).to_excel(w, sheet_name="Deskriptiv", index=False)
        omnibus.round(4).to_excel(w, sheet_name="Omnibus", index=False)
        diffs.round(3).to_excel(w, sheet_name="Diff_zu_Deutsch", index=False)
        gender.round(3).to_excel(w, sheet_name="Geschlecht", index=False)
        gender_holm.round(4).to_excel(w, sheet_name="Geschlecht_Holm", index=False)
        pooled.round(3).to_excel(w, sheet_name="Gepoolt", index=False)
        names.round(2).to_excel(w, sheet_name="Namen", index=False)
        spread.round(2).to_excel(w, sheet_name="Namensspanne", index=False)
    print(f"\nResults written to {OUTPUT}")
