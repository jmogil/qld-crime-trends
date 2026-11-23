#!/usr/bin/env python3
"""
Build data/custody.json for the Custody tab: people in custody in Queensland.

Adults (ABS, Prisoners in Australia; prisoners at 30 June each year, 2000 onwards), assembled from several releases,
later releases taking precedence for overlapping years:
  --current  data cube "2. Prisoner characteristics, States and territories (Tables 15-35)" of the latest release
             (ten years; Tables 16 and 18)
  --earlier  folder with
             45170do002_2010.xlsx (2000-2010; Tables 4, 5) and 45170do003_2010.xlsx (Indigenous rates; Table 3)
             45170do002_2015.xlsx (2005-2015; Tables 14, 18, 19)   (.xls originals converted to .xlsx)
             abs_prisoner_characteristics_<year>.xlsx: the same data cube as --current from earlier releases (2025 on),
             kept so years that drop out of the latest ten-year table remain covered
Young people (AIHW, Youth detention population in Australia; average night in each quarter, June quarter 2021 onwards):
  AIHW-juv_*_supplementary-tables-*.xlsx  (Tables S1-S18, S29-S33)
Population for derived figures: ABS 3101053.xlsx (Queensland ERP by single year of age and sex, 30 June).

Derived figures (flagged in the output):
  adults 2000-2015: Indigenous and non-Indigenous numbers from the published % Indigenous; 2000-2010 male and female
  numbers from the published % female; 2011-2015 male and female numbers from the published sex-specific rates and
  ABS adult population, scaled to the published total.
  young people: rates by sex, from AIHW counts and ABS population aged 10-17 of the same sex.

Requires: pandas, numpy, openpyxl.
"""
import argparse, glob, json, os, re, sys
import numpy as np
import openpyxl
import pandas as pd


def log(*a):
    print(*a, file=sys.stderr, flush=True)


def num(c):
    try:
        return float(str(c).replace(",", "").strip())
    except (TypeError, ValueError):
        return None


def rows(path, sheet):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    sh = next(s for s in wb.sheetnames if re.sub(r"[^0-9A-Za-z]", "", s).lower() == re.sub(r"[^0-9A-Za-z]", "", sheet).lower())
    return [list(r) for r in wb[sh].iter_rows(values_only=True)]


def year_block(R, start_pat):
    """rows 'YYYY ...' following the first row whose label matches start_pat, until the next non-year label"""
    i = next(k for k, r in enumerate(R) if r[0] and re.search(start_pat, str(r[0]), re.I))
    out = {}
    for r in R[i + 1:]:
        lab = str(r[0]).strip() if r[0] is not None else ""
        m = re.match(r"^(\d{4})", lab)
        if not m:
            if lab:
                break
            continue
        out[int(m.group(1))] = r
    return out


def col_of(R, name, hdr_rows=12):
    for r in R[:hdr_rows]:
        for j, c in enumerate(r):
            if c is not None and str(c).strip().rstrip(".") in (name, name + "."):
                return j
    raise KeyError(name)


def qld_erp(path):
    sheets = pd.read_excel(path, sheet_name=None, header=None)
    series = {"Persons": {}, "Male": {}, "Female": {}}
    years = None
    for name, sh in sheets.items():
        if not name.startswith("Data"):
            continue
        hdr, data = sh.iloc[0], sh.iloc[10:]
        years = [d.year for d in pd.to_datetime(data.iloc[:, 0])]
        for c in range(1, sh.shape[1]):
            parts = [p.strip() for p in str(hdr.iloc[c]).split(";")]
            if len(parts) >= 3 and parts[1] in series:
                a = 100 if parts[2].startswith("100") else int(parts[2])
                series[parts[1]][a] = data.iloc[:, c].to_numpy(float)
    out = {}
    for sx, d in series.items():
        arr = np.array([d[a] for a in range(101)])
        out[sx] = {y: arr[:, i] for i, y in enumerate(years)}
    return out


# ------------------------------------------------------------------------------------------------ adults
def adults(cur, e2015, e2010, e2010i, erp):
    S = {k: {} for k in ["total", "male", "female", "ind", "nonind", "r_total", "r_male", "r_female", "r_ind", "r_nonind", "r_ratio", "unsent"]}
    D = {k: set() for k in S}                     # years where the value is derived, not published
    SRC = {}
    # 2000-2010 (2010 release)
    R = rows(e2010, "Table_5"); q = year_block(R, r"QUEENSLAND")
    for y, r in q.items():
        tot, fem, ind, uns = num(r[1]), num(r[3]), num(r[4]), num(r[6])
        S["total"][y] = tot; S["unsent"][y] = uns
        S["female"][y] = round(tot * fem / 100); S["male"][y] = tot - S["female"][y]; D["female"].add(y); D["male"].add(y)
        S["ind"][y] = round(tot * ind / 100); S["nonind"][y] = tot - S["ind"][y]; D["ind"].add(y); D["nonind"].add(y)
        SRC[y] = "2010"
    R = rows(e2010, "Table_4"); c = col_of(R, "Qld")
    for r in R:
        m = re.match(r"^(\d{4})$", str(r[0]).strip()) if r[0] is not None else None
        if m: S["r_total"][int(m.group(1))] = num(r[c])
    R = rows(e2010i, "Table_3"); c = col_of(R, "Qld")
    for y, r in year_block(R, r"^INDIGENOUS").items(): S["r_ind"][y] = num(r[c])
    for y, r in year_block(R, r"^NON-INDIGENOUS").items(): S["r_nonind"][y] = num(r[c])
    for y in S["r_ind"]:
        if S["r_nonind"].get(y): S["r_ratio"][y] = round(S["r_ind"][y] / S["r_nonind"][y], 1)
    # 2005-2015 (2015 release) overrides
    R = rows(e2015, "Table_14"); q = year_block(R, r"^QUEENSLAND")
    for y, r in q.items():
        tot, rm, rf, ind, uns = num(r[1]), num(r[2]), num(r[3]), num(r[6]), num(r[8])
        S["total"][y] = tot; S["r_male"][y] = rm; S["r_female"][y] = rf; S["unsent"][y] = uns
        S["ind"][y] = round(tot * ind / 100); S["nonind"][y] = tot - S["ind"][y]; D["ind"].add(y); D["nonind"].add(y)
        SRC[y] = "2015"
        # numbers by sex are not published in this release: derive from sex-specific rates and ERP 18+, scaled to
        # the published total (checked below against 2016-2025, where both are published)
        pm, pf = erp["Male"][y][18:].sum(), erp["Female"][y][18:].sum()
        m_, f_ = rm * pm / 1e5, rf * pf / 1e5
        if y <= 2010: log("  %d female share: %% female (2010 release) %.1f, from rates %.1f" % (y, S["female"][y] / S["total"][y] * 100 if y in S["female"] else -1, f_ / (m_ + f_) * 100))
        S["male"][y] = round(tot * m_ / (m_ + f_)); S["female"][y] = tot - S["male"][y]
    R = rows(e2015, "Table_18"); c = col_of(R, "Qld")
    for r in R:
        m = re.match(r"^(\d{4})$", str(r[0]).strip()) if r[0] is not None else None
        if m: S["r_total"][int(m.group(1))] = num(r[c])
    R = rows(e2015, "Table_19"); c = col_of(R, "Qld")
    for y, r in year_block(R, r"^ABORIGINAL").items(): S["r_ind"][y] = num(r[c])
    for y, r in year_block(R, r"^NON-INDIGENOUS").items(): S["r_nonind"][y] = num(r[c])
    for y, r in year_block(R, r"^RATIO").items(): S["r_ratio"][y] = num(r[c])
    # 2016 onwards: current-format data cubes, oldest first, so the latest release wins
    first_new = None
    for path in cur:
        R = rows(path, "Table 16"); q = year_block(R, r"^Queensland")
        rel = max(q)
        for y, r in q.items():
            for k, j in (("total", 1), ("male", 2), ("female", 3), ("ind", 4), ("nonind", 5), ("unsent", 7), ("r_male", 8), ("r_female", 9), ("r_total", 10)):
                S[k][y] = num(r[j]); D[k].discard(y)
            SRC[y] = str(rel)
        first_new = min(q) if first_new is None else min(first_new, min(q))
        R = rows(path, "Table 18"); c = col_of(R, "Qld")
        for y, r in year_block(R, r"^Aboriginal").items(): S["r_ind"][y] = num(r[c])
        for y, r in year_block(R, r"^Non-Indigenous").items(): S["r_nonind"][y] = num(r[c])
        for y, r in year_block(R, r"^Ratio").items(): S["r_ratio"][y] = num(r[c])
        log("  %s: Queensland %d-%d" % (os.path.basename(path), min(q), rel))
    assert first_new == 2016, "years between the 2015 release and the earliest current-format cube are missing (%s)" % first_new
    # check the derivation of male/female numbers on years where both are published
    errs = []
    for y in range(2016, max(S["total"]) + 1):
        rm, rf, tot = S["r_male"][y], S["r_female"][y], S["total"][y]
        pm, pf = erp["Male"][min(y, max(erp["Male"]))][18:].sum(), erp["Female"][min(y, max(erp["Female"]))][18:].sum()
        m_, f_ = rm * pm / 1e5, rf * pf / 1e5
        errs.append(abs(round(tot * m_ / (m_ + f_)) - S["male"][y]) / S["male"][y] * 100)
    log("adult male numbers derived from rates vs published, 2016 onwards: mean abs %% diff %.2f, max %.2f" % (np.mean(errs), max(errs)))
    for y in range(2005, 2016):
        D["male"].add(y); D["female"].add(y)
    D["male"] = {y for y in D["male"] if y < 2016}; D["female"] = {y for y in D["female"] if y < 2016}
    Y = list(range(2000, max(S["total"]) + 1))
    arr = lambda k: [S[k].get(y) for y in Y]
    return {"years": Y, "source_release": [SRC.get(y) for y in Y],
            "n": {k: arr(k) for k in ["total", "male", "female", "ind", "nonind"]},
            "rate": {k[2:]: arr(k) for k in ["r_total", "r_male", "r_female", "r_ind", "r_nonind", "r_ratio"]},
            "unsent": arr("unsent"),
            "derived": {k: sorted(D[k]) for k in ["male", "female", "ind", "nonind"]},
            "breaks": [2005, 2016]}


# ------------------------------------------------------------------------------------------------ young people
def aihw_sheet(path, table, blocks=("Male", "Female", "Total")):
    """Queensland column of an AIHW table with row blocks (sex or Indigenous status) x quarter"""
    R = rows(path, "Table " + table)
    hi = next(i for i, r in enumerate(R) if r[1] and str(r[1]).strip() == "Quarter")
    q = [j for j, c in enumerate(R[hi]) if c and str(c).strip() == "Qld"][0]
    out, cur, quarters = {}, None, []
    for r in R[hi + 1:]:
        a, b = (str(r[0]).strip() if r[0] else ""), (str(r[1]).strip() if r[1] else "")
        if a.startswith("Notes"):
            break
        if a:
            cur = next((k for k in blocks if a.replace("–", "-").lower().startswith(k.replace("–", "-").lower()[:6])), None)
        if cur and re.match(r"^(Jun|Sep|Dec|Mar) qtr \d{4}$", b):
            v = r[q]
            out.setdefault(cur, []).append(None if v in (None, "n.p.", "n.a.", "—") and v != "—" else (0.0 if v == "—" else num(v)))
            if cur == blocks[0]:
                quarters.append(b)
    return quarters, out


def youth(path, erp):
    Q = None
    def get(table, blocks=("Male", "Female", "Total")):
        nonlocal Q
        q, d = aihw_sheet(path, table, blocks)
        if Q is None: Q = q
        assert q == Q, (table, q[:2], Q[:2])
        return d
    ages = {"all": ("S11", "S1", "S6"), "10-13": ("S12", "S2", "S7"), "14-17": ("S13", "S3", "S8"), "10-17": ("S14", "S4", "S9"), "18+": ("S15", "S5", "S10")}
    n = {}
    for a, (tt, ti, tn) in ages.items():
        T, I, N = get(tt), get(ti), get(tn)
        n[a] = {"total": T["Total"], "male": T["Male"], "female": T["Female"], "ind": I["Total"], "nonind": N["Total"],
                "ind_male": I["Male"], "ind_female": I["Female"], "nonind_male": N["Male"], "nonind_female": N["Female"]}
    uns = {}
    for a, (tt, ti, tn) in {"all": ("S29", "S19", "S24"), "10-17": ("S32", "S22", "S27")}.items():
        T, I, N = get(tt), get(ti), get(tn)
        uns[a] = {"total": T["Total"], "male": T["Male"], "female": T["Female"], "ind": I["Total"], "nonind": N["Total"]}
    rate = {}
    for a, t in (("10-13", "S16"), ("14-17", "S17"), ("10-17", "S18")):
        d = get(t, ("First Nations", "Non", "Total"))
        rate[a] = {"total": d["Total"], "ind": d["First Nations"], "nonind": d["Non"],
                   "ratio": [None if (x is None or y in (None, 0)) else round(x / y, 1) for x, y in zip(d["First Nations"], d["Non"])]}
    # rates by sex (derived): ABS ERP aged 10-17 by sex; quarter population interpolated between 30 June estimates
    def pop(sx, lo, hi, label):
        qn = {"Mar": 3, "Jun": 6, "Sep": 9, "Dec": 12}
        mon, yr = label.split()[0], int(label.split()[-1])
        t = yr + (qn[mon] - 1.5 - 6) / 12          # quarter midpoint, in years relative to 30 June of yr
        y0 = int(np.floor(t)); f = t - y0
        last = max(erp[sx]); y0 = min(y0, last - 1)
        p0, p1 = erp[sx][y0][lo:hi + 1].sum(), erp[sx][min(y0 + 1, last)][lo:hi + 1].sum()
        return p0 + (p1 - p0) * min(f, 2)
    for a, (lo, hi) in (("10-17", (10, 17)), ("10-13", (10, 13)), ("14-17", (14, 17))):
        for sx, k in (("Male", "male"), ("Female", "female"), ("Persons", "total_check")):
            src = n[a]["total" if k == "total_check" else k]
            rate[a][k] = [None if v is None else round(v / pop(sx, lo, hi, q) * 1e4, 3) for v, q in zip(src, Q)]
        chk = [abs(x - y) / y * 100 for x, y in zip(rate[a]["total_check"], rate[a]["total"]) if x and y]
        log("youth %s: rate from ABS ERP vs AIHW published rate, mean abs %% diff %.2f, max %.2f" % (a, np.mean(chk), max(chk)))
        del rate[a]["total_check"]
    return {"quarters": Q, "n": n, "unsent": uns, "rate": rate}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--current", required=True, help="ABS Prisoners in Australia: 2. Prisoner characteristics, States and territories (latest)")
    ap.add_argument("--earlier", required=True, help="folder with 45170do002_2015.xlsx, 45170do002_2010.xlsx, 45170do003_2010.xlsx")
    ap.add_argument("--aihw", required=True, help="AIHW youth detention supplementary tables (.xlsx)")
    ap.add_argument("--erp", required=True, help="ABS 3101053.xlsx")
    ap.add_argument("--out", default="data/custody.json")
    a = ap.parse_args()
    erp = qld_erp(a.erp)
    E = lambda n: os.path.join(a.earlier, n)
    cur = sorted(glob.glob(E("abs_prisoner_characteristics_*.xlsx"))) + [a.current]
    out = {"adults": adults(cur, E("45170do002_2015.xlsx"), E("45170do002_2010.xlsx"), E("45170do003_2010.xlsx"), erp),
           "youth": youth(a.aihw, erp),
           "sources": {"adults": "ABS, Prisoners in Australia (2010 and 2015 releases and the latest)", "youth": "AIHW, Youth detention population in Australia 2025"}}
    with open(a.out, "w") as fh:
        json.dump(out, fh, separators=(",", ":"))
    log("wrote", a.out, os.path.getsize(a.out), "bytes")


if __name__ == "__main__":
    main()
