#!/usr/bin/env python3
"""
Build data/victims.json for the Victims tab: Queensland figures from ABS Recorded Crime - Victims.

Inputs
  --victims DIR   folder with two sub-folders:
        current/   the workbooks of the latest ABS release (cubes 2, 3 and 5 are used):
                     "2. Victims of crime, states and territories (Tables 9 to 16).xlsx"
                     "3. Victims of crime, relationship of offender to victims (Tables 17 to 22).xlsx"
                     "5. Victims of family and domestic violence related offences (Tables 29 to 38).xlsx"
        earlier/   one workbook per earlier reference year holding that year's state age-by-sex table,
                   named victims_<year>.xlsx (2010 onwards; .xls originals converted to .xlsx)
  --erp FILE      ABS 3101053.xlsx (Queensland ERP by single year of age and sex)
  --out FILE      output JSON (default data/victims.json)

When ABS publishes a new release: move the old "2. Victims of crime, states and territories" workbook from
current/ to earlier/ as victims_<its year>.xlsx, put the new workbooks in current/, and re-run.

Requires: pandas, numpy, openpyxl.
"""
import argparse, glob, json, os, re, sys
import numpy as np
import openpyxl
import pandas as pd

STATES = ["New South Wales", "Victoria", "Queensland", "South Australia", "Western Australia", "Tasmania",
          "Northern Territory", "Australian Capital Territory", "Australia"]
SEXKEY = {"males": "Male", "females": "Female", "persons": "Persons"}


def log(*a):
    print(*a, file=sys.stderr, flush=True)


def s(c):
    return "" if c is None else str(c).strip()


def num(c):
    t = s(c).replace(",", "")
    try:
        return float(t)
    except ValueError:
        return None


def clean(t):
    """strip ABS footnote markers such as (a)(b)"""
    return re.sub(r"\s+", " ", re.sub(r"\([a-z]{1,2}\)", "", s(t))).strip()


def rows(wb, sheet):
    return [list(r) for r in wb[sheet].iter_rows(values_only=True)]


def sheet_no(wb, n):
    return [x for x in wb.sheetnames if re.sub(r"\D", "", x) == str(n)][0]


def state_of(r):
    for c in r[:2]:
        t = clean(c)
        if t in STATES:
            return t
    return None


def qld_block(R, start=0):
    """rows of the Queensland block (between the Queensland heading and the next state heading)"""
    q = next(i for i in range(start, len(R)) if state_of(R[i]) == "Queensland" and num(R[i][2] if len(R[i]) > 2 else None) is None)
    e = next((i for i in range(q + 1, len(R)) if state_of(R[i]) and state_of(R[i]) != "Queensland"), len(R))
    return R[q + 1:e]


def footnotes(R):
    out = {}
    for r in R:
        m = re.match(r"^\(([a-z]{1,2})\)\s*(.+)$", s(r[0]))
        if m:
            out[m.group(1)] = m.group(2).strip()
    return out


def ivals(r, a, b):
    return [None if num(c) is None else (int(num(c)) if float(num(c)).is_integer() else num(c)) for c in r[a:b]]


# ----------------------------------------------------------------------------- offence trends (Tables 9, 10)
OFF_TREE = [("Homicide", 0), ("Murder", 1), ("Attempted murder", 1), ("Manslaughter", 1), ("Assault", 0), ("Sexual assault", 0),
            ("Kidnapping/abduction", 0), ("Robbery", 0), ("Armed robbery", 1), ("Unarmed robbery", 1), ("Blackmail/other extortion", 0),
            ("Burglary", 0), ("Involving the taking of property", 1), ("Other", 1), ("Motor vehicle theft", 0), ("Other theft", 0)]
PERSON_ONLY = {"Homicide", "Murder", "Attempted murder", "Manslaughter", "Assault", "Sexual assault", "Kidnapping/abduction"}


def offence_trends(wb2, erp_total):
    R9, R10 = rows(wb2, sheet_no(wb2, 9)), rows(wb2, sheet_no(wb2, 10))
    hdr = [s(c) for c in R9[4]]
    ycols = [(i, int(h[:4]), re.findall(r"\(([a-z]{1,2})\)", h)) for i, h in enumerate(hdr) if re.match(r"^\d{4}", h)]
    years = [y for _, y, _ in ycols]
    fn = footnotes(R9)
    other_states = [x for x in STATES if x not in ("Queensland", "Australia")]
    breaks = [{"year": y, "note": fn[m]} for _, y, marks in ycols for m in marks
              if m in fn and ("Queensland" in fn[m] or not any(st in fn[m] for st in other_states))]
    def block(R):
        out = {}
        for r in qld_block(R):
            lab = clean(r[0])
            if lab:
                out[lab] = r
        return out
    b9, b10 = block(R9), block(R10)
    hdr10 = [s(c) for c in R10[4]]
    y10 = {int(h[:4]): i for i, h in enumerate(hdr10) if re.match(r"^\d{4}", h)}
    offs = []
    parent = None
    for name, lvl in OFF_TREE:
        if lvl == 0:
            parent = name
        r = b9[name]
        counts = [None if num(r[i]) is None else int(num(r[i])) for i, _, _ in ycols]
        rate_abs = [None if (name not in b10 or y not in y10 or num(b10[name][y10[y]]) is None) else num(b10[name][y10[y]]) for y in years]
        rate_calc = [None if c is None else round(c / erp_total[y] * 1e5, 2) for c, y in zip(counts, years)]
        label = name if lvl == 0 else {"Involving the taking of property": "Burglary involving the taking of property", "Other": "Other burglary"}.get(name, name)
        offs.append({"name": label, "lvl": lvl, "parent": parent if lvl else None, "person_only": name in PERSON_ONLY,
                     "counts": counts, "rate_abs": rate_abs, "rate_calc": rate_calc})
    notes = [fn[k] for k in fn if "Queensland" in fn[k]]
    return {"years": years, "offences": offs, "breaks": breaks, "qld_notes": notes}


# ----------------------------------------------------------------------------- age by sex, one release per year
def norm_off(t):
    t = clean(t).lower().replace(" /", "/").replace("/ ", "/")
    for key, name in (("homicide", "Homicide"), ("sexual assault", "Sexual assault"), ("kidnapping", "Kidnapping/abduction"),
                      ("robbery", "Robbery"), ("blackmail", "Blackmail/other extortion"), ("assault", "Assault")):
        if t.startswith(key):
            return name
    return None


def norm_age(t):
    t = clean(t).replace("–", "-").replace(" years", "").replace(" and over", "+").replace(" ", "")
    return t.replace("Under18", "0-17")


def is_age(lab):
    return bool(re.match(r"^(\d+\s*[–-]\s*\d+|\d+ years and over|Under 18|18 years and over|Total)", lab))


def extract_age(path, year):
    """[(offence, sex, age label, count, ABS rate or None)] for Queensland from one release's state age table."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    out = []
    if year <= 2012:                                   # one sheet per state, offences in columns
        R = rows(wb, sheet_no(wb, 3))
        assert "Queensland" in s(R[3][0]), s(R[3][0])
        hdr = [norm_off(c) for c in R[4]]
        sex = None
        for r in R[5:]:
            lab = s(r[0])
            if lab.lower() in SEXKEY:
                sex = SEXKEY[lab.lower()]
                continue
            if sex and is_age(lab):
                for j, o in enumerate(hdr):
                    if o and num(r[j]) is not None:
                        out.append((o, sex, norm_age(lab), num(r[j]), None))
        return out
    R = rows(wb, sheet_no(wb, 7 if year <= 2019 else 11))
    hrow = next(i for i, r in enumerate(R[:12]) if any(clean(c) in ("Qld", "MALES") for c in r))
    if any(s(c) == "MALES" for c in R[hrow]):          # sex x offence in columns, states as row blocks
        offrow = [norm_off(c) for c in R[hrow + 1]]
        cur, colsex = None, []
        for c in R[hrow]:
            if s(c).lower() in SEXKEY:
                cur = SEXKEY[s(c).lower()]
            colsex.append(cur)
        inq = False
        for r in R[hrow + 2:]:
            lab, st = s(r[0]), s(r[1])
            if lab == "" and st and num(r[1]) is None and st.isupper():
                inq = st == "QUEENSLAND"
                continue
            if inq and is_age(lab):
                for j, o in enumerate(offrow):
                    if o and colsex[j] and num(r[j]) is not None:
                        out.append((o, colsex[j], norm_age(lab), num(r[j]), None))
        return out
    hdr = [clean(c) for c in R[hrow]]                  # offences as row blocks, states in columns
    qcols = [i for i, h in enumerate(hdr) if h == "Qld"]
    q, qr = qcols[0], (qcols[1] if len(qcols) > 1 else None)
    off = sex = None
    for r in R[hrow + 1:]:
        lab, c1 = s(r[0]), s(r[1])
        if lab == "" and c1 and num(r[1]) is None:
            off, sex = norm_off(c1), None
            continue
        if clean(lab).lower() in SEXKEY:
            sex = SEXKEY[clean(lab).lower()]
            continue
        if off and sex and is_age(lab) and num(r[q]) is not None:
            out.append((off, sex, norm_age(lab), num(r[q]), num(r[qr]) if qr is not None else None))
    return out


# harmonised age groups: each is a list of alternative recipes (the first whose parts are all present is used)
H7 = [("0-9", [["0-9"]]), ("10-24", [["10-17", "18-24"], ["10-14", "15-19", "20-24"]]), ("25-34", [["25-34"]]), ("35-44", [["35-44"]]),
      ("45-54", [["45-54"]]), ("55-64", [["55-64"]]), ("65+", [["65+"]])]
H3 = [("0-34", [["0-17", "18-34"], ["0-19", "20-34"], ["0-9", "10-14", "15-19", "20-24", "25-34"]]),
      ("35-54", [["35-54"], ["35-44", "45-54"]]), ("55+", [["55+"], ["55-64", "65+"]])]
P9 = ["0-9", "10-17", "18-24", "25-34", "35-44", "45-54", "55-64", "65+"]
SCHEME = {"Sexual assault": "H7", "Robbery": "H7", "Assault": "H7", "Homicide": "H3", "Kidnapping/abduction": "H3", "Blackmail/other extortion": "H3"}
BAND_RANGE = {"0-9": (0, 9), "10-24": (10, 24), "25-34": (25, 34), "35-44": (35, 44), "45-54": (45, 54), "55-64": (55, 64), "65+": (65, 100),
              "0-34": (0, 34), "35-54": (35, 54), "55+": (55, 100), "10-17": (10, 17), "18-24": (18, 24), "0-17": (0, 17), "18-34": (18, 34)}


def harmonise(cells, scheme):
    """cells: {age label: count}. Returns list of counts per harmonised band (None where not derivable)."""
    out = []
    for _, recipes in (H7 if scheme == "H7" else H3):
        v = None
        for parts in recipes:
            if all(p in cells for p in parts):
                v = int(sum(cells[p] for p in parts))
                break
        out.append(v)
    return out


def qld_erp(path):
    """{'Persons'|'Male'|'Female': {year: np.array(ages 0..100)}} at 30 June."""
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
    for sex, d in series.items():
        arr = np.array([d[a] for a in range(101)])
        out[sex] = {y: arr[:, i] for i, y in enumerate(years)}
    return out


def age_section(vdir, wb2, wb3, erp, latest):
    years = list(range(2010, latest + 1))
    raw = {}
    for y in years:
        if y == latest:
            raw[y] = extract_age(glob.glob(os.path.join(vdir, "current", "2.*.xlsx"))[0], y)
        else:
            p = os.path.join(vdir, "earlier", f"victims_{y}.xlsx")
            raw[y] = extract_age(p, y) if os.path.exists(p) else []
    # sexual assault 2014 onwards in the current release (Table 21, 'Total' relationship block): revised, consistent age groups
    R21 = rows(wb3, sheet_no(wb3, 21))
    grp = [(i, clean(c)) for i, c in enumerate(R21[4]) if c]
    tot0 = next(i for i, g in grp if g.startswith("Total") and "known" not in g.lower())
    y21 = [int(s(c)[:4]) for c in R21[5][tot0:tot0 + 30] if re.match(r"^\d{4}", s(c))]
    sa21 = {}
    sex = None
    for r in qld_block(R21):
        lab = clean(r[0])
        if lab.lower() in SEXKEY:
            sex = SEXKEY[lab.lower()]
            continue
        if sex and is_age(s(r[0])):
            for k, y in enumerate(y21):
                v = num(r[tot0 + k])
                if v is not None:
                    sa21.setdefault(y, {}).setdefault(sex, {})[norm_age(s(r[0]))] = v
    def cells(y, off, sex):
        if off == "Sexual assault" and y in sa21 and sex in sa21[y]:
            return sa21[y][sex]
        return {a: n for o, sx, a, n, _ in raw[y] if o == off and sx == sex}
    def popband(sex, band, y):
        a, b = BAND_RANGE[band]
        yy = min(y, max(erp[sex]))
        return float(erp[sex][yy][a:b + 1].sum())
    trend, published = {}, {}
    for off, scheme in SCHEME.items():
        bands = [b for b, _ in (H7 if scheme == "H7" else H3)]
        t = {"scheme": scheme, "bands": bands}
        for sex in ("Persons", "Male", "Female"):
            mat, tot = [], []
            for y in years:
                c = cells(y, off, sex)
                mat.append(harmonise(c, scheme) if c else [None] * len(bands))
                tot.append(int(c["Total"]) if "Total" in c else None)
            t[sex] = {"n": [[mat[k][b] for k in range(len(years))] for b in range(len(bands))], "total": tot}
        trend[off] = t
        if scheme == "H7":                                 # published groups, for the years that have them
            p = {}
            for sex in ("Persons", "Male", "Female"):
                p[sex] = [[(int(cells(y, off, sex)[b]) if b in cells(y, off, sex) else None) for y in years] for b in P9]
            published[off] = p
    pop = {sc: {sex: [[round(popband(sex, b, y)) for y in years] for b in [x for x, _ in (H7 if sc == "H7" else H3)]] for sex in ("Persons", "Male", "Female")} for sc in ("H7", "H3")}
    pop["P9"] = {sex: [[round(popband(sex, b, y)) for y in years] for b in P9] for sex in ("Persons", "Male", "Female")}
    # profile for the latest year, as published (with ABS rates)
    prof = {}
    order = {}
    for o, sx, a, n, rate in raw[latest]:
        prof.setdefault(o, {}).setdefault(sx, []).append({"band": a, "n": int(n), "rate": rate})
    # checks
    chk = []
    for o, sx, a, n, rate in raw[latest]:
        if rate and a in BAND_RANGE and n >= 100:
            mine = n / popband(sx, a, latest) * 1e5
            chk.append(abs(mine - rate) / rate)
    if chk:
        log("age rates: my rate vs ABS published rate, latest year, cells with n>=100: median diff %.2f%%, max %.2f%%" % (np.median(chk) * 100, max(chk) * 100))
    for y in y21:
        a = {x: n for o, sx, x, n, _ in raw[y] if o == "Sexual assault" and sx == "Persons"}
        if "Total" in a and "Total" in sa21[y]["Persons"] and abs(a["Total"] - sa21[y]["Persons"]["Total"]) > 30:
            log(f"  note: sexual assault total {y}: first published {a['Total']:.0f}, current release {sa21[y]['Persons']['Total']:.0f}")
    return {"years": years, "trend": trend, "published": published, "published_bands": P9, "pop": pop, "profile": prof, "profile_year": latest}


# ----------------------------------------------------------------------------- relationship (Tables 17, 18)
def relationship(wb3):
    R = rows(wb3, sheet_no(wb3, 17))
    hdr = [s(c) for c in R[4]]
    ycols = [(i, int(h[:4])) for i, h in enumerate(hdr) if re.match(r"^\d{4}", h)]
    n = len(ycols) // 2
    ycols = ycols[:n]                                   # number block (the proportion block follows)
    years = [y for _, y in ycols]
    out, off = {}, None
    for r in qld_block(R):
        lab = clean(r[0])
        if not lab:
            continue
        if all(num(r[i]) is None and s(r[i]) != "np" for i, _ in ycols):
            off = lab
            out[off] = {}
            continue
        key = {"Known to victim": "known", "Family member": "family", "Non-family member": "nonfamily", "Stranger": "stranger", "Total": "total"}.get(lab)
        if off and key:
            out[off][key] = [None if num(r[i]) is None else int(num(r[i])) for i, _ in ycols]
    R = rows(wb3, sheet_no(wb3, 18))
    grp = [(i, clean(c)) for i, c in enumerate(R[4]) if c]
    y18 = [int(s(c)[:4]) for c in R[5][grp[0][0]:grp[1][0]] if re.match(r"^\d{4}", s(c))]
    keymap = {}
    for i, g in grp:
        gl = g.lower()
        keymap["family" if gl.startswith("family") else "nonfamily" if gl.startswith("non-family") else "known" if "known" in gl else "stranger" if gl.startswith("stranger") else "total"] = i
    bysex, off = {}, None
    for r in qld_block(R):
        lab = clean(r[0])
        if not lab:
            continue
        if lab.lower() in SEXKEY:
            if off:
                bysex[off][SEXKEY[lab.lower()]] = {k: [None if num(c) is None else int(num(c)) for c in r[i:i + len(y18)]] for k, i in keymap.items()}
        else:
            off = lab
            bysex[off] = {}
    return {"years": years, "offences": out, "years_sex": y18, "by_sex": bysex}


# ----------------------------------------------------------------------------- family and domestic violence (Table 30)
def fdv(wb5):
    R = rows(wb5, sheet_no(wb5, 30))
    grp = [(i, clean(c)) for i, c in enumerate(R[3]) if c]
    n0, r0, p0 = grp[0][0], grp[1][0], grp[2][0]
    years = [int(s(c)[:4]) for c in R[4][n0:r0] if re.match(r"^\d{4}", s(c))]
    k = len(years)
    out, sex = {}, None
    for r in qld_block(R):
        lab = clean(r[0])
        if not lab:
            continue
        if lab.lower() in SEXKEY:
            sex = SEXKEY[lab.lower()]
            continue
        if sex:
            out.setdefault(lab, {})[sex] = {"n": ivals(r, n0, n0 + k), "rate": [num(c) for c in r[r0:r0 + k]], "share": [num(c) for c in r[p0:p0 + k]]}
    fn = footnotes(R)
    defn = next((s(r[0]) for r in R if s(r[0]).startswith("Victims of selected offences have been determined to be FDV")), "")
    return {"years": years, "offences": out, "definition": defn, "qld_notes": [v for v in fn.values() if "Queensland" in v]}


# ----------------------------------------------------------------------------- sexual assault: time to report, age at incident
def sexual_assault(wb2):
    R = rows(wb2, sheet_no(wb2, 15))
    grp = [(i, clean(c)) for i, c in enumerate(R[3]) if c]
    n0 = grp[0][0]
    p0 = grp[1][0] if len(grp) > 1 else None
    years = [int(s(c)[:4]) for c in R[4][n0:(p0 or n0 + 40)] if re.match(r"^\d{4}", s(c))]
    k = len(years)
    LEAF = ["Less than 6 months", "6 months to less than 1 year", "1 to less than 2 years", "2 to less than 5 years", "5 to less than 10 years",
            "10 to less than 20 years", "20 years and over"]
    out, sex = {}, None
    for r in qld_block(R):
        lab = clean(r[0])
        if lab.lower() in SEXKEY:
            sex = SEXKEY[lab.lower()]
            out[sex] = {}
            continue
        if sex and (lab in LEAF or lab == "Total"):
            out[sex][lab] = ivals(r, n0, n0 + k)
    R = rows(wb2, sheet_no(wb2, 16))
    hdr = [clean(c) for c in R[4]]
    q = hdr.index("Qld")
    inc, sex = {}, None
    for r in R[5:]:
        lab = clean(r[0])
        if lab.lower() in SEXKEY:
            sex = SEXKEY[lab.lower()]
            inc[sex] = {}
            continue
        if sex and is_age(s(r[0])) and num(r[q]) is not None:
            inc[sex][norm_age(s(r[0]))] = int(num(r[q]))
    return {"years": years, "categories": LEAF, "report": out, "incident": inc}


# ----------------------------------------------------------------------------- latest-year snapshot (Tables 12-14)
TOP = {"location": ["Residential", "Community", "Other location"],
       "weapon": ["Weapon used", "No weapon used"],
       "outcome": ["Investigation not finalised", "Investigation finalised - no offender proceeded against", "Investigation finalised - offender proceeded against"]}


def snapshot(wb2):
    out = {}
    for key, tno in (("location", 12), ("weapon", 13), ("outcome", 14)):
        R = rows(wb2, sheet_no(wb2, tno))
        hdr = [clean(c) for c in R[4]]
        qs = [i for i, h in enumerate(hdr) if h == "Qld"]
        qn, qp = qs[0], qs[1]
        sec, off = {}, None
        for r in R[5:]:
            lab, c1 = clean(r[0]), clean(r[1])
            if lab == "" and c1 and num(r[1]) is None:
                off = {"Burglary - Involving the taking of property": "Burglary involving the taking of property", "Burglary - Other": "Other burglary"}.get(c1, c1)
                sec[off] = []
                continue
            if off is None or not lab or lab == "Footnotes" or len(lab) > 70:
                if lab == "Footnotes":
                    break
                continue
            if num(r[qn]) is None and s(r[qn]) != "np":
                continue
            lab2 = "Other location" if lab.lower() == "other location" else lab
            lab2 = "Other weapon type" if (key == "weapon" and lab2 == "Other") else lab2
            if lab2.startswith("Total"):
                lvl = -1
            elif lab2 in TOP[key]:
                lvl = 0
            else:
                lvl = 2 if (key == "weapon" and lab2 not in ("Firearm", "Other weapon", "Weapon n.f.d.")) else 1
            sec[off].append({"label": lab2, "lvl": lvl, "n": None if num(r[qn]) is None else int(num(r[qn])), "pct": num(r[qp])})
        out[key] = sec
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--victims", required=True)
    ap.add_argument("--erp", required=True)
    ap.add_argument("--out", default="data/victims.json")
    a = ap.parse_args()
    cur = os.path.join(a.victims, "current")
    def load(prefix):
        return openpyxl.load_workbook(glob.glob(os.path.join(cur, prefix + "*.xlsx"))[0], read_only=True, data_only=True)
    wb2, wb3, wb5 = load("2."), load("3."), load("5.")
    erp = qld_erp(a.erp)
    erp_total = {y: float(v.sum()) for y, v in erp["Persons"].items()}
    R = rows(wb2, "Contents")
    release = next((s(r[0]) for r in R if s(r[0]).startswith("Recorded Crime")), "Recorded Crime - Victims")
    released = next((s(r[0]) for r in R if s(r[0]).startswith("Released")), "")
    trends = offence_trends(wb2, erp_total)
    latest = trends["years"][-1]
    out = {"release": release, "released": released, "latest": latest,
           "trends": trends,
           "age": age_section(a.victims, wb2, wb3, erp, latest),
           "relationship": relationship(wb3),
           "fdv": fdv(wb5),
           "sexual_assault": sexual_assault(wb2),
           "snapshot": snapshot(wb2)}
    os.makedirs(os.path.dirname(os.path.abspath(a.out)), exist_ok=True)
    with open(a.out, "w", encoding="utf-8") as fh:
        json.dump(out, fh, separators=(",", ":"), ensure_ascii=False)
    log("wrote", a.out, os.path.getsize(a.out), "bytes;", release, "|", released)


if __name__ == "__main__":
    main()
