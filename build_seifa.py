#!/usr/bin/env python3
"""
Build data/seifa.json for the dashboard: ABS SEIFA 2021 (all four indexes) for Queensland LGAs and QPS police divisions.

SEIFA is a five-yearly Census product, so this script is run when ABS publishes a new SEIFA (next after the 2026
Census), not in the monthly refresh.

Inputs (--raw):
  SEIFA2021_LGA_Indexes.xlsx      ABS, "Local Government Area, Indexes, SEIFA 2021"
  SEIFA2021_SA1_Indexes.xlsx      ABS, "Statistical Area Level 1, Indexes, SEIFA 2021"
  SA1_2021_AUST_SHP_GDA2020.zip   ABS ASGS Edition 3 SA1 boundaries
  LGA_2025_AUST_GDA2020.zip, QPS_DIVISIONS.zip   the boundaries the dashboard maps
  --meta data/meta.json           the dashboard's list of LGAs and divisions (names in display order)

Method
  LGAs: the score and deciles ABS publishes for each 2021 LGA.
  Police divisions (ABS does not publish SEIFA for these): each Queensland SA1 is assigned to the division (and,
  for the comparison, the 2025 LGA) that contains a point inside it; the division score is the mean of its SA1
  scores weighted by SA1 usual resident population. For both geographies the share of residents living in SA1s
  in the lowest national quintile (deciles 1-2) is also reported.

Requires: pandas, openpyxl, shapely, pyshp.
"""
import argparse, json, os, re, sys, zipfile, io
import numpy as np
import pandas as pd
import shapefile
from shapely.geometry import shape, Point
from shapely.strtree import STRtree

IDX = ["IRSD", "IRSAD", "IER", "IEO"]
FULL = {"IRSD": "Index of Relative Socio-economic Disadvantage", "IRSAD": "Index of Relative Socio-economic Advantage and Disadvantage",
        "IER": "Index of Economic Resources", "IEO": "Index of Education and Occupation"}


def log(*a):
    print(*a, file=sys.stderr, flush=True)


def table1(path):
    """Table 1 of a SEIFA 'Indexes' workbook -> DataFrame with code, [name], score/decile per index, urp."""
    raw = pd.read_excel(path, sheet_name="Table 1", header=None)
    hdr_i = next(i for i in range(12) if str(raw.iloc[i, 0]).startswith("2021 "))
    top = raw.iloc[hdr_i - 1].ffill()
    cols = []
    for j, (a, b) in enumerate(zip(top, raw.iloc[hdr_i])):
        a, b = str(a), str(b)
        k = next((k for k, v in FULL.items() if a.lower().startswith(v.lower())), None)
        if j == 0: cols.append("code")
        elif b.startswith("2021 "): cols.append("name")
        elif k and b in ("Score", "Decile"): cols.append(f"{k}_{b.lower()}")
        elif b.startswith("Usual Resident"): cols.append("urp")
        else: cols.append(f"x{j}")
    df = raw.iloc[hdr_i + 1:].copy(); df.columns = cols
    df = df[df["code"].astype(str).str.fullmatch(r"\d+")].copy()
    df["code"] = df["code"].astype(str)
    for c in df.columns:
        if c not in ("code", "name") and not c.startswith("x"):
            df[c] = pd.to_numeric(df[c], errors="coerce")
    assert all(f"{k}_score" in df for k in IDX), cols
    return df[[c for c in df.columns if not c.startswith("x")]]


def state_deciles(path):
    """Tables 2-5: decile within state, per index (LGA workbook)."""
    out = {}
    for t, k in zip((2, 3, 4, 5), IDX):
        raw = pd.read_excel(path, sheet_name=f"Table {t}", header=None)
        title = str(raw.iloc[3, 0])
        assert FULL[k].lower() in title.lower(), (t, title)
        hdr_i = next(i for i in range(12) if str(raw.iloc[i, 0]).startswith("2021 "))
        h = [str(x) for x in raw.iloc[hdr_i]]
        st = h.index("State"); dec = st + 2
        assert h[dec] == "Decile", h
        d = raw.iloc[hdr_i + 1:]
        d = d[d[0].astype(str).str.fullmatch(r"\d+")]
        out[k] = dict(zip(d[0].astype(str), pd.to_numeric(d[dec], errors="coerce")))
    return out


def norm(s):
    s = re.sub(r"\(.*?\)", "", str(s)).lower()
    s = re.sub(r"\b(aboriginal shire|regional|city|shire|town|council)\b", "", s)
    return re.sub(r"[^a-z]", "", s)


def read_layer(zpath, keep=None):
    with zipfile.ZipFile(zpath) as z:
        stem = next(n[:-4] for n in z.namelist() if n.lower().endswith(".shp"))
        shp, dbf, shx = (io.BytesIO(z.read(stem + e)) for e in (".shp", ".dbf", ".shx"))
    r = shapefile.Reader(shp=shp, dbf=dbf, shx=shx, encoding="latin1")
    fields = [f[0] for f in r.fields[1:]]
    out = []
    for sr in r.iterShapeRecords():
        rec = dict(zip(fields, sr.record))
        if keep and not keep(rec):
            continue
        if not sr.shape.points:
            continue
        out.append((rec, shape(sr.shape.__geo_interface__).buffer(0)))
    return fields, out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--raw", required=True)
    ap.add_argument("--meta", required=True)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    P = lambda n: os.path.join(a.raw, n)
    meta = json.load(open(a.meta))
    lgas, divs = meta["lgas"], meta["divisions"]

    sa1 = table1(P("SEIFA2021_SA1_Indexes.xlsx")); sa1 = sa1[sa1["code"].str.startswith("3")].set_index("code")
    lga = table1(P("SEIFA2021_LGA_Indexes.xlsx")); lga = lga[lga["code"].str.startswith("3")].set_index("code")
    sdec = state_deciles(P("SEIFA2021_LGA_Indexes.xlsx"))
    log("Queensland SA1s with SEIFA:", len(sa1), "| LGAs:", len(lga))

    # match the dashboard's (2025) LGA names to the 2021 SEIFA LGAs
    by_norm = {norm(n): c for c, n in lga["name"].items()}
    lga_code = []
    for n in lgas:
        c = by_norm.get(norm(n))
        lga_code.append(c)
        if c is None:
            log("  no 2021 SEIFA LGA for", n)

    # geometry: SA1 points -> divisions and LGAs
    _, s1 = read_layer(P("SA1_2021_AUST_SHP_GDA2020.zip"), keep=lambda r: str(r.get("STE_CODE21")) == "3")
    pts = {str(r["SA1_CODE21"]): g.representative_point() for r, g in s1}
    log("Queensland SA1 polygons:", len(pts))
    dfields, dlay = read_layer(P("QPS_DIVISIONS.zip"))
    dname_f = next(f for f in dfields if f.upper() in ("DIVISION", "DIV_NAME", "NAME", "DIVISION_N"))
    lfields, llay = read_layer(P("LGA_2025_AUST_GDA2020.zip"), keep=lambda r: str(r.get("STE_CODE21")) == "3")
    dpos = {n.upper(): i for i, n in enumerate(divs)}
    lpos = {norm(n): i for i, n in enumerate(lgas)}

    def assign(layer, key):
        geoms = [g for _, g in layer]; idx = [key(r) for r, _ in layer]
        tree = STRtree(geoms)
        res = {}
        for code, p in pts.items():
            for j in tree.query(p):
                if geoms[j].contains(p):
                    res[code] = idx[j]; break
        return res, geoms, idx

    d_of, dgeoms, didx = assign(dlay, lambda r: dpos.get(str(r[dname_f]).strip().upper()))
    l_of, _, _ = assign(llay, lambda r: lpos.get(norm(r["LGA_NAME25"])))
    log("SA1s placed: division", sum(v is not None for v in d_of.values()), "| LGA", sum(v is not None for v in l_of.values()), "of", len(pts))
    unmatched = [str(r[dname_f]) for r, _ in dlay if dpos.get(str(r[dname_f]).strip().upper()) is None]
    if unmatched:
        log("  division polygons not in meta:", unmatched[:10])

    s = sa1.copy()
    s["div"] = [d_of.get(c) for c in s.index]; s["lga"] = [l_of.get(c) for c in s.index]
    s = s[s["urp"] > 0]
    have = s.dropna(subset=["IRSD_score"])
    log("SA1s with a score:", len(have), "of", len(s), "| residents in SA1s without a score: %.2f%%" % (100 * s.loc[s["IRSD_score"].isna(), "urp"].sum() / s["urp"].sum()))

    def agg(group_col, n):
        score = [[None] * len(IDX) for _ in range(n)]; q1 = [[None] * len(IDX) for _ in range(n)]
        pop = [0] * n; nsa1 = [0] * n
        for gi, g in have.dropna(subset=[group_col]).groupby(group_col):
            gi = int(gi); w = g["urp"].to_numpy(float)
            pop[gi] = int(w.sum()); nsa1[gi] = len(g)
            for k, ix in enumerate(IDX):
                sc = g[f"{ix}_score"].to_numpy(float); dc = g[f"{ix}_decile"].to_numpy(float)
                score[gi][k] = round(float((sc * w).sum() / w.sum()), 1)
                q1[gi][k] = round(float(100 * w[dc <= 2].sum() / w.sum()), 1)
        return score, q1, pop, nsa1

    d_score, d_q1, d_pop, d_n = agg("div", len(divs))
    # divisions with no SA1 point inside: use the SA1 that contains the division's own interior point
    sa1_geoms = {str(r["SA1_CODE21"]): g for r, g in s1}
    codes = list(sa1_geoms); stree = STRtree([sa1_geoms[c] for c in codes])
    filled = []
    for gi in range(len(divs)):
        if d_score[gi][0] is not None:
            continue
        j = didx.index(gi) if gi in didx else None
        if j is None:
            continue
        p = dgeoms[j].representative_point()
        hit = next((codes[k] for k in stree.query(p) if sa1_geoms[codes[k]].contains(p)), None)
        if hit and hit in have.index:
            row = have.loc[hit]
            d_score[gi] = [round(float(row[f"{ix}_score"]), 1) for ix in IDX]
            d_q1[gi] = [100.0 if row[f"{ix}_decile"] <= 2 else 0.0 for ix in IDX]
            filled.append(divs[gi])
    log("divisions with no SA1 inside, filled from the SA1 at their centre:", len(filled), filled[:8])
    log("divisions still without SEIFA:", [divs[i] for i in range(len(divs)) if d_score[i][0] is None])

    l_mean, l_q1, l_pop, l_n = agg("lga", len(lgas))
    # check: SA1-weighted means against the published LGA scores
    pub = np.array([[lga.loc[c, f"{ix}_score"] if c else np.nan for ix in IDX] for c in lga_code], float)
    est = np.array([[np.nan if v is None else v for v in r] for r in l_mean], float)
    for k, ix in enumerate(IDX):
        m = ~np.isnan(pub[:, k]) & ~np.isnan(est[:, k])
        log("%s: published LGA score vs SA1-weighted mean, r = %.3f, mean abs diff %.1f (n=%d)" % (ix, np.corrcoef(pub[m, k], est[m, k])[0, 1], np.abs(pub[m, k] - est[m, k]).mean(), m.sum()))

    def rec(c, k, field):
        if c is None or c not in lga.index: return None
        v = lga.loc[c, f"{IDX[k]}_{field}"]
        return None if pd.isna(v) else round(float(v), 1) if field == "score" else int(v)

    out = {
        "source": "ABS, Socio-Economic Indexes for Areas (SEIFA), Australia, 2021 (released 2023)",
        "indexes": [{"id": ix, "name": FULL[ix]} for ix in IDX],
        "l": {"score": [[rec(c, k, "score") for k in range(4)] for c in lga_code],
              "dec_aus": [[rec(c, k, "decile") for k in range(4)] for c in lga_code],
              "dec_qld": [[(None if c is None or pd.isna(sdec[ix].get(c)) else int(sdec[ix][c])) for ix in IDX] for c in lga_code],
              "q1": l_q1, "urp": [None if c is None else int(lga.loc[c, "urp"]) for c in lga_code], "name21": [None if c is None else lga.loc[c, "name"] for c in lga_code]},
        "d": {"score": d_score, "q1": d_q1, "urp": d_pop, "nsa1": d_n, "filled": filled},
    }
    with open(a.out, "w") as fh:
        json.dump(out, fh, separators=(",", ":"))
    log("wrote", a.out, os.path.getsize(a.out), "bytes")


if __name__ == "__main__":
    main()
