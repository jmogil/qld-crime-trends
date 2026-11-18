#!/usr/bin/env python3
"""
Build the data files for the Queensland crime trends dashboard.

Inputs (all in --raw):
  QPS (https://www.police.qld.gov.au/maps-and-statistics):
    QLD_Reported_Offences_Number.csv / _Rates.csv
    LGA_Reported_Offences_Number.csv / _Rates.csv
    division_Reported_Offences_Number.csv / _Rates.csv
    QLD_Reported_Offenders_Number.csv, LGA_Reported_Offenders_Number.csv
    QLD_Reported_victims_Number.xlsx, LGA_Reported_victims_Number.xlsx
    QPS_DIVISIONS.zip (shapefile; or already-extracted QPS_DIVISIONS.shp/.dbf)
  ABS:
    3101053.xlsx            ERP by single year of age, Queensland (National, state and territory population, Table 53)
    32350DS0006_*.xlsx      ERP by age and sex by LGA, 2001 onwards (Regional population by age and sex)
    LGA_2025_AUST_GDA2020.zip  LGA digital boundaries (ASGS Edition 3)

Usage:
  python3 build_data.py --raw RAW_DIR --out site/data            # process files already in RAW_DIR
  python3 build_data.py --raw RAW_DIR --out site/data --download # fetch the QPS + ABS files first
  use  --download-qps instead of --download to fetch only the ten monthly QPS files (ABS and boundary files already in RAW_DIR)
  add  --tidy DIR     to also write long-format analysis files (csv.gz)
  add  --report FILE  to also write a markdown report on the latest month

Requires: pandas, numpy, openpyxl. Geometry handling is pure Python (no GDAL needed).
"""
import argparse, csv, glob, io, json, math, os, struct, sys, zipfile, datetime
import numpy as np
import pandas as pd

QPS = "https://open-crime-data.s3-ap-southeast-2.amazonaws.com/Crime%20Statistics/"
DOWNLOADS = {
    "QLD_Reported_Offences_Number.csv": QPS + "QLD_Reported_Offences_Number.csv",
    "QLD_Reported_Offences_Rates.csv": QPS + "QLD_Reported_Offences_Rates.csv",
    "LGA_Reported_Offences_Number.csv": QPS + "LGA_Reported_Offences_Number.csv",
    "LGA_Reported_Offences_Rates.csv": QPS + "LGA_Reported_Offences_Rates.csv",
    "division_Reported_Offences_Number.csv": QPS + "division_Reported_Offences_Number.csv",
    "division_Reported_Offences_Rates.csv": QPS + "division_Reported_Offences_Rates.csv",
    "QLD_Reported_Offenders_Number.csv": QPS + "QLD_Reported_Offenders_Number.csv",
    "LGA_Reported_Offenders_Number.csv": QPS + "LGA_Reported_Offenders_Number.csv",
    "QLD_Reported_victims_Number.xlsx": QPS + "QLD_Reported_victims_Number.xlsx",
    "LGA_Reported_victims_Number.xlsx": QPS + "LGA_Reported_victims_Number.xlsx",
    "QPS_DIVISIONS.zip": "https://open-crime-data.s3-ap-southeast-2.amazonaws.com/document/QPS_DIVISIONS.zip",
    "3101053.xlsx": "https://www.abs.gov.au/statistics/people/population/national-state-and-territory-population/latest-release/3101053.xlsx",
    # The two ABS files below are release-specific; update the URLs when ABS publishes a new release.
    "32350DS0006_2001-25.xlsx": "https://www.abs.gov.au/statistics/people/population/regional-population-age-and-sex/2025/32350DS0006_2001-25.xlsx",
    "LGA_2025_AUST_GDA2020.zip": "https://www.abs.gov.au/statistics/standards/australian-statistical-geography-standard-asgs/edition-3-july-2021-june-2026/access-and-downloads/digital-boundary-files/LGA_2025_AUST_GDA2020.zip",
}

# Offence hierarchy: (indent level, QPS column name). Verified against column sums in the statewide file.
TREE = [
    (0, "Offences Against the Person"),
    (1, "Homicide (Murder)"),
    (1, "Other Homicide"),
    (2, "Attempted Murder"), (2, "Conspiracy to Murder"), (2, "Manslaughter (excl. by driving)"),
    (2, "Manslaughter Unlawful Striking Causing Death"), (2, "Driving Causing Death"),
    (1, "Assault"),
    (2, "Grievous Assault"), (2, "Serious Assault"), (2, "Serious Assault (Other)"), (2, "Common Assault'"),
    (1, "Sexual Offences"),
    (2, "Rape and Attempted Rape"), (2, "Other Sexual Offences"),
    (1, "Robbery"),
    (2, "Armed Robbery"), (2, "Unarmed Robbery"),
    (1, "Other Offences Against the Person"),
    (2, "Kidnapping & Abduction etc."), (2, "Coercive Control"), (2, "Extortion"), (2, "Stalking"),
    (2, "Life Endangering Acts"), (2, "Voluntary Assisted Dying"), (2, "Other Miscellaneous"),
    (0, "Offences Against Property"),
    (1, "Unlawful Entry"),
    (2, "Unlawful Entry With Intent - Dwelling"),
    (3, "Unlawful Entry Without Violence - Dwelling"), (3, "Unlawful Entry With Violence - Dwelling"),
    (2, "Unlawful Entry With Intent - Shop"), (2, "Unlawful Entry With Intent - Other"),
    (1, "Arson"),
    (1, "Other Property Damage"),
    (1, "Unlawful Use of Motor Vehicle"),
    (1, "Other Theft (excl. Unlawful Entry)"),
    (2, "Stealing from Dwellings"), (2, "Shop Stealing"), (2, "Vehicles (steal from/enter with intent)"), (2, "Other Stealing"),
    (1, "Fraud"),
    (2, "Fraud by Computer"), (2, "Fraud by Cheque"), (2, "Fraud by Credit Card"), (2, "Identity Fraud"), (2, "Other Fraud"),
    (1, "Handling Stolen Goods"),
    (2, "Possess Property Suspected Stolen"), (2, "Receiving Stolen Property"), (2, "Possess etc. Tainted Property"), (2, "Other Handling Stolen Goods"),
    (0, "Other Offences"),
    (1, "Drug Offences"),
    (2, "Trafficking Drugs"), (2, "Possess Drugs"), (2, "Produce Drugs"), (2, "Sell Supply Drugs"), (2, "Other Drug Offences"),
    (1, "Prostitution Offences"),
    (2, "Found in Places Used for Purpose of Prostitution Offences"), (2, "Have Interest in Premises Used for Prostitution Offences"),
    (2, "Knowingly Participate in Provision Prostitution Offences"), (2, "Public Soliciting"), (2, "Procuring Prostitution"),
    (2, "Permit Minor to be at a Place Used for Prostitution Offences"), (2, "Advertising Prostitution"), (2, "Other Prostitution Offences"),
    (1, "Liquor (excl. Drunkenness)"),
    (1, "Gaming Racing & Betting Offences"),
    (1, "Breach Domestic Violence Protection Order"),
    (1, "Trespassing and Vagrancy"),
    (1, "Weapons Act Offences"),
    (2, "Unlawful Possess Concealable Firearm"), (2, "Unlawful Possess Firearm - Other"), (2, "Bomb Possess and/or use of"),
    (2, "Possess and/or use other weapons; restricted items"), (2, "Weapons Act Offences - Other"),
    (1, "Good Order Offences"),
    (2, "Disobey Move-on Direction"), (2, "Resist Incite Hinder Obstruct Police"), (2, "Fare Evasion"), (2, "Public Nuisance"),
    (1, "Stock Related Offences"),
    (1, "Traffic and Related Offences"),
    (2, "Dangerous Operation of a Vehicle"), (2, "Drink Driving"), (2, "Disqualified Driving"),
    (2, "Interfere with Mechanism of Motor Vehicle"), (2, "E-mobility"),
    (1, "Miscellaneous Offences"),
]
MON = ["JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC"]
START_YEAR = 2001          # LGA, division and offender series all begin January 2001
JUV_CHANGE = (2018, 2)     # 17-year-olds counted as juveniles from 12 Feb 2018 (applied from Feb 2018)


def log(*a):
    print(*a, file=sys.stderr, flush=True)


# ----------------------------------------------------------------------------- download
def download(raw, qps_only=False):
    import urllib.request
    os.makedirs(raw, exist_ok=True)
    for name, url in DOWNLOADS.items():
        if qps_only and not (name.endswith(".csv") or name.endswith("_victims_Number.xlsx")):
            continue
        dest = os.path.join(raw, name)
        log("downloading", name)
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (qld-crime-dashboard build script)"})
        with urllib.request.urlopen(req, timeout=300) as r, open(dest, "wb") as fh:
            while True:
                chunk = r.read(1 << 20)
                if not chunk:
                    break
                fh.write(chunk)


# ----------------------------------------------------------------------------- QPS csv
def month_index(s):
    """'JAN01' -> months since Jan START_YEAR (negative before)."""
    m = MON.index(s[:3].upper())
    yy = int(s[3:])
    y = 1900 + yy if yy >= 90 else 2000 + yy
    return (y - START_YEAR) * 12 + m


def read_qps(path):
    """QPS files other than the statewide offence files carry an unnamed trailing row-id column."""
    with open(path, newline="", encoding="utf-8-sig") as fh:
        r = csv.reader(fh)
        header = next(r)
        first = next(r)
    names = header + ["_rowid"] if len(first) == len(header) + 1 else header
    df = pd.read_csv(path, header=None, skiprows=1, names=names, encoding="utf-8-sig")
    df["t"] = df["Month Year"].map(month_index)
    return df[df["t"] >= 0].copy()


def read_victims(path, key=None):
    """QPS victim workbooks: sheet 1 = offences against the person by Age (Adult = 18+, Juvenile = 0-17) and Sex;
    sheet 2 = all victims (offences against the person, unlawful entry, unlawful use of a motor vehicle), no age or sex.
    'Month Year' is a date; trailing unnamed row-id columns are dropped."""
    sheets = pd.read_excel(path, sheet_name=None)
    a, b = list(sheets.values())[:2]
    out = []
    for df in (a, b):
        df = df.loc[:, [c for c in df.columns if not str(c).startswith("Unnamed") and str(c) != "None"]].copy()
        d = pd.to_datetime(df["Month Year"])
        df["t"] = (d.dt.year - START_YEAR) * 12 + d.dt.month - 1
        out.append(df[df["t"] >= 0].copy())
    return out


def implied_pop(num, rate, cols):
    """QPS rates are per 100,000 persons, so population = number / rate * 1e5 wherever rate > 0."""
    N = num[cols].to_numpy(float)
    R = rate[cols].to_numpy(float)
    with np.errstate(divide="ignore", invalid="ignore"):
        ip = np.where(R > 0, N / R * 1e5, np.nan)
        med = np.nanmedian(ip, axis=1)
        spread = (np.nanmax(ip, axis=1) - np.nanmin(ip, axis=1)) / med
    assert np.nanmax(spread) < 1e-6, "implied populations disagree across offence columns"
    return med


def to_matrix(df, key, areas, col, T):
    """rows=areas, cols=months."""
    out = np.zeros((len(areas), T))
    idx = {a: i for i, a in enumerate(areas)}
    ai = df[key].map(idx).to_numpy()
    out[ai, df["t"].to_numpy()] = df[col].to_numpy(float)
    return out


def fill_rows(M):
    """forward- then back-fill NaNs along each row."""
    df = pd.DataFrame(M.T)
    return df.ffill().bfill().to_numpy().T


# ----------------------------------------------------------------------------- ABS population
def qld_single_year(path):
    """ERP by single year of age at 30 June, Queensland, by sex.
    Returns {'Persons'|'Male'|'Female': {year: np.array(ages 0..100)}}."""
    wb_sheets = pd.read_excel(path, sheet_name=None, header=None)
    series = {"Persons": {}, "Male": {}, "Female": {}}
    years = None
    for name, sh in wb_sheets.items():
        if not name.startswith("Data"):
            continue
        hdr = sh.iloc[0]
        datarows = sh.iloc[10:]
        years = [d.year for d in pd.to_datetime(datarows.iloc[:, 0])]
        for c in range(1, sh.shape[1]):
            parts = [p.strip() for p in str(hdr.iloc[c]).split(";")]
            if len(parts) >= 3 and parts[1] in series:
                a = 100 if parts[2].startswith("100") else int(parts[2])
                series[parts[1]][a] = datarows.iloc[:, c].to_numpy(float)
    out = {}
    for sex, d in series.items():
        assert len(d) == 101, (sex, len(d))
        arr = np.array([d[a] for a in range(101)])  # age x year
        out[sex] = {y: arr[:, i] for i, y in enumerate(years)}
    return out


def norm_lga(s):
    s = s.replace("(T)", "").replace("(Qld)", "").replace("-", " ").strip()
    for suf in ["Aboriginal Shire Council", "Regional Council", "City Council", "Shire Council", "Town Council", "Council"]:
        if s.endswith(suf):
            s = s[: -len(suf)].strip()
    return " ".join(s.lower().split())


def lga_age(path):
    """ABS LGA ERP by five-year age group, Queensland, by sex.
    Returns ({'Persons'|'Male'|'Female': DataFrame}, age labels). The sex of each table is read from its title."""
    out = {}
    ages = None
    for sheet, sh in pd.read_excel(path, sheet_name=None, header=None).items():
        if not sheet.lower().startswith("table"):
            continue
        title = " ".join(str(x) for x in sh.iloc[:4, 0]).lower()
        sex = "Female" if "female" in title else "Male" if "male" in title else "Persons" if "person" in title else None
        if sex is None:
            continue
        hdr_row = next(i for i in range(12) if str(sh.iloc[i, 0]).strip() == "Year")
        ages = [str(x).replace("\u2013", "-").strip() for x in sh.iloc[hdr_row - 1, 5:]]
        df = sh.iloc[hdr_row + 1:].copy()
        df.columns = ["year", "ste", "ste_name", "code", "name"] + ages
        df = df[pd.to_numeric(df["ste"], errors="coerce") == 3].copy()
        df["year"] = df["year"].astype(int)
        df["key"] = df["name"].map(norm_lga)
        for a in ages:
            df[a] = pd.to_numeric(df[a], errors="coerce")
        out[sex] = df
    assert set(out) == {"Persons", "Male", "Female"}, set(out)
    return out, ages


def monthly_from_june(years, values, T):
    """Linear interpolation of 30-June estimates to mid-month values; linear extrapolation at both ends."""
    years = np.asarray(years, float) + 0.5
    values = np.asarray(values, float)
    t = START_YEAR + (np.arange(T) + 0.5) / 12.0
    out = np.interp(t, years, values)
    lo = t < years[0]
    out[lo] = values[0] + (values[1] - values[0]) * (t[lo] - years[0])
    hi = t > years[-1]
    out[hi] = values[-1] + (values[-1] - values[-2]) * (t[hi] - years[-1])
    return np.maximum(out, 0)


# ----------------------------------------------------------------------------- geometry (pure python)
def read_shp(buf):
    pos = 100
    out = []
    while pos < len(buf):
        _, clen = struct.unpack(">ii", buf[pos:pos + 8])
        pos += 8
        rec = buf[pos:pos + clen * 2]
        pos += clen * 2
        st = struct.unpack("<i", rec[:4])[0]
        if st == 0:
            out.append([])
            continue
        assert st in (5, 15, 25), st
        nparts, npts = struct.unpack("<ii", rec[36:44])
        parts = struct.unpack("<%di" % nparts, rec[44:44 + 4 * nparts])
        o = 44 + 4 * nparts
        pts = struct.unpack("<%dd" % (2 * npts), rec[o:o + 16 * npts])
        rings = []
        for i, s in enumerate(parts):
            e = parts[i + 1] if i + 1 < nparts else npts
            rings.append([(pts[2 * j], pts[2 * j + 1]) for j in range(s, e)])
        out.append(rings)
    return out


def read_dbf(buf, enc="latin1"):
    n, hl, rl = struct.unpack("<IHH", buf[4:12])
    fields = []
    p = 32
    while buf[p] != 0x0D:
        fields.append((buf[p:p + 11].split(b"\0")[0].decode(), buf[p + 16]))
        p += 32
    out = []
    for i in range(n):
        o = hl + i * rl + 1
        r = {}
        for name, l in fields:
            r[name] = buf[o:o + l].decode(enc, "replace").strip()
            o += l
        out.append(r)
    return out


def load_layer(raw, stem):
    """Read <stem>.shp/.dbf from RAW_DIR, a sub-folder, or <stem>.zip."""
    for base in [raw] + [d for d in glob.glob(os.path.join(raw, "*")) if os.path.isdir(d)]:
        if os.path.exists(os.path.join(base, stem + ".shp")):
            return (read_dbf(open(os.path.join(base, stem + ".dbf"), "rb").read()),
                    read_shp(open(os.path.join(base, stem + ".shp"), "rb").read()))
    z = zipfile.ZipFile(os.path.join(raw, stem + ".zip"))
    names = {os.path.basename(n).lower(): n for n in z.namelist()}
    return read_dbf(z.read(names[(stem + ".dbf").lower()])), read_shp(z.read(names[(stem + ".shp").lower()]))


def ring_area_km2(r):
    a = 0.0
    for i in range(len(r) - 1):
        a += r[i][0] * r[i + 1][1] - r[i + 1][0] * r[i][1]
    lat = sum(p[1] for p in r) / len(r)
    return abs(a) / 2 * math.cos(math.radians(lat)) * 111.32 * 110.57


def dp(pts, tol):
    n = len(pts)
    if n < 3:
        return list(pts)
    keep = [False] * n
    keep[0] = keep[-1] = True
    stack = [(0, n - 1)]
    t2 = tol * tol
    while stack:
        a, b = stack.pop()
        ax, ay = pts[a]
        bx, by = pts[b]
        dx = bx - ax
        dy = by - ay
        L = dx * dx + dy * dy
        md = -1.0
        mi = -1
        for i in range(a + 1, b):
            px, py = pts[i]
            if L == 0:
                dd = (px - ax) ** 2 + (py - ay) ** 2
            else:
                t = ((px - ax) * dx + (py - ay) * dy) / L
                t = 0 if t < 0 else 1 if t > 1 else t
                dd = (px - ax - t * dx) ** 2 + (py - ay - t * dy) ** 2
            if dd > md:
                md = dd
                mi = i
        if md > t2:
            keep[mi] = True
            stack.append((a, mi))
            stack.append((mi, b))
    return [p for p, k in zip(pts, keep) if k]


def arc_tol(arc, tmin, tmax):
    xs = [p[0] for p in arc]
    ys = [p[1] for p in arc]
    ext = math.hypot(max(xs) - min(xs), max(ys) - min(ys))
    return min(tmax, max(tmin, ext * 0.004))


def simp_arc(arc, tmin, tmax):
    """Direction-independent simplification so both sides of a shared border come out identical."""
    rev = arc[0] > arc[-1]
    a = arc[::-1] if rev else arc
    s = dp(a, arc_tol(a, tmin, tmax))
    return s[::-1] if rev else s


def simplify_layer(features, tmin=0.0002, tmax=0.004, min_island_km2=1.0):
    """features: list of ring lists. Topology-aware: shared borders are split at junctions and simplified once."""
    key = lambda p: (round(p[0], 6), round(p[1], 6))
    # drop tiny islands (keep the largest ring of every feature)
    feats = []
    for rings in features:
        if not rings:
            feats.append([])
            continue
        areas = [ring_area_km2(r) for r in rings]
        big = max(areas)
        feats.append([r for r, a in zip(rings, areas) if a == big or a >= max(min_island_km2, 0.004 * big)])
    owners = {}
    for fi, rings in enumerate(feats):
        for r in rings:
            for p in r:
                owners.setdefault(key(p), set()).add(fi)
    owners = {k: frozenset(v) for k, v in owners.items()}
    out = []
    for fi, rings in enumerate(feats):
        new = []
        for r in rings:
            if r[0] == r[-1]:
                r = r[:-1]
            n = len(r)
            if n < 3:
                continue
            ow = [owners[key(p)] for p in r]
            brk = [i for i in range(n) if ow[i] != ow[i - 1] or ow[i] != ow[(i + 1) % n]]
            if not brk:
                closed = r + [r[0]]
                x0, y0 = closed[0]
                far = max(range(len(closed)), key=lambda i: (closed[i][0] - x0) ** 2 + (closed[i][1] - y0) ** 2)
                a = simp_arc(closed[:far + 1], tmin, tmax)
                b = simp_arc(closed[far:], tmin, tmax)
                ring = a + b[1:]
            else:
                ring = []
                for j, s in enumerate(brk):
                    e = brk[(j + 1) % len(brk)]
                    arc = r[s:e + 1] if e > s else r[s:] + r[:e + 1]
                    if len(brk) == 1:
                        arc = r[s:] + r[:s + 1]
                        x0, y0 = arc[0]
                        far = max(range(len(arc)), key=lambda i: (arc[i][0] - x0) ** 2 + (arc[i][1] - y0) ** 2)
                        sa = simp_arc(arc[:far + 1], tmin, tmax) + simp_arc(arc[far:], tmin, tmax)[1:]
                    else:
                        sa = simp_arc(arc, tmin, tmax)
                    ring.extend(sa[:-1])
                ring.append(ring[0])
            # quantise to 1e-4 degrees and drop repeats
            q = []
            for x, y in ring:
                pt = (int(round(x * 1e4)), int(round(y * 1e4)))
                if not q or q[-1] != pt:
                    q.append(pt)
            if len(q) >= 4:
                new.append(q)
        out.append(new)
    return out


def encode_rings(rings):
    """Each ring -> flat delta-encoded integer list [x0, y0, dx1, dy1, ...] in 1e-4 degrees."""
    enc = []
    for q in rings:
        flat = [q[0][0], q[0][1]]
        for i in range(1, len(q)):
            flat.append(q[i][0] - q[i - 1][0])
            flat.append(q[i][1] - q[i - 1][1])
        enc.append(flat)
    return enc


# ----------------------------------------------------------------------------- output helpers
def sparse(v):
    """ints joined by commas, zeros left empty."""
    return ",".join("" if x == 0 else str(int(x)) for x in v)


def rle(v):
    out = []
    prev = None
    n = 0
    for x in v:
        if x == prev:
            n += 1
        else:
            if prev is not None:
                out += [prev, n]
            prev = x
            n = 1
    out += [prev, n]
    return out


def popround(x):
    return round(float(x), 1) if x < 2000 else int(round(float(x)))


def dump(obj, path):
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(obj, fh, separators=(",", ":"), ensure_ascii=False)
    return os.path.getsize(path)


# ----------------------------------------------------------------------------- tidy analysis files
def month_str(t):
    return "%04d-%02d" % (START_YEAR + t // 12, t % 12 + 1)


def write_tidy(folder, c):
    """Long-format csv.gz files: one row per location x offence x month (x age group x sex for offenders)."""
    os.makedirs(folder, exist_ok=True)
    T, offs = c["T"], c["offences"]
    names = [o["name"] for o in offs]
    parent = [names[o["parent"]] if o["parent"] is not None else "" for o in offs]
    months = [month_str(t) for t in range(T)]
    NO = len(offs)

    def frame(geo_labels, loc_labels, counts, pops, extra=None, idx=None):
        """counts: (NO, L, T); pops: (L, T). Row order: location, offence, month. idx: subset of offence ids."""
        offs_ = offs if idx is None else [offs[k] for k in idx]
        counts = counts if idx is None else counts[idx]
        names_ = [o["name"] for o in offs_]
        parent_ = [names[o["parent"]] if o["parent"] is not None else "" for o in offs_]
        NO = len(offs_)
        L = len(loc_labels)
        cnt = np.transpose(counts, (1, 0, 2)).reshape(-1)
        pp = np.repeat(pops[:, None, :], NO, axis=1).reshape(-1)
        with np.errstate(divide="ignore", invalid="ignore"):
            rate = np.where(pp > 0, cnt / pp * 1e5, np.nan)
        loc = np.repeat(np.arange(L), NO * T)
        off = np.tile(np.repeat(np.arange(NO), T), L)
        df = pd.DataFrame({
            "geography": pd.Categorical.from_codes(pd.factorize(np.array(geo_labels, dtype=object))[0][loc], list(pd.unique(pd.Series(geo_labels)))),
            "location": pd.Categorical.from_codes(loc, loc_labels),
            "month": pd.Categorical.from_codes(np.tile(np.arange(T), L * NO), months),
            "offence": pd.Categorical.from_codes(off, names_),
            "offence_level": np.array([o["lvl"] for o in offs_], np.int8)[off],
            "parent_offence": pd.Categorical.from_codes(pd.factorize(np.array(parent_, dtype=object))[0][off], list(pd.unique(pd.Series(parent_)))),
        })
        if extra:
            for k, v in extra.items():
                df[k] = v
        return df, cnt, pp, rate

    out = {}
    # offences: Queensland + LGAs
    loc_labels = ["Queensland"] + c["lgas"]
    counts = np.concatenate([c["C_q"][:, None, :], c["C_l"]], axis=1)
    pops = np.vstack([c["pop_q"][None, :], c["pop_l"]])
    geo = ["Queensland"] + ["LGA"] * len(c["lgas"])
    df, cnt, pp, rate = frame(geo, loc_labels, counts, pops)
    df["offences"] = cnt; df["population"] = np.round(pp, 1); df["rate_per_100k"] = np.round(rate, 3)
    out["offences_qld_lga.csv.gz"] = df
    # offences: police divisions
    df, cnt, pp, rate = frame(["Police division"] * len(c["divs"]), c["divs"], c["C_d"], c["pop_d"])
    df["offences"] = cnt; df["population"] = np.round(pp, 1); df["rate_per_100k"] = np.round(rate, 3)
    out["offences_division.csv.gz"] = df
    # offenders: Queensland + LGAs, by age group and sex
    parts = []
    for a, alabel in (("a", "Adult"), ("j", "Juvenile")):
        for sx, slabel in (("", "Persons"), ("m", "Male"), ("f", "Female")):
            counts = np.concatenate([c["O_q"][a + sx][:, None, :], c["O_l"][a + sx]], axis=1)
            pops = np.vstack([c["den"]["q" + a + sx][None, :], c["den"]["l" + a + sx]])
            df, cnt, pp, rate = frame(geo, loc_labels, counts, pops)
            df["age_group"] = alabel; df["sex"] = slabel
            df["offenders"] = cnt; df["population"] = np.round(pp, 1); df["rate_per_100k"] = np.round(rate, 3)
            parts.append(df)
    out["offenders_qld_lga.csv.gz"] = pd.concat(parts, ignore_index=True)
    # victims: Queensland + LGAs; all victims (all ages, persons) for every offence with victim data, and by age group and sex
    # for offences against the person. Juvenile victims are aged 0-17 (QPS definition).
    v_all = [o["id"] for o in offs if o.get("vic")]; v_as = [o["id"] for o in offs if o.get("vic") == 2]
    parts = []
    counts = np.concatenate([c["V_q"][""][:, None, :], c["V_l"][""]], axis=1)
    pops = np.vstack([c["pop_q"][None, :], c["pop_l"]])
    df, cnt, pp, rate = frame(geo, loc_labels, counts, pops, idx=v_all)
    df["age_group"] = "All ages"; df["sex"] = "Persons"; df["victims"] = cnt; df["population"] = np.round(pp, 1); df["rate_per_100k"] = np.round(rate, 3)
    parts.append(df)
    for a, alabel in (("a", "Adult (18+)"), ("j", "Juvenile (0-17)")):
        for sx, slabel in (("", "Persons"), ("m", "Male"), ("f", "Female")):
            counts = np.concatenate([c["V_q"][a + sx][:, None, :], c["V_l"][a + sx]], axis=1)
            pops = np.vstack([c["den"]["qv" + a + sx][None, :], c["den"]["lv" + a + sx]])
            df, cnt, pp, rate = frame(geo, loc_labels, counts, pops, idx=v_as)
            df["age_group"] = alabel; df["sex"] = slabel
            df["victims"] = cnt; df["population"] = np.round(pp, 1); df["rate_per_100k"] = np.round(rate, 3)
            parts.append(df)
    out["victims_qld_lga.csv.gz"] = pd.concat(parts, ignore_index=True)
    for name, df in out.items():
        path = os.path.join(folder, name)
        df.to_csv(path, index=False, compression={"method": "gzip", "compresslevel": 6})
        log("tidy:", name, len(df), "rows,", os.path.getsize(path), "bytes")


# ----------------------------------------------------------------------------- refresh report
def write_report(path, c, min_mean=20, hist=60):
    """Markdown summary of the latest month. Each series is compared with a seasonal-naive forecast with drift:
    forecast(t) = y(t-12) * mean(y[t-12..t-1]) / mean(y[t-24..t-13]); the 95% interval comes from the empirical
    2.5th and 97.5th percentiles of the previous `hist` one-step log errors of the same forecast."""
    T, offs = c["T"], c["offences"]
    main = [o for o in offs if o["lvl"] <= 1]
    t = T - 1

    def check(cnt, pop):
        y = cnt / pop * 1e5
        if t < 24 + hist or cnt[t - 24:t].mean() < min_mean:
            return None
        def fc(u):
            g = y[u - 12:u].mean() / y[u - 24:u - 12].mean() if y[u - 24:u - 12].mean() > 0 else 1.0
            return y[u - 12] * g
        errs = []
        for u in range(t - hist, t):
            f = fc(u)
            if f > 0 and y[u] > 0:
                errs.append(math.log(y[u] / f))
        if len(errs) < hist * 0.8:
            return None
        f = fc(t)
        lo, hi = f * math.exp(np.percentile(errs, 2.5)), f * math.exp(np.percentile(errs, 97.5))
        return y[t], f, lo, hi

    def annual(cnt, pop, e):
        return cnt[e - 11:e + 1].sum() / pop[e - 11:e + 1].mean() * 1e5

    L = []
    L.append(f"# Queensland crime data refresh: data to {MON[t % 12].title()} {START_YEAR + t // 12}\n")
    L.append(f"Built {datetime.date.today().isoformat()}. Series run from Jan {START_YEAR}; {T} months.\n")
    L.append("## Queensland, main categories\n")
    L.append("| Offence | Rate, latest 12 months | Change on previous 12 months | Latest month rate | Forecast (95% interval) | Flag |")
    L.append("|---|---:|---:|---:|---|---|")
    for o in main:
        cnt, pop = c["C_q"][o["id"]].astype(float), c["pop_q"]
        r1, r0 = annual(cnt, pop, t), annual(cnt, pop, t - 12)
        ch = (r1 / r0 - 1) * 100 if r0 > 0 else float("nan")
        chs = f"{ch:+.1f}%" if ch == ch else "n/a"
        k = check(cnt, pop)
        if k:
            y, f, lo, hi = k
            flag = "above" if y > hi else "below" if y < lo else ""
            L.append(f"| {o['name']} | {r1:,.1f} | {chs} | {y:,.2f} | {f:,.2f} ({lo:,.2f} to {hi:,.2f}) | {flag} |")
        else:
            L.append(f"| {o['name']} | {r1:,.1f} | {chs} | {cnt[t] / pop[t] * 1e5:,.2f} | too few to forecast | |")
    flagged = []
    for o in main:
        for i, name in enumerate(c["lgas"]):
            cnt, pop = c["C_l"][o["id"], i].astype(float), c["pop_l"][i]
            k = check(cnt, pop)
            if k and (k[0] > k[3] or k[0] < k[2]):
                y, f, lo, hi = k
                flagged.append((abs(math.log(max(y, 1e-9) / f)), name, o["name"], y, f, lo, hi, int(cnt[t])))
    flagged.sort(reverse=True)
    L.append(f"\n## LGA series outside their forecast interval in the latest month ({len(flagged)} of the series with at least {min_mean} offences a month)\n")
    if flagged:
        L.append("| LGA | Offence | Latest month rate | Forecast (95% interval) | Offences |")
        L.append("|---|---|---:|---|---:|")
        for _, name, off, y, f, lo, hi, n in flagged[:25]:
            L.append(f"| {name} | {off} | {y:,.1f} | {f:,.1f} ({lo:,.1f} to {hi:,.1f}) | {n:,} |")
    L.append("\nForecast: same month a year earlier, scaled by the ratio of the latest 12 months to the 12 months before. "
             f"Interval: 2.5th to 97.5th percentile of that forecast's errors over the previous {hist} months. "
             "With a 95% interval, about 1 series in 20 is expected to fall outside by chance.\n")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write("\n".join(L))
    log("report written:", path)


# ----------------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--raw", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--download", action="store_true", help="fetch all QPS and ABS source files into --raw first")
    ap.add_argument("--download-qps", action="store_true", help="fetch only the eight monthly QPS csv files into --raw first")
    ap.add_argument("--tidy", help="folder for long-format analysis files (csv.gz)")
    ap.add_argument("--report", help="path for a markdown report on the latest month")
    args = ap.parse_args()
    raw, out = args.raw, args.out
    if args.download or args.download_qps:
        download(raw, qps_only=args.download_qps and not args.download)
    os.makedirs(os.path.join(out, "o"), exist_ok=True)
    P = lambda n: os.path.join(raw, n)

    # ---- QPS offences
    qn, qr = read_qps(P("QLD_Reported_Offences_Number.csv")), read_qps(P("QLD_Reported_Offences_Rates.csv"))
    ln, lr = read_qps(P("LGA_Reported_Offences_Number.csv")), read_qps(P("LGA_Reported_Offences_Rates.csv"))
    dn, dr = read_qps(P("division_Reported_Offences_Number.csv")), read_qps(P("division_Reported_Offences_Rates.csv"))
    qo, lo = read_qps(P("QLD_Reported_Offenders_Number.csv")), read_qps(P("LGA_Reported_Offenders_Number.csv"))
    qv_as, qv_all = read_victims(P("QLD_Reported_victims_Number.xlsx"))
    lv_as, lv_all = read_victims(P("LGA_Reported_victims_Number.xlsx"))
    cols = [c for _, c in TREE]
    V_AS = [c for c in cols if c in qv_as.columns]      # offences with victims by age and sex (offences against the person)
    V_ALL = [c for c in cols if c in qv_all.columns]    # offences with victim totals (adds unlawful entry, unlawful use of a motor vehicle)
    for df, nm, want in [(qv_as, "QLD victims", V_AS), (lv_as, "LGA victims", V_AS), (qv_all, "QLD victims (all)", V_ALL), (lv_all, "LGA victims (all)", V_ALL)]:
        extra = set(df.columns) - set(cols) - {"Month Year", "LGA Name", "Age", "Sex", "t"}
        assert not extra and set(want) <= set(df.columns), (nm, extra)
    assert len(V_AS) >= 20 and "Unlawful Entry" in V_ALL, (V_AS, V_ALL)
    log("victim offence columns: by age and sex", len(V_AS), "| totals", len(V_ALL))
    for df, nm in [(qn, "QLD offences"), (ln, "LGA offences"), (dn, "division offences"), (qo, "QLD offenders"), (lo, "LGA offenders")]:
        have = set(df.columns) - {"Month Year", "LGA Name", "Division", "Age", "Sex", "_rowid", "t"}
        assert have == set(cols), (nm, have ^ set(cols))
    T = int(min(qn["t"].max(), ln["t"].max(), dn["t"].max(), qo["t"].max(), lo["t"].max(),
                qv_as["t"].max(), lv_as["t"].max(), qv_all["t"].max(), lv_all["t"].max())) + 1
    for df in (qn, qr, ln, lr, dn, dr, qo, lo, qv_as, lv_as, qv_all, lv_all):
        df.drop(df[df["t"] >= T].index, inplace=True)
    lgas = sorted(ln["LGA Name"].unique())
    divs = sorted(dn["Division"].unique())
    assert sorted(lo["LGA Name"].unique()) == lgas
    assert sorted(lv_as["LGA Name"].unique()) == lgas and sorted(lv_all["LGA Name"].unique()) == lgas, "victim LGA names differ"
    log(f"months={T} (Jan {START_YEAR} to {MON[(T-1)%12]} {START_YEAR+(T-1)//12}), LGAs={len(lgas)}, divisions={len(divs)}")

    # ---- populations implied by QPS rates (offence-rate denominators)
    qn = qn.sort_values("t"); qr = qr.sort_values("t")
    assert (ln["LGA Name"].values == lr["LGA Name"].values).all() and (ln["t"].values == lr["t"].values).all()
    assert (dn["Division"].values == dr["Division"].values).all() and (dn["t"].values == dr["t"].values).all()
    qr = qr[qn.columns]
    pop_q = pd.Series(implied_pop(qn, qr, cols)).ffill().bfill().to_numpy()
    ln["_pop"] = implied_pop(ln, lr, cols)
    dn["_pop"] = implied_pop(dn, dr, cols)
    pop_l_raw = to_matrix(ln.assign(_p=ln["_pop"].fillna(-1)), "LGA Name", lgas, "_p", T)
    pop_l_raw[pop_l_raw < 0] = np.nan
    pop_d_raw = to_matrix(dn.assign(_p=dn["_pop"].fillna(-1)), "Division", divs, "_p", T)
    pop_d_raw[pop_d_raw < 0] = np.nan
    log("months with no implied population (filled from neighbours): LGA", int(np.isnan(pop_l_raw).sum()), "division", int(np.isnan(pop_d_raw).sum()))
    pop_l, pop_d = fill_rows(pop_l_raw), fill_rows(pop_d_raw)

    # ---- ABS age-specific denominators (offender rates): persons, males, females
    sy_all = qld_single_year(P("3101053.xlsx"))
    lga_file = sorted(glob.glob(P("32350DS0006*.xlsx")))[-1]
    la_all, ages = lga_age(lga_file)
    t_new = (JUV_CHANGE[0] - START_YEAR) * 12 + (JUV_CHANGE[1] - 1)
    newdef = np.arange(T) >= t_new

    def pick(old, new):
        return np.where(newdef, new, old)

    older = [a for a in ages if a not in ("0-4", "5-9", "10-14", "15-19") and "total" not in a.lower() and "all" not in a.lower()]
    assert "10-14" in ages and "15-19" in ages and len(older) >= 14, ages
    keymap = {norm_lga(n): n for n in lgas}
    SEX = {"": "Persons", "m": "Male", "f": "Female"}   # suffix -> ABS label
    den = {}                                            # e.g. den['qj'], den['qam'], den['ljf'] ...
    for sx, label in SEX.items():
        sy, la = sy_all[label], la_all[label]
        yrs = sorted(y for y in sy if y >= START_YEAR)
        lyrs = sorted(la["year"].unique())
        missing = set(keymap) - set(la["key"])
        assert not missing, ("LGAs with no ABS population", missing)
        # Queensland: exact single years of age
        j_old = [sy[y][10:17].sum() for y in yrs]; j_new = [sy[y][10:18].sum() for y in yrs]
        a_old = [sy[y][17:].sum() for y in yrs];   a_new = [sy[y][18:].sum() for y in yrs]
        den["qj" + sx] = pick(monthly_from_june(yrs, j_old, T), monthly_from_june(yrs, j_new, T))
        den["qvj" + sx] = monthly_from_june(yrs, [sy[y][0:18].sum() for y in yrs], T)   # victims: juvenile = 0-17 throughout
        den["qva" + sx] = monthly_from_june(yrs, [sy[y][18:].sum() for y in yrs], T)
        den["qa" + sx] = pick(monthly_from_june(yrs, a_old, T), monthly_from_june(yrs, a_new, T))
        # share of the 15-19 group that is juvenile under each definition (Queensland-wide, same sex, by year)
        sh_old = {y: sy[y][15:17].sum() / sy[y][15:20].sum() for y in yrs}
        sh_new = {y: sy[y][15:18].sum() / sy[y][15:20].sum() for y in yrs}
        lj = np.zeros((len(lgas), T)); laa = np.zeros((len(lgas), T)); lvj = np.zeros((len(lgas), T)); lva = np.zeros((len(lgas), T))
        for i, name in enumerate(lgas):
            d = la[la["key"] == norm_lga(name)].set_index("year").sort_index()
            ys = list(d.index)
            p1014 = d["10-14"].to_numpy(float); p1519 = d["15-19"].to_numpy(float); p20 = d[older].sum(axis=1).to_numpy(float)
            so = np.array([sh_old[y] for y in ys]); sn = np.array([sh_new[y] for y in ys])
            lj[i] = pick(monthly_from_june(ys, p1014 + p1519 * so, T), monthly_from_june(ys, p1014 + p1519 * sn, T))
            laa[i] = pick(monthly_from_june(ys, p20 + p1519 * (1 - so), T), monthly_from_june(ys, p20 + p1519 * (1 - sn), T))
            p09 = d["0-4"].to_numpy(float) + d["5-9"].to_numpy(float)
            lvj[i] = monthly_from_june(ys, p09 + p1014 + p1519 * sn, T)
            lva[i] = monthly_from_june(ys, p20 + p1519 * (1 - sn), T)
        den["lj" + sx], den["la" + sx] = lj, laa
        den["lvj" + sx], den["lva" + sx] = lvj, lva
    log("ABS Qld single-year ERP years", yrs[0], "-", yrs[-1], "| ABS LGA ERP years", lyrs[0], "-", lyrs[-1])
    chk = (den["lj"].sum(0) + den["la"].sum(0)) / (den["qj"] + den["qa"])
    log("sum of LGA 10+ denominators / Qld 10+ denominator: min %.4f max %.4f" % (chk.min(), chk.max()))
    chk2 = (den["qjm"] + den["qjf"]) / den["qj"]
    log("Qld male + female juvenile denominator / persons: min %.5f max %.5f" % (chk2.min(), chk2.max()))
    chk3 = (den["lvj"].sum(0) + den["lva"].sum(0)) / (den["qvj"] + den["qva"])
    log("sum of LGA all-age victim denominators / Qld: min %.4f max %.4f" % (chk3.min(), chk3.max()))
    den_qj, den_qa, den_lj, den_la = den["qj"], den["qa"], den["lj"], den["la"]

    # ---- offender matrices: persons (all sexes, including sex not stated), males, females
    OTAGS = [a + sx for a in ("a", "j") for sx in ("", "m", "f")]

    def offender_mats(df, key, areas, col):
        res = {}
        for age, a in (("Adult", "a"), ("Juvenile", "j")):
            for sx, label in SEX.items():
                sub = df[(df["Age"] == age) & ((df["Sex"] == label) if sx else True)]
                if key:
                    g = sub.groupby([key, "t"], sort=False)[col].sum().reset_index()
                    res[a + sx] = to_matrix(g, key, areas, col, T)
                else:
                    g = sub.groupby("t", sort=False)[col].sum()
                    v = np.zeros(T); v[g.index.to_numpy()] = g.to_numpy(float); res[a + sx] = v
        return res

    def victim_total(df, key, areas, col):
        if key:
            return to_matrix(df.groupby([key, "t"], sort=False)[col].sum().reset_index(), key, areas, col, T)
        g = df.groupby("t", sort=False)[col].sum(); v = np.zeros(T); v[g.index.to_numpy()] = g.to_numpy(float); return v

    # ---- summary windows (12 months ending at the latest month, stepping back a year at a time)
    nwin = (T - 12) // 12 + 1
    win_end = [T - 1 - 12 * k for k in range(nwin)]

    def wsum(M):  # M: (..., T) -> (..., nwin)
        return np.stack([M[..., e - 11:e + 1].sum(-1) for e in win_end], -1)

    def wmean(M):
        return np.stack([M[..., e - 11:e + 1].mean(-1) for e in win_end], -1)

    summary = {"c_q": [], "c_l": [], "c_d": [], "a_q": [], "j_q": [], "a_l": [], "j_l": [],
               "v_q": [], "v_l": [], "va_q": [], "vj_q": [], "va_l": [], "vj_l": []}
    sizes = []
    offences = []
    NO = len(TREE)
    C_q = np.zeros((NO, T), np.int32); C_l = np.zeros((NO, len(lgas), T), np.int32); C_d = np.zeros((NO, len(divs), T), np.int32)
    O_q = {k: np.zeros((NO, T), np.int32) for k in OTAGS}; O_l = {k: np.zeros((NO, len(lgas), T), np.int32) for k in OTAGS}
    VTAGS = [""] + OTAGS                                 # '' = all victims; a, am, af, j, jm, jf = by age group and sex
    V_q = {k: np.zeros((NO, T), np.int32) for k in VTAGS}; V_l = {k: np.zeros((NO, len(lgas), T), np.int32) for k in VTAGS}
    vic_level = []
    for oi, (lvl, col) in enumerate(TREE):
        mq = np.zeros(T); mq[qn["t"].to_numpy()] = qn[col].to_numpy(float)
        ml = to_matrix(ln, "LGA Name", lgas, col, T)
        md = to_matrix(dn, "Division", divs, col, T)
        oq = offender_mats(qo, None, None, col)
        ol = offender_mats(lo, "LGA Name", lgas, col)
        C_q[oi], C_l[oi], C_d[oi] = mq, ml, md
        for k in OTAGS:
            O_q[k][oi], O_l[k][oi] = oq[k], ol[k]
        f = {"q": sparse(mq), "l": [sparse(r) for r in ml], "d": [sparse(r) for r in md]}
        for k in OTAGS:
            f["q" + k] = sparse(oq[k])
            f["l" + k] = [sparse(r) for r in ol[k]]
        vl = 2 if col in V_AS else 1 if col in V_ALL else 0
        vic_level.append(vl)
        if vl:
            vq, vlm = victim_total(qv_all, None, None, col), victim_total(lv_all, "LGA Name", lgas, col)
            V_q[""][oi], V_l[""][oi] = vq, vlm
            f["qv"] = sparse(vq); f["lv"] = [sparse(r) for r in vlm]
            summary["v_q"].append(wsum(vq).astype(int).tolist()); summary["v_l"].append(wsum(vlm).astype(int).tolist())
        else:
            for k in ("v_q", "v_l"):
                summary[k].append(None)
        if vl == 2:
            vqa, vla = offender_mats(qv_as, None, None, col), offender_mats(lv_as, "LGA Name", lgas, col)
            for k in OTAGS:
                V_q[k][oi], V_l[k][oi] = vqa[k], vla[k]
                f["qv" + k] = sparse(vqa[k]); f["lv" + k] = [sparse(r) for r in vla[k]]
            summary["va_q"].append(wsum(vqa["a"]).astype(int).tolist()); summary["vj_q"].append(wsum(vqa["j"]).astype(int).tolist())
            summary["va_l"].append(wsum(vla["a"]).astype(int).tolist()); summary["vj_l"].append(wsum(vla["j"]).astype(int).tolist())
        else:
            for k in ("va_q", "vj_q", "va_l", "vj_l"):
                summary[k].append(None)
        sizes.append(dump(f, os.path.join(out, "o", f"{oi}.json")))
        summary["c_q"].append(wsum(mq).astype(int).tolist())
        summary["c_l"].append(wsum(ml).astype(int).tolist())
        summary["c_d"].append(wsum(md).astype(int).tolist())
        summary["a_q"].append(wsum(oq["a"]).astype(int).tolist()); summary["j_q"].append(wsum(oq["j"]).astype(int).tolist())
        summary["a_l"].append(wsum(ol["a"]).astype(int).tolist()); summary["j_l"].append(wsum(ol["j"]).astype(int).tolist())
        offences.append({"id": oi, "name": col.rstrip("'"), "lvl": lvl, "total": int(mq.sum()), "vic": vl})
    for i, o in enumerate(offences):  # parent pointers
        o["parent"] = next((j for j in range(i - 1, -1, -1) if offences[j]["lvl"] < o["lvl"]), None)
    summary["win_end"] = win_end
    summary["p_q"] = [popround(x) for x in wmean(pop_q)]
    summary["p_l"] = [[popround(x) for x in r] for r in wmean(pop_l)]
    summary["p_d"] = [[popround(x) for x in r] for r in wmean(pop_d)]
    summary["pj_q"] = [popround(x) for x in wmean(den_qj)]; summary["pa_q"] = [popround(x) for x in wmean(den_qa)]
    summary["pj_l"] = [[popround(x) for x in r] for r in wmean(den_lj)]
    summary["pa_l"] = [[popround(x) for x in r] for r in wmean(den_la)]
    summary["pvj_q"] = [popround(x) for x in wmean(den["qvj"])]; summary["pva_q"] = [popround(x) for x in wmean(den["qva"])]
    summary["pvj_l"] = [[popround(x) for x in r] for r in wmean(den["lvj"])]
    summary["pva_l"] = [[popround(x) for x in r] for r in wmean(den["lva"])]
    oap = cols.index("Offences Against the Person")
    log("victims, offences against the person, all months: Qld %d | sum of LGAs / Qld %.4f | adult + juvenile / all %.4f" % (
        V_q[""][oap].sum(), V_l[""][oap].sum() / max(1, V_q[""][oap].sum()), (V_q["a"][oap].sum() + V_q["j"][oap].sum()) / max(1, V_q[""][oap].sum())))
    s_sum = dump(summary, os.path.join(out, "summary.json"))

    pop = {"q": rle([popround(x) for x in pop_q]),
           "l": [rle([popround(x) for x in r]) for r in pop_l],
           "d": [rle([popround(x) for x in r]) for r in pop_d]}
    for k, v in den.items():   # qj, qa, qjm, ... (vectors) and lj, la, ljm, ... (LGA x month)
        pop[k] = [popround(x) for x in v] if v.ndim == 1 else [[popround(x) for x in r] for r in v]
    s_pop = dump(pop, os.path.join(out, "pop.json"))

    # ---- geometry
    drec, dshp = load_layer(raw, "QPS_DIVISIONS")
    by = {r["DIVISION"].upper(): g for r, g in zip(drec, dshp)}
    assert set(by) == {d.upper() for d in divs}, set(by) ^ {d.upper() for d in divs}
    dgeo = simplify_layer([by[d.upper()] for d in divs], tmin=0.00015, tmax=0.004, min_island_km2=0.5)
    lstem = os.path.splitext(os.path.basename(sorted(glob.glob(P("LGA_20*_AUST_GDA2020*")))[-1]))[0]
    lrec, lshp = load_layer(raw, lstem)
    namecol = next(k for k in lrec[0] if k.startswith("LGA_NAME"))
    lby = {norm_lga(r[namecol]): g for r, g in zip(lrec, lshp) if r.get("STE_CODE21") == "3"}
    assert not (set(keymap) - set(lby)), set(keymap) - set(lby)
    lgeo = simplify_layer([lby[norm_lga(n)] for n in lgas], tmin=0.0004, tmax=0.006, min_island_km2=2.0)
    geo = {"l": [encode_rings(r) for r in lgeo], "d": [encode_rings(r) for r in dgeo]}
    s_geo = dump(geo, os.path.join(out, "geo.json"))
    log("geometry points: division", sum(len(q) for f in dgeo for q in f), "LGA", sum(len(q) for f in lgeo for q in f))

    meta = {
        "built": datetime.date.today().isoformat(),
        "start_year": START_YEAR, "months": T,
        "juv_change_t": int(t_new),
        "abs_last_june": int(yrs[-1]), "abs_lga_last_june": int(lyrs[-1]),
        "offences": offences, "lgas": lgas, "divisions": divs,
        "win_end": win_end,
    }
    s_meta = dump(meta, os.path.join(out, "meta.json"))
    log("sizes (bytes): meta", s_meta, "pop", s_pop, "summary", s_sum, "geo", s_geo,
        "offence files total", sum(sizes), "max", max(sizes))

    ctx = dict(T=T, offences=offences, lgas=lgas, divs=divs, C_q=C_q, C_l=C_l, C_d=C_d, O_q=O_q, O_l=O_l, V_q=V_q, V_l=V_l,
               pop_q=pop_q, pop_l=pop_l, pop_d=pop_d, den=den)
    if args.report:
        write_report(args.report, ctx)
    if args.tidy:
        write_tidy(args.tidy, ctx)


if __name__ == "__main__":
    main()
