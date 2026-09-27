#!/usr/bin/env python3
"""
ro_sites.py - Rebuild the radiation oncology practice-site dataset from
Yu, Lu, Arons, Rowley, Sindhu (IJROBP 2026, "Structural Vulnerability in the US
RO Delivery System"), then extend it into a freestanding-practice target list.

What the paper did (and what this script reproduces):
  1. Source: CMS Doctors & Clinicians National Downloadable File (NDF),
     final release of each year 2018-2025.
  2. Keep radiation oncologists; collapse rows to unique physical sites.
     Paper: Google Address Validation API -> Google placeId per address.
     Here: rule-based address normalization by default, optional Google API.
  3. Hospital affiliation = any hospital CCN (hosp_afl_* columns) on the NDF
     row. CMS dropped these columns after 2021, so "freestanding" is only
     knowable from pre-2022 files. Assigned at the site's first observed year.
  4. Rurality = 2023 USDA RUCC via HUD ZIP->county crosswalk, dominant county
     (largest TOT_RATIO). Urban 1-3, rural-adjacent 4/6/8, nonadjacent 5/7/9.
  5. Model: discrete-time logit, P(site present at t is gone at t+2),
     predictors freestanding, rural, log(org size); state + year FE;
     SEs clustered by site. Sample: sites first seen 2018-2021.

Extension for current targeting (not in the paper):
  Because affiliation isn't in post-2021 NDFs, we flag CURRENT office-based
  (i.e., freestanding) practice from the Medicare Physician & Other
  Practitioners "by Provider and Service" file: radiation oncologists billing
  treatment management (77427) or delivery codes with Place_Of_Srvc == "O"
  (non-facility) are practicing in a freestanding setting.

Outputs (in --out):
  ro_clinics.xlsx          Excel database: All Sites (latest year), Target List,
                           Physicians, Site Counts, Model (if fit)
  sites_by_year.csv        every site x year
  disappearance_panel.csv  site-year panel for the t -> t+2 model
  target_list.csv/.xlsx    likely-freestanding, small-org sites
  model_table1.csv         Table 1 replication (if statsmodels + pre-2022 data)

Usage:
  python ro_sites.py \
      --ndf 2019=DAC_NationalDownloadableFile_2019.csv \
      --ndf 2021=DAC_NationalDownloadableFile_2021.csv \
      --ndf 2025=DAC_NationalDownloadableFile.csv \
      --pup MUP_PHY_R25_P05_V10_D23_Prov_Svc.csv \
      --zip-county ZIP_COUNTY_122024.xlsx \
      --rucc Ruralurbancontinuumcodes2023.csv \
      [--google-key KEY] [--max-org-size 50] --out out/
"""
import argparse, json, os, re, sys, time
import numpy as np
import pandas as pd

# ------------------------------------------------------------------ columns
# NDF headers changed across years; map every known alias to one name.
# Older archive exports use long human-readable headers ("Line 1 Street
# Address"), 2020-2021 use short codes (adr_ln_1), 2022+ mix both.
ALIASES = {
    "npi": ["NPI"],
    "last": ["Provider Last Name", "lst_nm", "Last Name"],
    "first": ["Provider First Name", "frst_nm", "First Name"],
    "cred": ["Cred", "Credential"],
    "grd_yr": ["Grd_yr", "Graduation year"],
    "pri_spec": ["pri_spec", "Primary specialty"],
    "sec_spec_all": ["sec_spec_all", "All secondary specialties"],
    "org_name": ["Facility Name", "org_nm", "Organization legal name"],
    "org_pac_id": ["org_pac_id", "Group Practice PAC ID"],
    "num_org_mem": ["num_org_mem", "Number of Group Practice members"],
    "adr1": ["adr_ln_1", "Line 1 Street Address"],
    "adr2": ["adr_ln_2", "Line 2 Street Address"],
    "city": ["City/Town", "cty", "City"],
    "state": ["State", "st"],
    "zip": ["ZIP Code", "zip", "Zip Code"],
    "phone": ["Telephone Number", "phn_numbr", "Phone Number"],
}
# Hospital CCNs, present through 2021 (short or long header form).
HOSP_PAT = re.compile(r"^(hosp_afl_|Hospital affiliation CCN )\d+$", re.I)

RO_SPEC = "RADIATION ONCOLOGY"

# Delivery / management codes (pre-2026 code set, matching PUP data years).
RO_CODES = {"77427", "77385", "77386", "77402", "77407", "77412",
            "77373", "G6015", "G6003", "G6007", "G6011"}


def _resolve(cols):
    lower = {c.lower().strip(): c for c in cols}
    out = {}
    for canon, alist in ALIASES.items():
        for a in alist:
            if a.lower() in lower:
                out[lower[a.lower()]] = canon
                break
    hosp = [c for c in cols if HOSP_PAT.match(c.strip())]
    missing = {"npi", "pri_spec", "adr1", "zip"} - set(out.values())
    if missing:
        sys.exit(f"NDF is missing required columns {sorted(missing)}; "
                 f"add aliases to ALIASES. Header was: {cols[:40]}")
    return out, hosp


# ------------------------------------------------------------------ address
SUFFIX = {"STREET": "ST", "AVENUE": "AVE", "ROAD": "RD", "DRIVE": "DR",
          "BOULEVARD": "BLVD", "LANE": "LN", "PARKWAY": "PKWY", "HIGHWAY": "HWY",
          "COURT": "CT", "PLACE": "PL", "SUITE": "STE", "CIRCLE": "CIR",
          "NORTH": "N", "SOUTH": "S", "EAST": "E", "WEST": "W",
          "TERRACE": "TER", "EXPRESSWAY": "EXPY", "FREEWAY": "FWY",
          "SQUARE": "SQ", "TRAIL": "TRL", "PIKE": "PIKE", "CENTER": "CTR"}
UNIT = re.compile(r"\b(STE|SUITE|UNIT|APT|BLDG|FL|FLOOR|RM|ROOM)\b\s*\S*.*$|#.*$")


DIRS = {"N", "S", "E", "W", "NE", "NW", "SE", "SW"}


def norm_addr(a1, zip5):
    """Site key. Numbered addresses collapse to house number + first street-name
    word + ZIP, so '1000 JOHNSON FERRY RD NE' == '1000 Johnson Ferry Road',
    '1001 S GEORGE ST 2ND' == '1001 S George St', '100 CASA ST C' == '100 Casa St'."""
    s = str(a1 or "").upper()
    s = re.sub(r"[.,]", " ", s)
    s = UNIT.sub("", s)  # suite/floor -> same building = same site
    s = re.sub(r"\b(N|S|E|W|NE|NW|SE|SW)(\d)", r"\1 \2", s)  # NW22ND -> NW 22ND
    toks = [SUFFIX.get(t, t) for t in s.split()]
    m = re.match(r"^(\d+)[A-Z]?$", toks[0]) if toks else None
    name = [t for t in toks[1:] if t not in DIRS and len(t) > 1]
    if m and name:
        return f"{m.group(1)} {name[0]}|{zip5}"
    return " ".join(toks).strip() + "|" + zip5


def google_place_ids(df, key, cache_path):
    """Paper's approach: Address Validation API -> placeId. Cached."""
    import urllib.request
    cache = json.load(open(cache_path)) if os.path.exists(cache_path) else {}
    url = f"https://addressvalidation.googleapis.com/v1:validateAddress?key={key}"
    todo = df.drop_duplicates("addr_key")
    for _, r in todo.iterrows():
        k = r.addr_key
        if k in cache:
            continue
        body = {"address": {"regionCode": "US", "addressLines": [
            f"{r.adr1}, {r.city}, {r.state} {r.zip5}"]}}
        req = urllib.request.Request(url, json.dumps(body).encode(),
                                     {"Content-Type": "application/json"})
        try:
            res = json.load(urllib.request.urlopen(req, timeout=20))
            cache[k] = res.get("result", {}).get("geocode", {}).get("placeId") or k
        except Exception as e:  # fall back to normalized key
            print(f"  google fail {k}: {e}", file=sys.stderr)
            cache[k] = k
        time.sleep(0.02)
    json.dump(cache, open(cache_path, "w"))
    return df.addr_key.map(cache)


# ------------------------------------------------------------------ loaders
def load_ndf(path, year):
    head = pd.read_csv(path, nrows=0, dtype=str, encoding_errors="replace")
    colmap, hosp = _resolve(list(head.columns))
    need = list(colmap) + hosp
    parts = []
    for ch in pd.read_csv(path, usecols=need, dtype=str, chunksize=500_000,
                          encoding_errors="replace", low_memory=False):
        ch = ch.rename(columns=colmap)
        # primary specialty only: secondary RO pulls in diagnostic/IR radiologists
        parts.append(ch[ch["pri_spec"].fillna("").str.upper().str.strip() == RO_SPEC])
    df = pd.concat(parts, ignore_index=True)
    for c in ALIASES:  # uniform schema even when a year lacks a column
        if c not in df:
            df[c] = np.nan
    df["year"] = int(year)
    df["zip5"] = df["zip"].fillna("").str.extract(r"(\d{5})", expand=False).fillna("")
    df["addr_key"] = [norm_addr(a, z) for a, z in zip(df["adr1"], df["zip5"])]
    df["num_org_mem"] = pd.to_numeric(df["num_org_mem"], errors="coerce")
    if hosp:
        h = df[hosp].fillna("").apply(lambda s: s.str.strip())
        df["hosp_affiliated"] = (h != "").any(axis=1)
    else:
        df["hosp_affiliated"] = np.nan  # unknown (2022+ files)
    df = df.drop(columns=hosp)
    print(f"  {year}: {len(df):,} RO rows, {df.npi.nunique():,} RO NPIs, "
          f"affiliation {'present' if hosp else 'absent'}")
    return df


def load_crosswalk(zip_county, rucc):
    xw = (pd.read_excel(zip_county, dtype=str) if zip_county.endswith("xlsx")
          else pd.read_csv(zip_county, dtype=str))
    xw.columns = [c.upper() for c in xw.columns]
    xw["ZIP"] = xw["ZIP"].str.zfill(5)
    xw["COUNTY"] = xw["COUNTY"].str.zfill(5)
    xw["TOT_RATIO"] = pd.to_numeric(xw["TOT_RATIO"], errors="coerce")
    xw = (xw.sort_values("TOT_RATIO", ascending=False)
            .drop_duplicates("ZIP")[["ZIP", "COUNTY"]]
            .rename(columns={"ZIP": "zip5", "COUNTY": "fips"}))
    r = pd.read_csv(rucc, dtype=str, encoding_errors="replace")
    if "Attribute" in r.columns:  # 2023 file is long format
        r = r[r["Attribute"].str.upper().str.startswith("RUCC_2023")]
        r = r.rename(columns={"Value": "rucc"})
    else:
        rc = [c for c in r.columns if c.upper().startswith("RUCC")][0]
        r = r.rename(columns={rc: "rucc"})
    r["fips"] = r["FIPS"].str.zfill(5)
    r["rucc"] = pd.to_numeric(r["rucc"], errors="coerce")
    r["rurality"] = r["rucc"].map(lambda v: "urban" if v <= 3 else
                                  ("rural_adjacent" if v in (4, 6, 8) else
                                   ("rural_nonadjacent" if v in (5, 7, 9) else None)))
    return xw.merge(r[["fips", "rucc", "rurality"]], on="fips", how="left")


def office_billing(pup_path):
    """NPI-level: does this RO bill management/delivery in office setting?"""
    usecols = ["Rndrng_NPI", "Rndrng_Prvdr_Type", "HCPCS_Cd",
               "Place_Of_Srvc", "Tot_Srvcs"]
    parts = []
    for ch in pd.read_csv(pup_path, usecols=usecols, dtype=str, chunksize=1_000_000,
                          encoding_errors="replace", low_memory=False):
        ch = ch[ch["HCPCS_Cd"].isin(RO_CODES)]
        parts.append(ch)
    p = pd.concat(parts, ignore_index=True)
    p["Tot_Srvcs"] = pd.to_numeric(p["Tot_Srvcs"], errors="coerce").fillna(0)
    office = p["Place_Of_Srvc"].str.upper().eq("O")
    p["office_srvcs"] = p["Tot_Srvcs"].where(office, 0)
    p["facility_srvcs"] = p["Tot_Srvcs"].where(~office, 0)
    g = p.groupby("Rndrng_NPI")[["office_srvcs", "facility_srvcs"]].sum()
    g.index.name = "npi"
    return g.reset_index()


# ------------------------------------------------------------------ sites
def _join(s):
    return "; ".join(sorted(set(s.dropna().astype(str))))


def build_sites(rows):
    rows = rows.assign(physician=(rows["first"].fillna("") + " " + rows["last"].fillna("")
                                  + (", " + rows["cred"]).fillna("")).str.strip())
    agg = rows.groupby(["year", "site_id"]).agg(
        adr1=("adr1", "first"), adr2=("adr2", "first"), city=("city", "first"),
        state=("state", "first"), zip5=("zip5", "first"),
        phone=("phone", _join), org_names=("org_name", _join),
        org_pac_ids=("org_pac_id", _join),
        org_size=("num_org_mem", "max"),
        n_rad_oncs=("npi", "nunique"),
        physicians=("physician", _join),
        npis=("npi", lambda s: ";".join(sorted(set(s.dropna())))),
        hosp_affiliated=("hosp_affiliated",
                         lambda s: np.nan if s.isna().all() else bool(s.fillna(False).any())),
    ).reset_index()
    return agg


def panel(sites):
    years = sorted(sites.year.unique())
    present = sites.pivot_table(index="site_id", columns="year", values="n_rad_oncs",
                                aggfunc="size", fill_value=0).astype(bool)
    first = sites.groupby("site_id").year.min().rename("first_year")
    last = sites.groupby("site_id").year.max().rename("last_year")
    base = (sites.sort_values("year").drop_duplicates("site_id")
            [["site_id", "hosp_affiliated", "org_size", "state", "rurality"]]
            .set_index("site_id"))
    rec = []
    for t in years:
        if t + 2 not in years:
            continue
        ids = present.index[present[t]]
        gone = ~present.loc[ids, t + 2]
        for sid, g in gone.items():
            rec.append((sid, t, int(g)))
    pn = pd.DataFrame(rec, columns=["site_id", "year", "disappear"])
    pn = pn.join(base, on="site_id").join(first, on="site_id")
    return pn, first, last


def replicate_model(pn, out):
    try:
        import statsmodels.formula.api as smf
    except ImportError:
        print("  statsmodels not installed; skipping model"); return None
    use_rural = pn.rurality.notna().any()  # no crosswalk -> drop the rural term
    d = pn[(pn.first_year <= 2021) & pn.hosp_affiliated.notna()
           & (pn.rurality.notna() | (not use_rural))].copy()
    if d.empty or d.disappear.nunique() < 2:
        print("  not enough affiliation-era data to fit model"); return None
    if not use_rural:
        print("  no rurality data; fitting model without the rural term")
    d["freestanding"] = (~d.hosp_affiliated.astype(bool)).astype(int)
    d["rural"] = (d.rurality != "urban").astype(int)
    d["log_org"] = np.log(d.org_size.fillna(1).clip(lower=1))
    terms = ["freestanding", "rural", "log_org"] if use_rural else ["freestanding", "log_org"]
    fe = " + C(year)" if d.year.nunique() > 1 else ""
    try:
        m = smf.logit(f"disappear ~ {' + '.join(terms)} + C(state){fe}",
                      data=d).fit(disp=0, cov_type="cluster",
                                  cov_kwds={"groups": pd.factorize(d.site_id)[0]})
    except Exception as e:
        print(f"  model failed to fit: {e}"); return None
    res = pd.DataFrame({"OR": np.exp(m.params), "lo": np.exp(m.conf_int()[0]),
                        "hi": np.exp(m.conf_int()[1]), "p": m.pvalues})
    res = res.loc[terms]
    res.to_csv(os.path.join(out, "model_table1.csv"))
    print("  Table 1 replication:\n" + res.round(3).to_string())
    return res


# ------------------------------------------------------------------ excel
def write_workbook(path, sheets):
    """Multi-sheet Excel database with frozen headers, filters, sane widths."""
    from openpyxl.styles import Font
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for name, df in sheets.items():
            if df is None:
                continue
            df.to_excel(xw, sheet_name=name, index=name == "Model")
            ws = xw.sheets[name]
            ws.freeze_panes = "B2" if name == "Model" else "A2"
            ws.auto_filter.ref = ws.dimensions
            for col in ws.columns:
                w = max((len(str(c.value)) for c in col[:200] if c.value is not None), default=8)
                ws.column_dimensions[col[0].column_letter].width = min(max(w + 2, 8), 60)
            for c in ws[1]:
                c.font = Font(bold=True)


# ------------------------------------------------------------------ main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ndf", action="append", required=True, help="YEAR=path")
    ap.add_argument("--pup", help="Physician & Other Practitioners by Provider and Service CSV")
    ap.add_argument("--zip-county", help="HUD ZIP->COUNTY crosswalk")
    ap.add_argument("--rucc", help="USDA RUCC 2023 CSV")
    ap.add_argument("--google-key")
    ap.add_argument("--max-org-size", type=int, default=50,
                    help="exclude orgs larger than this from target list (health systems)")
    ap.add_argument("--out", default="out")
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)

    print("Loading NDF files")
    rows = pd.concat([load_ndf(p, y) for y, p in (s.split("=", 1) for s in a.ndf)],
                     ignore_index=True)
    rows["site_id"] = (google_place_ids(rows, a.google_key, os.path.join(a.out, "place_cache.json"))
                       if a.google_key else rows["addr_key"])

    sites = build_sites(rows)
    if a.zip_county and a.rucc:
        sites = sites.merge(load_crosswalk(a.zip_county, a.rucc), on="zip5", how="left")
    else:
        sites["fips"] = sites["rucc"] = sites["rurality"] = np.nan

    # carry baseline affiliation forward (paper: first observed year)
    base_aff = (sites.dropna(subset=["hosp_affiliated"]).sort_values("year")
                .drop_duplicates("site_id").set_index("site_id").hosp_affiliated)
    sites["baseline_hosp_affiliated"] = sites.site_id.map(base_aff)
    sites.to_csv(os.path.join(a.out, "sites_by_year.csv"), index=False)

    counts = sites.groupby("year").site_id.nunique().rename("sites").to_frame()
    counts["rad_oncs"] = rows.groupby("year").npi.nunique()
    print("Site counts by year:")
    print(counts.to_string())

    pn, first, last = panel(sites)
    pn.to_csv(os.path.join(a.out, "disappearance_panel.csv"), index=False)
    model = replicate_model(pn, a.out)

    # ---------------- current sites (latest year) + freestanding flag
    latest = sites.year.max()
    cur = sites[sites.year == latest].copy()
    cur = cur.join(first, on="site_id")
    if a.pup:
        print("Scanning PUP for office-setting RO billing")
        ob = office_billing(a.pup)
        npi_site = rows[rows.year == latest][["npi", "site_id"]].drop_duplicates()
        s = npi_site.merge(ob, on="npi").groupby("site_id")[["office_srvcs", "facility_srvcs"]].sum()
        cur = cur.join(s, on="site_id")
        cur[["office_srvcs", "facility_srvcs"]] = cur[["office_srvcs", "facility_srvcs"]].fillna(0)
        cur["office_share"] = cur.office_srvcs / (cur.office_srvcs + cur.facility_srvcs).replace(0, np.nan)
    else:
        cur["office_srvcs"] = cur["facility_srvcs"] = cur["office_share"] = np.nan

    cur["likely_freestanding"] = np.select(
        [cur.office_share >= 0.5, cur.baseline_hosp_affiliated == False,
         cur.office_share < 0.5, cur.baseline_hosp_affiliated == True],
        ["yes (office billing)", "yes (pre-2022 NDF)", "no (facility billing)",
         "no (pre-2022 NDF)"], default="unknown")
    cols = ["org_names", "adr1", "adr2", "city", "state", "zip5", "phone", "n_rad_oncs",
            "physicians", "org_size", "likely_freestanding", "office_srvcs",
            "facility_srvcs", "office_share", "rurality", "rucc", "fips",
            "first_year", "org_pac_ids", "npis", "site_id"]
    all_cur = (cur.sort_values(["state", "city", "adr1"])
               [[c for c in cols if c in cur]])

    tgt = cur[(cur.org_size.fillna(0) <= a.max_org_size)
              & ~cur.likely_freestanding.str.startswith("no")]
    tgt = tgt.sort_values(["likely_freestanding", "office_srvcs", "n_rad_oncs"],
                          ascending=[True, False, False])
    tgt = tgt[[c for c in cols if c in tgt and c != "site_id"]]
    tgt.to_csv(os.path.join(a.out, "target_list.csv"), index=False)
    try:
        tgt.to_excel(os.path.join(a.out, "target_list.xlsx"), index=False)
    except Exception as e:
        print(f"  target_list.xlsx not written: {e}")

    # one row per physician x site in the latest year
    docs = (rows[rows.year == latest]
            .drop_duplicates(["npi", "site_id"])
            [["npi", "first", "last", "cred", "grd_yr", "pri_spec", "org_name",
              "adr1", "city", "state", "zip5", "phone", "site_id"]]
            .sort_values(["state", "last", "first"]))

    wb = os.path.join(a.out, "ro_clinics.xlsx")
    write_workbook(wb, {
        f"All Sites {latest}": all_cur,
        "Target List": tgt,
        "Physicians": docs,
        "Site Counts": counts.reset_index(),
        "Model": model,
    })
    print(f"All sites ({latest}): {len(all_cur):,} -> {wb}")
    print(f"Target list: {len(tgt):,} sites -> {a.out}/target_list.csv")


if __name__ == "__main__":
    main()
