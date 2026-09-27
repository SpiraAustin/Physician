#!/usr/bin/env python3
"""
fetch_data.py - Download the public inputs ro_sites.py needs into data/.

  NDF (current)  data.cms.gov provider-data API, dataset mj5m-pzi6
  NDF (archive)  data.cms.gov provider-data archive, December release per year
  PUP            data.cms.gov "Medicare Physician & Other Practitioners - by
                 Provider and Service" (latest year), pulled via the data API
                 filtered to Radiation Oncology so it's MBs, not ~3 GB
  RUCC 2023      USDA ERS
  HUD ZIP-COUNTY Needs a free HUD USER API token (--hud-token), or download
                 manually from https://www.huduser.gov/portal/datasets/usps_crosswalk.html

Archive URLs move occasionally; if one 404s, grab the file by hand from
https://data.cms.gov/provider-data/archived-data/doctors-clinicians and pass
it to ro_sites.py with --ndf YEAR=path.
"""
import argparse, csv, datetime, json, os, sys, urllib.parse, urllib.request, zipfile

UA = {"User-Agent": "ro-sites/1.0"}
NDF_META = "https://data.cms.gov/provider-data/api/1/metastore/schemas/dataset/items/mj5m-pzi6"
NDF_ARCHIVE = ("https://data.cms.gov/provider-data/sites/default/files/archive/"
               "Doctors%20and%20clinicians/{y}/doctors_and_clinicians_12_{y}.zip")
PUP_TITLE = "Medicare Physician & Other Practitioners - by Provider and Service"
RUCC_URL = ("https://ers.usda.gov/sites/default/files/_laserfiche/DataFiles/53251/"
            "Ruralurbancontinuumcodes2023.csv")
HUD_API = "https://www.huduser.gov/hudapi/public/usps?type=2&query=All"


def get(url, **kw):
    req = urllib.request.Request(url, headers={**UA, **kw.pop("headers", {})})
    return urllib.request.urlopen(req, timeout=kw.pop("timeout", 600))


def save(url, path):
    if os.path.exists(path):
        print(f"  have {path}"); return path
    print(f"  GET {url}")
    with get(url) as r, open(path + ".part", "wb") as f:
        while chunk := r.read(1 << 20):
            f.write(chunk)
    os.replace(path + ".part", path)
    return path


def ndf_current(d):
    meta = json.load(get(NDF_META))
    url = next(x["downloadURL"] for x in meta["distribution"] if x["downloadURL"].endswith(".csv"))
    return save(url, os.path.join(d, "ndf_current.csv"))


def ndf_year(d, y):
    out = os.path.join(d, f"ndf_{y}.csv")
    if os.path.exists(out):
        print(f"  have {out}"); return out
    z = save(NDF_ARCHIVE.format(y=y), os.path.join(d, f"ndf_{y}.zip"))
    with zipfile.ZipFile(z) as zf:
        name = max((n for n in zf.namelist() if n.lower().endswith(".csv")
                    and "national" in n.lower()), key=lambda n: zf.getinfo(n).file_size)
        with zf.open(name) as src, open(out, "wb") as dst:
            while chunk := src.read(1 << 20):
                dst.write(chunk)
    return out


def pup(d):
    out = os.path.join(d, "pup_radonc.csv")
    if os.path.exists(out):
        print(f"  have {out}"); return out
    cat = json.load(get("https://data.cms.gov/data.json"))
    ds = next(x for x in cat["dataset"] if x["title"].strip() == PUP_TITLE)
    api = next(x["accessURL"] for x in ds["distribution"]
               if x.get("format") == "API" and "data-api" in x.get("accessURL", ""))
    print(f"  PUP API {api}")
    rows, off = [], 0
    while True:
        q = urllib.parse.urlencode({"filter[Rndrng_Prvdr_Type]": "Radiation Oncology",
                                    "offset": off, "size": 5000})
        page = json.load(get(f"{api}?{q}"))
        if not page:
            break
        rows += page; off += len(page)
        print(f"    {off:,} rows", end="\r")
    with open(out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0])); w.writeheader(); w.writerows(rows)
    print(f"\n  wrote {out}")
    return out


def hud(d, token):
    out = os.path.join(d, "zip_county.csv")
    if os.path.exists(out):
        print(f"  have {out}"); return out
    res = json.load(get(HUD_API, headers={"Authorization": f"Bearer {token}"}))
    rows = res["data"]["results"]
    with open(out, "w", newline="") as f:
        w = csv.DictWriter(f, fieldnames=[k.upper() for k in rows[0]])
        w.writeheader(); w.writerows({k.upper(): v for k, v in r.items()} for r in rows)
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--years", default="2018,2019,2020,2021,2022,2023,2024,2025",
                    help="archive NDF years (December release); current file always fetched")
    ap.add_argument("--hud-token", default=os.environ.get("HUD_TOKEN"))
    ap.add_argument("--dir", default="data")
    a = ap.parse_args()
    os.makedirs(a.dir, exist_ok=True)
    args = []
    steps = [("NDF current", lambda: args.append(f"--ndf {datetime.date.today().year}={ndf_current(a.dir)}"))]
    for y in a.years.split(","):
        steps.append((f"NDF {y}", lambda y=y: args.append(f"--ndf {y}={ndf_year(a.dir, y)}")))
    steps += [("PUP", lambda: args.append(f"--pup {pup(a.dir)}")),
              ("RUCC", lambda: args.append(
                  f"--rucc {save(RUCC_URL, os.path.join(a.dir, 'rucc2023.csv'))}"))]
    if a.hud_token:
        steps.append(("HUD", lambda: args.append(f"--zip-county {hud(a.dir, a.hud_token)}")))
    else:
        print("No --hud-token: download ZIP_COUNTY xlsx manually for rurality.")
    failed = []
    for name, fn in steps:
        print(name)
        try:
            fn()
        except Exception as e:
            print(f"  FAILED: {e}", file=sys.stderr); failed.append(name)
    print("\nRun:\n  python ro_sites.py " + " \\\n    ".join(args) + " \\\n    --out out")
    if failed:
        print(f"\nFailed: {', '.join(failed)} (download manually)", file=sys.stderr)


if __name__ == "__main__":
    main()
