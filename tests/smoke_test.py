"""End-to-end smoke test of ro_sites.py on small synthetic CMS-shaped files."""
import os, subprocess, sys, tempfile
import numpy as np
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
rng = np.random.default_rng(0)


def make(d):
    states = ["TX", "OH", "CA", "NY", "GA"]
    sites = [(f"{100 + i} Main Street", f"Town{i}", states[i % 5], f"{75000 + i:05d}")
             for i in range(60)]
    files = {}
    for yr in (2019, 2021, 2023, 2025):
        rec = []
        for i, (adr, cty, st, z) in enumerate(sites):
            if rng.random() < (0.15 if i % 3 else 0.05):  # some sites vanish
                continue
            for k in range(1 + i % 3):
                rec.append({"npi": str(1000000000 + i * 10 + k), "lst": f"Doc{i}_{k}",
                            "frst": "A", "adr": adr + (" Suite 2" if k else ""),
                            "cty": cty, "st": st, "zip": z + "1234",
                            "org": f"Org {i // 2}", "pac": str(900 + i // 2),
                            "mem": str(5 + i % 80), "ccn": "" if i % 3 else "450001"})
        rec.append({"npi": "1999999999", "lst": "Cardio", "frst": "B", "adr": "1 X St",
                    "cty": "Y", "st": "TX", "zip": "75001", "org": "", "pac": "",
                    "mem": "", "ccn": ""})
        df = pd.DataFrame(rec)
        spec = np.where(df.lst == "Cardio", "CARDIOLOGY", "RADIATION ONCOLOGY")
        if yr == 2019:  # long-header archive style
            out = pd.DataFrame({"NPI": df.npi, "Last Name": df.lst, "First Name": df.frst,
                                "Credential": "MD", "Primary specialty": spec,
                                "All secondary specialties": "",
                                "Organization legal name": df.org,
                                "Group Practice PAC ID": df.pac,
                                "Number of Group Practice members": df.mem,
                                "Line 1 Street Address": df.adr, "Line 2 Street Address": "",
                                "City": df.cty, "State": df.st, "Zip Code": df.zip,
                                "Phone Number": "5550001111",
                                "Hospital affiliation CCN 1": df.ccn,
                                "Hospital affiliation LBN 1": ""})
        elif yr == 2021:  # short-code style with hosp_afl
            out = pd.DataFrame({"NPI": df.npi, "lst_nm": df.lst, "frst_nm": df.frst,
                                "Cred": "MD", "pri_spec": spec, "sec_spec_all": "",
                                "org_nm": df.org, "org_pac_id": df.pac, "num_org_mem": df.mem,
                                "adr_ln_1": df.adr, "adr_ln_2": "", "cty": df.cty, "st": df.st,
                                "zip": df.zip, "phn_numbr": "5550001111",
                                "hosp_afl_1": df.ccn, "hosp_afl_lbn_1": ""})
        else:  # current style, no affiliation
            out = pd.DataFrame({"NPI": df.npi, "Provider Last Name": df.lst,
                                "Provider First Name": df.frst, "Cred": "MD", "pri_spec": spec,
                                "sec_spec_all": "", "Facility Name": df.org,
                                "org_pac_id": df.pac, "num_org_mem": df.mem, "adr_ln_1": df.adr,
                                "adr_ln_2": "", "City/Town": df.cty, "State": df.st,
                                "ZIP Code": df.zip, "Telephone Number": "5550001111"})
        p = os.path.join(d, f"ndf_{yr}.csv"); out.to_csv(p, index=False); files[yr] = p
    last = pd.read_csv(files[2025], dtype=str)
    pup = pd.DataFrame({"Rndrng_NPI": last.NPI, "Rndrng_Prvdr_Type": "Radiation Oncology",
                        "HCPCS_Cd": "77427",
                        "Place_Of_Srvc": np.where(last.NPI.str[-2].astype(int) % 2, "O", "F"),
                        "Tot_Srvcs": "40"})
    pup.to_csv(os.path.join(d, "pup.csv"), index=False)
    pd.DataFrame({"ZIP": [s[3] for s in sites], "COUNTY": [f"{48001 + i}" for i in range(60)],
                  "TOT_RATIO": 1.0}).to_csv(os.path.join(d, "xw.csv"), index=False)
    pd.DataFrame({"FIPS": [f"{48001 + i}" for i in range(60)], "Attribute": "RUCC_2023",
                  "Value": [str(1 + i % 9) for i in range(60)]}).to_csv(
        os.path.join(d, "rucc.csv"), index=False)
    return files


def main():
    with tempfile.TemporaryDirectory() as d:
        f = make(d)
        out = os.path.join(d, "out")
        cmd = [sys.executable, os.path.join(HERE, "..", "ro_sites.py")]
        for y, p in f.items():
            cmd += ["--ndf", f"{y}={p}"]
        cmd += ["--pup", os.path.join(d, "pup.csv"), "--zip-county", os.path.join(d, "xw.csv"),
                "--rucc", os.path.join(d, "rucc.csv"), "--out", out]
        subprocess.run(cmd, check=True)
        x = pd.read_excel(os.path.join(out, "ro_clinics.xlsx"), sheet_name=None)
        print({k: v.shape for k, v in x.items()})
        allsites = x["All Sites 2025"]
        assert "CARDIO" not in " ".join(x["Physicians"]["last"].astype(str)).upper()
        # suites collapse into one site: never more sites than distinct buildings
        assert allsites.site_id.is_unique and len(allsites) <= 60
        assert allsites.rurality.notna().all()
        assert set(allsites.likely_freestanding) <= {
            "yes (office billing)", "no (facility billing)", "yes (pre-2022 NDF)",
            "no (pre-2022 NDF)", "unknown"}
        assert os.path.exists(os.path.join(out, "model_table1.csv"))
        print("SMOKE TEST PASSED")


if __name__ == "__main__":
    main()
