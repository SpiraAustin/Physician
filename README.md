# US Radiation Oncology Clinics Database

Builds an Excel database of every radiation oncology practice site in the US
from public CMS data, following Yu et al. (IJROBP 2026, "Structural
Vulnerability in the US RO Delivery System"), and flags likely freestanding
(non-hospital) practices.

## Run

```bash
pip install -r requirements.txt
python fetch_data.py --hud-token $HUD_TOKEN   # downloads into data/, prints the ro_sites.py command
python ro_sites.py --ndf 2019=data/ndf_2019.csv ... --out out   # paste the printed command
```

A HUD USER API token is free at https://www.huduser.gov/portal/dataset/uspszip-api.html.
Without it, download the ZIP-COUNTY crosswalk xlsx by hand and pass `--zip-county`.

## Output: `out/ro_clinics.xlsx`

| Sheet | Contents |
|---|---|
| All Sites *YYYY* | Every RO site in the latest NDF: address, phone, orgs, # rad oncs, physicians, org size, freestanding flag, office/facility billing, rurality |
| Target List | Likely-freestanding sites with org size <= `--max-org-size` (default 50) |
| Physicians | One row per radiation oncologist x site (latest year) |
| Site Counts | Sites and rad oncs per year |
| Model | Replication of the paper's Table 1 (odds ratios for site disappearance) |

Also written: `sites_by_year.csv`, `disappearance_panel.csv`, `target_list.csv/.xlsx`,
`model_table1.csv`.

## Notes

- A "site" is a normalized street address + ZIP (suite/floor stripped). Pass
  `--google-key` to use Google Address Validation placeIds, as the paper did.
- Hospital affiliation only exists in NDFs through 2021; later sites get
  "freestanding" from office-setting (POS = O) billing of 77427/delivery codes in the PUP file.
- `python tests/smoke_test.py` runs the pipeline end to end on synthetic data.
