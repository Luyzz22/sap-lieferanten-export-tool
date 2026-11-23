# Supplier reporting prototype

Python CSV-to-Excel reporting with duplicate removal, missing-field checks and an illustrative company-name matcher. This repository is an earlier standalone experiment; it is not a production sanctions-screening service or a direct SAP integration.

## Implemented behavior

The script reads `sap_lieferanten.csv`, keeps the first row for each supplier number and creates a workbook with `DASHBOARD`, `Alle Lieferanten`, `Nach Land` and `Pruefliste`. A `SANKTIONSLISTE` worksheet is added only when the demonstration matcher finds a match.

The matcher removes a fixed set of company suffixes, compares word sets and skips reference names with fewer than two remaining words. It does not resolve legal identities, transliterations or changing official lists. A missing or unreadable reference file currently produces an empty list; an absence of matches must not be interpreted as clearance.

## Run the example

```bash
git clone https://github.com/Luyzz22/sap-lieferanten-export-tool.git
cd sap-lieferanten-export-tool
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python create_test_data.py
python sap_export_tool.py
```

On Windows, activate with `.venv\Scripts\activate`. The generator overwrites `sap_lieferanten.csv`. The included `eu_sanctions.csv` is **synthetic demonstration data**, despite its historical filename. It is not an official list and makes no assertion about any real company.

Input columns used by the program: `Lieferantennr`, `Name`, `Land`, `Telefon`. Other columns are retained in the workbook. Outputs are `Lieferanten_Report_<timestamp>.xlsx` and `sap_export_log.txt`.

## Verification

```bash
python -m unittest discover -s tests -v
```

Tests execute the report command against temporary synthetic inputs, including a duplicate in the middle of the input and a missing name. They check row counts, review rows, matching output and row-color alignment. GitHub Actions runs the same checks with Python 3.12.

## Scope and limitations

The repository does not implement official-list updates, an SAP RFC client, access control, audit-log retention, cloud deployment or production monitoring. Earlier accuracy, savings, production-readiness and compliance claims were not backed by a reproducible evaluation and have been removed.

The name matcher can miss relevant entries and match unrelated names. It must not be used for sanctions decisions. Reports and verbose logs can contain input names; use synthetic data for development and apply an appropriate data-handling process before any real use.

Generated datasets, workbooks and logs are excluded from new commits. Historical versions remain in Git history. Current SBS offerings are listed at [sbsdeutschland.com](https://sbsdeutschland.com/).
