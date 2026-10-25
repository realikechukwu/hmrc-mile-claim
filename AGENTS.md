# Instructions for the AI assistant

You are helping a non-technical NHS doctor (or their partner) prepare an HMRC claim for mileage allowance relief: the miles driven from home to temporary hospital placements. Follow this file exactly. The user will paste a starter message that points here.

## Ground rules
- Everything stays local. No uploads, no web services that receive the user's data, no git commits of their documents.
- Never modify or move the user's original files. All output goes to `my-documents/tax-claim/`.
- Never invent data. Anything missing or unreadable goes in `my-documents/tax-claim/issues.md` and you ask the user.
- Stop at every CHECKPOINT and wait for the user's reply.
- Talk in plain English. No jargon without a one-line explanation. Short messages. Offer to explain.
- Do not use em dashes in anything you write for the user.
- Ask before installing anything. Ask for a password if a PDF is protected, and never store it.
- This is a calculation and evidence aid, not tax advice. Flag borderline points for HMRC or an accountant. Where evidence is weak, use the lower estimate and say so.
- Do not look up distances or postcodes online unless the user says you may.

## The tool
Everything is in `tool/`:
- `tool/run_all.sh` converts PDFs, extracts payslips and writes a draft config on the first run, then runs the checks and the calculation on the second run.
- `tool/build_claim.py` has the steps as commands: `prepare`, `extract`, `init`, `check`, `attendance`, `calc`, `evidence`, `all`.
- It needs `python3`, `pdftotext` (poppler) and, for the spreadsheet, `openpyxl`. `tesseract` is optional for scanned PDFs. It is built for Mac. On Windows the user needs WSL.
Use the tool. Do not write your own PDF converter. If conversion fails on a file, log it in `issues.md` and tell the user. Do not switch method without asking.

## Steps
Step 0, plan. Give a short numbered plan and the folders you will create. Check the tools above are installed and say what is missing. Wait for OK.

Step 1, convert. Check `my-documents/` has `payslips/`, `placements/` and `homes.csv`. Run `tool/run_all.sh` (add `"ASSIGNMENT=SHORTNAME"` pairs only if the user names their employers). Open `tax-claim/conversion_report.md`. Report: files converted, failures, duplicates, missing months. Show one sample converted payslip. CHECKPOINT 1.

Step 2, payslip data. Show the per-tax-year summary from `payslips.csv` (pay, taxable pay, tax, parking, any travel lines) and every line in `payslip_line_codes.md` that is unclassified. Ask what each unclassified line is. Confirm whether any mileage or excess travel was ever paid (it should be zero). Parking counts only if taken after tax; the tool tests this. CHECKPOINT 2.

Step 3, placements and risks. `init` drafts `tax-claim/config.yaml` from the TIS printout. Ask the user for what files cannot give: hospital postcodes, home address move dates, own car and car-sharing (only the driver can claim), annual and study leave per year, regional teaching or other regular trips, any rota exports (optional), and unreimbursed tolls and parking. Add placements missing from TIS (for example a current post or a local fellow job).
Then check each placement against the temporary workplace rules and log risks in `issues.md`:
- the same site for more than 24 months, including repeated returns that add up to roughly 40 percent of working time over 24 months
- consecutive sites where the journey from home is much the same (HMRC may treat them as one workplace)
- a job where all or almost all of the employment was at one site (a permanent workplace, so not claimable)
- gaps between placements, or a period when the employer was not the lead employer
- leave periods such as parental leave (no attendance)
Show the placement table and flagged risks. Set `claim: true/false` and a `risk` label per placement with the user. CHECKPOINT 3.

Step 4, distances. For each home postcode and site pair, ask the user for the normal-route return distance in miles from Google Maps. Show them how: driving directions, home postcode to hospital postcode, double the one-way figure (or add both directions). Put the figures in `config.yaml`. Run `python3 tool/build_claim.py check --config my-documents/tax-claim/config.yaml --out my-documents/tax-claim` and fix every problem it lists. CHECKPOINT 4: confirm the distance table.

Step 5, attendance and calculation. Use rota exports if the user has them. Otherwise use the working-day method, and say clearly that dates are placeholders, not real shifts. Run `tool/run_all.sh` again. Review `attendance_crosscheck.md` and explain any flags. Present both cases from `summary.md`: all claimable placements, and the lower case that drops the riskiest site. Rates: 45p a mile for the first 10,000 miles in a tax year and 25p after. Route: online form P87 if a year's expenses are 2,500 pounds or less, otherwise Self Assessment. Deadline: 4 years after the tax year ends (5 April).

Step 6, the pack. Write `evidence_index.md` (already generated), `checklist.md`, `README.md` (method and every assumption) and `sar_emails/`: one subject access request email per host trust and one for the lead employer, under UK GDPR Article 15, asking for rostered and worked shifts and annual and study leave for the placement dates, electronic export preferred, with placeholders for the GMC number, date of birth and contact details. The checklist starts with the earliest deadline and includes: calling HMRC on 0300 200 3310 to confirm how to file back years, and whether Self Assessment registration is needed. FINAL CHECKPOINT: show `summary.md` and `issues.md` and list what the user must supply or check before filing.

## Reusing for another person
Each person gets their own folder copy of `my-documents/` and their own `config.yaml`. Never mix two people's files.
