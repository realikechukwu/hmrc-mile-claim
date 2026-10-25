# HMRC Mile Claim

Get back tax on the miles you drove to hospital placements.

If you are an NHS doctor in training who drives from home to a different hospital every few months, HMRC lets you claim tax relief on those miles. Most people never do. This tool works out the claim for you from your own payslips and placement list. You use it by chatting with an AI assistant on your own computer.

**Interactive guide with tick boxes:** open `docs/index.html` in your browser, or visit https://realikechukwu.github.io/hmrc-mile-claim/ if the page is switched on.

## Is this for you?
- You are on an NHS rotation, and each hospital placement is under 24 months.
- You drive your own car from home to the hospital.
- Your employer has never paid you mileage or "excess travel".
- You have your payslips (the NHS ESR kind) from April 2022 onwards.
- You use a Mac. On Windows you need WSL, which is harder.

## What you need
1. An AI app that can open a folder on your computer and run commands. Examples are the Claude desktop app with Claude Code, or ChatGPT's Codex app. A normal chat window where you only upload files will not work.
2. Your documents: payslip PDFs, your placement list from TIS saved as a PDF, your home addresses with move dates, and Google Maps distances.
3. About half a day, spread over a few sessions.

## How to use it
1. **Download this folder.** Click the green Code button, then Download ZIP. Unzip it.
2. **Put your documents in `my-documents/`.** Payslips go in `payslips/`, your TIS page and ESR pages go in `placements/`, and copy `templates/homes.csv` to `my-documents/homes.csv` and add your addresses. Read `my-documents/README.txt`.
3. **Open the unzipped folder in your AI app.** Then copy the message in `PROMPT.md` and paste it in.
4. **Answer its questions.** It stops at checkpoints so you can check everything. Take your time.
5. **Read your results** in `my-documents/tax-claim/`: `summary.md` (the answer), `claim_summary.xlsx` (a spreadsheet you can edit), `issues.md` (things to check), `checklist.md` (what to do next) and `sar_emails/` (emails to ask hospitals for proof of the days you worked).

## Your privacy
- The scripts here never send anything anywhere. They run on your computer.
- The AI app you choose will read your documents to help you. That is covered by its own privacy terms, so check them first.
- Your documents and results stay in `my-documents/`, which git ignores, so they are never published by accident.

## What it cannot do
- It cannot tell you for certain that a hospital counts as a "temporary workplace". It flags risks. Check with HMRC or an accountant.
- Days of attendance are estimated from working days and leave, unless you supply rota records.
- It only reads ESR payslips. Scottish tax bands are not included.
- It is not tax advice. You are responsible for what you file.

## For helpers and developers
`AGENTS.md` has the assistant's full instructions. `tool/build_claim.py` has the commands. Run `python3 tool/build_claim.py --help`. MIT licence.
