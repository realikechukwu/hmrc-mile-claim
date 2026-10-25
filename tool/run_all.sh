#!/bin/bash
# Usage: tool/run_all.sh [person-folder] ["ASSIGNMENT=SHORTNAME,..."]
# Default person folder is my-documents. Layout:
#   payslips/      payslip PDFs, any depth
#   placements/    TIS placements printout (PDF) and any ESR assignment printouts
#   homes.csv      lines: postcode,from,to   (to left blank for the current address)
#   tax-claim/     created here: everything the tool writes
# The first run stops after writing a draft config.yaml. Fill the TODOs, run again.
# Your original files are never changed. Nothing leaves your computer.
set -euo pipefail
TOOL="$(cd "$(dirname "$0")" && pwd)"
P="${1:-$TOOL/../my-documents}"; EMAP="${2:-}"
PDF2MD="$TOOL/pdf2md.sh"
if [ ! -d "$P/payslips" ]; then
  echo "I cannot find $P/payslips. Create it and put your payslip PDFs inside, then run this again."; exit 1
fi
OUT="$P/tax-claim"; RAW="$OUT/_raw"; mkdir -p "$RAW"

convert_dir () {  # convert the PDFs directly inside $1 into $RAW/$2
  local n; n=$(find "$1" -maxdepth 1 -iname '*.pdf' 2>/dev/null | wc -l | tr -d ' ')
  if [ "$n" -gt 0 ]; then "$PDF2MD" -o "$RAW/$2" "$1"/*.pdf | grep -v '^ok' || true; fi
}
i=0
while IFS= read -r -d '' d; do i=$((i+1)); convert_dir "$d" "pay$i"; done \
  < <(find "$P/payslips" -type d -print0)
convert_dir "$P/placements" placements

python3 "$TOOL/build_claim.py" prepare --raw "$RAW" --out "$OUT/payslips_md" ${EMAP:+--employer-map "$EMAP"}
python3 "$TOOL/build_claim.py" extract --md "$OUT/payslips_md" --out "$OUT"

if [ ! -f "$OUT/config.yaml" ]; then
  TIS=$(ls "$RAW"/placements/*tis*.md 2>/dev/null | head -1 || true)
  python3 "$TOOL/build_claim.py" init --out "$OUT" ${TIS:+--tis "$TIS"} --homes "$P/homes.csv"
  echo "Next: fill every TODO in $OUT/config.yaml, then run this script again."
  exit 0
fi
python3 "$TOOL/build_claim.py" all --md "$OUT/payslips_md" --out "$OUT"
echo "Done. Open $OUT/summary.md and $OUT/claim_summary.xlsx"
