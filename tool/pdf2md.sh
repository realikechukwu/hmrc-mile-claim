#!/bin/zsh
# Usage: pdf2md.sh [-f] [-o outdir] <file.pdf | dir> ...
# Converts PDFs to .md next to the source (or into outdir). Text layer via pdftotext;
# falls back to tesseract OCR for scanned pages. Skips existing .md unless -f.
force=0; outdir=""
while getopts "fo:" o; do case $o in f) force=1;; o) outdir=$OPTARG;; esac; done
shift $((OPTIND-1))
[[ $# -eq 0 ]] && set -- .

convert() {
  local pdf=$1 dest="${outdir:-${1:h}}/${${1:t}%.[pP][dD][fF]}.md"
  [[ -n $outdir ]] && mkdir -p "$outdir"
  if [[ -e $dest && $force -eq 0 ]]; then echo "skip  $dest"; return; fi
  local npages=$(pdfinfo "$pdf" 2>/dev/null | awk '/^Pages:/{print $2}')
  local tmp=$(mktemp -d) out="" p txt
  for ((p=1; p<=${npages:-1}; p++)); do
    txt=$(pdftotext -f $p -l $p -layout "$pdf" - 2>/dev/null)
    if [[ $(print -r -- "$txt" | wc -w) -lt 15 ]] && command -v tesseract >/dev/null; then
      pdftoppm -r 200 -f $p -l $p -png "$pdf" "$tmp/pg" 2>/dev/null
      txt=$(tesseract "$(print $tmp/pg*.png | awk '{print $1}')" - 2>/dev/null)
      rm -f $tmp/pg*.png
    fi
    out+="$txt"$'\n\n'
  done
  rm -rf "$tmp"
  { echo "# ${${pdf:t}%.[pP][dD][fF]}"; echo; print -r -- "$out" | cat -s; } > "$dest"
  echo "ok    $dest ($(wc -w < "$dest") words)"
}

for arg in "$@"; do
  if [[ -d $arg ]]; then
    find "$arg" -iname '*.pdf' -print0 | while IFS= read -r -d '' f; do convert "$f"; done
  else convert "$arg"; fi
done
