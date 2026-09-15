#!/usr/bin/env bash
set -e
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO="$(cd "$HERE/.." && pwd)"
export HSBG_WORK="${HSBG_WORK:-$REPO/build}"
DOCX="$REPO/Ho_so_bai_giang_lai_xe_oto_2026_v3.docx"
for i in 1 2 3; do
  echo "── Lượt $i ──"
  python3 "$HERE/build_v3.py" >/dev/null
  rm -rf "$HSBG_WORK/pdf" && mkdir -p "$HSBG_WORK/pdf"
  soffice --headless --norestore --convert-to pdf --outdir "$HSBG_WORK/pdf" "$DOCX" >/dev/null 2>&1
  cp -f "$HSBG_WORK/toc_v3.json" "$HSBG_WORK/toc_v3_prev.json" 2>/dev/null || true
  python3 "$HERE/measure_v3.py"
  if cmp -s "$HSBG_WORK/toc_v3.json" "$HSBG_WORK/toc_v3_prev.json"; then echo "   → ổn định"; break; fi
done
pdfinfo "$HSBG_WORK"/pdf/*.pdf | grep Pages
