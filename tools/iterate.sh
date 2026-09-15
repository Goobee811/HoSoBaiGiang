#!/usr/bin/env bash
# Dựng giáo trình Word và lặp cho tới khi số trang trong mục lục ổn định.
#   HSBG_WORK  thư mục làm việc (mặc định: ../build so với tools/)
set -e
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
REPO="$(cd "$HERE/.." && pwd)"
export HSBG_WORK="${HSBG_WORK:-$REPO/build}"
DOCX="$REPO/Ho_so_bai_giang_lai_xe_2026_v1.0.docx"

for i in 1 2 3; do
  echo "── Lượt $i ──"
  python3 "$HERE/build.py" >/dev/null
  rm -rf "$HSBG_WORK/pdf" && mkdir -p "$HSBG_WORK/pdf"
  soffice --headless --norestore --convert-to pdf --outdir "$HSBG_WORK/pdf" "$DOCX" >/dev/null 2>&1
  cp -f "$HSBG_WORK/toc.json" "$HSBG_WORK/toc_prev.json" 2>/dev/null || true
  python3 "$HERE/measure.py"
  if cmp -s "$HSBG_WORK/toc.json" "$HSBG_WORK/toc_prev.json"; then
    echo "   → số trang ĐÃ ỔN ĐỊNH"; break
  fi
done
pdfinfo "$HSBG_WORK"/pdf/*.pdf | grep Pages
