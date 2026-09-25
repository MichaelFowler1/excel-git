#!/usr/bin/env bash
# Download real-world workbooks for fuzz/merge_fuzz.py into DIR (default
# ./corpus). Needs git, curl and python with py7zr (pip install py7zr).
#
#   fuzz/fetch_corpus.sh [DIR] [--no-enron]
#
# - Enron: 15,871 .xlsx files from the Enron spreadsheet corpus (Hermans &
#   Murphy-Hill, MSR 2015), figshare 1221767. About 1 GB download, 1.8 GB
#   unpacked.
# - Test suites: the .xlsx/.xlsm test files of Apache POI (many of them
#   attached to bug reports by users), calamine, readxl, ClosedXML, exceljs,
#   Apache Tika and pandas. About 1,000 files.
set -euo pipefail
DIR=${1:-corpus}
[[ "${1:-}" == --* ]] && DIR=corpus
mkdir -p "$DIR/files" "$DIR/src"
cd "$DIR"

fetch() {  # fetch OWNER/REPO [PATH_IN_REPO]
  local name=${1#*/}
  [ -d "src/$name" ] || git clone -q --depth 1 --filter=blob:none --no-checkout "https://github.com/$1" "src/$name"
  (cd "src/$name" && git ls-tree -r --name-only HEAD -- "${2:-.}" | grep -iE '\.xls[xm]$' > ../$name.list || true
   git sparse-checkout init --no-cone && sed 's|^|/|' ../$name.list > .git/info/sparse-checkout && git checkout -q HEAD)
}
fetch apache/poi test-data/spreadsheet
for r in tafia/calamine tidyverse/readxl ClosedXML/ClosedXML exceljs/exceljs apache/tika pandas-dev/pandas; do
  fetch "$r"
done
python3 - <<'EOF'
import hashlib, os, shutil
seen = set()
for src in sorted(os.listdir("src")):
    for dp, _, fs in os.walk(os.path.join("src", src)):
        if "/.git" in dp:
            continue
        for f in fs:
            if f.lower().endswith((".xlsx", ".xlsm")):
                p = os.path.join(dp, f)
                h = hashlib.sha1(open(p, "rb").read()).hexdigest()
                if h not in seen:
                    seen.add(h)
                    shutil.copy(p, os.path.join("files", f"{src}__{h[:6]}__{f}"))
print(f"{len(seen)} test-suite workbooks in {os.path.abspath('files')}")
EOF

if [[ " $* " != *" --no-enron "* ]]; then
  [ -f enron.7z ] || [ -d enron ] || curl -sSL -o enron.7z https://ndownloader.figshare.com/files/3242531
  if [ -f enron.7z ]; then
    python3 -c "import py7zr; py7zr.SevenZipFile('enron.7z').extractall('enron')" && rm enron.7z
  fi
  echo "$(find enron -name '*.xlsx' | wc -l) Enron workbooks in $(pwd)/enron"
fi
