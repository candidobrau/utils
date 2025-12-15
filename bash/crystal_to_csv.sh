#!/usr/bin/env bash
# convert_xls_to_csv.sh
#
# Why this script exists:
# Benchmark provides system-generated Crystal Reports in .xls that are
# not standard .xls files. This script uses LibreOffice to convert them
# to .csv so they can be processed by other workflows
# What this script does:
# - Looks in a directory provided (defaults to the current directory)
# - Finds all .xls files in that directory (non-recursive)
# - Converts each .xls to .csv using LibreOffice (headless mode)
# - Writes the converted files into a new folder named:
#     converted_files_YYYY-MM-DD

# Some error proofing
# -e : exit on error
# -u : undefined variables are errors
# -o pipefail : pipeline fails if any command fails, not just the last one
set -euo pipefail

# Configuration: LibreOffice binary location on macOS (hardcoded)
SOFFICE="/Applications/LibreOffice.app/Contents/MacOS/soffice"

# Input directory: first argument, or current working directory as default
IN_DIR="${1:-$(pwd)}"

# Output directory named with today's date (local date)
TODAY="$(date +%F)"  # YYYY-MM-DD
OUT_DIR="${IN_DIR}/converted_files_${TODAY}"

# LibreOffice CSV filter options (keep as ONE argument)
# 44 = comma delimiter
# 34 = double quote as text qualifier
# 0  = UTF-8 encoding
# 1  = quote all text cells
CONVERT_TO='csv:Text - txt - csv (StarCalc):44,34,0,1'

echo "=== LibreOffice .xls → .csv Converter ==="
echo "Input directory : ${IN_DIR}"
echo "Output directory: ${OUT_DIR}"
echo

# Check to ensure $IN_DIR exists
if [[ ! -d "$IN_DIR" ]]; then
  echo "ERROR: Input directory does not exist: $IN_DIR" >&2
  exit 1
fi

# Check to ensure hard-coded path to LibreOffice is found
if [[ ! -x "$SOFFICE" ]]; then
  echo "ERROR: LibreOffice 'soffice' not found (or not executable) at:" >&2
  echo "  $SOFFICE" >&2
  echo "If LibreOffice is installed elsewhere, edit SOFFICE at the top of this script." >&2
  exit 1
fi

echo "OK: Found LibreOffice at: $SOFFICE"
echo

# Create output folder if needed
mkdir -p "$OUT_DIR"
echo "OK: Ensured output folder exists: $OUT_DIR"
echo

# Find .xls files (non-recursive) safely even if there are none
shopt -s nullglob
xls_files=( "$IN_DIR"/*.xls )
count="${#xls_files[@]}"

echo "Found $count .xls file(s) to convert."
echo

# Check if count of .xls files is 0
if [[ "$count" -eq 0 ]]; then
  echo "Nothing to do. Exiting."
  exit 0
fi

# Convert each file, initiating count variables first
success=0
failed=0
skipped=0

for f in "${xls_files[@]}"; do
  base="$(basename "$f")"
  out_csv="${OUT_DIR}/${base%.xls}.csv"

  echo "----------------------------------------"
  echo "FILE : $f"
  echo "OUT  : $out_csv"

  if [[ -f "$out_csv" ]]; then
    echo "SKIP : Output already exists, not overwriting."
    skipped=$((skipped + 1))
    continue
  fi

  # Build args as an array so the convert-to string stays intact (spaces/parentheses safe)
  args=(
    --headless
    --nologo
    --nolockcheck
    --norestore
    --convert-to "$CONVERT_TO"
    --outdir "$OUT_DIR"
    "$f"
  )

  # Print the exact command (copy/paste safe)
  printf "CMD  : %q " "$SOFFICE" "${args[@]}"
  printf "\n"

  # Run conversion
  if "$SOFFICE" "${args[@]}"; then
    # LibreOffice exits 0 does not always guarantee output file exists, so verify
    if [[ -f "$out_csv" ]]; then
      echo "OK   : Conversion succeeded."
      success=$((success + 1))
    else
      echo "WARN : LibreOffice reported success but output CSV not found."
      failed=$((failed + 1))
    fi
  else
    echo "ERROR: Conversion failed (LibreOffice returned non-zero)."
    failed=$((failed + 1))
  fi
done

echo "========================================"
echo "Done."
echo "Succeeded: $success"
echo "Skipped  : $skipped"
echo "Failed   : $failed"
echo "Output folder: $OUT_DIR"
echo "========================================"
