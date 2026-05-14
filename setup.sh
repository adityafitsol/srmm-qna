#!/bin/bash
set -e

echo "================================================"
echo "  SRMM Scraper — Setup"
echo "================================================"

echo "[1/2] Installing Python dependencies..."
pip3 install -r requirements.txt

echo ""
echo "[2/2] Checking required files..."
MISSING=0
for f in "quiet-mechanic-451307-s9-1bd5db312124.json" "companies.json" "companies_list.txt"; do
    if [ ! -f "$f" ]; then
        echo "  MISSING: $f"
        MISSING=1
    else
        echo "  OK: $f"
    fi
done

echo ""
if [ $MISSING -eq 1 ]; then
    echo "Upload the missing files above, then run: bash run.sh"
else
    echo "All good! Run: bash run.sh"
fi
