#!/bin/bash
# ─────────────────────────────────────────────────────────────────────────────
# COS™ Report Generator — Run All
# CM Academy | Susil Bhandari, CCM
# DOI: 10.5281/zenodo.18802971
# ─────────────────────────────────────────────────────────────────────────────

set -e

NAVY='\033[0;34m'
GOLD='\033[0;33m'
GREEN='\033[0;32m'
RED='\033[0;31m'
GREY='\033[0;37m'
BOLD='\033[1m'
NC='\033[0m'

echo ""
echo -e "${BOLD}${NAVY}╔══════════════════════════════════════════════════════════════╗${NC}"
echo -e "${BOLD}${NAVY}║         CM ACADEMY — COS™ REPORT GENERATOR                  ║${NC}"
echo -e "${BOLD}${GOLD}║         Compliance · Oversight · Sustainability              ║${NC}"
echo -e "${BOLD}${GREY}║         DOI: 10.5281/zenodo.18802971                         ║${NC}"
echo -e "${BOLD}${NAVY}╚══════════════════════════════════════════════════════════════╝${NC}"
echo ""

# Create output directory
mkdir -p output

# Check Node.js
if ! command -v node &> /dev/null; then
  echo -e "${RED}✗ Node.js not found. Install from https://nodejs.org${NC}"
  exit 1
fi

# Check docx module
if ! node -e "require('docx')" 2>/dev/null; then
  echo -e "${GOLD}⚠  Installing docx module...${NC}"
  npm install -g docx
fi

echo -e "${BOLD}Generating 6 COS™ Reports...${NC}"
echo ""

REPORTS=(
  "generate_report.js|QA/QC Site Audit Report"
  "generate_osh_report.js|OSH Field Audit Report"
  "generate_ncr_register.js|NCR Register & Closure Report"
  "generate_governance_report.js|Project Governance Assessment"
  "generate_esg_report.js|ESG Alignment Report"
  "generate_donor_report.js|Donor Readiness Report"
)

PASS=0
FAIL=0

for entry in "${REPORTS[@]}"; do
  script="${entry%%|*}"
  name="${entry##*|}"
  
  printf "  %-45s " "[$((PASS+FAIL+1))/6] $name..."
  
  if node "$script" 2>/dev/null; then
    echo -e "${GREEN}✓ DONE${NC}"
    PASS=$((PASS+1))
  else
    echo -e "${RED}✗ FAILED${NC}"
    FAIL=$((FAIL+1))
    node "$script" 2>&1 | tail -5
  fi
done

echo ""
echo -e "${BOLD}Converting to PDF...${NC}"
echo ""

PDF_PASS=0
PDF_FAIL=0

for docx in output/*.docx; do
  name=$(basename "$docx" .docx)
  printf "  %-50s " "$name..."
  
  # Try LibreOffice
  if command -v libreoffice &> /dev/null; then
    if libreoffice --headless --convert-to pdf "$docx" --outdir output/ 2>/dev/null; then
      echo -e "${GREEN}✓ PDF${NC}"
      PDF_PASS=$((PDF_PASS+1))
    else
      echo -e "${GOLD}⚠  SKIPPED${NC}"
      PDF_FAIL=$((PDF_FAIL+1))
    fi
  elif command -v soffice &> /dev/null; then
    if soffice --headless --convert-to pdf "$docx" --outdir output/ 2>/dev/null; then
      echo -e "${GREEN}✓ PDF${NC}"
      PDF_PASS=$((PDF_PASS+1))
    else
      echo -e "${GOLD}⚠  SKIPPED${NC}"
      PDF_FAIL=$((PDF_FAIL+1))
    fi
  else
    echo -e "${GOLD}⚠  LibreOffice not found — DOCX only${NC}"
    PDF_FAIL=$((PDF_FAIL+1))
    break
  fi
done

echo ""
echo -e "${BOLD}${NAVY}─────────────────────────────────────────────────────────────${NC}"
echo -e "${BOLD}Summary${NC}"
echo -e "  Reports generated:  ${GREEN}$PASS / $((PASS+FAIL))${NC}"
if [ $PDF_FAIL -eq 0 ]; then
  echo -e "  PDF conversions:    ${GREEN}$PDF_PASS / $((PDF_PASS+PDF_FAIL))${NC}"
else
  echo -e "  PDF conversions:    ${GOLD}$PDF_PASS / $((PDF_PASS+PDF_FAIL))  (install LibreOffice for PDF)${NC}"
fi
echo ""
echo -e "  Output directory:   ${BOLD}./output/${NC}"
echo ""
ls -lh output/ | awk 'NR>1 {printf "    %-6s  %s\n", $5, $9}'
echo ""
echo -e "${GOLD}COS™ Methodology — Ethics-First Construction Management${NC}"
echo -e "${GREY}Bhandari, S. (2026). Zenodo. DOI: 10.5281/zenodo.18802971${NC}"
echo -e "${GREY}© CM Academy | NeoPlan Consult Pvt. Ltd. | CC BY 4.0${NC}"
echo ""
