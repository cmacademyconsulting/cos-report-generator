# COS™ Report Generator

**Ethics-First Construction Management Reports**  
*CM Academy | Susil Bhandari, CCM*  
*DOI: 10.5281/zenodo.18802971*

---

## Overview

The COS™ Report Generator produces fully-branded, professional-grade construction and infrastructure governance reports using the **COS™ Methodology — Compliance · Oversight · Sustainability**, developed by CM Academy (Nepal).

Each report is automatically generated as a `.docx` (Word) and `.pdf` file with:
- CM Academy letterhead and COS™ branding
- Colour-coded compliance scoring across all three pillars
- Auto-inserted DOI citation (`10.5281/zenodo.18802971`)
- CCM credential footer and LinkedIn reference
- Dual sign-off block for auditor and client

---

## Report Suite (6 Types)

| # | Report | Target | Price Point |
|---|--------|--------|-------------|
| 1 | **QA/QC Site Audit** | Contractors, PMCs | $500–$1,500 |
| 2 | **OSH Field Audit** | Contractors, HSE teams | $500–$1,500 |
| 3 | **NCR Register & CAPA** | QA/QC Engineers, PMCs | $500–$1,000 |
| 4 | **Project Governance Assessment** | Government, Donors | $1,500–$3,000 |
| 5 | **ESG Alignment Report** | Donor projects, Banks | $1,000–$2,500 |
| 6 | **Donor Readiness Report** | ADB/WB Executing Agencies | $1,500–$2,500 |

> Full engagement (all 6 reports) = **$5,500–$12,000 per project**

---

## Quick Start

### Prerequisites
```bash
node --version    # v18+ required
npm install -g docx
pip install python-docx  # for PDF conversion via LibreOffice
```

### Generate a Single Report
```bash
# QA/QC Site Audit
node generate_report.js

# OSH Field Audit
node generate_osh_report.js

# NCR Register
node generate_ncr_register.js

# Project Governance Assessment
node generate_governance_report.js

# ESG Alignment Report
node generate_esg_report.js

# Donor Readiness Report
node generate_donor_report.js
```

### Generate All 6 Reports
```bash
chmod +x run_all.sh
./run_all.sh
```

All outputs are saved to the `output/` directory as `.docx` and `.pdf`.

---

## Customising for a Client

Each generator has a **report data block** (`const R = { ... }`) near the top of the file. Edit only this block:

```javascript
const R = {
  project_name:    "Your Client's Project Name",
  client:          "Client Organization",
  contractor:      "Contractor Name",
  location:        "Project Location",
  audit_date:      "DD Month YYYY",
  audit_ref:       "COS-QA-YYYY-XXX-001",
  // COS™ Scores (0–100)
  c_score: 82,   // Compliance
  o_score: 74,   // Oversight
  s_score: 68,   // Sustainability
  // Checklist items — update status: "PASS" | "FAIL" | "OBS" | "N/A"
  ...
};
```

Then run the generator. A branded report is produced in under 5 seconds.

---

## Repository Structure

```
cos-report-generator/
├── README.md                       # This file
├── run_all.sh                      # Generate all 6 reports
├── cos_helpers.js                  # Shared brand + layout helpers
├── generate_report.js              # Report 1: QA/QC Site Audit
├── generate_osh_report.js          # Report 2: OSH Field Audit
├── generate_ncr_register.js        # Report 3: NCR Register & CAPA
├── generate_governance_report.js   # Report 4: Project Governance Assessment
├── generate_esg_report.js          # Report 5: ESG Alignment Report
├── generate_donor_report.js        # Report 6: Donor Readiness Report
├── web/
│   └── index.html                  # Client intake web form (open in browser)
├── docs/
│   └── COS_METHODOLOGY_WHITE_PAPER.pdf
├── scripts/
│   └── pdf_convert.sh              # LibreOffice PDF batch converter
└── output/                         # All generated reports land here
```

---

## COS™ Methodology

**Compliance · Oversight · Sustainability**

| Pillar | Purpose | Standards Integrated |
|--------|---------|---------------------|
| **C — Compliance** | Legality, transparency, audit-readiness | CMAA · FIDIC · ISO 9001 · IFC · GCF · Local Laws |
| **O — Oversight** | Ethical supervision, accountability, resilience | CMAA Duty of Care · Audit Frameworks · Donor Protocols |
| **S — Sustainability** | Climate alignment, ESG, SDG integration | UN SDGs · GRI · SASB · TCFD · Carbon Systems |

> *"COS™ operationalizes governance in real-time project delivery. It is not a parallel system but a unifying methodology."*

---

## About

**Susil Bhandari, CCM**  
Certified Construction Manager (CMAA)  
Founder, CM Academy | Director, NeoPlan Consult Pvt. Ltd.  
22 years experience across Nepal, Bahrain, and Qatar  
LinkedIn: [linkedin.com/in/ccm-susil-bhandari](https://linkedin.com/in/ccm-susil-bhandari)  
Email: cm.academy.consulting@gmail.com

**CM Academy** — *Developed in Nepal. Leading the World.™*  
© 2026 Complete Construction Management Developers Pvt. Ltd. (Reg. No. 275143/078/079)  
Licensed CC BY 4.0 | DOI: [10.5281/zenodo.18802971](https://doi.org/10.5281/zenodo.18802971)
