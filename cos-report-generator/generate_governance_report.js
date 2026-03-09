/**
 * COS™ Project Governance Assessment Report
 * CM Academy | Susil Bhandari, CCM
 * DOI: 10.5281/zenodo.18802971
 *
 * Comprehensive governance health check across all 3 COS™ pillars
 * Targeted at: PMCs, donor-funded projects, government infrastructure
 * Price point: $1,500 – $3,000 per report
 */

'use strict';
const {
  Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign
} = require('docx');
const fs = require('fs');
const H = require('./cos_helpers');

// ─── REPORT DATA ──────────────────────────────────────────────────────────────
const R = {
  project_name:    "Melamchi Water Supply Distribution Network Expansion",
  project_code:    "MWSDB-2026-GOV-003",
  client:          "Melamchi Water Supply Development Board (MWSDB)",
  contractor:      "Joint Venture: Kalika–Srinagar Construction",
  pmc:             "CM Academy | NeoPlan Consult Pvt. Ltd.",
  funder:          "Asian Development Bank (ADB Loan 2776-NEP) / GoN",
  location:        "Kathmandu Valley Distribution Zones 1–4",
  contract_value:  "NPR 3.84 Billion (USD ~28.9 Million)",
  contract_no:     "MWSDB/NCB/2026/003",
  assessment_date: "09 March 2026",
  report_ref:      "COS-GOV-2026-003-001",
  assessor:        "Susil Bhandari, CCM",
  report_version:  "v1.0 — Issued for Client Review",
  project_phase:   "Construction Phase — 34% Complete",

  // Governance domain scores (0–100)
  scores: {
    contract_admin:    78,
    financial_controls: 71,
    risk_management:   65,
    stakeholder_mgmt:  80,
    document_control:  74,
    change_management: 62,
    quality_systems:   77,
    safety_governance: 69,
    sustainability_gov: 58,
    donor_reporting:   83,
  },

  // COS™ Pillar aggregates
  c_score: 76,
  o_score: 70,
  s_score: 61,

  exec_summary: "This COS™ Project Governance Assessment evaluates the governance maturity of the Melamchi Water Supply Distribution Network Expansion project at 34% construction completion. The assessment covers ten governance domains mapped to the COS™ Compliance, Oversight, and Sustainability pillars. The overall governance health stands at 69% — PARTIAL. Strengths are evident in donor reporting alignment, stakeholder management, and contract administration. Immediate attention is required in three areas: (1) Sustainability governance — carbon tracking and SDG reporting are absent; (2) Change management — variation order backlog of 14 unprocessed VOs is creating financial and schedule risk; and (3) Risk management — the risk register has not been updated in 47 days against a required 14-day cycle.",

  // Governance domains
  domains: [
    {
      id: "GOV-01", name: "Contract Administration",
      pillar: "C", score: 78,
      description: "Review of contract management, FIDIC compliance, correspondence registers, and contractual obligation tracking.",
      findings: [
        { item: "Contract documents — signed originals available and version-controlled", status: "PASS", ref: "FIDIC 1.5" },
        { item: "Correspondence register maintained — all letters logged within 24 hours", status: "PASS", ref: "FIDIC" },
        { item: "Site instruction register complete — all verbal instructions confirmed in writing", status: "OBS", ref: "FIDIC 3.4" },
        { item: "Contractor's programme (baseline + current revision) submitted and approved", status: "PASS", ref: "FIDIC 8.3" },
        { item: "Extension of Time (EoT) claims processed within contractual timeline", status: "PASS", ref: "FIDIC 20.1" },
        { item: "Variation Order (VO) register — all VOs logged, priced, and approved", status: "FAIL", ref: "FIDIC 13.3" },
        { item: "Payment certificates issued within contractual period", status: "PASS", ref: "FIDIC 14" },
        { item: "Defects notification register maintained and tracked", status: "OBS", ref: "FIDIC 11" },
      ]
    },
    {
      id: "GOV-02", name: "Financial Controls & Audit Readiness",
      pillar: "C", score: 71,
      description: "Assessment of financial management, cost control, audit trail completeness, and donor financial reporting.",
      findings: [
        { item: "Approved budget and cost-to-complete forecast current and reconciled", status: "PASS", ref: "ADB FM" },
        { item: "Earned Value Management (EVM) — SPI and CPI calculated monthly", status: "OBS", ref: "PMI EVM" },
        { item: "Procurement records — all contracts competitively tendered and documented", status: "PASS", ref: "ADB Proc." },
        { item: "Bank reconciliation performed and filed monthly", status: "PASS", ref: "ADB FM" },
        { item: "Contingency drawdown — properly authorised and recorded", status: "PASS", ref: "ADB FM" },
        { item: "Audit findings (last internal audit) — all findings closed", status: "FAIL", ref: "ADB Audit" },
        { item: "Fixed asset register maintained for all donor-procured equipment", status: "OBS", ref: "ADB FM" },
      ]
    },
    {
      id: "GOV-03", name: "Risk Management",
      pillar: "O", score: 65,
      description: "Evaluation of risk register currency, risk treatment, and residual risk monitoring.",
      findings: [
        { item: "Risk register established — all identified risks logged with owners", status: "PASS", ref: "ISO 31000" },
        { item: "Risk register reviewed and updated on 14-day cycle", status: "FAIL", ref: "ISO 31000" },
        { item: "Top 5 risks have active treatment plans with measurable KPIs", status: "OBS", ref: "ISO 31000" },
        { item: "Risk review conducted with PMC and client at monthly project meetings", status: "PASS", ref: "CMAA" },
        { item: "Force majeure and climate risk formally assessed and documented", status: "FAIL", ref: "GCF" },
        { item: "Insurance coverage current — all required policies in force", status: "PASS", ref: "FIDIC 19" },
      ]
    },
    {
      id: "GOV-04", name: "Stakeholder Management",
      pillar: "O", score: 80,
      description: "Review of community engagement, public communication, grievance mechanisms, and multi-agency coordination.",
      findings: [
        { item: "Stakeholder mapping completed — all affected parties identified", status: "PASS", ref: "IFC PS1" },
        { item: "Community grievance mechanism established and publicised", status: "PASS", ref: "IFC PS1" },
        { item: "Grievance register maintained — all complaints logged and responded to within 14 days", status: "PASS", ref: "IFC PS1" },
        { item: "Public disclosure — project information board at all active sites", status: "PASS", ref: "ADB SPS" },
        { item: "Multi-agency coordination meetings — minutes available and action items tracked", status: "PASS", ref: "CMAA" },
        { item: "Vulnerable/affected household resettlement — monitoring reports current", status: "OBS", ref: "ADB SPS" },
      ]
    },
    {
      id: "GOV-05", name: "Document Control",
      pillar: "C", score: 74,
      description: "Assessment of document management systems, drawing control, and records accessibility.",
      findings: [
        { item: "Document Management System (DMS) in use — all documents version-controlled", status: "PASS", ref: "ISO 9001" },
        { item: "Drawing register current — latest revisions issued to all relevant parties", status: "PASS", ref: "FIDIC" },
        { item: "Transmittal records — all document transfers logged with receipt confirmation", status: "OBS", ref: "ISO 9001" },
        { item: "RFI register maintained — all RFIs responded to within contractual period", status: "PASS", ref: "CMAA" },
        { item: "Submittals log current — all contractor submittals tracked with approval status", status: "OBS", ref: "CMAA" },
      ]
    },
    {
      id: "GOV-06", name: "Change Management",
      pillar: "C", score: 62,
      description: "Evaluation of change control procedures, variation order processing, and scope change governance.",
      findings: [
        { item: "Change control procedure established and communicated to all parties", status: "PASS", ref: "FIDIC 13" },
        { item: "All scope changes formally submitted as VOs before work proceeds", status: "FAIL", ref: "FIDIC 13.1" },
        { item: "VO backlog — all VOs assessed, priced, and approved within 28 days", status: "FAIL", ref: "FIDIC 13.3" },
        { item: "Variation log reconciled monthly against contract sum adjustment", status: "OBS", ref: "FIDIC" },
        { item: "Scope creep risk assessed — budget impact of all pending VOs quantified", status: "FAIL", ref: "CMAA" },
      ]
    },
    {
      id: "GOV-07", name: "Quality Management Systems",
      pillar: "C", score: 77,
      description: "Review of ISO 9001-aligned QMS implementation, ITP compliance, and audit closure.",
      findings: [
        { item: "Quality Management Plan (QMP) approved and version-controlled", status: "PASS", ref: "ISO 9001" },
        { item: "ITP (Inspection and Test Plans) in use for all major construction activities", status: "PASS", ref: "ISO 9001" },
        { item: "Internal quality audits conducted quarterly — reports available", status: "PASS", ref: "ISO 9001" },
        { item: "Calibration records for all testing equipment current", status: "OBS", ref: "ISO 9001" },
        { item: "NCR register maintained — closure rate ≥ 80% at any time", status: "PASS", ref: "ISO 9001" },
        { item: "Lessons learned register established and updated", status: "OBS", ref: "CMAA" },
      ]
    },
    {
      id: "GOV-08", name: "Safety Governance",
      pillar: "O", score: 69,
      description: "OSH governance systems, incident accountability, and duty of care oversight.",
      findings: [
        { item: "Safety Management Plan approved and available on all sites", status: "PASS", ref: "ISO 45001" },
        { item: "Monthly HSE performance report submitted to client/donor", status: "PASS", ref: "ADB SPS" },
        { item: "Lost Time Injury (LTI) frequency rate tracked and reported", status: "PASS", ref: "ILO" },
        { item: "All fatalities/serious injuries reported to donor within 24 hours", status: "N/A", ref: "ADB SPS" },
        { item: "Safety improvement plan in place following last audit findings", status: "OBS", ref: "ISO 45001" },
        { item: "Duty of care obligations formally allocated to named individuals", status: "FAIL", ref: "CMAA" },
      ]
    },
    {
      id: "GOV-09", name: "Sustainability Governance",
      pillar: "S", score: 58,
      description: "Climate resilience integration, ESG metrics, SDG alignment, and carbon management.",
      findings: [
        { item: "Environmental Management Plan (EMP) approved and in use", status: "PASS", ref: "IFC PS3" },
        { item: "Carbon / GHG emission monitoring system established", status: "FAIL", ref: "TCFD" },
        { item: "SDG contribution documented — project mapped to relevant SDGs", status: "FAIL", ref: "UN SDG" },
        { item: "ESG performance metrics tracked and reported to donor quarterly", status: "OBS", ref: "GRI" },
        { item: "Climate resilience measures in design verified against climate projections", status: "OBS", ref: "GCF" },
        { item: "Biodiversity and cultural heritage — no-go zones respected and monitored", status: "PASS", ref: "IFC PS8" },
        { item: "TCFD-aligned disclosure included in progress reports", status: "FAIL", ref: "TCFD" },
      ]
    },
    {
      id: "GOV-10", name: "Donor Reporting & Compliance",
      pillar: "S", score: 83,
      description: "ADB reporting obligations, disbursement conditions, safeguard compliance, and audit alignment.",
      findings: [
        { item: "Quarterly progress reports submitted on schedule to ADB", status: "PASS", ref: "ADB Loan" },
        { item: "Safeguard monitoring reports (social and environmental) filed bi-annually", status: "PASS", ref: "ADB SPS" },
        { item: "Financial management reports submitted within 45 days of each quarter", status: "PASS", ref: "ADB FM" },
        { item: "Procurement plan current and accessible on ADB portal", status: "PASS", ref: "ADB Proc." },
        { item: "Anti-corruption, fraud, and integrity officer designated", status: "PASS", ref: "ADB ICA" },
        { item: "Disbursement conditions met — no blocked tranches", status: "PASS", ref: "ADB Loan" },
        { item: "Gender action plan implementation progress reported", status: "OBS", ref: "ADB GAP" },
      ]
    },
  ],

  // Priority actions
  priority_actions: [
    { priority: "CRITICAL", domain: "Change Management", action: "Clear VO backlog (14 unprocessed VOs) within 21 days. Assign dedicated QS resource. Escalate to client if contractor non-responsive.", deadline: "30 March 2026" },
    { priority: "CRITICAL", domain: "Risk Management", action: "Update risk register immediately. Climate and force majeure risk assessment to be completed and submitted to ADB within 14 days.", deadline: "23 March 2026" },
    { priority: "HIGH", domain: "Sustainability Governance", action: "Establish GHG/carbon monitoring log. Map project to SDGs 6, 9, 11, 13. Include TCFD disclosure in next quarterly report.", deadline: "09 April 2026" },
    { priority: "HIGH", domain: "Safety Governance", action: "Formally assign duty of care responsibilities in writing to PM, RE, and HSE Officer. Submit to consultant for record.", deadline: "16 March 2026" },
    { priority: "MEDIUM", domain: "Financial Controls", action: "Close all outstanding internal audit findings. Submit closure evidence with next quarterly financial report.", deadline: "30 March 2026" },
    { priority: "MEDIUM", domain: "Contract Administration", action: "Implement EVM. Calculate and report CPI/SPI in next progress report. VO register to be reconciled.", deadline: "09 April 2026" },
  ],
};

// ─── BUILD ─────────────────────────────────────────────────────────────────────
function build(r) {
  const children = [];

  // Cover
  children.push(...H.makeCover(
    r.report_type || "PROJECT GOVERNANCE ASSESSMENT REPORT",
    "COS™ Governance Health Check — Compliance · Oversight · Sustainability",
    H.C.NAVY,
    [
      ["Project",         r.project_name],
      ["Client / Owner",  r.client],
      ["Funder",          r.funder],
      ["Contract Value",  r.contract_value],
      ["Project Phase",   r.project_phase],
      ["PMC",             r.pmc],
      ["Assessment Ref",  r.report_ref],
      ["Assessment Date", r.assessment_date],
      ["Lead Assessor",   r.assessor + ", CCM"],
      ["Version",         r.report_version],
    ]
  ));

  // Section 1 — Overview
  children.push(
    H.h1("1. PROJECT & ASSESSMENT OVERVIEW"),
    H.divider(),
    H.spacer(80),
    overviewGrid(r),
    H.spacer(160),
  );

  // Section 2 — Executive Summary + Radar
  children.push(
    H.h1("2. EXECUTIVE SUMMARY"),
    H.divider(),
    H.spacer(80),
    H.body(r.exec_summary, { size: 21 }),
    H.spacer(120),
    H.h2("2.1 COS™ Tri-Pillar Governance Score"),
    H.spacer(60),
    H.scoreSummary(r.c_score, r.o_score, r.s_score),
    H.spacer(120),
    H.h2("2.2 Governance Domain Scorecard"),
    H.spacer(60),
    domainScorecard(r.domains),
    H.spacer(160),
  );

  // Section 3 — Priority Actions
  children.push(
    H.h1("3. PRIORITY ACTION PLAN"),
    H.divider(H.C.RED),
    H.spacer(80),
    priorityTable(r.priority_actions),
    H.spacer(160),
  );

  // Section 4 — Domain Assessments
  children.push(
    H.h1("4. GOVERNANCE DOMAIN ASSESSMENTS"),
    H.divider(),
    H.spacer(80),
  );

  r.domains.forEach(domain => {
    children.push(
      H.h2(`${domain.id} — ${domain.name}  [Score: ${domain.score}%  |  Pillar: ${domain.pillar}]`),
      H.body(domain.description, { italic: true }),
      H.spacer(60),
      H.checklistTable(domain.findings),
      H.spacer(120),
    );
  });

  // Section 5 — Governance Maturity Summary
  children.push(
    H.h1("5. GOVERNANCE MATURITY SUMMARY"),
    H.divider(),
    H.spacer(80),
    maturityTable(r.domains),
    H.spacer(120),
    H.body("Governance maturity levels: Level 1 (Ad Hoc, <50%) → Level 2 (Developing, 50–64%) → Level 3 (Defined, 65–79%) → Level 4 (Managed, 80–89%) → Level 5 (Optimised, 90%+)", { italic: true, size: 17 }),
    H.spacer(160),
  );

  // Section 6 — COS™ Integration Statement
  children.push(
    H.h1("6. COS™ METHODOLOGY — GOVERNANCE INTEGRATION"),
    H.divider(H.C.GOLD),
    H.spacer(80),
    cosIntegrationTable(),
    H.spacer(120),
    H.body("This assessment was conducted using the COS™ Methodology (Compliance · Oversight · Sustainability), an ethics-first governance framework developed by CM Academy (Nepal). The COS™ Methodology integrates CMAA, FIDIC, ISO, IFC/GCF, UN SDGs, GRI, SASB, and TCFD into a single verifiable governance ecosystem, validated through projects in Nepal, Bahrain, and Qatar over 22 years of practice.", { italic: true }),
    H.spacer(160),
  );

  // Section 7 — Sign-Off
  children.push(
    H.h1("7. CERTIFICATION & SIGN-OFF"),
    H.divider(H.C.GOLD),
    H.spacer(80),
    ...H.signOff(r.assessment_date, "Project Director / Client Representative"),
  );

  return H.makeDoc(r.report_ref, r.assessment_date, children);
}

// ─── SUB-COMPONENTS ───────────────────────────────────────────────────────────
function overviewGrid(r) {
  const pairs = [
    ["Project Name", r.project_name], ["Project Code", r.project_code],
    ["Client / Owner", r.client],     ["Funder", r.funder],
    ["PMC / Consultant", r.pmc],      ["Contractor", r.contractor],
    ["Contract Value", r.contract_value], ["Contract No.", r.contract_no],
    ["Project Phase", r.project_phase],   ["Assessment Date", r.assessment_date],
    ["Lead Assessor", r.assessor + ", CCM"], ["Report Version", r.report_version],
  ];
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2000, 2680, 2000, 2680],
    rows: pairs.reduce((rows, _, i, a) => {
      if (i % 2 === 0) rows.push(new TableRow({
        children: a.slice(i, i+2).flatMap(([l,v]) => [
          H.cell(l, 2000, H.C.NAVY, { bold: true, size: 18, color: H.C.WHITE }),
          H.cell(v, 2680, i % 4 === 0 ? H.C.LIGHT : H.C.WHITE, { size: 18 }),
        ])
      }));
      return rows;
    }, [])
  });
}

function domainScorecard(domains) {
  const cols = [800, 3200, 1000, 1200, 1560, 1600];
  const hdr = new TableRow({ tableHeader: true, children:
    ["ID", "Governance Domain", "Pillar", "Score", "Status", "COS™ Standards Anchored"].map((h, i) =>
      H.cell(h, cols[i], H.C.NAVY, { bold: true, size: 17, color: H.C.WHITE, center: true })
    )
  });
  const pillarColor = p => p === "C" ? H.C.GREEN : p === "O" ? H.C.NAVY : H.C.GOLD;
  const rows = domains.map((d, idx) => {
    const bg = idx % 2 === 0 ? H.C.WHITE : H.C.LIGHT;
    const sc = d.score;
    const statusBg = sc >= 80 ? H.C.PASS : sc >= 65 ? H.C.OBS : H.C.FAIL;
    const statusLabel = sc >= 80 ? "STRONG" : sc >= 65 ? "ADEQUATE" : "WEAK";
    const statusColor = sc >= 80 ? H.C.GREEN : sc >= 65 ? H.C.AMBER : H.C.RED;
    const standards = { C: "CMAA · FIDIC · ISO · IFC", O: "CMAA · ISO 31000 · ADB", S: "SDGs · GRI · TCFD · GCF" }[d.pillar];
    return new TableRow({ children: [
      H.cell(d.id, cols[0], bg, { size: 17, bold: true }),
      H.cell(d.name, cols[1], bg, { size: 17 }),
      new TableCell({ borders: H.BORDERS, shading: { fill: bg, type: ShadingType.CLEAR }, width: { size: cols[2], type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 80, right: 80 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: d.pillar, bold: true, size: 18, color: pillarColor(d.pillar), font: "Arial" })] })]
      }),
      new TableCell({ borders: H.BORDERS, shading: { fill: statusBg, type: ShadingType.CLEAR }, width: { size: cols[3], type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 80, right: 80 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `${sc}%`, bold: true, size: 19, color: H.C.NAVY, font: "Arial" })] })]
      }),
      new TableCell({ borders: H.BORDERS, shading: { fill: statusBg, type: ShadingType.CLEAR }, width: { size: cols[4], type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 80, right: 80 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: statusLabel, bold: true, size: 17, color: statusColor, font: "Arial" })] })]
      }),
      H.cell(standards, cols[5], bg, { size: 16, italic: true }),
    ]});
  });
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: cols, rows: [hdr, ...rows] });
}

function priorityTable(actions) {
  const cols = [1000, 2000, 4360, 2000];
  const hdr = new TableRow({ tableHeader: true, children:
    ["Priority", "Governance Domain", "Required Action", "Deadline"].map((h, i) =>
      H.cell(h, cols[i], H.C.NAVY, { bold: true, size: 17, color: H.C.WHITE, center: true })
    )
  });
  const pColor = p => p === "CRITICAL" ? H.C.RED : p === "HIGH" ? H.C.ORANGE : "6B4E00";
  const pBg    = p => p === "CRITICAL" ? "FFE8E8" : p === "HIGH" ? "FFF0D6" : H.C.OBS;
  const rows = actions.map((a, idx) => {
    const bg = idx % 2 === 0 ? H.C.WHITE : H.C.LIGHT;
    return new TableRow({ children: [
      new TableCell({ borders: H.BORDERS, shading: { fill: pBg(a.priority), type: ShadingType.CLEAR }, width: { size: cols[0], type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 80, right: 80 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: a.priority, bold: true, size: 16, color: pColor(a.priority), font: "Arial" })] })]
      }),
      H.cell(a.domain, cols[1], bg, { size: 17, bold: true }),
      H.cell(a.action, cols[2], bg, { size: 17 }),
      H.cell(a.deadline, cols[3], bg, { size: 17, bold: true }),
    ]});
  });
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: cols, rows: [hdr, ...rows] });
}

function maturityTable(domains) {
  const cols = [800, 3000, 1000, 1200, 3360];
  const hdr = new TableRow({ tableHeader: true, children:
    ["ID", "Domain", "Score", "Maturity Level", "Recommendation"].map((h,i) =>
      H.cell(h, cols[i], H.C.NAVY, { bold: true, size: 17, color: H.C.WHITE, center: true })
    )
  });
  const maturity = s => s >= 90 ? { level: "L5 — Optimised", rec: "Maintain and share best practices across project." } :
    s >= 80 ? { level: "L4 — Managed",    rec: "Document processes and replicate across contract packages." } :
    s >= 65 ? { level: "L3 — Defined",    rec: "Strengthen weak items; implement continuous improvement cycle." } :
    s >= 50 ? { level: "L2 — Developing", rec: "Immediate process improvement required; assign responsible owner." } :
    { level: "L1 — Ad Hoc", rec: "Critical intervention required; escalate to senior management." };

  const rows = domains.map((d, idx) => {
    const bg = idx % 2 === 0 ? H.C.WHITE : H.C.LIGHT;
    const m = maturity(d.score);
    const sc = d.score;
    const scoreBg = sc >= 80 ? H.C.PASS : sc >= 65 ? H.C.OBS : H.C.FAIL;
    return new TableRow({ children: [
      H.cell(d.id, cols[0], bg, { size: 16, bold: true }),
      H.cell(d.name, cols[1], bg, { size: 16 }),
      new TableCell({ borders: H.BORDERS, shading: { fill: scoreBg, type: ShadingType.CLEAR }, width: { size: cols[2], type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 80, right: 80 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `${sc}%`, bold: true, size: 17, color: H.C.NAVY, font: "Arial" })] })]
      }),
      H.cell(m.level, cols[3], scoreBg, { size: 16, bold: true }),
      H.cell(m.rec, cols[4], bg, { size: 16, italic: true }),
    ]});
  });
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: cols, rows: [hdr, ...rows] });
}

function cosIntegrationTable() {
  const cols = [2000, 2453, 2453, 2454];
  const hdr = new TableRow({ tableHeader: true, children:
    ["COS™ Pillar", "C — Compliance", "O — Oversight", "S — Sustainability"].map((h,i) =>
      H.cell(h, cols[i], i === 0 ? H.C.NAVY : i === 1 ? H.C.GREEN : i === 2 ? H.C.NAVY : H.C.GOLD, { bold: true, size: 18, color: H.C.WHITE, center: i > 0 })
    )
  });
  const rowData = [
    ["Governance Focus", "Legality, transparency, audit-readiness", "Ethical supervision, accountability, resilience", "Climate alignment, ESG, SDG integration"],
    ["Standards Integrated", "CMAA · FIDIC · ISO 9001/31000 · IFC · GCF · Local Laws", "CMAA Duty of Care · Audit Frameworks · Donor Oversight Protocols", "UN SDGs · GRI · SASB · TCFD · Carbon Credit Systems"],
    ["Governance Domains", "Contract Admin · Finance · Doc Control · Change Mgmt · QMS", "Risk Mgmt · Stakeholder Mgmt · Safety Governance", "Sustainability Gov · Donor Reporting"],
    ["Project Score", `${R.c_score}% — ${R.c_score >= 80 ? 'COMPLIANT' : R.c_score >= 65 ? 'PARTIAL' : 'NON-COMPLIANT'}`, `${R.o_score}% — ${R.o_score >= 80 ? 'COMPLIANT' : R.o_score >= 65 ? 'PARTIAL' : 'NON-COMPLIANT'}`, `${R.s_score}% — ${R.s_score >= 80 ? 'COMPLIANT' : R.s_score >= 65 ? 'PARTIAL' : 'NON-COMPLIANT'}`],
  ];
  const rows = rowData.map((row, idx) => new TableRow({ children:
    row.map((val, i) => H.cell(val, cols[i], i === 0 ? H.C.NAVY : idx % 2 === 0 ? H.C.LIGHT : H.C.WHITE, { bold: i === 0, size: 17, color: i === 0 ? H.C.WHITE : H.C.GREY }))
  }));
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: cols, rows: [hdr, ...rows] });
}

// ─── RUN ──────────────────────────────────────────────────────────────────────
R.report_type = "PROJECT GOVERNANCE ASSESSMENT REPORT";
const doc = build(R);
H.save(doc, `output/COS_Governance_Assessment_${R.report_ref}.docx`)
  .catch(err => { console.error(err); process.exit(1); });
