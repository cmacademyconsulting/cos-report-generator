/**
 * COS™ Donor Readiness Report
 * CM Academy | Susil Bhandari, CCM
 * DOI: 10.5281/zenodo.18802971
 *
 * Pre-submission governance assessment for donor-funded projects
 * Confirms audit-readiness before: loan disbursement, supervision missions,
 * mid-term reviews, project completion reports (PCR)
 * Targets: ADB, World Bank, UN, bilateral donor projects
 * Price point: $1,500 – $2,500 per report
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
  project_name:    "Integrated Urban Development Project — Pokhara Sub-Project",
  project_code:    "IUDP-2026-DONOR-POK",
  executing_agency:"Pokhara Metropolitan City (PMC-Pokhara)",
  implementing:    "Department of Urban Development & Building Construction (DUDBC)",
  donor:           "Asian Development Bank (ADB)",
  loan_no:         "ADB Loan 3580-NEP",
  grant_no:        "ADB Grant 0563-NEP",
  total_financing: "USD 150 Million (ADB) + USD 15 Million (GoN counterpart)",
  consultant:      "CM Academy | NeoPlan Consult Pvt. Ltd.",
  mission_type:    "ADB Mid-Term Review Mission",
  mission_date:    "16–20 March 2026",
  report_date:     "09 March 2026",
  report_ref:      "COS-DONOR-2026-POK-001",
  assessor:        "Susil Bhandari, CCM",
  report_version:  "v1.0 — Pre-Mission Readiness Assessment",
  readiness_level: "SUBSTANTIALLY READY",

  // Overall scores
  c_score: 81,
  o_score: 76,
  s_score: 73,

  exec_summary: "This COS™ Donor Readiness Report has been prepared in advance of the ADB Mid-Term Review Mission scheduled 16–20 March 2026 for the Integrated Urban Development Project — Pokhara Sub-Project. The report assesses governance readiness across seven critical domains required for a successful donor supervision mission. Overall readiness is rated SUBSTANTIALLY READY at 77% COS™ composite score. Key strengths: financial management systems are robust, procurement compliance is strong, and safeguard monitoring is up to date. Two areas require immediate attention before the mission: (1) the Gender Action Plan (GAP) progress report is 6 weeks overdue and must be filed before 12 March 2026; and (2) three environment safeguard monitoring observations from the last mission remain unresponded to.",

  // Readiness domains
  domains: [
    {
      id: "RD-01",
      name: "Financial Management & Disbursement",
      score: 84,
      mission_priority: "HIGH",
      description: "ADB will review SOE (Statement of Expenditure), financial statements, counterpart fund releases, and disbursement projections.",
      items: [
        { item: "Statement of Expenditure (SOE) — current quarter complete and reconciled", status: "PASS", ref: "ADB FM" },
        { item: "Audited project financial statements (FY2025) submitted to ADB on time", status: "PASS", ref: "ADB FM" },
        { item: "Counterpart fund (GoN) releases on schedule — no funding gap", status: "PASS", ref: "Loan Agmt" },
        { item: "Disbursement forecast (D-forecast) submitted and consistent with progress", status: "PASS", ref: "ADB FM" },
        { item: "Fixed asset register current — all ADB-procured assets recorded", status: "OBS", ref: "ADB FM" },
        { item: "No ineligible expenditures in last two quarters (ADB review findings)", status: "PASS", ref: "ADB FM" },
        { item: "Financial management improvement plan from last audit — actions closed", status: "PASS", ref: "ADB Audit" },
      ]
    },
    {
      id: "RD-02",
      name: "Procurement & Contract Management",
      score: 82,
      mission_priority: "HIGH",
      description: "Mission will review active contracts, procurement plan compliance, and contract performance.",
      items: [
        { item: "Procurement plan current and uploaded to ADB Online Procurement System", status: "PASS", ref: "ADB Proc." },
        { item: "All contracts awarded through ADB-approved procurement methods", status: "PASS", ref: "ADB Proc." },
        { item: "Contract performance — all contracts progressing within approved schedule", status: "OBS", ref: "FIDIC" },
        { item: "No misprocurement findings from prior review unresolved", status: "PASS", ref: "ADB Proc." },
        { item: "Variation orders processed — no unapproved cost overruns", status: "PASS", ref: "FIDIC 13" },
        { item: "Bid evaluation reports available and retained for all awarded contracts", status: "PASS", ref: "ADB Proc." },
        { item: "Integrity due diligence — no debarred contractors engaged", status: "PASS", ref: "ADB ICA" },
      ]
    },
    {
      id: "RD-03",
      name: "Progress & Physical Completion",
      score: 78,
      mission_priority: "HIGH",
      description: "Mission will assess physical progress against loan covenants and updated baseline.",
      items: [
        { item: "Quarterly progress report (Q4 2025) submitted to ADB", status: "PASS", ref: "ADB TA" },
        { item: "Progress aligned with updated implementation schedule — no critical delay", status: "OBS", ref: "CMAA" },
        { item: "Milestone covenants — all due milestones achieved", status: "PASS", ref: "Loan Agmt" },
        { item: "Design changes properly authorised and communicated to ADB", status: "PASS", ref: "ADB TA" },
        { item: "EVM / S-curve available — physical progress quantified", status: "OBS", ref: "PMI EVM" },
        { item: "Photo documentation current — site photos with date and location stamp", status: "PASS", ref: "ADB TA" },
      ]
    },
    {
      id: "RD-04",
      name: "Environmental Safeguards",
      score: 71,
      mission_priority: "CRITICAL",
      description: "Mission will verify EMP implementation, environmental monitoring, and response to prior mission observations.",
      items: [
        { item: "Environmental Monitoring Report (EMR) for H2 2025 submitted to ADB", status: "PASS", ref: "ADB SPS" },
        { item: "EMP implementation self-assessment — current and filed", status: "PASS", ref: "ADB SPS" },
        { item: "Prior mission environmental observations — all responded to in writing", status: "FAIL", ref: "ADB SPS" },
        { item: "Construction Environmental Management Plan (CEMP) approved and in use", status: "PASS", ref: "ADB SPS" },
        { item: "Environmental compliance certificates from local authorities current", status: "OBS", ref: "Local Law" },
        { item: "Climate adaptation measures in design documented", status: "OBS", ref: "GCF" },
      ]
    },
    {
      id: "RD-05",
      name: "Social Safeguards & Gender",
      score: 68,
      mission_priority: "CRITICAL",
      description: "Mission will verify resettlement implementation, GAP progress, and social monitoring.",
      items: [
        { item: "Resettlement Monitoring Report (RMR) — semi-annual report current", status: "PASS", ref: "ADB SPS" },
        { item: "All affected households compensated before civil works", status: "PASS", ref: "ADB SPS" },
        { item: "Gender Action Plan (GAP) — semi-annual progress report submitted on time", status: "FAIL", ref: "ADB GAP" },
        { item: "Gender targets (female employment ≥ 20%) — current status reported", status: "OBS", ref: "ADB GAP" },
        { item: "Indigenous Peoples Plan (IPP) — implementation current (if applicable)", status: "N/A", ref: "ADB SPS" },
        { item: "Community grievance mechanism — register current, all complaints responded to", status: "PASS", ref: "IFC PS1" },
        { item: "Labour retrenchment plan in place (if applicable)", status: "N/A", ref: "ILO" },
      ]
    },
    {
      id: "RD-06",
      name: "Governance & Anti-Corruption",
      score: 80,
      mission_priority: "MEDIUM",
      description: "Mission will review integrity systems, disclosure compliance, and project governance structures.",
      items: [
        { item: "Anti-corruption and integrity officer designated and active", status: "PASS", ref: "ADB ICA" },
        { item: "Whistleblower/reporting mechanism — staff aware and functional", status: "PASS", ref: "ADB ICA" },
        { item: "No integrity violations reported or unresolved in ADB system", status: "PASS", ref: "ADB ICA" },
        { item: "PIU staffing — all key positions filled (PMC, FM, Safeguards, Procurement)", status: "OBS", ref: "ADB TA" },
        { item: "No covenant breaches outstanding from last loan review", status: "PASS", ref: "Loan Agmt" },
        { item: "Project Steering Committee meetings — minutes available and actions tracked", status: "PASS", ref: "CMAA" },
      ]
    },
    {
      id: "RD-07",
      name: "Development Effectiveness & Results",
      score: 74,
      mission_priority: "MEDIUM",
      description: "Mission will review results framework, DMF targets, and development impact evidence.",
      items: [
        { item: "Design and Monitoring Framework (DMF) — outputs and outcomes updated", status: "PASS", ref: "ADB DMF" },
        { item: "Baseline data collected for all DMF indicators with baselines", status: "PASS", ref: "ADB DMF" },
        { item: "Mid-term targets on track — deviation analysis prepared for off-track indicators", status: "OBS", ref: "ADB DMF" },
        { item: "Beneficiary survey / data collection current — sex-disaggregated data available", status: "PASS", ref: "ADB M&E" },
        { item: "Project completion report (PCR) preparation plan (if applicable)", status: "N/A", ref: "ADB TA" },
        { item: "Lessons learned documented for sharing at mission", status: "OBS", ref: "CMAA" },
      ]
    },
  ],

  // Pre-mission action items
  pre_mission_actions: [
    { id: "PMA-001", priority: "CRITICAL", domain: "Environmental Safeguards", action: "Prepare written responses to all 3 prior mission environment observations. Submit to ADB portal by 12 March 2026.", owner: "EA Environment Officer", deadline: "12 March 2026" },
    { id: "PMA-002", priority: "CRITICAL", domain: "Social Safeguards & Gender", action: "Complete and submit overdue GAP semi-annual progress report. Include female employment data and targets vs actuals.", owner: "Social/Gender Officer", deadline: "12 March 2026" },
    { id: "PMA-003", priority: "HIGH", domain: "Financial Management", action: "Update and reconcile fixed asset register. Prepare summary for mission presentation.", owner: "Finance Manager", deadline: "14 March 2026" },
    { id: "PMA-004", priority: "HIGH", domain: "Progress & Physical", action: "Prepare updated S-curve and EVM analysis showing SPI/CPI. Have schedule narrative ready for delays in Zone 3.", owner: "Project Manager", deadline: "14 March 2026" },
    { id: "PMA-005", priority: "HIGH", domain: "Gender", action: "Calculate and document current female employment percentage against 20% target. Prepare catch-up plan if below target.", owner: "Gender Officer", deadline: "14 March 2026" },
    { id: "PMA-006", priority: "MEDIUM", domain: "Governance", action: "Fill vacant Procurement Officer position or confirm acting arrangement in writing. Notify ADB mission leader.", owner: "EA Director", deadline: "15 March 2026" },
  ],

  // Documents for the mission
  documents: [
    { doc: "Project Progress Report (Q4 2025)", status: "READY", location: "ADB Portal + Hard copy" },
    { doc: "Audited Financial Statements (FY2025)", status: "READY", location: "ADB Portal" },
    { doc: "Statement of Expenditure (Q4 2025)", status: "READY", location: "ADB Portal" },
    { doc: "Procurement Plan (updated March 2026)", status: "READY", location: "ADB Online Procurement" },
    { doc: "Environmental Monitoring Report (H2 2025)", status: "READY", location: "ADB Portal" },
    { doc: "Resettlement Monitoring Report (H2 2025)", status: "READY", location: "ADB Portal" },
    { doc: "Gender Action Plan Progress Report (H2 2025)", status: "OVERDUE", location: "Not yet submitted" },
    { doc: "Response to prior mission observations", status: "PENDING", location: "In preparation" },
    { doc: "Design & Monitoring Framework (updated)", status: "READY", location: "ADB Portal + Hard copy" },
    { doc: "Updated implementation schedule / S-curve", status: "PENDING", location: "In preparation" },
    { doc: "Site photographs (Q4 2025)", status: "READY", location: "SharePoint + Hard copy" },
    { doc: "Risk register (latest)", status: "READY", location: "PMU Files" },
  ],
};

// ─── BUILD ─────────────────────────────────────────────────────────────────────
function build(r) {
  const children = [];

  children.push(...H.makeCover(
    "DONOR READINESS REPORT",
    `Pre-Mission Governance Assessment  |  ${r.mission_type}`,
    H.C.NAVY,
    [
      ["Project",            r.project_name],
      ["Executing Agency",   r.executing_agency],
      ["Donor / Funder",     r.donor],
      ["Loan / Grant",       `${r.loan_no}  |  ${r.grant_no}`],
      ["Total Financing",    r.total_financing],
      ["Mission Type",       r.mission_type],
      ["Mission Dates",      r.mission_date],
      ["Readiness Rating",   r.readiness_level],
      ["Report Reference",   r.report_ref],
      ["Prepared by",        r.assessor + ", CCM"],
    ]
  ));

  // Section 1 — Project Details
  children.push(
    H.h1("1. PROJECT & MISSION OVERVIEW"),
    H.divider(),
    H.spacer(80),
    projectOverviewGrid(r),
    H.spacer(160),
  );

  // Section 2 — Readiness Summary
  children.push(
    H.h1("2. READINESS ASSESSMENT SUMMARY"),
    H.divider(),
    H.spacer(80),
    readinessBanner(r.readiness_level),
    H.spacer(100),
    H.body(r.exec_summary, { size: 21 }),
    H.spacer(120),
    H.h2("2.1 COS™ Tri-Pillar Readiness Score"),
    H.spacer(60),
    H.scoreSummary(r.c_score, r.o_score, r.s_score),
    H.spacer(120),
    H.h2("2.2 Domain Readiness Scorecard"),
    H.spacer(60),
    domainScorecard(r.domains),
    H.spacer(160),
  );

  // Section 3 — Pre-Mission Actions
  children.push(
    H.h1("3. PRE-MISSION ACTION ITEMS — MUST COMPLETE BEFORE 16 MARCH 2026"),
    H.divider(H.C.RED),
    H.spacer(80),
    preMissionTable(r.pre_mission_actions),
    H.spacer(160),
  );

  // Section 4 — Domain Checklists
  children.push(
    H.h1("4. DOMAIN READINESS CHECKLISTS"),
    H.divider(),
    H.spacer(80),
  );
  r.domains.forEach(d => {
    children.push(
      H.h2(`${d.id} — ${d.name}  [${d.score}%  |  Mission Priority: ${d.mission_priority}]`),
      H.body(d.description, { italic: true }),
      H.spacer(60),
      H.checklistTable(d.items),
      H.spacer(120),
    );
  });

  // Section 5 — Document Checklist
  children.push(
    H.h1("5. MISSION DOCUMENT CHECKLIST"),
    H.divider(),
    H.spacer(80),
    H.body("All documents listed below should be available in hard copy and/or uploaded to ADB Portal before the mission commences on 16 March 2026.", { italic: true }),
    H.spacer(80),
    documentTable(r.documents),
    H.spacer(160),
  );

  // Section 6 — COS™ Statement
  children.push(
    H.h1("6. COS™ GOVERNANCE STATEMENT"),
    H.divider(H.C.GOLD),
    H.spacer(80),
    cosStatement(),
    H.spacer(160),
  );

  // Section 7 — Sign-Off
  children.push(
    H.h1("7. CERTIFICATION & SIGN-OFF"),
    H.divider(H.C.GOLD),
    H.spacer(80),
    ...H.signOff(r.report_date, "Project Director / EA Representative"),
  );

  return H.makeDoc(r.report_ref, r.report_date, children);
}

// ─── SUB-COMPONENTS ───────────────────────────────────────────────────────────
function projectOverviewGrid(r) {
  const pairs = [
    ["Project Name", r.project_name], ["Project Code", r.project_code],
    ["Executing Agency", r.executing_agency], ["Implementing Agency", r.implementing],
    ["Donor", r.donor], ["Loan / Grant Reference", `${r.loan_no} | ${r.grant_no}`],
    ["Total Financing", r.total_financing], ["Consultant / PMC", r.consultant],
    ["Mission Type", r.mission_type], ["Mission Dates", r.mission_date],
    ["Readiness Rating", r.readiness_level], ["Report Date", r.report_date],
  ];
  return new Table({
    width: { size: 9360, type: WidthType.DXA }, columnWidths: [2000, 2680, 2000, 2680],
    rows: pairs.reduce((rows, _, i, a) => {
      if (i % 2 === 0) rows.push(new TableRow({ children: a.slice(i, i+2).flatMap(([l,v]) => [
        H.cell(l, 2000, H.C.NAVY, { bold: true, size: 18, color: H.C.WHITE }),
        H.cell(v, 2680, i % 4 === 0 ? H.C.LIGHT : H.C.WHITE, { size: 18 }),
      ])}));
      return rows;
    }, [])
  });
}

function readinessBanner(level) {
  const fill  = level.includes("READY") ? H.C.GREEN : level.includes("PARTIAL") ? H.C.AMBER : H.C.RED;
  const color = H.C.WHITE;
  return new Table({
    width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
    rows: [new TableRow({ children: [new TableCell({
      borders: H.NO_BORDERS, shading: { fill, type: ShadingType.CLEAR },
      width: { size: 9360, type: WidthType.DXA },
      margins: { top: 180, bottom: 180, left: 240, right: 240 },
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "OVERALL READINESS RATING", bold: true, size: 20, color, font: "Arial" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: level, bold: true, size: 44, color, font: "Arial" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Assessed using COS™ Methodology — Compliance · Oversight · Sustainability", size: 17, color, italic: true, font: "Arial" })] }),
      ]
    })]})],
  });
}

function domainScorecard(domains) {
  const cols = [800, 3200, 1000, 1200, 1360, 1800];
  const hdr = new TableRow({ tableHeader: true, children:
    ["ID", "Readiness Domain", "Score", "Status", "Mission Priority", "Key Risk if Not Ready"].map((h, i) =>
      H.cell(h, cols[i], H.C.NAVY, { bold: true, size: 16, color: H.C.WHITE, center: true })
    )
  });
  const riskMap = {
    "RD-01": "Disbursement suspension",
    "RD-02": "Misprocurement finding",
    "RD-03": "Loan covenant breach",
    "RD-04": "Safeguard non-compliance notice",
    "RD-05": "Social non-compliance / GAP penalty",
    "RD-06": "Integrity investigation",
    "RD-07": "Performance downgrade",
  };
  const rows = domains.map((d, idx) => {
    const bg = idx % 2 === 0 ? H.C.WHITE : H.C.LIGHT;
    const sc = d.score;
    const statusBg = sc >= 80 ? H.C.PASS : sc >= 65 ? H.C.OBS : H.C.FAIL;
    const statusLabel = sc >= 80 ? "READY" : sc >= 65 ? "PARTIAL" : "NOT READY";
    const statusColor = sc >= 80 ? H.C.GREEN : sc >= 65 ? H.C.AMBER : H.C.RED;
    const mpColor = d.mission_priority === "CRITICAL" ? H.C.RED : d.mission_priority === "HIGH" ? H.C.AMBER : H.C.GREY;
    return new TableRow({ children: [
      H.cell(d.id, cols[0], bg, { size: 16, bold: true }),
      H.cell(d.name, cols[1], bg, { size: 16 }),
      new TableCell({ borders: H.BORDERS, shading: { fill: statusBg, type: ShadingType.CLEAR }, width: { size: cols[2], type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 80, right: 80 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `${sc}%`, bold: true, size: 17, color: H.C.NAVY, font: "Arial" })] })]
      }),
      new TableCell({ borders: H.BORDERS, shading: { fill: statusBg, type: ShadingType.CLEAR }, width: { size: cols[3], type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 80, right: 80 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: statusLabel, bold: true, size: 15, color: statusColor, font: "Arial" })] })]
      }),
      new TableCell({ borders: H.BORDERS, shading: { fill: bg, type: ShadingType.CLEAR }, width: { size: cols[4], type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 80, right: 80 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: d.mission_priority, bold: true, size: 15, color: mpColor, font: "Arial" })] })]
      }),
      H.cell(riskMap[d.id] || "", cols[5], bg, { size: 15, italic: true, color: H.C.RED }),
    ]});
  });
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: cols, rows: [hdr, ...rows] });
}

function preMissionTable(actions) {
  const cols = [840, 1000, 1960, 3600, 1480, 480];
  const hdr = new TableRow({ tableHeader: true, children:
    ["Action", "Priority", "Domain", "Required Action", "Owner", "By"].map((h, i) =>
      H.cell(h, cols[i], H.C.NAVY, { bold: true, size: 16, color: H.C.WHITE, center: true })
    )
  });
  const pColor = p => p === "CRITICAL" ? H.C.RED : p === "HIGH" ? H.C.ORANGE : H.C.GREY;
  const pBg    = p => p === "CRITICAL" ? "FFE8E8" : p === "HIGH" ? "FFF4E0" : H.C.OBS;
  const rows = actions.map((a, idx) => {
    const bg = idx % 2 === 0 ? H.C.WHITE : H.C.LIGHT;
    return new TableRow({ children: [
      H.cell(a.id, cols[0], bg, { size: 15, bold: true }),
      new TableCell({ borders: H.BORDERS, shading: { fill: pBg(a.priority), type: ShadingType.CLEAR }, width: { size: cols[1], type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 60, right: 60 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: a.priority, bold: true, size: 14, color: pColor(a.priority), font: "Arial" })] })]
      }),
      H.cell(a.domain, cols[2], bg, { size: 15, bold: true }),
      H.cell(a.action, cols[3], bg, { size: 15 }),
      H.cell(a.owner, cols[4], bg, { size: 15 }),
      H.cell(a.deadline.replace("March 2026", "Mar"), cols[5], pBg(a.priority), { size: 13, bold: true, color: pColor(a.priority) }),
    ]});
  });
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: cols, rows: [hdr, ...rows] });
}

function documentTable(docs) {
  const cols = [4400, 2560, 2400];
  const hdr = new TableRow({ tableHeader: true, children:
    ["Document", "Status", "Location / Notes"].map((h, i) =>
      H.cell(h, cols[i], H.C.NAVY, { bold: true, size: 17, color: H.C.WHITE, center: true })
    )
  });
  const rows = docs.map((d, idx) => {
    const bg = idx % 2 === 0 ? H.C.WHITE : H.C.LIGHT;
    const statusBg = d.status === "READY" ? H.C.PASS : d.status === "PENDING" ? H.C.OBS : H.C.FAIL;
    const statusColor = d.status === "READY" ? H.C.GREEN : d.status === "PENDING" ? H.C.AMBER : H.C.RED;
    return new TableRow({ children: [
      H.cell(d.doc, cols[0], bg, { size: 17 }),
      new TableCell({ borders: H.BORDERS, shading: { fill: statusBg, type: ShadingType.CLEAR }, width: { size: cols[1], type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 80, right: 80 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: d.status, bold: true, size: 16, color: statusColor, font: "Arial" })] })]
      }),
      H.cell(d.location, cols[2], bg, { size: 16, italic: true }),
    ]});
  });
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: cols, rows: [hdr, ...rows] });
}

function cosStatement() {
  return new Table({
    width: { size: 9360, type: WidthType.DXA }, columnWidths: [9360],
    rows: [new TableRow({ children: [new TableCell({
      borders: { top: { style: BorderStyle.SINGLE, size: 8, color: H.C.GOLD }, bottom: { style: BorderStyle.SINGLE, size: 8, color: H.C.GOLD }, left: { style: BorderStyle.SINGLE, size: 6, color: H.C.GOLD }, right: { style: BorderStyle.SINGLE, size: 6, color: H.C.GOLD } },
      shading: { fill: H.C.LIGHT, type: ShadingType.CLEAR },
      width: { size: 9360, type: WidthType.DXA },
      margins: { top: 200, bottom: 200, left: 280, right: 280 },
      children: [
        new Paragraph({ children: [new TextRun({ text: "COS™ Methodology — Compliance · Oversight · Sustainability", bold: true, size: 24, color: H.C.NAVY, font: "Arial" })] }),
        H.spacer(80),
        new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "This Donor Readiness Report was prepared using the COS™ Methodology — an ethics-first governance framework developed by CM Academy (Nepal), published under DOI: 10.5281/zenodo.18802971. The COS™ Methodology integrates CMAA Standards of Practice, FIDIC Contract Conditions, ISO 9001/31000, IFC Performance Standards, Green Climate Fund Safeguards, UN SDGs, GRI, SASB, and TCFD into a single verifiable governance ecosystem.", size: 19, color: H.C.GREY, font: "Arial" })] }),
        H.spacer(60),
        new Paragraph({ spacing: { before: 40, after: 60 }, children: [new TextRun({ text: "The framework has been validated through 22 years of field practice across Nepal, Bahrain, and Qatar — including Nepal's first underground HV cable installation, Bahrain QA audits for 66–400 kV XLPE cable circuits, and the Eastern Nepal Electric Minibus Climate Pilot.", size: 19, color: H.C.GREY, italic: true, font: "Arial" })] }),
        H.spacer(60),
        new Paragraph({ children: [new TextRun({ text: "Citation: Bhandari, S. (2026). COS™ Methodology White Paper: Compliance · Oversight · Sustainability. Zenodo. DOI: 10.5281/zenodo.18802971", size: 17, color: H.C.NAVY, font: "Arial" })] }),
        H.spacer(60),
        new Paragraph({ children: [new TextRun({ text: "© CM Academy — Complete Construction Management Developers Pvt. Ltd. (Nepal, Reg. No. 275143/078/079) | CC BY 4.0", size: 16, color: H.C.GREY, font: "Arial" })] }),
      ]
    })]})],
  });
}

// ─── RUN ──────────────────────────────────────────────────────────────────────
const doc = build(R);
H.save(doc, `output/COS_Donor_Readiness_Report_${R.report_ref}.docx`)
  .catch(err => { console.error(err); process.exit(1); });
