/**
 * COS™ OSH Field Audit Report Generator
 * CM Academy | Susil Bhandari, CCM
 * DOI: 10.5281/zenodo.18802971
 *
 * Covers: Occupational Safety & Health field audit aligned to
 * COS™ Oversight Pillar — ISO 45001, OSHA, ILO, CMAA standards
 */

'use strict';
const {
  Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, LevelFormat, PageBreak
} = require('docx');
const fs = require('fs');
const H = require('./cos_helpers');

// ─── REPORT DATA ──────────────────────────────────────────────────────────────
const R = {
  project_name:    "Kathmandu Ring Road Improvement Project — Phase III",
  project_code:    "KRR-2026-OSH-012",
  client:          "Department of Roads, Government of Nepal",
  contractor:      "Himalayan Infrastructure Pvt. Ltd.",
  consultant:      "CM Academy | NeoPlan Consult Pvt. Ltd.",
  location:        "Ring Road — Koteshwor to Satdobato Section (3.2 km)",
  contract_no:     "DoR/2026/KRR/012",
  audit_date:      "09 March 2026",
  audit_ref:       "COS-OSH-2026-012-001",
  auditor:         "Susil Bhandari, CCM",
  report_type:     "OSH FIELD AUDIT REPORT",
  report_version:  "v1.0 — Issued for Review",
  workers_on_site: "147",
  site_area:       "3.2 km linear site, active traffic zone",

  // COS™ Pillar Scores
  c_score: 71,   // Compliance — legal/regulatory
  o_score: 68,   // Oversight — duty of care, incident systems
  s_score: 74,   // Sustainability — worker wellbeing, ESG

  // Summary
  exec_summary: "This COS™ OSH Field Audit was conducted on 09 March 2026 across the Koteshwor–Satdobato section of the Kathmandu Ring Road Improvement Project. The audit assessed Occupational Safety and Health (OSH) governance against the COS™ Methodology — Compliance, Oversight, and Sustainability pillars — benchmarked to ISO 45001, OSHA standards, ILO conventions, and CMAA duty of care requirements. 147 workers were present on site at time of audit. The overall OSH compliance stands at 71% (PARTIAL). Immediate corrective action is required for two critical findings: unsafe excavation edge protection at chainage 2+400 and absence of a posted emergency evacuation plan.",

  critical_findings: [
    "CRITICAL — Excavation at Ch. 2+400: No edge protection or shoring installed. Immediate stop-work required until compliance restored.",
    "CRITICAL — Emergency Evacuation Plan not posted at site office or worker camps. Contractor to rectify within 24 hours.",
    "HIGH — Night-shift workers observed without reflective vests in active traffic zone (Midnight–04:00 shift). Immediate PPE enforcement required.",
  ],

  commendations: [
    "First aid kits fully stocked and certified first-aider present on all shifts",
    "Toolbox talk records maintained daily — all workers signed attendance for last 30 days",
    "All heavy machinery operators hold valid licenses — records available on site",
    "Site perimeter hoarding complete and well-maintained along 2.8 km of active frontage",
  ],

  incidents_30days: [
    { date: "14 Feb 2026", type: "Near-Miss",  description: "Falling material from stockpile — no injury", action: "Stockpile re-graded, safety net installed", status: "CLOSED" },
    { date: "21 Feb 2026", type: "First Aid",  description: "Minor hand laceration — worker using grinder without gloves", action: "Re-training conducted, gloves issued", status: "CLOSED" },
    { date: "02 Mar 2026", type: "Near-Miss",  description: "Vehicle reversal near workers — no spotter present", action: "Spotter protocol enforced for all reversals", status: "OPEN" },
  ],
};

// ─── BUILD ────────────────────────────────────────────────────────────────────

function build(r) {
  const children = [];

  // Cover
  children.push(...H.makeCover(
    r.report_type,
    "Compliance · Oversight · Sustainability — Ethics-First OSH Governance",
    H.C.ORANGE,
    [
      ["Project", r.project_name],
      ["Client / Owner", r.client],
      ["Contractor", r.contractor],
      ["Location", r.location],
      ["Workers On Site", r.workers_on_site],
      ["Audit Reference", r.audit_ref],
      ["Audit Date", r.audit_date],
      ["Lead Auditor", r.auditor + ", CCM"],
      ["Report Version", r.report_version],
    ]
  ));

  // ── Section 1: Project Overview ───────────────────────────────────────────
  children.push(
    H.h1("1. PROJECT & AUDIT OVERVIEW"),
    H.divider(),
    H.spacer(80),
    infoGrid([
      ["Project Name", r.project_name], ["Project Code", r.project_code],
      ["Client / Owner", r.client],     ["Contract No.", r.contract_no],
      ["Contractor", r.contractor],     ["Consultant", r.consultant],
      ["Location / Section", r.location], ["Workers Present", r.workers_on_site],
      ["Site Area / Scope", r.site_area], ["Audit Date", r.audit_date],
      ["Lead Auditor", r.auditor + ", CCM"], ["Report Version", r.report_version],
    ]),
    H.spacer(160),
  );

  // ── Section 2: Executive Summary ──────────────────────────────────────────
  children.push(
    H.h1("2. EXECUTIVE SUMMARY"),
    H.divider(),
    H.spacer(80),
    H.body(r.exec_summary, { size: 21 }),
    H.spacer(120),
    H.h2("2.1 COS™ Tri-Pillar OSH Compliance Score"),
    H.spacer(60),
    H.scoreSummary(r.c_score, r.o_score, r.s_score),
    H.spacer(160),
  );

  // ── Section 3: Critical Findings (red-flag box) ───────────────────────────
  children.push(
    H.h1("3. CRITICAL FINDINGS — IMMEDIATE ACTION REQUIRED"),
    H.divider(H.C.RED),
    H.spacer(80),
    criticalBox(r.critical_findings),
    H.spacer(160),
  );

  // ── Section 4: Compliance Pillar ──────────────────────────────────────────
  children.push(
    H.h1("4. COMPLIANCE PILLAR — C"),
    H.divider(H.C.GREEN),
    H.spacer(80),
    H.pillarBanner("C", "Compliance", H.C.GREEN, "Regulatory compliance, legal OSH obligations, documentation standards  |  ISO 45001 · OSHA · ILO · Local Labour Law · CMAA"),
    H.spacer(120),
    H.h2("4.1 Legal & Regulatory Compliance"),
    H.spacer(60),
    H.checklistTable([
      { item: "OSH Management Plan approved by client/consultant and available on site", party: "Contractor / PMC", status: "PASS", ref: "ISO 45001" },
      { item: "OSH legal register maintained and reviewed for applicable Nepal OSH regulations", party: "HSE Officer", status: "PASS", ref: "Labour Act" },
      { item: "All workers registered for Employees Provident Fund (EPF/SSF) as required", party: "Contractor HR", status: "OBS", ref: "Labour Act" },
      { item: "Workers compensation insurance policy current and posted on site", party: "Contractor", status: "PASS", ref: "Labour Act" },
      { item: "Female workers — separate facilities and no night-shift without consent", party: "Site Manager", status: "N/A", ref: "ILO C89" },
      { item: "Child labour prohibition — no workers under 18 employed on site", party: "Contractor HR", status: "PASS", ref: "ILO C138" },
      { item: "Working hours within legal maximum — overtime records maintained", party: "Contractor HR", status: "OBS", ref: "Labour Act" },
      { item: "OSH committee established and meeting minutes available (>50 workers)", party: "HSE Officer", status: "PASS", ref: "Nepal OSH" },
    ]),
    H.spacer(120),
    H.h2("4.2 Documentation & Certification"),
    H.spacer(60),
    H.checklistTable([
      { item: "All workers hold valid induction certificate for this project", party: "HSE Officer", status: "PASS", ref: "CMAA" },
      { item: "Heavy equipment operators — valid licenses and medical fitness certificates", party: "Plant Manager", status: "PASS", ref: "Local" },
      { item: "Chemical MSDS/SDS sheets available and understood by users", party: "HSE Officer", status: "OBS", ref: "ISO 45001" },
      { item: "Permit-to-Work system active for confined space, electrical, hot work", party: "HSE Officer", status: "PASS", ref: "ISO 45001" },
      { item: "Scaffold inspection records current — signed by competent inspector", party: "Site Engineer", status: "PASS", ref: "OSHA" },
    ]),
    H.spacer(160),
  );

  // ── Section 5: Oversight Pillar ───────────────────────────────────────────
  children.push(
    H.h1("5. OVERSIGHT PILLAR — O"),
    H.divider(H.C.NAVY),
    H.spacer(80),
    H.pillarBanner("O", "Oversight", H.C.NAVY, "Duty of care, incident reporting, real-time risk management, proactive supervision  |  CMAA · ISO 45001 · FIDIC"),
    H.spacer(120),
    H.h2("5.1 PPE & Personal Safety"),
    H.spacer(60),
    H.checklistTable([
      { item: "Safety helmets worn by 100% of workers in all active zones", party: "HSE Officer", status: "PASS", ref: "ISO 45001" },
      { item: "Safety footwear — steel-toe boots observed in all work areas", party: "HSE Officer", status: "PASS", ref: "OSHA" },
      { item: "High-visibility vests — worn by all workers, including night shift", party: "HSE Officer", status: "FAIL", ref: "OSHA" },
      { item: "Hearing protection provided and worn near loud machinery", party: "HSE Officer", status: "PASS", ref: "ISO 45001" },
      { item: "Eye protection worn by all grinder, cutter, welder operators", party: "HSE Officer", status: "OBS", ref: "OSHA" },
      { item: "Respiratory protection available in dusty zones", party: "HSE Officer", status: "PASS", ref: "ISO 45001" },
    ]),
    H.spacer(120),
    H.h2("5.2 Site Safety Infrastructure"),
    H.spacer(60),
    H.checklistTable([
      { item: "Site perimeter hoarding / fencing complete and secure", party: "Site Manager", status: "PASS", ref: "Local" },
      { item: "Traffic management plan active — spotters and signage in place", party: "Traffic Marshal", status: "PASS", ref: "DoR" },
      { item: "Excavation edge protection / shoring in place at all open trenches", party: "Site Engineer", status: "FAIL", ref: "OSHA" },
      { item: "Temporary lighting adequate for night works — minimum 50 lux", party: "Site Electrician", status: "PASS", ref: "OSHA" },
      { item: "Fire extinguishers — placed at fuel storage, welding, office areas", party: "HSE Officer", status: "PASS", ref: "NFPA" },
      { item: "First aid kits stocked and accessible — list of contents posted", party: "HSE Officer", status: "PASS", ref: "ISO 45001" },
      { item: "Emergency evacuation plan posted — mustering point identified and marked", party: "HSE Officer", status: "FAIL", ref: "ISO 45001" },
      { item: "Crane / lifting equipment — annual inspection certificate current", party: "Plant Manager", status: "PASS", ref: "Local" },
    ]),
    H.spacer(120),
    H.h2("5.3 Incident Register — Last 30 Days"),
    H.spacer(60),
    incidentTable(r.incidents_30days),
    H.spacer(120),
    H.h2("5.4 Toolbox Talk & Training Records"),
    H.spacer(60),
    H.checklistTable([
      { item: "Daily toolbox talks conducted and attendance signed by all workers", party: "HSE Officer", status: "PASS", ref: "CMAA" },
      { item: "Weekly HSE inspection by resident engineer — report on file", party: "Consultant", status: "PASS", ref: "FIDIC" },
      { item: "Emergency drill conducted in last 90 days — records available", party: "HSE Officer", status: "OBS", ref: "ISO 45001" },
      { item: "Manual handling training provided to all excavation and lifting crews", party: "HSE Officer", status: "PASS", ref: "ISO 45001" },
    ]),
    H.spacer(160),
  );

  // ── Section 6: Sustainability Pillar ─────────────────────────────────────
  children.push(
    H.h1("6. SUSTAINABILITY PILLAR — S"),
    H.divider(H.C.GOLD),
    H.spacer(80),
    H.pillarBanner("S", "Sustainability", H.C.GOLD, "Worker wellbeing, community protection, ESG reporting, long-term resilience  |  UN SDGs · ILO · IFC PS · GRI"),
    H.spacer(120),
    H.h2("6.1 Worker Welfare & Wellbeing"),
    H.spacer(60),
    H.checklistTable([
      { item: "Drinking water — clean potable water available at all work zones", party: "Site Manager", status: "PASS", ref: "ILO" },
      { item: "Sanitation — adequate toilet facilities for workers (1 per 25 workers)", party: "Site Manager", status: "PASS", ref: "ILO" },
      { item: "Rest area / shade — covered rest shelter for workers during breaks", party: "Contractor", status: "OBS", ref: "ILO" },
      { item: "Worker accommodation (if applicable) — inspected and meets basic standards", party: "Contractor", status: "N/A", ref: "IFC PS2" },
      { item: "Grievance mechanism communicated to all workers", party: "Site Manager", status: "OBS", ref: "IFC PS2" },
      { item: "No forced labour or debt bondage — workers free to leave employment", party: "Contractor HR", status: "PASS", ref: "ILO C29" },
    ]),
    H.spacer(120),
    H.h2("6.2 Community & Environmental Protection"),
    H.spacer(60),
    H.checklistTable([
      { item: "Dust suppression active — water spraying in dry conditions", party: "Site Supervisor", status: "PASS", ref: "EMP" },
      { item: "Noise limits observed — no heavy percussion work beyond 18:00", party: "Site Supervisor", status: "PASS", ref: "Local" },
      { item: "Community access maintained — pedestrian diversions safe and lit", party: "Traffic Marshal", status: "PASS", ref: "DoR" },
      { item: "Construction waste segregated — hazardous waste separately stored", party: "Site Supervisor", status: "OBS", ref: "GRI" },
      { item: "No uncontrolled discharge to drains or watercourses", party: "Site Supervisor", status: "PASS", ref: "IFC PS3" },
      { item: "Carbon/GHG emissions log maintained for equipment and transport", party: "Sustainability Officer", status: "OBS", ref: "TCFD" },
    ]),
    H.spacer(160),
  );

  // ── Section 7: Commendations ──────────────────────────────────────────────
  children.push(
    H.h1("7. COMMENDATIONS & GOOD PRACTICE"),
    H.divider(H.C.GREEN),
    H.spacer(80),
    H.body("The following practices were observed as exemplary and are formally commended:", { italic: true }),
    H.spacer(60),
    ...r.commendations.map(c => H.bullet(c, H.C.GREEN)),
    H.spacer(160),
  );

  // ── Section 8: Corrective Action Plan ────────────────────────────────────
  children.push(
    H.h1("8. CORRECTIVE ACTION PLAN (CAP)"),
    H.divider(),
    H.spacer(80),
    H.body("The contractor is required to submit a written Corrective Action Plan (CAP) within 48 hours of receipt of this report. The CAP must address all FAIL and OBS items. All CRITICAL items require immediate action before resumption of affected works.", { size: 20 }),
    H.spacer(80),
    capTable([
      { no: "CAP-001", item: "Excavation edge protection — Ch. 2+400", priority: "CRITICAL", action: "Install timber shoring and edge barriers immediately. Stop work until compliant.", responsible: "Contractor Site Manager", deadline: "Immediate — same day" },
      { no: "CAP-002", item: "Emergency evacuation plan not posted", priority: "CRITICAL", action: "Print, laminate and post evacuation plan at site office, welfare area, and main site entrance.", responsible: "HSE Officer", deadline: "Within 24 hours" },
      { no: "CAP-003", item: "Night-shift reflective vests — non-compliance", priority: "HIGH", action: "Issue mandatory reflective vests to all night-shift workers. Conduct re-induction. HSE officer to monitor first 3 nights.", responsible: "HSE Officer", deadline: "Before next night shift" },
      { no: "CAP-004", item: "EPF/SSF registration — workers not yet registered", priority: "MEDIUM", action: "Submit registration for all unregistered workers within 7 days. Provide evidence to consultant.", responsible: "Contractor HR", deadline: "Within 7 days" },
      { no: "CAP-005", item: "Carbon / GHG emissions log absent", priority: "MEDIUM", action: "Establish equipment fuel consumption log. Submit template to consultant for approval.", responsible: "Sustainability Officer", deadline: "Within 14 days" },
      { no: "CAP-006", item: "Worker rest/shade facility inadequate", priority: "LOW", action: "Erect temporary shade shelter at main work zone welfare area.", responsible: "Site Manager", deadline: "Within 7 days" },
    ]),
    H.spacer(160),
  );

  // ── Section 9: Sign-Off ───────────────────────────────────────────────────
  children.push(
    H.h1("9. CERTIFICATION & SIGN-OFF"),
    H.divider(H.C.GOLD),
    H.spacer(80),
    ...H.signOff(r.audit_date, "Contractor Site Manager / HSE Officer"),
  );

  return H.makeDoc(r.audit_ref, r.audit_date, children);
}

// ─── SUB-COMPONENTS ───────────────────────────────────────────────────────────

function infoGrid(pairs) {
  const rows = [];
  for (let i = 0; i < pairs.length; i += 2) {
    const left  = pairs[i];
    const right = pairs[i + 1] || ["", ""];
    rows.push(new TableRow({
      children: [
        H.cell(left[0],  2000, H.C.NAVY,  { bold: true, size: 18, color: H.C.WHITE }),
        H.cell(left[1],  2680, i % 4 === 0 ? H.C.LIGHT : H.C.WHITE, { size: 18 }),
        H.cell(right[0], 2000, H.C.NAVY,  { bold: true, size: 18, color: H.C.WHITE }),
        H.cell(right[1], 2680, i % 4 === 0 ? H.C.LIGHT : H.C.WHITE, { size: 18 }),
      ]
    }));
  }
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: [2000, 2680, 2000, 2680], rows });
}

function criticalBox(findings) {
  const { Table, TableRow, TableCell, Paragraph, TextRun, BorderStyle, WidthType, ShadingType } = require('docx');
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [
      new TableRow({
        children: [new TableCell({
          borders: { top: { style: BorderStyle.SINGLE, size: 8, color: H.C.RED }, bottom: { style: BorderStyle.SINGLE, size: 8, color: H.C.RED }, left: { style: BorderStyle.SINGLE, size: 8, color: H.C.RED }, right: { style: BorderStyle.SINGLE, size: 8, color: H.C.RED } },
          shading: { fill: "FFF0F0", type: ShadingType.CLEAR },
          width: { size: 9360, type: WidthType.DXA },
          margins: { top: 160, bottom: 160, left: 240, right: 240 },
          children: [
            new Paragraph({ children: [new TextRun({ text: "⚠  CRITICAL — IMMEDIATE ACTION REQUIRED", bold: true, size: 22, color: H.C.RED, font: "Arial" })] }),
            H.spacer(80),
            ...findings.map(f => new Paragraph({
              numbering: { reference: "numbers", level: 0 },
              spacing: { before: 60, after: 60 },
              children: [new TextRun({ text: f, size: 20, bold: true, color: H.C.RED, font: "Arial" })]
            }))
          ]
        })]
      })
    ]
  });
}

function incidentTable(incidents) {
  const cols = [1200, 1200, 3000, 2160, 800];
  const { Table, TableRow, TableCell, Paragraph, TextRun, BorderStyle, WidthType, ShadingType } = require('docx');

  if (!incidents || incidents.length === 0) {
    return H.body("No incidents recorded in the last 30 days. ✓", { color: H.C.GREEN, bold: true });
  }

  const hdr = new TableRow({
    tableHeader: true,
    children: ["Date", "Type", "Description", "Corrective Action", "Status"].map((h, i) =>
      new TableCell({
        borders: H.BORDERS, shading: { fill: H.C.NAVY, type: ShadingType.CLEAR },
        width: { size: cols[i], type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 100, right: 100 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, bold: true, size: 17, color: H.C.WHITE, font: "Arial" })] })]
      })
    )
  });

  const rows = incidents.map((inc, idx) => {
    const bg = idx % 2 === 0 ? H.C.WHITE : H.C.LIGHT;
    const statusBg = inc.status === "CLOSED" ? H.C.PASS : H.C.FAIL;
    const statusColor = inc.status === "CLOSED" ? H.C.GREEN : H.C.RED;
    return new TableRow({
      children: [
        H.cell(inc.date, cols[0], bg, { size: 17 }),
        H.cell(inc.type, cols[1], bg, { size: 17, bold: true }),
        H.cell(inc.description, cols[2], bg, { size: 17 }),
        H.cell(inc.action, cols[3], bg, { size: 17 }),
        new TableCell({
          borders: H.BORDERS, shading: { fill: statusBg, type: ShadingType.CLEAR },
          width: { size: cols[4], type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 80, right: 80 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: inc.status, bold: true, size: 16, color: statusColor, font: "Arial" })] })]
        })
      ]
    });
  });

  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: cols, rows: [hdr, ...rows] });
}

function capTable(items) {
  const cols = [800, 2200, 900, 2500, 1800, 1160];
  const { Table, TableRow, ShadingType } = require('docx');

  const headers = ["CAP No.", "Non-Conformance / Observation", "Priority", "Required Corrective Action", "Responsible", "Deadline"];
  const priorityColor = p => p === "CRITICAL" ? H.C.RED : p === "HIGH" ? H.C.ORANGE : p === "MEDIUM" ? "6B4E00" : H.C.GREY;
  const priorityBg   = p => p === "CRITICAL" ? "FFE0E0" : p === "HIGH" ? "FFF0D6" : p === "MEDIUM" ? H.C.OBS : H.C.LIGHT;

  const hdr = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => H.cell(h, cols[i], H.C.NAVY, { bold: true, size: 17, color: H.C.WHITE, center: true }))
  });

  const rows = items.map((item, idx) => {
    const bg = idx % 2 === 0 ? H.C.WHITE : H.C.LIGHT;
    return new TableRow({
      children: [
        H.cell(item.no, cols[0], bg, { size: 17, bold: true, center: true }),
        H.cell(item.item, cols[1], bg, { size: 17 }),
        H.cell(item.priority, cols[2], priorityBg(item.priority), { size: 16, bold: true, color: priorityColor(item.priority), center: true }),
        H.cell(item.action, cols[3], bg, { size: 17 }),
        H.cell(item.responsible, cols[4], bg, { size: 17 }),
        H.cell(item.deadline, cols[5], bg, { size: 17, bold: true }),
      ]
    });
  });

  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: cols, rows: [hdr, ...rows] });
}

// ─── RUN ──────────────────────────────────────────────────────────────────────
const doc = build(R);
H.save(doc, `output/COS_OSH_Audit_Report_${R.audit_ref}.docx`)
  .catch(err => { console.error(err); process.exit(1); });
