/**
 * COS™ NCR Register & Closure Report Generator
 * CM Academy | Susil Bhandari, CCM
 * DOI: 10.5281/zenodo.18802971
 *
 * Non-Conformance Report (NCR) Register with:
 *   - Full NCR log with status tracking
 *   - Root cause analysis per NCR
 *   - Corrective & Preventive Action (CAPA) tracking
 *   - COS™ Tri-Pillar compliance scoring
 *   - Closure certification
 */

'use strict';
const {
  Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
  VerticalAlign, LevelFormat
} = require('docx');
const fs = require('fs');
const H = require('./cos_helpers');

// ─── REPORT DATA ──────────────────────────────────────────────────────────────
const R = {
  project_name:    "Hetauda–Narayanghat Road Upgrading Project",
  project_code:    "HNR-2026-NCR-REG",
  client:          "Department of Roads, Government of Nepal / ADB Loan 4412-NEP",
  contractor:      "Everest Road Builders Consortium",
  consultant:      "CM Academy | NeoPlan Consult Pvt. Ltd.",
  location:        "Hetauda–Narayanghat Highway, Full Alignment (76.4 km)",
  contract_no:     "DoR/ADB/2026/HNR/008",
  period:          "01 January 2026 — 09 March 2026",
  issue_date:      "09 March 2026",
  report_ref:      "COS-NCR-2026-REG-001",
  auditor:         "Susil Bhandari, CCM",
  report_version:  "v1.0 — Issued for Review",

  // Summary statistics
  total_ncrs:    12,
  closed:         7,
  open:           3,
  overdue:        2,
  obs:            0,

  // COS™ Scores
  c_score: 74,
  o_score: 69,
  s_score: 72,

  exec_summary: "This COS™ NCR Register covers the period 01 January 2026 to 09 March 2026 for the Hetauda–Narayanghat Road Upgrading Project, financed under ADB Loan 4412-NEP. A total of 12 Non-Conformance Reports (NCRs) have been raised. Of these, 7 are closed with verified corrective actions, 3 are open within deadline, and 2 are overdue and require immediate escalation. The NCR register has been maintained in accordance with ISO 9001:2015 Clause 10.2, FIDIC contract conditions, and the COS™ Compliance Pillar requirements. Root cause analysis and CAPA have been completed for all closed NCRs.",

  ncrs: [
    {
      ref: "NCR-001",
      date_raised: "05 Jan 2026",
      description: "Concrete compressive strength test results below specification (25 MPa required, 19.4 MPa achieved) — Ch. 14+200 culvert headwall",
      location: "Ch. 14+200, Culvert HW-14",
      category: "Quality — Materials",
      raised_by: "QA/QC Engineer",
      responsible: "Contractor",
      root_cause: "Incorrect water-cement ratio used by batching plant operator. No independent verification of mix design before pour.",
      corrective_action: "Defective headwall demolished and reconstructed. Mix design re-verified by approved lab. Batching plant operator re-trained.",
      preventive_action: "Pre-pour checklist introduced requiring QA/QC sign-off on W/C ratio before any concrete pour.",
      due_date: "25 Jan 2026",
      closure_date: "22 Jan 2026",
      status: "CLOSED",
      verified_by: "Susil Bhandari, CCM",
    },
    {
      ref: "NCR-002",
      date_raised: "09 Jan 2026",
      description: "Subgrade compaction test failing — density ratio 88% achieved vs 95% minimum specified — Ch. 22+400 to 23+000",
      location: "Ch. 22+400–23+000",
      category: "Quality — Earthworks",
      raised_by: "Resident Engineer",
      responsible: "Contractor",
      root_cause: "Roller passes insufficient. Equipment malfunction (roller drum cracked) not reported to supervisor.",
      corrective_action: "Section re-scarified, moisture-conditioned and re-compacted with operational roller. 3 test points all PASS at 97%.",
      preventive_action: "Daily plant checklist introduced. Compaction test required every 500m or at end of each day's work.",
      due_date: "28 Jan 2026",
      closure_date: "26 Jan 2026",
      status: "CLOSED",
      verified_by: "Susil Bhandari, CCM",
    },
    {
      ref: "NCR-003",
      date_raised: "14 Jan 2026",
      description: "Bituminous surface course laid without prime coat application — Ch. 31+000 to 31+600",
      location: "Ch. 31+000–31+600",
      category: "Quality — Pavement",
      raised_by: "QA/QC Engineer",
      responsible: "Contractor",
      root_cause: "Foreman proceeded without checking ITP. Prime coat material was awaiting delivery from Kathmandu — no hold-point enforced.",
      corrective_action: "Surface course milled and removed. Prime coat applied. Surface course re-laid after 24-hour cure.",
      preventive_action: "ITP hold-points now require written consultant sign-off before proceeding to next activity.",
      due_date: "31 Jan 2026",
      closure_date: "29 Jan 2026",
      status: "CLOSED",
      verified_by: "Susil Bhandari, CCM",
    },
    {
      ref: "NCR-004",
      date_raised: "20 Jan 2026",
      description: "Reinforcement bar diameter non-conforming — 12mm bars used in place of 16mm specified — Bridge pier column at Ch. 42+700",
      location: "Ch. 42+700, Bridge Pier P3",
      category: "Quality — Structural",
      raised_by: "Structural Engineer",
      responsible: "Contractor",
      root_cause: "Material delivery labelling error. No incoming material verification conducted before placement.",
      corrective_action: "Pier column demolished. Correct 16mm bars delivered with mill certificates. Pier reconstructed and tested.",
      preventive_action: "Incoming materials verification procedure updated — all rebar to be measured and tagged before placement.",
      due_date: "15 Feb 2026",
      closure_date: "12 Feb 2026",
      status: "CLOSED",
      verified_by: "Susil Bhandari, CCM",
    },
    {
      ref: "NCR-005",
      date_raised: "28 Jan 2026",
      description: "Survey setting-out error — road centreline deviated 340mm from design alignment at Ch. 55+100",
      location: "Ch. 55+100",
      category: "Survey / Setting Out",
      raised_by: "Survey Engineer",
      responsible: "Contractor",
      root_cause: "Control point benchmark disturbed by vehicle. Not re-checked before resuming setting-out works.",
      corrective_action: "All benchmarks re-established and verified with authority. Affected earthwork section re-graded to correct alignment.",
      preventive_action: "Benchmark inspection added to daily site checklist. All control points fenced and clearly marked.",
      due_date: "18 Feb 2026",
      closure_date: "14 Feb 2026",
      status: "CLOSED",
      verified_by: "Susil Bhandari, CCM",
    },
    {
      ref: "NCR-006",
      date_raised: "05 Feb 2026",
      description: "Stone masonry retaining wall — joints not fully filled with mortar, >15% void ratio observed — Ch. 61+200",
      location: "Ch. 61+200, RW-61A",
      category: "Quality — Masonry",
      raised_by: "QA/QC Engineer",
      responsible: "Contractor",
      root_cause: "Labour productivity pressure led to rushing of mortar application. No QA check before backfilling.",
      corrective_action: "Wall dismantled and rebuilt with compliant mortar filling. QA/QC inspection conducted before backfilling.",
      preventive_action: "Void ratio test now required per FIDIC for all masonry sections before backfill approval.",
      due_date: "25 Feb 2026",
      closure_date: "21 Feb 2026",
      status: "CLOSED",
      verified_by: "Susil Bhandari, CCM",
    },
    {
      ref: "NCR-007",
      date_raised: "12 Feb 2026",
      description: "Drainage pipe installation — incorrect gradient (0.1% instead of 0.5% minimum) causing standing water — Ch. 67+400",
      location: "Ch. 67+400, Drain D-067",
      category: "Quality — Drainage",
      raised_by: "Resident Engineer",
      responsible: "Contractor",
      root_cause: "Laser level equipment battery failed mid-installation. Work continued by eye without instrument verification.",
      corrective_action: "Drainage pipe excavated, re-graded to 0.6% gradient and reinstated. As-built survey confirms compliance.",
      preventive_action: "Gradient verification by instrument mandatory at each 10m interval during pipe installation.",
      due_date: "04 Mar 2026",
      closure_date: "01 Mar 2026",
      status: "CLOSED",
      verified_by: "Susil Bhandari, CCM",
    },
    {
      ref: "NCR-008",
      date_raised: "18 Feb 2026",
      description: "Concrete bridge deck surface — honeycombing observed over 2.4 m² area on underside of Span 2 — Ch. 42+700 bridge",
      location: "Ch. 42+700, Bridge Span 2 soffit",
      category: "Quality — Concrete",
      raised_by: "QA/QC Engineer",
      responsible: "Contractor",
      root_cause: "Inadequate vibration during concrete pour in confined formwork area.",
      corrective_action: "Structural engineer assessment commissioned. Repair specification approved. Epoxy injection and surface repair ongoing.",
      preventive_action: "Vibrator spacing protocol tightened for confined deck sections. Second vibrator operator assigned.",
      due_date: "10 Mar 2026",
      closure_date: null,
      status: "OPEN",
      verified_by: null,
    },
    {
      ref: "NCR-009",
      date_raised: "22 Feb 2026",
      description: "Slope protection — geotextile installed without approved shop drawings or consultant approval — Ch. 48+600",
      location: "Ch. 48+600, Slope SP-48",
      category: "Quality — Geotechnical",
      raised_by: "Resident Engineer",
      responsible: "Contractor",
      root_cause: "Contractor proceeded without hold-point clearance. Shop drawings submitted but approval pending.",
      corrective_action: "Works suspended. Shop drawings under review. Will reinstate under approved drawings.",
      preventive_action: "Revised ITP issued — all geotextile works require written approval before mobilisation.",
      due_date: "14 Mar 2026",
      closure_date: null,
      status: "OPEN",
      verified_by: null,
    },
    {
      ref: "NCR-010",
      date_raised: "28 Feb 2026",
      description: "Asphalt binder course temperature below specification (140°C minimum) on delivery — 4 loads rejected — Ch. 70+000",
      location: "Ch. 70+000",
      category: "Quality — Pavement",
      raised_by: "QA/QC Engineer",
      responsible: "Contractor",
      root_cause: "Transport distance from plant (~82 km) causing temperature loss. No insulated tarpaulin on trucks.",
      corrective_action: "All 4 rejected loads returned. Insulated tarpaulins now fitted. Temperature logs at plant and site established.",
      preventive_action: "Maximum transport time set at 45 minutes. All trucks require insulated covers. Temperature log mandatory.",
      due_date: "15 Mar 2026",
      closure_date: null,
      status: "OPEN",
      verified_by: null,
    },
    {
      ref: "NCR-011",
      date_raised: "01 Mar 2026",
      description: "OSH — workers operating jackhammers without hearing protection — Ch. 35+200 rock cutting zone",
      location: "Ch. 35+200, Rock Cut RC-35",
      category: "OSH — PPE",
      raised_by: "HSE Officer",
      responsible: "Contractor",
      root_cause: "Hearing protection not included in daily PPE issue for rock cutting crew. Oversight by site supervisor.",
      corrective_action: "Work stopped. Hearing protection issued to all 8 workers. Supervisor issued formal warning.",
      preventive_action: "Rock cutting activity added to high-noise register. Hearing protection mandatory checklist item for this zone.",
      due_date: "02 Mar 2026",
      closure_date: null,
      status: "OVERDUE",
      verified_by: null,
    },
    {
      ref: "NCR-012",
      date_raised: "05 Mar 2026",
      description: "Environmental — cement washout discharged directly to natural drainage channel — Ch. 14+800 batching area",
      location: "Ch. 14+800, Batching Plant Area",
      category: "Environment / Sustainability",
      raised_by: "Sustainability Officer",
      responsible: "Contractor",
      root_cause: "Washout pit overflowed due to heavy rainfall. No overflow containment provision.",
      corrective_action: "Overflow stopped. Washout pit capacity increased. Containment bund constructed.",
      preventive_action: "All washout facilities inspected before every rain event. Minimum 1.5x overflows capacity required.",
      due_date: "06 Mar 2026",
      closure_date: null,
      status: "OVERDUE",
      verified_by: null,
    },
  ],
};

// ─── BUILD ────────────────────────────────────────────────────────────────────

function build(r) {
  const children = [];

  // Cover
  children.push(...H.makeCover(
    r.report_type,
    "Non-Conformance Report Register & Closure Tracking",
    H.C.NAVY,
    [
      ["Project",       r.project_name],
      ["Client / Funder", r.client],
      ["Contractor",    r.contractor],
      ["Location",      r.location],
      ["Reporting Period", r.period],
      ["Report Reference", r.report_ref],
      ["Issue Date",    r.issue_date],
      ["Lead Auditor",  r.auditor + ", CCM"],
      ["Report Version", r.report_version],
    ]
  ));

  // ── Section 1: Overview ───────────────────────────────────────────────────
  children.push(
    H.h1("1. PROJECT & REGISTER OVERVIEW"),
    H.divider(),
    H.spacer(80),
    overviewGrid(r),
    H.spacer(160),
  );

  // ── Section 2: Executive Summary ──────────────────────────────────────────
  children.push(
    H.h1("2. EXECUTIVE SUMMARY"),
    H.divider(),
    H.spacer(80),
    H.body(r.exec_summary, { size: 21 }),
    H.spacer(120),
    H.h2("2.1 NCR Status Dashboard"),
    H.spacer(60),
    ncrDashboard(r),
    H.spacer(120),
    H.h2("2.2 COS™ Tri-Pillar Compliance Score"),
    H.spacer(60),
    H.scoreSummary(r.c_score, r.o_score, r.s_score),
    H.spacer(160),
  );

  // ── Section 3: Overdue NCRs (escalation) ─────────────────────────────────
  const overdueNcrs = r.ncrs.filter(n => n.status === "OVERDUE");
  if (overdueNcrs.length > 0) {
    children.push(
      H.h1("3. OVERDUE NCRs — ESCALATION REQUIRED"),
      H.divider(H.C.RED),
      H.spacer(80),
      overdueBox(overdueNcrs),
      H.spacer(160),
    );
  }

  // ── Section 4: Full NCR Register ──────────────────────────────────────────
  children.push(
    H.h1("4. FULL NCR REGISTER"),
    H.divider(),
    H.spacer(80),
    H.body("The following table presents all NCRs raised during the reporting period. Status key: CLOSED = verified and accepted  |  OPEN = within deadline  |  OVERDUE = past deadline, escalation required.", { italic: true }),
    H.spacer(80),
    fullNcrRegister(r.ncrs),
    H.spacer(160),
  );

  // ── Section 5: Closed NCR Detail (CAPA) ───────────────────────────────────
  children.push(
    H.h1("5. CLOSED NCR DETAIL — ROOT CAUSE & CAPA"),
    H.divider(H.C.GREEN),
    H.spacer(80),
    H.body("The following section provides full root cause analysis and Corrective & Preventive Action (CAPA) records for all closed NCRs, in accordance with ISO 9001:2015 Clause 10.2.", { italic: true }),
    H.spacer(80),
  );

  r.ncrs.filter(n => n.status === "CLOSED").forEach((ncr, idx) => {
    children.push(...ncrDetailCard(ncr, idx));
  });

  children.push(H.spacer(160));

  // ── Section 6: Open NCR Summary ───────────────────────────────────────────
  children.push(
    H.h1("6. OPEN NCR STATUS SUMMARY"),
    H.divider(H.C.AMBER),
    H.spacer(80),
    H.body("The following NCRs remain open. The contractor is required to complete corrective actions and submit closure evidence to the consultant for verification.", { size: 20 }),
    H.spacer(80),
  );

  r.ncrs.filter(n => n.status === "OPEN" || n.status === "OVERDUE").forEach((ncr, idx) => {
    children.push(...ncrDetailCard(ncr, idx));
  });

  children.push(H.spacer(160));

  // ── Section 7: Sign-Off ───────────────────────────────────────────────────
  children.push(
    H.h1("7. CERTIFICATION & SIGN-OFF"),
    H.divider(H.C.GOLD),
    H.spacer(80),
    ...H.signOff(r.issue_date, "Contractor Quality Manager"),
  );

  return H.makeDoc(r.report_ref, r.issue_date, children);
}

// ─── SUB-COMPONENTS ───────────────────────────────────────────────────────────

function overviewGrid(r) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2000, 2680, 2000, 2680],
    rows: [
      [["Project Name", r.project_name], ["Project Code", r.project_code]],
      [["Client / Funder", r.client], ["Contract No.", r.contract_no]],
      [["Contractor", r.contractor], ["Consultant", r.consultant]],
      [["Location", r.location], ["Reporting Period", r.period]],
      [["Lead Auditor", r.auditor + ", CCM"], ["Report Version", r.report_version]],
    ].map((pair, i) => new TableRow({
      children: pair.flatMap(([label, value]) => [
        H.cell(label, 2000, H.C.NAVY, { bold: true, size: 18, color: H.C.WHITE }),
        H.cell(value, 2680, i % 2 === 0 ? H.C.LIGHT : H.C.WHITE, { size: 18 }),
      ])
    }))
  });
}

function ncrDashboard(r) {
  const stats = [
    { label: "Total NCRs Raised", value: String(r.total_ncrs), fill: H.C.NAVY,   color: H.C.WHITE },
    { label: "Closed",            value: String(r.closed),     fill: H.C.PASS,   color: H.C.GREEN },
    { label: "Open",              value: String(r.open),       fill: H.C.OBS,    color: H.C.AMBER },
    { label: "Overdue",           value: String(r.overdue),    fill: "FFD0D0",   color: H.C.RED   },
  ];
  const cols = [2340, 2340, 2340, 2340];
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: cols,
    rows: [
      new TableRow({
        children: stats.map((s, i) => new TableCell({
          borders: H.BORDERS,
          shading: { fill: s.fill, type: ShadingType.CLEAR },
          width: { size: cols[i], type: WidthType.DXA },
          margins: { top: 160, bottom: 160, left: 120, right: 120 },
          children: [
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: s.value, bold: true, size: 52, color: s.color, font: "Arial" })] }),
            new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: s.label, bold: true, size: 18, color: s.color, font: "Arial" })] }),
          ]
        }))
      })
    ]
  });
}

function overdueBox(ncrs) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [new TableRow({
      children: [new TableCell({
        borders: { top: { style: BorderStyle.SINGLE, size: 8, color: H.C.RED }, bottom: { style: BorderStyle.SINGLE, size: 8, color: H.C.RED }, left: { style: BorderStyle.SINGLE, size: 8, color: H.C.RED }, right: { style: BorderStyle.SINGLE, size: 8, color: H.C.RED } },
        shading: { fill: "FFF0F0", type: ShadingType.CLEAR },
        width: { size: 9360, type: WidthType.DXA },
        margins: { top: 160, bottom: 160, left: 240, right: 240 },
        children: [
          new Paragraph({ children: [new TextRun({ text: "⚠  OVERDUE NCRs — IMMEDIATE ESCALATION TO CLIENT REQUIRED", bold: true, size: 22, color: H.C.RED, font: "Arial" })] }),
          H.spacer(80),
          ...ncrs.map(n => new Paragraph({
            numbering: { reference: "numbers", level: 0 },
            spacing: { before: 60, after: 60 },
            children: [
              new TextRun({ text: `${n.ref}: `, bold: true, size: 20, color: H.C.RED, font: "Arial" }),
              new TextRun({ text: `${n.description}  |  Due: ${n.due_date}`, size: 20, color: H.C.RED, font: "Arial" }),
            ]
          }))
        ]
      })]
    })]
  });
}

function fullNcrRegister(ncrs) {
  const cols = [640, 1040, 2600, 1400, 840, 1000, 840, 1000];
  const headers = ["No.", "NCR Ref", "Non-Conformance Description", "Location", "Category", "Raised By", "Due Date", "Status"];

  const hdr = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => new TableCell({
      borders: H.BORDERS,
      shading: { fill: H.C.NAVY, type: ShadingType.CLEAR },
      width: { size: cols[i], type: WidthType.DXA },
      margins: { top: 80, bottom: 80, left: 80, right: 80 },
      children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, bold: true, size: 16, color: H.C.WHITE, font: "Arial" })] })]
    }))
  });

  const rows = ncrs.map((n, idx) => {
    const bg = idx % 2 === 0 ? H.C.WHITE : H.C.LIGHT;
    const statusBg    = n.status === "CLOSED" ? H.C.PASS : n.status === "OVERDUE" ? "FFD0D0" : H.C.OBS;
    const statusColor = n.status === "CLOSED" ? H.C.GREEN : n.status === "OVERDUE" ? H.C.RED : H.C.AMBER;
    return new TableRow({
      children: [
        H.cell(`${idx + 1}`, cols[0], bg, { size: 16, center: true }),
        H.cell(n.ref, cols[1], bg, { size: 16, bold: true }),
        H.cell(n.description, cols[2], bg, { size: 16 }),
        H.cell(n.location, cols[3], bg, { size: 16 }),
        H.cell(n.category, cols[4], bg, { size: 16 }),
        H.cell(n.raised_by, cols[5], bg, { size: 16 }),
        H.cell(n.due_date, cols[6], bg, { size: 16 }),
        new TableCell({
          borders: H.BORDERS,
          shading: { fill: statusBg, type: ShadingType.CLEAR },
          width: { size: cols[7], type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 80, right: 80 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: n.status, bold: true, size: 16, color: statusColor, font: "Arial" })] })]
        })
      ]
    });
  });

  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: cols, rows: [hdr, ...rows] });
}

function ncrDetailCard(ncr, idx) {
  const isOpen    = ncr.status === "OPEN" || ncr.status === "OVERDUE";
  const headerBg  = ncr.status === "CLOSED" ? H.C.GREEN : ncr.status === "OVERDUE" ? H.C.RED : H.C.AMBER;
  const statusTxt = ncr.status === "CLOSED" ? `CLOSED — Verified by: ${ncr.verified_by} on ${ncr.closure_date}` : ncr.status === "OVERDUE" ? `OVERDUE — Was due: ${ncr.due_date}` : `OPEN — Due: ${ncr.due_date}`;

  const fields = [
    ["NCR Reference", ncr.ref],
    ["Date Raised", ncr.date_raised],
    ["Location", ncr.location],
    ["Category", ncr.category],
    ["Raised By", ncr.raised_by],
    ["Responsible Party", ncr.responsible],
  ];

  return [
    H.spacer(80),
    // Header bar
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [9360],
      rows: [new TableRow({
        children: [new TableCell({
          borders: H.NO_BORDERS,
          shading: { fill: headerBg, type: ShadingType.CLEAR },
          width: { size: 9360, type: WidthType.DXA },
          margins: { top: 120, bottom: 120, left: 200, right: 200 },
          children: [
            new Paragraph({ children: [new TextRun({ text: `${ncr.ref}  —  ${statusTxt}`, bold: true, size: 20, color: H.C.WHITE, font: "Arial" })] }),
          ]
        })]
      })]
    }),
    // Body
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [9360],
      rows: [new TableRow({
        children: [new TableCell({
          borders: { top: H.NO_BORDERS.top, bottom: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" }, left: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" }, right: { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" } },
          shading: { fill: H.C.LIGHT, type: ShadingType.CLEAR },
          width: { size: 9360, type: WidthType.DXA },
          margins: { top: 140, bottom: 140, left: 200, right: 200 },
          children: [
            // Meta grid
            new Table({
              width: { size: 8960, type: WidthType.DXA },
              columnWidths: [1800, 2680, 1800, 2680],
              rows: [
                new TableRow({ children: fields.slice(0, 2).flatMap(([l, v]) => [H.cell(l, 1800, H.C.NAVY, { bold: true, size: 17, color: H.C.WHITE }), H.cell(v, 2680, H.C.WHITE, { size: 17 })]) }),
                new TableRow({ children: fields.slice(2, 4).flatMap(([l, v]) => [H.cell(l, 1800, H.C.NAVY, { bold: true, size: 17, color: H.C.WHITE }), H.cell(v, 2680, H.C.LIGHT, { size: 17 })]) }),
                new TableRow({ children: fields.slice(4, 6).flatMap(([l, v]) => [H.cell(l, 1800, H.C.NAVY, { bold: true, size: 17, color: H.C.WHITE }), H.cell(v, 2680, H.C.WHITE, { size: 17 })]) }),
              ]
            }),
            H.spacer(100),
            // Description
            new Paragraph({ children: [new TextRun({ text: "Non-Conformance Description:", bold: true, size: 19, color: H.C.NAVY, font: "Arial" })] }),
            new Paragraph({ spacing: { before: 40, after: 80 }, children: [new TextRun({ text: ncr.description, size: 19, color: H.C.GREY, font: "Arial" })] }),
            // Root Cause
            new Paragraph({ children: [new TextRun({ text: "Root Cause Analysis:", bold: true, size: 19, color: H.C.NAVY, font: "Arial" })] }),
            new Paragraph({ spacing: { before: 40, after: 80 }, children: [new TextRun({ text: ncr.root_cause, size: 19, color: H.C.GREY, font: "Arial" })] }),
            // Corrective Action
            new Paragraph({ children: [new TextRun({ text: "Corrective Action Taken:", bold: true, size: 19, color: isOpen ? H.C.AMBER : H.C.GREEN, font: "Arial" })] }),
            new Paragraph({ spacing: { before: 40, after: 80 }, children: [new TextRun({ text: ncr.corrective_action, size: 19, color: H.C.GREY, font: "Arial" })] }),
            // Preventive Action
            new Paragraph({ children: [new TextRun({ text: "Preventive Action (CAPA):", bold: true, size: 19, color: H.C.NAVY, font: "Arial" })] }),
            new Paragraph({ spacing: { before: 40, after: 60 }, children: [new TextRun({ text: ncr.preventive_action, size: 19, color: H.C.GREY, font: "Arial" })] }),
          ]
        })]
      })]
    }),
  ];
}

// ─── RUN ──────────────────────────────────────────────────────────────────────
const doc = build(R);
H.save(doc, `output/COS_NCR_Register_${R.report_ref}.docx`)
  .catch(err => { console.error(err); process.exit(1); });
