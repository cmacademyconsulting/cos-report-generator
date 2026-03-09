/**
 * COS™ ESG Alignment Report
 * CM Academy | Susil Bhandari, CCM
 * DOI: 10.5281/zenodo.18802971
 *
 * ESG performance reporting for construction & infrastructure projects
 * Aligned to: GRI Standards · SASB · TCFD · UN SDGs · IFC PS · GCF
 * Targets: Donor-funded projects, private sector, banks, corporates
 * Price point: $1,000 – $2,500 per report
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
  project_name:    "Upper Trishuli 1 Hydropower Project — Construction Phase",
  project_code:    "UT1-2026-ESG-001",
  organization:    "Nepal Water & Energy Development Company (NWEDC)",
  project_type:    "216 MW Run-of-River Hydropower Project",
  location:        "Rasuwa District, Bagmati Province, Nepal",
  funder:          "IFC · ADB · AIIB · OECD-backed Lenders",
  reporting_period:"Q4 2025 (October – December 2025)",
  report_date:     "09 March 2026",
  report_ref:      "COS-ESG-2026-001-001",
  preparer:        "Susil Bhandari, CCM",
  report_version:  "v1.0 — Issued for Lender Review",

  // Workforce data
  workforce: {
    total: 2847,
    local: 1943,
    national: 712,
    international: 192,
    female_pct: "14%",
    lti_rate: "0.42",
    near_miss: 7,
    fatalities: 0,
  },

  // Environmental data
  environment: {
    ghg_scope1_tco2: "3,847",
    ghg_scope2_tco2: "124",
    ghg_intensity:   "1.34 tCO₂/NPR million spent",
    water_used_m3:   "187,400",
    waste_total_t:   "1,240",
    waste_recycled_pct: "62%",
    hazardous_waste_t:  "18.4",
    land_disturbed_ha:  "94.2",
    trees_compensated:  "12,840",
  },

  // COS™ Pillar scores
  c_score: 77,
  o_score: 72,
  s_score: 80,

  exec_summary: "This COS™ ESG Alignment Report covers Q4 2025 performance for the Upper Trishuli 1 Hydropower Project (216 MW), financed by IFC, ADB, AIIB, and OECD-backed lenders. The report is structured against the COS™ Sustainability Pillar and aligned to GRI Standards (Core), SASB Electric Utilities & Power Generators standard, TCFD recommendations, IFC Performance Standards, and the UN Sustainable Development Goals. Overall ESG performance stands at 76% — demonstrating strong governance foundations with targeted improvement areas in climate risk disclosure (TCFD) and supply chain sustainability. Zero fatalities were recorded in Q4. GHG intensity has improved 8% quarter-on-quarter. The project has compensated 12,840 trees against 9,200 disturbed — exceeding the 1:1.4 offset commitment.",

  // SDG mapping
  sdgs: [
    { sdg: "SDG 7", title: "Affordable & Clean Energy", contribution: "216 MW renewable energy capacity reducing Nepal's fossil fuel dependency. 387,000 households to benefit.", status: "ON TRACK" },
    { sdg: "SDG 8", title: "Decent Work & Economic Growth", contribution: "2,847 direct jobs. 68% local workforce. Minimum wage compliance verified monthly.", status: "ON TRACK" },
    { sdg: "SDG 9", title: "Industry, Innovation, Infrastructure", contribution: "Nepal's largest private infrastructure investment. First AIIB-financed hydropower project.", status: "ON TRACK" },
    { sdg: "SDG 13", title: "Climate Action", contribution: "Displacing ~1.2M tCO₂/yr vs coal baseline. Climate resilience measures in design.", status: "PARTIAL" },
    { sdg: "SDG 15", title: "Life on Land", contribution: "1.4x tree offset ratio. Fish passage structure installed. Biodiversity monitoring quarterly.", status: "ON TRACK" },
    { sdg: "SDG 16", title: "Peace, Justice & Strong Institutions", contribution: "Grievance mechanism active. 47 grievances received, 44 closed within 14 days.", status: "ON TRACK" },
  ],

  // GRI disclosures
  gri: [
    { code: "GRI 2-7",   title: "Employees",                  value: "2,847 workers (1,943 local Nepali, 712 national, 192 international)", status: "DISCLOSED" },
    { code: "GRI 3-3",   title: "Management of Material Topics", value: "Material topics: GHG, water, biodiversity, labour rights, community", status: "DISCLOSED" },
    { code: "GRI 302-1", title: "Energy Consumption",          value: "Diesel: 4,847 kL  |  Grid electricity: 2,140 MWh  |  Total: ~52,300 GJ", status: "DISCLOSED" },
    { code: "GRI 303-3", title: "Water Withdrawal",            value: "187,400 m³ (river + groundwater). 100% within permitted allocation.", status: "DISCLOSED" },
    { code: "GRI 305-1", title: "Direct GHG Emissions (Scope 1)", value: "3,847 tCO₂e — diesel combustion, blasting, mobile plant", status: "DISCLOSED" },
    { code: "GRI 305-2", title: "Indirect GHG Emissions (Scope 2)", value: "124 tCO₂e — purchased grid electricity (NEA)", status: "DISCLOSED" },
    { code: "GRI 306-3", title: "Waste Generated",             value: "1,240 t total (62% recycled, 37% landfilled, 1% incinerated)", status: "DISCLOSED" },
    { code: "GRI 403-2", title: "Hazard Identification",       value: "ISO 45001 compliant HIRA completed for all activities. 8 critical risks active.", status: "DISCLOSED" },
    { code: "GRI 403-9", title: "Work-related Injuries",       value: "LTI rate: 0.42. Near misses: 7. Fatalities: 0. FAC: 4.", status: "DISCLOSED" },
    { code: "GRI 413-1", title: "Local Community Engagement",  value: "47 grievances received, 44 closed. 3 open < 7 days. Community forum: quarterly.", status: "DISCLOSED" },
    { code: "GRI 201-1", title: "Economic Value Generated",    value: "NPR 2.87B disbursed Q4. Local procurement: 64% of total spend.", status: "DISCLOSED" },
    { code: "GRI 205-2", title: "Anti-Corruption Communication", value: "100% staff trained. Zero corruption allegations in Q4.", status: "DISCLOSED" },
  ],

  // TCFD disclosures
  tcfd: [
    { pillar: "Governance", req: "Board oversight of climate-related risks and opportunities", status: "PARTIAL", note: "Board sustainability committee established Q3. First climate briefing held Nov 2025." },
    { pillar: "Strategy", req: "Climate-related risks and opportunities across short, medium, long-term", status: "PARTIAL", note: "Physical climate risks (glacial lake outburst, seismicity) assessed. Transition risk not yet fully quantified." },
    { pillar: "Risk Mgmt", req: "Process for identifying, assessing, managing climate risks", status: "DISCLOSED", note: "Climate risk integrated into ISO 31000 risk register. Reviewed quarterly." },
    { pillar: "Metrics", req: "GHG emissions (Scope 1, 2, 3) and climate-related targets", status: "PARTIAL", note: "Scope 1 & 2 disclosed. Scope 3 (supply chain) assessment initiated — due Q2 2026." },
  ],

  // IFC Performance Standards
  ifc: [
    { ps: "IFC PS1", title: "Assessment & Mgmt of E&S Risks", status: "COMPLIANT", note: "ESIA approved. ESMP updated quarterly. Independent E&S Monitor (IEM) engaged." },
    { ps: "IFC PS2", title: "Labour & Working Conditions",    status: "COMPLIANT", note: "Workers' Accommodation Standard met. Grievance mechanism active. No forced labour." },
    { ps: "IFC PS3", title: "Resource Efficiency & Pollution", status: "PARTIAL",  note: "GHG monitoring in place. Scope 3 incomplete. Water discharge within permit limits." },
    { ps: "IFC PS4", title: "Community Health, Safety, Security", status: "COMPLIANT", note: "Traffic management plan active. Security personnel trained in Voluntary Principles." },
    { ps: "IFC PS5", title: "Land Acquisition & Involuntary Resettlement", status: "COMPLIANT", note: "RAP implemented. 312 households compensated. No outstanding claims." },
    { ps: "IFC PS6", title: "Biodiversity Conservation",     status: "PARTIAL",  note: "Biodiversity offset ratio: 1:1.4 trees. Fish passage installed. Cumulative impact not yet assessed." },
    { ps: "IFC PS7", title: "Indigenous Peoples",            status: "N/A",      note: "No formally designated Indigenous Peoples communities within project area." },
    { ps: "IFC PS8", title: "Cultural Heritage",             status: "COMPLIANT", note: "Chance find procedure active. One find reported in Q4 — DARC notified and managed." },
  ],

  priority_improvements: [
    { area: "TCFD Scope 3 Emissions", action: "Complete supply chain GHG mapping. Include Scope 3 estimate in Q1 2026 report.", deadline: "Q1 2026 Report" },
    { area: "Biodiversity Cumulative Impact", action: "Commission cumulative impact assessment for Trishuli watershed. Due Q3 2026.", deadline: "Q3 2026" },
    { area: "Climate Transition Risk", action: "Quantify regulatory and market transition risks. Include in next TCFD disclosure.", deadline: "Q2 2026 Report" },
    { area: "Supply Chain Sustainability", action: "Develop supply chain code of conduct. Pilot with top 10 suppliers by Q2 2026.", deadline: "Q2 2026" },
  ],
};

// ─── BUILD ─────────────────────────────────────────────────────────────────────
function build(r) {
  const children = [];

  children.push(...H.makeCover(
    "ESG ALIGNMENT REPORT",
    "Environmental · Social · Governance  |  GRI · SASB · TCFD · IFC PS · UN SDGs",
    H.C.GREEN,
    [
      ["Project / Organization", r.project_name],
      ["Project Type", r.project_type],
      ["Funder / Lender", r.funder],
      ["Location", r.location],
      ["Reporting Period", r.reporting_period],
      ["Report Reference", r.report_ref],
      ["Report Date", r.report_date],
      ["Prepared by", r.preparer + ", CCM"],
      ["Report Version", r.report_version],
    ]
  ));

  // Section 1
  children.push(H.h1("1. PROJECT OVERVIEW"), H.divider(), H.spacer(80), projectOverviewGrid(r), H.spacer(160));

  // Section 2 — Summary
  children.push(
    H.h1("2. EXECUTIVE SUMMARY"),
    H.divider(),
    H.spacer(80),
    H.body(r.exec_summary, { size: 21 }),
    H.spacer(120),
    H.h2("2.1 COS™ Tri-Pillar ESG Score"),
    H.spacer(60),
    H.scoreSummary(r.c_score, r.o_score, r.s_score),
    H.spacer(120),
    H.h2("2.2 Key ESG Metrics — Snapshot"),
    H.spacer(60),
    esgSnapshot(r),
    H.spacer(160),
  );

  // Section 3 — Environmental
  children.push(
    H.h1("3. ENVIRONMENTAL PERFORMANCE"),
    H.divider(H.C.GREEN),
    H.spacer(80),
    H.pillarBanner("E", "Environmental", H.C.GREEN, "GHG · Energy · Water · Waste · Biodiversity  |  GRI 302/303/305/306 · IFC PS3/6 · TCFD · GCF"),
    H.spacer(120),
    H.h2("3.1 GHG Emissions & Climate"),
    H.spacer(60),
    ghgTable(r.environment),
    H.spacer(120),
    H.h2("3.2 TCFD Climate Disclosure"),
    H.spacer(60),
    tcfdTable(r.tcfd),
    H.spacer(160),
  );

  // Section 4 — Social
  children.push(
    H.h1("4. SOCIAL PERFORMANCE"),
    H.divider(H.C.NAVY),
    H.spacer(80),
    H.pillarBanner("S", "Social", H.C.NAVY, "Workforce · Safety · Community · Labour Rights  |  GRI 403/413 · IFC PS2/4 · ILO · ADB SPS"),
    H.spacer(120),
    H.h2("4.1 Workforce Summary"),
    H.spacer(60),
    workforceTable(r.workforce),
    H.spacer(120),
    H.h2("4.2 IFC Performance Standards Compliance"),
    H.spacer(60),
    ifcTable(r.ifc),
    H.spacer(160),
  );

  // Section 5 — Governance
  children.push(
    H.h1("5. GOVERNANCE PERFORMANCE"),
    H.divider(H.C.GOLD),
    H.spacer(80),
    H.pillarBanner("G", "Governance", H.C.GOLD, "Transparency · Ethics · Compliance · Accountability  |  GRI 2/205 · SASB · CMAA · FIDIC · Local Laws"),
    H.spacer(120),
    H.h2("5.1 GRI Standards Disclosure"),
    H.spacer(60),
    griTable(r.gri),
    H.spacer(160),
  );

  // Section 6 — SDG Alignment
  children.push(
    H.h1("6. UN SDG ALIGNMENT"),
    H.divider(H.C.NAVY),
    H.spacer(80),
    H.body("The following SDGs are directly contributed to by this project. Contributions have been assessed against project outcomes, workforce data, and environmental performance.", { italic: true }),
    H.spacer(80),
    sdgTable(r.sdgs),
    H.spacer(160),
  );

  // Section 7 — Priority Improvements
  children.push(
    H.h1("7. PRIORITY IMPROVEMENTS"),
    H.divider(),
    H.spacer(80),
    improvementsTable(r.priority_improvements),
    H.spacer(160),
  );

  // Section 8 — Sign-Off
  children.push(
    H.h1("8. CERTIFICATION & SIGN-OFF"),
    H.divider(H.C.GOLD),
    H.spacer(80),
    ...H.signOff(r.report_date, "Organization ESG Representative / Lender"),
  );

  return H.makeDoc(r.report_ref, r.report_date, children);
}

// ─── SUB-COMPONENTS ───────────────────────────────────────────────────────────
function projectOverviewGrid(r) {
  const pairs = [
    ["Project", r.project_name], ["Project Code", r.project_code],
    ["Organization", r.organization], ["Project Type", r.project_type],
    ["Location", r.location], ["Funder / Lender", r.funder],
    ["Reporting Period", r.reporting_period], ["Report Date", r.report_date],
    ["Preparer", r.preparer + ", CCM"], ["Report Version", r.report_version],
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

function esgSnapshot(r) {
  const metrics = [
    { label: "Total Workforce", value: String(r.workforce.total), sub: `${r.workforce.local} local Nepali`, color: H.C.NAVY },
    { label: "LTI Rate",        value: r.workforce.lti_rate,      sub: "per 200,000 hrs",         color: H.C.GREEN },
    { label: "Fatalities",      value: String(r.workforce.fatalities), sub: "Q4 2025",           color: H.C.GREEN },
    { label: "GHG Scope 1",     value: r.environment.ghg_scope1_tco2, sub: "tCO₂e",             color: H.C.AMBER },
    { label: "Water Used",      value: r.environment.water_used_m3,   sub: "m³ Q4",             color: H.C.NAVY },
    { label: "Waste Recycled",  value: r.environment.waste_recycled_pct, sub: "of total waste",  color: H.C.GREEN },
  ];
  const cols = Array(6).fill(1560);
  return new Table({
    width: { size: 9360, type: WidthType.DXA }, columnWidths: cols,
    rows: [new TableRow({ children: metrics.map((m, i) => new TableCell({
      borders: H.BORDERS, shading: { fill: i % 2 === 0 ? H.C.LIGHT : H.C.WHITE, type: ShadingType.CLEAR },
      width: { size: 1560, type: WidthType.DXA }, margins: { top: 140, bottom: 140, left: 100, right: 100 },
      children: [
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: m.value, bold: true, size: 36, color: m.color, font: "Arial" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: m.label, bold: true, size: 16, color: H.C.NAVY, font: "Arial" })] }),
        new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: m.sub, size: 14, color: H.C.GREY, italic: true, font: "Arial" })] }),
      ]
    }))})]
  });
}

function ghgTable(e) {
  const metrics = [
    ["GHG Scope 1 (Direct)", `${e.ghg_scope1_tco2} tCO₂e`, "Diesel combustion, blasting, mobile equipment"],
    ["GHG Scope 2 (Indirect)", `${e.ghg_scope2_tco2} tCO₂e`, "Purchased grid electricity (NEA)"],
    ["GHG Intensity", e.ghg_intensity, "Improving: -8% vs Q3 2025"],
    ["Total Water Withdrawal", `${e.water_used_m3} m³`, "Within permitted allocation. No watercourse violations."],
    ["Total Waste Generated", `${e.waste_total_t} tonnes`, `${e.waste_recycled_pct} recycled | Hazardous: ${e.hazardous_waste_t} t (licensed disposal)`],
    ["Land Disturbance", `${e.land_disturbed_ha} ha`, "Restoration plan active. Topsoil stockpiled for reinstatement."],
    ["Tree Compensation", `${e.trees_compensated} trees`, "Against 9,200 disturbed — 1.4:1 ratio (commitment: 1:1)"],
  ];
  const cols = [2800, 1800, 4760];
  const hdr = new TableRow({ tableHeader: true, children:
    ["Environmental Metric", "Q4 2025 Value", "Notes / Context"].map((h, i) =>
      H.cell(h, cols[i], H.C.NAVY, { bold: true, size: 17, color: H.C.WHITE, center: true })
    )
  });
  const rows = metrics.map(([metric, value, note], idx) => new TableRow({ children: [
    H.cell(metric, cols[0], idx % 2 === 0 ? H.C.WHITE : H.C.LIGHT, { size: 17, bold: true }),
    H.cell(value, cols[1], idx % 2 === 0 ? H.C.LIGHT : H.C.WHITE, { size: 17, bold: true, color: H.C.NAVY }),
    H.cell(note, cols[2], idx % 2 === 0 ? H.C.WHITE : H.C.LIGHT, { size: 17, italic: true }),
  ]}));
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: cols, rows: [hdr, ...rows] });
}

function tcfdTable(items) {
  const cols = [1800, 3600, 1400, 2560];
  const hdr = new TableRow({ tableHeader: true, children:
    ["TCFD Pillar", "Disclosure Requirement", "Status", "Current Disclosure"].map((h, i) =>
      H.cell(h, cols[i], H.C.NAVY, { bold: true, size: 17, color: H.C.WHITE, center: true })
    )
  });
  const rows = items.map((item, idx) => {
    const bg = idx % 2 === 0 ? H.C.WHITE : H.C.LIGHT;
    const statusBg = item.status === "DISCLOSED" ? H.C.PASS : item.status === "PARTIAL" ? H.C.OBS : H.C.FAIL;
    const statusColor = item.status === "DISCLOSED" ? H.C.GREEN : item.status === "PARTIAL" ? H.C.AMBER : H.C.RED;
    return new TableRow({ children: [
      H.cell(item.pillar, cols[0], bg, { size: 17, bold: true }),
      H.cell(item.req, cols[1], bg, { size: 17 }),
      new TableCell({ borders: H.BORDERS, shading: { fill: statusBg, type: ShadingType.CLEAR }, width: { size: cols[2], type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 80, right: 80 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: item.status, bold: true, size: 16, color: statusColor, font: "Arial" })] })]
      }),
      H.cell(item.note, cols[3], bg, { size: 16, italic: true }),
    ]});
  });
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: cols, rows: [hdr, ...rows] });
}

function workforceTable(w) {
  const stats = [
    ["Total Workers", String(w.total)], ["Local (Nepali)", String(w.local)],
    ["National", String(w.national)], ["International", String(w.international)],
    ["Female Workers", w.female_pct], ["LTI Frequency Rate", w.lti_rate],
    ["Near Misses (Q4)", String(w.near_miss)], ["Fatalities (Q4)", String(w.fatalities)],
  ];
  return new Table({
    width: { size: 9360, type: WidthType.DXA }, columnWidths: [2340, 2340, 2340, 2340],
    rows: stats.reduce((rows, _, i, a) => {
      if (i % 2 === 0) rows.push(new TableRow({ children: a.slice(i, i+2).flatMap(([l,v]) => [
        H.cell(l, 2340, H.C.NAVY, { bold: true, size: 18, color: H.C.WHITE }),
        H.cell(v, 2340, i % 4 === 0 ? H.C.LIGHT : H.C.WHITE, { bold: true, size: 20, color: H.C.NAVY }),
      ])}));
      return rows;
    }, [])
  });
}

function ifcTable(items) {
  const cols = [1000, 2600, 1400, 4360];
  const hdr = new TableRow({ tableHeader: true, children:
    ["IFC PS", "Performance Standard", "Status", "Evidence / Notes"].map((h, i) =>
      H.cell(h, cols[i], H.C.NAVY, { bold: true, size: 17, color: H.C.WHITE, center: true })
    )
  });
  const rows = items.map((item, idx) => {
    const bg = idx % 2 === 0 ? H.C.WHITE : H.C.LIGHT;
    const statusBg = item.status === "COMPLIANT" ? H.C.PASS : item.status === "PARTIAL" ? H.C.OBS : item.status === "N/A" ? H.C.NA : H.C.FAIL;
    const statusColor = item.status === "COMPLIANT" ? H.C.GREEN : item.status === "PARTIAL" ? H.C.AMBER : H.C.GREY;
    return new TableRow({ children: [
      H.cell(item.ps, cols[0], bg, { size: 17, bold: true }),
      H.cell(item.title, cols[1], bg, { size: 17 }),
      new TableCell({ borders: H.BORDERS, shading: { fill: statusBg, type: ShadingType.CLEAR }, width: { size: cols[2], type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 80, right: 80 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: item.status, bold: true, size: 16, color: statusColor, font: "Arial" })] })]
      }),
      H.cell(item.note, cols[3], bg, { size: 16, italic: true }),
    ]});
  });
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: cols, rows: [hdr, ...rows] });
}

function griTable(items) {
  const cols = [1000, 2800, 3800, 1760];
  const hdr = new TableRow({ tableHeader: true, children:
    ["GRI Code", "Disclosure Title", "Data / Information Disclosed", "Status"].map((h, i) =>
      H.cell(h, cols[i], H.C.NAVY, { bold: true, size: 17, color: H.C.WHITE, center: true })
    )
  });
  const rows = items.map((item, idx) => {
    const bg = idx % 2 === 0 ? H.C.WHITE : H.C.LIGHT;
    return new TableRow({ children: [
      H.cell(item.code, cols[0], bg, { size: 17, bold: true }),
      H.cell(item.title, cols[1], bg, { size: 17, bold: true }),
      H.cell(item.value, cols[2], bg, { size: 17 }),
      new TableCell({ borders: H.BORDERS, shading: { fill: H.C.PASS, type: ShadingType.CLEAR }, width: { size: cols[3], type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 80, right: 80 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: item.status, bold: true, size: 16, color: H.C.GREEN, font: "Arial" })] })]
      }),
    ]});
  });
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: cols, rows: [hdr, ...rows] });
}

function sdgTable(items) {
  const cols = [1000, 2400, 4560, 1400];
  const hdr = new TableRow({ tableHeader: true, children:
    ["SDG", "Goal Title", "Project Contribution", "Status"].map((h, i) =>
      H.cell(h, cols[i], H.C.NAVY, { bold: true, size: 17, color: H.C.WHITE, center: true })
    )
  });
  const rows = items.map((item, idx) => {
    const bg = idx % 2 === 0 ? H.C.WHITE : H.C.LIGHT;
    const statusBg = item.status === "ON TRACK" ? H.C.PASS : H.C.OBS;
    const statusColor = item.status === "ON TRACK" ? H.C.GREEN : H.C.AMBER;
    return new TableRow({ children: [
      H.cell(item.sdg, cols[0], bg, { size: 17, bold: true }),
      H.cell(item.title, cols[1], bg, { size: 17, bold: true }),
      H.cell(item.contribution, cols[2], bg, { size: 17 }),
      new TableCell({ borders: H.BORDERS, shading: { fill: statusBg, type: ShadingType.CLEAR }, width: { size: cols[3], type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 80, right: 80 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: item.status, bold: true, size: 15, color: statusColor, font: "Arial" })] })]
      }),
    ]});
  });
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: cols, rows: [hdr, ...rows] });
}

function improvementsTable(items) {
  const cols = [2400, 4960, 2000];
  const hdr = new TableRow({ tableHeader: true, children:
    ["ESG Area", "Required Action", "Target Deadline"].map((h, i) =>
      H.cell(h, cols[i], H.C.NAVY, { bold: true, size: 17, color: H.C.WHITE, center: true })
    )
  });
  const rows = items.map((item, idx) => new TableRow({ children: [
    H.cell(item.area, cols[0], idx % 2 === 0 ? H.C.WHITE : H.C.LIGHT, { size: 17, bold: true }),
    H.cell(item.action, cols[1], idx % 2 === 0 ? H.C.WHITE : H.C.LIGHT, { size: 17 }),
    H.cell(item.deadline, cols[2], idx % 2 === 0 ? H.C.LIGHT : H.C.WHITE, { size: 17, bold: true }),
  ]}));
  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: cols, rows: [hdr, ...rows] });
}

// ─── RUN ──────────────────────────────────────────────────────────────────────
const doc = build(R);
H.save(doc, `output/COS_ESG_Alignment_Report_${R.report_ref}.docx`)
  .catch(err => { console.error(err); process.exit(1); });
