/**
 * COS™ Methodology Report Generator
 * CM Academy | Susil Bhandari, CCM
 * DOI: 10.5281/zenodo.18802971
 *
 * Generates branded professional reports:
 *   - QA/QC Site Audit Report
 *   - OSH Field Audit Report
 *   - NCR Register & Closure Report
 *   - Project Governance Assessment
 *   - ESG Alignment Report
 *   - Donor Readiness Report
 */

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageBreak, LevelFormat,
  TabStopType, TabStopPosition, UnderlineType,
  SimpleField
} = require('docx');
const fs = require('fs');

// ─── BRAND COLOURS ────────────────────────────────────────────────────────────
const NAVY    = "1F3864";   // CM Academy Navy
const GOLD    = "C9A84C";   // COS™ Gold
const ORANGE  = "E07B00";   // Highlight orange
const WHITE   = "FFFFFF";
const LIGHT   = "F2F7FB";   // Light blue-grey background
const GREY    = "595959";
const GREEN   = "1E6B3C";   // Compliance green
const BORDER  = { style: BorderStyle.SINGLE, size: 4, color: "CCCCCC" };
const BORDERS = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };
const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const NO_BORDERS = { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER };

// ─── HELPER FUNCTIONS ─────────────────────────────────────────────────────────

function divider(color = GOLD) {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color, space: 1 } },
    spacing: { before: 80, after: 80 },
    children: []
  });
}

function spacer(pts = 120) {
  return new Paragraph({ spacing: { before: pts, after: 0 }, children: [] });
}

function heading1(text, color = NAVY) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 280, after: 100 },
    children: [new TextRun({ text, bold: true, size: 28, color, font: "Arial" })]
  });
}

function heading2(text, color = NAVY) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, bold: true, size: 24, color, font: "Arial" })]
  });
}

function bodyText(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [new TextRun({
      text,
      size: opts.size || 20,
      bold: opts.bold || false,
      italics: opts.italic || false,
      color: opts.color || GREY,
      font: "Arial"
    })]
  });
}

function labelValue(label, value) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    children: [
      new TextRun({ text: `${label}: `, bold: true, size: 20, color: NAVY, font: "Arial" }),
      new TextRun({ text: value || "—", size: 20, color: GREY, font: "Arial" })
    ]
  });
}

function pillHeader(letter, title, color, subtitle) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [1200, 8160],
    rows: [new TableRow({
      children: [
        new TableCell({
          borders: NO_BORDERS,
          shading: { fill: color, type: ShadingType.CLEAR },
          width: { size: 1200, type: WidthType.DXA },
          verticalAlign: VerticalAlign.CENTER,
          margins: { top: 120, bottom: 120, left: 120, right: 120 },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: letter, bold: true, size: 52, color: WHITE, font: "Arial" })]
          })]
        }),
        new TableCell({
          borders: NO_BORDERS,
          shading: { fill: LIGHT, type: ShadingType.CLEAR },
          width: { size: 8160, type: WidthType.DXA },
          verticalAlign: VerticalAlign.CENTER,
          margins: { top: 120, bottom: 120, left: 200, right: 120 },
          children: [
            new Paragraph({
              children: [new TextRun({ text: title, bold: true, size: 28, color: NAVY, font: "Arial" })]
            }),
            new Paragraph({
              children: [new TextRun({ text: subtitle, size: 18, color: GREY, italic: true, font: "Arial" })]
            })
          ]
        })
      ]
    })]
  });
}

function checklistTable(items, tableWidth = 9360) {
  const colWidths = [4500, 2500, 1680, 680];
  const headerRow = new TableRow({
    tableHeader: true,
    children: ["Inspection Item / Verification Point", "Responsible Party", "Status", "Ref"].map((h, i) =>
      new TableCell({
        borders: BORDERS,
        shading: { fill: NAVY, type: ShadingType.CLEAR },
        width: { size: colWidths[i], type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: h, bold: true, size: 18, color: WHITE, font: "Arial" })]
        })]
      })
    )
  });

  const dataRows = items.map((item, idx) =>
    new TableRow({
      children: [
        new TableCell({
          borders: BORDERS,
          shading: { fill: idx % 2 === 0 ? WHITE : LIGHT, type: ShadingType.CLEAR },
          width: { size: colWidths[0], type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({
            children: [new TextRun({ text: item.item, size: 18, color: GREY, font: "Arial" })]
          })]
        }),
        new TableCell({
          borders: BORDERS,
          shading: { fill: idx % 2 === 0 ? WHITE : LIGHT, type: ShadingType.CLEAR },
          width: { size: colWidths[1], type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({
            children: [new TextRun({ text: item.party, size: 18, color: GREY, font: "Arial" })]
          })]
        }),
        new TableCell({
          borders: BORDERS,
          shading: {
            fill: item.status === "PASS" ? "D4EDDA" : item.status === "FAIL" ? "F8D7DA" : item.status === "N/A" ? "F8F9FA" : "FFF3CD",
            type: ShadingType.CLEAR
          },
          width: { size: colWidths[2], type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({
              text: item.status || "□",
              bold: true,
              size: 18,
              color: item.status === "PASS" ? GREEN : item.status === "FAIL" ? "B00020" : GREY,
              font: "Arial"
            })]
          })]
        }),
        new TableCell({
          borders: BORDERS,
          shading: { fill: idx % 2 === 0 ? WHITE : LIGHT, type: ShadingType.CLEAR },
          width: { size: colWidths[3], type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 80, right: 80 },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: item.ref || "", size: 16, color: GREY, font: "Arial" })]
          })]
        })
      ]
    })
  );

  return new Table({
    width: { size: tableWidth, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [headerRow, ...dataRows]
  });
}

function scoreTable(c_score, o_score, s_score, overall) {
  const getColor = (score) => score >= 80 ? "D4EDDA" : score >= 60 ? "FFF3CD" : "F8D7DA";
  const getLabel = (score) => score >= 80 ? "COMPLIANT" : score >= 60 ? "PARTIAL" : "NON-COMPLIANT";

  const headerRow = new TableRow({
    tableHeader: true,
    children: ["COS™ Pillar", "Score", "Status", "Key Finding"].map((h, i) =>
      new TableCell({
        borders: BORDERS,
        shading: { fill: NAVY, type: ShadingType.CLEAR },
        width: { size: [2800, 1200, 2000, 3360][i], type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ text: h, bold: true, size: 18, color: WHITE, font: "Arial" })]
        })]
      })
    )
  });

  const pillars = [
    { name: "C — Compliance", score: c_score, finding: "CMAA, FIDIC, ISO, IFC/GCF, Local Laws" },
    { name: "O — Oversight", score: o_score, finding: "Duty of Care, Audit Frameworks, Donor Protocols" },
    { name: "S — Sustainability", score: s_score, finding: "UN SDGs, GRI, SASB, TCFD, Carbon Systems" },
  ];

  const pillarRows = pillars.map(p =>
    new TableRow({
      children: [
        new TableCell({
          borders: BORDERS, width: { size: 2800, type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: p.name, bold: true, size: 19, color: NAVY, font: "Arial" })] })]
        }),
        new TableCell({
          borders: BORDERS,
          shading: { fill: getColor(p.score), type: ShadingType.CLEAR },
          width: { size: 1200, type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: `${p.score}%`, bold: true, size: 20, color: NAVY, font: "Arial" })]
          })]
        }),
        new TableCell({
          borders: BORDERS,
          shading: { fill: getColor(p.score), type: ShadingType.CLEAR },
          width: { size: 2000, type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: getLabel(p.score), bold: true, size: 18, color: p.score >= 80 ? GREEN : p.score >= 60 ? ORANGE : "B00020", font: "Arial" })]
          })]
        }),
        new TableCell({
          borders: BORDERS, width: { size: 3360, type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: p.finding, size: 17, color: GREY, italic: true, font: "Arial" })] })]
        })
      ]
    })
  );

  const overallRow = new TableRow({
    children: [
      new TableCell({
        borders: BORDERS,
        shading: { fill: NAVY, type: ShadingType.CLEAR },
        width: { size: 2800, type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: "OVERALL COS™ SCORE", bold: true, size: 20, color: WHITE, font: "Arial" })] })]
      }),
      new TableCell({
        borders: BORDERS,
        shading: { fill: getColor(overall), type: ShadingType.CLEAR },
        width: { size: 1200, type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `${overall}%`, bold: true, size: 24, color: NAVY, font: "Arial" })] })]
      }),
      new TableCell({
        borders: BORDERS,
        shading: { fill: getColor(overall), type: ShadingType.CLEAR },
        columnSpan: 2,
        width: { size: 5360, type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: getLabel(overall), bold: true, size: 22, color: overall >= 80 ? GREEN : overall >= 60 ? ORANGE : "B00020", font: "Arial" })] })]
      })
    ]
  });

  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2800, 1200, 2000, 3360],
    rows: [headerRow, ...pillarRows, overallRow]
  });
}

// ─── REPORT DATA (Customise per client) ───────────────────────────────────────
const report = {
  // Project Information
  project_name:    "Road Infrastructure Rehabilitation Project — Phase II",
  project_code:    "NEA-2026-RD-047",
  client:          "Department of Roads, Government of Nepal",
  contractor:      "Nepal Construction Co. Pvt. Ltd.",
  consultant:      "CM Academy | NeoPlan Consult Pvt. Ltd.",
  location:        "Kathmandu–Hetauda Fast Track, Section 3",
  contract_no:     "DoR/NCB/2026/047",
  audit_date:      "09 March 2026",
  audit_ref:       "COS-QA-2026-047-001",
  auditor:         "Susil Bhandari, CCM",
  report_type:     "QA/QC SITE AUDIT REPORT",
  report_version:  "v1.0 — Issued for Review",

  // COS™ Pillar Scores (0–100)
  c_score: 82,
  o_score: 74,
  s_score: 68,
  overall: 75,

  // Executive Summary
  exec_summary: "This COS™ QA/QC Site Audit was conducted on 09 March 2026 at Section 3 of the Kathmandu–Hetauda Fast Track road project. The audit assessed project governance against all three pillars of the COS™ Methodology — Compliance, Oversight, and Sustainability. The overall project compliance stands at 75% (PARTIAL). Key strengths include ISO-aligned material certification procedures and active NCR logging. Key findings requiring immediate action: (1) incomplete survey benchmark documentation, (2) outstanding NCRs older than 14 days without closure action, and (3) absence of carbon emission tracking in the site environmental management plan.",

  // Observations
  critical_findings: [
    "3 NCRs overdue beyond 14-day closure deadline — Broker escalation required",
    "Survey benchmark records incomplete — reference points unverified for chainage 12+400 to 13+200",
    "Carbon emission log absent — ESG reporting gap for donor milestone reporting",
  ],
  commendations: [
    "100% material certificates of compliance maintained and traceable",
    "Daily QA/QC inspection register updated with no gaps in last 30 days",
    "PPE compliance observed at 94% across all active work zones",
  ],
};

// ─── BUILD DOCUMENT ────────────────────────────────────────────────────────────

function buildReport(r) {

  // ── HEADER ─────────────────────────────────────────────────────────────────
  const header = new Header({
    children: [
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [5500, 3860],
        rows: [new TableRow({
          children: [
            new TableCell({
              borders: NO_BORDERS,
              width: { size: 5500, type: WidthType.DXA },
              verticalAlign: VerticalAlign.CENTER,
              margins: { top: 60, bottom: 60, left: 0, right: 120 },
              children: [
                new Paragraph({ children: [new TextRun({ text: "CM ACADEMY", bold: true, size: 28, color: NAVY, font: "Arial" })] }),
                new Paragraph({ children: [new TextRun({ text: "Developed in Nepal. Leading the World.™", size: 16, color: GOLD, italic: true, font: "Arial" })] }),
              ]
            }),
            new TableCell({
              borders: NO_BORDERS,
              width: { size: 3860, type: WidthType.DXA },
              verticalAlign: VerticalAlign.CENTER,
              margins: { top: 60, bottom: 60, left: 120, right: 0 },
              children: [
                new Paragraph({
                  alignment: AlignmentType.RIGHT,
                  children: [new TextRun({ text: "COS™ METHODOLOGY", bold: true, size: 18, color: NAVY, font: "Arial" })]
                }),
                new Paragraph({
                  alignment: AlignmentType.RIGHT,
                  children: [new TextRun({ text: "Compliance · Oversight · Sustainability", size: 16, color: GOLD, italic: true, font: "Arial" })]
                }),
                new Paragraph({
                  alignment: AlignmentType.RIGHT,
                  children: [new TextRun({ text: "DOI: 10.5281/zenodo.18802971", size: 14, color: GREY, font: "Arial" })]
                }),
              ]
            })
          ]
        })]
      }),
      new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 1 } },
        spacing: { before: 60, after: 0 },
        children: []
      })
    ]
  });

  // ── FOOTER ─────────────────────────────────────────────────────────────────
  const footer = new Footer({
    children: [
      new Paragraph({
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: GOLD, space: 1 } },
        spacing: { before: 80, after: 0 },
        children: []
      }),
      new Paragraph({
        tabStops: [{ type: TabStopType.CENTER, position: 4680 }, { type: TabStopType.RIGHT, position: 9360 }],
        children: [
          new TextRun({ text: `${r.audit_ref}  ·  ${r.audit_date}`, size: 16, color: GREY, font: "Arial" }),
          new TextRun({ text: "\t", size: 16 }),
          new TextRun({ text: "© CM Academy | NeoPlan Consult Pvt. Ltd. | CC BY 4.0", size: 16, color: GREY, italic: true, font: "Arial" }),
          new TextRun({ text: "\t", size: 16 }),
          new TextRun({ text: "Page ", size: 16, color: GREY, font: "Arial" }),
          new SimpleField({ instruction: "PAGE", cachedValue: "1", dirty: false }),
        ]
      })
    ]
  });

  // ── COVER PAGE ─────────────────────────────────────────────────────────────
  const coverPage = [
    spacer(400),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "CM ACADEMY", bold: true, size: 64, color: NAVY, font: "Arial" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Developed in Nepal. Leading the World.™", size: 26, color: GOLD, italic: true, font: "Arial" })]
    }),
    spacer(200),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      border: {
        top: { style: BorderStyle.SINGLE, size: 8, color: GOLD },
        bottom: { style: BorderStyle.SINGLE, size: 8, color: GOLD }
      },
      spacing: { before: 120, after: 120 },
      children: [new TextRun({ text: "COS™ METHODOLOGY", bold: true, size: 42, color: NAVY, font: "Arial" })]
    }),
    spacer(60),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: r.report_type, bold: true, size: 36, color: ORANGE, font: "Arial" })]
    }),
    spacer(60),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Compliance · Oversight · Sustainability", size: 24, color: GREY, italic: true, font: "Arial" })]
    }),
    spacer(300),

    // Project info box
    new Table({
      width: { size: 7000, type: WidthType.DXA },
      columnWidths: [2400, 4600],
      rows: [
        ["Project", r.project_name],
        ["Client", r.client],
        ["Location", r.location],
        ["Audit Reference", r.audit_ref],
        ["Audit Date", r.audit_date],
        ["Auditor", r.auditor],
        ["Version", r.report_version],
      ].map(([label, value], i) => new TableRow({
        children: [
          new TableCell({
            borders: BORDERS,
            shading: { fill: NAVY, type: ShadingType.CLEAR },
            width: { size: 2400, type: WidthType.DXA },
            margins: { top: 100, bottom: 100, left: 160, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 20, color: WHITE, font: "Arial" })] })]
          }),
          new TableCell({
            borders: BORDERS,
            shading: { fill: i % 2 === 0 ? LIGHT : WHITE, type: ShadingType.CLEAR },
            width: { size: 4600, type: WidthType.DXA },
            margins: { top: 100, bottom: 100, left: 160, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: value, size: 20, color: GREY, font: "Arial" })] })]
          })
        ]
      }))
    }),

    spacer(300),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Generated using COS™ Methodology — Ethics-First Construction Management", size: 18, color: GREY, italic: true, font: "Arial" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "Bhandari, S. (2026). COS™ Methodology White Paper. Zenodo. DOI: 10.5281/zenodo.18802971", size: 16, color: GREY, italic: true, font: "Arial" })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "© CM Academy | NeoPlan Consult Pvt. Ltd. (Nepal, Reg. No. 275143/078/079) | CC BY 4.0", size: 16, color: GREY, font: "Arial" })]
    }),
    new Paragraph({ children: [new PageBreak()] }),
  ];

  // ── SECTION 1 — PROJECT OVERVIEW ───────────────────────────────────────────
  const section1 = [
    heading1("1. PROJECT & AUDIT OVERVIEW"),
    divider(),
    spacer(80),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [2200, 2480, 2200, 2480],
      rows: [
        [["Project Name", r.project_name], ["Project Code", r.project_code]],
        [["Client / Owner", r.client], ["Contract No.", r.contract_no]],
        [["Contractor", r.contractor], ["Consultant", r.consultant]],
        [["Project Location", r.location], ["Audit Date", r.audit_date]],
        [["Lead Auditor", r.auditor + ", CCM"], ["Report Version", r.report_version]],
      ].map((row, i) => new TableRow({
        children: row.flatMap(([label, value]) => [
          new TableCell({
            borders: BORDERS,
            shading: { fill: NAVY, type: ShadingType.CLEAR },
            width: { size: 2200, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 18, color: WHITE, font: "Arial" })] })]
          }),
          new TableCell({
            borders: BORDERS,
            shading: { fill: i % 2 === 0 ? LIGHT : WHITE, type: ShadingType.CLEAR },
            width: { size: 2480, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: value, size: 18, color: GREY, font: "Arial" })] })]
          })
        ])
      }))
    }),
    spacer(160),
  ];

  // ── SECTION 2 — EXECUTIVE SUMMARY ─────────────────────────────────────────
  const section2 = [
    heading1("2. EXECUTIVE SUMMARY"),
    divider(),
    spacer(80),
    bodyText(r.exec_summary, { size: 21 }),
    spacer(120),
    heading2("2.1 COS™ Tri-Pillar Compliance Score"),
    spacer(60),
    scoreTable(r.c_score, r.o_score, r.s_score, r.overall),
    spacer(120),
  ];

  // ── SECTION 3 — COMPLIANCE PILLAR ─────────────────────────────────────────
  const section3 = [
    heading1("3. COMPLIANCE PILLAR — C"),
    divider(GREEN),
    spacer(80),
    pillHeader("C", "Compliance", GREEN, "Ensures legality, transparency, and audit-readiness  |  Anchored in: CMAA · FIDIC · ISO 9001 · IFC · GCF · Local Laws"),
    spacer(120),
    heading2("3.1 QA/QC Site Inspection Checklist"),
    spacer(60),
    checklistTable([
      { item: "Material certificates of compliance (ISO/ASTM) checked and filed", party: "QA/QC Engineer", status: "PASS", ref: "ISO 9001" },
      { item: "Approved shop drawings available at site and current revision confirmed", party: "Site Engineer", status: "PASS", ref: "FIDIC" },
      { item: "Method statements approved and available for active work packages", party: "Consultant / QA", status: "PASS", ref: "CMAA" },
      { item: "NCR register maintained and all NCRs logged with date, description, and status", party: "QA/QC Engineer", status: "FAIL", ref: "ISO 9001" },
      { item: "NCR closure within 14-day contractual deadline — all open NCRs reviewed", party: "Consultant / QA", status: "FAIL", ref: "FIDIC" },
      { item: "Daily QA/QC inspection record updated with no gaps (last 30 days)", party: "QA/QC Engineer", status: "PASS", ref: "ISO 9001" },
      { item: "Workmanship inspections conducted against approved method statements", party: "Consultant / QA", status: "PASS", ref: "CMAA" },
      { item: "Test results (compaction, concrete, asphalt) logged and within specification", party: "QA/QC Engineer", status: "PASS", ref: "ASTM" },
      { item: "ITP (Inspection and Test Plan) followed for current construction phase", party: "Consultant", status: "PASS", ref: "ISO 9001" },
      { item: "As-built records updated and traceable to approved drawings", party: "Site Engineer", status: "OBS", ref: "FIDIC" },
    ]),
    spacer(120),
    heading2("3.2 Survey Verification Checklist"),
    spacer(60),
    checklistTable([
      { item: "Survey benchmarks established and reference points verified with authority", party: "Survey Engineer", status: "FAIL", ref: "DoR" },
      { item: "Alignment cross-checked with approved drawings — chainage confirmed", party: "Consultant / QA", status: "PASS", ref: "FIDIC" },
      { item: "Depth and clearance verified against design and safety codes", party: "Contractor / QA", status: "PASS", ref: "ISO" },
      { item: "Utility conflicts identified and documented before excavation", party: "Utilities Engineer", status: "N/A", ref: "Local" },
      { item: "As-built survey records updated after each construction phase", party: "Site Engineer", status: "OBS", ref: "FIDIC" },
    ]),
    spacer(120),
  ];

  // ── SECTION 4 — OVERSIGHT PILLAR ──────────────────────────────────────────
  const section4 = [
    heading1("4. OVERSIGHT PILLAR — O"),
    divider(NAVY),
    spacer(80),
    pillHeader("O", "Oversight", NAVY, "Provides ethical supervision, accountability, and resilience  |  Duty of Care · Real-Time Dashboards · Incident Reporting"),
    spacer(120),
    heading2("4.1 OSH Field Safety Checklist"),
    spacer(60),
    checklistTable([
      { item: "PPE compliance — helmets, vests, boots observed across all work zones", party: "HSE Officer", status: "PASS", ref: "OSHA" },
      { item: "Toolbox talks conducted and records maintained (minimum weekly)", party: "HSE Officer", status: "PASS", ref: "CMAA" },
      { item: "Site perimeter secured — hoarding, signage, and barriers in place", party: "Site Supervisor", status: "PASS", ref: "Local" },
      { item: "Incident register maintained — all near-misses and incidents logged", party: "HSE Officer", status: "PASS", ref: "ISO 45001" },
      { item: "Emergency response plan posted and communicated to all workers", party: "HSE Officer", status: "OBS", ref: "ISO 45001" },
      { item: "First aid kit stocked and accessible — certified first-aider present", party: "Site Supervisor", status: "PASS", ref: "Local" },
      { item: "Heavy machinery — operators certified and equipment inspected daily", party: "Plant Manager", status: "PASS", ref: "Local" },
      { item: "Excavation shoring and edge protection in place at all open trenches", party: "Contractor", status: "PASS", ref: "OSHA" },
    ]),
    spacer(120),
    heading2("4.2 Governance & Accountability Checklist"),
    spacer(60),
    checklistTable([
      { item: "Project Manager available on site minimum 5 days/week", party: "Client / PMC", status: "PASS", ref: "FIDIC" },
      { item: "Resident Engineer's daily diary maintained and countersigned", party: "Consultant", status: "PASS", ref: "FIDIC" },
      { item: "Progress reports submitted on schedule (weekly/monthly as per contract)", party: "Contractor", status: "PASS", ref: "Contract" },
      { item: "Site instruction register maintained — all verbal instructions confirmed in writing", party: "Consultant", status: "OBS", ref: "FIDIC" },
      { item: "Variation orders processed within contractual timeline", party: "QS / PMC", status: "PASS", ref: "FIDIC" },
      { item: "Sub-contractor approvals obtained before mobilisation", party: "Consultant", status: "PASS", ref: "Contract" },
    ]),
    spacer(120),
  ];

  // ── SECTION 5 — SUSTAINABILITY PILLAR ─────────────────────────────────────
  const section5 = [
    heading1("5. SUSTAINABILITY PILLAR — S"),
    divider(GOLD),
    spacer(80),
    pillHeader("S", "Sustainability", GOLD, "Aligns projects with climate and sustainability goals  |  UN SDGs · ESG · GRI · SASB · TCFD · Carbon Credit Systems"),
    spacer(120),
    heading2("5.1 Sustainability & Environmental Checklist"),
    spacer(60),
    checklistTable([
      { item: "Environmental Management Plan (EMP) approved and communicated to contractor", party: "Consultant", status: "PASS", ref: "IFC PS" },
      { item: "Waste disposal procedures documented — materials segregated on site", party: "Site Supervisor", status: "PASS", ref: "GRI" },
      { item: "Dust and noise control measures active during all construction activities", party: "Site Supervisor", status: "PASS", ref: "Local" },
      { item: "Carbon emission log maintained — GHG tracking per project phase", party: "Sustainability Officer", status: "FAIL", ref: "TCFD" },
      { item: "Reinstatement plan in place — surfaces to be restored to original or better", party: "Contractor / QA", status: "PASS", ref: "COS™" },
      { item: "Community liaison record maintained — complaints logged and resolved", party: "Project Manager", status: "OBS", ref: "IFC PS" },
      { item: "Energy efficiency measures — equipment idling controls in place", party: "Plant Manager", status: "OBS", ref: "SDG 7" },
      { item: "Water use monitored — no uncontrolled discharge to watercourses", party: "Site Supervisor", status: "PASS", ref: "IFC PS" },
      { item: "SDG alignment documented — project contribution to SDG 9, 11, 13 recorded", party: "Project Manager", status: "OBS", ref: "UN SDG" },
      { item: "ESG report submitted to donor per agreed milestone schedule", party: "Project Manager", status: "PASS", ref: "GCF" },
    ]),
    spacer(120),
  ];

  // ── SECTION 6 — FINDINGS & ACTIONS ────────────────────────────────────────
  const section6 = [
    heading1("6. KEY FINDINGS & REQUIRED ACTIONS"),
    divider(),
    spacer(80),
    heading2("6.1 Critical Findings — Immediate Action Required"),
    spacer(60),
  ];

  r.critical_findings.forEach((finding, i) => {
    section6.push(new Paragraph({
      numbering: { reference: "numbers", level: 0 },
      spacing: { before: 60, after: 60 },
      children: [new TextRun({ text: finding, size: 20, color: "B00020", bold: true, font: "Arial" })]
    }));
  });

  section6.push(
    spacer(120),
    heading2("6.2 Commendations — Good Practice Observed"),
    spacer(60),
  );

  r.commendations.forEach((item) => {
    section6.push(new Paragraph({
      numbering: { reference: "bullets", level: 0 },
      spacing: { before: 60, after: 60 },
      children: [new TextRun({ text: item, size: 20, color: GREEN, font: "Arial" })]
    }));
  });

  section6.push(
    spacer(120),
    heading2("6.3 Observations — Monitor and Resolve within 7 Days"),
    spacer(60),
    bodyText("Items marked OBS in the checklists above require corrective action within 7 days. The contractor shall submit a written response to the consultant confirming action taken. All OBS items unresolved at next audit cycle will be escalated to NCR status."),
    spacer(120),
  );

  // ── SECTION 7 — CERTIFICATION & SIGN-OFF ──────────────────────────────────
  const section7 = [
    heading1("7. COS™ CERTIFICATION & SIGN-OFF"),
    divider(GOLD),
    spacer(80),
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [4680, 4680],
      rows: [
        new TableRow({
          children: [
            new TableCell({
              borders: BORDERS,
              shading: { fill: LIGHT, type: ShadingType.CLEAR },
              width: { size: 4680, type: WidthType.DXA },
              margins: { top: 200, bottom: 200, left: 200, right: 200 },
              children: [
                new Paragraph({ children: [new TextRun({ text: "Lead Auditor", bold: true, size: 20, color: NAVY, font: "Arial" })] }),
                spacer(120),
                new Paragraph({ children: [new TextRun({ text: "Susil Bhandari, CCM", bold: true, size: 22, color: NAVY, font: "Arial" })] }),
                new Paragraph({ children: [new TextRun({ text: "Certified Construction Manager (CMAA)", size: 18, color: GREY, font: "Arial" })] }),
                new Paragraph({ children: [new TextRun({ text: "Founder, CM Academy", size: 18, color: GREY, font: "Arial" })] }),
                new Paragraph({ children: [new TextRun({ text: "Director, NeoPlan Consult Pvt. Ltd.", size: 18, color: GREY, font: "Arial" })] }),
                new Paragraph({ children: [new TextRun({ text: "linkedin.com/in/ccm-susil-bhandari", size: 16, color: GOLD, font: "Arial" })] }),
                spacer(160),
                new Paragraph({
                  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: NAVY } },
                  children: [new TextRun({ text: "Signature: ________________________________", size: 18, color: GREY, font: "Arial" })]
                }),
                spacer(60),
                new Paragraph({ children: [new TextRun({ text: `Date: ${r.audit_date}`, size: 18, color: GREY, font: "Arial" })] }),
              ]
            }),
            new TableCell({
              borders: BORDERS,
              shading: { fill: LIGHT, type: ShadingType.CLEAR },
              width: { size: 4680, type: WidthType.DXA },
              margins: { top: 200, bottom: 200, left: 200, right: 200 },
              children: [
                new Paragraph({ children: [new TextRun({ text: "Client / Project Manager", bold: true, size: 20, color: NAVY, font: "Arial" })] }),
                spacer(120),
                new Paragraph({ children: [new TextRun({ text: "Name: ________________________________", size: 18, color: GREY, font: "Arial" })] }),
                spacer(40),
                new Paragraph({ children: [new TextRun({ text: "Designation: _________________________", size: 18, color: GREY, font: "Arial" })] }),
                spacer(40),
                new Paragraph({ children: [new TextRun({ text: "Organisation: ________________________", size: 18, color: GREY, font: "Arial" })] }),
                spacer(160),
                new Paragraph({
                  border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: NAVY } },
                  children: [new TextRun({ text: "Signature: ________________________________", size: 18, color: GREY, font: "Arial" })]
                }),
                spacer(60),
                new Paragraph({ children: [new TextRun({ text: "Date: ________________________________", size: 18, color: GREY, font: "Arial" })] }),
              ]
            })
          ]
        })
      ]
    }),
    spacer(200),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      border: {
        top: { style: BorderStyle.SINGLE, size: 6, color: GOLD },
        bottom: { style: BorderStyle.SINGLE, size: 6, color: GOLD }
      },
      spacing: { before: 120, after: 120 },
      children: [
        new TextRun({ text: "Generated using COS™ Methodology — Ethics-First Construction Management  |  ", size: 17, color: GREY, italic: true, font: "Arial" }),
        new TextRun({ text: "DOI: 10.5281/zenodo.18802971", size: 17, color: NAVY, font: "Arial" }),
      ]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 80, after: 80 },
      children: [new TextRun({ text: "© 2026 CM Academy – A Division of Complete Construction Management Developers Pvt. Ltd. (Nepal, Reg. No. 275143/078/079). CC BY 4.0", size: 15, color: GREY, font: "Arial" })]
    }),
  ];

  // ── ASSEMBLE DOCUMENT ──────────────────────────────────────────────────────
  const doc = new Document({
    numbering: {
      config: [
        {
          reference: "bullets",
          levels: [{ level: 0, format: LevelFormat.BULLET, text: "✓", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } }, run: { color: GREEN, bold: true } } }]
        },
        {
          reference: "numbers",
          levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } }]
        }
      ]
    },
    styles: {
      default: { document: { run: { font: "Arial", size: 20 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 28, bold: true, font: "Arial", color: NAVY },
          paragraph: { spacing: { before: 280, after: 100 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 24, bold: true, font: "Arial", color: NAVY },
          paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 1 } },
      ]
    },
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 1440, right: 1080, bottom: 1080, left: 1080 }
        }
      },
      headers: { default: header },
      footers: { default: footer },
      children: [
        ...coverPage,
        ...section1,
        ...section2,
        ...section3,
        ...section4,
        ...section5,
        ...section6,
        ...section7,
      ]
    }]
  });

  return doc;
}

// ─── GENERATE ─────────────────────────────────────────────────────────────────
const doc = buildReport(report);
Packer.toBuffer(doc).then(buffer => {
  const outPath = `output/COS_QA_Audit_Report_${report.audit_ref}.docx`;
  fs.writeFileSync(outPath, buffer);
  console.log(`✅ Report generated: ${outPath}`);
}).catch(err => {
  console.error("Error:", err);
  process.exit(1);
});
