/**
 * COS™ Report Generator — Shared Brand Helpers
 * CM Academy | Susil Bhandari, CCM
 * DOI: 10.5281/zenodo.18802971
 */

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, HeadingLevel, BorderStyle, WidthType,
  ShadingType, VerticalAlign, PageBreak, LevelFormat,
  TabStopType, SimpleField
} = require('docx');

// ─── BRAND COLOURS ─────────────────────────────────────────────────────────────
const C = {
  NAVY:   "1F3864",
  GOLD:   "C9A84C",
  ORANGE: "E07B00",
  WHITE:  "FFFFFF",
  LIGHT:  "F2F7FB",
  GREY:   "595959",
  GREEN:  "1E6B3C",
  RED:    "B00020",
  AMBER:  "E07B00",
  PASS:   "D4EDDA",
  FAIL:   "F8D7DA",
  OBS:    "FFF3CD",
  NA:     "F8F9FA",
};

const BORDER     = { style: BorderStyle.SINGLE, size: 4,  color: "CCCCCC" };
const BORDERS    = { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER };
const NO_BORDER  = { style: BorderStyle.NONE,   size: 0,  color: "FFFFFF" };
const NO_BORDERS = { top: NO_BORDER, bottom: NO_BORDER, left: NO_BORDER, right: NO_BORDER };
const THICK_BORDER = { style: BorderStyle.SINGLE, size: 8, color: C.NAVY };

// ─── PRIMITIVES ────────────────────────────────────────────────────────────────

function spacer(pts = 120) {
  return new Paragraph({ spacing: { before: pts, after: 0 }, children: [] });
}

function divider(color = C.GOLD) {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color, space: 1 } },
    spacing: { before: 80, after: 80 },
    children: []
  });
}

function h1(text, color = C.NAVY) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 280, after: 100 },
    children: [new TextRun({ text, bold: true, size: 28, color, font: "Arial" })]
  });
}

function h2(text, color = C.NAVY) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, bold: true, size: 24, color, font: "Arial" })]
  });
}

function h3(text, color = C.GREY) {
  return new Paragraph({
    spacing: { before: 160, after: 60 },
    children: [new TextRun({ text, bold: true, size: 21, color, font: "Arial" })]
  });
}

function body(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
    children: [new TextRun({
      text,
      size:    opts.size   || 20,
      bold:    opts.bold   || false,
      italics: opts.italic || false,
      color:   opts.color  || C.GREY,
      font: "Arial"
    })]
  });
}

function bullet(text, color = C.GREY, ref = "bullets") {
  return new Paragraph({
    numbering: { reference: ref, level: 0 },
    spacing: { before: 50, after: 50 },
    children: [new TextRun({ text, size: 20, color, font: "Arial" })]
  });
}

function numbered(text, color = C.GREY) {
  return bullet(text, color, "numbers");
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// ─── HEADER / FOOTER ──────────────────────────────────────────────────────────

function makeHeader() {
  return new Header({
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
                new Paragraph({ children: [new TextRun({ text: "CM ACADEMY", bold: true, size: 28, color: C.NAVY, font: "Arial" })] }),
                new Paragraph({ children: [new TextRun({ text: "Developed in Nepal. Leading the World.™", size: 16, color: C.GOLD, italic: true, font: "Arial" })] }),
              ]
            }),
            new TableCell({
              borders: NO_BORDERS,
              width: { size: 3860, type: WidthType.DXA },
              verticalAlign: VerticalAlign.CENTER,
              margins: { top: 60, bottom: 60, left: 120, right: 0 },
              children: [
                new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "COS™ METHODOLOGY", bold: true, size: 18, color: C.NAVY, font: "Arial" })] }),
                new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "Compliance · Oversight · Sustainability", size: 16, color: C.GOLD, italic: true, font: "Arial" })] }),
                new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: "DOI: 10.5281/zenodo.18802971", size: 14, color: C.GREY, font: "Arial" })] }),
              ]
            })
          ]
        })]
      }),
      new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.NAVY, space: 1 } },
        spacing: { before: 60, after: 0 },
        children: []
      })
    ]
  });
}

function makeFooter(ref, date) {
  return new Footer({
    children: [
      new Paragraph({
        border: { top: { style: BorderStyle.SINGLE, size: 6, color: C.GOLD, space: 1 } },
        spacing: { before: 80, after: 0 }, children: []
      }),
      new Paragraph({
        tabStops: [{ type: TabStopType.CENTER, position: 4680 }, { type: TabStopType.RIGHT, position: 9360 }],
        children: [
          new TextRun({ text: `${ref}  ·  ${date}`, size: 16, color: C.GREY, font: "Arial" }),
          new TextRun({ text: "\t", size: 16 }),
          new TextRun({ text: "© CM Academy | NeoPlan Consult Pvt. Ltd. | CC BY 4.0", size: 16, color: C.GREY, italic: true, font: "Arial" }),
          new TextRun({ text: "\t", size: 16 }),
          new TextRun({ text: "Page ", size: 16, color: C.GREY, font: "Arial" }),
          new SimpleField({ instruction: "PAGE", cachedValue: "1", dirty: false }),
        ]
      })
    ]
  });
}

// ─── COVER PAGE ───────────────────────────────────────────────────────────────

function makeCover(title, subtitle, color, infoRows) {
  return [
    spacer(360),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "CM ACADEMY", bold: true, size: 64, color: C.NAVY, font: "Arial" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Developed in Nepal. Leading the World.™", size: 26, color: C.GOLD, italic: true, font: "Arial" })] }),
    spacer(160),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      border: { top: { style: BorderStyle.SINGLE, size: 8, color: C.GOLD }, bottom: { style: BorderStyle.SINGLE, size: 8, color: C.GOLD } },
      spacing: { before: 120, after: 120 },
      children: [new TextRun({ text: "COS™ METHODOLOGY", bold: true, size: 42, color: C.NAVY, font: "Arial" })]
    }),
    spacer(60),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: title, bold: true, size: 34, color, font: "Arial" })] }),
    spacer(40),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: subtitle, size: 22, color: C.GREY, italic: true, font: "Arial" })] }),
    spacer(260),
    new Table({
      width: { size: 7200, type: WidthType.DXA },
      columnWidths: [2400, 4800],
      rows: infoRows.map(([label, value], i) => new TableRow({
        children: [
          new TableCell({
            borders: BORDERS,
            shading: { fill: C.NAVY, type: ShadingType.CLEAR },
            width: { size: 2400, type: WidthType.DXA },
            margins: { top: 100, bottom: 100, left: 160, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: label, bold: true, size: 20, color: C.WHITE, font: "Arial" })] })]
          }),
          new TableCell({
            borders: BORDERS,
            shading: { fill: i % 2 === 0 ? C.LIGHT : C.WHITE, type: ShadingType.CLEAR },
            width: { size: 4800, type: WidthType.DXA },
            margins: { top: 100, bottom: 100, left: 160, right: 120 },
            children: [new Paragraph({ children: [new TextRun({ text: value, size: 20, color: C.GREY, font: "Arial" })] })]
          })
        ]
      }))
    }),
    spacer(260),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Generated using COS™ Methodology — Ethics-First Construction Management", size: 17, color: C.GREY, italic: true, font: "Arial" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "Bhandari, S. (2026). COS™ Methodology. Zenodo. DOI: 10.5281/zenodo.18802971", size: 16, color: C.GREY, italic: true, font: "Arial" })] }),
    new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "© CM Academy — Complete Construction Management Developers Pvt. Ltd. (Nepal, Reg. No. 275143/078/079) | CC BY 4.0", size: 15, color: C.GREY, font: "Arial" })] }),
    pageBreak(),
  ];
}

// ─── PILLAR BANNER ────────────────────────────────────────────────────────────

function pillarBanner(letter, title, color, subtitle) {
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
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: letter, bold: true, size: 52, color: C.WHITE, font: "Arial" })] })]
        }),
        new TableCell({
          borders: NO_BORDERS,
          shading: { fill: C.LIGHT, type: ShadingType.CLEAR },
          width: { size: 8160, type: WidthType.DXA },
          verticalAlign: VerticalAlign.CENTER,
          margins: { top: 120, bottom: 120, left: 200, right: 120 },
          children: [
            new Paragraph({ children: [new TextRun({ text: title, bold: true, size: 28, color: C.NAVY, font: "Arial" })] }),
            new Paragraph({ children: [new TextRun({ text: subtitle, size: 17, color: C.GREY, italic: true, font: "Arial" })] })
          ]
        })
      ]
    })]
  });
}

// ─── CHECKLIST TABLE ──────────────────────────────────────────────────────────

function checklistTable(items) {
  const cols = [4200, 2600, 1800, 760];
  const hdr = new TableRow({
    tableHeader: true,
    children: ["Inspection Item / Verification Point", "Responsible Party", "Status", "Ref"].map((h, i) =>
      new TableCell({
        borders: BORDERS,
        shading: { fill: C.NAVY, type: ShadingType.CLEAR },
        width: { size: cols[i], type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, bold: true, size: 18, color: C.WHITE, font: "Arial" })] })]
      })
    )
  });

  const rows = items.map((item, idx) => {
    const bg = idx % 2 === 0 ? C.WHITE : C.LIGHT;
    const statusBg = item.status === "PASS" ? C.PASS : item.status === "FAIL" ? C.FAIL : item.status === "N/A" ? C.NA : C.OBS;
    const statusColor = item.status === "PASS" ? C.GREEN : item.status === "FAIL" ? C.RED : C.GREY;
    return new TableRow({
      children: [
        cell(item.item, cols[0], bg, { size: 18 }),
        cell(item.party, cols[1], bg, { size: 18 }),
        new TableCell({
          borders: BORDERS,
          shading: { fill: statusBg, type: ShadingType.CLEAR },
          width: { size: cols[2], type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: item.status || "□", bold: true, size: 18, color: statusColor, font: "Arial" })] })]
        }),
        cell(item.ref || "", cols[3], bg, { size: 16, center: true }),
      ]
    });
  });

  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: cols, rows: [hdr, ...rows] });
}

// ─── NCR TABLE ────────────────────────────────────────────────────────────────

function ncrTable(items) {
  const cols = [600, 1200, 2400, 1600, 1400, 1160, 400];
  const headers = ["No.", "NCR Ref", "Non-Conformance Description", "Location / Activity", "Responsible Party", "Due Date", "Status"];
  const hdr = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) =>
      new TableCell({
        borders: BORDERS,
        shading: { fill: C.NAVY, type: ShadingType.CLEAR },
        width: { size: cols[i], type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 80, right: 80 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, bold: true, size: 16, color: C.WHITE, font: "Arial" })] })]
      })
    )
  });

  const rows = items.map((ncr, idx) => {
    const bg = idx % 2 === 0 ? C.WHITE : C.LIGHT;
    const statusBg = ncr.status === "CLOSED" ? C.PASS : ncr.status === "OPEN" ? C.FAIL : ncr.status === "OVERDUE" ? "FFD6D6" : C.OBS;
    const statusColor = ncr.status === "CLOSED" ? C.GREEN : ncr.status === "OPEN" ? C.NAVY : ncr.status === "OVERDUE" ? C.RED : C.AMBER;
    return new TableRow({
      children: [
        cell(`${idx + 1}`, cols[0], bg, { size: 17, center: true }),
        cell(ncr.ref, cols[1], bg, { size: 17, bold: true }),
        cell(ncr.description, cols[2], bg, { size: 17 }),
        cell(ncr.location, cols[3], bg, { size: 17 }),
        cell(ncr.responsible, cols[4], bg, { size: 17 }),
        cell(ncr.due_date, cols[5], bg, { size: 17 }),
        new TableCell({
          borders: BORDERS,
          shading: { fill: statusBg, type: ShadingType.CLEAR },
          width: { size: cols[6], type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 80, right: 80 },
          children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: ncr.status, bold: true, size: 16, color: statusColor, font: "Arial" })] })]
        })
      ]
    });
  });

  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: cols, rows: [hdr, ...rows] });
}

// ─── SCORE SUMMARY TABLE ──────────────────────────────────────────────────────

function scoreSummary(c, o, s) {
  const overall = Math.round((c + o + s) / 3);
  const getColor = v => v >= 80 ? C.PASS : v >= 60 ? C.OBS : C.FAIL;
  const getLabel = v => v >= 80 ? "COMPLIANT" : v >= 60 ? "PARTIAL" : "NON-COMPLIANT";
  const getLabelColor = v => v >= 80 ? C.GREEN : v >= 60 ? C.AMBER : C.RED;

  const cols = [3000, 1200, 2000, 3160];
  const hdr = new TableRow({
    tableHeader: true,
    children: ["COS™ Pillar", "Score", "Status", "Standards Anchored"].map((h, i) =>
      new TableCell({
        borders: BORDERS, shading: { fill: C.NAVY, type: ShadingType.CLEAR },
        width: { size: cols[i], type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, bold: true, size: 18, color: C.WHITE, font: "Arial" })] })]
      })
    )
  });

  const pillars = [
    { name: "C — Compliance", score: c, standards: "CMAA · FIDIC · ISO 9001 · IFC · GCF · Local Laws" },
    { name: "O — Oversight",  score: o, standards: "Duty of Care · Audit Frameworks · Donor Protocols" },
    { name: "S — Sustainability", score: s, standards: "UN SDGs · GRI · SASB · TCFD · Carbon Systems" },
  ];

  const pillarRows = pillars.map(p => new TableRow({
    children: [
      cell(p.name, cols[0], C.WHITE, { bold: true, size: 19, color: C.NAVY }),
      new TableCell({
        borders: BORDERS, shading: { fill: getColor(p.score), type: ShadingType.CLEAR },
        width: { size: cols[1], type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `${p.score}%`, bold: true, size: 20, color: C.NAVY, font: "Arial" })] })]
      }),
      new TableCell({
        borders: BORDERS, shading: { fill: getColor(p.score), type: ShadingType.CLEAR },
        width: { size: cols[2], type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: getLabel(p.score), bold: true, size: 18, color: getLabelColor(p.score), font: "Arial" })] })]
      }),
      cell(p.standards, cols[3], C.LIGHT, { size: 17, italic: true }),
    ]
  }));

  const totalRow = new TableRow({
    children: [
      new TableCell({
        borders: BORDERS, shading: { fill: C.NAVY, type: ShadingType.CLEAR },
        width: { size: cols[0], type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [new Paragraph({ children: [new TextRun({ text: "OVERALL COS™ SCORE", bold: true, size: 20, color: C.WHITE, font: "Arial" })] })]
      }),
      new TableCell({
        borders: BORDERS, shading: { fill: getColor(overall), type: ShadingType.CLEAR },
        width: { size: cols[1], type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: `${overall}%`, bold: true, size: 24, color: C.NAVY, font: "Arial" })] })]
      }),
      new TableCell({
        borders: BORDERS, shading: { fill: getColor(overall), type: ShadingType.CLEAR },
        columnSpan: 2, width: { size: cols[2] + cols[3], type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 120, right: 120 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: getLabel(overall), bold: true, size: 22, color: getLabelColor(overall), font: "Arial" })] })]
      }),
    ]
  });

  return new Table({ width: { size: 9360, type: WidthType.DXA }, columnWidths: cols, rows: [hdr, ...pillarRows, totalRow] });
}

// ─── SIGN-OFF BLOCK ───────────────────────────────────────────────────────────

function signOff(date, clientLabel = "Client / Project Manager") {
  return [
    new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [4680, 4680],
      rows: [new TableRow({
        children: [
          new TableCell({
            borders: BORDERS, shading: { fill: C.LIGHT, type: ShadingType.CLEAR },
            width: { size: 4680, type: WidthType.DXA },
            margins: { top: 200, bottom: 200, left: 200, right: 200 },
            children: [
              new Paragraph({ children: [new TextRun({ text: "Lead Auditor / Author", bold: true, size: 20, color: C.NAVY, font: "Arial" })] }),
              spacer(100),
              new Paragraph({ children: [new TextRun({ text: "Susil Bhandari, CCM", bold: true, size: 22, color: C.NAVY, font: "Arial" })] }),
              new Paragraph({ children: [new TextRun({ text: "Certified Construction Manager (CMAA)", size: 18, color: C.GREY, font: "Arial" })] }),
              new Paragraph({ children: [new TextRun({ text: "Founder, CM Academy", size: 18, color: C.GREY, font: "Arial" })] }),
              new Paragraph({ children: [new TextRun({ text: "Director, NeoPlan Consult Pvt. Ltd.", size: 18, color: C.GREY, font: "Arial" })] }),
              new Paragraph({ children: [new TextRun({ text: "linkedin.com/in/ccm-susil-bhandari", size: 16, color: C.GOLD, font: "Arial" })] }),
              spacer(140),
              new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.NAVY } }, children: [new TextRun({ text: "Signature: ________________________________", size: 18, color: C.GREY, font: "Arial" })] }),
              spacer(60),
              new Paragraph({ children: [new TextRun({ text: `Date: ${date}`, size: 18, color: C.GREY, font: "Arial" })] }),
            ]
          }),
          new TableCell({
            borders: BORDERS, shading: { fill: C.LIGHT, type: ShadingType.CLEAR },
            width: { size: 4680, type: WidthType.DXA },
            margins: { top: 200, bottom: 200, left: 200, right: 200 },
            children: [
              new Paragraph({ children: [new TextRun({ text: clientLabel, bold: true, size: 20, color: C.NAVY, font: "Arial" })] }),
              spacer(100),
              new Paragraph({ children: [new TextRun({ text: "Name: ________________________________", size: 18, color: C.GREY, font: "Arial" })] }),
              spacer(40),
              new Paragraph({ children: [new TextRun({ text: "Designation: _________________________", size: 18, color: C.GREY, font: "Arial" })] }),
              spacer(40),
              new Paragraph({ children: [new TextRun({ text: "Organisation: ________________________", size: 18, color: C.GREY, font: "Arial" })] }),
              spacer(140),
              new Paragraph({ border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: C.NAVY } }, children: [new TextRun({ text: "Signature: ________________________________", size: 18, color: C.GREY, font: "Arial" })] }),
              spacer(60),
              new Paragraph({ children: [new TextRun({ text: "Date: ________________________________", size: 18, color: C.GREY, font: "Arial" })] }),
            ]
          })
        ]
      })]
    }),
    spacer(160),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      border: { top: { style: BorderStyle.SINGLE, size: 6, color: C.GOLD }, bottom: { style: BorderStyle.SINGLE, size: 6, color: C.GOLD } },
      spacing: { before: 120, after: 120 },
      children: [
        new TextRun({ text: "Generated using COS™ Methodology — Ethics-First Construction Management  |  ", size: 17, color: C.GREY, italic: true, font: "Arial" }),
        new TextRun({ text: "DOI: 10.5281/zenodo.18802971", size: 17, color: C.NAVY, font: "Arial" }),
      ]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: "© 2026 CM Academy – Complete Construction Management Developers Pvt. Ltd. (Nepal, Reg. No. 275143/078/079). CC BY 4.0", size: 15, color: C.GREY, font: "Arial" })]
    }),
  ];
}

// ─── DOC WRAPPER ──────────────────────────────────────────────────────────────

function makeDoc(ref, date, children) {
  return new Document({
    numbering: {
      config: [
        { reference: "bullets",
          levels: [{ level: 0, format: LevelFormat.BULLET, text: "✓", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } }, run: { color: C.GREEN, bold: true } } }] },
        { reference: "warn",
          levels: [{ level: 0, format: LevelFormat.BULLET, text: "▶", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } }, run: { color: C.AMBER, bold: true } } }] },
        { reference: "numbers",
          levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      ]
    },
    styles: {
      default: { document: { run: { font: "Arial", size: 20 } } },
      paragraphStyles: [
        { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 28, bold: true, font: "Arial", color: C.NAVY },
          paragraph: { spacing: { before: 280, after: 100 }, outlineLevel: 0 } },
        { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: 24, bold: true, font: "Arial", color: C.NAVY },
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
      headers: { default: makeHeader() },
      footers: { default: makeFooter(ref, date) },
      children
    }]
  });
}

// ─── UTILITY ──────────────────────────────────────────────────────────────────

function cell(text, width, fill = C.WHITE, opts = {}) {
  return new TableCell({
    borders: BORDERS,
    shading: { fill, type: ShadingType.CLEAR },
    width: { size: width, type: WidthType.DXA },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    children: [new Paragraph({
      alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT,
      children: [new TextRun({ text, size: opts.size || 18, bold: opts.bold || false, italics: opts.italic || false, color: opts.color || C.GREY, font: "Arial" })]
    })]
  });
}

function save(doc, path) {
  return Packer.toBuffer(doc).then(buf => {
    require('fs').writeFileSync(path, buf);
    console.log(`✅  Saved: ${path}`);
  });
}

module.exports = {
  C, BORDERS, NO_BORDERS,
  spacer, divider, h1, h2, h3, body, bullet, numbered, pageBreak,
  makeHeader, makeFooter, makeCover, pillarBanner,
  checklistTable, ncrTable, scoreSummary, signOff,
  makeDoc, cell, save
};
