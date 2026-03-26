const { Document, Packer, Paragraph, TextRun, Table, TableRow,
        TableCell, HeadingLevel, AlignmentType, BorderStyle,
        WidthType, ShadingType, LevelFormat, PageBreak,
        Header, Footer, SimpleField } = require('docx');
const fs = require('fs');

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };

function para(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text, size: 22, font: "Arial", ...opts })]
  });
}

function spacer() {
  return new Paragraph({ children: [new TextRun("")] });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, size: 22, font: "Arial" })]
  });
}

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 120 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "2E75B6", space: 4 } },
    children: [new TextRun({ text, bold: true, size: 28, color: "1B3A6B", font: "Arial" })]
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 80 },
    children: [new TextRun({ text, bold: true, size: 24, color: "2E75B6", font: "Arial" })]
  });
}

function riskColor(level) {
  return level === "HIGH" ? "C00000" : level === "MEDIUM" ? "D06000" : "375623";
}

function riskBg(level) {
  return level === "HIGH" ? "FCE4E4" : level === "MEDIUM" ? "FFF3CD" : "E9F7EF";
}

function findingTable(num, title, risk, cvss, description, impact, remediation) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2200, 7160],
    rows: [
      new TableRow({ children: [
        new TableCell({
          columnSpan: 2, borders,
          shading: { fill: "1B3A6B", type: ShadingType.CLEAR },
          margins: { top: 100, bottom: 100, left: 150, right: 150 },
          children: [new Paragraph({ children: [
            new TextRun({ text: `Finding ${num}: ${title}   `, bold: true, size: 24, color: "FFFFFF", font: "Arial" }),
            new TextRun({ text: `[${risk}]`, bold: true, size: 24, color: risk === "HIGH" ? "FFB3B3" : risk === "MEDIUM" ? "FFE082" : "A8E6A3", font: "Arial" })
          ]})]
        })
      ]}),
      ...[ 
        ["Risk Level",      risk,        true],
        ["CVSS Score",      cvss,        false],
        ["Description",     description, false],
        ["Business Impact", impact,      false],
        ["Remediation",     remediation, false],
      ].map(([label, value, isRisk]) => new TableRow({ children: [
        new TableCell({
          borders, width: { size: 2200, type: WidthType.DXA },
          shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [
            new TextRun({ text: label, bold: true, size: 22, font: "Arial" })
          ]})]
        }),
        new TableCell({
          borders, width: { size: 7160, type: WidthType.DXA },
          shading: { fill: isRisk ? riskBg(risk) : "FFFFFF", type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [
            new TextRun({ text: value, size: 22, font: "Arial", bold: isRisk, color: isRisk ? riskColor(risk) : "000000" })
          ]})]
        }),
      ]}))
    ]
  });
}

const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0, format: LevelFormat.BULLET, text: "\u2022",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } }
      }]
    }]
  },
  styles: {
    default: { document: { run: { font: "Arial", size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal",
        run: { size: 28, bold: true, font: "Arial", color: "1B3A6B" },
        paragraph: { spacing: { before: 360, after: 120 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal",
        run: { size: 24, bold: true, font: "Arial", color: "2E75B6" },
        paragraph: { spacing: { before: 240, after: 80 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 }
      }
    },
    headers: {
      default: new Header({ children: [
        new Paragraph({
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "2E75B6", space: 4 } },
          spacing: { before: 0, after: 120 },
          children: [
            new TextRun({ text: "VULNERABILITY ASSESSMENT REPORT", bold: true, size: 20, color: "1B3A6B", font: "Arial" }),
            new TextRun({ text: "   |   testphp.vulnweb.com   |   Confidential", size: 18, color: "888888", font: "Arial" }),
          ]
        })
      ]})
    },
    footers: {
      default: new Footer({ children: [
        new Paragraph({
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: "2E75B6", space: 4 } },
          alignment: AlignmentType.CENTER,
          spacing: { before: 120, after: 0 },
          children: [
            new TextRun({ text: "Prepared by: Malebo Nkuna   |   Future Interns Cybersecurity Internship   |   Page ", size: 18, color: "888888", font: "Arial" }),
            new SimpleField({ instruction: "PAGE", cachedValue: "1" }),
          ]
        })
      ]})
    },
    children: [

      // ── COVER PAGE ──
      spacer(), spacer(), spacer(),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 480, after: 80 },
        children: [new TextRun({ text: "VULNERABILITY ASSESSMENT REPORT", bold: true, size: 52, color: "1B3A6B", font: "Arial" })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 80 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 6 } },
        children: [new TextRun({ text: "Target: testphp.vulnweb.com", size: 28, color: "2E75B6", font: "Arial" })]
      }),
      spacer(),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3000, 6360],
        rows: [
          ["Prepared By",    "Malebo Nkuna"],
          ["CIN ID",         "FIT/MAR26/CS7289"],
          ["Program",        "Future Interns - Cyber Security Internship"],
          ["Date",           "March 26, 2026"],
          ["Classification", "Confidential"],
          ["Task",           "Task 1 - FUTURE_CS_01"],
        ].map(([label, value]) => new TableRow({ children: [
          new TableCell({
            borders, width: { size: 3000, type: WidthType.DXA },
            shading: { fill: "1B3A6B", type: ShadingType.CLEAR },
            margins: { top: 100, bottom: 100, left: 150, right: 150 },
            children: [new Paragraph({ children: [
              new TextRun({ text: label, bold: true, size: 22, color: "FFFFFF", font: "Arial" })
            ]})]
          }),
          new TableCell({
            borders, width: { size: 6360, type: WidthType.DXA },
            margins: { top: 100, bottom: 100, left: 150, right: 150 },
            children: [new Paragraph({ children: [
              new TextRun({ text: value, size: 22, font: "Arial", bold: label === "Classification", color: label === "Classification" ? "C00000" : "000000" })
            ]})]
          }),
        ]}))
      }),
      new Paragraph({ children: [new PageBreak()] }),

      // ── SECTION 1 ──
      heading1("1. Executive Summary"),
      para("This report presents the findings of a vulnerability assessment conducted on testphp.vulnweb.com — a deliberately vulnerable demo application maintained by Acunetix, designed for authorized security testing and educational purposes."),
      spacer(),
      para("A total of 3 vulnerabilities were identified spanning High and Medium severity ratings. All findings are documented with business impact explanations and actionable remediation steps."),
      spacer(),
      heading2("1.1 Risk Summary"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [1800, 5760, 1800],
        rows: [
          new TableRow({ children: [
            ...["Severity", "Finding", "CVSS"].map(h => new TableCell({
              borders,
              shading: { fill: "1B3A6B", type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              width: { size: h === "Finding" ? 5760 : 1800, type: WidthType.DXA },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
                new TextRun({ text: h, bold: true, size: 22, color: "FFFFFF", font: "Arial" })
              ]})]
            }))
          ]}),
          ...[ 
            ["HIGH",   "SQL Injection (SQLi)",        "9.8"],
            ["HIGH",   "Cross-Site Scripting (XSS)",  "8.2"],
            ["MEDIUM", "Missing HTTP Security Headers","5.1"],
          ].map(([risk, finding, cvss]) => new TableRow({ children: [
            new TableCell({
              borders, width: { size: 1800, type: WidthType.DXA },
              shading: { fill: riskBg(risk), type: ShadingType.CLEAR },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
                new TextRun({ text: risk, bold: true, size: 22, color: riskColor(risk), font: "Arial" })
              ]})]
            }),
            new TableCell({
              borders, width: { size: 5760, type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ children: [
                new TextRun({ text: finding, size: 22, font: "Arial" })
              ]})]
            }),
            new TableCell({
              borders, width: { size: 1800, type: WidthType.DXA },
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
                new TextRun({ text: cvss, size: 22, font: "Arial" })
              ]})]
            }),
          ]}))
        ]
      }),
      spacer(),

      // ── SECTION 2 ──
      new Paragraph({ children: [new PageBreak()] }),
      heading1("2. Scope & Methodology"),
      heading2("2.1 Target Information"),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3000, 6360],
        rows: [
          ["Target URL",       "http://testphp.vulnweb.com"],
          ["Technology Stack", "PHP, MySQL, Apache HTTP Server"],
          ["Assessment Type",  "Black-box Vulnerability Assessment"],
          ["Authorization",    "Authorized demo target (Acunetix test site)"],
          ["Assessment Date",  "March 26, 2026"],
        ].map(([k, v]) => new TableRow({ children: [
          new TableCell({
            borders, width: { size: 3000, type: WidthType.DXA },
            shading: { fill: "D5E8F0", type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({ children: [
              new TextRun({ text: k, bold: true, size: 22, font: "Arial" })
            ]})]
          }),
          new TableCell({
            borders, width: { size: 6360, type: WidthType.DXA },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({ children: [
              new TextRun({ text: v, size: 22, font: "Arial" })
            ]})]
          }),
        ]}))
      }),
      spacer(),
      heading2("2.2 Tools Used"),
      bullet("Nmap — Network port scanning and service detection"),
      bullet("OWASP ZAP (Passive Mode) — Automated vulnerability scanning"),
      bullet("Browser DevTools (Chrome) — Manual inspection of headers and cookies"),
      spacer(),

      // ── SECTION 3 ──
      new Paragraph({ children: [new PageBreak()] }),
      heading1("3. Detailed Findings"),
      spacer(),
      findingTable(1, "SQL Injection (SQLi)", "HIGH", "9.8 (Critical)",
        "SQL Injection was identified in multiple URL parameters. User input is appended directly into database queries without sanitisation.",
        "An attacker could dump the entire database, bypass authentication, or delete records — causing a severe data breach.",
        "Replace all dynamic SQL queries with parameterised queries. Apply least privilege to database accounts. Deploy a Web Application Firewall."
      ),
      spacer(),
      findingTable(2, "Cross-Site Scripting (XSS)", "HIGH", "8.2 (High)",
        "Reflected XSS was found in the search functionality. User input is echoed back to the page without output encoding.",
        "An attacker can steal session cookies, redirect users to fake login pages, or perform actions on behalf of the victim.",
        "Implement strict output encoding. Adopt a Content Security Policy (CSP) header. Validate all input on both client and server sides."
      ),
      spacer(),
      findingTable(3, "Missing HTTP Security Headers", "MEDIUM", "5.1 (Medium)",
        "Several important security headers are absent including X-Frame-Options, Content-Security-Policy, and Strict-Transport-Security.",
        "Missing headers expose the app to clickjacking, XSS, and man-in-the-middle attacks.",
        "Add security headers to all HTTP responses via the web server. Use Mozilla Observatory to validate after implementation."
      ),
      spacer(),

      // ── SECTION 4 ──
      new Paragraph({ children: [new PageBreak()] }),
      heading1("4. Conclusion"),
      para("The vulnerability assessment of testphp.vulnweb.com revealed significant security weaknesses. The two High-severity findings must be addressed as a first priority. All remediation steps are straightforward and well-documented."),
      spacer(),
      heading2("4.1 Top Recommendations"),
      bullet("Fix SQL Injection immediately — use parameterised queries in all database calls."),
      bullet("Fix XSS — implement output encoding and a Content Security Policy."),
      bullet("Add security headers — a server configuration change requiring less than one hour."),
      bullet("Schedule regular security assessments — at least quarterly and after major releases."),
      spacer(),
      para("This report was produced as part of the Future Interns Cybersecurity Internship Programme (Task 1 — FUTURE_CS_01). All testing was conducted on a legally authorised demo target.", { italic: true, color: "666666" }),

    ]
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync("my_report.docx", buffer);
  console.log("Done! Open my_report.docx to see your report!");
});