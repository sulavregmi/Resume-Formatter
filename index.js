const express = require('express');
const { Document, Packer, Paragraph, TextRun, LevelFormat,
        AlignmentType, BorderStyle, TabStopType } = require('docx');

const app = express();
app.use(express.json({ limit: '10mb' }));

const BLACK = "000000";

function sectionHeading(text) {
  return new Paragraph({
    children: [new TextRun({ text: text.toUpperCase(), bold: true, size: 22, color: BLACK, font: "Calibri" })],
    spacing: { before: 240, after: 80 }
  });
}
function divider() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: BLACK, space: 4 } },
    spacing: { after: 100 }
  });
}
function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children: [new TextRun({ text, size: 20, color: BLACK, font: "Calibri" })],
    spacing: { after: 60 }
  });
}
function jobHeader(title, company, dates) {
  return new Paragraph({
    children: [
      new TextRun({ text: title, bold: true, size: 22, color: BLACK, font: "Calibri" }),
      new TextRun({ text: "  |  ", size: 20, color: BLACK, font: "Calibri" }),
      new TextRun({ text: company, italics: true, size: 20, color: BLACK, font: "Calibri" }),
      new TextRun({ text: "\t", size: 20, font: "Calibri" }),
      new TextRun({ text: dates, size: 20, color: BLACK, font: "Calibri" })
    ],
    tabStops: [{ type: TabStopType.RIGHT, position: 9360 }],
    spacing: { before: 160, after: 60 }
  });
}
function skillRow(label, value) {
  return new Paragraph({
    children: [
      new TextRun({ text: label + ":  ", bold: true, size: 20, color: BLACK, font: "Calibri" }),
      new TextRun({ text: value, size: 20, color: BLACK, font: "Calibri" })
    ],
    spacing: { after: 80 }
  });
}
function projectTitle(title) {
  return new Paragraph({
    children: [new TextRun({ text: title, bold: true, size: 20, color: BLACK, font: "Calibri" })],
    spacing: { before: 140, after: 60 }
  });
}
function eduRow(degree, dates) {
  return new Paragraph({
    children: [
      new TextRun({ text: degree, bold: true, size: 20, color: BLACK, font: "Calibri" }),
      new TextRun({ text: "\t", size: 20, font: "Calibri" }),
      new TextRun({ text: dates, size: 20, color: BLACK, font: "Calibri" })
    ],
    tabStops: [{ type: TabStopType.RIGHT, position: 9360 }],
    spacing: { before: 140, after: 40 }
  });
}
function subLine(text) {
  return new Paragraph({
    children: [new TextRun({ text, italics: true, size: 20, color: BLACK, font: "Calibri" })],
    spacing: { after: 40 }
  });
}

app.post('/generate', async (req, res) => {
  try {
    const { name, contact, profile, skills, experience, education, projects, fileName } = req.body;

    const children = [
      // Name
      new Paragraph({
        children: [new TextRun({ text: name.toUpperCase(), bold: true, size: 52, color: BLACK, font: "Calibri" })],
        alignment: AlignmentType.CENTER, spacing: { after: 60 }
      }),
      // Contact
      new Paragraph({
        children: [new TextRun({ text: contact, size: 20, color: BLACK, font: "Calibri" })],
        alignment: AlignmentType.CENTER, spacing: { after: 160 }
      }),
      // Divider under header
      new Paragraph({
        border: { bottom: { style: BorderStyle.SINGLE, size: 16, color: BLACK, space: 1 } },
        spacing: { after: 200 }
      }),

      // Profile
      sectionHeading("Career Profile"),
      divider(),
      new Paragraph({
        children: [new TextRun({ text: profile, size: 20, color: BLACK, font: "Calibri" })],
        spacing: { after: 200 }
      }),

      // Skills
      sectionHeading("Technical Skills"),
      divider(),
      ...skills.map(s => skillRow(s.label, s.value)),

      // Experience
      new Paragraph({ spacing: { after: 0 } }),
      sectionHeading("Employment Experience"),
      divider(),
      ...experience.flatMap(job => [
        jobHeader(job.title, job.company, job.dates),
        ...job.bullets.map(b => bullet(b))
      ]),

      // Education
      new Paragraph({ spacing: { after: 0 } }),
      sectionHeading("Education"),
      divider(),
      ...education.flatMap(edu => [
        eduRow(edu.degree, edu.dates),
        ...(edu.sub ? [subLine(edu.sub)] : [])
      ]),

      // Projects
      new Paragraph({ spacing: { after: 0 } }),
      sectionHeading("Engineering & Analytics Projects"),
      divider(),
      ...projects.flatMap(p => [
        projectTitle(p.title),
        ...p.bullets.map(b => bullet(b))
      ])
    ];

    const doc = new Document({
      numbering: {
        config: [{
          reference: "bullets",
          levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 480, hanging: 280 } } } }]
        }]
      },
      sections: [{
        properties: {
          page: {
            size: { width: 12240, height: 15840 },
            margin: { top: 900, right: 1080, bottom: 900, left: 1080 }
          }
        },
        children
      }]
    });

    const buffer = await Packer.toBuffer(doc);
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename="${fileName || 'resume.docx'}"`,
      'Content-Length': buffer.length
    });
    res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

app.get('/', (req, res) => res.send('docx-service running'));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Listening on ${PORT}`));
