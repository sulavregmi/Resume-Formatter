const express = require('express');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  LevelFormat, BorderStyle, TabStopType, UnderlineType
} = require('docx');

const app = express();
app.use(express.json({ limit: '10mb' }));

const B = "000000";
const FONT = "Calibri";
const BODY_SIZE = 20;  // 10pt
const NAME_SIZE = 28;  // 14pt

function sectionHeading(text) {
  return new Paragraph({
    children: [new TextRun({
      text,
      bold: true,
      size: BODY_SIZE,
      color: B,
      font: FONT,
      underline: { type: UnderlineType.SINGLE }
    })],
    spacing: { before: 160, after: 60 }
  });
}

function bullet(text) {
  return new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    children: [new TextRun({ text, size: BODY_SIZE, color: B, font: FONT })],
    spacing: { before: 0, after: 40 }
  });
}

function jobLine(title, company, dates) {
  return new Paragraph({
    children: [
      new TextRun({ text: title, bold: true, size: BODY_SIZE, color: B, font: FONT }),
      new TextRun({ text: " | ", size: BODY_SIZE, color: B, font: FONT }),
      new TextRun({ text: company, size: BODY_SIZE, color: B, font: FONT }),
      new TextRun({ text: "\t", size: BODY_SIZE, font: FONT }),
      new TextRun({ text: dates, size: BODY_SIZE, color: B, font: FONT })
    ],
    tabStops: [{ type: TabStopType.RIGHT, position: 9360 }],
    spacing: { before: 100, after: 40 }
  });
}

function skillRow(label, value) {
  return new Paragraph({
    children: [
      new TextRun({ text: label + ": ", bold: true, size: BODY_SIZE, color: B, font: FONT }),
      new TextRun({ text: value, size: BODY_SIZE, color: B, font: FONT })
    ],
    spacing: { before: 0, after: 40 }
  });
}

function projectTitle(title) {
  return new Paragraph({
    children: [new TextRun({ text: title, bold: true, size: BODY_SIZE, color: B, font: FONT })],
    spacing: { before: 100, after: 40 }
  });
}

function buildDoc(data) {
  const { name, contact, profile, skills, experience, education, projects } = data;

  const children = [
    // Name
    new Paragraph({
      children: [new TextRun({ text: name, bold: true, size: NAME_SIZE, color: B, font: FONT })],
      spacing: { after: 0 }
    }),

    // Contact with bottom border
    new Paragraph({
      children: [new TextRun({ text: contact, size: BODY_SIZE, color: B, font: FONT })],
      spacing: { after: 80 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: B, space: 4 } }
    }),

    // Career Profile
    sectionHeading("CAREER PROFILE"),
    new Paragraph({
      children: [new TextRun({ text: profile, size: BODY_SIZE, color: B, font: FONT })],
      spacing: { after: 80 }
    }),

    // Technical Skills
    sectionHeading("TECHNICAL SKILLS"),
    ...skills.map(s => skillRow(s.label, s.value)),

    // Employment Experience
    sectionHeading("EMPLOYMENT EXPERIENCE"),
    ...experience.flatMap(job => [
      jobLine(job.title, job.company, job.dates),
      ...job.bullets.map(b => bullet(b))
    ]),

    // Education
    sectionHeading("EDUCATION"),
    ...education.flatMap(edu => {
      const rows = [];

      // School + location line
      rows.push(new Paragraph({
        children: [new TextRun({
          text: edu.school + (edu.location ? " | " + edu.location : ""),
          bold: true, size: BODY_SIZE, color: B, font: FONT
        })],
        spacing: { before: 100, after: 0 }
      }));

      // Degree + dates flush right
      rows.push(new Paragraph({
        children: [
          new TextRun({ text: edu.degree, size: BODY_SIZE, color: B, font: FONT }),
          new TextRun({ text: "\t", size: BODY_SIZE, font: FONT }),
          new TextRun({ text: edu.dates, size: BODY_SIZE, color: B, font: FONT })
        ],
        tabStops: [{ type: TabStopType.RIGHT, position: 9360 }],
        spacing: { before: 0, after: 0 }
      }));

      if (edu.minor) {
        rows.push(new Paragraph({
          children: [new TextRun({ text: "Minor: " + edu.minor, size: BODY_SIZE, color: B, font: FONT })],
          spacing: { before: 0, after: 0 }
        }));
      }
      if (edu.coursework) {
        rows.push(new Paragraph({
          children: [new TextRun({ text: "Relevant Coursework: " + edu.coursework, size: BODY_SIZE, color: B, font: FONT })],
          spacing: { before: 0, after: 40 }
        }));
      }

      return rows;
    }),

    // Projects
    sectionHeading("ENGINEERING & ANALYTICS PROJECTS"),
    ...projects.flatMap(p => [
      projectTitle(p.title),
      ...p.bullets.map(b => bullet(b))
    ])
  ];

  return new Document({
    numbering: {
      config: [{
        reference: "bullets",
        levels: [{
          level: 0,
          format: LevelFormat.BULLET,
          text: "•",
          alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 480, hanging: 280 } } }
        }]
      }]
    },
    sections: [{
      properties: {
        page: {
          size: { width: 12240, height: 15840 },
          margin: { top: 720, right: 1080, bottom: 720, left: 1080 }
        }
      },
      children
    }]
  });
}

app.post('/generate', async (req, res) => {
  try {
    const { fileName, ...resumeData } = req.body;
    const doc = buildDoc(resumeData);
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
