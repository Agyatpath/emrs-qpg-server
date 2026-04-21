const express = require('express');
const cors = require('cors');
const axios = require('axios');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
  TabStopType, TabStopPosition, PageNumber, Header,
  UnderlineType, HeadingLevel, PageBreak
} = require('docx');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

const PORT = process.env.PORT || 3001;
const OPENROUTER_KEY = process.env.OPENROUTER_KEY || '';

// ── SCHOOL INFO ───────────────────────────────────────────────────
const SCHOOL_EN = 'Eklavya Model Residential School, Bansla-Bagidora, Banswara';
const SCHOOL_HI = '\u090f\u0915\u0932\u0935\u094d\u092f \u0906\u0926\u0930\u094d\u0936 \u0906\u0935\u093e\u0938\u0940\u092f \u0935\u093f\u0926\u094d\u092f\u093e\u0932\u092f, \u092c\u093e\u0902\u0938\u0932\u093e-\u092c\u093e\u0917\u0940\u0926\u094c\u0930\u093e, \u092c\u093e\u0902\u0938\u0935\u093e\u0921\u093c\u093e';
const SCHOOL_SUB = 'NESTS \u2013 Ministry of Tribal Affairs, Govt. of India';
const MOTTO = '\u0938\u093e \u0935\u093f\u0926\u094d\u092f\u093e \u092f\u093e \u0935\u093f\u092e\u0941\u0915\u094d\u0924\u092f\u0947';

// ── HEALTH CHECK ─────────────────────────────────────────────────
app.get('/', (req, res) => {
  res.json({ status: 'ok', message: 'EMRS QPG Server running' });
});

// ── FETCH NCERT PDF ───────────────────────────────────────────────
app.post('/api/fetch-ncert', async (req, res) => {
  const { bookCode, chapterNumbers } = req.body;
  if (!bookCode || !chapterNumbers || !chapterNumbers.length) {
    return res.status(400).json({ error: 'bookCode and chapterNumbers required' });
  }

  let allText = '';
  let fetched = 0;
  let failed = 0;

  for (const chNum of chapterNumbers) {
    const padded = String(chNum).padStart(2, '0');
    const url = `https://ncert.nic.in/textbook/pdf/${bookCode}${padded}.pdf`;
    try {
      const response = await axios.get(url, {
        responseType: 'arraybuffer',
        timeout: 20000,
        headers: {
          'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36',
          'Referer': 'https://ncert.nic.in/textbook.php',
          'Accept': 'application/pdf,*/*'
        }
      });

      // Extract text from PDF buffer
      const pdfParse = require('pdf-parse');
      const data = await pdfParse(Buffer.from(response.data));
      const text = data.text.replace(/\s+/g, ' ').trim().slice(0, 8000);
      allText += `\n\n=== CHAPTER ${chNum} ===\n${text}`;
      fetched++;
    } catch (e) {
      failed++;
    }
  }

  res.json({
    success: true,
    text: allText.trim(),
    fetched,
    failed,
    totalChars: allText.length
  });
});

// ── CALL AI ───────────────────────────────────────────────────────
app.post('/api/generate', async (req, res) => {
  const { prompt } = req.body;
  if (!prompt) return res.status(400).json({ error: 'prompt required' });

  try {
    const response = await axios.post(
      'https://openrouter.ai/api/v1/chat/completions',
      {
        model: 'openrouter/auto',
        messages: [{ role: 'user', content: prompt }],
        max_tokens: 8192,
        temperature: 0.7
      },
      {
        headers: {
          'Authorization': `Bearer ${OPENROUTER_KEY}`,
          'Content-Type': 'application/json',
          'HTTP-Referer': 'https://emrs.school',
          'X-Title': 'EMRS Question Paper Generator'
        },
        timeout: 120000
      }
    );

    const text = response.data.choices[0].message.content;
    res.json({ success: true, text });
  } catch (e) {
    const msg = e.response?.data?.error?.message || e.message;
    res.status(500).json({ error: msg });
  }
});

// ── CREATE DOCX ───────────────────────────────────────────────────
app.post('/api/create-docx', async (req, res) => {
  const { paper, answerKey, blueprint, config } = req.body;
  if (!paper) return res.status(400).json({ error: 'paper required' });

  try {
    const buffer = await buildDocx(paper, answerKey || '', blueprint || [], config || {});
    const filename = [
      config.cls || 'Class',
      config.sub || 'Subject',
      config.exam || 'Exam',
      config.session || '2026-27'
    ].join('_').replace(/\s+/g, '_') + '.docx';

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');
    res.send(buffer);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── BUILD DOCX ────────────────────────────────────────────────────
async function buildDocx(paperText, answerKey, blueprint, config) {
  const stream = config.stream ? ` (${config.stream})` : '';
  const children = [];

  // ── HEADER TABLE ──
  const headerRows = [
    makeCenteredPara(SCHOOL_EN, 16, true, 'Times New Roman'),
    makeCenteredPara(SCHOOL_HI, 13, false, 'Arial'),
    makeCenteredPara(`${SCHOOL_SUB} | ${MOTTO}`, 10, false, 'Times New Roman', true),
    makeCenteredPara('─'.repeat(80), 7, false, 'Times New Roman'),
    makeCenteredPara(
      `Class: ${config.cls || ''}${stream}   |   Subject: ${config.sub || ''}   |   Exam: ${config.exam || ''}`,
      11, true, 'Times New Roman'
    ),
    makeCenteredPara(
      `Max Marks: ${config.marks || ''}   |   Duration: ${config.duration || ''}   |   ${config.setNo || 'Set A'}   |   Session: ${config.session || '2026-27'}`,
      11, false, 'Times New Roman'
    ),
  ];

  if (config.teacher || config.date) {
    let line = '';
    if (config.teacher) line += `Prepared by: ${config.teacher}`;
    if (config.teacher && config.date) line += '   |   ';
    if (config.date) line += `Date: ${config.date}`;
    headerRows.push(makeCenteredPara(line, 10, false, 'Times New Roman'));
  }

  const headerTable = new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    borders: {
      top: { style: BorderStyle.DOUBLE, size: 6, color: '000000' },
      bottom: { style: BorderStyle.DOUBLE, size: 6, color: '000000' },
      left: { style: BorderStyle.DOUBLE, size: 6, color: '000000' },
      right: { style: BorderStyle.DOUBLE, size: 6, color: '000000' },
    },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            width: { size: 9360, type: WidthType.DXA },
            margins: { top: 100, bottom: 100, left: 150, right: 150 },
            children: headerRows
          })
        ]
      })
    ]
  });

  children.push(headerTable);
  children.push(new Paragraph({ text: '', spacing: { before: 120, after: 120 } }));

  // ── PAPER CONTENT ──
  const paperLines = paperText.split('\n');
  for (const line of paperLines) {
    const trimmed = line.trim();

    // Section headers — SECTION A, GENERAL INSTRUCTIONS etc.
    if (trimmed.match(/^SECTION\s+[A-Z]/i) || trimmed.match(/^GENERAL\s+INSTRUCTIONS/i)) {
      children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 160, after: 80 },
        children: [new TextRun({
          text: trimmed,
          bold: true,
          underline: { type: UnderlineType.SINGLE },
          size: 24,
          font: 'Times New Roman'
        })]
      }));
      continue;
    }

    // Question lines — Q1. Q2. etc with marks on right
    const qMatch = trimmed.match(/^(Q\d+\.?\s*)(.*?)(\[(\d+)\])?$/);
    if (trimmed.match(/^Q\d+/)) {
      // Check if line has marks like [1] [2] [3] at end
      const marksMatch = trimmed.match(/\[(\d+)\]\s*$/);
      const marks = marksMatch ? marksMatch[1] : '';
      const qText = marksMatch ? trimmed.slice(0, trimmed.lastIndexOf('[' + marks + ']')).trim() : trimmed;

      children.push(new Paragraph({
        spacing: { before: 120, after: 40 },
        tabStops: [
          { type: TabStopType.RIGHT, position: 9360 }
        ],
        children: [
          new TextRun({
            text: qText,
            size: 22,
            font: 'Times New Roman'
          }),
          ...(marks ? [
            new TextRun({ text: '\t', size: 22 }),
            new TextRun({
              text: `[${marks}]`,
              size: 22,
              bold: true,
              font: 'Times New Roman'
            })
          ] : [])
        ]
      }));
      continue;
    }

    // OR separator
    if (trimmed === 'OR') {
      children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 80, after: 80 },
        children: [new TextRun({
          text: 'OR',
          bold: true,
          size: 22,
          font: 'Times New Roman'
        })]
      }));
      continue;
    }

    // Options (a) (b) (c) (d)
    if (trimmed.match(/^\([a-d]\)/i)) {
      children.push(new Paragraph({
        indent: { left: 360 },
        spacing: { before: 20, after: 20 },
        children: [new TextRun({
          text: trimmed,
          size: 22,
          font: 'Times New Roman'
        })]
      }));
      continue;
    }

    // Section divider ***
    if (trimmed === '***') {
      children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 160, after: 160 },
        children: [new TextRun({
          text: '* * *',
          bold: true,
          size: 24,
          font: 'Times New Roman'
        })]
      }));
      continue;
    }

    // Instruction numbered lines
    if (trimmed.match(/^\d+\.\s/)) {
      children.push(new Paragraph({
        indent: { left: 360 },
        spacing: { before: 40, after: 40 },
        children: [new TextRun({
          text: trimmed,
          size: 21,
          font: 'Times New Roman'
        })]
      }));
      continue;
    }

    // Empty line
    if (!trimmed) {
      children.push(new Paragraph({
        text: '',
        spacing: { before: 60, after: 60 }
      }));
      continue;
    }

    // Default text
    children.push(new Paragraph({
      spacing: { before: 40, after: 40 },
      children: [new TextRun({
        text: trimmed,
        size: 22,
        font: 'Times New Roman'
      })]
    }));
  }

  // ── ANSWER KEY ──
  if (answerKey) {
    children.push(new Paragraph({ children: [new PageBreak()] }));

    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 80 },
      children: [new TextRun({
        text: 'ANSWER KEY / MARKING SCHEME',
        bold: true,
        size: 28,
        font: 'Times New Roman'
      })]
    }));

    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 120 },
      children: [new TextRun({
        text: '\u0909\u0924\u094d\u0924\u0930 \u0915\u0941\u0902\u091c\u0940 / \u0905\u0902\u0915\u0928 \u092f\u094b\u091c\u0928\u093e',
        bold: true,
        size: 24,
        font: 'Arial'
      })]
    }));

    // Divider
    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 120 },
      border: { bottom: { style: BorderStyle.DOUBLE, size: 4, color: '000000', space: 1 } },
      children: [new TextRun({ text: '', size: 4 })]
    }));

    const akLines = answerKey.split('\n');
    for (const line of akLines) {
      const trimmed = line.trim();
      const isHeader = trimmed.match(/^(SECTION\s+[A-Z]|Q\d+)/i);
      children.push(new Paragraph({
        spacing: { before: isHeader ? 120 : 40, after: 40 },
        children: [new TextRun({
          text: trimmed || ' ',
          bold: !!isHeader,
          size: isHeader ? 22 : 21,
          font: 'Times New Roman'
        })]
      }));
    }
  }

  // ── BLUEPRINT ──
  if (blueprint && blueprint.length > 0) {
    children.push(new Paragraph({ children: [new PageBreak()] }));

    children.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 120 },
      children: [new TextRun({
        text: 'MARKING SCHEME BLUEPRINT',
        bold: true,
        size: 28,
        font: 'Times New Roman'
      })]
    }));

    // Blueprint table
    const bpRows = [
      // Header row
      new TableRow({
        tableHeader: true,
        children: ['Section', 'Question Type', 'Questions', 'Marks Each', 'Total'].map(h =>
          new TableCell({
            shading: { fill: 'D4A017', type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({ text: h, bold: true, size: 20, font: 'Times New Roman', color: '000000' })]
            })]
          })
        )
      }),
      // Data rows
      ...blueprint.map((b, i) => new TableRow({
        children: [
          b.sec, b.type, String(b.q), String(b.m), String(b.tot)
        ].map((val, ci) => new TableCell({
          shading: { fill: i % 2 === 0 ? 'FFFFFF' : 'F9F5E8', type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({
            alignment: ci === 0 || ci >= 2 ? AlignmentType.CENTER : AlignmentType.LEFT,
            children: [new TextRun({
              text: val,
              bold: ci === 0 || ci === 4,
              size: 20,
              font: ci === 0 ? 'Times New Roman' : 'Times New Roman',
              color: ci === 0 ? '8B6914' : '000000'
            })]
          })]
        }))
      })),
      // Total row
      new TableRow({
        children: [
          new TableCell({
            columnSpan: 4,
            shading: { fill: 'F0E8D0', type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({
              alignment: AlignmentType.RIGHT,
              children: [new TextRun({ text: 'GRAND TOTAL', bold: true, size: 20, font: 'Times New Roman' })]
            })]
          }),
          new TableCell({
            shading: { fill: 'F0E8D0', type: ShadingType.CLEAR },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun({
                text: String(blueprint.reduce((a, b) => a + b.tot, 0)),
                bold: true,
                size: 22,
                font: 'Times New Roman'
              })]
            })]
          })
        ]
      })
    ];

    children.push(new Table({
      width: { size: 9360, type: WidthType.DXA },
      columnWidths: [800, 4560, 1000, 1000, 1000],
      rows: bpRows
    }));

    children.push(new Paragraph({ text: '', spacing: { before: 120 } }));
    children.push(new Paragraph({
      spacing: { before: 0, after: 0 },
      children: [new TextRun({
        text: `* Minimum 50% Competency Based Questions (CBQ) as per CBSE NEP 2020`,
        italics: true,
        size: 18,
        font: 'Times New Roman'
      })]
    }));
    children.push(new Paragraph({
      children: [new TextRun({
        text: `* Passing Marks: 33% of ${config.marks} = ${Math.ceil((config.marks || 80) * 0.33)}`,
        italics: true,
        size: 18,
        font: 'Times New Roman'
      })]
    }));
  }

  // ── BUILD DOCUMENT ──
  const doc = new Document({
    styles: {
      default: {
        document: { run: { font: 'Times New Roman', size: 22 } }
      }
    },
    sections: [{
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 720, right: 900, bottom: 720, left: 900 }
        }
      },
      children
    }]
  });

  return await Packer.toBuffer(doc);
}

// ── HELPER: Centered paragraph ────────────────────────────────────
function makeCenteredPara(text, size, bold, font, italic) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 40, after: 40 },
    children: [new TextRun({
      text,
      bold: bold || false,
      italics: italic || false,
      size: (size || 11) * 2,
      font: font || 'Times New Roman'
    })]
  });
}

app.listen(PORT, () => {
  console.log(`EMRS QPG Server running on port ${PORT}`);
});
