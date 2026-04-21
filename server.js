const express = require('express');
const cors = require('cors');
const axios = require('axios');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
  TabStopType, UnderlineType, PageBreak, VerticalAlign
} = require('docx');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

const PORT = process.env.PORT || 3001;
const OPENROUTER_KEY = process.env.OPENROUTER_KEY || '';

// ── COLORS (from reference DOCX) ─────────────────────────────────
const C = {
  qNum:    '1C3A6B',  // Q1. Q2. dark blue bold
  qText:   '111111',  // question text near black
  marks:   '2E6DA4',  // [1 Mark] blue bold
  hindi:   '333333',  // Hindi question text
  option:  '2E6DA4',  // (a)(b)(c)(d) blue bold
  optTxt:  '000000',  // option text
  optSep:  'AAAAAA',  // / separator
  optHi:   '444444',  // Hindi option text
  answer:  '666666',  // Answer / उत्तर line
  partHdr: '2E6DA4',  // Part-I: Part-II headers
  partSub: '666666',  // (Q.1-Q.6, 1 mark each)
  secHdr:  '1C3A6B',  // section subheadings
  footer:  'AAAAAA',  // footer text
  gold:    'D4A017',  // section table background
  white:   'FFFFFF',
  black:   '000000',
};

// ── SCHOOL INFO ───────────────────────────────────────────────────
const SCHOOL_EN  = 'EKLAVYA MODEL RESIDENTIAL SCHOOL';
const SCHOOL_LOC = 'BANSLA-BAGIDORA, BANSWARA (RAJASTHAN)';
const SCHOOL_HI  = '\u090f\u0915\u0932\u0935\u094d\u092f \u0906\u0926\u0930\u094d\u0936 \u0906\u0935\u093e\u0938\u0940\u092f \u0935\u093f\u0926\u094d\u092f\u093e\u0932\u092f, \u092c\u093e\u0902\u0938\u0932\u093e-\u092c\u093e\u0917\u0940\u0926\u094c\u0930\u093e, \u092c\u093e\u0902\u0938\u0935\u093e\u0921\u093c\u093e (\u0930\u093e\u091c\u0938\u094d\u0925\u093e\u0928)';
const AFFIL      = 'Affiliated to CBSE, New Delhi  |  CBSE, \u0928\u0908 \u0926\u093f\u0932\u094d\u0932\u0940 \u0938\u0947 \u0938\u0902\u092c\u0926\u094d\u0927';

// ── HEALTH CHECK ─────────────────────────────────────────────────
app.get('/', (req, res) => res.json({ status: 'ok', message: 'EMRS QPG Server running' }));

// ── FETCH NCERT PDF ───────────────────────────────────────────────
app.post('/api/fetch-ncert', async (req, res) => {
  const { bookCode, chapterNumbers } = req.body;
  if (!bookCode || !chapterNumbers || !chapterNumbers.length)
    return res.status(400).json({ error: 'bookCode and chapterNumbers required' });
  let allText = '', fetched = 0, failed = 0;
  for (const chNum of chapterNumbers) {
    const padded = String(chNum).padStart(2, '0');
    const url = `https://ncert.nic.in/textbook/pdf/${bookCode}${padded}.pdf`;
    try {
      const response = await axios.get(url, {
        responseType: 'arraybuffer', timeout: 20000,
        headers: { 'User-Agent': 'Mozilla/5.0', 'Referer': 'https://ncert.nic.in/textbook.php' }
      });
      const pdfParse = require('pdf-parse');
      const data = await pdfParse(Buffer.from(response.data));
      allText += `\n\n=== CHAPTER ${chNum} ===\n${data.text.replace(/\s+/g,' ').trim().slice(0,8000)}`;
      fetched++;
    } catch(e) { failed++; }
  }
  res.json({ success: true, text: allText.trim(), fetched, failed, totalChars: allText.length });
});

// ── CALL AI ───────────────────────────────────────────────────────
app.post('/api/generate', async (req, res) => {
  const { prompt } = req.body;
  if (!prompt) return res.status(400).json({ error: 'prompt required' });
  try {
    const response = await axios.post('https://openrouter.ai/api/v1/chat/completions', {
      model: 'openrouter/auto',
      messages: [{ role: 'user', content: prompt }],
      max_tokens: 8192, temperature: 0.7
    }, {
      headers: {
        'Authorization': `Bearer ${OPENROUTER_KEY}`,
        'Content-Type': 'application/json',
        'HTTP-Referer': 'https://emrs.school',
        'X-Title': 'EMRS Question Paper Generator'
      },
      timeout: 120000
    });
    res.json({ success: true, text: response.data.choices[0].message.content });
  } catch(e) {
    res.status(500).json({ error: e.response?.data?.error?.message || e.message });
  }
});

// ── CREATE DOCX ───────────────────────────────────────────────────
app.post('/api/create-docx', async (req, res) => {
  const { paper, answerKey, blueprint, config } = req.body;
  if (!paper) return res.status(400).json({ error: 'paper required' });
  try {
    const buffer = await buildDocx(paper, answerKey || '', blueprint || [], config || {});
    const filename = [config.cls||'Class', config.sub||'Subject', config.exam||'Exam', config.session||'2026-27']
      .join('_').replace(/\s+/g,'_') + '.docx';
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Access-Control-Expose-Headers', 'Content-Disposition');
    res.send(buffer);
  } catch(e) { res.status(500).json({ error: e.message }); }
});

// ══════════════════════════════════════════════════════════════════
// BUILD DOCX — exact reference formatting
// ══════════════════════════════════════════════════════════════════
async function buildDocx(paperText, answerKey, blueprint, config) {
  const stream = config.stream ? ` (${config.stream})` : '';
  const cls    = (config.cls || '') + stream;
  const sub    = config.sub || '';
  const exam   = config.exam || '';
  const marks  = config.marks || 40;
  const dur    = config.duration || '1 Hour 30 Minutes';
  const setNo  = config.setNo || 'Set A';
  const sess   = config.session || '2026-27';
  const W      = 9360; // usable width in DXA (A4 with 900 margins each side)

  const children = [];

  // ── TABLE 0: School Header ────────────────────────────────────
  children.push(makeTable([[
    [
      makeRun(SCHOOL_EN,  17, true,  'Arial',  C.black),
      '\n',
      makeRun(SCHOOL_LOC, 13, true,  'Arial',  C.black),
      '\n',
      makeRun(SCHOOL_HI,  13, true,  'Mangal', C.black),
      '\n',
      makeRun(AFFIL,      9.5,false, 'Arial',  C.black, true),
    ]
  ]], W, true, AlignmentType.CENTER));

  children.push(spacer());

  // ── TABLE 1: Exam Info (3 cols) ───────────────────────────────
  const examName = exam + '  /  ' + getHindiExam(exam);
  const subHi    = getHindiSubject(sub);
  children.push(new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [3200, 3200, 2960],
    borders: allBorders(),
    rows: [new TableRow({ children: [
      new TableCell({
        width: { size: 3200, type: WidthType.DXA },
        margins: cellMargin(),
        children: [
          makePara([makeRun(examName, 11.5, true, 'Arial', C.black)], AlignmentType.LEFT),
          makePara([makeRun('Academic Session: ' + sess + '  /  \u0936\u0948\u0915\u094d\u0937\u0923\u093f\u0915 \u0938\u0924\u094d\u0930: ' + sess, 9, false, 'Arial', C.black)], AlignmentType.LEFT),
        ]
      }),
      new TableCell({
        width: { size: 3200, type: WidthType.DXA },
        margins: cellMargin(),
        verticalAlign: VerticalAlign.CENTER,
        children: [
          makePara([makeRun(sub.toUpperCase(), 12, true, 'Arial', C.black)], AlignmentType.CENTER),
          makePara([makeRun(subHi, 11, true, 'Mangal', C.black)], AlignmentType.CENTER),
        ]
      }),
      new TableCell({
        width: { size: 2960, type: WidthType.DXA },
        margins: cellMargin(),
        children: [
          makePara([makeRun('Class / \u0915\u0915\u094d\u0937\u093e: ' + cls, 10.5, true, 'Arial', C.black)], AlignmentType.LEFT),
          makePara([makeRun('Max. Marks / \u0905\u0902\u0915: ' + marks, 9.5, false, 'Arial', C.black)], AlignmentType.LEFT),
          makePara([makeRun('Time / \u0938\u092e\u092f: ' + dur, 9.5, false, 'Arial', C.black)], AlignmentType.LEFT),
          ...(config.teacher ? [makePara([makeRun('Teacher: ' + config.teacher, 9, false, 'Arial', C.black)], AlignmentType.LEFT)] : []),
          ...(config.date   ? [makePara([makeRun('Date: ' + config.date, 9, false, 'Arial', C.black)], AlignmentType.LEFT)] : []),
        ]
      }),
    ]})]
  }));

  children.push(spacer());

  // ── TABLE 2: Student Info ─────────────────────────────────────
  children.push(new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [4000, 2800, 2560],
    borders: allBorders(),
    rows: [new TableRow({ children: [
      new TableCell({ margins: cellMargin(), children: [makePara([makeRun('Name / \u0928\u093e\u092e : __________________________________', 10, false, 'Arial', C.black)])] }),
      new TableCell({ margins: cellMargin(), children: [makePara([makeRun('Roll No. / \u0915\u094d\u0930\u092e\u093e\u0902\u0915 : _____________', 10, false, 'Arial', C.black)])] }),
      new TableCell({ margins: cellMargin(), children: [makePara([makeRun('Section / \u0905\u0928\u0941\u092d\u093e\u0917 : ________', 10, false, 'Arial', C.black)])] }),
    ]})]
  }));

  children.push(spacer());

  // ── General Instructions heading ──────────────────────────────
  children.push(makePara([makeRun('GENERAL INSTRUCTIONS  /  \u0938\u093e\u092e\u093e\u0928\u094d\u092f \u0928\u093f\u0930\u094d\u0926\u0947\u0936', 12, true, 'Arial', C.qNum)], AlignmentType.CENTER));
  children.push(spacer());

  // ── TABLE 3: Instructions bilingual 2-col ─────────────────────
  const instructions = [
    ['This paper has FOUR sections: A, B, C and D.', '\u0907\u0938 \u092a\u094d\u0930\u0936\u094d\u0928-\u092a\u0924\u094d\u0930 \u092e\u0947\u0902 \u091a\u093e\u0930 \u0916\u0902\u0921 \u0939\u0948\u0902: \u0905, \u092c, \u0938 \u0914\u0930 \u0926\u0964'],
    ['All questions are compulsory. However, internal choices are given where indicated.', '\u0938\u092d\u0940 \u092a\u094d\u0930\u0936\u094d\u0928 \u0905\u0928\u093f\u0935\u093e\u0930\u094d\u092f \u0939\u0948\u0902\u0964 \u091c\u0939\u093e\u0901 \u0906\u0902\u0924\u0930\u093f\u0915 \u0935\u093f\u0915\u0932\u094d\u092a \u0926\u093f\u090f \u0917\u090f \u0939\u094b\u0902, \u0935\u0939\u093e\u0901 \u0915\u094b\u0908 \u090f\u0915 \u0909\u0924\u094d\u0924\u0930 \u0926\u0947\u0902\u0964'],
    ['Section A \u2013 Objective (MCQ / Fill in Blanks / True-False): 1 mark each.', '\u0916\u0902\u0921 \u0905 \u2013 \u0935\u0938\u094d\u0924\u0941\u0928\u093f\u0937\u094d\u0920 (\u092c\u0939\u0941\u0935\u093f\u0915\u0932\u094d\u092a\u0940\u092f / \u0930\u093f\u0915\u094d\u0924 \u0938\u094d\u0925\u093e\u0928 / \u0938\u0924\u094d\u092f-\u0905\u0938\u0924\u094d\u092f): 1-1 \u0905\u0902\u0915\u0964'],
    ['Section B \u2013 Very Short Answer: 2 marks each (2\u20133 sentences).', '\u0916\u0902\u0921 \u092c \u2013 \u0905\u0924\u093f \u0932\u0918\u0941 \u0909\u0924\u094d\u0924\u0930\u0940\u092f: 2-2 \u0905\u0902\u0915 (2\u20133 \u0935\u093e\u0915\u094d\u092f)\u0964'],
    ['Section C \u2013 Short Answer: 3 marks each (4\u20136 sentences).', '\u0916\u0902\u0921 \u0938 \u2013 \u0932\u0918\u0941 \u0909\u0924\u094d\u0924\u0930\u0940\u092f: 3-3 \u0905\u0902\u0915 (4\u20136 \u0935\u093e\u0915\u094d\u092f)\u0964'],
    ['Section D \u2013 Long Answer / Case-based / Map: as marked.', '\u0916\u0902\u0921 \u0926 \u2013 \u0926\u0940\u0930\u094d\u0918 \u0909\u0924\u094d\u0924\u0930\u0940\u092f / \u092a\u094d\u0930\u0915\u0930\u0923 \u0906\u0927\u093e\u0930\u093f\u0924 / \u092e\u093e\u0928\u091a\u093f\u0924\u094d\u0930: \u0928\u093f\u0930\u094d\u0927\u093e\u0930\u093f\u0924 \u0905\u0902\u0915\u0964'],
  ];
  const instrRows = instructions.map((pair, idx) => new TableRow({ children: [
    new TableCell({
      width: { size: 4680, type: WidthType.DXA },
      margins: cellMargin(60),
      children: [makePara([makeRun(`${idx+1}. `, 9.5, true, 'Arial', C.black), makeRun(pair[0], 9.5, false, 'Arial', C.black)])]
    }),
    new TableCell({
      width: { size: 4680, type: WidthType.DXA },
      margins: cellMargin(60),
      children: [makePara([makeRun(`${idx+1}. `, 9.5, true, 'Arial', C.black), makeRun(pair[1], 9.5, false, 'Mangal', C.black)])]
    }),
  ]}));
  children.push(new Table({ width:{size:W,type:WidthType.DXA}, columnWidths:[4680,4680], borders:allBorders(), rows:instrRows }));
  children.push(spacer());

  // ── PARSE PAPER AND BUILD SECTIONS ───────────────────────────
  const lines = paperText.split('\n');
  let i = 0;

  // Section name map for Hindi
  const SEC_HI = { A:'\u0905', B:'\u092c', C:'\u0938', D:'\u0926', E:'\u0907', F:'\u091c' };
  const SEC_NAME_HI = {
    A:'\u0935\u0938\u094d\u0924\u0941\u0928\u093f\u0937\u094d\u0920 \u092a\u094d\u0930\u0936\u094d\u0928',
    B:'\u0905\u0924\u093f \u0932\u0918\u0941 \u0909\u0924\u094d\u0924\u0930\u0940\u092f \u092a\u094d\u0930\u0936\u094d\u0928',
    C:'\u0932\u0918\u0941 \u0909\u0924\u094d\u0924\u0930\u0940\u092f \u092a\u094d\u0930\u0936\u094d\u0928',
    D:'\u0926\u0940\u0930\u094d\u0918 \u0909\u0924\u094d\u0924\u0930\u0940\u092f / \u092a\u094d\u0930\u0915\u0930\u0923 \u0906\u0927\u093e\u0930\u093f\u0924 / \u092e\u093e\u0928\u091a\u093f\u0924\u094d\u0930',
    E:'\u0905\u0924\u093f\u0930\u093f\u0915\u094d\u0924', F:'\u0905\u0924\u093f\u0930\u093f\u0915\u094d\u0924'
  };

  while (i < lines.length) {
    const raw  = lines[i];
    const line = raw.replace(/\*\*/g, '').replace(/^#+\s*/, '').trim();

    // Skip empty
    if (!line) { children.push(spacer(40)); i++; continue; }

    // ── Section header (SECTION A / SECTION – A etc.) ──────────
    const secMatch = line.match(/^SECTION[\s\-–]+([A-F])\b/i);
    if (secMatch) {
      const secLetter = secMatch[1].toUpperCase();
      // Find section type text after the letter
      const afterLetter = line.replace(/^SECTION[\s\-–]+[A-F]\s*/i, '').trim();
      const secName = afterLetter || blueprint.find(b => b.sec === secLetter)?.type || '';
      // Find marks formula from blueprint
      const bpRow = blueprint.find(b => b.sec === secLetter);
      const marksFormula = bpRow ? `${bpRow.q} \u00d7 ${bpRow.m} = ${bpRow.tot} Marks` : '';
      const secHiLetter = SEC_HI[secLetter] || secLetter;
      const secHiName   = SEC_NAME_HI[secLetter] || secName;

      // Section header table (2-col shaded)
      children.push(new Table({
        width: { size: W, type: WidthType.DXA },
        columnWidths: [7200, 2160],
        borders: allBorders(),
        rows: [new TableRow({ children: [
          new TableCell({
            width: { size: 7200, type: WidthType.DXA },
            shading: { fill: 'F5F0E0', type: ShadingType.CLEAR },
            margins: cellMargin(80),
            children: [makePara([
              makeRun(`SECTION \u2013 ${secLetter}   ${secName}`, 12, true, 'Arial', C.black),
              makeRun('   |   ', 10, false, 'Arial', C.black),
              makeRun(`\u0916\u0902\u0921 \u2013 ${secHiLetter}   ${secHiName}`, 11, true, 'Mangal', C.black),
            ])]
          }),
          new TableCell({
            width: { size: 2160, type: WidthType.DXA },
            shading: { fill: 'F5F0E0', type: ShadingType.CLEAR },
            margins: cellMargin(80),
            verticalAlign: VerticalAlign.CENTER,
            children: [makePara([makeRun(marksFormula, 10, true, 'Arial', C.black)], AlignmentType.RIGHT)]
          }),
        ]})]
      }));
      children.push(spacer(80));
      i++; continue;
    }

    // ── Part header (Part-I: MCQ etc.) ──────────────────────────
    if (line.match(/^Part[\-\s]*(I|II|III|IV|V)/i)) {
      const partMain = line.replace(/\(.*?\)/, '').trim();
      const partSub  = line.match(/\((.*?)\)/)?.[1] || '';
      children.push(makePara([
        makeRun(partMain, 10.5, true, 'Arial', C.partHdr),
        ...(partSub ? [makeRun('   (' + partSub + ')', 9.5, false, 'Arial', C.partSub)] : [])
      ]));
      children.push(spacer(40));
      i++; continue;
    }

    // ── Question line (Q1. Q2. etc.) ────────────────────────────
    const qMatch = line.match(/^(Q\d+[\.\)]\s*)(.*?)(\s*\[(\d+)\s*[Mm]arks?\s*(?:\/[^]]+)?\])?\s*$/);
    if (qMatch) {
      const qNum  = qMatch[1].trim();
      let   qText = qMatch[2].trim();
      const mText = qMatch[3] ? qMatch[3].trim() : '';

      // If no marks in square brackets, check if raw line ends with [N]
      const altMark = line.match(/\[(\d+)\]\s*$/);
      const marksDisplay = mText || (altMark ? `[${altMark[1]} Mark${altMark[1]>'1'?'s':''}]` : '');
      if (altMark && !mText) qText = line.replace(/^Q\d+[\.\)]\s*/,'').replace(/\[\d+\]\s*$/,'').trim();

      // Q number + question text + marks
      const qRuns = [
        makeRun(qNum + '  ', 11, true, 'Arial', C.qNum),
        makeRun(qText, 11, false, 'Arial', C.qText),
        ...(marksDisplay ? [makeRun('   ' + marksDisplay, 10, true, 'Arial', C.marks)] : [])
      ];
      children.push(makePara(qRuns, AlignmentType.LEFT, { before: 120, after: 20 }));

      // Look ahead for Hindi translation (next non-empty line that's Hindi or matches Hindi pattern)
      if (i+1 < lines.length) {
        const nextLine = lines[i+1].replace(/\*\*/g,'').trim();
        if (nextLine && !nextLine.match(/^Q\d+/) && !nextLine.match(/^\([a-d]\)/i) && !nextLine.match(/^SECTION/i) && isHindi(nextLine)) {
          children.push(makePara([makeRun(nextLine, 10, false, 'Mangal', C.hindi)], AlignmentType.LEFT, { before: 0, after: 20 }));
          i++;
        }
      }
      i++; continue;
    }

    // ── Options (a) (b) (c) (d) ──────────────────────────────────
    const optMatch = line.match(/^\(([a-d])\)\s*(.*)/i);
    if (optMatch) {
      const optLetter = optMatch[1].toLowerCase();
      const optText   = optMatch[2].trim();
      // Split English / Hindi if present (separated by   /   )
      const parts = optText.split(/\s{2,}\/\s{2,}|\s*\/\s*(?=[^\u0000-\u007F])/);
      const engPart = parts[0] ? parts[0].trim() : optText;
      const hiPart  = parts[1] ? parts[1].trim() : '';

      const optRuns = [
        makeRun(`(${optLetter})  `, 10, true, 'Arial', C.option),
        makeRun(engPart, 10, false, 'Arial', C.optTxt),
        ...(hiPart ? [
          makeRun('   /   ', 9, false, 'Arial', C.optSep),
          makeRun(hiPart, 9.5, false, 'Mangal', C.optHi)
        ] : [])
      ];
      children.push(makePara(optRuns, AlignmentType.LEFT, { before: 20, after: 20 }, 360));
      i++; continue;
    }

    // ── Answer blank line ─────────────────────────────────────────
    if (line.match(/^Answer\s*\/?\s*\u0909\u0924\u094d\u0924\u0930/i) || line.match(/^Answer\s*:/i)) {
      children.push(makePara([makeRun('Answer / \u0909\u0924\u094d\u0924\u0930 : ___________', 10, false, 'Arial', C.answer)], AlignmentType.LEFT, { before: 20, after: 60 }));
      i++; continue;
    }

    // ── OR separator ──────────────────────────────────────────────
    if (line === 'OR' || line.match(/^OR\s*\/\s*\u0905\u0925\u0935\u093e/i)) {
      children.push(makePara([makeRun('OR  /  \u0905\u0925\u0935\u093e', 10, true, 'Arial', C.marks)], AlignmentType.CENTER, { before: 80, after: 80 }));
      i++; continue;
    }

    // ── Case study / boxed content ────────────────────────────────
    if (line.match(/^(Read the (following|passage)|Case[\s\-]*(Study|Based)|Source[\s\-]*Based|MAP SKILL)/i)) {
      // Collect boxed content until next Q or section
      const boxLines = [];
      let j = i;
      while (j < lines.length) {
        const bl = lines[j].replace(/\*\*/g,'').trim();
        if (j > i && (bl.match(/^Q\d+[\.\)]/) || bl.match(/^SECTION[\s\-–]+[A-F]/i))) break;
        boxLines.push(bl);
        j++;
      }
      // Build box table
      const boxParas = boxLines.map(bl => {
        if (!bl) return makePara([makeRun(' ', 10, false, 'Arial', C.black)], AlignmentType.LEFT, { before: 20, after: 20 });
        const isHdr = bl.match(/^(READ|MAP SKILL|Case|Source)/i);
        return makePara([makeRun(bl, isHdr ? 10.5 : 10, isHdr, 'Arial', C.black)], AlignmentType.LEFT, { before: 20, after: 20 });
      });
      children.push(new Table({
        width: { size: W, type: WidthType.DXA },
        columnWidths: [W],
        borders: allBorders('#000000', 4),
        rows: [new TableRow({ children: [
          new TableCell({ margins: cellMargin(120), children: boxParas })
        ]})]
      }));
      children.push(spacer(80));
      i = j; continue;
    }

    // ── *** end marker ────────────────────────────────────────────
    if (line === '***' || line === '* * *') {
      children.push(makePara([makeRun('* * *', 12, true, 'Arial', C.black)], AlignmentType.CENTER, { before: 160, after: 80 }));
      i++; continue;
    }

    // ── Default paragraph ─────────────────────────────────────────
    const isHdrLine = line.match(/^(GENERAL\s+INSTRUCTIONS|Part-|Answer each)/i);
    children.push(makePara([makeRun(line, isHdrLine ? 10.5 : 10.5, !!isHdrLine, 'Arial', C.black)], AlignmentType.LEFT, { before: 40, after: 40 }));
    i++;
  }

  // ── Answer key page ───────────────────────────────────────────
  if (answerKey && answerKey.trim()) {
    children.push(new Paragraph({ children: [new PageBreak()] }));
    children.push(makePara([makeRun('ANSWER KEY / MARKING SCHEME', 14, true, 'Arial', C.qNum)], AlignmentType.CENTER, { before:0, after:80 }));
    children.push(makePara([makeRun('\u0909\u0924\u094d\u0924\u0930 \u0915\u0941\u0902\u091c\u0940 / \u0905\u0902\u0915\u0928 \u092f\u094b\u091c\u0928\u093e', 12, true, 'Mangal', C.qNum)], AlignmentType.CENTER, { before:0, after:120 }));
    for (const line of answerKey.split('\n')) {
      const t = line.replace(/\*\*/g,'').trim();
      const isH = !!t.match(/^(SECTION|Q\d+)/i);
      children.push(makePara([makeRun(t||' ', 10.5, isH, 'Arial', C.black)], AlignmentType.LEFT, { before: isH?100:30, after:30 }));
    }
  }

  // ── Marks Distribution Table ──────────────────────────────────
  children.push(spacer());
  children.push(makePara([makeRun('Marks Distribution  /  \u0905\u0902\u0915 \u0935\u093f\u0924\u0930\u0923', 11, true, 'Arial', C.qNum)], AlignmentType.CENTER, { before:120, after:80 }));

  const secLabels = blueprint.map(b => `Section ${b.sec} / \u0916\u0902\u0921 ${SEC_HI[b.sec]||b.sec}`);
  secLabels.push('Total / \u0915\u0941\u0932');
  const secVals = blueprint.map(b => String(b.tot));
  secVals.push(String(blueprint.reduce((a,b) => a + b.tot, 0)));

  const colW = Math.floor(W / secLabels.length);
  children.push(new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: secLabels.map(() => colW),
    borders: allBorders(),
    rows: [
      new TableRow({ children: secLabels.map(lbl => new TableCell({
        shading: { fill: 'F5F0E0', type: ShadingType.CLEAR },
        margins: cellMargin(60),
        children: [makePara([makeRun(lbl, 9, true, 'Arial', C.black)], AlignmentType.CENTER)]
      }))}),
      new TableRow({ children: secVals.map((v, idx) => new TableCell({
        margins: cellMargin(60),
        children: [makePara([makeRun(v, idx === secVals.length-1 ? 15 : 13, true, 'Arial', C.black)], AlignmentType.CENTER)]
      }))})
    ]
  }));

  children.push(spacer(120));
  children.push(makePara([
    makeRun('\u2014 \u2014  ', 10, false, 'Arial', C.footer),
    makeRun('Best of Luck! / \u0936\u0941\u092d\u0915\u093e\u092e\u0928\u093e\u090f\u0901!', 13, true, 'Arial', C.qNum),
    makeRun('  \u2014 \u2014', 10, false, 'Arial', C.footer),
  ], AlignmentType.CENTER, { before:80, after:40 }));
  children.push(makePara([makeRun('Eklavya Model Residential School, Bansla-Bagidora (Banswara)', 8, false, 'Arial', C.footer)], AlignmentType.CENTER));

  // ── Assemble Document ─────────────────────────────────────────
  const doc = new Document({
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

// ══════════════════════════════════════════════════════════════════
// HELPER FUNCTIONS
// ══════════════════════════════════════════════════════════════════

function makeRun(text, size, bold, font, color, italic) {
  return new TextRun({
    text: String(text),
    bold: !!bold,
    italics: !!italic,
    size: Math.round((size||11) * 2),
    font: font || 'Arial',
    color: color || '000000'
  });
}

function makePara(runs, align, spacing, indent) {
  return new Paragraph({
    alignment: align || AlignmentType.LEFT,
    spacing: spacing || { before: 40, after: 40 },
    indent: indent ? { left: indent } : undefined,
    children: runs.flatMap(r => typeof r === 'string' ? [new TextRun({ text: r, break: r==='\n'?1:0 })] : [r])
  });
}

function spacer(size) {
  return new Paragraph({ text: '', spacing: { before: size||60, after: size||60 } });
}

function cellMargin(v) {
  const m = v || 80;
  return { top: m, bottom: m, left: m+20, right: m+20 };
}

function allBorders(color, size) {
  const b = { style: BorderStyle.SINGLE, size: size||4, color: color||'000000' };
  return { top:b, bottom:b, left:b, right:b, insideHorizontal:b, insideVertical:b };
}

function makeTable(cells, width, centered, align) {
  return new Table({
    width: { size: width, type: WidthType.DXA },
    columnWidths: [width],
    borders: allBorders(),
    rows: [new TableRow({ children: [
      new TableCell({
        margins: cellMargin(100),
        children: cells[0].map(content => {
          if (Array.isArray(content)) {
            return makePara(content.filter(r => typeof r !== 'string' || r !== '\n'), align || AlignmentType.CENTER);
          }
          return content;
        })
      })
    ]})]
  });
}

function isHindi(text) {
  return /[\u0900-\u097F]/.test(text);
}

function getHindiExam(exam) {
  const map = {
    'Unit Test 1': '\u0907\u0915\u093e\u0908 \u092a\u0930\u0940\u0915\u094d\u0937\u0923 \u2013 I',
    'Unit Test 2': '\u0907\u0915\u093e\u0908 \u092a\u0930\u0940\u0915\u094d\u0937\u0923 \u2013 II',
    'Unit Test 3': '\u0907\u0915\u093e\u0908 \u092a\u0930\u0940\u0915\u094d\u0937\u0923 \u2013 III',
    'Unit Test 4': '\u0907\u0915\u093e\u0908 \u092a\u0930\u0940\u0915\u094d\u0937\u0923 \u2013 IV',
    'Half-Yearly': '\u0905\u0930\u094d\u0927\u0935\u093e\u0930\u094d\u0937\u093f\u0915 \u092a\u0930\u0940\u0915\u094d\u0937\u093e',
    'Annual': '\u0935\u093e\u0930\u094d\u0937\u093f\u0915 \u092a\u0930\u0940\u0915\u094d\u0937\u093e',
    'Annual (Board)': '\u0935\u093e\u0930\u094d\u0937\u093f\u0915 \u092a\u0930\u0940\u0915\u094d\u0937\u093e (\u092c\u094b\u0930\u094d\u0921)',
  };
  return map[exam] || exam;
}

function getHindiSubject(sub) {
  const map = {
    'Mathematics': '\u0917\u0923\u093f\u0924', 'Science': '\u0935\u093f\u091c\u094d\u091e\u093e\u0928',
    'Social Science': '\u0938\u093e\u092e\u093e\u091c\u093f\u0915 \u0935\u093f\u091c\u094d\u091e\u093e\u0928',
    'Hindi': '\u0939\u093f\u0902\u0926\u0940', 'English': '\u0905\u0902\u0917\u094d\u0930\u0947\u091c\u093c\u0940',
    'Sanskrit': '\u0938\u0902\u0938\u094d\u0915\u0943\u0924', 'Physics': '\u092d\u094c\u0924\u093f\u0915\u0940',
    'Chemistry': '\u0930\u0938\u093e\u092f\u0928 \u0935\u093f\u091c\u094d\u091e\u093e\u0928',
    'Biology': '\u091c\u0940\u0935 \u0935\u093f\u091c\u094d\u091e\u093e\u0928',
    'History': '\u0907\u0924\u093f\u0939\u093e\u0938', 'Geography': '\u092d\u0942\u0917\u094b\u0932',
    'Economics': '\u0905\u0930\u094d\u0925\u0936\u093e\u0938\u094d\u0924\u094d\u0930',
    'Political Science': '\u0930\u093e\u091c\u0928\u0940\u0924\u093f \u0935\u093f\u091c\u094d\u091e\u093e\u0928',
    'Hindi Core': '\u0939\u093f\u0902\u0926\u0940 (\u0906\u0927\u093e\u0930)',
    'English Core': '\u0905\u0902\u0917\u094d\u0930\u0947\u091c\u093c\u0940 (\u0906\u0927\u093e\u0930)',
    'Business Studies': '\u0935\u094d\u092f\u0935\u0938\u093e\u092f \u0905\u0927\u094d\u092f\u092f\u0928',
    'Accountancy': '\u0932\u0947\u0916\u093e\u0936\u093e\u0938\u094d\u0924\u094d\u0930',
    'Physical Education': '\u0936\u093e\u0930\u0940\u0930\u093f\u0915 \u0936\u093f\u0915\u094d\u0937\u093e',
  };
  return map[sub] || sub;
}

app.listen(PORT, () => console.log(`EMRS QPG Server running on port ${PORT}`));
