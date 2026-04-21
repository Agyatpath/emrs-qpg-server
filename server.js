const express = require('express');
const cors = require('cors');
const axios = require('axios');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
  PageBreak, VerticalAlign, ImageRun
} = require('docx');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));
const PORT = process.env.PORT || 3001;

// ── KEYS (from environment — never hardcode) ─────────────────────
const ANTHROPIC_KEY  = process.env.ANTHROPIC_KEY  || '';
const OPENROUTER_KEY = process.env.OPENROUTER_KEY || '';

// ── COLORS ───────────────────────────────────────────────────────
const C = {
  qNum:'1C3A6B', qText:'111111', marks:'2E6DA4',
  hindi:'333333', option:'2E6DA4', optTxt:'000000',
  optSep:'AAAAAA', optHi:'444444', answer:'666666',
  black:'000000', footer:'AAAAAA', shade:'F5F0E0'
};

// ── SCHOOL ────────────────────────────────────────────────────────
const SCH_EN  = 'EKLAVYA MODEL RESIDENTIAL SCHOOL';
const SCH_LOC = 'BANSLA-BAGIDORA, BANSWARA (RAJASTHAN)';
const SCH_HI  = '\u090f\u0915\u0932\u0935\u094d\u092f \u0906\u0926\u0930\u094d\u0936 \u0906\u0935\u093e\u0938\u0940\u092f \u0935\u093f\u0926\u094d\u092f\u093e\u0932\u092f, \u092c\u093e\u0902\u0938\u0932\u093e-\u092c\u093e\u0917\u0940\u0926\u094c\u0930\u093e, \u092c\u093e\u0902\u0938\u0935\u093e\u0921\u093c\u093e (\u0930\u093e\u091c\u0938\u094d\u0925\u093e\u0928)';
const SCH_AFF = 'Affiliated to CBSE, New Delhi | CBSE, \u0928\u0908 \u0926\u093f\u0932\u094d\u0932\u0940 \u0938\u0947 \u0938\u0902\u092c\u0926\u094d\u0927 | NESTS, Ministry of Tribal Affairs, Govt. of India';

// ── SEC HINDI NAMES ───────────────────────────────────────────────
const SEC_HI = {A:'\u0905',B:'\u092c',C:'\u0938',D:'\u0926',E:'\u0907',F:'\u091c'};
const SEC_NAME_HI = {
  A:'\u0935\u0938\u094d\u0924\u0941\u0928\u093f\u0937\u094d\u0920 \u092a\u094d\u0930\u0936\u094d\u0928',
  B:'\u0905\u0924\u093f \u0932\u0918\u0941 \u0909\u0924\u094d\u0924\u0930\u0940\u092f \u092a\u094d\u0930\u0936\u094d\u0928',
  C:'\u0932\u0918\u0941 \u0909\u0924\u094d\u0924\u0930\u0940\u092f \u092a\u094d\u0930\u0936\u094d\u0928',
  D:'\u0926\u0940\u0930\u094d\u0918 \u0909\u0924\u094d\u0924\u0930\u0940\u092f \u092a\u094d\u0930\u0936\u094d\u0928',
  E:'\u092a\u094d\u0930\u0915\u0930\u0923 \u0906\u0927\u093e\u0930\u093f\u0924 \u092a\u094d\u0930\u0936\u094d\u0928',
  F:'\u092e\u093e\u0928\u091a\u093f\u0924\u094d\u0930 \u0915\u094c\u0936\u0932'
};

// ── NCERT BOOK CODES ──────────────────────────────────────────────
const NCERT_BOOKS = {
  '6_Mathematics':   {code:'fmms1', name:'Ganita Prakash'},
  '6_Science':       {code:'fesc1', name:'Curiosity'},
  '6_Social Science':{code:'fees1', name:'Exploring Society'},
  '6_Hindi':         {code:'fhvs1', name:'Vasant'},
  '6_English':       {code:'feas1', name:'Poorvi'},
  '6_Sanskrit':      {code:'fsans1',name:'Ruchira'},
  '7_Mathematics':   {code:'gema1', name:'Mathematics'},
  '7_Science':       {code:'gesc1', name:'Science'},
  '7_Social Science':{code:'gess1', name:'Social Science'},
  '7_Hindi':         {code:'ghvs1', name:'Vasant'},
  '7_English':       {code:'gehs1', name:'Honeycomb'},
  '7_Sanskrit':      {code:'gsans1',name:'Ruchira'},
  '8_Mathematics':   {code:'hema1', name:'Mathematics'},
  '8_Science':       {code:'hesc1', name:'Science'},
  '8_Social Science':{code:'hees1', name:'Exploring Society'},
  '8_Hindi':         {code:'hhvs1', name:'Vasant'},
  '8_English':       {code:'hehs1', name:'Honeydew'},
  '8_Sanskrit':      {code:'hsans1',name:'Ruchira'},
  '9_Mathematics':   {code:'iema1', name:'Mathematics'},
  '9_Science':       {code:'iesc1', name:'Science'},
  '9_Social Science':{code:'iess1', name:'Social Science'},
  '9_Hindi':         {code:'ihks1', name:'Kshitij'},
  '9_English':       {code:'iebs1', name:'Beehive'},
  '9_Sanskrit':      {code:'isans1',name:'Shemushi'},
  '10_Mathematics':  {code:'jema1', name:'Mathematics'},
  '10_Science':      {code:'jesc1', name:'Science'},
  '10_Social Science':{code:'jess1',name:'Social Science'},
  '10_Hindi':        {code:'jhks1', name:'Kshitij'},
  '10_English':      {code:'jefl1', name:'First Flight'},
  '10_Sanskrit':     {code:'jsans1',name:'Shemushi'},
  '11_Mathematics':  {code:'kema1', name:'Mathematics'},
  '11_Physics':      {code:'keph1', name:'Physics Part I'},
  '11_Chemistry':    {code:'kech1', name:'Chemistry Part I'},
  '11_Biology':      {code:'kebo1', name:'Biology'},
  '11_History':      {code:'kehb1', name:'Themes in World History'},
  '11_Geography':    {code:'kefg1', name:'Fundamentals of Physical Geography'},
  '11_Economics':    {code:'keie1', name:'Indian Economic Development'},
  '11_Political Science':{code:'kept1',name:'Political Theory'},
  '11_Hindi':        {code:'kehb1', name:'Aroh'},
  '11_English':      {code:'keeg1', name:'Hornbill'},
  '12_Mathematics':  {code:'lema1', name:'Mathematics Part I'},
  '12_Physics':      {code:'leph1', name:'Physics Part I'},
  '12_Chemistry':    {code:'lech1', name:'Chemistry Part I'},
  '12_Biology':      {code:'lebo1', name:'Biology'},
  '12_History':      {code:'lehis1',name:'Themes in Indian History I'},
  '12_Geography':    {code:'lefg1', name:'Fundamentals of Human Geography'},
  '12_Economics':    {code:'leie1', name:'Introductory Microeconomics'},
  '12_Political Science':{code:'leps1',name:'Contemporary World Politics'},
  '12_Hindi':        {code:'lehb1', name:'Aroh'},
  '12_English':      {code:'lefl1', name:'Flamingo'},
};

// ── HEALTH CHECK ─────────────────────────────────────────────────
app.get('/', (req,res) => res.json({status:'ok', message:'EMRS QPG v2.0 Server'}));

// ── ADMIN PANEL ───────────────────────────────────────────────────
app.get('/admin', (req,res) => {
  res.send(`<!DOCTYPE html>
<html><head><title>EMRS QPG Admin</title>
<style>body{font-family:Arial;padding:20px;background:#f5f0e0}
h2{color:#1C3A6B}table{width:100%;border-collapse:collapse}
td,th{border:1px solid #ccc;padding:8px;font-size:12px}
th{background:#1C3A6B;color:white}
input{width:90%;padding:4px}
button{background:#1C3A6B;color:white;padding:8px 16px;border:none;cursor:pointer;margin:10px 0}
</style></head><body>
<h2>EMRS QPG — Admin Panel</h2>
<p>Book codes update when NCERT releases new textbooks.</p>
<table id="tbl">
<tr><th>Class_Subject</th><th>Book Code</th><th>Book Name</th></tr>
${Object.entries(NCERT_BOOKS).map(([k,v])=>`
<tr><td>${k}</td>
<td><input id="code_${k}" value="${v.code}"/></td>
<td><input id="name_${k}" value="${v.name}"/></td></tr>`).join('')}
</table>
<button onclick="alert('In production: codes update via env vars or database. Contact developer to update.')">Save Changes</button>
</body></html>`);
});

// ── FETCH NCERT PDF (4-layer fallback) ────────────────────────────
app.post('/api/fetch-ncert', async (req,res) => {
  const {bookCode, chapterNumbers} = req.body;
  if(!bookCode || !chapterNumbers?.length)
    return res.status(400).json({error:'bookCode and chapterNumbers required'});

  let allText='', fetched=0, failed=0;
  const sources = [
    ch => `https://ncert.nic.in/textbook/pdf/${bookCode}${String(ch).padStart(2,'0')}.pdf`,
    ch => `https://ncert.nic.in/textbook/pdf/${bookCode}0${ch}.pdf`,
  ];

  for(const chNum of chapterNumbers){
    let chText = '';
    for(const srcFn of sources){
      try{
        const url = srcFn(chNum);
        const r = await axios.get(url, {
          responseType:'arraybuffer', timeout:25000,
          headers:{'User-Agent':'Mozilla/5.0','Referer':'https://ncert.nic.in/'}
        });
        const pdfParse = require('pdf-parse');
        const d = await pdfParse(Buffer.from(r.data));
        chText = d.text.replace(/\s+/g,' ').trim().slice(0,10000);
        if(chText.length > 200){ fetched++; break; }
      }catch(e){}
    }
    if(chText) allText += `\n\n=== CHAPTER ${chNum} ===\n${chText}`;
    else failed++;
  }
  res.json({success:true, text:allText.trim(), fetched, failed, totalChars:allText.length});
});

// ── GET BOOK INFO ─────────────────────────────────────────────────
app.post('/api/book-info', (req,res) => {
  const {cls, subject} = req.body;
  const key = `${cls}_${subject}`;
  const book = NCERT_BOOKS[key];
  if(!book) return res.json({found:false});
  res.json({found:true, ...book, key});
});

// ── AI CHAPTER SUGGESTIONS ────────────────────────────────────────
app.post('/api/suggest-chapters', async (req,res) => {
  const {cls, subject, exam, totalMarks, bookName} = req.body;
  const prompt = `You are a CBSE curriculum expert. Suggest which chapters to cover for:
Class: ${cls}
Subject: ${subject}
Book: ${bookName}
Exam: ${exam}
Total Marks: ${totalMarks}

Give chapter numbers and names with brief reason why each is important for this exam.
Format as JSON array: [{"chapter": 1, "name": "Chapter Name", "reason": "why important", "priority": "high/medium/low"}]
Return ONLY the JSON array, no other text.`;

  try{
    const r = await callClaude(prompt, 1000);
    const clean = r.replace(/```json|```/g,'').trim();
    const suggestions = JSON.parse(clean);
    res.json({success:true, suggestions});
  }catch(e){
    res.json({success:false, error:e.message});
  }
});

// ── CALL CLAUDE API ───────────────────────────────────────────────
async function callClaude(prompt, maxTokens=8000) {
  if(!ANTHROPIC_KEY) throw new Error('No Anthropic API key configured');
  const r = await axios.post('https://api.anthropic.com/v1/messages', {
    model: 'claude-sonnet-4-20250514',
    max_tokens: maxTokens,
    messages: [{role:'user', content:prompt}]
  }, {
    headers: {
      'x-api-key': ANTHROPIC_KEY,
      'anthropic-version': '2023-06-01',
      'content-type': 'application/json'
    },
    timeout: 180000
  });
  return r.data.content[0].text;
}

// ── CALL OPENROUTER (fallback) ────────────────────────────────────
async function callOpenRouter(prompt, maxTokens=8000) {
  if(!OPENROUTER_KEY) throw new Error('No OpenRouter key configured');
  const r = await axios.post('https://openrouter.ai/api/v1/chat/completions', {
    model: 'openrouter/auto',
    messages: [{role:'user', content:prompt}],
    max_tokens: maxTokens, temperature:0.7
  }, {
    headers: {
      'Authorization':`Bearer ${OPENROUTER_KEY}`,
      'Content-Type':'application/json',
      'HTTP-Referer':'https://emrs.school',
      'X-Title':'EMRS QPG'
    },
    timeout: 120000
  });
  return r.data.choices[0].message.content;
}

// ── CALL AI (with fallback) ───────────────────────────────────────
async function callAI(prompt, maxTokens=8000) {
  try {
    return await callClaude(prompt, maxTokens);
  } catch(e) {
    console.log('Claude failed, trying OpenRouter:', e.message);
    return await callOpenRouter(prompt, maxTokens);
  }
}

// ── BUILD PROMPT ──────────────────────────────────────────────────
function buildPrompt(config, ncertText, blueprint) {
  const {cls, subject, exam, session, marks, duration, medium, difficulty,
         teacher, date, chapters, setLabel, stream} = config;

  const diffMap = {easy:'40% Easy, 40% Medium, 20% Hard',
    medium:'20% Easy, 50% Medium, 30% Hard',
    hard:'10% Easy, 40% Medium, 50% Hard',
    mixed:'30% Easy, 40% Medium, 30% Hard'};

  const isSci = ['Science','Physics','Chemistry','Biology'].includes(subject);
  const isGeo = ['Social Science','Geography','History'].includes(subject);
  const isMath = subject === 'Mathematics';
  const isSans = subject === 'Sanskrit';
  const isHindi = subject === 'Hindi' || subject === 'Hindi Core';

  let p = `You are an expert CBSE question paper setter for Class ${cls}${stream?' ('+stream+')':''} ${subject}.
Generate a complete question paper for EMRS Bansla-Bagidora, Banswara.

PAPER DETAILS:
Exam: ${exam} | Session: ${session} | Set: ${setLabel}
Total Marks: ${marks} | Duration: ${duration}
Medium: ${medium} | Difficulty: ${diffMap[difficulty]||diffMap.mixed}
${teacher?'Teacher: '+teacher:''}${date?'  Date: '+date:''}

BLUEPRINT TO FOLLOW EXACTLY:
${blueprint.map(b=>`Section ${b.sec}: ${b.type} — ${b.q} questions × ${b.m} marks = ${b.tot} marks`).join('\n')}

${ncertText ? `NCERT CONTENT (use this as the primary source):\n${ncertText.slice(0,12000)}\n` : ''}
${chapters?.length ? `CHAPTERS COVERED: ${chapters.join(', ')}\n` : ''}

CRITICAL OUTPUT FORMAT RULES — FOLLOW ALL 15:
RULE 1: NO markdown. No **, no ##, no --- lines, no | table syntax.
RULE 2: Do NOT write paper header/school/subject in the paper body.
RULE 3: Start with: GENERAL INSTRUCTIONS
RULE 4: All questions numbered Q1. Q2. continuously.
RULE 5: Marks at END of question: Q1. Question text here [2 Marks]
RULE 6: Section headers exactly: SECTION A  (on its own line)
RULE 7: MCQ options on separate lines: (A) option text
RULE 8: After EVERY MCQ/FIB/TF: Answer / \u0909\u0924\u094d\u0924\u0930 : ___________
RULE 9: Internal choice: write OR on its own line only.
RULE 10: Match Column — plain text NO markdown tables:
   Column I          Column II
   A. Term1          1. Description1
RULE 11: Fill in Blanks: Q7. The ___________ is the capital. [1 Mark]
RULE 12: True/False: Q9. Sun rises in east. (True/False) [1 Mark]
RULE 13: Case Study: plain paragraph then (i)(ii)(iii) sub-questions with marks
RULE 14: ${medium==='Bilingual'?'Write English question then Hindi translation on next line (no Q number repeat)':'Write questions in '+medium+' only'}
RULE 15: End paper with exactly: ***
${isSci?'Include at least one DIAGRAM BASED question (describe diagram in text).':''}
${isGeo?'Include at least one MAP BASED question.':''}
${isMath?'Include NUMERICAL/PROOF type questions. Use proper mathematical notation.':''}
${isSans?'Include Sanskrit grammar questions (sandhi, samas, pratyaya).':''}

After *** write exactly: ===ANSWER KEY===
Write complete answer key section-wise.
For MCQ: Q1. (B) — one line reason
For VSA: 2-3 line answer
For SA: key points with marks breakdown
For LA: detailed outline with marks
For Case Study: answer each sub-question

After answer key write exactly: ===BLUEPRINT===
Write marks distribution table as plain text.

Generate complete paper now following ALL rules.`;
  return p;
}

// ── SHUFFLE ARRAY ─────────────────────────────────────────────────
function shuffle(arr, seed=1) {
  const a = [...arr];
  let s = seed * 9301 + 49297;
  for(let i=a.length-1; i>0; i--){
    s = (s * 9301 + 49297) % 233280;
    const j = Math.floor((s/233280) * (i+1));
    [a[i],a[j]] = [a[j],a[i]];
  }
  return a;
}

// ── SHUFFLE PAPER TEXT ────────────────────────────────────────────
function shufflePaper(paperText, seed) {
  const lines = paperText.split('\n');
  const sections = [];
  let current = {header:'', questions:[]};
  let currentQ = null;

  for(const line of lines){
    const clean = line.replace(/\*\*/g,'').trim();
    if(clean.match(/^SECTION[\s\-–]+[A-F]/i)){
      if(currentQ) current.questions.push(currentQ);
      currentQ = null;
      if(current.header || current.questions.length)
        sections.push(current);
      current = {header:line, questions:[]};
    } else if(clean.match(/^Q\d+[\.\)]/)){
      if(currentQ) current.questions.push(currentQ);
      currentQ = {lines:[line], marks:0};
      const mMatch = clean.match(/\[(\d+)\s*[Mm]arks?\]/);
      if(mMatch) currentQ.marks = parseInt(mMatch[1]);
    } else if(clean === '***' || clean.match(/^===ANSWER KEY===/)){
      if(currentQ) current.questions.push(currentQ);
      currentQ = null;
      sections.push(current);
      // Everything after *** stays as-is
      const restIdx = lines.indexOf(line);
      return sections.map(sec => {
        const shuffled = shuffle(sec.questions, seed);
        return sec.header + '\n' + shuffled.map((q,i) => {
          return q.lines.join('\n').replace(/^Q\d+[\.\)]/,`Q${i+1}.`);
        }).join('\n');
      }).join('\n') + '\n' + lines.slice(lines.indexOf(line)).join('\n');
    } else {
      if(currentQ) currentQ.lines.push(line);
      else current.questions.push({lines:[line], marks:0});
    }
  }
  if(currentQ) current.questions.push(currentQ);
  sections.push(current);
  return sections.map(sec => {
    const shuffled = shuffle(sec.questions, seed);
    return sec.header + '\n' + shuffled.map((q,i) => {
      if(q.lines[0]?.match(/^Q\d+[\.\)]/))
        return q.lines.join('\n').replace(/^Q\d+[\.\)]/,`Q${i+1}.`);
      return q.lines.join('\n');
    }).join('\n');
  }).join('\n');
}

// ── PARSE PAPER INTO SECTIONS ─────────────────────────────────────
function parsePaperAndKey(fullText) {
  const akIdx = fullText.indexOf('===ANSWER KEY===');
  const bpIdx = fullText.indexOf('===BLUEPRINT===');
  const paperText = akIdx > 0 ? fullText.slice(0, akIdx).trim() : fullText;
  const answerKey = akIdx > 0 && bpIdx > 0 ? fullText.slice(akIdx+16, bpIdx).trim()
                  : akIdx > 0 ? fullText.slice(akIdx+16).trim() : '';
  const blueprint = bpIdx > 0 ? fullText.slice(bpIdx+15).trim() : '';
  return {paperText, answerKey, blueprint};
}

// ── GENERATE PAPER ────────────────────────────────────────────────
app.post('/api/generate', async (req,res) => {
  const {config, ncertText, blueprint} = req.body;
  if(!config) return res.status(400).json({error:'config required'});

  try{
    const numSets = config.numSets || 1;
    const results = [];

    // Generate Set A (original)
    const promptA = buildPrompt({...config, setLabel:'Set A / सेट अ'}, ncertText||'', blueprint||[]);
    const rawA = await callAI(promptA);
    const {paperText:ptA, answerKey:akA, blueprint:bpA} = parsePaperAndKey(rawA);
    results.push({set:'A', paperText:ptA, answerKey:akA, blueprint:bpA});

    // Generate Set B and C if needed (shuffle)
    if(numSets >= 3){
      const ptB = shufflePaper(ptA, 42);
      results.push({set:'B', paperText:ptB, answerKey:akA, blueprint:bpA});
      const ptC = shufflePaper(ptA, 137);
      results.push({set:'C', paperText:ptC, answerKey:akA, blueprint:bpA});
    }

    res.json({success:true, sets:results});
  }catch(e){
    console.error(e);
    res.status(500).json({error:e.message});
  }
});

// ── CREATE DOCX ───────────────────────────────────────────────────
app.post('/api/create-docx', async (req,res) => {
  const {paperText, answerKey, blueprintText, config, docType} = req.body;
  if(!paperText) return res.status(400).json({error:'paperText required'});

  try{
    let buf;
    if(docType === 'answerkey')
      buf = await buildAnswerKeyDocx(answerKey, config);
    else if(docType === 'blueprint')
      buf = await buildBlueprintDocx(blueprintText, config);
    else
      buf = await buildPaperDocx(paperText, config);

    const setLabel = config.setLabel||'A';
    const fn = `EMRS_${config.cls}_${config.sub}_${config.exam}_Set${setLabel}_${docType||'Paper'}.docx`
      .replace(/\s+/g,'_');
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition',`attachment; filename="${fn}"`);
    res.setHeader('Access-Control-Expose-Headers','Content-Disposition');
    res.send(buf);
  }catch(e){
    res.status(500).json({error:e.message});
  }
});

// ══════════════════════════════════════════════════════════════════
// DOCX BUILDERS
// ══════════════════════════════════════════════════════════════════

// ── HELPERS ───────────────────────────────────────────────────────
function tr(text,size,bold,font,color,italic){
  const f = font||'Times New Roman';
  return new TextRun({
    text:String(text), bold:!!bold, italics:!!italic,
    size:Math.round((size||11)*2), font:f, color:color||'000000'
  });
}
function trH(text,size,bold,color){ return tr(text,size,bold,'Kokila',color); }
function cp(runs,align,spacing,indent){
  return new Paragraph({
    alignment:align||AlignmentType.LEFT,
    spacing:spacing||{before:40,after:40},
    indent:indent?{left:indent}:undefined,
    children:Array.isArray(runs)?runs:[runs]
  });
}
function sp(sz){ return new Paragraph({text:'',spacing:{before:sz||60,after:sz||60}}); }
function mg(v){ const m=v||80; return {top:m,bottom:m,left:m+20,right:m+20}; }
function bdr(color,size){
  const b={style:BorderStyle.SINGLE,size:size||4,color:color||'000000'};
  return {top:b,bottom:b,left:b,right:b,insideHorizontal:b,insideVertical:b};
}
function noBdr(){
  const b={style:BorderStyle.NONE,size:0,color:'FFFFFF'};
  return {top:b,bottom:b,left:b,right:b,insideHorizontal:b,insideVertical:b};
}
function isHin(t){ return /[\u0900-\u097F]/.test(t); }

function hiExam(e){
  const m={
    'Unit Test 1':'\u0907\u0915\u093e\u0908 \u092a\u0930\u0940\u0915\u094d\u0937\u0923 - I',
    'Unit Test 2':'\u0907\u0915\u093e\u0908 \u092a\u0930\u0940\u0915\u094d\u0937\u0923 - II',
    'Unit Test 3':'\u0907\u0915\u093e\u0908 \u092a\u0930\u0940\u0915\u094d\u0937\u0923 - III',
    'Unit Test 4':'\u0907\u0915\u093e\u0908 \u092a\u0930\u0940\u0915\u094d\u0937\u0923 - IV',
    'Half-Yearly':'\u0905\u0930\u094d\u0927\u0935\u093e\u0930\u094d\u0937\u093f\u0915 \u092a\u0930\u0940\u0915\u094d\u0937\u093e',
    'Annual':'\u0935\u093e\u0930\u094d\u0937\u093f\u0915 \u092a\u0930\u0940\u0915\u094d\u0937\u093e',
    'Annual (Board)':'\u0935\u093e\u0930\u094d\u0937\u093f\u0915 \u092a\u0930\u0940\u0915\u094d\u0937\u093e (\u092c\u094b\u0930\u094d\u0921)'
  };
  return m[e]||e;
}
function hiSub(s){
  const m={
    'Mathematics':'\u0917\u0923\u093f\u0924','Science':'\u0935\u093f\u091c\u094d\u091e\u093e\u0928',
    'Social Science':'\u0938\u093e\u092e\u093e\u091c\u093f\u0915 \u0935\u093f\u091c\u094d\u091e\u093e\u0928',
    'Hindi':'\u0939\u093f\u0902\u0926\u0940','English':'\u0905\u0902\u0917\u094d\u0930\u0947\u091c\u093c\u0940',
    'Sanskrit':'\u0938\u0902\u0938\u094d\u0915\u0943\u0924','Physics':'\u092d\u094c\u0924\u093f\u0915\u0940',
    'Chemistry':'\u0930\u0938\u093e\u092f\u0928 \u0935\u093f\u091c\u094d\u091e\u093e\u0928',
    'Biology':'\u091c\u0940\u0935 \u0935\u093f\u091c\u094d\u091e\u093e\u0928',
    'History':'\u0907\u0924\u093f\u0939\u093e\u0938','Geography':'\u092d\u0942\u0917\u094b\u0932',
    'Political Science':'\u0930\u093e\u091c\u0928\u0940\u0924\u093f \u0935\u093f\u091c\u094d\u091e\u093e\u0928',
    'Economics':'\u0905\u0930\u094d\u0925\u0936\u093e\u0938\u094d\u0924\u094d\u0930',
    'Physical Education':'\u0936\u093e\u0930\u0940\u0930\u093f\u0915 \u0936\u093f\u0915\u094d\u0937\u093e',
    'Business Studies':'\u0935\u094d\u092f\u0935\u0938\u093e\u092f \u0905\u0927\u094d\u092f\u092f\u0928',
    'Accountancy':'\u0932\u0947\u0916\u093e\u0936\u093e\u0938\u094d\u0924\u094d\u0930'
  };
  return m[s]||s;
}

// ── BUILD PAPER HEADER ────────────────────────────────────────────
function buildHeader(config, W) {
  const {cls, sub, exam, session, marks, duration, teacher, date, setLabel, stream} = config;
  const clsDisplay = `${cls}${stream?' ('+stream+')':''}`;
  const children = [];

  // Table 0: School Header
  children.push(new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[W],
    borders:bdr(), rows:[new TableRow({children:[new TableCell({
      margins:mg(100),
      children:[
        cp([tr(SCH_EN,17,true,'Times New Roman',C.black)]),
        cp([trH(SCH_HI,13,true,C.black)]),
        cp([tr(SCH_LOC,12,true,'Times New Roman',C.black)]),
        cp([tr(SCH_AFF,9,false,'Times New Roman',C.black,true)]),
      ]
    })]})]
  }));
  children.push(sp(40));

  // Table 1: Exam Info
  children.push(new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[3200,3200,2960],
    borders:bdr(), rows:[new TableRow({children:[
      new TableCell({width:{size:3200,type:WidthType.DXA},margins:mg(80),children:[
        cp([tr(exam+'  /  '+hiExam(exam),11.5,true,'Times New Roman',C.black)]),
        cp([tr('Session: '+session+'  /  \u0938\u0924\u094d\u0930: '+session,9,false,'Times New Roman',C.black)]),
        cp([tr('Set / \u0938\u0947\u091f: '+(setLabel||'A'),10,true,'Times New Roman',C.marks)]),
      ]}),
      new TableCell({width:{size:3200,type:WidthType.DXA},margins:mg(80),verticalAlign:VerticalAlign.CENTER,children:[
        cp([tr(sub.toUpperCase(),12,true,'Times New Roman',C.black)],AlignmentType.CENTER),
        cp([trH(hiSub(sub),11,true,C.black)],AlignmentType.CENTER),
      ]}),
      new TableCell({width:{size:2960,type:WidthType.DXA},margins:mg(80),children:[
        cp([tr('Class / \u0915\u0915\u094d\u0937\u093e: '+clsDisplay,10.5,true,'Times New Roman',C.black)]),
        cp([tr('Max. Marks / \u0905\u0902\u0915: '+marks,9.5,false,'Times New Roman',C.black)]),
        cp([tr('Time / \u0938\u092e\u092f: '+duration,9.5,false,'Times New Roman',C.black)]),
        ...(teacher?[cp([tr('Teacher: '+teacher,9,false,'Times New Roman',C.black)])]:[]),
        ...(date?[cp([tr('Date: '+date,9,false,'Times New Roman',C.black)])]:[]),
      ]}),
    ]})]
  }));
  children.push(sp(40));

  // Table 2: Student Info
  children.push(new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[4000,2800,2560],
    borders:bdr(), rows:[new TableRow({children:[
      new TableCell({margins:mg(70),children:[cp([tr('Name / \u0928\u093e\u092e : __________________________________',10,false,'Times New Roman',C.black)])]}),
      new TableCell({margins:mg(70),children:[cp([tr('Roll No. / \u0915\u094d\u0930\u092e\u093e\u0902\u0915 : __________',10,false,'Times New Roman',C.black)])]}),
      new TableCell({margins:mg(70),children:[cp([tr('Section / \u0905\u0928\u0941\u092d\u093e\u0917 : _____',10,false,'Times New Roman',C.black)])]}),
    ]})]
  }));
  children.push(sp(40));

  // General Instructions heading
  children.push(cp([tr('GENERAL INSTRUCTIONS  /  \u0938\u093e\u092e\u093e\u0928\u094d\u092f \u0928\u093f\u0930\u094d\u0926\u0947\u0936',12,true,'Times New Roman',C.qNum)],AlignmentType.CENTER));
  children.push(sp(30));

  // Table 3: Instructions bilingual
  const INSTRS = [
    ['All questions are compulsory.','\u0938\u092d\u0940 \u092a\u094d\u0930\u0936\u094d\u0928 \u0905\u0928\u093f\u0935\u093e\u0930\u094d\u092f \u0939\u0948\u0902\u0964'],
    ['This paper is divided into sections as shown above.','\u092a\u094d\u0930\u0936\u094d\u0928-\u092a\u0924\u094d\u0930 \u0909\u092a\u0930\u094b\u0915\u094d\u0924 \u0916\u0902\u0921\u094b\u0902 \u092e\u0947\u0902 \u0935\u093f\u092d\u093e\u091c\u093f\u0924 \u0939\u0948\u0964'],
    ['Internal choices are given where indicated.','\u091c\u0939\u093e\u0901 \u0906\u0902\u0924\u0930\u093f\u0915 \u0935\u093f\u0915\u0932\u094d\u092a \u0926\u093f\u090f \u0917\u090f \u0939\u094b\u0902, \u0935\u0939\u093e\u0901 \u0915\u094b\u0908 \u090f\u0915 \u0909\u0924\u094d\u0924\u0930 \u0926\u0947\u0902\u0964'],
    ['Draw neat diagrams wherever required.','\u091c\u0939\u093e\u0901 \u0906\u0935\u0936\u094d\u092f\u0915 \u0939\u094b \u0938\u094d\u092a\u0937\u094d\u091f \u0906\u0930\u0947\u0916 \u092c\u0928\u093e\u0907\u090f\u0964'],
    ['Read each question carefully before answering.','\u0909\u0924\u094d\u0924\u0930 \u0926\u0947\u0928\u0947 \u0938\u0947 \u092a\u0939\u0932\u0947 \u092a\u094d\u0930\u0924\u094d\u092f\u0947\u0915 \u092a\u094d\u0930\u0936\u094d\u0928 \u0927\u094d\u092f\u093e\u0928\u092a\u0942\u0930\u094d\u0935\u0915 \u092a\u095d\u093c\u0947\u0902\u0964'],
    ['Marks are indicated against each question.','\u092a\u094d\u0930\u0924\u094d\u092f\u0947\u0915 \u092a\u094d\u0930\u0936\u094d\u0928 \u0915\u0947 \u0938\u093e\u092e\u0928\u0947 \u0905\u0902\u0915 \u0926\u0930\u094d\u0936\u093e\u090f \u0917\u090f \u0939\u0948\u0902\u0964'],
  ];
  children.push(new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[4680,4680], borders:bdr(),
    rows:INSTRS.map((pair,idx)=>new TableRow({children:[
      new TableCell({width:{size:4680,type:WidthType.DXA},margins:mg(60),children:[
        cp([tr(`${idx+1}. `,9.5,true,'Times New Roman',C.black),tr(pair[0],9.5,false,'Times New Roman',C.black)])
      ]}),
      new TableCell({width:{size:4680,type:WidthType.DXA},margins:mg(60),children:[
        cp([tr(`${idx+1}. `,9.5,true,'Kokila',C.black),trH(pair[1],9.5,false,C.black)])
      ]}),
    ]}))
  }));
  children.push(sp(60));
  return children;
}

// ── BUILD SECTION HEADER ──────────────────────────────────────────
function buildSectionHeader(secLetter, secName, marksFormula, W) {
  const hiL = SEC_HI[secLetter]||secLetter;
  const hiN = SEC_NAME_HI[secLetter]||secName;
  return new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[7200,2160], borders:bdr(),
    rows:[new TableRow({children:[
      new TableCell({
        width:{size:7200,type:WidthType.DXA},
        shading:{fill:C.shade,type:ShadingType.CLEAR},
        margins:mg(80),
        children:[cp([
          tr(`SECTION \u2013 ${secLetter}   ${secName}`,12,true,'Times New Roman',C.black),
          tr('   |   ',10,false,'Times New Roman',C.black),
          trH(`\u0916\u0902\u0921 \u2013 ${hiL}   ${hiN}`,11,true,C.black),
        ])]
      }),
      new TableCell({
        width:{size:2160,type:WidthType.DXA},
        shading:{fill:C.shade,type:ShadingType.CLEAR},
        margins:mg(80), verticalAlign:VerticalAlign.CENTER,
        children:[cp([tr(marksFormula,10,true,'Times New Roman',C.black)],AlignmentType.RIGHT)]
      }),
    ]})]
  });
}

// ── PARSE AND RENDER PAPER ────────────────────────────────────────
function renderPaperLines(lines, blueprint, W) {
  const children = [];
  let i = 0;
  let inCaseBox = false, caseLines = [];

  const flushCase = () => {
    if(!caseLines.length) return;
    const paras = caseLines.map(bl => {
      if(!bl.trim()) return sp(20);
      const isHdr = !!(bl.match(/^(READ|MAP|Case|Source|PASSAGE|Study|Observe)/i));
      const isH = isHin(bl);
      return cp([isH?trH(bl,10,isHdr,C.black):tr(bl,10,isHdr,'Times New Roman',C.black)],AlignmentType.LEFT,{before:20,after:20});
    });
    children.push(new Table({
      width:{size:W,type:WidthType.DXA}, columnWidths:[W],
      borders:bdr(C.black,6),
      rows:[new TableRow({children:[new TableCell({margins:mg(120),children:paras})]})]
    }));
    children.push(sp(80));
    caseLines = [];
    inCaseBox = false;
  };

  while(i < lines.length){
    const raw = lines[i];
    const line = raw.replace(/\*\*/g,'').replace(/^#+\s*/,'').trim();

    if(!line){ inCaseBox?caseLines.push(''):children.push(sp(40)); i++;continue; }
    if(line.match(/^\|[-:\s|]+\|$/)||line.match(/^-{3,}$/)){ i++;continue; }
    if(line.match(/^(EMRS |Subject:|Maximum Marks:|Chapters Covered:|Set [ABC])/i)){ i++;continue; }

    // Section header
    const secM = line.match(/^SECTION[\s\-–]+([A-F])\b(.*)$/i);
    if(secM){
      flushCase();
      const L = secM[1].toUpperCase();
      const after = secM[2].trim();
      const bpRow = Array.isArray(blueprint)?blueprint.find(b=>b.sec===L):null;
      const mf = bpRow?`${bpRow.q} \u00d7 ${bpRow.m} = ${bpRow.tot} Marks`:'';
      children.push(buildSectionHeader(L, after, mf, W));
      children.push(sp(80));
      i++;continue;
    }

    // Case/Source/Map box trigger
    if(!inCaseBox && line.match(/^(Read the|Case[\s\-]*(Study|Based)|Source[\s\-]*Based|MAP SKILL|Observe|Study the)/i)){
      flushCase();
      inCaseBox = true;
      caseLines.push(line);
      i++;continue;
    }
    if(inCaseBox){
      if(line.match(/^Q\d+[\.\)]/) && !line.match(/^\(([ivx]+)\)/i)) flushCase();
      else{ caseLines.push(line.startsWith('|')?line.split('|').map(s=>s.trim()).filter(Boolean).join('   '):line); i++;continue; }
    }

    // Question
    const qM = line.match(/^(Q\d+[\.\)])\s*(.*)/);
    if(qM){
      const qNum = qM[1]; let rest = qM[2].trim();
      const mkM = rest.match(/\s*\[(\d+)\s*[Mm]arks?[^\]]*\]\s*$/) || rest.match(/\s*\[(\d+)\]\s*$/);
      let marksStr = '';
      if(mkM){ marksStr=`[${mkM[1]} Mark${mkM[1]==='1'?'':'s'} / ${mkM[1]} \u0905\u0902\u0915]`; rest=rest.slice(0,rest.lastIndexOf(mkM[0])).trim(); }
      children.push(cp([
        tr(qNum+'  ',11,true,'Times New Roman',C.qNum),
        tr(rest,11,false,'Times New Roman',C.qText),
        ...(marksStr?[tr('   '+marksStr,10,true,'Times New Roman',C.marks)]:[])
      ],AlignmentType.LEFT,{before:140,after:20}));
      // Hindi translation on next line
      let j=i+1;
      while(j<lines.length&&!lines[j].trim()) j++;
      if(j<lines.length){
        const nl=lines[j].replace(/\*\*/g,'').trim();
        if(nl&&isHin(nl)&&!nl.match(/^Q\d+/)&&!nl.match(/^\([a-d]\)/i)&&!nl.match(/^SECTION/i)){
          children.push(cp([trH(nl,10,false,C.hindi)],AlignmentType.LEFT,{before:0,after:20}));
          i=j;
        }
      }
      i++;continue;
    }

    // Sub-question (i)(ii)(iii)
    const subM = line.match(/^\(([ivxIVX]+)\)\s*(.*)/);
    if(subM){
      const rest2=subM[2].trim();
      const mkM2=rest2.match(/\s*\[(\d+)[^\]]*\]\s*$/) || rest2.match(/\s*\[(\d+)\]\s*$/);
      let sMarks='',sTxt=rest2;
      if(mkM2){sMarks=`[${mkM2[1]} \u0905\u0902\u0915]`;sTxt=rest2.slice(0,rest2.lastIndexOf(mkM2[0])).trim();}
      children.push(cp([
        tr(`(${subM[1]})  `,10,true,'Times New Roman',C.qNum),
        tr(sTxt,10.5,false,'Times New Roman',C.qText),
        ...(sMarks?[tr('   '+sMarks,10,true,'Times New Roman',C.marks)]:[])
      ],AlignmentType.LEFT,{before:60,after:20},360));
      i++;continue;
    }

    // Options (A)(B)(C)(D) — CBSE capital letters
    const optM = line.match(/^\(([A-Da-d])\)\s*(.*)/);
    if(optM){
      const optL=optM[1].toUpperCase(); const optRest=optM[2].trim();
      const parts=optRest.split(/\s{2,}\/\s{2,}|\s+\/\s+(?=[\u0900-\u097F])/);
      const en=parts[0]?parts[0].trim():optRest;
      const hi=parts[1]?parts[1].trim():'';
      children.push(cp([
        tr(`(${optL})  `,10,true,'Times New Roman',C.option),
        tr(en,10,false,'Times New Roman',C.optTxt),
        ...(hi?[tr('   /   ',9,false,'Times New Roman',C.optSep),trH(hi,9.5,false,C.optHi)]:[])
      ],AlignmentType.LEFT,{before:20,after:20},360));
      i++;continue;
    }

    // Markdown table row
    if(line.startsWith('|')){
      const cells=line.split('|').map(s=>s.trim()).filter(Boolean);
      if(cells.length) children.push(cp([tr(cells.join('   '),10,false,'Times New Roman',C.black)],AlignmentType.LEFT,{before:20,after:20},360));
      i++;continue;
    }

    // OR separator
    if(line==='OR'||line.match(/^OR\s*\/\s*\u0905\u0925\u0935\u093e/i)){
      children.push(cp([tr('OR  /  \u0905\u0925\u0935\u093e',10,true,'Times New Roman',C.marks)],AlignmentType.CENTER,{before:80,after:80}));
      i++;continue;
    }

    // Answer blank
    if(line.match(/^Answer\s*[\/|:]/i)){
      children.push(cp([tr('Answer / \u0909\u0924\u094d\u0924\u0930 : ___________',10,false,'Times New Roman',C.answer)],AlignmentType.LEFT,{before:20,after:60}));
      i++;continue;
    }

    // *** end marker
    if(line==='***'||line==='* * *'){
      children.push(cp([tr('\u2022 \u2022 \u2022',12,true,'Times New Roman',C.black)],AlignmentType.CENTER,{before:160,after:80}));
      i++;continue;
    }

    // Part header
    if(line.match(/^Part[\-\s]*(I|II|III|IV|V)\b/i)){
      children.push(cp([tr(line,10.5,true,'Times New Roman',C.marks)],AlignmentType.LEFT,{before:80,after:40}));
      i++;continue;
    }

    // Hindi only line
    if(isHin(line)&&!line.match(/^Q\d+/)){
      children.push(cp([trH(line,10,false,C.hindi)],AlignmentType.LEFT,{before:0,after:20}));
      i++;continue;
    }

    // Assertion/Reason label lines
    if(line.match(/^(Assertion|Reason|Statement)/i)){
      children.push(cp([tr(line,10.5,false,'Times New Roman',C.black,true)],AlignmentType.LEFT,{before:20,after:20},360));
      i++;continue;
    }

    // Default
    const isBold=!!(line.match(/^(GENERAL INSTRUCTIONS|Answer each)/i));
    children.push(cp([tr(line,10.5,isBold,'Times New Roman',C.black)],AlignmentType.LEFT,{before:40,after:40}));
    i++;
  }
  flushCase();
  return children;
}

// ── BUILD PAPER DOCX ──────────────────────────────────────────────
async function buildPaperDocx(paperText, config) {
  const W = 9360;
  const blueprint = config.blueprint||[];
  const children = [...buildHeader(config, W)];
  const lines = paperText.split('\n');
  children.push(...renderPaperLines(lines, blueprint, W));

  // Footer
  children.push(sp(120));
  children.push(cp([
    tr('\u2014\u2014  ',10,false,'Times New Roman',C.footer),
    tr('Best of Luck! / \u0936\u0941\u092d\u0915\u093e\u092e\u0928\u093e\u090f\u0901!',13,true,'Times New Roman',C.qNum),
    tr('  \u2014\u2014',10,false,'Times New Roman',C.footer),
  ],AlignmentType.CENTER,{before:80,after:40}));
  children.push(cp([tr('Eklavya Model Residential School, Bansla-Bagidora, Banswara (Rajasthan)',8,false,'Times New Roman',C.footer)],AlignmentType.CENTER));

  const doc = new Document({sections:[{
    properties:{page:{size:{width:11906,height:16838},margin:{top:720,right:900,bottom:720,left:900}}},
    children
  }]});
  return await Packer.toBuffer(doc);
}

// ── BUILD ANSWER KEY DOCX ─────────────────────────────────────────
async function buildAnswerKeyDocx(answerKey, config) {
  const W = 9360;
  const children = [];

  // Header box
  children.push(new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[W], borders:bdr(),
    rows:[new TableRow({children:[new TableCell({
      margins:mg(100),
      shading:{fill:C.shade,type:ShadingType.CLEAR},
      children:[
        cp([tr('ANSWER KEY / MARKING SCHEME',14,true,'Times New Roman',C.qNum)],AlignmentType.CENTER),
        cp([trH('\u0909\u0924\u094d\u0924\u0930 \u0915\u0941\u0902\u091c\u0940 / \u0905\u0902\u0915\u0928 \u092f\u094b\u091c\u0928\u093e',12,true,C.qNum)],AlignmentType.CENTER),
        cp([tr(`${config.sub||''} | Class ${config.cls||''} | ${config.exam||''} | Set ${config.setLabel||'A'} | Session: ${config.session||''}`,10,false,'Times New Roman',C.black)],AlignmentType.CENTER),
      ]
    })]})]
  }));
  children.push(sp(80));

  for(const line of (answerKey||'').split('\n')){
    const t = line.replace(/\*\*/g,'').trim();
    if(!t){ children.push(sp(40)); continue; }
    const isSecH = !!(t.match(/^SECTION/i));
    const isQH = !!(t.match(/^Q\d+[\.\)]/));
    const isH = isHin(t);
    if(isSecH){
      children.push(sp(60));
      children.push(cp([tr(t,12,true,'Times New Roman',C.qNum)],AlignmentType.LEFT,{before:80,after:40}));
    } else if(isQH){
      children.push(cp([tr(t,11,true,'Times New Roman',C.black)],AlignmentType.LEFT,{before:60,after:20}));
    } else if(isH){
      children.push(cp([trH(t,10,false,C.hindi)],AlignmentType.LEFT,{before:20,after:20}));
    } else {
      children.push(cp([tr(t,10.5,false,'Times New Roman',C.black)],AlignmentType.LEFT,{before:20,after:20},360));
    }
  }

  children.push(sp(80));
  children.push(cp([tr('— Prepared by EMRS QPG System —',9,false,'Times New Roman',C.footer)],AlignmentType.CENTER));

  const doc = new Document({sections:[{
    properties:{page:{size:{width:11906,height:16838},margin:{top:720,right:900,bottom:720,left:900}}},
    children
  }]});
  return await Packer.toBuffer(doc);
}

// ── BUILD BLUEPRINT DOCX ──────────────────────────────────────────
async function buildBlueprintDocx(bpText, config) {
  const W = 9360;
  const children = [];

  // Title
  children.push(new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[W], borders:bdr(),
    rows:[new TableRow({children:[new TableCell({
      margins:mg(100),
      shading:{fill:'1C3A6B',type:ShadingType.CLEAR},
      children:[
        cp([tr('QUESTION PAPER BLUEPRINT',14,true,'Times New Roman','FFFFFF')],AlignmentType.CENTER),
        cp([trH('\u092a\u094d\u0930\u0936\u094d\u0928 \u092a\u0924\u094d\u0930 \u092c\u094d\u0932\u0942\u092a\u094d\u0930\u093f\u0902\u091f',12,true,'FFFFFF')],AlignmentType.CENTER),
        cp([tr(`${config.sub||''} | Class ${config.cls||''} | ${config.exam||''} | Marks: ${config.marks||40}`,10,false,'Times New Roman','FFFFFF')],AlignmentType.CENTER),
      ]
    })]})]
  }));
  children.push(sp(80));

  // Blueprint table
  const colW = Math.floor(W/6);
  children.push(new Table({
    width:{size:W,type:WidthType.DXA},
    columnWidths:[colW,colW*2,colW,colW,colW,colW],
    borders:bdr(),
    rows:[
      // Header row
      new TableRow({children:[
        'Section','Question Type','No. of Qs','Marks Each','Total Marks','Internal Choice'
      ].map(h=>new TableCell({
        shading:{fill:C.shade,type:ShadingType.CLEAR},
        margins:mg(60),
        children:[cp([tr(h,9.5,true,'Times New Roman',C.qNum)],AlignmentType.CENTER)]
      }))}),
      // Data rows from blueprint config
      ...(Array.isArray(config.blueprint)?config.blueprint.map(b=>
        new TableRow({children:[
          new TableCell({margins:mg(60),children:[cp([tr('Sec '+b.sec,10,true,'Times New Roman',C.black)],AlignmentType.CENTER)]}),
          new TableCell({margins:mg(60),children:[cp([tr(b.type||'',10,false,'Times New Roman',C.black)])]}),
          new TableCell({margins:mg(60),children:[cp([tr(String(b.q),12,true,'Times New Roman',C.black)],AlignmentType.CENTER)]}),
          new TableCell({margins:mg(60),children:[cp([tr(String(b.m),12,true,'Times New Roman',C.black)],AlignmentType.CENTER)]}),
          new TableCell({margins:mg(60),children:[cp([tr(String(b.tot),12,true,'Times New Roman',C.marks)],AlignmentType.CENTER)]}),
          new TableCell({margins:mg(60),children:[cp([tr(b.choice||'No',10,false,'Times New Roman',C.black)],AlignmentType.CENTER)]}),
        ]})
      ):[]),
      // Total row
      ...(Array.isArray(config.blueprint)?[new TableRow({children:[
        new TableCell({columnSpan:4,margins:mg(60),shading:{fill:C.shade,type:ShadingType.CLEAR},children:[cp([tr('TOTAL',11,true,'Times New Roman',C.qNum)],AlignmentType.RIGHT)]}),
        new TableCell({margins:mg(60),shading:{fill:C.shade,type:ShadingType.CLEAR},children:[cp([tr(String(config.blueprint.reduce((a,b)=>a+b.tot,0)),14,true,'Times New Roman',C.marks)],AlignmentType.CENTER)]}),
        new TableCell({margins:mg(60),shading:{fill:C.shade,type:ShadingType.CLEAR},children:[cp([tr('',10,false,'Times New Roman',C.black)])]}),
      ]})]:[]),
    ]
  }));

  children.push(sp(80));

  // Also render any blueprint text from AI
  if(bpText){
    children.push(cp([tr('AI Generated Blueprint Details:',11,true,'Times New Roman',C.qNum)],AlignmentType.LEFT,{before:60,after:40}));
    for(const line of bpText.split('\n')){
      const t=line.replace(/\*\*/g,'').trim();
      if(!t){children.push(sp(30));continue;}
      children.push(cp([tr(t,10,false,'Times New Roman',C.black)],AlignmentType.LEFT,{before:20,after:20}));
    }
  }

  const doc = new Document({sections:[{
    properties:{page:{size:{width:11906,height:16838},margin:{top:720,right:900,bottom:720,left:900}}},
    children
  }]});
  return await Packer.toBuffer(doc);
}

app.listen(PORT,()=>console.log(`EMRS QPG v2.0 running on port ${PORT}`));
