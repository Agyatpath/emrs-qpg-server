const express = require('express');
const cors = require('cors');
const axios = require('axios');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType,
  UnderlineType, PageBreak, VerticalAlign
} = require('docx');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

const PORT = process.env.PORT || 3001;
const OPENROUTER_KEY = process.env.OPENROUTER_KEY || '';

// ── COLORS ───────────────────────────────────────────────────────
const C = {
  qNum:'1C3A6B', qText:'111111', marks:'2E6DA4',
  hindi:'333333', option:'2E6DA4', optTxt:'000000',
  optSep:'AAAAAA', optHi:'444444', answer:'666666',
  black:'000000', footer:'AAAAAA', gold:'D4A017'
};

// ── SCHOOL ────────────────────────────────────────────────────────
const SCH_EN  = 'EKLAVYA MODEL RESIDENTIAL SCHOOL';
const SCH_LOC = 'BANSLA-BAGIDORA, BANSWARA (RAJASTHAN)';
const SCH_HI  = '\u090f\u0915\u0932\u0935\u094d\u092f \u0906\u0926\u0930\u094d\u0936 \u0906\u0935\u093e\u0938\u0940\u092f \u0935\u093f\u0926\u094d\u092f\u093e\u0932\u092f, \u092c\u093e\u0902\u0938\u0932\u093e-\u092c\u093e\u0917\u0940\u0926\u094c\u0930\u093e, \u092c\u093e\u0902\u0938\u0935\u093e\u0921\u093c\u093e (\u0930\u093e\u091c\u0938\u094d\u0925\u093e\u0928)';
const SCH_AFF = 'Affiliated to CBSE, New Delhi  |  CBSE, \u0928\u0908 \u0926\u093f\u0932\u094d\u0932\u0940 \u0938\u0947 \u0938\u0902\u092c\u0926\u094d\u0927';

const SEC_HI = {A:'\u0905',B:'\u092c',C:'\u0938',D:'\u0926',E:'\u0907',F:'\u091c'};
const SEC_NAME_HI = {
  A:'\u0935\u0938\u094d\u0924\u0941\u0928\u093f\u0937\u094d\u0920 \u092a\u094d\u0930\u0936\u094d\u0928',
  B:'\u0905\u0924\u093f \u0932\u0918\u0941 \u0909\u0924\u094d\u0924\u0930\u0940\u092f \u092a\u094d\u0930\u0936\u094d\u0928',
  C:'\u0932\u0918\u0941 \u0909\u0924\u094d\u0924\u0930\u0940\u092f \u092a\u094d\u0930\u0936\u094d\u0928',
  D:'\u0926\u0940\u0930\u094d\u0918 \u0909\u0924\u094d\u0924\u0930\u0940\u092f / \u092a\u094d\u0930\u0915\u0930\u0923 \u0906\u0927\u093e\u0930\u093f\u0924',
  E:'\u0905\u0924\u093f\u0930\u093f\u0915\u094d\u0924', F:'\u0905\u0924\u093f\u0930\u093f\u0915\u094d\u0924'
};

app.get('/', (req,res) => res.json({status:'ok',message:'EMRS QPG Server'}));

// ── FETCH NCERT ───────────────────────────────────────────────────
app.post('/api/fetch-ncert', async (req,res) => {
  const {bookCode,chapterNumbers} = req.body;
  if(!bookCode||!chapterNumbers||!chapterNumbers.length)
    return res.status(400).json({error:'bookCode and chapterNumbers required'});
  let allText='',fetched=0,failed=0;
  for(const chNum of chapterNumbers){
    const padded=String(chNum).padStart(2,'0');
    const url=`https://ncert.nic.in/textbook/pdf/${bookCode}${padded}.pdf`;
    try{
      const r=await axios.get(url,{responseType:'arraybuffer',timeout:20000,
        headers:{'User-Agent':'Mozilla/5.0','Referer':'https://ncert.nic.in/textbook.php'}});
      const pdfParse=require('pdf-parse');
      const d=await pdfParse(Buffer.from(r.data));
      allText+=`\n\n=== CHAPTER ${chNum} ===\n${d.text.replace(/\s+/g,' ').trim().slice(0,8000)}`;
      fetched++;
    }catch(e){failed++;}
  }
  res.json({success:true,text:allText.trim(),fetched,failed,totalChars:allText.length});
});

// ── CALL AI ───────────────────────────────────────────────────────
app.post('/api/generate', async (req,res) => {
  const {prompt} = req.body;
  if(!prompt) return res.status(400).json({error:'prompt required'});
  try{
    const r=await axios.post('https://openrouter.ai/api/v1/chat/completions',{
      model:'openrouter/auto',
      messages:[{role:'user',content:prompt}],
      max_tokens:8192,temperature:0.7
    },{
      headers:{'Authorization':`Bearer ${OPENROUTER_KEY}`,'Content-Type':'application/json',
        'HTTP-Referer':'https://emrs.school','X-Title':'EMRS QPG'},
      timeout:120000
    });
    res.json({success:true,text:r.data.choices[0].message.content});
  }catch(e){
    res.status(500).json({error:e.response?.data?.error?.message||e.message});
  }
});

// ── CREATE DOCX ───────────────────────────────────────────────────
app.post('/api/create-docx', async (req,res) => {
  const {paper,answerKey,blueprint,config} = req.body;
  if(!paper) return res.status(400).json({error:'paper required'});
  try{
    const buf=await buildDocx(paper,answerKey||'',blueprint||[],config||{});
    const fn=[config.cls||'Class',config.sub||'Subject',config.exam||'Exam',config.session||'2026-27']
      .join('_').replace(/\s+/g,'_')+'.docx';
    res.setHeader('Content-Type','application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition',`attachment; filename="${fn}"`);
    res.setHeader('Access-Control-Expose-Headers','Content-Disposition');
    res.send(buf);
  }catch(e){res.status(500).json({error:e.message});}
});

// ══════════════════════════════════════════════════════════════════
// DOCX BUILDER
// ══════════════════════════════════════════════════════════════════
async function buildDocx(paperText, answerKey, blueprint, config) {
  const stream = config.stream?` (${config.stream})`:'';
  const cls    = (config.cls||'')+stream;
  const sub    = config.sub||'';
  const marks  = config.marks||40;
  const dur    = config.duration||'1 Hour 30 Minutes';
  const sess   = config.session||'2026-27';
  const exam   = config.exam||'';
  const W      = 9360;
  const children = [];

  // ── TABLE 0: School Header ────────────────────────────────────
  children.push(new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[W],
    borders:borders(), rows:[new TableRow({children:[new TableCell({
      margins:mg(100),
      children:[
        cp([tr(SCH_EN, 17,true,'Arial',C.black)]),
        cp([tr(SCH_LOC,13,true,'Arial',C.black)]),
        cp([tr(SCH_HI, 13,true,'Mangal',C.black)]),
        cp([tr(SCH_AFF,9.5,false,'Arial',C.black,true)]),
      ]
    })]})]
  }));
  children.push(sp());

  // ── TABLE 1: Exam Info ────────────────────────────────────────
  children.push(new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[3200,3200,2960],
    borders:borders(), rows:[new TableRow({children:[
      new TableCell({width:{size:3200,type:WidthType.DXA},margins:mg(80),children:[
        cp([tr(exam+'  /  '+hiExam(exam),11.5,true,'Arial',C.black)]),
        cp([tr('Academic Session: '+sess+'  /  \u0936\u0948\u0915\u094d\u0937\u0923\u093f\u0915 \u0938\u0924\u094d\u0930: '+sess,9,false,'Arial',C.black)]),
      ]}),
      new TableCell({width:{size:3200,type:WidthType.DXA},margins:mg(80),verticalAlign:VerticalAlign.CENTER,children:[
        cp([tr(sub.toUpperCase(),12,true,'Arial',C.black)],AlignmentType.CENTER),
        cp([tr(hiSub(sub),11,true,'Mangal',C.black)],AlignmentType.CENTER),
      ]}),
      new TableCell({width:{size:2960,type:WidthType.DXA},margins:mg(80),children:[
        cp([tr('Class / \u0915\u0915\u094d\u0937\u093e: '+cls,10.5,true,'Arial',C.black)]),
        cp([tr('Max. Marks / \u0905\u0902\u0915: '+marks,9.5,false,'Arial',C.black)]),
        cp([tr('Time / \u0938\u092e\u092f: '+dur,9.5,false,'Arial',C.black)]),
        ...(config.teacher?[cp([tr('Teacher: '+config.teacher,9,false,'Arial',C.black)])]:[]),
        ...(config.date?[cp([tr('Date: '+config.date,9,false,'Arial',C.black)])]:[]),
      ]}),
    ]})]
  }));
  children.push(sp());

  // ── TABLE 2: Student Info ─────────────────────────────────────
  children.push(new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[4000,2800,2560],
    borders:borders(), rows:[new TableRow({children:[
      new TableCell({margins:mg(70),children:[cp([tr('Name / \u0928\u093e\u092e : __________________________________',10,false,'Arial',C.black)])]}),
      new TableCell({margins:mg(70),children:[cp([tr('Roll No. / \u0915\u094d\u0930\u092e\u093e\u0902\u0915 : _____________',10,false,'Arial',C.black)])]}),
      new TableCell({margins:mg(70),children:[cp([tr('Section / \u0905\u0928\u0941\u092d\u093e\u0917 : ________',10,false,'Arial',C.black)])]}),
    ]})]
  }));
  children.push(sp());

  // ── General Instructions Heading ──────────────────────────────
  children.push(cp([tr('GENERAL INSTRUCTIONS  /  \u0938\u093e\u092e\u093e\u0928\u094d\u092f \u0928\u093f\u0930\u094d\u0926\u0947\u0936',12,true,'Arial',C.qNum)],AlignmentType.CENTER,{before:60,after:60}));
  children.push(sp(40));

  // ── TABLE 3: Instructions bilingual ───────────────────────────
  const INSTRS=[
    ['This paper has FOUR sections: A, B, C and D.','\u0907\u0938 \u092a\u094d\u0930\u0936\u094d\u0928-\u092a\u0924\u094d\u0930 \u092e\u0947\u0902 \u091a\u093e\u0930 \u0916\u0902\u0921 \u0939\u0948\u0902: \u0905, \u092c, \u0938 \u0914\u0930 \u0926\u0964'],
    ['All questions are compulsory. Internal choices are given where indicated.','\u0938\u092d\u0940 \u092a\u094d\u0930\u0936\u094d\u0928 \u0905\u0928\u093f\u0935\u093e\u0930\u094d\u092f \u0939\u0948\u0902\u0964 \u091c\u0939\u093e\u0901 \u0906\u0902\u0924\u0930\u093f\u0915 \u0935\u093f\u0915\u0932\u094d\u092a \u0939\u094b, \u0935\u0939\u093e\u0901 \u0915\u094b\u0908 \u090f\u0915 \u0909\u0924\u094d\u0924\u0930 \u0926\u0947\u0902\u0964'],
    ['Section A \u2013 Objective (MCQ/Fill in Blanks/True-False): 1 mark each.','\u0916\u0902\u0921 \u0905 \u2013 \u0935\u0938\u094d\u0924\u0941\u0928\u093f\u0937\u094d\u0920: 1-1 \u0905\u0902\u0915\u0964'],
    ['Section B \u2013 Very Short Answer: 2 marks each.','\u0916\u0902\u0921 \u092c \u2013 \u0905\u0924\u093f \u0932\u0918\u0941 \u0909\u0924\u094d\u0924\u0930\u0940\u092f: 2-2 \u0905\u0902\u0915\u0964'],
    ['Section C \u2013 Short Answer: 3 marks each.','\u0916\u0902\u0921 \u0938 \u2013 \u0932\u0918\u0941 \u0909\u0924\u094d\u0924\u0930\u0940\u092f: 3-3 \u0905\u0902\u0915\u0964'],
    ['Section D \u2013 Long Answer/Case-based/Map: as marked.','\u0916\u0902\u0921 \u0926 \u2013 \u0926\u0940\u0930\u094d\u0918 \u0909\u0924\u094d\u0924\u0930\u0940\u092f / \u092a\u094d\u0930\u0915\u0930\u0923 \u0906\u0927\u093e\u0930\u093f\u0924: \u0928\u093f\u0930\u094d\u0927\u093e\u0930\u093f\u0924 \u0905\u0902\u0915\u0964'],
  ];
  children.push(new Table({
    width:{size:W,type:WidthType.DXA}, columnWidths:[4680,4680], borders:borders(),
    rows:INSTRS.map((pair,idx)=>new TableRow({children:[
      new TableCell({width:{size:4680,type:WidthType.DXA},margins:mg(60),children:[
        cp([tr(`${idx+1}. `,9.5,true,'Arial',C.black),tr(pair[0],9.5,false,'Arial',C.black)])
      ]}),
      new TableCell({width:{size:4680,type:WidthType.DXA},margins:mg(60),children:[
        cp([tr(`${idx+1}. `,9.5,true,'Arial',C.black),tr(pair[1],9.5,false,'Mangal',C.black)])
      ]}),
    ]}))
  }));
  children.push(sp());

  // ── PARSE PAPER ───────────────────────────────────────────────
  const lines = paperText.split('\n');
  let i=0;
  let inCaseBox=false, caseBoxLines=[];

  const flushCaseBox = ()=>{
    if(caseBoxLines.length===0) return;
    const boxParas = caseBoxLines.map(bl=>{
      if(!bl.trim()) return cp([tr(' ',10,false,'Arial',C.black)],AlignmentType.LEFT,{before:10,after:10});
      const isHdr=!!(bl.match(/^(READ|MAP|Case|Source|PASSAGE)/i));
      const isHindi=isHin(bl);
      return cp([tr(bl,isHdr?10.5:10,isHdr,isHindi?'Mangal':'Arial',isHindi?C.hindi:C.black)],AlignmentType.LEFT,{before:20,after:20});
    });
    children.push(new Table({
      width:{size:W,type:WidthType.DXA}, columnWidths:[W],
      borders:borders(C.black,6),
      rows:[new TableRow({children:[new TableCell({margins:mg(120),children:boxParas})]})]
    }));
    children.push(sp(80));
    caseBoxLines=[];
    inCaseBox=false;
  };

  while(i<lines.length){
    const raw=lines[i];
    const line=raw.replace(/\*\*/g,'').replace(/^#+\s*/,'').trim();

    // Empty line
    if(!line){
      if(inCaseBox) caseBoxLines.push('');
      else children.push(sp(40));
      i++; continue;
    }

    // Skip markdown table separators
    if(line.match(/^\|[-:\s|]+\|$/)){i++;continue;}

    // Skip --- dividers (AI generates these)
    if(line.match(/^-{3,}$/)){i++;continue;}

    // Skip duplicate header lines AI generates (EMRS..., Subject:..., etc.)
    if(line.match(/^(EMRS |Subject:|Maximum Marks:|Chapters Covered:|Unit Test \d - 20|Academic Session)/i)){i++;continue;}

    // Section header
    const secM=line.match(/^SECTION[\s\-–]+([A-F])\b/i);
    if(secM){
      flushCaseBox();
      const L=secM[1].toUpperCase();
      const after=line.replace(/^SECTION[\s\-–]+[A-F]\s*/i,'').trim();
      const bpRow=blueprint.find(b=>b.sec===L);
      const mf=bpRow?`${bpRow.q} \u00d7 ${bpRow.m} = ${bpRow.tot} Marks`:'';
      children.push(new Table({
        width:{size:W,type:WidthType.DXA}, columnWidths:[7200,2160],
        borders:borders(),
        rows:[new TableRow({children:[
          new TableCell({width:{size:7200,type:WidthType.DXA},
            shading:{fill:'F5F0E0',type:ShadingType.CLEAR},margins:mg(80),
            children:[cp([
              tr(`SECTION \u2013 ${L}   ${after}`,12,true,'Arial',C.black),
              tr('   |   ',10,false,'Arial',C.black),
              tr(`\u0916\u0902\u0921 \u2013 ${SEC_HI[L]||L}   ${SEC_NAME_HI[L]||after}`,11,true,'Mangal',C.black),
            ])]}),
          new TableCell({width:{size:2160,type:WidthType.DXA},
            shading:{fill:'F5F0E0',type:ShadingType.CLEAR},margins:mg(80),
            verticalAlign:VerticalAlign.CENTER,
            children:[cp([tr(mf,10,true,'Arial',C.black)],AlignmentType.RIGHT)]}),
        ]})]
      }));
      children.push(sp(80));
      i++;continue;
    }

    // Case/Map/Source box trigger
    if(!inCaseBox && line.match(/^(Read the|Case[\s\-]*(Study|Based)|Source[\s\-]*Based|MAP SKILL|Observe the|Study the passage)/i)){
      flushCaseBox();
      inCaseBox=true;
      caseBoxLines.push(line);
      i++;continue;
    }

    // If inside case box, keep collecting until next Q or SECTION
    if(inCaseBox){
      if(line.match(/^Q\d+[\.\)]/) && !line.match(/^\(([ivx]+)\)/i)){
        flushCaseBox();
      } else {
        // Convert markdown table rows to plain text
        if(line.startsWith('|')){
          const cells=line.split('|').map(s=>s.trim()).filter(Boolean);
          caseBoxLines.push(cells.join('   '));
        } else {
          caseBoxLines.push(line);
        }
        i++;continue;
      }
    }

    // Question line
    const qM=line.match(/^(Q\d+[\.\)])\s*(.*)/);
    if(qM){
      const qNum=qM[1].replace('.','.');
      let rest=qM[2].trim();
      // Extract marks from end: [1 Mark] [2 Marks] [1] [2] etc.
      const mkM=rest.match(/\s*\[(\d+)\s*[Mm]arks?[^\]]*\]\s*$/) || rest.match(/\s*\((\d+)\s*[Mm]arks?\)\s*$/) || rest.match(/\s*\[(\d+)\]\s*$/);
      let marksStr='';
      if(mkM){
        const n=mkM[1];
        marksStr=`[${n} Mark${n==='1'?'':'s'} / ${n} \u0905\u0902\u0915]`;
        rest=rest.slice(0,rest.lastIndexOf(mkM[0])).trim();
      }
      // Build question paragraph
      children.push(cp([
        tr(qNum+'  ',11,true,'Arial',C.qNum),
        tr(rest,11,false,'Arial',C.qText),
        ...(marksStr?[tr('   '+marksStr,10,true,'Arial',C.marks)]:[])
      ],AlignmentType.LEFT,{before:140,after:20}));

      // Check if NEXT non-empty line is Hindi translation of this question
      let j=i+1;
      while(j<lines.length && !lines[j].trim()) j++;
      if(j<lines.length){
        const nl=lines[j].replace(/\*\*/g,'').trim();
        if(nl && isHin(nl) && !nl.match(/^Q\d+[\.\)]/) && !nl.match(/^\([a-d]\)/i) && !nl.match(/^SECTION/i) && !nl.match(/^\(([ivx]+)\)/i)){
          children.push(cp([tr(nl,10,false,'Mangal',C.hindi)],AlignmentType.LEFT,{before:0,after:20}));
          i=j;
        }
      }
      i++;continue;
    }

    // Sub-question (i)(ii)(iii)(iv)
    const subqM=line.match(/^\(([ivxIVX]+)\)\s*(.*)/);
    if(subqM){
      const subRest=subqM[2].trim();
      const mkM2=subRest.match(/\s*\[(\d+)\s*[Mm]arks?[^\]]*\]\s*$/) || subRest.match(/\s*\[(\d+)\]\s*$/);
      let subMarks='',subText=subRest;
      if(mkM2){subMarks=`[${mkM2[1]} \u0905\u0902\u0915]`;subText=subRest.slice(0,subRest.lastIndexOf(mkM2[0])).trim();}
      children.push(cp([
        tr(`(${subqM[1]})  `,10,true,'Arial',C.qNum),
        tr(subText,10.5,false,'Arial',C.qText),
        ...(subMarks?[tr('   '+subMarks,10,true,'Arial',C.marks)]:[])
      ],AlignmentType.LEFT,{before:60,after:20},360));
      i++;continue;
    }

    // Option (a)(b)(c)(d)
    const optM=line.match(/^\(([a-d])\)\s*(.*)/i);
    if(optM){
      const optL=optM[1].toLowerCase();
      const optRest=optM[2].trim();
      // Split on   /   for bilingual
      const parts=optRest.split(/\s{2,}\/\s{2,}|\s+\/\s+(?=[^\u0000-\u007F])|\s+\/\s+(?=[\u0900-\u097F])/);
      const en=parts[0]?parts[0].trim():optRest;
      const hi=parts[1]?parts[1].trim():'';
      children.push(cp([
        tr(`(${optL})  `,10,true,'Arial',C.option),
        tr(en,10,false,'Arial',C.optTxt),
        ...(hi?[tr('   /   ',9,false,'Arial',C.optSep),tr(hi,9.5,false,'Mangal',C.optHi)]:[])
      ],AlignmentType.LEFT,{before:20,after:20},360));
      i++;continue;
    }

    // Markdown table row — convert to simple indented text
    if(line.startsWith('|')){
      const cells=line.split('|').map(s=>s.trim()).filter(Boolean);
      if(cells.length>0){
        children.push(cp([tr(cells.join('   '),10,false,'Arial',C.black)],AlignmentType.LEFT,{before:20,after:20},360));
      }
      i++;continue;
    }

    // OR separator
    if(line==='OR'||line.match(/^OR\s*\/\s*\u0905\u0925\u0935\u093e/i)||line==='OR  /  \u0905\u0925\u0935\u093e'){
      children.push(cp([tr('OR  /  \u0905\u0925\u0935\u093e',10,true,'Arial',C.marks)],AlignmentType.CENTER,{before:80,after:80}));
      i++;continue;
    }

    // Answer blank
    if(line.match(/^Answer\s*[\/|]\s*/i)||line.match(/^Answer\s*:/i)){
      children.push(cp([tr('Answer / \u0909\u0924\u094d\u0924\u0930 : ___________',10,false,'Arial',C.answer)],AlignmentType.LEFT,{before:20,after:60}));
      i++;continue;
    }

    // *** end
    if(line==='***'||line==='* * *'){
      children.push(cp([tr('* * *',12,true,'Arial',C.black)],AlignmentType.CENTER,{before:160,after:80}));
      i++;continue;
    }

    // Part header (Part-I, Part-II)
    if(line.match(/^Part[\-\s]*(I|II|III|IV|V)\b/i)){
      const pm=line.replace(/\s*\(.*?\)\s*$/,'').trim();
      const ps=line.match(/\((.*?)\)/)?.[1]||'';
      children.push(cp([
        tr(pm,10.5,true,'Arial',C.marks),
        ...(ps?[tr('   ('+ps+')',9.5,false,'Arial',C.answer)]:[])
      ],AlignmentType.LEFT,{before:80,after:40}));
      i++;continue;
    }

    // Hindi-only line (translation)
    if(isHin(line)&&!line.match(/^Q\d+/)){
      children.push(cp([tr(line,10,false,'Mangal',C.hindi)],AlignmentType.LEFT,{before:0,after:20}));
      i++;continue;
    }

    // Default
    const isBold=!!(line.match(/^(GENERAL INSTRUCTIONS|Answer each question)/i));
    children.push(cp([tr(line,10.5,isBold,'Arial',C.black)],AlignmentType.LEFT,{before:40,after:40}));
    i++;
  }
  flushCaseBox();

  // ── ANSWER KEY ────────────────────────────────────────────────
  if(answerKey&&answerKey.trim()){
    children.push(new Paragraph({children:[new PageBreak()]}));
    children.push(cp([tr('ANSWER KEY / MARKING SCHEME',14,true,'Arial',C.qNum)],AlignmentType.CENTER,{before:0,after:80}));
    children.push(cp([tr('\u0909\u0924\u094d\u0924\u0930 \u0915\u0941\u0902\u091c\u0940 / \u0905\u0902\u0915\u0928 \u092f\u094b\u091c\u0928\u093e',12,true,'Mangal',C.qNum)],AlignmentType.CENTER,{before:0,after:120}));
    for(const ln of answerKey.split('\n')){
      const t=ln.replace(/\*\*/g,'').trim();
      const isH=!!(t.match(/^(SECTION|Q\d+)/i));
      children.push(cp([tr(t||' ',10.5,isH,'Arial',C.black)],AlignmentType.LEFT,{before:isH?100:30,after:30}));
    }
  }

  // ── MARKS TABLE ───────────────────────────────────────────────
  children.push(sp());
  children.push(cp([tr('Marks Distribution  /  \u0905\u0902\u0915 \u0935\u093f\u0924\u0930\u0923',11,true,'Arial',C.qNum)],AlignmentType.CENTER,{before:120,after:80}));
  if(blueprint&&blueprint.length>0){
    const lbls=blueprint.map(b=>`Section ${b.sec} / \u0916\u0902\u0921 ${SEC_HI[b.sec]||b.sec}`);
    lbls.push('Total / \u0915\u0941\u0932');
    const vals=blueprint.map(b=>String(b.tot));
    vals.push(String(blueprint.reduce((a,b)=>a+b.tot,0)));
    const cw=Math.floor(W/lbls.length);
    children.push(new Table({
      width:{size:W,type:WidthType.DXA}, columnWidths:lbls.map(()=>cw),
      borders:borders(), rows:[
        new TableRow({children:lbls.map(l=>new TableCell({shading:{fill:'F5F0E0',type:ShadingType.CLEAR},margins:mg(60),children:[cp([tr(l,9,true,'Arial',C.black)],AlignmentType.CENTER)]}))}),
        new TableRow({children:vals.map((v,idx)=>new TableCell({margins:mg(60),children:[cp([tr(v,idx===vals.length-1?15:13,true,'Arial',C.black)],AlignmentType.CENTER)]}))})
      ]
    }));
  }

  // ── FOOTER ────────────────────────────────────────────────────
  children.push(sp(120));
  children.push(cp([
    tr('\u2014 \u2014  ',10,false,'Arial',C.footer),
    tr('Best of Luck! / \u0936\u0941\u092d\u0915\u093e\u092e\u0928\u093e\u090f\u0901!',13,true,'Arial',C.qNum),
    tr('  \u2014 \u2014',10,false,'Arial',C.footer),
  ],AlignmentType.CENTER,{before:80,after:40}));
  children.push(cp([tr('Eklavya Model Residential School, Bansla-Bagidora (Banswara)',8,false,'Arial',C.footer)],AlignmentType.CENTER));

  const doc=new Document({sections:[{
    properties:{page:{size:{width:11906,height:16838},margin:{top:720,right:900,bottom:720,left:900}}},
    children
  }]});
  return await Packer.toBuffer(doc);
}

// ── HELPERS ───────────────────────────────────────────────────────
function tr(text,size,bold,font,color,italic){
  return new TextRun({text:String(text),bold:!!bold,italics:!!italic,
    size:Math.round((size||11)*2),font:font||'Arial',color:color||'000000'});
}
function cp(runs,align,spacing,indent){
  return new Paragraph({
    alignment:align||AlignmentType.LEFT,
    spacing:spacing||{before:40,after:40},
    indent:indent?{left:indent}:undefined,
    children:runs
  });
}
function sp(size){return new Paragraph({text:'',spacing:{before:size||60,after:size||60}});}
function mg(v){const m=v||80;return {top:m,bottom:m,left:m+20,right:m+20};}
function borders(color,size){
  const b={style:BorderStyle.SINGLE,size:size||4,color:color||'000000'};
  return {top:b,bottom:b,left:b,right:b,insideHorizontal:b,insideVertical:b};
}
function isHin(t){return /[\u0900-\u097F]/.test(t);}
function hiExam(e){
  const m={'Unit Test 1':'\u0907\u0915\u093e\u0908 \u092a\u0930\u0940\u0915\u094d\u0937\u0923 \u2013 I','Unit Test 2':'\u0907\u0915\u093e\u0908 \u092a\u0930\u0940\u0915\u094d\u0937\u0923 \u2013 II','Unit Test 3':'\u0907\u0915\u093e\u0908 \u092a\u0930\u0940\u0915\u094d\u0937\u0923 \u2013 III','Unit Test 4':'\u0907\u0915\u093e\u0908 \u092a\u0930\u0940\u0915\u094d\u0937\u0923 \u2013 IV','Half-Yearly':'\u0905\u0930\u094d\u0927\u0935\u093e\u0930\u094d\u0937\u093f\u0915','Annual':'\u0935\u093e\u0930\u094d\u0937\u093f\u0915','Annual (Board)':'\u0935\u093e\u0930\u094d\u0937\u093f\u0915 (\u092c\u094b\u0930\u094d\u0921)'};
  return m[e]||e;
}
function hiSub(s){
  const m={'Mathematics':'\u0917\u0923\u093f\u0924','Science':'\u0935\u093f\u091c\u094d\u091e\u093e\u0928','Social Science':'\u0938\u093e\u092e\u093e\u091c\u093f\u0915 \u0935\u093f\u091c\u094d\u091e\u093e\u0928','Hindi':'\u0939\u093f\u0902\u0926\u0940','English':'\u0905\u0902\u0917\u094d\u0930\u0947\u091c\u093c\u0940','Sanskrit':'\u0938\u0902\u0938\u094d\u0915\u0943\u0924','Physics':'\u092d\u094c\u0924\u093f\u0915\u0940','Chemistry':'\u0930\u0938\u093e\u092f\u0928','Biology':'\u091c\u0940\u0935 \u0935\u093f\u091c\u094d\u091e\u093e\u0928','History':'\u0907\u0924\u093f\u0939\u093e\u0938','Geography':'\u092d\u0942\u0917\u094b\u0932','Economics':'\u0905\u0930\u094d\u0925\u0936\u093e\u0938\u094d\u0924\u094d\u0930','Political Science':'\u0930\u093e\u091c\u0928\u0940\u0924\u093f','Hindi Core':'\u0939\u093f\u0902\u0926\u0940','English Core':'\u0905\u0902\u0917\u094d\u0930\u0947\u091c\u093c\u0940','Business Studies':'\u0935\u094d\u092f\u0935\u0938\u093e\u092f \u0905\u0927\u094d\u092f\u092f\u0928','Accountancy':'\u0932\u0947\u0916\u093e\u0936\u093e\u0938\u094d\u0924\u094d\u0930','Physical Education':'\u0936\u093e\u0930\u0940\u0930\u093f\u0915 \u0936\u093f\u0915\u094d\u0937\u093e'};
  return m[s]||s;
}

app.listen(PORT,()=>console.log(`EMRS QPG Server running on port ${PORT}`));
