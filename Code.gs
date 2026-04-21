// ═══════════════════════════════════════════════════════════════════
// EMRS QUESTION PAPER GENERATOR - Google Apps Script
// Eklavya Model Residential School, Bansla-Bagidora, Banswara
// NESTS - Ministry of Tribal Affairs, Govt. of India
// ═══════════════════════════════════════════════════════════════════

// Replace this with your Render.com URL after deployment
var BACKEND_URL = "https://emrs-qpg-server.onrender.com";

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("EMRS Question Paper Generator")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function fetchNcertContent(data) {
  var url = BACKEND_URL + "/api/fetch-ncert";
  var options = {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify(data),
    muteHttpExceptions: true
  };
  try {
    var response = UrlFetchApp.fetch(url, options);
    return JSON.parse(response.getContentText());
  } catch (e) {
    return { success: false, error: e.toString(), text: "" };
  }
}

function buildBlueprint(cls, sub, exam, marks) {
  var isUT = exam.indexOf("Unit Test") === 0;
  var subL = sub.toLowerCase();
  var isMath = subL.indexOf("math") >= 0;
  var isLang = subL.indexOf("english") >= 0 || subL.indexOf("hindi") >= 0 || subL.indexOf("sanskrit") >= 0;
  if (isUT) {
    if (isLang) return [
      { sec:"A", type:"Reading Comprehension (CBQ)", q:1, m:8, tot:8 },
      { sec:"B", type:"Grammar / Vocabulary (MCQ + Fill in Blanks)", q:6, m:1, tot:6 },
      { sec:"C", type:"Short Answer / Literature — Choice", q:4, m:3, tot:12 },
      { sec:"D", type:"Long Writing / Detailed Answer — Choice", q:2, m:7, tot:14 }
    ];
    return [
      { sec:"A", type:"MCQ / Assertion-Reason / Match Column (CBQ)", q:10, m:1, tot:10 },
      { sec:"B", type:"Fill in Blanks / True-False with Reason", q:5, m:2, tot:10 },
      { sec:"C", type:"Short Answer (SA) — Internal Choice", q:4, m:3, tot:12 },
      { sec:"D", type:"Case Study / Long Answer — Choice", q:2, m:4, tot:8 }
    ];
  }
  if (marks === 100) return [
    { sec:"A", type:"MCQ / Assertion-Reason / Match Column (CBQ)", q:20, m:1, tot:20 },
    { sec:"B", type:"Fill in Blanks / True-False with Justification", q:5, m:2, tot:10 },
    { sec:"C", type:"Very Short Answer (VSA)", q:5, m:2, tot:10 },
    { sec:"D", type:"Short Answer (SA) — Internal Choice", q:8, m:3, tot:24 },
    { sec:"E", type:"Case Study / Source Based (CBQ)", q:3, m:4, tot:12 },
    { sec:"F", type:"Long Answer / Map Based — Choice", q:4, m:6, tot:24 }
  ];
  if (isLang) return [
    { sec:"A", type:"Reading — Factual / Discursive Passages (CBQ)", q:2, m:10, tot:20 },
    { sec:"B", type:"Grammar in Context / Fill in Blanks / Match", q:6, m:2, tot:12 },
    { sec:"C", type:"Writing Skills — Letter / Essay / Notice", q:2, m:6, tot:12 },
    { sec:"D", type:"Literature — Extract Based (CBQ)", q:2, m:4, tot:8 },
    { sec:"E", type:"Literature — Short / Long Answer — Choice", q:4, m:7, tot:28 }
  ];
  if (marks === 70) return [
    { sec:"A", type:"MCQ / Assertion-Reason / Match Column (CBQ)", q:16, m:1, tot:16 },
    { sec:"B", type:"Fill in Blanks / True-False", q:4, m:1, tot:4 },
    { sec:"C", type:"Very Short Answer (VSA)", q:5, m:2, tot:10 },
    { sec:"D", type:"Short Answer (SA) — Choice", q:7, m:3, tot:21 },
    { sec:"E", type:"Case Study / Image / Diagram (CBQ)", q:2, m:4, tot:8 },
    { sec:"F", type:"Long Answer / Derivation — Choice", q:3, m:5, tot:15 }
  ];
  if (isMath) return [
    { sec:"A", type:"MCQ + Assertion-Reason", q:20, m:1, tot:20 },
    { sec:"B", type:"Fill in Blanks / Match Column", q:5, m:1, tot:5 },
    { sec:"C", type:"Very Short Answer (VSA)", q:5, m:2, tot:10 },
    { sec:"D", type:"Short Answer (SA) — Internal Choice", q:6, m:3, tot:18 },
    { sec:"E", type:"Case-Based Integrated (CBQ)", q:3, m:4, tot:12 },
    { sec:"F", type:"Long Answer / Problem — Choice", q:3, m:5, tot:15 }
  ];
  return [
    { sec:"A", type:"MCQ / Assertion-Reason / Match Column (CBQ)", q:18, m:1, tot:18 },
    { sec:"B", type:"Fill in Blanks / True-False with Justification", q:4, m:1, tot:4 },
    { sec:"C", type:"Very Short Answer (VSA)", q:6, m:2, tot:12 },
    { sec:"D", type:"Short Answer (SA) — Choice", q:5, m:3, tot:15 },
    { sec:"E", type:"Case / Source / Map / Image Based (CBQ)", q:3, m:4, tot:12 },
    { sec:"F", type:"Long Answer (LA) — Choice", q:3, m:5, tot:15 }
  ];
}

function buildPrompt(config, ncertText, blueprint) {
  var bpLines = "";
  for (var i = 0; i < blueprint.length; i++) {
    bpLines += "Section " + blueprint[i].sec + ": " + blueprint[i].type + " — " + blueprint[i].q + " x " + blueprint[i].m + "m = " + blueprint[i].tot + "m\n";
  }
  var mediumInstr = "";
  if (config.medium === "english") mediumInstr = "Generate ENTIRE paper in ENGLISH only.";
  else if (config.medium === "hindi") mediumInstr = "Generate ENTIRE paper in HINDI only (Devanagari script).";
  else if (config.bilingualFormat === "below") mediumInstr = "BILINGUAL: Each question FIRST in ENGLISH then below HINDI translation. Format: Q1. [English]\nQ1. [Hindi in Devanagari]";
  else if (config.bilingualFormat === "side") mediumInstr = "BILINGUAL: Left column English | Right column Hindi in Devanagari. Use | separator.";
  else mediumInstr = "BILINGUAL: Complete paper in ENGLISH first, then '--- Hindi Section ---', then complete paper in HINDI.";
  var qtInstr = (config.mode === "manual" && config.selectedQTypes && config.selectedQTypes.length > 0)
    ? "QUESTION TYPES — use ONLY these: " + config.selectedQTypes.join(", ")
    : "QUESTION TYPES — use all CBSE types: MCQ, Assertion-Reason, Match the Column, Fill in the Blanks, True/False with Justification, VSA, SA, LA, Case Study, Source Based, Map Based, Image/Diagram Based, Data Interpretation, Comprehension Based.";
  var diffMap = { easy:"Easy — direct NCERT based.", mixed:"Mixed — 30% easy, 40% application, 30% HOTS.", medium:"Medium — balanced.", hard:"Hard — majority HOTS.", hots:"HOTS Heavy — 70% Higher Order Thinking." };
  var subL = (config.sub || "").toLowerCase();
  var isGeo = subL.indexOf("geography") >= 0 || subL.indexOf("social") >= 0 || subL.indexOf("history") >= 0;
  var isSci = subL.indexOf("science") >= 0 || subL.indexOf("physics") >= 0 || subL.indexOf("chemistry") >= 0 || subL.indexOf("biology") >= 0;
  var stream = config.stream ? " (" + config.stream + ")" : "";
  var p = "You are expert CBSE paper setter for EMRS Bansla-Bagidora, Banswara (NESTS, MoTA, Govt. of India).\n\n";
  p += "Class: " + config.cls + stream + " | Subject: " + config.sub + " | Exam: " + config.exam + " | Marks: " + config.marks + " | Duration: " + config.duration + " | " + config.setNo + " | Session: " + config.session + "\n";
  if (config.teacher) p += "Prepared by: " + config.teacher + "\n";
  if (config.date) p += "Date: " + config.date + "\n";
  p += "Chapters: " + (config.selectedChapters && config.selectedChapters.length > 0 ? config.selectedChapters.join(", ") : config.chaptersText || "Full syllabus") + "\n";
  if (ncertText) p += "\nNCERT CONTENT — generate questions EXCLUSIVELY from this:\n\"\"\"\n" + ncertText.slice(0, 15000) + "\n\"\"\"\n";
  p += "\nMEDIUM: " + mediumInstr + "\n";
  p += "BLUEPRINT (total = " + config.marks + " marks):\n" + bpLines;
  p += qtInstr + "\n";
  p += "DIFFICULTY: " + (diffMap[config.difficulty] || diffMap.mixed) + "\n";
  if (isGeo) p += "Include at least one MAP BASED question.\n";
  if (isSci) p += "Include DIAGRAM BASED questions.\n";
  p += "\nFORMATTING RULES:\n";
  p += "- Number questions Q1. Q2. Q3. throughout\n";
  p += "- Put marks in brackets at END of question on SAME LINE: Q1. Question text here [2]\n";
  p += "- MCQ options: (a) (b) (c) (d) each on new indented line\n";
  p += "- Internal choices: OR on separate centered line\n";
  p += "- Section headers: SECTION A, SECTION B etc. centered\n";
  p += "- Start with GENERAL INSTRUCTIONS (5-6 points numbered)\n";
  p += "- End with: ***\n";
  p += "- NEP 2020: min 50% CBQ, 20% HOTS, 30% recall\n";
  p += "- Connect to tribal/EMRS context where possible\n\n";
  p += "Then write: ===ANSWER KEY===\n";
  p += "Section wise answers. MCQ: Q1. (b) — reason. SA: key points [marks]. LA: outline with marks.\n";
  if (config.special) p += "\nTEACHER INSTRUCTIONS: " + config.special + "\n";
  p += "\nGenerate now.";
  return p;
}

function generatePaper(config) {
  try {
    var ncertText = "";
    if (config.bookCode && config.chapterNumbers && config.chapterNumbers.length > 0) {
      var fetchResult = fetchNcertContent({ bookCode: config.bookCode, chapters: config.chapterNumbers });
      if (fetchResult.success) ncertText = fetchResult.text || "";
    }
    var blueprint = buildBlueprint(config.cls, config.sub, config.exam, config.marks);
    var prompt = buildPrompt(config, ncertText, blueprint);
    var aiResponse = UrlFetchApp.fetch(BACKEND_URL + "/api/generate", {
      method: "POST",
      contentType: "application/json",
      payload: JSON.stringify({ prompt: prompt }),
      muteHttpExceptions: true
    });
    var aiResult = JSON.parse(aiResponse.getContentText());
    if (!aiResult.success) throw new Error(aiResult.error);
    var fullText = aiResult.text;
    var splitIdx = fullText.indexOf("===ANSWER KEY===");
    return {
      success: true,
      paper: splitIdx >= 0 ? fullText.slice(0, splitIdx).trim() : fullText.trim(),
      answerKey: splitIdx >= 0 ? fullText.slice(splitIdx + 16).trim() : "",
      blueprint: blueprint,
      ncertFetched: ncertText.length > 0
    };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function getBackendUrl() {
  return BACKEND_URL;
}
