/* global Office, Word */

let GLOSSARY = null; // { direct: {}, xref: {}, paraTexts: [] }

Office.onReady(() => {
  document.getElementById("refresh").addEventListener("click", buildGlossary);

  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    onSelectionChanged
  );
});

async function buildGlossary() {
  await Word.run(async (context) => {
    const paras = context.document.body.paragraphs;
    paras.load("items/text");
    await context.sync();

    const paraTexts = paras.items.map((p) => p.text || "");
    const { direct, xref } = extractDefsFromParagraphs(paraTexts);

    GLOSSARY = { direct, xref, paraTexts };
  });

  setUI("Glossary built. Now select a term.", "");
}

async function onSelectionChanged() {
  if (!GLOSSARY) return;

  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.load("text");
    await context.sync();

    const rawSelected = cleanText(selection.text || "");
    const selectedKey = normalizeTerm(rawSelected);
    if (!selectedKey) return;

    const candidates = unique([selectedKey, singularize(selectedKey)].filter(Boolean));

    // 1) Direct definition
    for (const key of candidates) {
      if (GLOSSARY.direct[key]) {
        setUI(rawSelected, GLOSSARY.direct[key]);
        return;
      }
    }

    // 2) Cross-reference: find embedded definition paragraph anywhere
    for (const key of candidates) {
      if (GLOSSARY.xref[key]) {
        const clauseRef = GLOSSARY.xref[key].clauseRef;
        const hit = findEmbeddedDefinitionParagraphAndExtract(GLOSSARY.paraTexts, key);

        setUI(
          rawSelected,
          hit ? hit : `No embedded definition found (cross-ref: clause ${clauseRef}).`
        );
        return;
      }
    }

    setUI(rawSelected, "No definition found.");
  });
}

function setUI(term, definition) {
  document.getElementById("term").textContent = term || "—";
  document.getElementById("definition").textContent = definition || "—";
}

/* ===========================
   Helpers (ported from Script Lab)
=========================== */

function extractDefsFromParagraphs(paragraphTexts) {
  const direct = {};
  const xref = {};

  const directRe =
    /^"([^"]+)"\s+(means|shall mean|includes|shall include|has the following meaning)\s+(.+?)\s*[.;:]?\s*$/i;
  const xrefRe =
    /^"([^"]+)"\s+has the meaning\s+(given in|set out in|set forth in)\s+clause\s+([0-9]+(?:\.[0-9]+)*)\b.*\s*[.;:]?\s*$/i;

  for (const raw of paragraphTexts) {
    const t = cleanText(raw);
    if (!t) continue;

    let m = t.match(directRe);
    if (m) {
      direct[normalizeTerm(m[1])] = m[3].trim();
      continue;
    }

    m = t.match(xrefRe);
    if (m) {
      xref[normalizeTerm(m[1])] = { clauseRef: m[3], rawLine: t };
      continue;
    }
  }

  return { direct, xref };
}

function findEmbeddedDefinitionParagraphAndExtract(paragraphTexts, termKey) {
  for (let i = 0; i < paragraphTexts.length; i++) {
    const p = cleanText(paragraphTexts[i] || "");
    if (!p.includes("(") || !p.includes(")")) continue;

    const extraction = extractMeaningFromParentheticalSentence(p, termKey);
    if (extraction) return truncate(extraction, 260);
  }
  return null;
}

function extractMeaningFromParentheticalSentence(paragraph, termKey) {
  const parenRe = /\(([^)]+)\)/g;
  let m;

  while ((m = parenRe.exec(paragraph)) !== null) {
    const inside = cleanText(m[1] || "");
    const insideNorm = normalizeTerm(inside);

    const matches =
      insideNorm.includes(termKey) ||
      (singularize(termKey) && insideNorm.includes(singularize(termKey)));

    if (!matches) continue;

    const before = paragraph.slice(0, m.index).trim();
    const phrase = extractNearestPhrase(before);

    if (phrase && phrase.length >= 20) return phrase;

    const sent = lastSentence(before);
    return sent || null;
  }

  return null;
}

function extractNearestPhrase(before) {
  const b = before;

  const cues = [" via ", " as ", " in the form of ", " through ", " using "];
  let cuePos = -1;
  for (const cue of cues) {
    const pos = b.toLowerCase().lastIndexOf(cue);
    if (pos > cuePos) cuePos = pos;
  }

  const strongDelims = [".", ";", ":"];
  let delimPos = -1;
  for (const d of strongDelims) {
    const pos = b.lastIndexOf(d);
    if (pos > delimPos) delimPos = pos;
  }

  let start = 0;
  if (delimPos !== -1) start = delimPos + 1;
  if (cuePos !== -1) start = Math.max(start, cuePos);

  let candidate = b.slice(start).trim();
  candidate = candidate.replace(/^(via|as|through|using)\s+/i, "");

  const firstComma = candidate.indexOf(",");
  if (firstComma !== -1 && firstComma < 40) {
    candidate = candidate.slice(firstComma + 1).trim();
  }

  return candidate;
}

// Lookbehind-free "last sentence"
function lastSentence(text) {
  const t = cleanText(text);
  if (!t) return null;

  // Split on sentence-ending punctuation followed by space/newline
  const parts = t.split(/[.?!]\s+/).map(s => s.trim()).filter(Boolean);
  if (parts.length === 0) return t;

  // We split *without* keeping punctuation; return the last chunk
  return parts[parts.length - 1];
}

function truncate(s, maxLen) {
  if (!s) return s;
  if (s.length <= maxLen) return s;
  return s.slice(0, maxLen - 1).trim() + "…";
}

function cleanText(s) {
  return (s || "")
    .replace(/[\u0000-\u001F\u007F-\u009F]/g, "")
    .replace(/[“”]/g, '"')
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeTerm(s) {
  return cleanText(s)
    .replace(/^["'\s]+|["'\s]+$/g, "")
    .replace(/^[^\w]+|[^\w]+$/g, "")
    .toLowerCase();
}

function unique(arr) {
  return Array.from(new Set(arr));
}

function singularize(term) {
  if (!term) return term;
  if (term.endsWith("ies")) return term.slice(0, -3) + "y";
  if (term.endsWith("ses")) return term.slice(0, -2);
  if (term.endsWith("s") && !term.endsWith("ss")) return term.slice(0, -1);
  return term;
}
