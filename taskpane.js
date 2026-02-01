/* global Office, Word */

let GLOSSARY = null; // { direct: {}, xref: {}, paraTexts: [] }

Office.onReady(() => {
  document.getElementById("refresh").addEventListener("click", buildGlossary);

  // Update when the user changes selection
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
  if (!GLOSSARY) return; // user must click refresh first

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

    // 2) Cross-reference: look for embedded definition paragraph
    for (const key of candidates) {
      if (GLOSSARY.xref[key]) {
        const clauseRef = GLOSSARY.xref[key].clauseRef;
        const hit = findEmbeddedDefinitionParagraphAndExtract(GLOSSARY.paraTexts, key);

        if (hit) {
          setUI(rawSelected, hit);
        } else {
          setUI(rawSelected, `No embedded definition found (cross-ref: clause ${clauseRef}).`);
        }
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

/* =========================================================
  function extractDefsFromParagraphs(paragraphTexts) {
  const direct = {};
  const xref = {};

  const directRe = /^"([^"]+)"\s+(means|shall mean|includes|shall include|has the following meaning)\s+(.+?)\s*[.;:]?\s*$/i;
  const xrefRe = /^"([^"]+)"\s+has the meaning\s+(given in|set out in|set forth in)\s+clause\s+([0-9]+(?:\.[0-9]+)*)\b.*\s*[.;:]?\s*$/i;

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

/* ===========================
   Embedded definition extraction (NEW)
   Return the phrase immediately before (an "Term") rather than whole paragraph
=========================== */

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
  // Find a parenthetical that contains the term
  const parenRe = /\(([^)]+)\)/g;
  let m;

  while ((m = parenRe.exec(paragraph)) !== null) {
    const inside = cleanText(m[1] || "");
    const insideNorm = normalizeTerm(inside);

    const matches =
      insideNorm.includes(termKey) ||
      (singularize(termKey) && insideNorm.includes(singularize(termKey)));

    if (!matches) continue;

    // We have a matching parenthetical at m.index..m.index+m[0].length
    const before = paragraph.slice(0, m.index).trim();

    // Prefer extracting just the definitional phrase nearest the parenthetical:
    // Take from the last strong delimiter (. ; :) OR a helpful cue word (via/as/in the form of)
    const phrase = extractNearestPhrase(before);

    // If phrase is too short / empty, fallback to last sentence
    if (phrase && phrase.length >= 20) {
      return phrase;
    }

    const sent = lastSentence(before);
    return sent || null;
  }

  return null;
}

function extractNearestPhrase(before) {
  const b = before;

  // Cue words (common in LPAs)
  const cues = [" via ", " as ", " in the form of ", " through ", " using "];

  let cuePos = -1;
  for (const cue of cues) {
    const pos = b.toLowerCase().lastIndexOf(cue);
    if (pos > cuePos) cuePos = pos;
  }

  // Strong delimiter position
  const strongDelims = [".", ";", ":"];
  let delimPos = -1;
  for (const d of strongDelims) {
    const pos = b.lastIndexOf(d);
    if (pos > delimPos) delimPos = pos;
  }

  // Start index: prefer the later of (cue end) vs (strong delimiter + 1)
  let start = 0;
  if (delimPos !== -1) start = delimPos + 1;
  if (cuePos !== -1) start = Math.max(start, cuePos); // include cue word, we'll trim it below

  let candidate = b.slice(start).trim();

  // Trim leading cue words like "via", "as"
  candidate = candidate.replace(/^(via|as|through|using)\s+/i, "");

  // If there's a comma very near the start, drop leading clause fragment
  // (helps remove "Limited Partners shall be entitled..." etc.)
  const firstComma = candidate.indexOf(",");
  if (firstComma !== -1 && firstComma < 40) {
    candidate = candidate.slice(firstComma + 1).trim();
  }

  return candidate;
}

function lastSentence(text) {
  const parts = cleanText(text)
    .split(/(?<=[.?!])\s+/)
    .map(s => s.trim())
    .filter(Boolean);
  return parts.length ? parts[parts.length - 1] : null;
}

function truncate(s, maxLen) {
  if (!s) return s;
  if (s.length <= maxLen) return s;
  return s.slice(0, maxLen - 1).trim() + "…";
}

/* ===========================
   Normalisation helpers
=========================== */

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

main();
