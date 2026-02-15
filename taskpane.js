/* global Office, Word */

let GLOSSARY = null; // { direct: {}, xref: {}, paraTexts: [] }

Office.onReady(() => {
  const refreshBtn = document.getElementById("refresh");
  if (refreshBtn) refreshBtn.addEventListener("click", buildGlossary);

  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    onSelectionChanged
  );

  // Auto-build glossary when add-in opens
  setStatus("Building glossary…");
  buildGlossary();
});

async function buildGlossary() {
  try {
    setStatus("Scanning document for definitions…");

    await Word.run(async (context) => {
      const paras = context.document.body.paragraphs;
      paras.load("items/text");
      await context.sync();

      const paraTexts = paras.items.map((p) => p.text || "");
      const { direct, xref } = extractDefsFromParagraphs(paraTexts);

      GLOSSARY = { direct, xref, paraTexts };
    });

    setUI("Ready", "Select a defined term in the document.");
    setStatus("Glossary ready ✓");
  } catch (e) {
    console.error("buildGlossary failed", e);
    setStatus("Error building glossary — check console.");
  }
}

async function onSelectionChanged() {
  try {
    if (!GLOSSARY) {
      setStatus("Glossary not ready yet…");
      return;
    }

    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      const rawSelected = cleanText(selection.text || "");
      const selectedKey = normalizeTerm(rawSelected);

      if (!selectedKey) {
        setStatus("No term selected.");
        return;
      }

      setStatus(`Looking up "${rawSelected}"…`);
      const candidates = unique([selectedKey, singularize(selectedKey)].filter(Boolean));

      for (const key of candidates) {
        if (GLOSSARY.direct[key]) {
          setUI(rawSelected, GLOSSARY.direct[key]);
          setStatus("Definition found ✓");
          return;
        }
      }

      for (const key of candidates) {
        if (GLOSSARY.xref[key]) {
          const clauseRef = GLOSSARY.xref[key].clauseRef;
          const hit = findEmbeddedDefinitionParagraphAndExtract(GLOSSARY.paraTexts, key);

          if (hit) {
            setUI(rawSelected, hit);
            setStatus(`Definition found via clause ${clauseRef} ✓`);
          } else {
            setUI(rawSelected, "No embedded definition located.");
            setStatus(`Cross-reference found (clause ${clauseRef}) but definition not extracted`);
          }
          return;
        }
      }

      setUI(rawSelected, "No definition found.");
      setStatus("No definition found.");
    });
  } catch (e) {
    console.error("onSelectionChanged failed", e);
    setStatus("Error during lookup — check console.");
  }
}

function setStatus(msg) {
  const el = document.getElementById("status");
  if (!el) return;
  el.textContent = msg || "";
}

function setUI(term, definition) {
  const termEl = document.getElementById("term");
  const defEl = document.getElementById("definition");
  if (termEl) termEl.textContent = term || "—";
  if (defEl) defEl.textContent = definition || "—";
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

    // Prefer structured legal patterns inside the parenthetical
    const fromParen = extractFromParentheticalItself(inside, termKey);
    if (fromParen) return fromParen;

    // Otherwise, use the text before the parenthetical
    const before = cleanText(paragraph.slice(0, m.index).trim());

    // If the parenthetical starts with "any such amount" / "such amount", use a wider slice
    if (/^(any\s+such\s+amount|such\s+amount|any\s+amount)\b/i.test(inside)) {
      const wider = extractAmountsReferent(before);
      if (wider) return wider;
    }

    // Fallback heuristic
    const phrase = extractNearestPhrase(before);
    if (phrase && phrase.length >= 20) return phrase;

    const sent = lastSentence(before);
    return sent || null;
  }

  return null;
}
function extractFromParentheticalItself(parenText, termKey) {
  const t = cleanText(parenText);

  // Pattern: "<X> being the "Term""
  // Example: "the amount by which ... being the "Clawback Amount""
  let m = t.match(/^(.*)\bbeing\s+the\s+"[^"]+"\s*$/i);
  if (m && m[1]) {
    const candidate = cleanText(m[1]);
    if (candidate.length >= 15) return candidate;
  }

  // Pattern: "<X> being the Term" (no quotes)
  m = t.match(/^(.*)\bbeing\s+the\s+(.+)\s*$/i);
  if (m && m[1] && normalizeTerm(m[2] || "").includes(termKey)) {
    const candidate = cleanText(m[1]);
    if (candidate.length >= 15) return candidate;
  }

  return null;
}

function extractAmountsReferent(beforeText) {
  // Aim: pull the chunk that defines the "amount(s)" referred to by
  // "any such amount" / "such amount" parentheticals.

  const b = cleanText(beforeText);

  // Look for the last occurrence of “such amount(s)”
  const idx = b.toLowerCase().lastIndexOf("such amounts");
  const idx2 = b.toLowerCase().lastIndexOf("such amount");

  const startIdx = Math.max(idx, idx2);
  if (startIdx !== -1) {
    const candidate = b.slice(startIdx).trim();
    // Remove leading “such amount(s) as” → convert to a more helpful phrase
    return candidate
      .replace(/^such\s+amounts?\s+as\s+/i, "amounts ")
      .replace(/^such\s+amounts?\s+/i, "amounts ");
  }

  // Otherwise, fall back to last sentence (better than grabbing the tail)
  return lastSentence(b);
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
