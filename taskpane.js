/* global Office, Word */

let GLOSSARY = null; // { direct: {}, xref: {}, paraTexts: [] }

Office.onReady(() => {
  const refreshBtn = document.getElementById("refresh");
  if (refreshBtn) refreshBtn.addEventListener("click", buildGlossary);

  Office.context.document.addHandlerAsync(
    Office.EventType.DocumentSelectionChanged,
    onSelectionChanged
  );

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

      // 1) Direct definition
      for (const key of candidates) {
        if (GLOSSARY.direct[key]) {
          setUI(rawSelected, GLOSSARY.direct[key]);
          setStatus("Definition found ✓");
          return;
        }
      }

      // 2) Cross-reference
      for (const key of candidates) {
        if (GLOSSARY.xref[key]) {
          const clauseRef = GLOSSARY.xref[key].clauseRef;
          // Fix 2: pass clauseRef so search starts at the correct clause paragraph
          const hit = findEmbeddedDefinitionParagraphAndExtract(GLOSSARY.paraTexts, key, clauseRef);

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
    /^"([^"]+)"\s+has the meaning\s+(given in|given to the term in|set out in|set forth in)\s+clause\s+([0-9]+(?:\.[0-9]+)*)\b.*\s*[.;:]?\s*$/i;
  // Fix 1: fallback for definitions without a "means" keyword
  const fallbackDirectRe =
    /^"([^"]+)"\s+(?!has\s+the\s+meaning\b)(.{15,}?)\s*[.;]?\s*$/i;

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

    // Fix 1: fallback — bare "Term" <definition> lines (no "means")
    m = t.match(fallbackDirectRe);
    if (m) {
      direct[normalizeTerm(m[1])] = m[2].trim();
      continue;
    }
  }

  // Phase 2: inline body definitions — capture (the "Term") and
  // (together, the "Term") patterns that appear in body text, not in the
  // definitions section (e.g. "Departing GP", "Major Matters", "GP").
  // When the parenthetical starts the paragraph (empty before-text), fall back
  // to the preceding paragraph's tail for context.
  const inlineBodyRe = /\((?:together,\s+)?the\s+"([A-Z][^"]{1,60})"\)/g;
  let prevInlinePara = "";
  for (const raw of paragraphTexts) {
    const t = cleanText(raw);
    if (!t) { prevInlinePara = ""; continue; }

    for (const match of t.matchAll(inlineBodyRe)) {
      const termName = match[1].trim();
      const key = normalizeTerm(termName);
      if (!key || direct[key] || xref[key]) continue; // don't overwrite

      let before = t.slice(0, match.index).trim();

      // If no before-text on this line, use the preceding paragraph's tail
      if (before.length < 15 && prevInlinePara.length >= 15) {
        before = prevInlinePara.slice(-400);
      }

      if (before.length < 15) continue; // still not enough context

      // extractNearestPhrase may return a very short fragment (e.g. "123" from
      // a decimal like "234.123") — fall back to lastSentence in that case.
      const nearestPhrase = extractNearestPhrase(before);
      const phrase = (nearestPhrase && nearestPhrase.length >= 15)
        ? nearestPhrase
        : (lastSentence(before) || "");

      if (phrase && phrase.length >= 15) {
        direct[key] = phrase;
      }
    }

    // Also handle "together, the "Term"" at paragraph START where the opening
    // '(' is at the end of the previous paragraph (Word sometimes splits these).
    const splitTogetherRe = /^together,\s+the\s+"([A-Z][^"]{1,60})"\s*[),]/;
    const sm = t.match(splitTogetherRe);
    if (sm) {
      const termName = sm[1].trim();
      const key = normalizeTerm(termName);
      if (key && !direct[key] && !xref[key] && prevInlinePara.length >= 15) {
        const nearestPhrase = extractNearestPhrase(prevInlinePara);
        const phrase = (nearestPhrase && nearestPhrase.length >= 15)
          ? nearestPhrase
          : (lastSentence(prevInlinePara) || "");
        if (phrase && phrase.length >= 15) direct[key] = phrase;
      }
    }

    prevInlinePara = t;
  }

  return { direct, xref };
}

// Fix 2: clause-anchored search with full-document fallback.
// Clause numbers rarely start paragraphs in complex LPAs, so when the
// clause anchor is not found we must still scan the whole document.
function findEmbeddedDefinitionParagraphAndExtract(paragraphTexts, termKey, clauseRef) {
  let startIdx = 0;
  let clauseFound = false;

  if (clauseRef) {
    const clausePat = new RegExp("^" + clauseRef.replace(".", "\\.") + "\\b");
    const clauseIdx = paragraphTexts.findIndex((p) =>
      clausePat.test(cleanText(p || ""))
    );
    if (clauseIdx !== -1) {
      startIdx = clauseIdx;
      clauseFound = true;
    }
  }

  // Search range: from the clause paragraph (or doc start) up to 30 paras,
  // then fall back to the rest of the document.
  const windowEnd = clauseFound
    ? Math.min(startIdx + 30, paragraphTexts.length)
    : paragraphTexts.length;

  // Build the ordered list of paragraph indices to search
  const searchIndices = [];
  for (let i = startIdx; i < windowEnd; i++) searchIndices.push(i);
  if (clauseFound && startIdx > 0) {
    for (let i = 0; i < startIdx; i++) searchIndices.push(i);
  }

  // Two full passes across all candidate paragraphs:
  // Pass 1 — only quoted occurrences (defining parentheticals like (the "Term"))
  // Pass 2 — plain-text occurrences (referencing parentheticals like (excluding Term))
  // This prevents an early plain-text reference in one paragraph from shadowing
  // a quoted defining occurrence in a later paragraph.
  for (const quotedPass of [true, false]) {
    for (const i of searchIndices) {
      const p = cleanText(paragraphTexts[i] || "");
      if (!p.includes("(") || !p.includes(")")) continue;
      const enriched = enrichWithPrecedingContext(p, paragraphTexts, i);
      const extraction = extractMeaningFromParentheticalSentenceForPass(enriched, termKey, quotedPass);
      if (extraction) return extraction;
    }
  }

  return null;
}

// Fix 2b: when a paragraph starts with '(' there is no "before" text on the
// same line. Prepend the tail of the preceding paragraph so extraction has
// something to work with (e.g. "(together, the 'Fund Expenses')" after a list).
function enrichWithPrecedingContext(paragraph, paragraphTexts, idx) {
  if (!paragraph.startsWith("(") || idx === 0) return paragraph;
  const prev = cleanText(paragraphTexts[idx - 1] || "");
  if (!prev) return paragraph;
  // Take last 400 chars of the preceding paragraph to avoid overly long strings
  return prev.slice(-400) + " " + paragraph;
}

// extractMeaningFromParentheticalSentenceForPass handles a single pass (quoted or plain).
// Called from findEmbeddedDefinitionParagraphAndExtract which runs two full-document
// passes so that a quoted defining parenthetical in a later paragraph always wins
// over a plain referencing parenthetical in an earlier paragraph.
function extractMeaningFromParentheticalSentenceForPass(paragraph, termKey, quotedPass) {
  const singKey = singularize(termKey);
  const parenRe = /\(([^)]+)\)/g;
  let m;

  while ((m = parenRe.exec(paragraph)) !== null) {
    const inside     = cleanText(m[1] || "");
    const insideNorm = normalizeTerm(inside);

    // Fix 4: word-boundary matching to prevent substring false positives
    const matches =
      termMatchesWord(insideNorm, termKey) ||
      (singKey && singKey !== termKey && termMatchesWord(insideNorm, singKey));

    if (!matches) continue;

    // Detect whether the term appears in quotes inside this parenthetical
    const isQuotedOccurrence =
      insideNorm.includes('"' + termKey) ||
      (singKey && singKey !== termKey && insideNorm.includes('"' + singKey));

    if (quotedPass && !isQuotedOccurrence) continue;
    if (!quotedPass && isQuotedOccurrence) continue;

    const fromParen = extractFromParentheticalItself(inside, termKey);
    if (fromParen) return fromParen;

    const before = cleanText(paragraph.slice(0, m.index).trim());

    // Skip if before-text is the definition line of a *different* term
    if (/^"[^"]+"\s+(means|shall mean|includes)\b/i.test(before)) continue;

    if (/^(any\s+such\s+amount|such\s+amount|any\s+amount)\b/i.test(inside)) {
      const wider = extractAmountsReferent(before);
      if (wider) return wider;
    }

    const phrase = extractNearestPhrase(before);
    if (phrase && phrase.length >= 20) return phrase;

    const sent = lastSentence(before);
    return sent || null;
  }

  return null;
}

// Kept for compatibility — used by the inline body scanner in extractDefsFromParagraphs
function extractMeaningFromParentheticalSentence(paragraph, termKey) {
  return extractMeaningFromParentheticalSentenceForPass(paragraph, termKey, true)
      || extractMeaningFromParentheticalSentenceForPass(paragraph, termKey, false);
}

function extractFromParentheticalItself(parenText, termKey) {
  const t = cleanText(parenText);

  // "<X> being the "Term""
  let m = t.match(/^(.*)\bbeing\s+the\s+"[^"]+"\s*$/i);
  if (m && m[1]) {
    const candidate = cleanText(m[1]);
    if (candidate.length >= 15) return candidate;
  }

  // "<X> being the Term" (no quotes)
  m = t.match(/^(.*)\bbeing\s+the\s+(.+)\s*$/i);
  if (m && m[1] && normalizeTerm(m[2] || "").includes(termKey)) {
    const candidate = cleanText(m[1]);
    if (candidate.length >= 15) return candidate;
  }

  return null;
}

function extractAmountsReferent(beforeText) {
  const b = cleanText(beforeText);

  const idx = b.toLowerCase().lastIndexOf("such amounts");
  const idx2 = b.toLowerCase().lastIndexOf("such amount");

  const startIdx = Math.max(idx, idx2);
  if (startIdx !== -1) {
    const candidate = b.slice(startIdx).trim();
    return candidate
      .replace(/^such\s+amounts?\s+as\s+/i, "amounts ")
      .replace(/^such\s+amounts?\s+/i, "amounts ");
  }

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

function lastSentence(text) {
  const t = cleanText(text);
  if (!t) return null;

  const parts = t.split(/[.?!]\s+/).map((s) => s.trim()).filter(Boolean);
  if (parts.length === 0) return t;
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
    .replace(/\((ies|es|s)\)\s*$/i, "") // Fix 3: strip parenthetical plural suffix e.g. Investment(s)
    .replace(/^["'\s]+|["'\s]+$/g, "")
    .replace(/^[^\w]+|[^\w]+$/g, "")
    .toLowerCase();
}

function unique(arr) {
  return Array.from(new Set(arr));
}

// Fix 4: word-boundary-safe term matching (prevents "investor" matching "investors")
function termMatchesWord(text, termKey) {
  if (!termKey) return false;
  const escaped = termKey.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  return new RegExp("\\b" + escaped + "\\b", "i").test(text);
}

function singularize(term) {
  if (!term) return term;
  if (term.endsWith("ies")) return term.slice(0, -3) + "y";
  if (term.endsWith("ses")) return term.slice(0, -2);
  if (term.endsWith("s") && !term.endsWith("ss")) return term.slice(0, -1);
  return term;
}
