import React, { useState, useEffect, useCallback, useRef } from "react";
import mammoth from "mammoth";
import * as pdfjsLib from "pdfjs-dist";
import { Upload, Truck, ArrowUpRight, ArrowDownLeft, Trash2, Settings, HelpCircle, X, ChevronDown, ChevronUp, FileWarning } from "lucide-react";

pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js`;

// ---------- style tokens ----------
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Oswald:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&family=Inter:wght@400;500;600;700&display=swap');

.ml-root {
  --ink: #1c2b33;
  --ink-soft: #4a5d66;
  --paper: #eef2f0;
  --paper-raised: #ffffff;
  --line: #c3cdc8;
  --line-soft: #d9e0dc;
  --accent: #b23a2e;
  --accent-soft: #f1ddda;
  --teal: #2f6f62;
  --teal-soft: #dbe9e5;
  --amber: #c07f1f;
  --amber-soft: #f3e4cc;
  font-family: 'Inter', sans-serif;
  color: var(--ink);
  background: var(--paper);
  background-image:
    linear-gradient(var(--line-soft) 1px, transparent 1px);
  background-size: 100% 32px;
  min-height: 100%;
  padding: 28px 20px 60px;
}
.ml-root * { box-sizing: border-box; }
.ml-shell { max-width: 980px; margin: 0 auto; }

.ml-header { display: flex; align-items: flex-end; justify-content: space-between; gap: 16px; margin-bottom: 22px; flex-wrap: wrap; }
.ml-title { font-family: 'Oswald', sans-serif; font-weight: 700; letter-spacing: 0.5px; font-size: 30px; margin: 0; text-transform: uppercase; }
.ml-subtitle { font-size: 13px; color: var(--ink-soft); margin-top: 4px; }
.ml-stamp-corner { font-family: 'IBM Plex Mono', monospace; font-size: 11px; color: var(--ink-soft); text-align: right; line-height: 1.5; }

.ml-panel { background: var(--paper-raised); border: 1px solid var(--line); border-radius: 3px; padding: 18px 20px; margin-bottom: 18px; }
.ml-panel-head { display: flex; align-items: center; justify-content: space-between; cursor: pointer; user-select: none; }
.ml-panel-title { font-family: 'Oswald', sans-serif; text-transform: uppercase; font-size: 13px; letter-spacing: 0.8px; color: var(--ink-soft); display: flex; align-items: center; gap: 8px; }

.ml-config-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-top: 16px; }
@media (max-width: 640px) { .ml-config-grid { grid-template-columns: 1fr; } }
.ml-field label { display: block; font-size: 11px; text-transform: uppercase; letter-spacing: 0.6px; color: var(--ink-soft); margin-bottom: 5px; font-weight: 600; }
.ml-field input[type=text] {
  width: 100%; padding: 9px 10px; border: 1px solid var(--line); border-radius: 2px; font-family: 'IBM Plex Mono', monospace; font-size: 13px; background: var(--paper);
}
.ml-field input[type=text]:focus { outline: 2px solid var(--teal); outline-offset: 1px; }
.ml-hint { font-size: 11.5px; color: var(--ink-soft); margin-top: 4px; }
.ml-checkbox-row { display: flex; align-items: center; gap: 8px; margin-top: 6px; font-size: 12.5px; color: var(--ink-soft); }

.ml-drop {
  border: 2px dashed var(--line);
  border-radius: 4px;
  padding: 30px 20px;
  text-align: center;
  background: var(--paper-raised);
  transition: border-color .15s, background .15s;
  cursor: pointer;
}
.ml-drop.dragging { border-color: var(--teal); background: var(--teal-soft); }
.ml-drop-title { font-family: 'Oswald', sans-serif; font-size: 15px; text-transform: uppercase; letter-spacing: .5px; }
.ml-drop-sub { font-size: 12px; color: var(--ink-soft); margin-top: 6px; }

.ml-summary-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 14px; margin-bottom: 18px; }
@media (max-width: 700px) { .ml-summary-grid { grid-template-columns: 1fr; } }
.ml-summary-card { background: var(--paper-raised); border: 1px solid var(--line); border-radius: 3px; padding: 16px 18px; position: relative; overflow: hidden; }
.ml-summary-label { font-family: 'Oswald', sans-serif; text-transform: uppercase; font-size: 11.5px; letter-spacing: .6px; color: var(--ink-soft); }
.ml-summary-value { font-family: 'IBM Plex Mono', monospace; font-size: 34px; font-weight: 600; margin-top: 4px; line-height: 1; }
.ml-summary-sub { font-size: 12px; color: var(--ink-soft); margin-top: 6px; }
.ml-summary-card.out { border-left: 4px solid var(--accent); }
.ml-summary-card.in { border-left: 4px solid var(--teal); }
.ml-summary-card.truck { border-left: 4px solid var(--amber); }

.ml-entries-head { display: flex; align-items: center; justify-content: space-between; margin-bottom: 10px; }
.ml-count-note { font-size: 12px; color: var(--ink-soft); }
.ml-clear-btn { font-size: 12px; color: var(--accent); background: none; border: none; cursor: pointer; text-decoration: underline; padding: 0; }

.ml-entry { background: var(--paper-raised); border: 1px solid var(--line); border-radius: 3px; margin-bottom: 10px; padding: 14px 16px; }
.ml-entry-top { display: flex; align-items: flex-start; justify-content: space-between; gap: 12px; }
.ml-entry-file { font-family: 'IBM Plex Mono', monospace; font-size: 11px; color: var(--ink-soft); word-break: break-all; }

.ml-stamp {
  font-family: 'IBM Plex Mono', monospace;
  font-weight: 600;
  font-size: 11px;
  letter-spacing: 1px;
  text-transform: uppercase;
  border: 2px solid currentColor;
  border-radius: 3px;
  padding: 3px 9px;
  display: inline-flex;
  align-items: center;
  gap: 5px;
  transform: rotate(-2deg);
  mix-blend-mode: multiply;
  white-space: nowrap;
}
.ml-stamp.out { color: var(--accent); }
.ml-stamp.in { color: var(--teal); }
.ml-stamp.unclear { color: var(--ink-soft); border-style: dashed; }
.ml-stamp.truck { color: var(--amber); transform: rotate(2deg); margin-left: 6px; }

.ml-entry-fields { display: grid; grid-template-columns: 1fr 1fr 1fr auto; gap: 12px; margin-top: 12px; align-items: end; }
@media (max-width: 780px) { .ml-entry-fields { grid-template-columns: 1fr 1fr; } }
.ml-entry-fields label { display: block; font-size: 10.5px; text-transform: uppercase; letter-spacing: .5px; color: var(--ink-soft); margin-bottom: 4px; font-weight: 600; }
.ml-entry-fields input[type=text], .ml-entry-fields input[type=number] {
  width: 100%; padding: 7px 8px; border: 1px solid var(--line); border-radius: 2px; font-size: 12.5px; font-family: 'IBM Plex Mono', monospace;
}
.ml-dir-select { display: flex; gap: 4px; }
.ml-dir-btn { border: 1px solid var(--line); background: var(--paper); padding: 7px 8px; font-size: 11px; cursor: pointer; border-radius: 2px; font-family: 'IBM Plex Mono', monospace; }
.ml-dir-btn.active-out { background: var(--accent-soft); border-color: var(--accent); color: var(--accent); }
.ml-dir-btn.active-in { background: var(--teal-soft); border-color: var(--teal); color: var(--teal); }
.ml-dir-btn.active-unclear { background: var(--line-soft); }
.ml-entry-actions { display: flex; gap: 8px; align-items: center; }
.ml-icon-btn { background: none; border: 1px solid var(--line); border-radius: 2px; padding: 6px; cursor: pointer; color: var(--ink-soft); display: inline-flex; }
.ml-icon-btn:hover { color: var(--accent); border-color: var(--accent); }

.ml-raw-toggle { font-size: 11px; color: var(--teal); background: none; border: none; cursor: pointer; margin-top: 10px; display: inline-flex; align-items: center; gap: 4px; padding: 0; }
.ml-raw-text { margin-top: 8px; max-height: 160px; overflow: auto; background: var(--paper); border: 1px solid var(--line-soft); padding: 8px; font-family: 'IBM Plex Mono', monospace; font-size: 10.5px; white-space: pre-wrap; color: var(--ink-soft); }

.ml-warn-badge { display: inline-flex; align-items: center; gap: 4px; font-size: 11px; color: var(--amber); background: var(--amber-soft); padding: 2px 8px; border-radius: 10px; margin-top: 4px; }

.ml-empty { text-align: center; padding: 40px 20px; color: var(--ink-soft); font-size: 13px; }

.ml-doc-note { display: flex; gap: 10px; align-items: flex-start; background: var(--amber-soft); border: 1px solid var(--amber); border-radius: 3px; padding: 12px 14px; font-size: 12.5px; color: #6b4a12; margin-bottom: 18px; }

.ml-new-month-btn { font-size: 12px; font-weight: 600; color: var(--accent); background: var(--accent-soft); border: 1px solid var(--accent); border-radius: 2px; cursor: pointer; padding: 6px 12px; display: inline-flex; align-items: center; gap: 6px; }
.ml-modal-overlay { position: fixed; inset: 0; background: rgba(28,43,51,0.55); z-index: 999; display: flex; align-items: center; justify-content: center; padding: 16px; }
.ml-modal-box { background: var(--paper-raised); border: 1px solid var(--line); border-radius: 3px; padding: 24px 26px; max-width: 400px; width: 100%; }
.ml-modal-title { font-family: 'Oswald', sans-serif; font-weight: 700; text-transform: uppercase; letter-spacing: 0.5px; font-size: 17px; margin-bottom: 10px; }
.ml-modal-body { font-size: 13px; color: var(--ink-soft); line-height: 1.6; margin-bottom: 20px; }
.ml-modal-actions { display: flex; gap: 10px; }
.ml-modal-btn { flex: 1; padding: 9px; border-radius: 2px; font-size: 13px; font-weight: 600; cursor: pointer; }
.ml-modal-btn.cancel { background: var(--paper); border: 1px solid var(--line); color: var(--ink-soft); }
.ml-modal-btn.confirm { background: var(--accent); border: 1px solid var(--accent); color: #fff; }
`;

// ---------- .doc (legacy binary) best-effort text extraction ----------
function extractLegacyDocText(arrayBuffer) {
  const bytes = new Uint8Array(arrayBuffer);
  let out = "";
  for (let i = 0; i < bytes.length - 1; i += 2) {
    const lo = bytes[i];
    const hi = bytes[i + 1];
    if (hi === 0 && lo >= 32 && lo < 127) {
      out += String.fromCharCode(lo);
    } else if (hi === 0 && (lo === 10 || lo === 13)) {
      out += "\n";
    } else {
      out += " ";
    }
  }
  out = out.replace(/[ \t]{3,}/g, "  ").replace(/\n{2,}/g, "\n");
  return out;
}

async function extractDocxText(arrayBuffer) {
  const result = await mammoth.extractRawText({ arrayBuffer });
  return result.value || "";
}

async function extractPdfText(arrayBuffer) {
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let fullText = "";
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    fullText += content.items.map((item) => item.str).join(" ") + " ";
  }
  return fullText;
}

// ---------- field parsing heuristics ----------
function grabBetween(text, startLabels, endLabels) {
  const upper = text.toUpperCase();
  let startIdx = -1;
  for (const s of startLabels) {
    const idx = upper.indexOf(s);
    if (idx !== -1 && (startIdx === -1 || idx < startIdx)) startIdx = idx + s.length;
  }
  if (startIdx === -1) return "";
  let endIdx = text.length;
  for (const e of endLabels) {
    const idx = upper.indexOf(e, startIdx);
    if (idx !== -1 && idx < endIdx) endIdx = idx;
  }
  return text.slice(startIdx, endIdx).replace(/\s+/g, " ").trim().slice(0, 160);
}

function guessGrandTotalPallets(text) {
  const upper = text.toUpperCase();
  const gtIdx = upper.indexOf("GRAND TOTAL");
  if (gtIdx !== -1) {
    const window = text.slice(gtIdx, gtIdx + 60);
    const nums = window.match(/\d+/g);
    if (nums && nums.length) return parseInt(nums[0], 10);
  }
  // No reliable "Grand Total" anchor found — this happens most often on
  // legacy .doc files, where the crude byte-scan can garble or cut off that
  // label entirely. Guessing at a number here (e.g. "largest number in the
  // document") produced confidently wrong pallet counts in practice — a
  // blank field the person has to fill in is safer than a wrong-looking
  // number that's easy to miss.
  return null;
}

// Canonical addresses by short site code — shared by both the content-based
// address matcher below and the filename-based direction parser.
const SITE_ADDRESS = {
  WSI: "WSI\n1147 Wedeking Ave, Building 1\nEvansville, IN 47711",
  WS2: "WS2\n701 Pennell St\nHenderson, KY 42420",
  EAB: "EAB\n500 N. Second Ave\nEvansville, IN 47710",
  MET: "Metronet\n300 E Walnut St\nEvansville, IN 47713",
};

// Known site addresses — matched anywhere in the extracted text, rather than
// relying on their position relative to "Ship From"/"Ship To" labels. This is
// necessary because the real BOL template positions the Ship From address as
// a Word floating frame, which plain-text extractors (mammoth, and the
// legacy .doc byte scan) pull out of its expected reading-order position —
// it ends up appearing after the *next* cell's label instead of right after
// its own. Address-matching sidesteps that entirely.
const KNOWN_ADDRESSES = [
  { key: "Wedeking", code: "WSI", label: SITE_ADDRESS.WSI },
  { key: "Pennell", code: "WS2", label: SITE_ADDRESS.WS2 },
  { key: "Second Ave", code: "EAB", label: SITE_ADDRESS.EAB },
  { key: "300 E Walnut", code: "MET", label: SITE_ADDRESS.MET },
];

// Staff name BOL files like "8.13.26 Henderson to MET.docx" or
// "WSI to MET.docx" — the filename itself says the direction. That's a much
// more reliable signal than fighting a scrambled document layout, so it's
// checked and cross-referenced against the extracted content rather than
// requiring the direction to be picked by hand.
const FILENAME_SITE_TOKENS = [
  { pattern: /\bWSI\b|\bWEDEKING\b/i, code: "WSI" },
  { pattern: /\bWS2\b|\bHENDERSON\b|\bPENNELL\b/i, code: "WS2" },
  { pattern: /\bEAB\b/i, code: "EAB" },
  { pattern: /\bMET\b|\bMETRONET\b|\bMAIN\b/i, code: "MET" },
];

function parseFilenameDirection(fileName) {
  const normalized = (fileName || "").replace(/\.(docx?|pdf)$/i, "").replace(/[._-]+/g, " ");
  const toMatch = normalized.match(/\bto\b/i);
  if (!toMatch) return null;

  const before = normalized.slice(0, toMatch.index);
  const after = normalized.slice(toMatch.index + toMatch[0].length);
  const findSite = (segment) => {
    for (const t of FILENAME_SITE_TOKENS) {
      if (t.pattern.test(segment)) return t.code;
    }
    return null;
  };

  const siteA = findSite(before);
  const siteB = findSite(after);
  if (!siteA || !siteB || siteA === siteB) return null;
  if (siteA === "MET" && siteB !== "MET") return { direction: "out", code: siteB };
  if (siteB === "MET" && siteA !== "MET") return { direction: "in", code: siteA };
  return null; // neither side is MET — not a pattern this can trust
}

function parseBolText(text) {
  const found = KNOWN_ADDRESSES
    .map((m) => ({ ...m, idx: text.indexOf(m.key) }))
    .filter((m) => m.idx !== -1)
    .sort((a, b) => a.idx - b.idx);

  let shipFrom = found[0] ? found[0].label : "";
  let shipTo = found[1] ? found[1].label : "";

  // Fall back to the old label-position method if none of the known
  // addresses were found (e.g. a new site not in the list yet, or a
  // differently-formatted BOL).
  if (!found.length) {
    shipFrom = grabBetween(text, ["SHIP FROM"], ["BILL OF LADING", "SHIP TO"]);
    shipTo = grabBetween(text, ["SHIP TO"], ["CARRIER NAME", "THIRD PARTY"]);
  }

  // Carrier name sits right after the Ship To address's city/state/zip and
  // before "Third Party Freight Charges" / "SCAC" in this template.
  let carrier = "";
  if (found[1]) {
    const afterAddr = text.slice(found[1].idx);
    const upperAfter = afterAddr.toUpperCase();
    let endIdx = afterAddr.length;
    for (const e of ["THIRD PARTY", "SCAC"]) {
      const i = upperAfter.indexOf(e);
      if (i !== -1 && i < endIdx) endIdx = i;
    }
    const span = afterAddr.slice(0, endIdx);
    const zipMatches = [...span.matchAll(/\d{5}/g)];
    const lastZip = zipMatches[zipMatches.length - 1];
    carrier = lastZip
      ? span.slice(lastZip.index + lastZip[0].length).replace(/\s+/g, " ").trim()
      : "";
  } else {
    carrier = grabBetween(text, ["CARRIER NAME"], ["SCAC", "SPECIAL INSTRUCTIONS", "MASTER BILL"]);
  }

  const pallets = guessGrandTotalPallets(text);
  return { shipFrom, shipTo, carrier, pallets };
}

function classifyDirection(shipFrom, shipTo, homeKeywords) {
  const keys = homeKeywords.split(",").map((k) => k.trim().toUpperCase()).filter(Boolean);
  if (!keys.length) return "unclear";
  const fromHit = keys.some((k) => shipFrom.toUpperCase().includes(k));
  const toHit = keys.some((k) => shipTo.toUpperCase().includes(k));
  if (fromHit && !toHit) return "out";
  if (toHit && !fromHit) return "in";
  return "unclear";
}

function classifyBoxTruck(carrier, boxKeywords, treatBlankAsTruck) {
  const c = (carrier || "").trim();
  if (!c) return !!treatBlankAsTruck;
  const keys = boxKeywords.split(",").map((k) => k.trim().toUpperCase()).filter(Boolean);
  if (!keys.length) return false;
  return keys.some((k) => c.toUpperCase().includes(k));
}

// siteKeywords string format: "Name:keyword|Name:keyword|..."
function parseSiteDefs(siteKeywords) {
  return siteKeywords
    .split("|")
    .map((pair) => {
      const idx = pair.indexOf(":");
      if (idx === -1) return null;
      return { name: pair.slice(0, idx).trim(), keyword: pair.slice(idx + 1).trim() };
    })
    .filter((s) => s && s.name && s.keyword);
}

function identifySite(addressText, siteKeywords) {
  const defs = parseSiteDefs(siteKeywords);
  const upper = (addressText || "").toUpperCase();
  const hit = defs.find((d) => upper.includes(d.keyword.toUpperCase()));
  return hit ? hit.name : null;
}

// ---------- storage helpers ----------
const CONFIG_KEY = "ml-config-v1";
const ENTRIES_KEY = "ml-entries-v1";

export default function ManifestLedger() {
  const [homeKeywords, setHomeKeywords] = useState("Metronet, 300 E Walnut");
  const [siteKeywords, setSiteKeywords] = useState("EAB:500 N. Second Ave|WS2:701 Pennell St|WSI:1147 Wedeking Ave");
  const [boxKeywords, setBoxKeywords] = useState("");
  const [treatBlankAsTruck, setTreatBlankAsTruck] = useState(false);
  const [configOpen, setConfigOpen] = useState(false);
  const [showNewMonth, setShowNewMonth] = useState(false);
  const [entries, setEntries] = useState([]);
  const [dragging, setDragging] = useState(false);
  const [loaded, setLoaded] = useState(false);
  const [expandedRaw, setExpandedRaw] = useState({});
  const fileInputRef = useRef(null);

  // load persisted state
  useEffect(() => {
    try {
      const cfg = localStorage.getItem(CONFIG_KEY);
      if (cfg) {
        const parsed = JSON.parse(cfg);
        if (parsed.homeKeywords !== undefined) setHomeKeywords(parsed.homeKeywords);
        if (parsed.siteKeywords !== undefined) setSiteKeywords(parsed.siteKeywords);
        if (parsed.boxKeywords !== undefined) setBoxKeywords(parsed.boxKeywords);
        if (parsed.treatBlankAsTruck !== undefined) setTreatBlankAsTruck(parsed.treatBlankAsTruck);
      }
    } catch (e) { /* no config saved yet */ }
    try {
      const ent = localStorage.getItem(ENTRIES_KEY);
      if (ent) setEntries(JSON.parse(ent));
    } catch (e) { /* no entries saved yet */ }
    setLoaded(true);
  }, []);

  // persist config
  useEffect(() => {
    if (!loaded) return;
    try {
      localStorage.setItem(CONFIG_KEY, JSON.stringify({ homeKeywords, siteKeywords, boxKeywords, treatBlankAsTruck }));
    } catch (e) { /* storage unavailable */ }
  }, [homeKeywords, siteKeywords, boxKeywords, treatBlankAsTruck, loaded]);

  // persist entries
  useEffect(() => {
    if (!loaded) return;
    try {
      localStorage.setItem(ENTRIES_KEY, JSON.stringify(entries));
    } catch (e) { /* storage unavailable */ }
  }, [entries, loaded]);

  const processFiles = useCallback(async (fileList) => {
    const files = Array.from(fileList);
    for (const file of files) {
      const lower = file.name.toLowerCase();
      if (!lower.endsWith(".doc") && !lower.endsWith(".docx") && !lower.endsWith(".pdf")) continue;
      const buf = await file.arrayBuffer();
      let text = "";
      let extractionMode = "";
      try {
        if (lower.endsWith(".docx")) {
          text = await extractDocxText(buf);
          extractionMode = "docx";
        } else if (lower.endsWith(".pdf")) {
          text = await extractPdfText(buf);
          extractionMode = "pdf";
        } else {
          text = extractLegacyDocText(buf);
          extractionMode = "legacy-doc";
        }
      } catch (e) {
        text = "";
        extractionMode = "failed";
      }
      let { shipFrom, shipTo, carrier, pallets } = parseBolText(text);
      let direction = classifyDirection(shipFrom, shipTo, homeKeywords);
      const filenameHint = parseFilenameDirection(file.name);
      let directionSource = "content";
      let directionMismatch = false;

      if (filenameHint) {
        if (direction === "unclear") {
          // Content extraction couldn't tell — trust the filename instead
          // of leaving this for manual review.
          direction = filenameHint.direction;
          directionSource = "filename";
          // If content extraction also failed to find any address at all,
          // fill in the canonical addresses from the filename too, so there's
          // nothing left to type by hand.
          if (!shipFrom && !shipTo) {
            const siteAddr = SITE_ADDRESS[filenameHint.code];
            if (filenameHint.direction === "in") {
              shipFrom = siteAddr;
              shipTo = SITE_ADDRESS.MET;
            } else {
              shipFrom = SITE_ADDRESS.MET;
              shipTo = siteAddr;
            }
          }
        } else if (direction !== filenameHint.direction) {
          // Both signals exist but disagree — a real discrepancy worth a
          // second look (wrong file attached, mislabeled name, etc.), so this
          // gets flagged rather than silently picking one.
          directionMismatch = true;
        }
      }

      const boxTruck = classifyBoxTruck(carrier, boxKeywords, treatBlankAsTruck);
      const id = (crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random()}`);
      setEntries((prev) => [
        {
          id,
          fileName: file.name,
          shipFrom, shipTo, carrier,
          pallets: pallets === null ? "" : pallets,
          direction, boxTruck,
          directionSource, directionMismatch,
          rawText: text.slice(0, 4000),
          extractionMode,
        },
        ...prev,
      ]);
    }
  }, [homeKeywords, boxKeywords, treatBlankAsTruck]);

  const onDrop = (e) => {
    e.preventDefault();
    setDragging(false);
    if (e.dataTransfer.files && e.dataTransfer.files.length) processFiles(e.dataTransfer.files);
  };

  const updateEntry = (id, patch) => {
    setEntries((prev) => prev.map((en) => (en.id === id ? { ...en, ...patch } : en)));
  };

  const removeEntry = (id) => setEntries((prev) => prev.filter((en) => en.id !== id));
  const clearAll = () => { if (entries.length) setEntries([]); };

  const outCount = entries.filter((e) => e.direction === "out").length;
  const inCount = entries.filter((e) => e.direction === "in").length;
  const truckCount = entries.filter((e) => e.boxTruck).length;
  const outPallets = entries.filter((e) => e.direction === "out").reduce((s, e) => s + (parseInt(e.pallets, 10) || 0), 0);
  const inPallets = entries.filter((e) => e.direction === "in").reduce((s, e) => s + (parseInt(e.pallets, 10) || 0), 0);
  const truckPallets = entries.filter((e) => e.boxTruck).reduce((s, e) => s + (parseInt(e.pallets, 10) || 0), 0);
  const hasLegacyDoc = entries.some((e) => e.extractionMode === "legacy-doc");

  return (
    <div className="ml-root">
      <style>{CSS}</style>
      <div className="ml-shell">
        <div className="ml-header">
          <div>
            <h1 className="ml-title">Monthly BOL Tracker</h1>
            <div className="ml-subtitle">Drop in BOLs — auto-count outbound, inbound, and box-truck runs</div>
          </div>
          <div className="ml-stamp-corner">
            {entries.length} BOL{entries.length === 1 ? "" : "s"} logged
          </div>
        </div>

        {hasLegacyDoc && (
          <div className="ml-doc-note">
            <FileWarning size={16} style={{ flexShrink: 0, marginTop: 2 }} />
            <div>
              <strong>.doc files are read with a best-effort text scan</strong> — old Word 97-2003 files don't have a
              reliable way to extract text in-browser, so fields pulled from them may be rough. Check the highlighted
              entries below and correct any fields by hand. If your system can export as .docx or PDF instead, that
              will parse much more cleanly.
            </div>
          </div>
        )}

        <div className="ml-panel">
          <div className="ml-panel-head" onClick={() => setConfigOpen((o) => !o)}>
            <div className="ml-panel-title"><Settings size={14} /> Settings</div>
            {configOpen ? <ChevronUp size={16} /> : <ChevronDown size={16} />}
          </div>
          {configOpen && (
            <div className="ml-config-grid">
              <div className="ml-field">
                <label>Home / main warehouse address keywords</label>
                <input type="text" value={homeKeywords} onChange={(e) => setHomeKeywords(e.target.value)} placeholder="Metronet, 300 E Walnut" />
                <div className="ml-hint">Comma-separated. If it appears in "Ship From," the BOL counts as shipped off-site. If it appears in "Ship To," it counts as brought back.</div>
              </div>
              <div className="ml-field">
                <label>Known offsite warehouses</label>
                <input type="text" value={siteKeywords} onChange={(e) => setSiteKeywords(e.target.value)} placeholder="Name:address fragment|Name:address fragment" />
                <div className="ml-hint">Format is <code>Name:address fragment</code>, separated by <code>|</code>. Used to label which warehouse each BOL is going to/from — pre-filled with EAB, WS2, and WSI.</div>
              </div>
              <div className="ml-field">
                <label>Box truck carrier keywords</label>
                <input type="text" value={boxKeywords} onChange={(e) => setBoxKeywords(e.target.value)} placeholder="e.g. Metronet Truck, in-house" />
                <div className="ml-hint">Comma-separated text that shows up in the Carrier Name field when you used your own box truck.</div>
                <div className="ml-checkbox-row">
                  <input type="checkbox" id="blank-truck" checked={treatBlankAsTruck} onChange={(e) => setTreatBlankAsTruck(e.target.checked)} />
                  <label htmlFor="blank-truck" style={{ marginBottom: 0, textTransform: "none", fontWeight: 400, letterSpacing: 0 }}>
                    Treat a blank Carrier Name as the box truck
                  </label>
                </div>
              </div>
            </div>
          )}
        </div>

        <div
          className={`ml-drop ${dragging ? "dragging" : ""}`}
          onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
          onDragLeave={() => setDragging(false)}
          onDrop={onDrop}
          onClick={() => fileInputRef.current && fileInputRef.current.click()}
        >
          <Upload size={22} style={{ marginBottom: 8 }} />
          <div className="ml-drop-title">Drop BOL files here, or click to browse</div>
          <div className="ml-drop-sub">.pdf, .docx, and .doc — you can select several at once</div>
          <input
            ref={fileInputRef}
            type="file"
            accept=".pdf,.doc,.docx"
            multiple
            style={{ display: "none" }}
            onChange={(e) => { if (e.target.files) processFiles(e.target.files); e.target.value = ""; }}
          />
        </div>

        <div style={{ height: 22 }} />

        <div className="ml-summary-grid">
          <div className="ml-summary-card out">
            <div className="ml-summary-label"><ArrowUpRight size={12} style={{ verticalAlign: -1 }} /> Shipped off-site</div>
            <div className="ml-summary-value">{outCount}</div>
            <div className="ml-summary-sub">{outPallets} pallets total</div>
          </div>
          <div className="ml-summary-card in">
            <div className="ml-summary-label"><ArrowDownLeft size={12} style={{ verticalAlign: -1 }} /> Brought back</div>
            <div className="ml-summary-value">{inCount}</div>
            <div className="ml-summary-sub">{inPallets} pallets total</div>
          </div>
          <div className="ml-summary-card truck">
            <div className="ml-summary-label"><Truck size={12} style={{ verticalAlign: -1 }} /> Box truck runs</div>
            <div className="ml-summary-value">{truckCount}</div>
            <div className="ml-summary-sub">{truckPallets} pallets total</div>
          </div>
        </div>

        <div className="ml-entries-head">
          <div className="ml-count-note">Every uploaded BOL, most recent first</div>
          {entries.length > 0 && (
            <button className="ml-new-month-btn" onClick={() => setShowNewMonth(true)}>
              <Trash2 size={12} /> Start New Month
            </button>
          )}
        </div>

        {showNewMonth && (
          <div className="ml-modal-overlay">
            <div className="ml-modal-box">
              <div className="ml-modal-title">Start a new month?</div>
              <div className="ml-modal-body">
                This will clear all {entries.length} logged BOL{entries.length === 1 ? "" : "s"} and reset the outbound,
                inbound, and box-truck totals back to zero. This can't be undone.
              </div>
              <div className="ml-modal-actions">
                <button className="ml-modal-btn cancel" onClick={() => setShowNewMonth(false)}>Cancel</button>
                <button className="ml-modal-btn confirm" onClick={() => { clearAll(); setShowNewMonth(false); }}>
                  Clear & start fresh
                </button>
              </div>
            </div>
          </div>
        )}

        {entries.length === 0 && (
          <div className="ml-empty">No BOLs logged yet. Drop files above to get started.</div>
        )}

        {entries.map((en) => (
          <div className="ml-entry" key={en.id}>
            <div className="ml-entry-top">
              <div>
                <div className="ml-entry-file">{en.fileName}</div>
                {en.extractionMode === "legacy-doc" && (
                  <div className="ml-warn-badge"><HelpCircle size={11} /> rough read — please verify</div>
                )}
                {en.extractionMode === "failed" && (
                  <div className="ml-warn-badge"><HelpCircle size={11} /> couldn't read file — fill in by hand</div>
                )}
                {en.directionMismatch && (
                  <div className="ml-warn-badge"><HelpCircle size={11} /> filename says {parseFilenameDirection(en.fileName)?.direction === "in" ? "inbound" : "outbound"}, but the document reads {en.direction === "in" ? "inbound" : en.direction === "out" ? "outbound" : "unclear"} — please verify</div>
                )}
              </div>
              <div style={{ textAlign: "right" }}>
                <div>
                  <span className={`ml-stamp ${en.direction}`}>
                    {en.direction === "out" ? <ArrowUpRight size={12} /> : en.direction === "in" ? <ArrowDownLeft size={12} /> : null}
                    {en.direction === "out" ? "Outbound" : en.direction === "in" ? "Inbound" : "Unclear"}
                  </span>
                  {en.boxTruck && <span className="ml-stamp truck"><Truck size={12} /> Box Truck</span>}
                </div>
                {(() => {
                  const site = identifySite(en.direction === "out" ? en.shipTo : en.shipFrom, siteKeywords);
                  return site ? <div className="ml-entry-file" style={{ marginTop: 4 }}>Warehouse: {site}</div> : null;
                })()}
              </div>
            </div>

            <div className="ml-entry-fields">
              <div>
                <label>Ship From</label>
                <input type="text" value={en.shipFrom} onChange={(e) => updateEntry(en.id, { shipFrom: e.target.value, direction: classifyDirection(e.target.value, en.shipTo, homeKeywords) })} />
              </div>
              <div>
                <label>Ship To</label>
                <input type="text" value={en.shipTo} onChange={(e) => updateEntry(en.id, { shipTo: e.target.value, direction: classifyDirection(en.shipFrom, e.target.value, homeKeywords) })} />
              </div>
              <div>
                <label>Carrier</label>
                <input type="text" value={en.carrier} onChange={(e) => updateEntry(en.id, { carrier: e.target.value, boxTruck: classifyBoxTruck(e.target.value, boxKeywords, treatBlankAsTruck) })} />
              </div>
              <div>
                <label>Pallets</label>
                <input type="number" value={en.pallets} onChange={(e) => updateEntry(en.id, { pallets: e.target.value })} style={{ width: 80 }} />
              </div>
            </div>

            <div className="ml-entry-fields" style={{ marginTop: 8 }}>
              <div>
                <label>Direction (override)</label>
                <div className="ml-dir-select">
                  <button className={`ml-dir-btn ${en.direction === "out" ? "active-out" : ""}`} onClick={() => updateEntry(en.id, { direction: "out" })}>Out</button>
                  <button className={`ml-dir-btn ${en.direction === "in" ? "active-in" : ""}`} onClick={() => updateEntry(en.id, { direction: "in" })}>In</button>
                  <button className={`ml-dir-btn ${en.direction === "unclear" ? "active-unclear" : ""}`} onClick={() => updateEntry(en.id, { direction: "unclear" })}>?</button>
                </div>
              </div>
              <div>
                <label>Box truck (override)</label>
                <div className="ml-dir-select">
                  <button className={`ml-dir-btn ${en.boxTruck ? "active-out" : ""}`} onClick={() => updateEntry(en.id, { boxTruck: !en.boxTruck })}>
                    {en.boxTruck ? "Yes" : "No"}
                  </button>
                </div>
              </div>
              <div />
              <div className="ml-entry-actions">
                <button className="ml-icon-btn" title="Remove" onClick={() => removeEntry(en.id)}><Trash2 size={14} /></button>
              </div>
            </div>

            <button className="ml-raw-toggle" onClick={() => setExpandedRaw((p) => ({ ...p, [en.id]: !p[en.id] }))}>
              {expandedRaw[en.id] ? <ChevronUp size={12} /> : <ChevronDown size={12} />}
              {expandedRaw[en.id] ? "Hide extracted text" : "Show extracted text"}
            </button>
            {expandedRaw[en.id] && <div className="ml-raw-text">{en.rawText || "(nothing extracted)"}</div>}
          </div>
        ))}
      </div>
    </div>
  );
}
