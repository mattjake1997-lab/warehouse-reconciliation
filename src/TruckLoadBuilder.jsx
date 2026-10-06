import React, { useState, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";
import * as pdfjsLib from "pdfjs-dist";
import { generateBOL, generateBolNumber, carrierForWarehouse } from "./bolGenerator";
import { estimateWeight, UNRELIABLE_WEIGHT_ITEMS } from "./weightLookup";

// Point pdf.js at its worker. If your app already configures this elsewhere
// (e.g. for the Bills of Lading tab), remove this line to avoid setting it twice.
pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js`;

const TARGET_SKIDS = 26;

// The people who sign these — Matthew is the primary user, so he's the
// default selection, but anyone in this list can be picked per BOL.
const DEFAULT_SIGNATURE_NAMES = [
  "Matthew Jake",
  "Nathan Cobb",
  "Jerold Wilkinson",
  "Derek Rister",
  "Jackie Ori",
  "Shelly Hayhurst",
];

// Parts that can be double-stacked on a single trailer spot — one spot holds
// two skids of these, so they only count as half a spot each toward the
// 26-spot target. Everything else counts as a full spot per skid.
const DOUBLE_STACKABLE_ITEMS = new Set([
  "102838", "117627", "112153", "102109", "101190", "102027",
  "102037", "102049", "102078", "102092", "111844", "116278",
]);

function isDoubleStackable(itemNumber) {
  return DOUBLE_STACKABLE_ITEMS.has(String(itemNumber));
}

function spotsForSkids(itemNumber, skids) {
  return isDoubleStackable(itemNumber) ? skids / 2 : skids;
}

const WAREHOUSE_LABELS = {
  "EVVLIN.WSI": "WSI",
  "EVVLIN.WS2": "WS2",
  "EVVLIN.EAB": "EAB",
};

function warehouseLabel(sub) {
  return WAREHOUSE_LABELS[sub] || sub || "Unknown";
}

// ---------------------------------------------------------------------------
// Excel parsing — offsite availability export (one row per skid/LPN)
// Columns used: Sub, Item, Item Description, LPN. Everything else is ignored.
// ---------------------------------------------------------------------------
function parseAvailabilityWorkbook(workbook) {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: null });

  // itemNumber -> { description, byWarehouse: { WSI: [lpn, lpn...], ... } }
  const byItem = new Map();

  for (const row of rows) {
    const item = row["Item"];
    if (item === null || item === undefined || item === "") continue;
    const itemKey = String(item).trim();
    const sub = row["Sub"];
    const wh = warehouseLabel(sub);
    const description = row["Item Description"] || "";
    const lpn = row["LPN"];

    if (!byItem.has(itemKey)) {
      byItem.set(itemKey, { description, byWarehouse: {} });
    }
    const entry = byItem.get(itemKey);
    if (!entry.byWarehouse[wh]) entry.byWarehouse[wh] = [];
    entry.byWarehouse[wh].push(lpn ?? entry.byWarehouse[wh].length + 1);
  }

  return byItem;
}

// ---------------------------------------------------------------------------
// PDF parsing — "Items Needing Replenishment" export
// Only two things are read from each row: the ITEM number and the EVV Move
// pallet count. Everything else on the row (description, on-site / off-site
// counts, status...) is ignored — item descriptions come from the inventory
// file instead.
// ---------------------------------------------------------------------------

// Pulls the text out of a PDF as plain lines (tokens grouped by vertical
// position, read left to right).
async function extractPdfLines(file) {
  const buf = await file.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: buf }).promise;

  const lines = [];
  let textItems = 0;
  for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
    const page = await pdf.getPage(pageNum);
    const content = await page.getTextContent();
    textItems += content.items.filter((it) => it.str && it.str.trim()).length;

    const rows = new Map(); // roundedY -> [{x, str}]
    for (const item of content.items) {
      const y = Math.round(item.transform[5]);
      const x = item.transform[4];
      // group text within 3px of an existing row's y value
      let bucketY = y;
      for (const existingY of rows.keys()) {
        if (Math.abs(existingY - y) <= 3) {
          bucketY = existingY;
          break;
        }
      }
      if (!rows.has(bucketY)) rows.set(bucketY, []);
      rows.get(bucketY).push({ x, str: item.str });
    }

    const sortedYs = Array.from(rows.keys()).sort((a, b) => b - a);
    for (const y of sortedYs) {
      const tokens = rows.get(y).sort((a, b) => a.x - b.x);
      const text = tokens.map((t) => t.str).join(" ").replace(/\s+/g, " ").trim();
      if (text) lines.push(text);
    }
  }

  if (textItems === 0) {
    throw new Error(
      "This PDF has no readable text — it's a picture of the table (like a screenshot), so the numbers can't be read reliably. Save or export the report to PDF straight from the source instead."
    );
  }
  return lines;
}

// A data row starts with a 5-6 digit item number. The EVV Move numbers sit at
// the far right of the row, so they're read from the end of the line — that
// keeps stray digits in the description or date columns from interfering.
//
// The export's right-hand columns are  EVV Move Qty | EVV Move Pallets, and
// sometimes a trailing Pallet Qty (units per pallet) after them. When that
// trailing column is there, Move Qty ÷ Pallet Qty = Move Pallets, which is
// how the layout is detected (and how a misread row gets caught).
function parseNeedsLines(lines) {
  const rows = [];
  for (const line of lines) {
    const m = line.match(/^(\d{5,6})\s+(.*)$/);
    if (!m) continue;
    const nums = [...m[2].matchAll(/\d[\d,]*/g)].map((n) => Number(n[0].replace(/,/g, "")));
    if (nums.length >= 2) rows.push({ item: m[1], nums });
  }

  const addsUp = (nums) => {
    if (nums.length < 3) return false;
    const [qty, pallets, perPallet] = nums.slice(-3);
    return perPallet > 0 && Math.abs(qty / perPallet - pallets) <= 1;
  };
  const hasPalletQtyColumn = rows.filter((r) => addsUp(r.nums)).length >= rows.length / 2;

  const needs = [];
  const flagged = []; // rows whose numbers didn't line up — worth a second look
  for (const r of rows) {
    const movePallets = hasPalletQtyColumn ? r.nums[r.nums.length - 2] : r.nums[r.nums.length - 1];
    if (!Number.isFinite(movePallets) || movePallets <= 0) continue;

    const implausible = movePallets > 100; // a trailer holds ~26 spots, so this is a misread
    if (implausible || (hasPalletQtyColumn && !addsUp(r.nums))) flagged.push(r.item);
    if (implausible) continue;

    needs.push({ item: r.item, description: "", movePallets });
  }

  // Keep the export's own ordering (highest EVV Move first = build priority).
  needs.sort((a, b) => b.movePallets - a.movePallets);
  return { needs, flagged };
}

async function parseNeedsPdf(file) {
  return parseNeedsLines(await extractPdfLines(file));
}

// Warehouses are tried in this order when picking which single warehouse to
// build the load from — EAB and WS2 before WSI. Between EAB and WS2 there
// was no stated preference, so EAB is tried first here; swap the order below
// if WS2 should go first.
const WAREHOUSE_ORDER = ["EAB", "WS2", "WSI"];

// ---------------------------------------------------------------------------
// Load building — loads never mix warehouses. This builds one candidate load
// per warehouse, using only that warehouse's stock, filling toward
// TARGET_SKIDS trailer spots in needs-list order (already ranked by Move
// Pallets, highest first).
// ---------------------------------------------------------------------------
function buildLoadForWarehouse(needs, availability, warehouse, target = TARGET_SKIDS) {
  const lineItems = [];
  const unmet = [];
  let spotsUsed = 0;

  for (const need of needs) {
    if (spotsUsed >= target) {
      unmet.push({ ...need, reason: "Load already full" });
      continue;
    }

    const avail = availability.get(need.item);
    const countAtWarehouse = avail?.byWarehouse?.[warehouse]?.length || 0;

    if (!avail) {
      unmet.push({ ...need, reason: "Not in offsite inventory" });
      continue;
    }
    if (countAtWarehouse === 0) {
      unmet.push({ ...need, reason: `Not available at ${warehouse}` });
      continue;
    }

    const doubleStack = isDoubleStackable(need.item);
    const spotsRemaining = target - spotsUsed;
    const maxSkidsBySpots = doubleStack
      ? Math.floor(spotsRemaining * 2)
      : Math.floor(spotsRemaining);

    const take = Math.min(countAtWarehouse, need.movePallets, maxSkidsBySpots);

    if (take > 0) {
      lineItems.push({
        id: `${need.item}-${warehouse}`,
        item: need.item,
        description: need.description || avail.description,
        warehouse,
        skids: take,
        availableAtWarehouse: countAtWarehouse,
        needed: need.movePallets,
        doubleStack,
        weight: estimateWeight(need.item, take) ?? "",
        weightIsEstimate: estimateWeight(need.item, take) != null,
      });
      spotsUsed += spotsForSkids(need.item, take);
    }

    if (need.movePallets > countAtWarehouse) {
      unmet.push({
        ...need,
        reason: `Only ${countAtWarehouse} of ${need.movePallets} available at ${warehouse}`,
      });
    } else if (take < need.movePallets) {
      unmet.push({ ...need, reason: "Load already full" });
    }
  }

  return { warehouse, lineItems, unmet, spotsUsed, isFull: spotsUsed >= target };
}

// Builds one candidate per warehouse that actually has any stock at all, and
// picks a default: the first (in priority order) that fills a full truck;
// if none do, the one that gets closest, flagged as not full.
function buildAllCandidates(needs, availability) {
  const warehousesPresent = WAREHOUSE_ORDER.filter((wh) =>
    Array.from(availability.values()).some((entry) => (entry.byWarehouse[wh]?.length || 0) > 0)
  );

  const candidates = warehousesPresent.map((wh) =>
    buildLoadForWarehouse(needs, availability, wh)
  );

  const fullCandidate = candidates.find((c) => c.isFull);
  const defaultCandidate =
    fullCandidate || candidates.reduce((best, c) => (c.spotsUsed > (best?.spotsUsed ?? -1) ? c : best), null);

  return { candidates, defaultWarehouse: defaultCandidate?.warehouse ?? null };
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------
export default function TruckLoadBuilder() {
  const [availability, setAvailability] = useState(null); // Map
  const [needs, setNeeds] = useState(null); // array
  const [needsFlagged, setNeedsFlagged] = useState([]); // item numbers whose pallet count looked off
  const [candidates, setCandidates] = useState([]); // one per warehouse with stock
  const [selectedWarehouse, setSelectedWarehouse] = useState(null);
  const [lineItems, setLineItems] = useState([]);
  const [unmet, setUnmet] = useState([]);
  const [error, setError] = useState(null);
  const [loadingAvail, setLoadingAvail] = useState(false);
  const [loadingNeeds, setLoadingNeeds] = useState(false);
  const [carrierName, setCarrierName] = useState("");
  const [scac, setScac] = useState("");
  const [bolNumber, setBolNumber] = useState("");
  const [generatingBol, setGeneratingBol] = useState(false);

  const [signatureNames, setSignatureNames] = useState(() => {
    try {
      const saved = localStorage.getItem("truckLoadBuilder.signatureNames");
      return saved ? JSON.parse(saved) : DEFAULT_SIGNATURE_NAMES;
    } catch {
      return DEFAULT_SIGNATURE_NAMES;
    }
  });
  const [selectedSignature, setSelectedSignature] = useState("Matthew Jake");
  const [newSignatureName, setNewSignatureName] = useState("");

  const [loadHistory, setLoadHistory] = useState(() => {
    try {
      const saved = localStorage.getItem("truckLoadBuilder.loadHistory");
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  const todayStr = () => new Date().toLocaleDateString("en-US");

  const logFinalizedLoad = (record) => {
    setLoadHistory((prev) => {
      const updated = [record, ...prev].slice(0, 200); // keep the log from growing forever
      try {
        localStorage.setItem("truckLoadBuilder.loadHistory", JSON.stringify(updated));
      } catch {
        // localStorage unavailable — history just won't persist across reloads
      }
      return updated;
    });
  };

  // Same BOL number generated more than once today is almost always a
  // mistake (duplicate booking, or the load changed but the number wasn't
  // regenerated) — not a hard block, since a legitimate re-print is
  // possible, just a confirmation before it happens silently.
  const findTodaysDuplicate = (candidateBolNumber) =>
    loadHistory.find((h) => h.bolNumber === candidateBolNumber && h.date === todayStr());

  const addSignatureName = () => {
    const name = newSignatureName.trim();
    if (!name || signatureNames.includes(name)) return;
    const updated = [...signatureNames, name];
    setSignatureNames(updated);
    setSelectedSignature(name);
    setNewSignatureName("");
    try {
      localStorage.setItem("truckLoadBuilder.signatureNames", JSON.stringify(updated));
    } catch {
      // localStorage unavailable — the list just won't persist across reloads
    }
  };

  const setLineWeight = (id, weight) => {
    setLineItems((prev) =>
      prev.map((li) => (li.id === id ? { ...li, weight, weightIsEstimate: false } : li))
    );
  };

  const handleGenerateBol = async () => {
    const duplicate = findTodaysDuplicate(bolNumber);
    if (duplicate) {
      const proceed = window.confirm(
        `BOL number ${bolNumber} was already generated today (${duplicate.warehouse}, ${duplicate.totalSkids} skids). Generate it again anyway?`
      );
      if (!proceed) return;
    }

    setGeneratingBol(true);
    try {
      await generateBOL({
        warehouse: selectedWarehouse,
        lineItems,
        carrierName,
        scac,
        bolNumber,
        signatureName: selectedSignature,
      });
      logFinalizedLoad({
        date: todayStr(),
        timestamp: new Date().toISOString(),
        warehouse: selectedWarehouse,
        bolNumber,
        carrierName,
        signatureName: selectedSignature,
        totalSkids: skidsUsed,
        totalSpots: Math.ceil(spotsUsed),
        items: lineItems.map((li) => ({ item: li.item, skids: li.skids, weight: li.weight })),
      });
    } catch (err) {
      setError(`Couldn't generate the BOL: ${err.message}`);
    } finally {
      setGeneratingBol(false);
    }
  };

  const skidsUsed = useMemo(
    () => lineItems.reduce((sum, li) => sum + li.skids, 0),
    [lineItems]
  );

  const spotsUsed = useMemo(
    () => lineItems.reduce((sum, li) => sum + spotsForSkids(li.item, li.skids), 0),
    [lineItems]
  );

  const isFullLoad = spotsUsed >= TARGET_SKIDS;

  const selectWarehouse = useCallback(
    (wh, candidateList) => {
      const list = candidateList || candidates;
      const candidate = list.find((c) => c.warehouse === wh);
      if (!candidate) return;
      setSelectedWarehouse(wh);
      setLineItems(candidate.lineItems);
      setUnmet(candidate.unmet);
      setBolNumber(generateBolNumber(candidate.lineItems));
      setCarrierName(carrierForWarehouse(wh));
    },
    [candidates]
  );

  const rebuild = useCallback(
    (needsList, availMap) => {
      if (!needsList || !availMap) return;
      const { candidates: newCandidates, defaultWarehouse } = buildAllCandidates(needsList, availMap);
      setCandidates(newCandidates);
      if (defaultWarehouse) {
        selectWarehouse(defaultWarehouse, newCandidates);
      } else {
        setSelectedWarehouse(null);
        setLineItems([]);
        setUnmet([]);
      }
    },
    [selectWarehouse]
  );

  const handleAvailabilityUpload = async (file) => {
    if (!file) return;
    setError(null);
    setLoadingAvail(true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const parsed = parseAvailabilityWorkbook(wb);
      setAvailability(parsed);
      rebuild(needs, parsed);
    } catch (err) {
      setError(`Couldn't read that Excel file: ${err.message}`);
    } finally {
      setLoadingAvail(false);
    }
  };

  const handleNeedsUpload = async (file) => {
    if (!file) return;
    setError(null);
    setLoadingNeeds(true);
    try {
      const { needs: parsed, flagged } = await parseNeedsPdf(file);
      setNeedsFlagged(flagged);
      if (parsed.length === 0) {
        setNeeds(null);
        setError(
          "Read the PDF's text but couldn't find any rows with an item number followed by an EVV Move pallet count. Check that this is the Items Needing Replenishment export."
        );
        return;
      }
      setNeeds(parsed);
      rebuild(parsed, availability);
    } catch (err) {
      setError(`Couldn't read that PDF: ${err.message}`);
    } finally {
      setLoadingNeeds(false);
    }
  };

  const removeLineItem = (id) => {
    setLineItems((prev) => prev.filter((li) => li.id !== id));
  };

  const adjustSkids = (id, delta) => {
    setLineItems((prev) =>
      prev.map((li) => {
        if (li.id !== id) return li;
        const next = Math.max(0, Math.min(li.availableAtWarehouse, li.skids + delta));
        const updated = { ...li, skids: next };
        if (li.weightIsEstimate) {
          const est = estimateWeight(li.item, next);
          updated.weight = est ?? "";
        }
        return updated;
      })
    );
  };

  const addUnmetItem = (need) => {
    const avail = availability?.get(need.item);
    const countAtWarehouse = avail?.byWarehouse?.[selectedWarehouse]?.length || 0;
    if (!avail || countAtWarehouse === 0) return;
    setLineItems((prev) => [
      ...prev,
      {
        id: `${need.item}-${selectedWarehouse}-${Date.now()}`,
        item: need.item,
        description: need.description || avail.description,
        warehouse: selectedWarehouse,
        skids: Math.min(1, countAtWarehouse),
        availableAtWarehouse: countAtWarehouse,
        needed: need.movePallets,
        doubleStack: isDoubleStackable(need.item),
        weight: estimateWeight(need.item, Math.min(1, countAtWarehouse)) ?? "",
        weightIsEstimate: estimateWeight(need.item, Math.min(1, countAtWarehouse)) != null,
      },
    ]);
    setUnmet((prev) => prev.filter((u) => u.item !== need.item));
  };

  return (
    <div style={{ fontFamily: "inherit" }}>
      <div
        style={{
          display: "grid",
          gridTemplateColumns: "1fr 1fr",
          gap: 16,
          marginBottom: 16,
        }}
      >
        <UploadCard
          header="Off Site Combined Inv"
          label="Offsite inventory (Excel)"
          hint="Combined WSI / WS2 / EAB export"
          accept=".xlsx,.xls"
          onFile={handleAvailabilityUpload}
          loading={loadingAvail}
          status={
            availability
              ? `${availability.size} parts loaded`
              : "No file uploaded yet"
          }
        />
        <UploadCard
          header="EVV DC PDF"
          label="Needs list (PDF)"
          hint='"Items Needing Replenishment" export'
          accept=".pdf"
          onFile={handleNeedsUpload}
          loading={loadingNeeds}
          status={needs && needs.length ? `${needs.length} items needed` : "No file uploaded yet"}
        />
      </div>

      {candidates.length > 0 && (
        <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
          {candidates.map((c) => (
            <button
              key={c.warehouse}
              onClick={() => selectWarehouse(c.warehouse)}
              style={{
                flex: 1,
                padding: "10px 12px",
                textAlign: "left",
                borderRadius: 8,
                border:
                  c.warehouse === selectedWarehouse
                    ? "2px solid var(--accent)"
                    : "1px solid var(--border-mid)",
                background: "var(--bg-card)",
                color: "var(--text-primary)",
              }}
            >
              <div style={{ fontSize: 13, fontWeight: 500 }}>{c.warehouse}</div>
              <div style={{ fontSize: 12, color: c.isFull ? "var(--green-light)" : "var(--amber-light)" }}>
                {Math.ceil(c.spotsUsed)} / {TARGET_SKIDS} spots
                {c.isFull ? " · full" : ""}
              </div>
            </button>
          ))}
        </div>
      )}

      {error && (
        <div
          style={{
            background: "var(--red-bg)",
            color: "var(--red-light)",
            padding: "10px 14px",
            borderRadius: 8,
            fontSize: 13,
            marginBottom: 16,
            border: "1px solid #5a1515",
          }}
        >
          {error}
        </div>
      )}

      {needsFlagged.length > 0 && (
        <div
          style={{
            background: "var(--amber-bg)",
            color: "var(--amber-light)",
            padding: "10px 14px",
            borderRadius: 8,
            fontSize: 13,
            marginBottom: 16,
            border: "1px solid #5a3d0a",
          }}
        >
          The EVV Move numbers for {needsFlagged.join(", ")} didn't line up with the rest of the row — double-check
          those pallet counts against the PDF before booking.
        </div>
      )}

      {selectedWarehouse && !isFullLoad && (
        <div
          style={{
            background: "var(--amber-bg)",
            color: "var(--amber-light)",
            padding: "10px 14px",
            borderRadius: 8,
            fontSize: 13,
            marginBottom: 16,
            border: "1px solid #5a3d0a",
          }}
        >
          No single warehouse can fill a full truck right now — this is the closest, at{" "}
          {Math.ceil(spotsUsed)} of {TARGET_SKIDS} spots from {selectedWarehouse}.
        </div>
      )}

      <div style={{ display: "flex", gap: 12, marginBottom: 16 }}>
        <StatCard
          label="Trailer spots"
          value={`${Math.ceil(spotsUsed)} / ${TARGET_SKIDS}`}
        />
        <StatCard label="Skids loaded" value={`${skidsUsed}`} />
        <StatCard label="Warehouse" value={selectedWarehouse || "—"} />
        <StatCard
          label="Still needed"
          value={unmet.length ? `${unmet.length} item${unmet.length === 1 ? "" : "s"}` : "None"}
          warn={unmet.length > 0}
        />
      </div>

      {lineItems.length > 0 && (
        <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13, marginBottom: 24 }}>
          <thead>
            <tr style={{ textAlign: "left", color: "var(--text-muted)" }}>
              <th style={cellStyle}>Item</th>
              <th style={cellStyle}>Description</th>
              <th style={cellStyle}>From</th>
              <th style={cellStyle}>Need (EVV Move)</th>
              <th style={cellStyle}>Skids</th>
              <th style={cellStyle}>Spots</th>
              <th style={cellStyle}>Weight</th>
              <th style={cellStyle}></th>
            </tr>
          </thead>
          <tbody>
            {lineItems.map((li) => (
              <tr key={li.id}>
                <td style={cellStyle}>{li.item}</td>
                <td style={cellStyle}>
                  {li.description}
                  {li.doubleStack && (
                    <span
                      style={{
                        marginLeft: 8,
                        fontSize: 11,
                        color: "var(--green-light)",
                        background: "var(--green-bg)",
                        padding: "2px 6px",
                        borderRadius: 4,
                      }}
                    >
                      double-stacks
                    </span>
                  )}
                </td>
                <td style={cellStyle}>{li.warehouse}</td>
                <td style={cellStyle}>{li.needed ?? "—"}</td>
                <td style={cellStyle}>
                  <button onClick={() => adjustSkids(li.id, -1)}>-</button>
                  <span style={{ margin: "0 8px" }}>{li.skids}</span>
                  <button onClick={() => adjustSkids(li.id, 1)}>+</button>
                  <span style={{ color: "var(--text-muted)", marginLeft: 8 }}>
                    / {li.availableAtWarehouse} avail
                  </span>
                </td>
                <td style={cellStyle}>{spotsForSkids(li.item, li.skids)}</td>
                <td style={cellStyle}>
                  <input
                    type="number"
                    min="0"
                    value={li.weight || ""}
                    onChange={(e) => setLineWeight(li.id, e.target.value)}
                    placeholder="lbs"
                    style={{ width: 70 }}
                  />
                  {li.weightIsEstimate && (
                    <span style={{ fontSize: 11, color: "var(--text-muted)", marginLeft: 4 }}>est.</span>
                  )}
                  {UNRELIABLE_WEIGHT_ITEMS.has(String(li.item)) && (
                    <span
                      style={{
                        marginLeft: 4,
                        fontSize: 11,
                        color: "var(--amber-light)",
                        background: "var(--amber-bg)",
                        padding: "1px 5px",
                        borderRadius: 4,
                      }}
                      title="Weight has varied a lot for this part across past loads — double check it"
                    >
                      check
                    </span>
                  )}
                </td>
                <td style={{ ...cellStyle, textAlign: "right" }}>
                  <button onClick={() => removeLineItem(li.id)} aria-label="Remove">
                    ✕
                  </button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}

      {lineItems.length > 0 && (
        <div
          style={{
            background: "var(--bg-card)",
            border: "1px solid var(--border)",
            borderRadius: 12,
            padding: "1rem 1.25rem",
            marginBottom: 24,
          }}
        >
          <p style={{ fontSize: 13, color: "var(--text-secondary)", margin: "0 0 10px" }}>
            Generate Bill of Lading for this load ({selectedWarehouse})
          </p>
          <div style={{ display: "flex", gap: 12, marginBottom: 12, flexWrap: "wrap" }}>
            <input
              placeholder="Carrier name"
              value={carrierName}
              onChange={(e) => setCarrierName(e.target.value)}
            />
            <input placeholder="SCAC" value={scac} onChange={(e) => setScac(e.target.value)} />
            <input
              placeholder="BOL number (auto-filled, editable)"
              value={bolNumber}
              onChange={(e) => setBolNumber(e.target.value)}
            />
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center", marginBottom: 12, flexWrap: "wrap" }}>
            <label style={{ fontSize: 13, color: "var(--text-secondary)" }}>Signature:</label>
            <select
              value={selectedSignature}
              onChange={(e) => setSelectedSignature(e.target.value)}
            >
              <option value="">(none selected)</option>
              {signatureNames.map((name) => (
                <option key={name} value={name}>
                  {name}
                </option>
              ))}
            </select>
            <input
              placeholder="Add a name…"
              value={newSignatureName}
              onChange={(e) => setNewSignatureName(e.target.value)}
              style={{ width: 140 }}
            />
            <button onClick={addSignatureName}>Add</button>
          </div>
          <button onClick={handleGenerateBol} disabled={generatingBol}>
            {generatingBol ? "Generating…" : "Download BOL (.docx)"}
          </button>
          {lineItems.some((li) => !li.weight) && (
            <span style={{ marginLeft: 12, fontSize: 12, color: "var(--amber-light)" }}>
              Some lines have no weight history yet — those cells are blank for manual entry.
            </span>
          )}
        </div>
      )}

      {loadHistory.length > 0 && (
        <div
          style={{
            background: "var(--bg-card)",
            border: "1px solid var(--border)",
            borderRadius: 12,
            padding: "1rem 1.25rem",
            marginBottom: 24,
          }}
        >
          <p style={{ fontSize: 13, color: "var(--text-secondary)", margin: "0 0 10px" }}>
            Load history ({loadHistory.length})
          </p>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <thead>
              <tr style={{ textAlign: "left", color: "var(--text-muted)" }}>
                <th style={cellStyle}>Date</th>
                <th style={cellStyle}>Warehouse</th>
                <th style={cellStyle}>BOL #</th>
                <th style={cellStyle}>Skids</th>
                <th style={cellStyle}>Signed by</th>
              </tr>
            </thead>
            <tbody>
              {loadHistory.slice(0, 20).map((h, i) => (
                <tr key={h.timestamp || i}>
                  <td style={cellStyle}>{h.date}</td>
                  <td style={cellStyle}>{h.warehouse}</td>
                  <td style={cellStyle}>{h.bolNumber}</td>
                  <td style={cellStyle}>{h.totalSkids}</td>
                  <td style={cellStyle}>{h.signatureName || "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
          {loadHistory.length > 20 && (
            <p style={{ fontSize: 12, color: "var(--text-muted)", marginTop: 8 }}>
              Showing the 20 most recent — {loadHistory.length} total logged.
            </p>
          )}
        </div>
      )}

      {unmet.length > 0 && (
        <div>
          <p style={{ fontSize: 13, color: "var(--text-secondary)", marginBottom: 8 }}>
            Not on this load
          </p>
          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13 }}>
            <tbody>
              {unmet.map((u) => (
                <tr key={u.item}>
                  <td style={cellStyle}>{u.item}</td>
                  <td style={cellStyle}>{u.description || availability?.get(u.item)?.description || ""}</td>
                  <td style={{ ...cellStyle, color: "var(--amber-light)" }}>{u.reason}</td>
                  <td style={{ ...cellStyle, textAlign: "right" }}>
                    {(availability?.get(u.item)?.byWarehouse?.[selectedWarehouse]?.length || 0) > 0 && (
                      <button onClick={() => addUnmetItem(u)}>Add to load</button>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

const cellStyle = { padding: "8px 4px", borderBottom: "1px solid var(--border)" };

function UploadCard({ header, label, hint, accept, onFile, loading, status }) {
  const [drag, setDrag] = useState(false);
  const loaded = !!status && status !== "No file uploaded yet";

  const openBrowser = () => {
    const input = document.createElement("input");
    input.type = "file";
    input.accept = accept;
    input.onchange = (ev) => onFile(ev.target.files?.[0]);
    input.click();
  };

  return (
    <div
      style={{
        background: "var(--bg-card)",
        border: "1px solid var(--border)",
        borderRadius: 12,
        padding: "1rem 1.25rem",
      }}
    >
      {header && (
        <p style={{ fontSize: 15, fontWeight: 700, color: "var(--text-primary)", margin: "0 0 6px" }}>
          {header}
        </p>
      )}
      <p style={{ fontSize: 13, color: "var(--text-secondary)", margin: "0 0 4px" }}>{label}</p>
      <p style={{ fontSize: 12, color: "var(--text-muted)", margin: "0 0 12px" }}>{hint}</p>
      <div
        onDragOver={(e) => { e.preventDefault(); setDrag(true); }}
        onDragLeave={() => setDrag(false)}
        onDrop={(e) => {
          e.preventDefault();
          setDrag(false);
          const file = e.dataTransfer.files?.[0];
          if (file) onFile(file);
        }}
        onClick={() => !loading && openBrowser()}
        style={{
          border: `1.5px dashed ${loaded ? "var(--green)" : drag ? "var(--accent-light)" : "var(--border-mid)"}`,
          borderRadius: "var(--radius)",
          background: loaded ? "var(--green-bg)" : drag ? "var(--bg-card-hover)" : "var(--bg-input)",
          padding: "20px 16px",
          textAlign: "center",
          cursor: loading ? "default" : "pointer",
          userSelect: "none",
          transition: "all 0.15s",
        }}
      >
        <div style={{ fontSize: 20, marginBottom: 6, color: loaded ? "var(--green-light)" : "var(--text-muted)" }}>
          {loading ? "⏳" : loaded ? "✓" : "↑"}
        </div>
        <div style={{ fontSize: 13, fontWeight: 600, color: loaded ? "var(--green-light)" : "var(--text-primary)" }}>
          {loading ? "Reading file…" : loaded ? status : "Drop file here, or click to browse"}
        </div>
      </div>
    </div>
  );
}

function StatCard({ label, value, warn }) {
  return (
    <div
      style={{
        flex: 1,
        background: "var(--bg-input)",
        border: "1px solid var(--border)",
        borderRadius: 8,
        padding: "1rem",
        textAlign: "center",
      }}
    >
      <p style={{ fontSize: 13, color: "var(--text-secondary)", margin: "0 0 4px" }}>{label}</p>
      <p
        style={{
          fontSize: 22,
          fontWeight: 500,
          margin: 0,
          color: warn ? "var(--amber-light)" : "var(--text-primary)",
        }}
      >
        {value}
      </p>
    </div>
  );
}
