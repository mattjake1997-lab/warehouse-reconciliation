import React, { useState, useCallback, useMemo } from 'react'
import * as XLSX from 'xlsx'
import PrintReport from './PrintReport.jsx'
import MonthlyBolTracker from './manifest-ledger.jsx'
import TruckLoadBuilder from './TruckLoadBuilder.jsx'

// ─── helpers ────────────────────────────────────────────────────────────────

// ─── EAB tracked parts ────────────────────────────────────────────────────
// Only these part numbers are pulled from the EAB workbook — every other
// sheet in that file is ignored. To track a new part later, just add its
// number to this list (as a string, matching digits only).
const TARGET_EAB_PARTS = new Set([
  '100727', '102552', '102838', '103324', '103604', '104115', '104382',
  '104910', '106371', '107092', '109635', '110606', '113998', '114800',
  '116278', '117627',
])

// ─── Exception reason presets ─────────────────────────────────────────────
const EXCEPTION_REASON_OPTIONS = [
  'Missing Pallet',
  'Load is still on trailer',
  'Load has not been added to Inv',
  'Load has not been brought back to MET Inv',
  'Other',
]

// Selecting one of these reasons means the discrepancy is explained — it
// still shows in the exceptions list, but no longer counts against the
// reconciled percentage. "Missing Pallet" is deliberately NOT in this set:
// it's still a real, unresolved discrepancy. "Other" IS included since it's
// always paired with a typed-in explanation in the notes field.
const EXPLAINED_REASONS = new Set([
  'Load is still on trailer',
  'Load has not been added to Inv',
  'Load has not been brought back to MET Inv',
  'Other',
])

function isLineAccounted(line, warehouse, exceptionCategories) {
  if (line.status === 'ok') return true
  const reasonKey = `${warehouse}-${line.part}`
  return EXPLAINED_REASONS.has((exceptionCategories || {})[reasonKey])
}

function normalizeKey(s) {
  return String(s ?? '').trim().toLowerCase().replace(/\s+/g, ' ')
}

function parseNumber(v) {
  const n = parseFloat(String(v ?? '').replace(/[^0-9.-]/g, ''))
  return isNaN(n) ? 0 : n
}

function parseDate(v) {
  if (!v) return null
  if (v instanceof Date) return v
  if (typeof v === 'number') {
    const d = XLSX.SSF.parse_date_code(v)
    if (d) return new Date(d.y, d.m - 1, d.d)
  }
  const d = new Date(v)
  return isNaN(d.getTime()) ? null : d
}

function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array', cellDates: false })
        const ws = wb.Sheets[wb.SheetNames[0]]
        const rows = XLSX.utils.sheet_to_json(ws, { defval: '' })
        resolve(rows)
      } catch (err) { reject(err) }
    }
    reader.onerror = reject
    reader.readAsArrayBuffer(file)
  })
}

// Special reader for EAB: uses the "Index of Part No.s" sheet as the
// authoritative map of sheet-number -> real part number (each sheet's own
// "Name" column is unreliable), then only reads sheets whose part number is
// in TARGET_EAB_PARTS. Header row is located dynamically per sheet by
// scanning for "Pallets per Space" text, since title rows vary in height.
function readExcelEAB(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array', cellDates: false })

        const indexSheetName = wb.SheetNames.find(n => n.toLowerCase().includes('index of part'))
        const partNumberBySheetIndex = {}
        if (indexSheetName) {
          const idxRows = XLSX.utils.sheet_to_json(wb.Sheets[indexSheetName], { defval: '', header: 1 })
          for (let i = 1; i < idxRows.length; i++) {
            const row = idxRows[i]
            const num = row[0]
            if (num === '' || num === undefined) continue
            const digits = String(row[1] ?? '').replace(/[^0-9]/g, '')
            if (digits) partNumberBySheetIndex[Number(num)] = digits
          }
        }

        const allRows = []
        for (const sheetName of wb.SheetNames) {
          if (sheetName === indexSheetName) continue

          const leadingNumMatch = sheetName.match(/^(\d+)\./)
          if (!leadingNumMatch) continue // not a numbered part sheet (Information, Sheet1, etc.)
          const sheetIndex = parseInt(leadingNumMatch[1])
          const partNumber = partNumberBySheetIndex[sheetIndex]
          if (!partNumber) continue
          if (!TARGET_EAB_PARTS.has(partNumber)) continue // only pull sheets for tracked parts

          const ws = wb.Sheets[sheetName]
          const raw = XLSX.utils.sheet_to_json(ws, { defval: '', header: 1 })
          if (raw.length < 2) continue

          let headerRowIdx = -1
          const scanLimit = Math.min(raw.length, 8)
          for (let i = 0; i < scanLimit; i++) {
            const row = raw[i]
            if (row.some(cell => String(cell ?? '').trim().toLowerCase().includes('pallets per space'))) {
              headerRowIdx = i
              break
            }
          }
          if (headerRowIdx === -1) continue

          const headers = raw[headerRowIdx]
          for (let i = headerRowIdx + 1; i < raw.length; i++) {
            const rowArr = raw[i]
            if (!rowArr.some(v => v !== '')) continue
            const obj = { __sheet__: sheetName, __partNumber__: partNumber }
            headers.forEach((h, idx) => { obj[h] = rowArr[idx] ?? '' })
            allRows.push(obj)
          }
        }
        resolve(allRows)
      } catch (err) { reject(err) }
    }
    reader.onerror = reject
    reader.readAsArrayBuffer(file)
  })
}

function findCol(row, candidates) {
  const keys = Object.keys(row)
  for (const c of candidates) {
    const match = keys.find(k => normalizeKey(k) === normalizeKey(c))
    if (match) return match
  }
  for (const c of candidates) {
    const match = keys.find(k => normalizeKey(k).includes(normalizeKey(c)))
    if (match) return match
  }
  return null
}

// ─── parsers ─────────────────────────────────────────────────────────────────

// Main warehouse: one row per pallet (LPN), destination in Locator column
// Columns: Org, Sub, Locator, Item, Item Description, Rev, Primary U, On-hand,
//          Receiving, Inbound, Ordered Q, Unpacked, Packed, Cost Group, LPN, Loaded, LPN Conte...
function parseMainWarehouse(rows) {
  if (!rows.length) return []
  const s = rows[0]
  const itemCol   = findCol(s, ['item','sku','part number','part no'])
  const descCol   = findCol(s, ['item description','description','desc'])
  const locCol    = findCol(s, ['locator','location','destination','dest','sub'])
  const lpnCol    = findCol(s, ['lpn','pallet id','pallet'])

  // Each row = one pallet (qty = 1 per row), group by item+destination
  const parsed = rows.map((r,i) => {
    const rawPart = String(r[itemCol]??'').trim()
    const rawDest = String(r[locCol]??'').trim().toUpperCase()
    if (!rawPart || rawPart === '0') return null

    // Destination: extract warehouse from locator like RECEIVED.EVVLIN.WS2, RECEIVED.EVVLIN.WSI, RECEIVED.EVVLIN.EAB
    let destination = 'UNKNOWN'
    if (rawDest.includes('WS2')) destination = 'WS2'
    else if (rawDest.includes('WSI')) destination = 'WSI'
    else if (rawDest.includes('EAB')) destination = 'EAB'

    return {
      id: i,
      partNumber: rawPart.toUpperCase(),
      description: String(r[descCol]??'').trim(),
      destination,
      qty: 1, // each row is one pallet
      lpn: String(r[lpnCol]??'').trim(),
      _raw: r,
    }
  }).filter(r => r && r.partNumber)

  return parsed
}

// EAB: part number comes from readExcelEAB's __partNumber__ (looked up via
// the Index of Part No.s sheet), not from the per-row "Name" column, which is
// unreliable. Only counts rows where Date Shipped is empty = still in storage.
function parseEAB(rows) {
  if (!rows.length) return []
  const result = []

  for (const r of rows) {
    const partNumber = r.__partNumber__
    if (!partNumber) continue

    const palletsCol = findCol(r, ['pallets per space', 'pallets/space', 'pallets'])
    const shippedCol = findCol(r, ['date shipped', 'date - return shipment'])

    const palletsRaw = palletsCol ? r[palletsCol] : ''
    const dateShipped = shippedCol ? r[shippedCol] : null

    const qty = parseNumber(palletsRaw)
    if (qty <= 0) continue // skips filler/leftover blank rows with no real pallet data

    // Only count pallets still in storage (Date Shipped is empty)
    if (dateShipped) continue

    result.push({
      partNumber,
      qty,
      sheetName: r.__sheet__,
      dateReceived: r['Date Received'] ?? null,
      _raw: r,
    })
  }
  return result
}

// WSI: one row per pallet. SKU = part number, Units = 1 per row.
// Columns: SKU, Lot, LPN, Units (with trailing space in WSI), Weight, Location, SKU Description (WSI only)
// WS2 columns: SKU, Lot, LPN, Units, Weight, Location (no description)
function parseWSI_WS2(rows) {
  if (!rows.length) return []
  const s = rows[0]
  // Handle "Units " with trailing space by checking all keys
  const skuCol  = findCol(s, ['sku','item','part number','part no'])
  const lpnCol  = findCol(s, ['lpn','pallet id','pallet'])
  const locCol  = findCol(s, ['location','loc','bay'])
  const descCol = findCol(s, ['sku description','description','desc','item description'])

  return rows.map((r,i) => {
    // Normalize keys to handle trailing spaces
    const keys = Object.keys(r)
    const skuKey = keys.find(k => k.trim().toLowerCase() === (skuCol||'sku').toLowerCase()) || skuCol
    const rawPart = String(r[skuKey]??'').trim()
    if (!rawPart || rawPart === '0' || rawPart.toLowerCase() === 'sku') return null
    return {
      id: i,
      partNumber: rawPart.toUpperCase(),
      description: descCol ? String(r[descCol]??'').trim() : '',
      qty: 1, // each row = one pallet
      lpn: lpnCol ? String(r[lpnCol]??'').trim() : '',
      location: locCol ? String(r[locCol]??'').trim() : '',
      _raw: r,
    }
  }).filter(r => r && r.partNumber)
}

// ─── reconciliation ──────────────────────────────────────────────────────────

function reconcileWarehouse(shipped, warehouseRows) {
  const shippedMap = {}
  for (const s of shipped) {
    if (!shippedMap[s.partNumber]) shippedMap[s.partNumber] = { qty: 0, rows: [] }
    shippedMap[s.partNumber].qty += s.qty
    shippedMap[s.partNumber].rows.push(s)
  }
  const warehouseMap = {}
  for (const w of warehouseRows) {
    if (!warehouseMap[w.partNumber]) warehouseMap[w.partNumber] = { qty: 0, rows: [] }
    warehouseMap[w.partNumber].qty += w.qty
    warehouseMap[w.partNumber].rows.push(w)
  }
  const allParts = new Set([...Object.keys(shippedMap), ...Object.keys(warehouseMap)])
  const lines = []
  for (const part of allParts) {
    const s = shippedMap[part]
    const w = warehouseMap[part]
    const shippedQty = s?.qty ?? 0
    const reportedQty = w?.qty ?? 0
    const variance = reportedQty - shippedQty
    let status = 'ok'
    if (!s) status = 'unmatched-warehouse'
    else if (!w) status = 'missing'
    else if (variance !== 0) status = variance > 0 ? 'over' : 'short'
    lines.push({ part, shippedQty, reportedQty, variance, status, shippedRows: s?.rows??[], warehouseRows: w?.rows??[] })
  }
  return lines.sort((a,b) => {
    const o = { missing:0, 'unmatched-warehouse':1, short:2, over:3, ok:4 }
    return (o[a.status]??5)-(o[b.status]??5)
  })
}

// ─── reconciliation % that accounts for ALL exceptions ──────────────────────
// Counts a part as reconciled if it's a true match, OR if its exception has
// been explained with one of the EXPLAINED_REASONS (e.g. "still on trailer").
// warehouse + exceptionCategories are needed to look up each line's reason.
function calcReconciliationPct(lines, warehouse, exceptionCategories) {
  if (!lines.length) return 0
  const matched = lines.filter(l => isLineAccounted(l, warehouse, exceptionCategories)).length
  return Math.round((matched / lines.length) * 100)
}

// ─── components ──────────────────────────────────────────────────────────────

const STATUS_STYLES = {
  ok:                   { label: 'Match',       bg: '#071a0f', color: '#3dba78', border: '#1a4d2e' },
  missing:              { label: 'Missing',      bg: '#1a0808', color: '#f06060', border: '#5a1515' },
  short:                { label: 'Short',        bg: '#1a1205', color: '#e8b84b', border: '#5a3d0a' },
  over:                 { label: 'Over',         bg: '#1a1205', color: '#e8b84b', border: '#5a3d0a' },
  'unmatched-warehouse':{ label: 'Not shipped',  bg: '#150a1f', color: '#b07de0', border: '#4a1f70' },
}

function Badge({ status }) {
  const s = STATUS_STYLES[status] || { label: status, bg: '#1a1a1a', color: '#aaa', border: '#333' }
  return (
    <span style={{
      display: 'inline-block', fontSize: 12, fontWeight: 600, letterSpacing: '0.02em',
      padding: '3px 10px', borderRadius: 5,
      background: s.bg, color: s.color, border: `1px solid ${s.border}`
    }}>{s.label}</span>
  )
}

function StatCard({ label, value, sub, accent, warn }) {
  return (
    <div style={{
      background: 'var(--bg-card)',
      border: `1px solid ${warn ? '#5a1515' : 'var(--border)'}`,
      borderTop: `3px solid ${accent || 'var(--border-mid)'}`,
      borderRadius: 'var(--radius-lg)', padding: '20px 22px',
    }}>
      <div style={{ fontSize: 12, color: 'var(--text-muted)', textTransform: 'uppercase', letterSpacing: '0.1em', marginBottom: 10, fontWeight: 600 }}>{label}</div>
      <div style={{ fontSize: 32, fontWeight: 700, fontFamily: 'var(--mono)', color: warn ? '#f06060' : 'var(--text-primary)', lineHeight: 1 }}>{value}</div>
      {sub && <div style={{ fontSize: 13, color: 'var(--text-secondary)', marginTop: 8 }}>{sub}</div>}
    </div>
  )
}

function UploadZone({ label, subtitle, onFile, loaded, fileName }) {
  const [drag, setDrag] = useState(false)
  const handle = useCallback(file => {
    if (!file) return
    onFile(file)
  }, [onFile])
  return (
    <div
      onDragOver={e => { e.preventDefault(); setDrag(true) }}
      onDragLeave={() => setDrag(false)}
      onDrop={e => { e.preventDefault(); setDrag(false); handle(e.dataTransfer.files[0]) }}
      onClick={() => { const i = document.createElement('input'); i.type='file'; i.accept='.xlsx,.xls,.csv'; i.onchange=ev=>handle(ev.target.files[0]); i.click() }}
      style={{
        border: `1.5px dashed ${loaded ? '#2d7a52' : drag ? 'var(--accent-light)' : 'var(--border-mid)'}`,
        borderRadius: 'var(--radius-lg)',
        background: loaded ? '#071a0f' : drag ? '#0a1525' : 'var(--bg-input)',
        padding: '28px 20px', cursor: 'pointer', transition: 'all 0.15s',
        textAlign: 'center', userSelect: 'none',
      }}
    >
      <div style={{ fontSize: 24, marginBottom: 10, color: loaded ? '#3dba78' : 'var(--text-muted)' }}>
        {loaded ? '✓' : '↑'}
      </div>
      <div style={{ fontWeight: 600, color: loaded ? '#3dba78' : 'var(--text-primary)', fontSize: 15, marginBottom: 4 }}>{label}</div>
      <div style={{ fontSize: 13, color: 'var(--text-muted)' }}>{loaded ? fileName : subtitle}</div>
    </div>
  )
}

function FilterBar({ lines, filter, setFilter }) {
  const counts = {
    all: lines.length,
    missing: lines.filter(l=>l.status==='missing').length,
    short: lines.filter(l=>l.status==='short').length,
    over: lines.filter(l=>l.status==='over').length,
    'unmatched-warehouse': lines.filter(l=>l.status==='unmatched-warehouse').length,
    ok: lines.filter(l=>l.status==='ok').length,
  }
  const labels = { all:'All', missing:'Missing', short:'Short', over:'Over', 'unmatched-warehouse':'Not shipped', ok:'Match' }
  return (
    <div style={{ display:'flex', gap:8, flexWrap:'wrap', marginBottom:18 }}>
      {Object.entries(labels).map(([f, lbl]) => (
        <button key={f} onClick={() => setFilter(f)} style={{
          padding:'5px 14px', borderRadius:6, fontSize:13, fontWeight:600,
          background: filter===f ? 'var(--accent)' : 'transparent',
          color: filter===f ? '#fff' : 'var(--text-secondary)',
          border: `1px solid ${filter===f ? 'var(--accent)' : 'var(--border-mid)'}`,
          transition:'all 0.1s'
        }}>{lbl} <span style={{ opacity:0.7, fontWeight:400 }}>({counts[f]})</span></button>
      ))}
    </div>
  )
}

function ReconciliationTable({ lines }) {
  const [filter, setFilter] = useState('all')
  const filtered = filter==='all' ? lines : lines.filter(l=>l.status===filter)
  const issues = lines.filter(l=>l.status!=='ok').length
  return (
    <div>
      <FilterBar lines={lines} filter={filter} setFilter={setFilter} />
      {issues > 0 && (
        <div style={{ fontSize:13, color:'#f06060', marginBottom:14, fontWeight:500 }}>
          ⚠ {issues} exception{issues!==1?'s':''} found in this warehouse
        </div>
      )}
      {issues === 0 && lines.length > 0 && (
        <div style={{ fontSize:13, color:'#3dba78', marginBottom:14, fontWeight:500 }}>
          ✓ All parts reconciled
        </div>
      )}
      <div style={{ overflowX:'auto', borderRadius:'var(--radius)', border:'1px solid var(--border)' }}>
        <table>
          <thead>
            <tr style={{ background:'#0e1219', borderBottom:'1px solid var(--border)' }}>
              {['Part number','Main Warehouse Count','Off Site Count','Variance','Status'].map(h => (
                <th key={h} style={{ padding:'12px 16px', color:'var(--text-muted)', fontWeight:700, fontSize:12, textTransform:'uppercase', letterSpacing:'0.08em', whiteSpace:'nowrap' }}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {filtered.length===0 && (
              <tr><td colSpan={5} style={{ padding:'32px 16px', textAlign:'center', color:'var(--text-muted)', fontSize:15 }}>No records match this filter.</td></tr>
            )}
            {filtered.map(line => (
              <tr key={line.part} style={{ borderBottom:'1px solid var(--border)', background: line.status!=='ok' ? 'rgba(255,255,255,0.015)' : 'transparent' }}>
                <td style={{ padding:'13px 16px', fontFamily:'var(--mono)', fontSize:13, color:'var(--text-primary)', fontWeight:600 }}>{line.part}</td>
                <td style={{ padding:'13px 16px', fontFamily:'var(--mono)', fontSize:14, textAlign:'right' }}>{line.shippedQty}</td>
                <td style={{ padding:'13px 16px', fontFamily:'var(--mono)', fontSize:14, textAlign:'right' }}>{line.reportedQty}</td>
                <td style={{ padding:'13px 16px', fontFamily:'var(--mono)', fontSize:14, textAlign:'right', fontWeight:600,
                  color: line.variance>0 ? '#e8b84b' : line.variance<0 ? '#f06060' : 'var(--text-muted)' }}>
                  {line.variance===0 ? '—' : (line.variance>0?'+':'')+line.variance}
                </td>
                <td style={{ padding:'13px 16px' }}><Badge status={line.status} /></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  )
}

function WarehousePanel({ name, lines, totalShipped, totalReported, exceptionCategories, loaded }) {
  const pct = calcReconciliationPct(lines, name, exceptionCategories)
  const variance = totalReported - totalShipped
  const issues = lines.filter(l=>l.status!=='ok').length
  const pctColor = pct===100 ? '#3dba78' : pct>=80 ? '#e8b84b' : '#f06060'

  if (!loaded) {
    return (
      <div style={{ background:'var(--bg-card)', border:'1px dashed var(--border-mid)', borderRadius:'var(--radius-lg)', padding:48, marginBottom:20, textAlign:'center' }}>
        <div style={{ fontSize:17, fontWeight:700, color:'var(--text-primary)', marginBottom:8 }}>{name} report not uploaded yet</div>
        <div style={{ fontSize:14, color:'var(--text-secondary)' }}>Go to the Upload tab to add this warehouse's report — the other loaded warehouses aren't affected.</div>
      </div>
    )
  }

  return (
    <div style={{ background:'var(--bg-card)', border:'1px solid var(--border)', borderRadius:'var(--radius-lg)', padding:28, marginBottom:20 }}>
      <div style={{ display:'flex', alignItems:'flex-start', justifyContent:'space-between', marginBottom:24, flexWrap:'wrap', gap:16 }}>
        <div>
          <h2 style={{ fontSize:22, fontWeight:700, color:'var(--text-primary)', marginBottom:4 }}>{name}</h2>
          <div style={{ fontSize:14, color:'var(--text-secondary)' }}>{lines.length} part{lines.length!==1?'s':''} · {issues} exception{issues!==1?'s':''}</div>
        </div>
        <div style={{ display:'flex', gap:28, alignItems:'flex-start' }}>
          <div style={{ textAlign:'right' }}>
            <div style={{ fontSize:12, color:'var(--text-muted)', textTransform:'uppercase', letterSpacing:'0.08em', fontWeight:700, marginBottom:4 }}>Reconciled</div>
            <div style={{ fontFamily:'var(--mono)', fontSize:28, fontWeight:700, color:pctColor }}>{pct}%</div>
          </div>
          <div style={{ textAlign:'right' }}>
            <div style={{ fontSize:12, color:'var(--text-muted)', textTransform:'uppercase', letterSpacing:'0.08em', fontWeight:700, marginBottom:4 }}>Variance</div>
            <div style={{ fontFamily:'var(--mono)', fontSize:28, fontWeight:700, color: variance===0?'#3dba78':'#e8b84b' }}>
              {variance===0 ? '0' : (variance>0?'+':'')+variance}
            </div>
          </div>
        </div>
      </div>
      <ReconciliationTable lines={lines} />
    </div>
  )
}

// ─── main App ─────────────────────────────────────────────────────────────────

const TABS = ['Upload','Summary','EAB','WSI','WS2','Exceptions','BOL Tracker','Truck Load Builder']

export default function App() {
  const [tab, setTab] = useState('Upload')
  const [data, setData] = useState({ main:null, mainName:null, eab:null, eabName:null, wsi:null, wsiName:null, ws2:null, ws2Name:null })
  const [showPrint, setShowPrint] = useState(false)
  const [reportName, setReportName] = useState('')
  const [showReset, setShowReset] = useState(false)
  const [exceptionReasons, setExceptionReasons] = useState({})
  const [exceptionCategories, setExceptionCategories] = useState({})

  const setFile = (key, nameKey, isEAB = false) => (rows, name) => setData(d => ({...d, [key]:rows, [nameKey]:name}))

  const handleFileUpload = (key, nameKey, isEAB = false) => (file) => {
    if (!file) return
    const reader = isEAB ? readExcelEAB : readExcel
    reader(file).then(rows => setData(d => ({...d, [key]:rows, [nameKey]:file.name}))).catch(() => alert(`Could not read ${file.name}. Make sure it's an Excel or CSV file.`))
  }

  const mainRows = useMemo(() => data.main ? parseMainWarehouse(data.main) : [], [data.main])
  const eabRows  = useMemo(() => data.eab  ? parseEAB(data.eab)            : [], [data.eab])
  const wsiRows  = useMemo(() => data.wsi  ? parseWSI_WS2(data.wsi)        : [], [data.wsi])
  const ws2Rows  = useMemo(() => data.ws2  ? parseWSI_WS2(data.ws2)        : [], [data.ws2])

  const mainEAB = useMemo(() => mainRows.filter(r=>r.destination==='EAB'), [mainRows])
  const mainWSI = useMemo(() => mainRows.filter(r=>r.destination==='WSI'), [mainRows])
  const mainWS2 = useMemo(() => mainRows.filter(r=>r.destination==='WS2'), [mainRows])

  const eabLoaded = !!data.eab
  const wsiLoaded = !!data.wsi
  const ws2Loaded = !!data.ws2

  const eabRec = useMemo(() => eabLoaded ? reconcileWarehouse(mainEAB, eabRows) : [], [eabLoaded, mainEAB, eabRows])
  const wsiRec = useMemo(() => wsiLoaded ? reconcileWarehouse(mainWSI, wsiRows) : [], [wsiLoaded, mainWSI, wsiRows])
  const ws2Rec = useMemo(() => ws2Loaded ? reconcileWarehouse(mainWS2, ws2Rows) : [], [ws2Loaded, mainWS2, ws2Rows])

  const tot = (arr, key) => arr.reduce((a,r)=>a+(r[key]??0),0)
  const eabShipped=tot(mainEAB,'qty'), wsiShipped=tot(mainWSI,'qty'), ws2Shipped=tot(mainWS2,'qty')
  const eabReported=tot(eabRows,'qty'), wsiReported=tot(wsiRows,'qty'), ws2Reported=tot(ws2Rows,'qty')
  const totalShipped=eabShipped+wsiShipped+ws2Shipped
  const totalReported=eabReported+wsiReported+ws2Reported
  const totalVariance=totalReported-totalShipped

  const allLines = [...eabRec, ...wsiRec, ...ws2Rec]
  const allIssues = allLines.filter(l=>l.status!=='ok')
  const eabMatched = eabRec.filter(l => isLineAccounted(l, 'EAB', exceptionCategories)).length
  const wsiMatched = wsiRec.filter(l => isLineAccounted(l, 'WSI', exceptionCategories)).length
  const ws2Matched = ws2Rec.filter(l => isLineAccounted(l, 'WS2', exceptionCategories)).length
  const overallPct = allLines.length ? Math.round(((eabMatched + wsiMatched + ws2Matched) / allLines.length) * 100) : 0
  const filesLoaded = [data.main,data.eab,data.wsi,data.ws2].filter(Boolean).length
  const hasData = filesLoaded > 0
  // Main plus at least one offsite warehouse report is enough to reconcile —
  // no need to wait on all three sites before getting useful results.
  const canReconcile = !!data.main && (eabLoaded || wsiLoaded || ws2Loaded)

  const prevCanReconcile = React.useRef(false)
  React.useEffect(() => {
    if (canReconcile && !prevCanReconcile.current) setTab('Summary')
    prevCanReconcile.current = canReconcile
  }, [canReconcile])

  const handleReset = () => {
    setData({ main:null, mainName:null, eab:null, eabName:null, wsi:null, wsiName:null, ws2:null, ws2Name:null })
    setTab('Upload')
    setShowReset(false)
    setReportName('')
  }

  const tabStyle = (t) => ({
    padding:'10px 20px', fontSize:15, fontWeight:600,
    background:'transparent', border:'none',
    borderBottom: tab===t ? '2px solid var(--accent-light)' : '2px solid transparent',
    color: tab===t ? 'var(--text-primary)' : 'var(--text-secondary)',
    cursor: (!hasData && t!=='Upload') ? 'not-allowed' : 'pointer',
    marginBottom:-1, display:'flex', alignItems:'center', gap:7,
    opacity: (!hasData && t!=='Upload') ? 0.35 : 1,
    transition:'color 0.15s',
  })

  return (
    <div style={{ maxWidth:1140, margin:'0 auto', padding:'40px 24px' }}>

      {/* header */}
      <div style={{ marginBottom:36, display:'flex', alignItems:'flex-start', justifyContent:'space-between', flexWrap:'wrap', gap:16 }}>
        <div>
          <div style={{ display:'flex', alignItems:'center', gap:10, marginBottom:8 }}>
            <div style={{ width:8, height:8, borderRadius:'50%', background:'var(--accent-light)' }} />
            <span style={{ fontSize:12, color:'var(--text-muted)', textTransform:'uppercase', letterSpacing:'0.12em', fontWeight:700 }}>Metronet</span>
          </div>
          <h1 style={{ fontSize:34, fontWeight:800, color:'var(--text-primary)', letterSpacing:'-0.03em', lineHeight:1.1 }}>Off Site Pallet Audit</h1>
          <div style={{ fontSize:15, color:'var(--text-secondary)', marginTop:8 }}>
            {filesLoaded===0 && 'Upload your shipment data and warehouse reports to begin.'}
            {filesLoaded>0 && !canReconcile && `${filesLoaded} of 4 files loaded — upload Main plus at least one warehouse report to reconcile.`}
            {canReconcile && filesLoaded<4 && `Reconciling with ${filesLoaded} of 4 files loaded — upload the rest whenever they're ready.`}
            {filesLoaded===4 && 'All files loaded. Review the reconciliation below.'}
          </div>
        </div>
        <div style={{ display:'flex', flexDirection:'column', alignItems:'flex-end', gap:12 }}>
          {hasData && !canReconcile && (
            <button onClick={() => setShowReset(true)} style={{
              padding:'9px 18px', borderRadius:7, border:'1px solid var(--border-mid)',
              background:'var(--bg-input)', color:'var(--text-secondary)', fontWeight:700, fontSize:14, display:'flex', alignItems:'center', gap:8
            }}>↺ Refresh</button>
          )}
          {canReconcile && (
            <div style={{ textAlign:'right' }}>
              <div style={{ fontSize:12, color:'var(--text-muted)', textTransform:'uppercase', letterSpacing:'0.08em', fontWeight:700, marginBottom:6 }}>Overall</div>
              <div style={{ fontFamily:'var(--mono)', fontSize:38, fontWeight:800, color: overallPct===100?'#3dba78':overallPct>=80?'#e8b84b':'#f06060', lineHeight:1 }}>{overallPct}%</div>
              <div style={{ fontSize:13, color:'var(--text-secondary)', marginTop:4 }}>
                {allIssues.length===0 ? 'Fully reconciled' : `${allIssues.length} exception${allIssues.length!==1?'s':''} require attention`}
              </div>
            </div>
          )}
          {canReconcile && (
            <div style={{ display:'flex', gap:10, alignItems:'center' }}>
              <button onClick={() => setShowPrint(true)} style={{
                padding:'9px 18px', borderRadius:7, border:'1px solid var(--accent)',
                background:'var(--accent)', color:'#fff', fontWeight:700, fontSize:14, display:'flex', alignItems:'center', gap:8
              }}>🖨 Print / Save PDF</button>
              <button onClick={() => setShowReset(true)} style={{
                padding:'9px 18px', borderRadius:7, border:'1px solid #5a1515',
                background:'#1a0808', color:'#f06060', fontWeight:700, fontSize:14, display:'flex', alignItems:'center', gap:8
              }}>↺ New Month</button>
            </div>
          )}
        </div>
      </div>

      {/* reset confirmation modal */}
      {showReset && (
        <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.8)', zIndex:999, display:'flex', alignItems:'center', justifyContent:'center' }}>
          <div style={{ background:'var(--bg-card)', border:'1px solid #5a1515', borderRadius:12, padding:32, maxWidth:420, width:'90%' }}>
            <div style={{ fontSize:18, fontWeight:800, marginBottom:10, color:'var(--text-primary)' }}>Start a new month?</div>
            <div style={{ fontSize:15, color:'var(--text-secondary)', marginBottom:24, lineHeight:1.6 }}>
              This will clear all uploaded files and reconciliation data. Make sure you've saved your PDF report first.
            </div>
            <div style={{ display:'flex', gap:10 }}>
              <button onClick={() => setShowReset(false)} style={{ flex:1, padding:'10px', borderRadius:7, border:'1px solid var(--border-mid)', background:'transparent', color:'var(--text-secondary)', fontWeight:600, fontSize:14 }}>Cancel</button>
              <button onClick={handleReset} style={{ flex:1, padding:'10px', borderRadius:7, border:'none', background:'#dc2626', color:'#fff', fontWeight:700, fontSize:14 }}>Clear & start fresh</button>
            </div>
          </div>
        </div>
      )}

      {/* print modal */}
      {showPrint && (
        <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.8)', zIndex:998, display:'flex', alignItems:'center', justifyContent:'center' }}>
          <div style={{ background:'var(--bg-card)', border:'1px solid var(--border-mid)', borderRadius:12, padding:32, maxWidth:420, width:'90%' }}>
            <div style={{ fontSize:18, fontWeight:800, marginBottom:10, color:'var(--text-primary)' }}>Name this report</div>
            <div style={{ fontSize:14, color:'var(--text-secondary)', marginBottom:16 }}>This will appear on the PDF as the report title.</div>
            <input
              type="text"
              placeholder="e.g. May 2025 Reconciliation"
              value={reportName}
              onChange={e => setReportName(e.target.value)}
              onKeyDown={e => e.key==='Enter' && reportName.trim() && setShowPrint(false) && setTimeout(()=>setShowPrint('preview'),50)}
              style={{
                width:'100%', padding:'10px 14px', borderRadius:7,
                border:'1px solid var(--border-mid)', background:'var(--bg-input)',
                color:'var(--text-primary)', fontSize:15, marginBottom:16, outline:'none'
              }}
              autoFocus
            />
            <div style={{ display:'flex', gap:10 }}>
              <button onClick={() => setShowPrint(false)} style={{ flex:1, padding:'10px', borderRadius:7, border:'1px solid var(--border-mid)', background:'transparent', color:'var(--text-secondary)', fontWeight:600, fontSize:14 }}>Cancel</button>
              <button onClick={() => { if(reportName.trim()) setShowPrint('preview') }} style={{ flex:1, padding:'10px', borderRadius:7, border:'none', background: reportName.trim()?'var(--accent)':'var(--border-mid)', color:'#fff', fontWeight:700, fontSize:14, cursor: reportName.trim()?'pointer':'not-allowed' }}>Preview report →</button>
            </div>
          </div>
        </div>
      )}

      {showPrint==='preview' && (
        <PrintReport
          reportName={reportName}
          date={new Date().toLocaleDateString('en-US',{month:'long',day:'numeric',year:'numeric'})}
          eabRec={eabRec} wsiRec={wsiRec} ws2Rec={ws2Rec}
          eabShipped={eabShipped} wsiShipped={wsiShipped} ws2Shipped={ws2Shipped}
          eabReported={eabReported} wsiReported={wsiReported} ws2Reported={ws2Reported}
          overallPct={overallPct}
          allIssues={allIssues}
          exceptionReasons={exceptionReasons}
          exceptionCategories={exceptionCategories}
          eabLoaded={eabLoaded} wsiLoaded={wsiLoaded} ws2Loaded={ws2Loaded}
          onClose={() => setShowPrint(false)}
        />
      )}

      {/* divider */}
      <div style={{ height:1, background:'var(--border)', marginBottom:0 }} />

      {/* tabs */}
      <div style={{ display:'flex', gap:0, borderBottom:'1px solid var(--border)', marginBottom:32, flexWrap:'wrap' }}>
        {TABS.map(t => {
          const alwaysEnabled = t === 'Upload' || t === 'BOL Tracker' || t === 'Truck Load Builder'
          const disabled = !hasData && !alwaysEnabled
          const isActive = tab === t
          const badgeCount = t === 'Exceptions' ? allIssues.length : null
          return (
            <button key={t} onClick={() => !disabled && setTab(t)} disabled={disabled} style={{
              padding:'10px 18px', fontSize:14, fontWeight:600,
              background:'transparent', border:'none',
              borderBottom: isActive ? '2px solid var(--accent-light)' : '2px solid transparent',
              color: disabled ? 'var(--text-muted)' : isActive ? 'var(--text-primary)' : 'var(--text-secondary)',
              cursor: disabled ? 'not-allowed' : 'pointer',
              marginBottom:-1, display:'flex', alignItems:'center', gap:7,
              opacity: disabled ? 0.35 : 1,
              transition:'color 0.15s', whiteSpace:'nowrap'
            }}>
              {t}
              {badgeCount > 0 && (
                <span style={{ fontSize:11, background:'#f06060', color:'#fff', borderRadius:10, padding:'1px 7px', fontWeight:700 }}>{badgeCount}</span>
              )}
            </button>
          )
        })}
      </div>

      {/* ── Upload ── */}
      {tab==='Upload' && (
        <div>
          <div style={{ marginBottom:24 }}>
            <h2 style={{ fontSize:20, fontWeight:700, marginBottom:6 }}>Upload files</h2>
            <p style={{ fontSize:15, color:'var(--text-secondary)' }}>Accepts .xlsx, .xls, or .csv. Column headers are matched automatically.</p>
          </div>
          <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit, minmax(230px, 1fr))', gap:14 }}>
            <UploadZone label="Main warehouse shipments" subtitle="Your internal shipment records" onFile={handleFileUpload('main','mainName')} loaded={!!data.main} fileName={data.mainName} />
            <UploadZone label="WSI warehouse report"     subtitle="WSI inventory report"          onFile={handleFileUpload('wsi','wsiName')} loaded={!!data.wsi}  fileName={data.wsiName} />
            <UploadZone label="WS2 warehouse report"     subtitle="WS2 inventory report"          onFile={handleFileUpload('ws2','ws2Name')} loaded={!!data.ws2}  fileName={data.ws2Name} />
            <UploadZone label="EAB warehouse report"     subtitle="Multi-sheet EAB workbook"       onFile={handleFileUpload('eab','eabName',true)} loaded={!!data.eab}  fileName={data.eabName} />
          </div>
          {filesLoaded>0 && (
            <div style={{ marginTop:28, background:'var(--bg-card)', border:'1px solid var(--border)', borderRadius:'var(--radius-lg)', padding:22 }}>
              <div style={{ fontSize:13, color:'var(--text-muted)', marginBottom:14, textTransform:'uppercase', letterSpacing:'0.08em', fontWeight:700 }}>Loaded files</div>
              {[
                {key:'main',name:data.mainName,rows:mainRows,label:'Main warehouse'},
                {key:'wsi', name:data.wsiName, rows:wsiRows, label:'WSI'},
                {key:'ws2', name:data.ws2Name, rows:ws2Rows, label:'WS2'},
                {key:'eab', name:data.eabName, rows:eabRows, label:'EAB'},
              ].filter(f=>f.name).map(f=>(
                <div key={f.key} style={{ display:'flex', justifyContent:'space-between', alignItems:'center', padding:'10px 0', borderBottom:'1px solid var(--border)', fontSize:15 }}>
                  <span style={{ color:'var(--text-secondary)', fontWeight:600, minWidth:160 }}>{f.label}</span>
                  <span style={{ color:'var(--text-primary)', fontFamily:'var(--mono)', fontSize:13 }}>{f.name}</span>
                  <span style={{ color:'#3dba78', fontSize:14, fontWeight:600 }}>{f.rows.length} rows</span>
                </div>
              ))}
              {canReconcile && (
                <button onClick={()=>setTab('Summary')} style={{
                  marginTop:18, width:'100%', padding:'13px', borderRadius:'var(--radius)',
                  background:'var(--accent)', color:'#fff', border:'none', fontWeight:700, fontSize:15,
                  letterSpacing:'0.01em'
                }}>View reconciliation →</button>
              )}
            </div>
          )}
        </div>
      )}

      {/* ── Summary ── */}
      {tab==='Summary' && (
        <div>
          <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit, minmax(160px,1fr))', gap:14, marginBottom:32 }}>
            <StatCard label="Main WHS Total"  value={totalShipped.toLocaleString()}  sub="Pallets sent to all warehouses"   accent="var(--accent)" />
            <StatCard label="Off Site Qty" value={totalReported.toLocaleString()} sub="Pallets reported by warehouses"    accent="var(--accent)" />
            <StatCard label="Net variance"   value={(totalVariance>0?'+':'')+totalVariance}
              sub={totalVariance===0?'Quantities match':'Pallet count difference'}
              accent={totalVariance===0?'#2d7a52':'#c47a15'}
              warn={totalVariance!==0} />
            <StatCard label="Total Accuracy" value={`${overallPct}%`}
              sub={allIssues.length===0?'All parts match':`${allIssues.length} exception${allIssues.length!==1?'s':''}`}
              accent={overallPct===100?'#2d7a52':overallPct>=80?'#c47a15':'#a32020'}
              warn={overallPct<80} />
          </div>

          <div style={{ display:'grid', gridTemplateColumns:'repeat(3,1fr)', gap:14, marginBottom:32 }}>
            {[
              {name:'EAB',shipped:eabShipped,reported:eabReported,rec:eabRec,loaded:eabLoaded},
              {name:'WSI',shipped:wsiShipped,reported:wsiReported,rec:wsiRec,loaded:wsiLoaded},
              {name:'WS2',shipped:ws2Shipped,reported:ws2Reported,rec:ws2Rec,loaded:ws2Loaded},
            ].map(w => {
              if (!w.loaded) {
                return (
                  <div key={w.name} onClick={()=>setTab('Upload')} style={{
                    background:'var(--bg-card)', border:'1px dashed var(--border-mid)',
                    borderRadius:'var(--radius-lg)', padding:'20px 24px', cursor:'pointer', textAlign:'center'
                  }}>
                    <div style={{ fontWeight:800, marginBottom:10, fontSize:18, letterSpacing:'-0.01em' }}>{w.name}</div>
                    <div style={{ fontSize:13, color:'var(--text-muted)' }}>Not uploaded yet</div>
                  </div>
                )
              }
              const v = w.reported - w.shipped
              const p = calcReconciliationPct(w.rec, w.name, exceptionCategories)
              const issues = w.rec.filter(l=>l.status!=='ok').length
              const pColor = p===100?'#3dba78':p>=80?'#e8b84b':'#f06060'
              return (
                <div key={w.name} onClick={()=>setTab(w.name)} style={{
                  background:'var(--bg-card)', border:`1px solid ${issues>0?'var(--border-mid)':'var(--border)'}`,
                  borderRadius:'var(--radius-lg)', padding:'20px 24px', cursor:'pointer', transition:'border-color 0.15s, background 0.15s'
                }}
                  onMouseEnter={e=>{e.currentTarget.style.borderColor='var(--border-focus)';e.currentTarget.style.background='var(--bg-card-hover)'}}
                  onMouseLeave={e=>{e.currentTarget.style.borderColor=issues>0?'var(--border-mid)':'var(--border)';e.currentTarget.style.background='var(--bg-card)'}}
                >
                  <div style={{ fontWeight:800, marginBottom:16, fontSize:18, letterSpacing:'-0.01em' }}>{w.name}</div>
                  <div style={{ display:'flex', justifyContent:'space-between', fontSize:14, color:'var(--text-muted)', marginBottom:6 }}>
                    <span>Main WHS Count</span><span style={{ fontFamily:'var(--mono)', color:'var(--text-secondary)', fontWeight:600 }}>{w.shipped}</span>
                  </div>
                  <div style={{ display:'flex', justifyContent:'space-between', fontSize:14, color:'var(--text-muted)', marginBottom:16, paddingBottom:16, borderBottom:'1px solid var(--border)' }}>
                    <span>Off Site Count</span><span style={{ fontFamily:'var(--mono)', color:'var(--text-secondary)', fontWeight:600 }}>{w.reported}</span>
                  </div>
                  <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center' }}>
                    <span style={{ fontSize:13, color:issues>0?'#f06060':'#3dba78', fontWeight:600 }}>
                      {issues>0 ? `${issues} exception${issues!==1?'s':''}` : '✓ All clear'}
                    </span>
                    <span style={{ fontFamily:'var(--mono)', fontWeight:800, fontSize:20, color:pColor }}>{p}%</span>
                  </div>
                </div>
              )
            })}
          </div>

          {allIssues.length>0 && (
            <div style={{ background:'var(--bg-card)', border:'1px solid #5a1515', borderRadius:'var(--radius-lg)', padding:24 }}>
              <div style={{ fontSize:16, fontWeight:700, marginBottom:18, color:'#f06060' }}>
                ⚠ {allIssues.length} exception{allIssues.length!==1?'s':''} require attention
              </div>
              <div style={{ overflowX:'auto', borderRadius:'var(--radius)', border:'1px solid var(--border)' }}>
                <table style={{ fontSize:14 }}>
                  <thead>
                    <tr style={{ background:'#0e1219', borderBottom:'1px solid var(--border)' }}>
                      {['Warehouse','Part number','Main Warehouse Count','Off Site Count','Variance','Status'].map(h=>(
                        <th key={h} style={{ padding:'11px 16px', color:'var(--text-muted)', fontWeight:700, fontSize:12, textTransform:'uppercase', letterSpacing:'0.08em' }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {allIssues.map((issue,i) => {
                      const wh = eabRec.includes(issue)?'EAB':wsiRec.includes(issue)?'WSI':'WS2'
                      return (
                        <tr key={i} style={{ borderBottom:'1px solid var(--border)' }}>
                          <td style={{ padding:'12px 16px' }}>
                            <span style={{ fontWeight:700, fontSize:13, background:'var(--bg-input)', border:'1px solid var(--border-mid)', borderRadius:5, padding:'3px 10px' }}>{wh}</span>
                          </td>
                          <td style={{ padding:'12px 16px', fontFamily:'var(--mono)', fontSize:13, fontWeight:700 }}>{issue.part}</td>
                          <td style={{ padding:'12px 16px', fontFamily:'var(--mono)', fontSize:14, textAlign:'right' }}>{issue.shippedQty}</td>
                          <td style={{ padding:'12px 16px', fontFamily:'var(--mono)', fontSize:14, textAlign:'right' }}>{issue.reportedQty}</td>
                          <td style={{ padding:'12px 16px', fontFamily:'var(--mono)', fontSize:14, textAlign:'right', fontWeight:700,
                            color:issue.variance>0?'#e8b84b':'#f06060' }}>
                            {(issue.variance>0?'+':'')+issue.variance}
                          </td>
                          <td style={{ padding:'12px 16px' }}><Badge status={issue.status} /></td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}
          {allIssues.length===0 && canReconcile && (
            <div style={{ textAlign:'center', padding:48, color:'#3dba78', fontSize:17, fontWeight:600 }}>
              ✓ All pallets reconciled across {[eabLoaded&&'EAB',wsiLoaded&&'WSI',ws2Loaded&&'WS2'].filter(Boolean).join(', ')}.
            </div>
          )}
        </div>
      )}

      {tab==='EAB' && <WarehousePanel name="EAB" lines={eabRec} totalShipped={eabShipped} totalReported={eabReported} exceptionCategories={exceptionCategories} loaded={eabLoaded} />}
      {tab==='WSI' && <WarehousePanel name="WSI" lines={wsiRec} totalShipped={wsiShipped} totalReported={wsiReported} exceptionCategories={exceptionCategories} loaded={wsiLoaded} />}
      {tab==='WS2' && <WarehousePanel name="WS2" lines={ws2Rec} totalShipped={ws2Shipped} totalReported={ws2Reported} exceptionCategories={exceptionCategories} loaded={ws2Loaded} />}

      {/* ── Exceptions ── */}
      {tab==='Exceptions' && (
        <div>
          <div style={{ marginBottom:24 }}>
            <h2 style={{ fontSize:20, fontWeight:700, marginBottom:6 }}>All exceptions</h2>
            <p style={{ fontSize:15, color:'var(--text-secondary)' }}>Every discrepancy across all three warehouses in one view.</p>
          </div>
          {allIssues.length===0
            ? <div style={{ textAlign:'center', padding:72, color:'#3dba78', fontSize:17, fontWeight:600 }}>✓ No exceptions found.</div>
            : (
              <div style={{ background:'var(--bg-card)', border:'1px solid var(--border)', borderRadius:'var(--radius-lg)', padding:24 }}>
                <div style={{ overflowX:'auto', borderRadius:'var(--radius)', border:'1px solid var(--border)' }}>
                  <table style={{ fontSize:14 }}>
                    <thead>
                      <tr style={{ background:'#0e1219', borderBottom:'1px solid var(--border)' }}>
                        {['Warehouse','Part number','Main Warehouse Count','Off Site Count','Variance','Status','Reason'].map(h=>(
                          <th key={h} style={{ padding:'12px 16px', color:'var(--text-muted)', fontWeight:700, fontSize:12, textTransform:'uppercase', letterSpacing:'0.08em' }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {allIssues.map((issue,i) => {
                        const wh = eabRec.includes(issue)?'EAB':wsiRec.includes(issue)?'WSI':'WS2'
                        const reasonKey = `${wh}-${issue.part}`
                        return (
                          <tr key={i} style={{ borderBottom:'1px solid var(--border)' }}>
                            <td style={{ padding:'13px 16px' }}>
                              <span style={{ fontWeight:700, fontSize:13, background:'var(--bg-input)', border:'1px solid var(--border-mid)', borderRadius:5, padding:'3px 10px' }}>{wh}</span>
                            </td>
                            <td style={{ padding:'13px 16px', fontFamily:'var(--mono)', fontSize:13, fontWeight:700 }}>{issue.part}</td>
                            <td style={{ padding:'13px 16px', fontFamily:'var(--mono)', fontSize:14, textAlign:'right' }}>{issue.shippedQty}</td>
                            <td style={{ padding:'13px 16px', fontFamily:'var(--mono)', fontSize:14, textAlign:'right' }}>{issue.reportedQty}</td>
                            <td style={{ padding:'13px 16px', fontFamily:'var(--mono)', fontSize:14, textAlign:'right', fontWeight:700,
                              color:issue.variance>0?'#e8b84b':'#f06060' }}>
                              {(issue.variance>0?'+':'')+issue.variance}
                            </td>
                            <td style={{ padding:'13px 16px' }}><Badge status={issue.status} /></td>
                            <td style={{ padding:'8px 16px' }}>
                              <select
                                value={exceptionCategories[reasonKey] || ''}
                                onChange={e => setExceptionCategories(prev => ({ ...prev, [reasonKey]: e.target.value }))}
                                style={{
                                  width:'100%', minWidth:200, padding:'6px 8px', borderRadius:6,
                                  border:'1px solid var(--border-mid)', background:'var(--bg-input)',
                                  color:'var(--text-primary)', fontSize:13, outline:'none', marginBottom:6
                                }}
                              >
                                <option value="">Select reason...</option>
                                {EXCEPTION_REASON_OPTIONS.map(opt => (
                                  <option key={opt} value={opt}>{opt}</option>
                                ))}
                              </select>
                              <input
                                type="text"
                                placeholder="Additional notes (optional)..."
                                value={exceptionReasons[reasonKey] || ''}
                                onChange={e => setExceptionReasons(prev => ({ ...prev, [reasonKey]: e.target.value }))}
                                style={{
                                  width:'100%', minWidth:160, padding:'6px 10px', borderRadius:6,
                                  border:'1px solid var(--border-mid)', background:'var(--bg-input)',
                                  color:'var(--text-primary)', fontSize:13, outline:'none'
                                }}
                              />
                            </td>
                          </tr>
                        )
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )
          }
        </div>
      )}

      {tab==='BOL Tracker' && (
        <MonthlyBolTracker />
      )}

      {tab==='Truck Load Builder' && (
        <TruckLoadBuilder />
      )}

      {/* footer */}
      <div style={{ marginTop:56, paddingTop:24, borderTop:'1px solid var(--border)', fontSize:13, color:'var(--text-muted)', display:'flex', justifyContent:'space-between', alignItems:'center', flexWrap:'wrap', gap:16 }}>
        <span>All processing happens in your browser — no data is sent to any server.</span>
        <img src="/metronet-logo.png" alt="Metronet" style={{ height:28, width:'auto', opacity:0.9 }} />
      </div>
    </div>
  )
}
