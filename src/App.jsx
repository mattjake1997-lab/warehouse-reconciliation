import React, { useState, useCallback, useMemo } from 'react'
import * as XLSX from 'xlsx'
import PrintReport from './PrintReport.jsx'
import BillOfLanding from './BillOfLanding.jsx'

// ─── helpers ────────────────────────────────────────────────────────────────

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

// Special reader for EAB: reads ALL sheets, skips row 1 (title), uses row 2 as headers
// Injects __sheet__ on every row and skips already-shipped pallets
function readExcelEAB(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = e => {
      try {
        const wb = XLSX.read(e.target.result, { type: 'array', cellDates: false })
        const SKIP = ['Information', 'Index of Part No.s']
        const allRows = []
        for (const sheetName of wb.SheetNames) {
          if (SKIP.some(s => sheetName.toLowerCase().includes(s.toLowerCase()))) continue
          const ws = wb.Sheets[sheetName]
          // sheet_to_json uses first row as headers by default
          // We need to skip row 1 (title) and use row 2 as headers
          const raw = XLSX.utils.sheet_to_json(ws, { defval: '', header: 1 })
          if (raw.length < 3) continue
          // row index 1 = headers (0-indexed), row index 2+ = data
          const headers = raw[1]
          for (let i = 2; i < raw.length; i++) {
            const rowArr = raw[i]
            if (!rowArr.some(v => v !== '')) continue
            const obj = { __sheet__: sheetName }
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

// EAB: multi-sheet workbook. Each sheet = one part number.
// Row 1 = title, Row 2 = headers, Row 3+ = data.
// Part number is in the "Name" column (col index 3).
// Only count rows where Date Shipped (col 7) is empty = still in storage.
// The XLSX library flattens all sheets into one array with a __sheet__ marker,
// so we use a special approach: we pass all rows and detect sheet breaks via
// a special _sheet field we inject during parsing in readExcelAllSheets().
function parseEAB(rows) {
  if (!rows.length) return []
  // rows here come from readExcelAllSheets which includes _sheetName on each row
  // Structure per data row: Total Spaces(0), Spaces by Receipt(1), Make(2), Name(3),
  //   Space Description(4), Pallets per Space(5), Date Received(6), Date Shipped(7),
  //   Spaces by Return Shipment(8), Comments(9)
  const SKIP_SHEETS = ['Information', 'Index of Part No.s']
  const result = []

  for (const r of rows) {
    const sheetName = r.__sheet__ || ''
    if (SKIP_SHEETS.some(s => sheetName.includes(s))) continue
    if (r.__isHeader__) continue // skip header rows

    const partRaw = r['Name'] ?? r['name'] ?? ''
    const palletsRaw = r['Pallets per Space'] ?? r['Pallets Per Space'] ?? r['pallets per space'] ?? 1
    const dateShipped = r['Date Shipped'] ?? r['date shipped'] ?? null

    // Only count pallets still in storage (Date Shipped is empty)
    if (dateShipped) continue

    const partStr = String(partRaw).trim()
    if (!partStr || partStr === '' || partStr === 'n/a' || partStr === 'Name') continue

    // Normalize part number to string of digits
    let partNumber = partStr.replace(/[^0-9]/g, '')
    if (!partNumber) continue

    result.push({
      partNumber: partNumber.toUpperCase(),
      qty: parseNumber(palletsRaw),
      sheetName,
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
// Only count parts that fully match; anything with any issue brings the score down
function calcReconciliationPct(lines) {
  if (!lines.length) return 0
  const matched = lines.filter(l => l.status === 'ok').length
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
              {['Part number','Shipped qty','Reported qty','Variance','Status'].map(h => (
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

function WarehousePanel({ name, lines, totalShipped, totalReported }) {
  const pct = calcReconciliationPct(lines)
  const variance = totalReported - totalShipped
  const issues = lines.filter(l=>l.status!=='ok').length
  const pctColor = pct===100 ? '#3dba78' : pct>=80 ? '#e8b84b' : '#f06060'
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

const TABS = ['Upload','Summary','EAB','WSI','WS2','Exceptions','Bill of Landing']

export default function App() {
  const [tab, setTab] = useState('Upload')
  const [data, setData] = useState({ main:null, mainName:null, eab:null, eabName:null, wsi:null, wsiName:null, ws2:null, ws2Name:null })
  const [showPrint, setShowPrint] = useState(false)
  const [reportName, setReportName] = useState('')
  const [showReset, setShowReset] = useState(false)
  const [bols, setBols] = useState([])

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

  const eabRec = useMemo(() => reconcileWarehouse(mainEAB, eabRows), [mainEAB, eabRows])
  const wsiRec = useMemo(() => reconcileWarehouse(mainWSI, wsiRows), [mainWSI, wsiRows])
  const ws2Rec = useMemo(() => reconcileWarehouse(mainWS2, ws2Rows), [mainWS2, ws2Rows])

  const tot = (arr, key) => arr.reduce((a,r)=>a+(r[key]??0),0)
  const eabShipped=tot(mainEAB,'qty'), wsiShipped=tot(mainWSI,'qty'), ws2Shipped=tot(mainWS2,'qty')
  const eabReported=tot(eabRows,'qty'), wsiReported=tot(wsiRows,'qty'), ws2Reported=tot(ws2Rows,'qty')
  const totalShipped=eabShipped+wsiShipped+ws2Shipped
  const totalReported=eabReported+wsiReported+ws2Reported
  const totalVariance=totalReported-totalShipped

  const allLines = [...eabRec, ...wsiRec, ...ws2Rec]
  const allIssues = allLines.filter(l=>l.status!=='ok')
  const overallPct = calcReconciliationPct(allLines)
  const filesLoaded = [data.main,data.eab,data.wsi,data.ws2].filter(Boolean).length
  const hasData = filesLoaded > 0

  const prevLoaded = React.useRef(0)
  React.useEffect(() => {
    if (filesLoaded===4 && prevLoaded.current<4) setTab('Summary')
    prevLoaded.current = filesLoaded
  }, [filesLoaded])

  const handleReset = () => {
    setData({ main:null, mainName:null, eab:null, eabName:null, wsi:null, wsiName:null, ws2:null, ws2Name:null })
    setTab('Upload')
    setShowReset(false)
    setReportName('')
    // Note: BOLs are kept across resets — they are a running log
  }

  const handleAddBOL = (bol) => setBols(prev => [...prev, bol])
  const handleRemoveBOL = (id) => setBols(prev => prev.filter(b => b.id !== id))

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
          <h1 style={{ fontSize:34, fontWeight:800, color:'var(--text-primary)', letterSpacing:'-0.03em', lineHeight:1.1 }}>Inventory Reconciliation</h1>
          <div style={{ fontSize:15, color:'var(--text-secondary)', marginTop:8 }}>
            {filesLoaded===0 && 'Upload your shipment data and warehouse reports to begin.'}
            {filesLoaded>0 && filesLoaded<4 && `${filesLoaded} of 4 files loaded — upload remaining files to reconcile.`}
            {filesLoaded===4 && 'All files loaded. Review the reconciliation below.'}
          </div>
        </div>
        <div style={{ display:'flex', flexDirection:'column', alignItems:'flex-end', gap:12 }}>
          {filesLoaded===4 && (
            <div style={{ textAlign:'right' }}>
              <div style={{ fontSize:12, color:'var(--text-muted)', textTransform:'uppercase', letterSpacing:'0.08em', fontWeight:700, marginBottom:6 }}>Overall</div>
              <div style={{ fontFamily:'var(--mono)', fontSize:38, fontWeight:800, color: overallPct===100?'#3dba78':overallPct>=80?'#e8b84b':'#f06060', lineHeight:1 }}>{overallPct}%</div>
              <div style={{ fontSize:13, color:'var(--text-secondary)', marginTop:4 }}>
                {allIssues.length===0 ? 'Fully reconciled' : `${allIssues.length} exception${allIssues.length!==1?'s':''} require attention`}
              </div>
            </div>
          )}
          {filesLoaded===4 && (
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
          onClose={() => setShowPrint(false)}
        />
      )}

      {/* divider */}
      <div style={{ height:1, background:'var(--border)', marginBottom:0 }} />

      {/* tabs */}
      <div style={{ display:'flex', gap:0, borderBottom:'1px solid var(--border)', marginBottom:32, flexWrap:'wrap' }}>
        {TABS.map(t => {
          const alwaysEnabled = t === 'Upload' || t === 'Bill of Landing'
          const disabled = !hasData && !alwaysEnabled
          const isActive = tab === t
          const badgeCount = t === 'Exceptions' ? allIssues.length : t === 'Bill of Lading' ? bols.length : null
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
                <span style={{ fontSize:11, background: t==='Bill of Lading'?'var(--accent)':'#f06060', color:'#fff', borderRadius:10, padding:'1px 7px', fontWeight:700 }}>{badgeCount}</span>
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
            <UploadZone label="EAB warehouse report"     subtitle="Multi-sheet EAB workbook"       onFile={handleFileUpload('eab','eabName',true)} loaded={!!data.eab}  fileName={data.eabName} />
            <UploadZone label="WSI warehouse report"     subtitle="WSI inventory report"          onFile={handleFileUpload('wsi','wsiName')} loaded={!!data.wsi}  fileName={data.wsiName} />
            <UploadZone label="WS2 warehouse report"     subtitle="WS2 inventory report"          onFile={handleFileUpload('ws2','ws2Name')} loaded={!!data.ws2}  fileName={data.ws2Name} />
          </div>
          {filesLoaded>0 && (
            <div style={{ marginTop:28, background:'var(--bg-card)', border:'1px solid var(--border)', borderRadius:'var(--radius-lg)', padding:22 }}>
              <div style={{ fontSize:13, color:'var(--text-muted)', marginBottom:14, textTransform:'uppercase', letterSpacing:'0.08em', fontWeight:700 }}>Loaded files</div>
              {[
                {key:'main',name:data.mainName,rows:mainRows,label:'Main warehouse'},
                {key:'eab', name:data.eabName, rows:eabRows, label:'EAB'},
                {key:'wsi', name:data.wsiName, rows:wsiRows, label:'WSI'},
                {key:'ws2', name:data.ws2Name, rows:ws2Rows, label:'WS2'},
              ].filter(f=>f.name).map(f=>(
                <div key={f.key} style={{ display:'flex', justifyContent:'space-between', alignItems:'center', padding:'10px 0', borderBottom:'1px solid var(--border)', fontSize:15 }}>
                  <span style={{ color:'var(--text-secondary)', fontWeight:600, minWidth:160 }}>{f.label}</span>
                  <span style={{ color:'var(--text-primary)', fontFamily:'var(--mono)', fontSize:13 }}>{f.name}</span>
                  <span style={{ color:'#3dba78', fontSize:14, fontWeight:600 }}>{f.rows.length} rows</span>
                </div>
              ))}
              {filesLoaded===4 && (
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
            <StatCard label="Total shipped"  value={totalShipped.toLocaleString()}  sub="Pallets sent to all warehouses"   accent="var(--accent)" />
            <StatCard label="Total reported" value={totalReported.toLocaleString()} sub="Pallets reported by warehouses"    accent="var(--accent)" />
            <StatCard label="Net variance"   value={(totalVariance>0?'+':'')+totalVariance}
              sub={totalVariance===0?'Quantities match':'Pallet count difference'}
              accent={totalVariance===0?'#2d7a52':'#c47a15'}
              warn={totalVariance!==0} />
            <StatCard label="Parts reconciled" value={`${overallPct}%`}
              sub={allIssues.length===0?'All parts match':`${allIssues.length} exception${allIssues.length!==1?'s':''}`}
              accent={overallPct===100?'#2d7a52':overallPct>=80?'#c47a15':'#a32020'}
              warn={overallPct<80} />
          </div>

          <div style={{ display:'grid', gridTemplateColumns:'repeat(3,1fr)', gap:14, marginBottom:32 }}>
            {[
              {name:'EAB',shipped:eabShipped,reported:eabReported,rec:eabRec},
              {name:'WSI',shipped:wsiShipped,reported:wsiReported,rec:wsiRec},
              {name:'WS2',shipped:ws2Shipped,reported:ws2Reported,rec:ws2Rec},
            ].map(w => {
              const v = w.reported - w.shipped
              const p = calcReconciliationPct(w.rec)
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
                    <span>Shipped</span><span style={{ fontFamily:'var(--mono)', color:'var(--text-secondary)', fontWeight:600 }}>{w.shipped}</span>
                  </div>
                  <div style={{ display:'flex', justifyContent:'space-between', fontSize:14, color:'var(--text-muted)', marginBottom:16, paddingBottom:16, borderBottom:'1px solid var(--border)' }}>
                    <span>Reported</span><span style={{ fontFamily:'var(--mono)', color:'var(--text-secondary)', fontWeight:600 }}>{w.reported}</span>
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
                      {['Warehouse','Part number','Shipped','Reported','Variance','Status'].map(h=>(
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
          {allIssues.length===0 && filesLoaded===4 && (
            <div style={{ textAlign:'center', padding:48, color:'#3dba78', fontSize:17, fontWeight:600 }}>✓ All pallets reconciled across all three warehouses.</div>
          )}
        </div>
      )}

      {tab==='EAB' && <WarehousePanel name="EAB" lines={eabRec} totalShipped={eabShipped} totalReported={eabReported} />}
      {tab==='WSI' && <WarehousePanel name="WSI" lines={wsiRec} totalShipped={wsiShipped} totalReported={wsiReported} />}
      {tab==='WS2' && <WarehousePanel name="WS2" lines={ws2Rec} totalShipped={ws2Shipped} totalReported={ws2Reported} />}

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
                        {['Warehouse','Part number','Shipped qty','Reported qty','Variance','Status'].map(h=>(
                          <th key={h} style={{ padding:'12px 16px', color:'var(--text-muted)', fontWeight:700, fontSize:12, textTransform:'uppercase', letterSpacing:'0.08em' }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {allIssues.map((issue,i) => {
                        const wh = eabRec.includes(issue)?'EAB':wsiRec.includes(issue)?'WSI':'WS2'
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

      {tab==='Bill of Landing' && (
        <BillOfLanding bols={bols} onAddBOL={handleAddBOL} onRemoveBOL={handleRemoveBOL} />
      )}

      {/* footer */}
      <div style={{ marginTop:56, paddingTop:20, borderTop:'1px solid var(--border)', fontSize:13, color:'var(--text-muted)', display:'flex', justifyContent:'space-between', flexWrap:'wrap', gap:8 }}>
        <span style={{ fontWeight:600 }}>Metronet · Inventory Reconciliation</span>
        <span>All processing happens in your browser — no data is sent to any server.</span>
      </div>
    </div>
  )
}
