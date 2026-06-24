import React, { useState, useCallback } from 'react'

async function loadPdfJs() {
  if (window.pdfjsLib) return window.pdfjsLib
  return new Promise((resolve, reject) => {
    const script = document.createElement('script')
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js'
    script.onload = () => {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc =
        'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js'
      resolve(window.pdfjsLib)
    }
    script.onerror = reject
    document.head.appendChild(script)
  })
}

async function readPdfBOL(file) {
  const pdfjsLib = await loadPdfJs()
  const arrayBuffer = await file.arrayBuffer()
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise
  let fullText = ''
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i)
    const content = await page.getTextContent()
    fullText += content.items.map(item => item.str).join(' ') + ' '
  }
  return parseBOLText(fullText, file.name)
}

function parseBOLText(text, filename) {
  const t = text.replace(/\s+/g, ' ').trim()

  // Date — PDF has spaces: "Date 5 - 19 - 26"
  const dateMatch = t.match(/Date\s+([\d]{1,2})\s*[-\/]\s*([\d]{1,2})\s*[-\/]\s*([\d]{2,4})/)
  const date = dateMatch ? `${dateMatch[1]}-${dateMatch[2]}-${dateMatch[3]}` : ''

  // Carrier
  const carrierRaw = t.match(/\b(Walts|Tyme|Werner|Estes|FedEx|UPS|XPO|Averitt|Old Dominion|Saia)\b/i)
  const carrier = carrierRaw ? carrierRaw[1] : ''

  // Warehouse addresses
  const ADDRESSES = {
    metronet: ['300 E Walnut', 'Metronet', '47713'],
    wsi:      ['Wedeking', '47711'],
    ws2:      ['Pennell', '42420', 'Henderson'],
    eab:      ['Second Ave', '47710', 'Blind', 'Evansville Association'],
  }

  const shipFromIdx = t.search(/SHIP FROM/i)
  const shipToIdx   = t.search(/SHIP TO/i)
  const shipFromCtx = shipFromIdx >= 0 ? t.slice(shipFromIdx, shipFromIdx + 300) : ''
  const shipToCtx   = shipToIdx   >= 0 ? t.slice(shipToIdx,   shipToIdx   + 300) : ''

  let direction = 'outbound'
  let warehouse = 'Unknown'

  const fromIsMetronet = ADDRESSES.metronet.some(a => shipFromCtx.includes(a))
  const toIsMetronet   = ADDRESSES.metronet.some(a => shipToCtx.includes(a))

  if (fromIsMetronet && !toIsMetronet) {
    direction = 'outbound'
    if      (ADDRESSES.wsi.some(a => shipToCtx.includes(a)))  warehouse = 'WSI'
    else if (ADDRESSES.ws2.some(a => shipToCtx.includes(a)))  warehouse = 'WS2'
    else if (ADDRESSES.eab.some(a => shipToCtx.includes(a)))  warehouse = 'EAB'
  } else if (toIsMetronet && !fromIsMetronet) {
    direction = 'inbound'
    if      (ADDRESSES.wsi.some(a => shipFromCtx.includes(a)))  warehouse = 'WSI'
    else if (ADDRESSES.ws2.some(a => shipFromCtx.includes(a)))  warehouse = 'WS2'
    else if (ADDRESSES.eab.some(a => shipFromCtx.includes(a)))  warehouse = 'EAB'
  } else {
    for (const [wh, keys] of [['WSI', ADDRESSES.wsi], ['WS2', ADDRESSES.ws2], ['EAB', ADDRESSES.eab]]) {
      for (const key of keys) {
        const keyIdx = t.indexOf(key)
        if (keyIdx === -1) continue
        warehouse = wh
        const distFrom = Math.abs(keyIdx - shipFromIdx)
        const distTo   = Math.abs(keyIdx - shipToIdx)
        direction = distFrom < distTo ? 'inbound' : 'outbound'
        break
      }
      if (warehouse !== 'Unknown') break
    }
  }

  // Pallet count from Grand Total
  const grandMatch = t.match(/Grand Total\s+(\d+)/i)
  const totalPallets = grandMatch ? parseInt(grandMatch[1]) : 0

  // Line items — part number and pallet count only
  const lineItems = []
  const skip = new Set(['14706','47713','47710','47711','42420','360','1','49'])
  const itemPattern = /\b(\d{5,7})\s+(\d{1,3})\b/g
  let m
  while ((m = itemPattern.exec(t)) !== null) {
    const part = m[1], pallets = parseInt(m[2])
    if (skip.has(part) || pallets < 1 || pallets > 200) continue
    if (!lineItems.find(x => x.partNumber === part)) lineItems.push({ partNumber: part, pallets })
  }

  return {
    id: Date.now() + Math.random(),
    filename, date, carrier,
    direction, warehouse,
    lineItems, totalPallets,
    uploadedAt: new Date().toISOString(),
  }
}

function BOLCard({ bol, onRemove }) {
  const [expanded, setExpanded] = useState(false)
  const isOutbound = bol.direction === 'outbound'

  const fromLabel = isOutbound ? 'Metronet → 300 E Walnut, Evansville IN' : {
    WSI: 'WSI → 1147 Wedeking Ave, Evansville IN',
    WS2: 'WS2 → 701 Pennell St, Henderson KY',
    EAB: 'EAB → 500 N Second Ave, Evansville IN',
  }[bol.warehouse] || `${bol.warehouse}`

  const toLabel = isOutbound ? {
    WSI: 'WSI · 1147 Wedeking Ave, Evansville IN',
    WS2: 'WS2 · 701 Pennell St, Henderson KY',
    EAB: 'EAB · 500 N Second Ave, Evansville IN',
  }[bol.warehouse] || bol.warehouse : 'Metronet · 300 E Walnut, Evansville IN'

  return (
    <div style={{ background:'var(--bg-card)', border:'1px solid var(--border)', borderLeft:`3px solid ${isOutbound?'var(--accent-light)':'#3dba78'}`, borderRadius:'var(--radius-lg)', marginBottom:10, overflow:'hidden' }}>
      <div onClick={()=>setExpanded(!expanded)} style={{ padding:'16px 20px', cursor:'pointer', display:'flex', alignItems:'center', gap:14, flexWrap:'wrap' }}>

        {/* Direction badge */}
        <span style={{ fontSize:11, fontWeight:700, padding:'3px 10px', borderRadius:4, background:isOutbound?'#0a1525':'#071a0f', color:isOutbound?'var(--accent-light)':'#3dba78', border:`1px solid ${isOutbound?'#1e3a5a':'#1a4d2e'}`, textTransform:'uppercase', letterSpacing:'0.06em', whiteSpace:'nowrap' }}>
          {isOutbound ? `→ To ${bol.warehouse}` : `← From ${bol.warehouse}`}
        </span>

        {/* Date */}
        <span style={{ fontSize:15, fontWeight:700, color:'var(--text-primary)' }}>{bol.date || 'No date'}</span>

        {/* Carrier */}
        {bol.carrier && <span style={{ fontSize:14, color:'var(--text-secondary)' }}>{bol.carrier}</span>}

        {/* Pallet count */}
        <span style={{ marginLeft:'auto', fontSize:15, fontWeight:700, fontFamily:'var(--mono)', color:'var(--text-primary)' }}>
          {bol.totalPallets} pallets
        </span>

        <span style={{ fontSize:12, color:'var(--text-muted)' }}>{expanded?'▲':'▼'}</span>
        <button onClick={e=>{e.stopPropagation();onRemove(bol.id)}} style={{ background:'transparent', border:'none', color:'var(--text-muted)', fontSize:18, cursor:'pointer', padding:'0 4px', lineHeight:1 }}>×</button>
      </div>

      {expanded && (
        <div style={{ borderTop:'1px solid var(--border)', padding:'16px 20px' }}>
          {/* From / To addresses */}
          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:16, marginBottom:16 }}>
            <div style={{ background:'var(--bg-input)', borderRadius:'var(--radius)', padding:'12px 16px' }}>
              <div style={{ fontSize:11, color:'var(--text-muted)', textTransform:'uppercase', letterSpacing:'0.08em', fontWeight:700, marginBottom:6 }}>From</div>
              <div style={{ fontSize:14, color:'var(--text-primary)', fontWeight:600 }}>
                {isOutbound ? 'Metronet' : bol.warehouse === 'WSI' ? 'WSI — Wedeking' : bol.warehouse === 'WS2' ? 'WS2 — Pennell St' : 'EAB'}
              </div>
              <div style={{ fontSize:13, color:'var(--text-secondary)', marginTop:2 }}>
                {isOutbound ? '300 E Walnut St, Evansville IN 47713' :
                  bol.warehouse === 'WSI' ? '1147 Wedeking Ave, Evansville IN 47711' :
                  bol.warehouse === 'WS2' ? '701 Pennell St, Henderson KY 42420' :
                  '500 N Second Ave, Evansville IN 47710'}
              </div>
            </div>
            <div style={{ background:'var(--bg-input)', borderRadius:'var(--radius)', padding:'12px 16px' }}>
              <div style={{ fontSize:11, color:'var(--text-muted)', textTransform:'uppercase', letterSpacing:'0.08em', fontWeight:700, marginBottom:6 }}>To</div>
              <div style={{ fontSize:14, color:'var(--text-primary)', fontWeight:600 }}>
                {!isOutbound ? 'Metronet' : bol.warehouse === 'WSI' ? 'WSI — Wedeking' : bol.warehouse === 'WS2' ? 'WS2 — Pennell St' : 'EAB'}
              </div>
              <div style={{ fontSize:13, color:'var(--text-secondary)', marginTop:2 }}>
                {!isOutbound ? '300 E Walnut St, Evansville IN 47713' :
                  bol.warehouse === 'WSI' ? '1147 Wedeking Ave, Evansville IN 47711' :
                  bol.warehouse === 'WS2' ? '701 Pennell St, Henderson KY 42420' :
                  '500 N Second Ave, Evansville IN 47710'}
              </div>
            </div>
          </div>

          {/* Line items */}
          {bol.lineItems?.length > 0 && (
            <table style={{ width:'100%', fontSize:13 }}>
              <thead>
                <tr style={{ borderBottom:'1px solid var(--border)' }}>
                  {['Part Number','# Pallets'].map(h=>(
                    <th key={h} style={{ padding:'6px 10px', color:'var(--text-muted)', fontWeight:700, fontSize:11, textTransform:'uppercase', letterSpacing:'0.06em' }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {bol.lineItems.map((item,i)=>(
                  <tr key={i} style={{ borderBottom:'1px solid var(--border)' }}>
                    <td style={{ padding:'8px 10px', fontFamily:'var(--mono)', fontWeight:600 }}>{item.partNumber}</td>
                    <td style={{ padding:'8px 10px', fontFamily:'var(--mono)', textAlign:'right' }}>{item.pallets}</td>
                  </tr>
                ))}
                <tr style={{ background:'var(--bg-input)', fontWeight:700 }}>
                  <td style={{ padding:'8px 10px', color:'var(--text-muted)', fontSize:11, textTransform:'uppercase' }}>Total</td>
                  <td style={{ padding:'8px 10px', fontFamily:'var(--mono)', textAlign:'right' }}>{bol.totalPallets}</td>
                </tr>
              </tbody>
            </table>
          )}
          <div style={{ fontSize:11, color:'var(--text-muted)', marginTop:10 }}>File: {bol.filename}</div>
        </div>
      )}
    </div>
  )
}

export default function BillOfLanding({ bols, onAddBOL, onRemoveBOL }) {
  const [dragging, setDragging] = useState(false)
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const [filterWarehouse, setFilterWarehouse] = useState('all')
  const [filterDirection, setFilterDirection] = useState('all')

  const handleFiles = useCallback(async (files) => {
    setLoading(true); setError('')
    for (const file of Array.from(files)) {
      try { onAddBOL(await readPdfBOL(file)) }
      catch(err) { console.error(err); setError(`Could not read ${file.name}. Make sure it's a PDF.`) }
    }
    setLoading(false)
  }, [onAddBOL])

  const outbound = bols.filter(b=>b.direction==='outbound')
  const inbound  = bols.filter(b=>b.direction==='inbound')
  const byWH = {
    EAB: bols.filter(b=>b.warehouse==='EAB'),
    WSI: bols.filter(b=>b.warehouse==='WSI'),
    WS2: bols.filter(b=>b.warehouse==='WS2'),
  }
  const filtered = bols
    .filter(b=>(filterWarehouse==='all'||b.warehouse===filterWarehouse)&&(filterDirection==='all'||b.direction===filterDirection))
    .sort((a,b)=>(b.uploadedAt||'').localeCompare(a.uploadedAt||''))

  return (
    <div>
      <div style={{ marginBottom:24 }}>
        <h2 style={{ fontSize:20, fontWeight:700, marginBottom:6 }}>Bill of Landing Log</h2>
        <p style={{ fontSize:15, color:'var(--text-secondary)' }}>Upload Bill of Landing PDFs to track shipments to and from each warehouse.</p>
      </div>

      {bols.length > 0 && (
        <div style={{ display:'grid', gridTemplateColumns:'repeat(auto-fit, minmax(150px, 1fr))', gap:12, marginBottom:28 }}>
          <div style={{ background:'var(--bg-card)', border:'1px solid var(--border)', borderTop:'3px solid var(--accent)', borderRadius:'var(--radius-lg)', padding:'16px 20px' }}>
            <div style={{ fontSize:11, color:'var(--text-muted)', textTransform:'uppercase', letterSpacing:'0.1em', fontWeight:700, marginBottom:8 }}>Total BOLs</div>
            <div style={{ fontSize:30, fontWeight:800, fontFamily:'var(--mono)' }}>{bols.length}</div>
            <div style={{ fontSize:13, color:'var(--text-secondary)', marginTop:6 }}>{outbound.length} out · {inbound.length} in</div>
          </div>
          {['EAB','WSI','WS2'].map(wh => {
            const w = byWH[wh]
            return (
              <div key={wh} style={{ background:'var(--bg-card)', border:'1px solid var(--border)', borderTop:'3px solid var(--border-mid)', borderRadius:'var(--radius-lg)', padding:'16px 20px' }}>
                <div style={{ fontSize:11, color:'var(--text-muted)', textTransform:'uppercase', letterSpacing:'0.1em', fontWeight:700, marginBottom:8 }}>{wh}</div>
                <div style={{ fontSize:30, fontWeight:800, fontFamily:'var(--mono)' }}>{w.length}</div>
                <div style={{ fontSize:13, color:'var(--text-secondary)', marginTop:6 }}>
                  {w.filter(b=>b.direction==='outbound').length} out · {w.filter(b=>b.direction==='inbound').length} in · {w.reduce((a,b)=>a+(b.totalPallets||0),0)} pallets
                </div>
              </div>
            )
          })}
          <div style={{ background:'var(--bg-card)', border:'1px solid var(--border)', borderTop:'3px solid var(--border-mid)', borderRadius:'var(--radius-lg)', padding:'16px 20px' }}>
            <div style={{ fontSize:11, color:'var(--text-muted)', textTransform:'uppercase', letterSpacing:'0.1em', fontWeight:700, marginBottom:8 }}>Total Pallets</div>
            <div style={{ fontSize:30, fontWeight:800, fontFamily:'var(--mono)' }}>{bols.reduce((a,b)=>a+(b.totalPallets||0),0).toLocaleString()}</div>
            <div style={{ fontSize:13, color:'var(--text-secondary)', marginTop:6 }}>this month</div>
          </div>
        </div>
      )}

      <div onDragOver={e=>{e.preventDefault();setDragging(true)}} onDragLeave={()=>setDragging(false)}
        onDrop={e=>{e.preventDefault();setDragging(false);handleFiles(e.dataTransfer.files)}}
        onClick={()=>{const i=document.createElement('input');i.type='file';i.accept='.pdf';i.multiple=true;i.onchange=ev=>handleFiles(ev.target.files);i.click()}}
        style={{ border:`1.5px dashed ${dragging?'var(--accent-light)':'var(--border-mid)'}`, borderRadius:'var(--radius-lg)', background:dragging?'#0a1525':'var(--bg-input)', padding:'28px 20px', cursor:'pointer', textAlign:'center', marginBottom:24, transition:'all 0.15s', userSelect:'none' }}>
        <div style={{ fontSize:24, marginBottom:10, color:'var(--text-muted)' }}>{loading?'⏳':'↑'}</div>
        <div style={{ fontWeight:600, fontSize:15, marginBottom:4, color:'var(--text-primary)' }}>{loading?'Reading...':'Upload Bill of Landing PDFs'}</div>
        <div style={{ fontSize:13, color:'var(--text-muted)' }}>Drop .pdf files here, or click to browse. Multiple files supported.</div>
      </div>

      {error && <div style={{ background:'#1a0808', border:'1px solid #5a1515', borderRadius:'var(--radius)', padding:'12px 16px', marginBottom:16, fontSize:14, color:'#f06060' }}>⚠ {error}</div>}

      {bols.length > 0 && (
        <>
          <div style={{ display:'flex', gap:8, marginBottom:16, flexWrap:'wrap', alignItems:'center' }}>
            <span style={{ fontSize:13, color:'var(--text-muted)', fontWeight:600 }}>Filter:</span>
            {['all','EAB','WSI','WS2'].map(f=>(
              <button key={f} onClick={()=>setFilterWarehouse(f)} style={{ padding:'4px 14px', borderRadius:5, fontSize:13, fontWeight:600, background:filterWarehouse===f?'var(--accent)':'transparent', color:filterWarehouse===f?'#fff':'var(--text-secondary)', border:`1px solid ${filterWarehouse===f?'var(--accent)':'var(--border-mid)'}` }}>
                {f==='all'?'All warehouses':f}
              </button>
            ))}
            <div style={{ width:1, height:20, background:'var(--border)', margin:'0 4px' }} />
            {['all','outbound','inbound'].map(f=>(
              <button key={f} onClick={()=>setFilterDirection(f)} style={{ padding:'4px 14px', borderRadius:5, fontSize:13, fontWeight:600, background:filterDirection===f?'#0d2318':'transparent', color:filterDirection===f?'#3dba78':'var(--text-secondary)', border:`1px solid ${filterDirection===f?'#2d7a52':'var(--border-mid)'}` }}>
                {f==='all'?'Both directions':f==='outbound'?'→ Outbound':'← Inbound'}
              </button>
            ))}
            <span style={{ marginLeft:'auto', fontSize:13, color:'var(--text-muted)' }}>{filtered.length} record{filtered.length!==1?'s':''}</span>
          </div>
          {filtered.map(bol=><BOLCard key={bol.id} bol={bol} onRemove={onRemoveBOL}/>)}
        </>
      )}

      {bols.length===0 && !loading && (
        <div style={{ textAlign:'center', padding:'40px 20px', color:'var(--text-muted)', fontSize:15 }}>
          No Bills of Landing uploaded yet. Upload your first PDF above.
        </div>
      )}
    </div>
  )
}