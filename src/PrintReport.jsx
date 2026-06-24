import React from 'react'

export default function PrintReport({ reportName, date, summary, eabRec, wsiRec, ws2Rec, eabShipped, wsiShipped, ws2Shipped, eabReported, wsiReported, ws2Reported, overallPct, allIssues, onClose }) {

  const handlePrint = () => window.print()

  const warehouseSections = [
    { name: 'EAB', lines: eabRec, shipped: eabShipped, reported: eabReported },
    { name: 'WSI', lines: wsiRec, shipped: wsiShipped, reported: wsiReported },
    { name: 'WS2', lines: ws2Rec, shipped: ws2Shipped, reported: ws2Reported },
  ]

  const statusLabel = { ok: 'Match', missing: 'Missing', short: 'Short', over: 'Over', 'unmatched-warehouse': 'Not Shipped' }

  return (
    <>
      {/* screen overlay */}
      <div className="print-overlay" style={{
        position: 'fixed', inset: 0, background: 'rgba(0,0,0,0.85)',
        zIndex: 1000, display: 'flex', alignItems: 'flex-start', justifyContent: 'center',
        padding: '32px 16px', overflowY: 'auto'
      }}>
        <div style={{ background: '#fff', borderRadius: 10, width: '100%', maxWidth: 860, color: '#111', position: 'relative' }}>

          {/* screen-only toolbar */}
          <div className="no-print" style={{
            padding: '16px 24px', borderBottom: '1px solid #e5e7eb',
            display: 'flex', alignItems: 'center', justifyContent: 'space-between', gap: 12
          }}>
            <span style={{ fontWeight: 700, fontSize: 15, color: '#111' }}>Print preview — {reportName}</span>
            <div style={{ display: 'flex', gap: 10 }}>
              <button onClick={onClose} style={{
                padding: '8px 18px', borderRadius: 6, border: '1px solid #d1d5db',
                background: '#fff', fontSize: 14, fontWeight: 600, cursor: 'pointer', color: '#374151'
              }}>Cancel</button>
              <button onClick={handlePrint} style={{
                padding: '8px 20px', borderRadius: 6, border: 'none',
                background: '#1d4ed8', color: '#fff', fontSize: 14, fontWeight: 700, cursor: 'pointer'
              }}>Print / Save as PDF</button>
            </div>
          </div>

          {/* report body */}
          <div id="print-body" style={{ padding: '40px 48px', fontFamily: 'Arial, sans-serif' }}>

            {/* header */}
            <div style={{ borderBottom: '3px solid #1d4ed8', paddingBottom: 18, marginBottom: 28 }}>
              <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start' }}>
                <div>
                  <div style={{ fontSize: 11, color: '#6b7280', textTransform: 'uppercase', letterSpacing: '0.1em', fontWeight: 700, marginBottom: 4 }}>Metronet</div>
                  <h1 style={{ fontSize: 26, fontWeight: 800, color: '#111', margin: 0, letterSpacing: '-0.02em' }}>Inventory Reconciliation Report</h1>
                  <div style={{ fontSize: 14, color: '#6b7280', marginTop: 6 }}>{reportName}</div>
                </div>
                <div style={{ textAlign: 'right' }}>
                  <div style={{ fontSize: 11, color: '#6b7280', textTransform: 'uppercase', letterSpacing: '0.08em', fontWeight: 700, marginBottom: 4 }}>Generated</div>
                  <div style={{ fontSize: 14, color: '#374151', fontWeight: 600 }}>{date}</div>
                  <div style={{ fontSize: 28, fontWeight: 800, color: overallPct === 100 ? '#15803d' : overallPct >= 80 ? '#b45309' : '#dc2626', marginTop: 6 }}>{overallPct}%</div>
                  <div style={{ fontSize: 12, color: '#6b7280' }}>Overall reconciled</div>
                </div>
              </div>
            </div>

            {/* executive summary */}
            <h2 style={{ fontSize: 14, fontWeight: 800, textTransform: 'uppercase', letterSpacing: '0.08em', color: '#374151', marginBottom: 14 }}>Executive Summary</h2>
            <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: 32, fontSize: 14 }}>
              <thead>
                <tr style={{ background: '#1d4ed8', color: '#fff' }}>
                  {['Warehouse', 'Pallets Shipped', 'Pallets Reported', 'Variance', 'Parts Reconciled', 'Exceptions'].map(h => (
                    <th key={h} style={{ padding: '9px 14px', textAlign: h === 'Warehouse' ? 'left' : 'right', fontWeight: 700, fontSize: 12 }}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {warehouseSections.map((w, i) => {
                  const v = w.reported - w.shipped
                  const matched = w.lines.filter(l => l.status === 'ok').length
                  const pct = w.lines.length ? Math.round((matched / w.lines.length) * 100) : 0
                  const issues = w.lines.filter(l => l.status !== 'ok').length
                  return (
                    <tr key={w.name} style={{ background: i % 2 === 0 ? '#f9fafb' : '#fff', borderBottom: '1px solid #e5e7eb' }}>
                      <td style={{ padding: '9px 14px', fontWeight: 700 }}>{w.name}</td>
                      <td style={{ padding: '9px 14px', textAlign: 'right', fontFamily: 'monospace' }}>{w.shipped}</td>
                      <td style={{ padding: '9px 14px', textAlign: 'right', fontFamily: 'monospace' }}>{w.reported}</td>
                      <td style={{ padding: '9px 14px', textAlign: 'right', fontFamily: 'monospace', color: v === 0 ? '#374151' : v > 0 ? '#b45309' : '#dc2626', fontWeight: 600 }}>
                        {v === 0 ? '—' : (v > 0 ? '+' : '') + v}
                      </td>
                      <td style={{ padding: '9px 14px', textAlign: 'right', fontWeight: 700, color: pct === 100 ? '#15803d' : pct >= 80 ? '#b45309' : '#dc2626' }}>{pct}%</td>
                      <td style={{ padding: '9px 14px', textAlign: 'right', color: issues > 0 ? '#dc2626' : '#15803d', fontWeight: 600 }}>{issues}</td>
                    </tr>
                  )
                })}
                {/* totals row */}
                <tr style={{ background: '#1e293b', color: '#fff', fontWeight: 700 }}>
                  <td style={{ padding: '9px 14px' }}>Total</td>
                  <td style={{ padding: '9px 14px', textAlign: 'right', fontFamily: 'monospace' }}>{eabShipped + wsiShipped + ws2Shipped}</td>
                  <td style={{ padding: '9px 14px', textAlign: 'right', fontFamily: 'monospace' }}>{eabReported + wsiReported + ws2Reported}</td>
                  <td style={{ padding: '9px 14px', textAlign: 'right', fontFamily: 'monospace' }}>
                    {(() => { const v = (eabReported+wsiReported+ws2Reported)-(eabShipped+wsiShipped+ws2Shipped); return v===0?'—':(v>0?'+':'')+v })()}
                  </td>
                  <td style={{ padding: '9px 14px', textAlign: 'right' }}>{overallPct}%</td>
                  <td style={{ padding: '9px 14px', textAlign: 'right' }}>{allIssues.length}</td>
                </tr>
              </tbody>
            </table>

            {/* exceptions */}
            {allIssues.length > 0 && (
              <>
                <h2 style={{ fontSize: 14, fontWeight: 800, textTransform: 'uppercase', letterSpacing: '0.08em', color: '#374151', marginBottom: 14 }}>Exceptions Requiring Attention</h2>
                <table style={{ width: '100%', borderCollapse: 'collapse', marginBottom: 32, fontSize: 14 }}>
                  <thead>
                    <tr style={{ background: '#dc2626', color: '#fff' }}>
                      {['Warehouse', 'Part Number', 'Shipped', 'Reported', 'Variance', 'Status'].map(h => (
                        <th key={h} style={{ padding: '9px 14px', textAlign: h === 'Warehouse' || h === 'Part Number' || h === 'Status' ? 'left' : 'right', fontWeight: 700, fontSize: 12 }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {allIssues.map((issue, i) => {
                      const wh = eabRec.includes(issue) ? 'EAB' : wsiRec.includes(issue) ? 'WSI' : 'WS2'
                      return (
                        <tr key={i} style={{ background: i % 2 === 0 ? '#fef2f2' : '#fff', borderBottom: '1px solid #fecaca' }}>
                          <td style={{ padding: '9px 14px', fontWeight: 700 }}>{wh}</td>
                          <td style={{ padding: '9px 14px', fontFamily: 'monospace', fontWeight: 700 }}>{issue.part}</td>
                          <td style={{ padding: '9px 14px', textAlign: 'right', fontFamily: 'monospace' }}>{issue.shippedQty}</td>
                          <td style={{ padding: '9px 14px', textAlign: 'right', fontFamily: 'monospace' }}>{issue.reportedQty}</td>
                          <td style={{ padding: '9px 14px', textAlign: 'right', fontFamily: 'monospace', fontWeight: 700, color: issue.variance > 0 ? '#b45309' : '#dc2626' }}>
                            {(issue.variance > 0 ? '+' : '') + issue.variance}
                          </td>
                          <td style={{ padding: '9px 14px', fontWeight: 600, color: issue.status === 'missing' ? '#dc2626' : issue.status === 'unmatched-warehouse' ? '#7c3aed' : '#b45309' }}>
                            {statusLabel[issue.status] || issue.status}
                          </td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              </>
            )}

            {allIssues.length === 0 && (
              <div style={{ textAlign: 'center', padding: '24px', background: '#f0fdf4', border: '1px solid #bbf7d0', borderRadius: 8, marginBottom: 32, color: '#15803d', fontWeight: 700, fontSize: 15 }}>
                ✓ All pallets fully reconciled across all three warehouses.
              </div>
            )}

            {/* per-warehouse detail */}
            {warehouseSections.map(w => (
              <div key={w.name} style={{ marginBottom: 32, pageBreakInside: 'avoid' }}>
                <h2 style={{ fontSize: 14, fontWeight: 800, textTransform: 'uppercase', letterSpacing: '0.08em', color: '#374151', marginBottom: 14 }}>{w.name} — Full Detail</h2>
                <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
                  <thead>
                    <tr style={{ background: '#374151', color: '#fff' }}>
                      {['Part Number', 'Shipped', 'Reported', 'Variance', 'Status'].map(h => (
                        <th key={h} style={{ padding: '8px 12px', textAlign: h === 'Part Number' || h === 'Status' ? 'left' : 'right', fontWeight: 700, fontSize: 12 }}>{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {w.lines.map((line, i) => (
                      <tr key={line.part} style={{ background: line.status !== 'ok' ? '#fef9c3' : i % 2 === 0 ? '#f9fafb' : '#fff', borderBottom: '1px solid #e5e7eb' }}>
                        <td style={{ padding: '7px 12px', fontFamily: 'monospace', fontWeight: 600 }}>{line.part}</td>
                        <td style={{ padding: '7px 12px', textAlign: 'right', fontFamily: 'monospace' }}>{line.shippedQty}</td>
                        <td style={{ padding: '7px 12px', textAlign: 'right', fontFamily: 'monospace' }}>{line.reportedQty}</td>
                        <td style={{ padding: '7px 12px', textAlign: 'right', fontFamily: 'monospace', fontWeight: line.variance !== 0 ? 700 : 400, color: line.variance > 0 ? '#b45309' : line.variance < 0 ? '#dc2626' : '#9ca3af' }}>
                          {line.variance === 0 ? '—' : (line.variance > 0 ? '+' : '') + line.variance}
                        </td>
                        <td style={{ padding: '7px 12px', fontWeight: 600, color: line.status === 'ok' ? '#15803d' : line.status === 'missing' ? '#dc2626' : line.status === 'unmatched-warehouse' ? '#7c3aed' : '#b45309' }}>
                          {statusLabel[line.status] || line.status}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ))}

            {/* footer */}
            <div style={{ borderTop: '1px solid #e5e7eb', paddingTop: 14, marginTop: 8, display: 'flex', justifyContent: 'space-between', fontSize: 11, color: '#9ca3af' }}>
              <span>Metronet · Inventory Reconciliation</span>
              <span>Generated {date} · Confidential</span>
            </div>

          </div>
        </div>
      </div>

      {/* print-only styles */}
      <style>{`
        @media print {
          body * { visibility: hidden; }
          #print-body, #print-body * { visibility: visible; }
          #print-body { position: fixed; top: 0; left: 0; width: 100%; padding: 24px 32px; font-size: 12px; }
          .no-print { display: none !important; }
          .print-overlay { display: block !important; position: static !important; background: none !important; padding: 0 !important; }
          @page { margin: 0.75in; size: letter; }
        }
      `}</style>
    </>
  )
}
