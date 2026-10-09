import { useLayoutEffect, useMemo, useRef, useState } from 'react'
import { normKey } from '../lib/excelParser'
import { charsToPx, ptToPx } from '../lib/sheetModel'
import { CELL_INSET_PX } from '../lib/textGeom'

const LANE_H = 62 // tinggi lajur TTD di preview (px) — di file asli baris ini kosong
const SIG_H = 56 // tinggi gambar TTD di preview (px)

/** Zoom proporsional agar muat layar TANPA merusak perbandingan font/kolom. */
function useFitZoom(totalPx, enabled) {
  const ref = useRef(null)
  const [cw, setCw] = useState(0)
  useLayoutEffect(() => {
    if (!enabled) return
    const el = ref.current
    if (!el) return
    const ro = new ResizeObserver((es) => {
      const w = es[0]?.contentRect?.width || 0
      setCw(w)
    })
    ro.observe(el)
    setCw(el.clientWidth || 0)
    return () => ro.disconnect()
  }, [enabled])
  const zoom = enabled && cw > 0 && totalPx > 0 ? Math.min(1, (cw - 2) / totalPx) : 1
  return [ref, zoom]
}

/**
 * Preview SEMIRIP Excel: tabel di-render 1:1 dari model sheet —
 * nilai, merge, font, warna, border, alignment, tinggi baris & lebar kolom
 * persis seperti file yang di-upload. TTD tampil di lajur kosong tepat di
 * atas baris nama (baris yang sama dipakai saat export Excel).
 */
export default function ReportPreview({
  sheet, blocks, signatures,
  fileUid = 'f', spotMissing = false, fit = true, onEditSig,
}) {
  const pages = useMemo(() => {
    if (!sheet || !blocks?.length) return []
    return blocks.map((b, i) => ({
      n: b.page,
      block: b,
      start: i === 0 ? 0 : blocks[i - 1].nameRow + 2,
      end: i === blocks.length - 1 ? sheet.nrows : b.nameRow + 2,
    }))
  }, [sheet, blocks])

  const colWpx = useMemo(
    () => (sheet ? sheet.colW.map(charsToPx) : []),
    [sheet]
  )
  const totalPx = useMemo(() => colWpx.reduce((a, w) => a + w, 0), [colWpx])
  const [wrapRef, zoom] = useFitZoom(totalPx, fit)

  if (!sheet || !pages.length) {
    return <div className="card mut">Model preview belum siap — coba upload ulang file.</div>
  }

  const colOffset = (c0) => {
    let w = 0
    for (let c = 0; c < c0; c++) w += colWpx[c] || 0
    return w
  }
  const colSpanW = (c0, span) => {
    let w = 0
    for (let c = c0; c < c0 + span; c++) w += colWpx[c] || 0
    return Math.max(w, 1)
  }

  // Posisi horizontal tiap TTD di atas tengah namanya (% dari lebar penuh),
  // memakai geometri teks yang sama dengan export Excel.
  const sigSpots = (block) => {
    const out = []
    ;(block.slots || []).forEach((slot) => {
      const off = colOffset(slot.nameCol)
      const wPx = colSpanW(slot.nameCol, slot.spanCols || 1)
      const n = Math.max(slot.names.length, 1)
      slot.names.forEach((nm, i) => {
        let cx
        if (slot.geom && slot.geom.centers && slot.geom.centers[i] != null && slot.geom.centers.length === slot.names.length) {
          cx = CELL_INSET_PX + slot.geom.centers[i]
        } else {
          cx = (wPx / n) * (i + 0.5)
        }
        out.push({ slot, name: nm, key: normKey(nm), leftPct: Math.min(100, Math.max(0, ((off + cx) / totalPx) * 100)) })
      })
    })
    return out
  }

  const renderLane = (block) => (
    <div className="xlane" style={{ height: LANE_H }}>
      {sigSpots(block).map((s) => {
        const sig = signatures[s.key]
        const on = sig?.dataUrl && sig?.enabled !== false
        if (!on && !spotMissing) return null
        return (
          <button
            key={`${s.slot.nameCol}-${s.name}`}
            className={`xsig${on ? '' : ' miss'}`}
            style={{ left: `${s.leftPct}%` }}
            title={on ? `${s.name} (${s.slot.role}) — klik untuk ubah` : `${s.name} (${s.slot.role}) — belum TTD, klik untuk isi`}
            onClick={() => onEditSig?.(s.key)}
          >
            {on
              ? <img src={sig.dataUrl} alt={s.name} style={{ height: SIG_H }} />
              : <span className="xmissbox">✍️ {s.name}</span>}
          </button>
        )
      })}
    </div>
  )

  const cellStyle = (cell, rowHPt) => {
    if (!cell) return { td: { padding: '1px 3px' }, box: null, wrap: false }
    const f = sheet.fonts[cell.f]
    const bd = sheet.bords[cell.bd]
    const al = sheet.aligns[cell.al]
    // line-height rapat ala Excel: teks pas dalam tinggi barisnya.
    const lh = Math.max(1.02, Math.min(1.25, (rowHPt || 15) / (f.size || 11)))
    const wrap = al.wrap === true
    return {
      td: {
        font: f.css,
        color: f.color,
        textDecoration: f.underline ? 'underline' : undefined,
        background: sheet.fills[cell.bg] || undefined,
        borderLeft: bd[0] || undefined,
        borderRight: bd[1] || undefined,
        borderTop: bd[2] || undefined,
        borderBottom: bd[3] || undefined,
        textAlign: al.h,
        verticalAlign: al.v,
        whiteSpace: wrap ? 'pre-wrap' : 'nowrap',
        lineHeight: wrap ? undefined : 1.05, // teks 1 baris: kotak pas (Excel ~1.0em)
        padding: wrap ? 0 : `1px ${3 + (al.indent || 0) * 10}px`,
      },
      // Sel wrap: bungkus diklip setinggi baris (Excel memotong teks berlebih).
      // Sel satu baris: teks langsung (meluber ke sel kosong tetangga ala Excel).
      box: wrap ? {
        lineHeight: lh,
        textAlign: al.h,
        whiteSpace: 'pre-wrap',
        // Excel membungkus di batas kata; pecah tengah kata hanya bila terpaksa.
        overflowWrap: 'break-word',
        wordBreak: 'normal',
        overflow: 'hidden',
        padding: `0 ${3 + (al.indent || 0) * 10}px`,
        display: 'flex',
        flexDirection: 'column',
        justifyContent: al.v === 'middle' ? 'center' : al.v === 'bottom' ? 'flex-end' : 'flex-start',
      } : null,
      wrap,
    }
  }

  // Tinggi blokir teks (px) agar klip tepat di batas baris seperti Excel.
  const boxH = (r, rs) => {
    let h = 0
    for (let k = 0; k < (rs || 1); k++) h += ptToPx(sheet.rows[r + k]?.h)
    return Math.max(h - 1, 4)
  }

  return (
    <div className="xreport">
      {pages.map((pg) => {
        const b = pg.block
        const laneRow = b.nameRow - 1 // lajur kosong di atas baris nama (sama dgn export)
        return (
          <section key={pg.n} id={`${fileUid}-pg${pg.n}`} className="xpage">
            <div className="xpagetag">
              Hlm {pg.n}
            </div>
            <div className="xtable-wrap" ref={pg.n === pages[0]?.n ? wrapRef : undefined}>
              <table className="xsheet" style={{ width: totalPx, zoom }}>
                <colgroup>
                  {colWpx.map((w, c) => (
                    <col key={c} style={{ width: w }} />
                  ))}
                </colgroup>
                <tbody>
                  {Array.from({ length: pg.end - pg.start }, (_, k) => {
                    const r = pg.start + k
                    const row = sheet.rows[r]
                    if (!row) return null
                    // Lajur TTD: baris kosong diperlebar di preview agar TTD terbaca
                    // (di file export Excel baris ini pun ditinggikan ke 72pt).
                    if (r === laneRow) {
                      const st = cellStyle(row.cells[0] || null, row.h)
                      return (
                        <tr key={r} className="xlanerow">
                          <td colSpan={Math.max(1, sheet.ncols)} style={{ ...st.td, padding: 2 }}>
                            {renderLane(b)}
                          </td>
                        </tr>
                      )
                    }
                    return (
                      <tr key={r} style={{ height: ptToPx(row.h) }}>
                        {Array.from({ length: sheet.ncols }, (_, c) => {
                          if (row.skip[c]) return null // tertutup merge dari sel lain
                          const cell = row.cells[c]
                          const span = row.spans[c]
                          const rs = span && span.rs > 1 ? Math.min(span.rs, pg.end - r) : 1
                          const st = cellStyle(cell || null, row.h)
                          return (
                            <td
                              key={c}
                              colSpan={span ? Math.min(span.cs, sheet.ncols - c) : undefined}
                              rowSpan={rs > 1 ? rs : undefined}
                              style={st.td}
                            >
                              {st.box
                                ? <div className="xcell" style={{ ...st.box, height: boxH(r, rs) }}>{cell ? cell.v : ''}</div>
                                : (cell ? cell.v : '')}
                            </td>
                          )
                        })}
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          </section>
        )
      })}
    </div>
  )
}
