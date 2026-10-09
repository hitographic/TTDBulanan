import { useEffect, useRef, useState } from 'react'
import { enhanceImageData } from '../lib/sigImage'

/**
 * Crop + enhance untuk foto/upload TTD (port dari Confidential/SignatureCropModal).
 * - Seret kotak / titik sudut / gambar kotak baru di luar area
 * - Enhance: background kertas jadi transparan, tinta jadi hitam pekat
 */
export default function SigCropModal({ src, onCancel, onApply }) {
  const editorRef = useRef(null)
  const previewRef = useRef(null)
  const dragRef = useRef(null)
  const [img, setImg] = useState(null)
  const [crop, setCrop] = useState({ x: 0.08, y: 0.3, w: 0.84, h: 0.4 })
  const [enhance, setEnhance] = useState(true)
  const [threshold, setThreshold] = useState(175)
  const [resultUrl, setResultUrl] = useState(null)

  useEffect(() => {
    if (!src) return
    const im = new Image()
    im.onload = () => {
      setImg(im)
      setCrop({ x: 0.08, y: 0.3, w: 0.84, h: 0.4 })
    }
    im.src = src
  }, [src])

  // Editor + overlay kotak crop
  useEffect(() => {
    const canvas = editorRef.current
    if (!img || !canvas) return
    const maxW = 480
    const scale = Math.min(maxW / img.naturalWidth, 300 / img.naturalHeight)
    canvas.width = Math.max(1, Math.round(img.naturalWidth * scale))
    canvas.height = Math.max(1, Math.round(img.naturalHeight * scale))
    const ctx = canvas.getContext('2d')
    ctx.drawImage(img, 0, 0, canvas.width, canvas.height)

    const r = { x: crop.x * canvas.width, y: crop.y * canvas.height, w: crop.w * canvas.width, h: crop.h * canvas.height }
    ctx.save()
    ctx.fillStyle = 'rgba(15,23,42,0.55)'
    ctx.fillRect(0, 0, canvas.width, r.y)
    ctx.fillRect(0, r.y, r.x, r.h)
    ctx.fillRect(r.x + r.w, r.y, canvas.width - (r.x + r.w), r.h)
    ctx.fillRect(0, r.y + r.h, canvas.width, canvas.height - (r.y + r.h))
    ctx.restore()

    ctx.save()
    ctx.strokeStyle = '#0ea5e9'
    ctx.lineWidth = 2
    ctx.setLineDash([6, 4])
    ctx.strokeRect(r.x, r.y, r.w, r.h)
    ctx.setLineDash([])
    const hs = 14
    ctx.fillStyle = '#fff'
    ctx.strokeStyle = '#0ea5e9'
    ;[[r.x, r.y], [r.x + r.w, r.y], [r.x, r.y + r.h], [r.x + r.w, r.y + r.h]].forEach(([cx, cy]) => {
      ctx.beginPath()
      ctx.rect(cx - hs / 2, cy - hs / 2, hs, hs)
      ctx.fill()
      ctx.stroke()
    })
    ctx.restore()
  }, [img, crop])

  // Hasil: crop + enhance -> preview
  useEffect(() => {
    const pv = previewRef.current
    if (!img || !pv) return
    const iw = img.naturalWidth, ih = img.naturalHeight
    const sx = Math.max(0, Math.round(crop.x * iw))
    const sy = Math.max(0, Math.round(crop.y * ih))
    let sw = Math.round(crop.w * iw), sh = Math.round(crop.h * ih)
    if (sw < 10 || sh < 10) return
    if (sx + sw > iw) sw = iw - sx
    if (sy + sh > ih) sh = ih - sy

    const tmp = document.createElement('canvas')
    tmp.width = sw
    tmp.height = sh
    const tctx = tmp.getContext('2d')
    tctx.drawImage(img, sx, sy, sw, sh, 0, 0, sw, sh)
    if (enhance) {
      try {
        const id = tctx.getImageData(0, 0, sw, sh)
        enhanceImageData(id.data, threshold)
        tctx.putImageData(id, 0, 0)
      } catch { /* tainted canvas — pakai crop mentah */ }
    }
    let url = null
    try { url = tmp.toDataURL('image/png') } catch { url = null }
    // eslint-disable-next-line react/set-state-in-effect -- hasil dibaca dari canvas (sistem eksternal)
    setResultUrl(url)

    const s = Math.min(440 / sw, 120 / sh)
    pv.width = Math.max(1, Math.round(sw * s))
    pv.height = Math.max(1, Math.round(sh * s))
    const pctx = pv.getContext('2d')
    pctx.clearRect(0, 0, pv.width, pv.height)
    pctx.drawImage(tmp, 0, 0, pv.width, pv.height)
  }, [img, crop, enhance, threshold])

  const toLocal = (e) => {
    const canvas = editorRef.current
    const r = canvas.getBoundingClientRect()
    return {
      x: (e.clientX - r.left) * (canvas.width / r.width),
      y: (e.clientY - r.top) * (canvas.height / r.height),
    }
  }

  const hit = (p, r) => {
    const t = 16
    if (Math.abs(p.x - r.x) < t && Math.abs(p.y - r.y) < t) return 'nw'
    if (Math.abs(p.x - (r.x + r.w)) < t && Math.abs(p.y - r.y) < t) return 'ne'
    if (Math.abs(p.x - r.x) < t && Math.abs(p.y - (r.y + r.h)) < t) return 'sw'
    if (Math.abs(p.x - (r.x + r.w)) < t && Math.abs(p.y - (r.y + r.h)) < t) return 'se'
    if (p.x >= r.x && p.x <= r.x + r.w && p.y >= r.y && p.y <= r.y + r.h) return 'move'
    return 'new'
  }

  const onDown = (e) => {
    const canvas = editorRef.current
    if (!canvas) return
    e.preventDefault()
    const p = toLocal(e)
    const r = { x: crop.x * canvas.width, y: crop.y * canvas.height, w: crop.w * canvas.width, h: crop.h * canvas.height }
    try { canvas.setPointerCapture(e.pointerId) } catch { /* abaikan */ }
    dragRef.current = { mode: hit(p, r), startX: p.x, startY: p.y, orig: { ...crop }, dw: canvas.width, dh: canvas.height }
  }

  const onMove = (e) => {
    const drag = dragRef.current
    if (!drag) return
    e.preventDefault()
    const p = toLocal(e)
    const dx = (p.x - drag.startX) / drag.dw
    const dy = (p.y - drag.startY) / drag.dh
    const o = drag.orig
    const minS = 0.05
    let n = { ...o }
    if (drag.mode === 'move') {
      n.x = Math.min(1 - o.w, Math.max(0, o.x + dx))
      n.y = Math.min(1 - o.h, Math.max(0, o.y + dy))
    } else if (drag.mode === 'new') {
      const ax = drag.startX / drag.dw, ay = drag.startY / drag.dh
      const bx = p.x / drag.dw, by = p.y / drag.dh
      n.x = Math.min(ax, bx); n.y = Math.min(ay, by)
      n.w = Math.abs(bx - ax); n.h = Math.abs(by - ay)
    } else if (drag.mode === 'nw') {
      const nx = Math.min(o.x + o.w - minS, Math.max(0, o.x + dx))
      const ny = Math.min(o.y + o.h - minS, Math.max(0, o.y + dy))
      n.w = o.w + (o.x - nx); n.h = o.h + (o.y - ny); n.x = nx; n.y = ny
    } else if (drag.mode === 'ne') {
      const ny = Math.min(o.y + o.h - minS, Math.max(0, o.y + dy))
      n.y = ny; n.h = o.h + (o.y - ny)
      n.w = Math.max(minS, Math.min(1 - o.x, o.w + dx))
    } else if (drag.mode === 'sw') {
      const nx = Math.min(o.x + o.w - minS, Math.max(0, o.x + dx))
      n.x = nx; n.w = o.w + (o.x - nx)
      n.h = Math.max(minS, Math.min(1 - o.y, o.h + dy))
    } else if (drag.mode === 'se') {
      n.w = Math.max(minS, Math.min(1 - o.x, o.w + dx))
      n.h = Math.max(minS, Math.min(1 - o.y, o.h + dy))
    }
    n.x = Math.min(0.95, Math.max(0, n.x))
    n.y = Math.min(0.95, Math.max(0, n.y))
    n.w = Math.min(1 - n.x, Math.max(0.03, n.w))
    n.h = Math.min(1 - n.y, Math.max(0.03, n.h))
    setCrop(n)
  }

  const onUp = () => { dragRef.current = null }

  if (!src) return null
  return (
    <div className="modal" onClick={onCancel} style={{ zIndex: 60 }} role="dialog" aria-modal="true" aria-label="Atur area tanda tangan">
      <div className="box" onClick={(e) => e.stopPropagation()}>
        <div className="mhead">
          <h3>Crop TTD</h3>
          <button className="iconbtn" onClick={onCancel} aria-label="Tutup">✕</button>
        </div>
        <p className="mut" style={{ marginTop: 0 }}>Seret kotak / titik sudut. Hasil enhance otomatis di bawah.</p>

        <div className="cropwrap">
          <canvas
            ref={editorRef}
            onPointerDown={onDown} onPointerMove={onMove}
            onPointerUp={onUp} onPointerCancel={onUp}
          />
        </div>
        <div className="toolbar" style={{ marginTop: 8 }}>
          <button onClick={() => setCrop({ x: 0.08, y: 0.3, w: 0.84, h: 0.4 })}>Reset</button>
          <button onClick={() => setCrop({ x: 0, y: 0, w: 1, h: 1 })}>Semua</button>
          <span style={{ flex: 1 }} />
          <label className="switch mini" title="Background kertas jadi transparan, tinta jadi hitam pekat">
            <input type="checkbox" checked={enhance} onChange={(e) => setEnhance(e.target.checked)} />
            <span>Enhance</span>
          </label>
        </div>
        {enhance && (
          <div className="throw">
            <span className="mut">Sensitivitas</span>
            <input type="range" min={100} max={220} value={threshold} onChange={(e) => setThreshold(Number(e.target.value))} />
            <b>{threshold}</b>
          </div>
        )}

        <div className="mut" style={{ margin: '8px 0 4px' }}>Hasil:</div>
        <div className="checker">
          <canvas ref={previewRef} />
        </div>

        <div className="toolbar" style={{ marginTop: 10 }}>
          <span style={{ flex: 1 }} />
          <button onClick={onCancel}>Batal</button>
          <button className="primary" disabled={!resultUrl} onClick={() => resultUrl && onApply(resultUrl)}>Gunakan</button>
        </div>
      </div>
    </div>
  )
}
