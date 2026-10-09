import { useEffect, useRef, useState } from 'react'
import SigCropModal from './SigCropModal'
import { findInkBounds, padBounds } from '../lib/sigImage'

const COLORS = [
  ['#00008B', 'Dongker'],
  ['#111111', 'Hitam'],
  ['#1E90FF', 'Biru'],
  ['#7C3AED', 'Ungu'],
]

/** Potong canvas ke bounding-box goresan (hemat tempat, TTD lebih besar di export). */
function cropToInk(cv) {
  try {
    const ctx = cv.getContext('2d')
    const id = ctx.getImageData(0, 0, cv.width, cv.height)
    const alpha = new Uint8Array(cv.width * cv.height)
    for (let i = 0, j = 0; i < id.data.length; i += 4, j++) alpha[j] = id.data[i + 3]
    const box = padBounds(findInkBounds(alpha, cv.width, cv.height), cv.width, cv.height, 12)
    if (!box) return cv.toDataURL('image/png')
    const tmp = document.createElement('canvas')
    tmp.width = box.w
    tmp.height = box.h
    tmp.getContext('2d').drawImage(cv, box.x, box.y, box.w, box.h, 0, 0, box.w, box.h)
    return tmp.toDataURL('image/png')
  } catch {
    return cv.toDataURL('image/png')
  }
}

/** Canvas TTD: gambar (mouse/sentuh/stylus) + ketik nama + upload (crop + enhance). Ramah HP. */
export default function SignaturePad({ initial, enabled = true, personName, onSave, onSaveNext, hasNext, nextName, onToggleEnabled, onDelete, onClose }) {
  const ref = useRef(null)
  const drawing = useRef(false)
  const last = useRef(null)
  const [tab, setTab] = useState('gambar')
  const [typed, setTyped] = useState(personName || '')
  const [font, setFont] = useState("'Brush Script MT','Segoe Script',cursive")
  const [color, setColor] = useState('#00008B')
  const [thickness, setThickness] = useState(2)
  const [pendingSrc, setPendingSrc] = useState(null)
  const [err, setErr] = useState('')

  // Muat TTD lama ke canvas saat tab gambar dibuka
  useEffect(() => {
    if (tab === 'gambar' && initial) {
      const t = setTimeout(() => {
        const cv = ref.current
        if (!cv) return
        const img = new Image()
        img.onload = () => {
          const ctx = cv.getContext('2d')
          ctx.clearRect(0, 0, cv.width, cv.height)
          const s = Math.min(cv.width / img.width, cv.height / img.height, 1)
          const w = img.width * s, h = img.height * s
          ctx.drawImage(img, (cv.width - w) / 2, (cv.height - h) / 2, w, h)
        }
        img.src = initial
      }, 30)
      return () => clearTimeout(t)
    }
  }, [tab, initial])

  // Kunci scroll halaman saat modal terbuka (penting di HP)
  useEffect(() => {
    const prev = document.body.style.overflow
    document.body.style.overflow = 'hidden'
    const onKey = (e) => { if (e.key === 'Escape') onClose() }
    window.addEventListener('keydown', onKey)
    return () => {
      document.body.style.overflow = prev
      window.removeEventListener('keydown', onKey)
    }
  }, [onClose])

  const pos = (clientX, clientY) => {
    const cv = ref.current
    const r = cv.getBoundingClientRect()
    return {
      x: ((clientX - r.left) / r.width) * cv.width,
      y: ((clientY - r.top) / r.height) * cv.height,
    }
  }

  const onPointerDown = (e) => {
    e.preventDefault()
    drawing.current = true
    last.current = pos(e.clientX, e.clientY)
    try { ref.current.setPointerCapture(e.pointerId) } catch { /* abaikan */ }
  }
  const onPointerMove = (e) => {
    if (!drawing.current) return
    e.preventDefault()
    const p = pos(e.clientX, e.clientY)
    const ctx = ref.current.getContext('2d')
    const scale = ref.current.width / Math.max(ref.current.getBoundingClientRect().width, 1)
    ctx.strokeStyle = color
    ctx.lineWidth = Math.max(1.5, thickness * scale)
    ctx.lineCap = 'round'
    ctx.lineJoin = 'round'
    ctx.beginPath()
    ctx.moveTo(last.current.x, last.current.y)
    ctx.lineTo(p.x, p.y)
    ctx.stroke()
    last.current = p
  }
  const onPointerUp = () => { drawing.current = false }

  const clear = () => {
    const cv = ref.current
    if (cv) cv.getContext('2d').clearRect(0, 0, cv.width, cv.height)
  }

  const canvasEmpty = (cv) => {
    const d = cv.getContext('2d').getImageData(0, 0, cv.width, cv.height).data
    for (let i = 3; i < d.length; i += 40) if (d[i] > 10) return false
    return true
  }

  const saveCanvas = (next) => {
    const cv = ref.current
    if (!cv) return
    if (canvasEmpty(cv)) { setErr('Canvas masih kosong — gambar dulu atau pakai tab Ketik/Upload.'); return }
    setErr('')
    ;(next && hasNext ? onSaveNext : onSave)(cropToInk(cv))
  }

  const saveTyped = (next) => {
    if (!typed.trim()) { setErr('Ketik nama dulu.'); return }
    setErr('')
    const cv = document.createElement('canvas')
    cv.width = 900; cv.height = 600
    const ctx = cv.getContext('2d')
    ctx.fillStyle = color; ctx.textAlign = 'center'; ctx.textBaseline = 'middle'
    ctx.font = `110px ${font}`
    ctx.fillText(typed.trim(), 450, 310)
    ;(next && hasNext ? onSaveNext : onSave)(cropToInk(cv))
  }

  const onFile = (e) => {
    const f = e.target.files?.[0]
    if (!f) return
    const r = new FileReader()
    r.onload = (ev) => { setPendingSrc(ev.target.result); setErr('') }
    r.readAsDataURL(f)
    e.target.value = ''
  }

  const pickTab = (k) => { setTab(k); setErr('') }

  const swatches = (compact) => (
    <span className={`swatches${compact ? ' mini' : ''}`}>
      {COLORS.map(([c, label]) => (
        <button
          key={c}
          type="button"
          className={`sw${color === c ? ' on' : ''}`}
          style={{ background: c }}
          title={label}
          aria-label={`Warna ${label}`}
          onClick={() => setColor(c)}
        />
      ))}
      <label className={`sw custom${!COLORS.some(([c]) => c === color) ? ' on' : ''}`} title="Warna lain">
        <input type="color" value={color} onChange={(e) => setColor(e.target.value)} aria-label="Warna lain" />
        <span aria-hidden>+</span>
      </label>
    </span>
  )

  return (
    <div className="modal" onClick={onClose} role="dialog" aria-modal="true" aria-label={`Tanda tangan ${personName}`}>
      <div className="box" onClick={(e) => e.stopPropagation()}>
        <div className="mhead">
          <h3>{personName}</h3>
          <button className="iconbtn" onClick={onClose} aria-label="Tutup">✕</button>
        </div>
        {hasNext && nextName && <p className="mut" style={{ marginTop: 0 }}>Lanjut: <b>{nextName}</b></p>}
        {initial && (
          <div className="mrow">
            <label className="switch mini">
              <input type="checkbox" checked={enabled} onChange={onToggleEnabled} />
              <span>Pakai di export</span>
            </label>
            <button className="linkbtn danger-text" onClick={onDelete}>Hapus</button>
          </div>
        )}
        <div className="tabs">
          {[['gambar', 'Gambar'], ['ketik', 'Ketik'], ['upload', 'Upload']].map(([k, l]) => (
            <button key={k} className={tab === k ? 'on' : ''} onClick={() => pickTab(k)} style={{ flex: '1 1 auto' }}>{l}</button>
          ))}
        </div>
        {err && <div className="formerr" role="alert">{err}</div>}
        {tab === 'gambar' && (
          <>
            <canvas
              ref={ref} width={720} height={540} className="pad"
              style={{ touchAction: 'none' }}
              onPointerDown={onPointerDown} onPointerMove={onPointerMove}
              onPointerUp={onPointerUp} onPointerCancel={onPointerUp} onPointerLeave={onPointerUp}
            />
            <div className="drawtools">
              {swatches()}
              <span className="vdiv" />
              <input
                type="range" min={1} max={6} step={1} value={thickness}
                onChange={(e) => setThickness(Number(e.target.value))}
                title={`Ketebalan ${thickness}`} aria-label="Ketebalan garis"
              />
              <span style={{ flex: 1 }} />
              <button onClick={clear} title="Hapus gambar">Hapus</button>
            </div>
            <div className="toolbar" style={{ marginTop: 8 }}>
              <button className="primary" onClick={() => saveCanvas(false)}>Simpan</button>
              {hasNext && <button className="ok" onClick={() => saveCanvas(true)}>Lanjut →</button>}
            </div>
          </>
        )}
        {tab === 'ketik' && (
          <>
            <input type="text" enterKeyHint="done" autoComplete="off" value={typed} onChange={(e) => setTyped(e.target.value)} placeholder="cth. Afida Ayu A." />
            <div className="fontrow">
              {[["'Brush Script MT','Segoe Script',cursive", 'Script'], ['Georgia,serif', 'Serif'], ['Arial,sans-serif', 'Sans']].map(([f, l]) => (
                <button key={l} className={font === f ? 'primary' : ''} onClick={() => setFont(f)}>{l}</button>
              ))}
            </div>
            <div className="drawtools" style={{ marginBottom: 8 }}>
              {swatches(true)}
            </div>
            <div className="typepreview" style={{ fontFamily: font, color }}>
              {typed || '—'}
            </div>
            <div className="toolbar" style={{ marginTop: 8 }}>
              <button className="primary" onClick={() => saveTyped(false)}>Simpan</button>
              {hasNext && <button className="ok" onClick={() => saveTyped(true)}>Lanjut →</button>}
            </div>
          </>
        )}
        {tab === 'upload' && (
          <>
            <label className="btn" style={{ display: 'flex', minHeight: 52, alignItems: 'center', justifyContent: 'center', fontWeight: 700 }}>
              📷 Pilih / foto TTD
              <input type="file" accept="image/*" hidden onChange={onFile} />
            </label>
            <p className="mut">Potong area lalu enhance otomatis (background jadi transparan).</p>
          </>
        )}
      </div>

      {pendingSrc && (
        <SigCropModal
          src={pendingSrc}
          onCancel={() => setPendingSrc(null)}
          onApply={(url) => { setPendingSrc(null); onSave(url) }}
        />
      )}
    </div>
  )
}
