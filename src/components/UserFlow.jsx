import { useMemo, useState } from 'react'
import ReportPreview from './ReportPreview'

const shortName = (name) =>
  String(name || '').replace(/\.xlsx?$/i, '').trim() || name

const Confetti = () => (
  <div className="confetti" aria-hidden>
    {Array.from({ length: 24 }, (_, i) => (
      <span key={i} style={{ '--d': `${(i % 12) * 0.15}s`, '--x': `${(i * 37 + 5) % 100}%` }} />
    ))}
  </div>
)

/**
 * Alur game untuk user biasa: 1 Pilih laporan -> 2 Sebelum/Sesudah TTD ->
 * 3 Simpan (finale + confetti). Admin tetap memakai dashboard penuh.
 */
export default function UserFlow({
  session, files, sigOf, isMine, effSigsFor,
  onRequestSign, onFinalSave, onLogout,
}) {
  const [selId, setSelId] = useState(null)
  const [stage, setStage] = useState('pick') // pick | sign | done
  const [showDoc, setShowDoc] = useState(false)
  const [saveState, setSaveState] = useState(null) // null | 'saving' | 'ok' | 'fail'

  const sel = files.find((f) => f.id === selId) ?? null
  const mySlots = useMemo(
    () => (sel?.persons ?? []).filter((p) => isMine(p.name)),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [sel]
  )
  const doneCount = mySlots.filter((p) => sigOf(p.name)?.dataUrl).length
  const allDone = mySlots.length > 0 && doneCount === mySlots.length

  const fileState = (f) => {
    const m = (f.persons ?? []).filter((p) => isMine(p.name))
    const d = m.filter((p) => sigOf(p.name)?.dataUrl).length
    return { total: m.length, done: d, complete: m.length > 0 && d === m.length }
  }
  const completed = files.filter((f) => fileState(f).complete).length

  const months = useMemo(() => {
    const map = new Map()
    files.forEach((f) => {
      const k = f.month || 'Lainnya'
      if (!map.has(k)) map.set(k, [])
      map.get(k).push(f)
    })
    return [...map.entries()]
  }, [files])

  const pick = (f) => {
    setSelId(f.id)
    setStage('sign')
    setShowDoc(false)
    setSaveState(null)
  }
  const back = () => {
    setStage('pick')
    setSelId(null)
    setShowDoc(false)
    setSaveState(null)
  }

  const doSave = async () => {
    if (!sel || saveState === 'saving') return
    setSaveState('saving')
    const ok = await onFinalSave(sel)
    setSaveState(ok ? 'ok' : 'fail')
  }

  return (
    <div className="gwrap">
      <header className="ghead">
        <div className="guser">
          <span className="avatar">{(session.name || session.nik || '?').trim()[0]?.toUpperCase()}</span>
          <div>
            <b>Halo, {(session.name || session.nik || '').split(' ')[0]}! 👋</b>
            <small>{completed}/{files.length} laporan selesai {session.offline ? '• offline' : ''}</small>
          </div>
        </div>
        <button className="glogout" onClick={onLogout}>Keluar</button>
      </header>

      <div className="steps" aria-label="Tahapan">
        {['Pilih', 'Tanda Tangan', 'Simpan'].map((l, i) => {
          const n = i + 1
          const cur = (stage === 'pick' && n === 1) || (stage === 'sign' && n === 2) || (stage === 'done' && n === 3)
          const ok = (n === 1 && stage !== 'pick') || (n === 2 && stage === 'done')
          return <span key={l} className={`step${cur ? ' cur' : ''}${ok ? ' ok' : ''}`}><b>{ok ? '★' : n}</b>{l}</span>
        })}
      </div>

      {stage === 'pick' && (
        <>
          <div className="greet"><b>🎮 Misi TTD dimulai!</b><span>Pilih laporan untuk ditandatangani.</span></div>
          {!files.length && (
            <div className="card empty">
              <div className="emptyart">📄</div>
              <p className="mut">Belum ada laporan untukmu.</p>
            </div>
          )}
          {months.map(([m, list]) => (
            <div key={m}>
              <div className="monthhead">📁 {m}</div>
              {list.map((f) => {
                const st = fileState(f)
                return (
                  <button key={f.id} className="stagecard" onClick={() => pick(f)}>
                    <span className="scicon">📑</span>
                    <span className="sctext">
                      <b>{shortName(f.name)}</b>
                      <small>{st.done}/{st.total} tanda tangan</small>
                    </span>
                    <span className="stars" aria-hidden>
                      {'★'.repeat(st.done)}{'☆'.repeat(Math.max(0, st.total - st.done))}
                    </span>
                    <span className={`pill${st.complete ? ' ok' : st.done ? ' half' : ''}`}>
                      {st.complete ? 'Selesai ✅' : st.done ? 'Lanjut ▶' : 'Mulai 🚀'}
                    </span>
                  </button>
                )
              })}
            </div>
          ))}
        </>
      )}

      {stage === 'sign' && sel && (
        <>
          <button className="gback" onClick={back}>← Laporan</button>
          <div className="gtitle">
            <b>{shortName(sel.name)}</b>
            <small>{sel.month ? `${sel.month} • ` : ''}{doneCount}/{mySlots.length} terisi</small>
          </div>
          <div className="batabs">
            <span>SEBELUM ➜ SESUDAH</span>
            <button className="linkbtn" onClick={() => setShowDoc(true)}>👁 Preview dokumen</button>
          </div>
          {mySlots.map((p) => {
            const s = sigOf(p.name)
            const on = Boolean(s?.dataUrl)
            return (
              <div key={p.key} className={`bacard${on ? ' done' : ''}`}>
                <div className="ba">
                  <div className="babefore">
                    <small>SEBELUM</small>
                    <span className="baname">{p.name}</span>
                    <span className="badashed">— kosong —</span>
                  </div>
                  <div className="baarrow" aria-hidden>➜</div>
                  <div className="baafter">
                    <small>SESUDAH</small>
                    {on ? <img src={s.dataUrl} alt={p.name} /> : <span className="badashed">?</span>}
                  </div>
                </div>
                <button className={on ? '' : 'primary'} onClick={() => onRequestSign(p.key)}>
                  {on ? '✏️ Ubah' : '✍️ Isi Tanda Tangan'}
                </button>
              </div>
            )
          })}
          <button className="bignext" disabled={!allDone} onClick={() => { setSaveState(null); setStage('done') }}>
            {allDone ? 'Lanjut ke Simpan 🎉' : `Isi dulu (${doneCount}/${mySlots.length})`}
          </button>
          {showDoc && (
            <div className="gamemodal" role="dialog" aria-label="Preview dokumen">
              <div className="gamedoc">
                <div className="mhead">
                  <h3>{shortName(sel.name)}</h3>
                  <button className="iconbtn" onClick={() => setShowDoc(false)} aria-label="Tutup">✕</button>
                </div>
                <ReportPreview
                  sheet={sel.sheet}
                  blocks={sel.blocks}
                  signatures={effSigsFor(sel)}
                  fileUid={sel.id}
                  spotMissing
                  fit
                  onEditSig={(key) => onRequestSign(key)}
                />
              </div>
            </div>
          )}
        </>
      )}

      {stage === 'done' && sel && (
        <>
          <button className="gback" onClick={() => setStage('sign')}>← Kembali</button>
          <div className="donecard">
            {saveState === 'ok' && <Confetti />}
            <div className="doneart" aria-hidden>🏆</div>
            <h2>Laporan siap disimpan!</h2>
            <p className="mut">{shortName(sel.name)}{sel.month ? ` • ${sel.month}` : ''}</p>
            <div className="donerow">
              {mySlots.map((p) => {
                const s = sigOf(p.name)
                return s?.dataUrl
                  ? <img key={p.key} src={s.dataUrl} alt={p.name} title={p.name} />
                  : null
              })}
            </div>
            {saveState === 'fail' && (
              <div className="formerr" role="alert">Gagal menyimpan — periksa internet lalu coba lagi.</div>
            )}
            {saveState === 'ok' ? (
              <>
                <div className="pill ok big">✅ Tersimpan di cloud!</div>
                <div className="toolbar center">
                  <button className="primary" onClick={back}>＋ Laporan lain</button>
                  <button onClick={onLogout}>Keluar</button>
                </div>
              </>
            ) : (
              <button className="bigsave" disabled={saveState === 'saving'} onClick={doSave}>
                {saveState === 'saving' ? 'Menyimpan…' : '💾 SIMPAN'}
              </button>
            )}
          </div>
        </>
      )}
    </div>
  )
}
