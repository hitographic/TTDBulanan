import { useEffect, useMemo, useState } from 'react'
import ReportPreview from './ReportPreview'
import { loadRead, saveRead } from '../lib/auth'

const shortName = (name) =>
  String(name || '').replace(/\.xlsx?$/i, '').trim() || name

const STEPS = ['Baca', 'Pilih', 'Isi', 'Selesai']

/**
 * Mode "semua sekaligus": satu TTD untuk seluruh slot di semua laporan.
 * Tanpa list — langsung satu kartu + popup tanda tangan.
 */
function AllAtOnce({ session, files, queue, sigOf, onRequestSign, setStage }) {
  const mine = sigOf(session.nameInFile)
  const on = Boolean(mine?.dataUrl)
  const repKey = queue[0]?.key
  return (
    <>
      <div className="onecard">
        <small className="mut">Satu tanda tangan untuk</small>
        <div className="baname">{queue.length} slot • {files.length} laporan</div>
        <div className="onesig">
          {on
            ? <img src={mine.dataUrl} alt="Tanda tanganku" />
            : <span className="badashed">Belum ada TTD — ketuk tombol di bawah</span>}
        </div>
        {repKey ? (
          <button className="primary" style={{ width: '100%', minHeight: 54 }} onClick={() => onRequestSign(repKey)}>
            {on ? '✏️ Ubah Tanda Tangan' : '✍️ Isi Tanda Tangan'}
          </button>
        ) : (
          <p className="mut">Tidak ada slot untuk akun ini.</p>
        )}
      </div>
      <div className="navrow">
        <button className="ghost" onClick={() => setStage('mode')}>← Kembali</button>
        <button className="primary" disabled={!on} onClick={() => setStage('done')}>
          {on ? 'Selesai →' : 'Isi dulu'}
        </button>
      </div>
    </>
  )
}

/**
 * Walkthrough user biasa: 1 Baca laporan -> 2 Pilih mode ->
 * 3 Isi TTD (satu per satu / semua sekaligus) -> 4 Selesai.
 * Tiap halaman ada tombol ← →. Admin tetap memakai dashboard penuh.
 */
export default function UserWizard({
  session, files, sigOf, isMine, effSigsFor, onRequestSign, onLogout,
}) {
  const [step, setStep] = useState('list') // list | mode | sign | done
  const [mode, setMode] = useState(null) // 'one' | 'all'
  const [read, setRead] = useState(() => loadRead(session.nik))
  const [previewId, setPreviewId] = useState(null)
  const [oneIdx, setOneIdx] = useState(0)
  const [confirmSkip, setConfirmSkip] = useState(false)

  useEffect(() => {
    saveRead(session.nik, read)
  }, [read, session.nik])

  const markRead = (f) => setRead((prev) => new Set(prev).add(f.name))
  const allRead = files.length > 0 && files.every((f) => read.has(f.name))
  const unread = files.filter((f) => !read.has(f.name)).length
  const previewFile = previewId ? (files.find((f) => f.id === previewId) ?? null) : null

  const fileState = (f) => {
    const m = (f.persons ?? []).filter((p) => isMine(p.name))
    const d = m.filter((p) => sigOf(p.name)?.dataUrl).length
    return { total: m.length, done: d, complete: m.length > 0 && d === m.length }
  }

  const months = useMemo(() => {
    const map = new Map()
    files.forEach((f) => {
      const k = f.month || 'Lainnya'
      if (!map.has(k)) map.set(k, [])
      map.get(k).push(f)
    })
    return [...map.entries()]
  }, [files])

  // Antrean slot milik user lintas semua file
  const queue = useMemo(() => {
    const out = []
    files.forEach((f) => {
      ;(f.persons ?? []).forEach((p) => {
        if (isMine(p.name)) {
          out.push({ fileId: f.id, fileName: f.name, month: f.month, key: p.key, name: p.name })
        }
      })
    })
    return out
  }, [files, session]) // eslint-disable-line react-hooks/exhaustive-deps

  const signedCount = queue.filter((q) => sigOf(q.name)?.dataUrl).length
  const allSigned = queue.length > 0 && signedCount === queue.length
  const unsigned = queue.filter((q) => !sigOf(q.name)?.dataUrl)
  const cur = unsigned.length ? unsigned[Math.min(oneIdx, unsigned.length - 1)] : null

  // Mode satu-per-satu: maju otomatis ke slot berikut setelah terisi
  useEffect(() => {
    if (step === 'sign' && mode === 'one' && cur && sigOf(cur.name)?.dataUrl) {
      const t = setTimeout(() => setOneIdx((i) => i + 1), 700)
      return () => clearTimeout(t)
    }
  }) // eslint-disable-line react-hooks/exhaustive-deps

  const goMode = () => {
    setMode(null)
    setStage('mode')
  }
  const setStage = (s) => {
    setPreviewId(null)
    setStep(s)
  }

  const stepIdx = { list: 1, mode: 2, sign: 3, done: 4 }[step]

  return (
    <div className="gwrap">
      <header className="ghead">
        <div className="guser">
          <span className="avatar">{(session.name || session.nik || '?').trim()[0]?.toUpperCase()}</span>
          <div>
            <b>Halo, {(session.name || session.nik || '').split(' ')[0]}! 👋</b>
            <small>{signedCount}/{queue.length} TTD{session.offline ? ' • offline' : ''}</small>
          </div>
        </div>
        <button className="glogout" onClick={onLogout}>Keluar</button>
      </header>

      <div className="steps" aria-label="Tahapan">
        {STEPS.map((l, i) => {
          const n = i + 1
          return (
            <span key={l} className={`step${stepIdx === n ? ' cur' : ''}${stepIdx > n ? ' ok' : ''}`}>
              <b>{stepIdx > n ? '★' : n}</b>{l}
            </span>
          )
        })}
      </div>

      {step === 'list' && (
        <>
          <div className="greet">
            <b>📖 Langkah 1 — Baca laporan</b>
            <span>Buka tiap laporan, baca, lalu tandai sudah dibaca.</span>
          </div>
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
                const rd = read.has(f.name)
                const st = fileState(f)
                return (
                  <div key={f.id} className="slotrow">
                    <span className="scicon" aria-hidden>📑</span>
                    <span className="sinfo">
                      <b>{shortName(f.name)}</b>
                      <small>{st.done}/{st.total} TTD</small>
                    </span>
                    <span className={`readpill${rd ? ' ok' : ''}`}>{rd ? 'Dibaca ✓' : 'Baru'}</span>
                    <button onClick={() => setPreviewId(f.id)} title="Baca laporan">👁</button>
                  </div>
                )
              })}
            </div>
          ))}
          {!!files.length && (
            <div className="navrow">
              <button
                className="primary"
                onClick={() => { allRead ? goMode() : setConfirmSkip(true) }}
              >
                {allRead ? 'Lanjut →' : `Lanjut (${files.length - unread}/${files.length} dibaca) →`}
              </button>
            </div>
          )}
        </>
      )}

      {confirmSkip && (
        <div className="modal" role="dialog" aria-modal="true" aria-label="Konfirmasi lewati bacaan">
          <div className="box" style={{ textAlign: 'center' }}>
            <div style={{ fontSize: 44 }} aria-hidden>📖</div>
            <h3 style={{ margin: '8px 0 4px' }}>Yakin tidak baca dulu?</h3>
            <p className="mut">
              Masih ada {unread} laporan belum dibaca. Tanda tangan berarti kamu menyetujui isinya.
            </p>
            <div className="toolbar center">
              <button onClick={() => setConfirmSkip(false)}>← Baca dulu</button>
              <button className="primary" onClick={() => { setConfirmSkip(false); goMode() }}>
                Ya, lanjut →
              </button>
            </div>
          </div>
        </div>
      )}

      {step === 'mode' && (
        <>
          <div className="greet">
            <b>✍️ Langkah 2 — Pilih cara mengisi</b>
            <span>Sekaligus atau satu per satu, bebas.</span>
          </div>
          <button className={`modecard${mode === 'one' ? ' sel' : ''}`} onClick={() => setMode('one')}>
            <span className="scicon" aria-hidden>🚶</span>
            <span className="sctext">
              <b>Satu per satu</b>
              <small>Dipandu slot per slot, otomatis lanjut.</small>
            </span>
          </button>
          <button className={`modecard${mode === 'all' ? ' sel' : ''}`} onClick={() => setMode('all')}>
            <span className="scicon" aria-hidden>⚡</span>
            <span className="sctext">
              <b>Semua sekaligus</b>
              <small>Satu TTD berlaku untuk semua laporan.</small>
            </span>
          </button>
          <div className="navrow">
            <button className="ghost" onClick={() => setStage('list')}>← Kembali</button>
            <button
              className="primary"
              disabled={!mode}
              onClick={() => { setOneIdx(0); setStage('sign') }}
            >
              Lanjut →
            </button>
          </div>
        </>
      )}

      {step === 'sign' && (
        <>
          <div className="greet">
            <b>✍️ Langkah 3 — {mode === 'one' ? 'Isi satu per satu' : 'Isi semua'}</b>
            <span>{signedCount}/{queue.length} terisi • tersimpan otomatis ke cloud.</span>
          </div>

          {mode === 'one' && (
            <>
              {!cur ? (
                <div className="onecard">
                  <div className="doneart" aria-hidden>✅</div>
                  <b>Semua sudah terisi!</b>
                  <div className="navrow">
                    <button className="primary" onClick={() => setStage('done')}>Lihat hasil →</button>
                  </div>
                </div>
              ) : (
                <div className="onecard">
                  <small className="mut">Slot {oneIdx + 1} dari {unsigned.length} tersisa</small>
                  <div className="baname">{cur.name}</div>
                  <small className="mut">{shortName(cur.fileName)}{cur.month ? ` • ${cur.month}` : ''}</small>
                  <div className="onesig">
                    {sigOf(cur.name)?.dataUrl
                      ? <img src={sigOf(cur.name).dataUrl} alt={cur.name} />
                      : <span className="badashed">Belum ada TTD — ketuk tombol di bawah</span>}
                  </div>
                  <button className="primary" style={{ width: '100%', minHeight: 54 }} onClick={() => onRequestSign(cur.key)}>
                    ✍️ Isi Tanda Tangan
                  </button>
                  <div className="navrow">
                    <button className="ghost" disabled={oneIdx === 0} onClick={() => setOneIdx((i) => Math.max(0, i - 1))}>
                      ← Sebelumnya
                    </button>
                    <button className="ghost" onClick={() => setPreviewId(cur.fileId)}>👁 Dokumen</button>
                  </div>
                </div>
              )}
              <div className="navrow">
                <button className="ghost" onClick={() => setStage('mode')}>← Pilih mode</button>
              </div>
            </>
          )}

          {mode === 'all' && <AllAtOnce
            session={session}
            files={files}
            queue={queue}
            sigOf={sigOf}
            onRequestSign={onRequestSign}
            setStage={setStage}
          />}
        </>
      )}

      {step === 'done' && (
        <div className="donecard">
          <div className="doneart" aria-hidden>✅</div>
          <h2>Semua selesai!</h2>
          <p className="mut">{signedCount}/{queue.length} tanda tangan • tersimpan di cloud</p>
          <div className="toolbar center">
            <button className="primary" onClick={() => { setMode(null); setOneIdx(0); setStage('list') }}>
              ＋ Daftar laporan
            </button>
            <button onClick={onLogout}>Keluar</button>
          </div>
        </div>
      )}

      {previewFile && (
        <div className="gamemodal" role="dialog" aria-label="Preview dokumen">
          <div className="gamedoc">
            <div className="mhead">
              <h3>{shortName(previewFile.name)}</h3>
              <button className="iconbtn" onClick={() => setPreviewId(null)} aria-label="Tutup">✕</button>
            </div>
            <ReportPreview
              sheet={previewFile.sheet}
              blocks={previewFile.blocks}
              signatures={effSigsFor(previewFile)}
              fileUid={previewFile.id}
              spotMissing
              fit
              onEditSig={(key) => onRequestSign(key)}
            />
            <div className="toolbar center">
              <button className="primary" onClick={() => { markRead(previewFile); setPreviewId(null) }}>
                Sudah dibaca ✓
              </button>
              <button onClick={() => setPreviewId(null)}>Tutup</button>
            </div>
          </div>
        </div>
      )}
    </div>
  )
}
