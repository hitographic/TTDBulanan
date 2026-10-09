import { useEffect, useMemo, useRef, useState } from 'react'
import './styles.css'
import { parseWorkbook, normKey } from './lib/excelParser'
import { buildSheetModel, normalizeWorkbookBuffer } from './lib/sheetModel'
import { exportSignedExcel, buildSignedWorkbook } from './lib/excelExport'
import { exportAllZip } from './lib/zipExport'
import { exportPreviewPdf } from './lib/pdfExport'
import {
  loadCloud, saveCloud, cloudPing, cloudPull, cloudFiles, cloudGet,
  cloudPush, cloudSave, mergeSignatures, bufToB64, b64ToBuf,
} from './lib/cloud'
import { browserMeasure, geomForSide } from './lib/textGeom'
import SignaturePad from './components/SignaturePad'
import ReportPreview from './components/ReportPreview'

let fileSeq = 0
let toastSeq = 0
const SIG_STORE_KEY = 'lapbul-signatures-v1'
const SAMPLE_NAME = 'contoh-laporan-bulanan.xlsx'

const shortName = (name) =>
  String(name || '').replace(/\.xlsx?$/i, '').trim() || name

const loadSigs = () => {
  try {
    const raw = localStorage.getItem(SIG_STORE_KEY)
    if (!raw) return {}
    const o = JSON.parse(raw)
    return o && typeof o === 'object' ? o : {}
  } catch {
    return {}
  }
}

const withGeom = (parsed) => {
  try {
    const measure = browserMeasure()
    parsed.blocks.forEach((b) => {
      ;(b.slots || []).forEach((s) => {
        s.geom = geomForSide(s.raw, s.names, measure)
      })
    })
  } catch { /* fallback: tengah slot */ }
  return parsed
}

export default function App() {
  const [files, setFiles] = useState([]) // [{id,name,buffer,sheetName,aoa,blocks,persons}]
  const [activeId, setActiveId] = useState(null)
  const [signatures, setSignatures] = useState(loadSigs) // key -> {dataUrl, enabled}, tersimpan otomatis
  const [editing, setEditing] = useState(null)
  const [q, setQ] = useState('')
  const [filter, setFilter] = useState('semua') // semua | belum | sudah
  const [view, setView] = useState('laporan') // 'laporan' | 'ringkas'
  const [previewFit, setPreviewFit] = useState(true) // true = pas lebar, false = ukuran asli Excel
  const [busy, setBusy] = useState('')
  const [mtab, setMtab] = useState('preview') // HP: file | nama | preview
  const [toasts, setToasts] = useState([])
  const [drag, setDrag] = useState(false)
  const [confirmDel, setConfirmDel] = useState(null)
  const [spot, setSpot] = useState(true) // sorot slot belum TTD di preview
  const [cloud, setCloud] = useState(loadCloud) // {url, autoSync} backend Apps Script
  const [cloudUrl, setCloudUrl] = useState(cloud.url) // draft input URL
  const [cloudOk, setCloudOk] = useState(null) // null=belum dites, true/false
  const [showCloud, setShowCloud] = useState(false) // panel pengaturan cloud
  const [driveFiles, setDriveFiles] = useState([]) // [{id,name,size,modified}]
  const delTimer = useRef(null)

  const notify = (msg, kind = 'info') => {
    const id = `${Date.now()}-${toastSeq++}`
    setToasts((t) => [...t.slice(-2), { id, msg, kind }])
    setTimeout(() => setToasts((t) => t.filter((x) => x.id !== id)), 3400)
  }

  // TTD tersimpan otomatis di browser — aman di-refresh
  useEffect(() => {
    try {
      localStorage.setItem(SIG_STORE_KEY, JSON.stringify(signatures))
    } catch { /* kuota penuh / mode privat: abaikan */ }
  }, [signatures])

  // Pengaturan cloud tersimpan otomatis; tes koneksi diam-diam bila ada URL
  useEffect(() => {
    try {
      saveCloud(cloud)
    } catch { /* abaikan */ }
  }, [cloud])

  useEffect(() => {
    const c = loadCloud()
    if (c.url) {
      cloudPing(c.url).then(() => setCloudOk(true)).catch(() => setCloudOk(false))
    }
  }, []) // eslint-disable-line

  useEffect(() => () => clearTimeout(delTimer.current), [])

  async function addOneBatch(items) {
    const next = []
    const failed = []
    setBusy('Membaca file…')
    try {
      for (const { name, buf } of items) {
        try {
          // .xls lawas dikonversi dulu ke .xlsx agar preview & export jalan
          const xbuf = normalizeWorkbookBuffer(buf.slice(0))
          const parsed = withGeom(parseWorkbook(xbuf.slice(0)))
          if (!parsed.blocks.length) {
            failed.push(`${name} (blok TTD Dibuat/Diketahui/Disetujui tak ditemukan)`)
            continue
          }
          // Model setia-1:1 untuk preview semirip Excel (style, merge, ukuran)
          let sheet = null
          try {
            sheet = await buildSheetModel(xbuf.slice(0))
          } catch (e) {
            console.warn('model preview gagal, pakai ringkas:', name, e)
          }
          next.push({
            id: `${Date.now()}-${fileSeq++}`,
            name,
            buffer: xbuf.slice(0),
            sheetName: parsed.sheetName,
            aoa: parsed.aoa,
            blocks: parsed.blocks,
            persons: parsed.persons,
            sheet,
          })
        } catch (e) {
          console.error(name, e)
          failed.push(`${name} (${e.message})`)
        }
      }
    } finally {
      setBusy('')
    }
    failed.forEach((m) => notify(`Gagal membaca ${m}`, 'err'))
    if (!next.length) return
    setFiles((prev) => {
      const map = new Map(prev.map((f) => [f.name, f]))
      next.forEach((f) => map.set(f.name, f))
      const arr = [...map.values()]
      if (!activeId) setActiveId(arr[0]?.id ?? null)
      else if (!arr.find((f) => f.id === activeId)) setActiveId(next[0].id)
      return arr
    })
    setActiveId((cur) => cur ?? next[0]?.id ?? null)
    setMtab((cur) => (cur === 'file' ? 'nama' : cur))
    const pages = next.reduce((n, f) => n + f.blocks.length, 0)
    notify(`${next.length} file dimuat • ${pages} halaman. Lanjut isi TTD.`, 'ok')
  }

  // Contoh bawaan bila belum ada file
  useEffect(() => {
    fetch(SAMPLE_NAME)
      .then((r) => (r.ok ? r.arrayBuffer() : Promise.reject()))
      .then((buf) => addOneBatch([{ name: SAMPLE_NAME, buf }]))
      .catch(() => {})
  }, []) // eslint-disable-line

  const readInputFiles = async (fileList) => {
    const list = [...(fileList || [])]
    if (!list.length) return
    setBusy('Membaca file…')
    try {
      const items = []
      for (const f of list) items.push({ name: f.name, buf: await f.arrayBuffer() })
      addOneBatch(items)
    } finally {
      setBusy('')
    }
  }

  const onUpload = async (e) => {
    await readInputFiles(e.target.files)
    e.target.value = ''
  }

  const loadSample = async () => {
    setBusy('Memuat contoh…')
    try {
      const r = await fetch(SAMPLE_NAME)
      if (!r.ok) throw new Error('contoh tak tersedia')
      const buf = await r.arrayBuffer()
      addOneBatch([{ name: SAMPLE_NAME, buf }])
    } catch {
      notify('File contoh tak tersedia.', 'err')
    } finally {
      setBusy('')
    }
  }

  const doRemove = (id) => {
    setFiles((prev) => {
      const arr = prev.filter((f) => f.id !== id)
      if (id === activeId) setActiveId(arr[0]?.id ?? null)
      return arr
    })
    setConfirmDel(null)
    notify('File dihapus dari daftar.', 'info')
  }

  const askRemove = (id) => {
    if (confirmDel === id) return doRemove(id)
    setConfirmDel(id)
    clearTimeout(delTimer.current)
    delTimer.current = setTimeout(() => setConfirmDel(null), 2600)
  }

  const askClearAll = () => {
    if (confirmDel === '__all__') {
      setFiles([])
      setActiveId(null)
      setConfirmDel(null)
      notify('Semua file dihapus.', 'info')
      return
    }
    setConfirmDel('__all__')
    clearTimeout(delTimer.current)
    delTimer.current = setTimeout(() => setConfirmDel(null), 2600)
  }

  // Gabungan nama lintas SEMUA file — 1 TTD berlaku ke semua file sekaligus
  const persons = useMemo(() => {
    const map = new Map()
    files.forEach((f) => {
      f.persons.forEach((p) => {
        if (!map.has(p.key)) map.set(p.key, { key: p.key, name: p.name, count: 0, files: new Set(), occ: [] })
        const g = map.get(p.key)
        if (g.name.length < p.name.length) g.name = p.name // ejaan paling lengkap
        g.count += p.count
        g.files.add(f.name)
        p.occ.forEach((o) => g.occ.push({ ...o, file: shortName(f.name), fileId: f.id }))
      })
    })
    return [...map.values()].sort((a, b) => b.count - a.count || a.name.localeCompare(b.name))
  }, [files])

  const activeFile = useMemo(
    () => files.find((f) => f.id === activeId) ?? files[0] ?? null,
    [files, activeId]
  )

  const isSigned = (key) => Boolean(signatures[key]?.dataUrl && signatures[key]?.enabled !== false)

  const filtered = useMemo(() => {
    const s = q.trim().toLowerCase()
    let arr = persons
    if (filter === 'belum') arr = arr.filter((p) => !signatures[p.key]?.dataUrl)
    else if (filter === 'sudah') arr = arr.filter((p) => signatures[p.key]?.dataUrl)
    if (s) arr = arr.filter((p) => p.name.toLowerCase().includes(s))
    return arr
  }, [persons, q, filter, signatures])

  const signedCount = persons.filter((p) => signatures[p.key]?.dataUrl).length
  const unsignedList = useMemo(() => persons.filter((p) => !signatures[p.key]?.dataUrl), [persons, signatures])
  const nextUp = unsignedList[0] ?? null
  const totalBlocks = files.reduce((n, f) => n + f.blocks.length, 0)
  const pct = persons.length ? Math.round((signedCount / persons.length) * 100) : 0
  const remaining = persons.length - signedCount

  const setSig = (key, dataUrl, ts = Date.now()) =>
    setSignatures((m) => ({ ...m, [key]: { dataUrl, enabled: true, updatedAt: ts } }))

  // Kirim 1 TTD ke cloud di background (tak memblokir UI)
  const pushSig = (key, name, dataUrl, ts) => {
    if (!cloud.url || !cloud.autoSync) return
    cloudPush(cloud.url, { key, name, dataUrl, updatedAt: ts }).catch((e) => {
      console.warn('auto-sync gagal:', e.message)
    })
  }

  const saveAndNext = (url) => {
    if (!editing) return
    const ts = Date.now()
    const nm = persons.find((p) => p.key === editing)?.name ?? editing
    setSig(editing, url, ts)
    pushSig(editing, nm, url, ts)
    const idx = unsignedList.findIndex((p) => p.key === editing)
    const next = unsignedList[idx + 1] ?? unsignedList.find((p) => p.key !== editing) ?? null
    if (next) {
      setEditing(next.key)
      notify(`TTD ${persons.find((p) => p.key === editing)?.name ?? ''} tersimpan. Lanjut: ${next.name}.`, 'ok')
    } else {
      setEditing(null)
      notify('Semua TTD lengkap. Siap export. 🎉', 'ok')
      setMtab('preview')
    }
  }

  const toggleEnabled = (key) =>
    setSignatures((m) => ({
      ...m,
      [key]: { dataUrl: m[key]?.dataUrl, enabled: m[key]?.enabled === false ? true : false },
    }))

  const clearSig = (key) => {
    setSignatures((m) => { const n = { ...m }; delete n[key]; return n })
    notify('TTD dihapus.', 'info')
  }

  const slotNames = (f) => {
    const out = []
    f.blocks.forEach((b) => {
      ;(b.slots || []).forEach((s) => {
        s.names.forEach((nm) => out.push({ blockId: b.id, name: nm }))
      })
    })
    return out
  }

  const fileSignedInfo = (f) => {
    let hit = 0, need = 0
    const seen = new Set()
    slotNames(f).forEach(({ blockId, name }) => {
      need += 1
      const k = normKey(name)
      if (isSigned(k) && !seen.has(`${blockId}-${k}`)) {
        seen.add(`${blockId}-${k}`)
        hit += 1
      }
    })
    return { hit, need, done: need > 0 && hit >= need }
  }

  const doExportOne = async (f) => {
    if (!f) return notify('Pilih file dulu.', 'err')
    if (signedCount === 0) return notify('Isi minimal 1 tanda tangan dulu.', 'err')
    try {
      await exportSignedExcel({
        originalBuffer: f.buffer,
        blocks: f.blocks,
        signatures,
        outName: f.name.replace(/\.xlsx?$/i, '') + '-signed.xlsx',
      })
      notify(`${shortName(f.name)} terdownload.`, 'ok')
    } catch (e) {
      console.error(e)
      notify('Gagal export: ' + e.message, 'err')
    }
  }

  const doExportZip = async () => {
    if (!files.length) return notify('Upload file Excel dulu.', 'err')
    if (signedCount === 0) return notify('Isi minimal 1 tanda tangan dulu.', 'err')
    setBusy('Menyiapkan ZIP…')
    try {
      await exportAllZip({
        files,
        signatures,
        outName: `Laporan-Bulanan-signed-${files.length}-file.zip`,
        onProgress: (a, b, name) => {
          if (b === 100) setBusy(`Mengompres ZIP… ${Math.round(a)}%`)
          else setBusy(`Menyiapkan ${a}/${b}… ${shortName(name)}`)
        },
      })
      notify(
        remaining === 0
          ? `ZIP ${files.length} file terdownload. Selesai 🎉`
          : `ZIP terdownload. Catatan: ${remaining} nama belum TTD.`,
        remaining === 0 ? 'ok' : 'info'
      )
    } catch (e) {
      console.error(e)
      notify('Gagal membuat ZIP: ' + e.message, 'err')
    } finally {
      setBusy('')
    }
  }

  const doExportPdf = async () => {
    if (!activeFile) return notify('Pilih file dulu.', 'err')
    if (signedCount === 0) return notify('Isi minimal 1 tanda tangan dulu.', 'err')
    try {
      await exportPreviewPdf('ttdPreview', activeFile.name.replace(/\.xlsx?$/i, '') + '-signed.pdf')
      notify('PDF preview terdownload.', 'ok')
    } catch (e) {
      console.error(e)
      notify('Gagal export PDF: ' + e.message, 'err')
    }
  }

  /* ===== Cloud: Google Sheets (database TTD) + Drive (file) ===== */
  const saveCloudUrl = async () => {
    const c = { url: cloudUrl.trim(), autoSync: cloud.autoSync }
    setCloud(c)
    if (!c.url) {
      setCloudOk(null)
      notify('Cloud dimatikan — mode lokal.', 'info')
      return
    }
    setBusy('Menghubungi cloud…')
    try {
      await cloudPing(c.url)
      setCloudOk(true)
      notify('Cloud tersambung. ✅', 'ok')
    } catch (e) {
      setCloudOk(false)
      notify(e.message, 'err')
    } finally {
      setBusy('')
    }
  }

  const doPull = async () => {
    if (!cloud.url) return notify('Isi URL Apps Script dulu.', 'err')
    setBusy('Mengambil TTD dari cloud…')
    try {
      const j = await cloudPull(cloud.url)
      const { merged, added, updated } = mergeSignatures(signatures, j.signatures)
      setSignatures(merged)
      setCloudOk(true)
      notify(
        added || updated
          ? `Sinkron: +${added} baru, ${updated} diperbarui dari cloud.`
          : 'Cloud sudah sama dengan lokal.',
        added || updated ? 'ok' : 'info'
      )
    } catch (e) {
      setCloudOk(false)
      notify(e.message, 'err')
    } finally {
      setBusy('')
    }
  }

  const doPushAll = async () => {
    if (!cloud.url) return notify('Isi URL Apps Script dulu.', 'err')
    const entries = Object.entries(signatures).filter(([, s]) => s?.dataUrl)
    if (!entries.length) return notify('Belum ada TTD lokal.', 'err')
    const nameOf = (key) => persons.find((p) => p.key === key)?.name ?? key
    setBusy('Mengirim TTD ke cloud…')
    try {
      let n = 0
      for (const [key, s] of entries) {
        setBusy(`Mengirim ${++n}/${entries.length}…`)
        await cloudPush(cloud.url, {
          key,
          name: nameOf(key),
          dataUrl: s.dataUrl,
          updatedAt: s.updatedAt || 0,
        })
      }
      setCloudOk(true)
      notify(`${entries.length} TTD terkirim ke cloud.`, 'ok')
    } catch (e) {
      setCloudOk(false)
      notify(e.message, 'err')
    } finally {
      setBusy('')
    }
  }

  const doDriveList = async (silent = false) => {
    if (!cloud.url) {
      if (!silent) notify('Isi URL Apps Script dulu.', 'err')
      return
    }
    if (!silent) setBusy('Memuat daftar Drive…')
    try {
      const j = await cloudFiles(cloud.url)
      setDriveFiles(j.files || [])
      setCloudOk(true)
      if (!silent) {
        notify(
          j.files?.length ? `${j.files.length} file Excel di Drive.` : 'Folder Drive kosong — upload file dulu.',
          'info'
        )
      }
    } catch (e) {
      setCloudOk(false)
      if (!silent) notify(e.message, 'err')
    } finally {
      if (!silent) setBusy('')
    }
  }

  const doDriveLoad = async (df) => {
    setBusy(`Mengunduh ${shortName(df.name)}…`)
    try {
      const j = await cloudGet(cloud.url, df.id)
      await addOneBatch([{ name: j.name || df.name, buf: b64ToBuf(j.base64) }])
    } catch (e) {
      notify(e.message, 'err')
    } finally {
      setBusy('')
    }
  }

  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

  const doDriveUpload = async (f) => {
    if (!f) return notify('Pilih file dulu.', 'err')
    setBusy(`Mengunggah ${shortName(f.name)}…`)
    try {
      await cloudSave(cloud.url, {
        name: f.name,
        base64: bufToB64(f.buffer.slice(0)),
        mime: XLSX_MIME,
      })
      notify(`${shortName(f.name)} tersimpan di Drive.`, 'ok')
      doDriveList(true)
    } catch (e) {
      notify(e.message, 'err')
    } finally {
      setBusy('')
    }
  }

  const doDriveSaveSigned = async (f) => {
    if (!f) return notify('Pilih file dulu.', 'err')
    if (signedCount === 0) return notify('Isi minimal 1 tanda tangan dulu.', 'err')
    setBusy(`Menyimpan hasil ${shortName(f.name)}…`)
    try {
      const wb = await buildSignedWorkbook({
        originalBuffer: f.buffer.slice(0),
        blocks: f.blocks,
        signatures,
      })
      const buf = await wb.xlsx.writeBuffer()
      const outName = f.name.replace(/\.xlsx?$/i, '') + '-signed.xlsx'
      await cloudSave(cloud.url, { name: outName, base64: bufToB64(buf), mime: XLSX_MIME })
      notify(`${outName} tersimpan di Drive.`, 'ok')
      doDriveList(true)
    } catch (e) {
      console.error(e)
      notify('Gagal menyimpan ke Drive: ' + e.message, 'err')
    } finally {
      setBusy('')
    }
  }

  const primaryAction = () => {    if (!files.length) {
      notify('Upload file Excel dulu — bisa pilih banyak sekaligus.', 'info')
      setMtab('file')
      return
    }
    if (nextUp) {
      setEditing(nextUp.key)
      setMtab('nama')
      return
    }
    doExportZip()
  }

  const ctaLabel = !files.length
    ? '📤 Upload file'
    : nextUp
      ? `Isi: ${nextUp.name}`
      : `⬇ ZIP (${files.length})`

  const jumpToPage = (v) => {
    if (!v || !activeFile) return
    document.getElementById(`${activeFile.id}-pg${v}`)?.scrollIntoView({ behavior: 'smooth', block: 'start' })
  }

  const editingPerson = editing ? persons.find((p) => p.key === editing) : null
  const editIdx = editingPerson ? unsignedList.findIndex((p) => p.key === editing) : -1
  const nextAfterEdit = editIdx >= 0
    ? (unsignedList[editIdx + 1] ?? unsignedList.find((p) => p.key !== editing) ?? null)
    : unsignedList[0] ?? null

  return (
    <>
      <header className="top">
        <div className="brand">
          <h1>✍️ TTD Laporan Bulanan</h1>
        </div>
        <div className="prog" role="status" aria-label={`Progres ${signedCount} dari ${persons.length}`}>
          <div className="progbar"><span style={{ width: `${pct}%` }} /></div>
          <span className="progtext">{persons.length ? `${signedCount}/${persons.length}` : '—'}</span>
        </div>
        <button className={`hcta ${!files.length || nextUp ? 'primary' : 'ok'}`} onClick={primaryAction}>
          {!files.length ? 'Mulai' : nextUp ? 'Isi TTD' : '⬇ ZIP'}
        </button>
      </header>
      {busy && <div className="busybar" role="status">{busy}</div>}

      {/* Tab navigasi HP */}
      <nav className="mtab" aria-label="Navigasi">
        {[
          ['file', 'File', files.length || ''],
          ['nama', 'TTD', persons.length ? `${signedCount}/${persons.length}` : ''],
          ['preview', 'Preview', ''],
        ].map(([k, label, n]) => (
          <button key={k} className={mtab === k ? 'on' : ''} onClick={() => setMtab(k)}>
            {label}{n !== '' && <span className="n">{n}</span>}
          </button>
        ))}
      </nav>

      <div className="layout" data-mtab={mtab}>
        {/* LANGKAH 1 — FILE */}
        <div className="pane pane-file">
          <div className="card">
            <h2>File <span className="mut">{files.length ? `• ${files.length} file, ${totalBlocks} hlm` : ''}</span></h2>
            <div
              className={`dropzone mini${drag ? ' over' : ''}`}
              onDragOver={(e) => { e.preventDefault(); setDrag(true) }}
              onDragLeave={() => setDrag(false)}
              onDrop={(e) => { e.preventDefault(); setDrag(false); readInputFiles(e.dataTransfer.files) }}
            >
              <label className="btn upload">
                ＋ Tambah file Excel
                <input type="file" accept=".xlsx,.xls" multiple hidden onChange={onUpload} />
              </label>
              {!files.length && (
                <button className="linkbtn" onClick={loadSample}>atau coba file contoh →</button>
              )}
            </div>

            {files.map((f) => {
              const st = fileSignedInfo(f)
              const active = activeFile?.id === f.id
              const p = st.need ? Math.round((st.hit / st.need) * 100) : 0
              return (
                <div key={f.id} className={`frow ${active ? 'active' : ''}`} onClick={() => { setActiveId(f.id); setMtab('preview') }} title={`${f.name} — ${st.hit}/${st.need} slot TTD`}>
                  <span className={`dot${st.done ? ' ok' : st.hit ? ' half' : ''}`} />
                  <span className="fname">{shortName(f.name)}</span>
                  <button
                    className="iconbtn"
                    onClick={(e) => { e.stopPropagation(); doExportOne(f) }}
                    title="Download XLSX bertanda tangan"
                    aria-label={`Download ${shortName(f.name)}`}
                  >⬇</button>
                  <button
                    className={`iconbtn${confirmDel === f.id ? ' danger' : ''}`}
                    onClick={(e) => { e.stopPropagation(); askRemove(f.id) }}
                    title="Hapus dari daftar"
                    aria-label={`Hapus ${shortName(f.name)}`}
                  >
                    {confirmDel === f.id ? '?' : '✕'}
                  </button>
                  <span className="fbar"><span style={{ width: `${p}%` }} /></span>
                </div>
              )
            })}
            {files.length > 1 && (
              <button className={`linkbtn dim${confirmDel === '__all__' ? ' danger-text' : ''}`} onClick={askClearAll}>
                {confirmDel === '__all__' ? 'Klik lagi untuk hapus semua' : 'Hapus semua'}
              </button>
            )}
          </div>

          {/* CLOUD — database TTD di Google Sheets */}
          <div className="card" style={{ marginTop: 12 }}>
            <h2>Cloud <span className="mut">• {!cloud.url ? 'mati (lokal saja)' : cloudOk ? '🟢 tersambung' : '⚪ belum dites'}</span></h2>
            {(showCloud || !cloud.url) && (
              <>
                <input
                  type="url" inputMode="url" autoComplete="off"
                  placeholder="https://script.google.com/macros/s/…/exec"
                  value={cloudUrl} onChange={(e) => setCloudUrl(e.target.value)}
                  aria-label="URL Web App Apps Script"
                />
                <div className="toolbar" style={{ marginTop: 8 }}>
                  <button className="primary" onClick={saveCloudUrl}>Simpan & Tes</button>
                  {cloud.url && <button onClick={() => setShowCloud(false)}>Tutup</button>}
                </div>
              </>
            )}
            {cloud.url && !showCloud && (
              <button className="linkbtn" onClick={() => setShowCloud(true)}>Ubah URL</button>
            )}
            {cloud.url && (
              <>
                <label className="switch" style={{ marginTop: 4 }}>
                  <input
                    type="checkbox" checked={cloud.autoSync}
                    onChange={(e) => setCloud({ ...cloud, autoSync: e.target.checked })}
                  />
                  <span>Kirim otomatis tiap TTD baru</span>
                </label>
                <div className="toolbar" style={{ marginTop: 8 }}>
                  <button onClick={doPull}>⬇ Ambil TTD</button>
                  <button onClick={doPushAll}>⬆ Kirim semua</button>
                </div>
              </>
            )}
          </div>

          {/* DRIVE — file laporan di Google Drive */}
          <div className="card" style={{ marginTop: 12 }}>
            <h2>Drive <span className="mut">• {driveFiles.length ? `${driveFiles.length} file` : 'folder laporan'}</span></h2>
            {!cloud.url && <p className="mut" style={{ margin: '4px 0' }}>Hubungkan cloud dulu untuk akses Drive.</p>}
            {cloud.url && (
              <div className="toolbar" style={{ marginBottom: 6 }}>
                <button onClick={() => doDriveList(false)}>Muat daftar</button>
              </div>
            )}
            {driveFiles.map((df) => (
              <div key={df.id} className="frow" title={`${df.name} — klik untuk memuat`}>
                <span className="fname" onClick={() => doDriveLoad(df)}>{shortName(df.name)}</span>
                <button
                  className="iconbtn"
                  onClick={() => doDriveLoad(df)}
                  title={`Muat ${shortName(df.name)} ke app`}
                  aria-label={`Muat ${shortName(df.name)}`}
                >⬇</button>
              </div>
            ))}
            <div className="explinks">
              <button className="linkbtn" onClick={() => activeFile && doDriveUpload(activeFile)} disabled={!activeFile || !cloud.url}>Upload file ini (asli)</button>
              <span aria-hidden>·</span>
              <button className="linkbtn" onClick={() => activeFile && doDriveSaveSigned(activeFile)} disabled={!activeFile || !cloud.url}>Simpan hasil signed</button>
            </div>
          </div>
        </div>

        {/* LANGKAH 2 — TANDA TANGAN */}
        <div className="pane pane-nama">
          <div className="card">
            <h2>Tanda tangan <span className="mut">• {signedCount}/{persons.length || '–'}</span></h2>
            {!persons.length && (
              <div className="empty">
                <div className="emptyart">📁</div>
                <p className="mut">Upload file dulu.</p>
                <button className="primary" onClick={() => setMtab('file')}>Upload →</button>
              </div>
            )}
            {!!persons.length && nextUp && (
              <button className="nextup" onClick={() => setEditing(nextUp.key)}>
                <span className="nu-text">Berikutnya: <b>{nextUp.name}</b> <span className="mut">({nextUp.count}×)</span></span>
                <span className="nu-go">Isi →</span>
              </button>
            )}
            {!!persons.length && !nextUp && (
              <button className="nextup done" onClick={() => setMtab('preview')}>
                <span className="nu-text">✅ Semua lengkap — siap export</span>
                <span className="nu-go">→</span>
              </button>
            )}
            {!!persons.length && (
              <div className="chips" role="tablist" aria-label="Filter nama">
                {[['semua', 'Semua'], ['belum', `Belum ${remaining}`], ['sudah', `Sudah ${signedCount}`]].map(([k, l]) => (
                  <button key={k} className={filter === k ? 'on' : ''} onClick={() => setFilter(k)}>{l}</button>
                ))}
                <input
                  type="search" enterKeyHint="search" autoComplete="off"
                  placeholder="Cari…" value={q} onChange={(e) => setQ(e.target.value)}
                  aria-label="Cari nama"
                />
              </div>
            )}
            {filtered.map((p) => {
              const sig = signatures[p.key]
              const done = Boolean(sig?.dataUrl)
              return (
                <button key={p.key} className={`person${done ? ' done' : ''}`} onClick={() => setEditing(p.key)} title={done ? 'Klik untuk ubah' : 'Klik untuk isi'}>
                  <span className="pthumb">{sig?.dataUrl ? <img src={sig.dataUrl} alt="" /> : '＋'}</span>
                  <span className="ptext">
                    <b>{p.name}</b>
                    <small>{p.count}× • {p.files.size} file • {done ? 'sudah' : 'belum'}</small>
                  </span>
                  <span className="pgo">{done ? 'Ubah' : 'Isi'}</span>
                </button>
              )
            })}
            {!!persons.length && !filtered.length && <div className="mut">Tidak ketemu. <button className="linkbtn" onClick={() => { setQ(''); setFilter('semua') }}>Reset</button></div>}
          </div>
        </div>

        {/* LANGKAH 3 — PERIKSA & EXPORT */}
        <div className="pane pane-preview">
          <div className="card export">
            <div className="expgrid">
              <div className="expinfo">
                <b>Export</b>
                <span className="mut">{files.length} file • {signedCount}/{persons.length} TTD{remaining > 0 && persons.length > 0 ? ` • ${remaining} belum` : ''}</span>
              </div>
              <button className="ok" onClick={doExportZip} disabled={!files.length || signedCount === 0}>
                ⬇ ZIP{files.length ? ` (${files.length})` : ''}
              </button>
            </div>
            <div className="explinks">
              <button className="linkbtn" onClick={() => activeFile && doExportOne(activeFile)} disabled={!activeFile}>Excel file ini</button>
              <span aria-hidden>·</span>
              <button className="linkbtn" onClick={doExportPdf} disabled={!activeFile}>PDF file ini</button>
            </div>
          </div>

          <div className="card previewbar" style={{ marginTop: 12 }}>
            <div className="pvrow">
              <b className="pvtitle">{activeFile ? shortName(activeFile.name) : 'Preview'}</b>
              <div className="seg" role="tablist" aria-label="Mode preview">
                <button className={view === 'laporan' ? 'on' : ''} onClick={() => setView('laporan')}>Laporan</button>
                <button className={view === 'ringkas' ? 'on' : ''} onClick={() => setView('ringkas')}>Ringkas</button>
              </div>
            </div>
            {files.length > 1 && (
              <div className="filetabs">
                {files.map((f) => {
                  const st = fileSignedInfo(f)
                  return (
                    <button
                      key={f.id}
                      className={activeFile?.id === f.id ? 'on' : ''}
                      onClick={() => { setActiveId(f.id); setMtab('preview') }}
                      title={`${f.name} — ${st.hit}/${st.need}`}
                    >
                      {st.done ? '✓ ' : ''}{shortName(f.name)}
                    </button>
                  )
                })}
              </div>
            )}
            {activeFile && view === 'laporan' && (
              <div className="pvrow tools">
                <select
                  key={activeFile.id}
                  defaultValue=""
                  onChange={(e) => jumpToPage(e.target.value)}
                  aria-label="Lompat ke halaman"
                >
                  <option value="">Halaman…</option>
                  {activeFile.blocks.map((b) => <option key={b.id} value={b.page}>{b.page}</option>)}
                </select>
                <div className="seg small" role="tablist" aria-label="Ukuran preview">
                  <button className={previewFit ? 'on' : ''} onClick={() => setPreviewFit(true)} title="Sesuaikan lebar layar">Pas</button>
                  <button className={!previewFit ? 'on' : ''} onClick={() => setPreviewFit(false)} title="Ukuran asli Excel">1:1</button>
                </div>
                <label className="switch mini" title="Tandai slot yang belum ada TTD-nya">
                  <input type="checkbox" checked={spot} onChange={(e) => setSpot(e.target.checked)} />
                  <span>Sorot</span>
                </label>
              </div>
            )}
          </div>

          <div id="ttdPreview" style={{ marginTop: 12 }}>
            {!activeFile && (
              <div className="card empty">
                <div className="emptyart">📄</div>
                <p className="mut">Belum ada preview. Upload file di langkah 1.</p>
              </div>
            )}
            {activeFile && view === 'laporan' && (
              <ReportPreview
                sheet={activeFile.sheet}
                aoa={activeFile.aoa}
                blocks={activeFile.blocks}
                signatures={signatures}
                fileUid={activeFile.id}
                spotMissing={spot}
                fit={previewFit}
                onEditSig={(key) => setEditing(key)}
              />
            )}
            {activeFile && view === 'ringkas' && (
              activeFile.blocks.map((b) => (
                <div key={b.id} className="block" id={`${activeFile.id}-pg${b.page}`}>
                  <div className="bh"><span>Hlm {b.page}</span></div>
                  <div className="cols" style={{ gridTemplateColumns: `repeat(${Math.min((b.slots || []).length, 3)}, 1fr)` }}>
                    {(b.slots || []).map((s, si) => (
                      <div key={si}>
                        <div className="lab">{s.role}</div>
                        <div className="names">
                          {s.names.map((nm) => {
                            const sig = signatures[normKey(nm)]
                            const on = sig?.dataUrl && sig?.enabled !== false
                            return (
                              <button
                                key={nm}
                                className={`chipbtn chip ${on ? 'signed' : 'miss'}`}
                                onClick={() => { setEditing(normKey(nm)); setMtab('nama') }}
                                title={on ? `${nm} (${s.role}) — klik untuk ubah` : `${nm} (${s.role}) — belum TTD, klik untuk isi`}
                              >
                                {on && <img src={sig.dataUrl} alt="" />}
                                <span>{nm}{on ? '' : ' • belum'}</span>
                              </button>
                            )
                          })}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              ))
            )}
          </div>
        </div>
      </div>

      {/* CTA lengket di HP */}
      <div className="actionbar">
        <span className="mut">{persons.length ? `${signedCount}/${persons.length} TTD` : 'Siap?'}</span>
        <button className={!files.length || nextUp ? 'primary' : 'ok'} onClick={primaryAction}>{ctaLabel}</button>
      </div>

      <div className="toasts" aria-live="polite">
        {toasts.map((t) => <div key={t.id} className={`toast ${t.kind}`}>{t.msg}</div>)}
      </div>

      {editingPerson && (
        <SignaturePad
          key={editing}
          personName={editingPerson.name}
          initial={signatures[editingPerson.key]?.dataUrl}
          enabled={signatures[editingPerson.key]?.enabled !== false}
          hasNext={Boolean(nextAfterEdit)}
          nextName={nextAfterEdit?.name}
          onClose={() => setEditing(null)}
          onToggleEnabled={() => toggleEnabled(editingPerson.key)}
          onDelete={() => { clearSig(editingPerson.key); setEditing(null) }}
          onSave={(url) => {
            const ts = Date.now()
            setSig(editingPerson.key, url, ts)
            pushSig(editingPerson.key, editingPerson.name, url, ts)
            setEditing(null)
            notify(`TTD ${editingPerson.name} tersimpan.`, 'ok')
          }}
          onSaveNext={saveAndNext}
        />
      )}
    </>
  )
}
