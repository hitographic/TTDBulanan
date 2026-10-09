import { useEffect, useMemo, useRef, useState } from 'react'
import './styles.css'
import { parseWorkbook, normKey, matchNameToUser } from './lib/excelParser'
import { buildSheetModel, normalizeWorkbookBuffer } from './lib/sheetModel'
import { exportSignedExcel, buildSignedWorkbook } from './lib/excelExport'
import { exportAllZip } from './lib/zipExport'
import { exportPreviewPdf } from './lib/pdfExport'
import {
  loadCloud, saveCloud, cloudPing, cloudPull, cloudFiles, cloudGet,
  cloudPush, cloudSave, cloudUsers, cloudLogin, cloudChanges,
  mergeSignatures, pickSig, diffSince, flushPending, bufToB64, b64ToBuf,
} from './lib/cloud'
import { loadSession, saveSession, clearSession } from './lib/auth'
import LoginPage from './components/LoginPage'
import { browserMeasure, geomForSide } from './lib/textGeom'
import SignaturePad from './components/SignaturePad'
import ReportPreview from './components/ReportPreview'
import UserFlow from './components/UserFlow'

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
  const [driveFiles, setDriveFiles] = useState([]) // [{id,name,month,monthId,size,modified}]
  const [session, setSession] = useState(loadSession) // {nik,name,nameInFile,role,aliases} | null
  const [users, setUsers] = useState([]) // daftar user (tanpa password) untuk alias
  const [months, setMonths] = useState([]) // [{id,name}] folder bulan di Drive
  const [monthSel, setMonthSel] = useState('__all__')
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

  // Tarik TTD baru bila ada perubahan di cloud (dipakai polling + tombol refresh)
  const checkUpdates = async (manual) => {
    if (!cloud.url || !session || session.offline) return false
    if (manual) setBusy('Memeriksa pembaruan…')
    try {
      const ch = await cloudChanges(cloud.url)
      const snap = JSON.stringify(ch.changes || {})
      if (snap === lastSnapRef.current) return false
      lastSnapRef.current = snap
      if (!diffSince(signaturesRef.current, ch.changes)) return false
      const j = await cloudPull(cloud.url)
      const applied = mergeSignatures(signaturesRef.current, j.signatures)
      if (applied.added + applied.updated === 0) return false
      setSignatures((cur) => mergeSignatures(cur, j.signatures).merged)
      if (applied.added > 0) notify(`Ada TTD baru dari rekan 🎉 (+${applied.added})`, 'ok')
      else notify('Ada pembaruan TTD dari cloud.', 'info')
      setCloudOk(true)
      return true
    } catch (e) {
      if (manual) notify(e.message, 'err')
      return false
    } finally {
      if (manual) setBusy('')
    }
  }

  // Poll cloud tiap 20 detik: TTD rekan langsung muncul + antrean kirim di-flush
  useEffect(() => {
    if (!session || session.offline) return undefined
    let stop = false
    const tick = async () => {
      if (stop || document.hidden) return
      try {
        await flushPending(pendingRef.current, (p) =>
          cloudPush(cloud.url, { key: p.key, name: p.name, dataUrl: p.dataUrl, updatedAt: p.updatedAt })
        )
        await checkUpdates(false)
      } catch { /* diam: coba lagi periode berikut */ }
    }
    const id = setInterval(tick, 20000)
    const onVis = () => { if (!document.hidden) tick() }
    document.addEventListener('visibilitychange', onVis)
    return () => {
      stop = true
      clearInterval(id)
      document.removeEventListener('visibilitychange', onVis)
    }
  }, [session, cloud.url]) // eslint-disable-line

  useEffect(() => () => clearTimeout(delTimer.current), [])

  // Parse mentah tanpa state/toast — dipakai upload manual & boot login
  async function parseItems(items) {
    const next = []
    const failed = []
    setBusy('Membaca file…')
    try {
      for (const { name, buf, month, monthId } of items) {
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
            month: month || '',
            monthId: monthId || '',
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
    return { next, failed }
  }

  async function addOneBatch(items, opts = {}) {
    const { next, failed } = await parseItems(items)
    failed.forEach((m) => notify(`Gagal membaca ${m}`, 'err'))
    if (!next.length) return []
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
    if (!opts.quiet) {
      const pages = next.reduce((n, f) => n + f.blocks.length, 0)
      notify(`${next.length} file dimuat • ${pages} halaman. Lanjut isi TTD.`, 'ok')
    }
    return next
  }

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

  /* ===== Auth & cakupan: user biasa hanya miliknya, admin semua ===== */
  const isAdmin = session?.role === 'admin'

  // Peta alias global: normKey(varian) -> kunci kanonis pemilik
  const aliasToKey = useMemo(() => {
    const m = new Map()
    users.forEach((u) => {
      const canon = normKey(u.nameInFile || '')
      if (!canon) return
      ;(u.aliases || []).forEach((a) => {
        const k = normKey(a)
        if (k && k !== canon) m.set(k, canon)
      })
    })
    return m
  }, [users])

  // Kunci kanonis + varian milik akun login (untuk fuzzy inisial)
  const myCanonKey = session ? normKey(session.nameInFile) : null
  const scopeList = useMemo(() => {
    if (!session || isAdmin) return []
    const out = []
    if (session.nameInFile) out.push({ raw: session.nameInFile, key: myCanonKey })
    ;(session.aliases || []).forEach((a) => out.push({ raw: a, key: myCanonKey }))
    return out
  }, [session, isAdmin, myCanonKey])

  // TTD efektif untuk sebuah nama di file
  const sigOf = (name) => pickSig(signatures, aliasToKey, name, scopeList)

  const fileHasUser = (f, ses) => {
    if (!ses || ses.role === 'admin') return true
    const mine = { nameInFile: ses.nameInFile, aliases: ses.aliases }
    return f.persons.some((p) => matchNameToUser(p.name, mine))
  }

  // Peta TTD per file untuk export/preview (varian ejaan ikut terisi)
  const effSigsFor = (f) => {
    const m = { ...signatures }
    slotNames(f).forEach(({ name }) => {
      const s = sigOf(name)
      if (s) m[normKey(name)] = s
    })
    return m
  }

  // Kunci penyimpanan: user biasa selalu ke kunci kanonisnya (1 TTD utk semua varian)
  const storeKeyFor = (key) => (isAdmin || !session ? key : myCanonKey || key)

  const canSign = (key, name) => {
    if (isAdmin) return true
    if (!session) return false
    if (myCanonKey && (key === myCanonKey || aliasToKey.get(key) === myCanonKey)) return true
    return matchNameToUser(name || key, session)
  }

  const handleEditSig = (key, name) => {
    if (canSign(key, name)) {
      setEditing(key)
      return true
    }
    notify(`Slot ini milik ${name || key} — hanya pemilik yang bisa mengisi.`, 'err')
    return false
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

  // User biasa: progres & daftar hanya atas namanya sendiri
  const scopePersons = useMemo(() => {
    if (isAdmin || !session) return persons
    const mine = { nameInFile: session.nameInFile, aliases: session.aliases }
    return persons.filter((p) => matchNameToUser(p.name, mine))
  }, [persons, session, isAdmin])

  const filtered = useMemo(() => {
    const s = q.trim().toLowerCase()
    let arr = scopePersons
    if (filter === 'belum') arr = arr.filter((p) => !sigOf(p.name)?.dataUrl)
    else if (filter === 'sudah') arr = arr.filter((p) => sigOf(p.name)?.dataUrl)
    if (s) arr = arr.filter((p) => p.name.toLowerCase().includes(s))
    return arr
  }, [scopePersons, q, filter, signatures, aliasToKey, session, isAdmin])

  const signedCount = scopePersons.filter((p) => sigOf(p.name)?.dataUrl).length
  const unsignedList = useMemo(
    () => scopePersons.filter((p) => !sigOf(p.name)?.dataUrl),
    [scopePersons, signatures, aliasToKey, session, isAdmin]
  )
  const nextUp = unsignedList[0] ?? null
  const totalBlocks = files.reduce((n, f) => n + f.blocks.length, 0)
  const pct = scopePersons.length ? Math.round((signedCount / scopePersons.length) * 100) : 0
  const remaining = scopePersons.length - signedCount

  const setSig = (key, dataUrl, ts = Date.now()) =>
    setSignatures((m) => ({ ...m, [key]: { dataUrl, enabled: true, updatedAt: ts } }))

  // Simpan lokal + kirim ke cloud. 'ok' | 'local' | 'fail' (fail = masuk antrean coba lagi)
  const SAVE_MSG = 'Menyimpan TTD ke cloud…'
  const pendingRef = useRef({}) // antrean kirim ulang bila gagal
  const signaturesRef = useRef(signatures)
  const lastSnapRef = useRef('')
  useEffect(() => {
    signaturesRef.current = signatures
  }, [signatures])

  const syncOne = async (key, name, dataUrl) => {
    const ts = Date.now()
    setSig(key, dataUrl, ts)
    if (!cloud.url) return 'local'
    setBusy(SAVE_MSG)
    try {
      await cloudPush(cloud.url, { key, name, dataUrl, updatedAt: ts })
      return 'ok'
    } catch (e) {
      console.warn('kirim TTD gagal, masuk antrean:', e.message)
      pendingRef.current[key] = { key, name, dataUrl, updatedAt: ts }
      return 'fail'
    } finally {
      setBusy((b) => (b === SAVE_MSG ? '' : b))
    }
  }

  const saveAndNext = async (url) => {
    if (!editing) return
    const storeKey = storeKeyFor(editing)
    const nm = persons.find((p) => p.key === editing)?.name ?? editing
    const idx = unsignedList.findIndex((p) => p.key === editing)
    const next = unsignedList[idx + 1] ?? unsignedList.find((p) => p.key !== editing) ?? null
    if (next) {
      setEditing(next.key)
      notify(`TTD ${nm} tersimpan. Lanjut: ${next.name}.`, 'ok')
    } else {
      setEditing(null)
      notify('Semua TTD lengkap. Siap export. 🎉', 'ok')
      setMtab('preview')
    }
    if (!cloud.autoSync) {
      setSig(storeKey, url)
      return
    }
    const r = await syncOne(storeKey, nm, url)
    if (r === 'fail') notify('Gagal kirim ke cloud — tersimpan lokal, dicoba otomatis. 🔄', 'err')
  }

  const toggleEnabled = (key) => {
    const sk = storeKeyFor(key)
    setSignatures((m) => ({
      ...m,
      [sk]: { dataUrl: m[sk]?.dataUrl, enabled: m[sk]?.enabled === false ? true : false },
    }))
  }

  const clearSig = (key) => {
    const sk = storeKeyFor(key)
    setSignatures((m) => { const n = { ...m }; delete n[sk]; return n })
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
      const s = sigOf(name)
      if (s?.dataUrl && s?.enabled !== false && !seen.has(`${blockId}-${k}`)) {
        seen.add(`${blockId}-${k}`)
        hit += 1
      }
    })
    return { hit, need, done: need > 0 && hit >= need }
  }

  const doExportOne = async (f) => {
    if (!isAdmin) return notify('Hanya admin yang bisa download.', 'err')
    if (!f) return notify('Pilih file dulu.', 'err')
    if (signedCount === 0) return notify('Isi minimal 1 tanda tangan dulu.', 'err')
    try {
      await exportSignedExcel({
        originalBuffer: f.buffer,
        blocks: f.blocks,
        signatures: effSigsFor(f),
        outName: f.name.replace(/\.xlsx?$/i, '') + '-signed.xlsx',
      })
      notify(`${shortName(f.name)} terdownload.`, 'ok')
    } catch (e) {
      console.error(e)
      notify('Gagal export: ' + e.message, 'err')
    }
  }

  const doExportZip = async () => {
    if (!isAdmin) return notify('Hanya admin yang bisa download.', 'err')
    if (!files.length) return notify('Upload file Excel dulu.', 'err')
    if (signedCount === 0) return notify('Isi minimal 1 tanda tangan dulu.', 'err')
    setBusy('Menyiapkan ZIP…')
    try {
      await exportAllZip({
        files,
        signatures,
        signaturesForFile: (f) => effSigsFor(f),
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
    if (!isAdmin) return notify('Hanya admin yang bisa download.', 'err')
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
    // User biasa hanya boleh mendorong TTD miliknya sendiri
    const mine = (key) => isAdmin || !session || (myCanonKey && (key === myCanonKey || aliasToKey.get(key) === myCanonKey))
    const entries = Object.entries(signatures).filter(([key, s]) => s?.dataUrl && mine(key))
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
      await addOneBatch([{
        name: j.name || df.name,
        buf: b64ToBuf(j.base64),
        month: df.month || '',
        monthId: df.monthId || '',
      }])
    } catch (e) {
      notify(e.message, 'err')
    } finally {
      setBusy('')
    }
  }

  // Admin: muat seluruh file dalam satu bulan sekaligus
  const doDriveLoadMonth = async (monthId) => {
    const list = driveFiles.filter((df) => df.monthId === monthId)
    if (!list.length) return notify('Bulan ini kosong.', 'info')
    const mname = months.find((m) => m.id === monthId)?.name || ''
    setBusy(`Mengunduh ${list.length} file ${mname}…`)
    try {
      const items = []
      let n = 0
      for (const df of list) {
        setBusy(`Mengunduh ${++n}/${list.length}… ${shortName(df.name)}`)
        try {
          const j = await cloudGet(cloud.url, df.id)
          items.push({
            name: j.name || df.name,
            buf: b64ToBuf(j.base64),
            month: df.month || '',
            monthId: df.monthId || '',
          })
        } catch (e) { console.warn('unduh gagal:', df.name, e.message) }
      }
      await addOneBatch(items)
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
        folderId: f.monthId || '',
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
        signatures: effSigsFor(f),
      })
      const buf = await wb.xlsx.writeBuffer()
      const outName = f.name.replace(/\.xlsx?$/i, '') + '-signed.xlsx'
      await cloudSave(cloud.url, { name: outName, base64: bufToB64(buf), mime: XLSX_MIME, folderId: f.monthId || '' })
      notify(`${outName} tersimpan di Drive.`, 'ok')
      doDriveList(true)
    } catch (e) {
      console.error(e)
      notify('Gagal menyimpan ke Drive: ' + e.message, 'err')
    } finally {
      setBusy('')
    }
  }

  /* ===== Auth: login NIK + boot sesuai akun ===== */
  const doLogin = async (nik, password) => {
    setBusy('Masuk…')
    try {
      const j = await cloudLogin(cloud.url, nik, password)
      const ses = { ...j.user, ts: Date.now(), offline: false }
      setSession(ses)
      saveSession(ses)
      notify(`Halo, ${ses.name || ses.nik} 👋`, 'ok')
      await boot(ses, false)
    } catch (e) {
      // Offline: sesi yang pernah login di HP ini boleh lanjut tanpa password
      const cached = loadSession()
      if (cached && cached.nik === String(nik).trim()) {
        const ses = { ...cached, offline: true }
        setSession(ses)
        notify('Offline — memakai sesi tersimpan di HP ini.', 'info')
        await boot(ses, true)
      } else {
        throw e
      }
    } finally {
      setBusy('')
    }
  }

  const doLogout = () => {
    clearSession()
    setSession(null)
    setFiles([])
    setActiveId(null)
    setUsers([])
    setDriveFiles([])
    setMonths([])
    setMonthSel('__all__')
    setEditing(null)
    setMtab('preview')
    notify('Keluar. Sampai jumpa 👋', 'info')
  }

  // Muat TTD + user + file Drive sesuai akun (admin = semua, user = miliknya)
  const boot = async (ses, offline) => {
    const admin = ses.role === 'admin'
    setBusy('Menyiapkan…')
    try {
      if (!offline) {
        try {
          const u = await cloudUsers(cloud.url)
          setUsers(u.users || [])
        } catch (e) { console.warn('daftar user gagal:', e.message) }
        try {
          const j = await cloudPull(cloud.url)
          setSignatures((prev) => mergeSignatures(prev, j.signatures).merged)
          setCloudOk(true)
        } catch (e) { console.warn('pull gagal:', e.message) }
      }
      let tree = { months: [], files: [] }
      if (!offline) {
        try {
          tree = await cloudFiles(cloud.url)
          setDriveFiles(tree.files || [])
          setMonths(tree.months || [])
          setMonthSel('__all__')
        } catch (e) { console.warn('daftar drive gagal:', e.message) }
      }
      const items = []
      let n = 0
      if (admin) {
        // Admin mengatur sendiri: pilih bulan di kartu Drive untuk dimuat
        if (!(tree.files || []).length) {
          await loadSample()
        } else {
          notify(`Drive siap: ${(tree.months || []).length} bulan, ${tree.files.length} file. Pilih bulan di kartu Drive.`, 'info')
        }
        return
      }
      for (const df of tree.files || []) {
        setBusy(`Mengunduh ${++n}/${tree.files.length}… ${shortName(df.name)}`)
        try {
          const j = await cloudGet(cloud.url, df.id)
          items.push({
            name: j.name || df.name,
            buf: b64ToBuf(j.base64),
            month: df.month || '',
            monthId: df.monthId || '',
          })
        } catch (e) { console.warn('unduh gagal:', df.name, e.message) }
      }
      const { next, failed } = await parseItems(items)
      failed.forEach((m) => notify(`Gagal membaca ${m}`, 'err'))
      const kept = admin ? next : next.filter((f) => fileHasUser(f, ses))
      setFiles(kept)
      setActiveId(kept[0]?.id ?? null)
      if (kept.length) {
        const hit = new Set(kept.map((f) => f.month || 'Lainnya')).size
        notify(`${kept.length} file • ${hit} bulan dimuat. Selamat bekerja! 🎉`, 'ok')
      } else if (admin) {
        await loadSample()
      } else {
        notify('Belum ada file laporan untuk akun ini di Drive.', 'info')
      }
    } finally {
      setBusy('')
    }
  }

  const primaryAction = () => {
    if (!files.length) {
      if (isAdmin) {
        notify('Upload file Excel dulu — bisa pilih banyak sekaligus.', 'info')
        setMtab('file')
      } else {
        notify('Belum ada file laporan untuk akun ini.', 'info')
      }
      return
    }
    if (nextUp) {
      handleEditSig(nextUp.key, nextUp.name)
      setMtab('nama')
      return
    }
    if (isAdmin) {
      doExportZip()
      return
    }
    notify('Semua TTD lengkap ✅', 'ok')
    setMtab('preview')
  }

  const ctaLabel = !files.length
    ? (isAdmin ? '📤 Upload file' : '📄 Laporan')
    : nextUp
      ? `Isi: ${nextUp.name}`
      : (isAdmin ? `⬇ ZIP (${files.length})` : '✅ Lengkap')

  const jumpToPage = (v) => {
    if (!v || !activeFile) return
    document.getElementById(`${activeFile.id}-pg${v}`)?.scrollIntoView({ behavior: 'smooth', block: 'start' })
  }

  const editingPerson = editing ? persons.find((p) => p.key === editing) : null
  const editIdx = editingPerson ? unsignedList.findIndex((p) => p.key === editing) : -1
  const nextAfterEdit = editIdx >= 0
    ? (unsignedList[editIdx + 1] ?? unsignedList.find((p) => p.key !== editing) ?? null)
    : unsignedList[0] ?? null

  // Finale user: pastikan TTD miliknya untuk file ini terkirim ke cloud
  const pushOwnForFile = async (f) => {
    if (!cloud.url) return false
    try {
      const jobs = []
      slotNames(f).forEach(({ name }) => {
        let k = normKey(name)
        if (!isAdmin && session) {
          if (!matchNameToUser(name, session)) return
          k = myCanonKey
        }
        const s = signaturesRef.current[k]
        if (s?.dataUrl && !jobs.find((j) => j.key === k)) {
          jobs.push({ key: k, name, dataUrl: s.dataUrl, updatedAt: s.updatedAt || 0 })
        }
      })
      for (const j of jobs) {
        await cloudPush(cloud.url, j)
      }
      await flushPending(pendingRef.current, (p) =>
        cloudPush(cloud.url, { key: p.key, name: p.name, dataUrl: p.dataUrl, updatedAt: p.updatedAt })
      )
      return true
    } catch {
      return false
    }
  }

  const toastEl = (
    <div className="toasts" aria-live="polite">
      {toasts.map((t) => <div key={t.id} className={`toast ${t.kind}`}>{t.msg}</div>)}
    </div>
  )

  const padEl = editingPerson && (
    <SignaturePad
      key={editing}
      personName={editingPerson.name}
      initial={sigOf(editingPerson.name)?.dataUrl}
      enabled={sigOf(editingPerson.name)?.enabled !== false}
      hasNext={isAdmin && Boolean(nextAfterEdit)}
      nextName={nextAfterEdit?.name}
      onClose={() => setEditing(null)}
      onToggleEnabled={() => toggleEnabled(editingPerson.key)}
      onDelete={() => { clearSig(editingPerson.key); setEditing(null) }}
      onSave={async (url) => {
        const storeKey = storeKeyFor(editingPerson.key)
        const nm = editingPerson.name
        setEditing(null)
        if (!cloud.autoSync) {
          setSig(storeKey, url)
          notify(`TTD ${nm} tersimpan.`, 'ok')
          return
        }
        const r = await syncOne(storeKey, nm, url)
        notify(
          r === 'ok' ? `TTD ${nm} tersimpan di cloud ✅`
          : r === 'local' ? `TTD ${nm} tersimpan.`
          : `TTD ${nm} tersimpan lokal ⚠️ — koneksi gagal, dicoba otomatis. 🔄`,
          r === 'fail' ? 'err' : 'ok'
        )
      }}
      onSaveNext={saveAndNext}
    />
  )

  if (!session) {
    return (
      <>
        {busy && <div className="busybar" role="status">{busy}</div>}
        <LoginPage onLogin={doLogin} />
        {toastEl}
      </>
    )
  }

  // User biasa: alur game 3 tahap (admin: dashboard penuh di bawah)
  if (!isAdmin) {
    return (
      <>
        {busy && <div className="busybar" role="status">{busy}</div>}
        <UserFlow
          session={session}
          files={files}
          sigOf={sigOf}
          isMine={(name) => matchNameToUser(name, session)}
          effSigsFor={effSigsFor}
          onRequestSign={(key) => handleEditSig(key, persons.find((p) => p.key === key)?.name ?? key)}
          onFinalSave={pushOwnForFile}
          onLogout={doLogout}
        />
        {toastEl}
        {padEl}
      </>
    )
  }

  return (
    <>
      <header className="top">
        <div className="brand">
          <h1>✍️ TTD Laporan Bulanan</h1>
          {session && <p>{session.name || session.nik}{isAdmin ? ' 👑 admin' : ''}{session.offline ? ' (offline)' : ''}</p>}
        </div>
        <div className="prog" role="status" aria-label={`Progres ${signedCount} dari ${scopePersons.length}`}>
          <div className="progbar"><span style={{ width: `${pct}%` }} /></div>
          <span className="progtext">{scopePersons.length ? `${signedCount}/${scopePersons.length}` : '—'}</span>
        </div>
        {session && <button onClick={doLogout} title="Keluar dari akun">Keluar</button>}
        <button className={`hcta ${!files.length || nextUp ? 'primary' : 'ok'}`} onClick={primaryAction}>
          {!files.length ? 'Mulai' : nextUp ? 'Isi TTD' : '⬇ ZIP'}
        </button>
      </header>
      {busy && <div className="busybar" role="status">{busy}</div>}

      {/* Tab navigasi HP */}
      <nav className="mtab" aria-label="Navigasi">
        {[
          ['file', 'File', files.length || ''],
          ['nama', 'TTD', scopePersons.length ? `${signedCount}/${scopePersons.length}` : ''],
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
            {isAdmin && (
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
            )}
            {!isAdmin && !files.length && (
              <div className="empty">
                <div className="emptyart">📄</div>
                <p className="mut">Belum ada laporan untukmu.</p>
              </div>
            )}

            {files.map((f) => {
              const st = fileSignedInfo(f)
              const active = activeFile?.id === f.id
              const p = st.need ? Math.round((st.hit / st.need) * 100) : 0
              return (
                <div key={f.id} className={`frow ${active ? 'active' : ''}`} onClick={() => { setActiveId(f.id); setMtab('preview') }} title={`${f.month ? `${f.month} • ` : ''}${f.name} — ${st.hit}/${st.need} slot TTD`}>
                  <span className={`dot${st.done ? ' ok' : st.hit ? ' half' : ''}`} />
                  <span className="fname">{shortName(f.name)}</span>
                  {f.month && <span className="badge">{f.month}</span>}
                  {isAdmin && (
                  <button
                    className="iconbtn"
                    onClick={(e) => { e.stopPropagation(); doExportOne(f) }}
                    title="Download XLSX bertanda tangan"
                    aria-label={`Download ${shortName(f.name)}`}
                  >⬇</button>
                  )}
                  {isAdmin && (
                  <button
                    className={`iconbtn${confirmDel === f.id ? ' danger' : ''}`}
                    onClick={(e) => { e.stopPropagation(); askRemove(f.id) }}
                    title="Hapus dari daftar"
                    aria-label={`Hapus ${shortName(f.name)}`}
                  >
                    {confirmDel === f.id ? '?' : '✕'}
                  </button>
                  )}
                  <span className="fbar"><span style={{ width: `${p}%` }} /></span>
                </div>
              )
            })}
            {isAdmin && files.length > 1 && (
              <button className={`linkbtn dim${confirmDel === '__all__' ? ' danger-text' : ''}`} onClick={askClearAll}>
                {confirmDel === '__all__' ? 'Klik lagi untuk hapus semua' : 'Hapus semua'}
              </button>
            )}
          </div>

          {/* CLOUD — database TTD di Google Sheets (admin saja) */}
          {isAdmin && (
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
          )}

          {/* DRIVE — file laporan di Google Drive (admin saja) */}
          {isAdmin && (
          <div className="card" style={{ marginTop: 12 }}>
            <h2>Drive <span className="mut">• {driveFiles.length ? `${driveFiles.length} file` : 'folder laporan'}</span></h2>
            {!cloud.url && <p className="mut" style={{ margin: '4px 0' }}>Hubungkan cloud dulu untuk akses Drive.</p>}
            {cloud.url && (
              <div className="toolbar" style={{ marginBottom: 6 }}>
                <button onClick={() => doDriveList(false)}>Muat daftar</button>
              </div>
            )}
            {months.length > 0 && (
              <div className="chips" role="tablist" aria-label="Filter bulan">
                {[['__all__', 'Semua'], ...months.map((m) => [m.id, m.name])].map(([k, l]) => (
                  <button key={k} className={monthSel === k ? 'on' : ''} onClick={() => setMonthSel(k)}>{l}</button>
                ))}
              </div>
            )}
            {monthSel !== '__all__' && (
              <div className="toolbar" style={{ marginBottom: 6 }}>
                <button className="primary" onClick={() => doDriveLoadMonth(monthSel)}>
                  ⬇ Muat bulan {months.find((m) => m.id === monthSel)?.name || ''}
                </button>
              </div>
            )}
            {(monthSel === '__all__' ? driveFiles : driveFiles.filter((df) => df.monthId === monthSel)).map((df) => (
              <div key={df.id} className="frow" title={`${df.month ? `${df.month} • ` : ''}${df.name} — klik untuk memuat`}>
                <span className="fname" onClick={() => doDriveLoad(df)}>{shortName(df.name)}</span>
                {monthSel === '__all__' && df.month && <span className="badge">{df.month}</span>}
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
          )}
        </div>

        {/* LANGKAH 2 — TANDA TANGAN */}
        <div className="pane pane-nama">
          <div className="card">
            <h2>Tanda tangan <span className="mut">• {signedCount}/{scopePersons.length || '–'}</span></h2>
            {!scopePersons.length && (
              <div className="empty">
                <div className="emptyart">📁</div>
                <p className="mut">{isAdmin ? 'Upload file dulu.' : 'Belum ada laporan untukmu.'}</p>
                {isAdmin && <button className="primary" onClick={() => setMtab('file')}>Upload →</button>}
              </div>
            )}
            {!!scopePersons.length && nextUp && (
              <button className="nextup" onClick={() => handleEditSig(nextUp.key, nextUp.name)}>
                <span className="nu-text">Berikutnya: <b>{nextUp.name}</b> <span className="mut">({nextUp.count}×)</span></span>
                <span className="nu-go">Isi →</span>
              </button>
            )}
            {!!scopePersons.length && !nextUp && (
              <button className="nextup done" onClick={() => setMtab('preview')}>
                <span className="nu-text">✅ Semua lengkap — siap export</span>
                <span className="nu-go">→</span>
              </button>
            )}
            {!!scopePersons.length && (
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
              const sig = sigOf(p.name)
              const done = Boolean(sig?.dataUrl)
              return (
                <button key={p.key} className={`person${done ? ' done' : ''}`} onClick={() => handleEditSig(p.key, p.name)} title={done ? 'Klik untuk ubah' : 'Klik untuk isi'}>
                  <span className="pthumb">{sig?.dataUrl ? <img src={sig.dataUrl} alt="" /> : '＋'}</span>
                  <span className="ptext">
                    <b>{p.name}</b>
                    <small>{p.count}× • {p.files.size} file • {done ? 'sudah' : 'belum'}</small>
                  </span>
                  <span className="pgo">{done ? 'Ubah' : 'Isi'}</span>
                </button>
              )
            })}
            {!!scopePersons.length && !filtered.length && <div className="mut">Tidak ketemu. <button className="linkbtn" onClick={() => { setQ(''); setFilter('semua') }}>Reset</button></div>}
          </div>
        </div>

        {/* LANGKAH 3 — PERIKSA & EXPORT (export khusus admin) */}
        <div className="pane pane-preview">
          {isAdmin && (
          <div className="card export">
            <div className="expgrid">
              <div className="expinfo">
                <b>Export</b>
                <span className="mut">{files.length} file • {signedCount}/{scopePersons.length} TTD{remaining > 0 && scopePersons.length > 0 ? ` • ${remaining} belum` : ''}</span>
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
          )}

          <div className="card previewbar" style={{ marginTop: 12 }}>
            <div className="pvrow">
              <b className="pvtitle">{activeFile ? shortName(activeFile.name) : 'Preview'}</b>
              <button className="iconbtn" onClick={() => checkUpdates(true)} title="Perbarui TTD dari cloud" aria-label="Perbarui dari cloud">⟳</button>
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
                      title={`${f.month ? `${f.month} • ` : ''}${f.name} — ${st.hit}/${st.need}`}
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
                <p className="mut">{isAdmin ? 'Belum ada preview. Upload file di langkah 1.' : 'Belum ada laporan untuk ditampilkan.'}</p>
              </div>
            )}
            {activeFile && view === 'laporan' && (
              <ReportPreview
                sheet={activeFile.sheet}
                aoa={activeFile.aoa}
                blocks={activeFile.blocks}
                signatures={effSigsFor(activeFile)}
                fileUid={activeFile.id}
                spotMissing={spot}
                fit={previewFit}
                onEditSig={(key) => handleEditSig(key, persons.find((p) => p.key === key)?.name ?? key)}
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
                            const sig = sigOf(nm)
                            const on = sig?.dataUrl && sig?.enabled !== false
                            return (
                              <button
                                key={nm}
                                className={`chipbtn chip ${on ? 'signed' : 'miss'}`}
                                onClick={() => { if (handleEditSig(normKey(nm), nm)) setMtab('nama') }}
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
        <span className="mut">{scopePersons.length ? `${signedCount}/${scopePersons.length} TTD` : 'Siap?'}</span>
        <button className={!files.length || nextUp ? 'primary' : 'ok'} onClick={primaryAction}>{ctaLabel}</button>
      </div>

      {toastEl}
      {padEl}
    </>
  )
}
