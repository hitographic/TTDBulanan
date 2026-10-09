import * as XLSX from 'xlsx'

/**
 * Kunci unik nama — titik diabaikan agar varian ejaan seperti
 * "Mukharom JS." dan "Mukharom J.S." memakai 1 TTD yang sama.
 */
export const normKey = (s) =>
  String(s || '').replace(/\./g, '').replace(/\s+/g, ' ').trim().toLowerCase()

/** Satu sel bisa berisi beberapa nama dipisah "/". */
export const splitNames = (v) =>
  String(v || '')
    .split('/')
    .map((s) => s.replace(/\s+/g, ' ').trim())
    .filter(Boolean)

/**
 * Samakan ejaan nama vs akun user.
 * - Cocok persis (normKey, toleran titik): "Mukharom J.S." = "Mukharom JS."
 * - Cocok inisial per token: "A. Wahid A.W." = "Abdul Wahid A.W.",
 *   "Satria W.K." = "Satria Wijaya K.", "Zaidhiya R.R." = "Zaidhiya Rizqi R."
 * - Varian lain (mis. "Zaidhiya R.") didaftarkan di kolom alias pada tab User.
 */
const toks = (s) => String(s || '').toLowerCase().split(/[\s.]+/).filter(Boolean)
const tokEq = (a, b) =>
  a === b || (a.length === 1 && b.startsWith(a)) || (b.length === 1 && a.startsWith(b))

export const fuzzyNameMatch = (a, b) => {
  const A = toks(a), B = toks(b)
  return A.length > 0 && A.length === B.length && A.every((t, i) => tokEq(t, B[i]))
}

/** Nama di file milik user? user = {nameInFile, aliases: []}. */
export const matchNameToUser = (fileName, user) => {
  if (!user) return false
  const cands = [user.nameInFile, ...(user.aliases || [])].filter(Boolean)
  return cands.some(
    (c) => normKey(fileName) === normKey(c) || fuzzyNameMatch(fileName, c)
  )
}

const clean = (v) => String(v ?? '').replace(/\s+/g, ' ').trim().toLowerCase()

// Label peran di baris footer. Urutan penting: yang spesifik dulu.
const LABELS = [
  [/diketahui|mengetahui/, 'Diketahui'],
  [/disetujui|menyetujui/, 'Disetujui'],
  [/diperiksa|pemeriksa/, 'Diperiksa'],
  [/dibuat/, 'Dibuat'],
  [/auditor/, 'Auditor'],
  [/auditee/, 'Auditee'],
]

const roleOf = (text) => {
  const t = clean(text)
  for (const [re, role] of LABELS) {
    if (re.test(t)) return role
  }
  return null
}

// Baris jabatan di bawah baris nama (Staff, Spv, BPDQCM, ...) — bukan nama.
const ROLE_WORDS =
  /staff|spv|supervisor|qc\b|bpdqcm|manager|mgr|admin|operator|leader|foreman|kepala|dept|section|bagian|unit|koordinator|analis|teknisi|technician/i
const DATE_CITY =
  /cibitung|jakarta|bekasi|\b20\d\d\b|januari|februari|maret|april|mei\b|juni\b|juli\b|agustus|september|oktober|november|desember/i

/** Sel layak dianggap nama orang. */
const isNameLike = (v) => {
  const t = clean(v)
  if (!/[a-z]/.test(t)) return false // butuh huruf (menyingkirkan ".", "-", angka)
  if (roleOf(v)) return false
  if (DATE_CITY.test(t)) return false
  return t.replace(/[^a-z]/g, '').length >= 2
}

/**
 * Deteksi blok footer TTD generik.
 *
 * Pola umum laporan: satu baris label ("Dibuat oleh", "Diketahui oleh",
 * "Disetujui oleh", ...), 2–3 baris kosong, baris nama, baris jabatan.
 * Tiap sel nama jadi 1 slot; perannya mengikuti label yang menaunginya
 * (lebar merge label, mis. "Diketahui oleh," bisa menaungi 3 nama).
 * Sel berisi "A / B" dipecah jadi beberapa nama (1 TTD per orang).
 */
export function parseWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' })
  const sheetName = wb.SheetNames[0]
  const ws = wb.Sheets[sheetName]
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: '' })
  const merges = ws['!merges'] || []

  // aoa dimulai dari sudut !ref (tak selalu A1) — selaraskan ke koordinat
  // Excel 0-indexed agar cocok dengan !merges, ExcelJS, & model preview.
  let rowOff = 0, colOff = 0
  try {
    const rng = XLSX.utils.decode_range(ws['!ref'])
    rowOff = rng.s.r
    colOff = rng.s.c
  } catch { /* sheet kosong */ }
  const T = (r) => r + rowOff // indeks aoa -> baris Excel 0-indexed
  const C = (c) => c + colOff // indeks aoa -> kolom Excel 0-indexed

  const mergeAt = (r, c) => merges.find((m) => r >= m.s.r && r <= m.e.r && c >= m.s.c && c <= m.e.c)

  const blocks = []
  for (let r = 0; r < aoa.length; r++) {
    const row = aoa[r] || []
    const labels = [] // [{c, role, c0, c1}] — c/c0/c1 koordinat Excel 0-indexed
    for (let c = 0; c < row.length; c++) {
      const role = roleOf(row[c])
      if (!role) continue
      const m = mergeAt(T(r), C(c))
      labels.push({ c: C(c), role, c0: m ? m.s.c : C(c), c1: m ? m.e.c : C(c) })
    }
    if (!labels.length) continue

    // Baris nama = baris tak-kosong pertama setelah label (maks. 6 baris di bawah).
    let nameRow = -1
    for (let nr = r + 1; nr <= Math.min(r + 6, aoa.length - 1); nr++) {
      const nrow = aoa[nr] || []
      const cands = []
      for (let c = 0; c < nrow.length; c++) {
        if (isNameLike(nrow[c])) cands.push(c)
      }
      if (!cands.length) continue
      // Kalau semua kandidat mirip jabatan (tanpa nama), lewati baris ini.
      const raws = cands.map((c) => String(nrow[c] ?? ''))
      if (raws.every((t) => ROLE_WORDS.test(t))) continue
      nameRow = nr
      break
    }
    if (nameRow < 0) continue

    const minLabelCol = Math.min(...labels.map((l) => l.c))
    const maxLabelEnd = Math.max(...labels.map((l) => l.c1))
    const nrow = aoa[nameRow] || []
    const slots = []
    for (let c = 0; c < nrow.length; c++) {
      const raw = String(nrow[c] ?? '')
      if (!isNameLike(raw)) continue
      const tc = C(c)
      if (tc < minLabelCol || tc > maxLabelEnd) continue
      // Peran: label yang merge-nya menaungi kolom nama, atau label terdekat di kiri.
      let role = null
      for (const l of labels) {
        if (tc >= l.c0 && tc <= l.c1) { role = l.role; break }
      }
      if (!role) {
        const left = labels.filter((l) => l.c0 <= tc)
        role = (left.length ? left[left.length - 1] : labels[0]).role
      }
      const m = mergeAt(T(nameRow), tc)
      const names = splitNames(raw)
      if (!names.length) continue
      slots.push({
        role,
        labelCol: labels[0].c,
        nameCol: tc, // 0-indexed (koordinat Excel)
        spanCols: m ? m.e.c - m.s.c + 1 : 1,
        raw: raw.replace(/\s+/g, ' ').trim(),
        names,
        geom: null, // diisi browserMeasure di App (withGeom)
      })
    }
    if (!slots.length) continue
    blocks.push({
      id: blocks.length,
      page: blocks.length + 1,
      labelRow: T(r), // 0-indexed (koordinat Excel)
      nameRow: T(nameRow), // 0-indexed (koordinat Excel)
      slots,
    })
  }

  // Daftar orang unik + hitung kemunculan
  const map = new Map()
  blocks.forEach((b) => {
    b.slots.forEach((s) => {
      s.names.forEach((name) => {
        const key = normKey(name)
        if (!map.has(key)) map.set(key, { key, name, count: 0, occ: [] })
        const p = map.get(key)
        if (name.length > p.name.length) p.name = name // ejaan paling lengkap
        p.count += 1
        p.occ.push({ blockId: b.id, page: b.page, side: s.role })
      })
    })
  })
  const persons = [...map.values()].sort((a, b) => b.count - a.count || a.name.localeCompare(b.name))
  return { sheetName, aoa, blocks, persons }
}
