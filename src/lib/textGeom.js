/**
 * Geometri teks nama untuk penempatan TTD tepat di atas tiap nama
 * (bukan asal tengah slot — nama panjangnya beda-beda, mis. "Afida Ayu A."
 * vs "Hero Wiwoho", jadi tengah slot jatuh di atas pemisah "/").
 *
 * Sel diekspor apa adanya (Calibri 11, rata kiri), jadi posisi tengah tiap
 * nama = inset sel + tengah segmen teks hasil split '/'.
 */

export function segmentCenters(raw, measure) {
  const parts = String(raw ?? '').split('/')
  if (!parts.length) return null
  const out = []
  let x = 0
  for (let i = 0; i < parts.length; i++) {
    const part = parts[i]
    const core = part.trim()
    if (!core) return null
    const lead = part.length - part.trimStart().length
    const leadW = measure(part.slice(0, lead))
    const coreW = measure(core)
    out.push(x + leadW + coreW / 2)
    x += measure(part) + (i < parts.length - 1 ? measure('/') : 0)
  }
  return { centers: out, total: x }
}

// Jarak teks dari tepi kiri sel di Excel (cell padding), px
export const CELL_INSET_PX = 3

let _ctx = null
/** Measurer Calibri 11 (font sel nama laporan). Hanya di browser. */
export function browserMeasure() {
  if (!_ctx) {
    _ctx = document.createElement('canvas').getContext('2d')
    _ctx.font = '11pt Calibri, Arial, sans-serif'
  }
  const ctx = _ctx
  return (s) => ctx.measureText(s).width
}

/** Hitung geometri tiap sisi blok; mengembalikan null jika tak bisa (fallback: tengah slot). */
export function geomForSide(raw, names, measure) {
  if (!names.length) return null
  try {
    const g = segmentCenters(raw, measure)
    if (!g || g.centers.length !== names.length) return null
    return g
  } catch {
    return null
  }
}
