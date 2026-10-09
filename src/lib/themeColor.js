// Resolusi warna Excel (theme + tint/shade) ke CSS hex.
// ExcelJS tidak me-resolve warna tema saat read, jadi kita hitung manual.
// Urutan indeks theme di SpreadsheetML: lt1, dk1, lt2, dk2, accent1-6,
// hlink, folHlink (terverifikasi dari file: theme 2 = lt2 "tan" untuk
// band section, theme 7 = accent4 untuk header).

const SCHEME = [
  'FFFFFF', // 0 lt1
  '000000', // 1 dk1
  'EEECE1', // 2 lt2
  '1F497D', // 3 dk2
  '4F81BD', // 4 accent1
  'C0504D', // 5 accent2
  '9BBB59', // 6 accent3
  '8064A2', // 7 accent4
  '4BACC6', // 8 accent5
  'F79646', // 9 accent6
  '0000FF', // 10 hlink
  '800080', // 11 folHlink
]

const hex2rgb = (h) => [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16))
const rgb2hex = ([r, g, b]) =>
  '#' + [r, g, b].map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('').toUpperCase()

// Rumus tint/shade sesuai spec OOXML: tint>0 campur putih, tint<0 (shade) gelapkan.
export function themeToCss(theme, tint = 0) {
  const base = SCHEME[theme] ?? '000000'
  let [r, g, b] = hex2rgb(base)
  if (tint > 0) {
    r = r * (1 - tint) + 255 * tint
    g = g * (1 - tint) + 255 * tint
    b = b * (1 - tint) + 255 * tint
  } else if (tint < 0) {
    r = r * (1 + tint)
    g = g * (1 + tint)
    b = b * (1 + tint)
  }
  return rgb2hex([r, g, b])
}

/** Normalisasi objek warna ExcelJS {argb, theme, tint, indexed} ke CSS atau null. */
export function excelColorToCss(c) {
  if (!c) return null
  try {
    if (c.argb && typeof c.argb === 'string' && /^[0-9A-Fa-f]{8}$/.test(c.argb)) {
      if (c.argb === '00000000') return null // Automatic / tak ada warna
      return '#' + c.argb.slice(2).toUpperCase()
    }
    if (typeof c.theme === 'number') return themeToCss(c.theme, Number(c.tint) || 0)
  } catch { /* abaikan, fallback null */ }
  return null
}
