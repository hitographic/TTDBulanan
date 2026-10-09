import ExcelJS from 'exceljs'
import { normKey } from './excelParser.js'
import { CELL_INSET_PX } from './textGeom.js'

// Ukuran render tetap 2,4 x 2,4 cm.
// ExcelJS `ext` dalam pixel (96 dpi): 2,4 / 2,54 * 96 = ±91 px.
const SIG_W_PX = 91
const SIG_H_PX = 91
const EMU_PER_PX = 9525
const SIG_H_EMU = SIG_W_PX * EMU_PER_PX // 866775 ≈ 2,41 cm
// Tinggi baris area TTD (pt) — 2,4 cm ≈ 68 pt, diberi napas 72 pt
const SIG_ROW_H_PT = 72
const EMU_PER_PT = 12700

const dataUrlExt = (dataUrl) => {
  if (dataUrl.startsWith('data:image/jpeg')) return 'jpeg'
  return 'png'
}

// Lebar kolom (satuan karakter Excel) -> pixel, terkalibrasi ke render
// aktual Excel Mac (Calibri 11). Di Excel Windows posisi bisa meleset
// ±3mm; masih dalam toleransi visual.
const charsToPx = (chars) => Math.round(chars * 8)
function colChars(ws, col1) {
  return ws.getColumn(col1).width || 8.43
}

/**
 * Cari anchor sel + offset EMU untuk posisi X (pixel) dihitung dari awal
 * kolom startCol0 (0-based). Mengembalikan nativeCol/nativeColOff mentah
 * (EMU) — presisi penuh, tidak lewat konversi fraksi ExcelJS yang error
 * untuk kolom custom-width.
 */
function anchorForOffset(ws, startCol0, targetXpx) {
  let acc = 0
  for (let c = startCol0; c < startCol0 + 30; c++) {
    const wpx = charsToPx(colChars(ws, c + 1))
    if (targetXpx < acc + wpx) {
      return { nativeCol: c, nativeColOff: Math.round((targetXpx - acc) * EMU_PER_PX) }
    }
    acc += wpx
  }
  const last = startCol0 + 29
  return { nativeCol: last, nativeColOff: Math.round(charsToPx(colChars(ws, last + 1)) * EMU_PER_PX) }
}

/** Titik tengah tiap nama di atas selnya (px dari awal sel). */
function centersForSlot(slot, totalPx) {
  const n = Math.max(slot.names.length, 1)
  const g = slot.geom
  if (g && g.centers && g.centers.length === slot.names.length) {
    return slot.names.map((_, i) => CELL_INSET_PX + g.centers[i])
  }
  return slot.names.map((_, i) => (totalPx / n) * (i + 0.5))
}

/**
 * Membangun workbook bertanda tangan (tanpa download — bisa diuji di Node).
 * Tiap nama yang punya TTD aktif ditempel gambar 2,4x2,4 cm tepat di atas
 * tengah namanya, di baris kosong persis di atas baris nama.
 */
export async function buildSignedWorkbook({ originalBuffer, blocks, signatures }) {
  const wb = new ExcelJS.Workbook()
  await wb.xlsx.load(originalBuffer)
  const ws = wb.worksheets[0]
  // TTD 2,41 cm di tengah vertikal baris 72 pt (2,54 cm)
  const rowOffY = Math.round((SIG_ROW_H_PT * EMU_PER_PT - SIG_H_EMU) / 2)

  const placeSlot = (sigExcelRow, slot) => {
    let total = 0
    for (let k = 0; k < (slot.spanCols || 1); k++) {
      total += charsToPx(colChars(ws, slot.nameCol + 1 + k))
    }
    const centers = centersForSlot(slot, total)
    slot.names.forEach((name, i) => {
      const sig = signatures[normKey(name)]
      if (!sig?.dataUrl || sig?.enabled === false) return
      const base64 = sig.dataUrl.split(',')[1]
      if (!base64) return
      const imageId = wb.addImage({ base64, extension: dataUrlExt(sig.dataUrl) })
      // Baris signature seluruhnya kosong sehingga toleransi luapan aman;
      // jepit kiri saja agar tak keluar sel awal.
      const x = Math.max(1, centers[i] - SIG_W_PX / 2)
      const { nativeCol, nativeColOff } = anchorForOffset(ws, slot.nameCol, x)
      ws.addImage(imageId, {
        tl: { nativeCol, nativeColOff, nativeRow: sigExcelRow - 1, nativeRowOff: rowOffY },
        ext: { width: SIG_W_PX, height: SIG_H_PX },
        editAs: 'oneCell',
      })
    })
  }

  blocks.forEach((blk) => {
    // Baris kosong persis di atas baris nama (nameRow 0-indexed -> +1 = 1-indexed).
    const sigExcelRow = blk.nameRow
    try {
      ws.getRow(sigExcelRow).height = SIG_ROW_H_PT
    } catch { /* abaikan */ }
    ;(blk.slots || []).forEach((slot) => placeSlot(sigExcelRow, slot))
  })
  return wb
}

/** Export + download XLSX bertanda tangan (dipakai browser). */
export async function exportSignedExcel({ originalBuffer, blocks, signatures, outName }) {
  const wb = await buildSignedWorkbook({ originalBuffer, blocks, signatures })
  const buf = await wb.xlsx.writeBuffer()
  const blob = new Blob([buf], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = outName || 'laporan-signed.xlsx'
  a.click()
  setTimeout(() => URL.revokeObjectURL(url), 2000)
}
