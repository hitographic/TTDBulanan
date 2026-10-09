import JSZip from 'jszip'
import { buildSignedWorkbook } from './excelExport.js'

/**
 * Gabungkan SEMUA file bertanda tangan ke dalam 1 file ZIP.
 * Jauh lebih ramah user dibanding N kali download yang sering diblokir browser/HP.
 */
export async function exportAllZip({ files, signatures, outName, onProgress }) {
  if (!files?.length) throw new Error('Belum ada file.')
  const zip = new JSZip()
  for (let i = 0; i < files.length; i++) {
    const f = files[i]
    onProgress?.(i + 1, files.length, f.name)
    const wb = await buildSignedWorkbook({
      originalBuffer: f.buffer,
      blocks: f.blocks,
      signatures,
    })
    const buf = await wb.xlsx.writeBuffer()
    const entryName = f.name.replace(/\.xlsx?$/i, '') + '-signed.xlsx'
    zip.file(entryName, buf)
  }
  const blob = await zip.generateAsync(
    { type: 'blob', compression: 'DEFLATE' },
    ({ percent }) => onProgress?.(percent, 100, 'Mengompres ZIP…')
  )
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = outName || `Laporan-Bulanan-signed-${files.length}-file.zip`
  document.body.appendChild(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 5000)
}
