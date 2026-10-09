import jsPDF from 'jspdf'
import html2canvas from 'html2canvas'

/** Export area preview menjadi PDF portrait A4 multi-halaman (mirip confidential). */
export async function exportPreviewPdf(elementId = 'ttdPreview', filename = 'laporan-signed.pdf') {
  const el = document.getElementById(elementId)
  if (!el) throw new Error('Elemen preview tidak ditemukan')
  const canvas = await html2canvas(el, { scale: 2, useCORS: true, backgroundColor: '#f1f5f9' })
  const img = canvas.toDataURL('image/png')
  const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' })
  const pageW = pdf.internal.pageSize.getWidth()
  const pageH = pdf.internal.pageSize.getHeight()
  const imgW = pageW
  const imgH = (canvas.height * imgW) / canvas.width
  let y = 0
  let first = true
  // potong vertikal per halaman A4
  while (y < imgH) {
    if (!first) pdf.addPage()
    pdf.addImage(img, 'PNG', 0, -y, imgW, imgH)
    y += pageH
    first = false
  }
  pdf.save(filename)
}
