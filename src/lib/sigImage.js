/**
 * Util gambar TTD murni (tanpa DOM) — dipakai SignaturePad & SigCropModal.
 * enhance: ubah foto jadi hitam-putih + background transparan (seperti Confidential).
 */

/** Threshold enhance ke tinta hitam pekat / transparan. Mutasi `data` in-place. */
export function enhanceImageData(data, threshold = 175) {
  let ink = 0
  const n = data.length
  for (let i = 0; i < n; i += 4) {
    const gray = 0.299 * data[i] + 0.587 * data[i + 1] + 0.114 * data[i + 2]
    if (gray > threshold) {
      data[i + 3] = 0 // kertas -> transparan
    } else {
      data[i] = 0; data[i + 1] = 0; data[i + 2] = 0; data[i + 3] = 255
      ink++
    }
  }
  return { inkRatio: n ? ink / (n / 4) : 0 }
}

/** Cari bounding-box piksel bertinta dari channel alpha. */
export function findInkBounds(alpha, w, h, step = 2, minAlpha = 10) {
  let minX = w, minY = h, maxX = -1, maxY = -1
  for (let y = 0; y < h; y += step) {
    for (let x = 0; x < w; x += step) {
      if (alpha[y * w + x] > minAlpha) {
        if (x < minX) minX = x
        if (x > maxX) maxX = x
        if (y < minY) minY = y
        if (y > maxY) maxY = y
      }
    }
  }
  if (maxX < 0) return null
  return { minX, minY, maxX, maxY }
}

/** Crop bounding-box + padding ke dalam batas kanvas. */
export function padBounds(b, w, h, pad = 10) {
  if (!b) return null
  const minX = Math.max(0, b.minX - pad)
  const minY = Math.max(0, b.minY - pad)
  const maxX = Math.min(w - 1, b.maxX + pad)
  const maxY = Math.min(h - 1, b.maxY + pad)
  if (maxX <= minX || maxY <= minY) return null
  return { x: minX, y: minY, w: maxX - minX + 1, h: maxY - minY + 1 }
}
