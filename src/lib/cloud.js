/**
 * Client backend cloud (Google Sheets + Drive via Apps Script Web App).
 * Tanpa URL yang dikonfigurasi, app tetap jalan 100% lokal (localStorage).
 */

import { normKey, fuzzyNameMatch } from './excelParser.js'

const LS_KEY = 'lapbul-cloud-v1'

// URL bawaan agar tim langsung tersambung (bisa diganti per-HP di kartu Cloud).
export const DEFAULT_CLOUD_URL =
  'https://script.google.com/macros/s/AKfycbxmKvI7dV7aUQCohyI2YyYHJbdfVwUNgpQyKEHCtQUAZktcwLai9UnUZ8YHTPOmLAEgtA/exec'

export const loadCloud = () => {
  try {
    const raw = localStorage.getItem(LS_KEY)
    if (!raw) return { url: DEFAULT_CLOUD_URL, autoSync: true }
    const o = JSON.parse(raw)
    return { url: String(o.url || '').trim() || DEFAULT_CLOUD_URL, autoSync: o.autoSync !== false }
  } catch {
    return { url: DEFAULT_CLOUD_URL, autoSync: true }
  }
}

export const saveCloud = (c) => {
  try {
    localStorage.setItem(LS_KEY, JSON.stringify({ url: cleanUrl(c.url), autoSync: c.autoSync !== false }))
  } catch { /* mode privat: abaikan */ }
}

const cleanUrl = (url) => String(url || '').trim().replace(/\/+$/, '')
const needUrl = (url) => {
  const u = cleanUrl(url)
  if (!u) throw new Error('URL Apps Script belum diisi (Pengaturan Cloud).')
  return u
}

async function get(url, params) {
  const u = needUrl(url)
  let r
  try {
    r = await fetch(`${u}?${new URLSearchParams(params)}`)
  } catch {
    throw new Error('Tak bisa menghubungi cloud — periksa URL & internet.')
  }
  if (!r.ok) throw new Error(`Cloud menjawab ${r.status}.`)
  const j = await r.json()
  if (!j || j.ok !== true) throw new Error(j?.error || 'Cloud mengembalikan error.')
  return j
}

async function post(url, body) {
  const u = needUrl(url)
  let r
  try {
    // text/plain agar jadi "simple request" (tanpa preflight CORS)
    r = await fetch(u, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(body),
    })
  } catch {
    throw new Error('Tak bisa menghubungi cloud — periksa URL & internet.')
  }
  if (!r.ok) throw new Error(`Cloud menjawab ${r.status}.`)
  const j = await r.json()
  if (!j || j.ok !== true) throw new Error(j?.error || 'Cloud mengembalikan error.')
  return j
}

export const cloudPing = (url) => get(url, { action: 'ping' })
export const cloudPull = (url) => get(url, { action: 'pull' }) // -> {signatures:[{key,name,dataUrl,updatedAt}]}
export const cloudFiles = (url) => get(url, { action: 'files' }) // -> {months, files:[{id,name,month,monthId,size,modified}]}
export const cloudGet = (url, fileId) => get(url, { action: 'get', fileId }) // -> {name,mime,base64}
export const cloudUsers = (url) => get(url, { action: 'users' }) // -> {users:[{nik,name,nameInFile,role,aliases}]} (tanpa password)
export const cloudLogin = (url, nik, password) =>
  get(url, { action: 'login', nik, password }) // -> {user:{nik,name,nameInFile,role,aliases}}
export const cloudPush = (url, { key, name, dataUrl, updatedAt }) =>
  post(url, { action: 'push', key, name, dataUrl, updatedAt })
export const cloudSave = (url, { name, base64, mime, folderId }) =>
  post(url, { action: 'save', name, base64, mime, folderId: folderId || '' })

/**
 * Gabung TTD lokal + cloud. Yang lebih baru (updatedAt) menang.
 * Mengembalikan {merged, added, updated} untuk notifikasi.
 */
export function mergeSignatures(local, remote) {
  const merged = { ...local }
  let added = 0, updated = 0
  ;(remote || []).forEach((r) => {
    if (!r || !r.key || !r.dataUrl) return
    const cur = merged[r.key]
    const rt = Number(r.updatedAt) || 0
    const lt = Number(cur?.updatedAt) || 0
    if (!cur) {
      merged[r.key] = { dataUrl: r.dataUrl, enabled: true, updatedAt: rt }
      added += 1
    } else if (rt > lt) {
      merged[r.key] = { dataUrl: r.dataUrl, enabled: cur.enabled !== false, updatedAt: rt }
      updated += 1
    }
  })
  return { merged, added, updated }
}

/** ArrayBuffer -> base64 (chunked agar tak overflow stack untuk file besar). */
export function bufToB64(ab) {
  const bytes = new Uint8Array(ab)
  let s = ''
  const CH = 0x8000
  for (let i = 0; i < bytes.length; i += CH) {
    s += String.fromCharCode.apply(null, bytes.subarray(i, i + CH))
  }
  return btoa(s)
}

/** base64 -> ArrayBuffer. */
export function b64ToBuf(b64) {
  const s = atob(b64)
  const bytes = new Uint8Array(s.length)
  for (let i = 0; i < s.length; i++) bytes[i] = s.charCodeAt(i)
  return bytes.buffer
}

/**
 * Ambil TTD efektif untuk sebuah nama di file.
 * Urutan: kunci persis -> alias terdaftar -> inisial milik sendiri.
 * - aliasToKey: Map normKey(varian) -> kunci kanonis pemilik
 * - scopeKeys: [{raw, key}] nama-nama milik akun login (untuk fuzzy inisial)
 */
export function pickSig(signatures, aliasToKey, name, scopeKeys = []) {
  const k = normKey(name)
  if (signatures[k]?.dataUrl) return signatures[k]
  const canon = aliasToKey?.get?.(k)
  if (canon && signatures[canon]?.dataUrl) return signatures[canon]
  for (const { raw, key } of scopeKeys) {
    if (signatures[key]?.dataUrl && fuzzyNameMatch(name, raw)) return signatures[key]
  }
  return null
}
