/** Sesi login (NIK + password terverifikasi server; password tak disimpan). */
const SES_KEY = 'lapbul-session-v1'

export const loadSession = () => {
  try {
    const raw = localStorage.getItem(SES_KEY)
    if (!raw) return null
    const o = JSON.parse(raw)
    if (!o || !o.nik) return null
    return {
      nik: String(o.nik),
      name: String(o.name || ''),
      nameInFile: String(o.nameInFile || ''),
      role: o.role === 'admin' ? 'admin' : 'user',
      aliases: Array.isArray(o.aliases) ? o.aliases.map(String) : [],
      ts: Number(o.ts) || 0,
      offline: o.offline === true,
    }
  } catch {
    return null
  }
}

export const saveSession = (s) => {
  try {
    localStorage.setItem(
      SES_KEY,
      JSON.stringify({
        nik: s.nik,
        name: s.name || '',
        nameInFile: s.nameInFile || '',
        role: s.role === 'admin' ? 'admin' : 'user',
        aliases: s.aliases || [],
        ts: Date.now(),
        offline: s.offline === true,
      })
    )
  } catch { /* mode privat: abaikan */ }
}

export const clearSession = () => {
  try {
    localStorage.removeItem(SES_KEY)
  } catch { /* abaikan */ }
}
