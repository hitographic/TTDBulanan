import { useState } from 'react'

/** Halaman login NIK + password (terverifikasi ke Google Sheet via cloud). */
export default function LoginPage({ onLogin }) {
  const [nik, setNik] = useState('')
  const [pw, setPw] = useState('')
  const [err, setErr] = useState('')
  const [loading, setLoading] = useState(false)

  const submit = async (e) => {
    e.preventDefault()
    if (!nik.trim() || !pw) {
      setErr('Isi NIK dan password dulu.')
      return
    }
    setErr('')
    setLoading(true)
    try {
      await onLogin(nik.trim(), pw)
    } catch (ex) {
      setErr(ex?.message || 'Gagal masuk.')
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="loginwrap">
      <form className="card logincard" onSubmit={submit}>
        <div className="loginlogo" aria-hidden>✍️</div>
        <h1>TTD Laporan Bulanan</h1>
        <p className="mut">Masuk dengan NIK & password.</p>
        {err && <div className="formerr" role="alert">{err}</div>}
        <label className="llab">
          NIK
          <input
            type="text" inputMode="numeric" autoComplete="username"
            value={nik} onChange={(e) => setNik(e.target.value)}
            placeholder="cth. 50086913"
          />
        </label>
        <label className="llab">
          Password
          <input
            type="password" autoComplete="current-password"
            value={pw} onChange={(e) => setPw(e.target.value)}
            placeholder="••••••••"
          />
        </label>
        <button className="primary big" style={{ width: '100%', marginTop: 10 }} disabled={loading}>
          {loading ? 'Memeriksa…' : 'Masuk'}
        </button>
      </form>
    </div>
  )
}
