import { useState } from 'react'

/** Halaman login NIK + password (terverifikasi ke Google Sheet via cloud). */
export default function LoginPage({ onLogin }) {
  const [nik, setNik] = useState('')
  const [pw, setPw] = useState('')
  const [show, setShow] = useState(false)
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
      <div className="loginbg" aria-hidden />
      <form className="card logincard" onSubmit={submit}>
        <div className="loginlogo" aria-hidden><span>✍️</span></div>
        <h1>TTD Laporan Bulanan</h1>
        <p className="loginsub">Tanda tangan digital laporan bulanan</p>
        {err && <div className="formerr" role="alert">{err}</div>}
        <label className="llab">
          <span>NIK</span>
          <input
            type="text" inputMode="numeric" autoComplete="username"
            value={nik} onChange={(e) => setNik(e.target.value)}
            placeholder="cth. 50086913"
          />
        </label>
        <label className="llab">
          <span>Password</span>
          <div className="pwrow">
            <input
              type={show ? 'text' : 'password'} autoComplete="current-password"
              value={pw} onChange={(e) => setPw(e.target.value)}
              placeholder="••••••••"
            />
            <button
              type="button" className="pweye"
              onClick={() => setShow((s) => !s)}
              aria-label={show ? 'Sembunyikan password' : 'Tampilkan password'}
              title={show ? 'Sembunyikan' : 'Tampilkan'}
            >
              {show ? '🙈' : '👁️'}
            </button>
          </div>
        </label>
        <button className="loginbtn" disabled={loading}>
          {loading && <span className="spin" aria-hidden />}
          {loading ? 'Memeriksa…' : 'Masuk →'}
        </button>
        <p className="loginfoot">TTDBULANAN</p>
      </form>
    </div>
  )
}
