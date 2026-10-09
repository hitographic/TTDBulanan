/**
 * TTDBulanan — backend Google Sheets + Drive untuk app statis (GitHub Pages).
 *
 * sheets  : database TTD (tab "TTD": key | name | fileId | updatedAt)
 * drive   : file PNG tiap TTD di subfolder "TTD", file laporan di folder utama
 *
 * CARA PAKAI (sekali saja):
 * 1. Buka Google Sheet laporan -> Extensions -> Apps Script
 * 2. Hapus isi editor, tempel seluruh file ini, Save
 * 3. Deploy -> New deployment -> type "Web app"
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 4. Copy Web app URL (/exec) -> tempel di Pengaturan > Cloud pada aplikasi
 */

// === ISI DENGAN ID MILIKMU (sudah terisi dari link yang diberikan) ===
var SHEET_ID = '1Dex3iv4toDo1gfEs2r5-0LrZyeRp3omDdcKaNZdNmYk';
var DRIVE_FOLDER_ID = '1TOQ8S1_QaaMBxfOi2xMYYI4E4wtI5Ctw';
// ============================================================

var TAB = 'TTD';
var SIG_DIR = 'TTD';
var USER_TAB = 'User';

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

function rootFolder() {
  return DriveApp.getFolderById(DRIVE_FOLDER_ID);
}

function sigDir() {
  var r = rootFolder();
  var it = r.getFoldersByName(SIG_DIR);
  return it.hasNext() ? it.next() : r.createFolder(SIG_DIR);
}

function ttdSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(TAB);
  if (!sh) {
    sh = ss.insertSheet(TAB);
    sh.appendRow(['key', 'name', 'fileId', 'updatedAt']);
  }
  return sh;
}

function findRow(sh, key) {
  var v = sh.getDataRange().getValues();
  for (var i = 1; i < v.length; i++) {
    if (String(v[i][0]) === key) {
      return { row: i + 1, name: v[i][1], fileId: v[i][2], updatedAt: Number(v[i][3]) || 0 };
    }
  }
  return null;
}

// Tab "User": NIK | Name | Name in File | password | role? | alias?
// - role: "admin" bila selnya berisi admin (default "user")
// - alias: varian ejaan lain dipisah ";" (opsional, mis. "Satria Wijaya K.;Satria W K")
function parseUsers() {
  var sh = SpreadsheetApp.openById(SHEET_ID).getSheetByName(USER_TAB);
  if (!sh) throw new Error('Tab "User" tidak ditemukan di Sheet.');
  var v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  var head = v[0].map(function (h) { return String(h || '').toLowerCase().trim(); });
  var ci = function (names) {
    for (var k = 0; k < names.length; k++) {
      var i = head.indexOf(names[k]);
      if (i >= 0) return i;
    }
    return -1;
  };
  var cNik = ci(['nik']);
  var cName = ci(['name', 'nama']);
  var cNif = ci(['name in file', 'name_in_file', 'nameinfile']);
  var cPw = ci(['password', 'pass']);
  var cRole = ci(['role', 'peran']);
  var cAlias = ci(['alias']);
  var out = [];
  for (var i = 1; i < v.length; i++) {
    var nik = String(v[i][cNik] || '').trim();
    if (!nik) continue;
    var aliases = [];
    if (cAlias >= 0) {
      aliases = String(v[i][cAlias] || '').split(';')
        .map(function (s) { return String(s || '').replace(/\s+/g, ' ').trim(); })
        .filter(function (s) { return !!s; });
    }
    out.push({
      nik: nik,
      name: cName >= 0 ? String(v[i][cName] || '').trim() : '',
      nameInFile: cNif >= 0 ? String(v[i][cNif] || '').trim() : '',
      password: cPw >= 0 ? String(v[i][cPw] || '') : '',
      role: (cRole >= 0 && /admin/i.test(String(v[i][cRole] || ''))) ? 'admin' : 'user',
      aliases: aliases
    });
  }
  return out;
}

// Versi publik user — password TIDAK pernah dikirim ke client.
function publicUser(u) {
  return { nik: u.nik, name: u.name, nameInFile: u.nameInFile, role: u.role, aliases: u.aliases };
}

function fileDataUrl(fileId) {
  var f = DriveApp.getFileById(fileId);
  var b = f.getBlob();
  return 'data:' + (b.getContentType() || 'image/png') + ';base64,' +
    Utilities.base64Encode(b.getBytes());
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var a = p.action || 'ping';
  try {
    if (a === 'ping') {
      return json({ ok: true, service: 'TTDBulanan', time: new Date().toISOString() });
    }
    if (a === 'users') {
      // Daftar user TANPA password (untuk sinkronisasi nama & alias)
      var users = parseUsers().map(publicUser);
      return json({ ok: true, users: users });
    }
    if (a === 'login') {
      // Verifikasi NIK + password di server; password tak pernah dikirim balik
      var nik = String(p.nik || '').trim();
      var pw = String(p.password || '');
      var all = parseUsers();
      var found = null;
      for (var u = 0; u < all.length; u++) {
        if (all[u].nik === nik) { found = all[u]; break; }
      }
      if (!found || found.password !== pw) {
        return json({ ok: false, error: 'NIK atau password salah.' });
      }
      return json({ ok: true, user: publicUser(found) });
    }
    if (a === 'pull') {
      // Semua TTD + gambarnya (last-write-wins diurutkan client via updatedAt)
      var sh = ttdSheet();
      var v = sh.getDataRange().getValues();
      var out = [];
      for (var i = 1; i < v.length; i++) {
        var key = String(v[i][0] || '');
        if (!key) continue;
        var dataUrl = '';
        try { dataUrl = fileDataUrl(String(v[i][2] || '')); } catch (err) { /* file hilang: lewati */ }
        out.push({ key: key, name: String(v[i][1] || key), dataUrl: dataUrl, updatedAt: Number(v[i][3]) || 0 });
      }
      return json({ ok: true, signatures: out });
    }
    if (a === 'files') {
      // File Excel di folder utama + tiap subfolder bulan (mis. "Oktober").
      // Folder sistem "TTD" (gambar tanda tangan) dikecualikan dari bulan.
      var root = rootFolder();
      var folders = [];
      var fit = root.getFolders();
      while (fit.hasNext()) {
        var fd = fit.next();
        if (fd.getName() === SIG_DIR) continue;
        folders.push({ id: fd.getId(), name: fd.getName() });
      }
      folders.sort(function (x, y) { return x.name < y.name ? -1 : 1; });
      var files = [];
      var pushXls = function (folder, month, monthId) {
        var it = folder.getFiles();
        while (it.hasNext()) {
          var f = it.next();
          if (!/\.xlsx?$/i.test(f.getName())) continue;
          files.push({
            id: f.getId(), name: f.getName(),
            month: month, monthId: monthId,
            size: f.getSize(), modified: f.getLastUpdated().toISOString()
          });
        }
      };
      pushXls(root, '', '');
      for (var m = 0; m < folders.length; m++) {
        try {
          pushXls(DriveApp.getFolderById(folders[m].id), folders[m].name, folders[m].id);
        } catch (err) { /* lewati folder bermasalah */ }
      }
      files.sort(function (x, y) {
        if (x.month !== y.month) return x.month < y.month ? -1 : 1;
        return x.name < y.name ? -1 : 1;
      });
      return json({ ok: true, months: folders, files: files });
    }
    if (a === 'get') {
      var g = DriveApp.getFileById(String(p.fileId || ''));
      return json({ ok: true, name: g.getName(), mime: g.getMimeType(), base64: Utilities.base64Encode(g.getBlob().getBytes()) });
    }
    return json({ ok: false, error: 'action tak dikenal: ' + a });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.action === 'push') {
      // Simpan 1 TTD: gambar -> Drive, index -> Sheet (yang lebih baru menang)
      var key = String(body.key || '').trim();
      if (!key) return json({ ok: false, error: 'key kosong' });
      var updatedAt = Number(body.updatedAt) || Date.now();
      var sh = ttdSheet();
      var ex = findRow(sh, key);
      if (ex && ex.updatedAt > updatedAt) {
        return json({ ok: true, skipped: true, reason: 'cloud lebih baru' });
      }
      var bytes = Utilities.base64Decode(String(body.dataUrl).split(',')[1]);
      var blob = Utilities.newBlob(bytes, 'image/png', key + '.png');
      var fileId = '';
      if (ex && ex.fileId) {
        try {
          var old = DriveApp.getFileById(ex.fileId);
          old.setContent(bytes);
          fileId = old.getId();
        } catch (err) {
          fileId = sigDir().createFile(blob).getId();
        }
      } else {
        fileId = sigDir().createFile(blob).getId();
      }
      var row = [key, String(body.name || key), fileId, updatedAt];
      if (ex) sh.getRange(ex.row, 1, 1, 4).setValues([row]);
      else sh.appendRow(row);
      return json({ ok: true, fileId: fileId });
    }
    if (body.action === 'save') {
      // Simpan file laporan (asli / bertanda tangan) ke Drive.
      // folderId opsional: simpan ke folder bulan yang sama (default folder utama).
      var name = String(body.name || 'laporan-signed.xlsx');
      var data = Utilities.base64Decode(String(body.base64));
      var mime = body.mime || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
      var target = rootFolder();
      if (body.folderId) {
        try { target = DriveApp.getFolderById(String(body.folderId)); } catch (err) { /* fallback root */ }
      }
      var f = target.createFile(Utilities.newBlob(data, mime, name));
      return json({ ok: true, id: f.getId(), name: f.getName() });
    }
    return json({ ok: false, error: 'action tak dikenal' });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}
