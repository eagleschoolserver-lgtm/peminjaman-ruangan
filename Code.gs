/**
 * BACKEND PEMINJAMAN RUANGAN - Google Apps Script
 * Tempel di Apps Script yang terhubung ke Google Spreadsheet (Extensions > Apps Script).
 * Jalankan fungsi setup() sekali, lalu Deploy > New deployment > Web app
 * (Execute as: Me, Who has access: Anyone).
 */
const SS = SpreadsheetApp.getActiveSpreadsheet();
const TZ = 'Asia/Jakarta';
const SESI = ['Jam ke-1', 'Jam ke-2', 'Jam ke-3', 'Pulang sekolah'];
// Jam pengajuan (WIB). User hanya bisa mengajukan peminjaman pada rentang ini (admin tidak dibatasi).
const JAM_BUKA = '06:00';
const JAM_TUTUP = '15:00';
function pengajuanBuka() {
  const now = Utilities.formatDate(new Date(), TZ, 'HH:mm');
  return now >= JAM_BUKA && now < JAM_TUTUP;
}

function sh(n) { return SS.getSheetByName(n); }

// Jalankan sekali: membuat sheet + akun admin awal (admin / admin123)
function setup() {
  mk('Users', ['username', 'password', 'nama', 'role', 'wa', 'kelola']);
  const uh = sh('Users');
  if (uh.getRange('E1').getValue() !== 'wa') uh.getRange('E1').setValue('wa'); // untuk sheet lama
  if (uh.getRange('F1').getValue() !== 'kelola') uh.getRange('F1').setValue('kelola');
  mk('Ruang', ['nama']);
  mk('Bookings', ['id', 'tanggal', 'sesi', 'ruang', 'username', 'nama', 'keperluan', 'status', 'catatan', 'dibuat']);
  if (sh('Users').getLastRow() < 2) sh('Users').appendRow(['admin', hash('admin123'), 'Administrator', 'admin', '', '']);
  if (sh('Ruang').getLastRow() < 2) ['Aula', 'Lab Komputer', 'Perpustakaan'].forEach(r => sh('Ruang').appendRow([r]));
}
function mk(n, h) {
  const s = sh(n) || SS.insertSheet(n);
  if (s.getLastRow() === 0) { s.appendRow(h); s.setFrozenRows(1); s.getRange('A:J').setNumberFormat('@'); }
}
function hash(p) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, p)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}
function read(n) {
  const v = sh(n).getDataRange().getValues(); const h = v.shift();
  return v.map((r, i) => {
    const o = { _row: i + 2 };
    h.forEach((k, j) => o[k] = r[j] instanceof Date ? Utilities.formatDate(r[j], TZ, 'yyyy-MM-dd') : String(r[j]));
    return o;
  });
}
const pub = b => { const o = Object.assign({}, b); delete o._row; return o; };

function doGet() { return ContentService.createTextOutput('API peminjaman ruangan aktif'); }

function doPost(e) {
  let r;
  try { r = route(JSON.parse(e.postData.contents)); }
  catch (err) { r = { ok: false, error: err.message }; }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}

function route(d) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    ensureCols(); // otomatis memperbaiki header kolom wa & kelola jika setup() belum dijalankan ulang
    if (d.action === 'login') return login(d);
    const raw = CacheService.getScriptCache().get('t_' + d.token);
    if (!raw) return { ok: false, error: 'AUTH' };
    const me = JSON.parse(raw);
    // Admin dengan kolom 'kelola' kosong = Super Admin (semua ruang). Jika terisi = admin ruang tertentu.
    const admin = () => { if (me.role !== 'admin') throw new Error('Hanya admin yang boleh melakukan ini'); };
    const sup = () => { if (me.role !== 'admin' || me.k) throw new Error('Hanya Super Admin yang boleh melakukan ini'); };
    const can = ruang => me.role === 'admin' && (!me.k || me.k.split('|').includes(ruang));
    const today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');

    switch (d.action) {
      case 'rooms': return { ok: true, data: read('Ruang').map(r => r.nama) };
      case 'week': return { ok: true, info: { buka: JAM_BUKA, tutup: JAM_TUTUP, open: pengajuanBuka() }, data: read('Bookings').filter(b => b.tanggal >= d.from && b.tanggal <= d.to && b.status !== 'Ditolak').map(pub) };
      case 'mine': return { ok: true, data: read('Bookings').filter(b => b.username === me.u).reverse().map(pub) };
      case 'book': {
        if (me.role !== 'admin' && !pengajuanBuka())
          throw new Error('Pengajuan hanya dapat dilakukan pukul ' + JAM_BUKA.replace(':', '.') + ' - ' + JAM_TUTUP.replace(':', '.') + ' WIB');
        if (!SESI.includes(d.sesi)) throw new Error('Sesi tidak valid');
        if (!read('Ruang').some(r => r.nama === d.ruang)) throw new Error('Ruang tidak ditemukan');
        const dow = Number(Utilities.formatDate(new Date(d.tanggal + 'T12:00:00+07:00'), TZ, 'u')); // 1=Senin..7=Minggu
        if (dow > 5) throw new Error('Peminjaman hanya Senin - Jumat');
        if (d.tanggal < today) throw new Error('Tanggal sudah lewat');
        if (!String(d.keperluan || '').trim()) throw new Error('Keperluan wajib diisi');
        const bentrok = read('Bookings').some(b => b.tanggal === d.tanggal && b.sesi === d.sesi && b.ruang === d.ruang && (b.status === 'Menunggu' || b.status === 'Disetujui'));
        if (bentrok) throw new Error('Slot sudah dipesan orang lain');
        sh('Bookings').appendRow([Utilities.getUuid().slice(0, 8), d.tanggal, d.sesi, d.ruang, me.u, me.nama, String(d.keperluan).trim(), 'Menunggu', '', Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm')]);
        notifAdmins(d.ruang,
          '*Pengajuan peminjaman ruangan baru*\n\nRuang: ' + d.ruang + '\nTanggal: ' + d.tanggal + '\nSesi: ' + d.sesi +
          '\nPemohon: ' + me.nama + '\nKeperluan: ' + String(d.keperluan).trim() + '\n\nMohon diverifikasi melalui website.');
        const meWA = (read('Users').find(u => u.username === me.u) || {}).wa;
        if (meWA) kirimWA(meWA, pesanKonfirmasi({ nama: me.nama, ruang: d.ruang, tanggal: d.tanggal, sesi: d.sesi, keperluan: String(d.keperluan).trim() }));
        return { ok: true };
      }
      case 'cancel': {
        const b = read('Bookings').find(x => x.id === d.id);
        if (!b) throw new Error('Data tidak ditemukan');
        if (!(b.username === me.u && b.status === 'Menunggu') && !can(b.ruang)) throw new Error('Anda tidak berhak menghapus pengajuan ini');
        sh('Bookings').deleteRow(b._row); return { ok: true };
      }
      case 'all': admin(); return { ok: true, data: read('Bookings').filter(b => can(b.ruang)).reverse().map(pub) };
      case 'verify': {
        admin();
        if (!['Disetujui', 'Ditolak'].includes(d.status)) throw new Error('Status tidak valid');
        const b = read('Bookings').find(x => x.id === d.id);
        if (!b) throw new Error('Data tidak ditemukan');
        if (!can(b.ruang)) throw new Error('Anda bukan admin untuk ruang ' + b.ruang);
        sh('Bookings').getRange(b._row, 8, 1, 2).setValues([[d.status, d.catatan || '']]);
        const pem = read('Users').find(u => u.username === b.username);
        if (pem && pem.wa) kirimWA(pem.wa, d.status === 'Disetujui' ? pesanSetuju(b) : pesanTolak(b, d.catatan));
        return { ok: true };
      }
      case 'users': sup(); return { ok: true, data: read('Users').map(u => ({ username: u.username, nama: u.nama, role: u.role, wa: u.wa, kelola: u.kelola || '' })) };
      case 'addUser': {
        sup();
        const un = String(d.username || '').trim().toLowerCase();
        if (!un || !d.nama) throw new Error('Username dan nama wajib diisi');
        const role = d.role === 'admin' ? 'admin' : 'user';
        const kel = role === 'admin' ? (d.kelola || []).join('|') : '';
        const ex = read('Users').find(u => u.username === un);
        if (ex) {
          if (d.password) sh('Users').getRange(ex._row, 2).setValue(hash(d.password));
          sh('Users').getRange(ex._row, 3).setValue(d.nama);
          sh('Users').getRange(ex._row, 4).setValue(role);
          sh('Users').getRange(ex._row, 5).setValue(normWA(d.wa));
          sh('Users').getRange(ex._row, 6).setValue(kel);
        } else {
          if (!d.password) throw new Error('Password wajib diisi');
          sh('Users').appendRow([un, hash(d.password), d.nama, role, normWA(d.wa), kel]);
        }
        return { ok: true };
      }
      case 'delUser': {
        sup();
        if (d.username === me.u) throw new Error('Tidak bisa menghapus akun sendiri');
        const u = read('Users').find(x => x.username === d.username);
        if (u) sh('Users').deleteRow(u._row);
        return { ok: true };
      }
      case 'addRoom': {
        sup();
        const n = String(d.nama || '').trim();
        if (!n) throw new Error('Nama ruang wajib diisi');
        if (read('Ruang').some(r => r.nama === n)) throw new Error('Ruang sudah ada');
        sh('Ruang').appendRow([n]); return { ok: true };
      }
      case 'delRoom': {
        sup();
        const r = read('Ruang').find(x => x.nama === d.nama);
        if (r) sh('Ruang').deleteRow(r._row);
        return { ok: true };
      }
    }
    throw new Error('Aksi tidak dikenal');
  } finally { lock.releaseLock(); }
}

function login(d) {
  const u = read('Users').find(x => x.username === String(d.username || '').trim().toLowerCase());
  if (!u || u.password !== hash(String(d.password || ''))) throw new Error('Username atau password salah');
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('t_' + token, JSON.stringify({ u: u.username, nama: u.nama, role: u.role, k: u.kelola || '' }), 21600);
  return { ok: true, token, user: { username: u.username, nama: u.nama, role: u.role, kelola: u.kelola ? u.kelola.split('|') : [] } };
}

/* ================= NOTIFIKASI WHATSAPP (FONNTE) ================= */
// Token disimpan di Project Settings > Script properties, key: FONNTE_TOKEN

// ---- Template pesan WA ke pemohon ----
function hariTgl(t) {
  const H = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  const B = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  const dt = new Date(t + 'T12:00:00+07:00');
  return H[Number(Utilities.formatDate(dt, TZ, 'u')) % 7] + ', ' + Number(t.slice(8, 10)) + ' ' + B[Number(t.slice(5, 7)) - 1] + ' ' + t.slice(0, 4);
}
function detail(b) {
  return '*Ruang:* ' + b.ruang + '\n*Hari/Tanggal:* ' + hariTgl(b.tanggal) + '\n*Sesi:* ' + b.sesi + '\n*Keperluan:* ' + b.keperluan;
}
function pesanKonfirmasi(b) {
  return 'Halo ' + b.nama + ',\n\n*Pengajuan peminjaman ruangan Anda sudah kami terima* dan sedang menunggu verifikasi admin.\n\n' + detail(b) +
    '\n\nAnda akan mendapat notifikasi lagi setelah pengajuan disetujui atau ditolak.';
}
function pesanSetuju(b) {
  return 'Halo ' + b.nama + ',\n\n*Pengajuan peminjaman ruangan Anda DISETUJUI* ✅\n\n' + detail(b) +
    '\n\nSilakan gunakan ruangan sesuai jadwal dan jaga kebersihannya. Terima kasih.';
}
function pesanTolak(b, alasan) {
  return 'Halo ' + b.nama + ',\n\nMohon maaf, *pengajuan peminjaman ruangan Anda DITOLAK* ❌\n\n' + detail(b) +
    '\n*Alasan:* ' + (alasan || '-') + '\n\nAnda dapat mengajukan kembali pada jadwal atau ruangan lain melalui website.';
}

// Admin penerima notifikasi untuk satu ruang: admin ruang tsb; jika tidak ada, Super Admin
function adminsFor(ruang) {
  const ad = read('Users').filter(u => u.role === 'admin' && u.wa);
  const khusus = ad.filter(u => u.kelola && u.kelola.split('|').includes(ruang));
  return khusus.length ? khusus : ad.filter(u => !u.kelola);
}
function notifAdmins(ruang, pesan) { adminsFor(ruang).forEach(a => kirimWA(a.wa, pesan)); }

function ensureCols() {
  const s = sh('Users'), h = s.getRange('A1:F1').getValues()[0];
  if (h[4] !== 'wa') s.getRange('E1').setValue('wa');
  if (h[5] !== 'kelola') s.getRange('F1').setValue('kelola');
}

function normWA(n) { return String(n || '').replace(/[^0-9]/g, ''); }

function kirimWA(nomor, pesan) {
  try {
    const token = PropertiesService.getScriptProperties().getProperty('FONNTE_TOKEN');
    const target = normWA(nomor);
    if (!token || !target) return;
    const res = UrlFetchApp.fetch('https://api.fonnte.com/send', {
      method: 'post',
      headers: { Authorization: token },
      payload: { target: target, message: pesan, countryCode: '62' },
      muteHttpExceptions: true
    });
    console.log('Fonnte: ' + res.getContentText());
  } catch (err) {
    console.error('Gagal kirim WA: ' + err.message); // kegagalan WA tidak menggagalkan peminjaman
  }
}

// Jalankan manual untuk tes: isi nomor Anda (format 08xxxxxxxxxx)
function testWA() {
  kirimWA('08xxxxxxxxxx', 'Tes notifikasi peminjaman ruangan berhasil.');
}
