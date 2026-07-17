/**
 * MINI GAME BALAP KARUNG - LIGHT VERSION
 * Satu file Code.gs: game member + admin panel + Spreadsheet.
 *
 * 1) Ganti SPREADSHEET_ID dan ADMIN_PIN.
 * 2) Jalankan setupMiniGame() satu kali.
 * 3) Deploy sebagai Web App:
 *    Execute as: Me
 *    Who has access: Anyone
 *
 * URL Member : .../exec
 * URL Admin  : .../exec?page=admin
 */

const SPREADSHEET_ID = 'GANTI_DENGAN_ID_SPREADSHEET';
const ADMIN_PIN = '1708';

const MEMBER_SHEET = 'MEMBERS';
const LOG_SHEET = 'GAME_LOGS';
const SETTINGS_KEY = 'SACK_RACE_SETTINGS';

function doGet(e) {
  const page = e && e.parameter && e.parameter.page === 'admin' ? 'admin' : 'member';
  return HtmlService.createHtmlOutput(buildHtml_(page))
    .setTitle(page === 'admin' ? 'Admin Balap Karung' : 'Mini Game Balap Karung')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1');
}

/* =========================
   SETUP SPREADSHEET
========================= */

function setupMiniGame() {
  const ss = getSpreadsheet_();

  const member = getOrCreateSheet_(ss, MEMBER_SHEET);
  if (member.getLastRow() === 0) {
    member.appendRow([
      'USER_ID',
      'TICKETS_AVAILABLE',
      'TOTAL_PLAYED',
      'TOTAL_PRIZE',
      'LAST_PLAYED',
      'CREATED_AT',
      'STATUS'
    ]);
    member.setFrozenRows(1);
  }

  const logs = getOrCreateSheet_(ss, LOG_SHEET);
  if (logs.getLastRow() === 0) {
    logs.appendRow([
      'PLAY_ID',
      'USER_ID',
      'STARTED_AT',
      'FINISHED_AT',
      'PRIZE',
      'TICKET_USED',
      'TICKETS_LEFT',
      'TOTAL_PLAYED',
      'CHOSEN_PLAYER',
      'WINNER',
      'STATUS'
    ]);
    logs.setFrozenRows(1);
  }

  const settings = getSettings_();
  saveSettings_(settings);

  member.autoResizeColumns(1, 7);
  logs.autoResizeColumns(1, 11);

  return 'Setup selesai. Sheet MEMBERS dan GAME_LOGS sudah tersedia.';
}

function getSpreadsheet_() {
  if (SPREADSHEET_ID && SPREADSHEET_ID !== 'GANTI_DENGAN_ID_SPREADSHEET') {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (!active) {
    throw new Error('Isi SPREADSHEET_ID terlebih dahulu.');
  }
  return active;
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function getSettings_() {
  const raw = PropertiesService.getScriptProperties().getProperty(SETTINGS_KEY);
  const defaults = {
    enabled: true,
    duration: 12,
    prize: 10000,
    players: ['Budi Kuat', 'Putri Lompat', 'Jono Cepat', 'Siti Juara']
  };

  if (!raw) return defaults;

  try {
    const value = JSON.parse(raw);
    return {
      enabled: value.enabled !== false,
      duration: Math.max(5, Math.min(30, Number(value.duration) || 12)),
      prize: Math.max(0, Number(value.prize) || 10000),
      players: Array.isArray(value.players) && value.players.length >= 2
        ? value.players.slice(0, 6).map(function(name) {
            return cleanText_(name, 24);
          })
        : defaults.players
    };
  } catch (error) {
    return defaults;
  }
}

function saveSettings_(settings) {
  PropertiesService.getScriptProperties()
    .setProperty(SETTINGS_KEY, JSON.stringify(settings));
}

/* =========================
   MEMBER API
========================= */

function getPublicSettings() {
  const settings = getSettings_();
  return {
    enabled: settings.enabled,
    duration: settings.duration,
    prize: settings.prize,
    prizeLabel: formatRupiah_(settings.prize),
    players: settings.players
  };
}

function memberLogin(userId) {
  userId = normalizeUserId_(userId);
  if (userId.length < 3) throw new Error('User ID minimal 3 karakter.');

  const sheet = getSpreadsheet_().getSheetByName(MEMBER_SHEET);
  if (!sheet) throw new Error('Jalankan setupMiniGame() terlebih dahulu.');

  const row = findMemberRow_(sheet, userId);
  if (!row) throw new Error('User ID belum didaftarkan oleh admin.');

  const values = sheet.getRange(row, 1, 1, 7).getValues()[0];
  if (String(values[6]).toUpperCase() !== 'ACTIVE') {
    throw new Error('User ID sedang tidak aktif.');
  }

  return memberDataFromRow_(values);
}

function beginGame(userId, chosenPlayer) {
  userId = normalizeUserId_(userId);
  chosenPlayer = cleanText_(chosenPlayer, 24);

  const settings = getSettings_();
  if (!settings.enabled) throw new Error('Mini game sedang dinonaktifkan.');
  if (settings.players.indexOf(chosenPlayer) === -1) {
    throw new Error('Peserta tidak valid.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = getSpreadsheet_();
    const memberSheet = ss.getSheetByName(MEMBER_SHEET);
    const logSheet = ss.getSheetByName(LOG_SHEET);

    const row = findMemberRow_(memberSheet, userId);
    if (!row) throw new Error('User ID belum didaftarkan.');

    const member = memberSheet.getRange(row, 1, 1, 7).getValues()[0];
    const tickets = Number(member[1]) || 0;
    const status = String(member[6] || '').toUpperCase();

    if (status !== 'ACTIVE') throw new Error('User ID sedang tidak aktif.');
    if (tickets < 1) throw new Error('Tiket bermain sudah habis.');

    const ticketsLeft = tickets - 1;
    memberSheet.getRange(row, 2).setValue(ticketsLeft);

    const playId = Utilities.getUuid();
    logSheet.appendRow([
      playId,
      userId,
      new Date(),
      '',
      settings.prize,
      1,
      ticketsLeft,
      Number(member[2]) || 0,
      chosenPlayer,
      chosenPlayer,
      'PLAYING'
    ]);

    return {
      playId: playId,
      userId: userId,
      chosenPlayer: chosenPlayer,
      duration: settings.duration,
      prize: settings.prize,
      prizeLabel: formatRupiah_(settings.prize),
      ticketsLeft: ticketsLeft
    };
  } finally {
    lock.releaseLock();
  }
}

function finishGame(playId, userId) {
  playId = cleanText_(playId, 60);
  userId = normalizeUserId_(userId);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const ss = getSpreadsheet_();
    const memberSheet = ss.getSheetByName(MEMBER_SHEET);
    const logSheet = ss.getSheetByName(LOG_SHEET);

    const logRow = findLogRow_(logSheet, playId);
    if (!logRow) throw new Error('Data permainan tidak ditemukan.');

    const log = logSheet.getRange(logRow, 1, 1, 11).getValues()[0];
    if (String(log[1]) !== userId) throw new Error('User ID tidak sesuai.');

    if (String(log[10]) === 'FINISHED') {
      return {
        winner: String(log[9]),
        prize: Number(log[4]) || 0,
        prizeLabel: formatRupiah_(Number(log[4]) || 0),
        ticketsLeft: Number(log[6]) || 0,
        totalPlayed: Number(log[7]) || 0
      };
    }

    const memberRow = findMemberRow_(memberSheet, userId);
    if (!memberRow) throw new Error('Data member tidak ditemukan.');

    const member = memberSheet.getRange(memberRow, 1, 1, 7).getValues()[0];
    const totalPlayed = (Number(member[2]) || 0) + 1;
    const prize = Number(log[4]) || 0;
    const totalPrize = (Number(member[3]) || 0) + prize;
    const finishedAt = new Date();

    memberSheet.getRange(memberRow, 3, 1, 3).setValues([[
      totalPlayed,
      totalPrize,
      finishedAt
    ]]);

    logSheet.getRange(logRow, 4).setValue(finishedAt);
    logSheet.getRange(logRow, 8).setValue(totalPlayed);
    logSheet.getRange(logRow, 11).setValue('FINISHED');

    return {
      winner: String(log[9]),
      prize: prize,
      prizeLabel: formatRupiah_(prize),
      ticketsLeft: Number(log[6]) || 0,
      totalPlayed: totalPlayed
    };
  } finally {
    lock.releaseLock();
  }
}

/* =========================
   ADMIN API
========================= */

function adminLogin(pin) {
  checkAdmin_(pin);
  return getAdminData_(pin);
}

function saveMemberTickets(pin, userId, ticketAmount, mode) {
  checkAdmin_(pin);

  userId = normalizeUserId_(userId);
  ticketAmount = Math.max(0, Math.floor(Number(ticketAmount) || 0));
  mode = mode === 'SET' ? 'SET' : 'ADD';

  if (userId.length < 3) throw new Error('User ID minimal 3 karakter.');
  if (ticketAmount < 1 && mode === 'ADD') {
    throw new Error('Jumlah tiket minimal 1.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = getSpreadsheet_().getSheetByName(MEMBER_SHEET);
    let row = findMemberRow_(sheet, userId);

    if (!row) {
      sheet.appendRow([
        userId,
        ticketAmount,
        0,
        0,
        '',
        new Date(),
        'ACTIVE'
      ]);
    } else {
      const current = Number(sheet.getRange(row, 2).getValue()) || 0;
      const newValue = mode === 'SET' ? ticketAmount : current + ticketAmount;
      sheet.getRange(row, 2).setValue(newValue);
      sheet.getRange(row, 7).setValue('ACTIVE');
    }

    return getAdminData_(pin);
  } finally {
    lock.releaseLock();
  }
}

function setMemberStatus(pin, userId, status) {
  checkAdmin_(pin);
  userId = normalizeUserId_(userId);
  status = status === 'INACTIVE' ? 'INACTIVE' : 'ACTIVE';

  const sheet = getSpreadsheet_().getSheetByName(MEMBER_SHEET);
  const row = findMemberRow_(sheet, userId);
  if (!row) throw new Error('User ID tidak ditemukan.');

  sheet.getRange(row, 7).setValue(status);
  return getAdminData_(pin);
}

function saveGameSettings(pin, payload) {
  checkAdmin_(pin);
  payload = payload || {};

  const settings = getSettings_();
  settings.enabled = payload.enabled !== false;
  settings.duration = Math.max(5, Math.min(30, Number(payload.duration) || 12));
  settings.prize = Math.max(0, Math.floor(Number(payload.prize) || 0));

  if (Array.isArray(payload.players)) {
    const players = payload.players
      .map(function(name) { return cleanText_(name, 24); })
      .filter(Boolean)
      .slice(0, 6);

    if (players.length < 2) throw new Error('Minimal harus ada 2 peserta.');
    settings.players = players;
  }

  saveSettings_(settings);
  return getAdminData_(pin);
}

function getAdminData(pin) {
  checkAdmin_(pin);
  return getAdminData_(pin);
}

function getAdminData_() {
  const ss = getSpreadsheet_();
  const memberSheet = ss.getSheetByName(MEMBER_SHEET);
  const logSheet = ss.getSheetByName(LOG_SHEET);
  const settings = getSettings_();

  const members = readData_(memberSheet, 7).map(function(row) {
    return {
      userId: String(row[0]),
      tickets: Number(row[1]) || 0,
      totalPlayed: Number(row[2]) || 0,
      totalPrize: Number(row[3]) || 0,
      totalPrizeLabel: formatRupiah_(Number(row[3]) || 0),
      lastPlayed: dateText_(row[4]),
      status: String(row[6] || 'ACTIVE')
    };
  }).reverse();

  const logs = readData_(logSheet, 11).map(function(row) {
    return {
      playId: String(row[0]),
      userId: String(row[1]),
      startedAt: dateText_(row[2]),
      finishedAt: dateText_(row[3]),
      prize: Number(row[4]) || 0,
      prizeLabel: formatRupiah_(Number(row[4]) || 0),
      ticketsLeft: Number(row[6]) || 0,
      totalPlayed: Number(row[7]) || 0,
      chosenPlayer: String(row[8] || ''),
      winner: String(row[9] || ''),
      status: String(row[10] || '')
    };
  }).reverse().slice(0, 100);

  const stats = members.reduce(function(result, member) {
    result.totalTickets += member.tickets;
    result.totalPlayed += member.totalPlayed;
    result.totalPrize += member.totalPrize;
    return result;
  }, {
    totalMembers: members.length,
    totalTickets: 0,
    totalPlayed: 0,
    totalPrize: 0
  });

  stats.totalPrizeLabel = formatRupiah_(stats.totalPrize);

  return {
    settings: {
      enabled: settings.enabled,
      duration: settings.duration,
      prize: settings.prize,
      prizeLabel: formatRupiah_(settings.prize),
      players: settings.players
    },
    stats: stats,
    members: members.slice(0, 100),
    logs: logs
  };
}

/* =========================
   HELPERS
========================= */

function checkAdmin_(pin) {
  if (String(pin || '') !== String(ADMIN_PIN)) {
    throw new Error('PIN admin salah.');
  }
}

function findMemberRow_(sheet, userId) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const finder = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
    .createTextFinder(userId)
    .matchEntireCell(true)
    .findNext();
  return finder ? finder.getRow() : 0;
}

function findLogRow_(sheet, playId) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const finder = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
    .createTextFinder(playId)
    .matchEntireCell(true)
    .findNext();
  return finder ? finder.getRow() : 0;
}

function readData_(sheet, columns) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, columns).getValues();
}

function memberDataFromRow_(row) {
  return {
    userId: String(row[0]),
    tickets: Number(row[1]) || 0,
    totalPlayed: Number(row[2]) || 0,
    totalPrize: Number(row[3]) || 0,
    totalPrizeLabel: formatRupiah_(Number(row[3]) || 0),
    lastPlayed: dateText_(row[4]),
    status: String(row[6] || 'ACTIVE')
  };
}

function normalizeUserId_(value) {
  return String(value || '')
    .replace(/[<>]/g, '')
    .replace(/\s+/g, '')
    .trim()
    .toUpperCase()
    .slice(0, 40);
}

function cleanText_(value, maxLength) {
  return String(value || '')
    .replace(/[<>]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, maxLength || 100);
}

function formatRupiah_(amount) {
  return 'Rp' + Math.floor(Number(amount) || 0)
    .toString()
    .replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}

function dateText_(value) {
  if (!value) return '-';
  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return String(value);
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'dd/MM/yyyy HH:mm:ss'
  );
}

/* =========================
   HTML + CSS + JS
========================= */

function buildHtml_(page) {
  return `<!DOCTYPE html>
<html lang="id">
<head>
<meta charset="UTF-8">
<style>
*{box-sizing:border-box}body{margin:0;font-family:Arial,sans-serif;background:#f3f5f9;color:#18202c}button,input,select{font:inherit}.hide{display:none!important}.page{min-height:100vh;padding:18px;display:grid;place-items:center}.card{width:min(980px,100%);background:#fff;border-radius:20px;box-shadow:0 18px 55px #0002;padding:24px}.small{width:min(440px,100%);text-align:center}.red{color:#d90022}.muted{color:#738095}.btn{border:0;border-radius:11px;padding:12px 15px;font-weight:bold;cursor:pointer;background:#273142;color:#fff}.primary{background:#d90022}.light{background:#edf1f6;color:#17202c}.danger{background:#8d1024}.full{width:100%}input,select{width:100%;padding:12px;border:1px solid #dfe3ea;border-radius:10px;outline:0}label{display:block;margin:0 0 6px;font-size:12px;font-weight:bold}.field{margin-bottom:13px;text-align:left}.top{display:flex;justify-content:space-between;align-items:center;gap:12px}.badge{padding:9px 12px;border-radius:10px;background:#f1f4f8}.players{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-top:18px}.player{padding:17px 8px;border:1px solid #e2e6ed;border-radius:14px;background:#fafbfc;cursor:pointer;text-align:center}.player:hover{border-color:#d90022}.avatar{font-size:44px;margin-bottom:7px}.timer{margin:15px 0;padding:12px;border-radius:12px;background:#f2f4f8}.bar{height:9px;background:#dfe4eb;border-radius:99px;overflow:hidden;margin-top:8px}.bar i{display:block;width:100%;height:100%;background:#d90022;transform-origin:left}.arena{position:relative;height:390px;overflow:hidden;border-radius:15px;background:repeating-linear-gradient(0deg,#c98a4c 0 24%,#fff 24% 25%)}.finish{position:absolute;right:40px;top:0;bottom:0;width:45px;background:repeating-conic-gradient(#111 0 25%,#fff 0 50%) 50%/20px 20px}.tracks{position:absolute;inset:0}.lane{position:relative;height:25%;border-bottom:1px dashed #fff}.runner{position:absolute;left:12px;top:50%;transform:translateY(-50%);width:70px;text-align:center;transition:left .08s linear}.runner b{display:block;font-size:11px}.runner.chosen{filter:drop-shadow(0 0 8px #ffd33d)}.sack{font-size:42px;animation:hop .35s infinite alternate}@keyframes hop{to{transform:translateY(-7px)}}.result{text-align:center}.result .cup{font-size:70px}.result h1{font-size:42px;margin:8px 0;color:#d90022}.stats{display:grid;grid-template-columns:repeat(4,1fr);gap:9px;margin:15px 0}.stat{padding:13px;border-radius:12px;background:#f4f6f9}.stat small{display:block;color:#738095}.stat strong{font-size:21px}.admin{width:min(1250px,100%)}.grid{display:grid;grid-template-columns:360px 1fr;gap:12px}.box{border:1px solid #e1e5eb;border-radius:13px;padding:14px}.actions{display:flex;gap:8px;flex-wrap:wrap}.table{overflow:auto;max-height:350px;border:1px solid #e1e5eb;border-radius:11px}table{width:100%;border-collapse:collapse;min-width:700px;font-size:12px}th,td{padding:9px;border-bottom:1px solid #e8ebf0;text-align:left}th{position:sticky;top:0;background:#f2f4f7}.player-inputs{display:grid;grid-template-columns:1fr 1fr;gap:7px}.toast{position:fixed;right:15px;bottom:15px;padding:12px 15px;border-radius:10px;background:#162131;color:#fff;opacity:0;transform:translateY(20px);transition:.2s}.toast.show{opacity:1;transform:none}@media(max-width:760px){.players{grid-template-columns:repeat(2,1fr)}.grid{grid-template-columns:1fr}.stats{grid-template-columns:repeat(2,1fr)}.top{align-items:flex-start;flex-direction:column}.arena{height:350px}.finish{right:10px}}
</style>
</head>
<body>

<section id="memberLogin" class="page ${page === 'admin' ? 'hide' : ''}">
  <div class="card small">
    <div style="font-size:52px">🇮🇩</div>
    <h1 class="red">Balap Karung</h1>
    <p class="muted">Login menggunakan User ID yang sudah diberikan tiket oleh admin.</p>
    <div class="field"><label>User ID</label><input id="memberId" placeholder="Contoh: CRB12345"></div>
    <div id="memberError" class="red" style="min-height:22px;font-size:12px"></div>
    <button class="btn primary full" onclick="loginMember()">Masuk Mini Game</button>
    <button class="btn light full" style="margin-top:8px" onclick="openAdmin()">Panel Admin</button>
  </div>
</section>

<section id="memberPanel" class="page hide">
  <div class="card">
    <div class="top">
      <div><h2>Halo, <span id="memberName"></span></h2><p class="muted">Pilih satu peserta untuk memulai permainan.</p></div>
      <div class="badge">Tiket tersedia: <b id="ticketCount">0</b></div>
    </div>
    <div class="players" id="playerList"></div>
    <div style="margin-top:16px"><button class="btn light" onclick="logoutMember()">Keluar</button></div>
  </div>
</section>

<section id="gamePanel" class="page hide">
  <div class="card">
    <div class="top"><h2>Balapan Berlangsung</h2><div class="badge">Pilihan: <b id="chosenName"></b></div></div>
    <div class="timer">Sisa waktu: <b id="timeLeft">12.0</b> detik<div class="bar"><i id="timeBar"></i></div></div>
    <div class="arena"><div class="finish"></div><div class="tracks" id="tracks"></div></div>
  </div>
</section>

<section id="resultPanel" class="page hide">
  <div class="card small result">
    <div class="cup">🏆</div>
    <p class="red"><b>PERTANDINGAN SELESAI</b></p>
    <h1 id="winnerName">-</h1>
    <p>Hadiah yang didapat:</p>
    <h2 id="prizeResult">Rp0</h2>
    <p class="muted">Total bermain: <b id="playedResult">0</b> kali<br>Tiket tersisa: <b id="ticketResult">0</b></p>
    <button class="btn primary full" onclick="backToMember()">Main Lagi</button>
  </div>
</section>

<section id="adminLogin" class="page ${page === 'admin' ? '' : 'hide'}">
  <div class="card small">
    <div style="font-size:54px">🔐</div>
    <h2>Admin Panel</h2>
    <div class="field"><label>PIN Admin</label><input id="adminPin" type="password"></div>
    <div id="adminError" class="red" style="min-height:22px;font-size:12px"></div>
    <button class="btn primary full" onclick="loginAdmin()">Masuk</button>
    <button class="btn light full" style="margin-top:8px" onclick="openMember()">Kembali ke Game</button>
  </div>
</section>

<section id="adminPanel" class="page hide">
  <div class="card admin">
    <div class="top">
      <div><h2>Dashboard Admin Balap Karung</h2><p class="muted">Input tiket, atur hadiah, dan lihat laporan Spreadsheet.</p></div>
      <div class="actions"><button class="btn light" onclick="refreshAdmin()">Refresh</button><button class="btn danger" onclick="logoutAdmin()">Keluar</button></div>
    </div>

    <div class="stats">
      <div class="stat"><small>Total Member</small><strong id="sMembers">0</strong></div>
      <div class="stat"><small>Tiket Aktif</small><strong id="sTickets">0</strong></div>
      <div class="stat"><small>Total Permainan</small><strong id="sPlayed">0</strong></div>
      <div class="stat"><small>Total Hadiah</small><strong id="sPrize">Rp0</strong></div>
    </div>

    <div class="grid">
      <div>
        <div class="box">
          <h3>Input Tiket Member</h3>
          <div class="field"><label>User ID</label><input id="ticketUserId"></div>
          <div class="field"><label>Jumlah Tiket</label><input id="ticketAmount" type="number" min="0" value="1"></div>
          <div class="field"><label>Metode</label><select id="ticketMode"><option value="ADD">Tambah tiket</option><option value="SET">Set jumlah tiket</option></select></div>
          <button class="btn primary full" onclick="saveTickets()">Simpan Tiket</button>
        </div>

        <div class="box" style="margin-top:12px">
          <h3>Pengaturan Game</h3>
          <div class="field"><label><input id="gameEnabled" type="checkbox" style="width:auto"> Game aktif</label></div>
          <div class="field"><label>Durasi (detik)</label><input id="gameDuration" type="number" min="5" max="30"></div>
          <div class="field"><label>Hadiah per permainan</label><input id="gamePrize" type="number" min="0"></div>
          <label>Nama Peserta</label>
          <div class="player-inputs" id="playerInputs"></div>
          <button class="btn primary full" style="margin-top:10px" onclick="saveSettings()">Simpan Pengaturan</button>
        </div>
      </div>

      <div>
        <div class="box">
          <h3>Data Member</h3>
          <div class="table"><table><thead><tr><th>User ID</th><th>Tiket</th><th>Main</th><th>Total Hadiah</th><th>Terakhir Main</th><th>Status</th></tr></thead><tbody id="memberRows"></tbody></table></div>
        </div>
        <div class="box" style="margin-top:12px">
          <h3>Riwayat Permainan</h3>
          <div class="table"><table><thead><tr><th>Tanggal</th><th>User ID</th><th>Pilihan</th><th>Hadiah</th><th>Tiket Sisa</th><th>Total Main</th><th>Status</th></tr></thead><tbody id="logRows"></tbody></table></div>
        </div>
      </div>
    </div>
  </div>
</section>

<div class="toast" id="toast"></div>

<script>
var settings=null,member=null,selected=null,play=null,adminPinValue='',raf=0,startTime=0;

function show(id){['memberLogin','memberPanel','gamePanel','resultPanel','adminLogin','adminPanel'].forEach(function(x){document.getElementById(x).classList.add('hide')});document.getElementById(id).classList.remove('hide')}
function toast(text){var el=document.getElementById('toast');el.textContent=text;el.classList.add('show');clearTimeout(toast.t);toast.t=setTimeout(function(){el.classList.remove('show')},2200)}
function fail(e){toast(e.message||'Terjadi kesalahan')}
function openAdmin(){show('adminLogin')}
function openMember(){show('memberLogin')}
function logoutMember(){member=null;show('memberLogin')}
function logoutAdmin(){adminPinValue='';show('adminLogin')}

function loadSettings(done){
 google.script.run.withSuccessHandler(function(x){settings=x;if(done)done()}).withFailureHandler(fail).getPublicSettings()
}

function loginMember(){
 var id=document.getElementById('memberId').value;
 document.getElementById('memberError').textContent='';
 google.script.run.withSuccessHandler(function(x){
   member=x;
   loadSettings(function(){
     document.getElementById('memberName').textContent=member.userId;
     document.getElementById('ticketCount').textContent=member.tickets;
     renderPlayers();
     show('memberPanel');
   });
 }).withFailureHandler(function(e){document.getElementById('memberError').textContent=e.message}).memberLogin(id)
}

function renderPlayers(){
 var box=document.getElementById('playerList');box.innerHTML='';
 settings.players.forEach(function(name,i){
   var b=document.createElement('button');b.className='player';
   b.innerHTML='<div class="avatar">'+['🧑','👩','👨','👧','🧔','👩‍🦱'][i%6]+'</div><b>'+safe(name)+'</b>';
   b.onclick=function(){startGame(name)};box.appendChild(b)
 })
}

function startGame(name){
 if(member.tickets<1){toast('Tiket bermain sudah habis');return}
 selected=name;
 google.script.run.withSuccessHandler(function(x){
   play=x;document.getElementById('chosenName').textContent=name;
   renderTracks();show('gamePanel');startTime=performance.now();raf=requestAnimationFrame(raceLoop)
 }).withFailureHandler(fail).beginGame(member.userId,name)
}

function renderTracks(){
 var box=document.getElementById('tracks');box.innerHTML='';
 settings.players.slice(0,4).forEach(function(name,i){
   var lane=document.createElement('div');lane.className='lane';
   var r=document.createElement('div');r.className='runner'+(name===selected?' chosen':'');r.dataset.name=name;
   r.innerHTML='<div class="sack">🧍</div><b>'+safe(name)+'</b>';
   lane.appendChild(r);box.appendChild(lane)
 })
}

function raceLoop(now){
 var duration=play.duration,elapsed=Math.min(duration,(now-startTime)/1000),p=elapsed/duration;
 document.getElementById('timeLeft').textContent=(duration-elapsed).toFixed(1);
 document.getElementById('timeBar').style.transform='scaleX('+(1-p)+')';
 var width=document.getElementById('tracks').clientWidth-120;
 Array.prototype.forEach.call(document.querySelectorAll('.runner'),function(r,i){
   var chosen=r.dataset.name===selected;
   var max=chosen?1:(.72+i*.045);
   var move=(1-Math.pow(1-p,3))*max+(Math.sin(elapsed*3+i)*.01);
   if(chosen&&p>.98)move=1;
   r.style.left=(12+Math.max(0,Math.min(max,move))*width)+'px'
 });
 if(elapsed<duration){raf=requestAnimationFrame(raceLoop)}else{finishGame()}
}

function finishGame(){
 cancelAnimationFrame(raf);
 google.script.run.withSuccessHandler(function(x){
   document.getElementById('winnerName').textContent=x.winner;
   document.getElementById('prizeResult').textContent=x.prizeLabel;
   document.getElementById('playedResult').textContent=x.totalPlayed;
   document.getElementById('ticketResult').textContent=x.ticketsLeft;
   member.tickets=x.ticketsLeft;member.totalPlayed=x.totalPlayed;show('resultPanel')
 }).withFailureHandler(fail).finishGame(play.playId,member.userId)
}

function backToMember(){
 if(member.tickets>0){document.getElementById('ticketCount').textContent=member.tickets;show('memberPanel')}
 else{toast('Tiket sudah habis');show('memberLogin')}
}

function loginAdmin(){
 var pin=document.getElementById('adminPin').value;
 document.getElementById('adminError').textContent='';
 google.script.run.withSuccessHandler(function(x){adminPinValue=pin;renderAdmin(x);show('adminPanel')})
 .withFailureHandler(function(e){document.getElementById('adminError').textContent=e.message}).adminLogin(pin)
}

function refreshAdmin(){
 google.script.run.withSuccessHandler(function(x){renderAdmin(x);toast('Data diperbarui')}).withFailureHandler(fail).getAdminData(adminPinValue)
}

function renderAdmin(x){
 document.getElementById('sMembers').textContent=x.stats.totalMembers;
 document.getElementById('sTickets').textContent=x.stats.totalTickets;
 document.getElementById('sPlayed').textContent=x.stats.totalPlayed;
 document.getElementById('sPrize').textContent=x.stats.totalPrizeLabel;
 document.getElementById('gameEnabled').checked=x.settings.enabled;
 document.getElementById('gameDuration').value=x.settings.duration;
 document.getElementById('gamePrize').value=x.settings.prize;
 var inputs=document.getElementById('playerInputs');inputs.innerHTML='';
 x.settings.players.forEach(function(name){var i=document.createElement('input');i.className='pname';i.value=name;inputs.appendChild(i)});
 document.getElementById('memberRows').innerHTML=x.members.length?x.members.map(function(m){return '<tr><td><b>'+safe(m.userId)+'</b></td><td>'+m.tickets+'</td><td>'+m.totalPlayed+'</td><td>'+m.totalPrizeLabel+'</td><td>'+m.lastPlayed+'</td><td>'+m.status+'</td></tr>'}).join(''):'<tr><td colspan="6">Belum ada member.</td></tr>';
 document.getElementById('logRows').innerHTML=x.logs.length?x.logs.map(function(l){return '<tr><td>'+l.finishedAt+'</td><td><b>'+safe(l.userId)+'</b></td><td>'+safe(l.chosenPlayer)+'</td><td>'+l.prizeLabel+'</td><td>'+l.ticketsLeft+'</td><td>'+l.totalPlayed+'</td><td>'+l.status+'</td></tr>'}).join(''):'<tr><td colspan="7">Belum ada permainan.</td></tr>'
}

function saveTickets(){
 google.script.run.withSuccessHandler(function(x){renderAdmin(x);toast('Tiket berhasil disimpan');document.getElementById('ticketUserId').value=''})
 .withFailureHandler(fail).saveMemberTickets(adminPinValue,document.getElementById('ticketUserId').value,document.getElementById('ticketAmount').value,document.getElementById('ticketMode').value)
}

function saveSettings(){
 var players=Array.prototype.map.call(document.querySelectorAll('.pname'),function(i){return i.value});
 google.script.run.withSuccessHandler(function(x){renderAdmin(x);toast('Pengaturan disimpan')}).withFailureHandler(fail)
 .saveGameSettings(adminPinValue,{enabled:document.getElementById('gameEnabled').checked,duration:document.getElementById('gameDuration').value,prize:document.getElementById('gamePrize').value,players:players})
}

function safe(x){return String(x||'').replace(/[&<>"']/g,function(c){return{'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]})}

if('${page}'==='member'){loadSettings()}
</script>
</body>
</html>`;
}
