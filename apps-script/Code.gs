// ====================================================================
// CONFIGURATION & CONSTANTS
// ====================================================================
var SPREADSHEET_ID = ""; // <-- TAMPAL GOOGLE SHEET ID ANDA DI SINI

var SESSION_TTL_HOURS = 12; 
var MAX_LOGIN_ATTEMPTS = 8; 
var RATE_LIMIT_MINUTES = 5; 

function pastikanTab(namaTab, senaraiHeader) {
  var ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(namaTab);

  if (!sheet) {
    sheet = ss.insertSheet(namaTab);
    sheet.getRange(1, 1, 1, senaraiHeader.length).setValues([senaraiHeader]).setFontWeight("bold");
    sheet.setFrozenRows(1);
    return sheet;
  }

  var currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
  var missingHeaders = senaraiHeader.filter(function(h) { return currentHeaders.indexOf(h) === -1; });

  if (missingHeaders.length > 0 && currentHeaders[0] !== "") {
    var startCol = currentHeaders.length + 1;
    sheet.getRange(1, startCol, 1, missingHeaders.length).setValues([missingHeaders]).setFontWeight("bold");
  } else if (currentHeaders[0] === "") {
    sheet.getRange(1, 1, 1, senaraiHeader.length).setValues([senaraiHeader]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getAppSalt() {
  var props = PropertiesService.getScriptProperties();
  var salt = props.getProperty("APP_SECURITY_SALT");
  if (!salt) { salt = Utilities.getUuid() + "-" + new Date().getTime(); props.setProperty("APP_SECURITY_SALT", salt); }
  return salt;
}

function hashPassword(password, userSalt) {
  var appSalt = getAppSalt();
  var combined = password + userSalt + appSalt;
  var byteSignature = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, combined, Utilities.Charset.UTF_8);
  var hexString = "";
  for (var i = 0; i < byteSignature.length; i++) {
    var byteVal = byteSignature[i]; if (byteVal < 0) byteVal += 256;
    var byteHex = byteVal.toString(16); if (byteHex.length == 1) byteHex = "0" + byteHex;
    hexString += byteHex;
  }
  return hexString;
}

function sanitizeFormula(inputVal) {
  if (typeof inputVal !== "string") return inputVal;
  var trimmed = inputVal.trim();
  if (/^[=+\-@]/.test(trimmed)) return "'" + trimmed;
  return trimmed;
}

function logActivity(emel, tindakan, status) {
  try {
    var sheetLog = pastikanTab("Log", ["Masa", "Emel", "Tindakan", "Status"]);
    sheetLog.appendRow([new Date(), sanitizeFormula(emel), sanitizeFormula(tindakan), sanitizeFormula(status)]);
  } catch (e) {}
}

function checkRateLimit(emel) {
  var cache = CacheService.getScriptCache();
  var key = "LOGIN_ATTEMPTS_" + emel.toLowerCase();
  var attempts = parseInt(cache.get(key) || "0", 10);
  if (attempts >= MAX_LOGIN_ATTEMPTS) return false;
  cache.put(key, (attempts + 1).toString(), RATE_LIMIT_MINUTES * 60);
  return true;
}

function clearRateLimit(emel) {
  CacheService.getScriptCache().remove("LOGIN_ATTEMPTS_" + emel.toLowerCase());
}

function validateSession(token, requiredRole) {
  if (!token) return { valid: false, error: "Token sesi tidak wujud." };
  var sheetSessions = pastikanTab("Sessions", ["Token", "Emel", "Role", "ExpiryTimestamp"]);
  var data = sheetSessions.getDataRange().getValues();
  var now = new Date().getTime();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === token) {
      if (now > parseInt(data[i][3], 10)) return { valid: false, error: "Sesi telah tamat tempoh. Sila log masuk semula." };
      if (requiredRole && data[i][2] !== requiredRole) return { valid: false, error: "Capaian ditolak. Kebenaran tidak mencukupi." };
      return { valid: true, emel: data[i][1], role: data[i][2] };
    }
  }
  return { valid: false, error: "Token sesi tidak sah." };
}

function setupAdminPertama() {
  var EMEL_ADMIN         = "admin@smiaag.edu.my"; 
  var PASSWORD_SEMENTARA = "TukarSaya123";           

  var sheetUsers = pastikanTab("Users", ["Emel", "Nama", "Role", "PasswordHash", "Salt", "MestiTukarPassword", "ResetKodHash", "ResetLuput", "Status", "TarikhDaftar"]);
  pastikanTab("Sessions", ["Token", "Emel", "Role", "ExpiryTimestamp"]);
  pastikanTab("Log", ["Masa", "Emel", "Tindakan", "Status"]);

  var dataUsers = sheetUsers.getDataRange().getValues();
  for (var i = 1; i < dataUsers.length; i++) {
    if (dataUsers[i][0].toString().toLowerCase() === EMEL_ADMIN.toLowerCase()) return "Akaun Admin sudah wujud!";
  }

  var userSalt = Utilities.getUuid();
  var passHash = hashPassword(PASSWORD_SEMENTARA, userSalt);
  sheetUsers.appendRow([sanitizeFormula(EMEL_ADMIN.toLowerCase()), "Pentadbir Utama", "ADMIN", passHash, userSalt, true, "", "", "ACTIVE", new Date()]);

  logActivity(EMEL_ADMIN, "SETUP_ADMIN_PERTAMA", "SUCCESS");
  Logger.log("BERJAYA! Akaun Admin dicipta: " + EMEL_ADMIN);
  return "Berjaya mencipta Akaun Admin!";
}

function doGet(e) { return ContentService.createTextOutput(JSON.stringify({ status: "online" })).setMimeType(ContentService.MimeType.JSON); }

function jsonResponse(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

function handleMintaKodReset(payload) {
  var emel = (payload.emel || "").trim().toLowerCase();
  var GENERIC_MSG = "Kalau emel ini berdaftar, kod telah dihantar."; 
  
  if (!emel) return jsonResponse({ success: true, message: GENERIC_MSG });

  var cache = CacheService.getScriptCache();
  var cacheKey = "RESET_REQ_" + emel;
  var reqCount = parseInt(cache.get(cacheKey) || "0", 10);
  if (reqCount >= 3) return jsonResponse({ success: true, message: GENERIC_MSG }); 
  cache.put(cacheKey, (reqCount + 1).toString(), 3600); 

  var sheetUsers = pastikanTab("Users", ["Emel", "Nama", "Role", "PasswordHash", "Salt", "MestiTukarPassword", "ResetKodHash", "ResetLuput", "Status", "TarikhDaftar"]);
  var data = sheetUsers.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === emel && data[i][8] === "ACTIVE") {
      var kodRawak = Math.floor(100000 + Math.random() * 900000).toString(); 
      var hashedKod = hashPassword(kodRawak, data[i][4]); 
      var luput = new Date().getTime() + (15 * 60 * 1000); 

      sheetUsers.getRange(i + 1, 7).setValue(hashedKod); 
      sheetUsers.getRange(i + 1, 8).setValue(luput);     
      
      try {
        MailApp.sendEmail(emel, "[e-Cuti] Kod Tetapan Semula Kata Laluan", 
          "Seseorang telah meminta untuk menetapkan semula kata laluan anda.\n\n" +
          "Kod 6-digit anda ialah: " + kodRawak + "\n\n" + 
          "Kod ini sah selama 15 minit. Jika anda tidak memintanya, abaikan emel ini."
        );
      } catch(e) {}
      logActivity(emel, "MINTA_KOD_RESET", "SENT");
      break; 
    }
  }
  return jsonResponse({ success: true, message: GENERIC_MSG });
}

function handleSahkanKodReset(payload) {
  var emel = (payload.emel || "").trim().toLowerCase();
  var kod = payload.kod || "";
  var newPassword = payload.newPassword || "";
  
  if (!emel || !kod || newPassword.length < 6) return jsonResponse({ success: false, message: "Maklumat tidak lengkap atau kata laluan terlalu pendek." });

  var cache = CacheService.getScriptCache();
  var attemptKey = "RESET_ATTEMPT_" + emel;
  var attempts = parseInt(cache.get(attemptKey) || "0", 10);
  
  if (attempts >= 5) return jsonResponse({ success: false, message: "Terlalu banyak percubaan gagal. Kod telah dibatalkan." });

  var sheetUsers = pastikanTab("Users", ["Emel", "Nama", "Role", "PasswordHash", "Salt", "MestiTukarPassword", "ResetKodHash", "ResetLuput", "Status", "TarikhDaftar"]);
  var data = sheetUsers.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === emel && data[i][8] === "ACTIVE") {
      var dbKodHash = data[i][6];
      var dbLuput = parseInt(data[i][7], 10);
      var now = new Date().getTime();

      if (!dbKodHash || now > dbLuput) {
        logActivity(emel, "SAHKAN_KOD_RESET", "EXPIRED_OR_EMPTY");
        return jsonResponse({ success: false, message: "Kod tidak sah atau telah tamat tempoh." });
      }

      var inputKodHash = hashPassword(kod, data[i][4]);
      if (inputKodHash === dbKodHash) {
        var newSalt = Utilities.getUuid();
        sheetUsers.getRange(i + 1, 4).setValue(hashPassword(newPassword, newSalt)); 
        sheetUsers.getRange(i + 1, 5).setValue(newSalt); 
        sheetUsers.getRange(i + 1, 6).setValue(false); 
        sheetUsers.getRange(i + 1, 7).setValue(""); 
        sheetUsers.getRange(i + 1, 8).setValue("");
        
        cache.remove(attemptKey);
        invalidateUserSessions(emel); 
        logActivity(emel, "TUKAR_PASSWORD_RESET", "SUCCESS");
        return jsonResponse({ success: true, message: "Kata laluan berjaya ditukar! Sila log masuk." });
      } else {
        attempts++;
        cache.put(attemptKey, attempts.toString(), 15 * 60); 
        if (attempts >= 5) {
          sheetUsers.getRange(i + 1, 7).setValue(""); 
          sheetUsers.getRange(i + 1, 8).setValue("");
          logActivity(emel, "SAHKAN_KOD_RESET", "LOCKED_MAX_ATTEMPTS");
          return jsonResponse({ success: false, message: "Terlalu banyak percubaan salah. Sila minta kod baharu." });
        }
        logActivity(emel, "SAHKAN_KOD_RESET", "WRONG_CODE");
        return jsonResponse({ success: false, message: "Kod salah. Baki percubaan: " + (5 - attempts) });
      }
    }
  }
  return jsonResponse({ success: false, message: "Ralat. Kod tidak sah." });
}

function handleGetUsers(payload, sessionCheck) {
  if (sessionCheck.role !== "ADMIN") return jsonResponse({ success: false, message: "Capaian ditolak." });
  var data = pastikanTab("Users", ["Emel", "Nama", "Role", "PasswordHash", "Salt", "MestiTukarPassword", "ResetKodHash", "ResetLuput", "Status", "TarikhDaftar"]).getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    if(data[i][8] === "ACTIVE") users.push({ emel: data[i][0], nama: data[i][1], role: data[i][2], tarikhDaftar: data[i][9] });
  }
  return jsonResponse({ success: true, data: users });
}

function handleDeleteUser(payload, sessionCheck) {
  if (sessionCheck.role !== "ADMIN") return jsonResponse({ success: false, message: "Capaian ditolak." });
  var targetEmel = payload.emelTarget.toLowerCase();
  
  var sheetUsers = pastikanTab("Users", ["Emel", "Nama", "Role", "PasswordHash", "Salt", "MestiTukarPassword", "ResetKodHash", "ResetLuput", "Status", "TarikhDaftar"]);
  var data = sheetUsers.getDataRange().getValues();
  
  var totalAdmin = 0;
  var targetRow = -1;
  var targetRole = "";
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][8] === "ACTIVE" && data[i][2] === "ADMIN") totalAdmin++;
    if (data[i][0].toString().toLowerCase() === targetEmel && data[i][8] === "ACTIVE") { targetRow = i + 1; targetRole = data[i][2]; }
  }

  if (targetRow === -1) return jsonResponse({ success: false, message: "Pengguna tidak dijumpai." });
  if (targetRole === "ADMIN" && totalAdmin <= 1) {
    logActivity(sessionCheck.emel, "CUBA_PADAM_ADMIN_TERAKHIR", "BLOCKED");
    return jsonResponse({ success: false, message: "TINDAKAN DISEKAT: Anda tidak boleh memadam pentadbir tunggal yang terakhir di dalam sistem." });
  }

  sheetUsers.getRange(targetRow, 9).setValue("DELETED");
  invalidateUserSessions(targetEmel);
  logActivity(sessionCheck.emel, "PADAM_PENGGUNA", targetEmel);
  
  return jsonResponse({ success: true, message: "Pengguna berjaya dipadamkan." });
}

// ====================================================================
// HELPER TARIKH & PENGGUNA (BAHARU)
// ====================================================================
var TZ = "Asia/Kuala_Lumpur";
var MAX_GURU_SEHARI = 3;
var ROLE_SAH = ["ADMIN", "GURU", "STAF"];
var KATEGORI_SAH = ["Urusan Peribadi / Keluarga", "Kecemasan & Kesihatan", "Kursus / Bengkel", "Lain-Lain"];
var MIME_SAH = ["application/pdf", "image/png", "image/jpeg"];
var HEADER_USERS = ["Emel", "Nama", "Role", "PasswordHash", "Salt", "MestiTukarPassword", "ResetKodHash", "ResetLuput", "Status", "TarikhDaftar"];
var HEADER_SESSIONS = ["Token", "Emel", "Role", "ExpiryTimestamp"];
var HEADER_CUTI = ["ID", "NamaGuru", "Emel", "TarikhMohon", "TarikhMula", "TarikhAkhir", "JumlahHari", "SebabKategori", "SebabButiran", "Status", "JadualRelief", "NamaFailSurat", "UlasanPengetua"];

// Tukar nilai sel (Date atau teks) kepada "yyyy-MM-dd"
function toYMD(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, TZ, "yyyy-MM-dd");
  var s = val.toString().trim().replace(/^'/, "");
  return /^\d{4}-\d{2}-\d{2}/.test(s) ? s.substring(0, 10) : s;
}

function isValidYMD(s) {
  if (typeof s !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  var p = s.split("-"), d = new Date(Date.UTC(+p[0], +p[1] - 1, +p[2]));
  return d.getUTCFullYear() === +p[0] && d.getUTCMonth() === +p[1] - 1 && d.getUTCDate() === +p[2];
}

function formatTarikhDMY(ymd) {
  var p = (ymd || "").split("-");
  return p.length === 3 ? p[2] + "/" + p[1] + "/" + p[0] : ymd;
}

function cariPengguna(emel) {
  var sheet = pastikanTab("Users", HEADER_USERS);
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === emel.toLowerCase() && data[i][8] === "ACTIVE") {
      return { sheet: sheet, row: i + 1, data: data[i] };
    }
  }
  return null;
}

// Senarai tarikh khas: TANPA_HAD (kuota dilepaskan) dan SABTU (Sabtu yang ada program = hari bekerja)
var TAB_TARIKH = { TANPA_HAD: "TarikhTanpaHad", SABTU: "SabtuProgram" };
function jenisTarikh(payload) { return payload.jenis === "SABTU" ? "SABTU" : "TANPA_HAD"; }

function getTarikhList(jenis) {
  var data = pastikanTab(TAB_TARIKH[jenis], ["Tarikh"]).getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) { var v = toYMD(data[i][0]); if (v) out.push(v); }
  return out;
}
function getTarikhTanpaHadList() { return getTarikhList("TANPA_HAD"); }

// Hari persekolahan antara dua tarikh: Isnin–Jumaat, + Sabtu yang ditanda program. Ahad sentiasa dikecualikan.
// Kira dalam UTC supaya tak lari zon masa.
function senaraiHariBekerja(mula, akhir, sabtuProgram) {
  sabtuProgram = sabtuProgram || [];
  var a = mula.split("-"), b = akhir.split("-");
  var cur = Date.UTC(+a[0], +a[1] - 1, +a[2]), end = Date.UTC(+b[0], +b[1] - 1, +b[2]);
  var out = [];
  while (cur <= end) {
    var d = new Date(cur), ymd = Utilities.formatDate(d, "UTC", "yyyy-MM-dd"), hari = d.getUTCDay();
    if (hari >= 1 && hari <= 5) out.push(ymd);
    else if (hari === 6 && sabtuProgram.indexOf(ymd) > -1) out.push(ymd);
    cur += 86400000;
  }
  return out;
}

// Kiraan penggunaan kuota sehari (APPROVED + PENDING)
function kiraPenggunaanKuota(dataCuti, sabtuProgram) {
  var usage = {};
  for (var i = 1; i < dataCuti.length; i++) {
    var st = dataCuti[i][9].toString();
    if (st !== "APPROVED" && st !== "PENDING") continue;
    var m = toYMD(dataCuti[i][4]), a = toYMD(dataCuti[i][5]);
    if (!isValidYMD(m) || !isValidYMD(a)) continue;
    senaraiHariBekerja(m, a, sabtuProgram).forEach(function(d) { usage[d] = (usage[d] || 0) + 1; });
  }
  return usage;
}

function handleGetUnlimitedDates() {
  return jsonResponse({ success: true, data: getTarikhList("TANPA_HAD"), sabtu: getTarikhList("SABTU") });
}

function handleAddUnlimitedDate(payload, sessionCheck) {
  if (sessionCheck.role !== "ADMIN") return jsonResponse({ success: false, message: "Ditolak" });
  var jenis = jenisTarikh(payload);
  var dates = payload.dates || [];
  if (payload.date && dates.length === 0) dates = [payload.date];
  dates = dates.filter(isValidYMD);
  if (jenis === "SABTU") dates = dates.filter(function(d) { var p = d.split("-"); return new Date(Date.UTC(+p[0], +p[1] - 1, +p[2])).getUTCDay() === 6; });
  if (dates.length === 0) return jsonResponse({ success: false, message: jenis === "SABTU" ? "Tiada hari Sabtu dalam pilihan." : "Tiada tarikh sah" });

  var sedia = getTarikhList(jenis);
  var baru = dates.filter(function(d, idx) { return sedia.indexOf(d) === -1 && dates.indexOf(d) === idx; });
  if (baru.length > 0) {
    var sheet = pastikanTab(TAB_TARIKH[jenis], ["Tarikh"]);
    sheet.getRange(sheet.getLastRow() + 1, 1, baru.length, 1).setValues(baru.map(function(d) { return ["'" + d]; }));
  }
  logActivity(sessionCheck.emel, "TAMBAH_" + jenis, baru.join(","));
  return jsonResponse({ success: true, message: "Tarikh ditambah", added: baru });
}

function handleRemoveUnlimitedDate(payload, sessionCheck) {
  if (sessionCheck.role !== "ADMIN") return jsonResponse({ success: false, message: "Ditolak" });
  var jenis = jenisTarikh(payload);
  var sheet = pastikanTab(TAB_TARIKH[jenis], ["Tarikh"]);
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (toYMD(data[i][0]) === payload.date) {
      sheet.deleteRow(i + 1);
      logActivity(sessionCheck.emel, "BUANG_" + jenis, payload.date);
      return jsonResponse({ success: true, message: "Dibuang" });
    }
  }
  return jsonResponse({ success: false, message: "Tidak dijumpai" });
}

function bersihkanSesi(tapis) {
  var sheet = pastikanTab("Sessions", HEADER_SESSIONS);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;
  var simpan = [data[0]];
  for (var i = 1; i < data.length; i++) { if (!tapis(data[i])) simpan.push(data[i]); }
  if (simpan.length === data.length) return;
  sheet.getRange(1, 1, data.length, data[0].length).clearContent();
  sheet.getRange(1, 1, simpan.length, data[0].length).setValues(simpan);
}

function invalidateUserSessions(emel) {
  var e = emel.toLowerCase();
  bersihkanSesi(function(r) { return r[1].toString().toLowerCase() === e; });
}

function buangSesiTamat() {
  var now = new Date().getTime();
  bersihkanSesi(function(r) { return now > parseInt(r[3], 10); });
}

function uploadFileToDrive(base64Data, fileName, mimeType) {
  // Fail TIDAK dikongsi awam. Admin/pemilik buka melalui action GET_ATTACHMENT.
  var folderName = "eCuti_Surat_Permohonan";
  var folders = DriveApp.getFoldersByName(folderName);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
  var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
  return folder.createFile(blob).getId();
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) return jsonResponse({ success: false, message: "Permintaan tidak sah." });
    var payload = JSON.parse(e.postData.contents);

    // 1. ACTION AWAM (tanpa token)
    if (payload.action === "LOGIN") return handleLogin(payload);
    if (payload.action === "MINTA_KOD_RESET") return handleMintaKodReset(payload);
    if (payload.action === "SAHKAN_KOD_RESET") return handleSahkanKodReset(payload);
    if (payload.action === "GET_UNLIMITED_DATES") return handleGetUnlimitedDates();
    if (payload.action === "GET_LEAVE_APPLICATIONS" && payload.token === "PUBLIC_GHOST") {
      return handleGetLeaveApplications(payload, { valid: true, role: "PUBLIC", emel: "" });
    }

    // 2. ACTION WAJIB TOKEN
    var sessionCheck = validateSession(payload.token);
    if (!sessionCheck.valid) return jsonResponse({ success: false, message: sessionCheck.error, forceLogout: true });

    if (payload.action === "LOGOUT") return handleLogout(payload);

    // Pengguna yang belum tukar kata laluan sementara hanya boleh tukar kata laluan
    if (sessionCheck.role === "PENDING_PW") {
      if (payload.action === "CHANGE_PASSWORD") return handleChangePassword(payload, sessionCheck);
      return jsonResponse({ success: false, mustChangePassword: true, message: "Sila tukar kata laluan sementara dahulu." });
    }

    switch (payload.action) {
      case "CHANGE_PASSWORD": return handleChangePassword(payload, sessionCheck);
      case "SUBMIT_LEAVE": return handleSubmitLeave(payload, sessionCheck);
      case "GET_LEAVE_APPLICATIONS": return handleGetLeaveApplications(payload, sessionCheck);
      case "GET_ATTACHMENT": return handleGetAttachment(payload, sessionCheck);
      case "APPROVE_LEAVE": return handleApproveLeave(payload, sessionCheck);
      case "ADD_UNLIMITED_DATE": return handleAddUnlimitedDate(payload, sessionCheck);
      case "REMOVE_UNLIMITED_DATE": return handleRemoveUnlimitedDate(payload, sessionCheck);
      case "GET_USERS": return handleGetUsers(payload, sessionCheck);
      case "ADD_USER": return handleAddUser(payload, sessionCheck);
      case "EDIT_USER": return handleEditUser(payload, sessionCheck);
      case "DELETE_USER": return handleDeleteUser(payload, sessionCheck);
      case "GET_DATA_GURU": return handleGetDataGuru(sessionCheck);
      default: return jsonResponse({ success: false, message: "Aksi tidak dikenali." });
    }
  } catch (err) { return jsonResponse({ success: false, message: "Ralat Pelayan: " + err.toString() }); }
}

function handleLogin(payload) {
  var emel = (payload.emel || "").trim().toLowerCase();
  var password = payload.password || "";
  if (!emel || !password) return jsonResponse({ success: false, message: "Sila isi maklumat penuh." });
  if (!checkRateLimit(emel)) {
    logActivity(emel, "LOGIN", "FAILED_RATELIMIT");
    return jsonResponse({ success: false, message: "Terlalu banyak percubaan. Sila tunggu 5 minit." });
  }

  var u = cariPengguna(emel);
  if (u && hashPassword(password, u.data[4]) === u.data[3]) {
    clearRateLimit(emel);
    buangSesiTamat();
    var mesti = (u.data[5] === true || u.data[5].toString().toUpperCase() === "TRUE");
    var token = Utilities.getUuid() + "-" + Utilities.getUuid();
    var roleSesi = mesti ? "PENDING_PW" : u.data[2];
    pastikanTab("Sessions", HEADER_SESSIONS).appendRow([token, emel, roleSesi, new Date().getTime() + (SESSION_TTL_HOURS * 60 * 60 * 1000)]);
    logActivity(emel, "LOGIN", "SUCCESS");
    return jsonResponse({ success: true, token: token, emel: emel, nama: u.data[1], role: u.data[2], mestiTukarPassword: mesti });
  }
  logActivity(emel, "LOGIN", "FAILED_WRONG_CREDENTIALS");
  return jsonResponse({ success: false, message: "Emel atau kata laluan tidak sah." });
}

function handleLogout(payload) {
  var t = payload.token;
  bersihkanSesi(function(r) { return r[0] === t; });
  return jsonResponse({ success: true });
}

function handleChangePassword(payload, sessionCheck) {
  var newP = payload.newPassword || "", oldP = payload.oldPassword || "";
  if (newP.length < 6) return jsonResponse({ success: false, message: "Kata laluan baharu minimum 6 aksara." });
  if (!checkRateLimit(sessionCheck.emel)) return jsonResponse({ success: false, message: "Terlalu banyak percubaan. Sila tunggu 5 minit." });

  var u = cariPengguna(sessionCheck.emel);
  if (!u) return jsonResponse({ success: false, message: "Ralat teknikal." });
  if (hashPassword(oldP, u.data[4]) !== u.data[3]) {
    logActivity(sessionCheck.emel, "TUKAR_PASSWORD", "FAILED_OLD_PASSWORD");
    return jsonResponse({ success: false, message: "Kata laluan semasa salah." });
  }
  if (newP === oldP) return jsonResponse({ success: false, message: "Kata laluan baharu mesti berbeza." });

  clearRateLimit(sessionCheck.emel);
  var newSalt = Utilities.getUuid();
  u.sheet.getRange(u.row, 4, 1, 3).setValues([[hashPassword(newP, newSalt), newSalt, false]]);
  invalidateUserSessions(sessionCheck.emel);
  logActivity(sessionCheck.emel, "TUKAR_PASSWORD", "SUCCESS");
  return jsonResponse({ success: true, message: "Kata laluan berjaya dikemaskini! Sila log masuk semula." });
}

function handleAddUser(payload, sessionCheck) {
  if (sessionCheck.role !== "ADMIN") return jsonResponse({ success: false, message: "Capaian ditolak." });
  if (!payload.emel || !payload.nama || !payload.role || !payload.tempPassword) return jsonResponse({ success: false, message: "Sila isi semua maklumat." });

  var emelBaru = payload.emel.toLowerCase().trim();
  var role = payload.role.toString().toUpperCase().trim();
  if (ROLE_SAH.indexOf(role) === -1) return jsonResponse({ success: false, message: "Peranan tidak sah." });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(emelBaru)) return jsonResponse({ success: false, message: "Format emel tidak sah." });
  if (cariPengguna(emelBaru)) return jsonResponse({ success: false, message: "Emel ini telah wujud dalam sistem." });

  var sheetUsers = pastikanTab("Users", HEADER_USERS);
  var salt = Utilities.getUuid();
  sheetUsers.appendRow([sanitizeFormula(emelBaru), sanitizeFormula(payload.nama), role, hashPassword(payload.tempPassword, salt), salt, true, "", "", "ACTIVE", new Date()]);

  try {
    MailApp.sendEmail(emelBaru, "[e-Cuti] Akaun Portal SMIAAG Telah Dicipta",
      "Salam sejahtera,\n\nAkaun portal e-Cuti SMIAAG anda telah didaftarkan dengan peranan " + role + ".\n\n" +
      "ID Log Masuk: " + emelBaru + "\n" +
      "Kata Laluan Sementara: " + payload.tempPassword + "\n\n" +
      "Sila log masuk ke portal dan ANDA DIWAJIBKAN menukar kata laluan pada log masuk pertama.\n\nTerima kasih."
    );
  } catch(e) {}

  logActivity(sessionCheck.emel, "TAMBAH_PENGGUNA", emelBaru);
  return jsonResponse({ success: true, message: "Pengguna berjaya ditambah dan emel kata laluan sementara telah dihantar!" });
}

function handleEditUser(payload, sessionCheck) {
  if (sessionCheck.role !== "ADMIN") return jsonResponse({ success: false, message: "Capaian ditolak." });

  var targetEmel = (payload.targetEmel || "").trim().toLowerCase();
  var newEmel = (payload.emel || "").trim().toLowerCase();
  var newNama = (payload.nama || "").trim();
  var newRole = (payload.role || "").trim().toUpperCase();

  if (!targetEmel || !newEmel || !newNama || !newRole) return jsonResponse({ success: false, message: "Sila lengkapkan semua maklumat." });
  if (ROLE_SAH.indexOf(newRole) === -1) return jsonResponse({ success: false, message: "Peranan tidak sah." });

  var sheetUsers = pastikanTab("Users", HEADER_USERS);
  var data = sheetUsers.getDataRange().getValues();
  var targetRow = -1, currentRole = "", totalAdmin = 0;

  for (var i = 1; i < data.length; i++) {
    if (data[i][8] === "ACTIVE" && data[i][2] === "ADMIN") totalAdmin++;
    if (data[i][8] === "ACTIVE" && data[i][0].toString().toLowerCase() === targetEmel) { targetRow = i + 1; currentRole = data[i][2]; }
  }
  if (targetRow === -1) return jsonResponse({ success: false, message: "Pengguna tidak dijumpai." });

  if (currentRole === "ADMIN" && newRole !== "ADMIN" && totalAdmin <= 1) {
    logActivity(sessionCheck.emel, "CUBA_UBAH_PERANAN_ADMIN_TERAKHIR", "BLOCKED");
    return jsonResponse({ success: false, message: "TINDAKAN DISEKAT: Anda tidak boleh mengubah peranan pentadbir tunggal yang terakhir di dalam sistem." });
  }
  if (newEmel !== targetEmel && cariPengguna(newEmel)) {
    return jsonResponse({ success: false, message: "Alamat emel baharu ini telah pun berdaftar di bawah pengguna lain." });
  }

  sheetUsers.getRange(targetRow, 1, 1, 3).setValues([[sanitizeFormula(newEmel), sanitizeFormula(newNama), newRole]]);

  // Emel ATAU peranan berubah: batalkan semua sesi supaya peranan lama tak kekal
  if (newEmel !== targetEmel || newRole !== currentRole) invalidateUserSessions(targetEmel);

  logActivity(sessionCheck.emel, "KEMASKINI_PENGGUNA", targetEmel + " -> " + newEmel + " (" + newRole + ")");
  return jsonResponse({ success: true, message: "Maklumat pengguna berjaya dikemaskini." });
}

function handleGetDataGuru(sessionCheck) {
  var data = pastikanTab("DATA_GURU", ["NAMA_GURU", "BIL_CUTI_TAHUNAN_BERBAKI"]).getDataRange().getValues();
  var isAdmin = sessionCheck.role === "ADMIN";
  var u = isAdmin ? null : cariPengguna(sessionCheck.emel);
  var namaSendiri = u ? u.data[1].toString().trim() : "";
  var teachers = [];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    var nama = data[i][0].toString().trim();
    // Bukan admin: nampak senarai nama (untuk jadual relief), kelayakan hanya milik sendiri
    var kelayakan = (isAdmin || nama === namaSendiri) ? (parseInt(data[i][1], 10) || 0) : null;
    teachers.push({ NAMA_GURU: nama, BIL_CUTI_TAHUNAN_BERBAKI: kelayakan });
  }
  return jsonResponse({ success: true, data: teachers });
}

function bacaLampiran(raw) {
  try { var o = JSON.parse(raw); if (o && o.fileId) return o; } catch(e) {}
  return null;
}

function handleGetLeaveApplications(payload, sessionCheck) {
  var data = pastikanTab("PermohonanCuti", HEADER_CUTI).getDataRange().getValues();
  var apps = [];
  var isAdmin = sessionCheck.role === "ADMIN";
  var userEmail = (sessionCheck.emel || "").toLowerCase();

  for (var i = 1; i < data.length; i++) {
    var status = data[i][9].toString();
    var emelPemohon = data[i][2].toString().toLowerCase();
    var milikSendiri = userEmail && emelPemohon === userEmail;
    if (!(isAdmin || status === "APPROVED" || milikSendiri)) continue;

    var mohon = data[i][3], appTs = 0, appStr = "";
    if (mohon instanceof Date) { appTs = mohon.getTime(); appStr = Utilities.formatDate(mohon, TZ, "dd/MM/yyyy h.mma").toLowerCase(); }
    else { appStr = mohon.toString(); }

    var item = {
      id: data[i][0].toString(), teacherName: data[i][1].toString(),
      startDate: toYMD(data[i][4]), endDate: toYMD(data[i][5]), totalDays: parseInt(data[i][6], 10) || 0, status: status
    };

    if (isAdmin || milikSendiri) {
      var lamp = bacaLampiran(data[i][11]);
      var relief = []; try { relief = JSON.parse(data[i][10]); } catch(e) {}
      item.teacherEmail = data[i][2].toString();
      item.appDate = appStr; item.appTs = appTs;
      item.reasonCategory = data[i][7]; item.reasonDetails = data[i][8];
      item.relief = relief; item.pengetuaNotes = data[i][12];
      item.attachmentName = lamp ? lamp.name : ""; item.hasAttachment = !!lamp;
    } else {
      // Rekod orang lain: nama & tarikh sahaja
      item.teacherEmail = ""; item.appDate = ""; item.appTs = 0;
      item.reasonCategory = "Sulit"; item.reasonDetails = ""; item.relief = [];
      item.pengetuaNotes = ""; item.attachmentName = ""; item.hasAttachment = false;
    }
    apps.push(item);
  }
  return jsonResponse({ success: true, data: apps, usage: kiraPenggunaanKuota(data, getTarikhList("SABTU")) });
}

function handleGetAttachment(payload, sessionCheck) {
  var data = pastikanTab("PermohonanCuti", HEADER_CUTI).getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() !== payload.appId) continue;
    var milik = data[i][2].toString().toLowerCase() === sessionCheck.emel.toLowerCase();
    if (sessionCheck.role !== "ADMIN" && !milik) return jsonResponse({ success: false, message: "Capaian ditolak." });
    var lamp = bacaLampiran(data[i][11]);
    if (!lamp) return jsonResponse({ success: false, message: "Tiada lampiran." });
    var blob = DriveApp.getFileById(lamp.fileId).getBlob();
    logActivity(sessionCheck.emel, "LIHAT_LAMPIRAN", payload.appId);
    return jsonResponse({ success: true, name: lamp.name, mime: blob.getContentType(), base64: Utilities.base64Encode(blob.getBytes()) });
  }
  return jsonResponse({ success: false, message: "Permohonan tidak dijumpai." });
}

function handleSubmitLeave(payload, sessionCheck) {
  var app = payload.application || {};
  var u = cariPengguna(sessionCheck.emel);
  if (!u) return jsonResponse({ success: false, message: "Akaun tidak dijumpai." });

  var namaGuru = u.data[1].toString().trim();
  var emelGuru = sessionCheck.emel.toLowerCase();
  var mula = (app.startDate || "").toString(), akhir = (app.endDate || "").toString();
  var kategori = (app.reasonCategory || "").toString();
  var butiran = (app.reasonDetails || "").toString().trim().substring(0, 500);

  // --- Validasi asas ---
  if (!isValidYMD(mula) || !isValidYMD(akhir)) return jsonResponse({ success: false, message: "Format tarikh tidak sah." });
  if (akhir < mula) return jsonResponse({ success: false, message: "Tarikh akhir mesti selepas tarikh mula." });
  if (KATEGORI_SAH.indexOf(kategori) === -1) return jsonResponse({ success: false, message: "Kategori sebab tidak sah." });
  if (!butiran) return jsonResponse({ success: false, message: "Sila isi butiran lanjut." });

  var sabtuProgram = getTarikhList("SABTU");
  var hari = senaraiHariBekerja(mula, akhir, sabtuProgram);
  if (hari.length === 0) return jsonResponse({ success: false, message: "Tiada hari persekolahan dalam julat ini (Sabtu & Ahad cuti, kecuali Sabtu yang ada program)." });
  if (hari.length > 60) return jsonResponse({ success: false, message: "Julat cuti terlalu panjang (maksimum 60 hari)." });

  if (!app.fileBase64) return jsonResponse({ success: false, message: "Lampiran surat rasmi adalah wajib." });
  if (MIME_SAH.indexOf(app.fileMimeType) === -1) return jsonResponse({ success: false, message: "Jenis fail mesti PDF, PNG atau JPG." });
  if (app.fileBase64.length > 7 * 1024 * 1024) return jsonResponse({ success: false, message: "Fail melebihi 5MB." });
  var namaFail = (app.fileName || "surat").toString().replace(/[\\\/:*?"<>|]/g, "_").substring(0, 120);

  var relief = (Array.isArray(app.relief) ? app.relief : []).slice(0, 30).map(function(r) {
    // Disimpan sebagai JSON (sel bermula dengan "["), jadi tiada risiko formula
    var bersih = function(v) { return (v || "").toString().substring(0, 100); };
    return { day: bersih(r.day), time: bersih(r.time), class: bersih(r.class), replacement: bersih(r.replacement) };
  });

  // --- Bahagian kritikal: kunci supaya dua permohonan serentak tak langgar kuota / ID ---
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return jsonResponse({ success: false, message: "Sistem sibuk. Sila cuba sebentar lagi." });
  try {
    var sheet = pastikanTab("PermohonanCuti", HEADER_CUTI);
    var data = sheet.getDataRange().getValues();

    // Tindihan dengan permohonan sendiri
    for (var i = 1; i < data.length; i++) {
      var st = data[i][9].toString();
      if (data[i][2].toString().toLowerCase() !== emelGuru || (st !== "PENDING" && st !== "APPROVED")) continue;
      var m = toYMD(data[i][4]), a = toYMD(data[i][5]);
      if (m <= akhir && a >= mula) return jsonResponse({ success: false, message: "Tarikh bertindih dengan permohonan anda " + data[i][0] + "." });
    }

    // Kuota harian (APPROVED + PENDING), kecuali tarikh tanpa had
    var usage = kiraPenggunaanKuota(data, sabtuProgram), tanpaHad = getTarikhTanpaHadList();
    var penuh = hari.filter(function(d) { return tanpaHad.indexOf(d) === -1 && (usage[d] || 0) >= MAX_GURU_SEHARI; });
    if (penuh.length > 0) {
      return jsonResponse({ success: false, message: "Kuota penuh pada: " + penuh.map(formatTarikhDMY).join(", ") + ". Sila pilih tarikh lain." });
    }

    // ID unik: CT-TAHUN-nnn (nombor terbesar tahun ini + 1)
    var tahun = new Date().getFullYear(), prefix = "CT-" + tahun + "-", maks = 0;
    for (var j = 1; j < data.length; j++) {
      var id = data[j][0].toString();
      if (id.indexOf(prefix) === 0) maks = Math.max(maks, parseInt(id.substring(prefix.length), 10) || 0);
    }
    var newId = prefix + ("00" + (maks + 1)).slice(-3);

    var fileId;
    try { fileId = uploadFileToDrive(app.fileBase64, namaFail, app.fileMimeType); }
    catch (e) { return jsonResponse({ success: false, message: "Gagal memuat naik fail. Sila cuba lagi." }); }

    sheet.appendRow([newId, sanitizeFormula(namaGuru), sanitizeFormula(emelGuru), new Date(), mula, akhir, hari.length,
      kategori, sanitizeFormula(butiran), "PENDING", JSON.stringify(relief),
      JSON.stringify({ name: namaFail, fileId: fileId }), ""]);
    SpreadsheetApp.flush();
    logActivity(emelGuru, "MOHON_CUTI", newId);
    return jsonResponse({ success: true, message: "Permohonan berjaya!", id: newId });
  } finally {
    lock.releaseLock();
  }
}

function handleApproveLeave(payload, sessionCheck) {
  if (sessionCheck.role !== "ADMIN") return jsonResponse({ success: false, message: "Hanya admin boleh meluluskan cuti." });
  if (payload.newStatus !== "APPROVED" && payload.newStatus !== "REJECTED") return jsonResponse({ success: false, message: "Status tidak sah." });

  var sheet = pastikanTab("PermohonanCuti", HEADER_CUTI);
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() !== payload.appId) continue;
    if (data[i][9].toString() !== "PENDING") return jsonResponse({ success: false, message: "Permohonan ini telah diproses sebelum ini." });

    var catatan = (payload.notes || "").toString().substring(0, 500);
    sheet.getRange(i + 1, 10).setValue(payload.newStatus);
    sheet.getRange(i + 1, 13).setValue(sanitizeFormula(catatan));
    logActivity(sessionCheck.emel, "PROSES_CUTI_" + payload.newStatus, payload.appId);

    var emelGuru = data[i][2].toString();
    if (emelGuru.indexOf("@") > -1) {
      var lulus = payload.newStatus === "APPROVED";
      var body = "Salam sejahtera " + data[i][1] + ",\n\n" +
                 "Permohonan cuti tahunan anda (" + payload.appId + ") dari " + formatTarikhDMY(toYMD(data[i][4])) + " hingga " + formatTarikhDMY(toYMD(data[i][5])) + " telah disemak oleh Pengetua.\n\n" +
                 "STATUS KELULUSAN : " + (lulus ? "DILULUSKAN" : "DITOLAK") + "\n" +
                 "CATATAN PENTADBIR : " + (catatan || "Tiada catatan tambahan.") + "\n\n" +
                 "Sila ambil maklum dan uruskan tugas serta jadual relief anda (jika berkaitan).\n\nTerima kasih.\n\nSistem e-Cuti SMIAAG";
      try { MailApp.sendEmail(emelGuru, "Status Permohonan Cuti e-Cuti SMIAAG: " + (lulus ? "DILULUSKAN" : "DITOLAK"), body); } catch(e) {}
    }
    return jsonResponse({ success: true, message: "Status dikemaskini & emel dihantar." });
  }
  return jsonResponse({ success: false, message: "Permohonan tidak dijumpai." });
}

