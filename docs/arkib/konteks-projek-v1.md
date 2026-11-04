> ARKIB: dokumen konteks versi 1 (sebelum v2). Kod di sini sudah lapuk, rujuk `public/` dan `apps-script/`. Pautan Publish-to-Web telah diredact.

# PROJECT CONTEXT: e-Cuti SMIAAG (Sistem Pengurusan Cuti Tahunan Guru)

## 1. Project Overview
* **Type:** Web Application (Single Page Application - SPA)
* **Tech Stack:** Google Apps Script (Backend), HTML5, Tailwind CSS, Vanilla JavaScript (Frontend), Google Sheets (Database), Google Drive (File Storage).
* **Architecture:** Serverless architecture using Google Apps Script `doGet` and `doPost` to serve the web app and act as a RESTful API.
* **Purpose:** A centralized system for school teachers (Guru) and staff (Staf) to apply for annual leave, and for the principal (Admin) to manage, approve, and track leave quotas.

## 2. Core Features & Business Logic
1. **Role-Based Access Control (RBAC):** `ADMIN` (Full access), `GURU` / `STAF` (Can only view own records and apply).
2. **Custom Authentication:** 
   * Token-based session management (`localStorage` + Google Sheets).
   * Secure Password Hashing (SHA-256 + User Salt + App Salt).
   * Forgot Password via Auto-Email (OTP with 15-min expiry).
   * Brute-force protection (Rate limiting via `CacheService`).
   * "Force Password Change" on first login.
3. **Leave Application Engine:**
   * Validates date overlapping and automatic day counting (excluding Sundays).
   * Compulsory PDF/Image attachment (Uploaded to Google Drive via Base64).
   * Dynamic Relief Teacher (Jadual Relief) table.
4. **Calendar & Quota System:**
   * Max 3 teachers allowed per day (Default).
   * Unlimited Quota Exception (Admin can set specific dates to bypass the 3-person limit).
   * Syncs real-time with Google Sheets.
5. **Admin Approval & Auto-Email:**
   * "First Come First Serve" sorting based on exact timestamp.
   * Dynamic leave balance calculation.
   * Auto-sends official approval/rejection emails via `MailApp.sendEmail()`.
6. **UI/UX:**
   * Tailwind CSS with Dark/Light Mode.
   * Mobile-first design (Bottom Navigation Bar for mobile).
   * Native Chart.js for analytics.

## 3. Database Schema (Google Sheets Tabs)
1. **`Users`**: `Emel`, `Nama`, `Role`, `PasswordHash`, `Salt`, `MestiTukarPassword`, `ResetKodHash`, `ResetLuput`, `Status`, `TarikhDaftar`
2. **`PermohonanCuti`**: `ID`, `NamaGuru`, `Emel`, `TarikhMohon`, `TarikhMula`, `TarikhAkhir`, `JumlahHari`, `SebabKategori`, `SebabButiran`, `Status`, `JadualRelief`, `NamaFailSurat`, `UlasanPengetua`
3. **`Sessions`**: `Token`, `Emel`, `Role`, `ExpiryTimestamp`
4. **`Log`**: `Masa`, `Emel`, `Tindakan`, `Status`
5. **`DATA_GURU`**: `NAMA_GURU`, `BIL_CUTI_TAHUNAN_BERBAKI` (Base quota for teachers).
6. **`TarikhTanpaHad`**: `Tarikh` (List of dates bypassing the quota).

---

## 4. Source Code

### A. Backend (`Code.gs`)
```javascript
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

function invalidateUserSessions(emel) {
  var sheet = pastikanTab("Sessions", ["Token", "Emel", "Role", "ExpiryTimestamp"]);
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][1].toString().toLowerCase() === emel.toLowerCase()) {
      sheet.deleteRow(i + 1);
    }
  }
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

function uploadFileToDrive(base64Data, fileName, mimeType) {
  try {
    if (!base64Data) return null;
    var folderName = "eCuti_Surat_Permohonan";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);
    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, mimeType, fileName);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (e) { return null; }
}

function doGet(e) { return ContentService.createTextOutput(JSON.stringify({ status: "online" })).setMimeType(ContentService.MimeType.JSON); }

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) return jsonResponse({ success: false, message: "Permintaan tidak sah." });
    var payload = JSON.parse(e.postData.contents);
    
    // 1. ACTION TANPA TOKEN (Public & Auth Flow)
    if (payload.action === "LOGIN") return handleLogin(payload);
    if (payload.action === "MINTA_KOD_RESET") return handleMintaKodReset(payload);
    if (payload.action === "SAHKAN_KOD_RESET") return handleSahkanKodReset(payload);
    if (payload.action === "SUBMIT_LEAVE") return handleSubmitLeave(payload);
    if (payload.action === "GET_UNLIMITED_DATES") return handleGetUnlimitedDates();
    
    // 2. CHECK TOKEN PUBLIC (Untuk Render Kalendar Secara Awam)
    if (payload.action === "GET_LEAVE_APPLICATIONS" && payload.token === "PUBLIC_GHOST") {
      return handleGetLeaveApplications(payload, { valid: true, role: "PUBLIC", emel: "public" });
    }

    // 3. ACTION WAJIB TOKEN (Bagi Pengguna Log Masuk)
    var sessionCheck = validateSession(payload.token);
    if (!sessionCheck.valid) return jsonResponse({ success: false, message: sessionCheck.error, forceLogout: true });

    switch (payload.action) {
      case "CHANGE_PASSWORD": return handleChangePassword(payload, sessionCheck);
      case "GET_LEAVE_APPLICATIONS": return handleGetLeaveApplications(payload, sessionCheck);
      case "APPROVE_LEAVE": return handleApproveLeave(payload, sessionCheck);
      case "ADD_UNLIMITED_DATE": return handleAddUnlimitedDate(payload, sessionCheck);
      case "REMOVE_UNLIMITED_DATE": return handleRemoveUnlimitedDate(payload, sessionCheck);
      case "GET_USERS": return handleGetUsers(payload, sessionCheck);
      case "ADD_USER": return handleAddUser(payload, sessionCheck);
      case "EDIT_USER": return handleEditUser(payload, sessionCheck);
      case "DELETE_USER": return handleDeleteUser(payload, sessionCheck);
      case "GET_DATA_GURU": return handleGetDataGuru();
      default: return jsonResponse({ success: false, message: "Aksi tidak dikenali." });
    }
  } catch (err) { return jsonResponse({ success: false, message: "Ralat Pelayan: " + err.toString() }); }
}

function jsonResponse(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

function handleLogin(payload) {
  var emel = (payload.emel || "").trim().toLowerCase();
  var password = payload.password || "";
  if (!emel || !password) return jsonResponse({ success: false, message: "Sila isi maklumat penuh." });
  if (!checkRateLimit(emel)) {
    logActivity(emel, "LOGIN", "FAILED_RATELIMIT");
    return jsonResponse({ success: false, message: "Terlalu banyak percubaan. Sila tunggu 5 minit." });
  }

  var sheetUsers = pastikanTab("Users", ["Emel", "Nama", "Role", "PasswordHash", "Salt", "MestiTukarPassword", "ResetKodHash", "ResetLuput", "Status", "TarikhDaftar"]);
  var data = sheetUsers.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === emel && data[i][8] === "ACTIVE") {
      var computedHash = hashPassword(password, data[i][4]);
      if (computedHash === data[i][3]) {
        clearRateLimit(emel);
        logActivity(emel, "LOGIN", "SUCCESS");
        var token = Utilities.getUuid() + "-" + new Date().getTime();
        pastikanTab("Sessions", ["Token", "Emel", "Role", "ExpiryTimestamp"]).appendRow([token, emel, data[i][2], new Date().getTime() + (SESSION_TTL_HOURS * 60 * 60 * 1000)]);
        return jsonResponse({ success: true, token: token, emel: emel, nama: data[i][1], role: data[i][2], mestiTukarPassword: (data[i][5] === true || data[i][5].toString().toUpperCase() === "TRUE") });
      }
    }
  }
  logActivity(emel, "LOGIN", "FAILED_WRONG_CREDENTIALS");
  return jsonResponse({ success: false, message: "Emel atau kata laluan tidak sah." });
}

function handleChangePassword(payload, sessionCheck) {
  if ((payload.newPassword || "").length < 6) return jsonResponse({ success: false, message: "Min 6 aksara." });
  var sheetUsers = pastikanTab("Users", ["Emel", "Nama", "Role", "PasswordHash", "Salt", "MestiTukarPassword", "ResetKodHash", "ResetLuput", "Status", "TarikhDaftar"]);
  var data = sheetUsers.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === sessionCheck.emel.toLowerCase()) {
      var newSalt = Utilities.getUuid();
      sheetUsers.getRange(i + 1, 4).setValue(hashPassword(payload.newPassword, newSalt)); 
      sheetUsers.getRange(i + 1, 5).setValue(newSalt); 
      sheetUsers.getRange(i + 1, 6).setValue(false); 
      
      invalidateUserSessions(sessionCheck.emel);
      logActivity(sessionCheck.emel, "TUKAR_PASSWORD", "SUCCESS");
      return jsonResponse({ success: true, message: "Kata laluan berjaya dikemaskini! Sila log masuk semula." });
    }
  }
  return jsonResponse({ success: false, message: "Ralat teknikal." });
}

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

function handleAddUser(payload, sessionCheck) {
  if (sessionCheck.role !== "ADMIN") return jsonResponse({ success: false, message: "Capaian ditolak." });
  if (!payload.emel || !payload.nama || !payload.role || !payload.tempPassword) return jsonResponse({ success: false, message: "Sila isi semua maklumat." });
  
  var emelBaru = payload.emel.toLowerCase().trim();
  var sheetUsers = pastikanTab("Users", ["Emel", "Nama", "Role", "PasswordHash", "Salt", "MestiTukarPassword", "ResetKodHash", "ResetLuput", "Status", "TarikhDaftar"]);
  var data = sheetUsers.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === emelBaru && data[i][8] === "ACTIVE") {
      return jsonResponse({ success: false, message: "Emel ini telah wujud dalam sistem." });
    }
  }

  var salt = Utilities.getUuid();
  var hash = hashPassword(payload.tempPassword, salt);
  sheetUsers.appendRow([sanitizeFormula(emelBaru), sanitizeFormula(payload.nama), sanitizeFormula(payload.role), hash, salt, true, "", "", "ACTIVE", new Date()]);

  try {
    MailApp.sendEmail(emelBaru, "[e-Cuti] Akaun Portal SMIAAG Telah Dicipta", 
      "Salam sejahtera,\n\nAkaun portal e-Cuti SMIAAG anda telah didaftarkan dengan peranan " + payload.role + ".\n\n" +
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
  
  var sheetUsers = pastikanTab("Users", ["Emel", "Nama", "Role", "PasswordHash", "Salt", "MestiTukarPassword", "ResetKodHash", "ResetLuput", "Status", "TarikhDaftar"]);
  var data = sheetUsers.getDataRange().getValues();
  
  var targetRow = -1;
  var currentRole = "";
  var totalAdmin = 0;
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][8] === "ACTIVE" && data[i][2] === "ADMIN") totalAdmin++;
    if (data[i][8] === "ACTIVE" && data[i][0].toString().toLowerCase() === targetEmel) {
      targetRow = i + 1;
      currentRole = data[i][2];
    }
  }
  
  if (targetRow === -1) return jsonResponse({ success: false, message: "Pengguna tidak dijumpai." });
  
  if (currentRole === "ADMIN" && newRole !== "ADMIN" && totalAdmin <= 1) {
    logActivity(sessionCheck.emel, "CUBA_UBAH_PERANAN_ADMIN_TERAKHIR", "BLOCKED");
    return jsonResponse({ success: false, message: "TINDAKAN DISEKAT: Anda tidak boleh mengubah peranan pentadbir tunggal yang terakhir di dalam sistem." });
  }
  
  if (newEmel !== targetEmel) {
    for (var j = 1; j < data.length; j++) {
      if (data[j][8] === "ACTIVE" && data[j][0].toString().toLowerCase() === newEmel) {
        return jsonResponse({ success: false, message: "Alamat emel baharu ini telah pun berdaftar di bawah pengguna lain." });
      }
    }
  }
  
  sheetUsers.getRange(targetRow, 1).setValue(sanitizeFormula(newEmel));
  sheetUsers.getRange(targetRow, 2).setValue(sanitizeFormula(newNama));
  sheetUsers.getRange(targetRow, 3).setValue(sanitizeFormula(newRole));
  
  if (newEmel !== targetEmel) {
    invalidateUserSessions(targetEmel);
  }
  
  logActivity(sessionCheck.emel, "KEMASKINI_PENGGUNA", targetEmel + " -> " + newEmel);
  return jsonResponse({ success: true, message: "Maklumat pengguna berjaya dikemaskini." });
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

function handleGetDataGuru() {
  var data = pastikanTab("DATA_GURU", ["NAMA_GURU", "BIL_CUTI_TAHUNAN_BERBAKI"]).getDataRange().getValues();
  var teachers = [];
  for (var i = 1; i < data.length; i++) { if (data[i][0]) teachers.push({ NAMA_GURU: data[i][0].toString(), BIL_CUTI_TAHUNAN_BERBAKI: data[i][1] || 0 }); }
  return jsonResponse({ success: true, data: teachers });
}

function handleGetLeaveApplications(payload, sessionCheck) {
  var sheet = pastikanTab("PermohonanCuti", ["ID", "NamaGuru", "Emel", "TarikhMohon", "TarikhMula", "TarikhAkhir", "JumlahHari", "SebabKategori", "SebabButiran", "Status", "JadualRelief", "NamaFailSurat", "UlasanPengetua"]);
  var data = sheet.getDataRange().getValues();
  var apps = [];
  var isAdmin = (sessionCheck && sessionCheck.role === "ADMIN");
  var userEmail = (sessionCheck && sessionCheck.emel) ? sessionCheck.emel.toLowerCase() : "";

  for (var i = 1; i < data.length; i++) {
    var status = data[i][9].toString();
    var emelPemohon = data[i][2].toString().toLowerCase();
    
    if (isAdmin || status === "APPROVED" || emelPemohon === userEmail) {
      var reliefParsed = [];
      try { reliefParsed = JSON.parse(data[i][10]); } catch(e){}
      
      var reasonCat = data[i][7], reasonDet = data[i][8], attach = data[i][11], notes = data[i][12];

      if (!isAdmin && emelPemohon !== userEmail) {
        reasonCat = "Sulit";
        reasonDet = "Maklumat Terhad";
        attach = "Tiada Fail";
        notes = "";
        reliefParsed = [];
      }

      apps.push({
        id: data[i][0], teacherName: data[i][1], teacherEmail: data[i][2], appDate: data[i][3],
        startDate: data[i][4], endDate: data[i][5], totalDays: data[i][6],
        reasonCategory: reasonCat, reasonDetails: reasonDet, status: status,
        relief: reliefParsed, attachmentName: attach, pengetuaNotes: notes
      });
    }
  }
  return jsonResponse({ success: true, data: apps });
}

function handleSubmitLeave(payload) {
  var app = payload.application;
  if (!app || !app.teacherName) return jsonResponse({ success: false, message: "Borang tidak lengkap." });
  var msiaTime = Utilities.formatDate(new Date(), "Asia/Kuala_Lumpur", "dd/MM/yyyy h.mma").toLowerCase();
  var finalAttachmentData = app.fileName || "Tiada_Fail.pdf";
  if (app.fileBase64) {
    var fileUrl = uploadFileToDrive(app.fileBase64, app.fileName, app.fileMimeType);
    if (fileUrl) finalAttachmentData = JSON.stringify({ name: app.fileName, url: fileUrl });
  }

  var sheet = pastikanTab("PermohonanCuti", ["ID", "NamaGuru", "Emel", "TarikhMohon", "TarikhMula", "TarikhAkhir", "JumlahHari", "SebabKategori", "SebabButiran", "Status", "JadualRelief", "NamaFailSurat", "UlasanPengetua"]);
  var newId = "CT-" + new Date().getFullYear() + "-" + ("00" + (sheet.getLastRow())).slice(-3);

  sheet.appendRow([ newId, sanitizeFormula(app.teacherName), sanitizeFormula(app.teacherEmail), msiaTime, sanitizeFormula(app.startDate), sanitizeFormula(app.endDate), app.totalDays, sanitizeFormula(app.reasonCategory), sanitizeFormula(app.reasonDetails), "PENDING", JSON.stringify(app.relief || []), sanitizeFormula(finalAttachmentData), "" ]);
  return jsonResponse({ success: true, message: "Permohonan berjaya!", id: newId });
}

function handleApproveLeave(payload, sessionCheck) {
  if (sessionCheck.role !== "ADMIN") return jsonResponse({ success: false, message: "Hanya admin boleh meluluskan cuti." });
  var sheet = pastikanTab("PermohonanCuti", ["ID", "NamaGuru", "Emel", "TarikhMohon", "TarikhMula", "TarikhAkhir", "JumlahHari", "SebabKategori", "SebabButiran", "Status", "JadualRelief", "NamaFailSurat", "UlasanPengetua"]);
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === payload.appId) {
      sheet.getRange(i + 1, 10).setValue(payload.newStatus);
      sheet.getRange(i + 1, 13).setValue(sanitizeFormula(payload.notes || ""));
      
      var emelGuru = data[i][2];
      if (emelGuru && emelGuru.indexOf("@") > -1) {
        var subject = "Status Permohonan Cuti e-Cuti SMIAAG: " + (payload.newStatus === "APPROVED" ? "DILULUSKAN" : "DITOLAK");
        var startDateStr = data[i][4], endDateStr = data[i][5];
        try { if (!isNaN(new Date(data[i][4]).getTime())) startDateStr = new Date(data[i][4]).toDateString();
              if (!isNaN(new Date(data[i][5]).getTime())) endDateStr = new Date(data[i][5]).toDateString(); } catch(e) {}

        var body = "Salam sejahtera " + data[i][1] + ",\n\n" +
                   "Permohonan cuti tahunan anda (" + payload.appId + ") dari tarikh " + startDateStr + " hingga " + endDateStr + " telah disemak oleh Pengetua.\n\n" +
                   "STATUS KELULUSAN : " + (payload.newStatus === "APPROVED" ? "DILULUSKAN" : "DITOLAK") + "\n" +
                   "CATATAN PENTADBIR : " + (payload.notes || "Tiada catatan tambahan.") + "\n\n" +
                   "Sila ambil maklum dan uruskan tugas serta jadual relief anda (jika berkaitan).\n\n" +
                   "Terima kasih.\n\n" +
                   "Sistem e-Cuti SMIAAG";
        try { MailApp.sendEmail(emelGuru, subject, body); } catch(e) {}
      }
      return jsonResponse({ success: true, message: "Status dikemaskini & emel dihantar." });
    }
  }
  return jsonResponse({ success: false, message: "Permohonan tidak dijumpai." });
}

function handleGetUnlimitedDates() {
  var data = pastikanTab("TarikhTanpaHad", ["Tarikh"]).getDataRange().getValues();
  var dates = [];
  for (var i = 1; i < data.length; i++) { 
    var val = data[i][0];
    if (!val) continue;
    
    if (val instanceof Date) {
      dates.push(Utilities.formatDate(val, "Asia/Kuala_Lumpur", "yyyy-MM-dd"));
    } else {
      dates.push(val.toString().trim());
    }
  }
  return jsonResponse({ success: true, data: dates });
}

function handleAddUnlimitedDate(payload, sessionCheck) {
  if (sessionCheck.role !== "ADMIN") return jsonResponse({ success: false, message: "Ditolak" });
  
  var datesToProcess = payload.dates || [];
  if (payload.date && datesToProcess.length === 0) datesToProcess = [payload.date];
  if (datesToProcess.length === 0) return jsonResponse({ success: false, message: "Tiada tarikh" });
  
  var sheet = pastikanTab("TarikhTanpaHad", ["Tarikh"]);
  for (var i=0; i<datesToProcess.length; i++) { 
    sheet.appendRow(["'" + datesToProcess[i]]); 
  }
  return jsonResponse({ success: true, message: "Tarikh ditambah" });
}

function handleRemoveUnlimitedDate(payload, sessionCheck) {
  if (sessionCheck.role !== "ADMIN") return jsonResponse({ success: false, message: "Ditolak" });
  var sheet = pastikanTab("TarikhTanpaHad", ["Tarikh"]);
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    var val = data[i][0];
    var strVal = (val instanceof Date) ? Utilities.formatDate(val, "Asia/Kuala_Lumpur", "yyyy-MM-dd") : val.toString().trim();
    
    if (strVal === payload.date) { 
      sheet.deleteRow(i + 1); 
      return jsonResponse({ success: true, message: "Dibuang" }); 
    }
  }
  return jsonResponse({ success: false, message: "Tidak dijumpai" });
}
```

### B. Frontend (`index.html`)
*(Note to Claude: Please treat the `GAS_URL` constant below as the Web App execution endpoint).*

```html
<!DOCTYPE html>
<html lang="ms" class="light">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>e-Cuti SMIAAG | Sistem Pengurusan Cuti Tahunan Guru</title>
  
  <script src="https://cdn.tailwindcss.com"></script>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@300;400;500;600;700;800&display=swap" rel="stylesheet">
  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.4.1/papaparse.min.js"></script>

  <script>
    tailwind.config = {
      darkMode: 'class',
      theme: { extend: { fontFamily: { sans: ['"Plus Jakarta Sans"', 'sans-serif'] } } }
    }
  </script>

  <style>
    .bg-aurora-light { background: #f8fafc; background-image: radial-gradient(at 0% 0%, rgba(192, 132, 252, 0.15) 0px, transparent 50%), radial-gradient(at 100% 0%, rgba(125, 211, 252, 0.2) 0px, transparent 50%), radial-gradient(at 50% 100%, rgba(251, 113, 133, 0.1) 0px, transparent 50%); background-attachment: fixed; }
    .bg-aurora-dark { background: #090d16; background-image: radial-gradient(at 10% 10%, rgba(139, 92, 246, 0.25) 0px, transparent 50%), radial-gradient(at 90% 10%, rgba(14, 165, 233, 0.25) 0px, transparent 50%), radial-gradient(at 50% 90%, rgba(236, 72, 153, 0.2) 0px, transparent 50%); background-attachment: fixed; }
    .glass-card { backdrop-filter: blur(16px); -webkit-backdrop-filter: blur(16px); background: rgba(255, 255, 255, 0.85); border: 1px solid rgba(226, 232, 240, 0.8); box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.04); }
    .dark .glass-card { background: rgba(15, 23, 42, 0.7); border: 1px solid rgba(255, 255, 255, 0.08); box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.37); }
    .tab-active { position: relative; }
    .tab-active::after { content: ''; position: absolute; bottom: -2px; left: 15%; width: 70%; height: 3px; background: linear-gradient(90deg, #0ea5e9, #8b5cf6, #ec4899); border-radius: 9999px; }
  </style>
</head>

<body class="bg-aurora-light text-slate-800 dark:text-slate-100 font-sans min-h-screen transition-colors duration-300 antialiased">

  <header class="sticky top-0 z-40 border-b border-slate-200/60 dark:border-slate-800/60 backdrop-blur-xl bg-white/80 dark:bg-slate-900/80">
    <div class="max-w-7xl mx-auto px-4 h-20 flex items-center justify-between gap-4">
      <div class="flex items-center gap-3">
        <div class="w-11 h-11 rounded-2xl bg-gradient-to-tr from-indigo-600 to-pink-500 p-0.5 shadow-md flex items-center justify-center">
          <div class="w-full h-full bg-slate-900 rounded-[14px] flex items-center justify-center"><i class="fa-solid fa-graduation-cap text-lg text-transparent bg-clip-text bg-gradient-to-r from-cyan-400 to-purple-400"></i></div>
        </div>
        <div>
          <h1 class="text-xl font-extrabold tracking-tight bg-gradient-to-r from-slate-900 via-indigo-950 to-slate-800 dark:from-white dark:to-cyan-200 bg-clip-text text-transparent">e-Cuti SMIAAG</h1>
          <p class="text-xs text-slate-500 font-medium">Sistem Cuti Tahunan Guru</p>
        </div>
      </div>
      <nav class="hidden lg:flex items-center gap-1 bg-slate-100/80 dark:bg-slate-800/50 p-1.5 rounded-2xl border">
        <button onclick="switchTab('calendar')" id="nav-calendar" class="nav-btn px-4 py-2 text-xs font-semibold rounded-xl flex items-center gap-2 text-indigo-600 tab-active"><i class="fa-solid fa-calendar-days"></i> Kalendar Bulanan</button>
        <button onclick="switchTab('form')" id="nav-form" class="nav-btn px-4 py-2 text-xs font-semibold rounded-xl flex items-center gap-2 text-slate-600"><i class="fa-solid fa-paper-plane"></i> Borang Permohonan</button>
        <button onclick="switchTab('admin')" id="nav-admin" class="nav-btn px-4 py-2 text-xs font-semibold rounded-xl flex items-center gap-2 text-slate-600 relative"><i class="fa-solid fa-user-shield"></i> Portal Pengguna<span id="nav-admin-badge" class="hidden absolute -top-1 -right-1 flex h-3 w-3"><span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-rose-400 opacity-75"></span><span class="relative inline-flex rounded-full h-3 w-3 bg-rose-500"></span></span></button>
      </nav>
      <div class="flex items-center gap-2">
        <button onclick="toggleTheme()" class="w-10 h-10 rounded-xl bg-slate-100 dark:bg-slate-800 text-slate-600 flex items-center justify-center"><i class="fa-solid fa-moon dark:hidden"></i><i class="fa-solid fa-sun hidden dark:block text-amber-400"></i></button>
      </div>
    </div>
  </header>

  <main class="max-w-7xl mx-auto px-4 py-8 pb-24 lg:pb-8">
    
    <!-- TOAST NOTIFICATION -->
    <div id="toast" class="fixed top-24 right-4 z-50 transform translate-x-full transition-transform duration-300 max-w-sm w-full glass-card p-4 rounded-2xl shadow-2xl border-l-4 flex items-start gap-3">
      <div class="w-8 h-8 rounded-xl bg-indigo-500/20 text-indigo-500 flex items-center justify-center shrink-0"><i class="fa-solid fa-circle-info"></i></div>
      <div class="flex-1"><h4 id="toast-title" class="text-sm font-bold">Mesej Sistem</h4><p id="toast-message" class="text-xs text-slate-500 mt-0.5">Sila tunggu...</p></div>
      <button type="button" onclick="hideToast()" class="text-slate-400 hover:text-rose-500 p-1 z-50 cursor-pointer"><i class="fa-solid fa-xmark text-lg"></i></button>
    </div>

    <!-- MODAL CONFIRMATION (Ganti windows.confirm) -->
    <div id="custom-confirm-modal" class="hidden fixed inset-0 z-[110] flex items-center justify-center bg-slate-900/60 backdrop-blur-sm">
      <div class="glass-card max-w-sm w-full p-6 rounded-3xl mx-4 shadow-2xl text-center border-t-4 border-rose-500">
        <i class="fa-solid fa-triangle-exclamation text-4xl text-rose-500 mb-3"></i>
        <h3 id="confirm-title" class="text-lg font-extrabold text-slate-800">Sahkan Tindakan</h3>
        <p id="confirm-message" class="text-xs text-slate-500 mb-6">Adakah anda pasti?</p>
        <div class="flex gap-3">
          <button onclick="closeConfirmModal()" class="flex-1 py-3 rounded-xl bg-slate-100 text-slate-600 font-bold text-xs shadow-sm hover:bg-slate-200">Batal</button>
          <button id="confirm-btn-yes" class="flex-1 py-3 rounded-xl bg-rose-600 text-white font-bold text-xs shadow-sm hover:bg-rose-700">Ya, Teruskan</button>
        </div>
      </div>
    </div>

    <!-- MODAL WAJIB TUKAR PASSWORD -->
    <div id="force-password-modal" class="hidden fixed inset-0 z-[100] flex items-center justify-center bg-slate-900/60 backdrop-blur-sm">
      <div class="glass-card max-w-sm w-full p-6 rounded-3xl mx-4 shadow-2xl border-t-4 border-amber-500 text-center">
        <i class="fa-solid fa-triangle-exclamation text-4xl text-amber-500 mb-3"></i>
        <h3 class="text-lg font-extrabold text-slate-800">Tukar Kata Laluan Wajib</h3>
        <p class="text-xs text-slate-500 mb-5">Untuk keselamatan, anda wajib menukar kata laluan sementara sebelum menggunakan sistem.</p>
        <div class="space-y-3">
          <input type="password" id="force-new-password" placeholder="Kata Laluan Baharu (Min 6 aksara)" class="w-full px-4 py-3 rounded-xl border text-xs focus:ring-2 outline-none">
          <button onclick="forceChangePassword()" class="w-full py-3 rounded-xl bg-amber-500 text-white font-bold text-xs shadow hover:bg-amber-600 transition-colors">Simpan Kata Laluan</button>
        </div>
      </div>
    </div>

    <!-- TAB 1: CALENDAR VIEW -->
    <section id="tab-calendar" class="tab-content space-y-4">
      <div class="glass-card rounded-2xl p-4 sm:p-5 flex flex-col md:flex-row justify-between gap-4">
        <div>
          <div class="flex items-center gap-2"><span class="w-2.5 h-2.5 rounded-full bg-emerald-500"></span><span class="text-[10px] font-bold uppercase tracking-wider text-emerald-600">Jadual Cuti Pendidik</span></div>
          <div class="flex items-center gap-2 mt-2">
            <select id="calendar-month-select" onchange="jumpToSelectedMonth()" class="appearance-none px-4 py-1.5 rounded-xl bg-white border text-sm font-extrabold text-slate-700 outline-none cursor-pointer shadow-sm focus:ring-2 focus:ring-indigo-100">
              <option value="0">Januari</option><option value="1">Februari</option><option value="2">Mac</option><option value="3">April</option><option value="4">Mei</option><option value="5">Jun</option><option value="6">Julai</option><option value="7">Ogos</option><option value="8">September</option><option value="9">Oktober</option><option value="10">November</option><option value="11">Disember</option>
            </select>
            <span class="text-lg font-black text-slate-800 dark:text-white" id="calendar-year-display"></span>
          </div>
        </div>
        <div class="flex items-center bg-slate-100 p-1 rounded-xl border h-fit shadow-sm">
          <button onclick="changeMonth(-1)" class="px-3 py-1.5 rounded-lg hover:bg-white text-slate-600 transition-colors"><i class="fa-solid fa-chevron-left text-xs"></i></button>
          <button onclick="goToCurrentMonth()" class="px-3 py-1.5 text-xs font-bold hover:bg-white text-slate-700 transition-colors">Kini</button>
          <button onclick="changeMonth(1)" class="px-3 py-1.5 rounded-lg hover:bg-white text-slate-600 transition-colors"><i class="fa-solid fa-chevron-right text-xs"></i></button>
        </div>
      </div>
      <div class="glass-card rounded-2xl p-2 sm:p-5 border-0">
        <div class="grid grid-cols-7 gap-1 sm:gap-2 mb-2 text-center text-[9px] sm:text-[10px] font-bold uppercase tracking-wider text-slate-400 border-b pb-2"><div class="text-rose-500">Ahd</div><div>Isn</div><div>Sel</div><div>Rab</div><div>Kha</div><div>Jum</div><div class="text-indigo-500">Sab</div></div>
        <div id="calendar-grid" class="grid grid-cols-7 gap-1 sm:gap-2 w-full"></div>
      </div>
    </section>

    <!-- TAB 2: FORM PERMOHONAN -->
    <section id="tab-form" class="tab-content hidden space-y-6">
      <div class="glass-card rounded-3xl p-6 sm:p-8 max-w-4xl mx-auto">
        <h2 class="text-xl font-extrabold tracking-tight border-b pb-4 mb-6">Borang Cuti Tahunan</h2>
        <form id="leave-application-form" onsubmit="handleLeaveSubmit(event)" class="space-y-8">
          <div>
            <h3 class="text-xs font-bold uppercase tracking-wider text-indigo-600 flex items-center gap-2 mb-4"><i class="fa-solid fa-user"></i> 1. Maklumat Pemohon</h3>
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label class="block text-xs font-bold text-slate-500 mb-1">Nama Guru *</label>
                <select id="form-teacher-select" required onchange="updateSelectedTeacherInfo()" class="w-full px-4 py-2.5 rounded-xl border text-xs outline-none focus:ring-2 focus:ring-indigo-100"><option value="">-- Pilih Guru --</option></select>
                <div id="form-teacher-info" class="hidden mt-2 p-3 rounded-xl bg-indigo-50 border flex justify-between items-center"><span class="text-xs font-bold text-indigo-800">Baki Terkini:</span><span id="form-teacher-balance" class="text-sm font-black text-indigo-600 bg-white rounded px-2">0 Hari</span></div>
              </div>
              <div>
                <label class="block text-xs font-bold text-slate-500 mb-1">Alamat Emel Guru *</label>
                <input type="email" id="form-teacher-email" required placeholder="Untuk makluman status kelulusan automatik" class="w-full px-4 py-2.5 rounded-xl border text-xs outline-none focus:ring-2 focus:ring-indigo-100">
              </div>
            </div>
          </div>
          <div>
            <h3 class="text-xs font-bold uppercase tracking-wider text-indigo-600 flex items-center gap-2 mb-4"><i class="fa-solid fa-calendar"></i> 2. Maklumat Cuti</h3>
            <div class="grid grid-cols-1 md:grid-cols-3 gap-4">
              <div><label class="block text-xs font-bold text-slate-500 mb-1">Tarikh Mula *</label><input type="date" id="form-start-date" required onchange="calculateLeaveDays()" class="w-full px-4 py-2.5 rounded-xl border text-xs focus:ring-2 focus:ring-indigo-100"></div>
              <div><label class="block text-xs font-bold text-slate-500 mb-1">Tarikh Akhir *</label><input type="date" id="form-end-date" required onchange="calculateLeaveDays()" class="w-full px-4 py-2.5 rounded-xl border text-xs focus:ring-2 focus:ring-indigo-100"></div>
              <div><label class="block text-xs font-bold text-slate-500 mb-1">Jumlah Hari (Kecuali Ahad)</label><input type="text" id="form-total-days" readonly placeholder="0 Hari" class="w-full px-4 py-2.5 rounded-xl bg-indigo-50 text-indigo-600 font-bold border text-xs outline-none"></div>
            </div>
            <div class="grid grid-cols-1 md:grid-cols-3 gap-4 mt-4">
              <div><label class="block text-xs font-bold text-slate-500 mb-1">Sebab *</label><select id="form-reason-category" required class="w-full px-4 py-2.5 rounded-xl border text-xs focus:ring-2 focus:ring-indigo-100"><option>Urusan Peribadi / Keluarga</option><option>Kecemasan & Kesihatan</option><option>Kursus / Bengkel</option><option>Lain-Lain</option></select></div>
              <div class="md:col-span-2"><label class="block text-xs font-bold text-slate-500 mb-1">Butiran Lanjut *</label><input type="text" id="form-reason-details" required placeholder="Cth: Menghadiri kenduri..." class="w-full px-4 py-2.5 rounded-xl border text-xs focus:ring-2 focus:ring-indigo-100"></div>
            </div>
          </div>
          <div>
            <div class="flex items-center justify-between mb-3"><h3 class="text-xs font-bold uppercase tracking-wider text-indigo-600 flex items-center gap-2"><i class="fa-solid fa-chalkboard-user"></i> 3. Jadual Relief</h3><button type="button" onclick="addReliefRow()" class="px-3 py-1.5 rounded-lg bg-indigo-100 text-indigo-600 text-[10px] font-bold shadow-sm hover:bg-indigo-200"><i class="fa-solid fa-plus"></i> Tambah</button></div>
            <div class="overflow-x-auto border rounded-xl"><table class="w-full text-left text-xs bg-slate-50"><thead class="bg-slate-100 text-slate-600 border-b"><tr><th class="p-2">Hari</th><th class="p-2">Masa</th><th class="p-2">Kelas</th><th class="p-2">Guru Ganti</th><th class="p-2"></th></tr></thead><tbody id="relief-table-body" class="bg-white divide-y divide-slate-100"></tbody></table></div>
          </div>
          <div>
            <h3 class="text-xs font-bold uppercase tracking-wider text-indigo-600 flex items-center gap-2 mb-2"><i class="fa-solid fa-paperclip"></i> 4. Lampiran Wajib</h3>
            <div class="border-2 border-dashed border-slate-300 hover:border-indigo-400 transition-colors rounded-2xl p-6 text-center bg-slate-50">
              <input type="file" id="form-attachment-file" accept=".pdf,.png,.jpg,.jpeg" required onchange="handleFileSelect(event)" class="hidden">
              <i class="fa-solid fa-cloud-arrow-up text-3xl text-indigo-500 mb-3"></i>
              <p class="text-xs font-bold text-slate-600" id="file-name-display">Muat naik Surat Rasmi (Max 5MB).</p>
              <button type="button" onclick="document.getElementById('form-attachment-file').click()" class="mt-3 px-5 py-2 rounded-xl bg-indigo-600 text-white text-xs font-bold shadow-sm hover:bg-indigo-700">Pilih Fail</button>
            </div>
          </div>
          <div class="pt-4 border-t flex justify-end"><button type="submit" class="w-full sm:w-auto px-6 py-3 rounded-xl bg-gradient-to-r from-indigo-600 to-purple-600 text-white text-xs font-extrabold shadow-md hover:shadow-lg hover:scale-[1.02] transition-all flex items-center justify-center gap-2"><i class="fa-solid fa-paper-plane"></i> Hantar Permohonan</button></div>
        </form>
      </div>
    </section>

    <!-- TAB 3: ADMIN PORTAL -->
    <section id="tab-admin" class="tab-content hidden space-y-6">
      
      <!-- KAD AUTH (LOGIN & LUPA PASSWORD) -->
      <div id="admin-auth-card" class="glass-card rounded-3xl p-8 max-w-md mx-auto shadow-lg border-t-4 border-t-indigo-500 relative overflow-hidden">
        
        <!-- PANDANGAN LOGIN -->
        <div id="login-view" class="text-center space-y-4">
          <div class="w-14 h-14 rounded-2xl bg-indigo-100 text-indigo-600 flex items-center justify-center text-2xl mx-auto shadow-inner"><i class="fa-solid fa-lock"></i></div>
          <h3 class="text-lg font-extrabold text-slate-800">Portal Pengguna</h3>
          <p class="text-xs text-slate-500 mb-2">Akses untuk pentadbir dan warga pendidik.</p>
          <div class="space-y-3 text-left">
            <input type="email" id="admin-login-email" placeholder="Alamat Emel" class="w-full px-4 py-3 rounded-xl border text-xs focus:ring-2 outline-none">
            <input type="password" id="admin-login-pass" placeholder="Kata Laluan" class="w-full px-4 py-3 rounded-xl border text-xs focus:ring-2 outline-none">
          </div>
          <div class="text-right">
            <button onclick="toggleAuthView('forgot')" class="text-[10px] text-indigo-600 hover:underline font-bold">Lupa Kata Laluan?</button>
          </div>
          <button onclick="performLogin()" class="w-full py-3 rounded-xl bg-indigo-600 text-white font-bold text-xs shadow-md mt-2 hover:bg-indigo-700 transition-colors">Log Masuk</button>
        </div>

        <!-- PANDANGAN LUPA PASSWORD -->
        <div id="forgot-view" class="hidden text-center space-y-4">
          <button onclick="toggleAuthView('login')" class="absolute top-4 left-4 text-slate-400 hover:text-slate-600"><i class="fa-solid fa-arrow-left"></i></button>
          <div class="w-14 h-14 rounded-2xl bg-purple-100 text-purple-600 flex items-center justify-center text-2xl mx-auto shadow-inner"><i class="fa-solid fa-key"></i></div>
          <h3 class="text-lg font-extrabold text-slate-800">Lupa Kata Laluan</h3>
          
          <div id="forgot-step-1" class="space-y-3 text-left">
            <p class="text-[10px] text-slate-500 text-center mb-4">Sila masukkan emel berdaftar anda. Kami akan hantar kod 6-digit untuk menetapkan semula kata laluan.</p>
            <input type="email" id="forgot-email" placeholder="Alamat Emel Berdaftar" class="w-full px-4 py-3 rounded-xl border text-xs outline-none">
            <button onclick="requestResetCode()" class="w-full py-3 rounded-xl bg-purple-600 text-white font-bold text-xs shadow-md hover:bg-purple-700 transition-colors">Minta Kod Akses</button>
          </div>

          <div id="forgot-step-2" class="hidden space-y-3 text-left">
            <p class="text-[10px] text-slate-500 text-center mb-4">Kod 6-digit telah dihantar ke emel anda (Sila semak Spam jika tiada). Kod sah selama 15 minit.</p>
            <input type="text" id="forgot-code" placeholder="Kod 6-Digit (Cth: 123456)" maxlength="6" class="w-full px-4 py-3 rounded-xl border text-xs font-mono text-center tracking-[0.5em] outline-none">
            <input type="password" id="forgot-new-pass" placeholder="Kata Laluan Baharu" class="w-full px-4 py-3 rounded-xl border text-xs outline-none mt-2">
            <button onclick="confirmResetCode()" class="w-full py-3 rounded-xl bg-emerald-600 text-white font-bold text-xs shadow-md mt-2 hover:bg-emerald-700 transition-colors">Sahkan & Tukar Password</button>
          </div>
        </div>
      </div>

      <div id="admin-main-panel" class="hidden space-y-6">
        <div class="glass-card rounded-2xl p-4 flex flex-wrap items-center justify-between gap-4 shadow-sm">
          <div class="flex items-center gap-3">
            <div class="w-9 h-9 rounded-xl bg-indigo-600 text-white flex items-center justify-center font-bold text-xs"><i class="fa-solid fa-circle-user"></i></div>
            <div>
              <h3 class="text-sm font-extrabold text-slate-800" id="portal-user-name">Pengguna</h3>
              <p class="text-[10px] text-slate-500 font-bold" id="portal-user-role">Role</p>
            </div>
          </div>
          <div class="flex items-center gap-2 flex-wrap">
            <div class="flex bg-slate-100 rounded-xl p-1 text-xs font-bold border flex-wrap" id="portal-nav-group">
              <!-- Butang Navbar Dimasukkan secara Dinamik Ikut Role -->
            </div>
            <button onclick="performLogout()" class="px-3 py-1.5 rounded-xl bg-rose-100 text-rose-500 font-bold text-xs hover:bg-rose-200 transition-colors"><i class="fa-solid fa-right-from-bracket"></i></button>
          </div>
        </div>

        <!-- SUB 1: PENGURUSAN REKOD / KELULUSAN -->
        <div id="admin-view-approvals" class="admin-subview space-y-4">
          <div class="flex flex-col sm:flex-row gap-2">
            <select id="admin-filter-status" onchange="renderApprovalsList()" class="px-3 py-2.5 rounded-xl bg-white border text-xs font-bold outline-none shadow-sm focus:ring-2 focus:ring-indigo-100">
              <option value="PENDING">Permohonan Baru / Dalam Semakan</option><option value="ALL">Semua Rekod Sejarah</option><option value="APPROVED">Telah Diluluskan</option><option value="REJECTED">Ditolak</option>
            </select>
            <select id="admin-filter-teacher" onchange="renderApprovalsList()" class="px-3 py-2.5 rounded-xl bg-white border text-xs font-bold flex-1 outline-none hidden admin-only-ui shadow-sm focus:ring-2 focus:ring-indigo-100"><option value="ALL">-- Carian Semua Guru --</option></select>
          </div>
          
          <div id="approvals-teacher-badge" class="hidden bg-indigo-50 border border-indigo-100 p-4 rounded-xl items-center justify-between shadow-sm"></div>
          
          <div id="approvals-list-container" class="space-y-4"></div>
        </div>

        <!-- SUB 2: BAKI CUTI KESELURUHAN (ADMIN SAHAJA) -->
        <div id="admin-view-teachers" class="admin-subview hidden space-y-4 admin-only-ui">
          <div class="glass-card rounded-2xl p-5 border"><table class="w-full text-left text-xs"><thead class="bg-slate-100 text-slate-600 border-b"><tr><th class="p-3">Nama Guru</th><th class="p-3 text-center">Baki Terkini</th></tr></thead><tbody id="admin-teachers-table-body" class="divide-y divide-slate-100"></tbody></table></div>
        </div>

        <!-- SUB 3: ANALITIK -->
        <div id="admin-view-analytics" class="admin-subview hidden space-y-4">
          <div class="grid grid-cols-2 md:grid-cols-3 gap-4">
            <div class="glass-card p-4 border-l-4 border-indigo-500 shadow-sm"><span class="text-[10px] font-bold text-slate-400">Total Permohonan</span><h3 class="text-xl font-black mt-1 text-slate-700" id="stat-total-apps">0</h3></div>
            <div class="glass-card p-4 border-l-4 border-emerald-500 shadow-sm"><span class="text-[10px] font-bold text-slate-400">Diluluskan</span><h3 class="text-xl font-black text-emerald-500 mt-1" id="stat-approved-apps">0</h3></div>
            <div class="glass-card p-4 border-l-4 border-purple-500 shadow-sm col-span-2 md:col-span-1"><span class="text-[10px] font-bold text-slate-400">Kadar Lulus</span><h3 class="text-xl font-black text-purple-500 mt-1" id="stat-avg-rate">0%</h3></div>
          </div>
          <div class="glass-card rounded-2xl p-5 h-64 border"><canvas id="reasonPieChart"></canvas></div>
        </div>

        <!-- SUB 4: KUOTA (ADMIN SAHAJA) -->
        <div id="admin-view-quota" class="admin-subview hidden space-y-4 admin-only-ui">
          <div class="glass-card rounded-2xl p-6 border-l-4 border-purple-500 shadow-sm">
            <h3 class="text-sm font-extrabold text-slate-800">Pengecualian Kuota (Tanpa Had)</h3>
            <p class="text-[10px] text-slate-500 mb-4">Lepaskan limit maksimum 3 orang guru sehari pada tarikh yang terpilih.</p>
            
            <div class="flex flex-wrap items-center gap-2 mb-6 p-4 bg-slate-50 rounded-xl border border-slate-200">
              <div><label class="block text-[10px] font-bold text-slate-500 mb-1">Tarikh Mula</label><input type="date" id="admin-unlim-start" class="px-3 py-2 rounded-xl text-xs outline-none border focus:ring-2"></div>
              <div class="self-end pb-2"><span class="text-xs font-bold text-slate-400">hingga</span></div>
              <div><label class="block text-[10px] font-bold text-slate-500 mb-1">Tarikh Akhir (Pilihan)</label><input type="date" id="admin-unlim-end" class="px-3 py-2 rounded-xl text-xs outline-none border focus:ring-2"></div>
              <div class="self-end"><button onclick="addUnlimitedBulk()" class="px-5 py-2 rounded-xl bg-purple-600 text-white text-xs font-bold hover:bg-purple-700 shadow-sm transition-colors">Set Kuota Tanpa Had</button></div>
            </div>
            
            <div class="border rounded-xl bg-white p-4 max-w-sm mx-auto sm:mx-0 shadow-sm">
              <div class="flex items-center justify-between mb-4"><button onclick="changeQuotaMonth(-1)" class="p-2 bg-slate-100 rounded-lg hover:bg-slate-200"><i class="fa-solid fa-chevron-left text-xs"></i></button><h4 id="quota-calendar-title" class="font-bold text-sm text-slate-700">Bulan</h4><button onclick="changeQuotaMonth(1)" class="p-2 bg-slate-100 rounded-lg hover:bg-slate-200"><i class="fa-solid fa-chevron-right text-xs"></i></button></div>
              <div class="grid grid-cols-7 gap-1 text-center text-[10px] font-bold text-slate-400 mb-2 border-b pb-2"><div class="text-rose-500">Ahd</div><div>Isn</div><div>Sel</div><div>Rab</div><div>Kha</div><div>Jum</div><div class="text-indigo-500">Sab</div></div>
              <div id="quota-calendar-grid" class="grid grid-cols-7 gap-1 sm:gap-2"></div>
              <div class="mt-4 flex items-center justify-center gap-4 text-[9px] font-bold text-slate-500"><div class="flex items-center gap-1"><span class="w-3 h-3 rounded bg-purple-500"></span> Tanpa Had</div><div class="flex items-center gap-1"><span class="w-3 h-3 rounded bg-slate-100 border"></span> Biasa (Max 3)</div></div>
            </div>
          </div>
        </div>

        <!-- SUB 5: SETTINGS (ADMIN + SEMUA USER) -->
        <div id="admin-view-settings" class="admin-subview hidden space-y-6">
          
          <!-- TUKAR PASSWORD SENDIRI (SEMUA USER BOLEH) -->
          <div class="glass-card rounded-2xl p-6 border-l-4 border-indigo-500 shadow-sm">
            <h3 class="text-sm font-extrabold text-slate-800 mb-1"><i class="fa-solid fa-key mr-2"></i>Tukar Kata Laluan</h3>
            <p class="text-[10px] text-slate-500 mb-4">Penukaran kata laluan akan memaksa log keluar automatik dari peranti lain.</p>
            <div class="flex flex-col sm:flex-row gap-3">
              <input type="password" id="setting-new-pass" placeholder="Kata Laluan Baharu (Min 6 Aksara)" class="flex-1 px-4 py-2.5 rounded-xl border text-xs outline-none focus:ring-2 focus:ring-indigo-100">
              <button onclick="changeOwnPassword()" class="px-5 py-2.5 rounded-xl bg-indigo-600 text-white font-bold text-xs shadow-sm hover:bg-indigo-700 transition-colors">Simpan Kata Laluan</button>
            </div>
          </div>

          <!-- PENGURUSAN PENGGUNA (ADMIN SAHAJA) -->
          <div class="glass-card rounded-2xl p-6 border-l-4 border-emerald-500 admin-only-ui shadow-sm">
            <h3 class="text-sm font-extrabold text-slate-800 mb-1"><i class="fa-solid fa-users-gear mr-2"></i>Pengurusan Akses Pengguna</h3>
            <p class="text-[10px] text-slate-500 mb-4">Tambah atau kemaskini pengguna untuk sistem ini. Data guru diselaraskan dengan Google Sheet.</p>
            
            <datalist id="teacher-names-list"></datalist>

            <div class="bg-slate-50 p-4 rounded-xl border border-slate-200 mb-6 flex flex-col md:flex-row gap-3 items-end shadow-inner">
              <div class="flex-1 w-full"><label class="text-[10px] font-bold text-slate-500">Alamat Emel</label><input type="email" id="set-add-email" class="w-full px-3 py-2.5 rounded-lg border text-xs focus:ring-2 focus:ring-emerald-100 outline-none"></div>
              <div class="flex-1 w-full"><label class="text-[10px] font-bold text-slate-500">Nama Penuh</label><input type="text" id="set-add-nama" list="teacher-names-list" placeholder="Taip / pilih dari senarai" class="w-full px-3 py-2.5 rounded-lg border text-xs focus:ring-2 focus:ring-emerald-100 outline-none"></div>
              <div class="w-full md:w-32"><label class="text-[10px] font-bold text-slate-500">Peranan (Role)</label><select id="set-add-role" class="w-full px-3 py-2.5 rounded-lg border text-xs focus:ring-2 focus:ring-emerald-100 outline-none"><option value="GURU">GURU</option><option value="STAF">STAF</option><option value="ADMIN">ADMIN</option></select></div>
              <div class="flex-1 w-full"><label class="text-[10px] font-bold text-slate-500">Kata Laluan Sementara</label><input type="text" id="set-add-pass" class="w-full px-3 py-2.5 rounded-lg border text-xs focus:ring-2 focus:ring-emerald-100 outline-none" value="TukarSaya123"></div>
              <button onclick="addNewUser()" class="w-full md:w-auto px-5 py-2.5 bg-emerald-600 text-white font-bold text-xs rounded-lg whitespace-nowrap shadow-sm hover:bg-emerald-700 transition-colors">Tambah User</button>
            </div>

            <div class="overflow-x-auto border rounded-xl shadow-sm"><table class="w-full text-left text-xs"><thead class="bg-slate-100 text-slate-600"><tr><th class="p-3">Emel</th><th class="p-3">Nama</th><th class="p-3">Peranan</th><th class="p-3 text-right">Tindakan</th></tr></thead><tbody id="settings-users-table" class="bg-white divide-y divide-slate-100"></tbody></table></div>
          </div>
        </div>
      </div>
      
      <!-- MODAL KEMASKINI PENGGUNA -->
      <div id="edit-user-modal" class="hidden fixed inset-0 z-[110] flex items-center justify-center bg-slate-900/60 backdrop-blur-sm">
        <div class="glass-card max-w-sm w-full p-6 rounded-3xl mx-4 shadow-2xl border-t-4 border-indigo-500">
          <h3 class="text-lg font-extrabold text-slate-800 mb-1">Kemaskini Pengguna</h3>
          <p class="text-xs text-slate-500 mb-5">Ubah maklumat akses pengguna ini.</p>
          <div class="space-y-4">
            <input type="hidden" id="edit-original-email">
            <div><label class="block text-[10px] font-bold text-slate-500 mb-1">Alamat Emel</label><input type="email" id="edit-user-email" class="w-full px-3 py-2.5 rounded-xl border text-xs outline-none focus:ring-2"></div>
            <div><label class="block text-[10px] font-bold text-slate-500 mb-1">Nama Penuh</label><input type="text" id="edit-user-nama" list="teacher-names-list" class="w-full px-3 py-2.5 rounded-xl border text-xs outline-none focus:ring-2"></div>
            <div><label class="block text-[10px] font-bold text-slate-500 mb-1">Peranan</label><select id="edit-user-role" class="w-full px-3 py-2.5 rounded-xl border text-xs outline-none focus:ring-2"><option value="GURU">GURU</option><option value="STAF">STAF</option><option value="ADMIN">ADMIN</option></select></div>
            <div class="flex gap-3 pt-2">
              <button onclick="closeEditUserModal()" class="flex-1 py-2.5 rounded-xl bg-slate-100 text-slate-600 font-bold text-xs shadow-sm hover:bg-slate-200">Batal</button>
              <button onclick="submitEditUser()" class="flex-1 py-2.5 rounded-xl bg-indigo-600 text-white font-bold text-xs shadow-sm hover:bg-indigo-700">Simpan Perubahan</button>
            </div>
          </div>
        </div>
      </div>

    </section>
  </main>

  <!-- BOTTOM NAVIGATION FOR MOBILE -->
  <nav class="lg:hidden fixed bottom-0 left-0 w-full bg-white/90 backdrop-blur-xl border-t border-slate-200 z-50 flex justify-around items-center p-2 pb-safe shadow-[0_-4px_20px_-10px_rgba(0,0,0,0.1)]">
    <button onclick="switchTab('calendar')" id="mob-nav-calendar" class="mob-nav-btn flex-1 flex flex-col items-center justify-center p-2 text-indigo-600">
      <i class="fa-solid fa-calendar-days text-xl mb-1"></i>
      <span class="text-[9px] font-bold">Kalendar</span>
    </button>
    <button onclick="switchTab('form')" id="mob-nav-form" class="mob-nav-btn flex-1 flex flex-col items-center justify-center p-2 text-slate-400 hover:text-slate-600">
      <i class="fa-solid fa-paper-plane text-xl mb-1"></i>
      <span class="text-[9px] font-bold">Mohon</span>
    </button>
    <button onclick="switchTab('admin')" id="mob-nav-admin" class="mob-nav-btn flex-1 flex flex-col items-center justify-center p-2 text-slate-400 hover:text-slate-600 relative">
      <i class="fa-solid fa-user-shield text-xl mb-1"></i>
      <span class="text-[9px] font-bold">Portal</span>
      <span id="mob-nav-admin-badge" class="hidden absolute top-2 right-6 flex h-2.5 w-2.5"><span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-rose-400 opacity-75"></span><span class="relative inline-flex rounded-full h-2.5 w-2.5 bg-rose-500"></span></span>
    </button>
  </nav>

  <style>
    @supports (padding-bottom: env(safe-area-inset-bottom)) {
      .pb-safe { padding-bottom: calc(0.5rem + env(safe-area-inset-bottom)); }
    }
  </style>

  <script>
    // ====================================================================
    // ===== TAMPAL URL APPS SCRIPT DI SINI SELEPAS DEPLOY =====
    const GAS_URL = "https://script.google.com/macros/s/AKfycbyIQ15mj6bJN1dbQ2HyZ-L03GlGqjd-xvaVwfhPB5HUp8MmLzDZZE0vNMsV6jarca90ZQ/exec"; 
    // ====================================================================

    // STATE VARIABLES
    let teachersData = [], leaveApplications = [], unlimitedDates = [];
    let currentCalendarDate = new Date(), quotaCalendarDate = new Date();
    let selectedFileBase64 = null, selectedFileMime = null, selectedFileName = null;
    let reasonChartInstance = null;
    let currentConfirmCallback = null;
    
    // SESSION VARIABLES
    let currentSessionToken = localStorage.getItem('smiaag_admin_token');
    let currentEmel = localStorage.getItem('smiaag_user_email');
    let currentRole = localStorage.getItem('smiaag_user_role') || "GURU";
    let resetEmailCache = "";

    // UTILITIES
    const escapeHTML = str => !str ? '' : String(str).replace(/[&<>'"]/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[m]));
    function formatTarikhBM(isoStr) { try { const d = new Date(isoStr); if (isNaN(d)) return escapeHTML(isoStr); return `${String(d.getDate()).padStart(2,'0')}/${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`; } catch(e) { return escapeHTML(isoStr); } }

    function showToast(title, msg, type = 'info') {
      document.getElementById('toast-title').innerText = title; document.getElementById('toast-message').innerText = msg;
      const t = document.getElementById('toast');
      t.className = `fixed top-24 right-4 z-50 transform transition-transform duration-300 max-w-sm w-full glass-card p-4 rounded-2xl shadow-2xl border-l-4 ${type === 'success' ? 'border-emerald-500' : type === 'error' ? 'border-rose-500' : 'border-indigo-500'} flex items-start gap-3 translate-x-0`;
      setTimeout(hideToast, 5000);
    }
    function hideToast() { const t = document.getElementById('toast'); if(t) { t.classList.remove('translate-x-0'); t.classList.add('translate-x-full'); } }

    function showConfirmModal(title, message, onConfirm) {
      document.getElementById('confirm-title').innerText = title;
      document.getElementById('confirm-message').innerText = message;
      currentConfirmCallback = onConfirm;
      document.getElementById('custom-confirm-modal').classList.remove('hidden');
    }
    function closeConfirmModal() { document.getElementById('custom-confirm-modal').classList.add('hidden'); currentConfirmCallback = null; }
    document.getElementById('confirm-btn-yes').addEventListener('click', () => { if (currentConfirmCallback) currentConfirmCallback(); closeConfirmModal(); });

    async function apiPost(payload) {
      if(currentSessionToken && !payload.token && payload.action !== 'LOGIN' && !payload.action.includes('RESET')) payload.token = currentSessionToken;
      try {
        const res = await fetch(GAS_URL, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(payload) });
        const json = await res.json();
        if(json.forceLogout) { showToast("Sesi Tamat", json.message, "error"); performLogout(); }
        return json;
      } catch(e) { return { success: false, message: "Ralat rangkaian pelayan." }; }
    }

    document.addEventListener("DOMContentLoaded", () => {
      addReliefRow();
      fetchPublicData();
      if (currentSessionToken) prepareDashboard();
    });

    async function fetchPublicData() {
      try {
        const urlParams = new URLSearchParams(window.location.search);
        let directDataUrl = "[PAUTAN-PUBLISH-TO-WEB-DIREDACT]/pub?output=csv";
        Papa.parse(directDataUrl, { download: true, header: true, complete: function(results) {
          if (results && results.data && results.data.length > 0) {
            teachersData = results.data.filter(r => r.NAMA_GURU).map(r => ({ NAMA_GURU: r.NAMA_GURU.trim(), BIL_CUTI_TAHUNAN_BERBAKI: parseInt(r.BIL_CUTI_TAHUNAN_BERBAKI) || 0 }));
            populateTeacherDropdowns();
          }
        }});
      } catch (e) {}

      const pApps = apiPost({ action: "GET_LEAVE_APPLICATIONS", token: "PUBLIC_GHOST" }); 
      const pUnlim = apiPost({ action: "GET_UNLIMITED_DATES" });
      const [resApps, resUnlim] = await Promise.all([pApps, pUnlim]);
      if(resApps.success) leaveApplications = resApps.data;
      if(resUnlim.success) unlimitedDates = resUnlim.data;
      renderCalendar();
    }

    function getBakiSemasa(nama) {
      const t = teachersData.find(x => x.NAMA_GURU === nama);
      let asas = t ? parseInt(t.BIL_CUTI_TAHUNAN_BERBAKI) || 0 : 0;
      let guna = leaveApplications.filter(a => a.teacherName === nama && a.status === 'APPROVED').reduce((sum, a) => sum + (parseInt(a.totalDays)||0), 0);
      return asas - guna;
    }

    function populateTeacherDropdowns() {
      let opts = '<option value="">-- Pilih Guru --</option>' + teachersData.map(t => `<option value="${escapeHTML(t.NAMA_GURU)}">${escapeHTML(t.NAMA_GURU)}</option>`).join('');
      document.getElementById('form-teacher-select').innerHTML = opts;
      
      let datalistOpts = teachersData.map(t => `<option value="${escapeHTML(t.NAMA_GURU)}">`).join('');
      const dl = document.getElementById('teacher-names-list');
      if (dl) dl.innerHTML = datalistOpts;

      document.querySelectorAll('.relief-replacement').forEach(s => { const val = s.value; s.innerHTML = opts; s.value = val; });
      if(currentSessionToken && currentRole === 'ADMIN') {
        document.getElementById('admin-filter-teacher').innerHTML = '<option value="ALL">-- Carian Semua Guru --</option>' + teachersData.map(t => `<option value="${escapeHTML(t.NAMA_GURU)}">${escapeHTML(t.NAMA_GURU)}</option>`).join('');
        const tbody = document.getElementById('admin-teachers-table-body');
        if(tbody) tbody.innerHTML = teachersData.map(t => { let baki = getBakiSemasa(t.NAMA_GURU); return `<tr class="border-b"><td class="p-3 font-bold">${escapeHTML(t.NAMA_GURU)}</td><td class="p-3 text-center font-black ${baki<=0?'text-rose-500':'text-indigo-600'}">${baki} Hari</td></tr>`; }).join('');
      }
    }

    function updateSelectedTeacherInfo() {
      const name = document.getElementById('form-teacher-select').value;
      const infoBox = document.getElementById('form-teacher-info');
      const balSpan = document.getElementById('form-teacher-balance');
      if(!name) { infoBox.classList.add('hidden'); return; }
      let baki = getBakiSemasa(name); balSpan.innerText = baki + " Hari";
      balSpan.className = `text-sm font-black px-2 py-0.5 rounded shadow-sm bg-white ${baki<=0 ? 'text-rose-500' : 'text-indigo-600'}`;
      infoBox.classList.remove('hidden');
    }

    function renderCalendar() {
      const grid = document.getElementById('calendar-grid'); if (!grid) return; grid.innerHTML = '';
      const year = currentCalendarDate.getFullYear(), month = currentCalendarDate.getMonth();
      document.getElementById('calendar-month-select').value = month; document.getElementById('calendar-year-display').innerText = year;
      const firstDay = new Date(year, month, 1).getDay(), daysInMonth = new Date(year, month + 1, 0).getDate();
      for (let i = 0; i < firstDay; i++) { grid.innerHTML += `<div class="p-1 min-h-[70px] sm:min-h-[90px] rounded-xl bg-slate-100/30 opacity-30"></div>`; }

      for (let day = 1; day <= daysInMonth; day++) {
        const dObj = new Date(year, month, day); dObj.setHours(0,0,0,0);
        const dayOfWeek = dObj.getDay(), yyyymmdd = `${year}-${String(month+1).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
        let approvedApps = leaveApplications.filter(a => a.status === 'APPROVED' && (new Date(a.startDate).setHours(0,0,0,0) <= dObj && new Date(a.endDate).setHours(0,0,0,0) >= dObj));
        let isSunday = (dayOfWeek === 0), isUnlimited = unlimitedDates.includes(yyyymmdd), count = approvedApps.length, isFull = (!isUnlimited && count >= 3);
        
        let headerHtml = `<div class="flex items-center justify-between"><span class="text-xs sm:text-sm font-black ${isSunday?'text-rose-400':(dayOfWeek===6?'text-indigo-500':'text-slate-700')}">${day}</span>`;
        if(isSunday) headerHtml += `<span class="hidden sm:inline text-[9px] text-slate-400 font-bold">Cuti</span></div>`;
        else if(isUnlimited) headerHtml += `<span class="px-1 py-0.5 rounded text-[8px] sm:text-[9px] font-bold bg-purple-100 text-purple-700">0/∞</span></div>`;
        else if(!isFull) headerHtml += `<span class="px-1 py-0.5 rounded text-[8px] sm:text-[9px] font-bold bg-emerald-100 text-emerald-700">${count}/3<span class="hidden sm:inline"> Sedia</span></span></div>`;
        else headerHtml += `<span class="px-1 py-0.5 rounded text-[8px] sm:text-[9px] font-bold bg-rose-100 text-rose-700">PENUH</span></div>`;
        
        let bodyHtml = `<div class="my-1 space-y-0.5 sm:space-y-1 flex-1">`;
        if(isUnlimited && count > 0) bodyHtml += `<div class="px-1 py-0.5 mt-1 rounded bg-purple-50 text-purple-700 text-[8px] sm:text-[10px] font-extrabold flex items-center justify-center border border-purple-100 shadow-sm">${count} Cuti</div>`;
        else if(!isUnlimited) bodyHtml += approvedApps.map(a => `<div class="px-1 py-0.5 rounded bg-indigo-50 text-indigo-700 text-[8px] sm:text-[9px] font-bold truncate border shadow-sm"><i class="fa-solid fa-check text-emerald-500 text-[7px] mr-0.5"></i>${escapeHTML(a.teacherName.split(' ')[0])}</div>`).join('');
        bodyHtml += `</div>`;

        let btnHtml = (isSunday || isFull) ? "" : `<button onclick="prefillFormDate('${yyyymmdd}'); event.stopPropagation();" class="mt-auto w-full py-1 sm:py-1.5 rounded-md sm:rounded-lg ${isUnlimited?'bg-gradient-to-r from-purple-500 to-indigo-500':'bg-indigo-600 hover:bg-indigo-700'} text-white font-bold text-[9px] sm:text-[10px] shadow-sm flex items-center justify-center gap-1 transition-all hover:scale-[1.02]"><i class="fa-solid fa-plus"></i><span class="hidden sm:inline"> Mohon</span></button>`;
        const bgClass = isSunday ? "bg-slate-100/50 opacity-60" : isFull ? "bg-rose-50/40 border-rose-200" : isUnlimited ? "bg-purple-50/20 border-purple-200 cursor-pointer hover:shadow-md" : "bg-white hover:border-indigo-400 cursor-pointer hover:shadow-md";
        grid.innerHTML += `<div class="p-1 sm:p-2 min-h-[85px] sm:min-h-[105px] rounded-lg sm:rounded-xl border flex flex-col justify-between transition-all ${bgClass}" ${(!isSunday&&!isFull)?`onclick="prefillFormDate('${yyyymmdd}')"`:''}>${headerHtml}${bodyHtml}${btnHtml}</div>`;
      }
    }

    function changeMonth(o) { currentCalendarDate.setMonth(currentCalendarDate.getMonth() + o); renderCalendar(); }
    function jumpToSelectedMonth() { currentCalendarDate.setMonth(parseInt(document.getElementById('calendar-month-select').value)); renderCalendar(); }
    function goToCurrentMonth() { currentCalendarDate = new Date(); renderCalendar(); }
    function prefillFormDate(d) { switchTab('form'); document.getElementById('form-start-date').value = d; document.getElementById('form-end-date').value = d; calculateLeaveDays(); }
    
    function calculateLeaveDays() {
      const sVal = document.getElementById('form-start-date').value, eVal = document.getElementById('form-end-date').value, tInp = document.getElementById('form-total-days');
      if(!sVal || !eVal) { tInp.value = "0 Hari"; return; }
      let s = new Date(sVal), e = new Date(eVal), c=0;
      if(e<s) { tInp.value = "Ralat Tarikh"; return; }
      while(s<=e) { if(s.getDay()>=1 && s.getDay()<=6) c++; s.setDate(s.getDate()+1); }
      tInp.value = `${c} Hari`;
    }

    function addReliefRow() {
      const row = document.createElement('tr');
      row.innerHTML = `<td class="p-1"><select class="relief-day w-full p-2 border rounded-lg text-[10px] outline-none focus:ring-2"><option>Isnin</option><option>Selasa</option><option>Rabu</option><option>Khamis</option><option>Jumaat</option><option>Sabtu</option></select></td><td class="p-1"><input type="text" placeholder="Cth: 8AM-9AM" class="relief-time w-full p-2 border rounded-lg text-[10px] outline-none focus:ring-2"></td><td class="p-1"><input type="text" placeholder="Cth: 3 Bestari" class="relief-class w-full p-2 border rounded-lg text-[10px] outline-none focus:ring-2"></td><td class="p-1"><select class="relief-replacement w-full p-2 border rounded-lg text-[10px] outline-none focus:ring-2"><option value="">Pilih Guru</option>${teachersData.map(t=>`<option value="${escapeHTML(t.NAMA_GURU)}">${escapeHTML(t.NAMA_GURU)}</option>`).join('')}</select></td><td class="p-1 text-center"><button type="button" onclick="this.closest('tr').remove()" class="text-rose-500 hover:bg-rose-50 p-1.5 rounded-md"><i class="fa-solid fa-trash-can"></i></button></td>`;
      document.getElementById('relief-table-body').appendChild(row);
    }

    function handleFileSelect(e) {
      const f = e.target.files[0]; if(!f) return;
      if(f.size > 5*1024*1024) { showToast("Ralat Saiz", "Fail melebihi 5MB.", "error"); e.target.value=""; return; }
      selectedFileName = f.name; selectedFileMime = f.type;
      const r = new FileReader();
      r.onload = ev => { selectedFileBase64 = ev.target.result.split(',')[1]; document.getElementById('file-name-display').innerHTML = `<span class="text-indigo-600 font-bold"><i class="fa-solid fa-check-circle"></i> Terpilih: ${escapeHTML(f.name)}</span>`; };
      r.readAsDataURL(f);
    }

    async function handleLeaveSubmit(e) {
      e.preventDefault();
      const totalDays = parseInt(document.getElementById('form-total-days').value)||0;
      if(totalDays===0) return showToast("Ralat Pengiraan", "Jumlah hari sifar atau tiada hari bersekolah yang valid.", "error");
      if(!selectedFileBase64) return showToast("Sila Muat Naik Surat", "Lampiran rasmi (PDF/Imej) adalah diwajibkan.", "error");
      
      const relief = [];
      document.querySelectorAll('#relief-table-body tr').forEach(r => {
        const d = r.querySelector('.relief-day')?.value, t = r.querySelector('.relief-time')?.value, c = r.querySelector('.relief-class')?.value, rep = r.querySelector('.relief-replacement')?.value;
        if(c && rep) relief.push({ day: d, time: t, class: c, replacement: rep });
      });
      
      const payload = { action: "SUBMIT_LEAVE", application: { teacherName: document.getElementById('form-teacher-select').value, teacherEmail: document.getElementById('form-teacher-email').value, startDate: document.getElementById('form-start-date').value, endDate: document.getElementById('form-end-date').value, totalDays: totalDays, reasonCategory: document.getElementById('form-reason-category').value, reasonDetails: document.getElementById('form-reason-details').value, relief: relief, fileName: selectedFileName, fileMimeType: selectedFileMime, fileBase64: selectedFileBase64 } };
      
      showToast("Sedang Menghantar...", "Merekod permohonan dan memuat naik surat ke Google Drive...", "info");
      const res = await apiPost(payload);
      if(res.success) { showToast("Berjaya Dihantar", "Permohonan direkod. Sila tunggu kelulusan.", "success"); document.getElementById('leave-application-form').reset(); selectedFileBase64=null; document.getElementById('file-name-display').innerText="Muat naik Surat Rasmi (Max 5MB)."; document.getElementById('relief-table-body').innerHTML=''; document.getElementById('form-teacher-info').classList.add('hidden'); addReliefRow(); fetchPublicData(); switchTab('calendar'); }
      else showToast("Gagal", res.message, "error");
    }

    function toggleAuthView(view) {
      document.getElementById('login-view').classList.add('hidden'); document.getElementById('forgot-view').classList.add('hidden');
      document.getElementById(`${view}-view`).classList.remove('hidden');
      if(view === 'forgot') { document.getElementById('forgot-step-1').classList.remove('hidden'); document.getElementById('forgot-step-2').classList.add('hidden'); }
    }

    async function requestResetCode() {
      const email = document.getElementById('forgot-email').value; if(!email) return showToast("Ralat", "Masukkan emel berdaftar", "error");
      showToast("Memproses...", "Menyemak kelayakan dan menghantar emel...", "info");
      const res = await apiPost({ action: "MINTA_KOD_RESET", emel: email });
      if(res.success) {
        showToast("Status Mesej", res.message, "success");
        resetEmailCache = email;
        document.getElementById('forgot-step-1').classList.add('hidden');
        document.getElementById('forgot-step-2').classList.remove('hidden');
      }
    }

    async function confirmResetCode() {
      const kod = document.getElementById('forgot-code').value, newP = document.getElementById('forgot-new-pass').value;
      if(!kod || newP.length < 6) return showToast("Ralat", "Sila periksa format kod dan pastikan password minimum 6 aksara.", "error");
      showToast("Mengesahkan...", "Menukar kata laluan...", "info");
      const res = await apiPost({ action: "SAHKAN_KOD_RESET", emel: resetEmailCache, kod: kod, newPassword: newP });
      if(res.success) { showToast("Selesai", res.message, "success"); toggleAuthView('login'); }
      else showToast("Gagal", res.message, "error");
    }

    async function performLogin() {
      showToast("Log Masuk...", "Mengesahkan maklumat anda...", "info");
      const res = await apiPost({ action: "LOGIN", emel: document.getElementById('admin-login-email').value, password: document.getElementById('admin-login-pass').value });
      if(res.success) {
        currentSessionToken = res.token; currentEmel = res.emel; currentRole = res.role;
        localStorage.setItem('smiaag_admin_token', res.token); localStorage.setItem('smiaag_user_email', res.emel); localStorage.setItem('smiaag_user_role', res.role);
        showToast("Akses Dibenarkan", `Selamat kembali, ${res.nama}`, "success");
        if(res.mestiTukarPassword) { document.getElementById('force-password-modal').classList.remove('hidden'); }
        else { prepareDashboard(); }
      } else showToast("Akses Ditolak", res.message, "error");
    }

    async function forceChangePassword() {
      const newP = document.getElementById('force-new-password').value; if(newP.length<6) return showToast("Ralat", "Min 6 aksara", "error");
      showToast("Menyimpan...", "Menukar kata laluan wajib...", "info");
      const res = await apiPost({ action: "CHANGE_PASSWORD", newPassword: newP });
      if(res.success) { showToast("Berjaya", "Sila log masuk semula menggunakan kata laluan baharu.", "success"); document.getElementById('force-password-modal').classList.add('hidden'); performLogout(); }
      else showToast("Gagal", res.message, "error");
    }

    function performLogout() {
      currentSessionToken = null; localStorage.removeItem('smiaag_admin_token');
      document.getElementById('admin-main-panel').classList.add('hidden'); document.getElementById('admin-auth-card').classList.remove('hidden');
      switchTab('calendar'); showToast("Sesi Ditamatkan", "Anda telah dilog keluar dengan selamat.", "info");
    }

    async function prepareDashboard() {
      document.getElementById('admin-auth-card').classList.add('hidden');
      document.getElementById('admin-main-panel').classList.remove('hidden');
      document.getElementById('portal-user-name').innerText = currentEmel;
      document.getElementById('portal-user-role').innerText = "PERANAN: " + currentRole;

      document.querySelectorAll('.admin-only-ui').forEach(el => { el.style.display = (currentRole === 'ADMIN') ? '' : 'none'; });
      
      const navGroup = document.getElementById('portal-nav-group');
      if (currentRole === 'ADMIN') {
        navGroup.innerHTML = `<button onclick="switchAdminSubtab('approvals')" id="subnav-approvals" class="px-3 py-1.5 rounded-lg">Kelulusan</button><button onclick="switchAdminSubtab('teachers')" id="subnav-teachers" class="px-3 py-1.5 rounded-lg">Baki Cuti</button><button onclick="switchAdminSubtab('analytics')" id="subnav-analytics" class="px-3 py-1.5 rounded-lg">Analisis</button><button onclick="switchAdminSubtab('quota')" id="subnav-quota" class="px-3 py-1.5 rounded-lg">Kuota</button><button onclick="switchAdminSubtab('settings')" id="subnav-settings" class="px-3 py-1.5 rounded-lg"><i class="fa-solid fa-gear"></i></button>`;
      } else {
        navGroup.innerHTML = `<button onclick="switchAdminSubtab('approvals')" id="subnav-approvals" class="px-3 py-1.5 rounded-lg">Rekod Cuti Saya</button><button onclick="switchAdminSubtab('analytics')" id="subnav-analytics" class="px-3 py-1.5 rounded-lg">Analisis</button><button onclick="switchAdminSubtab('settings')" id="subnav-settings" class="px-3 py-1.5 rounded-lg"><i class="fa-solid fa-gear"></i></button>`;
      }

      showToast("Menyegerak Data...", "Mengambil rekod terkini...", "info");
      const res = await apiPost({ action: "GET_LEAVE_APPLICATIONS" });
      if(res.success) { leaveApplications = res.data; populateTeacherDropdowns(); switchAdminSubtab('approvals'); }
      if(currentRole === 'ADMIN') loadUsersForAdmin();
    }

    function switchAdminSubtab(sub) {
      document.querySelectorAll('.admin-subview').forEach(e => e.classList.add('hidden'));
      document.querySelectorAll('#portal-nav-group button').forEach(b => b.className = "px-3 py-1.5 rounded-lg text-slate-600 hover:bg-white transition-colors");
      const btn = document.getElementById(`subnav-${sub}`);
      if (btn) btn.className = "px-3 py-1.5 rounded-lg bg-indigo-600 text-white shadow-sm";
      const view = document.getElementById(`admin-view-${sub}`);
      if(view) view.classList.remove('hidden');
      
      if(sub==='approvals') renderApprovalsList();
      if(sub==='analytics') initAnalytics();
      if(sub==='quota') { apiPost({action:"GET_UNLIMITED_DATES"}).then(r=>{ if(r.success) unlimitedDates=r.data; renderAdminQuotaCalendar();}); }
    }

    async function changeOwnPassword() {
      const newP = document.getElementById('setting-new-pass').value; if(newP.length<6) return showToast("Ralat", "Min 6 aksara", "error");
      const res = await apiPost({ action: "CHANGE_PASSWORD", newPassword: newP });
      if(res.success) { showToast("Berjaya", res.message, "success"); performLogout(); }
      else showToast("Gagal", res.message, "error");
    }

    async function loadUsersForAdmin() {
      const res = await apiPost({ action: "GET_USERS" });
      if(res.success) {
        document.getElementById('settings-users-table').innerHTML = res.data.map(u => `
          <tr class="hover:bg-slate-50 transition-colors"><td class="p-3 font-bold">${escapeHTML(u.emel)}</td><td class="p-3">${escapeHTML(u.nama)}</td><td class="p-3"><span class="bg-indigo-100 text-indigo-700 px-2 py-1 rounded text-[10px] font-bold shadow-sm">${u.role}</span></td>
          <td class="p-3 text-right">
            <button onclick="openEditUserModal('${escapeHTML(u.emel)}', '${escapeHTML(u.nama)}', '${u.role}')" class="text-amber-500 hover:bg-amber-100 p-2 rounded shadow-sm transition-colors"><i class="fa-solid fa-pen-to-square"></i></button>
            <button onclick="deleteUser('${escapeHTML(u.emel)}')" class="text-rose-500 hover:bg-rose-100 p-2 rounded ml-1 shadow-sm transition-colors"><i class="fa-solid fa-trash-can"></i></button>
          </td></tr>
        `).join('');
      }
    }

    async function addNewUser() {
      const e = document.getElementById('set-add-email').value, n = document.getElementById('set-add-nama').value, r = document.getElementById('set-add-role').value, p = document.getElementById('set-add-pass').value;
      if(!e||!n||!p) return showToast("Ralat", "Lengkapkan semua", "error");
      showToast("Menambah...", "Mendaftar dan menghantar emel pendaftaran...", "info");
      const res = await apiPost({ action: "ADD_USER", emel: e, nama: n, role: r, tempPassword: p });
      if(res.success) { showToast("Selesai", res.message, "success"); loadUsersForAdmin(); } else showToast("Gagal", res.message, "error");
    }

    function deleteUser(emel) {
      showConfirmModal("Padam Pengguna", `Adakah anda pasti untuk memadam akses bagi pengguna: ${emel}? Tindakan ini tidak boleh berpatah balik.`, async () => {
        showToast("Memadam...", "Melaksanakan arahan padaman pada server...", "info");
        const res = await apiPost({ action: "DELETE_USER", emelTarget: emel });
        if(res.success) { showToast("Berjaya", res.message, "success"); loadUsersForAdmin(); } else showToast("Gagal", res.message, "error");
      });
    }

    function openEditUserModal(emel, nama, role) {
      document.getElementById('edit-original-email').value = emel;
      document.getElementById('edit-user-email').value = emel;
      document.getElementById('edit-user-nama').value = nama;
      document.getElementById('edit-user-role').value = role;
      document.getElementById('edit-user-modal').classList.remove('hidden');
    }
    function closeEditUserModal() { document.getElementById('edit-user-modal').classList.add('hidden'); }
    
    async function submitEditUser() {
      const oriEmel = document.getElementById('edit-original-email').value;
      const newEmel = document.getElementById('edit-user-email').value;
      const newNama = document.getElementById('edit-user-nama').value;
      const newRole = document.getElementById('edit-user-role').value;
      
      if(!newEmel || !newNama) return showToast("Ralat", "Sila isikan maklumat yang sah.", "error");
      showToast("Mengemaskini...", "Menghantar arahan ke server...", "info");
      const res = await apiPost({ action: "EDIT_USER", targetEmel: oriEmel, emel: newEmel, nama: newNama, role: newRole });
      if(res.success) { showToast("Selesai", res.message, "success"); closeEditUserModal(); loadUsersForAdmin(); } 
      else showToast("Gagal", res.message, "error");
    }

    function renderApprovalsList() {
      const cont = document.getElementById('approvals-list-container');
      const badgeBox = document.getElementById('approvals-teacher-badge');
      const st = document.getElementById('admin-filter-status').value, tc = document.getElementById('admin-filter-teacher')?.value || 'ALL';
      
      let fil = leaveApplications.filter(a => { if(st !== 'ALL' && a.status !== st) return false; if(currentRole === 'ADMIN' && tc !== 'ALL' && a.teacherName !== tc) return false; return true; });
      fil.sort((a,b) => {
        const parseD = (s) => { try { const p = s.match(/(\d{2})\/(\d{2})\/(\d{4}) (\d{1,2})\.(\d{2})(am|pm)/i); if(!p) return new Date(s).getTime()||0; let h = parseInt(p[4]); if(p[6].toLowerCase()==='pm'&&h<12)h+=12; if(p[6].toLowerCase()==='am'&&h===12)h=0; return new Date(`${p[3]}-${p[2]}-${p[1]}T${String(h).padStart(2,'0')}:${p[5]}:00`).getTime(); } catch(e){return 0;} };
        return parseD(a.appDate) - parseD(b.appDate);
      });

      if(currentRole === 'ADMIN') {
        const navBadge = document.getElementById('nav-admin-badge');
        const mobNavBadge = document.getElementById('mob-nav-admin-badge');
        const hasPending = leaveApplications.filter(x=>x.status==='PENDING').length > 0;
        if(navBadge) { if(hasPending) navBadge.classList.remove('hidden'); else navBadge.classList.add('hidden'); }
        if(mobNavBadge) { if(hasPending) mobNavBadge.classList.remove('hidden'); else mobNavBadge.classList.add('hidden'); }

        if(tc !== 'ALL') {
          let baki = getBakiSemasa(tc), totApp = leaveApplications.filter(x=>x.teacherName===tc).length, totAppr = leaveApplications.filter(x=>x.teacherName===tc && x.status==='APPROVED').length;
          badgeBox.innerHTML = `<div><h4 class="font-bold text-sm text-indigo-900"><i class="fa-solid fa-filter text-indigo-400 mr-2"></i>Sejarah: ${escapeHTML(tc)}</h4><p class="text-[10px] text-indigo-700 mt-1">${totApp} Permohonan | ${totAppr} Lulus</p></div><div class="text-right"><span class="block text-[10px] text-indigo-400 font-bold">BAKI SEMASA</span><span class="text-xl font-black ${baki<=0?'text-rose-500':'text-indigo-600'}">${baki} Hari</span></div>`;
          badgeBox.classList.remove('hidden'); badgeBox.classList.add('flex');
        } else { badgeBox.classList.add('hidden'); badgeBox.classList.remove('flex'); }
      }

      if(fil.length===0) { cont.innerHTML = '<p class="text-xs text-slate-400 text-center p-8 border-2 border-dashed rounded-2xl">Tiada rekod dijumpai.</p>'; return; }
      
      cont.innerHTML = fil.map(a => {
        let bg = a.status==='APPROVED'?'bg-emerald-100 text-emerald-700':a.status==='REJECTED'?'bg-rose-100 text-rose-700':'bg-amber-100 text-amber-700 border', txt = a.status==='APPROVED'?'DILULUSKAN':a.status==='REJECTED'?'DITOLAK': (currentRole==='ADMIN'?'SEMAKAN BARU':'DALAM SEMAKAN');
        let act = "", drive = `<span class="text-slate-500 italic">Tiada Fail</span>`;
        if(currentRole === 'ADMIN' && a.status === 'PENDING') {
          let warn = getBakiSemasa(a.teacherName) <= 0 ? `<div class="bg-rose-50 text-rose-600 border border-rose-200 text-[10px] p-2.5 rounded-lg mb-3 font-bold flex items-center gap-2"><i class="fa-solid fa-triangle-exclamation"></i> Peringatan: Baki cuti tahunan guru ini telah habis (0).</div>` : '';
          act = `${warn}<div class="flex flex-col sm:flex-row gap-2 mt-4 pt-3 border-t"><input type="text" id="n-${a.id}" placeholder="Catatan (Pilihan)" class="flex-1 text-xs px-3 py-2 border rounded-lg outline-none focus:ring-2 focus:ring-indigo-100"><div class="flex gap-2"><button onclick="processApp('${a.id}','REJECTED')" class="flex-1 sm:flex-none px-5 py-2.5 bg-rose-100 text-rose-600 rounded-lg text-xs font-bold shadow-sm hover:bg-rose-200 transition-colors">Tolak</button><button onclick="processApp('${a.id}','APPROVED')" class="flex-1 sm:flex-none px-5 py-2.5 bg-emerald-600 text-white rounded-lg text-xs font-bold shadow-sm hover:bg-emerald-700 transition-colors">Luluskan</button></div></div>`;
        } else if (a.status !== 'PENDING') { act = `<div class="bg-slate-50 p-3 rounded-lg mt-3 text-[10px] border border-slate-100 shadow-inner"><b>Catatan Pengurusan:</b> ${escapeHTML(a.pengetuaNotes)||'Tiada catatan tambahan'}</div>`; }
        try { let o = JSON.parse(a.attachmentName); drive = `<a href="${o.url}" target="_blank" class="text-[10px] bg-indigo-50 text-indigo-700 px-3 py-1.5 rounded font-bold border shadow-sm hover:bg-indigo-100 transition-colors block text-center truncate"><i class="fa-solid fa-file-pdf"></i> Lihat Surat Rasmi</a>`; } catch(e){}
        let rT = a.relief&&a.relief.length>0 ? `<div class="border rounded-lg overflow-x-auto mt-2"><table class="w-full text-[9px] min-w-[200px]"><thead class="bg-slate-100 border-b"><tr><th class="p-1.5 text-left">Hari</th><th class="p-1.5 text-left">Masa</th><th class="p-1.5 text-left">Kelas</th><th class="p-1.5 text-left">Ganti</th></tr></thead><tbody class="divide-y divide-slate-100">${a.relief.map(r=>`<tr class="bg-white"><td class="p-1.5">${escapeHTML(r.day)}</td><td class="p-1.5">${escapeHTML(r.time)}</td><td class="p-1.5 font-bold">${escapeHTML(r.class)}</td><td class="p-1.5 text-indigo-600">${escapeHTML(r.replacement)}</td></tr>`).join('')}</tbody></table></div>` : `<p class="text-[9px] text-slate-400 mt-1">Tiada tugas relief didaftarkan.</p>`;
        return `<div class="border rounded-2xl p-4 sm:p-5 bg-white shadow-sm mb-4 hover:shadow-md transition-shadow"><div class="flex flex-col sm:flex-row justify-between items-start gap-2 mb-3"><div><span class="text-[9px] font-mono bg-slate-100 px-2 py-0.5 border rounded-md shadow-sm">${a.id}</span><p class="text-[10px] text-slate-500 font-bold mt-2"><i class="fa-regular fa-clock mr-1"></i>Dimohon: ${escapeHTML(a.appDate)}</p><h4 class="font-bold text-sm mt-1 text-slate-800">${escapeHTML(a.teacherName)}</h4></div><span class="px-2 py-1 rounded-md text-[9px] font-black shadow-sm self-start ${bg}">${txt}</span></div><div class="grid grid-cols-1 sm:grid-cols-2 gap-3 sm:gap-4 mt-3 bg-slate-50 p-3 rounded-xl border shadow-inner"><div><span class="text-[9px] text-slate-400 font-bold"><i class="fa-regular fa-calendar-check mr-1"></i>TARIKH CUTI</span><p class="text-xs font-bold mt-0.5">${formatTarikhBM(a.startDate)} - ${formatTarikhBM(a.endDate)}</p></div><div><span class="text-[9px] text-slate-400 font-bold"><i class="fa-solid fa-list-check mr-1"></i>SEBAB</span><p class="text-xs font-bold mt-0.5">${escapeHTML(a.reasonCategory)}</p></div></div><div class="grid grid-cols-1 sm:grid-cols-2 gap-4 mt-4"><div><span class="text-[9px] text-slate-400 font-bold uppercase tracking-wider mb-2 block">Lampiran</span>${drive}</div><div class="overflow-hidden"><span class="text-[9px] text-slate-400 font-bold uppercase tracking-wider block">Jadual Relief</span>${rT}</div></div>${act}</div>`;
      }).join('');
    }

    async function processApp(id, stat) {
      showConfirmModal("Sahkan Status", `Hantar status keputusan ini ke emel pemohon dan sistem?`, async () => {
        showToast("Memproses & Emel...", "Sistem sedang mengemaskini rekod dan menghantar emel ke latar belakang...", "info");
        const res = await apiPost({ action: "APPROVE_LEAVE", appId: id, newStatus: stat, notes: document.getElementById(`n-${id}`).value });
        if(res.success) { showToast("Selesai", res.message, "success"); apiPost({ action: "GET_LEAVE_APPLICATIONS" }).then(r=>{if(r.success) {leaveApplications=r.data; renderApprovalsList(); populateTeacherDropdowns();} }); }
        else showToast("Ralat", res.message, "error");
      });
    }

    function initAnalytics() {
      let myApps = (currentRole === 'ADMIN') ? leaveApplications : leaveApplications.filter(a => a.teacherEmail.toLowerCase() === currentEmel.toLowerCase());
      const tot = myApps.length, app = myApps.filter(a=>a.status==='APPROVED').length;
      document.getElementById('stat-total-apps').innerText = tot; document.getElementById('stat-approved-apps').innerText = app; document.getElementById('stat-avg-rate').innerText = tot>0 ? Math.round((app/tot)*100)+'%' : '0%';
      const ctx = document.getElementById('reasonPieChart'); if(reasonChartInstance) reasonChartInstance.destroy();
      const cats = {}; myApps.forEach(a=>{ cats[a.reasonCategory]=(cats[a.reasonCategory]||0)+1; });
      reasonChartInstance = new Chart(ctx, { type: 'doughnut', data: { labels: Object.keys(cats).length?Object.keys(cats):["Tiada Data"], datasets: [{ data: Object.values(cats).length?Object.values(cats):[1], backgroundColor: ['#6366f1','#ec4899','#10b981','#f59e0b'] }] }, options: { maintainAspectRatio: false } });
    }

    function changeQuotaMonth(o) { quotaCalendarDate.setMonth(quotaCalendarDate.getMonth() + o); renderAdminQuotaCalendar(); }
    
    function renderAdminQuotaCalendar() {
      const grid = document.getElementById('quota-calendar-grid'), title = document.getElementById('quota-calendar-title');
      if(!grid) return; const year = quotaCalendarDate.getFullYear(), month = quotaCalendarDate.getMonth();
      title.innerText = `${["Januari","Februari","Mac","April","Mei","Jun","Julai","Ogos","September","Oktober","November","Disember"][month]} ${year}`; grid.innerHTML = '';
      const firstDay = new Date(year, month, 1).getDay(), daysInMonth = new Date(year, month + 1, 0).getDate();
      for (let i = 0; i < firstDay; i++) grid.innerHTML += `<div class="p-2"></div>`;
      for (let day = 1; day <= daysInMonth; day++) {
        const yyyymmdd = `${year}-${String(month+1).padStart(2,'0')}-${String(day).padStart(2,'0')}`;
        const isUnlimited = unlimitedDates.includes(yyyymmdd), dayOfWeek = new Date(year, month, day).getDay();
        let bg = isUnlimited ? "bg-purple-500 text-white shadow-md border-purple-600 scale-105" : (dayOfWeek===0 ? "bg-rose-50 text-rose-400" : (dayOfWeek===6 ? "bg-slate-50 text-indigo-500 hover:bg-purple-50" : "bg-white text-slate-600 hover:bg-purple-50"));
        grid.innerHTML += `<button onclick="toggleUnlimitedDate('${yyyymmdd}')" class="p-3 sm:p-2 rounded-xl border text-xs font-black transition-all ${bg}">${day}</button>`;
      }
    }

    async function toggleUnlimitedDate(dateStr) {
      const isUnlim = unlimitedDates.includes(dateStr); showToast("Menyimpan...", "Mengemaskini status tarikh", "info");
      const res = await apiPost({ action: isUnlim ? "REMOVE_UNLIMITED_DATE" : "ADD_UNLIMITED_DATE", [isUnlim?'date':'dates']: isUnlim?dateStr:[dateStr] });
      if(res.success) { if(isUnlim) unlimitedDates = unlimitedDates.filter(d=>d!==dateStr); else unlimitedDates.push(dateStr); renderAdminQuotaCalendar(); fetchPublicData(); hideToast(); }
    }

    async function addUnlimitedBulk() {
      const s = document.getElementById('admin-unlim-start').value, e = document.getElementById('admin-unlim-end').value || s;
      if(!s) return showToast("Ralat", "Sila pilih sekurang-kurangnya tarikh mula.", "error"); 
      let curr = new Date(s), end = new Date(e), dates = [];
      while(curr <= end) { let d = `${curr.getFullYear()}-${String(curr.getMonth()+1).padStart(2,'0')}-${String(curr.getDate()).padStart(2,'0')}`; if(!unlimitedDates.includes(d)) dates.push(d); curr.setDate(curr.getDate()+1); }
      if(dates.length > 0) {
        showToast("Memproses...", "Mengeset pukal ke sistem...", "info"); const res = await apiPost({ action: "ADD_UNLIMITED_DATE", dates: dates });
        if(res.success) { unlimitedDates.push(...dates); renderAdminQuotaCalendar(); fetchPublicData(); document.getElementById('admin-unlim-start').value=""; document.getElementById('admin-unlim-end').value=""; hideToast(); }
      }
    }

    function switchTab(id) { 
      // Desktop Nav update
      document.querySelectorAll('.tab-content').forEach(e=>e.classList.add('hidden')); 
      document.querySelectorAll('.nav-btn').forEach(b=>b.className=b.className.replace('text-indigo-600 tab-active', 'text-slate-600')); 
      const active = document.getElementById(`nav-${id}`); 
      if(active) active.className=active.className.replace('text-slate-600', 'text-indigo-600 tab-active'); 
      
      // Mobile Nav Update
      document.querySelectorAll('.mob-nav-btn').forEach(b=>{
        b.classList.remove('text-indigo-600');
        b.classList.add('text-slate-400');
      });
      const activeMob = document.getElementById(`mob-nav-${id}`);
      if(activeMob) {
        activeMob.classList.remove('text-slate-400');
        activeMob.classList.add('text-indigo-600');
      }

      document.getElementById(`tab-${id}`).classList.remove('hidden'); 
      if(id==='calendar') renderCalendar(); 
    }
    
    function toggleTheme() { document.documentElement.classList.toggle('dark'); }
  </script>
</body>
</html>
```

## 5. Deployment / Setup Instructions (For Claude Reference)
1. **Google Sheets Setup:** Create a new Google Sheet. Copy its ID.
2. **Apps Script Setup:** 
   * Open the sheet -> Extensions -> Apps Script.
   * Paste the `Code.gs` content into the default file.
   * Create an `index.html` file and paste the HTML content.
   * Put the Google Sheet ID inside the `SPREADSHEET_ID` variable in `Code.gs`.
3. **Execution:** Run `setupAdminPertama()` in the App Script editor once to create the first root Admin account. Provide required Google Authorizations.
4. **Deploy:** Deploy as a Web App (Execute as: "Me", Access: "Anyone").
5. **Final Linking:** Copy the generated Web App URL and paste it into the `GAS_URL` variable inside `index.html`. Deploy again to apply the frontend changes.
`