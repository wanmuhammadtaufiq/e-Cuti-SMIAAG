# e-Cuti SMIAAG — Audit Keselamatan & Bug (02/11/2026)

Status: sistem dah deploy di https://ecutismiaag.netlify.app/ tapi belum digunakan.
Repo GitHub (03/11/2026): https://github.com/WeBake/e-Cuti-SMIAAG (private, branch main). Repo lama wanmuhammadtaufiq/e-Cuti-SMIAAG tidak lagi dikemaskini.
Aliran kerja: perubahan baru → branch + Pull Request → Netlify Deploy Preview → Merge ke main.

## KEMASKINI 02/11/2026 — v2 (fail dalam folder e-cuti/v2/)
Keputusan pemilik:
- Publish-to-Web DATA_GURU sudah dihentikan. Senarai guru kini melalui action GET_DATA_GURU (perlu login).
- Password admin dalam setupAdminPertama sudah ditukar.
- Default "TukarSaya123" untuk user baru DIKEKALKAN atas permintaan (mudah untuk guru). Mitigasi: sesi PENDING_PW, tak boleh guna sistem sebelum tukar.
- BIL_CUTI_TAHUNAN_BERBAKI = kelayakan asal setahun. Baki = kelayakan − cuti APPROVED yang bermula tahun semasa.

- Sabtu = hari cuti. Admin boleh tanda "Sabtu Ada Program" (tab SabtuProgram) dan barulah Sabtu itu dikira hari persekolahan & boleh dimohon cuti (kuota 3 terpakai). Ahad sentiasa cuti.

Dibaiki dalam v2: semua item A (1–5), item C kecuali had minimum 6 aksara dan OTP Math.random, semua item D. Cuti umum belum dikecualikan secara automatik (boleh ditambah kemudian).
Nota: amaran konsol "cdn.tailwindcss.com should not be used in production" hanya amaran prestasi, bukan ralat. Boleh tukar ke Tailwind CLI kemudian jika perlu.
Tambahan v2: bug kalendar lompat bulan pada 31hb (setDate(1)); permohonan hanya boleh diproses sekali; kuota dikira lulus + dalam semakan; lampiran dibuka melalui GET_ATTACHMENT (admin/pemilik).

---
Senarai asal audit (sebelum v2):

## A. Kritikal (baiki sebelum guna)

1. **Publish-to-Web CSV (DATA_GURU)** — URL `docs.google.com/.../pub?output=csv` hardcoded dalam index.html. Nama guru + baki cuti terdedah awam. Kalau setting Publish = "Entire document", tab `Users` (PasswordHash, Salt, ResetKodHash) dan `Sessions` (token aktif!) juga boleh dibaca melalui `&gid=`. Semak: Sheets > File > Share > Publish to web. Fix: Stop publishing, guna action `GET_DATA_GURU` (jadikan public, return nama sahaja).
2. **SUBMIT_LEAVE tanpa token** — sesiapa boleh hantar permohonan atas nama mana-mana guru, upload fail tanpa had ke Drive pemilik skrip. Server tak semak kuota 3 orang/hari, overlap, baki, atau totalDays. Fix: wajibkan login; ambil nama & emel dari session, kira hari di server, enforce kuota di server.
3. **XSS curi token admin** — `attachmentName` = `app.fileName` bila tiada base64. Penyerang hantar fileName `{"name":"x","url":"javascript:..."}` → jadi link "Lihat Surat Rasmi" di portal admin → klik → token admin dalam localStorage dicuri. Fix: server tolak permohonan tanpa fail; frontend hanya terima URL bermula `https://drive.google.com/`.
4. **Surat cuti dikongsi ANYONE_WITH_LINK** — surat (mungkin MC/rekod kesihatan) boleh dibuka sesiapa ada link. Isu PDPA. Fix: jangan setSharing; admin buka melalui akaun Google pemilik.
5. **Tukar role tak batalkan sesi** — role disimpan dalam tab Sessions; admin yang diturunkan ke GURU kekal ADMIN sehingga 12 jam. Fix: `invalidateUserSessions` bila role berubah, atau baca role terkini dari Users dalam validateSession.

## B. Password hardcoded

- `Code.gs` `setupAdminPertama()`: `PASSWORD_SEMENTARA = "TukarSaya123"` + emel `admin@smiaag.edu.my`. Ada force change, tapi semak admin dah tukar password.
- `index.html` input `set-add-pass` default `value="TukarSaya123"` → semua user baru dapat password sama dan boleh diteka; password juga dihantar plaintext dalam emel. Fix: jana password rawak di server.
- Tiada API key/secret lain dalam frontend. GAS_URL memang public — OK. APP_SECURITY_SALT dalam PropertiesService — OK.

## C. Sederhana

- `totalDays` ditulis ke Sheet tanpa sanitizeFormula (formula injection).
- `openEditUserModal('${escapeHTML(...)}')` dalam onclick — escapeHTML tak cukup dalam konteks JS attribute (&#39; ditukar balik jadi '). `u.role` tak di-escape langsung.
- `CHANGE_PASSWORD` tak minta password lama.
- Logout tak padam sesi di server; tab Sessions tak pernah dibersihkan.
- PUBLIC_GHOST dedahkan emel guru untuk semua cuti APPROVED.
- Password minimum 6 aksara (cadang 8+). OTP guna Math.random.
- Tiada notis privasi PDPA pada borang.

## D. Bug fungsi

1. "Rekod Cuti Saya" (GURU) turut papar cuti APPROVED guru lain (dengan label Sulit) — renderApprovalsList tak tapis ikut emel sendiri.
2. ID permohonan = `getLastRow()` + tiada LockService → ID boleh berganda bila hantar serentak atau bila row dipadam; APPROVE_LEAVE akan kena row pertama yang sama ID.
3. Kuota 3/hari hanya kira APPROVED; PENDING tak dikira. Borang boleh pilih tarikh penuh terus.
4. Emel dalam borang ditaip bebas — kalau beza dengan emel login, rekod tak muncul dalam "Rekod Cuti Saya".
5. Baki = BIL_CUTI_TAHUNAN_BERBAKI − semua cuti APPROVED. Kalau kolum itu sudah "baki" (bukan kelayakan asal), jadi tolak dua kali. Padanan ikut nama — tukar ejaan nama, baki lari.
6. Dark mode: input bg-white tapi teks warisi warna cerah → tak boleh baca; bg-aurora-dark tak pernah dipakai; tema tak disimpan.
7. Chart.js dari jsdelivr tanpa versi (`npm/chart.js`) — boleh rosak bila major version baru. Pin versi.
8. Logout tak clear emel/role dari localStorage.
9. Sabtu dikira hari cuti; cuti umum tak dikecualikan (sahkan polisi sekolah).
