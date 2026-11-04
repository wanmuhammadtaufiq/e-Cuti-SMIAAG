# e-Cuti SMIAAG

Sistem permohonan cuti tahunan guru SMI Al-Amin Gombak.

- Frontend: `public/index.html` (Tailwind CDN, vanilla JS), dihoskan di Netlify — https://ecutismiaag.netlify.app/
- Backend: `apps-script/Code.gs` (Google Apps Script + Google Sheets + Google Drive)

## Struktur

```
public/index.html      → halaman web (yang Netlify publish)
apps-script/Code.gs    → salinan kod Apps Script (TIDAK auto-deploy)
netlify.toml           → tetapan Netlify (folder publish + header keselamatan)
```

## Cara kemaskini (dengan preview dulu)

1. Buat perubahan di branch baharu, bukan terus di `main`.
2. Buka Pull Request ke `main`. Netlify akan jana **Deploy Preview** (link dalam PR, bahagian checks / komen netlify bot).
3. Test di link preview. Kalau OK, klik **Merge** — Netlify deploy ke production secara automatik.

> Deploy Preview menggunakan backend Apps Script yang SAMA dengan production (GAS_URL sama). Jadi permohonan yang dihantar semasa test akan masuk ke Sheet sebenar. Padam baris ujian selepas test, atau test guna akaun ujian.

## Bila ubah Code.gs

Netlify tidak deploy Apps Script. Selepas merge:

1. Buka Apps Script projek e-Cuti, tampal kandungan `apps-script/Code.gs`.
2. **Deploy → Manage deployments → Edit (ikon pensel) → Version: New version → Deploy.**
   Jangan buat "New deployment" — URL akan berubah dan `GAS_URL` dalam index.html perlu ditukar.

Kalau Code.gs dan index.html sama-sama berubah, deploy Apps Script DULU, baru merge PR.

## Tab Google Sheets

`Users`, `Sessions`, `PermohonanCuti`, `Log`, `DATA_GURU`, `TarikhTanpaHad`, `SabtuProgram` — semua dicipta automatik oleh `pastikanTab()` kecuali `DATA_GURU` (diisi manual: `NAMA_GURU`, `BIL_CUTI_TAHUNAN_BERBAKI` = kelayakan setahun).

## Keselamatan

- Jangan "Publish to web" mana-mana tab Sheet.
- Repo ini public: jangan commit Sheet ID, password, atau data guru.
- Salt password disimpan dalam Script Properties (`APP_SECURITY_SALT`), bukan dalam kod.
