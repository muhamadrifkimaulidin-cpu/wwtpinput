# wwtpinput
Backup aplikasi penginputan WWTP (Google Apps Script web app, data di Google Sheets).

## File
| File | Isi |
|---|---|
| `Code.gs` | `doGet`, `include` |
| `Auth.gs` | Login, sesi token server-side, hash password, role & approver |
| `Helpers.gs` | Tanggal/shift server, ID unik, lock, cache berversi, config, util sheet |
| `Modules.gs` | Registry kolom semua modul (Operasional, FIK, Softener, Maintenance, Pressure Gauge) |
| `Crud.gs` | Simpan & baca entri semua modul, data shift sebelumnya |
| `Requests.gs` | Request ubah/hapus + approval (request baru menggantikan request pending di record yang sama) |
| `Notifications.gs` | Notifikasi anomali + email |
| `Report.gs` | Laporan sheet / Excel / CSV / PDF |
| `Setup.gs` | Setup sheet, migrasi, auto-backup |
| `Transfer.gs` | Menu "WWTP Tools" di Spreadsheet |
| `Index.html`, `Stylesheet.html`, `JavaScript.html` | UI |

## Deploy dari v5 ke v6
1. Di editor Apps Script **hapus** file lama: `Dashboard.gs`, `Crudproses.gs`, `Crudmaintenance.gs`.
2. Tambah file baru: `Modules.gs`, `Requests.gs`, `Notifications.gs`. Timpa file lain dengan versi repo ini.
3. Buka Spreadsheet → refresh → menu **WWTP Tools › Admin › Siapkan Sheet & Kolom Baru** (tidak menghapus data; membuat `DataPressureGauge`, menambah kolom `Requests.module`, `Notifications.user_id`, key Config `approver_ids`, `pg_unit`).
4. Cek sheet `Config`: `approver_ids` (default `USR005,USR006,USR003`), `pg_unit` (default `bar`), `admin_email`.
5. Ganti password semua user di sheet `Users`, lalu **WWTP Tools › Admin › Amankan Password Users (hash)**.
6. Deploy › Manage deployments › Edit › New version.

Setelah mengubah sheet `Users` atau `Config`: **WWTP Tools › Admin › Muat Ulang Cache Users & Config**.
