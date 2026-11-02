// ============================================================
// GOOGLE APPS SCRIPT — PINJAM IMIP BACKEND (INVENTARIS & PEMINJAMAN)
// ============================================================
// CARA SETUP:
// 1. Buat Spreadsheet baru di Google Drive (misal: "DB Peminjaman Barang IMIP")
// 2. Klik Extensions > Apps Script
// 3. Hapus semua kode default, paste SELURUH kode ini ke editor
// 4. Klik menu dropdown fungsi, pilih "setupSheetsDanDrive" lalu klik Run (Jalankan)
//    -> Setujui izin akses Google Drive dan Google Sheets
// 5. Klik Deploy > New Deployment > Web App
//    - Description: "PinjamIMIP API v1"
//    - Execute as: "Me" (Email Anda)
//    - Who has access: "Anyone" (Siapa saja)  <-- PENTING!
// 6. Copy URL Web App yang didapat, lalu masukkan ke web PinjamIMIP di bagian Admin!
// ============================================================

const SHEET_BARANG     = "barang";
const SHEET_PEMINJAMAN = "peminjaman";
const FOLDER_DRIVE     = "PinjamIMIP - Foto & Bukti";

// PIN Admin Default (Bisa Anda ubah di sini)
const ADMIN_PIN_DEFAULT = "123456";

// Header Default jika sheet baru dibuat
const HEADERS_BARANG = [
  "id_barang", "nama_barang", "kategori", "deskripsi", 
  "foto_url", "stok_total", "stok_tersedia", "kondisi", "status", 
  "created_at", "updated_at"
];

const HEADERS_PEMINJAMAN = [
  "id_pinjam", "id_barang", "nama_barang", "nama_peminjam", 
  "kontak_wa", "divisi_instansi", "tgl_pinjam", "tgl_kembali", 
  "keperluan", "foto_identitas", "status", "catatan_admin", 
  "created_at", "updated_at"
];

// ============================================================
// FUNGSI INISIALISASI OTOMATIS (JALANKAN SEKALI SAAT PERTAMA KALI)
// ============================================================
function setupSheetsDanDrive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Buat Sheet Barang jika belum ada
  let sheetB = ss.getSheetByName(SHEET_BARANG);
  if (!sheetB) {
    sheetB = ss.insertSheet(SHEET_BARANG);
    sheetB.appendRow(HEADERS_BARANG);
    sheetB.getRange(1, 1, 1, HEADERS_BARANG.length).setFontWeight("bold").setBackground("#080C14").setFontColor("#FFFFFF");
    
    // Data contoh awal
    sheetB.appendRow([
      "BRG-001", "Proyektor Epson EB-E500", "Elektronik", 
      "Proyektor HDMI/VGA 3300 Lumens untuk presentasi & kajian",
      "https://images.unsplash.com/photo-1517604931442-7e0c8ed2963c?w=600",
      2, 2, "Bagus", "Tersedia", new Date(), new Date()
    ]);
    sheetB.appendRow([
      "BRG-002", "Wireless Mic Clip-On (2 Pcs)", "Audio & Sound", 
      "Mic wireless receiver Type-C/Lightning untuk dokumentasi/kajian",
      "https://images.unsplash.com/photo-1590658268037-6bf12165a8df?w=600",
      3, 3, "Bagus", "Tersedia", new Date(), new Date()
    ]);
    sheetB.appendRow([
      "BRG-003", "Kabel Rol Listrik 15 Meter", "Kelistrikan", 
      "Kabel sambung tebal 4 stop kontak",
      "https://images.unsplash.com/photo-1558494949-ef010cbdcc31?w=600",
      4, 4, "Bagus", "Tersedia", new Date(), new Date()
    ]);
  }

  // 2. Buat Sheet Peminjaman jika belum ada
  let sheetP = ss.getSheetByName(SHEET_PEMINJAMAN);
  if (!sheetP) {
    sheetP = ss.insertSheet(SHEET_PEMINJAMAN);
    sheetP.appendRow(HEADERS_PEMINJAMAN);
    sheetP.getRange(1, 1, 1, HEADERS_PEMINJAMAN.length).setFontWeight("bold").setBackground("#080C14").setFontColor("#FFFFFF");
  }

  // 3. Buat Folder Google Drive jika belum ada
  getOrCreateFolder(FOLDER_DRIVE);
  
  Logger.log("SETUP BERHASIL! Sheet barang, peminjaman, dan folder Drive siap digunakan.");
}

// ============================================================
// HANDLER REQUEST GET (READ DATA)
// ============================================================
function doGet(e) {
  try {
    const p = e.parameter || {};
    const action = p.action || "get_all";

    if (action === "get_all" || action === "get_katalog") {
      const dataBarang = getTableData(SHEET_BARANG);
      return jsonResponse({
        success: true,
        barang: dataBarang
      });
    }

    if (action === "get_admin_data") {
      const pin = p.pin || "";
      if (pin !== ADMIN_PIN_DEFAULT) {
        return jsonResponse({ success: false, message: "PIN Admin salah!" });
      }
      const dataBarang = getTableData(SHEET_BARANG);
      const dataPeminjaman = getTableData(SHEET_PEMINJAMAN);
      return jsonResponse({
        success: true,
        barang: dataBarang,
        peminjaman: dataPeminjaman
      });
    }

    if (action === "check_status") {
      const query = (p.query || "").toString().trim().toLowerCase();
      if (!query) {
        return jsonResponse({ success: false, message: "Masukkan nomor WA atau ID Peminjaman!" });
      }
      const allPeminjaman = getTableData(SHEET_PEMINJAMAN);
      const filtered = allPeminjaman.filter(row => {
        const wa = (row.kontak_wa || "").toString().replace(/[^0-9]/g, "");
        const cleanQuery = query.replace(/[^0-9]/g, "");
        const id = (row.id_pinjam || "").toString().toLowerCase();
        return (cleanQuery && wa.includes(cleanQuery)) || id.includes(query);
      });
      return jsonResponse({ success: true, hasil: filtered });
    }

    return jsonResponse({ success: true, message: "PinjamIMIP API is running." });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ============================================================
// HANDLER REQUEST POST (SUBMIT, APPROVE, REJECT, RETURN, CRUD BARANG)
// ============================================================
function doPost(e) {
  try {
    let payload = {};
    if (e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (errJson) {
        payload = e.parameter || {};
      }
    } else {
      payload = e.parameter || {};
    }

    const action = payload.action;

    // 1. PUBLIK: AJUKAN PINJAM BARANG
    if (action === "ajukan_pinjam") {
      return handleAjukanPinjam(payload);
    }

    // 2. ADMIN: AUTH CHECK
    const pin = payload.pin || "";
    if (pin !== ADMIN_PIN_DEFAULT) {
      return jsonResponse({ success: false, message: "Akses ditolak: PIN Admin salah!" });
    }

    // 3. ADMIN: VERIFIKASI / TINJAU PINJAM (TERIMA, TOLAK, KEMBALI)
    if (action === "admin_tinjau") {
      return handleAdminTinjau(payload);
    }

    // 4. ADMIN: TAMBAH BARANG
    if (action === "admin_tambah_barang") {
      return handleTambahBarang(payload);
    }

    // 5. ADMIN: EDIT BARANG
    if (action === "admin_edit_barang") {
      return handleEditBarang(payload);
    }

    // 6. ADMIN: HAPUS BARANG
    if (action === "admin_hapus_barang") {
      return handleHapusBarang(payload);
    }

    return jsonResponse({ success: false, message: "Action tidak dikenali." });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ============================================================
// LOGIKA BISNIS
// ============================================================

function handleAjukanPinjam(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetP = ss.getSheetByName(SHEET_PEMINJAMAN);
  const sheetB = ss.getSheetByName(SHEET_BARANG);
  if (!sheetP || !sheetB) throw new Error("Sheet tidak ditemukan. Jalankan setup terlebih dahulu.");

  // Validasi ketersediaan barang
  const barangRows = getTableData(SHEET_BARANG);
  const targetBarang = barangRows.find(b => String(b.id_barang) === String(data.id_barang));
  if (!targetBarang) {
    return jsonResponse({ success: false, message: "Barang tidak ditemukan!" });
  }
  if (parseInt(targetBarang.stok_tersedia, 10) <= 0) {
    return jsonResponse({ success: false, message: "Mohon maaf, stok barang ini sedang habis / dipinjam." });
  }

  // Generate ID Pinjam: PJM-YYYYMMDD-XXXX
  const todayStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyyMMdd");
  const randomSuffix = Math.floor(1000 + Math.random() * 9000);
  const idPinjam = "PJM-" + todayStr + "-" + randomSuffix;

  // Upload foto identitas / KTM jika ada
  let fotoUrl = "";
  if (data.foto_identitas && data.foto_identitas.startsWith("data:")) {
    fotoUrl = saveImageToDrive(data.foto_identitas, "ID_" + idPinjam);
  } else {
    fotoUrl = data.foto_identitas || "";
  }

  const now = new Date();
  const newRow = [
    idPinjam,
    data.id_barang,
    targetBarang.nama_barang,
    data.nama_peminjam || "",
    data.kontak_wa || "",
    data.divisi_instansi || "",
    data.tgl_pinjam || "",
    data.tgl_kembali || "",
    data.keperluan || "",
    fotoUrl,
    "MENUNGGU", // Status awal
    "",         // Catatan admin
    now,
    now
  ];

  sheetP.appendRow(newRow);

  return jsonResponse({
    success: true,
    message: "Pengajuan berhasil dikirim! Silakan tunggu konfirmasi admin.",
    id_pinjam: idPinjam,
    nama_barang: targetBarang.nama_barang
  });
}

function handleAdminTinjau(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetP = ss.getSheetByName(SHEET_PEMINJAMAN);
  const sheetB = ss.getSheetByName(SHEET_BARANG);
  
  const idPinjam = data.id_pinjam;
  const statusBaru = (data.status || "").toUpperCase(); // DISETUJUI, DITOLAK, DIKEMBALIKAN
  const catatan = data.catatan || "";

  const pData = sheetP.getDataRange().getValues();
  const pHeaders = pData[0].map(h => String(h).trim().toLowerCase());
  const idxIdPinjam = pHeaders.indexOf("id_pinjam");
  const idxIdBarang = pHeaders.indexOf("id_barang");
  const idxStatus = pHeaders.indexOf("status");
  const idxCatatan = pHeaders.indexOf("catatan_admin");
  const idxUpdated = pHeaders.indexOf("updated_at");

  let targetRowIdx = -1;
  let idBarang = "";
  let oldStatus = "";

  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][idxIdPinjam]).trim() === String(idPinjam).trim()) {
      targetRowIdx = i + 1; // 1-based index
      idBarang = String(pData[i][idxIdBarang]).trim();
      oldStatus = String(pData[i][idxStatus]).trim().toUpperCase();
      break;
    }
  }

  if (targetRowIdx === -1) {
    return jsonResponse({ success: false, message: "Data peminjaman tidak ditemukan!" });
  }

  // Update status peminjaman
  sheetP.getRange(targetRowIdx, idxStatus + 1).setValue(statusBaru);
  if (catatan) {
    sheetP.getRange(targetRowIdx, idxCatatan + 1).setValue(catatan);
  }
  sheetP.getRange(targetRowIdx, idxUpdated + 1).setValue(new Date());

  // Update stok barang jika status berubah
  if (statusBaru === "DISETUJUI" && oldStatus !== "DISETUJUI") {
    ubahStokBarang(sheetB, idBarang, -1);
  } else if (statusBaru === "DIKEMBALIKAN" && oldStatus === "DISETUJUI") {
    ubahStokBarang(sheetB, idBarang, 1);
  }

  return jsonResponse({
    success: true,
    message: "Status peminjaman berhasil diubah menjadi: " + statusBaru
  });
}

function ubahStokBarang(sheetB, idBarang, delta) {
  const bData = sheetB.getDataRange().getValues();
  const bHeaders = bData[0].map(h => String(h).trim().toLowerCase());
  const idxId = bHeaders.indexOf("id_barang");
  const idxTersedia = bHeaders.indexOf("stok_tersedia");
  const idxTotal = bHeaders.indexOf("stok_total");
  const idxStatus = bHeaders.indexOf("status");

  for (let i = 1; i < bData.length; i++) {
    if (String(bData[i][idxId]).trim() === String(idBarang).trim()) {
      const curTersedia = parseInt(bData[i][idxTersedia], 10) || 0;
      const total = parseInt(bData[i][idxTotal], 10) || 1;
      let newTersedia = curTersedia + delta;
      if (newTersedia < 0) newTersedia = 0;
      if (newTersedia > total) newTersedia = total;

      sheetB.getRange(i + 1, idxTersedia + 1).setValue(newTersedia);
      sheetB.getRange(i + 1, idxStatus + 1).setValue(newTersedia > 0 ? "Tersedia" : "Dipinjam");
      break;
    }
  }
}

function handleTambahBarang(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetB = ss.getSheetByName(SHEET_BARANG);

  const nextId = "BRG-" + String(Math.floor(100 + Math.random() * 900));
  
  let fotoUrl = "";
  if (data.foto_url && data.foto_url.startsWith("data:")) {
    fotoUrl = saveImageToDrive(data.foto_url, nextId);
  } else {
    fotoUrl = data.foto_url || "";
  }

  const stok = parseInt(data.stok_total, 10) || 1;
  const now = new Date();

  sheetB.appendRow([
    nextId,
    data.nama_barang || "Barang Baru",
    data.kategori || "Umum",
    data.deskripsi || "",
    fotoUrl,
    stok,
    stok, // Stok tersedia awal sama dengan stok total
    data.kondisi || "Bagus",
    "Tersedia",
    now,
    now
  ]);

  return jsonResponse({
    success: true,
    message: "Barang baru berhasil ditambahkan!",
    id_barang: nextId
  });
}

function handleEditBarang(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetB = ss.getSheetByName(SHEET_BARANG);
  const idBarang = data.id_barang;

  const bData = sheetB.getDataRange().getValues();
  const bHeaders = bData[0].map(h => String(h).trim().toLowerCase());
  const idxId = bHeaders.indexOf("id_barang");
  const idxNama = bHeaders.indexOf("nama_barang");
  const idxKategori = bHeaders.indexOf("kategori");
  const idxDeskripsi = bHeaders.indexOf("deskripsi");
  const idxFoto = bHeaders.indexOf("foto_url");
  const idxTotal = bHeaders.indexOf("stok_total");
  const idxTersedia = bHeaders.indexOf("stok_tersedia");
  const idxKondisi = bHeaders.indexOf("kondisi");
  const idxStatus = bHeaders.indexOf("status");
  const idxUpdated = bHeaders.indexOf("updated_at");

  let rowIdx = -1;
  for (let i = 1; i < bData.length; i++) {
    if (String(bData[i][idxId]).trim() === String(idBarang).trim()) {
      rowIdx = i + 1;
      break;
    }
  }

  if (rowIdx === -1) {
    return jsonResponse({ success: false, message: "Barang tidak ditemukan!" });
  }

  if (data.nama_barang !== undefined) sheetB.getRange(rowIdx, idxNama + 1).setValue(data.nama_barang);
  if (data.kategori !== undefined) sheetB.getRange(rowIdx, idxKategori + 1).setValue(data.kategori);
  if (data.deskripsi !== undefined) sheetB.getRange(rowIdx, idxDeskripsi + 1).setValue(data.deskripsi);
  if (data.kondisi !== undefined) sheetB.getRange(rowIdx, idxKondisi + 1).setValue(data.kondisi);

  if (data.foto_url && data.foto_url.startsWith("data:")) {
    const newFoto = saveImageToDrive(data.foto_url, idBarang + "_update");
    sheetB.getRange(rowIdx, idxFoto + 1).setValue(newFoto);
  }

  if (data.stok_total !== undefined) {
    const newTotal = parseInt(data.stok_total, 10) || 1;
    sheetB.getRange(rowIdx, idxTotal + 1).setValue(newTotal);
    // Sesuaikan stok tersedia jika stok baru lebih kecil
    const curTersedia = parseInt(bData[rowIdx - 1][idxTersedia], 10) || 0;
    if (curTersedia > newTotal) {
      sheetB.getRange(rowIdx, idxTersedia + 1).setValue(newTotal);
    }
  }

  sheetB.getRange(rowIdx, idxUpdated + 1).setValue(new Date());

  return jsonResponse({ success: true, message: "Data barang berhasil diperbarui!" });
}

function handleHapusBarang(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetB = ss.getSheetByName(SHEET_BARANG);
  const idBarang = data.id_barang;

  const bData = sheetB.getDataRange().getValues();
  const bHeaders = bData[0].map(h => String(h).trim().toLowerCase());
  const idxId = bHeaders.indexOf("id_barang");

  for (let i = 1; i < bData.length; i++) {
    if (String(bData[i][idxId]).trim() === String(idBarang).trim()) {
      sheetB.deleteRow(i + 1);
      return jsonResponse({ success: true, message: "Barang berhasil dihapus." });
    }
  }

  return jsonResponse({ success: false, message: "Barang tidak ditemukan." });
}

// ============================================================
// HELPER DRIVE & SPREADSHEET
// ============================================================

function getTableData(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) return [];

  const headers = values[0].map(h => String(h).trim().toLowerCase());
  const result = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const obj = {};
    for (let c = 0; c < headers.length; c++) {
      let val = row[c];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
      }
      obj[headers[c]] = val;
    }
    result.push(obj);
  }
  return result;
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  while (folders.hasNext()) {
    const f = folders.next();
    if (!f.isTrashed()) return f;
  }
  const created = DriveApp.createFolder(folderName);
  try {
    created.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {}
  return created;
}

function saveImageToDrive(base64Data, filenamePrefix) {
  if (!base64Data || !base64Data.startsWith("data:")) return base64Data;
  try {
    const commaIndex = base64Data.indexOf(",");
    if (commaIndex === -1) return base64Data;

    const meta = base64Data.substring(0, commaIndex);
    const rawBase64 = base64Data.substring(commaIndex + 1);
    
    let mimeType = "image/jpeg";
    const mimeMatch = meta.match(/:(.*?);/);
    if (mimeMatch && mimeMatch[1]) mimeType = mimeMatch[1];

    let ext = ".jpg";
    if (mimeType.includes("png")) ext = ".png";
    else if (mimeType.includes("webp")) ext = ".webp";

    const bytes = Utilities.base64Decode(rawBase64);
    const blob = Utilities.newBlob(bytes, mimeType, filenamePrefix + "_" + Date.now() + ext);

    const folder = getOrCreateFolder(FOLDER_DRIVE);
    const file = folder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(e) {}

    const fileId = file.getId();
    // Gunakan direct thumbnail URL agar mudah diakses tanpa kendala izin browser
    return "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w1000";
  } catch (err) {
    Logger.log("Error saveImage: " + err.toString());
    return "";
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
