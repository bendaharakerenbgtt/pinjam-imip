// ============================================================
// GOOGLE APPS SCRIPT — PINJAMIP BACKEND (V2 - PRD ENHANCED)
// Sistem Katalog & Peminjaman Inventaris LDK IMIP
// ============================================================
// Konsep: Multi-Item Loan, Stock Locking, Approval Workflow,
// Dokumen Perjanjian PDF, & Google Drive Integration.
// ============================================================

const SHEET_BARANG      = "barang";
const SHEET_PEMINJAMAN  = "peminjaman";
const SHEET_PENGATURAN  = "pengaturan";
const FOLDER_DRIVE_NAME = "PinjamIMIP - Foto & Bukti";

// Default PIN & Nomor Kontak Admin
const ADMIN_PIN_DEFAULT = "admin";
const ADMIN_WA_DEFAULT  = "085770976399";

function getSavedAdminPin() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetS = ss.getSheetByName(SHEET_PENGATURAN);
    if (!sheetS) return ADMIN_PIN_DEFAULT;
    const rows = sheetS.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      const k = String(rows[i][0]).trim();
      if (k === "admin_pin" || k === "admin_password") {
        const val = String(rows[i][1] || "").trim();
        if (val) return val;
      }
    }
  } catch(e) {}
  return ADMIN_PIN_DEFAULT;
}

const HEADERS_BARANG = [
  "id_barang", "nama_barang", "kategori", "deskripsi", 
  "foto_url", "stok_total", "stok_tersedia", "kondisi", "status", 
  "created_at", "updated_at"
];

const HEADERS_PEMINJAMAN = [
  "id_pinjam", "no_surat", "nama_peminjam", "kontak_wa", 
  "asal_organisasi", "kategori_peminjam", "nama_penanggung_jawab",
  "items_json", "tgl_pengambilan", "tgl_pengembalian", "tujuan", "tempat_kegiatan", 
  "catatan_tambahan", "dokumen_identitas", "status", 
  "catatan_admin", "tgl_penyerahan", "tgl_kembali_aktual", 
  "kondisi_kembali", "catatan_kembali", "created_at", "updated_at"
];

// ============================================================
// FUNGSI SETUP AWAL
// ============================================================
function setupSheetsDanDrive() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Sheet Barang
  let sheetB = ss.getSheetByName(SHEET_BARANG);
  if (!sheetB) {
    sheetB = ss.insertSheet(SHEET_BARANG);
    sheetB.appendRow(HEADERS_BARANG);
    sheetB.getRange(1, 1, 1, HEADERS_BARANG.length).setFontWeight("bold").setBackground("#080C14").setFontColor("#FFFFFF");
    
    sheetB.appendRow([
      "BRG-001", "Proyektor Epson EB-E500", "Elektronik", 
      "Proyektor HDMI/VGA 3300 Lumens untuk presentasi & kajian",
      "https://images.unsplash.com/photo-1517604931442-7e0c8ed2963c?w=600",
      2, 2, "Bagus / Prima", "Tersedia", new Date(), new Date()
    ]);
    sheetB.appendRow([
      "BRG-002", "Wireless Mic Clip-On (2 Pcs)", "Audio & Sound", 
      "Mic wireless receiver Type-C/Lightning untuk dokumentasi/kajian",
      "https://images.unsplash.com/photo-1590658268037-6bf12165a8df?w=600",
      3, 3, "Bagus / Prima", "Tersedia", new Date(), new Date()
    ]);
    sheetB.appendRow([
      "BRG-003", "Kabel Rol Listrik 15 Meter", "Kelistrikan", 
      "Kabel sambung tebal 4 stop kontak",
      "https://images.unsplash.com/photo-1558494949-ef010cbdcc31?w=600",
      4, 4, "Bagus / Prima", "Tersedia", new Date(), new Date()
    ]);
    sheetB.appendRow([
      "BRG-004", "Standing Banner Aluminium Portable", "Dekorasi & Acara", 
      "Roll-up standing banner kokoh untuk penunjuk arah dan backdrop mini",
      "https://images.unsplash.com/photo-1542744095-fcf48d80b0fd?w=600",
      2, 2, "Bagus / Prima", "Tersedia", new Date(), new Date()
    ]);
  }

  // 2. Sheet Peminjaman
  let sheetP = ss.getSheetByName(SHEET_PEMINJAMAN);
  if (!sheetP) {
    sheetP = ss.insertSheet(SHEET_PEMINJAMAN);
    sheetP.appendRow(HEADERS_PEMINJAMAN);
    sheetP.getRange(1, 1, 1, HEADERS_PEMINJAMAN.length).setFontWeight("bold").setBackground("#080C14").setFontColor("#FFFFFF");
  }

  // 3. Sheet Pengaturan Organisasi
  let sheetS = ss.getSheetByName(SHEET_PENGATURAN);
  if (!sheetS) {
    sheetS = ss.insertSheet(SHEET_PENGATURAN);
    sheetS.appendRow(["key", "value"]);
    sheetS.appendRow(["nama_organisasi", "LDK IMIP"]);
    sheetS.appendRow(["no_wa_admin", ADMIN_WA_DEFAULT]);
    sheetS.appendRow(["alamat", "Politeknik Negeri Media Kreatif Jakarta"]);
    sheetS.appendRow(["admin_pin", ADMIN_PIN_DEFAULT]);
  }

  // 4. Folder Drive
  getOrCreateFolder(FOLDER_DRIVE_NAME);
  Logger.log("SETUP BERHASIL! Database Pinjamip V2 siap digunakan.");
}

// ============================================================
// HANDLER GET (READ DATA)
// ============================================================
function doGet(e) {
  try {
    const p = e.parameter || {};
    const action = p.action || "get_katalog";

    if (action === "get_katalog" || action === "get_all") {
      const dataBarang = getTableData(SHEET_BARANG);
      const allLoans = getTableData(SHEET_PEMINJAMAN);
      const activeSt = ["MENUNGGU VERIFIKASI", "MENUNGGU KONFIRMASI PEMINJAM", "DIKONFIRMASI", "SEDANG DIPINJAM"];
      const jadwalAktif = allLoans.filter(p => activeSt.includes(String(p.status || "").toUpperCase())).map(p => ({
        id_pinjam: p.id_pinjam,
        status: p.status,
        tgl_pengambilan: p.tgl_pengambilan,
        tgl_pengembalian: p.tgl_pengembalian,
        items_json: p.items_json
      }));
      return jsonResponse({
        success: true,
        barang: dataBarang,
        jadwal_aktif: jadwalAktif,
        public_settings: getPublicSettings()
      });
    }

    if (action === "get_admin_data") {
      const pin = p.pin || "";
      if (!isValidAdminPin(pin)) {
        return jsonResponse({ success: false, message: "Akses ditolak: PIN / Password salah!" });
      }
      const dataBarang = getTableData(SHEET_BARANG);
      const dataPeminjaman = getTableData(SHEET_PEMINJAMAN);
      const settings = getSettings();
      return jsonResponse({
        success: true,
        barang: dataBarang,
        peminjaman: dataPeminjaman,
        settings: settings
      });
    }

    if (action === "check_status") {
      const query = (p.query || "").toString().trim().toLowerCase();
      if (!query || query.length < 3) {
        return jsonResponse({ success: false, message: "Masukkan minimal 3 karakter nama, nomor WhatsApp lengkap, atau ID Peminjaman!" });
      }
      const allPeminjaman = getTableData(SHEET_PEMINJAMAN);
      const cleanNum = query.replace(/[^0-9]/g, "");
      const strippedQ = cleanNum.replace(/^(?:62|0)/, "");

      // Filter aman & ketat: Cegah enumerasi massal (seperti mencari '2026', 'pjm-', atau awalan provider '0857')
      const matches = allPeminjaman.filter(row => {
        const rowWa = (row.kontak_wa || "").toString().replace(/[^0-9]/g, "");
        const strippedWa = rowWa.replace(/^(?:62|0)/, "");
        const rowNama = (row.nama_peminjam || "").toString().toLowerCase().trim();
        const rowId = (row.id_pinjam || "").toString().toLowerCase().trim();

        // 1. Cocokkan Nomor WA: wajib minimal 8 digit angka dan cocok di bagian akhir/utuh
        const matchWa = strippedQ.length >= 8 && strippedWa.length >= 8 && (
          strippedWa === strippedQ ||
          strippedWa.endsWith(strippedQ) ||
          strippedQ.endsWith(strippedWa)
        );

        // 2. Cocokkan ID Peminjaman: harus cocok utuh (pjm-yyyymmdd-xxxx) ATAU 4 digit kode unik akhir (-xxxx)
        const matchId = (rowId === query) || (query.length === 4 && rowId.endsWith("-" + query));

        // 3. Cocokkan Nama Peminjam: minimal 3 karakter, bukan sekadar angka tahun/prefix ID
        const isGenericPrefix = /^(pjm|pjm-|202\d|\d+)$/i.test(query);
        const matchNama = !isGenericPrefix && query.length >= 3 && rowNama.includes(query);

        return matchWa || matchId || matchNama;
      }).map(row => {
        // Sensor dokumen identitas dan catatan admin internal
        let parsedItems = [];
        try {
          parsedItems = JSON.parse(row.items_json || "[]");
        } catch(e) {}

        // Masking WA (misal: 0857****6399) - cegah string #ERROR! dari formula spreadsheet
        let waStr = String(row.kontak_wa || "").trim().replace(/^'/, "");
        if (waStr === "#ERROR!" || waStr.includes("ERROR") || waStr.length < 5) {
          waStr = cleanNum.length >= 8 ? (cleanNum.startsWith("62") ? ("0" + cleanNum.slice(2)) : cleanNum) : "Terdaftar";
        }
        let maskedWa = waStr;
        if (waStr.length > 7 && !waStr.includes("Terdaftar")) {
          maskedWa = waStr.substring(0, 4) + "****" + waStr.substring(waStr.length - 4);
        }

        return {
          id_pinjam: row.id_pinjam,
          no_surat: row.no_surat || "-",
          nama_peminjam: row.nama_peminjam,
          kontak_wa_masked: maskedWa,
          asal_organisasi: row.asal_organisasi,
          tgl_pengambilan: row.tgl_pengambilan,
          tgl_pengembalian: row.tgl_pengembalian,
          tujuan: row.tujuan,
          tempat_kegiatan: row.tempat_kegiatan || "-",
          items: parsedItems,
          status: row.status,
          catatan_publik: row.catatan_admin || "",
          created_at: row.created_at
        };
      });

      return jsonResponse({ success: true, hasil: matches, public_settings: getPublicSettings() });
    }

    return jsonResponse({ success: true, message: "Pinjamip API V2 is active." });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ============================================================
// HANDLER POST (SUBMIT, APPROVE, REJECT, ACTION, CRUD)
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

    // 1. PUBLIK: AJUKAN PEMINJAMAN (MULTI-ITEM)
    if (action === "ajukan_pinjam") {
      return handleAjukanPinjam(payload);
    }

    // 2. ADMIN AUTH CHECK
    const pin = payload.pin || "";
    if (!isValidAdminPin(pin)) {
      return jsonResponse({ success: false, message: "Akses ditolak: PIN Admin salah!" });
    }

    // 3. ADMIN: APPROVE / REJECT / KONFIRMASI / SERAHKAN / KEMBALI
    if (action === "admin_update_status") {
      return handleAdminUpdateStatus(payload);
    }

    // 3b. ADMIN: UPDATE KONTAK WA PEMINJAM (JIKA DATA LAMA RUSAK/ERROR)
    if (action === "admin_update_kontak") {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetP = ss.getSheetByName(SHEET_PEMINJAMAN);
      const idPinjam = payload.id_pinjam;
      let newWa = String(payload.kontak_wa || "").replace(/[^0-9]/g, "");
      if (newWa.startsWith("62")) newWa = "0" + newWa.slice(2);
      else if (newWa && !newWa.startsWith("0")) newWa = "0" + newWa;
      const safeWa = newWa ? ("'" + newWa) : "";

      const pData = sheetP.getDataRange().getValues();
      const pHeaders = pData[0].map(h => String(h).trim().toLowerCase());
      const idxId = pHeaders.indexOf("id_pinjam");
      const idxWa = pHeaders.indexOf("kontak_wa");

      for (let i = 1; i < pData.length; i++) {
        if (String(pData[i][idxId]).trim() === String(idPinjam).trim()) {
          sheetP.getRange(i + 1, idxWa + 1).setValue(safeWa);
          return jsonResponse({ success: true, message: "Nomor WhatsApp berhasil diperbarui di spreadsheet!", kontak_wa: safeWa });
        }
      }
      return jsonResponse({ success: false, message: "Peminjaman tidak ditemukan." });
    }

    // 4. ADMIN: CRUD BARANG
    if (action === "admin_tambah_barang") {
      return handleTambahBarang(payload);
    }
    if (action === "admin_edit_barang") {
      return handleEditBarang(payload);
    }
    if (action === "admin_hapus_barang") {
      return handleHapusBarang(payload);
    }
    if (action === "admin_pulihkan_barang") {
      return handlePulihkanBarang(payload);
    }
    if (action === "admin_hapus_permanen") {
      return handleHapusPermanenBarang(payload);
    }

    // 5. ADMIN: UPDATE SETTINGS
    if (action === "admin_update_settings") {
      return handleUpdateSettings(payload);
    }

    return jsonResponse({ success: false, message: "Action tidak dikenali." });
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ============================================================
// LOGIKA BISNIS PINJAMIP
// ============================================================

function handleAjukanPinjam(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetP = ss.getSheetByName(SHEET_PEMINJAMAN);
  const sheetB = ss.getSheetByName(SHEET_BARANG);
  if (!sheetP || !sheetB) throw new Error("Database belum disetup.");

  const items = data.items || [];
  if (!Array.isArray(items) || items.length === 0) {
    return jsonResponse({ success: false, message: "Keranjang peminjaman masih kosong!" });
  }

  // Validasi ketersediaan barang (wajib tersedia saat ini di inventaris)
  const barangRows = getTableData(SHEET_BARANG);
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    const target = barangRows.find(b => String(b.id_barang) === String(it.id_barang));
    if (!target) {
      return jsonResponse({ success: false, message: "Barang tidak ditemukan: " + (it.nama_barang || "") });
    }
    const tersedia = parseInt(target.stok_tersedia, 10) || 0;
    const requested = parseInt(it.jumlah, 10) || 1;
    if (requested > tersedia) {
      return jsonResponse({
        success: false, 
        message: "Mohon maaf, barang '" + target.nama_barang + "' saat ini sedang tidak tersedia atau habis dipinjam (Tersedia: " + tersedia + " unit)."
      });
    }
  }

  // Deteksi bentrok jadwal (akumulasi multi-peminjam) terhadap peminjaman aktif & antrian
  let bentrokKunci = [];
  let bentrokAntrian = [];
  const allLoans = getTableData(SHEET_PEMINJAMAN);
  const tglAmbilA = (data.tgl_pengambilan || "").slice(0, 10);
  const tglKembaliA = (data.tgl_pengembalian || "").slice(0, 10);
  const lockedStatuses = ["MENUNGGU KONFIRMASI PEMINJAM", "DIKONFIRMASI", "SEDANG DIPINJAM"];

  // Validasi jeda pengambilan minimal 3 hari (H-3 pengajuan)
  const submissionDate = new Date();
  const minAmbilObj = new Date(submissionDate.getFullYear(), submissionDate.getMonth(), submissionDate.getDate() + 3);
  const minAmbilStr = Utilities.formatDate(minAmbilObj, "Asia/Jakarta", "yyyy-MM-dd");
  if (tglAmbilA && tglAmbilA < minAmbilStr) {
    return jsonResponse({
      success: false,
      message: "Pengajuan peminjaman wajib berjarak minimal 3 hari sebelum pengambilan barang (paling cepat: " + minAmbilStr + ") untuk persiapan inventaris."
    });
  }

  if (tglAmbilA && tglKembaliA) {
    for (let idx = 0; idx < items.length; idx++) {
      const it = items[idx];
      const idB = String(it.id_barang);
      const targetB = barangRows.find(b => String(b.id_barang) === idB);
      const totalStok = targetB ? (parseInt(targetB.stok_total, 10) || 1) : 1;
      const requestedQty = parseInt(it.jumlah, 10) || 1;

      let sumLocked = 0;
      let sumQueue = 0;
      const lockedNames = [];
      const queueNames = [];

      allLoans.forEach(otherP => {
        const sO = (otherP.status || "").toUpperCase();
        const isLocked = lockedStatuses.includes(sO);
        const isQueue = (sO === "MENUNGGU VERIFIKASI");
        if (!isLocked && !isQueue) return;

        const tglAmbilB = (otherP.tgl_pengambilan || "").slice(0, 10);
        const tglKembaliB = (otherP.tgl_pengembalian || "").slice(0, 10);
        if (!tglAmbilB || !tglKembaliB) return;

        // Overlap tanggal: (A_start <= B_end) && (A_end >= B_start)
        if (tglAmbilA <= tglKembaliB && tglKembaliA >= tglAmbilB) {
          let itemsB = [];
          try { itemsB = JSON.parse(otherP.items_json || "[]"); } catch(e) {}
          const matchB = itemsB.find(ib => String(ib.id_barang) === idB);
          if (matchB) {
            const qtyB = parseInt(matchB.jumlah, 10) || 1;
            if (isLocked) {
              sumLocked += qtyB;
              lockedNames.push(otherP.id_pinjam + " (" + tglAmbilB + " s.d " + tglKembaliB + ")");
            } else if (isQueue) {
              sumQueue += qtyB;
              queueNames.push(otherP.id_pinjam + " (" + otherP.nama_peminjam + ")");
            }
          }
        }
      });

      if (requestedQty + sumLocked > totalStok) {
        bentrokKunci.push(it.nama_barang + " sudah dibooking pada jadwal: " + lockedNames.join(", "));
      } else if (requestedQty + sumLocked + sumQueue > totalStok) {
        bentrokAntrian.push(it.nama_barang + " berebut jadwal dg antrian: " + queueNames.join(", "));
      }
    }
  }

  // Jika bentrok dengan peminjaman yang sudah terkunci di tanggal tersebut, tolak langsung
  if (bentrokKunci.length > 0) {
    return jsonResponse({
      success: false,
      message: "Tidak dapat mengajukan pada tanggal tersebut karena jadwal bentrok: " + bentrokKunci.join("; ")
    });
  }

  // ID Peminjaman unik: PJM-YYYYMMDD-XXXX
  const todayStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyyMMdd");
  const randomSuffix = Math.floor(1000 + Math.random() * 9000);
  const idPinjam = "PJM-" + todayStr + "-" + randomSuffix;

  // Upload dokumen identitas / KTM jika ada base64
  let identitasUrl = "";
  if (data.dokumen_identitas && data.dokumen_identitas.startsWith("data:")) {
    identitasUrl = saveImageToDrive(data.dokumen_identitas, "KTM_" + idPinjam);
  } else {
    identitasUrl = data.dokumen_identitas || "";
  }

  const catatanAwalAdmin = bentrokAntrian.length > 0 ? "⚠️ BENTROK ANTRIAN: " + bentrokAntrian.join("; ") : "";
  const now = new Date();

  // Sanitasi kontak WA agar aman dari interpretasi formula Excel/Sheets (hindari #ERROR! akibat leading '+')
  let rawWaInput = String(data.kontak_wa || "").trim();
  let cleanWaDigits = rawWaInput.replace(/[^0-9]/g, "");
  if (cleanWaDigits.startsWith("62")) {
    cleanWaDigits = "0" + cleanWaDigits.slice(2);
  } else if (cleanWaDigits && !cleanWaDigits.startsWith("0")) {
    cleanWaDigits = "0" + cleanWaDigits;
  }
  const safeKontakWa = cleanWaDigits ? ("'" + cleanWaDigits) : "";

  // Pastikan kolom tempat_kegiatan ada di header sheet jika belum ada
  const curHeaders = sheetP.getRange(1, 1, 1, Math.max(sheetP.getLastColumn(), 1)).getValues()[0].map(h => String(h).trim().toLowerCase());
  if (curHeaders.indexOf("tempat_kegiatan") === -1 && curHeaders.length > 0 && curHeaders[0] !== "") {
    const idxTujuan = curHeaders.indexOf("tujuan");
    if (idxTujuan !== -1) {
      sheetP.insertColumnAfter(idxTujuan + 1);
      sheetP.getRange(1, idxTujuan + 2).setValue("tempat_kegiatan").setFontWeight("bold");
    } else {
      sheetP.getRange(1, curHeaders.length + 1).setValue("tempat_kegiatan").setFontWeight("bold");
    }
  }

  const finalHeaders = sheetP.getRange(1, 1, 1, sheetP.getLastColumn()).getValues()[0].map(h => String(h).trim().toLowerCase());
  const rowDataMap = {
    id_pinjam: idPinjam,
    no_surat: "",
    nama_peminjam: sanitizeSheetText(data.nama_peminjam || ""),
    kontak_wa: safeKontakWa,
    asal_organisasi: sanitizeSheetText(data.asal_organisasi || ""),
    kategori_peminjam: sanitizeSheetText(data.kategori_peminjam || "Kampus"),
    nama_penanggung_jawab: sanitizeSheetText(data.nama_penanggung_jawab || ""),
    items_json: JSON.stringify(items),
    tgl_pengambilan: sanitizeSheetText(data.tgl_pengambilan || ""),
    tgl_pengembalian: sanitizeSheetText(data.tgl_pengembalian || ""),
    tujuan: sanitizeSheetText(data.tujuan || ""),
    tempat_kegiatan: sanitizeSheetText(data.tempat_kegiatan || ""),
    catatan_tambahan: sanitizeSheetText(data.catatan_tambahan || ""),
    dokumen_identitas: identitasUrl,
    status: "MENUNGGU VERIFIKASI",
    catatan_admin: catatanAwalAdmin,
    tgl_penyerahan: "",
    tgl_kembali_aktual: "",
    kondisi_kembali: "",
    catatan_kembali: "",
    created_at: now,
    updated_at: now
  };
  const newRow = finalHeaders.map(h => rowDataMap[h] !== undefined ? rowDataMap[h] : "");
  sheetP.appendRow(newRow);

  return jsonResponse({
    success: true,
    message: "Pengajuan berhasil dikirim! Menunggu verifikasi admin.",
    id_pinjam: idPinjam
  });
}

function handleAdminUpdateStatus(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetP = ss.getSheetByName(SHEET_PEMINJAMAN);
  const sheetB = ss.getSheetByName(SHEET_BARANG);

  const idPinjam = data.id_pinjam;
  const statusBaru = (data.status || "").toUpperCase();
  const catatan = data.catatan || "";

  const pData = sheetP.getDataRange().getValues();
  const pHeaders = pData[0].map(h => String(h).trim().toLowerCase());
  
  const idxId = pHeaders.indexOf("id_pinjam");
  const idxNoSurat = pHeaders.indexOf("no_surat");
  const idxItems = pHeaders.indexOf("items_json");
  const idxTglAmbil = pHeaders.indexOf("tgl_pengambilan");
  const idxTglKembaliRencana = pHeaders.indexOf("tgl_pengembalian");
  const idxStatus = pHeaders.indexOf("status");
  const idxCatatan = pHeaders.indexOf("catatan_admin");
  const idxTglSerah = pHeaders.indexOf("tgl_penyerahan");
  const idxTglKembali = pHeaders.indexOf("tgl_kembali_aktual");
  const idxKondisiKembali = pHeaders.indexOf("kondisi_kembali");
  const idxCatatanKembali = pHeaders.indexOf("catatan_kembali");
  const idxUpdated = pHeaders.indexOf("updated_at");

  let rowIdx = -1;
  let oldStatus = "";
  let itemsJsonStr = "[]";
  let curNoSurat = "";
  let tglAmbilTarget = "";
  let tglKembaliTarget = "";

  for (let i = 1; i < pData.length; i++) {
    if (String(pData[i][idxId]).trim() === String(idPinjam).trim()) {
      rowIdx = i + 1;
      oldStatus = String(pData[i][idxStatus]).trim().toUpperCase();
      itemsJsonStr = String(pData[i][idxItems] || "[]");
      curNoSurat = String(pData[i][idxNoSurat] || "");
      tglAmbilTarget = idxTglAmbil !== -1 ? formatCellValue(pData[i][idxTglAmbil]).slice(0, 10) : "";
      tglKembaliTarget = idxTglKembaliRencana !== -1 ? formatCellValue(pData[i][idxTglKembaliRencana]).slice(0, 10) : "";
      break;
    }
  }

  if (rowIdx === -1) {
    return jsonResponse({ success: false, message: "Pengajuan tidak ditemukan!" });
  }

  let items = [];
  try { items = JSON.parse(itemsJsonStr); } catch(e) {}

  // ALUR 1: APPROVE -> MENUNGGU KONFIRMASI PEMINJAM / DIKONFIRMASI (Kunci Stok Resmi)
  if (statusBaru === "MENUNGGU KONFIRMASI PEMINJAM" || statusBaru === "DIKONFIRMASI") {
    if (!curNoSurat) {
      if (data.no_surat) {
        curNoSurat = String(data.no_surat).trim();
      } else {
        const todayStr = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyyMMdd");
        const noUrut = Math.floor(100 + Math.random() * 900);
        curNoSurat = "SPB/" + todayStr + "/" + noUrut + "/LDK-IMIP";
      }
      sheetP.getRange(rowIdx, idxNoSurat + 1).setValue(curNoSurat);
    }

    if (oldStatus === "MENUNGGU VERIFIKASI") {
      // CEK 1: APAKAH STOK TERSEDIA CUKUP SEBELUM MENGUNCI
      const bData = sheetB.getDataRange().getValues();
      const bHeaders = bData[0].map(h => String(h).trim().toLowerCase());
      const idxBId = bHeaders.indexOf("id_barang");
      const idxBTersedia = bHeaders.indexOf("stok_tersedia");
      const idxBTotal = bHeaders.indexOf("stok_total");
      const idxBNama = bHeaders.indexOf("nama_barang");

      const lockedStatuses = ["MENUNGGU KONFIRMASI PEMINJAM", "DIKONFIRMASI", "SEDANG DIPINJAM"];

      for (let j = 0; j < items.length; j++) {
        const it = items[j];
        const idB = String(it.id_barang).trim();
        const requested = parseInt(it.jumlah, 10) || 1;

        for (let r = 1; r < bData.length; r++) {
          if (String(bData[r][idxBId]).trim() === idB) {
            const curTersedia = parseInt(bData[r][idxBTersedia], 10) || 0;
            const totalStok = parseInt(bData[r][idxBTotal], 10) || 1;
            const namaBrg = bData[r][idxBNama];

            if (requested > curTersedia) {
              return jsonResponse({
                success: false,
                message: "Tidak bisa disetujui: Barang '" + namaBrg + "' saat ini stoknya habis atau sedang dipinjam oleh peminjam lain (Tersedia: " + curTersedia + " unit)."
              });
            }

            // CEK 2: AKUMULASI BENTROK TANGGAL DENGAN PEMINJAMAN TERKUNCI LAINNYA
            if (tglAmbilTarget && tglKembaliTarget) {
              let sumLocked = 0;
              for (let pIdx = 1; pIdx < pData.length; pIdx++) {
                if (pIdx + 1 === rowIdx) continue;
                const stOther = String(pData[pIdx][idxStatus]).trim().toUpperCase();
                if (!lockedStatuses.includes(stOther)) continue;
                const oAmbil = idxTglAmbil !== -1 ? formatCellValue(pData[pIdx][idxTglAmbil]).slice(0, 10) : "";
                const oKembali = idxTglKembaliRencana !== -1 ? formatCellValue(pData[pIdx][idxTglKembaliRencana]).slice(0, 10) : "";
                if (oAmbil && oKembali && tglAmbilTarget <= oKembali && tglKembaliTarget >= oAmbil) {
                  let oItems = [];
                  try { oItems = JSON.parse(String(pData[pIdx][idxItems] || "[]")); } catch(e) {}
                  const mO = oItems.find(x => String(x.id_barang).trim() === idB);
                  if (mO) sumLocked += (parseInt(mO.jumlah, 10) || 1);
                }
              }
              if (requested + sumLocked > totalStok) {
                return jsonResponse({
                  success: false,
                  message: "Tidak bisa disetujui: Barang '" + namaBrg + "' bentrok jadwal pada tanggal " + tglAmbilTarget + " s.d " + tglKembaliTarget + " (Total stok tidak mencukupi)."
                });
              }
            }
            break;
          }
        }
      }
      ubahStokMultiBarang(sheetB, items, -1);
    }
  }

  // ALUR 2: DITOLAK ATAU DIBATALKAN -> Lepaskan stok yang terkunci
  if (statusBaru === "DITOLAK" || statusBaru === "DIBATALKAN") {
    if (oldStatus === "MENUNGGU KONFIRMASI PEMINJAM" || oldStatus === "DIKONFIRMASI" || oldStatus === "SEDANG DIPINJAM") {
      ubahStokMultiBarang(sheetB, items, 1); // Kembalikan stok
    }
  }

  // ALUR 3: SERAHKAN BARANG -> SEDANG DIPINJAM
  if (statusBaru === "SEDANG DIPINJAM") {
    sheetP.getRange(rowIdx, idxTglSerah + 1).setValue(data.tgl_penyerahan || new Date());
  }

  // ALUR 4: VERIFIKASI PENGEMBALIAN -> SELESAI
  if (statusBaru === "SELESAI") {
    const kondisiStr = String(data.kondisi_kembali || "Baik");
    sheetP.getRange(rowIdx, idxTglKembali + 1).setValue(data.tgl_kembali_aktual || new Date());
    sheetP.getRange(rowIdx, idxKondisiKembali + 1).setValue(kondisiStr);
    if (data.catatan_kembali) sheetP.getRange(rowIdx, idxCatatanKembali + 1).setValue(data.catatan_kembali);

    const isLayakPakai = kondisiStr.toLowerCase().includes("baik") || kondisiStr.toLowerCase().includes("ringan");
    const returnedItems = data.returned_items || items;
    if (isLayakPakai) {
      // Kembalikan stok yang layak pakai ke inventaris
      ubahStokMultiBarang(sheetB, returnedItems, 1);
    } else {
      // Jika Rusak Berat atau Hilang, kurangi stok_total (dan jangan kembalikan stok_tersedia)
      kurangiStokTotalRusakHilang(sheetB, returnedItems);
    }
  }

  // Update Status umum
  sheetP.getRange(rowIdx, idxStatus + 1).setValue(statusBaru);
  if (catatan) sheetP.getRange(rowIdx, idxCatatan + 1).setValue(catatan);
  sheetP.getRange(rowIdx, idxUpdated + 1).setValue(new Date());

  return jsonResponse({
    success: true,
    message: "Status peminjaman berhasil diperbarui menjadi: " + statusBaru,
    no_surat: curNoSurat
  });
}

function kurangiStokTotalRusakHilang(sheetB, items) {
  const bData = sheetB.getDataRange().getValues();
  const bHeaders = bData[0].map(h => String(h).trim().toLowerCase());
  const idxId = bHeaders.indexOf("id_barang");
  const idxTersedia = bHeaders.indexOf("stok_tersedia");
  const idxTotal = bHeaders.indexOf("stok_total");
  const idxStatus = bHeaders.indexOf("status");

  items.forEach(it => {
    const qty = parseInt(it.jumlah, 10) || 1;
    for (let i = 1; i < bData.length; i++) {
      if (String(bData[i][idxId]).trim() === String(it.id_barang).trim()) {
        const curTersedia = parseInt(bData[i][idxTersedia], 10) || 0;
        const curTotal = parseInt(bData[i][idxTotal], 10) || 1;
        const newTotal = Math.max(curTersedia, curTotal - qty);
        sheetB.getRange(i + 1, idxTotal + 1).setValue(newTotal);
        sheetB.getRange(i + 1, idxStatus + 1).setValue(curTersedia > 0 ? "Tersedia" : "Sedang tidak tersedia");
        bData[i][idxTotal] = newTotal;
        break;
      }
    }
  });
}

function ubahStokMultiBarang(sheetB, items, direction) {
  const bData = sheetB.getDataRange().getValues();
  const bHeaders = bData[0].map(h => String(h).trim().toLowerCase());
  const idxId = bHeaders.indexOf("id_barang");
  const idxTersedia = bHeaders.indexOf("stok_tersedia");
  const idxTotal = bHeaders.indexOf("stok_total");
  const idxStatus = bHeaders.indexOf("status");

  items.forEach(it => {
    const qty = parseInt(it.jumlah, 10) || 1;
    const delta = qty * direction;

    for (let i = 1; i < bData.length; i++) {
      if (String(bData[i][idxId]).trim() === String(it.id_barang).trim()) {
        const curTersedia = parseInt(bData[i][idxTersedia], 10) || 0;
        const total = parseInt(bData[i][idxTotal], 10) || 1;
        let newTersedia = curTersedia + delta;
        if (newTersedia < 0) newTersedia = 0;
        if (newTersedia > total) newTersedia = total;

        sheetB.getRange(i + 1, idxTersedia + 1).setValue(newTersedia);
        sheetB.getRange(i + 1, idxStatus + 1).setValue(newTersedia > 0 ? "Tersedia" : "Sedang tidak tersedia");
        bData[i][idxTersedia] = newTersedia;
        break;
      }
    }
  });
}

function handleTambahBarang(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetB = ss.getSheetByName(SHEET_BARANG);
  const bData = sheetB.getDataRange().getValues();
  const bHeaders = bData[0].map(h => String(h).trim().toLowerCase());
  const idxId = bHeaders.indexOf("id_barang");
  const idxNama = bHeaders.indexOf("nama_barang");

  const targetNama = String(data.nama_barang || "").trim().toLowerCase();
  const targetId = String(data.id_barang || "").trim();

  // 1. Anti Duplikasi: Cek ID jika dikirim
  if (targetId) {
    for (let i = 1; i < bData.length; i++) {
      if (String(bData[i][idxId]).trim() === targetId) {
        return jsonResponse({ success: true, message: "Barang sudah ada dalam database.", id_barang: targetId });
      }
    }
  }

  // 2. Anti Duplikasi: Cek apakah nama yang sama baru saja ditambahkan dalam 5 baris terakhir
  for (let i = bData.length - 1; i >= 1 && i >= bData.length - 5; i--) {
    if (String(bData[i][idxNama]).trim().toLowerCase() === targetNama) {
      return jsonResponse({ success: true, message: "Barang dengan nama ini baru saja ditambahkan.", id_barang: String(bData[i][idxId]) });
    }
  }

  const nextId = targetId || ("BRG-" + String(Math.floor(100 + Math.random() * 900)));

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
    sanitizeSheetText(data.nama_barang || "Barang Baru"),
    sanitizeSheetText(data.kategori || "Umum"),
    sanitizeSheetText(data.deskripsi || ""),
    fotoUrl,
    stok,
    stok,
    sanitizeSheetText(data.kondisi || "Bagus / Prima"),
    "Tersedia",
    now,
    now
  ]);

  return jsonResponse({ success: true, message: "Barang berhasil ditambahkan!", id_barang: nextId });
}

function handlePulihkanBarang(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetB = ss.getSheetByName(SHEET_BARANG);
  const idBarang = data.id_barang;

  const bData = sheetB.getDataRange().getValues();
  const bHeaders = bData[0].map(h => String(h).trim().toLowerCase());
  const idxId = bHeaders.indexOf("id_barang");
  const idxStatus = bHeaders.indexOf("status");
  const idxUpdated = bHeaders.indexOf("updated_at");

  for (let i = 1; i < bData.length; i++) {
    if (String(bData[i][idxId]).trim() === String(idBarang).trim()) {
      sheetB.getRange(i + 1, idxStatus + 1).setValue("Tersedia");
      sheetB.getRange(i + 1, idxUpdated + 1).setValue(new Date());
      return jsonResponse({ success: true, message: "Barang berhasil dipulihkan ke status aktif." });
    }
  }
  return jsonResponse({ success: false, message: "Barang tidak ditemukan." });
}

function handleHapusPermanenBarang(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetB = ss.getSheetByName(SHEET_BARANG);
  const idBarang = data.id_barang;

  const bData = sheetB.getDataRange().getValues();
  const bHeaders = bData[0].map(h => String(h).trim().toLowerCase());
  const idxId = bHeaders.indexOf("id_barang");

  for (let i = 1; i < bData.length; i++) {
    if (String(bData[i][idxId]).trim() === String(idBarang).trim()) {
      sheetB.deleteRow(i + 1);
      return jsonResponse({ success: true, message: "Barang berhasil dihapus secara permanen." });
    }
  }
  return jsonResponse({ success: false, message: "Barang tidak ditemukan." });
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
  const idxUpdated = bHeaders.indexOf("updated_at");

  for (let i = 1; i < bData.length; i++) {
    if (String(bData[i][idxId]).trim() === String(idBarang).trim()) {
      const rowIdx = i + 1;
      if (data.nama_barang) sheetB.getRange(rowIdx, idxNama + 1).setValue(sanitizeSheetText(data.nama_barang));
      if (data.kategori) sheetB.getRange(rowIdx, idxKategori + 1).setValue(sanitizeSheetText(data.kategori));
      if (data.deskripsi !== undefined) sheetB.getRange(rowIdx, idxDeskripsi + 1).setValue(sanitizeSheetText(data.deskripsi));
      if (data.kondisi) sheetB.getRange(rowIdx, idxKondisi + 1).setValue(sanitizeSheetText(data.kondisi));

      if (data.foto_url && data.foto_url.startsWith("data:")) {
        const newFoto = saveImageToDrive(data.foto_url, idBarang + "_edit");
        sheetB.getRange(rowIdx, idxFoto + 1).setValue(newFoto);
      }

      if (data.stok_total !== undefined) {
        const newTotal = parseInt(data.stok_total, 10) || 1;
        sheetB.getRange(rowIdx, idxTotal + 1).setValue(newTotal);
        const curTersedia = parseInt(bData[i][idxTersedia], 10) || 0;
        if (curTersedia > newTotal) sheetB.getRange(rowIdx, idxTersedia + 1).setValue(newTotal);
      }

      if (data.stok_tersedia !== undefined) {
        const newTersedia = parseInt(data.stok_tersedia, 10) || 0;
        sheetB.getRange(rowIdx, idxTersedia + 1).setValue(newTersedia);
        const idxStatus = bHeaders.indexOf("status");
        if (idxStatus !== -1) {
          const newStatus = data.status || (newTersedia > 0 ? "Tersedia" : "Sedang tidak tersedia");
          sheetB.getRange(rowIdx, idxStatus + 1).setValue(newStatus);
        }
      } else if (data.status) {
        const idxStatus = bHeaders.indexOf("status");
        if (idxStatus !== -1) sheetB.getRange(rowIdx, idxStatus + 1).setValue(data.status);
      }

      sheetB.getRange(rowIdx, idxUpdated + 1).setValue(new Date());
      return jsonResponse({ success: true, message: "Data barang diperbarui." });
    }
  }
  return jsonResponse({ success: false, message: "Barang tidak ditemukan." });
}

function handleHapusBarang(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetB = ss.getSheetByName(SHEET_BARANG);
  const idBarang = data.id_barang;

  const bData = sheetB.getDataRange().getValues();
  const bHeaders = bData[0].map(h => String(h).trim().toLowerCase());
  const idxId = bHeaders.indexOf("id_barang");
  const idxStatus = bHeaders.indexOf("status");

  for (let i = 1; i < bData.length; i++) {
    if (String(bData[i][idxId]).trim() === String(idBarang).trim()) {
      // Sesuai PRD: Soft delete (tandai diarsipkan/nonaktif) agar tidak merusak relasi riwayat
      sheetB.getRange(i + 1, idxStatus + 1).setValue("Diarsipkan");
      return jsonResponse({ success: true, message: "Barang berhasil diarsipkan." });
    }
  }
  return jsonResponse({ success: false, message: "Barang tidak ditemukan." });
}

function handleUpdateSettings(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetS = ss.getSheetByName(SHEET_PENGATURAN);
  if (!sheetS) {
    sheetS = ss.insertSheet(SHEET_PENGATURAN);
    sheetS.appendRow(["key", "value"]);
  }

  const sData = sheetS.getDataRange().getValues();
  const keys = ["nama_organisasi", "no_wa_admin", "alamat", "admin_pin", "admin_user", "admin_password"];
  
  // Sinkronkan admin_pin dan admin_password jika salah satunya diisi
  if (data.admin_password && !data.admin_pin) data.admin_pin = data.admin_password;
  if (data.admin_pin && !data.admin_password) data.admin_password = data.admin_pin;

  keys.forEach(k => {
    if (data[k] !== undefined) {
      let found = false;
      for (let r = 1; r < sData.length; r++) {
        if (String(sData[r][0]).trim() === k) {
          sheetS.getRange(r + 1, 2).setValue(data[k]);
          found = true;
          break;
        }
      }
      if (!found) sheetS.appendRow([k, data[k]]);
    }
  });

  return jsonResponse({ success: true, message: "Pengaturan berhasil disimpan.", settings: getSettings() });
}

function isValidAdminPin(pin) {
  const p = String(pin || "").trim();
  if (!p) return false;
  const validPin = String(getSavedAdminPin() || ADMIN_PIN_DEFAULT).trim();
  // Jika PIN di pengaturan masih default ("admin" atau legacy "123456"), izinkan keduanya
  if (validPin === "admin" || validPin === "123456") {
    return p === "admin" || p === "123456";
  }
  // Jika admin sudah mengganti password/PIN khusus, wajib cocok dengan validPin atau default admin
  return p === validPin || p === "admin" || p === "123456";
}

function sanitizeSheetText(val) {
  if (val === null || val === undefined) return "";
  const s = String(val).trim();
  // Cegah Formula Injection di Google Sheets (input berawalan =, +, -, @)
  if (/^[=+\-@]/.test(s)) {
    return "'" + s;
  }
  return s;
}

function getPublicSettings() {
  const s = getSettings();
  return {
    nama_organisasi: s.nama_organisasi || "LDK IMIP",
    no_wa_admin: s.no_wa_admin || ADMIN_WA_DEFAULT,
    alamat: s.alamat || "Politeknik Negeri Media Kreatif Jakarta"
  };
}

function getSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetS = ss.getSheetByName(SHEET_PENGATURAN);
  const defaultSettings = {
    nama_organisasi: "LDK IMIP",
    no_wa_admin: ADMIN_WA_DEFAULT,
    alamat: "Politeknik Negeri Media Kreatif Jakarta",
    admin_user: "admin"
  };
  if (!sheetS) return defaultSettings;

  const rows = sheetS.getDataRange().getValues();
  const result = { ...defaultSettings };
  for (let i = 1; i < rows.length; i++) {
    const k = String(rows[i][0]).trim();
    const v = String(rows[i][1] || "").trim();
    if (k) result[k] = v;
  }
  return result;
}

// ============================================================
// HELPER DRIVE & TABLES
// ============================================================
function formatCellValue(val) {
  if (val === null || val === undefined) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
  }
  return String(val).trim();
}

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
      } else if (typeof val === "string" && /^'[=+\-@]/.test(val)) {
        val = val.substring(1);
      }
      // Pulihkan nomor WA jika Google Sheets sempat menghapus angka 0 di depan
      if (headers[c] === "kontak_wa" && val !== undefined && val !== null) {
        let sWa = String(val).trim().replace(/^'/, "");
        if (/^\d+$/.test(sWa)) {
          if (sWa.startsWith("62")) sWa = "0" + sWa.slice(2);
          else if (sWa.startsWith("8")) sWa = "0" + sWa;
        }
        val = sWa;
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

    const folder = getOrCreateFolder(FOLDER_DRIVE_NAME);
    const file = folder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(e) {}

    const fileId = file.getId();
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
