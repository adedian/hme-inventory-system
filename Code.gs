/**
 * HME INVENTORY MANAGEMENT SYSTEM
 * Designed by: Ade Dian Sukmana, S.Kom
 * Backend: Google Apps Script
 */

// Konfigurasi Email Admin (Silakan ubah dengan email tujuan notifikasi)
var ADMIN_EMAIL = "email.admin@domain.com"; 

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Sistem Peminjaman HME')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ================= DATABASE SELECTOR =================
function getSheet(name) {
  // Masukkan ID Spreadsheet Anda di sini
  var ss = SpreadsheetApp.openById("MASUKKAN_ID_SPREADSHEET_ANDA_DISINI");
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if(name === "Data Barang Masuk-Keluar") {
      sheet.appendRow(["No", "Nama Alat", "Nama Pembawa", "Tanggal Dibawa", "Tanggal Kembali", "Tim", "Lokasi", "Keterangan", "Status", "ID Transaksi", "Username", "ID Barang"]);
    } else if(name === "Master_Barang") {
      sheet.appendRow(["ID Barang", "Nama Barang", "URL Foto", "Kategori", "Status"]);
    } else if(name === "Data_Users") {
      sheet.appendRow(["Username", "Password", "Role", "Nama Lengkap", "Label"]);
      sheet.appendRow(["admin", "1234", "admin", "Admin HME", "AD"]);
      sheet.appendRow(["gabriel", "1234", "staff", "Gabriel", "GB"]);
    }
  }
  return sheet;
}

// ================= FUNGSI NOTIFIKASI GMAIL =================
function sendHmeEmail(recipient, subject, bodyContent) {
  try {
    MailApp.sendEmail({
      to: recipient,
      subject: "[INVENTORY HME] " + subject,
      htmlBody: `
        <div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #e2e8f0; border-radius: 10px;">
          <h2 style="color: #010031;">Sistem Peminjaman HME</h2>
          <hr style="border: 0; border-top: 1px solid #eee;">
          <div style="padding: 10px 0; color: #4a5568;">${bodyContent}</div>
          <hr style="border: 0; border-top: 1px solid #eee;">
          <p style="font-size: 11px; color: #a0aec0;">Email otomatis. Mohon tidak membalas.</p>
        </div>`
    });
  } catch (e) { console.log("Email error: " + e.toString()); }
}

// ================= UPDATE STATUS MASTER BARANG =================
function updateMasterStatus(idBarang, statusBaru) {
  if (!idBarang) return;
  var sheet = getSheet("Master_Barang");
  var data = sheet.getDataRange().getValues();
  if (sheet.getLastColumn() < 5) { sheet.getRange(1, 5).setValue("Status"); }
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString() == idBarang.toString()) {
      sheet.getRange(i + 1, 5).setValue(statusBaru);
      break;
    }
  }
}

// ================= LOGIKA TRANSAKSI PEMINJAMAN =================
function getLoans() {
  try {
    var data = getSheet("Data Barang Masuk-Keluar").getDataRange().getValues();
    var loans = [];
    for(var i = 1; i < data.length; i++) {
       if(!data[i][9]) continue; 
       loans.push({
          no: (data[i][0] || "").toString(), 
          item: (data[i][1] || "").toString(), 
          namaPembawa: (data[i][2] || "").toString(),
          tglBawa: (data[i][3] instanceof Date) ? Utilities.formatDate(data[i][3], "GMT+7", "dd/MM/yyyy HH.mm") : (data[i][3] || "").toString(),
          tglKembali: (data[i][4] instanceof Date) ? Utilities.formatDate(data[i][4], "GMT+7", "dd/MM/yyyy HH.mm") : (data[i][4] || "").toString(),
          tim: (data[i][5] || "").toString(), 
          lokasi: (data[i][6] || "").toString(), 
          ket: (data[i][7] || "").toString(), 
          status: (data[i][8] || "").toString(),
          id: (data[i][9] || "").toString(), 
          user: (data[i][10] || "").toString(), 
          idBarang: (data[i][11] || "").toString()
       });
    }
    return loans.reverse();
  } catch(e) { return []; }
}

function saveLoan(data) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    var sheet = getSheet("Data Barang Masuk-Keluar");
    var vals = sheet.getDataRange().getValues();
    var rowIndex = -1;
    var currentItem = {};

    for(var i = 1; i < vals.length; i++) { 
      if(vals[i][9] == data.id) { 
        rowIndex = i + 1; 
        currentItem = { idBarang: vals[i][11], namaAlat: vals[i][1], peminjam: vals[i][2] };
        break; 
      } 
    }
    
    if(rowIndex > -1) {
       if(data.status === "returned") {
          sheet.getRange(rowIndex, 5).setValue(Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH.mm"));
          updateMasterStatus(currentItem.idBarang, "Tersedia");
          sendHmeEmail(ADMIN_EMAIL, "Barang Kembali", `<b>${currentItem.peminjam}</b> telah mengembalikan <b>${currentItem.namaAlat}</b>.`);
       } else if (data.status === "approved") {
          updateMasterStatus(currentItem.idBarang, "Dipinjam");
       }
       sheet.getRange(rowIndex, 9).setValue(data.status);
    } else {
       var maxNo = 0;
       for(var j = 1; j < vals.length; j++) {
         var valA = parseInt(vals[j][0], 10);
         if(!isNaN(valA) && valA > maxNo) maxNo = valA;
       }
       sheet.appendRow([maxNo + 1, data.namaAlat, data.namaPembawa, data.tanggalDibawa, "", data.tim, data.lokasi, data.keterangan, data.status, data.id, data.username, data.idBarang]);
    }
    SpreadsheetApp.flush();
    return "Success";
  } catch (e) { return "Error"; } finally { lock.releaseLock(); }
}

function saveMultipleLoans(dataArray) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    var sheet = getSheet("Data Barang Masuk-Keluar");
    var vals = sheet.getDataRange().getValues();
    var maxNo = 0;
    for(var j = 1; j < vals.length; j++) {
      var valA = parseInt(vals[j][0], 10);
      if(!isNaN(valA) && valA > maxNo) maxNo = valA;
    }
    
    var listTools = "<ul>";
    for (var k = 0; k < dataArray.length; k++) {
      var d = dataArray[k];
      maxNo++;
      sheet.appendRow([maxNo, d.namaAlat, d.namaPembawa, d.tanggalDibawa, "", d.tim, d.lokasi, d.keterangan, d.status, d.id, d.username, d.idBarang]);
      listTools += `<li>${d.namaAlat} (${d.idBarang})</li>`;
    }
    listTools += "</ul>";
    sendHmeEmail(ADMIN_EMAIL, "Permohonan Baru", `Ada pengajuan dari <b>${dataArray[0].namaPembawa}</b>:<br>${listTools}<br>Lokasi: ${dataArray[0].lokasi}`);
    SpreadsheetApp.flush();
    return "Success";
  } catch (e) { return "Error"; } finally { lock.releaseLock(); }
}

// ================= LOGIKA INVENTARIS =================
function getInventory() {
  try {
    var sheet = getSheet("Master_Barang");
    var data = sheet.getDataRange().getValues();
    var items = [];
    if (data[0].length < 5) { sheet.getRange(1, 5).setValue("Status"); }
    for(var i = 1; i < data.length; i++) {
      if(!data[i][0]) continue;
      var st = (data[i].length > 4 && data[i][4]) ? data[i][4].toString() : "";
      if(!st) { 
        st = "Tersedia";
        sheet.getRange(i+1, 5).setValue("Tersedia"); 
      }
      items.push({ 
        id: data[i][0].toString(), 
        nama: (data[i][1] || "").toString(), 
        foto: (data[i][2] || "").toString(), 
        kategori: (data[i][3] || "").toString(), 
        status_master: st 
      });
    }
    return JSON.stringify(items);
  } catch(e) { return JSON.stringify([]); }
}

function dbAddInventory(item) { getSheet("Master_Barang").appendRow([item.id, item.nama, item.foto, item.kategori, "Tersedia"]); return "Success"; }
function dbDeleteInventory(id) { var s = getSheet("Master_Barang"); var d = s.getDataRange().getValues(); for(var i = 1; i < d.length; i++) { if(d[i][0] == id) { s.deleteRow(i + 1); break; } } return "Success"; }

// ================= LOGIKA DATABASE USER =================
function getAllUsers() { try { var data = getSheet("Data_Users").getDataRange().getValues(); var usersObj = {}; for(var i = 1; i < data.length; i++) { if(!data[i][0]) continue; usersObj[data[i][0].toString().toLowerCase()] = { pass: data[i][1].toString(), role: data[i][2].toString(), name: data[i][3].toString(), label: data[i][4].toString() }; } return JSON.stringify(usersObj); } catch(e) { return "{}"; } }
function dbAddUser(u) { getSheet("Data_Users").appendRow([u.username, u.pass, u.role, u.name, u.label]); return "Success"; }
function dbDeleteUser(un) { var s = getSheet("Data_Users"); var d = s.getDataRange().getValues(); for(var i=1; i<d.length; i++) { if(d[i][0].toLowerCase() === un.toLowerCase()) { s.deleteRow(i+1); break; } } return "Success"; }
function dbChangePwd(un, newPwd) { var s = getSheet("Data_Users"); var d = s.getDataRange().getValues(); for(var i=1; i<d.length; i++) { if(d[i][0].toLowerCase() === un.toLowerCase()) { s.getRange(i+1, 2).setValue(newPwd); break; } } return "Success"; }

// ================= REKAP & EXPORT =================
function getMonthlyStats(m,y) { 
  var data = getSheet("Data Barang Masuk-Keluar").getDataRange().getValues(); 
  var s = { total:0, approved:0, returned:0, pending:0, rejected:0, topItems:{} }; 
  for (var i = 1; i < data.length; i++) { 
    var tgl = data[i][3]; if (!tgl || !data[i][9]) continue; 
    var mm = -1, yy = -1; 
    if (tgl instanceof Date) { mm = tgl.getMonth()+1; yy = tgl.getFullYear(); } 
    else { var p = tgl.toString().split(' ')[0].split('/'); mm = parseInt(p[1]); yy = parseInt(p[2]); } 
    if (mm == m && yy == y) { 
      s.total++; var st = data[i][8]; 
      if(st=='approved') s.approved++; 
      if(st=='returned') s.returned++; 
      if(st=='pending') s.pending++; 
      if(st=='rejected') s.rejected++; 
      s.topItems[data[i][1]] = (s.topItems[data[i][1]] || 0) + 1; 
    } 
  } 
  return s; 
}

function generateExcelReport(m,y) { var ssName = "Laporan_HME_" + m + "_" + y; var ss = SpreadsheetApp.create(ssName); var reportSheet = ss.getSheets()[0]; var masterData = getSheet("Master_Barang").getDataRange().getValues(); var photoMap = {}; for(var j=1; j<masterData.length; j++) { photoMap[masterData[j][0]] = masterData[j][2]; } reportSheet.appendRow(["No", "ID Barang", "Foto Alat", "Nama Alat", "Peminjam", "Tgl Pinjam", "Tgl Kembali", "Lokasi", "Status"]); var loanData = getSheet("Data Barang Masuk-Keluar").getDataRange().getValues(); var count = 1; for (var i = 1; i < loanData.length; i++) { var tgl = loanData[i][3]; if(!tgl || !loanData[i][9]) continue; var mm = -1, yy = -1; if (tgl instanceof Date) { mm = tgl.getMonth()+1; yy = tgl.getFullYear(); } else { var p = tgl.toString().split(' ')[0].split('/'); mm = parseInt(p[1]); yy = parseInt(p[2]); } if (mm == m && yy == y) { var idB = loanData[i][11]; var fotoUrl = photoMap[idB] || ""; var fotoFormula = fotoUrl ? '=IMAGE("' + fotoUrl + '", 1)' : "No Photo"; reportSheet.appendRow([count++, idB, fotoFormula, loanData[i][1], loanData[i][2], loanData[i][3], loanData[i][4], loanData[i][6], loanData[i][8]]); } } reportSheet.setRowHeights(2, reportSheet.getLastRow() > 1 ? reportSheet.getLastRow() : 2, 70); reportSheet.setColumnWidth(3, 120); reportSheet.setColumnWidth(4, 150); reportSheet.setColumnWidth(5, 120); reportSheet.setColumnWidth(6, 130); reportSheet.setColumnWidth(7, 130); reportSheet.getRange("A1:I1").setFontWeight("bold").setBackground("#f0f0f0"); return ss.getUrl(); }