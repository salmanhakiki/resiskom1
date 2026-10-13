// =====================================================
// GOOGLE APPS SCRIPT - WEB FORM HANDLER
// Terima POST request dari web form, update spreadsheet
// =====================================================

const SPREADSHEET_ID = '1cgEfwLtxTXynCAQgu8NFaOW6e2Yr2lZbC74fhhMCOHA';
const SHEET_KAS = 'KAS HIMPUNAN';
const SHEET_RIWAYAT = 'Riwayat Cicilan';

// Mapping bulan ke column index (KAS HIMPUNAN)
const BULAN_KAS_COL = {
  'juni': 4, 'juli': 8, 'agustus': 12, 'september': 16,
  'oktober': 20, 'november': 24, 'desember': 28, 'januari': 32,
  'februari': 36, 'maret': 40, 'april': 44, 'mei': 48
};

// =====================================================
// POST HANDLER - Terima request dari web form
// =====================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    
    if (data.type === 'kas') {
      return handleKas(data);
    } else if (data.type === 'jaket') {
      return handleJaket(data);
    }
    
    return createResponse(false, 'Tipe pembayaran tidak valid');
    
  } catch (error) {
    return createResponse(false, error.toString());
  }
}

// =====================================================
// HANDLE KAS - Update KAS HIMPUNAN
// =====================================================
function handleKas(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_KAS);
  
  // Find row untuk nama
  const namaCol = sheet.getRange('C6:C25').getValues();
  let targetRow = -1;
  
  for (let i = 0; i < namaCol.length; i++) {
    if (namaCol[i][0] === data.nama) {
      targetRow = i + 6;
      break;
    }
  }
  
  if (targetRow === -1) {
    return createResponse(false, 'Nama tidak ditemukan di sheet');
  }
  
  // Get base column untuk bulan
  const baseCol = BULAN_KAS_COL[data.bulan];
  if (!baseCol) {
    return createResponse(false, 'Bulan tidak valid');
  }
  
  // Update cells untuk setiap minggu
  for (let minggu of data.minggu) {
    const col = baseCol + (parseInt(minggu) - 1);
    sheet.getRange(targetRow, col).setValue(true);
  }
  
  const nominal = data.minggu.length * 5000;
  return createResponse(true, `Kas ${data.bulan} minggu ${data.minggu.join(', ')} - Rp${nominal.toLocaleString('id-ID')}`);
}

// =====================================================
// HANDLE JAKET - Add to Riwayat Cicilan
// =====================================================
function handleJaket(data) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_RIWAYAT);
  
  // Find last row
  const dataRange = sheet.getRange('C7:C200').getValues();
  let lastRow = 6;
  for (let i = 0; i < dataRange.length; i++) {
    if (dataRange[i][0]) {
      lastRow = i + 7;
    }
  }
  
  const newRow = lastRow + 1;
  const today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy');
  
  // Write data (B=Tanggal, C=Nama, D=Nominal, E=Keterangan)
  sheet.getRange(newRow, 2).setValue(today);
  sheet.getRange(newRow, 3).setValue(data.nama);
  sheet.getRange(newRow, 4).setValue(data.nominal);
  sheet.getRange(newRow, 5).setValue(data.keterangan);
  
  return createResponse(true, `${data.keterangan} - Rp${data.nominal.toLocaleString('id-ID')}`);
}

// =====================================================
// CREATE RESPONSE - JSON response untuk frontend
// =====================================================
function createResponse(success, message) {
  const response = {
    success: success,
    message: message
  };
  
  return ContentService
    .createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}
