/**
 * Google Apps Script API cho Dashboard Quản Lý Tổng Hợp CDX
 * CHUYỂN ĐỔI THÀNH REST API ĐỂ DÙNG VỚI GITHUB PAGES
 */

function doGet(e) {
  const action = e.parameter.action;
  const args = e.parameter.args ? JSON.parse(e.parameter.args) : [];
  
  let result;
  try {
    switch (action) {
      case 'getInitialData':
        result = getInitialData();
        break;
      case 'loginUser':
        result = loginUser(args[0], args[1]);
        break;
      case 'saveData':
        result = saveData(args[0], args[1]);
        break;
      case 'updateData':
        result = updateData(args[0], args[1], args[2]);
        break;
      case 'deleteData':
        result = deleteData(args[0], args[1]);
        break;
      default:
        // Nếu không truyền action, mặc định serve HTML (để vẫn dùng được trên GAS nếu cần)
        return HtmlService.createTemplateFromFile('index')
          .evaluate()
          .setTitle('Hệ Thống Quản Lý CDX (GAS Mirror)')
          .addMetaTag('viewport', 'width=device-width, initial-scale=1')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
  } catch (err) {
    result = { status: 'error', message: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  // Dự phòng cho các request POST nếu cần
  return doGet(e);
}

// === CONFIGURATION ===
const CONFIG = {
  SPREADSHEET_ID: '1JSII6jKG1of1CsCYsn_75Va8dGLIYAsBd3JIIaP1Djc',
  SHEETS: {
    USER: 'User',
    CHIPHI: 'Chiphi',
    CHIPHI_CHITIET: 'Chiphichitiet',
    PHIEU_NHAP_XUAT: 'PhieuNhapXuat',
    PNX_CHI_TIET: 'PNXChiTiet',
    VAT_LIEU: 'VatLieu',
    DS_KHO: 'DS_kho',
    TON_KHO: 'Tonkho',
    PHIEU_CHUYEN_KHO: 'Phieuchuyenkho',
    CHUYEN_KHO_CHI_TIET: 'Chuyenkhochitiet',
    DOI_TAC: 'Doitac',
    CHAM_CONG: 'Chamcong',
    LICH_SU_LUONG: 'LichSuLuong',
    GIAO_DICH_LUONG: 'GiaoDichLuong',
    BANG_LUONG_THANG: 'BangLuongThang'
  }
};

function getSS() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

/**
 * Lấy dữ liệu ban đầu cho ứng dụng
 */
function getInitialData() {
  try {
    const ss = getSS();
    const results = {};
    
    for (const [key, name] of Object.entries(CONFIG.SHEETS)) {
      const sheet = ss.getSheetByName(name);
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        if (data.length > 1) {
          results[name] = rangeToObjects(data);
        } else {
          results[name] = [];
        }
      }
    }
    
    return results;
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

function rangeToObjects(range) {
  const headers = range[0];
  const data = range.slice(1);
  return data.map(row => {
    const obj = {};
    headers.forEach((header, i) => {
      let val = row[i];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss'Z'");
      }
      obj[header] = val;
    });
    return obj;
  });
}

function loginUser(username, password) {
  const ss = getSS();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.USER);
  const data = rangeToObjects(sheet.getDataRange().getValues());
  
  const user = data.find(u => u['ID'] == username && u['App_pass'] == password);
  if (user) {
    return {
      user: {
        id: user['ID'],
        name: user['Họ và tên'],
        role: user['Chức vụ'] || 'Nhân viên'
      }
    };
  }
  return { error: 'Sai tài khoản hoặc mật khẩu' };
}

function saveData(sheetName, data) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName(sheetName);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    const newRow = headers.map(h => data[h] || '');
    sheet.appendRow(newRow);
    
    return { status: 'success' };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

function updateData(sheetName, data, rowIndex) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName(sheetName);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // rowIndex từ Frontend là index trong mảng (0-indexed, không tính header)
    // Trong Google Sheets, dòng đầu tiên là 1, header là dòng 1 => dòng dữ liệu đầu tiên là 2
    const sheetRowIndex = parseInt(rowIndex) + 2;
    
    const rowValues = headers.map(h => data[h] !== undefined ? data[h] : '');
    sheet.getRange(sheetRowIndex, 1, 1, headers.length).setValues([rowValues]);
    
    return { status: 'success' };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

function deleteData(sheetName, rowIndex) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName(sheetName);
    const sheetRowIndex = parseInt(rowIndex) + 2;
    
    // Thay vì xóa hẳn, ta đánh dấu 'X' vào cột 'Delete' nếu có
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const deleteColIdx = headers.indexOf('Delete');
    
    if (deleteColIdx !== -1) {
      sheet.getRange(sheetRowIndex, deleteColIdx + 1).setValue('X');
    } else {
      // Nếu không có cột Delete, tiến hành xóa dòng (Cẩn trọng)
      sheet.deleteRow(sheetRowIndex);
    }
    
    return { status: 'success' };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}
