/*******************************************************
 * QUẢN LÝ VĂN BẢN - GOOGLE APPS SCRIPT
 * Backend: Code.gs (bản đã sửa lỗi)
 *
 * Yêu cầu: tạo script từ trong Google Sheet
 *   (Tiện ích mở rộng -> Apps Script).
 * File giao diện phải tên đúng là: Index
 *
 * Cấu trúc Sheet:
 * 01_CAUHINH, 02_NGUOIDUNG, 03_DANHMUC, 04_VANBAN_DEN,
 * 05_VANBAN_DI, 06_NGUOI_XULY, 07_LICHSU_XULY,
 * 08_LICHSU_EMAIL, 09_NHATKY_HE_THONG
 *******************************************************/

const APP = {
  SHEETS: {
    CONFIG: '01_CAUHINH',
    USERS: '02_NGUOIDUNG',
    CATEGORIES: '03_DANHMUC',
    HANDLERS: '06_NGUOI_XULY',
    INCOMING: '04_VANBAN_DEN',
    OUTGOING: '05_VANBAN_DI',
    HISTORY: '07_LICHSU_XULY',
    EMAIL_LOG: '08_LICHSU_EMAIL',
    AUDIT: '09_NHATKY_HE_THONG'
  },
  ROOT_FOLDER: 'QUAN_LY_VAN_BAN',
  MAX_UPLOAD_MB: 15
};

// Các cột luôn lưu dạng văn bản để Google Sheets không tự đổi kiểu
const TEXT_COLUMNS = ['ID','SO_VAN_BAN','NGUOI_CHU_TRI_ID','NGUOI_PHU_HOP_ID','NGUOI_SOAN_ID',
  'VAN_BAN_ID','FILE_ID','VALUE','ACTIVE','TRANG_THAI','TEN_DANG_NHAP','MAT_KHAU_HASH','SALT'];

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Quản lý văn bản')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * CỔNG API CHO GITHUB PAGES (giao diện đặt ngoài Apps Script gọi bằng fetch POST).
 * Chỉ cho phép 3 lệnh: login, logout, api. Mọi nghiệp vụ khác vẫn đi qua api() có kiểm tra phiên.
 * Body gửi lên: {"fn":"login|logout|api","args":[...]}
 */
function doPost(e) {
  let out;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const args = Array.isArray(body.args) ? body.args : [];
    let result;
    if (body.fn === 'login') result = login.apply(null, args);
    else if (body.fn === 'logout') result = logout.apply(null, args);
    else if (body.fn === 'api') result = api.apply(null, args);
    else throw new Error('Chức năng không hợp lệ.');
    out = { ok: true, data: result === undefined ? null : result };
  } catch (err) {
    out = { ok: false, error: (err && err.message) ? err.message : String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// Dán ID của Google Sheet dữ liệu vào đây nếu script KHÔNG được tạo từ trong Sheet.
// ID là đoạn giữa /d/ và /edit trong link Sheet: https://docs.google.com/spreadsheets/d/<ID>/edit
const SPREADSHEET_ID = '';

function getSS_() {
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (ss) return ss;
  let id = SPREADSHEET_ID;
  if (!id) {
    try { id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID') || ''; } catch (e) {}
  }
  if (!id) {
    throw new Error('Chưa biết Google Sheet dữ liệu. Dán ID của Sheet vào dòng const SPREADSHEET_ID = \'\'; ở đầu file Code.gs (ID là đoạn giữa /d/ và /edit trong link Sheet).');
  }
  try { return SpreadsheetApp.openById(id); }
  catch (e) { throw new Error('Không mở được Google Sheet có ID đã nhập. Kiểm tra lại ID và quyền truy cập của tài khoản chạy script.'); }
}

/* =========================
   KHỞI TẠO
========================= */

function setupSystem() {
  requireOwnerContext_();
  const ss = getSS_();

  const defs = {};
  defs[APP.SHEETS.CONFIG] = [
    ['KEY','VALUE','DESCRIPTION'],
    ['APP_NAME','QUẢN LÝ VĂN BẢN','Tên ứng dụng'],
    ['EMAIL_ENABLED','TRUE','Bật/tắt nhắc việc email'],
    ['REMIND_BEFORE_DAYS','3','Số ngày trước hạn bắt đầu nhắc'],
    ['SEND_TIMES','08:00,15:00','Các giờ gửi, phân cách bằng dấu phẩy'],
    ['REMIND_DUE_TODAY','TRUE','Nhắc khi đến hạn hôm nay'],
    ['OVERDUE_DAYS','MOI_NGAY','MOI_NGAY = nhắc mỗi ngày khi quá hạn; hoặc danh sách ngày, vd 1,3,5'],
    ['MAX_EMAILS_PER_RUN','50','Số email tối đa mỗi lần chạy'],
    ['ROOT_FOLDER_ID','','ID thư mục gốc Google Drive'],
    ['LAST_REMINDER_RUN','','Lần chạy nhắc việc gần nhất']
  ];

  defs[APP.SHEETS.USERS] = [
    ['ID','HO_TEN','EMAIL','BO_PHAN','VAI_TRO','TRANG_THAI','NGAY_TAO','TEN_DANG_NHAP','MAT_KHAU_HASH','SALT']
  ];

  defs[APP.SHEETS.CATEGORIES] = [
    ['ID','NHOM','GIA_TRI','TRANG_THAI','THU_TU'],
    ['1','LOAI_VAN_BAN','Thông tư','Đang dùng','1'],
    ['2','LOAI_VAN_BAN','Nghị định','Đang dùng','2'],
    ['3','LOAI_VAN_BAN','Quyết định','Đang dùng','3'],
    ['4','LOAI_VAN_BAN','Quy trình','Đang dùng','4'],
    ['5','LOAI_VAN_BAN','Nội quy','Đang dùng','5'],
    ['6','LOAI_VAN_BAN','Quy chế','Đang dùng','6'],
    ['7','LOAI_VAN_BAN','Công văn','Đang dùng','7'],
    ['8','LOAI_VAN_BAN','Thông báo','Đang dùng','8'],
    ['9','LOAI_VAN_BAN','Hướng dẫn','Đang dùng','9'],
    ['10','LOAI_VAN_BAN','Kế hoạch','Đang dùng','10'],
    ['11','LOAI_VAN_BAN','Biên bản','Đang dùng','11'],
    ['12','LOAI_VAN_BAN','Báo cáo','Đang dùng','12'],
    ['13','LOAI_VAN_BAN','Tờ trình','Đang dùng','13'],
    ['20','TINH_TRANG','Chưa xử lý','Đang dùng','1'],
    ['21','TINH_TRANG','Đang xử lý','Đang dùng','2'],
    ['22','TINH_TRANG','Chờ phối hợp','Đang dùng','3'],
    ['23','TINH_TRANG','Chờ phê duyệt','Đang dùng','4'],
    ['24','TINH_TRANG','Đã hoàn thành','Đang dùng','5'],
    ['25','TINH_TRANG','Lưu trữ','Đang dùng','6'],
    ['30','MUC_DO','Bình thường','Đang dùng','1'],
    ['31','MUC_DO','Quan trọng','Đang dùng','2'],
    ['32','MUC_DO','Khẩn','Đang dùng','3'],
    ['33','MUC_DO','Rất khẩn','Đang dùng','4']
  ];

  defs[APP.SHEETS.HANDLERS] = [
    ['ID','HO_TEN','EMAIL','BO_PHAN','TRANG_THAI','NGAY_TAO']
  ];

  defs[APP.SHEETS.INCOMING] = [[
    'ID','STT','SO_VAN_BAN','NGAY_VAN_BAN','NGAY_NHAN','LOAI_VAN_BAN',
    'CO_QUAN_BAN_HANH','TEN_VAN_BAN','TRICH_YEU','NGUOI_CHU_TRI_ID',
    'NGUOI_PHU_HOP_ID','BO_PHAN','MUC_DO','HAN_XU_LY','TINH_TRANG',
    'TIEN_DO','TRANG_THAI_HAN','GHI_CHU','FILE_ID','FILE_NAME','FILE_URL',
    'NGUOI_NHAP','NGAY_NHAP','NGAY_CAP_NHAT','ACTIVE'
  ]];

  defs[APP.SHEETS.OUTGOING] = [[
    'ID','STT','SO_VAN_BAN','NGAY_VAN_BAN','LOAI_VAN_BAN','NOI_NHAN',
    'TEN_VAN_BAN','TRICH_YEU','NGUOI_SOAN_ID','TINH_TRANG','FILE_ID',
    'FILE_NAME','FILE_URL','GHI_CHU','NGUOI_NHAP','NGAY_NHAP',
    'NGAY_CAP_NHAT','ACTIVE'
  ]];

  defs[APP.SHEETS.HISTORY] = [[
    'ID','VAN_BAN_ID','LOAI_VAN_BAN','THOI_GIAN','NGUOI_THUC_HIEN',
    'HANH_DONG','NOI_DUNG'
  ]];

  defs[APP.SHEETS.EMAIL_LOG] = [[
    'ID','VAN_BAN_ID','SO_VAN_BAN','EMAIL','THOI_GIAN','LOAI_NHAC',
    'SO_NGAY','KET_QUA','ERROR'
  ]];

  defs[APP.SHEETS.AUDIT] = [[
    'ID','THOI_GIAN','EMAIL','HANH_DONG','DOI_TUONG','DOI_TUONG_ID','CHI_TIET'
  ]];

  Object.keys(defs).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const rows = defs[name];

    // Đặt định dạng văn bản cho các cột nhạy cảm TRƯỚC khi ghi dữ liệu
    const headers = rows[0];
    headers.forEach((h, i) => {
      if (TEXT_COLUMNS.indexOf(h) >= 0 || (name === APP.SHEETS.EMAIL_LOG && h === 'THOI_GIAN')) {
        sh.getRange(1, i + 1, Math.max(sh.getMaxRows(), 2), 1).setNumberFormat('@');
      }
    });

    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    } else if (name === APP.SHEETS.CONFIG) {
      // Bổ sung các cấu hình còn thiếu nếu chạy setup lần nữa
      const existing = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 3).getValues();
      const keys = new Set(existing.slice(1).map(r => String(r[0])));
      rows.slice(1).forEach(r => {
        if (!keys.has(r[0])) sh.appendRow(r);
      });
    }
    // Bổ sung cột còn thiếu (khi nâng cấp từ bản cũ)
    const curHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
    headers.forEach(h => {
      if (curHeaders.indexOf(h) < 0) {
        const col = sh.getLastColumn() + 1;
        if (TEXT_COLUMNS.indexOf(h) >= 0) sh.getRange(1, col, Math.max(sh.getMaxRows(), 2), 1).setNumberFormat('@');
        sh.getRange(1, col).setValue(h);
        curHeaders.push(h);
      }
    });
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold');
  });

  const root = getOrCreateRootFolder_();
  setConfig_('ROOT_FOLDER_ID', root.getId());

  // Tài khoản chạy setup (chủ script): bảo đảm có trong danh sách người xử lý
  const ownerEmail = getCurrentEmail_();
  if (ownerEmail) {
    const handlers = readObjects_(APP.SHEETS.HANDLERS);
    if (!handlers.some(h => String(h.EMAIL).toLowerCase() === ownerEmail.toLowerCase())) {
      appendObject_(APP.SHEETS.HANDLERS, {
        ID: uid_(), HO_TEN: ownerEmail.split('@')[0], EMAIL: ownerEmail,
        BO_PHAN: 'Quản trị', TRANG_THAI: 'Đang hoạt động', NGAY_TAO: now_()
      });
    }
  }
  const loginInfo = bootstrapLogins_(ownerEmail);

  createReminderTrigger_();
  audit_('SETUP', 'SYSTEM', '', 'Khởi tạo/cập nhật hệ thống');
  return 'Đã khởi tạo hệ thống. Thư mục Drive: ' + root.getUrl() + (loginInfo ? '\n' + loginInfo : '');
}

/* =========================
   API GIAO DIỆN
========================= */

function getAppData() {
  requireLogin_();
  const user = getCurrentUser_();
  const incoming = readObjects_(APP.SHEETS.INCOMING).filter(isActive_);
  const outgoing = readObjects_(APP.SHEETS.OUTGOING).filter(isActive_);
  const incomingAll = readObjects_(APP.SHEETS.INCOMING);
  const archivedCount = incomingAll.filter(r => String(r.ACTIVE).toUpperCase() === 'FALSE'
    || String(r.TINH_TRANG) === 'Lưu trữ').length;

  const stats = buildStats_(incoming, outgoing);
  stats.archived = archivedCount;

  return {
    user: user,
    config: getConfigObject_(),
    categories: getCategories_(),
    handlers: readObjects_(APP.SHEETS.HANDLERS).filter(r => String(r.TRANG_THAI) !== 'Ngừng sử dụng'),
    users: isAdmin_(user) ? readObjects_(APP.SHEETS.USERS).map(publicUser_) : [],
    incoming: enrichIncoming_(incoming),
    outgoing: enrichOutgoing_(outgoing),
    history: readObjects_(APP.SHEETS.HISTORY).slice(-200).reverse(),
    emailLog: isAdmin_(user) ? readObjects_(APP.SHEETS.EMAIL_LOG).slice(-200).reverse() : [],
    stats: stats,
    triggerCount: isAdmin_(user) ? getReminderTriggerCount_() : 0
  };
}

function saveIncoming(payload) {
  requirePermission_(['ADMIN','QUAN_LY','NHAN_VIEN']);
  payload = payload || {};
  validateRequired_(payload, ['SO_VAN_BAN','NGAY_VAN_BAN','TEN_VAN_BAN','TRICH_YEU','NGUOI_CHU_TRI_ID']);

  const user = getCurrentUser_();
  const now = now_();
  const id = payload.ID || uid_();
  const existing = readObjects_(APP.SHEETS.INCOMING).find(r => String(r.ID) === String(id));

  if (existing && !canEditRecord_(existing, user)) throw new Error('Bạn không có quyền sửa văn bản này.');

  const file = payload.FILE && payload.FILE.base64 ? saveUploadedFile_(payload.FILE, 'DEN') : null;
  const status = payload.TINH_TRANG || 'Chưa xử lý';
  const row = {
    ID: id,
    STT: existing ? existing.STT : nextStt_(APP.SHEETS.INCOMING),
    SO_VAN_BAN: clean_(payload.SO_VAN_BAN),
    NGAY_VAN_BAN: payload.NGAY_VAN_BAN,
    NGAY_NHAN: payload.NGAY_NHAN || '',
    LOAI_VAN_BAN: clean_(payload.LOAI_VAN_BAN),
    CO_QUAN_BAN_HANH: clean_(payload.CO_QUAN_BAN_HANH),
    TEN_VAN_BAN: clean_(payload.TEN_VAN_BAN),
    TRICH_YEU: clean_(payload.TRICH_YEU),
    NGUOI_CHU_TRI_ID: clean_(payload.NGUOI_CHU_TRI_ID),
    NGUOI_PHU_HOP_ID: Array.isArray(payload.NGUOI_PHU_HOP_ID) ? payload.NGUOI_PHU_HOP_ID.join(',') : clean_(payload.NGUOI_PHU_HOP_ID),
    BO_PHAN: clean_(payload.BO_PHAN),
    MUC_DO: clean_(payload.MUC_DO),
    HAN_XU_LY: payload.HAN_XU_LY || '',
    TINH_TRANG: status,
    TIEN_DO: Number(payload.TIEN_DO || 0),
    TRANG_THAI_HAN: calcDueStatus_(payload.HAN_XU_LY, status),
    GHI_CHU: clean_(payload.GHI_CHU),
    FILE_ID: file ? file.id : (existing ? existing.FILE_ID : ''),
    FILE_NAME: file ? file.name : (existing ? existing.FILE_NAME : ''),
    FILE_URL: file ? file.url : (existing ? existing.FILE_URL : ''),
    NGUOI_NHAP: existing ? existing.NGUOI_NHAP : user.EMAIL,
    NGAY_NHAP: existing ? existing.NGAY_NHAP : now,
    NGAY_CAP_NHAT: now,
    ACTIVE: 'TRUE'
  };

  upsertObject_(APP.SHEETS.INCOMING, row);
  appendHistory_(id, 'VĂN BẢN ĐẾN', existing ? 'CẬP NHẬT' : 'TẠO MỚI',
    existing ? 'Cập nhật thông tin văn bản' : 'Tạo văn bản mới');
  audit_(existing ? 'UPDATE_INCOMING' : 'CREATE_INCOMING', 'VĂN BẢN ĐẾN', id, row.SO_VAN_BAN);
  return {ok:true, id:id};
}

function saveOutgoing(payload) {
  requirePermission_(['ADMIN','QUAN_LY','NHAN_VIEN']);
  payload = payload || {};
  validateRequired_(payload, ['SO_VAN_BAN','NGAY_VAN_BAN','TEN_VAN_BAN','TRICH_YEU']);

  const user = getCurrentUser_();
  const now = now_();
  const id = payload.ID || uid_();
  const existing = readObjects_(APP.SHEETS.OUTGOING).find(r => String(r.ID) === String(id));

  if (existing && !canEditRecord_(existing, user)) throw new Error('Bạn không có quyền sửa văn bản này.');

  const file = payload.FILE && payload.FILE.base64 ? saveUploadedFile_(payload.FILE, 'DI') : null;
  const row = {
    ID: id,
    STT: existing ? existing.STT : nextStt_(APP.SHEETS.OUTGOING),
    SO_VAN_BAN: clean_(payload.SO_VAN_BAN),
    NGAY_VAN_BAN: payload.NGAY_VAN_BAN,
    LOAI_VAN_BAN: clean_(payload.LOAI_VAN_BAN),
    NOI_NHAN: clean_(payload.NOI_NHAN),
    TEN_VAN_BAN: clean_(payload.TEN_VAN_BAN),
    TRICH_YEU: clean_(payload.TRICH_YEU),
    NGUOI_SOAN_ID: clean_(payload.NGUOI_SOAN_ID),
    TINH_TRANG: payload.TINH_TRANG || 'Đang xử lý',
    FILE_ID: file ? file.id : (existing ? existing.FILE_ID : ''),
    FILE_NAME: file ? file.name : (existing ? existing.FILE_NAME : ''),
    FILE_URL: file ? file.url : (existing ? existing.FILE_URL : ''),
    GHI_CHU: clean_(payload.GHI_CHU),
    NGUOI_NHAP: existing ? existing.NGUOI_NHAP : user.EMAIL,
    NGAY_NHAP: existing ? existing.NGAY_NHAP : now,
    NGAY_CAP_NHAT: now,
    ACTIVE: 'TRUE'
  };

  upsertObject_(APP.SHEETS.OUTGOING, row);
  appendHistory_(id, 'VĂN BẢN ĐI', existing ? 'CẬP NHẬT' : 'TẠO MỚI',
    existing ? 'Cập nhật thông tin văn bản' : 'Tạo văn bản mới');
  audit_(existing ? 'UPDATE_OUTGOING' : 'CREATE_OUTGOING', 'VĂN BẢN ĐI', id, row.SO_VAN_BAN);
  return {ok:true, id:id};
}

function completeRecord(type, id) {
  const user = requirePermission_(['ADMIN','QUAN_LY','NHAN_VIEN']);
  const sheet = type === 'DEN' ? APP.SHEETS.INCOMING : APP.SHEETS.OUTGOING;
  const r = readObjects_(sheet).find(x => String(x.ID) === String(id));
  if (!r) throw new Error('Không tìm thấy văn bản.');
  if (!canEditRecord_(r, user)) throw new Error('Bạn không có quyền hoàn thành văn bản này.');
  if (String(r.TINH_TRANG) === 'Đã hoàn thành') return {ok:true};
  r.TINH_TRANG = 'Đã hoàn thành';
  if (type === 'DEN') r.TIEN_DO = 100;
  r.NGAY_CAP_NHAT = now_();
  upsertObject_(sheet, r);
  appendHistory_(id, type === 'DEN' ? 'VĂN BẢN ĐẾN' : 'VĂN BẢN ĐI', 'HOÀN THÀNH', 'Đánh dấu đã hoàn thành');
  audit_('COMPLETE', type, id, r.SO_VAN_BAN);
  return {ok:true};
}

function addNote(type, id, text) {
  requirePermission_(['ADMIN','QUAN_LY','NHAN_VIEN']);
  text = String(text || '').trim();
  if (!text) throw new Error('Nhập nội dung ghi chú.');
  if (text.length > 1000) throw new Error('Ghi chú tối đa 1000 ký tự.');
  const sheet = type === 'DEN' ? APP.SHEETS.INCOMING : APP.SHEETS.OUTGOING;
  const r = readObjects_(sheet).find(x => String(x.ID) === String(id));
  if (!r) throw new Error('Không tìm thấy văn bản.');
  appendHistory_(id, type === 'DEN' ? 'VĂN BẢN ĐẾN' : 'VĂN BẢN ĐI', 'GHI CHÚ', text);
  return {ok:true};
}

function archiveRecord(type, id) {
  requirePermission_(['ADMIN','QUAN_LY']);
  const sheet = type === 'DEN' ? APP.SHEETS.INCOMING : APP.SHEETS.OUTGOING;
  const r = readObjects_(sheet).find(x => String(x.ID) === String(id));
  if (!r) throw new Error('Không tìm thấy văn bản.');
  r.ACTIVE = 'FALSE';
  r.TINH_TRANG = 'Lưu trữ';
  r.NGAY_CAP_NHAT = now_();
  upsertObject_(sheet, r);
  appendHistory_(id, type === 'DEN' ? 'VĂN BẢN ĐẾN' : 'VĂN BẢN ĐI', 'LƯU TRỮ', 'Đưa văn bản vào lưu trữ');
  audit_('ARCHIVE', type, id, r.SO_VAN_BAN);
  return {ok:true};
}

function deleteRecord(type, id) {
  requirePermission_(['ADMIN']);
  const sheet = type === 'DEN' ? APP.SHEETS.INCOMING : APP.SHEETS.OUTGOING;
  const r = readObjects_(sheet).find(x => String(x.ID) === String(id));
  if (!r) throw new Error('Không tìm thấy văn bản.');

  // Xóa file đính kèm trong Drive (chuyển vào Thùng rác của Drive, giữ được 30 ngày)
  let fileMsg = 'không có file đính kèm';
  if (r.FILE_ID) {
    try {
      DriveApp.getFileById(String(r.FILE_ID)).setTrashed(true);
      fileMsg = 'đã xóa file "' + r.FILE_NAME + '" trong Drive';
    } catch (e) {
      fileMsg = 'KHÔNG xóa được file trong Drive (' + (e && e.message ? e.message : e) + ')';
    }
  }
  const oldName = r.FILE_NAME;
  r.ACTIVE = 'DELETED';
  r.FILE_ID = ''; r.FILE_NAME = ''; r.FILE_URL = '';
  r.NGAY_CAP_NHAT = now_();
  upsertObject_(sheet, r);
  appendHistory_(id, type === 'DEN' ? 'VĂN BẢN ĐẾN' : 'VĂN BẢN ĐI', 'XÓA', 'Xóa văn bản; ' + fileMsg);
  audit_('DELETE', type, id, (r.SO_VAN_BAN || '') + ' | ' + fileMsg);
  return {ok:true, fileDeleted: !!oldName && fileMsg.indexOf('KHÔNG') < 0, message: 'Đã xóa văn bản; ' + fileMsg + '.'};
}

/* =========================
   NGƯỜI DÙNG / DANH MỤC
========================= */

function saveUser(payload) {
  const me = requirePermission_(['ADMIN']);
  payload = payload || {};
  validateRequired_(payload, ['HO_TEN','EMAIL','VAI_TRO','TEN_DANG_NHAP']);
  const username = normUsername_(payload.TEN_DANG_NHAP);
  if (!/^[a-z0-9._-]{3,30}$/.test(username)) throw new Error('Tên đăng nhập gồm 3-30 ký tự: chữ không dấu, số, dấu . _ -');
  const role = clean_(payload.VAI_TRO).toUpperCase();
  if (['ADMIN','QUAN_LY','NHAN_VIEN','CHI_XEM'].indexOf(role) < 0) throw new Error('Vai trò không hợp lệ.');

  const all = readObjects_(APP.SHEETS.USERS);
  const id = payload.ID || uid_();
  const existing = all.find(r => String(r.ID) === String(id));
  if (all.some(r => String(r.ID) !== String(id) && normUsername_(r.TEN_DANG_NHAP) === username)) {
    throw new Error('Tên đăng nhập "' + username + '" đã tồn tại.');
  }
  const status = payload.TRANG_THAI || 'Đang hoạt động';
  if (String(id) === String(me.ID) && (role !== 'ADMIN' || status !== 'Đang hoạt động')) {
    throw new Error('Không thể tự hạ quyền hoặc khóa chính tài khoản đang đăng nhập.');
  }

  const row = Object.assign({}, existing || {}, {
    ID:id, HO_TEN:clean_(payload.HO_TEN), EMAIL:clean_(payload.EMAIL).toLowerCase(),
    BO_PHAN:clean_(payload.BO_PHAN), VAI_TRO:role, TRANG_THAI:status,
    TEN_DANG_NHAP:username,
    NGAY_TAO:existing ? existing.NGAY_TAO : now_()
  });
  const pw = String(payload.MAT_KHAU || '');
  if (pw) {
    if (pw.length < 6) throw new Error('Mật khẩu tối thiểu 6 ký tự.');
    setPassword_(row, pw);
  } else if (!existing || !row.MAT_KHAU_HASH) {
    throw new Error('Tài khoản mới phải có mật khẩu (tối thiểu 6 ký tự).');
  }
  upsertObject_(APP.SHEETS.USERS, row);
  syncHandlerFromUser_(row);
  audit_('SAVE_USER','NGƯỜI DÙNG',id,username + ' / ' + role);
  return {ok:true};
}

function saveHandler(payload) {
  requirePermission_(['ADMIN']);
  payload = payload || {};
  validateRequired_(payload, ['HO_TEN','EMAIL']);
  const id = payload.ID || uid_();
  const row = {
    ID:id, HO_TEN:clean_(payload.HO_TEN), EMAIL:clean_(payload.EMAIL).toLowerCase(),
    BO_PHAN:clean_(payload.BO_PHAN), TRANG_THAI:payload.TRANG_THAI || 'Đang hoạt động',
    NGAY_TAO:payload.NGAY_TAO || now_()
  };
  upsertObject_(APP.SHEETS.HANDLERS,row);
  audit_('SAVE_HANDLER','NGƯỜI XỬ LÝ',id,row.EMAIL);
  return {ok:true};
}

function saveCategory(payload) {
  requirePermission_(['ADMIN']);
  payload = payload || {};
  validateRequired_(payload, ['NHOM','GIA_TRI']);
  const id = payload.ID || uid_();
  const row = {
    ID:id, NHOM:clean_(payload.NHOM), GIA_TRI:clean_(payload.GIA_TRI),
    TRANG_THAI:payload.TRANG_THAI || 'Đang dùng',
    THU_TU:payload.THU_TU || 99
  };
  upsertObject_(APP.SHEETS.CATEGORIES,row);
  audit_('SAVE_CATEGORY','DANH MỤC',id,row.NHOM + ': ' + row.GIA_TRI);
  return {ok:true};
}

/* =========================
   CẤU HÌNH EMAIL
========================= */

function saveConfig(payload) {
  requirePermission_(['ADMIN']);
  payload = payload || {};
  const allowed = [
    'EMAIL_ENABLED','REMIND_BEFORE_DAYS','SEND_TIMES','REMIND_DUE_TODAY',
    'OVERDUE_DAYS','MAX_EMAILS_PER_RUN'
  ];
  allowed.forEach(k => {
    if (payload[k] !== undefined) setConfig_(k, String(payload[k]));
  });
  createReminderTrigger_();
  audit_('SAVE_CONFIG','CẤU HÌNH','EMAIL','Cập nhật cấu hình nhắc việc');
  return getConfigObject_();
}

function sendTestEmail() {
  requirePermission_(['ADMIN']);
  const email = getCurrentEmail_();
  if (!email) throw new Error('Không xác định được email tài khoản.');
  MailApp.sendEmail({
    to: email,
    subject: '[Quản lý văn bản] Email kiểm tra',
    htmlBody: '<h3>Email kiểm tra thành công</h3><p>Hệ thống nhắc việc đang hoạt động.</p>'
  });
  return 'Đã gửi email kiểm tra tới ' + email;
}

function runReminderNow() {
  requirePermission_(['ADMIN']);
  return processReminders_(true);
}

/**
 * Được trigger gọi theo các giờ trong SEND_TIMES (mỗi giờ một trigger hằng ngày).
 */
function scheduledReminder() {
  processReminders_(false);
}

/**
 * Tạo lại trigger: mỗi giờ trong SEND_TIMES (vd 08:00,15:00) có một trigger chạy hằng ngày.
 * Trigger thuộc về tài khoản đang chạy hàm này (thường là ADMIN).
 */
function createReminderTrigger_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'scheduledReminder')
    .forEach(t => ScriptApp.deleteTrigger(t));

  const tz = Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh';
  const hours = parseSendHours_(getConfigValue_('SEND_TIMES'));
  hours.forEach(h => {
    ScriptApp.newTrigger('scheduledReminder')
      .timeBased()
      .inTimezone(tz)
      .atHour(h)
      .nearMinute(5)
      .everyDays(1)
      .create();
  });
  return hours;
}

function parseSendHours_(text) {
  const hours = String(text || '08:00').split(',')
    .map(s => parseInt(String(s).trim().split(':')[0], 10))
    .filter(h => !isNaN(h) && h >= 0 && h <= 23);
  const uniq = Array.from(new Set(hours));
  return uniq.length ? uniq : [8];
}

/** Số lịch nhắc đang được cài cho tài khoản hiện tại. */
function getReminderTriggerCount_() {
  try {
    return ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'scheduledReminder').length;
  } catch (e) { return 0; }
}

/** Xác định các loại nhắc áp dụng cho một văn bản (days = số ngày còn lại đến hạn). */
function getReminderTypes_(days, cfg) {
  const reminders = [];
  const beforeDays = Number(cfg.REMIND_BEFORE_DAYS || 3);
  if (days > 0 && days <= beforeDays) reminders.push('TRƯỚC HẠN');
  if (days === 0 && String(cfg.REMIND_DUE_TODAY).toUpperCase() !== 'FALSE') reminders.push('ĐẾN HẠN HÔM NAY');
  if (days < 0) {
    const od = String(cfg.OVERDUE_DAYS === undefined || cfg.OVERDUE_DAYS === '' ? 'MOI_NGAY' : cfg.OVERDUE_DAYS).trim().toUpperCase();
    if (od === 'MOI_NGAY' || od === '*') {
      reminders.push('QUÁ HẠN');
    } else {
      const list = od.split(',').map(x => Number(x.trim())).filter(x => !isNaN(x));
      if (list.indexOf(Math.abs(days)) >= 0) reminders.push('QUÁ HẠN');
    }
  }
  return reminders;
}

/** Quét toàn bộ văn bản đến và gửi nhắc việc cho người xử lý. */
function processReminders_(manual) {
  const cfg = getConfigObject_();
  if (String(cfg.EMAIL_ENABLED).toUpperCase() === 'FALSE') return 'Nhắc việc email đang tắt.';

  const tz = Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh';
  const now = new Date();
  const handlerMap = {};
  readObjects_(APP.SHEETS.HANDLERS).forEach(h => handlerMap[String(h.ID)] = h);

  const ctx = {
    cfg: cfg, tz: tz, now: now,
    slotKey: Utilities.formatDate(now, tz, 'yyyy-MM-dd HH'),
    handlerMap: handlerMap,
    logs: readObjects_(APP.SHEETS.EMAIL_LOG),
    force: !!manual,
    max: Number(cfg.MAX_EMAILS_PER_RUN || 50),
    sent: 0, skipped: 0, noEmail: 0, errors: 0
  };

  let checked = 0, due = 0;
  readObjects_(APP.SHEETS.INCOMING).filter(isActive_).forEach(r => {
    if (ctx.sent >= ctx.max) return;
    if (!r.HAN_XU_LY) return;
    if (['Đã hoàn thành','Lưu trữ'].includes(String(r.TINH_TRANG))) return;
    const dueDate = parseDate_(r.HAN_XU_LY);
    if (!dueDate) return;
    checked++;
    const days = daysBetween_(now, dueDate);
    const reminders = getReminderTypes_(days, cfg);
    if (!reminders.length) return;
    due++;
    deliverReminder_(r, days, reminders, ctx);
  });

  setConfig_('LAST_REMINDER_RUN', Utilities.formatDate(now, tz, 'yyyy-MM-dd HH:mm:ss'));
  audit_('REMINDER_RUN','EMAIL','',`Gửi ${ctx.sent}, bỏ qua ${ctx.skipped}, thiếu email ${ctx.noEmail}, lỗi ${ctx.errors}`);
  return `Đã quét ${checked} văn bản có hạn, ${due} văn bản cần nhắc. Gửi ${ctx.sent} email, bỏ qua ${ctx.skipped} (đã gửi trong giờ này), thiếu email ${ctx.noEmail}, lỗi ${ctx.errors}.`;
}

/** Gửi nhắc việc ngay cho một văn bản (nút "Nhắc ngay" trên giao diện). */
function sendReminderForDocument(id) {
  const user = requirePermission_(['ADMIN','QUAN_LY','NHAN_VIEN']);
  const r = readObjects_(APP.SHEETS.INCOMING).find(x => String(x.ID) === String(id));
  if (!r) throw new Error('Không tìm thấy văn bản.');
  if (!canEditRecord_(r, user)) throw new Error('Bạn không có quyền nhắc việc văn bản này.');
  if (!r.HAN_XU_LY) throw new Error('Văn bản chưa có hạn xử lý.');

  const cfg = getConfigObject_();
  const tz = Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh';
  const now = new Date();
  const handlerMap = {};
  readObjects_(APP.SHEETS.HANDLERS).forEach(h => handlerMap[String(h.ID)] = h);
  const ctx = {
    cfg: cfg, tz: tz, now: now,
    slotKey: Utilities.formatDate(now, tz, 'yyyy-MM-dd HH'),
    handlerMap: handlerMap, logs: [], force: true, max: 999,
    sent: 0, skipped: 0, noEmail: 0, errors: 0
  };
  const days = daysBetween_(now, parseDate_(r.HAN_XU_LY));
  deliverReminder_(r, days, ['NHẮC THỦ CÔNG'], ctx);
  appendHistory_(id, 'VĂN BẢN ĐẾN', 'NHẮC VIỆC', `Gửi ${ctx.sent} email nhắc việc`);
  if (ctx.sent === 0 && ctx.noEmail > 0) throw new Error('Người xử lý chưa có email. Hãy cập nhật email ở mục Người xử lý.');
  if (ctx.sent === 0 && ctx.errors > 0) throw new Error('Gửi email bị lỗi. Xem chi tiết trong mục Nhắc việc > Lịch sử gửi email.');
  return `Đã gửi ${ctx.sent} email nhắc việc.`;
}

/** Gửi email cho người chủ trì và người phối hợp của một văn bản, ghi log. */
function deliverReminder_(r, days, reminders, ctx) {
  const type = reminders.join(' + ');
  const main = ctx.handlerMap[String(r.NGUOI_CHU_TRI_ID)];

  const recipients = [];
  if (main) recipients.push({ email: String(main.EMAIL || '').trim().toLowerCase(), name: main.HO_TEN, role: 'Người chủ trì' });
  String(r.NGUOI_PHU_HOP_ID || '').split(',').map(x => x.trim()).filter(Boolean).forEach(hid => {
    const h = ctx.handlerMap[hid];
    if (h) recipients.push({ email: String(h.EMAIL || '').trim().toLowerCase(), name: h.HO_TEN, role: 'Người phối hợp' });
  });

  if (!recipients.length || recipients.every(x => !x.email)) {
    if (!ctx.force && alreadyLogged_(ctx.logs, r.ID, '', type, ctx.slotKey, 'Thiếu email')) return;
    logEmail_(ctx, r, '', type, days, 'Thiếu email', 'Văn bản chưa có người xử lý hoặc người xử lý chưa có email');
    ctx.noEmail++;
    return;
  }

  const seen = {};
  recipients.forEach(rc => {
    if (!rc.email || seen[rc.email]) return;
    seen[rc.email] = true;
    if (ctx.sent >= ctx.max) return;

    if (!ctx.force && alreadyLogged_(ctx.logs, r.ID, rc.email, type, ctx.slotKey, 'Đã gửi')) {
      ctx.skipped++;
      return;
    }
    try {
      MailApp.sendEmail({
        to: rc.email,
        subject: buildEmailSubject_(r, days),
        htmlBody: buildReminderHtml_(r, days, main, ctx.cfg, rc.role)
      });
      logEmail_(ctx, r, rc.email, type, days, 'Đã gửi', '');
      ctx.sent++;
    } catch (err) {
      logEmail_(ctx, r, rc.email, type, days, 'Lỗi', String(err));
      ctx.errors++;
    }
  });
}

function logEmail_(ctx, r, email, type, days, result, error) {
  const entry = {
    ID: uid_(), VAN_BAN_ID: r.ID, SO_VAN_BAN: r.SO_VAN_BAN, EMAIL: email,
    THOI_GIAN: Utilities.formatDate(new Date(), ctx.tz, 'yyyy-MM-dd HH:mm:ss'),
    LOAI_NHAC: type, SO_NGAY: days, KET_QUA: result, ERROR: error
  };
  appendObject_(APP.SHEETS.EMAIL_LOG, entry);
  ctx.logs.push(entry);
}

function alreadyLogged_(logs, id, email, type, slotKey, result) {
  return logs.some(l => String(l.VAN_BAN_ID) === String(id)
    && String(l.EMAIL).toLowerCase() === String(email).toLowerCase()
    && String(l.LOAI_NHAC) === String(type)
    && String(l.THOI_GIAN).substring(0, 13) === slotKey
    && String(l.KET_QUA) === result);
}

/* =========================
   PREVIEW / DRIVE
========================= */

/**
 * Lấy nội dung file đính kèm cho giao diện (xem trước / tải về).
 * Việc hiển thị Word, Excel, PowerPoint do trình duyệt đảm nhiệm, máy chủ chỉ trả dữ liệu file.
 */
function getFileData(fileId) {
  requireLogin_();
  if (!fileId) throw new Error('Thiếu file ID.');
  const known = readObjects_(APP.SHEETS.INCOMING).concat(readObjects_(APP.SHEETS.OUTGOING))
    .some(r => String(r.FILE_ID) === String(fileId));
  if (!known) throw new Error('File không thuộc hệ thống.');
  const file = DriveApp.getFileById(fileId);
  if (file.getSize() > APP.MAX_UPLOAD_MB * 1024 * 1024) throw new Error('File quá lớn để xem trực tiếp.');
  const blob = file.getBlob();
  const name = file.getName();
  return {
    id:file.getId(), name:name, mimeType:guessMime_(blob.getContentType(), name), size:file.getSize(),
    base64:Utilities.base64Encode(blob.getBytes()), converted:false, note:''
  };
}

function guessMime_(mime, name) {
  const m = String(mime || '').toLowerCase();
  if (m && m !== 'application/octet-stream') return m;
  const ext = String(name || '').toLowerCase().split('.').pop();
  const map = {
    pdf:'application/pdf', png:'image/png', jpg:'image/jpeg', jpeg:'image/jpeg', gif:'image/gif', webp:'image/webp', bmp:'image/bmp',
    txt:'text/plain', csv:'text/csv', html:'text/html', rtf:'application/rtf',
    doc:'application/msword', docx:'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xls:'application/vnd.ms-excel', xlsx:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    ppt:'application/vnd.ms-powerpoint', pptx:'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    odt:'application/vnd.oasis.opendocument.text', ods:'application/vnd.oasis.opendocument.spreadsheet',
    odp:'application/vnd.oasis.opendocument.presentation'
  };
  return map[ext] || m || 'application/octet-stream';
}

function saveUploadedFile_(fileObj, type) {
  const bytes = Utilities.base64Decode(String(fileObj.base64).split(',').pop());
  const sizeMb = bytes.length / 1024 / 1024;
  if (sizeMb > APP.MAX_UPLOAD_MB) {
    throw new Error('File vượt quá ' + APP.MAX_UPLOAD_MB + ' MB.');
  }

  const blob = Utilities.newBlob(bytes, fileObj.mimeType || 'application/octet-stream',
    fileObj.name || ('file_' + Date.now()));
  const folder = getDocumentFolder_(type);
  const file = folder.createFile(blob);
  return {id:file.getId(), name:file.getName(), url:file.getUrl()};
}

function getOrCreateRootFolder_() {
  const id = getConfigValue_('ROOT_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch(e) {}
  }
  const it = DriveApp.getFoldersByName(APP.ROOT_FOLDER);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(APP.ROOT_FOLDER);
}

function getDocumentFolder_(type) {
  const root = getOrCreateRootFolder_();
  const year = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh', 'yyyy');
  const mainName = type === 'DEN' ? 'VAN_BAN_DEN' : 'VAN_BAN_DI';
  const main = getOrCreateChild_(root, mainName);
  return getOrCreateChild_(main, year);
}

function getOrCreateChild_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

/* =========================
   TỔNG QUAN / DỮ LIỆU
========================= */

function buildStats_(incoming, outgoing) {
  const today = new Date();
  const before = Number(getConfigValue_('REMIND_BEFORE_DAYS') || 3);
  const closed = ['Đã hoàn thành','Lưu trữ'];
  const stats = {
    totalIncoming: incoming.length,
    totalOutgoing: outgoing.length,
    processing:0, dueSoon:0, dueToday:0, overdue:0,
    completed:0, archived:0, noDueDate:0
  };

  incoming.forEach(r => {
    const status = String(r.TINH_TRANG || '');
    if (['Đang xử lý','Chưa xử lý','Chờ phối hợp','Chờ phê duyệt'].includes(status)) stats.processing++;
    if (status === 'Đã hoàn thành') stats.completed++;
    if (status === 'Lưu trữ') stats.archived++;

    const d = parseDate_(r.HAN_XU_LY);
    if (!d) { stats.noDueDate++; return; }
    if (closed.includes(status)) return;
    const diff = daysBetween_(today, d);
    if (diff < 0) stats.overdue++;
    else if (diff === 0) stats.dueToday++;
    else if (diff <= before) stats.dueSoon++;
  });
  return stats;
}

function enrichIncoming_(rows) {
  const map = {};
  readObjects_(APP.SHEETS.HANDLERS).forEach(h => map[h.ID] = h);
  return rows.map(r => {
    const out = Object.assign({}, r);
    const main = map[r.NGUOI_CHU_TRI_ID];
    out.NGUOI_CHU_TRI_TEN = main ? main.HO_TEN : '';
    out.NGUOI_CHU_TRI_EMAIL = main ? main.EMAIL : '';
    out.NGUOI_PHU_HOP_TEN = String(r.NGUOI_PHU_HOP_ID || '').split(',')
      .map(id => map[id.trim()] ? map[id.trim()].HO_TEN : '').filter(Boolean).join(', ');
    out.NGUOI_PHU_HOP_EMAIL = String(r.NGUOI_PHU_HOP_ID || '').split(',')
      .map(id => map[id.trim()] ? map[id.trim()].EMAIL : '').filter(Boolean).join(', ');
    out.TRANG_THAI_HAN = calcDueStatus_(r.HAN_XU_LY, r.TINH_TRANG);
    return out;
  });
}

function enrichOutgoing_(rows) {
  const map = {};
  readObjects_(APP.SHEETS.HANDLERS).forEach(h => map[h.ID] = h);
  return rows.map(r => {
    const out = Object.assign({}, r);
    const p = map[r.NGUOI_SOAN_ID];
    out.NGUOI_SOAN_TEN = p ? p.HO_TEN : '';
    return out;
  });
}

function calcDueStatus_(dateValue, status) {
  if (!dateValue) return 'Không có hạn';
  if (['Đã hoàn thành','Lưu trữ'].includes(String(status))) return 'Đã đóng';
  const d = parseDate_(dateValue);
  if (!d) return 'Không có hạn';
  const diff = daysBetween_(new Date(), d);
  const before = Number(getConfigValue_('REMIND_BEFORE_DAYS') || 3);
  if (diff < 0) return 'Quá hạn';
  if (diff === 0) return 'Đến hạn hôm nay';
  if (diff <= before) return 'Sắp đến hạn';
  return 'Chưa đến hạn';
}

/* =========================
   HỖ TRỢ / PHÂN QUYỀN
========================= */

/* ---- Phiên đăng nhập: do ứng dụng tự quản lý (tên đăng nhập + mật khẩu) ---- */
var CURRENT_USER_ID_ = null;
var CURRENT_EMAIL_ = '';
const SESSION_TTL_ = 21600;   // 6 giờ (tối đa của CacheService)
const MAX_FAILS_ = 5;
const LOCK_SECONDS_ = 600;

function getCurrentEmail_() {
  if (CURRENT_EMAIL_) return CURRENT_EMAIL_;
  let e = '';
  try { e = Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail() || ''; } catch (x) {}
  return String(e).trim().toLowerCase();
}

/** Chỉ cho chạy từ trình soạn thảo Apps Script của chủ script. */
function requireOwnerContext_() {
  let active = '', eff = '';
  try { active = Session.getActiveUser().getEmail(); eff = Session.getEffectiveUser().getEmail(); } catch (e) {}
  if (!active || !eff || active.toLowerCase() !== eff.toLowerCase()) {
    throw new Error('Chức năng này chỉ chủ script được chạy trong trình soạn thảo Apps Script.');
  }
}

function normUsername_(v) { return String(v === null || v === undefined ? '' : v).trim().toLowerCase(); }
function isUserActive_(u) {
  return String(u.TRANG_THAI || '').normalize('NFC').trim().toLowerCase() === 'đang hoạt động';
}
function publicUser_(u) {
  const o = Object.assign({}, u);
  delete o.MAT_KHAU_HASH; delete o.SALT;
  o.VAI_TRO = String(o.VAI_TRO || '').toUpperCase();
  return o;
}

function hashPassword_(pw, salt) {
  let h = salt + ':' + pw;
  for (let i = 0; i < 300; i++) {
    h = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, h + salt, Utilities.Charset.UTF_8));
  }
  return h;
}
function setPassword_(row, pw) {
  row.SALT = Utilities.getUuid().replace(/-/g, '');
  row.MAT_KHAU_HASH = hashPassword_(pw, row.SALT);
}
function randomPassword_() {
  const chars = 'abcdefghjkmnpqrstuvwxyzABCDEFGHJKMNPQRSTUVWXYZ23456789';
  let p = '';
  for (let i = 0; i < 10; i++) p += chars.charAt(Math.floor(Math.random() * chars.length));
  return p;
}

/**
 * Chạy trong setupSystem: bổ sung tên đăng nhập cho người dùng cũ,
 * và cấp mật khẩu tạm cho ADMIN chưa có mật khẩu.
 */
function bootstrapLogins_(ownerEmail) {
  let users = readObjects_(APP.SHEETS.USERS);
  const used = {};
  users.forEach(u => { if (u.TEN_DANG_NHAP) used[normUsername_(u.TEN_DANG_NHAP)] = true; });
  const notes = [];

  if (!users.some(u => String(u.VAI_TRO).toUpperCase() === 'ADMIN') && ownerEmail) {
    appendObject_(APP.SHEETS.USERS, {
      ID: uid_(), HO_TEN: ownerEmail.split('@')[0], EMAIL: ownerEmail,
      BO_PHAN: 'Quản trị', VAI_TRO: 'ADMIN', TRANG_THAI: 'Đang hoạt động', NGAY_TAO: now_()
    });
    users = readObjects_(APP.SHEETS.USERS);
  }

  users.forEach(u => {
    let changed = false;
    if (!u.ID) { u.ID = uid_(); changed = true; }
    if (!u.TRANG_THAI) { u.TRANG_THAI = 'Đang hoạt động'; changed = true; }
    if (!u.TEN_DANG_NHAP) {
      let base = String(u.EMAIL || 'user').split('@')[0].toLowerCase().replace(/[^a-z0-9._-]/g, '');
      if (base.length < 3) base = 'user' + base;
      let name = base, n = 1;
      while (used[name]) { n++; name = base + n; }
      used[name] = true;
      u.TEN_DANG_NHAP = name; changed = true;
    }
    if (!u.MAT_KHAU_HASH && String(u.VAI_TRO).toUpperCase() === 'ADMIN') {
      const pw = randomPassword_();
      setPassword_(u, pw);
      notes.push('ADMIN "' + u.TEN_DANG_NHAP + '" — mật khẩu tạm: ' + pw + ' (hãy đổi sau khi đăng nhập)');
      changed = true;
    }
    if (changed) { writeUserRow_(u); }
  });

  const missing = users.filter(u => !u.MAT_KHAU_HASH && String(u.VAI_TRO).toUpperCase() !== 'ADMIN');
  if (missing.length) notes.push('Người dùng chưa có mật khẩu (ADMIN đặt trong mục Người dùng): ' + missing.map(u => u.TEN_DANG_NHAP).join(', '));
  const out = notes.join('\n');
  if (out) Logger.log(out);
  return out;
}

function writeUserRow_(u) {
  const row = Object.assign({}, u);
  row.ID = String(row.ID);
  upsertObject_(APP.SHEETS.USERS, row);
}

/** Chạy tay trong trình soạn thảo để đặt lại mật khẩu: datLaiMatKhau('tendangnhap'). */
function datLaiMatKhauAdmin() {
  requireOwnerContext_();
  const username = 'admin';   // đổi thành tên đăng nhập cần đặt lại
  const u = readObjects_(APP.SHEETS.USERS).find(x => normUsername_(x.TEN_DANG_NHAP) === username);
  if (!u) throw new Error('Không thấy tên đăng nhập "' + username + '" trong 02_NGUOIDUNG.');
  const pw = randomPassword_();
  setPassword_(u, pw);
  u.TRANG_THAI = 'Đang hoạt động';
  writeUserRow_(u);
  Logger.log('Mật khẩu mới của ' + username + ': ' + pw);
  return 'Mật khẩu mới của ' + username + ': ' + pw;
}

function login(username, password) {
  const name = normUsername_(username);
  const pw = String(password || '');
  if (!name || !pw) throw new Error('Nhập tên đăng nhập và mật khẩu.');
  const cache = CacheService.getScriptCache();
  const failKey = 'fail_' + name.replace(/[^a-z0-9._@-]/g, '').substring(0, 60);
  const fails = Number(cache.get(failKey) || 0);
  if (fails >= MAX_FAILS_) throw new Error('Đăng nhập sai quá nhiều lần. Thử lại sau 10 phút.');

  const u = readObjects_(APP.SHEETS.USERS).find(x =>
    normUsername_(x.TEN_DANG_NHAP) === name || (name.indexOf('@') > 0 && normUsername_(x.EMAIL) === name));
  const ok = u && u.MAT_KHAU_HASH && u.SALT && hashPassword_(pw, String(u.SALT)) === String(u.MAT_KHAU_HASH);
  if (!ok) {
    cache.put(failKey, String(fails + 1), LOCK_SECONDS_);
    throw new Error('Sai tên đăng nhập hoặc mật khẩu.');
  }
  if (!isUserActive_(u)) throw new Error('Tài khoản đã bị khóa. Liên hệ quản trị viên.');

  cache.remove(failKey);
  const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  cache.put('sess_' + token, String(u.ID), SESSION_TTL_);
  CURRENT_USER_ID_ = String(u.ID);
  CURRENT_EMAIL_ = String(u.EMAIL || '').toLowerCase();
  audit_('LOGIN','TÀI KHOẢN',u.ID,u.TEN_DANG_NHAP);
  return { token: token };
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('sess_' + token);
  return {ok:true};
}

function changePassword(oldPw, newPw) {
  const me = requireLogin_();
  newPw = String(newPw || '');
  if (newPw.length < 6) throw new Error('Mật khẩu mới tối thiểu 6 ký tự.');
  const row = readObjects_(APP.SHEETS.USERS).find(x => String(x.ID) === String(me.ID));
  if (!row || hashPassword_(String(oldPw || ''), String(row.SALT)) !== String(row.MAT_KHAU_HASH)) {
    throw new Error('Mật khẩu hiện tại không đúng.');
  }
  setPassword_(row, newPw);
  writeUserRow_(row);
  audit_('CHANGE_PASSWORD','TÀI KHOẢN',row.ID,row.TEN_DANG_NHAP);
  return 'Đã đổi mật khẩu.';
}

/**
 * Cổng gọi duy nhất từ giao diện: kiểm tra phiên đăng nhập rồi mới chạy hàm.
 * Các hàm nghiệp vụ không có phiên sẽ tự báo PHIEN_HET_HAN.
 */
function api(token, fn, args) {
  const allowed = {
    getAppData:getAppData, saveIncoming:saveIncoming, saveOutgoing:saveOutgoing,
    archiveRecord:archiveRecord, deleteRecord:deleteRecord, completeRecord:completeRecord, addNote:addNote, saveUser:saveUser,
    saveHandler:saveHandler, saveCategory:saveCategory, saveConfig:saveConfig,
    sendTestEmail:sendTestEmail, runReminderNow:runReminderNow,
    sendReminderForDocument:sendReminderForDocument, getFileData:getFileData,
    changePassword:changePassword
  };
  if (!allowed[fn]) throw new Error('Chức năng không hợp lệ.');
  const id = token ? CacheService.getScriptCache().get('sess_' + token) : null;
  if (!id) throw new Error('PHIEN_HET_HAN');
  const u = readObjects_(APP.SHEETS.USERS).find(x => String(x.ID) === String(id));
  if (!u || !isUserActive_(u)) throw new Error('PHIEN_HET_HAN');
  CacheService.getScriptCache().put('sess_' + token, String(id), SESSION_TTL_);  // gia hạn phiên
  CURRENT_USER_ID_ = String(id);
  CURRENT_EMAIL_ = String(u.EMAIL || '').toLowerCase();
  return allowed[fn].apply(null, Array.isArray(args) ? args : []);
}

function getCurrentUser_() {
  if (!CURRENT_USER_ID_) throw new Error('PHIEN_HET_HAN');
  const u = readObjects_(APP.SHEETS.USERS).find(x => String(x.ID) === String(CURRENT_USER_ID_));
  if (!u || !isUserActive_(u)) throw new Error('PHIEN_HET_HAN');
  return publicUser_(u);
}

function requireLogin_() { return getCurrentUser_(); }

function isAdmin_(u) { return String(u.VAI_TRO).toUpperCase() === 'ADMIN'; }

function requirePermission_(roles) {
  const u = requireLogin_();
  const role = String(u.VAI_TRO).toUpperCase();
  if (!roles.map(x => String(x).toUpperCase()).includes(role)) {
    throw new Error('Bạn không có quyền thực hiện thao tác này.');
  }
  return u;
}

function canEditRecord_(r, user) {
  const role = String(user.VAI_TRO).toUpperCase();
  if (role === 'ADMIN' || role === 'QUAN_LY') return true;
  if (role === 'NHAN_VIEN') {
    return String(r.NGUOI_NHAP).toLowerCase() === String(user.EMAIL).toLowerCase()
      || String(r.NGUOI_CHU_TRI_ID) === findHandlerIdByEmail_(user.EMAIL)
      || String(r.NGUOI_SOAN_ID) === findHandlerIdByEmail_(user.EMAIL);
  }
  return false;
}

function findHandlerIdByEmail_(email) {
  const h = readObjects_(APP.SHEETS.HANDLERS).find(x => String(x.EMAIL).toLowerCase() === String(email).toLowerCase());
  return h ? String(h.ID) : '';
}

function syncHandlerFromUser_(u) {
  const handlers = readObjects_(APP.SHEETS.HANDLERS);
  const existing = handlers.find(h => String(h.EMAIL).toLowerCase() === String(u.EMAIL).toLowerCase());
  const row = {
    ID: existing ? existing.ID : uid_(),
    HO_TEN:u.HO_TEN, EMAIL:u.EMAIL, BO_PHAN:u.BO_PHAN,
    TRANG_THAI: String(u.TRANG_THAI) === 'Đang hoạt động' ? 'Đang hoạt động' : 'Ngừng sử dụng',
    NGAY_TAO:existing ? existing.NGAY_TAO : now_()
  };
  upsertObject_(APP.SHEETS.HANDLERS,row);
}

function getCategories_() {
  const rows = readObjects_(APP.SHEETS.CATEGORIES)
    .filter(r => String(r.TRANG_THAI) !== 'Ngừng dùng')
    .sort((a, b) => Number(a.THU_TU || 99) - Number(b.THU_TU || 99));
  const out = {LOAI_VAN_BAN:[],TINH_TRANG:[],MUC_DO:[]};
  rows.forEach(r => { if (out[r.NHOM]) out[r.NHOM].push(r.GIA_TRI); });
  return out;
}

/* =========================
   SHEET CRUD
========================= */

function readObjects_(sheetName) {
  const sh = getSS_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(String);
  return values.slice(1)
    .filter(row => row.some(c => c !== '' && c !== null))
    .map(row => {
      const o = {};
      headers.forEach((h,i) => o[h] = normalizeValue_(row[i]));
      return o;
    });
}

function appendObject_(sheetName, obj) {
  const sh = getSS_().getSheetByName(sheetName);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  sh.appendRow(headers.map(h => obj[h] !== undefined ? obj[h] : ''));
}

function upsertObject_(sheetName, obj) {
  const sh = getSS_().getSheetByName(sheetName);
  const headers = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const idIndex = headers.indexOf('ID');
  if (idIndex < 0) throw new Error('Sheet thiếu cột ID.');
  const values = sh.getDataRange().getValues();
  let rowNumber = -1;
  for (let i=1;i<values.length;i++) {
    if (String(values[i][idIndex]) === String(obj.ID)) { rowNumber = i+1; break; }
  }
  const row = headers.map(h => obj[h] !== undefined ? obj[h] : '');
  if (rowNumber > 0) sh.getRange(rowNumber,1,1,headers.length).setValues([row]);
  else sh.appendRow(row);
}

function nextStt_(sheetName) {
  const rows = readObjects_(sheetName);
  return rows.reduce((m,r) => Math.max(m, Number(r.STT)||0), 0) + 1;
}

/**
 * ACTIVE có thể là chuỗi 'TRUE' hoặc giá trị logic true (Google Sheets tự đổi kiểu).
 * Ô trống được coi là đang hoạt động.
 */
function isActive_(r) {
  const v = (r.ACTIVE === '' || r.ACTIVE === null || r.ACTIVE === undefined) ? 'TRUE' : r.ACTIVE;
  return String(v).toUpperCase() === 'TRUE';
}

/* =========================
   HISTORY / AUDIT
========================= */

function appendHistory_(id,type,action,content) {
  appendObject_(APP.SHEETS.HISTORY,{
    ID:uid_(), VAN_BAN_ID:id, LOAI_VAN_BAN:type, THOI_GIAN:now_(),
    NGUOI_THUC_HIEN:getCurrentEmail_(), HANH_DONG:action, NOI_DUNG:content
  });
}

function audit_(action,type,id,detail) {
  try {
    appendObject_(APP.SHEETS.AUDIT,{
      ID:uid_(), THOI_GIAN:now_(), EMAIL:getCurrentEmail_(),
      HANH_DONG:action, DOI_TUONG:type, DOI_TUONG_ID:id, CHI_TIET:detail
    });
  } catch(e) {}
}

/* =========================
   CONFIG / EMAIL
========================= */

function getConfigObject_() {
  const o = {};
  readObjects_(APP.SHEETS.CONFIG).forEach(r => o[r.KEY] = r.VALUE);
  return o;
}

function getConfigValue_(key) {
  const r = readObjects_(APP.SHEETS.CONFIG).find(x => String(x.KEY) === String(key));
  return r ? r.VALUE : '';
}

function setConfig_(key,value) {
  const sh = getSS_().getSheetByName(APP.SHEETS.CONFIG);
  const values = sh.getDataRange().getValues();
  for (let i=1;i<values.length;i++) {
    if (String(values[i][0]) === String(key)) {
      sh.getRange(i+1,2).setNumberFormat('@').setValue(String(value));
      return;
    }
  }
  sh.appendRow([key,String(value),'']);
}

function buildEmailSubject_(r,days) {
  if (days < 0) return '[QUẢN LÝ VĂN BẢN] QUÁ HẠN: ' + r.SO_VAN_BAN;
  if (days === 0) return '[QUẢN LÝ VĂN BẢN] ĐẾN HẠN HÔM NAY: ' + r.SO_VAN_BAN;
  return '[QUẢN LÝ VĂN BẢN] CÒN ' + days + ' NGÀY: ' + r.SO_VAN_BAN;
}

function buildReminderHtml_(r,days,main,cfg,role) {
  const title = days < 0 ? 'CẢNH BÁO VĂN BẢN QUÁ HẠN'
    : days === 0 ? 'VĂN BẢN ĐẾN HẠN HÔM NAY'
    : 'NHẮC VIỆC XỬ LÝ VĂN BẢN';
  const color = days < 0 ? '#dc2626' : days === 0 ? '#ea580c' : '#2563eb';
  const due = r.HAN_XU_LY ? formatDisplayDate_(r.HAN_XU_LY) : '';
  return `
    <div style="font-family:Arial,sans-serif;max-width:680px">
      <h2 style="color:${color}">${title}</h2>
      <table cellpadding="7" style="border-collapse:collapse">
        <tr><td><b>Số văn bản</b></td><td>${esc_(r.SO_VAN_BAN)}</td></tr>
        <tr><td><b>Ngày văn bản</b></td><td>${esc_(formatDisplayDate_(r.NGAY_VAN_BAN))}</td></tr>
        <tr><td><b>Trích yếu</b></td><td>${esc_(r.TRICH_YEU)}</td></tr>
        <tr><td><b>Người chủ trì</b></td><td>${esc_(main ? main.HO_TEN : '')} ${main && main.EMAIL ? '(' + esc_(main.EMAIL) + ')' : ''}</td></tr>
        <tr><td><b>Vai trò của bạn</b></td><td>${esc_(role || '')}</td></tr>
        <tr><td><b>Hạn xử lý</b></td><td>${esc_(due)}</td></tr>
        <tr><td><b>Trạng thái</b></td><td>${esc_(calcDueStatus_(r.HAN_XU_LY,r.TINH_TRANG))}</td></tr>
      </table>
      <p>Vui lòng kiểm tra và thực hiện xử lý văn bản theo quy định.</p>
    </div>`;
}

/* =========================
   TIỆN ÍCH
========================= */

function uid_() { return Utilities.getUuid(); }

function now_() { return new Date(); }

function clean_(v) { return v === null || v === undefined ? '' : String(v).trim(); }

/** Chuyển Date thành chuỗi để google.script.run trả về được (Date không serialize được). */
function normalizeValue_(v) {
  if (v instanceof Date) {
    if (isNaN(v.getTime())) return '';
    const tz = Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh';
    const hasTime = v.getHours() || v.getMinutes() || v.getSeconds();
    return Utilities.formatDate(v, tz, hasTime ? 'yyyy-MM-dd HH:mm:ss' : 'yyyy-MM-dd');
  }
  return v;
}

function validateRequired_(obj, fields) {
  fields.forEach(f => {
    if (obj[f] === undefined || obj[f] === null || String(obj[f]).trim() === '') {
      throw new Error('Thiếu trường bắt buộc: ' + f);
    }
  });
}

function parseDate_(v) {
  if (!v) return null;
  if (v instanceof Date && !isNaN(v.getTime())) return v;
  const s = String(v).trim();
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function stripTime_(d) {
  const x = new Date(d);
  x.setHours(0,0,0,0);
  return x;
}

/** Số ngày từ a đến b (b - a), theo ngày lịch. */
function daysBetween_(a, b) {
  return Math.round((stripTime_(b) - stripTime_(a)) / 86400000);
}

function formatDateKey_(v) {
  if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}/.test(v)) return v.substring(0,10);
  const d = parseDate_(v) || new Date(v);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
}

function formatDisplayDate_(v) {
  const d = parseDate_(v);
  return d ? Utilities.formatDate(d, Session.getScriptTimeZone() || 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy') : '';
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
}
