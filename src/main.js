// --- CẤU HÌNH API ---
// Thay thế URL này bằng URL của Web App sau khi bạn Deploy (Triển khai) trong Google Apps Script
const API_URL = 'https://script.google.com/macros/s/AKfycbyPf9oYjWdG1VROFGTNpXdV1zw86BSCnPnsbdP9yImHeonUDbL1PxY1uAqSR5jsTBoUxw/exec';

// Supabase Configuration
const SUPABASE_URL = 'https://weipegqglhqsqvdzgztb.supabase.co';
const SUPABASE_KEY = 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6IndlaXBlZ3FnbGhxc3F2ZHpnenRiIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzAzODcyNzAsImV4cCI6MjA4NTk2MzI3MH0.v9ylIEOGNrbTyDpjbWst4KoXtKL1cJ58A-LEkXKCd5Q';
const supabaseClient = supabase.createClient(SUPABASE_URL, SUPABASE_KEY);

// Network Monitoring
window.isOnline = navigator.onLine;
window.addEventListener('online', updateNetworkStatus);
window.addEventListener('offline', updateNetworkStatus);

function handleError(err) {
    console.error('Lỗi ứng dụng:', err);
    showLoading(false);
    Swal.fire({
        icon: 'error',
        title: 'Đã xảy ra lỗi',
        text: err.message || 'Không thể kết nối với máy chủ.',
        confirmButtonText: 'Đóng'
    });
}

function updateNetworkStatus() {
    window.isOnline = navigator.onLine;
    const indicator = document.getElementById('network-status');
    const text = document.getElementById('status-text');
    if (indicator) {
        indicator.classList.remove('d-none');
        if (window.isOnline) {
            indicator.className = 'position-fixed top-0 end-0 m-2 badge bg-success';
            text.innerText = 'Đang trực tuyến';
            setTimeout(() => indicator.classList.add('d-none'), 3000);
            processOfflineQueue(); // Try to sync when back online
        } else {
            indicator.className = 'position-fixed top-0 end-0 m-2 badge bg-danger';
            text.innerText = 'Đang ngoại tuyến (Offline)';
        }
    }
}

// --- BIẾN TOÀN CỤC ---
var GLOBAL_DATA = {};
var CURRENT_USER = null;
var CURRENT_SHEET = '';
var sidebarBS = null;
var FILTER_TIMER = null;
var LOADING_TIMEOUT = null;
var EDIT_INDEX = -1;
var CURRENT_MENU_ID = '';

// --- CẤU HÌNH MENU ---
const MENU_STRUCTURE = [
    {
        group: 'QUẢN LÝ KHO & TÀI CHÍNH', items: [
            { id: 'chiphi', sheet: 'Chiphi', icon: 'fas fa-coins', title: 'Quản lý Chi phí' },
            { id: 'phieunhap', sheet: 'Phieunhap', icon: 'fas fa-arrow-circle-down', title: 'Nhập kho' },
            { id: 'phieuxuat', sheet: 'Phieuxuat', icon: 'fas fa-arrow-circle-up', title: 'Xuất kho' },
            { id: 'phieuchuyen', sheet: 'Phieuchuyenkho', icon: 'fas fa-exchange-alt', title: 'Điều chuyển' },
            { id: 'tonkho', sheet: 'Tonkho', icon: 'fas fa-boxes', title: 'Báo cáo Tồn kho' }
        ]
    },
    {
        group: 'QUẢN TRỊ NHÂN SỰ', items: [
            { id: 'chamcong', sheet: 'Chamcong', icon: 'fas fa-calendar-check', title: 'Bảng Chấm công' },
            { id: 'tamung', sheet: 'GiaoDichLuong', icon: 'fas fa-hand-holding-usd', title: 'Tạm ứng / PC' },
            { id: 'bangluong', sheet: 'BangLuongThang', icon: 'fas fa-money-check-alt', title: 'Bảng Lương tháng' },
            { id: 'nhansu', sheet: 'User', icon: 'fas fa-users', title: 'Hồ sơ Nhân sự' }
        ]
    },
    {
        group: 'DANH MỤC HỆ THỐNG', items: [
            { id: 'dskho', sheet: 'DS_kho', icon: 'fas fa-warehouse', title: 'Danh sách Kho' },
            { id: 'vattu', sheet: 'VatLieu', icon: 'fas fa-tools', title: 'Danh mục Vật tư' },
            { id: 'doitac', sheet: 'Doitac', icon: 'fas fa-handshake', title: 'Đối tác / NCC' }
        ]
    }
];

const COLUMN_MAP = {
    // Thông tin chung
    'ID': 'Mã hệ thống',
    'Ngày': 'Ngày lập',
    'Ngay': 'Ngày lập',
    'Nội dung': 'Nội dung',
    'Noi_Dung': 'Nội dung',
    'Diễn giải': 'Diễn giải',
    'Ghi chú': 'Ghi chú',
    'Ghi_Chu': 'Ghi chú',
    'Trạng thái': 'Trạng thái',
    'Trang_Thai': 'Trạng thái',
    'Người lập': 'Người lập',
    'NguoiLap': 'Người lập',
    'Họ và tên': 'Họ và tên',
    'SĐT': 'Số điện thoại',
    'Email': 'Email',
    'Địa chỉ': 'Địa chỉ',
    'Chức vụ': 'Chức vụ',
    'Bộ phận': 'Bộ phận',
    'Username': 'Tên đăng nhập',
    'Password': 'Mật khẩu',
    'Role': 'Quyền hạn',
    'Mã số thuế': 'Mã số thuế',

    // Quản lý Kho (Nhập/Xuất/Tồn)
    'ID phieu nhap': 'Mã Phiếu Nhập',
    'ID phieu Xuat': 'Mã Phiếu Xuất',
    'ID Chkho': 'Mã Chuyển Kho',
    'Tên kho': 'Tên Kho',
    'Ten_Kho': 'Tên Kho',
    'Tên Kho': 'Tên Kho',
    'Loại': 'Loại phiếu',
    'Tổng tiền': 'Tổng tiền',
    'Tong_Tien': 'Tổng tiền',
    'Đường dẫn file': 'Tài liệu đính kèm',
    'Phê duyệt': 'Trưởng bộ phận',
    'Người nhập': 'Người nhập kho',
    'Người xuất': 'Người xuất kho',
    'Người Xuất': 'Người xuất kho',
    'Lý do': 'Lý do xuất/nhập',

    // Chi tiết (Sub-tables)
    'ID nhap chi tiet': 'Mã chi tiết',
    'ID xuat chi tiet': 'Mã chi tiết',
    'Mã vật tư': 'Mã vật tư',
    'Ma_VT': 'Mã vật tư',
    'Tên vật tư': 'Tên vật tư',
    'Ten_VT': 'Tên vật tư',
    'Đơn vị tính': 'Đơn vị tính',
    'ĐVT': 'Đơn vị tính',
    'Số lượng': 'Số lượng',
    'So_Luong': 'Số lượng',
    'Đơn giá': 'Đơn giá',
    'Don_Gia': 'Đơn giá',
    'Thành tiền': 'Thành tiền',
    'Thanh_Tien': 'Thành tiền',
    'Ghi chú chi tiết': 'Ghi chú chi tiết',

    // Vật tư
    'ID vật tư': 'Mã vật tư',
    'Nhóm vật tư': 'Nhóm vật tư',
    'Quy cách': 'Quy cách/Kích thước',
    'Tồn kho': 'Tồn kho hiện tại',
    'Giá nhập': 'Giá nhập gần nhất',

    // Chi phí
    'ID_ChiPhi': 'Mã Chi Phí',
    'Tên chi phí': 'Tên chi phí',
    'Số tiền': 'Số tiền chi',
    'Hạng mục': 'Hạng mục chi',
    'Who delete': 'Người xóa',

    // Nhân sự / User
    'Phân quyền': 'Phân quyền',
    'App_pass': 'Mật khẩu ứng dụng',
    'Ngày sinh': 'Ngày sinh',
    'CCCD': 'Căn cước công dân',
    'Ngày vào làm': 'Ngày vào làm',
    'Lương cơ bản': 'Lương cơ bản',
    'Phụ cấp': 'Phụ cấp',
    'Hình ảnh': 'Hình ảnh',

    // Chấm công
    'ID_ChamCong': 'Mã chấm công',
    'Tháng': 'Tháng',
    'Năm': 'Năm',
    'Ngày công': 'Ngày công',
    'Số công': 'Số công',
    'Ngày nghỉ': 'Ngày nghỉ',
    'Ghi chú chấm công': 'Ghi chú chấm công',

    // Giao dịch lương / Tạm ứng
    'ID_GiaoDich': 'Mã giao dịch',
    'Loại giao dịch': 'Loại giao dịch',
    'Số tiền': 'Số tiền',
    'Ngày giao dịch': 'Ngày giao dịch',
    'Mô tả': 'Mô tả',
    'Nhân viên': 'Nhân viên',

    // Bảng lương tháng
    'ID_BangLuong': 'Mã bảng lương',
    'Tổng lương': 'Tổng lương',
    'Thực lĩnh': 'Thực lĩnh',
    'Khấu trừ': 'Khấu trừ',
    'Thưởng': 'Thưởng',

    // Đối tác / NCC
    'ID đối tác': 'Mã đối tác',
    'Tên đối tác': 'Tên đối tác',
    'Loại đối tác': 'Loại đối tác',
    'Người liên hệ': 'Người liên hệ',
    'Website': 'Website',

    // Kho
    'ID kho': 'Mã kho',
    'Địa điểm': 'Địa điểm',
    'Quản lý kho': 'Quản lý kho',

    // Tồn kho
    'ID_TonKho': 'Mã tồn kho',
    'Số lượng tồn': 'Số lượng tồn',

    // Chuyển kho chi tiết
    'ID CKchitiet': 'Mã chi tiết chuyển',
    'Kho nguồn': 'Kho nguồn',
    'Kho đích': 'Kho đích',

    // Nhóm vật tư
    'ID NVT': 'Mã nhóm vật tư',
    'Tên nhóm': 'Tên nhóm',

    // Chi phí chi tiết
    'ID_ChiTiet': 'Mã chi tiết chi phí'
};


const RELATION_CONFIG = {
    'Phieunhap': { child: 'NhapChiTiet', foreignKey: 'ID phieu nhap', title: 'CHI TIẾT NHẬP' },
    'Phieuxuat': { child: 'XuatChiTiet', foreignKey: 'ID phieu Xuat', title: 'CHI TIẾT XUẤT' },
    'Phieuchuyenkho': { child: 'Chuyenkhochitiet', foreignKey: 'ID Chkho', title: 'CHI TIẾT CHUYỂN' },
    'Chiphi': { child: 'Chiphichitiet', foreignKey: 'ID_ChiPhi', title: 'CHI TIẾT CHI' }
};

// --- HÀM GỌI API SUPABASE ---
async function callSupabase(action, table, data = null, id = null) {
    if (!window.isOnline && action === 'read') {
        console.log(`[Offline] Đang lấy dữ liệu từ CACHE cho ${table}`);
        return { status: 'success', data: JSON.parse(localStorage.getItem(`cache_${table}`) || '[]') };
    }

    if (!window.isOnline && (action === 'insert' || action === 'update' || action === 'delete')) {
        addToOfflineQueue(action, table, data, id);
        return { status: 'offline-queued', message: 'Bạn đang ngoại tuyến. Dữ liệu đã được lưu tạm và sẽ tự động đồng bộ khi có mạng.' };
    }

    try {
        let result;
        const mappedTable = TABLE_MAP[table] || table;

        if (action === 'read') {
            result = await supabaseClient.from(mappedTable).select('*');
            if (result.error) throw result.error;
            localStorage.setItem(`cache_${table}`, JSON.stringify(result.data));
            return { status: 'success', data: result.data };
        } else if (action === 'insert') {
            result = await supabaseClient.from(mappedTable).insert([data]);
            if (result.error) throw result.error;
            return { status: 'success' };
        } else if (action === 'update') {
            const pk = PK_MAP[mappedTable] || 'ID';
            result = await supabaseClient.from(mappedTable).update(data).eq(pk, id);
            if (result.error) throw result.error;
            return { status: 'success' };
        } else if (action === 'delete') {
            const pk = PK_MAP[mappedTable] || 'ID';
            result = await supabaseClient.from(mappedTable).update({ 'Delete': 'X' }).eq(pk, id);
            if (result.error) throw result.error;
            return { status: 'success' };
        }
    } catch (err) {
        console.error('Supabase Error:', err);
        return { status: 'error', message: err.message };
    }
}

// Map các tên sheet sang tên bảng Supabase (nếu khác nhau)
const TABLE_MAP = {
    "Chamcong": "Chamcong", "LichSuLuong": "LichSuLuong", "GiaoDichLuong": "GiaoDichLuong",
    "BanGiaoCongCu": "BanGiaoCongCu", "Phieunhap": "PhieuNhap", "Phieuxuat": "PhieuXuat",
    "XuatChiTiet": "XuatChiTiet", "Phieuchuyenkho": "Phieuchuyenkho", "NhapChiTiet": "NhapChiTiet",
    "DS_kho": "DS_kho", "VatLieu": "Vat_tu", "Tonkho": "Tonkho",
    "Chuyenkhochitiet": "Chuyenkhochitiet", "Nhomvattu": "Nhomvattu", "BangLuongThang": "BangLuongThang",
    "Notes": "Notes", "LichNhac": "LichNhac", "Chiphi": "Chiphi", "User": "User",
    "Chiphichitiet": "Chiphichitiet", "Luongphucap": "Luongphucap", "Doitac": "Doitac"
};

const PK_MAP = {
    "Chamcong": "ID_ChamCong", "LichSuLuong": "ID_LichSu", "GiaoDichLuong": "ID_GiaoDich",
    "BanGiaoCongCu": "ID_BGCC", "PhieuNhap": "ID phieu nhap", "PhieuXuat": "ID phieu Xuat",
    "XuatChiTiet": "ID xuat chi tiet", "Phieuchuyenkho": "ID Chkho", "NhapChiTiet": "ID nhap chi tiet",
    "DS_kho": "ID kho", "Vat_tu": "ID vật tư", "Tonkho": "ID_TonKho",
    "Chuyenkhochitiet": "ID CKchitiet", "Nhomvattu": "ID NVT", "BangLuongThang": "ID_BangLuong",
    "Notes": "ID_Note", "LichNhac": "ID_Nhac", "Chiphi": "ID_ChiPhi", "User": "ID",
    "Chiphichitiet": "ID_ChiTiet", "Luongphucap": "ID lương", "Doitac": "ID đối tác"
};

// --- HÀM TƯƠNG THÍCH CODE CŨ ---
async function callGAS(action, sheet, data = null, id = null) {
    if (action === 'loginUser') {
        // ID là username, data là password. Trong Supabase bảng User dùng cột 'App_pass'
        const { data: user, error } = await supabaseClient.from('User').select('*').eq('ID', sheet).eq('App_pass', data).single();
        if (error) return { status: 'error', error: 'Sai ID hoặc mật khẩu' };

        // Chuẩn hóa dữ liệu user cho frontend (Map các cột tiếng Việt sang tiếng Anh cho dễ dùng)
        const normalizedUser = {
            ...user,
            name: user['Họ và tên'] || user['ID'],
            role: user['Phân quyền'] || user['Chức vụ'] || 'Nhân viên'
        };
        return { status: 'success', user: normalizedUser };
    }
    return callSupabase(action, sheet, data, id);
}

// --- KHỞI TẠO ---
document.addEventListener('DOMContentLoaded', function () {
    const cachedData = localStorage.getItem('GLOBAL_DATA_CACHE');
    if (cachedData) {
        try {
            GLOBAL_DATA = JSON.parse(cachedData);
            console.log('Loaded data from local cache');
        } catch (e) { console.error('Cache corruption', e); }
    }

    // Tự động đăng nhập nếu đã có session
    const cachedUser = localStorage.getItem('CURRENT_USER_SESSION');
    if (cachedUser) {
        try {
            CURRENT_USER = JSON.parse(cachedUser);
            document.getElementById('user-display-name').innerText = CURRENT_USER.name;
            document.getElementById('login-container').classList.add('d-none');
            document.getElementById('app-container').classList.remove('d-none');
            renderDashboard();
            getInitialData(); // Cập nhật data mới nhất
        } catch (e) { console.error('Session corruption', e); }
    }

    showLoading(false);
    renderSidebar();

    var sbEl = document.getElementById('sidebarMenu');
    if (sbEl && typeof bootstrap !== 'undefined') sidebarBS = new bootstrap.Offcanvas(sbEl);

    document.getElementById('manual-sidebar-btn').addEventListener('click', function () {
        if (window.innerWidth >= 992) document.body.classList.toggle('sidebar-closed');
        else if (sidebarBS) sidebarBS.toggle();
    });

    document.getElementById('form-login').addEventListener('submit', function (e) {
        e.preventDefault();
        var u = document.getElementById('username').value.trim();
        var p = document.getElementById('password').value.trim();
        if (!u || !p) return;
        showLoading(true);

        callGAS('loginUser', u, p)
            .then(handleLogin)
            .catch(handleError);
    });
});

// --- XỬ LÝ ĐĂNG NHẬP & TẢI FULL DỮ LIỆU ---
function handleLogin(res) {
    if (!res || res.status === 'error') {
        showLoading(false);
        Swal.fire('Lỗi', res ? res.error : 'Lỗi đăng nhập', 'error');
        return;
    }
    CURRENT_USER = res.user;
    localStorage.setItem('CURRENT_USER_SESSION', JSON.stringify(CURRENT_USER));
    document.getElementById('user-display-name').innerText = CURRENT_USER.name;

    if (Object.keys(GLOBAL_DATA).length > 0) {
        document.getElementById('login-container').classList.add('d-none');
        document.getElementById('app-container').classList.remove('d-none');
        renderDashboard();
    }
    getInitialData();
}

function refreshAllData() {
    return getInitialData();
}

function refreshSingleSheet(sheet) {
    showLoading(true);
    return callSupabase('read', sheet).then(res => {
        showLoading(false);
        if (res.status == 'success') {
            GLOBAL_DATA[sheet] = res.data;
            if (CURRENT_SHEET == sheet) renderTable(res.data);
        }
    });
}

function getInitialData() {
    showLoading(true);
    var promises = [];
    MENU_STRUCTURE.forEach(group => {
        group.items.forEach(m => {
            promises.push(callSupabase('read', m.sheet).then(res => {
                if (res.status == 'success') GLOBAL_DATA[m.sheet] = res.data;
            }));
        });
    });

    const extraSheets = ['User', 'DS_kho', 'NhapChiTiet', 'XuatChiTiet', 'Chiphichitiet', 'VatLieu', 'Nhomvattu'];
    extraSheets.forEach(s => {
        promises.push(callSupabase('read', s).then(res => {
            if (res.status == 'success') GLOBAL_DATA[s] = res.data;
        }));
    });

    return Promise.all(promises).then(() => {
        showLoading(false);
        if (document.getElementById('app-container').classList.contains('d-none')) {
            document.getElementById('login-container').classList.add('d-none');
            document.getElementById('app-container').classList.remove('d-none');
            renderDashboard();
        } else if (CURRENT_SHEET) {
            renderTable(GLOBAL_DATA[CURRENT_SHEET] || []);
        }
    }).catch(handleError);
}

// --- HÀM ĐĂNG XUẤT ---
window.doLogout = function () {
    showLoading(true);
    setTimeout(function () {
        CURRENT_USER = null;
        localStorage.removeItem('CURRENT_USER_SESSION');
        GLOBAL_DATA = {};
        document.getElementById('form-login').reset();
        document.getElementById('app-container').classList.add('d-none');
        document.getElementById('login-container').classList.remove('d-none');
        if (window.innerWidth < 992 && sidebarBS) sidebarBS.hide();
        var modalEl = document.getElementById('dataModal');
        var modalInstance = bootstrap.Modal.getInstance(modalEl);
        if (modalInstance) modalInstance.hide();
        showLoading(false);
        Swal.fire({
            icon: 'success',
            title: 'Hẹn gặp lại!',
            text: 'Bạn đã đăng xuất thành công.',
            toast: true,
            position: 'top-end',
            timer: 3000
        });
    }, 500);
}

// Quản lý Offline Queue
function addToOfflineQueue(action, table, data, id) {
    const queue = JSON.parse(localStorage.getItem('offline_queue') || '[]');
    queue.push({ action, table, data, id, timestamp: new Date().getTime() });
    localStorage.setItem('offline_queue', JSON.stringify(queue));

    Swal.fire({
        icon: 'warning',
        title: 'Đang ngoại tuyến',
        text: 'Dữ liệu đã được lưu vào hàng chờ và sẽ tự đồng bộ khi có mạng lại.',
        toast: true,
        position: 'top-end',
        timer: 3000
    });
}

async function processOfflineQueue() {
    const queue = JSON.parse(localStorage.getItem('offline_queue') || '[]');
    if (queue.length === 0) return;

    console.log(`Đang đồng bộ ${queue.length} lệnh offline...`);
    for (const item of queue) {
        await callSupabase(item.action, item.table, item.data, item.id);
    }
    localStorage.removeItem('offline_queue');

    Swal.fire({
        icon: 'success',
        title: 'Đã đồng bộ',
        text: 'Toàn bộ dữ liệu offline đã được đẩy lên hệ thống.',
        toast: true,
        position: 'top-end',
        timer: 3000
    });
}


// --- RENDER DASHBOARD ---
function renderDashboard() {
    var db = document.getElementById('tab-dashboard');
    db.innerHTML = '';
    var container = document.createElement('div');
    container.className = 'dashboard-container container-fluid pt-3';

    var stats = {
        user: (GLOBAL_DATA['User'] || []).length,
        vattu: (GLOBAL_DATA['VatLieu'] || []).length,
        kho: (GLOBAL_DATA['DS_kho'] || []).length,
        phieu: (GLOBAL_DATA['PhieuNhap'] || []).length + (GLOBAL_DATA['PhieuXuat'] || []).length
    };

    var kpiHTML = `
    <h5 class="fw-bold text-success mb-3"><i class="fas fa-chart-pie me-2"></i>TỔNG QUAN</h5>
    <div class="row g-3 mb-4">
        <div class="col-12 col-sm-6 col-xl-3">${kpiCard('Nhân sự', stats.user, 'fas fa-users', '#1E88E5', 'nhansu')}</div>
        <div class="col-12 col-sm-6 col-xl-3">${kpiCard('Vật tư', stats.vattu, 'fas fa-box', '#43A047', 'vattu')}</div>
        <div class="col-12 col-sm-6 col-xl-3">${kpiCard('Kho bãi', stats.kho, 'fas fa-warehouse', '#FB8C00', 'dskho')}</div>
        <div class="col-12 col-sm-6 col-xl-3">${kpiCard('Phiếu NX', stats.phieu, 'fas fa-file-invoice', '#E53935', 'phieunhap')}</div>
    </div>`;
    container.innerHTML = kpiHTML;

    var menuHTML = `<h5 class="fw-bold text-success mb-3"><i class="fas fa-th me-2"></i>CHỨC NĂNG</h5>`;
    MENU_STRUCTURE.forEach(g => {
        menuHTML += `<div class="mb-4"><h6 class="text-muted text-uppercase fw-bold small border-bottom pb-2 mb-3">${g.group}</h6><div class="menu-grid">`;
        g.items.forEach(i => {
            menuHTML += `<div class="menu-card-modern" onclick="switchTab('${i.id}', null)"><i class="${i.icon}"></i><span>${i.title}</span></div>`;
        });
        menuHTML += `</div></div>`;
    });
    container.innerHTML += menuHTML;
    db.appendChild(container);
}

function kpiCard(l, v, i, c, tabId) {
    return `
    <div class="kpi-card-premium" onclick="switchTab('${tabId}', null)">
        <div class="kpi-content">
            <div class="kpi-label">${l}</div>
            <div class="kpi-value" style="color: ${c}">${v}</div>
        </div>
        <div class="kpi-icon-wrap" style="background: linear-gradient(135deg, ${c}, ${c}dd)">
            <i class="${i}"></i>
        </div>
    </div>`;
}

// --- CHUYỂN TAB ---
window.switchTab = function (id, el) {
    document.querySelectorAll('.sidebar-item').forEach(e => e.classList.remove('active'));
    if (el) el.classList.add('active');

    document.querySelectorAll('.content-tab').forEach(e => e.classList.add('d-none'));

    if (id === 'dashboard') {
        document.getElementById('tab-dashboard').classList.remove('d-none');
    } else {
        var item = null;
        MENU_STRUCTURE.forEach(g => { var f = g.items.find(x => x.id == id); if (f) item = f; });

        if (item) {
            CURRENT_SHEET = item.sheet;
            CURRENT_MENU_ID = id;
            document.getElementById('tab-generic').classList.remove('d-none');
            document.getElementById('page-title').innerHTML = `<i class="${item.icon} me-2"></i> ${item.title}`;
            document.getElementById('extra-btn-area').innerHTML = `<button class="btn btn-success rounded-pill fw-bold shadow-sm px-3" onclick="openAddModal()"><i class="fas fa-plus me-1"></i> Thêm mới</button>`;
            document.getElementById('filter-area').innerHTML = `<div class="row g-2"><div class="col-12 col-md-4"><div class="input-group"><span class="input-group-text bg-white border-end-0"><i class="fas fa-search text-muted"></i></span><input type="text" id="search-box" class="form-control border-start-0 ps-0" placeholder="Tìm kiếm dữ liệu..." onkeyup="debounceSearch()"></div></div></div>`;

            renderTable(GLOBAL_DATA[CURRENT_SHEET] || []);
        }
    }
    if (window.innerWidth < 992 && sidebarBS) sidebarBS.hide();
}

// --- TABLE RENDER ---
function renderTable(data) {
    if (!document.getElementById('filter-bar')) {
        var filterHtml = `
            <div id="filter-bar" class="bg-white p-3 rounded-3 shadow-sm mb-3 border">
                <div class="row g-2 align-items-end">
                    <div class="col-md-2">
                        <label class="form-label small fw-bold text-muted mb-1">Từ ngày</label>
                        <input type="date" id="filter-from" class="form-control" oninput="renderTable(GLOBAL_DATA[CURRENT_SHEET] || [])" onchange="renderTable(GLOBAL_DATA[CURRENT_SHEET] || [])">
                    </div>
                    <div class="col-md-2">
                        <label class="form-label small fw-bold text-muted mb-1">Đến ngày</label>
                        <input type="date" id="filter-to" class="form-control" oninput="renderTable(GLOBAL_DATA[CURRENT_SHEET] || [])" onchange="renderTable(GLOBAL_DATA[CURRENT_SHEET] || [])">
                    </div>
                    <div class="col-md-2">
                        <label class="form-label small fw-bold text-muted mb-1">Kho</label>
                        <select id="filter-kho" class="form-select" onchange="renderTable(GLOBAL_DATA[CURRENT_SHEET] || [])">
                            <option value="">-- Tất cả kho --</option>
                        </select>
                    </div>
                    <div class="col-md-2">
                        <label class="form-label small fw-bold text-muted mb-1">Nhân sự</label>
                        <select id="filter-user" class="form-select" onchange="renderTable(GLOBAL_DATA[CURRENT_SHEET] || [])">
                            <option value="">-- Tất cả --</option>
                        </select>
                    </div>
                    <div class="col-md-4">
                        <label class="form-label small fw-bold text-muted mb-1">Tìm kiếm nhanh</label>
                        <div class="input-group">
                            <span class="input-group-text bg-light"><i class="fas fa-search"></i></span>
                            <input type="text" id="main-search" class="form-control" placeholder="Gõ để tìm kiếm..." oninput="renderTable(GLOBAL_DATA[CURRENT_SHEET] || [])">
                        </div>
                    </div>
                </div>
            </div>`;
        document.getElementById('app-container').querySelector('.main-content').insertAdjacentHTML('afterbegin', filterHtml);

        // Populate Dropdowns
        var khoData = GLOBAL_DATA['DS_kho'] || [];
        var userData = GLOBAL_DATA['User'] || [];

        var khoSelect = document.getElementById('filter-kho');
        if (khoSelect) {
            khoData.forEach(k => {
                var name = k['Tên kho'] || k['Ten_Kho'] || k['Tên Kho'];
                if (name) {
                    var opt = document.createElement('option');
                    opt.value = name; opt.innerText = name;
                    khoSelect.appendChild(opt);
                }
            });
        }

        var userSelect = document.getElementById('filter-user');
        if (userSelect) {
            userData.forEach(u => {
                var name = u['Họ và tên'] || u['Name'];
                if (name) {
                    var opt = document.createElement('option');
                    opt.value = name; opt.innerText = name;
                    userSelect.appendChild(opt);
                }
            });
        }
    }

    var tb = document.querySelector('#data-table tbody');
    var th = document.querySelector('#data-table thead');
    if (!tb || !th) return;
    th.innerHTML = ''; tb.innerHTML = '';

    var filterVal = (document.getElementById('main-search')?.value || '').toLowerCase();
    var filterKho = document.getElementById('filter-kho')?.value || '';
    var filterUser = document.getElementById('filter-user')?.value || '';
    var filterFrom = document.getElementById('filter-from')?.value || '';
    var filterTo = document.getElementById('filter-to')?.value || '';

    var activeData = data.filter(r => {
        if (r['Delete'] == 'X') return false;

        var rowText = Object.values(r).join(' ').toLowerCase();
        if (filterVal && !rowText.includes(filterVal)) return false;

        if (filterKho) {
            var kho = r['Tên kho'] || r['Ten_Kho'] || r['Tên Kho'] || '';
            if (kho != filterKho) return false;
        }

        if (filterUser) {
            var user = r['Người nhập'] || r['Người xuất'] || r['Người lập'] || r['Nhân viên'] || r['Họ và tên'] || '';
            if (user != filterUser) return false;
        }

        var rowDate = '';
        for (var k in r) { if (k.toLowerCase().includes('ngày') || k.toLowerCase().includes('date')) { rowDate = r[k]; break; } }
        if (rowDate) {
            var d = new Date(rowDate);
            if (filterFrom && d < new Date(filterFrom)) return false;
            if (filterTo) {
                var dTo = new Date(filterTo); dTo.setHours(23, 59, 59, 999);
                if (d > dTo) return false;
            }
        }

        return true;
    });

    if (activeData.length === 0) { tb.innerHTML = '<tr><td colspan="100%" class="text-center p-5 text-muted fst-italic">Không có dữ liệu hiển thị</td></tr>'; return; }

    var keys = Object.keys(activeData[0]);
    var skip = ['Delete', 'Update', 'App_pass', 'Who delete', 'Path file', 'file path'];
    var trH = document.createElement('tr');
    keys.forEach(k => { if (!skip.includes(k)) { var t = document.createElement('th'); t.innerText = COLUMN_MAP[k] || k; trH.appendChild(t); } });
    trH.appendChild(document.createElement('th'));
    th.appendChild(trH);

    activeData.forEach((r) => {
        var absoluteIdx = data.indexOf(r);
        var tr = document.createElement('tr');
        tr.style.cursor = 'pointer';
        keys.forEach(k => {
            if (!skip.includes(k)) {
                var td = document.createElement('td'); var v = r[k];
                if (k.toUpperCase().includes('NGAY') || k.toUpperCase().includes('DATE')) v = formatSafeDate(v);
                if ((k.toUpperCase().includes('GIA') || k.toUpperCase().includes('TIEN')) && !isNaN(v) && v !== '') { v = Number(v).toLocaleString('vi-VN'); td.className = 'text-end fw-bold text-success'; }
                td.innerText = v || ''; tr.appendChild(td);
            }
        });

        var tdA = document.createElement('td'); tdA.className = 'text-center sticky-col';
        var btnGroup = document.createElement('div'); btnGroup.className = 'btn-group';

        var btnEdit = document.createElement('button'); btnEdit.className = 'btn btn-sm btn-light border text-success shadow-sm'; btnEdit.innerHTML = '<i class="fas fa-pen"></i>'; btnEdit.title = 'Sửa';
        btnEdit.onclick = (e) => { e.stopPropagation(); openEditModal(absoluteIdx); };

        var btnDel = document.createElement('button'); btnDel.className = 'btn btn-sm btn-light border text-danger shadow-sm'; btnDel.innerHTML = '<i class="fas fa-trash-alt"></i>'; btnDel.title = 'Xóa';
        btnDel.onclick = (e) => { e.stopPropagation(); deleteRow(absoluteIdx); };

        btnGroup.appendChild(btnEdit); btnGroup.appendChild(btnDel);
        tdA.appendChild(btnGroup); tr.appendChild(tdA);

        tr.onclick = () => showRowDetail(r, CURRENT_SHEET, absoluteIdx);
        tb.appendChild(tr);
    });
}

function formatSafeDate(d) {
    if (!d || d === 'Invalid Date') return '';
    var dt = new Date(d);
    if (!isNaN(dt.getTime())) return dt.toLocaleDateString('vi-VN');
    return d;
}

// --- UTILS ---
window.debounceSearch = function () {
    if (FILTER_TIMER) clearTimeout(FILTER_TIMER);
    FILTER_TIMER = setTimeout(() => {
        var key = document.getElementById('search-box').value.toLowerCase();
        var raw = GLOBAL_DATA[CURRENT_SHEET] || [];
        var filtered = raw.filter(r => Object.values(r).join(' ').toLowerCase().includes(key));
        renderTable(filtered);
    }, 300);
}

function showLoading(s, text) {
    var overlay = document.getElementById('loading-overlay');
    if (text) overlay.querySelector('.loader-text').innerText = text;
    else overlay.querySelector('.loader-text').innerText = 'Đang xử lý dữ liệu...';

    if (s) {
        overlay.style.display = 'flex';
        if (LOADING_TIMEOUT) clearTimeout(LOADING_TIMEOUT);
        LOADING_TIMEOUT = setTimeout(function () { overlay.style.display = 'none'; }, 30000);
    } else {
        overlay.style.display = 'none';
        if (LOADING_TIMEOUT) clearTimeout(LOADING_TIMEOUT);
    }
}

// function handleError is already defined at the top

function renderSidebar() {
    var h = '<div class="py-2">';
    MENU_STRUCTURE.forEach(g => {
        h += `<div class="sidebar-heading">${g.group}</div>`;
        g.items.forEach(i => {
            h += `<div class="sidebar-item" onclick="switchTab('${i.id}', this)"><i class="${i.icon}"></i>${i.title}</div>`;
        });
    });
    document.querySelector('#sidebarMenu .offcanvas-body').innerHTML = h + '</div>';
}

window.openAddModal = function () { EDIT_INDEX = -1; document.getElementById('modalTitle').innerText = 'Thêm Mới'; buildForm(CURRENT_SHEET, null); new bootstrap.Modal(document.getElementById('dataModal')).show(); }

window.openEditModal = function (i) { EDIT_INDEX = i; document.getElementById('modalTitle').innerText = 'Cập Nhật'; var d = (GLOBAL_DATA[CURRENT_SHEET] || [])[i]; buildForm(CURRENT_SHEET, d); new bootstrap.Modal(document.getElementById('dataModal')).show(); }

function buildForm(s, d) {
    var f = document.getElementById('dynamic-form'); f.innerHTML = '';
    var smp = GLOBAL_DATA[s]?.[0] || (GLOBAL_DATA[s] && GLOBAL_DATA[s].length > 0 ? GLOBAL_DATA[s][0] : {});
    if (Object.keys(smp).length === 0) { f.innerHTML = '<div class="alert alert-warning">Chưa có dữ liệu mẫu.</div>'; return; }
    var skip = ['Delete', 'Update', 'App_pass'];
    Object.keys(smp).forEach(k => {
        if (skip.includes(k)) return;
        var v = d ? d[k] : '';
        if (d && (k.includes('Ngay') || k.includes('Date'))) {
            try { v = new Date(v).toISOString().split('T')[0]; } catch (e) { v = ''; }
        }
        var div = document.createElement('div');
        div.className = 'col-md-6';
        div.innerHTML = `<label class="form-label small fw-bold text-muted">${COLUMN_MAP[k] || k}</label><input class="form-control rounded-3" name="${k}" value="${v}" ${k.includes('Ngay') ? 'type="date"' : ''}>`;
        f.appendChild(div);
    });
}

window.submitData = function () {
    var f = document.getElementById('dynamic-form'); var fd = {};
    f.querySelectorAll('input').forEach(i => fd[i.name] = i.value);
    if (!fd['NguoiLap'] && CURRENT_USER) fd['NguoiLap'] = CURRENT_USER.name;
    showLoading(true);
    bootstrap.Modal.getInstance(document.getElementById('dataModal')).hide();

    var promise = (EDIT_INDEX === -1)
        ? callGAS('saveData', CURRENT_SHEET, fd)
        : callGAS('updateData', CURRENT_SHEET, fd, EDIT_INDEX);

    promise.then(r => {
        if (r.status == 'success') {
            const Toast = Swal.mixin({ toast: true, position: 'top-end', showConfirmButton: false, timer: 2000 });
            Toast.fire({ icon: 'success', title: 'Đã lưu thành công' });

            var savedSheet = CURRENT_SHEET;
            // Nếu có SHEET_BEFORE_ADD (đang thêm bản ghi con), khôi phục lại sheet chính sau khi load xong data con
            if (window.SHEET_BEFORE_ADD) {
                refreshSingleSheet(savedSheet).then(() => {
                    CURRENT_SHEET = window.SHEET_BEFORE_ADD;
                    delete window.SHEET_BEFORE_ADD;
                    // Tùy chọn: Mở lại detail của cha để thấy dòng mới (Phức tạp hơn, để sau)
                });
            } else {
                refreshSingleSheet(CURRENT_SHEET);
            }
        } else { showLoading(false); alert(r.message); }
    }).catch(handleError);
}

window.deleteRow = function (i, sheet) {
    var targetSheet = sheet || CURRENT_SHEET;
    var rowData = (GLOBAL_DATA[targetSheet] || [])[i];
    if (!rowData) return;

    var mappedTable = TABLE_MAP[targetSheet] || targetSheet;
    var pkField = PK_MAP[mappedTable] || 'ID';
    var idValue = rowData[pkField];

    Swal.fire({
        title: 'Xác nhận xóa?',
        text: "Dữ liệu sẽ được đánh dấu xóa!",
        icon: 'warning',
        showCancelButton: true,
        confirmButtonColor: '#2E7D32',
        cancelButtonColor: '#d33',
        confirmButtonText: 'Đồng ý',
        cancelButtonText: 'Hủy'
    }).then((result) => {
        if (result.isConfirmed) {
            showLoading(true);
            callGAS('delete', targetSheet, null, idValue).then(r => {
                if (r.status == 'success' || r.status == 'online-success') {
                    refreshSingleSheet(targetSheet);
                } else { showLoading(false); alert(r.message || 'Lỗi khi xóa dữ liệu'); }
            }).catch(handleError);
        }
    });
}

function showRowDetail(r, s, idx) {
    if (!r && idx !== -1) r = (GLOBAL_DATA[s] || [])[idx];
    if (!r) return;

    var menuTitle = 'bản ghi';
    MENU_STRUCTURE.forEach(g => { var f = g.items.find(x => x.sheet == s); if (f) menuTitle = f.title; });

    var html = '<div class="row text-start" style="font-size:12px !important; max-height: 400px; overflow-y: auto; overflow-x: hidden; padding: 15px; background: #fdfdfd; border-radius: 12px;">';
    var skip = ['Delete', 'Update', 'App_pass', 'Who delete', 'ID_Phieu', 'ID phieu nhap', 'ID phieu Xuat', 'ID Chkho', 'Path file', 'file path'];
    for (var k in r) {
        if (!skip.includes(k) && r[k]) {
            var val = r[k] || '';
            if (k.toUpperCase().includes('NGAY') || k.toUpperCase().includes('DATE')) val = formatSafeDate(val);

            var displayVal = val;
            if (typeof val === 'string' && (val.startsWith('http') || val.startsWith('https'))) {
                if (k.toLowerCase().includes('hình ảnh') || val.match(/\.(jpeg|jpg|gif|png)$/i)) {
                    displayVal = `<div class="mt-1"><img src="${val}" style="max-height:180px; max-width:100%; border-radius:8px; cursor:zoom-in; border: 2px solid #fff; shadow: 0 4px 6px rgba(0,0,0,0.1)" onclick="window.open('${val}')"></div>`;
                } else {
                    displayVal = `<a href="${val}" target="_blank" class="btn btn-xs btn-outline-primary py-0 px-2 mt-1" style="font-size:10px">Mở tài liệu <i class="fas fa-external-link-alt ms-1"></i></a>`;
                }
            } else if (!isNaN(val) && val !== '' && (k.toLowerCase().includes('tiền') || k.toLowerCase().includes('giá') || k.toLowerCase().includes('thành'))) {
                displayVal = `<span class="text-success fw-bold">${Number(val).toLocaleString('vi-VN')}</span>`;
            }

            html += `<div class="col-md-4 mb-3 border-bottom pb-2"><small class="text-muted fw-bold d-block text-uppercase" style="font-size:9px; letter-spacing: 0.5px;">${COLUMN_MAP[k] || k}</small><span class="fw-medium text-dark" style="font-size:12px">${displayVal}</span></div>`;
        }
    }
    html += '</div>';

    if (RELATION_CONFIG[s]) {
        var cf = RELATION_CONFIG[s];
        var cS = cf.child;
        var fk = cf.foreignKey;

        var mappedTable = TABLE_MAP[s] || s;
        var pkField = PK_MAP[mappedTable] || 'ID';
        var pid = r[pkField];

        var childs = (GLOBAL_DATA[cS] || []).filter(c => String(c[fk]) == String(pid) && c['Delete'] != 'X');

        html += `<div class="mt-3 bg-light p-2 rounded-3 border">
            <h6 class="text-success fw-bold mb-2 d-flex justify-content-between align-items-center" style="font-size:14px">
                <span><i class="fas fa-list-ul me-2"></i>${cf.title}</span> 
                <div>
                    <span class="badge bg-success rounded-pill me-2">${childs.length}</span>
                    <button class="btn btn-xs btn-primary rounded-pill px-2" onclick="Swal.close(); setTimeout(() => openAddChildModal('${cS}', '${fk}', '${pid}'), 200)" style="font-size:10px"><i class="fas fa-plus me-1"></i>Thêm dòng</button>
                </div>
            </h6>`;

        if (childs.length > 0) {
            html += `<div class="table-responsive" style="max-height: 250px; overflow-y: auto; border-radius: 8px;"><table class="table table-sm table-bordered custom-table mb-0" style="font-size:11px"><thead><tr>`;
            var ck = Object.keys(childs[0]);
            ck.forEach(x => { if (!skip.includes(x) && x != fk) html += `<th>${COLUMN_MAP[x] || x}</th>`; });
            html += '<th class="text-center" style="width:40px">#</th></tr></thead><tbody>';
            childs.forEach((cr) => {
                var originalCIdx = (GLOBAL_DATA[cS] || []).indexOf(cr);
                html += '<tr>';
                ck.forEach(x => {
                    if (!skip.includes(x) && x != fk) {
                        var cv = cr[x] || '';
                        if (typeof cv === 'string' && (cv.startsWith('http') || cv.startsWith('https'))) {
                            cv = `<a href="${cv}" target="_blank" class="text-primary"><i class="fas fa-paperclip"></i></a>`;
                        } else if (!isNaN(cv) && cv !== '' && (x.toLowerCase().includes('tiền') || x.toLowerCase().includes('giá') || x.toLowerCase().includes('thành'))) {
                            cv = `<b>${Number(cv).toLocaleString('vi-VN')}</b>`;
                        }
                        html += `<td>${cv}</td>`;
                    }
                });
                html += `<td class="text-center"><button class="btn btn-xs btn-info py-0 px-1 shadow-sm" title="Xem chi tiết" onclick="Swal.close(); setTimeout(() => showRowDetail(null, '${cS}', ${originalCIdx}), 200);"><i class="fas fa-eye text-white" style="font-size:10px"></i></button></td>`;
                html += '</tr>';
            });
            html += '</tbody></table></div>';
        } else {
            html += '<div class="text-center p-3 text-muted small fst-italic">Chưa có dữ liệu chi tiết</div>';
        }
        html += '</div>';
    }

    Swal.fire({
        title: 'Chi tiết ' + menuTitle,
        html: html,
        width: '1050px',
        showCloseButton: true,
        showCancelButton: true,
        showDenyButton: true,
        confirmButtonText: '<i class="fas fa-save me-1"></i> Lưu',
        denyButtonText: '<i class="fas fa-edit me-1"></i> Sửa',
        cancelButtonText: '<i class="fas fa-times me-1"></i> Hủy',
        footer: `<div class="w-100 d-flex justify-content-between align-items-center">
                    <span class="text-muted small">Mã hệ thống: ${pid || idx}</span>
                    <button class="btn btn-sm btn-danger px-4 shadow-sm rounded-pill" onclick="Swal.close(); deleteRow(${idx}, '${s}')"><i class="fas fa-trash-alt me-1"></i> XÓA BẢN GHI NÀY</button>
                 </div>`,
        customClass: {
            title: 'detail-modal-title',
            popup: 'rounded-4 shadow-lg',
            confirmButton: 'btn btn-success rounded-pill px-4 shadow-sm',
            denyButton: 'btn btn-warning text-dark rounded-pill px-4 ms-2 shadow-sm',
            cancelButton: 'btn btn-light rounded-pill px-4 border ms-2'
        },
        buttonsStyling: false
    }).then((result) => {
        if (result.isConfirmed) {
            Swal.fire({ icon: 'info', title: 'Thông báo', text: 'Vui lòng nhấn nút Sửa để thay đổi dữ liệu!', timer: 1500, showConfirmButton: false });
        } else if (result.isDenied) {
            openEditModal(idx, s);
        }
    });
}

window.openAddChildModal = function (childSheet, foreignKey, parentId) {
    EDIT_INDEX = -1;
    window.SHEET_BEFORE_ADD = CURRENT_SHEET; // Lưu lại sheet cha
    CURRENT_SHEET = childSheet; // Tạm thời chuyển sheet để build form
    document.getElementById('modalTitle').innerText = 'Thêm chi tiết: ' + (COLUMN_MAP[childSheet] || childSheet);

    var smp = GLOBAL_DATA[childSheet]?.[0] || {};
    var prefillData = {};
    prefillData[foreignKey] = parentId;

    buildForm(childSheet, prefillData);
    new bootstrap.Modal(document.getElementById('dataModal')).show();
}

window.viewProfile = function () {
    if (!CURRENT_USER) return;
    var html = `<div class="text-center mb-3"><div class="avatar-circle mx-auto bg-success text-white" style="width:80px;height:80px;font-size:2rem"><i class="fas fa-user"></i></div><h4 class="mt-2">${CURRENT_USER.name}</h4><span class="badge bg-secondary">${CURRENT_USER.role || 'Nhân viên'}</span></div>`;
    Swal.fire({ html: html, showConfirmButton: false });
}
