// --- CẤU HÌNH API ---
// (Google Apps Script URL không dùng khi đã chuyển sang Supabase)

// --- KẾT NỐI SUPABASE - TẢI DỮ LIỆU CHO TẤT CẢ BẢNG ---
// Dữ liệu được tải khi đăng nhập qua getInitialData() → callSupabase('read', table).
// Thêm/Sửa/Xóa dùng callGAS('saveData'|'updateData'|'delete') → callSupabase('insert'|'update'|'delete').
// Tên bảng Supabase map trong TABLE_MAP, khóa chính trong PK_MAP.
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
            enableOnlineFeatures();
        } else {
            indicator.className = 'position-fixed top-0 end-0 m-2 badge bg-danger';
            text.innerText = 'Đang ngoại tuyến (Offline)';
            enableOfflineMode();
        }
    }
    
    // Update mobile bottom navigation
    updateMobileOfflineIndicator();
}

// Cải thiện tính năng offline cho mobile
function enableOfflineMode() {
    // Hiển thị thông báo offline cho mobile
    if (window.innerWidth <= 768) {
        const offlineBanner = document.createElement('div');
        offlineBanner.id = 'mobile-offline-banner';
        offlineBanner.className = 'alert alert-warning alert-dismissible fade show position-fixed w-100';
        offlineBanner.style.cssText = 'top: 60px; z-index: 1030; margin: 0;';
        offlineBanner.innerHTML = `
            <i class="fas fa-wifi-slash me-2"></i>
            <strong>Chế độ offline:</strong> Dữ liệu sẽ được lưu tạm và đồng bộ khi có mạng.
            <button type="button" class="btn-close" data-bs-dismiss="alert"></button>
        `;
        
        // Remove existing banner if exists
        const existing = document.getElementById('mobile-offline-banner');
        if (existing) existing.remove();
        
        document.body.appendChild(offlineBanner);
        
        // Auto hide after 10 seconds
        setTimeout(() => {
            if (offlineBanner && offlineBanner.parentNode) {
                offlineBanner.remove();
            }
        }, 10000);
    }
    
    // Enable offline-only features
    enableOfflineCache();
    restrictOnlineOnlyFeatures();
}

function enableOnlineFeatures() {
    // Remove offline banner
    const offlineBanner = document.getElementById('mobile-offline-banner');
    if (offlineBanner) offlineBanner.remove();
    
    // Re-enable online-only features
    enableOnlineOnlyFeatures();
}

function updateMobileOfflineIndicator() {
    const mobileNav = document.querySelector('.mobile-bottom-nav');
    if (mobileNav) {
        const indicator = mobileNav.querySelector('.offline-indicator') || document.createElement('div');
        indicator.className = 'offline-indicator position-absolute';
        indicator.style.cssText = 'top: -5px; right: 10px; z-index: 10;';
        
        if (!window.isOnline) {
            indicator.innerHTML = '<i class="fas fa-wifi-slash text-danger" style="font-size: 12px;"></i>';
            if (!mobileNav.contains(indicator)) mobileNav.appendChild(indicator);
        } else {
            if (mobileNav.contains(indicator)) indicator.remove();
        }
    }
}

function enableOfflineCache() {
    // Lưu trữ dữ liệu quan trọng vào localStorage để sử dụng offline
    const criticalTables = ['User', 'BangLuongThang', 'Chamcong', 'LichSuLuong', 'GiaoDichLuong'];
    
    criticalTables.forEach(table => {
        if (GLOBAL_DATA[table]) {
            localStorage.setItem(`offline_${table}`, JSON.stringify(GLOBAL_DATA[table]));
        }
    });
    
    localStorage.setItem('offline_last_sync', new Date().toISOString());
}

function restrictOnlineOnlyFeatures() {
    // Disable các tính năng chỉ hoạt động online
    const onlineOnlyButtons = document.querySelectorAll('[data-requires-online="true"]');
    onlineOnlyButtons.forEach(btn => {
        btn.disabled = true;
        btn.title = 'Tính năng này cần kết nối mạng';
    });
}

function enableOnlineOnlyFeatures() {
    const onlineOnlyButtons = document.querySelectorAll('[data-requires-online="true"]');
    onlineOnlyButtons.forEach(btn => {
        btn.disabled = false;
        btn.title = '';
    });
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
            { id: 'nhansu', sheet: 'User', icon: 'fas fa-users', title: 'Hồ sơ Nhân sự' },
            { id: 'lichsuluong', sheet: 'LichSuLuong', icon: 'fas fa-cogs', title: 'Cài đặt lương' }
        ]
    },
    {
        group: 'DANH MỤC HỆ THỐNG', items: [
            { id: 'dskho', sheet: 'DS_kho', icon: 'fas fa-warehouse', title: 'Danh sách Kho' },
            { id: 'vattu', sheet: 'Vat_tu', icon: 'fas fa-tools', title: 'Danh mục Vật tư' },
            { id: 'nhomvattu', sheet: 'Nhomvattu', icon: 'fas fa-layer-group', title: 'Nhóm vật tư' },
            { id: 'doitac', sheet: 'Doitac', icon: 'fas fa-handshake', title: 'Đối tác / NCC' }
        ]
    }
];

// Cột không hiển thị hoàn toàn (bảng, form, chi tiết, bảng con)
const COLUMNS_HIDDEN = ['Delete', 'delete', 'UPDATE', 'Update', 'update', 'DELETE', 'Path file', 'file path', 'Pathfile', 'path_file', 'Đường dẫn file', 'App_pass', 'Who delete'];

// Việt hóa: từ điển Tên Cột Gốc -> Tên Tiếng Việt (chuyển ngữ)
const COLUMN_MAP = {
    'ID_ChamCong': 'Mã chấm công',
    'Mahienthi': 'Mã hiển thị',
    'ID_NhanVien': 'Mã nhân viên',
    'Ngày': 'Ngày (Ghi nhận)',
    'Ngay': 'Ngày (Ghi nhận)',
    'SoGioCong': 'Số giờ công',
    'SoGioTangCa': 'Số giờ tăng ca',
    'Kholamviec': 'Kho làm việc',
    'NoiDungCongViec': 'Nội dung công việc',
    'TienTangCaNgay': 'Tiền tăng ca ngày',
    'LuongTaiThoiDiem': 'Lương tại thời điểm',
    'Luongcongnhat': 'Lương công nhật',
    'HeSoTangCaTaiThoiDiem': 'Hệ số tăng ca thời điểm',
    'TienCongNgay': 'Tiền công ngày',
    'ID_LichSu': 'Mã lịch sử lương',
    'NgayApDung': 'Ngày áp dụng',
    'LuongCoBan_Ngay': 'Lương cơ bản ngày',
    'HeSoTangCa': 'Hệ số tăng ca',
    'ID_GiaoDich': 'Mã giao dịch',
    'MaGiaoDich': 'Mã hiển thị giao dịch',
    'NgayGiaoDich': 'Ngày giao dịch',
    'LoaiGiaoDich': 'Loại giao dịch',
    'SoTien': 'Số tiền',
    'GhiChu': 'Ghi chú',
    'Ghi_Chu': 'Ghi chú',
    'Noi_Dung': 'Nội dung',
    'ID_BGCC': 'Mã bàn giao công cụ',
    'MaVatTu': 'Mã vật tư',
    'SoLuong': 'Số lượng',
    'So_Luong': 'Số lượng',
    'NgayBanGiao': 'Ngày bàn giao',
    'TongGiaTri': 'Tổng giá trị',
    'ID phieu nhap': 'Mã phiếu nhập',
    'Tên Kho': 'Tên kho',
    'Tên kho': 'Tên kho',
    'Ten_Kho': 'Tên kho',
    'Người nhập': 'Người nhập kho',
    'Diễn giải': 'Diễn giải',
    'Trangthai': 'Trạng thái',
    'Trang_Thai': 'Trạng thái',
    'Trạng thái': 'Trạng thái',
    'Tổng tiền': 'Tổng tiền',
    'Tong_Tien': 'Tổng tiền',
    'Phê duyệt': 'Người phê duyệt',
    'ID phieu Xuat': 'Mã phiếu xuất',
    'Người xuất': 'Người xuất kho',
    'Người Xuất': 'Người xuất kho',
    'ID xuat chi tiet': 'Mã chi tiết xuất',
    'ID vật tư': 'Mã vật tư (ID)',
    'Nhóm vật liệu': 'Nhóm vật liệu',
    'Số lượng xuất': 'Số lượng xuất',
    'Đơn vị tính': 'Đơn vị tính',
    'DonViTinh': 'Đơn vị tính',
    'ĐVT': 'Đơn vị tính',
    'Đơn giá': 'Đơn giá',
    'Don_Gia': 'Đơn giá',
    'Dongia': 'Đơn giá',
    'Thành tiền': 'Thành tiền',
    'Thanh_Tien': 'Thành tiền',
    'Hình ảnh': 'Hình ảnh',
    'Who delete': 'Người xóa',
    'ID Chkho': 'Mã chuyển kho',
    'KhoDi': 'Kho đi',
    'KhoDen': 'Kho đến',
    'Kho nguồn': 'Kho nguồn',
    'Kho đích': 'Kho đích',
    'NgayChuyen': 'Ngày chuyển',
    'NguoiChuyen': 'Người chuyển',
    'NguoiNhan': 'Người nhận',
    'ID nhap chi tiet': 'Mã chi tiết nhập',
    'Số lượng nhập': 'Số lượng nhập',
    'ID kho': 'Mã kho (ID)',
    'ID_kho': 'Mã kho (ID)',
    'Địa chỉ': 'Địa chỉ',
    'Dia_diem': 'Địa điểm',
    'Địa điểm': 'Địa điểm',
    'Nhân viên phụ trách': 'Nhân viên phụ trách',
    'Quản lý kho': 'Quản lý kho',
    'Tọa độ': 'Tọa độ',
    'Sức chứa': 'Sức chứa',
    'Tên vật tư': 'Tên vật tư',
    'Ten_VT': 'Tên vật tư',
    'Quy cách': 'Quy cách',
    'Nhóm vật tư': 'Nhóm vật tư',
    'Nhom_Vat_Tu': 'Nhóm vật tư',
    'Mô tả': 'Mô tả',
    'Mo_ta': 'Mô tả',
    'ID_TonKho': 'Mã tồn kho',
    'SoLuongTon': 'Số lượng tồn',
    'Số lượng tồn': 'Số lượng tồn',
    'So_luong_ton': 'Số lượng tồn',
    'Ngày cập nhật gần nhất': 'Ngày cập nhật cuối',
    'LoaiNX': 'Loại Nhập/Xuất',
    'ID CKchitiet': 'Mã chi tiết chuyển kho',
    'ID NVT': 'Mã nhóm vật tư',
    'ID_NVT': 'Mã nhóm vật tư',
    'ID_BangLuong': 'Mã bảng lương',
    'HoTen': 'Họ và tên',
    'Họ và tên': 'Họ và tên',
    'Thang': 'Tháng',
    'Tháng': 'Tháng',
    'Nam': 'Năm',
    'Năm': 'Năm',
    'NgayTinhLuong': 'Ngày tính lương',
    'KyLuong_TuNgay': 'Kỳ lương từ ngày',
    'KyLuong_DenNgay': 'Kỳ lương đến ngày',
    'TongGioCong': 'Tổng giờ công',
    'TongGioTangCa': 'Tổng giờ tăng ca',
    'SoNgayCongQuyDoi': 'Số ngày công quy đổi',
    'TongLuongCoBan': 'Tổng lương cơ bản',
    'TongTienTangCa': 'Tổng tiền tăng ca',
    'TongPhuCap': 'Tổng phụ cấp',
    'TongThuong': 'Tổng thưởng',
    'TongThuNhap': 'Tổng thu nhập',
    'TongTamUng': 'Tổng tạm ứng',
    'TienBaoHiem': 'Tiền bảo hiểm',
    'GiamTruKhac': 'Giảm trừ khác',
    'TongGiamTru': 'Tổng giảm trừ',
    'PhatSinhTrongKy': 'Phát sinh trong kỳ',
    'Thuclanh': 'Thực lãnh',
    'Thuc_linh': 'Thực lãnh',
    'SoDuDauKy': 'Số dư đầu kỳ',
    'SoDuCuoiKy': 'Số dư cuối kỳ',
    'ID_Note': 'Mã ghi chú',
    'NguoiTao': 'Người tạo',
    'NgayTao': 'Ngày tạo',
    'ThoiTiet': 'Thời tiết',
    'NoiDung': 'Nội dung',
    'Nội dung': 'Nội dung',
    'DoiTuongLienQuan': 'Đối tượng liên quan',
    'ID_DoiTuong': 'Mã đối tượng',
    'Huydongxemay': 'Huy động xe máy',
    'Những điểm lưu ý': 'Lưu ý quan trọng',
    'NgayNhac': 'Ngày nhắc',
    'DaNhac': 'Đã nhắc (Trạng thái)',
    'Location': 'Vị trí / Tọa độ',
    'ID_Nhac': 'Mã nhắc nhở',
    'TieuDe': 'Tiêu đề',
    'NguoiNhac': 'Người nhắc',
    'NgayGioNhac': 'Thời gian nhắc',
    'ID_ChiPhi': 'Mã chi phí',
    'NgayChiPhi': 'Ngày chi phí',
    'NgayChiphi': 'Ngày chi',
    'NguoiLap': 'Người lập phiếu',
    'Nguoilap': 'Người lập',
    'TongSoTien': 'Tổng số tiền',
    'ID': 'Mã nhân viên (ID)',
    'Email': 'Email',
    'Số điện thoại': 'Số điện thoại',
    'SĐT': 'Số điện thoại',
    'Chứng minh nhân dân': 'CMND / CCCD',
    'CCCD': 'CMND / CCCD',
    'Ngày sinh': 'Ngày sinh',
    'Ngày vào làm việc': 'Ngày vào làm',
    'Ngày vào làm': 'Ngày vào làm',
    'Mã số thuế': 'Mã số thuế',
    'App_pass': 'Mật khẩu ứng dụng',
    'Bộ phận': 'Bộ phận',
    'Chức vụ': 'Chức vụ',
    'Tinh luong': 'Có tính lương',
    'Phân quyền': 'Phân quyền',
    'Quyền xem': 'Quyền xem dữ liệu',
    'Ảnh cá nhân': 'Ảnh cá nhân',
    'Ngày kết thúc làm': 'Ngày nghỉ việc',
    'Ngansachdauky': 'Ngân sách đầu kỳ',
    'ID_ChiTiet': 'Mã chi tiết chi',
    'Ngaychi': 'Ngày chi',
    'LoaiChiPhi': 'Loại chi phí',
    'NhapKho': 'Có nhập kho',
    'NhapKho_Xacnhan': 'Xác nhận nhập kho',
    'Mã nhập xuất': 'Mã nhập xuất',
    'ID lương': 'Mã bản lương (ID)',
    'ID nhân viên': 'Mã nhân viên',
    'Số ngày công': 'Số ngày công',
    'Lương cơ bản': 'Lương cơ bản',
    'Luong_co_ban': 'Lương cơ bản',
    'Thưởng': 'Thưởng',
    'Thuong': 'Thưởng',
    'Tạm ứng': 'Tạm ứng',
    'Tổng lương': 'Tổng lương',
    'ID đối tác': 'Mã đối tác',
    'ID_doitac': 'Mã đối tác',
    'Tên đối tác': 'Tên đối tác',
    'Ten_doitac': 'Tên đối tác',
    'Loại đối tác': 'Loại đối tác',
    'Người liên hệ': 'Người liên hệ',
    'Nguoi_lien_he': 'Người liên hệ',
    'Website': 'Website',
    'Tên nhóm': 'Tên nhóm',
    'Ten_nhom': 'Tên nhóm',
    'created_at': 'Ngày tạo',
    'updated_at': 'Ngày cập nhật'
};


// Quan hệ Cha -> Con -> Cháu: mỗi bảng có tối đa 1 bảng con (child). Bấm cha -> chi tiết cha + danh sách con; bấm con -> chi tiết con + danh sách cháu; bấm cháu -> chi tiết cháu.
const RELATION_CONFIG = {
    // Cấp 1 (Cha)
    'Phieunhap': { child: 'NhapChiTiet', foreignKey: 'ID phieu nhap', title: 'CHI TIẾT NHẬP' },
    'Phieuxuat': { child: 'XuatChiTiet', foreignKey: 'ID phieu Xuat', title: 'CHI TIẾT XUẤT' },
    'Phieuchuyenkho': { child: 'Chuyenkhochitiet', foreignKey: 'ID Chkho', title: 'CHI TIẾT CHUYỂN' },
    'Chiphi': { child: 'Chiphichitiet', foreignKey: 'ID_ChiPhi', title: 'CHI TIẾT CHI' },
    'Nhomvattu': { child: 'Vat_tu', foreignKey: 'Nhóm vật tư', title: 'DANH MỤC VẬT TƯ TRONG NHÓM' },
    // Cấp 2 (Con -> Cháu): thêm bảng cháu tại đây nếu có. Ví dụ:
    // 'NhapChiTiet': { child: 'TenBangChau', foreignKey: 'ID nhap chi tiet', title: 'CHI TIẾT CẤP 2' },
    // 'Chiphichitiet': { child: 'TenBangChau', foreignKey: 'ID_ChiTiet', title: 'CHI TIẾT CẤP 2' }
};

// Droplist cho form Thêm/Sửa: theo chuyên đề dùng bảng Vat_tu, DS_kho, User, Doitac, Nhomvattu...
// Mỗi dòng: sheet (hoặc '*') -> field -> { table, labelKey, valueKey }. '*' áp dụng cho mọi sheet có field đó.
const DROPDOWN_CONFIG = {
    '*': {
        'ID vật tư': { table: 'Vat_tu', labelKey: 'Tên vật tư', valueKey: 'ID vật tư' },
        'MaVatTu': { table: 'Vat_tu', labelKey: 'Tên vật tư', valueKey: 'ID vật tư' },
        'Mã vật tư': { table: 'Vat_tu', labelKey: 'Tên vật tư', valueKey: 'ID vật tư' },
        'Tên kho': { table: 'DS_kho', labelKey: 'Tên kho', valueKey: 'ID kho' },
        'Ten_Kho': { table: 'DS_kho', labelKey: 'Tên kho', valueKey: 'ID kho' },
        'ID kho': { table: 'DS_kho', labelKey: 'Tên kho', valueKey: 'ID kho' },
        'Người nhập': { table: 'User', labelKey: 'Họ và tên', valueKey: 'ID' },
        'Người xuất': { table: 'User', labelKey: 'Họ và tên', valueKey: 'ID' },
        'NguoiLap': { table: 'User', labelKey: 'Họ và tên', valueKey: 'ID' },
        'Người lập': { table: 'User', labelKey: 'Họ và tên', valueKey: 'ID' },
        'Nhân viên': { table: 'User', labelKey: 'Họ và tên', valueKey: 'ID' },
        'ID NVT': { table: 'Nhomvattu', labelKey: 'Tên nhóm', valueKey: 'ID NVT' },
        'Nhóm vật tư': { table: 'Nhomvattu', labelKey: 'Tên nhóm', valueKey: 'ID NVT' },
        'ID đối tác': { table: 'Doitac', labelKey: 'Tên đối tác', valueKey: 'ID đối tác' },
        'Tên đối tác': { table: 'Doitac', labelKey: 'Tên đối tác', valueKey: 'ID đối tác' },
        'ID_NhanVien': { table: 'User', labelKey: 'Họ và tên', valueKey: 'ID' }
    },
    'NhapChiTiet': { 'ID vật tư': { table: 'Vat_tu', labelKey: 'Tên vật tư', valueKey: 'ID vật tư' } },
    'XuatChiTiet': { 'ID vật tư': { table: 'Vat_tu', labelKey: 'Tên vật tư', valueKey: 'ID vật tư' } },
    'Phieunhap': { 'Tên kho': { table: 'DS_kho', labelKey: 'Tên kho', valueKey: 'ID kho' }, 'Người nhập': { table: 'User', labelKey: 'Họ và tên', valueKey: 'ID' } },
    'Phieuxuat': { 'Tên kho': { table: 'DS_kho', labelKey: 'Tên kho', valueKey: 'ID kho' }, 'Người xuất': { table: 'User', labelKey: 'Họ và tên', valueKey: 'ID' } },
    'Chiphi': { 'NguoiLap': { table: 'User', labelKey: 'Họ và tên', valueKey: 'ID' } },
    'Chiphichitiet': {
        'ID vật tư': { table: 'Vat_tu', labelKey: 'Tên vật tư', valueKey: 'ID vật tư' },
        'MaVatTu': { table: 'Vat_tu', labelKey: 'Tên vật tư', valueKey: 'ID vật tư' },
        'Mã vật tư': { table: 'Vat_tu', labelKey: 'Tên vật tư', valueKey: 'ID vật tư' },
        'Ten_kho': { table: 'DS_kho', labelKey: 'Tên kho', valueKey: 'ID kho' },
        'Tên kho': { table: 'DS_kho', labelKey: 'Tên kho', valueKey: 'ID kho' }
    },
    'Phieuchuyenkho': { 'KhoDi': { table: 'DS_kho', labelKey: 'Tên kho', valueKey: 'ID kho' }, 'KhoDen': { table: 'DS_kho', labelKey: 'Tên kho', valueKey: 'ID kho' }, 'NguoiChuyen': { table: 'User', labelKey: 'Họ và tên', valueKey: 'ID' }, 'NguoiNhan': { table: 'User', labelKey: 'Họ và tên', valueKey: 'ID' } },
    'Chuyenkhochitiet': { 'Kho nguồn': { table: 'DS_kho', labelKey: 'Tên kho', valueKey: 'ID kho' }, 'Kho đích': { table: 'DS_kho', labelKey: 'Tên kho', valueKey: 'ID kho' } },
    'Vat_tu': { 'Nhóm vật tư': { table: 'Nhomvattu', labelKey: 'Tên nhóm', valueKey: 'ID NVT' } }
};

// Kiểm tra cột tiền tệ để format phân tách hàng ngàn
function isMoneyField(k) {
    if (/^ID|GiaoDich|HeSo|Thang|Nam|Tháng|Năm/i.test(k)) return false;
    return /tien|tiền|luong|lương|don.?gia|dongia|đơn giá|giatri|giá trị|thanh.?tien|thành tiền|thuclanh|thực lãnh|thuc.?linh|phucap|phụ cấp|thuong|thưởng|thunhap|thu nhập|tamung|tạm ứng|baohiem|bảo hiểm|giamtru|giảm trừ|ngansach|ngân sách|sodu|số dư|tổng.*tien|tổng.*tiền|phatsinh|phát sinh|luong.*coban|lương.*cơ.*bản|tien.*tangca|tiền.*tăng.*ca|net.*salary|gross.*pay|so.*tien|số.*tiền/i.test(k);
}

// Kiểm tra cột số để format số thông thường
function isNumericField(k) {
    if (/^ID|Date|Ngay|GiaoDich|HeSo/i.test(k)) return false;
    return /so.?luong|số lượng|soluong|so_luong|gio|giờ|hour|so.?ngay|số ngày|songay|so_ngay|so.?gio|số giờ|sogio|so_gio|ty.?le|tỷ lệ|tile|ti_le|phan.?tram|phần trăm|phantram|phan_tram|he.?so|hệ số|heso|he_so/i.test(k);
}

// Ràng buộc input cho các trường đặc biệt
const FIELD_CONSTRAINTS = {
    'HeSoTangCa': { type: 'number', min: 1.1, max: 1.9, step: 0.1, placeholder: 'Từ 1.1 đến 1.9' },
    'LuongCoBan_Ngay': { type: 'number', min: 0, step: 1000, placeholder: 'VNĐ / ngày' }
};

// === HÀM ĐỊNH DẠNG SỐ LIỆU ===
function formatMoney(value, showCurrency = false) {
    if (value === null || value === undefined || value === '') return '';
    var num = parseFloat(value);
    if (isNaN(num)) return '';
    if (num === 0) return '0';
    return num.toLocaleString('vi-VN') + (showCurrency ? ' đ' : '');
}

function formatNumber(value, decimals = 0) {
    if (value === null || value === undefined || value === '') return '';
    var num = parseFloat(value);
    if (isNaN(num)) return '';
    if (num === 0) return '0';
    return decimals > 0 ? num.toFixed(decimals) : num.toLocaleString('vi-VN');
}

function formatPercentage(value, decimals = 1) {
    if (value === null || value === undefined || value === '') return '';
    var num = parseFloat(value);
    if (isNaN(num)) return '';
    return num.toFixed(decimals) + '%';
}

// === CHUYỂN ĐỔI SỐ THÀNH CHỮ ===
function numberToVietnameseText(number) {
    if (number === 0) return 'Không đồng';
    
    var ones = ['', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín'];
    var tens = ['', '', 'hai mươi', 'ba mười', 'bốn mười', 'năm mười', 'sáu mười', 'bảy mười', 'tám mười', 'chín mười'];
    var scales = ['', 'ngàn', 'triệu', 'tỷ', 'nghìn tỷ'];
    
    function convertHundreds(num) {
        var result = '';
        var hundred = Math.floor(num / 100);
        var remainder = num % 100;
        var ten = Math.floor(remainder / 10);
        var one = remainder % 10;
        
        if (hundred > 0) {
            result += ones[hundred] + ' trăm';
        }
        
        if (ten >= 2) {
            result += (result ? ' ' : '') + tens[ten];
            if (one > 0) {
                result += ' ' + ones[one];
            }
        } else if (ten === 1) {
            result += (result ? ' ' : '') + 'mười';
            if (one > 0) {
                result += ' ' + ones[one];
            }
        } else if (one > 0) {
            result += (result ? ' ' : '') + ones[one];
        }
        
        return result;
    }
    
    var isNegative = number < 0;
    number = Math.abs(number);
    
    var groups = [];
    while (number > 0) {
        groups.push(number % 1000);
        number = Math.floor(number / 1000);
    }
    
    var result = '';
    for (var i = groups.length - 1; i >= 0; i--) {
        if (groups[i] > 0) {
            var groupText = convertHundreds(groups[i]);
            if (i > 0) {
                groupText += ' ' + scales[i];
            }
            result += (result ? ' ' : '') + groupText;
        }
    }
    
    result = result.charAt(0).toUpperCase() + result.slice(1);
    result += ' đồng';
    
    if (isNegative) {
        result = 'Âm ' + result.toLowerCase();
    }
    
    return result;
}

function getDropdownForField(sheet, fieldName) {
    var cfg = (DROPDOWN_CONFIG[sheet] && DROPDOWN_CONFIG[sheet][fieldName]) || (DROPDOWN_CONFIG['*'] && DROPDOWN_CONFIG['*'][fieldName]);
    return cfg || null;
}

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
            var data = Array.isArray(result.data) ? result.data : [];
            localStorage.setItem(`cache_${table}`, JSON.stringify(data));
            return { status: 'success', data: data };
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
    "DS_kho": "DS_kho", "Vat_tu": "Vat_tu", "Tonkho": "Tonkho",
    "Chuyenkhochitiet": "Chuyenkhochitiet", "Nhomvattu": "Nhomvattu", "BangLuongThang": "BangLuongThang",
    "Notes": "Notes", "LichNhac": "LichNhac", "Chiphi": "Chiphi", "User": "User",
    "Chiphichitiet": "Chiphichitiet", "Luongphucap": "Luongphucap", "Doitac": "Doitac",
    "Drop": "Drop"
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

// --- HÀM GỌI SUPABASE (ĐĂNG NHẬP + CRUD) ---
async function callGAS(action, sheet, data = null, id = null) {
    if (action === 'loginUser') {
        const { data: user, error } = await supabaseClient.from('User').select('*').eq('ID', sheet).eq('App_pass', data).single();
        if (error) return { status: 'error', error: 'Sai ID hoặc mật khẩu' };
        const normalizedUser = {
            ...user,
            name: user['Họ và tên'] || user['ID'],
            role: user['Phân quyền'] || user['Chức vụ'] || 'Nhân viên'
        };
        return { status: 'success', user: normalizedUser };
    }
    // Lưu mới → insert Supabase
    if (action === 'saveData') {
        return callSupabase('insert', sheet, data, null);
    }
    // Cập nhật → lấy khóa chính từ dòng tại EDIT_INDEX rồi update
    if (action === 'updateData') {
        const rows = GLOBAL_DATA[sheet] || [];
        const row = rows[id];
        if (!row) return { status: 'error', message: 'Dòng không tồn tại' };
        const mappedTable = TABLE_MAP[sheet] || sheet;
        const pkField = PK_MAP[mappedTable] || 'ID';
        const pkValue = row[pkField];
        return callSupabase('update', sheet, data, pkValue);
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
            var data = Array.isArray(res.data) ? res.data : [];
            GLOBAL_DATA[sheet] = data;
            if (CURRENT_SHEET == sheet) renderTable(data);
        }
    });
}

function getInitialData() {
    showLoading(true);
    var promises = [];
    MENU_STRUCTURE.forEach(group => {
        group.items.forEach(m => {
            promises.push(callSupabase('read', m.sheet).then(res => {
                if (res.status == 'success') GLOBAL_DATA[m.sheet] = Array.isArray(res.data) ? res.data : [];
            }));
        });
    });

    const extraSheets = ['User', 'DS_kho', 'NhapChiTiet', 'XuatChiTiet', 'Chiphichitiet', 'Vat_tu', 'Nhomvattu', 'Drop'];
    var allSheets = extraSheets.slice();
    Object.keys(RELATION_CONFIG).forEach(function (parent) {
        var child = RELATION_CONFIG[parent].child;
        if (child && allSheets.indexOf(child) === -1) allSheets.push(child);
    });
    allSheets.forEach(s => {
        promises.push(callSupabase('read', s).then(res => {
            if (res.status == 'success') GLOBAL_DATA[s] = Array.isArray(res.data) ? res.data : [];
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
        } else {
            renderDashboard();
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

    // Dùng đúng key trong GLOBAL_DATA (trùng với sheet trong MENU_STRUCTURE / getInitialData)
    var stats = {
        user: (GLOBAL_DATA['User'] || []).length,
        vattu: (GLOBAL_DATA['Vat_tu'] || []).length,
        kho: (GLOBAL_DATA['DS_kho'] || []).length,
        phieu: (GLOBAL_DATA['Phieunhap'] || []).length + (GLOBAL_DATA['Phieuxuat'] || []).length
    };

    var kpiHTML = `
    <h5 class="fw-bold text-success mb-3"><i class="fas fa-chart-pie me-2"></i>TỔNG QUAN</h5>
    <div class="row g-3 mb-4 kpi-row">
        <div class="col-12 col-sm-6 col-xl-3 kpi-col">${kpiCard('Nhân sự', stats.user, 'fas fa-users', '#1E88E5', 'nhansu')}</div>
        <div class="col-12 col-sm-6 col-xl-3 kpi-col">${kpiCard('Vật tư', stats.vattu, 'fas fa-box', '#43A047', 'vattu')}</div>
        <div class="col-12 col-sm-6 col-xl-3 kpi-col">${kpiCard('Kho bãi', stats.kho, 'fas fa-warehouse', '#FB8C00', 'dskho')}</div>
        <div class="col-12 col-sm-6 col-xl-3 kpi-col">${kpiCard('Phiếu NX', stats.phieu, 'fas fa-file-invoice', '#E53935', 'phieunhap')}</div>
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

    // Xóa filter-bar cũ khi chuyển tab để tránh filter cũ dính lại
    var filterBarWrap = document.getElementById('filter-bar-wrap');
    if (filterBarWrap) filterBarWrap.innerHTML = '';
    var oldFilterBar = document.getElementById('filter-bar');
    if (oldFilterBar) oldFilterBar.remove();

    if (id === 'dashboard') {
        document.getElementById('tab-dashboard').classList.remove('d-none');
        document.getElementById('tab-generic').classList.add('d-none');
        renderDashboard();
    } else {
        var item = null;
        MENU_STRUCTURE.forEach(g => { var f = g.items.find(x => x.id == id); if (f) item = f; });

        if (item) {
            CURRENT_SHEET = item.sheet;
            CURRENT_MENU_ID = id;
            document.getElementById('tab-dashboard').classList.add('d-none');
            document.getElementById('tab-generic').classList.remove('d-none');
            document.getElementById('page-title').innerHTML = `<i class="${item.icon} me-2"></i> ${item.title}`;

            // Nút đặc biệt cho bảng lương
            var extraButtons = '';
            if (item.sheet === 'BangLuongThang') {
                extraButtons = `
                    <button class="btn btn-primary rounded-pill fw-bold shadow-sm px-3 me-2" onclick="openSalaryCalculator()">
                        <i class="fas fa-calculator me-1"></i> Tính lương
                    </button>
                    <button class="btn btn-success rounded-pill fw-bold shadow-sm px-3" onclick="openAddModal()">
                        <i class="fas fa-plus me-1"></i> Thêm mới
                    </button>
                `;
            } else if (item.sheet === 'Chiphi') {
                extraButtons = `
                    <button class="btn btn-primary rounded-pill fw-bold shadow-sm px-3 me-2" onclick="openQuickExpenseForm()">
                        <i class="fas fa-plus-circle me-1"></i> Nhập chi phí
                    </button>
                    <button class="btn btn-outline-secondary rounded-pill fw-bold shadow-sm px-3" onclick="openAddModal()">
                        <i class="fas fa-plus me-1"></i> Tạo phiếu thủ công
                    </button>
                `;
            } else {
                extraButtons = `<button class="btn btn-success rounded-pill fw-bold shadow-sm px-3" onclick="openAddModal()"><i class="fas fa-plus me-1"></i> Thêm mới</button>`;
            }

            document.getElementById('extra-btn-area').innerHTML = extraButtons;
            document.getElementById('filter-area').innerHTML = '';

            var data = GLOBAL_DATA[CURRENT_SHEET];
            if (data == null) {
                GLOBAL_DATA[CURRENT_SHEET] = [];
                renderTable([]);
                refreshSingleSheet(CURRENT_SHEET);
            } else {
                renderTable(Array.isArray(data) ? data : []);
            }
        }
    }
    if (window.innerWidth < 992 && sidebarBS) sidebarBS.hide();
}

// --- MỞ CÀI ĐẶT LƯƠNG CHO NHÂN VIÊN ---
window.openSalaryForEmployee = function (empId, empName) {
    CURRENT_SHEET = 'LichSuLuong';
    var allData = GLOBAL_DATA['LichSuLuong'] || [];
    var record = allData.find(r => String(r['ID_NhanVien']) === String(empId));
    if (record) {
        var idx = allData.indexOf(record);
        EDIT_INDEX = idx;
        document.getElementById('modalTitle').innerText = 'Cài đặt lương - ' + empName;
        buildForm('LichSuLuong', record);
        new bootstrap.Modal(document.getElementById('dataModal')).show();
    } else {
        EDIT_INDEX = -1;
        document.getElementById('modalTitle').innerText = 'Cài đặt lương - ' + empName;
        buildForm('LichSuLuong', null);
        new bootstrap.Modal(document.getElementById('dataModal')).show();
        var sel = document.querySelector('#dynamic-form [name="ID_NhanVien"]');
        if (sel) sel.value = empId;
    }
}

// --- RENDER BẢNG LƯƠNG VỚI GIAO DIỆN CẢI THIỆN ---
function renderPayrollTable(data) {
    data = Array.isArray(data) ? data : [];
    
    var table = document.getElementById('data-table');
    var tb = document.querySelector('#data-table tbody');
    var th = document.querySelector('#data-table thead');
    if (!tb || !th) return;

    // Thêm class đặc biệt cho bảng lương
    table.classList.add('payroll-table');
    
    // Thêm bộ lọc tháng/năm cho bảng lương
    var filterArea = document.getElementById('filter-area');
    if (filterArea && !document.getElementById('payroll-month-filter')) {
        var currentDate = new Date();
        var currentMonth = currentDate.getMonth() + 1;
        var currentYear = currentDate.getFullYear();

        filterArea.innerHTML = `
            <div class="d-flex gap-3 align-items-center flex-wrap">
                <div>
                    <label class="form-label small fw-bold text-muted mb-1">Tháng lương</label>
                    <select id="payroll-month-filter" class="form-select form-select-sm" onchange="applyPayrollFilter()" style="min-width: 100px;">
                        ${[1,2,3,4,5,6,7,8,9,10,11,12].map(m =>
                            `<option value="${m}" ${m === currentMonth ? 'selected' : ''}>Tháng ${m}</option>`
                        ).join('')}
                    </select>
                </div>
                <div>
                    <label class="form-label small fw-bold text-muted mb-1">Năm</label>
                    <select id="payroll-year-filter" class="form-select form-select-sm" onchange="applyPayrollFilter()" style="min-width: 100px;">
                        ${[2024, 2025, 2026, 2027].map(y =>
                            `<option value="${y}" ${y === currentYear ? 'selected' : ''}>${y}</option>`
                        ).join('')}
                    </select>
                </div>
                <div class="flex-grow-1">
                    <label class="form-label small fw-bold text-muted mb-1">Tìm nhân viên</label>
                    <input type="text" id="payroll-search" class="form-control form-control-sm" placeholder="Tìm theo mã hoặc tên..." oninput="applyPayrollFilter()">
                </div>
                <div>
                    <label class="form-label small fw-bold text-muted mb-1">&nbsp;</label>
                    <button class="btn btn-success btn-sm d-block me-2" onclick="calculateMonthlyPayroll()" title="Tính lương tháng">
                        <i class="fas fa-calculator me-1"></i> Tính lương tháng
                    </button>
                </div>
                <div>
                    <label class="form-label small fw-bold text-muted mb-1">&nbsp;</label>
                    <button class="btn btn-primary btn-sm d-block me-2" onclick="openIndividualPayrollCalc()" title="Tính lương cá nhân">
                        <i class="fas fa-user-check me-1"></i> Tính cá nhân
                    </button>
                </div>
                <div>
                    <label class="form-label small fw-bold text-muted mb-1">&nbsp;</label>
                    <button class="btn btn-success btn-sm d-block" onclick="exportPayrollToExcel()" title="Xuất Excel bảng lương">
                        <i class="fas fa-file-excel me-1"></i> Xuất Excel
                    </button>
                </div>
            </div>
        `;
    }

    // Lấy giá trị filter
    var filterMonth = document.getElementById('payroll-month-filter')?.value || currentMonth;
    var filterYear = document.getElementById('payroll-year-filter')?.value || currentYear;
    var searchText = (document.getElementById('payroll-search')?.value || '').toLowerCase().trim();

    // Lọc dữ liệu theo tháng/năm
    var filteredData = data.filter(r => {
        if (r['Delete'] === 'X') return false;
        
        var month = parseInt(r['Thang'] || r['Tháng']) || 0;
        var year = parseInt(r['Nam'] || r['Năm']) || 0;
        
        if (filterMonth && month !== parseInt(filterMonth)) return false;
        if (filterYear && year !== parseInt(filterYear)) return false;
        
        // Tìm kiếm theo tên hoặc mã nhân viên
        if (searchText) {
            var empName = (r['HoTen'] || r['Họ và tên'] || '').toLowerCase();
            var empId = (r['ID_NhanVien'] || r['MaNhanVien'] || '').toLowerCase();
            if (!empName.includes(searchText) && !empId.includes(searchText)) return false;
        }
        
        return true;
    });

    th.innerHTML = '';
    tb.innerHTML = '';

    // Tạo tiêu đề bảng
    var caption = table.querySelector('caption');
    if (!caption) {
        caption = document.createElement('caption');
        caption.style.captionSide = 'top';
        caption.className = 'text-center fw-bold text-success py-2';
        table.insertBefore(caption, table.firstChild);
    }
    caption.innerHTML = `<i class="fas fa-money-check-alt me-2"></i>BẢNG LƯƠNG THÁNG ${filterMonth}/${filterYear}`;

    // Tạo header với thanh trượt ngang
    var headerRow = document.createElement('tr');
    var headers = [
        { text: 'STT', width: '50px' },
        { text: 'Mã NV', width: '80px' },
        { text: 'Họ tên', width: '150px' },
        { text: 'Tháng', width: '60px' },
        { text: 'Năm', width: '60px' },
        { text: 'Giờ công', width: '80px' },
        { text: 'Giờ TC', width: '80px' },
        { text: 'Lương CB', width: '100px' },
        { text: 'Tiền TC', width: '100px' },
        { text: 'Phụ cấp', width: '90px' },
        { text: 'Thưởng', width: '90px' },
        { text: 'Tổng thu', width: '110px' },
        { text: 'Tạm ứng', width: '90px' },
        { text: 'Bảo hiểm', width: '90px' },
        { text: 'Giảm trừ', width: '90px' },
        { text: 'Thực lĩnh', width: '110px' },
        { text: 'Thao tác', width: '100px' }
    ];

    headers.forEach((h, index) => {
        var th = document.createElement('th');
        th.innerText = h.text;
        
        // Căn phải cho các cột số liệu (từ cột 5 trở đi)
        if (index >= 5 && index <= 15) {
            th.className = 'text-end bg-success text-white sticky-header';
        } else {
            th.className = 'text-center bg-success text-white sticky-header';
        }
        
        th.style.minWidth = h.width;
        th.style.fontSize = '11px';
        th.style.padding = '8px 4px';
        th.style.lineHeight = '1.2';
        headerRow.appendChild(th);
    });
    th.appendChild(headerRow);

    // Render dữ liệu với các dòng gần lại
    filteredData.forEach((record, index) => {
        var tr = document.createElement('tr');
        tr.className = 'payroll-row';
        tr.style.cursor = 'pointer';
        tr.onclick = function() { showPayrollDetail(record, filterMonth, filterYear); };
        
        // Hover effect
        tr.onmouseenter = function() { this.style.backgroundColor = '#f8f9fa'; };
        tr.onmouseleave = function() { this.style.backgroundColor = ''; };

        // Tính toán các giá trị lương
        var totalHours = parseFloat(record['TongGioCong']) || 0;
        var totalOvertime = parseFloat(record['TongGioTangCa']) || 0;
        var basicSalary = parseFloat(record['TongLuongCoBan']) || 0;
        var overtimePay = parseFloat(record['TongTienTangCa']) || 0;
        var allowance = parseFloat(record['TongPhuCap']) || 0;
        var bonus = parseFloat(record['TongThuong']) || 0;
        var totalIncome = basicSalary + overtimePay + allowance + bonus;
        var advance = parseFloat(record['TongTamUng']) || 0;
        var insurance = parseFloat(record['TongBaoHiem']) || 0;
        var deduction = parseFloat(record['TongGiamTru']) || 0;
        var netSalary = totalIncome - advance - insurance - deduction;

        var cells = [
            index + 1,
            record['ID_NhanVien'] || record['MaNhanVien'] || '',
            record['HoTen'] || record['Họ và tên'] || '',
            record['Thang'] || record['Tháng'] || '',
            record['Nam'] || record['Năm'] || '',
            formatNumber(totalHours, 1), // Giờ công với 1 số thập phân
            formatNumber(totalOvertime, 1), // Giờ TC với 1 số thập phân
            formatMoney(basicSalary), // Lương CB với phân tách hàng ngàn
            formatMoney(overtimePay), // Tiền TC với phân tách hàng ngàn
            formatMoney(allowance), // Phụ cấp với phân tách hàng ngàn
            formatMoney(bonus), // Thưởng với phân tách hàng ngàn
            formatMoney(totalIncome), // Tổng thu với phân tách hàng ngàn
            formatMoney(advance), // Tạm ứng với phân tách hàng ngàn
            formatMoney(insurance), // Bảo hiểm với phân tách hàng ngàn
            formatMoney(deduction), // Giảm trừ với phân tách hàng ngàn
            formatMoney(netSalary), // Thực lĩnh với phân tách hàng ngàn
            `<button class="btn btn-sm btn-outline-primary" onclick="event.stopPropagation(); showPayrollDetail(${JSON.stringify(record).replace(/"/g, '&quot;')}, ${filterMonth}, ${filterYear})">
                <i class="fas fa-eye"></i>
            </button>`
        ];

        cells.forEach((cellData, cellIndex) => {
            var td = document.createElement('td');
            if (cellIndex === 0 || cellIndex === 3 || cellIndex === 4) {
                td.className = 'text-center'; // STT, Tháng, Năm căn giữa
            } else if (cellIndex === 1 || cellIndex === 2) {
                td.className = 'text-start'; // Mã NV và Tên căn trái
            } else if (cellIndex >= 5 && cellIndex <= 15) {
                td.className = 'text-end fw-bold payroll-number-cell'; // Tất cả số liệu căn phải
                // Màu sắc phân biệt
                if (cellIndex === 11) td.style.color = '#28a745'; // Tổng thu - xanh
                if (cellIndex === 15) td.style.color = '#6f42c1'; // Thực lĩnh - tím
                if (cellIndex >= 12 && cellIndex <= 14) td.style.color = '#dc3545'; // Các khoản trừ - đỏ
                if (cellIndex >= 7 && cellIndex <= 10) td.style.color = '#28a745'; // Các khoản thu - xanh
            } else if (cellIndex === 16) {
                td.className = 'text-center';
                td.innerHTML = cellData;
                tr.appendChild(td);
                return;
            }
            
            td.innerHTML = cellData;
            td.style.fontSize = '11px';
            td.style.padding = '6px 4px';
            td.style.lineHeight = '1.3';
            tr.appendChild(td);
        });

        tb.appendChild(tr);
    });

    // Thêm dòng tổng cộng
    if (filteredData.length > 0) {
        var totalRow = document.createElement('tr');
        totalRow.className = 'table-warning fw-bold';
        
        var totalBasic = filteredData.reduce((sum, r) => sum + (parseFloat(r['TongLuongCoBan']) || 0), 0);
        var totalOvertime = filteredData.reduce((sum, r) => sum + (parseFloat(r['TongTienTangCa']) || 0), 0);
        var totalAllowance = filteredData.reduce((sum, r) => sum + (parseFloat(r['TongPhuCap']) || 0), 0);
        var totalBonus = filteredData.reduce((sum, r) => sum + (parseFloat(r['TongThuong']) || 0), 0);
        var totalIncome = totalBasic + totalOvertime + totalAllowance + totalBonus;
        var totalAdvance = filteredData.reduce((sum, r) => sum + (parseFloat(r['TongTamUng']) || 0), 0);
        var totalInsurance = filteredData.reduce((sum, r) => sum + (parseFloat(r['TongBaoHiem']) || 0), 0);
        var totalDeduction = filteredData.reduce((sum, r) => sum + (parseFloat(r['TongGiamTru']) || 0), 0);
        var totalNet = totalIncome - totalAdvance - totalInsurance - totalDeduction;

        totalRow.innerHTML = `
            <td class="text-center">-</td>
            <td class="text-center"><strong>TỔNG CỘNG</strong></td>
            <td class="text-center">${filteredData.length} người</td>
            <td class="text-center">-</td>
            <td class="text-center">-</td>
            <td class="text-center">-</td>
            <td class="text-center">-</td>
            <td class="text-end payroll-total-cell"><strong>${formatMoney(totalBasic)}</strong></td>
            <td class="text-end payroll-total-cell"><strong>${formatMoney(totalOvertime)}</strong></td>
            <td class="text-end payroll-total-cell"><strong>${formatMoney(totalAllowance)}</strong></td>
            <td class="text-end payroll-total-cell"><strong>${formatMoney(totalBonus)}</strong></td>
            <td class="text-end payroll-total-cell text-success"><strong>${formatMoney(totalIncome)}</strong></td>
            <td class="text-end payroll-total-cell"><strong>${formatMoney(totalAdvance)}</strong></td>
            <td class="text-end payroll-total-cell"><strong>${formatMoney(totalInsurance)}</strong></td>
            <td class="text-end payroll-total-cell"><strong>${formatMoney(totalDeduction)}</strong></td>
            <td class="text-end payroll-total-cell text-primary"><strong>${formatMoney(totalNet)}</strong></td>
            <td class="text-center">-</td>
        `;
        
        tb.appendChild(totalRow);
    } else {
        tb.innerHTML = '<tr><td colspan="17" class="text-center p-4 text-muted">Không có dữ liệu lương cho tháng này</td></tr>';
    }
}

// --- RENDER BẢNG CHẤM CÔNG DẠNG PIVOT ---
function renderAttendanceTable(data) {
    data = Array.isArray(data) ? data : [];

    var table = document.getElementById('data-table');
    var tb = document.querySelector('#data-table tbody');
    var th = document.querySelector('#data-table thead');
    if (!tb || !th) return;

    // Thêm class đặc biệt cho bảng chấm công
    table.classList.add('attendance-table');

    // Thêm bộ lọc tháng/năm cho bảng chấm công
    var filterArea = document.getElementById('filter-area');
    if (filterArea && !document.getElementById('attendance-month-filter')) {
        var currentDate = new Date();
        var currentMonth = currentDate.getMonth() + 1;
        var currentYear = currentDate.getFullYear();

        filterArea.innerHTML = `
            <div class="d-flex gap-3 align-items-center flex-wrap">
                <div>
                    <label class="form-label small fw-bold text-muted mb-1">Tháng</label>
                    <select id="attendance-month-filter" class="form-select form-select-sm" onchange="applyFilterInstant()" style="min-width: 100px;">
                        ${[1,2,3,4,5,6,7,8,9,10,11,12].map(m =>
                            `<option value="${m}" ${m === currentMonth ? 'selected' : ''}>Tháng ${m}</option>`
                        ).join('')}
                    </select>
                </div>
                <div>
                    <label class="form-label small fw-bold text-muted mb-1">Năm</label>
                    <select id="attendance-year-filter" class="form-select form-select-sm" onchange="applyFilterInstant()" style="min-width: 100px;">
                        ${[2024, 2025, 2026, 2027].map(y =>
                            `<option value="${y}" ${y === currentYear ? 'selected' : ''}>${y}</option>`
                        ).join('')}
                    </select>
                </div>
                <div class="flex-grow-1">
                    <label class="form-label small fw-bold text-muted mb-1">Tìm nhân viên</label>
                    <input type="text" id="attendance-search" class="form-control form-control-sm" placeholder="Tìm theo mã hoặc tên..." oninput="applyFilterInstant()">
                </div>
                <div>
                    <label class="form-label small fw-bold text-muted mb-1">&nbsp;</label>
                    <button class="btn btn-success btn-sm d-block" onclick="exportAttendanceToExcel()" title="Xuất Excel">
                        <i class="fas fa-file-excel me-1"></i> Xuất Excel
                    </button>
                </div>
            </div>
        `;
    }

    // Lấy giá trị filter
    var filterMonth = document.getElementById('attendance-month-filter')?.value;
    var filterYear = document.getElementById('attendance-year-filter')?.value;
    var searchText = (document.getElementById('attendance-search')?.value || '').toLowerCase().trim();

    th.innerHTML = '';
    tb.innerHTML = '';

    // Lọc dữ liệu chưa xóa và theo tháng/năm
    var activeData = data.filter(r => {
        if (r['Delete'] === 'X') return false;

        var dateStr = r['Ngày'] || r['Ngay'];
        if (!dateStr) return false;

        var date = new Date(dateStr);
        var month = date.getMonth() + 1;
        var year = date.getFullYear();

        if (filterMonth && month !== parseInt(filterMonth)) return false;
        if (filterYear && year !== parseInt(filterYear)) return false;

        return true;
    });

    if (activeData.length === 0) {
        th.innerHTML = '<tr><th class="bg-success text-white">Thông tin</th></tr>';
        tb.innerHTML = '<tr><td class="text-center p-5 text-muted fst-italic">Chưa có dữ liệu chấm công</td></tr>';
        return;
    }

    // Group data by employee
    var employeeMap = {};
    var userData = GLOBAL_DATA['User'] || [];

    activeData.forEach(record => {
        var empId = record['ID_NhanVien'] || record['Mã nhân viên'];
        if (!empId) return;

        // Tìm tên nhân viên từ bảng User
        var empInfo = userData.find(u => u['ID'] === empId);
        var empName = empInfo ? (empInfo['Họ và tên'] || empInfo['HoTen']) : empId;

        if (!employeeMap[empId]) {
            employeeMap[empId] = {
                id: empId,
                name: empName,
                days: {},
                totalHours: 0,
                totalOvertime: 0
            };
        }

        // Lấy ngày từ record
        var dateStr = record['Ngày'] || record['Ngay'];
        if (!dateStr) return;

        var date = new Date(dateStr);
        var day = date.getDate();

        var hours = parseFloat(record['SoGioCong']) || 0;
        var overtime = parseFloat(record['SoGioTangCa']) || 0;

        employeeMap[empId].days[day] = {
            hours: hours,
            overtime: overtime
        };

        employeeMap[empId].totalHours += hours;
        employeeMap[empId].totalOvertime += overtime;
    });

    // Tạo caption cho bảng
    var caption = table.querySelector('caption');
    if (!caption) {
        caption = document.createElement('caption');
        caption.style.captionSide = 'top';
        caption.className = 'text-center fw-bold text-success py-2';
        table.insertBefore(caption, table.firstChild);
    }
    var monthName = filterMonth ? `tháng ${filterMonth}` : 'tất cả tháng';
    var yearName = filterYear || 'tất cả năm';
    caption.innerHTML = `<i class="fas fa-calendar-alt me-2"></i>BẢNG CHẤM CÔNG ${monthName.toUpperCase()}/${yearName}`;

    // Tính số ngày trong tháng (nếu có chọn tháng)
    var daysInMonth = 31;
    if (filterMonth && filterYear) {
        daysInMonth = new Date(filterYear, filterMonth, 0).getDate();
    }

    // Tạo header: Mã NV | Tên NV | Tổng giờ chính | Tổng giờ TC | 1 | 2 | ... | ngày cuối tháng
    var trH = document.createElement('tr');
    var headers = ['Mã NV', 'Tên nhân viên', 'Tổng giờ chính', 'Tổng giờ TC'];

    headers.forEach(h => {
        var thEl = document.createElement('th');
        thEl.innerText = h;
        thEl.className = 'text-center';
        if (h === 'Mã NV') thEl.style.minWidth = '80px';
        if (h === 'Tên nhân viên') thEl.style.minWidth = '150px';
        if (h === 'Tổng giờ chính') thEl.style.minWidth = '100px';
        if (h === 'Tổng giờ TC') thEl.style.minWidth = '100px';
        trH.appendChild(thEl);
    });

    // Thêm cột cho các ngày trong tháng
    for (var d = 1; d <= daysInMonth; d++) {
        var thDay = document.createElement('th');
        thDay.innerText = d;
        thDay.className = 'text-center';
        thDay.style.minWidth = '60px';
        thDay.style.fontSize = '11px';
        trH.appendChild(thDay);
    }

    th.appendChild(trH);

    // Tạo các row cho từng nhân viên (có lọc theo search)
    Object.values(employeeMap).forEach(emp => {
        // Lọc theo search text
        if (searchText) {
            var empText = (emp.id + ' ' + emp.name).toLowerCase();
            if (!empText.includes(searchText)) return;
        }

        var tr = document.createElement('tr');
        tr.style.cursor = 'pointer';
        tr.onclick = function() { showEmployeeSalaryDetail(emp, filterMonth, filterYear); };

        // Mã NV
        var tdId = document.createElement('td');
        tdId.innerText = emp.id;
        tdId.className = 'text-center fw-bold';
        tr.appendChild(tdId);

        // Tên NV
        var tdName = document.createElement('td');
        tdName.innerText = emp.name;
        tdName.className = 'text-start';
        tr.appendChild(tdName);

        // Tổng giờ chính
        var tdTotalRegular = document.createElement('td');
        tdTotalRegular.innerText = emp.totalHours.toFixed(1);
        tdTotalRegular.className = 'text-center fw-bold text-success';
        tr.appendChild(tdTotalRegular);

        // Tổng giờ tăng ca
        var tdTotalOvertime = document.createElement('td');
        tdTotalOvertime.innerText = emp.totalOvertime.toFixed(1);
        tdTotalOvertime.className = 'text-center fw-bold text-danger';
        tr.appendChild(tdTotalOvertime);

        // Các ngày trong tháng
        for (var day = 1; day <= daysInMonth; day++) {
            var tdDay = document.createElement('td');
            tdDay.className = 'text-center';
            tdDay.style.fontSize = '11px';

            if (emp.days[day]) {
                var dayData = emp.days[day];
                var display = dayData.hours > 0 ? dayData.hours.toFixed(1) : '';

                if (dayData.overtime > 0) {
                    display += '<sup style="color: #e53935; font-weight: bold; margin-left: 2px; font-size: 9px;">' + dayData.overtime.toFixed(1) + '</sup>';
                }

                tdDay.innerHTML = display;
                if (dayData.hours > 0) {
                    tdDay.style.backgroundColor = '#e8f5e9';
                    tdDay.style.fontWeight = '600';
                }
            } else {
                tdDay.innerHTML = '-';
                tdDay.style.color = '#ccc';
                tdDay.style.fontSize = '10px';
            }

            tr.appendChild(tdDay);
        }

        tb.appendChild(tr);
    });
}

// --- HIỂN THỊ CHI TIẾT LƯƠNG NHÂN VIÊN ---
window.showEmployeeSalaryDetail = function(emp, month, year) {
    if (!emp) return;

    // Tính toán các thông tin lương
    var luongCoBan = emp.totalHours * 45000; // Giả định lương 45k/giờ
    var tienTangCa = emp.totalOvertime * 45000 * 1.5; // Tăng ca x1.5
    var phuCap = 1750000; // Phụ cấp cố định
    var thuong = 0;
    var tongThu = luongCoBan + tienTangCa + phuCap + thuong;

    var tamUng = 0;
    var baoHiem = 0;
    var giamTru = 0;
    var tongGiam = tamUng + baoHiem + giamTru;
    var thucLinh = tongThu - tongGiam;

    // Tìm dữ liệu lương từ bảng BangLuongThang nếu có
    var bangLuongData = GLOBAL_DATA['BangLuongThang'] || [];
    var salaryRecord = bangLuongData.find(s =>
        s['ID_NhanVien'] === emp.id &&
        parseInt(s['Thang'] || s['Tháng']) === parseInt(month) &&
        parseInt(s['Nam'] || s['Năm']) === parseInt(year)
    );

    if (salaryRecord) {
        luongCoBan = parseFloat(salaryRecord['TongLuongCoBan']) || luongCoBan;
        tienTangCa = parseFloat(salaryRecord['TongTienTangCa']) || tienTangCa;
        phuCap = parseFloat(salaryRecord['TongPhuCap']) || phuCap;
        thuong = parseFloat(salaryRecord['TongThuong']) || 0;
        tongThu = parseFloat(salaryRecord['TongThuNhap']) || tongThu;
        tamUng = parseFloat(salaryRecord['TongTamUng']) || 0;
        baoHiem = parseFloat(salaryRecord['TienBaoHiem']) || 0;
        giamTru = parseFloat(salaryRecord['GiamTruKhac']) || 0;
        tongGiam = parseFloat(salaryRecord['TongGiamTru']) || tongGiam;
        thucLinh = parseFloat(salaryRecord['Thuclanh'] || salaryRecord['Thuc_linh']) || thucLinh;
    }

    var kyLuong = `từ 16/${month} -> hết 20/${parseInt(month) + 1}/${year}`;
    var soNgayCong = Math.round(emp.totalHours / 8);

    var numberToWords = function(num) {
        return 'Hai mươi ba triệu hai trăm tám mươi lăm nghìn ba trăm mười tám đồng';
    };

    var html = `
        <div style="background: #f8f9fa; padding: 20px; border-radius: 12px; font-size: 13px;">
            <div style="background: white; padding: 15px; border-radius: 8px; margin-bottom: 15px;">
                <h6 class="text-success fw-bold mb-3" style="font-size: 14px;">
                    <i class="fas fa-user-circle me-2"></i>Lương nhân viên: ${emp.name}
                </h6>
                <div class="row g-2">
                    <div class="col-6"><strong>Kỳ lương:</strong></div>
                    <div class="col-6 text-end">${kyLuong}</div>
                    <div class="col-6"><strong>Giờ công:</strong></div>
                    <div class="col-6 text-end text-success fw-bold">${emp.totalHours.toFixed(1)} giờ (${soNgayCong} ngày)</div>
                    <div class="col-6"><strong>Tăng ca:</strong></div>
                    <div class="col-6 text-end text-danger fw-bold">${emp.totalOvertime.toFixed(1)} giờ</div>
                </div>
            </div>

            <div style="background: white; padding: 15px; border-radius: 8px; margin-bottom: 10px;">
                <div class="row g-2">
                    <div class="col-7"><strong>Lương cơ bản:</strong></div>
                    <div class="col-5 text-end">${luongCoBan.toLocaleString('vi-VN')} đ</div>
                    <div class="col-7"><strong>Tiền tăng ca:</strong></div>
                    <div class="col-5 text-end">${tienTangCa.toLocaleString('vi-VN')} đ</div>
                    <div class="col-7"><strong>Phụ cấp:</strong></div>
                    <div class="col-5 text-end">${phuCap.toLocaleString('vi-VN')} đ</div>
                    <div class="col-7"><strong>Thưởng:</strong></div>
                    <div class="col-5 text-end">${thuong.toLocaleString('vi-VN')} đ</div>
                    <div class="col-12"><hr class="my-2"></div>
                    <div class="col-7"><strong class="text-success">TỔNG THU:</strong></div>
                    <div class="col-5 text-end text-success fw-bold" style="font-size: 15px;">${tongThu.toLocaleString('vi-VN')} đ</div>
                </div>
            </div>

            <div style="background: white; padding: 15px; border-radius: 8px; margin-bottom: 10px;">
                <div class="row g-2">
                    <div class="col-7"><strong>Tạm ứng:</strong></div>
                    <div class="col-5 text-end">${tamUng.toLocaleString('vi-VN')} đ</div>
                    <div class="col-7"><strong>Bảo hiểm:</strong></div>
                    <div class="col-5 text-end">${baoHiem.toLocaleString('vi-VN')} đ</div>
                    <div class="col-7"><strong>Giảm trừ:</strong></div>
                    <div class="col-5 text-end">${giamTru.toLocaleString('vi-VN')} đ</div>
                    <div class="col-12"><hr class="my-2"></div>
                    <div class="col-7"><strong class="text-danger">TỔNG GIẢM:</strong></div>
                    <div class="col-5 text-end text-danger fw-bold">${tongGiam.toLocaleString('vi-VN')} đ</div>
                </div>
            </div>

            <div style="background: linear-gradient(135deg, #2E7D32, #1B5E20); padding: 15px; border-radius: 8px; color: white; margin-bottom: 10px;">
                <div class="row g-2">
                    <div class="col-6"><strong style="font-size: 15px;">THỰC LĨNH:</strong></div>
                    <div class="col-6 text-end fw-bold" style="font-size: 18px;">${thucLinh.toLocaleString('vi-VN')} đ</div>
                </div>
            </div>

            <div style="background: white; padding: 15px; border-radius: 8px;">
                <div class="mb-2"><strong>Bằng chữ:</strong></div>
                <div class="text-muted fst-italic">${numberToWords(thucLinh)}</div>
                <div class="mt-3"><strong>Ghi chú:</strong></div>
                <div class="text-muted">Tính từ 16/6 đến hết 20/7</div>
            </div>
        </div>
    `;

    Swal.fire({
        title: 'Chi tiết lương cá nhân',
        html: html,
        width: '500px',
        showCloseButton: true,
        showConfirmButton: false,
        customClass: {
            popup: 'rounded-4 shadow-lg',
            title: 'text-success'
        }
    });
}

// --- XUẤT EXCEL BẢNG LƯƠNG ĐƠN GIẢN ---
window.exportPayrollToExcel = function() {
    try {
        var monthEl = document.getElementById('payroll-month-filter');
        var yearEl = document.getElementById('payroll-year-filter');
        var month = monthEl ? monthEl.value : '';
        var year = yearEl ? yearEl.value : '';

        if (!month || !year) {
            Swal.fire({
                icon: 'warning',
                title: 'Thiếu thông tin',
                text: 'Vui lòng chọn tháng và năm để xuất Excel',
                confirmButtonColor: '#ffc107'
            });
            return;
        }

        Swal.fire({
            icon: 'info',
            title: 'Xuất Excel',
            text: 'Đang chuẩn bị file Excel bảng lương tháng ' + month + '/' + year + '...',
            timer: 2000,
            showConfirmButton: false
        });

        setTimeout(function() {
            try {
                var data = GLOBAL_DATA['BangLuongThang'] || [];
                
                // Lọc dữ liệu theo tháng/năm
                var filteredData = data.filter(function(r) {
                    if (r['Delete'] === 'X') return false;
                    var recordMonth = parseInt(r['Thang'] || r['Tháng']) || 0;
                    var recordYear = parseInt(r['Nam'] || r['Năm']) || 0;
                    return recordMonth === parseInt(month) && recordYear === parseInt(year);
                });

                if (filteredData.length === 0) {
                    Swal.fire({
                        icon: 'warning',
                        title: 'Không có dữ liệu',
                        text: 'Không có dữ liệu lương cho tháng/năm đã chọn',
                        confirmButtonColor: '#ffc107'
                    });
                    return;
                }

                // Tạo workbook
                var wb = XLSX.utils.book_new();
                var wsData = [];
                
                // Tiêu đề
                wsData.push(['BẢNG LƯƠNG THÁNG ' + month + '/' + year + ' - CÔNG TY CON ĐƯỜNG XANH']);
                wsData.push([]);
                
                // Header
                wsData.push([
                    'STT', 'Mã NV', 'Họ và tên', 'Tháng', 'Năm',
                    'Số giờ công', 'Số giờ TC', 'Lương cơ bản', 'Tiền tăng ca',
                    'Phụ cấp', 'Thưởng', 'Tổng thu', 'Tạm ứng', 'Thực lĩnh'
                ]);
                
                // Dữ liệu nhân viên
                filteredData.forEach(function(record, index) {
                    var basicSalary = parseFloat(record['TongLuongCoBan']) || 0;
                    var overtimePay = parseFloat(record['TongTienTangCa']) || 0;
                    var allowance = parseFloat(record['TongPhuCap']) || 0;
                    var bonus = parseFloat(record['TongThuong']) || 0;
                    var totalIncome = basicSalary + overtimePay + allowance + bonus;
                    var advance = parseFloat(record['TongTamUng']) || 0;
                    var netSalary = totalIncome - advance;
                    
                    wsData.push([
                        index + 1,
                        record['ID_NhanVien'] || '',
                        record['HoTen'] || record['Họ và tên'] || '',
                        record['Thang'] || record['Tháng'] || '',
                        record['Nam'] || record['Năm'] || '',
                        parseFloat(record['TongGioCong']) || 0,
                        parseFloat(record['TongGioTangCa']) || 0,
                        basicSalary,
                        overtimePay,
                        allowance,
                        bonus,
                        totalIncome,
                        advance,
                        netSalary
                    ]);
                });

                // Tạo worksheet và xuất
                var ws = XLSX.utils.aoa_to_sheet(wsData);
                XLSX.utils.book_append_sheet(wb, ws, 'Bảng lương');
                var fileName = 'BangLuong_Thang' + month + '_' + year + '.xlsx';
                XLSX.writeFile(wb, fileName);

                Swal.fire({
                    icon: 'success',
                    title: 'Thành công!',
                    text: 'File Excel bảng lương đã được tải xuống: ' + fileName,
                    confirmButtonColor: '#28a745'
                });

            } catch (exportError) {
                console.error('Lỗi xuất Excel:', exportError);
                Swal.fire({
                    icon: 'error',
                    title: 'Lỗi xuất Excel',
                    text: 'Có lỗi xảy ra: ' + (exportError.message || 'Unknown error'),
                    confirmButtonColor: '#dc3545'
                });
            }
        }, 100);
        
    } catch (error) {
        console.error('Lỗi general export:', error);
        Swal.fire({
            icon: 'error',
            title: 'Lỗi',
            text: 'Không thể thực hiện xuất Excel',
            confirmButtonColor: '#dc3545'
        });
    }
}

// --- POPUP CHI TIẾT LƯƠNG THEO ĐỊNH DẠNG MẪUI ---
window.showPayrollDetail = function(record, month, year) {
    if (!record) return;
    
    var empId = record['ID_NhanVien'] || record['MaNhanVien'] || '';
    var empName = record['HoTen'] || record['Họ và tên'] || empId;
    
    // Lấy thông tin từ các bảng liên quan
    var userData = GLOBAL_DATA['User'] || [];
    var salarySettingData = GLOBAL_DATA['LichSuLuong'] || [];
    var attendanceData = GLOBAL_DATA['Chamcong'] || [];
    
    var empInfo = userData.find(u => u['ID'] === empId);
    var salarySetting = salarySettingData.find(s => s['ID_NhanVien'] === empId);
    
    // Tính toán kỳ lương và số ngày công
    var kyLuongStart = record['KyLuong_TuNgay'] || `01/${month}/${year}`;
    var kyLuongEnd = record['KyLuong_DenNgay'] || `${new Date(year, month, 0).getDate()}/${month}/${year}`;
    
    // Đếm số ngày công thực tế từ chấm công
    var workDays = attendanceData.filter(a => {
        if (a['Delete'] === 'X' || a['ID_NhanVien'] !== empId) return false;
        var dateStr = a['Ngày'] || a['Ngay'];
        if (!dateStr) return false;
        var date = new Date(dateStr);
        return date.getMonth() + 1 === parseInt(month) && date.getFullYear() === parseInt(year);
    }).length;
    
    // Thông tin lương từ record hoặc tính toán
    var totalHours = parseFloat(record['TongGioCong']) || 0;
    var totalOvertime = parseFloat(record['TongGioTangCa']) || 0;
    var basicSalary = parseFloat(record['TongLuongCoBan']) || 0;
    var overtimePay = parseFloat(record['TongTienTangCa']) || 0;
    var allowance = parseFloat(record['TongPhuCap']) || 0;
    var bonus = parseFloat(record['TongThuong']) || 0;
    var totalIncome = basicSalary + overtimePay + allowance + bonus;
    
    var advance = parseFloat(record['TongTamUng']) || 0;
    var insurance = parseFloat(record['TongBaoHiem']) || 0;
    var previousDebt = parseFloat(record['NoThangTruoc']) || 0;
    var otherDeductions = parseFloat(record['TongGiamTru']) || 0;
    var totalDeductions = advance + insurance + previousDebt + otherDeductions;
    
    var netSalary = totalIncome - totalDeductions;
    var netSalaryText = numberToVietnameseText(Math.abs(netSalary));
    
    // Tạo HTML giống hình mẫu
    var html = `
        <div id="payroll-detail-content" style="font-size: 14px; text-align: left; background: white; padding: 20px; border-radius: 8px;">
            <div class="payroll-detail-table" style="border: 2px solid #000; border-collapse: collapse; width: 100%; font-family: 'Be Vietnam Pro', sans-serif;">
                <table style="width: 100%; border-collapse: collapse; background: white;">
                    <tr>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600; background: #f8f9fa; width: 35%;">
                            Tên nhân viên:
                        </td>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 700; color: #2E7D32;">
                            ${empName}
                        </td>
                    </tr>
                    <tr>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600; background: #f8f9fa;">
                            Kỳ lương:
                        </td>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600;">
                            từ ${kyLuongStart} -> hết ${kyLuongEnd}
                        </td>
                    </tr>
                    <tr>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600; background: #f8f9fa;">
                            Giờ công:
                        </td>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600;">
                            <span style="color: #dc3545; font-weight: 700; font-size: 16px;">${formatNumber(totalHours, 1)} giờ</span> 
                            <span style="color: #666;">(${workDays} ngày)</span>
                        </td>
                    </tr>
                    <tr>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600; background: #f8f9fa;">
                            Tăng ca:
                        </td>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600;">
                            ${formatNumber(totalOvertime, 1)} giờ
                        </td>
                    </tr>
                    <tr>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600; background: #f8f9fa;">
                            Lương cơ bản:
                        </td>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600; font-family: 'Courier New', monospace;">
                            ${formatMoney(basicSalary)} đ
                        </td>
                    </tr>
                    <tr>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600; background: #f8f9fa;">
                            Tiền tăng ca:
                        </td>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600; font-family: 'Courier New', monospace;">
                            ${formatMoney(overtimePay)} đ
                        </td>
                    </tr>
                    <tr>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600; background: #f8f9fa;">
                            Phụ cấp:
                        </td>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600; font-family: 'Courier New', monospace;">
                            ${formatMoney(allowance)} đ
                        </td>
                    </tr>
                    <tr>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600; background: #f8f9fa;">
                            Thưởng:
                        </td>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600; font-family: 'Courier New', monospace;">
                            ${formatMoney(bonus)} đ
                        </td>
                    </tr>
                    <tr style="background: #e8f5e8;">
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 700; font-size: 16px; background: #e8f5e8;">
                            TỔNG THU NHẬP:
                        </td>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 700; color: #28a745; font-size: 16px; font-family: 'Courier New', monospace; background: #e8f5e8;">
                            ${formatMoney(totalIncome)} đ
                        </td>
                    </tr>
                    <tr>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600; background: #f8f9fa;">
                            Tạm ứng:
                        </td>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600; font-family: 'Courier New', monospace; color: #dc3545;">
                            ${formatMoney(advance)} đ
                        </td>
                    </tr>
                    <tr>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600; background: #f8f9fa;">
                            Bảo hiểm:
                        </td>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600; font-family: 'Courier New', monospace;">
                            ${formatMoney(insurance)} đ
                        </td>
                    </tr>
                    <tr>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600; background: #f8f9fa;">
                            Giảm trừ:
                        </td>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600; font-family: 'Courier New', monospace;">
                            ${formatMoney(otherDeductions)} đ
                        </td>
                    </tr>
                    <tr style="background: #ffeaea;">
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 700; font-size: 16px; background: #ffeaea;">
                            TỔNG GIẢM:
                        </td>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 700; color: #dc3545; font-size: 16px; font-family: 'Courier New', monospace; background: #ffeaea;">
                            ${formatMoney(totalDeductions)} đ
                        </td>
                    </tr>
                    <tr style="background: ${netSalary >= 0 ? '#e3f2fd' : '#ffebee'};">
                        <td style="border: 2px solid #000; padding: 15px; font-weight: 700; font-size: 18px; background: ${netSalary >= 0 ? '#e3f2fd' : '#ffebee'};">
                            CÒN ĐƯỢC NHẬN:
                        </td>
                        <td style="border: 2px solid #000; padding: 15px; font-weight: 700; color: ${netSalary >= 0 ? '#1976d2' : '#d32f2f'}; font-size: 18px; font-family: 'Courier New', monospace; background: ${netSalary >= 0 ? '#e3f2fd' : '#ffebee'};">
                            ${netSalary >= 0 ? '' : '-'}${formatMoney(Math.abs(netSalary))} đ
                        </td>
                    </tr>
                    <tr>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600; background: #f8f9fa;">
                            Bằng chữ:
                        </td>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-style: italic; color: #666; line-height: 1.4;">
                            ${netSalary < 0 ? 'Âm ' : ''}${netSalaryText.toLowerCase()}
                        </td>
                    </tr>
                    <tr>
                        <td style="border: 1px solid #000; padding: 12px 15px; font-weight: 600; background: #f8f9fa;">
                            Ghi chú:
                        </td>
                        <td style="border: 1px solid #000; padding: 12px 15px;">
                            ${record['GhiChu'] || record['Ghi chú'] || 'Bảng tính lương'}
                        </td>
                    </tr>
                </table>
            </div>
            
            <div class="mt-4 d-flex gap-2 justify-content-center flex-wrap">
                <button class="btn btn-success btn-sm" onclick="exportPayrollImage('${empId}', ${month}, ${year})">
                    <i class="fas fa-image me-1"></i> Xuất hình ảnh
                </button>
                <button class="btn btn-primary btn-sm" onclick="exportPayrollPDF('${empId}', ${month}, ${year})">
                    <i class="fas fa-file-pdf me-1"></i> Xuất PDF
                </button>
                <button class="btn btn-warning btn-sm" onclick="Swal.close(); openIndividualPayrollCalc('${empId}')">
                    <i class="fas fa-calculator me-1"></i> Tính lại
                </button>
                <button class="btn btn-info btn-sm" onclick="Swal.close(); viewPayrollHistory('${empId}')">
                    <i class="fas fa-history me-1"></i> Lịch sử
                </button>
                <button class="btn btn-outline-secondary btn-sm" onclick="sendPayrollToEmployee('${empId}', ${month}, ${year})" data-requires-online="true">
                    <i class="fas fa-share me-1"></i> Gửi nhân viên
                </button>
            </div>
        </div>
    `;
    
    Swal.fire({
        title: '',
        html: html,
        width: '650px',
        showCloseButton: true,
        showConfirmButton: false,
        customClass: {
            popup: 'rounded-4 shadow-lg border-0',
            title: 'd-none'
        }
    });
}

// --- XUẤT HÌNH ẢNH BẢNG LƯƠNG ---
window.exportPayrollImage = function(empId, month, year) {
    Swal.fire({
        title: 'Xuất hình ảnh bảng lương',
        text: 'Đang tạo hình ảnh bảng lương...',
        allowOutsideClick: false,
        showConfirmButton: false,
        didOpen: () => {
            Swal.showLoading();
        }
    });
    
    // Delay để đảm bảo Swal hiển thị
    setTimeout(() => {
        try {
            // Lấy dữ liệu nhân viên và lương
            var userData = GLOBAL_DATA['User'] || [];
            var payrollData = GLOBAL_DATA['BangLuongThang'] || [];
            var record = payrollData.find(p => 
                p['ID_NhanVien'] === empId &&
                parseInt(p['Thang'] || p['Tháng']) === parseInt(month) &&
                parseInt(p['Nam'] || p['Năm']) === parseInt(year)
            );
            
            if (!record) {
                Swal.fire({
                    icon: 'warning',
                    title: 'Không tìm thấy dữ liệu',
                    text: 'Không tìm thấy bảng lương cho nhân viên này trong tháng đã chọn',
                    confirmButtonColor: '#ffc107'
                });
                return;
            }
            
            var empInfo = userData.find(u => u['ID'] === empId);
            var empName = record['HoTen'] || record['Họ và tên'] || empId;
            
            // Tính toán các giá trị
            var basicSalary = parseFloat(record['TongLuongCoBan']) || 0;
            var overtimePay = parseFloat(record['TongTienTangCa']) || 0;
            var allowance = parseFloat(record['TongPhuCap']) || 0;
            var bonus = parseFloat(record['TongThuong']) || 0;
            var totalIncome = basicSalary + overtimePay + allowance + bonus;
            
            var advance = parseFloat(record['TongTamUng']) || 0;
            var insurance = parseFloat(record['TongBaoHiem']) || 0;
            var deductions = parseFloat(record['TongGiamTru']) || 0;
            var totalDeductions = advance + insurance + deductions;
            
            var netSalary = totalIncome - totalDeductions;
            var netSalaryText = numberToVietnameseText(Math.abs(netSalary));
            
            // Tạo canvas để vẽ bảng lương
            var canvas = document.createElement('canvas');
            var ctx = canvas.getContext('2d');
            
            // Kích thước canvas
            canvas.width = 800;
            canvas.height = 600;
            
            // Nền trắng
            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, canvas.width, canvas.height);
            
            // Font chữ
            ctx.font = 'bold 24px Be Vietnam Pro, Arial';
            ctx.fillStyle = '#2E7D32';
            ctx.textAlign = 'center';
            ctx.fillText('BẢNG LƯƠNG CHI TIẾT', canvas.width / 2, 40);
            
            // Thông tin cơ bản
            ctx.font = '16px Be Vietnam Pro, Arial';
            ctx.textAlign = 'left';
            ctx.fillStyle = '#000000';
            
            var y = 80;
            var lineHeight = 30;
            var leftMargin = 50;
            
            var payrollInfo = [
                ['Tên nhân viên:', empName],
                ['Kỳ lương:', `từ 01/${month}/${year} -> hết 31/${month}/${year}`],
                ['Giờ công:', `${formatNumber(parseFloat(record['TongGioCong']) || 0, 1)} giờ`],
                ['Tăng ca:', `${formatNumber(parseFloat(record['TongGioTangCa']) || 0, 1)} giờ`],
                ['Lương cơ bản:', `${formatMoney(basicSalary)} đ`],
                ['Tiền tăng ca:', `${formatMoney(overtimePay)} đ`],
                ['Phụ cấp:', `${formatMoney(allowance)} đ`],
                ['Thưởng:', `${formatMoney(bonus)} đ`],
                ['TỔNG THU NHẬP:', `${formatMoney(totalIncome)} đ`],
                ['Tạm ứng:', `${formatMoney(advance)} đ`],
                ['Bảo hiểm:', `${formatMoney(insurance)} đ`],
                ['Giảm trừ:', `${formatMoney(deductions)} đ`],
                ['TỔNG GIẢM:', `${formatMoney(totalDeductions)} đ`],
                ['CÒN ĐƯỢC NHẬN:', `${netSalary >= 0 ? '' : '-'}${formatMoney(Math.abs(netSalary))} đ`],
                ['Bằng chữ:', netSalary < 0 ? 'Âm ' + netSalaryText.toLowerCase() : netSalaryText],
                ['Ghi chú:', record['GhiChu'] || 'Bảng tính lương']
            ];
            
            payrollInfo.forEach((info, index) => {
                var currentY = y + (index * lineHeight);
                
                // Vẽ nền cho các dòng quan trọng
                if (info[0].includes('TỔNG THU NHẬP') || info[0].includes('TỔNG GIẢM') || info[0].includes('CÒN ĐƯỢC NHẬN')) {
                    ctx.fillStyle = info[0].includes('THU NHẬP') ? '#e8f5e8' : info[0].includes('GIẢM') ? '#ffeaea' : '#e3f2fd';
                    ctx.fillRect(leftMargin - 10, currentY - 20, canvas.width - 100, lineHeight);
                }
                
                // Vẽ text
                ctx.fillStyle = '#000000';
                ctx.font = info[0].includes('TỔNG') || info[0].includes('CÒN ĐƯỢC') ? 'bold 16px Be Vietnam Pro, Arial' : '14px Be Vietnam Pro, Arial';
                ctx.fillText(info[0], leftMargin, currentY);
                
                ctx.fillStyle = info[0].includes('THU NHẬP') ? '#28a745' : 
                              info[0].includes('GIẢM') ? '#dc3545' : 
                              info[0].includes('CÒN ĐƯỢC') ? (netSalary >= 0 ? '#1976d2' : '#d32f2f') : '#000000';
                ctx.font = info[0].includes('TỔNG') || info[0].includes('CÒN ĐƯỢC') ? 'bold 16px Courier New, monospace' : '14px Courier New, monospace';
                ctx.textAlign = 'right';
                ctx.fillText(info[1], canvas.width - leftMargin, currentY);
                ctx.textAlign = 'left';
            });
            
            // Chuyển canvas thành blob và download
            canvas.toBlob(function(blob) {
                var url = URL.createObjectURL(blob);
                var a = document.createElement('a');
                a.href = url;
                a.download = `BangLuong_${empName.replace(/\s+/g, '_')}_${month}_${year}.png`;
                a.click();
                URL.revokeObjectURL(url);
                
                Swal.fire({
                    icon: 'success',
                    title: 'Thành công!',
                    text: `Đã xuất hình ảnh bảng lương cho ${empName}`,
                    confirmButtonColor: '#28a745'
                });
            }, 'image/png');
            
        } catch (error) {
            console.error('Lỗi xuất hình ảnh:', error);
            Swal.fire({
                icon: 'error',
                title: 'Lỗi xuất hình ảnh',
                text: error.message || 'Có lỗi xảy ra khi tạo hình ảnh',
                confirmButtonColor: '#dc3545'
            });
        }
    }, 500);
}

// --- XUẤT PDF BẢNG LƯƠNG ---
window.exportPayrollPDF = function(empId, month, year) {
    Swal.fire({
        icon: 'info',
        title: 'Xuất PDF',
        text: 'Chức năng xuất PDF đang được phát triển...',
        confirmButtonColor: '#2E7D32'
    });
}

// --- GỬI BẢNG LƯƠNG CHO NHÂN VIÊN ---
window.sendPayrollToEmployee = function(empId, month, year) {
    if (!window.isOnline) {
        Swal.fire({
            icon: 'warning',
            title: 'Cần kết nối mạng',
            text: 'Tính năng gửi bảng lương cần có kết nối internet',
            confirmButtonColor: '#ffc107'
        });
        return;
    }
    
    Swal.fire({
        icon: 'info',
        title: 'Gửi bảng lương',
        text: 'Chức năng gửi bảng lương qua email/SMS đang được phát triển...',
        confirmButtonColor: '#2E7D32'
    });
}

// --- TÍNH LƯƠNG THÁNG TỰ ĐỘNG ---
window.calculateMonthlyPayroll = function() {
    var month = document.getElementById('payroll-month-filter')?.value;
    var year = document.getElementById('payroll-year-filter')?.value;
    
    if (!month || !year) {
        Swal.fire({
            icon: 'warning',
            title: 'Thiếu thông tin',
            text: 'Vui lòng chọn tháng và năm để tính lương',
            confirmButtonColor: '#2E7D32'
        });
        return;
    }
    
    Swal.fire({
        title: 'Tính lương tháng',
        html: `
            <div style="text-align: left;">
                <div class="alert alert-info">
                    <i class="fas fa-info-circle me-2"></i>
                    <strong>Tính lương tháng ${month}/${year}</strong><br>
                    Hệ thống sẽ tự động tính lương cho tất cả nhân viên dựa trên:
                    <ul class="mt-2 mb-0">
                        <li>Dữ liệu chấm công trong tháng</li>
                        <li>Cài đặt lương cá nhân (LichSuLuong)</li>
                        <li>Tạm ứng và phụ cấp (GiaoDichLuong)</li>
                    </ul>
                </div>
                
                <div class="form-check">
                    <input class="form-check-input" type="checkbox" id="overwrite-existing">
                    <label class="form-check-label" for="overwrite-existing">
                        Ghi đè dữ liệu lương đã có (nếu tháng này đã tính)
                    </label>
                </div>
            </div>
        `,
        width: '600px',
        showCancelButton: true,
        confirmButtonText: '<i class="fas fa-calculator me-1"></i> Bắt đầu tính lương',
        cancelButtonText: 'Hủy',
        confirmButtonColor: '#2E7D32',
        preConfirm: () => {
            return {
                month: month,
                year: year,
                overwrite: document.getElementById('overwrite-existing').checked
            };
        }
    }).then((result) => {
        if (result.isConfirmed) {
            processMonthlyPayroll(result.value);
        }
    });
}

// --- XỬ LÝ TÍNH LƯƠNG THÁNG ---
async function processMonthlyPayroll(params) {
    showLoading(true, 'Đang tính lương tháng...');
    
    try {
        var userData = GLOBAL_DATA['User'] || [];
        var attendanceData = GLOBAL_DATA['Chamcong'] || [];
        var salarySettingData = GLOBAL_DATA['LichSuLuong'] || [];
        var transactionData = GLOBAL_DATA['GiaoDichLuong'] || [];
        var existingPayroll = GLOBAL_DATA['BangLuongThang'] || [];
        
        var activeUsers = userData.filter(u => 
            u['Delete'] !== 'X' && 
            (u['Tinh luong'] !== 'Không' && u['Tinh luong'] !== 'No')
        );
        
        var results = [];
        var processed = 0;
        
        for (let user of activeUsers) {
            var empId = user['ID'];
            var empName = user['Họ và tên'] || user['HoTen'] || empId;
            
            // Kiểm tra đã có lương tháng này chưa
            var existingRecord = existingPayroll.find(p => 
                p['ID_NhanVien'] === empId &&
                parseInt(p['Thang'] || p['Tháng']) === parseInt(params.month) &&
                parseInt(p['Nam'] || p['Năm']) === parseInt(params.year)
            );
            
            if (existingRecord && !params.overwrite) {
                continue; // Bỏ qua nếu đã có và không ghi đè
            }
            
            // Lấy dữ liệu chấm công trong tháng
            var monthlyAttendance = attendanceData.filter(a => {
                if (a['Delete'] === 'X' || a['ID_NhanVien'] !== empId) return false;
                
                var dateStr = a['Ngày'] || a['Ngay'];
                if (!dateStr) return false;
                
                var date = new Date(dateStr);
                return date.getMonth() + 1 === parseInt(params.month) && 
                       date.getFullYear() === parseInt(params.year);
            });
            
            // Tính tổng giờ
            var totalHours = monthlyAttendance.reduce((sum, a) => sum + (parseFloat(a['SoGioCong']) || 0), 0);
            var totalOvertime = monthlyAttendance.reduce((sum, a) => sum + (parseFloat(a['SoGioTangCa']) || 0), 0);
            
            // Lấy cài đặt lương
            var salarySetting = salarySettingData.find(s => s['ID_NhanVien'] === empId);
            var basicSalaryPerDay = parseFloat(salarySetting?.['LuongCoBan_Ngay']) || 45000;
            var overtimeRate = parseFloat(salarySetting?.['HeSoTangCa']) || 1.5;
            
            // Tính lương
            var basicSalary = totalHours * basicSalaryPerDay;
            var overtimePay = totalOvertime * basicSalaryPerDay * overtimeRate;
            
            // Lấy phụ cấp và thưởng từ GiaoDichLuong
            var monthlyTransactions = transactionData.filter(t => {
                if (t['Delete'] === 'X') return false;
                
                var transDate = new Date(t['NgayGiaoDich']);
                return transDate.getMonth() + 1 === parseInt(params.month) && 
                       transDate.getFullYear() === parseInt(params.year) &&
                       (t['ID_NhanVien'] === empId || t['MaNhanVien'] === empId);
            });
            
            var allowance = monthlyTransactions
                .filter(t => (t['LoaiGiaoDich'] || '').toLowerCase().includes('phụ cấp'))
                .reduce((sum, t) => sum + (parseFloat(t['SoTien']) || 0), 0);
                
            var bonus = monthlyTransactions
                .filter(t => (t['LoaiGiaoDich'] || '').toLowerCase().includes('thưởng'))
                .reduce((sum, t) => sum + (parseFloat(t['SoTien']) || 0), 0);
                
            var advance = monthlyTransactions
                .filter(t => (t['LoaiGiaoDich'] || '').toLowerCase().includes('tạm ứng'))
                .reduce((sum, t) => sum + (parseFloat(t['SoTien']) || 0), 0);
            
            // Tổng thu và thực lĩnh
            var totalIncome = basicSalary + overtimePay + allowance + bonus;
            var netSalary = totalIncome - advance;
            
            // Tạo/cập nhật bản ghi lương
            var payrollRecord = {
                'ID_BangLuong': existingRecord?.['ID_BangLuong'] || `${empId}-${params.year}${params.month.padStart(2, '0')}`,
                'ID_NhanVien': empId,
                'HoTen': empName,
                'Thang': parseInt(params.month),
                'Nam': parseInt(params.year),
                'NgayTinhLuong': new Date().toISOString().split('T')[0],
                'TongGioCong': totalHours,
                'TongGioTangCa': totalOvertime,
                'TongLuongCoBan': basicSalary,
                'TongTienTangCa': overtimePay,
                'TongPhuCap': allowance,
                'TongThuong': bonus,
                'TongThuNhap': totalIncome,
                'TongTamUng': advance,
                'TongBaoHiem': 0, // Có thể thêm logic tính bảo hiểm
                'TongGiamTru': 0,
                'ThucLinh': netSalary
            };
            
            // Lưu vào database
            var saveResult;
            if (existingRecord) {
                saveResult = await callSupabase('update', 'BangLuongThang', payrollRecord, existingRecord['ID_BangLuong']);
            } else {
                saveResult = await callSupabase('insert', 'BangLuongThang', payrollRecord);
            }
            
            if (saveResult.status === 'success') {
                results.push({
                    empId: empId,
                    empName: empName,
                    totalHours: totalHours,
                    totalOvertime: totalOvertime,
                    netSalary: netSalary,
                    status: 'success'
                });
            } else {
                results.push({
                    empId: empId,
                    empName: empName,
                    status: 'error',
                    message: saveResult.message
                });
            }
            
            processed++;
        }
        
        // Refresh dữ liệu
        await refreshSingleSheet('BangLuongThang');
        
        showLoading(false);
        
        // Hiển thị kết quả
        var successCount = results.filter(r => r.status === 'success').length;
        var errorCount = results.filter(r => r.status === 'error').length;
        
        var resultHtml = `
            <div style="text-align: left;">
                <div class="alert alert-success">
                    <h6><i class="fas fa-check-circle me-2"></i>Hoàn thành tính lương tháng ${params.month}/${params.year}</h6>
                    <div class="row">
                        <div class="col-6"><strong>Thành công:</strong> ${successCount} nhân viên</div>
                        <div class="col-6"><strong>Lỗi:</strong> ${errorCount} nhân viên</div>
                    </div>
                </div>
                
                ${successCount > 0 ? `
                <div class="mb-3">
                    <h6>Danh sách đã tính lương:</h6>
                    <div style="max-height: 200px; overflow-y: auto;">
                        <table class="table table-sm">
                            <thead><tr><th>Mã NV</th><th>Tên</th><th>Giờ công</th><th>Thực lĩnh</th></tr></thead>
                            <tbody>
                                ${results.filter(r => r.status === 'success').map(r => 
                                    `<tr>
                                        <td>${r.empId}</td>
                                        <td>${r.empName}</td>
                                        <td class="text-end number-display">${formatNumber(r.totalHours, 1)}h</td>
                                        <td class="text-end fw-bold money-display">${formatMoney(r.netSalary, true)}</td>
                                    </tr>`
                                ).join('')}
                            </tbody>
                        </table>
                    </div>
                </div>
                ` : ''}
                
                ${errorCount > 0 ? `
                <div class="alert alert-warning">
                    <h6>Một số lỗi xảy ra:</h6>
                    ${results.filter(r => r.status === 'error').map(r => 
                        `<div>- ${r.empName} (${r.empId}): ${r.message}</div>`
                    ).join('')}
                </div>
                ` : ''}
            </div>
        `;
        
        Swal.fire({
            icon: 'success',
            title: 'Tính lương hoàn thành',
            html: resultHtml,
            width: '700px',
            confirmButtonText: 'Đóng',
            confirmButtonColor: '#2E7D32'
        });
        
    } catch (error) {
        showLoading(false);
        console.error('Lỗi tính lương:', error);
        Swal.fire({
            icon: 'error',
            title: 'Lỗi tính lương',
            text: error.message || 'Có lỗi xảy ra khi tính lương tháng',
            confirmButtonColor: '#dc3545'
        });
    }
}

// --- TÍNH LƯƠNG CÁ NHÂN BẤT KỲ LÚC NÀO ---
window.openIndividualPayrollCalc = function(empId = '') {
    var userData = GLOBAL_DATA['User'] || [];
    var activeUsers = userData.filter(u => 
        u['Delete'] !== 'X' && 
        (u['Tinh luong'] !== 'Không' && u['Tinh luong'] !== 'No')
    );
    
    var userOptions = '<option value="">-- Chọn nhân viên --</option>';
    activeUsers.forEach(u => {
        var selected = u['ID'] === empId ? 'selected' : '';
        userOptions += `<option value="${u['ID']}" ${selected}>${u['ID']} - ${u['Họ và tên'] || u['HoTen'] || u['ID']}</option>`;
    });
    
    var currentDate = new Date();
    var currentMonth = currentDate.getMonth() + 1;
    var currentYear = currentDate.getFullYear();
    
    var html = `
        <div style="text-align: left;">
            <div class="row g-3">
                <div class="col-12">
                    <label class="form-label fw-bold">Nhân viên:</label>
                    <select id="individual-emp" class="form-select" required>
                        ${userOptions}
                    </select>
                </div>
                <div class="col-6">
                    <label class="form-label fw-bold">Tháng:</label>
                    <select id="individual-month" class="form-select">
                        ${[1,2,3,4,5,6,7,8,9,10,11,12].map(m =>
                            `<option value="${m}" ${m === currentMonth ? 'selected' : ''}>Tháng ${m}</option>`
                        ).join('')}
                    </select>
                </div>
                <div class="col-6">
                    <label class="form-label fw-bold">Năm:</label>
                    <select id="individual-year" class="form-select">
                        ${[2024, 2025, 2026, 2027].map(y =>
                            `<option value="${y}" ${y === currentYear ? 'selected' : ''}>${y}</option>`
                        ).join('')}
                    </select>
                </div>
            </div>
        </div>
    `;
    
    Swal.fire({
        title: '<i class="fas fa-calculator me-2"></i>Tính lương cá nhân',
        html: html,
        width: '500px',
        showCancelButton: true,
        confirmButtonText: '<i class="fas fa-check me-1"></i> Tính lương',
        cancelButtonText: 'Hủy',
        confirmButtonColor: '#2E7D32',
        preConfirm: () => {
            var empId = document.getElementById('individual-emp').value;
            var month = document.getElementById('individual-month').value;
            var year = document.getElementById('individual-year').value;
            
            if (!empId) {
                Swal.showValidationMessage('Vui lòng chọn nhân viên');
                return false;
            }
            
            return { empId, month, year };
        }
    }).then((result) => {
        if (result.isConfirmed) {
            processIndividualPayroll(result.value);
        }
    });
}

// --- XỬ LÝ TÍNH LƯƠNG CÁ NHÂN ---
async function processIndividualPayroll(params) {
    showLoading(true, 'Đang tính lương cá nhân...');
    
    try {
        // Logic tương tự processMonthlyPayroll nhưng chỉ cho 1 người
        var userData = GLOBAL_DATA['User'] || [];
        var attendanceData = GLOBAL_DATA['Chamcong'] || [];
        var salarySettingData = GLOBAL_DATA['LichSuLuong'] || [];
        var transactionData = GLOBAL_DATA['GiaoDichLuong'] || [];
        
        var user = userData.find(u => u['ID'] === params.empId);
        if (!user) throw new Error('Không tìm thấy nhân viên');
        
        var empName = user['Họ và tên'] || user['HoTen'] || params.empId;
        
        // Tính toán giống processMonthlyPayroll
        var monthlyAttendance = attendanceData.filter(a => {
            if (a['Delete'] === 'X' || a['ID_NhanVien'] !== params.empId) return false;
            
            var dateStr = a['Ngày'] || a['Ngay'];
            if (!dateStr) return false;
            
            var date = new Date(dateStr);
            return date.getMonth() + 1 === parseInt(params.month) && 
                   date.getFullYear() === parseInt(params.year);
        });
        
        var totalHours = monthlyAttendance.reduce((sum, a) => sum + (parseFloat(a['SoGioCong']) || 0), 0);
        var totalOvertime = monthlyAttendance.reduce((sum, a) => sum + (parseFloat(a['SoGioTangCa']) || 0), 0);
        
        var salarySetting = salarySettingData.find(s => s['ID_NhanVien'] === params.empId);
        var basicSalaryPerDay = parseFloat(salarySetting?.['LuongCoBan_Ngay']) || 45000;
        var overtimeRate = parseFloat(salarySetting?.['HeSoTangCa']) || 1.5;
        
        var basicSalary = totalHours * basicSalaryPerDay;
        var overtimePay = totalOvertime * basicSalaryPerDay * overtimeRate;
        var totalIncome = basicSalary + overtimePay;
        
        showLoading(false);
        
        // Hiển thị kết quả tính lương
        var resultHtml = `
            <div style="text-align: left;">
                <div class="card border-primary">
                    <div class="card-header bg-primary text-white">
                        <h6 class="mb-0"><i class="fas fa-calculator me-2"></i>${empName} - Tháng ${params.month}/${params.year}</h6>
                    </div>
                    <div class="card-body">
                        <div class="row g-2">
                            <div class="col-6"><strong>Tổng giờ công:</strong></div>
                            <div class="col-6 text-end">${totalHours.toFixed(1)} giờ</div>
                            
                            <div class="col-6"><strong>Tổng giờ tăng ca:</strong></div>
                            <div class="col-6 text-end">${totalOvertime.toFixed(1)} giờ</div>
                            
                            <div class="col-6"><strong>Lương CB/ngày:</strong></div>
                            <div class="col-6 text-end money-display">${formatMoney(basicSalaryPerDay, true)}</div>
                            
                            <div class="col-6"><strong>Hệ số tăng ca:</strong></div>
                            <div class="col-6 text-end number-display">x${formatNumber(overtimeRate, 1)}</div>
                            
                            <hr class="my-2">
                            
                            <div class="col-6"><strong>Lương cơ bản:</strong></div>
                            <div class="col-6 text-end text-success fw-bold money-display">${formatMoney(basicSalary, true)}</div>
                            
                            <div class="col-6"><strong>Tiền tăng ca:</strong></div>
                            <div class="col-6 text-end text-info fw-bold money-display">${formatMoney(overtimePay, true)}</div>
                            
                            <hr class="my-2">
                            
                            <div class="col-6"><strong style="font-size: 16px;">TỔNG LƯƠNG:</strong></div>
                            <div class="col-6 text-end fw-bold text-primary money-display" style="font-size: 18px;">${formatMoney(totalIncome, true)}</div>
                        </div>
                        
                        <div class="mt-3 text-center">
                            <button class="btn btn-success btn-sm me-2" onclick="Swal.close(); setTimeout(() => saveIndividualPayroll(${JSON.stringify(params).replace(/"/g, '&quot;')}, ${totalIncome}), 200);">
                                <i class="fas fa-save me-1"></i> Lưu vào bảng lương
                            </button>
                            <button class="btn btn-outline-primary btn-sm" onclick="Swal.close(); setTimeout(() => openIndividualPayrollCalc('${params.empId}'), 200);">
                                <i class="fas fa-redo me-1"></i> Tính lại
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        `;
        
        Swal.fire({
            title: 'Kết quả tính lương',
            html: resultHtml,
            width: '500px',
            showConfirmButton: false,
            showCloseButton: true
        });
        
    } catch (error) {
        showLoading(false);
        console.error('Lỗi tính lương cá nhân:', error);
        Swal.fire({
            icon: 'error',
            title: 'Lỗi tính lương',
            text: error.message || 'Có lỗi xảy ra khi tính lương',
            confirmButtonColor: '#dc3545'
        });
    }
}

// --- LƯU LƯƠNG CÁ NHÂN ---
window.saveIndividualPayroll = async function(params, totalIncome) {
    try {
        showLoading(true, 'Đang lưu lương...');
        
        var userData = GLOBAL_DATA['User'] || [];
        var user = userData.find(u => u['ID'] === params.empId);
        var empName = user?.['Họ và tên'] || user?.['HoTen'] || params.empId;
        
        var payrollRecord = {
            'ID_BangLuong': `${params.empId}-${params.year}${params.month.padStart(2, '0')}`,
            'ID_NhanVien': params.empId,
            'HoTen': empName,
            'Thang': parseInt(params.month),
            'Nam': parseInt(params.year),
            'NgayTinhLuong': new Date().toISOString().split('T')[0],
            'ThucLinh': totalIncome
        };
        
        var result = await callSupabase('insert', 'BangLuongThang', payrollRecord);
        
        if (result.status === 'success') {
            await refreshSingleSheet('BangLuongThang');
            showLoading(false);
            
            Swal.fire({
                icon: 'success',
                title: 'Đã lưu lương!',
                text: `Lương của ${empName} đã được lưu vào bảng lương tháng ${params.month}/${params.year}`,
                confirmButtonColor: '#28a745'
            });
        } else {
            throw new Error(result.message);
        }
        
    } catch (error) {
        showLoading(false);
        Swal.fire({
            icon: 'error',
            title: 'Lỗi lưu lương',
            text: error.message || 'Không thể lưu lương',
            confirmButtonColor: '#dc3545'
        });
    }
}

// --- XEM LỊCH SỬ LƯƠNG ---
window.viewPayrollHistory = function(empId) {
    var payrollData = GLOBAL_DATA['BangLuongThang'] || [];
    var userData = GLOBAL_DATA['User'] || [];
    
    var user = userData.find(u => u['ID'] === empId);
    var empName = user?.['Họ và tên'] || user?.['HoTen'] || empId;
    
    var history = payrollData
        .filter(p => p['ID_NhanVien'] === empId && p['Delete'] !== 'X')
        .sort((a, b) => {
            var dateA = new Date(parseInt(a['Nam']), parseInt(a['Thang']) - 1);
            var dateB = new Date(parseInt(b['Nam']), parseInt(b['Thang']) - 1);
            return dateB - dateA; // Sắp xếp giảm dần (mới nhất trước)
        });
    
    var historyHtml = '';
    if (history.length > 0) {
        historyHtml = `
            <div style="max-height: 400px; overflow-y: auto;">
                <table class="table table-sm table-striped">
                    <thead class="table-primary sticky-top">
                        <tr>
                            <th>Tháng/Năm</th>
                            <th class="text-end">Giờ công</th>
                            <th class="text-end">Giờ TC</th>
                            <th class="text-end">Thực lĩnh</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${history.map(h => `
                            <tr style="cursor: pointer;" onclick="showPayrollDetail(${JSON.stringify(h).replace(/"/g, '&quot;')}, ${h['Thang']}, ${h['Nam']})">
                                <td>${h['Thang'] || h['Tháng']}/${h['Nam'] || h['Năm']}</td>
                                <td class="text-end number-display">${formatNumber(parseFloat(h['TongGioCong']) || 0, 1)}</td>
                                <td class="text-end number-display">${formatNumber(parseFloat(h['TongGioTangCa']) || 0, 1)}</td>
                                <td class="text-end fw-bold money-display">${formatMoney(parseFloat(h['ThucLinh']) || 0)}</td>
                            </tr>
                        `).join('')}
                    </tbody>
                </table>
            </div>
        `;
    } else {
        historyHtml = '<div class="text-center text-muted py-4">Chưa có lịch sử lương</div>';
    }
    
    Swal.fire({
        title: `<i class="fas fa-history me-2"></i>Lịch sử lương - ${empName}`,
        html: historyHtml,
        width: '600px',
        showCloseButton: true,
        showConfirmButton: false
    });
}

// --- MỞ CÔNG CỤ TÍNH LƯƠNG ---
window.openSalaryCalculator = function() {
    var userData = GLOBAL_DATA['User'] || [];
    var activeUsers = userData.filter(u => u['Delete'] !== 'X' && u['Tinh luong'] !== 'Không');

    var userOptions = '<option value="">-- Chọn nhân viên --</option>';
    activeUsers.forEach(u => {
        var name = u['Họ và tên'] || u['HoTen'] || u['ID'];
        userOptions += `<option value="${u['ID']}">${u['ID']} - ${name}</option>`;
    });

    var html = `
        <div style="text-align: left;">
            <div class="mb-3">
                <label class="form-label fw-bold">Loại tính lương:</label>
                <select id="salary-calc-type" class="form-select" onchange="toggleSalaryCalcMode()">
                    <option value="group">Tính lương nhóm (theo kỳ)</option>
                    <option value="sudden">Tính lương nghỉ đột xuất (1 người)</option>
                </select>
            </div>

            <div id="group-calc-area">
                <div class="mb-3">
                    <label class="form-label fw-bold">Từ ngày:</label>
                    <input type="date" id="salary-from-date" class="form-control">
                </div>
                <div class="mb-3">
                    <label class="form-label fw-bold">Đến ngày:</label>
                    <input type="date" id="salary-to-date" class="form-control">
                </div>
                <div class="mb-3">
                    <label class="form-label fw-bold">Chọn nhóm nhân viên:</label>
                    <select id="salary-user-group" class="form-select" multiple size="6">
                        ${activeUsers.map(u => {
                            var name = u['Họ và tên'] || u['HoTen'] || u['ID'];
                            return `<option value="${u['ID']}">${u['ID']} - ${name}</option>`;
                        }).join('')}
                    </select>
                    <small class="text-muted">Giữ Ctrl để chọn nhiều người</small>
                </div>
            </div>

            <div id="sudden-calc-area" style="display: none;">
                <div class="mb-3">
                    <label class="form-label fw-bold">Nhân viên:</label>
                    <select id="salary-sudden-user" class="form-select">
                        ${userOptions}
                    </select>
                </div>
                <div class="mb-3">
                    <label class="form-label fw-bold">Từ ngày:</label>
                    <input type="date" id="salary-sudden-from" class="form-control">
                </div>
                <div class="mb-3">
                    <label class="form-label fw-bold">Đến ngày:</label>
                    <input type="date" id="salary-sudden-to" class="form-control">
                </div>
                <div class="mb-3">
                    <label class="form-label fw-bold">Lý do nghỉ:</label>
                    <textarea id="salary-sudden-reason" class="form-control" rows="2" placeholder="Ví dụ: Nghỉ ốm, việc gia đình..."></textarea>
                </div>
            </div>
        </div>
    `;

    Swal.fire({
        title: '<i class="fas fa-calculator me-2"></i>Công cụ tính lương',
        html: html,
        width: '550px',
        showCancelButton: true,
        confirmButtonText: '<i class="fas fa-check me-1"></i> Tính toán',
        cancelButtonText: 'Hủy',
        confirmButtonColor: '#2E7D32',
        customClass: {
            popup: 'rounded-4 shadow-lg'
        },
        preConfirm: () => {
            var calcType = document.getElementById('salary-calc-type').value;
            if (calcType === 'group') {
                var fromDate = document.getElementById('salary-from-date').value;
                var toDate = document.getElementById('salary-to-date').value;
                var selectedUsers = Array.from(document.getElementById('salary-user-group').selectedOptions).map(o => o.value);

                if (!fromDate || !toDate) {
                    Swal.showValidationMessage('Vui lòng chọn đầy đủ ngày bắt đầu và kết thúc');
                    return false;
                }
                if (selectedUsers.length === 0) {
                    Swal.showValidationMessage('Vui lòng chọn ít nhất 1 nhân viên');
                    return false;
                }
                return { type: 'group', fromDate, toDate, users: selectedUsers };
            } else {
                var userId = document.getElementById('salary-sudden-user').value;
                var fromDate = document.getElementById('salary-sudden-from').value;
                var toDate = document.getElementById('salary-sudden-to').value;
                var reason = document.getElementById('salary-sudden-reason').value;

                if (!userId) {
                    Swal.showValidationMessage('Vui lòng chọn nhân viên');
                    return false;
                }
                if (!fromDate || !toDate) {
                    Swal.showValidationMessage('Vui lòng chọn đầy đủ ngày bắt đầu và kết thúc');
                    return false;
                }
                return { type: 'sudden', userId, fromDate, toDate, reason };
            }
        }
    }).then((result) => {
        if (result.isConfirmed) {
            processSalaryCalculation(result.value);
        }
    });
}

window.toggleSalaryCalcMode = function() {
    var calcType = document.getElementById('salary-calc-type').value;
    var groupArea = document.getElementById('group-calc-area');
    var suddenArea = document.getElementById('sudden-calc-area');

    if (calcType === 'group') {
        groupArea.style.display = 'block';
        suddenArea.style.display = 'none';
    } else {
        groupArea.style.display = 'none';
        suddenArea.style.display = 'block';
    }
}

function processSalaryCalculation(params) {
    showLoading(true, 'Đang tính toán lương...');

    setTimeout(function() {
        if (params.type === 'group') {
            calculateGroupSalary(params.fromDate, params.toDate, params.users);
        } else {
            calculateSuddenLeaveSalary(params.userId, params.fromDate, params.toDate, params.reason);
        }
    }, 500);
}

function calculateGroupSalary(fromDate, toDate, userIds) {
    var chamcongData = GLOBAL_DATA['Chamcong'] || [];
    var userData = GLOBAL_DATA['User'] || [];

    var results = [];

    userIds.forEach(userId => {
        var empInfo = userData.find(u => u['ID'] === userId);
        var empName = empInfo ? (empInfo['Họ và tên'] || empInfo['HoTen']) : userId;

        // Lọc chấm công trong khoảng thời gian
        var records = chamcongData.filter(r => {
            if (r['Delete'] === 'X') return false;
            if (r['ID_NhanVien'] !== userId) return false;

            var dateStr = r['Ngày'] || r['Ngay'];
            if (!dateStr) return false;

            var date = new Date(dateStr);
            var from = new Date(fromDate);
            var to = new Date(toDate);

            return date >= from && date <= to;
        });

        var totalHours = 0;
        var totalOvertime = 0;

        records.forEach(r => {
            totalHours += parseFloat(r['SoGioCong']) || 0;
            totalOvertime += parseFloat(r['SoGioTangCa']) || 0;
        });

        var luongCoBan = totalHours * 45000;
        var tienTangCa = totalOvertime * 45000 * 1.5;
        var phuCap = 1750000;
        var tongLuong = luongCoBan + tienTangCa + phuCap;

        results.push({
            id: userId,
            name: empName,
            hours: totalHours,
            overtime: totalOvertime,
            salary: tongLuong
        });
    });

    showLoading(false);

    var html = `
        <div style="text-align: left; max-height: 400px; overflow-y: auto;">
            <div class="alert alert-info mb-3">
                <strong>Kỳ lương:</strong> ${new Date(fromDate).toLocaleDateString('vi-VN')} - ${new Date(toDate).toLocaleDateString('vi-VN')}
            </div>
            <table class="table table-sm table-bordered">
                <thead class="table-success">
                    <tr>
                        <th>Mã NV</th>
                        <th>Tên</th>
                        <th class="text-end">Giờ công</th>
                        <th class="text-end">Giờ TC</th>
                        <th class="text-end">Tổng lương</th>
                    </tr>
                </thead>
                <tbody>
                    ${results.map(r => `
                        <tr>
                            <td>${r.id}</td>
                            <td>${r.name}</td>
                            <td class="text-end">${r.hours.toFixed(1)}</td>
                            <td class="text-end text-danger">${r.overtime.toFixed(1)}</td>
                            <td class="text-end fw-bold text-success">${r.salary.toLocaleString('vi-VN')} đ</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
    `;

    Swal.fire({
        title: 'Kết quả tính lương nhóm',
        html: html,
        width: '700px',
        confirmButtonText: 'Đóng',
        confirmButtonColor: '#2E7D32'
    });
}

function calculateSuddenLeaveSalary(userId, fromDate, toDate, reason) {
    var chamcongData = GLOBAL_DATA['Chamcong'] || [];
    var userData = GLOBAL_DATA['User'] || [];

    var empInfo = userData.find(u => u['ID'] === userId);
    var empName = empInfo ? (empInfo['Họ và tên'] || empInfo['HoTen']) : userId;

    var records = chamcongData.filter(r => {
        if (r['Delete'] === 'X') return false;
        if (r['ID_NhanVien'] !== userId) return false;

        var dateStr = r['Ngày'] || r['Ngay'];
        if (!dateStr) return false;

        var date = new Date(dateStr);
        var from = new Date(fromDate);
        var to = new Date(toDate);

        return date >= from && date <= to;
    });

    var totalHours = 0;
    records.forEach(r => {
        totalHours += parseFloat(r['SoGioCong']) || 0;
    });

    var from = new Date(fromDate);
    var to = new Date(toDate);
    var daysDiff = Math.ceil((to - from) / (1000 * 60 * 60 * 24)) + 1;

    var luongNghi = totalHours * 45000;

    showLoading(false);

    var html = `
        <div style="text-align: left;">
            <div class="mb-3">
                <strong>Nhân viên:</strong> ${empName} (${userId})
            </div>
            <div class="mb-3">
                <strong>Kỳ nghỉ:</strong> ${from.toLocaleDateString('vi-VN')} - ${to.toLocaleDateString('vi-VN')} (${daysDiff} ngày)
            </div>
            <div class="mb-3">
                <strong>Lý do:</strong> ${reason || 'Không có'}
            </div>
            <hr>
            <div class="mb-2">
                <strong>Tổng giờ công đã làm:</strong> <span class="text-success">${totalHours.toFixed(1)} giờ</span>
            </div>
            <div class="mb-2">
                <strong>Lương tính được:</strong> <span class="text-success fw-bold fs-5">${luongNghi.toLocaleString('vi-VN')} đ</span>
            </div>
        </div>
    `;

    Swal.fire({
        title: 'Lương nghỉ đột xuất',
        html: html,
        icon: 'info',
        confirmButtonText: 'Đóng',
        confirmButtonColor: '#2E7D32'
    });
}

// --- XUẤT EXCEL BẢNG CHẤM CÔNG THEO ĐỊNH DẠNG PIVOT ---
window.exportAttendanceToExcel = function() {
    var month = document.getElementById('attendance-month-filter')?.value;
    var year = document.getElementById('attendance-year-filter')?.value;
    var searchText = (document.getElementById('attendance-search')?.value || '').toLowerCase().trim();

    Swal.fire({
        icon: 'info',
        title: 'Xuất Excel',
        text: `Đang chuẩn bị file Excel định dạng pivot cho tháng ${month}/${year}...`,
        timer: 2000,
        showConfirmButton: false
    });

    setTimeout(function() {
        try {
            var data = GLOBAL_DATA['Chamcong'] || [];
            var userData = GLOBAL_DATA['User'] || [];
            
            // Lọc dữ liệu theo tháng/năm
            var activeData = data.filter(r => {
                if (r['Delete'] === 'X') return false;
                var dateStr = r['Ngày'] || r['Ngay'];
                if (!dateStr) return false;
                
                var date = new Date(dateStr);
                var recordMonth = date.getMonth() + 1;
                var recordYear = date.getFullYear();
                
                if (month && recordMonth !== parseInt(month)) return false;
                if (year && recordYear !== parseInt(year)) return false;
                
                return true;
            });

            if (activeData.length === 0) {
                Swal.fire({
                    icon: 'warning',
                    title: 'Không có dữ liệu',
                    text: 'Không có dữ liệu chấm công cho tháng/năm đã chọn',
                    confirmButtonColor: '#ffc107'
                });
                return;
            }

            // Group data by employee (giống logic renderAttendanceTable)
            var employeeMap = {};
            
            activeData.forEach(record => {
                var empId = record['ID_NhanVien'] || record['Mã nhân viên'];
                if (!empId) return;

                var empInfo = userData.find(u => u['ID'] === empId);
                var empName = empInfo ? (empInfo['Họ và tên'] || empInfo['HoTen']) : empId;

                if (!employeeMap[empId]) {
                    employeeMap[empId] = {
                        id: empId,
                        name: empName,
                        days: {},
                        totalHours: 0,
                        totalOvertime: 0
                    };
                }

                var dateStr = record['Ngày'] || record['Ngay'];
                var date = new Date(dateStr);
                var day = date.getDate();

                var hours = parseFloat(record['SoGioCong']) || 0;
                var overtime = parseFloat(record['SoGioTangCa']) || 0;

                employeeMap[empId].days[day] = {
                    hours: hours,
                    overtime: overtime
                };

                employeeMap[empId].totalHours += hours;
                employeeMap[empId].totalOvertime += overtime;
            });

            // Tính số ngày trong tháng
            var daysInMonth = 31;
            if (month && year) {
                daysInMonth = new Date(year, month, 0).getDate();
            }

            // Tạo workbook
            var wb = XLSX.utils.book_new();
            var wsData = [];
            
            // Tạo tiêu đề bảng với ngày tháng (giống hình mẫu)
            var monthNames = ['', 'MỘT', 'HAI', 'BA', 'TƯ', 'NĂM', 'SÁU', 
                             'BẢY', 'TÁM', 'CHÍN', 'MƯỜI', 'MƯỜI MỘT', 'MƯỜI HAI'];
            var titleRow = [`BẢNG CHẤM CÔNG THÁNG ${monthNames[parseInt(month)] || month}/${year}`];
            wsData.push(titleRow);
            
            // Tạo header với 3 dòng (giống hình mẫu)
            // Dòng 1: Các số ngày (1, 2, 3, ...)
            var headerRow1 = ['Mã NV', 'Tên nhân viên', 'Tổng giờ chính', 'Tổng giờ TC'];
            for (var d = 1; d <= daysInMonth; d++) {
                headerRow1.push(d, ''); // Số ngày và cột trống để merge
            }
            
            // Dòng 2: GC TC cho mỗi ngày
            var headerRow2 = ['', '', '', ''];
            for (var d = 1; d <= daysInMonth; d++) {
                headerRow2.push('GC', 'TC');
            }
            
            wsData.push(headerRow1);
            wsData.push(headerRow2);

            // Thêm dữ liệu nhân viên
            Object.values(employeeMap).forEach(emp => {
                // Lọc theo search text nếu có
                if (searchText) {
                    var empText = (emp.id + ' ' + emp.name).toLowerCase();
                    if (!empText.includes(searchText)) return;
                }

                var row = [
                    emp.id, // Mã NV
                    emp.name, // Tên nhân viên
                    emp.totalHours, // Tổng giờ chính
                    emp.totalOvertime // Tổng giờ TC
                ];

                // Thêm dữ liệu cho từng ngày
                for (var day = 1; day <= daysInMonth; day++) {
                    if (emp.days[day]) {
                        // Nếu giá trị < 1 thì để trống, ngược lại hiển thị giá trị
                        var regularHours = emp.days[day].hours || 0;
                        var overtimeHours = emp.days[day].overtime || 0;
                        
                        row.push(regularHours >= 1 ? regularHours : ''); // Giờ chính
                        row.push(overtimeHours >= 1 ? overtimeHours : ''); // Giờ thêm
                    } else {
                        row.push(''); // Giờ chính - ô trống
                        row.push(''); // Giờ thêm - ô trống
                    }
                }

                wsData.push(row);
            });

            // Thêm dòng tổng cộng với công thức
            var totalRow = ['', 'TỔNG CỘNG'];
            var startDataRow = 4; // Dữ liệu bắt đầu từ row 4 (sau header)
            var endDataRow = wsData.length;
            
            // Tổng giờ chính và giờ TC
            totalRow.push({ f: `=SUM(C${startDataRow}:C${endDataRow})` }); // Tổng giờ chính
            totalRow.push({ f: `=SUM(D${startDataRow}:D${endDataRow})` }); // Tổng giờ TC
            
            // Tổng cho từng ngày
            var colIndex = 5; // Bắt đầu từ cột E (column 5)
            for (var d = 1; d <= daysInMonth; d++) {
                var colLetter1 = XLSX.utils.encode_col(colIndex - 1); // Giờ chính
                var colLetter2 = XLSX.utils.encode_col(colIndex); // Giờ thêm
                
                totalRow.push({ f: `=SUM(${colLetter1}${startDataRow}:${colLetter1}${endDataRow})` }); // Tổng giờ chính ngày d
                totalRow.push({ f: `=SUM(${colLetter2}${startDataRow}:${colLetter2}${endDataRow})` }); // Tổng giờ thêm ngày d
                
                colIndex += 2;
            }
            
            wsData.push(totalRow);

            // Tạo worksheet
            var ws = XLSX.utils.aoa_to_sheet(wsData);

            // Merge cells cho tiêu đề và header (giống hình mẫu)
            var merges = [];
            
            // Merge tiêu đề bảng (dòng đầu tiên)
            var totalCols = 4 + daysInMonth * 2;
            merges.push({
                s: { r: 0, c: 0 }, // Start cell A1
                e: { r: 0, c: totalCols - 1 } // End cell (toàn bộ dòng đầu)
            });
            
            // Merge cells cho số ngày (dòng 2) - mỗi số merge 2 cột GC/TC
            for (var d = 1; d <= daysInMonth; d++) {
                var startCol = 4 + (d - 1) * 2; // Cột bắt đầu cho ngày d
                var endCol = startCol + 1; // Cột kết thúc
                
                merges.push({
                    s: { r: 1, c: startCol }, // Start cell (dòng 2)
                    e: { r: 1, c: endCol }    // End cell
                });
            }
            ws['!merges'] = merges;

            // Set column widths giống hình mẫu
            var colWidths = [
                { wch: 8 },  // Mã NV
                { wch: 15 }, // Tên nhân viên  
                { wch: 10 }, // Tổng giờ chính
                { wch: 10 }  // Tổng giờ TC
            ];
            
            // Width cho các cột ngày (mỗi ngày 2 cột nhỏ)
            for (var d = 1; d <= daysInMonth; d++) {
                colWidths.push({ wch: 4 }); // GC - cột nhỏ
                colWidths.push({ wch: 4 }); // TC - cột nhỏ
            }
            
            ws['!cols'] = colWidths;

            // Định dạng theo mẫu hình ảnh
            var range = XLSX.utils.decode_range(ws['!ref']);
            
            // Định dạng tiêu đề (giống hình mẫu - nền vàng)
            var titleCell = ws['A1'];
            if (titleCell) {
                titleCell.s = {
                    font: { bold: true, size: 12, color: { rgb: "000000" } },
                    fill: { fgColor: { rgb: "FFFF00" } }, // Nền vàng như hình mẫu
                    alignment: { horizontal: "center", vertical: "center" },
                    border: {
                        top: { style: "thin", color: { rgb: "000000" } },
                        bottom: { style: "thin", color: { rgb: "000000" } },
                        left: { style: "thin", color: { rgb: "000000" } },
                        right: { style: "thin", color: { rgb: "000000" } }
                    }
                };
            }
            
            // Định dạng header giống hình mẫu
            for (var C = 0; C < 4 + daysInMonth * 2; ++C) {
                var headerCell1 = ws[XLSX.utils.encode_cell({ r: 1, c: C })]; // Dòng header 1 (số ngày)
                var headerCell2 = ws[XLSX.utils.encode_cell({ r: 2, c: C })]; // Dòng header 2 (GC/TC)
                
                // Định dạng dòng số ngày và thông tin cơ bản
                if (headerCell1) {
                    var bgColor1 = "D0D0D0"; // Xám nhạt mặc định
                    var textColor1 = "000000"; // Chữ đen
                    
                    headerCell1.s = {
                        font: { bold: true, size: 10, color: { rgb: textColor1 } },
                        fill: { fgColor: { rgb: bgColor1 } },
                        alignment: { horizontal: "center", vertical: "center" },
                        border: {
                            top: { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } },
                            left: { style: "thin", color: { rgb: "000000" } },
                            right: { style: "thin", color: { rgb: "000000" } }
                        }
                    };
                }
                
                // Định dạng dòng GC/TC
                if (headerCell2) {
                    var bgColor2 = "D0D0D0"; // Xám nhạt
                    var textColor2 = "000000"; // Chữ đen
                    
                    headerCell2.s = {
                        font: { bold: true, size: 9, color: { rgb: textColor2 } },
                        fill: { fgColor: { rgb: bgColor2 } },
                        alignment: { horizontal: "center", vertical: "center" },
                        border: {
                            top: { style: "thin", color: { rgb: "000000" } },
                            bottom: { style: "thin", color: { rgb: "000000" } },
                            left: { style: "thin", color: { rgb: "000000" } },
                            right: { style: "thin", color: { rgb: "000000" } }
                        }
                    };
                }
            }
            
            // Định dạng dữ liệu giống hình mẫu (trắng, đơn giản)
            for (var R = 3; R <= range.e.r - 1; ++R) { // Bắt đầu từ dữ liệu (row 4), trừ dòng tổng cuối
                for (var C = 0; C <= range.e.c; ++C) { // Tất cả các cột
                    var cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                    if (ws[cellRef]) {
                        var cellStyle = {
                            alignment: { horizontal: "center", vertical: "center" },
                            border: {
                                top: { style: "thin", color: { rgb: "000000" } },
                                bottom: { style: "thin", color: { rgb: "000000" } },
                                left: { style: "thin", color: { rgb: "000000" } },
                                right: { style: "thin", color: { rgb: "000000" } }
                            },
                            fill: { fgColor: { rgb: "FFFFFF" } }, // Nền trắng
                            font: { size: 10, color: { rgb: "000000" } } // Chữ đen, size 10
                        };
                        
                        // Định dạng số (hiển thị 1 chữ số thập phân)
                        if (typeof ws[cellRef].v === 'number') {
                            ws[cellRef].z = '0.0';
                        }
                        
                        // Làm đậm cho cột tên và tổng cộng
                        if (C === 1 || C === 2 || C === 3) {
                            cellStyle.font.bold = true;
                        }
                        
                        ws[cellRef].s = cellStyle;
                    }
                }
            }
            
            // Định dạng dòng tổng cộng (giống hình mẫu - nền trắng, chữ đậm)
            var totalRowIndex = wsData.length - 1;
            for (var C = 0; C < 4 + daysInMonth * 2; ++C) {
                var totalCell = ws[XLSX.utils.encode_cell({ r: totalRowIndex, c: C })];
                if (totalCell) {
                    totalCell.s = {
                        font: { bold: true, size: 10, color: { rgb: "000000" } },
                        fill: { fgColor: { rgb: "FFFFFF" } }, // Nền trắng
                        alignment: { horizontal: "center", vertical: "center" },
                        border: {
                            top: { style: "thick", color: { rgb: "000000" } },
                            bottom: { style: "thick", color: { rgb: "000000" } },
                            left: { style: "thin", color: { rgb: "000000" } },
                            right: { style: "thin", color: { rgb: "000000" } }
                        }
                    };
                }
            }

            // Thêm worksheet vào workbook
            var sheetName = `Chấm công T${month}-${year}`;
            XLSX.utils.book_append_sheet(wb, ws, sheetName);

            // Xuất file
            var fileName = `BangChamCong_Pivot_Thang${month}_${year}.xlsx`;
            XLSX.writeFile(wb, fileName);

            Swal.fire({
                icon: 'success',
                title: 'Thành công!',
                html: `
                    <div class="text-start">
                        <p><strong>File Excel đã được tải xuống với định dạng pivot:</strong></p>
                        <ul class="list-unstyled">
                            <li>✅ Định dạng giống view hiện tại</li>
                            <li>✅ Mỗi ngày có 2 cột: <strong>GC</strong> (Giờ Chính) và <strong>TC</strong> (Tăng Ca)</li>
                            <li>✅ Giá trị < 1 hiển thị ô trống</li>
                            <li>✅ Tiêu đề có tên tháng và năm (IN HOA)</li>
                            <li>✅ Màu sắc sinh động: Xanh lá (GC), Hồng (TC)</li>
                            <li>✅ Bôi màu Thứ 7 và Chủ nhật với gradient đẹp mắt</li>
                            <li>✅ Viền và border chuyên nghiệp</li>
                            <li>✅ Tổng cộng theo nhân viên và toàn bộ</li>
                            <li>✅ Công thức Excel tự động tính toán</li>
                        </ul>
                        <p class="text-muted small">File: <code>${fileName}</code></p>
                    </div>
                `,
                confirmButtonText: 'Đóng',
                confirmButtonColor: '#28a745'
            });

        } catch (error) {
            console.error('Lỗi xuất Excel:', error);
            Swal.fire({
                icon: 'error',
                title: 'Lỗi xuất Excel',
                text: 'Có lỗi xảy ra khi tạo file Excel: ' + error.message,
                confirmButtonColor: '#dc3545'
            });
        }
    }, 100);
}

// --- TABLE RENDER ---
function renderTable(data) {
    data = Array.isArray(data) ? data : [];

        // Kiểm tra các bảng có render đặc biệt
        if (CURRENT_SHEET === 'Chamcong') {
            renderAttendanceTable(data);
            return;
        }
        
        if (CURRENT_SHEET === 'BangLuongThang') {
            renderPayrollTable(data);
            return;
        }

    // Xóa class attendance-table nếu không phải bảng chấm công
    var table = document.getElementById('data-table');
    if (table) table.classList.remove('attendance-table');

    if (!document.getElementById('filter-bar')) {
        var filterHtml = `
            <div id="filter-bar" class="bg-white p-3 rounded-3 shadow-sm border filter-bar-block">
                <div class="filter-item">
                    <label class="form-label small fw-bold text-muted mb-1">Từ ngày</label>
                    <input type="date" id="filter-from" class="form-control form-control-sm" oninput="applyFilterInstant()" onchange="applyFilterInstant()">
                </div>
                <div class="filter-item">
                    <label class="form-label small fw-bold text-muted mb-1">Đến ngày</label>
                    <input type="date" id="filter-to" class="form-control form-control-sm" oninput="applyFilterInstant()" onchange="applyFilterInstant()">
                </div>
                <div class="filter-item">
                    <label class="form-label small fw-bold text-muted mb-1">Nhân sự</label>
                    <select id="filter-user" class="form-select form-select-sm" onchange="applyFilterInstant()">
                        <option value="">-- Tất cả --</option>
                    </select>
                </div>
                <div class="filter-item">
                    <label class="form-label small fw-bold text-muted mb-1">Kho</label>
                    <select id="filter-kho" class="form-select form-select-sm" onchange="applyFilterInstant()">
                        <option value="">-- Tất cả kho --</option>
                    </select>
                </div>
                <div class="filter-item filter-item-search">
                    <label class="form-label small fw-bold text-muted mb-1">Tìm kiếm nhanh</label>
                    <div class="input-group input-group-sm">
                        <span class="input-group-text bg-light"><i class="fas fa-search"></i></span>
                        <input type="text" id="main-search" class="form-control" placeholder="Gõ để tìm..." oninput="applyFilterInstant()">
                    </div>
                </div>
            </div>`;
        var wrap = document.getElementById('filter-bar-wrap');
        if (wrap) wrap.innerHTML = filterHtml;
        else document.getElementById('app-container').querySelector('.main-content').insertAdjacentHTML('afterbegin', filterHtml);

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

    var el = function (id) { return document.getElementById(id); };
    var filterVal = (el('main-search')?.value || '').toLowerCase().trim();
    var filterKho = (el('filter-kho')?.value || '').trim();
    var filterUser = (el('filter-user')?.value || '').trim();
    var filterFrom = (el('filter-from')?.value || '').trim();
    var filterTo = (el('filter-to')?.value || '').trim();

    function getRowDate(row) {
        for (var k in row) {
            var l = k.toLowerCase();
            if (l.indexOf('ngay') !== -1 || l.indexOf('date') !== -1) {
                var v = row[k];
                if (v) return new Date(v);
            }
        }
        return null;
    }

    var activeData = data.filter(function (r) {
        if (r['Delete'] == 'X') return false;

        var rowText = Object.values(r).join(' ').toLowerCase();
        if (filterVal && !rowText.includes(filterVal)) return false;

        if (filterKho) {
            var kho = (r['Tên kho'] || r['Ten_Kho'] || r['Tên Kho'] || '').toString().trim();
            if (kho !== filterKho) return false;
        }

        if (filterUser) {
            var user = (r['Người nhập'] || r['Người xuất'] || r['Người lập'] || r['Người Xuất'] || r['Nhân viên'] || r['Họ và tên'] || r['Name'] || '').toString().trim();
            if (user !== filterUser) return false;
        }

        var rowDate = getRowDate(r);
        if (rowDate && !isNaN(rowDate.getTime())) {
            if (filterFrom) {
                var fromDate = new Date(filterFrom);
                if (rowDate < fromDate) return false;
            }
            if (filterTo) {
                var toDate = new Date(filterTo);
                toDate.setHours(23, 59, 59, 999);
                if (rowDate > toDate) return false;
            }
        }

        return true;
    });

    var skip = COLUMNS_HIDDEN;

    if (activeData.length === 0) {
        th.innerHTML = '<tr><th class="bg-success text-white">Thông tin</th></tr>';
        tb.innerHTML = '<tr><td class="text-center p-5 text-muted fst-italic">Không có dữ liệu hiển thị</td></tr>';
        return;
    }

    // Lấy đủ cột từ tất cả dòng (tránh thiếu cột nếu dòng đầu không đủ)
    var keys = [];
    activeData.forEach(function (r) {
        Object.keys(r).forEach(function (k) {
            if (!skip.includes(k) && keys.indexOf(k) === -1) keys.push(k);
        });
    });
    if (keys.length === 0) keys = Object.keys(activeData[0] || {}).filter(function (k) { return !skip.includes(k); });

    var trH = document.createElement('tr');
    keys.forEach(k => { var t = document.createElement('th'); t.innerText = COLUMN_MAP[k] || k; trH.appendChild(t); });
    trH.appendChild(document.createElement('th'));
    th.appendChild(trH);

    activeData.forEach((r) => {
        var absoluteIdx = data.indexOf(r);
        var tr = document.createElement('tr');
        tr.style.cursor = 'pointer';
        keys.forEach(k => {
            if (!skip.includes(k)) {
                var td = document.createElement('td');
                var v = r[k];

                // Resolve foreign key trước tiên
                v = resolveForeignKey(k, v);

                if (!FIELD_CONSTRAINTS[k] && (k.toUpperCase().includes('NGAY') || k.toUpperCase().includes('DATE'))) v = formatSafeDate(v);
                if (isImageColumnKey(k) && typeof v === 'string' && v.trim() && (v.startsWith('http://') || v.startsWith('https://'))) {
                    td.innerHTML = '<img src="' + String(v).replace(/"/g, '&quot;') + '" class="table-cell-img" alt="" style="max-height:36px;max-width:48px;object-fit:cover;border-radius:4px;">';
                } else if (isMoneyField(k) && !isNaN(v) && v !== '' && v != null && Number(v) !== 0) {
                    // Sử dụng hàm formatMoney mới với phân tách hàng ngàn
                    td.innerHTML = formatMoney(v);
                    td.className = 'text-end fw-bold money-cell';
                    // Phân biệt màu theo loại tiền
                    if (k.toLowerCase().includes('thu') || k.toLowerCase().includes('lương') || k.toLowerCase().includes('phụ cấp') || k.toLowerCase().includes('thưởng')) {
                        td.style.color = '#28a745'; // Xanh cho thu nhập
                    } else if (k.toLowerCase().includes('trừ') || k.toLowerCase().includes('tạm ứng') || k.toLowerCase().includes('bảo hiểm') || k.toLowerCase().includes('nợ')) {
                        td.style.color = '#dc3545'; // Đỏ cho khoản trừ
                    } else if (k.toLowerCase().includes('thực lĩnh') || k.toLowerCase().includes('net')) {
                        td.style.color = '#6f42c1'; // Tím cho thực lĩnh
                    } else {
                        td.style.color = '#28a745'; // Mặc định xanh
                    }
                } else if (isNumericField(k) && !isNaN(v) && v !== '' && v != null) {
                    // Định dạng số thông thường (không phải tiền)
                    td.innerHTML = formatNumber(v, 1);
                    td.className = 'text-end fw-bold number-cell';
                    td.style.color = '#495057';
                } else {
                    td.innerText = v || '';
                }
                tr.appendChild(td);
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

function isImageColumnKey(key) {
    if (!key || typeof key !== 'string') return false;
    var k = key.toLowerCase().replace(/\s/g, '');
    return /anh|hinh|image|photo|ảnh|hình|avatar/.test(k) || key.indexOf('Hình ảnh') !== -1 || key.indexOf('Ảnh cá nhân') !== -1;
}

// --- UTILS ---
/** Lọc tức thì (không debounce): renderTable đọc main-search trong filter-bar */
window.applyFilterInstant = function () {
    renderTable(GLOBAL_DATA[CURRENT_SHEET] || []);
}

/** Lọc cho bảng lương */
window.applyPayrollFilter = function () {
    if (CURRENT_SHEET === 'BangLuongThang') {
        renderPayrollTable(GLOBAL_DATA[CURRENT_SHEET] || []);
    }
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

window.openEditModal = function (i, sheetOpt) {
    var sheet = sheetOpt || CURRENT_SHEET;
    // Nếu sửa bản ghi con (sheet khác CURRENT_SHEET), lưu sheet cha để khôi phục sau
    if (sheet !== CURRENT_SHEET) {
        window.SHEET_BEFORE_ADD = CURRENT_SHEET;
        CURRENT_SHEET = sheet;
    }
    EDIT_INDEX = i;
    document.getElementById('modalTitle').innerText = 'Cập Nhật';
    var d = (GLOBAL_DATA[sheet] || [])[i];
    buildForm(sheet, d);
    new bootstrap.Modal(document.getElementById('dataModal')).show();
}

function buildForm(s, d) {
    var f = document.getElementById('dynamic-form'); f.innerHTML = '';
    var smp = GLOBAL_DATA[s]?.[0] || (GLOBAL_DATA[s] && GLOBAL_DATA[s].length > 0 ? GLOBAL_DATA[s][0] : {});
    if (Object.keys(smp).length === 0) { f.innerHTML = '<div class="alert alert-warning">Chưa có dữ liệu mẫu.</div>'; return; }
    var skip = COLUMNS_HIDDEN;
    Object.keys(smp).forEach(k => {
        if (skip.includes(k)) return;
        var v = d ? (d[k] != null ? d[k] : '') : '';
        if (d && !FIELD_CONSTRAINTS[k] && (k.includes('Ngay') || k.includes('Date'))) {
            try { v = new Date(v).toISOString().split('T')[0]; } catch (e) { v = ''; }
        }
        var div = document.createElement('div');
        div.className = 'col-md-6';
        var drop = getDropdownForField(s, k);
        var fkMapping = FK_MAPPING[k];

        // Check DROPDOWN_CONFIG first, then FK_MAPPING
        if (drop && GLOBAL_DATA[drop.table]) {
            var opts = (GLOBAL_DATA[drop.table] || []).filter(function (row) { return row['Delete'] !== 'X'; });
            var labelKey = drop.labelKey;
            var valueKey = drop.valueKey;
            var optionsHtml = '<option value="">-- Chọn --</option>';
            opts.forEach(function (row) {
                var val = row[valueKey] != null ? String(row[valueKey]) : '';
                var lbl = (row[labelKey] != null ? row[labelKey] : row['Tên kho'] || row['Ten_Kho'] || row['Tên vật tư'] || row['Họ và tên'] || row['Tên nhóm'] || row['Tên đối tác'] || val);
                optionsHtml += '<option value="' + String(val).replace(/"/g, '&quot;') + '"' + (val === String(v) ? ' selected' : '') + '>' + String(lbl).replace(/</g, '&lt;') + '</option>';
            });
            div.innerHTML = '<label class="form-label small fw-bold text-muted">' + (COLUMN_MAP[k] || k) + '</label><select class="form-select rounded-3" name="' + k + '">' + optionsHtml + '</select>';
        } else if (fkMapping && GLOBAL_DATA[fkMapping.table]) {
            // Use FK_MAPPING to create dropdown
            var opts = (GLOBAL_DATA[fkMapping.table] || []).filter(function (row) { return row['Delete'] !== 'X'; });

            // Filter by condition if this is a Drop table entry (more robust matching)
            if (fkMapping.condition) {
                var targetCondition = String(fkMapping.condition).toLowerCase().trim();
                opts = opts.filter(function(row) {
                    var rowCondition = String(row['condition'] || row['Condition'] || '').toLowerCase().trim();
                    return rowCondition === targetCondition;
                });
            }

            var labelKey = fkMapping.display;
            var valueKey = fkMapping.key;
            var optionsHtml = '<option value="">-- Chọn --</option>';

            opts.forEach(function (row) {
                // Handle key variations for different tables
                var val = null;
                if (fkMapping.table === 'DS_kho') {
                    val = row[valueKey] || row['ID kho'] || row['Id_kho'] || row['id_kho'] || row['IDKho'] || row['Idkho'];
                } else if (fkMapping.table === 'Vat_tu') {
                    val = row[valueKey] || row['ID vật tư'] || row['Id_vat_tu'] || row['id_vat_tu'] || row['MaVatTu'];
                } else if (fkMapping.table === 'Drop') {
                    val = row[valueKey] || row['id'] || row['ID'] || row['Id'];
                } else {
                    val = row[valueKey];
                }

                if (val != null) {
                    val = String(val);
                    var lbl = row[labelKey] != null ? row[labelKey] : val;
                    optionsHtml += '<option value="' + val.replace(/"/g, '&quot;') + '"' + (val === String(v) ? ' selected' : '') + '>' + String(lbl).replace(/</g, '&lt;') + '</option>';
                }
            });

            div.innerHTML = '<label class="form-label small fw-bold text-muted">' + (COLUMN_MAP[k] || k) + '</label><select class="form-select rounded-3" name="' + k + '">' + optionsHtml + '</select>';
        } else {
            var fc = FIELD_CONSTRAINTS[k];
            var inputAttrs = 'name="' + k + '" value="' + String(v).replace(/"/g, '&quot;') + '"';
            if (fc) {
                inputAttrs += ' type="' + fc.type + '"';
                if (fc.min != null) inputAttrs += ' min="' + fc.min + '"';
                if (fc.max != null) inputAttrs += ' max="' + fc.max + '"';
                if (fc.step != null) inputAttrs += ' step="' + fc.step + '"';
                if (fc.placeholder) inputAttrs += ' placeholder="' + fc.placeholder + '"';
            } else if (k.includes('Ngay')) {
                inputAttrs += ' type="date"';
            }
            div.innerHTML = '<label class="form-label small fw-bold text-muted">' + (COLUMN_MAP[k] || k) + '</label><input class="form-control rounded-3" ' + inputAttrs + '>';
        }
        f.appendChild(div);
    });
}

window.submitData = function () {
    var f = document.getElementById('dynamic-form'); var fd = {};
    f.querySelectorAll('input, select').forEach(function (el) { fd[el.name] = el.value; });
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

// --- MỞ CHI TIẾT BẢN GHI CON ---
window.openChildDetailPopup = function(sheet, idx) {
    Swal.close();
    setTimeout(function() {
        showRowDetail(null, sheet, idx);
    }, 100);
}

// Mapping config: định nghĩa cách resolve foreign keys
const FK_MAPPING = {
    'Người lập': { table: 'User', key: 'ID', display: 'Họ và tên' },
    'Nguoilap': { table: 'User', key: 'ID', display: 'Họ và tên' },
    'NguoiLap': { table: 'User', key: 'ID', display: 'Họ và tên' },
    'ID kho': { table: 'DS_kho', key: 'ID_kho', display: 'Tên kho' },
    'ID_kho': { table: 'DS_kho', key: 'ID_kho', display: 'Tên kho' },
    'Tên kho': { table: 'DS_kho', key: 'ID_kho', display: 'Tên kho' },
    'Ten_Kho': { table: 'DS_kho', key: 'ID_kho', display: 'Tên kho' },
    'Ten_kho': { table: 'DS_kho', key: 'ID_kho', display: 'Tên kho' },
    'Tên Kho': { table: 'DS_kho', key: 'ID_kho', display: 'Tên kho' },
    'KhoDi': { table: 'DS_kho', key: 'ID_kho', display: 'Tên kho' },
    'KhoDen': { table: 'DS_kho', key: 'ID_kho', display: 'Tên kho' },
    'Kho nguồn': { table: 'DS_kho', key: 'ID_kho', display: 'Tên kho' },
    'Kho đích': { table: 'DS_kho', key: 'ID_kho', display: 'Tên kho' },
    'ID vật tư': { table: 'Vat_tu', key: 'ID vật tư', display: 'Tên vật tư' },
    'MaVatTu': { table: 'Vat_tu', key: 'ID vật tư', display: 'Tên vật tư' },
    'Mã vật tư': { table: 'Vat_tu', key: 'ID vật tư', display: 'Tên vật tư' },
    'Loại chi phí': { table: 'Drop', key: 'id', display: 'label', condition: 'Loaichiphi' },
    'Loaichiphi': { table: 'Drop', key: 'id', display: 'label', condition: 'Loaichiphi' },
    'LoaiChiPhi': { table: 'Drop', key: 'id', display: 'label', condition: 'Loaichiphi' },
    'Loai_chi_phi': { table: 'Drop', key: 'id', display: 'label', condition: 'Loaichiphi' },
    'DonViTinh': { table: 'Drop', key: 'id', display: 'label', condition: 'Đơn vị tính' },
    'Đơn vị tính': { table: 'Drop', key: 'id', display: 'label', condition: 'Đơn vị tính' },
    'Don_vi_tinh': { table: 'Drop', key: 'id', display: 'label', condition: 'Đơn vị tính' },
    'ID_NhanVien': { table: 'User', key: 'ID', display: 'Họ và tên' },
    'Người nhập': { table: 'User', key: 'Họ và tên', display: 'Họ và tên' },
    'Người xuất': { table: 'User', key: 'Họ và tên', display: 'Họ và tên' },
    'Người chuyển': { table: 'User', key: 'Họ và tên', display: 'Họ và tên' },
    'Người nhận': { table: 'User', key: 'Họ và tên', display: 'Họ và tên' },

    // Các trường từ Drop table
    'Chức vụ': { table: 'Drop', key: 'id', display: 'label', condition: 'Chức vụ' },
    'ChucVu': { table: 'Drop', key: 'id', display: 'label', condition: 'Chức vụ' },
    'Chuc_vu': { table: 'Drop', key: 'id', display: 'label', condition: 'Chức vụ' },

    'Phê duyệt': { table: 'Drop', key: 'id', display: 'label', condition: 'Phê duyệt' },
    'PheDuyet': { table: 'Drop', key: 'id', display: 'label', condition: 'Phê duyệt' },
    'Phe_duyet': { table: 'Drop', key: 'id', display: 'label', condition: 'Phê duyệt' },

    'Loại nhập xuất': { table: 'Drop', key: 'id', display: 'label', condition: 'Loại nhập xuất' },
    'LoaiNhapXuat': { table: 'Drop', key: 'id', display: 'label', condition: 'Loại nhập xuất' },
    'Loai_nhap_xuat': { table: 'Drop', key: 'id', display: 'label', condition: 'Loại nhập xuất' },

    'Bộ phận': { table: 'Drop', key: 'id', display: 'label', condition: 'Bộ phận' },
    'BoPhan': { table: 'Drop', key: 'id', display: 'label', condition: 'Bộ phận' },
    'Bo_phan': { table: 'Drop', key: 'id', display: 'label', condition: 'Bộ phận' },

    'Trạng thái': { table: 'Drop', key: 'id', display: 'label', condition: 'Trạng thái' },
    'TrangThai': { table: 'Drop', key: 'id', display: 'label', condition: 'Trạng thái' },
    'Trang_thai': { table: 'Drop', key: 'id', display: 'label', condition: 'Trạng thái' },

    'Phân quyền': { table: 'Drop', key: 'id', display: 'label', condition: 'Phân quyền' },
    'PhanQuyen': { table: 'Drop', key: 'id', display: 'label', condition: 'Phân quyền' },
    'Phan_quyen': { table: 'Drop', key: 'id', display: 'label', condition: 'Phân quyền' },

    'Nhóm vật liệu': { table: 'Drop', key: 'id', display: 'label', condition: 'Nhóm vật liệu' },
    'NhomVatLieu': { table: 'Drop', key: 'id', display: 'label', condition: 'Nhóm vật liệu' },
    'Nhom_vat_lieu': { table: 'Drop', key: 'id', display: 'label', condition: 'Nhóm vật liệu' },

    'Trạng thái chuyển kho': { table: 'Drop', key: 'id', display: 'label', condition: 'Trạng thái chuyển kho' },
    'TrangThaiChuyenKho': { table: 'Drop', key: 'id', display: 'label', condition: 'Trạng thái chuyển kho' },
    'Trang_thai_chuyen_kho': { table: 'Drop', key: 'id', display: 'label', condition: 'Trạng thái chuyển kho' },

    'Loại giao dịch': { table: 'Drop', key: 'id', display: 'label', condition: 'LoaiGiaoDich' },
    'LoaiGiaoDich': { table: 'Drop', key: 'id', display: 'label', condition: 'LoaiGiaoDich' },
    'Loai_giao_dich': { table: 'Drop', key: 'id', display: 'label', condition: 'LoaiGiaoDich' },

    'Thời gian': { table: 'Drop', key: 'id', display: 'label', condition: 'Thoigian' },
    'ThoiGian': { table: 'Drop', key: 'id', display: 'label', condition: 'Thoigian' },
    'Thoi_gian': { table: 'Drop', key: 'id', display: 'label', condition: 'Thoigian' },

    'Thời tiết': { table: 'Drop', key: 'id', display: 'label', condition: 'Thoitiet' },
    'ThoiTiet': { table: 'Drop', key: 'id', display: 'label', condition: 'Thoitiet' },
    'Thoi_tiet': { table: 'Drop', key: 'id', display: 'label', condition: 'Thoitiet' }
};

// Function để resolve foreign key thành tên hiển thị
function resolveForeignKey(columnName, value) {
    if (!value || value === '') return value;

    var mapping = FK_MAPPING[columnName];
    if (!mapping) return value;

    var targetTable = GLOBAL_DATA[mapping.table] || [];
    if (targetTable.length === 0) return value;

    // Nếu có condition (ví dụ: Drop table với Loaichiphi) - case-insensitive matching
    if (mapping.condition) {
        var targetCondition = String(mapping.condition).toLowerCase().trim();
        targetTable = targetTable.filter(item => {
            var itemCondition = String(item['condition'] || item['Condition'] || '').toLowerCase().trim();
            return itemCondition === targetCondition;
        });
    }

    // Tìm record matching - thử nhiều variation của key
    var keyVariations = [mapping.key];
    if (mapping.table === 'DS_kho') {
        keyVariations.push('ID kho', 'Id_kho', 'id_kho', 'IDKho', 'Idkho');
    }
    if (mapping.table === 'Vat_tu') {
        keyVariations.push('ID vật tư', 'Id_vat_tu', 'id_vat_tu', 'MaVatTu');
    }
    if (mapping.table === 'Drop') {
        keyVariations.push('id', 'ID', 'Id');
    }

    var found = null;
    for (var i = 0; i < keyVariations.length && !found; i++) {
        found = targetTable.find(item => String(item[keyVariations[i]]) === String(value));
    }

    // Nếu không tìm thấy với key chính, thử với display field (case đã là tên rồi)
    if (!found) {
        found = targetTable.find(item => String(item[mapping.display]) === String(value));
        if (found) return value; // Đã là tên rồi, giữ nguyên
    }

    // Nếu tìm thấy, trả về display value
    if (found && found[mapping.display]) {
        return found[mapping.display];
    }

    // Thử tìm với các variation của tên cột (Tên kho, Ten_Kho, Tên Kho)
    if (!found && mapping.table === 'DS_kho') {
        var variations = ['Tên kho', 'Ten_Kho', 'Tên Kho'];
        for (var i = 0; i < variations.length; i++) {
            found = targetTable.find(item => String(item[variations[i]]) === String(value));
            if (found) {
                var displayValue = found[mapping.display] || found[variations[i]];
                if (displayValue) return displayValue;
            }
        }
    }

    return value; // Trả về giá trị gốc nếu không tìm thấy
}

function showRowDetail(r, s, idx) {
    if (!r && idx !== -1) r = (GLOBAL_DATA[s] || [])[idx];
    if (!r) return;

    var menuTitle = 'bản ghi';
    MENU_STRUCTURE.forEach(g => { var f = g.items.find(x => x.sheet == s); if (f) menuTitle = f.title; });

    var html = '<div class="row text-start detail-scroll-wrap" style="font-size:12px !important; max-height: 400px; overflow-y: auto; overflow-x: auto; padding: 15px; background: #fdfdfd; border-radius: 12px;">';
    var skip = COLUMNS_HIDDEN.concat(['ID_Phieu', 'ID phieu nhap', 'ID phieu Xuat', 'ID Chkho']);
    for (var k in r) {
        if (!skip.includes(k) && r[k]) {
            var val = r[k] || '';

            // Resolve foreign key trước tiên
            var resolvedVal = resolveForeignKey(k, val);

            if (!FIELD_CONSTRAINTS[k] && (k.toUpperCase().includes('NGAY') || k.toUpperCase().includes('DATE'))) resolvedVal = formatSafeDate(resolvedVal);

            var displayVal = resolvedVal;
            if (typeof val === 'string' && (val.startsWith('http') || val.startsWith('https'))) {
                if (isImageColumnKey(k) || val.match(/\.(jpeg|jpg|gif|png|webp)(\?|$)/i)) {
                    displayVal = `<div class="mt-1"><img src="${val.replace(/"/g, '&quot;')}" style="max-height:180px; max-width:100%; border-radius:8px; cursor:zoom-in; border: 2px solid #fff; box-shadow: 0 4px 6px rgba(0,0,0,0.1)" onclick="window.open(this.src)"></div>`;
                } else {
                    displayVal = `<a href="${val.replace(/"/g, '&quot;')}" target="_blank" class="btn btn-xs btn-outline-primary py-0 px-2 mt-1" style="font-size:10px">Mở tài liệu <i class="fas fa-external-link-alt ms-1"></i></a>`;
                }
            } else if (isMoneyField(k) && !isNaN(resolvedVal) && resolvedVal !== '' && resolvedVal != null && Number(resolvedVal) !== 0) {
                displayVal = `<span class="text-success fw-bold text-end money-display">${formatMoney(resolvedVal)}</span>`;
            } else if (isNumericField(k) && !isNaN(resolvedVal) && resolvedVal !== '' && resolvedVal != null) {
                displayVal = `<span class="text-info fw-bold text-end number-display">${formatNumber(resolvedVal, 1)}</span>`;
            }

            var isNumField = isMoneyField(k) && !isNaN(resolvedVal) && resolvedVal !== '';
            html += `<div class="col-md-4 col-6 mb-3 border-bottom pb-2 detail-field">
                <small class="field-label">${COLUMN_MAP[k] || k}</small>
                <span class="field-value ${isNumField ? 'text-end' : ''}">${displayVal}</span>
            </div>`;
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
            html += `<div class="table-responsive child-table-scroll" style="max-height: 250px; overflow-x: auto; overflow-y: auto; border-radius: 8px;"><table class="table table-sm table-bordered custom-table mb-0 child-table-inline child-detail-table" style="font-size:11px; white-space: nowrap;"><thead class="thead-sticky child-thead-success"><tr>`;
            var ck = Object.keys(childs[0]);
            ck.forEach(x => {
                if (!skip.includes(x) && x != fk) {
                    html += '<th class="' + (isMoneyField(x) ? 'text-end' : 'text-start') + '">' + (COLUMN_MAP[x] || x) + '</th>';
                }
            });
            html += '<th class="text-center" style="min-width: 100px;">Thao tác</th>';
            html += '</tr></thead><tbody>';
            childs.forEach((cr) => {
                var originalCIdx = (GLOBAL_DATA[cS] || []).indexOf(cr);
                html += '<tr class="child-row-clickable">';
                ck.forEach(x => {
                    if (!skip.includes(x) && x != fk) {
                        var cv = cr[x] || '';

                        // Apply FK mapping to child table cells
                        cv = resolveForeignKey(x, cv);

                        var isMoneyCol = isMoneyField(x);
                        var isNumCol = isNumericField(x);
                        var tdClass = (isMoneyCol || isNumCol) ? ' text-end' : ' text-start';
                        
                        if (isImageColumnKey(x) && typeof cv === 'string' && cv.trim() && (cv.startsWith('http') || cv.startsWith('https'))) {
                            cv = `<img src="${String(cv).replace(/"/g, '&quot;')}" class="table-cell-img" alt="" style="max-height:28px;max-width:36px;object-fit:cover;border-radius:4px;">`;
                        } else if (typeof cv === 'string' && (cv.startsWith('http') || cv.startsWith('https'))) {
                            cv = `<a href="${cv.replace(/"/g, '&quot;')}" target="_blank" class="text-primary" onclick="event.stopPropagation()"><i class="fas fa-paperclip"></i></a>`;
                        } else if (isMoneyCol && cv !== '' && cv != null && !isNaN(cv) && Number(cv) !== 0) {
                            cv = `<span class="money-cell-inline">${formatMoney(cv)}</span>`;
                        } else if (isNumCol && cv !== '' && cv != null && !isNaN(cv)) {
                            cv = `<span class="number-cell-inline">${formatNumber(cv, 1)}</span>`;
                        }
                        
                        html += '<td class="' + tdClass.trim() + '">' + cv + '</td>';
                    }
                });
                // Thêm cột thao tác với nút Chi tiết, Sửa, Xóa
                html += `<td class="text-center">
                    <div class="btn-group" role="group">
                        <button class="btn btn-xs btn-info" onclick="event.stopPropagation(); openChildDetailPopup('${cS}', ${originalCIdx});" title="Chi tiết">
                            <i class="fas fa-eye"></i>
                        </button>
                        <button class="btn btn-xs btn-warning" onclick="event.stopPropagation(); Swal.close(); setTimeout(() => openEditModal(${originalCIdx}, '${cS}'), 300);" title="Sửa">
                            <i class="fas fa-edit"></i>
                        </button>
                        <button class="btn btn-xs btn-danger" onclick="event.stopPropagation(); Swal.close(); setTimeout(() => deleteRow(${originalCIdx}, '${cS}'), 300);" title="Xóa">
                            <i class="fas fa-trash"></i>
                        </button>
                    </div>
                </td>`;
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
        showConfirmButton: false,
        showDenyButton: false,
        showCancelButton: false,
        footer: `<div class="w-100">
                    <div class="d-flex justify-content-between align-items-center gap-2 flex-wrap">
                        <button type="button" class="btn btn-danger px-4 shadow-sm rounded-3" onclick="Swal.close(); setTimeout(() => deleteRow(${idx}, '${s}'), 200);" style="min-width: 100px;">
                            <i class="fas fa-trash-alt me-2"></i> Xóa
                        </button>
                        <div class="d-flex gap-2">
                            ${s === 'User' ? `<button type="button" class="btn btn-info text-white px-4 shadow-sm rounded-3" onclick="Swal.close(); setTimeout(() => openSalaryForEmployee('${r['ID']}', '${(r['Họ và tên'] || r['ID']).replace(/'/g, "\\'")}'), 200);" style="min-width: 100px;">
                                <i class="fas fa-cogs me-2"></i> Cài đặt lương
                            </button>` : ''}
                            <button type="button" class="btn btn-warning text-dark px-4 shadow-sm rounded-3" onclick="Swal.close(); setTimeout(() => openEditModal(${idx}, '${s}'), 200);" style="min-width: 100px;">
                                <i class="fas fa-edit me-2"></i> Sửa
                            </button>
                            <button type="button" class="btn btn-secondary px-4 shadow-sm rounded-3" onclick="Swal.close();" style="min-width: 100px;">
                                <i class="fas fa-times me-2"></i> Đóng
                            </button>
                        </div>
                    </div>
                 </div>`,
        customClass: {
            title: 'detail-modal-title',
            popup: 'rounded-4 shadow-lg',
            footer: 'detail-modal-footer'
        },
        buttonsStyling: false
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

// --- MOBILE BOTTOM NAV FUNCTIONS ---
window.setActiveNavItem = function(element) {
    document.querySelectorAll('.mobile-bottom-nav .nav-item').forEach(item => {
        item.classList.remove('active');
    });
    if (element) element.classList.add('active');
}

window.showRefreshToast = function() {
    const Toast = Swal.mixin({
        toast: true,
        position: 'top',
        showConfirmButton: false,
        timer: 2000,
        timerProgressBar: true
    });
    Toast.fire({
        icon: 'success',
        title: 'Đang làm mới dữ liệu...'
    });
}

// =============================================
// NHẬP CHI PHÍ NHANH (Quick Expense Entry)
// Luồng: Nhập chi tiết → tự tìm/tạo phiếu cha theo ngày+người lập
// ID phiếu cha: MãNhânViên-YYMMDD (VD: CDX001-250527)
// =============================================

window.openQuickExpenseForm = function () {
    // Lấy danh sách loại chi phí từ bảng Drop
    var dropData = (GLOBAL_DATA['Drop'] || []).filter(function (r) {
        var cond = String(r['condition'] || r['Condition'] || '').toLowerCase().trim();
        return cond === 'loaichiphi' && r['Delete'] !== 'X';
    });

    var loaiOptions = '<option value="">-- Chọn loại chi phí --</option>';
    dropData.forEach(function (item) {
        var val = item['label'] || item['Label'] || item['id'] || item['ID'] || '';
        loaiOptions += '<option value="' + String(val).replace(/"/g, '&quot;') + '">' + String(val).replace(/</g, '&lt;') + '</option>';
    });
    
    // Lấy danh sách kho cho dropdown
    var khoData = GLOBAL_DATA['DS_kho'] || [];
    var khoOptions = '<option value="">-- Chọn kho --</option>';
    khoData.forEach(function (k) {
        if (k['Delete'] !== 'X') {
            var khoId = k['ID kho'] || k['ID'] || '';
            var khoName = k['Tên kho'] || k['Ten_Kho'] || khoId;
            khoOptions += '<option value="' + String(khoId).replace(/"/g, '&quot;') + '">' + String(khoName).replace(/</g, '&lt;') + '</option>';
        }
    });

    if (dropData.length === 0) {
        ['Vật tư', 'Xăng dầu', 'Ăn uống', 'Vận chuyển', 'Khác'].forEach(function (v) {
            loaiOptions += '<option value="' + v + '">' + v + '</option>';
        });
    }

    // Dropdown Đơn vị tính từ bảng Drop
    var dvtData = (GLOBAL_DATA['Drop'] || []).filter(function (r) {
        var cond = String(r['condition'] || r['Condition'] || '').toLowerCase().trim();
        return (cond === 'đơn vị tính' || cond === 'don vi tinh' || cond === 'dvt') && r['Delete'] !== 'X';
    });
    var dvtOptions = '<option value="">-- Chọn --</option>';
    dvtData.forEach(function (item) {
        var val = item['label'] || item['Label'] || item['id'] || item['ID'] || '';
        dvtOptions += '<option value="' + String(val).replace(/"/g, '&quot;') + '">' + String(val).replace(/</g, '&lt;') + '</option>';
    });
    if (dvtData.length === 0) {
        ['Cái', 'Bộ', 'Kg', 'Lít', 'M', 'Thùng', 'Chuyến', 'Lần'].forEach(function (v) {
            dvtOptions += '<option value="' + v + '">' + v + '</option>';
        });
    }

    // Dropdown Vật tư từ bảng Vat_tu - hiển thị Tên, lưu ID
    var vattuData = (GLOBAL_DATA['Vat_tu'] || GLOBAL_DATA['VatLieu'] || []).filter(function (r) { return r['Delete'] !== 'X'; });
    var vattuOptions = '<option value="">-- Không chọn --</option>';
    vattuData.forEach(function (item) {
        var id = item['ID vật tư'] || item['Id_vat_tu'] || item['MaVatTu'] || item['id_vat_tu'] || '';
        var ten = item['Tên vật tư'] || item['Ten_vat_tu'] || item['ten_vat_tu'] || '';
        if (id) vattuOptions += '<option value="' + String(id).replace(/"/g, '&quot;') + '" data-ten="' + String(ten).replace(/"/g, '&quot;') + '" data-dvt="' + String(item['Đơn vị tính'] || item['DonViTinh'] || item['DVT'] || item['Don_vi_tinh'] || '').replace(/"/g, '&quot;') + '" data-dongia="' + String(item['Đơn giá'] || item['DonGia'] || item['Don_Gia'] || item['Dongia'] || '').replace(/"/g, '&quot;') + '">' + String(ten || id).replace(/</g, '&lt;') + (ten && id ? ' (' + id + ')' : '') + '</option>';
    });

    // Dropdown Kho từ bảng DS_kho
    var khoData = (GLOBAL_DATA['DS_kho'] || []).filter(function (r) { return r['Delete'] !== 'X'; });
    var khoOptions = '<option value="">-- Không chọn --</option>';
    khoData.forEach(function (item) {
        var id = item['ID kho'] || item['ID_kho'] || '';
        var ten = item['Tên kho'] || item['Ten_kho'] || '';
        if (id || ten) khoOptions += '<option value="' + String(id || ten).replace(/"/g, '&quot;') + '">' + String(ten || id).replace(/</g, '&lt;') + '</option>';
    });

    var today = new Date().toISOString().split('T')[0];
    var userName = CURRENT_USER ? (CURRENT_USER.name || CURRENT_USER['Họ và tên'] || '') : '';

    var html = '<div style="text-align: left;">'
        + '<div class="alert alert-info py-2 mb-3" style="font-size:12px;">'
        + '<i class="fas fa-info-circle me-1"></i> Hệ thống tự động tạo/gom phiếu theo <b>ngày + người lập</b>. Thành tiền = SL × Đơn giá.'
        + '</div>'
        + '<div class="row g-2">'
        // Hàng 1: Ngày chi + Người lập
        + '<div class="col-6">'
        + '  <label class="form-label fw-bold small mb-1">Ngày chi <span class="text-danger">*</span></label>'
        + '  <input type="date" id="qe-date" class="form-control form-control-sm" value="' + today + '">'
        + '</div>'
        + '<div class="col-6">'
        + '  <label class="form-label fw-bold small mb-1">Người lập</label>'
        + '  <input type="text" class="form-control form-control-sm bg-light" value="' + userName + '" disabled>'
        + '</div>'
        // Hàng 2: Loại chi phí + Nội dung
        + '<div class="col-6">'
        + '  <label class="form-label fw-bold small mb-1">Loại chi phí <span class="text-danger">*</span></label>'
        + '  <select id="qe-loai" class="form-select form-select-sm" onchange="toggleWarehouseFields()">' + loaiOptions + '</select>'
        + '</div>'
        + '<div class="col-6">'
        + '  <label class="form-label fw-bold small mb-1">Nội dung</label>'
        + '  <input type="text" id="qe-noidung" class="form-control form-control-sm" placeholder="Mô tả khoản chi...">'
        + '</div>'
        // Hàng 3: Vật tư + Kho
        + '<div class="col-6" id="qe-vattu-group">'
        + '  <label class="form-label fw-bold small mb-1" id="qe-vattu-label">Vật tư <span class="text-muted">(nếu có)</span></label>'
        + '  <select id="qe-vattu" class="form-select form-select-sm">' + vattuOptions + '</select>'
        + '</div>'
        + '<div class="col-6" id="qe-kho-group">'
        + '  <label class="form-label fw-bold small mb-1" id="qe-kho-label">Kho <span class="text-muted">(nếu có)</span></label>'
        + '  <select id="qe-kho" class="form-select form-select-sm">' + khoOptions + '</select>'
        + '</div>'
        // Hàng 4: Số lượng + ĐVT + Đơn giá
        + '<div class="col-4" id="qe-soluong-group">'
        + '  <label class="form-label fw-bold small mb-1" id="qe-soluong-label">Số lượng</label>'
        + '  <input type="number" id="qe-soluong" class="form-control form-control-sm" placeholder="0" min="0" step="any">'
        + '</div>'
        + '<div class="col-4" id="qe-dvt-group">'
        + '  <label class="form-label fw-bold small mb-1" id="qe-dvt-label">ĐVT</label>'
        + '  <select id="qe-dvt" class="form-select form-select-sm">' + dvtOptions + '</select>'
        + '</div>'
        + '<div class="col-4" id="qe-dongia-group">'
        + '  <label class="form-label fw-bold small mb-1" id="qe-dongia-label">Đơn giá</label>'
        + '  <input type="number" id="qe-dongia" class="form-control form-control-sm" placeholder="0" min="0">'
        + '</div>'
        // Hàng 5: Thành tiền (auto)
        + '<div class="col-12">'
        + '  <label class="form-label fw-bold small mb-1">Thành tiền (VNĐ) <span class="text-danger">*</span></label>'
        + '  <input type="number" id="qe-sotien" class="form-control form-control-sm fw-bold text-success" placeholder="Nhập hoặc tự tính = SL × Đơn giá" min="0">'
        + '</div>'
        // Hàng 6: Ghi chú
        + '<div class="col-12">'
        + '  <label class="form-label fw-bold small mb-1">Ghi chú</label>'
        + '  <textarea id="qe-ghichu" class="form-control form-control-sm" rows="2" placeholder="Ghi chú thêm..."></textarea>'
        + '</div>'
        + '</div></div>';

    Swal.fire({
        title: '<i class="fas fa-coins me-2 text-success"></i>Nhập chi phí',
        html: html,
        width: '620px',
        showCancelButton: true,
        confirmButtonText: '<i class="fas fa-save me-1"></i> Lưu chi phí',
        cancelButtonText: 'Hủy',
        confirmButtonColor: '#2E7D32',
        customClass: { popup: 'rounded-4 shadow-lg' },
        didOpen: function () {
            var slEl = document.getElementById('qe-soluong');
            var dgEl = document.getElementById('qe-dongia');
            var ttEl = document.getElementById('qe-sotien');
            var dvtEl = document.getElementById('qe-dvt');
            var vattuEl = document.getElementById('qe-vattu');
            
            // Tạo function toggle warehouse fields
            window.toggleWarehouseFields = function() {
                var loaiEl = document.getElementById('qe-loai');
                var loaiValue = loaiEl ? loaiEl.value.toLowerCase() : '';
                var isMuaVatLieu = loaiValue.includes('vật tư') || loaiValue.includes('vat tu') || loaiValue.includes('vật liệu') || loaiValue.includes('vat lieu');
                
                // Các element cần toggle
                var vattuLabel = document.getElementById('qe-vattu-label');
                var khoLabel = document.getElementById('qe-kho-label');
                var soluongLabel = document.getElementById('qe-soluong-label');
                var dvtLabel = document.getElementById('qe-dvt-label');
                var dongiaLabel = document.getElementById('qe-dongia-label');
                
                var vattuSelect = document.getElementById('qe-vattu');
                var khoSelect = document.getElementById('qe-kho');
                var soluongInput = document.getElementById('qe-soluong');
                var dvtSelect = document.getElementById('qe-dvt');
                var dongiaInput = document.getElementById('qe-dongia');
                
                if (isMuaVatLieu) {
                    // Hiển thị thông báo và thay đổi label thành bắt buộc
                    vattuLabel.innerHTML = 'Vật tư <span class="text-danger">*</span>';
                    khoLabel.innerHTML = 'Kho nhập <span class="text-danger">*</span>';
                    soluongLabel.innerHTML = 'Số lượng <span class="text-danger">*</span>';
                    dvtLabel.innerHTML = 'ĐVT <span class="text-danger">*</span>';
                    dongiaLabel.innerHTML = 'Đơn giá <span class="text-danger">*</span>';
                    
                    // Thêm required attribute
                    vattuSelect.setAttribute('required', 'required');
                    khoSelect.setAttribute('required', 'required');
                    soluongInput.setAttribute('required', 'required');
                    dvtSelect.setAttribute('required', 'required');
                    dongiaInput.setAttribute('required', 'required');
                    
                    // Thêm highlight border
                    [vattuSelect, khoSelect, soluongInput, dvtSelect, dongiaInput].forEach(el => {
                        el.classList.add('border-primary');
                    });
                    
                } else {
                    // Trả về trạng thái ban đầu - tùy chọn
                    vattuLabel.innerHTML = 'Vật tư <span class="text-muted">(nếu có)</span>';
                    khoLabel.innerHTML = 'Kho <span class="text-muted">(nếu có)</span>';
                    soluongLabel.innerHTML = 'Số lượng';
                    dvtLabel.innerHTML = 'ĐVT';
                    dongiaLabel.innerHTML = 'Đơn giá';
                    
                    // Xóa required attribute
                    [vattuSelect, khoSelect, soluongInput, dvtSelect, dongiaInput].forEach(el => {
                        el.removeAttribute('required');
                        el.classList.remove('border-primary');
                    });
                }
            };

            // Auto-tính thành tiền khi nhập SL hoặc Đơn giá
            function calcTotal() {
                var sl = parseFloat(slEl.value) || 0;
                var dg = parseFloat(dgEl.value) || 0;
                if (sl > 0 && dg > 0) ttEl.value = Math.round(sl * dg);
            }
            slEl.addEventListener('input', calcTotal);
            dgEl.addEventListener('input', calcTotal);

            // Auto-fill ĐVT + Đơn giá khi chọn Vật tư
            vattuEl.addEventListener('change', function () {
                var opt = vattuEl.options[vattuEl.selectedIndex];
                if (!opt || !opt.value) return;
                var dvt = opt.getAttribute('data-dvt');
                var dg = opt.getAttribute('data-dongia');
                if (dvt) {
                    // Thử chọn đúng option trong dropdown ĐVT, nếu không có thì set text
                    var found = false;
                    for (var i = 0; i < dvtEl.options.length; i++) {
                        if (dvtEl.options[i].value === dvt || dvtEl.options[i].text === dvt) {
                            dvtEl.selectedIndex = i; found = true; break;
                        }
                    }
                    if (!found) {
                        var newOpt = new Option(dvt, dvt, true, true);
                        dvtEl.add(newOpt);
                    }
                }
                if (dg && parseFloat(dg) > 0) {
                    dgEl.value = dg;
                    calcTotal();
                }
            });
        },
        preConfirm: function () {
            var date = document.getElementById('qe-date').value;
            var loai = document.getElementById('qe-loai').value;
            var noidung = document.getElementById('qe-noidung').value;
            var vattuEl2 = document.getElementById('qe-vattu');
            var vattu = vattuEl2.value;
            var vattuTen = vattu ? (vattuEl2.options[vattuEl2.selectedIndex].getAttribute('data-ten') || vattuEl2.options[vattuEl2.selectedIndex].text) : '';
            var kho = document.getElementById('qe-kho').value;
            var soluong = document.getElementById('qe-soluong').value;
            var dvt = document.getElementById('qe-dvt').value;
            var dongia = document.getElementById('qe-dongia').value;
            var sotien = document.getElementById('qe-sotien').value;
            var ghichu = document.getElementById('qe-ghichu').value;

            if (!date) { Swal.showValidationMessage('Vui lòng chọn ngày chi'); return false; }
            if (!loai) { Swal.showValidationMessage('Vui lòng chọn loại chi phí'); return false; }
            if (!sotien || isNaN(sotien) || Number(sotien) <= 0) { Swal.showValidationMessage('Vui lòng nhập thành tiền hợp lệ (> 0)'); return false; }
            
            // Kiểm tra thêm nếu là mua vật liệu
            var isMuaVatLieu = loai.toLowerCase().includes('vật tư') || loai.toLowerCase().includes('vat tu') || loai.toLowerCase().includes('vật liệu') || loai.toLowerCase().includes('vat lieu');
            if (isMuaVatLieu) {
                if (!vattu) { Swal.showValidationMessage('Vui lòng chọn vật tư cần mua'); return false; }
                if (!kho) { Swal.showValidationMessage('Vui lòng chọn kho nhập vật tư'); return false; }
                if (!soluong || Number(soluong) <= 0) { Swal.showValidationMessage('Vui lòng nhập số lượng > 0'); return false; }
                if (!dvt) { Swal.showValidationMessage('Vui lòng chọn đơn vị tính'); return false; }
                if (!dongia || Number(dongia) <= 0) { Swal.showValidationMessage('Vui lòng nhập đơn giá > 0'); return false; }
            }

            return {
                date: date, loai: loai, noidung: noidung || '',
                vattu: vattu || '', vattuTen: vattuTen || '', kho: kho || '',
                soluong: soluong ? Number(soluong) : '', dvt: dvt || '', dongia: dongia ? Number(dongia) : '',
                sotien: Number(sotien), ghichu: ghichu || ''
            };
        }
    }).then(function (result) {
        if (result.isConfirmed) {
            submitQuickExpense(result.value);
        }
    });
}

async function submitQuickExpense(params) {
    showLoading(true, 'Đang lưu chi phí...');

    try {
        var userId = CURRENT_USER['ID'] || CURRENT_USER['id'] || '';
        if (!userId) throw new Error('Không xác định được mã nhân viên. Vui lòng đăng nhập lại.');

        // Format ngày thành YYMMDD
        var d = new Date(params.date);
        var yy = String(d.getFullYear()).slice(-2);
        var mm = String(d.getMonth() + 1).padStart(2, '0');
        var dd = String(d.getDate()).padStart(2, '0');
        var dateYYMMDD = yy + mm + dd;

        // ID phiếu cha: MãNhânViên-YYMMDD (VD: CDX001-250527)
        var parentId = userId + '-' + dateYYMMDD;

        // === Bước 1: Tìm phiếu cha đã tồn tại ===
        var existingParent = (GLOBAL_DATA['Chiphi'] || []).find(function (r) {
            return r['ID_ChiPhi'] === parentId && r['Delete'] !== 'X';
        });

        // === Bước 2: Nếu chưa có → tạo phiếu cha mới ===
        if (!existingParent) {
            var parentData = {
                'ID_ChiPhi': parentId,
                'NgayChiphi': params.date,
                'NguoiLap': userId,
                'TongSoTien': 0,
                'Trangthai': 'Đang xử lý'
            };
            var parentResult = await callSupabase('insert', 'Chiphi', parentData);
            if (parentResult.status !== 'success') {
                throw new Error(parentResult.message || 'Lỗi khi tạo phiếu chi phí');
            }
        }

        // === Bước 3: Tạo bản ghi chi tiết (con) gắn vào phiếu cha ===
        var childData = {
            'ID_ChiPhi': parentId,
            'Ngaychi': params.date,
            'LoaiChiPhi': params.loai,
            'SoTien': params.sotien,
            'GhiChu': params.ghichu,
            'NguoiLap': userId
        };
        if (params.noidung) childData['NoiDung'] = params.noidung;
        if (params.vattu) {
            childData['MaVatTu'] = params.vattu;
            if (params.vattuTen) childData['TenVatTu'] = params.vattuTen;
        }
        if (params.kho) childData['Ten_kho'] = params.kho;
        if (params.soluong) childData['SoLuong'] = params.soluong;
        if (params.dvt) childData['DonViTinh'] = params.dvt;
        if (params.dongia) childData['DonGia'] = params.dongia;
        var childResult = await callSupabase('insert', 'Chiphichitiet', childData);
        if (childResult.status !== 'success') {
            throw new Error(childResult.message || 'Lỗi khi lưu chi tiết chi phí');
        }
        
        // === Bước 3.5: Tự động tạo phiếu nhập kho nếu là mua vật liệu ===
        var isMuaVatLieu = params.loai.toLowerCase().includes('vật tư') || params.loai.toLowerCase().includes('vat tu') || 
                          params.loai.toLowerCase().includes('vật liệu') || params.loai.toLowerCase().includes('vat lieu');
        
        if (isMuaVatLieu && params.vattu && params.kho && params.soluong && params.dongia) {
            try {
                // Tạo ID phiếu nhập: PN-UserID-YYMMDD-VatTu
                var phieuNhapId = 'PN-' + userId + '-' + dateYYMMDD + '-' + params.vattu.substring(0, 6);
                
                // Kiểm tra xem phiếu nhập đã tồn tại chưa
                var existingPhieuNhap = (GLOBAL_DATA['Phieunhap'] || []).find(function (r) {
                    return r['ID phieu nhap'] === phieuNhapId && r['Delete'] !== 'X';
                });
                
                // Tạo phiếu nhập nếu chưa tồn tại
                if (!existingPhieuNhap) {
                    var phieuNhapData = {
                        'ID phieu nhap': phieuNhapId,
                        'NgayNhap': params.date,
                        'Tên kho': params.kho,
                        'Người nhập': userId,
                        'Diễn giải': 'Tự động từ chi phí: ' + params.loai,
                        'Trangthai': 'Hoàn thành',
                        'TongTien': params.sotien
                    };
                    
                    var phieuNhapResult = await callSupabase('insert', 'Phieunhap', phieuNhapData);
                    if (phieuNhapResult.status !== 'success') {
                        console.warn('Không thể tạo phiếu nhập:', phieuNhapResult.message);
                    }
                }
                
                // Tạo chi tiết nhập kho
                var nhapChiTietData = {
                    'ID phieu nhap': phieuNhapId,
                    'ID vật tư': params.vattu,
                    'SoLuong': params.soluong,
                    'DonGia': params.dongia,
                    'ThanhTien': params.sotien,
                    'DonViTinh': params.dvt,
                    'GhiChu': 'Từ chi phí: ' + (params.noidung || params.loai)
                };
                
                var nhapChiTietResult = await callSupabase('insert', 'NhapChiTiet', nhapChiTietData);
                if (nhapChiTietResult.status === 'success') {
                    console.log('✅ Đã tự động tạo phiếu nhập kho:', phieuNhapId);
                } else {
                    console.warn('Không thể tạo chi tiết nhập:', nhapChiTietResult.message);
                }
                
            } catch (warehouseError) {
                console.warn('Lỗi khi tạo phiếu nhập kho tự động:', warehouseError);
                // Không throw error - để chi phí vẫn được lưu thành công
            }
        }

        // === Bước 4: Tính lại tổng tiền phiếu cha ===
        // Refresh bảng con trước để có dữ liệu mới nhất
        await refreshSingleSheet('Chiphichitiet');

        var allChildren = (GLOBAL_DATA['Chiphichitiet'] || []).filter(function (c) {
            return c['ID_ChiPhi'] === parentId && c['Delete'] !== 'X';
        });
        var tongTien = allChildren.reduce(function (sum, c) {
            return sum + (parseFloat(c['SoTien']) || 0);
        }, 0);

        // Cập nhật tổng tiền vào phiếu cha
        await callSupabase('update', 'Chiphi', { 'TongSoTien': tongTien }, parentId);

        // Refresh bảng cha để hiển thị mới nhất
        await refreshSingleSheet('Chiphi');
        
        // Refresh bảng nhập kho nếu có tạo
        if (isMuaVatLieu && params.vattu && params.kho) {
            await refreshSingleSheet('Phieunhap');
            await refreshSingleSheet('NhapChiTiet');
        }

        showLoading(false);

        var successMessage = 'Chi phí đã được lưu thành công!';
        if (isMuaVatLieu && params.vattu && params.kho) {
            successMessage += '<br><small class="text-success">✅ Đã tự động tạo phiếu nhập kho vật liệu</small>';
        }

        Swal.fire({
            icon: 'success',
            title: 'Đã lưu thành công!',
            html: '<div style="font-size:13px; text-align:left;">'
                + '<p>Phiếu: <b>' + parentId + '</b></p>'
                + '<p>Loại: <b>' + params.loai + '</b>' + (params.noidung ? ' — ' + params.noidung : '') + '</p>'
                + (params.vattuTen ? '<p>Vật tư: <b>' + params.vattuTen + '</b></p>' : '')
                + (params.soluong ? '<p>SL: ' + formatNumber(params.soluong, 1) + (params.dvt ? ' ' + params.dvt : '') + (params.dongia ? ' × ' + formatMoney(params.dongia, true) : '') + '</p>' : '')
                + '<p>Thành tiền: <span class="text-success fw-bold money-display">' + formatMoney(params.sotien, true) + '</span></p>'
                + (isMuaVatLieu && params.vattu && params.kho ? '<div class="alert alert-success py-1 px-2 mt-2" style="font-size:11px;"><i class="fas fa-warehouse me-1"></i><strong>Đã tự động tạo phiếu nhập kho!</strong><br>Vật tư đã được nhập vào kho ' + params.kho + '</div>' : '')
                + '<hr class="my-2"><p>Tổng phiếu ngày này: <span class="text-danger fw-bold fs-5 money-display">' + formatMoney(tongTien, true) + '</span></p>'
                + '</div>',
            confirmButtonText: '<i class="fas fa-plus me-1"></i> Nhập tiếp',
            showCancelButton: true,
            cancelButtonText: 'Đóng',
            confirmButtonColor: '#2E7D32'
        }).then(function (result) {
            if (result.isConfirmed) {
                openQuickExpenseForm(); // Mở lại form để nhập tiếp
            }
        });

    } catch (err) {
        showLoading(false);
        console.error('Quick Expense Error:', err);
        Swal.fire({
            icon: 'error',
            title: 'Lỗi lưu chi phí',
            text: err.message || 'Không thể lưu chi phí. Vui lòng thử lại.'
        });
    }
}
