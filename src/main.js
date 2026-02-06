// --- CẤU HÌNH API ---
// Thay thế URL này bằng URL của Web App sau khi bạn Deploy (Triển khai) trong Google Apps Script
const API_URL = 'https://script.google.com/macros/s/AKfycbyPf9oYjWdG1VROFGTNpXdV1zw86BSCnPnsbdP9yImHeonUDbL1PxY1uAqSR5jsTBoUxw/exec';

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
            { id: 'phieunhap', sheet: 'PhieuNhapXuat', icon: 'fas fa-arrow-circle-down', title: 'Phiếu Nhập kho' },
            { id: 'phieuxuat', sheet: 'PhieuNhapXuat', icon: 'fas fa-arrow-circle-up', title: 'Phiếu Xuất kho' },
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

const COLUMN_MAP = { 'ID': 'Mã', 'Họ và tên': 'Họ tên', 'Ten_Kho': 'Tên kho', 'TenVatTu': 'Vật tư', 'SoLuong': 'SL', 'DonGia': 'Đơn giá', 'ThanhTien': 'Thành tiền', 'NguoiLap': 'Người lập', 'NgayLap': 'Ngày', 'TongSoTien': 'Tổng tiền', 'NoiDung': 'Nội dung', 'NgayChiPhi': 'Ngày chi', 'SoTien': 'Số tiền' };

const RELATION_CONFIG = {
    'PhieuNhapXuat': { child: 'PNXChiTiet', foreignKey: 'ID_Phieu', title: 'CHI TIẾT PHIẾU' },
    'Phieuchuyenkho': { child: 'Chuyenkhochitiet', foreignKey: 'ID Chkho', title: 'CHI TIẾT CHUYỂN' },
    'Chiphi': { child: 'Chiphichitiet', foreignKey: 'ID_ChiPhi', title: 'CHI TIẾT CHI' }
};

// --- HÀM GỌI API GAS ---
async function callGAS(action, ...args) {
    if (API_URL === 'YOUR_GAS_WEB_APP_URL_HERE') {
        throw new Error('Bạn chưa cấu hình API_URL trong src/main.js');
    }

    // Sử dụng tham số URL cho các tham số GET đơn giản
    const url = new URL(API_URL);
    url.searchParams.set('action', action);
    if (args.length > 0) {
        url.searchParams.set('args', JSON.stringify(args));
    }

    try {
        console.log('Calling API:', url.toString());
        const response = await fetch(url.toString(), {
            method: 'GET',
            mode: 'cors',
            cache: 'no-cache',
            redirect: 'follow'
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const result = await response.json();
        return result;
    } catch (error) {
        console.error('API Error:', error);
        if (error.message === 'Failed to fetch') {
            throw new Error('Không thể kết nối API. Vui lòng kiểm tra: 1. Đã Deploy Apps Script là "Anyone". 2. Bạn đã cấp quyền truy cập Sheet trong Apps Script Editor.');
        }
        throw error;
    }
}

// --- KHỞI TẠO ---
document.addEventListener('DOMContentLoaded', function () {
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
    if (!res || res.error) { showLoading(false); alert(res ? res.error : 'Lỗi đăng nhập'); return; }
    CURRENT_USER = res.user;
    document.getElementById('user-display-name').innerText = CURRENT_USER.name;

    // Gọi hàm tải toàn bộ dữ liệu
    callGAS('getInitialData')
        .then(d => {
            showLoading(false);
            if (d.status === 'error') { alert(d.message); return; }
            GLOBAL_DATA = d;
            document.getElementById('login-container').classList.add('d-none');
            document.getElementById('app-container').classList.remove('d-none');
            renderDashboard();
        })
        .catch(handleError);
}

// --- HÀM ĐĂNG XUẤT MỀM (Soft Logout) ---
window.doLogout = function () {
    showLoading(true);
    setTimeout(function () {
        CURRENT_USER = null;
        GLOBAL_DATA = {};
        document.getElementById('form-login').reset();
        document.getElementById('app-container').classList.add('d-none');
        document.getElementById('login-container').classList.remove('d-none');
        if (window.innerWidth < 992 && sidebarBS) sidebarBS.hide();
        var modalEl = document.getElementById('dataModal');
        var modalInstance = bootstrap.Modal.getInstance(modalEl);
        if (modalInstance) modalInstance.hide();
        showLoading(false);
        const Toast = Swal.mixin({
            toast: true, position: 'top-end', showConfirmButton: false, timer: 3000, timerProgressBar: true
        });
        Toast.fire({ icon: 'success', title: 'Đã đăng xuất an toàn' });
    }, 500);
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
        phieu: (GLOBAL_DATA['PhieuNhapXuat'] || []).length
    };

    var kpiHTML = `
    <h5 class="fw-bold text-success mb-3"><i class="fas fa-chart-pie me-2"></i>TỔNG QUAN</h5>
    <div class="row g-3 mb-4">
        <div class="col-12 col-sm-6 col-xl-3">${kpiCard('Nhân sự', stats.user, 'fas fa-users', '#1E88E5')}</div>
        <div class="col-12 col-sm-6 col-xl-3">${kpiCard('Vật tư', stats.vattu, 'fas fa-box', '#43A047')}</div>
        <div class="col-12 col-sm-6 col-xl-3">${kpiCard('Kho bãi', stats.kho, 'fas fa-warehouse', '#FB8C00')}</div>
        <div class="col-12 col-sm-6 col-xl-3">${kpiCard('Phiếu NX', stats.phieu, 'fas fa-file-invoice', '#E53935')}</div>
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

function kpiCard(l, v, i, c) {
    return `<div class="kpi-card"><div><h2 class="mb-0 fw-bold text-dark">${v}</h2><div class="text-muted small fw-bold text-uppercase mt-1">${l}</div></div><div class="kpi-icon shadow-sm" style="background-color: ${c};"><i class="${i}"></i></div></div>`;
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
            document.getElementById('filter-area').innerHTML = `<div class="row g-2"><div class="col-md-4"><div class="input-group"><span class="input-group-text bg-white border-end-0"><i class="fas fa-search text-muted"></i></span><input type="text" id="search-box" class="form-control border-start-0 ps-0" placeholder="Tìm kiếm dữ liệu..." onkeyup="debounceSearch()"></div></div></div>`;

            renderTable(GLOBAL_DATA[CURRENT_SHEET] || []);
        }
    }
    if (window.innerWidth < 992 && sidebarBS) sidebarBS.hide();
}

// --- TABLE RENDER ---
function renderTable(data) {
    var tb = document.querySelector('#data-table tbody');
    var th = document.querySelector('#data-table thead');
    th.innerHTML = ''; tb.innerHTML = '';

    var activeData = data.filter(r => r['Delete'] != 'X');

    if (CURRENT_SHEET === 'PhieuNhapXuat') {
        if (CURRENT_MENU_ID === 'phieunhap') activeData = activeData.filter(r => String(r['ID']).startsWith('PN') || String(r['Loai']).includes('Nhập'));
        if (CURRENT_MENU_ID === 'phieuxuat') activeData = activeData.filter(r => String(r['ID']).startsWith('PX') || String(r['Loai']).includes('Xuất'));
    }

    if (activeData.length === 0) { tb.innerHTML = '<tr><td colspan="100%" class="text-center p-5 text-muted fst-italic">Không có dữ liệu hiển thị</td></tr>'; return; }

    var keys = Object.keys(activeData[0]);
    var skip = ['Delete', 'Update', 'App_pass', 'Who delete'];
    var trH = document.createElement('tr');
    keys.forEach(k => { if (!skip.includes(k)) { var t = document.createElement('th'); t.innerText = COLUMN_MAP[k] || k; trH.appendChild(t); } });
    trH.appendChild(document.createElement('th'));
    th.appendChild(trH);

    activeData.forEach((r) => {
        var originalIdx = data.indexOf(r);
        var tr = document.createElement('tr');
        tr.onclick = (e) => { if (!e.target.closest('button')) showRowDetail(r, CURRENT_SHEET, originalIdx); };
        tr.style.cursor = 'pointer';
        keys.forEach(k => {
            if (!skip.includes(k)) {
                var td = document.createElement('td'); var v = r[k];
                if (k.toUpperCase().includes('NGAY') || k.toUpperCase().includes('DATE')) v = formatSafeDate(v);
                if ((k.toUpperCase().includes('GIA') || k.toUpperCase().includes('TIEN')) && !isNaN(v) && v !== '') { v = Number(v).toLocaleString('vi-VN'); td.className = 'text-end fw-bold text-success'; }
                td.innerText = v; tr.appendChild(td);
            }
        });
        var tdA = document.createElement('td'); tdA.className = 'text-center';
        tdA.innerHTML = `
            <div class="btn-group">
                <button class="btn btn-sm btn-light border text-primary shadow-sm" onclick="openEditModal(${originalIdx})" title="Sửa"><i class="fas fa-pen"></i></button>
                <button class="btn btn-sm btn-light border text-danger shadow-sm" onclick="deleteRow(${originalIdx})" title="Xóa"><i class="fas fa-trash"></i></button>
            </div>`;
        tr.appendChild(tdA); tb.appendChild(tr);
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

function showLoading(s) {
    var overlay = document.getElementById('loading-overlay');
    if (s) {
        overlay.style.display = 'flex';
        if (LOADING_TIMEOUT) clearTimeout(LOADING_TIMEOUT);
        LOADING_TIMEOUT = setTimeout(function () { overlay.style.display = 'none'; }, 30000);
    } else {
        overlay.style.display = 'none';
        if (LOADING_TIMEOUT) clearTimeout(LOADING_TIMEOUT);
    }
}

function handleError(e) { showLoading(false); alert('Lỗi hệ thống: ' + e.message); }

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
            callGAS('getInitialData').then(d => {
                showLoading(false); GLOBAL_DATA = d;
                renderTable(GLOBAL_DATA[CURRENT_SHEET] || []);
            });
        } else { showLoading(false); alert(r.message); }
    }).catch(handleError);
}

window.deleteRow = function (i) {
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
            callGAS('deleteData', CURRENT_SHEET, i).then(r => {
                if (r.status == 'success') {
                    callGAS('getInitialData').then(d => {
                        showLoading(false); GLOBAL_DATA = d;
                        renderTable(GLOBAL_DATA[CURRENT_SHEET] || []);
                    });
                } else { showLoading(false); alert(r.message); }
            }).catch(handleError);
        }
    });
}

function showRowDetail(r, s, idx) {
    var html = '<div class="row text-start">'; var skip = ['Delete', 'Update', 'App_pass'];
    for (var k in r) { if (!skip.includes(k) && r[k]) { var val = r[k]; if (k.includes('Ngay')) val = formatSafeDate(val); html += `<div class="col-6 mb-3"><small class="text-muted fw-bold d-block text-uppercase" style="font-size:0.75rem">${COLUMN_MAP[k] || k}</small><span class="fw-medium">${val}</span></div>`; } }
    html += '</div>';

    if (RELATION_CONFIG[s]) {
        var cf = RELATION_CONFIG[s]; var cS = cf.child; var fk = cf.foreignKey; var pid = Object.values(r)[0];
        var childs = (GLOBAL_DATA[cS] || []).filter(c => c[fk] == pid && c['Delete'] != 'X');
        if (childs.length > 0) {
            html += `<h6 class="text-success border-top pt-3 mt-2 fw-bold">${cf.title} (${childs.length})</h6><div class="table-responsive"><table class="table table-sm table-bordered small"><thead><tr>`;
            var ck = Object.keys(childs[0]);
            ck.forEach(x => { if (!skip.includes(x) && x != fk) html += `<th>${COLUMN_MAP[x] || x}</th>`; });
            html += '</tr></thead><tbody>';
            childs.forEach(cr => { html += '<tr>'; ck.forEach(x => { if (!skip.includes(x) && x != fk) html += `<td>${cr[x]}</td>`; }); html += '</tr>'; });
            html += '</tbody></table></div>';
        }
    }
    Swal.fire({ title: 'Chi tiết bản ghi', html: html, width: '800px', showCloseButton: true, showConfirmButton: false });
}

window.viewProfile = function () {
    if (!CURRENT_USER) return;
    var html = `<div class="text-center mb-3"><div class="avatar-circle mx-auto bg-success text-white" style="width:80px;height:80px;font-size:2rem"><i class="fas fa-user"></i></div><h4 class="mt-2">${CURRENT_USER.name}</h4><span class="badge bg-secondary">${CURRENT_USER.role || 'Nhân viên'}</span></div>`;
    Swal.fire({ html: html, showConfirmButton: false });
}
