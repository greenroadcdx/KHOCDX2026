// Script để fix vấn đề hiển thị dữ liệu
console.log("🔧 Bắt đầu fix vấn đề dữ liệu...");

// 1. Clear toàn bộ cache
function clearAllCache() {
    console.log("🧹 Đang xóa cache...");
    localStorage.clear();
    sessionStorage.clear();
    console.log("✅ Đã xóa cache!");
}

// 2. Force reload từ Supabase
async function forceReloadData() {
    console.log("🔄 Đang tải lại dữ liệu từ Supabase...");
    
    // Kiểm tra kết nối Supabase
    if (typeof supabaseClient === 'undefined') {
        console.error("❌ Supabase client không tồn tại!");
        return;
    }
    
    try {
        // Test với bảng Chiphi (Chi phí)
        const { data, error } = await supabaseClient.from('Chiphi').select('*');
        
        if (error) {
            console.error("❌ Lỗi Supabase:", error);
            return;
        }
        
        console.log(`✅ Tải được ${data.length} bản ghi từ Supabase`);
        console.log("📋 Sample data:", data.slice(0, 3));
        
        // Cập nhật GLOBAL_DATA
        if (typeof GLOBAL_DATA !== 'undefined') {
            GLOBAL_DATA['Chiphi'] = data;
            console.log("✅ Đã cập nhật GLOBAL_DATA");
            
            // Render lại table nếu đang ở trang Chi phí
            if (typeof CURRENT_SHEET !== 'undefined' && CURRENT_SHEET === 'Chiphi') {
                if (typeof renderTable === 'function') {
                    renderTable(data);
                    console.log("✅ Đã render lại table");
                }
            }
        }
        
    } catch (err) {
        console.error("❌ Lỗi:", err);
    }
}

// 3. Debug function
function debugCurrentState() {
    console.log("🔍 Debug thông tin hiện tại:");
    console.log("- CURRENT_SHEET:", typeof CURRENT_SHEET !== 'undefined' ? CURRENT_SHEET : 'undefined');
    console.log("- GLOBAL_DATA keys:", typeof GLOBAL_DATA !== 'undefined' ? Object.keys(GLOBAL_DATA) : 'undefined');
    
    if (typeof GLOBAL_DATA !== 'undefined' && GLOBAL_DATA['Chiphi']) {
        console.log("- Chiphi records:", GLOBAL_DATA['Chiphi'].length);
        console.log("- Sample Chiphi:", GLOBAL_DATA['Chiphi'].slice(0, 2));
    }
    
    // Kiểm tra localStorage
    console.log("- LocalStorage items:", localStorage.length);
    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key.includes('cache') || key.includes('GLOBAL')) {
            const size = localStorage.getItem(key).length;
            console.log(`  - ${key}: ${(size/1024).toFixed(2)} KB`);
        }
    }
}

// 4. Full fix function
async function fullFix() {
    console.log("🚀 Bắt đầu full fix...");
    
    // Bước 1: Clear cache
    clearAllCache();
    
    // Bước 2: Đợi một chút
    await new Promise(resolve => setTimeout(resolve, 1000));
    
    // Bước 3: Reload trang
    console.log("🔄 Reload trang...");
    location.reload();
}

// Export functions to global scope
window.clearAllCache = clearAllCache;
window.forceReloadData = forceReloadData;
window.debugCurrentState = debugCurrentState;
window.fullFix = fullFix;

console.log(`
🔧 Fix script loaded! Các lệnh có sẵn:
- clearAllCache() - Xóa toàn bộ cache
- forceReloadData() - Tải lại dữ liệu từ Supabase
- debugCurrentState() - Debug thông tin hiện tại
- fullFix() - Fix toàn bộ (clear cache + reload)

Thử chạy: fullFix()
`);