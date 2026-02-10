// Script để fix tất cả optional chaining operators
const fs = require('fs');

// Đọc file main.js
const filePath = './src/main.js';
let content = fs.readFileSync(filePath, 'utf8');

console.log('🔧 Bắt đầu fix optional chaining operators...');

// Fix các patterns phổ biến
const fixes = [
    // document.getElementById('id')?.value
    {
        pattern: /document\.getElementById\('([^']+)'\)\?\.value/g,
        replacement: '(function(el) { return el ? el.value : \'\'; })(document.getElementById(\'$1\'))'
    },
    // element?.property
    {
        pattern: /(\w+)\?\.\['([^']+)'\]/g,
        replacement: '($1 && $1[\'$2\'] ? $1[\'$2\'] : undefined)'
    },
    // Simple property access with optional chaining
    {
        pattern: /(\w+)\?\.\[([^\]]+)\]/g,
        replacement: '($1 && $1[$2] ? $1[$2] : undefined)'
    },
    // Method calls with optional chaining
    {
        pattern: /(\w+)\?\.([a-zA-Z_][a-zA-Z0-9_]*)\(/g,
        replacement: '($1 && typeof $1.$2 === \'function\' ? $1.$2('
    }
];

let fixCount = 0;

fixes.forEach(fix => {
    const before = content;
    content = content.replace(fix.pattern, fix.replacement);
    const matches = (before.match(fix.pattern) || []).length;
    if (matches > 0) {
        console.log(`✅ Fixed ${matches} instances of pattern: ${fix.pattern}`);
        fixCount += matches;
    }
});

// Manual fixes cho specific cases
const manualFixes = [
    // filterMonth?.value || default
    {
        from: /document\.getElementById\('payroll-month-filter'\)\?\.value/g,
        to: '(function(el) { return el ? el.value : null; })(document.getElementById(\'payroll-month-filter\'))'
    },
    {
        from: /document\.getElementById\('payroll-year-filter'\)\?\.value/g,
        to: '(function(el) { return el ? el.value : null; })(document.getElementById(\'payroll-year-filter\'))'
    },
    {
        from: /document\.getElementById\('payroll-search'\)\?\.value/g,
        to: '(function(el) { return el ? el.value : null; })(document.getElementById(\'payroll-search\'))'
    },
    {
        from: /document\.getElementById\('attendance-month-filter'\)\?\.value/g,
        to: '(function(el) { return el ? el.value : null; })(document.getElementById(\'attendance-month-filter\'))'
    },
    {
        from: /document\.getElementById\('attendance-year-filter'\)\?\.value/g,
        to: '(function(el) { return el ? el.value : null; })(document.getElementById(\'attendance-year-filter\'))'
    },
    {
        from: /document\.getElementById\('attendance-search'\)\?\.value/g,
        to: '(function(el) { return el ? el.value : null; })(document.getElementById(\'attendance-search\'))'
    },
    // empInfo?.property
    {
        from: /empInfo\?\.\['([^']+)'\]/g,
        to: '(empInfo && empInfo[\'$1\'] ? empInfo[\'$1\'] : \'\')'
    },
    // salarySetting?.property  
    {
        from: /salarySetting\?\.\['([^']+)'\]/g,
        to: '(salarySetting && salarySetting[\'$1\'] ? salarySetting[\'$1\'] : undefined)'
    },
    // user?.property
    {
        from: /user\?\.\['([^']+)'\]/g,
        to: '(user && user[\'$1\'] ? user[\'$1\'] : undefined)'
    },
    // existingRecord?.property
    {
        from: /existingRecord\?\.\['([^']+)'\]/g,
        to: '(existingRecord && existingRecord[\'$1\'] ? existingRecord[\'$1\'] : undefined)'
    },
    // GLOBAL_DATA[s]?.[0]
    {
        from: /GLOBAL_DATA\[s\]\?\.\[0\]/g,
        to: '(GLOBAL_DATA[s] && GLOBAL_DATA[s][0] ? GLOBAL_DATA[s][0] : {})'
    },
    // GLOBAL_DATA[childSheet]?.[0]
    {
        from: /GLOBAL_DATA\[childSheet\]\?\.\[0\]/g,
        to: '(GLOBAL_DATA[childSheet] && GLOBAL_DATA[childSheet][0] ? GLOBAL_DATA[childSheet][0] : {})'
    }
];

manualFixes.forEach(fix => {
    const before = content;
    content = content.replace(fix.from, fix.to);
    const matches = (before.match(fix.from) || []).length;
    if (matches > 0) {
        console.log(`✅ Manual fix: ${matches} instances`);
        fixCount += matches;
    }
});

// Ghi lại file
fs.writeFileSync(filePath, content, 'utf8');

console.log(`🎉 Hoàn thành! Đã fix ${fixCount} optional chaining operators.`);
console.log('📝 File src/main.js đã được cập nhật để tương thích với build environment.');