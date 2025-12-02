// patch-libreoffice.js
const fs = require('fs');
const path = require('path');

// 修复 libreoffice-convert 模块的 bug
function patchLibreOfficeConvert() {
    const modulePath = path.dirname(require.resolve('libreoffice-convert'));
    const indexFile = path.join(modulePath, 'index.js');
    
    try {
        let content = fs.readFileSync(indexFile, 'utf8');
        
        // 修复 includes is not a function 错误
        if (content.includes('(filter ?? "").includes')) {
            content = content.replace(
                /let fmt = !\(filter \?\? ""\)\.includes\(" "\) \?/g,
                'let fmt = !(filter ? String(filter) : "").includes(" ") ?'
            );
            fs.writeFileSync(indexFile, content);
            console.log('成功修复 libreoffice-convert 模块');
        }
    } catch (error) {
        console.warn('无法自动修复模块，将使用备用方案:', error.message);
    }
}

// 应用补丁
patchLibreOfficeConvert();