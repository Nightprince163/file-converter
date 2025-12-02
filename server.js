const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const { exec, execSync} = require('child_process');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3001;

// 应用补丁
require('./patch-libreoffice');

// 中间件配置
app.use(cors());
app.use(express.json({ limit: '100mb' })); // 增加JSON解析限制
app.use(express.urlencoded({ extended: true, limit: '100mb' }));

// 创建详细的日志函数
function logToFile(message, level = 'INFO') {
    const timestamp = new Date().toISOString();
    const logMessage = `[${timestamp}] [${level}] ${message}\n`;
    
    // 同时输出到控制台和文件，不能用console.log，避免递归调用
    process.stdout.write(logMessage);
    
    // 写入日志文件（如果是windows系统，路径改为当前目录下的file-converter.log）
    if (process.platform === 'win32') {
        fs.appendFileSync(path.join(__dirname, 'file-converter.log'), logMessage, 'utf8');
        return;
    }
    fs.appendFileSync('/var/log/file-converter.log', logMessage, 'utf8');
}

// 替换所有的 console.log
console.log = function(message) {
    logToFile(message, 'INFO');
};

console.error = function(message) {
    logToFile(message, 'ERROR');
};

// 在应用启动时记录
logToFile('文件转换服务启动', 'INFO');

// 获取磁盘空间信息
function getDiskSpaceInfo(path = '/') {
    try {
        // 方法1: 使用 fs.statfs (Node.js v18.15.0+)
        if (fs.statfsSync) {
            try {
                const stats = fs.statfsSync(path);
                const total = stats.blocks * stats.bsize;
                const free = stats.bfree * stats.bsize;
                const used = total - free;
                
                return {
                    total: formatBytes(total),
                    free: formatBytes(free),
                    used: formatBytes(used),
                    usagePercentage: ((used / total) * 100).toFixed(2),
                    path: path,
                    method: 'fs.statfs'
                };
            } catch (fsError) {
                console.log('fs.statfs 不可用，使用备用方法:', fsError.message);
            }
        }

        // 方法2: 使用 df 命令（跨平台）
        try {
            let command;
            let parseFunction;
            
            if (process.platform === 'win32') {
                // Windows 系统
                command = `wmic logicaldisk where "DeviceID='${path.substring(0, 2)}'" get Size,FreeSpace`;
                parseFunction = parseWindowsDF;
            } else {
                // Linux/Unix 系统
                command = `df -k "${path}"`;
                parseFunction = parseUnixDF;
            }
            
            const output = execSync(command, { encoding: 'utf8' });
            return parseFunction(output, path);
            
        } catch (dfError) {
            console.error('df 命令执行失败:', dfError.message);
            
            // 方法3: 使用 fs.stat（基础方法，只获取当前目录信息）
            const stats = fs.statSync(path);
            const total = 0; // 无法获取总量
            const free = 0;  // 无法获取空闲空间
            
            return {
                total: 'N/A',
                free: 'N/A',
                used: 'N/A',
                usagePercentage: 'N/A',
                path: path,
                warning: '无法获取完整磁盘信息，请检查系统权限',
                method: 'fallback'
            };
        }
        
    } catch (error) {
        return {
            error: error.message,
            path: path,
            method: 'error'
        };
    }
}

// 解析 Unix/Linux df 命令输出
function parseUnixDF(output, path) {
    const lines = output.trim().split('\n');
    if (lines.length < 2) {
        throw new Error('df 命令输出格式不正确');
    }
    
    const dataLine = lines[1].split(/\s+/);
    const total = parseInt(dataLine[1]) * 1024; // 1K blocks to bytes
    const used = parseInt(dataLine[2]) * 1024;
    const free = parseInt(dataLine[3]) * 1024;
    
    return {
        total: formatBytes(total),
        free: formatBytes(free),
        used: formatBytes(used),
        usagePercentage: ((used / total) * 100).toFixed(2),
        path: path,
        method: 'df command'
    };
}

// 解析 Windows df 命令输出
function parseWindowsDF(output, path) {
    const lines = output.trim().split('\n');
    if (lines.length < 2) {
        throw new Error('wmic 命令输出格式不正确');
    }
    
    const dataLine = lines[1].split(/\s+/).filter(Boolean);
    const free = parseInt(dataLine[0]);
    const total = parseInt(dataLine[1]);
    const used = total - free;
    
    return {
        total: formatBytes(total),
        free: formatBytes(free),
        used: formatBytes(used),
        usagePercentage: ((used / total) * 100).toFixed(2),
        path: path,
        method: 'wmic command'
    };
}

// 字节格式化函数
function formatBytes(bytes, decimals = 2) {
    if (bytes === 0) return '0 Bytes';
    if (bytes === 'N/A') return 'N/A';
    
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB', 'TB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(decimals)) + ' ' + sizes[i];
}

// 检查磁盘空间是否足够
function checkDiskSpace(minFreeSpace = 100 * 1024 * 1024) { // 默认 100MB
    const diskInfo = getDiskSpaceInfo(process.cwd());
    
    if (diskInfo.error) {
        return {
            sufficient: false,
            reason: `无法检查磁盘空间: ${diskInfo.error}`,
            info: diskInfo
        };
    }
    
    if (diskInfo.free === 'N/A') {
        return {
            sufficient: true, // 假设足够，因为无法检测
            warning: '无法准确检测磁盘空间',
            info: diskInfo
        };
    }
    
    // 提取数字部分进行比较
    const freeBytes = parseFloat(diskInfo.free) * 
        Math.pow(1024, ['Bytes', 'KB', 'MB', 'GB', 'TB'].indexOf(diskInfo.free.split(' ')[1]));
    
    return {
        sufficient: freeBytes >= minFreeSpace,
        free: diskInfo.free,
        required: formatBytes(minFreeSpace),
        usagePercentage: diskInfo.usagePercentage,
        info: diskInfo
    };
}

// 修复文件权限和路径的函数
function fixEnvironment() {
    const outputDir = 'converted/';
    
    // 确保输出目录存在且有写权限
    if (!fs.existsSync(outputDir)) {
        fs.mkdirSync(outputDir, { recursive: true, mode: 0o755 });
    }
    
    // 检查磁盘空间
    const stats = fs.statSync(outputDir);
    if (stats.size === 0) {
        console.log('输出目录可用');
    }
    
    return outputDir;
}

async function robustConvert(inputPath, outputDir) {
    const absoluteInputPath = path.resolve(inputPath);
    const absoluteOutputDir = path.resolve(outputDir);
    const platform = process.platform;
    
    // 如果是 Linux 系统且使用 flatpak
    if (platform === 'linux') {
        return await convertWithFlatpak(absoluteInputPath, absoluteOutputDir);
    } else {
        // Windows 系统使用 soffice 命令
        return await convertWithSoffice(absoluteInputPath, absoluteOutputDir);
    }
}

// 修改：返回实际生成的PDF文件名
async function convertWithFlatpak(inputPath, outputDir) {
    console.log('使用 flatpak 转换模式...');
    
    // 记录转换前的文件列表
    const filesBefore = fs.readdirSync(outputDir);
    
    try {
        const command = `flatpak run org.libreoffice.LibreOffice --headless --convert-to 'pdf:writer_pdf_Export:Zoom=100' --outdir "${outputDir}" "${inputPath}"`;
        console.log('执行转换命令:', command);
        execSync(command, { encoding: 'utf8', timeout: 120000 });
        
        // 查找新生成的PDF文件
        const generatedPdf = findNewlyGeneratedPdf(inputPath, outputDir, filesBefore);
        if (generatedPdf) {
            return generatedPdf; // 返回生成的PDF文件名
        }
        
        throw new Error('未找到转换后的PDF文件');
    } catch (error) {
        console.log('转换失败:', error.message);
        throw error;
    }
}

// 修改：返回实际生成的PDF文件名
async function convertWithSoffice(inputPath, outputDir) {
    // 记录转换前的文件列表
    const filesBefore = fs.readdirSync(outputDir);
    
    const command = `soffice --headless --convert-to pdf --outdir "${outputDir}" "${inputPath}"`;
    console.log(`使用 soffice 转换: ${command}`);
    
    try {
        execSync(command, { encoding: 'utf8', timeout: 120000 });
        
        // 查找新生成的PDF文件
        const generatedPdf = findNewlyGeneratedPdf(inputPath, outputDir, filesBefore);
        if (generatedPdf) {
            return generatedPdf;
        }
        
        throw new Error('未找到转换后的PDF文件');
} catch (error) {
        throw new Error(`soffice 转换失败: ${error.message}`);
    }
}

// 新增：查找新生成的PDF文件
function findNewlyGeneratedPdf(inputPath, outputDir, filesBefore) {
    const inputFileName = path.basename(inputPath, path.extname(inputPath));
    const expectedPdfName = inputFileName + '.pdf';
    
    // 获取转换后的文件列表
    const filesAfter = fs.readdirSync(outputDir);
    
    // 查找新生成的文件
    const newFiles = filesAfter.filter(file => !filesBefore.includes(file));
    console.log(`新生成的文件: ${newFiles}`);
    
    // 优先查找与输入文件同名的PDF
    if (filesAfter.includes(expectedPdfName) && !filesBefore.includes(expectedPdfName)) {
        return expectedPdfName;
    }
    
    // 查找包含输入文件名的PDF
    const matchingPdfs = newFiles.filter(file => 
        file.endsWith('.pdf') && 
        file.toLowerCase().includes(inputFileName.toLowerCase())
    );
    
    if (matchingPdfs.length > 0) {
        // 返回最匹配的文件（按文件名相似度排序）
        return matchingPdfs.sort((a, b) => {
            const aSimilarity = stringSimilarity(a, inputFileName);
            const bSimilarity = stringSimilarity(b, inputFileName);
            return bSimilarity - aSimilarity;
        })[0];
    }
    
    // 返回任何新生成的PDF文件
    const newPdfs = newFiles.filter(file => file.endsWith('.pdf'));
    if (newPdfs.length > 0) {
        return newPdfs[0];
    }
    
    return null;
}

// 新增：简单的字符串相似度计算（用于文件名匹配）
function stringSimilarity(str1, str2) {
    const longer = str1.length > str2.length ? str1 : str2;
    const shorter = str1.length > str2.length ? str2 : str1;
    
    if (longer.length === 0) return 1.0;
    
    // 检查是否包含
    if (longer.includes(shorter)) return 0.9;
    
    // 简单的前缀匹配
    if (longer.startsWith(shorter) || shorter.startsWith(longer)) return 0.7;
    
    return 0.1;
}

// 检查 LibreOffice 环境
function checkLibreOfficeEnvironment() {
    try {
        // 测试 LibreOffice 是否正常工作
        const result = require('child_process').execSync('soffice --version', { encoding: 'utf8' });
        console.log('LibreOffice 版本:', result.trim());
        return true;
    } catch (error) {
        console.error('LibreOffice 环境检查失败:', error.message);
        return false;
    }
}
// 在应用启动时检查
if (!checkLibreOfficeEnvironment()) {
    console.error('请检查 LibreOffice 安装是否完整');
    // 不退出进程，而是提供降级方案
}

// 确保目录存在
const ensureDirectoryExists = (dir) => {
    if (!fs.existsSync(dir)) {
        fs.mkdirSync(dir, { recursive: true });
    }
};

// 配置 multer 用于文件上传
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        const uploadDir = 'uploads/';
        ensureDirectoryExists(uploadDir);
        cb(null, uploadDir);
    },
    filename: (req, file, cb) => {
        // 生成唯一文件名，避免中文乱码问题
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        const originalName = Buffer.from(file.originalname, 'latin1').toString('utf8');
        const safeName = originalName.replace(/[^a-zA-Z0-9.\-_]/g, '_');
        cb(null, uniqueSuffix + '-' + safeName);
    }
});

// 优化multer配置，处理大文件
const upload = multer({
    storage: storage,
    fileFilter: (req, file, cb) => {
        const allowedMimeTypes = [
            'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            'application/msword',
            'application/vnd.ms-excel',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'application/vnd.ms-powerpoint',
            'application/vnd.openxmlformats-officedocument.presentationml.presentation',
            'application/octet-stream' // 增加二进制流类型
        ];
        
        const allowedExtensions = ['.docx', '.doc', '.xls', '.xlsx', '.ppt', '.pptx'];
        const fileExt = path.extname(file.originalname).toLowerCase();
        
        if (allowedMimeTypes.includes(file.mimetype) || 
            allowedExtensions.includes(fileExt)) {
            cb(null, true);
        } else {
            console.error(`不支持的文件类型: ${fileExt}, 只支持doc, docx, xls, xlsx, ppt, pptx文件`);
            cb(new Error(`不支持的文件类型: ${fileExt}, 只支持doc, docx, xls, xlsx, ppt, pptx文件`), false);
        }
    },
    limits: {
        fileSize: 100 * 1024 * 1024, // 增加到100MB限制
        fieldSize: 100 * 1024 * 1024 // 增加字段大小限制
    }
});

// 流式文件上传处理（替代multer，用于超大文件）
const handleStreamUpload = (req, res) => {
    return new Promise((resolve, reject) => {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        const uploadDir = 'uploads/';
        ensureDirectoryExists(uploadDir);
        
        const originalName = Buffer.from(req.headers['x-file-name'] || 'unknown.docx', 'latin1').toString('utf8');
        const safeName = originalName.replace(/[^a-zA-Z0-9.\-_]/g, '_');
        const filename = uniqueSuffix + '-' + safeName;
        const filePath = path.join(uploadDir, filename);
        
        const fileStream = fs.createWriteStream(filePath);
        let fileSize = 0;
        
        req.on('data', (chunk) => {
            fileSize += chunk.length;
            // 检查文件大小限制
            if (fileSize > 100 * 1024 * 1024) {
                fileStream.destroy();
                fs.unlinkSync(filePath);
                reject(new Error('文件大小超过100MB限制'));
            }
        });
        
        req.pipe(fileStream);
        
        fileStream.on('finish', () => {
            resolve({
                path: filePath,
                originalname: originalName,
                size: fileSize
            });
        });
        
        fileStream.on('error', (error) => {
            reject(error);
        });
        
        req.on('error', (error) => {
            fileStream.destroy();
            reject(error);
        });
    });
};


// 修改转换接口，支持两种上传方式
app.post('/api/convert-docx-to-pdf', async (req, res) => {
    let tempFilePath = '';
    
    // 设置长超时时间
    req.setTimeout(600000); // 10分钟
    res.setTimeout(600000);
    
    try {
        let fileInfo;
        
        // 检查内容类型，决定使用哪种上传方式
        if (req.headers['content-type'] && req.headers['content-type'].includes('application/octet-stream')) {
            // 流式上传
            fileInfo = await handleStreamUpload(req, res);
            tempFilePath = fileInfo.path;
        } else {
            // 传统的multer上传
            await new Promise((resolve, reject) => {
                upload.single('file')(req, res, (err) => {
                    if (err) reject(err);
                    else resolve();
                });
            });
            
            if (!req.file) {
                return res.status(400).json({
                    success: false,
                    error: '没有上传文件或文件格式不正确'
                });
            }
            
            fileInfo = req.file;
            tempFilePath = req.file.path;
        }

        const outputDir = 'converted/';
        
        // 修复环境
        fixEnvironment();
        
        console.log(`开始转换: ${fileInfo.originalname}`);
        console.log(`文件大小: ${fileInfo.size} 字节`);
        console.log(`临时文件路径: ${tempFilePath}`);

        // 检查磁盘空间（至少需要文件大小的2倍空间）
        const diskCheck = checkDiskSpace(fileInfo.size * 2);
        if (!diskCheck.sufficient) {
            throw new Error(`磁盘空间不足。可用空间: ${diskCheck.free}, 需要: ${diskCheck.required}`);
        }

        let outputPath;
        
        try {
            // 使用robustConvert进行转换，现在返回生成的PDF文件名
            outputPdfName = await robustConvert(tempFilePath, outputDir);
            outputPath = path.join(outputDir, outputPdfName);
            
            console.log(`找到转换后的PDF文件: ${outputPdfName}`);
            
        } catch (conversionError) {
            console.error(`转换失败: ${conversionError.message}`);
            throw new Error(`文件转换失败: ${conversionError.message}`);
        }

        // 检查输出文件
        if (!fs.existsSync(outputPath)) {
            throw new Error(`转换后的文件不存在: ${outputPath}`);
        }
        
        const stats = fs.statSync(outputPath);
        if (stats.size === 0) {
            throw new Error('转换后的文件为空');
        }

        console.log(`转换成功: ${outputPath}, 文件大小: ${stats.size} 字节`);

        // 设置响应头，支持大文件下载
        res.setHeader('Content-Type', 'application/pdf');
        res.setHeader('Content-Disposition', `attachment; filename="${path.basename(outputPath)}"`);
        res.setHeader('Content-Length', stats.size);
        
        // 使用流式下载，避免内存问题
        const fileStream = fs.createReadStream(outputPath);
        fileStream.pipe(res);
        
        fileStream.on('end', () => {
            // 清理文件
            cleanupFile(tempFilePath);
            cleanupFile(outputPath);
        });
        
        fileStream.on('error', (error) => {
            console.error('文件流错误:', error);
            cleanupFile(tempFilePath);
            cleanupFile(outputPath);
            res.status(500).json({
                success: false,
                error: '文件下载失败'
            });
        });
        
    } catch (error) {
        console.error('转换过程错误:', error.message);
        
        // 清理临时文件
        if (tempFilePath) {
            cleanupFile(tempFilePath);
        }
        
        res.status(500).json({
            success: false,
            error: '文件转换失败: ' + error.message
        });
    }
});


// 文件清理函数
function cleanupFile(filePath) {
    try {
        if (fs.existsSync(filePath)) {
            fs.unlinkSync(filePath);
        }
    } catch (error) {
        console.error('文件清理错误:', error);
    }
}

// 日志查看端点 - 返回HTML格式
app.get('/log', (req, res) => {
    try {
        // 根据操作系统确定日志文件路径
        let logFilePath;
        if (process.platform === 'win32') {
            logFilePath = path.join(__dirname, 'file-converter.log');
        } else {
            logFilePath = '/var/log/file-converter.log';
        }
        
        console.log(`正在读取日志文件: ${logFilePath}`);
        
        // 检查日志文件是否存在
        if (!fs.existsSync(logFilePath)) {
            const html = `
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>日志文件查看器 - 文件不存在</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .header { background: #f44336; color: white; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        .file-info { background: #e3f2fd; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        .error { background: #ffebee; padding: 15px; border-radius: 5px; color: #c62828; }
        .log-content { background: #fafafa; padding: 15px; border-radius: 5px; font-family: monospace; white-space: pre-wrap; }
        .stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px; margin-bottom: 20px; }
        .stat-item { background: #e8f5e8; padding: 10px; border-radius: 5px; text-align: center; }
        .refresh-btn { background: #2196f3; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; margin: 10px 0; }
        .refresh-btn:hover { background: #1976d2; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📄 日志文件查看器</h1>
        </div>
        <div class="error">
            <h2>❌ 日志文件不存在</h2>
            <p><strong>文件路径:</strong> ${logFilePath}</p>
            <p><strong>建议:</strong> 请确认服务已正常运行并生成日志</p>
        </div>
        <button class="refresh-btn" onclick="location.reload()">🔄 刷新页面</button>
    </div>
</body>
</html>`;
            res.setHeader('Content-Type', 'text/html; charset=utf-8');
            return res.status(404).send(html);
        }
        
        // 获取文件统计信息
        const stats = fs.statSync(logFilePath);
        
        // 支持查询参数：lines（返回最后N行）
        const lines = parseInt(req.query.lines) || 1000; // 默认返回最后1000行
        const maxLines = 5000; // 最大返回行数限制
        
        // 读取日志文件内容
        const logContent = fs.readFileSync(logFilePath, 'utf8');
        const logLines = logContent.split('\n').filter(line => line.trim() !== '');
        
        // 计算实际返回的行数
        const actualLines = Math.min(lines, maxLines, logLines.length);
        const startIndex = Math.max(0, logLines.length - actualLines);
        const recentLogs = logLines.slice(startIndex).join('\n');
        
        // 格式化文件大小
        function formatFileSize(bytes) {
            if (bytes === 0) return '0 Bytes';
            const k = 1024;
            const sizes = ['Bytes', 'KB', 'MB', 'GB'];
            const i = Math.floor(Math.log(bytes) / Math.log(k));
            return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
        }
        
        // 生成HTML响应
        const html = `
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>日志文件查看器 - ${path.basename(logFilePath)}</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .header { background: #2196f3; color: white; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        .file-info { background: #e3f2fd; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        .stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 10px; margin-bottom: 20px; }
        .stat-item { background: #e8f5e8; padding: 10px; border-radius: 5px; text-align: center; }
        .stat-item .label { font-weight: bold; color: #2e7d32; }
        .stat-item .value { font-size: 1.2em; color: #1b5e20; }
        .controls { background: #fff3e0; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        .log-content { background: #263238; color: #eceff1; padding: 15px; border-radius: 5px; font-family: 'Courier New', monospace; white-space: pre-wrap; overflow-x: auto; max-height: 70vh; overflow-y: auto; }
        .log-line { margin: 2px 0; }
        .log-line:hover { background: #37474f; }
        .timestamp { color: #81d4fa; }
        .level-info { color: #4caf50; }
        .level-error { color: #f44336; }
        .level-warning { color: #ff9800; }
        .refresh-btn, .lines-btn { background: #2196f3; color: white; border: none; padding: 10px 20px; border-radius: 5px; cursor: pointer; margin: 5px; }
        .refresh-btn:hover, .lines-btn:hover { background: #1976d2; }
        .lines-selector { display: inline-block; margin-left: 10px; }
        .lines-selector select { padding: 8px; border-radius: 4px; border: 1px solid #ccc; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📄 日志文件查看器 - ${path.basename(logFilePath)}</h1>
        </div>
        
        <div class="file-info">
            <h2>📋 文件信息</h2>
            <p><strong>完整路径:</strong> ${logFilePath}</p>
            <p><strong>操作系统:</strong> ${process.platform}</p>
            <p><strong>最后修改:</strong> ${new Date(stats.mtime).toLocaleString('zh-CN')}</p>
        </div>
        
        <div class="stats">
            <div class="stat-item">
                <div class="label">文件大小</div>
                <div class="value">${formatFileSize(stats.size)}</div>
            </div>
            <div class="stat-item">
                <div class="label">总行数</div>
                <div class="value">${logLines.length}</div>
            </div>
            <div class="stat-item">
                <div class="label">显示行数</div>
                <div class="value">${actualLines}</div>
            </div>
            <div class="stat-item">
                <div class="label">显示范围</div>
                <div class="value">${startIndex + 1} - ${logLines.length}</div>
            </div>
        </div>
        
        <div class="controls">
            <button class="refresh-btn" onclick="location.reload()">🔄 刷新页面</button>
            <div class="lines-selector">
                <label for="lines">显示行数: </label>
                <select id="lines" onchange="changeLines(this.value)">
                    <option value="100" ${lines === 100 ? 'selected' : ''}>最后100行</option>
                    <option value="500" ${lines === 500 ? 'selected' : ''}>最后500行</option>
                    <option value="1000" ${lines === 1000 ? 'selected' : ''}>最后1000行</option>
                    <option value="2000" ${lines === 2000 ? 'selected' : ''}>最后2000行</option>
                    <option value="5000" ${lines === 5000 ? 'selected' : ''}>最后5000行</option>
                </select>
            </div>
        </div>
        
        <div class="log-content" id="logContent">
            ${recentLogs.split('\n').map(line => {
                // 简单的日志级别颜色标记
                let levelClass = 'level-info';
                if (line.includes('[ERROR]')) levelClass = 'level-error';
                else if (line.includes('[WARNING]') || line.includes('[WARN]')) levelClass = 'level-warning';
                
                // 提取时间戳部分
                const timestampMatch = line.match(/\[(.*?)\]/);
                const timestamp = timestampMatch ? timestampMatch[1] : '';
                const contentAfterTimestamp = line.replace(/\[.*?\]\s*/, '');
                
                return `<div class="log-line"><span class="timestamp">[${timestamp}]</span> <span class="${levelClass}">${contentAfterTimestamp}</span></div>`;
            }).join('')}
        </div>
        
        <script>
            function changeLines(lines) {
                const url = new URL(window.location.href);
                url.searchParams.set('lines', lines);
                window.location.href = url.toString();
            }
            
            // 自动滚动到底部
            window.addEventListener('load', function() {
                const logContent = document.getElementById('logContent');
                logContent.scrollTop = logContent.scrollHeight;
            });
            
            // 自动刷新（可选）
            // setInterval(() => location.reload(), 30000); // 每30秒自动刷新
        </script>
    </div>
</body>
</html>`;
        
        res.setHeader('Content-Type', 'text/html; charset=utf-8');
        res.send(html);
        
    } catch (error) {
        console.error('读取日志文件失败:', error);
        const errorHtml = `
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>日志文件查看器 - 错误</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; background-color: #f5f5f5; }
        .container { max-width: 1200px; margin: 0 auto; background: white; padding: 20px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        .header { background: #f44336; color: white; padding: 15px; border-radius: 5px; margin-bottom: 20px; }
        .error { background: #ffebee; padding: 15px; border-radius: 5px; color: #c62828; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>❌ 日志文件查看器 - 错误</h1>
        </div>
        <div class="error">
            <h2>读取日志文件失败</h2>
            <p><strong>错误信息:</strong> ${error.message}</p>
        </div>
    </div>
</body>
</html>`;
        res.setHeader('Content-Type', 'text/html; charset=utf-8');
        res.status(500).send(errorHtml);
    }
});

// 测试转换端点
app.post('/api/test-convert', upload.single('file'), async (req, res) => {
    if (!req.file) {
        return res.status(400).json({ error: '没有上传文件' });
    }

    try {
        // 简单的文件检查
        const stats = fs.statSync(req.file.path);
        res.json({
            success: true,
            message: '文件接收成功',
            fileInfo: {
                originalName: req.file.originalname,
                size: stats.size,
                uploadPath: req.file.path
            }
        });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// 其他端点保持不变...
app.get('/api/info', (req, res) => {
    res.json({
        service: 'DOCX to PDF Converter',
        version: '2.0.0',
        supportedFormats: ['docx', 'doc'],
        maxFileSize: '100MB',
        conversionMethod: 'LibreOffice Command Line'
    });
});

// 错误处理中间件
app.use((error, req, res, next) => {
    if (error instanceof multer.MulterError) {
        if (error.code === 'LIMIT_FILE_SIZE') {
            return res.status(400).json({
                success: false,
                error: '文件太大，请上传小于100MB的文件'
            });
        }
    }
    
    res.status(500).json({
        success: false,
        error: '服务器内部错误: ' + error.message
    });
});

// 404 处理
app.use('*', (req, res) => {
    res.status(404).json({
        success: false,
        error: '接口不存在'
    });
});

// 启动服务
app.listen(PORT, '0.0.0.0', () => {
    console.log(`文件转换服务运行在 http://0.0.0.0:${PORT}`);
    console.log('健康检查: http://localhost:3001/health');
    console.log('服务信息: http://localhost:3001/api/info');
    console.log('日志查看: http://localhost:3001/log'); // 添加日志端点提示
});

process.on('SIGTERM', () => {
    console.log('收到终止信号，正在关闭服务...');
    process.exit(0);
});