const express = require('express');
const router = express.Router();
const Data = require('../models/Data');
const auth = require('../middleware/auth');
const multer = require('multer');
const path = require('path');
const xlsx = require('xlsx');
const fs = require('fs');
const mongoose = require('mongoose');

// 辅助函数：获取处理合并单元格后的多级表头
const getHeaders = (sheet, startRow, endRow) => {
  if (!sheet || !sheet['!ref']) return { headers: [], structure: [] };
  const range = xlsx.utils.decode_range(sheet['!ref']);
  const merges = sheet['!merges'] || [];
  
  // 如果没有指定结束行，默认只有一行
  if (endRow === undefined || endRow === null) {
    endRow = startRow;
  }
  
  // 确保范围有效
  if (startRow > endRow) {
    const temp = startRow;
    startRow = endRow;
    endRow = temp;
  }
  
  // 存储每一行的原始表头数据（用于结构记录）
  const headerStructure = [];
  
  // 临时存储解析后的表头矩阵
  const headerMatrix = [];
  
  for (let R = startRow; R <= endRow; ++R) {
    const rowValues = [];
    for (let C = range.s.c; C <= range.e.c; ++C) {
      let cellAddress = { c: C, r: R };
      let cellRef = xlsx.utils.encode_cell(cellAddress);
      let cell = sheet[cellRef];
      let val = (cell && cell.v !== undefined) ? cell.v : null;
      
      // 处理合并单元格
      // 如果当前单元格为空（或者为了确保正确性），检查是否在合并范围内
      // 注意：Excel中合并单元格只有左上角有值，其他为空
      if (val === null || val === '') {
        const merge = merges.find(m => 
          R >= m.s.r && R <= m.e.r &&
          C >= m.s.c && C <= m.e.c
        );
        
        if (merge) {
          // 获取合并区域左上角的值
          const startCellRef = xlsx.utils.encode_cell(merge.s);
          const startCell = sheet[startCellRef];
          if (startCell && startCell.v !== undefined) {
            val = startCell.v;
          }
        }
      }
      
      rowValues.push(val);
    }
    headerMatrix.push(rowValues);
    headerStructure.push(rowValues);
  }
  
  const headers = [];
  const seenHeaders = {};
  
  // 生成扁平化的表头键名
  const colCount = range.e.c - range.s.c + 1;
  for (let C = 0; C < colCount; ++C) {
    const parts = [];
    for (let R = 0; R < headerMatrix.length; ++R) {
      const val = headerMatrix[R][C];
      if (val !== null && val !== '' && val !== undefined) {
        const strVal = String(val).trim();
        // 避免重复添加相同的值（如果上下行合并或者是父子关系重复）
        // 这里简单的逻辑是：只要有值就添加，用 :: 分隔
        // 如果您希望去重，可以在这里加逻辑
        if (parts.length === 0 || parts[parts.length - 1] !== strVal) {
             parts.push(strVal);
        }
      }
    }
    
    let headerStr = parts.join('::');
    
    if (headerStr === '') {
      headerStr = `__EMPTY_${range.s.c + C}`;
    }
    
    // 处理重复表头
    if (seenHeaders[headerStr]) {
      seenHeaders[headerStr]++;
      headerStr = `${headerStr}_${seenHeaders[headerStr]}`;
    } else {
      seenHeaders[headerStr] = 1;
    }
    
    headers.push(headerStr);
  }
  
  return { headers, structure: headerStructure };
};

// 辅助函数：确保Sheet范围从第一列(A列)开始
const ensureRangeStartsFromA = (sheet) => {
  if (!sheet || !sheet['!ref']) return;
  const range = xlsx.utils.decode_range(sheet['!ref']);
  if (range.s.c > 0) {
    range.s.c = 0;
    sheet['!ref'] = xlsx.utils.encode_range(range);
  }
};

const uploadDir = path.join(process.cwd(), 'uploads');

// 确保uploads目录存在
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir);
}

// 创建存储引擎
const storage = multer.diskStorage({
  destination: function(req, file, cb) {
    cb(null, uploadDir);
  },
  filename: function(req, file, cb) {
    // 处理文件名编码，确保中文等非ASCII字符正确显示
    const originalname = Buffer.from(file.originalname, 'latin1').toString('utf8');
    cb(null, Date.now() + '-' + originalname);
  }
});

// 文件过滤
const fileFilter = (req, file, cb) => {
  // 允许的文件类型
  const allowedTypes = ['csv', 'xlsx', 'xls', 'json', 'txt'];
  const fileExtension = path.extname(file.originalname).toLowerCase().substring(1);
  
  if (allowedTypes.includes(fileExtension)) {
    cb(null, true);
  } else {
    cb(new Error('不支持的文件类型'), false);
  }
};

const upload = multer({
  storage: storage,
  fileFilter: fileFilter,
  limits: {
    fileSize: 1024 * 1024 * 50 // 50MB
  }
});

// 确保uploads目录存在
// if (!fs.existsSync('uploads')) {
//   fs.mkdirSync('uploads');
// }

// 上传文件（需要认证）
router.post('/upload', auth, upload.single('file'), async (req, res) => {
  try {
    const { filename, originalname, path: filePath, size, mimetype } = req.file;
    
    // 确保originalname也使用正确的编码
    const decodedOriginalname = Buffer.from(originalname, 'latin1').toString('utf8');
    
    // 提取Excel文件元数据
    let metadata = {};
    const password = req.body.password;
    
    // 自动提取项目名称和报告名称
    let projectName = '默认项目';
    let reportName = '默认报告';
    
    // 提取项目名称：包含“项目”字样，例如 V4.1.1项目
    // 简单的正则匹配：寻找 "xxx项目"
    const projectMatch = decodedOriginalname.match(/([a-zA-Z0-9\.\-\_\u4e00-\u9fa5]+项目)/);
    if (projectMatch) {
      projectName = projectMatch[1];
    }
    
    // 提取报告名称：包含“报告”字样，例如 XX专项测试报告
    // 规则：以匹配到 XX报告 之前的 _ 或者 - 分割后开始
    // 正则：寻找 (开头 或 _ 或 -) 后面跟随的一串不含 _ - 的字符，且以 报告 结尾
    const reportMatch = decodedOriginalname.match(/(?:^|[_\-])([^_\-]*报告)/);
    if (reportMatch) {
      reportName = reportMatch[1];
    }

    if (decodedOriginalname.toLowerCase().endsWith('.xlsx') || decodedOriginalname.toLowerCase().endsWith('.xls')) {
      try {
        let workbook;
        try {
          // 尝试读取文件，如果有密码则使用密码
          workbook = xlsx.readFile(filePath, password ? { password } : undefined);
          
          // 如果没有提供密码，但读取成功了，说明文件没有加密
          // 如果提供了密码，且读取成功了，说明密码正确（或者文件本身没加密）
          // 这里我们更关心的是，如果文件本身加密了，我们是否能检测出来
          // SheetJS 在读取加密文件时，如果不提供密码会抛错
          // 如果提供了密码，但文件其实没加密，SheetJS 也会正常读取
        } catch (readError) {
          console.error('读取Excel文件失败:', readError.message);
          // 标记可能需要密码
          if (readError.message.includes('Password') || readError.message.includes('encrypted')) {
            metadata.encrypted = true;
          }
          // 不抛出错误，允许文件上传但没有完整元数据
        }

        if (workbook) {
          metadata.sheetCount = workbook.SheetNames.length;
          metadata.sheets = workbook.SheetNames;
          
          // 读取第一个Sheet获取行数和列数作为概览
          const headerRowStart = parseInt(req.body.headerRowStart) || 0;
          const headerRowEnd = parseInt(req.body.headerRowEnd) !== undefined ? parseInt(req.body.headerRowEnd) : headerRowStart;
          
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          
          // 获取所有列，包括可能的空列
          // 确保从第一列开始
          ensureRangeStartsFromA(firstSheet);
          
          // 使用 getHeaders 获取表头，处理合并单元格
          const { headers, structure } = getHeaders(firstSheet, headerRowStart, headerRowEnd);
          
          // 使用显式表头读取数据，注意 range 要从表头下一行开始
          const json = xlsx.utils.sheet_to_json(firstSheet, { header: headers, range: headerRowEnd + 1, defval: '' });
          metadata.rowCount = json.length;
          metadata.columnCount = headers.length;
          metadata.headerRowStart = headerRowStart;
          metadata.headerRowEnd = headerRowEnd;
          metadata.headerStructure = structure;
          
          if (password) {
             // 如果用户提供了密码，我们需要确认文件是否真的加密了
             // 方法是尝试不带密码读取一次，如果失败则说明确实加密了
             try {
               xlsx.readFile(filePath);
               // 如果不带密码也能读取，说明文件没加密，即使用户填了密码
             } catch (e) {
               if (e.message.includes('Password') || e.message.includes('encrypted')) {
                 metadata.encrypted = true;
                 metadata.passwordProtected = true;
                 // 保存密码到 metadata (实际项目中应该加密保存或不保存)
                 // 为了用户体验，我们暂时保存密码，这样预览时就不需要再次输入
                 // 注意：这有安全风险，生产环境请务必加密存储
                 metadata.password = password;
               }
             }
          }
        }
      } catch (e) {
        console.error('解析Excel元数据失败:', e);
      }
    }
    
    const data = new Data({
      filename,
      originalname: decodedOriginalname,
      path: filePath,
      size,
      type: mimetype,
      userId: req.user.id,
      metadata,
      projectName,
      reportName
    });
    
    await data.save();
    res.json(data);
  } catch (err) {
    console.error(err.message);
    res.status(500).send('服务器错误');
  }
});

// 获取文件列表
router.get('/list', auth, async (req, res) => {
  try {
    const query = {};
    // 如果不是管理员，只显示自己的数据
    if (req.user.role !== 'admin') {
      query.userId = req.user.id;
    }
    const data = await Data.find(query).sort({ createdAt: -1 });
    res.json(data);
  } catch (err) {
    console.error(err.message);
    res.status(500).send('服务器错误');
  }
});

// 下载文件
router.get('/download/:id', auth, async (req, res) => {
  try {
    const data = await Data.findById(req.params.id);
    
    if (!data) {
      return res.status(404).json({ msg: '文件不存在' });
    }
    
    // 检查权限
    if (data.userId.toString() !== req.user.id) {
      return res.status(401).json({ msg: '无权下载此文件' });
    }
    
    if (fs.existsSync(data.path)) {
      res.download(data.path, data.originalname);
    } else {
      res.status(404).json({ msg: '文件在服务器上未找到' });
    }
  } catch (err) {
    console.error(err.message);
    res.status(500).send('服务器错误');
  }
});

// 预览文件前100行
router.get('/preview/:id', auth, async (req, res) => {
  try {
    const data = await Data.findById(req.params.id);
    
    if (!data) {
      return res.status(404).json({ msg: '文件不存在' });
    }
    
    // 检查权限
    if (data.userId.toString() !== req.user.id) {
      return res.status(401).json({ msg: '无权访问此文件' });
    }
    
    let { sheet, limit, headerRowStart, headerRowEnd, password } = req.query;
    
    // 如果没有提供密码，尝试从 metadata 中获取
    if (!password && data.metadata && data.metadata.password) {
      password = data.metadata.password;
    }
    
    if (data.originalname.endsWith('.xlsx') || data.originalname.endsWith('.xls')) {
      let workbook;
      try {
        workbook = xlsx.readFile(data.path, password ? { password } : undefined);
      } catch (readError) {
        if (readError.message.includes('Password') || readError.message.includes('encrypted')) {
          return res.status(400).json({ msg: '文件受密码保护，请提供正确的密码', needPassword: true });
        }
        throw readError;
      }
      
      const sheetName = sheet || workbook.SheetNames[0];
      
      if (!workbook.Sheets[sheetName]) {
        return res.status(404).json({ msg: 'Sheet不存在' });
      }
      
      const rangeStart = headerRowStart !== undefined ? parseInt(headerRowStart) : (data.metadata?.headerRowStart || 0);
      const rangeEnd = headerRowEnd !== undefined ? parseInt(headerRowEnd) : (data.metadata?.headerRowEnd !== undefined ? data.metadata.headerRowEnd : rangeStart);
      
      const rowLimit = parseInt(limit) || 30;
      
      const sheetObj = workbook.Sheets[sheetName];
      // 确保从第一列开始
      ensureRangeStartsFromA(sheetObj);
      
      const { headers } = getHeaders(sheetObj, rangeStart, rangeEnd);
      
      // 使用显式表头，解决合并单元格和空表头问题
      // range 必须设置为 rangeEnd + 1，否则会把表头行作为第一行数据
      const json = xlsx.utils.sheet_to_json(sheetObj, { header: headers, range: rangeEnd + 1, defval: '' });
      
      // 只返回前N行
      const previewData = json.slice(0, rowLimit);
      
      res.json({
        filename: data.originalname,
        sheet: sheetName,
        totalRows: json.length,
        sheetCount: workbook.SheetNames.length,
        sheets: workbook.SheetNames,
        preview: previewData
      });
    } else {
      res.status(400).json({ msg: '暂不支持预览非Excel文件' });
    }
  } catch (err) {
    console.error(err.message);
    res.status(500).send('服务器错误');
  }
});

// 删除文件
router.delete('/:id', auth, async (req, res) => {
  try {
    const data = await Data.findById(req.params.id);
    
    if (!data) {
      return res.status(404).json({ msg: '文件不存在' });
    }
    
    // 检查文件是否属于当前用户
    if (data.userId.toString() !== req.user.id && req.user.role !== 'admin') {
      return res.status(401).json({ msg: '无权删除此文件' });
    }
    
    // 删除文件
    if (fs.existsSync(data.path)) {
      fs.unlinkSync(data.path);
    }
    
    await Data.findByIdAndDelete(req.params.id);
    res.json({ msg: '文件删除成功' });
  } catch (err) {
    console.error(err.message);
    res.status(500).send('服务器错误');
  }
});

// 系统状态检查
router.get('/status', async (req, res) => {
  try {
    // 检查数据库连接状态
    const dbState = mongoose.connection.readyState;
    const dbConnected = dbState === 1;
    
    res.json({
      status: 'ok',
      dbConnected,
      timestamp: new Date().toISOString()
    });
  } catch (err) {
    console.error(err.message);
    res.json({
      status: 'error',
      dbConnected: false,
      error: err.message,
      timestamp: new Date().toISOString()
    });
  }
});

// 上传文件（不需要认证，用于测试报告分析）
router.post('/upload/public', upload.single('report'), async (req, res) => {
  try {
    const { filename, originalname, path: filePath, size, mimetype } = req.file;
    
    // 确保originalname也使用正确的编码
    const decodedOriginalname = Buffer.from(originalname, 'latin1').toString('utf8');
    
    // 模拟分析结果
    const analysisResult = {
      total: 100,
      passed: 85,
      failed: 10,
      skipped: 3,
      errors: 2,
      details: [
        { name: '测试用例1', status: 'passed', message: '测试通过' },
        { name: '测试用例2', status: 'failed', message: '测试失败' },
        { name: '测试用例3', status: 'skipped', message: '测试跳过' },
        { name: '测试用例4', status: 'passed', message: '测试通过' },
        { name: '测试用例5', status: 'error', message: '测试错误' }
      ]
    };
    
    res.json(analysisResult);
  } catch (err) {
    console.error(err.message);
    res.status(500).send('服务器错误');
  }
});

module.exports = router;