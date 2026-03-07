const express = require('express');
const router = express.Router();
const Data = require('../models/Data');
const auth = require('../middleware/auth');
const multer = require('multer');
const path = require('path');
const xlsx = require('xlsx');
const fs = require('fs');

// 创建存储引擎
const storage = multer.diskStorage({
  destination: function(req, file, cb) {
    cb(null, 'uploads/');
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
if (!fs.existsSync('uploads')) {
  fs.mkdirSync('uploads');
}

// 上传文件
router.post('/upload', auth, upload.single('file'), async (req, res) => {
  try {
    const { filename, originalname, path: filePath, size, mimetype } = req.file;
    
    // 确保originalname也使用正确的编码
    const decodedOriginalname = Buffer.from(originalname, 'latin1').toString('utf8');
    
    // 提取Excel文件元数据
    let metadata = {};
    const password = req.body.password;
    
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
          const headerRow = parseInt(req.body.headerRow) || 0;
          const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
          const json = xlsx.utils.sheet_to_json(firstSheet, { range: headerRow });
          metadata.rowCount = json.length;
          metadata.columnCount = json.length > 0 ? Object.keys(json[0]).length : 0;
          metadata.headerRow = headerRow;
          
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
      metadata
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
    const data = await Data.find({ userId: req.user.id }).sort({ createdAt: -1 });
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
    
    const { sheet, limit, headerRow, password } = req.query;
    
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
      
      const range = req.query.headerRow !== undefined ? parseInt(req.query.headerRow) : (data.metadata?.headerRow || 0);
      const rowLimit = parseInt(limit) || 30;
      
      const json = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName], { range });
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
    if (data.userId.toString() !== req.user.id) {
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

module.exports = router;