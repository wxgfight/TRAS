const express = require('express');
const router = express.Router();
const Data = require('../models/Data');
const Comparison = require('../models/Comparison');
const auth = require('../middleware/auth');
const xlsx = require('xlsx');
const path = require('path');
const pd = require('node-pandas');
const { v4: uuidv4 } = require('uuid');
const fs = require('fs');

// 智能主键推荐函数
function recommendPrimaryKeys(columns) {
  const keywordPatterns = [
    /id/i, /编号/i, /序号/i, /名称/i, /code/i, /key/i,
    /test.*id/i, /case.*id/i, /item.*id/i
  ];
  
  const recommendedKeys = [];
  
  columns.forEach(column => {
    for (const pattern of keywordPatterns) {
      if (pattern.test(column)) {
        recommendedKeys.push(column);
        break;
      }
    }
  });
  
  return recommendedKeys.length > 0 ? recommendedKeys : [columns[0]];
}

// 生成主键值函数
function generatePrimaryKeyValue(row, primaryKeys) {
  if (!Array.isArray(primaryKeys)) {
    primaryKeys = [primaryKeys];
  }
  return primaryKeys.map(key => row[key] || '').join('|');
}

// 空值处理函数
function handleNullValues(value, strategy = 'ignore') {
  if (value === null || value === undefined || value === '') {
    switch (strategy) {
      case 'ignore':
        return null;
      case 'treat_as_empty':
        return '';
      case 'treat_as_zero':
        return 0;
      default:
        return null;
    }
  }
  return value;
}

// 数据类型转换函数
function convertDataType(value) {
  // 尝试转换为数字
  if (!isNaN(value) && value !== '') {
    return Number(value);
  }
  
  // 尝试转换为日期
  const dateRegex = /^\d{4}-\d{2}-\d{2}$|^\d{2}\/\d{2}\/\d{4}$|^\d{4}\/\d{2}\/\d{2}$/;
  if (dateRegex.test(value)) {
    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      return date;
    }
  }
  
  // 保持为字符串
  return value;
}

// 获取文件的sheet列表
    router.get('/sheets/:id', auth, async (req, res) => {
      try {
        const data = await Data.findById(req.params.id);
        
        if (!data) {
          return res.status(404).json({ msg: '文件不存在' });
        }
        
        // 检查文件是否属于当前用户
        if (data.userId.toString() !== req.user.id) {
          return res.status(401).json({ msg: '无权访问此文件' });
        }
        
        // 检查文件类型
        if (!data.originalname.endsWith('.xlsx') && !data.originalname.endsWith('.xls')) {
          return res.status(400).json({ msg: '仅支持Excel文件' });
        }
        
        const password = req.query.password;
        
        // 读取Excel文件
        let workbook;
        try {
          workbook = xlsx.readFile(data.path, password ? { password } : undefined);
        } catch (readError) {
           if (readError.message.includes('Password') || readError.message.includes('encrypted')) {
             return res.status(400).json({ msg: '文件受密码保护，请提供正确的密码', needPassword: true });
           }
           throw readError;
        }
        
        const sheets = workbook.SheetNames;
        
        res.json({ sheets });
      } catch (err) {
        console.error(err.message);
        res.status(500).send('服务器错误');
      }
    });

// 对比两个文件
    router.post('/compare', auth, async (req, res) => {
      try {
        const { 
          file1Id, 
          file2Id, 
          sheet1, 
          sheet2, 
          primaryKeys = [], 
          ignoreColumns = [], 
          headerRow = 0, 
          nullValueStrategy = 'ignore',
          detectMoved = false,
          file1Password,
          file2Password
        } = req.body;
        
        // 获取两个文件
        const file1 = await Data.findById(file1Id);
        const file2 = await Data.findById(file2Id);
        
        if (!file1 || !file2) {
          return res.status(404).json({ msg: '文件不存在' });
        }
        
        // 检查文件是否属于当前用户
        if (file1.userId.toString() !== req.user.id || file2.userId.toString() !== req.user.id) {
          return res.status(401).json({ msg: '无权访问此文件' });
        }
        
        // 检查文件类型
        if (!file1.originalname.endsWith('.xlsx') && !file1.originalname.endsWith('.xls')) {
          return res.status(400).json({ msg: '第一个文件必须是Excel文件' });
        }
        
        if (!file2.originalname.endsWith('.xlsx') && !file2.originalname.endsWith('.xls')) {
          return res.status(400).json({ msg: '第二个文件必须是Excel文件' });
        }
        
        // 读取Excel文件
        let workbook1, workbook2;
        try {
          workbook1 = xlsx.readFile(file1.path, file1Password ? { password: file1Password } : undefined);
        } catch (readError) {
           if (readError.message.includes('Password') || readError.message.includes('encrypted')) {
             return res.status(400).json({ msg: `文件1 "${file1.originalname}" 受密码保护，请提供正确的密码`, needFile1Password: true });
           }
           throw readError;
        }

        try {
          workbook2 = xlsx.readFile(file2.path, file2Password ? { password: file2Password } : undefined);
        } catch (readError) {
           if (readError.message.includes('Password') || readError.message.includes('encrypted')) {
             return res.status(400).json({ msg: `文件2 "${file2.originalname}" 受密码保护，请提供正确的密码`, needFile2Password: true });
           }
           throw readError;
        }
        
        // 获取指定sheet的数据
        const sheet1Data = xlsx.utils.sheet_to_json(workbook1.Sheets[sheet1]);
        const sheet2Data = xlsx.utils.sheet_to_json(workbook2.Sheets[sheet2]);
    
    // 获取列名
    const columns1 = Object.keys(sheet1Data[0] || {});
    const columns2 = Object.keys(sheet2Data[0] || {});
    
    // 智能主键推荐
    const recommendedPrimaryKeys = recommendPrimaryKeys([...new Set([...columns1, ...columns2])]);
    
    // 如果没有指定主键，使用推荐的主键
    const finalPrimaryKeys = primaryKeys.length > 0 ? primaryKeys : recommendedPrimaryKeys;
    
    // 过滤掉忽略的列
    const filteredColumns1 = columns1.filter(col => !ignoreColumns.includes(col));
    const filteredColumns2 = columns2.filter(col => !ignoreColumns.includes(col));
    
    // 为每行生成主键值
    const file1Rows = sheet1Data.map((row, index) => {
      // 转换数据类型并处理空值
      const processedRow = {};
      Object.keys(row).forEach(key => {
        if (!ignoreColumns.includes(key)) {
          processedRow[key] = convertDataType(handleNullValues(row[key], nullValueStrategy));
        }
      });
      return {
        ...processedRow,
        __primaryKey: generatePrimaryKeyValue(processedRow, finalPrimaryKeys),
        __index: index
      };
    });
    
    const file2Rows = sheet2Data.map((row, index) => {
      // 转换数据类型并处理空值
      const processedRow = {};
      Object.keys(row).forEach(key => {
        if (!ignoreColumns.includes(key)) {
          processedRow[key] = convertDataType(handleNullValues(row[key], nullValueStrategy));
        }
      });
      return {
        ...processedRow,
        __primaryKey: generatePrimaryKeyValue(processedRow, finalPrimaryKeys),
        __index: index
      };
    });
    
    // 创建主键到行的映射
    const file1Map = new Map();
    file1Rows.forEach(row => {
      if (row.__primaryKey) {
        file1Map.set(row.__primaryKey, row);
      }
    });
    
    const file2Map = new Map();
    file2Rows.forEach(row => {
      if (row.__primaryKey) {
        file2Map.set(row.__primaryKey, row);
      }
    });
    
    // 对比数据
    const differences = [];
    
    // 检查删除的记录（file1中有，file2中没有）
    file1Rows.forEach(row1 => {
      if (row1.__primaryKey && !file2Map.has(row1.__primaryKey)) {
        differences.push({
          type: 'deleted',
          file: 'file1',
          rowIndex: row1.__index,
          rowData: row1,
          primaryKey: row1.__primaryKey
        });
      }
    });
    
    // 检查新增的记录（file2中有，file1中没有）
    file2Rows.forEach(row2 => {
      if (row2.__primaryKey && !file1Map.has(row2.__primaryKey)) {
        differences.push({
          type: 'added',
          file: 'file2',
          rowIndex: row2.__index,
          rowData: row2,
          primaryKey: row2.__primaryKey
        });
      }
    });
    
    // 检查修改的记录（主键匹配，但内容不同）
    file1Rows.forEach(row1 => {
      if (row1.__primaryKey && file2Map.has(row1.__primaryKey)) {
        const row2 = file2Map.get(row1.__primaryKey);
        
        // 检查每行的差异
        const rowDifferences = [];
        filteredColumns1.forEach(column => {
          if (row1[column] !== row2[column]) {
            rowDifferences.push({
              column,
              oldValue: row1[column],
              newValue: row2[column]
            });
          }
        });
        
        if (rowDifferences.length > 0) {
          differences.push({
            type: 'modified',
            file: 'both',
            rowIndex1: row1.__index,
            rowIndex2: row2.__index,
            rowData1: row1,
            rowData2: row2,
            primaryKey: row1.__primaryKey,
            changes: rowDifferences
          });
        } else if (detectMoved && row1.__index !== row2.__index) {
          // 检查移动的记录（内容相同但位置不同）
          differences.push({
            type: 'moved',
            file: 'both',
            oldIndex: row1.__index,
            newIndex: row2.__index,
            rowData: row1,
            primaryKey: row1.__primaryKey
          });
        }
      }
    });
    
    // 统计差异数量
    const addedCount = differences.filter(d => d.type === 'added').length;
    const deletedCount = differences.filter(d => d.type === 'deleted').length;
    const modifiedCount = differences.filter(d => d.type === 'modified').length;
    const movedCount = differences.filter(d => d.type === 'moved').length;
    
    // 生成任务ID
    const taskId = uuidv4();
    
    // 生成Excel报表
    const reportFileName = `report-${taskId}.xlsx`;
    const reportPath = path.join('uploads', 'reports', reportFileName);
    
    // 创建新的工作簿
    const reportWorkbook = xlsx.utils.book_new();
    
    // 1. 差异汇总概览 Sheet
    const summaryData = [
      ['对比报告概览'],
      ['任务ID', taskId],
      ['对比时间', new Date().toLocaleString()],
      ['文件1', file1.originalname],
      ['文件2', file2.originalname],
      ['Sheet1', sheet1],
      ['Sheet2', sheet2],
      ['主键', finalPrimaryKeys.join(', ')],
      [''],
      ['统计信息'],
      ['文件1行数', file1Rows.length],
      ['文件2行数', file2Rows.length],
      ['新增记录数', addedCount],
      ['删除记录数', deletedCount],
      ['修改记录数', modifiedCount],
      ['移动记录数', movedCount],
      ['差异总计', differences.length]
    ];
    const summarySheet = xlsx.utils.aoa_to_sheet(summaryData);
    xlsx.utils.book_append_sheet(reportWorkbook, summarySheet, '差异汇总概览');
    
    // 2. 差异详情列表 Sheet
    const detailsData = differences.map(diff => {
      let changeDetail = '';
      if (diff.type === 'modified') {
        changeDetail = diff.changes.map(c => `${c.column}: ${c.oldValue} -> ${c.newValue}`).join('; ');
      } else if (diff.type === 'moved') {
        changeDetail = `行号: ${diff.oldIndex} -> ${diff.newIndex}`;
      }
      
      return {
        '差异类型': diff.type === 'added' ? '新增' : diff.type === 'deleted' ? '删除' : diff.type === 'modified' ? '修改' : '移动',
        '主键值': diff.primaryKey,
        '行号(文件1)': diff.rowIndex1 !== undefined ? diff.rowIndex1 + 1 : (diff.rowIndex !== undefined && diff.file === 'file1' ? diff.rowIndex + 1 : ''),
        '行号(文件2)': diff.rowIndex2 !== undefined ? diff.rowIndex2 + 1 : (diff.rowIndex !== undefined && diff.file === 'file2' ? diff.rowIndex + 1 : ''),
        '变更详情': changeDetail
      };
    });
    const detailsSheet = xlsx.utils.json_to_sheet(detailsData);
    xlsx.utils.book_append_sheet(reportWorkbook, detailsSheet, '差异详情列表');
    
    // 3. 新增项明细 Sheet
    const addedItems = differences.filter(d => d.type === 'added').map(d => d.rowData);
    if (addedItems.length > 0) {
      const addedSheet = xlsx.utils.json_to_sheet(addedItems);
      xlsx.utils.book_append_sheet(reportWorkbook, addedSheet, '新增项明细');
    } else {
      xlsx.utils.book_append_sheet(reportWorkbook, xlsx.utils.aoa_to_sheet([['无新增项']]), '新增项明细');
    }
    
    // 4. 删除项明细 Sheet
    const deletedItems = differences.filter(d => d.type === 'deleted').map(d => d.rowData);
    if (deletedItems.length > 0) {
      const deletedSheet = xlsx.utils.json_to_sheet(deletedItems);
      xlsx.utils.book_append_sheet(reportWorkbook, deletedSheet, '删除项明细');
    } else {
      xlsx.utils.book_append_sheet(reportWorkbook, xlsx.utils.aoa_to_sheet([['无删除项']]), '删除项明细');
    }
    
    // 5. 修改项对比 Sheet
    const modifiedItems = differences.filter(d => d.type === 'modified').map(d => {
      const item = { '主键值': d.primaryKey };
      d.changes.forEach(c => {
        item[`${c.column} (旧值)`] = c.oldValue;
        item[`${c.column} (新值)`] = c.newValue;
      });
      return item;
    });
    if (modifiedItems.length > 0) {
      const modifiedSheet = xlsx.utils.json_to_sheet(modifiedItems);
      xlsx.utils.book_append_sheet(reportWorkbook, modifiedSheet, '修改项对比');
    } else {
      xlsx.utils.book_append_sheet(reportWorkbook, xlsx.utils.aoa_to_sheet([['无修改项']]), '修改项对比');
    }
    
    // 写入文件
    xlsx.writeFile(reportWorkbook, reportPath);
    
    // 保存对比记录到数据库
    const comparison = new Comparison({
      taskId,
      userId: req.user.id,
      file1Id,
      file2Id,
      sheet1,
      sheet2,
      primaryKey: finalPrimaryKeys,
      ignoreColumns,
      headerRow,
      nullStrategy: nullValueStrategy, // 注意字段名匹配
      stats: {
        file1Rows: file1Rows.length,
        file2Rows: file2Rows.length,
        added: addedCount,
        deleted: deletedCount,
        modified: modifiedCount,
        totalDifferences: differences.length
      },
      reportPath,
      status: 'completed'
    });
    
    await comparison.save();
    
    res.json({
      taskId,
      file1: file1.originalname,
      file2: file2.originalname,
      sheet1,
      sheet2,
      primaryKeys: finalPrimaryKeys,
      recommendedPrimaryKeys,
      ignoreColumns,
      headerRow,
      nullValueStrategy,
      detectMoved,
      differences,
      reportPath
    });
  } catch (err) {
    console.error(err.message);
    res.status(500).send('服务器错误');
  }
});

// 获取文件的列信息和推荐主键
    router.get('/columns/:id/:sheet', auth, async (req, res) => {
      try {
        const { id, sheet } = req.params;
        const password = req.query.password;
        
        const data = await Data.findById(id);
        
        if (!data) {
          return res.status(404).json({ msg: '文件不存在' });
        }
        
        // 检查文件是否属于当前用户
        if (data.userId.toString() !== req.user.id) {
          return res.status(401).json({ msg: '无权访问此文件' });
        }
        
        // 检查文件类型
        if (!data.originalname.endsWith('.xlsx') && !data.originalname.endsWith('.xls')) {
          return res.status(400).json({ msg: '仅支持Excel文件' });
        }
        
        // 读取Excel文件
        let workbook;
        try {
          workbook = xlsx.readFile(data.path, password ? { password } : undefined);
        } catch (readError) {
           if (readError.message.includes('Password') || readError.message.includes('encrypted')) {
             return res.status(400).json({ msg: '文件受密码保护，请提供正确的密码', needPassword: true });
           }
           throw readError;
        }
        
        const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheet]);
        
        // 获取列名
        const columns = Object.keys(sheetData[0] || {});
        
        // 智能主键推荐
        const recommendedPrimaryKeys = recommendPrimaryKeys(columns);
        
        res.json({
          columns,
          recommendedPrimaryKeys
        });
      } catch (err) {
        console.error(err.message);
        res.status(500).send('服务器错误');
      }
    });

// 下载报表
router.get('/report/:taskId', auth, async (req, res) => {
  try {
    const comparison = await Comparison.findOne({ taskId: req.params.taskId });
    
    if (!comparison) {
      return res.status(404).json({ msg: '记录不存在' });
    }
    
    // 检查权限
    if (comparison.userId.toString() !== req.user.id && req.user.role !== 'admin') {
      return res.status(401).json({ msg: '无权访问此文件' });
    }
    
    const filePath = comparison.reportPath;
    
    if (fs.existsSync(filePath)) {
      res.download(filePath);
    } else {
      res.status(404).json({ msg: '报表文件不存在' });
    }
  } catch (err) {
    console.error(err.message);
    res.status(500).send('服务器错误');
  }
});

module.exports = router;