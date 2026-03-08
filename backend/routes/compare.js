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

// 辅助函数：确保Sheet范围从第一列(A列)开始
const ensureRangeStartsFromA = (sheet) => {
  if (!sheet || !sheet['!ref']) return;
  const range = xlsx.utils.decode_range(sheet['!ref']);
  if (range.s.c > 0) {
    range.s.c = 0;
    sheet['!ref'] = xlsx.utils.encode_range(range);
  }
};

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
  if (value === null || value === undefined || String(value).trim() === '') {
    return value;
  }
  
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

// 辅助函数：获取处理合并单元格后的多级表头 (复用 data.js 中的逻辑)
const getHeaders = (sheet, startRow, endRow) => {
  if (!sheet || !sheet['!ref']) return { headers: [], structure: [] };
  const range = xlsx.utils.decode_range(sheet['!ref']);
  const merges = sheet['!merges'] || [];
  
  if (endRow === undefined || endRow === null) {
    endRow = startRow;
  }
  
  if (startRow > endRow) {
    const temp = startRow;
    startRow = endRow;
    endRow = temp;
  }
  
  const headerMatrix = [];
  
  for (let R = startRow; R <= endRow; ++R) {
    const rowValues = [];
    for (let C = range.s.c; C <= range.e.c; ++C) {
      let cellAddress = { c: C, r: R };
      let cellRef = xlsx.utils.encode_cell(cellAddress);
      let cell = sheet[cellRef];
      let val = (cell && cell.v !== undefined) ? cell.v : null;
      
      if (val === null || val === '') {
        const merge = merges.find(m => 
          R >= m.s.r && R <= m.e.r &&
          C >= m.s.c && C <= m.e.c
        );
        
        if (merge) {
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
  }
  
  const headers = [];
  const seenHeaders = {};
  
  const colCount = range.e.c - range.s.c + 1;
  for (let C = 0; C < colCount; ++C) {
    const parts = [];
    for (let R = 0; R < headerMatrix.length; ++R) {
      const val = headerMatrix[R][C];
      if (val !== null && val !== '' && val !== undefined) {
        const strVal = String(val).trim();
        if (parts.length === 0 || parts[parts.length - 1] !== strVal) {
             parts.push(strVal);
        }
      }
    }
    
    let headerStr = parts.join('::');
    
    if (headerStr === '') {
      headerStr = `__EMPTY_${range.s.c + C}`;
    }
    
    if (seenHeaders[headerStr]) {
      seenHeaders[headerStr]++;
      headerStr = `${headerStr}_${seenHeaders[headerStr]}`;
    } else {
      seenHeaders[headerStr] = 1;
    }
    
    headers.push(headerStr);
  }
  
  return headers;
};

// 辅助函数：处理合并空表头 (并恢复多级表头结构)
const mergeEmptyHeaders = (ws) => {
  if (!ws || !ws['!ref']) return;
  const range = xlsx.utils.decode_range(ws['!ref']);
  const merges = ws['!merges'] || [];
  
  // 1. 先进行列合并（针对 __EMPTY, 0, 0_1 等）
  let currentStartCol = -1;
  let isMerging = false;

  // 遍历第一行（表头行）
  // 获取所有列名
  const headers = [];
  for (let C = range.s.c; C <= range.e.c; ++C) {
    const cellRef = xlsx.utils.encode_cell({c: C, r: range.s.r});
    const cell = ws[cellRef];
    headers.push(cell ? cell.v : '');
  }

  // 检查是否有多级表头（包含 :: 分隔符）
  const hasMultiLevel = headers.some(h => String(h).includes('::'));
  
  if (hasMultiLevel) {
     // 计算最大层级深度
     let maxDepth = 0;
     const headerStructure = headers.map(h => {
       const parts = String(h).split('::');
       if (parts.length > maxDepth) maxDepth = parts.length;
       return parts;
     });
     
     // 如果有多级，需要向下移动数据，腾出表头空间
     if (maxDepth > 1) {
        const dataRange = xlsx.utils.decode_range(ws['!ref']);
        // 移动数据：从最后一行开始向上移动
        for (let R = dataRange.e.r; R > dataRange.s.r; --R) {
           for (let C = dataRange.s.c; C <= dataRange.e.c; ++C) {
              const fromRef = xlsx.utils.encode_cell({c: C, r: R});
              const toRef = xlsx.utils.encode_cell({c: C, r: R + maxDepth - 1});
              ws[toRef] = ws[fromRef];
           }
        }
        
        // 更新 range
        ws['!ref'] = xlsx.utils.encode_range({
           s: { c: dataRange.s.c, r: dataRange.s.r },
           e: { c: dataRange.e.c, r: dataRange.e.r + maxDepth - 1 }
        });
        
        // 填充多级表头
        for (let C = 0; C < headers.length; ++C) {
           const parts = headerStructure[C];
           for (let level = 0; level < maxDepth; ++level) {
              const cellRef = xlsx.utils.encode_cell({c: range.s.c + C, r: range.s.r + level});
              let val = parts[level] || parts[parts.length - 1] || '';
              
              if (String(val).startsWith('__EMPTY') || String(val) === '0' || /^0_\d+$/.test(String(val))) {
                  val = ''; 
              }

              ws[cellRef] = { v: val, t: 's' };
           }
        }
        
        // 应用合并逻辑（横向和纵向）
        // 1. 横向合并
        for (let level = 0; level < maxDepth; ++level) {
           let mergeStart = -1;
           let mergeVal = null;
           
           for (let C = range.s.c; C <= range.e.c; ++C) {
              const cellRef = xlsx.utils.encode_cell({c: C, r: range.s.r + level});
              const cell = ws[cellRef];
              const val = cell ? cell.v : '';
              
              if (val === mergeVal && val !== '') {
                 // 继续合并
              } else {
                 if (mergeStart !== -1 && (C - 1) > mergeStart) {
                    merges.push({
                       s: { r: range.s.r + level, c: mergeStart },
                       e: { r: range.s.r + level, c: C - 1 }
                    });
                    for (let mC = mergeStart + 1; mC < C; ++mC) {
                       const mRef = xlsx.utils.encode_cell({c: mC, r: range.s.r + level});
                       if (ws[mRef]) ws[mRef].v = '';
                    }
                 }
                 mergeStart = C;
                 mergeVal = val;
              }
           }
           if (mergeStart !== -1 && range.e.c > mergeStart) {
              merges.push({
                 s: { r: range.s.r + level, c: mergeStart },
                 e: { r: range.s.r + level, c: range.e.c }
              });
               for (let mC = mergeStart + 1; mC <= range.e.c; ++mC) {
                  const mRef = xlsx.utils.encode_cell({c: mC, r: range.s.r + level});
                  if (ws[mRef]) ws[mRef].v = '';
               }
           }
        }
        
        // 2. 纵向合并
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const parts = headerStructure[C - range.s.c];
            const firstVal = parts[0];
            if (firstVal && !String(firstVal).startsWith('__EMPTY') && firstVal !== '0') {
                let mergeStartRow = -1;
                let mergeVal = null;
                
                for (let level = 0; level < maxDepth; ++level) {
                    const cellRef = xlsx.utils.encode_cell({c: C, r: range.s.r + level});
                    const val = ws[cellRef] ? ws[cellRef].v : '';
                    
                    if (val === mergeVal && val !== '') {
                       // 继续
                    } else {
                       if (mergeStartRow !== -1 && (level - 1) > mergeStartRow) {
                          merges.push({
                             s: { r: range.s.r + mergeStartRow, c: C },
                             e: { r: range.s.r + level - 1, c: C }
                          });
                          for (let mR = mergeStartRow + 1; mR < level; ++mR) {
                              const mRef = xlsx.utils.encode_cell({c: C, r: range.s.r + mR});
                              if (ws[mRef]) ws[mRef].v = '';
                          }
                       }
                       mergeStartRow = level;
                       mergeVal = val;
                    }
                }
                if (mergeStartRow !== -1 && (maxDepth - 1) > mergeStartRow) {
                    merges.push({
                       s: { r: range.s.r + mergeStartRow, c: C },
                       e: { r: range.s.r + maxDepth - 1, c: C }
                    });
                    for (let mR = mergeStartRow + 1; mR < maxDepth; ++mR) {
                        const mRef = xlsx.utils.encode_cell({c: C, r: range.s.r + mR});
                        if (ws[mRef]) ws[mRef].v = '';
                    }
                }
            }
        }
     } else {
       // 单级表头，使用之前的逻辑
       for (let C = range.s.c; C <= range.e.c; ++C) {
          const cellRef = xlsx.utils.encode_cell({c: C, r: range.s.r});
          const cell = ws[cellRef];
          
          let shouldMerge = false;
          
          if (cell && cell.v !== undefined) {
             const val = String(cell.v).trim();
             if (val.startsWith('__EMPTY') || val === '0' || val === '' || /^0_\d+$/.test(val)) {
               shouldMerge = true;
             }
          } else {
             shouldMerge = true;
          }
          
          if (shouldMerge) {
             if (!isMerging) {
                if (C > range.s.c) {
                   currentStartCol = C - 1;
                   isMerging = true;
                }
             }
          } else {
             if (isMerging) {
                if (currentStartCol >= range.s.c && (C - 1) > currentStartCol) {
                    merges.push({
                      s: { r: range.s.r, c: currentStartCol },
                      e: { r: range.s.r, c: C - 1 }
                    });
                    for (let mC = currentStartCol + 1; mC < C; ++mC) {
                        const mCellRef = xlsx.utils.encode_cell({c: mC, r: range.s.r});
                        if (ws[mCellRef]) {
                          ws[mCellRef].v = ''; 
                          ws[mCellRef].t = 's';
                        }
                    }
                }
                isMerging = false;
             }
          }
        }
        
        if (isMerging) {
           if (currentStartCol >= range.s.c && range.e.c > currentStartCol) {
              merges.push({
                 s: { r: range.s.r, c: currentStartCol },
                 e: { r: range.s.r, c: range.e.c }
              });
               for (let mC = currentStartCol + 1; mC <= range.e.c; ++mC) {
                    const mCellRef = xlsx.utils.encode_cell({c: mC, r: range.s.r});
                    if (ws[mCellRef]) {
                      ws[mCellRef].v = '';
                      ws[mCellRef].t = 's';
                    }
               }
           }
        }
     }
  }
  
  if (merges.length > 0) {
    ws['!merges'] = merges;
  }
};

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
          headerRowStart = 0,
          headerRowEnd = 0,
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
        const ws1 = workbook1.Sheets[sheet1];
        const ws2 = workbook2.Sheets[sheet2];
        
        ensureRangeStartsFromA(ws1);
        ensureRangeStartsFromA(ws2);
        
        const sheet1Data = xlsx.utils.sheet_to_json(ws1, { header: getHeaders(ws1, headerRowStart, headerRowEnd), range: headerRowEnd + 1, defval: '' });
        const sheet2Data = xlsx.utils.sheet_to_json(ws2, { header: getHeaders(ws2, headerRowStart, headerRowEnd), range: headerRowEnd + 1, defval: '' });
    
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
    const sameItems = [];
    
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
        // 跳过自动生成的空列名
        if (column.startsWith('__EMPTY_')) return;
        
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
        } else {
          sameItems.push(row1);
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
    const addedItems = differences.filter(d => d.type === 'added').map(d => {
      // 移除内部字段
      const { __primaryKey, __index, ...rest } = d.rowData;
      return rest;
    });
    if (addedItems.length > 0) {
      const addedSheet = xlsx.utils.json_to_sheet(addedItems);
      mergeEmptyHeaders(addedSheet);
      xlsx.utils.book_append_sheet(reportWorkbook, addedSheet, '新增项明细');
    } else {
      xlsx.utils.book_append_sheet(reportWorkbook, xlsx.utils.aoa_to_sheet([['无新增项']]), '新增项明细');
    }
    
    // 4. 删除项明细 Sheet
    const deletedItems = differences.filter(d => d.type === 'deleted').map(d => {
      // 移除内部字段
      const { __primaryKey, __index, ...rest } = d.rowData;
      return rest;
    });
    if (deletedItems.length > 0) {
      const deletedSheet = xlsx.utils.json_to_sheet(deletedItems);
      mergeEmptyHeaders(deletedSheet);
      xlsx.utils.book_append_sheet(reportWorkbook, deletedSheet, '删除项明细');
    } else {
      xlsx.utils.book_append_sheet(reportWorkbook, xlsx.utils.aoa_to_sheet([['无删除项']]), '删除项明细');
    }
    
    // 5. 修改项对比 Sheet
    const modifiedItems = differences.filter(d => d.type === 'modified').map(d => {
      const item = { '主键值': d.primaryKey };
      d.changes.forEach(c => {
        // 处理表头合并逻辑：如果列名是自动生成的（__EMPTY_开头），则尝试与前一列合并显示
        // 但在这里我们拿到的是已经解析好的列名，如果是自动生成的，说明原文件就是空的
        // 用户需求是：某一列的 表头 如果为空，则与前面一列合并
        // 这里我们在生成报告时，如果列名以 __EMPTY_ 开头，我们在展示时可以做特殊处理
        // 或者，在更早的数据读取阶段处理表头
        
        // 实际上，xlsx.utils.json_to_sheet 会使用对象的 key 作为表头
        // 我们可以在这里构造更有意义的 key
        
        let displayColumn = c.column;
        if (displayColumn.startsWith('__EMPTY_')) {
           // 尝试找到它的前一列（这需要知道列的顺序，目前 changes 数组不一定有序，且不包含所有列）
           // 简单起见，我们直接保留原样，或者在数据读取时就处理好表头
           // 更好的方式是在生成 sheet 之前，对 item 的 key 进行重命名
        }
        
        item[`${displayColumn} (旧值)`] = c.oldValue;
        item[`${displayColumn} (新值)`] = c.newValue;
      });
      return item;
    });
    


    if (modifiedItems.length > 0) {
      const modifiedSheet = xlsx.utils.json_to_sheet(modifiedItems);
      // 处理表头
      mergeEmptyHeaders(modifiedSheet);
      xlsx.utils.book_append_sheet(reportWorkbook, modifiedSheet, '修改项对比');
    } else {
      xlsx.utils.book_append_sheet(reportWorkbook, xlsx.utils.aoa_to_sheet([['无修改项']]), '修改项对比');
    }

    // 6. 相同项明细 Sheet
    if (sameItems.length > 0) {
      // 移除内部字段
      const cleanSameItems = sameItems.map(item => {
        const { __primaryKey, __index, ...rest } = item;
        return rest;
      });
      const sameSheet = xlsx.utils.json_to_sheet(cleanSameItems);
      mergeEmptyHeaders(sameSheet);
      xlsx.utils.book_append_sheet(reportWorkbook, sameSheet, '相同项明细');
    } else {
      xlsx.utils.book_append_sheet(reportWorkbook, xlsx.utils.aoa_to_sheet([['无相同项']]), '相同项明细');
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
      headerRowStart,
      headerRowEnd,
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
      headerRowStart,
      headerRowEnd,
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
        const headerRowStart = parseInt(req.query.headerRowStart) || 0;
        const headerRowEnd = parseInt(req.query.headerRowEnd) !== undefined ? parseInt(req.query.headerRowEnd) : headerRowStart;
        
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
        
        const ws = workbook.Sheets[sheet];
        ensureRangeStartsFromA(ws);
        
        const sheetData = xlsx.utils.sheet_to_json(ws, { header: getHeaders(ws, headerRowStart, headerRowEnd), range: headerRowEnd + 1, defval: '' });
        
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