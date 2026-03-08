
const express = require('express');
const router = express.Router();
const Data = require('../models/Data');
const Analysis = require('../models/Analysis');
const auth = require('../middleware/auth');
const xlsx = require('xlsx');

// 辅助函数：确保Sheet范围从第一列(A列)开始
const ensureRangeStartsFromA = (sheet) => {
  if (!sheet || !sheet['!ref']) return;
  const range = xlsx.utils.decode_range(sheet['!ref']);
  if (range.s.c > 0) {
    range.s.c = 0;
    sheet['!ref'] = xlsx.utils.encode_range(range);
  }
};

// 辅助函数：获取处理合并单元格后的多级表头 (复用 compare.js 中的逻辑)
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

// 获取文件指定Sheet的列信息（用于配置图表）
router.get('/columns/:id', auth, async (req, res) => {
  try {
    const { id } = req.params;
    const { sheet, password } = req.query;
    
    const data = await Data.findById(id);
    if (!data) return res.status(404).json({ msg: '文件不存在' });
    
    // 检查权限
    if (data.userId.toString() !== req.user.id && req.user.role !== 'admin') {
      return res.status(401).json({ msg: '无权访问此文件' });
    }
    
    let workbook;
    try {
      workbook = xlsx.readFile(data.path, password ? { password } : undefined);
    } catch (error) {
       if (error.message.includes('Password') || error.message.includes('encrypted')) {
         return res.status(400).json({ msg: '文件受密码保护，请提供正确的密码', needPassword: true });
       }
       throw error;
    }

    const sheetName = sheet || workbook.SheetNames[0];
    const ws = workbook.Sheets[sheetName];
    if (!ws) return res.status(404).json({ msg: 'Sheet不存在' });

    ensureRangeStartsFromA(ws);
    
    const headerRowStart = data.metadata?.headerRowStart || 0;
    const headerRowEnd = data.metadata?.headerRowEnd !== undefined ? data.metadata.headerRowEnd : (data.metadata?.headerRowStart || 0);
    
    const columns = getHeaders(ws, headerRowStart, headerRowEnd);
    
    res.json({ columns, sheet: sheetName, sheets: workbook.SheetNames });
  } catch (err) {
    console.error(err.message);
    res.status(500).send('服务器错误');
  }
});

// 执行数据分析（聚合/统计）
router.post('/analyze', auth, async (req, res) => {
  try {
    const { fileId, sheet, xAxis, yAxis, aggregation, password, chartType } = req.body;
    
    const data = await Data.findById(fileId);
    if (!data) return res.status(404).json({ msg: '文件不存在' });
    
    // 检查权限
    if (data.userId.toString() !== req.user.id && req.user.role !== 'admin') {
      return res.status(401).json({ msg: '无权访问此文件' });
    }
    
    let workbook;
    try {
      workbook = xlsx.readFile(data.path, password ? { password } : undefined);
    } catch (error) {
       if (error.message.includes('Password') || error.message.includes('encrypted')) {
         return res.status(400).json({ msg: '文件受密码保护', needPassword: true });
       }
       throw error;
    }

    const ws = workbook.Sheets[sheet];
    ensureRangeStartsFromA(ws);
    
    const headerRowStart = data.metadata?.headerRowStart || 0;
    const headerRowEnd = data.metadata?.headerRowEnd !== undefined ? data.metadata.headerRowEnd : (data.metadata?.headerRowStart || 0);
    
    const headers = getHeaders(ws, headerRowStart, headerRowEnd);
    
    // 读取所有数据
    const jsonData = xlsx.utils.sheet_to_json(ws, { header: headers, range: headerRowEnd + 1, defval: null });
    
    // 数据处理逻辑
    // 1. 过滤掉 X轴 或 Y轴 为空的行
    // 2. 根据 X轴 分组
    // 3. 对 Y轴 进行聚合 (Count, Sum, Avg, Max, Min)
    
    const groupedData = {};
    
    jsonData.forEach(row => {
      const xValue = row[xAxis];
      // 如果 X轴值为空，且不是 0，则跳过
      if (xValue === null || xValue === undefined || xValue === '') return;
      
      const key = String(xValue);
      
      if (!groupedData[key]) {
        groupedData[key] = [];
      }
      
      // 解析 Y轴数值
      let yValue = row[yAxis];
      if (aggregation !== 'count') {
         yValue = parseFloat(yValue);
         if (isNaN(yValue)) yValue = 0;
      }
      
      groupedData[key].push(yValue);
    });
    
    const resultX = [];
    const resultY = [];
    
    // 对分组后的数据进行聚合
    Object.keys(groupedData).forEach(key => {
      const values = groupedData[key];
      let aggValue = 0;
      
      switch (aggregation) {
        case 'count':
          aggValue = values.length;
          break;
        case 'sum':
          aggValue = values.reduce((a, b) => a + b, 0);
          break;
        case 'avg':
          aggValue = values.length > 0 ? values.reduce((a, b) => a + b, 0) / values.length : 0;
          break;
        case 'max':
          aggValue = Math.max(...values);
          break;
        case 'min':
          aggValue = Math.min(...values);
          break;
        default:
          aggValue = values.length;
      }
      
      // 保留两位小数
      if (typeof aggValue === 'number' && !Number.isInteger(aggValue)) {
          aggValue = parseFloat(aggValue.toFixed(2));
      }

      resultX.push(key);
      resultY.push(aggValue);
    });
    
    // 简单的排序 (可选: 根据 X轴 排序)
    // 这里暂时不做复杂排序，保持 Object.keys 的顺序 (通常是插入顺序或字典序)
    
    const analysisResult = {
      xAxis: resultX,
      series: resultY,
      chartType,
      title: `${aggregation.toUpperCase()} of ${yAxis.replace(/::/g, ' > ')} by ${xAxis.replace(/::/g, ' > ')}`
    };

    // 保存分析结果到数据库
    const newAnalysis = new Analysis({
      userId: req.user.id,
      fileId,
      sheet,
      chartType,
      xAxis,
      yAxis,
      aggregation,
      result: analysisResult
    });
    
    await newAnalysis.save();
    
    res.json(analysisResult);

  } catch (err) {
    console.error(err.message);
    res.status(500).send('服务器错误');
  }
});

// 获取历史分析记录
router.get('/history', auth, async (req, res) => {
  try {
    const query = {};
    if (req.user.role !== 'admin') {
      query.userId = req.user.id;
    }
    const history = await Analysis.find(query)
      .populate('fileId', 'originalname')
      .sort({ createdAt: -1 });
    res.json(history);
  } catch (err) {
    console.error(err.message);
    res.status(500).send('服务器错误');
  }
});

// 生成综合报表
router.post('/report', auth, async (req, res) => {
  try {
    const { fileId, sheet, password } = req.body;
    
    const data = await Data.findById(fileId);
    if (!data) return res.status(404).json({ msg: '文件不存在' });
    
    // 检查权限
    if (data.userId.toString() !== req.user.id && req.user.role !== 'admin') {
      return res.status(401).json({ msg: '无权访问此文件' });
    }
    
    let workbook;
    try {
      workbook = xlsx.readFile(data.path, password ? { password } : undefined);
    } catch (error) {
       if (error.message.includes('Password') || error.message.includes('encrypted')) {
         return res.status(400).json({ msg: '文件受密码保护', needPassword: true });
       }
       throw error;
    }

    const ws = workbook.Sheets[sheet];
    ensureRangeStartsFromA(ws);
    
    const headerRowStart = data.metadata?.headerRowStart || 0;
    const headerRowEnd = data.metadata?.headerRowEnd !== undefined ? data.metadata.headerRowEnd : (data.metadata?.headerRowStart || 0);
    
    const headers = getHeaders(ws, headerRowStart, headerRowEnd);
    
    // 读取所有数据
    const jsonData = xlsx.utils.sheet_to_json(ws, { header: headers, range: headerRowEnd + 1, defval: null });
    
    // 创建新的 Workbook
    const reportWorkbook = xlsx.utils.book_new();
    
    // 汇总 Sheet 数据
    const summaryRows = [];
    summaryRows.push(['列名', '类型', '有效值数量', '空值数量', '唯一值数量', '数值统计 (Sum/Avg/Max/Min)', 'Top 1 值', 'Top 1 数量', 'Top 2 值', 'Top 2 数量', 'Top 3 值', 'Top 3 数量']);
    
    // 遍历每一个表头列，生成统计信息
    headers.forEach(column => {
       // 跳过空列名
       if (!column || column.startsWith('__EMPTY')) return;
       
       const cleanColumnName = column.replace(/::/g, '_').substring(0, 31); // Sheet名最长31字符
       
       // 统计逻辑
       // 1. 值分布统计 (Value Distribution)
       const valueCounts = {};
       let numericValues = [];
       let totalCount = 0;
       let nullCount = 0;
       
       jsonData.forEach(row => {
          const val = row[column];
          if (val === null || val === undefined || val === '') {
             nullCount++;
             return;
          }
          
          totalCount++;
          const key = String(val);
          valueCounts[key] = (valueCounts[key] || 0) + 1;
          
          // 尝试转数字
          if (!isNaN(val) && val !== '') {
             numericValues.push(Number(val));
          }
       });
       
       // 生成统计数据行
       const statsRows = [];
       statsRows.push(['统计项', '值']);
       statsRows.push(['列名', column]);
       statsRows.push(['总行数', jsonData.length]);
       statsRows.push(['有效值数量', totalCount]);
       statsRows.push(['空值数量', nullCount]);
       statsRows.push(['唯一值数量', Object.keys(valueCounts).length]);
       
       let colType = '文本型';
       let numericStats = '-';
       
       if (numericValues.length > 0) {
          colType = '数值型';
          const sum = numericValues.reduce((a, b) => a + b, 0);
          const avg = sum / numericValues.length;
          const max = Math.max(...numericValues);
          const min = Math.min(...numericValues);
          
          numericStats = `Sum:${sum}, Avg:${avg.toFixed(2)}, Max:${max}, Min:${min}`;
          
          statsRows.push(['类型', '数值型']);
          statsRows.push(['求和 (Sum)', sum]);
          statsRows.push(['平均值 (Avg)', avg.toFixed(2)]);
          statsRows.push(['最大值 (Max)', max]);
          statsRows.push(['最小值 (Min)', min]);
       } else {
          statsRows.push(['类型', '文本型']);
       }
       
       statsRows.push([]); // 空行
       statsRows.push(['值分布 (Top 50)', '数量']);
       
       // 按数量排序
       const sortedValues = Object.entries(valueCounts)
          .sort((a, b) => b[1] - a[1]);
          
       const topValues = sortedValues.slice(0, 3);
       const topValuesStr = [];
       for (let i = 0; i < 3; i++) {
          if (topValues[i]) {
            topValuesStr.push(topValues[i][0]);
            topValuesStr.push(topValues[i][1]);
          } else {
            topValuesStr.push('');
            topValuesStr.push('');
          }
       }
       
       // 添加到汇总行
       summaryRows.push([
          column, 
          colType, 
          totalCount, 
          nullCount, 
          Object.keys(valueCounts).length, 
          numericStats,
          ...topValuesStr
       ]);

       // 详情页只取前50个
       const detailSortedValues = sortedValues.slice(0, 50); 
       detailSortedValues.forEach(([val, count]) => {
          statsRows.push([val, count]);
       });
       
       // 修正逻辑：先生成数据，再生成Sheet，最后添加
       const wsStats = xlsx.utils.aoa_to_sheet(statsRows);
       
       // 设置列宽
       wsStats['!cols'] = [{ wch: 20 }, { wch: 15 }];
       
       // 尝试生成 Sheet 名称
       let sheetName = cleanColumnName.replace(/[\\/?*[\]]/g, '').substring(0, 31);
       
       // 确保 Sheet 名唯一且有效
       let counter = 1;
       while (reportWorkbook.SheetNames.includes(sheetName)) {
          sheetName = cleanColumnName.substring(0, 28) + `_${counter}`;
          counter++;
       }
       
       try {
         xlsx.utils.book_append_sheet(reportWorkbook, wsStats, sheetName);
       } catch (e) {
          console.error('Sheet添加失败', sheetName, e);
       }
    });
    
    // 添加汇总 Sheet 到最前面
    if (summaryRows.length > 0) {
       const wsSummary = xlsx.utils.aoa_to_sheet(summaryRows);
       wsSummary['!cols'] = [
          { wch: 20 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 40 },
          { wch: 15 }, { wch: 10 }, { wch: 15 }, { wch: 10 }, { wch: 15 }, { wch: 10 }
       ];
       
       // xlsx 库没有直接的 insert sheet 方法，只能 append
       // 如果非要放第一页，需要重新构建 workbook，或者在 append 之前就处理
       // 但由于我们需要遍历 columns 才能生成 summary，所以先收集数据，最后再处理 workbook 的 Sheets 顺序
       // 实际上，我们可以先 append Summary，然后再 append 其他的，但是我们是在循环里生成的其他 Sheet
       
       // 方案：创建一个临时的 workbook 用来存放 detail sheets，最后合并
       // 或者更简单：先生成 summary sheet，然后把 reportWorkbook 里的 sheet 挪到后面？
       // xlsx 的 book_append_sheet 是按顺序加的。
       // 我们可以在循环里不 append，而是把 sheet 对象存数组里，最后一起 append
    }
    
    // 重新组织 Workbook (把 Summary 放到第一个)
    const finalWorkbook = xlsx.utils.book_new();
    const wsSummary = xlsx.utils.aoa_to_sheet(summaryRows);
    wsSummary['!cols'] = [
          { wch: 20 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 40 },
          { wch: 15 }, { wch: 10 }, { wch: 15 }, { wch: 10 }, { wch: 15 }, { wch: 10 }
    ];
    xlsx.utils.book_append_sheet(finalWorkbook, wsSummary, '全表汇总');
    
    // 把之前 reportWorkbook 里的 sheet 复制过来
    reportWorkbook.SheetNames.forEach(name => {
       xlsx.utils.book_append_sheet(finalWorkbook, reportWorkbook.Sheets[name], name);
    });
    
    // 写入Buffer
    const buffer = xlsx.write(finalWorkbook, { type: 'buffer', bookType: 'xlsx' });
    
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=Analysis_Report_${Date.now()}.xlsx`);
    res.send(buffer);

  } catch (err) {
    console.error(err.message);
    res.status(500).send('服务器错误');
  }
});

module.exports = router;
