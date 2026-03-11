const express = require('express');
const router = express.Router();
const Template = require('../models/Template');
const auth = require('../middleware/auth');
const multer = require('multer');
const path = require('path');
const fs = require('fs');
const xlsx = require('xlsx');

// Helper function: Get multi-level headers after handling merged cells
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
  
  const headerStructure = [];
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
    headerStructure.push(rowValues);
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
  
  return { headers, structure: headerStructure };
};

// Helper function: Ensure Sheet range starts from column A
const ensureRangeStartsFromA = (sheet) => {
  if (!sheet || !sheet['!ref']) return;
  const range = xlsx.utils.decode_range(sheet['!ref']);
  if (range.s.c > 0) {
    range.s.c = 0;
    sheet['!ref'] = xlsx.utils.encode_range(range);
  }
};

const uploadDir = path.join(process.cwd(), 'uploads', 'templates');

// Ensure templates directory exists
if (!fs.existsSync(uploadDir)) {
  fs.mkdirSync(uploadDir, { recursive: true });
}

// Storage configuration
const storage = multer.diskStorage({
  destination: function(req, file, cb) {
    cb(null, uploadDir);
  },
  filename: function(req, file, cb) {
    // Handle filename encoding
    const originalname = Buffer.from(file.originalname, 'latin1').toString('utf8');
    cb(null, Date.now() + '-' + originalname);
  }
});

const upload = multer({
  storage: storage,
  limits: { fileSize: 1024 * 1024 * 50 } // 50MB
});

// Upload template
router.post('/upload', auth, upload.single('file'), async (req, res) => {
  try {
    const { filename, originalname, path: filePath, size, mimetype } = req.file;
    const decodedOriginalname = Buffer.from(originalname, 'latin1').toString('utf8');

    const template = new Template({
      filename,
      originalname: decodedOriginalname,
      path: filePath,
      size,
      type: mimetype,
      uploadedBy: req.user.id
    });

    await template.save();
    res.json(template);
  } catch (err) {
    console.error(err.message);
    res.status(500).send('Server Error');
  }
});

// Get template list
router.get('/list', auth, async (req, res) => {
  try {
    const templates = await Template.find().sort({ createdAt: -1 }).populate('uploadedBy', 'username');
    res.json(templates);
  } catch (err) {
    console.error(err.message);
    res.status(500).send('Server Error');
  }
});

// Download template
router.get('/download/:id', auth, async (req, res) => {
  try {
    const template = await Template.findById(req.params.id);
    if (!template) return res.status(404).json({ msg: 'Template not found' });

    if (fs.existsSync(template.path)) {
      res.download(template.path, template.originalname);
    } else {
      res.status(404).json({ msg: 'File not found on server' });
    }
  } catch (err) {
    console.error(err.message);
    res.status(500).send('Server Error');
  }
});

// Preview template
router.get('/preview/:id', auth, async (req, res) => {
  try {
    const template = await Template.findById(req.params.id);
    
    if (!template) {
      return res.status(404).json({ msg: 'Template not found' });
    }
    
    if (!fs.existsSync(template.path)) {
      return res.status(404).json({ msg: 'File not found on server' });
    }
    
    let { sheet, limit, headerRowStart, headerRowEnd, password } = req.query;
    
    if (template.originalname.endsWith('.xlsx') || template.originalname.endsWith('.xls')) {
      let workbook;
      try {
        workbook = xlsx.readFile(template.path, password ? { password } : undefined);
      } catch (readError) {
        if (readError.message.includes('Password') || readError.message.includes('encrypted')) {
          return res.status(400).json({ msg: 'File is password protected', needPassword: true });
        }
        throw readError;
      }
      
      const sheetName = sheet || workbook.SheetNames[0];
      
      if (!workbook.Sheets[sheetName]) {
        return res.status(404).json({ msg: 'Sheet not found' });
      }
      
      const rangeStart = headerRowStart !== undefined ? parseInt(headerRowStart) : 0;
      const rangeEnd = headerRowEnd !== undefined ? parseInt(headerRowEnd) : rangeStart;
      
      const rowLimit = parseInt(limit) || 30;
      
      const sheetObj = workbook.Sheets[sheetName];
      ensureRangeStartsFromA(sheetObj);
      
      const { headers } = getHeaders(sheetObj, rangeStart, rangeEnd);
      
      // Use explicit headers, reading from the next row
      const json = xlsx.utils.sheet_to_json(sheetObj, { header: headers, range: rangeEnd + 1, defval: '' });
      
      const previewData = json.slice(0, rowLimit);
      
      res.json({
        filename: template.originalname,
        sheet: sheetName,
        totalRows: json.length,
        sheetCount: workbook.SheetNames.length,
        sheets: workbook.SheetNames,
        preview: previewData
      });
    } else {
      res.status(400).json({ msg: 'Preview supported only for Excel files' });
    }
  } catch (err) {
    console.error(err.message);
    res.status(500).send('Server Error');
  }
});

// Delete template
router.delete('/:id', auth, async (req, res) => {
  try {
    const template = await Template.findById(req.params.id);
    if (!template) return res.status(404).json({ msg: 'Template not found' });

    // Only uploader or admin can delete
    if (template.uploadedBy.toString() !== req.user.id && req.user.role !== 'admin') {
      return res.status(401).json({ msg: 'Unauthorized' });
    }

    if (fs.existsSync(template.path)) {
      fs.unlinkSync(template.path);
    }

    await Template.findByIdAndDelete(req.params.id);
    res.json({ msg: 'Template deleted' });
  } catch (err) {
    console.error(err.message);
    res.status(500).send('Server Error');
  }
});

module.exports = router;
