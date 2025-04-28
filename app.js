const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const { promisify } = require('util');
const crypto = require('crypto');
const app = express();

// 配置和中间件
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ extended: true, limit: '50mb' }));
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'ejs');

// 辅助函数
const generateUniqueSuffix = () => crypto.randomBytes(3).toString('hex');

const getStructureHash = (obj) => {
  if (Array.isArray(obj) && obj.some(item => typeof item !== 'object')) {
    return crypto.createHash('md5').update('primitive_array').digest('hex').slice(0, 6);
  }
  if ((Array.isArray(obj) && obj.length === 0) || (typeof obj === 'object' && obj !== null && Object.keys(obj).length === 0)) {
    return crypto.createHash('md5').update('empty').digest('hex').slice(0, 6);
  }
  if (Array.isArray(obj) && typeof obj[0] === 'object') {
    const keys = Object.keys(obj[0]).sort();
    const types = keys.map(key => typeof obj[0][key]);
    return crypto.createHash('md5').update(JSON.stringify({ keys, types })).digest('hex').slice(0, 6);
  }
  if (typeof obj === 'object' && obj !== null) {
    const keys = Object.keys(obj).sort();
    const types = keys.map(key => typeof obj[key]);
    return crypto.createHash('md5').update(JSON.stringify({ keys, types })).digest('hex').slice(0, 6);
  }
  return crypto.createHash('md5').update(typeof obj).digest('hex').slice(0, 6);
};

const generateSheetName = (fieldName, structureHash) => {
  const base = fieldName.replace(/[\\\/:*?[\]]/g, '').replace(/_/g, '').substring(0, 20);
  return `${base}_${structureHash}`;
};

// 核心处理函数
const processNestedData = (data, fieldName, parentId, workbook, globalStructureMap = new Map()) => {
  if (!data || data.length === 0) return '';

  // 处理基本类型数组
  const isPrimitiveArray = Array.isArray(data) && data.some(item => typeof item !== 'object');
  if (isPrimitiveArray) {
    const sheetName = generateSheetName(fieldName, getStructureHash(data));
    if (!globalStructureMap.has(sheetName)) {
      const worksheet = XLSX.utils.json_to_sheet([], { header: ["parent_id", "value"] });
      globalStructureMap.set(sheetName, { worksheet, data: [] });
    }
    const currentSheet = globalStructureMap.get(sheetName);
    data.forEach(value => {
      const row = { parent_id: parentId, value };
      currentSheet.data.push(row);
      XLSX.utils.sheet_add_json(currentSheet.worksheet, [row], { skipHeader: true, origin: -1 });
    });
    if (!workbook.Sheets[sheetName]) {
      XLSX.utils.book_append_sheet(workbook, currentSheet.worksheet, sheetName);
    }
    return sheetName;
  }

  // 处理对象/嵌套数组
  const sheetName = generateSheetName(fieldName, getStructureHash(data[0]));
  if (!globalStructureMap.has(sheetName)) {
    const headers = Object.keys(data[0]).sort();
    const worksheet = XLSX.utils.json_to_sheet([], { header: ['parent_id', ...headers] });
    globalStructureMap.set(sheetName, { worksheet, data: [] });
  }

  const currentSheet = globalStructureMap.get(sheetName);
  data.forEach(item => {
    const row = { parent_id: parentId };
    Object.entries(item).forEach(([key, value]) => {
      if (Array.isArray(value)) {
        const childRef = processNestedData(value, key, item.id || parentId, workbook, globalStructureMap);
        row[`${key}_ref`] = childRef;
      } else if (typeof value === 'object' && value !== null) {
        Object.entries(value).forEach(([subKey, subValue]) => {
          row[`${key}_${subKey}`] = subValue;
        });
      } else {
        row[key] = value;
      }
    });
    currentSheet.data.push(row);
    XLSX.utils.sheet_add_json(currentSheet.worksheet, [row], { skipHeader: true, origin: -1 });
  });

  if (!workbook.Sheets[sheetName]) {
    XLSX.utils.book_append_sheet(workbook, currentSheet.worksheet, sheetName);
  }
  return sheetName;
};

// 文件上传配置 - 统一使用uploads目录
const storage = multer.diskStorage({
  destination: 'uploads/',
  filename: (req, file, cb) => cb(null, `${Date.now()}-${file.originalname}`)
});

const upload = multer({
  storage,
  limits: { fileSize: 50 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const allowedTypes = [
      'application/json',
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'application/vnd.ms-excel'
    ];
    
    allowedTypes.includes(file.mimetype)
      ? cb(null, true)
      : cb(new Error('请上传JSON或Excel文件'));
  }
});

// 改进的通用解析函数
const parseGenericExcel = (workbook) => {
  const sheets = workbook.SheetNames;
  const sheetData = {};
  
  // 1. 收集所有工作表数据
  sheets.forEach(sheetName => {
    sheetData[sheetName] = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
  });

  // 2. 自动检测关联关系
  const detectRelations = () => {
    const relations = {};
    
    // 遍历所有工作表
    Object.entries(sheetData).forEach(([sheetName, rows]) => {
      if (!rows.length) return;
      
      const firstRow = rows[0];
      
      // 检测可能的关联字段（包含_id或_ref后缀的字段）
      Object.keys(firstRow).forEach(key => {
        if (key.endsWith('_id') || key.endsWith('_ref')) {
          const potentialRef = firstRow[key];
          
          // 如果_ref指向的工作表存在
          if (typeof potentialRef === 'string' && sheetData[potentialRef]) {
            relations[sheetName] = relations[sheetName] || [];
            relations[sheetName].push({
              sourceField: key,
              targetSheet: potentialRef,
              relationType: key.endsWith('_ref') ? 'nested' : 'reference'
            });
          }
        }
      });
    });
    
    return relations;
  };

  // 3. 智能推断关联键
  const inferRelationKeys = (sourceRow, targetSheet) => {
    const possibleKeys = [
      'id', '_generated_id', 'parent_id', 
      'location_id', 'evse_uid', 'connector_id'
    ];
    
    // 尝试找到两个表共有的字段
    const targetFirstRow = sheetData[targetSheet][0] || {};
    const commonFields = Object.keys(sourceRow).filter(
      key => key in targetFirstRow
    );
    
    // 优先使用预设键，然后使用共有字段
    const candidateKeys = [...possibleKeys, ...commonFields];
    
    for (const key of candidateKeys) {
      if (key in sourceRow && sheetData[targetSheet].some(
        row => row[key] === sourceRow[key]
      )) {
        return key;
      }
    }
    
    // 默认使用第一个字段
    return Object.keys(targetFirstRow)[0];
  };

  // 4. 构建嵌套结构
  const buildNestedStructure = (data, relations, currentSheet) => {
    return data.map(item => {
      const nestedItem = { ...item };
      
      // 处理当前表的所有关联关系
      (relations[currentSheet] || []).forEach(relation => {
        const { sourceField, targetSheet, relationType } = relation;
        const relationKey = inferRelationKeys(item, targetSheet);
        
        const relatedItems = sheetData[targetSheet].filter(
          row => row[relationKey] === item[relationKey] || 
                 row[relationKey] === item[sourceField]
        );
        
        if (relatedItems.length) {
          const fieldName = sourceField.replace(/_id$|_ref$/, '');
          
          if (relationType === 'nested') {
            // 嵌套关联
            nestedItem[fieldName] = buildNestedStructure(
              relatedItems, 
              relations, 
              targetSheet
            );
          } else {
            // 简单引用
            nestedItem[fieldName] = relatedItems.length === 1 
              ? relatedItems[0] 
              : relatedItems;
          }
          
          // 移除原始引用字段
          delete nestedItem[sourceField];
        }
      });
      
      return nestedItem;
    });
  };

  // 5. 主流程
  const relations = detectRelations();
  const mainSheet = sheets.find(sheet => sheet === 'MainData') || sheets[0];
  
  return buildNestedStructure(sheetData[mainSheet], relations, mainSheet);
};

// 清理上传目录
const cleanupUploads = async () => {
  try {
    const files = await promisify(fs.readdir)('uploads');
    const now = Date.now();
    
    for (const file of files) {
      const filePath = path.join('uploads', file);
      const stats = await promisify(fs.stat)(filePath);
      // 删除超过1小时的文件
      if (now - stats.mtimeMs > 3600000) {
        await promisify(fs.unlink)(filePath);
      }
    }
  } catch (error) {
    console.error('清理上传目录失败:', error);
  }
};

// 路由配置
app.get('/', (req, res) => res.render('index', { jsonData: null, error: null, cleanupMessage: null }));

app.post('/upload', upload.single('excelFile'), async (req, res) => {
  try {
    if (!req.file) throw new Error('请选择要上传的文件');

    let jsonData;
    if (req.file.mimetype === 'application/json') {
      // 处理JSON文件
      const fileContent = await promisify(fs.readFile)(req.file.path, 'utf8');
      jsonData = JSON.parse(fileContent);
    } else {
      // 处理Excel文件
      const workbook = XLSX.readFile(req.file.path);
      jsonData = parseGenericExcel(workbook);
    }

    // 上传后立即删除文件
    await promisify(fs.unlink)(req.file.path);

    res.render('index', {
      jsonData: JSON.stringify(jsonData, null, 2),
      error: null,
      cleanupMessage: '上传文件已自动清理'
    });
  } catch (error) {
    // 出错时也尝试删除文件
    if (req.file) {
      try {
        await promisify(fs.unlink)(req.file.path);
      } catch (cleanupError) {
        console.error('清理上传文件失败:', cleanupError);
      }
    }
    
    res.render('index', { 
      jsonData: null, 
      error: error.message, 
      cleanupMessage: '上传文件已自动清理'
    });
  }
});

app.post('/generate-excel', async (req, res) => {
  try {
    let jsonData = JSON.parse(req.body.jsonData);
    if (!Array.isArray(jsonData)) jsonData = [jsonData];

    const workbook = XLSX.utils.book_new();
    const globalStructureMap = new Map();
    const idMap = new Map();

    const mainData = jsonData.map((item, index) => {
      const parentId = item.id || `rec_${crypto.randomBytes(4).toString('hex')}`;
      idMap.set(index, parentId);

      const processedItem = { _generated_id: parentId };

      Object.entries(item).forEach(([key, value]) => {
        if (Array.isArray(value)) {
          processedItem[`${key}_ref`] = processNestedData(value, key, parentId, workbook, globalStructureMap);
        } else if (typeof value === 'object' && value !== null) {
          Object.entries(value).forEach(([subKey, subValue]) => {
            processedItem[`${key}_${subKey}`] = subValue;
          });
        } else {
          processedItem[key] = value;
        }
      });

      return processedItem;
    });

    const mainWorksheet = XLSX.utils.json_to_sheet(mainData);
    XLSX.utils.book_append_sheet(workbook, mainWorksheet, "MainData");

    const fileName = `export-${Date.now()}.xlsx`;
    const filePath = path.join(__dirname, 'uploads', fileName);

    XLSX.writeFile(workbook, filePath);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=${fileName}`);

    const fileStream = fs.createReadStream(filePath);
    fileStream.pipe(res);
    
    fileStream.on('end', async () => {
      try {
        await promisify(fs.unlink)(filePath);
      } catch (error) {
        console.error('删除临时Excel文件失败:', error);
      }
    });

  } catch (error) {
    res.render('index', {
      jsonData: req.body.jsonData,
      error: `生成Excel失败: ${error.message}`,
      cleanupMessage: null
    });
  }
});

// 服务器启动
const PORT = process.env.PORT || 3000;
app.listen(PORT, async () => {
  console.log(`Server running on port ${PORT}`);
  
  // 创建必要的目录
  if (!fs.existsSync('uploads')) fs.mkdirSync('uploads');
  
  // 启动时清理旧文件
  await cleanupUploads();
  
  // 每小时清理一次上传目录
  setInterval(cleanupUploads, 3600000);
});
