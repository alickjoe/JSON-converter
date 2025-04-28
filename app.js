const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');
const { promisify } = require('util');
const crypto = require('crypto');
const app = express();

// 增加请求体大小限制配置
app.use(express.json({ limit: '50mb' }));  // 增加到50MB
app.use(express.urlencoded({ extended: true, limit: '50mb' }));  // 增加到50MB

// 中间件配置
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'ejs');

// 辅助函数 ---------------------------------------------------
const generateUniqueSuffix = () => {
  return crypto.randomBytes(3).toString('hex');
};

// 增强版结构哈希算法
const getStructureHash = (obj) => {
  // 处理基本类型数组的特殊标记
  if (Array.isArray(obj) && obj.some(item => typeof item !== 'object')) {
    return crypto.createHash('md5').update('primitive_array').digest('hex').slice(0,6);
  }

  // 处理空数组/对象
  if ((Array.isArray(obj) && obj.length === 0) || (typeof obj === 'object' && obj !== null && Object.keys(obj).length === 0)) {
    return crypto.createHash('md5').update('empty').digest('hex').slice(0,6);
  }

  // 处理对象数组
  if (Array.isArray(obj) && typeof obj[0] === 'object') {
    const sample = obj[0];
    const keys = Object.keys(sample).sort();
    const types = keys.map(key => typeof sample[key]);
    return crypto.createHash('md5')
      .update(JSON.stringify({ keys, types }))
      .digest('hex')
      .slice(0, 6);
  }

  // 处理单个对象
  if (typeof obj === 'object' && obj !== null) {
    const keys = Object.keys(obj).sort();
    const types = keys.map(key => typeof obj[key]);
    return crypto.createHash('md5')
      .update(JSON.stringify({ keys, types }))
      .digest('hex')
      .slice(0, 6);
  }

  return crypto.createHash('md5').update(typeof obj).digest('hex').slice(0,6);
};

// 优化后的工作表命名逻辑
const generateSheetName = (fieldName, structureHash) => {
  const base = fieldName.replace(/[\\\/:*?[\]]/g, '')
    .replace(/_/g, '')
    .substring(0, 20);
  return `${base}_${structureHash}`; // 总长度<=28
};

// 改进后的核心处理函数
const processNestedData = (data, fieldName, parentSheetNames, parentId, workbook, globalStructureMap = new Map()) => {
  if (!data || data.length === 0) return '';

  // 处理基本类型数组（字符串/数字等）
  const isPrimitiveArray = Array.isArray(data) && data.some(item => typeof item !== 'object');
  if (isPrimitiveArray) {
    const structureHash = getStructureHash(data);
    const sheetName = generateSheetName(fieldName, structureHash);

    if (!globalStructureMap.has(sheetName)) {
      const worksheet = XLSX.utils.json_to_sheet([], {
        header: ["parent_id", "value"]
      });
      globalStructureMap.set(sheetName, {
        worksheet,
        data: []
      });
    }

    const currentSheet = globalStructureMap.get(sheetName);
    data.forEach(value => {
      const row = { parent_id: parentId, value };
      currentSheet.data.push(row);
      XLSX.utils.sheet_add_json(currentSheet.worksheet, [row], {
        skipHeader: true,
        origin: -1
      });
    });

    if (!workbook.Sheets[sheetName]) {
      XLSX.utils.book_append_sheet(workbook, currentSheet.worksheet, sheetName);
    }
    return sheetName;
  }

  // 处理对象/嵌套数组
  const structureHash = getStructureHash(data[0]);
  const sheetName = generateSheetName(fieldName, structureHash);

  if (!globalStructureMap.has(sheetName)) {
    // 使用第一个对象的键作为表头
    const headers = Object.keys(data[0]).sort();
    const worksheet = XLSX.utils.json_to_sheet([], { header: ['parent_id', ...headers] });
    globalStructureMap.set(sheetName, {
      worksheet,
      data: []
    });
  }

  const currentSheet = globalStructureMap.get(sheetName);
  
  data.forEach(item => {
    const row = { parent_id: parentId };
    
    Object.entries(item).forEach(([key, value]) => {
      if (Array.isArray(value)) {
        const childRef = processNestedData(
          value,
          key,
          parentSheetNames,
          item.id || parentId,
          workbook,
          globalStructureMap
        );
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
    XLSX.utils.sheet_add_json(currentSheet.worksheet, [row], {
      skipHeader: true,
      origin: -1
    });
  });

  if (!workbook.Sheets[sheetName]) {
    XLSX.utils.book_append_sheet(workbook, currentSheet.worksheet, sheetName);
  }

  return sheetName;
};

// 文件上传配置 -------------------------------------------------
const storage = multer.diskStorage({
  destination: 'uploads/',
  filename: (req, file, cb) => {
    cb(null, `${Date.now()}-${file.originalname}`);
  }
});

const upload = multer({
  storage,
  limits: { fileSize: 50 * 1024 * 1024 }, // 增加文件大小限制到50MB
  fileFilter: (req, file, cb) => {
    if (file.mimetype.includes('excel') || file.mimetype.includes('spreadsheetml')) {
      cb(null, true);
    } else {
      cb(new Error('请上传Excel文件'));
    }
  }
});

// 修改后的解析函数，处理只有parent_id和value列的情况
const parseMultiTabExcel = (workbook) => {
  const sheets = workbook.SheetNames;
  const result = {};
  const relations = {};
  
  // 第一阶段：收集所有工作表数据并识别关系
  sheets.forEach(sheetName => {
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet);
    result[sheetName] = jsonData;
    
    // 检测关系（通过 parent_id 字段）
    if (jsonData.length > 0) {
      const firstRow = jsonData[0];
      const columns = Object.keys(firstRow);
      
      // 如果是只有parent_id和value的特殊表
      if (columns.length === 2 && columns.includes('parent_id') && columns.includes('value')) {
        relations[sheetName] = {
          parentField: 'parent_id',
          childField: sheetName.replace(/_\w+$/, ''), // 移除哈希后缀获取字段名
          isValueOnly: true
        };
      } else if ('parent_id' in firstRow) {
        relations[sheetName] = {
          parentField: 'parent_id',
          childField: sheetName.replace(/_\w+$/, ''),
          isValueOnly: false
        };
      }
    }
  });

  // 第二阶段：构建嵌套结构
  const buildNested = (parentItems, relations) => {
    return parentItems.map(parent => {
      const nested = { ...parent };
      
      // 查找所有子表关系
      Object.entries(relations).forEach(([sheetName, relation]) => {
        const children = result[sheetName].filter(
          child => child[relation.parentField] === parent.id || 
                  child[relation.parentField] === parent._generated_id
        );
        
        if (children.length > 0) {
          // 处理只有value的特殊表
          if (relation.isValueOnly) {
            nested[relation.childField] = children.map(child => child.value);
          } else {
            // 处理普通表
            nested[relation.childField] = children.map(child => {
              const childCopy = { ...child };
              delete childCopy[relation.parentField]; // 移除 parent_id
              
              // 递归处理更深层次的嵌套
              const hasNested = Object.keys(relations).some(
                s => s !== sheetName && result[s][0]?.parent_id === child.id
              );
              
              return hasNested ? buildNested([childCopy], relations)[0] : childCopy;
            });
          }
        }
      });
      
      return nested;
    });
  };
  
  // 第三阶段：确定根表并构建结构
  const rootSheets = sheets.filter(sheet => 
    !Object.values(relations).some(r => r.childField === sheet.replace(/_\w+$/, ''))
  );
  
  return rootSheets.length > 0 
    ? buildNested(result[rootSheets[0]], relations) 
    : result;
};

// 路由配置 -----------------------------------------------------
app.get('/', (req, res) => {
  res.render('index', { jsonData: null, error: null, cleanupMessage: null });
});

app.post('/upload', upload.single('excelFile'), async (req, res) => {
  try {
    if (!req.file) throw new Error('请选择要上传的文件');

    const workbook = XLSX.readFile(req.file.path);
    const jsonData = parseMultiTabExcel(workbook);

    await promisify(fs.unlink)(req.file.path);

    res.render('index', {
      jsonData: JSON.stringify(jsonData, null, 2),
      error: null,
      cleanupMessage: '文件已自动清理'
    });
  } catch (error) {
    res.render('index', { jsonData: null, error: error.message, cleanupMessage: null });
  }
});

app.post('/generate-excel', (req, res) => {
  try {
    let jsonData = JSON.parse(req.body.jsonData);
    if (!Array.isArray(jsonData)) jsonData = [jsonData];

    const workbook = XLSX.utils.book_new();
    const globalStructureMap = new Map();
    const idMap = new Map(); // 新增：用于存储生成的ID映射

    const mainData = jsonData.map((item, index) => {
      // 生成唯一ID并保存映射关系
      const parentId = item.id || `rec_${crypto.randomBytes(4).toString('hex')}`;
      idMap.set(index, parentId);

      const processedItem = {
        _generated_id: parentId // 将生成的ID存入主表
      };

      // 处理所有属性，不再排除特定字段
      Object.entries(item).forEach(([key, value]) => {
        if (Array.isArray(value)) {
          processedItem[`${key}_ref`] = processNestedData(
            value,
            key,
            [],
            parentId,
            workbook,
            globalStructureMap
          );
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

    // 文件生成和下载
    const fileName = `export-${Date.now()}.xlsx`;
    const filePath = path.join(__dirname, 'temp', fileName);

    if (!fs.existsSync(path.join(__dirname, 'temp'))) {
      fs.mkdirSync(path.join(__dirname, 'temp'));
    }

    XLSX.writeFile(workbook, filePath);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename=${fileName}`);

    fs.createReadStream(filePath)
      .pipe(res)
      .on('finish', () => fs.unlinkSync(filePath));

  } catch (error) {
    res.render('index', {
      jsonData: req.body.jsonData,
      error: `生成Excel失败: ${error.message}`,
      cleanupMessage: null
    });
  }
});

// 服务器启动 --------------------------------------------------
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
  if (!fs.existsSync('uploads')) fs.mkdirSync('uploads');
  if (!fs.existsSync('temp')) fs.mkdirSync('temp');
});
