const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const app = express();

// 配置EJS模板引擎
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'ejs');

// 配置静态文件目录
//app.use(express.static('public'));

// 配置Multer文件上传
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, 'uploads/');
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + '-' + file.originalname);
  }
});

const upload = multer({
  storage: storage,
  fileFilter: (req, file, cb) => {
    if (file.mimetype.includes('excel') || file.mimetype.includes('spreadsheetml')) {
      cb(null, true);
    } else {
      cb(new Error('请上传Excel文件'));
    }
  }
});

// 路由配置
app.get('/', (req, res) => {
  res.render('index', { jsonData: null, error: null });
});

app.post('/upload', upload.single('excelFile'), async (req, res) => {
  try {
    if (!req.file) {
      throw new Error('请选择要上传的文件');
    }

    // 读取Excel文件
    const workbook = XLSX.readFile(req.file.path);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    // 转换为JSON
    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    res.render('index', {
      jsonData: JSON.stringify(jsonData, null, 2),
      error: null
    });
  } catch (error) {
    res.render('index', {
      jsonData: null,
      error: error.message
    });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});