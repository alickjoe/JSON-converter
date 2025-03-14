const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const app = express();
const fs = require('fs');
const { promisify } = require('util');
const readdir = promisify(fs.readdir);
const unlink = promisify(fs.unlink);
const rmdir = promisify(fs.rmdir);

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
  res.render('index', { jsonData: null, error: null, cleanupMessage: null });
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
    // 转换完成后立即清理文件
    if (req.file) {
      await unlink(req.file.path);
    }
    res.render('index', {
      jsonData: JSON.stringify(jsonData, null, 2),
      error: null,
      cleanupMessage: '文件已自动清理'
    });
  } catch (error) {
    res.render('index', {
      jsonData: null,
      error: error.message,      
      cleanupMessage: null
    });
  }
});

// 在路由配置部分添加清理路由
// app.post('/cleanup', async (req, res) => {
//   try {
//     const uploadDir = path.join(__dirname, 'uploads');

//     // 检查目录是否存在
//     if (!fs.existsSync(uploadDir)) {
//       return res.render('index', {
//         jsonData: null,
//         error: null,
//         cleanupMessage: '上传目录不存在，无需清理'
//       });
//     }

//     // 获取所有文件
//     const files = await readdir(uploadDir);

//     // 删除所有文件
//     await Promise.all(files.map(file =>
//       unlink(path.join(uploadDir, file))
//     ));

//     // 可选：删除空目录
//     // await rmdir(uploadDir);

//     res.render('index', {
//       jsonData: null,
//       error: null,
//       cleanupMessage: `成功清理 ${files.length} 个临时文件`
//     });
//   } catch (error) {
//     res.render('index', {
//       jsonData: null,
//       error: null,
//       cleanupMessage: `清理失败: ${error.message}`
//     });
//   }
// });

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});