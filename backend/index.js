const express = require('express');
const mongoose = require('mongoose');
const cors = require('cors');
const path = require('path');
const authRoutes = require('./routes/auth');
const dataRoutes = require('./routes/data');
const compareRoutes = require('./routes/compare');
const historyRoutes = require('./routes/history');
const dashboardRoutes = require('./routes/dashboard');
const analysisRoutes = require('./routes/analysis');

const app = express();

// 中间件
app.use(cors());
app.use(express.json());

// 静态文件托管 - 必须放在 API 路由之前，或者是特定路径
// 这里我们将整个 public 目录托管为根路径
app.use(express.static(path.join(__dirname, 'public')));

// API 路由
app.use('/api/auth', authRoutes);
app.use('/api/data', dataRoutes);
app.use('/api/compare', compareRoutes);
app.use('/api/history', historyRoutes);
app.use('/api/dashboard', dashboardRoutes);
app.use('/api/analysis', analysisRoutes);

// 所有其他 GET 请求返回 index.html (SPA 支持)
app.get(/(.*)/, (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// 连接数据库
mongoose.connect('mongodb://localhost:27017/data-analysis-system').then(() => {
  console.log('数据库连接成功');
  
  // 启动服务器
  const PORT = process.env.PORT || 5000;
  app.listen(PORT, () => {
    console.log(`服务器运行在端口 ${PORT}`);
  });
}).catch(err => {
  console.error('数据库连接失败:', err);
  process.exit(1);
});