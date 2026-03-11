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
const templateRoutes = require('./routes/template');

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
app.use('/api/template', templateRoutes);

// 全局错误处理中间件
app.use((err, req, res, next) => {
  console.error('Global Error Handler:', err);
  res.status(500).json({ msg: 'Server Error: ' + err.message });
});

// 所有其他 GET 请求返回 index.html (SPA 支持)
app.get(/(.*)/, (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// 连接数据库
console.log('正在连接数据库...');

process.on('exit', (code) => {
  console.log(`Process exited with code: ${code}`);
});

process.on('uncaughtException', (err) => {
  console.error('Uncaught Exception:', err);
});

mongoose.connect('mongodb://localhost:27017/data-analysis-system').then(() => {
  console.log('数据库连接成功');
  
  // 启动服务器
  const PORT = process.env.PORT || 6060;
  console.log(`正在启动服务器，监听端口 ${PORT}...`);
  const server = app.listen(PORT, '0.0.0.0', () => {
    console.log(`服务器运行在端口 ${PORT} (0.0.0.0)`);
  });

  server.on('error', (e) => {
    console.error('Server error:', e);
  });
}).catch(err => {
  console.error('数据库连接失败:', err);
  // process.exit(1); // 不要退出，保持尝试
});

// 防止进程退出
setInterval(() => {
  // 保持心跳
}, 10000);