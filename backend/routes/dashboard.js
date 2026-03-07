const express = require('express');
const router = express.Router();
const Data = require('../models/Data');
const Comparison = require('../models/Comparison');
const User = require('../models/User');
const auth = require('../middleware/auth');

// 获取仪表盘数据
router.get('/stats', auth, async (req, res) => {
  try {
    // 统计概览
    const totalFiles = await Data.countDocuments({ userId: req.user.id });
    const totalComparisons = await Comparison.countDocuments({ userId: req.user.id });
    
    // 今日任务数
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayComparisons = await Comparison.countDocuments({
      userId: req.user.id,
      createdAt: { $gte: today }
    });
    
    // 近7天任务趋势
    const last7Days = [];
    for (let i = 6; i >= 0; i--) {
      const date = new Date();
      date.setDate(date.getDate() - i);
      date.setHours(0, 0, 0, 0);
      
      const nextDate = new Date(date);
      nextDate.setDate(nextDate.getDate() + 1);
      
      const count = await Comparison.countDocuments({
        userId: req.user.id,
        createdAt: { $gte: date, $lt: nextDate }
      });
      
      last7Days.push({
        date: date.toISOString().split('T')[0],
        count
      });
    }
    
    // 差异类型分布
    const comparisonStats = await Comparison.find({ userId: req.user.id }, 'stats');
    let totalAdded = 0;
    let totalDeleted = 0;
    let totalModified = 0;
    
    comparisonStats.forEach(stat => {
      if (stat.stats) {
        totalAdded += stat.stats.added || 0;
        totalDeleted += stat.stats.deleted || 0;
        totalModified += stat.stats.modified || 0;
      }
    });
    
    // 管理员可以查看所有用户的统计数据
    if (req.user.role === 'admin') {
      const adminTotalFiles = await Data.countDocuments();
      const adminTotalComparisons = await Comparison.countDocuments();
      const adminTodayComparisons = await Comparison.countDocuments({
        createdAt: { $gte: today }
      });
      
      const adminComparisonStats = await Comparison.find({}, 'stats');
      let adminTotalAdded = 0;
      let adminTotalDeleted = 0;
      let adminTotalModified = 0;
      
      adminComparisonStats.forEach(stat => {
        if (stat.stats) {
          adminTotalAdded += stat.stats.added || 0;
          adminTotalDeleted += stat.stats.deleted || 0;
          adminTotalModified += stat.stats.modified || 0;
        }
      });
      
      return res.json({
        user: {
          totalFiles,
          totalComparisons,
          todayComparisons,
          last7Days,
          differenceDistribution: {
            added: totalAdded,
            deleted: totalDeleted,
            modified: totalModified
          }
        },
        admin: {
          totalFiles: adminTotalFiles,
          totalComparisons: adminTotalComparisons,
          todayComparisons: adminTodayComparisons,
          differenceDistribution: {
            added: adminTotalAdded,
            deleted: adminTotalDeleted,
            modified: adminTotalModified
          }
        }
      });
    }
    
    res.json({
      totalFiles,
      totalComparisons,
      todayComparisons,
      last7Days,
      differenceDistribution: {
        added: totalAdded,
        deleted: totalDeleted,
        modified: totalModified
      }
    });
  } catch (err) {
    console.error(err.message);
    res.status(500).send('服务器错误');
  }
});

// 获取文件对比趋势
router.get('/comparison-trend', auth, async (req, res) => {
  try {
    // 这里可以添加更详细的趋势分析
    res.json({ msg: '趋势分析功能开发中' });
  } catch (err) {
    console.error(err.message);
    res.status(500).send('服务器错误');
  }
});

module.exports = router;