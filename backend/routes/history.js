const express = require('express');
const router = express.Router();
const Comparison = require('../models/Comparison');
const auth = require('../middleware/auth');

// 获取历史记录列表
router.get('/list', auth, async (req, res) => {
  try {
    const { page = 1, limit = 10, search, startDate, endDate, userId } = req.query;
    
    const query = {};
    
    // 如果不是管理员，只能查看自己的记录
    if (req.user.role !== 'admin') {
      query.userId = req.user.id;
    } else if (userId) {
      // 管理员可以按特定用户筛选
      query.userId = userId;
    }
    
    // 搜索条件
    if (search) {
      // 可以在这里增加对文件名搜索的支持，但需要联表查询，比较复杂
      // 这里暂时只支持对 remark 的搜索
      query.remark = { $regex: search, $options: 'i' };
    }
    
    // 时间范围
    if (startDate && endDate) {
      query.createdAt = {
        $gte: new Date(startDate),
        $lte: new Date(endDate)
      };
    } else if (startDate) {
      query.createdAt = { $gte: new Date(startDate) };
    } else if (endDate) {
      query.createdAt = { $lte: new Date(endDate) };
    }
    
    const total = await Comparison.countDocuments(query);
    const comparisons = await Comparison.find(query)
      .populate('file1Id', 'originalname')
      .populate('file2Id', 'originalname')
      .sort({ createdAt: -1 })
      .skip((page - 1) * limit)
      .limit(parseInt(limit));
    
    res.json({
      total,
      page: parseInt(page),
      limit: parseInt(limit),
      data: comparisons
    });
  } catch (err) {
    console.error(err.message);
    res.status(500).send('服务器错误');
  }
});

// 获取历史记录详情
router.get('/:id', auth, async (req, res) => {
  try {
    const comparison = await Comparison.findById(req.params.id)
      .populate('file1Id', 'originalname path')
      .populate('file2Id', 'originalname path');
    
    if (!comparison) {
      return res.status(404).json({ msg: '记录不存在' });
    }
    
    // 检查权限
    if (comparison.userId.toString() !== req.user.id && req.user.role !== 'admin') {
      return res.status(401).json({ msg: '无权访问此记录' });
    }
    
    res.json(comparison);
  } catch (err) {
    console.error(err.message);
    res.status(500).send('服务器错误');
  }
});

// 重新对比
router.post('/:id/rerun', auth, async (req, res) => {
  try {
    const comparison = await Comparison.findById(req.params.id);
    
    if (!comparison) {
      return res.status(404).json({ msg: '记录不存在' });
    }
    
    // 检查权限
    if (comparison.userId.toString() !== req.user.id && req.user.role !== 'admin') {
      return res.status(401).json({ msg: '无权操作此记录' });
    }
    
    // 这里可以添加重新对比的逻辑
    // 暂时返回成功
    res.json({ msg: '重新对比任务已启动' });
  } catch (err) {
    console.error(err.message);
    res.status(500).send('服务器错误');
  }
});

// 删除历史记录
router.delete('/:id', auth, async (req, res) => {
  try {
    const comparison = await Comparison.findById(req.params.id);
    
    if (!comparison) {
      return res.status(404).json({ msg: '记录不存在' });
    }
    
    // 检查权限
    if (comparison.userId.toString() !== req.user.id && req.user.role !== 'admin') {
      return res.status(401).json({ msg: '无权删除此记录' });
    }
    
    await Comparison.deleteOne({ _id: req.params.id });
    res.json({ msg: '记录删除成功' });
  } catch (err) {
    console.error(err.message);
    res.status(500).send('服务器错误');
  }
});

module.exports = router;