const jwt = require('jsonwebtoken');

module.exports = function(req, res, next) {
  // 从请求头获取token
  const token = req.header('x-auth-token');
  
  // 检查token是否存在
  if (!token) {
    return res.status(401).json({ msg: '没有token，拒绝访问' });
  }
  
  try {
    // 验证token
    const decoded = jwt.verify(token, 'secret');
    req.user = decoded;
    next();
  } catch (err) {
    res.status(401).json({ msg: 'token无效' });
  }
};