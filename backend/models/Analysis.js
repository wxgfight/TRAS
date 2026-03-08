
const mongoose = require('mongoose');

const analysisSchema = new mongoose.Schema({
  userId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: 'User',
    required: true
  },
  fileId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: 'Data',
    required: true
  },
  sheet: {
    type: String,
    required: true
  },
  chartType: {
    type: String,
    enum: ['bar', 'line', 'pie', 'scatter'],
    required: true
  },
  xAxis: {
    type: String,
    required: true
  },
  yAxis: {
    type: String,
    required: true
  },
  aggregation: {
    type: String,
    enum: ['count', 'sum', 'avg', 'max', 'min'],
    required: true
  },
  result: {
    type: Object, // 存储分析结果数据 { xAxis: [], series: [], title: '' }
    required: true
  },
  createdAt: {
    type: Date,
    default: Date.now
  }
});

module.exports = mongoose.model('Analysis', analysisSchema);
