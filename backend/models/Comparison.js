const mongoose = require('mongoose');

const comparisonSchema = new mongoose.Schema({
  taskId: {
    type: String,
    required: true,
    unique: true
  },
  userId: {
    type: mongoose.Schema.Types.ObjectId,
    ref: 'User',
    required: true
  },
  file1Id: {
    type: mongoose.Schema.Types.ObjectId,
    ref: 'Data',
    required: true
  },
  file2Id: {
    type: mongoose.Schema.Types.ObjectId,
    ref: 'Data',
    required: true
  },
  sheet1: {
    type: String,
    required: true
  },
  sheet2: {
    type: String,
    required: true
  },
  primaryKey: {
    type: [String],
    required: true
  },
  ignoreColumns: {
    type: [String],
    default: []
  },
  headerRow: {
    type: Number,
    default: 0
  },
  nullStrategy: {
    type: String,
    enum: ['ignore', 'treat-as-empty', 'treat-as-value'],
    default: 'ignore'
  },
  createdAt: {
    type: Date,
    default: Date.now
  },
  stats: {
    type: Object,
    default: {
      file1Rows: 0,
      file2Rows: 0,
      added: 0,
      deleted: 0,
      modified: 0,
      totalDifferences: 0
    }
  },
  reportPath: {
    type: String
  },
  status: {
    type: String,
    enum: ['pending', 'completed', 'failed'],
    default: 'pending'
  }
});

module.exports = mongoose.model('Comparison', comparisonSchema);