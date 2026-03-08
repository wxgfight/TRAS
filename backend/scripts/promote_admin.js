const mongoose = require('mongoose');
const User = require('../models/User');

async function promoteAdmin() {
  try {
    await mongoose.connect('mongodb://localhost:27017/data-analysis-system');
    const result = await User.updateOne({ username: 'admin' }, { $set: { role: 'admin' } });
    if (result.modifiedCount > 0) {
      console.log('User "admin" has been promoted to administrator role.');
    } else {
      console.log('User "admin" was already an administrator or not found.');
    }
    process.exit(0);
  } catch (err) {
    console.error(err);
    process.exit(1);
  }
}

promoteAdmin();
