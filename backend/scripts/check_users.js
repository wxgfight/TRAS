const mongoose = require('mongoose');
const User = require('../models/User');

async function checkUsers() {
  try {
    await mongoose.connect('mongodb://localhost:27017/data-analysis-system');
    const users = await User.find({}, 'username role');
    if (users.length === 0) {
      console.log('No users found in database.');
    } else {
      console.log('Current users in database:');
      users.forEach(u => console.log(`- ${u.username}: ${u.role}`));
    }
    process.exit(0);
  } catch (err) {
    console.error(err);
    process.exit(1);
  }
}

checkUsers();
