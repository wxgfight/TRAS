const axios = require('axios');

async function testRegister() {
  try {
    const response = await axios.post('http://localhost:5000/api/auth/register', {
      username: 'test',
      email: 'test@example.com',
      password: '123456'
    });
    console.log('注册成功:', response.data);
  } catch (error) {
    console.error('注册失败:', error.response ? error.response.data : error.message);
  }
}

testRegister();