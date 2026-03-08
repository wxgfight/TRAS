const http = require('http');

const testUrl = (path) => {
  return new Promise((resolve, reject) => {
    http.get(`http://localhost:5000${path}`, (res) => {
      console.log(`GET ${path}: Status ${res.statusCode}, Content-Type: ${res.headers['content-type']}`);
      res.resume();
      resolve();
    }).on('error', (e) => {
      console.error(`GET ${path}: Error ${e.message}`);
      resolve();
    });
  });
};

(async () => {
  await testUrl('/');
  await testUrl('/assets/index-BXoPYq-1.js'); // Note: hash might be different, check file name
  await testUrl('/api/data/list'); // Should be 401 Unauthorized
})();
