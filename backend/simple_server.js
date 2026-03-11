const express = require('express');
const app = express();

app.get('/', (req, res) => {
  res.send('Hello World');
});

const PORT = 6001;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});

process.on('exit', (code) => {
  console.log(`Process exited with code: ${code}`);
});

setInterval(() => {
  console.log('Heartbeat...');
}, 2000);
