const express = require('express');
const path = require('path');
const app = express();

const distPath = path.join(__dirname);
app.use(express.static(__dirname));

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'taskpane.html'));
});


app.get('/auth-callback.html', (req, res) => {
  res.sendFile(path.join(__dirname, 'auth-callback.html'));
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
