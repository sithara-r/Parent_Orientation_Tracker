const express = require('express');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 8080;

app.use(express.static(__dirname));

// Accept raw binary data up to 10MB
app.use(express.raw({ type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', limit: '10mb' }));

app.post('/api/uploadExcel', (req, res) => {
  try {
    fs.writeFileSync(path.join(__dirname, 'Parent_Orientation_Tracker.xlsx'), req.body);
    res.sendStatus(200);
  } catch (err) {
    console.error('Error saving Excel:', err);
    res.status(500).send('Error saving file.');
  }
});

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'Parent_Orientation_Tracker.html'));
});

app.listen(PORT, () => {
  console.log(`\n======================================================`);
  console.log(`Local Development Server running at http://localhost:${PORT}`);
  console.log(`======================================================`);
  console.log(`NOTE: This server allows you to save data locally to the Excel file.`);
  console.log(`======================================================\n`);
});
