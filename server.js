const express = require('express');
const multer = require('multer');
const path = require('path');
const { processAttendance } = require('./processor');

const app = express();
const port = process.env.PORT || 3000;

// Setup multer for memory storage
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

app.use(express.static('public'));

app.post('/upload', upload.single('file'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).send('No file uploaded.');
    }

    const outputBuffer = await processAttendance(req.file.buffer);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=ketqua.xlsx');
    res.send(outputBuffer);
  } catch (error) {
    console.error('Error processing file:', error);
    res.status(500).send('An error occurred during processing.');
  }
});

app.listen(port, () => {
  console.log(`Server is running at http://localhost:${port}`);
});
