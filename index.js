
const express = require('express');
const multer = require('multer');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const cors = require('cors');

const app = express();
app.use(cors());
const upload = multer({ storage: multer.memoryStorage() });

app.post('/merge', upload.single('template'), (req, res) => {
  try {
    const templateBuffer = req.file.buffer;
    const data = JSON.parse(req.body.data);

    const zip = new PizZip(templateBuffer);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true
    });

    doc.setData(data);
    doc.render();

    const outputBuffer = doc.getZip().generate({ type: 'nodebuffer' });

    res.setHeader('Content-Disposition', 'attachment; filename=merged.docx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.send(outputBuffer);
  } catch (error) {
    console.error('Merge error:', error);
    res.status(500).send('Failed to merge template.');
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Docx Merge API running on port ${PORT}`));
