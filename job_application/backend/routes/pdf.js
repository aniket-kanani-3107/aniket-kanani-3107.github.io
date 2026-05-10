const express = require('express');
const path = require('path');
const { createResumePdf } = require('../services/pdfService');
const { logActivity } = require('../services/activityService');

const router = express.Router();

router.post('/generate', async (req, res) => {
  const db = req.app.locals.db;
  const payload = req.body;

  if (!payload?.name || !payload?.sections) {
    return res.status(400).json({ error: 'Resume payload is required.' });
  }

  try {
    const outputDir = path.join(__dirname, '..', 'storage', 'generated');
    const result = await createResumePdf({ payload, outputDir });

    db.prepare('INSERT INTO resumes (version_name, file_path) VALUES (?, ?)')
      .run(payload.versionName || 'PDF Export', result.filePath);

    logActivity(db, 'pdf', `Generated PDF resume: ${result.filename}`);
    res.json(result);
  } catch (error) {
    res.status(500).json({ error: 'PDF generation failed', details: error.message });
  }
});

module.exports = router;
