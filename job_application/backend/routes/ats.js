const express = require('express');
const { analyzeAts } = require('../services/atsService');
const { logActivity } = require('../services/activityService');

const router = express.Router();

router.post('/analyze', (req, res) => {
  const db = req.app.locals.db;
  const { resumeText, jobDescription } = req.body;

  if (!resumeText || !jobDescription) {
    return res.status(400).json({ error: 'Resume text and job description are required.' });
  }

  const analysis = analyzeAts({ resumeText, jobDescription });
  logActivity(db, 'ats', 'ATS analysis completed.');
  return res.json(analysis);
});

module.exports = router;
