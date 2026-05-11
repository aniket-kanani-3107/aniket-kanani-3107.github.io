const express = require('express');
const fs = require('fs');
const path = require('path');
const multer = require('multer');
const { parseResume } = require('../services/resumeParser');
const { buildResumePrompt, buildCoverLetterPrompt } = require('../services/promptService');
const { logActivity } = require('../services/activityService');

const router = express.Router();
const upload = multer({
  dest: path.join(__dirname, '..', 'storage', 'uploads'),
});

router.get('/', (req, res) => {
  const db = req.app.locals.db;
  const resumes = db
    .prepare('SELECT id, job_id, version_name, file_path, created_at FROM resumes ORDER BY created_at DESC')
    .all();
  res.json(resumes);
});

router.post('/upload', upload.single('file'), async (req, res) => {
  const db = req.app.locals.db;
  if (!req.file) {
    return res.status(400).json({ error: 'File is required.' });
  }

  const parsedText = await parseResume(req.file.path);
  const versionName = req.body.version_name || req.file.originalname;

  const info = db
    .prepare(
      'INSERT INTO resumes (job_id, version_name, file_path, parsed_text) VALUES (?, ?, ?, ?)'
    )
    .run(req.body.job_id || null, versionName, req.file.path, parsedText);

  logActivity(db, 'resume_upload', `Resume uploaded: ${versionName}`);
  return res.json({ id: info.lastInsertRowid, parsedText });
});

router.post('/save', (req, res) => {
  const db = req.app.locals.db;
  const { content, version_name, job_id } = req.body;

  if (!content) {
    return res.status(400).json({ error: 'Resume content is required.' });
  }

  const safeName = (version_name || 'optimized')
    .replace(/[^a-z0-9-_]/gi, '_')
    .toLowerCase();
  const filename = `${safeName}_${Date.now()}.txt`;
  const filePath = path.join(__dirname, '..', 'storage', 'resumes', filename);

  fs.writeFileSync(filePath, content, 'utf-8');

  const info = db
    .prepare(
      'INSERT INTO resumes (job_id, version_name, file_path, parsed_text) VALUES (?, ?, ?, ?)'
    )
    .run(job_id || null, version_name || 'Optimized Resume', filePath, content);

  logActivity(db, 'resume_saved', `Resume saved: ${version_name || 'Optimized Resume'}`);
  return res.json({ id: info.lastInsertRowid, filePath });
});

router.post('/optimize', (req, res) => {
  const db = req.app.locals.db;
  const { resumeText, jobDescription } = req.body;

  if (!resumeText || !jobDescription) {
    return res.status(400).json({ error: 'Resume text and job description are required.' });
  }

  const resumePrompt = buildResumePrompt({ resumeText, jobDescription });
  const coverLetterPrompt = buildCoverLetterPrompt({ resumeText, jobDescription });

  db.prepare('INSERT INTO prompts (type, content) VALUES (?, ?)')
    .run('resume_optimization', resumePrompt);
  db.prepare('INSERT INTO prompts (type, content) VALUES (?, ?)')
    .run('cover_letter', coverLetterPrompt);

  logActivity(db, 'prompt', 'Generated resume optimization prompts.');

  res.json({ resumePrompt, coverLetterPrompt });
});

module.exports = router;
