const express = require('express');
const { logActivity } = require('../services/activityService');

const router = express.Router();

router.get('/', (req, res) => {
  const db = req.app.locals.db;
  const prompts = db
    .prepare('SELECT id, type, content, created_at FROM prompts ORDER BY created_at DESC')
    .all();
  res.json(prompts);
});

router.post('/', (req, res) => {
  const db = req.app.locals.db;
  const { type, content } = req.body;

  if (!content) {
    return res.status(400).json({ error: 'Prompt content is required.' });
  }

  const info = db.prepare('INSERT INTO prompts (type, content) VALUES (?, ?)')
    .run(type || 'custom', content);

  logActivity(db, 'prompt', 'Saved a custom prompt.');
  return res.json({ id: info.lastInsertRowid });
});

module.exports = router;
