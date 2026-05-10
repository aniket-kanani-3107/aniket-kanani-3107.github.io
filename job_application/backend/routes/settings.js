const express = require('express');
const { logActivity } = require('../services/activityService');
const { startScheduler } = require('../services/scheduler');

const router = express.Router();

router.get('/', (req, res) => {
  const db = req.app.locals.db;
  const settings = db
    .prepare('SELECT * FROM settings ORDER BY id DESC LIMIT 1')
    .get();
  res.json(settings || {});
});

router.put('/', (req, res) => {
  const db = req.app.locals.db;
  const { target_roles, locations, resume_path, scraping_schedule, api_keys } = req.body;

  db.prepare(
    `INSERT INTO settings (target_roles, locations, resume_path, scraping_schedule, api_keys, updated_at)
     VALUES (?, ?, ?, ?, ?, datetime('now'))`
  ).run(target_roles, locations, resume_path, scraping_schedule, api_keys);

  logActivity(db, 'settings', 'Settings updated.');
  startScheduler(db);
  res.json({ status: 'ok' });
});

module.exports = router;
