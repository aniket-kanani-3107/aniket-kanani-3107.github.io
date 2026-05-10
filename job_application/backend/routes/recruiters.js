const express = require('express');
const { logActivity } = require('../services/activityService');

const router = express.Router();

router.get('/', (req, res) => {
  const db = req.app.locals.db;
  const recruiters = db
    .prepare('SELECT * FROM recruiters ORDER BY created_at DESC')
    .all();
  res.json(recruiters);
});

router.post('/', (req, res) => {
  const db = req.app.locals.db;
  const { company, recruiter_name, email, status, last_contacted, notes } = req.body;

  const info = db
    .prepare(
      `INSERT INTO recruiters (company, recruiter_name, email, status, last_contacted, notes)
       VALUES (?, ?, ?, ?, ?, ?)`
    )
    .run(company, recruiter_name, email, status, last_contacted, notes);

  logActivity(db, 'recruiter', `Recruiter added for ${company}`);
  res.json({ id: info.lastInsertRowid });
});

module.exports = router;
