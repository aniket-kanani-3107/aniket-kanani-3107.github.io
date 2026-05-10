const express = require('express');
const { logActivity } = require('../services/activityService');

const router = express.Router();

router.get('/', (req, res) => {
  const db = req.app.locals.db;
  const applications = db
    .prepare(
      `SELECT applications.*, jobs.company, jobs.title, jobs.location
       FROM applications
       JOIN jobs ON jobs.id = applications.job_id
       ORDER BY applications.updated_at DESC`
    )
    .all();
  res.json(applications);
});

router.post('/', (req, res) => {
  const db = req.app.locals.db;
  const { job_id, status } = req.body;

  if (!job_id) {
    return res.status(400).json({ error: 'job_id is required.' });
  }

  const info = db
    .prepare('INSERT INTO applications (job_id, status) VALUES (?, ?)')
    .run(job_id, status || 'saved');

  logActivity(db, 'application', `Application created for job ${job_id}`);
  return res.json({ id: info.lastInsertRowid });
});

router.patch('/:id/status', (req, res) => {
  const db = req.app.locals.db;
  const { status } = req.body;
  const { id } = req.params;

  db.prepare(
    'UPDATE applications SET status = ?, updated_at = datetime(\'now\') WHERE id = ?'
  ).run(status, id);

  logActivity(db, 'application', `Application ${id} status updated to ${status}`);
  res.json({ status: 'ok' });
});

module.exports = router;
