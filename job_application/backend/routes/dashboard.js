const express = require('express');
const { listActivity } = require('../services/activityService');

const router = express.Router();

router.get('/', (req, res) => {
  const db = req.app.locals.db;
  const totalJobs = db.prepare('SELECT COUNT(*) as count FROM jobs').get().count;
  const appliedJobs = db
    .prepare("SELECT COUNT(*) as count FROM applications WHERE status = 'applied'")
    .get().count;
  const avgAts = db.prepare('SELECT AVG(ats_score) as avg FROM jobs').get().avg || 0;
  const latestJobs = db
    .prepare('SELECT * FROM jobs ORDER BY date_found DESC LIMIT 5')
    .all();
  const recentResumes = db
    .prepare('SELECT id, version_name, created_at FROM resumes ORDER BY created_at DESC LIMIT 5')
    .all();

  res.json({
    metrics: {
      totalJobs,
      appliedJobs,
      avgAts: Math.round(avgAts),
    },
    latestJobs,
    recentResumes,
    activity: listActivity(db),
  });
});

module.exports = router;
