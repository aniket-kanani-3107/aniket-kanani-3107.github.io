const express = require('express');

const router = express.Router();

router.get('/', (req, res) => {
  const db = req.app.locals.db;
  const history = db
    .prepare(
      `SELECT id, role, location, remote, include_wellfound, created_at
       FROM search_history
       ORDER BY created_at DESC
       LIMIT 10`
    )
    .all();
  res.json(history);
});

module.exports = router;
