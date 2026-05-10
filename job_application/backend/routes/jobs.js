const express = require('express');
const { scrapeAllSources, loadSources } = require('../services/scraperService');
const { logActivity } = require('../services/activityService');

const router = express.Router();

router.get('/', (req, res) => {
  const db = req.app.locals.db;
  const jobs = db
    .prepare('SELECT * FROM jobs ORDER BY date_found DESC, created_at DESC')
    .all();
  res.json(jobs);
});

router.post('/', (req, res) => {
  const db = req.app.locals.db;
  const { company, title, location, url, description, source } = req.body;

  const stmt = db.prepare(`
    INSERT OR IGNORE INTO jobs (company, title, location, url, description, date_found, status, source)
    VALUES (@company, @title, @location, @url, @description, @date_found, 'saved', @source)
  `);

  const info = stmt.run({
    company,
    title,
    location,
    url,
    description,
    date_found: new Date().toISOString(),
    source: source || 'Manual',
  });

  logActivity(db, 'job_saved', `Job saved manually: ${title} at ${company}`);
  res.json({ id: info.lastInsertRowid });
});

router.post('/scrape', async (req, res) => {
  const db = req.app.locals.db;
  const { role, location, remote, includeWellfound } = req.body;
  try {
    const jobs = await scrapeAllSources(db, {
      role: role || '',
      location: location || '',
      remote: Boolean(remote),
      includeWellfound: Boolean(includeWellfound),
    });

    res.json({ count: jobs.length, jobs });
  } catch (error) {
    res.status(500).json({ error: 'Scrape failed', details: error.message });
  }
});

router.get('/sources', (req, res) => {
  res.json(loadSources());
});

module.exports = router;
