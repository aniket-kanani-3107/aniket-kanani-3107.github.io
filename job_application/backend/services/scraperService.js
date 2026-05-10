const fs = require('fs');
const path = require('path');
const { fetchGreenhouseJobs, normalizeGreenhouseJob } = require('../scrapers/greenhouse');
const { fetchLeverJobs, normalizeLeverJob } = require('../scrapers/lever');
const { fetchWellfoundJobs, normalizeWellfoundJob } = require('../scrapers/wellfound');
const { logActivity } = require('./activityService');

const sourcesPath = path.join(__dirname, '..', 'storage', 'sources.json');
const HOURS_48_MS = 48 * 60 * 60 * 1000;

const defaultSources = {
  greenhouse: [
    { token: 'stripe', label: 'Stripe' },
    { token: 'figma', label: 'Figma' },
  ],
  lever: [
    { token: 'notion', label: 'Notion' },
    { token: 'robinhood', label: 'Robinhood' },
  ],
};

const loadSources = () => {
  if (!fs.existsSync(sourcesPath)) {
    fs.writeFileSync(sourcesPath, JSON.stringify(defaultSources, null, 2));
    return defaultSources;
  }
  return JSON.parse(fs.readFileSync(sourcesPath, 'utf-8'));
};

const isRecent = (dateString) => {
  const date = new Date(dateString);
  if (Number.isNaN(date.getTime())) return true;
  return Date.now() - date.getTime() <= HOURS_48_MS;
};

const matchFilters = (job, filters) => {
  const roleMatch = filters.role
    ? job.title.toLowerCase().includes(filters.role.toLowerCase())
    : true;
  const locationMatch = filters.location
    ? job.location.toLowerCase().includes(filters.location.toLowerCase())
    : true;
  const remoteMatch = filters.remote
    ? job.location.toLowerCase().includes('remote')
    : true;

  return roleMatch && locationMatch && remoteMatch && isRecent(job.date_found);
};

const scrapeAllSources = async (db, filters) => {
  const sources = loadSources();
  const results = [];

  for (const greenhouse of sources.greenhouse || []) {
    try {
      const jobs = await fetchGreenhouseJobs(greenhouse.token);
      jobs.forEach((job) => {
        results.push(normalizeGreenhouseJob(job, greenhouse.label));
      });
    } catch (error) {
      logActivity(db, 'scrape_error', `Greenhouse ${greenhouse.label} failed.`);
    }
  }

  for (const lever of sources.lever || []) {
    try {
      const jobs = await fetchLeverJobs(lever.token);
      jobs.forEach((job) => {
        results.push(normalizeLeverJob(job, lever.label));
      });
    } catch (error) {
      logActivity(db, 'scrape_error', `Lever ${lever.label} failed.`);
    }
  }

  if (filters.includeWellfound) {
    try {
      const jobs = await fetchWellfoundJobs(filters);
      jobs.forEach((job) => {
        results.push(normalizeWellfoundJob(job));
      });
    } catch (error) {
      logActivity(db, 'scrape_error', 'Wellfound scraping failed.');
    }
  }

  const filtered = results.filter((job) => matchFilters(job, filters));
  const insert = db.prepare(`
    INSERT OR IGNORE INTO jobs (company, title, location, url, description, date_found, status, ats_score, source)
    VALUES (@company, @title, @location, @url, @description, @date_found, @status, @ats_score, @source)
  `);

  const insertedJobs = [];
  const transaction = db.transaction((jobs) => {
    jobs.forEach((job) => {
      const info = insert.run({ ...job, status: 'saved', ats_score: null });
      if (info.changes > 0) {
        insertedJobs.push(job);
        logActivity(db, 'job_found', `New job found: ${job.title} at ${job.company}`);
      }
    });
  });

  transaction(filtered);

  return insertedJobs;
};

module.exports = { scrapeAllSources, loadSources };
