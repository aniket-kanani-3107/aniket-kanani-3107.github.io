const cron = require('node-cron');
const { scrapeAllSources } = require('./scraperService');
const { logActivity } = require('./activityService');

let task = null;

const startScheduler = (db) => {
  const settings = db.prepare('SELECT * FROM settings ORDER BY id DESC LIMIT 1').get();
  const schedule = settings?.scraping_schedule || '0 8 * * *';

  if (task) {
    task.stop();
  }

  task = cron.schedule(schedule, async () => {
    try {
      logActivity(db, 'automation', 'Running scheduled job hunt.');
      await scrapeAllSources(db, {
        role: settings?.target_roles || '',
        location: settings?.locations || '',
        remote: true,
        includeWellfound: false,
      });
      logActivity(db, 'automation', 'Scheduled scan completed.');
    } catch (error) {
      logActivity(db, 'automation_error', 'Scheduled scan failed.');
    }
  });

  return task;
};

module.exports = { startScheduler };
