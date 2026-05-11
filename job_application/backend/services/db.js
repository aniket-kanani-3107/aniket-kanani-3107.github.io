const fs = require('fs');
const path = require('path');
const Database = require('better-sqlite3');

const dbPath = path.join(__dirname, '..', 'database', 'job_hunter.db');
const seedPath = path.join(__dirname, '..', '..', 'jobs', 'sample_jobs.json');

const seedSampleData = (db) => {
  const count = db.prepare('SELECT COUNT(*) as count FROM jobs').get();
  if (count.count > 0) {
    return;
  }

  if (!fs.existsSync(seedPath)) {
    return;
  }

  const sampleJobs = JSON.parse(fs.readFileSync(seedPath, 'utf-8'));
  const insertJob = db.prepare(`
    INSERT INTO jobs (company, title, location, url, description, date_found, status, ats_score, source)
    VALUES (@company, @title, @location, @url, @description, @date_found, @status, @ats_score, @source)
  `);
  const insertApplication = db.prepare(`
    INSERT INTO applications (job_id, status, stage_notes)
    VALUES (@job_id, @status, @stage_notes)
  `);

  const transaction = db.transaction((jobs) => {
    jobs.forEach((job) => {
      const info = insertJob.run(job);
      insertApplication.run({
        job_id: info.lastInsertRowid,
        status: job.status || 'saved',
        stage_notes: 'Sample application seeded for dashboard preview.',
      });
    });
  });

  transaction(sampleJobs);
};

const initDb = () => {
  const db = new Database(dbPath);
  db.pragma('journal_mode = WAL');

  db.exec(`
    CREATE TABLE IF NOT EXISTS jobs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      company TEXT NOT NULL,
      title TEXT NOT NULL,
      location TEXT,
      url TEXT UNIQUE NOT NULL,
      description TEXT,
      date_found TEXT,
      status TEXT DEFAULT 'saved',
      ats_score INTEGER,
      source TEXT,
      created_at TEXT DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS applications (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      job_id INTEGER NOT NULL,
      status TEXT DEFAULT 'saved',
      stage_notes TEXT,
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now')),
      FOREIGN KEY (job_id) REFERENCES jobs(id)
    );

    CREATE TABLE IF NOT EXISTS resumes (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      job_id INTEGER,
      version_name TEXT,
      file_path TEXT,
      parsed_text TEXT,
      created_at TEXT DEFAULT (datetime('now')),
      FOREIGN KEY (job_id) REFERENCES jobs(id)
    );

    CREATE TABLE IF NOT EXISTS prompts (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      type TEXT,
      content TEXT,
      created_at TEXT DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS search_history (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      role TEXT,
      location TEXT,
      remote INTEGER,
      include_wellfound INTEGER,
      created_at TEXT DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS settings (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      target_roles TEXT,
      locations TEXT,
      resume_path TEXT,
      scraping_schedule TEXT,
      api_keys TEXT,
      created_at TEXT DEFAULT (datetime('now')),
      updated_at TEXT DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS recruiters (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      company TEXT,
      recruiter_name TEXT,
      email TEXT,
      status TEXT,
      last_contacted TEXT,
      notes TEXT,
      created_at TEXT DEFAULT (datetime('now'))
    );

    CREATE TABLE IF NOT EXISTS activity (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      type TEXT,
      message TEXT,
      created_at TEXT DEFAULT (datetime('now'))
    );
  `);

  seedSampleData(db);

  return db;
};

module.exports = { initDb };
