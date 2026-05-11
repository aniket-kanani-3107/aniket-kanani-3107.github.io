const express = require('express');
const cors = require('cors');
const { ensureStorage } = require('../services/storage');
const { initDb } = require('../services/db');
const { startScheduler } = require('../services/scheduler');

const jobsRoutes = require('../routes/jobs');
const atsRoutes = require('../routes/ats');
const resumeRoutes = require('../routes/resumes');
const promptRoutes = require('../routes/prompts');
const applicationRoutes = require('../routes/applications');
const pdfRoutes = require('../routes/pdf');
const dashboardRoutes = require('../routes/dashboard');
const settingsRoutes = require('../routes/settings');
const recruiterRoutes = require('../routes/recruiters');
const searchRoutes = require('../routes/searches');

ensureStorage();
const db = initDb();

const app = express();
const PORT = process.env.PORT || 5000;

app.use(cors({ origin: 'http://localhost:5173' }));
app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true }));

app.locals.db = db;

app.get('/api/health', (req, res) => res.json({ status: 'ok' }));

app.use('/api/jobs', jobsRoutes);
app.use('/api/ats', atsRoutes);
app.use('/api/resumes', resumeRoutes);
app.use('/api/prompts', promptRoutes);
app.use('/api/applications', applicationRoutes);
app.use('/api/pdf', pdfRoutes);
app.use('/api/dashboard', dashboardRoutes);
app.use('/api/settings', settingsRoutes);
app.use('/api/recruiters', recruiterRoutes);
app.use('/api/searches', searchRoutes);

startScheduler(db);

app.listen(PORT, () => {
  console.log(`Job Hunter backend running on http://localhost:${PORT}`);
});
