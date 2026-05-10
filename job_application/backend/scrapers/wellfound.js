const { chromium } = require('playwright');

const fetchWellfoundJobs = async ({ role, location, remote }) => {
  const params = new URLSearchParams();
  if (role) params.set('query', role);
  if (location) params.set('location', location);
  if (remote) params.set('remote', 'true');

  const url = `https://wellfound.com/jobs${params.toString() ? `?${params.toString()}` : ''}`;
  const browser = await chromium.launch();
  const page = await browser.newPage();

  await page.goto(url, { waitUntil: 'networkidle' });
  const data = await page.evaluate(() => {
    const nextData = window.__NEXT_DATA__;
    if (nextData?.props?.pageProps?.jobs) {
      return nextData.props.pageProps.jobs;
    }
    return null;
  });

  await browser.close();

  const jobs = data?.jobs || data?.listings || data?.items || [];
  return Array.isArray(jobs) ? jobs : [];
};

const normalizeWellfoundJob = (job) => ({
  company: job.startup?.name || job.company_name || 'Wellfound',
  title: job.title || job.role || '',
  location: job.location || job.location_name || '',
  url: job.url || job.redirect_url || job.link || 'https://wellfound.com/jobs',
  description: job.description || job.description_plain || '',
  date_found: job.published_at || job.created_at || new Date().toISOString(),
  source: 'Wellfound',
});

module.exports = { fetchWellfoundJobs, normalizeWellfoundJob };
