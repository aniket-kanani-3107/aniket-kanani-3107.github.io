const fetchGreenhouseJobs = async (boardToken) => {
  const url = `https://boards-api.greenhouse.io/v1/boards/${boardToken}/jobs?content=true`;
  const response = await fetch(url);

  if (!response.ok) {
    throw new Error(`Greenhouse fetch failed for ${boardToken}`);
  }

  const data = await response.json();
  return data.jobs || [];
};

const normalizeGreenhouseJob = (job, sourceLabel) => ({
  company: sourceLabel,
  title: job.title,
  location: job.location?.name || '',
  url: job.absolute_url,
  description: job.content || '',
  date_found: job.updated_at || job.created_at || new Date().toISOString(),
  source: 'Greenhouse',
});

module.exports = { fetchGreenhouseJobs, normalizeGreenhouseJob };
