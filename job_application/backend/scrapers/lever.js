const fetchLeverJobs = async (company) => {
  const url = `https://api.lever.co/v0/postings/${company}?mode=json`;
  const response = await fetch(url);

  if (!response.ok) {
    throw new Error(`Lever fetch failed for ${company}`);
  }

  return response.json();
};

const normalizeLeverJob = (job, companyLabel) => ({
  company: companyLabel,
  title: job.text,
  location: job.categories?.location || '',
  url: job.hostedUrl,
  description: job.descriptionPlain || job.description || '',
  date_found: job.createdAt ? new Date(job.createdAt).toISOString() : new Date().toISOString(),
  source: 'Lever',
});

module.exports = { fetchLeverJobs, normalizeLeverJob };
