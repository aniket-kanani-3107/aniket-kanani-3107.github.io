import { useState } from 'react';
import api from '../api/client';
import PageShell from '../components/layout/PageShell';
import { Button } from '../components/ui/button';
import { Input } from '../components/ui/input';
import { Badge } from '../components/ui/badge';

const Scraper = () => {
  const [filters, setFilters] = useState({
    role: 'Software Engineer',
    location: 'Remote',
    remote: true,
    includeWellfound: false,
  });
  const [results, setResults] = useState([]);
  const [loading, setLoading] = useState(false);

  const handleChange = (field, value) => {
    setFilters((prev) => ({ ...prev, [field]: value }));
  };

  const runScrape = async () => {
    setLoading(true);
    const response = await api.post('/jobs/scrape', filters);
    setResults(response.data.jobs);
    setLoading(false);
  };

  return (
    <div className="flex flex-col gap-6">
      <PageShell
        title="Job Scraper"
        subtitle="Pull fresh roles from Greenhouse, Lever, and Wellfound. Filter to the last 48 hours."
        action={
          <Button onClick={runScrape} disabled={loading}>
            {loading ? 'Scanning...' : 'Run Job Hunt'}
          </Button>
        }
      >
        <div className="grid gap-4 md:grid-cols-4">
          <div>
            <p className="mb-2 text-xs text-slate-400">Target Role</p>
            <Input value={filters.role} onChange={(e) => handleChange('role', e.target.value)} />
          </div>
          <div>
            <p className="mb-2 text-xs text-slate-400">Location</p>
            <Input value={filters.location} onChange={(e) => handleChange('location', e.target.value)} />
          </div>
          <div>
            <p className="mb-2 text-xs text-slate-400">Remote Only</p>
            <select
              className="glass-input w-full"
              value={filters.remote ? 'yes' : 'no'}
              onChange={(e) => handleChange('remote', e.target.value === 'yes')}
            >
              <option value="yes">Remote</option>
              <option value="no">Any</option>
            </select>
          </div>
          <div>
            <p className="mb-2 text-xs text-slate-400">Include Wellfound</p>
            <select
              className="glass-input w-full"
              value={filters.includeWellfound ? 'yes' : 'no'}
              onChange={(e) => handleChange('includeWellfound', e.target.value === 'yes')}
            >
              <option value="no">Skip</option>
              <option value="yes">Include</option>
            </select>
          </div>
        </div>
      </PageShell>

      <PageShell
        title="Latest Matches"
        subtitle="Results stored locally in SQLite. Duplicates are automatically removed."
      >
        <div className="flex flex-col gap-3">
          {results.map((job) => (
            <div key={job.url} className="rounded-2xl border border-white/10 bg-white/5 p-4">
              <div className="flex flex-wrap items-center justify-between gap-2">
                <div>
                  <p className="text-sm font-semibold text-white">{job.title}</p>
                  <p className="text-xs text-slate-400">{job.company} • {job.location}</p>
                </div>
                <Badge>{job.source}</Badge>
              </div>
              <a className="mt-2 inline-block text-xs text-accentSoft" href={job.url} target="_blank" rel="noreferrer">
                View posting
              </a>
            </div>
          ))}
          {!results.length && (
            <p className="text-sm text-slate-400">Run a scan to load jobs.</p>
          )}
        </div>
      </PageShell>
    </div>
  );
};

export default Scraper;
