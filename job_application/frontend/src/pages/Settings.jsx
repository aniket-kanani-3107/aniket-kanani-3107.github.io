import { useEffect, useState } from 'react';
import api from '../api/client';
import PageShell from '../components/layout/PageShell';
import { Button } from '../components/ui/button';
import { Input } from '../components/ui/input';
import { Textarea } from '../components/ui/textarea';

const Settings = () => {
  const [form, setForm] = useState({
    target_roles: '',
    locations: '',
    resume_path: '',
    scraping_schedule: '0 8 * * *',
    api_keys: '',
  });
  const [saved, setSaved] = useState(false);

  useEffect(() => {
    const loadSettings = async () => {
      const response = await api.get('/settings');
      if (response.data) {
        setForm((prev) => ({ ...prev, ...response.data }));
      }
    };
    loadSettings();
  }, []);

  const updateField = (field, value) => {
    setForm((prev) => ({ ...prev, [field]: value }));
  };

  const saveSettings = async () => {
    await api.put('/settings', form);
    setSaved(true);
    setTimeout(() => setSaved(false), 2000);
  };

  return (
    <PageShell
      title="Settings"
      subtitle="Configure targets, schedules, and local file paths."
      action={<Button onClick={saveSettings}>Save Settings</Button>}
    >
      <div className="grid gap-4 md:grid-cols-2">
        <div>
          <p className="mb-2 text-xs text-slate-400">Target Roles</p>
          <Input value={form.target_roles || ''} onChange={(e) => updateField('target_roles', e.target.value)} />
        </div>
        <div>
          <p className="mb-2 text-xs text-slate-400">Locations</p>
          <Input value={form.locations || ''} onChange={(e) => updateField('locations', e.target.value)} />
        </div>
        <div>
          <p className="mb-2 text-xs text-slate-400">Resume Path</p>
          <Input value={form.resume_path || ''} onChange={(e) => updateField('resume_path', e.target.value)} />
        </div>
        <div>
          <p className="mb-2 text-xs text-slate-400">Scraping Schedule (cron)</p>
          <Input value={form.scraping_schedule || ''} onChange={(e) => updateField('scraping_schedule', e.target.value)} />
        </div>
      </div>
      <div className="mt-4">
        <p className="mb-2 text-xs text-slate-400">API Keys (stored locally)</p>
        <Textarea value={form.api_keys || ''} onChange={(e) => updateField('api_keys', e.target.value)} />
      </div>
      {saved && <p className="mt-3 text-xs text-emerald-300">Settings saved.</p>}
    </PageShell>
  );
};

export default Settings;
