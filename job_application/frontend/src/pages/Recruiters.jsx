import { useEffect, useState } from 'react';
import api from '../api/client';
import PageShell from '../components/layout/PageShell';
import { Button } from '../components/ui/button';
import { Input } from '../components/ui/input';

const Recruiters = () => {
  const [recruiters, setRecruiters] = useState([]);
  const [form, setForm] = useState({
    company: '',
    recruiter_name: '',
    email: '',
    status: 'prospect',
    last_contacted: '',
    notes: '',
  });

  const loadRecruiters = async () => {
    const response = await api.get('/recruiters');
    setRecruiters(response.data);
  };

  useEffect(() => {
    loadRecruiters();
  }, []);

  const updateField = (field, value) => {
    setForm((prev) => ({ ...prev, [field]: value }));
  };

  const saveRecruiter = async () => {
    await api.post('/recruiters', form);
    setForm({ company: '', recruiter_name: '', email: '', status: 'prospect', last_contacted: '', notes: '' });
    loadRecruiters();
  };

  return (
    <div className="flex flex-col gap-6">
      <PageShell
        title="Recruiter Tracker"
        subtitle="Keep tabs on recruiter touchpoints and outreach status."
        action={<Button onClick={saveRecruiter}>Add Recruiter</Button>}
      >
        <div className="grid gap-4 md:grid-cols-2">
          <div>
            <p className="mb-2 text-xs text-slate-400">Company</p>
            <Input value={form.company} onChange={(e) => updateField('company', e.target.value)} />
          </div>
          <div>
            <p className="mb-2 text-xs text-slate-400">Recruiter Name</p>
            <Input value={form.recruiter_name} onChange={(e) => updateField('recruiter_name', e.target.value)} />
          </div>
          <div>
            <p className="mb-2 text-xs text-slate-400">Email</p>
            <Input value={form.email} onChange={(e) => updateField('email', e.target.value)} />
          </div>
          <div>
            <p className="mb-2 text-xs text-slate-400">Status</p>
            <select
              className="glass-input w-full"
              value={form.status}
              onChange={(e) => updateField('status', e.target.value)}
            >
              <option value="prospect">Prospect</option>
              <option value="contacted">Contacted</option>
              <option value="responsive">Responsive</option>
              <option value="inactive">Inactive</option>
            </select>
          </div>
          <div>
            <p className="mb-2 text-xs text-slate-400">Last Contacted</p>
            <Input value={form.last_contacted} onChange={(e) => updateField('last_contacted', e.target.value)} />
          </div>
          <div>
            <p className="mb-2 text-xs text-slate-400">Notes</p>
            <Input value={form.notes} onChange={(e) => updateField('notes', e.target.value)} />
          </div>
        </div>
      </PageShell>

      <PageShell title="Recruiter Pipeline">
        <div className="grid gap-3 md:grid-cols-2">
          {recruiters.map((recruiter) => (
            <div key={recruiter.id} className="rounded-2xl border border-white/10 bg-white/5 p-4">
              <p className="text-sm font-semibold text-white">{recruiter.company}</p>
              <p className="text-xs text-slate-400">{recruiter.recruiter_name} · {recruiter.email}</p>
              <p className="mt-2 text-xs text-slate-300">Status: {recruiter.status}</p>
              <p className="text-xs text-slate-500">Last contacted: {recruiter.last_contacted || '—'}</p>
            </div>
          ))}
        </div>
      </PageShell>
    </div>
  );
};

export default Recruiters;
