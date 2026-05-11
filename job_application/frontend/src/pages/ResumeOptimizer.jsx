import { useEffect, useState } from 'react';
import api from '../api/client';
import PageShell from '../components/layout/PageShell';
import { Button } from '../components/ui/button';
import { Textarea } from '../components/ui/textarea';

const buildPdfPayload = (resumeText) => {
  const lines = resumeText.split('\n').filter(Boolean);
  const name = lines[0] || 'Candidate Name';
  const summary = lines[1] || 'Professional summary will appear here.';
  const bullets = lines.filter((line) => line.startsWith('-')).map((line) => line.replace(/^-\s*/, ''));

  return {
    name,
    summary,
    sections: [
      {
        title: 'Experience Highlights',
        content: [
          {
            type: 'list',
            items: bullets.length ? bullets : ['Add bullet points to generate a richer PDF.'],
          },
        ],
      },
    ],
  };
};

const ResumeOptimizer = () => {
  const [resumeText, setResumeText] = useState('');
  const [jobDescription, setJobDescription] = useState('');
  const [prompts, setPrompts] = useState(null);
  const [pdfStatus, setPdfStatus] = useState('');
  const [optimizedText, setOptimizedText] = useState('');
  const [versionName, setVersionName] = useState('Tailored Resume');
  const [versions, setVersions] = useState([]);

  const loadVersions = async () => {
    const response = await api.get('/resumes');
    setVersions(response.data);
  };

  useEffect(() => {
    loadVersions();
  }, []);

  const uploadResume = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    const formData = new FormData();
    formData.append('file', file);
    const response = await api.post('/resumes/upload', formData);
    setResumeText(response.data.parsedText || '');
    loadVersions();
  };

  const generatePrompts = async () => {
    const response = await api.post('/resumes/optimize', { resumeText, jobDescription });
    setPrompts(response.data);
  };

  const generatePdf = async () => {
    setPdfStatus('Generating PDF...');
    const payload = buildPdfPayload(resumeText);
    const response = await api.post('/pdf/generate', payload);
    setPdfStatus(`Saved PDF: ${response.data.filename}`);
  };

  const saveVersion = async () => {
    if (!optimizedText.trim()) return;
    await api.post('/resumes/save', {
      content: optimizedText,
      version_name: versionName,
    });
    setOptimizedText('');
    loadVersions();
  };

  return (
    <div className="flex flex-col gap-6">
      <PageShell
        title="Resume Optimizer"
        subtitle="Generate Claude-ready prompts and polished PDFs while preserving factual experience."
        action={
          <div className="flex flex-wrap gap-3">
            <Button variant="secondary" onClick={generatePrompts}>Generate Prompts</Button>
            <Button onClick={generatePdf}>Export PDF</Button>
          </div>
        }
      >
        <div className="grid gap-4 lg:grid-cols-2">
          <div>
            <p className="mb-2 text-xs text-slate-400">Resume Text</p>
            <Textarea value={resumeText} onChange={(e) => setResumeText(e.target.value)} />
            <label className="mt-3 inline-flex cursor-pointer items-center gap-2 text-xs text-accentSoft">
              <input type="file" className="hidden" onChange={uploadResume} />
              Upload PDF or DOCX
            </label>
          </div>
          <div>
            <p className="mb-2 text-xs text-slate-400">Job Description</p>
            <Textarea value={jobDescription} onChange={(e) => setJobDescription(e.target.value)} />
          </div>
        </div>
      </PageShell>

      {prompts && (
        <PageShell title="Generated Prompts" subtitle="Paste these into Claude for manual AI refinement.">
          <div className="grid gap-4 md:grid-cols-2">
            <div className="glass-card p-4">
              <p className="text-sm font-semibold text-white">Resume Optimization Prompt</p>
              <Textarea className="mt-3 min-h-[220px]" value={prompts.resumePrompt} readOnly />
            </div>
            <div className="glass-card p-4">
              <p className="text-sm font-semibold text-white">Cover Letter Prompt</p>
              <Textarea className="mt-3 min-h-[220px]" value={prompts.coverLetterPrompt} readOnly />
            </div>
          </div>
        </PageShell>
      )}

      <PageShell
        title="Save Tailored Version"
        subtitle="Paste Claude's rewritten resume and save it as a version."
        action={<Button onClick={saveVersion}>Save Version</Button>}
      >
        <div className="grid gap-4 lg:grid-cols-[1fr_220px]">
          <Textarea
            value={optimizedText}
            onChange={(e) => setOptimizedText(e.target.value)}
            placeholder="Paste Claude's tailored resume here."
          />
          <div>
            <p className="mb-2 text-xs text-slate-400">Version Name</p>
            <input
              className="glass-input"
              value={versionName}
              onChange={(e) => setVersionName(e.target.value)}
            />
          </div>
        </div>
      </PageShell>

      <PageShell title="Resume Version Manager" subtitle="Stored locally for quick reuse.">
        <div className="grid gap-3 md:grid-cols-2">
          {versions.map((version) => (
            <div key={version.id} className="rounded-2xl border border-white/10 bg-white/5 p-4">
              <p className="text-sm font-semibold text-white">{version.version_name}</p>
              <p className="text-xs text-slate-400">Saved: {version.created_at}</p>
              <p className="mt-2 text-xs text-slate-500">File: {version.file_path}</p>
            </div>
          ))}
          {!versions.length && (
            <p className="text-sm text-slate-400">No resume versions saved yet.</p>
          )}
        </div>
      </PageShell>

      {pdfStatus && (
        <div className="rounded-2xl border border-white/10 bg-white/5 px-4 py-3 text-sm text-slate-200">
          {pdfStatus}
        </div>
      )}
    </div>
  );
};

export default ResumeOptimizer;
