import { useState } from 'react';
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

      {pdfStatus && (
        <div className="rounded-2xl border border-white/10 bg-white/5 px-4 py-3 text-sm text-slate-200">
          {pdfStatus}
        </div>
      )}
    </div>
  );
};

export default ResumeOptimizer;
