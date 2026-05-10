import { useState } from 'react';
import api from '../api/client';
import PageShell from '../components/layout/PageShell';
import { Button } from '../components/ui/button';
import { Textarea } from '../components/ui/textarea';
import { Badge } from '../components/ui/badge';

const AtsAnalyzer = () => {
  const [resumeText, setResumeText] = useState('');
  const [jobDescription, setJobDescription] = useState('');
  const [analysis, setAnalysis] = useState(null);

  const uploadResume = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    const formData = new FormData();
    formData.append('file', file);
    const response = await api.post('/resumes/upload', formData);
    setResumeText(response.data.parsedText);
  };

  const analyze = async () => {
    const response = await api.post('/ats/analyze', { resumeText, jobDescription });
    setAnalysis(response.data);
  };

  return (
    <div className="flex flex-col gap-6">
      <PageShell
        title="ATS Analyzer"
        subtitle="Upload a resume, paste a job description, and score the match."
        action={<Button onClick={analyze}>Analyze Match</Button>}
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

      {analysis && (
        <PageShell title="ATS Results" subtitle="Keyword coverage and improvement guidance.">
          <div className="grid gap-4 md:grid-cols-3">
            <div className="glass-card p-4 text-center">
              <p className="text-sm text-slate-400">Match Score</p>
              <p className="mt-2 text-3xl font-semibold text-white">{analysis.score}%</p>
            </div>
            <div className="glass-card p-4">
              <p className="text-sm text-slate-400">Missing Keywords</p>
              <div className="mt-2 flex flex-wrap gap-2">
                {analysis.missingKeywords.slice(0, 12).map((keyword) => (
                  <Badge key={keyword}>{keyword}</Badge>
                ))}
              </div>
            </div>
            <div className="glass-card p-4">
              <p className="text-sm text-slate-400">Suggestions</p>
              <ul className="mt-2 space-y-2 text-xs text-slate-300">
                {analysis.suggestions.map((suggestion) => (
                  <li key={suggestion.keyword}>{suggestion.recommendation}</li>
                ))}
              </ul>
            </div>
          </div>
        </PageShell>
      )}
    </div>
  );
};

export default AtsAnalyzer;
