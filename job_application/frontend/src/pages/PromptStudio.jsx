import { useEffect, useState } from 'react';
import api from '../api/client';
import PageShell from '../components/layout/PageShell';
import { Button } from '../components/ui/button';
import { Textarea } from '../components/ui/textarea';

const PromptStudio = () => {
  const [prompts, setPrompts] = useState([]);
  const [content, setContent] = useState('');

  const loadPrompts = async () => {
    const response = await api.get('/prompts');
    setPrompts(response.data);
  };

  useEffect(() => {
    loadPrompts();
  }, []);

  const savePrompt = async () => {
    if (!content.trim()) return;
    await api.post('/prompts', { type: 'custom', content });
    setContent('');
    loadPrompts();
  };

  return (
    <div className="flex flex-col gap-6">
      <PageShell
        title="Prompt Studio"
        subtitle="Store reusable prompts for resume tailoring, outreach, and interview prep."
        action={<Button onClick={savePrompt}>Save Prompt</Button>}
      >
        <Textarea value={content} onChange={(e) => setContent(e.target.value)} />
      </PageShell>

      <PageShell title="Saved Prompts">
        <div className="flex flex-col gap-3">
          {prompts.map((prompt) => (
            <div key={prompt.id} className="rounded-2xl border border-white/10 bg-white/5 p-4">
              <p className="text-xs uppercase text-slate-400">{prompt.type}</p>
              <p className="mt-2 text-sm text-slate-200 whitespace-pre-wrap">{prompt.content}</p>
            </div>
          ))}
        </div>
      </PageShell>
    </div>
  );
};

export default PromptStudio;
