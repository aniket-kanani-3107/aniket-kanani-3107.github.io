const { extractKeywords } = require('./atsService');

const buildResumePrompt = ({ resumeText, jobDescription }) => {
  const keywords = extractKeywords(jobDescription).slice(0, 15);
  return `You are an expert resume editor. Rewrite the resume to maximize ATS alignment while preserving factual experience. Do not add new experience.

Job description summary keywords: ${keywords.join(', ')}.

Instructions:
- Preserve every employer, title, timeline, and metric.
- Rewrite bullets to emphasize impact and quantify results.
- Add the missing keywords naturally where they truthfully fit.
- Keep the formatting concise.

Resume:
${resumeText}

Return a revised resume in plain text with clear section headings.`;
};

const buildCoverLetterPrompt = ({ resumeText, jobDescription }) => {
  const keywords = extractKeywords(jobDescription).slice(0, 12);
  return `Write a tailored cover letter based on the resume and job description. Keep it professional, concise, and factual.

Key focus keywords: ${keywords.join(', ')}.

Resume:
${resumeText}

Job Description:
${jobDescription}

Return a complete cover letter in plain text.`;
};

module.exports = { buildResumePrompt, buildCoverLetterPrompt };
