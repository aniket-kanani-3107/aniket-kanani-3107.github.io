const tokenize = (text) =>
  text
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .split(/\s+/)
    .filter((token) => token.length > 2);

const extractKeywords = (text) => {
  const tokens = tokenize(text);
  const counts = tokens.reduce((acc, token) => {
    acc[token] = (acc[token] || 0) + 1;
    return acc;
  }, {});

  return Object.entries(counts)
    .sort((a, b) => b[1] - a[1])
    .slice(0, 40)
    .map(([token]) => token);
};

const analyzeAts = ({ resumeText, jobDescription }) => {
  const resumeTokens = new Set(tokenize(resumeText));
  const jobKeywords = extractKeywords(jobDescription);

  const matched = jobKeywords.filter((keyword) => resumeTokens.has(keyword));
  const missing = jobKeywords.filter((keyword) => !resumeTokens.has(keyword));

  const score =
    jobKeywords.length === 0
      ? 0
      : Math.round((matched.length / jobKeywords.length) * 100);

  const suggestions = missing.slice(0, 8).map((keyword) => ({
    keyword,
    recommendation: `Consider weaving "${keyword}" into a relevant bullet or skills section.`,
  }));

  return {
    score,
    matchedKeywords: matched,
    missingKeywords: missing,
    suggestions,
  };
};

module.exports = { analyzeAts, extractKeywords };
