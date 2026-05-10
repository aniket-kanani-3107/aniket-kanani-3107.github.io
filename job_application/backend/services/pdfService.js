const fs = require('fs');
const path = require('path');
const puppeteer = require('puppeteer');

const buildResumeHtml = ({ name, summary, sections }) => `
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <style>
      body { font-family: "Inter", Arial, sans-serif; margin: 32px; color: #0f172a; }
      h1 { font-size: 24px; margin-bottom: 4px; }
      h2 { font-size: 16px; margin-top: 18px; border-bottom: 1px solid #e2e8f0; padding-bottom: 4px; }
      p { margin: 6px 0; line-height: 1.5; }
      ul { margin: 6px 0 12px 18px; }
      li { margin-bottom: 4px; }
      .summary { font-size: 13px; color: #475569; }
    </style>
  </head>
  <body>
    <h1>${name}</h1>
    <p class="summary">${summary}</p>
    ${sections
      .map(
        (section) => `
      <h2>${section.title}</h2>
      ${section.content
        .map((block) =>
          block.type === 'list'
            ? `<ul>${block.items.map((item) => `<li>${item}</li>`).join('')}</ul>`
            : `<p>${block.text}</p>`
        )
        .join('')}
    `
      )
      .join('')}
  </body>
</html>
`;

const generatePdf = async ({ html, outputPath }) => {
  const browser = await puppeteer.launch({ headless: 'new' });
  const page = await browser.newPage();
  await page.setContent(html, { waitUntil: 'networkidle0' });
  await page.pdf({
    path: outputPath,
    format: 'A4',
    printBackground: true,
    margin: { top: '20mm', bottom: '20mm', left: '16mm', right: '16mm' },
  });
  await browser.close();
};

const createResumePdf = async ({ payload, outputDir }) => {
  const safeName = payload.name.replace(/[^a-z0-9-_]/gi, '_').toLowerCase();
  const filename = `${safeName}_resume_${Date.now()}.pdf`;
  const outputPath = path.join(outputDir, filename);

  const html = buildResumeHtml(payload);
  await generatePdf({ html, outputPath });

  return { filePath: outputPath, filename };
};

module.exports = { createResumePdf, buildResumeHtml };
