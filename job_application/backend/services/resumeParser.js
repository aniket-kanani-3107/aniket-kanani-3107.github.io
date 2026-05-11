const fs = require('fs');
const path = require('path');
const mammoth = require('mammoth');
const pdf = require('pdf-parse');

const uploadsDir = path.join(__dirname, '..', 'storage', 'uploads');

const ensureSafePath = (filePath) => {
  const resolvedPath = path.resolve(filePath);
  const resolvedUploads = path.resolve(uploadsDir);
  if (!resolvedPath.startsWith(resolvedUploads)) {
    throw new Error('Invalid resume path.');
  }
};

const parseResume = async (filePath) => {
  ensureSafePath(filePath);
  const ext = path.extname(filePath).toLowerCase();

  if (ext === '.pdf') {
    const dataBuffer = fs.readFileSync(filePath);
    const data = await pdf(dataBuffer);
    return data.text || '';
  }

  if (ext === '.docx') {
    const data = await mammoth.extractRawText({ path: filePath });
    return data.value || '';
  }

  if (ext === '.txt' || ext === '.md') {
    return fs.readFileSync(filePath, 'utf-8');
  }

  return '';
};

module.exports = { parseResume };
