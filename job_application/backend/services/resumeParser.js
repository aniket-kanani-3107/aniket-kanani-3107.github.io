const fs = require('fs');
const path = require('path');
const mammoth = require('mammoth');
const pdf = require('pdf-parse');

const uploadsDir = path.join(__dirname, '..', 'storage', 'uploads');

const resolveUploadPath = (filename) => {
  const safeFilename = path.basename(filename);
  const resolvedUploads = path.resolve(uploadsDir);
  const filePath = path.join(resolvedUploads, safeFilename);
  if (!filePath.startsWith(resolvedUploads)) {
    throw new Error('Invalid resume path.');
  }
  return filePath;
};

const parseResume = async (filename) => {
  const filePath = resolveUploadPath(filename);
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
