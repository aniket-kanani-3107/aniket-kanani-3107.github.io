const fs = require('fs');
const path = require('path');

const storageRoot = path.join(__dirname, '..', 'storage');

const folders = [
  storageRoot,
  path.join(storageRoot, 'uploads'),
  path.join(storageRoot, 'generated'),
  path.join(storageRoot, 'prompts'),
  path.join(storageRoot, 'resumes'),
  path.join(storageRoot, 'logs'),
];

const ensureStorage = () => {
  folders.forEach((folder) => {
    if (!fs.existsSync(folder)) {
      fs.mkdirSync(folder, { recursive: true });
    }
  });
};

module.exports = { ensureStorage, storageRoot };
