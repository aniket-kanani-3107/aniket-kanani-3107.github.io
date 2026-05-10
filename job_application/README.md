# Offline AI Job Hunter

Local-first job hunting workspace with a React + Node.js + SQLite stack. Everything runs on your machine, with scraping and manual AI workflows.

## Getting Started

```bash
cd job_application
npm install
cd backend && npm install
cd ../frontend && npm install
```

Run both servers:

```bash
cd job_application
npm run app
```

- Frontend: `http://localhost:5173`
- Backend: `http://localhost:5000`

## Folder Structure

```
job_application/
├── backend/
├── frontend/
├── resumes/
├── generated/
├── jobs/
├── exports/
├── temp/
```

## Notes

- Update `backend/storage/sources.json` with company tokens for Greenhouse and Lever.
- Use the Resume Optimizer module to generate prompts for Claude, then paste responses back.
