import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { createBrowserRouter, RouterProvider } from 'react-router-dom';
import './index.css';
import App from './App.jsx';
import Dashboard from './pages/Dashboard.jsx';
import Scraper from './pages/Scraper.jsx';
import AtsAnalyzer from './pages/AtsAnalyzer.jsx';
import ResumeOptimizer from './pages/ResumeOptimizer.jsx';
import Tracker from './pages/Tracker.jsx';
import PromptStudio from './pages/PromptStudio.jsx';
import Settings from './pages/Settings.jsx';
import Recruiters from './pages/Recruiters.jsx';

const router = createBrowserRouter([
  {
    path: '/',
    element: <App />,
    children: [
      { index: true, element: <Dashboard /> },
      { path: 'scraper', element: <Scraper /> },
      { path: 'ats', element: <AtsAnalyzer /> },
      { path: 'resume', element: <ResumeOptimizer /> },
      { path: 'tracker', element: <Tracker /> },
      { path: 'prompts', element: <PromptStudio /> },
      { path: 'recruiters', element: <Recruiters /> },
      { path: 'settings', element: <Settings /> },
    ],
  },
]);

createRoot(document.getElementById('root')).render(
  <StrictMode>
    <RouterProvider router={router} />
  </StrictMode>
);
