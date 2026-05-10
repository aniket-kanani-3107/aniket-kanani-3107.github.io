import { NavLink } from 'react-router-dom';
import {
  LayoutDashboard,
  Radar,
  FileSearch,
  Sparkles,
  ClipboardList,
  MessageSquareText,
  Settings,
  Users,
} from 'lucide-react';

const links = [
  { to: '/', label: 'Dashboard', icon: LayoutDashboard },
  { to: '/scraper', label: 'Job Scraper', icon: Radar },
  { to: '/ats', label: 'ATS Analyzer', icon: FileSearch },
  { to: '/resume', label: 'Resume Optimizer', icon: Sparkles },
  { to: '/tracker', label: 'Application Tracker', icon: ClipboardList },
  { to: '/prompts', label: 'Prompt Studio', icon: MessageSquareText },
  { to: '/recruiters', label: 'Recruiter CRM', icon: Users },
  { to: '/settings', label: 'Settings', icon: Settings },
];

const Sidebar = () => (
  <aside className="sticky top-6 hidden h-[calc(100vh-48px)] w-64 flex-col gap-3 rounded-3xl border border-white/10 bg-white/5 p-6 text-slate-200 backdrop-blur-xl lg:flex">
    <div className="mb-4">
      <p className="text-xs uppercase tracking-widest text-slate-400">Offline AI Job Hunter</p>
      <h1 className="text-lg font-semibold text-white">Job Hunt HQ</h1>
    </div>
    <nav className="flex flex-1 flex-col gap-2">
      {links.map((link) => {
        const Icon = link.icon;
        return (
          <NavLink
            key={link.to}
            to={link.to}
            className={({ isActive }) =>
              `flex items-center gap-3 rounded-xl px-3 py-2 text-sm font-medium transition ${
                isActive ? 'bg-white/15 text-white' : 'text-slate-300 hover:bg-white/10'
              }`
            }
          >
            <Icon size={18} />
            {link.label}
          </NavLink>
        );
      })}
    </nav>
    <div className="rounded-xl border border-white/10 bg-white/5 p-3 text-xs text-slate-300">
      Local-first · Runs on localhost · Offline-ready
    </div>
  </aside>
);

export default Sidebar;
