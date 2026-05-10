import { Bell, Search } from 'lucide-react';

const Topbar = () => (
  <header className="flex flex-wrap items-center justify-between gap-4 rounded-2xl border border-white/10 bg-white/5 px-6 py-4 backdrop-blur-xl">
    <div>
      <p className="text-xs uppercase tracking-widest text-slate-400">Job Hunt Command Center</p>
      <h2 className="text-2xl font-semibold text-white">Welcome back, Aniket</h2>
    </div>
    <div className="flex items-center gap-3">
      <div className="hidden items-center gap-2 rounded-full border border-white/10 bg-white/5 px-3 py-2 text-sm text-slate-300 md:flex">
        <Search size={16} />
        Quick search jobs or companies
      </div>
      <button className="rounded-full border border-white/10 bg-white/5 p-2 text-slate-200">
        <Bell size={18} />
      </button>
    </div>
  </header>
);

export default Topbar;
