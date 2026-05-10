const PageShell = ({ title, subtitle, children, action }) => (
  <section className="glass-card p-6">
    <div className="flex flex-wrap items-start justify-between gap-4">
      <div>
        <h3 className="section-title">{title}</h3>
        {subtitle && <p className="mt-1 text-sm text-slate-400">{subtitle}</p>}
      </div>
      {action}
    </div>
    <div className="mt-6">{children}</div>
  </section>
);

export default PageShell;
