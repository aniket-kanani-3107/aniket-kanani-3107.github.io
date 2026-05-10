import { cn } from '../../lib/utils';

const Badge = ({ className, ...props }) => (
  <span
    className={cn(
      'inline-flex items-center rounded-full bg-white/10 px-3 py-1 text-xs font-medium text-slate-200',
      className
    )}
    {...props}
  />
);

export { Badge };
