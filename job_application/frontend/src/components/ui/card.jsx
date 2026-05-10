import { cn } from '../../lib/utils';

const Card = ({ className, ...props }) => (
  <div className={cn('glass-card p-5', className)} {...props} />
);

const CardHeader = ({ className, ...props }) => (
  <div className={cn('flex items-center justify-between', className)} {...props} />
);

const CardTitle = ({ className, ...props }) => (
  <h3 className={cn('text-lg font-semibold text-white', className)} {...props} />
);

const CardContent = ({ className, ...props }) => (
  <div className={cn('mt-4 text-sm text-slate-200', className)} {...props} />
);

export { Card, CardHeader, CardTitle, CardContent };
