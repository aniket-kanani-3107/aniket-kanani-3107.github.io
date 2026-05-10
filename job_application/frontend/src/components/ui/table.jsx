import { cn } from '../../lib/utils';

const Table = ({ className, ...props }) => (
  <div className={cn('w-full overflow-auto', className)}>
    <table className="w-full text-sm text-slate-200" {...props} />
  </div>
);

const TableHead = ({ className, ...props }) => (
  <th className={cn('px-3 py-2 text-left text-xs uppercase tracking-wide text-slate-400', className)} {...props} />
);

const TableCell = ({ className, ...props }) => (
  <td className={cn('px-3 py-3 align-top text-sm text-slate-200', className)} {...props} />
);

const TableRow = ({ className, ...props }) => (
  <tr className={cn('border-b border-white/5', className)} {...props} />
);

export { Table, TableHead, TableCell, TableRow };
