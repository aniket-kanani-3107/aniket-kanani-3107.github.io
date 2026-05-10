import { cn } from '../../lib/utils';

const Input = ({ className, ...props }) => (
  <input className={cn('glass-input w-full', className)} {...props} />
);

export { Input };
