import { cn } from '../../lib/utils';

const Textarea = ({ className, ...props }) => (
  <textarea className={cn('glass-input w-full min-h-[140px] resize-y', className)} {...props} />
);

export { Textarea };
