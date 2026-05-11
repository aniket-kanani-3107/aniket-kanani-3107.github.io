import { cva } from 'class-variance-authority';
import { cn } from '../../lib/utils';

const buttonVariants = cva(
  'inline-flex items-center justify-center rounded-lg px-4 py-2 text-sm font-medium transition hover:opacity-90 focus:outline-none focus:ring-2 focus:ring-accent/40 disabled:opacity-50',
  {
    variants: {
      variant: {
        default: 'bg-accent text-white',
        secondary: 'bg-white/10 text-white',
        ghost: 'bg-transparent text-slate-200 hover:bg-white/10',
      },
    },
    defaultVariants: {
      variant: 'default',
    },
  }
);

const Button = ({ className, variant, ...props }) => (
  <button className={cn(buttonVariants({ variant }), className)} {...props} />
);

export { Button };
