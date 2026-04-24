import type { ButtonHTMLAttributes, ReactNode } from "react";

import { cn } from "@/lib/utils";

export function PrimaryButton({
  children,
  className,
  disabled,
  ...props
}: ButtonHTMLAttributes<HTMLButtonElement> & { children: ReactNode }) {
  return (
    <button
      {...props}
      disabled={disabled}
      className={cn(
        "inline-flex min-h-12 items-center justify-center rounded-full px-5 text-sm font-semibold tracking-[0.01em] transition",
        disabled
          ? "cursor-not-allowed bg-[color:oklch(0.89_0.01_95)] text-[color:oklch(0.58_0.02_230)]"
          : "bg-[linear-gradient(135deg,oklch(0.42_0.08_182),oklch(0.56_0.08_198))] text-white shadow-[0_14px_32px_rgba(31,92,93,0.25)] hover:translate-y-[-1px]",
        className,
      )}
    >
      {children}
    </button>
  );
}

export function SecondaryLink({
  href,
  children,
  className,
}: {
  href: string;
  children: ReactNode;
  className?: string;
}) {
  return (
    <a
      href={href}
      className={cn(
        "inline-flex min-h-12 items-center justify-center rounded-full border border-[color:oklch(0.88_0.01_95)] bg-white/88 px-5 text-sm font-semibold text-[color:oklch(0.34_0.02_232)] transition hover:bg-[color:oklch(0.975_0.004_95)]",
        className,
      )}
    >
      {children}
    </a>
  );
}
