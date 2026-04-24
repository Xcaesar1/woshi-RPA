import { cn } from "@/lib/utils";

const STATUS_THEME: Record<string, string> = {
  PENDING: "bg-stone-200/80 text-stone-700",
  QUEUED: "bg-amber-100 text-amber-800",
  RUNNING: "bg-cyan-100 text-cyan-800",
  SUCCESS: "bg-emerald-100 text-emerald-800",
  PARTIAL_SUCCESS: "bg-orange-100 text-orange-800",
  FAILED: "bg-rose-100 text-rose-800",
};

export function StatusPill({
  status,
  label,
  className,
}: {
  status: string;
  label: string;
  className?: string;
}) {
  return (
    <span
      className={cn(
        "inline-flex min-h-8 items-center rounded-full px-3 py-1 text-xs font-semibold tracking-[0.02em]",
        STATUS_THEME[status] ?? "bg-stone-200 text-stone-700",
        className,
      )}
    >
      {label}
    </span>
  );
}
