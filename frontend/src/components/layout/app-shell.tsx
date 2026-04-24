import { Boxes, Clock3, ListTodo, PlusCircle } from "lucide-react";
import type { ReactNode } from "react";

import { Spotlight } from "@/components/ui/spotlight-new";
import { cn } from "@/lib/utils";

type NavItem = {
  href: string;
  label: string;
  current?: boolean;
  icon?: "new" | "list";
};

type AppShellProps = {
  eyebrow: string;
  title: ReactNode;
  subtitle: ReactNode;
  callout?: ReactNode;
  nav: NavItem[];
  actions?: ReactNode;
  children: ReactNode;
  aside?: ReactNode;
  pageClassName?: string;
};

const navIcons = {
  new: PlusCircle,
  list: ListTodo,
} as const;

export function AppShell({
  eyebrow,
  title,
  subtitle,
  callout,
  nav,
  actions,
  children,
  aside,
  pageClassName,
}: AppShellProps) {
  return (
    <div className="relative min-h-screen overflow-hidden">
      <Spotlight
        width={520}
        height={1200}
        smallWidth={220}
        duration={10}
        xOffset={64}
        gradientFirst="radial-gradient(68.54% 68.72% at 55.02% 31.46%, hsla(174, 54%, 62%, .16) 0, hsla(174, 54%, 42%, .05) 50%, transparent 80%)"
        gradientSecond="radial-gradient(50% 50% at 50% 50%, hsla(36, 82%, 70%, .10) 0, hsla(36, 82%, 58%, .03) 80%, transparent 100%)"
        gradientThird="radial-gradient(50% 50% at 50% 50%, hsla(201, 72%, 78%, .08) 0, hsla(201, 72%, 55%, .02) 80%, transparent 100%)"
      />
      <div className="relative z-10 mx-auto flex min-h-screen w-full max-w-7xl flex-col px-4 py-5 sm:px-6 lg:px-8">
        <header className="rounded-[28px] border border-white/70 bg-white/72 px-5 py-4 shadow-[0_24px_80px_rgba(42,61,51,0.08)] backdrop-blur-xl">
          <div className="flex flex-col gap-4 lg:flex-row lg:items-center lg:justify-between">
            <div className="flex items-start gap-3">
              <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-[linear-gradient(135deg,oklch(0.44_0.08_182),oklch(0.64_0.08_196))] text-white shadow-[0_14px_30px_rgba(49,101,97,0.22)]">
                <Boxes className="h-5 w-5" />
              </div>
              <div className="space-y-1">
                <p className="text-[0.7rem] font-semibold uppercase tracking-[0.28em] text-[color:oklch(0.51_0.04_205)]">
                  领星自动化任务中心
                </p>
                <div className="flex flex-wrap items-center gap-2 text-sm text-[color:oklch(0.39_0.03_232)]">
                  <span>上传清单</span>
                  <span className="text-[color:oklch(0.72_0.02_210)]">/</span>
                  <span>后台排队</span>
                  <span className="text-[color:oklch(0.72_0.02_210)]">/</span>
                  <span>自动下载与整理</span>
                </div>
              </div>
            </div>
            <div className="flex flex-wrap items-center gap-3">
              <div className="inline-flex items-center gap-2 rounded-full border border-[color:oklch(0.88_0.015_95)] bg-[color:oklch(0.985_0.005_95)] px-3 py-2 text-xs font-medium text-[color:oklch(0.39_0.03_232)]">
                <Clock3 className="h-3.5 w-3.5" />
                多人提交，后台按顺序稳稳处理
              </div>
              <nav className="flex flex-wrap items-center gap-2 rounded-full border border-[color:oklch(0.9_0.01_95)] bg-white/72 p-1 shadow-[inset_0_1px_0_rgba(255,255,255,0.7)]">
                {nav.map((item) => {
                  const Icon = item.icon ? navIcons[item.icon] : null;
                  return (
                    <a
                      key={item.href}
                      aria-current={item.current ? "page" : undefined}
                      href={item.href}
                      className={cn(
                        "inline-flex items-center gap-2 rounded-full px-4 py-2.5 text-sm font-semibold transition focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-[color:oklch(0.65_0.08_182)]",
                        item.current
                          ? "border border-[color:oklch(0.78_0.05_182)] bg-[linear-gradient(180deg,oklch(0.965_0.035_176),oklch(0.93_0.04_188))] text-[color:oklch(0.26_0.055_196)] shadow-[0_10px_24px_rgba(42,106,99,0.14)]"
                          : "text-[color:oklch(0.42_0.025_232)] hover:bg-[color:oklch(0.97_0.01_95)] hover:text-[color:oklch(0.26_0.04_196)]",
                      )}
                    >
                      {Icon ? <Icon className="h-4 w-4" /> : null}
                      {item.label}
                    </a>
                  );
                })}
              </nav>
            </div>
          </div>
        </header>

        <main className={cn("mt-6 flex-1 space-y-6", pageClassName)}>
          <section className="grid gap-5 lg:grid-cols-[minmax(0,1.2fr)_minmax(280px,0.8fr)]">
            <div className="relative overflow-hidden rounded-[32px] border border-white/70 bg-[linear-gradient(145deg,rgba(255,255,255,0.92),rgba(250,247,241,0.8))] px-6 py-7 shadow-[0_28px_100px_rgba(34,54,44,0.10)]">
              <div className="absolute inset-x-6 top-0 h-px bg-[linear-gradient(90deg,transparent,rgba(76,115,111,0.4),transparent)]" />
              <div className="space-y-4">
                <span className="inline-flex w-fit rounded-full border border-[color:oklch(0.87_0.02_95)] bg-white/80 px-3 py-1 text-[0.68rem] font-semibold uppercase tracking-[0.24em] text-[color:oklch(0.51_0.04_205)]">
                  {eyebrow}
                </span>
                <div className="space-y-3">
                  <h1 className="max-w-4xl font-[family-name:var(--font-display)] text-[clamp(2rem,3vw,3.75rem)] leading-[1.02] font-semibold tracking-[-0.04em] text-[color:oklch(0.22_0.025_242)]">
                    {title}
                  </h1>
                  <div className="max-w-3xl text-[0.98rem] leading-7 text-[color:oklch(0.42_0.03_228)]">
                    {subtitle}
                  </div>
                  {callout ? <div className="pt-2">{callout}</div> : null}
                </div>
              </div>
            </div>

            <aside className="space-y-4">
              <div className="rounded-[28px] border border-white/70 bg-white/78 p-5 shadow-[0_18px_70px_rgba(36,53,44,0.08)] backdrop-blur-xl">
                <p className="text-[0.72rem] font-semibold uppercase tracking-[0.24em] text-[color:oklch(0.55_0.03_205)]">
                  小白上手指南
                </p>
                <div className="mt-4 space-y-3">
                  <div className="rounded-2xl bg-[color:oklch(0.968_0.014_92)] px-4 py-3">
                    <div className="text-sm font-semibold text-[color:oklch(0.26_0.03_232)]">1. 准备清单</div>
                    <div className="mt-1 text-sm leading-6 text-[color:oklch(0.36_0.035_228)]">
                      不会做表也没关系，先下载示例文件，把 FBA 号替换进去。
                    </div>
                  </div>
                  <div className="rounded-2xl bg-[color:oklch(0.968_0.014_92)] px-4 py-3">
                    <div className="text-sm font-semibold text-[color:oklch(0.26_0.03_232)]">2. 上传提交</div>
                    <div className="mt-1 text-sm leading-6 text-[color:oklch(0.36_0.035_228)]">
                      选择 `.txt` 或 `.xlsx`，填写提交人，然后点“开始处理”。
                    </div>
                  </div>
                  <div className="rounded-2xl bg-[color:oklch(0.968_0.014_92)] px-4 py-3">
                    <div className="text-sm font-semibold text-[color:oklch(0.26_0.03_232)]">3. 等待下载</div>
                    <div className="mt-1 text-sm leading-6 text-[color:oklch(0.36_0.035_228)]">
                      去“任务列表”看进度，显示可下载后直接拿结果包。
                    </div>
                  </div>
                </div>
              </div>

              {aside}
            </aside>
          </section>

          {actions ? <div className="flex flex-wrap items-center gap-3">{actions}</div> : null}
          {children}
        </main>
      </div>
    </div>
  );
}
