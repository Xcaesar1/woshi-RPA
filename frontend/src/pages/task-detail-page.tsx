import { AppShell } from "@/components/layout/app-shell";
import { BentoGrid, BentoGridItem } from "@/components/ui/bento-grid";
import { PointerHighlight } from "@/components/ui/pointer-highlight";
import { PrimaryButton, SecondaryLink } from "@/components/ui/primary-button";
import { StatusPill } from "@/components/ui/status-pill";
import { fetchJson } from "@/lib/http";
import { readPageData } from "@/lib/page-data";
import { Activity, Boxes, Download, FileArchive, FileText, ListTodo, PackageCheck, UserRound } from "lucide-react";
import { StrictMode, startTransition, useEffect, useEffectEvent, useMemo, useState } from "react";
import { createRoot } from "react-dom/client";

type FbaResult = {
  fba_code: string;
  status: string;
  status_label: string;
  warehouse_count?: number;
  download_count: number;
  output_workbook?: string | null;
  report_file?: string | null;
  error?: string | null;
};

type TaskDetail = {
  id: string;
  original_filename: string;
  workflow_label: string;
  submitter: string;
  status: string;
  status_label: string;
  current_stage: string;
  total_fba_count: number;
  success_fba_count: number;
  failed_fba_count: number;
  created_at: string;
  started_at?: string | null;
  finished_at?: string | null;
  error_message?: string | null;
  recent_log: string;
  can_download: boolean;
  download_url: string;
  fba_results: FbaResult[];
};

type TaskDetailPayload = {
  task: TaskDetail;
};

function usePolling(callback: () => Promise<void>, enabled: boolean, intervalMs: number) {
  const onTick = useEffectEvent(callback);
  useEffect(() => {
    if (!enabled) {
      return;
    }
    const timer = window.setInterval(() => {
      void onTick();
    }, intervalMs);
    return () => {
      window.clearInterval(timer);
    };
  }, [enabled, intervalMs, onTick]);
}

function TaskDetailPage() {
  const payload = useMemo(() => readPageData<TaskDetailPayload>(), []);
  const [task, setTask] = useState(payload.task);
  const [hint, setHint] = useState("");

  const isTerminal = task.status === "SUCCESS" || task.status === "PARTIAL_SUCCESS" || task.status === "FAILED";

  const refresh = async () => {
    const response = await fetchJson<TaskDetail>(`/api/tasks/${task.id}`);
    startTransition(() => {
      setTask(response);
    });
  };

  usePolling(refresh, !isTerminal, 5000);

  async function manualRefresh() {
    setHint("正在获取最新执行状态。");
    try {
      await refresh();
      setHint("已更新。");
    } catch (error) {
      setHint(error instanceof Error ? error.message : "刷新失败");
    }
  }

  return (
    <AppShell
      eyebrow="任务详情"
      nav={[
        { href: "/tasks/new", label: "新建任务", icon: "new" },
        { href: "/tasks", label: "任务列表", current: true, icon: "list" },
      ]}
      title={
        <>
          这一页把任务状态、FBA 明细和最近日志，
          <br />
          都压缩到一个更容易排错的视图里。
        </>
      }
      subtitle={
        <>
          你不用再到多个目录里翻文件。当前任务的执行阶段、结果包是否可下载、每个 FBA 的下载数量和异常信息，都能直接在这里看到。
        </>
      }
      callout={
        <PointerHighlight
          rectangleClassName="rounded-full border-[color:oklch(0.73_0.03_188)]"
          pointerClassName="text-[color:oklch(0.53_0.08_188)]"
          containerClassName="max-w-fit rounded-full bg-white/70 px-4 py-2"
        >
          <span className="text-sm font-medium text-[color:oklch(0.35_0.03_230)]">
            当前任务编号：{task.id}
          </span>
        </PointerHighlight>
      }
      actions={
        <>
          <SecondaryLink href="/tasks">
            <ListTodo className="mr-2 h-4 w-4" />
            返回列表
          </SecondaryLink>
          <PrimaryButton type="button" onClick={() => void manualRefresh()}>
            <Activity className="mr-2 h-4 w-4" />
            立即刷新
          </PrimaryButton>
          {task.can_download ? (
            <a
              href={task.download_url}
              className="inline-flex min-h-12 items-center justify-center rounded-full bg-[linear-gradient(135deg,oklch(0.45_0.09_164),oklch(0.58_0.08_182))] px-5 text-sm font-semibold text-white shadow-[0_14px_32px_rgba(31,92,74,0.24)] transition hover:translate-y-[-1px]"
            >
              <Download className="mr-2 h-4 w-4" />
              下载结果
            </a>
          ) : null}
          <div className="text-sm text-[color:oklch(0.46_0.03_228)]">{hint}</div>
        </>
      }
      aside={
        <div className="rounded-[28px] border border-white/70 bg-white/78 p-5 shadow-[0_18px_70px_rgba(36,53,44,0.08)] backdrop-blur-xl">
          <p className="text-[0.72rem] font-semibold uppercase tracking-[0.24em] text-[color:oklch(0.55_0.03_205)]">
            当前状态
          </p>
          <div className="mt-4 space-y-4">
            <StatusPill status={task.status} label={task.status_label} className="text-sm" />
            <div className="text-sm leading-6 text-[color:oklch(0.42_0.03_228)]">
              <div>当前阶段：{task.current_stage || "-"}</div>
              <div>开始时间：{task.started_at || "-"}</div>
              <div>完成时间：{task.finished_at || "-"}</div>
            </div>
          </div>
        </div>
      }
    >
      <section className="space-y-6">
        <BentoGrid className="mx-0 max-w-none md:auto-rows-[14rem] md:grid-cols-4">
          <BentoGridItem
            className="border-[color:oklch(0.89_0.02_95)] bg-white/92 p-5 shadow-[0_18px_55px_rgba(41,59,49,0.08)]"
            icon={<FileText className="h-5 w-5 text-[color:oklch(0.52_0.08_190)]" />}
            title={task.original_filename}
            description="原始上传文件"
            header={<MetricHeader label="文件" />}
          />
          <BentoGridItem
            className="border-[color:oklch(0.89_0.02_95)] bg-white/92 p-5 shadow-[0_18px_55px_rgba(41,59,49,0.08)]"
            icon={<UserRound className="h-5 w-5 text-[color:oklch(0.55_0.08_40)]" />}
            title={task.submitter}
            description="提交人"
            header={<MetricHeader label="提交人" />}
          />
          <BentoGridItem
            className="border-[color:oklch(0.89_0.02_95)] bg-white/92 p-5 shadow-[0_18px_55px_rgba(41,59,49,0.08)]"
            icon={<PackageCheck className="h-5 w-5 text-[color:oklch(0.55_0.08_165)]" />}
            title={`${task.success_fba_count}/${task.total_fba_count}`}
            description="成功 FBA / 总 FBA"
            header={<MetricHeader label="完成度" />}
          />
          <BentoGridItem
            className="border-[color:oklch(0.89_0.02_95)] bg-white/92 p-5 shadow-[0_18px_55px_rgba(41,59,49,0.08)]"
            icon={<FileArchive className="h-5 w-5 text-[color:oklch(0.58_0.11_23)]" />}
            title={task.can_download ? "可下载" : "处理中"}
            description={task.workflow_label}
            header={<MetricHeader label="结果包" />}
          />
        </BentoGrid>

        <section className="overflow-hidden rounded-[30px] border border-white/70 bg-white/84 shadow-[0_20px_80px_rgba(36,56,43,0.08)] backdrop-blur-xl">
          <div className="border-b border-[color:oklch(0.92_0.01_95)] px-6 py-5">
            <h2 className="font-[family-name:var(--font-display)] text-2xl font-semibold tracking-[-0.03em] text-[color:oklch(0.22_0.025_242)]">
              FBA 执行明细
            </h2>
            <p className="mt-2 text-sm leading-6 text-[color:oklch(0.46_0.03_228)]">
              每个 FBA 的下载文件数量、处理结果和错误信息都在这里。
            </p>
          </div>

          <div className="overflow-x-auto">
            <table className="min-w-full border-collapse">
              <thead>
                <tr className="bg-[color:oklch(0.985_0.003_95)] text-left text-xs uppercase tracking-[0.14em] text-[color:oklch(0.55_0.02_228)]">
                  <th className="px-6 py-4 font-semibold">FBA</th>
                  <th className="px-6 py-4 font-semibold">状态</th>
                  <th className="px-6 py-4 font-semibold">下载数</th>
                  <th className="px-6 py-4 font-semibold">输出文件</th>
                  <th className="px-6 py-4 font-semibold">错误</th>
                </tr>
              </thead>
              <tbody>
                {task.fba_results.length === 0 ? (
                  <tr>
                    <td colSpan={5} className="px-6 py-12 text-center text-sm text-[color:oklch(0.48_0.03_228)]">
                      任务还没有产出 FBA 明细。
                    </td>
                  </tr>
                ) : (
                  task.fba_results.map((item) => (
                    <tr key={`${item.fba_code}-${item.status}`} className="border-t border-[color:oklch(0.93_0.008_95)] align-top">
                      <td className="px-6 py-5 text-sm font-medium text-[color:oklch(0.26_0.02_232)]">{item.fba_code}</td>
                      <td className="px-6 py-5">
                        <StatusPill status={item.status} label={item.status_label || item.status} />
                      </td>
                      <td className="px-6 py-5 text-sm text-[color:oklch(0.33_0.02_232)]">{item.download_count ?? 0}</td>
                      <td className="px-6 py-5 text-sm text-[color:oklch(0.46_0.03_228)]">{item.output_workbook || "-"}</td>
                      <td className="px-6 py-5 text-sm leading-6 text-[color:oklch(0.46_0.03_228)]">{item.error || "-"}</td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </section>

        <section className="grid gap-6 lg:grid-cols-[minmax(0,0.88fr)_minmax(0,1.12fr)]">
          <article className="rounded-[30px] border border-white/70 bg-white/84 p-6 shadow-[0_20px_80px_rgba(36,56,43,0.08)] backdrop-blur-xl">
            <div className="flex items-center gap-3">
              <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-[linear-gradient(135deg,rgba(74,138,130,0.12),rgba(106,163,158,0.22))] text-[color:oklch(0.43_0.08_182)]">
                <Boxes className="h-5 w-5" />
              </div>
              <div>
                <h2 className="font-[family-name:var(--font-display)] text-2xl font-semibold tracking-[-0.03em] text-[color:oklch(0.22_0.025_242)]">
                  错误信息
                </h2>
                <p className="mt-1 text-sm text-[color:oklch(0.46_0.03_228)]">如果任务失败，这里会直接给出总错误。</p>
              </div>
            </div>
            <pre className="mt-5 overflow-x-auto rounded-[24px] bg-[color:oklch(0.97_0.01_95)] p-5 text-sm leading-7 text-[color:oklch(0.33_0.02_232)] whitespace-pre-wrap">
              {task.error_message || "无"}
            </pre>
          </article>

          <article className="rounded-[30px] border border-white/70 bg-[linear-gradient(180deg,rgba(28,39,43,0.98),rgba(36,46,52,0.96))] p-6 shadow-[0_24px_90px_rgba(22,31,36,0.2)]">
            <div className="flex items-center justify-between gap-3">
              <div>
                <h2 className="font-[family-name:var(--font-display)] text-2xl font-semibold tracking-[-0.03em] text-[color:oklch(0.95_0.006_95)]">
                  最近日志
                </h2>
                <p className="mt-1 text-sm text-[color:oklch(0.8_0.02_230)]">实时轮询时会自动刷新这一块。</p>
              </div>
            </div>
            <pre className="mt-5 max-h-[420px] overflow-auto whitespace-pre-wrap rounded-[24px] border border-white/10 bg-black/15 p-5 text-sm leading-7 text-[color:oklch(0.9_0.01_230)]">
              {task.recent_log || "暂无日志"}
            </pre>
          </article>
        </section>
      </section>
    </AppShell>
  );
}

function MetricHeader({ label }: { label: string }) {
  return (
    <div className="rounded-2xl bg-[linear-gradient(135deg,rgba(255,255,255,0.94),rgba(245,242,235,0.88))] p-3 text-xs font-medium uppercase tracking-[0.18em] text-[color:oklch(0.52_0.03_205)]">
      {label}
    </div>
  );
}

const root = document.getElementById("root");
if (root) {
  createRoot(root).render(
    <StrictMode>
      <TaskDetailPage />
    </StrictMode>,
  );
}
