import { AppShell } from "@/components/layout/app-shell";
import { BentoGrid, BentoGridItem } from "@/components/ui/bento-grid";
import { PointerHighlight } from "@/components/ui/pointer-highlight";
import { PrimaryButton, SecondaryLink } from "@/components/ui/primary-button";
import { StatusPill } from "@/components/ui/status-pill";
import { fetchJson } from "@/lib/http";
import { readPageData } from "@/lib/page-data";
import { Activity, Download, Eye, Filter, HardDriveDownload, LayoutList, PlusCircle, Timer } from "lucide-react";
import {
  StrictMode,
  startTransition,
  useEffect,
  useEffectEvent,
  useMemo,
  useState,
} from "react";
import { createRoot } from "react-dom/client";

type TaskView = {
  id: string;
  original_filename: string;
  workflow_label: string;
  submitter: string;
  created_at: string;
  status: string;
  status_label: string;
  success_fba_count: number;
  failed_fba_count: number;
  can_download: boolean;
  detail_url: string;
  download_url: string;
};

type SystemStatus = {
  queued_count: number;
  running_count: number;
  queue_depth: number;
  worker_alive: boolean;
  worker_recent_heartbeat: string | null;
  browser_slots_total: number;
  browser_slots_in_use: number;
  redis_error?: string | null;
};

type TaskListPayload = {
  tasks: TaskView[];
  system_status: SystemStatus;
  submitter: string;
  status: string;
  status_choices: [string, string][];
};

function usePolling(callback: () => Promise<void>, intervalMs: number) {
  const onTick = useEffectEvent(callback);
  useEffect(() => {
    const timer = window.setInterval(() => {
      void onTick();
    }, intervalMs);
    return () => {
      window.clearInterval(timer);
    };
  }, [intervalMs, onTick]);
}

function TaskListPage() {
  const payload = useMemo(() => readPageData<TaskListPayload>(), []);
  const [tasks, setTasks] = useState(payload.tasks);
  const [systemStatus, setSystemStatus] = useState(payload.system_status);
  const [submitter, setSubmitter] = useState(payload.submitter);
  const [status, setStatus] = useState(payload.status);
  const [loading, setLoading] = useState(false);
  const [hint, setHint] = useState("");

  const refresh = async () => {
    const params = new URLSearchParams();
    const effectiveSubmitter = submitter.trim();
    if (effectiveSubmitter) {
      params.set("submitter", effectiveSubmitter);
    }
    if (status) {
      params.set("status", status);
    }
    const query = params.toString();
    const [taskPayload, statusPayload] = await Promise.all([
      fetchJson<{ tasks: TaskView[] }>(`/api/tasks${query ? `?${query}` : ""}`),
      fetchJson<SystemStatus>("/api/system/status"),
    ]);
    startTransition(() => {
      setTasks(taskPayload.tasks ?? []);
      setSystemStatus(statusPayload);
      window.history.replaceState({}, "", `/tasks${query ? `?${query}` : ""}`);
    });
  };

  useEffect(() => {
    const remembered = window.localStorage.getItem("lingxing_submitter");
    if (!payload.submitter && remembered) {
      setSubmitter(remembered);
      void (async () => {
        await fetchWithState(remembered, payload.status);
      })();
    }
  }, [payload.status, payload.submitter]);

  async function fetchWithState(nextSubmitter = submitter, nextStatus = status) {
    setLoading(true);
    setHint("正在刷新任务列表与系统状态。");
    try {
      const params = new URLSearchParams();
      const effectiveSubmitter = nextSubmitter.trim();
      if (effectiveSubmitter) {
        params.set("submitter", effectiveSubmitter);
      }
      if (nextStatus) {
        params.set("status", nextStatus);
      }
      const query = params.toString();
      const [taskPayload, statusPayload] = await Promise.all([
        fetchJson<{ tasks: TaskView[] }>(`/api/tasks${query ? `?${query}` : ""}`),
        fetchJson<SystemStatus>("/api/system/status"),
      ]);
      startTransition(() => {
        setTasks(taskPayload.tasks ?? []);
        setSystemStatus(statusPayload);
        window.history.replaceState({}, "", `/tasks${query ? `?${query}` : ""}`);
      });
      if (effectiveSubmitter) {
        window.localStorage.setItem("lingxing_submitter", effectiveSubmitter);
      }
      setHint("已同步最新状态。");
    } catch (error) {
      setHint(error instanceof Error ? error.message : "刷新失败");
    } finally {
      setLoading(false);
    }
  }

  usePolling(refresh, 5000);

  return (
    <AppShell
        eyebrow="任务列表"
        nav={[
          { href: "/tasks/new", label: "新建任务", icon: "new" },
          { href: "/tasks", label: "任务列表", current: true, icon: "list" },
        ]}
        title={
          <>
            把排队、执行和下载状态，
            <br />
            放到一个更容易扫读的工作台里。
          </>
        }
        subtitle={
          <>
            页面会自动刷新，不需要反复手动点按钮。你能一眼看清当前有多少任务在排队、浏览器执行槽是否被占用、哪条任务已经可以下载结果。
          </>
        }
        callout={
          <PointerHighlight
            rectangleClassName="rounded-full border-[color:oklch(0.73_0.03_188)]"
            pointerClassName="text-[color:oklch(0.53_0.08_188)]"
            containerClassName="max-w-fit rounded-full bg-white/70 px-4 py-2"
          >
            <span className="text-sm font-medium text-[color:oklch(0.35_0.03_230)]">
              就算 10 台电脑同时提交，这里也只会排队，不会同时拉起 10 个浏览器。
            </span>
          </PointerHighlight>
        }
        actions={
          <>
            <SecondaryLink href="/tasks/new">
              <PlusCircle className="mr-2 h-4 w-4" />
              新建任务
            </SecondaryLink>
            <PrimaryButton type="button" onClick={() => void fetchWithState()} disabled={loading}>
              <Activity className="mr-2 h-4 w-4" />
              {loading ? "刷新中" : "立即刷新"}
            </PrimaryButton>
            <div className="text-sm text-[color:oklch(0.46_0.03_228)]">{hint}</div>
          </>
        }
        aside={
          <div className="rounded-[28px] border border-white/70 bg-white/78 p-5 shadow-[0_18px_70px_rgba(36,53,44,0.08)] backdrop-blur-xl">
            <p className="text-[0.72rem] font-semibold uppercase tracking-[0.24em] text-[color:oklch(0.55_0.03_205)]">
              观察建议
            </p>
            <div className="mt-4 space-y-3 text-sm leading-6 text-[color:oklch(0.42_0.03_228)]">
              <p>如果 `worker 在线` 正常、`执行槽` 长时间为 0，但任务没有启动，优先看 Redis 或 worker 日志。</p>
              <p>部分成功任务也能下载结果，失败明细会一并打包，方便继续排错。</p>
            </div>
          </div>
        }
      >
        <section className="space-y-6">
          <BentoGrid className="mx-0 max-w-none md:auto-rows-[14rem] md:grid-cols-4">
            <BentoGridItem
              className="border-[color:oklch(0.89_0.02_95)] bg-white/92 p-5 shadow-[0_18px_55px_rgba(41,59,49,0.08)]"
              icon={<LayoutList className="h-5 w-5 text-[color:oklch(0.52_0.08_190)]" />}
              title={`${systemStatus.queued_count ?? 0}`}
              description="排队中的任务"
              header={<MetricHeader label="队列" />}
            />
            <BentoGridItem
              className="border-[color:oklch(0.89_0.02_95)] bg-white/92 p-5 shadow-[0_18px_55px_rgba(41,59,49,0.08)]"
              icon={<Timer className="h-5 w-5 text-[color:oklch(0.56_0.1_78)]" />}
              title={`${systemStatus.running_count ?? 0}`}
              description="当前运行中的任务"
              header={<MetricHeader label="执行" />}
            />
            <BentoGridItem
              className="border-[color:oklch(0.89_0.02_95)] bg-white/92 p-5 shadow-[0_18px_55px_rgba(41,59,49,0.08)]"
              icon={<HardDriveDownload className="h-5 w-5 text-[color:oklch(0.55_0.08_165)]" />}
              title={`${systemStatus.browser_slots_in_use ?? 0}/${systemStatus.browser_slots_total ?? 1}`}
              description="浏览器执行槽占用"
              header={<MetricHeader label="浏览器" />}
            />
            <BentoGridItem
              className="border-[color:oklch(0.89_0.02_95)] bg-white/92 p-5 shadow-[0_18px_55px_rgba(41,59,49,0.08)]"
              icon={<Activity className="h-5 w-5 text-[color:oklch(0.58_0.11_23)]" />}
              title={systemStatus.worker_alive ? "在线" : "离线"}
              description={systemStatus.worker_recent_heartbeat || "暂无心跳"}
              header={<MetricHeader label="Worker" />}
            />
          </BentoGrid>

          <section className="rounded-[30px] border border-white/70 bg-white/82 p-6 shadow-[0_20px_80px_rgba(36,56,43,0.08)] backdrop-blur-xl">
            <div className="flex flex-col gap-4 lg:flex-row lg:items-end lg:justify-between">
              <div>
                <p className="text-[0.72rem] font-semibold uppercase tracking-[0.24em] text-[color:oklch(0.55_0.03_205)]">
                  任务筛选
                </p>
                <h2 className="mt-2 font-[family-name:var(--font-display)] text-2xl font-semibold tracking-[-0.03em] text-[color:oklch(0.22_0.025_242)]">
                  查看“我的任务”或按状态筛选
                </h2>
              </div>
              <div className="text-sm text-[color:oklch(0.46_0.03_228)]">任务列表每 5 秒自动刷新一次。</div>
            </div>

            <div className="mt-6 grid gap-4 md:grid-cols-[minmax(0,1fr)_220px_auto_auto]">
              <label className="space-y-2">
                <span className="text-sm font-medium text-[color:oklch(0.32_0.02_232)]">提交人</span>
                <input
                  value={submitter}
                  onChange={(event) => setSubmitter(event.target.value)}
                  placeholder="输入姓名或工号"
                  className="min-h-12 w-full rounded-2xl border border-[color:oklch(0.88_0.01_95)] bg-[color:oklch(0.99_0.002_95)] px-4 text-sm text-[color:oklch(0.28_0.02_232)] outline-none transition placeholder:text-[color:oklch(0.68_0.02_228)] focus:border-[color:oklch(0.6_0.09_190)] focus:ring-2 focus:ring-[color:oklch(0.86_0.03_190)]"
                />
              </label>
              <label className="space-y-2">
                <span className="text-sm font-medium text-[color:oklch(0.32_0.02_232)]">任务状态</span>
                <select
                  value={status}
                  onChange={(event) => setStatus(event.target.value)}
                  className="min-h-12 w-full rounded-2xl border border-[color:oklch(0.88_0.01_95)] bg-[color:oklch(0.99_0.002_95)] px-4 text-sm text-[color:oklch(0.28_0.02_232)] outline-none transition focus:border-[color:oklch(0.6_0.09_190)] focus:ring-2 focus:ring-[color:oklch(0.86_0.03_190)]"
                >
                  <option value="">全部状态</option>
                  {payload.status_choices.map(([value, label]) => (
                    <option key={value} value={value}>
                      {label}
                    </option>
                  ))}
                </select>
              </label>
              <div className="flex items-end">
                <PrimaryButton
                  type="button"
                  onClick={() => void fetchWithState()}
                  className="w-full"
                  disabled={loading}
                >
                  <Filter className="mr-2 h-4 w-4" />
                  应用筛选
                </PrimaryButton>
              </div>
              <div className="flex items-end">
                <button
                  type="button"
                  onClick={() => {
                    setSubmitter("");
                    setStatus("");
                    void fetchWithState("", "");
                  }}
                  className="inline-flex min-h-12 w-full items-center justify-center rounded-full border border-[color:oklch(0.88_0.01_95)] bg-white/88 px-5 text-sm font-semibold text-[color:oklch(0.34_0.02_232)] transition hover:bg-[color:oklch(0.975_0.004_95)]"
                >
                  清空
                </button>
              </div>
            </div>
          </section>

          <section className="overflow-hidden rounded-[30px] border border-white/70 bg-white/84 shadow-[0_20px_80px_rgba(36,56,43,0.08)] backdrop-blur-xl">
            <div className="border-b border-[color:oklch(0.92_0.01_95)] px-6 py-5">
              <h2 className="font-[family-name:var(--font-display)] text-2xl font-semibold tracking-[-0.03em] text-[color:oklch(0.22_0.025_242)]">
                最近任务
              </h2>
              <p className="mt-2 text-sm leading-6 text-[color:oklch(0.46_0.03_228)]">
                结果可下载的任务会直接显示下载入口；还在排队的任务则保留详情入口。
              </p>
            </div>

            <div className="overflow-x-auto">
              <table className="min-w-full border-collapse">
                <thead>
                  <tr className="bg-[color:oklch(0.985_0.003_95)] text-left text-xs uppercase tracking-[0.14em] text-[color:oklch(0.55_0.02_228)]">
                    <th className="px-6 py-4 font-semibold">任务编号</th>
                    <th className="px-6 py-4 font-semibold">文件</th>
                    <th className="px-6 py-4 font-semibold">流程</th>
                    <th className="px-6 py-4 font-semibold">提交人</th>
                    <th className="px-6 py-4 font-semibold">提交时间</th>
                    <th className="px-6 py-4 font-semibold">状态</th>
                    <th className="px-6 py-4 font-semibold">成功 / 失败</th>
                    <th className="px-6 py-4 font-semibold">操作</th>
                  </tr>
                </thead>
                <tbody>
                  {tasks.length === 0 ? (
                    <tr>
                      <td colSpan={8} className="px-6 py-12 text-center text-sm text-[color:oklch(0.48_0.03_228)]">
                        当前筛选条件下还没有任务。
                      </td>
                    </tr>
                  ) : (
                    tasks.map((task) => (
                      <tr key={task.id} className="border-t border-[color:oklch(0.93_0.008_95)] align-top">
                        <td className="px-6 py-5 text-sm font-medium text-[color:oklch(0.26_0.02_232)]">{task.id}</td>
                        <td className="px-6 py-5 text-sm text-[color:oklch(0.33_0.02_232)]">{task.original_filename}</td>
                        <td className="px-6 py-5 text-sm text-[color:oklch(0.46_0.03_228)]">{task.workflow_label}</td>
                        <td className="px-6 py-5 text-sm text-[color:oklch(0.33_0.02_232)]">{task.submitter}</td>
                        <td className="px-6 py-5 text-sm text-[color:oklch(0.46_0.03_228)]">{task.created_at || "-"}</td>
                        <td className="px-6 py-5">
                          <StatusPill status={task.status} label={task.status_label} />
                        </td>
                        <td className="px-6 py-5 text-sm text-[color:oklch(0.33_0.02_232)]">
                          {task.success_fba_count}/{task.failed_fba_count}
                        </td>
                        <td className="px-6 py-5">
                          <div className="flex flex-wrap gap-2">
                            <a
                              href={task.detail_url}
                              className="inline-flex items-center rounded-full border border-[color:oklch(0.88_0.01_95)] bg-white px-3 py-2 text-sm font-medium text-[color:oklch(0.33_0.02_232)] transition hover:bg-[color:oklch(0.975_0.004_95)]"
                            >
                              <Eye className="mr-2 h-4 w-4" />
                              详情
                            </a>
                            {task.can_download ? (
                              <a
                                href={task.download_url}
                                className="inline-flex items-center rounded-full bg-[color:oklch(0.955_0.03_165)] px-3 py-2 text-sm font-medium text-[color:oklch(0.3_0.05_170)] transition hover:bg-[color:oklch(0.93_0.04_165)]"
                              >
                                <Download className="mr-2 h-4 w-4" />
                                下载
                              </a>
                            ) : (
                              <span className="inline-flex items-center rounded-full bg-[color:oklch(0.97_0.01_95)] px-3 py-2 text-sm text-[color:oklch(0.48_0.02_228)]">
                                等待结果
                              </span>
                            )}
                          </div>
                        </td>
                      </tr>
                    ))
                  )}
                </tbody>
              </table>
            </div>
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
      <TaskListPage />
    </StrictMode>,
  );
}
