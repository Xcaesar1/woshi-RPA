import { BentoGrid, BentoGridItem } from "@/components/ui/bento-grid";
import { PointerHighlight } from "@/components/ui/pointer-highlight";
import { AppShell } from "@/components/layout/app-shell";
import { PrimaryButton } from "@/components/ui/primary-button";
import { fetchJson } from "@/lib/http";
import { readPageData } from "@/lib/page-data";
import { CloudUpload, FileSpreadsheet, FileText, PackageCheck, ShieldCheck, Upload } from "lucide-react";
import { createRoot } from "react-dom/client";
import { StrictMode, useEffect, useMemo, useState } from "react";

type WorkflowOption = {
  name: string;
  label: string;
};

type TaskNewPayload = {
  workflows: WorkflowOption[];
  example_files?: string[];
  exampleFiles?: string[];
};

type CreateTaskResponse = {
  redirect_url: string;
};

function NewTaskPage() {
  const payload = useMemo(() => readPageData<TaskNewPayload>(), []);
  const exampleFiles = payload.exampleFiles ?? payload.example_files ?? [];
  const [workflow, setWorkflow] = useState(payload.workflows[0]?.name ?? "");
  const [submitter, setSubmitter] = useState("");
  const [remark, setRemark] = useState("");
  const [file, setFile] = useState<File | null>(null);
  const [submitting, setSubmitting] = useState(false);
  const [hint, setHint] = useState("");

  useEffect(() => {
    const remembered = window.localStorage.getItem("lingxing_submitter");
    if (remembered) {
      setSubmitter((current) => current || remembered);
    }
  }, []);

  async function handleSubmit(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (!file) {
      setHint("请先选择清单文件");
      return;
    }

    setSubmitting(true);
    setHint("任务正在创建，校验完成后会立即进入排队。");
    try {
      const formData = new FormData();
      formData.append("workflow_name", workflow);
      formData.append("submitter", submitter.trim());
      formData.append("remark", remark.trim());
      formData.append("manifest_file", file);
      const response = await fetchJson<CreateTaskResponse>("/api/tasks", {
        method: "POST",
        body: formData,
      });
      window.localStorage.setItem("lingxing_submitter", submitter.trim());
      window.location.href = response.redirect_url;
    } catch (error) {
      setHint(error instanceof Error ? error.message : "创建任务失败");
    } finally {
      setSubmitting(false);
    }
  }

  return (
    <AppShell
      eyebrow="新建任务"
      nav={[
        { href: "/tasks/new", label: "新建任务", current: true, icon: "new" },
        { href: "/tasks", label: "任务列表", icon: "list" },
      ]}
      title={
        <>
          把原来“命令行操作”的那条链路，
          <br />
          收敛成同事也能直接上手的上传页面。
        </>
      }
      subtitle={
        <>
          支持上传 `.txt` 和 `.xlsx` 清单，系统会自动解析 FBA、进入后台排队，并在完成后生成可下载的结果包。
          整体风格更偏内网工具里的“高级实用派”，重点是稳、清晰、好操作。
        </>
      }
      callout={
        <PointerHighlight
          rectangleClassName="rounded-full border-[color:oklch(0.73_0.03_188)]"
          pointerClassName="text-[color:oklch(0.53_0.08_188)]"
          containerClassName="max-w-fit rounded-full bg-white/70 px-4 py-2"
        >
          <span className="text-sm font-medium text-[color:oklch(0.35_0.03_230)]">
            当前模式是“多人可提交，浏览器执行始终保持单槽稳定运行”。
          </span>
        </PointerHighlight>
      }
      aside={
        <div className="rounded-[28px] border border-white/70 bg-white/78 p-5 shadow-[0_18px_70px_rgba(36,53,44,0.08)] backdrop-blur-xl">
          <p className="text-[0.72rem] font-semibold uppercase tracking-[0.24em] text-[color:oklch(0.55_0.03_205)]">
            适合谁用
          </p>
          <div className="mt-4 space-y-3 text-sm leading-6 text-[color:oklch(0.42_0.03_228)]">
            <p>适合内网同事按模板上传，无需再手工打开命令行。</p>
            <p>上传后可以关闭页面，稍后再回来查看排队状态和下载结果。</p>
            <p>每个任务都独立保留日志、报告和结果包，排错比以前直观很多。</p>
          </div>
        </div>
      }
    >
      <section className="grid gap-6 xl:grid-cols-[minmax(0,1.1fr)_360px]">
        <div className="rounded-[30px] border border-white/70 bg-white/82 p-6 shadow-[0_20px_80px_rgba(36,56,43,0.08)] backdrop-blur-xl">
          <div className="flex flex-wrap items-center justify-between gap-4">
            <div>
              <p className="text-[0.72rem] font-semibold uppercase tracking-[0.24em] text-[color:oklch(0.55_0.03_205)]">
                提交流程
              </p>
              <h2 className="mt-2 font-[family-name:var(--font-display)] text-2xl font-semibold tracking-[-0.03em] text-[color:oklch(0.22_0.025_242)]">
                上传清单并开始处理
              </h2>
            </div>
            <div className="inline-flex items-center gap-2 rounded-full bg-[color:oklch(0.97_0.01_95)] px-3 py-2 text-xs font-medium text-[color:oklch(0.43_0.03_228)]">
              <ShieldCheck className="h-4 w-4 text-[color:oklch(0.5_0.08_172)]" />
              自动校验格式，非法文件不会入队
            </div>
          </div>

          <form className="mt-6 space-y-5" onSubmit={handleSubmit}>
            <div className="grid gap-5 md:grid-cols-2">
              <label className="space-y-2">
                <span className="text-sm font-medium text-[color:oklch(0.32_0.02_232)]">流程类型</span>
                <select
                  value={workflow}
                  onChange={(event) => setWorkflow(event.target.value)}
                  className="min-h-12 w-full rounded-2xl border border-[color:oklch(0.88_0.01_95)] bg-[color:oklch(0.99_0.002_95)] px-4 text-sm text-[color:oklch(0.28_0.02_232)] outline-none transition focus:border-[color:oklch(0.6_0.09_190)] focus:ring-2 focus:ring-[color:oklch(0.86_0.03_190)]"
                >
                  {payload.workflows.map((item) => (
                    <option key={item.name} value={item.name}>
                      {item.label}
                    </option>
                  ))}
                </select>
              </label>

              <label className="space-y-2">
                <span className="text-sm font-medium text-[color:oklch(0.32_0.02_232)]">提交人</span>
                <input
                  value={submitter}
                  onChange={(event) => setSubmitter(event.target.value)}
                  placeholder="填写姓名或工号"
                  required
                  className="min-h-12 w-full rounded-2xl border border-[color:oklch(0.88_0.01_95)] bg-[color:oklch(0.99_0.002_95)] px-4 text-sm text-[color:oklch(0.28_0.02_232)] outline-none transition placeholder:text-[color:oklch(0.68_0.02_228)] focus:border-[color:oklch(0.6_0.09_190)] focus:ring-2 focus:ring-[color:oklch(0.86_0.03_190)]"
                />
              </label>
            </div>

            <label className="space-y-2">
              <span className="text-sm font-medium text-[color:oklch(0.32_0.02_232)]">备注</span>
              <input
                value={remark}
                onChange={(event) => setRemark(event.target.value)}
                placeholder="可选，方便在任务列表里快速识别"
                className="min-h-12 w-full rounded-2xl border border-[color:oklch(0.88_0.01_95)] bg-[color:oklch(0.99_0.002_95)] px-4 text-sm text-[color:oklch(0.28_0.02_232)] outline-none transition placeholder:text-[color:oklch(0.68_0.02_228)] focus:border-[color:oklch(0.6_0.09_190)] focus:ring-2 focus:ring-[color:oklch(0.86_0.03_190)]"
              />
            </label>

            <label className="block cursor-pointer rounded-[28px] border border-dashed border-[color:oklch(0.82_0.03_186)] bg-[linear-gradient(180deg,rgba(244,250,248,0.96),rgba(255,255,255,0.92))] p-6 transition hover:border-[color:oklch(0.62_0.08_188)]">
              <input
                type="file"
                accept=".txt,.xlsx"
                className="hidden"
                onChange={(event) => setFile(event.target.files?.[0] ?? null)}
              />
              <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                <div className="flex items-start gap-3">
                  <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-[linear-gradient(135deg,rgba(74,138,130,0.12),rgba(106,163,158,0.22))] text-[color:oklch(0.43_0.08_182)]">
                    <CloudUpload className="h-5 w-5" />
                  </div>
                  <div className="space-y-1">
                    <div className="text-sm font-semibold text-[color:oklch(0.24_0.02_232)]">
                      选择清单文件
                    </div>
                    <div className="text-sm leading-6 text-[color:oklch(0.47_0.03_228)]">
                      支持 `.txt` 与 `.xlsx`。如果是 Excel，系统会读取第一个工作表并自动识别 FBA 列。
                    </div>
                  </div>
                </div>
                <div className="rounded-full bg-white px-4 py-2 text-sm font-medium text-[color:oklch(0.33_0.03_232)] shadow-[0_12px_30px_rgba(41,61,52,0.08)]">
                  {file ? file.name : "点击选择文件"}
                </div>
              </div>
            </label>

            <div className="flex flex-wrap items-center gap-3">
              <PrimaryButton type="submit" disabled={submitting}>
                <Upload className="mr-2 h-4 w-4" />
                {submitting ? "正在提交" : "开始处理"}
              </PrimaryButton>
              <div className="min-h-6 text-sm text-[color:oklch(0.46_0.03_228)]">{hint}</div>
            </div>
          </form>
        </div>

        <div className="space-y-5">
          <BentoGrid className="mx-0 max-w-none md:auto-rows-[14rem] md:grid-cols-2">
            <BentoGridItem
              className="md:col-span-2 border-[color:oklch(0.89_0.02_95)] bg-[linear-gradient(180deg,rgba(255,255,255,0.96),rgba(246,244,239,0.88))] p-5 shadow-[0_18px_55px_rgba(41,59,49,0.08)]"
              icon={<PackageCheck className="h-5 w-5 text-[color:oklch(0.52_0.08_176)]" />}
              title="这页里最重要的两个文件"
              description="你可以直接下载示例，再把自己的 FBA 替换进去。这样对不会写格式的人也更友好。"
              header={
                <div className="rounded-2xl bg-[linear-gradient(135deg,rgba(66,145,137,0.12),rgba(223,182,106,0.12))] p-3 text-xs leading-6 text-[color:oklch(0.42_0.03_230)]">
                  示例文件会一直跟当前解析规则保持一致。
                </div>
              }
            />

            {exampleFiles.map((filename) => {
              const isTxt = filename.endsWith(".txt");
              return (
                <BentoGridItem
                  key={filename}
                  className="border-[color:oklch(0.89_0.02_95)] bg-white/92 p-5 shadow-[0_18px_55px_rgba(41,59,49,0.08)]"
                  icon={
                    isTxt ? (
                      <FileText className="h-5 w-5 text-[color:oklch(0.57_0.09_210)]" />
                    ) : (
                      <FileSpreadsheet className="h-5 w-5 text-[color:oklch(0.55_0.09_156)]" />
                    )
                  }
                  title={<a href={`/api/examples/${filename}`}>{filename}</a>}
                  description={isTxt ? "每行一个 FBA，可写注释。" : "自动识别列名，适合批量提交。"}
                  header={
                    <div className="rounded-2xl bg-[color:oklch(0.975_0.006_95)] p-3 text-sm text-[color:oklch(0.34_0.02_232)]">
                      <a className="font-medium hover:underline" href={`/api/examples/${filename}`}>
                        立即下载示例
                      </a>
                    </div>
                  }
                />
              );
            })}
          </BentoGrid>
        </div>
      </section>
    </AppShell>
  );
}

const root = document.getElementById("root");
if (root) {
  createRoot(root).render(
    <StrictMode>
      <NewTaskPage />
    </StrictMode>,
  );
}
