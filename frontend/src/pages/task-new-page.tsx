import { AppShell } from "@/components/layout/app-shell";
import { PrimaryButton } from "@/components/ui/primary-button";
import { fetchJson } from "@/lib/http";
import { readPageData } from "@/lib/page-data";
import { AlertCircle, CheckCircle2, ClipboardList, CloudUpload, FileText, ShieldCheck, Upload } from "lucide-react";
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

const AMAZON_HL_WORKFLOW_NAME = "amazon_hl_csv_process";

function NewTaskPage() {
  const payload = useMemo(() => readPageData<TaskNewPayload>(), []);
  const exampleFiles = (payload.exampleFiles ?? payload.example_files ?? []).filter((filename) => filename.endsWith(".csv"));
  const [workflow, setWorkflow] = useState(payload.workflows[0]?.name ?? "");
  const [submitter, setSubmitter] = useState("");
  const [fbaText, setFbaText] = useState("");
  const [files, setFiles] = useState<File[]>([]);
  const [submitting, setSubmitting] = useState(false);
  const [hint, setHint] = useState("");
  const fbaCheck = useMemo(() => parseFbaTextForPreview(fbaText), [fbaText]);
  const isAmazonHlWorkflow = workflow === AMAZON_HL_WORKFLOW_NAME;

  useEffect(() => {
    const remembered = window.localStorage.getItem("lingxing_submitter");
    if (remembered) {
      setSubmitter((current) => current || remembered);
    }
  }, []);

  async function handleSubmit(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault();
    if (isAmazonHlWorkflow) {
      if (!files.length) {
        setHint("HL 发货请上传一个或多个 Amazon 后台导出的 CSV 文件。");
        return;
      }
      const invalidFiles = files.filter((item) => !item.name.toLowerCase().endsWith(".csv"));
      if (invalidFiles.length) {
        setHint(`HL 发货只支持 .csv 文件：${invalidFiles.slice(0, 3).map((item) => item.name).join("、")}`);
        return;
      }
    } else {
      if (!fbaCheck.codes.length) {
        setHint("正常/UPS 流程请直接粘贴 FBA 号，一行一个。");
        return;
      }
      if (fbaCheck.invalid.length) {
        setHint(`FBA号格式不正确，请先检查：${fbaCheck.invalid.slice(0, 3).join("、")}`);
        return;
      }
    }

    setSubmitting(true);
    setHint("任务正在创建，校验完成后会立即进入排队。");
    try {
      const formData = new FormData();
      formData.append("workflow_name", workflow);
      formData.append("submitter", submitter.trim());
      if (isAmazonHlWorkflow) {
        files.forEach((item) => formData.append("manifest_files", item));
      } else {
        formData.append("fba_text", fbaCheck.codes.join("\n"));
      }
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
      introMode="none"
      title={
        <>
          上传清单并开始处理
        </>
      }
      subtitle={
        <>
          正常/UPS 粘贴 FBA 号；HL 发货上传 Amazon 后台 CSV。系统会自动校验并排队处理。
        </>
      }
    >
      <section className="grid gap-5 xl:grid-cols-[minmax(0,1.14fr)_320px]">
        <div className="rounded-[30px] border border-white/70 bg-white/82 p-6 shadow-[0_20px_80px_rgba(36,56,43,0.08)] backdrop-blur-xl">
          <div className="flex flex-wrap items-center justify-between gap-4">
            <div>
              <p className="text-[0.72rem] font-semibold uppercase tracking-[0.24em] text-[color:oklch(0.55_0.03_205)]">
                新建任务
              </p>
              <h1 className="mt-2 font-[family-name:var(--font-display)] text-2xl font-semibold tracking-[-0.03em] text-[color:oklch(0.22_0.025_242)]">
                上传清单并开始处理
              </h1>
              <p className="mt-2 max-w-2xl text-sm leading-6 text-[color:oklch(0.46_0.03_228)]">
                正常/UPS 直接粘贴 FBA 号；领星暂时抓不到的 HL 货件, 选择 HL 流程后上传 Amazon CSV。
              </p>
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
                  onChange={(event) => {
                    setWorkflow(event.target.value);
                    setFiles([]);
                    setHint("");
                  }}
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

            <section className="rounded-[28px] border border-[color:oklch(0.88_0.018_95)] bg-[linear-gradient(180deg,rgba(255,255,255,0.92),rgba(248,247,242,0.78))] p-4">
              <div className="flex flex-col gap-3 sm:flex-row sm:items-start sm:justify-between">
                <div className="flex items-start gap-3">
                  <div className="flex h-11 w-11 shrink-0 items-center justify-center rounded-2xl bg-[color:oklch(0.94_0.035_178)] text-[color:oklch(0.4_0.08_182)]">
                    <ClipboardList className="h-5 w-5" />
                  </div>
                  <div>
                    <div className="text-sm font-semibold text-[color:oklch(0.24_0.02_232)]">方式一：直接粘贴 FBA 号 (正常/UPS)</div>
                    <div className="mt-1 text-sm leading-6 text-[color:oklch(0.45_0.03_228)]">
                      多个 FBA 一行一个, 或从表格整列复制过来。正常发货和 UPS 发货都走这个入口。
                    </div>
                  </div>
                </div>
                <FbaCheckBadge codes={fbaCheck.codes.length} invalid={fbaCheck.invalid.length} />
              </div>
              <textarea
                name="fba_text"
                value={fbaText}
                onChange={(event) => setFbaText(event.target.value)}
                placeholder={"例如：\nFBA19C2P8D5D\nFBA19BXBL1MT"}
                disabled={isAmazonHlWorkflow}
                className="mt-4 min-h-32 w-full resize-y rounded-2xl border border-[color:oklch(0.86_0.016_95)] bg-white/92 px-4 py-3 text-sm leading-6 text-[color:oklch(0.27_0.025_232)] outline-none transition placeholder:text-[color:oklch(0.67_0.02_228)] focus:border-[color:oklch(0.6_0.09_190)] focus:ring-2 focus:ring-[color:oklch(0.86_0.03_190)]"
              />
              {fbaCheck.invalid.length ? (
                <div className="mt-3 flex items-start gap-2 rounded-2xl bg-[color:oklch(0.965_0.025_35)] px-3 py-2 text-sm font-medium text-[color:oklch(0.43_0.08_35)]">
                  <AlertCircle className="mt-0.5 h-4 w-4 shrink-0" />
                  以下内容不是有效 FBA 号：{fbaCheck.invalid.slice(0, 5).join("、")}
                </div>
              ) : null}
            </section>

            <label className="block cursor-pointer rounded-[28px] border border-dashed border-[color:oklch(0.82_0.03_186)] bg-[linear-gradient(180deg,rgba(244,250,248,0.96),rgba(255,255,255,0.92))] p-6 transition hover:border-[color:oklch(0.62_0.08_188)]">
              <input
                type="file"
                accept=".csv"
                multiple
                className="hidden"
                disabled={!isAmazonHlWorkflow}
                onChange={(event) => setFiles(Array.from(event.target.files ?? []))}
              />
              <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                <div className="flex items-start gap-3">
                  <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-[linear-gradient(135deg,rgba(74,138,130,0.12),rgba(106,163,158,0.22))] text-[color:oklch(0.43_0.08_182)]">
                    <CloudUpload className="h-5 w-5" />
                  </div>
                  <div className="space-y-1">
                    <div className="text-sm font-semibold text-[color:oklch(0.24_0.02_232)]">
                      上传 HL Amazon CSV 文件
                    </div>
                    <div className="text-sm leading-6 text-[color:oklch(0.47_0.03_228)]">
                      方式二：一次可上传多个 Amazon CSV；也支持一个 CSV 内包含多个 FBA 货件。
                    </div>
                    {!isAmazonHlWorkflow ? (
                      <div className="text-xs font-medium text-[color:oklch(0.5_0.04_70)]">
                        需要先在“流程类型”里选择 HL 发货 Amazon CSV 整理。
                      </div>
                    ) : null}
                  </div>
                </div>
                <div className="rounded-full bg-white px-4 py-2 text-sm font-medium text-[color:oklch(0.33_0.03_232)] shadow-[0_12px_30px_rgba(41,61,52,0.08)]">
                  {files.length ? `已选 ${files.length} 个 CSV` : isAmazonHlWorkflow ? "点击选择 CSV 文件" : "选择 HL 流程后启用"}
                </div>
              </div>
              {files.length ? (
                <div className="mt-4 flex flex-wrap gap-2">
                  {files.slice(0, 6).map((item, index) => (
                    <span
                      key={`${item.name}-${item.size}-${index}`}
                      className="rounded-full bg-white/86 px-3 py-1 text-xs font-medium text-[color:oklch(0.42_0.03_228)]"
                    >
                      {item.name}
                    </span>
                  ))}
                  {files.length > 6 ? (
                    <span className="rounded-full bg-white/86 px-3 py-1 text-xs font-medium text-[color:oklch(0.42_0.03_228)]">
                      另外 {files.length - 6} 个
                    </span>
                  ) : null}
                </div>
              ) : null}
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

        <div className="rounded-[30px] border border-white/70 bg-white/82 p-5 shadow-[0_20px_70px_rgba(36,56,43,0.08)] backdrop-blur-xl">
          <p className="text-[0.72rem] font-semibold uppercase tracking-[0.24em] text-[color:oklch(0.55_0.03_205)]">
            示例文件
          </p>
          <h2 className="mt-2 font-[family-name:var(--font-display)] text-2xl font-semibold tracking-[-0.03em] text-[color:oklch(0.22_0.025_242)]">
            Amazon CSV 格式
          </h2>
          <p className="mt-2 text-sm leading-6 text-[color:oklch(0.46_0.03_228)]">
            HL 发货只需要上传 Amazon 后台导出的货件信息 CSV, 可一次选择多个文件。
          </p>
          <div className="mt-4 rounded-[22px] bg-[color:oklch(0.97_0.012_95)] p-4 text-sm leading-6 text-[color:oklch(0.42_0.03_228)]">
            必须包含: 货件编号, SKU, 商品名称, FNSKU, 原厂包装模板名称, 每箱件数, 箱子总数, 商品总数, 箱号。
          </div>
          <div className="mt-5 space-y-3">
            {exampleFiles.map((filename) => {
              return (
                <a
                  key={filename}
                  href={`/api/examples/${filename}`}
                  className="flex items-start gap-3 rounded-[22px] border border-[color:oklch(0.89_0.02_95)] bg-white/88 p-4 text-[color:oklch(0.28_0.025_232)] shadow-[0_12px_32px_rgba(41,59,49,0.06)] transition hover:translate-y-[-1px] hover:bg-white"
                >
                  <span className="flex h-10 w-10 shrink-0 items-center justify-center rounded-2xl bg-[color:oklch(0.94_0.035_178)] text-[color:oklch(0.42_0.08_182)]">
                    <FileText className="h-5 w-5" />
                  </span>
                  <span>
                    <span className="block text-sm font-semibold">{filename}</span>
                    <span className="mt-1 block text-xs leading-5 text-[color:oklch(0.48_0.03_228)]">
                      Amazon 后台导出的 HL 货件 CSV 示例。
                    </span>
                  </span>
                </a>
              );
            })}
            {!exampleFiles.length ? (
              <div className="rounded-[22px] border border-[color:oklch(0.89_0.02_95)] bg-white/88 p-4 text-sm leading-6 text-[color:oklch(0.48_0.03_228)]">
                当前环境暂未提供示例文件, 请直接使用 Amazon 后台导出的 CSV。
              </div>
            ) : null}
          </div>
        </div>
      </section>
    </AppShell>
  );
}

function parseFbaTextForPreview(text: string): { codes: string[]; invalid: string[] } {
  const codes: string[] = [];
  const invalid: string[] = [];
  const seen = new Set<string>();
  for (const rawLine of text.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line || line.startsWith("#")) {
      continue;
    }
    const tokens = line.match(/[A-Za-z0-9-]+/g) ?? [];
    for (const token of tokens) {
      const normalized = token.toUpperCase();
      if (/^FBA[A-Z0-9-]+$/.test(normalized)) {
        if (!seen.has(normalized)) {
          seen.add(normalized);
          codes.push(normalized);
        }
      } else {
        invalid.push(token);
      }
    }
  }
  return { codes, invalid };
}

function FbaCheckBadge({ codes, invalid }: { codes: number; invalid: number }) {
  if (invalid > 0) {
    return (
      <div className="inline-flex items-center gap-2 rounded-full bg-[color:oklch(0.965_0.025_35)] px-3 py-2 text-xs font-semibold text-[color:oklch(0.43_0.08_35)]">
        <AlertCircle className="h-3.5 w-3.5" />
        {invalid} 个格式待检查
      </div>
    );
  }
  if (codes > 0) {
    return (
      <div className="inline-flex items-center gap-2 rounded-full bg-[color:oklch(0.945_0.04_165)] px-3 py-2 text-xs font-semibold text-[color:oklch(0.32_0.06_170)]">
        <CheckCircle2 className="h-3.5 w-3.5" />
        已识别 {codes} 个 FBA
      </div>
    );
  }
  return (
    <div className="inline-flex items-center gap-2 rounded-full bg-[color:oklch(0.975_0.008_95)] px-3 py-2 text-xs font-semibold text-[color:oklch(0.44_0.025_232)]">
      等待粘贴
    </div>
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
