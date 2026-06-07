"use client";

import Link from "next/link";
import { useCallback, useRef, useState } from "react";
import { processWorkbook, type ProcessResult, type SheetReport } from "./lib/processTags";

type Status =
  | { kind: "idle" }
  | { kind: "working" }
  | { kind: "done"; result: ProcessResult; sourceName: string }
  | { kind: "error"; message: string };

const STEPS = [
  {
    n: 1,
    title: "Rebuild the pixel column",
    body: "Every pixel cell becomes a fresh ft.event URL built from the client, campaign and per-row placement id.",
  },
  {
    n: 2,
    title: "Clean the clicktags",
    body: "Strips the us_privacy=${US_PRIVACY} parameter out of every Update_Clicktag1 URL.",
  },
  {
    n: 3,
    title: "Drop the static column",
    body: "Deletes the entire Static_Clicktag1 column and shifts the rest of the sheet left.",
  },
];

export default function Home() {
  const [status, setStatus] = useState<Status>({ kind: "idle" });
  const [dragging, setDragging] = useState(false);
  const inputRef = useRef<HTMLInputElement>(null);

  const handleFile = useCallback(async (file: File) => {
    if (!/\.xlsx$/i.test(file.name)) {
      setStatus({ kind: "error", message: "That isn’t a .xlsx file — please pick an Excel workbook." });
      return;
    }
    setStatus({ kind: "working" });
    try {
      const buffer = await file.arrayBuffer();
      const result = processWorkbook(buffer, file.name);
      setStatus({ kind: "done", result, sourceName: file.name });
    } catch (err) {
      setStatus({
        kind: "error",
        message: err instanceof Error ? err.message : "Could not process the file.",
      });
    }
  }, []);

  const onDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragging(false);
      const file = e.dataTransfer.files?.[0];
      if (file) handleFile(file);
    },
    [handleFile]
  );

  const download = useCallback(() => {
    if (status.kind !== "done") return;
    const blob = new Blob([new Uint8Array(status.result.data)], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = status.result.fileName;
    a.click();
    URL.revokeObjectURL(url);
  }, [status]);

  const reset = () => {
    setStatus({ kind: "idle" });
    if (inputRef.current) inputRef.current.value = "";
  };

  return (
    <div className="relative min-h-full overflow-hidden bg-zinc-50 text-zinc-900 dark:bg-[#08080a] dark:text-zinc-100">
      {/* ambient glow */}
      <div
        aria-hidden
        className="pointer-events-none absolute inset-x-0 top-0 h-[420px] bg-[radial-gradient(60%_100%_at_50%_0%,rgba(16,185,129,0.14),transparent_70%)]"
      />

      <div className="relative mx-auto w-full max-w-[1440px] px-6 py-10 sm:px-10 lg:py-16">
        {/* header */}
        <header className="flex flex-col gap-6 sm:flex-row sm:items-center sm:justify-between">
          <div className="flex items-center gap-4">
            <div className="flex h-11 w-11 items-center justify-center rounded-xl bg-gradient-to-br from-emerald-400 to-teal-600 text-white shadow-lg shadow-emerald-500/20">
              <BoltIcon className="h-6 w-6" />
            </div>
            <div>
              <h1 className="text-xl font-semibold tracking-tight sm:text-2xl">Flashtalking Tag Cleaner</h1>
              <p className="text-sm text-zinc-500 dark:text-zinc-400">
                Fix ft_tags trafficking sheets in one drop.
              </p>
            </div>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <Link
              href="/placements"
              className="inline-flex w-fit items-center gap-2 rounded-full border border-zinc-200 bg-white px-3 py-1.5 text-xs font-medium text-zinc-600 transition-colors hover:bg-zinc-100 dark:border-zinc-800 dark:bg-zinc-900 dark:text-zinc-400 dark:hover:bg-zinc-800"
            >
              Placement Merge →
            </Link>
            <span className="inline-flex w-fit items-center gap-2 rounded-full border border-zinc-200 bg-white px-3 py-1.5 text-xs font-medium text-zinc-600 dark:border-zinc-800 dark:bg-zinc-900 dark:text-zinc-400">
              <span className="h-2 w-2 rounded-full bg-emerald-500" />
              Runs in your browser · nothing uploaded
            </span>
          </div>
        </header>

        {/* steps */}
        <div className="mt-10 grid gap-4 sm:grid-cols-2 lg:grid-cols-3">
          {STEPS.map((s) => (
            <div
              key={s.n}
              className="rounded-2xl border border-zinc-200 bg-white p-5 dark:border-zinc-800 dark:bg-zinc-900/50"
            >
              <div className="flex h-8 w-8 items-center justify-center rounded-lg bg-zinc-100 text-sm font-semibold text-zinc-700 dark:bg-zinc-800 dark:text-zinc-200">
                {s.n}
              </div>
              <h3 className="mt-3 text-sm font-semibold">{s.title}</h3>
              <p className="mt-1 text-sm leading-6 text-zinc-500 dark:text-zinc-400">{s.body}</p>
            </div>
          ))}
        </div>

        {/* dropzone */}
        <label
          onDragOver={(e) => {
            e.preventDefault();
            setDragging(true);
          }}
          onDragLeave={() => setDragging(false)}
          onDrop={onDrop}
          className={`group mt-6 flex cursor-pointer flex-col items-center justify-center gap-3 rounded-3xl border-2 border-dashed px-6 py-16 text-center transition-all ${
            dragging
              ? "border-emerald-500 bg-emerald-50 dark:bg-emerald-950/30"
              : "border-zinc-300 bg-white hover:border-emerald-400 hover:bg-emerald-50/40 dark:border-zinc-700 dark:bg-zinc-900/40 dark:hover:border-emerald-500/60 dark:hover:bg-emerald-950/10"
          }`}
        >
          <input
            ref={inputRef}
            type="file"
            accept=".xlsx"
            className="hidden"
            onChange={(e) => {
              const file = e.target.files?.[0];
              if (file) handleFile(file);
            }}
          />
          <div
            className={`flex h-14 w-14 items-center justify-center rounded-2xl transition-colors ${
              dragging
                ? "bg-emerald-500 text-white"
                : "bg-zinc-100 text-zinc-500 group-hover:bg-emerald-500 group-hover:text-white dark:bg-zinc-800"
            }`}
          >
            {status.kind === "working" ? (
              <Spinner className="h-6 w-6" />
            ) : (
              <UploadIcon className="h-6 w-6" />
            )}
          </div>
          <div className="space-y-1">
            <p className="text-base font-medium">
              {status.kind === "working"
                ? "Processing…"
                : "Drop a .xlsx file here, or click to browse"}
            </p>
            <p className="text-sm text-zinc-500 dark:text-zinc-400">
              Your logo, fonts and styling are preserved.
            </p>
          </div>
        </label>

        {status.kind === "error" && (
          <div className="mt-6 flex items-start gap-3 rounded-2xl border border-red-300 bg-red-50 px-5 py-4 text-sm text-red-700 dark:border-red-900/70 dark:bg-red-950/40 dark:text-red-300">
            <WarnIcon className="mt-0.5 h-5 w-5 shrink-0" />
            <p className="flex-1">{status.message}</p>
            <button onClick={reset} className="font-medium underline underline-offset-2">
              try again
            </button>
          </div>
        )}

        {status.kind === "done" && (
          <Results status={status} onDownload={download} onReset={reset} />
        )}
      </div>
    </div>
  );
}

function Results({
  status,
  onDownload,
  onReset,
}: {
  status: Extract<Status, { kind: "done" }>;
  onDownload: () => void;
  onReset: () => void;
}) {
  const { result, sourceName } = status;
  const totalPixels = result.sheets.reduce((n, s) => n + s.pixelsReplaced, 0);
  const totalClicktags = result.sheets.reduce((n, s) => n + s.urlsCleaned, 0);

  return (
    <section className="mt-8 space-y-8">
      {/* success banner */}
      <div className="flex flex-col gap-4 rounded-2xl border border-emerald-300 bg-gradient-to-br from-emerald-50 to-teal-50 px-6 py-5 dark:border-emerald-900/70 dark:from-emerald-950/40 dark:to-teal-950/20 sm:flex-row sm:items-center sm:justify-between">
        <div className="flex items-center gap-3">
          <div className="flex h-10 w-10 items-center justify-center rounded-full bg-emerald-500 text-white">
            <CheckIcon className="h-5 w-5" />
          </div>
          <div>
            <p className="font-semibold text-emerald-900 dark:text-emerald-200">All done</p>
            <p className="text-sm text-emerald-700 dark:text-emerald-300/80">
              {totalPixels} pixel · {totalClicktags} clicktag URLs updated in{" "}
              <span className="font-medium">{sourceName}</span>
            </p>
          </div>
        </div>
        <div className="flex shrink-0 items-center gap-2">
          <button
            onClick={onDownload}
            className="inline-flex items-center gap-2 rounded-full bg-emerald-600 px-5 py-2.5 text-sm font-medium text-white shadow-sm transition-colors hover:bg-emerald-700"
          >
            <DownloadIcon className="h-4 w-4" />
            Download cleaned file
          </button>
          <button
            onClick={onReset}
            className="rounded-full border border-zinc-300 bg-white px-4 py-2.5 text-sm font-medium text-zinc-700 transition-colors hover:bg-zinc-100 dark:border-zinc-700 dark:bg-zinc-900 dark:text-zinc-200 dark:hover:bg-zinc-800"
          >
            Clear
          </button>
        </div>
      </div>

      {result.sheets.map((s) => (
        <SheetResult key={s.sheetFile} sheet={s} />
      ))}
    </section>
  );
}

function SheetResult({ sheet: s }: { sheet: SheetReport }) {
  return (
    <div className="space-y-5">
      {/* stat cards */}
      <div className="grid gap-4 sm:grid-cols-3">
        <StatCard
          label="Pixel URLs rebuilt"
          value={s.pixelsReplaced}
          sub={
            s.clientId && s.campaignId
              ? `client ${s.clientId} · campaign ${s.campaignId}`
              : s.pixelColumn
              ? `column ${s.pixelColumn}`
              : undefined
          }
          tone="emerald"
        />
        <StatCard
          label="Clicktags cleaned"
          value={s.urlsCleaned}
          sub={`Update_Clicktag1 · column ${s.updateColumn}`}
          tone="sky"
        />
        <StatCard
          label="Column deleted"
          value={s.staticColumn ? 1 : 0}
          sub={s.staticColumn ? `Static_Clicktag1 · column ${s.staticColumn}` : "none found"}
          tone="violet"
        />
      </div>

      {/* example URLs */}
      <div className="grid gap-4 lg:grid-cols-2">
        {s.pixelSample && (
          <CodePanel title="New pixel URL">
            <span className="break-all text-emerald-700 dark:text-emerald-400">{s.pixelSample}</span>
          </CodePanel>
        )}
        {s.sample && (
          <CodePanel title="Clicktag change">
            <span className="break-all text-red-600 line-through decoration-red-400/50 dark:text-red-400">
              {s.sample.before}
            </span>
            <span className="mt-1 break-all text-emerald-700 dark:text-emerald-400">{s.sample.after}</span>
          </CodePanel>
        )}
      </div>

      <PreviewTable headers={s.headers} rows={s.rows} pixelCol={s.pixelColumn} updateCol={s.updateColumn} />
    </div>
  );
}

const TONES = {
  emerald: "text-emerald-600 dark:text-emerald-400",
  sky: "text-sky-600 dark:text-sky-400",
  violet: "text-violet-600 dark:text-violet-400",
} as const;

function StatCard({
  label,
  value,
  sub,
  tone,
}: {
  label: string;
  value: number;
  sub?: string;
  tone: keyof typeof TONES;
}) {
  return (
    <div className="rounded-2xl border border-zinc-200 bg-white p-5 dark:border-zinc-800 dark:bg-zinc-900/50">
      <p className="text-sm text-zinc-500 dark:text-zinc-400">{label}</p>
      <p className={`mt-1 text-3xl font-semibold tabular-nums ${TONES[tone]}`}>{value}</p>
      {sub && <p className="mt-1 truncate text-xs text-zinc-500 dark:text-zinc-500" title={sub}>{sub}</p>}
    </div>
  );
}

function CodePanel({ title, children }: { title: string; children: React.ReactNode }) {
  return (
    <div className="overflow-hidden rounded-2xl border border-zinc-200 dark:border-zinc-800">
      <div className="border-b border-zinc-200 bg-zinc-50 px-4 py-2.5 text-xs font-semibold uppercase tracking-wide text-zinc-500 dark:border-zinc-800 dark:bg-zinc-900">
        {title}
      </div>
      <div className="flex flex-col gap-1 bg-white px-4 py-3 font-mono text-xs leading-5 dark:bg-zinc-950/40">
        {children}
      </div>
    </div>
  );
}

function PreviewTable({
  headers,
  rows,
  pixelCol,
  updateCol,
}: {
  headers: string[];
  rows: string[][];
  pixelCol: string | null;
  updateCol: string;
}) {
  if (!headers.length) return null;
  // After processing, columns are contiguous from A. Map header index -> letter.
  const letterAt = (i: number) => String.fromCharCode(65 + i);
  const highlight = (i: number) => {
    const l = letterAt(i);
    return l === pixelCol || l === updateCol;
  };

  return (
    <div className="overflow-hidden rounded-2xl border border-zinc-200 dark:border-zinc-800">
      <div className="border-b border-zinc-200 bg-zinc-50 px-4 py-2.5 text-xs font-semibold uppercase tracking-wide text-zinc-500 dark:border-zinc-800 dark:bg-zinc-900">
        Preview · {rows.length} row{rows.length === 1 ? "" : "s"}
      </div>
      <div className="overflow-x-auto">
        <table className="min-w-full border-collapse text-left text-xs">
          <thead className="bg-zinc-50 dark:bg-zinc-900/80">
            <tr>
              {headers.map((h, i) => (
                <th
                  key={i}
                  className={`whitespace-nowrap px-3 py-2.5 font-semibold ${
                    highlight(i)
                      ? "text-emerald-700 dark:text-emerald-400"
                      : "text-zinc-600 dark:text-zinc-300"
                  }`}
                >
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, ri) => (
              <tr
                key={ri}
                className="border-t border-zinc-100 even:bg-zinc-50/60 dark:border-zinc-800/70 dark:even:bg-zinc-900/30"
              >
                {row.map((cell, ci) => (
                  <td
                    key={ci}
                    title={cell}
                    className={`max-w-[24rem] truncate px-3 py-2 ${
                      highlight(ci)
                        ? "font-mono text-zinc-700 dark:text-zinc-300"
                        : "text-zinc-500 dark:text-zinc-400"
                    }`}
                  >
                    {cell}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------- icons ---------- */

function UploadIcon({ className }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
      <polyline points="17 8 12 3 7 8" />
      <line x1="12" y1="3" x2="12" y2="15" />
    </svg>
  );
}

function DownloadIcon({ className }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
      <polyline points="7 10 12 15 17 10" />
      <line x1="12" y1="15" x2="12" y2="3" />
    </svg>
  );
}

function CheckIcon({ className }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.5" strokeLinecap="round" strokeLinejoin="round">
      <polyline points="20 6 9 17 4 12" />
    </svg>
  );
}

function WarnIcon({ className }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M10.29 3.86 1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z" />
      <line x1="12" y1="9" x2="12" y2="13" />
      <line x1="12" y1="17" x2="12.01" y2="17" />
    </svg>
  );
}

function BoltIcon({ className }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="currentColor">
      <path d="M13 2 4.5 13.5H11l-1 8.5 8.5-11.5H12l1-8.5z" />
    </svg>
  );
}

function Spinner({ className }: { className?: string }) {
  return (
    <svg className={`animate-spin ${className ?? ""}`} viewBox="0 0 24 24" fill="none">
      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
      <path className="opacity-90" fill="currentColor" d="M4 12a8 8 0 0 1 8-8V0C5.37 0 0 5.37 0 12h4z" />
    </svg>
  );
}
