"use client";

import Link from "next/link";
import { useCallback, useMemo, useRef, useState } from "react";
import {
  analyze,
  build,
  type AnalyzeResult,
  type BuildResult,
  type PlacementPlan,
  type SiteSelections,
} from "../lib/mergePlacements";

type Phase =
  | { kind: "idle" }
  | { kind: "working"; label: string }
  | { kind: "review"; analysis: AnalyzeResult; targetName: string }
  | { kind: "done"; result: BuildResult; targetName: string }
  | { kind: "error"; message: string };

export default function PlacementsPage() {
  const [src, setSrc] = useState<File | null>(null);
  const [tgt, setTgt] = useState<File | null>(null);
  const [phase, setPhase] = useState<Phase>({ kind: "idle" });
  // candidate index chosen per placement name (-1 = leave blank)
  const [selections, setSelections] = useState<Record<string, number>>({});
  const buffers = useRef<{ src: ArrayBuffer; tgt: ArrayBuffer } | null>(null);

  const runAnalyze = useCallback(async (source: File, target: File) => {
    setPhase({ kind: "working", label: "Reading files…" });
    try {
      const [s, t] = await Promise.all([source.arrayBuffer(), target.arrayBuffer()]);
      buffers.current = { src: s, tgt: t };
      const analysis = analyze(s, t);
      // default each appendable placement to its first candidate
      const defaults: Record<string, number> = {};
      for (const p of analysis.plans) {
        if (p.action === "append") defaults[p.name] = p.candidates.length ? 0 : -1;
      }
      setSelections(defaults);
      setPhase({ kind: "review", analysis, targetName: target.name });
    } catch (err) {
      setPhase({ kind: "error", message: err instanceof Error ? err.message : "Could not read the files." });
    }
  }, []);

  const pickSrc = useCallback(
    (file: File) => {
      setSrc(file);
      if (tgt) runAnalyze(file, tgt);
      else setPhase({ kind: "idle" });
    },
    [tgt, runAnalyze]
  );
  const pickTgt = useCallback(
    (file: File) => {
      setTgt(file);
      if (src) runAnalyze(src, file);
      else setPhase({ kind: "idle" });
    },
    [src, runAnalyze]
  );

  const apply = useCallback(() => {
    if (phase.kind !== "review" || !buffers.current) return;
    const { analysis, targetName } = phase;
    setPhase({ kind: "working", label: "Writing placements…" });
    // let the working state paint before the (synchronous) build
    setTimeout(() => {
      try {
        const picks: SiteSelections = {};
        for (const p of analysis.plans) {
          if (p.action !== "append") continue;
          const idx = selections[p.name] ?? -1;
          picks[p.name] = idx >= 0 ? p.candidates[idx] : null;
        }
        const result = build(buffers.current!.src, buffers.current!.tgt, targetName, picks);
        setPhase({ kind: "done", result, targetName });
      } catch (err) {
        setPhase({ kind: "error", message: err instanceof Error ? err.message : "Could not build the file." });
      }
    }, 20);
  }, [phase, selections]);

  const download = useCallback(() => {
    if (phase.kind !== "done") return;
    const blob = new Blob([new Uint8Array(phase.result.data)], {
      type: "application/vnd.ms-excel.sheet.macroEnabled.12",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = phase.result.fileName;
    a.click();
    URL.revokeObjectURL(url);
  }, [phase]);

  const reset = () => {
    setSrc(null);
    setTgt(null);
    setSelections({});
    buffers.current = null;
    setPhase({ kind: "idle" });
  };

  return (
    <div className="relative min-h-full overflow-hidden bg-zinc-50 text-zinc-900 dark:bg-[#08080a] dark:text-zinc-100">
      <div
        aria-hidden
        className="pointer-events-none absolute inset-x-0 top-0 h-[420px] bg-[radial-gradient(60%_100%_at_50%_0%,rgba(99,102,241,0.16),transparent_70%)]"
      />
      <div className="relative mx-auto w-full max-w-[1440px] px-6 py-10 sm:px-10 lg:py-16">
        <header className="flex flex-col gap-6 sm:flex-row sm:items-center sm:justify-between">
          <div className="flex items-center gap-4">
            <div className="flex h-11 w-11 items-center justify-center rounded-xl bg-gradient-to-br from-indigo-400 to-violet-600 text-white shadow-lg shadow-indigo-500/20">
              <LayersIcon className="h-6 w-6" />
            </div>
            <div>
              <h1 className="text-xl font-semibold tracking-tight sm:text-2xl">Placement Merge</h1>
              <p className="text-sm text-zinc-500 dark:text-zinc-400">
                Add media-plan placements + resolve sites into the campaign import sheet.
              </p>
            </div>
          </div>
          <Link
            href="/"
            className="inline-flex w-fit items-center gap-2 rounded-full border border-zinc-200 bg-white px-3 py-1.5 text-xs font-medium text-zinc-600 transition-colors hover:bg-zinc-100 dark:border-zinc-800 dark:bg-zinc-900 dark:text-zinc-400 dark:hover:bg-zinc-800"
          >
            <ArrowLeftIcon className="h-3.5 w-3.5" />
            Tag Cleaner
          </Link>
        </header>

        <div className="mt-10 grid gap-4 lg:grid-cols-2">
          <FileCard
            step={1}
            label="Prerequisites / media plan"
            hint="The .xlsx with “Placement Name” + “Publisher” columns."
            accept=".xlsx"
            file={src}
            onPick={pickSrc}
          />
          <FileCard
            step={2}
            label="Campaign import sheet"
            hint="The .xlsm with the Import + “Site Lookups” sheets (this file gets updated)."
            accept=".xlsm,.xlsx"
            file={tgt}
            onPick={pickTgt}
          />
        </div>

        {phase.kind === "working" && (
          <p className="mt-6 flex items-center gap-2 text-sm text-zinc-500 dark:text-zinc-400">
            <Spinner className="h-4 w-4" /> {phase.label}
          </p>
        )}

        {phase.kind === "error" && (
          <div className="mt-6 flex items-start gap-3 rounded-2xl border border-red-300 bg-red-50 px-5 py-4 text-sm text-red-700 dark:border-red-900/70 dark:bg-red-950/40 dark:text-red-300">
            <WarnIcon className="mt-0.5 h-5 w-5 shrink-0" />
            <p className="flex-1">{phase.message}</p>
            <button onClick={reset} className="font-medium underline underline-offset-2">
              start over
            </button>
          </div>
        )}

        {phase.kind === "review" && (
          <Review
            analysis={phase.analysis}
            selections={selections}
            onSelect={(name, idx) => setSelections((s) => ({ ...s, [name]: idx }))}
            onApply={apply}
          />
        )}

        {phase.kind === "done" && <Done result={phase.result} targetName={phase.targetName} onDownload={download} onReset={reset} />}
      </div>
    </div>
  );
}

function FileCard({
  step,
  label,
  hint,
  accept,
  file,
  onPick,
}: {
  step: number;
  label: string;
  hint: string;
  accept: string;
  file: File | null;
  onPick: (f: File) => void;
}) {
  const [dragging, setDragging] = useState(false);
  return (
    <label
      onDragOver={(e) => {
        e.preventDefault();
        setDragging(true);
      }}
      onDragLeave={() => setDragging(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragging(false);
        const f = e.dataTransfer.files?.[0];
        if (f) onPick(f);
      }}
      className={`flex cursor-pointer flex-col gap-4 rounded-3xl border-2 border-dashed p-6 transition-all ${
        dragging
          ? "border-indigo-500 bg-indigo-50/60 dark:bg-indigo-950/20"
          : file
          ? "border-zinc-200 bg-white dark:border-zinc-800 dark:bg-zinc-900/50"
          : "border-zinc-300 bg-white hover:border-indigo-400 hover:bg-indigo-50/30 dark:border-zinc-700 dark:bg-zinc-900/40 dark:hover:border-indigo-500/60"
      }`}
    >
      <input
        type="file"
        accept={accept}
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) onPick(f);
        }}
      />
      <div className="flex items-center gap-3">
        <div className="flex h-10 w-10 items-center justify-center rounded-xl bg-gradient-to-br from-indigo-400 to-violet-600 text-white shadow-lg shadow-indigo-500/20">
          {step}
        </div>
        <div>
          <p className="text-sm font-semibold">{label}</p>
          <p className="text-xs text-zinc-500 dark:text-zinc-400">{hint}</p>
        </div>
      </div>
      {file ? (
        <div className="flex items-center gap-2 rounded-xl bg-zinc-100 px-3 py-2 text-sm dark:bg-zinc-800/70">
          <FileIcon className="h-4 w-4 shrink-0 text-zinc-500" />
          <span className="truncate font-medium">{file.name}</span>
          <span className="ml-auto shrink-0 text-xs text-zinc-500">{(file.size / 1024).toFixed(0)} KB</span>
        </div>
      ) : (
        <p className="text-sm text-zinc-400 dark:text-zinc-500">Drop a file here, or click to browse</p>
      )}
    </label>
  );
}

function Review({
  analysis,
  selections,
  onSelect,
  onApply,
}: {
  analysis: AnalyzeResult;
  selections: Record<string, number>;
  onSelect: (name: string, idx: number) => void;
  onApply: () => void;
}) {
  const appendable = analysis.plans.filter((p) => p.action === "append");
  const existing = analysis.plans.length - appendable.length;
  const noMatch = appendable.filter((p) => p.candidates.length === 0).length;

  return (
    <section className="mt-8 space-y-5">
      <div className="flex flex-col gap-4 rounded-2xl border border-zinc-200 bg-white px-6 py-5 dark:border-zinc-800 dark:bg-zinc-900/50 sm:flex-row sm:items-center sm:justify-between">
        <div>
          <p className="font-semibold">Match each publisher to a site</p>
          <p className="text-sm text-zinc-500 dark:text-zinc-400">
            {appendable.length} new placement{appendable.length === 1 ? "" : "s"} for “{analysis.importSheetName}”
            {existing ? ` · ${existing} already present` : ""}
            {noMatch ? ` · ${noMatch} with no site match` : ""}.
          </p>
          {(analysis.startDate || analysis.endDate) && (
            <p className="mt-1 text-xs text-zinc-400 dark:text-zinc-500">
              Auto-filling Start_Date <span className="font-medium text-zinc-500 dark:text-zinc-400">{analysis.startDate || "—"}</span>,
              End_Date <span className="font-medium text-zinc-500 dark:text-zinc-400">{analysis.endDate || "—"}</span>, and Dimensions from the plan.
            </p>
          )}
        </div>
        <button
          onClick={onApply}
          className="inline-flex shrink-0 items-center gap-2 rounded-full bg-indigo-600 px-5 py-2.5 text-sm font-medium text-white shadow-sm transition-colors hover:bg-indigo-700"
        >
          Apply & build file
          <ArrowRightIcon className="h-4 w-4" />
        </button>
      </div>

      <div className="overflow-hidden rounded-2xl border border-zinc-200 dark:border-zinc-800">
        <table className="min-w-full border-collapse text-left text-sm">
          <thead className="bg-zinc-50 text-xs uppercase tracking-wide text-zinc-500 dark:bg-zinc-900 dark:text-zinc-400">
            <tr>
              <th className="px-4 py-3 font-semibold">Placement</th>
              <th className="px-4 py-3 font-semibold">Publisher</th>
              <th className="px-4 py-3 font-semibold">Site &amp; Site_ID</th>
            </tr>
          </thead>
          <tbody>
            {analysis.plans.map((p) => (
              <PlanRow
                key={p.name}
                plan={p}
                value={selections[p.name] ?? -1}
                onSelect={(idx) => onSelect(p.name, idx)}
              />
            ))}
          </tbody>
        </table>
      </div>
    </section>
  );
}

function PlanRow({
  plan,
  value,
  onSelect,
}: {
  plan: PlacementPlan;
  value: number;
  onSelect: (idx: number) => void;
}) {
  const chosen = value >= 0 ? plan.candidates[value] : undefined;
  return (
    <tr className="border-t border-zinc-100 align-top dark:border-zinc-800/70">
      <td className="max-w-[28rem] px-4 py-3">
        <span className="block truncate font-mono text-xs text-zinc-700 dark:text-zinc-300" title={plan.name}>
          {plan.name}
        </span>
        {plan.action === "existing" && (
          <span className="mt-1 inline-block rounded-full bg-zinc-100 px-2 py-0.5 text-[11px] font-medium text-zinc-500 dark:bg-zinc-800 dark:text-zinc-400">
            already in sheet
          </span>
        )}
      </td>
      <td className="px-4 py-3">
        <span className="inline-flex items-center rounded-md bg-indigo-50 px-2 py-1 text-xs font-medium text-indigo-700 dark:bg-indigo-950/40 dark:text-indigo-300">
          {plan.publisher || "—"}
        </span>
      </td>
      <td className="px-4 py-3">
        {plan.action === "existing" ? (
          <span className="text-xs text-zinc-400">—</span>
        ) : plan.candidates.length === 0 ? (
          <span className="inline-flex items-center gap-1.5 rounded-md bg-red-50 px-2 py-1 text-xs font-medium text-red-600 dark:bg-red-950/30 dark:text-red-400">
            <WarnIcon className="h-3.5 w-3.5" /> No name found
          </span>
        ) : (
          <div className="flex flex-col gap-1">
            <select
              value={value}
              onChange={(e) => onSelect(Number(e.target.value))}
              className="w-full max-w-md rounded-lg border border-zinc-300 bg-white px-3 py-2 text-sm text-zinc-800 shadow-sm focus:border-indigo-500 focus:outline-none focus:ring-1 focus:ring-indigo-500 dark:border-zinc-700 dark:bg-zinc-900 dark:text-zinc-100"
            >
              <option value={-1}>— leave blank —</option>
              {plan.candidates.map((c, i) => (
                <option key={i} value={i}>
                  {c.site}
                </option>
              ))}
            </select>
            <span className="text-[11px] text-zinc-400">
              {plan.candidates.length} match{plan.candidates.length === 1 ? "" : "es"}
              {chosen ? ` · Site_ID ${chosen.siteId}` : ""}
            </span>
          </div>
        )}
      </td>
    </tr>
  );
}

function Done({
  result,
  targetName,
  onDownload,
  onReset,
}: {
  result: BuildResult;
  targetName: string;
  onDownload: () => void;
  onReset: () => void;
}) {
  const r = result.report;
  const newNames = useMemo(() => new Set(r.appended.map((a) => a.name)), [r.appended]);
  const withSite = r.appended.filter((a) => a.site).length;

  return (
    <section className="mt-8 space-y-8">
      <div className="flex flex-col gap-4 rounded-2xl border border-emerald-300 bg-gradient-to-br from-emerald-50 to-teal-50 px-6 py-5 dark:border-emerald-900/70 dark:from-emerald-950/40 dark:to-teal-950/20 sm:flex-row sm:items-center sm:justify-between">
        <div className="flex items-center gap-3">
          <div className="flex h-10 w-10 items-center justify-center rounded-full bg-emerald-500 text-white">
            <CheckIcon className="h-5 w-5" />
          </div>
          <div>
            <p className="font-semibold text-emerald-900 dark:text-emerald-200">
              {r.appended.length} placement{r.appended.length === 1 ? "" : "s"} added · {withSite} with a site
            </p>
            <p className="text-sm text-emerald-700 dark:text-emerald-300/80">
              into “{r.importSheetName}” · <span className="font-medium">{targetName}</span>
            </p>
          </div>
        </div>
        <div className="flex shrink-0 items-center gap-2">
          <button
            onClick={onDownload}
            className="inline-flex items-center gap-2 rounded-full bg-emerald-600 px-5 py-2.5 text-sm font-medium text-white shadow-sm transition-colors hover:bg-emerald-700"
          >
            <DownloadIcon className="h-4 w-4" />
            Download updated file
          </button>
          <button
            onClick={onReset}
            className="rounded-full border border-zinc-300 bg-white px-4 py-2.5 text-sm font-medium text-zinc-700 transition-colors hover:bg-zinc-100 dark:border-zinc-700 dark:bg-zinc-900 dark:text-zinc-200 dark:hover:bg-zinc-800"
          >
            Clear
          </button>
        </div>
      </div>

      <PreviewTable
        headers={r.headers}
        rows={r.rows}
        highlightCols={[
          r.columns.siteId,
          r.columns.site,
          r.columns.placement,
          r.columns.dimensions,
          r.columns.start,
          r.columns.end,
          r.columns.clicktag,
        ].filter((i) => i >= 0)}
        placementColIndex={r.columns.placement}
        isNew={(row) => newNames.has((row[r.columns.placement] ?? "").trim())}
      />
    </section>
  );
}

function PreviewTable({
  headers,
  rows,
  highlightCols,
  placementColIndex,
  isNew,
}: {
  headers: string[];
  rows: string[][];
  highlightCols: number[];
  placementColIndex: number;
  isNew: (row: string[]) => boolean;
}) {
  if (!headers.length) return null;
  const hl = new Set(highlightCols);
  return (
    <div className="overflow-hidden rounded-2xl border border-zinc-200 dark:border-zinc-800">
      <div className="flex items-center gap-3 border-b border-zinc-200 bg-zinc-50 px-4 py-2.5 text-xs font-semibold uppercase tracking-wide text-zinc-500 dark:border-zinc-800 dark:bg-zinc-900">
        Import sheet · {rows.length} row{rows.length === 1 ? "" : "s"}
        <span className="ml-auto inline-flex items-center gap-1.5 normal-case">
          <span className="h-2.5 w-2.5 rounded-sm bg-emerald-500" />
          <span className="text-zinc-500">newly added</span>
        </span>
      </div>
      <div className="overflow-x-auto">
        <table className="min-w-full border-collapse text-left text-xs">
          <thead className="bg-zinc-50 dark:bg-zinc-900/80">
            <tr>
              {headers.map((h, i) => (
                <th
                  key={i}
                  className={`whitespace-nowrap px-3 py-2.5 font-semibold ${
                    hl.has(i) ? "text-indigo-700 dark:text-indigo-400" : "text-zinc-600 dark:text-zinc-300"
                  }`}
                >
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, ri) => {
              const added = isNew(row);
              return (
                <tr
                  key={ri}
                  className={`border-t border-zinc-100 dark:border-zinc-800/70 ${
                    added ? "bg-emerald-50/70 dark:bg-emerald-950/20" : "even:bg-zinc-50/60 dark:even:bg-zinc-900/30"
                  }`}
                >
                  {row.map((cell, ci) => (
                    <td
                      key={ci}
                      title={cell}
                      className={`max-w-[24rem] truncate px-3 py-2 ${
                        hl.has(ci) ? "font-medium text-zinc-700 dark:text-zinc-200" : "text-zinc-500 dark:text-zinc-400"
                      }`}
                    >
                      {ci === placementColIndex && added ? (
                        <span className="inline-flex items-center gap-1.5">
                          <span className="h-1.5 w-1.5 rounded-full bg-emerald-500" />
                          {cell}
                        </span>
                      ) : (
                        cell
                      )}
                    </td>
                  ))}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ---------- icons ---------- */
function LayersIcon({ className }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <polygon points="12 2 2 7 12 12 22 7 12 2" />
      <polyline points="2 17 12 22 22 17" />
      <polyline points="2 12 12 17 22 12" />
    </svg>
  );
}
function ArrowLeftIcon({ className }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <line x1="19" y1="12" x2="5" y2="12" />
      <polyline points="12 19 5 12 12 5" />
    </svg>
  );
}
function ArrowRightIcon({ className }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <line x1="5" y1="12" x2="19" y2="12" />
      <polyline points="12 5 19 12 12 19" />
    </svg>
  );
}
function FileIcon({ className }: { className?: string }) {
  return (
    <svg className={className} viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
      <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
      <polyline points="14 2 14 8 20 8" />
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
function Spinner({ className }: { className?: string }) {
  return (
    <svg className={`animate-spin ${className ?? ""}`} viewBox="0 0 24 24" fill="none">
      <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" />
      <path className="opacity-90" fill="currentColor" d="M4 12a8 8 0 0 1 8-8V0C5.37 0 0 5.37 0 12h4z" />
    </svg>
  );
}
