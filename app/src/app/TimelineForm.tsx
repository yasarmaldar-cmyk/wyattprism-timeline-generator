"use client";

import { useEffect, useMemo, useState } from "react";
import type { ProjectInputs, ReportType } from "@/lib/types";

type LibraryData = {
  reportTypes: { id: string; label: string }[];
  optional_modules: Record<string, { id: string; question: string; default: boolean; task_ids: string[] }[]>;
  rules: Record<string, { closure_offset_from_agm_days: number | null }>;
  templates: Record<string, { name: string; phases: { id: string; name: string }[]; taskCount: number }>;
};

type Preview = {
  anchors: { kick_off: string; closure: string; board_meeting: string | null; agm: string | null };
  tasks: { id: string; name: string; phase: string; responsibility: string; startDate: string; endDate: string; durationDays: number }[];
};

// Editable, in-memory representation of a task — used for Ops to tweak names,
// dates, responsibility, add custom tasks, or remove tasks that don't apply.
type EditableTask = {
  id: string;
  name: string;
  phase: string;
  startDate: string;  // yyyy-mm-dd
  endDate: string;    // yyyy-mm-dd
  responsibility: "wp" | "client" | "both";
};

function daysBetween(a: string, b: string): number {
  const start = new Date(a);
  const end = new Date(b);
  if (isNaN(start.getTime()) || isNaN(end.getTime())) return 0;
  return Math.max(1, Math.round((end.getTime() - start.getTime()) / 86400000) + 1);
}

const inputCls =
  "mt-1 block w-full rounded-md border border-zinc-300 dark:border-zinc-700 bg-white dark:bg-zinc-900 px-3 py-2 text-sm shadow-sm focus:border-blue-500 focus:outline-none focus:ring-1 focus:ring-blue-500";
const labelCls = "block text-sm font-medium text-zinc-700 dark:text-zinc-300";

export default function TimelineForm({ data }: { data: LibraryData }) {
  const [reportType, setReportType] = useState<ReportType>("annual_report");
  const [clientName, setClientName] = useState("");
  const [reportingPeriod, setReportingPeriod] = useState("2025-26");
  const [kickOffDate, setKickOffDate] = useState("");
  const [agmDate, setAgmDate] = useState("");
  const [boardMeetingDate, setBoardMeetingDate] = useState("");
  const [closureDate, setClosureDate] = useState("");
  const [spocTeam, setSpocTeam] = useState("");
  const [spocName, setSpocName] = useState("");
  const [enabledModules, setEnabledModules] = useState<Record<string, boolean>>({});
  const [preview, setPreview] = useState<Preview | null>(null);
  const [editableTasks, setEditableTasks] = useState<EditableTask[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  // Wyattprism Platform integration — set when launched from the shell
  const [wpProjectId, setWpProjectId] = useState<string | null>(null);
  const [wpProjectCode, setWpProjectCode] = useState<string | null>(null);
  const [sendResult, setSendResult] = useState<string | null>(null);

  // Read URL query params on first render — when the shell deep-links in,
  // pre-fill the form so Ops doesn't re-type anything.
  useEffect(() => {
    const sp = new URLSearchParams(window.location.search);
    const id = sp.get("wp_project_id");
    if (id) {
      setWpProjectId(id);
      setWpProjectCode(sp.get("wp_project_code"));
      const client = sp.get("wp_client_name");
      if (client) setClientName(client);
      const rt = sp.get("wp_report_type");
      if (rt) setReportType(rt as ReportType);
      const kickOff = sp.get("wp_kick_off_date");
      if (kickOff) setKickOffDate(kickOff);
      const closure = sp.get("wp_closure_date");
      if (closure) setClosureDate(closure);
      const agm = sp.get("wp_agm_date");
      if (agm) setAgmDate(agm);
      const board = sp.get("wp_board_meeting_date");
      if (board) setBoardMeetingDate(board);
      const period = sp.get("wp_reporting_period");
      if (period) setReportingPeriod(period);
    }
  }, []);

  const modules = data.optional_modules[reportType] ?? [];
  const rules = data.rules[reportType];
  const requiresAgm = rules.closure_offset_from_agm_days != null;

  // Initialize module defaults when report type changes
  useMemo(() => {
    const next: Record<string, boolean> = {};
    for (const m of modules) next[m.id] = m.default;
    setEnabledModules(next);
    setPreview(null);
  }, [reportType]); // eslint-disable-line react-hooks/exhaustive-deps

  const inputs: ProjectInputs = {
    clientName,
    reportType,
    reportingPeriod,
    kickOffDate,
    agmDate: agmDate || undefined,
    boardMeetingDate: boardMeetingDate || undefined,
    closureDate: closureDate || undefined,
    enabledModules,
    spoc: spocTeam || spocName ? { team: spocTeam, name: spocName } : undefined,
  };

  async function handlePreview() {
    setError(null);
    setBusy(true);
    try {
      const res = await fetch("/api/timeline?preview=1", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(inputs),
      });
      if (!res.ok) {
        const j = await res.json().catch(() => ({}));
        throw new Error(j.error || `HTTP ${res.status}`);
      }
      const result: Preview = await res.json();
      setPreview(result);
      // Populate the editable copy with the computed tasks so Ops can tweak.
      setEditableTasks(
        result.tasks.map((t) => ({
          id: t.id,
          name: t.name,
          phase: t.phase,
          startDate: t.startDate.slice(0, 10),
          endDate: t.endDate.slice(0, 10),
          responsibility: (t.responsibility as EditableTask["responsibility"]) || "wp",
        }))
      );
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }

  // Editable-tasks helpers ------------------------------------------------
  function updateTask(idx: number, patch: Partial<EditableTask>) {
    setEditableTasks((prev) =>
      prev ? prev.map((t, i) => (i === idx ? { ...t, ...patch } : t)) : null
    );
  }
  function deleteTask(idx: number) {
    setEditableTasks((prev) => (prev ? prev.filter((_, i) => i !== idx) : null));
  }
  function addTask() {
    const today = new Date().toISOString().slice(0, 10);
    const fallback = kickOffDate || today;
    setEditableTasks((prev) => [
      ...(prev ?? []),
      {
        id: `custom-${Date.now()}`,
        name: "New task",
        phase: "Custom",
        startDate: fallback,
        endDate: fallback,
        responsibility: "wp",
      },
    ]);
  }
  function recomputeFromInputs() {
    // Confirm before discarding edits — Ops will lose any tweaks.
    if (editableTasks && editableTasks.length > 0) {
      const ok = window.confirm(
        "This will discard all your edits and regenerate the timeline from the project inputs. Continue?"
      );
      if (!ok) return;
    }
    // Re-run preview, which overwrites editableTasks with fresh computed values.
    void handlePreview();
  }

  async function handleDownload() {
    setError(null);
    setBusy(true);
    try {
      const res = await fetch("/api/timeline", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(inputs),
      });
      if (!res.ok) {
        const j = await res.json().catch(() => ({}));
        throw new Error(j.error || `HTTP ${res.status}`);
      }
      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${clientName || "client"}_${reportType}_${reportingPeriod}.xlsx`;
      a.click();
      URL.revokeObjectURL(url);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }

  async function handleSendToWyattprism() {
    setError(null);
    setSendResult(null);
    if (!editableTasks || editableTasks.length === 0) {
      setError("Click 'Preview timeline' first to generate tasks.");
      return;
    }
    setBusy(true);
    try {
      // Send the edited tasks directly — Ops' tweaks override the auto-computed timeline.
      const payload = {
        ...inputs,
        wp_project_id: wpProjectId,
        wp_project_code: wpProjectCode,
        anchors: preview?.anchors,
        tasks: editableTasks.map((t) => ({
          name: t.name,
          phase: t.phase,
          plannedStart: t.startDate,
          plannedEnd: t.endDate,
          responsibility: t.responsibility,
        })),
      };
      const res = await fetch("/api/timeline?send=wyattprism", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const j = await res.json().catch(() => ({}));
      if (!res.ok || !j.ok) {
        throw new Error(j.error || `HTTP ${res.status}`);
      }
      setSendResult(
        `✓ Sent — ${j.taskCount ?? "?"} tasks now live on Wyattprism project ${wpProjectCode}.`
      );
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="space-y-6">
    {/* Top section: form (left) + anchors (right) */}
    <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
      <section className="rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-900 p-6 shadow-sm">
        {wpProjectId && (
          <div className="mb-4 rounded-md border border-blue-300 bg-blue-50 dark:bg-blue-950 dark:border-blue-800 p-3 text-sm">
            <strong>Linked to Wyattprism project {wpProjectCode}.</strong> When the
            timeline is ready, click <em>Send to Wyattprism</em> to push the tasks back
            to the platform as a live tracker.
          </div>
        )}
        <h2 className="text-lg font-semibold mb-4">Project details</h2>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div className="sm:col-span-2">
            <label className={labelCls}>Client name</label>
            <input className={inputCls} value={clientName} onChange={(e) => setClientName(e.target.value)} placeholder="e.g. Greaves Cotton Ltd." />
          </div>
          <div>
            <label className={labelCls}>Report type</label>
            <select className={inputCls} value={reportType} onChange={(e) => setReportType(e.target.value as ReportType)}>
              {data.reportTypes.map((r) => (
                <option key={r.id} value={r.id}>{r.label}</option>
              ))}
            </select>
            <p className="mt-1 text-xs text-zinc-500">{data.templates[reportType].taskCount} tasks · {data.templates[reportType].phases.length} phases</p>
          </div>
          <div>
            <label className={labelCls}>Reporting period</label>
            <input className={inputCls} value={reportingPeriod} onChange={(e) => setReportingPeriod(e.target.value)} placeholder="2025-26" />
          </div>

          <div>
            <label className={labelCls}>Kick-off date</label>
            <input type="date" className={inputCls} value={kickOffDate} onChange={(e) => setKickOffDate(e.target.value)} />
          </div>
          {requiresAgm ? (
            <div>
              <label className={labelCls}>AGM date {requiresAgm && <span className="text-red-500">*</span>}</label>
              <input type="date" className={inputCls} value={agmDate} onChange={(e) => setAgmDate(e.target.value)} />
              <p className="mt-1 text-xs text-zinc-500">Closure auto-set to AGM − 30 days</p>
            </div>
          ) : (
            <div>
              <label className={labelCls}>Closure / publication date <span className="text-red-500">*</span></label>
              <input type="date" className={inputCls} value={closureDate} onChange={(e) => setClosureDate(e.target.value)} />
            </div>
          )}
          <div>
            <label className={labelCls}>Board meeting date</label>
            <input type="date" className={inputCls} value={boardMeetingDate} onChange={(e) => setBoardMeetingDate(e.target.value)} />
          </div>
          {requiresAgm && (
            <div>
              <label className={labelCls}>Closure (override)</label>
              <input type="date" className={inputCls} value={closureDate} onChange={(e) => setClosureDate(e.target.value)} />
            </div>
          )}
          <div>
            <label className={labelCls}>SPOC team</label>
            <input className={inputCls} value={spocTeam} onChange={(e) => setSpocTeam(e.target.value)} placeholder="IR team" />
          </div>
          <div>
            <label className={labelCls}>SPOC name</label>
            <input className={inputCls} value={spocName} onChange={(e) => setSpocName(e.target.value)} />
          </div>
        </div>

        {modules.length > 0 && (
          <div className="mt-6">
            <h3 className="text-sm font-semibold mb-2">Customisation</h3>
            <div className="space-y-3">
              {modules.map((m) => (
                <label key={m.id} className="flex items-start gap-3 cursor-pointer">
                  <input
                    type="checkbox"
                    className="mt-0.5"
                    checked={enabledModules[m.id] ?? m.default}
                    onChange={(e) => setEnabledModules({ ...enabledModules, [m.id]: e.target.checked })}
                  />
                  <span className="text-sm text-zinc-700 dark:text-zinc-300">
                    {m.question}
                    <span className="ml-1 text-xs text-zinc-500">({m.task_ids.length} tasks)</span>
                  </span>
                </label>
              ))}
            </div>
          </div>
        )}

        <div className="mt-6 flex gap-3 flex-wrap">
          <button onClick={handlePreview} disabled={busy} className="rounded-md border border-zinc-300 dark:border-zinc-700 px-4 py-2 text-sm font-medium hover:bg-zinc-100 dark:hover:bg-zinc-800 disabled:opacity-50">
            {busy ? "Working…" : "Preview timeline"}
          </button>
          <button onClick={handleDownload} disabled={busy} className="rounded-md border border-zinc-300 dark:border-zinc-700 px-4 py-2 text-sm font-medium hover:bg-zinc-100 dark:hover:bg-zinc-800 disabled:opacity-50">
            {busy ? "Working…" : "Download Excel"}
          </button>
          {wpProjectId && (
            <button onClick={handleSendToWyattprism} disabled={busy} className="rounded-md bg-blue-600 text-white px-4 py-2 text-sm font-medium hover:bg-blue-700 disabled:opacity-50">
              {busy ? "Sending…" : "Send to Wyattprism"}
            </button>
          )}
        </div>
        {editableTasks && editableTasks.length > 0 && (
          <p className="mt-2 text-[11px] text-zinc-500 italic">
            Note: <strong>Send to Wyattprism</strong> uses your edited tasks. The Excel download
            still uses the auto-computed timeline.
          </p>
        )}
        {error && <p className="mt-3 text-sm text-red-600">{error}</p>}
        {sendResult && <p className="mt-3 text-sm text-green-700">{sendResult}</p>}
      </section>

      {/* Right column: anchors only (compact summary) */}
      <section className="rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-900 p-6 shadow-sm">
        <h2 className="text-lg font-semibold mb-4">Key dates</h2>
        {!preview ? (
          <p className="text-sm text-zinc-500">Click &ldquo;Preview timeline&rdquo; to compute the kick-off, closure and any other anchor dates.</p>
        ) : (
          <div className="grid grid-cols-2 gap-2 text-xs">
            <KV label="Kick-off" value={preview.anchors.kick_off} />
            <KV label="Closure" value={preview.anchors.closure} />
            {preview.anchors.board_meeting && <KV label="Board Meeting" value={preview.anchors.board_meeting} />}
            {preview.anchors.agm && <KV label="AGM" value={preview.anchors.agm} />}
          </div>
        )}
      </section>
    </div>

    {/* Tasks — full width below the form so the task names have room to breathe */}
    {editableTasks && editableTasks.length > 0 && (
      <section className="rounded-lg border border-zinc-200 dark:border-zinc-800 bg-white dark:bg-zinc-900 p-6 shadow-sm">
        <div className="flex items-baseline justify-between mb-4 gap-3 flex-wrap">
          <div>
            <h2 className="text-lg font-semibold">Tasks</h2>
            <p className="text-xs text-zinc-500 mt-0.5">
              {editableTasks.length} task{editableTasks.length === 1 ? "" : "s"} · click any field to edit · changes go to Wyattprism when you click Send
            </p>
          </div>
          <div className="flex items-center gap-2">
            <button
              onClick={recomputeFromInputs}
              disabled={busy}
              className="rounded-md border border-zinc-300 dark:border-zinc-700 px-2.5 py-1 text-xs font-medium hover:bg-zinc-100 dark:hover:bg-zinc-800 disabled:opacity-50"
              title="Discard edits and recompute from project inputs"
            >
              ↺ Recompute
            </button>
            <button
              onClick={addTask}
              className="rounded-md border border-zinc-300 dark:border-zinc-700 px-2.5 py-1 text-xs font-medium hover:bg-zinc-100 dark:hover:bg-zinc-800"
            >
              + Add task
            </button>
          </div>
        </div>

        <div className="overflow-y-auto max-h-[70vh] border border-zinc-200 dark:border-zinc-800 rounded-md">
          <table className="w-full text-sm" style={{ tableLayout: "fixed" }}>
            <colgroup>
              <col style={{ width: "13%" }} />
              <col />
              <col style={{ width: "100px" }} />
              <col style={{ width: "150px" }} />
              <col style={{ width: "150px" }} />
              <col style={{ width: "70px" }} />
              <col style={{ width: "40px" }} />
            </colgroup>
            <thead className="bg-zinc-100 dark:bg-zinc-800 sticky top-0 z-10">
              <tr>
                <th className="text-left px-3 py-2 font-medium text-xs uppercase tracking-wider text-zinc-500">Phase</th>
                <th className="text-left px-3 py-2 font-medium text-xs uppercase tracking-wider text-zinc-500">Task</th>
                <th className="text-center px-3 py-2 font-medium text-xs uppercase tracking-wider text-zinc-500">By</th>
                <th className="text-left px-3 py-2 font-medium text-xs uppercase tracking-wider text-zinc-500">Start</th>
                <th className="text-left px-3 py-2 font-medium text-xs uppercase tracking-wider text-zinc-500">End</th>
                <th className="text-center px-3 py-2 font-medium text-xs uppercase tracking-wider text-zinc-500">Days</th>
                <th className="px-1 py-2"></th>
              </tr>
            </thead>
            <tbody>
              {editableTasks.map((t, i) => (
                <tr key={t.id} className="border-t border-zinc-200 dark:border-zinc-800 hover:bg-zinc-50 dark:hover:bg-zinc-800/40">
                  <td className="px-2 py-1.5">
                    <input
                      className="w-full bg-transparent border border-transparent hover:border-zinc-300 focus:border-blue-500 rounded px-2 py-1 text-zinc-600 dark:text-zinc-400 text-xs focus:outline-none focus:bg-white dark:focus:bg-zinc-800"
                      value={t.phase}
                      onChange={(e) => updateTask(i, { phase: e.target.value })}
                    />
                  </td>
                  <td className="px-2 py-1.5">
                    <input
                      className="w-full bg-transparent border border-transparent hover:border-zinc-300 focus:border-blue-500 rounded px-2 py-1 text-sm focus:outline-none focus:bg-white dark:focus:bg-zinc-800"
                      value={t.name}
                      onChange={(e) => updateTask(i, { name: e.target.value })}
                    />
                  </td>
                  <td className="px-2 py-1.5 text-center">
                    <select
                      className="bg-transparent border border-transparent hover:border-zinc-300 focus:border-blue-500 rounded px-1.5 py-1 text-xs focus:outline-none focus:bg-white dark:focus:bg-zinc-800"
                      value={t.responsibility}
                      onChange={(e) => updateTask(i, { responsibility: e.target.value as EditableTask["responsibility"] })}
                    >
                      <option value="wp">WP</option>
                      <option value="client">Client</option>
                      <option value="both">Both</option>
                    </select>
                  </td>
                  <td className="px-2 py-1.5">
                    <input
                      type="date"
                      className="w-full bg-transparent border border-transparent hover:border-zinc-300 focus:border-blue-500 rounded px-2 py-1 text-xs focus:outline-none focus:bg-white dark:focus:bg-zinc-800"
                      value={t.startDate}
                      onChange={(e) => updateTask(i, { startDate: e.target.value })}
                    />
                  </td>
                  <td className="px-2 py-1.5">
                    <input
                      type="date"
                      className="w-full bg-transparent border border-transparent hover:border-zinc-300 focus:border-blue-500 rounded px-2 py-1 text-xs focus:outline-none focus:bg-white dark:focus:bg-zinc-800"
                      value={t.endDate}
                      onChange={(e) => updateTask(i, { endDate: e.target.value })}
                    />
                  </td>
                  <td className="px-2 py-1.5 text-center text-zinc-500 text-xs">
                    {daysBetween(t.startDate, t.endDate)}
                  </td>
                  <td className="px-1 py-1.5 text-center">
                    <button
                      onClick={() => deleteTask(i)}
                      title="Remove this task"
                      className="text-zinc-400 hover:text-red-600 text-lg leading-none cursor-pointer"
                    >
                      ×
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    )}
    </div>
  );
}

function KV({ label, value }: { label: string; value: string }) {
  return (
    <div className="rounded border border-zinc-200 dark:border-zinc-800 px-2 py-1.5">
      <div className="text-zinc-500">{label}</div>
      <div className="font-medium">{new Date(value).toLocaleDateString("en-GB")}</div>
    </div>
  );
}

