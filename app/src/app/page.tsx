import { taskLibrary, REPORT_TYPE_LABELS } from "@/lib/library";
import TimelineForm from "./TimelineForm";

export default function Home() {
  const reportTypes = Object.keys(taskLibrary.templates) as Array<keyof typeof taskLibrary.templates>;
  const initialData = {
    reportTypes: reportTypes.map((id) => ({ id, label: REPORT_TYPE_LABELS[id] })),
    optional_modules: taskLibrary.optional_modules,
    rules: taskLibrary.rules,
    templates: Object.fromEntries(
      reportTypes.map((id) => [
        id,
        {
          name: taskLibrary.templates[id].name,
          phases: taskLibrary.templates[id].phases,
          taskCount: taskLibrary.templates[id].tasks.length,
        },
      ])
    ),
  };

  return (
    <main className="min-h-screen bg-zinc-50 dark:bg-zinc-950">
      <div className="mx-auto max-w-5xl px-6 py-10">
        <header className="mb-8">
          <h1 className="text-3xl font-semibold tracking-tight">Timeline Generator</h1>
          <p className="mt-1 text-sm text-zinc-600 dark:text-zinc-400">
            Generate consistent project timelines for AR, IR, Sustainability and ESG reports.
          </p>
        </header>
        <TimelineForm data={initialData} />
      </div>
    </main>
  );
}
