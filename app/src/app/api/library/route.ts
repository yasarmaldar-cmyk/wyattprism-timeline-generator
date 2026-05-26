import { NextResponse } from "next/server";
import { taskLibrary } from "@/lib/library";

export const runtime = "nodejs";

export async function GET() {
  return NextResponse.json({
    reportTypes: Object.keys(taskLibrary.templates),
    optional_modules: taskLibrary.optional_modules,
    rules: taskLibrary.rules,
    templates: Object.fromEntries(
      Object.entries(taskLibrary.templates).map(([k, v]) => [
        k,
        { name: v.name, phases: v.phases, taskCount: v.tasks.length },
      ])
    ),
  });
}
