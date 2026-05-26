import { addDays, format } from "date-fns";

// Parse "YYYY-MM-DD" as UTC midnight so ExcelJS writes a clean integer Excel serial.
// date-fns parseISO treats date-only strings as local time, which causes a TZ shift in Excel.
function parseISO(s: string): Date {
  const [y, m, d] = s.split("-").map(Number);
  return new Date(Date.UTC(y, m - 1, d));
}
import { getOptionalModules, getRules, getTemplate } from "./library";
import type {
  AnchorDates,
  ComputedTask,
  ProjectInputs,
  TaskDef,
} from "./types";

export class EngineError extends Error {}

export function resolveAnchors(inputs: ProjectInputs): AnchorDates {
  const kick = parseISO(inputs.kickOffDate);
  const agm = inputs.agmDate ? parseISO(inputs.agmDate) : undefined;
  const board = inputs.boardMeetingDate ? parseISO(inputs.boardMeetingDate) : undefined;
  const rules = getRules(inputs.reportType);

  let closure: Date | undefined;
  if (inputs.closureDate) {
    closure = parseISO(inputs.closureDate);
  } else if (rules.closure_offset_from_agm_days != null && agm) {
    closure = addDays(agm, rules.closure_offset_from_agm_days);
  }

  if (!closure) {
    throw new EngineError(
      "Closure date could not be determined. Provide a Closure date, or for AR/IR provide an AGM date."
    );
  }

  return { kick_off: kick, board_meeting: board, agm, closure };
}

export function filterTasksByModules(tasks: TaskDef[], inputs: ProjectInputs): TaskDef[] {
  const modules = getOptionalModules(inputs.reportType);
  const disabledTaskIds = new Set<string>();
  for (const m of modules) {
    const enabled = inputs.enabledModules[m.id] ?? m.default;
    if (!enabled) m.task_ids.forEach((id) => disabledTaskIds.add(id));
  }
  return tasks.filter((t) => !disabledTaskIds.has(t.id));
}

function anchorDate(anchors: AnchorDates, anchor: TaskDef["anchor"]): Date {
  switch (anchor) {
    case "kick_off":
      return anchors.kick_off;
    case "closure":
      return anchors.closure;
    case "agm":
      if (!anchors.agm) throw new EngineError("Task is anchored to AGM but no AGM date was provided.");
      return anchors.agm;
    case "board_meeting":
      if (!anchors.board_meeting)
        throw new EngineError("Task is anchored to Board Meeting but no Board Meeting date was provided.");
      return anchors.board_meeting;
  }
}

export function computeTimeline(inputs: ProjectInputs): {
  anchors: AnchorDates;
  tasks: ComputedTask[];
} {
  const anchors = resolveAnchors(inputs);
  const template = getTemplate(inputs.reportType);
  const tasks = filterTasksByModules(template.tasks, inputs);

  const computed: ComputedTask[] = tasks.map((t) => {
    const base = anchorDate(anchors, t.anchor);
    // Delivery-anchored: offset_days = end date; duration_days = work span backwards (default 1 = milestone)
    const endDate = addDays(base, t.offset_days);
    const dur = Math.max(1, t.duration_days ?? 1);
    const startDate = addDays(endDate, -(dur - 1));
    const dateLabel =
      dur > 1
        ? `${format(startDate, "dd-MM-yyyy")} to ${format(endDate, "dd-MM-yyyy")}`
        : format(endDate, "dd-MM-yyyy");
    return { ...t, startDate, endDate, dateLabel };
  });

  computed.sort((a, b) => a.startDate.getTime() - b.startDate.getTime());
  return { anchors, tasks: computed };
}

export function groupByMonth(tasks: ComputedTask[]): Map<string, ComputedTask[]> {
  const map = new Map<string, ComputedTask[]>();
  for (const t of tasks) {
    const key = format(t.startDate, "MMMM yyyy");
    if (!map.has(key)) map.set(key, []);
    map.get(key)!.push(t);
  }
  return map;
}
