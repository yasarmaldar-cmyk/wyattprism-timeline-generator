export type ReportType = "annual_report" | "integrated_report" | "sustainability_report" | "esg_report";
export type Anchor = "kick_off" | "board_meeting" | "agm" | "closure";
export type Responsibility = "client" | "wp" | "both";

export interface TaskDef {
  id: string;
  name: string;
  phase: string;
  responsibility: Responsibility;
  anchor: Anchor;
  offset_days: number;
  duration_days?: number;
  depends_on?: string[];
  notes?: string;
}

export interface PhaseDef {
  id: string;
  name: string;
  order: number;
}

export interface OptionalModule {
  id: string;
  question: string;
  default: boolean;
  task_ids: string[];
}

export interface Template {
  name: string;
  phases: PhaseDef[];
  tasks: TaskDef[];
}

export interface TaskLibrary {
  version: string;
  anchors: Record<Anchor, string>;
  rules: Record<ReportType, { closure_offset_from_agm_days: number | null }>;
  optional_modules: Record<ReportType, OptionalModule[]>;
  templates: Record<ReportType, Template>;
}

export interface AnchorDates {
  kick_off: Date;
  board_meeting?: Date;
  agm?: Date;
  closure: Date;
}

export interface ProjectInputs {
  clientName: string;
  reportType: ReportType;
  reportingPeriod: string; // e.g. "2025-26"
  kickOffDate: string;     // ISO yyyy-mm-dd
  agmDate?: string;
  boardMeetingDate?: string;
  closureDate?: string;    // override; if absent and AR/IR, derived from AGM
  enabledModules: Record<string, boolean>;
  spoc?: { team: string; name: string };
}

export interface ComputedTask extends TaskDef {
  startDate: Date;
  endDate: Date;
  dateLabel: string;
}
