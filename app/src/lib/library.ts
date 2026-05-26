import library from "../../data/task-library.json";
import type { TaskLibrary, ReportType } from "./types";

export const taskLibrary = library as unknown as TaskLibrary;

export function getTemplate(type: ReportType) {
  return taskLibrary.templates[type];
}

export function getOptionalModules(type: ReportType) {
  return taskLibrary.optional_modules[type] ?? [];
}

export function getRules(type: ReportType) {
  return taskLibrary.rules[type];
}

export const REPORT_TYPE_LABELS: Record<ReportType, string> = {
  annual_report: "Annual Report",
  integrated_report: "Integrated Report",
  sustainability_report: "Sustainability Report",
  esg_report: "ESG / BRSR Report",
};
