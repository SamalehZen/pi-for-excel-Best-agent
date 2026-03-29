export interface TemplateInfo {
  id: string;
  name: string;
  category: string;
  description: string;
  previewUrl: string;
  xlsxFile: string;
  primaryColor: string;
  fontFamily: string;
  tags: string[];
}

export const TEMPLATE_CATALOG: TemplateInfo[] = [
  {
    id: "monthly-time-sheet",
    name: "Monthly Time Sheet",
    category: "Timesheet",
    description: "Track daily clock-in/out, shift hours, and overtime with weekly subtotals.",
    previewUrl: "/templates/previews/monthly-time-sheet.png",
    xlsxFile: "/templates/files/timesheet_templates_Monthly-Time-Sheet.xlsx",
    primaryColor: "#052E34",
    fontFamily: "Calibri",
    tags: ["timesheet", "hours", "attendance", "clock", "overtime", "weekly"],
  },
  {
    id: "meeting-attendance",
    name: "Meeting Attendance",
    category: "Attendance",
    description: "Track meeting attendees with roles, departments, and attendance status.",
    previewUrl: "/templates/previews/meeting-attendance.png",
    xlsxFile: "/templates/files/attendance_templates_Meeting-Attendance-Sheet.xlsx",
    primaryColor: "#4472C4",
    fontFamily: "Calibri",
    tags: ["meeting", "attendance", "roll call", "participants", "tracking"],
  },
  {
    id: "sales-contest-tracker",
    name: "Sales Contest Tracker",
    category: "Sales",
    description: "Track sales competitions with leaderboards, targets, and achievement rates.",
    previewUrl: "/templates/previews/sales-contest-tracker.png",
    xlsxFile: "/templates/files/sales_templates_Sales-Contest-Tracker.xlsx",
    primaryColor: "#2E3BF6",
    fontFamily: "Helvetica",
    tags: ["sales", "contest", "leaderboard", "competition", "tracker", "target"],
  },
  {
    id: "sales-forecast-12m",
    name: "12-Month Sales Forecast",
    category: "Sales",
    description: "Project monthly revenue, units, and growth across a 12-month horizon.",
    previewUrl: "/templates/previews/sales-forecast-12m.png",
    xlsxFile: "/templates/files/sales_templates_12-Month-Sales-Forecast-Sheet.xlsx",
    primaryColor: "#548235",
    fontFamily: "Helvetica",
    tags: ["forecast", "projection", "12 month", "annual", "revenue", "growth"],
  },
  {
    id: "daily-sales-report",
    name: "Daily Sales Report",
    category: "Sales",
    description: "Track daily item sales with unit prices, quantities, tax, and line totals.",
    previewUrl: "/templates/previews/daily-sales-report.png",
    xlsxFile: "/templates/files/sales_templates_Daily-Sales-Report-Sheet.xlsx",
    primaryColor: "#11545B",
    fontFamily: "Franklin Gothic Book",
    tags: ["daily", "report", "invoice", "receipt", "sales log", "items"],
  },
  {
    id: "sales-lead-tracking",
    name: "Sales Lead Tracking",
    category: "Sales",
    description: "Manage sales pipeline with lead status, follow-ups, and conversion tracking.",
    previewUrl: "/templates/previews/sales-lead-tracking.png",
    xlsxFile: "/templates/files/sales_templates_Sales-Lead-Tracking-Sheet.xlsx",
    primaryColor: "#2F5496",
    fontFamily: "Calibri",
    tags: ["leads", "pipeline", "CRM", "follow-up", "conversion", "prospects"],
  },
  {
    id: "employee-schedule",
    name: "Employee Schedule",
    category: "Schedule",
    description: "Plan weekly employee shifts with time slots, departments, and coverage.",
    previewUrl: "/templates/previews/employee-schedule.png",
    xlsxFile: "/templates/files/schedule_sheets_Employee-Schedule-Sheet.xlsx",
    primaryColor: "#305496",
    fontFamily: "Calibri",
    tags: ["schedule", "shifts", "roster", "employees", "weekly", "planning"],
  },
  {
    id: "resource-planning",
    name: "Resource Planning",
    category: "Planning",
    description: "Allocate team resources across projects with capacity and utilization tracking.",
    previewUrl: "/templates/previews/resource-planning.png",
    xlsxFile: "/templates/files/plan_templates_Resource-Planning-Sheet.xlsx",
    primaryColor: "#4472C4",
    fontFamily: "Calibri",
    tags: ["resource", "planning", "allocation", "capacity", "utilization", "projects"],
  },
  {
    id: "work-planner",
    name: "Work Planner",
    category: "Planning",
    description: "Organize tasks and milestones with deadlines, priorities, and progress tracking.",
    previewUrl: "/templates/previews/work-planner.png",
    xlsxFile: "/templates/files/plan_templates_Work-Planner-Sheet.xlsx",
    primaryColor: "#375623",
    fontFamily: "Calibri",
    tags: ["planner", "tasks", "milestones", "deadlines", "progress", "project"],
  },
  {
    id: "goal-tracking",
    name: "Goal Tracking",
    category: "Tracking",
    description: "Set and monitor goals with KPIs, milestones, and completion percentages.",
    previewUrl: "/templates/previews/goal-tracking.png",
    xlsxFile: "/templates/files/tracking_templates_Goal-Tracking-Sheet.xlsx",
    primaryColor: "#385723",
    fontFamily: "Calibri",
    tags: ["goals", "KPI", "OKR", "tracking", "milestones", "objectives"],
  },
];

export function findRecommendedTemplates(
  dataHints: DataAnalysisHints,
): { id: string; score: number }[] {
  const scores: { id: string; score: number }[] = [];

  for (const template of TEMPLATE_CATALOG) {
    let score = 0;
    const allKeywords = [
      ...template.tags,
      template.category.toLowerCase(),
      template.name.toLowerCase(),
    ];

    for (const hint of dataHints.keywords) {
      const h = hint.toLowerCase();
      for (const kw of allKeywords) {
        if (kw.includes(h) || h.includes(kw)) {
          score += 10;
          break;
        }
      }
    }

    if (dataHints.hasDateColumns) {
      if (template.tags.includes("timesheet") || template.tags.includes("schedule") || template.tags.includes("daily")) {
        score += 5;
      }
    }
    if (dataHints.hasNumericColumns) {
      if (template.tags.includes("sales") || template.tags.includes("forecast") || template.tags.includes("tracking")) {
        score += 3;
      }
    }
    if (dataHints.columnCount > 6 && template.tags.includes("report")) {
      score += 2;
    }

    if (score > 0) {
      scores.push({ id: template.id, score });
    }
  }

  scores.sort((a, b) => b.score - a.score);
  return scores;
}

export interface DataAnalysisHints {
  keywords: string[];
  hasDateColumns: boolean;
  hasNumericColumns: boolean;
  columnCount: number;
  rowCount: number;
  headers: string[];
}
