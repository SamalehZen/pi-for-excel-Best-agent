export { MONTHLY_TIME_SHEET } from "./monthly-time-sheet.js";
export { MEETING_ATTENDANCE } from "./meeting-attendance.js";
export { SALES_BALANCED_SCORECARD } from "./sales-balanced-scorecard.js";
export { SALES_FORECAST_12M } from "./sales-forecast-12m.js";
export { SALES_CONTEST_TRACKER } from "./sales-contest-tracker.js";
export { DAILY_SALES_REPORT } from "./daily-sales-report.js";

import { MONTHLY_TIME_SHEET } from "./monthly-time-sheet.js";
import { MEETING_ATTENDANCE } from "./meeting-attendance.js";
import { SALES_BALANCED_SCORECARD } from "./sales-balanced-scorecard.js";
import { SALES_FORECAST_12M } from "./sales-forecast-12m.js";
import { SALES_CONTEST_TRACKER } from "./sales-contest-tracker.js";
import { DAILY_SALES_REPORT } from "./daily-sales-report.js";

export const BUNDLED_TEMPLATE_DEFINITIONS = [
  MONTHLY_TIME_SHEET,
  MEETING_ATTENDANCE,
  SALES_BALANCED_SCORECARD,
  SALES_FORECAST_12M,
  SALES_CONTEST_TRACKER,
  DAILY_SALES_REPORT,
] as const;
