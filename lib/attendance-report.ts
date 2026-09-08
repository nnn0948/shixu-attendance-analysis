// Report totals already use two decimal places. Integer hundredths avoid
// floating-point remainders at the full-day and half-day boundaries.
export function formatWorkDays(totalHours: number, standardHours = 9): string {
  if (!Number.isFinite(standardHours) || standardHours <= 0) return '待确认';
  const units = Math.max(0, Math.round(totalHours * 100));
  const dayUnits = Math.round(standardHours * 100);
  const halfDayUnits = dayUnits / 2;
  const days = Math.floor(units / dayUnits);
  let remainder = units % dayUnits;
  const parts: string[] = [];
  if (days > 0) parts.push(`${days}天`);
  if (remainder >= halfDayUnits) {
    parts.push('半天');
    remainder -= halfDayUnits;
  }
  if (remainder > 0) parts.push(`${remainder / 100}小时`);
  return parts.join('＋') || '0天';
}

export type ReportDay = {
  name: string; date: string; morning: number; afternoon: number;
  total: number; type: 'ok' | 'half' | 'warn'; issues: string[];
};
export type MonthlyStandard = {
  month: string; hours: number | null; employees: number;
  supportingEmployees: number; sampleDays: number; explanation: string;
};
export type SummaryRow = {
  name: string; month: string; days: number; total: number;
  exceptionDates: string[]; standard: MonthlyStandard;
};

function uniqueMode(counts: Map<number, number>): number | null {
  const sorted = [...counts].sort((a, b) => b[1] - a[1]);
  return sorted.length && (!sorted[1] || sorted[0][1] > sorted[1][1]) ? sorted[0][0] : null;
}

export function inferMonthlyStandards(rows: ReportDay[]): Map<string, MonthlyStandard> {
  const months = new Map<string, ReportDay[]>();
  for (const row of rows) {
    const month = row.date.slice(0, 7);
    const days = months.get(month) ?? [];
    days.push(row);
    months.set(month, days);
  }
  const standards = new Map<string, MonthlyStandard>();
  for (const [month, days] of months) {
    const samples = days.filter((row) => row.type === 'ok' && !row.issues.length && row.morning > 0 && row.afternoon > 0 && Number.isFinite(row.total) && row.total > 0);
    const employees = new Map<string, Map<number, number>>();
    for (const row of samples) {
      // Small clock differences vote for the nearest half-hour standard.
      const hours = Math.round(row.total * 2) / 2;
      if (hours <= 0) continue;
      const counts = employees.get(row.name) ?? new Map<number, number>();
      counts.set(hours, (counts.get(hours) ?? 0) + 1);
      employees.set(row.name, counts);
    }
    // One vote per employee prevents employees with more recorded days from
    // outweighing the majority of staff. Tied employee modes abstain.
    const votes = new Map<number, number>();
    for (const counts of employees.values()) {
      const mode = uniqueMode(counts);
      if (mode !== null) votes.set(mode, (votes.get(mode) ?? 0) + 1);
    }
    const candidate = uniqueMode(votes);
    const support = candidate === null ? 0 : votes.get(candidate)!;
    const confirmed = employees.size >= 2 && support > employees.size / 2;
    const hours = confirmed ? candidate : null;
    const explanation = hours !== null
      ? `自动识别：${support}/${employees.size}名员工的常见完整工作日为${hours}小时（${samples.length}个有效工作日样本）`
      : employees.size < 2
        ? '有效员工样本不足，暂不折算工作天数，仅保留总小时数'
        : '员工常见时长并列或未过半，无法确定本月标准，暂不折算工作天数';
    standards.set(month, { month, hours, employees: employees.size, supportingEmployees: support, sampleDays: samples.length, explanation });
  }
  return standards;
}

export function summarizeMonthly(rows: ReportDay[]): SummaryRow[] {
  const standards = inferMonthlyStandards(rows);
  const summaries = new Map<string, SummaryRow>();
  for (const row of rows) {
    const month = row.date.slice(0, 7);
    const key = `${month}\u0000${row.name}`;
    const summary = summaries.get(key) ?? { name: row.name, month, days: 0, total: 0, exceptionDates: [], standard: standards.get(month)! };
    if (row.total > 0) summary.days += 1;
    summary.total = Math.round((summary.total + row.total) * 100) / 100;
    if (row.issues.length || row.type === 'warn') summary.exceptionDates.push(row.date);
    summaries.set(key, summary);
  }
  return [...summaries.values()]
    .map((row) => ({ ...row, exceptionDates: [...new Set(row.exceptionDates)] }))
    .sort((a, b) => a.month.localeCompare(b.month) || a.name.localeCompare(b.name, 'zh-CN'));
}

export function monthlyWorkDays(row: SummaryRow): string {
  return row.standard.hours === null ? `待确认（${row.total}小时）` : formatWorkDays(row.total, row.standard.hours);
}

export type IssueRow = { name: string; date: string; period: string; punches: string; reason: string };

export function groupIssuesByEmployee(rows: IssueRow[]) {
  const employees = new Map<string, IssueRow[]>();
  for (const row of rows) {
    const group = employees.get(row.name) ?? [];
    group.push(row);
    employees.set(row.name, group);
  }
  const periods: Record<string, number> = { 全天: 0, 上午: 1, 下午: 2 };
  return [...employees.entries()]
    .sort(([a], [b]) => a.localeCompare(b, 'zh-CN'))
    .map(([name, issues]) => ({
      name,
      rows: issues.sort((a, b) => a.date.localeCompare(b.date) || (periods[a.period] ?? 3) - (periods[b.period] ?? 3)),
    }));
}
