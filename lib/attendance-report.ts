// Report totals already use two decimal places. Integer hundredths avoid
// floating-point remainders at the full-day and half-day boundaries.
export function formatWorkDays(totalHours: number): string {
  const units = Math.max(0, Math.round(totalHours * 100));
  const days = Math.floor(units / 900);
  let remainder = units % 900;
  const parts: string[] = [];
  if (days > 0) parts.push(`${days}天`);
  if (remainder >= 450) {
    parts.push('半天');
    remainder -= 450;
  }
  if (remainder > 0) parts.push(`${remainder / 100}小时`);
  return parts.join('＋') || '0天';
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
