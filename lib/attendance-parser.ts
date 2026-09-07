export type PunchGroup = { name: string; date: string; times: number[] };
export type AttendanceSheetData = {
  matrix: unknown[][];
  records: Record<string, unknown>[];
};

export function parseAttendanceSheets(sheets: AttendanceSheetData[]) {
  const parsed: PunchGroup[] = [];
  let rowCount = 0;

  for (const { matrix, records } of sheets) {
    rowCount += matrix.length;
    parsed.push(...parseDeviceReport(matrix), ...parseRecordRows(records));
  }

  return { groups: mergePunchGroups(parsed), rowCount };
}

function parseDeviceReport(rows: unknown[][]): PunchGroup[] {
  const groups: PunchGroup[] = [];
  let period: { year: number; month: number } | null = null;
  let dayColumns: Array<{ column: number; day: number }> = [];

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = rows[rowIndex] ?? [];
    period = findPeriod(row) ?? period;

    const possibleDays = findDayColumns(row);
    if (possibleDays.length >= 7) {
      dayColumns = possibleDays;
      continue;
    }

    if (!period || !dayColumns.length) continue;
    const name = findLabeledValue(row, /^(姓名|员工姓名)[:：]?$/);
    if (!name) continue;

    const punchRow = findPunchRow(rows, rowIndex + 1, dayColumns);
    if (!punchRow) continue;

    for (const { column, day } of dayColumns) {
      if (!isValidCalendarDay(period.year, period.month, day)) continue;
      const times = extractTimes(punchRow[column]);
      if (!times.length) continue;
      groups.push({
        name,
        date: `${period.year}-${String(period.month).padStart(2, '0')}-${String(day).padStart(2, '0')}`,
        times,
      });
    }
  }

  return groups;
}

function findPeriod(row: unknown[]) {
  for (const value of row) {
    const text = cellText(value).trim();
    const match = text.match(/(20\d{2})[年/-](\d{1,2})[月/-](\d{1,2})日?\s*(?:~|～|至|—|–|-)/);
    if (!match) continue;
    const year = Number(match[1]);
    const month = Number(match[2]);
    if (month >= 1 && month <= 12) return { year, month };
  }
  return null;
}

function findDayColumns(row: unknown[]) {
  const values = row.map((value, column) => {
    const text = cellText(value).trim();
    return /^\d{1,2}$/.test(text) ? { column, day: Number(text) } : null;
  }).filter((item): item is { column: number; day: number } => Boolean(item && item.day >= 1 && item.day <= 31));

  if (values.length < 7) return [];
  const ordered = values.every((item, index) => index === 0 || item.day === values[index - 1].day + 1);
  return ordered ? values : [];
}

function findLabeledValue(row: unknown[], labelPattern: RegExp) {
  for (let column = 0; column < row.length; column += 1) {
    const raw = cellText(row[column]).trim();
    const compact = raw.replace(/\s+/g, '');
    if (!labelPattern.test(compact)) continue;

    const inline = compact.replace(labelPattern, '').trim();
    if (inline) return inline;
    for (let next = column + 1; next <= Math.min(row.length - 1, column + 4); next += 1) {
      const candidate = cleanName(row[next]);
      if (candidate) return candidate;
    }
  }
  return '';
}

function findPunchRow(rows: unknown[][], startIndex: number, dayColumns: Array<{ column: number; day: number }>) {
  for (let index = startIndex; index < Math.min(rows.length, startIndex + 3); index += 1) {
    const row = rows[index] ?? [];
    if (findLabeledValue(row, /^(姓名|员工姓名)[:：]?$/)) return null;
    const punchCount = dayColumns.reduce((sum, { column }) => sum + extractTimes(row[column]).length, 0);
    if (punchCount > 0) return row;
  }
  return null;
}

function parseRecordRows(rows: Record<string, unknown>[]): PunchGroup[] {
  const groups: PunchGroup[] = [];
  let lastName = '';
  let lastDate = '';

  for (const row of rows) {
    const entries = Object.entries(row);
    const nameEntry = entries.find(([key]) => /^(姓名|员工姓名|员工|名字|name|employee)$/i.test(normalizeKey(key)));
    const dateEntry = entries.find(([key]) => /^(日期|考勤日期|打卡日期|签到日期|date|day)$/i.test(normalizeKey(key)));
    const currentName = cleanName(nameEntry?.[1]) || lastName;
    const explicitDate = parseDate(dateEntry?.[1]) || '';
    if (currentName) lastName = currentName;
    if (explicitDate) lastDate = explicitDate;
    if (!currentName) continue;

    const byDate = new Map<string, number[]>();
    for (const [key, value] of entries) {
      if (key === nameEntry?.[0]) continue;
      const text = cellText(value).trim();
      if (!text) continue;
      const date = parseDate(value) || explicitDate || lastDate;
      const times = extractTimes(value);
      const isPunchField = /(打卡|签到|签退|时间|上班|下班|上午|下午|punch|clock|check|time)/i.test(key);
      if (date && times.length && (isPunchField || key === dateEntry?.[0] || /\d{1,2}:\d{2}/.test(text))) {
        byDate.set(date, [...(byDate.get(date) ?? []), ...times]);
      }
    }
    for (const [date, times] of byDate) {
      if (times.length) groups.push({ name: currentName, date, times });
    }
  }
  return groups;
}

function mergePunchGroups(groups: PunchGroup[]) {
  const map = new Map<string, PunchGroup>();
  for (const group of groups) {
    if (!group.name || !group.date || !group.times.length) continue;
    const key = `${group.name}\u0000${group.date}`;
    const item = map.get(key) ?? { name: group.name, date: group.date, times: [] };
    item.times.push(...group.times);
    item.times = [...new Set(item.times)].sort((a, b) => a - b);
    map.set(key, item);
  }
  return [...map.values()];
}

export function extractTimes(value: unknown): number[] {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return [value.getHours() * 60 + value.getMinutes()];
  }
  const text = cellText(value);
  const result: number[] = [];
  const regex = /([01]?\d|2[0-3]):([0-5]\d)/g;
  let match: RegExpExecArray | null;
  while ((match = regex.exec(text))) result.push(Number(match[1]) * 60 + Number(match[2]));
  return result;
}

function parseDate(value: unknown): string | null {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return localDate(value);
  const text = cellText(value).trim();
  if (!text) return null;
  const match = text.match(/(20\d{2})[年/-](\d{1,2})[月/-](\d{1,2})日?/);
  if (match) return `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : localDate(parsed);
}

function isValidCalendarDay(year: number, month: number, day: number) {
  const date = new Date(year, month - 1, day);
  return date.getFullYear() === year && date.getMonth() === month - 1 && date.getDate() === day;
}

function localDate(date: Date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function cleanName(value: unknown) {
  const text = cellText(value).trim();
  return text && !/姓名|员工姓名/i.test(text) ? text : '';
}

function cellText(value: unknown) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? '' : value.toISOString();
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean' || typeof value === 'bigint') return String(value);
  return '';
}

function normalizeKey(value: string) {
  return value.replace(/[\s_\-（）()]/g, '').toLowerCase();
}
