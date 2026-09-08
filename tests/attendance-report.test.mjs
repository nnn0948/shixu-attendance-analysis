import test from 'node:test';
import assert from 'node:assert/strict';
import { formatWorkDays, groupIssuesByEmployee } from '../lib/attendance-report.ts';

test('181 hours becomes 20 days plus 1 hour', () => {
  assert.equal(formatWorkDays(181), '20天＋1小时');
});

test('work days preserve half days and sub-half-day remainder hours', () => {
  for (const [hours, expected] of [
    [0, '0天'], [1, '1小时'], [4.49, '4.49小时'], [4.5, '半天'],
    [5.5, '半天＋1小时'], [8.99, '半天＋4.49小时'], [9, '1天'],
    [13.5, '1天＋半天'], [180, '20天'], [184.5, '20天＋半天'],
    [185.75, '20天＋半天＋1.25小时'], [180.01, '20天＋0.01小时'],
  ]) assert.equal(formatWorkDays(hours), expected);
});

test('issues are grouped by employee and sorted by date then period without mutating input', () => {
  const make = (name, date, period) => ({ name, date, period, punches: '08:00', reason: '缺卡' });
  const rows = [make('张三', '2026-08-03', '下午'), make('李四', '2026-08-01', '上午'), make('张三', '2026-08-02', '上午'), make('张三', '2026-08-03', '上午')];
  const snapshot = structuredClone(rows);
  const groups = groupIssuesByEmployee(rows);
  assert.deepEqual(groups.map((group) => group.name), ['李四', '张三']);
  assert.deepEqual(groups[1].rows.map((row) => [row.date, row.period]), [
    ['2026-08-02', '上午'], ['2026-08-03', '上午'], ['2026-08-03', '下午'],
  ]);
  assert.equal(groups.flatMap((group) => group.rows).length, rows.length);
  assert.deepEqual(rows, snapshot);
  assert.deepEqual(groupIssuesByEmployee([]), []);
});
