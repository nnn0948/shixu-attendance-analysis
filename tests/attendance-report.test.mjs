import test from 'node:test';
import assert from 'node:assert/strict';
import { formatWorkDays, groupIssuesByEmployee, inferMonthlyStandards, summarizeMonthly, monthlyWorkDays } from '../lib/attendance-report.ts';

test('181 hours becomes 20 days plus 1 hour', () => {
  assert.equal(formatWorkDays(181), '20天＋1小时');
});

const day = (name, date, total, extra = {}) => ({ name, date, morning: total / 2, afternoon: total / 2, total, type: 'ok', issues: [], ...extra });

test('automatically identifies 8, 8.5 and 9-hour standards independently per month', () => {
  const rows = [8, 8.5, 9].flatMap((hours, index) => ['甲', '乙', '丙'].map((name) => day(name, `2026-0${index + 6}-01`, hours)));
  const standards = inferMonthlyStandards(rows);
  assert.deepEqual([...standards.values()].map((row) => row.hours), [8, 8.5, 9]);
});

test('one employee with many overtime days cannot outweigh most employees', () => {
  const rows = [day('甲', '2026-08-01', 8), day('乙', '2026-08-01', 8),
    ...Array.from({ length: 20 }, (_, i) => day('丙', `2026-08-${String(i + 1).padStart(2, '0')}`, 10))];
  assert.equal(inferMonthlyStandards(rows).get('2026-08').hours, 8);
});

test('occasional overtime, half days, and abnormal days do not set the standard', () => {
  const rows = ['甲', '乙'].flatMap((name) => [
    day(name, '2026-08-01', 8.5), day(name, '2026-08-02', 8.5), day(name, '2026-08-03', 10),
    day(name, '2026-08-04', 4, { morning: 4, afternoon: 0, type: 'half' }),
    day(name, '2026-08-05', 12, { type: 'warn', issues: ['缺卡'] }),
  ]);
  const standard = inferMonthlyStandards(rows).get('2026-08');
  assert.equal(standard.hours, 8.5);
  assert.equal(standard.sampleDays, 6);
});

test('minor clock differences cluster around a half-hour standard', () => {
  assert.equal(inferMonthlyStandards([day('甲', '2026-08-01', 8.42), day('乙', '2026-08-01', 8.53)]).get('2026-08').hours, 8.5);
});

test('ties and insufficient data stay unconfirmed instead of silently defaulting to 9', () => {
  for (const rows of [
    [day('甲', '2026-08-01', 8)],
    [day('甲', '2026-08-01', 8), day('乙', '2026-08-01', 9)],
    [day('甲', '2026-08-01', 4, { afternoon: 0, type: 'half' })],
    [day('甲', '2026-08-01', 8), day('甲', '2026-08-02', 9), day('乙', '2026-08-01', 8)],
  ]) assert.equal(inferMonthlyStandards(rows).get('2026-08').hours, null);
  assert.equal(inferMonthlyStandards([]).size, 0);
});

test('work day formatting uses the inferred full day and half day without dropping residual hours', () => {
  assert.equal(formatWorkDays(181, 8), '22天＋半天＋1小时');
  assert.equal(formatWorkDays(181, 8.5), '21天＋2.5小时');
  assert.equal(formatWorkDays(181, 9), '20天＋1小时');
  assert.equal(formatWorkDays(13.75, 8.5), '1天＋半天＋1小时');
});

test('monthly reports separate the same employee across months and preserve attendance days', () => {
  const rows = [day('甲', '2026-08-01', 8), day('乙', '2026-08-01', 8),
    day('甲', '2026-09-01', 9), day('乙', '2026-09-01', 9),
    day('甲', '2026-09-02', 1, { morning: 1, afternoon: 0, type: 'half' })];
  const summaries = summarizeMonthly(rows);
  const august = summaries.find((row) => row.name === '甲' && row.month === '2026-08');
  const september = summaries.find((row) => row.name === '甲' && row.month === '2026-09');
  assert.equal(summaries.length, 4);
  assert.equal(august.total, 8);
  assert.equal(august.standard.hours, 8);
  assert.equal(september.total, 10);
  assert.equal(september.standard.hours, 9);
  assert.equal(september.days, 2);
  assert.equal(monthlyWorkDays(september), '1天＋1小时');
  assert.match(monthlyWorkDays(summarizeMonthly([day('甲', '2026-08-01', 8)])[0]), /待确认.*8小时/);
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
