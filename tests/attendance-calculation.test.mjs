import test from 'node:test';
import assert from 'node:assert/strict';
import { analyzePeriods, analyzeSegment, countedMinutes, DEFAULT_BOUNDARY } from '../lib/attendance-calculation.ts';

const times = (...values) => values.map((value) => {
  const [hours, minutes] = value.split(':').map(Number);
  return hours * 60 + minutes;
});
const segment = (...values) => analyzeSegment(times(...values), '上午', 10);

test('06:50–11:30 counts as 4.5 hours', () => {
  assert.equal(segment('06:50', '11:30').hours, 4.5);
});

test('allowance includes exactly 15 minutes, but retains 16–29 minutes', () => {
  for (const [raw, counted] of [[270, 270], [271, 270], [284, 270], [285, 270], [286, 286], [299, 299], [300, 300], [314, 300]]) {
    assert.equal(countedMinutes(raw), counted);
  }
});

test('user multiple-punch example sums 145 + 109 = 254 minutes, then counts 240', () => {
  const result = segment('06:55', '09:20', '09:40', '11:29');
  assert.equal(result.hours, 4);
  assert.equal(result.issue, undefined);
});

test('round once after accumulating, never round each interval', () => {
  assert.equal(segment('07:00', '09:10', '09:20', '11:30').hours, 260 / 60);
});

test('short break is retained and not mistaken for duplicate punches', () => {
  assert.equal(segment('07:00', '09:20', '09:25', '11:30').hours, 265 / 60);
});

test('six punches are three intervals and out-of-order records are sorted', () => {
  assert.equal(segment('11:30', '10:10', '07:00', '08:00', '08:20', '09:20').hours, 200 / 60);
});

test('12:10 belongs to afternoon, 12:09 belongs to morning', () => {
  assert.equal(DEFAULT_BOUNDARY, '12:10');
  const result = analyzePeriods(times('08:00', '12:09', '12:10', '16:40'));
  assert.deepEqual(result.morningTimes, times('08:00', '12:09'));
  assert.deepEqual(result.afternoonTimes, times('12:10', '16:40'));
  assert.equal(result.morning.hours, 4);
  assert.equal(result.afternoon.hours, 4.5);
});

test('half-day allowance is independent of the other half-day', () => {
  const result = analyzePeriods(times('07:00', '11:10', '13:00', '17:10'));
  assert.equal(result.morning.hours + result.afternoon.hours, 8);
});

test('empty or single-punch period has no hours; only single-punch is an error', () => {
  assert.deepEqual(segment(), { hours: 0, valid: false });
  const result = analyzePeriods(times('07:00', '11:30', '13:00'));
  assert.equal(result.morning.hours, 4.5);
  assert.equal(result.afternoon.hours, 0);
  assert.match(result.afternoon.issue, /下午.*缺少对应签到/);
});

test('odd punches retain completed pairs and identify unmatched time', () => {
  const result = segment('06:55', '09:20', '09:40', '11:29', '12:00');
  assert.equal(result.hours, 4);
  assert.match(result.issue, /12:00 缺少对应签到/);
});

test('original unambiguous duplicate example uses 07:30–11:30 and flags it', () => {
  const result = segment('07:28', '07:30', '11:30');
  assert.equal(result.hours, 4);
  assert.match(result.issue, /忽略重复记录.*07:30–11:30/);
});

test('identical timestamps do not create a zero-length work pair', () => {
  assert.equal(segment('07:00', '07:00', '11:30').hours, 4.5);
});
