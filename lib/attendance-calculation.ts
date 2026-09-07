export const DEFAULT_BOUNDARY = '12:10';
export const DEFAULT_BOUNDARY_MINUTES = 12 * 60 + 10;

export type SegmentResult = { hours: number; valid: boolean; issue?: string };

// Apply the allowance once to the accumulated half-day, not to each interval.
export function countedMinutes(minutes: number): number {
  const remainder = minutes % 30;
  return remainder <= 15 ? minutes - remainder : minutes;
}

export function analyzeSegment(times: number[], label: string, window: number): SegmentResult {
  const sorted = [...new Set(times)].sort((a, b) => a - b);
  if (!sorted.length) return { hours: 0, valid: false };
  if (sorted.length === 1) return { hours: 0, valid: false, issue: `${label}仅有 1 次打卡，缺少对应签到` };

  const records = sorted.map(formatTime).join('、');
  // Even records represent complete in/out pairs, including short breaks.
  // Only use the legacy duplicate heuristic for an odd, two-cluster record.
  if (sorted.length % 2 === 1) {
    const clusters: number[][] = [];
    for (const time of sorted) {
      const last = clusters.at(-1);
      if (last && time - last[0] <= window) last.push(time);
      else clusters.push([time]);
    }
    if (clusters.length === 2) {
      const start = clusters[0].at(-1)!;
      const end = clusters[1][0];
      return {
        hours: countedMinutes(end - start) / 60,
        valid: true,
        issue: `${label}出现 ${sorted.length} 次打卡，已忽略重复记录并取 ${formatTime(start)}–${formatTime(end)}；原记录：${records}`,
      };
    }
  }

  let minutes = 0;
  const intervals: string[] = [];
  for (let index = 0; index + 1 < sorted.length; index += 2) {
    minutes += sorted[index + 1] - sorted[index];
    intervals.push(`${formatTime(sorted[index])}–${formatTime(sorted[index + 1])}`);
  }
  return {
    hours: countedMinutes(minutes) / 60,
    valid: true,
    ...(sorted.length % 2 === 1 ? {
      issue: `${label}出现 ${sorted.length} 次打卡，${formatTime(sorted.at(-1)!)} 缺少对应签到，该次不计时；已按顺序配对计算 ${intervals.join('、')}；原记录：${records}`,
    } : {}),
  };
}

export function analyzePeriods(times: number[], boundary = DEFAULT_BOUNDARY_MINUTES, duplicateWindow = 10) {
  const morningTimes = times.filter((time) => time < boundary);
  const afternoonTimes = times.filter((time) => time >= boundary);
  return {
    morningTimes,
    afternoonTimes,
    morning: analyzeSegment(morningTimes, '上午', duplicateWindow),
    afternoon: analyzeSegment(afternoonTimes, '下午', duplicateWindow),
  };
}

function formatTime(value: number) {
  return `${String(Math.floor(value / 60)).padStart(2, '0')}:${String(value % 60).padStart(2, '0')}`;
}
