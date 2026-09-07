'use client';

import { DragEvent, useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';
import { AlertTriangle, CalendarDays, CheckCircle2, ChevronDown, Clock3, Download, FileSpreadsheet, Filter, Info, Search, Settings2, UploadCloud, Users, XCircle } from 'lucide-react';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle, DialogTrigger } from '@/components/ui/dialog';
import { Input } from '@/components/ui/input';
import { Label } from '@/components/ui/label';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from '@/components/ui/select';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from '@/components/ui/table';
import { parseAttendanceSheets, type PunchGroup } from '@/lib/attendance-parser';
import { analyzePeriods, DEFAULT_BOUNDARY, DEFAULT_BOUNDARY_MINUTES } from '@/lib/attendance-calculation';

type DailyRow = { name: string; date: string; morning: number; afternoon: number; total: number; status: string; type: 'ok' | 'half' | 'warn'; morningTimes: number[]; afternoonTimes: number[]; issues: string[] };
type SummaryRow = { name: string; days: number; total: number; exceptionDates: string[] };
type IssueRow = { name: string; date: string; period: string; punches: string; reason: string };

const sampleGroups: PunchGroup[] = [
  { name: '林晓雯', date: '2026-08-03', times: [485, 720, 810, 1065] },
  { name: '陈嘉豪', date: '2026-08-03', times: [481, 723, 811] },
  { name: '周宁', date: '2026-08-03', times: [805, 1056] },
  { name: '林晓雯', date: '2026-08-04', times: [478, 715, 808, 1067] },
  { name: '陈嘉豪', date: '2026-08-04', times: [448, 450, 558, 690, 810] },
  { name: '周宁', date: '2026-08-04', times: [492, 718, 813, 1060] },
];

export default function Home() {
  const fileRef = useRef<HTMLInputElement>(null);
  const [groups, setGroups] = useState<PunchGroup[]>(sampleGroups);
  const [fileName, setFileName] = useState('示例考勤数据');
  const [sourceRows, setSourceRows] = useState(24);
  const [boundary, setBoundary] = useState(DEFAULT_BOUNDARY);
  const [duplicateWindow, setDuplicateWindow] = useState(10);
  const [query, setQuery] = useState('');
  const [filter, setFilter] = useState('all');
  const [error, setError] = useState('');
  const [dragging, setDragging] = useState(false);

  const boundaryMinutes = timeToMinutes(boundary) ?? DEFAULT_BOUNDARY_MINUTES;
  const daily = useMemo(() => analyzeGroups(groups, boundaryMinutes, duplicateWindow), [groups, boundaryMinutes, duplicateWindow]);
  const summaries = useMemo(() => summarize(daily), [daily]);
  const issues = useMemo(() => issueRows(daily), [daily]);
  const visibleDaily = daily.filter((row) => matches(row.name, row.date, query) && (filter === 'all' || (filter === 'normal' ? row.type === 'ok' : filter === 'half' ? row.type === 'half' : row.type === 'warn')));
  const visibleSummaries = summaries.filter((row) => matches(row.name, '', query));
  const visibleIssues = issues.filter((row) => matches(row.name, row.date, query));
  const totalHours = summaries.reduce((sum, row) => sum + row.total, 0);
  const employeeCount = summaries.length;
  const attendanceDays = daily.filter((row) => row.total > 0).length;
  const completeness = daily.length ? Math.max(0, 100 - (issues.length / daily.length) * 100) : 0;
  const monthLabel = formatMonth(daily[0]?.date);
  const stateRef = useRef({ daily, summaries, issues });
  stateRef.current = { daily, summaries, issues };

  useEffect(() => {
    const modelContext = (document as Document & { modelContext?: { registerTool: (tool: unknown, options?: { signal?: AbortSignal }) => void | Promise<void> } }).modelContext;
    if (!modelContext?.registerTool) return;
    const lifecycle = new AbortController();
    void Promise.resolve(modelContext.registerTool({
      name: 'read_attendance_summary',
      title: '读取考勤分析结果',
      description: '读取当前页面已经完成的考勤汇总和异常数量。',
      inputSchema: { type: 'object', properties: {}, additionalProperties: false },
      annotations: { readOnlyHint: true, untrustedContentHint: true },
      execute: () => ({ employees: stateRef.current.summaries.length, attendanceDays: stateRef.current.daily.filter((row) => row.total > 0).length, totalHours: round2(stateRef.current.summaries.reduce((sum, row) => sum + row.total, 0)), exceptionItems: stateRef.current.issues.length }),
    }, { signal: lifecycle.signal })).catch(() => undefined);
    return () => lifecycle.abort();
  }, []);

  async function handleFile(file?: File) {
    if (!file) return;
    setError('');
    try {
      const buffer = await file.arrayBuffer();
      const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });
      const sheets = workbook.SheetNames.map((sheetName) => ({
        matrix: XLSX.utils.sheet_to_json<unknown[]>(workbook.Sheets[sheetName], { header: 1, defval: '', raw: false, dateNF: 'yyyy-mm-dd hh:mm:ss' }),
        records: XLSX.utils.sheet_to_json<Record<string, unknown>>(workbook.Sheets[sheetName], { defval: '', raw: false, dateNF: 'yyyy-mm-dd hh:mm:ss' }),
      }));
      const { groups: merged, rowCount } = parseAttendanceSheets(sheets);
      if (!merged.length) throw new Error('未识别到“姓名、日期、打卡时间”数据，请检查表头或时间格式。');
      setGroups(merged); setFileName(file.name); setSourceRows(rowCount);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : '文件解析失败，请换一个 Excel 或 CSV 文件重试。');
    } finally { if (fileRef.current) fileRef.current.value = ''; }
  }

  function exportResults() {
    const detailData = daily.map((row) => ({ 姓名: row.name, 日期: row.date, '上午时长(小时)': row.morning, '下午时长(小时)': row.afternoon, '当天总时长(小时)': row.total, '状态/异常说明': row.status }));
    const summaryData = summaries.map((row) => ({ 姓名: row.name, 本月出勤天数: row.days, '本月总工作时长(小时)': row.total, 异常天数汇总: row.exceptionDates.join('、') || '无' }));
    const issueData = issues.map((row) => ({ 姓名: row.name, 日期: row.date, 时间段: row.period, 打卡记录: row.punches, 异常原因: row.reason }));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(detailData), '每日考勤明细');
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summaryData), '月度考勤汇总');
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(issueData), '异常与待确认事项');
    XLSX.writeFile(workbook, `考勤分析结果_${daily[0]?.date.slice(0, 7) || '本月'}.xlsx`);
  }

  function onDrop(event: DragEvent) { event.preventDefault(); setDragging(false); void handleFile(event.dataTransfer.files?.[0]); }

  return (
    <main className="min-h-screen bg-background text-foreground">
      <header className="topbar">
        <div className="brand"><div className="brand-mark"><Clock3 size={20} strokeWidth={2.4} /></div><div><div className="brand-name">时序</div><div className="brand-sub">考勤分析工作台</div></div></div>
        <div className="header-actions"><RuleDialog boundary={boundary} setBoundary={setBoundary} duplicateWindow={duplicateWindow} setDuplicateWindow={setDuplicateWindow} /><div className="header-divider" /><div className="avatar">HR</div></div>
      </header>

      <section className="workspace">
        <div className="page-heading">
          <div><div className="eyebrow"><span /> {monthLabel || '考勤数据待上传'}</div><h1>考勤核算</h1><p>上传原始签到表，系统将按上午与下午分别核算工时并识别异常。</p></div>
          <div className="heading-actions"><Button variant="outline" className="export-btn" onClick={exportResults} disabled={!daily.length}><Download size={17} />导出结果</Button><Button className="upload-btn" onClick={() => fileRef.current?.click()}><UploadCloud size={17} />上传新表</Button></div>
        </div>
        <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" className="sr-only" onChange={(event) => void handleFile(event.target.files?.[0])} />

        <div className={`file-strip ${dragging ? 'dragging' : ''}`} onDragOver={(event) => { event.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={onDrop}>
          <div className="file-icon"><FileSpreadsheet size={21} /></div><div className="file-copy"><strong>{fileName}</strong><span>{fileName === '示例考勤数据' ? '当前展示示例结果 · 可点击“上传新表”替换' : `已读取 ${sourceRows.toLocaleString()} 行原始记录 · 最近更新 刚刚`}</span></div>
          <div className="file-progress"><span>数据完整度</span><strong>{completeness.toFixed(1)}%</strong><div><i style={{ width: `${completeness}%` }} /></div></div>
          <Badge className="ready-badge">{fileName === '示例考勤数据' ? '示例数据' : '分析完成'}</Badge><Button variant="ghost" size="icon" aria-label="选择其他文件" onClick={() => fileRef.current?.click()}><ChevronDown size={17} /></Button>
        </div>
        {error && <div className="error-banner"><XCircle size={17} /><span>{error}</span><button onClick={() => setError('')}>关闭</button></div>}

        <div className="metric-grid">
          <Metric icon={<Users />} label="员工人数" value={String(employeeCount)} note="本月有签到记录" tone="blue" />
          <Metric icon={<CalendarDays />} label="出勤人次" value={String(attendanceDays)} note={`共 ${daily.length} 个考勤日期`} tone="violet" />
          <Metric icon={<Clock3 />} label="总工作时长" value={totalHours.toFixed(1)} unit="小时" note={employeeCount ? `人均 ${(totalHours / employeeCount).toFixed(1)} 小时` : '暂无数据'} tone="teal" />
          <Metric icon={<AlertTriangle />} label="异常事项" value={String(issues.length)} note={`涉及 ${new Set(issues.map((row) => row.name)).size} 名员工`} tone="orange" />
        </div>

        <Tabs defaultValue="detail" className="result-panel">
          <div className="result-toolbar">
            <TabsList className="tab-list"><TabsTrigger value="detail">每日明细 <span>{daily.length}</span></TabsTrigger><TabsTrigger value="summary">月度汇总 <span>{summaries.length}</span></TabsTrigger><TabsTrigger value="issues">异常事项 <span className="issue-count">{issues.length}</span></TabsTrigger></TabsList>
            <div className="table-tools"><div className="search-wrap"><Search size={16} /><Input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="搜索姓名或日期" /></div><Select value={filter} onValueChange={(value) => setFilter(value ?? 'all')}><SelectTrigger size="sm" className="filter-select"><Filter size={15} /><SelectValue /></SelectTrigger><SelectContent><SelectItem value="all">全部状态</SelectItem><SelectItem value="normal">正常</SelectItem><SelectItem value="half">半天</SelectItem><SelectItem value="exception">异常</SelectItem></SelectContent></Select></div>
          </div>
          <TabsContent value="detail" className="tab-content"><DailyTable rows={visibleDaily} /></TabsContent>
          <TabsContent value="summary" className="tab-content"><SummaryTable rows={visibleSummaries} /></TabsContent>
          <TabsContent value="issues" className="tab-content"><IssueTable rows={visibleIssues} /></TabsContent>
        </Tabs>
        <div className="logic-note"><Info size={16} /><span>计算口径：<strong>{boundary}</strong> 前为上午，起为下午；半天内按顺序两两配对后累加。每半天超出半小时档位的部分 ≤15 分钟舍去，超过 15 分钟保留实际时长；未配对打卡不计时并标记异常。</span></div>
      </section>
    </main>
  );
}

function RuleDialog({ boundary, setBoundary, duplicateWindow, setDuplicateWindow }: { boundary: string; setBoundary: (value: string) => void; duplicateWindow: number; setDuplicateWindow: (value: number) => void }) {
  const [nextBoundary, setNextBoundary] = useState(boundary); const [nextWindow, setNextWindow] = useState(String(duplicateWindow));
  return <Dialog><DialogTrigger render={<Button variant="ghost" size="sm" className="rule-button" />}><Settings2 size={16} />计算规则</DialogTrigger><DialogContent className="rule-dialog"><DialogHeader><DialogTitle>计算规则</DialogTitle><DialogDescription>应用后将重新计算每日明细、月度汇总和导出结果。</DialogDescription></DialogHeader><div className="rule-fields">
    <div><Label htmlFor="boundary">上午 / 下午分界（默认 12:10）</Label><Input id="boundary" type="time" value={nextBoundary} onChange={(event) => setNextBoundary(event.target.value)} /></div>
    <div><Label htmlFor="duplicate">奇数次打卡的重复识别窗口（分钟）</Label><Input id="duplicate" type="number" min="1" max="60" value={nextWindow} onChange={(event) => setNextWindow(event.target.value)} /></div>
    <div className="rule-summary"><CheckCircle2 size={17} /><span>分界时刻之前为上午，从分界时刻起为下午。半天内按时间排序，第 1–2 次、第 3–4 次依次配对并累加，中间离岗时间不计入。</span></div>
    <div className="rule-summary"><CheckCircle2 size={17} /><span>每半天合计后，超出半小时档位的部分不超过 15 分钟则舍去；超过 15 分钟保留实际时长。例如 4小时40分 → 4小时30分，4小时46分保持不变。</span></div>
    <div className="rule-summary"><Info size={17} /><span>偶数次打卡直接配对，不合并短暂离岗。奇数次仅在记录可归为上班、下班两组时按窗口去重并提示；其他情况计算完整配对，剩余单次不计时并标记缺卡。</span></div>
  </div><DialogFooter><Button onClick={() => { if (timeToMinutes(nextBoundary) !== null) setBoundary(nextBoundary); setDuplicateWindow(Math.min(60, Math.max(1, Number(nextWindow) || 10))); }}>应用规则</Button></DialogFooter></DialogContent></Dialog>;
}

function Metric({ icon, label, value, unit, note, tone }: { icon: React.ReactNode; label: string; value: string; unit?: string; note: string; tone: string }) { return <article className="metric-card"><div className={`metric-icon ${tone}`}>{icon}</div><div><p>{label}</p><div className="metric-value">{value}{unit && <small>{unit}</small>}</div><span>{note}</span></div></article>; }

function DailyTable({ rows }: { rows: DailyRow[] }) {
  if (!rows.length) return <EmptyTable />;
  return <div className="table-scroll"><Table><TableHeader><TableRow><TableHead>姓名</TableHead><TableHead>日期</TableHead><TableHead>打卡记录</TableHead><TableHead className="text-right">上午时长</TableHead><TableHead className="text-right">下午时长</TableHead><TableHead className="text-right">当天总时长</TableHead><TableHead>状态 / 异常说明</TableHead></TableRow></TableHeader><TableBody>{rows.map((row) => <TableRow key={`${row.name}-${row.date}`}><TableCell className="employee"><span>{row.name.slice(0, 1)}</span>{row.name}</TableCell><TableCell className="date-cell">{row.date}</TableCell><TableCell className="punches">{[...row.morningTimes, ...row.afternoonTimes].map(formatTime).join(' · ')}</TableCell><Hours value={row.morning} /><Hours value={row.afternoon} /><Hours value={row.total} total /><TableCell><Badge className={`status ${row.type}`}>{row.status}</Badge></TableCell></TableRow>)}</TableBody></Table></div>;
}

function SummaryTable({ rows }: { rows: SummaryRow[] }) {
  if (!rows.length) return <EmptyTable />;
  return <div className="table-scroll"><Table><TableHeader><TableRow><TableHead>姓名</TableHead><TableHead className="text-right">本月出勤天数</TableHead><TableHead className="text-right">本月总工作时长</TableHead><TableHead>异常天数汇总</TableHead></TableRow></TableHeader><TableBody>{rows.map((row) => <TableRow key={row.name}><TableCell className="employee"><span>{row.name.slice(0, 1)}</span>{row.name}</TableCell><TableCell className="hours total">{row.days}<small> 天</small></TableCell><Hours value={row.total} total /><TableCell>{row.exceptionDates.length ? <div className="date-badges">{row.exceptionDates.map((date) => <Badge key={date} className="status warn">{date}</Badge>)}</div> : <span className="no-issue"><CheckCircle2 size={15} />无异常</span>}</TableCell></TableRow>)}</TableBody></Table></div>;
}

function IssueTable({ rows }: { rows: IssueRow[] }) {
  if (!rows.length) return <div className="all-clear"><CheckCircle2 size={34} /><strong>没有待处理异常</strong><span>当前筛选范围内的打卡记录均可正常计算。</span></div>;
  return <div className="table-scroll"><Table><TableHeader><TableRow><TableHead>员工 / 日期</TableHead><TableHead>时间段</TableHead><TableHead>原始打卡记录</TableHead><TableHead>异常原因与处理结果</TableHead></TableRow></TableHeader><TableBody>{rows.map((row, index) => <TableRow key={`${row.name}-${row.date}-${index}`}><TableCell><div className="issue-person"><strong>{row.name}</strong><span>{row.date}</span></div></TableCell><TableCell><Badge className="period-badge">{row.period}</Badge></TableCell><TableCell className="punches">{row.punches || '无'}</TableCell><TableCell className="reason"><AlertTriangle size={15} />{row.reason}</TableCell></TableRow>)}</TableBody></Table></div>;
}

function Hours({ value, total }: { value: number; total?: boolean }) { return <TableCell className={`hours ${total ? 'total' : ''}`}>{value.toFixed(2)}<small> h</small></TableCell>; }
function EmptyTable() { return <div className="empty-view"><Search size={30} /><strong>没有匹配记录</strong><span>请调整搜索关键词或筛选条件。</span></div>; }

function analyzeGroups(groups: PunchGroup[], boundary: number, duplicateWindow: number): DailyRow[] {
  return mergeGroups(groups).map((group) => {
    const { morningTimes, afternoonTimes, morning, afternoon } = analyzePeriods(group.times, boundary, duplicateWindow);
    const issues = [morning.issue, afternoon.issue].filter(Boolean) as string[]; const total = round2(morning.hours + afternoon.hours);
    let status = '正常'; let type: DailyRow['type'] = 'ok';
    if (issues.length) { status = issues.join('；'); type = 'warn'; } else if ((morning.valid && !afternoon.valid) || (!morning.valid && afternoon.valid)) { status = '工作半天'; type = 'half'; } else if (!morning.valid && !afternoon.valid) { status = '无有效工时'; type = 'warn'; }
    return { name: group.name, date: group.date, morning: round2(morning.hours), afternoon: round2(afternoon.hours), total, status, type, morningTimes, afternoonTimes, issues };
  }).sort((a, b) => a.date.localeCompare(b.date) || a.name.localeCompare(b.name, 'zh-CN'));
}

function summarize(rows: DailyRow[]): SummaryRow[] {
  const map = new Map<string, SummaryRow>();
  for (const row of rows) { const item = map.get(row.name) ?? { name: row.name, days: 0, total: 0, exceptionDates: [] }; if (row.total > 0) item.days += 1; item.total = round2(item.total + row.total); if (row.issues.length || row.type === 'warn') item.exceptionDates.push(row.date); map.set(row.name, item); }
  return [...map.values()].map((row) => ({ ...row, exceptionDates: [...new Set(row.exceptionDates)] })).sort((a, b) => a.name.localeCompare(b.name, 'zh-CN'));
}

function issueRows(rows: DailyRow[]): IssueRow[] {
  const result: IssueRow[] = [];
  for (const row of rows) { if (!row.issues.length && row.type === 'warn') result.push({ name: row.name, date: row.date, period: '全天', punches: [...row.morningTimes, ...row.afternoonTimes].map(formatTime).join('、'), reason: row.status }); for (const reason of row.issues) { const period = reason.startsWith('上午') ? '上午' : '下午'; const times = period === '上午' ? row.morningTimes : row.afternoonTimes; result.push({ name: row.name, date: row.date, period, punches: times.map(formatTime).join('、'), reason }); } }
  return result;
}

function mergeGroups(groups: PunchGroup[]): PunchGroup[] {
  const map = new Map<string, PunchGroup>();
  for (const group of groups) { if (!group.name || !group.date || !group.times.length) continue; const key = `${group.name}\u0000${group.date}`; const item = map.get(key) ?? { name: group.name, date: group.date, times: [] }; item.times.push(...group.times); item.times = [...new Set(item.times)].sort((a, b) => a - b); map.set(key, item); }
  return [...map.values()];
}

function timeToMinutes(value: string) { const match = value.match(/^(\d{1,2}):([0-5]\d)$/); if (!match || Number(match[1]) > 23) return null; return Number(match[1]) * 60 + Number(match[2]); }
function formatTime(value: number) { return `${String(Math.floor(value / 60)).padStart(2, '0')}:${String(value % 60).padStart(2, '0')}`; }
function formatMonth(value?: string) { if (!value) return ''; const [year, month] = value.split('-'); return `${year} 年 ${Number(month)} 月考勤`; }
function matches(name: string, date: string, query: string) { const needle = query.trim().toLowerCase(); return !needle || name.toLowerCase().includes(needle) || date.includes(needle); }
function round2(value: number) { return Math.round(value * 100) / 100; }
