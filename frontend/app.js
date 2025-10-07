import { setupDatePickers } from './calendar.js';

const fileInput = document.getElementById('file-input');
const startDateInput = document.getElementById('start-date');
const timezoneInput = document.getElementById('timezone');
const parseBtn = document.getElementById('parse-btn');
const downloadBtn = document.getElementById('download-btn');
const statusEl = document.getElementById('status');
const summarySection = document.getElementById('summary');
const summaryTitle = document.getElementById('summary-title');
const summaryMeta = document.getElementById('summary-meta');
const eventsContainer = document.getElementById('events');
const summaryCourses = document.getElementById('summary-courses');
const summaryOccurrences = document.getElementById('summary-occurrences');

let parsedResult = null;
let fileName = '';
let icsContent = '';

const STATUS_PRESETS = {
  info: {
    classes: 'border-neutral-200 bg-neutral-50 text-neutral-600'
  },
  success: {
    classes: 'border-emerald-200 bg-emerald-50 text-emerald-700'
  },
  error: {
    classes: 'border-rose-200 bg-rose-50 text-rose-700'
  }
};

function setStatus(message, kind = 'info') {
  const preset = STATUS_PRESETS[kind] || STATUS_PRESETS.info;
  statusEl.className = `rounded-2xl border px-4 py-2 text-sm transition ${preset.classes}`;
  statusEl.textContent = message;
}

function resetSummary() {
  summarySection.classList.add('hidden');
  eventsContainer.innerHTML = '';
  parsedResult = null;
  icsContent = '';
  downloadBtn.disabled = true;
}

function formatOccurrence(occurrence) {
  const dtStart = occurrence.dtStart;
  const dtEnd = occurrence.dtEnd;
  const dateObj = buildDate(dtStart);
  const weekday = dateObj.toLocaleDateString('zh-Hans', { weekday: 'long' });
  const dateLabel = dateObj.toLocaleDateString('zh-Hans', { month: '2-digit', day: '2-digit' });
  const startTime = `${dtStart.slice(9, 11)}:${dtStart.slice(11, 13)}`;
  const endTime = `${dtEnd.slice(9, 11)}:${dtEnd.slice(11, 13)}`;
  return { date: `${dtStart.slice(0, 4)}-${dtStart.slice(4, 6)}-${dtStart.slice(6, 8)}`, dateLabel, weekday, startTime, endTime };
}

function renderSummary(result) {
  summarySection.classList.remove('hidden');
  summaryTitle.textContent = result.title || '课程表';
  const totalEvents = result.events.length;
  const totalOccurrences = result.events.reduce((sum, event) => sum + event.occurrences.length, 0);
  summaryMeta.textContent = `时区 ${result.timeZone} · 开始于 ${result.startDate}`;
  summaryCourses.textContent = `课程 ${totalEvents} 门`;
  summaryOccurrences.textContent = `排课 ${totalOccurrences} 次`;

  eventsContainer.innerHTML = '';
  for (const event of result.events) {
    const wrapper = document.createElement('article');
    wrapper.className = 'flex flex-col gap-3 rounded-[24px] border border-black/5 bg-white px-5 py-5 text-[#1d1d1f] shadow-[0_16px_30px_rgba(0,0,0,0.06)] transition hover:-translate-y-0.5 hover:shadow-[0_20px_40px_rgba(0,0,0,0.08)]';

    const heading = document.createElement('div');
    heading.className = 'flex items-start justify-between gap-4';

    const titleBlock = document.createElement('div');
    titleBlock.className = 'flex flex-col gap-2';

    const titleRow = document.createElement('div');
    titleRow.className = 'flex flex-wrap items-center gap-2';

    const dayBadge = document.createElement('span');
    dayBadge.className = 'inline-flex items-center rounded-full bg-neutral-100 px-3 py-1 text-xs font-medium text-neutral-600';
    dayBadge.textContent = event.dayLabel;
    titleRow.appendChild(dayBadge);

    const title = document.createElement('h4');
    title.className = 'text-lg font-semibold';
    title.textContent = event.courseName || event.courseRaw;
    titleRow.appendChild(title);

    titleBlock.appendChild(titleRow);

    const chipRow = document.createElement('div');
    chipRow.className = 'flex flex-wrap gap-2 text-xs font-medium text-neutral-600';
    const periodLabel = event.endPeriod && event.endPeriod !== event.startPeriod
      ? `第${event.startPeriod}-${event.endPeriod}节`
      : `第${event.startPeriod}节`;
    chipRow.appendChild(createChip(periodLabel));
    chipRow.appendChild(createChip(`共 ${event.occurrences.length} 次`));
    if (event.weeksRaw) {
      chipRow.appendChild(createChip(`周次 ${event.weeksRaw}`));
    }
    titleBlock.appendChild(chipRow);

    heading.appendChild(titleBlock);

    if (event.teacher) {
      const teacherBadge = createChip(event.teacher);
      teacherBadge.classList.add('shrink-0');
      heading.appendChild(teacherBadge);
    }

    wrapper.appendChild(heading);

    const firstOccurrence = formatOccurrence(event.occurrences[0]);
    const details = document.createElement('div');
    details.className = 'space-y-1.5 text-sm text-neutral-600';
    details.appendChild(createInfoRow('日期', `${firstOccurrence.dateLabel} · ${firstOccurrence.weekday}`));
    details.appendChild(createInfoRow('时间', `${firstOccurrence.startTime} - ${firstOccurrence.endTime}`));
    if (event.location) {
      details.appendChild(createInfoRow('地点', event.location));
    }
    if (event.weeksRaw) {
      details.appendChild(createInfoRow('覆盖周次', event.weeksRaw));
    }
    wrapper.appendChild(details);

    const nextOccurrence = firstOccurrence;
    const nextHint = document.createElement('p');
    nextHint.className = 'rounded-2xl border border-black/5 bg-neutral-50 px-3 py-2 text-xs text-neutral-500';
    nextHint.textContent = `最近排课：${nextOccurrence.date} ${nextOccurrence.startTime}-${nextOccurrence.endTime}`;
    wrapper.appendChild(nextHint);

    eventsContainer.appendChild(wrapper);
  }
}

async function handleParse() {
  const file = fileInput.files && fileInput.files[0];
  if (!file) {
    setStatus('请先选择课程表文件', 'error');
    resetSummary();
    return;
  }
  const startDate = startDateInput.value;
  const timeZone = timezoneInput.value.trim() || 'Asia/Shanghai';
  if (!startDate) {
    setStatus('请输入第一周周一的日期', 'error');
    resetSummary();
    return;
  }
  setStatus('解析中，请稍候…', 'info');
  try {
    const pkgData = await file.arrayBuffer();
    const response = await fetch(`/api/parse?start=${encodeURIComponent(startDate)}&tz=${encodeURIComponent(timeZone)}`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/octet-stream',
        'X-Filename': encodeURIComponent(file.name)
      },
      body: pkgData
    });
    if (!response.ok) {
      const errorPayload = await safeParseJson(response);
      const message = errorPayload?.error || `解析失败（HTTP ${response.status}）`;
      throw new Error(message);
    }
    const result = await response.json();
    parsedResult = result;
    icsContent = result.ics;
    fileName = file.name.replace(/\.[^.]+$/, '') || 'schedule';
    renderSummary(result);
    downloadBtn.disabled = false;
    setStatus('解析完成，可以下载 ICS', 'success');
  } catch (error) {
    console.error(error);
    setStatus(`解析失败：${error.message}`, 'error');
    resetSummary();
  }
}

function handleDownload() {
  if (!parsedResult || !icsContent) {
    setStatus('请先完成解析', 'error');
    return;
  }
  try {
    const blob = new Blob([icsContent], { type: 'text/calendar;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `${fileName || 'schedule'}.ics`;
    document.body.appendChild(anchor);
    anchor.click();
    document.body.removeChild(anchor);
    URL.revokeObjectURL(url);
  } catch (error) {
    console.error(error);
    setStatus(`导出失败：${error.message}`, 'error');
  }
}

parseBtn.addEventListener('click', handleParse);
downloadBtn.addEventListener('click', handleDownload);
fileInput.addEventListener('change', () => {
  setStatus('文件已准备，点击“解析课程表”开始处理', 'info');
  resetSummary();
});

async function safeParseJson(response) {
  try {
    return await response.json();
  } catch (error) {
    return null;
  }
}

setStatus('请选择课表文件并点击“解析课程表”', 'info');
setupDatePickers();

function createChip(text) {
  const span = document.createElement('span');
  span.className = 'inline-flex items-center rounded-full border border-neutral-200 bg-neutral-50 px-3 py-1 text-xs';
  span.textContent = text;
  return span;
}

function createInfoRow(icon, text) {
  const row = document.createElement('div');
  row.className = 'flex items-center gap-2';
  const label = document.createElement('span');
  label.className = 'w-12 text-xs text-neutral-500';
  label.textContent = icon;
  const textEl = document.createElement('span');
  textEl.className = 'flex-1';
  textEl.textContent = text;
  row.append(label, textEl);
  return row;
}

function buildDate(dtString) {
  const year = Number(dtString.slice(0, 4));
  const month = Number(dtString.slice(4, 6));
  const day = Number(dtString.slice(6, 8));
  const hour = Number(dtString.slice(9, 11));
  const minute = Number(dtString.slice(11, 13));
  return new Date(year, month - 1, day, hour, minute);
}
