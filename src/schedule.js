const ZIP_SIG_LOCAL_FILE_HEADER = 0x04034b50;
const ZIP_SIG_CENTRAL_DIRECTORY = 0x02014b50;
const ZIP_SIG_END_OF_CENTRAL_DIRECTORY = 0x06054b50;
const ZIP_COMPRESSION_STORE = 0;
const ZIP_COMPRESSION_DEFLATE = 8;

const textDecoder = new TextDecoder('utf-8');
const textEncoder = new TextEncoder();
const globalScope = typeof globalThis !== 'undefined'
  ? globalThis
  : typeof window !== 'undefined'
    ? window
    : typeof global !== 'undefined'
      ? global
      : {};

const CELL_PARAGRAPH_REGEX = /<w:p[\s\S]*?<\/w:p>/g;
const CELL_REGEX = /<w:tc[\s\S]*?<\/w:tc>/g;
const ROW_REGEX = /<w:tr[\s\S]*?<\/w:tr>/g;
const TABLE_REGEX = /<w:tbl[\s\S]*?<\/w:tbl>/g;
const TEXT_REGEX = /<w:t[^>]*>([\s\S]*?)<\/w:t>/g;

const DAY_OFFSETS = {
  '星期一': 0,
  '星期二': 1,
  '星期三': 2,
  '星期四': 3,
  '星期五': 4,
  '星期六': 5,
  '星期日': 6,
  '星期天': 6
};

const ZIP_DECODER_NOT_AVAILABLE_MESSAGE = '环境不支持 ZIP/deflate 解压，请上传 Flat OPC XML 文件';

export async function parseSchedule(input, options = {}) {
  const { startDate = '2025-09-08', timeZone = 'Asia/Shanghai' } = options;
  const documentXml = await resolveDocumentXml(input);
  const title = extractTitle(documentXml);
  const tables = parseTables(documentXml);
  if (!tables.length) {
    throw new Error('未找到课程表的表格信息');
  }
  const timetableRows = parseTable(tables[0]);
  const { grid } = buildGrid(timetableRows);
  const periods = extractPeriods(grid);
  const columnDayMap = buildColumnDayMap(grid);
  const baseDate = createBaseDate(startDate, timeZone);
  const events = collectEvents(grid, columnDayMap, periods, baseDate, timeZone);
  return { title, events, periods, columnDayMap, timeZone, startDate };
}

export async function parseScheduleFromPackage(input, options = {}) {
  return parseSchedule(input, options);
}

export function buildIcs(events, timeZone) {
  const lines = [];
  const dtstamp = new Date().toISOString().replace(/[-:]/g, '').replace(/\.\d{3}Z$/, 'Z');
  lines.push('BEGIN:VCALENDAR');
  lines.push('PRODID:-//calendar-helper//CN');
  lines.push('VERSION:2.0');
  lines.push('CALSCALE:GREGORIAN');
  lines.push('METHOD:PUBLISH');
  lines.push(`X-WR-TIMEZONE:${timeZone}`);
  for (const event of events) {
    if (!event.occurrences.length) continue;
    const descriptionParts = [];
    if (event.courseType) descriptionParts.push(event.courseType);
    if (event.teacher) descriptionParts.push(`教师: ${event.teacher}`);
    if (event.enrollment != null) descriptionParts.push(`人数: ${event.enrollment}`);
    if (event.weeksRaw) descriptionParts.push(`周次: ${event.weeksRaw}`);
    const description = escapeIcsText(descriptionParts.join('\n'));
    const summary = escapeIcsText(event.courseName || event.courseRaw);
    const location = escapeIcsText(event.location || '');
    const first = event.occurrences[0];
    lines.push('BEGIN:VEVENT');
    lines.push(`UID:${event.uid}`);
    lines.push(`DTSTAMP:${dtstamp}`);
    lines.push(`SUMMARY:${summary}`);
    if (event.location) lines.push(`LOCATION:${location}`);
    lines.push(`DESCRIPTION:${description}`);
    lines.push(`DTSTART;TZID=${timeZone}:${first.dtStart}`);
    lines.push(`DTEND;TZID=${timeZone}:${first.dtEnd}`);
    if (event.rrule) {
      lines.push(`RRULE:FREQ=WEEKLY;INTERVAL=${event.rrule.interval};COUNT=${event.rrule.count}`);
    } else if (event.rdates.length) {
      lines.push(`RDATE;TZID=${timeZone}:${event.rdates.join(',')}`);
    }
    lines.push('END:VEVENT');
  }
  lines.push('END:VCALENDAR');
  return lines.join('\r\n');
}

async function resolveDocumentXml(input) {
  if (typeof input === 'string') {
    const trimmed = input.trimStart();
    if (trimmed.startsWith('<?xml')) {
      if (trimmed.includes('<pkg:package')) {
        return extractDocumentXmlFromFlatOpc(trimmed);
      }
      return trimmed;
    }
    input = textEncoder.encode(input);
  }

  if (input instanceof ArrayBuffer) {
    input = new Uint8Array(input);
  } else if (ArrayBuffer.isView(input)) {
    input = new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  }

  if (input instanceof Uint8Array) {
    if (input.length < 4) {
      throw new Error('文件数据过短，无法解析');
    }
    const signature = readUint32LE(input, 0);
    if (signature === ZIP_SIG_LOCAL_FILE_HEADER || signature === ZIP_SIG_END_OF_CENTRAL_DIRECTORY) {
      return extractDocumentXmlFromZip(input);
    }
    const text = textDecoder.decode(input);
    const trimmed = text.trimStart();
    if (trimmed.startsWith('<?xml')) {
      if (trimmed.includes('<pkg:package')) {
        return extractDocumentXmlFromFlatOpc(trimmed);
      }
      return trimmed;
    }
    throw new Error('无法识别的文件格式，请确认是 Word 导出的课表文件');
  }

  throw new Error('不支持的输入类型，必须是字符串或字节数组');
}

function extractDocumentXmlFromFlatOpc(pkgXml) {
  const partMatch = pkgXml.match(/<pkg:part[^>]*pkg:name="\/word\/document\.xml"[^>]*>([\s\S]*?)<\/pkg:part>/);
  if (!partMatch) {
    throw new Error('Flat OPC 中缺少 /word/document.xml');
  }
  const xmlDataMatch = partMatch[1].match(/<pkg:xmlData>([\s\S]*?)<\/pkg:xmlData>/);
  if (!xmlDataMatch) {
    throw new Error('Flat OPC 中缺少 document.xml 的 xmlData');
  }
  return xmlDataMatch[1];
}

async function extractDocumentXmlFromZip(bytes) {
  try {
    return await extractDocumentXmlFromZipManual(bytes);
  } catch (error) {
    const fallback = await extractDocumentXmlUsingJsZip(bytes, error);
    if (fallback) return fallback;
    throw error;
  }
}

async function extractDocumentXmlFromZipManual(bytes) {
  const entry = locateZipEntry(bytes, 'word/document.xml');
  if (!entry) {
    throw new Error('ZIP 包中缺少 word/document.xml');
  }
  const { compression, compressedSize, localHeaderOffset } = entry;
  const nameLength = readUint16LE(bytes, localHeaderOffset + 26);
  const extraLength = readUint16LE(bytes, localHeaderOffset + 28);
  const dataStart = localHeaderOffset + 30 + nameLength + extraLength;
  const dataEnd = dataStart + compressedSize;
  if (dataEnd > bytes.length) {
    throw new Error('ZIP 数据损坏，无法读取完整文件');
  }
  const fileSlice = bytes.subarray(dataStart, dataEnd);
  if (compression === ZIP_COMPRESSION_STORE) {
    return textDecoder.decode(fileSlice);
  }
  if (compression === ZIP_COMPRESSION_DEFLATE) {
    const inflated = await inflateRaw(fileSlice);
    return textDecoder.decode(inflated);
  }
  throw new Error(`不支持的 ZIP 压缩方式: ${compression}`);
}

function locateZipEntry(bytes, targetName) {
  const eocdOffset = findEndOfCentralDirectory(bytes);
  if (eocdOffset === -1) {
    throw new Error('未找到 ZIP 尾部目录');
  }
  const centralDirectoryOffset = readUint32LE(bytes, eocdOffset + 16);
  const totalEntries = readUint16LE(bytes, eocdOffset + 10);
  let offset = centralDirectoryOffset;
  const decoder = new TextDecoder('utf-8');
  for (let i = 0; i < totalEntries; i++) {
    const signature = readUint32LE(bytes, offset);
    if (signature !== ZIP_SIG_CENTRAL_DIRECTORY) {
      throw new Error('ZIP 中央目录损坏');
    }
    const compression = readUint16LE(bytes, offset + 10);
    const compressedSize = readUint32LE(bytes, offset + 20);
    const fileNameLength = readUint16LE(bytes, offset + 28);
    const extraLength = readUint16LE(bytes, offset + 30);
    const commentLength = readUint16LE(bytes, offset + 32);
    const localHeaderOffset = readUint32LE(bytes, offset + 42);
    const nameStart = offset + 46;
    const nameEnd = nameStart + fileNameLength;
    const fileName = decoder.decode(bytes.subarray(nameStart, nameEnd));
    if (fileName === targetName) {
      return { compression, compressedSize, localHeaderOffset };
    }
    offset = nameEnd + extraLength + commentLength;
  }
  return null;
}

function findEndOfCentralDirectory(bytes) {
  const minEOCD = 22;
  const maxCommentLength = 0xffff;
  const start = Math.max(0, bytes.length - (minEOCD + maxCommentLength));
  for (let i = bytes.length - minEOCD; i >= start; i--) {
    if (readUint32LE(bytes, i) === ZIP_SIG_END_OF_CENTRAL_DIRECTORY) {
      return i;
    }
  }
  return -1;
}

async function inflateRaw(data) {
  if (typeof process !== 'undefined' && process.versions?.node) {
    const { inflateRawSync } = await import('node:zlib');
    const result = inflateRawSync(data);
    return result instanceof Uint8Array
      ? result
      : new Uint8Array(result.buffer, result.byteOffset, result.byteLength);
  }
  if (typeof DecompressionStream !== 'undefined') {
    try {
      const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
      const buffer = await new Response(stream).arrayBuffer();
      return new Uint8Array(buffer);
    } catch (error) {
      try {
        const stream = new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate'));
        const buffer = await new Response(stream).arrayBuffer();
        return new Uint8Array(buffer);
      } catch (innerError) {
        // fall through to JSZip fallback below
      }
    }
  }
  throw new Error(ZIP_DECODER_NOT_AVAILABLE_MESSAGE);
}

function readUint16LE(bytes, offset) {
  return bytes[offset] | (bytes[offset + 1] << 8);
}

function readUint32LE(bytes, offset) {
  return (
    bytes[offset] |
    (bytes[offset + 1] << 8) |
    (bytes[offset + 2] << 16) |
    (bytes[offset + 3] << 24)
  ) >>> 0;
}

function decodeXmlEntities(raw) {
  if (!raw) return '';
  return raw
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&#x([0-9a-fA-F]+);/g, (_, hex) => String.fromCharCode(parseInt(hex, 16)))
    .replace(/&#(\d+);/g, (_, dec) => String.fromCharCode(parseInt(dec, 10)));
}

function collectTexts(xmlSegment) {
  if (!xmlSegment) return [];
  const out = [];
  let match;
  while ((match = TEXT_REGEX.exec(xmlSegment))) {
    out.push(decodeXmlEntities(match[1]));
  }
  return out;
}

function extractTitle(documentXml) {
  const [beforeTable] = documentXml.split(/<w:tbl/);
  const pieces = collectTexts(beforeTable);
  const cleaned = pieces.map((text) => text.trim()).filter(Boolean);
  return cleaned.join(' ');
}

function parseTables(documentXml) {
  const tables = [];
  let match;
  while ((match = TABLE_REGEX.exec(documentXml))) {
    tables.push(match[0]);
  }
  return tables;
}

function parseTable(tableXml) {
  const rows = [];
  let rowMatch;
  while ((rowMatch = ROW_REGEX.exec(tableXml))) {
    const rowXml = rowMatch[0];
    const cells = [];
    let cellMatch;
    while ((cellMatch = CELL_REGEX.exec(rowXml))) {
      const cellXml = cellMatch[0];
      const content = extractCellContent(cellXml);
      const colspan = extractGridSpan(cellXml);
      const vMerge = extractVMerge(cellXml);
      cells.push({ ...content, colspan, vMerge });
    }
    rows.push(cells);
  }
  return rows;
}

function extractCellContent(cellXml) {
  const paragraphs = [];
  let paraMatch;
  while ((paraMatch = CELL_PARAGRAPH_REGEX.exec(cellXml))) {
    const paraXml = paraMatch[0];
    const fragments = collectTexts(paraXml);
    const text = fragments.join('').replace(/\u000b/g, '\n');
    if (!text.trim()) continue;
    paragraphs.push(text.trim());
  }
  const text = paragraphs.join('\n');
  return { text, paragraphs };
}

function extractGridSpan(cellXml) {
  const match = cellXml.match(/<w:gridSpan[^>]*w:val="(\d+)"[^>]*\/>/);
  return match ? parseInt(match[1], 10) : 1;
}

function extractVMerge(cellXml) {
  const match = cellXml.match(/<w:vMerge(?:[^>]*w:val="([^"]+)")?[^>]*\/>/);
  if (!match) return null;
  return match[1] || 'continue';
}

function buildGrid(rows) {
  const totalCols = rows[0].reduce((sum, cell) => sum + cell.colspan, 0);
  const grid = [];
  const active = new Array(totalCols).fill(null);
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    const row = rows[rowIndex];
    const expanded = new Array(totalCols);
    let colCursor = 0;
    for (const cell of row) {
      const span = cell.colspan || 1;
      if (cell.vMerge === 'continue') {
        for (let i = 0; i < span; i++) {
          const activeCell = active[colCursor];
          if (!activeCell) {
            throw new Error(`表格合并信息异常，定位在行 ${rowIndex} 列 ${colCursor}`);
          }
          expanded[colCursor] = activeCell;
          colCursor++;
        }
        continue;
      }
      const newCell = {
        text: cell.text,
        paragraphs: cell.paragraphs,
        colspan: span,
        vMerge: cell.vMerge,
        rowStart: rowIndex,
        startColumn: colCursor
      };
      for (let i = 0; i < span; i++) {
        expanded[colCursor] = newCell;
        if (cell.vMerge === 'restart') {
          active[colCursor] = newCell;
        } else {
          active[colCursor] = null;
        }
        colCursor++;
      }
    }
    for (let col = 0; col < totalCols; col++) {
      if (!expanded[col] && active[col]) {
        expanded[col] = active[col];
      }
      if (!expanded[col]) {
        active[col] = null;
      }
    }
    grid.push(expanded);
  }
  return { grid, totalCols };
}

function extractPeriods(grid) {
  const periods = [];
  for (let rowIndex = 1; rowIndex < grid.length; rowIndex++) {
    const cell = grid[rowIndex][0];
    if (!cell || !cell.text) continue;
    const text = cell.text.replace(/\s+/g, ' ').trim();
    const match = text.match(/^第(\d+)节\s+(\d{2}:\d{2})~(\d{2}:\d{2})$/);
    if (!match) continue;
    periods[rowIndex] = {
      period: parseInt(match[1], 10),
      start: match[2],
      end: match[3]
    };
  }
  return periods;
}

function buildColumnDayMap(grid) {
  const header = grid[0];
  const map = new Array(header.length).fill(null);
  for (let col = 1; col < header.length; col++) {
    const cell = header[col];
    const label = cell && cell.text ? cell.text.trim() : '';
    map[col] = label;
  }
  return map;
}

function parseWeeks(raw) {
  const segments = raw.split(/[、,，]/).map((item) => item.trim()).filter(Boolean);
  const weeks = new Set();
  for (const seg of segments) {
    const match = seg.match(/^(\d+)(?:-(\d+))?(.+)?$/);
    if (!match) continue;
    const start = parseInt(match[1], 10);
    const end = match[2] ? parseInt(match[2], 10) : start;
    const tail = match[3] || '';
    const filter = tail.includes('单周') ? 'odd' : tail.includes('双周') ? 'even' : 'all';
    for (let w = start; w <= end; w++) {
      if (filter === 'odd' && w % 2 === 0) continue;
      if (filter === 'even' && w % 2 === 1) continue;
      weeks.add(w);
    }
  }
  return Array.from(weeks).sort((a, b) => a - b);
}

function parseCourseCell(rawText) {
  const parts = rawText.split('/').map((item) => item.trim());
  if (parts.length !== 5) {
    throw new Error(`无法解析课程信息：${rawText}`);
  }
  const [weeksRaw, courseRaw, teacher, location, sizeRaw] = parts;
  let courseType = null;
  let courseName = courseRaw;
  const typeMatch = courseRaw.match(/^(.*?[)）])(.*)$/);
  if (typeMatch) {
    courseType = typeMatch[1].trim();
    courseName = typeMatch[2].trim() || courseName;
  }
  const enrollmentMatch = sizeRaw.match(/(\d+)/);
  return {
    weeksRaw,
    weeks: parseWeeks(weeksRaw),
    courseRaw,
    courseType,
    courseName,
    teacher,
    location: location || '',
    enrollment: enrollmentMatch ? parseInt(enrollmentMatch[1], 10) : null
  };
}

function createBaseDate(start, timeZone) {
  const [yearStr, monthStr, dayStr] = start.split('-');
  const year = parseInt(yearStr, 10);
  const month = parseInt(monthStr, 10);
  const day = parseInt(dayStr, 10);
  if (!year || !month || !day) {
    throw new Error(`无效的日期：${start}`);
  }
  const approxUtc = new Date(Date.UTC(year, month - 1, day));
  const offset = getTimeZoneOffset(approxUtc, timeZone);
  return new Date(approxUtc.getTime() - offset);
}

function getTimeZoneOffset(date, timeZone) {
  const dtf = new Intl.DateTimeFormat('en-CA', {
    timeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
    hour12: false
  });
  const parts = dtf.formatToParts(date);
  const data = Object.fromEntries(parts.map((part) => [part.type, part.value]));
  const localMillis = Date.UTC(
    parseInt(data.year, 10),
    parseInt(data.month, 10) - 1,
    parseInt(data.day, 10),
    parseInt(data.hour, 10),
    parseInt(data.minute, 10),
    parseInt(data.second, 10)
  );
  const utcMillis = Date.UTC(
    date.getUTCFullYear(),
    date.getUTCMonth(),
    date.getUTCDate(),
    date.getUTCHours(),
    date.getUTCMinutes(),
    date.getUTCSeconds()
  );
  return localMillis - utcMillis;
}

function addDays(date, days) {
  const copy = new Date(date.getTime());
  copy.setUTCDate(copy.getUTCDate() + days);
  return copy;
}

function formatDateParts(date, timeZone) {
  const formatter = new Intl.DateTimeFormat('en-CA', {
    timeZone,
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  });
  const formatted = formatter.format(date);
  return formatted.split('-');
}

function formatDateTimeLocal(date, time, timeZone) {
  const [hourStr, minuteStr] = time.split(':');
  const hour = parseInt(hourStr, 10);
  const minute = parseInt(minuteStr, 10);
  const [year, month, day] = formatDateParts(date, timeZone);
  const hh = hour.toString().padStart(2, '0');
  const mm = minute.toString().padStart(2, '0');
  return `${year}${month}${day}T${hh}${mm}00`;
}

function computeEventOccurrences(weeks, baseDate, dayOffset, startRow, endRow, periods, timeZone) {
  if (!weeks.length) return [];
  const occurrences = [];
  const startTime = periods[startRow]?.start;
  const endTime = periods[endRow]?.end;
  if (!startTime || !endTime) return [];
  for (const week of weeks) {
    const dayCount = (week - 1) * 7 + dayOffset;
    const eventDate = addDays(baseDate, dayCount);
    const dtStart = formatDateTimeLocal(eventDate, startTime, timeZone);
    const dtEnd = formatDateTimeLocal(eventDate, endTime, timeZone);
    occurrences.push({ week, dtStart, dtEnd });
  }
  return occurrences;
}

function deriveRRule(weeks) {
  if (weeks.length <= 1) return null;
  const deltas = [];
  for (let i = 1; i < weeks.length; i++) {
    deltas.push(weeks[i] - weeks[i - 1]);
  }
  const uniqueDeltas = Array.from(new Set(deltas));
  if (uniqueDeltas.length !== 1) return null;
  const interval = uniqueDeltas[0];
  if (interval <= 0) return null;
  return { interval, count: weeks.length };
}

function fnv1a(value) {
  let hash = 0x811c9dc5;
  for (let i = 0; i < value.length; i++) {
    hash ^= value.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, '0');
}

function splitCourseEntries(cell) {
  const entries = [];
  const sources = cell.paragraphs && cell.paragraphs.length ? cell.paragraphs : [cell.text || ''];
  for (const paragraph of sources) {
    if (!paragraph) continue;
    const parts = paragraph
      .split(/\n|\s{2,}/)
      .map((item) => item.trim())
      .filter(Boolean);
    for (const part of parts) {
      if (part.includes('/')) {
        entries.push(part);
      }
    }
  }
  return entries;
}

function collectEvents(grid, columnDayMap, periods, baseDate, timeZone) {
  const eventMap = new Map();
  for (let rowIndex = 1; rowIndex < grid.length; rowIndex++) {
    for (let col = 1; col < grid[rowIndex].length; col++) {
      const cell = grid[rowIndex][col];
      if (!cell) continue;
      if (cell.rowStart !== rowIndex) continue;
      if (cell.startColumn !== col) continue;
      const dayLabel = columnDayMap[col];
      const dayOffset = DAY_OFFSETS[dayLabel];
      if (dayOffset == null) continue;
      const entries = splitCourseEntries(cell);
      if (!entries.length) continue;
      const endRow = findCellEndRow(grid, rowIndex, col, cell);
      for (const entry of entries) {
        let parsed;
        try {
          parsed = parseCourseCell(entry);
        } catch (error) {
          continue;
        }
        const occurrences = computeEventOccurrences(
          parsed.weeks,
          baseDate,
          dayOffset,
          rowIndex,
          endRow,
          periods,
          timeZone
        );
        if (!occurrences.length) continue;
        const startPeriod = periods[rowIndex]?.period ?? null;
        const endPeriod = periods[endRow]?.period ?? null;
        const key = [
          parsed.courseRaw,
          parsed.teacher,
          parsed.location,
          dayLabel,
          startPeriod,
          endPeriod
        ].join('||');
        let bucket = eventMap.get(key);
        if (!bucket) {
          bucket = {
            courseRaw: parsed.courseRaw,
            courseType: parsed.courseType,
            courseName: parsed.courseName,
            teacher: parsed.teacher,
            location: parsed.location,
            enrollment: parsed.enrollment,
            dayLabel,
            startPeriod,
            endPeriod,
            weeksRawSegments: new Set(),
            weeks: new Set(),
            occurrences: [],
            timeZone
          };
          eventMap.set(key, bucket);
        }
        bucket.weeksRawSegments.add(parsed.weeksRaw);
        for (const week of parsed.weeks) {
          bucket.weeks.add(week);
        }
        for (const occurrence of occurrences) {
          bucket.occurrences.push(occurrence);
        }
        if (parsed.enrollment && !bucket.enrollment) {
          bucket.enrollment = parsed.enrollment;
        }
      }
    }
  }

  const events = [];
  for (const [key, bucket] of eventMap.entries()) {
    const occurrences = Array.from(new Map(bucket.occurrences.map((occ) => [occ.dtStart, occ])).values()).sort(
      (a, b) => (a.dtStart < b.dtStart ? -1 : a.dtStart > b.dtStart ? 1 : 0)
    );
    const weeks = Array.from(bucket.weeks).sort((a, b) => a - b);
    const rrule = deriveRRule(weeks);
    const rdates = [];
    if (!rrule && occurrences.length > 1) {
      for (let i = 1; i < occurrences.length; i++) {
        rdates.push(occurrences[i].dtStart);
      }
    }
    const uid = `cal-${fnv1a(key)}`;
    events.push({
      courseRaw: bucket.courseRaw,
      courseType: bucket.courseType,
      courseName: bucket.courseName,
      teacher: bucket.teacher,
      location: bucket.location,
      enrollment: bucket.enrollment ?? null,
      dayLabel: bucket.dayLabel,
      startPeriod: bucket.startPeriod,
      endPeriod: bucket.endPeriod,
      weeksRaw: Array.from(bucket.weeksRawSegments).join('、'),
      weeks,
      occurrences,
      rrule,
      rdates,
      uid,
      timeZone
    });
  }
  return events;
}

function findCellEndRow(grid, startRow, col, targetCell) {
  let endRow = startRow;
  for (let r = startRow + 1; r < grid.length; r++) {
    const cell = grid[r][col];
    if (cell !== targetCell) break;
    endRow = r;
  }
  return endRow;
}

function escapeIcsText(value) {
  return value
    .replace(/\\/g, '\\\\')
    .replace(/\n/g, '\\n')
    .replace(/;/g, '\\;')
    .replace(/,/g, '\\,');
}

let jsZipPromise = null;

async function ensureJsZip() {
  if (globalScope && globalScope.JSZip) {
    return globalScope.JSZip;
  }
  if (jsZipPromise) {
    return jsZipPromise;
  }
  jsZipPromise = (async () => {
    try {
      // For Node.js environment
      if (typeof process !== 'undefined' && process.versions?.node) {
        const jszip = await import('jszip');
        return jszip.default || jszip;
      }
      // For browser environment with dynamic import
      if (typeof window !== 'undefined' && typeof window.fetch === 'function') {
        if (!globalScope.JSZip) {
          await new Promise((resolve, reject) => {
            const script = document.createElement('script');
            script.src = 'https://cdn.jsdelivr.net/npm/jszip@3.10.1/dist/jszip.min.js';
            script.async = true;
            script.onload = () => resolve();
            script.onerror = () => reject(new Error('JSZip 脚本加载失败'));
            document.head.appendChild(script);
          });
        }
        return globalScope.JSZip;
      }
    } catch (error) {
      console.error('Failed to load JSZip:', error);
    }
    return null;
  })();
  return jsZipPromise;
}

async function extractDocumentXmlUsingJsZip(bytes, originalError) {
  const JSZip = await ensureJsZip();
  if (!JSZip) {
    console.error('JSZip is not available, cannot process ZIP file.');
    const message = typeof originalError?.message === 'string' ? originalError.message : '';
    const isZipRelated = message.includes('ZIP') || message.includes('zip') || message.includes('解压');
    if (isZipRelated) {
      throw originalError;
    }
    return null;
  }
  try {
    const zip = await JSZip.loadAsync(bytes);
    const docFile = zip.file('word/document.xml');
    if (!docFile) {
      throw new Error('ZIP 包中缺少 word/document.xml');
    }
    return await docFile.async('text');
  } catch (error) {
    console.error('JSZip 解压失败:', error);
    // If JSZip fails, we throw the original error if it was zip-related
    if (originalError) {
      const message = typeof originalError?.message === 'string' ? originalError.message : '';
      if (message.includes('ZIP') || message.includes('zip') || message.includes('解压')) {
        throw originalError;
      }
    }
    throw error; // otherwise throw the new error from JSZip
  }
}
