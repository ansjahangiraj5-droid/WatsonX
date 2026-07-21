import { useEffect, useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';

/* ─────────────────────────── shared constants ─────────────────────────── */

const allowedExcelTypes = new Set([
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel'
]);
const allOption = 'ALL';
const dateRangeOptions = ['Today', 'Last 7 Days', 'Last 30 Days', 'ALL Time'];
const sessionStorageKey = 'watsonx-ticket-dashboard-state';
const requiredFieldAliases = {
  number: ['number'],
  assignmentGroup: ['assignment group'],
  assignedTo: ['assigned to'],
  priority: ['priority']
};
const optionalFieldAliases = {
  created: ['created', 'opened', 'open date', 'date created', 'creation date', 'date opened']
};

/* ────────────────────────── SLA table definition ───────────────────────── */

const SLA_ROWS = [
  { key: 'p1_response',    label: 'Severity 1 - Response Time',    target: '20 mins',          expectedPct: 99.90, priority: 1, type: 'response'    },
  { key: 'p1_resolution',  label: 'Severity 1 - Resolution Time',  target: '4 Hours',           expectedPct: 99.90, priority: 1, type: 'resolution'  },
  { key: 'p2_response',    label: 'Severity 2 - Response Time',    target: '4 Hours',           expectedPct: 99.90, priority: 2, type: 'response'    },
  { key: 'p2_resolution',  label: 'Severity 2 - Resolution Time',  target: '8 Hours',           expectedPct: 99.90, priority: 2, type: 'resolution'  },
  { key: 'p3_response',    label: 'Severity 3 - Response Time',    target: '4 Hours',           expectedPct: 95,    priority: 3, type: 'response'    },
  { key: 'p3_resolution',  label: 'Severity 3 - Resolution Time',  target: '3 Business days',  expectedPct: 95,    priority: 3, type: 'resolution'  },
  { key: 'p4_response',    label: 'Severity 4 - Response Time',    target: '8 Hours',           expectedPct: 95,    priority: 4, type: 'response'    },
  { key: 'p4_resolution',  label: 'Severity 4 - Resolution Time',  target: '6 Business days',  expectedPct: 95,    priority: 4, type: 'resolution'  },
];

/* ──────────────────────────── helper functions ─────────────────────────── */

function normalizeKey(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function getCellValue(value) {
  if (value === null || value === undefined || value === '') {
    return '-';
  }
  return String(value);
}

function formatNumber(value) {
  return new Intl.NumberFormat('en-US').format(value);
}

function formatShortDate(value) {
  return value.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
}

function findHeaderIndex(headers, aliases) {
  return headers.findIndex((header) => {
    const normalized = normalizeKey(header);
    return aliases.some((alias) => normalized === alias || normalized.includes(alias) || alias.includes(normalized));
  });
}

function buildBreakdown(items, keyName, valueName) {
  return Object.entries(items)
    .map(([key, total]) => ({ [keyName]: key, [valueName]: total }))
    .sort((left, right) => right[valueName] - left[valueName]);
}

function startOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function isWeekend(date) {
  const day = date.getDay();
  return day === 0 || day === 6;
}

function setBusinessBoundary(date, hour, minute = 0, second = 0, millisecond = 0) {
  const nextDate = new Date(date);
  nextDate.setHours(hour, minute, second, millisecond);
  return nextDate;
}

function moveToNextBusinessStart(date) {
  const nextDate = new Date(date);

  while (isWeekend(nextDate)) {
    nextDate.setDate(nextDate.getDate() + 1);
    nextDate.setHours(9, 0, 0, 0);
  }

  const dayStart = setBusinessBoundary(nextDate, 9);
  const dayEnd = setBusinessBoundary(nextDate, 17);

  if (nextDate < dayStart) {
    return dayStart;
  }

  if (nextDate >= dayEnd) {
    nextDate.setDate(nextDate.getDate() + 1);
    nextDate.setHours(9, 0, 0, 0);
    return moveToNextBusinessStart(nextDate);
  }

  return nextDate;
}

function calculateBusinessDays(startDate, endDate) {
  if (!startDate || !endDate || endDate < startDate) {
    return 0;
  }

  const current = startOfDay(startDate);
  const last = startOfDay(endDate);
  let total = 0;

  while (current <= last) {
    if (!isWeekend(current)) {
      total += 1;
    }
    current.setDate(current.getDate() + 1);
  }

  return total;
}

function calculateBusinessHours(startDate, endDate) {
  if (!startDate || !endDate || endDate <= startDate) {
    return 0;
  }

  let current = moveToNextBusinessStart(startDate);
  let totalMilliseconds = 0;

  while (current < endDate) {
    if (isWeekend(current)) {
      current = moveToNextBusinessStart(current);
      continue;
    }

    const dayEnd = setBusinessBoundary(current, 17);
    const segmentEnd = dayEnd < endDate ? dayEnd : endDate;

    if (segmentEnd > current) {
      totalMilliseconds += segmentEnd.getTime() - current.getTime();
    }

    if (segmentEnd >= endDate) {
      break;
    }

    const nextDay = new Date(current);
    nextDay.setDate(nextDay.getDate() + 1);
    nextDay.setHours(9, 0, 0, 0);
    current = moveToNextBusinessStart(nextDay);
  }

  return Math.max(totalMilliseconds / (1000 * 60 * 60), 0);
}

function getPriorityTarget(priority) {
  const normalizedPriority = normalizeKey(priority);

  if (normalizedPriority === '1' || normalizedPriority.includes('priority 1')) {
    return { limit: 4, unit: 'hours' };
  }

  if (normalizedPriority === '2' || normalizedPriority.includes('priority 2')) {
    return { limit: 8, unit: 'hours' };
  }

  if (normalizedPriority === '3' || normalizedPriority.includes('priority 3')) {
    return { limit: 3, unit: 'days' };
  }

  if (normalizedPriority === '4' || normalizedPriority.includes('priority 4')) {
    return { limit: 6, unit: 'days' };
  }

  return { limit: 0, unit: 'days' };
}

function formatDuration(incident) {
  if (!incident.createdDate || !incident.resolvedDate) {
    return '-';
  }

  return incident.target.unit === 'hours'
    ? `${incident.resolutionHours.toFixed(1)} hrs`
    : `${incident.resolutionDays} days`;
}

function serializeDate(value) {
  return value instanceof Date ? value.toISOString() : value;
}

function parseDate(value) {
  if (!value || value === '-') {
    return null;
  }

  const parsedDate = new Date(value);
  return Number.isNaN(parsedDate.getTime()) ? null : parsedDate;
}

function readStoredState() {
  if (typeof window === 'undefined') {
    return null;
  }

  const savedState = window.sessionStorage.getItem(sessionStorageKey);

  if (!savedState) {
    return null;
  }

  try {
    return JSON.parse(savedState);
  } catch {
    return null;
  }
}

function createIncidentConfig(headers) {
  if (headers.length === 0) {
    return null;
  }

  const config = {
    numberIndex: findHeaderIndex(headers, requiredFieldAliases.number),
    assignmentGroupIndex: findHeaderIndex(headers, requiredFieldAliases.assignmentGroup),
    assignedToIndex: findHeaderIndex(headers, requiredFieldAliases.assignedTo),
    priorityIndex: findHeaderIndex(headers, requiredFieldAliases.priority),
    createdIndex: findHeaderIndex(headers, optionalFieldAliases.created),
    breachedIndex: findHeaderIndex(headers, ['has breached']),
    categoryIndex: findHeaderIndex(headers, ['category']),
    slaIndex: findHeaderIndex(headers, ['sla definition']),
    stageIndex: findHeaderIndex(headers, ['stage']),
    resolvedIndex: findHeaderIndex(headers, ['resolved'])
  };

  if (Object.values(requiredFieldAliases).some((aliases) => findHeaderIndex(headers, aliases) === -1)) {
    return null;
  }

  return config;
}

function buildIncident(row, index, incidentConfig, commentMap = {}) {
  const createdDate = incidentConfig.createdIndex >= 0 ? parseDate(row[incidentConfig.createdIndex]) : null;
  const resolvedDate = incidentConfig.resolvedIndex >= 0 ? parseDate(row[incidentConfig.resolvedIndex]) : null;
  const priority = incidentConfig.priorityIndex >= 0 ? getCellValue(row[incidentConfig.priorityIndex]) : '-';
  const target = getPriorityTarget(priority);
  const resolutionDays = calculateBusinessDays(createdDate, resolvedDate);
  const resolutionHours = calculateBusinessHours(createdDate, resolvedDate);
  const actualDuration = target.unit === 'hours' ? resolutionHours : resolutionDays;
  const withinTarget = Boolean(resolvedDate) && target.limit > 0 && actualDuration <= target.limit;
  const number = getCellValue(row[incidentConfig.numberIndex]);

  return {
    id: `${number}-${index}`,
    number,
    assignmentGroup: incidentConfig.assignmentGroupIndex >= 0 ? getCellValue(row[incidentConfig.assignmentGroupIndex]) : '-',
    assignedTo: incidentConfig.assignedToIndex >= 0 ? getCellValue(row[incidentConfig.assignedToIndex]) : '-',
    priority,
    created: incidentConfig.createdIndex >= 0 ? getCellValue(row[incidentConfig.createdIndex]) : '-',
    createdDate,
    hasBreached: incidentConfig.breachedIndex >= 0 ? getCellValue(row[incidentConfig.breachedIndex]) : '-',
    category: incidentConfig.categoryIndex >= 0 ? getCellValue(row[incidentConfig.categoryIndex]) : '-',
    slaDefinition: incidentConfig.slaIndex >= 0 ? getCellValue(row[incidentConfig.slaIndex]) : '-',
    stage: incidentConfig.stageIndex >= 0 ? getCellValue(row[incidentConfig.stageIndex]) : '-',
    resolved: incidentConfig.resolvedIndex >= 0 ? getCellValue(row[incidentConfig.resolvedIndex]) : '-',
    resolvedDate,
    resolutionDays,
    resolutionHours,
    actualDuration,
    target,
    withinTarget,
    comment: commentMap[number] || '',
    isCommentEditing: false,
    originalRow: row
  };
}

/* ──────────────────────── breach analysis helpers ──────────────────────── */

/**
 * Reads an Excel file and returns { headers, rows }.
 */
async function readExcelFile(file) {
  const fileBuffer = await file.arrayBuffer();
  const workbook = XLSX.read(fileBuffer, { type: 'array', cellDates: true });
  const firstSheetName = workbook.SheetNames[0];

  if (!firstSheetName) {
    throw new Error('The uploaded Excel file does not contain any sheets.');
  }

  const worksheet = workbook.Sheets[firstSheetName];
  const sheetData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
  const nonEmptyRows = sheetData.filter((row) => Array.isArray(row) && row.some((cell) => String(cell).trim() !== ''));

  if (nonEmptyRows.length === 0) {
    throw new Error('The uploaded Excel file is empty.');
  }

  const headers = nonEmptyRows[0].map((h, i) => String(h).trim() || `Column ${i + 1}`);
  const rows = nonEmptyRows.slice(1);
  return { headers, rows };
}

/**
 * Extracts priority number (1-5) from a cell value.
 * Handles: "1", "P1", "Priority 1", "1 - Critical", etc.
 */
function extractPriorityNumber(value) {
  const str = normalizeKey(String(value ?? ''));
  const match = str.match(/\b([1-5])\b/);
  return match ? parseInt(match[1], 10) : null;
}

/**
 * Count tickets per priority from Monthly Incident Count sheet.
 * Finds the "Priority" column and tallies rows.
 */
function countByPriority(headers, rows) {
  const priorityIdx = findHeaderIndex(headers, ['priority']);

  if (priorityIdx === -1) {
    return null;
  }

  const counts = { 1: 0, 2: 0, 3: 0, 4: 0, 5: 0 };

  for (const row of rows) {
    const p = extractPriorityNumber(row[priorityIdx]);
    if (p !== null && counts[p] !== undefined) {
      counts[p] += 1;
    }
  }

  return counts;
}

/**
 * Detect whether an SLA definition cell refers to response or resolution.
 * Returns 'response', 'resolution', or null if unrecognisable.
 */
function extractSlaType(value) {
  const str = normalizeKey(String(value ?? ''));
  if (str.includes('resolution') || str.includes('resolve')) return 'resolution';
  if (str.includes('response') || str.includes('respond')) return 'response';
  return null;
}

/**
 * Count breached tickets per priority AND per SLA type (response / resolution)
 * from the Breached Sheet.
 *
 * Requires a Priority column.
 * If an "SLA Definition" column is also present the counts are split by type;
 * otherwise the same count is used for both response and resolution rows.
 *
 * Returns: { 1: { response: n, resolution: n }, 2: { … }, … }
 * or null when no Priority column is found.
 */
function countBreachedByPriority(headers, rows) {
  const priorityIdx = findHeaderIndex(headers, ['priority']);

  if (priorityIdx === -1) {
    return null;
  }

  const slaDefIdx = findHeaderIndex(headers, ['sla definition', 'sla type', 'sla name', 'type']);

  // initialise counts for P1-P5, both types
  const counts = {};
  for (let p = 1; p <= 5; p++) {
    counts[p] = { response: 0, resolution: 0 };
  }

  for (const row of rows) {
    const p = extractPriorityNumber(row[priorityIdx]);
    if (p === null || !counts[p]) continue;

    if (slaDefIdx !== -1) {
      const slaType = extractSlaType(row[slaDefIdx]);
      if (slaType === 'response') {
        counts[p].response += 1;
      } else if (slaType === 'resolution') {
        counts[p].resolution += 1;
      } else {
        // SLA Definition present but not classifiable — count for both
        counts[p].response += 1;
        counts[p].resolution += 1;
      }
    } else {
      // No SLA Definition column — apply to both types
      counts[p].response += 1;
      counts[p].resolution += 1;
    }
  }

  return counts;
}

/**
 * Compute SLA achieved % for a given priority.
 * formula: (total - breached) / total * 100
 */
function computeSlaAchieved(total, breached) {
  if (!total || total === 0) {
    return null; // no data
  }
  const achieved = Math.max(total - breached, 0);
  return (achieved / total) * 100;
}

/* ═══════════════════════════════════════════════════════════════════════════
   HOME PAGE
═══════════════════════════════════════════════════════════════════════════ */

function HomePage({ onNavigate }) {
  return (
    <main className="page homePage">
      <div className="homeHero">
        <div className="brandBadge homeHeroBadge">WX</div>
        <p className="eyebrow" style={{ color: '#93c5fd', textAlign: 'center' }}>Service Operations</p>
        <h1 style={{ textAlign: 'center', marginTop: 8 }}>Ticket Command Center</h1>
        <p style={{ color: '#cbd5e1', textAlign: 'center', marginTop: 12, maxWidth: '52ch' }}>
          Select a module below to get started.
        </p>
      </div>

      <div className="homeCardGrid">
        <button type="button" className="homeModuleCard" onClick={() => onNavigate('daily')}>
          <div className="homeModuleIcon homeModuleIconBlue">DT</div>
          <h2 className="homeModuleTitle">Daily Tracker</h2>
          <p className="homeModuleDesc">
            Upload your service-ticket Excel workbook and explore interactive charts for assignment
            groups, assignees, resolution performance and ticket-creation trends with live filtering.
          </p>
          <span className="homeModuleCta">Open Daily Tracker →</span>
        </button>

        <button type="button" className="homeModuleCard" onClick={() => onNavigate('breach')}>
          <div className="homeModuleIcon homeModuleIconPurple">BA</div>
          <h2 className="homeModuleTitle">Breach Analysis</h2>
          <p className="homeModuleDesc">
            Upload your Monthly Incident Count and Breached Sheet to get a full SLA achievement
            summary broken down by severity and SLA type.
          </p>
          <span className="homeModuleCta">Open Breach Analysis →</span>
        </button>
      </div>
    </main>
  );
}

/* ═══════════════════════════════════════════════════════════════════════════
   BREACH ANALYSIS PAGE
═══════════════════════════════════════════════════════════════════════════ */

const breachStorageKey = 'watsonx-breach-analysis-state';

function readBreachStoredState() {
  if (typeof window === 'undefined') return null;
  const saved = window.sessionStorage.getItem(breachStorageKey);
  if (!saved) return null;
  try { return JSON.parse(saved); } catch { return null; }
}

function BreachAnalysisPage({ onBack }) {
  const countFileRef = useRef(null);
  const breachFileRef = useRef(null);

  const initialBreachState = useMemo(() => readBreachStoredState(), []);

  const [countFileName, setCountFileName] = useState(initialBreachState?.countFileName || '');
  const [countError, setCountError] = useState('');
  const [priorityCounts, setPriorityCounts] = useState(initialBreachState?.priorityCounts || null);

  const [breachFileName, setBreachFileName] = useState(initialBreachState?.breachFileName || '');
  const [breachError, setBreachError] = useState('');
  const [breachCounts, setBreachCounts] = useState(initialBreachState?.breachCounts || null);

  /* ── persist to session storage whenever data changes ── */

  useEffect(() => {
    if (typeof window === 'undefined') return;
    if (!priorityCounts && !breachCounts) {
      window.sessionStorage.removeItem(breachStorageKey);
      return;
    }
    window.sessionStorage.setItem(
      breachStorageKey,
      JSON.stringify({ countFileName, priorityCounts, breachFileName, breachCounts })
    );
  }, [countFileName, priorityCounts, breachFileName, breachCounts]);

  /* ── upload handlers ── */

  const handleCountUpload = async (event) => {
    const file = event.target.files?.[0];
    event.target.value = '';

    if (!file) return;

    const hasExcelExtension = /\.(xlsx|xls)$/i.test(file.name);
    const hasAllowedMimeType = !file.type || allowedExcelTypes.has(file.type);

    if (!hasExcelExtension || !hasAllowedMimeType) {
      setCountError('Only Excel files (.xlsx or .xls) are allowed.');
      setCountFileName('');
      setPriorityCounts(null);
      return;
    }

    try {
      const { headers, rows } = await readExcelFile(file);
      const counts = countByPriority(headers, rows);

      if (!counts) {
        setCountError('Could not find a "Priority" column in this file. Please verify the column header.');
        setCountFileName(file.name);
        setPriorityCounts(null);
        return;
      }

      setCountFileName(file.name);
      setCountError('');
      setPriorityCounts(counts);
    } catch (err) {
      setCountError(err.message || 'Unable to read the file.');
      setCountFileName('');
      setPriorityCounts(null);
    }
  };

  const handleClearCount = () => {
    setCountFileName('');
    setCountError('');
    setPriorityCounts(null);
    if (countFileRef.current) countFileRef.current.value = '';
  };

  const handleBreachUpload = async (event) => {
    const file = event.target.files?.[0];
    event.target.value = '';

    if (!file) return;

    const hasExcelExtension = /\.(xlsx|xls)$/i.test(file.name);
    const hasAllowedMimeType = !file.type || allowedExcelTypes.has(file.type);

    if (!hasExcelExtension || !hasAllowedMimeType) {
      setBreachError('Only Excel files (.xlsx or .xls) are allowed.');
      setBreachFileName('');
      setBreachCounts(null);
      return;
    }

    try {
      const { headers, rows } = await readExcelFile(file);
      const counts = countBreachedByPriority(headers, rows);

      if (!counts) {
        setBreachError('Could not find a "Priority" column in this file. Please verify the column header.');
        setBreachFileName(file.name);
        setBreachCounts(null);
        return;
      }

      setBreachFileName(file.name);
      setBreachError('');
      setBreachCounts(counts);
    } catch (err) {
      setBreachError(err.message || 'Unable to read the file.');
      setBreachFileName('');
      setBreachCounts(null);
    }
  };

  const handleClearBreach = () => {
    setBreachFileName('');
    setBreachError('');
    setBreachCounts(null);
    if (breachFileRef.current) breachFileRef.current.value = '';
  };

  /* ── derived SLA table ── */

  const slaTableRows = useMemo(() => {
    return SLA_ROWS.map((row) => {
      const total = priorityCounts ? (priorityCounts[row.priority] ?? 0) : null;
      // breachCounts[p] is now { response, resolution } — pick the right sub-type
      const breached = breachCounts
        ? (breachCounts[row.priority]?.[row.type] ?? 0)
        : null;
      const achievedPct = total !== null && breached !== null ? computeSlaAchieved(total, breached) : null;

      return { ...row, total, breached, achievedPct };
    });
  }, [priorityCounts, breachCounts]);

  const handleExportSlaTable = () => {
    const headers = ['SLA Definition', 'Target', 'Expected SLA %', 'SLA Achieved %'];
    const dataRows = slaTableRows.map((row) => [
      row.label,
      row.target,
      `${row.expectedPct}%`,
      row.achievedPct === null ? '—' : `${row.achievedPct.toFixed(2)}%`
    ]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet([headers, ...dataRows]), 'SLA Achievement');
    XLSX.writeFile(workbook, 'sla-achievement-summary.xlsx');
  };

  /* ── render ── */

  return (
    <main className="page">
      {/* top nav */}
      <div className="baNavBar">
        <button type="button" className="actionButton ghostButton baBackBtn" onClick={onBack}>
          ← Back to Home
        </button>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <div className="brandBadge" style={{ width: 40, height: 40, borderRadius: 12, fontSize: '0.78rem' }}>WX</div>
          <div>
            <p className="eyebrow" style={{ color: '#93c5fd' }}>Breach Analysis</p>
            <h2 style={{ color: '#ffffff', fontSize: '1.1rem' }}>SLA Achievement Report</h2>
          </div>
        </div>
      </div>

      {/* upload section */}
      <section className="baUploadGrid">
        {/* Card 1 – Monthly Incident Count */}
        <div className="baUploadCard">
          <div className="baUploadCardHeader">
            <span className="baUploadBadge baUploadBadgeBlue">1</span>
            <div>
              <strong style={{ color: '#ffffff' }}>Monthly Incident Count</strong>
              <p style={{ color: '#94a3b8', fontSize: '0.88rem', marginTop: 4 }}>
                Used to calculate the total number of tickets per priority.
              </p>
            </div>
          </div>

          <input ref={countFileRef} type="file" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={handleCountUpload} />
          <div style={{ color: '#cbd5e1', fontSize: '0.88rem', minHeight: 20 }}>
            {countFileName ? `📄 ${countFileName}` : 'No file selected'}
          </div>

          {countError && <p className="bannerError" style={{ marginTop: 6 }}>{countError}</p>}

          <div className="buttonRow">
            <button type="button" className="actionButton primaryButton" onClick={() => countFileRef.current?.click()}>
              Choose file
            </button>
            {priorityCounts && (
              <button type="button" className="actionButton ghostButton" onClick={handleClearCount}>
                Clear
              </button>
            )}
          </div>

          {/* priority count summary */}
          {priorityCounts && (
            <div className="baPrioritySummary">
              <p style={{ color: '#93c5fd', fontWeight: 700, marginBottom: 8 }}>Ticket count by priority</p>
              {[1, 2, 3, 4, 5].map((p) => (
                <div key={p} className="baPriorityRow">
                  <span className={`baPriorityBadge baPriority${p}`}>P{p}</span>
                  <span style={{ color: '#e2e8f0' }}>Total no of P{p}</span>
                  <span style={{ color: '#ffffff', fontWeight: 700 }}>{priorityCounts[p]}</span>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Card 2 – Breached Sheet */}
        <div className="baUploadCard">
          <div className="baUploadCardHeader">
            <span className="baUploadBadge baUploadBadgePurple">2</span>
            <div>
              <strong style={{ color: '#ffffff' }}>Breached Sheet</strong>
              <p style={{ color: '#94a3b8', fontSize: '0.88rem', marginTop: 4 }}>
                Contains the breached tickets — used to compute the SLA achieved %.
              </p>
            </div>
          </div>

          <input ref={breachFileRef} type="file" accept=".xlsx,.xls" style={{ display: 'none' }} onChange={handleBreachUpload} />
          <div style={{ color: '#cbd5e1', fontSize: '0.88rem', minHeight: 20 }}>
            {breachFileName ? `📄 ${breachFileName}` : 'No file selected'}
          </div>

          {breachError && <p className="bannerError" style={{ marginTop: 6 }}>{breachError}</p>}

          <div className="buttonRow">
            <button type="button" className="actionButton primaryButton" onClick={() => breachFileRef.current?.click()}>
              Choose file
            </button>
            {breachCounts && (
              <button type="button" className="actionButton ghostButton" onClick={handleClearBreach}>
                Clear
              </button>
            )}
          </div>

          {/* breached count summary */}
          {breachCounts && (
            <div className="baPrioritySummary">
              <p style={{ color: '#c4b5fd', fontWeight: 700, marginBottom: 8 }}>Breached count by priority &amp; type</p>
              {[1, 2, 3, 4, 5].map((p) => (
                <div key={p} style={{ display: 'flex', flexDirection: 'column', gap: 4 }}>
                  <div className="baPriorityRow">
                    <span className={`baPriorityBadge baPriority${p}`}>P{p}</span>
                    <span style={{ color: '#e2e8f0' }}>Response breached</span>
                    <span style={{ color: '#ffffff', fontWeight: 700 }}>{breachCounts[p]?.response ?? 0}</span>
                  </div>
                  <div className="baPriorityRow">
                    <span className={`baPriorityBadge baPriority${p}`} style={{ opacity: 0 }}>P{p}</span>
                    <span style={{ color: '#e2e8f0' }}>Resolution breached</span>
                    <span style={{ color: '#ffffff', fontWeight: 700 }}>{breachCounts[p]?.resolution ?? 0}</span>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      </section>

      {/* SLA table – shows when at least one file is loaded */}
      {(priorityCounts || breachCounts) && (
        <section className="tableCard" style={{ marginTop: 24, borderRadius: 24, padding: 28 }}>
          <div className="tableHeader tableActions fixedHeader">
            <div>
              <h2 style={{ color: '#0f172a' }}>SLA Achievement Summary</h2>
              <p style={{ color: '#475569', fontSize: '0.9rem', marginTop: 4 }}>
                Formula: SLA Achieved % = (Total − Breached) ÷ Total × 100
                {(!priorityCounts || !breachCounts) && (
                  <span style={{ color: '#d97706', marginLeft: 8 }}>
                    (Upload both files to see full results)
                  </span>
                )}
              </p>
            </div>
            <button type="button" className="actionButton exportButton" onClick={handleExportSlaTable}>
              Download Excel
            </button>
          </div>

          <div className="tableWrapper">
            <table className="slaTable">
              <thead>
                <tr>
                  <th>SLA Definition</th>
                  <th>Target</th>
                  <th>Expected SLA %</th>
                  <th>SLA Achieved %</th>
                </tr>
              </thead>
              <tbody>
                {slaTableRows.map((row) => {
                  const achievedDisplay = row.achievedPct === null
                    ? '—'
                    : `${row.achievedPct.toFixed(2)}%`;

                  return (
                    <tr key={row.key}>
                      <td style={{ fontWeight: 600 }}>{row.label}</td>
                      <td>{row.target}</td>
                      <td>{row.expectedPct}%</td>
                      <td style={{ fontWeight: 700 }}>{achievedDisplay}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </section>
      )}
    </main>
  );
}

/* ═══════════════════════════════════════════════════════════════════════════
   DAILY TRACKER PAGE  (original App — zero logic changes)
═══════════════════════════════════════════════════════════════════════════ */

function DailyTrackerPage({ onBack }) {
  const fileInputRef = useRef(null);
  const dashboardExportRef = useRef(null);
  const initialState = useMemo(() => readStoredState(), []);
  const [fileName, setFileName] = useState(initialState?.fileName || '');
  const [headers, setHeaders] = useState(initialState?.headers || []);
  const [rows, setRows] = useState(initialState?.rows || []);
  const [sheetName, setSheetName] = useState(initialState?.sheetName || '');
  const [error, setError] = useState('');
  const [selectedAssignmentGroup, setSelectedAssignmentGroup] = useState(initialState?.selectedAssignmentGroup || allOption);
  const [selectedAssignee, setSelectedAssignee] = useState(initialState?.selectedAssignee || allOption);
  const [selectedBreachStatus, setSelectedBreachStatus] = useState(initialState?.selectedBreachStatus || allOption);
  const [selectedDateRange, setSelectedDateRange] = useState(initialState?.selectedDateRange || 'ALL Time');
  const [commentsByNumber, setCommentsByNumber] = useState(initialState?.commentsByNumber || {});
  const [editingComments, setEditingComments] = useState({});
  // draft text per ticket (only what the user is currently typing — NOT the saved history)
  const [draftComments, setDraftComments] = useState({});
  // toast: { id, ticketNumber }
  const [savedToast, setSavedToast] = useState(null);
  const toastTimerRef = useRef(null);
  // custom columns: [{ id, label }]
  const [customColumns, setCustomColumns] = useState(initialState?.customColumns || []);
  // custom column values: { [ticketNumber]: { [colId]: string } }
  const [customColValues, setCustomColValues] = useState(initialState?.customColValues || {});
  // which custom cells are being edited: { [ticketNumber-colId]: bool }
  const [editingCustomCells, setEditingCustomCells] = useState({});
  // custom cell saved toast
  const [customCellToast, setCustomCellToast] = useState(null);
  const customCellToastTimerRef = useRef(null);
  // add-column dialog
  const [showAddColDialog, setShowAddColDialog] = useState(false);
  const [newColLabel, setNewColLabel] = useState('');
  // fullscreen ticket preview
  const [previewFullscreen, setPreviewFullscreen] = useState(false);

  const resetData = (message = '', nextFileName = '') => {
    setFileName(nextFileName);
    setHeaders([]);
    setRows([]);
    setSheetName('');
    setSelectedAssignmentGroup(allOption);
    setSelectedAssignee(allOption);
    setSelectedBreachStatus(allOption);
    setSelectedDateRange('ALL Time');
    setCommentsByNumber({});
    setEditingComments({});
    setCustomColumns([]);
    setCustomColValues({});
    setEditingCustomCells({});
    setError(message);

    if (typeof window !== 'undefined') {
      window.sessionStorage.removeItem(sessionStorageKey);
    }
  };

  const handleClear = () => {
    resetData();

    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  // ── AI Chat State ────────────────────────────────────────
  const [chatInput, setChatInput] = useState('');
  const [chatResponse, setChatResponse] = useState('');
  const [isChatLoading, setIsChatLoading] = useState(false);
  const [chatError, setChatError] = useState('');

  const handleAskAI = async () => {
    console.log('User question:', chatInput);
    if (!chatInput.trim() || filteredIncidents.length === 0) return;
    setIsChatLoading(true);
    setChatError('');
    setChatResponse('');

    // LLM ke liye clean JSON banate hain taaki tokens bachein aur answer accurate aaye
    const ticketDataForAI = filteredIncidents.map((inc) => ({
      Number: inc.number,
      Priority: inc.priority,
      AssignmentGroup: inc.assignmentGroup,
      AssignedTo: inc.assignedTo,
      Status: inc.withinTarget ? 'Within Target' : 'Breached',
      Duration: `${inc.actualDuration} ${inc.target.unit}`,
      Comments: inc.comment || 'None'
    }));

    try {
      const response = await fetch('http://127.0.0.1:3001/api/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          question: chatInput,
          tickets: ticketDataForAI
        })
      });

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.error || 'Failed to fetch AI response.');
      }

      setChatResponse(data.answer);
    } catch (err) {
      setChatError(err.message);
    } finally {
      setIsChatLoading(false);
    }
  };

  const incidentConfig = useMemo(() => createIncidentConfig(headers), [headers]);

  const incidents = useMemo(() => {
    if (!incidentConfig) {
      return [];
    }

    return rows.map((row, index) => buildIncident(row, index, incidentConfig, commentsByNumber));
  }, [commentsByNumber, incidentConfig, rows]);

  useEffect(() => {
    if (typeof window === 'undefined') {
      return;
    }

    if (headers.length === 0 && rows.length === 0) {
      return;
    }

    window.sessionStorage.setItem(
      sessionStorageKey,
      JSON.stringify({
        fileName,
        headers,
        rows: rows.map((row) => row.map((cell) => serializeDate(cell))),
        sheetName,
        selectedAssignmentGroup,
        selectedAssignee,
        selectedBreachStatus,
        selectedDateRange,
        commentsByNumber,
        customColumns,
        customColValues
      })
    );
  }, [commentsByNumber, customColValues, customColumns, fileName, headers, rows, selectedAssignmentGroup, selectedAssignee, selectedBreachStatus, selectedDateRange, sheetName]);

  const assignmentGroups = useMemo(() => {
    return [allOption, ...Array.from(new Set(incidents.map((item) => item.assignmentGroup).filter((item) => item !== '-'))).sort()];
  }, [incidents]);

  const assignees = useMemo(() => {
    return [allOption, ...Array.from(new Set(incidents.map((item) => item.assignedTo).filter((item) => item !== '-'))).sort()];
  }, [incidents]);

  const breachStatuses = useMemo(() => {
    return [allOption, ...Array.from(new Set(incidents.map((item) => item.hasBreached).filter((item) => item !== '-'))).sort()];
  }, [incidents]);

  const filteredIncidents = useMemo(() => {
    const today = startOfDay(new Date());

    return incidents.filter((incident) => {
      const matchesAssignmentGroup = selectedAssignmentGroup === allOption || incident.assignmentGroup === selectedAssignmentGroup;
      const matchesAssignee = selectedAssignee === allOption || incident.assignedTo === selectedAssignee;
      const matchesBreachStatus = selectedBreachStatus === allOption || incident.hasBreached === selectedBreachStatus;

      let matchesDateRange = true;

      if (selectedDateRange !== 'ALL Time') {
        if (!incident.createdDate) {
          matchesDateRange = false;
        } else {
          const createdDay = startOfDay(incident.createdDate);
          const diffDays = Math.floor((today.getTime() - createdDay.getTime()) / (1000 * 60 * 60 * 24));

          if (selectedDateRange === 'Today') {
            matchesDateRange = diffDays === 0;
          }

          if (selectedDateRange === 'Last 7 Days') {
            matchesDateRange = diffDays >= 0 && diffDays < 7;
          }

          if (selectedDateRange === 'Last 30 Days') {
            matchesDateRange = diffDays >= 0 && diffDays < 30;
          }
        }
      }

      return matchesAssignmentGroup && matchesAssignee && matchesBreachStatus && matchesDateRange;
    });
  }, [incidents, selectedAssignmentGroup, selectedAssignee, selectedBreachStatus, selectedDateRange]);

  const metrics = useMemo(() => {
    const breachedCount = filteredIncidents.filter((incident) => !incident.withinTarget && incident.resolvedDate).length;
    const uniqueGroups = new Set(filteredIncidents.map((incident) => incident.assignmentGroup).filter((item) => item !== '-')).size;
    const resolvedWithinTarget = filteredIncidents.filter((incident) => incident.resolvedDate && incident.withinTarget).length;
    const resolvedOutOfTarget = filteredIncidents.filter((incident) => incident.resolvedDate && !incident.withinTarget).length;

    return {
      totalIncidents: filteredIncidents.length,
      breachedCount,
      uniqueGroups,
      resolvedWithinTarget,
      resolvedOutOfTarget
    };
  }, [filteredIncidents]);

  const assignmentGroupBreakdown = useMemo(() => {
    const grouped = filteredIncidents.reduce((accumulator, incident) => {
      accumulator[incident.assignmentGroup] = (accumulator[incident.assignmentGroup] || 0) + 1;
      return accumulator;
    }, {});

    return buildBreakdown(grouped, 'assignmentGroup', 'total');
  }, [filteredIncidents]);

  const assigneeBreakdown = useMemo(() => {
    const grouped = filteredIncidents.reduce((accumulator, incident) => {
      accumulator[incident.assignedTo] = (accumulator[incident.assignedTo] || 0) + 1;
      return accumulator;
    }, {});

    return buildBreakdown(grouped, 'assignedTo', 'total');
  }, [filteredIncidents]);

  const resolutionDurationBreakdown = useMemo(() => {
    return filteredIncidents
      .filter((incident) => incident.resolvedDate)
      .map((incident) => ({
        number: incident.number,
        value: incident.actualDuration,
        target: incident.target.limit,
        unit: incident.target.unit,
        withinTarget: incident.withinTarget
      }));
  }, [filteredIncidents]);

  const recentIncidentTrend = useMemo(() => {
    const grouped = filteredIncidents.reduce((accumulator, incident) => {
      if (!incident.createdDate) {
        return accumulator;
      }

      const key = incident.createdDate.toISOString().slice(0, 10);
      accumulator[key] = (accumulator[key] || 0) + 1;
      return accumulator;
    }, {});

    const orderedItems = Object.entries(grouped)
      .map(([date, total]) => ({ date, total }))
      .sort((left, right) => left.date.localeCompare(right.date));

    if (selectedDateRange === 'Today') {
      return orderedItems.slice(-1);
    }

    if (selectedDateRange === 'Last 7 Days') {
      return orderedItems.slice(-7);
    }

    if (selectedDateRange === 'Last 30 Days') {
      return orderedItems.slice(-30);
    }

    return orderedItems;
  }, [filteredIncidents, selectedDateRange]);

  const maxAssignmentGroupValue = Math.max(...assignmentGroupBreakdown.map((item) => item.total), 1);
  const maxAssigneeValue = Math.max(...assigneeBreakdown.map((item) => item.total), 1);
  const maxResolutionValue = Math.max(...resolutionDurationBreakdown.map((item) => Math.max(item.value, item.target)), 1);
  const maxTrendValue = Math.max(...recentIncidentTrend.map((item) => item.total), 1);

  const handleFileUpload = async (event) => {
    const selectedFile = event.target.files?.[0];

    if (!selectedFile) {
      return;
    }

    const hasExcelExtension = /\.(xlsx|xls)$/i.test(selectedFile.name);
    const hasAllowedMimeType = !selectedFile.type || allowedExcelTypes.has(selectedFile.type);

    if (!hasExcelExtension || !hasAllowedMimeType) {
      resetData('Only Excel files (.xlsx or .xls) are allowed. Please upload a valid Excel file.');
      event.target.value = '';
      return;
    }

    try {
      const fileBuffer = await selectedFile.arrayBuffer();
      const workbook = XLSX.read(fileBuffer, { type: 'array', cellDates: true });
      const firstSheetName = workbook.SheetNames[0];

      if (!firstSheetName) {
        resetData('The uploaded Excel file does not contain any sheets.', selectedFile.name);
        return;
      }

      const worksheet = workbook.Sheets[firstSheetName];
      const sheetData = XLSX.utils.sheet_to_json(worksheet, {
        header: 1,
        defval: ''
      });

      const nonEmptyRows = sheetData.filter((row) => Array.isArray(row) && row.some((cell) => String(cell).trim() !== ''));

      if (nonEmptyRows.length === 0) {
        resetData('The uploaded Excel file is empty.', selectedFile.name);
        return;
      }

      const computedColumns = ['sla duration', 'sla status', 'comments'];
      const rawHeaders = nonEmptyRows[0].map((header, index) => String(header).trim() || `Column ${index + 1}`);
      const computedIndexes = new Set(rawHeaders.map((h, i) => computedColumns.includes(normalizeKey(h)) ? i : -1).filter((i) => i !== -1));
      const nextHeaders = rawHeaders.filter((_, i) => !computedIndexes.has(i));
      const commentsColIndex = rawHeaders.findIndex((h) => normalizeKey(h) === 'comments');
      const nextRows = nonEmptyRows.slice(1).map((row) => nextHeaders.map((_, i) => {
        const originalIndex = rawHeaders.indexOf(nextHeaders[i]);
        return row[originalIndex] ?? '';
      }));
      const nextConfig = createIncidentConfig(nextHeaders);

      if (!nextConfig) {
        resetData(
          'The Excel file must contain Number, Assignment Group, Assigned To, and Priority columns. Extra fields are allowed.',
          selectedFile.name
        );
        return;
      }

      const restoredComments = {};
      if (commentsColIndex >= 0) {
        const numberColIndex = rawHeaders.findIndex((h) => normalizeKey(h) === 'number');
        nonEmptyRows.slice(1).forEach((row) => {
          const num = String(row[numberColIndex] ?? '').trim();
          const comment = String(row[commentsColIndex] ?? '').trim();
          if (num && comment) restoredComments[num] = comment;
        });
      }

      setFileName(selectedFile.name);
      setSheetName(firstSheetName);
      setError('');
      setSelectedAssignmentGroup(allOption);
      setSelectedAssignee(allOption);
      setSelectedBreachStatus(allOption);
      setSelectedDateRange('ALL Time');
      setCommentsByNumber(restoredComments);
      setEditingComments({});
      setHeaders(nextHeaders);
      setRows(nextRows);
    } catch (uploadError) {
      resetData('Unable to read the uploaded file. Please upload a valid Excel file.', selectedFile.name);
    } finally {
      event.target.value = '';
    }
  };

  const toggleCommentEditor = (ticketNumber) => {
    setEditingComments((currentState) => {
      const nowOpen = !currentState[ticketNumber];
      if (nowOpen) {
        // always start with a blank draft — saved history stays separate
        setDraftComments((prev) => ({ ...prev, [ticketNumber]: '' }));
      }
      return { ...currentState, [ticketNumber]: nowOpen };
    });
  };

  const handleCommentDraftChange = (ticketNumber, value) => {
    setDraftComments((prev) => ({ ...prev, [ticketNumber]: value }));
  };

  const handleSaveComment = (ticketNumber) => {
    const draft = (draftComments[ticketNumber] || '').trim();
    if (draft) {
      // append the new entry to the saved history with a newline separator
      setCommentsByNumber((prev) => {
        const existing = (prev[ticketNumber] || '').trimEnd();
        return { ...prev, [ticketNumber]: existing ? `${existing}\n${draft}` : draft };
      });
    }
    // clear the draft and close the editor
    setDraftComments((prev) => ({ ...prev, [ticketNumber]: '' }));
    setEditingComments((currentState) => ({ ...currentState, [ticketNumber]: false }));
    // show toast
    if (toastTimerRef.current) clearTimeout(toastTimerRef.current);
    setSavedToast(ticketNumber);
    toastTimerRef.current = setTimeout(() => setSavedToast(null), 2200);
  };

  // ── custom columns ────────────────────────────────────────
  const handleAddColumn = () => {
    const label = newColLabel.trim();
    if (!label) return;
    const id = `cc_${Date.now()}`;
    setCustomColumns((prev) => [...prev, { id, label }]);
    setNewColLabel('');
    setShowAddColDialog(false);
  };

  const handleRemoveColumn = (colId) => {
    setCustomColumns((prev) => prev.filter((c) => c.id !== colId));
    setCustomColValues((prev) => {
      const next = { ...prev };
      Object.keys(next).forEach((num) => {
        const copy = { ...next[num] };
        delete copy[colId];
        next[num] = copy;
      });
      return next;
    });
  };

  const cellKey = (ticketNumber, colId) => `${ticketNumber}__${colId}`;

  const toggleCustomCell = (ticketNumber, colId) => {
    const key = cellKey(ticketNumber, colId);
    setEditingCustomCells((prev) => ({ ...prev, [key]: !prev[key] }));
  };

  const handleCustomCellChange = (ticketNumber, colId, value) => {
    setCustomColValues((prev) => ({
      ...prev,
      [ticketNumber]: { ...(prev[ticketNumber] || {}), [colId]: value }
    }));
  };

  const handleSaveCustomCell = (ticketNumber, colId) => {
    const key = cellKey(ticketNumber, colId);
    setEditingCustomCells((prev) => ({ ...prev, [key]: false }));
    if (customCellToastTimerRef.current) clearTimeout(customCellToastTimerRef.current);
    setCustomCellToast(key);
    customCellToastTimerRef.current = setTimeout(() => setCustomCellToast(null), 2200);
  };

  const handleExportTicketData = async () => {
    const extraHeaders = ['SLA Duration', 'SLA Status', 'Comments', ...customColumns.map((c) => c.label)];
    const worksheetData = [
      [...headers, ...extraHeaders],
      ...filteredIncidents.map((incident) => [
        ...headers.map((_, i) => getCellValue(incident.originalRow[i])),
        formatDuration(incident),
        incident.withinTarget ? 'Within target' : 'Breached',
        incident.comment || '',
        ...customColumns.map((c) => (customColValues[incident.number]?.[c.id] || ''))
      ])
    ];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.aoa_to_sheet(worksheetData), 'Ticket Data');

    if (window.showSaveFilePicker) {
      try {
        const fileHandle = await window.showSaveFilePicker({
          suggestedName: 'ticket-data-preview.xlsx',
          types: [{ description: 'Excel Workbook', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }]
        });
        const writable = await fileHandle.createWritable();
        const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
        await writable.write(new Blob([buffer], { type: 'application/octet-stream' }));
        await writable.close();
      } catch (err) {
        if (err.name !== 'AbortError') {
          XLSX.writeFile(workbook, 'ticket-data-preview.xlsx');
        }
      }
    } else {
      XLSX.writeFile(workbook, 'ticket-data-preview.xlsx');
    }
  };

  return (
    <main className="page">
      {/* back button bar */}
      <div className="baNavBar" style={{ marginBottom: 20 }}>
        <button type="button" className="actionButton ghostButton baBackBtn" onClick={onBack}>
          ← Back to Home
        </button>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <div className="brandBadge" style={{ width: 40, height: 40, borderRadius: 12, fontSize: '0.78rem' }}>WX</div>
          <div>
            <p className="eyebrow" style={{ color: '#93c5fd' }}>Daily Tracker</p>
            <h2 style={{ color: '#ffffff', fontSize: '1.1rem' }}>Ticket Command Center</h2>
          </div>
        </div>
      </div>

      <section className="appShell">
        <aside className="sidePanel">
          <div className="brandBlock">
            <div className="brandBadge">WX</div>
            <div>
              <p className="eyebrow">Service Operations</p>
              <h2>Ticket Command Center</h2>
            </div>
          </div>

          <div className="sideCard brightCard">
            <span className="sideLabel">Workbook</span>
            <strong>{fileName || 'No Excel file uploaded'}</strong>
            <small>{sheetName ? `Sheet: ${sheetName}` : 'Upload an Excel file to begin'}</small>
          </div>

          <div className="sideCard">
            <label htmlFor="assignment-group-filter" className="sideLabel">Assignment group</label>
            <select
              id="assignment-group-filter"
              className="filterSelect"
              value={selectedAssignmentGroup}
              onChange={(event) => setSelectedAssignmentGroup(event.target.value)}
              disabled={assignmentGroups.length === 1}
            >
              {assignmentGroups.map((group) => (
                <option key={group} value={group}>{group}</option>
              ))}
            </select>
          </div>

          <div className="sideCard">
            <label htmlFor="assignee-filter" className="sideLabel">Assigned to</label>
            <select
              id="assignee-filter"
              className="filterSelect"
              value={selectedAssignee}
              onChange={(event) => setSelectedAssignee(event.target.value)}
              disabled={assignees.length === 1}
            >
              {assignees.map((user) => (
                <option key={user} value={user}>{user}</option>
              ))}
            </select>
          </div>

          <div className="sideCard">
            <label htmlFor="breach-filter" className="sideLabel">Has breached</label>
            <select
              id="breach-filter"
              className="filterSelect"
              value={selectedBreachStatus}
              onChange={(event) => setSelectedBreachStatus(event.target.value)}
              disabled={breachStatuses.length === 1}
            >
              {breachStatuses.map((status) => (
                <option key={status} value={status}>{status}</option>
              ))}
            </select>
          </div>

          <div className="sideCard">
            <label htmlFor="date-range-filter" className="sideLabel">Created date range</label>
            <select
              id="date-range-filter"
              className="filterSelect"
              value={selectedDateRange}
              onChange={(event) => setSelectedDateRange(event.target.value)}
            >
              {dateRangeOptions.map((option) => (
                <option key={option} value={option}>{option}</option>
              ))}
            </select>
            <small>Choose Today, Last 7 Days, Last 30 Days, or ALL Time.</small>
          </div>

          <div className="sideCard compactList">
            <span className="sideLabel">Required fields</span>
            <span>Number</span>
            <span>Assignment Group</span>
            <span>Assigned To</span>
            <span>Priority</span>
          </div>
        </aside>

        <section className="dashboardPanel" ref={dashboardExportRef}>
          <section className="hero heroGradient">
            <div>
              <p className="eyebrow highlight">Operations dashboard</p>
              <h1>Excel-driven ticket analytics</h1>
              <p className="subtitle lightText">
                Upload a workbook containing service tickets and explore visible charts for assignment groups,
                assignees, resolution performance, and ticket creation trends with live filtering.
              </p>
            </div>

            <div className="uploadCard gradientCard">
              <span>Upload Excel file</span>
              <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFileUpload} style={{display:"none"}}/>
              <strong>{fileName || 'No Excel file selected'}</strong>
              <small>{sheetName ? `Sheet: ${sheetName}` : 'Accepted formats: .xlsx and .xls only'}</small>
              <div className="buttonRow">
                <button type="button" className="actionButton primaryButton" onClick={() => fileInputRef.current?.click()}>
                  Choose file
                </button>
                <button type="button" className="actionButton ghostButton" onClick={handleClear}>
                  Clear
                </button>
              </div>

              {/* ── inline filters shown only on mobile ── */}
              <div className="mobileFilters">
                <div className="mobileFilterRow">
                  <label htmlFor="m-assignment-group-filter" className="sideLabel">Assignment group</label>
                  <select
                    id="m-assignment-group-filter"
                    className="filterSelect"
                    value={selectedAssignmentGroup}
                    onChange={(e) => setSelectedAssignmentGroup(e.target.value)}
                    disabled={assignmentGroups.length === 1}
                  >
                    {assignmentGroups.map((g) => <option key={g} value={g}>{g}</option>)}
                  </select>
                </div>
                <div className="mobileFilterRow">
                  <label htmlFor="m-assignee-filter" className="sideLabel">Assigned to</label>
                  <select
                    id="m-assignee-filter"
                    className="filterSelect"
                    value={selectedAssignee}
                    onChange={(e) => setSelectedAssignee(e.target.value)}
                    disabled={assignees.length === 1}
                  >
                    {assignees.map((u) => <option key={u} value={u}>{u}</option>)}
                  </select>
                </div>
                <div className="mobileFilterRow">
                  <label htmlFor="m-breach-filter" className="sideLabel">Has breached</label>
                  <select
                    id="m-breach-filter"
                    className="filterSelect"
                    value={selectedBreachStatus}
                    onChange={(e) => setSelectedBreachStatus(e.target.value)}
                    disabled={breachStatuses.length === 1}
                  >
                    {breachStatuses.map((s) => <option key={s} value={s}>{s}</option>)}
                  </select>
                </div>
                <div className="mobileFilterRow">
                  <label htmlFor="m-date-range-filter" className="sideLabel">Created date range</label>
                  <select
                    id="m-date-range-filter"
                    className="filterSelect"
                    value={selectedDateRange}
                    onChange={(e) => setSelectedDateRange(e.target.value)}
                  >
                    {dateRangeOptions.map((o) => <option key={o} value={o}>{o}</option>)}
                  </select>
                </div>
              </div>
            </div>
          </section>

          {error ? <p className="bannerError">{error}</p> : null}

          <section className="metrics">
            <article className="metricCard pinkCard">
              <span>Total tickets</span>
              <strong>{formatNumber(metrics.totalIncidents)}</strong>
            </article>
            <article className="metricCard blueCard">
              <span>Breached</span>
              <strong>{formatNumber(metrics.breachedCount)}</strong>
            </article>
            <article className="metricCard purpleCard">
              <span>Within target</span>
              <strong>{formatNumber(metrics.resolvedWithinTarget)}</strong>
            </article>
            <article className="metricCard tealCard">
              <span>Out of target</span>
              <strong>{formatNumber(metrics.resolvedOutOfTarget)}</strong>
            </article>
          </section>

          {/* ──  AI Chat Assistant Block ── */}
          <section className="tableCard darkPanel" style={{ marginBottom: 20 }}>
            <div className="tableHeader">
              <div>
                <h2 style={{ color: '#fff' }}>✨ AI Ticket Analyst</h2>
                <p className="tableSubtitle" style={{ color: '#94a3b8' }}>Ask questions about the current filtered ticket data.</p>
              </div>
              <span className="pill">Powered by Gemini</span>
            </div>
            
            <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
              <textarea
                className="commentInput"
                style={{ minHeight: 60, width: '100%' }}
                placeholder="E.g., 'Which assignee has the most breached P1 tickets?' or 'Summarize the comments.'"
                value={chatInput}
                onChange={(e) => setChatInput(e.target.value)}
                onKeyDown={(e) => {
                  if (e.key === 'Enter' && !e.shiftKey) {
                    e.preventDefault();
                    handleAskAI();
                  }
                }}
              />
              <div className="buttonRow">
                <button 
                  type="button" 
                  className="actionButton primaryButton" 
                  onClick={handleAskAI}
                  disabled={isChatLoading || !chatInput.trim() || filteredIncidents.length === 0}
                >
                  {isChatLoading ? 'Analyzing Data...' : 'Ask AI'}
                </button>
              </div>

              {chatError && <p className="bannerError">{chatError}</p>}
              
              {chatResponse && (
                <div style={{ background: 'rgba(255,255,255,0.05)', padding: 16, borderRadius: 14, marginTop: 10, border: '1px solid rgba(255,255,255,0.1)' }}>
                  <strong style={{ color: '#22d3ee', display: 'block', marginBottom: 8 }}>AI Response:</strong>
                  <div style={{ color: '#f8fafc', whiteSpace: 'pre-wrap', lineHeight: 1.6 }}>{chatResponse}</div>
                </div>
              )}
            </div>
          </section>

          <section className="insightsGrid wideGrid">
            <article className="tableCard overviewCard darkPanel">
              <div className="tableHeader fixedHeader">
                <h2>Assignment group workload</h2>
                <span className="pill">{selectedAssignmentGroup}</span>
              </div>
              <div className="chartPanel cardScrollArea">
                {assignmentGroupBreakdown.length > 0 ? (
                  assignmentGroupBreakdown.map((item) => (
                    <div className="chartRow improvedChartRow" key={item.assignmentGroup}>
                      <div className="chartLabel lightText">{item.assignmentGroup}</div>
                      <div className="chartTrack darkTrack">
                        <div
                          className="chartBar gradientPink"
                          style={{ width: `${(item.total / maxAssignmentGroupValue) * 100}%` }}
                        />
                      </div>
                      <div className="chartValue lightText">{item.total}</div>
                    </div>
                  ))
                ) : (
                  <div className="emptyState lightText">Upload valid data to see assignment group trends.</div>
                )}
              </div>
            </article>

            <article className="tableCard overviewCard">
              <div className="tableHeader fixedHeader">
                <h2>Resolution time vs SLA</h2>
                <span className="pill neutral">Created to resolved</span>
              </div>
              <div className="chartPanel cardScrollArea">
                {resolutionDurationBreakdown.length > 0 ? (
                  resolutionDurationBreakdown.map((item) => (
                    <div className="chartRow improvedChartRow" key={item.number}>
                      <div className="chartLabel multiLineLabel">
                        <strong>{item.number}</strong>
                        <span>{item.unit === 'hours' ? `${item.target} hrs target` : `${item.target} day target`}</span>
                      </div>
                      <div className="chartTrack slaTrack">
                        <div
                          className={`chartBar ${item.withinTarget ? 'gradientGreen' : 'gradientOrange'}`}
                          style={{ width: `${(item.value / maxResolutionValue) * 100}%` }}
                        />
                      </div>
                      <div className="chartValue darkValue">{item.unit === 'hours' ? `${item.value.toFixed(1)}h` : `${item.value}d`}</div>
                    </div>
                  ))
                ) : (
                  <div className="emptyState">Resolution duration appears here for resolved tickets.</div>
                )}
              </div>
            </article>
          </section>

          <section className="insightsGrid">
            <article className="tableCard overviewCard">
              <div className="tableHeader fixedHeader">
                <h2>Assigned to distribution</h2>
                <span className="pill neutral">Filtered view</span>
              </div>
              <div className="chartPanel cardScrollArea">
                {assigneeBreakdown.length > 0 ? (
                  assigneeBreakdown.map((item) => (
                    <div className="chartRow improvedChartRow" key={item.assignedTo}>
                      <div className="chartLabel darkValue">{item.assignedTo}</div>
                      <div className="chartTrack">
                        <div
                          className="chartBar gradientBlue"
                          style={{ width: `${(item.total / maxAssigneeValue) * 100}%` }}
                        />
                      </div>
                      <div className="chartValue darkValue">{item.total}</div>
                    </div>
                  ))
                ) : (
                  <div className="emptyState">Assignee charts appear here after upload.</div>
                )}
              </div>
            </article>

            <article className="tableCard overviewCard summaryGradient">
              <div className="tableHeader fixedHeader trendHeader">
                <div>
                  <h2>Recent ticket creation</h2>
                  <p className="tableSubtitle darkSubtitle">Y-axis shows ticket count, X-axis shows created dates.</p>
                </div>
                <span className="pill neutral">{selectedDateRange}</span>
              </div>
              {recentIncidentTrend.length > 0 ? (
                <div className="axisChartCard cardScrollArea">
                  <div className="lineChart">
                    <div className="axisY lineAxisY">
                      {[maxTrendValue, Math.ceil(maxTrendValue / 2), 0].map((tick) => (
                        <span key={tick}>{tick}</span>
                      ))}
                    </div>
                    <div className="lineChartArea">
                      <svg className="lineSvg" viewBox={`0 0 ${Math.max(recentIncidentTrend.length - 1, 1) * 120 + 40} 220`} preserveAspectRatio="none">
                        {[40, 110, 180].map((y) => (
                          <line key={y} x1="0" y1={y} x2={Math.max(recentIncidentTrend.length - 1, 1) * 120 + 40} y2={y} className="lineGrid" />
                        ))}
                        <polyline
                          className="linePath"
                          points={recentIncidentTrend.map((item, index) => {
                            const x = index * 120 + 20;
                            const y = 180 - ((item.total / maxTrendValue) * 140);
                            return `${x},${y}`;
                          }).join(' ')}
                        />
                        {recentIncidentTrend.map((item, index) => {
                          const x = index * 120 + 20;
                          const y = 180 - ((item.total / maxTrendValue) * 140);

                          return (
                            <g key={item.date}>
                              <circle cx={x} cy={y} r="6" className="linePoint" />
                              <text x={x} y={y - 12} textAnchor="middle" className="lineValue">{item.total}</text>
                            </g>
                          );
                        })}
                      </svg>
                      <div className="lineLabels">
                        {recentIncidentTrend.map((item) => (
                          <span key={item.date}>{formatShortDate(new Date(item.date))}</span>
                        ))}
                      </div>
                    </div>
                  </div>
                </div>
              ) : (
                <div className="emptyState">Created-date activity appears here when the data is available.</div>
              )}
            </article>
          </section>

          <section className={`tableCard overviewCard${previewFullscreen ? ' previewFullscreen' : ''}`}>
            <div className="tableHeader tableActions fixedHeader">
              <div>
                <h2>Ticket data preview</h2>
                <p className="tableSubtitle darkSubtitle">
                  Showing {formatNumber(filteredIncidents.length)} records after applying the selected filters.
                </p>
              </div>
              <div className="pillButtonGroup">
                <span className="pill neutral">First worksheet</span>
                <button
                  type="button"
                  className="actionButton addColButton"
                  onClick={() => { setShowAddColDialog(true); setNewColLabel(''); }}
                  disabled={headers.length === 0}
                  title="Add a custom column to this table"
                >
                  + Add Column
                </button>
                <button type="button" className="actionButton exportButton" onClick={handleExportTicketData} disabled={filteredIncidents.length === 0}>
                  Download Excel
                </button>
                <button
                  type="button"
                  className="iconButton fullscreenToggleBtn"
                  onClick={() => setPreviewFullscreen((v) => !v)}
                  title={previewFullscreen ? 'Minimise preview' : 'Expand to full screen'}
                  aria-label={previewFullscreen ? 'Minimise preview' : 'Expand to full screen'}
                >
                  {previewFullscreen ? (
                    /* compress / minimise icon */
                    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                      <polyline points="4 14 10 14 10 20"/><polyline points="20 10 14 10 14 4"/>
                      <line x1="10" y1="14" x2="3" y2="21"/><line x1="21" y1="3" x2="14" y2="10"/>
                    </svg>
                  ) : (
                    /* expand / fullscreen icon */
                    <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
                      <polyline points="15 3 21 3 21 9"/><polyline points="9 21 3 21 3 15"/>
                      <line x1="21" y1="3" x2="14" y2="10"/><line x1="3" y1="21" x2="10" y2="14"/>
                    </svg>
                  )}
                </button>
              </div>
            </div>

            {/* ── Add Column dialog ── */}
            {showAddColDialog && (
              <div className="addColDialog">
                <div className="addColDialogInner">
                  <p className="addColDialogTitle">New column name</p>
                  <input
                    className="addColInput"
                    type="text"
                    placeholder="e.g. Change Request, Transport Number…"
                    value={newColLabel}
                    onChange={(e) => setNewColLabel(e.target.value)}
                    onKeyDown={(e) => { if (e.key === 'Enter') handleAddColumn(); if (e.key === 'Escape') setShowAddColDialog(false); }}
                    autoFocus
                  />
                  <div className="addColDialogButtons">
                    <button type="button" className="actionButton primaryButton" onClick={handleAddColumn} disabled={!newColLabel.trim()}>
                      Add
                    </button>
                    <button type="button" className="actionButton ghostButton addColCancelBtn" onClick={() => setShowAddColDialog(false)}>
                      Cancel
                    </button>
                  </div>
                </div>
              </div>
            )}

            {/* ── custom column chips ── */}
            {customColumns.length > 0 && (
              <div className="customColChips">
                {customColumns.map((col) => (
                  <span key={col.id} className="customColChip">
                    {col.label}
                    <button
                      type="button"
                      className="customColChipRemove"
                      onClick={() => handleRemoveColumn(col.id)}
                      title={`Remove column "${col.label}"`}
                    >
                      ×
                    </button>
                  </span>
                ))}
              </div>
            )}

            {headers.length > 0 ? (
              <div className={`tableWrapper${previewFullscreen ? ' dataPreviewFullscreenArea' : ' dataPreviewScrollArea'}`}>
                <table>
                  <thead>
                    <tr>
                      {headers.map((header, index) => (
                        <th key={`${header}-${index}`}>{header}</th>
                      ))}
                      <th>SLA Duration</th>
                      <th>SLA Status</th>
                      <th>Recent Update</th>
                      <th>Comments</th>
                      {customColumns.map((col) => (
                        <th key={col.id}>{col.label}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredIncidents.map((incident) => (
                      <tr key={incident.id}>
                        {headers.map((_, columnIndex) => (
                          <td key={`${incident.id}-${columnIndex}`}>{getCellValue(incident.originalRow[columnIndex])}</td>
                        ))}
                        <td>{formatDuration(incident)}</td>
                        <td>
                          <span className={`statusBadge ${incident.withinTarget ? 'statusGood' : 'statusBad'}`}>
                            {incident.withinTarget ? 'Within target' : 'Breached'}
                          </span>
                        </td>
                        <td className="commentCell">{incident.comment || '-'}</td>
                        <td>
                          <div className="commentBox">
                            <button
                              type="button"
                              className="actionButton smallButton"
                              onClick={() => toggleCommentEditor(incident.number)}
                            >
                              {editingComments[incident.number] ? 'Hide' : 'Add comments'}
                            </button>
                            {editingComments[incident.number] && (
                              <>
                                <textarea
                                  className="commentInput"
                                  value={draftComments[incident.number] || ''}
                                  onChange={(event) => handleCommentDraftChange(incident.number, event.target.value)}
                                  onKeyDown={(event) => {
                                    if (event.key === 'Enter' && (event.ctrlKey || event.metaKey)) {
                                      event.preventDefault();
                                      handleSaveComment(incident.number);
                                    }
                                  }}
                                  placeholder="Add new comment&#10;Ctrl+Enter to save"
                                  autoFocus
                                />
                                <button
                                  type="button"
                                  className="actionButton saveCommentButton"
                                  onClick={() => handleSaveComment(incident.number)}
                                >
                                  Save comment
                                </button>
                              </>
                            )}
                            {savedToast === incident.number && (
                              <span className="commentSavedToast">✓ Comment saved</span>
                            )}
                          </div>
                        </td>
                        {customColumns.map((col) => {
                          const key = cellKey(incident.number, col.id);
                          const isEditing = !!editingCustomCells[key];
                          const value = customColValues[incident.number]?.[col.id] || '';
                          return (
                            <td key={col.id}>
                              <div className="commentBox">
                                {isEditing ? (
                                  <>
                                    <input
                                      className="customCellInput"
                                      type="text"
                                      value={value}
                                      onChange={(e) => handleCustomCellChange(incident.number, col.id, e.target.value)}
                                      onKeyDown={(e) => {
                                        if (e.key === 'Enter') { e.preventDefault(); handleSaveCustomCell(incident.number, col.id); }
                                        if (e.key === 'Escape') toggleCustomCell(incident.number, col.id);
                                      }}
                                      placeholder={`Enter ${col.label}`}
                                      autoFocus
                                    />
                                    <button
                                      type="button"
                                      className="actionButton saveCommentButton"
                                      onClick={() => handleSaveCustomCell(incident.number, col.id)}
                                    >
                                      Save
                                    </button>
                                  </>
                                ) : (
                                  <span
                                    className="customCellValue"
                                    onClick={() => toggleCustomCell(incident.number, col.id)}
                                    title="Click to edit"
                                  >
                                    {value || <span className="customCellPlaceholder">—  click to add</span>}
                                  </span>
                                )}
                                {customCellToast === key && (
                                  <span className="commentSavedToast">✓ Saved</span>
                                )}
                              </div>
                            </td>
                          );
                        })}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="emptyState">Upload an Excel file to display ticket data here.</div>
            )}
          </section>
        </section>
      </section>
    </main>
  );
}

/* ═══════════════════════════════════════════════════════════════════════════
   ROOT APP — routing shell
═══════════════════════════════════════════════════════════════════════════ */

const pageStorageKey = 'watsonx-active-page';



function App() {
  const [page, setPage] = useState(() => {
    try {
      const saved = window.sessionStorage.getItem(pageStorageKey);
      if (saved === 'daily' || saved === 'breach') return saved;
    } catch {}
    return 'home';
  });

  const navigate = (target) => {
    try { window.sessionStorage.setItem(pageStorageKey, target); } catch {}
    setPage(target);
  };

  const goHome = () => {
    try { window.sessionStorage.setItem(pageStorageKey, 'home'); } catch {}
    setPage('home');
  };

  if (page === 'daily') {
    return <DailyTrackerPage onBack={goHome} />;
  }

  if (page === 'breach') {
    return <BreachAnalysisPage onBack={goHome} />;
  }

  return <HomePage onNavigate={navigate} />;
}

export default App;
