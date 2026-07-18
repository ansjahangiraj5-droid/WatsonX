import { useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';

const allowedExcelTypes = new Set([
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel'
]);
const allOption = 'ALL';
const dateRangeOptions = ['Today', 'Last 7 Days', 'Last 30 Days', 'ALL Time'];

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
  return value.toLocaleDateString('en-US', {
    month: 'short',
    day: 'numeric'
  });
}

function findHeaderIndex(headers, aliases) {
  return headers.findIndex((header) => aliases.includes(normalizeKey(header)));
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
  if (!startDate || !endDate || endDate < startDate) {
    return 0;
  }

  let current = new Date(startDate);
  let totalMilliseconds = 0;

  while (current < endDate) {
    if (!isWeekend(current)) {
      const endOfDay = new Date(current);
      endOfDay.setHours(23, 59, 59, 999);
      const segmentEnd = endOfDay < endDate ? endOfDay : endDate;
      totalMilliseconds += segmentEnd.getTime() - current.getTime();
    }

    const nextDay = startOfDay(current);
    nextDay.setDate(nextDay.getDate() + 1);
    current = nextDay;
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

function App() {
  const fileInputRef = useRef(null);
  const [fileName, setFileName] = useState('');
  const [headers, setHeaders] = useState([]);
  const [rows, setRows] = useState([]);
  const [sheetName, setSheetName] = useState('');
  const [error, setError] = useState('');
  const [selectedAssignmentGroup, setSelectedAssignmentGroup] = useState(allOption);
  const [selectedAssignee, setSelectedAssignee] = useState(allOption);
  const [selectedBreachStatus, setSelectedBreachStatus] = useState(allOption);
  const [selectedDateRange, setSelectedDateRange] = useState('ALL Time');

  const resetData = (message = '', nextFileName = '') => {
    setFileName(nextFileName);
    setHeaders([]);
    setRows([]);
    setSheetName('');
    setSelectedAssignmentGroup(allOption);
    setSelectedAssignee(allOption);
    setSelectedBreachStatus(allOption);
    setSelectedDateRange('ALL Time');
    setError(message);
  };

  const handleClear = () => {
    resetData();

    if (fileInputRef.current) {
      fileInputRef.current.value = '';
    }
  };

  const incidentConfig = useMemo(() => {
    if (headers.length === 0) {
      return null;
    }

    const numberIndex = findHeaderIndex(headers, ['number']);
    const assignmentGroupIndex = findHeaderIndex(headers, ['assignment group']);
    const assignedToIndex = findHeaderIndex(headers, ['assigned to']);
    const priorityIndex = findHeaderIndex(headers, ['priority']);
    const createdIndex = findHeaderIndex(headers, ['created']);
    const breachedIndex = findHeaderIndex(headers, ['has breached']);
    const categoryIndex = findHeaderIndex(headers, ['category']);
    const slaIndex = findHeaderIndex(headers, ['sla definition']);
    const stageIndex = findHeaderIndex(headers, ['stage']);
    const resolvedIndex = findHeaderIndex(headers, ['resolved']);

    if ([
      numberIndex,
      assignmentGroupIndex,
      assignedToIndex,
      priorityIndex,
      createdIndex,
      breachedIndex,
      categoryIndex,
      slaIndex,
      stageIndex,
      resolvedIndex
    ].some((index) => index === -1)) {
      return null;
    }

    return {
      numberIndex,
      assignmentGroupIndex,
      assignedToIndex,
      priorityIndex,
      createdIndex,
      breachedIndex,
      categoryIndex,
      slaIndex,
      stageIndex,
      resolvedIndex
    };
  }, [headers]);

  const incidents = useMemo(() => {
    if (!incidentConfig) {
      return [];
    }

    return rows.map((row, index) => {
      const createdRaw = row[incidentConfig.createdIndex];
      const resolvedRaw = row[incidentConfig.resolvedIndex];
      const createdDate = createdRaw instanceof Date ? createdRaw : new Date(createdRaw);
      const resolvedDate = resolvedRaw instanceof Date ? resolvedRaw : new Date(resolvedRaw);
      const safeCreatedDate = Number.isNaN(createdDate.getTime()) ? null : createdDate;
      const safeResolvedDate = Number.isNaN(resolvedDate.getTime()) ? null : resolvedDate;
      const target = getPriorityTarget(row[incidentConfig.priorityIndex]);
      const resolutionDays = calculateBusinessDays(safeCreatedDate, safeResolvedDate);
      const resolutionHours = calculateBusinessHours(safeCreatedDate, safeResolvedDate);
      const actualDuration = target.unit === 'hours' ? resolutionHours : resolutionDays;
      const withinTarget = target.limit > 0 && actualDuration <= target.limit;

      return {
        id: `${getCellValue(row[incidentConfig.numberIndex])}-${index}`,
        number: getCellValue(row[incidentConfig.numberIndex]),
        assignmentGroup: getCellValue(row[incidentConfig.assignmentGroupIndex]),
        assignedTo: getCellValue(row[incidentConfig.assignedToIndex]),
        priority: getCellValue(row[incidentConfig.priorityIndex]),
        created: getCellValue(row[incidentConfig.createdIndex]),
        createdDate: safeCreatedDate,
        hasBreached: getCellValue(row[incidentConfig.breachedIndex]),
        category: getCellValue(row[incidentConfig.categoryIndex]),
        slaDefinition: getCellValue(row[incidentConfig.slaIndex]),
        stage: getCellValue(row[incidentConfig.stageIndex]),
        resolved: getCellValue(row[incidentConfig.resolvedIndex]),
        resolvedDate: safeResolvedDate,
        resolutionDays,
        resolutionHours,
        actualDuration,
        target,
        withinTarget,
        originalRow: row
      };
    });
  }, [incidentConfig, rows]);

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
    const breachedCount = filteredIncidents.filter((incident) => /true|yes|breached/i.test(incident.hasBreached)).length;
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

      const nextHeaders = nonEmptyRows[0].map((header, index) => String(header).trim() || `Column ${index + 1}`);
      const nextRows = nonEmptyRows.slice(1).map((row) => nextHeaders.map((_, index) => row[index] ?? ''));
      const requiredIndexes = [
        findHeaderIndex(nextHeaders, ['number']),
        findHeaderIndex(nextHeaders, ['assignment group']),
        findHeaderIndex(nextHeaders, ['assigned to']),
        findHeaderIndex(nextHeaders, ['priority']),
        findHeaderIndex(nextHeaders, ['created']),
        findHeaderIndex(nextHeaders, ['has breached']),
        findHeaderIndex(nextHeaders, ['category']),
        findHeaderIndex(nextHeaders, ['sla definition']),
        findHeaderIndex(nextHeaders, ['stage']),
        findHeaderIndex(nextHeaders, ['resolved'])
      ];

      if (requiredIndexes.some((index) => index === -1)) {
        resetData(
          'The Excel file must contain Number, Assignment Group, Assigned To, Priority, Created, Has Breached, Category, SLA Definition, Stage, and Resolved columns.',
          selectedFile.name
        );
        return;
      }

      setFileName(selectedFile.name);
      setSheetName(firstSheetName);
      setError('');
      setSelectedAssignmentGroup(allOption);
      setSelectedAssignee(allOption);
      setSelectedBreachStatus(allOption);
      setSelectedDateRange('ALL Time');
      setHeaders(nextHeaders);
      setRows(nextRows);
    } catch (uploadError) {
      resetData('Unable to read the uploaded file. Please upload a valid Excel file.', selectedFile.name);
    } finally {
      event.target.value = '';
    }
  };

  return (
    <main className="page">
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
            <small>Choose Today, Last 7 Days, Last 30 Days, or All Time.</small>
          </div>

          <div className="sideCard compactList">
            <span className="sideLabel">SLA targets</span>
            <span>Priority 4: 6 business days</span>
            <span>Priority 3: 3 business days</span>
            <span>Priority 2: 8 hours</span>
            <span>Priority 1: 4 hours</span>
          </div>
        </aside>

        <section className="dashboardPanel">
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
              <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />
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

          <section className="tableCard overviewCard">
            <div className="tableHeader tableActions fixedHeader">
              <div>
                <h2>Ticket data preview</h2>
                <p className="tableSubtitle darkSubtitle">
                  Showing {formatNumber(filteredIncidents.length)} records after applying the selected filters.
                </p>
              </div>
              <span className="pill neutral">First worksheet</span>
            </div>

            {headers.length > 0 ? (
              <div className="tableWrapper dataPreviewScrollArea">
                <table>
                  <thead>
                    <tr>
                      {headers.map((header, index) => (
                        <th key={`${header}-${index}`}>{header}</th>
                      ))}
                      <th>SLA Duration</th>
                      <th>SLA Status</th>
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
                            {incident.withinTarget ? 'Within target' : 'Out of target'}
                          </span>
                        </td>
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

export default App;
