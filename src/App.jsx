import { useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';

const allowedExcelTypes = new Set([
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel'
]);
const allOption = 'All';

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

function findHeaderIndex(headers, aliases) {
  return headers.findIndex((header) => aliases.includes(normalizeKey(header)));
}

function buildBreakdown(items, keyName, valueName) {
  return Object.entries(items)
    .map(([key, total]) => ({ [keyName]: key, [valueName]: total }))
    .sort((left, right) => right[valueName] - left[valueName]);
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

  const resetData = (message = '', nextFileName = '') => {
    setFileName(nextFileName);
    setHeaders([]);
    setRows([]);
    setSheetName('');
    setSelectedAssignmentGroup(allOption);
    setSelectedAssignee(allOption);
    setSelectedBreachStatus(allOption);
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
      const createdDate = createdRaw instanceof Date ? createdRaw : new Date(createdRaw);

      return {
        id: `${getCellValue(row[incidentConfig.numberIndex])}-${index}`,
        number: getCellValue(row[incidentConfig.numberIndex]),
        assignmentGroup: getCellValue(row[incidentConfig.assignmentGroupIndex]),
        assignedTo: getCellValue(row[incidentConfig.assignedToIndex]),
        priority: getCellValue(row[incidentConfig.priorityIndex]),
        created: getCellValue(row[incidentConfig.createdIndex]),
        createdDate: Number.isNaN(createdDate.getTime()) ? null : createdDate,
        hasBreached: getCellValue(row[incidentConfig.breachedIndex]),
        category: getCellValue(row[incidentConfig.categoryIndex]),
        slaDefinition: getCellValue(row[incidentConfig.slaIndex]),
        stage: getCellValue(row[incidentConfig.stageIndex]),
        resolved: getCellValue(row[incidentConfig.resolvedIndex]),
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
    return incidents.filter((incident) => {
      const matchesAssignmentGroup = selectedAssignmentGroup === allOption || incident.assignmentGroup === selectedAssignmentGroup;
      const matchesAssignee = selectedAssignee === allOption || incident.assignedTo === selectedAssignee;
      const matchesBreachStatus = selectedBreachStatus === allOption || incident.hasBreached === selectedBreachStatus;

      return matchesAssignmentGroup && matchesAssignee && matchesBreachStatus;
    });
  }, [incidents, selectedAssignmentGroup, selectedAssignee, selectedBreachStatus]);

  const metrics = useMemo(() => {
    const breachedCount = filteredIncidents.filter((incident) => /true|yes|breached/i.test(incident.hasBreached)).length;
    const unassignedCount = filteredIncidents.filter((incident) => incident.assignedTo === '-').length;
    const uniqueGroups = new Set(filteredIncidents.map((incident) => incident.assignmentGroup).filter((item) => item !== '-')).size;

    return {
      totalIncidents: filteredIncidents.length,
      breachedCount,
      uniqueGroups,
      unassignedCount
    };
  }, [filteredIncidents]);

  const assignmentGroupBreakdown = useMemo(() => {
    const grouped = filteredIncidents.reduce((accumulator, incident) => {
      accumulator[incident.assignmentGroup] = (accumulator[incident.assignmentGroup] || 0) + 1;
      return accumulator;
    }, {});

    return buildBreakdown(grouped, 'assignmentGroup', 'total').slice(0, 8);
  }, [filteredIncidents]);

  const assigneeBreakdown = useMemo(() => {
    const grouped = filteredIncidents.reduce((accumulator, incident) => {
      accumulator[incident.assignedTo] = (accumulator[incident.assignedTo] || 0) + 1;
      return accumulator;
    }, {});

    return buildBreakdown(grouped, 'assignedTo', 'total').slice(0, 8);
  }, [filteredIncidents]);

  const breachBreakdown = useMemo(() => {
    const grouped = filteredIncidents.reduce((accumulator, incident) => {
      accumulator[incident.hasBreached] = (accumulator[incident.hasBreached] || 0) + 1;
      return accumulator;
    }, {});

    return buildBreakdown(grouped, 'status', 'total');
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

    return buildBreakdown(grouped, 'date', 'total').slice(-7);
  }, [filteredIncidents]);

  const maxAssignmentGroupValue = Math.max(...assignmentGroupBreakdown.map((item) => item.total), 1);
  const maxAssigneeValue = Math.max(...assigneeBreakdown.map((item) => item.total), 1);
  const maxBreachValue = Math.max(...breachBreakdown.map((item) => item.total), 1);
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

          <div className="sideCard">
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
            <small>Users can combine these filters to inspect the dashboard.</small>
          </div>

          <div className="sideCard compactList">
            <span className="sideLabel">Expected columns</span>
            <span>Number</span>
            <span>Assignment group</span>
            <span>Assigned to</span>
            <span>Has breached</span>
          </div>
        </aside>

        <section className="dashboardPanel">
          <section className="hero heroGradient">
            <div>
              <p className="eyebrow highlight">Operations dashboard</p>
              <h1>Excel-driven ticket analytics</h1>
              <p className="subtitle lightText">
                Upload a workbook containing service tickets and explore visual charts for assignment groups,
                assignees, and breached records with live filtering.
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
              <span>Assignment groups</span>
              <strong>{formatNumber(metrics.uniqueGroups)}</strong>
            </article>
            <article className="metricCard tealCard">
              <span>Unassigned</span>
              <strong>{formatNumber(metrics.unassignedCount)}</strong>
            </article>
          </section>

          <section className="insightsGrid wideGrid">
            <article className="tableCard overviewCard darkPanel">
              <div className="tableHeader">
                <h2>Assignment group workload</h2>
                <span className="pill">{selectedAssignmentGroup}</span>
              </div>
              <div className="chartPanel">
                {assignmentGroupBreakdown.length > 0 ? (
                  assignmentGroupBreakdown.map((item) => (
                    <div className="chartRow" key={item.assignmentGroup}>
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
              <div className="tableHeader">
                <h2>Has breached distribution</h2>
                <span className="pill neutral">Filtered view</span>
              </div>
              <div className="chartPanel verticalPanel compactVerticalPanel">
                {breachBreakdown.length > 0 ? (
                  breachBreakdown.map((item) => (
                    <div className="verticalMetric" key={item.status}>
                      <div className="verticalChart">
                        <div
                          className="verticalBar"
                          style={{ height: `${(item.total / maxBreachValue) * 100}%` }}
                        />
                      </div>
                      <strong>{item.total}</strong>
                      <span>{item.status}</span>
                    </div>
                  ))
                ) : (
                  <div className="emptyState">Breach summary appears here after upload.</div>
                )}
              </div>
            </article>
          </section>

          <section className="insightsGrid">
            <article className="tableCard overviewCard">
              <div className="tableHeader">
                <h2>Assigned to distribution</h2>
                <span className="pill neutral">Filtered view</span>
              </div>
              <div className="chartPanel">
                {assigneeBreakdown.length > 0 ? (
                  assigneeBreakdown.map((item) => (
                    <div className="chartRow" key={item.assignedTo}>
                      <div className="chartLabel">{item.assignedTo}</div>
                      <div className="chartTrack">
                        <div
                          className="chartBar gradientBlue"
                          style={{ width: `${(item.total / maxAssigneeValue) * 100}%` }}
                        />
                      </div>
                      <div className="chartValue">{item.total}</div>
                    </div>
                  ))
                ) : (
                  <div className="emptyState">Assignee charts appear here after upload.</div>
                )}
              </div>
            </article>

            <article className="tableCard overviewCard summaryGradient">
              <div className="tableHeader">
                <h2>Recent ticket creation</h2>
                <span className="pill neutral">Last 7 dates</span>
              </div>
              <div className="chartPanel verticalPanel compactVerticalPanel">
                {recentIncidentTrend.length > 0 ? (
                  recentIncidentTrend.map((item) => (
                    <div className="verticalMetric" key={item.date}>
                      <div className="verticalChart trendChart">
                        <div
                          className="verticalBar trendBar"
                          style={{ height: `${(item.total / maxTrendValue) * 100}%` }}
                        />
                      </div>
                      <strong>{item.total}</strong>
                      <span>{item.date}</span>
                    </div>
                  ))
                ) : (
                  <div className="emptyState">Created-date activity appears here when the data is available.</div>
                )}
              </div>
            </article>
          </section>

          <section className="tableCard">
            <div className="tableHeader tableActions">
              <div>
                <h2>Ticket data preview</h2>
                <p className="tableSubtitle">
                  Showing {formatNumber(filteredIncidents.length)} records after applying the selected filters.
                </p>
              </div>
              <span className="pill neutral">First worksheet</span>
            </div>

            {headers.length > 0 ? (
              <div className="tableWrapper">
                <table>
                  <thead>
                    <tr>
                      {headers.map((header, index) => (
                        <th key={`${header}-${index}`}>{header}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {filteredIncidents.map((incident) => (
                      <tr key={incident.id}>
                        {headers.map((_, columnIndex) => (
                          <td key={`${incident.id}-${columnIndex}`}>{getCellValue(incident.originalRow[columnIndex])}</td>
                        ))}
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
