import { useMemo, useRef, useState } from 'react';
import * as XLSX from 'xlsx';

const allowedExcelTypes = new Set([
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel'
]);

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

function App() {
  const fileInputRef = useRef(null);
  const [fileName, setFileName] = useState('');
  const [headers, setHeaders] = useState([]);
  const [rows, setRows] = useState([]);
  const [sheetName, setSheetName] = useState('');
  const [error, setError] = useState('');
  const [selectedUser, setSelectedUser] = useState('All Users');

  const resetData = (message = '', nextFileName = '') => {
    setFileName(nextFileName);
    setHeaders([]);
    setRows([]);
    setSheetName('');
    setSelectedUser('All Users');
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

    const incidentIndex = findHeaderIndex(headers, ['incident number', 'incident', 'ticket number', 'ticket', 'incident id']);
    const daysIndex = findHeaderIndex(headers, ['no of days from ticket raised', 'days from ticket raised', 'days open', 'ageing days', 'aging days', 'days']);
    const userIndex = findHeaderIndex(headers, ['users name assigned to', 'assigned to', 'assignee', 'user name', 'assigned user']);
    const priorityIndex = findHeaderIndex(headers, ['priority level', 'priority', 'severity']);

    if ([incidentIndex, daysIndex, userIndex, priorityIndex].some((index) => index === -1)) {
      return null;
    }

    return {
      incidentIndex,
      daysIndex,
      userIndex,
      priorityIndex
    };
  }, [headers]);

  const incidents = useMemo(() => {
    if (!incidentConfig) {
      return [];
    }

    return rows.map((row, index) => {
      const incidentNumber = getCellValue(row[incidentConfig.incidentIndex]);
      const assignedTo = getCellValue(row[incidentConfig.userIndex]);
      const priority = getCellValue(row[incidentConfig.priorityIndex]);
      const daysRaw = row[incidentConfig.daysIndex];
      const daysOpen = Number(daysRaw);

      return {
        id: `${incidentNumber}-${index}`,
        incidentNumber,
        assignedTo,
        priority,
        daysOpen: Number.isFinite(daysOpen) ? daysOpen : 0,
        originalRow: row
      };
    });
  }, [incidentConfig, rows]);

  const users = useMemo(() => {
    const names = Array.from(new Set(incidents.map((item) => item.assignedTo).filter((name) => name !== '-')));
    return ['All Users', ...names];
  }, [incidents]);

  const filteredIncidents = useMemo(() => {
    if (selectedUser === 'All Users') {
      return incidents;
    }

    return incidents.filter((incident) => incident.assignedTo === selectedUser);
  }, [incidents, selectedUser]);

  const metrics = useMemo(() => {
    const overdue = filteredIncidents.filter((incident) => incident.daysOpen > 30).length;
    const critical = filteredIncidents.filter((incident) => /p1|critical|high/i.test(incident.priority)).length;
    const averageAge = filteredIncidents.length > 0
      ? filteredIncidents.reduce((sum, incident) => sum + incident.daysOpen, 0) / filteredIncidents.length
      : 0;

    return {
      totalIncidents: filteredIncidents.length,
      overdue,
      critical,
      averageAge: averageAge.toFixed(1)
    };
  }, [filteredIncidents]);

  const priorityBreakdown = useMemo(() => {
    const grouped = filteredIncidents.reduce((accumulator, incident) => {
      accumulator[incident.priority] = (accumulator[incident.priority] || 0) + 1;
      return accumulator;
    }, {});

    return Object.entries(grouped).map(([priority, total]) => ({ priority, total }));
  }, [filteredIncidents]);

  const ageBuckets = useMemo(() => {
    const buckets = [
      { label: '0-7 days', min: 0, max: 7, total: 0 },
      { label: '8-15 days', min: 8, max: 15, total: 0 },
      { label: '16-30 days', min: 16, max: 30, total: 0 },
      { label: '31-60 days', min: 31, max: 60, total: 0 },
      { label: '60+ days', min: 61, max: Number.POSITIVE_INFINITY, total: 0 }
    ];

    filteredIncidents.forEach((incident) => {
      const bucket = buckets.find((item) => incident.daysOpen >= item.min && incident.daysOpen <= item.max);
      if (bucket) {
        bucket.total += 1;
      }
    });

    return buckets;
  }, [filteredIncidents]);

  const topUsers = useMemo(() => {
    const grouped = incidents.reduce((accumulator, incident) => {
      accumulator[incident.assignedTo] = (accumulator[incident.assignedTo] || 0) + 1;
      return accumulator;
    }, {});

    return Object.entries(grouped)
      .filter(([name]) => name !== '-')
      .map(([name, total]) => ({ name, total }))
      .sort((left, right) => right.total - left.total)
      .slice(0, 5);
  }, [incidents]);

  const maxPriorityValue = Math.max(...priorityBreakdown.map((item) => item.total), 1);
  const maxAgeValue = Math.max(...ageBuckets.map((item) => item.total), 1);
  const maxUserValue = Math.max(...topUsers.map((item) => item.total), 1);

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
      const workbook = XLSX.read(fileBuffer, { type: 'array' });
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
        findHeaderIndex(nextHeaders, ['incident number', 'incident', 'ticket number', 'ticket', 'incident id']),
        findHeaderIndex(nextHeaders, ['no of days from ticket raised', 'days from ticket raised', 'days open', 'ageing days', 'aging days', 'days']),
        findHeaderIndex(nextHeaders, ['users name assigned to', 'assigned to', 'assignee', 'user name', 'assigned user']),
        findHeaderIndex(nextHeaders, ['priority level', 'priority', 'severity'])
      ];

      if (requiredIndexes.some((index) => index === -1)) {
        resetData(
          'The Excel file must contain Incident Number, No of Days from Ticket Raised, Users Name Assigned To, and Priority Level columns.',
          selectedFile.name
        );
        return;
      }

      setFileName(selectedFile.name);
      setSheetName(firstSheetName);
      setError('');
      setSelectedUser('All Users');
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
              <h2>Incident Command Center</h2>
            </div>
          </div>

          <div className="sideCard">
            <span className="sideLabel">Workbook</span>
            <strong>{fileName || 'No Excel file uploaded'}</strong>
            <small>{sheetName ? `Sheet: ${sheetName}` : 'Upload an Excel file to begin'}</small>
          </div>

          <div className="sideCard">
            <label htmlFor="user-filter" className="sideLabel">Assigned user view</label>
            <select
              id="user-filter"
              className="filterSelect"
              value={selectedUser}
              onChange={(event) => setSelectedUser(event.target.value)}
              disabled={users.length === 1}
            >
              {users.map((user) => (
                <option key={user} value={user}>{user}</option>
              ))}
            </select>
            <small>See incidents for a specific assignee in charts and cards.</small>
          </div>

          <div className="sideCard compactList">
            <span className="sideLabel">Expected columns</span>
            <span>Incident Number</span>
            <span>No of Days from Ticket Raised</span>
            <span>Users Name Assigned To</span>
            <span>Priority Level</span>
          </div>
        </aside>

        <section className="dashboardPanel">
          <section className="hero heroGradient">
            <div>
              <p className="eyebrow highlight">Operations dashboard</p>
              <h1>Excel-driven incident dashboard</h1>
              <p className="subtitle lightText">
                Upload an incident workbook and monitor ageing, assignee workload, and priority mix in a
                colorful dashboard inspired by enterprise support consoles.
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
              <span>Total incidents</span>
              <strong>{formatNumber(metrics.totalIncidents)}</strong>
            </article>
            <article className="metricCard blueCard">
              <span>Over 30 days</span>
              <strong>{formatNumber(metrics.overdue)}</strong>
            </article>
            <article className="metricCard purpleCard">
              <span>High / critical</span>
              <strong>{formatNumber(metrics.critical)}</strong>
            </article>
            <article className="metricCard tealCard">
              <span>Average age</span>
              <strong>{metrics.averageAge}</strong>
            </article>
          </section>

          <section className="insightsGrid wideGrid">
            <article className="tableCard overviewCard darkPanel">
              <div className="tableHeader">
                <h2>Incident priority mix</h2>
                <span className="pill">{selectedUser}</span>
              </div>
              <div className="chartPanel">
                {priorityBreakdown.length > 0 ? (
                  priorityBreakdown.map((item) => (
                    <div className="chartRow" key={item.priority}>
                      <div className="chartLabel lightText">{item.priority}</div>
                      <div className="chartTrack darkTrack">
                        <div
                          className="chartBar gradientPink"
                          style={{ width: `${(item.total / maxPriorityValue) * 100}%` }}
                        />
                      </div>
                      <div className="chartValue lightText">{item.total}</div>
                    </div>
                  ))
                ) : (
                  <div className="emptyState lightText">Upload valid incident data to see the priority graph.</div>
                )}
              </div>
            </article>

            <article className="tableCard overviewCard">
              <div className="tableHeader">
                <h2>Incidents by ageing bucket</h2>
                <span className="pill neutral">Current selection</span>
              </div>
              <div className="chartPanel verticalPanel">
                {ageBuckets.map((bucket) => (
                  <div className="verticalMetric" key={bucket.label}>
                    <div className="verticalChart">
                      <div
                        className="verticalBar"
                        style={{ height: `${(bucket.total / maxAgeValue) * 100 || 0}%` }}
                      />
                    </div>
                    <strong>{bucket.total}</strong>
                    <span>{bucket.label}</span>
                  </div>
                ))}
              </div>
            </article>
          </section>

          <section className="insightsGrid">
            <article className="tableCard overviewCard">
              <div className="tableHeader">
                <h2>Assigned user workload</h2>
                <span className="pill neutral">All users</span>
              </div>
              <div className="chartPanel">
                {topUsers.length > 0 ? (
                  topUsers.map((item) => (
                    <div className="chartRow" key={item.name}>
                      <div className="chartLabel">{item.name}</div>
                      <div className="chartTrack">
                        <div
                          className="chartBar gradientBlue"
                          style={{ width: `${(item.total / maxUserValue) * 100}%` }}
                        />
                      </div>
                      <div className="chartValue">{item.total}</div>
                    </div>
                  ))
                ) : (
                  <div className="emptyState">User workload appears here after upload.</div>
                )}
              </div>
            </article>

            <article className="tableCard overviewCard summaryGradient">
              <div className="tableHeader">
                <h2>User incident spotlight</h2>
                <span className="pill neutral">Filtered</span>
              </div>
              <div className="summaryList spotlightList">
                <div>
                  <span>Selected user</span>
                  <strong>{selectedUser}</strong>
                </div>
                <div>
                  <span>Visible incidents</span>
                  <strong>{formatNumber(filteredIncidents.length)}</strong>
                </div>
                <div>
                  <span>Oldest incident age</span>
                  <strong>{filteredIncidents.length > 0 ? `${Math.max(...filteredIncidents.map((item) => item.daysOpen))} days` : '-'}</strong>
                </div>
                <div>
                  <span>Highest priority in view</span>
                  <strong>{filteredIncidents[0]?.priority || '-'}</strong>
                </div>
              </div>
            </article>
          </section>

          <section className="tableCard">
            <div className="tableHeader tableActions">
              <div>
                <h2>Incident data preview</h2>
                <p className="tableSubtitle">
                  {selectedUser === 'All Users'
                    ? 'Showing all incidents from the uploaded worksheet.'
                    : `Showing incidents assigned to ${selectedUser}.`}
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
              <div className="emptyState">Upload an Excel file to display incident data here.</div>
            )}
          </section>
        </section>
      </section>
    </main>
  );
}

export default App;
