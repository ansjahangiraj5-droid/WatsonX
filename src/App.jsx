import { useMemo, useState } from 'react';
import * as XLSX from 'xlsx';

const allowedExcelTypes = new Set([
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel'
]);

function getCellValue(value) {
  if (value === null || value === undefined || value === '') {
    return '-';
  }

  return String(value);
}

function App() {
  const [fileName, setFileName] = useState('');
  const [headers, setHeaders] = useState([]);
  const [rows, setRows] = useState([]);
  const [sheetName, setSheetName] = useState('');
  const [error, setError] = useState('');

  const metrics = useMemo(() => {
    const populatedCells = rows.reduce((count, row) => {
      return count + row.filter((cell) => String(cell ?? '').trim() !== '').length;
    }, 0);

    const numericCells = rows.flatMap((row) => row).filter((cell) => typeof cell === 'number');
    const totalValue = numericCells.reduce((sum, value) => sum + value, 0);
    const averageValue = numericCells.length > 0 ? totalValue / numericCells.length : 0;

    return {
      totalRows: rows.length,
      totalColumns: headers.length,
      populatedCells,
      numericCells: numericCells.length,
      averageValue: averageValue.toFixed(2)
    };
  }, [headers, rows]);

  const chartData = useMemo(() => {
    if (headers.length === 0 || rows.length === 0) {
      return [];
    }

    return headers.slice(0, 6).map((header, columnIndex) => {
      const filledCount = rows.reduce((count, row) => {
        return String(row[columnIndex] ?? '').trim() !== '' ? count + 1 : count;
      }, 0);

      return {
        header,
        filledCount
      };
    });
  }, [headers, rows]);

  const maxChartValue = useMemo(() => {
    if (chartData.length === 0) {
      return 1;
    }

    return Math.max(...chartData.map((item) => item.filledCount), 1);
  }, [chartData]);

  const resetData = (message, nextFileName = '') => {
    setFileName(nextFileName);
    setHeaders([]);
    setRows([]);
    setSheetName('');
    setError(message);
  };

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

      setFileName(selectedFile.name);
      setSheetName(firstSheetName);
      setError('');
      setHeaders(nonEmptyRows[0].map((header, index) => String(header).trim() || `Column ${index + 1}`));
      setRows(nonEmptyRows.slice(1).map((row) => nonEmptyRows[0].map((_, index) => row[index] ?? '')));
    } catch (uploadError) {
      resetData('Unable to read the uploaded file. Please upload a valid Excel file.', selectedFile.name);
    } finally {
      event.target.value = '';
    }
  };

  return (
    <main className="page">
      <section className="hero">
        <div>
          <p className="eyebrow">Operations dashboard</p>
          <h1>Upload an Excel file to generate a live service dashboard</h1>
          <p className="subtitle">
            This dashboard accepts only Excel files and turns the first worksheet into summary cards,
            a visual activity board, and a data table.
          </p>
        </div>
        <label className="uploadCard">
          <span>Upload Excel file</span>
          <input type="file" accept=".xlsx,.xls" onChange={handleFileUpload} />
          <strong>{fileName || 'No Excel file selected'}</strong>
          <small>{sheetName ? `Sheet: ${sheetName}` : 'Accepted formats: .xlsx, .xls'}</small>
        </label>
      </section>

      {error ? <p className="bannerError">{error}</p> : null}

      <section className="metrics">
        <article>
          <span>Total records</span>
          <strong>{metrics.totalRows}</strong>
        </article>
        <article>
          <span>Total fields</span>
          <strong>{metrics.totalColumns}</strong>
        </article>
        <article>
          <span>Populated cells</span>
          <strong>{metrics.populatedCells}</strong>
        </article>
        <article>
          <span>Numeric cells</span>
          <strong>{metrics.numericCells}</strong>
        </article>
        <article>
          <span>Average numeric value</span>
          <strong>{metrics.averageValue}</strong>
        </article>
      </section>

      <section className="insightsGrid">
        <article className="tableCard overviewCard">
          <div className="tableHeader">
            <h2>Dataset health</h2>
            <span className="pill">Live</span>
          </div>
          <div className="chartPanel">
            {chartData.length > 0 ? (
              chartData.map((item) => (
                <div className="chartRow" key={item.header}>
                  <div className="chartLabel">{item.header}</div>
                  <div className="chartTrack">
                    <div
                      className="chartBar"
                      style={{ width: `${(item.filledCount / maxChartValue) * 100}%` }}
                    />
                  </div>
                  <div className="chartValue">{item.filledCount}</div>
                </div>
              ))
            ) : (
              <div className="emptyState">Upload an Excel file to see graphical insights.</div>
            )}
          </div>
        </article>

        <article className="tableCard overviewCard">
          <div className="tableHeader">
            <h2>Platform summary</h2>
            <span className="pill neutral">Workbook</span>
          </div>
          <div className="summaryList">
            <div><span>File name</span><strong>{fileName || '-'}</strong></div>
            <div><span>Worksheet</span><strong>{sheetName || '-'}</strong></div>
            <div><span>Visible columns</span><strong>{headers.slice(0, 6).join(', ') || '-'}</strong></div>
            <div><span>Dashboard status</span><strong>{headers.length > 0 ? 'Ready' : 'Waiting for upload'}</strong></div>
          </div>
        </article>
      </section>

      <section className="tableCard">
        <div className="tableHeader">
          <h2>Excel data preview</h2>
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
                {rows.map((row, rowIndex) => (
                  <tr key={`${rowIndex}-${row.join('-')}`}>
                    {headers.map((_, columnIndex) => (
                      <td key={`${rowIndex}-${columnIndex}`}>{getCellValue(row[columnIndex])}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : (
          <div className="emptyState">Upload an Excel file to display the worksheet data here.</div>
        )}
      </section>
    </main>
  );
}

export default App;
