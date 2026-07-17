import { useMemo, useState } from 'react';

const defaultRows = [];

function App() {
  const [fileName, setFileName] = useState('');
  const [headers, setHeaders] = useState([]);
  const [rows, setRows] = useState(defaultRows);
  const [error, setError] = useState('');

  const metrics = useMemo(() => {
    return {
      totalRows: rows.length,
      totalColumns: headers.length,
      populatedCells: rows.reduce((count, row) => {
        return count + row.filter((cell) => String(cell ?? '').trim() !== '').length;
      }, 0)
    };
  }, [headers, rows]);

  const handleFileUpload = async (event) => {
    const selectedFile = event.target.files?.[0];

    if (!selectedFile) {
      return;
    }

    setFileName(selectedFile.name);
    setError('');

    try {
      const text = await selectedFile.text();
      const lines = text
        .split(/\r?\n/)
        .map((line) => line.trim())
        .filter(Boolean);

      if (lines.length === 0) {
        setHeaders([]);
        setRows([]);
        setError('The uploaded file is empty.');
        return;
      }

      const parsedRows = lines.map((line) => line.split(',').map((cell) => cell.trim()));
      setHeaders(parsedRows[0]);
      setRows(parsedRows.slice(1));
    } catch (uploadError) {
      setHeaders([]);
      setRows([]);
      setError('Unable to read the uploaded file. Please use a CSV file.');
    }
  };

  return (
    <main className="page">
      <section className="hero">
        <div>
          <p className="eyebrow">React dashboard</p>
          <h1>Upload a CSV file and preview the data</h1>
          <p className="subtitle">
            Choose a local CSV file to see summary metrics and a data table directly in the UI.
          </p>
        </div>
        <label className="uploadCard">
          <span>Upload file</span>
          <input type="file" accept=".csv" onChange={handleFileUpload} />
          <strong>{fileName || 'No file selected'}</strong>
        </label>
      </section>

      <section className="metrics">
        <article>
          <span>Rows</span>
          <strong>{metrics.totalRows}</strong>
        </article>
        <article>
          <span>Columns</span>
          <strong>{metrics.totalColumns}</strong>
        </article>
        <article>
          <span>Filled cells</span>
          <strong>{metrics.populatedCells}</strong>
        </article>
      </section>

      <section className="tableCard">
        <div className="tableHeader">
          <h2>Uploaded data</h2>
          {error ? <p className="error">{error}</p> : null}
        </div>

        {headers.length > 0 ? (
          <div className="tableWrapper">
            <table>
              <thead>
                <tr>
                  {headers.map((header) => (
                    <th key={header}>{header}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {rows.map((row, rowIndex) => (
                  <tr key={`${rowIndex}-${row.join('-')}`}>
                    {headers.map((_, columnIndex) => (
                      <td key={`${rowIndex}-${columnIndex}`}>{row[columnIndex] || '-'}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : (
          <div className="emptyState">Upload a CSV file to display data here.</div>
        )}
      </section>
    </main>
  );
}

export default App;
