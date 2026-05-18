import React from 'react'

// Visar en resultattabell ({columns, rows, row_count, truncated}).
export default function DataTable({ table, copyColumnLabel, onCopyColumn }) {
  if (!table || table.columns.length === 0) {
    return <div className="empty-state">Ingen data.</div>
  }
  if (table.rows.length === 0) {
    return <div className="empty-state">Tabellen är tom ({table.row_count} rader).</div>
  }
  return (
    <div className="table-wrap">
      <table className="data-table">
        <thead>
          <tr>
            <th className="row-num">#</th>
            {table.columns.map((c, index) => (
              <th key={`${c}-${index}`}>
                {onCopyColumn ? (
                  <button
                    type="button"
                    className="copy-header-btn"
                    title={`Kopiera ${c}`}
                    onClick={() => onCopyColumn(index)}
                  >
                    {copyColumnLabel || c}
                  </button>
                ) : (
                  c
                )}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {table.rows.map((row, i) => (
            <tr key={i}>
              <td className="row-num">{i + 1}</td>
              {row.map((cell, j) => (
                <td key={j} title={cell}>
                  {cell}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
      {table.truncated && (
        <div className="table-note">
          Visar de första {table.rows.length} av {table.row_count} raderna. Öppna i Excel för hela
          resultatet.
        </div>
      )}
    </div>
  )
}
