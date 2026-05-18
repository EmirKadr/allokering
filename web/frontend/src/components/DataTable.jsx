import React from 'react'

// Visar en resultattabell ({columns, rows, row_count, truncated}).
export default function DataTable({ table }) {
  if (!table || table.columns.length === 0) {
    return <div className="empty-state">Ingen data.</div>
  }
  if (table.rows.length === 0) {
    return <div className="empty-state">Tabellen ar tom ({table.row_count} rader).</div>
  }
  return (
    <div className="table-wrap">
      <table className="data-table">
        <thead>
          <tr>
            <th className="row-num">#</th>
            {table.columns.map((c) => (
              <th key={c}>{c}</th>
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
          Visar de forsta {table.rows.length} av {table.row_count} raderna. Oppna i Excel for hela
          resultatet.
        </div>
      )}
    </div>
  )
}
