import React, { useState } from 'react'
import DataTable from './DataTable.jsx'
import { downloadUrl, openExcel, tableColumnText } from '../api.js'

// Visar resultatet från ett flöde: summeringskort, tabellflikar,
// fritext-rapport och logg.
export default function ResultPanel({ result, onError }) {
  const [activeTab, setActiveTab] = useState(result.tables[0]?.key || null)
  const [copyStatus, setCopyStatus] = useState('')

  const active = result.tables.find((t) => t.key === activeTab) || result.tables[0]
  const summaryEntries = Object.entries(result.summary || {})
  const copyColumns = result.flow_id === 'split-values'

  const handleExcel = async (key) => {
    try {
      await openExcel(result.session_id, key)
    } catch (err) {
      onError('Kunde inte öppna i Excel', String(err.message || err))
    }
  }

  const copyText = async (text) => {
    if (navigator.clipboard?.writeText) {
      await navigator.clipboard.writeText(text)
      return
    }

    const field = document.createElement('textarea')
    field.value = text
    field.setAttribute('readonly', '')
    field.style.position = 'fixed'
    field.style.top = '-9999px'
    document.body.appendChild(field)
    field.select()
    const ok = document.execCommand('copy')
    document.body.removeChild(field)
    if (!ok) throw new Error('Urklipp kunde inte användas.')
  }

  const handleCopyColumn = async (key, columnIndex) => {
    try {
      const text = await tableColumnText(result.session_id, key, columnIndex)
      await copyText(text)
      setCopyStatus('Kopierat.')
    } catch (err) {
      onError('Kunde inte kopiera', String(err.message || err))
    }
  }

  return (
    <div className="result-panel">
      {summaryEntries.length > 0 && (
        <div className="summary-cards">
          {summaryEntries.map(([label, value]) => (
            <div key={label} className="summary-card">
              <span className="summary-value">{String(value)}</span>
              <span className="summary-label">{label}</span>
            </div>
          ))}
        </div>
      )}

      {result.text && <pre className="report-text">{result.text}</pre>}

      {result.tables.length > 0 && (
        <>
          <div className="tab-toolbar">
            <div className="tab-row">
              {result.tables.map((t) => (
                <button
                  key={t.key}
                  className={`tab ${active?.key === t.key ? 'active' : ''}`}
                  onClick={() => setActiveTab(t.key)}
                >
                  {t.label}
                  <span className="tab-count">{t.table.row_count}</span>
                </button>
              ))}
            </div>
            {active && (
              <div className="tab-actions">
                {copyStatus && <span className="status-text">{copyStatus}</span>}
                <button className="btn-sm" onClick={() => handleExcel(active.key)}>
                  Öppna i Excel
                </button>
                <a className="btn-sm link" href={downloadUrl(result.session_id, active.key)}>
                  Ladda ner CSV
                </a>
              </div>
            )}
          </div>
          {active && (
            <DataTable
              table={active.table}
              copyColumnLabel={copyColumns ? 'Kopiera' : undefined}
              onCopyColumn={copyColumns ? (index) => handleCopyColumn(active.key, index) : undefined}
            />
          )}
        </>
      )}

      {result.tables.length === 0 && !result.text && (
        <div className="empty-state">Flödet kördes - ingen tabelldata.</div>
      )}

      {result.log?.length > 0 && (
        <details className="log-panel">
          <summary>Logg ({result.log.length} rader)</summary>
          <pre>{result.log.join('\n')}</pre>
        </details>
      )}
    </div>
  )
}
