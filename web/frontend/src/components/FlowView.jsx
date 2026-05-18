import React, { useState } from 'react'
import DropZone from './DropZone.jsx'
import FileSlot from './FileSlot.jsx'
import ResultPanel from './ResultPanel.jsx'
import { detect, runFlow } from '../api.js'

// Renderar ett enskilt flode: indatafalt fran flow-deskriptorn, kor-knapp
// och resultat. Komponenten remountas per flode (key={flow.id}).
export default function FlowView({ flow, onError }) {
  const [values, setValues] = useState({}) // key -> {name,file} | string
  const [busy, setBusy] = useState(false)
  const [status, setStatus] = useState('')
  const [result, setResult] = useState(null)

  const fileInputs = flow.inputs.filter((i) => i.type === 'file')
  const fieldInputs = flow.inputs.filter((i) => i.type !== 'file')

  const setFile = (key, file) =>
    setValues((v) => ({ ...v, [key]: { name: file.name, file } }))
  const clearFile = (key) =>
    setValues((v) => {
      const next = { ...v }
      delete next[key]
      return next
    })
  const setField = (key, val) => setValues((v) => ({ ...v, [key]: val }))

  const routeDroppedFiles = async (files) => {
    setStatus('Identifierar filer...')
    const unknown = []
    for (const file of files) {
      try {
        const { file_type } = await detect(file)
        const target = fileInputs.find((i) => (i.detect || []).includes(file_type))
        if (target) setFile(target.key, file)
        else unknown.push(file.name)
      } catch {
        unknown.push(file.name)
      }
    }
    setStatus('')
    if (unknown.length) {
      onError(
        'Okand filtyp',
        `Kunde inte sortera automatiskt: ${unknown.join(', ')}. Dra till ratt ruta eller anvand "Valj".`,
        'warn',
      )
    }
  }

  const missing = flow.inputs.filter((i) => i.required && !values[i.key])
  const canRun = missing.length === 0 && !busy

  const run = async () => {
    if (!canRun) return
    setBusy(true)
    setStatus('Kor ' + flow.label + '...')
    setResult(null)
    const fd = new FormData()
    for (const inp of flow.inputs) {
      const v = values[inp.key]
      if (v === undefined || v === '') continue
      if (inp.type === 'file') fd.append(inp.key, v.file, v.file.name)
      else fd.append(inp.key, v)
    }
    try {
      const data = await runFlow(flow.id, fd)
      setResult(data)
      setStatus('Klart.')
    } catch (err) {
      onError('Fel i ' + flow.label, String(err.message || err))
      setStatus('')
    } finally {
      setBusy(false)
    }
  }

  return (
    <div className="flow-view">
      <div className="flow-header">
        <h1>{flow.label}</h1>
        <p className="flow-desc">{flow.description}</p>
      </div>

      <section className="panel">
        <h2 className="panel-title">Indata</h2>

        {fileInputs.length > 0 && (
          <DropZone onFiles={routeDroppedFiles} busy={busy} />
        )}

        {fileInputs.length > 0 && (
          <div className="slots">
            {fileInputs.map((inp) => (
              <FileSlot
                key={inp.key}
                slot={inp}
                entry={values[inp.key] || null}
                onSet={setFile}
                onClear={clearFile}
              />
            ))}
          </div>
        )}

        {fieldInputs.length > 0 && (
          <div className="fields">
            {fieldInputs.map((inp) => (
              <div key={inp.key} className="field">
                <label className="field-label">
                  {inp.label}
                  {inp.required && <span className="req">*</span>}
                </label>
                {inp.type === 'textarea' ? (
                  <textarea
                    className="field-input"
                    rows={6}
                    value={values[inp.key] || ''}
                    onChange={(e) => setField(inp.key, e.target.value)}
                  />
                ) : (
                  <input
                    className="field-input"
                    type={inp.type === 'number' ? 'number' : 'text'}
                    value={values[inp.key] ?? (inp.default || '')}
                    onChange={(e) => setField(inp.key, e.target.value)}
                  />
                )}
              </div>
            ))}
          </div>
        )}

        {flow.inputs.length === 0 && (
          <p className="muted">Detta flode behover ingen indata.</p>
        )}

        <div className="run-row">
          <button className="btn primary" disabled={!canRun} onClick={run}>
            {busy ? 'Kor...' : 'Kor ' + flow.label}
          </button>
          <span className="status-text">{status}</span>
        </div>
        {missing.length > 0 && (
          <p className="hint">
            Kravs: {missing.map((m) => m.label).join(', ')}
          </p>
        )}
      </section>

      {result && (
        <section className="panel">
          <h2 className="panel-title">Resultat</h2>
          <ResultPanel key={result.session_id} result={result} onError={onError} />
        </section>
      )}
    </div>
  )
}
