import React, { useState } from 'react'
import DropArea from './DropArea.jsx'
import ResultPanel from './ResultPanel.jsx'
import { runFlow } from '../api.js'
import { routeFilesToSlots } from '../fileRouting.js'
import { fileInputKey, slotLabel } from '../poolSlots.js'

// Renderar ett enskilt flöde: indatafält från flow-deskriptorn, kör-knapp
// och resultat. Komponenten remountas per flöde (key={flow.id}).
export default function FlowView({ flow, allSlots, files, onSet, onError, onGoToUpload }) {
  const [values, setValues] = useState({}) // text/number/textarea values
  const [busy, setBusy] = useState(false)
  const [status, setStatus] = useState('')
  const [result, setResult] = useState(null)

  const fileInputs = flow.inputs.filter((i) => i.type === 'file')
  const fieldInputs = flow.inputs.filter((i) => i.type !== 'file')

  const setFile = (key, file) => onSet(key, file)
  const setField = (key, val) => setValues((v) => ({ ...v, [key]: val }))

  const routeDroppedFiles = (dropped) => {
    if (busy) {
      setStatus('Vänta tills körningen är klar.')
      return
    }
    return routeFilesToSlots(dropped, allSlots, setFile, {
      setStatus,
      onNoSlots: () =>
        onError('Inga filfält', 'Den här vyn tar inte emot filer.', 'warn'),
      onUnknown: (unknown) =>
        onError(
          'Okänd filtyp',
          `Kunde inte sortera automatiskt: ${unknown.join(', ')}. Kontrollera filen i Datauppladdning.`,
          'warn',
        ),
    })
  }

  const missing = flow.inputs.filter((i) => {
    if (!i.required) return false
    if (i.type === 'file') return !files[fileInputKey(i)]
    return !values[i.key]
  })
  const canRun = missing.length === 0 && !busy

  const run = async () => {
    if (!canRun) return
    setBusy(true)
    setStatus('Kör ' + flow.label + '...')
    setResult(null)
    const fd = new FormData()
    for (const inp of flow.inputs) {
      const v = inp.type === 'file' ? files[fileInputKey(inp)] : values[inp.key]
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
    <DropArea onFiles={routeDroppedFiles}>
      <div className="flow-view">
        <div className="flow-header">
          <h1>{flow.label}</h1>
          <p className="flow-desc">{flow.description}</p>
        </div>

        <section className="panel">
          <div className="panel-head">
            <h2 className="panel-title">Indata</h2>
            {fileInputs.length > 0 && (
              <button type="button" className="btn-sm" onClick={onGoToUpload}>
                Gå till Datauppladdning
              </button>
            )}
          </div>

          {fileInputs.length > 0 && (
            <>
              <p className="drop-hint">
                Filerna hämtas från centrala Datauppladdning. Statusen nedan visar vad detta
                flöde använder.
              </p>
              <div className="flow-file-list">
                {fileInputs.map((inp) => {
                  const poolKey = fileInputKey(inp)
                  const filled = !!files[poolKey]
                  const cls = filled ? 'ok' : inp.required ? 'missing' : 'opt'
                  return (
                    <div
                      key={inp.key}
                      className={`flow-file-row ${filled ? 'flow-file-row-filled' : ''}`}
                    >
                      <span className={`file-tag ${cls}`}>
                        {filled ? '✓' : inp.required ? '✗' : '○'} {slotLabel(poolKey)}
                        {!inp.required && ' (valfri)'}
                      </span>
                      <span className="flow-file-name">
                        {filled ? files[poolKey].name : 'Ingen fil i Datauppladdning'}
                      </span>
                    </div>
                  )
                })}
              </div>
            </>
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
            <p className="muted">Detta flöde behöver ingen indata.</p>
          )}

          <div className="run-row">
            <button className="btn primary" disabled={!canRun} onClick={run}>
              {busy ? 'Kör...' : 'Kör ' + flow.label}
            </button>
            <span className="status-text">{status}</span>
          </div>
          {missing.length > 0 && (
            <p className="hint">
              Krävs: {missing.map((m) => m.label).join(', ')}
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
    </DropArea>
  )
}
