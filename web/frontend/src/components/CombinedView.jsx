import React, { useState } from 'react'
import DropArea from './DropArea.jsx'
import FilePickerButton from './FilePickerButton.jsx'
import ResultPanel from './ResultPanel.jsx'
import { runFlow } from '../api.js'
import { routeFilesToSlots } from '../fileRouting.js'
import { fileInputKey, slotLabel } from '../poolSlots.js'

// Huvudvyn: körningsknappar för alla "combined"-flöden. Filerna kommer
// från den delade datapoolen och kan även släppas direkt i denna vy.
export default function CombinedView({ flows, allSlots, files, onSet, onError, onGoToUpload }) {
  const [busyId, setBusyId] = useState(null)
  const [status, setStatus] = useState('')
  const [dropStatus, setDropStatus] = useState('')
  const [result, setResult] = useState(null) // { label, data }

  const routeDropped = (dropped) => {
    if (busyId) {
      setDropStatus('Vänta tills körningen är klar.')
      return
    }
    return routeFilesToSlots(dropped, allSlots, onSet, {
      setStatus: setDropStatus,
      onUnknown: (unknown) =>
        onError(
          'Okänd filtyp',
          `Kunde inte sortera automatiskt: ${unknown.join(', ')}. Använd "Välj filer" på Datauppladdning om filen saknar igenkännbar typ.`,
          'warn',
        ),
    })
  }

  const missingFor = (flow) =>
    flow.inputs.filter((i) => i.required && !files[fileInputKey(i)])

  const run = async (flow) => {
    if (missingFor(flow).length || busyId) return
    setBusyId(flow.id)
    setStatus('Kör ' + flow.label + '...')
    const fd = new FormData()
    for (const inp of flow.inputs) {
      const v = files[fileInputKey(inp)]
      if (v) fd.append(inp.key, v.file, v.file.name)
    }
    try {
      const data = await runFlow(flow.id, fd)
      setResult({ label: flow.label, data })
      setStatus('Klart: ' + flow.label)
    } catch (err) {
      onError('Fel i ' + flow.label, String(err.message || err))
      setStatus('')
    } finally {
      setBusyId(null)
    }
  }

  // Körningsknappar grupperade per kategori.
  const groups = []
  for (const flow of flows) {
    let group = groups.find((g) => g.name === flow.category)
    if (!group) {
      group = { name: flow.category, flows: [] }
      groups.push(group)
    }
    group.flows.push(flow)
  }

  const anyFile = Object.keys(files).length > 0

  return (
    <DropArea onFiles={routeDropped}>
      <div className="flow-view">
        <div className="flow-header">
          <h1>Allokering &amp; analys</h1>
          <p className="flow-desc">
            Kör valfri analys. Släpp filer var som helst i vyn eller gå via Datauppladdning.
            Grön text = uppladdad, röd text = saknas.
          </p>
        </div>

        {!anyFile && (
          <div className="upload-prompt">
            <span>Inga filer uppladdade an.</span>
            <FilePickerButton onFiles={routeDropped} disabled={!!busyId}>
              Välj filer
            </FilePickerButton>
            <button type="button" className="btn-sm" onClick={onGoToUpload}>
              Gå till Datauppladdning
            </button>
          </div>
        )}

        <section className="panel">
          <div className="panel-head">
            <h2 className="panel-title">Körningar</h2>
            {anyFile && (
              <FilePickerButton onFiles={routeDropped} disabled={!!busyId}>
                Välj fler filer
              </FilePickerButton>
            )}
          </div>
          <p className="drop-hint">
            Hela vyn tar emot filer. Uppladdade filer sparas i samma datapool som
            Datauppladdning.
          </p>
          {dropStatus && <p className="status-text upload-status">{dropStatus}</p>}
          {groups.map((group) => (
            <div key={group.name} className="action-group">
              <div className="action-group-title">{group.name}</div>
              <div className="action-grid">
                {group.flows.map((flow) => {
                  const ready = missingFor(flow).length === 0
                  return (
                    <div key={flow.id} className={`action-card ${ready ? 'ready' : ''}`}>
                      <h3 className="action-title">{flow.label}</h3>
                      <p className="action-desc">{flow.description}</p>
                      <div className="action-files">
                        {flow.inputs.map((inp) => {
                          const filled = !!files[fileInputKey(inp)]
                          const cls = filled ? 'ok' : inp.required ? 'missing' : 'opt'
                          return (
                            <span key={inp.key} className={`file-tag ${cls}`}>
                              {filled ? '✓' : inp.required ? '✗' : '○'} {slotLabel(inp.key)}
                              {!inp.required && ' (valfri)'}
                            </span>
                          )
                        })}
                      </div>
                      <button
                        className="btn primary action-btn"
                        disabled={!ready || !!busyId}
                        onClick={() => run(flow)}
                      >
                        {busyId === flow.id ? 'Kör...' : 'Kör ' + flow.label}
                      </button>
                    </div>
                  )
                })}
              </div>
            </div>
          ))}
          {status && (
            <div className="run-row">
              <span className="status-text">{status}</span>
            </div>
          )}
        </section>

        {result && (
          <section className="panel">
            <h2 className="panel-title">Resultat — {result.label}</h2>
            <ResultPanel key={result.data.session_id} result={result.data} onError={onError} />
          </section>
        )}
      </div>
    </DropArea>
  )
}
