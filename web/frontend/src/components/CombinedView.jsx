import React, { useState } from 'react'
import ResultPanel from './ResultPanel.jsx'
import { runFlow } from '../api.js'
import { logicalKey, slotLabel } from '../poolSlots.js'

// Huvudvyn: korningsknappar for alla "combined"-floden. Filerna laddas upp
// pa den separata Datauppladdning-sidan och kommer in via props.
export default function CombinedView({ flows, files, onError, onGoToUpload }) {
  const [busyId, setBusyId] = useState(null)
  const [status, setStatus] = useState('')
  const [result, setResult] = useState(null) // { label, data }

  const missingFor = (flow) =>
    flow.inputs.filter((i) => i.required && !files[logicalKey(i.key)])

  const run = async (flow) => {
    if (missingFor(flow).length || busyId) return
    setBusyId(flow.id)
    setStatus('Kor ' + flow.label + '...')
    const fd = new FormData()
    for (const inp of flow.inputs) {
      const v = files[logicalKey(inp.key)]
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

  // Korningsknappar grupperade per kategori.
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
    <div className="flow-view">
      <div className="flow-header">
        <h1>Allokering &amp; analys</h1>
        <p className="flow-desc">
          Kor valfri analys. Varje knapp visar vilka filer den behover - gron text = uppladdad,
          rod text = saknas. Filerna laddas upp under <strong>Datauppladdning</strong>.
        </p>
      </div>

      {!anyFile && (
        <div className="upload-prompt">
          <span>Inga filer uppladdade an.</span>
          <button className="btn-sm" onClick={onGoToUpload}>
            Ga till Datauppladdning
          </button>
        </div>
      )}

      <section className="panel">
        <h2 className="panel-title">Korningar</h2>
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
                        const filled = !!files[logicalKey(inp.key)]
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
                      {busyId === flow.id ? 'Kor...' : 'Kor ' + flow.label}
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
  )
}
