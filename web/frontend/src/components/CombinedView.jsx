import React, { useMemo, useState } from 'react'
import DropZone from './DropZone.jsx'
import FileSlot from './FileSlot.jsx'
import ResultPanel from './ResultPanel.jsx'
import { detect, runFlow } from '../api.js'

// Huvudvyn: alla "combined"-floden delar en uppsattning filrutor, precis
// som det gamla tkinter-GUI:t. Ladda upp en gang, kor valfri analys.

// "details" och "orders" ar samma fil (bestallningslinjer) - en gemensam ruta.
const logicalKey = (key) => (key === 'details' ? 'orders' : key)

const SLOT_LABELS = {
  orders: 'Bestallningslinjer',
  buffer: 'Buffertpallar',
  overview: 'Orderoversikt',
  dispatch: 'Dispatchpallar',
  saldo: 'Saldo / automation',
  items: 'Item option',
  not_putaway: 'Ej inlagrade',
  prognos: 'Prognosfil',
  campaign: 'Kampanjfil',
  max_csv: 'artikel_max.csv (valfri)',
}
const SLOT_ORDER = [
  'orders', 'buffer', 'overview', 'dispatch', 'saldo',
  'items', 'not_putaway', 'prognos', 'campaign', 'max_csv',
]

export default function CombinedView({ flows, onError }) {
  const [values, setValues] = useState({}) // logicalKey -> {name, file}
  const [busyId, setBusyId] = useState(null)
  const [status, setStatus] = useState('')
  const [result, setResult] = useState(null) // {label, data}

  // Union av alla filrutor over alla combined-floden.
  const slots = useMemo(() => {
    const map = new Map()
    for (const flow of flows) {
      for (const inp of flow.inputs) {
        const lk = logicalKey(inp.key)
        if (!map.has(lk)) {
          map.set(lk, {
            key: lk,
            label: SLOT_LABELS[lk] || inp.label,
            detect: new Set(inp.detect || []),
          })
        } else {
          ;(inp.detect || []).forEach((d) => map.get(lk).detect.add(d))
        }
      }
    }
    const keys = SLOT_ORDER.filter((k) => map.has(k)).concat(
      [...map.keys()].filter((k) => !SLOT_ORDER.includes(k)),
    )
    return keys.map((k) => ({ ...map.get(k), detect: [...map.get(k).detect] }))
  }, [flows])

  const setFile = (key, file) =>
    setValues((v) => ({ ...v, [key]: { name: file.name, file } }))
  const clearFile = (key) =>
    setValues((v) => {
      const next = { ...v }
      delete next[key]
      return next
    })

  const routeDropped = async (files) => {
    setStatus('Identifierar filer...')
    const unknown = []
    for (const file of files) {
      try {
        const { file_type } = await detect(file)
        const target = slots.find((s) => s.detect.includes(file_type))
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

  const missingFor = (flow) =>
    flow.inputs.filter((i) => i.required && !values[logicalKey(i.key)])

  const run = async (flow) => {
    if (missingFor(flow).length || busyId) return
    setBusyId(flow.id)
    setStatus('Kor ' + flow.label + '...')
    const fd = new FormData()
    for (const inp of flow.inputs) {
      const v = values[logicalKey(inp.key)]
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

  const filledCount = slots.filter((s) => values[s.key]).length

  return (
    <div className="flow-view">
      <div className="flow-header">
        <h1>Allokering &amp; analys</h1>
        <p className="flow-desc">
          Ladda upp filerna en gang - de delas mellan alla korningar nedan. Varje knapp visar
          vilka filer den behover: gron text = uppladdad, rod text = saknas.
        </p>
      </div>

      <section className="panel">
        <h2 className="panel-title">
          Datauppladdning · {filledCount}/{slots.length} filer inlagda
        </h2>
        <DropZone onFiles={routeDropped} busy={!!busyId} />
        <div className="pool-grid">
          {slots.map((slot) => (
            <FileSlot
              key={slot.key}
              slot={slot}
              entry={values[slot.key] || null}
              onSet={setFile}
              onClear={clearFile}
            />
          ))}
        </div>
      </section>

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
                        const filled = !!values[logicalKey(inp.key)]
                        const cls = filled ? 'ok' : inp.required ? 'missing' : 'opt'
                        const label = SLOT_LABELS[logicalKey(inp.key)] || inp.label
                        return (
                          <span key={inp.key} className={`file-tag ${cls}`}>
                            {filled ? '✓' : inp.required ? '✗' : '○'} {label}
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
