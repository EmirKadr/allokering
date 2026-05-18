import React, { useEffect, useMemo, useState } from 'react'
import DropZone from './components/DropZone.jsx'
import FileSlot from './components/FileSlot.jsx'
import DataTable from './components/DataTable.jsx'
import Modal from './components/Modal.jsx'
import { allocate, detect, downloadUrl, health, openExcel } from './api.js'

const SLOTS = [
  { key: 'orders', label: 'Bestallningslinjer', required: true },
  { key: 'buffer', label: 'Buffertpallar', required: true },
  { key: 'saldo', label: 'Saldo / automation', required: false },
  { key: 'items', label: 'Item option', required: false },
  { key: 'not_putaway', label: 'Ej inlagrade', required: false },
]

const TABS = [
  { key: 'result', label: 'Resultat', metric: 'result_rows' },
  { key: 'near_miss', label: 'Near-miss', metric: 'near_miss_rows' },
  { key: 'refill_hp', label: 'Refill Huvudplock', metric: 'refill_hp_rows' },
  { key: 'refill_autostore', label: 'Refill AutoStore', metric: 'refill_autostore_rows' },
  { key: 'pallet_spaces', label: 'Pallplatser', metric: 'pallet_space_rows' },
]

const EMPTY_SLOTS = Object.fromEntries(SLOTS.map((s) => [s.key, null]))

export default function App() {
  const [slots, setSlots] = useState(EMPTY_SLOTS)
  const [busy, setBusy] = useState(false)
  const [result, setResult] = useState(null)
  const [activeTab, setActiveTab] = useState('result')
  const [modal, setModal] = useState(null)
  const [status, setStatus] = useState('')
  const [info, setInfo] = useState({ version: '', title: 'Allokering' })

  useEffect(() => {
    health()
      .then(setInfo)
      .catch(() => setInfo({ version: '?', title: 'Allokering' }))
  }, [])

  const setSlot = (key, file) =>
    setSlots((prev) => ({ ...prev, [key]: { name: file.name, file } }))
  const clearSlot = (key) => setSlots((prev) => ({ ...prev, [key]: null }))

  const onFiles = async (files) => {
    setStatus('Identifierar filer...')
    const unknown = []
    for (const file of files) {
      try {
        const { slot } = await detect(file)
        if (slot) {
          setSlot(slot, file)
        } else {
          unknown.push(file.name)
        }
      } catch {
        unknown.push(file.name)
      }
    }
    setStatus('')
    if (unknown.length) {
      setModal({
        title: 'Okand filtyp',
        tone: 'warn',
        body: (
          <>
            <p>Foljande filer kunde inte sorteras automatiskt:</p>
            <ul>
              {unknown.map((n) => (
                <li key={n}>{n}</li>
              ))}
            </ul>
            <p>Dra dem direkt till ratt ruta eller anvand knappen "Valj".</p>
          </>
        ),
      })
    }
  }

  const canRun = slots.orders && slots.buffer && !busy

  const runAllocation = async () => {
    if (!canRun) return
    setBusy(true)
    setStatus('Kor allokering...')
    setResult(null)
    try {
      const data = await allocate(slots)
      setResult(data)
      setActiveTab('result')
      setStatus('Allokering klar.')
    } catch (err) {
      setModal({ title: 'Fel under allokering', tone: 'error', body: <p>{String(err.message || err)}</p> })
      setStatus('')
    } finally {
      setBusy(false)
    }
  }

  const handleOpenExcel = async (key) => {
    if (!result) return
    try {
      await openExcel(result.session_id, key)
      setStatus('Oppnar i Excel...')
    } catch (err) {
      setModal({ title: 'Kunde inte oppna i Excel', tone: 'error', body: <p>{String(err.message || err)}</p> })
    }
  }

  const showHelp = () =>
    setModal({
      title: 'Sa har fungerar demon',
      tone: 'info',
      body: (
        <div className="help-body">
          <p>
            <strong>1. Lagg in filer.</strong> Slapp dem i den stora zonen - filtypen kanns igen
            automatiskt - eller dra/valj dem per ruta. Bestallningslinjer och buffertpallar kravs,
            ovriga ar valfria.
          </p>
          <p>
            <strong>2. Kor allokering.</strong> Samma motor som CLI-kommandot{' '}
            <code>allocate</code>: Helpall → AutoStore → Huvudplock, FIFO, near-miss-loggning.
          </p>
          <p>
            <strong>3. Granska resultatet.</strong> Flikar for resultat, near-miss, refill och
            pallplatser. Oppna valfri tabell i Excel eller ladda ner som CSV.
          </p>
          <p className="muted">
            API-styrd: allt gar via samma HTTP-API som kan koras som webbapp senare. CLI:t paverkas
            inte.
          </p>
        </div>
      ),
    })

  const activeTable = useMemo(
    () => (result ? result.tables[activeTab] : null),
    [result, activeTab],
  )

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <span className="brand-mark">A</span>
          <div>
            <div className="brand-title">{info.title}</div>
            <div className="brand-sub">Allokeringsdemo · v{info.version}</div>
          </div>
        </div>
        <button className="btn ghost" onClick={showHelp}>
          ? Hjalp
        </button>
      </header>

      <main className="layout">
        <section className="panel inputs">
          <h2 className="panel-title">Indata</h2>
          <DropZone onFiles={onFiles} busy={busy} />
          <div className="slots">
            {SLOTS.map((slot) => (
              <FileSlot
                key={slot.key}
                slot={slot}
                entry={slots[slot.key]}
                onSet={setSlot}
                onClear={clearSlot}
              />
            ))}
          </div>
          <div className="run-row">
            <button className="btn primary" disabled={!canRun} onClick={runAllocation}>
              {busy ? 'Kor...' : 'Kor allokering'}
            </button>
            <span className="status-text">{status}</span>
          </div>
          {!slots.orders || !slots.buffer ? (
            <p className="hint">Bestallningslinjer och buffertpallar maste valjas.</p>
          ) : null}
        </section>

        <section className="panel results">
          <h2 className="panel-title">Resultat</h2>
          {!result ? (
            <div className="empty-state big">
              Inga resultat an. Lagg in filer och kor allokeringen.
            </div>
          ) : (
            <>
              <div className="summary-cards">
                {TABS.map((t) => (
                  <button
                    key={t.key}
                    className={`summary-card ${activeTab === t.key ? 'active' : ''}`}
                    onClick={() => setActiveTab(t.key)}
                  >
                    <span className="summary-value">{result.summary[t.metric]}</span>
                    <span className="summary-label">{t.label}</span>
                  </button>
                ))}
              </div>

              <div className="tab-toolbar">
                <div className="tab-row">
                  {TABS.map((t) => (
                    <button
                      key={t.key}
                      className={`tab ${activeTab === t.key ? 'active' : ''}`}
                      onClick={() => setActiveTab(t.key)}
                    >
                      {t.label}
                    </button>
                  ))}
                </div>
                <div className="tab-actions">
                  <button className="btn-sm" onClick={() => handleOpenExcel(activeTab)}>
                    Oppna i Excel
                  </button>
                  <a
                    className="btn-sm link"
                    href={downloadUrl(result.session_id, activeTab)}
                  >
                    Ladda ner CSV
                  </a>
                </div>
              </div>

              <DataTable table={activeTable} />

              {result.log?.length > 0 && (
                <details className="log-panel">
                  <summary>Logg ({result.log.length} rader)</summary>
                  <pre>{result.log.join('\n')}</pre>
                </details>
              )}
            </>
          )}
        </section>
      </main>

      <Modal
        open={!!modal}
        title={modal?.title}
        tone={modal?.tone}
        onClose={() => setModal(null)}
      >
        {modal?.body}
      </Modal>
    </div>
  )
}
