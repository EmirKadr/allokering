import React, { useEffect, useState } from 'react'
import Sidebar from './components/Sidebar.jsx'
import FlowView from './components/FlowView.jsx'
import Modal from './components/Modal.jsx'
import { getFlows, health } from './api.js'

function useTheme() {
  const [theme, setTheme] = useState(
    () => localStorage.getItem('allok-theme') || 'dark',
  )
  useEffect(() => {
    document.documentElement.dataset.theme = theme
    localStorage.setItem('allok-theme', theme)
  }, [theme])
  return [theme, () => setTheme((t) => (t === 'dark' ? 'light' : 'dark'))]
}

export default function App() {
  const [theme, toggleTheme] = useTheme()
  const [flows, setFlows] = useState([])
  const [activeId, setActiveId] = useState(null)
  const [info, setInfo] = useState({ version: '', title: 'Allokering' })
  const [modal, setModal] = useState(null)
  const [loadError, setLoadError] = useState('')

  useEffect(() => {
    health()
      .then(setInfo)
      .catch(() => {})
    getFlows()
      .then((list) => {
        setFlows(list)
        setActiveId(list[0]?.id || null)
      })
      .catch((err) => setLoadError(String(err.message || err)))
  }, [])

  const showError = (title, body, tone = 'error') =>
    setModal({ title, tone, body: <p>{body}</p> })

  const showHelp = () =>
    setModal({
      title: 'Om appen',
      tone: 'info',
      body: (
        <div className="help-body">
          <p>
            Hela allokerings-appen som ett <strong>API-styrt</strong> granssnitt. Varje flode i
            menyn motsvarar ett CLI-kommando och kor exakt samma motor.
          </p>
          <p>
            <strong>Indata:</strong> slapp filer i drop-zonen sa sorteras de automatiskt, eller
            valj per ruta. Textfalt fylls i for hand.
          </p>
          <p>
            <strong>Resultat:</strong> flikar per tabell. Oppna i Excel eller ladda ner som CSV.
          </p>
          <p className="muted">
            CLI och det gamla tkinter-GUI:t ar oroda - detta ar ett nytt lager ovanpa samma logik.
          </p>
        </div>
      ),
    })

  const activeFlow = flows.find((f) => f.id === activeId)

  return (
    <div className="app">
      <header className="topbar">
        <div className="brand">
          <span className="brand-mark">A</span>
          <div>
            <div className="brand-title">Allokering</div>
            <div className="brand-sub">API-styrd app · v{info.version || '...'}</div>
          </div>
        </div>
        <div className="topbar-actions">
          <button className="btn ghost" onClick={toggleTheme} title="Vaxla tema">
            {theme === 'dark' ? '☀ Ljust' : '☾ Morkt'}
          </button>
          <button className="btn ghost" onClick={showHelp}>
            ? Hjalp
          </button>
        </div>
      </header>

      {loadError ? (
        <div className="load-error">
          Kunde inte ladda floden: {loadError}
          <div className="muted">Ar API-servern igang?</div>
        </div>
      ) : (
        <div className="shell">
          <Sidebar flows={flows} activeId={activeId} onSelect={setActiveId} />
          <main className="content">
            {activeFlow ? (
              <FlowView key={activeFlow.id} flow={activeFlow} onError={showError} />
            ) : (
              <div className="empty-state big">Laddar...</div>
            )}
          </main>
        </div>
      )}

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
