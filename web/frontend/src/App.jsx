import React, { useEffect, useMemo, useState } from 'react'
import Sidebar from './components/Sidebar.jsx'
import FlowView from './components/FlowView.jsx'
import CombinedView from './components/CombinedView.jsx'
import UploadView from './components/UploadView.jsx'
import Modal from './components/Modal.jsx'
import { getFlows, health } from './api.js'

const UPLOAD_ID = '__upload__'
const COMBINED_ID = '__combined__'

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
  const [activeId, setActiveId] = useState(UPLOAD_ID)
  const [info, setInfo] = useState({ version: '' })
  const [modal, setModal] = useState(null)
  const [loadError, setLoadError] = useState('')

  // Uppladdade filer lever har sa de delas mellan Datauppladdning-sidan
  // och korningarna i huvudvyn. logicalKey -> { name, file }.
  const [poolFiles, setPoolFiles] = useState({})

  useEffect(() => {
    health()
      .then(setInfo)
      .catch(() => {})
    getFlows()
      .then(setFlows)
      .catch((err) => setLoadError(String(err.message || err)))
  }, [])

  const combinedFlows = useMemo(
    () => flows.filter((f) => f.view === 'combined'),
    [flows],
  )
  const soloFlows = useMemo(() => flows.filter((f) => f.view === 'solo'), [flows])

  const setPoolFile = (key, file) =>
    setPoolFiles((v) => ({ ...v, [key]: { name: file.name, file } }))
  const clearPoolFile = (key) =>
    setPoolFiles((v) => {
      const next = { ...v }
      delete next[key]
      return next
    })

  // Sidebar-grupper: huvudvy (datauppladdning + analys), sedan solo-floden.
  const navGroups = useMemo(() => {
    const groups = [
      {
        name: 'Huvudvy',
        items: [
          {
            id: UPLOAD_ID,
            label: 'Datauppladdning',
            description: 'Ladda upp alla filer pa ett stalle.',
          },
          {
            id: COMBINED_ID,
            label: 'Allokering & analys',
            description: 'Allokering, ordersaldo, kontroller m.m. - delade filer.',
          },
        ],
      },
    ]
    for (const flow of soloFlows) {
      let group = groups.find((g) => g.name === flow.category)
      if (!group) {
        group = { name: flow.category, items: [] }
        groups.push(group)
      }
      group.items.push({ id: flow.id, label: flow.label, description: flow.description })
    }
    return groups
  }, [soloFlows])

  const showError = (title, body, tone = 'error') =>
    setModal({ title, tone, body: <p>{body}</p> })

  const showHelp = () =>
    setModal({
      title: 'Om appen',
      tone: 'info',
      body: (
        <div className="help-body">
          <p>
            Hela allokerings-appen som ett <strong>API-styrt</strong> granssnitt. Varje flode kor
            exakt samma motor som CLI:t.
          </p>
          <p>
            <strong>Datauppladdning</strong> ar en egen sida - ladda upp filerna en gang dar.
          </p>
          <p>
            <strong>Allokering & analys</strong> samlar allokering, ordersaldo, LYX, pafyllnadsprio,
            kontroller och prognos. Varje knapp visar vilka uppladdade filer den behover.
          </p>
          <p>
            <strong>Eftersok</strong> och <strong>Data & verktyg</strong> har egna vyer.
          </p>
          <p className="muted">
            CLI och det gamla tkinter-GUI:t ar oroda - detta ar ett nytt lager ovanpa samma logik.
          </p>
        </div>
      ),
    })

  const activeSolo = soloFlows.find((f) => f.id === activeId)

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
          <Sidebar groups={navGroups} activeId={activeId} onSelect={setActiveId} />
          <main className="content">
            {flows.length === 0 ? (
              <div className="empty-state big">Laddar...</div>
            ) : activeId === UPLOAD_ID ? (
              <UploadView
                flows={combinedFlows}
                files={poolFiles}
                onSet={setPoolFile}
                onClear={clearPoolFile}
                onError={showError}
              />
            ) : activeId === COMBINED_ID ? (
              <CombinedView
                flows={combinedFlows}
                files={poolFiles}
                onError={showError}
                onGoToUpload={() => setActiveId(UPLOAD_ID)}
              />
            ) : activeSolo ? (
              <FlowView key={activeSolo.id} flow={activeSolo} onError={showError} />
            ) : (
              <div className="empty-state big">Valj ett flode.</div>
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
