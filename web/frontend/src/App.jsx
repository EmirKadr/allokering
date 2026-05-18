import React, { useEffect, useMemo, useRef, useState } from 'react'
import Sidebar from './components/Sidebar.jsx'
import FlowView from './components/FlowView.jsx'
import CombinedView from './components/CombinedView.jsx'
import UploadView from './components/UploadView.jsx'
import Modal from './components/Modal.jsx'
import { getFlows, health, updateObservations } from './api.js'
import { deriveSlots } from './poolSlots.js'

const UPLOAD_ID = '__upload__'
const COMBINED_ID = '__combined__'
const HIDDEN_FLOW_IDS = new Set(['observations-update', 'observations-sync'])

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
  const [autoStatus, setAutoStatus] = useState('')
  const lastObservationsFile = useRef('')

  // Uppladdade filer lever här så de delas mellan Datauppladdning-sidan
  // och alla körningar. logicalKey -> { name, file }.
  const [poolFiles, setPoolFiles] = useState({})

  useEffect(() => {
    health()
      .then(setInfo)
      .catch(() => {})
    getFlows()
      .then(setFlows)
      .catch((err) => setLoadError(String(err.message || err)))
  }, [])

  const visibleFlows = useMemo(
    () => flows.filter((f) => !HIDDEN_FLOW_IDS.has(f.id)),
    [flows],
  )
  const combinedFlows = useMemo(
    () => visibleFlows.filter((f) => f.view === 'combined'),
    [visibleFlows],
  )
  const soloFlows = useMemo(
    () => visibleFlows.filter((f) => f.view === 'solo'),
    [visibleFlows],
  )
  const allFileSlots = useMemo(() => deriveSlots(visibleFlows), [visibleFlows])

  const triggerObservationsUpdate = async (file) => {
    const signature = `${file.name}:${file.size}:${file.lastModified || ''}`
    if (lastObservationsFile.current === signature) return
    lastObservationsFile.current = signature
    setAutoStatus('Observations uppdateras från buffertpallar...')
    try {
      const result = await updateObservations(file)
      if (result.new_rows > 0) {
        setAutoStatus(
          `Observations uppdaterad: ${result.new_rows} nya pallid. Artikel_max: ${result.article_max_rows} rader.`,
        )
      } else {
        setAutoStatus('Observations kontrollerad: inga nya pallid.')
      }
    } catch (err) {
      lastObservationsFile.current = ''
      setAutoStatus('')
      showError('Observations kunde inte uppdateras', String(err.message || err), 'warn')
    }
  }

  const setPoolFile = (key, file) => {
    setPoolFiles((v) => ({ ...v, [key]: { name: file.name, file } }))
    if (key === 'buffer') triggerObservationsUpdate(file)
  }
  const clearPoolFile = (key) =>
    setPoolFiles((v) => {
      const next = { ...v }
      delete next[key]
      return next
    })

  // Sidebar-grupper: huvudvy (datauppladdning + analys), sedan solo-flöden.
  const navGroups = useMemo(() => {
    const groups = [
      {
        name: 'Huvudvy',
        items: [
          {
            id: UPLOAD_ID,
            label: 'Datauppladdning',
            description: 'Ladda upp alla filer på ett ställe.',
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
            Hela allokerings-appen som ett <strong>API-styrt</strong> gränssnitt. Varje flöde kör
            exakt samma motor som CLI:t.
          </p>
          <p>
            <strong>Datauppladdning</strong> är en egen sida för filstatus, men filer kan släppas
            i hela den aktiva vyn.
          </p>
          <p>
            <strong>Allokering & analys</strong> samlar allokering, ordersaldo, LYX, påfyllnadsprio,
            kontroller och prognos. Varje knapp visar vilka uppladdade filer den behöver.
          </p>
          <p>
            <strong>Eftersök</strong> och <strong>Data & verktyg</strong> har egna vyer.
          </p>
          <p className="muted">
            CLI och det gamla tkinter-GUI:t är orörda - detta är ett nytt lager ovanpå samma logik.
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
          <button className="btn ghost" onClick={toggleTheme} title="Växla tema">
            {theme === 'dark' ? '☀ Ljust' : '☾ Mörkt'}
          </button>
          <button className="btn ghost" onClick={showHelp}>
            ? Hjälp
          </button>
        </div>
      </header>

      {loadError ? (
        <div className="load-error">
          Kunde inte ladda flöden: {loadError}
          <div className="muted">Är API-servern igång?</div>
        </div>
      ) : (
        <div className="shell">
          <Sidebar groups={navGroups} activeId={activeId} onSelect={setActiveId} />
          <main className="content">
            {flows.length === 0 ? (
              <div className="empty-state big">Laddar...</div>
            ) : activeId === UPLOAD_ID ? (
              <UploadView
                slots={allFileSlots}
                files={poolFiles}
                onSet={setPoolFile}
                onClear={clearPoolFile}
                onError={showError}
                autoStatus={autoStatus}
              />
            ) : activeId === COMBINED_ID ? (
              <CombinedView
                flows={combinedFlows}
                allSlots={allFileSlots}
                files={poolFiles}
                onSet={setPoolFile}
                onError={showError}
                onGoToUpload={() => setActiveId(UPLOAD_ID)}
              />
            ) : activeSolo ? (
              <FlowView
                key={activeSolo.id}
                flow={activeSolo}
                allSlots={allFileSlots}
                files={poolFiles}
                onSet={setPoolFile}
                onError={showError}
                onGoToUpload={() => setActiveId(UPLOAD_ID)}
              />
            ) : (
              <div className="empty-state big">Välj ett flöde.</div>
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
