import React, { useRef, useState } from 'react'

// Stor drag&drop-zon som auto-routar slappta filer till ratt slot
// (samma filtypsdetektering som GUI:ts globala drop-handler).
export default function DropZone({ onFiles, busy }) {
  const inputRef = useRef(null)
  const [hover, setHover] = useState(false)

  const handleDrop = (e) => {
    e.preventDefault()
    setHover(false)
    if (busy) return
    const files = [...e.dataTransfer.files]
    if (files.length) onFiles(files)
  }

  const handlePick = (e) => {
    const files = [...e.target.files]
    if (files.length) onFiles(files)
    e.target.value = ''
  }

  return (
    <div
      className={`dropzone ${hover ? 'dropzone-hover' : ''} ${busy ? 'dropzone-busy' : ''}`}
      onDragOver={(e) => {
        e.preventDefault()
        if (!busy) setHover(true)
      }}
      onDragLeave={() => setHover(false)}
      onDrop={handleDrop}
      onClick={() => !busy && inputRef.current?.click()}
    >
      <input
        ref={inputRef}
        type="file"
        multiple
        hidden
        onChange={handlePick}
        accept=".csv,.xlsx,.xlsm,.xls,.txt"
      />
      <div className="dropzone-icon">⤓</div>
      <div className="dropzone-title">Slapp filer har eller klicka for att valja</div>
      <div className="dropzone-sub">
        Filtypen kanns igen automatiskt - bestallningslinjer, buffertpallar, saldo och item
        sorteras till ratt ruta.
      </div>
    </div>
  )
}
