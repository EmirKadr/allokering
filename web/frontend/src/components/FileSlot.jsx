import React, { useRef, useState } from 'react'

// En filruta. Accepterar aven egen drop/klick for manuell tilldelning
// (motsvarar GUI:ts enskilda filval).
export default function FileSlot({ slot, entry, onSet, onClear }) {
  const inputRef = useRef(null)
  const [hover, setHover] = useState(false)
  const filled = !!entry

  const handleDrop = (e) => {
    e.preventDefault()
    e.stopPropagation()
    setHover(false)
    const file = e.dataTransfer.files?.[0]
    if (file) onSet(slot.key, file)
  }

  return (
    <div
      className={`file-slot ${filled ? 'file-slot-filled' : ''} ${hover ? 'file-slot-hover' : ''}`}
      onDragOver={(e) => {
        e.preventDefault()
        e.stopPropagation()
        setHover(true)
      }}
      onDragLeave={() => setHover(false)}
      onDrop={handleDrop}
    >
      <input
        ref={inputRef}
        type="file"
        hidden
        accept=".csv,.xlsx,.xlsm,.xls,.txt"
        onChange={(e) => {
          const file = e.target.files?.[0]
          if (file) onSet(slot.key, file)
          e.target.value = ''
        }}
      />
      <div className="file-slot-main">
        <div className="file-slot-label">
          {slot.label}
          {slot.required && <span className="req">*</span>}
        </div>
        <div className="file-slot-name">
          {filled ? entry.name : <span className="muted">Ingen fil vald</span>}
        </div>
      </div>
      <div className="file-slot-actions">
        <span className={`status-pill ${filled ? 'ok' : 'none'}`}>
          {filled ? 'Uppladdad' : 'Ej fil'}
        </span>
        <button className="btn-sm" onClick={() => inputRef.current?.click()}>
          Valj
        </button>
        <button
          className="btn-sm danger"
          disabled={!filled}
          onClick={() => onClear(slot.key)}
        >
          ✕
        </button>
      </div>
    </div>
  )
}
