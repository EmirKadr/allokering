import React, { useRef } from 'react'

// En filrad med manuell filväljare. Drag & drop hanteras av hela vyn.
export default function FileSlot({ slot, entry, onSet, onClear }) {
  const inputRef = useRef(null)
  const filled = !!entry

  return (
    <div className={`file-slot ${filled ? 'file-slot-filled' : ''}`}>
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
        <button type="button" className="btn-sm" onClick={() => inputRef.current?.click()}>
          Välj
        </button>
        <button
          type="button"
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
