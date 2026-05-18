import React from 'react'

// Ersätter tkinter messagebox/popups - fel, hjälp och bekräftelser.
export default function Modal({ open, title, tone = 'info', onClose, children }) {
  if (!open) return null
  return (
    <div className="modal-backdrop" onClick={onClose}>
      <div className={`modal modal-${tone}`} onClick={(e) => e.stopPropagation()}>
        <header className="modal-head">
          <h3>{title}</h3>
          <button className="icon-btn" onClick={onClose} aria-label="Stäng">
            ✕
          </button>
        </header>
        <div className="modal-body">{children}</div>
        <footer className="modal-foot">
          <button className="btn" onClick={onClose}>
            Stäng
          </button>
        </footer>
      </div>
    </div>
  )
}
