import React, { useRef, useState } from 'react'

// Gör hela det inneslutna området till en drop-yta. Filer kan släppas
// var som helst i vyn - ingen särskild drop-ruta behövs.
export default function DropArea({ onFiles, children, disabled = false }) {
  const [over, setOver] = useState(false)
  const depth = useRef(0)

  const hasFiles = (e) => [...(e.dataTransfer?.types || [])].includes('Files')
  const reset = () => {
    depth.current = 0
    setOver(false)
  }

  return (
    <div
      className={`drop-area ${over ? 'drop-area-over' : ''}`}
      onDragEnter={(e) => {
        if (!hasFiles(e)) return
        e.preventDefault()
        if (disabled) {
          e.dataTransfer.dropEffect = 'none'
          return
        }
        e.dataTransfer.dropEffect = 'copy'
        depth.current += 1
        setOver(true)
      }}
      onDragOver={(e) => {
        if (!hasFiles(e)) return
        e.preventDefault()
        if (disabled) {
          e.dataTransfer.dropEffect = 'none'
          return
        }
        e.dataTransfer.dropEffect = 'copy'
      }}
      onDragLeave={() => {
        depth.current = Math.max(0, depth.current - 1)
        if (depth.current === 0) setOver(false)
      }}
      onDrop={(e) => {
        if (!hasFiles(e)) return
        e.preventDefault()
        reset()
        if (disabled) return
        const files = [...e.dataTransfer.files]
        if (files.length) onFiles(files)
      }}
    >
      {children}
      {over && (
        <div className="drop-overlay">
          <div className="drop-overlay-msg">Släpp filerna här</div>
        </div>
      )}
    </div>
  )
}
