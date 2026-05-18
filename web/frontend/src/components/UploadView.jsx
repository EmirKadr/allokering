import React, { useState } from 'react'
import DropArea from './DropArea.jsx'
import FileSlot from './FileSlot.jsx'
import FilePickerButton from './FilePickerButton.jsx'
import { routeFilesToSlots } from '../fileRouting.js'

// Egen sida för all datauppladdning. Filerna lagras i App och delas av
// alla körningar.
export default function UploadView({ slots, files, onSet, onClear, onError, autoStatus }) {
  const [status, setStatus] = useState('')

  const routeDropped = (dropped) =>
    routeFilesToSlots(dropped, slots, onSet, {
      setStatus,
      onUnknown: (unknown) =>
        onError(
          'Okänd filtyp',
          `Kunde inte sortera automatiskt: ${unknown.join(', ')}. Använd "Välj" på rätt filrad om filen saknar igenkännbar typ.`,
          'warn',
        ),
    })

  const filledCount = slots.filter((s) => files[s.key]).length

  return (
    <DropArea onFiles={routeDropped}>
      <div className="flow-view">
        <div className="flow-header">
          <h1>Datauppladdning</h1>
          <p className="flow-desc">
            Ladda upp alla filer här - en gång. De delas av allokering, analys, eftersök och
            verktyg. Släpp filer var som helst i vyn så sorteras igenkända filtyper automatiskt.
          </p>
        </div>

        <section className="panel">
          <div className="panel-head">
            <h2 className="panel-title">Filer · {filledCount}/{slots.length} inlagda</h2>
            <FilePickerButton onFiles={routeDropped}>Välj filer</FilePickerButton>
          </div>
          <p className="drop-hint">
            Släpp filer var som helst i vyn. Filraderna nedan visar vad som är uppladdat.
          </p>
          {autoStatus && <p className="status-text upload-status">{autoStatus}</p>}
          <div className="pool-grid">
            {slots.map((slot) => (
              <FileSlot
                key={slot.key}
                slot={slot}
                entry={files[slot.key] || null}
                onSet={onSet}
                onClear={onClear}
              />
            ))}
          </div>
          {status && <p className="status-text upload-status">{status}</p>}
        </section>
      </div>
    </DropArea>
  )
}
