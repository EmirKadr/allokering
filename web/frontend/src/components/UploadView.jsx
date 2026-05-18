import React, { useState } from 'react'
import DropZone from './DropZone.jsx'
import FileSlot from './FileSlot.jsx'
import { detect } from '../api.js'
import { deriveSlots } from '../poolSlots.js'

// Egen sida for all datauppladdning. Filerna lagras i App och delas av
// alla korningar i "Allokering & analys".
export default function UploadView({ flows, files, onSet, onClear, onError }) {
  const [status, setStatus] = useState('')
  const slots = deriveSlots(flows)

  const routeDropped = async (dropped) => {
    setStatus('Identifierar filer...')
    const unknown = []
    for (const file of dropped) {
      try {
        const { file_type } = await detect(file)
        const target = slots.find((s) => s.detect.includes(file_type))
        if (target) onSet(target.key, file)
        else unknown.push(file.name)
      } catch {
        unknown.push(file.name)
      }
    }
    setStatus('')
    if (unknown.length) {
      onError(
        'Okand filtyp',
        `Kunde inte sortera automatiskt: ${unknown.join(', ')}. Dra till ratt ruta eller anvand "Valj".`,
        'warn',
      )
    }
  }

  const filledCount = slots.filter((s) => files[s.key]).length

  return (
    <div className="flow-view">
      <div className="flow-header">
        <h1>Datauppladdning</h1>
        <p className="flow-desc">
          Ladda upp alla filer har - en gang. De delas av alla korningar under "Allokering &amp;
          analys". Filtypen kanns igen automatiskt nar du slapper en fil.
        </p>
      </div>

      <section className="panel">
        <h2 className="panel-title">
          Filer · {filledCount}/{slots.length} inlagda
        </h2>
        <DropZone onFiles={routeDropped} busy={false} />
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
  )
}
