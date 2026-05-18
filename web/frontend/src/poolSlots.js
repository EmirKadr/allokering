// Delad logik för datapoolens filrader.
// Flera flödes-inputs kan dela samma fysiska fil i poolen.
const KEY_OVERRIDES = {
  details: 'orders',
  wms_buffert: 'buffer',
}

export const logicalKey = (key) => KEY_OVERRIDES[key] || key

const SLOT_LABELS = {
  orders: 'Beställningslinjer',
  buffer: 'Buffertpallar',
  overview: 'Orderöversikt',
  dispatch: 'Dispatchpallar',
  saldo: 'Saldo / automation',
  items: 'Item option',
  not_putaway: 'Ej inlagrade',
  prognos: 'Prognosfil',
  campaign: 'Kampanjfil',
  max_csv: 'artikel_max.csv',
  wms_receive: 'Mottagningslogg',
  wms_booking: 'Inlagringslogg',
  wms_trans: 'Transaktionslogg',
  wms_pick: 'Plocklogg',
  wms_correct: 'Korrigeringslogg',
  remote_file: 'Observationsfil',
  values_file: 'Textfil med värden',
}

const SLOT_ORDER = [
  'orders',
  'buffer',
  'overview',
  'dispatch',
  'saldo',
  'items',
  'not_putaway',
  'prognos',
  'campaign',
  'max_csv',
  'wms_receive',
  'wms_booking',
  'wms_trans',
  'wms_pick',
  'wms_correct',
  'remote_file',
  'values_file',
]

export function slotLabel(key) {
  const lk = logicalKey(key)
  return SLOT_LABELS[lk] || lk
}

export function fileInputKey(input) {
  return logicalKey(input.pool || input.key)
}

// Union av alla filrader över de givna flödena, i fast ordning.
export function deriveSlots(flows) {
  const map = new Map()
  for (const flow of flows) {
    for (const inp of flow.inputs || []) {
      if (inp.type && inp.type !== 'file') continue
      const lk = fileInputKey(inp)
      if (!map.has(lk)) {
        map.set(lk, {
          key: lk,
          label: SLOT_LABELS[lk] || inp.label,
          detect: new Set(inp.detect || []),
        })
      } else {
        ;(inp.detect || []).forEach((d) => map.get(lk).detect.add(d))
      }
    }
  }
  const keys = SLOT_ORDER.filter((k) => map.has(k)).concat(
    [...map.keys()].filter((k) => !SLOT_ORDER.includes(k)),
  )
  return keys.map((k) => ({ ...map.get(k), detect: [...map.get(k).detect] }))
}
