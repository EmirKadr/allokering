// Delad logik for datapoolens filrutor - anvands av UploadView och CombinedView.
// "details" och "orders" ar samma filformat (bestallningslinjer) -> en gemensam ruta.
export const logicalKey = (key) => (key === 'details' ? 'orders' : key)

const SLOT_LABELS = {
  orders: 'Bestallningslinjer',
  buffer: 'Buffertpallar',
  overview: 'Orderoversikt',
  dispatch: 'Dispatchpallar',
  saldo: 'Saldo / automation',
  items: 'Item option',
  not_putaway: 'Ej inlagrade',
  prognos: 'Prognosfil',
  campaign: 'Kampanjfil',
  max_csv: 'artikel_max.csv',
}

const SLOT_ORDER = [
  'orders', 'buffer', 'overview', 'dispatch', 'saldo',
  'items', 'not_putaway', 'prognos', 'campaign', 'max_csv',
]

export function slotLabel(key) {
  const lk = logicalKey(key)
  return SLOT_LABELS[lk] || lk
}

// Union av alla filrutor over de givna floden, i fast ordning.
export function deriveSlots(flows) {
  const map = new Map()
  for (const flow of flows) {
    for (const inp of flow.inputs || []) {
      if (inp.type && inp.type !== 'file') continue
      const lk = logicalKey(inp.key)
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
