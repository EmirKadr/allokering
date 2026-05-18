import { detect } from './api.js'

const FILE_WORDS = {
  max_csv: ['artikel_max', 'article_max'],
  not_putaway: ['not_putaway', 'not putaway', 'ej_inlag', 'ej inlag', 'ejinlag'],
  remote_file: ['observations', 'observationer'],
  values_file: ['values', 'varden', 'värden'],
}

const fallbackSlot = (file, slots, droppedCount) => {
  const name = (file.name || '').toLowerCase()
  const hinted = slots.find((slot) =>
    (FILE_WORDS[slot.key] || []).some((word) => name.includes(word)),
  )
  if (hinted) return hinted

  return droppedCount === 1 && slots.length === 1 ? slots[0] : null
}

export async function routeFilesToSlots(droppedFiles, slots, onSet, options = {}) {
  const files = [...(droppedFiles || [])]
  const targetSlots = [...(slots || [])]
  if (!files.length) return { assigned: [], unknown: [] }

  const setStatus = options.setStatus || (() => {})
  const onUnknown = options.onUnknown || (() => {})
  const onNoSlots = options.onNoSlots || (() => {})

  if (!targetSlots.length) {
    onNoSlots(files)
    return { assigned: [], unknown: files.map((file) => file.name) }
  }

  setStatus('Identifierar filer...')
  const assigned = []
  const unknown = []

  for (const file of files) {
    let target = null
    try {
      const { file_type } = await detect(file)
      target = targetSlots.find((slot) => (slot.detect || []).includes(file_type))
    } catch {
      target = null
    }

    if (!target) target = fallbackSlot(file, targetSlots, files.length)

    if (target) {
      onSet(target.key, file)
      assigned.push({ file, slot: target })
    } else {
      unknown.push(file.name)
    }
  }

  if (assigned.length === 1) setStatus('1 fil inlagd.')
  else if (assigned.length > 1) setStatus(`${assigned.length} filer inlagda.`)
  else setStatus('')

  if (unknown.length) onUnknown(unknown)
  return { assigned, unknown }
}
