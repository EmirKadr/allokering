// Tunt klientlager mot FastAPI. Relativa sokvagar fungerar bade bakom
// Vite-proxyn (npm run dev) och nar FastAPI serverar den byggda frontenden.

async function parseError(res) {
  let detail
  try {
    const data = await res.json()
    detail = data.detail
  } catch {
    detail = res.statusText
  }
  if (detail && typeof detail === 'object') return detail.message || JSON.stringify(detail)
  return detail || `Fel ${res.status}`
}

export async function health() {
  const res = await fetch('/api/health')
  if (!res.ok) throw new Error(await parseError(res))
  return res.json()
}

export async function detect(file) {
  const fd = new FormData()
  fd.append('file', file)
  const res = await fetch('/api/detect', { method: 'POST', body: fd })
  if (!res.ok) throw new Error(await parseError(res))
  return res.json()
}

export async function allocate(slots) {
  const fd = new FormData()
  for (const [key, entry] of Object.entries(slots)) {
    if (entry && entry.file) fd.append(key, entry.file, entry.file.name)
  }
  const res = await fetch('/api/allocate', { method: 'POST', body: fd })
  if (!res.ok) throw new Error(await parseError(res))
  return res.json()
}

export async function openExcel(sessionId, key) {
  const res = await fetch('/api/open-excel', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ session_id: sessionId, key }),
  })
  if (!res.ok) throw new Error(await parseError(res))
  return res.json()
}

export function downloadUrl(sessionId, key) {
  return `/api/download/${sessionId}/${key}`
}
