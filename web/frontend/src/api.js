// Tunt klientlager mot FastAPI. Relativa sokvagar fungerar bade bakom
// Vite-proxyn (npm run dev) och nar FastAPI serverar den byggda frontenden.

async function parseError(res) {
  let detail
  try {
    detail = (await res.json()).detail
  } catch {
    detail = res.statusText
  }
  if (detail && typeof detail === 'object') return detail.message || JSON.stringify(detail)
  return detail || `Fel ${res.status}`
}

async function jsonOrThrow(res) {
  if (!res.ok) throw new Error(await parseError(res))
  return res.json()
}

export async function health() {
  return jsonOrThrow(await fetch('/api/health'))
}

export async function getFlows() {
  const data = await jsonOrThrow(await fetch('/api/flows'))
  return data.flows
}

export async function detect(file) {
  const fd = new FormData()
  fd.append('file', file)
  return jsonOrThrow(await fetch('/api/detect', { method: 'POST', body: fd }))
}

// formData: FormData med filer (UploadFile) och textfalt.
export async function runFlow(flowId, formData) {
  return jsonOrThrow(
    await fetch(`/api/flow/${flowId}`, { method: 'POST', body: formData }),
  )
}

export async function openExcel(sessionId, key) {
  return jsonOrThrow(
    await fetch('/api/open-excel', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ session_id: sessionId, key }),
    }),
  )
}

export function downloadUrl(sessionId, key) {
  return `/api/download/${sessionId}/${key}`
}
