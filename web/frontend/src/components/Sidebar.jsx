import React from 'react'

// Flodesnavigation grupperad per kategori.
export default function Sidebar({ flows, activeId, onSelect }) {
  const categories = []
  for (const flow of flows) {
    let cat = categories.find((c) => c.name === flow.category)
    if (!cat) {
      cat = { name: flow.category, flows: [] }
      categories.push(cat)
    }
    cat.flows.push(flow)
  }

  return (
    <nav className="sidebar">
      {categories.map((cat) => (
        <div key={cat.name} className="nav-group">
          <div className="nav-group-title">{cat.name}</div>
          {cat.flows.map((flow) => (
            <button
              key={flow.id}
              className={`nav-item ${flow.id === activeId ? 'active' : ''}`}
              onClick={() => onSelect(flow.id)}
              title={flow.description}
            >
              {flow.label}
            </button>
          ))}
        </div>
      ))}
    </nav>
  )
}
