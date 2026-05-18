import React from 'react'

// Navigation. groups = [{ name, items: [{ id, label, description }] }].
export default function Sidebar({ groups, activeId, onSelect }) {
  return (
    <nav className="sidebar">
      {groups.map((group) => (
        <div key={group.name} className="nav-group">
          <div className="nav-group-title">{group.name}</div>
          {group.items.map((item) => (
            <button
              key={item.id}
              className={`nav-item ${item.id === activeId ? 'active' : ''}`}
              onClick={() => onSelect(item.id)}
              title={item.description || ''}
            >
              {item.label}
            </button>
          ))}
        </div>
      ))}
    </nav>
  )
}
