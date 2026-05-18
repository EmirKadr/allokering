import React, { useRef } from 'react'

export default function FilePickerButton({
  onFiles,
  disabled = false,
  children = 'Välj filer',
  multiple = true,
}) {
  const inputRef = useRef(null)

  return (
    <>
      <input
        ref={inputRef}
        type="file"
        multiple={multiple}
        hidden
        disabled={disabled}
        accept=".csv,.xlsx,.xlsm,.xls,.txt"
        onChange={(e) => {
          const files = [...e.target.files]
          if (files.length) onFiles(files)
          e.target.value = ''
        }}
      />
      <button
        type="button"
        className="btn-sm"
        disabled={disabled}
        onClick={() => inputRef.current?.click()}
      >
        {children}
      </button>
    </>
  )
}
