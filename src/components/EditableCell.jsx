import React from "react";

/**
 * Inline editable cell component for Quick Edit functionality
 * Supports both single-line and multiline editing with save button
 */
const EditableCell = ({
  value,
  onChange,
  onSave,
  multiline = false,
  placeholder = "",
  inputWidth = "w-full",
  saveBelow = false,
}) => {
  const onKeyDown = (e) => {
    // Ctrl/Cmd + S to save
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") {
      e.preventDefault();
      onSave?.();
    }
    // Enter to save (single-line only)
    if (!multiline && e.key === "Enter") {
      e.preventDefault();
      onSave?.();
    }
    e.stopPropagation();
  };

  const commonProps = {
    value,
    placeholder,
    onChange: (e) => onChange(e.target.value),
    onKeyDown,
    onClick: (e) => e.stopPropagation(),
    onMouseDown: (e) => e.stopPropagation(),
    className: `${inputWidth} text-sm border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-transparent`,
  };

  const saveButton = (
    <button
      type="button"
      className="px-3 py-2 bg-green-600 text-white rounded-md hover:bg-green-700"
      onClick={(e) => {
        e.stopPropagation();
        onSave?.();
      }}
      title="Save now (Ctrl+S)"
    >
      Save
    </button>
  );

  // Multiline (textarea)
  if (multiline) {
    return (
      <div className="grid grid-cols-1 gap-2">
        <textarea
          {...commonProps}
          rows={6}
          style={{ minHeight: "7rem" }}
          className={`${inputWidth} note-input text-sm border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-transparent`}
        />
        <div>{saveButton}</div>
      </div>
    );
  }

  // Single-line with save button below
  if (saveBelow) {
    return (
      <div className="grid grid-cols-1 gap-2">
        <input {...commonProps} />
        <div>{saveButton}</div>
      </div>
    );
  }

  // Single-line with save button beside
  return (
    <div className="flex items-center gap-2">
      <input {...commonProps} />
      {saveButton}
    </div>
  );
};

export default React.memo(EditableCell);
