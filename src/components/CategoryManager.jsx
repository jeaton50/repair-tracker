import React, { useState } from "react";

/**
 * Category to PM Mapping Manager Modal
 * Allows users to create, view, and delete category-to-PM mappings
 */
const CategoryManager = ({
  categoryMapping,
  allCategories,
  unmatchedCategories,
  onAddMapping,
  onRemoveMapping,
  onExport,
  onClose,
}) => {
  const [newCategory, setNewCategory] = useState("");
  const [newPM, setNewPM] = useState("");
  const [newDepartment, setNewDepartment] = useState("");
  const [newCategoryText, setNewCategoryText] = useState("");

  const handleAddMapping = () => {
    if (newCategory && newPM) {
      onAddMapping(newCategory, newPM, newDepartment, newCategoryText);
      setNewCategory("");
      setNewPM("");
      setNewDepartment("");
      setNewCategoryText("");
    }
  };

  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-lg shadow-xl max-w-6xl w-full max-h-[90vh] overflow-hidden flex flex-col">
        {/* Header */}
        <div className="p-6 border-b flex justify-between items-center">
          <h2 className="text-2xl font-bold text-gray-800">
            Category to PM Mapping Manager
          </h2>
          <button
            onClick={onClose}
            className="text-gray-500 hover:text-gray-700 text-2xl"
          >
            ×
          </button>
        </div>

        {/* Body */}
        <div className="p-6 space-y-6 overflow-y-auto flex-1">
          {/* Add New Mapping Form */}
          <div className="bg-blue-50 p-4 rounded-lg">
            <h3 className="font-semibold text-blue-900 mb-3">Add New Mapping</h3>
            <div className="grid grid-cols-2 gap-3 mb-3">
              <div>
                <label className="text-xs text-gray-600 mb-1 block">
                  Category Code
                </label>
                <select
                  value={newCategory}
                  onChange={(e) => setNewCategory(e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg"
                >
                  <option value="">Select Category</option>
                  {allCategories.map((cat) => (
                    <option key={cat} value={cat}>
                      {cat}
                    </option>
                  ))}
                </select>
              </div>
              <div>
                <label className="text-xs text-gray-600 mb-1 block">PM Name</label>
                <input
                  type="text"
                  placeholder="PM Name"
                  value={newPM}
                  onChange={(e) => setNewPM(e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg"
                />
              </div>
              <div>
                <label className="text-xs text-gray-600 mb-1 block">
                  Department
                </label>
                <input
                  type="text"
                  placeholder="Department (optional)"
                  value={newDepartment}
                  onChange={(e) => setNewDepartment(e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg"
                />
              </div>
              <div>
                <label className="text-xs text-gray-600 mb-1 block">
                  Category Description
                </label>
                <input
                  type="text"
                  placeholder="Category description (optional)"
                  value={newCategoryText}
                  onChange={(e) => setNewCategoryText(e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg"
                />
              </div>
            </div>
            <button
              onClick={handleAddMapping}
              disabled={!newCategory || !newPM}
              className="w-full px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 disabled:opacity-50"
            >
              Add Mapping
            </button>
          </div>

          {/* Unmatched Categories Alert */}
          {unmatchedCategories.length > 0 && (
            <div className="bg-red-50 p-4 rounded-lg">
              <h3 className="font-semibold text-red-900 mb-2">
                Unmatched Categories ({unmatchedCategories.length})
              </h3>
              <p className="text-sm text-red-800 mb-3">
                These categories don't have PM assignments:
              </p>
              <div className="flex flex-wrap gap-2">
                {unmatchedCategories.map((cat) => (
                  <span
                    key={cat}
                    className="px-3 py-1 bg-red-100 text-red-800 rounded-full text-sm"
                  >
                    {cat}
                  </span>
                ))}
              </div>
            </div>
          )}

          {/* Current Mappings List */}
          <div>
            <div className="flex justify-between items-center mb-3">
              <h3 className="font-semibold text-gray-800">
                Current Mappings ({categoryMapping.length})
              </h3>
              <button
                onClick={onExport}
                className="text-sm px-3 py-1 bg-gray-600 text-white rounded hover:bg-gray-700"
              >
                Export JSON
              </button>
            </div>
            <div className="space-y-2 max-h-96 overflow-y-auto">
              {categoryMapping.map((m, idx) => (
                <div
                  key={idx}
                  className="flex items-start justify-between p-4 bg-gray-50 rounded-lg border"
                >
                  <div className="flex-1 space-y-1">
                    <div className="flex items-center gap-2">
                      <span className="font-bold text-gray-900">{m.category}</span>
                      <span className="text-gray-400">→</span>
                      <span className="font-semibold text-blue-600">{m.pm}</span>
                    </div>
                    {m.category_text && (
                      <p className="text-sm text-gray-600">{m.category_text}</p>
                    )}
                    {m.department && (
                      <p className="text-xs text-gray-500">
                        Department: {m.department}
                      </p>
                    )}
                  </div>
                  <button
                    onClick={() => onRemoveMapping(m.category)}
                    className="text-red-600 hover:text-red-800 text-sm ml-4"
                  >
                    Remove
                  </button>
                </div>
              ))}
            </div>
          </div>

          {/* All Categories Overview */}
          <div>
            <h3 className="font-semibold text-gray-800 mb-3">
              All Categories in Data ({allCategories.length})
            </h3>
            <div className="flex flex-wrap gap-2">
              {allCategories.map((cat) => {
                const hasMapping = categoryMapping.some(
                  (m) =>
                    m.category.trim().toUpperCase() === cat.trim().toUpperCase()
                );
                return (
                  <span
                    key={cat}
                    className={`px-3 py-1 rounded-full text-sm ${
                      hasMapping
                        ? "bg-green-100 text-green-800"
                        : "bg-gray-100 text-gray-800"
                    }`}
                  >
                    {cat} {hasMapping && "✓"}
                  </span>
                );
              })}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};

export default CategoryManager;
