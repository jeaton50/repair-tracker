import React from "react";
import { ChevronUp, ChevronDown, ChevronLeft, ChevronRight } from "lucide-react";
import EditableCell from "./EditableCell";

/**
 * Paginated table component with inline editing support
 * Displays repair data with sorting, pagination, and quick edit functionality
 */
const PaginatedTable = ({
  data,
  columns,
  onRowClick,
  activeTab,
  currentPage,
  setCurrentPage,
  itemsPerPage,
  sortConfig,
  onSort,
  // Inline editing props
  notesService,
  onInlineNoteChange,
  onInlineFollowUpChange,
  onInlineSaveNow,
}) => {
  const totalPages = Math.ceil(data.length / itemsPerPage);
  const startIdx = (currentPage - 1) * itemsPerPage;
  const endIdx = startIdx + itemsPerPage;
  const paginatedData = data.slice(startIdx, endIdx);

  const isInlineCol = (col) =>
    activeTab === "combined" &&
    (col === "Meeting Note" || col === "Requires Follow Up");

  return (
    <div className="h-full flex flex-col bg-white rounded-lg shadow">
      {/* Scrollable table wrapper */}
      <div className="flex-1 overflow-auto">
        <table className="w-full border-collapse col-18ch-table">
          {/* Column sizing */}
          <colgroup>
            {columns.map((c) => (
              <col
                key={c}
                style={
                  c === "Requires Follow Up"
                    ? {
                        width: "calc(12ch + 64px + 2rem)",
                        minWidth: "calc(12ch + 64px + 2rem)",
                      }
                    : undefined
                }
              />
            ))}
          </colgroup>

          {/* Header */}
          <thead className="bg-gray-50 border-b sticky top-0 z-10">
            <tr>
              {columns.map((col) => {
                const thExtra =
                  col === "Meeting Note"
                    ? "note-col"
                    : col === "Requires Follow Up"
                    ? "followup-col"
                    : "";
                return (
                  <th
                    key={col}
                    onClick={() => onSort(col)}
                    className={`px-4 py-3 text-left text-xs font-medium text-gray-700 uppercase tracking-wider whitespace-normal bg-gray-50 cursor-pointer hover:bg-gray-100 ${thExtra}`}
                    onMouseDown={(e) => e.stopPropagation()}
                  >
                    <div className="flex items-center gap-2">
                      {col}
                      {isInlineCol(col) && <span className="text-blue-500">✏️</span>}
                      {sortConfig.key === col ? (
                        sortConfig.direction === "asc" ? (
                          <ChevronUp size={14} />
                        ) : (
                          <ChevronDown size={14} />
                        )
                      ) : null}
                    </div>
                  </th>
                );
              })}
            </tr>
          </thead>

          {/* Body */}
          <tbody className="bg-white divide-y">
            {paginatedData.map((row, idx) => {
              const hasAssignment = row["Assigned To"] && row["Assigned To"] !== "";
              const rowBg =
                activeTab === "combined" && !hasAssignment ? "bg-red-50" : "";
              const actualIndex = startIdx + idx;

              return (
                <tr
                  key={actualIndex}
                  className={`${rowBg} hover:bg-gray-50 cursor-pointer`}
                  onClick={() => onRowClick(actualIndex)}
                >
                  {columns.map((col) => {
                    // Inline editors for note columns in Combined view
                    if (isInlineCol(col)) {
                      const barcode = row["Barcode#"] || row["Barcode"];
                      const noteObj = notesService?.getNote(barcode) || {
                        meetingNote: "",
                        requiresFollowUp: "",
                      };
                      const value =
                        col === "Meeting Note"
                          ? noteObj.meetingNote
                          : noteObj.requiresFollowUp;
                      const handleChange =
                        col === "Meeting Note"
                          ? (v) => onInlineNoteChange(barcode, v)
                          : (v) => onInlineFollowUpChange(barcode, v);

                      const tdExtra =
                        col === "Meeting Note" ? "note-col" : "followup-col";

                      return (
                        <td
                          key={col}
                          className={`px-4 py-3 text-sm text-gray-900 whitespace-normal break-words ${tdExtra}`}
                          onMouseDown={(e) => e.stopPropagation()}
                          onClick={(e) => e.stopPropagation()}
                        >
                          <EditableCell
                            value={value}
                            onChange={handleChange}
                            onSave={onInlineSaveNow}
                            multiline={col === "Meeting Note"}
                            placeholder={
                              col === "Meeting Note"
                                ? "Type meeting note…"
                                : "Follow up…"
                            }
                            inputWidth={
                              col === "Requires Follow Up" ? "w-followup" : "w-full"
                            }
                            saveBelow={col === "Requires Follow Up"}
                          />
                        </td>
                      );
                    }

                    // Regular cells
                    const content = String(row[col] ?? "");
                    return (
                      <td
                        key={col}
                        className="px-4 py-3 text-sm text-gray-900 whitespace-normal break-words"
                        style={{ maxWidth: 300 }}
                      >
                        {content}
                      </td>
                    );
                  })}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* Pagination footer */}
      {totalPages > 1 && itemsPerPage < 99999 && (
        <div className="border-t bg-white px-6 py-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-2">
              <button
                onClick={() => setCurrentPage(1)}
                disabled={currentPage === 1}
                className="px-3 py-1 text-sm border border-gray-300 rounded hover:bg-gray-50 disabled:opacity-50 disabled:cursor-not-allowed"
              >
                First
              </button>
              <button
                onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                disabled={currentPage === 1}
                className="flex items-center gap-1 px-3 py-1 text-sm border border-gray-300 rounded hover:bg-gray-50 disabled:opacity-50 disabled:cursor-not-allowed"
              >
                <ChevronLeft size={16} />
                Previous
              </button>
            </div>

            <div className="flex items-center gap-2">
              <span className="text-sm text-gray-600">
                Page {currentPage} of {totalPages}
              </span>
              <span className="text-sm text-gray-400">|</span>
              <span className="text-sm text-gray-600">
                Showing {startIdx + 1}-{Math.min(endIdx, data.length)} of{" "}
                {data.length}
              </span>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                disabled={currentPage === totalPages}
                className="flex items-center gap-1 px-3 py-1 text-sm border border-gray-300 rounded hover:bg-gray-50 disabled:opacity-50 disabled:cursor-not-allowed"
              >
                Next
                <ChevronRight size={16} />
              </button>
              <button
                onClick={() => setCurrentPage(totalPages)}
                disabled={currentPage === totalPages}
                className="px-3 py-1 text-sm border border-gray-300 rounded hover:bg-gray-50 disabled:opacity-50 disabled:cursor-not-allowed"
              >
                Last
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default PaginatedTable;
