import React, { useState, useEffect } from "react";
import toast from "react-hot-toast";

/**
 * Modal editor for detailed repair item editing
 * Supports editing meeting notes and follow-up status
 */
const RowEditor = ({ row, onClose, notesService, onSave }) => {
  const barcode = row?.["Barcode#"] || "";
  const [meetingNote, setMeetingNote] = useState("");
  const [followUp, setFollowUp] = useState("");
  const [isSaving, setIsSaving] = useState(false);
  const [isLoading, setIsLoading] = useState(true);
  const [hasChanges, setHasChanges] = useState(false);
  const [lastSaved, setLastSaved] = useState(null);

  // Load current note from service
  useEffect(() => {
    if (!barcode || !notesService) {
      setIsLoading(false);
      return;
    }
    const note = notesService.getNote(barcode);
    setMeetingNote(note.meetingNote || "");
    setFollowUp(note.requiresFollowUp || "");
    setIsLoading(false);
  }, [barcode, notesService]);

  // Track changes
  useEffect(() => {
    if (isLoading) return;
    const original = notesService?.getNote(barcode) || {
      meetingNote: "",
      requiresFollowUp: ""
    };
    setHasChanges(
      meetingNote !== (original.meetingNote || "") ||
      followUp !== (original.requiresFollowUp || "")
    );
  }, [meetingNote, followUp, barcode, notesService, isLoading]);

  const handleSave = async () => {
    if (!barcode || !notesService) return;
    setIsSaving(true);

    const savePromise = (async () => {
      notesService.updateNote(barcode, meetingNote, followUp);
      await notesService.saveToSharePoint();
      setLastSaved(new Date());
      setHasChanges(false);
      onSave?.();
    })();

    toast.promise(savePromise, {
      loading: "Saving to SharePoint...",
      success: "Saved successfully!",
      error: "Failed to save. Please try again.",
    });

    try {
      await savePromise;
    } catch (e) {
      console.error("Save error:", e);
    } finally {
      setIsSaving(false);
    }
  };

  const handleClose = () => {
    if (hasChanges && !window.confirm("You have unsaved changes. Close anyway?")) {
      return;
    }
    onClose?.();
  };

  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-lg shadow-xl w-full max-w-4xl max-h-[90vh] overflow-hidden flex flex-col">
        {/* Header */}
        <div className="p-6 border-b flex justify-between items-center">
          <div>
            <h2 className="text-xl font-bold text-gray-800">Edit Repair Item</h2>
            <p className="text-sm text-gray-500 mt-1">
              {barcode} {row?.["Equipment"] ? `- ${row["Equipment"]}` : ""}
            </p>
            {isLoading && <p className="text-xs text-blue-600 mt-1">⟳ Loading…</p>}
            {isSaving && <p className="text-xs text-blue-600 mt-1">⟳ Saving to SharePoint…</p>}
            {hasChanges && !isSaving && (
              <p className="text-xs text-orange-600 mt-1">⚠️ Unsaved changes</p>
            )}
            {lastSaved && !isSaving && !hasChanges && (
              <p className="text-xs text-green-600 mt-1">
                ✓ Saved at {lastSaved.toLocaleTimeString()}
              </p>
            )}
          </div>
          <button
            onClick={handleClose}
            className="text-gray-500 hover:text-gray-700 text-2xl"
          >
            ×
          </button>
        </div>

        {/* Body */}
        <div className="p-6 space-y-6 overflow-y-auto flex-1">
          {/* Meeting Note */}
          <div>
            <div className="flex items-center justify-between mb-2">
              <label className="block text-sm font-semibold text-gray-700">
                Meeting Note
              </label>
              <button
                type="button"
                onClick={() => setMeetingNote("")}
                className="text-xs px-2 py-1 text-red-600 border border-red-300 rounded hover:bg-red-50"
              >
                Clear
              </button>
            </div>
            <textarea
              value={meetingNote}
              onChange={(e) => setMeetingNote(e.target.value)}
              placeholder="Add meeting notes here…"
              className="w-full h-48 px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent resize-none"
              disabled={isLoading || isSaving}
            />
          </div>

          {/* Follow Up */}
          <div>
            <div className="flex items-center justify-between mb-2">
              <label className="block text-sm font-semibold text-gray-700">
                Requires Follow Up
              </label>
              <div className="flex gap-2">
                <button
                  type="button"
                  onClick={() => setFollowUp(meetingNote)}
                  className="text-xs px-2 py-1 border border-gray-300 rounded hover:bg-gray-50"
                >
                  Copy from Meeting Note
                </button>
                <button
                  type="button"
                  onClick={() => setFollowUp("")}
                  className="text-xs px-2 py-1 text-red-600 border border-red-300 rounded hover:bg-red-50"
                >
                  Clear
                </button>
              </div>
            </div>
            <textarea
              value={followUp}
              onChange={(e) => setFollowUp(e.target.value)}
              placeholder="Add follow-up notes here…"
              className="w-full h-32 px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent resize-none"
              disabled={isLoading || isSaving}
            />
          </div>

          {/* Context Information */}
          <div className="p-4 bg-gray-50 rounded-lg text-sm space-y-2">
            <div>
              <span className="font-semibold text-gray-700">Damage:</span>{" "}
              {row?.["Damage Description"] ?? ""}
            </div>
            <div>
              <span className="font-semibold text-gray-700">Ticket Notes:</span>{" "}
              {row?.["Ticket Description"] ?? ""}
            </div>
            <div>
              <span className="font-semibold text-gray-700">Reason:</span>{" "}
              {row?.["Repair Reason"] ?? ""}
            </div>
          </div>
        </div>

        {/* Footer */}
        <div className="p-6 border-t flex justify-between">
          <div />
          <div className="flex gap-2">
            <button
              onClick={handleSave}
              disabled={!hasChanges || isSaving || isLoading}
              className="flex items-center gap-2 px-6 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 disabled:opacity-50 disabled:cursor-not-allowed"
            >
              Save to SharePoint
            </button>
            <button
              onClick={handleClose}
              className="px-6 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700"
            >
              Close
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

export default RowEditor;
