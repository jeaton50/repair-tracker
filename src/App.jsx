// src/App.jsx - SHAREPOINT VERSION (No Firebase3426) + Quick Edit wiring1
import React, { useState, useMemo, useEffect, useRef, useCallback } from "react";
import { Client } from "@microsoft/microsoft-graph-client";
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig, loginRequest, graphConfig } from "./authConfig";
import OneDriveService from "./oneDriveService.js";
import SharePointNotesService from "./SharePointNotesService.js";
import * as XLSX from "xlsx";
import {
  Search,
  Download,
  ChevronDown,
  ChevronUp,
  Upload,
  FileSpreadsheet,
  RefreshCw,
  Cloud,
  ChevronLeft,
  ChevronRight,
  Save,
} from "lucide-react";

// ---------- MSAL instance & helpers ----------
const msalInstance = new PublicClientApplication(msalConfig);

async function ensureAccessToken() {
  let account = msalInstance.getAllAccounts()[0];
  if (!account) {
    await msalInstance.loginPopup(loginRequest);
    account = msalInstance.getAllAccounts()[0];
  }
  try {
    const r = await msalInstance.acquireTokenSilent({ ...loginRequest, account });
    return r.accessToken;
  } catch {
    const r = await msalInstance.acquireTokenPopup({ ...loginRequest, account });
    return r.accessToken;
  }
}

// ---------- Custom Hook: Debounce ----------
const useDebounce = (callback, delay) => {
  const timeoutRef = useRef(null);
  return useCallback(
    (...args) => {
      if (timeoutRef.current) clearTimeout(timeoutRef.current);
      timeoutRef.current = setTimeout(() => callback(...args), delay);
    },
    [callback, delay]
  );
};

/* ---------- Tiny inline editor for Quick Edit ---------- */
// ---------- Inline editable cell ----------
// ---------- EditableCell (merged) ----------
const EditableCell = ({
  value,
  onChange,
  onSave,
  multiline = false,
  placeholder = "",
  inputWidth = "w-full", // e.g. "w-20ch" to force ~20 characters width
}) => {
  // Save on Ctrl/Cmd+S, stop row clicks while editing
  const onKeyDown = (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "s") {
      e.preventDefault();
      onSave?.();
    }
    // For single-line, Enter saves
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

  if (multiline) {
    return (
      <div className="flex items-start gap-2">
        <textarea
          {...commonProps}
          rows={6}
          style={{ minHeight: "7rem" }} // nice multi-line height
		  className={`${inputWidth} note-input text-sm border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-transparent`}
        />
		<input
  {...commonProps}
  className={`${inputWidth} followup-input text-sm border border-gray-300 rounded-md px-3 py-2 focus:ring-2 focus:ring-blue-500 focus:border-transparent`}
/>
        <button
          type="button"
          className="shrink-0 px-3 py-2 bg-green-600 text-white rounded-md hover:bg-green-700"
          onClick={(e) => {
            e.stopPropagation();
            onSave?.();
          }}
          title="Save now"
        >
          Save
        </button>
      </div>
    );
  }

  return (
    <div className="flex items-center gap-2">
      <input {...commonProps} />
      <button
        type="button"
        className="shrink-0 px-3 py-2 bg-green-600 text-white rounded-md hover:bg-green-700"
        onClick={(e) => {
          e.stopPropagation();
          onSave?.();
        }}
        title="Save now"
      >
        Save
      </button>
    </div>
  );
};


// ---------- Row Editor (modal) ----------
const RowEditor = ({ row, onClose, notesService, onSave }) => {
  const barcode = row?.["Barcode#"] || "";
  const [meetingNote, setMeetingNote] = React.useState("");
  const [followUp, setFollowUp] = React.useState("");
  const [isSaving, setIsSaving] = React.useState(false);
  const [isLoading, setIsLoading] = React.useState(true);
  const [hasChanges, setHasChanges] = React.useState(false);
  const [lastSaved, setLastSaved] = React.useState(null);

  // load current note
  React.useEffect(() => {
    if (!barcode || !notesService) {
      setIsLoading(false);
      return;
    }
    const note = notesService.getNote(barcode);
    setMeetingNote(note.meetingNote || "");
    setFollowUp(note.requiresFollowUp || "");
    setIsLoading(false);
  }, [barcode, notesService]);

  // change tracking
  React.useEffect(() => {
    if (isLoading) return;
    const original = notesService?.getNote(barcode) || { meetingNote: "", requiresFollowUp: "" };
    setHasChanges(
      meetingNote !== (original.meetingNote || "") ||
      followUp !== (original.requiresFollowUp || "")
    );
  }, [meetingNote, followUp, barcode, notesService, isLoading]);

  const handleSave = async () => {
    if (!barcode || !notesService) return;
    setIsSaving(true);
    try {
      notesService.updateNote(barcode, meetingNote, followUp);
      await notesService.saveToSharePoint();
      setLastSaved(new Date());
      setHasChanges(false);
      onSave?.();
      alert("✅ Saved to SharePoint successfully!");
    } catch (e) {
      console.error("Save error:", e);
      alert("❌ Failed to save. Please try again.");
    } finally {
      setIsSaving(false);
    }
  };

  const handleClose = () => {
    if (hasChanges && !window.confirm("You have unsaved changes. Close anyway?")) return;
    onClose?.();
  };

  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-lg shadow-xl w-full max-w-4xl max-h-[90vh] overflow-hidden flex flex-col">
        <div className="p-6 border-b flex justify-between items-center">
          <div>
            <h2 className="text-xl font-bold text-gray-800">Edit Repair Item</h2>
            <p className="text-sm text-gray-500 mt-1">
              {barcode} {row?.["Equipment"] ? `- ${row["Equipment"]}` : ""}
            </p>
            {isLoading && <p className="text-xs text-blue-600 mt-1">⟳ Loading…</p>}
            {isSaving && <p className="text-xs text-blue-600 mt-1">⟳ Saving to SharePoint…</p>}
            {hasChanges && !isSaving && <p className="text-xs text-orange-600 mt-1">⚠️ Unsaved changes</p>}
            {lastSaved && !isSaving && !hasChanges && (
              <p className="text-xs text-green-600 mt-1">✓ Saved at {lastSaved.toLocaleTimeString()}</p>
            )}
          </div>
          <button onClick={handleClose} className="text-gray-500 hover:text-gray-700 text-2xl">×</button>
        </div>

        <div className="p-6 space-y-6 overflow-y-auto flex-1">
          <div>
            <div className="flex items-center justify-between mb-2">
              <label className="block text-sm font-semibold text-gray-700">Meeting Note</label>
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

          <div>
            <div className="flex items-center justify-between mb-2">
              <label className="block text-sm font-semibold text-gray-700">Requires Follow Up</label>
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

          <div className="p-4 bg-gray-50 rounded-lg text-sm space-y-2">
            <div><span className="font-semibold text-gray-700">Damage:</span> {row?.["Damage Description"] ?? ""}</div>
            <div><span className="font-semibold text-gray-700">Ticket Notes:</span> {row?.["Ticket Description"] ?? ""}</div>
            <div><span className="font-semibold text-gray-700">Reason:</span> {row?.["Repair Reason"] ?? ""}</div>
          </div>
        </div>

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
            <button onClick={handleClose} className="px-6 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700">
              Close
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};


// ---------- Paginated Table Component ----------
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
  // inline editing wiring
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
      {/* scrollable table wrapper */}
     <div className="flex-1 overflow-auto">
  <table className="w-full border-collapse col-18ch-table">
    {/* NEW: colgroup to pin the Requires Follow Up width (+16px) */}
    <colgroup>
      {columns.map((c) => (
        <col
          key={c}
          // +2rem accounts for px-4 padding on td/th (1rem left + 1rem right)
          style={
            c === "Requires Follow Up"
              ? { width: "calc(12ch + 46px + 2rem)", minWidth: "calc(12ch + 46px + 2rem)" }
              : undefined
          }
        />
      ))}
    </colgroup>

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
                {sortConfig.key === col
                  ? sortConfig.direction === "asc"
                    ? <ChevronUp size={14} />
                    : <ChevronDown size={14} />
                  : null}
              </div>
            </th>
          );
        })}
      </tr>
    </thead>

          <tbody className="bg-white divide-y">
            {paginatedData.map((row, idx) => {
              const hasAssignment = row["Assigned To"] && row["Assigned To"] !== "";
              const rowBg =
                activeTab === "combined" && !hasAssignment ? "bg-red-50" : "";
              const actualIndex = startIdx + idx;

              return (
                <tr
                  key={actualIndex}
                  className={`${rowBg} hover:bg-gray-50`}
                  onClick={() => onRowClick(actualIndex)}
                >
                  {columns.map((col) => {
                    // Inline editors only on Combined for the two note columns
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
                              col === "Meeting Note" ? "Type meeting note…" : "Follow up…"
                            }
                            // follow-up ~12ch wide, meeting note fills the wider cell
                            inputWidth={col === "Requires Follow Up" ? "w-followup" : "w-full"}

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

      {/* pagination footer */}
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
                Showing {startIdx + 1}-{Math.min(endIdx, data.length)} of {data.length}
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




// ---------- Main component ----------
const RepairTrackerSheet = () => {
  const [activeTab, setActiveTab] = useState("combined");
  const [searchInput, setSearchInput] = useState("");
  const [searchTerm, setSearchTerm] = useState("");
  const [sortConfig, setSortConfig] = useState({ key: null, direction: "asc" });

  const [ticketData, setTicketData] = useState([]);
  const [reportData, setReportData] = useState([]);
  const [categoryMapping, setCategoryMapping] = useState([]);

  const [showCategoryManager, setShowCategoryManager] = useState(false);
  const [unmatchedCategories, setUnmatchedCategories] = useState([]);
  const [editingRow, setEditingRow] = useState(null);
  const [editingRowIndex, setEditingRowIndex] = useState(null);

  const [combinedDataWithNotes, setCombinedDataWithNotes] = useState([]);
  const [notesMap, setNotesMap] = useState(new Map());

  const [isImporting, setIsImporting] = useState(false);
  const [notesService, setNotesService] = useState(null);

  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [accessToken, setAccessToken] = useState(null);
  const [userName, setUserName] = useState("");
  const [lastSync, setLastSync] = useState(null);
  const [lastNotesSync, setLastNotesSync] = useState(null);
  const [autoRefresh, setAutoRefresh] = useState(true);
  const [msalInitialized, setMsalInitialized] = useState(false);
  const [loading, setLoading] = useState(false);

  const [locationFilter, setLocationFilter] = useState("");
  const [pmFilter, setPmFilter] = useState("");

  const [currentPage, setCurrentPage] = useState(1);
  const [itemsPerPage, setItemsPerPage] = useState(100);
  const ITEMS_PER_PAGE = itemsPerPage;

  // Quick Edit save orchestration
  const [pendingSaves, setPendingSaves] = useState(0);
  const saveTimer = useRef(null);

  const debouncedSetSearch = useDebounce((value) => setSearchTerm(value), 300);

  // MSAL init
  useEffect(() => {
    msalInstance
      .initialize()
      .then(() => setMsalInitialized(true))
      .catch((e) => console.error("MSAL initialization failed:", e));
  }, []);

  // Existing session
  useEffect(() => {
    if (!msalInitialized) return;
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      msalInstance
        .acquireTokenSilent({ ...loginRequest, account: accounts[0] })
        .then((resp) => {
          setAccessToken(resp.accessToken);
          setIsAuthenticated(true);
          setUserName(accounts[0].name);
        })
        .catch((e) => console.error("Silent token acquisition failed:", e));
    }
  }, [msalInitialized]);

  // Initialize SharePoint Notes Service
  useEffect(() => {
    if (!isAuthenticated || !accessToken) return;

    const initNotesService = async () => {
      try {
        const token = await ensureAccessToken();
        const graph = Client.init({ authProvider: (done) => done(null, token) });
        const ods = new OneDriveService(graph);

        const service = new SharePointNotesService(ods, {
          spHostname: graphConfig.spHostname,
          spSitePath: graphConfig.spSitePath,
          spBasePath: graphConfig.spBasePath,
        });

        setNotesService(service);

        // Load all notes on startup
        console.log("📥 Loading notes from SharePoint...");
        const notes = await service.loadAllNotes();
        setNotesMap(notes);
        setLastNotesSync(new Date());
        console.log(`✅ Loaded ${notes.size} notes from SharePoint`);
      } catch (error) {
        console.error("Failed to initialize notes service:", error);
      }
    };

    initNotesService();
  }, [isAuthenticated, accessToken]);

  // Periodic refresh of notes (every 30 seconds)
  useEffect(() => {
    if (!notesService || !isAuthenticated) return;

    const interval = setInterval(async () => {
      try {
        console.log("🔄 Refreshing notes from SharePoint...");
        const notes = await notesService.loadAllNotes();
        setNotesMap(notes);
        setLastNotesSync(new Date());
      } catch (error) {
        console.error("Failed to refresh notes:", error);
      }
    }, 30000);

    return () => clearInterval(interval);
  }, [notesService, isAuthenticated]);

  // SharePoint data loader
  const loadFromSharePoint = useCallback(async (silent = false) => {
    if (!silent) setLoading(true);
    try {
      const token = await ensureAccessToken();
      const graph = Client.init({ authProvider: (done) => done(null, token) });
      const ods = new OneDriveService(graph);

      const HOST = graphConfig.spHostname;
      const SITE = graphConfig.spSitePath;
      const BASE = graphConfig.spBasePath;

      const [tickets, reports, mapping] = await Promise.all([
        ods
          .readExcelFromSharePoint({
            hostname: HOST,
            sitePath: SITE,
            fileRelativePath: `${BASE}/${graphConfig.ticketsFile}`,
          })
          .catch(() => []),
        ods
          .readExcelFromSharePoint({
            hostname: HOST,
            sitePath: SITE,
            fileRelativePath: `${BASE}/${graphConfig.reportsFile}`,
          })
          .catch(() => []),
        ods
          .readJsonFromSharePoint({
            hostname: HOST,
            sitePath: SITE,
            fileRelativePath: `${BASE}/${graphConfig.mappingFile}`,
          })
          .catch(() => []),
      ]);

      setTicketData(tickets);
      setReportData(reports);
      setCategoryMapping(mapping);
      setLastSync(new Date());

      if (!silent) {
        alert(
          `Loaded from SharePoint:\n${tickets.length} tickets\n${reports.length} reports\n${mapping.length} category mappings`
        );
      }
    } catch (e) {
      console.error("SharePoint load failed:", e);
      if (!silent) alert(e.message || "Failed to load from SharePoint.");
    } finally {
      if (!silent) setLoading(false);
    }
  }, []);

  // Load on auth
  useEffect(() => {
    if (!isAuthenticated || !accessToken) return;
    loadFromSharePoint();
  }, [isAuthenticated, accessToken, loadFromSharePoint]);

  // Auto refresh base data every 5 minutes
  useEffect(() => {
    if (!autoRefresh || !isAuthenticated || !accessToken) return;
    const interval = setInterval(() => loadFromSharePoint(true), 5 * 60 * 1000);
    return () => clearInterval(interval);
  }, [autoRefresh, isAuthenticated, accessToken, loadFromSharePoint]);

  const handleLogin = async () => {
    try {
      await msalInstance.loginPopup(loginRequest);
      const account = msalInstance.getAllAccounts()[0];
      const { accessToken } = await msalInstance.acquireTokenSilent({ ...loginRequest, account });
      setAccessToken(accessToken);
      setIsAuthenticated(true);
      setUserName(account.name);
    } catch (err) {
      console.error("Login failed:", err);
      alert("Failed to sign in to Microsoft. Please try again.");
    }
  };

  const handleLogout = () => {
    msalInstance.logoutPopup();
    setIsAuthenticated(false);
    setAccessToken(null);
    setUserName("");
    setTicketData([]);
    setReportData([]);
    setCategoryMapping([]);
  };

  // Build combined data
  const baseCombinedData = useMemo(() => {
    if (reportData.length === 0) return [];

    const normalizeBarcode = (x) => (x ? String(x).trim().toUpperCase() : "");

    const ticketMap = new Map();
    ticketData.forEach((t) => {
      const bc = normalizeBarcode(t["Barcode"]);
      if (bc) ticketMap.set(bc, t);
    });

    const categoryToPM = new Map();
    categoryMapping.forEach((m) => {
      if (m.category && m.pm) {
        categoryToPM.set(m.category.trim().toUpperCase(), {
          pm: m.pm,
          department: m.department || "",
          categoryText: m.category_text || "",
        });
      }
    });

    const unmatchedSet = new Set();

    const ageInDays = (dateStr) => {
      if (!dateStr) return "";
      const d = new Date(dateStr);
      if (isNaN(d)) return "";
      const today = new Date();
      return Math.ceil(Math.abs(today - d) / (1000 * 60 * 60 * 24));
    };

    const formatTicketNumber = (ticket) => {
      if (!ticket) return "";
      const cleaned = String(ticket).replace(/[^0-9.]/g, "");
      const num = parseFloat(cleaned);
      if (isNaN(num)) return "";
      return String(Math.floor(num));
    };

    const out = reportData.map((r) => {
      const bc = normalizeBarcode(r["Barcode#"]);
      const t = ticketMap.get(bc) || {};
      const category = (r["Category"] || "").trim();
      const mapInfo = categoryToPM.get(category.toUpperCase());
      const assignedPM = mapInfo ? mapInfo.pm : "";

      if (category && !assignedPM) unmatchedSet.add(category);

      return {
        "Meeting Note": "",
        "Requires Follow Up": "",
        "Assigned To": assignedPM,
        Location: r["Repair Location"] || t["Location"] || "",
        "Repair Ticket": formatTicketNumber(r["Ticket"]),
        "Asset Repair Age": ageInDays(r["Date In"]),
        "Barcode#": r["Barcode#"] || t["Barcode"] || "",
        Equipment: `(${r["Equipment"]}) - ${r["Description"]}`,
        "Damage Description": r["Notes"] || "",
        "Ticket Description": t["Notes"] || "",
        "Repair Reason": r["Repair Reason"] || "",
        "Last Order#": r["Last Order#"] || t["Order# to Bill"] || "",
        "Reference#": r["Reference#"] || "",
        Customer: r["Customer"] || t["Customer"] || "",
        "Customer Title": r["Customer Title"] || "",
        "Repair Cost": r["Repair Cost"] || "0",
        "Date In": r["Date In"] || t["Creation Date"] || "",
        Department: r["Department"] || "",
        Category: r["Category"] || "",
        Billable: t["Billable"] || r["Billable"] || "",
        "Created By": t["Created By"] || r["User In"] || "",
        "Repair Price": r["Repair Price"] || "0",
        "Repair Vendor": r["Repair Vendor"] || "",
        _TicketMatched: Object.keys(t).length > 0 ? "Yes" : "No",
      };
    });

    setUnmatchedCategories(Array.from(unmatchedSet).sort());
    return out;
  }, [ticketData, reportData, categoryMapping]);

  // Merge notes with base data
  useEffect(() => {
    const merged = baseCombinedData.map((row) => {
      const bc = row["Barcode#"];
      const note = notesMap.get(bc) || { meetingNote: "", requiresFollowUp: "" };
      return {
        ...row,
        "Meeting Note": note.meetingNote,
        "Requires Follow Up": note.requiresFollowUp,
      };
    });
    setCombinedDataWithNotes(merged);
  }, [baseCombinedData, notesMap]);

  // Derived UI values
  const getCurrentData = useCallback(() => {
    switch (activeTab) {
      case "tickets":
        return ticketData;
      case "reports":
        return reportData;
      case "combined":
        return combinedDataWithNotes;
      default:
        return [];
    }
  }, [activeTab, ticketData, reportData, combinedDataWithNotes]);

  const currentData = getCurrentData();
  const columns =
    currentData.length > 0 ? Object.keys(currentData[0]).filter((c) => !c.startsWith("_")) : [];

  const uniqueLocations = useMemo(() => {
    const s = new Set();
    if (activeTab === "combined" || activeTab === "reports") {
      reportData.forEach((r) => {
        const loc = r["Repair Location"];
        if (loc && String(loc).trim()) s.add(String(loc).trim());
      });
    }
    if (activeTab === "combined" || activeTab === "tickets") {
      ticketData.forEach((r) => {
        const loc = r["Location"];
        if (loc && String(loc).trim()) s.add(String(loc).trim());
      });
    }
    return Array.from(s).sort();
  }, [reportData, ticketData, activeTab]);

  const [allCategories, uniquePMs] = useMemo(() => {
    const cats = new Set();
    const pms = new Set();
    reportData.forEach((r) => r["Category"] && cats.add(r["Category"].trim()));
    combinedDataWithNotes.forEach((row) => {
      const pm = row["Assigned To"];
      if (pm && pm.trim()) pms.add(pm.trim());
    });
    return [Array.from(cats).sort(), Array.from(pms).sort()];
  }, [reportData, combinedDataWithNotes]);

  // Category mapping helpers
  const addCategoryMapping = (category, pm, department = "", categoryText = "") => {
    setCategoryMapping((prev) => {
      const copy = [...prev];
      const i = copy.findIndex(
        (m) => m.category.trim().toUpperCase() === category.trim().toUpperCase()
      );
      const entry = {
        category: category.trim(),
        pm: pm.trim(),
        department: department.trim(),
        category_text: categoryText.trim(),
      };
      if (i >= 0) copy[i] = entry;
      else copy.push(entry);
      return copy;
    });
  };

  const removeCategoryMapping = (category) => {
    setCategoryMapping((prev) =>
      prev.filter((m) => m.category.trim().toUpperCase() !== category.trim().toUpperCase())
    );
  };

  const exportCategoryMapping = () => {
    const json = JSON.stringify(categoryMapping, null, 2);
    const blob = new Blob([json], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `category_mapping_${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const downloadNotesTemplate = () => {
    const templateData = [
      {
        "Barcode#": "RV123456",
        "Meeting Note": "Example: Cable tested and working properly",
        "Requires Follow Up": "Example: Ship to customer location",
      },
      {
        "Barcode#": "MC987654",
        "Meeting Note": "Example: Screen cracked, needs replacement",
        "Requires Follow Up": "Example: Order new screen from vendor",
      },
      {
        "Barcode#": "RV555555",
        "Meeting Note": "Example: Battery issue resolved",
        "Requires Follow Up": "",
      },
    ];

    const ws = XLSX.utils.json_to_sheet(templateData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Notes Template");
    ws["!cols"] = [{ wch: 15 }, { wch: 50 }, { wch: 40 }];
    XLSX.writeFile(wb, "notes_import_template.xlsx");
  };

  // Import notes from Excel
  const importNotesFromExcel = async (file) => {
    if (!file || !notesService) return;

    setIsImporting(true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws);

      const notesArray = rows
        .map((row) => ({
          barcode: row["Barcode#"] || row["Barcode"],
          meetingNote: row["Meeting Note"] || "",
          requiresFollowUp: row["Requires Follow Up"] || "",
        }))
        .filter((note) => note.barcode);

      await notesService.importNotes(notesArray);

      // Refresh display
      const notes = await notesService.loadAllNotes();
      setNotesMap(notes);
      setLastNotesSync(new Date());

      alert(`✅ Imported ${notesArray.length} notes to SharePoint`);
    } catch (error) {
      console.error("Import failed:", error);
      alert(`❌ Import failed: ${error.message}`);
    } finally {
      setIsImporting(false);
    }
  };

  const CategoryManager = () => {
    const [newCategory, setNewCategory] = useState("");
    const [newPM, setNewPM] = useState("");
    const [newDepartment, setNewDepartment] = useState("");
    const [newCategoryText, setNewCategoryText] = useState("");

    return (
      <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
        <div className="bg-white rounded-lg shadow-xl max-w-6xl w-full max-h-[90vh] overflow-hidden flex flex-col">
          <div className="p-6 border-b flex justify-between items-center">
            <h2 className="text-2xl font-bold text-gray-800">Category to PM Mapping Manager</h2>
            <button
              onClick={() => setShowCategoryManager(false)}
              className="text-gray-500 hover:text-gray-700 text-2xl"
            >
              ×
            </button>
          </div>

          <div className="p-6 space-y-6 overflow-y-auto flex-1">
            <div className="bg-blue-50 p-4 rounded-lg">
              <h3 className="font-semibold text-blue-900 mb-3">Add New Mapping</h3>
              <div className="grid grid-cols-2 gap-3 mb-3">
                <div>
                  <label className="text-xs text-gray-600 mb-1 block">Category Code</label>
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
                  <label className="text-xs text-gray-600 mb-1 block">Department</label>
                  <input
                    type="text"
                    placeholder="Department (optional)"
                    value={newDepartment}
                    onChange={(e) => setNewDepartment(e.target.value)}
                    className="w-full px-3 py-2 border border-gray-300 rounded-lg"
                  />
                </div>
                <div>
                  <label className="text-xs text-gray-600 mb-1 block">Category Description</label>
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
                onClick={() => {
                  if (newCategory && newPM) {
                    addCategoryMapping(newCategory, newPM, newDepartment, newCategoryText);
                    setNewCategory("");
                    setNewPM("");
                    setNewDepartment("");
                    setNewCategoryText("");
                  }
                }}
                disabled={!newCategory || !newPM}
                className="w-full px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 disabled:opacity-50"
              >
                Add Mapping
              </button>
            </div>

            {unmatchedCategories.length > 0 && (
              <div className="bg-red-50 p-4 rounded-lg">
                <h3 className="font-semibold text-red-900 mb-2">
                  Unmatched Categories ({unmatchedCategories.length})
                </h3>
                <p className="text-sm text-red-800 mb-3">These categories don't have PM assignments:</p>
                <div className="flex flex-wrap gap-2">
                  {unmatchedCategories.map((cat) => (
                    <span key={cat} className="px-3 py-1 bg-red-100 text-red-800 rounded-full text-sm">
                      {cat}
                    </span>
                  ))}
                </div>
              </div>
            )}

            <div>
              <div className="flex justify-between items-center mb-3">
                <h3 className="font-semibold text-gray-800">Current Mappings ({categoryMapping.length})</h3>
                <button
                  onClick={exportCategoryMapping}
                  className="text-sm px-3 py-1 bg-gray-600 text-white rounded hover:bg-gray-700"
                >
                  Export JSON
                </button>
              </div>
              <div className="space-y-2 max-h-96 overflow-y-auto">
                {categoryMapping.map((m, idx) => (
                  <div key={idx} className="flex items-start justify-between p-4 bg-gray-50 rounded-lg border">
                    <div className="flex-1 space-y-1">
                      <div className="flex items-center gap-2">
                        <span className="font-bold text-gray-900">{m.category}</span>
                        <span className="text-gray-400">→</span>
                        <span className="font-semibold text-blue-600">{m.pm}</span>
                      </div>
                      {m.category_text && <p className="text-sm text-gray-600">{m.category_text}</p>}
                      {m.department && <p className="text-xs text-gray-500">Department: {m.department}</p>}
                    </div>
                    <button
                      onClick={() => removeCategoryMapping(m.category)}
                      className="text-red-600 hover:text-red-800 text-sm ml-4"
                    >
                      Remove
                    </button>
                  </div>
                ))}
              </div>
            </div>

            <div>
              <h3 className="font-semibold text-gray-800 mb-3">
                All Categories in Data ({allCategories.length})
              </h3>
              <div className="flex flex-wrap gap-2">
                {allCategories.map((cat) => {
                  const hasMapping = categoryMapping.some(
                    (m) => m.category.trim().toUpperCase() === cat.trim().toUpperCase()
                  );
                  return (
                    <span
                      key={cat}
                      className={`px-3 py-1 rounded-full text-sm ${
                        hasMapping ? "bg-green-100 text-green-800" : "bg-gray-100 text-gray-800"
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

  // Filtering, sorting, export
  const hasData = ticketData.length > 0 || reportData.length > 0;

  useEffect(() => {
    setCurrentPage(1);
  }, [searchTerm, locationFilter, pmFilter, activeTab]);

  const filteredAndSortedData = useMemo(() => {
    let rows = getCurrentData();

    if (locationFilter) {
      rows = rows.filter((r) => {
        const loc = r["Location"] || r["Repair Location"];
        return String(loc || "") === locationFilter;
      });
    }

    if (activeTab === "combined" && pmFilter) {
      rows = rows.filter((r) => {
        if (pmFilter === "__unassigned__") return !r["Assigned To"] || r["Assigned To"] === "";
        return r["Assigned To"] === pmFilter;
      });
    }

    if (searchTerm) {
      const q = searchTerm.toLowerCase();
      rows = rows.filter((r) => Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q)));
    }

    if (sortConfig.key) {
      rows = [...rows].sort((a, b) => {
        const av = a[sortConfig.key];
        const bv = b[sortConfig.key];
        if (av === bv) return 0;
        return (av > bv ? 1 : -1) * (sortConfig.direction === "asc" ? 1 : -1);
      });
    }

    return rows;
  }, [getCurrentData, activeTab, locationFilter, pmFilter, searchTerm, sortConfig]);

  const handleSort = (key) =>
    setSortConfig((prev) => ({
      key,
      direction: prev.key === key && prev.direction === "asc" ? "desc" : "asc",
    }));

  const openRowEditor = (idx) => {
    setEditingRowIndex(idx);
    setEditingRow(filteredAndSortedData[idx]);
  };

  const closeRowEditor = () => {
    setEditingRow(null);
    setEditingRowIndex(null);
  };

  const handleNoteSaved = async () => {
    // Refresh notes after modal save
    if (notesService) {
      const notes = await notesService.loadAllNotes();
      setNotesMap(notes);
      setLastNotesSync(new Date());
    }
  };

  const exportToCSV = () => {
    const rows = filteredAndSortedData;
    if (rows.length === 0) return;
    const headers = columns.join(",");
    const body = rows
      .map((row) =>
        columns
          .map((col) => {
            const val = row[col];
            const text = String(val ?? "").replace(/"/g, '""');
            return `"${text}"`;
          })
          .join(",")
      )
      .join("\n");
    const csv = [headers, body].join("\n");
    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${activeTab}_export_${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  /* ---------------- Quick Edit save orchestration ---------------- */
  const queueSaveAndAutoFlush = useCallback(() => {
    setPendingSaves((n) => n + 1);
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      try {
        await notesService?.saveToSharePoint();
        setPendingSaves(0);
        setLastNotesSync(new Date());
      } catch (e) {
        console.error("Save error:", e);
      }
    }, 10000);
  }, [notesService]);

  const saveNow = useCallback(async () => {
    clearTimeout(saveTimer.current);
    try {
      await notesService?.saveToSharePoint();
      setPendingSaves(0);
      setLastNotesSync(new Date());
    } catch (e) {
      console.error("Save error:", e);
    }
  }, [notesService]);

  // Inline handlers: update service cache + live UI map so table refreshes immediately
  const onInlineNoteChange = useCallback(
    (barcode, next) => {
      if (!barcode || !notesService) return;
      notesService.updateNote(barcode, next, undefined);
      setNotesMap((prev) => {
        const copy = new Map(prev);
        const cur = copy.get(barcode) || { barcode, meetingNote: "", requiresFollowUp: "" };
        copy.set(barcode, { ...cur, meetingNote: next });
        return copy;
      });
      queueSaveAndAutoFlush();
    },
    [notesService, queueSaveAndAutoFlush]
  );

  const onInlineFollowUpChange = useCallback(
    (barcode, next) => {
      if (!barcode || !notesService) return;
      notesService.updateNote(barcode, undefined, next);
      setNotesMap((prev) => {
        const copy = new Map(prev);
        const cur = copy.get(barcode) || { barcode, meetingNote: "", requiresFollowUp: "" };
        copy.set(barcode, { ...cur, requiresFollowUp: next });
        return copy;
      });
      queueSaveAndAutoFlush();
    },
    [notesService, queueSaveAndAutoFlush]
  );

  // Render
  const hasDataNow = ticketData.length > 0 || reportData.length > 0;

  return (
    <div className="w-full h-screen flex flex-col bg-gray-50">
      {/* Header */}
      <div className="bg-white border-b px-6 py-4">
        <div className="flex items-center justify-between">
          <div>
            <h1 className="text-2xl font-bold text-gray-800">Repair Tracker Dashboard</h1>
            <div className="flex items-center gap-3 mt-1">
              <p className="text-sm text-gray-500">
                {isAuthenticated ? `Connected as ${userName}` : "Not connected"}
              </p>
              {isAuthenticated && (
                <div className="flex items-center gap-1 text-xs text-green-600 bg-green-50 px-2 py-1 rounded">
                  <Cloud size={12} />
                  SharePoint
                </div>
              )}
              {lastSync && <span className="text-xs text-gray-500">Data: {lastSync.toLocaleTimeString()}</span>}
              {lastNotesSync && (
                <span className="text-xs text-blue-500">Notes: {lastNotesSync.toLocaleTimeString()}</span>
              )}
            </div>
          </div>

          <div className="flex gap-2 items-center flex-wrap">
            {!isAuthenticated ? (
              <button
                onClick={handleLogin}
                className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 text-sm"
              >
                <Cloud size={16} />
                Sign in
              </button>
            ) : (
              <>
                {/* Pending changes badge */}
                {pendingSaves > 0 ? (
                  <div className="text-amber-600 text-sm">💾 {pendingSaves} pending change(s)</div>
                ) : (
                  <div className="text-emerald-600 text-sm">✓ All changes saved</div>
                )}

                <button
                  onClick={saveNow}
                  className="flex items-center gap-2 px-4 py-2 bg-cyan-600 text-white rounded-lg hover:bg-cyan-700 text-sm"
                >
                  <Save size={16} />
                  Save now
                </button>

                <button
                  onClick={() => loadFromSharePoint(false)}
                  disabled={loading}
                  className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 disabled:opacity-50 text-sm"
                >
                  <RefreshCw size={16} className={loading ? "animate-spin" : ""} />
                  Refresh
                </button>
                <button
                  onClick={() => setShowCategoryManager(true)}
                  className="px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 text-sm"
                >
                  Manage Categories
                  {unmatchedCategories.length > 0 && (
                    <span className="ml-2 bg-red-500 text-white px-2 py-0.5 rounded-full text-xs">
                      {unmatchedCategories.length}
                    </span>
                  )}
                </button>
                <button
                  onClick={handleLogout}
                  className="px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 text-sm"
                >
                  Sign Out
                </button>
              </>
            )}

            <label className="flex items-center gap-2 px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700 cursor-pointer transition-colors text-sm">
              <Upload size={16} />
              Upload Mapping
              <input
                type="file"
                accept=".json"
                onChange={async (e) => {
                  const f = e.target.files?.[0];
                  if (!f) return;
                  const text = await f.text();
                  try {
                    const json = JSON.parse(text);
                    setCategoryMapping(json);
                    alert(`Loaded ${json.length} category mappings`);
                  } catch (err) {
                    alert("Invalid JSON");
                  }
                }}
                className="hidden"
                disabled={loading}
              />
            </label>

            <label className="flex items-center gap-2 px-4 py-2 bg-emerald-600 text-white rounded-lg hover:bg-emerald-700 cursor-pointer transition-colors text-sm">
              <Upload size={16} />
              Upload Notes
              <input
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={(e) => importNotesFromExcel(e.target.files?.[0])}
                className="hidden"
                disabled={loading || isImporting || !notesService}
              />
            </label>

            <button
              onClick={downloadNotesTemplate}
              className="flex items-center gap-2 px-4 py-2 bg-teal-600 text-white rounded-lg hover:bg-teal-700 transition-colors text-sm"
              title="Download Excel template for notes import"
            >
              <Download size={16} />
              Notes Template
            </button>
          </div>
        </div>
      </div>

      {/* Toolbar */}
      <div className="bg-white border-b">
        <div className="flex items-center justify-between px-6 pt-3 border-b">
          <div className="flex">
            <button
              onClick={() => setActiveTab("combined")}
              className={`px-4 py-2 font-medium border-b-2 transition-colors ${
                activeTab === "combined" ? "border-blue-500 text-blue-600" : "border-transparent text-gray-500"
              }`}
            >
              Combined ({combinedDataWithNotes.length})
            </button>
            <button
              onClick={() => setActiveTab("tickets")}
              className={`px-4 py-2 font-medium border-b-2 transition-colors ${
                activeTab === "tickets" ? "border-blue-500 text-blue-600" : "border-transparent text-gray-500"
              }`}
            >
              Tickets ({ticketData.length})
            </button>
            <button
              onClick={() => setActiveTab("reports")}
              className={`px-4 py-2 font-medium border-b-2 transition-colors ${
                activeTab === "reports" ? "border-blue-500 text-blue-600" : "border-transparent text-gray-500"
              }`}
            >
              Reports ({reportData.length})
            </button>
            <button
              onClick={() => setActiveTab("diagnostics")}
              className={`px-4 py-2 font-medium border-b-2 transition-colors ${
                activeTab === "diagnostics" ? "border-orange-500 text-orange-600" : "border-transparent text-gray-500"
              }`}
            >
              Diagnostics
            </button>
          </div>

          <button
            onClick={exportToCSV}
            disabled={getCurrentData().length === 0}
            className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors disabled:opacity-50 text-sm"
          >
            <Download size={18} />
            Export
          </button>
        </div>

        <div className="flex items-center px-6 py-3 gap-3">
          <div className="flex-1 relative max-w-md">
            <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
            <input
              type="text"
              placeholder="Search across all columns..."
              value={searchInput}
              onChange={(e) => {
                setSearchInput(e.target.value);
                debouncedSetSearch(e.target.value);
              }}
              className="w-full pl-10 pr-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm"
            />
          </div>

          <div className="flex items-center gap-2">
            <label className="text-sm text-gray-600 whitespace-nowrap">Rows per page:</label>
            <select
              value={itemsPerPage}
              onChange={(e) => {
                const newSize = parseInt(e.target.value);
                setItemsPerPage(newSize);
                setCurrentPage(1);
              }}
              className="px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white text-sm"
            >
              <option value={50}>50</option>
              <option value={100}>100</option>
              <option value={200}>200</option>
              <option value={500}>500</option>
              <option value={1000}>1,000</option>
              <option value={99999}>All</option>
            </select>
          </div>

          <select
            value={locationFilter}
            onChange={(e) => setLocationFilter(e.target.value)}
            className="px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white text-sm"
          >
            <option value="">All Locations</option>
            {uniqueLocations.map((loc) => (
              <option key={loc} value={loc}>
                {loc}
              </option>
            ))}
          </select>

          {locationFilter && (
            <button
              onClick={() => setLocationFilter("")}
              className="px-3 py-2 text-xs text-gray-600 hover:text-gray-800 hover:bg-gray-100 rounded-lg transition-colors"
            >
              Clear Location
            </button>
          )}

          {activeTab === "combined" && (
            <>
              <select
                value={pmFilter}
                onChange={(e) => setPmFilter(e.target.value)}
                className="px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white text-sm"
              >
                <option value="">All Assigned To</option>
                <option value="__unassigned__">Unassigned</option>
                {uniquePMs.map((pm) => (
                  <option key={pm} value={pm}>
                    {pm}
                  </option>
                ))}
              </select>

              {pmFilter && (
                <button
                  onClick={() => setPmFilter("")}
                  className="px-3 py-2 text-xs text-gray-600 hover:text-gray-800 hover:bg-gray-100 rounded-lg transition-colors"
                >
                  Clear PM
                </button>
              )}
            </>
          )}
        </div>
      </div>

      {/* Body */}
      <div className="flex-1 overflow-hidden px-6 py-4">
        {loading ? (
          <div className="flex items-center justify-center h-full">
            <div className="text-center">
              <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4" />
              <p className="text-gray-500">Loading...</p>
            </div>
          </div>
        ) : activeTab === "diagnostics" ? (
          <div className="max-w-6xl mx-auto space-y-6 overflow-y-auto h-full pb-8">
            <div className="bg-white p-6 rounded-lg shadow">
              <h2 className="text-xl font-semibold text-gray-800 mb-4">System Diagnostics</h2>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
                <div className="p-4 bg-blue-50 rounded-lg">
                  <h3 className="font-semibold text-blue-900 mb-2">Repair Ticket List</h3>
                  <p className="text-2xl font-bold text-blue-800">{ticketData.length}</p>
                  <p className="text-sm text-blue-700">records</p>
                </div>
                <div className="p-4 bg-green-50 rounded-lg">
                  <h3 className="font-semibold text-green-900 mb-2">Repair Report</h3>
                  <p className="text-2xl font-bold text-green-800">{reportData.length}</p>
                  <p className="text-sm text-green-700">records</p>
                </div>
                <div className="p-4 bg-purple-50 rounded-lg">
                  <h3 className="font-semibold text-purple-900 mb-2">Notes (SharePoint)</h3>
                  <p className="text-2xl font-bold text-purple-800">{notesMap.size}</p>
                  <p className="text-sm text-purple-700">notes stored</p>
                </div>
              </div>

              <div className="mt-6 p-4 bg-green-50 rounded-lg border border-green-200">
                <h3 className="font-semibold text-green-900 mb-3">💰 SharePoint Storage Benefits</h3>
                <div className="space-y-2 text-sm text-green-800">
                  <div className="flex items-start gap-2">
                    <span className="text-green-600">✓</span>
                    <div>
                      <strong>Zero Firebase costs</strong>
                      <p className="text-xs text-green-700">All notes stored in SharePoint Excel file</p>
                    </div>
                  </div>
                  <div className="flex items-start gap-2">
                    <span className="text-green-600">✓</span>
                    <div>
                      <strong>Manual save (prevents accidental overwrites)</strong>
                      <p className="text-xs text-green-700">Click "Save to SharePoint" button when ready</p>
                    </div>
                  </div>
                  <div className="flex items-start gap-2">
                    <span className="text-green-600">✓</span>
                    <div>
                      <strong>30-second refresh cycle</strong>
                      <p className="text-xs text-green-700">Automatically syncs with SharePoint</p>
                    </div>
                  </div>
                  <div className="flex items-start gap-2">
                    <span className="text-green-600">✓</span>
                    <div>
                      <strong>All data in one place</strong>
                      <p className="text-xs text-green-700">Notes stored alongside tickets and reports</p>
                    </div>
                  </div>
                </div>
                <div className="mt-4 pt-4 border-t border-green-300">
                  <p className="text-sm font-semibold text-green-900">💡 Estimated annual savings: $120-600 vs Firebase</p>
                  <p className="text-xs text-green-700 mt-1">
                    File location: SharePoint/Shared Documents/repair_notes.xlsx
                  </p>
                </div>
              </div>
            </div>
          </div>
        ) : !hasDataNow ? (
          <div className="flex items-center justify-center h-full">
            <div className="text-center max-w-lg bg-white p-12 rounded-lg shadow-lg">
              <FileSpreadsheet className="mx-auto text-blue-500 mb-6" size={64} />
              <h3 className="text-2xl font-semibold text-gray-800 mb-3">Welcome to Repair Tracker</h3>
              <p className="text-gray-600 mb-6">Sign in to load data from SharePoint.</p>
            </div>
          </div>
        ) : getCurrentData().length === 0 ? (
          <div className="flex items-center justify-center h-full">
            <div className="text-center bg-white p-8 rounded-lg shadow">
              <p className="text-gray-500">No data available</p>
            </div>
          </div>
        ) : (
          <PaginatedTable
            data={filteredAndSortedData}
            columns={columns}
            onRowClick={openRowEditor}
            activeTab={activeTab}
            currentPage={currentPage}
            setCurrentPage={setCurrentPage}
            itemsPerPage={ITEMS_PER_PAGE}
            sortConfig={sortConfig}
            onSort={handleSort}
            // inline editing
            notesService={notesService}
            onInlineNoteChange={onInlineNoteChange}
            onInlineFollowUpChange={onInlineFollowUpChange}
            onInlineSaveNow={saveNow}
          />
        )}
      </div>

      {/* Footer */}
      {hasDataNow && activeTab !== "diagnostics" && (
        <div className="bg-white border-t px-6 py-3">
          <div className="flex items-center justify-between text-sm text-gray-600">
            <span>
              {itemsPerPage >= 99999 ? (
                `Showing all ${filteredAndSortedData.length} records`
              ) : (
                <>
                  Page {currentPage} of {Math.ceil(filteredAndSortedData.length / ITEMS_PER_PAGE)}{" "}
                  ({filteredAndSortedData.length} total records)
                </>
              )}
            </span>
            <div className="flex items-center gap-4">
              {locationFilter && <span className="text-blue-600">Location: {locationFilter}</span>}
              {activeTab === "combined" && pmFilter && (
                <span className="text-green-600">
                  Assigned To: {pmFilter === "__unassigned__" ? "Unassigned" : pmFilter}
                </span>
              )}
              {searchTerm && <span className="text-blue-600">Search: "{searchTerm}"</span>}
            </div>
          </div>
        </div>
      )}

      {showCategoryManager && <CategoryManager />}

      {editingRow && (
        <RowEditor
          row={editingRow}
          rowIndex={editingRowIndex}
          onClose={closeRowEditor}
          notesService={notesService}
          onSave={handleNoteSaved}
        />
      )}
    </div>
  );
};

export default RepairTrackerSheet;
