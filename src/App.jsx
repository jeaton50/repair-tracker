// src/App.jsx - OPTIMIZED VERSION
import React, { useState, useMemo, useEffect, useRef, useCallback } from "react";
import { Client } from "@microsoft/microsoft-graph-client";
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig, loginRequest, graphConfig } from "./authConfig";
import OneDriveService from "./oneDriveService";
import { db } from "./firebase";
import * as XLSX from "xlsx";
import { doc, setDoc, onSnapshot, serverTimestamp, deleteDoc, writeBatch, getDoc } from "firebase/firestore";
import { Search, Download, ChevronDown, ChevronUp, Upload, FileSpreadsheet, RefreshCw, Cloud, ChevronLeft, ChevronRight } from "lucide-react";

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
  
  return useCallback((...args) => {
    if (timeoutRef.current) clearTimeout(timeoutRef.current);
    timeoutRef.current = setTimeout(() => callback(...args), delay);
  }, [callback, delay]);
};

// ---------- Row Editor (Optimized - Load on demand) ----------
const RowEditor = ({ row, rowIndex, onClose }) => {
  const [meetingNote, setMeetingNote] = useState("");
  const [followUp, setFollowUp] = useState("");
  const [lastSaved, setLastSaved] = useState(null);
  const [isSaving, setIsSaving] = useState(false);
  const [isLoading, setIsLoading] = useState(true);
  const saveTimeoutRef = useRef(null);
  const barcode = row["Barcode#"];

  // Load notes once when editor opens
  useEffect(() => {
    if (!barcode) {
      setIsLoading(false);
      return;
    }
    
    const loadNotes = async () => {
      try {
        const ref = doc(db, "repairNotes", barcode);
        const snap = await getDoc(ref);
        
        if (snap.exists()) {
          const d = snap.data();
          setMeetingNote(d.meetingNote || "");
          setFollowUp(d.requiresFollowUp || "");
        }
      } catch (e) {
        console.error("Failed to load notes:", e);
      } finally {
        setIsLoading(false);
      }
    };

    loadNotes();
  }, [barcode]);

  // Auto-save
  useEffect(() => {
    if (isLoading) return;
    if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current);
    saveTimeoutRef.current = setTimeout(async () => {
      if (!barcode) return;
      setIsSaving(true);
      try {
        const ref = doc(db, "repairNotes", barcode);
        await setDoc(
          ref,
          { barcode, meetingNote, requiresFollowUp: followUp, lastUpdated: serverTimestamp() },
          { merge: true }
        );
        setLastSaved(new Date());
      } catch (e) {
        console.error("Auto-save error:", e);
        alert("Failed to save. Check your Firebase connection.");
      } finally {
        setIsSaving(false);
      }
    }, 1000);
    return () => saveTimeoutRef.current && clearTimeout(saveTimeoutRef.current);
  }, [meetingNote, followUp, barcode, isLoading]);

  // Real-time sync only for this ONE row
  useEffect(() => {
    if (!barcode || isLoading) return;
    const ref = doc(db, "repairNotes", barcode);
    const unsub = onSnapshot(
      ref,
      (snap) => {
        if (snap.exists()) {
          const d = snap.data();
          if (document.activeElement?.name !== "meetingNote") setMeetingNote(d.meetingNote || "");
          if (document.activeElement?.name !== "followUp") setFollowUp(d.requiresFollowUp || "");
        }
      },
      (err) => console.error("Real-time sync error:", err)
    );
    return () => unsub();
  }, [barcode, isLoading]);

  const handleClearAll = () => {
    if (window.confirm("Clear all notes for this item?")) {
      setMeetingNote("");
      setFollowUp("");
    }
  };

  const handleDeleteFromDatabase = async () => {
    if (!window.confirm("Permanently delete all notes for this item from the database?")) return;
    try {
      await deleteDoc(doc(db, "repairNotes", barcode));
      setMeetingNote("");
      setFollowUp("");
      alert("Notes deleted from database");
    } catch (e) {
      console.error("Delete error:", e);
      alert("Failed to delete. Please try again.");
    }
  };

  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-lg shadow-xl w-full max-w-4xl max-h-[90vh] overflow-hidden flex flex-col">
        <div className="p-6 border-b flex justify-between items-center">
          <div>
            <h2 className="text-xl font-bold text-gray-800">Edit Repair Item</h2>
            <p className="text-sm text-gray-500 mt-1">
              {row["Barcode#"]} - {row["Equipment"]}
            </p>
            {isLoading && <p className="text-xs text-blue-600 mt-1">⟳ Loading...</p>}
            {isSaving && !isLoading && <p className="text-xs text-blue-600 mt-1">⟳ Saving...</p>}
            {lastSaved && !isSaving && !isLoading && (
              <p className="text-xs text-green-600 mt-1">✓ Saved at {lastSaved.toLocaleTimeString()}</p>
            )}
          </div>
          <button onClick={onClose} className="text-gray-500 hover:text-gray-700 text-2xl">
            ×
          </button>
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
              name="meetingNote"
              value={meetingNote}
              onChange={(e) => setMeetingNote(e.target.value)}
              placeholder="Add meeting notes here... (auto-saves)"
              className="w-full h-32 px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent resize-none"
              disabled={isLoading}
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
              name="followUp"
              value={followUp}
              onChange={(e) => setFollowUp(e.target.value)}
              placeholder="Add follow-up notes here... (auto-saves)"
              className="w-full h-24 px-4 py-3 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent resize-none"
              disabled={isLoading}
            />
          </div>

          <div className="p-4 bg-gray-50 rounded-lg text-sm space-y-2">
            <div>
              <span className="font-semibold text-gray-700">Damage:</span> {row["Damage Description"]}
            </div>
            <div>
              <span className="font-semibold text-gray-700">Ticket Notes:</span> {row["Ticket Description"]}
            </div>
            <div>
              <span className="font-semibold text-gray-700">Reason:</span> {row["Repair Reason"]}
            </div>
          </div>
        </div>

        <div className="p-6 border-t flex justify-between">
          <div className="flex gap-2">
            <button
              onClick={handleClearAll}
              className="px-6 py-2 text-orange-600 bg-orange-50 border border-orange-300 rounded-lg hover:bg-orange-100"
              disabled={isLoading}
            >
              Clear All Notes
            </button>
            <button
              onClick={handleDeleteFromDatabase}
              className="px-6 py-2 text-red-600 bg-red-50 border border-red-300 rounded-lg hover:bg-red-100"
              disabled={isLoading}
            >
              Delete from Database
            </button>
          </div>
          <button onClick={onClose} className="px-6 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700">
            Close
          </button>
        </div>
      </div>
    </div>
  );
};

// ---------- Paginated Table Component ----------
const PaginatedTable = ({ data, columns, onRowClick, activeTab, currentPage, setCurrentPage, itemsPerPage, sortConfig, onSort }) => {
  const totalPages = Math.ceil(data.length / itemsPerPage);
  const startIdx = (currentPage - 1) * itemsPerPage;
  const endIdx = startIdx + itemsPerPage;
  const paginatedData = data.slice(startIdx, endIdx);

  return (
    <div className="h-full flex flex-col bg-white rounded-lg shadow">
      {/* Table */}
      <div className="flex-1 overflow-auto">
        <table className="w-full border-collapse">
          <thead className="bg-gray-50 border-b sticky top-0 z-10">
            <tr>
              {columns.map((col) => {
                const isEditable =
                  activeTab === "combined" && (col === "Meeting Note" || col === "Requires Follow Up");
                return (
                  <th
                    key={col}
                    onClick={() => onSort(col)}
                    className="px-4 py-3 text-left text-xs font-medium text-gray-700 uppercase tracking-wider whitespace-normal bg-gray-50 cursor-pointer hover:bg-gray-100"
                  >
                    <div className="flex items-center gap-2">
                      {col}
                      {isEditable && <span className="text-blue-500">✏️</span>}
                      {sortConfig.key === col && (
                        sortConfig.direction === "asc" ? <ChevronUp size={14} /> : <ChevronDown size={14} />
                      )}
                    </div>
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody className="bg-white divide-y">
            {paginatedData.map((row, idx) => {
              const hasAssignment = row["Assigned To"] && row["Assigned To"] !== "";
              const rowBg = activeTab === "combined" && !hasAssignment ? "bg-red-50" : "";
              const isEditable = activeTab === "combined";
              const actualIndex = startIdx + idx;
              
              return (
                <tr
                  key={actualIndex}
                  className={`${rowBg} ${isEditable ? "hover:bg-blue-50 cursor-pointer" : "hover:bg-gray-50"}`}
                  onClick={() => isEditable && onRowClick(actualIndex)}
                >
                  {columns.map((col) => (
                    <td
                      key={col}
                      className="px-4 py-3 text-sm text-gray-900 whitespace-normal break-words"
                      style={{ maxWidth: 300 }}
                    >
                      {String(row[col] ?? "")}
                    </td>
                  ))}
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>

      {/* Pagination Controls */}
      {totalPages > 1 && (
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
                onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
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
                onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
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
  const [notesCache, setNotesCache] = useState(new Map());

  const [isImporting, setIsImporting] = useState(false);
  const isImportingRef = useRef(false);
  useEffect(() => { isImportingRef.current = isImporting; }, [isImporting]);
  const notesListenersRef = useRef(new Map());

  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [accessToken, setAccessToken] = useState(null);
  const [userName, setUserName] = useState("");
  const [lastSync, setLastSync] = useState(null);
  const [autoRefresh, setAutoRefresh] = useState(true);
  const [msalInitialized, setMsalInitialized] = useState(false);
  const [loading, setLoading] = useState(false);
  
  const [locationFilter, setLocationFilter] = useState("");
  const [pmFilter, setPmFilter] = useState("");
  
  // Pagination
  const [currentPage, setCurrentPage] = useState(1);
  const [itemsPerPage, setItemsPerPage] = useState(100);
  const ITEMS_PER_PAGE = itemsPerPage;
  
  // For tracking visible rows (now based on pagination)
  const MAX_LISTENERS = 100;

  // Debounced search
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

  // ---------- SharePoint data loader ----------
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

  // Auto refresh
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

  // ---------- Build combined data ----------
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

  // ---------- OPTIMIZED Notes live merge with LIMITED listeners ----------
  useEffect(() => {
    if (baseCombinedData.length === 0) {
      setCombinedDataWithNotes([]);
      return;
    }

    // Only subscribe to notes for current page + buffer
    const startIdx = Math.max(0, (currentPage - 1) * ITEMS_PER_PAGE - 25);
    const endIdx = Math.min(baseCombinedData.length, currentPage * ITEMS_PER_PAGE + 25);
    const pageBarcodes = baseCombinedData
      .slice(startIdx, endIdx)
      .map(row => row["Barcode#"])
      .filter(Boolean)
      .slice(0, MAX_LISTENERS);

    // Cleanup old listeners for barcodes no longer on current page
    notesListenersRef.current.forEach((unsub, bc) => {
      if (!pageBarcodes.includes(bc)) {
        unsub();
        notesListenersRef.current.delete(bc);
      }
    });

    // Subscribe to current page barcodes only
    pageBarcodes.forEach((bc) => {
      if (notesListenersRef.current.has(bc)) return;

      const ref = doc(db, "repairNotes", bc);
      const unsub = onSnapshot(
        ref,
        (snap) => {
          if (isImportingRef.current) return;
          
          const noteData = snap.exists() 
            ? {
                meetingNote: snap.data().meetingNote || "",
                requiresFollowUp: snap.data().requiresFollowUp || "",
              }
            : { meetingNote: "", requiresFollowUp: "" };

          setNotesMap(prev => {
            const n = new Map(prev);
            n.set(bc, noteData);
            return n;
          });
          
          setNotesCache(prev => {
            const n = new Map(prev);
            n.set(bc, noteData);
            return n;
          });
        },
        (err) => console.error(`Error syncing barcode ${bc}:`, err)
      );
      notesListenersRef.current.set(bc, unsub);
    });

    // Merge notes using cache for non-visible rows
    const merged = baseCombinedData.map((row) => {
      const bc = row["Barcode#"];
      const notes = notesMap.get(bc) || notesCache.get(bc) || { meetingNote: "", requiresFollowUp: "" };
      return {
        ...row,
        "Meeting Note": notes.meetingNote,
        "Requires Follow Up": notes.requiresFollowUp,
      };
    });
    setCombinedDataWithNotes(merged);

    return () => {
      notesListenersRef.current.forEach((unsub) => unsub());
      notesListenersRef.current.clear();
    };
  }, [baseCombinedData, notesMap, currentPage, notesCache]);

  // Clean up listeners when switching away from combined tab
  useEffect(() => {
    if (activeTab !== "combined") {
      notesListenersRef.current.forEach(unsub => unsub());
      notesListenersRef.current.clear();
    }
  }, [activeTab]);

  // ---------- Derived UI values ----------
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
      reportData.forEach(r => {
        const loc = r["Repair Location"];
        if (loc && String(loc).trim()) s.add(String(loc).trim());
      });
    }
    if (activeTab === "combined" || activeTab === "tickets") {
      ticketData.forEach(r => {
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

  // ---------- Category mapping helpers ----------
  const addCategoryMapping = (category, pm, department = "", categoryText = "") => {
    setCategoryMapping((prev) => {
      const copy = [...prev];
      const i = copy.findIndex((m) => m.category.trim().toUpperCase() === category.trim().toUpperCase());
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

  const importNotesFromExcel = async (file) => {
    if (!file) return;
    setIsImporting(true);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });

      const norm = (v) => String(v ?? "").trim();
      const getRow = (r) => {
        const barcode =
          norm(r["Barcode#"] || r["Barcode"] || r["BARCODE#"] || r["barcode"]);
        const meetingNote =
          r["Meeting Note"] ?? r["MeetingNote"] ?? r["Meeting_Notes"] ?? "";
        const requiresFollowUp =
          r["Requires Follow Up"] ??
          r["RequiresFollowUp"] ??
          r["Follow Up"] ??
          r["FollowUp"] ??
          "";
        return {
          barcode,
          meetingNote: String(meetingNote),
          requiresFollowUp: String(requiresFollowUp),
        };
      };

      const byBarcode = new Map();
      for (const raw of rows) {
        const { barcode, meetingNote, requiresFollowUp } = getRow(raw);
        if (!barcode) continue;
        byBarcode.set(barcode, { barcode, meetingNote, requiresFollowUp });
      }
      const deduped = Array.from(byBarcode.values());

      const chunkSize = 400;
      let written = 0;
      for (let i = 0; i < deduped.length; i += chunkSize) {
        const slice = deduped.slice(i, i + chunkSize);
        const batch = writeBatch(db);
        for (const { barcode, meetingNote, requiresFollowUp } of slice) {
          const ref = doc(db, "repairNotes", barcode);
          batch.set(
            ref,
            {
              barcode,
              meetingNote,
              requiresFollowUp,
              lastUpdated: serverTimestamp(),
            },
            { merge: true }
          );
          written++;
        }
        await batch.commit();
      }
      alert(`Imported/merged notes for ${written} unique barcodes`);
    } catch (e) {
      console.error("Notes import failed:", e);
      alert("Failed to import notes. Ensure the first sheet has columns: Barcode#, Meeting Note, Requires Follow Up.");
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
                <h3 className="font-semibold text-gray-800">
                  Current Mappings ({categoryMapping.length})
                </h3>
                <button
                  onClick={exportCategoryMapping}
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
                      {m.category_text && <p className="text-sm text-gray-600">{m.category_text}</p>}
                      {m.department && (
                        <p className="text-xs text-gray-500">Department: {m.department}</p>
                      )}
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

  // ---------- Filtering, sorting, export ----------
  const hasData = ticketData.length > 0 || reportData.length > 0;

  // Reset page when filters change
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

  const exportToCSV = () => {
    if (filteredAndSortedData.length === 0) return;
    const headers = columns.join(",");
    const rows = filteredAndSortedData.map((row) =>
      columns
        .map((col) => {
          const val = row[col];
          const text = String(val ?? "").replace(/"/g, '""');
          return `"${text}"`;
        })
        .join(",")
    );
    const csv = [headers, ...rows].join("\n");
    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${activeTab}_export_${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  // ---------- Render ----------
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
                  OneDrive/SharePoint synced
                </div>
              )}
              {lastSync && <span className="text-xs text-gray-500">Last sync: {lastSync.toLocaleTimeString()}</span>}
              {activeTab === "combined" && (
                <span className="text-xs text-blue-600 bg-blue-50 px-2 py-1 rounded">
                  Active listeners: {notesListenersRef.current.size} / {MAX_LISTENERS}
                </span>
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
                <button onClick={handleLogout} className="px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 text-sm">
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
              Upload Notes (Excel/CSV)
              <input
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={(e) => importNotesFromExcel(e.target.files?.[0])}
                className="hidden"
                disabled={loading || isImporting}
              />
            </label>
          </div>
        </div>
      </div>

      {/* Toolbar: Tabs on top, Search + Filters below */}
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

          {/* Page Size Selector */}
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
              <option key={loc} value={loc}>{loc}</option>
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
                  <option key={pm} value={pm}>{pm}</option>
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
        {/* Performance Warning for "All" */}
        {itemsPerPage >= 99999 && filteredAndSortedData.length > 500 && (
          <div className="mb-4 p-3 bg-yellow-50 border border-yellow-200 rounded-lg">
            <div className="flex items-start gap-2">
              <span className="text-yellow-600 text-sm">⚠️</span>
              <div className="flex-1">
                <p className="text-sm text-yellow-800 font-medium">Performance Note</p>
                <p className="text-xs text-yellow-700 mt-1">
                  Displaying all {filteredAndSortedData.length} rows at once. This may cause slower scrolling. 
                  Consider using pagination (100-500 rows) for better performance.
                </p>
              </div>
              <button
                onClick={() => setItemsPerPage(100)}
                className="text-xs text-yellow-800 hover:text-yellow-900 underline"
              >
                Switch to 100/page
              </button>
            </div>
          </div>
        )}
        
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
              <h2 className="text-xl font-semibold text-gray-800 mb-4">Data Matching Diagnostics</h2>
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
                  <h3 className="font-semibold text-purple-900 mb-2">Category Mappings</h3>
                  <p className="text-2xl font-bold text-purple-800">{categoryMapping.length}</p>
                  <p className="text-sm text-purple-700">categories mapped</p>
                </div>
              </div>
              
              <div className="mt-6 p-4 bg-blue-50 rounded-lg">
                <h3 className="font-semibold text-blue-900 mb-2">Memory Optimization Status</h3>
                <div className="space-y-2 text-sm">
                  <div className="flex justify-between">
                    <span>Active Firestore Listeners:</span>
                    <span className="font-mono font-bold">{notesListenersRef.current.size} / {MAX_LISTENERS}</span>
                  </div>
                  <div className="flex justify-between">
                    <span>Cached Notes:</span>
                    <span className="font-mono font-bold">{notesCache.size}</span>
                  </div>
                  <div className="flex justify-between">
                    <span>Total Rows:</span>
                    <span className="font-mono font-bold">{combinedDataWithNotes.length}</span>
                  </div>
                  <div className="flex justify-between">
                    <span>Filtered Rows:</span>
                    <span className="font-mono font-bold">{filteredAndSortedData.length}</span>
                  </div>
                </div>
              </div>
            </div>
          </div>
        ) : !hasData ? (
          <div className="flex items-center justify-center h-full">
            <div className="text-center max-w-lg bg-white p-12 rounded-lg shadow-lg">
              <FileSpreadsheet className="mx-auto text-blue-500 mb-6" size={64} />
              <h3 className="text-2xl font-semibold text-gray-800 mb-3">Welcome to Repair Tracker</h3>
              <p className="text-gray-600 mb-6">Sign in to load data from SharePoint, or upload files manually.</p>
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
          />
        )}
      </div>

      {/* Footer strip - Updated to show pagination info */}
      {hasData && activeTab !== "diagnostics" && (
        <div className="bg-white border-t px-6 py-3">
          <div className="flex items-center justify-between text-sm text-gray-600">
            <span>
              Page {currentPage} of {Math.ceil(filteredAndSortedData.length / ITEMS_PER_PAGE)} 
              {" "}({filteredAndSortedData.length} total records)
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
        <RowEditor row={editingRow} rowIndex={editingRowIndex} onClose={closeRowEditor} />
      )}
    </div>
  );
};

export default RepairTrackerSheet;