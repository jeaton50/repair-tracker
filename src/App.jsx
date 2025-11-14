import React, { useState, useMemo, useEffect, useRef, useCallback } from "react";
import { Client } from "@microsoft/microsoft-graph-client";
import { Toaster } from "react-hot-toast";
import toast from "react-hot-toast";
import * as XLSX from "xlsx";
import {
  Search,
  Download,
  Upload,
  FileSpreadsheet,
  RefreshCw,
  Cloud,
  Save,
} from "lucide-react";

// Components
import RowEditor from "./components/RowEditor";
import PaginatedTable from "./components/PaginatedTable";
import CategoryManager from "./components/CategoryManager";

// Hooks
import { useDebounce } from "./hooks/useDebounce";
import { useAuth, ensureAccessToken } from "./hooks/useAuth";

// Services
import OneDriveService from "./oneDriveService.js";
import SharePointNotesService from "./SharePointNotesService.js";

// Config & Utils
import { graphConfig } from "./authConfig";
import {
  DEBOUNCE_DELAY_MS,
  NOTES_REFRESH_INTERVAL_MS,
  DATA_AUTO_REFRESH_INTERVAL_MS,
  AUTO_SAVE_DELAY_MS,
  DEFAULT_ITEMS_PER_PAGE,
  ITEMS_PER_PAGE_OPTIONS,
} from "./utils/constants";
import { normalizeBarcode, ageInDays, formatTicketNumber } from "./utils/dataHelpers";

/**
 * Main Repair Tracker Dashboard Application
 * Manages equipment repair tracking with SharePoint/OneDrive integration
 */
const RepairTrackerSheet = () => {
  // Authentication
  const { isAuthenticated, accessToken, userName, handleLogin, handleLogout } = useAuth();

  // Tab and Search State
  const [activeTab, setActiveTab] = useState("combined");
  const [searchInput, setSearchInput] = useState("");
  const [searchTerm, setSearchTerm] = useState("");
  const [sortConfig, setSortConfig] = useState({ key: null, direction: "asc" });

  // Data State
  const [ticketData, setTicketData] = useState([]);
  const [reportData, setReportData] = useState([]);
  const [categoryMapping, setCategoryMapping] = useState([]);
  const [combinedDataWithNotes, setCombinedDataWithNotes] = useState([]);
  const [notesMap, setNotesMap] = useState(new Map());

  // UI State
  const [showCategoryManager, setShowCategoryManager] = useState(false);
  const [unmatchedCategories, setUnmatchedCategories] = useState([]);
  const [editingRow, setEditingRow] = useState(null);
  const [editingRowIndex, setEditingRowIndex] = useState(null);
  const [loading, setLoading] = useState(false);
  const [isImporting, setIsImporting] = useState(false);

  // Service State
  const [notesService, setNotesService] = useState(null);
  const [lastSync, setLastSync] = useState(null);
  const [lastNotesSync, setLastNotesSync] = useState(null);

  // Pagination and Filters
  const [currentPage, setCurrentPage] = useState(1);
  const [itemsPerPage, setItemsPerPage] = useState(DEFAULT_ITEMS_PER_PAGE);
  const [locationFilter, setLocationFilter] = useState("");
  const [pmFilter, setPmFilter] = useState("");

  // Quick Edit State
  const [pendingSaves, setPendingSaves] = useState(0);
  const saveTimer = useRef(null);

  const debouncedSetSearch = useDebounce((value) => setSearchTerm(value), DEBOUNCE_DELAY_MS);

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

        console.log("📥 Loading notes from SharePoint...");
        const notes = await service.loadAllNotes();
        setNotesMap(notes);
        setLastNotesSync(new Date());
        toast.success(`Loaded ${notes.size} notes from SharePoint`);
      } catch (error) {
        console.error("Failed to initialize notes service:", error);
        toast.error("Failed to load notes from SharePoint");
      }
    };

    initNotesService();
  }, [isAuthenticated, accessToken]);

  // Periodic refresh of notes
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
    }, NOTES_REFRESH_INTERVAL_MS);

    return () => clearInterval(interval);
  }, [notesService, isAuthenticated]);

  // SharePoint data loader
  const loadFromSharePoint = useCallback(
    async (silent = false) => {
      if (!silent) setLoading(true);

      const loadPromise = (async () => {
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

        return { tickets, reports, mapping };
      })();

      if (!silent) {
        toast.promise(loadPromise, {
          loading: "Loading from SharePoint...",
          success: (data) =>
            `Loaded ${data.tickets.length} tickets, ${data.reports.length} reports`,
          error: "Failed to load from SharePoint",
        });
      }

      try {
        await loadPromise;
      } catch (e) {
        console.error("SharePoint load failed:", e);
      } finally {
        if (!silent) setLoading(false);
      }
    },
    []
  );

  // Load data on authentication
  useEffect(() => {
    if (!isAuthenticated || !accessToken) return;
    loadFromSharePoint();
  }, [isAuthenticated, accessToken, loadFromSharePoint]);

  // Auto refresh data every 5 minutes
  useEffect(() => {
    if (!isAuthenticated || !accessToken) return;
    const interval = setInterval(
      () => loadFromSharePoint(true),
      DATA_AUTO_REFRESH_INTERVAL_MS
    );
    return () => clearInterval(interval);
  }, [isAuthenticated, accessToken, loadFromSharePoint]);

  // Build combined data with category mappings
  const baseCombinedData = useMemo(() => {
    if (reportData.length === 0) return [];

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

  // Merge notes with combined data
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

  // Get current data based on active tab
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
    currentData.length > 0
      ? Object.keys(currentData[0]).filter((c) => !c.startsWith("_"))
      : [];

  // Unique locations for filter
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

  // Unique categories and PMs
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
    toast.success("Category mapping exported");
  };

  // Download notes template
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
    toast.success("Template downloaded");
  };

  // Import notes from Excel
  const importNotesFromExcel = async (file) => {
    if (!file || !notesService) return;

    setIsImporting(true);
    const importPromise = (async () => {
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

      const notes = await notesService.loadAllNotes();
      setNotesMap(notes);
      setLastNotesSync(new Date());

      return notesArray.length;
    })();

    toast.promise(importPromise, {
      loading: "Importing notes...",
      success: (count) => `Imported ${count} notes to SharePoint`,
      error: "Failed to import notes",
    });

    try {
      await importPromise;
    } catch (error) {
      console.error("Import failed:", error);
    } finally {
      setIsImporting(false);
    }
  };

  // Reset page when filters change
  useEffect(() => {
    setCurrentPage(1);
  }, [searchTerm, locationFilter, pmFilter, activeTab]);

  // Filter and sort data
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
        if (pmFilter === "__unassigned__")
          return !r["Assigned To"] || r["Assigned To"] === "";
        return r["Assigned To"] === pmFilter;
      });
    }

    if (searchTerm) {
      const q = searchTerm.toLowerCase();
      rows = rows.filter((r) =>
        Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q))
      );
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
    if (notesService) {
      const notes = await notesService.loadAllNotes();
      setNotesMap(notes);
      setLastNotesSync(new Date());
    }
  };

  // Export to CSV
  const exportToCSV = () => {
    const rows = filteredAndSortedData;
    if (rows.length === 0) {
      toast.error("No data to export");
      return;
    }

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
    toast.success("Data exported successfully");
  };

  // Quick Edit save orchestration
  const queueSaveAndAutoFlush = useCallback(() => {
    setPendingSaves((n) => n + 1);
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(async () => {
      try {
        await notesService?.saveToSharePoint();
        setPendingSaves(0);
        setLastNotesSync(new Date());
        toast.success("Changes saved to SharePoint");
      } catch (e) {
        console.error("Save error:", e);
        toast.error("Failed to save changes");
      }
    }, AUTO_SAVE_DELAY_MS);
  }, [notesService]);

  const saveNow = useCallback(async () => {
    clearTimeout(saveTimer.current);

    const savePromise = (async () => {
      await notesService?.saveToSharePoint();
      setPendingSaves(0);
      setLastNotesSync(new Date());
    })();

    toast.promise(savePromise, {
      loading: "Saving to SharePoint...",
      success: "Saved successfully!",
      error: "Failed to save",
    });

    try {
      await savePromise;
    } catch (e) {
      console.error("Save error:", e);
    }
  }, [notesService]);

  // Inline edit handlers
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

  const hasDataNow = ticketData.length > 0 || reportData.length > 0;

  return (
    <div className="w-full h-screen flex flex-col bg-gray-50">
      <Toaster position="top-right" />

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
              {lastSync && (
                <span className="text-xs text-gray-500">
                  Data: {lastSync.toLocaleTimeString()}
                </span>
              )}
              {lastNotesSync && (
                <span className="text-xs text-blue-500">
                  Notes: {lastNotesSync.toLocaleTimeString()}
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
                {pendingSaves > 0 ? (
                  <div className="text-amber-600 text-sm">
                    💾 {pendingSaves} pending change(s)
                  </div>
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
                    toast.success(`Loaded ${json.length} category mappings`);
                  } catch {
                    toast.error("Invalid JSON file");
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
                activeTab === "combined"
                  ? "border-blue-500 text-blue-600"
                  : "border-transparent text-gray-500"
              }`}
            >
              Combined ({combinedDataWithNotes.length})
            </button>
            <button
              onClick={() => setActiveTab("tickets")}
              className={`px-4 py-2 font-medium border-b-2 transition-colors ${
                activeTab === "tickets"
                  ? "border-blue-500 text-blue-600"
                  : "border-transparent text-gray-500"
              }`}
            >
              Tickets ({ticketData.length})
            </button>
            <button
              onClick={() => setActiveTab("reports")}
              className={`px-4 py-2 font-medium border-b-2 transition-colors ${
                activeTab === "reports"
                  ? "border-blue-500 text-blue-600"
                  : "border-transparent text-gray-500"
              }`}
            >
              Reports ({reportData.length})
            </button>
            <button
              onClick={() => setActiveTab("diagnostics")}
              className={`px-4 py-2 font-medium border-b-2 transition-colors ${
                activeTab === "diagnostics"
                  ? "border-orange-500 text-orange-600"
                  : "border-transparent text-gray-500"
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
            <Search
              className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400"
              size={18}
            />
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
            <label className="text-sm text-gray-600 whitespace-nowrap">
              Rows per page:
            </label>
            <select
              value={itemsPerPage}
              onChange={(e) => {
                const newSize = parseInt(e.target.value);
                setItemsPerPage(newSize);
                setCurrentPage(1);
              }}
              className="px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white text-sm"
            >
              {ITEMS_PER_PAGE_OPTIONS.map((size) => (
                <option key={size} value={size}>
                  {size === 99999 ? "All" : size}
                </option>
              ))}
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
              <h2 className="text-xl font-semibold text-gray-800 mb-4">
                System Diagnostics
              </h2>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
                <div className="p-4 bg-blue-50 rounded-lg">
                  <h3 className="font-semibold text-blue-900 mb-2">
                    Repair Ticket List
                  </h3>
                  <p className="text-2xl font-bold text-blue-800">{ticketData.length}</p>
                  <p className="text-sm text-blue-700">records</p>
                </div>
                <div className="p-4 bg-green-50 rounded-lg">
                  <h3 className="font-semibold text-green-900 mb-2">Repair Report</h3>
                  <p className="text-2xl font-bold text-green-800">{reportData.length}</p>
                  <p className="text-sm text-green-700">records</p>
                </div>
                <div className="p-4 bg-purple-50 rounded-lg">
                  <h3 className="font-semibold text-purple-900 mb-2">
                    Notes (SharePoint)
                  </h3>
                  <p className="text-2xl font-bold text-purple-800">{notesMap.size}</p>
                  <p className="text-sm text-purple-700">notes stored</p>
                </div>
              </div>

              <div className="mt-6 p-4 bg-green-50 rounded-lg border border-green-200">
                <h3 className="font-semibold text-green-900 mb-3">
                  💰 SharePoint Storage Benefits
                </h3>
                <div className="space-y-2 text-sm text-green-800">
                  <div className="flex items-start gap-2">
                    <span className="text-green-600">✓</span>
                    <div>
                      <strong>All data in SharePoint</strong>
                      <p className="text-xs text-green-700">
                        Centralized storage with enterprise security
                      </p>
                    </div>
                  </div>
                  <div className="flex items-start gap-2">
                    <span className="text-green-600">✓</span>
                    <div>
                      <strong>Auto-save with manual control</strong>
                      <p className="text-xs text-green-700">
                        Changes queued and saved automatically after 10 seconds
                      </p>
                    </div>
                  </div>
                  <div className="flex items-start gap-2">
                    <span className="text-green-600">✓</span>
                    <div>
                      <strong>30-second refresh cycle</strong>
                      <p className="text-xs text-green-700">
                        Automatically syncs with SharePoint
                      </p>
                    </div>
                  </div>
                  <div className="flex items-start gap-2">
                    <span className="text-green-600">✓</span>
                    <div>
                      <strong>No external dependencies</strong>
                      <p className="text-xs text-green-700">
                        Removed Firebase - reduced bundle size by 300KB
                      </p>
                    </div>
                  </div>
                </div>
              </div>
            </div>
          </div>
        ) : !hasDataNow ? (
          <div className="flex items-center justify-center h-full">
            <div className="text-center max-w-lg bg-white p-12 rounded-lg shadow-lg">
              <FileSpreadsheet className="mx-auto text-blue-500 mb-6" size={64} />
              <h3 className="text-2xl font-semibold text-gray-800 mb-3">
                Welcome to Repair Tracker
              </h3>
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
            itemsPerPage={itemsPerPage}
            sortConfig={sortConfig}
            onSort={handleSort}
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
                  Page {currentPage} of{" "}
                  {Math.ceil(filteredAndSortedData.length / itemsPerPage)} (
                  {filteredAndSortedData.length} total records)
                </>
              )}
            </span>
            <div className="flex items-center gap-4">
              {locationFilter && (
                <span className="text-blue-600">Location: {locationFilter}</span>
              )}
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

      {/* Modals */}
      {showCategoryManager && (
        <CategoryManager
          categoryMapping={categoryMapping}
          allCategories={allCategories}
          unmatchedCategories={unmatchedCategories}
          onAddMapping={addCategoryMapping}
          onRemoveMapping={removeCategoryMapping}
          onExport={exportCategoryMapping}
          onClose={() => setShowCategoryManager(false)}
        />
      )}

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
