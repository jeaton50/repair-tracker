// src/App.jsx
import React, { useEffect, useMemo, useRef, useState } from "react";
import { Client } from "@microsoft/microsoft-graph-client";
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig, loginRequest, graphConfig } from "./authConfig";
import OneDriveService from "./oneDriveService";
import {
  Search, Download, ChevronDown, ChevronUp,
  Upload, FileSpreadsheet, RefreshCw, Cloud
} from "lucide-react";

import { db } from "./firebase";
import { doc, setDoc, onSnapshot, serverTimestamp, deleteDoc } from "firebase/firestore";

const msalInstance = new PublicClientApplication(msalConfig);

/* ----------------------------- Auth helper ----------------------------- */
async function ensureAccessToken() {
  // Ensure MSAL is initialized before any API calls
  if (!msalInstance || !msalInstance.initialize) {
    throw new Error("MSAL instance not available");
  }
  await msalInstance.initialize();

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

/* ----------------------------- Row Editor ------------------------------ */
const RowEditor = ({ row, rowIndex, onClose }) => {
  const [meetingNote, setMeetingNote] = useState(row["Meeting Note"] || "");
  const [followUp, setFollowUp] = useState(row["Requires Follow Up"] || "");
  const [lastSaved, setLastSaved] = useState(null);
  const [isSaving, setIsSaving] = useState(false);
  const saveTimeoutRef = useRef(null);
  const [locationFilter, setLocationFilter] = useState("");
  const [pmFilter, setPmFilter] = useState("");
  
  const barcode = row["Barcode#"];

  // Auto-save to Firebase
  useEffect(() => {
    if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current);
    saveTimeoutRef.current = setTimeout(async () => {
      if (!barcode) return;
      setIsSaving(true);
      try {
        const docRef = doc(db, "repairNotes", barcode);
        await setDoc(
          docRef,
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
  }, [meetingNote, followUp, barcode]);

  // Realtime sync from Firebase
  useEffect(() => {
    if (!barcode) return;
    const docRef = doc(db, "repairNotes", barcode);
    const unsub = onSnapshot(
      docRef,
      (snap) => {
        if (snap.exists()) {
          const data = snap.data();
          if (document.activeElement?.name !== "meetingNote") setMeetingNote(data.meetingNote || "");
          if (document.activeElement?.name !== "followUp") setFollowUp(data.requiresFollowUp || "");
        }
      },
      (err) => console.error("Realtime sync error:", err)
    );
    return () => unsub();
  }, [barcode]);

  const handleClearAll = () => {
    if (window.confirm("Clear all notes for this item?")) {
      setMeetingNote("");
      setFollowUp("");
    }
  };

  const handleDeleteFromDatabase = async () => {
    if (!window.confirm("Permanently delete all notes for this item from the database?")) return;
    try {
      const docRef = doc(db, "repairNotes", barcode);
      await deleteDoc(docRef);
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
            {isSaving && (
              <p className="text-xs text-blue-600 mt-1 flex items-center gap-1">
                <span className="inline-block animate-spin">⟳</span> Saving...
              </p>
            )}
            {lastSaved && !isSaving && (
              <p className="text-xs text-green-600 mt-1">✓ Saved at {lastSaved.toLocaleTimeString()}</p>
            )}
          </div>
          <button onClick={onClose} className="text-gray-500 hover:text-gray-700 text-2xl">×</button>
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
            />
          </div>

          <div className="p-4 bg-gray-50 rounded-lg text-sm space-y-2">
            <div><span className="font-semibold text-gray-700">Damage:</span> {row["Damage Description"]}</div>
            <div><span className="font-semibold text-gray-700">Ticket Notes:</span> {row["Ticket Description"]}</div>
            <div><span className="font-semibold text-gray-700">Reason:</span> {row["Repair Reason"]}</div>
          </div>
        </div>

        <div className="p-6 border-t flex justify-between">
          <div className="flex gap-2">
            <button
              onClick={handleClearAll}
              className="px-6 py-2 text-orange-600 bg-orange-50 border border-orange-300 rounded-lg hover:bg-orange-100"
            >
              Clear All Notes
            </button>
            <button
              onClick={handleDeleteFromDatabase}
              className="px-6 py-2 text-red-600 bg-red-50 border border-red-300 rounded-lg hover:bg-red-100"
            >
              Delete from Database
            </button>
          </div>
          <button
            onClick={onClose}
            className="px-6 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700"
          >
            Close
          </button>
        </div>
      </div>
    </div>
  );
};

/* ------------------------------ Main App ------------------------------- */
const App = () => {
  // MSAL ready gate
  const [msalReady, setMsalReady] = useState(false);

  // Core state
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [accessToken, setAccessToken] = useState(null);
  const [userName, setUserName] = useState("");
  const [loading, setLoading] = useState(false);
  const [lastSync, setLastSync] = useState(null);
  const [autoRefresh, setAutoRefresh] = useState(true);

  // Data sets
  const [ticketData, setTicketData] = useState([]);
  const [reportData, setReportData] = useState([]);
  const [categoryMapping, setCategoryMapping] = useState([]);

  // Notes sync
  const [combinedDataWithNotes, setCombinedDataWithNotes] = useState([]);
  const [notesMap, setNotesMap] = useState(new Map());
  const notesListenersRef = useRef(new Map());

  // UI state
  const [activeTab, setActiveTab] = useState("combined");
  const [searchTerm, setSearchTerm] = useState("");
  const [sortConfig, setSortConfig] = useState({ key: null, direction: "asc" });
  const [showCategoryManager, setShowCategoryManager] = useState(false);
  const [unmatchedCategories, setUnmatchedCategories] = useState([]);
  const [editingRow, setEditingRow] = useState(null);
  const [editingRowIndex, setEditingRowIndex] = useState(null);

  /* ----------------------- Initialize MSAL once ------------------------ */
  useEffect(() => {
    let mounted = true;
    (async () => {
      try {
        await msalInstance.initialize();
        if (!mounted) return;

        setMsalReady(true);

        // Try restoring a session silently
        const account = msalInstance.getAllAccounts()[0];
        if (!account) return;

        const r = await msalInstance.acquireTokenSilent({ ...loginRequest, account });
        setAccessToken(r.accessToken);
        setIsAuthenticated(true);
        setUserName(account.name || "");
        await loadFromSharePoint(true);
      } catch (e) {
        // Not signed in or init failed; UI will allow login
        console.debug("MSAL init/restore:", e?.message || e);
        setMsalReady(true);
      }
    })();
    return () => { mounted = false; };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  /* ----------------------- Helpers / formatters ------------------------ */
  const formatCell = (value) => {
    if (value == null) return "";
    const str = String(value);
    if (str.includes("T") && str.includes("Z")) {
      const d = new Date(str);
      if (!isNaN(d.getTime())) return d.toLocaleDateString();
    }
    const num = parseFloat(str.replace(/,/g, ""));
    if (!isNaN(num) && /^[\d,.\s-]+$/.test(str)) {
      return Number.isInteger(num) ? String(Math.trunc(num)) : String(num);
    }
    return str;
  };

  const calculateAge = (dateStr) => {
    if (!dateStr) return "";
    const d = new Date(dateStr);
    if (isNaN(d.getTime())) return "";
    const days = Math.round((Date.now() - d.getTime()) / (1000 * 60 * 60 * 24));
    return days >= 0 ? days : "";
  };

  /* ---------------------- Build combined base data --------------------- */
  const baseCombinedData = useMemo(() => {
    if (reportData.length === 0) return [];

    const norm = (s) => (s ? String(s).trim().toUpperCase() : "");
    const ticketMap = new Map();
    ticketData.forEach((t) => {
      const bc = norm(t["Barcode"]);
      if (bc) ticketMap.set(bc, t);
    });

    const categoryToPM = new Map();
    categoryMapping.forEach((m) => {
      const cat = m.category?.trim();
      if (cat) {
        categoryToPM.set(cat.toUpperCase(), {
          pm: m.pm || "",
          department: m.department || "",
          category_text: m.category_text || "",
        });
      }
    });

    const unmatched = new Set();

    const combined = reportData.map((r) => {
      const bc = norm(r["Barcode#"]);
      const ticket = ticketMap.get(bc) || {};
      const cat = (r["Category"] || "").trim();
      const mapping = categoryToPM.get(cat.toUpperCase());
      const assignedPM = mapping?.pm || "";

      if (cat && !assignedPM) unmatched.add(cat);

      return {
        "Meeting Note": "",
        "Requires Follow Up": "",
        "Assigned To": assignedPM,
        "Location": r["Repair Location"] || ticket["Location"] || "",
        "Repair Ticket": r["Ticket"] || "",
        "Asset Repair Age": calculateAge(r["Date In"]),
        "Barcode#": r["Barcode#"] || ticket["Barcode"] || "",
        "Equipment": `(${r["Equipment"] || ""}) - ${r["Description"] || ""}`,
        "Damage Description": r["Notes"] || "",
        "Ticket Description": ticket["Notes"] || "",
        "Repair Reason": r["Repair Reason"] || "",
        "Last Order#": r["Last Order#"] || ticket["Order# to Bill"] || "",
        "Reference#": r["Reference#"] || "",
        "Customer": r["Customer"] || ticket["Customer"] || "",
        "Customer Title": r["Customer Title"] || "",
        "Repair Cost": r["Repair Cost"] || "0",
        "Date In": r["Date In"] || ticket["Creation Date"] || "",
        "Department": r["Department"] || "",
        "Category": r["Category"] || "",
        "Billable": ticket["Billable"] || r["Billable"] || "",
        "Created By": ticket["Created By"] || r["User In"] || "",
        "Repair Price": r["Repair Price"] || "0",
        "Repair Vendor": r["Repair Vendor"] || "",
        "_TicketMatched": Object.keys(ticket).length > 0 ? "Yes" : "No",
      };
    });

    setUnmatchedCategories(Array.from(unmatched).sort());
    return combined;
  }, [ticketData, reportData, categoryMapping]);

  /* ---------------------- Live notes merge (Firebase) ------------------- */
  useEffect(() => {
    if (baseCombinedData.length === 0) {
      setCombinedDataWithNotes([]);
      return;
    }

    const barcodes = baseCombinedData.map((row) => row["Barcode#"]).filter(Boolean);

    // remove stale listeners
    notesListenersRef.current.forEach((unsub, bc) => {
      if (!barcodes.includes(bc)) {
        unsub();
        notesListenersRef.current.delete(bc);
      }
    });

    // add listeners for new barcodes
    barcodes.forEach((bc) => {
      if (notesListenersRef.current.has(bc)) return;
      const docRef = doc(db, "repairNotes", bc);
      const unsub = onSnapshot(
        docRef,
        (snap) => {
          setNotesMap((prev) => {
            const next = new Map(prev);
            if (snap.exists()) {
              const d = snap.data();
              next.set(bc, {
                meetingNote: d.meetingNote || "",
                requiresFollowUp: d.requiresFollowUp || "",
              });
            } else {
              next.set(bc, { meetingNote: "", requiresFollowUp: "" });
            }
            return next;
          });
        },
        (e) => console.error(`Error syncing barcode ${bc}:`, e)
      );
      notesListenersRef.current.set(bc, unsub);
    });

    // merge
    const merged = baseCombinedData.map((row) => {
      const bc = row["Barcode#"];
      const n = notesMap.get(bc);
      return {
        ...row,
        "Meeting Note": n?.meetingNote || "",
        "Requires Follow Up": n?.requiresFollowUp || "",
      };
    });
    setCombinedDataWithNotes(merged);

    return () => {
      notesListenersRef.current.forEach((unsub) => unsub());
      notesListenersRef.current.clear();
    };
  }, [baseCombinedData, notesMap]);

  /* ------------------------ Category Manager helpers -------------------- */
/* ------------------------ Category Manager helpers -------------------- */
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
    prev.filter(
      (m) => m.category.trim().toUpperCase() !== category.trim().toUpperCase()
    )
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

/* ------------------------ Category Manager UI ------------------------- */
const CategoryManager = () => {
  const [newCategory, setNewCategory] = useState("");
  const [newPM, setNewPM] = useState("");
  const [newDepartment, setNewDepartment] = useState("");
  const [newCategoryText, setNewCategoryText] = useState("");

  return (
    <div className="fixed inset-0 bg-black/50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-lg shadow-xl max-w-6xl w-full max-h-[90vh] overflow-hidden flex flex-col">
        <div className="p-6 border-b flex justify-between items-center">
          <h2 className="text-2xl font-bold text-gray-800">
            Category to PM Mapping Manager
          </h2>
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


  /* ------------------------ SharePoint data loader ---------------------- */
  const loadFromSharePoint = async (silent = false) => {
    if (!silent) setLoading(true);
    try {
      const token = await ensureAccessToken();
      const graph = Client.init({ authProvider: (done) => done(null, token) });
      const ods = new OneDriveService(graph);

      const HOST = graphConfig.spHostname;
      const SITE = graphConfig.spSitePath;
      const BASE = graphConfig.spBasePath;

      const [tickets, reports, mapping] = await Promise.all([
        ods.readExcelFromSharePoint({
          hostname: HOST, sitePath: SITE,
          fileRelativePath: `${BASE}/${graphConfig.ticketsFile}`,
        }).catch(() => []),
        ods.readExcelFromSharePoint({
          hostname: HOST, sitePath: SITE,
          fileRelativePath: `${BASE}/${graphConfig.reportsFile}`,
        }).catch(() => []),
        ods.readJsonFromSharePoint({
          hostname: HOST, sitePath: SITE,
          fileRelativePath: `${BASE}/${graphConfig.mappingFile}`,
        }).catch(() => []),
      ]);

      setTicketData(tickets);
      setReportData(reports);
      setCategoryMapping(mapping);
      setLastSync(new Date());

      if (!silent) {
        alert(`Loaded from SharePoint:\n${tickets.length} tickets\n${reports.length} reports\n${mapping.length} category mappings`);
      }
    } catch (e) {
      console.error("SharePoint load failed:", e);
      if (!silent) alert(e.message || "Failed to load from SharePoint.");
    } finally {
      if (!silent) setLoading(false);
    }
  };

  /* ------------------------------ Auth handlers ------------------------ */
  const handleLogin = async () => {
    try {
      // Make sure MSAL is initialized before showing popup
      await msalInstance.initialize();
      await msalInstance.loginPopup(loginRequest);
      const account = msalInstance.getAllAccounts()[0];
      const { accessToken } = await msalInstance.acquireTokenSilent({ ...loginRequest, account });
      setAccessToken(accessToken);
      setIsAuthenticated(true);
      setUserName(account.name);
      await loadFromSharePoint(false);
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
    setCombinedDataWithNotes([]);
    setNotesMap(new Map());
    notesListenersRef.current.forEach((u) => u());
    notesListenersRef.current.clear();
  };

  /* -------------------------------- Effects ---------------------------- */
  // Auto-refresh every 5 minutes
  useEffect(() => {
    if (!autoRefresh || !isAuthenticated || !accessToken) return;
    const interval = setInterval(() => loadFromSharePoint(true), 5 * 60 * 1000);
    return () => clearInterval(interval);
  }, [autoRefresh, isAuthenticated, accessToken]);

  /* ------------------------------ Derived UI --------------------------- */
  const hasData = ticketData.length > 0 || reportData.length > 0;

  const getCurrentData = () => {
    switch (activeTab) {
      case "tickets": return ticketData;
      case "reports": return reportData;
      case "combined": return combinedDataWithNotes;
      default: return [];
    }
  };
  const currentData = getCurrentData();

  const columns = currentData.length > 0
    ? Object.keys(currentData[0]).filter((c) => !c.startsWith("_"))
    : [];

  const filteredAndSortedData = useMemo(() => {
  let rows = currentData;

  // 1) Location filter (matches "Location" or "Repair Location")
  if (locationFilter) {
    rows = rows.filter((r) => {
      const loc = r["Location"] || r["Repair Location"];
      return String(loc || "").trim() === locationFilter;
    });
  }

  // 2) PM filter (useful on Combined tab; harmless elsewhere)
  if (pmFilter) {
    rows = rows.filter((r) => {
      const assigned = String(r["Assigned To"] || "").trim();
      return pmFilter === "__unassigned__" ? assigned === "" : assigned === pmFilter;
    });
  }

  // 3) Search
  if (searchTerm) {
    const q = searchTerm.toLowerCase();
    rows = rows.filter((r) =>
      Object.values(r).some((v) => String(v ?? "").toLowerCase().includes(q))
    );
  }

  // 4) Sort
  if (sortConfig.key) {
    rows = [...rows].sort((a, b) => {
      const av = a[sortConfig.key];
      const bv = b[sortConfig.key];
      if (av === bv) return 0;
      return (av > bv ? 1 : -1) * (sortConfig.direction === "asc" ? 1 : -1);
    });
  }

  return rows;
}, [currentData, locationFilter, pmFilter, searchTerm, sortConfig]);

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

  /* -------------------------------- RENDER ----------------------------- */
  return (
    <div className="w-full min-h-screen flex flex-col bg-gray-50">
      {/* Header */}
      <div className="bg-white border-b px-6 py-4">
        <div className="flex items-center justify-between">
          <div>
            <h1 className="text-2xl font-bold text-gray-800">Repair Tracker Dashboard</h1>
            <div className="flex items-center gap-3 mt-1">
              <p className="text-sm text-gray-500">
                {!msalReady
                  ? "Initializing…"
                  : isAuthenticated
                  ? `Connected as ${userName}`
                  : "Not connected"}
              </p>
              {isAuthenticated && (
                <div className="flex items-center gap-1 text-xs text-green-600 bg-green-50 px-2 py-1 rounded">
                  <Cloud size={12} /> SharePoint synced
                </div>
              )}
              {lastSync && (
                <span className="text-xs text-gray-500">Last sync: {lastSync.toLocaleTimeString()}</span>
              )}
            </div>
          </div>

          <div className="flex gap-2 items-center flex-wrap">
            {!isAuthenticated ? (
              <button
                onClick={handleLogin}
                disabled={!msalReady}
                className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 text-sm disabled:opacity-50"
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
                  onClick={handleLogout}
                  className="px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 text-sm"
                >
                  Sign Out
                </button>
                <label className="flex items-center gap-2 text-sm ml-2">
                  <input
                    type="checkbox"
                    checked={autoRefresh}
                    onChange={(e) => setAutoRefresh(e.target.checked)}
                  />
                  Auto refresh
                </label>
              </>
            )}

            {/* Admin: upload mapping/tickets/reports (optional manual loads) */}
            <label className="flex items-center gap-2 px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700 cursor-pointer transition-colors text-sm">
              <Upload size={16} />
              Upload Mapping
              <input
                type="file"
                accept=".json"
                className="hidden"
                onChange={async (e) => {
                  const f = e.target.files?.[0];
                  if (!f) return;
                  try {
                    const txt = await f.text();
                    const json = JSON.parse(txt);
                    setCategoryMapping(json);
                    alert(`Loaded ${json.length} category mappings from ${f.name}`);
                  } catch (err) {
                    alert("Invalid JSON file");
                  }
                }}
              />
            </label>
            <label className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 cursor-pointer transition-colors text-sm">
              <Upload size={16} />
              Upload Tickets
              <input
                type="file"
                accept=".xlsx,.xls"
                className="hidden"
                onChange={async (e) => {
                  const f = e.target.files?.[0];
                  if (!f) return;
                  setLoading(true);
                  try {
                    if (typeof window.XLSX === "undefined") {
                      await new Promise((res, rej) => {
                        const s = document.createElement("script");
                        s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
                        s.onload = res; s.onerror = rej; document.head.appendChild(s);
                      });
                    }
                    const buf = await f.arrayBuffer();
                    const wb = window.XLSX.read(buf, { type: "array", cellDates: true });
                    const sheet = wb.Sheets[wb.SheetNames[0]];
                    const raw = window.XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
                    if (raw.length < 2) throw new Error("Empty or incorrect format");
                    const headers = (raw[1] || []).map((h) => String(h || "").trim()).filter(Boolean);
                    const rows = raw.slice(2)
                      .filter((r) => r && r.some((c) => c !== "" && c != null))
                      .map((r) => Object.fromEntries(headers.map((h, i) => [h, r[i] ?? ""])));
                    setTicketData(rows);
                    alert(`Loaded ${rows.length} ticket rows from ${f.name}`);
                  } catch (err) {
                    console.error(err);
                    alert("Failed to parse Excel");
                  } finally {
                    setLoading(false);
                  }
                }}
              />
            </label>
            <label className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 cursor-pointer transition-colors text-sm">
              <Upload size={16} />
              Upload Reports
              <input
                type="file"
                accept=".xlsx,.xls"
                className="hidden"
                onChange={async (e) => {
                  const f = e.target.files?.[0];
                  if (!f) return;
                  setLoading(true);
                  try {
                    if (typeof window.XLSX === "undefined") {
                      await new Promise((res, rej) => {
                        const s = document.createElement("script");
                        s.src = "https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js";
                        s.onload = res; s.onerror = rej; document.head.appendChild(s);
                      });
                    }
                    const buf = await f.arrayBuffer();
                    const wb = window.XLSX.read(buf, { type: "array", cellDates: true });
                    const sheet = wb.Sheets[wb.SheetNames[0]];
                    const raw = window.XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
                    if (raw.length < 2) throw new Error("Empty or incorrect format");
                    const headers = (raw[1] || []).map((h) => String(h || "").trim()).filter(Boolean);
                    const rows = raw.slice(2)
                      .filter((r) => r && r.some((c) => c !== "" && c != null))
                      .map((r) => Object.fromEntries(headers.map((h, i) => [h, r[i] ?? ""])));
                    setReportData(rows);
                    alert(`Loaded ${rows.length} report rows from ${f.name}`);
                  } catch (err) {
                    console.error(err);
                    alert("Failed to parse Excel");
                  } finally {
                    setLoading(false);
                  }
                }}
              />
            </label>

            {hasData && (
              <button
                onClick={() => setShowCategoryManager(true)}
                className="flex items-center gap-2 px-4 py-2 bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 cursor-pointer transition-colors text-sm"
              >
                Manage Categories
                {unmatchedCategories.length > 0 && (
                  <span className="bg-red-500 text-white px-2 py-0.5 rounded-full text-xs">
                    {unmatchedCategories.length}
                  </span>
                )}
              </button>
            )}
          </div>
        </div>
      </div>

      {/* Search / Tabs */}
      {hasData && (
        <div className="bg-white border-b px-6 py-3">
          <div className="flex items-center justify-between mb-3">
            <div className="flex-1 relative">
              <Search className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
              <input
                type="text"
                placeholder="Search across all columns..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full pl-10 pr-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
              />
            </div>

            <div className="flex ml-4">
              <button
                onClick={() => setActiveTab("combined")}
                className={`px-6 py-3 font-medium border-b-2 transition-colors ${
                  activeTab === "combined" ? "border-blue-500 text-blue-600" : "border-transparent text-gray-500 hover:text-gray-700"
                }`}
              >
                Combined <span className="ml-2 text-xs bg-gray-100 px-2 py-1 rounded-full">{combinedDataWithNotes.length}</span>
              </button>
              <button
                onClick={() => setActiveTab("tickets")}
                className={`px-6 py-3 font-medium border-b-2 transition-colors ${
                  activeTab === "tickets" ? "border-blue-500 text-blue-600" : "border-transparent text-gray-500 hover:text-gray-700"
                }`}
              >
                Tickets <span className="ml-2 text-xs bg-gray-100 px-2 py-1 rounded-full">{ticketData.length}</span>
              </button>
              <button
                onClick={() => setActiveTab("reports")}
                className={`px-6 py-3 font-medium border-b-2 transition-colors ${
                  activeTab === "reports" ? "border-blue-500 text-blue-600" : "border-transparent text-gray-500 hover:text-gray-700"
                }`}
              >
                Reports <span className="ml-2 text-xs bg-gray-100 px-2 py-1 rounded-full">{reportData.length}</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Body */}
      <div className="flex-1 overflow-hidden px-6 py-4">
        {loading ? (
          <div className="flex items-center justify-center h-full">
            <div className="text-center">
              <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
              <p className="text-gray-500">Loading...</p>
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
        ) : currentData.length === 0 ? (
          <div className="flex items-center justify-center h-full">
            <div className="text-center bg-white p-8 rounded-lg shadow">
              <p className="text-gray-500">No data available</p>
            </div>
          </div>
        ) : (
          <div className="bg-white rounded-lg shadow h-full overflow-auto">
            <table className="w-full border-collapse">
              <thead className="bg-gray-50 border-b border-gray-200 sticky top-0 z-10">
                <tr>
                  {columns.map((col) => (
                    <th
                      key={col}
                      onClick={() => handleSort(col)}
                      className="px-4 py-3 text-left text-xs font-medium text-gray-700 uppercase tracking-wider cursor-pointer hover:bg-gray-100 bg-gray-50 whitespace-normal"
                    >
                      <div className="flex items-center gap-2">
                        {col}
                        {sortConfig.key === col &&
                          (sortConfig.direction === "asc" ? <ChevronUp size={14} /> : <ChevronDown size={14} />)}
                      </div>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody className="bg-white divide-y divide-gray-200">
                {filteredAndSortedData.map((row, idx) => (
                  <tr
                    key={idx}
                    className="hover:bg-blue-50 cursor-pointer"
                    onClick={() => activeTab === "combined" && openRowEditor(idx)}
                  >
                    {columns.map((col) => (
                      <td
                        key={col}
                        className="px-4 py-3 text-sm text-gray-900 whitespace-normal break-words"
                        style={{ maxWidth: "300px" }}
                      >
                        {formatCell(row[col])}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* Footer */}
      {hasData && (
        <div className="bg-white border-t px-6 py-3">
          <div className="flex items-center justify-between text-sm text-gray-600">
            <span>Showing {filteredAndSortedData.length} of {currentData.length} records</span>
            <button
              onClick={() => {
                const headers = columns.join(",");
                const rows = filteredAndSortedData.map((r) =>
                  columns.map((c) => `"${String(formatCell(r[c]) || "").replace(/"/g, '""')}"`).join(",")
                );
                const csv = [headers, ...rows].join("\n");
                const blob = new Blob([csv], { type: "text/csv" });
                const url = URL.createObjectURL(blob);
                const a = document.createElement("a");
                a.href = url;
                a.download = `export_${activeTab}_${new Date().toISOString().slice(0, 10)}.csv`;
                a.click();
                URL.revokeObjectURL(url);
              }}
              className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
            >
              <Download size={18} />
              Export
            </button>
          </div>
        </div>
      )}

      {showCategoryManager && <CategoryManager />}

      {editingRow && (
        <RowEditor
          row={editingRow}
          rowIndex={editingRowIndex}
          onClose={closeRowEditor}
        />
      )}
    </div>
  );
};

export default App;
