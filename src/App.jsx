// src/App.jsx (or wherever this component lives)

import { Client } from "@microsoft/microsoft-graph-client";
import { loginRequest, graphConfig } from "./authConfig";
import React, { useState, useMemo, useEffect, useRef } from 'react';
import { Search, Download, ChevronDown, ChevronUp, Upload, FileSpreadsheet, RefreshCw, Cloud } from 'lucide-react';
import { db } from './firebase';
import { doc, setDoc, onSnapshot, serverTimestamp, deleteDoc } from 'firebase/firestore';
import { PublicClientApplication } from '@azure/msal-browser';
import { msalConfig } from "./authConfig";
import OneDriveService from './oneDriveService';

const msalInstance = new PublicClientApplication(msalConfig);

// --- ✅ put the helper RIGHT HERE (top-level, not inside JSX/components) ---
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

async function debugListFolder(graph, folderPath) {
  const encodePath = (p) => p.split("/").map(encodeURIComponent).join("/");
  try {
    const encoded = encodePath(folderPath || "RepairTracker");
    const resp = await graph.api(`/me/drive/root:/${encoded}:/children`).get();
    console.log("DEBUG children of", folderPath, resp.value.map(v => ({
      name: v.name, id: v.id, isFolder: !!v.folder
    })));
  } catch (e) {
    console.error("DEBUG failed to list folder:", folderPath, e);
  }
}

/* ---------------- Row Editor (unchanged except for imports) ---------------- */

const SHARED_URL =
  "https://rentexinc-my.sharepoint.com/:f:/g/personal/justin_eaton_rentex_com/EksOqWxAKtVCjk5AnQaOasMBMow-ZyhwkFQ26M-mYyEcPw?e=EbTzEP";

const RowEditor = ({ row, rowIndex, onClose }) => {
  const [meetingNote, setMeetingNote] = useState(row['Meeting Note'] || '');
  const [followUp, setFollowUp] = useState(row['Requires Follow Up'] || '');
  const [lastSaved, setLastSaved] = useState(null);
  const [isSaving, setIsSaving] = useState(false);
  const saveTimeoutRef = useRef(null);
  const barcode = row['Barcode#'];

  // Auto-save
  useEffect(() => {
    if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current);

    saveTimeoutRef.current = setTimeout(async () => {
      if (!barcode) return;
      setIsSaving(true);
      try {
        const docRef = doc(db, 'repairNotes', barcode);
        await setDoc(
          docRef,
          { barcode, meetingNote, requiresFollowUp: followUp, lastUpdated: serverTimestamp() },
          { merge: true }
        );
        setLastSaved(new Date());
      } catch (error) {
        console.error('Auto-save error:', error);
        alert('Failed to save. Check your Firebase connection.');
      } finally {
        setIsSaving(false);
      }
    }, 1000);

    return () => {
      if (saveTimeoutRef.current) clearTimeout(saveTimeoutRef.current);
    };
  }, [meetingNote, followUp, barcode]);

  // Real-time sync
  useEffect(() => {
    if (!barcode) return;
    const docRef = doc(db, 'repairNotes', barcode);
    const unsubscribe = onSnapshot(
      docRef,
      (snapshot) => {
        if (snapshot.exists()) {
          const data = snapshot.data();
          if (document.activeElement?.name !== 'meetingNote') setMeetingNote(data.meetingNote || '');
          if (document.activeElement?.name !== 'followUp') setFollowUp(data.requiresFollowUp || '');
        }
      },
      (error) => console.error('Real-time sync error:', error)
    );
    return () => unsubscribe();
  }, [barcode]);

  const handleClearAll = () => {
    if (window.confirm('Clear all notes for this item?')) {
      setMeetingNote(''); setFollowUp('');
    }
  };

  const handleDeleteFromDatabase = async () => {
    if (!window.confirm('Permanently delete all notes for this item from the database?')) return;
    try {
      const docRef = doc(db, 'repairNotes', barcode);
      await deleteDoc(docRef);
      setMeetingNote(''); setFollowUp('');
      alert('Notes deleted from database');
    } catch (error) {
      console.error('Delete error:', error);
      alert('Failed to delete. Please try again.');
    }
  };

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
      <div className="bg-white rounded-lg shadow-xl w-full max-w-4xl max-h-[90vh] overflow-hidden flex flex-col">
        <div className="p-6 border-b border-gray-200 flex justify-between items-center">
          <div>
            <h2 className="text-xl font-bold text-gray-800">Edit Repair Item</h2>
            <p className="text-sm text-gray-500 mt-1">{row['Barcode#']} - {row['Equipment']}</p>
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
              <button type="button" onClick={() => setMeetingNote('')} className="text-xs px-2 py-1 text-red-600 border border-red-300 rounded hover:bg-red-50">Clear</button>
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
                <button type="button" onClick={() => setFollowUp(meetingNote)} className="text-xs px-2 py-1 border border-gray-300 rounded hover:bg-gray-50">Copy from Meeting Note</button>
                <button type="button" onClick={() => setFollowUp('')} className="text-xs px-2 py-1 text-red-600 border border-red-300 rounded hover:bg-red-50">Clear</button>
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
            <div><span className="font-semibold text-gray-700">Damage:</span> {row['Damage Description']}</div>
            <div><span className="font-semibold text-gray-700">Ticket Notes:</span> {row['Ticket Description']}</div>
            <div><span className="font-semibold text-gray-700">Reason:</span> {row['Repair Reason']}</div>
          </div>
        </div>

        <div className="p-6 border-t border-gray-200 flex justify-between">
          <div className="flex gap-2">
            <button onClick={handleClearAll} className="px-6 py-2 text-orange-600 bg-orange-50 border border-orange-300 rounded-lg hover:bg-orange-100">Clear All Notes</button>
            <button onClick={handleDeleteFromDatabase} className="px-6 py-2 text-red-600 bg-red-50 border border-red-300 rounded-lg hover:bg-red-100">Delete from Database</button>
          </div>
          <button onClick={onClose} className="px-6 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700">Close</button>
        </div>
      </div>
    </div>
  );
};

/* --------------------------- Main Component --------------------------- */

const RepairTrackerSheet = () => {
  const [activeTab, setActiveTab] = useState('combined');
  const [searchTerm, setSearchTerm] = useState('');
  const [locationFilter, setLocationFilter] = useState('');
  const [pmFilter, setPmFilter] = useState('');
  const [sortConfig, setSortConfig] = useState({ key: null, direction: 'asc' });
  const [ticketData, setTicketData] = useState([]);
  const [reportData, setReportData] = useState([]);
  const [categoryMapping, setCategoryMapping] = useState([]);
  const [loading, setLoading] = useState(false);
  const [wrapText, setWrapText] = useState(false);
  const [showCategoryManager, setShowCategoryManager] = useState(false);
  const [unmatchedCategories, setUnmatchedCategories] = useState([]);

  const [editingRow, setEditingRow] = useState(null);
  const [editingRowIndex, setEditingRowIndex] = useState(null);

  const [combinedDataWithNotes, setCombinedDataWithNotes] = useState([]);
  const [notesMap, setNotesMap] = useState(new Map());
  const notesListenersRef = useRef(new Map());

  // Auth / Graph
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [accessToken, setAccessToken] = useState(null);
  const [userName, setUserName] = useState('');
  const [lastSync, setLastSync] = useState(null);
  const [autoRefresh, setAutoRefresh] = useState(true);
  const [msalInitialized, setMsalInitialized] = useState(false);

  // Initialize MSAL
  useEffect(() => {
    const initializeMsal = async () => {
      try {
        await msalInstance.initialize();
        setMsalInitialized(true);
      } catch (error) {
        console.error('MSAL initialization failed:', error);
      }
    };
    initializeMsal();
  }, []);

  // Silent sign-in if possible
  useEffect(() => {
    if (!msalInitialized) return;
    const checkAuth = async () => {
      const accounts = msalInstance.getAllAccounts();
      if (accounts.length > 0) {
        try {
          const response = await msalInstance.acquireTokenSilent({ ...loginRequest, account: accounts[0] });
          setAccessToken(response.accessToken);
          setIsAuthenticated(true);
          setUserName(accounts[0].name);
        } catch (error) {
          console.error('Silent token acquisition failed:', error);
        }
      }
    };
    checkAuth();
  }, [msalInitialized]);

  // Auto-load after sign-in
  useEffect(() => {
    if (isAuthenticated && accessToken) {
      loadFromOneDrive(true);
    }
  }, [isAuthenticated, accessToken]);

  // Optional periodic refresh
  // Auto-refresh every 5 minutes while authenticated
useEffect(() => {
  if (!autoRefresh || !isAuthenticated) return;
  const interval = setInterval(() => loadFromOneDrive(true), 5 * 60 * 1000);
  return () => clearInterval(interval);
}, [autoRefresh, isAuthenticated]); // note: no accessToken in deps

// Login (don’t store token in state; fetch fresh inside loadFromOneDrive)
const handleLogin = async () => {
  try {
    await msalInstance.loginPopup(loginRequest);
    const account = msalInstance.getAllAccounts()[0];
    setIsAuthenticated(true);
    setUserName(account?.name || "");
    await loadFromOneDrive(false); // will call ensureAccessToken() internally
  } catch (err) {
    console.error("Login failed:", err);
    alert("Failed to sign in to Microsoft. Please try again.");
  }
};


  const handleLogout = () => {
    const account = msalInstance.getAllAccounts()[0];
    msalInstance.logoutPopup({ account });
    setIsAuthenticated(false);
    setAccessToken(null);
    setUserName('');
    setTicketData([]); setReportData([]); setCategoryMapping([]);
  };

  // Build Graph client from token
  const getGraphClient = () => Client.init({ authProvider: (done) => done(null, accessToken) });

// Load from OneDrive/SharePoint
const loadFromOneDrive = async (silent = false) => {
  if (!silent) setLoading(true);
  try {
    // 1) Always fetch a fresh token right before calling Graph
    const token = await ensureAccessToken();
    if (!token) throw new Error("No access token returned");

    // 2) Build a Graph client from the token
    const graph = Client.init({
      authProvider: (done) => done(null, token),
    });

    // 3) Create the OneDrive service (folderPath from config, default to "RepairTracker")
    const ods = new OneDriveService(graph, graphConfig.folderPath || "RepairTracker");

    // (Optional) quick one-time probe to verify the folder is resolvable
    // await debugListFolder(graph, graphConfig.folderPath || "RepairTracker");

    // 4) Helpers that try multiple filenames (real names first, then canonical)
    const tryExcel = async (names) => {
      for (const n of names) {
        try {
          return await ods.readExcelFileShared(n);
        } catch (e) {
          // If it's a not-found, try the next candidate
          if (e?.statusCode === 404 || String(e?.message || "").includes("not found")) continue;
          // Log other errors but keep trying next name
          console.warn("Excel read error for", n, e);
        }
      }
      return [];
    };

    const tryJson = async (names) => {
      for (const n of names) {
        try {
          return await ods.readJsonFileShared(n);
        } catch (e) {
          if (e?.statusCode === 404 || String(e?.message || "").includes("not found")) continue;
          console.warn("JSON read error for", n, e);
        }
      }
      return [];
    };

    // 5) Load all three in parallel (adjust filenames to match your drive)
    const [tickets, reports, mapping] = await Promise.all([
      tryExcel(["repair ticket list example.xlsx", "ticket_list.xlsx"]),
      tryExcel(["rtp repair report example.xlsx", "repair_report.xlsx"]),
      tryJson(["cleaned-pm-mapping (2).json", "category_mapping.json"]),
    ]);

    // 6) Update state
    setTicketData(tickets);
    setReportData(reports);
    setCategoryMapping(mapping);
    setLastSync(new Date());

    if (!silent) {
      alert(
        `Loaded from OneDrive:\n` +
        `${tickets.length} tickets\n` +
        `${reports.length} reports\n` +
        `${mapping.length} category mappings`
      );
    }
  } catch (e) {
    console.error("OneDrive load failed:", e);
    if (!silent) alert(e.message || "Failed to load from OneDrive.");
  } finally {
    if (!silent) setLoading(false);
  }
};



  /* ---------------------- XLSX loader for manual uploads ---------------------- */

  useEffect(() => {
    const loadXLSX = async () => {
      if (typeof XLSX !== 'undefined') return;
      const script = document.createElement('script');
      script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
      return new Promise((resolve, reject) => {
        script.onload = resolve;
        script.onerror = reject;
        document.head.appendChild(script);
      });
    };
    loadXLSX().catch(err => console.error('Failed to load XLSX:', err));
  }, []);

  const handleFileUpload = async (e, type) => {
    const file = e.target.files[0];
    if (!file) return;
    setLoading(true);
    try {
      if (type === 'mapping') {
        const text = await file.text();
        const jsonData = JSON.parse(text);
        setCategoryMapping(jsonData);
        alert(`Successfully loaded ${Array.isArray(jsonData) ? jsonData.length : Object.keys(jsonData || {}).length} category mappings`);
        setLoading(false);
        return;
      }

      if (typeof XLSX === 'undefined') {
        alert('Please wait for the library to load and try again');
        setLoading(false);
        return;
      }

      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
      const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1, defval: "", raw: false });

      if (rawData.length < 2) {
        alert('File appears to be empty or has incorrect format');
        setLoading(false);
        return;
      }

      const headers = rawData[1].map(h => String(h || '').trim()).filter(h => h);
      const dataRows = rawData.slice(2);

      const data = dataRows
        .filter(row => row && row.some(cell => cell !== "" && cell !== null))
        .map(row => {
          const obj = {};
          headers.forEach((header, index) => {
            obj[header] = row[index] !== undefined && row[index] !== null ? String(row[index]) : "";
          });
          return obj;
        });

      if (type === 'tickets') setTicketData(data);
      else setReportData(data);

      alert(`Successfully loaded ${data.length} records from ${file.name}`);
    } catch (err) {
      alert(`Error reading file: ${err.message}`);
      console.error('File upload error:', err);
    } finally {
      setLoading(false);
    }
  };

  /* ------------------------------ Merge logic ------------------------------ */

  const baseCombinedData = useMemo(() => {
    if (reportData.length === 0) return [];

    const normalizeBarcode = (barcode) => (!barcode ? '' : String(barcode).trim().toUpperCase());

    const ticketMap = new Map();
    ticketData.forEach(ticket => {
      const barcode = normalizeBarcode(ticket['Barcode']);
      if (barcode) ticketMap.set(barcode, ticket);
    });

    const categoryToPM = new Map();
    categoryMapping.forEach(item => {
      if (item.category && item.pm) {
        categoryToPM.set(item.category.trim().toUpperCase(), {
          pm: item.pm,
          department: item.department || '',
          categoryText: item.category_text || ''
        });
      }
    });

    const unmatchedSet = new Set();

    const calculateAge = (dateStr) => {
      if (!dateStr) return '';
      try {
        const date = new Date(dateStr);
        const today = new Date();
        const diffDays = Math.ceil(Math.abs(today - date) / (1000 * 60 * 60 * 24));
        return diffDays;
      } catch { return ''; }
    };

    const result = reportData.map(report => {
      const reportBarcode = normalizeBarcode(report['Barcode#']);
      const ticket = ticketMap.get(reportBarcode) || {};
      const ticketNotes = ticket['Notes'] || '';

      const category = (report['Category'] || '').trim();
      const mappingInfo = categoryToPM.get(category.toUpperCase());
      const assignedPM = mappingInfo ? mappingInfo.pm : '';

      if (category && !assignedPM) unmatchedSet.add(category);

      return {
        'Meeting Note': '',
        'Requires Follow Up': '',
        'Assigned To': assignedPM,
        'Location': report['Repair Location'] || ticket['Location'] || '',
        'Repair Ticket': report['Ticket'] || '',
        'Asset Repair Age': calculateAge(report['Date In']),
        'Barcode#': report['Barcode#'] || ticket['Barcode'] || '',
        'Equipment': `(${report['Equipment']}) - ${report['Description']}`,
        'Damage Description': report['Notes'] || '',
        'Ticket Description': ticketNotes,
        'Repair Reason': report['Repair Reason'] || '',
        'Last Order#': report['Last Order#'] || ticket['Order# to Bill'] || '',
        'Reference#': report['Reference#'] || '',
        'Customer': report['Customer'] || ticket['Customer'] || '',
        'Customer Title': report['Customer Title'] || '',
        'Repair Cost': report['Repair Cost'] || '0',
        'Date In': report['Date In'] || ticket['Creation Date'] || '',
        'Department': report['Department'] || '',
        'Category': report['Category'] || '',
        'Billable': ticket['Billable'] || report['Billable'] || '',
        'Created By': ticket['Created By'] || report['User In'] || '',
        'Repair Price': report['Repair Price'] || '0',
        'Repair Vendor': report['Repair Vendor'] || '',
        '_TicketMatched': Object.keys(ticket).length > 0 ? 'Yes' : 'No'
      };
    });

    setUnmatchedCategories(Array.from(unmatchedSet).sort());
    return result;
  }, [ticketData, reportData, categoryMapping]);

  // Subscribe to Firestore notes for current rows and merge
  useEffect(() => {
    if (baseCombinedData.length === 0) { setCombinedDataWithNotes([]); return; }

    const barcodes = baseCombinedData.map(row => row['Barcode#']).filter(Boolean);
    if (barcodes.length === 0) { setCombinedDataWithNotes(baseCombinedData); return; }

    // Clean up listeners for removed barcodes
    notesListenersRef.current.forEach((unsubscribe, barcode) => {
      if (!barcodes.includes(barcode)) { unsubscribe(); notesListenersRef.current.delete(barcode); }
    });

    // New listeners
    barcodes.forEach((barcode) => {
      if (notesListenersRef.current.has(barcode)) return;
      const docRef = doc(db, 'repairNotes', barcode);
      const unsubscribe = onSnapshot(
        docRef,
        (snapshot) => {
          setNotesMap(prevMap => {
            const newMap = new Map(prevMap);
            if (snapshot.exists()) {
              const data = snapshot.data();
              newMap.set(barcode, { meetingNote: data.meetingNote || '', requiresFollowUp: data.requiresFollowUp || '' });
            } else {
              newMap.set(barcode, { meetingNote: '', requiresFollowUp: '' });
            }
            return newMap;
          });
        },
        (error) => console.error(`Error syncing barcode ${barcode}:`, error)
      );
      notesListenersRef.current.set(barcode, unsubscribe);
    });

    // Merge
    const merged = baseCombinedData.map(row => {
      const barcode = row['Barcode#'];
      const notes = notesMap.get(barcode);
      return { ...row, 'Meeting Note': notes?.meetingNote || '', 'Requires Follow Up': notes?.requiresFollowUp || '' };
    });
    setCombinedDataWithNotes(merged);

    // Cleanup on unmount
    return () => {
      notesListenersRef.current.forEach(unsubscribe => unsubscribe());
      notesListenersRef.current.clear();
    };
  }, [baseCombinedData, notesMap]);

  const openRowEditor = (rowIndex) => {
    setEditingRowIndex(rowIndex);
    setEditingRow(filteredAndSortedData[rowIndex]);
  };
  const closeRowEditor = () => { setEditingRow(null); setEditingRowIndex(null); };

  const getCurrentData = () => {
    switch (activeTab) {
      case 'tickets': return ticketData;
      case 'reports': return reportData;
      case 'combined': return combinedDataWithNotes;
      case 'diagnostics': return [];
      default: return [];
    }
  };

  const currentData = getCurrentData();
  const columns = currentData.length > 0 ? Object.keys(currentData[0]).filter(col => !col.startsWith('_')) : [];

  const uniqueLocations = useMemo(() => {
    const locations = new Set();
    currentData.forEach(row => {
      const location = row['Location'] || row['Repair Location'];
      if (location && location.trim()) locations.add(location.trim());
    });
    return Array.from(locations).sort();
  }, [currentData]);

  const uniquePMs = useMemo(() => {
    const pms = new Set();
    combinedDataWithNotes.forEach(row => {
      const pm = row['Assigned To'];
      if (pm && pm.trim()) pms.add(pm.trim());
    });
    return Array.from(pms).sort();
  }, [combinedDataWithNotes]);

  const pmWorkload = useMemo(() => {
    const workload = {};
    combinedDataWithNotes.forEach(row => {
      const pm = row['Assigned To'] || 'Unassigned';
      if (!workload[pm]) workload[pm] = { count: 0, totalCost: 0, avgAge: 0, ages: [] };
      workload[pm].count++;
      const cost = parseFloat(row['Repair Cost']) || 0; workload[pm].totalCost += cost;
      const age = parseInt(row['Asset Repair Age']) || 0; if (age > 0) workload[pm].ages.push(age);
    });
    Object.keys(workload).forEach(pm => {
      if (workload[pm].ages.length > 0) {
        const sum = workload[pm].ages.reduce((a, b) => a + b, 0);
        workload[pm].avgAge = Math.round(sum / workload[pm].ages.length);
      }
    });
    return workload;
  }, [combinedDataWithNotes]);

  const allCategories = useMemo(() => {
    const cats = new Set();
    reportData.forEach(report => {
      const category = report['Category'];
      if (category && category.trim()) cats.add(category.trim());
    });
    return Array.from(cats).sort();
  }, [reportData]);

  const addCategoryMapping = (category, pm, department = '', categoryText = '') => {
    const newMapping = [...categoryMapping];
    const existingIndex = newMapping.findIndex(m => m.category.trim().toUpperCase() === category.trim().toUpperCase());
    if (existingIndex >= 0) {
      newMapping[existingIndex] = { category: category.trim(), pm: pm.trim(), department: department.trim(), category_text: categoryText.trim() };
    } else {
      newMapping.push({ category: category.trim(), pm: pm.trim(), department: department.trim(), category_text: categoryText.trim() });
    }
    setCategoryMapping(newMapping);
  };

  const removeCategoryMapping = (category) => {
    setCategoryMapping(prev => prev.filter(m => m.category.trim().toUpperCase() !== category.trim().toUpperCase()));
  };

  const exportCategoryMapping = () => {
    const json = JSON.stringify(categoryMapping, null, 2);
    const blob = new Blob([json], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `category_mapping_${new Date().toISOString().split('T')[0]}.json`;
    a.click(); URL.revokeObjectURL(url);
  };

  const filteredAndSortedData = useMemo(() => {
    let filtered = currentData;

    if (locationFilter) filtered = filtered.filter(row => (row['Location'] || row['Repair Location']) === locationFilter);
    if (pmFilter) {
      filtered = filtered.filter(row => pmFilter === '__unassigned__' ? !row['Assigned To'] : row['Assigned To'] === pmFilter);
    }
    if (searchTerm) {
      const s = searchTerm.toLowerCase();
      filtered = filtered.filter(row => Object.values(row).some(val => String(val).toLowerCase().includes(s)));
    }
    if (sortConfig.key) {
      filtered = [...filtered].sort((a, b) => {
        const aVal = a[sortConfig.key]; const bVal = b[sortConfig.key];
        if (aVal === bVal) return 0;
        const cmp = aVal > bVal ? 1 : -1;
        return sortConfig.direction === 'asc' ? cmp : -cmp;
      });
    }
    return filtered;
  }, [currentData, searchTerm, locationFilter, pmFilter, sortConfig]);

  const handleSort = (key) => {
    setSortConfig(prev => ({ key, direction: prev.key === key && prev.direction === 'asc' ? 'desc' : 'asc' }));
  };

  const formatCell = (value) => {
    if (!value && value !== 0) return '';
    const str = String(value);
    if (str.includes('T') && str.includes('Z')) {
      try {
        const date = new Date(str);
        if (!isNaN(date.getTime())) return date.toLocaleDateString();
      } catch {}
    }
    const num = parseFloat(str.replace(/,/g, ''));
    if (!isNaN(num) && str.match(/^[\d,\.]+$/)) return (num === Math.floor(num)) ? Math.floor(num).toString() : num.toString();
    return str;
  };

  const exportToCSV = () => {
    const headers = columns.join(',');
    const rows = filteredAndSortedData.map(row =>
      columns.map(col => `"${String(formatCell(row[col]) || '').replace(/"/g, '""')}"`).join(',')
    );
    const csv = [headers, ...rows].join('\n');
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = `${activeTab}_export_${new Date().toISOString().split('T')[0]}.csv`;
    a.click(); URL.revokeObjectURL(url);
  };

  const hasData = ticketData.length > 0 || reportData.length > 0;

  /* ------------------------- Category Manager modal ------------------------- */

  const CategoryManager = () => {
    const [newCategory, setNewCategory] = useState('');
    const [newPM, setNewPM] = useState('');
    const [newDepartment, setNewDepartment] = useState('');
    const [newCategoryText, setNewCategoryText] = useState('');
    return (
      <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center z-50 p-4">
        <div className="bg-white rounded-lg shadow-xl max-w-6xl w-full max-h-[90vh] overflow-hidden flex flex-col">
          <div className="p-6 border-b border-gray-200 flex justify-between items-center">
            <h2 className="text-2xl font-bold text-gray-800">Category to PM Mapping Manager</h2>
            <button onClick={() => setShowCategoryManager(false)} className="text-gray-500 hover:text-gray-700 text-2xl">×</button>
          </div>

          <div className="p-6 space-y-6 overflow-y-auto flex-1">
            <div className="bg-blue-50 p-4 rounded-lg">
              <h3 className="font-semibold text-blue-900 mb-3">Add New Mapping</h3>
              <div className="grid grid-cols-2 gap-3 mb-3">
                <div>
                  <label className="text-xs text-gray-600 mb-1 block">Category Code</label>
                  <select value={newCategory} onChange={(e) => setNewCategory(e.target.value)} className="w-full px-3 py-2 border border-gray-300 rounded-lg">
                    <option value="">Select Category</option>
                    {allCategories.map(cat => (<option key={cat} value={cat}>{cat}</option>))}
                  </select>
                </div>
                <div>
                  <label className="text-xs text-gray-600 mb-1 block">PM Name</label>
                  <input type="text" placeholder="PM Name" value={newPM} onChange={(e) => setNewPM(e.target.value)} className="w-full px-3 py-2 border border-gray-300 rounded-lg" />
                </div>
                <div>
                  <label className="text-xs text-gray-600 mb-1 block">Department</label>
                  <input type="text" placeholder="Department (optional)" value={newDepartment} onChange={(e) => setNewDepartment(e.target.value)} className="w-full px-3 py-2 border border-gray-300 rounded-lg" />
                </div>
                <div>
                  <label className="text-xs text-gray-600 mb-1 block">Category Description</label>
                  <input type="text" placeholder="Category description (optional)" value={newCategoryText} onChange={(e) => setNewCategoryText(e.target.value)} className="w-full px-3 py-2 border border-gray-300 rounded-lg" />
                </div>
              </div>
              <button
                onClick={() => {
                  if (newCategory && newPM) {
                    addCategoryMapping(newCategory, newPM, newDepartment, newCategoryText);
                    setNewCategory(''); setNewPM(''); setNewDepartment(''); setNewCategoryText('');
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
                <h3 className="font-semibold text-red-900 mb-2">Unmatched Categories ({unmatchedCategories.length})</h3>
                <p className="text-sm text-red-800 mb-3">These categories don't have PM assignments:</p>
                <div className="flex flex-wrap gap-2">
                  {unmatchedCategories.map(cat => (
                    <span key={cat} className="px-3 py-1 bg-red-100 text-red-800 rounded-full text-sm">{cat}</span>
                  ))}
                </div>
              </div>
            )}

            <div>
              <div className="flex justify-between items-center mb-3">
                <h3 className="font-semibold text-gray-800">Current Mappings ({categoryMapping.length})</h3>
                <button onClick={exportCategoryMapping} className="text-sm px-3 py-1 bg-gray-600 text-white rounded hover:bg-gray-700">Export JSON</button>
              </div>
              <div className="space-y-2 max-h-96 overflow-y-auto">
                {categoryMapping.map((mapping, idx) => (
                  <div key={idx} className="flex items-start justify-between p-4 bg-gray-50 rounded-lg border border-gray-200">
                    <div className="flex-1 space-y-1">
                      <div className="flex items-center gap-2">
                        <span className="font-bold text-gray-900">{mapping.category}</span>
                        <span className="text-gray-400">→</span>
                        <span className="font-semibold text-blue-600">{mapping.pm}</span>
                      </div>
                      {mapping.category_text && (<p className="text-sm text-gray-600">{mapping.category_text}</p>)}
                      {mapping.department && (<p className="text-xs text-gray-500">Department: {mapping.department}</p>)}
                    </div>
                    <button onClick={() => removeCategoryMapping(mapping.category)} className="text-red-600 hover:text-red-800 text-sm ml-4">Remove</button>
                  </div>
                ))}
              </div>
            </div>

            <div>
              <h3 className="font-semibold text-gray-800 mb-3">All Categories in Data ({allCategories.length})</h3>
              <div className="flex flex-wrap gap-2">
                {allCategories.map(cat => {
                  const hasMapping = categoryMapping.some(m => m.category.trim().toUpperCase() === cat.trim().toUpperCase());
                  return (
                    <span key={cat} className={`px-3 py-1 rounded-full text-sm ${hasMapping ? 'bg-green-100 text-green-800' : 'bg-gray-100 text-gray-800'}`}>
                      {cat} {hasMapping && '✓'}
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

  /* ---------------------------------- UI ---------------------------------- */

  return (
    <div className="w-full h-screen flex flex-col bg-gray-50">
      <div className="bg-white border-b border-gray-200 px-6 py-4">
        <div className="flex items-center justify-between">
          <div>
            <h1 className="text-2xl font-bold text-gray-800">Repair Tracker Dashboard</h1>
            <div className="flex items-center gap-3 mt-1">
              <p className="text-sm text-gray-500">
                {isAuthenticated ? `Connected as ${userName}` : 'Not connected to OneDrive'}
              </p>
              {isAuthenticated && (
                <div className="flex items-center gap-1 text-xs text-green-600 bg-green-50 px-2 py-1 rounded">
                  <Cloud size={12} /> OneDrive synced
                </div>
              )}
              {lastSync && (
                <span className="text-xs text-gray-500">
                  Last sync: {lastSync.toLocaleTimeString()}
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
                Sign in to OneDrive
              </button>
            ) : (
              <>
                <button
                  onClick={() => loadFromOneDrive(false)}
                  disabled={loading}
                  className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 disabled:opacity-50 text-sm"
                >
                  <RefreshCw size={16} className={loading ? 'animate-spin' : ''} />
                  Refresh
                </button>

                <button
                  onClick={handleLogout}
                  className="px-4 py-2 bg-gray-600 text-white rounded-lg hover:bg-gray-700 text-sm"
                >
                  Sign Out
                </button>
              </>
            )}

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

            <label className="flex items-center gap-2 px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700 cursor-pointer transition-colors text-sm">
              <Upload size={16} />
              Upload Mapping
              <input type="file" accept=".json" onChange={(e) => handleFileUpload(e, 'mapping')} className="hidden" disabled={loading} />
            </label>

            <label className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 cursor-pointer transition-colors text-sm">
              <Upload size={16} />
              Upload Tickets
              <input type="file" accept=".xlsx,.xls" onChange={(e) => handleFileUpload(e, 'tickets')} className="hidden" disabled={loading} />
            </label>

            <label className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 cursor-pointer transition-colors text-sm">
              <Upload size={16} />
              Upload Reports
              <input type="file" accept=".xlsx,.xls" onChange={(e) => handleFileUpload(e, 'reports')} className="hidden" disabled={loading} />
            </label>
          </div>
        </div>
      </div>

      <div className="bg-white border-b border-gray-200">
        <div className="flex px-6 overflow-x-auto">
          <button
            onClick={() => setActiveTab('combined')}
            className={`px-6 py-3 font-medium border-b-2 transition-colors whitespace-nowrap ${
              activeTab === 'combined' ? 'border-blue-500 text-blue-600' : 'border-transparent text-gray-500 hover:text-gray-700'
            }`}
          >
            Combined Report
            <span className="ml-2 text-xs bg-gray-100 px-2 py-1 rounded-full">{combinedDataWithNotes.length}</span>
          </button>
          <button
            onClick={() => setActiveTab('tickets')}
            className={`px-6 py-3 font-medium border-b-2 transition-colors whitespace-nowrap ${
              activeTab === 'tickets' ? 'border-blue-500 text-blue-600' : 'border-transparent text-gray-500 hover:text-gray-700'
            }`}
          >
            Repair Ticket List
            <span className="ml-2 text-xs bg-gray-100 px-2 py-1 rounded-full">{ticketData.length}</span>
          </button>
          <button
            onClick={() => setActiveTab('reports')}
            className={`px-6 py-3 font-medium border-b-2 transition-colors whitespace-nowrap ${
              activeTab === 'reports' ? 'border-blue-500 text-blue-600' : 'border-transparent text-gray-500 hover:text-gray-700'
            }`}
          >
            Repair Report
            <span className="ml-2 text-xs bg-gray-100 px-2 py-1 rounded-full">{reportData.length}</span>
          </button>
          <button
            onClick={() => setActiveTab('diagnostics')}
            className={`px-6 py-3 font-medium border-b-2 transition-colors whitespace-nowrap ${
              activeTab === 'diagnostics' ? 'border-orange-500 text-orange-600' : 'border-transparent text-gray-500 hover:text-gray-700'
            }`}
          >
            Diagnostics
          </button>
        </div>
      </div>

      {/* Filters + Export */}
      {hasData && activeTab !== 'diagnostics' && (
        <div className="bg-white border-b border-gray-200 px-6 py-3">
          <div className="flex items-center justify-between mb-3">
            <div className="flex-1 relative">
              <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400" size={18} />
              <input
                type="text"
                placeholder="Search across all columns..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full pl-10 pr-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
              />
            </div>

            <div className="flex items-center gap-2 ml-4">
             <select
  value={locationFilter}
  onChange={(e) => setLocationFilter(e.target.value)}
  className="px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white min-w-[180px]"
>
  <option value="">All Locations</option>
  {uniqueLocations.map((location) => (
    <option key={location} value={location}>
      {location}
    </option>
  ))}
</select>



              {locationFilter && (
                <button onClick={() => setLocationFilter('')} className="px-3 py-2 text-sm text-gray-600 hover:text-gray-800 hover:bg-gray-100 rounded-lg transition-colors">
                  Clear
                </button>
              )}

              {activeTab === 'combined' && (
                <>
                  <select
                    value={pmFilter}
                    onChange={(e) => setPmFilter(e.target.value)}
                    className="px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white min-w-[200px]"
                  >
                    <option value="">All Assigned To</option>
                    <option value="__unassigned__">Unassigned</option>
                    {uniquePMs.map(pm => (<option key={pm} value={pm}>{pm}</option>))}
                  </select>

                  {pmFilter && (
                    <button onClick={() => setPmFilter('')} className="px-3 py-2 text-sm text-gray-600 hover:text-gray-800 hover:bg-gray-100 rounded-lg transition-colors">
                      Clear
                    </button>
                  )}
                </>
              )}
            </div>

            <div className="flex items-center gap-2 ml-4">
              <button
                onClick={() => setWrapText(!wrapText)}
                className={`px-4 py-2 border rounded-lg transition-colors text-sm ${wrapText ? 'bg-blue-600 text-white border-blue-600 hover:bg-blue-700' : 'bg-white text-gray-700 border-gray-300 hover:bg-gray-50'}`}
              >
                {wrapText ? 'Unwrap' : 'Wrap'} Text
              </button>

              <button
                onClick={exportToCSV}
                disabled={currentData.length === 0}
                className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors disabled:opacity-50"
              >
                <Download size={18} />
                Export
              </button>
            </div>
          </div>

          {activeTab === 'combined' && (
            <div className="flex items-center gap-4 text-xs text-gray-500">
              <span>Total: {combinedDataWithNotes.length} items</span>
              <span className="text-green-600">
                Assigned: {combinedDataWithNotes.filter(row => row['Assigned To'] && row['Assigned To'] !== '').length}
              </span>
              <span className="text-red-600 font-semibold">
                Unassigned: {combinedDataWithNotes.filter(row => !row['Assigned To'] || row['Assigned To'] === '').length}
              </span>
              {categoryMapping.length > 0 && <span className="text-purple-600">{categoryMapping.length} category mappings</span>}
              {unmatchedCategories.length > 0 && <span className="text-orange-600 font-semibold">{unmatchedCategories.length} categories without PM</span>}
              <span className="text-blue-600">{uniquePMs.length} unique PMs</span>
            </div>
          )}
        </div>
      )}

      {/* Main table / diagnostics / empty states */}
      <div className="flex-1 overflow-hidden px-6 py-4">
        {loading ? (
          <div className="flex items-center justify-center h-full">
            <div className="text-center">
              <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
              <p className="text-gray-500">Loading...</p>
            </div>
          </div>
        ) : activeTab === 'diagnostics' ? (
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

              {combinedDataWithNotes.length > 0 && (
                <div className="p-4 bg-indigo-50 rounded-lg">
                  <h3 className="font-semibold text-indigo-900 mb-4">Workload by Assigned PM</h3>
                  <div className="space-y-3">
                    {Object.entries(pmWorkload)
                      .sort((a, b) => b[1].count - a[1].count)
                      .map(([pm, stats]) => (
                        <div key={pm} className="p-3 bg-white rounded-lg shadow-sm">
                          <div className="flex justify-between items-start mb-2">
                            <span className={`font-semibold ${pm === 'Unassigned' ? 'text-red-600' : 'text-gray-800'}`}>{pm}</span>
                            <span className="text-lg font-bold text-indigo-600">{stats.count} items</span>
                          </div>
                          <div className="grid grid-cols-3 gap-2 text-sm">
                            <div><span className="text-gray-600">Total Cost:</span><p className="font-semibold">${stats.totalCost.toFixed(2)}</p></div>
                            <div><span className="text-gray-600">Avg Age:</span><p className="font-semibold">{stats.avgAge} days</p></div>
                            <div><span className="text-gray-600">Workload %:</span><p className="font-semibold">{((stats.count / combinedDataWithNotes.length) * 100).toFixed(1)}%</p></div>
                          </div>
                          <div className="mt-2 h-2 bg-gray-200 rounded-full overflow-hidden">
                            <div className={`h-full ${pm === 'Unassigned' ? 'bg-red-500' : 'bg-indigo-500'}`} style={{ width: `${(stats.count / combinedDataWithNotes.length) * 100}%` }} />
                          </div>
                        </div>
                      ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        ) : !hasData ? (
          <div className="flex items-center justify-center h-full">
            <div className="text-center max-w-lg bg-white p-12 rounded-lg shadow-lg">
              <FileSpreadsheet className="mx-auto text-blue-500 mb-6" size={64} />
              <h3 className="text-2xl font-semibold text-gray-800 mb-3">Welcome to Repair Tracker</h3>
              <p className="text-gray-600 mb-6">Sign in to OneDrive to load shared data, or upload files manually</p>
            </div>
          </div>
        ) : currentData.length === 0 ? (
          <div className="flex items-center justify-center h-full">
            <div className="text-center bg-white p-8 rounded-lg shadow"><p className="text-gray-500">No data available</p></div>
          </div>
        ) : (
          <div className="bg-white rounded-lg shadow h-full overflow-auto">
            <table className="w-full border-collapse">
              <thead className="bg-gray-50 border-b border-gray-200 sticky top-0 z-10">
                <tr>
                  {columns.map((col) => {
                    const isEditable = activeTab === 'combined' && (col === 'Meeting Note' || col === 'Requires Follow Up');
                    return (
                      <th
                        key={col}
                        onClick={() => handleSort(col)}
                        className={`px-4 py-3 text-left text-xs font-medium text-gray-700 uppercase tracking-wider cursor-pointer hover:bg-gray-100 bg-gray-50 ${wrapText ? 'whitespace-normal' : 'whitespace-nowrap'}`}
                      >
                        <div className="flex items-center gap-2">
                          {col}
                          {isEditable && <span className="text-blue-500">✏️</span>}
                          {sortConfig.key === col && (sortConfig.direction === 'asc' ? <ChevronUp size={14} /> : <ChevronDown size={14} />)}
                        </div>
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody className="bg-white divide-y divide-gray-200">
                {filteredAndSortedData.map((row, idx) => {
                  const hasAssignment = row['Assigned To'] && row['Assigned To'] !== '';
                  const rowBgColor = activeTab === 'combined' && !hasAssignment ? 'bg-red-50' : '';
                  const isEditable = activeTab === 'combined';
                  return (
                    <tr key={idx} className={`${rowBgColor} ${isEditable ? 'hover:bg-blue-50 cursor-pointer' : 'hover:bg-gray-50'}`} onClick={() => isEditable && openRowEditor(idx)}>
                      {columns.map((col) => (
                        <td key={col} className={`px-4 py-3 text-sm text-gray-900 ${wrapText ? 'whitespace-normal break-words' : 'whitespace-nowrap'}`} style={wrapText ? { maxWidth: '300px' } : undefined}>
                          {formatCell(row[col])}
                        </td>
                      ))}
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {hasData && activeTab !== 'diagnostics' && (
        <div className="bg-white border-t border-gray-200 px-6 py-3">
          <div className="flex items-center justify-between text-sm text-gray-600">
            <span>Showing {filteredAndSortedData.length} of {currentData.length} records</span>
            <div className="flex items-center gap-4">
              {locationFilter && (<span className="text-blue-600">Location: {locationFilter}</span>)}
              {pmFilter && (<span className="text-green-600">Assigned To: {pmFilter === '__unassigned__' ? 'Unassigned' : pmFilter}</span>)}
              {searchTerm && (<span className="text-blue-600">Search: "{searchTerm}"</span>)}
            </div>
          </div>
        </div>
      )}

      {showCategoryManager && <CategoryManager />}

      {editingRow && <RowEditor row={editingRow} rowIndex={editingRowIndex} onClose={closeRowEditor} />}
    </div>
  );
};

export default RepairTrackerSheet;
