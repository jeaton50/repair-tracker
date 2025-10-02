import React, { useState, useMemo, useEffect } from 'react';
import { Search, Download, ChevronDown, ChevronUp, Upload, FileSpreadsheet, Save, Cloud } from 'lucide-react';
import { db } from './firebase';
import { collection, doc, setDoc, getDoc, updateDoc } from 'firebase/firestore';

const RepairTrackerSheet = () => {
  const [activeTab, setActiveTab] = useState('combined');
  const [searchTerm, setSearchTerm] = useState('');
  const [locationFilter, setLocationFilter] = useState('');
  const [sortConfig, setSortConfig] = useState({ key: null, direction: 'asc' });
  const [ticketData, setTicketData] = useState([]);
  const [reportData, setReportData] = useState([]);
  const [loading, setLoading] = useState(false);
  const [wrapText, setWrapText] = useState(false);
  const [editingCell, setEditingCell] = useState(null);
  const [saving, setSaving] = useState(false);
  const [lastSaved, setLastSaved] = useState(null);

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
      if (typeof XLSX === 'undefined') {
        alert('Please wait a moment for the library to load and try again');
        setLoading(false);
        return;
      }

      const arrayBuffer = await file.arrayBuffer();
      const workbook = XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
      
      const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { 
        header: 1, 
        defval: "", 
        raw: false 
      });
      
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

      if (type === 'tickets') {
        setTicketData(data);
      } else {
        setReportData(data);
      }
      
      alert(`Successfully loaded ${data.length} records from ${file.name}`);
    } catch (err) {
      alert(`Error reading file: ${err.message}`);
      console.error('File upload error:', err);
    } finally {
      setLoading(false);
    }
  };

  const combinedData = useMemo(() => {
    if (reportData.length === 0) return [];
    
    const normalizeBarcode = (barcode) => {
      if (!barcode) return '';
      return String(barcode).trim().toUpperCase();
    };
    
    const ticketMap = new Map();
    ticketData.forEach(ticket => {
      const barcode = normalizeBarcode(ticket['Barcode']);
      if (barcode) {
        ticketMap.set(barcode, ticket);
      }
    });

    const calculateAge = (dateStr) => {
      if (!dateStr) return '';
      try {
        const date = new Date(dateStr);
        const today = new Date();
        const diffTime = Math.abs(today - date);
        const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
        return diffDays;
      } catch {
        return '';
      }
    };

    const result = reportData.map(report => {
      const reportBarcode = normalizeBarcode(report['Barcode#']);
      const ticket = ticketMap.get(reportBarcode) || {};
      const ticketNotes = ticket['Notes'] || '';
      
      return {
        'Meeting Note': '',
        'Requires Follow Up': '',
        'Dept': '',
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
    
    return result;
  }, [ticketData, reportData]);

  const [editedCombinedData, setEditedCombinedData] = useState([]);
  
  // Load data from Firebase when combined data changes
  useEffect(() => {
    const loadNotesFromFirebase = async () => {
      if (combinedData.length === 0) return;
      
      try {
        const updatedData = [...combinedData];
        
        // Load notes for each barcode
        for (let i = 0; i < updatedData.length; i++) {
          const barcode = updatedData[i]['Barcode#'];
          if (!barcode) continue;
          
          const docRef = doc(db, 'repairNotes', barcode);
          const docSnap = await getDoc(docRef);
          
          if (docSnap.exists()) {
            const data = docSnap.data();
            updatedData[i]['Meeting Note'] = data.meetingNote || '';
            updatedData[i]['Requires Follow Up'] = data.requiresFollowUp || '';
          }
        }
        
        setEditedCombinedData(updatedData);
        console.log('Loaded notes from Firebase');
      } catch (error) {
        console.error('Error loading notes:', error);
        setEditedCombinedData(combinedData);
      }
    };
    
    loadNotesFromFirebase();
  }, [combinedData]);

  const handleCellEdit = (rowIndex, column, value) => {
    const updatedData = [...editedCombinedData];
    updatedData[rowIndex][column] = value;
    setEditedCombinedData(updatedData);
  };

  const saveNotesToFirebase = async () => {
    setSaving(true);
    try {
      const savePromises = editedCombinedData.map(async (row) => {
        const barcode = row['Barcode#'];
        if (!barcode) return;
        
        const docRef = doc(db, 'repairNotes', barcode);
        
        try {
          await setDoc(docRef, {
            barcode: barcode,
            meetingNote: row['Meeting Note'] || '',
            requiresFollowUp: row['Requires Follow Up'] || '',
            lastUpdated: new Date().toISOString()
          }, { merge: true });
        } catch (error) {
          console.error(`Error saving note for ${barcode}:`, error);
        }
      });
      
      await Promise.all(savePromises);
      setLastSaved(new Date());
      alert('Notes saved successfully to Firebase!');
    } catch (error) {
      console.error('Error saving notes:', error);
      alert('Error saving notes. Please try again.');
    } finally {
      setSaving(false);
    }
  };

  const getCurrentData = () => {
    switch (activeTab) {
      case 'tickets': return ticketData;
      case 'reports': return reportData;
      case 'combined': return editedCombinedData;
      case 'diagnostics': return [];
      default: return [];
    }
  };

  const currentData = getCurrentData();
  const columns = currentData.length > 0 ? Object.keys(currentData[0]) : [];

  const uniqueLocations = useMemo(() => {
    const locations = new Set();
    currentData.forEach(row => {
      const location = row['Location'] || row['Repair Location'];
      if (location && location.trim()) {
        locations.add(location.trim());
      }
    });
    return Array.from(locations).sort();
  }, [currentData]);

  const filteredAndSortedData = useMemo(() => {
    let filtered = currentData;
    
    if (locationFilter) {
      filtered = filtered.filter(row => {
        const location = row['Location'] || row['Repair Location'];
        return location === locationFilter;
      });
    }
    
    if (searchTerm) {
      filtered = filtered.filter(row =>
        Object.values(row).some(val =>
          String(val).toLowerCase().includes(searchTerm.toLowerCase())
        )
      );
    }

    if (sortConfig.key) {
      filtered.sort((a, b) => {
        const aVal = a[sortConfig.key];
        const bVal = b[sortConfig.key];
        if (aVal === bVal) return 0;
        const comparison = aVal > bVal ? 1 : -1;
        return sortConfig.direction === 'asc' ? comparison : -comparison;
      });
    }

    return filtered;
  }, [currentData, searchTerm, locationFilter, sortConfig]);

  const handleSort = (key) => {
    setSortConfig(prev => ({
      key,
      direction: prev.key === key && prev.direction === 'asc' ? 'desc' : 'asc'
    }));
  };

  const formatCell = (value) => {
    if (!value && value !== 0) return '';
    const str = String(value);
    
    if (str.includes('T') && str.includes('Z')) {
      try {
        const date = new Date(str);
        if (!isNaN(date.getTime())) {
          return date.toLocaleDateString();
        }
      } catch {}
    }
    
    const num = parseFloat(str.replace(/,/g, ''));
    if (!isNaN(num) && str.match(/^[\d,\.]+$/)) {
      if (num === Math.floor(num)) {
        return Math.floor(num).toString();
      }
      return num.toString();
    }
    
    return str;
  };

  const exportToCSV = () => {
    const headers = columns.join(',');
    const rows = filteredAndSortedData.map(row =>
      columns.map(col => {
        const value = row[col];
        const formattedValue = formatCell(value);
        return `"${String(formattedValue || '').replace(/"/g, '""')}"`;
      }).join(',')
    );
    const csv = [headers, ...rows].join('\n');
    
    const blob = new Blob([csv], { type: 'text/csv' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${activeTab}_export_${new Date().toISOString().split('T')[0]}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const hasData = ticketData.length > 0 || reportData.length > 0;

  return (
    <div className="w-full h-screen flex flex-col bg-gray-50">
      <div className="bg-white border-b border-gray-200 px-6 py-4">
        <div className="flex items-center justify-between">
          <div>
            <h1 className="text-2xl font-bold text-gray-800">Repair Tracker Dashboard</h1>
            <p className="text-sm text-gray-500 mt-1">Manage repair tickets and reports</p>
          </div>
          
          <div className="flex gap-2 items-center">
            {lastSaved && (
              <span className="text-xs text-gray-500 flex items-center gap-1">
                <Cloud size={14} />
                Last saved: {lastSaved.toLocaleTimeString()}
              </span>
            )}
            
            {activeTab === 'combined' && editedCombinedData.length > 0 && (
              <button
                onClick={saveNotesToFirebase}
                disabled={saving}
                className="flex items-center gap-2 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 cursor-pointer transition-colors text-sm disabled:opacity-50"
              >
                <Save size={16} />
                {saving ? 'Saving...' : 'Save Notes'}
              </button>
            )}
            
            <label className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 cursor-pointer transition-colors text-sm">
              <Upload size={16} />
              {ticketData.length > 0 ? 'Reload' : 'Upload'} Ticket List
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => handleFileUpload(e, 'tickets')}
                className="hidden"
                disabled={loading}
              />
            </label>
            <label className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 cursor-pointer transition-colors text-sm">
              <Upload size={16} />
              {reportData.length > 0 ? 'Reload' : 'Upload'} Repair Report
              <input
                type="file"
                accept=".xlsx,.xls"
                onChange={(e) => handleFileUpload(e, 'reports')}
                className="hidden"
                disabled={loading}
              />
            </label>
          </div>
        </div>
      </div>

      <div className="bg-white border-b border-gray-200">
        <div className="flex px-6 overflow-x-auto">
          <button
            onClick={() => setActiveTab('combined')}
            className={`px-6 py-3 font-medium border-b-2 transition-colors whitespace-nowrap ${
              activeTab === 'combined'
                ? 'border-blue-500 text-blue-600'
                : 'border-transparent text-gray-500 hover:text-gray-700'
            }`}
          >
            Combined Report
            <span className="ml-2 text-xs bg-gray-100 px-2 py-1 rounded-full">
              {editedCombinedData.length}
            </span>
          </button>
          <button
            onClick={() => setActiveTab('tickets')}
            className={`px-6 py-3 font-medium border-b-2 transition-colors whitespace-nowrap ${
              activeTab === 'tickets'
                ? 'border-blue-500 text-blue-600'
                : 'border-transparent text-gray-500 hover:text-gray-700'
            }`}
          >
            Repair Ticket List
            <span className="ml-2 text-xs bg-gray-100 px-2 py-1 rounded-full">
              {ticketData.length}
            </span>
          </button>
          <button
            onClick={() => setActiveTab('reports')}
            className={`px-6 py-3 font-medium border-b-2 transition-colors whitespace-nowrap ${
              activeTab === 'reports'
                ? 'border-blue-500 text-blue-600'
                : 'border-transparent text-gray-500 hover:text-gray-700'
            }`}
          >
            Repair Report
            <span className="ml-2 text-xs bg-gray-100 px-2 py-1 rounded-full">
              {reportData.length}
            </span>
          </button>
          <button
            onClick={() => setActiveTab('diagnostics')}
            className={`px-6 py-3 font-medium border-b-2 transition-colors whitespace-nowrap ${
              activeTab === 'diagnostics'
                ? 'border-orange-500 text-orange-600'
                : 'border-transparent text-gray-500 hover:text-gray-700'
            }`}
          >
            🔍 Diagnostics
          </button>
        </div>
      </div>

      {hasData && activeTab !== 'diagnostics' && (
        <div className="bg-white border-b border-gray-200 px-6 py-3 flex items-center gap-4">
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
          
          <div className="flex items-center gap-2">
            <select
              value={locationFilter}
              onChange={(e) => setLocationFilter(e.target.value)}
              className="px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent bg-white min-w-[180px]"
            >
              <option value="">All Locations</option>
              {uniqueLocations.map(location => (
                <option key={location} value={location}>
                  {location}
                </option>
              ))}
            </select>
            
            {locationFilter && (
              <button
                onClick={() => setLocationFilter('')}
                className="px-3 py-2 text-sm text-gray-600 hover:text-gray-800 hover:bg-gray-100 rounded-lg transition-colors"
              >
                Clear
              </button>
            )}
          </div>
          
          <button
            onClick={() => setWrapText(!wrapText)}
            className={`px-4 py-2 border rounded-lg transition-colors text-sm ${
              wrapText 
                ? 'bg-blue-600 text-white border-blue-600 hover:bg-blue-700' 
                : 'bg-white text-gray-700 border-gray-300 hover:bg-gray-50'
            }`}
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
      )}

      <div className="flex-1 overflow-hidden px-6 py-4">
        {loading ? (
          <div className="flex items-center justify-center h-full">
            <div className="text-center">
              <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
              <p className="text-gray-500">Loading file...</p>
            </div>
          </div>
        ) : activeTab === 'diagnostics' ? (
          <div className="max-w-4xl mx-auto space-y-6 overflow-y-auto h-full">
            <div className="bg-white p-6 rounded-lg shadow">
              <h2 className="text-xl font-semibold text-gray-800 mb-4">Data Matching Diagnostics</h2>
              
              <div className="space-y-4">
                <div className="p-4 bg-blue-50 rounded-lg">
                  <h3 className="font-semibold text-blue-900 mb-2">📋 Repair Ticket List</h3>
                  <p className="text-sm text-blue-800">Records: {ticketData.length}</p>
                </div>
                
                <div className="p-4 bg-green-50 rounded-lg">
                  <h3 className="font-semibold text-green-900 mb-2">🔧 Repair Report</h3>
                  <p className="text-sm text-green-800">Records: {reportData.length}</p>
                </div>
              </div>
            </div>
          </div>
        ) : !hasData ? (
          <div className="flex items-center justify-center h-full">
            <div className="text-center max-w-lg bg-white p-12 rounded-lg shadow-lg">
              <FileSpreadsheet className="mx-auto text-blue-500 mb-6" size={64} />
              <h3 className="text-2xl font-semibold text-gray-800 mb-3">Welcome to Repair Tracker</h3>
              <p className="text-gray-600 mb-6">Upload your Excel files to get started.</p>
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
                  {columns.map((col) => {
                    const isEditable = activeTab === 'combined' && (col === 'Meeting Note' || col === 'Requires Follow Up');
                    return (
                      <th
                        key={col}
                        onClick={() => handleSort(col)}
                        className={`px-4 py-3 text-left text-xs font-medium text-gray-700 uppercase tracking-wider cursor-pointer hover:bg-gray-100 bg-gray-50 ${
                          wrapText ? 'whitespace-normal' : 'whitespace-nowrap'
                        }`}
                      >
                        <div className="flex items-center gap-2">
                          {col}
                          {isEditable && <span className="text-blue-500">✏️</span>}
                          {sortConfig.key === col && (
                            sortConfig.direction === 'asc' ? <ChevronUp size={14} /> : <ChevronDown size={14} />
                          )}
                        </div>
                      </th>
                    );
                  })}
                </tr>
              </thead>
              <tbody className="bg-white divide-y divide-gray-200">
                {filteredAndSortedData.map((row, idx) => (
                  <tr key={idx} className="hover:bg-gray-50">
                    {columns.map((col) => {
                      const isEditable = activeTab === 'combined' && (col === 'Meeting Note' || col === 'Requires Follow Up');
                      const cellValue = row[col];
                      const isEditing = editingCell?.rowIndex === idx && editingCell?.column === col;
                      
                      return (
                        <td 
                          key={col} 
                          className={`px-4 py-3 text-sm text-gray-900 ${
                            wrapText ? 'whitespace-normal break-words' : 'whitespace-nowrap'
                          } ${isEditable ? 'cursor-text' : ''}`}
                          style={wrapText ? { maxWidth: '300px' } : undefined}
                          onClick={() => {
                            if (isEditable) {
                              setEditingCell({ rowIndex: idx, column: col });
                            }
                          }}
                        >
                          {isEditable && isEditing ? (
                            <textarea
                              autoFocus
                              value={cellValue}
                              onChange={(e) => handleCellEdit(idx, col, e.target.value)}
                              onBlur={() => setEditingCell(null)}
                              className="w-full min-h-[60px] p-2 border border-blue-500 rounded focus:outline-none focus:ring-2 focus:ring-blue-500 resize-y"
                              style={{ minWidth: '200px' }}
                              onKeyDown={(e) => {
                                if (e.key === 'Escape') {
                                  setEditingCell(null);
                                }
                              }}
                            />
                          ) : isEditable ? (
                            <div 
                              className={`min-h-[40px] p-2 border border-transparent hover:border-gray-300 rounded whitespace-pre-wrap ${
                                !cellValue ? 'text-gray-400 italic' : ''
                              }`}
                            >
                              {cellValue || 'Click to add note...'}
                            </div>
                          ) : (
                            formatCell(cellValue)
                          )}
                        </td>
                      );
                    })}
                  </tr>
                ))}
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
              {locationFilter && (
                <span className="text-blue-600">Location: {locationFilter}</span>
              )}
              {searchTerm && (
                <span className="text-blue-600">Search: "{searchTerm}"</span>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default RepairTrackerSheet;