// src/sharePointNotesService.js
import * as XLSX from "xlsx";

/**
 * Service for reading/writing repair notes to SharePoint Excel file
 * Replaces Firebase Firestore.
 */
class SharePointNotesService {
  constructor(oneDriveService, config) {
    this.ods = oneDriveService;
    this.config = config;
    this.notesCache = new Map(); // In-memory cache
    this.saveQueue = new Map(); // Pending saves
    this.isSaving = false;
  }

  /**
   * Load all notes from SharePoint Excel file
   */
  async loadAllNotes() {
    try {
      console.log("📥 Loading notes from SharePoint...");
      
      const notes = await this.ods.readExcelFromSharePoint({
        hostname: this.config.spHostname,
        sitePath: this.config.spSitePath,
        fileRelativePath: `${this.config.spBasePath}/repair_notes.xlsx`,
      });

      // Convert array to Map for faster lookup
      const notesMap = new Map();
      notes.forEach(note => {
        if (note["Barcode#"]) {
          notesMap.set(note["Barcode#"], {
            barcode: note["Barcode#"],
            meetingNote: note["Meeting Note"] || "",
            requiresFollowUp: note["Requires Follow Up"] || "",
            lastUpdated: note["Last Updated"] || new Date().toISOString(),
          });
        }
      });

      this.notesCache = notesMap;
      console.log(`✅ Loaded ${notesMap.size} notes from SharePoint`);
      return notesMap;
    } catch (error) {
      // File doesn't exist yet, return empty map
      if (error.message?.includes("404") || error.message?.includes("not found")) {
        console.log("📝 Notes file doesn't exist yet, starting fresh");
        this.notesCache = new Map();
        return new Map();
      }
      console.error("❌ Error loading notes:", error);
      throw error;
    }
  }

  /**
   * Get note for specific barcode (from cache)
   */
  getNote(barcode) {
    return this.notesCache.get(barcode) || {
      barcode,
      meetingNote: "",
      requiresFollowUp: "",
      lastUpdated: null,
    };
  }

  /**
   * Update note in cache and queue for save
   */
  updateNote(barcode, meetingNote, requiresFollowUp) {
    const note = {
      barcode,
      meetingNote: meetingNote || "",
      requiresFollowUp: requiresFollowUp || "",
      lastUpdated: new Date().toISOString(),
    };

    // Update cache immediately
    this.notesCache.set(barcode, note);

    // Queue for save
    this.saveQueue.set(barcode, note);

    return note;
  }

  /**
   * Delete note
   */
  deleteNote(barcode) {
    this.notesCache.delete(barcode);
    this.saveQueue.set(barcode, null); // Mark for deletion
  }

  /**
   * Save all queued changes to SharePoint
   */
  async saveToSharePoint() {
    if (this.saveQueue.size === 0) {
      console.log("No changes to save");
      return;
    }

    if (this.isSaving) {
      console.log("Save already in progress, skipping...");
      return;
    }

    this.isSaving = true;
    const changesToSave = new Map(this.saveQueue);
    this.saveQueue.clear();

    try {
      console.log(`💾 Saving ${changesToSave.size} note changes to SharePoint...`);

      // Apply queued changes to cache
      changesToSave.forEach((note, barcode) => {
        if (note === null) {
          // Delete
          this.notesCache.delete(barcode);
        } else {
          // Update
          this.notesCache.set(barcode, note);
        }
      });

      // Convert Map to array for Excel
      const notesArray = Array.from(this.notesCache.values()).map(note => ({
        "Barcode#": note.barcode,
        "Meeting Note": note.meetingNote,
        "Requires Follow Up": note.requiresFollowUp,
        "Last Updated": note.lastUpdated,
      }));

      // Create Excel workbook
      const ws = XLSX.utils.json_to_sheet(notesArray);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Repair Notes");

      // Set column widths
      ws['!cols'] = [
        { wch: 15 },  // Barcode#
        { wch: 50 },  // Meeting Note
        { wch: 40 },  // Requires Follow Up
        { wch: 20 },  // Last Updated
      ];

      // Convert to binary
      const excelBuffer = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });

      // Upload to SharePoint
      await this.ods.uploadExcelToSharePoint({
        hostname: this.config.spHostname,
        sitePath: this.config.spSitePath,
        fileRelativePath: `${this.config.spBasePath}/repair_notes.xlsx`,
        data: excelBuffer,
      });

      console.log(`✅ Saved ${changesToSave.size} notes to SharePoint`);
    } catch (error) {
      console.error("❌ Error saving notes:", error);
      // Re-queue failed saves
      changesToSave.forEach((note, barcode) => {
        this.saveQueue.set(barcode, note);
      });
      throw error;
    } finally {
      this.isSaving = false;
    }
  }

  /**
   * Batch import notes from array
   */
  async importNotes(notesArray) {
    console.log(`📥 Importing ${notesArray.length} notes...`);
    
    notesArray.forEach(note => {
      if (note.barcode) {
        this.updateNote(note.barcode, note.meetingNote, note.requiresFollowUp);
      }
    });

    await this.saveToSharePoint();
    console.log(`✅ Imported ${notesArray.length} notes`);
  }

  /**
   * Get all notes as array
   */
  getAllNotes() {
    return Array.from(this.notesCache.values());
  }
}

export default SharePointNotesService;
