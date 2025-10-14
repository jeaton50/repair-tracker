// src/SharePointNotesService.js
import * as XLSX from "xlsx";

/**
 * Service for reading/writing repair notes to a SharePoint Excel file via OneDriveService.
 * Replaces Firebase Firestore.
 *
 * Columns (Excel):
 *   "Barcode#", "Meeting Note", "Requires Follow Up", "Last Updated"
 *
 * Usage:
 *   const notesSvc = new SharePointNotesService(oneDriveService, {
 *     spHostname: "rentexinc.sharepoint.com",
 *     spSitePath: "/sites/ProductManagers",
 *     spBasePath: "General/Repairs/RepairTracker",
 *     // notesFile: "repair_notes.xlsx" // optional
 *   });
 */
export default class SharePointNotesService {
  constructor(oneDriveService, config) {
    this.ods = oneDriveService;
    this.config = {
      ...config,
      notesFile: config?.notesFile || "repair_notes.xlsx",
    };
    this.notesCache = new Map(); // In-memory cache { BARCODE -> {barcode, meetingNote, requiresFollowUp, lastUpdated} }
    this.saveQueue = new Map();  // Pending saves { BARCODE -> note | null (delete) }
    this.isSaving = false;
  }

  /* ------------------------------ internals -------------------------------- */

  _filePath() {
    const base = (this.config.spBasePath || "").replace(/\/+$/g, "");
    return `${base}/${this.config.notesFile}`.replace(/^\/+/g, "");
  }

  _normBarcode(bc) {
    return bc ? String(bc).trim().toUpperCase() : "";
  }

  _noteShape(note = {}) {
    return {
      barcode: this._normBarcode(note.barcode ?? note["Barcode#"] ?? note["Barcode"]),
      meetingNote: String(note.meetingNote ?? note["Meeting Note"] ?? "").trim(),
      requiresFollowUp: String(note.requiresFollowUp ?? note["Requires Follow Up"] ?? "").trim(),
      lastUpdated: note.lastUpdated ?? note["Last Updated"] ?? new Date().toISOString(),
    };
  }

  /* --------------------------------- API ----------------------------------- */

  /** Load all notes from SharePoint Excel file into cache and return a Map. */
  async loadAllNotes() {
    try {
      console.log("📥 Loading notes from SharePoint…", this._filePath());
      const rows = await this.ods.readExcelFromSharePoint({
        hostname: this.config.spHostname,
        sitePath: this.config.spSitePath,
        fileRelativePath: this._filePath(),
      });

      const map = new Map();
      (rows || []).forEach(r => {
        const shaped = this._noteShape(r);
        if (shaped.barcode) map.set(shaped.barcode, shaped);
      });

      this.notesCache = map;
      console.log(`✅ Loaded ${map.size} notes from ${this._filePath()}`);
      return map;
    } catch (e) {
      // Common "file not found" signatures from Graph SDK
      const status = e?.statusCode || e?.status;
      const code = e?.code || e?.body?.error?.code;
      if (status === 404 || code === "itemNotFound" || /404|not\s*found/i.test(String(e?.message))) {
        console.warn("📝 Notes file not found; starting with empty cache:", this._filePath());
        this.notesCache = new Map();
        return this.notesCache;
      }
      console.error("❌ Error loading notes:", e);
      throw e;
    }
  }

  /** Get a single note from cache; returns an empty shell if missing. */
  getNote(barcode) {
    const key = this._normBarcode(barcode);
    return this.notesCache.get(key) || {
      barcode: key,
      meetingNote: "",
      requiresFollowUp: "",
      lastUpdated: null,
    };
  }

  /** Upsert a note in cache and queue it for save. */
  updateNote(barcode, meetingNote, requiresFollowUp) {
    const key = this._normBarcode(barcode);
    if (!key) return null;

    const note = {
      barcode: key,
      meetingNote: meetingNote || "",
      requiresFollowUp: requiresFollowUp || "",
      lastUpdated: new Date().toISOString(),
    };

    this.notesCache.set(key, note);
    this.saveQueue.set(key, note);
    return note;
  }

  /** Delete a note (queue the deletion). */
  deleteNote(barcode) {
    const key = this._normBarcode(barcode);
    if (!key) return;
    this.notesCache.delete(key);
    this.saveQueue.set(key, null); // mark for deletion
  }

  /** Persist queued changes to SharePoint (rebuilds & uploads the Excel). */
  async saveToSharePoint() {
    if (this.saveQueue.size === 0) {
      console.log("ℹ️ No changes to save.");
      return;
    }
    if (this.isSaving) {
      console.log("⏳ Save already in progress; skipping this call.");
      return;
    }

    this.isSaving = true;
    const batch = new Map(this.saveQueue);
    this.saveQueue.clear();

    try {
      console.log(`💾 Applying ${batch.size} change(s) to cache and uploading workbook…`);

      // Apply queued changes to the cache snapshot
      batch.forEach((note, bc) => {
        if (note === null) this.notesCache.delete(bc);
        else this.notesCache.set(bc, note);
      });

      // Convert cache to array of rows for Excel
      const rows = Array.from(this.notesCache.values()).map(n => ({
        "Barcode#": n.barcode,
        "Meeting Note": n.meetingNote,
        "Requires Follow Up": n.requiresFollowUp,
        "Last Updated": n.lastUpdated,
      }));

      // Build the worksheet with modest column widths (optional)
      const ws = XLSX.utils.json_to_sheet(rows, { skipHeader: false });
      ws["!cols"] = [
        { wch: 16 }, // Barcode#
        { wch: 60 }, // Meeting Note
        { wch: 28 }, // Requires Follow Up
        { wch: 24 }, // Last Updated
      ];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Repair Notes");
      const fileContent = XLSX.write(wb, { bookType: "xlsx", type: "array" }); // ArrayBuffer

      await this.ods.uploadExcelToSharePoint({
        hostname: this.config.spHostname,
        sitePath: this.config.spSitePath,
        fileRelativePath: this._filePath(),
        fileContent, // <-- IMPORTANT: OneDriveService expects fileContent
      });

      console.log(`✅ Saved ${batch.size} change(s) to ${this._filePath()}`);
    } catch (e) {
      console.error("❌ Error saving notes; re-queuing failed batch:", e);
      // Re-queue the failed changes
      batch.forEach((note, bc) => this.saveQueue.set(bc, note));
      throw e;
    } finally {
      this.isSaving = false;
    }
  }

  /** Bulk import notes (array of { barcode, meetingNote, requiresFollowUp, lastUpdated? }) and persist. */
  async importNotes(notesArray) {
    const list = Array.isArray(notesArray) ? notesArray : [];
    console.log(`📥 Importing ${list.length} note(s)…`);
    for (const n of list) {
      const shaped = this._noteShape(n);
      if (shaped.barcode) {
        this.updateNote(shaped.barcode, shaped.meetingNote, shaped.requiresFollowUp);
      }
    }
    await this.saveToSharePoint();
    console.log("✅ Import complete.");
  }

  /** Get all notes as a plain array. */
  getAllNotes() {
    return Array.from(this.notesCache.values());
  }
}
