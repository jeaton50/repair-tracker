// src/SharePointNotesService.js
import * as XLSX from "xlsx";

/**
 * Service for reading/writing repair notes to a SharePoint Excel file via OneDriveService.
 * Expected columns:
 *   "Barcode#", "Meeting Note", "Requires Follow Up", "Last Updated"
 *
 * NOTE: OneDriveService._xlsxToRows() assumes:
 *   - headers on row 2
 *   - data starts row 3
 * This writer preserves that by inserting a blank first row.
 */
export default class SharePointNotesService {
  /**
   * @param {OneDriveService} oneDriveService
   * @param {{
   *   spHostname: string,
   *   spSitePath: string,
   *   spBasePath: string,       // e.g. "General/Repairs/RepairTracker"
   *   notesFile?: string        // default "repair_notes.xlsx"
   * }} config
   */
  constructor(oneDriveService, config) {
    this.ods = oneDriveService;
    this.config = {
      ...config,
      notesFile: config?.notesFile || "repair_notes.xlsx",
    };

    this.notesCache = new Map(); // { BARCODE -> {barcode, meetingNote, requiresFollowUp, lastUpdated} }
    this.saveQueue = new Map();  // { BARCODE -> note | null }
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

  _shapeFromRow(row) {
    // Accept either "Barcode#" or "Barcode" just in case
    const bc = row["Barcode#"] ?? row["Barcode"];
    return {
      barcode: this._normBarcode(bc),
      meetingNote: String(row["Meeting Note"] ?? "").trim(),
      requiresFollowUp: String(row["Requires Follow Up"] ?? "").trim(),
      lastUpdated: row["Last Updated"] || new Date().toISOString(),
    };
  }

  _shapeForCache(note) {
    return {
      barcode: this._normBarcode(note.barcode),
      meetingNote: String(note.meetingNote || "").trim(),
      requiresFollowUp: String(note.requiresFollowUp || "").trim(),
      lastUpdated: note.lastUpdated || new Date().toISOString(),
    };
  }

  /* --------------------------------- API ----------------------------------- */

  /**
   * Load all notes from SharePoint Excel into memory.
   * Uses OneDriveService.readExcelFromSharePoint (headers row 2, data row 3).
   */
  async loadAllNotes() {
    try {
      console.log("📥 Loading notes from SharePoint...");
      const rows = await this.ods.readExcelFromSharePoint({
        hostname: this.config.spHostname,
        sitePath: this.config.spSitePath,
        fileRelativePath: this._filePath(),
      });

      const map = new Map();
      (rows || []).forEach(r => {
        const shaped = this._shapeFromRow(r);
        if (shaped.barcode) map.set(shaped.barcode, shaped);
      });

      this.notesCache = map;
      console.log(`✅ Loaded ${map.size} notes from ${this._filePath()}`);
      return map;
    } catch (e) {
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

  /** Get a note from cache (or an empty shell). */
  getNote(barcode) {
    const key = this._normBarcode(barcode);
    return this.notesCache.get(key) || {
      barcode: key,
      meetingNote: "",
      requiresFollowUp: "",
      lastUpdated: null,
    };
  }

  /** Upsert note in cache and queue it for save. */
  updateNote(barcode, meetingNote, requiresFollowUp) {
    const key = this._normBarcode(barcode);
    if (!key) return null;

    const note = this._shapeForCache({ barcode: key, meetingNote, requiresFollowUp });
    this.notesCache.set(key, note);
    this.saveQueue.set(key, note);
    return note;
  }

  /** Delete a note (queue deletion). */
  deleteNote(barcode) {
    const key = this._normBarcode(barcode);
    if (!key) return;
    this.notesCache.delete(key);
    this.saveQueue.set(key, null);
  }

  /**
   * Persist queued changes to SharePoint:
   * - Ensures folders exist (no 404 on first save)
   * - Builds workbook with a BLANK row 1 and headers on row 2
   * - Uploads as Blob (browser-safe)
   */
  async saveToSharePoint() {
    if (this.saveQueue.size === 0) {
      console.log("ℹ️ No changes to save.");
      return;
    }
    if (this.isSaving) {
      console.log("⏳ Save already in progress; skipping.");
      return;
    }

    this.isSaving = true;
    const batch = new Map(this.saveQueue);
    this.saveQueue.clear();

    try {
      // Apply batch to cache snapshot
      batch.forEach((note, bc) => {
        if (note === null) this.notesCache.delete(bc);
        else this.notesCache.set(bc, note);
      });

      // Convert cache to array of row objects for Excel
      const rows = Array.from(this.notesCache.values()).map(n => ({
        "Barcode#": n.barcode,
        "Meeting Note": n.meetingNote,
        "Requires Follow Up": n.requiresFollowUp,
        "Last Updated": n.lastUpdated,
      }));

      // ---- Build workbook with blank row1 + headers row2 (for your reader) ----
      const headers = ["Barcode#", "Meeting Note", "Requires Follow Up", "Last Updated"];
      const dataRows = rows.map(r => [
        r["Barcode#"], r["Meeting Note"], r["Requires Follow Up"], r["Last Updated"]
      ]);
      const aoa = [
        [],          // row 1: blank
        headers,     // row 2: headers
        ...dataRows, // row 3+: data
      ];

      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws["!cols"] = [
        { wch: 16 }, // Barcode#
        { wch: 60 }, // Meeting Note
        { wch: 28 }, // Requires Follow Up
        { wch: 24 }, // Last Updated
      ];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Repair Notes");

      // ArrayBuffer -> Blob (browser-safe; avoids Node Buffer)
      const excelArrayBuffer = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const fileContent = new Blob([excelArrayBuffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      // Ensure folder chain exists before upload
      const folderPath = (this.config.spBasePath || "").replace(/\/+$/, "");
      if (folderPath) {
        await this.ods.ensureFolderPath({
          hostname: this.config.spHostname,
          sitePath: this.config.spSitePath,
          folderRelativePath: folderPath,
        });
      }

      // Upload
      await this.ods.uploadExcelToSharePoint({
        hostname: this.config.spHostname,
        sitePath: this.config.spSitePath,
        fileRelativePath: this._filePath(),
        fileContent,
      });

      console.log(`✅ Saved ${batch.size} change(s) to ${this._filePath()}`);
    } catch (e) {
      console.error("❌ Error saving notes; re-queuing failed batch:", e);
      // Re-queue failed changes so user can try again
      batch.forEach((note, bc) => this.saveQueue.set(bc, note));
      throw e;
    } finally {
      this.isSaving = false;
    }
  }

  /** Bulk import notes and persist. */
  async importNotes(notesArray) {
    const list = Array.isArray(notesArray) ? notesArray : [];
    for (const n of list) {
      const shaped = this._shapeForCache(n);
      if (shaped.barcode) {
        this.notesCache.set(shaped.barcode, shaped);
        this.saveQueue.set(shaped.barcode, shaped);
      }
    }
    await this.saveToSharePoint();
  }

  /** Get all notes as a plain array. */
  getAllNotes() {
    return Array.from(this.notesCache.values());
  }
}
