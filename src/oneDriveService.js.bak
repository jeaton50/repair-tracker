// src/oneDriveService.js
import * as XLSX from "xlsx";
import { ResponseType } from "@microsoft/microsoft-graph-client";

/**
 * OneDrive/SharePoint file reader & writer focused on SharePoint site libraries.
 * Requires delegated Microsoft Graph scopes. Minimum:
 *  - Files.Read.All (read)
 *  - Files.ReadWrite.All (upload/replace)
 *
 * Example (read Excel - strict):
 *   const ods = new OneDriveService(graphClient);
 *   const rows = await ods.readExcelFromSharePoint({
 *     hostname: "rentexinc.sharepoint.com",
 *     sitePath: "/sites/ProductManagers",
 *     fileRelativePath: "General/Repairs/RepairTracker/ticket_list.xlsx"
 *   });
 *
 * Example (read Excel - flexible):
 *   const rows = await ods.readExcelFromSharePointFlexible({ ... });
 *
 * Example (upload Excel from ArrayBuffer/Blob):
 *   await ods.uploadExcelToSharePoint({
 *     hostname: "rentexinc.sharepoint.com",
 *     sitePath: "/sites/ProductManagers",
 *     fileRelativePath: "General/Repairs/RepairTracker/repair_notes.xlsx",
 *     fileContent: arrayBufferOrBlob
 *   });
 */
export default class OneDriveService {
  constructor(client) {
    this.client = client;
  }

  /* ----------------------------- SharePoint core ---------------------------- */

  // Resolve site + default "Documents" drive
  async _getSiteAndDefaultDrive(hostname, sitePath) {
    const site = await this.client.api(`/sites/${hostname}:${sitePath}`).get();
    const drive = await this.client.api(`/sites/${site.id}/drive`).get(); // default document library
    return { site, drive };
  }

  // Ensure folder chain exists under the site's default Documents drive
  async ensureFolderPath({ hostname, sitePath, folderRelativePath }) {
    const { site, drive } = await this._getSiteAndDefaultDrive(hostname, sitePath);
    const segments = (folderRelativePath || "").split("/").filter(Boolean);
    let currentPath = "";
    for (const seg of segments) {
      currentPath = currentPath ? `${currentPath}/${seg}` : seg;
      const encoded = currentPath.split("/").map(encodeURIComponent).join("/");
      try {
        // If the folder exists this will succeed
        await this.client.api(`/sites/${site.id}/drives/${drive.id}/root:/${encoded}`).get();
      } catch {
        // Create folder under its parent
        const parent = currentPath.includes("/") ? currentPath.split("/").slice(0, -1).join("/") : "";
        const parentEncoded = parent ? parent.split("/").map(encodeURIComponent).join("/") : "";
        await this.client
          .api(`/sites/${site.id}/drives/${drive.id}/root:/${parentEncoded}:/children`)
          .post({
            name: seg,
            folder: {},
            "@microsoft.graph.conflictBehavior": "replace",
          });
      }
    }
  }

  // Get file content (binary) from SharePoint by relative path within the library
  async _getSpBinary({ hostname, sitePath, fileRelativePath }) {
    const { site, drive } = await this._getSiteAndDefaultDrive(hostname, sitePath);
    const encoded = fileRelativePath.split("/").map(encodeURIComponent).join("/");
    const url = `/sites/${site.id}/drives/${drive.id}/root:/${encoded}:/content`;
    return this.client.api(url).responseType(ResponseType.ARRAYBUFFER).get();
  }

  // Get file content (text) from SharePoint by relative path within the library
  async _getSpText({ hostname, sitePath, fileRelativePath }) {
    const { site, drive } = await this._getSiteAndDefaultDrive(hostname, sitePath);
    const encoded = fileRelativePath.split("/").map(encodeURIComponent).join("/");
    const url = `/sites/${site.id}/drives/${drive.id}/root:/${encoded}:/content`;
    return this.client.api(url).responseType(ResponseType.TEXT).get();
  }

  // Put file content (binary) to SharePoint by relative path within the library
  async _putSpBinary({ hostname, sitePath, fileRelativePath, fileContent }) {
    const { site, drive } = await this._getSiteAndDefaultDrive(hostname, sitePath);
    const encoded = fileRelativePath.split("/").map(encodeURIComponent).join("/");
    const url = `/sites/${site.id}/drives/${drive.id}/root:/${encoded}:/content`;

    // In browsers, Graph prefers Blob; avoid Node Buffer (not available)
    let body = fileContent;
    const isBrowser = typeof window !== "undefined" && typeof window.Blob !== "undefined";
    const mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    if (isBrowser) {
      if (fileContent instanceof ArrayBuffer) {
        body = new Blob([fileContent], { type: mime });
      } else if (ArrayBuffer.isView(fileContent)) {
        body = new Blob([fileContent.buffer], { type: mime });
      }
    }

    return this.client.api(url).put(body);
  }

  /* ------------------------------ Public API -------------------------------- */

  // STRICT reader (headers on row 2, data from row 3) — use for tickets/reports
  async readExcelFromSharePoint({ hostname, sitePath, fileRelativePath }) {
    const content = await this._getSpBinary({ hostname, sitePath, fileRelativePath });
    return this._xlsxToRows(content);
  }

  // FLEXIBLE reader (first non-empty row is headers) — use for notes
  async readExcelFromSharePointFlexible({ hostname, sitePath, fileRelativePath }) {
    const content = await this._getSpBinary({ hostname, sitePath, fileRelativePath });
    return this._xlsxToRowsFlexible(content);
  }

  async readJsonFromSharePoint({ hostname, sitePath, fileRelativePath }) {
    const text = await this._getSpText({ hostname, sitePath, fileRelativePath });
    return JSON.parse(text);
  }

  /**
   * Upload (create or replace) an Excel file at the given SharePoint path.
   */
  async uploadExcelToSharePoint({ hostname, sitePath, fileRelativePath, fileContent }) {
    try {
      await this._putSpBinary({ hostname, sitePath, fileRelativePath, fileContent });
      console.log(`✅ Uploaded Excel to: ${fileRelativePath}`);
    } catch (error) {
      console.error(`❌ Error uploading Excel to SharePoint: ${fileRelativePath}`, error);
      throw error;
    }
  }

  // Optional: quick probe to verify connectivity & IDs
  async debugProbe({ hostname, sitePath }) {
    const { site, drive } = await this._getSiteAndDefaultDrive(hostname, sitePath);
    console.info("SharePoint probe OK", { siteId: site.id, driveId: drive.id });
    return { site, drive };
  }

  /* ------------------------------ Local helpers ----------------------------- */

  /**
   * STRICT parse:
   *   - Row 2 (index 1) is headers
   *   - Data starts at row 3
   * Used by tickets/reports — DO NOT CHANGE to avoid regressions.
   */
  _xlsxToRows(arrayBuffer) {
    const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });

    if (raw.length < 2) return [];

    const headers = (raw[1] || [])
      .map(h => String(h || "").trim())
      .filter(Boolean);

    if (!headers.length) return [];

    return raw
      .slice(2)
      .filter(row => row && row.some(c => c !== "" && c != null && String(c).trim() !== ""))
      .map(row => Object.fromEntries(headers.map((h, i) => [h, row[i] ?? ""])));
  }

  /**
   * FLEXIBLE parse:
   *   - First non-empty row is treated as headers
   *   - Works if row 1 is blank or not
   * Recommended for ad-hoc/notes sheets that may not follow the strict layout.
   */
  _xlsxToRowsFlexible(arrayBuffer) {
    const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });

    if (!raw.length) return [];

    const headerIdx = raw.findIndex(
      r => Array.isArray(r) && r.some(c => c != null && String(c).trim() !== "")
    );
    if (headerIdx === -1) return [];

    const headers = (raw[headerIdx] || [])
      .map(h => String(h || "").trim())
      .filter(Boolean);

    if (!headers.length) return [];

    const startIdx = headerIdx + 1;

    return raw
      .slice(startIdx)
      .filter(row => row && row.some(c => c != null && String(c).trim() !== ""))
      .map(row => Object.fromEntries(headers.map((h, i) => [h, row[i] ?? ""])));
  }

  /**
   * Build an Excel workbook (ArrayBuffer) from either a 2D array (AOA)
   * or an array of objects (rows). Returns an ArrayBuffer ready for upload.
   */
  buildExcelArrayBuffer({ aoa, rows, sheetName = "Sheet1" } = {}) {
    const wb = XLSX.utils.book_new();
    let ws;

    if (Array.isArray(aoa)) {
      ws = XLSX.utils.aoa_to_sheet(aoa);
    } else if (Array.isArray(rows)) {
      ws = XLSX.utils.json_to_sheet(rows, { skipHeader: false });
    } else {
      ws = XLSX.utils.aoa_to_sheet([["(empty)"]]);
    }

    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    return XLSX.write(wb, { bookType: "xlsx", type: "array" }); // ArrayBuffer
  }

  /**
   * Build an Excel workbook as a Blob (best for browser uploads).
   */
  buildExcelBlob({ aoa, rows, sheetName = "Sheet1" } = {}) {
    const arrayBuf = this.buildExcelArrayBuffer({ aoa, rows, sheetName });
    return new Blob([arrayBuf], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
  }

  /* ---------------------- Large file uploads (optional) --------------------- */
  /**
   * For files >4MB, you can use an upload session:
   *   const { site, drive } = await this._getSiteAndDefaultDrive(hostname, sitePath);
   *   const encoded = fileRelativePath.split("/").map(encodeURIComponent).join("/");
   *   const session = await this.client
   *     .api(`/sites/${site.id}/drives/${drive.id}/root:/${encoded}:/createUploadSession`)
   *     .post({ item: { "@microsoft.graph.conflictBehavior": "replace" }});
   *   // then chunked PUTs to session.uploadUrl
   */
}
