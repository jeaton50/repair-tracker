// src/oneDriveService1.js
import * as XLSX from "xlsx";
import { ResponseType } from "@microsoft/microsoft-graph-client";

/**
 * OneDrive/SharePoint file reader & writer focused on SharePoint site libraries.
 * Requires delegated Microsoft Graph scopes. Minimum:
 *  - Files.Read.All (read)
 *  - Files.ReadWrite.All (upload/replace)
 *
 * Usage (read Excel):
 *   const ods = new OneDriveService(graphClient);
 *   const rows = await ods.readExcelFromSharePoint({
 *     hostname: "rentexinc.sharepoint.com",
 *     sitePath: "/sites/ProductManagers",
 *     fileRelativePath: "General/Repairs/RepairTracker/ticket_list.xlsx"
 *   });
 *
 * Usage (upload/replace Excel from an ArrayBuffer/Blob):
 *   await ods.uploadExcelToSharePoint({
 *     hostname: "rentexinc.sharepoint.com",
 *     sitePath: "/sites/ProductManagers",
 *     fileRelativePath: "General/Repairs/RepairTracker/ticket_list.xlsx",
 *     fileContent: arrayBufferOrBlob
 *   });
 *
 * Tip: For large files (>4MB), consider using upload sessions (createUploadSession).
 */
export default class OneDriveService {
  constructor(client) {
    this.client = client;
  }

  /* ----------------------------- SharePoint core ---------------------------- */

  // Resolve site + default "Documents" drive
  async _getSiteAndDefaultDrive(hostname, sitePath) {
    // Examples:
    //   hostname: "rentexinc.sharepoint.com"
    //   sitePath: "/sites/ProductManagers"
    const site = await this.client.api(`/sites/${hostname}:${sitePath}`).get();
    const drive = await this.client.api(`/sites/${site.id}/drive`).get(); // default document library
    return { site, drive };
  }

  // Get file content (binary) from SharePoint by relative path within the library
  async _getSpBinary({ hostname, sitePath, fileRelativePath }) {
    const { site, drive } = await this._getSiteAndDefaultDrive(hostname, sitePath);
    const encoded = fileRelativePath.split("/").map(encodeURIComponent).join("/");
    return this.client
      .api(`/sites/${site.id}/drives/${drive.id}/root:/${encoded}:/content`)
      .responseType(ResponseType.ARRAYBUFFER)
      .get();
  }

  // Get file content (text) from SharePoint by relative path within the library
  async _getSpText({ hostname, sitePath, fileRelativePath }) {
    const { site, drive } = await this._getSiteAndDefaultDrive(hostname, sitePath);
    const encoded = fileRelativePath.split("/").map(encodeURIComponent).join("/");
    return this.client
      .api(`/sites/${site.id}/drives/${drive.id}/root:/${encoded}:/content`)
      .responseType(ResponseType.TEXT)
      .get();
  }

  // Put file content (binary) to SharePoint by relative path within the library
  async _putSpBinary({ hostname, sitePath, fileRelativePath, fileContent }) {
    const { site, drive } = await this._getSiteAndDefaultDrive(hostname, sitePath);
    const encoded = fileRelativePath.split("/").map(encodeURIComponent).join("/");
    // For files <= ~4MB, simple PUT to /content is fine. For larger, use upload session.
    return this.client
      .api(`/sites/${site.id}/drives/${drive.id}/root:/${encoded}:/content`)
      .put(fileContent);
  }

  /* ------------------------------ Public API -------------------------------- */

  async readExcelFromSharePoint({ hostname, sitePath, fileRelativePath }) {
    const content = await this._getSpBinary({ hostname, sitePath, fileRelativePath });
    return this._xlsxToRows(content); // content is ArrayBuffer
  }

  async readJsonFromSharePoint({ hostname, sitePath, fileRelativePath }) {
    const text = await this._getSpText({ hostname, sitePath, fileRelativePath });
    return JSON.parse(text);
  }

  /**
   * Upload (create or replace) an Excel file at the given SharePoint path.
   * @param {Object} opts
   * @param {string} opts.hostname - e.g., "rentexinc.sharepoint.com"
   * @param {string} opts.sitePath - e.g., "/sites/ProductManagers"
   * @param {string} opts.fileRelativePath - e.g., "General/Repairs/RepairTracker/ticket_list.xlsx"
   * @param {ArrayBuffer|Blob|Uint8Array|ReadableStream} opts.fileContent - Excel binary to upload
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

  /* ------------------------------ Local helpers ----------------------------- */

  /**
   * Convert an XLSX ArrayBuffer in to an array of row objects.
   * Assumes: row 2 (index 1) contains headers, data starts at row 3.
   */
  _xlsxToRows(arrayBuffer) {
    const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });

    if (raw.length < 2) return [];

    const headers = (raw[1] || [])
      .map(h => String(h || "").trim())
      .filter(Boolean);

    return raw
      .slice(2)
      .filter(row => row && row.some(c => c !== "" && c != null))
      .map(row => Object.fromEntries(headers.map((h, i) => [h, row[i] ?? ""])));
  }

  /**
   * Optional helper: build an Excel workbook (ArrayBuffer) from either a 2D array (AOA)
   * or an array of objects (rows). Returns an ArrayBuffer ready for upload.
   * @param {Object} param0
   * @param {Array<Array<any>>} [param0.aoa] - 2D array including header row
   * @param {Array<Object>} [param0.rows] - array of objects (keys become headers)
   * @param {string} [param0.sheetName="Sheet1"]
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
    // Write as ArrayBuffer (type: "array")
    const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
    return out; // ArrayBuffer
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
