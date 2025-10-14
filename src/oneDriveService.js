// src/oneDriveService.js
import * as XLSX from "xlsx";
import { ResponseType } from "@microsoft/microsoft-graph-client";

/**
 * OneDrive/SharePoint file reader & writer focused on SharePoint site libraries.
 * Scopes needed:
 *  - Files.Read.All (read)
 *  - Files.ReadWrite.All (upload/replace)
 *
 * Example (read Excel):
 *   const ods = new OneDriveService(graphClient);
 *   const rows = await ods.readExcelFromSharePoint({
 *     hostname: "rentexinc.sharepoint.com",
 *     sitePath: "/sites/ProductManagers",
 *     fileRelativePath: "General/Repairs/RepairTracker/ticket_list.xlsx"
 *   });
 *
 * Example (upload Excel from ArrayBuffer/Blob):
 *   await ods.uploadExcelToSharePoint({
 *     hostname: "rentexinc.sharepoint.com",
 *     sitePath: "/sites/ProductManagers",
 *     fileRelativePath: "General/Repairs/RepairTracker/ticket_list.xlsx",
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
    // hostname: "rentexinc.sharepoint.com"
    // sitePath: "/sites/ProductManagers"
    const siteUrl = `/sites/${hostname}:${sitePath}`;
    try {
      const site = await this.client.api(siteUrl).get();
      const drive = await this.client.api(`/sites/${site.id}/drive`).get(); // default doc lib
      return { site, drive };
    } catch (e) {
      console.error("Failed to resolve site/drive", {
        hostname,
        sitePath,
        status: e?.statusCode,
        code: e?.code,
        message: e?.message,
      });
      throw e;
    }
  }

  // Get file content (binary) from SharePoint by relative path within the library
  async _getSpBinary({ hostname, sitePath, fileRelativePath }) {
    const { site, drive } = await this._getSiteAndDefaultDrive(hostname, sitePath);
    const encoded = fileRelativePath.split("/").map(encodeURIComponent).join("/");
    const url = `/sites/${site.id}/drives/${drive.id}/root:/${encoded}:/content`;
    try {
      return await this.client.api(url).responseType(ResponseType.ARRAYBUFFER).get();
    } catch (e) {
      console.error("Graph _getSpBinary error", {
        url,
        status: e?.statusCode,
        code: e?.code,
        message: e?.message,
        body: e?.body,
      });
      throw e;
    }
  }

  // Get file content (text) from SharePoint by relative path within the library
  async _getSpText({ hostname, sitePath, fileRelativePath }) {
    const { site, drive } = await this._getSiteAndDefaultDrive(hostname, sitePath);
    const encoded = fileRelativePath.split("/").map(encodeURIComponent).join("/");
    const url = `/sites/${site.id}/drives/${drive.id}/root:/${encoded}:/content`;
    try {
      return await this.client.api(url).responseType(ResponseType.TEXT).get();
    } catch (e) {
      console.error("Graph _getSpText error", {
        url,
        status: e?.statusCode,
        code: e?.code,
        message: e?.message,
        body: e?.body,
      });
      throw e;
    }
  }

  // Put file content (binary) to SharePoint by relative path within the library
  async _putSpBinary({ hostname, sitePath, fileRelativePath, fileContent }) {
    const { site, drive } = await this._getSiteAndDefaultDrive(hostname, sitePath);
    const encoded = fileRelativePath.split("/").map(encodeURIComponent).join("/");
    const url = `/sites/${site.id}/drives/${drive.id}/root:/${encoded}:/content`;
    try {
      return await this.client.api(url).put(fileContent);
    } catch (e) {
      console.error("Graph _putSpBinary error", {
        url,
        status: e?.statusCode,
        code: e?.code,
        message: e?.message,
        body: e?.body,
      });
      throw e;
    }
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
   * @param {string} opts.hostname
   * @param {string} opts.sitePath
   * @param {string} opts.fileRelativePath
   * @param {ArrayBuffer|Blob|Uint8Array|ReadableStream} opts.fileContent
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
   * Convert an XLSX ArrayBuffer into an array of row objects.
   * Tries row 2 as headers (your original layout), falls back to row 1 if needed.
   */
  _xlsxToRows(arrayBuffer) {
    const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });

    if (!raw.length) return [];

    // Prefer row 2 as headers; if empty, use row 1
    const headerRow = (raw[1] && raw[1].some(v => v !== "")) ? raw[1] : raw[0];
    const startIdx = headerRow === raw[1] ? 2 : 1;

    const headers = (headerRow || [])
      .map(h => String(h || "").trim())
      .filter(Boolean);

    if (!headers.length) return [];

    return raw
      .slice(startIdx)
      .filter(row => row && row.some(c => c !== "" && c != null))
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

  /* ---------------------- Large file uploads (optional) --------------------- */
  /**
   * For files >4MB, use an upload session:
   *
   *   const { site, drive } = await this._getSiteAndDefaultDrive(hostname, sitePath);
   *   const encoded = fileRelativePath.split("/").map(encodeURIComponent).join("/");
   *   const session = await this.client
   *     .api(`/sites/${site.id}/drives/${drive.id}/root:/${encoded}:/createUploadSession`)
   *     .post({ item: { "@microsoft.graph.conflictBehavior": "replace" }});
   *   // then PUT chunks to session.uploadUrl
   */
}
