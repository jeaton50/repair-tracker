// src/oneDriveService.js
import * as XLSX from "xlsx";
import { ResponseType } from "@microsoft/microsoft-graph-client";

/**
 * OneDrive/SharePoint file reader & writer focused on SharePoint site libraries.
 * Scopes needed:
 *  - Files.Read.All (read)
 *  - Files.ReadWrite.All (upload/replace)
 */
export default class OneDriveService {
  constructor(client) {
    this.client = client;
  }

  /* ----------------------------- SharePoint core ---------------------------- */

  // Resolve site + default "Documents" drive
  async _getSiteAndDefaultDrive(hostname, sitePath) {
    const site = await this.client.api(`/sites/${hostname}:${sitePath}`).get();
    const drive = await this.client.api(`/sites/${site.id}/drive`).get(); // default doc lib
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
        await this.client.api(`/sites/${site.id}/drives/${drive.id}/root:/${encoded}`).get(); // exists
      } catch {
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
    return this._xlsxToRows_STRICT(content);
  }

  // FLEXIBLE reader scanning ALL SHEETS — recommended for notes
  async readExcelFromSharePointFlexible({ hostname, sitePath, fileRelativePath, preferredSheetName }) {
    const content = await this._getSpBinary({ hostname, sitePath, fileRelativePath });
    return this._xlsxToRows_FLEX(content, preferredSheetName);
  }

  // Convenience wrapper specifically for notes (with smart defaults)
  async readExcelForNotes({ hostname, sitePath, fileRelativePath, preferredSheetName = "Repair Notes" }) {
    return this.readExcelFromSharePointFlexible({
      hostname,
      sitePath,
      fileRelativePath,
      preferredSheetName,
    });
  }

  async readJsonFromSharePoint({ hostname, sitePath, fileRelativePath }) {
    const text = await this._getSpText({ hostname, sitePath, fileRelativePath });
    return JSON.parse(text);
  }

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
   * STRICT parse used by tickets/reports:
   *   - Row 2 (index 1) is headers
   *   - Data starts at row 3
   */
  _xlsxToRows_STRICT(arrayBuffer) {
    const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
    const firstSheet = wb.SheetNames[0];
    const sheet = wb.Sheets[firstSheet];
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });

    if (raw.length < 2) return [];

    const headers = (raw[1] || [])
      .map(h => String(h || "").trim())
      .filter(Boolean);

    if (!headers.length) return [];

    return raw
      .slice(2)
      .filter(row => row && row.some(c => c != null && String(c).trim() !== ""))
      .map(row => Object.fromEntries(headers.map((h, i) => [h, row[i] ?? ""])));
  }

  /**
   * FLEXIBLE parse:
   * - Searches a preferred sheet first (if provided), then all sheets.
   * - For each sheet, finds the first non-empty row as headers.
   * - Picks the sheet whose headers include "Barcode#" or "Barcode" if possible.
   * - Falls back to the first parsable sheet if nothing matches the heuristic.
   */
  _xlsxToRows_FLEX(arrayBuffer, preferredSheetName) {
    const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });

    const tryParseSheet = (sheetName) => {
      const sheet = wb.Sheets[sheetName];
      if (!sheet) return null;
      const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
      if (!raw.length) return null;

      const headerIdx = raw.findIndex(r =>
        Array.isArray(r) && r.some(c => c != null && String(c).trim() !== "")
      );
      if (headerIdx === -1) return null;

      const headers = (raw[headerIdx] || [])
        .map(h => String(h || "").trim())
        .filter(Boolean);
      if (!headers.length) return null;

      const startIdx = headerIdx + 1;
      const rows = raw
        .slice(startIdx)
        .filter(row => row && row.some(c => c != null && String(c).trim() !== ""))
        .map(row => Object.fromEntries(headers.map((h, i) => [h, row[i] ?? ""])));

      return { sheetName, headers, rows };
    };

    // 1) Try the preferred sheet first (if any)
    if (preferredSheetName && wb.SheetNames.includes(preferredSheetName)) {
      const r = tryParseSheet(preferredSheetName);
      if (r && r.rows.length >= 0) return r.rows; // even empty is a valid parse
    }

    // 2) Try all sheets and rank by header match quality
    const candidates = [];
    for (const s of wb.SheetNames) {
      const r = tryParseSheet(s);
      if (r) {
        const headersUpper = r.headers.map(h => h.toUpperCase());
        const hasBarcode = headersUpper.includes("BARCODE#") || headersUpper.includes("BARCODE");
        const score = (hasBarcode ? 1000 : 0) + r.headers.length; // simple heuristic
        candidates.push({ score, ...r });
      }
    }

    if (!candidates.length) return [];

    // Prefer sheets with barcode-like headers; otherwise the widest header set
    candidates.sort((a, b) => b.score - a.score);
    return candidates[0].rows;
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
