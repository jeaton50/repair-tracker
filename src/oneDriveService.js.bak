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

  // Ensure folder chain exists under the site's default Documents drive
  async ensureFolderPath({ hostname, sitePath, folderRelativePath }) {
    const { site, drive } = await this._getSiteAndDefaultDrive(hostname, sitePath);
    const segments = (folderRelativePath || "").split("/").filter(Boolean);
    let currentPath = "";
    for (const seg of segments) {
      currentPath = currentPath ? `${currentPath}/${seg}` : seg;
      const encoded = currentPath.split("/").map(encodeURIComponent).join("/");
      const metaUrl = `/sites/${site.id}/drives/${drive.id}/root:/${encoded}`;
      try {
        await this.client.api(metaUrl).get(); // exists
      } catch {
        // create this level
        await this.client
          .api(`/sites/${site.id}/drives/${drive.id}/root:/${encoded.replace(/\/?[^/]+$/, "") || ""}:/children`)
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

  async readExcelFromSharePoint({
    hostname,
    sitePath,
    fileRelativePath,
    sheetName,         // optional: preferred sheet to read
    expectedHeaders,   // optional: list of headers to match
  }) {
    const content = await this._getSpBinary({ hostname, sitePath, fileRelativePath });
    return this._xlsxToRows(content, { sheetName, expectedHeaders }); // content is ArrayBuffer
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
   * Convert an XLSX ArrayBuffer into array of row objects.
   * - Auto-detects header row within the first 5 rows.
   * - Optionally targets a specific sheet and set of expected headers.
   */
  _xlsxToRows(arrayBuffer, opts = {}) {
    const expected = (opts.expectedHeaders && opts.expectedHeaders.length)
      ? opts.expectedHeaders
      : ["Barcode#", "Meeting Note", "Requires Follow Up", "Last Updated"];

    const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });

    const scanSheet = (sheet) => {
      const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
      if (!raw.length) return [];
      // search first 5 rows for a header row containing all expected headers (case-insensitive)
      for (let i = 0; i < Math.min(5, raw.length); i++) {
        const hdr = (raw[i] || []).map(h => String(h || "").trim());
        const norm = hdr.map(h => h.toLowerCase());
        const ok = expected.every(e => norm.includes(String(e).toLowerCase()));
        if (ok) {
          const start = i + 1;
          return raw
            .slice(start)
            .filter(r => r && r.some(c => c !== "" && c != null))
            .map(r => Object.fromEntries(hdr.map((h, idx) => [h, r[idx] ?? ""])));
        }
      }
      return [];
    };

    // Try preferred sheet name first (if provided), then fall back to any sheet that matches.
    if (opts.sheetName && wb.Sheets[opts.sheetName]) {
      const rows = scanSheet(wb.Sheets[opts.sheetName]);
      if (rows.length) return rows;
    }
    for (const name of wb.SheetNames) {
      const rows = scanSheet(wb.Sheets[name]);
      if (rows.length) return rows;
    }
    return [];
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
   * const { site, drive } = await this._getSiteAndDefaultDrive(hostname, sitePath);
   * const encoded = fileRelativePath.split("/").map(encodeURIComponent).join("/");
   * const session = await this.client
   *   .api(`/sites/${site.id}/drives/${drive.id}/root:/${encoded}:/createUploadSession`)
   *   .post({ item: { "@microsoft.graph.conflictBehavior": "replace" }});
   * // then PUT chunks to session.uploadUrl
   */
}
