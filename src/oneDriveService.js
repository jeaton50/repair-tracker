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
        // Create this level (POST to parent children)
        const parentPath = currentPath.includes("/") ? currentPath.split("/").slice(0, -1).join("/") : "";
        const parentEncoded = parentPath ? parentPath.split("/").map(encodeURIComponent).join("/") : "";
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
      // In browsers, upload as Blob (avoid Node Buffer)
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
      return await this.client.api(url).put(body);
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
    expectedHeaders,   // optional: headers to match (omit for tickets/reports)
  }) {
    const content = await this._getSpBinary({ hostname, sitePath, fileRelativePath });
    return this._xlsxToRows(content, { sheetName, expectedHeaders }); // ArrayBuffer -> rows
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
   * Convert an XLSX ArrayBuffer into array of row objects.
   * - If `expectedHeaders` is provided: find that header row (first 5 rows).
   * - Otherwise: use the first non-empty row as headers (generic mode).
   * - If `sheetName` is provided, try that first; otherwise scan all sheets.
   */
  _xlsxToRows(arrayBuffer, opts = {}) {
    const expected = (opts.expectedHeaders && opts.expectedHeaders.length)
      ? opts.expectedHeaders
      : null; // generic mode by default

    const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });

    const scanSheet = (sheet) => {
      const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
      if (!raw.length) return [];

      // Header-match mode (when expected headers are provided)
      if (expected) {
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
        // If nothing matched, fall through to generic parsing below.
      }

      // Generic mode: pick the first non-empty row as headers.
      const headerIdx = raw.findIndex(r => Array.isArray(r) && r.some(c => c !== "" && c != null));
      if (headerIdx === -1) return [];
      const hdr = (raw[headerIdx] || []).map(h => String(h || "").trim());
      const start = headerIdx + 1;
      return raw
        .slice(start)
        .filter(r => r && r.some(c => c !== "" && c != null))
        .map(r => Object.fromEntries(hdr.map((h, idx) => [h, r[idx] ?? ""])));
    };

    // Try preferred sheet first (if provided), then any sheet that matches.
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
   * Build an Excel workbook as ArrayBuffer (good for Node or to wrap in a Blob).
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
