// src/oneDriveService.js
import * as XLSX from "xlsx";
import { ResponseType } from "@microsoft/microsoft-graph-client";

/**
 * OneDrive/SharePoint file reader focused on SharePoint site libraries.
 * Requires delegated Microsoft Graph scope: Files.Read.All (admin consent).
 *
 * Usage:
 *   const ods = new OneDriveService(graph);
 *   const rows = await ods.readExcelFromSharePoint({
 *     hostname: "rentexinc.sharepoint.com",
 *     sitePath: "/sites/ProductManagers",
 *     fileRelativePath: "General/Repairs/RepairTracker/ticket_list.xlsx"
 *   });
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
    const drive = await this.client.api(`/sites/${site.id}/drive`).get(); // default doc lib
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

  /* ------------------------------ Public API -------------------------------- */

  async readExcelFromSharePoint({ hostname, sitePath, fileRelativePath }) {
    const content = await this._getSpBinary({ hostname, sitePath, fileRelativePath });
    return this._xlsxToRows(content); // content is ArrayBuffer
  }

  async readJsonFromSharePoint({ hostname, sitePath, fileRelativePath }) {
    const text = await this._getSpText({ hostname, sitePath, fileRelativePath });
    return JSON.parse(text);
  }

  /* ------------------------------ Local helpers ----------------------------- */

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
}
