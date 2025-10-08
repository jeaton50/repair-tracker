// src/oneDriveService.js
import * as XLSX from "xlsx";
import { ResponseType } from "@microsoft/microsoft-graph-client";

// ... keep encodePath / isNotFound helpers if you added them ...

export default class OneDriveService {
  constructor(client, folderPath = "RepairTracker") {
    this.client = client;
    this.folderPath = folderPath;
  }

  // ------- BINARY (Excel) fetch by path -------
  async getBinaryByPath(filePath) {
    const path = filePath.includes("/") ? filePath : `${this.folderPath}/${filePath}`;
    const encoded = path.split("/").map(encodeURIComponent).join("/");
    return this.client
      .api(`/me/drive/root:/${encoded}:/content`)
      .responseType(ResponseType.ARRAYBUFFER)   // 👈 force ArrayBuffer
      .get();
  }

  // ------- BINARY (Excel) fetch by id/remoteItem -------
  async getBinaryByItem(itemOrId) {
    try {
      const id = typeof itemOrId === "string" ? itemOrId : itemOrId.id;
      return await this.client
        .api(`/me/drive/items/${id}/content`)
        .responseType(ResponseType.ARRAYBUFFER) // 👈 force ArrayBuffer
        .get();
    } catch (err) {
      const item = typeof itemOrId === "string" ? null : itemOrId;
      const driveId = item?.remoteItem?.parentReference?.driveId;
      const remoteId = item?.remoteItem?.id;
      if (driveId && remoteId) {
        return this.client
          .api(`/drives/${driveId}/items/${remoteId}/content`)
          .responseType(ResponseType.ARRAYBUFFER) // 👈 force ArrayBuffer
          .get();
      }
      throw err;
    }
  }

  // ------- TEXT/JSON fetch by path -------
  async getTextByPath(filePath) {
    const path = filePath.includes("/") ? filePath : `${this.folderPath}/${filePath}`;
    const encoded = path.split("/").map(encodeURIComponent).join("/");
    return this.client
      .api(`/me/drive/root:/${encoded}:/content`)
      .responseType(ResponseType.TEXT)          // 👈 force text
      .get();
  }

  // ------- TEXT/JSON fetch by id/remoteItem -------
  async getTextByItem(itemOrId) {
    try {
      const id = typeof itemOrId === "string" ? itemOrId : itemOrId.id;
      return await this.client
        .api(`/me/drive/items/${id}/content`)
        .responseType(ResponseType.TEXT)        // 👈 force text
        .get();
    } catch (err) {
      const item = typeof itemOrId === "string" ? null : itemOrId;
      const driveId = item?.remoteItem?.parentReference?.driveId;
      const remoteId = item?.remoteItem?.id;
      if (driveId && remoteId) {
        return this.client
          .api(`/drives/${driveId}/items/${remoteId}/content`)
          .responseType(ResponseType.TEXT)      // 👈 force text
          .get();
      }
      throw err;
    }
  }

  // ---------------- public APIs ----------------
  async readExcelFileShared(fileName) {
    let content;
    try {
      content = await this.getBinaryByPath(fileName);
    } catch (err) {
      const item = await this.searchSharedFile(fileName);
      content = await this.getBinaryByItem(item);
    }
    return this._xlsxToRows(content); // already an ArrayBuffer
  }

  async readJsonFileShared(fileName) {
    let text;
    try {
      text = await this.getTextByPath(fileName);
    } catch (err) {
      const item = await this.searchSharedFile(fileName);
      text = await this.getTextByItem(item);
    }
    return JSON.parse(text);
  }

  // unchanged: searchSharedFile + parsers
  async searchSharedFile(fileName) {
    const resp = await this.client.api("/me/drive/sharedWithMe").get();
    let hit = resp.value.find(x => x.name === fileName && !x.folder);
    if (!hit) hit = resp.value.find(x => x.name?.toLowerCase() === fileName.toLowerCase() && !x.folder);
    if (!hit) throw new Error(`File not found in "Shared with me": ${fileName}`);
    return hit;
  }

  _xlsxToRows(arrayBuffer) {
    const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", raw: false });
    if (raw.length < 2) return [];
    const headers = (raw[1] || []).map(h => String(h || "").trim()).filter(Boolean);
    return raw
      .slice(2)
      .filter(row => row && row.some(c => c !== "" && c != null))
      .map(row => Object.fromEntries(headers.map((h, i) => [h, row[i] ?? ""])));
  }
}
