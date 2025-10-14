// src/oneDriveService.js - UPDATED with Upload Support
import * as XLSX from "xlsx";

class OneDriveService {
  constructor(graphClient) {
    this.graphClient = graphClient;
  }

  // ========== EXISTING READ METHODS ==========

  async readExcelFromSharePoint({ hostname, sitePath, fileRelativePath }) {
    try {
      const url = `/sites/${hostname}:${sitePath}:/drive/root:${fileRelativePath}:/content`;
      const fileStream = await this.graphClient.api(url).get();
      const buffer = await this.streamToBuffer(fileStream);
      const workbook = XLSX.read(buffer, { type: "buffer" });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      return XLSX.utils.sheet_to_json(sheet);
    } catch (error) {
      console.error(`Error reading Excel from SharePoint: ${fileRelativePath}`, error);
      throw error;
    }
  }

  async readJsonFromSharePoint({ hostname, sitePath, fileRelativePath }) {
    try {
      const url = `/sites/${hostname}:${sitePath}:/drive/root:${fileRelativePath}:/content`;
      const fileStream = await this.graphClient.api(url).get();
      const buffer = await this.streamToBuffer(fileStream);
      const text = buffer.toString("utf-8");
      return JSON.parse(text);
    } catch (error) {
      console.error(`Error reading JSON from SharePoint: ${fileRelativePath}`, error);
      throw error;
    }
  }

  // ========== NEW UPLOAD METHODS ==========

  /**
   * Upload Excel file to SharePoint
   * @param {Object} params
   * @param {string} params.hostname - SharePoint hostname
   * @param {string} params.sitePath - Site path
   * @param {string} params.fileRelativePath - File path (e.g., /Shared Documents/notes.xlsx)
   * @param {ArrayBuffer|Buffer} params.data - Excel data as buffer
   */
  async uploadExcelToSharePoint({ hostname, sitePath, fileRelativePath, data }) {
    try {
      const url = `/sites/${hostname}:${sitePath}:/drive/root:${fileRelativePath}:/content`;
      
      await this.graphClient
        .api(url)
        .putStream(data);

      console.log(`✅ Uploaded Excel to: ${fileRelativePath}`);
    } catch (error) {
      console.error(`Error uploading Excel to SharePoint: ${fileRelativePath}`, error);
      throw error;
    }
  }

  /**
   * Upload JSON file to SharePoint
   */
  async uploadJsonToSharePoint({ hostname, sitePath, fileRelativePath, data }) {
    try {
      const url = `/sites/${hostname}:${sitePath}:/drive/root:${fileRelativePath}:/content`;
      const jsonString = JSON.stringify(data, null, 2);
      const buffer = Buffer.from(jsonString, "utf-8");

      await this.graphClient
        .api(url)
        .header("Content-Type", "application/json")
        .putStream(buffer);

      console.log(`✅ Uploaded JSON to: ${fileRelativePath}`);
    } catch (error) {
      console.error(`Error uploading JSON to SharePoint: ${fileRelativePath}`, error);
      throw error;
    }
  }

  /**
   * Check if file exists on SharePoint
   */
  async fileExists({ hostname, sitePath, fileRelativePath }) {
    try {
      const url = `/sites/${hostname}:${sitePath}:/drive/root:${fileRelativePath}`;
      await this.graphClient.api(url).get();
      return true;
    } catch (error) {
      if (error.statusCode === 404) {
        return false;
      }
      throw error;
    }
  }

  /**
   * Delete file from SharePoint
   */
  async deleteFile({ hostname, sitePath, fileRelativePath }) {
    try {
      const url = `/sites/${hostname}:${sitePath}:/drive/root:${fileRelativePath}`;
      await this.graphClient.api(url).delete();
      console.log(`🗑️ Deleted file: ${fileRelativePath}`);
    } catch (error) {
      console.error(`Error deleting file from SharePoint: ${fileRelativePath}`, error);
      throw error;
    }
  }

  // ========== HELPER METHODS ==========

  async streamToBuffer(stream) {
    if (Buffer.isBuffer(stream)) {
      return stream;
    }
    if (stream instanceof ArrayBuffer) {
      return Buffer.from(stream);
    }
    const chunks = [];
    return new Promise((resolve, reject) => {
      stream.on("data", (chunk) => chunks.push(Buffer.from(chunk)));
      stream.on("error", reject);
      stream.on("end", () => resolve(Buffer.concat(chunks)));
    });
  }
}

export default OneDriveService;
