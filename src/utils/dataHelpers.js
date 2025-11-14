// Data transformation and validation utilities

/**
 * Normalize barcode to uppercase and trimmed
 */
export const normalizeBarcode = (barcode) => {
  return barcode ? String(barcode).trim().toUpperCase() : "";
};

/**
 * Calculate age in days from a date string
 */
export const ageInDays = (dateStr) => {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (isNaN(d)) return "";
  const today = new Date();
  return Math.ceil(Math.abs(today - d) / (1000 * 60 * 60 * 24));
};

/**
 * Format ticket number by cleaning and converting to integer
 */
export const formatTicketNumber = (ticket) => {
  if (!ticket) return "";
  const cleaned = String(ticket).replace(/[^0-9.]/g, "");
  const num = parseFloat(cleaned);
  if (isNaN(num)) return "";
  return String(Math.floor(num));
};

/**
 * Validate barcode format (e.g., RV123456, MC987654)
 */
export const validateBarcode = (barcode) => {
  return /^[A-Z]{2}\d{6}$/.test(barcode);
};

/**
 * Sanitize numeric input
 */
export const sanitizeNumeric = (val) => {
  const num = parseFloat(val);
  return isNaN(num) ? 0 : num;
};
