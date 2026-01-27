// Application Constants

// Debounce and refresh intervals
export const DEBOUNCE_DELAY_MS = 300;
export const NOTES_REFRESH_INTERVAL_MS = 30_000; // 30 seconds
export const DATA_AUTO_REFRESH_INTERVAL_MS = 5 * 60 * 1000; // 5 minutes
export const AUTO_SAVE_DELAY_MS = 10_000; // 10 seconds

// UI Constants
export const DEFAULT_TEXTAREA_ROWS = 6;
export const DEFAULT_ITEMS_PER_PAGE = 100;
export const ITEMS_PER_PAGE_OPTIONS = [50, 100, 200, 500, 1000, 99999];

// Column widths
export const COLUMN_WIDTHS = {
  BARCODE: 16,
  MEETING_NOTE: 60,
  FOLLOW_UP: 28,
  LAST_UPDATED: 24,
};
