/**
 * I2E Common Utilities
 * Shared utilities used across I2E Invoice Processor and Invoice Validator
 * 
 * @version 1.0
 * @author I2E Development Team
 */

// ===== FILE HANDLING UTILITIES =====

/**
 * Format file size in human-readable format
 * @param {number} bytes - File size in bytes
 * @returns {string} Formatted file size (e.g., "1.5 MB")
 */
function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    
    return parseFloat((bytes / Math.pow(k, i)).toFixed(2)) + ' ' + sizes[i];
}

/**
 * Validate file type against allowed types
 * @param {File} file - File object to validate
 * @param {string[]} allowedTypes - Array of allowed MIME types
 * @returns {boolean} True if file type is allowed
 */
function validateFileType(file, allowedTypes) {
    if (!file || !allowedTypes || allowedTypes.length === 0) {
        return false;
    }
    
    return allowedTypes.includes(file.type);
}

/**
 * Handle file upload errors with user-friendly messages
 * @param {Error} error - Error object
 * @param {string} filename - Name of the file that caused the error
 */
function handleFileUploadError(error, filename) {
    console.error(`File upload error for ${filename}:`, error);
    
    let message = `Error processing file "${filename}": `;
    
    if (error.message.includes('PDF')) {
        message += 'Invalid PDF file or corrupted content.';
    } else if (error.message.includes('memory')) {
        message += 'File too large. Please try a smaller file.';
    } else if (error.message.includes('permission')) {
        message += 'Permission denied. Please check file access.';
    } else {
        message += 'Unknown error occurred. Please try again.';
    }
    
    alert(message);
}

// ===== DATA PROCESSING UTILITIES =====

/**
 * Standardize WBS codes for consistent matching
 * @param {string} wbsCode - Raw WBS code
 * @returns {string} Standardized WBS code
 */
function standardizeWBS(wbsCode) {
    if (!wbsCode || typeof wbsCode !== 'string') {
        return '';
    }
    
    // Remove "-EXN" suffix from PPM data and trim/uppercase
    let standardized = wbsCode.replace(/-EXN$/, '').trim().toUpperCase();
    
    // Extract project ID from full WBS format (EN44-PRO0022640 -> PRO0022640) 
    // This handles both CTC full format and RTC partial format
    const projectMatch = standardized.match(/([A-Z]{2,4}\d{7})$/);
    if (projectMatch) {
        return projectMatch[1];
    }
    
    return standardized;
}

/**
 * Format currency values consistently
 * @param {number} amount - Numeric amount
 * @param {string} currency - Currency code (default: 'EUR')
 * @returns {string} Formatted currency string
 */
function formatCurrency(amount, currency = 'EUR') {
    if (amount === null || amount === undefined || isNaN(amount)) {
        return `0.00 ${currency}`;
    }
    
    const formatter = new Intl.NumberFormat('de-DE', {
        style: 'currency',
        currency: currency,
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    });
    
    return formatter.format(amount);
}

/**
 * Generate unique ID for various purposes
 * @param {string} prefix - Optional prefix for the ID
 * @returns {string} Unique identifier
 */
function generateUniqueId(prefix = 'id') {
    const timestamp = Date.now().toString(36);
    const randomStr = Math.random().toString(36).substr(2, 9);
    return `${prefix}_${timestamp}_${randomStr}`;
}

/**
 * Convert string to CSS-safe class name
 * @param {string} str - Input string
 * @returns {string} CSS-safe class name
 */
function makeCSSClass(str) {
    if (!str || typeof str !== 'string') {
        return 'unknown';
    }
    
    // Replace spaces and special characters with underscores
    // Keep only alphanumeric characters, hyphens, and underscores
    return str.replace(/[^a-zA-Z0-9_-]/g, '_').toLowerCase();
}

// ===== STORAGE UTILITIES =====

/**
 * Save data to localStorage with error handling
 * @param {string} key - Storage key
 * @param {any} data - Data to save (will be JSON stringified)
 * @returns {boolean} True if saved successfully
 */
function saveToLocalStorage(key, data) {
    try {
        const jsonData = JSON.stringify(data);
        localStorage.setItem(key, jsonData);
        return true;
    } catch (error) {
        console.error(`Failed to save to localStorage (${key}):`, error);
        return false;
    }
}

/**
 * Load data from localStorage with error handling
 * @param {string} key - Storage key
 * @param {any} defaultValue - Default value if key not found
 * @returns {any} Parsed data or default value
 */
function loadFromLocalStorage(key, defaultValue = null) {
    try {
        const jsonData = localStorage.getItem(key);
        if (jsonData === null) {
            return defaultValue;
        }
        return JSON.parse(jsonData);
    } catch (error) {
        console.error(`Failed to load from localStorage (${key}):`, error);
        return defaultValue;
    }
}

/**
 * Clear localStorage entries by prefix
 * @param {string} prefix - Prefix to match for deletion
 * @returns {number} Number of items deleted
 */
function clearStorageByPrefix(prefix) {
    let deletedCount = 0;
    const keysToDelete = [];
    
    // Collect keys to delete (can't delete while iterating)
    for (let i = 0; i < localStorage.length; i++) {
        const key = localStorage.key(i);
        if (key && key.startsWith(prefix)) {
            keysToDelete.push(key);
        }
    }
    
    // Delete collected keys
    keysToDelete.forEach(key => {
        localStorage.removeItem(key);
        deletedCount++;
    });
    
    return deletedCount;
}

// ===== DATE/TIME UTILITIES =====

/**
 * Format date string consistently
 * @param {string|Date} dateInput - Date string or Date object
 * @param {string} format - Format type ('short', 'long', 'iso')
 * @returns {string} Formatted date string
 */
function formatDate(dateInput, format = 'short') {
    if (!dateInput) return '';
    
    let date;
    if (typeof dateInput === 'string') {
        date = new Date(dateInput);
    } else if (dateInput instanceof Date) {
        date = dateInput;
    } else {
        return '';
    }
    
    if (isNaN(date.getTime())) {
        return '';
    }
    
    switch (format) {
        case 'long':
            return date.toLocaleDateString('en-US', { 
                year: 'numeric', 
                month: 'long', 
                day: 'numeric' 
            });
        case 'iso':
            return date.toISOString().split('T')[0];
        case 'short':
        default:
            return date.toLocaleDateString('en-US');
    }
}

/**
 * Parse invoice month string to standardized format
 * @param {string} monthString - Month string (e.g., "January 2025", "01/2025")
 * @returns {string} Standardized month format ("YYYY-MM")
 */
function parseInvoiceMonth(monthString) {
    if (!monthString || typeof monthString !== 'string') {
        return '';
    }
    
    // Try to parse different month formats
    const monthNames = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ];
    
    // Handle "Month Year" format (e.g., "January 2025")
    for (let i = 0; i < monthNames.length; i++) {
        if (monthString.includes(monthNames[i])) {
            const year = monthString.match(/\d{4}/)?.[0];
            if (year) {
                const monthNum = (i + 1).toString().padStart(2, '0');
                return `${year}-${monthNum}`;
            }
        }
    }
    
    // Handle "MM/YYYY" format
    const mmYyyyMatch = monthString.match(/(\d{1,2})\/(\d{4})/);
    if (mmYyyyMatch) {
        const month = mmYyyyMatch[1].padStart(2, '0');
        const year = mmYyyyMatch[2];
        return `${year}-${month}`;
    }
    
    // Handle "YYYY-MM" format (already correct)
    const yyyyMmMatch = monthString.match(/(\d{4})-(\d{1,2})/);
    if (yyyyMmMatch) {
        const year = yyyyMmMatch[1];
        const month = yyyyMmMatch[2].padStart(2, '0');
        return `${year}-${month}`;
    }
    
    return monthString; // Return as-is if no format matched
}

/**
 * Get current timestamp in ISO format
 * @returns {string} Current timestamp
 */
function getCurrentTimestamp() {
    return new Date().toISOString();
}

// ===== LOGGING UTILITIES =====

/**
 * Log information with consistent formatting
 * @param {string} message - Log message
 * @param {any} data - Optional data to log
 */
function logInfo(message, data = null) {
    const timestamp = new Date().toISOString();
    console.log(`[${timestamp}] INFO: ${message}`, data || '');
}

/**
 * Log error with consistent formatting
 * @param {string} message - Error message
 * @param {Error|any} error - Error object or data
 */
function logError(message, error = null) {
    const timestamp = new Date().toISOString();
    console.error(`[${timestamp}] ERROR: ${message}`, error || '');
}

/**
 * Log warning with consistent formatting
 * @param {string} message - Warning message
 * @param {any} data - Optional data to log
 */
function logWarning(message, data = null) {
    const timestamp = new Date().toISOString();
    console.warn(`[${timestamp}] WARNING: ${message}`, data || '');
}

// ===== NUMBER AND CURRENCY UTILITIES =====

/**
 * Parse currency string to numeric value
 * @param {string} currencyString - Currency string (e.g., "€1,234.56", "1.234,56")
 * @returns {number} Numeric value or 0 if parsing fails
 */
function parseCurrencyValue(currencyString) {
    if (!currencyString || typeof currencyString !== 'string') {
        return 0;
    }
    
    // Remove currency symbols and spaces
    let cleanStr = currencyString.replace(/[€$£¥₹]|\s/g, '');
    
    // Handle different decimal separators
    // If string contains both comma and dot, assume European format (1.234,56)
    if (cleanStr.includes(',') && cleanStr.includes('.')) {
        // European format: remove dots (thousands separator), replace comma with dot
        cleanStr = cleanStr.replace(/\./g, '').replace(',', '.');
    } else if (cleanStr.includes(',')) {
        // Could be thousands separator or decimal separator
        // If comma is followed by exactly 3 digits and end, it's thousands separator
        if (/,\d{3}$/.test(cleanStr)) {
            cleanStr = cleanStr.replace(',', '');
        } else {
            // Assume it's decimal separator
            cleanStr = cleanStr.replace(',', '.');
        }
    }
    
    const numValue = parseFloat(cleanStr);
    return isNaN(numValue) ? 0 : numValue;
}

/**
 * Round number to specified decimal places
 * @param {number} num - Number to round
 * @param {number} decimals - Number of decimal places (default: 2)
 * @returns {number} Rounded number
 */
function roundToDecimals(num, decimals = 2) {
    if (isNaN(num)) return 0;
    
    const factor = Math.pow(10, decimals);
    return Math.round(num * factor) / factor;
}

/**
 * Parse numeric value from Excel cell (handles formulas and regular values)
 * @param {any} cellValue - Raw cell value from ExcelJS
 * @returns {number} Parsed numeric value
 */
function parseExcelNumericValue(cellValue) {
    if (cellValue === null || cellValue === undefined) {
        return 0;
    }
    
    // Handle Excel error values (#N/A, #REF!, #DIV/0!, etc.)
    if (typeof cellValue === 'object' && cellValue.error !== undefined) {
        console.warn('📊 Excel error in numeric cell:', cellValue.error);
        return 0;
    }
    
    // Handle Excel formula objects
    if (typeof cellValue === 'object' && cellValue.result !== undefined) {
        // Check if formula result is also an error
        if (typeof cellValue.result === 'object' && cellValue.result.error !== undefined) {
            console.warn('📊 Excel formula error:', cellValue.result.error);
            return 0;
        }
        // ExcelJS formula object - use the calculated result
        return parseFloat(cellValue.result) || 0;
    }
    
    // Handle regular numeric values
    if (typeof cellValue === 'number') {
        return cellValue;
    }
    
    // Handle string values (try to parse as float)
    if (typeof cellValue === 'string') {
        // Check for Excel error strings
        if (cellValue.startsWith('#')) {
            console.warn('📊 Excel error string in numeric cell:', cellValue);
            return 0;
        }
        return parseFloat(cellValue) || 0;
    }
    
    // Fallback
    return parseFloat(cellValue) || 0;
}

/**
 * Parse text value from Excel cell (handles formulas and regular values)
 * @param {any} cellValue - Raw cell value from ExcelJS
 * @returns {string} Parsed text value
 */
function parseExcelTextValue(cellValue) {
    if (cellValue === null || cellValue === undefined) {
        return '';
    }
    
    // Handle Excel error values (#N/A, #REF!, #DIV/0!, etc.)
    if (typeof cellValue === 'object' && cellValue.error !== undefined) {
        console.warn('📊 Excel error in text cell:', cellValue.error);
        return '';
    }
    
    // Handle Excel formula objects
    if (typeof cellValue === 'object' && cellValue.result !== undefined) {
        // Check if formula result is an error
        if (typeof cellValue.result === 'object' && cellValue.result.error !== undefined) {
            console.warn('📊 Excel formula error in text:', cellValue.result.error);
            return '';
        }
        // ExcelJS formula object - use the calculated result
        return String(cellValue.result || '').trim();
    }
    
    // Handle Date objects
    if (cellValue instanceof Date) {
        return cellValue.toISOString().split('T')[0];
    }
    
    // Handle boolean values
    if (typeof cellValue === 'boolean') {
        return cellValue ? 'TRUE' : 'FALSE';
    }
    
    // Handle regular values
    const textValue = String(cellValue || '').trim();
    
    // Check for Excel error strings
    if (textValue.startsWith('#')) {
        console.warn('📊 Excel error string in text cell:', textValue);
        return '';
    }
    
    return textValue;
}

/**
 * Parse date value from Excel cell (handles formulas, serial numbers, and date objects)
 * @param {any} cellValue - Raw cell value from ExcelJS
 * @returns {Date|null} Parsed Date object or null if invalid
 */
function parseExcelDateValue(cellValue) {
    if (cellValue === null || cellValue === undefined) {
        return null;
    }
    
    // Handle Excel error values
    if (typeof cellValue === 'object' && cellValue.error !== undefined) {
        console.warn('📊 Excel error in date cell:', cellValue.error);
        return null;
    }
    
    // Handle Excel formula objects
    if (typeof cellValue === 'object' && cellValue.result !== undefined) {
        // Check if formula result is an error
        if (typeof cellValue.result === 'object' && cellValue.result.error !== undefined) {
            console.warn('📊 Excel formula error in date:', cellValue.result.error);
            return null;
        }
        // Recursively parse the formula result
        return parseExcelDateValue(cellValue.result);
    }
    
    // Handle Date objects
    if (cellValue instanceof Date) {
        return isNaN(cellValue.getTime()) ? null : cellValue;
    }
    
    // Handle Excel serial date numbers (days since 1900-01-01, accounting for Excel's leap year bug)
    if (typeof cellValue === 'number') {
        if (cellValue < 1 || cellValue > 2958465) { // Valid Excel date range
            return null;
        }
        const excelDate = new Date((cellValue - 25569) * 86400 * 1000);
        return isNaN(excelDate.getTime()) ? null : excelDate;
    }
    
    // Handle string date values
    if (typeof cellValue === 'string') {
        const trimmed = cellValue.trim();
        if (!trimmed || trimmed.startsWith('#')) {
            return null;
        }
        const date = new Date(trimmed);
        return isNaN(date.getTime()) ? null : date;
    }
    
    return null;
}

/**
 * Parse percentage value from Excel cell (handles both 0.15 and 15% formats)
 * @param {any} cellValue - Raw cell value from ExcelJS
 * @returns {number} Percentage as decimal (0.15 for 15%)
 */
function parseExcelPercentageValue(cellValue) {
    if (cellValue === null || cellValue === undefined) {
        return 0;
    }
    
    // Handle Excel error values
    if (typeof cellValue === 'object' && cellValue.error !== undefined) {
        console.warn('📊 Excel error in percentage cell:', cellValue.error);
        return 0;
    }
    
    // Handle Excel formula objects
    if (typeof cellValue === 'object' && cellValue.result !== undefined) {
        if (typeof cellValue.result === 'object' && cellValue.result.error !== undefined) {
            console.warn('📊 Excel formula error in percentage:', cellValue.result.error);
            return 0;
        }
        return parseExcelPercentageValue(cellValue.result);
    }
    
    // Handle numeric values
    if (typeof cellValue === 'number') {
        return cellValue; // Excel usually stores percentages as decimals (0.15 for 15%)
    }
    
    // Handle string percentage values
    if (typeof cellValue === 'string') {
        const trimmed = cellValue.trim();
        if (trimmed.startsWith('#')) {
            console.warn('📊 Excel error string in percentage cell:', trimmed);
            return 0;
        }
        
        if (trimmed.endsWith('%')) {
            // Convert "15%" to 0.15
            const numStr = trimmed.slice(0, -1);
            const num = parseFloat(numStr);
            return isNaN(num) ? 0 : num / 100;
        }
        
        const num = parseFloat(trimmed);
        return isNaN(num) ? 0 : num;
    }
    
    return 0;
}

/**
 * Check if Excel cell value represents a truly empty cell (not zero)
 * @param {any} cellValue - Raw cell value from ExcelJS
 * @returns {boolean} True if cell is genuinely empty
 */
function isExcelCellEmpty(cellValue) {
    if (cellValue === null || cellValue === undefined) {
        return true;
    }
    
    // Handle Excel formula objects
    if (typeof cellValue === 'object' && cellValue.result !== undefined) {
        return isExcelCellEmpty(cellValue.result);
    }
    
    // Handle error values as empty
    if (typeof cellValue === 'object' && cellValue.error !== undefined) {
        return true;
    }
    
    // Handle string values
    if (typeof cellValue === 'string') {
        const trimmed = cellValue.trim();
        return trimmed === '' || trimmed.startsWith('#');
    }
    
    // Number 0 is not empty, but NaN is
    if (typeof cellValue === 'number') {
        return isNaN(cellValue);
    }
    
    return false;
}

/**
 * Calculate ISO week number for a given date
 * @param {Date} date - Date object
 * @returns {number} ISO week number (1-53)
 */
function getWeekNumber(date) {
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

/**
 * Calculate ISO week number AND year for a given date
 * Returns both week number and the ISO week year (which can differ from calendar year)
 * @param {Date} date - Date object
 * @returns {Object} {week: number, year: number}
 */
function getWeekNumberAndYear(date) {
    const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
    const dayNum = d.getUTCDay() || 7;
    
    // Move to the Thursday of this week - this determines which year the week belongs to
    d.setUTCDate(d.getUTCDate() + 4 - dayNum);
    const isoYear = d.getUTCFullYear();
    
    // Calculate week number within that ISO year
    const yearStart = new Date(Date.UTC(isoYear, 0, 1));
    const weekNum = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
    
    return {
        week: weekNum,
        year: isoYear
    };
}

/**
 * Universal date parser for I2E application
 * Handles various date formats and returns standardized period keys
 * @param {any} dateValue - Raw date value from Excel or other sources
 * @param {string} periodType - 'Week', 'Month', or 'Year'
 * @returns {string} Standardized period key (e.g., '2025-10', '2025', 'Unknown')
 */
function parseI2EDate(dateValue, periodType = 'Month') {
    if (!dateValue) return 'Unknown';
    
    const dateStr = String(dateValue).trim();
    if (!dateStr) return 'Unknown';
    
    let date = null;
    
    // Try different date formats
    if (periodType === 'Week') {
        // For weeks, expect full date strings
        if (dateStr.match(/^\d{4}-\d{2}-\d{2}$/)) {
            // YYYY-MM-DD format
            date = new Date(dateStr);
        } else if (dateStr.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
            // DD/MM/YYYY format (European)
            const [day, month, year] = dateStr.split('/');
            date = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
        } else if (dateStr.match(/^\d{2}-\d{2}-\d{4}$/)) {
            // DD-MM-YYYY format
            const [day, month, year] = dateStr.split('-');
            date = new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
        } else {
            // Try direct Date parsing
            date = new Date(dateStr);
        }
        
        if (date && !isNaN(date.getTime())) {
            const weekData = getWeekNumberAndYear(date);
            return `${weekData.year}-W${String(weekData.week).padStart(2, '0')}`;
        }
    } else {
        // For Month/Year periods
        
        // First try month names (from PPM data)
        const monthNames = ['january', 'february', 'march', 'april', 'may', 'june',
                          'july', 'august', 'september', 'october', 'november', 'december'];
        const shortMonthNames = ['jan', 'feb', 'mar', 'apr', 'may', 'jun',
                               'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
        
        const lowerDateStr = dateStr.toLowerCase();
        let monthIndex = monthNames.findIndex(month => month === lowerDateStr);
        if (monthIndex === -1) {
            monthIndex = shortMonthNames.findIndex(month => month === lowerDateStr);
        }
        
        if (monthIndex !== -1) {
            // Found month name - use current year
            const currentYear = new Date().getFullYear();
            date = new Date(currentYear, monthIndex, 1);
        } else if (dateStr.match(/^\d+$/)) {
            // Numeric period (from EXT SAP data)
            const periodNum = parseInt(dateStr);
            if (periodNum >= 1 && periodNum <= 12) {
                const currentYear = new Date().getFullYear();
                date = new Date(currentYear, periodNum - 1, 1);
            }
        } else if (dateStr.match(/^\d{4}-\d{2}$/)) {
            // YYYY-MM format
            const [year, month] = dateStr.split('-');
            date = new Date(parseInt(year), parseInt(month) - 1, 1);
        } else if (dateStr.match(/^\d{2}\/\d{4}$/)) {
            // MM/YYYY format
            const [month, year] = dateStr.split('/');
            date = new Date(parseInt(year), parseInt(month) - 1, 1);
        } else if (dateStr.match(/^\d{2}\/\d{2}\/\d{4}$/)) {
            // DD/MM/YYYY format (European) - extract month
            const [day, month, year] = dateStr.split('/');
            date = new Date(parseInt(year), parseInt(month) - 1, 1);
        } else if (dateStr.match(/^\d{2}-\d{2}-\d{4}$/)) {
            // DD-MM-YYYY format - extract month
            const [day, month, year] = dateStr.split('-');
            date = new Date(parseInt(year), parseInt(month) - 1, 1);
        } else {
            // Try direct Date parsing as fallback
            date = new Date(dateStr);
        }
        
        if (date && !isNaN(date.getTime())) {
            if (periodType === 'Month') {
                return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
            } else if (periodType === 'Year') {
                return `${date.getFullYear()}`;
            }
        }
    }
    
    return 'Unknown';
}

// ===== VALIDATION UTILITIES =====

/**
 * Check if value is empty or null
 * @param {any} value - Value to check
 * @returns {boolean} True if empty/null/undefined
 */
function isEmpty(value) {
    return value === null || 
           value === undefined || 
           value === '' || 
           (Array.isArray(value) && value.length === 0) ||
           (typeof value === 'object' && Object.keys(value).length === 0);
}

/**
 * Validate email format
 * @param {string} email - Email string to validate
 * @returns {boolean} True if valid email format
 */
function isValidEmail(email) {
    if (!email || typeof email !== 'string') return false;
    
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
}

/**
 * Validate WBS code format
 * @param {string} wbsCode - WBS code to validate
 * @returns {boolean} True if valid WBS format
 */
function isValidWBS(wbsCode) {
    if (!wbsCode || typeof wbsCode !== 'string') return false;
    
    // Support both full WBS format (AA44-PRO0012345) and RTC project format (PRO0012345)
    const fullWbsRegex = /^[A-Z]{2}\d{2}-[A-Z]{3}\d{7}$/;  // CTC format
    const rtcProjectRegex = /^[A-Z]{3}\d{7}$/;             // RTC format
    const standardized = standardizeWBS(wbsCode);
    
    return fullWbsRegex.test(standardized) || rtcProjectRegex.test(standardized);
}

// ===== INDEXEDDB UTILITIES =====

/**
 * IndexedDB wrapper for large data storage (cost data, invoice data)
 * Provides localStorage-like interface but with much larger storage capacity
 */

let i2eDB = null;

/**
 * Initialize IndexedDB database
 * @returns {Promise<IDBDatabase>} Database instance
 */
async function initializeIndexedDB() {
    if (i2eDB) return i2eDB;
    
    return new Promise((resolve, reject) => {
        const request = indexedDB.open('I2E_Database', 1);
        
        request.onerror = () => {
            console.error('Failed to open IndexedDB:', request.error);
            reject(request.error);
        };
        
        request.onsuccess = () => {
            i2eDB = request.result;
            console.log('IndexedDB initialized successfully');
            resolve(i2eDB);
        };
        
        request.onupgradeneeded = (event) => {
            const db = event.target.result;
            
            // Create object stores for different data types
            if (!db.objectStoreNames.contains('costData')) {
                db.createObjectStore('costData', { keyPath: 'id' });
            }
            if (!db.objectStoreNames.contains('invoiceData')) {
                db.createObjectStore('invoiceData', { keyPath: 'id' });
            }
            
            console.log('IndexedDB object stores created');
        };
    });
}

/**
 * Save large data to IndexedDB
 * @param {string} key - Storage key
 * @param {any} data - Data to save
 * @returns {Promise<boolean>} True if saved successfully
 */
async function saveToIndexedDB(key, data) {
    try {
        const db = await initializeIndexedDB();
        const transaction = db.transaction(['costData'], 'readwrite');
        const store = transaction.objectStore('costData');
        
        const dataToStore = {
            id: key,
            data: data,
            timestamp: new Date().toISOString(),
            size: JSON.stringify(data).length
        };
        
        return new Promise((resolve, reject) => {
            const request = store.put(dataToStore);
            
            request.onsuccess = () => {
                console.log(`💾 IndexedDB: Saved ${dataToStore.size} characters to key "${key}"`);
                resolve(true);
            };
            
            request.onerror = () => {
                console.error(`💾 IndexedDB: Failed to save key "${key}":`, request.error);
                reject(request.error);
            };
        });
    } catch (error) {
        console.error(`💾 IndexedDB: Error saving key "${key}":`, error);
        return false;
    }
}

/**
 * Load large data from IndexedDB
 * @param {string} key - Storage key
 * @param {any} defaultValue - Default value if key not found
 * @returns {Promise<any>} Retrieved data or default value
 */
async function loadFromIndexedDB(key, defaultValue = null) {
    try {
        const db = await initializeIndexedDB();
        const transaction = db.transaction(['costData'], 'readonly');
        const store = transaction.objectStore('costData');
        
        return new Promise((resolve, reject) => {
            const request = store.get(key);
            
            request.onsuccess = () => {
                if (request.result) {
                    console.log(`💾 IndexedDB: Loaded ${request.result.size} characters from key "${key}"`);
                    resolve(request.result.data);
                } else {
                    console.log(`💾 IndexedDB: Key "${key}" not found, using default value`);
                    resolve(defaultValue);
                }
            };
            
            request.onerror = () => {
                console.error(`💾 IndexedDB: Failed to load key "${key}":`, request.error);
                resolve(defaultValue);
            };
        });
    } catch (error) {
        console.error(`💾 IndexedDB: Error loading key "${key}":`, error);
        return defaultValue;
    }
}

/**
 * Remove data from IndexedDB
 * @param {string} key - Storage key to remove
 * @returns {Promise<boolean>} True if removed successfully
 */
async function removeFromIndexedDB(key) {
    try {
        const db = await initializeIndexedDB();
        const transaction = db.transaction(['costData'], 'readwrite');
        const store = transaction.objectStore('costData');
        
        return new Promise((resolve, reject) => {
            const request = store.delete(key);
            
            request.onsuccess = () => {
                console.log(`💾 IndexedDB: Removed key "${key}"`);
                resolve(true);
            };
            
            request.onerror = () => {
                console.error(`💾 IndexedDB: Failed to remove key "${key}":`, request.error);
                resolve(false);
            };
        });
    } catch (error) {
        console.error(`💾 IndexedDB: Error removing key "${key}":`, error);
        return false;
    }
}

/**
 * Clear IndexedDB storage by prefix (similar to localStorage version)
 * @param {string} prefix - Prefix to match for deletion
 * @returns {Promise<number>} Number of items deleted
 */
async function clearIndexedDBByPrefix(prefix) {
    try {
        const db = await initializeIndexedDB();
        const transaction = db.transaction(['costData'], 'readwrite');
        const store = transaction.objectStore('costData');
        
        return new Promise((resolve, reject) => {
            const getAllRequest = store.getAllKeys();
            
            getAllRequest.onsuccess = () => {
                const keys = getAllRequest.result.filter(key => key.startsWith(prefix));
                let deletedCount = 0;
                
                if (keys.length === 0) {
                    resolve(0);
                    return;
                }
                
                keys.forEach((key, index) => {
                    const deleteRequest = store.delete(key);
                    
                    deleteRequest.onsuccess = () => {
                        deletedCount++;
                        if (deletedCount === keys.length) {
                            console.log(`💾 IndexedDB: Cleared ${deletedCount} keys with prefix "${prefix}"`);
                            resolve(deletedCount);
                        }
                    };
                    
                    deleteRequest.onerror = () => {
                        console.error(`💾 IndexedDB: Failed to delete key "${key}"`);
                        deletedCount++;
                        if (deletedCount === keys.length) {
                            resolve(deletedCount);
                        }
                    };
                });
            };
            
            getAllRequest.onerror = () => {
                console.error('💾 IndexedDB: Failed to get keys for prefix clearing');
                resolve(0);
            };
        });
    } catch (error) {
        console.error(`💾 IndexedDB: Error clearing prefix "${prefix}":`, error);
        return 0;
    }
}

/**
 * Normalize Excel field names to standardized format
 * Maps various field name formats to consistent internal names
 * @param {Object} rawRow - Raw row from Excel with original field names
 * @returns {Object} Row with normalized field names
 */
function normalizeExcelFieldNames(rawRow) {
    const fieldMap = {
        // Service Provision Period variations
        'serviceProvisionPeriod': ['service provision period', 'service_provision_period', 'serviceprovisionperiod', 'month of invoice'],
        // Week Starts On variations  
        'weekStartsOn': ['week starts on', 'week_starts_on', 'weekstartson'],
        // User variations
        'user': ['user', 'username', 'employee', 'name'],
        // Cost variations  
        'costEuro': ['cost €', 'cost', 'value', 'amount'],
        // Total/Hours variations
        'total': ['total', 'hours', 'quantity'],
        // Project Type variations
        'projectType': ['project type', 'project_type', 'type'],
        // Resource Manager variations
        'resourceManager': ['resource manager', 'manager', 'rm'],
        // WBS Element variations
        'wbsElement': ['wbs element id', 'wbs', 'wbs_element'],
        // Project Name variations
        'projectName': ['project name', 'project', 'project_name'],
        // Job Title/Role variations
        'jobTitle': ['job title', 'role', 'position', 'title']
    };
    
    const normalizedRow = {};
    const availableFields = Object.keys(rawRow);
    
    // First, copy all original fields
    Object.assign(normalizedRow, rawRow);
    
    // Then add normalized field names
    for (const [standardName, variations] of Object.entries(fieldMap)) {
        for (const variation of variations) {
            const matchingField = availableFields.find(field => 
                field.toLowerCase().replace(/[^a-z0-9]/g, '') === variation.replace(/[^a-z0-9]/g, '')
            );
            
            if (matchingField && rawRow[matchingField]) {
                normalizedRow[standardName] = rawRow[matchingField];
                break; // Use first match found
            }
        }
    }
    
    return normalizedRow;
}

/**
 * Clear all I2E data from both localStorage and IndexedDB
 * Comprehensive cleanup function for complete cache clearing
 * @returns {Promise<boolean>} True if cleared successfully
 */
async function clearAllI2EData() {
    try {
        console.log('🧹 Starting comprehensive I2E data cleanup...');
        
        // Clear localStorage with i2e_ prefix
        const localStorageCleared = clearStorageByPrefix('i2e_');
        console.log(`📦 localStorage: Cleared ${localStorageCleared} items`);
        
        // Clear IndexedDB with i2e_ prefix
        const indexedDBCleared = await clearIndexedDBByPrefix('i2e_');
        console.log(`💾 IndexedDB: Cleared ${indexedDBCleared} items`);
        
        console.log('✅ Comprehensive I2E data cleanup completed!');
        return true;
        
    } catch (error) {
        console.error('❌ Error during comprehensive cleanup:', error);
        return false;
    }
}

/**
 * Hybrid storage function - tries IndexedDB first for cost data, falls back to localStorage
 * This provides compatibility during migration period
 * @param {string} key - Storage key
 * @param {any} defaultValue - Default value if key not found
 * @returns {Promise<any>} Retrieved data or default value
 */
async function loadCostDataFromStorage(key, defaultValue = []) {
    // For cost data cache, use IndexedDB for better performance and capacity
    if (key === 'i2e_cost_data_cache') {
        try {
            // First try IndexedDB
            const indexedDBData = await loadFromIndexedDB(key, null);
            if (indexedDBData !== null) {
                return indexedDBData;
            }
            
            // If not in IndexedDB, check localStorage and migrate
            const localStorageData = loadFromLocalStorage(key, defaultValue);
            if (localStorageData && localStorageData.length > 0) {
                console.log(`📦 Migrating ${localStorageData.length} cost records from localStorage to IndexedDB`);
                await saveToIndexedDB(key, localStorageData);
                // Clean up localStorage after successful migration
                localStorage.removeItem(key);
                return localStorageData;
            }
            
            return defaultValue;
        } catch (error) {
            console.error('Error loading cost data, falling back to localStorage:', error);
            return loadFromLocalStorage(key, defaultValue);
        }
    }
    
    // For other data, use regular localStorage
    return loadFromLocalStorage(key, defaultValue);
}

// ===== EXPORT FOR MODULE USAGE =====

// If using as a module, export functions
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        // File handling
        formatFileSize,
        validateFileType,
        handleFileUploadError,
        
        // Data processing
        standardizeWBS,
        formatCurrency,
        generateUniqueId,
        makeCSSClass,
        
        // Storage
        saveToLocalStorage,
        loadFromLocalStorage,
        clearStorageByPrefix,
        
        // Date/time
        formatDate,
        parseInvoiceMonth,
        getCurrentTimestamp,
        
        // Logging
        logInfo,
        logError,
        logWarning,
        
        // Numbers/currency
        parseCurrencyValue,
        roundToDecimals,
        parseExcelNumericValue,
        parseExcelTextValue,
        parseExcelDateValue,
        parseExcelPercentageValue,
        isExcelCellEmpty,
        parseI2EDate,
        getWeekNumber,
        getWeekNumberAndYear,
        
        // Validation
        isEmpty,
        isValidEmail,
        isValidWBS,
        
        // IndexedDB utilities
        initializeIndexedDB,
        saveToIndexedDB,
        loadFromIndexedDB,        removeFromIndexedDB,
        clearIndexedDBByPrefix,
        loadCostDataFromStorage,
        clearAllI2EData,
        normalizeExcelFieldNames
    };
}

// Global availability for browser usage
if (typeof window !== 'undefined') {
    // Functions are already declared globally, no need to attach to window
    logInfo('I2E Common Utilities loaded successfully');
}