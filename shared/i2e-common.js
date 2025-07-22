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
        
        // Validation
        isEmpty,
        isValidEmail,
        isValidWBS
    };
}

// Global availability for browser usage
if (typeof window !== 'undefined') {
    // Functions are already declared globally, no need to attach to window
    logInfo('I2E Common Utilities loaded successfully');
}