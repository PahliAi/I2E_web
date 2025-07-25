/**
 * I2E Cache Management System
 * Handles browser localStorage for invoice workflow persistence
 * 
 * @version 1.0
 * @author I2E Development Team
 */

// ===== CACHE CONFIGURATION =====

const CACHE_KEYS = {
    PENDING: 'i2e_pending_invoices',
    APPROVED: 'i2e_approved_invoices',
    REJECTED: 'i2e_rejected_invoices',
    PREFERENCES: 'i2e_user_preferences'
};

const CACHE_VERSION = '1.0';
const MAX_CACHE_SIZE = 5 * 1024 * 1024; // 5MB limit for safety

// ===== CORE INVOICE LIFECYCLE MANAGEMENT =====

/**
 * Add extracted invoice to pending cache
 * @param {Object} invoiceData - Complete extracted invoice data
 * @returns {boolean} Success status
 */
async function addPendingInvoice(invoiceData) {
    try {
        if (!invoiceData || !invoiceData.invoiceNumber) {
            logError('addPendingInvoice: Invalid invoice data', invoiceData);
            return false;
        }

        const pendingInvoices = await getPendingInvoices();
        
        // Check for duplicates
        const existingIndex = pendingInvoices.findIndex(invoice => 
            invoice.invoiceNumber === invoiceData.invoiceNumber
        );

        console.log('🔍 addPendingInvoice called');
        console.log('🔍 invoiceData.fullInvoiceData length:', invoiceData.fullInvoiceData ? invoiceData.fullInvoiceData.length : 'undefined');
        
        // Debug credit note amounts
        if (invoiceData.fullInvoiceData && invoiceData.fullInvoiceData.length > 0) {
            const firstItem = invoiceData.fullInvoiceData[0];
            if (firstItem.creditNote) {
                console.log('💾 Credit note stored:', firstItem.invoiceNumber, 'Corrected:', firstItem.isCreditNoteCorrected);
            }
        }
        
        // If invoiceData already has fullInvoiceData, use it directly (don't nest it)
        // If invoiceData is an array, use it as the line items
        // If invoiceData is a single object without fullInvoiceData, wrap it in array
        let fullInvoiceDataToStore;
        if (invoiceData.fullInvoiceData && Array.isArray(invoiceData.fullInvoiceData)) {
            // invoiceData already has a proper fullInvoiceData array - use it directly
            fullInvoiceDataToStore = invoiceData.fullInvoiceData;
            console.log('🔍 Using existing fullInvoiceData array');
        } else if (Array.isArray(invoiceData)) {
            // invoiceData is itself an array of line items
            fullInvoiceDataToStore = invoiceData;
            console.log('🔍 Using invoiceData as line items array');
        } else {
            // invoiceData is a single invoice object, wrap in array
            fullInvoiceDataToStore = [invoiceData];
            console.log('🔍 Wrapping single invoiceData in array');
        }
        
        console.log('🔍 fullInvoiceDataToStore length:', fullInvoiceDataToStore.length);

        const cacheEntry = {
            invoiceNumber: invoiceData.invoiceNumber,
            status: 'pending',
            extractedDate: new Date().toISOString(),
            source: 'new',
            fullInvoiceData: fullInvoiceDataToStore,
            validationData: null,
            lastModified: new Date().toISOString()
        };

        if (existingIndex >= 0) {
            // Update existing pending invoice
            pendingInvoices[existingIndex] = cacheEntry;
            logInfo(`Updated existing pending invoice: ${invoiceData.invoiceNumber}`);
        } else {
            // Add new pending invoice
            pendingInvoices.push(cacheEntry);
            logInfo(`Added new pending invoice: ${invoiceData.invoiceNumber}`);
        }

        return await saveToIndexedDB(CACHE_KEYS.PENDING, pendingInvoices);

    } catch (error) {
        logError('Error adding pending invoice:', error);
        return false;
    }
}

/**
 * Get all pending invoices from cache
 * @returns {Array} Array of pending invoice objects
 */
async function getPendingInvoices() {
    try {
        const pending = await loadFromIndexedDB(CACHE_KEYS.PENDING, []);
        
        // Mark cache items as from cache (not new)
        return pending.map(invoice => ({
            ...invoice,
            source: invoice.source === 'new' ? 'cache' : invoice.source
        }));

    } catch (error) {
        logError('Error getting pending invoices:', error);
        return [];
    }
}

/**
 * Approve an invoice (move from pending to approved)
 * @param {string} invoiceNumber - Invoice number to approve
 * @param {string} userId - User performing the approval
 * @param {string} comments - Optional approval comments
 * @returns {boolean} Success status
 */
async function approveInvoice(invoiceNumber, userId = 'unknown', comments = '') {
    try {
        const pendingInvoices = await getPendingInvoices();
        const invoiceIndex = pendingInvoices.findIndex(inv => inv.invoiceNumber === invoiceNumber);
        
        if (invoiceIndex === -1) {
            logError(`Invoice not found in pending: ${invoiceNumber}`);
            return false;
        }

        const invoice = pendingInvoices[invoiceIndex];
        
        // Create approved invoice (keep full data since we have IndexedDB storage)
        const approvedInvoice = {
            ...invoice,
            status: 'approved',
            approvalDate: new Date().toISOString(),
            approvedBy: userId,
            comments: comments,
            summary: extractInvoiceSummary(invoice.fullInvoiceData),
            lastModified: new Date().toISOString()
        };

        // Remove from pending
        pendingInvoices.splice(invoiceIndex, 1);
        
        // Add to approved
        const approvedInvoices = await getApprovedInvoices();
        approvedInvoices.push(approvedInvoice);

        // Save both arrays
        const pendingSaved = await saveToIndexedDB(CACHE_KEYS.PENDING, pendingInvoices);
        const approvedSaved = await saveToIndexedDB(CACHE_KEYS.APPROVED, approvedInvoices);

        if (pendingSaved && approvedSaved) {
            logInfo(`Invoice approved: ${invoiceNumber} by ${userId}`);
            return true;
        } else {
            logError('Failed to save approval changes');
            return false;
        }

    } catch (error) {
        logError('Error approving invoice:', error);
        return false;
    }
}

/**
 * Reject an invoice (move from pending to rejected)
 * @param {string} invoiceNumber - Invoice number to reject
 * @param {string} userId - User performing the rejection
 * @param {string} comments - Rejection reason (recommended)
 * @returns {boolean} Success status
 */
async function rejectInvoice(invoiceNumber, userId = 'unknown', comments = '') {
    try {
        const pendingInvoices = await getPendingInvoices();
        const invoiceIndex = pendingInvoices.findIndex(inv => inv.invoiceNumber === invoiceNumber);
        
        if (invoiceIndex === -1) {
            logError(`Invoice not found in pending: ${invoiceNumber}`);
            return false;
        }

        const invoice = pendingInvoices[invoiceIndex];
        
        // Create rejected invoice (keep full data since we have IndexedDB storage)
        const rejectedInvoice = {
            ...invoice,
            status: 'rejected',
            rejectionDate: new Date().toISOString(),
            rejectedBy: userId,
            comments: comments,
            summary: extractInvoiceSummary(invoice.fullInvoiceData),
            lastModified: new Date().toISOString()
        };

        // Remove from pending
        pendingInvoices.splice(invoiceIndex, 1);
        
        // Add to rejected
        const rejectedInvoices = await getRejectedInvoices();
        rejectedInvoices.push(rejectedInvoice);

        // Save both arrays
        const pendingSaved = await saveToIndexedDB(CACHE_KEYS.PENDING, pendingInvoices);
        const rejectedSaved = await saveToIndexedDB(CACHE_KEYS.REJECTED, rejectedInvoices);

        if (pendingSaved && rejectedSaved) {
            logInfo(`Invoice rejected: ${invoiceNumber} by ${userId}`);
            return true;
        } else {
            logError('Failed to save rejection changes');
            return false;
        }

    } catch (error) {
        logError('Error rejecting invoice:', error);
        return false;
    }
}

/**
 * Get all approved invoices from cache
 * @returns {Array} Array of approved invoice summaries
 */
async function getApprovedInvoices() {
    try {
        return await loadFromIndexedDB(CACHE_KEYS.APPROVED, []);
    } catch (error) {
        logError('Error getting approved invoices:', error);
        return [];
    }
}

/**
 * Get all rejected invoices from cache
 * @returns {Array} Array of rejected invoice summaries
 */
async function getRejectedInvoices() {
    try {
        return await loadFromIndexedDB(CACHE_KEYS.REJECTED, []);
    } catch (error) {
        logError('Error getting rejected invoices:', error);
        return [];
    }
}

// ===== DATA CONSOLIDATION FUNCTIONS =====

/**
 * Get all invoices regardless of status
 * @returns {Array} Combined array of all invoices
 */
async function getAllInvoices() {
    try {
        const pending = await getPendingInvoices();
        const approved = await getApprovedInvoices();
        const rejected = await getRejectedInvoices();
        
        return [...pending, ...approved, ...rejected];
        
    } catch (error) {
        logError('Error getting all invoices:', error);
        return [];
    }
}

/**
 * Get invoices for validation calculations (exclude rejected)
 * @returns {Promise<Array>} Pending + approved invoices only
 */
async function getInvoicesForValidation() {
    try {
        const pending = await getPendingInvoices();
        const approved = await getApprovedInvoices();
        
        return [...pending, ...approved];
        
    } catch (error) {
        logError('Error getting invoices for validation:', error);
        return [];
    }
}

/**
 * Deduplicate invoices when merging cache with uploaded data
 * @param {Array} cachedInvoices - Invoices from cache
 * @param {Array} uploadedInvoices - Invoices from Excel upload
 * @returns {Array} Deduplicated invoice array
 */
function deduplicateInvoices(cachedInvoices, uploadedInvoices) {
    try {
        if (!Array.isArray(cachedInvoices)) cachedInvoices = [];
        if (!Array.isArray(uploadedInvoices)) uploadedInvoices = [];
        
        const result = [...cachedInvoices];
        let duplicateCount = 0;
        let newCount = 0;
        
        uploadedInvoices.forEach(uploaded => {
            const existingIndex = result.findIndex(cached => 
                getInvoiceNumber(cached) === getInvoiceNumber(uploaded)
            );
            
            if (existingIndex >= 0) {
                // Uploaded data takes precedence, but preserve approval status
                const existing = result[existingIndex];
                const updatedInvoice = mergeInvoiceData(existing, uploaded);
                result[existingIndex] = updatedInvoice;
                duplicateCount++;
                
                logInfo(`Updated duplicate invoice: ${getInvoiceNumber(uploaded)}`);
            } else {
                // Convert uploaded invoice to cache format
                const cacheEntry = convertUploadedToCacheFormat(uploaded);
                result.push(cacheEntry);
                newCount++;
                
                logInfo(`Added new invoice from upload: ${getInvoiceNumber(uploaded)}`);
            }
        });
        
        logInfo(`Deduplication complete: ${duplicateCount} updated, ${newCount} new`);
        return result;
        
    } catch (error) {
        logError('Error deduplicating invoices:', error);
        return cachedInvoices;
    }
}

/**
 * Merge invoice data with precedence to uploaded data but preserve approval status
 * @param {Object} existingInvoice - Existing invoice in cache
 * @param {Object} uploadedInvoice - New invoice from upload
 * @returns {Object} Merged invoice data
 */
function mergeInvoiceData(existingInvoice, uploadedInvoice) {
    try {
        // Convert uploaded to cache format
        const uploadedCache = convertUploadedToCacheFormat(uploadedInvoice);
        
        // Preserve approval status and metadata from existing
        return {
            ...uploadedCache,
            status: existingInvoice.status,
            approvalDate: existingInvoice.approvalDate,
            approvedBy: existingInvoice.approvedBy,
            rejectionDate: existingInvoice.rejectionDate,
            rejectedBy: existingInvoice.rejectedBy,
            comments: existingInvoice.comments,
            extractedDate: existingInvoice.extractedDate,
            lastModified: new Date().toISOString()
        };
        
    } catch (error) {
        logError('Error merging invoice data:', error);
        return existingInvoice;
    }
}

// ===== PROJECT CALCULATION FUNCTIONS =====

/**
 * Calculate project totals excluding rejected invoices
 * @param {Array} invoices - Array of invoices to calculate
 * @returns {Object} Project totals grouped by WBS/Project ID
 */
function calculateProjectTotals(invoices) {
    try {
        if (!Array.isArray(invoices)) {
            logError('calculateProjectTotals: Invalid invoices array');
            return {};
        }
        
        // Only include pending and approved (exclude rejected)
        const validInvoices = invoices.filter(invoice => 
            invoice.status === 'pending' || invoice.status === 'approved'
        );
        
        const projectGroups = {};
        
        validInvoices.forEach(invoice => {
            const projectId = getProjectId(invoice);
            const invoiceTotal = getInvoiceTotal(invoice);
            
            if (!projectId || !invoiceTotal) return;
            
            if (!projectGroups[projectId]) {
                projectGroups[projectId] = {
                    projectId: projectId,
                    projectName: getProjectName(invoice),
                    totalInvoiced: 0,
                    pendingCount: 0,
                    approvedCount: 0,
                    rejectedCount: 0,
                    invoices: []
                };
            }
            
            projectGroups[projectId].totalInvoiced += invoiceTotal;
            projectGroups[projectId].invoices.push(invoice);
            
            // Count by status
            if (invoice.status === 'pending') projectGroups[projectId].pendingCount++;
            if (invoice.status === 'approved') projectGroups[projectId].approvedCount++;
        });
        
        logInfo(`Calculated totals for ${Object.keys(projectGroups).length} projects`);
        return projectGroups;
        
    } catch (error) {
        logError('Error calculating project totals:', error);
        return {};
    }
}

/**
 * Aggregate invoices by WBS code
 * @param {Array} invoices - Array of invoices to aggregate
 * @returns {Object} Invoices grouped by WBS code
 */
function aggregateByWBS(invoices) {
    try {
        if (!Array.isArray(invoices)) return {};
        
        const wbsGroups = {};
        
        invoices.forEach(invoice => {
            const wbsCode = standardizeWBS(getProjectId(invoice));
            
            if (!wbsCode) return;
            
            if (!wbsGroups[wbsCode]) {
                wbsGroups[wbsCode] = [];
            }
            
            wbsGroups[wbsCode].push(invoice);
        });
        
        return wbsGroups;
        
    } catch (error) {
        logError('Error aggregating by WBS:', error);
        return {};
    }
}

// ===== UTILITY FUNCTIONS =====

/**
 * Clear all cached invoice data
 * @returns {Promise<boolean>} Success status
 */
async function clearCache() {
    try {
        // Clear IndexedDB invoice data
        await removeFromIndexedDB(CACHE_KEYS.PENDING);
        await removeFromIndexedDB(CACHE_KEYS.APPROVED);
        await removeFromIndexedDB(CACHE_KEYS.REJECTED);
        
        // Also clear any legacy localStorage invoice data 
        localStorage.removeItem(CACHE_KEYS.PENDING);
        localStorage.removeItem(CACHE_KEYS.APPROVED);
        localStorage.removeItem(CACHE_KEYS.REJECTED);
        
        logInfo('Cache cleared successfully (localStorage + IndexedDB)');
        return true;
        
    } catch (error) {
        logError('Error clearing cache:', error);
        return false;
    }
}

/**
 * Get cache statistics
 * @returns {Promise<Object>} Cache size and count information
 */
async function getCacheStats() {
    try {
        const pending = await getPendingInvoices() || [];
        const approved = await getApprovedInvoices() || [];
        const rejected = await getRejectedInvoices() || [];
        
        // Ensure all values are arrays
        const safePending = Array.isArray(pending) ? pending : [];
        const safeApproved = Array.isArray(approved) ? approved : [];
        const safeRejected = Array.isArray(rejected) ? rejected : [];
        
        const pendingSize = JSON.stringify(safePending).length;
        const approvedSize = JSON.stringify(safeApproved).length;
        const rejectedSize = JSON.stringify(safeRejected).length;
        const totalSize = pendingSize + approvedSize + rejectedSize;
        
        return {
            counts: {
                pending: safePending.length,
                approved: safeApproved.length,
                rejected: safeRejected.length,
                total: safePending.length + safeApproved.length + safeRejected.length
            },
            sizes: {
                pending: pendingSize,
                approved: approvedSize,
                rejected: rejectedSize,
                total: totalSize,
                totalMB: (totalSize / (1024 * 1024)).toFixed(2)
            },
            limits: {
                maxSize: MAX_CACHE_SIZE,
                maxSizeMB: (MAX_CACHE_SIZE / (1024 * 1024)).toFixed(2),
                percentUsed: ((totalSize / MAX_CACHE_SIZE) * 100).toFixed(1)
            },
            lastUpdated: new Date().toISOString()
        };
        
    } catch (error) {
        logError('Error getting cache stats:', error);
        return {
            counts: { pending: 0, approved: 0, rejected: 0, total: 0 },
            sizes: { total: 0, totalMB: '0.00' },
            error: error.message
        };
    }
}

/**
 * Export all cache data to Excel format
 * @returns {Array} Flattened array suitable for Excel export
 */
async function exportCacheToExcel() {
    try {
        const allInvoices = await getAllInvoices();
        const exportData = [];
        
        allInvoices.forEach(invoice => {
            if (invoice.fullInvoiceData && Array.isArray(invoice.fullInvoiceData)) {
                // Pending invoices have full data
                invoice.fullInvoiceData.forEach(lineItem => {
                    exportData.push({
                        ...lineItem,
                        approvalStatus: invoice.status,
                        approvalDate: invoice.approvalDate || invoice.rejectionDate || '',
                        approvedBy: invoice.approvedBy || invoice.rejectedBy || '',
                        comments: invoice.comments || '',
                        extractedDate: invoice.extractedDate,
                        lastModified: invoice.lastModified
                    });
                });
            } else if (invoice.summary) {
                // Approved/rejected invoices have summary only
                exportData.push({
                    ...invoice.summary,
                    approvalStatus: invoice.status,
                    approvalDate: invoice.approvalDate || invoice.rejectionDate || '',
                    approvedBy: invoice.approvedBy || invoice.rejectedBy || '',
                    comments: invoice.comments || '',
                    extractedDate: invoice.extractedDate || '',
                    lastModified: invoice.lastModified
                });
            }
        });
        
        logInfo(`Exported ${exportData.length} line items from cache`);
        return exportData;
        
    } catch (error) {
        logError('Error exporting cache to Excel:', error);
        return [];
    }
}

// ===== HELPER FUNCTIONS =====

/**
 * Extract summary data from full invoice data
 * @param {Array} fullInvoiceData - Complete invoice line items
 * @returns {Object} Summary object
 */
function extractInvoiceSummary(fullInvoiceData) {
    if (!Array.isArray(fullInvoiceData) || fullInvoiceData.length === 0) {
        return {};
    }
    
    const firstItem = fullInvoiceData[0];
    const totalAmount = firstItem.extractedInvoiceTotal || 
                       fullInvoiceData.reduce((sum, item) => sum + (item.positionTotal || 0), 0);
    
    return {
        projectId: firstItem.projectId,
        customerId: firstItem.customerId,
        fileName: firstItem.fileName,
        invoiceNumber: firstItem.invoiceNumber,
        invoiceDate: firstItem.dateOfInvoice,
        monthOfInvoice: firstItem.monthOfInvoice,
        currency: firstItem.currency,
        totalAmount: totalAmount,
        creditNote: firstItem.creditNote,
        lineItemCount: fullInvoiceData.length
    };
}

/**
 * Get invoice number from various invoice formats
 * @param {Object} invoice - Invoice object
 * @returns {string} Invoice number
 */
function getInvoiceNumber(invoice) {
    if (invoice.invoiceNumber) return invoice.invoiceNumber;
    if (invoice.fullInvoiceData && invoice.fullInvoiceData[0]) {
        return invoice.fullInvoiceData[0].invoiceNumber;
    }
    if (invoice.summary && invoice.summary.invoiceNumber) {
        return invoice.summary.invoiceNumber;
    }
    return '';
}

/**
 * Get project ID from various invoice formats
 * @param {Object} invoice - Invoice object
 * @returns {string} Project ID
 */
function getProjectId(invoice) {
    if (invoice.projectId) return invoice.projectId;
    if (invoice.fullInvoiceData && invoice.fullInvoiceData[0]) {
        return invoice.fullInvoiceData[0].projectId;
    }
    if (invoice.summary && invoice.summary.projectId) {
        return invoice.summary.projectId;
    }
    return '';
}

/**
 * Get project name from project ID
 * @param {Object} invoice - Invoice object
 * @returns {string} Project name
 */
function getProjectName(invoice) {
    const projectId = getProjectId(invoice);
    if (projectId && projectId.includes('PRO')) {
        return `Project ${projectId.split('-')[1] || projectId}`;
    }
    return projectId || 'Unknown Project';
}

/**
 * Get invoice total from various invoice formats
 * @param {Object} invoice - Invoice object
 * @returns {number} Invoice total amount
 */
function getInvoiceTotal(invoice) {
    if (invoice.summary && typeof invoice.summary.totalAmount === 'number') {
        return invoice.summary.totalAmount;
    }
    if (invoice.fullInvoiceData && invoice.fullInvoiceData[0]) {
        return invoice.fullInvoiceData[0].extractedInvoiceTotal || 0;
    }
    return 0;
}

/**
 * Convert uploaded invoice data to cache format
 * @param {Object} uploadedInvoice - Raw uploaded invoice data
 * @returns {Object} Cache-formatted invoice
 */
function convertUploadedToCacheFormat(uploadedInvoice) {
    return {
        invoiceNumber: getInvoiceNumber(uploadedInvoice),
        status: 'pending',
        extractedDate: new Date().toISOString(),
        source: 'upload',
        fullInvoiceData: Array.isArray(uploadedInvoice) ? uploadedInvoice : [uploadedInvoice],
        validationData: null,
        lastModified: new Date().toISOString()
    };
}

// ===== GLOBAL AVAILABILITY =====

// Make functions globally available for browser usage
if (typeof window !== 'undefined') {
    // Core functions
    window.addPendingInvoice = addPendingInvoice;
    window.getPendingInvoices = getPendingInvoices;
    window.approveInvoice = approveInvoice;
    window.rejectInvoice = rejectInvoice;
    window.getApprovedInvoices = getApprovedInvoices;
    window.getRejectedInvoices = getRejectedInvoices;
    
    // Data consolidation
    window.getAllInvoices = getAllInvoices;
    window.getInvoicesForValidation = getInvoicesForValidation;
    window.deduplicateInvoices = deduplicateInvoices;
    window.mergeInvoiceData = mergeInvoiceData;
    
    // Project calculations
    window.calculateProjectTotals = calculateProjectTotals;
    window.aggregateByWBS = aggregateByWBS;
    
    // Utilities
    window.clearCache = clearCache;
    window.getCacheStats = getCacheStats;
    window.exportCacheToExcel = exportCacheToExcel;
    
    logInfo('I2E Cache Management System loaded successfully');
}

// Export for module usage
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        // Core functions
        addPendingInvoice,
        getPendingInvoices,
        approveInvoice,
        rejectInvoice,
        getApprovedInvoices,
        getRejectedInvoices,
        
        // Data consolidation
        getAllInvoices,
        getInvoicesForValidation,
        deduplicateInvoices,
        mergeInvoiceData,
        
        // Project calculations
        calculateProjectTotals,
        aggregateByWBS,
        
        // Utilities
        clearCache,
        getCacheStats,
        exportCacheToExcel,
        
        // Constants
        CACHE_KEYS,
        CACHE_VERSION
    };
}