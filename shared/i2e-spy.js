/**
 * I2E Spy Modal - Shared Cache Debugging Tool
 * Provides hidden cache inspection capabilities for both Invoice Processor and Validator
 * 
 * @version 1.0
 * @author I2E Development Team
 */

// ===== SPY MODAL CONFIGURATION =====

const SPY_CONFIG = {
    icon: {
        emoji: '🕵️‍♂️',
        opacity: {
            hidden: '0.1',
            visible: '0.3'
        },
        position: {
            top: '10px',
            right: '77px'  // 1.5cm (approximately 57px) + original 20px = 77px
        }
    },
    modal: {
        zIndex: 1000,
        maxHeight: '60vh'
    }
};

// ===== SPY ICON INJECTION =====

/**
 * Inject spy icon into page header
 * @param {string} containerId - ID of header container (optional)
 */
function injectSpyIcon(containerId = null) {
    const headerSelector = containerId ? `#${containerId}` : '.header';
    const header = document.querySelector(headerSelector);
    
    console.log('🕵️ Spy icon: Attempting injection with selector:', headerSelector);
    console.log('🕵️ Spy icon: Header found:', !!header);
    
    if (!header) {
        console.warn('🕵️ Spy icon: Header not found for injection');
        console.log('🕵️ Spy icon: Available elements with class "header":', document.querySelectorAll('.header'));
        return;
    }
    
    // Check if spy icon already exists
    const existingSpyIcon = header.querySelector('.spy-icon');
    if (existingSpyIcon) {
        console.log('🕵️ Spy icon: Already exists, skipping injection');
        return;
    }
    
    // Make header relative positioned if not already
    if (getComputedStyle(header).position === 'static') {
        header.style.position = 'relative';
    }
    
    // Create spy icon
    const spyIcon = document.createElement('div');
    spyIcon.className = 'spy-icon';
    spyIcon.onclick = showSpyModal;
    spyIcon.style.cssText = `
        position: absolute; 
        top: ${SPY_CONFIG.icon.position.top}; 
        right: ${SPY_CONFIG.icon.position.right}; 
        cursor: pointer; 
        opacity: ${SPY_CONFIG.icon.opacity.visible}; 
        transition: opacity 0.3s ease; 
        font-size: 1.5rem;
        z-index: 1100;
        pointer-events: auto;
    `;
    spyIcon.title = '🕵️ Cache Inspector';
    spyIcon.textContent = SPY_CONFIG.icon.emoji;
    
    // Hover effects
    spyIcon.onmouseover = () => spyIcon.style.opacity = '1.0';
    spyIcon.onmouseout = () => spyIcon.style.opacity = SPY_CONFIG.icon.opacity.visible;
    
    header.appendChild(spyIcon);
    
    console.log('🕵️ Spy icon: Element created and added to DOM');
    console.log('🕵️ Spy icon: Final styles:', spyIcon.style.cssText);
    console.log('🕵️ Spy icon: Element in DOM:', document.querySelector('.spy-icon'));
    
    logInfo('🕵️ Spy icon injected successfully');
}

// ===== SPY MODAL CORE FUNCTIONS =====

/**
 * Show the spy modal with current cache data
 */
function showSpyModal() {
    // Create modal if it doesn't exist
    let modal = document.getElementById('spyModal');
    if (!modal) {
        modal = createSpyModal();
        document.body.appendChild(modal);
    }
    
    // Populate with fresh cache data
    populateSpyModal();
    
    // Show modal
    modal.style.display = 'block';
    
    logInfo('🕵️ Spy modal opened');
}

/**
 * Create the spy modal DOM structure
 * @returns {HTMLElement} Modal element
 */
function createSpyModal() {
    const modal = document.createElement('div');
    modal.id = 'spyModal';
    modal.style.display = 'none';
    
    const appName = getAppName();
    
    modal.innerHTML = `
        <div style="position: fixed; top: 0; left: 0; width: 100%; height: 100%; 
                   background: rgba(0,0,0,0.8); z-index: ${SPY_CONFIG.modal.zIndex};" 
             onclick="closeSpyModal()">
            <div style="position: relative; top: 5%; left: 5%; width: 90%; height: 90%; 
                       background: white; border-radius: 8px; padding: 20px; 
                       overflow-y: auto; box-shadow: 0 4px 20px rgba(0,0,0,0.3);"
                 onclick="event.stopPropagation()">
                
                <!-- Header -->
                <div style="display: flex; justify-content: space-between; align-items: center; 
                           margin-bottom: 20px; border-bottom: 2px solid #e5e7eb; padding-bottom: 10px;">
                    <h2 style="margin: 0; color: #1f2937;">🕵️‍♂️ Cache Inspector - ${appName}</h2>
                    <button onclick="closeSpyModal()" 
                           style="background: #ef4444; color: white; border: none; 
                                  border-radius: 4px; padding: 8px 16px; cursor: pointer;">
                        ✕ Close
                    </button>
                </div>
                
                <!-- Cache Stats -->
                <div id="spyStats" style="background: #f3f4f6; padding: 15px; border-radius: 6px; margin-bottom: 20px;">
                    Loading cache statistics...
                </div>
                
                <!-- Tabs for different cache types -->
                <div style="margin-bottom: 20px;">
                    <button class="spy-tab-btn active" onclick="showSpyTab('summary')" 
                           style="background: #3b82f6; color: white; border: none; padding: 10px 20px; 
                                  margin-right: 10px; border-radius: 4px; cursor: pointer;">
                        📋 Summary
                    </button>
                    <button class="spy-tab-btn" onclick="showSpyTab('pending')" 
                           style="background: #6b7280; color: white; border: none; padding: 10px 20px; 
                                  margin-right: 10px; border-radius: 4px; cursor: pointer;">
                        ⏳ Pending
                    </button>
                    <button class="spy-tab-btn" onclick="showSpyTab('approved')"
                           style="background: #6b7280; color: white; border: none; padding: 10px 20px; 
                                  margin-right: 10px; border-radius: 4px; cursor: pointer;">
                        ✅ Approved
                    </button>
                    <button class="spy-tab-btn" onclick="showSpyTab('rejected')"
                           style="background: #6b7280; color: white; border: none; padding: 10px 20px; 
                                  margin-right: 10px; border-radius: 4px; cursor: pointer;">
                        ❌ Rejected
                    </button>
                    <button class="spy-tab-btn" onclick="showSpyTab('raw')"
                           style="background: #6b7280; color: white; border: none; padding: 10px 20px; 
                                  border-radius: 4px; cursor: pointer;">
                        🔧 Raw Data
                    </button>
                </div>
                
                <!-- Content Area -->
                <div id="spyContent" style="font-family: 'Arial', sans-serif; 
                                           background: #f8fafc; color: #1f2937; 
                                           padding: 15px; border-radius: 6px; 
                                           overflow-x: auto; max-height: ${SPY_CONFIG.modal.maxHeight};">
                    Loading cache data...
                </div>
                
                <!-- Actions -->
                <div id="spyActions" style="margin-top: 20px; text-align: center;">
                    <button onclick="exportSpyData()" 
                           style="background: #10b981; color: white; border: none; 
                                  padding: 10px 20px; margin-right: 10px; border-radius: 4px; cursor: pointer;">
                        📊 Export Cache to Excel
                    </button>
                    <!-- Clear cache button will be added dynamically if storage >80% -->
                    ${getAppSpecificButtons()}
                </div>
            </div>
        </div>
    `;
    
    return modal;
}

/**
 * Populate spy modal with current cache data
 */
function populateSpyModal() {
    try {
        console.log('🕵️ Spy modal: populateSpyModal called');
        
        if (typeof getCacheStats !== 'function') {
            console.error('🕵️ getCacheStats function not available');
            return;
        }
        
        const cacheStats = getCacheStats();
        console.log('🕵️ Cache stats:', cacheStats);
        
        const statsElement = document.getElementById('spyStats');
        
        if (statsElement) {
            statsElement.innerHTML = `
                <h3 style="margin: 0 0 10px 0; color: #374151;">📊 Cache Statistics</h3>
                <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 15px;">
                    <div>
                        <strong>Invoice Counts:</strong><br>
                        Pending: ${cacheStats.counts.pending}<br>
                        Approved: ${cacheStats.counts.approved}<br>
                        Rejected: ${cacheStats.counts.rejected}<br>
                        <strong>Total: ${cacheStats.counts.total}</strong>
                    </div>
                    <div>
                        <strong>Storage Usage:</strong><br>
                        Size: ${cacheStats.sizes.totalMB} MB<br>
                        Limit: ${cacheStats.limits.maxSizeMB} MB<br>
                        <strong>Used: ${cacheStats.limits.percentUsed}%</strong>
                    </div>
                    <div>
                        <strong>Last Updated:</strong><br>
                        ${new Date(cacheStats.lastUpdated).toLocaleString()}
                    </div>
                </div>
            `;
            console.log('🕵️ Stats populated successfully');
        } else {
            console.error('🕵️ spyStats element not found');
        }
        
        // Add clear cache button if storage >80%
        updateClearCacheButton(cacheStats);
        
        // Show summary by default
        showSpyTab('summary');
        
    } catch (error) {
        console.error('🕵️ Error populating spy modal:', error);
        const statsElement = document.getElementById('spyStats');
        if (statsElement) {
            statsElement.innerHTML = `<div style="color: red;">Error loading cache data: ${error.message}</div>`;
        }
    }
}

/**
 * Generate summary view of all invoices
 * @returns {Array} Array of invoice summaries
 */
function generateSummaryView() {
    try {
        const pending = getPendingInvoices() || [];
        const approved = getApprovedInvoices() || [];
        const rejected = getRejectedInvoices() || [];
        
        const allInvoices = [...pending, ...approved, ...rejected];
        
        return allInvoices.map(invoice => {
            const invoiceNumber = invoice.invoiceNumber || 'Unknown';
            let totalAmount = 0;
            let lineItemCount = 0;
            
            // Extract amount from different sources
            if (invoice.fullInvoiceData && Array.isArray(invoice.fullInvoiceData)) {
                totalAmount = invoice.fullInvoiceData[0]?.extractedInvoiceTotal || 0;
                lineItemCount = invoice.fullInvoiceData.length;
            } else if (invoice.summary) {
                totalAmount = invoice.summary.totalAmount || 0;
                lineItemCount = invoice.summary.lineItemCount || 0;
            }
            
            return {
                invoiceNumber: invoiceNumber,
                status: invoice.status || 'unknown',
                totalAmount: totalAmount,
                currency: invoice.fullInvoiceData?.[0]?.currency || invoice.summary?.currency || 'EUR',
                lineItems: lineItemCount,
                date: invoice.fullInvoiceData?.[0]?.dateOfInvoice || invoice.summary?.invoiceDate || 'Unknown',
                projectId: invoice.fullInvoiceData?.[0]?.projectId || invoice.summary?.projectId || 'Unknown',
                creditNote: invoice.fullInvoiceData?.[0]?.creditNote || invoice.summary?.creditNote || false
            };
        });
        
    } catch (error) {
        console.error('Error generating summary view:', error);
        return [{ error: 'Failed to generate summary: ' + error.message }];
    }
}

/**
 * Show specific tab content in spy modal
 * @param {string} tabType - Type of tab to show (pending|approved|rejected|raw)
 */
function showSpyTab(tabType) {
    try {
        console.log(`🕵️ Spy tab: switching to ${tabType}`);
        
        // Update tab button styles (only if called from a button click)
        if (typeof event !== 'undefined' && event.target) {
            document.querySelectorAll('.spy-tab-btn').forEach(btn => {
                btn.style.background = '#6b7280';
            });
            event.target.style.background = '#3b82f6';
        } else {
            // Called programmatically - find and update the right button
            const targetBtn = document.querySelector(`[onclick*="${tabType}"]`);
            if (targetBtn) {
                document.querySelectorAll('.spy-tab-btn').forEach(btn => {
                    btn.style.background = '#6b7280';
                });
                targetBtn.style.background = '#3b82f6';
            }
        }
        
        const contentElement = document.getElementById('spyContent');
        let data;
        
        switch(tabType) {
            case 'summary':
                data = generateSummaryView();
                break;
            case 'pending':
                if (typeof getPendingInvoices === 'function') {
                    data = getPendingInvoices();
                } else {
                    data = { error: 'getPendingInvoices function not available' };
                }
                break;
            case 'approved':
                if (typeof getApprovedInvoices === 'function') {
                    data = getApprovedInvoices();
                } else {
                    data = { error: 'getApprovedInvoices function not available' };
                }
                break;
            case 'rejected':
                if (typeof getRejectedInvoices === 'function') {
                    data = getRejectedInvoices();
                } else {
                    data = { error: 'getRejectedInvoices function not available' };
                }
                break;
            case 'raw':
                data = {
                    pending: localStorage.getItem('i2e_pending_invoices'),
                    approved: localStorage.getItem('i2e_approved_invoices'),
                    rejected: localStorage.getItem('i2e_rejected_invoices'),
                    preferences: localStorage.getItem('i2e_user_preferences')
                };
                break;
            default:
                data = { error: 'Unknown tab type: ' + tabType };
        }
        
        if (contentElement) {
            if (tabType === 'raw') {
                // Keep JSON format for raw data
                contentElement.innerHTML = '<pre>' + JSON.stringify(data, null, 2) + '</pre>';
            } else {
                // Format as readable table
                contentElement.innerHTML = formatDataAsTable(data, tabType);
            }
            console.log(`🕵️ Tab content updated for ${tabType}`);
        } else {
            console.error('🕵️ spyContent element not found');
        }
        
        if (typeof logInfo === 'function') {
            logInfo(`🕵️ Spy tab switched to: ${tabType}`);
        }
        
    } catch (error) {
        console.error(`🕵️ Error switching to tab ${tabType}:`, error);
        const contentElement = document.getElementById('spyContent');
        if (contentElement) {
            contentElement.innerHTML = `<pre style="color: red;">Error loading ${tabType} data: ${error.message}</pre>`;
        }
    }
}

/**
 * Format data as a readable table instead of JSON
 * @param {Array|Object} data - Data to format
 * @param {string} tabType - Type of tab for context
 * @returns {string} HTML table string
 */
function formatDataAsTable(data, tabType) {
    try {
        if (!data || (Array.isArray(data) && data.length === 0)) {
            return `<div style="text-align: center; padding: 2rem; color: #6b7280;">
                        <div style="font-size: 1.2rem; margin-bottom: 0.5rem;">📭</div>
                        <div>No ${tabType} invoices found</div>
                    </div>`;
        }
        
        if (data.error) {
            return `<div style="color: red; padding: 1rem;">Error: ${data.error}</div>`;
        }
        
        if (tabType === 'summary') {
            return formatSummaryTable(data);
        }
        
        // For pending, approved, rejected
        if (Array.isArray(data)) {
            return formatInvoiceTable(data, tabType);
        }
        
        // Fallback to JSON if we can't format it
        return '<pre>' + JSON.stringify(data, null, 2) + '</pre>';
        
    } catch (error) {
        return `<div style="color: red; padding: 1rem;">Error formatting data: ${error.message}</div>`;
    }
}

/**
 * Format summary data as a simple table
 */
function formatSummaryTable(data) {
    if (!Array.isArray(data) || data.length === 0) {
        return '<div style="color: #6b7280; padding: 1rem;">No summary data available</div>';
    }
    
    let html = `
        <table style="width: 100%; border-collapse: collapse; background: white;">
            <thead>
                <tr style="background: #f8fafc; border-bottom: 2px solid #e2e8f0;">
                    <th style="padding: 0.75rem; text-align: left; border: 1px solid #e2e8f0;">Invoice #</th>
                    <th style="padding: 0.75rem; text-align: left; border: 1px solid #e2e8f0;">Status</th>
                    <th style="padding: 0.75rem; text-align: left; border: 1px solid #e2e8f0;">Amount</th>
                    <th style="padding: 0.75rem; text-align: left; border: 1px solid #e2e8f0;">Date</th>
                    <th style="padding: 0.75rem; text-align: left; border: 1px solid #e2e8f0;">Project</th>
                    <th style="padding: 0.75rem; text-align: left; border: 1px solid #e2e8f0;">Items</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    data.forEach(invoice => {
        const statusColor = getStatusColor(invoice.status);
        const creditNote = invoice.creditNote ? ' (Credit Note)' : '';
        
        html += `
            <tr style="border-bottom: 1px solid #e2e8f0;">
                <td style="padding: 0.75rem; border: 1px solid #e2e8f0;"><strong>${invoice.invoiceNumber}</strong></td>
                <td style="padding: 0.75rem; border: 1px solid #e2e8f0;">
                    <span style="background: ${statusColor}; color: white; padding: 0.25rem 0.5rem; border-radius: 4px; font-size: 0.8rem;">
                        ${invoice.status.toUpperCase()}
                    </span>
                </td>
                <td style="padding: 0.75rem; border: 1px solid #e2e8f0;">${formatAmount(invoice.totalAmount)} ${invoice.currency}${creditNote}</td>
                <td style="padding: 0.75rem; border: 1px solid #e2e8f0;">${invoice.date}</td>
                <td style="padding: 0.75rem; border: 1px solid #e2e8f0;">${invoice.projectId}</td>
                <td style="padding: 0.75rem; border: 1px solid #e2e8f0;">${invoice.lineItems} items</td>
            </tr>
        `;
    });
    
    html += '</tbody></table>';
    return html;
}

/**
 * Format invoice data as a simple table
 */
function formatInvoiceTable(data, tabType) {
    if (!Array.isArray(data) || data.length === 0) {
        return `<div style="color: #6b7280; padding: 1rem;">No ${tabType} invoices found</div>`;
    }
    
    let html = `
        <table style="width: 100%; border-collapse: collapse; background: white;">
            <thead>
                <tr style="background: #f8fafc; border-bottom: 2px solid #e2e8f0;">
                    <th style="padding: 0.75rem; text-align: left; border: 1px solid #e2e8f0;">Invoice #</th>
                    <th style="padding: 0.75rem; text-align: left; border: 1px solid #e2e8f0;">Status</th>
                    <th style="padding: 0.75rem; text-align: left; border: 1px solid #e2e8f0;">Amount</th>
                    <th style="padding: 0.75rem; text-align: left; border: 1px solid #e2e8f0;">Date</th>
                    <th style="padding: 0.75rem; text-align: left; border: 1px solid #e2e8f0;">Project</th>
    `;
    
    if (tabType === 'pending') {
        html += '<th style="padding: 0.75rem; text-align: left; border: 1px solid #e2e8f0;">Source</th>';
    } else {
        html += '<th style="padding: 0.75rem; text-align: left; border: 1px solid #e2e8f0;">Action Date</th>';
    }
    
    html += `
                </tr>
            </thead>
            <tbody>
    `;
    
    data.forEach(invoice => {
        const invoiceNumber = invoice.invoiceNumber || 'Unknown';
        const statusColor = getStatusColor(invoice.status);
        let amount = 'Unknown';
        let date = 'Unknown';
        let projectId = 'Unknown';
        
        // Extract data from different formats
        if (invoice.fullInvoiceData && invoice.fullInvoiceData[0]) {
            amount = `${formatAmount(invoice.fullInvoiceData[0].extractedInvoiceTotal)} ${invoice.fullInvoiceData[0].currency || 'EUR'}`;
            date = invoice.fullInvoiceData[0].dateOfInvoice || 'Unknown';
            projectId = invoice.fullInvoiceData[0].projectId || 'Unknown';
        } else if (invoice.summary) {
            amount = `${formatAmount(invoice.summary.totalAmount)} ${invoice.summary.currency || 'EUR'}`;
            date = invoice.summary.invoiceDate || 'Unknown';
            projectId = invoice.summary.projectId || 'Unknown';
        }
        
        const creditNote = invoice.fullInvoiceData?.[0]?.creditNote || invoice.summary?.creditNote ? ' (Credit Note)' : '';
        
        // Make date editable for pending invoices (for testing month filtering)
        const dateDisplay = tabType === 'pending' && date !== 'Unknown' ? 
            `<input type="text" value="${date}" onchange="updateInvoiceDate('${invoiceNumber}', this.value)" 
                    style="width: 100%; border: 1px solid #d1d5db; padding: 0.25rem; font-size: 0.8rem;" 
                    title="Edit date to test month filtering (format: dd.mm.yyyy)">` : date;
        
        html += `
            <tr style="border-bottom: 1px solid #e2e8f0;">
                <td style="padding: 0.75rem; border: 1px solid #e2e8f0;"><strong>${invoiceNumber}</strong></td>
                <td style="padding: 0.75rem; border: 1px solid #e2e8f0;">
                    <span style="background: ${statusColor}; color: white; padding: 0.25rem 0.5rem; border-radius: 4px; font-size: 0.8rem;">
                        ${invoice.status.toUpperCase()}
                    </span>
                </td>
                <td style="padding: 0.75rem; border: 1px solid #e2e8f0;">${amount}${creditNote}</td>
                <td style="padding: 0.75rem; border: 1px solid #e2e8f0;">${dateDisplay}</td>
                <td style="padding: 0.75rem; border: 1px solid #e2e8f0;">${projectId}</td>
        `;
        
        if (tabType === 'pending') {
            const sourceIcon = invoice.source === 'new' ? '🆕' : '💾';
            const sourceText = invoice.source === 'new' ? 'New' : 'Cache';
            html += `<td style="padding: 0.75rem; border: 1px solid #e2e8f0;">${sourceIcon} ${sourceText}</td>`;
        } else {
            const actionDate = invoice.approvalDate || invoice.rejectionDate || 'Unknown';
            html += `<td style="padding: 0.75rem; border: 1px solid #e2e8f0;">${formatDate(actionDate)}</td>`;
        }
        
        html += '</tr>';
    });
    
    html += '</tbody></table>';
    return html;
}

/**
 * Get status color for badges
 */
function getStatusColor(status) {
    switch(status) {
        case 'pending': return '#f59e0b';
        case 'approved': return '#10b981';
        case 'rejected': return '#ef4444';
        default: return '#6b7280';
    }
}

/**
 * Format amount with commas
 */
function formatAmount(amount) {
    if (typeof amount !== 'number') return amount || '0';
    return amount.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

/**
 * Format ISO date string to readable format
 */
function formatDate(dateString) {
    if (!dateString || dateString === 'Unknown') return 'Unknown';
    try {
        const date = new Date(dateString);
        return date.toLocaleDateString() + ' ' + date.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
    } catch (error) {
        return dateString;
    }
}

/**
 * Update invoice date for testing month filtering
 * @param {string} invoiceNumber - Invoice number to update
 * @param {string} newDate - New date in dd.mm.yyyy format
 */
function updateInvoiceDate(invoiceNumber, newDate) {
    try {
        console.log(`🕵️ Updating invoice ${invoiceNumber} date to: ${newDate}`);
        
        // Validate date format (basic check)
        const dateRegex = /^\d{2}\.\d{2}\.\d{4}$/;
        if (!dateRegex.test(newDate)) {
            alert('Please use date format: dd.mm.yyyy (e.g., 15.04.2025)');
            return;
        }
        
        // Get pending invoices
        const pendingInvoices = getPendingInvoices();
        const invoiceIndex = pendingInvoices.findIndex(inv => inv.invoiceNumber === invoiceNumber);
        
        if (invoiceIndex === -1) {
            console.error(`🕵️ Invoice ${invoiceNumber} not found in pending invoices`);
            return;
        }
        
        // Update the date in the full invoice data
        const invoice = pendingInvoices[invoiceIndex];
        if (invoice.fullInvoiceData && invoice.fullInvoiceData[0]) {
            invoice.fullInvoiceData[0].dateOfInvoice = newDate;
            
            // Also update month of invoice based on new date
            const dateParts = newDate.split('.');
            if (dateParts.length >= 3) {
                const monthNum = parseInt(dateParts[1]);
                const year = dateParts[2];
                const monthNames = ['', 'January', 'February', 'March', 'April', 'May', 'June', 
                                  'July', 'August', 'September', 'October', 'November', 'December'];
                const monthName = monthNames[monthNum] || 'Unknown';
                invoice.fullInvoiceData[0].monthOfInvoice = `${monthName} ${year}`;
                
                console.log(`🕵️ Updated month of invoice to: ${monthName} ${year}`);
            }
            
            // Update the cache
            if (typeof saveToLocalStorage === 'function') {
                saveToLocalStorage('i2e_pending_invoices', pendingInvoices);
            } else {
                // Fallback: direct localStorage update
                localStorage.setItem('i2e_pending_invoices', JSON.stringify(pendingInvoices));
            }
            
            console.log(`✅ Invoice ${invoiceNumber} date updated successfully`);
            
            // Refresh the spy modal display
            showSpyTab('pending');
            
            // If we're in the validator, refresh the pending invoices table too
            if (typeof displayPendingInvoices === 'function' && typeof getPendingInvoices === 'function') {
                displayPendingInvoices(getPendingInvoices());
            }
            
        } else {
            console.error(`🕵️ Invoice ${invoiceNumber} has no fullInvoiceData to update`);
        }
        
    } catch (error) {
        console.error('🕵️ Error updating invoice date:', error);
        alert('Error updating date: ' + error.message);
    }
}

/**
 * Update clear cache button visibility based on storage usage
 * @param {Object} cacheStats - Cache statistics object
 */
function updateClearCacheButton(cacheStats) {
    const actionsElement = document.getElementById('spyActions');
    if (!actionsElement) return;
    
    const percentUsed = parseFloat(cacheStats.limits?.percentUsed || 0);
    const existingButton = actionsElement.querySelector('.clear-cache-btn');
    
    if (percentUsed >= 0) {
        // Show clear cache button always (80% rule disabled for testing)
        if (!existingButton) {
            const clearButton = document.createElement('button');
            clearButton.className = 'clear-cache-btn';
            clearButton.onclick = clearCacheWithConfirm;
            clearButton.style.cssText = `
                background: #ef4444; color: white; border: none; 
                padding: 10px 20px; margin-right: 10px; border-radius: 4px; cursor: pointer;
            `;
            clearButton.innerHTML = '🗑️ Clear All Cache';
            
            // Insert before the last button (app-specific buttons)
            const exportButton = actionsElement.querySelector('button');
            exportButton.parentNode.insertBefore(clearButton, exportButton.nextSibling);
        }
        
        // Add Excel cache clearing button
        const existingExcelButton = actionsElement.querySelector('.clear-excel-btn');
        if (!existingExcelButton) {
            const clearExcelButton = document.createElement('button');
            clearExcelButton.className = 'clear-excel-btn';
            clearExcelButton.onclick = clearExcelCacheWithConfirm;
            clearExcelButton.style.cssText = `
                background: #f59e0b; color: white; border: none; 
                padding: 10px 20px; margin-right: 10px; border-radius: 4px; cursor: pointer;
            `;
            clearExcelButton.innerHTML = '📊 Clear Excel Files';
            
            // Insert after the clear cache button
            const clearCacheBtn = actionsElement.querySelector('.clear-cache-btn');
            if (clearCacheBtn) {
                clearCacheBtn.parentNode.insertBefore(clearExcelButton, clearCacheBtn.nextSibling);
            }
        }
    }
    // Note: 80% rule disabled - button always shown for testing
}

/**
 * Close the spy modal
 */
function closeSpyModal() {
    const modal = document.getElementById('spyModal');
    if (modal) {
        modal.style.display = 'none';
        logInfo('🕵️ Spy modal closed');
    }
}

// ===== SPY ACTIONS =====

/**
 * Export all cache data to Excel file
 */
async function exportSpyData() {
    try {
        const appName = getAppName();
        const timestamp = new Date().toISOString();
        
        // Check if ExcelJS is available
        if (typeof ExcelJS === 'undefined') {
            alert('❌ ExcelJS library not available. Excel export requires ExcelJS to be loaded.');
            logError('🕵️ ExcelJS not available for Excel export');
            return;
        }
        
        // Collect all cache data
        const pendingInvoices = getPendingInvoices() || [];
        const approvedInvoices = getApprovedInvoices() || [];
        const rejectedInvoices = getRejectedInvoices() || [];
        
        console.log('🕵️ Export Debug - Pending invoices:', pendingInvoices.length);
        console.log('🕵️ Export Debug - Approved invoices:', approvedInvoices.length);
        console.log('🕵️ Export Debug - Rejected invoices:', rejectedInvoices.length);
        
        // Flatten all invoice data for Excel export
        const allInvoiceData = [];
        
        // Add pending invoices
        pendingInvoices.forEach((invoice, invoiceIndex) => {
            console.log(`🕵️ Export Debug - Processing pending invoice ${invoiceIndex}:`, invoice.invoiceNumber);
            
            if (invoice.fullInvoiceData && Array.isArray(invoice.fullInvoiceData)) {
                let lineItems = invoice.fullInvoiceData;
                console.log(`🕵️ Export Debug - Initial array length: ${lineItems.length}`);
                
                // Handle nested fullInvoiceData structure (same as validator code)
                if (lineItems.length === 1 && lineItems[0].fullInvoiceData && Array.isArray(lineItems[0].fullInvoiceData)) {
                    console.log('🕵️ Export Debug - Detected nested structure, using inner fullInvoiceData');
                    lineItems = lineItems[0].fullInvoiceData;
                }
                
                console.log(`🕵️ Export Debug - Final line items count: ${lineItems.length}`);
                
                lineItems.forEach((lineItem, lineIndex) => {
                    console.log(`🕵️ Export Debug - Adding line item ${lineIndex}: ${lineItem.positionDescription || 'No description'}`);
                    allInvoiceData.push({
                        ...lineItem,
                        cacheStatus: 'Pending',
                        cacheSource: invoice.source || 'Unknown'
                    });
                });
            } else {
                console.log('🕵️ Export Debug - Invoice has no fullInvoiceData or not an array');
            }
        });
        
        // Add approved invoices
        approvedInvoices.forEach(invoice => {
            if (invoice.fullInvoiceData && Array.isArray(invoice.fullInvoiceData)) {
                let lineItems = invoice.fullInvoiceData;
                
                // Handle nested fullInvoiceData structure
                if (lineItems.length === 1 && lineItems[0].fullInvoiceData && Array.isArray(lineItems[0].fullInvoiceData)) {
                    lineItems = lineItems[0].fullInvoiceData;
                }
                
                lineItems.forEach(lineItem => {
                    allInvoiceData.push({
                        ...lineItem,
                        cacheStatus: 'Approved',
                        approvalDate: invoice.approvalDate || 'Unknown'
                    });
                });
            }
        });
        
        // Add rejected invoices
        rejectedInvoices.forEach(invoice => {
            if (invoice.fullInvoiceData && Array.isArray(invoice.fullInvoiceData)) {
                let lineItems = invoice.fullInvoiceData;
                
                // Handle nested fullInvoiceData structure
                if (lineItems.length === 1 && lineItems[0].fullInvoiceData && Array.isArray(lineItems[0].fullInvoiceData)) {
                    lineItems = lineItems[0].fullInvoiceData;
                }
                
                lineItems.forEach(lineItem => {
                    allInvoiceData.push({
                        ...lineItem,
                        cacheStatus: 'Rejected',
                        rejectionDate: invoice.rejectionDate || 'Unknown'
                    });
                });
            }
        });
        
        console.log(`🕵️ Export Debug - Total line items collected: ${allInvoiceData.length}`);
        
        if (allInvoiceData.length === 0) {
            alert('No invoice data found in cache to export.');
            return;
        }
        
        // Create Excel workbook
        const workbook = new ExcelJS.Workbook();
        
        // Create main data sheet
        const dataSheet = workbook.addWorksheet('Cache Data');
        
        // Define headers for cache export
        const headers = [
            'Cache Status', 'Invoice Number', 'Project ID', 'Customer ID', 
            'Invoice Date', 'Month of Invoice', 'Service Period', 'Position',
            'Description', 'Quantity', 'Unit Price', 'Position Total',
            'Cost Type', 'Currency', 'VAT', 'Credit Note', 'Page Number'
        ];
        
        // Add conditional headers based on status
        if (pendingInvoices.length > 0) headers.push('Cache Source');
        if (approvedInvoices.length > 0) headers.push('Approval Date');
        if (rejectedInvoices.length > 0) headers.push('Rejection Date');
        
        dataSheet.addRow(headers);
        
        // Style headers
        const headerRow = dataSheet.getRow(1);
        headerRow.font = { bold: true };
        headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F8FAFC' } };
        
        // Add data rows
        allInvoiceData.forEach(row => {
            const dataRow = [
                row.cacheStatus || '',
                row.invoiceNumber || '',
                row.projectId || '',
                row.customerId || '',
                row.dateOfInvoice || '',
                row.monthOfInvoice || '',
                row.serviceProvisionPeriod || '',
                row.position || '',
                row.positionDescription || '',
                row.positionQuantity || 1,
                row.unitPrice || 0,
                row.positionTotal || 0,
                row.typeCost || '',
                row.currency || '',
                row.vat || '',
                row.creditNote ? 'Yes' : 'No',
                row.pageNumber || ''
            ];
            
            // Add conditional data
            if (pendingInvoices.length > 0) dataRow.push(row.cacheSource || '');
            if (approvedInvoices.length > 0) dataRow.push(row.approvalDate || '');
            if (rejectedInvoices.length > 0) dataRow.push(row.rejectionDate || '');
            
            const excelRow = dataSheet.addRow(dataRow);
            
            // Color coding by status
            let fillColor = '';
            switch (row.cacheStatus) {
                case 'Pending': fillColor = 'FFF3CD'; break;  // Light yellow
                case 'Approved': fillColor = 'D1F2EB'; break; // Light green
                case 'Rejected': fillColor = 'F8D7DA'; break; // Light red
            }
            
            if (fillColor) {
                excelRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
            }
            
            // Highlight credit notes in baby blue
            if (row.creditNote) {
                excelRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'B0E6FF' } };
            }
        });
        
        // Auto-fit columns
        autoFitColumns(dataSheet);
        
        // Format amount columns
        const unitPriceColumn = dataSheet.getColumn(11);
        const totalColumn = dataSheet.getColumn(12);
        unitPriceColumn.numFmt = '#,##0.00';
        totalColumn.numFmt = '#,##0.00';
        
        // Add filters
        const lastColumnLetter = String.fromCharCode(65 + headers.length - 1);
        dataSheet.autoFilter = `A1:${lastColumnLetter}1`;
        
        // Create summary sheet
        createCacheSummarySheet(workbook, pendingInvoices, approvedInvoices, rejectedInvoices);
        
        // Generate filename and save
        const filename = generateFilename('I2E_Cache_Export');
        await saveWorkbook(workbook, filename);
        
        console.log(`✅ Cache data exported to Excel successfully from ${appName}!`);
        logInfo(`🕵️ Cache data exported to Excel from ${appName}`);
        
        // Offer to clear cache if storage is >80% full
        const stats = getCacheStats();
        const percentUsed = parseFloat(stats.limits?.percentUsed || 0);
        
        if (percentUsed > 80) {
            if (confirm(`📊 Excel export successful!\n\nStorage usage: ${percentUsed}% (${stats.sizes.totalMB} MB)\n\nWould you like to clear the cache to free up space?\n\n⚠️ This will delete all cached invoice data.`)) {
                try {
                    if (await clearCache()) {
                        console.log('✅ Cache cleared successfully after export!');
                        closeSpyModal();
                        
                        // Call app-specific refresh function if available
                        if (typeof refreshAfterCacheClear === 'function') {
                            refreshAfterCacheClear();
                        } else {
                            // Default: reload page
                            location.reload();
                        }
                    }
                } catch (error) {
                    console.error('❌ Error clearing cache after export:', error);
                    alert('❌ Error clearing cache after export!');
                }
            }
        }
        
    } catch (error) {
        alert('❌ Error exporting cache data to Excel: ' + error.message);
        logError('🕵️ Excel export error:', error);
    }
}

/**
 * Create cache summary sheet for Excel export
 * @param {ExcelJS.Workbook} workbook - Workbook instance
 * @param {Array} pendingInvoices - Pending invoices
 * @param {Array} approvedInvoices - Approved invoices  
 * @param {Array} rejectedInvoices - Rejected invoices
 */
function createCacheSummarySheet(workbook, pendingInvoices, approvedInvoices, rejectedInvoices) {
    const summarySheet = workbook.addWorksheet('Cache Summary');
    
    // Calculate totals
    const stats = {
        pending: { count: pendingInvoices.length, total: 0 },
        approved: { count: approvedInvoices.length, total: 0 },
        rejected: { count: rejectedInvoices.length, total: 0 }
    };
    
    // Calculate totals for each status
    pendingInvoices.forEach(invoice => {
        if (invoice.fullInvoiceData && Array.isArray(invoice.fullInvoiceData)) {
            // Use the first line item's extractedInvoiceTotal (it's the same for all line items from the same invoice)
            const firstLineItem = invoice.fullInvoiceData[0];
            if (firstLineItem && firstLineItem.extractedInvoiceTotal) {
                stats.pending.total += firstLineItem.extractedInvoiceTotal;
            }
        }
    });
    
    approvedInvoices.forEach(invoice => {
        if (invoice.fullInvoiceData && Array.isArray(invoice.fullInvoiceData)) {
            // Use the first line item's extractedInvoiceTotal (it's the same for all line items from the same invoice)
            const firstLineItem = invoice.fullInvoiceData[0];
            if (firstLineItem && firstLineItem.extractedInvoiceTotal) {
                stats.approved.total += firstLineItem.extractedInvoiceTotal;
            }
        }
    });
    
    rejectedInvoices.forEach(invoice => {
        if (invoice.fullInvoiceData && Array.isArray(invoice.fullInvoiceData)) {
            // Use the first line item's extractedInvoiceTotal (it's the same for all line items from the same invoice)
            const firstLineItem = invoice.fullInvoiceData[0];
            if (firstLineItem && firstLineItem.extractedInvoiceTotal) {
                stats.rejected.total += firstLineItem.extractedInvoiceTotal;
            }
        }
    });
    
    // Add summary headers
    const headers = ['Status', 'Invoice Count', 'Total Amount (EUR)', 'Percentage'];
    summarySheet.addRow(headers);
    
    // Style headers
    const headerRow = summarySheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F8FAFC' } };
    
    // Add summary data
    const totalInvoices = stats.pending.count + stats.approved.count + stats.rejected.count;
    const totalAmount = stats.pending.total + stats.approved.total + stats.rejected.total;
    
    // Pending row
    const pendingRow = summarySheet.addRow([
        'Pending',
        stats.pending.count,
        stats.pending.total,
        totalInvoices > 0 ? `${((stats.pending.count / totalInvoices) * 100).toFixed(1)}%` : '0%'
    ]);
    pendingRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF3CD' } };
    
    // Approved row
    const approvedRow = summarySheet.addRow([
        'Approved',
        stats.approved.count,
        stats.approved.total,
        totalInvoices > 0 ? `${((stats.approved.count / totalInvoices) * 100).toFixed(1)}%` : '0%'
    ]);
    approvedRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D1F2EB' } };
    
    // Rejected row
    const rejectedRow = summarySheet.addRow([
        'Rejected',
        stats.rejected.count,
        stats.rejected.total,
        totalInvoices > 0 ? `${((stats.rejected.count / totalInvoices) * 100).toFixed(1)}%` : '0%'
    ]);
    rejectedRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F8D7DA' } };
    
    // Total row
    const totalRow = summarySheet.addRow([
        'TOTAL',
        totalInvoices,
        totalAmount,
        '100%'
    ]);
    totalRow.font = { bold: true };
    totalRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'E5E7EB' } };
    
    // Auto-fit columns
    autoFitColumns(summarySheet);
    
    // Format amount column
    const amountColumn = summarySheet.getColumn(3);
    amountColumn.numFmt = '#,##0.00';
}

/**
 * Auto-fit columns based on content (utility function for Excel export)
 * @param {ExcelJS.Worksheet} worksheet - Worksheet to format
 * @param {number} maxWidth - Maximum column width (default: 50)
 */
function autoFitColumns(worksheet, maxWidth = 50) {
    worksheet.columns.forEach(column => {
        let maxLength = 0;
        column.eachCell({ includeEmpty: true }, (cell) => {
            const cellLength = cell.value ? cell.value.toString().length : 0;
            if (cellLength > maxLength) {
                maxLength = cellLength;
            }
        });
        column.width = Math.min(maxLength + 2, maxWidth);
    });
}

/**
 * Generate filename with timestamp (utility function for Excel export)
 * @param {string} prefix - Filename prefix
 * @returns {string} Generated filename
 */
function generateFilename(prefix) {
    const timestamp = new Date().toISOString().slice(0, 10);
    return `${prefix}_${timestamp}.xlsx`;
}

/**
 * Save workbook to file and trigger download (utility function for Excel export)
 * @param {ExcelJS.Workbook} workbook - Workbook to save
 * @param {string} filename - Filename for download
 * @returns {Promise<void>}
 */
async function saveWorkbook(workbook, filename) {
    // Generate buffer
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { 
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' 
    });
    
    // Download file
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

/**
 * Clear all cache with confirmation
 */
async function clearCacheWithConfirm() {
    if (confirm('🚨 Are you sure you want to clear ALL cache data?\n\nThis will delete all pending, approved, and rejected invoices.\n\nThis action cannot be undone!')) {
        try {
            if (await clearCache()) {
                console.log('✅ Cache cleared successfully!');
                closeSpyModal();
                
                // Call app-specific refresh function if available
                if (typeof refreshAfterCacheClear === 'function') {
                    refreshAfterCacheClear();
                } else {
                    // Default: reload page
                    location.reload();
                }
                
                logInfo('🕵️ Cache cleared successfully');
            } else {
                console.error('❌ Error clearing cache!');
                logError('🕵️ Cache clear failed');
            }
        } catch (error) {
            console.error('❌ Error clearing cache:', error);
            alert('Error clearing cache: ' + error.message);
            logError('🕵️ Cache clear failed');
        }
    }
}

/**
 * Clear Excel cache with confirmation
 */
async function clearExcelCacheWithConfirm() {
    if (confirm('📊 Are you sure you want to clear Excel file cache?\n\nThis will delete all uploaded cost data files (PPM, EXT SAP, I2E data).\n\nThis action cannot be undone!')) {
        try {
            // Clear the Excel/cost data cache from localStorage (legacy)
            localStorage.removeItem('i2e_cost_data_cache');
            
            // Also clear from IndexedDB (where it's actually stored now)
            if (typeof removeFromIndexedDB === 'function') {
                await removeFromIndexedDB('i2e_cost_data_cache');
                console.log('✅ Excel cache cleared from IndexedDB!');
            }
            
            console.log('✅ Excel cache cleared successfully!');
            closeSpyModal();
            
            // Call app-specific refresh function if available
            if (typeof refreshAfterCacheClear === 'function') {
                refreshAfterCacheClear();
            } else {
                // Default: reload page
                location.reload();
            }
            
            logInfo('🕵️ Excel cache cleared successfully (localStorage + IndexedDB)');
        } catch (error) {
            console.error('❌ Error clearing Excel cache:', error);
            alert('Error clearing Excel cache: ' + error.message);
            logError('🕵️ Excel cache clear failed');
        }
    }
}

// ===== UTILITY FUNCTIONS =====

/**
 * Get current application name based on page title or URL
 * @returns {string} Application name
 */
function getAppName() {
    const title = document.title;
    if (title.includes('Processor')) return 'Invoice Processor';
    if (title.includes('Validator')) return 'Invoice Validator';
    
    const path = window.location.pathname;
    if (path.includes('Processor')) return 'Invoice Processor';
    if (path.includes('Validator')) return 'Invoice Validator';
    
    return 'I2E Application';
}

/**
 * Get app-specific buttons for the modal
 * @returns {string} HTML for app-specific buttons
 */
function getAppSpecificButtons() {
    // No app-specific buttons needed
    return '';
}

// ===== GLOBAL AVAILABILITY =====

// Make functions globally available for browser usage
if (typeof window !== 'undefined') {
    // Core functions
    window.injectSpyIcon = injectSpyIcon;
    window.showSpyModal = showSpyModal;
    window.closeSpyModal = closeSpyModal;
    window.showSpyTab = showSpyTab;
    window.exportSpyData = exportSpyData;
    window.clearCacheWithConfirm = clearCacheWithConfirm;
    window.updateInvoiceDate = updateInvoiceDate;
    
    logInfo('🕵️ I2E Spy Module loaded successfully');
}

// Export for module usage
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        injectSpyIcon,
        showSpyModal,
        closeSpyModal,
        showSpyTab,
        exportSpyData,
        clearCacheWithConfirm,
        SPY_CONFIG
    };
}