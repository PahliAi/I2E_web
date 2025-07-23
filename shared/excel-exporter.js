/**
 * I2E Excel Exporter
 * Excel file generation and export functionality for I2E Invoice Processor
 * 
 * @version 1.0
 * @author I2E Development Team
 * @requires ExcelJS library
 * @requires i2e-common.js for utilities
 */

// ===== MAIN EXPORT FUNCTIONS =====

/**
 * Export data to Excel with standard format (all fields)
 * @param {Array} data - Array of extracted invoice data
 * @param {string} filename - Optional custom filename
 * @returns {Promise<void>} 
 */
async function exportToExcel(data, filename = null) {
    if (!data || data.length === 0) {
        alert('No data to export. Please process some files first.');
        return;
    }
    
    try {
        logInfo(`📊 Exporting ${data.length} items to Excel`);
        
        // Create workbook
        const workbook = new ExcelJS.Workbook();
        
        // Create "Invoice Details" sheet
        const detailsSheet = workbook.addWorksheet('Invoice Details');
        
        // Add headers
        const headers = [
            'ProjectID', 'CustomerID', 'InvoiceNumber', 'InvoiceDate', 'MonthOfInvoice',
            'CreditNote', 'ServiceProvisionPeriod', 'CostType', 'PositionDescription',
            'PositionQuantity', 'PositionTotal'
        ];
        
        detailsSheet.addRow(headers);
        
        // Style headers
        const headerRow = detailsSheet.getRow(1);
        headerRow.font = { bold: true };
        headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F8FAFC' } };
        
        // Add data rows
        data.forEach(row => {
            const dataRow = [
                row.projectId || '',
                row.customerId || '',
                row.invoiceNumber || '',
                row.dateOfInvoice || '',
                row.monthOfInvoice || '',
                row.creditNote ? 'Yes' : 'No',
                row.serviceProvisionPeriod || '',
                row.typeCost || '',
                row.positionDescription || '',
                row.positionQuantity || 1,
                row.positionTotal || 0
            ];
            
            const excelRow = detailsSheet.addRow(dataRow);
            
            // Highlight validation error rows in red
            if (row.hasValidationError) {
                excelRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FEE2E2' } };
            }
            // Highlight credit notes in baby blue
            else if (row.creditNote) {
                excelRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'B0E6FF' } };
            }
        });
        
        // Auto-fit columns
        autoFitColumns(detailsSheet);
        
        // Format amount column
        const totalColumn = detailsSheet.getColumn(11);
        totalColumn.numFmt = '#,##0.00';
        
        // Add filters
        detailsSheet.autoFilter = 'A1:K1';
        
        // Create summary sheet
        createInvoiceSummarySheet(workbook, data);
        
        // Generate filename and save
        const finalFilename = filename || generateFilename('I2E_Invoice_Data');
        await saveWorkbook(workbook, finalFilename);
        
        logInfo('✅ Excel file exported successfully!');
        
    } catch (error) {
        logError('Error exporting to Excel:', error);
        alert('An error occurred while exporting to Excel. Please try again.');
    }
}

/**
 * Export data to Excel with selected fields
 * @param {Array} data - Array of extracted invoice data
 * @param {Array} selectedFields - Array of field keys to include
 * @param {Object} allFields - Field definitions object
 * @param {string} filename - Optional custom filename
 * @returns {Promise<void>}
 */
async function exportToExcelWithFields(data, selectedFields, allFields, filename = null) {
    if (!data || data.length === 0) {
        alert('No data to export. Please process some files first.');
        return;
    }
    
    if (!selectedFields || selectedFields.length === 0) {
        alert('Please select at least one field to export.');
        return;
    }
    
    try {
        logInfo(`📊 Exporting with selected fields: ${selectedFields.join(', ')}`);
        
        // Create workbook
        const workbook = new ExcelJS.Workbook();
        
        // Create "Invoice Details" sheet
        const detailsSheet = workbook.addWorksheet('Invoice Details');
        
        // Add headers based on selected fields
        const headers = selectedFields.map(fieldKey => allFields[fieldKey].name);
        detailsSheet.addRow(headers);
        
        // Style headers
        const headerRow = detailsSheet.getRow(1);
        headerRow.font = { bold: true };
        headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F8FAFC' } };
        
        // Add data rows
        data.forEach(row => {
            const dataRow = selectedFields.map(fieldKey => {
                let value = row[fieldKey];
                
                // Format specific fields
                if (fieldKey === 'creditNote') {
                    value = value ? 'Yes' : 'No';
                } else if (fieldKey === 'positionTotal' || fieldKey === 'unitPrice') {
                    value = value || 0;
                } else if (value === null || value === undefined) {
                    value = '';
                }
                
                return value;
            });
            
            const excelRow = detailsSheet.addRow(dataRow);
            
            // Highlight validation error rows in red, credit notes in baby-blue
            if (row.hasValidationError) {
                excelRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FEE2E2' } };
            } else if (row.creditNote) {
                excelRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'B0E6FF' } };
            }
        });
        
        // Auto-fit columns
        autoFitColumnsWithHeaders(detailsSheet, headers);
        
        // Format amount columns
        selectedFields.forEach((fieldKey, index) => {
            if (fieldKey === 'positionTotal' || fieldKey === 'unitPrice') {
                const column = detailsSheet.getColumn(index + 1);
                column.numFmt = '#,##0.00';
            }
        });
        
        // Add filters
        const lastColumn = String.fromCharCode(65 + selectedFields.length - 1);
        detailsSheet.autoFilter = `A1:${lastColumn}1`;
        
        // Create summary sheet if relevant fields are selected
        if (selectedFields.includes('invoiceNumber') && 
            selectedFields.includes('typeCost') && 
            selectedFields.includes('positionTotal')) {
            createInvoiceSummarySheet(workbook, data);
        }
        
        // Create invoice status sheet if status data is available
        if (selectedFields.includes('approvalStatus') || 
            data.some(row => row.approvalStatus !== undefined)) {
            createInvoiceStatusSheet(workbook, data);
        }
        
        // Create individual invoice detail sheets
        await createIndividualInvoiceDetailSheets(workbook, data);
        
        // Generate filename and save
        const finalFilename = filename || generateFilename('I2E_Selected_Fields');
        await saveWorkbook(workbook, finalFilename);
        
        logInfo(`✅ Excel file exported successfully with ${selectedFields.length} fields!`);
        
    } catch (error) {
        logError('Error exporting to Excel:', error);
        alert('An error occurred while exporting to Excel. Please try again.');
    }
}

/**
 * Export validation results to Excel (for Validator)
 * @param {Object} validationResults - Validation comparison results
 * @param {string} filename - Optional custom filename
 * @returns {Promise<void>}
 */
async function exportValidationToExcel(validationResults, filename = null) {
    try {
        logInfo('📊 Exporting validation results to Excel');
        
        // Create workbook
        const workbook = new ExcelJS.Workbook();
        
        // Create main overview sheet
        createValidationOverviewSheet(workbook, validationResults.overview);
        
        // Create detailed breakdown sheet
        createValidationDetailSheet(workbook, validationResults.details);
        
        // Generate filename and save
        const finalFilename = filename || generateFilename('I2E_Validation_Results');
        await saveWorkbook(workbook, finalFilename);
        
        logInfo('✅ Validation results exported successfully!');
        
    } catch (error) {
        logError('Error exporting validation results:', error);
        alert('An error occurred while exporting validation results. Please try again.');
    }
}

// ===== SHEET CREATION FUNCTIONS =====

/**
 * Create invoice summary sheet
 * @param {ExcelJS.Workbook} workbook - Workbook instance
 * @param {Array} data - Invoice data
 */
function createInvoiceSummarySheet(workbook, data) {
    const summarySheet = workbook.addWorksheet('Invoice Total per Cost Type');
    
    // Aggregate data by cost type
    const aggregatedData = {};
    data.forEach(row => {
        if (!row.positionDescription) return; // Skip non-line items
        
        const key = `${row.invoiceNumber}_${row.typeCost}`;
        if (!aggregatedData[key]) {
            aggregatedData[key] = {
                invoiceNumber: row.invoiceNumber,
                typeCost: row.typeCost,
                totalQuantity: 0,
                totalAmount: 0
            };
        }
        aggregatedData[key].totalQuantity += row.positionQuantity || 1;
        aggregatedData[key].totalAmount += row.positionTotal || 0;
    });
    
    // Add summary headers
    const summaryHeaders = ['Invoice Number', 'Cost Type', 'Total Quantity', 'Total Amount'];
    summarySheet.addRow(summaryHeaders);
    
    // Style summary headers
    const summaryHeaderRow = summarySheet.getRow(1);
    summaryHeaderRow.font = { bold: true };
    summaryHeaderRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F8FAFC' } };
    
    // Add summary data
    Object.values(aggregatedData).forEach(row => {
        summarySheet.addRow([
            row.invoiceNumber,
            row.typeCost,
            row.totalQuantity,
            row.totalAmount
        ]);
    });
    
    // Auto-fit summary columns
    autoFitColumns(summarySheet, 30);
    
    // Format summary amount column
    const summaryTotalColumn = summarySheet.getColumn(4);
    summaryTotalColumn.numFmt = '#,##0.00';
    
    // Add summary filters
    summarySheet.autoFilter = 'A1:D1';
}

/**
 * Create invoice status sheet (1 line per invoice with status)
 * @param {ExcelJS.Workbook} workbook - Workbook instance
 * @param {Array} data - Invoice data with status information
 */
function createInvoiceStatusSheet(workbook, data) {
    const statusSheet = workbook.addWorksheet('Invoice Status');
    
    // Headers for the status sheet
    const headers = [
        'Invoice Number', 'Project ID', 'Customer ID', 'Invoice Date', 
        'Month of Invoice', 'Total Amount', 'Currency', 'Credit Note', 
        'Status', 'Approval Date', 'Approved/Rejected By', 'Comments', 
        'Extracted Date', 'Last Modified'
    ];
    statusSheet.addRow(headers);
    
    // Style headers
    const headerRow = statusSheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F8FAFC' } };
    
    // Group data by invoice number to get one line per invoice
    const invoiceGroups = {};
    data.forEach(row => {
        const invoiceNumber = row.invoiceNumber || 'Unknown';
        if (!invoiceGroups[invoiceNumber]) {
            // Store the first occurrence of each invoice with its metadata
            invoiceGroups[invoiceNumber] = {
                invoiceNumber: invoiceNumber,
                projectId: row.projectId || '',
                customerId: row.customerId || '',
                invoiceDate: row.dateOfInvoice || '',
                monthOfInvoice: row.monthOfInvoice || '',
                totalAmount: row.extractedInvoiceTotal || 0,
                currency: row.currency || 'EUR',
                creditNote: row.creditNote ? 'Yes' : 'No',
                status: row.approvalStatus || 'Unknown',
                approvalDate: row.approvalDate || '',
                approvedBy: row.approvedBy || '',
                comments: row.comments || '',
                extractedDate: row.extractedDate || '',
                lastModified: row.lastModified || ''
            };
        }
    });
    
    // Add data rows (one per invoice)
    Object.values(invoiceGroups).forEach(invoice => {
        const dataRow = [
            invoice.invoiceNumber,
            invoice.projectId,
            invoice.customerId,
            invoice.invoiceDate,
            invoice.monthOfInvoice,
            invoice.totalAmount,
            invoice.currency,
            invoice.creditNote,
            invoice.status,
            invoice.approvalDate,
            invoice.approvedBy,
            invoice.comments,
            invoice.extractedDate,
            invoice.lastModified
        ];
        
        const excelRow = statusSheet.addRow(dataRow);
        
        // Color coding based on status
        if (invoice.status === 'approved') {
            excelRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D1FAE5' } }; // Green
        } else if (invoice.status === 'rejected') {
            excelRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FEE2E2' } }; // Red
        } else if (invoice.status === 'pending') {
            excelRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FEF3C7' } }; // Yellow
        }
        
        // Highlight credit notes with blue background
        if (invoice.creditNote === 'Yes') {
            excelRow.getCell(8).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'B0E6FF' } };
        }
    });
    
    // Auto-fit columns
    autoFitColumnsWithHeaders(statusSheet, headers);
    
    // Format amount column
    const amountColumn = statusSheet.getColumn(6);
    amountColumn.numFmt = '#,##0.00';
    
    // Add filters
    statusSheet.autoFilter = 'A1:N1';
}

/**
 * Create validation overview sheet
 * @param {ExcelJS.Workbook} workbook - Workbook instance
 * @param {Array} overviewData - Overview validation data
 */
function createValidationOverviewSheet(workbook, overviewData) {
    const sheet = workbook.addWorksheet('Validation Overview');
    
    // Add headers
    const headers = [
        'WBS Code', 'Project Name', 'Internal Cost', 'External Cost', 
        'Total Cost', 'Total Invoiced', 'Delta', 'Status'
    ];
    sheet.addRow(headers);
    
    // Style headers
    const headerRow = sheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F8FAFC' } };
    
    // Add data rows
    overviewData.forEach(row => {
        const dataRow = [
            row.wbsCode,
            row.projectName,
            row.internalCost,
            row.externalCost,
            row.totalCost,
            row.totalInvoiced,
            row.delta,
            row.delta >= 0 ? 'Under Budget' : 'Over Budget'
        ];
        
        const excelRow = sheet.addRow(dataRow);
        
        // Color coding: Green for positive delta, Red for negative
        if (row.delta >= 0) {
            excelRow.getCell(8).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D1FAE5' } };
        } else {
            excelRow.getCell(8).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FEE2E2' } };
        }
    });
    
    // Auto-fit columns
    autoFitColumns(sheet);
    
    // Format currency columns
    [3, 4, 5, 6, 7].forEach(colIndex => {
        const column = sheet.getColumn(colIndex);
        column.numFmt = '#,##0.00';
    });
    
    // Add filters
    sheet.autoFilter = 'A1:H1';
}

/**
 * Create validation detail sheet
 * @param {ExcelJS.Workbook} workbook - Workbook instance
 * @param {Array} detailData - Detail validation data
 */
function createValidationDetailSheet(workbook, detailData) {
    const sheet = workbook.addWorksheet('Validation Details');
    
    // Add headers
    const headers = [
        'WBS Code', 'Invoice Month', 'Service Period', 'Source', 
        'Description', 'Quantity', 'Amount', 'Cost Type'
    ];
    sheet.addRow(headers);
    
    // Style headers
    const headerRow = sheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F8FAFC' } };
    
    // Add data rows
    detailData.forEach(row => {
        const dataRow = [
            row.wbsCode,
            row.invoiceMonth,
            row.servicePeriod,
            row.source, // 'Internal', 'External', 'Invoice'
            row.description,
            row.quantity,
            row.amount,
            row.costType
        ];
        
        sheet.addRow(dataRow);
    });
    
    // Auto-fit columns
    autoFitColumns(sheet);
    
    // Format amount column
    const amountColumn = sheet.getColumn(7);
    amountColumn.numFmt = '#,##0.00';
    
    // Add filters
    sheet.autoFilter = 'A1:H1';
}

// ===== UTILITY FUNCTIONS =====

/**
 * Auto-fit columns based on content
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
 * Auto-fit columns with specific headers
 * @param {ExcelJS.Worksheet} worksheet - Worksheet to format
 * @param {Array} headers - Array of header names
 * @param {number} maxWidth - Maximum column width (default: 50)
 */
function autoFitColumnsWithHeaders(worksheet, headers, maxWidth = 50) {
    worksheet.columns.forEach((column, index) => {
        let maxLength = headers[index] ? headers[index].length : 0;
        
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
 * Generate filename with timestamp
 * @param {string} prefix - Filename prefix
 * @returns {string} Generated filename
 */
function generateFilename(prefix) {
    const timestamp = new Date().toISOString().slice(0, 10);
    return `${prefix}_${timestamp}.xlsx`;
}

/**
 * Save workbook to file and trigger download
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

// ===== DATA FORMATTING FUNCTIONS =====

/**
 * Format data for export based on selected fields
 * @param {Array} extractedData - Raw extracted data
 * @param {Array} selectedFields - Fields to include
 * @returns {Array} Formatted data
 */
function formatDataForExport(extractedData, selectedFields) {
    return extractedData.map(row => {
        const formattedRow = {};
        
        selectedFields.forEach(fieldKey => {
            let value = row[fieldKey];
            
            // Apply field-specific formatting
            switch (fieldKey) {
                case 'creditNote':
                    value = value ? 'Yes' : 'No';
                    break;
                case 'positionTotal':
                case 'unitPrice':
                case 'extractedInvoiceTotal':
                    value = roundToDecimals(value || 0, 2);
                    break;
                case 'positionQuantity':
                    value = value || 1;
                    break;
                default:
                    value = value || '';
            }
            
            formattedRow[fieldKey] = value;
        });
        
        // Preserve validation status
        formattedRow.hasValidationError = row.hasValidationError;
        formattedRow.creditNote = row.creditNote;
        
        return formattedRow;
    });
}

/**
 * Get all available fields for export
 * @returns {Object} Field definitions
 */
function getAllAvailableFields() {
    return {
        fileName: { name: 'File Name', description: 'Original PDF filename' },
        projectId: { name: 'Project ID', description: 'Project identifier (e.g., BE44-PRO0023366)' },
        invoiceNumber: { name: 'Invoice Number', description: 'Invoice reference number' },
        customerId: { name: 'Customer ID', description: 'Customer identifier' },
        dateOfInvoice: { name: 'Invoice Date', description: 'Date when invoice was issued' },
        monthOfInvoice: { name: 'Month of Invoice', description: 'Month when invoice was issued' },
        currency: { name: 'Currency', description: 'Invoice currency (EUR, USD, etc.)' },
        vat: { name: 'VAT', description: 'VAT percentage or identifier' },
        creditNote: { name: 'Credit Note', description: 'Whether this is a credit note (Yes/No)' },
        serviceProvisionPeriod: { name: 'Service Period', description: 'Period when service was provided' },
        position: { name: 'Position', description: 'Line item position number' },
        material: { name: 'Material', description: 'Material/service code' },
        positionDescription: { name: 'Description', description: 'Line item description' },
        positionQuantity: { name: 'Quantity', description: 'Quantity of items/hours' },
        unit: { name: 'Unit', description: 'Unit of measurement (hours, pcs, etc.)' },
        unitPrice: { name: 'Unit Price', description: 'Price per unit' },
        positionTotal: { name: 'Line Total', description: 'Total amount for this line item' },
        typeCost: { name: 'Cost Type', description: 'Internal or External cost classification' },
        extractedInvoiceTotal: { name: 'Invoice Total', description: 'Total amount extracted from invoice' },
        pageNumber: { name: 'Page Number', description: 'PDF page number where item was found' }
    };
}

/**
 * Validate field selection
 * @param {Array} selectedFields - Selected field keys
 * @param {Object} allFields - All available fields
 * @returns {Object} Validation result
 */
function validateFieldSelection(selectedFields, allFields) {
    const result = {
        valid: true,
        errors: [],
        warnings: []
    };
    
    if (!selectedFields || selectedFields.length === 0) {
        result.valid = false;
        result.errors.push('No fields selected');
        return result;
    }
    
    // Check for invalid fields
    const invalidFields = selectedFields.filter(field => !allFields[field]);
    if (invalidFields.length > 0) {
        result.valid = false;
        result.errors.push(`Invalid fields: ${invalidFields.join(', ')}`);
    }
    
    // Warnings for missing important fields
    const importantFields = ['projectId', 'invoiceNumber', 'positionTotal'];
    const missingImportant = importantFields.filter(field => !selectedFields.includes(field));
    if (missingImportant.length > 0) {
        result.warnings.push(`Consider including: ${missingImportant.join(', ')}`);
    }
    
    return result;
}

/**
 * Create individual invoice detail sheets for audit trail
 * @param {Object} workbook - ExcelJS workbook
 * @param {Array} data - Export data array
 */
async function createIndividualInvoiceDetailSheets(workbook, data) {
    try {
        // Get all unique invoices from the data
        const invoiceMap = new Map();
        data.forEach(row => {
            if (row.invoiceNumber) {
                if (!invoiceMap.has(row.invoiceNumber)) {
                    invoiceMap.set(row.invoiceNumber, {
                        invoiceNumber: row.invoiceNumber,
                        projectId: row.projectId,
                        customerId: row.customerId,
                        dateOfInvoice: row.dateOfInvoice,
                        monthOfInvoice: row.monthOfInvoice,
                        approvalStatus: row.approvalStatus,
                        approvalDate: row.approvalDate,
                        approvedBy: row.approvedBy,
                        comments: row.comments,
                        lineItems: []
                    });
                }
                invoiceMap.get(row.invoiceNumber).lineItems.push(row);
            }
        });

        // Get detailed invoice data from cache to include cost comparison
        const allInvoices = typeof getAllInvoices === 'function' ? getAllInvoices() : [];
        
        // Create a sheet for each invoice
        for (const [invoiceNumber, invoiceInfo] of invoiceMap) {
            // Find the full invoice data from cache
            const fullInvoice = allInvoices.find(inv => inv.invoiceNumber === invoiceNumber);
            
            // Create safe sheet name (Excel limits: 31 chars, no special chars)
            const safeSheetName = `Invoice_${invoiceNumber}`.substring(0, 31).replace(/[\\\/\?\*\[\]]/g, '_');
            const sheet = workbook.addWorksheet(safeSheetName);
            
            // Invoice Header Section
            sheet.addRow(['INVOICE DETAILS']);
            sheet.addRow([]); // Empty row
            
            // Basic invoice information
            sheet.addRow(['Invoice Number:', invoiceNumber]);
            sheet.addRow(['Project ID:', invoiceInfo.projectId || 'Unknown']);
            sheet.addRow(['Customer ID:', invoiceInfo.customerId || 'Unknown']);
            sheet.addRow(['Invoice Date:', invoiceInfo.dateOfInvoice || 'Unknown']);
            sheet.addRow(['Month of Invoice:', invoiceInfo.monthOfInvoice || 'Unknown']);
            sheet.addRow(['Approval Status:', invoiceInfo.approvalStatus || 'Pending']);
            
            if (invoiceInfo.approvalDate) {
                sheet.addRow(['Approval Date:', invoiceInfo.approvalDate]);
            }
            if (invoiceInfo.approvedBy) {
                sheet.addRow(['Approved/Rejected By:', invoiceInfo.approvedBy]);
            }
            if (invoiceInfo.comments) {
                sheet.addRow(['Comments:', invoiceInfo.comments]);
            }
            
            sheet.addRow([]); // Empty row
            
            // Cost Comparison Section (if full invoice data is available)
            if (fullInvoice && typeof calculateInvoiceCosts === 'function') {
                try {
                    const costs = calculateInvoiceCosts(fullInvoice, invoiceInfo.monthOfInvoice);
                    
                    sheet.addRow(['COST COMPARISON']);
                    sheet.addRow(['Internal Cost (PPM):', costs.internalCost || 0, 'EUR']);
                    sheet.addRow(['External Cost (EXT SAP):', costs.externalCost || 0, 'EUR']);
                    sheet.addRow(['Total Expected Cost:', costs.totalCost || 0, 'EUR']);
                    sheet.addRow(['Total Invoiced Amount:', costs.invoicedAmount || 0, 'EUR']);
                    sheet.addRow(['Delta (Invoice - Cost):', costs.delta || 0, 'EUR']);
                    sheet.addRow([]); // Empty row
                    
                    // Add detailed cost breakdown if available
                    if (typeof getDetailedCostBreakdown === 'function') {
                        const breakdown = getDetailedCostBreakdown(fullInvoice, invoiceInfo.monthOfInvoice);
                        
                        if (breakdown.ppmData && breakdown.ppmData.length > 0) {
                            sheet.addRow(['PPM DATA MATCHES']);
                            sheet.addRow(['Employee', 'Role', 'Rate', 'Cost', 'WBS', 'Month']);
                            breakdown.ppmData.forEach(item => {
                                sheet.addRow([
                                    item.user || '',
                                    item.role || '',
                                    item.rate || 0,
                                    item.cost || 0,
                                    item.wbs || '',
                                    item.month || ''
                                ]);
                            });
                            sheet.addRow([]); // Empty row
                        }
                        
                        if (breakdown.extSapData && breakdown.extSapData.length > 0) {
                            sheet.addRow(['EXT SAP DATA MATCHES']);
                            sheet.addRow(['Supplier', 'Value', 'Period', 'Fiscal Year', 'Document Date', 'Document Type', 'Document Number']);
                            breakdown.extSapData.forEach(item => {
                                sheet.addRow([
                                    item.supplier || '',
                                    item.value || 0,
                                    item.period || '',
                                    item.fiscalYear || '',
                                    item.documentDate || '',
                                    item.documentType || '',
                                    item.documentNumber || ''
                                ]);
                            });
                            sheet.addRow([]); // Empty row
                        }
                    }
                } catch (error) {
                    console.warn(`Could not calculate costs for invoice ${invoiceNumber}:`, error);
                }
            }
            
            // Invoice Line Items Section
            sheet.addRow(['INVOICE LINE ITEMS']);
            
            // Headers for line items
            const lineItemHeaders = [
                'Position Description', 'Position Quantity', 'Position Total', 
                'Unit Price', 'Cost Type', 'Service Period', 'Currency'
            ];
            sheet.addRow(lineItemHeaders);
            
            // Add line item data
            invoiceInfo.lineItems.forEach(item => {
                sheet.addRow([
                    item.positionDescription || '',
                    item.positionQuantity || '',
                    item.positionTotal || 0,
                    item.unitPrice || '',
                    item.typeCost || '',
                    item.serviceProvisionPeriod || '',
                    item.currency || 'EUR'
                ]);
            });
            
            // Style the sheet
            // Header styling
            const headerRow1 = sheet.getRow(1);
            headerRow1.font = { bold: true, size: 14 };
            headerRow1.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D1FAE5' } };
            
            // Find and style section headers
            sheet.eachRow((row, rowNumber) => {
                const firstCell = row.getCell(1);
                if (firstCell.value && typeof firstCell.value === 'string' && 
                    (firstCell.value.includes('COST COMPARISON') || 
                     firstCell.value.includes('INVOICE LINE ITEMS') ||
                     firstCell.value.includes('PPM DATA MATCHES') ||
                     firstCell.value.includes('EXT SAP DATA MATCHES'))) {
                    row.font = { bold: true };
                    row.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F3F4F6' } };
                }
            });
            
            // Auto-fit columns
            sheet.columns.forEach(column => {
                let maxLength = 10;
                column.eachCell({ includeEmpty: true }, (cell) => {
                    const cellValue = cell.value ? cell.value.toString() : '';
                    maxLength = Math.max(maxLength, cellValue.length);
                });
                column.width = Math.min(maxLength + 2, 50); // Max width of 50
            });
        }
        
        console.log(`✅ Created ${invoiceMap.size} individual invoice detail sheets`);
        
    } catch (error) {
        console.error('❌ Error creating individual invoice detail sheets:', error);
    }
}

// ===== EXPORT FOR MODULE USAGE =====

// If using as a module, export functions
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        exportToExcel,
        exportToExcelWithFields,
        exportValidationToExcel,
        formatDataForExport,
        getAllAvailableFields,
        validateFieldSelection,
        generateFilename,
        saveWorkbook
    };
}

// Global availability for browser usage
if (typeof window !== 'undefined') {
    logInfo('I2E Excel Exporter loaded successfully');
}