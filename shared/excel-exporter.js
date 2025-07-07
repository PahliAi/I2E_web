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