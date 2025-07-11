/**
 * I2E PDF Extractor
 * PDF processing and data extraction logic for I2E Invoice Processor
 * 
 * @version 1.0
 * @author I2E Development Team
 * @requires PDF.js library
 * @requires i2e-common.js for utilities
 */

// ===== MAIN PDF PROCESSING FUNCTIONS =====

/**
 * Extract data from PDF file
 * @param {File} file - PDF file to process
 * @returns {Promise<Array>} Array of extracted invoice data
 */
async function extractDataFromPDF(file) {
    try {
        // Load PDF
        const arrayBuffer = await file.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        
        // Extract text from all pages
        let fullText = '';
        const pageTexts = [];
        
        for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
            const page = await pdf.getPage(pageNum);
            const textContent = await page.getTextContent();
            
            // Group text items by approximate Y position to preserve lines
            const lines = {};
            textContent.items.forEach(item => {
                const y = Math.round(item.transform[5] / 5) * 5; // Group by 5-pixel chunks
                if (!lines[y]) lines[y] = [];
                lines[y].push(item);
            });
            
            // Sort lines by Y position (top to bottom) and concatenate within lines
            const sortedY = Object.keys(lines).sort((a, b) => b - a); // Descending Y
            const pageText = sortedY.map(y => {
                // Sort items within line by X position (left to right)
                const lineItems = lines[y].sort((a, b) => a.transform[4] - b.transform[4]);
                return lineItems.map(item => item.str).join(' ').trim();
            }).filter(line => line.length > 0).join('\n');
            
            pageTexts.push(pageText);
            fullText += pageText + '\n';
        }
        
        // Extract invoice data
        const invoiceData = extractInvoiceData(fullText, pageTexts, file.name);
        
        return invoiceData;
        
    } catch (error) {
        console.error('Error processing PDF:', error);
        throw new Error(`Failed to process PDF: ${error.message}`);
    }
}

function extractInvoiceData(fullText, pageTexts, fileName) {
    // Extract invoice-level data
    const baseInvoiceInfo = {
        fileName: fileName,
        projectId: extractField(fullText, 'projectId'),
        invoiceNumber: extractField(fullText, 'invoiceNumber'),
        customerId: extractField(fullText, 'customerId'),
        dateOfInvoice: extractField(fullText, 'dateOfInvoice'),
        monthOfInvoice: extractMonthFromDate(extractField(fullText, 'dateOfInvoice')),
        currency: extractField(fullText, 'currency'),
        vat: extractField(fullText, 'vat'),
        creditNote: detectCreditNote(fullText)
    };
    
    // Extract line items from each page with page-specific service periods
    const lineItems = [];
    let extractedInvoiceTotal = null;
    const allPageTotals = [];
    
    pageTexts.forEach((pageText, pageIndex) => {
        // Extract service provision period for this specific page
        const pageServicePeriod = extractServicePeriodFromPage(pageText);
        
        // Extract potential totals from ALL pages, not just the first
        const pageTotal = extractInvoiceTotal(pageText);
        if (pageTotal !== null) {
            allPageTotals.push({
                amount: pageTotal,
                pageNumber: pageIndex + 1,
                pageText: pageText
            });
            console.log(`💰 Found potential total on page ${pageIndex + 1}: ${pageTotal}`);
        }
        
        const pageLineItems = extractLineItems(pageText, pageIndex + 1);
        pageLineItems.forEach(item => {
            lineItems.push({
                ...baseInvoiceInfo,
                serviceProvisionPeriod: pageServicePeriod,
                ...item,
                pageNumber: pageIndex + 1,
                extractedInvoiceTotal: null // Will be set after choosing best total
            });
        });
    });
    
    // Choose the best total from all pages
    if (allPageTotals.length > 0) {
        console.log(`🔍 Analyzing ${allPageTotals.length} potential totals from all pages:`, allPageTotals);
        
        // Check each total to see if it comes from a "Total" line vs "Subtotal" line
        const totalCandidates = allPageTotals.map(candidate => {
            const hasMainTotal = /\bTotal\b/i.test(candidate.pageText) && !/Subtotal/i.test(candidate.pageText.match(/.*Total.*$/im)?.[0] || '');
            const hasSubtotalOnly = /Subtotal/i.test(candidate.pageText) && !/\bTotal\b(?!.*Subtotal)/i.test(candidate.pageText);
            
            return {
                ...candidate,
                isMainTotal: hasMainTotal,
                isSubtotalOnly: hasSubtotalOnly,
                priority: hasMainTotal ? 1 : (hasSubtotalOnly ? 3 : 2) // 1=Total, 2=Mixed, 3=Subtotal only
            };
        });
        
        // Sort by priority (lower number = higher priority), then by page number (later page wins)
        totalCandidates.sort((a, b) => {
            if (a.priority !== b.priority) return a.priority - b.priority;
            return b.pageNumber - a.pageNumber; // Later page wins for same priority
        });
        
        extractedInvoiceTotal = totalCandidates[0].amount;
        console.log(`✅ Selected best total: ${extractedInvoiceTotal} from page ${totalCandidates[0].pageNumber} (priority: ${totalCandidates[0].priority}, isMainTotal: ${totalCandidates[0].isMainTotal})`);
    }
    
    // Update all line items with the final selected total
    lineItems.forEach(item => {
        item.extractedInvoiceTotal = extractedInvoiceTotal;
    });
    
    return lineItems.length > 0 ? lineItems : [{...baseInvoiceInfo, extractedInvoiceTotal}];
}

function extractServicePeriodFromPage(pageText) {
    // Look for month names in page headers (JAN 2024, FEB 2025, MAR 2022, etc.)
    // This pattern matches: "MONTH YYYY" where MONTH is 3-letter abbreviation
    
    const monthPattern = /\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+(\d{4})\b/i;
    const match = pageText.match(monthPattern);
    
    if (match) {
        const monthAbbr = match[1].toUpperCase();
        const year = match[2];
        
        // Convert 3-letter month abbreviation to full month name
        const monthMap = {
            'JAN': 'January', 'FEB': 'February', 'MAR': 'March',
            'APR': 'April', 'MAY': 'May', 'JUN': 'June',
            'JUL': 'July', 'AUG': 'August', 'SEP': 'September',
            'OCT': 'October', 'NOV': 'November', 'DEC': 'December'
        };
        
        const fullMonthName = monthMap[monthAbbr];
        const servicePeriod = `${fullMonthName} ${year}`;
        
        console.log(`📅 Extracted service period from page: "${servicePeriod}" (from pattern: "${match[0]}")`);
        return servicePeriod;
    }
    
    // Fallback: try to extract from more flexible patterns
    const fallbackPatterns = [
        // Handle variations like "Service Provision Period: 02/2024"
        /Service.*?Period.*?:?\s*(\d{1,2})[./](\d{4})/i,
        // Handle numeric month patterns
        /\b(\d{1,2})[./](\d{4})\b/
    ];
    
    for (const pattern of fallbackPatterns) {
        const fallbackMatch = pageText.match(pattern);
        if (fallbackMatch) {
            const month = parseInt(fallbackMatch[1], 10);
            const year = fallbackMatch[2];
            
            // Convert numeric month to name
            const monthNames = [
                '', 'January', 'February', 'March', 'April', 'May', 'June',
                'July', 'August', 'September', 'October', 'November', 'December'
            ];
            
            if (month >= 1 && month <= 12) {
                const servicePeriod = `${monthNames[month]} ${year}`;
                console.log(`📅 Extracted service period from fallback: "${servicePeriod}" (from pattern: "${fallbackMatch[0]}")`);
                return servicePeriod;
            }
        }
    }
    
    console.log('⚠️ Could not extract service period from page.');
    return 'Unknown Period';
}

function normalizeServicePeriod(periodStr) {
    const monthMap = {
        'JAN': 'January', 'FEB': 'February', 'MAR': 'March', 'APR': 'April',
        'MAY': 'May', 'JUN': 'June', 'JUL': 'July', 'AUG': 'August',
        'SEP': 'September', 'OCT': 'October', 'NOV': 'November', 'DEC': 'December'
    };
    
    // Extract month abbreviation
    for (const [abbr, fullName] of Object.entries(monthMap)) {
        if (periodStr.toUpperCase().includes(abbr)) {
            return fullName;
        }
    }
    
    return periodStr;
}

function extractField(text, fieldType) {
    const patterns = {
        projectId: [
            /([A-Z]{2}\d{2}-PRO\d{7})/,
            /(PRO\d{7})/,
            /([A-Z]{2}-PRO\d{7})/
        ],
        invoiceNumber: [
            /Invoice\s+No\.?\s*:?\s*(\d+)/i,
            /Credit\s+Note\s+No\.?\s*:?\s*(\d+)/i,
            /Invoice\s+Number\s*:?\s*(\d+)/i
        ],
        customerId: [
            /Customer\s+ID\s*:?\s*(\d+)/i,
            /Client\s+ID\s*:?\s*(\d+)/i
        ],
        dateOfInvoice: [
            /Date\s*:?\s*(\d{1,2}[./]\d{1,2}[./]\d{4})/i,
            /Invoice\s+Date\s*:?\s*(\d{1,2}[./]\d{1,2}[./]\d{4})/i
        ],
        currency: [
            /Currency\s*:?\s*([A-Z]{3})/i,
            /(EUR|USD|GBP)/
        ],
        vat: [
            /VAT\s*ID\s*:?\s*([A-Z0-9]+)/i,
            /BTW\s*:?\s*([A-Z0-9]+)/i
        ],
        serviceProvisionPeriod: [
            /Service\s+Provision\s+Period\s*:?\s*(\d{1,2}[./]\d{4}(?:\s*-\s*\d{1,2}[./]\d{4})?)/i,
            /Period\s*:?\s*(\d{1,2}[./]\d{4}(?:\s*-\s*\d{1,2}[./]\d{4})?)/i,
            /([A-Z]{3,9}\s+\d{4})/i,
            // More flexible patterns for various formats
            /Service.*?Period.*?:?\s*([0-9][0-9]?[./]\d{4}(?:\s*-\s*[0-9][0-9]?[./]\d{4})?)/i,
            /Provision.*?Period.*?:?\s*([0-9][0-9]?[./]\d{4}(?:\s*-\s*[0-9][0-9]?[./]\d{4})?)/i
        ]
    };
    
    const fieldPatterns = patterns[fieldType] || [];
    
    for (const pattern of fieldPatterns) {
        const match = text.match(pattern);
        if (match) {
            return match[1];
        }
    }
    
    return null;
}

function extractMonthFromDate(dateStr) {
    if (!dateStr) return null;
    
    const months = {
        '01': 'January', '02': 'February', '03': 'March', '04': 'April',
        '05': 'May', '06': 'June', '07': 'July', '08': 'August',
        '09': 'September', '10': 'October', '11': 'November', '12': 'December'
    };
    
    // Handle different date formats
    const parts = dateStr.split(/[./-]/);
    if (parts.length >= 3) {
        // Assuming DD.MM.YYYY format
        const monthNum = parts[1].padStart(2, '0');
        return months[monthNum] || null;
    }
    
    return null;
}

function detectCreditNote(text) {
    const creditIndicators = [
        'credit note', 'creditnote', 'credit memo', 'refund',
        'return', 'adjustment', 'reversal'
    ];
    
    const textLower = text.toLowerCase();
    return creditIndicators.some(indicator => textLower.includes(indicator));
}

function extractInvoiceTotal(pageText) {
    // Extract the invoice total from the PDF text
    // More flexible patterns that handle line breaks and spacing
    const totalPatterns = [
        // Direct patterns (amount on same line)
        /Total\s+([0-9.,]+[-]?)/gi,
        /Total\s*:?\s*([0-9.,]+[-]?)/gi,
        /Subtotal\s+([0-9.,]+[-]?)/gi,
        /Grand\s+Total\s+([0-9.,]+[-]?)/gi,
        // Subtotal variations
        /Subtotal\s*\(Net\)\s+([0-9.,]+[-]?)/gi,
        /\w+\s+\d+\s+Subtotal\s*\(Net\)\s+([0-9.,]+[-]?)/gi,
        // Flexible patterns (find Total, then look for nearby amounts)
        /Total[\s\S]{0,50}?([0-9.,]{4,}[-]?)/gi,
        /Subtotal[\s\S]{0,50}?([0-9.,]{4,}[-]?)/gi
    ];
    
    // Also try finding amounts near "Total" word
    const lines = pageText.split('\n');
    let extractedTotal = null;
    
    // Collect all potential totals, then choose the best one
    const potentialTotals = [];
    
    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        if (/\b(Total|Subtotal.*\(Net\)|Subtotal)\b/i.test(line) && !/Position.*Total/i.test(line)) {
            console.log(`🔍 Found total/subtotal line: "${line.trim()}"`);
            
            // Skip VAT lines - we want actual totals, not VAT amounts
            if (/VAT\s+[0-9.,]+%/i.test(line)) {
                console.log(`⚠️ Skipping VAT line: "${line.trim()}"`);
                continue;
            }
            
            const totalWordMatch = line.match(/\b(Total|Subtotal.*?\(Net\)?)\b/i);
            const isMainTotal = /\bTotal\b/i.test(line) && !/Subtotal/i.test(line);
            
            // First try to extract amount from the SAME line (most common case)
            const sameLineMatch = line.match(/([0-9]{1,3}(?:[.,][0-9]{3})*[.,][0-9]{2}[-]?)/g);
            if (sameLineMatch && totalWordMatch) {
                const totalWordIndex = line.indexOf(totalWordMatch[0]);
                const totalWordEnd = totalWordIndex + totalWordMatch[0].length;
                
                // Look for amounts that appear within 50 characters AFTER the total word
                let bestAmount = null;
                
                for (const amountStr of sameLineMatch) {
                    const amountIndex = line.indexOf(amountStr, totalWordEnd);
                    if (amountIndex !== -1 && amountIndex - totalWordEnd < 50) {
                        const amount = parseAmount(amountStr);
                        if (amount !== null && Math.abs(amount) > 1) { // Skip very small amounts like 0.00
                            // Prefer larger amounts (actual totals vs VAT amounts)
                            if (Math.abs(amount) > Math.abs(bestAmount || 0)) {
                                bestAmount = amount;
                            }
                        }
                    }
                }
                
                if (bestAmount !== null) {
                    potentialTotals.push({
                        amount: bestAmount,
                        type: isMainTotal ? 'Total' : 'Subtotal',
                        line: line.trim(),
                        priority: isMainTotal ? 1 : 2 // Main Total has higher priority
                    });
                    console.log(`💰 Found potential ${isMainTotal ? 'Total' : 'Subtotal'}: ${bestAmount} from "${line.trim()}"`);
                }
            }
            
            // If not found on same line, check next few lines (but skip VAT lines)
            if (!sameLineMatch || !potentialTotals.some(t => t.line === line.trim())) {
                for (let j = i + 1; j < Math.min(i + 3, lines.length); j++) {
                    const checkLine = lines[j];
                    
                    // Skip VAT lines when looking for totals
                    if (/VAT\s+[0-9.,]+%/i.test(checkLine)) {
                        console.log(`⚠️ Skipping VAT line in next-line search: "${checkLine.trim()}"`);
                        continue;
                    }
                    
                    const amountMatch = checkLine.match(/([0-9]{1,3}(?:[.,][0-9]{3})*[.,][0-9]{2}[-]?)/g);
                    if (amountMatch) {
                        // Take the largest amount from this line
                        let largestAmount = null;
                        for (const amountStr of amountMatch) {
                            const amount = parseAmount(amountStr);
                            if (amount !== null && Math.abs(amount) > Math.abs(largestAmount || 0)) {
                                largestAmount = amount;
                            }
                        }
                        if (largestAmount !== null && Math.abs(largestAmount) > 1) {
                            potentialTotals.push({
                                amount: largestAmount,
                                type: isMainTotal ? 'Total' : 'Subtotal',
                                line: `${line.trim()} -> ${checkLine.trim()}`,
                                priority: isMainTotal ? 1 : 2
                            });
                            console.log(`💰 Found potential ${isMainTotal ? 'Total' : 'Subtotal'}: ${largestAmount} from next line: "${checkLine.trim()}"`);
                            break;
                        }
                    }
                }
            }
        }
    }
    
    // Choose the best total: prefer "Total" over "Subtotal", then by position (later in document)
    if (potentialTotals.length > 0) {
        console.log(`🔍 All potential totals found:`, potentialTotals);
        
        // Sort by priority (1=Total, 2=Subtotal), then by index (later in document wins)
        potentialTotals.sort((a, b) => {
            if (a.priority !== b.priority) return a.priority - b.priority;
            return 0; // If same priority, keep original order (later wins)
        });
        
        extractedTotal = potentialTotals[0].amount;
        console.log(`✅ Selected best total: ${extractedTotal} from ${potentialTotals[0].type} line: "${potentialTotals[0].line}"`);
    }
    
    // Fallback to regex patterns
    if (extractedTotal === null) {
        for (const pattern of totalPatterns) {
            const matches = pageText.match(pattern);
            if (matches) {
                const lastMatch = matches[matches.length - 1];
                const amountMatch = lastMatch.match(/([0-9.,]+[-]?)/);
                if (amountMatch) {
                    const amount = parseAmount(amountMatch[1]);
                    if (amount !== null && amount !== 0) {
                        extractedTotal = amount;
                        console.log(`💰 Found invoice total: ${extractedTotal} from pattern: ${pattern}`);
                        break;
                    }
                }
            }
        }
    }
    
    return extractedTotal;
}

function extractLineItems(pageText, pageNumber) {
    const lineItems = [];
    
    console.log(`🔍 Debugging page ${pageNumber} text extraction:`);
    console.log('Raw page text length:', pageText.length);
    
    // Step 1: Filter text to section BEFORE Subtotal/Total boundaries
    const lines = pageText.split('\n');
    const lineItemSection = [];
    
    for (const line of lines) {
        // Stop at subtotal/total boundaries, but not "Position Total" headers
        if ((/\b(Subtotal|Total)\b/i.test(line) && !/Position.*Total/i.test(line))) {
            console.log(`🛑 Stopping at boundary: "${line.trim()}"`);
            break;
        }
        lineItemSection.push(line);
    }
    
    console.log(`📄 Page ${pageNumber} has ${lines.length} total lines, processing ${lineItemSection.length} lines before totals`);
    
    // Step 2: Find candidate lines using VAT pattern anchor
    const candidateLines = lineItemSection.filter(line => {
        const trimmed = line.trim();
        // Look for VAT percentage pattern: number%({code})
        return trimmed.length > 20 && /\d+[.,]?\d*%\([A-Z0-9]+\)/.test(trimmed);
    });
    
    console.log(`🎯 Found ${candidateLines.length} candidate lines with VAT patterns:`, candidateLines);
    
    // Step 3: Strategy 1 - Structured VAT-anchored parsing
    const structuredPattern = /(\d{4})\s+(\d{6})\s+(.+?)\s+(\d+[.,]?\d*)\s+([A-Z]+)\s+(\d+[.,]?\d*%\([A-Z0-9]+\))\s+([\d.,]+)\s+([\d.,]+-?)/g;
    
    for (const line of candidateLines) {
        const match = structuredPattern.exec(line);
        if (match) {
            const [fullMatch, position, material, description, quantity, unit, vat, unitPrice, total] = match;
            
            const item = {
                position: position,
                material: material,
                positionDescription: description.trim(),
                positionQuantity: parseFloat(quantity.replace(',', '.')),
                unit: unit,
                vat: vat,
                unitPrice: parseAmount(unitPrice),
                positionTotal: parseAmount(total),
                typeCost: classifyCostType(description)
            };
            
            if (item.positionTotal !== null) {
                lineItems.push(item);
                console.log(`✅ Strategy 1 (VAT-anchored) extracted: ${item.position} ${item.positionDescription} = ${item.positionTotal}`);
            }
        }
        // Reset regex for next line
        structuredPattern.lastIndex = 0;
    }
    
    // Step 4: Strategy 2 - Flexible parsing for lines with VAT patterns
    if (lineItems.length === 0) {
        console.log('📋 Strategy 1 failed, trying flexible VAT-based parsing...');
        
        for (const line of candidateLines) {
            console.log(`🔍 Analyzing VAT line: "${line.trim()}"`);
            
            const parts = line.trim().split(/\s+/);
            
            // Find position (4 digits at start)
            const posIndex = parts.findIndex(p => /^\d{4}$/.test(p));
            if (posIndex >= 0) {
                // Find VAT pattern index
                const vatIndex = parts.findIndex(p => /\d+[.,]?\d*%\([A-Z0-9]+\)/.test(p));
                // Find final amount (after VAT pattern)
                const amountIndex = parts.findIndex((p, i) => i > vatIndex && /^[\d.,]+-?$/.test(p));
                
                if (vatIndex > posIndex && amountIndex > vatIndex) {
                    // Extract description between material and quantity
                    const material = parts[posIndex + 1] || '';
                    let description = '';
                    let quantity = 1;
                    
                    // Find where description ends and quantity begins
                    for (let i = posIndex + 2; i < vatIndex - 2; i++) {
                        if (/^\d+[.,]?\d*$/.test(parts[i]) && !description.includes(parts[i])) {
                            quantity = parseFloat(parts[i].replace(',', '.'));
                            break;
                        } else {
                            description += parts[i] + ' ';
                        }
                    }
                    
                    const item = {
                        position: parts[posIndex],
                        material: material,
                        positionDescription: description.trim() || 'Unknown Service',
                        positionQuantity: quantity,
                        unit: parts[vatIndex - 1] || 'PU',
                        vat: parts[vatIndex],
                        unitPrice: parseAmount(parts[vatIndex + 1]) || 0,
                        positionTotal: parseAmount(parts[amountIndex]),
                        typeCost: classifyCostType(description)
                    };
                    
                    if (item.positionTotal !== null) {
                        lineItems.push(item);
                        console.log(`✅ Strategy 2 (flexible VAT) extracted: ${item.position} ${item.positionDescription} = ${item.positionTotal}`);
                    }
                }
            }
        }
    }
    
    // Step 5: Strategy 3 - Simple pattern extraction for any line with % and amounts
    if (lineItems.length === 0) {
        console.log('📋 Strategy 2 failed, trying simple pattern extraction...');
        
        for (const line of candidateLines) {
            // Look for: position number ... percentage pattern ... final amount
            const simpleMatch = line.match(/(\d{4})\s+.*?(\d+[.,]?\d*%\([A-Z0-9]+\)).*?([\d.,]+-?)$/);
            if (simpleMatch) {
                const position = simpleMatch[1];
                const vat = simpleMatch[2];
                const total = parseAmount(simpleMatch[3]);
                
                if (total !== null) {
                    // Extract material code if present
                    const materialMatch = line.match(/\d{4}\s+(\d{6})/);
                    const material = materialMatch ? materialMatch[1] : '';
                    
                    // Extract description (everything between material and numbers)
                    let description = line.replace(/^\d{4}\s+(\d{6}\s+)?/, '');
                    description = description.replace(/\s+\d+[.,]?\d*.*$/, '').trim();
                    
                    const item = {
                        position: position,
                        material: material,
                        positionDescription: description || 'Service Item',
                        positionQuantity: 1,
                        unit: 'PU',
                        vat: vat,
                        unitPrice: total,
                        positionTotal: total,
                        typeCost: classifyCostType(description)
                    };
                    
                    lineItems.push(item);
                    console.log(`✅ Strategy 3 (simple pattern) extracted: ${item.position} ${item.positionDescription} = ${item.positionTotal}`);
                }
            }
        }
    }
    
    console.log(`📊 Page ${pageNumber}: Final result = ${lineItems.length} line items`);
    return lineItems;
}

function extractDescription(parts, startIndex) {
    // Extract description between position/material and numbers
    let desc = '';
    for (let i = startIndex + 2; i < parts.length; i++) {
        if (/^\d/.test(parts[i]) && parts[i].length > 2) break;
        desc += parts[i] + ' ';
    }
    return desc.trim();
}

function extractQuantity(parts, startIndex, endIndex) {
    // Look for quantity between description and amount
    for (let i = startIndex + 2; i < endIndex; i++) {
        const num = parseInt(parts[i], 10);
        if (!isNaN(num) && num > 0 && num < 1000) {
            return num;
        }
    }
    return 1;
}

function parseAmount(amountStr) {
    if (!amountStr) return null;
    
    // Handle European format with trailing minus
    const isNegative = amountStr.endsWith('-');
    let cleanAmount = amountStr.replace('-', '');
    
    // Convert European format (123.456,78) to standard (123456.78)
    if (cleanAmount.includes(',') && cleanAmount.lastIndexOf(',') > cleanAmount.lastIndexOf('.')) {
        cleanAmount = cleanAmount.replace(/\./g, '').replace(',', '.');
    } else {
        cleanAmount = cleanAmount.replace(/,/g, '');
    }
    
    const amount = parseFloat(cleanAmount);
    return isNegative ? -amount : amount;
}

function classifyCostType(description) {
    const descLower = description.toLowerCase();
    
    // Internal: contains "Application" or "Infrastructure"
    if (descLower.includes('application') || descLower.includes('infrastructure')) {
        return 'Internal';
    }
    
    // External: contains "External" or "Other"
    if (descLower.includes('external') || descLower.includes('other')) {
        return 'External';
    }
    
    // Default classification
    if (descLower.includes('consultant') || descLower.includes('service') || descLower.includes('support')) {
        return 'External';
    }
    
    return 'Internal';
}

function createLineItem(match) {
    try {
        const [fullMatch, position, material, description, quantity, unit, vat, unitPrice, total] = match;
        
        return {
            position: position,
            material: material,
            positionDescription: description.trim(),
            positionQuantity: parseInt(quantity, 10),
            unit: unit,
            vat: vat,
            unitPrice: parseAmount(unitPrice),
            positionTotal: parseAmount(total),
            typeCost: classifyCostType(description)
        };
    } catch (error) {
        console.error('Error creating line item:', error);
        return null;
    }
}

function makeCSSClass(str) {
    // Convert string to CSS-safe class name by replacing spaces and special chars
    return str.replace(/[^a-zA-Z0-9_-]/g, '_');
}

// ===== EXPORT FOR MODULE USAGE =====

// If using as a module, export functions
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        extractDataFromPDF,
        extractInvoiceData,
        extractServicePeriodFromPage,
        extractField,
        extractMonthFromDate,
        detectCreditNote,
        extractInvoiceTotal,
        extractLineItems
    };
}

// Global availability for browser usage
if (typeof window !== 'undefined') {
    window.extractDataFromPDF = extractDataFromPDF;
    window.extractInvoiceData = extractInvoiceData;
    window.extractServicePeriodFromPage = extractServicePeriodFromPage;
    window.extractField = extractField;
    window.extractMonthFromDate = extractMonthFromDate;
    window.detectCreditNote = detectCreditNote;
    window.extractInvoiceTotal = extractInvoiceTotal;
    window.extractLineItems = extractLineItems;
    
    console.log('I2E PDF Extractor loaded successfully');
}