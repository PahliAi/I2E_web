# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is the I2E Invoice Processor - a zero-cost, frontend-only PDF invoice data extraction tool that exports to Excel. The application consists of two main HTML files with shared JavaScript modules for processing invoices and validating data.

## Architecture

### Main Applications
- **I2E_Invoice_Processor.html** - Primary invoice processing application with PDF upload, data extraction, and Excel export
- **I2E_Invoice_Validator.html** - Validation interface for reviewing and approving extracted invoice data

### Shared Module System
The application uses a modular architecture with shared JavaScript libraries:

- **shared/i2e-common.js** - Common utilities (file handling, currency formatting, WBS standardization)
- **shared/pdf-extractor.js** - PDF processing using PDF.js library for text extraction and invoice data parsing
- **shared/excel-exporter.js** - Excel file generation using ExcelJS library with configurable field selection
- **shared/i2e-cache.js** - Client-side caching system for invoice validation workflow
- **shared/i2e-spy.js** - Development utilities and debugging tools

### External Dependencies
- **PDF.js (3.11.174)** - Client-side PDF processing
- **ExcelJS (4.4.0)** - Excel file generation
- **i2e-styles.css** - Shared CSS framework with responsive design

## Key Features

### Invoice Processing Pipeline
1. **PDF Upload** - Drag-and-drop or file browser upload of multiple PDF files
2. **Data Extraction** - Text extraction from PDFs with pattern matching for invoice fields
3. **Data Validation** - Real-time validation of extracted totals vs. calculated line item sums
4. **Excel Export** - Configurable field selection and export to Excel with multiple sheets

### Data Structure
The application extracts and processes these key fields:
- Invoice metadata (number, date, customer, project ID)
- Line items with service periods, quantities, unit prices, totals
- Credit note detection and amount correction
- VAT information and currency handling

### Validation System
- **Duplicate Detection** - Prevents processing the same invoice twice
- **Total Validation** - Compares extracted invoice totals with calculated line item sums
- **Credit Note Handling** - Automatic detection and amount sign correction
- **Cache-based Workflow** - Pending/approved/rejected status tracking

## Development

### No Build Process
This is a pure frontend application with no build tools, package managers, or compilation steps required. All files are served directly from the filesystem.

### File Structure
```
I2E_web/
├── I2E_Invoice_Processor.html    # Main processing interface
├── I2E_Invoice_Validator.html    # Validation interface
├── assets/
│   └── i2e-styles.css           # Shared CSS framework
└── shared/
    ├── i2e-common.js            # Common utilities
    ├── pdf-extractor.js         # PDF processing logic
    ├── excel-exporter.js        # Excel export functionality
    ├── i2e-cache.js             # Client-side caching
    └── i2e-spy.js               # Development utilities
```

### Testing
- Manual testing through browser interface
- No automated test suite
- Validation testing through the I2E_Invoice_Validator.html interface

### Browser Compatibility
- Requires modern browser with ES6+ support
- Uses Web APIs: FileReader, localStorage, Blob, URL.createObjectURL
- External CDN dependencies must be accessible

## Common Operations

### Adding New Invoice Fields
1. Update field definitions in `I2E_Invoice_Processor.html` (allFields object)
2. Add extraction logic in `shared/pdf-extractor.js` (extractField function)
3. Update Excel export headers in `shared/excel-exporter.js`

### Modifying PDF Parsing
- Text extraction logic is in `shared/pdf-extractor.js`
- Pattern matching uses regular expressions for different invoice formats
- Service period extraction is page-specific to handle multi-page invoices

### Customizing Export Format
- Field selection UI is in the field selection modal
- Export configuration can be saved to localStorage
- Multiple export presets are available (default, detailed, summary)

## Data Flow

1. **File Upload** → PDF files uploaded via drag-and-drop or file input
2. **PDF Processing** → Text extraction using PDF.js, pattern matching for data fields
3. **Data Structuring** → Hierarchical organization by invoice → service period → line items
4. **Validation** → Real-time validation of totals, duplicate detection, credit note handling
5. **Caching** → Storage in localStorage for validation workflow
6. **Export** → Configurable Excel export with multiple sheets and formatting

## Security Notes

- All processing is client-side only
- No server communication or data transmission
- Files are processed in browser memory only
- localStorage used for caching and configuration
- No authentication or access control (frontend-only tool)