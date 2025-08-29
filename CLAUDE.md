# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

**I2E (Invoice to Excel) Cost Analysis & Invoice Management System** is a zero-cost, browser-based system for cost data analysis and invoice processing. The system has been redesigned with a **cost-focused architecture** that includes:

1. **Main Dashboard** (`index.html`) - Central entry point with data overview and navigation
2. **Cost Data Analysis** - Enhanced cost views with regular and project-based pivot analysis
3. **Invoice Extraction & Approval** - Streamlined invoice processing workflow

## Architecture

### Core Technologies
- **Frontend-only architecture** - no server dependencies
- **PDF.js** - client-side PDF text extraction
- **ExcelJS** - client-side Excel file generation
- **Vanilla JavaScript** - no frameworks, modular shared utilities
- **Local Storage** - configuration and cache persistence
- **CSS Grid/Flexbox** - responsive layouts

### File Structure
```
index.html                     # Main dashboard - central entry point
I2E_Invoice_Processor.html     # PDF extraction application (self-contained)
I2E_Invoice_Validator.html     # Cost validation and analysis application
shared/
├── i2e-common.js             # Shared utilities (file handling, currency, validation)
├── i2e-cache.js              # Local storage management and data caching
├── i2e-spy.js                # Performance monitoring and analytics
├── pdf-extractor.js          # PDF processing and data extraction logic
└── excel-exporter.js         # Excel generation and formatting
assets/
└── i2e-styles.css            # Shared CSS styles with dashboard components
design/                       # Technical documentation and design specs
archive/                      # Legacy files and test data
```

## Key Features

### Main Dashboard (`index.html`)
- **Cache data overview** - displays file statistics and invoice status counts
- **Smart navigation** - buttons enabled only when required data is available
- **Unified upload** - drag-and-drop areas for both cost data (XLSX) and invoices (PDF)
- **Help integration** - prominent help button with README modal
- **Responsive design** - three-section layout (overview, upload, navigation)

### Cost Data Analysis
- **Dual view modes** - Regular views and Project-based views
- **Regular views** - traditional field-by-field analysis (hours per employee, costs per supplier)
- **Project views** - expandable/collapsible project-centric analysis with pivot tables
- **Real-time filtering** - project name search and multi-select capabilities
- **Enhanced visualizations** - monthly costs, cost by type, supplier analysis

### Invoice Extraction & Approval
- **Streamlined workflow** - direct access to pending invoices from dashboard
- **VAT-pattern-anchored extraction** - uses VAT percentages as primary line detection
- **Multi-page intelligence** - analyzes all pages, prefers "Total" over "Subtotal"
- **Service period extraction** - extracts month names from headers (JAN 2024, FEB 2025)
- **Credit note support** - negative amounts with baby-blue highlighting
- **Cost validation** - compares against PPM and EXT SAP data

## Common Development Tasks

### Building and Testing
```bash
# No build process required - direct HTML file deployment
# Test by opening files in browser

# Main entry point - always start here
open index.html

# Direct navigation (for development/testing)
open I2E_Invoice_Validator.html?view=costviews    # Cost analysis view
open I2E_Invoice_Validator.html?view=pending      # Invoice approval view
```

### Data Extraction Patterns
Key regex patterns are located in `shared/pdf-extractor.js`:
- Project ID: `AA44-PRO0012345` or `PRO0012345` format
- VAT detection: Pattern-based line item identification
- Service periods: Month name extraction from headers
- Currency handling: Multiple format support (€1,234.56, 1.234,56)

### Storage Management
Use `shared/i2e-cache.js` for data persistence:
- `cacheInvoiceData()` - store processed invoices
- `getCachedInvoiceData()` - retrieve cached data
- `clearCacheByProject()` - project-specific cleanup
- Local storage keys use `i2e_` prefix for consistency

### Excel Export Configuration
Export field configuration in `shared/excel-exporter.js`:
- Default fields: filename, customer ID, invoice number, project ID, service period
- Detailed report: all available fields including position codes
- Custom selection: user-defined field ordering
- Uses ExcelJS for client-side generation

## Debugging and Troubleshooting

### Common Issues
- **PDF extraction failures**: Check for text-based PDFs (not scanned images)
- **Memory issues**: Large files may exceed browser memory limits
- **Data validation errors**: Red highlighting indicates extraction vs calculation mismatches
- **Cache corruption**: Use `clearStorageByPrefix('i2e_')` from i2e-common.js

### Performance Monitoring
The `i2e-spy.js` module provides performance tracking:
- Processing times for PDF extraction
- Memory usage monitoring
- Error rate tracking
- Cache hit/miss ratios

### Browser Compatibility
- Requires modern browser with JavaScript enabled
- PDF.js requires Chrome 80+, Firefox 75+, Safari 13+, Edge 80+
- Minimum 2GB RAM for large invoice processing

## Data Privacy and Security

All processing occurs client-side:
- No data transmission to external servers
- Local storage only for configuration persistence
- PDF processing happens entirely in browser memory
- Suitable for confidential financial data

## Integration Points

### Adding New Data Sources
1. Update main dashboard upload processing (`index.html` processCostFiles function)
2. Modify data parsing logic in respective modules
3. Update cache structure if needed (`shared/i2e-cache.js`)
4. Test with sample data files

### Extending Cost Views
1. Add new view types to `switchViewType()` function in validator
2. Create corresponding view generation functions
3. Update project-based pivot logic for new data dimensions
4. Add CSS styling for new view components

### Dashboard Navigation
1. URL parameter handling in `handleDashboardNavigation()` function
2. Tab hiding/showing logic in `hideTabs()` and `hideTabNavigation()`
3. Navigation button state management in `updateNavigationButtons()`
4. Cache monitoring functions for data availability checks

## Current Limitations

- **OCR not supported** - requires machine-readable PDFs
- **Client-side memory limits** - large batches may cause performance issues
- **Browser-specific storage** - data tied to specific browser instance
- **Manual template sharing** - no automated template distribution system

## Claude Code Guidelines

### UX and Interaction Design
- If you are in non-automatic mode then first explain what you want to do and only after approval do it. That is a better UX than having to click ESC each time

### Working with Functions and Tools
- Do NOT assume to know how a function you call is named and what the IO is, check to make sure.