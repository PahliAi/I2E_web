# I2E Invoice Processor

A client-side PDF invoice data extraction and Excel export tool. This application processes PDF invoices entirely in the browser without requiring server infrastructure or external services.
Warning: this version is specifically built for 1 company with it's very specific invoice structure and cost excels structure. this toll will not work for any other company without copying the code and customizing it. 

## Features

- PDF invoice data extraction using client-side processing
- Excel export with configurable field selection
- Invoice validation and duplicate detection
- Credit note handling with automatic amount correction
- Multi-page invoice support with service period extraction
- Hierarchical data organization by invoice and service period
- Real-time validation of extracted totals vs calculated line item sums
- Multi dynamic cost overviews: optionally per project and per year/month/week period

## Usage

### Basic Operation

1. Open `index.html` in a web browser
2. Select option Cost Analysis, Invoice 2 Excel or Invoice Validation
3. Upload PDF invoice files and excel Cost data files using drag-and-drop or file browser
4. Click "Continue" to visit the selected option page

## Technical Requirements

- Modern web browser with ES6+ support
- Internet connection for loading external libraries (PDF.js, ExcelJS)
- JavaScript enabled

## File Structure

```
I2E_web/
├── index.html                    # Landing page with options and upload area
├── I2E_Invoice_Processor.html    # Invoice extraction interface
├── I2E_Invoice_Validator.html    # Invoice validation interface
├── I2E_Cost_view.html            # Cost view interface (dynamic tables)
├── assets/
│   └── i2e-styles.css           # Shared CSS framework
└── shared/
    ├── i2e-common.js            # Common utilities
    ├── pdf-extractor.js         # PDF processing logic
    ├── excel-exporter.js        # Excel export functionality
    ├── i2e-cache.js             # Client-side caching
    └── i2e-spy.js               # Development utilities
```

## Data Extraction

The application extracts the following invoice information:

- Invoice metadata (number, date, customer ID, project ID)
- Line item details (position, description, quantity, unit price, total)
- Service provision periods
- VAT information and currency
- Credit note detection and handling

## Export Options

### Field Selection
- Default fields: File name, customer ID, invoice number, date, project ID, service period, quantity, unit price, total
- Detailed report: All available fields including material codes, cost types, VAT
- Summary report: Essential fields only
- Custom selection: Choose specific fields and ordering

### Export Formats
- Excel workbook with multiple sheets
- Invoice details sheet with line items
- Summary sheet with invoice totals
- Configurable field ordering and selection

## Browser Compatibility

- Chrome 80+
- Firefox 75+
- Safari 13+
- Edge 80+

## Security

- All processing occurs client-side
- No data transmission to external servers
- Files processed in browser memory only
- Configuration stored in browser localStorage
- No authentication required

## Installation

No installation required. Clone or download the repository and open the HTML files in a web browser.

```bash
git clone <repository-url>
cd I2E_web
```

## Development

This is a pure frontend application with no build process required. All files are served directly from the filesystem.

### Dependencies
- PDF.js (v3.11.174) - loaded via CDN
- ExcelJS (v4.4.0) - loaded via CDN

### Testing
Manual testing through browser interface. No automated test suite.

## License

See LICENSE file for details.