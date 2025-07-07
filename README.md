# I2E Invoice Processor

A client-side PDF invoice data extraction and Excel export tool. This application processes PDF invoices entirely in the browser without requiring server infrastructure or external services.

## Features

- PDF invoice data extraction using client-side processing
- Excel export with configurable field selection
- Invoice validation and duplicate detection
- Credit note handling with automatic amount correction
- Multi-page invoice support with service period extraction
- Hierarchical data organization by invoice and service period
- Real-time validation of extracted totals vs calculated line item sums

## Usage

### Basic Operation

1. Open `I2E_Invoice_Processor.html` in a web browser
2. Upload PDF invoice files using drag-and-drop or file browser
3. Click "Process Invoices" to extract data
4. Review extracted data in the hierarchical table
5. Configure export fields and export to Excel

### Invoice Validation

1. Open `I2E_Invoice_Validator.html` for reviewing processed invoices
2. Filter and validate extracted invoice data
3. Approve or reject invoices based on accuracy
4. Export validated data to Excel

## Technical Requirements

- Modern web browser with ES6+ support
- Internet connection for loading external libraries (PDF.js, ExcelJS)
- JavaScript enabled

## File Structure

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
# Open I2E_Invoice_Processor.html in your browser
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