# I2E (Invoice to Excel) - Cost Analysis & Invoice Management System

A browser-based system for cost data analysis and invoice processing. Runs entirely in the browser without server dependencies.

## Usage

1. Open `index.html` in your web browser
2. Use the DEMO card to load sample data, or upload your own files
3. Upload Excel files containing cost data (PPM, EXT SAP, I2E data)
4. Upload PDF invoices for data extraction and validation
5. Use the cost analysis tools to review project data
6. Process invoices through the validation and approval workflow

## Features

### Main Dashboard
- Data overview with file and invoice statistics
- File upload via drag-and-drop or file selection
- Navigation to different functional areas
- Demo mode with sample data

### Cost Data Analysis
- Regular views: field-based analysis by employee, supplier, cost type
- Project views: expandable project-centric analysis
- Project filtering and multi-select capabilities
- Monthly cost analysis and supplier breakdowns

### Invoice Processing
- PDF text extraction using VAT pattern detection
- Multi-page invoice support
- Service period extraction from headers
- Credit note handling with negative amounts
- Project ID extraction (AA44-PRO0012345 and PRO0012345 formats)

### Invoice Validation & Approval
- Cost validation by comparing against PPM and EXT SAP data
- Visual indicators for validation errors and credit notes
- Individual and batch invoice approval/rejection
- Comment tracking and audit trail

## Architecture

### Frontend-Only Design
- Browser-based processing with no server requirements
- Local storage for configuration and cache persistence
- Client-side PDF processing (PDF.js) and Excel generation (ExcelJS)

### File Structure
```
index.html                     # Main dashboard - start here
I2E_Invoice_Processor.html     # PDF extraction (standalone)
I2E_Invoice_Validator.html     # Cost validation & approval
shared/
├── i2e-common.js             # Shared utilities (file handling, currency, validation)
├── i2e-cache.js              # Local storage management and data caching
├── i2e-spy.js                # Performance monitoring and cache inspection
├── pdf-extractor.js          # PDF processing and data extraction
└── excel-exporter.js         # Excel generation and formatting
assets/
└── i2e-styles.css            # Unified styling with responsive design
```

## Use Cases

### Finance Teams
- Automated invoice data extraction from PDF files
- Cost validation against SAP and PPM data
- Structured approval workflows with audit trails

### Project Managers
- Project-specific cost analysis and reporting
- Budget monitoring and variance tracking
- Supplier cost analysis

### Controllers & Analysts
- Custom Excel exports with flexible field selection
- Multi-dimensional cost reporting
- Planned vs actual cost comparison

## Technical Requirements

### Browser Compatibility
- Chrome 80+, Firefox 75+, Safari 13+, Edge 80+
- JavaScript required
- 2GB+ RAM recommended for large files

### Data Processing
- PDF text extraction from machine-readable PDFs (no OCR support)
- Excel parsing (.xlsx files up to 50MB)
- Browser-based local storage caching (10MB+ capacity)

### Security & Privacy
- All processing occurs locally in the browser
- No external data transmission
- Configuration stored in browser local storage only

## Supported File Types

### Excel Files (Cost Data)
- PPM Data: Project cost data with WBS codes
- EXT SAP Data: External supplier costs
- I2E Data: Invoice data for validation

### PDF Files (Invoices)
- Text-based PDFs with machine-readable content
- Multi-page invoices with complex structures
- Various invoice layouts and formats

## 🎯 Getting Started

### Step 1: Setup
1. Download or clone the repository
2. Open `index.html` in your web browser
3. No installation or configuration required

### Step 2: Try the Demo (Recommended)
1. Click the **DEMO** card on the dashboard
2. System loads pre-sanitized sample data instantly
3. Explore all features without uploading your own files
4. Perfect for testing, training, or demonstrations

### Step 3: Upload Your Data
1. Drag Excel files to the dashboard upload area
2. System automatically detects PPM, EXT SAP, and I2E worksheets
3. Data is cached locally for instant access

### Step 4: Process Invoices
1. Upload PDF invoices via dashboard or processor
2. Review extracted line items and validate data
3. System compares against cost data for validation

### Step 5: Manage & Export
1. Use validator to approve/reject invoices
2. Access cost views for analysis
3. Export data to Excel with custom fields

## 🎬 Demo Mode

The I2E system includes a comprehensive demo mode that provides instant access to sample data for testing and demonstrations.

### Key Benefits
- **No Setup Required**: Click once to load pre-sanitized sample data
- **Full Feature Access**: Experience all system capabilities immediately
- **Risk-Free Testing**: No real data exposure during evaluation
- **Perfect for Training**: Ideal for onboarding new users or client demos

### Demo Data Included
- **Cost Data**: Sample PPM and EXT SAP data with realistic project structures
- **Invoice Data**: Pre-processed PDF invoices with various scenarios
- **Project Coverage**: Multiple projects with different cost patterns
- **Validation Examples**: Mix of valid and validation-requiring invoices

### How to Use
1. **Access**: Click the purple **DEMO** card on the main dashboard
2. **Explore**: All features become available with sample data
3. **Test**: Try cost views, invoice validation, and approval workflows
4. **Reset**: Use "Clear demo data" button to return to clean state

### Smart Demo Management
- **Auto-Hide**: Demo card disappears when you upload real data
- **State Persistence**: Demo remains available after page refresh
- **Clean Reset**: Clearing demo data returns to initial state
- **No Interference**: Demo and real data never mix

## 🚀 Navigation Tips

- **Always start with `index.html`** - it's your command center
- **Try the DEMO first** - loads sample data instantly for immediate exploration
- **Dashboard shows data availability** - buttons activate when data is ready
- **Cache inspector** (spy icon) provides detailed system information
- **Help button** opens comprehensive documentation modal

## 🔍 Troubleshooting

### Common Issues
- **PDF extraction fails**: Ensure PDFs contain machine-readable text (not scanned images)
- **Large file issues**: Browser memory limits may affect very large files
- **Cache problems**: Use the spy icon to inspect and clear cache if needed
- **Missing data**: Check dashboard for data availability before navigation

### Performance Tips
- **Upload data incrementally**: Process smaller batches for better performance
- **Clear cache regularly**: Use built-in cache management tools
- **Monitor storage**: Spy icon shows storage usage and recommendations

## 📋 Data Privacy & Compliance

- **Local Processing Only**: No data leaves your browser
- **GDPR Compliant**: No external data transmission
- **Secure**: Documents processed entirely on your machine
- **Auditable**: All operations logged locally for transparency

## 🚀 Production Ready

This system is ready for production use with:
- ✅ Clean, professional logging (debug noise removed)
- ✅ Comprehensive error handling
- ✅ User-friendly interfaces
- ✅ Complete documentation
- ✅ Responsive design for all screen sizes
- ✅ Robust data validation and processing

---

**I2E System** - Transforming invoice processing from manual work to automated intelligence.