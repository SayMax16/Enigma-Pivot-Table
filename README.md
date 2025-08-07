# Qlik Sense Pivot Table Data Extractor

A comprehensive Node.js application for extracting pivot table data from Qlik Sense applications using the enigma.js library.

## Features

- **Clean Pivot Extraction**: Extract simplified pivot tables with material descriptions and financial measures
- **Container-based Extraction**: Full container pivot table extraction with all dimensions and measures
- **Field Selection**: Automated field selection and filtering (Plant, Time Period)
- **Multiple Output Formats**: Export data as JSON and CSV
- **Certificate-based Authentication**: Secure connection using Qlik Sense certificates

## Quick Start

### Prerequisites

- Node.js 14.x or higher
- Access to Qlik Sense server
- Valid Qlik Sense certificates (client.pem, client_key.pem, root.pem)

### Installation

1. Clone the repository
2. Install dependencies:
   ```bash
   npm install
   ```
3. Configure certificates in `config/certificates/`
4. Set up environment variables in `.env` file

### Environment Configuration

Create a `.env` file with your Qlik Sense configuration:

```env
QLIK_ENGINE_HOST=your-qlik-server-ip
QLIK_ENGINE_PORT=4747
QLIK_APP_ID=your-app-id
QLIK_USER_DIRECTORY=your-domain
QLIK_USER_ID=your-username
QLIK_CERTIFICATES_PATH=./config/certificates
```

## Usage

### Clean Pivot Table Extraction

Extract a clean, simplified pivot table with material descriptions and 4 financial measures:

```bash
node extract-clean-pivot.js
```

**Output:**
- `clean_pivot_data.csv` - Material descriptions with financial measures
- `clean_pivot_data.json` - Same data in JSON format

**Data Structure:**
- **Краткий текст материала** - Material Description
- **На начало периода** - Beginning Balance
- **ПМ за период** - Goods Received
- **ОМ за период** - Goods Issued  
- **На конец периода** - Ending Balance

### Full Container Extraction

Extract complete pivot table data from container objects:

```bash
node src/index.js
```

## Customization

### Changing Selections

Edit `extract-clean-pivot.js` to modify plant and time period:

```javascript
const selections = [
  { fieldName: 'Завод', value: '1101' },        // Plant
  { fieldName: 'Год-Месяц', value: '2024-авг' } // Year-Month
];
```

### Changing Dimensions

Modify the dimension field in `extract-clean-pivot.js`:

```javascript
// Material Description (current)
qFieldDefs: ['Краткий текст материала']

// Alternative options:
qFieldDefs: ['Номер материала']           // Material Code
qFieldDefs: ['Группа товаров материала']   // Material Group
```

## Project Structure

```
├── config/
│   └── certificates/          # Qlik Sense certificates
├── src/
│   ├── session-manager.js     # Connection management
│   ├── field-selector.js      # Field selection logic
│   ├── container-extractor.js # Container extraction
│   ├── pivot-extractor.js     # Pivot data processing
│   └── index.js              # Main application
├── extract-clean-pivot.js     # Clean pivot extraction
├── package.json
└── README.md
```

## Data Quality

The extraction provides complete data from the Qlik Sense model:
- ✅ Beginning and ending balances match Qlik interface exactly
- ✅ ПМ/ОМ values may be higher as they capture all transaction records
- ✅ The Qlik interface may apply additional business rules or filtering

## Troubleshooting

**Connection Issues:**
- Verify certificates are correctly placed in `config/certificates/`
- Check server IP and port in `.env` file
- Ensure user has access to the specified app

**No Data Returned:**
- Verify field names exist in the Qlik app
- Check that the plant and time period values are valid
- Run with different time periods to find data

**Permission Errors:**
- Ensure user has read access to the Qlik app
- Verify certificate authentication is properly configured

## License

MIT License