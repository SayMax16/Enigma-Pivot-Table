# Enigma Pivot Table Extractor

An optimized Node.js application for extracting data from Qlik Sense pivot tables using enigma.js. This tool connects to your Qlik Sense app, makes specific field selections, and extracts pivot table data with performance optimizations for minimal CPU and memory usage.

## Features

- **Multiple Authentication Methods**: Certificate-based, API Key, and JWT authentication
- **Optimized Data Extraction**: Paginated data fetching to minimize memory usage
- **Automatic Retry Logic**: Handles aborted requests during heavy calculations
- **Field Selection**: Automated selection of Завод and YearMonth fields
- **Multiple Export Formats**: JSON and CSV output
- **Real-time Monitoring**: Optional WebSocket traffic logging
- **Performance Optimized**: Delta protocol, request batching, and memory management

## Prerequisites

- Node.js >= 14.0
- Access to Qlik Sense (on-premise or cloud)
- Qlik Sense app with a pivot table object
- Authentication credentials (certificates, API key, or JWT)

## Installation

1. Clone the repository:
```bash
git clone https://github.com/SayMax16/Enigma-Pivot-Table.git
cd Enigma-Pivot-Table
```
2. Install dependencies:
```bash
npm install
```

3. Copy the environment configuration:
```bash
cp .env.example .env
```

4. Configure your settings in `.env` file

## Configuration

### Environment Variables

Edit the `.env` file with your Qlik Sense configuration:

#### Connection Settings
```env
QLIK_ENGINE_HOST=your-qlik-server.com
QLIK_ENGINE_PORT=4747
QLIK_APP_ID=your-app-id-here
```

#### Authentication (choose one method)

**Certificate Authentication (On-premise):**
```env
QLIK_AUTH_METHOD=certificates
QLIK_USER_DIRECTORY=YOUR_DOMAIN
QLIK_USER_ID=your-username
QLIK_CERTIFICATES_PATH=./config/certificates
```

**API Key Authentication (Qlik Cloud):**
```env
QLIK_AUTH_METHOD=apikey
QLIK_API_KEY=your-api-key-here
```

**JWT Authentication:**
```env
QLIK_AUTH_METHOD=jwt
QLIK_JWT_TOKEN=your-jwt-token-here
```

#### Data Settings
```env
QLIK_PIVOT_OBJECT_ID=your-pivot-object-id
QLIK_ZAVOD_FIELD=Завод
QLIK_ZAVOD_VALUE=1101
QLIK_YEARMONTH_FIELD=YearMonth
QLIK_YEARMONTH_VALUE=2025.01
```

### Certificate Setup (for on-premise)

If using certificate authentication, place your certificate files in `config/certificates/`:
- `root.pem` - Root certificate
- `client.pem` - Client certificate  
- `client_key.pem` - Client private key

## Usage

### Basic Usage

Run the data extraction:
```bash
npm start
```

### Development Mode

Run with auto-restart on changes:
```bash
npm run dev
```

### Programmatic Usage

```javascript
const QlikPivotDataExtractor = require('./src/index');

const extractor = new QlikPivotDataExtractor();

extractor.run()
  .then((data) => {
    console.log('Data extracted successfully:', data.summary);
  })
  .catch((error) => {
    console.error('Extraction failed:', error);
  });
```

## Output

The application generates two output files:

1. **pivot_data.json** - Complete data with metadata in JSON format
2. **pivot_data.csv** - Formatted data in CSV format

### Data Structure

```javascript
{
  "headers": [
    { "name": "Завод", "type": "dimension", "index": 0 },
    { "name": "YearMonth", "type": "dimension", "index": 1 },
    { "name": "Revenue", "type": "measure", "index": 2 }
  ],
  "rows": [
    {
      "index": 0,
      "data": {
        "Завод": { "text": "1101", "number": 1101, "state": "S" },
        "YearMonth": { "text": "2025.01", "number": null, "state": "S" },
        "Revenue": { "text": "1,234,567", "number": 1234567, "state": "L" }
      }
    }
  ],
  "metadata": {
    "dimensions": [...],
    "measures": [...],
    "totalRows": 150,
    "extractedRows": 150
  },
  "summary": {
    "totalRows": 150,
    "totalColumns": 3,
    "dimensions": 2,
    "measures": 1
  }
}
```

## Performance Optimization

The application includes several optimizations:

- **Paginated Data Fetching**: Configurable page size (default: 1000 rows)
- **Request Retry Logic**: Automatic retry for aborted requests
- **Delta Protocol**: Reduced bandwidth usage (enabled by default)
- **Memory Management**: Proper cleanup of objects and sessions
- **Connection Reuse**: Single session for multiple operations

### Performance Settings

```env
QLIK_PAGE_SIZE=1000        # Rows per page
QLIK_MAX_PAGES=10          # Maximum pages to fetch
```

## Troubleshooting

### Common Issues

1. **Authentication Errors**
   - Verify certificate files are in correct location
   - Check user directory and user ID are correct
   - Ensure API key has proper permissions

2. **Object Not Found**
   - Verify the pivot object ID exists in the app
   - Check app ID is correct
   - Ensure user has access to the app

3. **Field Selection Errors**
   - Verify field names match exactly (case-sensitive)
   - Check that field values exist in the data
   - Ensure fields are available in the data model

4. **Performance Issues**
   - Reduce page size for large datasets
   - Increase max pages limit if needed
   - Enable traffic logging to debug requests

### Debug Mode

Enable traffic logging to see WebSocket communication:
```env
QLIK_ENABLE_TRAFFIC_LOGGING=true
```

## Architecture

The application consists of four main components:

- **SessionManager** (`src/session-manager.js`) - Handles Qlik Sense connections and authentication
- **FieldSelector** (`src/field-selector.js`) - Manages field selections and filtering
- **PivotExtractor** (`src/pivot-extractor.js`) - Extracts and formats pivot table data
- **Main Application** (`src/index.js`) - Orchestrates the entire process

## Error Handling

The application includes comprehensive error handling:
- Automatic retry for aborted requests
- Graceful session cleanup
- Detailed error logging
- Fallback methods for data extraction

## License

MIT License - see LICENSE file for details.

## Support

For issues related to:
- **enigma.js**: Check the [enigma.js documentation](https://github.com/qlik-oss/enigma.js)
- **Qlik Sense APIs**: Refer to [Qlik Sense Developer Documentation](https://help.qlik.com/en-US/sense-developer/)
- **This Application**: Create an issue in the project repository