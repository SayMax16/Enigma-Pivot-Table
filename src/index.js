const dotenv = require('dotenv');
const fs = require('fs');
const path = require('path');
const SessionManager = require('./session-manager');
const FieldSelector = require('./field-selector');
const PivotExtractor = require('./pivot-extractor');
const FieldDebugger = require('./field-debugger');

// Load environment variables
dotenv.config();

class QlikPivotDataExtractor {
  constructor() {
    this.sessionManager = null;
    this.fieldSelector = null;
    this.pivotExtractor = null;
    this.config = this.loadConfiguration();
  }

  // Load configuration from environment variables
  loadConfiguration() {
    const config = {
      // Connection settings
      engineHost: process.env.QLIK_ENGINE_HOST || 'localhost',
      enginePort: process.env.QLIK_ENGINE_PORT || 4747,
      appId: process.env.QLIK_APP_ID,
      
      // Authentication settings
      authMethod: process.env.QLIK_AUTH_METHOD || 'certificates', // certificates, apikey, jwt
      userDirectory: process.env.QLIK_USER_DIRECTORY,
      userId: process.env.QLIK_USER_ID,
      certificatesPath: process.env.QLIK_CERTIFICATES_PATH || './config/certificates',
      apiKey: process.env.QLIK_API_KEY,
      jwtToken: process.env.QLIK_JWT_TOKEN,
      
      // Data extraction settings
      pivotObjectId: process.env.QLIK_PIVOT_OBJECT_ID,
      
      // Field selection settings
      zavodField: process.env.QLIK_ZAVOD_FIELD || 'Завод',
      zavodValue: process.env.QLIK_ZAVOD_VALUE || '1101',
      yearMonthField: process.env.QLIK_YEARMONTH_FIELD || 'YearMonth',
      yearMonthValue: process.env.QLIK_YEARMONTH_VALUE || '2025.01',
      
      // Performance settings
      pageSize: parseInt(process.env.QLIK_PAGE_SIZE) || 1000,
      maxPages: parseInt(process.env.QLIK_MAX_PAGES) || 10,
      enableTrafficLogging: process.env.QLIK_ENABLE_TRAFFIC_LOGGING === 'true',
      debugFields: process.env.QLIK_DEBUG_FIELDS === 'true',
    };

    // Validate required configuration
    this.validateConfiguration(config);
    
    return config;
  }

  // Validate that required configuration is present
  validateConfiguration(config) {
    const required = ['appId', 'pivotObjectId'];
    
    if (config.authMethod === 'certificates') {
      required.push('userDirectory', 'userId');
    } else if (config.authMethod === 'apikey') {
      required.push('apiKey');
    } else if (config.authMethod === 'jwt') {
      required.push('jwtToken');
    }

    const missing = required.filter(key => !config[key]);
    
    if (missing.length > 0) {
      throw new Error(`Missing required configuration: ${missing.join(', ')}`);
    }
  }

  // Initialize all components
  async initialize() {
    try {
      console.log('Initializing Qlik Pivot Data Extractor...');
      
      // Create session manager
      this.sessionManager = new SessionManager(this.config);
      
      // Enable traffic logging if requested
      if (this.config.enableTrafficLogging) {
        this.sessionManager.enableTrafficLogging();
      }
      
      // Connect to Qlik Sense
      const { doc } = await this.sessionManager.connect();
      
      // Initialize field selector, pivot extractor, and debugger
      this.fieldSelector = new FieldSelector(doc);
      this.pivotExtractor = new PivotExtractor(doc);
      this.fieldDebugger = new FieldDebugger(doc);
      
      console.log('Initialization completed successfully');
      
    } catch (error) {
      console.error('Initialization failed:', error);
      throw error;
    }
  }

  // Make required field selections
  async makeSelections() {
    try {
      console.log('Making field selections...');
      
      // Debug field selection if enabled
      if (this.config.debugFields) {
        console.log('🔍 Debug mode enabled - analyzing fields...');
        
        // Debug both fields
        await this.fieldDebugger.debugFieldSelection(this.config.zavodField, this.config.zavodValue);
        await this.fieldDebugger.debugFieldSelection(this.config.yearMonthField, this.config.yearMonthValue);
        
        // Look for similar fields to yearMonth
        await this.fieldDebugger.findFieldsContaining('Год');
        await this.fieldDebugger.findFieldsContaining('Месяц');
        await this.fieldDebugger.findFieldsContaining('Month');
        await this.fieldDebugger.findFieldsContaining('Year');
      }
      
      const selections = [
        {
          fieldName: this.config.zavodField,
          value: this.config.zavodValue,
        },
        {
          fieldName: this.config.yearMonthField,
          value: this.config.yearMonthValue,
        },
      ];
      
      await this.fieldSelector.makeSelections(selections);
      
      // Verify selections
      await this.fieldSelector.verifySelections(selections);
      
      console.log('Field selections completed successfully');
      
    } catch (error) {
      console.error('Field selection failed:', error);
      throw error;
    }
  }

  // Extract pivot table data
  async extractData() {
    try {
      console.log('Starting data extraction...');
      
      // Get pivot object
      const pivotObject = await this.pivotExtractor.getPivotObject(this.config.pivotObjectId);
      
      // Extract data with optimization settings
      const extractedData = await this.pivotExtractor.extractPivotData(pivotObject, {
        pageSize: this.config.pageSize,
        maxPages: this.config.maxPages,
      });
      
      // Format data for easier consumption
      const formattedData = this.pivotExtractor.formatPivotData(extractedData);
      
      console.log('Data extraction completed successfully');
      console.log(`Extracted ${formattedData.summary.totalRows} rows, ${formattedData.summary.totalColumns} columns`);
      
      return formattedData;
      
    } catch (error) {
      console.error('Data extraction failed:', error);
      throw error;
    }
  }

  // Save data to file
  async saveDataToFile(data, filename = 'pivot_data.json') {
    try {
      const outputPath = path.resolve(filename);
      
      if (filename.endsWith('.csv')) {
        const csvData = this.pivotExtractor.exportToCSV(data);
        fs.writeFileSync(outputPath, csvData, 'utf8');
      } else {
        fs.writeFileSync(outputPath, JSON.stringify(data, null, 2), 'utf8');
      }
      
      console.log(`Data saved to: ${outputPath}`);
      return outputPath;
      
    } catch (error) {
      console.error('Failed to save data to file:', error);
      throw error;
    }
  }

  // Print data summary
  printDataSummary(data) {
    console.log('\n=== DATA EXTRACTION SUMMARY ===');
    console.log(`Total Rows: ${data.summary.totalRows}`);
    console.log(`Total Columns: ${data.summary.totalColumns}`);
    console.log(`Dimensions: ${data.summary.dimensions}`);
    console.log(`Measures: ${data.summary.measures}`);
    
    console.log('\nColumn Headers:');
    data.headers.forEach((header, index) => {
      console.log(`  ${index + 1}. ${header.name} (${header.type})`);
    });
    
    if (data.rows.length > 0) {
      console.log('\nFirst 3 rows (sample data):');
      data.rows.slice(0, 3).forEach((row, index) => {
        console.log(`Row ${index + 1}:`);
        Object.entries(row.data).forEach(([key, value]) => {
          console.log(`  ${key}: ${value.text}`);
        });
        console.log('');
      });
    }
  }

  // Clean shutdown
  async shutdown() {
    try {
      if (this.sessionManager) {
        await this.sessionManager.close();
      }
      console.log('Shutdown completed successfully');
    } catch (error) {
      console.error('Shutdown error:', error);
    }
  }

  // Main execution method
  async run() {
    try {
      console.log('Starting Qlik Pivot Data Extraction...');
      console.log('Configuration:');
      console.log(`- App ID: ${this.config.appId}`);
      console.log(`- Pivot Object ID: ${this.config.pivotObjectId}`);
      console.log(`- ${this.config.zavodField}: ${this.config.zavodValue}`);
      console.log(`- ${this.config.yearMonthField}: ${this.config.yearMonthValue}`);
      
      // Initialize connection
      await this.initialize();
      
      // Make required selections
      await this.makeSelections();
      
      // Extract data
      const data = await this.extractData();
      
      // Print summary
      this.printDataSummary(data);
      
      // Save data to files
      await this.saveDataToFile(data, 'pivot_data.json');
      await this.saveDataToFile(data, 'pivot_data.csv');
      
      console.log('\n✅ Data extraction completed successfully!');
      
      return data;
      
    } catch (error) {
      console.error('\n❌ Data extraction failed:', error);
      throw error;
    } finally {
      await this.shutdown();
    }
  }
}

// Run the extractor if this file is executed directly
if (require.main === module) {
  const extractor = new QlikPivotDataExtractor();
  
  extractor.run()
    .then(() => {
      console.log('Process completed successfully');
      process.exit(0);
    })
    .catch((error) => {
      console.error('Process failed:', error);
      process.exit(1);
    });
}

module.exports = QlikPivotDataExtractor;