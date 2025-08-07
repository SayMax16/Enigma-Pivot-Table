const dotenv = require('dotenv');
const SessionManager = require('./src/session-manager');
const FieldSelector = require('./src/field-selector');
const ContainerExtractor = require('./src/container-extractor');
const fs = require('fs').promises;

dotenv.config();

async function extractCleanPivot() {
  const config = {
    engineHost: '10.7.11.70',
    enginePort: process.env.QLIK_ENGINE_PORT || 4747,
    appId: process.env.QLIK_APP_ID,
    authMethod: process.env.QLIK_AUTH_METHOD || 'certificates',
    userDirectory: process.env.QLIK_USER_DIRECTORY,
    userId: process.env.QLIK_USER_ID,
    certificatesPath: process.env.QLIK_CERTIFICATES_PATH || './config/certificates',
  };

  const sessionManager = new SessionManager(config);
  
  try {
    console.log('🚀 Extracting Clean Pivot Table Data...');
    const { doc } = await sessionManager.connect();
    
    // Apply field selections
    console.log('\n📋 Applying Field Selections...');
    const fieldSelector = new FieldSelector(doc);
    const selections = [
      { fieldName: 'Завод', value: '1101' },
      { fieldName: 'Год-Месяц', value: '2024-авг' }
    ];
    
    await fieldSelector.makeSelections(selections);
    
    // Create a custom pivot table object with only the fields we want
    console.log('\n🔧 Creating Clean Pivot Table...');
    const customPivotObject = await doc.createSessionObject({
      qInfo: { qType: 'table' },
      qHyperCubeDef: {
        qDimensions: [
          {
            qDef: { 
              qFieldDefs: ['Краткий текст материала'],
              qSortCriterias: [{ qSortByLoadOrder: 1 }]
            },
            qNullSuppression: true
          }
        ],
        qMeasures: [
          {
            qDef: { 
              qDef: '[На начало периода]',
              qLabel: 'На начало периода'
            }
          },
          {
            qDef: { 
              qDef: '[ПМ за период]',
              qLabel: 'ПМ за период'
            }
          },
          {
            qDef: { 
              qDef: '[ОМ за период]',
              qLabel: 'ОМ за период'
            }
          },
          {
            qDef: { 
              qDef: '[На конец периода]',
              qLabel: 'На конец периода'
            }
          }
        ],
        qInitialDataFetch: [{
          qLeft: 0,
          qTop: 0,
          qWidth: 5,  // 1 dimension + 4 measures
          qHeight: 1000
        }]
      }
    });
    
    // Get the layout and data
    console.log('\n📊 Extracting Data...');
    const layout = await customPivotObject.getLayout();
    const hypercube = layout.qHyperCube;
    
    console.log(`Total rows available: ${hypercube.qSize.qcy}`);
    console.log(`Columns: ${hypercube.qSize.qcx}`);
    
    // Extract all data with pagination
    const allData = [];
    const pageSize = 100;
    let currentPage = 0;
    const totalRows = hypercube.qSize.qcy;
    
    while (currentPage * pageSize < totalRows) {
      const startRow = currentPage * pageSize;
      const endRow = Math.min(startRow + pageSize - 1, totalRows - 1);
      
      console.log(`Fetching page ${currentPage + 1}: rows ${startRow} to ${endRow}`);
      
      const dataPages = await customPivotObject.getHyperCubeData('/qHyperCubeDef', [{
        qLeft: 0,
        qTop: startRow,
        qWidth: hypercube.qSize.qcx,
        qHeight: endRow - startRow + 1
      }]);
      
      if (dataPages && dataPages[0] && dataPages[0].qMatrix) {
        allData.push(...dataPages[0].qMatrix);
        console.log(`Page ${currentPage + 1}: ${dataPages[0].qMatrix.length} rows fetched`);
      }
      
      currentPage++;
      
      // Safety break
      if (currentPage > 100) {
        console.log('⚠️ Safety limit reached (100 pages)');
        break;
      }
    }
    
    console.log(`\nTotal extracted: ${allData.length} rows`);
    
    // Format the data
    console.log('\n🔄 Formatting Data...');
    const formattedData = [];
    let nonZeroRowsCount = 0;
    
    allData.forEach((row, index) => {
      const materialCode = row[0]?.qText || '';
      const beginningBalance = parseFloat(row[1]?.qNum) || 0;
      const goodsReceived = parseFloat(row[2]?.qNum) || 0;
      const goodsIssued = parseFloat(row[3]?.qNum) || 0;
      const endingBalance = parseFloat(row[4]?.qNum) || 0;
      
      // Check if this row has any non-zero values
      const hasNonZeroValues = beginningBalance !== 0 || goodsReceived !== 0 || goodsIssued !== 0 || endingBalance !== 0;
      
      if (hasNonZeroValues) {
        nonZeroRowsCount++;
      }
      
      // Only include rows with material codes and some data
      if (materialCode.trim() !== '' || hasNonZeroValues) {
        formattedData.push({
          'Краткий текст материала': materialCode,
          'На начало периода': beginningBalance,
          'ПМ за период': goodsReceived,
          'ОМ за период': goodsIssued,
          'На конец периода': endingBalance
        });
      }
    });
    
    console.log(`\n📊 Data Analysis:`);
    console.log(`- Total formatted rows: ${formattedData.length}`);
    console.log(`- Rows with non-zero values: ${nonZeroRowsCount}`);
    
    // Show sample of data with values
    if (nonZeroRowsCount > 0) {
      console.log('\n📋 Sample rows with values (first 10):');
      let sampleCount = 0;
      formattedData.forEach((row, index) => {
        if (sampleCount >= 10) return;
        
        const hasNonZero = row['На начало периода'] !== 0 || 
                          row['ПМ за период'] !== 0 || 
                          row['ОМ за период'] !== 0 || 
                          row['На конец периода'] !== 0;
        
        if (hasNonZero && row['Краткий текст материала'].trim() !== '') {
          sampleCount++;
          console.log(`\n${sampleCount}. ${row['Краткий текст материала']}:`);
          if (row['На начало периода'] !== 0) console.log(`   На начало периода: ${row['На начало периода'].toLocaleString()}`);
          if (row['ПМ за период'] !== 0) console.log(`   ПМ за период: ${row['ПМ за период'].toLocaleString()}`);
          if (row['ОМ за период'] !== 0) console.log(`   ОМ за период: ${row['ОМ за период'].toLocaleString()}`);
          if (row['На конец периода'] !== 0) console.log(`   На конец периода: ${row['На конец периода'].toLocaleString()}`);
        }
      });
    } else {
      console.log('\n⚠️ No rows with non-zero values found');
      console.log('First 5 rows anyway:');
      formattedData.slice(0, 5).forEach((row, index) => {
        console.log(`\n${index + 1}. ${row['Краткий текст материала'] || 'No material code'}:`);
        console.log(`   На начало периода: ${row['На начало периода']}`);
        console.log(`   ПМ за период: ${row['ПМ за период']}`);
        console.log(`   ОМ за период: ${row['ОМ за период']}`);
        console.log(`   На конец периода: ${row['На конец периода']}`);
      });
    }
    
    // Save the clean data
    console.log('\n💾 Saving Clean Pivot Data...');
    
    // Save as JSON
    await fs.writeFile('clean_pivot_data.json', JSON.stringify(formattedData, null, 2));
    
    // Save as CSV
    const csvHeaders = 'Краткий текст материала,На начало периода,ПМ за период,ОМ за период,На конец периода';
    const csvRows = formattedData.map(row => {
      return [
        `"${row['Краткий текст материала']}"`,
        row['На начало периода'],
        row['ПМ за период'],
        row['ОМ за период'],
        row['На конец периода']
      ].join(',');
    });
    
    const csvContent = [csvHeaders, ...csvRows].join('\n');
    await fs.writeFile('clean_pivot_data.csv', csvContent);
    
    console.log(`\n✅ Clean pivot data saved:`);
    console.log(`- clean_pivot_data.json (${formattedData.length} rows)`);
    console.log(`- clean_pivot_data.csv (${formattedData.length} rows)`);
    
    // Summary statistics
    console.log('\n📊 SUMMARY STATISTICS:');
    const measures = ['На начало периода', 'ПМ за период', 'ОМ за период', 'На конец периода'];
    measures.forEach(measure => {
      const values = formattedData.map(row => row[measure]);
      const nonZeroValues = values.filter(v => v !== 0);
      const total = values.reduce((sum, val) => sum + val, 0);
      const avg = nonZeroValues.length > 0 ? total / nonZeroValues.length : 0;
      const max = Math.max(...values);
      
      console.log(`${measure}:`);
      console.log(`  - Total: ${total.toLocaleString()}`);
      console.log(`  - Non-zero entries: ${nonZeroValues.length} / ${values.length}`);
      console.log(`  - Average (non-zero): ${avg.toLocaleString()}`);
      console.log(`  - Maximum: ${max.toLocaleString()}`);
    });
    
    await sessionManager.close();
    console.log('\n✅ Clean pivot extraction completed!');
    
  } catch (error) {
    console.error('❌ Clean pivot extraction failed:', error);
    await sessionManager.close();
    process.exit(1);
  }
}

extractCleanPivot();