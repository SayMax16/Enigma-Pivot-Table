class PivotExtractor {
  constructor(doc) {
    this.doc = doc;
  }

  // Get existing pivot table object by ID
  async getPivotObject(objectId) {
    try {
      console.log(`Getting pivot object with ID: ${objectId}`);
      const pivotObject = await this.doc.getObject(objectId);
      console.log('Pivot object retrieved successfully');
      return pivotObject;
    } catch (error) {
      console.error(`Failed to get pivot object ${objectId}:`, error);
      throw error;
    }
  }

  // Get pivot table layout (metadata and structure)
  async getPivotLayout(pivotObject) {
    try {
      console.log('Getting pivot layout...');
      const layout = await pivotObject.getLayout();
      
      const hypercube = layout.qHyperCube;
      console.log('Pivot layout info:');
      console.log(`- Dimensions: ${hypercube.qDimensionInfo.length}`);
      console.log(`- Measures: ${hypercube.qMeasureInfo.length}`);
      console.log(`- Total rows: ${hypercube.qSize.qcy}`);
      console.log(`- Total columns: ${hypercube.qSize.qcx}`);
      console.log(`- Data mode: ${hypercube.qMode}`);
      
      return layout;
    } catch (error) {
      console.error('Failed to get pivot layout:', error);
      throw error;
    }
  }

  // Extract pivot data with optimization (pagination)
  async extractPivotData(pivotObject, options = {}) {
    try {
      const {
        pageSize = 1000,     // Number of rows per page
        maxPages = 10,       // Maximum pages to fetch (safety limit)
        startRow = 0,        // Starting row
        startCol = 0,        // Starting column
        columnCount = null,  // Number of columns (null = all)
      } = options;

      console.log('Starting optimized pivot data extraction...');
      
      // Get layout first to understand data structure
      const layout = await this.getPivotLayout(pivotObject);
      const hypercube = layout.qHyperCube;
      
      const totalRows = hypercube.qSize.qcy;
      const totalCols = columnCount || hypercube.qSize.qcx;
      
      console.log(`Extracting data: ${totalRows} rows, ${totalCols} columns`);
      
      const allData = [];
      let currentRow = startRow;
      let pagesProcessed = 0;

      // Extract data in chunks (pagination for memory optimization)
      while (currentRow < totalRows && pagesProcessed < maxPages) {
        const remainingRows = totalRows - currentRow;
        const rowsToFetch = Math.min(pageSize, remainingRows);
        
        console.log(`Fetching page ${pagesProcessed + 1}: rows ${currentRow} to ${currentRow + rowsToFetch - 1}`);
        
        try {
          let pageData;
          
          // Choose extraction method based on data mode
          if (hypercube.qMode === 'EQ_DATA_MODE_PIVOT') {
            // Use getHyperCubePivotData for pivot tables
            pageData = await pivotObject.getHyperCubePivotData('/qHyperCubeDef', [{
              qTop: currentRow,
              qLeft: startCol,
              qHeight: rowsToFetch,
              qWidth: totalCols,
            }]);
          } else {
            // Use getHyperCubeData for straight tables
            pageData = await pivotObject.getHyperCubeData('/qHyperCubeDef', [{
              qTop: currentRow,
              qLeft: startCol,
              qHeight: rowsToFetch,
              qWidth: totalCols,
            }]);
          }
          
          if (pageData && pageData.length > 0 && pageData[0].qMatrix) {
            const matrixData = pageData[0].qMatrix;
            allData.push(...matrixData);
            console.log(`Page ${pagesProcessed + 1}: ${matrixData.length} rows fetched`);
          } else {
            console.log(`Page ${pagesProcessed + 1}: No data returned`);
            break;
          }
          
          currentRow += rowsToFetch;
          pagesProcessed++;
          
          // Small delay to prevent overloading the server
          if (pagesProcessed < maxPages && currentRow < totalRows) {
            await new Promise(resolve => setTimeout(resolve, 10));
          }
          
        } catch (pageError) {
          console.error(`Error fetching page ${pagesProcessed + 1}:`, pageError);
          
          // Handle specific error types
          if (pageError.code === 6001 || pageError.parameter === 'Page(s) too large') {
            console.log(`Page too large, reducing page size from ${rowsToFetch} to ${Math.floor(rowsToFetch / 2)}`);
            // Reduce page size and retry
            const newPageSize = Math.max(10, Math.floor(rowsToFetch / 2));
            
            try {
              const pageData = await pivotObject.getHyperCubeData('/qHyperCubeDef', [{
                qTop: currentRow,
                qLeft: startCol,
                qHeight: newPageSize,
                qWidth: totalCols,
              }]);
              
              if (pageData && pageData.length > 0 && pageData[0].qMatrix) {
                const matrixData = pageData[0].qMatrix;
                allData.push(...matrixData);
                console.log(`Page ${pagesProcessed + 1}: ${matrixData.length} rows fetched (reduced size)`);
                currentRow += matrixData.length;
                pagesProcessed++;
                continue;
              }
            } catch (retryError) {
              console.error('Retry with smaller page size failed:', retryError);
              break;
            }
          } else if (pageError.code === 6002 || pageError.parameter === 'Not in pivot mode') {
            console.log('Object is not in pivot mode, trying straight table method...');
            // Try straight table method once
            try {
              const pageData = await pivotObject.getHyperCubeData('/qHyperCubeDef', [{
                qTop: currentRow,
                qLeft: startCol,
                qHeight: rowsToFetch,
                qWidth: totalCols,
              }]);
              
              if (pageData && pageData.length > 0 && pageData[0].qMatrix) {
                const matrixData = pageData[0].qMatrix;
                allData.push(...matrixData);
                console.log(`Page ${pagesProcessed + 1}: ${matrixData.length} rows fetched (straight table)`);
                currentRow += rowsToFetch;
                pagesProcessed++;
                continue;
              }
            } catch (straightError) {
              console.error('Straight table method also failed:', straightError);
              break;
            }
          } else if (pageError.code === 'LOCERR_GENERIC_ABORTED') {
            console.log('Request aborted, retrying...');
            continue; // Retry same page
          } else {
            console.log(`Stopping extraction due to error: ${pageError.message}`);
            break; // Stop on other errors
          }
        }
      }
      
      console.log(`Data extraction completed: ${allData.length} total rows`);
      
      return {
        data: allData,
        metadata: {
          dimensions: hypercube.qDimensionInfo,
          measures: hypercube.qMeasureInfo,
          totalRows: totalRows,
          totalColumns: totalCols,
          extractedRows: allData.length,
          pageSize: pageSize,
          pagesProcessed: pagesProcessed,
        },
      };
      
    } catch (error) {
      console.error('Failed to extract pivot data:', error);
      throw error;
    }
  }

  // Alternative method: Extract data using getLayout (for smaller datasets)
  async extractPivotDataSimple(pivotObject) {
    try {
      console.log('Extracting pivot data using simple method (getLayout)...');
      
      const layout = await pivotObject.getLayout();
      const hypercube = layout.qHyperCube;
      
      if (hypercube.qDataPages && hypercube.qDataPages.length > 0) {
        const allData = [];
        hypercube.qDataPages.forEach(page => {
          if (page.qMatrix) {
            allData.push(...page.qMatrix);
          }
        });
        
        return {
          data: allData,
          metadata: {
            dimensions: hypercube.qDimensionInfo,
            measures: hypercube.qMeasureInfo,
            totalRows: allData.length,
            extractedRows: allData.length,
          },
        };
      } else {
        console.log('No data pages found in layout, trying pivot-specific method...');
        return await this.extractPivotData(pivotObject);
      }
      
    } catch (error) {
      console.error('Simple extraction failed, falling back to paginated method:', error);
      return await this.extractPivotData(pivotObject);
    }
  }

  // Format extracted data for easier consumption
  formatPivotData(extractedData) {
    const { data, metadata } = extractedData;
    
    // Create column headers
    const headers = [];
    
    // Add dimension headers
    metadata.dimensions.forEach(dim => {
      headers.push({
        name: dim.qFallbackTitle,
        type: 'dimension',
        index: headers.length,
      });
    });
    
    // Add measure headers
    metadata.measures.forEach(measure => {
      headers.push({
        name: measure.qFallbackTitle,
        type: 'measure',
        index: headers.length,
      });
    });
    
    // Format data rows
    const formattedRows = data.map((row, rowIndex) => {
      const formattedRow = {};
      
      row.forEach((cell, cellIndex) => {
        const header = headers[cellIndex];
        if (header) {
          formattedRow[header.name] = {
            text: cell.qText,
            number: cell.qNum,
            state: cell.qState,
          };
        }
      });
      
      return {
        index: rowIndex,
        data: formattedRow,
      };
    });
    
    return {
      headers,
      rows: formattedRows,
      metadata,
      summary: {
        totalRows: formattedRows.length,
        totalColumns: headers.length,
        dimensions: metadata.dimensions.length,
        measures: metadata.measures.length,
      },
    };
  }

  // Export data to different formats
  exportToCSV(formattedData) {
    const { headers, rows } = formattedData;
    
    // Create CSV header
    const csvHeaders = headers.map(h => h.name).join(',');
    
    // Create CSV rows
    const csvRows = rows.map(row => {
      return headers.map(header => {
        const cellData = row.data[header.name];
        return cellData ? `"${cellData.text}"` : '""';
      }).join(',');
    });
    
    return [csvHeaders, ...csvRows].join('\n');
  }

  // Monitor pivot object for changes (real-time updates)
  monitorPivotChanges(pivotObject, callback) {
    console.log('Setting up pivot change monitoring...');
    
    pivotObject.on('changed', () => {
      console.log('Pivot object changed, triggering callback...');
      callback();
    });
    
    pivotObject.on('closed', () => {
      console.log('Pivot object closed');
    });
  }
}

module.exports = PivotExtractor;