class FieldSelector {
  constructor(doc) {
    this.doc = doc;
  }

  // Create a list object for field values to enable selection
  async createFieldListObject(fieldName) {
    const properties = {
      qInfo: {
        qType: `field-list-${fieldName}`,
      },
      qListObjectDef: {
        qDef: {
          qFieldDefs: [fieldName],
        },
        qInitialDataFetch: [{
          qTop: 0,
          qLeft: 0,
          qHeight: 1000, // Fetch up to 1000 values
          qWidth: 1,
        }],
      },
    };

    console.log(`Creating list object for field: ${fieldName}`);
    return await this.doc.createObject(properties);
  }

  // Find the index of a specific value in a field's list object
  async findValueIndex(listObject, targetValue) {
    const layout = await listObject.getLayout();
    const dataPages = layout.qListObject.qDataPages;
    
    console.log(`DEBUG: Looking for value "${targetValue}" in list object`);
    let globalIndex = 0;
    
    for (const page of dataPages) {
      console.log(`DEBUG: Processing page with ${page.qMatrix.length} rows`);
      for (let i = 0; i < page.qMatrix.length; i++) {
        const cell = page.qMatrix[i][0];
        console.log(`DEBUG: Index ${globalIndex}: "${cell.qText}" (elemNumber: ${cell.qElemNumber})`);
        if (cell.qText === targetValue) {
          console.log(`DEBUG: Found "${targetValue}" at globalIndex ${globalIndex}, elemNumber ${cell.qElemNumber}`);
          // Use the element number instead of the row index
          return cell.qElemNumber;
        }
        globalIndex++;
      }
    }
    
    throw new Error(`Value "${targetValue}" not found in field`);
  }

  // Select a specific value in a field
  async selectFieldValue(fieldName, value) {
    try {
      console.log(`Selecting ${fieldName} = ${value}`);
      
      // Create list object for the field
      const listObject = await this.createFieldListObject(fieldName);
      
      // Find the index of the target value
      const valueIndex = await this.findValueIndex(listObject, value);
      console.log(`Found value "${value}" at index ${valueIndex}`);
      
      // Make the selection
      await listObject.selectListObjectValues('/qListObjectDef', [valueIndex], false);
      console.log(`Successfully selected ${fieldName} = ${value}`);
      
      // Clean up the temporary list object
      await this.doc.destroyObject(listObject.id);
      
      return true;
    } catch (error) {
      console.error(`Failed to select ${fieldName} = ${value}:`, error);
      throw error;
    }
  }

  // Alternative method: Direct field selection (if you know the exact field values)
  async selectFieldValueDirect(fieldName, value) {
    try {
      console.log(`Direct selection: ${fieldName} = ${value}`);
      
      // Use the field's selectValues method directly
      const field = await this.doc.getField(fieldName);
      
      // Make the selection - use selectValues with proper toggle mode
      const result = await field.selectValues([{ qText: value }], true, false);
      console.log(`Selection result for ${fieldName}:`, result);
      
      // Verify the selection was applied by getting field info
      try {
        const fieldInfo = await field.getNxInfo();
        console.log(`Field ${fieldName} info after selection:`, fieldInfo);
      } catch (infoError) {
        console.log(`Could not get field info: ${infoError.message}`);
      }
      
      console.log(`Successfully selected ${fieldName} = ${value} (direct method)`);
      return result;
    } catch (error) {
      console.error(`Failed direct selection ${fieldName} = ${value}:`, error);
      // Fall back to list object method
      return await this.selectFieldValue(fieldName, value);
    }
  }

  // Select multiple fields with their values
  async makeSelections(selections) {
    console.log('Making selections:', selections);
    
    // Clear all existing selections first
    console.log('Clearing all existing selections...');
    await this.clearSelections();
    
    for (const { fieldName, value } of selections) {
      try {
        // Use list object method (more reliable than direct field selection)
        console.log(`\n--- Selecting ${fieldName} = ${value} ---`);
        const success = await this.selectFieldValue(fieldName, value);
        
        if (!success) {
          console.error(`❌ Failed to select ${fieldName} = ${value}`);
          throw new Error(`Selection failed for ${fieldName} = ${value}`);
        }
        
        // Wait a moment for the selection to be processed
        await new Promise(resolve => setTimeout(resolve, 200));
        
      } catch (error) {
        console.error(`Failed to select ${fieldName} = ${value}:`, error);
        throw error;
      }
    }
    
    console.log('\nAll selections completed successfully');
    
    // Verify selections by checking field states
    await this.verifyActualSelections(selections);
    
    return true;
  }

  // Verify that selections were actually applied by checking document selection state
  async verifyActualSelections(selections) {
    console.log('=== VERIFYING ACTUAL SELECTIONS ===');
    
    try {
      // Get the document's selection state
      const selectionObject = await this.doc.createObject({
        qInfo: { qType: 'selection-verification' },
        qSelectionObjectDef: {}
      });
      
      const layout = await selectionObject.getLayout();
      console.log('Current selection state:', JSON.stringify(layout.qSelectionObject, null, 2));
      
      // Clean up
      await this.doc.destroyObject(selectionObject.id);
      
      // Alternative: Check each field individually
      for (const { fieldName, value } of selections) {
        try {
          const field = await this.doc.getField(fieldName);
          
          // Try to get the field's selected values
          const fieldObject = await this.doc.createObject({
            qInfo: { qType: `verify-${fieldName}` },
            qListObjectDef: {
              qDef: { qFieldDefs: [fieldName] },
              qShowAlternatives: true,
              qInitialDataFetch: [{ qTop: 0, qLeft: 0, qHeight: 100, qWidth: 1 }]
            }
          });
          
          const fieldLayout = await fieldObject.getLayout();
          console.log(`\nField "${fieldName}" selection state:`);
          
          if (fieldLayout.qListObject && fieldLayout.qListObject.qDataPages) {
            let selectedCount = 0;
            let selectedValues = [];
            
            fieldLayout.qListObject.qDataPages.forEach(page => {
              page.qMatrix.forEach(row => {
                const cell = row[0];
                if (cell.qState === 'S') { // S = Selected
                  selectedCount++;
                  selectedValues.push(cell.qText);
                }
              });
            });
            
            console.log(`  Selected count: ${selectedCount}`);
            console.log(`  Selected values: [${selectedValues.join(', ')}]`);
            
            if (selectedCount === 0) {
              console.log(`  ❌ No values selected in field "${fieldName}"`);
            } else if (selectedCount === 1 && selectedValues.includes(value)) {
              console.log(`  ✅ Correct selection: "${value}"`);
            } else {
              console.log(`  ⚠️  Unexpected selection state`);
            }
          }
          
          // Clean up
          await this.doc.destroyObject(fieldObject.id);
          
        } catch (fieldError) {
          console.error(`Failed to verify field "${fieldName}":`, fieldError.message);
        }
      }
      
    } catch (error) {
      console.error('Failed to verify selections:', error.message);
    }
  }

  // Clear all selections
  async clearSelections() {
    try {
      console.log('Clearing all selections...');
      await this.doc.clearAll();
      console.log('All selections cleared');
      return true;
    } catch (error) {
      console.error('Failed to clear selections:', error);
      throw error;
    }
  }

  // Get current selections state
  async getCurrentSelections() {
    try {
      const selectionState = await this.doc.getSelectionState();
      return selectionState;
    } catch (error) {
      console.error('Failed to get current selections:', error);
      throw error;
    }
  }

  // Verify selections were applied correctly
  async verifySelections(expectedSelections) {
    try {
      const currentSelections = await this.getCurrentSelections();
      
      console.log('Current selections state:');
      if (currentSelections && currentSelections.qSelections && currentSelections.qSelections.length > 0) {
        currentSelections.qSelections.forEach(selection => {
          console.log(`- ${selection.qField}: ${selection.qSelected}`);
        });
      } else {
        console.log('- No active selections');
      }
      
      return currentSelections;
    } catch (error) {
      console.error('Failed to verify selections:', error);
      // Don't throw error for verification - it's not critical
      console.log('Continuing without selection verification');
      return null;
    }
  }
}

module.exports = FieldSelector;