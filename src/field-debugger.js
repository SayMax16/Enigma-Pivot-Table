class FieldDebugger {
  constructor(doc) {
    this.doc = doc;
  }

  // Get all available fields in the document
  async getAllFields() {
    try {
      console.log('=== DEBUGGING: Getting all available fields ===');
      
      const properties = {
        qInfo: {
          qType: 'field-list-debug',
        },
        qFieldListDef: {},
      };

      const fieldListObject = await this.doc.createObject(properties);
      const layout = await fieldListObject.getLayout();
      
      console.log('Available fields in the document:');
      if (layout.qFieldList && layout.qFieldList.qItems) {
        layout.qFieldList.qItems.forEach((field, index) => {
          console.log(`${index + 1}. "${field.qName}" (${field.qCardinal} values)`);
        });
      }
      
      // Clean up
      await this.doc.destroyObject(fieldListObject.id);
      
      return layout.qFieldList ? layout.qFieldList.qItems : [];
    } catch (error) {
      console.error('Failed to get field list:', error);
      return [];
    }
  }

  // Get field values for a specific field
  async getFieldValues(fieldName, maxValues = 20) {
    try {
      console.log(`=== DEBUGGING: Getting values for field "${fieldName}" ===`);
      
      const properties = {
        qInfo: {
          qType: `field-values-${fieldName}`,
        },
        qListObjectDef: {
          qDef: {
            qFieldDefs: [fieldName],
          },
          qInitialDataFetch: [{
            qTop: 0,
            qLeft: 0,
            qHeight: maxValues,
            qWidth: 1,
          }],
        },
      };

      const listObject = await this.doc.createObject(properties);
      const layout = await listObject.getLayout();
      
      console.log(`Field "${fieldName}" values:`);
      if (layout.qListObject && layout.qListObject.qDataPages) {
        layout.qListObject.qDataPages.forEach(page => {
          page.qMatrix.forEach((row, index) => {
            const cell = row[0];
            console.log(`${index + 1}. "${cell.qText}" (state: ${cell.qState})`);
          });
        });
      }
      
      // Clean up
      await this.doc.destroyObject(listObject.id);
      
      return layout.qListObject;
    } catch (error) {
      console.error(`Failed to get values for field "${fieldName}":`, error);
      return null;
    }
  }

  // Search for fields containing specific text
  async findFieldsContaining(searchText) {
    const fields = await this.getAllFields();
    const matches = fields.filter(field => 
      field.qName.toLowerCase().includes(searchText.toLowerCase())
    );
    
    console.log(`=== Fields containing "${searchText}" ===`);
    matches.forEach(field => {
      console.log(`- "${field.qName}"`);
    });
    
    return matches;
  }

  // Debug field selection
  async debugFieldSelection(fieldName, targetValue) {
    console.log(`=== DEBUGGING FIELD SELECTION ===`);
    console.log(`Field: "${fieldName}"`);
    console.log(`Target value: "${targetValue}"`);
    
    // Check if field exists
    const fields = await this.getAllFields();
    const fieldExists = fields.find(f => f.qName === fieldName);
    
    if (!fieldExists) {
      console.log(`❌ Field "${fieldName}" not found!`);
      
      // Look for similar fields
      const similar = fields.filter(f => 
        f.qName.includes('Год') || 
        f.qName.includes('Месяц') || 
        f.qName.includes('Month') ||
        f.qName.includes('Year')
      );
      
      if (similar.length > 0) {
        console.log('🔍 Similar fields found:');
        similar.forEach(field => {
          console.log(`  - "${field.qName}"`);
        });
      }
      
      return false;
    }
    
    console.log(`✅ Field "${fieldName}" exists (${fieldExists.qCardinal} values)`);
    
    // Get field values
    const fieldValues = await this.getFieldValues(fieldName, 50);
    
    // Check if target value exists
    if (fieldValues && fieldValues.qDataPages) {
      let valueFound = false;
      fieldValues.qDataPages.forEach(page => {
        page.qMatrix.forEach(row => {
          if (row[0].qText === targetValue) {
            valueFound = true;
            console.log(`✅ Value "${targetValue}" found in field`);
          }
        });
      });
      
      if (!valueFound) {
        console.log(`❌ Value "${targetValue}" not found in field`);
        console.log('Available values (first 20):');
        fieldValues.qDataPages.forEach(page => {
          page.qMatrix.slice(0, 20).forEach((row, index) => {
            console.log(`  ${index + 1}. "${row[0].qText}"`);
          });
        });
      }
      
      return valueFound;
    }
    
    return false;
  }
}

module.exports = FieldDebugger;