class ContainerExtractor {
  constructor(doc) {
    this.doc = doc;
  }

  // Get container object by ID
  async getContainer(containerId) {
    try {
      console.log(`Getting container with ID: ${containerId}`);
      const container = await this.doc.getObject(containerId);
      console.log('Container retrieved successfully');
      return container;
    } catch (error) {
      console.error(`Failed to get container ${containerId}:`, error);
      throw error;
    }
  }

  // Get container layout and analyze its structure
  async analyzeContainer(container) {
    try {
      console.log('Analyzing container structure...');
      const layout = await container.getLayout();
      
      console.log('Container info:');
      console.log(`- Type: ${layout.qInfo.qType}`);
      console.log(`- Title: ${layout.title || layout.qMeta?.title || 'No title'}`);
      
      // Look for child objects in the container
      if (layout.qChildList && layout.qChildList.qItems) {
        console.log(`- Child objects: ${layout.qChildList.qItems.length}`);
        
        layout.qChildList.qItems.forEach((child, index) => {
          console.log(`  ${index + 1}. ID: "${child.qInfo.qId}", Type: "${child.qInfo.qType}", Title: "${child.qData?.title || child.qMeta?.title || 'No title'}"`);
        });
      }
      
      return layout;
    } catch (error) {
      console.error('Failed to analyze container:', error);
      throw error;
    }
  }

  // Find and get a specific child object within the container
  async getChildObject(container, childObjectId) {
    try {
      console.log(`Looking for child object: ${childObjectId}`);
      
      // First, analyze the container to see its children
      const layout = await this.analyzeContainer(container);
      
      // Check if the child object exists in the container
      let childFound = false;
      if (layout.qChildList && layout.qChildList.qItems) {
        childFound = layout.qChildList.qItems.some(child => child.qInfo.qId === childObjectId);
      }
      
      if (childFound) {
        console.log(`✅ Child object "${childObjectId}" found in container`);
      } else {
        console.log(`⚠️ Child object "${childObjectId}" not found in container children list`);
        console.log('Attempting direct access...');
      }
      
      // Try to get the child object directly from the document
      const childObject = await this.doc.getObject(childObjectId);
      console.log(`✅ Successfully retrieved child object: ${childObjectId}`);
      
      return childObject;
    } catch (error) {
      console.error(`Failed to get child object ${childObjectId}:`, error);
      throw error;
    }
  }

  // Get all child objects from the container
  async getAllChildObjects(container) {
    try {
      const layout = await this.analyzeContainer(container);
      const childObjects = [];
      
      if (layout.qChildList && layout.qChildList.qItems) {
        console.log(`Retrieving ${layout.qChildList.qItems.length} child objects...`);
        
        for (const child of layout.qChildList.qItems) {
          try {
            const childObj = await this.doc.getObject(child.qInfo.qId);
            childObjects.push({
              id: child.qInfo.qId,
              type: child.qInfo.qType,
              object: childObj
            });
            console.log(`✅ Retrieved: ${child.qInfo.qId} (${child.qInfo.qType})`);
          } catch (error) {
            console.log(`❌ Failed to retrieve: ${child.qInfo.qId} - ${error.message}`);
          }
        }
      }
      
      return childObjects;
    } catch (error) {
      console.error('Failed to get child objects:', error);
      throw error;
    }
  }

  // Specialized method to extract pivot table from container
  async extractPivotFromContainer(containerId, pivotObjectId) {
    try {
      console.log('=== CONTAINER-BASED PIVOT EXTRACTION ===');
      console.log(`Container ID: ${containerId}`);
      console.log(`Target Pivot ID: ${pivotObjectId}`);
      
      // Step 1: Get the container
      const container = await this.getContainer(containerId);
      
      // Step 2: Analyze container structure
      await this.analyzeContainer(container);
      
      // Step 3: Get the specific pivot table object
      const pivotObject = await this.getChildObject(container, pivotObjectId);
      
      // Step 4: Verify it's a pivot table
      const pivotLayout = await pivotObject.getLayout();
      console.log('\nPivot object details:');
      console.log(`- Type: ${pivotLayout.qInfo.qType}`);
      console.log(`- Title: ${pivotLayout.title || pivotLayout.qMeta?.title || 'No title'}`);
      
      if (pivotLayout.qHyperCube) {
        console.log('- Structure: HyperCube detected');
        console.log(`- Dimensions: ${pivotLayout.qHyperCube.qDimensionInfo?.length || 0}`);
        console.log(`- Measures: ${pivotLayout.qHyperCube.qMeasureInfo?.length || 0}`);
        console.log(`- Data mode: ${pivotLayout.qHyperCube.qMode || 'N/A'}`);
        console.log(`- Total rows: ${pivotLayout.qHyperCube.qSize?.qcy || 'N/A'}`);
      }
      
      if (pivotLayout.qPivotTable) {
        console.log('- Structure: Native pivot table detected');
      }
      
      console.log('✅ Container-based pivot extraction setup completed');
      return pivotObject;
      
    } catch (error) {
      console.error('Container-based pivot extraction failed:', error);
      throw error;
    }
  }
}

module.exports = ContainerExtractor;