/**
 * PROJECT SETUP - Creation and Folder Management
 * Handles new project creation, folder structure, and asset tracker setup
 */

// ============================================================================
// MAIN PROJECT CREATION FUNCTION
// ============================================================================

function createNewProject(setupData) {
  try {
    // Validate required fields
    const required = ['projectNumber', 'projectName', 'client', 'description', 'emailNotifications', 'projectManagers'];
    for (const field of required) {
      if (!setupData[field] || setupData[field].toString().trim() === '') {
        return { 
          success: false, 
          message: `Required field missing: ${field}`
        };
      }
    }
    
    // Get HQ configuration
    const hqConfig = getHQConfig();
    const parentFolderId = hqConfig.PARENT_FOLDER_ID;
    const defaultEmail = hqConfig.DEFAULT_EMAIL;
    
    if (!parentFolderId) {
      return { 
        success: false, 
        message: 'Parent folder not configured. Check HQ Config sheet.'
      };
    }
    
    // Combine emails
    const projectEmails = setupData.emailNotifications.trim();
    const combinedEmails = defaultEmail + (projectEmails ? ', ' + projectEmails : '');
    setupData.alertEmails = combinedEmails;
    
    // Create folder structure
    const folderResult = createProjectFolderStructure(setupData, parentFolderId);
    if (!folderResult.success) {
      return folderResult;
    }
    
    // Create Asset Tracker
    const trackerResult = createAssetTracker(setupData, folderResult);
    if (!trackerResult.success) {
      return trackerResult;
    }
    
    // Add to Projects List sheet
    const projectSheetData = {
      projectNumber: setupData.projectNumber,
      projectName: setupData.projectName,
      client: setupData.client,
      description: setupData.description,
      projectManagers: Array.isArray(setupData.projectManagers) 
        ? setupData.projectManagers.join(', ') 
        : setupData.projectManagers,
      emailNotifications: combinedEmails,
      startDate: setupData.startDate,
      inHandsDate: setupData.inHandsDate,
      assetTrackerUrl: trackerResult.trackerUrl
    };
    
    addProjectToSheet(projectSheetData);
    
    return {
      success: true,
      message: 'Project created successfully!',
      projectNumber: setupData.projectNumber,
      assetTrackerUrl: trackerResult.trackerUrl
    };
    
  } catch (error) {
    console.error('Error creating project:', error);
    return { 
      success: false, 
      message: 'Error: ' + error.toString()
    };
  }
}

// ============================================================================
// FOLDER STRUCTURE CREATION
// ============================================================================

function createProjectFolderStructure(setupData, parentFolderId) {
  try {
    const parentFolder = DriveApp.getFolderById(parentFolderId);
    
    // Create main project folder
    const projectFolderName = setupData.projectNumber;
    const projectFolder = parentFolder.createFolder(projectFolderName);
    const projectFolderId = projectFolder.getId();
    
    // Create subfolder structure
    createSubfolders(projectFolder, setupData.projectName);
    
    return {
      success: true,
      projectFolder: projectFolder,
      projectFolderId: projectFolderId
    };
    
  } catch (error) {
    console.error('Error creating folder structure:', error);
    return { 
      success: false, 
      message: `Folder creation failed: ${error.toString()}`
    };
  }
}

function createSubfolders(projectFolder, projectName) {
  // Level 1 folders
  projectFolder.createFolder('01 - Admin Docs');
  const productionFiles = projectFolder.createFolder('02 - Production Files');
  const projectFiles = projectFolder.createFolder('03 - Project Files');
  projectFolder.createFolder('04 - Vendor Docs');
  
  // Level 2 folders under 02 - Production Files
  productionFiles.createFolder(`0 - ${projectName} Artwork Files`);
  productionFiles.createFolder('On Hold');
  
  // Level 2 folders under 03 - Project Files
  projectFiles.createFolder('Sections');
  projectFiles.createFolder(`Team Docs - ${projectName}`);
}

// ============================================================================
// ASSET TRACKER CREATION
// ============================================================================

function createAssetTracker(setupData, folderResult) {
  try {
    // Create new spreadsheet with proper OAuth scope
    const trackerName = `Production - ${setupData.projectNumber}`;
    const tracker = SpreadsheetApp.create(trackerName);
    const trackerId = tracker.getId();
    const trackerUrl = tracker.getUrl();
    
    // Move to project folder
    const trackerFile = DriveApp.getFileById(trackerId);
    trackerFile.moveTo(folderResult.projectFolder);
    
    // Set up the tracker
    setupAssetTrackerSheets(tracker, setupData, folderResult);
    
    return {
      success: true,
      trackerId: trackerId,
      trackerUrl: trackerUrl
    };
    
  } catch (error) {
    console.error('Error creating asset tracker:', error);
    return {
      success: false,
      message: `Asset tracker creation failed: ${error.toString()}`
    };
  }
}

function setupAssetTrackerSheets(tracker, setupData, folderResult) {
  // Rename first sheet to Master
  const mainSheet = tracker.getSheets()[0];
  mainSheet.setName(`${setupData.projectNumber} Master`);
  
  // Create headers
  const headers = [
    'ID', 'Area', 'Asset', 'Status', 'Dimensions', 'Quantity', 'Item', 'Material',
    'Due Date', 'Strike Date', 'Venue', 'Location', 'Artwork', 'Image Link',
    'Double Sided', 'Diecut', 'Production Status', 'Edit'
  ];
  
  mainSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  mainSheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  
  mainSheet.setFrozenRows(1);
  
  // Set column widths
  mainSheet.setColumnWidth(1, 80);   // ID
  mainSheet.setColumnWidth(2, 100);  // Area
  mainSheet.setColumnWidth(3, 200);  // Asset
  mainSheet.setColumnWidth(4, 120);  // Status
  mainSheet.setColumnWidth(5, 120);  // Dimensions
  mainSheet.setColumnWidth(13, 200); // Artwork
  mainSheet.setColumnWidth(14, 200); // Image Link
  
  // Create ProjectConfig sheet (HIDDEN)
  createProjectConfigSheet(tracker, setupData, folderResult);
  
  // Create MaterialIDMap sheet (starts empty)
  const materialSheet = tracker.insertSheet('MaterialIDMap');
  materialSheet.getRange(1, 1, 1, 2).setValues([['Material', 'ID Prefix']]);
  materialSheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e8f0fe');
  materialSheet.hideSheet();
  
  // Create AssetLog sheet
  const logSheet = tracker.insertSheet('AssetLog');
  logSheet.getRange(1, 1, 1, 4).setValues([['LogID', 'ProjectRow', 'Timestamp', 'FormData']]);
  logSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#e8f0fe');
  logSheet.hideSheet();
  
  // Create dropdown lists with initial values
  createDropdownSheets(tracker);
  
  // Install complete Asset Tracker code
  installAssetTrackerCode(tracker, setupData, folderResult);
  
  SpreadsheetApp.flush();
}

function createProjectConfigSheet(tracker, setupData, folderResult) {
  const configSheet = tracker.insertSheet('ProjectConfig', 0);
  
  const configData = [
    ['PROJECT CONFIGURATION', 'Value', 'Description'],
    ['', '', ''],
    ['Basic Information', '', ''],
    ['Project Name', setupData.projectName, 'Full name of the project'],
    ['Project Code', setupData.projectNumber, 'Project number/code'],
    ['Client Name', setupData.client, 'Client or company name'],
    ['', '', ''],
    ['Email Alerts', '', ''],
    ['Alert Recipients', setupData.alertEmails, 'Email addresses for alerts'],
    ['', '', ''],
    ['Google Drive Folders', '', ''],
    ['Main Folder ID', folderResult.projectFolderId, 'Google Drive folder ID'],
    ['Artwork Folder', `0 - ${setupData.projectName} Artwork Files`, 'Folder for artwork files'],
    ['Production Folder', '02 - Production Files', 'Production files folder'],
    ['Team Docs Folder', `Team Docs - ${setupData.projectName}`, 'Team documentation folder'],
    ['', '', ''],
    ['Sheet Configuration', '', ''],
    ['Master Sheet Name', `${setupData.projectNumber} Master`, 'Main asset tracking sheet'],
    ['Log Sheet Name', 'AssetLog', 'Change log sheet'],
    ['Material ID Map Sheet', 'MaterialIDMap', 'Material ID mapping sheet']
  ];
  
  configSheet.getRange(1, 1, configData.length, 3).setValues(configData);
  
  configSheet.setColumnWidth(1, 200);
  configSheet.setColumnWidth(2, 300);
  configSheet.setColumnWidth(3, 350);
  
  configSheet.getRange('A1:C1')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12);
  
  const sectionRows = [3, 8, 11, 17];
  sectionRows.forEach(row => {
    configSheet.getRange(row, 1, 1, 3).setBackground('#e8f0fe').setFontWeight('bold');
  });
  
  configSheet.getRange('B2:B30').setBackground('#f8f9fa');
  
  // HIDE the ProjectConfig sheet
  configSheet.hideSheet();
}

function createDropdownSheets(tracker) {
  const dropdownData = {
    'ItemsList': ['Banner', 'Sign', 'Poster', 'Decal', 'Display', 'Billboard'],
    'MaterialsList': ['Adhesive Vinyl - Matte', 'Foamcore - 1/4"', 'Foamcore - 1/2"', 'Gatorplast - 1/4"', 'Gatorplast - 1/2"'],
    'StatusesList': ['New Asset', 'In Progress', 'Awaiting Approval', 'Approved', 'In Production', 'Delivered', 'On Hold', 'Requires Attn'],
    'VenuesList': ['Main Hall', 'Conference Room', 'Lobby', 'Outdoor'],
    'AreasList': ['North', 'South', 'East', 'West', 'Central'],
    'ProductionStatusesList': ['Processing', 'Printing', 'Cutting', 'Finishing', 'Ready', 'Picked Up']
  };
  
  Object.entries(dropdownData).forEach(([sheetName, values]) => {
    const sheet = tracker.insertSheet(sheetName);
    sheet.getRange(1, 1).setValue(sheetName.replace('List', ''));
    sheet.getRange(1, 1).setFontWeight('bold').setBackground('#e8f0fe');
    
    if (values && values.length > 0) {
      const data = values.map(v => [v]);
      sheet.getRange(2, 1, data.length, 1).setValues(data);
    }
    
    sheet.hideSheet();
  });
}

function installAssetTrackerCode(tracker, setupData, folderResult) {
  try {
    // Get the spreadsheet ID
    const trackerId = tracker.getId();
    
    // Read the template Asset Tracker code
    const codeTemplate = getAssetTrackerCodeTemplate();
    const assetFormHtml = getAssetFormHtmlTemplate();
    const commentDialogHtml = getCommentDialogHtmlTemplate();
    const dropdownEditorHtml = getDropdownEditorHtmlTemplate();
    const reorderAssetFormHtml = getReorderAssetFormHtmlTemplate();
    
    // Configure the code with project-specific values
    const artworkFolderName = `0 - ${setupData.projectName} Artwork Files`;
    const masterSheetName = `${setupData.projectNumber} Master`;
    
    const configuredCode = codeTemplate
      .replace(/\{\{MASTER_SHEET_NAME\}\}/g, masterSheetName)
      .replace(/\{\{DRIVE_FOLDER_ID\}\}/g, folderResult.projectFolderId)
      .replace(/\{\{PROJECT_CODE\}\}/g, setupData.projectNumber)
      .replace(/\{\{PROJECT_NAME\}\}/g, setupData.projectName)
      .replace(/\{\{ARTWORK_FOLDER_NAME\}\}/g, artworkFolderName)
      .replace(/\{\{ALERT_EMAILS\}\}/g, setupData.alertEmails);
    
    // Share the artwork folder with alert recipients
    shareArtworkFolder(folderResult.projectFolder, artworkFolderName, setupData.alertEmails);
    
    // Create installation instructions with actual code
    const instructionSheet = tracker.insertSheet('📌 CODE INSTALLATION');
    
    const instructions = [
      ['ASSET TRACKER CODE - INSTALLATION INSTRUCTIONS'],
      [''],
      ['This Asset Tracker has been pre-configured for your project.'],
      ['All configuration values have been automatically set.'],
      [''],
      ['📋 INSTALLATION STEPS:'],
      [''],
      ['1. Open Extensions → Apps Script in this spreadsheet'],
      ['2. Delete the existing Code.gs file content'],
      ['3. Copy ALL the code from the "Code.gs Template" sheet below'],
      ['4. Paste it into Code.gs in Apps Script'],
      ['5. Create these HTML files in Apps Script (Files → + → HTML):'],
      ['   • AssetForm.html - Copy from "AssetForm" sheet'],
      ['   • CommentDialog.html - Copy from "CommentDialog" sheet'],
      ['   • DropdownEditor.html - Copy from "DropdownEditor" sheet'],
      ['   • ReorderAssetForm.html - Copy from "ReorderAssetForm" sheet'],
      ['6. Save all files (Ctrl+S or Cmd+S)'],
      ['7. Refresh this spreadsheet'],
      ['8. You will see the menu: ' + setupData.projectName],
      ['9. Delete this instruction sheet and the template sheets'],
      [''],
      ['✅ ALREADY CONFIGURED:'],
      [`   • Project Name: ${setupData.projectName}`],
      [`   • Project Code: ${setupData.projectNumber}`],
      [`   • Master Sheet: ${masterSheetName}`],
      [`   • Drive Folder: ${artworkFolderName}`],
      [`   • Alert Emails: ${setupData.alertEmails}`],
      [''],
      ['All dropdown lists start empty - you can add values as you work.'],
      ['MaterialIDMap will auto-assign letter prefixes as you add materials.']
    ];
    
    instructionSheet.getRange(1, 1, instructions.length, 1).setValues(instructions);
    instructionSheet.setColumnWidth(1, 800);
    instructionSheet.getRange('A1').setFontSize(14).setFontWeight('bold').setBackground('#fef3c7');
    instructionSheet.getRange('A6').setFontWeight('bold').setFontSize(12);
    instructionSheet.getRange('A3:A4').setFontColor('#059669');
    instructionSheet.getRange('A21').setFontWeight('bold').setFontSize(12).setBackground('#dcfce7');
    
    // Create template sheets with the actual code
    createCodeTemplateSheet(tracker, configuredCode);
    createHtmlTemplateSheet(tracker, 'AssetForm', assetFormHtml);
    createHtmlTemplateSheet(tracker, 'CommentDialog', commentDialogHtml);
    createHtmlTemplateSheet(tracker, 'DropdownEditor', dropdownEditorHtml);
    createHtmlTemplateSheet(tracker, 'ReorderAssetForm', reorderAssetFormHtml);
    
    // Make instruction sheet active
    tracker.setActiveSheet(instructionSheet);
    
    SpreadsheetApp.flush();
    
  } catch (error) {
    console.error('Error in installAssetTrackerCode:', error);
  }
}

function shareArtworkFolder(projectFolder, artworkFolderName, emailList) {
  try {
    const artworkFolders = projectFolder.getFoldersByName(artworkFolderName);
    if (artworkFolders.hasNext()) {
      const artworkFolder = artworkFolders.next();
      const emails = emailList.split(',').map(e => e.trim());
      
      emails.forEach(email => {
        try {
          artworkFolder.addEditor(email);
        } catch (e) {
          console.error(`Error sharing folder with ${email}:`, e);
        }
      });
    }
  } catch (error) {
    console.error('Error sharing artwork folder:', error);
  }
}

function createCodeTemplateSheet(tracker, code) {
  const sheet = tracker.insertSheet('Code.gs Template');
  
  // Split code into chunks that fit in cells (max 50000 characters per cell)
  const maxChunkSize = 40000;
  const chunks = [];
  
  for (let i = 0; i < code.length; i += maxChunkSize) {
    chunks.push([code.substring(i, i + maxChunkSize)]);
  }
  
  sheet.getRange(1, 1, chunks.length, 1).setValues(chunks);
  sheet.setColumnWidth(1, 1000);
  sheet.getRange(1, 1, chunks.length, 1).setWrap(true);
  
  // Add note at top
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1).setValue('⬇️ COPY ALL THE CODE BELOW (SELECT ALL CELLS AND COPY)');
  sheet.getRange(1, 1).setBackground('#fef3c7').setFontWeight('bold');
  
  sheet.hideSheet();
}

function createHtmlTemplateSheet(tracker, name, htmlContent) {
  const sheet = tracker.insertSheet(`${name} Template`);
  
  const maxChunkSize = 40000;
  const chunks = [];
  
  for (let i = 0; i < htmlContent.length; i += maxChunkSize) {
    chunks.push([htmlContent.substring(i, i + maxChunkSize)]);
  }
  
  sheet.getRange(1, 1, chunks.length, 1).setValues(chunks);
  sheet.setColumnWidth(1, 1000);
  sheet.getRange(1, 1, chunks.length, 1).setWrap(true);
  
  sheet.insertRowBefore(1);
  sheet.getRange(1, 1).setValue(`⬇️ COPY ALL THE HTML BELOW FOR ${name}`);
  sheet.getRange(1, 1).setBackground('#fef3c7').setFontWeight('bold');
  
  sheet.hideSheet();
}

function getAssetTrackerCodeTemplate() {
  // Return the complete Asset Tracker code as a string
  // This will be the contents of AssetTracker_Code.gs
  return `/**
 * @OnlyCurrentDoc
 * Asset Management System - Auto-configured by Projects HQ
 */

// Configuration - AUTO-POPULATED BY PROJECTS HQ
const CONFIG = {
  LOG_SHEET_NAME: 'AssetLog',
  LOG_ID_PREFIX: 'ASSET',
  MASTER_SHEET_NAME: '{{MASTER_SHEET_NAME}}',
  MATERIAL_ID_MAP_SHEET: 'MaterialIDMap',
  DRIVE_FOLDER_ID: '{{DRIVE_FOLDER_ID}}',
  PROJECT_CODE: '{{PROJECT_CODE}}',
  PROJECT_NAME: '{{PROJECT_NAME}}',
  ARTWORK_FOLDER_NAME: '{{ARTWORK_FOLDER_NAME}}',
  ALERT_EMAILS: '{{ALERT_EMAILS}}',
  
  COLUMN_MAP: {
    ID: 1, AREA: 2, ASSET: 3, STATUS: 4, DIMENSIONS: 5, QUANTITY: 6,
    ITEM: 7, MATERIAL: 8, DUE_DATE: 9, STRIKE_DATE: 10, VENUE: 11,
    LOCATION: 12, ARTWORK: 13, IMAGE_LINK: 14, DOUBLE_SIDED: 15, DIECUT: 16, 
    PRODUCTION_STATUS: 17, EDIT: 18
  },
  
  DROPDOWN_SHEETS: {
    items: 'ItemsList',
    statuses: 'StatusesList',
    venues: 'VenuesList',
    areas: 'AreasList',
    productionStatuses: 'ProductionStatusesList'
  }
};
` + '// ... [REST OF CODE WILL BE ADDED IN NEXT MESSAGE - THIS IS A PLACEHOLDER]';
}

// Template getter functions for Asset Tracker code
function getAssetTrackerCodeTemplate() {
  // This reads the actual AssetTracker_Code.gs file content
  // In production, this would contain the full code
  // For now, returning a placeholder that indicates where to insert the actual code
  return '/**
 * @OnlyCurrentDoc
 * Asset Management System with Google Drive Integration and Two-Way Sync
 * 
 * CONFIGURATION INSTRUCTIONS:
 * This file will be auto-configured by Projects HQ during project creation.
 * The CONFIG object below will be populated with project-specific values.
 */

// Configuration - AUTO-POPULATED BY PROJECTS HQ
const CONFIG = {
  LOG_SHEET_NAME: 'AssetLog',
  LOG_ID_PREFIX: 'ASSET',
  MASTER_SHEET_NAME: '{{MASTER_SHEET_NAME}}',  // Will be replaced with: ProjectNumber Master
  MATERIAL_ID_MAP_SHEET: 'MaterialIDMap',
  DRIVE_FOLDER_ID: '{{DRIVE_FOLDER_ID}}',  // Will be replaced with project folder ID
  PROJECT_CODE: '{{PROJECT_CODE}}',  // Will be replaced with project number
  PROJECT_NAME: '{{PROJECT_NAME}}',  // Will be replaced with project name
  ARTWORK_FOLDER_NAME: '{{ARTWORK_FOLDER_NAME}}',  // Will be replaced with: 0 - ProjectName Artwork Files
  ALERT_EMAILS: '{{ALERT_EMAILS}}',  // Will be replaced with comma-separated email list
  
  COLUMN_MAP: {
    ID: 1, AREA: 2, ASSET: 3, STATUS: 4, DIMENSIONS: 5, QUANTITY: 6,
    ITEM: 7, MATERIAL: 8, DUE_DATE: 9, STRIKE_DATE: 10, VENUE: 11,
    LOCATION: 12, ARTWORK: 13, IMAGE_LINK: 14, DOUBLE_SIDED: 15, DIECUT: 16, 
    PRODUCTION_STATUS: 17, EDIT: 18
  },
  
  DROPDOWN_SHEETS: {
    items: 'ItemsList',
    statuses: 'StatusesList',
    venues: 'VenuesList',
    areas: 'AreasList',
    productionStatuses: 'ProductionStatusesList'
  }
};

const assetApp = {
  showDialog: function() {
    const htmlOutput = HtmlService.createHtmlOutputFromFile('AssetForm')
        .setWidth(600).setHeight(900);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, `${CONFIG.PROJECT_NAME} : Add New Asset`);
  },

  openForEdit: function(rowNumber) {
    const formData = projectSheet.getRowData(rowNumber);
    if (formData) {
      const htmlOutput = HtmlService.createHtmlOutputFromFile('AssetForm')
          .setWidth(600).setHeight(900);
      const htmlContent = htmlOutput.getContent();
      const modifiedContent = htmlContent.replace('<script>', `<script>window.editFormData = ${JSON.stringify(formData)};`);
      const modifiedOutput = HtmlService.createHtmlOutput(modifiedContent).setWidth(600).setHeight(900);
      SpreadsheetApp.getUi().showModalDialog(modifiedOutput, 'Edit Asset');
    } else {
      SpreadsheetApp.getUi().alert('Error', 'Could not load row data.', SpreadsheetApp.getUi().ButtonSet.OK);
    }
  },

  addToProject: function(assetData) {
    try {
      return projectSheet.addProjectItem(assetData, CONFIG.LOG_ID_PREFIX, CONFIG.LOG_SHEET_NAME);
    } catch (e) {
      console.error("Error in assetApp.addToProject: " + e.toString());
      return { success: false, message: `Error adding to project: ${e.toString()}`, rowNumber: null, logId: null };
    }
  }
};

const dropdownManager = {
  getDropdownValues: function(fieldName) {
    // Special handling for materials - read from MaterialIDMap
    if (fieldName === 'materials') {
      return materialIDManager.getAllMaterials();
    }
    
    const sheetName = CONFIG.DROPDOWN_SHEETS[fieldName];
    if (!sheetName) return [];
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      sheet.getRange(1, 1).setValue(fieldName.charAt(0).toUpperCase() + fieldName.slice(1));
      sheet.hideSheet();
    }
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    return values.filter(row => row[0] && row[0].toString().trim() !== "").map(row => row[0].toString().trim());
  },

  addDropdownValue: function(fieldName, newValue) {
    // Special handling for materials - add to MaterialIDMap
    if (fieldName === 'materials') {
      const existingMaterials = materialIDManager.getAllMaterials();
      if (existingMaterials.includes(newValue)) {
        return { success: false, message: 'Material already exists' };
      }
      materialIDManager.assignIDToMaterial(newValue);
      return { success: true, message: 'Material added successfully', value: newValue };
    }
    
    const sheetName = CONFIG.DROPDOWN_SHEETS[fieldName];
    if (!sheetName) return { success: false, message: 'Invalid field name' };
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(sheetName);
      sheet.getRange(1, 1).setValue(fieldName.charAt(0).toUpperCase() + fieldName.slice(1));
      sheet.hideSheet();
    }
    const existingValues = this.getDropdownValues(fieldName);
    if (existingValues.includes(newValue)) return { success: false, message: 'Value already exists' };
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1).setValue(newValue);
    return { success: true, message: 'Value added successfully', value: newValue };
  },

  updateDropdownValue: function(fieldName, oldValue, newValue) {
    // Special handling for materials - update in MaterialIDMap
    if (fieldName === 'materials') {
      materialIDManager.updateMaterialName(oldValue, newValue);
      this.updateInMainSheet(fieldName, oldValue, newValue);
      return { success: true, message: 'Material updated successfully' };
    }
    
    const sheetName = CONFIG.DROPDOWN_SHEETS[fieldName];
    if (!sheetName) return { success: false, message: 'Invalid field name' };
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Sheet not found' };
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'No values to update' };
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] === oldValue) {
        sheet.getRange(i + 2, 1).setValue(newValue);
        this.updateInMainSheet(fieldName, oldValue, newValue);
        return { success: true, message: 'Value updated successfully' };
      }
    }
    return { success: false, message: 'Value not found' };
  },

  deleteDropdownValue: function(fieldName, value) {
    // Special handling for materials - delete from MaterialIDMap
    if (fieldName === 'materials') {
      materialIDManager.deleteMaterial(value);
      return { success: true, message: 'Material deleted successfully' };
    }
    
    const sheetName = CONFIG.DROPDOWN_SHEETS[fieldName];
    if (!sheetName) return { success: false, message: 'Invalid field name' };
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Sheet not found' };
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'No values to delete' };
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] === value) {
        sheet.deleteRow(i + 2);
        return { success: true, message: 'Value deleted successfully' };
      }
    }
    return { success: false, message: 'Value not found' };
  },

  updateInMainSheet: function(fieldName, oldValue, newValue) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const mainSheet = spreadsheet.getSheetByName(CONFIG.MASTER_SHEET_NAME);
    if (!mainSheet) return;
    const columnMap = {
      'items': CONFIG.COLUMN_MAP.ITEM, 'materials': CONFIG.COLUMN_MAP.MATERIAL,
      'statuses': CONFIG.COLUMN_MAP.STATUS, 'venues': CONFIG.COLUMN_MAP.VENUE, 
      'areas': CONFIG.COLUMN_MAP.AREA, 'productionStatuses': CONFIG.COLUMN_MAP.PRODUCTION_STATUS
    };
    const column = columnMap[fieldName];
    if (!column) return;
    const lastRow = mainSheet.getLastRow();
    if (lastRow <= 1) return;
    const range = mainSheet.getRange(2, column, lastRow - 1, 1);
    const values = range.getValues();
    let updated = false;
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] === oldValue) {
        values[i][0] = newValue;
        updated = true;
      }
    }
    if (updated) range.setValues(values);
  },

  getAllDropdownData: function() {
    return {
      items: this.getDropdownValues('items'), 
      materials: this.getDropdownValues('materials'),
      statuses: this.getDropdownValues('statuses'), 
      venues: this.getDropdownValues('venues'), 
      areas: this.getDropdownValues('areas'), 
      productionStatuses: this.getDropdownValues('productionStatuses')
    };
  }
};

const materialIDManager = {
  getOrCreateMaterialIDSheet: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(CONFIG.MATERIAL_ID_MAP_SHEET);
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet(CONFIG.MATERIAL_ID_MAP_SHEET);
      sheet.getRange(1, 1, 1, 2).setValues([['Material', 'ID Prefix']]);
      sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e8f0fe');
      sheet.hideSheet();
    } else {
      if (!sheet.isSheetHidden()) {
        sheet.hideSheet();
      }
    }
    
    return sheet;
  },

  getAllMaterials: function() {
    const sheet = this.getOrCreateMaterialIDSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    return values.filter(row => row[0] && row[0].toString().trim() !== "")
                 .map(row => row[0].toString().trim());
  },

  getMaterialIDMap: function() {
    const sheet = this.getOrCreateMaterialIDSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return {};
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    const map = {};
    values.forEach(row => {
      if (row[0] && row[1]) map[row[0].toString().trim()] = row[1].toString().trim();
    });
    return map;
  },

  getNextAvailableLetter: function() {
    const sheet = this.getOrCreateMaterialIDSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return 'A';
    const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    const usedLetters = values.map(row => row[0]).filter(letter => letter);
    let maxCharCode = 64;
    usedLetters.forEach(letter => {
      const charCode = letter.charCodeAt(0);
      if (charCode > maxCharCode) maxCharCode = charCode;
    });
    return String.fromCharCode(maxCharCode + 1);
  },

  assignIDToMaterial: function(materialName) {
    const sheet = this.getOrCreateMaterialIDSheet();
    const map = this.getMaterialIDMap();
    if (map[materialName]) return map[materialName];
    const nextLetter = this.getNextAvailableLetter();
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, 1, 2).setValues([[materialName, nextLetter]]);
    return nextLetter;
  },

  getMaterialPrefix: function(material) {
    const map = this.getMaterialIDMap();
    return map[material] || 'Z';
  },

  updateMaterialName: function(oldName, newName) {
    const sheet = this.getOrCreateMaterialIDSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] === oldName) {
        sheet.getRange(i + 2, 1).setValue(newName);
        return;
      }
    }
  },

  deleteMaterial: function(materialName) {
    const sheet = this.getOrCreateMaterialIDSheet();
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
    const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (values[i][0] === materialName) {
        sheet.deleteRow(i + 2);
        return;
      }
    }
  }
};

const driveManager = {
  getFolderStructure: function() {
    try {
      const primaryFolder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
      const folders = [{ id: primaryFolder.getId(), name: primaryFolder.getName() }];
      
      // Get the artwork folder specifically
      const artworkFolders = primaryFolder.getFoldersByName(CONFIG.ARTWORK_FOLDER_NAME);
      if (artworkFolders.hasNext()) {
        const artworkFolder = artworkFolders.next();
        folders.push({ id: artworkFolder.getId(), name: artworkFolder.getName() });
      }
      
      // Get other subfolders
      const subfolders = primaryFolder.getFolders();
      while (subfolders.hasNext()) {
        const folder = subfolders.next();
        if (folder.getName() !== CONFIG.ARTWORK_FOLDER_NAME) {
          folders.push({ id: folder.getId(), name: folder.getName() });
        }
      }
      
      return folders;
    } catch (e) {
      console.error('Error getting folder structure: ' + e.toString());
      return [];
    }
  },

  uploadFile: function(fileData, folderId, fileName) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      const blob = Utilities.newBlob(Utilities.base64Decode(fileData.split(',')[1]), fileData.split(';')[0].split(':')[1], fileName);
      const file = folder.createFile(blob);
      return { success: true, fileId: file.getId(), fileUrl: file.getUrl(), fileName: file.getName() };
    } catch (e) {
      console.error('Error uploading file: ' + e.toString());
      return { success: false, message: 'Error uploading file: ' + e.toString() };
    }
  },

  updateImageLinks: function() {
    try {
      const primaryFolder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
      const thumbnailFolders = primaryFolder.getFoldersByName('1 - Thumbnail Images');
      
      if (!thumbnailFolders.hasNext()) {
        return { success: false, message: 'Thumbnail folder "1 - Thumbnail Images" not found.' };
      }
      
      const thumbnailFolder = thumbnailFolders.next();
      const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
      
      if (!sheet) {
        return { success: false, message: 'Master sheet not found.' };
      }
      
      const files = thumbnailFolder.getFiles();
      let updateCount = 0;
      
      while (files.hasNext()) {
        const file = files.next();
        const fileName = file.getName();
        const fileBaseName = fileName.substring(0, fileName.lastIndexOf('.')) || fileName;
        const assetId = fileBaseName;
        
        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();
        
        for (let i = 1; i < values.length; i++) {
          const rowId = values[i][CONFIG.COLUMN_MAP.ID - 1];
          if (rowId === assetId) {
            const imageLink = file.getUrl();
            sheet.getRange(i + 1, CONFIG.COLUMN_MAP.IMAGE_LINK).setValue(imageLink);
            updateCount++;
            break;
          }
        }
      }
      
      return { success: true, message: `Updated ${updateCount} image links.`, updateCount: updateCount };
    } catch (e) {
      console.error('Error updating image links: ' + e.toString());
      return { success: false, message: 'Error updating image links: ' + e.toString() };
    }
  }
};

const projectSheet = {
  getActiveSheet: function() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(CONFIG.MASTER_SHEET_NAME);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(CONFIG.MASTER_SHEET_NAME);
      const headers = ['ID', 'Area', 'Asset', 'Status', 'Dimensions', 'Quantity', 'Item', 'Material', 'Due Date', 'Strike Date', 'Venue', 'Location', 'Artwork', 'Image Link', 'Double Sided', 'Diecut', 'Production Status', 'Edit'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
      this.applyColumnDropdowns(sheet);
      this.setupConditionalFormatting(sheet);
    }
    return sheet;
  },

  applyColumnDropdowns: function(sheet) {
    const maxRows = 1000;
    const items = dropdownManager.getDropdownValues('items');
    const materials = dropdownManager.getDropdownValues('materials');
    const statuses = dropdownManager.getDropdownValues('statuses');
    const venues = dropdownManager.getDropdownValues('venues');
    const areas = dropdownManager.getDropdownValues('areas');
    const productionStatuses = dropdownManager.getDropdownValues('productionStatuses');
    
    if (items.length > 0) {
      const itemRule = SpreadsheetApp.newDataValidation().requireValueInList(items, true).setAllowInvalid(true).build();
      sheet.getRange(2, CONFIG.COLUMN_MAP.ITEM, maxRows, 1).setDataValidation(itemRule);
    }
    if (materials.length > 0) {
      const materialRule = SpreadsheetApp.newDataValidation().requireValueInList(materials, true).setAllowInvalid(true).build();
      sheet.getRange(2, CONFIG.COLUMN_MAP.MATERIAL, maxRows, 1).setDataValidation(materialRule);
    }
    if (statuses.length > 0) {
      const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(statuses, true).setAllowInvalid(true).build();
      sheet.getRange(2, CONFIG.COLUMN_MAP.STATUS, maxRows, 1).setDataValidation(statusRule);
    }
    if (venues.length > 0) {
      const venueRule = SpreadsheetApp.newDataValidation().requireValueInList(venues, true).setAllowInvalid(true).build();
      sheet.getRange(2, CONFIG.COLUMN_MAP.VENUE, maxRows, 1).setDataValidation(venueRule);
    }
    if (areas.length > 0) {
      const areaRule = SpreadsheetApp.newDataValidation().requireValueInList(areas, true).setAllowInvalid(true).build();
      sheet.getRange(2, CONFIG.COLUMN_MAP.AREA, maxRows, 1).setDataValidation(areaRule);
    }
    if (productionStatuses.length > 0) {
      const productionStatusRule = SpreadsheetApp.newDataValidation().requireValueInList(productionStatuses, true).setAllowInvalid(true).build();
      sheet.getRange(2, CONFIG.COLUMN_MAP.PRODUCTION_STATUS, maxRows, 1).setDataValidation(productionStatusRule);
    }
  },

  setupConditionalFormatting: function(sheet) {
    sheet.clearConditionalFormatRules();
    const range = sheet.getRange(2, 1, 1000, CONFIG.COLUMN_MAP.EDIT);
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$D2="Delivered"`)
      .setFontColor('#38AE74')
      .setRanges([range])
      .build();
    sheet.setConditionalFormatRules([rule]);
  },

  getRowData: function(rowNumber) {
    try {
      const sheet = this.getActiveSheet();
      const row = parseInt(rowNumber);
      if (!row || isNaN(row) || row < 2) return null;
      const lastRow = sheet.getLastRow();
      if (row > lastRow) return null;
      const range = sheet.getRange(row, 1, 1, CONFIG.COLUMN_MAP.EDIT);
      const values = range.getValues()[0];
      let width = '', height = '';
      const dimensions = values[CONFIG.COLUMN_MAP.DIMENSIONS - 1];
      if (dimensions) {
        const dimMatch = dimensions.toString().match(/^([\d.]+)"\s*x\s*([\d.]+)"$/);
        if (dimMatch) { width = dimMatch[1]; height = dimMatch[2]; }
      }
      const parseDateToISO = (dateValue) => {
        if (!dateValue) return '';
        if (dateValue instanceof Date) return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        const dateStr = dateValue.toString();
        const dateMatch = dateStr.match(/\w+,\s+(\w+)\s+(\d+),\s+(\d+)/);
        if (dateMatch) {
          const date = new Date(`${dateMatch[1]} ${dateMatch[2]}, ${dateMatch[3]}`);
          return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        return '';
      };
      return {
        originalRowNumber: row, item: values[CONFIG.COLUMN_MAP.ITEM - 1] || '', material: values[CONFIG.COLUMN_MAP.MATERIAL - 1] || '',
        asset: values[CONFIG.COLUMN_MAP.ASSET - 1] || '', quantity: values[CONFIG.COLUMN_MAP.QUANTITY - 1] || '', width: width, height: height,
        dieCut: (values[CONFIG.COLUMN_MAP.DIECUT - 1] === true), doubleSided: (values[CONFIG.COLUMN_MAP.DOUBLE_SIDED - 1] === true),
        status: values[CONFIG.COLUMN_MAP.STATUS - 1] || '', dueDate: parseDateToISO(values[CONFIG.COLUMN_MAP.DUE_DATE - 1]),
        strikeDate: parseDateToISO(values[CONFIG.COLUMN_MAP.STRIKE_DATE - 1]), venue: values[CONFIG.COLUMN_MAP.VENUE - 1] || '',
        area: values[CONFIG.COLUMN_MAP.AREA - 1] || '', location: values[CONFIG.COLUMN_MAP.LOCATION - 1] || '', 
        artworkUrl: values[CONFIG.COLUMN_MAP.ARTWORK - 1] || '', imageLink: values[CONFIG.COLUMN_MAP.IMAGE_LINK - 1] || '',
        productionStatus: values[CONFIG.COLUMN_MAP.PRODUCTION_STATUS - 1] || ''
      };
    } catch (error) {
      console.error('Error getting row data:', error);
      return null;
    }
  },

  getNextId: function(material, isEdit = false, currentId = null) {
    const sheet = this.getActiveSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const prefix = materialIDManager.getMaterialPrefix(material);
    
    if (isEdit && currentId) {
      return currentId;
    }
    
    let maxNumber = 0;
    for (let i = 1; i < values.length; i++) {
      const cellValue = values[i][0];
      if (cellValue && typeof cellValue === 'string' && cellValue.startsWith(prefix)) {
        const numberPart = cellValue.substring(prefix.length);
        const number = parseInt(numberPart);
        if (!isNaN(number) && number > maxNumber) maxNumber = number;
      }
    }
    
    const nextNumber = maxNumber + 1;
    const formattedNumber = nextNumber < 10 ? `0${nextNumber}` : nextNumber.toString();
    return `${prefix}${formattedNumber}`;
  },

  getMaterialPrefix: function(material) { return materialIDManager.getMaterialPrefix(material); },
  getNextIdForMaterial: function(material) { return this.getNextId(material); },

  formatDate: function(dateString) {
    if (!dateString) return '';
    const date = new Date(dateString);
    const days = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    const months = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
    return `${days[date.getDay()]}, ${months[date.getMonth()]} ${date.getDate()}, ${date.getFullYear()}`;
  },

  logFormData: function(formData, projectRowNumber, logIdPrefix, logSheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      let logSheet = spreadsheet.getSheetByName(logSheetName);
      if (!logSheet) {
        logSheet = spreadsheet.insertSheet(logSheetName);
        logSheet.hideSheet();
        logSheet.getRange(1, 1, 1, 4).setValues([['LogID', 'ProjectRow', 'Timestamp', 'FormData']]);
      }
      const logId = `${logIdPrefix}_${Date.now()}_${projectRowNumber}`;
      const timestamp = new Date();
      const formDataWithRow = { ...formData, originalRowNumber: projectRowNumber };
      const formDataJson = JSON.stringify(formDataWithRow);
      const lastLogRow = logSheet.getLastRow();
      logSheet.getRange(lastLogRow + 1, 1, 1, 4).setValues([[logId, projectRowNumber, timestamp, formDataJson]]);
      return logId;
    } catch (error) {
      console.error('Error logging form data:', error);
      return null;
    }
  },

  updateLogFormData: function(formData, projectRowNumber, logIdPrefix, logSheetName) {
    try {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const logSheet = spreadsheet.getSheetByName(logSheetName);
      if (!logSheet) return this.logFormData(formData, projectRowNumber, logIdPrefix, logSheetName);
      const dataRange = logSheet.getDataRange();
      const values = dataRange.getValues();
      let existingRowIndex = -1;
      for (let i = 1; i < values.length; i++) {
        if (values[i][1] === projectRowNumber) { existingRowIndex = i + 1; break; }
      }
      const logId = `${logIdPrefix}_${Date.now()}_${projectRowNumber}`;
      const timestamp = new Date();
      const formDataWithRow = { ...formData, originalRowNumber: projectRowNumber };
      const formDataJson = JSON.stringify(formDataWithRow);
      if (existingRowIndex > 0) {
        logSheet.getRange(existingRowIndex, 1, 1, 4).setValues([[logId, projectRowNumber, timestamp, formDataJson]]);
      } else {
        const lastLogRow = logSheet.getLastRow();
        logSheet.getRange(lastLogRow + 1, 1, 1, 4).setValues([[logId, projectRowNumber, timestamp, formDataJson]]);
      }
      return logId;
    } catch (error) {
      console.error('Error updating log data:', error);
      return null;
    }
  },

  addProjectItem: function(assetData, logIdPrefix, logSheetName) {
    try {
      const sheet = this.getActiveSheet();
      const originalRowNumber = assetData.originalRowNumber || (assetData.formData && assetData.formData.originalRowNumber);
      if (originalRowNumber && originalRowNumber > 0) return this.updateProjectItem(assetData, logIdPrefix, logSheetName);
      
      const lastRow = sheet.getLastRow();
      const nextRow = lastRow + 1;
      
      const assetId = this.getNextId(assetData.material);
      const dueDate = this.formatDate(assetData.dueDate);
      const strikeDate = this.formatDate(assetData.strikeDate);
      const dimensions = `${assetData.width}" x ${assetData.height}"`;
      const rowData = [
        assetId, assetData.area || '', assetData.asset || '', 'New Asset', dimensions, assetData.quantity || '', 
        assetData.item || '', assetData.material || '', dueDate, strikeDate, assetData.venue || '', 
        assetData.location || '', assetData.artworkUrl || '', '', assetData.doubleSided || false, 
        assetData.dieCut || false, '', 'Edit'
      ];
      const range = sheet.getRange(nextRow, 1, 1, rowData.length);
      range.setValues([rowData]);
      sheet.getRange(nextRow, CONFIG.COLUMN_MAP.DOUBLE_SIDED).insertCheckboxes();
      sheet.getRange(nextRow, CONFIG.COLUMN_MAP.DIECUT).insertCheckboxes();
      range.setFontColor('#0062e2');
      
      const logId = this.logFormData(assetData.formData || assetData, nextRow, logIdPrefix, logSheetName);
      if (logId) {
        const editCell = sheet.getRange(nextRow, CONFIG.COLUMN_MAP.EDIT);
        editCell.setNote(`LogID: ${logId}\n\nTo edit this item:\n1. Select this cell\n2. Go to ${CONFIG.PROJECT_NAME} > Edit Selected Row`);
        editCell.setBackground('#e3f2fd');
        editCell.setFontColor('#1976d2');
        editCell.setFontWeight('bold');
      }
      
      this.applyDropdownsToRow(sheet, nextRow);
      this.forceSort(sheet);
      
      return { success: true, message: `Asset added to row ${nextRow}`, rowNumber: nextRow, logId: logId, isUpdate: false };
    } catch (error) {
      console.error('Error adding project item:', error);
      return { success: false, message: `Error adding item: ${error.message}`, rowNumber: null, logId: null, isUpdate: false };
    }
  },

  applyDropdownsToRow: function(sheet, rowNum) {
    try {
      const items = dropdownManager.getDropdownValues('items');
      const materials = dropdownManager.getDropdownValues('materials');
      const statuses = dropdownManager.getDropdownValues('statuses');
      const venues = dropdownManager.getDropdownValues('venues');
      const areas = dropdownManager.getDropdownValues('areas');
      const productionStatuses = dropdownManager.getDropdownValues('productionStatuses');
      
      if (items.length > 0) {
        const itemRule = SpreadsheetApp.newDataValidation().requireValueInList(items, true).setAllowInvalid(true).build();
        sheet.getRange(rowNum, CONFIG.COLUMN_MAP.ITEM).setDataValidation(itemRule);
      }
      if (materials.length > 0) {
        const materialRule = SpreadsheetApp.newDataValidation().requireValueInList(materials, true).setAllowInvalid(true).build();
        sheet.getRange(rowNum, CONFIG.COLUMN_MAP.MATERIAL).setDataValidation(materialRule);
      }
      if (statuses.length > 0) {
        const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(statuses, true).setAllowInvalid(true).build();
        sheet.getRange(rowNum, CONFIG.COLUMN_MAP.STATUS).setDataValidation(statusRule);
      }
      if (venues.length > 0) {
        const venueRule = SpreadsheetApp.newDataValidation().requireValueInList(venues, true).setAllowInvalid(true).build();
        sheet.getRange(rowNum, CONFIG.COLUMN_MAP.VENUE).setDataValidation(venueRule);
      }
      if (areas.length > 0) {
        const areaRule = SpreadsheetApp.newDataValidation().requireValueInList(areas, true).setAllowInvalid(true).build();
        sheet.getRange(rowNum, CONFIG.COLUMN_MAP.AREA).setDataValidation(areaRule);
      }
      if (productionStatuses.length > 0) {
        const productionStatusRule = SpreadsheetApp.newDataValidation().requireValueInList(productionStatuses, true).setAllowInvalid(true).build();
        sheet.getRange(rowNum, CONFIG.COLUMN_MAP.PRODUCTION_STATUS).setDataValidation(productionStatusRule);
      }
    } catch (error) {
      console.error('Error applying dropdowns to row:', error);
    }
  },

  updateProjectItem: function(assetData, logIdPrefix, logSheetName) {
    try {
      const sheet = this.getActiveSheet();
      const rowNum = parseInt(assetData.originalRowNumber);
      if (!rowNum || isNaN(rowNum) || rowNum < 1) throw new Error(`Invalid row number: ${assetData.originalRowNumber}`);
      const existingId = sheet.getRange(rowNum, CONFIG.COLUMN_MAP.ID).getValue();
      
      const assetId = existingId || this.getNextId(assetData.material, true, existingId);
      
      const dueDate = this.formatDate(assetData.dueDate);
      const strikeDate = this.formatDate(assetData.strikeDate);
      const dimensions = `${assetData.width}" x ${assetData.height}"`;
      const rowData = [
        assetId, assetData.area || '', assetData.asset || '', assetData.status || 'New Asset', dimensions, 
        assetData.quantity || '', assetData.item || '', assetData.material || '', dueDate, strikeDate, 
        assetData.venue || '', assetData.location || '', assetData.artworkUrl || '', assetData.imageLink || '', 
        assetData.doubleSided || false, assetData.dieCut || false, assetData.productionStatus || '', 'Edit'
      ];
      const range = sheet.getRange(rowNum, 1, 1, rowData.length - 1);
      range.setValues([rowData.slice(0, -1)]);
      const doubleSidedCell = sheet.getRange(rowNum, CONFIG.COLUMN_MAP.DOUBLE_SIDED);
      const diecutCell = sheet.getRange(rowNum, CONFIG.COLUMN_MAP.DIECUT);
      if (doubleSidedCell.getDataValidation() === null) doubleSidedCell.insertCheckboxes();
      if (diecutCell.getDataValidation() === null) diecutCell.insertCheckboxes();
      const rowRange = sheet.getRange(rowNum, 1, 1, CONFIG.COLUMN_MAP.EDIT);
      if (assetData.status === 'New Asset') rowRange.setFontColor('#0062e2');
      else if (assetData.status === 'Delivered') rowRange.setFontColor('#38AE74');
      else if (assetData.status === 'On Hold') rowRange.setFontColor('#f7c831');
      else rowRange.setFontColor('#000000');
      const logId = this.updateLogFormData(assetData.formData || assetData, rowNum, logIdPrefix, logSheetName);
      if (logId) {
        const editCell = sheet.getRange(rowNum, CONFIG.COLUMN_MAP.EDIT);
        editCell.setNote(`LogID: ${logId}\n\nTo edit this item:\n1. Select this cell\n2. Go to ${CONFIG.PROJECT_NAME} > Edit Selected Row\n\nLast updated: ${new Date().toLocaleString()}`);
      }
      
      this.forceSort(sheet);
      
      return { success: true, message: `Asset updated in row ${rowNum}`, rowNumber: rowNum, logId: logId, isUpdate: true };
    } catch (error) {
      console.error('Error updating project item:', error);
      return { success: false, message: `Error updating item: ${error.message}`, rowNumber: null, logId: null, isUpdate: false };
    }
  },

  forceSort: function(sheet) {
    try {
      Utilities.sleep(100);
      ensureNewAssetsAtTop(sheet);
    } catch (error) {
      console.error('Error in forceSort:', error);
    }
  }
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu(CONFIG.PROJECT_NAME)
    .addItem('Add New Asset', 'openAssetApp')
    .addSeparator()
    .addItem('Edit Selected Row', 'editSelectedRow')
    .addSeparator()
    .addItem('Reorder Asset', 'openReorderAssetApp')
    .addSeparator()
    .addSubMenu(ui.createMenu('Additional')
      .addItem('Edit All Dropdowns', 'openDropdownEditor')
      .addItem('Update All File Names', 'updateFileNames')
      .addItem('Update Image Links', 'updateImageLinks'))
    .addToUi();
  setupAutoSort();
  materialIDManager.getOrCreateMaterialIDSheet();
  initializeProductionStatusSheet();
}

function initializeProductionStatusSheet() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName('ProductionStatusesList');
    
    if (!sheet) {
      sheet = spreadsheet.insertSheet('ProductionStatusesList');
      sheet.getRange(1, 1).setValue('Production Statuses');
      sheet.getRange(1, 1).setFontWeight('bold');
      sheet.hideSheet();
    }
  } catch (error) {
    console.error('Error initializing Production Status sheet:', error);
  }
}

function openDropdownEditor() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('DropdownEditor').setTitle('Dropdown Editor').setWidth(320);
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function setupAutoSort() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onEditInstallable' || trigger.getHandlerFunction() === 'onChangeInstallable') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger('onEditInstallable').forSpreadsheet(SpreadsheetApp.getActive()).onEdit().create();
  ScriptApp.newTrigger('onChangeInstallable').forSpreadsheet(SpreadsheetApp.getActive()).onChange().create();
}

function updateFileNames() {
  try {
    const ui = SpreadsheetApp.getUi();
    const mainFolder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    const artworkFolders = mainFolder.getFoldersByName(CONFIG.ARTWORK_FOLDER_NAME);
    
    if (!artworkFolders.hasNext()) {
      ui.alert('Folder Not Found', `The "${CONFIG.ARTWORK_FOLDER_NAME}" folder was not found.`, ui.ButtonSet.OK);
      return;
    }
    
    const artworkFolder = artworkFolders.next();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
    if (!sheet) { ui.alert('Error', 'Master sheet not found.', ui.ButtonSet.OK); return; }
    const files = artworkFolder.getFiles();
    let renamedCount = 0, skippedCount = 0, linksAddedCount = 0;
    const renamedFiles = [];
    const fileMap = {};
    
    while (files.hasNext()) {
      const file = files.next();
      const fileName = file.getName();
      const idMatch = fileName.match(/^([A-Z]\d+)/);
      if (idMatch) {
        const assetId = idMatch[1];
        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();
        for (let i = 1; i < values.length; i++) {
          const rowId = values[i][CONFIG.COLUMN_MAP.ID - 1];
          if (rowId === assetId) {
            const assetName = values[i][CONFIG.COLUMN_MAP.ASSET - 1];
            if (assetName) {
              const lastDotIndex = fileName.lastIndexOf('.');
              const extension = lastDotIndex > -1 ? fileName.substring(lastDotIndex) : '';
              const sanitizedAssetName = assetName.toString().replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_-]/g, '_');
              const newFileName = `${assetId}_${CONFIG.PROJECT_CODE}_${sanitizedAssetName}${extension}`;
              if (fileName !== newFileName) {
                file.setName(newFileName);
                renamedFiles.push(`${fileName} → ${newFileName}`);
                renamedCount++;
              } else skippedCount++;
              
              fileMap[assetId] = file.getUrl();
            }
            break;
          }
        }
      }
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    for (let i = 1; i < values.length; i++) {
      const rowId = values[i][CONFIG.COLUMN_MAP.ID - 1];
      if (rowId && fileMap[rowId]) {
        const artworkCell = sheet.getRange(i + 1, CONFIG.COLUMN_MAP.ARTWORK);
        const currentValue = artworkCell.getValue();
        
        if (!currentValue || currentValue !== fileMap[rowId]) {
          artworkCell.setValue(fileMap[rowId]);
          linksAddedCount++;
        }
      }
    }
    
    let message = `File Update Complete!\n\n✅ Renamed: ${renamedCount} file(s)\n⏭️ Skipped: ${skippedCount} file(s) (already correct)\n🔗 Links Added: ${linksAddedCount} artwork link(s)\n`;
    if (renamedFiles.length > 0) {
      message += `\nRenamed Files:\n${renamedFiles.slice(0, 10).join('\n')}`;
      if (renamedFiles.length > 10) message += `\n... and ${renamedFiles.length - 10} more`;
    }
    ui.alert('Update File Names', message, ui.ButtonSet.OK);
  } catch (error) {
    console.error('Error in updateFileNames:', error);
    SpreadsheetApp.getUi().alert('Error', 'Failed to update file names: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function updateImageLinks() {
  try {
    const result = driveManager.updateImageLinks();
    const ui = SpreadsheetApp.getUi();
    
    if (result.success) {
      ui.alert('Update Image Links', `Successfully updated ${result.updateCount} image links from the "1 - Thumbnail Images" folder.`, ui.ButtonSet.OK);
    } else {
      ui.alert('Error', result.message, ui.ButtonSet.OK);
    }
  } catch (error) {
    console.error('Error in updateImageLinks:', error);
    SpreadsheetApp.getUi().alert('Error', 'Failed to update image links: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function onEditInstallable(e) { handleEdit(e); }

function onChangeInstallable(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
  if (sheet) ensureNewAssetsAtTop(sheet);
}

function handleEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== CONFIG.MASTER_SHEET_NAME) return;
  const range = e.range;
  const row = range.getRow();
  if (row < 2) return;
  ensureCheckboxesInRow(sheet, row);
  if (range.getColumn() === CONFIG.COLUMN_MAP.STATUS) {
    const newStatus = range.getValue();
    const rowRange = sheet.getRange(row, 1, 1, CONFIG.COLUMN_MAP.EDIT);
    if (newStatus === 'New Asset') rowRange.setFontColor('#0062e2');
    else if (newStatus === 'Requires Attn') {
      rowRange.setFontColor('#FF2093');
      showCommentDialog(sheet, row);
    } else if (newStatus === 'Delivered') rowRange.setFontColor('#38AE74');
    else if (newStatus === 'On Hold') rowRange.setFontColor('#f7c831');
    else if (newStatus && newStatus !== 'New Asset' && newStatus !== 'Requires Attn' && newStatus !== 'Delivered' && newStatus !== 'On Hold') rowRange.setFontColor('#000000');
    ensureNewAssetsAtTop(sheet);
  }
  if (range.getColumn() !== CONFIG.COLUMN_MAP.STATUS && range.getColumn() !== CONFIG.COLUMN_MAP.EDIT) {
    const statusCell = sheet.getRange(row, CONFIG.COLUMN_MAP.STATUS);
    const currentStatus = statusCell.getValue();
    if (currentStatus === 'New Asset') sheet.getRange(row, 1, 1, CONFIG.COLUMN_MAP.EDIT).setFontColor('#0062e2');
    else if (currentStatus === 'Requires Attn') sheet.getRange(row, 1, 1, CONFIG.COLUMN_MAP.EDIT).setFontColor('#FF2093');
    else if (currentStatus === 'Delivered') sheet.getRange(row, 1, 1, CONFIG.COLUMN_MAP.EDIT).setFontColor('#38AE74');
    else if (currentStatus === 'On Hold') sheet.getRange(row, 1, 1, CONFIG.COLUMN_MAP.EDIT).setFontColor('#f7c831');
  }
  if (range.getColumn() !== CONFIG.COLUMN_MAP.EDIT) logRowEdit(sheet, row);
}

function ensureCheckboxesInRow(sheet, row) {
  try {
    const doubleSidedCell = sheet.getRange(row, CONFIG.COLUMN_MAP.DOUBLE_SIDED);
    const diecutCell = sheet.getRange(row, CONFIG.COLUMN_MAP.DIECUT);
    const doubleSidedValue = doubleSidedCell.getValue();
    const diecutValue = diecutCell.getValue();
    const doubleSidedValidation = doubleSidedCell.getDataValidation();
    if (doubleSidedValidation === null || doubleSidedValidation.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
      doubleSidedCell.insertCheckboxes();
      if (doubleSidedValue === 'TRUE' || doubleSidedValue === true) doubleSidedCell.setValue(true);
      else if (doubleSidedValue === 'FALSE' || doubleSidedValue === false) doubleSidedCell.setValue(false);
    }
    const diecutValidation = diecutCell.getDataValidation();
    if (diecutValidation === null || diecutValidation.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.CHECKBOX) {
      diecutCell.insertCheckboxes();
      if (diecutValue === 'TRUE' || diecutValue === true) diecutCell.setValue(true);
      else if (diecutValue === 'FALSE' || diecutValue === false) diecutCell.setValue(false);
    }
  } catch (error) {
    console.error('Error ensuring checkboxes:', error);
  }
}

function showCommentDialog(sheet, row) {
  try {
    const rowData = sheet.getRange(row, 1, 1, CONFIG.COLUMN_MAP.EDIT).getValues()[0];
    const assetId = rowData[CONFIG.COLUMN_MAP.ID - 1];
    const assetName = rowData[CONFIG.COLUMN_MAP.ASSET - 1];
    const material = rowData[CONFIG.COLUMN_MAP.MATERIAL - 1];
    const item = rowData[CONFIG.COLUMN_MAP.ITEM - 1];
    const htmlOutput = HtmlService.createHtmlOutputFromFile('CommentDialog').setWidth(420).setHeight(280);
    const htmlContent = htmlOutput.getContent();
    const modifiedContent = htmlContent.replace('<script>', `<script>window.assetData = ${JSON.stringify({ row: row, assetId: assetId || 'N/A', assetName: assetName || 'N/A', item: item || 'N/A', material: material || 'N/A' })};`);
    const modifiedOutput = HtmlService.createHtmlOutput(modifiedContent).setWidth(420).setHeight(280);
    SpreadsheetApp.getUi().showModalDialog(modifiedOutput, '⚠️ Attention Required');
  } catch (error) {
    console.error('Error showing comment dialog:', error);
  }
}

function sendNotification(data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.MASTER_SHEET_NAME);
    const row = data.row;
    const spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    const sheetName = sheet.getName();
    const cellReference = `${sheetName}!D${row}`;
    const directLink = `${spreadsheetUrl}#gid=${sheet.getSheetId()}&range=D${row}`;
    const emailSubject = `🚨 Attention Required - ${CONFIG.PROJECT_NAME} - Asset ${data.assetId}`;
    const emailBody = `<html><body style="font-family: Arial, sans-serif; color: #333;"><h2 style="color: #FF2093;">⚠️ Asset Requires Attention</h2><div style="background-color: #fef2f2; padding: 15px; border-left: 4px solid #FF2093; margin: 20px 0;"><p><strong>Project:</strong> ${CONFIG.PROJECT_NAME}</p><p><strong>Asset Name:</strong> ${data.assetName}</p><p><strong>Asset ID:</strong> ${data.assetId}</p><p><strong>Item:</strong> ${data.item}</p><p><strong>Material:</strong> ${data.material}</p><p><strong>Location:</strong> ${cellReference}</p></div><div style="background-color: #f9fafb; padding: 15px; border-radius: 6px; margin: 20px 0;"><p style="margin: 0;"><strong>Comment:</strong></p><p style="margin: 10px 0 0 0; white-space: pre-wrap;">${data.comment}</p></div><p>This asset has been flagged as <strong style="color: #FF2093;">"Requires Attn"</strong> and needs your immediate attention.</p><p style="margin-top: 30px;"><a href="${directLink}" style="background-color: #FF2093; color: white; padding: 10px 20px; text-decoration: none; border-radius: 4px;">View Asset in Spreadsheet</a></p><hr style="margin-top: 30px; border: none; border-top: 1px solid #ddd;"><p style="font-size: 12px; color: #666;">This is an automated notification from the ${CONFIG.PROJECT_NAME} Asset Management System.</p></body></html>`;
    
    // Send single email to all recipients
    try { 
      MailApp.sendEmail({ 
        to: CONFIG.ALERT_EMAILS, 
        subject: emailSubject, 
        htmlBody: emailBody 
      }); 
    } catch (emailError) { 
      console.error('Error sending email:', emailError);
      return { success: false, message: 'Error sending email: ' + emailError.toString() };
    }
    
    return { success: true, message: 'Notification sent successfully' };
  } catch (error) {
    console.error('Error sending notification:', error);
    return { success: false, message: error.toString() };
  }
}

function logRowEdit(sheet, row) {
  try {
    const rowData = sheet.getRange(row, 1, 1, CONFIG.COLUMN_MAP.EDIT).getValues()[0];
    const columnsToCheck = [CONFIG.COLUMN_MAP.ID - 1, CONFIG.COLUMN_MAP.ASSET - 1, CONFIG.COLUMN_MAP.DIMENSIONS - 1, CONFIG.COLUMN_MAP.QUANTITY - 1, CONFIG.COLUMN_MAP.DUE_DATE - 1, CONFIG.COLUMN_MAP.STRIKE_DATE - 1, CONFIG.COLUMN_MAP.LOCATION - 1, CONFIG.COLUMN_MAP.ARTWORK - 1];
    const hasContent = columnsToCheck.some(index => { const cell = rowData[index]; return cell !== '' && cell !== null && cell !== undefined; });
    if (!hasContent) return;
    const parseDateToISO = (dateValue) => {
      if (!dateValue) return '';
      if (dateValue instanceof Date) return Utilities.formatDate(dateValue, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const dateStr = dateValue.toString();
      const dateMatch = dateStr.match(/\w+,\s+(\w+)\s+(\d+),\s+(\d+)/);
      if (dateMatch) {
        const date = new Date(`${dateMatch[1]} ${dateMatch[2]}, ${dateMatch[3]}`);
        return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      return '';
    };
    let width = '', height = '';
    const dimensions = rowData[CONFIG.COLUMN_MAP.DIMENSIONS - 1];
    if (dimensions) {
      const dimMatch = dimensions.toString().match(/^([\d.]+)"\s*x\s*([\d.]+)"$/);
      if (dimMatch) { width = dimMatch[1]; height = dimMatch[2]; }
    }
    const formData = {
      originalRowNumber: row, item: rowData[CONFIG.COLUMN_MAP.ITEM - 1] || '', material: rowData[CONFIG.COLUMN_MAP.MATERIAL - 1] || '',
      asset: rowData[CONFIG.COLUMN_MAP.ASSET - 1] || '', quantity: rowData[CONFIG.COLUMN_MAP.QUANTITY - 1] || '', width: width, height: height,
      dieCut: (rowData[CONFIG.COLUMN_MAP.DIECUT - 1] === true), doubleSided: (rowData[CONFIG.COLUMN_MAP.DOUBLE_SIDED - 1] === true),
      status: rowData[CONFIG.COLUMN_MAP.STATUS - 1] || '', dueDate: parseDateToISO(rowData[CONFIG.COLUMN_MAP.DUE_DATE - 1]),
      strikeDate: parseDateToISO(rowData[CONFIG.COLUMN_MAP.STRIKE_DATE - 1]), venue: rowData[CONFIG.COLUMN_MAP.VENUE - 1] || '',
      area: rowData[CONFIG.COLUMN_MAP.AREA - 1] || '', location: rowData[CONFIG.COLUMN_MAP.LOCATION - 1] || '', 
      artworkUrl: rowData[CONFIG.COLUMN_MAP.ARTWORK - 1] || '', imageLink: rowData[CONFIG.COLUMN_MAP.IMAGE_LINK - 1] || '',
      productionStatus: rowData[CONFIG.COLUMN_MAP.PRODUCTION_STATUS - 1] || ''
    };
    const logId = projectSheet.updateLogFormData(formData, row, CONFIG.LOG_ID_PREFIX, CONFIG.LOG_SHEET_NAME);
    const editCell = sheet.getRange(row, CONFIG.COLUMN_MAP.EDIT);
    editCell.setValue('Edit');
    if (logId) {
      editCell.setNote(`LogID: ${logId}\n\nTo edit this item:\n1. Select this cell\n2. Go to ${CONFIG.PROJECT_NAME} > Edit Selected Row\n\nLast updated: ${new Date().toLocaleString()}`);
      editCell.setBackground('#e3f2fd');
      editCell.setFontColor('#1976d2');
      editCell.setFontWeight('bold');
    }
  } catch (error) {
    console.error('Error logging row edit:', error);
  }
}

function ensureNewAssetsAtTop(sheet) {
  try {
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;
    const dataRange = sheet.getRange(2, 1, lastRow - 1, CONFIG.COLUMN_MAP.EDIT);
    const values = dataRange.getValues();
    const newAssetRows = [], otherRows = [];
    values.forEach((row, index) => {
      const status = row[CONFIG.COLUMN_MAP.STATUS - 1];
      if (status === 'New Asset') newAssetRows.push(row);
      else otherRows.push(row);
    });
    otherRows.sort((a, b) => {
      const idA = a[0] ? a[0].toString() : '';
      const idB = b[0] ? b[0].toString() : '';
      return idA.localeCompare(idB);
    });
    const sortedData = [...newAssetRows, ...otherRows];
    const currentOrder = values.map(row => row.join('|'));
    const newOrder = sortedData.map(row => row.join('|'));
    if (currentOrder.join('||') !== newOrder.join('||')) {
      dataRange.setValues(sortedData);
      for (let i = 0; i < sortedData.length; i++) {
        const rowNum = i + 2;
        const status = sortedData[i][CONFIG.COLUMN_MAP.STATUS - 1];
        const rowRange = sheet.getRange(rowNum, 1, 1, CONFIG.COLUMN_MAP.EDIT);
        if (status === 'New Asset') rowRange.setFontColor('#0062e2');
        else if (status === 'Requires Attn') rowRange.setFontColor('#FF2093');
        else if (status === 'Delivered') rowRange.setFontColor('#38AE74');
        else rowRange.setFontColor('#000000');
      }
    }
  } catch (error) {
    console.error('Error in ensureNewAssetsAtTop:', error);
  }
}

function sortNonNewAssets() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getSheetByName(CONFIG.MASTER_SHEET_NAME);
  if (!sheet) return;
  ensureNewAssetsAtTop(sheet);
}

function onEdit(e) { handleEdit(e); }
function openAssetApp() { assetApp.showDialog(); }

function editSelectedRow() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getActiveSheet();
    if (sheet.getName() !== CONFIG.MASTER_SHEET_NAME) {
      SpreadsheetApp.getUi().alert('Edit Row', `Please select a row in the "${CONFIG.MASTER_SHEET_NAME}" sheet to edit.`, SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    const activeCell = sheet.getActiveCell();
    const rowNumber = activeCell.getRow();
    if (rowNumber < 2) {
      SpreadsheetApp.getUi().alert('Edit Row', 'Please select a data row (not the header row) to edit.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    assetApp.openForEdit(rowNumber);
  } catch (error) {
    console.error('Error in editSelectedRow:', error);
    SpreadsheetApp.getUi().alert('Error', 'An error occurred while trying to edit the row: ' + error.message, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function getDropdownData() { return dropdownManager.getAllDropdownData(); }
function addNewDropdownValue(fieldName, value) { return dropdownManager.addDropdownValue(fieldName, value); }
function updateDropdownValue(fieldName, oldValue, newValue) { return dropdownManager.updateDropdownValue(fieldName, oldValue, newValue); }
function deleteDropdownValue(fieldName, value) { return dropdownManager.deleteDropdownValue(fieldName, value); }
function getDriveFolders() { return driveManager.getFolderStructure(); }
function uploadFileToDrive(fileData, folderId, fileName) { return driveManager.uploadFile(fileData, folderId, fileName); }
function addAssetToProject(assetData) { return assetApp.addToProject(assetData); }
function getMaterialPrefix(material) { return projectSheet.getMaterialPrefix(material); }
function getNextIdForMaterial(material) { return projectSheet.getNextIdForMaterial(material); }
function getRowDataForEdit(rowNumber) { 
  const rowData = projectSheet.getRowData(rowNumber);
  if (rowData) {
    const sheet = projectSheet.getActiveSheet();
    const idValue = sheet.getRange(rowNumber, CONFIG.COLUMN_MAP.ID).getValue();
    return { ...rowData, id: idValue };
  }
  return null;
}

function openReorderAssetApp() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('ReorderAssetForm')
      .setWidth(600).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Reorder Asset');
}

function getAvailableAssetsForReorder() {
  try {
    const sheet = projectSheet.getActiveSheet();
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) return [];
    
    const dataRange = sheet.getRange(2, 1, lastRow - 1, CONFIG.COLUMN_MAP.EDIT);
    const values = dataRange.getValues();
    
    const availableAssets = [];
    const seenAssets = new Set();
    
    for (let i = 0; i < values.length; i++) {
      const assetId = values[i][CONFIG.COLUMN_MAP.ID - 1];
      const assetName = values[i][CONFIG.COLUMN_MAP.ASSET - 1];
      
      if (!assetName) continue;
      
      const isConcatenated = /^[A-Z]\d+\s*-\s*.+/.test(assetName.toString());
      
      if (!isConcatenated && !seenAssets.has(assetName)) {
        seenAssets.add(assetName);
        availableAssets.push({
          name: assetName,
          rowNumber: i + 2
        });
      }
    }
    
    availableAssets.sort((a, b) => a.name.localeCompare(b.name));
    
    return availableAssets;
  } catch (error) {
    console.error('Error getting available assets:', error);
    throw new Error('Failed to load assets: ' + error.message);
  }
}

function reorderAsset(reorderData) {
  try {
    const sheet = projectSheet.getActiveSheet();
    const originalRowNumber = parseInt(reorderData.originalRowNumber);
    
    if (!originalRowNumber || originalRowNumber < 2) {
      return { success: false, message: 'Invalid asset selection' };
    }
    
    const originalRange = sheet.getRange(originalRowNumber, 1, 1, CONFIG.COLUMN_MAP.EDIT);
    const originalData = originalRange.getValues()[0];
    
    const originalId = originalData[CONFIG.COLUMN_MAP.ID - 1];
    const originalAsset = originalData[CONFIG.COLUMN_MAP.ASSET - 1];
    const originalMaterial = originalData[CONFIG.COLUMN_MAP.MATERIAL - 1];
    
    const newId = projectSheet.getNextId(originalMaterial);
    const newAssetName = `${originalId} - ${originalAsset}`;
    
    const lastRow = sheet.getLastRow();
    const newRow = lastRow + 1;
    
    const rowData = [
      newId,
      originalData[CONFIG.COLUMN_MAP.AREA - 1] || '',
      newAssetName,
      'New Asset',
      originalData[CONFIG.COLUMN_MAP.DIMENSIONS - 1] || '',
      reorderData.quantity,
      originalData[CONFIG.COLUMN_MAP.ITEM - 1] || '',
      originalMaterial || '',
      originalData[CONFIG.COLUMN_MAP.DUE_DATE - 1] || '',
      originalData[CONFIG.COLUMN_MAP.STRIKE_DATE - 1] || '',
      originalData[CONFIG.COLUMN_MAP.VENUE - 1] || '',
      originalData[CONFIG.COLUMN_MAP.LOCATION - 1] || '',
      originalData[CONFIG.COLUMN_MAP.ARTWORK - 1] || '',
      originalData[CONFIG.COLUMN_MAP.IMAGE_LINK - 1] || '',
      originalData[CONFIG.COLUMN_MAP.DOUBLE_SIDED - 1] || false,
      originalData[CONFIG.COLUMN_MAP.DIECUT - 1] || false,
      '',
      'Edit'
    ];
    
    const range = sheet.getRange(newRow, 1, 1, rowData.length);
    range.setValues([rowData]);
    
    sheet.getRange(newRow, CONFIG.COLUMN_MAP.DOUBLE_SIDED).insertCheckboxes();
    sheet.getRange(newRow, CONFIG.COLUMN_MAP.DIECUT).insertCheckboxes();
    
    range.setFontColor('#0062e2');
    
    projectSheet.applyDropdownsToRow(sheet, newRow);
    
    const editCell = sheet.getRange(newRow, CONFIG.COLUMN_MAP.EDIT);
    editCell.setBackground('#e3f2fd');
    editCell.setFontColor('#1976d2');
    editCell.setFontWeight('bold');
    editCell.setNote(`Reordered from: ${originalAsset} (ID: ${originalId})\nCreated: ${new Date().toLocaleString()}`);
    
    const logData = {
      originalRowNumber: newRow,
      reorderedFrom: originalId,
      originalAsset: originalAsset,
      newQuantity: reorderData.quantity
    };
    projectSheet.logFormData(logData, newRow, CONFIG.LOG_ID_PREFIX, CONFIG.LOG_SHEET_NAME);
    
    projectSheet.forceSort(sheet);
    
    return { 
      success: true, 
      message: `Asset reordered successfully! New ID: ${newId}`,
      newId: newId,
      rowNumber: newRow
    };
    
  } catch (error) {
    console.error('Error reordering asset:', error);
    return { success: false, message: 'Error reordering asset: ' + error.message };
  }
}';
}

function getAssetFormHtmlTemplate() {
  return '<!DOCTYPE html>
<html>
<head>
    <base target="_top">
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        input[type=number]::-webkit-inner-spin-button,
        input[type=number]::-webkit-outer-spin-button {
            -webkit-appearance: none;
            margin: 0;
        }
        input[type=number] {
            -moz-appearance: textfield;
        }
        .add-link {
            color: #2563eb;
            font-size: 0.75rem;
            font-weight: 600;
            cursor: pointer;
            display: inline-flex;
            align-items: center;
            gap: 0.25rem;
        }
        .add-link:hover {
            color: #1d4ed8;
        }
        .hidden { display: none; }
        #side-drawer {
            transition: transform 0.3s ease-in-out;
        }
        .accordion-item {
            border-bottom: 1px solid #e5e7eb;
        }
        .accordion-header {
            cursor: pointer;
            padding: 12px 16px;
            background: #f9fafb;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .accordion-header:hover {
            background: #f3f4f6;
        }
        .accordion-content {
            max-height: 0;
            overflow: hidden;
            transition: max-height 0.3s ease;
        }
        .accordion-content.active {
            max-height: 500px;
            overflow-y: auto;
        }
        .dropdown-list-item {
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 8px 16px;
            border-bottom: 1px solid #f3f4f6;
        }
        .dropdown-list-item input {
            flex: 1;
        }
        .dropdown-list-item button {
            padding: 4px 8px;
            font-size: 12px;
        }
        
        /* Compact form styling */
        label {
            font-size: 0.75rem;
            font-weight: 500;
            margin-bottom: 0.25rem;
        }
        
        input, select {
            font-size: 0.813rem;
            padding: 0.375rem 0.5rem;
        }
        
        .form-section {
            margin-bottom: 1rem;
        }
    </style>
</head>
<body class="bg-gray-100 p-4">
    <div class="flex">
        <!-- Main Form Content -->
        <div class="flex-1">
            <div class="max-w-3xl mx-auto bg-white rounded-lg shadow-lg p-5">
                <h1 class="text-xl font-semibold text-gray-800 mb-4">Add New Asset</h1>

                <!-- Asset Details Section -->
                <div class="form-section">
                    <h2 class="text-sm font-semibold text-gray-800 mb-3">Asset Details</h2>
                    
                    <!-- Item and Material - Same Row -->
                    <div class="grid grid-cols-2 gap-3 mb-3">
                        <div>
                            <label class="text-xs font-medium text-gray-700 block mb-1">Item</label>
                            <select id="item-dropdown" class="w-full px-2 py-1.5 border border-gray-300 rounded text-xs">
                                <option value="">Select Item</option>
                            </select>
                            <input id="item-input" type="text" class="hidden w-full px-2 py-1.5 border border-gray-300 rounded text-xs mt-1" placeholder="Enter new item" />
                            <a id="add-item-link" class="add-link mt-0.5">+ Add New Item</a>
                        </div>
                        <div>
                            <label class="text-xs font-medium text-gray-700 block mb-1">Material</label>
                            <select id="material-dropdown" class="w-full px-2 py-1.5 border border-gray-300 rounded text-xs">
                                <option value="">Select Material</option>
                            </select>
                            <input id="material-input" type="text" class="hidden w-full px-2 py-1.5 border border-gray-300 rounded text-xs mt-1" placeholder="Enter new material" />
                            <a id="add-material-link" class="add-link mt-0.5">+ Add New Material</a>
                        </div>
                    </div>

                    <!-- Asset Name, Quantity, Width, Height - Same Row -->
                    <div class="grid grid-cols-12 gap-2 mb-3">
                        <div class="col-span-6">
                            <label class="text-xs font-medium text-gray-700 block mb-1">Asset Name</label>
                            <input id="asset-name" type="text" class="w-full px-2 py-1.5 border border-gray-300 rounded text-xs" />
                        </div>
                        <div class="col-span-2">
                            <label class="text-xs font-medium text-gray-700 block mb-1">Quantity</label>
                            <input id="quantity" type="number" min="1" value="1" class="w-full px-2 py-1.5 border border-gray-300 rounded text-xs" />
                        </div>
                        <div class="col-span-2">
                            <label class="text-xs font-medium text-gray-700 block mb-1">Width</label>
                            <input id="width" type="number" step="0.1" class="w-full px-2 py-1.5 border border-gray-300 rounded text-xs" />
                        </div>
                        <div class="col-span-2">
                            <label class="text-xs font-medium text-gray-700 block mb-1">Height</label>
                            <input id="height" type="number" step="0.1" class="w-full px-2 py-1.5 border border-gray-300 rounded text-xs" />
                        </div>
                    </div>

                    <!-- Die Cut and Double Sided Checkboxes -->
                    <div class="flex gap-4 mb-3">
                        <label class="flex items-center">
                            <input type="checkbox" id="die-cut" class="mr-1.5 h-3.5 w-3.5 text-blue-600">
                            <span class="text-xs text-gray-700">Die Cut Decal</span>
                        </label>
                        <label class="flex items-center">
                            <input type="checkbox" id="double-sided" class="mr-1.5 h-3.5 w-3.5 text-blue-600">
                            <span class="text-xs text-gray-700">Double Sided</span>
                        </label>
                    </div>
                </div>

                <!-- Production Details Section -->
                <div class="form-section">
                    <h2 class="text-sm font-semibold text-gray-800 mb-3">Production Details</h2>
                    
                    <!-- Status, In-Hands Date, Strike Date - Same Row -->
                    <div class="grid grid-cols-12 gap-2 mb-3">
                        <div class="col-span-6">
                            <label class="text-xs font-medium text-gray-700 block mb-1">Status</label>
                            <select id="status-dropdown" class="w-full px-2 py-1.5 border border-gray-300 rounded text-xs">
                                <option value="">Select Status</option>
                            </select>
                            <input id="status-input" type="text" class="hidden w-full px-2 py-1.5 border border-gray-300 rounded text-xs mt-1" placeholder="Enter new status" />
                            <a id="add-status-link" class="add-link mt-0.5">+ Add New Status</a>
                        </div>
                        <div class="col-span-3">
                            <label class="text-xs font-medium text-gray-700 block mb-1">In-Hands Date</label>
                            <input id="due-date" type="date" class="w-full px-2 py-1.5 border border-gray-300 rounded text-xs" />
                        </div>
                        <div class="col-span-3">
                            <label class="text-xs font-medium text-gray-700 block mb-1">Strike Date</label>
                            <input id="strike-date" type="date" class="w-full px-2 py-1.5 border border-gray-300 rounded text-xs" />
                        </div>
                    </div>
                </div>

                <!-- Installation Details Section -->
                <div class="form-section">
                    <h2 class="text-sm font-semibold text-gray-800 mb-3">Installation Details</h2>
                    
                    <!-- Venue and Area - Same Row (50/50) -->
                    <div class="grid grid-cols-2 gap-3 mb-3">
                        <div>
                            <label class="text-xs font-medium text-gray-700 block mb-1">Venue</label>
                            <select id="venue-dropdown" class="w-full px-2 py-1.5 border border-gray-300 rounded text-xs">
                                <option value="">Select Venue</option>
                            </select>
                            <input id="venue-input" type="text" class="hidden w-full px-2 py-1.5 border border-gray-300 rounded text-xs mt-1" placeholder="Enter new venue" />
                            <a id="add-venue-link" class="add-link mt-0.5">+ Add New Venue</a>
                        </div>
                        <div>
                            <label class="text-xs font-medium text-gray-700 block mb-1">Area</label>
                            <select id="area-dropdown" class="w-full px-2 py-1.5 border border-gray-300 rounded text-xs">
                                <option value="">Select Area</option>
                            </select>
                            <input id="area-input" type="text" class="hidden w-full px-2 py-1.5 border border-gray-300 rounded text-xs mt-1" placeholder="Enter new area" />
                            <a id="add-area-link" class="add-link mt-0.5">+ Add New Area</a>
                        </div>
                    </div>

                    <!-- Location -->
                    <div class="mb-3">
                        <label class="text-xs font-medium text-gray-700 block mb-1">Location</label>
                        <input id="location" type="text" class="w-full px-2 py-1.5 border border-gray-300 rounded text-xs" />
                    </div>
                </div>

                <!-- File Upload Section -->
                <div class="form-section">
                    <div class="border-2 border-dashed border-gray-300 rounded-lg p-4">
                        <div class="mb-3">
                            <label class="text-xs font-medium text-gray-700 block mb-1">Select Folder</label>
                            <select id="folder-dropdown" class="w-full px-2 py-1.5 border border-gray-300 rounded text-xs">
                                <option value="">Select Folder</option>
                            </select>
                        </div>
                        <div class="flex items-center justify-between">
                            <span class="text-xs text-gray-600">File Upload</span>
                            <label for="file-input" class="bg-white border border-gray-300 px-3 py-1.5 rounded text-xs font-medium hover:bg-gray-50 cursor-pointer">
                                Select File
                            </label>
                            <input id="file-input" type="file" class="hidden" />
                        </div>
                        <div id="file-name" class="text-xs text-gray-500 mt-2"></div>
                    </div>
                </div>

                <!-- Add to Project Button -->
                <div class="flex justify-center mt-4">
                    <button id="add-to-project-btn" class="bg-blue-600 text-white px-6 py-2 rounded text-sm font-medium hover:bg-blue-700">
                        Add to Project
                    </button>
                </div>
            </div>
        </div>

        <!-- Side Drawer Trigger -->
        <div id="drawer-trigger" class="w-12 bg-gray-50 hover:bg-gray-200 transition-colors cursor-pointer flex items-center justify-center" style="box-shadow: -4px 0 12px -4px rgba(0,0,0,0.05)">
            <div class="flex flex-col items-center gap-2">
                <svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" class="text-gray-500"><polyline points="15 18 9 12 15 6"></polyline></svg>
                <span class="text-xs font-medium text-gray-500" style="writing-mode: vertical-rl; text-orientation: mixed; transform: rotate(180deg);">Options</span>
            </div>
        </div>
    </div>

    <!-- Side Drawer for Dropdown Management -->
    <div id="side-drawer" class="fixed top-0 right-0 h-full w-80 bg-white shadow-xl transform translate-x-full z-50">
        <div class="p-4 border-b flex justify-between items-center">
            <h2 class="text-lg font-semibold text-gray-800">Dropdown Options</h2>
            <button id="close-drawer" class="text-gray-500 hover:text-gray-800 text-2xl">&times;</button>
        </div>
        
        <div class="overflow-y-auto" style="height: calc(100% - 73px);">
            <!-- Item Accordion -->
            <div class="accordion-item">
                <div class="accordion-header" data-accordion="item">
                    <span class="text-sm font-medium">Item</span>
                    <svg class="w-4 h-4 transform transition-transform" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7" />
                    </svg>
                </div>
                <div class="accordion-content" id="item-accordion-content">
                    <div id="item-list" class="py-2"></div>
                    <div class="p-4 bg-gray-50">
                        <button id="add-new-item-drawer" class="w-full bg-blue-600 text-white px-3 py-2 rounded text-sm hover:bg-blue-700">+ Add New Item</button>
                    </div>
                </div>
            </div>

            <!-- Material Accordion -->
            <div class="accordion-item">
                <div class="accordion-header" data-accordion="material">
                    <span class="text-sm font-medium">Material</span>
                    <svg class="w-4 h-4 transform transition-transform" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7" />
                    </svg>
                </div>
                <div class="accordion-content" id="material-accordion-content">
                    <div id="material-list" class="py-2"></div>
                    <div class="p-4 bg-gray-50">
                        <button id="add-new-material-drawer" class="w-full bg-blue-600 text-white px-3 py-2 rounded text-sm hover:bg-blue-700">+ Add New Material</button>
                    </div>
                </div>
            </div>

            <!-- Status Accordion -->
            <div class="accordion-item">
                <div class="accordion-header" data-accordion="status">
                    <span class="text-sm font-medium">Status</span>
                    <svg class="w-4 h-4 transform transition-transform" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7" />
                    </svg>
                </div>
                <div class="accordion-content" id="status-accordion-content">
                    <div id="status-list" class="py-2"></div>
                    <div class="p-4 bg-gray-50">
                        <button id="add-new-status-drawer" class="w-full bg-blue-600 text-white px-3 py-2 rounded text-sm hover:bg-blue-700">+ Add New Status</button>
                    </div>
                </div>
            </div>

            <!-- Venue Accordion -->
            <div class="accordion-item">
                <div class="accordion-header" data-accordion="venue">
                    <span class="text-sm font-medium">Venue</span>
                    <svg class="w-4 h-4 transform transition-transform" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7" />
                    </svg>
                </div>
                <div class="accordion-content" id="venue-accordion-content">
                    <div id="venue-list" class="py-2"></div>
                    <div class="p-4 bg-gray-50">
                        <button id="add-new-venue-drawer" class="w-full bg-blue-600 text-white px-3 py-2 rounded text-sm hover:bg-blue-700">+ Add New Venue</button>
                    </div>
                </div>
            </div>

            <!-- Area Accordion -->
            <div class="accordion-item">
                <div class="accordion-header" data-accordion="area">
                    <span class="text-sm font-medium">Area</span>
                    <svg class="w-4 h-4 transform transition-transform" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7" />
                    </svg>
                </div>
                <div class="accordion-content" id="area-accordion-content">
                    <div id="area-list" class="py-2"></div>
                    <div class="p-4 bg-gray-50">
                        <button id="add-new-area-drawer" class="w-full bg-blue-600 text-white px-3 py-2 rounded text-sm hover:bg-blue-700">+ Add New Area</button>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
        let dropdownData = {};
        let driveFolders = [];
        let selectedFile = null;
        let isEditMode = false;
        let originalRowNumber = null;

        if (window.editFormData) {
            isEditMode = true;
            originalRowNumber = window.editFormData.originalRowNumber;
            document.querySelector('h1').textContent = 'Edit Asset';
            document.getElementById('add-to-project-btn').textContent = 'Update Project';
        }

        google.script.run
            .withSuccessHandler(function(data) {
                dropdownData = data;
                populateDropdowns();
                
                if (isEditMode && window.editFormData) {
                    populateFormData(window.editFormData);
                }
            })
            .getDropdownData();

        google.script.run
            .withSuccessHandler(function(folders) {
                driveFolders = folders;
                const folderDropdown = document.getElementById('folder-dropdown');
                folders.forEach(folder => {
                    const option = document.createElement('option');
                    option.value = folder.id;
                    option.textContent = folder.name;
                    folderDropdown.appendChild(option);
                });
            })
            .getDriveFolders();

        function populateDropdowns() {
            ['items', 'materials', 'statuses', 'venues', 'areas'].forEach(field => {
                const dropdown = document.getElementById(`${field === 'items' ? 'item' : field === 'materials' ? 'material' : field === 'statuses' ? 'status' : field === 'venues' ? 'venue' : 'area'}-dropdown`);
                
                if (dropdownData[field]) {
                    dropdownData[field].forEach(value => {
                        const option = document.createElement('option');
                        option.value = value;
                        option.textContent = value;
                        dropdown.appendChild(option);
                    });
                }
            });
        }

        function setupAddLink(fieldName) {
            const addLink = document.getElementById(`add-${fieldName}-link`);
            const dropdown = document.getElementById(`${fieldName}-dropdown`);
            const input = document.getElementById(`${fieldName}-input`);

            addLink.addEventListener('click', function() {
                if (dropdown.classList.contains('hidden')) {
                    dropdown.classList.remove('hidden');
                    input.classList.add('hidden');
                    input.value = '';
                    this.textContent = `+ Add New ${fieldName.charAt(0).toUpperCase() + fieldName.slice(1)}`;
                } else {
                    dropdown.classList.add('hidden');
                    input.classList.remove('hidden');
                    input.focus();
                    this.textContent = 'Cancel';
                }
            });

            input.addEventListener('blur', function() {
                if (this.value.trim()) {
                    addNewValue(fieldName, this.value.trim());
                }
            });

            input.addEventListener('keypress', function(e) {
                if (e.key === 'Enter') {
                    e.preventDefault();
                    if (this.value.trim()) {
                        addNewValue(fieldName, this.value.trim());
                    }
                }
            });
        }

        function addNewValue(fieldName, value) {
            const fieldMap = {
                'item': 'items',
                'material': 'materials',
                'status': 'statuses',
                'venue': 'venues',
                'area': 'areas'
            };

            google.script.run
                .withSuccessHandler(function(result) {
                    if (result.success) {
                        const dropdown = document.getElementById(`${fieldName}-dropdown`);
                        const option = document.createElement('option');
                        option.value = value;
                        option.textContent = value;
                        option.selected = true;
                        dropdown.appendChild(option);

                        dropdown.classList.remove('hidden');
                        const input = document.getElementById(`${fieldName}-input`);
                        input.classList.add('hidden');
                        input.value = '';
                        
                        const addLink = document.getElementById(`add-${fieldName}-link`);
                        addLink.textContent = `+ Add New ${fieldName.charAt(0).toUpperCase() + fieldName.slice(1)}`;
                    } else {
                        alert(result.message || 'Error adding value');
                    }
                })
                .addNewDropdownValue(fieldMap[fieldName], value);
        }

        ['item', 'material', 'status', 'venue', 'area'].forEach(field => {
            setupAddLink(field);
        });

        document.getElementById('file-input').addEventListener('change', function(e) {
            if (e.target.files.length > 0) {
                selectedFile = e.target.files[0];
                document.getElementById('file-name').textContent = selectedFile.name;
            }
        });

        document.getElementById('add-to-project-btn').addEventListener('click', function() {
            const formData = collectFormData();
            
            if (!validateForm(formData)) {
                return;
            }

            const button = this;
            button.disabled = true;
            button.textContent = isEditMode ? 'Updating...' : 'Adding...';
            button.classList.add('opacity-50');

            if (selectedFile) {
                const folderId = document.getElementById('folder-dropdown').value;
                
                if (!folderId) {
                    alert('Please select a folder for the file upload');
                    button.disabled = false;
                    button.textContent = isEditMode ? 'Update Project' : 'Add to Project';
                    button.classList.remove('opacity-50');
                    return;
                }

                if (isEditMode && originalRowNumber) {
                    google.script.run
                        .withSuccessHandler(function(rowData) {
                            const existingId = rowData ? rowData.id : null;
                            const projectCode = rowData.projectCode || 'PROJECT';
                            const assetName = formData.asset.replace(/\s+/g, '_');
                            const fileName = `${existingId}_${projectCode}_${assetName}`;

                            const reader = new FileReader();
                            reader.onload = function(e) {
                                google.script.run
                                    .withSuccessHandler(function(uploadResult) {
                                        if (uploadResult.success) {
                                            formData.artworkUrl = uploadResult.fileUrl;
                                            submitToProject(formData, button);
                                        } else {
                                            alert('Error uploading file: ' + uploadResult.message);
                                            button.disabled = false;
                                            button.textContent = isEditMode ? 'Update Project' : 'Add to Project';
                                            button.classList.remove('opacity-50');
                                        }
                                    })
                                    .uploadFileToDrive(e.target.result, folderId, fileName);
                            };
                            reader.readAsDataURL(selectedFile);
                        })
                        .getRowDataForEdit(originalRowNumber);
                } else {
                    google.script.run
                        .withSuccessHandler(function(nextId) {
                            const assetName = formData.asset.replace(/\s+/g, '_');
                            const fileName = `${nextId}_PROJECT_${assetName}`;

                            const reader = new FileReader();
                            reader.onload = function(e) {
                                google.script.run
                                    .withSuccessHandler(function(uploadResult) {
                                        if (uploadResult.success) {
                                            formData.artworkUrl = uploadResult.fileUrl;
                                            submitToProject(formData, button);
                                        } else {
                                            alert('Error uploading file: ' + uploadResult.message);
                                            button.disabled = false;
                                            button.textContent = isEditMode ? 'Update Project' : 'Add to Project';
                                            button.classList.remove('opacity-50');
                                        }
                                    })
                                    .uploadFileToDrive(e.target.result, folderId, fileName);
                            };
                            reader.readAsDataURL(selectedFile);
                        })
                        .getNextIdForMaterial(formData.material);
                }
            } else {
                submitToProject(formData, button);
            }
        });

        function collectFormData() {
            const formData = {
                item: document.getElementById('item-dropdown').value || document.getElementById('item-input').value,
                material: document.getElementById('material-dropdown').value || document.getElementById('material-input').value,
                asset: document.getElementById('asset-name').value,
                quantity: document.getElementById('quantity').value,
                width: document.getElementById('width').value,
                height: document.getElementById('height').value,
                dieCut: document.getElementById('die-cut').checked,
                doubleSided: document.getElementById('double-sided').checked,
                status: document.getElementById('status-dropdown').value || document.getElementById('status-input').value,
                dueDate: document.getElementById('due-date').value,
                strikeDate: document.getElementById('strike-date').value,
                venue: document.getElementById('venue-dropdown').value || document.getElementById('venue-input').value,
                area: document.getElementById('area-dropdown').value || document.getElementById('area-input').value,
                location: document.getElementById('location').value,
                artworkUrl: ''
            };

            if (isEditMode && originalRowNumber) {
                formData.originalRowNumber = originalRowNumber;
            }

            return formData;
        }

        function validateForm(formData) {
            if (!formData.asset) {
                alert('Please enter an Asset Name');
                return false;
            }
            if (!formData.material) {
                alert('Please select or enter a Material');
                return false;
            }
            if (!formData.quantity || formData.quantity <= 0) {
                alert('Please enter a valid Quantity');
                return false;
            }
            return true;
        }

        function submitToProject(formData, button) {
            google.script.run
                .withSuccessHandler(function(result) {
                    button.disabled = false;
                    button.textContent = isEditMode ? 'Update Project' : 'Add to Project';
                    button.classList.remove('opacity-50');

                    if (result.success) {
                        if (result.isUpdate) {
                            google.script.host.close();
                        } else {
                            resetForm();
                        }
                    } else {
                        alert('Error: ' + result.message);
                    }
                })
                .withFailureHandler(function(error) {
                    button.disabled = false;
                    button.textContent = isEditMode ? 'Update Project' : 'Add to Project';
                    button.classList.remove('opacity-50');
                    alert('Error: ' + error.message);
                })
                .addAssetToProject(formData);
        }

        function resetForm() {
            document.getElementById('item-dropdown').value = '';
            document.getElementById('material-dropdown').value = '';
            document.getElementById('asset-name').value = '';
            document.getElementById('quantity').value = '1';
            document.getElementById('width').value = '';
            document.getElementById('height').value = '';
            document.getElementById('die-cut').checked = false;
            document.getElementById('double-sided').checked = false;
            document.getElementById('status-dropdown').value = '';
            document.getElementById('due-date').value = '';
            document.getElementById('strike-date').value = '';
            document.getElementById('venue-dropdown').value = '';
            document.getElementById('area-dropdown').value = '';
            document.getElementById('location').value = '';
            document.getElementById('folder-dropdown').value = '';
            document.getElementById('file-input').value = '';
            document.getElementById('file-name').textContent = '';
            selectedFile = null;
        }

        function populateFormData(formData) {
            if (!formData) return;

            if (formData.item) document.getElementById('item-dropdown').value = formData.item;
            if (formData.material) document.getElementById('material-dropdown').value = formData.material;
            if (formData.asset) document.getElementById('asset-name').value = formData.asset;
            if (formData.quantity) document.getElementById('quantity').value = formData.quantity;
            if (formData.width) document.getElementById('width').value = formData.width;
            if (formData.height) document.getElementById('height').value = formData.height;
            if (formData.dieCut) document.getElementById('die-cut').checked = formData.dieCut;
            if (formData.doubleSided) document.getElementById('double-sided').checked = formData.doubleSided;
            if (formData.status) document.getElementById('status-dropdown').value = formData.status;
            if (formData.dueDate) document.getElementById('due-date').value = formData.dueDate;
            if (formData.strikeDate) document.getElementById('strike-date').value = formData.strikeDate;
            if (formData.venue) document.getElementById('venue-dropdown').value = formData.venue;
            if (formData.area) document.getElementById('area-dropdown').value = formData.area;
            if (formData.location) document.getElementById('location').value = formData.location;
        }

        const drawer = document.getElementById('side-drawer');
        const drawerTrigger = document.getElementById('drawer-trigger');
        const closeDrawer = document.getElementById('close-drawer');

        drawerTrigger.addEventListener('click', () => {
            drawer.classList.remove('translate-x-full');
        });

        closeDrawer.addEventListener('click', () => {
            drawer.classList.add('translate-x-full');
        });

        document.querySelectorAll('.accordion-header').forEach(header => {
            header.addEventListener('click', function() {
                const accordionType = this.getAttribute('data-accordion');
                const content = document.getElementById(`${accordionType}-accordion-content`);
                const svg = this.querySelector('svg');
                
                content.classList.toggle('active');
                svg.classList.toggle('rotate-180');
                
                if (content.classList.contains('active')) {
                    loadDrawerItems(accordionType);
                }
            });
        });

        function loadDrawerItems(fieldType) {
            const fieldMap = {
                'item': 'items',
                'material': 'materials',
                'status': 'statuses',
                'venue': 'venues',
                'area': 'areas'
            };
            
            const listContainer = document.getElementById(`${fieldType}-list`);
            listContainer.innerHTML = '<div class="text-center py-4 text-sm text-gray-500">Loading...</div>';
            
            const values = dropdownData[fieldMap[fieldType]] || [];
            
            listContainer.innerHTML = '';
            
            values.forEach(value => {
                const itemDiv = document.createElement('div');
                itemDiv.className = 'dropdown-list-item';
                itemDiv.innerHTML = `
                    <input type="text" value="${value}" data-original="${value}" class="flex-1 px-2 py-1 border border-gray-300 rounded text-sm" />
                    <button class="save-btn hidden bg-green-600 text-white rounded hover:bg-green-700">Save</button>
                    <button class="delete-btn bg-red-600 text-white rounded hover:bg-red-700">Delete</button>
                `;
                
                const input = itemDiv.querySelector('input');
                const saveBtn = itemDiv.querySelector('.save-btn');
                const deleteBtn = itemDiv.querySelector('.delete-btn');
                
                input.addEventListener('input', function() {
                    if (this.value !== this.getAttribute('data-original')) {
                        saveBtn.classList.remove('hidden');
                    } else {
                        saveBtn.classList.add('hidden');
                    }
                });
                
                saveBtn.addEventListener('click', function() {
                    const oldValue = input.getAttribute('data-original');
                    const newValue = input.value.trim();
                    
                    if (newValue && newValue !== oldValue) {
                        google.script.run
                            .withSuccessHandler(function(result) {
                                if (result.success) {
                                    input.setAttribute('data-original', newValue);
                                    saveBtn.classList.add('hidden');
                                    
                                    updateFormDropdown(fieldType, oldValue, newValue);
                                    
                                    google.script.run
                                        .withSuccessHandler(function(data) {
                                            dropdownData = data;
                                            populateDropdowns();
                                        })
                                        .getDropdownData();
                                } else {
                                    alert('Error: ' + result.message);
                                }
                            })
                            .updateDropdownValue(fieldMap[fieldType], oldValue, newValue);
                    }
                });
                
                deleteBtn.addEventListener('click', function() {
                    if (confirm(`Delete "${value}"?`)) {
                        google.script.run
                            .withSuccessHandler(function(result) {
                                if (result.success) {
                                    itemDiv.remove();
                                    
                                    const dropdown = document.getElementById(`${fieldType}-dropdown`);
                                    const options = dropdown.querySelectorAll('option');
                                    options.forEach(opt => {
                                        if (opt.value === value) opt.remove();
                                    });
                                    
                                    google.script.run
                                        .withSuccessHandler(function(data) {
                                            dropdownData = data;
                                        })
                                        .getDropdownData();
                                } else {
                                    alert('Error: ' + result.message);
                                }
                            })
                            .deleteDropdownValue(fieldMap[fieldType], value);
                    }
                });
                
                listContainer.appendChild(itemDiv);
            });
        }

        function updateFormDropdown(fieldType, oldValue, newValue) {
            const dropdown = document.getElementById(`${fieldType}-dropdown`);
            const options = dropdown.querySelectorAll('option');
            
            options.forEach(opt => {
                if (opt.value === oldValue) {
                    opt.value = newValue;
                    opt.textContent = newValue;
                }
            });
        }

        ['item', 'material', 'status', 'venue', 'area'].forEach(fieldType => {
            document.getElementById(`add-new-${fieldType}-drawer`).addEventListener('click', function() {
                const newValue = prompt(`Enter new ${fieldType}:`);
                if (newValue && newValue.trim()) {
                    const fieldMap = {
                        'item': 'items',
                        'material': 'materials',
                        'status': 'statuses',
                        'venue': 'venues',
                        'area': 'areas'
                    };
                    
                    google.script.run
                        .withSuccessHandler(function(result) {
                            if (result.success) {
                                loadDrawerItems(fieldType);
                                
                                const dropdown = document.getElementById(`${fieldType}-dropdown`);
                                const option = document.createElement('option');
                                option.value = newValue.trim();
                                option.textContent = newValue.trim();
                                dropdown.appendChild(option);
                                
                                google.script.run
                                    .withSuccessHandler(function(data) {
                                        dropdownData = data;
                                    })
                                    .getDropdownData();
                            } else {
                                alert('Error: ' + result.message);
                            }
                        })
                        .addNewDropdownValue(fieldMap[fieldType], newValue.trim());
                }
            });
        });
    </script>
</body>
</html>';
}

function getCommentDialogHtmlTemplate() {
  return '<!DOCTYPE html>
<html>
<head>
    <base target="_top">
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        body {
            font-family: -apple-system, Inter, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            margin: 0;
            padding: 16px;
            font-size: 13px;
        }
        .comment-box {
            width: 100%;
            min-height: 100px;
            padding: 10px;
            border: 1px solid #d1d5db;
            border-radius: 6px;
            font-size: 12px;
            resize: vertical;
            font-family: inherit;
        }
        .comment-box:focus {
            outline: none;
            border-color: #0D6EFD;
            box-shadow: 0 0 0 2px rgba(99, 102, 241, 0.1);
        }
        .post-btn {
            background-color: #0D6EFD;
            color: white;
            padding: 4px 40px;
            border: none;
            border-radius: 3px;
            font-weight: 400;
            cursor: pointer;
            font-size: 12px;
            transition: background-color 0.2s;
        }
        .post-btn:hover {
            background-color: #357DE9;
        }
        .post-btn:disabled {
            background-color: #9ca3af;
            cursor: not-allowed;
        }
        .asset-info {
            background-color: #fef2f2;
            border-left: 3px solid #FF2093;
            padding: 10px;
            margin-bottom: 12px;
            border-radius: 4px;
        }
        .asset-info p {
            margin: 0;
            margin-bottom: 4px;
            font-size: 12px;
        }
        .asset-info p:last-child {
            margin-bottom: 0;
        }
    </style>
</head>
<body>
    <div class="asset-info">
        <p class="text-gray-700"><strong>Asset:</strong> <span id="asset-name"></span></p>
        <p class="text-gray-700"><strong>ID:</strong> <span id="asset-id"></span></p>
        <p class="text-gray-700"><strong>Status:</strong> <span style="color: #FF2093; font-weight: 600;">Requires Attn</span></p>
    </div>

    <div class="mb-3">
        <textarea 
            id="comment-text" 
            class="comment-box" 
            placeholder="Add your comment..."
        ></textarea>
    </div>

    <div class="flex justify-end">
        <button id="post-btn" class="post-btn">
            Send
        </button>
    </div>

    <script>
        const assetData = window.assetData || {};
        
        document.getElementById('asset-name').textContent = assetData.assetName || 'N/A';
        document.getElementById('asset-id').textContent = assetData.assetId || 'N/A';

        const commentText = document.getElementById('comment-text');
        const postBtn = document.getElementById('post-btn');

        postBtn.addEventListener('click', function() {
            const comment = commentText.value.trim();
            
            if (!comment) {
                alert('Please enter a comment before sending.');
                return;
            }

            postBtn.disabled = true;
            postBtn.textContent = 'Sending...';

            google.script.run
                .withSuccessHandler(function(result) {
                    if (result.success) {
                        google.script.host.close();
                    } else {
                        alert('Error: ' + result.message);
                        postBtn.disabled = false;
                        postBtn.textContent = 'Send';
                    }
                })
                .withFailureHandler(function(error) {
                    alert('Error sending notification: ' + error.message);
                    postBtn.disabled = false;
                    postBtn.textContent = 'Send';
                })
                .sendNotification({
                    row: assetData.row,
                    assetId: assetData.assetId,
                    assetName: assetData.assetName,
                    item: assetData.item,
                    material: assetData.material,
                    comment: comment
                });
        });

        setTimeout(() => {
            commentText.focus();
        }, 100);
    </script>
</body>
</html>';
}

function getDropdownEditorHtmlTemplate() {
  return '<!DOCTYPE html>
<html>
<head>
    <base target="_top">
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
            margin: 0;
            padding: 0;
            background: white;
        }
        .accordion-item {
            border-bottom: 1px solid #e5e7eb;
        }
        .accordion-header {
            cursor: pointer;
            padding: 12px 16px;
            background: #f9fafb;
            display: flex;
            justify-content: space-between;
            align-items: center;
        }
        .accordion-header:hover {
            background: #f3f4f6;
        }
        .accordion-content {
            max-height: 0;
            overflow: hidden;
            transition: max-height 0.3s ease;
        }
        .accordion-content.active {
            max-height: 500px;
            overflow-y: auto;
        }
        .dropdown-list-item {
            display: flex;
            align-items: center;
            gap: 8px;
            padding: 8px 16px;
            border-bottom: 1px solid #f3f4f6;
        }
        .dropdown-list-item input {
            flex: 1;
        }
        .dropdown-list-item button {
            padding: 4px 8px;
            font-size: 12px;
        }
    </style>
</head>
<body>
    <div class="p-4 border-b bg-white">
        <h2 class="text-lg font-semibold text-gray-800">Dropdown Options</h2>
        <p class="text-xs text-gray-500 mt-1">Manage all dropdown lists</p>
    </div>
    
    <div class="overflow-y-auto" style="height: calc(100vh - 73px);">
        <!-- Item Accordion -->
        <div class="accordion-item">
            <div class="accordion-header" data-accordion="item">
                <span class="text-sm font-medium">Item</span>
                <svg class="w-4 h-4 transform transition-transform" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7" />
                </svg>
            </div>
            <div class="accordion-content" id="item-accordion-content">
                <div id="item-list" class="py-2"></div>
                <div class="p-4 bg-gray-50">
                    <button id="add-new-item-drawer" class="w-full bg-blue-600 text-white px-3 py-2 rounded text-sm hover:bg-blue-700">+ Add New Item</button>
                </div>
            </div>
        </div>

        <!-- Material Accordion -->
        <div class="accordion-item">
            <div class="accordion-header" data-accordion="material">
                <span class="text-sm font-medium">Material</span>
                <svg class="w-4 h-4 transform transition-transform" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7" />
                </svg>
            </div>
            <div class="accordion-content" id="material-accordion-content">
                <div id="material-list" class="py-2"></div>
                <div class="p-4 bg-gray-50">
                    <button id="add-new-material-drawer" class="w-full bg-blue-600 text-white px-3 py-2 rounded text-sm hover:bg-blue-700">+ Add New Material</button>
                </div>
            </div>
        </div>

        <!-- Status Accordion -->
        <div class="accordion-item">
            <div class="accordion-header" data-accordion="status">
                <span class="text-sm font-medium">Status</span>
                <svg class="w-4 h-4 transform transition-transform" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7" />
                </svg>
            </div>
            <div class="accordion-content" id="status-accordion-content">
                <div id="status-list" class="py-2"></div>
                <div class="p-4 bg-gray-50">
                    <button id="add-new-status-drawer" class="w-full bg-blue-600 text-white px-3 py-2 rounded text-sm hover:bg-blue-700">+ Add New Status</button>
                </div>
            </div>
        </div>

        <!-- Venue Accordion -->
        <div class="accordion-item">
            <div class="accordion-header" data-accordion="venue">
                <span class="text-sm font-medium">Venue</span>
                <svg class="w-4 h-4 transform transition-transform" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7" />
                </svg>
            </div>
            <div class="accordion-content" id="venue-accordion-content">
                <div id="venue-list" class="py-2"></div>
                <div class="p-4 bg-gray-50">
                    <button id="add-new-venue-drawer" class="w-full bg-blue-600 text-white px-3 py-2 rounded text-sm hover:bg-blue-700">+ Add New Venue</button>
                </div>
            </div>
        </div>

        <!-- Area Accordion -->
        <div class="accordion-item">
            <div class="accordion-header" data-accordion="area">
                <span class="text-sm font-medium">Area</span>
                <svg class="w-4 h-4 transform transition-transform" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7" />
                </svg>
            </div>
            <div class="accordion-content" id="area-accordion-content">
                <div id="area-list" class="py-2"></div>
                <div class="p-4 bg-gray-50">
                    <button id="add-new-area-drawer" class="w-full bg-blue-600 text-white px-3 py-2 rounded text-sm hover:bg-blue-700">+ Add New Area</button>
                </div>
            </div>
        </div>

        <!-- Production Status Accordion -->
        <div class="accordion-item">
            <div class="accordion-header" data-accordion="productionStatus">
                <span class="text-sm font-medium">Production Status</span>
                <svg class="w-4 h-4 transform transition-transform" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke="currentColor">
                    <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 9l-7 7-7-7" />
                </svg>
            </div>
            <div class="accordion-content" id="productionStatus-accordion-content">
                <div id="productionStatus-list" class="py-2"></div>
                <div class="p-4 bg-gray-50">
                    <button id="add-new-productionStatus-drawer" class="w-full bg-blue-600 text-white px-3 py-2 rounded text-sm hover:bg-blue-700">+ Add New Production Status</button>
                </div>
            </div>
        </div>
    </div>

    <script>
        let dropdownData = {};

        google.script.run
            .withSuccessHandler(function(data) {
                dropdownData = data;
            })
            .getDropdownData();

        document.querySelectorAll('.accordion-header').forEach(header => {
            header.addEventListener('click', function() {
                const accordionType = this.getAttribute('data-accordion');
                const content = document.getElementById(`${accordionType}-accordion-content`);
                const svg = this.querySelector('svg');
                
                content.classList.toggle('active');
                svg.classList.toggle('rotate-180');
                
                if (content.classList.contains('active')) {
                    loadDrawerItems(accordionType);
                }
            });
        });

        function loadDrawerItems(fieldType) {
            const fieldMap = {
                'item': 'items',
                'material': 'materials',
                'status': 'statuses',
                'venue': 'venues',
                'area': 'areas',
                'productionStatus': 'productionStatuses'
            };
            
            const listContainer = document.getElementById(`${fieldType}-list`);
            listContainer.innerHTML = '<div class="text-center py-4 text-sm text-gray-500">Loading...</div>';
            
            google.script.run
                .withSuccessHandler(function(data) {
                    dropdownData = data;
                    const values = dropdownData[fieldMap[fieldType]] || [];
                    
                    listContainer.innerHTML = '';
                    
                    values.forEach(value => {
                        const itemDiv = document.createElement('div');
                        itemDiv.className = 'dropdown-list-item';
                        itemDiv.innerHTML = `
                            <input type="text" value="${value}" data-original="${value}" class="flex-1 px-2 py-1 border border-gray-300 rounded text-sm" />
                            <button class="save-btn hidden bg-green-600 text-white rounded hover:bg-green-700">Save</button>
                            <button class="delete-btn bg-red-600 text-white rounded hover:bg-red-700">Delete</button>
                        `;
                        
                        const input = itemDiv.querySelector('input');
                        const saveBtn = itemDiv.querySelector('.save-btn');
                        const deleteBtn = itemDiv.querySelector('.delete-btn');
                        
                        input.addEventListener('input', function() {
                            if (this.value !== this.getAttribute('data-original')) {
                                saveBtn.classList.remove('hidden');
                            } else {
                                saveBtn.classList.add('hidden');
                            }
                        });
                        
                        saveBtn.addEventListener('click', function() {
                            const oldValue = input.getAttribute('data-original');
                            const newValue = input.value.trim();
                            
                            if (newValue && newValue !== oldValue) {
                                google.script.run
                                    .withSuccessHandler(function(result) {
                                        if (result.success) {
                                            input.setAttribute('data-original', newValue);
                                            saveBtn.classList.add('hidden');
                                            
                                            google.script.run
                                                .withSuccessHandler(function(data) {
                                                    dropdownData = data;
                                                })
                                                .getDropdownData();
                                        } else {
                                            alert('Error: ' + result.message);
                                        }
                                    })
                                    .updateDropdownValue(fieldMap[fieldType], oldValue, newValue);
                            }
                        });
                        
                        deleteBtn.addEventListener('click', function() {
                            if (confirm(`Delete "${value}"?`)) {
                                google.script.run
                                    .withSuccessHandler(function(result) {
                                        if (result.success) {
                                            itemDiv.remove();
                                            
                                            google.script.run
                                                .withSuccessHandler(function(data) {
                                                    dropdownData = data;
                                                })
                                                .getDropdownData();
                                        } else {
                                            alert('Error: ' + result.message);
                                        }
                                    })
                                    .deleteDropdownValue(fieldMap[fieldType], value);
                            }
                        });
                        
                        listContainer.appendChild(itemDiv);
                    });
                })
                .getDropdownData();
        }

        ['item', 'material', 'status', 'venue', 'area', 'productionStatus'].forEach(fieldType => {
            document.getElementById(`add-new-${fieldType}-drawer`).addEventListener('click', function() {
                const displayName = fieldType === 'productionStatus' ? 'production status' : fieldType;
                const newValue = prompt(`Enter new ${displayName}:`);
                if (newValue && newValue.trim()) {
                    const fieldMap = {
                        'item': 'items',
                        'material': 'materials',
                        'status': 'statuses',
                        'venue': 'venues',
                        'area': 'areas',
                        'productionStatus': 'productionStatuses'
                    };
                    
                    google.script.run
                        .withSuccessHandler(function(result) {
                            if (result.success) {
                                loadDrawerItems(fieldType);
                                
                                google.script.run
                                    .withSuccessHandler(function(data) {
                                        dropdownData = data;
                                    })
                                    .getDropdownData();
                            } else {
                                alert('Error: ' + result.message);
                            }
                        })
                        .addNewDropdownValue(fieldMap[fieldType], newValue.trim());
                }
            });
        });
    </script>
</body>
</html>';
}

function getReorderAssetFormHtmlTemplate() {
  return '<!DOCTYPE html>
<html>
<head>
    <base target="_top">
    <script src="https://cdn.tailwindcss.com"></script>
    <style>
        input[type=number]::-webkit-inner-spin-button,
        input[type=number]::-webkit-outer-spin-button {
            -webkit-appearance: none;
            margin: 0;
        }
        input[type=number] {
            -moz-appearance: textfield;
        }
        
        label {
            font-size: 0.875rem;
            font-weight: 500;
            margin-bottom: 0.5rem;
        }
        
        input, select {
            font-size: 0.875rem;
            padding: 0.5rem 0.75rem;
        }
        
        .hidden {
            display: none;
        }
    </style>
</head>
<body class="bg-gray-100 p-6">
    <div class="max-w-2xl mx-auto bg-white rounded-lg shadow-lg p-6">
        <h1 class="text-2xl font-semibold text-gray-800 mb-6">Reorder Asset</h1>

        <div class="mb-6">
            <label class="text-sm font-medium text-gray-700 block mb-2">Asset</label>
            <select id="asset-dropdown" class="w-full px-3 py-2 border border-gray-300 rounded text-sm" disabled>
                <option value="">Select Asset</option>
            </select>
            <div id="loading-indicator" class="text-xs text-gray-500 mt-2 flex items-center">
                <svg class="animate-spin h-4 w-4 mr-2 text-blue-600" xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24">
                    <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle>
                    <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"></path>
                </svg>
                Loading assets...
            </div>
        </div>

        <div class="mb-6">
            <label class="text-sm font-medium text-gray-700 block mb-2">Qty</label>
            <input id="quantity" type="number" min="1" placeholder="1" class="w-full px-3 py-2 border border-gray-300 rounded text-sm" />
        </div>

        <div class="flex justify-center mt-8">
            <button id="add-to-project-btn" class="bg-blue-600 text-white px-8 py-3 rounded text-base font-medium hover:bg-blue-700">
                Add to Project
            </button>
        </div>
    </div>

    <script>
        let availableAssets = [];

        google.script.run
            .withSuccessHandler(function(assets) {
                availableAssets = assets;
                populateAssetDropdown();
                
                document.getElementById('loading-indicator').classList.add('hidden');
                document.getElementById('asset-dropdown').disabled = false;
            })
            .withFailureHandler(function(error) {
                console.error('Error loading assets:', error);
                
                const loadingIndicator = document.getElementById('loading-indicator');
                loadingIndicator.innerHTML = '<span class="text-red-600">Error loading assets. Please refresh and try again.</span>';
                
                alert('Error loading assets: ' + error.message);
            })
            .getAvailableAssetsForReorder();

        function populateAssetDropdown() {
            const dropdown = document.getElementById('asset-dropdown');
            
            availableAssets.forEach(asset => {
                const option = document.createElement('option');
                option.value = asset.name;
                option.textContent = asset.name;
                option.dataset.rowNumber = asset.rowNumber;
                dropdown.appendChild(option);
            });
        }

        document.getElementById('add-to-project-btn').addEventListener('click', function() {
            const selectedAsset = document.getElementById('asset-dropdown').value;
            const quantity = document.getElementById('quantity').value;

            if (!selectedAsset) {
                alert('Please select an asset to reorder');
                return;
            }

            if (!quantity || quantity <= 0) {
                alert('Please enter a valid quantity');
                return;
            }

            const button = this;
            button.disabled = true;
            button.textContent = 'Adding...';
            button.classList.add('opacity-50');

            const selectedOption = document.getElementById('asset-dropdown').selectedOptions[0];
            const rowNumber = selectedOption.dataset.rowNumber;

            const reorderData = {
                assetName: selectedAsset,
                quantity: quantity,
                originalRowNumber: rowNumber
            };

            google.script.run
                .withSuccessHandler(function(result) {
                    button.disabled = false;
                    button.textContent = 'Add to Project';
                    button.classList.remove('opacity-50');

                    if (result.success) {
                        document.getElementById('asset-dropdown').value = '';
                        document.getElementById('quantity').value = '';
                    } else {
                        alert('Error: ' + result.message);
                    }
                })
                .withFailureHandler(function(error) {
                    button.disabled = false;
                    button.textContent = 'Add to Project';
                    button.classList.remove('opacity-50');
                    alert('Error: ' + error.message);
                })
                .reorderAsset(reorderData);
        });
    </script>
</body>
</html>';
}
