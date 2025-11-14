/**
 * PROJECTS HQ - MAIN SYSTEM
 * Central hub for project creation and management
 * 
 * @OnlyCurrentDoc
 */

// ============================================================================
// GLOBAL CONFIGURATION
// ============================================================================

const SYSTEM_CONFIG = {
  PARENT_FOLDER_ID: '19trzOWC-Orgb4gaDhHT7yMxG6MB-5mOZ',
  DEFAULT_EMAIL: 'projects@nichbranding.com',
  VERSION: '1.0.0'
};

// ============================================================================
// INITIALIZATION & MENU
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Projects HQ')
    .addItem('➕ Create New Project', 'showProjectSidebar')
    .addSeparator()
    .addItem('🔄 Refresh Data', 'refreshProjectList')
    .addItem('⚙️ Settings', 'showSettings')
    .addToUi();
  
  initializeProjectHQ();
}

function initializeProjectHQ() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create or get Projects List sheet (main visible sheet)
  let projectsSheet = ss.getSheetByName('Projects List');
  if (!projectsSheet) {
    projectsSheet = createProjectsListSheet(ss);
  }
  
  // Create HQ Config sheet if doesn't exist
  let configSheet = ss.getSheetByName('HQ Config');
  if (!configSheet) {
    configSheet = createHQConfigSheet(ss);
  }
  
  // Create team lists
  initializeTeamLists(ss);
  
  // Set Projects List as active sheet
  ss.setActiveSheet(projectsSheet);
}

// ============================================================================
// PROJECTS LIST SHEET (Main Visible Sheet)
// ============================================================================

function createProjectsListSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('Projects List', 0);
  
  const headers = [
    'Project Number', 'Project Name', 'Client', 'Description', 
    'Project Managers', 'Project Status', 'Email Notifications',
    'Start Date', 'In-Hands Date',
    'Invoice Amount', 'Invoice Status', 'Invoice Payment Link',
    'Production Tracker'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format header
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11)
    .setWrap(false);
  
  // Set column widths
  sheet.setColumnWidth(1, 130);  // Project Number
  sheet.setColumnWidth(2, 150);  // Project Name
  sheet.setColumnWidth(3, 120);  // Client
  sheet.setColumnWidth(4, 200);  // Description
  sheet.setColumnWidth(5, 120);  // Project Managers
  sheet.setColumnWidth(6, 110);  // Project Status
  sheet.setColumnWidth(7, 180);  // Email Notifications
  sheet.setColumnWidth(8, 100);  // Start Date
  sheet.setColumnWidth(9, 100);  // In-Hands Date
  sheet.setColumnWidth(10, 110); // Invoice Amount
  sheet.setColumnWidth(11, 110); // Invoice Status
  sheet.setColumnWidth(12, 200); // Invoice Payment Link
  sheet.setColumnWidth(13, 200); // Production Tracker
  
  sheet.setFrozenRows(1);
  
  // Add data validation for Invoice Status
  const invoiceStatusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Not Ready', 'Ready to Send', 'Sent', 'Overdue', 'Paid'], true)
    .setAllowInvalid(false)
    .build();
  
  // Add data validation for Project Status
  const projectStatusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Planning', 'Active', 'On Hold', 'Completed', 'Archived'], true)
    .setAllowInvalid(false)
    .build();
  
  // Apply to first 1000 rows (will apply to new rows)
  sheet.getRange('K2:K1000').setDataValidation(invoiceStatusRule);
  sheet.getRange('F2:F1000').setDataValidation(projectStatusRule);
  
  return sheet;
}

// ============================================================================
// HQ CONFIG SHEET
// ============================================================================

function createHQConfigSheet(spreadsheet) {
  const sheet = spreadsheet.insertSheet('HQ Config');
  
  const configData = [
    ['PROJECTS HQ CONFIGURATION', 'Value', 'Description'],
    ['', '', ''],
    ['System Settings', '', ''],
    ['Parent Projects Folder ID', SYSTEM_CONFIG.PARENT_FOLDER_ID, 'Main Google Drive folder for all projects'],
    ['Default Alert Email', SYSTEM_CONFIG.DEFAULT_EMAIL, 'Always included in project alerts'],
    ['System Version', SYSTEM_CONFIG.VERSION, 'Current version']
  ];
  
  sheet.getRange(1, 1, configData.length, 3).setValues(configData);
  
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 350);
  sheet.setColumnWidth(3, 400);
  
  sheet.getRange('A1:C1')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(12);
  
  sheet.getRange('A3:C3').setBackground('#e8f0fe').setFontWeight('bold');
  sheet.getRange('B2:B10').setBackground('#f8f9fa');
  
  sheet.hideSheet();
  
  return sheet;
}

// ============================================================================
// TEAM LISTS
// ============================================================================

function initializeTeamLists(spreadsheet) {
  const teamMembers = ['Diana', 'Heather'];
  
  let pmSheet = spreadsheet.getSheetByName('ProjectManagersList');
  if (!pmSheet) {
    pmSheet = spreadsheet.insertSheet('ProjectManagersList');
    pmSheet.getRange(1, 1).setValue('Project Managers');
    const data = teamMembers.map(m => [m]);
    pmSheet.getRange(2, 1, data.length, 1).setValues(data);
    pmSheet.hideSheet();
  }
}

function getProjectManagers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('ProjectManagersList');
  
  if (!sheet) return ['Diana', 'Heather'];
  
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return ['Diana', 'Heather'];
  
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return values.filter(row => row[0] && row[0].toString().trim() !== "")
                .map(row => row[0].toString().trim());
}

// ============================================================================
// CONFIGURATION HELPERS
// ============================================================================

function getHQConfig() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('HQ Config');
  
  if (!configSheet) {
    return {
      PARENT_FOLDER_ID: SYSTEM_CONFIG.PARENT_FOLDER_ID,
      DEFAULT_EMAIL: SYSTEM_CONFIG.DEFAULT_EMAIL
    };
  }
  
  return {
    PARENT_FOLDER_ID: configSheet.getRange('B4').getValue() || SYSTEM_CONFIG.PARENT_FOLDER_ID,
    DEFAULT_EMAIL: configSheet.getRange('B5').getValue() || SYSTEM_CONFIG.DEFAULT_EMAIL,
    VERSION: configSheet.getRange('B6').getValue() || SYSTEM_CONFIG.VERSION
  };
}

// ============================================================================
// UI FUNCTIONS
// ============================================================================

function showProjectSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('ProjectSidebar')
    .setTitle('Start a Project')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

function showSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = ss.getSheetByName('HQ Config');
  if (config) {
    config.showSheet();
    ss.setActiveSheet(config);
  }
}

function refreshProjectList() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Projects list refreshed!', 'Projects HQ', 3);
}

// ============================================================================
// PROJECT ADDITION TO SHEET
// ============================================================================

function addProjectToSheet(projectData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Projects List');
  
  if (!sheet) {
    throw new Error('Projects List sheet not found');
  }
  
  // Create tracker link
  const trackerLink = `=HYPERLINK("${projectData.assetTrackerUrl}", "Production Tracker - ${projectData.projectNumber}")`;
  
  const rowData = [
    projectData.projectNumber,
    projectData.projectName,
    projectData.client,
    projectData.description,
    projectData.projectManagers,
    'Active', // Project Status (default) - now in column 6
    projectData.emailNotifications,
    projectData.startDate || '',
    projectData.inHandsDate || '',
    '', // Invoice Amount (manual)
    '', // Invoice Status (manual)
    '', // Invoice Payment Link (manual)
    trackerLink
  ];
  
  sheet.appendRow(rowData);
  
  // Format the new row
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 10).setNumberFormat('$#,##0.00'); // Invoice Amount format
  
  return lastRow;
}
