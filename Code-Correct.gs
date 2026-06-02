const SPREADSHEET_ID = '1fFZHJMOypwFJTSYqCh0w4gnvHW2LXYBWKIhOLwwKqGk';
const DATA_RANGE = 'A1:P100';
const SETTINGS_SHEET_NAME = 'Settings';
const TYPES_SHEET_NAME = 'Types';

function getAllSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = ss.getSheets();
  const sheetNames = [];
  
  sheets.forEach(sheet => {
    const name = sheet.getName();
    if (name !== SETTINGS_SHEET_NAME && name !== TYPES_SHEET_NAME) {
      sheetNames.push(name);
    }
  });
  
  return sheetNames;
}

function getTypesList() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(TYPES_SHEET_NAME);
  const data = sheet.getRange('A:B').getValues();
  
  const types = {};
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][1]) {
      types[data[i][0].toString()] = data[i][1];
    }
  }
  return types;
}

function getCampaigns(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getRange(DATA_RANGE).getValues();
  const backgrounds = sheet.getRange(DATA_RANGE).getBackgrounds();
  
  const campaigns = [];
  for (let i = 2; i < data.length; i++) {
    // Skip rows with #808080 background color (sub-headers)
    if (backgrounds[i][0] === '#808080') {
      continue;
    }
    
    if (data[i][0] && !data[i][0].toString().toLowerCase().includes('total')) {
      campaigns.push(data[i]);
    }
  }
  
  return campaigns;
}

function getSummary(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const settingsSheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
  const settingsData = settingsSheet.getRange('A:B').getValues();
  
  let targetUnits = 0;
  for (let i = 1; i < settingsData.length; i++) {
    if (settingsData[i][0] === sheetName) {
      targetUnits = parseFloat(settingsData[i][1]) || 0;
      break;
    }
  }
  
  return { targetUnits: targetUnits };
}

function getDashboardSummary() {
  const sheetNames = getAllSheets();
  const types = getTypesList();
  
  let totalSpend = 0;
  let totalUnits = 0;
  const locations = [];
  const byType = {};
  const byCategory = {};
  
  sheetNames.forEach(sheetName => {
    const campaigns = getCampaigns(sheetName);
    const summary = getSummary(sheetName);
    
    let locationSpend = 0;
    
    campaigns.forEach(row => {
      if (!row[0] || row[0].toString().toLowerCase().includes('total')) return;
      
      // Column P (index 15) is the Annual Total
      const spend = parseFloat(row[15]) || 0;
      const typeNum = row[2];
      const category = row[1];
      
      locationSpend += spend;
      totalSpend += spend;
      
      const typeName = types[typeNum] || typeNum || 'Unknown';
      byType[typeName] = (byType[typeName] || 0) + spend;
      byCategory[category] = (byCategory[category] || 0) + spend;
    });
    
    const units = summary.targetUnits || 0;
    totalUnits += units;
    
    const spendPerUnit = units > 0 ? locationSpend / units : 0;
    
    locations.push({
      name: sheetName,
      spend: locationSpend,
      units: units,
      spendPerUnit: spendPerUnit
    });
  });
  
  return {
    totalSpend: totalSpend,
    totalUnits: totalUnits,
    locations: locations,
    byType: byType,
    byCategory: byCategory
  };
}

function updateSpend(sheetName, campaignName, months) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  const data = sheet.getRange(DATA_RANGE).getValues();
  
  // Find the actual row by searching for the campaign name (case-insensitive)
  let actualRowIndex = -1;
  const campaignNameLower = campaignName.toString().toLowerCase().trim();
  for (let i = 2; i < data.length; i++) {
    if (data[i][0]) {
      const cellNameLower = data[i][0].toString().toLowerCase().trim();
      if (cellNameLower === campaignNameLower) {
        actualRowIndex = i;
        break;
      }
    }
  }
  
  if (actualRowIndex === -1) {
    return { success: false, error: 'Campaign not found' };
  }
  
  // Columns D-O (indices 3-14) are the 12 months: Jan through Dec
  for (let i = 0; i < 12; i++) {
    data[actualRowIndex][3 + i] = months[i];
  }
  
  // Column P (index 15) is auto-calculated Annual Total - don't touch it
  
  sheet.getRange(DATA_RANGE).setValues(data);
  
  return { success: true, message: 'Campaign updated' };
}

function doGet(e) {
  const action = e.parameter.action;
  const sheet = e.parameter.sheet;
  
  let result;
  
  switch (action) {
    case 'getSheets':
      result = getAllSheets();
      break;
    case 'getCampaigns':
      result = getCampaigns(sheet);
      break;
    case 'getSummary':
      result = getSummary(sheet);
      break;
    case 'getDashboardSummary':
      result = getDashboardSummary();
      break;
    default:
      result = { error: 'Unknown action' };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const params = JSON.parse(e.postData.contents);
  
  let result;
  
  try {
    switch (params.action) {
      case 'updateSpend':
        result = updateSpend(params.sheet, params.campaignName, params.months);
        break;
      default:
        result = { error: 'Unknown action' };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}
