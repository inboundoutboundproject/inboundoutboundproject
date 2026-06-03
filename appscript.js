// Google Apps Script - Deploy as Web App
// File: Code.gs

const SHEET_NAME = 'ExchangeData';

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  const action = e.parameter.action || e.parameter.get?.('action');

  try {
    let result;
    switch (action) {
      case 'loadData':
        result = loadData();
        break;
      case 'addRecord':
        result = addRecord(JSON.parse(e.parameter.data || e.postData.contents));
        break;
      case 'updateRecord':
        result = updateRecord(
          parseInt(e.parameter.index),
          JSON.parse(e.parameter.data || e.postData.contents)
        );
        break;
      case 'deleteRecord':
        result = deleteRecord(parseInt(e.parameter.index));
        break;
      default:
        result = { success: false, error: 'Invalid action' };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        error: error.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Initialize headers
    sheet.getRange(1, 1, 1, 20).setValues([[
      'ID', 'Name', 'Student ID', 'Email', 'Phone', 'Passport Number',
      'Residence', 'Major', 'Year', 'Adviser', 'Country', 'University',
      'Project', 'Funder', 'Budget', 'Start Date', 'End Date',
      'Form Type', 'Details', 'Timestamp'
    ]]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function loadData() {
  const sheet = getSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return { success: true, data: [] };
  }

  const data = sheet.getRange(2, 1, lastRow - 1, 20).getValues();

  const records = data.map(row => ({
    id: row[0],
    name: row[1],
    studentId: row[2],
    email: row[3],
    phone: row[4],
    passportNumber: row[5],
    residence: row[6],
    major: row[7],
    year: row[8],
    adviser: row[9],
    country: row[10],
    university: row[11],
    project: row[12],
    funder: row[13],
    budget: row[14],
    fromdate: row[15],
    todate: row[16],
    formType: row[17],
    details: row[18],
    timestamp: row[19]
  }));

  return { success: true, data: records };
}

function addRecord(record) {
  const sheet = getSheet();
  const id = Utilities.getUuid();
  const timestamp = new Date().toISOString();

  const rowData = [
    id,
    record.name || '',
    record.studentId || '',
    record.email || '',
    record.phone || '',
    record.passportNumber || '',
    record.residence || '',
    record.major || '',
    record.year || '',
    record.adviser || '',
    record.country || '',
    record.university || '',
    record.project || '',
    record.funder || '',
    record.budget || '',
    record.fromdate || '',
    record.todate || '',
    record.formType || 'inbound',
    record.details || '',
    timestamp
  ];

  sheet.appendRow(rowData);

  return {
    success: true,
    record: { ...record, id, timestamp }
  };
}

function updateRecord(index, record) {
  const sheet = getSheet();
  const rowNumber = index + 2; // +2 because: array is 0-indexed, and row 1 is headers

  if (rowNumber > sheet.getLastRow()) {
    return { success: false, error: 'Record not found' };
  }

  const timestamp = new Date().toISOString();

  const rowData = [
    record.id || sheet.getRange(rowNumber, 1).getValue(),
    record.name || '',
    record.studentId || '',
    record.email || '',
    record.phone || '',
    record.passportNumber || '',
    record.residence || '',
    record.major || '',
    record.year || '',
    record.adviser || '',
    record.country || '',
    record.university || '',
    record.project || '',
    record.funder || '',
    record.budget || '',
    record.fromdate || '',
    record.todate || '',
    record.formType || 'inbound',
    record.details || '',
    timestamp
  ];

  sheet.getRange(rowNumber, 1, 1, 20).setValues([rowData]);

  return { success: true };
}

function deleteRecord(index) {
  try {
    const sheet = getSheet();
    const rowNumber = index + 2; // +2 because: array is 0-indexed, and row 1 is headers
    
    if (rowNumber > sheet.getLastRow()) {
      return { success: false, error: 'Record not found' };
    }
    
    sheet.deleteRow(rowNumber);
    
    return { success: true };
  } catch (error) {
    return { 
      success: false, 
      error: error.toString() 
    };
  }
}
