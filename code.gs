function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Paniniwala at Pananampalataya - Module')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

const SHEET_ID = '1dN1alYrLlcyqHwb3PnKUP0dqboTFqoSv30ZHByDrES4';
const SHEET_NAME = 'Sheet1'; // Default sheet name, can be changed if needed

function getSheet() {
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME) || SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
}

function loginUser(id) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();

  // Assuming row 0 is header: ID, Name, Score, Attempts
  for (let i = 1; i < data.length; i++) {
    // toString to ensure string comparison
    if (data[i][0].toString().toLowerCase() === id.toString().toLowerCase()) {
      return {
        success: true,
        name: data[i][1],
        score: data[i][2],
        attempts: data[i][3] || 0,
        row: i + 1
      };
    }
  }

  return { success: false, message: 'ID not found.' };
}

function saveScore(id, score) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toString().toLowerCase() === id.toString().toLowerCase()) {
      const row = i + 1;
      let currentAttempts = data[i][3] || 0;

      if (currentAttempts >= 3) {
        return { success: false, message: 'Maximum attempts reached.' };
      }

      // Update score (Column C - index 3) and attempts (Column D - index 4)
      sheet.getRange(row, 3).setValue(score);
      sheet.getRange(row, 4).setValue(currentAttempts + 1);

      return { success: true, attempts: currentAttempts + 1 };
    }
  }

  return { success: false, message: 'ID not found during save.' };
}
