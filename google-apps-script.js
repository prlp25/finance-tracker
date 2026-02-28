// ============================================
// POLEDGER - Google Apps Script (v3)
// Paste this entire code into your Google Sheet's Apps Script editor
// Supports: Expenses, Deposits, Settings, Photo Uploads
// ============================================

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'getAll';
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === 'getAll') {
    var expSheet = ss.getSheetByName('Expenses');
    var depSheet = ss.getSheetByName('Deposits');
    var setSheet = ss.getSheetByName('Settings');

    // Get expenses (12 columns)
    var expenses = [];
    if (expSheet && expSheet.getLastRow() > 1) {
      var data = expSheet.getRange(2, 1, expSheet.getLastRow() - 1, 12).getValues();
      for (var i = 0; i < data.length; i++) {
        if (data[i][0] === '' || data[i][0] === null) continue;
        expenses.push({
          id: data[i][0],
          date: data[i][1],
          description: data[i][2],
          amount: data[i][3],
          category: data[i][4],
          fundSource: data[i][5],
          user: data[i][6],
          reimbursable: data[i][7] === true || data[i][7] === 'TRUE',
          reimbTo: data[i][8] || null,
          reimbPct: data[i][9] || 50,
          reimbursed: data[i][10] === true || data[i][10] === 'TRUE',
          notes: data[i][11] || ''
        });
      }
    }

    // Get deposits (8 columns)
    var deposits = [];
    if (depSheet && depSheet.getLastRow() > 1) {
      var depData = depSheet.getRange(2, 1, depSheet.getLastRow() - 1, 8).getValues();
      for (var j = 0; j < depData.length; j++) {
        if (depData[j][0] === '' || depData[j][0] === null) continue;
        deposits.push({
          id: depData[j][0],
          date: depData[j][1],
          amount: depData[j][2],
          account: depData[j][3],
          description: depData[j][4],
          user: depData[j][5],
          photoUrl: depData[j][6] || '',
          notes: depData[j][7] || ''
        });
      }
    }

    // Get settings
    var settings = { savingsStart: 50000, opexStart: 30000 };
    if (setSheet && setSheet.getLastRow() >= 2) {
      var setData = setSheet.getRange(2, 1, 1, 2).getValues();
      settings.savingsStart = setData[0][0] || 50000;
      settings.opexStart = setData[0][1] || 30000;
    }

    return ContentService.createTextOutput(JSON.stringify({
      expenses: expenses,
      deposits: deposits,
      settings: settings
    })).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var expSheet = ss.getSheetByName('Expenses');
  var depSheet = ss.getSheetByName('Deposits');
  var setSheet = ss.getSheetByName('Settings');

  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action;

    // ---- EXPENSE ACTIONS ----

    if (action === 'addExpense') {
      var exp = payload.expense;
      expSheet.appendRow([
        exp.id, exp.date, exp.description, exp.amount, exp.category,
        exp.fundSource, exp.user, exp.reimbursable, exp.reimbTo || '',
        exp.reimbPct, exp.reimbursed, exp.notes || ''
      ]);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'editExpense') {
      var upd = payload.expense;
      if (expSheet.getLastRow() > 1) {
        var ids = expSheet.getRange(2, 1, expSheet.getLastRow() - 1, 1).getValues();
        for (var k = 0; k < ids.length; k++) {
          if (ids[k][0] == upd.id) {
            var row = k + 2;
            expSheet.getRange(row, 1, 1, 12).setValues([[
              upd.id, upd.date, upd.description, upd.amount, upd.category,
              upd.fundSource, upd.user, upd.reimbursable, upd.reimbTo || '',
              upd.reimbPct, upd.reimbursed, upd.notes || ''
            ]]);
            return ContentService.createTextOutput(JSON.stringify({ success: true }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ error: 'Not found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'toggleReimbursed') {
      var targetId = payload.id;
      if (expSheet.getLastRow() > 1) {
        var ids2 = expSheet.getRange(2, 1, expSheet.getLastRow() - 1, 1).getValues();
        for (var i = 0; i < ids2.length; i++) {
          if (ids2[i][0] == targetId) {
            var row2 = i + 2;
            var current = expSheet.getRange(row2, 11).getValue();
            expSheet.getRange(row2, 11).setValue(current === true || current === 'TRUE' ? false : true);
            return ContentService.createTextOutput(JSON.stringify({ success: true }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ error: 'Not found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'deleteExpense') {
      var delId = payload.id;
      if (expSheet.getLastRow() > 1) {
        var allIds = expSheet.getRange(2, 1, expSheet.getLastRow() - 1, 1).getValues();
        for (var j = 0; j < allIds.length; j++) {
          if (allIds[j][0] == delId) {
            expSheet.deleteRow(j + 2);
            return ContentService.createTextOutput(JSON.stringify({ success: true }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ error: 'Not found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ---- DEPOSIT ACTIONS ----

    if (action === 'addDeposit') {
      var dep = payload.deposit;
      var photoUrl = '';

      // Upload photo to Google Drive if base64 data is provided
      if (dep.photoBase64 && dep.photoBase64.length > 0) {
        photoUrl = uploadPhotoToDrive(dep.photoBase64, dep.photoName || 'deposit.jpg');
      }

      depSheet.appendRow([
        dep.id, dep.date, dep.amount, dep.account,
        dep.description || '', dep.user, photoUrl, dep.notes || ''
      ]);

      return ContentService.createTextOutput(JSON.stringify({ success: true, photoUrl: photoUrl }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'editDeposit') {
      var updDep = payload.deposit;
      if (depSheet.getLastRow() > 1) {
        var depIds = depSheet.getRange(2, 1, depSheet.getLastRow() - 1, 1).getValues();
        for (var m = 0; m < depIds.length; m++) {
          if (depIds[m][0] == updDep.id) {
            var depRow = m + 2;
            // Keep existing photo URL if no new photo
            var existingPhotoUrl = depSheet.getRange(depRow, 7).getValue() || '';
            var newPhotoUrl = existingPhotoUrl;
            if (updDep.photoBase64 && updDep.photoBase64.length > 0) {
              newPhotoUrl = uploadPhotoToDrive(updDep.photoBase64, updDep.photoName || 'deposit.jpg');
            }
            depSheet.getRange(depRow, 1, 1, 8).setValues([[
              updDep.id, updDep.date, updDep.amount, updDep.account,
              updDep.description || '', updDep.user, newPhotoUrl, updDep.notes || ''
            ]]);
            return ContentService.createTextOutput(JSON.stringify({ success: true, photoUrl: newPhotoUrl }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ error: 'Not found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'deleteDeposit') {
      var delDepId = payload.id;
      if (depSheet.getLastRow() > 1) {
        var allDepIds = depSheet.getRange(2, 1, depSheet.getLastRow() - 1, 1).getValues();
        for (var n = 0; n < allDepIds.length; n++) {
          if (allDepIds[n][0] == delDepId) {
            depSheet.deleteRow(n + 2);
            return ContentService.createTextOutput(JSON.stringify({ success: true }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ error: 'Not found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ---- SETTINGS ----

    if (action === 'updateSettings') {
      var s = payload.settings;
      setSheet.getRange(2, 1, 1, 2).setValues([[s.savingsStart, s.opexStart]]);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({ error: 'Unknown action' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- GOOGLE DRIVE PHOTO UPLOAD ----

function getOrCreateFolder(folderName) {
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}

function uploadPhotoToDrive(base64Data, fileName) {
  var folder = getOrCreateFolder('PoLedger Uploads');

  // Remove data URL prefix if present (e.g., "data:image/jpeg;base64,")
  var cleanBase64 = base64Data;
  if (cleanBase64.indexOf(',') !== -1) {
    cleanBase64 = cleanBase64.split(',')[1];
  }

  // Detect MIME type from data URL or default to jpeg
  var mimeType = 'image/jpeg';
  if (base64Data.indexOf('data:') === 0) {
    var mimeMatch = base64Data.match(/data:([^;]+);/);
    if (mimeMatch) mimeType = mimeMatch[1];
  }

  var blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), mimeType, fileName);
  var file = folder.createFile(blob);

  // Make the file viewable by anyone with the link
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file.getUrl();
}

// Run this ONCE to set up the sheets with proper headers
// If upgrading from v2, run setupSheets again to add the Deposits sheet
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Expenses sheet (12 columns)
  var expSheet = ss.getSheetByName('Expenses');
  if (!expSheet) {
    expSheet = ss.insertSheet('Expenses');
  }
  expSheet.getRange(1, 1, 1, 12).setValues([[
    'ID', 'Date', 'Description', 'Amount', 'Category',
    'Fund Source', 'User', 'Reimbursable', 'Reimb To',
    'Reimb %', 'Reimbursed', 'Notes'
  ]]);
  expSheet.getRange(1, 1, 1, 12).setFontWeight('bold');

  // Deposits sheet (8 columns) — NEW in v3
  var depSheet = ss.getSheetByName('Deposits');
  if (!depSheet) {
    depSheet = ss.insertSheet('Deposits');
  }
  depSheet.getRange(1, 1, 1, 8).setValues([[
    'ID', 'Date', 'Amount', 'Account', 'Description', 'User', 'PhotoURL', 'Notes'
  ]]);
  depSheet.getRange(1, 1, 1, 8).setFontWeight('bold');

  // Settings sheet (2 columns)
  var setSheet = ss.getSheetByName('Settings');
  if (!setSheet) {
    setSheet = ss.insertSheet('Settings');
  }
  setSheet.getRange(1, 1, 1, 2).setValues([['Savings Start', 'Opex Start']]);
  setSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
  if (setSheet.getLastRow() < 2) {
    setSheet.getRange(2, 1, 1, 2).setValues([[50000, 30000]]);
  }

  SpreadsheetApp.getUi().alert('Setup complete! Sheets "Expenses", "Deposits", and "Settings" are ready (v3 with Deposits & Photo Upload).');
}
