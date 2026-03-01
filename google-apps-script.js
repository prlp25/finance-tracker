// ============================================
// POLEDGER - Google Apps Script (v4)
// Paste this entire code into your Google Sheet's Apps Script editor
// Supports: Expenses, Deposits, Settings, Photo Uploads, Receipt OCR
// ============================================

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'getAll';
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === 'getAll') {
    var expSheet = ss.getSheetByName('Expenses');
    var depSheet = ss.getSheetByName('Deposits');
    var setSheet = ss.getSheetByName('Settings');

    // Get expenses (13 columns — added ReceiptURL in v4)
    var expenses = [];
    if (expSheet && expSheet.getLastRow() > 1) {
      var lastCol = Math.min(expSheet.getLastColumn(), 13);
      var data = expSheet.getRange(2, 1, expSheet.getLastRow() - 1, lastCol).getValues();
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
          notes: data[i][11] || '',
          receiptUrl: (data[i].length > 12 ? data[i][12] : '') || ''
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
      var setData = setSheet.getRange(2, 1, 1, Math.min(setSheet.getLastColumn(), 3)).getValues();
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
      var receiptUrl = '';
      if (exp.receiptBase64 && exp.receiptBase64.length > 0) {
        receiptUrl = uploadPhotoToDrive(exp.receiptBase64, exp.receiptName || 'receipt.jpg');
      }
      expSheet.appendRow([
        exp.id, exp.date, exp.description, exp.amount, exp.category,
        exp.fundSource, exp.user, exp.reimbursable, exp.reimbTo || '',
        exp.reimbPct, exp.reimbursed, exp.notes || '', receiptUrl
      ]);
      return ContentService.createTextOutput(JSON.stringify({ success: true, receiptUrl: receiptUrl }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'editExpense') {
      var upd = payload.expense;
      if (expSheet.getLastRow() > 1) {
        var ids = expSheet.getRange(2, 1, expSheet.getLastRow() - 1, 1).getValues();
        for (var k = 0; k < ids.length; k++) {
          if (ids[k][0] == upd.id) {
            var row = k + 2;
            var existingReceipt = '';
            try { existingReceipt = expSheet.getRange(row, 13).getValue() || ''; } catch(ex) {}
            expSheet.getRange(row, 1, 1, 13).setValues([[
              upd.id, upd.date, upd.description, upd.amount, upd.category,
              upd.fundSource, upd.user, upd.reimbursable, upd.reimbTo || '',
              upd.reimbPct, upd.reimbursed, upd.notes || '', upd.receiptUrl || existingReceipt
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

    // ---- OCR RECEIPT ----

    if (action === 'ocrReceipt') {
      var apiKey = getVisionApiKey();
      if (!apiKey) {
        return ContentService.createTextOutput(JSON.stringify({
          error: 'Vision API key not set. Add it to Settings sheet column C (row 2).'
        })).setMimeType(ContentService.MimeType.JSON);
      }

      var photoB64 = payload.photoBase64;
      // Strip data URL prefix
      var cleanB64 = photoB64;
      if (cleanB64.indexOf(',') !== -1) {
        cleanB64 = cleanB64.split(',')[1];
      }

      // Also upload the receipt to Drive for storage
      var receiptDriveUrl = '';
      try {
        receiptDriveUrl = uploadPhotoToDrive(photoB64, 'receipt-' + new Date().getTime() + '.jpg');
      } catch (uploadErr) {
        // Non-fatal — OCR can still proceed
      }

      // Call Google Cloud Vision API
      var visionUrl = 'https://vision.googleapis.com/v1/images:annotate?key=' + apiKey;
      var visionPayload = {
        requests: [{
          image: { content: cleanB64 },
          features: [{ type: 'TEXT_DETECTION', maxResults: 1 }]
        }]
      };

      var visionRes = UrlFetchApp.fetch(visionUrl, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify(visionPayload),
        muteHttpExceptions: true
      });

      var visionData = JSON.parse(visionRes.getContentText());

      if (visionData.error) {
        return ContentService.createTextOutput(JSON.stringify({
          error: 'Vision API error: ' + (visionData.error.message || JSON.stringify(visionData.error))
        })).setMimeType(ContentService.MimeType.JSON);
      }

      var rawText = '';
      if (visionData.responses && visionData.responses[0] && visionData.responses[0].fullTextAnnotation) {
        rawText = visionData.responses[0].fullTextAnnotation.text;
      } else if (visionData.responses && visionData.responses[0] && visionData.responses[0].textAnnotations) {
        rawText = visionData.responses[0].textAnnotations[0].description;
      }

      var parsed = parseReceiptText(rawText);

      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        storeName: parsed.storeName,
        amount: parsed.amount,
        date: parsed.date,
        category: parsed.category,
        receiptUrl: receiptDriveUrl,
        rawText: rawText
      })).setMimeType(ContentService.MimeType.JSON);
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

  var cleanBase64 = base64Data;
  if (cleanBase64.indexOf(',') !== -1) {
    cleanBase64 = cleanBase64.split(',')[1];
  }

  var mimeType = 'image/jpeg';
  if (base64Data.indexOf('data:') === 0) {
    var mimeMatch = base64Data.match(/data:([^;]+);/);
    if (mimeMatch) mimeType = mimeMatch[1];
  }

  var blob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), mimeType, fileName);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file.getUrl();
}

// ---- VISION API KEY ----

function getVisionApiKey() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var setSheet = ss.getSheetByName('Settings');
  if (!setSheet || setSheet.getLastRow() < 2) return '';
  try {
    var val = setSheet.getRange(2, 3).getValue();
    return val ? String(val).trim() : '';
  } catch (ex) {
    return '';
  }
}

// ---- RECEIPT OCR PARSER ----

function parseReceiptText(text) {
  var result = { storeName: '', amount: 0, date: '', category: '' };
  if (!text) return result;

  var lines = text.split('\n').map(function(l) { return l.trim(); }).filter(function(l) { return l.length > 0; });

  // --- Store name: first meaningful line (skip very short lines or numbers-only) ---
  for (var i = 0; i < Math.min(lines.length, 5); i++) {
    var line = lines[i];
    // Skip lines that are mostly numbers, dates, or very short
    if (line.length < 3) continue;
    if (/^\d[\d\s:.\-\/]+$/.test(line)) continue;
    if (/^(tel|phone|fax|tin|vat|address|branch)/i.test(line)) continue;
    result.storeName = line.replace(/[*#]+/g, '').trim();
    break;
  }

  // --- Amount: look for TOTAL, AMOUNT DUE, GRAND TOTAL patterns ---
  var amountPatterns = [
    /(?:GRAND\s*TOTAL|TOTAL\s*(?:DUE|AMOUNT|AMT)?|AMOUNT\s*DUE|NET\s*AMOUNT|BALANCE\s*DUE|TOTAL\s*SALE)\s*[:\s]*[₱P]?\s*([\d,]+\.?\d*)/i,
    /[₱P]\s*([\d,]+\.\d{2})\s*$/im,
    /(?:TOTAL)\s*[:\s]*\$?\s*([\d,]+\.?\d*)/i
  ];
  var bestAmount = 0;
  for (var p = 0; p < amountPatterns.length; p++) {
    var allMatches = text.match(new RegExp(amountPatterns[p].source, 'gi'));
    if (allMatches) {
      for (var am = 0; am < allMatches.length; am++) {
        var numMatch = allMatches[am].match(/[\d,]+\.?\d*/);
        if (numMatch) {
          var val = parseFloat(numMatch[0].replace(/,/g, ''));
          if (val > bestAmount) bestAmount = val;
        }
      }
    }
    if (bestAmount > 0) break;
  }
  result.amount = bestAmount;

  // --- Date: look for common date formats ---
  var datePatterns = [
    // MM/DD/YYYY or MM-DD-YYYY
    /(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/,
    // YYYY-MM-DD
    /(\d{4})-(\d{1,2})-(\d{1,2})/,
    // Mon DD, YYYY or DD Mon YYYY
    /(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*[\s.,]+(\d{1,2})[\s.,]+(\d{4})/i,
    /(\d{1,2})[\s.,]+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*[\s.,]+(\d{4})/i
  ];
  var monthMap = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };

  for (var dp = 0; dp < datePatterns.length; dp++) {
    var dm = text.match(datePatterns[dp]);
    if (dm) {
      var year, month, day;
      if (dp === 0) { // MM/DD/YYYY
        month = parseInt(dm[1], 10);
        day = parseInt(dm[2], 10);
        year = parseInt(dm[3], 10);
      } else if (dp === 1) { // YYYY-MM-DD
        year = parseInt(dm[1], 10);
        month = parseInt(dm[2], 10);
        day = parseInt(dm[3], 10);
      } else if (dp === 2) { // Mon DD, YYYY
        month = monthMap[dm[1].substring(0, 3).toLowerCase()] || 1;
        day = parseInt(dm[2], 10);
        year = parseInt(dm[3], 10);
      } else if (dp === 3) { // DD Mon YYYY
        day = parseInt(dm[1], 10);
        month = monthMap[dm[2].substring(0, 3).toLowerCase()] || 1;
        year = parseInt(dm[3], 10);
      }
      if (year > 2000 && year < 2100 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        result.date = year + '-' + String(month).padStart(2, '0') + '-' + String(day).padStart(2, '0');
        break;
      }
    }
  }

  // --- Category auto-detect from store name ---
  var upper = (result.storeName + ' ' + text.substring(0, 200)).toUpperCase();
  var groceryKw = ['SM SUPERMARKET', 'SM HYPER', 'ROBINSONS', 'PUREGOLD', 'LANDERS', 'S&R', 'METRO MART', 'WALTERMART', 'EVER GOTESCO', 'SAVEMORE', 'LANDMARK', 'SHOPWISE', 'GROCERY', 'SUPERMARKET', 'MARKET MARKET'];
  var restaurantKw = ['JOLLIBEE', 'MCDONALD', 'STARBUCKS', 'CHOWKING', 'GREENWICH', 'KFC', 'BURGER KING', 'PIZZA HUT', 'SHAKEY', 'YELLOW CAB', 'BONCHON', 'ARMY NAVY', 'MAX\'S', 'RESTAURANT', 'CAFE', 'DINER', 'EATERY', 'FOOD HALL', 'RAMEN', 'GRILL'];
  var gasKw = ['SHELL', 'PETRON', 'CALTEX', 'PHOENIX', 'SEAOIL', 'FLYING V', 'UNIOIL', 'PTT', 'TOTAL GAS', 'GASOLINE', 'FUEL'];
  var homeKw = ['ACE HARDWARE', 'WILCON', 'CW HOME', 'HANDYMAN', 'TRUE VALUE'];

  for (var gi = 0; gi < groceryKw.length; gi++) {
    if (upper.indexOf(groceryKw[gi]) !== -1) { result.category = 'Groceries'; break; }
  }
  if (!result.category) {
    for (var ri = 0; ri < restaurantKw.length; ri++) {
      if (upper.indexOf(restaurantKw[ri]) !== -1) { result.category = 'Restaurant / Eating Out'; break; }
    }
  }
  if (!result.category) {
    for (var fi = 0; fi < gasKw.length; fi++) {
      if (upper.indexOf(gasKw[fi]) !== -1) { result.category = 'Other Expenses'; result.storeName = result.storeName || 'Gas / Fuel'; break; }
    }
  }
  if (!result.category) {
    for (var hi = 0; hi < homeKw.length; hi++) {
      if (upper.indexOf(homeKw[hi]) !== -1) { result.category = 'Home Improvements'; break; }
    }
  }

  return result;
}

// Run this ONCE to set up the sheets with proper headers
// If upgrading from v3, run setupSheets again to add ReceiptURL column and Vision API Key
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Expenses sheet (13 columns — added ReceiptURL in v4)
  var expSheet = ss.getSheetByName('Expenses');
  if (!expSheet) {
    expSheet = ss.insertSheet('Expenses');
  }
  expSheet.getRange(1, 1, 1, 13).setValues([[
    'ID', 'Date', 'Description', 'Amount', 'Category',
    'Fund Source', 'User', 'Reimbursable', 'Reimb To',
    'Reimb %', 'Reimbursed', 'Notes', 'ReceiptURL'
  ]]);
  expSheet.getRange(1, 1, 1, 13).setFontWeight('bold');

  // Deposits sheet (8 columns)
  var depSheet = ss.getSheetByName('Deposits');
  if (!depSheet) {
    depSheet = ss.insertSheet('Deposits');
  }
  depSheet.getRange(1, 1, 1, 8).setValues([[
    'ID', 'Date', 'Amount', 'Account', 'Description', 'User', 'PhotoURL', 'Notes'
  ]]);
  depSheet.getRange(1, 1, 1, 8).setFontWeight('bold');

  // Settings sheet (3 columns — added Vision API Key in v4)
  var setSheet = ss.getSheetByName('Settings');
  if (!setSheet) {
    setSheet = ss.insertSheet('Settings');
  }
  setSheet.getRange(1, 1, 1, 3).setValues([['Savings Start', 'Opex Start', 'Vision API Key']]);
  setSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  if (setSheet.getLastRow() < 2) {
    setSheet.getRange(2, 1, 1, 3).setValues([[50000, 30000, '']]);
  }

  SpreadsheetApp.getUi().alert('Setup complete! Sheets ready (v4 with Receipt OCR). Remember to add your Vision API key to Settings column C.');
}
