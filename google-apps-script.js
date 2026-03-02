// ============================================
// POLEDGER - Google Apps Script (v8)
// Paste this entire code into your Google Sheet's Apps Script editor
// Supports: Expenses, Deposits, Bills, BankAccounts, CreditCards, CCStatements, MerchantMap, Settings, Photo Uploads, Receipt OCR
// Personal-first architecture with household tagging + Credit Card module
// ============================================

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'getAll';
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === 'getAll') {
    var expSheet = ss.getSheetByName('Expenses');
    var depSheet = ss.getSheetByName('Deposits');
    var setSheet = ss.getSheetByName('Settings');
    var billSheet = ss.getSheetByName('Bills');
    var bankSheet = ss.getSheetByName('BankAccounts');
    var ccSheet = ss.getSheetByName('CreditCards');
    var ccStmtSheet = ss.getSheetByName('CCStatements');
    var merchantSheet = ss.getSheetByName('MerchantMap');

    // Get expenses (14 columns)
    var expenses = [];
    if (expSheet && expSheet.getLastRow() > 1) {
      var lastCol = Math.min(expSheet.getLastColumn(), 14);
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
          receiptUrl: (data[i].length > 12 ? data[i][12] : '') || '',
          scope: (data[i].length > 13 ? data[i][13] : '') || 'shared'
        });
      }
    }

    // Get deposits (9 columns)
    var deposits = [];
    if (depSheet && depSheet.getLastRow() > 1) {
      var depLastCol = Math.min(depSheet.getLastColumn(), 9);
      var depData = depSheet.getRange(2, 1, depSheet.getLastRow() - 1, depLastCol).getValues();
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
          notes: depData[j][7] || '',
          scope: (depData[j].length > 8 ? depData[j][8] : '') || 'shared'
        });
      }
    }

    // Get bills (11 columns)
    var bills = [];
    if (billSheet && billSheet.getLastRow() > 1) {
      var billLastCol = Math.min(billSheet.getLastColumn(), 11);
      var billData = billSheet.getRange(2, 1, billSheet.getLastRow() - 1, billLastCol).getValues();
      for (var b = 0; b < billData.length; b++) {
        if (billData[b][0] === '' || billData[b][0] === null) continue;
        bills.push({
          id: billData[b][0],
          name: billData[b][1],
          amount: billData[b][2],
          dueDay: billData[b][3],
          type: billData[b][4] || 'fixed',
          account: billData[b][5] || '',
          assignedTo: billData[b][6] || 'Both',
          active: billData[b][7] !== false && billData[b][7] !== 'FALSE',
          notes: billData[b][8] || '',
          lastPaidDate: billData[b][9] || '',
          category: (billData[b].length > 10 ? billData[b][10] : '') || 'Other Expenses'
        });
      }
    }

    // Get bank accounts (5 columns)
    var bankAccounts = [];
    if (bankSheet && bankSheet.getLastRow() > 1) {
      var bankLastCol = Math.min(bankSheet.getLastColumn(), 5);
      var bankData = bankSheet.getRange(2, 1, bankSheet.getLastRow() - 1, bankLastCol).getValues();
      for (var ba = 0; ba < bankData.length; ba++) {
        if (bankData[ba][0] === '' || bankData[ba][0] === null) continue;
        bankAccounts.push({
          id: bankData[ba][0],
          name: bankData[ba][1],
          bankName: bankData[ba][2] || '',
          startingBalance: bankData[ba][3] || 0,
          user: bankData[ba][4] || ''
        });
      }
    }

    // Get credit cards (7 columns)
    var creditCards = [];
    if (ccSheet && ccSheet.getLastRow() > 1) {
      var ccLastCol = Math.min(ccSheet.getLastColumn(), 7);
      var ccData = ccSheet.getRange(2, 1, ccSheet.getLastRow() - 1, ccLastCol).getValues();
      for (var cc = 0; cc < ccData.length; cc++) {
        if (ccData[cc][0] === '' || ccData[cc][0] === null) continue;
        creditCards.push({
          id: ccData[cc][0],
          name: ccData[cc][1],
          bankName: ccData[cc][2] || '',
          lastFour: ccData[cc][3] || '',
          statementDueDay: ccData[cc][4] || 1,
          creditLimit: ccData[cc][5] || 0,
          user: ccData[cc][6] || ''
        });
      }
    }

    // Get CC statements (8 columns)
    var ccStatements = [];
    if (ccStmtSheet && ccStmtSheet.getLastRow() > 1) {
      var stmtLastCol = Math.min(ccStmtSheet.getLastColumn(), 8);
      var stmtData = ccStmtSheet.getRange(2, 1, ccStmtSheet.getLastRow() - 1, stmtLastCol).getValues();
      for (var st = 0; st < stmtData.length; st++) {
        if (stmtData[st][0] === '' || stmtData[st][0] === null) continue;
        ccStatements.push({
          id: stmtData[st][0],
          creditCardId: stmtData[st][1],
          billingMonth: stmtData[st][2] || '',
          totalAmount: stmtData[st][3] || 0,
          dueDate: stmtData[st][4] || '',
          paid: stmtData[st][5] === true || stmtData[st][5] === 'TRUE',
          paidDate: stmtData[st][6] || '',
          importedCount: stmtData[st][7] || 0
        });
      }
    }

    // Get merchant map
    var merchantMap = [];
    if (merchantSheet && merchantSheet.getLastRow() > 1) {
      var mLastCol = Math.min(merchantSheet.getLastColumn(), 3);
      var mData = merchantSheet.getRange(2, 1, merchantSheet.getLastRow() - 1, mLastCol).getValues();
      for (var mi = 0; mi < mData.length; mi++) {
        if (mData[mi][0] === '' || mData[mi][0] === null) continue;
        merchantMap.push({
          pattern: mData[mi][0],
          category: mData[mi][1] || '',
          user: mData[mi][2] || ''
        });
      }
    }

    // Get settings (5 columns — keep backward compat)
    var settings = { savingsStart: 50000, opexStart: 30000, personalBudgetPatrick: 0, personalBudgetAica: 0 };
    if (setSheet && setSheet.getLastRow() >= 2) {
      var setLastCol = Math.min(setSheet.getLastColumn(), 5);
      var setData = setSheet.getRange(2, 1, 1, setLastCol).getValues();
      settings.savingsStart = setData[0][0] || 50000;
      settings.opexStart = setData[0][1] || 30000;
      if (setData[0].length > 3) settings.personalBudgetPatrick = setData[0][3] || 0;
      if (setData[0].length > 4) settings.personalBudgetAica = setData[0][4] || 0;
    }

    return ContentService.createTextOutput(JSON.stringify({
      expenses: expenses,
      deposits: deposits,
      bills: bills,
      bankAccounts: bankAccounts,
      creditCards: creditCards,
      ccStatements: ccStatements,
      merchantMap: merchantMap,
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
  var billSheet = ss.getSheetByName('Bills');
  var bankSheet = ss.getSheetByName('BankAccounts');
  var ccSheet = ss.getSheetByName('CreditCards');
  var ccStmtSheet = ss.getSheetByName('CCStatements');
  var merchantSheet = ss.getSheetByName('MerchantMap');

  try {
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action;

    // ---- EXPENSE ACTIONS ----

    if (action === 'addExpense') {
      var exp = payload.expense;
      var receiptUrl = exp.receiptUrl || '';
      if (exp.receiptBase64 && exp.receiptBase64.length > 0) {
        receiptUrl = uploadPhotoToDrive(exp.receiptBase64, exp.receiptName || 'receipt.jpg');
      }
      expSheet.appendRow([
        exp.id, exp.date, exp.description, exp.amount, exp.category,
        exp.fundSource, exp.user, exp.reimbursable, exp.reimbTo || '',
        exp.reimbPct, exp.reimbursed, exp.notes || '', receiptUrl, exp.scope || 'personal'
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
            var existingScope = 'personal';
            try { existingScope = expSheet.getRange(row, 14).getValue() || 'personal'; } catch(ex2) {}
            expSheet.getRange(row, 1, 1, 14).setValues([[
              upd.id, upd.date, upd.description, upd.amount, upd.category,
              upd.fundSource, upd.user, upd.reimbursable, upd.reimbTo || '',
              upd.reimbPct, upd.reimbursed, upd.notes || '', upd.receiptUrl || existingReceipt,
              upd.scope || existingScope
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
        dep.description || '', dep.user, photoUrl, dep.notes || '', dep.scope || 'personal'
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
            depSheet.getRange(depRow, 1, 1, 9).setValues([[
              updDep.id, updDep.date, updDep.amount, updDep.account,
              updDep.description || '', updDep.user, newPhotoUrl, updDep.notes || '',
              updDep.scope || 'personal'
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

    // ---- BILL ACTIONS ----

    if (action === 'addBill') {
      var bill = payload.bill;
      billSheet.appendRow([
        bill.id, bill.name, bill.amount || 0, bill.dueDay,
        bill.type || 'fixed', bill.account || '', bill.assignedTo || 'Both',
        true, bill.notes || '', '', bill.category || 'Other Expenses'
      ]);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'editBill') {
      var updBill = payload.bill;
      if (billSheet.getLastRow() > 1) {
        var billIds = billSheet.getRange(2, 1, billSheet.getLastRow() - 1, 1).getValues();
        for (var bi = 0; bi < billIds.length; bi++) {
          if (billIds[bi][0] == updBill.id) {
            var billRow = bi + 2;
            billSheet.getRange(billRow, 1, 1, 11).setValues([[
              updBill.id, updBill.name, updBill.amount || 0, updBill.dueDay,
              updBill.type || 'fixed', updBill.account || '', updBill.assignedTo || 'Both',
              updBill.active !== false, updBill.notes || '', updBill.lastPaidDate || '',
              updBill.category || 'Other Expenses'
            ]]);
            return ContentService.createTextOutput(JSON.stringify({ success: true }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ error: 'Not found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'markBillPaid') {
      var paidId = payload.id;
      var paidDate = payload.date || new Date().toISOString().split('T')[0];
      var paidUser = payload.user || 'Patrick';
      var paidAmount = payload.amount;
      var paidFromAccount = payload.bankAccountId || '';
      if (billSheet.getLastRow() > 1) {
        var billLastColM = Math.min(billSheet.getLastColumn(), 11);
        var billRows = billSheet.getRange(2, 1, billSheet.getLastRow() - 1, billLastColM).getValues();
        for (var pi = 0; pi < billRows.length; pi++) {
          if (billRows[pi][0] == paidId) {
            var bRow = pi + 2;
            var bName = billRows[pi][1];
            var bAmount = paidAmount || billRows[pi][2] || 0;
            var bAccount = paidFromAccount || billRows[pi][5] || '';
            var bCategory = (billRows[pi].length > 10 ? billRows[pi][10] : '') || 'Other Expenses';

            billSheet.getRange(bRow, 10).setValue(paidDate);

            var maxExpId = 0;
            if (expSheet.getLastRow() > 1) {
              var expIds = expSheet.getRange(2, 1, expSheet.getLastRow() - 1, 1).getValues();
              for (var ei = 0; ei < expIds.length; ei++) {
                if (Number(expIds[ei][0]) > maxExpId) maxExpId = Number(expIds[ei][0]);
              }
            }
            var newExpId = maxExpId + 1;
            expSheet.appendRow([
              newExpId, paidDate, bName, bAmount, bCategory,
              bAccount, paidUser, false, '', 50, false,
              'Auto-recorded from bill payment', '', 'personal'
            ]);

            return ContentService.createTextOutput(JSON.stringify({
              success: true,
              expenseId: newExpId,
              expenseAmount: bAmount,
              expenseCategory: bCategory
            })).setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ error: 'Not found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'deleteBill') {
      var delBillId = payload.id;
      if (billSheet.getLastRow() > 1) {
        var allBillIds = billSheet.getRange(2, 1, billSheet.getLastRow() - 1, 1).getValues();
        for (var di = 0; di < allBillIds.length; di++) {
          if (allBillIds[di][0] == delBillId) {
            billSheet.deleteRow(di + 2);
            return ContentService.createTextOutput(JSON.stringify({ success: true }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ error: 'Not found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ---- BANK ACCOUNT ACTIONS ----

    if (action === 'addBankAccount') {
      var ba = payload.bankAccount;
      bankSheet.appendRow([
        ba.id, ba.name, ba.bankName || '', ba.startingBalance || 0, ba.user || ''
      ]);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'editBankAccount') {
      var updBA = payload.bankAccount;
      if (bankSheet.getLastRow() > 1) {
        var baIds = bankSheet.getRange(2, 1, bankSheet.getLastRow() - 1, 1).getValues();
        for (var bai = 0; bai < baIds.length; bai++) {
          if (baIds[bai][0] == updBA.id) {
            var baRow = bai + 2;
            bankSheet.getRange(baRow, 1, 1, 5).setValues([[
              updBA.id, updBA.name, updBA.bankName || '', updBA.startingBalance || 0, updBA.user || ''
            ]]);
            return ContentService.createTextOutput(JSON.stringify({ success: true }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ error: 'Not found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'deleteBankAccount') {
      var delBAId = payload.id;
      if (bankSheet.getLastRow() > 1) {
        var allBAIds = bankSheet.getRange(2, 1, bankSheet.getLastRow() - 1, 1).getValues();
        for (var bdi = 0; bdi < allBAIds.length; bdi++) {
          if (allBAIds[bdi][0] == delBAId) {
            bankSheet.deleteRow(bdi + 2);
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
      var cleanB64 = photoB64;
      if (cleanB64.indexOf(',') !== -1) {
        cleanB64 = cleanB64.split(',')[1];
      }

      var receiptDriveUrl = '';
      try {
        receiptDriveUrl = uploadPhotoToDrive(photoB64, 'receipt-' + new Date().getTime() + '.jpg');
      } catch (uploadErr) {}

      var visionUrl = 'https://vision.googleapis.com/v1/images:annotate?key=' + apiKey;
      var visionPayload = {
        requests: [{
          image: { content: cleanB64 },
          features: [{ type: 'DOCUMENT_TEXT_DETECTION', maxResults: 1 }]
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

    // ---- CREDIT CARD ACTIONS ----

    if (action === 'addCreditCard') {
      var cc = payload.creditCard;
      ccSheet.appendRow([
        cc.id, cc.name, cc.bankName || '', cc.lastFour || '',
        cc.statementDueDay || 1, cc.creditLimit || 0, cc.user || ''
      ]);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'editCreditCard') {
      var updCC = payload.creditCard;
      if (ccSheet.getLastRow() > 1) {
        var ccIds = ccSheet.getRange(2, 1, ccSheet.getLastRow() - 1, 1).getValues();
        for (var cci = 0; cci < ccIds.length; cci++) {
          if (ccIds[cci][0] == updCC.id) {
            var ccRow = cci + 2;
            ccSheet.getRange(ccRow, 1, 1, 7).setValues([[
              updCC.id, updCC.name, updCC.bankName || '', updCC.lastFour || '',
              updCC.statementDueDay || 1, updCC.creditLimit || 0, updCC.user || ''
            ]]);
            return ContentService.createTextOutput(JSON.stringify({ success: true }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ error: 'Not found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'deleteCreditCard') {
      var delCCId = payload.id;
      if (ccSheet.getLastRow() > 1) {
        var allCCIds = ccSheet.getRange(2, 1, ccSheet.getLastRow() - 1, 1).getValues();
        for (var dci = 0; dci < allCCIds.length; dci++) {
          if (allCCIds[dci][0] == delCCId) {
            ccSheet.deleteRow(dci + 2);
            return ContentService.createTextOutput(JSON.stringify({ success: true }))
              .setMimeType(ContentService.MimeType.JSON);
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ error: 'Not found' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ---- CC STATEMENT IMPORT (bulk add expenses from OCR) ----

    if (action === 'importCCStatement') {
      var stmt = payload.statement;
      var items = payload.items || [];
      // Save statement record
      ccStmtSheet.appendRow([
        stmt.id, stmt.creditCardId, stmt.billingMonth, stmt.totalAmount || 0,
        stmt.dueDate || '', false, '', items.length
      ]);
      // Bulk-add approved expenses
      var maxExpId = 0;
      if (expSheet.getLastRow() > 1) {
        var expIds = expSheet.getRange(2, 1, expSheet.getLastRow() - 1, 1).getValues();
        for (var ei = 0; ei < expIds.length; ei++) {
          if (Number(expIds[ei][0]) > maxExpId) maxExpId = Number(expIds[ei][0]);
        }
      }
      var createdIds = [];
      for (var it = 0; it < items.length; it++) {
        var item = items[it];
        if (item.skip) continue; // skipped duplicates
        maxExpId++;
        var reimb = item.reimbursable ? true : false;
        var reimbTo = item.reimbTo || '';
        var reimbPct = item.reimbPct || 50;
        expSheet.appendRow([
          maxExpId, item.date, item.description, item.amount, item.category,
          item.fundSource || '', item.user || '', reimb, reimbTo, reimbPct, false,
          'CC Import: ' + (stmt.billingMonth || ''), '', item.scope || 'personal'
        ]);
        createdIds.push(maxExpId);
        // Save merchant mapping if category was set
        if (item.merchantPattern && item.category) {
          saveMerchantMapping(merchantSheet, item.merchantPattern, item.category, item.user || '');
        }
      }
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        importedCount: createdIds.length,
        expenseIds: createdIds,
        statementId: stmt.id
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ---- MERCHANT MAP ----

    if (action === 'saveMerchantMap') {
      var mp = payload.mapping;
      saveMerchantMapping(merchantSheet, mp.pattern, mp.category, mp.user || '');
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ---- OCR STATEMENT (parse CC statement image) ----

    if (action === 'ocrStatement') {
      var apiKey = getVisionApiKey();
      if (!apiKey) {
        return ContentService.createTextOutput(JSON.stringify({
          error: 'Vision API key not set. Add it to Settings sheet column C (row 2).'
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var stmtB64 = payload.photoBase64;
      var cleanStmtB64 = stmtB64;
      if (cleanStmtB64.indexOf(',') !== -1) {
        cleanStmtB64 = cleanStmtB64.split(',')[1];
      }
      var visionUrl2 = 'https://vision.googleapis.com/v1/images:annotate?key=' + apiKey;
      var visionPayload2 = {
        requests: [{
          image: { content: cleanStmtB64 },
          features: [{ type: 'DOCUMENT_TEXT_DETECTION', maxResults: 1 }]
        }]
      };
      var visionRes2 = UrlFetchApp.fetch(visionUrl2, {
        method: 'POST',
        contentType: 'application/json',
        payload: JSON.stringify(visionPayload2),
        muteHttpExceptions: true
      });
      var visionData2 = JSON.parse(visionRes2.getContentText());
      if (visionData2.error) {
        return ContentService.createTextOutput(JSON.stringify({
          error: 'Vision API error: ' + (visionData2.error.message || JSON.stringify(visionData2.error))
        })).setMimeType(ContentService.MimeType.JSON);
      }
      var rawText2 = '';
      if (visionData2.responses && visionData2.responses[0] && visionData2.responses[0].fullTextAnnotation) {
        rawText2 = visionData2.responses[0].fullTextAnnotation.text;
      } else if (visionData2.responses && visionData2.responses[0] && visionData2.responses[0].textAnnotations) {
        rawText2 = visionData2.responses[0].textAnnotations[0].description;
      }
      var parsedItems = parseStatementText(rawText2);
      // Auto-categorize using merchant map
      var mMap = [];
      if (merchantSheet && merchantSheet.getLastRow() > 1) {
        var mData = merchantSheet.getRange(2, 1, merchantSheet.getLastRow() - 1, 3).getValues();
        for (var mm = 0; mm < mData.length; mm++) {
          if (mData[mm][0]) mMap.push({ pattern: String(mData[mm][0]).toUpperCase(), category: mData[mm][1], user: mData[mm][2] });
        }
      }
      for (var pi = 0; pi < parsedItems.length; pi++) {
        if (!parsedItems[pi].category) {
          var descUpper = String(parsedItems[pi].description).toUpperCase();
          for (var mi2 = 0; mi2 < mMap.length; mi2++) {
            if (descUpper.indexOf(mMap[mi2].pattern) !== -1) {
              parsedItems[pi].category = mMap[mi2].category;
              parsedItems[pi].autoCategory = true;
              break;
            }
          }
        }
      }
      return ContentService.createTextOutput(JSON.stringify({
        success: true,
        items: parsedItems,
        rawText: rawText2
      })).setMimeType(ContentService.MimeType.JSON);
    }

    // ---- SETTINGS ----

    if (action === 'updateSettings') {
      var s = payload.settings;
      setSheet.getRange(2, 1, 1, 2).setValues([[s.savingsStart || 0, s.opexStart || 0]]);
      return ContentService.createTextOutput(JSON.stringify({ success: true }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'updatePersonalBudget') {
      var pb = payload;
      setSheet.getRange(2, 4, 1, 2).setValues([[pb.patrick || 0, pb.aica || 0]]);
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
  var fullText = text.toUpperCase();

  // --- Grab screenshot detection ---
  var isGrab = /GRAB|GRABFOOD|GRABMART|GRABPAY|GRABEXPRESS/i.test(text);
  if (isGrab) {
    result.category = 'Grab Delivery';

    for (var gi = 0; gi < Math.min(lines.length, 15); gi++) {
      var gl = lines[gi];
      if (/^(grab|your order|order|receipt|completed|delivered|paid|subtotal|total|delivery|discount|promo|voucher|rating|rate|help|cancel|reorder|view|share|items?|x\d)/i.test(gl)) continue;
      if (gl.length < 3 || /^\d[\d\s:.\-\/₱]+$/.test(gl)) continue;
      if (/^[₱P]\s*\d/.test(gl)) continue;
      result.storeName = gl.replace(/[*#=]+/g, '').trim();
      break;
    }

    var grabTotal = 0;
    var grabTotalMatch = text.match(/(?:Total|You paid|Amount paid|Order total)\s*[:\s]*[₱P]\s*([\d,]+(?:\.\d{1,2})?)/i);
    if (grabTotalMatch) {
      grabTotal = parseFloat(grabTotalMatch[1].replace(/,/g, ''));
    }
    if (!grabTotal) {
      var grabAmounts = [];
      var grabPesoRe = /[₱P]\s*([\d,]+(?:\.\d{1,2})?)/g;
      var gm;
      while ((gm = grabPesoRe.exec(text)) !== null) {
        var gv = parseFloat(gm[1].replace(/,/g, ''));
        if (gv > 0) grabAmounts.push(gv);
      }
      if (grabAmounts.length > 0) {
        grabAmounts.sort(function(a, b) { return b - a; });
        grabTotal = grabAmounts[0];
      }
    }
    if (grabTotal > 0) result.amount = grabTotal;

    var grabDateMatch = text.match(/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*[\s.,]+(\d{1,2})[\s.,]+(\d{4})/i);
    if (grabDateMatch) {
      var gMonthMap = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
      var gMon = gMonthMap[grabDateMatch[1].substring(0, 3).toLowerCase()] || 1;
      result.date = parseInt(grabDateMatch[3], 10) + '-' + String(gMon).padStart(2, '0') + '-' + String(parseInt(grabDateMatch[2], 10)).padStart(2, '0');
    }
    if (!result.date) {
      var grabDate2 = text.match(/(\d{1,2})[\s.\-]+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*[\s.,]+(\d{4})/i);
      if (grabDate2) {
        var gMonthMap2 = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
        var gMon2 = gMonthMap2[grabDate2[2].substring(0, 3).toLowerCase()] || 1;
        result.date = parseInt(grabDate2[3], 10) + '-' + String(gMon2).padStart(2, '0') + '-' + String(parseInt(grabDate2[1], 10)).padStart(2, '0');
      }
    }

    return result;
  }

  // --- BDO card slip detection ---
  var isBDO = /BDO|BANCO DE ORO|CONTACTLESS/i.test(text);

  // --- Store name: first meaningful line ---
  for (var i = 0; i < Math.min(lines.length, 8); i++) {
    var line = lines[i];
    if (line.length < 3) continue;
    if (/^\d[\d\s:.\-\/]+$/.test(line)) continue;
    if (/^(tel|phone|fax|tin|vat|address|branch|station|receipt|invoice|or\s*#|si\s*#|date|time|cashier|bdo|banco|card|contactless|terminal|merchant|acq|aid|tc|tvr|appr|ref|trace|batch|stan)/i.test(line)) continue;
    if (/^\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}/.test(line)) continue;
    result.storeName = line.replace(/[*#=\-]+/g, '').trim();
    break;
  }

  // --- Amount ---
  var amounts = [];

  function cleanAmount(s) {
    return s.replace(/,/g, '').replace(/\s*\.\s*/g, '.').trim();
  }

  var totalDuePatterns = [
    /TOTAL\s*AMOUNT\s*DUE\s*[:\s=]*[₱PHP\s]*([\d,]+\s*\.\s*\d{1,2})/gi,
    /AMOUNT\s*DUE\s*[:\s=]*[₱PHP\s]*([\d,]+\s*\.\s*\d{1,2})/gi,
    /TOTAL\s*DUE\s*[:\s=]*[₱PHP\s]*([\d,]+\s*\.\s*\d{1,2})/gi
  ];
  for (var td = 0; td < totalDuePatterns.length; td++) {
    var tdMatch;
    var tdRe = totalDuePatterns[td];
    while ((tdMatch = tdRe.exec(text)) !== null) {
      var tdVal = parseFloat(cleanAmount(tdMatch[1]));
      if (tdVal > 0 && tdVal < 10000000) amounts.push({ val: tdVal, priority: 0 });
    }
  }

  var totalPatterns = [
    /(?:GRAND\s*TOTAL|TOTAL\s*(?:AMOUNT|AMT|SALE|SALES)?|NET\s*(?:AMOUNT|TOTAL)|BALANCE\s*DUE|AMOUNT\s*PAID)\s*[:\s=]*[₱PHP\s]*([\d,]+\s*\.\s*\d{1,2})/gi,
    /(?:TOTAL|AMOUNT|SALE|PESO)\s*[:\s=]*\s*([\d,]+\s*\.\s*\d{2})/gi
  ];
  for (var tp = 0; tp < totalPatterns.length; tp++) {
    var tMatch;
    var re = totalPatterns[tp];
    while ((tMatch = re.exec(text)) !== null) {
      var tVal = parseFloat(cleanAmount(tMatch[1]));
      if (tVal > 0 && tVal < 10000000) amounts.push({ val: tVal, priority: 1 });
    }
  }

  if (isBDO) {
    var bdoAmtMatch = text.match(/(?:AMOUNT|TOTAL|SALE)\s*[:\s]*[₱PHP\s]*([\d,]+\s*\.\s*\d{2})/i);
    if (bdoAmtMatch) {
      var bdoVal = parseFloat(cleanAmount(bdoAmtMatch[1]));
      if (bdoVal > 0 && bdoVal < 10000000) amounts.push({ val: bdoVal, priority: 0 });
    }
    for (var bii = 0; bii < lines.length; bii++) {
      var bdoLine = lines[bii].match(/^[₱PHP\s]*([\d,]+\.\d{2})\s*$/);
      if (bdoLine) {
        var bVal = parseFloat(cleanAmount(bdoLine[1]));
        if (bVal > 100 && bVal < 10000000) amounts.push({ val: bVal, priority: 2 });
      }
    }
  }

  var pesoPattern = /[₱P]\s*([\d,]+\s*\.\s*\d{1,2})/g;
  var pMatch;
  while ((pMatch = pesoPattern.exec(text)) !== null) {
    var pVal = parseFloat(cleanAmount(pMatch[1]));
    if (pVal > 0 && pVal < 10000000) amounts.push({ val: pVal, priority: 2 });
  }

  for (var li = 0; li < lines.length; li++) {
    var lineMatch = lines[li].match(/^\s*[₱PHP\s]*([\d,]+\s*\.\s*\d{2})\s*$/);
    if (lineMatch) {
      var lVal = parseFloat(cleanAmount(lineMatch[1]));
      if (lVal > 0 && lVal < 10000000) amounts.push({ val: lVal, priority: 3 });
    }
  }

  if (amounts.length > 0) {
    amounts.sort(function(a, b) { return a.priority - b.priority || b.val - a.val; });
    result.amount = amounts[0].val;
  }

  // --- Date ---
  var monthMap = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
  var foundDate = false;
  for (var dli = 0; dli < lines.length && !foundDate; dli++) {
    var dline = lines[dli];

    var bdoDate = dline.match(/(\d{1,2})\s*(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s*(\d{4})/i);
    if (bdoDate) {
      var bdoMon = monthMap[bdoDate[2].substring(0, 3).toLowerCase()] || 1;
      result.date = parseInt(bdoDate[3], 10) + '-' + String(bdoMon).padStart(2, '0') + '-' + String(parseInt(bdoDate[1], 10)).padStart(2, '0');
      foundDate = true; continue;
    }

    var ymd = dline.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,3})/);
    if (ymd) {
      var y1 = parseInt(ymd[1], 10), m1 = parseInt(ymd[2], 10), d1 = parseInt(ymd[3], 10);
      if (y1 > 2000 && y1 < 2100 && m1 >= 1 && m1 <= 12 && d1 >= 1 && d1 <= 31) {
        result.date = y1 + '-' + String(m1).padStart(2, '0') + '-' + String(d1).padStart(2, '0');
        foundDate = true; continue;
      }
    }

    var mdy = dline.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (mdy) {
      var a = parseInt(mdy[1], 10), b2 = parseInt(mdy[2], 10), c = parseInt(mdy[3], 10);
      if (c > 2000 && c < 2100) {
        if (a >= 1 && a <= 12 && b2 >= 1 && b2 <= 31) {
          result.date = c + '-' + String(a).padStart(2, '0') + '-' + String(b2).padStart(2, '0');
          foundDate = true; continue;
        }
        if (a > 12 && a <= 31 && b2 >= 1 && b2 <= 12) {
          result.date = c + '-' + String(b2).padStart(2, '0') + '-' + String(a).padStart(2, '0');
          foundDate = true; continue;
        }
      }
    }

    var mdyShort = dline.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2})(?!\d)/);
    if (mdyShort) {
      var p1 = parseInt(mdyShort[1], 10), p2 = parseInt(mdyShort[2], 10), ys = parseInt(mdyShort[3], 10) + 2000;
      if (p1 >= 1 && p1 <= 12 && p2 >= 1 && p2 <= 31) {
        result.date = ys + '-' + String(p1).padStart(2, '0') + '-' + String(p2).padStart(2, '0');
        foundDate = true; continue;
      }
      if (p1 > 12 && p1 <= 31 && p2 >= 1 && p2 <= 12) {
        result.date = ys + '-' + String(p2).padStart(2, '0') + '-' + String(p1).padStart(2, '0');
        foundDate = true; continue;
      }
    }

    var mdn = dline.match(/(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*[\s.,\-]+(\d{1,2})[\s.,\-]+(\d{4})/i);
    if (mdn) {
      var mn = monthMap[mdn[1].substring(0, 3).toLowerCase()] || 1;
      result.date = parseInt(mdn[3], 10) + '-' + String(mn).padStart(2, '0') + '-' + String(parseInt(mdn[2], 10)).padStart(2, '0');
      foundDate = true; continue;
    }
    var dmn = dline.match(/(\d{1,2})[\s.,\-]+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*[\s.,\-]+(\d{4})/i);
    if (dmn) {
      var mn2 = monthMap[dmn[2].substring(0, 3).toLowerCase()] || 1;
      result.date = parseInt(dmn[3], 10) + '-' + String(mn2).padStart(2, '0') + '-' + String(parseInt(dmn[1], 10)).padStart(2, '0');
      foundDate = true; continue;
    }
  }

  // --- Category auto-detect ---
  var fullUpper = (result.storeName + ' ' + text).toUpperCase();

  var grabKw = ['GRAB', 'GRABFOOD', 'GRABMART', 'GRABPAY', 'GRABEXPRESS', 'FOODPANDA', 'PANDA'];
  for (var gki = 0; gki < grabKw.length; gki++) {
    if (fullUpper.indexOf(grabKw[gki]) !== -1) { result.category = 'Grab Delivery'; break; }
  }

  if (!result.category) {
    var gasKw = ['SHELL', 'PETRON', 'CALTEX', 'CHEVRON', 'PHOENIX', 'SEAOIL', 'FLYING V', 'UNIOIL', 'PTT', 'TOTAL GAS', 'GASOLINE', 'FUEL', 'DIESEL', 'UNLEADED', 'GAS STATION', 'PETROLEUM', 'CLEANFUEL'];
    var groceryKw = ['SM SUPERMARKET', 'SM HYPER', 'ROBINSONS', 'PUREGOLD', 'LANDERS', 'S&R', 'METRO MART', 'WALTERMART', 'EVER GOTESCO', 'SAVEMORE', 'LANDMARK', 'SHOPWISE', 'GROCERY', 'SUPERMARKET', 'MARKET MARKET', 'RUSTANS', 'UNIMART'];
    var restaurantKw = ['JOLLIBEE', 'MCDONALD', 'STARBUCKS', 'CHOWKING', 'GREENWICH', 'KFC', 'BURGER KING', 'PIZZA HUT', 'SHAKEY', 'YELLOW CAB', 'BONCHON', 'ARMY NAVY', 'MAX\'S', 'RESTAURANT', 'CAFE', 'DINER', 'EATERY', 'FOOD HALL', 'RAMEN', 'GRILL', 'SAMGYUP', 'INASAL', 'KENNY ROGERS', 'YOSHINOYA', 'SUBWAY', 'TURKS', 'MINISTOP', 'BALESIN', 'RESORT', 'HOTEL', 'BAR & GRILL', 'BISTRO', 'TEPPANYAKI', 'SUSHI', 'STEAKHOUSE', 'BUFFET', 'WINE', 'COCKTAIL'];
    var homeKw = ['ACE HARDWARE', 'WILCON', 'CW HOME', 'HANDYMAN', 'TRUE VALUE', 'DATABLITZ'];
    var healthKw = ['MERCURY DRUG', 'WATSONS', 'SOUTHSTAR', 'GENERIKA', 'TGP', 'ROSE PHARMACY', 'HOSPITAL', 'CLINIC', 'DENTAL', 'MEDICAL'];

    var kwSets = [
      { kw: gasKw, cat: 'Gas / Fuel' },
      { kw: groceryKw, cat: 'Groceries' },
      { kw: restaurantKw, cat: 'Restaurant / Eating Out' },
      { kw: homeKw, cat: 'Home Improvements' },
      { kw: healthKw, cat: 'Health/Medical' }
    ];
    for (var ci = 0; ci < kwSets.length && !result.category; ci++) {
      for (var ki = 0; ki < kwSets[ci].kw.length; ki++) {
        if (fullUpper.indexOf(kwSets[ci].kw[ki]) !== -1) { result.category = kwSets[ci].cat; break; }
      }
    }
  }

  if (!result.category && (/SERVICE\s*CHARGE|12%\s*VAT|VATABLE\s*SALE/i.test(text))) {
    result.category = 'Restaurant / Eating Out';
  }

  return result;
}

// ---- MERCHANT MAP HELPER ----
function saveMerchantMapping(merchantSheet, pattern, category, user) {
  if (!merchantSheet || !pattern || !category) return;
  var upperPattern = String(pattern).toUpperCase().trim();
  // Check if mapping already exists — update it
  if (merchantSheet.getLastRow() > 1) {
    var existing = merchantSheet.getRange(2, 1, merchantSheet.getLastRow() - 1, 3).getValues();
    for (var i = 0; i < existing.length; i++) {
      if (String(existing[i][0]).toUpperCase().trim() === upperPattern) {
        merchantSheet.getRange(i + 2, 2).setValue(category);
        return;
      }
    }
  }
  merchantSheet.appendRow([upperPattern, category, user]);
}

// ---- CC STATEMENT OCR PARSER ----
function parseStatementText(text) {
  var items = [];
  if (!text) return items;

  // Normalize OCR quirks: collapse multiple spaces, fix common OCR artifacts
  text = text.replace(/\r\n/g, '\n').replace(/\r/g, '\n');
  // Normalize Unicode dashes, en-dash, em-dash to regular hyphen
  text = text.replace(/[\u2013\u2014\u2212]/g, '-');
  // Fix OCR sometimes putting comma as period or vice versa in amounts
  // (handled by being flexible in regex)

  var lines = text.split('\n').map(function(l) { return l.replace(/\s+/g, ' ').trim(); }).filter(function(l) { return l.length > 0; });

  var monthMap = { jan:1, feb:2, mar:3, apr:4, may:5, jun:6, jul:7, aug:8, sep:9, oct:10, nov:11, dec:12 };
  var currentYear = new Date().getFullYear();

  // Skip words — lines starting with these are headers/footers, not transactions
  var skipPattern = /^(statement|credit card|card no|card number|account|page|date|description|amount|total|minimum|payment due|balance|previous|new charges|transaction|posting|reference|instalment|credit limit|available|billing|closing|opening|finance charge|annual fee|late charge|over limit|interest|thank you|due date|your |please |effective|as of)/i;

  // Build a list of potential transaction lines using flexible matching
  // The key insight: look for lines that have an amount (X,XXX.XX or XXX.XX) at the end

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];

    // Skip obvious non-transaction lines
    if (skipPattern.test(line)) continue;
    if (/^[\s\-=_*]+$/.test(line)) continue;
    if (/^(INSTALMENT|Reference:|Ref\s)/i.test(line)) continue;
    if (/^PAYMENT RECEIVED/i.test(line)) continue;
    // Skip lines that are just numbers (page numbers, etc.)
    if (/^\d{1,4}$/.test(line)) continue;

    // Master approach: find an amount at the end of the line, then parse what's before it
    // Amount pattern: optional minus, digits with optional commas, dot, 2 decimal digits
    var amtMatch = line.match(/^(.+?)\s+(-?[\d,]+\.\d{2})\s*$/);
    if (!amtMatch) continue;

    var beforeAmt = amtMatch[1].trim();
    var amtStr = amtMatch[2];
    var amt = parseFloat(amtStr.replace(/,/g, ''));

    // Skip tiny amounts (likely page numbers or noise) and huge amounts
    if (Math.abs(amt) < 1 || Math.abs(amt) >= 10000000) continue;

    var date = '';
    var description = '';

    // Try to extract date from the beginning of beforeAmt

    // BDO format: MM/DD/YY MM/DD/YY DESCRIPTION
    var bdoMatch = beforeAmt.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+\d{1,2}\/\d{1,2}\/\d{2,4}\s+(.+)$/);
    if (bdoMatch) {
      var mon = parseInt(bdoMatch[1], 10);
      var day = parseInt(bdoMatch[2], 10);
      var yr = parseInt(bdoMatch[3], 10);
      if (yr < 100) yr += 2000;
      if (mon > 12) { var t = mon; mon = day; day = t; }
      date = yr + '-' + String(mon).padStart(2, '0') + '-' + String(day).padStart(2, '0');
      description = bdoMatch[4].trim();
    }

    // Single date with year: MM/DD/YY DESCRIPTION
    if (!description) {
      var singleYrMatch = beforeAmt.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})\s+(.+)$/);
      if (singleYrMatch) {
        var mon1 = parseInt(singleYrMatch[1], 10);
        var day1 = parseInt(singleYrMatch[2], 10);
        var yr1 = parseInt(singleYrMatch[3], 10);
        if (yr1 < 100) yr1 += 2000;
        if (mon1 > 12) { var t1 = mon1; mon1 = day1; day1 = t1; }
        date = yr1 + '-' + String(mon1).padStart(2, '0') + '-' + String(day1).padStart(2, '0');
        description = singleYrMatch[4].trim();
      }
    }

    // Two dates without year: MM/DD MM/DD DESCRIPTION
    if (!description) {
      var twoDatesMatch = beforeAmt.match(/^(\d{1,2})[\/.](\d{1,2})\s+\d{1,2}[\/.]?\d{1,2}\s+(.+)$/);
      if (twoDatesMatch) {
        var mon2 = parseInt(twoDatesMatch[1], 10);
        var day2 = parseInt(twoDatesMatch[2], 10);
        if (mon2 > 12) { var t2 = mon2; mon2 = day2; day2 = t2; }
        date = currentYear + '-' + String(mon2).padStart(2, '0') + '-' + String(day2).padStart(2, '0');
        description = twoDatesMatch[3].trim();
      }
    }

    // Single date without year: MM/DD DESCRIPTION
    if (!description) {
      var singleMatch = beforeAmt.match(/^(\d{1,2})[\/.](\d{1,2})\s+(.+)$/);
      if (singleMatch) {
        var mon3 = parseInt(singleMatch[1], 10);
        var day3 = parseInt(singleMatch[2], 10);
        if (mon3 > 12) { var t3 = mon3; mon3 = day3; day3 = t3; }
        date = currentYear + '-' + String(mon3).padStart(2, '0') + '-' + String(day3).padStart(2, '0');
        description = singleMatch[3].trim();
      }
    }

    // DD Mon or Mon DD format
    if (!description) {
      var dmMatch = beforeAmt.match(/^(\d{1,2})\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+(.+)$/i);
      if (dmMatch) {
        var mm = monthMap[dmMatch[2].substring(0, 3).toLowerCase()] || 1;
        date = currentYear + '-' + String(mm).padStart(2, '0') + '-' + String(parseInt(dmMatch[1], 10)).padStart(2, '0');
        description = dmMatch[3].trim();
      }
    }
    if (!description) {
      var mdMatch = beforeAmt.match(/^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+(\d{1,2})\s+(.+)$/i);
      if (mdMatch) {
        var mm2 = monthMap[mdMatch[1].substring(0, 3).toLowerCase()] || 1;
        date = currentYear + '-' + String(mm2).padStart(2, '0') + '-' + String(parseInt(mdMatch[2], 10)).padStart(2, '0');
        description = mdMatch[3].trim();
      }
    }

    // No date found — just use the whole text before amount as description
    if (!description) {
      // Only accept if it starts with a letter (not random numbers)
      if (/^[A-Za-z]/.test(beforeAmt)) {
        description = beforeAmt;
      } else {
        continue; // skip lines like "123 456 789.00" that are likely not transactions
      }
    }

    // Final skip check on the extracted description
    if (skipPattern.test(description)) continue;
    if (/^PAYMENT RECEIVED/i.test(description)) continue;

    items.push({
      date: date,
      description: description,
      amount: Math.abs(amt),
      category: '',
      merchantPattern: extractMerchantKey(description),
      isRefund: amt < 0
    });
  }

  // Second pass: try to reassemble split lines (OCR broke rows across 2-3 lines)
  // Handles: date-line → desc+amt-line  OR  date-line → desc-line → amt-line
  if (items.length === 0 && lines.length > 1) {
    var pendingDate = '';
    var pendingDesc = '';
    for (var j = 0; j < lines.length; j++) {
      var ln = lines[j];
      if (skipPattern.test(ln) || /^(INSTALMENT|Reference:|Ref\s|PAYMENT RECEIVED)/i.test(ln)) { pendingDate = ''; pendingDesc = ''; continue; }

      // Is this a date-only line? (one or two dates)
      var dateOnly = ln.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})(\s+\d{1,2}\/\d{1,2}\/\d{2,4})?\s*$/);
      if (dateOnly) {
        var dm = parseInt(dateOnly[1], 10); var dd = parseInt(dateOnly[2], 10); var dy = parseInt(dateOnly[3], 10);
        if (dy < 100) dy += 2000;
        if (dm > 12) { var tt = dm; dm = dd; dd = tt; }
        pendingDate = dy + '-' + String(dm).padStart(2, '0') + '-' + String(dd).padStart(2, '0');
        pendingDesc = '';
        continue;
      }

      // Is this a description+amount line following a date?
      if (pendingDate) {
        var descAmt = ln.match(/^(.+?)\s+(-?[\d,]+\.\d{2})\s*$/);
        if (descAmt) {
          var da = parseFloat(descAmt[2].replace(/,/g, ''));
          var dd2 = descAmt[1].trim();
          if (Math.abs(da) >= 1 && Math.abs(da) < 10000000 && !skipPattern.test(dd2) && !/^PAYMENT RECEIVED/i.test(dd2)) {
            items.push({ date: pendingDate, description: dd2, amount: Math.abs(da), category: '', merchantPattern: extractMerchantKey(dd2), isRefund: da < 0 });
          }
          pendingDate = ''; pendingDesc = '';
          continue;
        }
        // Is this a description-only line? (text, no amount at end)
        if (/^[A-Za-z]/.test(ln) && !ln.match(/-?[\d,]+\.\d{2}\s*$/)) {
          pendingDesc = ln.trim();
          continue;
        }
        // Is this an amount-only line following date+desc?
        if (pendingDesc) {
          var amtOnly = ln.match(/^(-?[\d,]+\.\d{2})\s*$/);
          if (amtOnly) {
            var da2 = parseFloat(amtOnly[1].replace(/,/g, ''));
            if (Math.abs(da2) >= 1 && Math.abs(da2) < 10000000 && !skipPattern.test(pendingDesc)) {
              items.push({ date: pendingDate, description: pendingDesc, amount: Math.abs(da2), category: '', merchantPattern: extractMerchantKey(pendingDesc), isRefund: da2 < 0 });
            }
            pendingDate = ''; pendingDesc = '';
            continue;
          }
        }
      }
      pendingDate = ''; pendingDesc = '';
    }
  }

  // Auto-categorize using built-in keywords (same as receipt parser)
  for (var ac = 0; ac < items.length; ac++) {
    if (items[ac].category) continue;
    items[ac].category = autoCategorize(items[ac].description);
  }

  return items;
}

function extractMerchantKey(desc) {
  // Extract a normalized merchant key from description for merchant mapping
  // Remove common prefixes like GRAB*, POS PURCHASE, etc.
  var clean = String(desc).toUpperCase()
    .replace(/^(GRAB\*|POS\s*(PURCHASE|DEBIT)|CARD\s*PURCHASE|ONLINE\s*PURCHASE|E-?COMMERCE|PAYMENT\s*TO)\s*/i, '')
    .replace(/\s+(BRANCH|STORE|SHOP|BGC|MAKATI|MANILA|TAGUIG|QC|QUEZON|CEBU|DAVAO).*$/i, '')
    .replace(/[*#\-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  // Take first 2-3 meaningful words
  var words = clean.split(' ').filter(function(w) { return w.length > 1; });
  return words.slice(0, 3).join(' ');
}

function autoCategorize(desc) {
  var upper = String(desc).toUpperCase();
  var grabKw = ['GRAB', 'GRABFOOD', 'GRABMART', 'GRABPAY', 'FOODPANDA'];
  var gasKw = ['SHELL', 'PETRON', 'CALTEX', 'PHOENIX', 'SEAOIL', 'UNIOIL', 'CLEANFUEL', 'FUEL', 'GASOLINE'];
  var groceryKw = ['SM SUPERMARKET', 'SM HYPER', 'ROBINSONS', 'PUREGOLD', 'LANDERS', 'S&R', 'METRO MART', 'WALTERMART', 'SAVEMORE', 'SHOPWISE', 'GROCERY', 'SUPERMARKET', 'RUSTANS'];
  var restaurantKw = ['JOLLIBEE', 'MCDONALD', 'STARBUCKS', 'CHOWKING', 'GREENWICH', 'KFC', 'BURGER KING', 'PIZZA HUT', 'SHAKEY', 'YELLOW CAB', 'BONCHON', 'RESTAURANT', 'CAFE', 'DINER', 'RAMEN', 'GRILL', 'SUSHI', 'BUFFET'];
  var homeKw = ['ACE HARDWARE', 'WILCON', 'CW HOME', 'HANDYMAN', 'TRUE VALUE'];
  var healthKw = ['MERCURY DRUG', 'WATSONS', 'SOUTHSTAR', 'GENERIKA', 'HOSPITAL', 'CLINIC', 'DENTAL', 'MEDICAL'];
  var shoppingKw = ['LAZADA', 'SHOPEE', 'ZALORA', 'UNIQLO', 'H&M', 'ZARA', 'SM STORE', 'SM DEPT', 'LANDMARK DEPT'];
  var subsKw = ['NETFLIX', 'SPOTIFY', 'YOUTUBE', 'APPLE', 'GOOGLE PLAY', 'ADOBE', 'MICROSOFT', 'SUBSCRIPTION'];

  var kwSets = [
    { kw: grabKw, cat: 'Grab Delivery' },
    { kw: gasKw, cat: 'Gas / Fuel' },
    { kw: groceryKw, cat: 'Groceries' },
    { kw: restaurantKw, cat: 'Restaurant / Eating Out' },
    { kw: homeKw, cat: 'Home Improvements' },
    { kw: healthKw, cat: 'Medical Expenses' },
    { kw: shoppingKw, cat: 'Shopping' },
    { kw: subsKw, cat: 'Subscriptions' }
  ];
  for (var ci = 0; ci < kwSets.length; ci++) {
    for (var ki = 0; ki < kwSets[ci].kw.length; ki++) {
      if (upper.indexOf(kwSets[ci].kw[ki]) !== -1) return kwSets[ci].cat;
    }
  }
  return '';
}

// Run this ONCE to set up the sheets with proper headers
// v8: Added CreditCards, CCStatements, MerchantMap sheets
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Expenses sheet (14 columns)
  var expSheet = ss.getSheetByName('Expenses');
  if (!expSheet) {
    expSheet = ss.insertSheet('Expenses');
  }
  expSheet.getRange(1, 1, 1, 14).setValues([[
    'ID', 'Date', 'Description', 'Amount', 'Category',
    'Fund Source', 'User', 'Reimbursable', 'Reimb To',
    'Reimb %', 'Reimbursed', 'Notes', 'ReceiptURL', 'Scope'
  ]]);
  expSheet.getRange(1, 1, 1, 14).setFontWeight('bold');

  // Deposits sheet (9 columns)
  var depSheet = ss.getSheetByName('Deposits');
  if (!depSheet) {
    depSheet = ss.insertSheet('Deposits');
  }
  depSheet.getRange(1, 1, 1, 9).setValues([[
    'ID', 'Date', 'Amount', 'Account', 'Description', 'User', 'PhotoURL', 'Notes', 'Scope'
  ]]);
  depSheet.getRange(1, 1, 1, 9).setFontWeight('bold');

  // Bills sheet (11 columns)
  var billSheet = ss.getSheetByName('Bills');
  if (!billSheet) {
    billSheet = ss.insertSheet('Bills');
  }
  billSheet.getRange(1, 1, 1, 11).setValues([[
    'ID', 'Name', 'Amount', 'DueDay', 'Type', 'Account', 'AssignedTo', 'Active', 'Notes', 'LastPaidDate', 'Category'
  ]]);
  billSheet.getRange(1, 1, 1, 11).setFontWeight('bold');

  // BankAccounts sheet (5 columns) — NEW in v7
  var bankSheet = ss.getSheetByName('BankAccounts');
  if (!bankSheet) {
    bankSheet = ss.insertSheet('BankAccounts');
  }
  bankSheet.getRange(1, 1, 1, 5).setValues([[
    'ID', 'Name', 'BankName', 'StartingBalance', 'User'
  ]]);
  bankSheet.getRange(1, 1, 1, 5).setFontWeight('bold');

  // CreditCards sheet (7 columns) — NEW in v8
  var ccSheet = ss.getSheetByName('CreditCards');
  if (!ccSheet) {
    ccSheet = ss.insertSheet('CreditCards');
  }
  ccSheet.getRange(1, 1, 1, 7).setValues([[
    'ID', 'Name', 'BankName', 'LastFour', 'StatementDueDay', 'CreditLimit', 'User'
  ]]);
  ccSheet.getRange(1, 1, 1, 7).setFontWeight('bold');

  // CCStatements sheet (8 columns) — NEW in v8
  var ccStmtSheet = ss.getSheetByName('CCStatements');
  if (!ccStmtSheet) {
    ccStmtSheet = ss.insertSheet('CCStatements');
  }
  ccStmtSheet.getRange(1, 1, 1, 8).setValues([[
    'ID', 'CreditCardID', 'BillingMonth', 'TotalAmount', 'DueDate', 'Paid', 'PaidDate', 'ImportedCount'
  ]]);
  ccStmtSheet.getRange(1, 1, 1, 8).setFontWeight('bold');

  // MerchantMap sheet (3 columns) — NEW in v8
  var merchantSheet = ss.getSheetByName('MerchantMap');
  if (!merchantSheet) {
    merchantSheet = ss.insertSheet('MerchantMap');
  }
  merchantSheet.getRange(1, 1, 1, 3).setValues([[
    'MerchantPattern', 'Category', 'User'
  ]]);
  merchantSheet.getRange(1, 1, 1, 3).setFontWeight('bold');

  // Settings sheet (5 columns)
  var setSheet = ss.getSheetByName('Settings');
  if (!setSheet) {
    setSheet = ss.insertSheet('Settings');
  }
  setSheet.getRange(1, 1, 1, 5).setValues([['Savings Start', 'Opex Start', 'Vision API Key', 'Personal Budget - Patrick', 'Personal Budget - Aica']]);
  setSheet.getRange(1, 1, 1, 5).setFontWeight('bold');
  if (setSheet.getLastRow() < 2) {
    setSheet.getRange(2, 1, 1, 5).setValues([[0, 0, '', 0, 0]]);
  }

  SpreadsheetApp.getUi().alert('Setup complete! Sheets ready (v8 - CC Module with MerchantMap). Remember to add your Vision API key to Settings column C.');
}
