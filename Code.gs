function doGet(e) {
  var template = HtmlService.createTemplateFromFile('index');
  var users = getAllUsers();
  template.usersData = JSON.stringify(users); 
  return template.evaluate()
      .setTitle('Moliya-Pro')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getAllUsers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    users.push({ u: data[i][0], p: data[i][1], r: data[i][2] });
  }
  return users;
}

function getFormData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data'); 
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  
  var points = [];
  var cats = [];
  var payments = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1] !== "") {
       points.push(data[i][1]);
    }
    if (data[i][2] && data[i][2] !== "") {
       payments.push(data[i][2]);
    }
    if (data[i][3] && data[i][3] !== "") {
       cats.push(data[i][3]);
    }
  }
  
  return { points: points, cats: cats, payments: payments };
}

function saveTransaction(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kirim Chiqim');
  var batchId = (data.formMode === 'add' && data.rowId) ? data.rowId : new Date().getTime().toString();
  var dateStr = data.date; 
  var firmVal = data.firm || "";

  if (data.formMode === 'add') {
      var rowsToAdd = [];
      
      if (data.type === 'Kirim') {
        for (var payType in data.payments) {
          var amount = parseFloat(data.payments[payType]);
          if (amount > 0) {
            rowsToAdd.push([dateStr, firmVal, data.point, "Savdo tushumi", payType, amount, 0, data.note, batchId]);
          }
        }
      } 
      else if (data.type === 'Chiqim') {
        var chiqim = parseFloat(data.amount);
        rowsToAdd.push([dateStr, firmVal, data.point, data.category, data.payment, 0, chiqim, data.note, batchId]);
      } 
      else if (data.type === 'Otkazma') {
        var mainAmount = parseFloat(data.amount);
        var usdAmount = parseFloat(data.amountIn);      
        var source = data.transferFrom; 
        var dest = data.transferTo;     
        
        var finalChiqim = (source === 'Dollar' && usdAmount > 0) ? usdAmount : mainAmount;
        var finalKirim = (dest === 'Dollar' && usdAmount > 0) ? usdAmount : mainAmount;

        rowsToAdd.push([dateStr, firmVal, "Transfer", "O'tkazma", source + " -> " + dest, finalKirim, finalChiqim, data.note, batchId, source, dest]);
      }
      
      if (rowsToAdd.length > 0) {
        var lastRow = sheet.getLastRow();
        var numCols = rowsToAdd[0].length;
        sheet.getRange(lastRow + 1, 1, rowsToAdd.length, numCols).setValues(rowsToAdd);
      }
  } 
  else {
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return "No Data";
      var idValues = sheet.getRange(2, 9, lastRow - 1, 1).getValues();
      var rowIndex = -1;

      if (data.type === 'Kirim') {
          var targetId = data.rowId;
          var rowsToDelete = [];
          for (var i = 0; i < idValues.length; i++) { if (idValues[i][0] == targetId) rowsToDelete.push(i + 2); }
          for (var i = rowsToDelete.length - 1; i >= 0; i--) { sheet.deleteRow(rowsToDelete[i]); }

          var rowsToAdd = [];
          for (var payType in data.payments) {
              var amount = parseFloat(data.payments[payType]);
              if (amount > 0) rowsToAdd.push([data.date, firmVal, data.point, "Savdo tushumi", payType, amount, 0, data.note, targetId]);
          }
          if (rowsToAdd.length > 0) sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, 9).setValues(rowsToAdd);
      }
      else if (data.type === 'Chiqim') {
          for (var i = 0; i < idValues.length; i++) { if (idValues[i][0] == data.rowId) { rowIndex = i + 2; break; } }
          if (rowIndex > 0) {
             sheet.getRange(rowIndex, 1, 1, 1).setValue(data.date);
             sheet.getRange(rowIndex, 4, 1, 5).setValues([[data.category, data.payment, 0, data.amount, data.note]]);
          }
      }
      else if (data.type === 'Otkazma') {
          for (var i = 0; i < idValues.length; i++) { if (idValues[i][0] == data.rowId) { rowIndex = i + 2; break; } }
          if (rowIndex > 0) {
             var mainAmount = parseFloat(data.amount);
             var usdAmount = parseFloat(data.amountIn);      
             var source = data.transferFrom; var dest = data.transferTo;     
             var finalChiqim = (source === 'Dollar' && usdAmount > 0) ? usdAmount : mainAmount;
             var finalKirim = (dest === 'Dollar' && usdAmount > 0) ? usdAmount : mainAmount;
             sheet.getRange(rowIndex, 1, 1, 1).setValue(data.date);
             sheet.getRange(rowIndex, 5, 1, 4).setValues([[source + " -> " + dest, finalKirim, finalChiqim, data.note]]);
             sheet.getRange(rowIndex, 10, 1, 2).setValues([[source, dest]]);
          }
      }
  }
  
  SpreadsheetApp.flush(); 
  return "Success";
}

function deleteTransaction(rowId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kirim Chiqim');
  var data = sheet.getDataRange().getValues();
  var deletedCount = 0;

  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][8] == rowId) {
      sheet.deleteRow(i + 1);
      deletedCount++;
    }
  }
  
  return deletedCount > 0 ? "Deleted" : "Not Found";
}

function getTransactionsList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kirim Chiqim');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  
  var result = [];
  for (var i = data.length - 1; i >= 0; i--) {
    var row = data[i];
    var dateVal = row[0];
    var dateStr = "";
    
    if (dateVal instanceof Date) {
        var y = dateVal.getFullYear();
        var m = ("0" + (dateVal.getMonth() + 1)).slice(-2);
        var d = ("0" + dateVal.getDate()).slice(-2);
        dateStr = y + "-" + m + "-" + d;
    } else {
        dateStr = String(dateVal);
    }

    result.push({
      date: dateStr,
      firm: row[1],
      point: row[2],
      cat: row[3],
      pay: row[4],
      kirim: row[5],
      chiqim: row[6],
      note: row[7],
      rowId: row[8]
    });
  }
  return result;
}

function getDashboardMetrics() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kirim Chiqim');
  if (!sheet) return null;

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues(); 
  
  var now = new Date();
  var currentMonth = now.getMonth(); 
  var currentYear = now.getFullYear();

  var stats = {
    monthIncome: 0,
    monthExpense: 0,
    monthProfit: 0,
    yearIncome: [0,0,0,0,0,0,0,0,0,0,0,0], 
    yearExpense: [0,0,0,0,0,0,0,0,0,0,0,0],
    pointsData: {} 
  };

  var blacklist = ["DONIYOR AKA", "BOSHQA", "DIREKTOR", "KASSA"];

  data.forEach(function(row) {
    var dateVal = new Date(row[0]);
    var point = row[2] ? row[2].toString().trim() : "";
    var category = row[3];
    var kirim = parseFloat(row[5]) || 0;
    var chiqim = parseFloat(row[6]) || 0;

    if (dateVal.getFullYear() === currentYear) {
       var m = dateVal.getMonth();
       
       var isValidIncome = (kirim > 0 && category === "Savdo tushumi" && blacklist.indexOf(point.toUpperCase()) === -1);
       var isValidExpense = (chiqim > 0 && category !== "O'tkazma" && category !== "Transfer");

       if (isValidIncome) {
           stats.yearIncome[m] += kirim;
           if (m === currentMonth) stats.monthIncome += kirim;
           
           if (!stats.pointsData[point]) stats.pointsData[point] = 0;
           stats.pointsData[point] += kirim;
       }

       if (isValidExpense) {
           stats.yearExpense[m] += chiqim;
           if (m === currentMonth) stats.monthExpense += chiqim;
       }
    }
  });

  stats.monthProfit = stats.monthIncome - stats.monthExpense;
  return stats;
}

function getRealTimeBalance() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kirim Chiqim');
  if (!sheet) return { Naqd: 0, P2P: 0, Dollar: 0 };
  
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { Naqd: 0, P2P: 0, Dollar: 0 };
  
  var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues(); 
  
  var balance = { "Naqd": 0, "P2P": 0, "Dollar": 0, "Bank": 0 };

  data.forEach(function(row) {
    var category = row[3]; 
    var payType = row[4];  
    var kirim = parseFloat(row[5]) || 0; 
    var chiqim = parseFloat(row[6]) || 0;
    
    if (category === "O'tkazma" || category === "Transfer") {
       var source = row[9]; 
       var dest = row[10];
       
       if (!source && payType.includes('->')) {
           var parts = payType.split('->');
           source = parts[0].trim();
           dest = parts[1].trim();
       }

       if (balance.hasOwnProperty(source)) balance[source] -= chiqim; 
       if (balance.hasOwnProperty(dest)) balance[dest] += kirim;
    } 
    else {
       if (balance.hasOwnProperty(payType)) {
          balance[payType] += (kirim - chiqim);
       }
    }
  });

  return balance;
}
