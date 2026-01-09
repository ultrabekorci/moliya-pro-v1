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
    var permString = (data[i].length > 4) ? data[i][4] : "";
    var permissions = {};
    

    try {
      if (permString && permString.toString().trim() !== "") {
        permissions = JSON.parse(permString);
      } else {
        if (data[i][2] === 'Admin') permissions = { admin: true }; 
      }
    } catch (e) {
      permissions = {};
    }

    users.push({ 
      u: data[i][0], 
      p: data[i][1], 
      r: data[i][2], 
      s: data[i][3] || 'active', 
      perms: permissions,
      row: i + 1 
    });
  }
  return users;
}
function getFormData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  if (!sheet) return null;
  
  // Ma'lumotlarni olish
  var data = sheet.getDataRange().getValues();
  var firms = [], points = [], payments = [], cats = [];


  function checkStatus(val) {
    if (!val) return 'active';
    var s = String(val).trim().toLowerCase();
    return (s === 'inactive') ? 'inactive' : 'active';
  }


  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowId = i + 1;


    if (row[0] && String(row[0]).trim() !== '') {
      firms.push({ 
        value: String(row[0]).trim(), 
        status: checkStatus(row.length > 4 ? row[4] : null), 
        row: rowId
      });
    }


    if (row.length > 1 && row[1] && String(row[1]).trim() !== '') {
      points.push({ 
        value: String(row[1]).trim(), 
        status: checkStatus(row.length > 5 ? row[5] : null), 
        row: rowId 
      });
    }


    if (row.length > 2 && row[2] && String(row[2]).trim() !== '') {
      payments.push({ 
        value: String(row[2]).trim(), 
        status: checkStatus(row.length > 6 ? row[6] : null), 
        row: rowId 
      });
    }


    if (row.length > 3 && row[3] && String(row[3]).trim() !== '') {
      cats.push({ 
        value: String(row[3]).trim(), 
        status: checkStatus(row.length > 7 ? row[7] : null), 
        row: rowId 
      });
    }
  }
  
  return { firms: firms, points: points, payments: payments, cats: cats };
}

function saveDataItem(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  var columnMap = { 'firm': {v:1, s:5}, 'point': {v:2, s:6}, 'payment': {v:3, s:7}, 'category': {v:4, s:8} };
  var cols = columnMap[data.type];
  
  var targetRow;

  if (data.row == -1) {
    targetRow = sheet.getLastRow() + 1;
    sheet.getRange(targetRow, cols.v).setValue(data.value);
    sheet.getRange(targetRow, cols.s).setValue('active');
  } else {
    targetRow = parseInt(data.row);
    sheet.getRange(targetRow, cols.v).setValue(data.value);
  }
  
  SpreadsheetApp.flush();
  return targetRow;
}

function toggleDataItemStatus(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  var columnMap = { 'firm': 5, 'point': 6, 'payment': 7, 'category': 8 };
  
  var targetRow = parseInt(data.row);
  var cell = sheet.getRange(targetRow, columnMap[data.type]);
  var newStatus = (cell.getValue() === 'active') ? 'inactive' : 'active';
  
  cell.setValue(newStatus);
  return newStatus;
}

function deleteDataItem(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  var columnMap = { 'firm': {v:1, s:5}, 'point': {v:2, s:6}, 'payment': {v:3, s:7}, 'category': {v:4, s:8} };
  var cols = columnMap[data.type];
  
  var targetRow = parseInt(data.row);
  sheet.getRange(targetRow, cols.v).clearContent();
  sheet.getRange(targetRow, cols.s).clearContent();
  return "Success";
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
        var originalUSD = (data.payment === 'Dollar' && data.originalUSD) ? data.originalUSD : null;
        rowsToAdd.push([dateStr, firmVal, data.point, data.category, data.payment, 0, chiqim, data.note, batchId, originalUSD, null]);
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
          var val = (kirim - chiqim);
          
          // Fix for Dollar expenses converted to UZS
          // If row[9] (Source/OriginalUSD) exists for a Dollar expense, use it as the deduction
          if (payType === 'Dollar' && category !== "O'tkazma" && category !== "Transfer" && row[9]) {
              var originalVal = parseFloat(row[9]);
              if (originalVal > 0) val = -originalVal;
          }
          
          balance[payType] += val;
       }
    }
  });

  return balance;
}



function saveUser(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  var targetRow;
  var permJson = JSON.stringify(data.permissions || {});

  if (data.row == -1) {
    targetRow = sheet.getLastRow() + 1;

    sheet.getRange(targetRow, 1, 1, 5).setValues([[data.login, data.password, data.role, 'active', permJson]]);
  } else {

    targetRow = parseInt(data.row);

    sheet.getRange(targetRow, 1, 1, 3).setValues([[data.login, data.password, data.role]]);
    sheet.getRange(targetRow, 5).setValue(permJson);
  }
  
  SpreadsheetApp.flush();
  return targetRow;
}

function toggleUserStatus(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  var targetRow = parseInt(row);
  

  var cell = sheet.getRange(targetRow, 4);
  var currentStatus = cell.getValue();
  



  var newStatus = (currentStatus === 'inactive') ? 'active' : 'inactive';
  cell.setValue(newStatus);
  
  return newStatus;
}

function deleteUser(row) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');

  sheet.deleteRow(parseInt(row));
  return "Success";
}

function getCurrencyRate(dateStr) {
  try {
    var url = "https://cbu.uz/ru/arkhiv-kursov-valyut/json/all/" + dateStr + "/";
    var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
    
    var rate = 0;
    var rateDate = "";
    
    if (response.getResponseCode() === 200) {
        var json = JSON.parse(response.getContentText());
        if (Array.isArray(json)) {
            for (var i = 0; i < json.length; i++) {
                if (json[i].Ccy === 'USD') {
                    rate = parseFloat(json[i].Rate) || 0;
                    rateDate = json[i].Date;
                    break;
                }
            }
        }
    }

    if (rate === 0) {
        var urlLatest = "https://cbu.uz/ru/arkhiv-kursov-valyut/json/";
        var respLatest = UrlFetchApp.fetch(urlLatest, {muteHttpExceptions: true});
        if (respLatest.getResponseCode() === 200) {
            var jsonLatest = JSON.parse(respLatest.getContentText());
            if (Array.isArray(jsonLatest)) {
              for (var i = 0; i < jsonLatest.length; i++) {
                  if (jsonLatest[i].Ccy === 'USD') {
                      rate = parseFloat(jsonLatest[i].Rate) || 0;
                      rateDate = jsonLatest[i].Date;
                      break;
                  }
              }
            }
        }
    }

    if (rate === 0) return { rate: 0, date: "", error: "USD Not Found (Date & Latest failed)" };
    
    return {
        rate: rate,
        date: rateDate // Return the date of the rate we found (either specific or latest)
    };

  } catch (e) {
    Logger.log("Error fetching rate: " + e.toString());
    return { rate: 0, date: "", error: "Exception: " + e.toString() };
  }
}

// *** MUHIM: RUXSATLARI YANGILASH UCHUN ***
// Ushbu funksiyani yuqoridagi menyudan tanlab "Run" tugmasini bosing
// Bu sizdan 'UrlFetchApp' (internetga ulanish) uchun ruxsat so'raydi.
function authorizeScript() {
  var url = "https://cbu.uz/ru/arkhiv-kursov-valyut/json/";
  console.log("Ruxsat tekshirilmoqda...");
  var response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  console.log("Ruxsat mavjud! Javob kodi: " + response.getResponseCode());
}