// ==========================================
// 1. SAHIFA YUKLASH (DO GET)
// ==========================================
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Moliya-Pro')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==========================================
// 2. USERLAR VA SOZLAMALARNI OLISH
// ==========================================
function getAllUsers() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Users');
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var users = [];
  // 1-qator sarlavha deb tashlab ketamiz
  for (var i = 1; i < data.length; i++) {
    users.push({ u: data[i][0], p: data[i][1], r: data[i][2] });
  }
  return users;
}

function getFormData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data'); 
  if (!sheet) return null;
  var data = sheet.getDataRange().getValues();
  
  var points = [];   // Savdo nuqtalari (B ustun)
  var cats = [];     // Kategoriyalar (D ustun)
  var payments = []; // To'lov turlari (C ustun)

  // 1-qator sarlavha deb tashlab ketamiz
  for (var i = 1; i < data.length; i++) {
    
    // B ustun (index 1) -> Savdo Nuqtalari
    if (data[i][1] && data[i][1] !== "") {
       points.push(data[i][1]);
    }

    // C ustun (index 2) -> To'lov Turlari
    if (data[i][2] && data[i][2] !== "") {
       payments.push(data[i][2]);
    }

    // D ustun (index 3) -> Kategoriyalar
    if (data[i][3] && data[i][3] !== "") {
       cats.push(data[i][3]);
    }
  }
  
  return { points: points, cats: cats, payments: payments };
}
// ==========================================
// 3. TRANZAKSIYALARNI SAQLASH
// ==========================================
// ==========================================
// 3. TRANZAKSIYALARNI SAQLASH (SUPER OPTIMIZED)
// ==========================================
function saveTransaction(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kirim Chiqim');
  var batchId = new Date().getTime().toString(); 
  var dateStr = data.date; 
  var firmVal = data.firm || ""; 

  // --- A) YANGI QO'SHISH (Batch operatsiya - Barcha turlar) ---
  if (!data.rowId || data.rowId === "") {
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
        
        var finalChiqim = 0; 
        var finalKirim = 0;

        if (source === 'Dollar') { 
          finalChiqim = (usdAmount > 0) ? usdAmount : mainAmount; 
          finalKirim = mainAmount; 
        }
        else if (dest === 'Dollar') { 
          finalChiqim = mainAmount; 
          finalKirim = (usdAmount > 0) ? usdAmount : mainAmount; 
        }
        else { 
          finalChiqim = mainAmount; 
          finalKirim = mainAmount; 
        }

        rowsToAdd.push([dateStr, firmVal, "Transfer", "O'tkazma", source + " -> " + dest, finalKirim, finalChiqim, data.note, batchId, source, dest]);
      }
      
      // BIR MARTA QO'SHISH
      if (rowsToAdd.length > 0) {
        var lastRow = sheet.getLastRow();
        var numCols = rowsToAdd[0].length;
        sheet.getRange(lastRow + 1, 1, rowsToAdd.length, numCols).setValues(rowsToAdd);
      }
  } 
  
  // --- B) TAHRIRLASH (EDIT) - Super Optimized ---
  else {
      var lastRow = sheet.getLastRow();
      if (lastRow < 2) return "No Data";
      
      // Faqat ID ustunini o'qish
      var idRange = sheet.getRange(2, 9, lastRow - 1, 1);
      var idValues = idRange.getValues();
      var rowIndex = -1;

      // 1. KIRIMNI TAHRIRLASH
      if (data.type === 'Kirim') {
          var targetId = data.rowId;
          var rowsToDelete = [];
          
          for (var i = 0; i < idValues.length; i++) {
              if (idValues[i][0] == targetId) {
                  rowsToDelete.push(i + 2);
              }
          }
          
          for (var i = rowsToDelete.length - 1; i >= 0; i--) {
              sheet.deleteRow(rowsToDelete[i]);
          }

          var rowsToAdd = [];
          for (var payType in data.payments) {
              var amount = parseFloat(data.payments[payType]);
              if (amount > 0) {
                  rowsToAdd.push([data.date, firmVal, data.point, "Savdo tushumi", payType, amount, 0, data.note, targetId]);
              }
          }
          
          if (rowsToAdd.length > 0) {
            var lastRow = sheet.getLastRow();
            sheet.getRange(lastRow + 1, 1, rowsToAdd.length, 9).setValues(rowsToAdd);
          }
      }
      
      // 2. CHIQIM TAHRIRLASH (Batch Update)
      else if (data.type === 'Chiqim') {
          for (var i = 0; i < idValues.length; i++) {
              if (idValues[i][0] == data.rowId) {
                  rowIndex = i + 2;
                  break;
              }
          }

          if (rowIndex > 0) {
              // BATCH UPDATE: Bir marta ko'p ustunlarni yangilash
              var updates = [
                [data.date, '', '', data.category, data.payment, 0, data.amount, data.note]
              ];
              // 1, 2, 3-ustunlar (sana, firma, point - point tahrir qilinmaydi)
              sheet.getRange(rowIndex, 1, 1, 1).setValue(data.date);
              sheet.getRange(rowIndex, 4, 1, 5).setValues([[data.category, data.payment, 0, data.amount, data.note]]);
          }
      } 
      
      // 3. O'TKAZMA TAHRIRLASH (Batch Update)
      else if (data.type === 'Otkazma') {
          for (var i = 0; i < idValues.length; i++) {
              if (idValues[i][0] == data.rowId) {
                  rowIndex = i + 2;
                  break;
              }
          }

          if (rowIndex > 0) {
              var mainAmount = parseFloat(data.amount);       
              var usdAmount = parseFloat(data.amountIn);      
              var source = data.transferFrom; 
              var dest = data.transferTo;     
              var finalChiqim = 0; 
              var finalKirim = 0;

              if (source === 'Dollar') { 
                finalChiqim = (usdAmount > 0) ? usdAmount : mainAmount; 
                finalKirim = mainAmount; 
              } 
              else if (dest === 'Dollar') { 
                finalChiqim = mainAmount; 
                finalKirim = (usdAmount > 0) ? usdAmount : mainAmount; 
              } 
              else { 
                finalChiqim = mainAmount; 
                finalKirim = mainAmount; 
              }
              
              // BATCH UPDATE: Ko'p ustunlarni bir marta
              sheet.getRange(rowIndex, 1, 1, 1).setValue(data.date);
              sheet.getRange(rowIndex, 5, 1, 4).setValues([[source + " -> " + dest, finalKirim, finalChiqim, data.note]]);
              sheet.getRange(rowIndex, 10, 1, 2).setValues([[source, dest]]);
          }
      }
  }
  
  SpreadsheetApp.flush(); 
  return "Success";
}

// ==========================================
// 4. O'CHIRISH
// ==========================================
function deleteTransaction(rowId) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kirim Chiqim');
  var data = sheet.getDataRange().getValues();
  for (var i = 0; i < data.length; i++) {
    // ID 8-indexda
    if (data[i][8] == rowId) {
      sheet.deleteRow(i + 1);
      return "Deleted";
    }
  }
  return "Not Found";
}

// ==========================================
// 5. JADVAL UCHUN MA'LUMOT
// ==========================================
// ==========================================
// 5. JADVAL UCHUN MA'LUMOT (TUZATILDI)
// ==========================================
function getTransactionsList() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kirim Chiqim');
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  // Ma'lumotlarni olamiz
  var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();
  
  var result = [];
  for (var i = data.length - 1; i >= 0; i--) {
    var row = data[i];
    
    // --- SANANI FORMATLASH (MUHIM) ---
    // Agar sana "Date" obyekti bo'lsa, uni chiroyli stringga aylantiramiz
    var dateVal = row[0];
    var dateStr = "";
    
    if (dateVal instanceof Date) {
        var y = dateVal.getFullYear();
        var m = ("0" + (dateVal.getMonth() + 1)).slice(-2);
        var d = ("0" + dateVal.getDate()).slice(-2);
        dateStr = y + "-" + m + "-" + d; // 2025-12-31 formatida
    } else {
        dateStr = String(dateVal); // Shunchaki tekst bo'lsa
    }

    result.push({
      date: dateStr,      // Formatlangan sana
      firm: row[1],
      point: row[2],
      cat: row[3],        // Kategoriya
      pay: row[4],
      kirim: row[5],
      chiqim: row[6],
      note: row[7],
      rowId: row[8]
    });
  }
  return result;
}

// ==========================================
// 6. DASHBOARD METRIKALARI (FILTR BILAN)
// ==========================================
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

  // --- BLACKLIST (KIRITILMAYDIGANLAR) ---
  // Bu nomdagi "Savdo nuqtalari" statistikaga kirmaydi
  var blacklist = ["DONIYOR AKA", "BOSHQA", "DIREKTOR", "KASSA"];

  data.forEach(function(row) {
    var dateVal = new Date(row[0]);
    var point = row[2] ? row[2].toString().trim() : "";
    var category = row[3];
    var kirim = parseFloat(row[5]) || 0;
    var chiqim = parseFloat(row[6]) || 0;

    if (dateVal.getFullYear() === currentYear) {
       var m = dateVal.getMonth();
       
       // --- FILTR ---
       // 1. Kirim: Faqat "Savdo tushumi" bo'lsa VA Blacklistda bo'lmasa
       var isValidIncome = (kirim > 0 && category === "Savdo tushumi" && blacklist.indexOf(point.toUpperCase()) === -1);
       
       // 2. Chiqim: O'tkazma emasligi
       var isValidExpense = (chiqim > 0 && category !== "O'tkazma" && category !== "Transfer");

       // KIRIM HISOBLASH
       if (isValidIncome) {
           stats.yearIncome[m] += kirim;
           if (m === currentMonth) stats.monthIncome += kirim;
           
           if (!stats.pointsData[point]) stats.pointsData[point] = 0;
           stats.pointsData[point] += kirim;
       }

       // CHIQIM HISOBLASH
       if (isValidExpense) {
           stats.yearExpense[m] += chiqim;
           if (m === currentMonth) stats.monthExpense += chiqim;
       }
    }
  });

  stats.monthProfit = stats.monthIncome - stats.monthExpense;
  return stats;
}

// ==========================================
// 7. REAL VAQT BALANSI (KASSA)
// ==========================================
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
    
    // O'TKAZMA
    if (category === "O'tkazma" || category === "Transfer") {
       var source = row[9]; 
       var dest = row[10];
       
       // Eski ma'lumotlar uchun (agar col 10-11 bo'sh bo'lsa)
       if (!source && payType.includes('->')) {
           var parts = payType.split('->');
           source = parts[0].trim();
           dest = parts[1].trim();
       }

       if (balance.hasOwnProperty(source)) balance[source] -= chiqim; 
       if (balance.hasOwnProperty(dest)) balance[dest] += kirim;
    } 
    // ODDIY KIRIM / CHIQIM
    else {
       if (balance.hasOwnProperty(payType)) {
          balance[payType] += (kirim - chiqim);
       }
    }
  });

  return balance;
}
