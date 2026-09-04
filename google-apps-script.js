/**
 * GOOGLE APPS SCRIPT CODE FOR MASTER PAYROLL GOOGLE SHEET
 * 
 * Instructions:
 * 1. Open your Master Google Sheet in Google Sheets.
 * 2. Click Extensions > Apps Script.
 * 3. Delete any default code in Code.gs and paste the code below.
 * 4. Click "Deploy" > "New deployment".
 * 5. Select type: "Web app".
 * 6. Set Description: "Payroll Sync API"
 * 7. Set "Execute as": "Me"
 * 8. Set "Who has access": "Anyone"  <-- CRITICAL so the browser app can append without oauth popups.
 * 9. Click "Deploy", authorize permissions when prompted, and COPY the Web App URL (ends in /exec).
 * 10. Paste that URL into the Excel/CSV Splitter app under "Google Sheet Webhook URL".
 */

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Ensure or create "Payroll Summary" tab
    var summarySheet = ss.getSheetByName("Payroll Summary");
    if (!summarySheet) {
      summarySheet = ss.insertSheet("Payroll Summary");
      summarySheet.appendRow([
        "Timestamp",
        "Pay Cycle",
        "Trainer",
        "Transactions / Sessions",
        "Total Sales ($)",
        "Commission Rate (%)",
        "Trainer Payout ($)",
        "Gym Net Retained ($)"
      ]);
      summarySheet.getRange("A1:H1").setFontWeight("bold").setBackground("#d9ead3");
      summarySheet.setFrozenRows(1);
    }
    
    var timestamp = new Date();
    var cycle = data.payCycle || "N/A";
    
    // Append trainer payroll breakdown rows
    if (data.trainers && data.trainers.length) {
      data.trainers.forEach(function(t) {
        summarySheet.appendRow([
          timestamp,
          cycle,
          t.name,
          t.count,
          t.totalSales,
          t.commissionRate + "%",
          t.payout,
          t.gymRetained
        ]);
      });
    }
    
    // Append member rollup row if provided
    if (data.memberTotal !== undefined) {
      summarySheet.appendRow([
        timestamp,
        cycle,
        "member (General)",
        data.memberCount || 0,
        data.memberTotal,
        "0%",
        0,
        data.memberTotal
      ]);
    }
    
    // 2. Optionally create/update a tab for the current pay cycle with all raw transactions
    if (data.rawTransactions && data.rawTransactions.length) {
      var tabName = ("Cycle_" + cycle).replace(/[\/\\?*:[\]]/g, "_").slice(0, 30);
      var cycleSheet = ss.getSheetByName(tabName);
      if (!cycleSheet) {
        cycleSheet = ss.insertSheet(tabName);
        if (data.headers && data.headers.length) {
          var headerRow = data.headers.slice();
          headerRow.unshift("Assigned Group");
          cycleSheet.appendRow(headerRow);
          cycleSheet.getRange(1, 1, 1, headerRow.length).setFontWeight("bold").setBackground("#e6b8af");
          cycleSheet.setFrozenRows(1);
        }
      }
      
      var rowsToInsert = data.rawTransactions.map(function(row) {
        var rowArr = [row._assignedGroup || ""];
        data.headers.forEach(function(h) {
          rowArr.push(row[h] !== undefined ? row[h] : "");
        });
        return rowArr;
      });
      
      if (rowsToInsert.length) {
        cycleSheet.getRange(cycleSheet.getLastRow() + 1, 1, rowsToInsert.length, rowsToInsert[0].length).setValues(rowsToInsert);
      }
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Payroll data appended successfully for cycle " + cycle
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
