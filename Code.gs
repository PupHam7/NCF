/*
❤️❤️PupHam❤️❤️
*/

// 1. MENU SETUP
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('⚙️ Management')
    .addItem('📦 Open Item Manager', 'showItemManager')
    .addItem('📃 Open Attendance UI', 'showSidebar')
    .addSeparator()
    .addItem('🔄 Refresh Inventory & Forms', 'refreshInventory')
    .addToUi();

  ui.createMenu('🔍 Search')
    .addItem('Search C.I. in Logs', 'runCISearch')
    .addToUi();

}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Student Attendance')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}

// 2. ATTENDANCE LOGIC
function handleAttendance(data, type) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Attendance");
  var lastRow = sheet.getLastRow();
  var range = sheet.getRange(1, 1, lastRow || 2, 6);
  var values = range.getValues();
  var now = new Date();

  if (type === 'in') {
    // --- 11-DIGIT CHECKER ---
    // Removes any spaces or dashes and checks if exactly 11 digits remain
    var cleanContact = data.contact.toString().replace(/[^0-9]/g, "");
    if (cleanContact.length !== 11) {
      return "❌ Error: Contact number must be exactly 11 digits!";
    }

    for (var i = 0; i < values.length; i++) {
      if (values[i][3] === data.email && (!values[i][5] || values[i][5] === "")) {
        return "⚠️ This email is already clocked in!";
      }
    }

    // Use the cleaned 11-digit number for the record
    sheet.appendRow([now, data.dept, data.name, data.email, "'" + cleanContact, ""]);
    sheet.getRange(sheet.getLastRow(), 1).setNumberFormat("M/d/yyyy H:mm:ss AM/PM");
    return "✅ Clocked In: " + now.toLocaleTimeString();
  }

  if (type === 'out') {
    for (var i = values.length - 1; i >= 0; i--) {
      if (values[i][3] === data.email && (!values[i][5] || values[i][5] === "")) {
        var outCell = sheet.getRange(i + 1, 6);
        outCell.setValue(now).setNumberFormat("M/d/yyyy H:mm:ss AM/PM");
        return "👋 Clocked Out: " + now.toLocaleTimeString();
      }
    }
    return "❌ No active 'Clock In' found for " + data.email;
  }
}

// 3. MAIN INVENTORY & FORM SYNC
function updateAndGuardInventory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invSheet = ss.getSheetByName("Inventory");
  var fr1 = ss.getSheetByName("Form Responses 1"); // Equipment
  var fr2 = ss.getSheetByName("Form Responses 2"); // Returns
  var fr3 = ss.getSheetByName("Form Responses 3"); // Consumables
  var ownerEmail = ss.getOwner().getEmail();

  // Determine latest form submission
  var t1 = getLatestTimestamp(fr1);
  var t2 = getLatestTimestamp(fr2);
  var t3 = getLatestTimestamp(fr3);
  var latest = Math.max(t1, t2, t3);

  // Email Notification Logic
  if (latest === t3 && t3 > 0) {
    processConsumableBorrow(fr3, invSheet, ownerEmail);
  } else if (latest === t1 && t1 > 0) {
    processEquipmentBorrow(fr1, ownerEmail);
  } else if (latest === t2 && t2 > 0) {
    processReturn(fr2, ownerEmail);
  }

  updateAllFormChoicesAndColors(invSheet);
}

// 4. PROCESS CONSUMABLES (With Expiry)
function processConsumableBorrow(sheet, invSheet, ownerEmail) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(lastRow, 1, 1, 4).getValues()[0];
  var email = data[1];
  var rawItem = data[2];
  var cleanItem = rawItem.split(" (")[0];

  if (rawItem.indexOf("OUT OF STOCK") !== -1) {
    MailApp.sendEmail(ownerEmail, "⚠️ OOS Alert", email + " tried to get " + cleanItem);
    sheet.deleteRow(lastRow);
    return;
  }

  // Lookup Expiry Date in Col J (Index 4 of the F-J range)
  var expiry = "N/A";
  var consData = invSheet.getRange(11, 6, invSheet.getLastRow() - 10, 5).getValues();
  for (var i = 0; i < consData.length; i++) {
    if (consData[i][0] === cleanItem) {
      expiry = consData[i][4] instanceof Date ? consData[i][4].toLocaleDateString() : consData[i][4];
      break;
    }
  }

  MailApp.sendEmail({
    to: email,
    cc: ownerEmail,
    subject: "✅ Consumable Borrowed",
    body: "Dear Clinical Instructors," + "\nGood day. This email serves as a confirmation that the Equipment and Consumables Borrowing Google Form has been successfully completed and submitted. The details of the requested equipment or consumable supplies have been recorded in the system for documentation and monitoring purposes." + "\n\nThank you for your time and continued support in maintaining an organized and efficient inventory management system." + "\n\nBorrowed: " + cleanItem + "\n⚠️ Expiry Date: " + expiry
  });
}

// 5. PROCESS EQUIPMENT 
function processEquipmentBorrow(sheet, ownerEmail) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(lastRow, 1, 1, 4).getValues()[0];
  var rawItem = data[1];
  var email = data[3];
  var cleanItem = rawItem.split(" (")[0];

  if (rawItem.indexOf("OUT OF STOCK") !== -1) {
    MailApp.sendEmail(ownerEmail, "⚠️ OOS Alert", email + " tried to borrow " + cleanItem);
    sheet.deleteRow(lastRow);
  } else if (email) {
    MailApp.sendEmail(email, "✅ Equipment Confirmed", "Dear Clinical Instructors," + "\nGood day. This email serves as a confirmation that the Equipment and Consumables Borrowing Google Form has been successfully completed and submitted. The details of the requested equipment or consumable supplies have been recorded in the system for documentation and monitoring purposes." + "\n\nThank you for your time and continued support in maintaining an organized and efficient inventory management system." + "\n\nBorrowed: " + cleanItem);
  }
}

// 6. PROCESS RETURNS
function processReturn(sheet, ownerEmail) {
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(lastRow, 1, 1, 3).getValues()[0];
  var email = data[1];
  var rawItem = data[2];
  var cleanItem = rawItem.split(" (")[0];

  MailApp.sendEmail(email, "🔄 Return Confirmed", "Dear Clinical Instructors," + "\nGood day. This email serves as a confirmation that the Equipment and Consumables Return Google Form has been successfully completed and submitted. The return details, including the items returned and the corresponding date and time, have been recorded in the system for documentation and inventory monitoring." + "\n\nThank you for your time and cooperation in helping maintain an organized and efficient inventory management system." + "\n\nThank you for returning: " + cleanItem);
}

function updateAllFormChoicesAndColors(invSheet) {
  var formEq = FormApp.openById("1rSJn5YLn82-66SQbQTHEkEjWMEIUQxwtPQPl27hlw3Q");
  var formRet = FormApp.openById("1NlpLyxbz6hDxceATUjCc0QZgATXr2fCOOHZWlO44giI");
  var formCons = FormApp.openById("1oTABVbzviWQ-x-_ufVm_1kq0KffLzR57-UpyMRCi4E8");

  var lastRow = invSheet.getLastRow();
  if (lastRow < 11) return;

  var eqChoices = [], consChoices = [];
  var now = new Date();
  now.setHours(0, 0, 0, 0);

  // --- 1. Process Equipment (A-E) ---
  // Range: Col A to E (5 columns)
  var eqData = invSheet.getRange(11, 1, lastRow - 10, 5).getValues();
  var currentEqRoom = "General"; 

  eqData.forEach(function (row, i) {
    var itemName = row[0] ? row[0].toString().trim() : "";
    if (itemName === "") return;

    if (itemName.toLowerCase().includes("room")) {
      currentEqRoom = itemName; 
    } else {
      var label = "[" + currentEqRoom + "] " + itemName + (row[3] > 0 ? " (" + row[3] + " available)" : " (OUT OF STOCK)");
      eqChoices.push(label);
      // Highlight A-D based on Available Unit (Col D / Index 3)
      invSheet.getRange(i + 11, 1, 1, 5).setBackground(row[3] <= 3 ? "#ff9999" : null);
    }
  });

  // --- 2. Process Consumables (G-L) ---
  // Range: Col G to L (6 columns). G=0, H=1, I=2, J=3, K=4, L=5
  var consData = invSheet.getRange(11, 7, lastRow - 10, 6).getValues();
  var currentConsRoom = "General";

  consData.forEach(function (row, i) {
    var itemName = row[0] ? row[0].toString().trim() : "";
    if (itemName === "") return;

    if (itemName.toLowerCase().includes("room")) {
      currentConsRoom = itemName;
    } else {
      var label = "[" + currentConsRoom + "] " + itemName + (row[3] > 0 ? " (" + row[3] + " available)" : " (OUT OF STOCK)");
      consChoices.push(label);

      var currentRow = i + 11;
      // Expiry Date is in Column L (Index 5 of this range)
      var expiryDate = row[5] instanceof Date ? new Date(row[5]) : null;
      var expiryCell = invSheet.getRange(currentRow, 12); // Column L is 12
      var expiryColor = null;

      if (expiryDate) {
        expiryDate.setHours(0, 0, 0, 0);
        var diffInMs = expiryDate.getTime() - now.getTime();
        var diffInDays = Math.ceil(diffInMs / (1000 * 60 * 60 * 24));
        if (diffInDays < 0) expiryColor = "#ff4d4d"; // Expired (Red)
        else if (diffInDays === 0) expiryColor = "#ffa500"; // Today (Orange)
        else if (diffInDays <= 30) expiryColor = "#ffff00"; // Within 30 days (Yellow)
      }
      expiryCell.setBackground(expiryColor);

      // Highlight G-J based on Available Unit (Col J / Index 3)
      if (row[3] <= 3) {
        invSheet.getRange(currentRow, 7, 1, 5).setBackground("#ff9999");
      } else {
        invSheet.getRange(currentRow, 7, 1, 5).setBackground(null);
      }
    }
  });


  var allChoices = [...new Set(eqChoices.concat(consChoices))];


  if (eqChoices.length > 0) formEq.getItems(FormApp.ItemType.MULTIPLE_CHOICE)[0].asMultipleChoiceItem().setChoiceValues([...new Set(eqChoices)]);
  if (consChoices.length > 0) formCons.getItems(FormApp.ItemType.MULTIPLE_CHOICE)[0].asMultipleChoiceItem().setChoiceValues([...new Set(consChoices)]);
  if (allChoices.length > 0) formRet.getItems(FormApp.ItemType.MULTIPLE_CHOICE)[0].asMultipleChoiceItem().setChoiceValues(allChoices);
}

// 8. OVERDUE REMINDERS (Time Trigger)
function sendOverdueReminders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var fr1 = ss.getSheetByName("Form Responses 1");
  var fr2 = ss.getSheetByName("Form Responses 2");
  var borrows = fr1.getDataRange().getValues();
  var returns = fr2.getDataRange().getValues();
  var now = new Date();

  for (var i = 1; i < borrows.length; i++) {
    var timestamp = new Date(borrows[i][0]);
    var item = borrows[i][1].split(" (")[0];
    var email = borrows[i][3];

    if ((now - timestamp) > (24 * 60 * 60 * 1000)) {
      var isReturned = returns.some(r => r[1] === email && r[2].indexOf(item) !== -1);
      if (!isReturned && email) {
        MailApp.sendEmail(email, "⚠️ OVERDUE: Return Reminder", "Dear Clinical Instructors," +
          "\n\nGood day. This email serves as a notification that the borrowed equipment or consumable supplies recorded in the system have not yet been returned within the expected one-day period. According to the borrowing record submitted through the Google Form, the items remain unreturned as of this time. We kindly request your assistance in verifying the status of the borrowed materials and facilitating their return at the earliest convenience." + "\nYour prompt attention to this matter will help ensure the proper monitoring and availability of equipment and consumable supplies for other users. Thank you for your understanding and cooperation."
          + "\n\nThe item '" + item + "' is overdue.");
      }
    }
  }
}

function getLatestTimestamp(sheet) {
  var lastRow = sheet.getLastRow();
  return lastRow > 1 ? new Date(sheet.getRange(lastRow, 1).getValue()).getTime() : 0;
}

function runCISearch() {

  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var logSheet = ss.getSheetByName("Logs");

  if (!logSheet) {
    ui.alert("❌ Error: The 'Logs' sheet was not found.");
    return;
  }

  var response = ui.prompt(
    'Strict CI Search',
    'Enter the EXACT C.I. Name to highlight logs:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  var nameToSearch = response.getResponseText().trim().toLowerCase();
  var lastRow = logSheet.getLastRow();

  if (lastRow < 4) {
    ui.alert("⚠️ No data found in Logs to search.");
    return;
  }

  var range = logSheet.getRange(4, 1, lastRow - 3, 29);
  var values = range.getValues();
  var matchCount = 0;

  // RESET COLORS (original ranges)
  logSheet.getRange(4, 1, lastRow - 3, 9).setBackground("#c3f3ca");   // A-I
  logSheet.getRange(4, 11, lastRow - 3, 9).setBackground("#c3f3ca");  // K-S
  logSheet.getRange(4, 21, lastRow - 3, 9).setBackground("#c3f3ca");  // U-AC

  if (nameToSearch === "") {
    ui.alert("ℹ️ Search cleared.");
    return;
  }

  for (var i = 0; i < values.length; i++) {

    var currentRow = i + 4;

    // Borrow Section
    if (values[i][6] && values[i][6].toString().toLowerCase().trim() === nameToSearch) {
      logSheet.getRange(currentRow, 1, 1, 9).setBackground("#ffff00");
      matchCount++;
    }

    // Return Section
    if (values[i][16] && values[i][16].toString().toLowerCase().trim() === nameToSearch) {
      logSheet.getRange(currentRow, 11, 1, 9).setBackground("#ffff00");
      matchCount++;
    }

    // Consumables Section
    if (values[i][26] && values[i][26].toString().toLowerCase().trim() === nameToSearch) {
      logSheet.getRange(currentRow, 21, 1, 9).setBackground("#ffff00");
      matchCount++;
    }
  }

  if (matchCount > 0) {
    ui.alert("✅ Found " + matchCount + " match(es) for '" + response.getResponseText() + "'.");
  } else {
    ui.alert("❌ No exact matches found for '" + response.getResponseText() + "'.");
  }
}

function refreshInventory() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var invSheet = ss.getSheetByName("Inventory");

  updateAllFormChoicesAndColors(invSheet);

  SpreadsheetApp.getActive().toast(
    "✅ Inventory and Forms refreshed!",
    "System Update",
    3
  );
}


function modifyInventory(action, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName("Inventory");

  const isEq = data.type.includes("Equipment");
  const startCol = isEq ? 1 : 7; // A (1) or G (7)
  const lastRow = invSheet.getLastRow();

  // 1. GLOBAL DUPLICATE CHECK
  const fullRange = invSheet.getRange(11, startCol, Math.max(lastRow - 10, 1), 1).getValues();


  const searchName = data.name.trim().toLowerCase();
  let existingRow = -1;

  for (let i = 0; i < fullRange.length; i++) {
    if (fullRange[i][0].toString().toLowerCase().trim() === searchName) {
      existingRow = i + 11;
      break;
    }
  }

  try {
    if (action === 'add') {
      if (existingRow !== -1) return "⚠️ Item already exists in row " + existingRow;
      if (!data.room) return "❌ Error: Please select a Room.";

      // 2. FIND THE ROOM HEADER
      let roomRow = -1;
      const roomSearch = data.room.trim().toLowerCase();
      const nameColValues = invSheet.getRange(1, startCol, lastRow, 1).getValues();

      for (let r = 0; r < nameColValues.length; r++) {
        if (nameColValues[r][0].toString().toLowerCase().includes(roomSearch)) {
          roomRow = r + 1;
          break;
        }
      }

      if (roomRow === -1) return "❌ Error: Header '" + data.room + "' not found on sheet.";

      // 3. FIND FIRST EMPTY ROW UNDER THAT ROOM
      let rowToAdd = roomRow + 1;
      while (invSheet.getRange(rowToAdd, startCol).getValue() !== "") {
        // If we hit another Room header, we must insert a row to avoid overwriting it
        if (invSheet.getRange(rowToAdd, startCol).getValue().toString().toLowerCase().includes("room")) {
          invSheet.insertRowBefore(rowToAdd);
          break; 
        }
        rowToAdd++;
        if (rowToAdd > 3000) break; 
      }

      // 4. SET DATA & FORMULAS
      // Name and Stock
      invSheet.getRange(rowToAdd, startCol).setValue(data.name);
      invSheet.getRange(rowToAdd, startCol + 1).setValue(data.stock);

      // Shelf (Eq) or Unit (Consumables)



      if (isEq) {
        invSheet.getRange(rowToAdd, 5).setValue(data.shelf); // Column E



      } else {
        invSheet.getRange(rowToAdd, 11).setValue(data.unit); // Column K



      }

      // Formulas
      var borrowedFormula = isEq ? 
        '=SUMIF(Logs!$B$4:$B$991, "*" & A' + rowToAdd + ' & "*", Logs!$A$4:$A$991) - SUMIF(Logs!$L$4:$L$991, "*" & A' + rowToAdd + ' & "*", Logs!$K$4:$K$991)' : 
        '=SUMIF(Logs!$V$4:$V$991, "*" & G' + rowToAdd + ' & "*", Logs!$U$4:$U$991) - SUMIF(Logs!$L$4:$L$991, "*" & G' + rowToAdd + ' & "*", Logs!$K$4:$K$991)';
      
      var availableFormula = isEq ? '=B' + rowToAdd + '- C' + rowToAdd : '=H' + rowToAdd + '- I' + rowToAdd;

      invSheet.getRange(rowToAdd, startCol + 2).setFormula(borrowedFormula);
      invSheet.getRange(rowToAdd, startCol + 3).setFormula(availableFormula);
      updateAllFormChoicesAndColors(invSheet);
      return "✅ Success: Added to " + data.room;

    } else if (action === 'update') {
      if (existingRow === -1) return "❌ Item not found.";
      invSheet.getRange(existingRow, startCol + 1).setValue(data.stock);
      // Update Shelf/Unit during update too
      if (isEq) invSheet.getRange(existingRow, 5).setValue(data.shelf);
      else invSheet.getRange(existingRow, 11).setValue(data.unit);
      updateAllFormChoicesAndColors(invSheet);

      return "✅ Update Successful!";

    } else if (action === 'remove') {
      if (existingRow === -1) return "❌ Item not found.";
      invSheet.getRange(existingRow, startCol, 1, (isEq ? 5 : 6)).clearContent();
      invSheet.getRange(existingRow, startCol, 1, 5).setBackground(null);
        updateAllFormChoicesAndColors(invSheet);
      return "✅ Item Removed.";
    }

    updateAllFormChoicesAndColors(invSheet);


  } catch (e) {
    return "❌ Error: " + e.toString();
  }
}

// Function to specifically open the Item Manager Sidebar
function showItemManager() {
  var html = HtmlService.createHtmlOutputFromFile('ItemManager') 
    .setTitle('NCF Item Manager')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}














function generateMonthlyEquipment() {
  const ss = SpreadsheetApp.getActive();
  const logSheet = ss.getSheetByName("Logs");
  const reportSheet = ss.getSheetByName("Monthly Report");
  
  // 1. Setup Date Range (Current Month)
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0);
  const end = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);

  // 2. Get all Log Data
  const lastLogRow = logSheet.getLastRow();
  if (lastLogRow < 4) return;
  const logData = logSheet.getRange(4, 1, lastLogRow - 3, 13).getValues(); 

  // 3. STEP ONE: DISCOVER UNIQUE NAMES FIRST
  let itemNamesSet = new Set();
  
  logData.forEach(row => {
    let bItem = cleanItemName(row[1]); // Borrowing Item (Col B)
    let bTS = row[2];                 // Borrowing TS (Col C)
    
    if (bItem && bTS instanceof Date && bTS >= start && bTS <= end) {
      itemNamesSet.add(bItem);
    }
  });

  // 4. STEP TWO: SORT NAMES ALPHABETICALLY
  let sortedNames = Array.from(itemNamesSet).sort((a, b) => a.localeCompare(b));

  // 5. STEP THREE: CALCULATE TOTALS BASED ON SORTED NAMES
  // This ensures Index 0 of Names always matches Index 0 of Totals
  let finalData = sortedNames.map(itemName => {
    let borrowedTotal = 0;
    let returnedTotal = 0;

    logData.forEach(row => {
      // Check Borrowing (Cols A, B, C)
      if (cleanItemName(row[1]) === itemName) {
        let bTS = row[2];
        if (bTS instanceof Date && bTS >= start && bTS <= end) {
          borrowedTotal += (parseFloat(row[0]) || 0);
        }
      }

      // Check Returning (Cols K, L, M)
      if (cleanItemName(row[11]) === itemName) {
        let rTS = row[12];
        if (rTS instanceof Date && rTS >= start && rTS <= end) {
          returnedTotal += (parseFloat(row[10]) || 0);
        }
      }
    });

    return [itemName, borrowedTotal, returnedTotal];
  });

  // 6. WRITE TO SHEET
  // Clear old content from B7:D
  const lastReportRow = reportSheet.getLastRow();
  if (lastReportRow >= 7) {
    reportSheet.getRange(7, 2, lastReportRow - 6, 3).clearContent();
  }

  if (finalData.length > 0) {
    reportSheet.getRange(7, 2, finalData.length, 3).setValues(finalData);
  }
  
  ss.toast("✅ Report generated: Items sorted alphabetically before calculation.");
}


function generateMonthlyConsumables() {
  const ss = SpreadsheetApp.getActive();
  const logSheet = ss.getSheetByName("Logs");
  const reportSheet = ss.getSheetByName("Monthly Report");
  
  // 1. Setup Date Range (Current Month)
  const now = new Date();
  const start = new Date(now.getFullYear(), now.getMonth(), 1, 0, 0, 0);
  const end = new Date(now.getFullYear(), now.getMonth() + 1, 0, 23, 59, 59);

  // 2. Get all Log Data (Scanning up to Column W for Consumables)
  const lastLogRow = logSheet.getLastRow();
  if (lastLogRow < 4) return;
  const logData = logSheet.getRange(4, 1, lastLogRow - 3, 23).getValues(); 

  // 3. STEP ONE: DISCOVER UNIQUE CONSUMABLE NAMES FIRST
  let itemNamesSet = new Set();
  
  logData.forEach(row => {
    let cItem = cleanItemName(row[21]); // Consumable Item (Col V)
    let cTS = row[22];                 // Consumable TS (Col W)
    
    if (cItem && cTS instanceof Date && cTS >= start && cTS <= end) {
      itemNamesSet.add(cItem);
    }
  });

  // 4. STEP TWO: SORT NAMES ALPHABETICALLY
  let sortedNames = Array.from(itemNamesSet).sort((a, b) => a.localeCompare(b));

  // 5. STEP THREE: CALCULATE TOTALS BASED ON SORTED NAMES
  let finalData = sortedNames.map(itemName => {
    let usedTotal = 0;
    let returnedTotal = 0;

    logData.forEach(row => {
      // Check Consumable Usage (Logs Cols U, V, W)
      if (cleanItemName(row[21]) === itemName) {
        let cTS = row[22];
        if (cTS instanceof Date && cTS >= start && cTS <= end) {
          usedTotal += (parseFloat(row[20]) || 0); // Col U
        }
      }

      // Check Returning (Logs Cols K, L, M)
      if (cleanItemName(row[11]) === itemName) {
        let rTS = row[12];
        if (rTS instanceof Date && rTS >= start && rTS <= end) {
          returnedTotal += (parseFloat(row[10]) || 0); // Col K
        }
      }
    });

    return [itemName, usedTotal, returnedTotal];
  });

  // 6. WRITE TO SHEET (Columns F, G, H starting at row 7)
  // Clear old content from F7:H
  const lastReportRow = reportSheet.getLastRow();
  if (lastReportRow >= 7) {
    reportSheet.getRange(7, 6, lastReportRow - 6, 3).clearContent();
  }

  if (finalData.length > 0) {
    reportSheet.getRange(7, 6, finalData.length, 3).setValues(finalData);
  }
  
  ss.toast("✅ Consumables report generated and sorted alphabetically.");
}

/**
 * Helper to remove room numbers and "(available)" text
 */
function cleanItemName(str) {
  if (!str) return "";
  // Removes [Room XXX] and everything after the first "("
  return str.replace(/\[.*?\]\s*/g, '').split('(')[0].trim();
}















function automatedMonthlyReportingCycle() {
  const ss = SpreadsheetApp.getActive();
  const ownerEmail = "cydjian1@gmail.com"; 
  
  // Calculate the month that just ended
  const now = new Date();
  const reportDate = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const monthLabel = Utilities.formatDate(reportDate, Session.getScriptTimeZone(), "MMMM yyyy");

  // 1. Process all report data
  updateMonthlyEquipmentBorrowing(reportDate);
  updateMonthlyConsumables(reportDate);
  categorizeReturnedItems(reportDate);
  consolidateInstructorsForFormatting(reportDate);
  
  // 2. Email the specific "Monthly Report" tab as an Excel attachment
  emailReportAsExcel("Monthly Report", ownerEmail, "Inventory Report: " + monthLabel);
  
  ss.toast("Monthly report cycle completed successfully!");
}


function consolidateInstructorsForFormatting(targetDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName("Logs");
  const formatSheet = ss.getSheetByName("For formatting");
  
  // 1. Setup Timeframe (Current Month)
  const now = new Date();
  const firstDay = new Date(now.getFullYear(), now.getMonth(), 1);
  const lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0);

  // 2. Define the Column Indices for Instructors and Timestamps
  // Borrowing: Instructor Col G (6), Timestamp Col C (2)
  // Returning: Instructor Col Q (16), Timestamp Col M (12)
  // Consumables: Instructor Col AA (26), Timestamp Col W (22)
  
  const lastRow = logSheet.getLastRow();
  if (lastRow < 4) return;
  const data = logSheet.getRange(4, 1, lastRow - 3, 27).getValues(); 
  
  let instructorList = [];

  // 3. Extract names from all three sections
  data.forEach(row => {
    // Section 1: Borrowing
    if (isValidEntry(row[2], firstDay, lastDay) && row[6]) instructorList.push([row[6]]);
    
    // Section 2: Returning
    if (isValidEntry(row[12], firstDay, lastDay) && row[16]) instructorList.push([row[16]]);
    
    // Section 3: Consumables
    if (isValidEntry(row[22], firstDay, lastDay) && row[26]) instructorList.push([row[26]]);
  });

  // 4. Update the "For formatting" sheet
  formatSheet.clearContents(); // Clear old data
  if (instructorList.length > 0) {
    formatSheet.getRange(1, 1, instructorList.length, 1).setValues(instructorList);
  }

  ss.toast("Instructors consolidated to 'For formatting' sheet!");
}

/**
 * EXCEL EXPORT LOGIC: Isolates the report and emails it
 */
function emailReportAsExcel(tabName, recipient, subject) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(tabName);
  
  // Create a temporary spreadsheet to hold only the report tab
  const tempSS = SpreadsheetApp.create("Temp_Monthly_Report");
  sheet.copyTo(tempSS).setName(tabName);
  tempSS.deleteSheet(tempSS.getSheets()[0]); // Delete default Sheet1
  
  const fileId = tempSS.getId();
  SpreadsheetApp.flush();
  
  // Convert the temp file to an Excel blob
  const url = "https://docs.google.com/spreadsheets/d/" + fileId + "/export?format=xlsx";
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, { headers: { 'Authorization': 'Bearer ' + token } });
  const blob = response.getBlob().setName(subject + ".xlsx");
  
  // Send the email with the attachment
  GmailApp.sendEmail(recipient, subject, "Please find attached the inventory report for the previous month.", {
    attachments: [blob]
  });
  
  // Move temp file to Trash
  DriveApp.getFileById(fileId).setTrashed(true);
}

function isValidEntry(dateVal, start, end) {
  return (dateVal instanceof Date && dateVal >= start && dateVal <= end);
}
