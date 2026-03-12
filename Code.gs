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
    MailApp.sendEmail(email, "✅ Equipment Confirmed","Dear Clinical Instructors," + "\nGood day. This email serves as a confirmation that the Equipment and Consumables Borrowing Google Form has been successfully completed and submitted. The details of the requested equipment or consumable supplies have been recorded in the system for documentation and monitoring purposes." + "\n\nThank you for your time and continued support in maintaining an organized and efficient inventory management system." + "\n\nBorrowed: " + cleanItem);
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

  // 1. Process Equipment (A-D)
  var eqData = invSheet.getRange(11, 1, lastRow - 10, 4).getValues();
  eqData.forEach(function(row, i) {
    if (row[0]) { 
      var label = row[0] + (row[3] > 0 ? " (" + row[3] + " available)" : " (OUT OF STOCK)");
      eqChoices.push(label);
      invSheet.getRange(i + 11, 1, 1, 4).setBackground(row[3] <= 3 ? "#ff9999" : null);
    }
  });

  // 2. Process Consumables (F-I)
  var consData = invSheet.getRange(11, 6, lastRow - 10, 5).getValues(); 
  var now = new Date();
  now.setHours(0, 0, 0, 0); // Normalize today's time

  consData.forEach(function(row, i) {
    if (row[0]) { 
      var label = row[0] + (row[3] > 0 ? " (" + row[3] + " available)" : " (OUT OF STOCK)");
      consChoices.push(label);
      
      var currentRow = i + 11;
      var expiryDate = row[4] instanceof Date ? new Date(row[4]) : null;
      var expiryCell = invSheet.getRange(currentRow, 10); // Target ONLY Column J
      var expiryColor = null;

      if (expiryDate) {
        expiryDate.setHours(0, 0, 0, 0); // Normalize expiry time
        var diffInMs = expiryDate.getTime() - now.getTime();
        var diffInDays = Math.ceil(diffInMs / (1000 * 60 * 60 * 24));

        if (diffInDays < 0) {
          expiryColor = "#ff4d4d"; // RED: Expired
        } else if (diffInDays === 0) {
          expiryColor = "#ffa500"; // ORANGE: Expires Today
        } else if (diffInDays <= 30) {
          expiryColor = "#ffff00"; // YELLOW: 30 days remaining
        }
      }

      expiryCell.setBackground(expiryColor);

      // Keep your separate logic for highlighting the whole row if stock is low
      if (row[3] <= 3 || row[3] === "OUT OF STOCK") {
        invSheet.getRange(currentRow, 6, 1, 4).setBackground("#ff9999"); // Rows F-I for Low Stock
      } else {
        invSheet.getRange(currentRow, 6, 1, 4).setBackground(null); // Reset F-I if stock is fine
      }
    }
  });

  // 3. Create Return List 
  var allChoices = [...new Set(eqChoices.concat(consChoices))];

  // 4. Update Forms with unique values only
  if (eqChoices.length > 0) {
    formEq.getItems(FormApp.ItemType.MULTIPLE_CHOICE)[0].asMultipleChoiceItem().setChoiceValues([...new Set(eqChoices)]);
  }
  
  if (consChoices.length > 0) {
    formCons.getItems(FormApp.ItemType.MULTIPLE_CHOICE)[0].asMultipleChoiceItem().setChoiceValues([...new Set(consChoices)]);
  }

  if (allChoices.length > 0) {
    formRet.getItems(FormApp.ItemType.MULTIPLE_CHOICE)[0].asMultipleChoiceItem().setChoiceValues(allChoices);
  }
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

/**
 * Handles the logic from the Sidebar UI to modify the Sheet
 */
function modifyInventory(action, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const invSheet = ss.getSheetByName("Inventory");
  
  // Define columns based on category
  // Equipment: A(1) to D(4) | Consumables: F(6) to I(9)
  const isEq = data.type.includes("Equipment");
  const startCol = isEq ? 1 : 6;
  const nameColOffset = 0; // The name is in the first col of the block
  const stockColOffset = 1; // Stock is in the second col (B or G)
  
  const lastRow = invSheet.getLastRow();
  const range = invSheet.getRange(11, startCol, lastRow - 10, 2); // Get Name and Stock columns
  const values = range.getValues();
  
  let targetRow = -1;
  const searchName = data.name.trim().toLowerCase();

  // Find the item if it exists
  for (let i = 0; i < values.length; i++) {
    if (values[i][0].toString().toLowerCase().trim() === searchName) {
      targetRow = i + 11;
      break;
    }
  }

  try {
    if (action === 'add') {
      if (targetRow !== -1) return "⚠️ Item already exists! Use 'Update' instead.";
      // Find first empty row in that specific column block
      let rowToAdd = 11;
      while (invSheet.getRange(rowToAdd, startCol).getValue() !== "") {
        rowToAdd++;
      }
      invSheet.getRange(rowToAdd, startCol).setValue(data.name);
      invSheet.getRange(rowToAdd, startCol + stockColOffset).setValue(data.stock);
      
    } else if (action === 'update') {
      if (targetRow === -1) return "❌ Item not found.";
      invSheet.getRange(targetRow, startCol + stockColOffset).setValue(data.stock);
      
    } else if (action === 'remove') {
      if (targetRow === -1) return "❌ Item not found.";
      // Clear only the 4-column wide block so we don't delete the other category
      invSheet.getRange(targetRow, startCol, 1, 4).clearContent();
    }

    // CRITICAL: Refresh the Forms and Colors immediately
    updateAllFormChoicesAndColors(invSheet);
    return "✅ Success: Inventory & Forms updated!";
    
  } catch (e) {
    return "❌ Error: " + e.toString();
  }
}

// Function to specifically open the Item Manager Sidebar
function showItemManager() {
  var html = HtmlService.createHtmlOutputFromFile('ItemManager') // Ensure your HTML file is named 'ItemManager'
    .setTitle('NCF Item Manager')
    .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
}
