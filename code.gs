function linkEpidFiles() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const folderId = '1PEpcKLEZS_beDhG6yOW1OgZYMWTO1QPS';
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();

  const fileMap = {};

  while (files.hasNext()) {
    const file = files.next();
    fileMap[file.getName()] = file.getUrl();
  }

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const epid = data[i][0];
    let matchFound = false;

    for (const fileName in fileMap) {
      if (fileName.includes(epid)) {
        const url = fileMap[fileName];
        sheet.getRange(i + 1, 2).setFormula(`=HYPERLINK("${url}", "Open File")`);
        sheet.getRange(i + 1, 3).setValue("✅ Linked");
        matchFound = true;
        break;
      }
    }

    if (!matchFound) {
      sheet.getRange(i + 1, 2).setValue(""); // Clear any old link
      sheet.getRange(i + 1, 3).setValue("❌ File not found");
    }
  }
}
function formatSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const range = sheet.getDataRange();
  const lastRow = sheet.getLastRow();

  // Header styling
  const header = sheet.getRange(1, 1, 1, 3);
  header.setFontWeight("bold").setBackground("#f0f0f0").setFontSize(12);
  sheet.setColumnWidth(1, 200); // EPID
  sheet.setColumnWidth(2, 250); // File link
  sheet.setColumnWidth(3, 150); // Status

  // Align all cells
  range.setVerticalAlignment("middle");

  // Status coloring
  for (let i = 2; i <= lastRow; i++) {
    const statusCell = sheet.getRange(i, 3);
    const status = statusCell.getValue();

    if (status === "✅ Linked") {
      statusCell.setBackground("#d9ead3"); // light green
    } else if (status === "❌ File not found") {
      statusCell.setBackground("#f4cccc"); // light red
    } else {
      statusCell.setBackground(null); // reset
    }
  }

  // Optional: Add borders to whole table
  range.setBorder(true, true, true, true, true, true);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🔄 File Linker")
    .addItem("Link EPID Files", "linkEpidFiles")
    .addItem("Format Sheet", "formatSheet")
    .addToUi();
}
