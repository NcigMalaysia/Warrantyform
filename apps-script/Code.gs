// Replace the existing doPost and jsonResponse functions in the warranty
// Apps Script project, then deploy a new version of its existing web app.
// Column M of the User and Seller tabs is reserved for Claim ID.
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const category = String(data.category || '').trim();
    const claimId = String(data.claimId || '').trim();

    if (category !== 'User' && category !== 'Seller') {
      throw new Error('Invalid category: ' + category);
    }
    if (!/^NCIG-\d+-[A-Z0-9]{6}$/.test(claimId)) {
      throw new Error('Invalid Claim ID.');
    }

    const sheet = spreadsheet.getSheetByName(category);
    if (!sheet) {
      throw new Error('Tab "' + category + '" tidak dijumpai.');
    }
    if (!Array.isArray(data.items) || data.items.length === 0) {
      throw new Error('Tiada produk diterima.');
    }

    const rows = data.items.map(function (item) {
      return [
        data.date || '',       // A: Date
        data.name || '',       // B: Full Name
        data.phone || '',      // C: Phone Number
        data.email || '',      // D: Email
        item.product || '',    // E: Product
        item.variant || '',    // F: Variant
        item.batch || '',      // G: Batch / Serial Number
        item.defect || '',     // H: Defect
        data.address || '',    // I: Address
        category,              // J: Category
        '',                    // K: Status
        '',                    // L: Tracking Number
        claimId                // M: Claim ID
      ];
    });

    // Serialize submissions so two customers cannot reserve the same next row.
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      if (sheet.getMaxColumns() < 13) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), 13 - sheet.getMaxColumns());
      }
      const headerCell = sheet.getRange(1, 13);
      const header = String(headerCell.getValue() || '').trim();
      if (header && header !== 'Claim ID') {
        throw new Error('Column M is already used. Please reserve it for Claim ID.');
      }
      if (!header) headerCell.setValue('Claim ID');

      const lastRow = sheet.getLastRow();
      const existing = lastRow > 1
        ? sheet.getRange(2, 13, lastRow - 1, 1)
            .createTextFinder(claimId)
            .matchEntireCell(true)
            .findAll()
        : [];

      // If a network response was lost, confirm the original write on retry.
      if (existing.length > 0) {
        if (existing.length !== rows.length) {
          throw new Error('Claim ID already exists with a different item count.');
        }
        return jsonResponse({
          success: true,
          claimId: claimId,
          savedTo: category,
          totalItems: existing.length,
          alreadySaved: true
        });
      }

      sheet.getRange(lastRow + 1, 1, rows.length, 13).setValues(rows);
      SpreadsheetApp.flush();

      return jsonResponse({
        success: true,
        claimId: claimId,
        savedTo: category,
        totalItems: rows.length
      });
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    return jsonResponse({ success: false, message: error.message });
  }
}

function jsonResponse(result) {
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}
