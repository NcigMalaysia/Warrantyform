// Replace the existing doPost and jsonResponse functions in the warranty
// Apps Script project, then deploy a new version of its existing web app.
// Column M of the User and Seller tabs is reserved for Claim ID.
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const category = String(data.category || '').trim();
    const claimId = String(data.claimId || '').trim();
    const warrantyDate = Utilities.formatDate(
      new Date(),
      'Asia/Kuala_Lumpur',
      'yyyy-MM-dd'
    );
    const customerName = toSmartTitleCase_(data.name);
    const shippingAddress = toSmartTitleCase_(data.address);

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
        warrantyDate,          // A: Date (current Malaysia date)
        customerName,          // B: Full Name
        data.phone || '',      // C: Phone Number
        data.email || '',      // D: Email
        item.product || '',    // E: Product
        item.variant || '',    // F: Variant
        item.batch || '',      // G: Batch / Serial Number
        item.defect || '',     // H: Defect
        shippingAddress,       // I: Address
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
      if (!header && sheet.getLastRow() > 1) {
        const existingValues = sheet.getRange(2, 13, sheet.getLastRow() - 1, 1).getValues();
        if (existingValues.some(function (row) { return String(row[0] || '').trim() !== ''; })) {
          throw new Error('Column M contains existing data. Please reserve it for Claim ID.');
        }
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

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || '').trim();

    if (action && action !== 'dashboard') {
      throw new Error('Invalid action.');
    }

    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheets = ['User', 'Seller'];
    const rows = [];

    sourceSheets.forEach(function (sheetName) {
      const sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) return;

      const values = sheet.getDataRange().getValues();
      const headers = values[0].map(normalizeHeader_);

      // New submissions always use A, E, F, G and H. Header aliases keep the
      // dashboard compatible with older sheets whose labels were renamed.
      const dateIndex = resolveColumnIndex_(headers, ['date', 'timestamp'], 0);
      const productIndex = resolveColumnIndex_(headers, ['product', 'produk'], 4);
      const variantIndex = resolveColumnIndex_(headers, ['variant', 'varian'], 5);
      const batchIndex = resolveColumnIndex_(headers, [
        'batch number',
        'batch / serial number',
        'batch serial number',
        'batch no',
        'batch',
        'serial number'
      ], 6);
      const defectIndex = resolveColumnIndex_(headers, [
        'defect',
        'defect type',
        'jenis defect',
        'issue'
      ], 7);

      values.slice(1).forEach(function (row) {
        const product = String(row[productIndex] || '').trim();
        const defect = String(row[defectIndex] || '').trim();

        if (!product || !defect) return;

        rows.push({
          Date: formatDashboardDate_(row[dateIndex]),
          Product: product,
          Variant: String(row[variantIndex] || '').trim(),
          'Batch Number': String(row[batchIndex] || '').trim(),
          Defect: defect,
          Category: sheetName
        });
      });
    });

    return jsonResponse({
      success: true,
      rows: rows,
      totalRecords: rows.length,
      generatedAt: new Date().toISOString()
    });
  } catch (error) {
    return jsonResponse({ success: false, message: error.message });
  }
}

function normalizeHeader_(header) {
  return String(header || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function resolveColumnIndex_(headers, candidates, fallbackIndex) {
  const normalizedCandidates = candidates.map(normalizeHeader_);

  for (let i = 0; i < normalizedCandidates.length; i++) {
    const index = headers.indexOf(normalizedCandidates[i]);
    if (index !== -1) return index;
  }

  return fallbackIndex;
}

function formatDashboardDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, 'Asia/Kuala_Lumpur', 'yyyy-MM-dd');
  }
  return String(value || '').trim();
}

function jsonResponse(result) {
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function toSmartTitleCase_(value) {
  return String(value || '')
    .trim()
    .split(/(\s+)/)
    .map(function (part) {
      if (/^\s+$/.test(part) || !part) return part;
      if (/^[A-Z0-9./-]+$/.test(part) && /[A-Z]{2}/.test(part)) return part;
      return part
        .toLowerCase()
        .replace(/(^|[-'’/])([a-z])/g, function (_, boundary, letter) {
          return boundary + letter.toUpperCase();
        });
    })
    .join('');
}
