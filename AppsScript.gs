/**
 * Apps Script V3 - Product Quality Register
 *
 * Sheets:
 * - PQ_Issues_Master
 * - PQ_Issue_Images
 * - PQ_Product_Master
 *
 * Features:
 * - Save issue
 * - Lookup item by code
 * - List issues for tracker page
 * - Update follow-up fields from tracker
 */

const SHEET_ISSUES = 'PQ_Issues_Master';
const SHEET_IMAGES = 'PQ_Issue_Images';
const SHEET_PRODUCTS = 'PQ_Product_Master';
const IMAGE_FOLDER_NAME = 'PQ_Product_Quality_Issue_Images';

// =========================
// POST
// =========================
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents || '{}');
    const action = String(payload.action || '').trim();

    if (action === 'updateIssue') {
      return updateIssue_(payload);
    }

    return saveIssue_(payload);

  } catch (error) {
    return jsonOutput_({
      status: 'error',
      message: error.message || String(error)
    });
  }
}

// =========================
// GET
// =========================
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? String(e.parameter.action).trim() : '';

    if (action === 'lookupItem') {
      const itemCode = (e.parameter.itemCode || '').trim();
      return lookupItemResponse_(itemCode);
    }

    if (action === 'listIssues') {
      return listIssues_();
    }

    if (action === 'health') {
      return jsonOutput_({
        status: 'success',
        message: 'Backend is live',
        sheets: [SHEET_ISSUES, SHEET_IMAGES, SHEET_PRODUCTS]
      });
    }

    return jsonOutput_({
      status: 'success',
      message: 'Apps Script backend is working'
    });

  } catch (error) {
    return jsonOutput_({
      status: 'error',
      message: error.message || String(error)
    });
  }
}

// =========================
// Save issue
// =========================
function saveIssue_(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const issuesSheet = ss.getSheetByName(SHEET_ISSUES);
  const imagesSheet = ss.getSheetByName(SHEET_IMAGES);

  if (!issuesSheet) throw new Error('Sheet not found: ' + SHEET_ISSUES);
  if (!imagesSheet) throw new Error('Sheet not found: ' + SHEET_IMAGES);

  const issueId = generateIssueId_(issuesSheet);
  const now = new Date();
  const timestamp = formatDateTime_(now);

  const files = Array.isArray(payload.files) ? payload.files : [];
  const uploadedImages = [];

  if (files.length > 0) {
    const folder = getOrCreateFolder_();

    files.forEach(function(file, index) {
      if (!file || !file.base64) return;

      const bytes = Utilities.base64Decode(file.base64);
      const blob = Utilities.newBlob(
        bytes,
        file.mimeType || 'application/octet-stream',
        file.name || ('image_' + (index + 1))
      );

      const createdFile = folder.createFile(blob);
      const fileUrl = createdFile.getUrl();
      const fileName = createdFile.getName();
      const imageRecordId = issueId + '-IMG-' + String(index + 1).padStart(2, '0');

      imagesSheet.appendRow([
        imageRecordId,                       // A Image_Record_ID
        issueId,                            // B Issue_ID
        fileName,                           // C File_Name
        fileUrl,                            // D File_URL
        timestamp,                          // E Upload_Date
        payload.Reporter_Name || '',        // F Uploaded_By
        'Uploaded from HTML form'           // G Source_Note
      ]);

      uploadedImages.push({
        name: fileName,
        url: fileUrl
      });
    });
  }

  const imageCount = uploadedImages.length;
  const imageLinks = uploadedImages.map(function(img) {
    return img.url;
  }).join(' | ');

  issuesSheet.appendRow([
    issueId,                                         // A Issue_ID
    payload.Date_Observed || '',                     // B Date_Observed
    payload.Source_Channel || '',                    // C Source_Channel
    payload.Entity_Name || '',                       // D Entity_Name
    payload.Reference_No || '',                      // E Reference_No
    payload.Item_Code || '',                         // F Item_Code
    payload.Item_Name || '',                         // G Item_Name
    payload.Product_Category || '',                  // H Product_Category
    payload.Problem_Description || '',               // I Problem_Description
    payload.Cause_Description || '',                 // J Cause_Description
    payload.Proposed_Action || '',                   // K Proposed_Action
    payload.Actual_Action_Taken || '',               // L Actual_Action_Taken
    payload.Approved_Final_Solution || '',           // M Approved_Final_Solution
    payload.Implementation_Effective_From || '',     // N Implementation_Effective_From
    payload.Implementation_Notes || '',              // O Implementation_Notes
    payload.Concerned_Departments || '',             // P Concerned_Departments
    payload.Severity_Level || '',                    // Q Severity_Level
    payload.Followup_Status || '',                   // R Followup_Status
    payload.Followup_Notes || '',                    // S Followup_Notes
    payload.Reporter_Name || '',                     // T Reporter_Name
    imageCount,                                      // U Image_Count
    imageLinks,                                      // V Image_Links
    payload.Action_Update_Date || '',                // W Action_Update_Date
    timestamp,                                       // X Created_At
    timestamp                                        // Y Updated_At
  ]);

  return jsonOutput_({
    status: 'success',
    issueId: issueId,
    imageCount: imageCount,
    imageLinks: imageLinks,
    createdAt: timestamp,
    updatedAt: timestamp,
    message: 'Issue saved successfully'
  });
}

// =========================
// List issues for tracker
// =========================
function listIssues_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ISSUES);

  if (!sheet) throw new Error('Sheet not found: ' + SHEET_ISSUES);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) {
    return jsonOutput_({
      status: 'success',
      count: 0,
      issues: []
    });
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const issues = values.map(function(row, index) {
    const obj = {};
    headers.forEach(function(header, colIndex) {
      obj[String(header)] = row[colIndex];
    });
    obj._rowNumber = index + 2;
    return obj;
  }).reverse();

  return jsonOutput_({
    status: 'success',
    count: issues.length,
    issues: issues
  });
}

// =========================
// Update selected issue
// =========================
function updateIssue_(payload) {
  const issueId = String(payload.Issue_ID || '').trim();
  if (!issueId) {
    return jsonOutput_({
      status: 'error',
      message: 'Missing Issue_ID'
    });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_ISSUES);
  if (!sheet) throw new Error('Sheet not found: ' + SHEET_ISSUES);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return jsonOutput_({
      status: 'error',
      message: 'No issue rows found'
    });
  }

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  let targetRow = null;
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0] || '').trim() === issueId) {
      targetRow = i + 2;
      break;
    }
  }

  if (!targetRow) {
    return jsonOutput_({
      status: 'error',
      message: 'Issue_ID not found: ' + issueId
    });
  }

  const now = formatDateTime_(new Date());

  // Columns in PQ_Issues_Master:
  // L=12 Actual_Action_Taken
  // M=13 Approved_Final_Solution
  // N=14 Implementation_Effective_From
  // R=18 Followup_Status
  // S=19 Followup_Notes
  // W=23 Action_Update_Date
  // Y=25 Updated_At

  if (payload.Actual_Action_Taken !== undefined) {
    sheet.getRange(targetRow, 12).setValue(payload.Actual_Action_Taken || '');
  }

  if (payload.Approved_Final_Solution !== undefined) {
    sheet.getRange(targetRow, 13).setValue(payload.Approved_Final_Solution || '');
  }

  if (payload.Implementation_Effective_From !== undefined) {
    sheet.getRange(targetRow, 14).setValue(payload.Implementation_Effective_From || '');
  }

  if (payload.Followup_Status !== undefined) {
    sheet.getRange(targetRow, 18).setValue(payload.Followup_Status || '');
  }

  if (payload.Followup_Notes !== undefined) {
    sheet.getRange(targetRow, 19).setValue(payload.Followup_Notes || '');
  }

  const actionUpdateDate = payload.Action_Update_Date || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  sheet.getRange(targetRow, 23).setValue(actionUpdateDate);
  sheet.getRange(targetRow, 25).setValue(now);

  return jsonOutput_({
    status: 'success',
    message: 'Issue updated successfully',
    Issue_ID: issueId,
    Updated_At: now
  });
}

// =========================
// Lookup item from PQ_Product_Master
// A = Item_Code
// B = Item_Name
// C = Product_Category
// =========================
function lookupItemResponse_(itemCode) {
  if (!itemCode) {
    return jsonOutput_({
      status: 'error',
      message: 'Missing itemCode'
    });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productSheet = ss.getSheetByName(SHEET_PRODUCTS);

  if (!productSheet) {
    throw new Error('Sheet not found: ' + SHEET_PRODUCTS);
  }

  const lastRow = productSheet.getLastRow();
  if (lastRow < 2) {
    return jsonOutput_({
      status: 'not_found',
      itemCode: itemCode,
      message: 'No product rows found'
    });
  }

  const values = productSheet.getRange(2, 1, lastRow - 1, 3).getValues();

  for (var i = 0; i < values.length; i++) {
    var rowCode = String(values[i][0] || '').trim();
    if (rowCode === itemCode) {
      return jsonOutput_({
        status: 'success',
        itemCode: itemCode,
        itemName: String(values[i][1] || '').trim(),
        productCategory: String(values[i][2] || '').trim()
      });
    }
  }

  return jsonOutput_({
    status: 'not_found',
    itemCode: itemCode,
    message: 'Item code not found'
  });
}

// =========================
// Generate Issue ID
// =========================
function generateIssueId_(issuesSheet) {
  const currentYear = new Date().getFullYear();
  const lastRow = issuesSheet.getLastRow();

  if (lastRow < 2) {
    return 'PQ-' + currentYear + '-0001';
  }

  const lastValue = issuesSheet.getRange(lastRow, 1).getValue();
  const match = String(lastValue).match(/PQ-(\d{4})-(\d+)/);

  if (!match) {
    return 'PQ-' + currentYear + '-0001';
  }

  const yearInSheet = Number(match[1]);
  const sequence = Number(match[2]);

  if (yearInSheet !== currentYear) {
    return 'PQ-' + currentYear + '-0001';
  }

  const nextSequence = String(sequence + 1).padStart(4, '0');
  return 'PQ-' + currentYear + '-' + nextSequence;
}

// =========================
// Drive folder for images
// =========================
function getOrCreateFolder_() {
  const folders = DriveApp.getFoldersByName(IMAGE_FOLDER_NAME);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(IMAGE_FOLDER_NAME);
}

// =========================
// Date time formatter
// =========================
function formatDateTime_(dateObj) {
  return Utilities.formatDate(
    dateObj,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH:mm:ss'
  );
}

// =========================
// JSON output
// =========================
function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
