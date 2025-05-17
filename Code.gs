function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Forms')
    .addItem('Upload Excel Data', 'showUploadForm')
    .addItem('Initialize Report Sheets', 'initializeAllReportSheets')
    .addToUi();
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('UploadForm')
    .setTitle('Report Submission Form')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function authenticateUser(employeeId, password) {
  try {
    const ss = SpreadsheetApp.openById('15Q_EMBhht_yw6BbY5Kx9BV4ictn2zKqoTiKuvFilxTo');
    const userSheet = ss.getSheetByName('User_Data');
    if (!userSheet) throw new Error("User database not found");

    const [headers, ...data] = userSheet.getDataRange().getValues();
    
    // Find column indexes in User_Data sheet
    const idCol = headers.findIndex(h => h.toString().trim().toLowerCase() === 'id');
    const nameCol = headers.findIndex(h => h.toString().trim().toLowerCase() === 'name');
    const zoneCol = headers.findIndex(h => h.toString().trim().toLowerCase() === 'zone');
    const passCol = headers.findIndex(h => h.toString().trim().toLowerCase() === 'password');
    const emailCol = headers.findIndex(h => h.toString().trim().toLowerCase() === 'email');

    if (idCol === -1 || nameCol === -1 || zoneCol === -1 || passCol === -1) {
      throw new Error("Required columns not found in User_Data");
    }

    // Hash password and find user
    const hashedPassword = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      password
    ).map(b => ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2)).join('');

    const user = data.find(row => 
      row[idCol]?.toString().trim().toLowerCase() === employeeId.toLowerCase().trim() && 
      row[passCol]?.toString().trim() === hashedPassword
    );

    if (!user) throw new Error("Invalid Employee ID or password");

    // Prepare user data object
    const userData = {
      id: user[idCol],       // From Column A in User_Data
      name: user[nameCol],   // From Column C in User_Data
      zone: user[zoneCol],   // From Column D in User_Data
      email: emailCol >= 0 ? user[emailCol] : '' // Include email if available
    };

    // Cache the user data for 30 minutes (1800 seconds)
    const cache = CacheService.getScriptCache();
    cache.put('currentUser', JSON.stringify(userData), 1800);

    return {
      success: true,
      message: "Authentication successful",
      userData: userData
    };
  } catch (e) {
    Logger.log('Authentication error: ' + e.message);
    return {
      success: false,
      message: e.message,
      error: e.stack
    };
  }
}

function hashPassword(password) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password)
    .map(b => ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2))
    .join('');
}

function getFieldsForReportType(reportType) {
  var ss = SpreadsheetApp.openById('15Q_EMBhht_yw6BbY5Kx9BV4ictn2zKqoTiKuvFilxTo');
  var dataSheet = ss.getSheetByName('Data-Report');
  var dropDataSheet = ss.getSheetByName('Data_Drop');
  
  var data = dataSheet.getRange(2, 1, dataSheet.getLastRow()-1, 3).getValues();
  var dropOptions = dropDataSheet.getRange(2, 1, dropDataSheet.getLastRow()-1, 2).getValues();
  var optionsMap = {};
  
  dropOptions.forEach(function(row) {
    var fieldLabel = row[0];
    var optionValue = row[1];
    if (!optionsMap[fieldLabel]) optionsMap[fieldLabel] = [];
    if (optionValue) optionsMap[fieldLabel].push(optionValue);
  });
  
  return data
    .filter(row => row[0] === reportType)
    .map(field => {
      var fieldLabel = field[1];
      var fieldType = field[2];
      var fieldName = fieldLabel.replace(/\s+/g, '_').toLowerCase();
      var fieldData = {
        label: fieldLabel,
        type: fieldType,
        name: fieldName
      };
      
      if (optionsMap[fieldLabel]?.length > 0) {
        fieldData.type = 'Dropdown';
        fieldData.options = optionsMap[fieldLabel];
      }
      return fieldData;
    });
}

function submitReport(formData) {
  var ss = SpreadsheetApp.openById('15Q_EMBhht_yw6BbY5Kx9BV4ictn2zKqoTiKuvFilxTo');

  try {
    var userData = formData.userData || { 
      id: '', 
      name: 'Unknown User', 
      zone: 'Unknown Zone' 
    };
    
    var sheetName = 'Sub_' + formData.reportType.replace(/\s+/g, '_').substring(0, 25).replace(/[\/\\?\*\[\]]/g, '');
    var sheet = ss.getSheetByName(sheetName) || createReportSheet(ss, sheetName, formData.reportType);

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var rowData = headers.map((h, i) => {
      if (i === 0) return new Date();                   // Column A: Timestamp
      if (i === 1) return formData.reportType;          // Column B: Report Type
      if (i === 2) return userData.id;                  // Column C: User ID (from User_Data Column A)
      if (i === 3) return userData.name;                // Column D: Name (from User_Data Column C)
      if (i === 4) return userData.zone;                // Column E: Zone (from User_Data Column D)
      
      // Handle report-specific fields
      var fieldName = h.replace(/\s+/g, '_').toLowerCase();
      return formData.fields[fieldName] || '';
    });

    sheet.appendRow(rowData);
    sendEmailNotification(userData, formData);
    return { success: true, message: 'Report submitted successfully!' };
  } catch (e) {
    Logger.log('Error in submitReport: ' + e.toString());
    return { success: false, message: 'Error submitting report: ' + e.message };
  }
}

function createReportSheet(ss, sheetName, reportType) {
  var sheet = ss.insertSheet(sheetName);
  var dataSheet = ss.getSheetByName('Data-Report');
  var data = dataSheet.getRange(2, 1, dataSheet.getLastRow()-1, 3).getValues();
  var reportFields = data.filter(row => row[0] === reportType).map(row => row[1]);
  
  // Fixed column structure
  var headers = [
    'Timestamp',     // Column A
    'Report Type',   // Column B
    'User ID',       // Column C (from User_Data Column A)
    'Name',          // Column D (from User_Data Column C)
    'Zone'           // Column E (from User_Data Column D)
  ].concat(reportFields);
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Format header row
  sheet.getRange(1, 1, 1, headers.length)
    .setBackground('#eeeeee')
    .setFontWeight('bold')
    .setBorder(true, true, true, true, true, true);
    
  // Auto-resize columns
  for (var i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  return sheet;
}

function sendEmailNotification(userData, formData) {
  var emailBody = `New report submission:\n\n
    Submitted By: ${userData.name}\n
    Zone: ${userData.zone}\n
    Email: ${userData.email || 'Not provided'}\n
    Type: ${formData.reportType}\n
    Time: ${new Date()}\n\nDetails:\n` +
  Object.entries(formData.fields).map(([key, value]) => `${key}: ${value}`).join('\n');

  try {
    MailApp.sendEmail({
      to: 'nflvigilance@gmail.com',
      subject: 'New Report: ' + formData.reportType,
      body: emailBody
    });
    Logger.log('Email notification sent successfully');
  } catch (e) {
    Logger.log('Error sending email: ' + e.toString());
  }
}

function initializeAllReportSheets() {
  var ss = SpreadsheetApp.openById('15Q_EMBhht_yw6BbY5Kx9BV4ictn2zKqoTiKuvFilxTo');
  var dataSheet = ss.getSheetByName('Data-Report');
  
  var reportTypes = dataSheet.getRange(2, 1, dataSheet.getLastRow()-1, 1).getValues()
    .flat()
    .filter(function(item, pos, self) {
      return item && self.indexOf(item) == pos;
    });
  
  var data = dataSheet.getRange(2, 1, dataSheet.getLastRow()-1, 3).getValues();
  
  reportTypes.forEach(function(reportType) {
    var reportFields = data.filter(function(row) {
      return row[0] === reportType;
    }).map(function(row) { return row[1]; });
    
    // New header structure with only ID, Name, Zone
    var headers = ['Timestamp', 'Report Type', 'ID', 'Name', 'Zone'].concat(reportFields);
    
    var sheetName = 'Sub_' + reportType.replace(/\s+/g, '_').substring(0, 25);
    sheetName = sheetName.replace(/[\/\\?*\[\]]/g, '');
    
    var sheet = ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
    
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } 
    else {
      var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      // Check if headers match (ignoring case and whitespace)
      var normalizedCurrent = currentHeaders.map(h => h.toString().trim().toLowerCase());
      var normalizedNew = headers.map(h => h.toString().trim().toLowerCase());
      
      var headersMatch = JSON.stringify(normalizedCurrent) === JSON.stringify(normalizedNew);
      
      if (!headersMatch) {
        var existingData = sheet.getDataRange().getValues();
        
        ss.deleteSheet(sheet);
        sheet = ss.insertSheet(sheetName);
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        
        if (existingData.length > 1) {
          var newData = [];
          
          for (var r = 1; r < existingData.length; r++) {
            var newRow = new Array(headers.length).fill('');
            
            // Map old data to new structure
            newRow[0] = existingData[r][0]; // Timestamp
            newRow[1] = existingData[r][1]; // Report Type
            
            // For existing data, we can't determine ID, so leave blank
            // Name and Zone will be in their original positions (assuming old structure was Timestamp, Report Type, Name, Zone)
            newRow[3] = existingData[r][2] || ''; // Name (from old column C)
            newRow[4] = existingData[r][3] || ''; // Zone (from old column D)
            
            // Map all other fields
            for (var c = 4; c < existingData[r].length; c++) {
              if (c >= existingData[0].length) continue;
              var oldHeader = existingData[0][c];
              var newIndex = headers.indexOf(oldHeader);
              if (newIndex !== -1) {
                newRow[newIndex] = existingData[r][c];
              }
            }
            
            newData.push(newRow);
          }
          
          if (newData.length > 0) {
            sheet.getRange(2, 1, newData.length, headers.length).setValues(newData);
          }
        }
      }
    }
    
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#eeeeee')
      .setFontWeight('bold')
      .setBorder(true, true, true, true, true, true);
  });
  
  return "Successfully initialized " + reportTypes.length + " report sheets (data preserved)";
}

function getUserData_() {
  try {
    if (!checkAuthorization()) {
      return {success: false, message: "Not authorized"};
    }

    var userEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
    var userSheet = SpreadsheetApp.openById('1dP4XXlbna4OghQaUgt8o2SKmgUBAHb7atIZx35AJ0zw');
    var data = userSheet.getDataRange().getValues();
    
    var headers = data[0].map(function(h) { return h.toString().toLowerCase(); });
    var emailCol = headers.findIndex(function(h) { return h.includes('email'); });
    var nameCol = headers.findIndex(function(h) { return h.includes('name'); });
    var zoneCol = headers.findIndex(function(h) { return h.includes('zone'); });
    var idCol = headers.findIndex(function(h) { return h.includes('id'); });

    if (emailCol === -1) throw new Error("Email column not found");
    
    for (var i = 1; i < data.length; i++) {
      var rowEmail = data[i][emailCol] ? data[i][emailCol].toString().toLowerCase().trim() : '';
      if (rowEmail === userEmail) {
        return {
          success: true,
          email: data[i][emailCol],
          name: nameCol >= 0 ? data[i][nameCol] : 'Unknown',
          zone: zoneCol >= 0 ? data[i][zoneCol] : 'Unknown',
          userId: idCol >= 0 ? data[i][idCol] : 'N/A'
        };
      }
    }
    return {success: false, message: "User not found in database"};
  } catch (e) {
    Logger.log("Error in getUserData_: " + e.message);
    return {success: false, message: "System error: " + e.message};
  }
}

function checkAuthorization() {
  try {
    var email = Session.getActiveUser().getEmail();
    if (!email) throw new Error("No active user session");
    return true;
  } catch (e) {
    Logger.log("Authorization check failed: " + e.message);
    return false;
  }
}

function hashAllPasswords() {
  var ss = SpreadsheetApp.openById('15Q_EMBhht_yw6BbY5Kx9BV4ictn2zKqoTiKuvFilxTo');
  var userSheet = ss.getSheetByName('User_Data');
  
  var data = userSheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return h.toString().toLowerCase(); });
  
  var emailCol = headers.findIndex(function(h) { return h.includes('email') || h.includes('login'); });
  var passwordCol = headers.findIndex(function(h) { return h.includes('password'); });
  
  if (emailCol === -1 || passwordCol === -1) {
    Logger.log("Required columns not found");
    return;
  }
  
  for (var i = 1; i < data.length; i++) {
    var password = data[i][passwordCol];
    if (password && !password.startsWith('$2a$') && !password.match(/^[a-f0-9]{64}$/i)) {
      var hashed = hashPassword(password);
      userSheet.getRange(i+1, passwordCol+1).setValue(hashed);
      Utilities.sleep(200);
    }
  }
  
  SpreadsheetApp.flush();
  Logger.log("Password hashing completed");
}

function showMainFormAfterLogin() {
  var html = HtmlService.createTemplateFromFile('FormUI').evaluate();
  SpreadsheetApp.getUi().showModalDialog(html, 'Report Submission Form');
}

function getDataRange(sheetName, startRow, startCol, numCols) {
  var ss = SpreadsheetApp.openById('15Q_EMBhht_yw6BbY5Kx9BV4ictn2zKqoTiKuvFilxTo');
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    throw new Error("Sheet '" + sheetName + "' not found");
  }
  
  var lastRow = sheet.getLastRow();
  var numRows = lastRow - startRow + 1;
  
  if (numRows <= 0) {
    return [];
  }
  
  return sheet.getRange(startRow, startCol, numRows, numCols).getValues();
}

function showUploadForm() {
  var html = HtmlService.createHtmlOutputFromFile('UploadForm')
    .setWidth(600)
    .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Upload Report Data');
}

function processUploadedFile(fileData, fileName, reportType) {
  try {
    // Validate inputs
    if (!fileData || !fileName || !reportType) {
      return { success: false, message: "Missing required parameters" };
    }
    
    // Check if file is Excel
    if (!fileName.toLowerCase().endsWith('.xlsx')) {
      return { success: false, message: "Only Excel (.xlsx) files are supported" };
    }
    
    // Convert base64 to blob
    const blob = Utilities.newBlob(Utilities.base64Decode(fileData), fileName);
    
    // Parse the Excel file
    const ss = SpreadsheetApp.openById('15Q_EMBhht_yw6BbY5Kx9BV4ictn2zKqoTiKuvFilxTo');
    const resource = {
      title: 'Temp Excel Import',
      mimeType: MimeType.MICROSOFT_EXCEL,
      parents: [{id: ss.getId()}]
    };
    
    // Upload and convert to Google Sheets format
    const file = Drive.Files.insert(resource, blob, {convert: true});
    const tempSheet = SpreadsheetApp.openById(file.id);
    const tempData = tempSheet.getSheets()[0].getDataRange().getValues();
    
    // Delete the temporary file
    Drive.Files.remove(file.id);
    
    // Determine target sheet name
    const sheetName = 'Sub_' + reportType.replace(/\s+/g, '_').substring(0, 25).replace(/[\/\\?\*\[\]]/g, '');
    const targetSheet = ss.getSheetByName(sheetName);
    
    if (!targetSheet) {
      return { success: false, message: "Target sheet not found for this report type" };
    }
    
    // Get headers from target sheet
    const headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    
    // Map Excel columns to target sheet columns
    const excelHeaders = tempData[0];
    const columnMapping = headers.map(header => {
      const headerStr = header.toString().trim().toLowerCase();
      const index = excelHeaders.findIndex(h => h.toString().trim().toLowerCase() === headerStr);
      return index >= 0 ? index : -1;
    });
    
    // Prepare data for insertion (skip header row from Excel)
    const rowsToInsert = [];
    for (let i = 1; i < tempData.length; i++) {
      const row = tempData[i];
      const newRow = columnMapping.map(colIndex => {
        if (colIndex === -1) return ''; // Column not found in Excel
        return row[colIndex] !== undefined ? row[colIndex] : '';
      });
      
      // Add timestamp to first column if it's empty
      if (!newRow[0]) {
        newRow[0] = new Date();
      }
      
      // Add report type to second column if it's empty
      if (!newRow[1]) {
        newRow[1] = reportType;
      }
      
      rowsToInsert.push(newRow);
    }
    
    // Insert data into target sheet
    if (rowsToInsert.length > 0) {
      targetSheet.getRange(targetSheet.getLastRow() + 1, 1, rowsToInsert.length, headers.length)
        .setValues(rowsToInsert);
      
      return { 
        success: true, 
        message: `Successfully imported ${rowsToInsert.length} rows to ${sheetName}` 
      };
    } else {
      return { success: false, message: "No valid data rows found in the Excel file" };
    }
    
  } catch (e) {
    Logger.log('Error in processUploadedFile: ' + e.toString());
    return { 
      success: false, 
      message: 'Error processing file: ' + e.message 
    };
  }
}

// Add this to get current user data
function getCurrentUser() {
  const cache = CacheService.getScriptCache();
  const cachedUser = cache.get('currentUser');
  if (cachedUser) {
    return JSON.parse(cachedUser);
  }
  // If not in cache (shouldn't happen), return empty object
  return {};
}

function validateExcelHeaders(fileData, fileName, reportType) {
  try {
    // Basic validation
    if (!fileData || !fileName || !reportType) {
      throw new Error("Missing required parameters");
    }
    
    if (!fileName.toLowerCase().endsWith('.xlsx')) {
      throw new Error("Only Excel (.xlsx) files are supported");
    }
    
    // Get target sheet info
    const ss = SpreadsheetApp.openById('15Q_EMBhht_yw6BbY5Kx9BV4ictn2zKqoTiKuvFilxTo');
    const sheetName = 'Sub_' + reportType.replace(/\s+/g, '_').substring(0, 25).replace(/[\/\\?\*\[\]]/g, '');
    const targetSheet = ss.getSheetByName(sheetName);
    
    if (!targetSheet) {
      throw new Error("Target sheet not found for this report type");
    }
    
    // Get expected headers (skip first 5 system columns)
    const allHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    const expectedHeaders = allHeaders.slice(5); // Skip first 5 system columns
    
    // Create temp file to read headers
    const blob = Utilities.newBlob(Utilities.base64Decode(fileData), fileName);
    const resource = {
      title: 'Temp Header Validation',
      mimeType: MimeType.MICROSOFT_EXCEL,
      parents: [{id: ss.getId()}]
    };
    
    const file = Drive.Files.insert(resource, blob, {convert: true});
    const tempSheet = SpreadsheetApp.openById(file.id);
    const tempHeaders = tempSheet.getSheets()[0].getRange(1, 1, 1, tempSheet.getSheets()[0].getLastColumn()).getValues()[0];
    
    // Clean up temp file
    Drive.Files.remove(file.id);
    
    // For debugging, log the headers
    Logger.log("Expected headers (after slice): " + JSON.stringify(expectedHeaders));
    Logger.log("Found headers in Excel: " + JSON.stringify(tempHeaders));
    
    // Normalize both header sets - consistently standardize case, whitespace and trim
    const normalizedExpected = expectedHeaders.map(h => h.toString().trim().toLowerCase());
    const normalizedFound = tempHeaders.map(h => h.toString().trim().toLowerCase());
    
    Logger.log("Normalized expected: " + JSON.stringify(normalizedExpected));
    Logger.log("Normalized found: " + JSON.stringify(normalizedFound));
    
    // More lenient matching - as long as each expected header has a counterpart in found headers
    // This approach focuses on ensuring all necessary columns are present, rather than requiring exact matches
    let missingHeaders = [];
    let extraHeaders = [];
    
    // Check for missing headers - these are required
    normalizedExpected.forEach((expectedHeader, index) => {
      if (!normalizedFound.some(foundHeader => foundHeader === expectedHeader)) {
        missingHeaders.push(expectedHeaders[index]);
      }
    });
    
    // Check for extra headers - these are acceptable but we'll report them
    normalizedFound.forEach((foundHeader, index) => {
      if (!normalizedExpected.some(expectedHeader => expectedHeader === foundHeader)) {
        extraHeaders.push(tempHeaders[index]);
      }
    });
    
    // If there are no missing headers, consider it valid - extra headers are okay
    const isValid = missingHeaders.length === 0;
    
    return {
      isValid: isValid,
      targetSheet: sheetName,
      expectedHeaders: expectedHeaders,
      foundHeaders: tempHeaders,
      missingHeaders: missingHeaders,
      extraHeaders: extraHeaders,
      message: isValid ? "Headers match" : "Headers don't match"
    };
    
  } catch (e) {
    Logger.log('Error in validateExcelHeaders: ' + e.toString());
    return {
      isValid: false,
      message: 'Error validating headers: ' + e.message
    };
  }
}

// Helper function to compare arrays
function arraysEqual(a, b) {
  if (a.length !== b.length) return false;
  const normalizedA = a.map(item => item.toString().trim().toLowerCase());
  const normalizedB = b.map(item => item.toString().trim().toLowerCase());
  
  for (let i = 0; i < normalizedA.length; i++) {
    if (normalizedA[i] !== normalizedB[i]) return false;
  }
  return true;
}

function processUploadedFile(fileData, fileName, reportType) {
  try {
    // Validate inputs
    if (!fileData || !fileName || !reportType) {
      return { success: false, message: "Missing required parameters" };
    }
    
    if (!fileName.toLowerCase().endsWith('.xlsx')) {
      return { success: false, message: "Only Excel (.xlsx) files are supported" };
    }
    
    // Get current user from cache
    const cache = CacheService.getScriptCache();
    const currentUser = JSON.parse(cache.get('currentUser'));
    
    if (!currentUser || !currentUser.id) {
      return { success: false, message: "User session expired. Please login again." };
    }
    
    // Get target sheet
    const ss = SpreadsheetApp.openById('15Q_EMBhht_yw6BbY5Kx9BV4ictn2zKqoTiKuvFilxTo');
    const sheetName = 'Sub_' + reportType.replace(/\s+/g, '_').substring(0, 25).replace(/[\/\\?\*\[\]]/g, '');
    const targetSheet = ss.getSheetByName(sheetName);
    
    if (!targetSheet) {
      return { success: false, message: "Target sheet not found for this report type" };
    }
    
    // Get all headers (system + data)
    const allHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
    const dataHeaders = allHeaders.slice(5); // Skip first 5 system columns
    
    // Parse the Excel file
    const blob = Utilities.newBlob(Utilities.base64Decode(fileData), fileName);
    const resource = {
      title: 'Temp Excel Import',
      mimeType: MimeType.MICROSOFT_EXCEL,
      parents: [{id: ss.getId()}]
    };
    
    const file = Drive.Files.insert(resource, blob, {convert: true});
    const tempSheet = SpreadsheetApp.openById(file.id);
    const tempData = tempSheet.getSheets()[0].getDataRange().getValues();
    const tempHeaders = tempData[0];
    
    // Normalize headers for comparison
    const normalizedExpected = dataHeaders.map(h => h.toString().trim().toLowerCase());
    const normalizedFound = tempHeaders.map(h => h.toString().trim().toLowerCase());
    
    // Create a mapping from source columns to target columns
    // This allows for flexibility in column order
    const columnMapping = [];
    let missingColumns = [];
    
    // For each target data column, find matching source column
    normalizedExpected.forEach((expectedHeader, targetIndex) => {
      const sourceIndex = normalizedFound.findIndex(h => h === expectedHeader);
      if (sourceIndex !== -1) {
        columnMapping.push({ 
          sourceIndex: sourceIndex,
          targetIndex: targetIndex + 5 // +5 for system columns
        });
      } else {
        missingColumns.push(dataHeaders[targetIndex]);
      }
    });
    
    // If any columns are missing, abort
    if (missingColumns.length > 0) {
      Drive.Files.remove(file.id);
      return { 
        success: false, 
        message: `Missing required columns in Excel file: ${missingColumns.join(', ')}` 
      };
    }
    
    // Prepare data for insertion
    const rowsToInsert = [];
    for (let i = 1; i < tempData.length; i++) {
      const row = tempData[i];
      
      // Create new row with all columns initialized to empty string
      const newRow = new Array(allHeaders.length).fill('');
      
      // Set system fields
      newRow[0] = new Date();             // Timestamp
      newRow[1] = reportType;             // Report Type
      newRow[2] = currentUser.id;         // ID
      newRow[3] = currentUser.name || ''; // Name
      newRow[4] = currentUser.zone || ''; // Zone
      
      // Map data columns from Excel to appropriate positions
      columnMapping.forEach(mapping => {
        newRow[mapping.targetIndex] = row[mapping.sourceIndex];
      });
      
      rowsToInsert.push(newRow);
    }
    
    // Insert data into target sheet
    if (rowsToInsert.length > 0) {
      targetSheet.getRange(targetSheet.getLastRow() + 1, 1, rowsToInsert.length, allHeaders.length)
        .setValues(rowsToInsert);
      
      // Clean up temp file
      Drive.Files.remove(file.id);
      
      return { 
        success: true, 
        message: `Successfully imported ${rowsToInsert.length} rows to ${sheetName}` 
      };
    } else {
      Drive.Files.remove(file.id);
      return { success: false, message: "No valid data rows found in the Excel file" };
    }
    
  } catch (e) {
    Logger.log('Error in processUploadedFile: ' + e.toString());
    return { 
      success: false, 
      message: 'Error processing file: ' + e.message 
    };
  }
}