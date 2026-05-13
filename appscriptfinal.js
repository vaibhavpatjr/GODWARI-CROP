// Google Apps Script - Modified to work with existing columns (Brand = Title of BioStimulant, Category = Brand Name)
function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var rows = data.slice(1);
  
  // Check if single product request
  if (e && e.parameter && e.parameter.product) {
    var productId = e.parameter.product;
    for (var i = 0; i < rows.length; i++) {
      var productIndex = headers.indexOf('PRODUCT_ID');
      if (productIndex !== -1 && rows[i][productIndex] === productId) {
        var product = {};
        for (var j = 0; j < headers.length; j++) {
          var headerName = headers[j];
          var value = rows[i][j];
          // Map existing columns to expected field names
          if (headerName === 'Brand') {
            product['Title of BioStimulant'] = value;
            product[headerName] = value;
          } else if (headerName === 'Category') {
            product['Brand Name'] = value;
            product[headerName] = value;
          } else {
            product[headerName] = value;
          }
        }
        return ContentService.createTextOutput(JSON.stringify(product))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({error: "Product not found"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // Return all data - map column names for frontend
  var mappedData = [['PRODUCT_ID', 'Product Name', 'Title of BioStimulant', 'Brand Name', 'Company', 'Batch Number', 'Expiry Date', 'Composition', 'Specification', 'Dose', 'Crop', 'Usage Instructions', 'Storage Instructions', 'Warning Notes', 'Status', 'Timestamp']];
  
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    var mappedRow = [];
    for (var j = 0; j < headers.length; j++) {
      var header = headers[j];
      var value = row[j];
      if (header === 'Brand') {
        mappedRow[2] = value; // Title of BioStimulant
      } else if (header === 'Category') {
        mappedRow[3] = value; // Brand Name
      } else if (header === 'Product Name') {
        mappedRow[1] = value;
      } else if (header === 'PRODUCT_ID') {
        mappedRow[0] = value;
      } else if (header === 'Company') {
        mappedRow[4] = value;
      } else if (header === 'Batch Number') {
        mappedRow[5] = value;
      } else if (header === 'Expiry Date') {
        mappedRow[6] = value;
      } else if (header === 'Composition') {
        mappedRow[7] = value;
      } else if (header === 'Specification') {
        mappedRow[8] = value;
      } else if (header === 'Dose') {
        mappedRow[9] = value;
      } else if (header === 'Crop') {
        mappedRow[10] = value;
      } else if (header === 'Usage Instructions') {
        mappedRow[11] = value;
      } else if (header === 'Storage Instructions') {
        mappedRow[12] = value;
      } else if (header === 'Warning Notes') {
        mappedRow[13] = value;
      } else if (header === 'Status') {
        mappedRow[14] = value;
      }
    }
    // Fill missing indices
    for (var k = 0; k < mappedRow.length; k++) {
      if (mappedRow[k] === undefined) mappedRow[k] = '';
    }
    mappedData.push(mappedRow);
  }
  
  var result = {
    success: true,
    data: mappedData
  };
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  
  // Get parameters
  var action = e.parameter.action;
  var productJson = e.parameter.product;
  
  if (!productJson) {
    return ContentService.createTextOutput(JSON.stringify({success: false, error: "No product data"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  var product;
  try {
    product = JSON.parse(productJson);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({success: false, error: "Invalid JSON: " + err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // Handle delete
  if (action === 'delete') {
    var productIdToDelete = e.parameter.productId || product.PRODUCT_ID;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === productIdToDelete) {
        sheet.deleteRow(i + 1);
        return ContentService.createTextOutput(JSON.stringify({success: true, message: "Deleted"}))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({success: false, message: "Product not found"}))
      .setMimeType(ContentService.MimeType.JSON);
  }
  
  // Handle update or add
  var productId = product.PRODUCT_ID;
  var foundRow = -1;
  
  // Find existing row by PRODUCT_ID (assuming it's in column A)
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === productId) {
      foundRow = i + 1;
      break;
    }
  }
  
  // Get existing headers or create if empty
  if (headers.length === 0 || headers[0] !== 'PRODUCT_ID') {
    headers = ['PRODUCT_ID', 'Product Name', 'Brand', 'Company', 'Category', 'Batch Number', 'Manufacture Date', 'Expiry Date', 'Composition', 'Specification', 'Dose', 'Crop', 'Usage Instructions', 'Storage Instructions', 'Warning Notes', 'Status', 'Timestamp'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  
  // Build row data based on existing sheet headers
  var rowData = [];
  for (var h = 0; h < headers.length; h++) {
    var headerName = headers[h];
    var value = '';
    
    // Map frontend fields to sheet columns
    if (headerName === 'PRODUCT_ID') {
      value = product.PRODUCT_ID || '';
    } else if (headerName === 'Product Name') {
      value = product['Product Name'] || '';
    } else if (headerName === 'Brand') {
      // Brand column stores Title of BioStimulant
      value = product['Title of BioStimulant'] || product.Brand || '';
    } else if (headerName === 'Company') {
      value = product.Company || '';
    } else if (headerName === 'Category') {
      // Category column stores Brand Name
      value = product['Brand Name'] || product.Category || '';
    } else if (headerName === 'Batch Number') {
      value = product['Batch Number'] || '';
    } else if (headerName === 'Manufacture Date') {
      value = product['Manufacture Date'] || '';
    } else if (headerName === 'Expiry Date') {
      value = product['Expiry Date'] || '';
    } else if (headerName === 'Composition') {
      value = typeof product.Composition === 'object' ? JSON.stringify(product.Composition) : (product.Composition || '');
    } else if (headerName === 'Specification') {
      value = typeof product.Specification === 'object' ? JSON.stringify(product.Specification) : (product.Specification || '');
    } else if (headerName === 'Dose') {
      value = product.Dose || '';
    } else if (headerName === 'Crop') {
      value = product.Crop || '';
    } else if (headerName === 'Usage Instructions') {
      value = product['Usage Instructions'] || '';
    } else if (headerName === 'Storage Instructions') {
      value = product['Storage Instructions'] || '';
    } else if (headerName === 'Warning Notes') {
      value = product['Warning Notes'] || '';
    } else if (headerName === 'Status') {
      value = product.Status || 'Active';
    } else if (headerName === 'Timestamp') {
      value = product.Timestamp || new Date().toISOString();
    }
    
    rowData.push(value);
  }
  
  // Update or insert
  if (foundRow > 0) {
    // UPDATE existing row
    for (var col = 0; col < rowData.length; col++) {
      if (rowData[col] !== undefined && rowData[col] !== '') {
        sheet.getRange(foundRow, col + 1).setValue(rowData[col]);
      }
    }
    var message = "Product updated successfully";
  } else {
    // INSERT new row
    sheet.appendRow(rowData);
    var message = "Product added successfully";
  }
  
  return ContentService.createTextOutput(JSON.stringify({
    success: true, 
    message: message,
    productId: productId
  })).setMimeType(ContentService.MimeType.JSON);
}
