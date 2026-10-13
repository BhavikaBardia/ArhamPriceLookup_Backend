/**
 * Arham Product Price Portal - Google Apps Script Backend
 * 
 * Instructions:
 * 1. Open your Google Sheet.
 * 2. Click Extensions > Apps Script.
 * 3. Paste this code into Code.gs and save.
 * 4. Click Deploy > New deployment.
 * 5. Select type: Web app.
 * 6. Set Description: "Arham Price Portal Web App".
 * 7. Set "Execute as": "Me".
 * 8. Set "Who has access": "Anyone".
 * 9. Click Deploy, copy the Web App URL, and paste it into the Web App Settings dialog!
 https://script.google.com/macros/s/AKfycbxpR4V1mzRxGovg1QXN99lUae12bFCqsGS03z-5bkEdl2K0GBujahleIKOwO3lTkDyB/exec
 */

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "getHealth";
  var result = {};
  
  try {
    if (action === "getCustomers") {
      result = getCustomers();
    } else if (action === "getProducts") {
      result = getProducts();
    } else if (action === "lookupPrices") {
      var customerId = (e.parameter.customer_id || "").trim();
      var query = (e.parameter.query || "").trim();
      result = lookupPrices(customerId, query);
    } else if (action === "createProduct") {
      var payload = {
        product_id: (e.parameter.product_id || "").trim(),
        standard_name: (e.parameter.standard_name || "").trim(),
        make: (e.parameter.make || "").trim()
      };
      result = createProduct(payload);
    } else if (action === "updateProduct") {
      var payload = {
        product_id: (e.parameter.product_id || "").trim(),
        standard_name: (e.parameter.standard_name || "").trim(),
        make: (e.parameter.make || "").trim()
      };
      result = updateProduct(payload);
    } else if (action === "deleteProduct") {
      var pid = (e.parameter.product_id || "").trim();
      result = deleteProduct(pid);
    } else if (action === "createCustomer") {
      var payload = {
        customer_id: (e.parameter.customer_id || "").trim(),
        customer_name: (e.parameter.customer_name || "").trim(),
        gstin: (e.parameter.gstin || "").trim()
      };
      result = createCustomer(payload);
    } else if (action === "updateCustomer") {
      var payload = {
        customer_id: (e.parameter.customer_id || "").trim(),
        customer_name: (e.parameter.customer_name || "").trim(),
        gstin: (e.parameter.gstin || "").trim()
      };
      result = updateCustomer(payload);
    } else if (action === "deleteCustomer") {
      var cid = (e.parameter.customer_id || "").trim();
      result = deleteCustomer(cid);
    } else if (action === "upsertPricing") {
      var payload = {
        customer_id: (e.parameter.customer_id || "").trim(),
        product_id: (e.parameter.product_id || "").trim(),
        price: parseFloat(e.parameter.price)
      };
      result = upsertPricing(payload);
    } else if (action === "getAllPricing") {
      result = getAllPricing();
    } else if (action === "getAllData") {
      result = getAllData();
    } else if (action === "saveQuotation") {
      var payload = null;
      if (e && e.parameter && e.parameter.data) {
        payload = JSON.parse(e.parameter.data);
      } else if (e && e.parameter && e.parameter.payload) {
        payload = JSON.parse(e.parameter.payload);
      } else {
        payload = e ? e.parameter : {};
      }
      result = saveQuotation(payload);
    } else if (action === "confirmOrder") {
      var payload = null;
      if (e && e.parameter && e.parameter.data) {
        payload = JSON.parse(e.parameter.data);
      } else if (e && e.parameter && e.parameter.payload) {
        payload = JSON.parse(e.parameter.payload);
      } else {
        payload = e ? e.parameter : {};
      }
      result = confirmQuotationOrder(payload);
    } else if (action === "getQuotations") {
      result = getQuotations();
    } else if (action === "getHealth") {
      result = getHealth();
    } else {
      result = { error: "Unknown action: " + action };
    }
  } catch (err) {
    result = { error: err.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var result = {};
  try {
    var data = {};
    if (e && e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    }
    
    var action = data.action || (e && e.parameter && e.parameter.action) || "";
    var payload = data.payload || data;
    
    if (action === "createProduct") {
      result = createProduct(payload);
    } else if (action === "updateProduct") {
      result = updateProduct(payload);
    } else if (action === "deleteProduct") {
      result = deleteProduct(payload.product_id || payload);
    } else if (action === "createCustomer") {
      result = createCustomer(payload);
    } else if (action === "updateCustomer") {
      result = updateCustomer(payload);
    } else if (action === "deleteCustomer") {
      result = deleteCustomer(payload.customer_id || payload);
    } else if (action === "upsertPricing") {
      result = upsertPricing(payload);
    } else if (action === "getAllPricing") {
      result = getAllPricing();
    } else if (action === "getAllData") {
      result = getAllData();
    } else if (action === "saveQuotation") {
      result = saveQuotation(payload);
    } else if (action === "confirmOrder") {
      result = confirmQuotationOrder(payload);
    } else if (action === "getQuotations") {
      result = getQuotations();
    } else if (action === "getHealth") {
      result = getHealth();
    } else {
      result = { error: "Unknown POST action: " + action };
    }
  } catch (err) {
    result = { error: err.toString() };
  }
  
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// -------------------------------------------------------------
// Helper: Get active spreadsheet instance
// -------------------------------------------------------------
function getSpreadsheet() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    return null;
  }
}

// -------------------------------------------------------------
// Action: Health Status & Setup Check
// -------------------------------------------------------------
function getHealth() {
  var ss = getSpreadsheet();
  if (!ss) {
    return {
      mode: "google_apps_script",
      status: "error",
      message: "Script must be bound to a Google Sheet",
      customer_count: 0,
      product_count: 0,
      pricing_count: 0
    };
  }
  
  ensureSheetsExist(ss);
  
  var custSheet = ss.getSheetByName("Customers");
  var prodSheet = ss.getSheetByName("Products");
  var priceSheet = ss.getSheetByName("Pricing");
  
  return {
    mode: "google_apps_script",
    status: "online",
    spreadsheet_name: ss.getName(),
    spreadsheet_id: ss.getId(),
    customer_count: Math.max(0, custSheet.getLastRow() - 1),
    product_count: Math.max(0, prodSheet.getLastRow() - 1),
    pricing_count: Math.max(0, priceSheet.getLastRow() - 1)
  };
}

// -------------------------------------------------------------
// Action: List Customers
// -------------------------------------------------------------
function getCustomers() {
  var ss = getSpreadsheet();
  if (!ss) return [];
  ensureSheetsExist(ss);
  
  var sheet = ss.getSheetByName("Customers");
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var headerRow = data[0];
  var cidIdx = 0, nameIdx = 1, gstinIdx = -1;
  
  for (var h = 0; h < headerRow.length; h++) {
    var hName = String(headerRow[h] || "").trim().toLowerCase();
    if (hName === "customer id" || hName === "customer_id" || hName === "id" || hName.indexOf("cust id") !== -1) {
      cidIdx = h;
    } else if (hName === "customer name" || hName === "customer_name" || hName === "name" || hName.indexOf("cust name") !== -1) {
      nameIdx = h;
    } else if (hName.indexOf("gst") !== -1 || hName.indexOf("tin") !== -1 || hName.indexOf("tax") !== -1) {
      gstinIdx = h;
    }
  }

  if (gstinIdx === -1) {
    for (var h = 0; h < headerRow.length; h++) {
      if (h !== cidIdx && h !== nameIdx) {
        gstinIdx = h;
        break;
      }
    }
    if (gstinIdx === -1) gstinIdx = 2;
  }
  
  var customers = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[cidIdx] && !row[nameIdx]) continue;

    var cidVal = String(row[cidIdx] || "").trim();
    var nameVal = String(row[nameIdx] || "").trim();
    var gstinVal = "";

    // Check designated gstinIdx first
    if (gstinIdx < row.length && row[gstinIdx] !== undefined && row[gstinIdx] !== null && String(row[gstinIdx]).trim() !== "") {
      gstinVal = String(row[gstinIdx]).trim();
    }
    
    // Comprehensive scan across all cells in this row if gstinVal is still empty
    if (!gstinVal) {
      for (var col = 0; col < row.length; col++) {
        if (col !== cidIdx && col !== nameIdx) {
          var val = String(row[col] || "").trim();
          if (val) {
            gstinVal = val;
            break;
          }
        }
      }
    }

    var contactVal = String(row[2] || "").trim();
    // If contact_person is identical to gstinVal (or contains GSTIN), clear contact_person to avoid duplicate display
    if (contactVal === gstinVal) contactVal = "";

    customers.push({
      customer_id: cidVal,
      customer_name: nameVal,
      contact_person: contactVal,
      phone: String(row[3] || "").trim(),
      email: String(row[4] || "").trim(),
      address: String(row[5] || "").trim(),
      gstin: gstinVal,
      GSTIN: gstinVal
    });
  }
  return customers;
}

// -------------------------------------------------------------
// Action: List Products
// -------------------------------------------------------------
function getProducts() {
  var ss = getSpreadsheet();
  if (!ss) return [];
  ensureSheetsExist(ss);
  
  var sheet = ss.getSheetByName("Products");
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var products = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0]) continue;
    products.push({
      product_id: String(row[0] || "").trim(),
      standard_name: String(row[1] || "").trim(),
      make: String(row[2] || "").trim()
    });
  }
  return products;
}

// -------------------------------------------------------------
// Action: List All Pricing Entries
// -------------------------------------------------------------
function getAllPricing() {
  var ss = getSpreadsheet();
  if (!ss) return [];
  ensureSheetsExist(ss);
  
  var sheet = ss.getSheetByName("Pricing");
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var records = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[0] || !row[1]) continue;
    var parsedPrice = parseNumericPrice(row[2]);
    records.push({
      customer_id: String(row[0] || "").trim(),
      product_id: String(row[1] || "").trim(),
      customer_price: parsedPrice,
      price: parsedPrice,
      previous_price: parseNumericPrice(row[3]),
      last_updated: String(row[4] || "").trim(),
      price_history: String(row[5] || "").trim()
    });
  }
  return records;
}

// -------------------------------------------------------------
// Action: Bulk Load All Spreadsheet Data (Customers, Products, Pricing)
// -------------------------------------------------------------
function getAllData() {
  return {
    customers: getCustomers(),
    products: getProducts(),
    pricing: getAllPricing()
  };
}

// Helper: Parse Numeric Price from sheet value safely
function parseNumericPrice(val) {
  if (val === null || val === undefined || val === "") return null;
  if (typeof val === "number") return isNaN(val) ? null : val;
  var str = String(val).replace(/[^0-9\.]/g, "");
  if (!str) return null;
  var num = parseFloat(str);
  return isNaN(num) ? null : num;
}

// -------------------------------------------------------------
// Action: Lookup Prices with Typo-Tolerant Fuzzy Logic
// -------------------------------------------------------------
function lookupPrices(customerId, query) {
  var ss = getSpreadsheet();
  if (!ss || !customerId) return [];
  ensureSheetsExist(ss);
  
  var targetStr = String(customerId || "").trim().toLowerCase();
  var customers = getCustomers();
  var custMatch = customers.find(function(c) { 
    var cid = String(c.customer_id || "").trim().toLowerCase();
    var cname = String(c.customer_name || "").trim().toLowerCase();
    return cid === targetStr || cname === targetStr || (cname && targetStr.indexOf(cname) !== -1) || (cname && cname.indexOf(targetStr) !== -1); 
  });
  
  var targetCid = custMatch ? String(custMatch.customer_id || "").trim().toLowerCase() : targetStr;
  var targetCustName = custMatch ? String(custMatch.customer_name || "").trim().toLowerCase() : targetStr;
  var actualCustName = custMatch ? custMatch.customer_name : customerId;
  var actualCustId = custMatch ? custMatch.customer_id : customerId;

  // Load Pricing Map
  var priceSheet = ss.getSheetByName("Pricing");
  var priceData = priceSheet.getDataRange().getValues();
  var priceMap = {};
  
  for (var i = 1; i < priceData.length; i++) {
    var row = priceData[i];
    if (!row[0]) continue;
    var rowCust = String(row[0]).trim().toLowerCase();
    
    // Check if row matches customer ID OR customer Name
    var isCustMatch = (rowCust === targetCid) || 
                      (targetCustName && rowCust === targetCustName) ||
                      (targetCustName && rowCust.indexOf(targetCustName) !== -1) ||
                      (targetCustName && targetCustName.indexOf(rowCust) !== -1);

    if (isCustMatch) {
      var prodKey = String(row[1] || "").trim().toLowerCase();
      var pVal = parseNumericPrice(row[2]);
      var prevVal = parseNumericPrice(row[3]);
      
      var priceEntry = {
        price: pVal,
        previous_price: prevVal,
        last_updated: String(row[4] || "").trim(),
        price_history: String(row[5] || "").trim()
      };

      // Store in map by whatever product identifier is in row[1] (Product ID or Standard Name)
      priceMap[prodKey] = priceEntry;
    }
  }
  
  var q = String(query || "").trim().toLowerCase();
  var products = getProducts();
  var results = [];
  
  if (!q) {
    for (var p = 0; p < products.length; p++) {
      var prod = products[p];
      var pid = String(prod.product_id || "").trim().toLowerCase();
      var pname = String(prod.standard_name || "").trim().toLowerCase();

      // Look up in priceMap by product ID OR standard product name
      var pr = priceMap[pid] || priceMap[pname] || null;
      var isConfigured = (pr && pr.price !== null && pr.price !== undefined);

      results.push({
        product_id: prod.product_id,
        standard_name: prod.standard_name,
        make: prod.make,
        customer_id: actualCustId,
        customer_name: actualCustName,
        price: isConfigured ? pr.price : null,
        previous_price: pr ? pr.previous_price : null,
        last_updated: pr ? pr.last_updated : "",
        price_history: pr ? pr.price_history : "",
        is_configured: isConfigured
      });
    }
    return results;
  }


  
  for (var p = 0; p < products.length; p++) {
    var prod = products[p];
    var pid = prod.product_id;
    var stdName = prod.standard_name;
    var make = prod.make;
    
    // Only include products related/configured for this customer
    var pr = priceMap[pid] || null;
    if (!pr || pr.price === null || pr.price === undefined) {
      continue;
    }

    var textTokens = (pid + " " + stdName + " " + make).replace(/[,-]/g, " ").toLowerCase().split(/\s+/).filter(Boolean);
    
    var wordScores = [];
    for (var w = 0; w < qWords.length; w++) {
      var qw = qWords[w];
      var bestScore = 0.0;
      
      for (var t = 0; t < textTokens.length; t++) {
        var pt = textTokens[t];
        var s = 0.0;
        if (pt.indexOf(qw) !== -1 || qw.indexOf(pt) !== -1) {
          s = 1.0;
        } else {
          s = getLevenshteinSimilarity(qw, pt);
          if (s >= 0.6 && qw[0] === pt[0]) {
            s += 0.05;
          }
        }
        if (s > bestScore) bestScore = s;
      }
      wordScores.push(bestScore);
    }
    
    var allMatched = wordScores.every(function(score) { return score >= 0.65; });
    if (allMatched && wordScores.length > 0) {
      var avgScore = wordScores.reduce(function(a, b) { return a + b; }, 0) / wordScores.length;
      
      results.push({
        score: avgScore,
        item: {
          product_id: pid,
          standard_name: stdName,
          make: make,
          customer_id: customerId,
          customer_name: custName,
          price: pr.price,
          previous_price: pr.previous_price,
          last_updated: pr.last_updated,
          price_history: pr.price_history,
          is_configured: true
        }
      });
    }
  }
  
  results.sort(function(a, b) { return b.score - a.score; });
  return results.map(function(r) { return r.item; });
}

// -------------------------------------------------------------
// Action: Create Product
// -------------------------------------------------------------
function createProduct(payload) {
  var ss = getSpreadsheet();
  if (!ss) throw new Error("Google Spreadsheet not accessible.");
  ensureSheetsExist(ss);
  
  var pid = String(payload.product_id || "").trim();
  var name = String(payload.standard_name || "").trim();
  var make = String(payload.make || "").trim();
  
  if (!pid || !name) throw new Error("Product ID and Standard Product Name are required.");
  
  var sheet = ss.getSheetByName("Products");
  var data = sheet.getDataRange().getValues();

  var normName = name.toLowerCase().replace(/[^a-z0-9]/g, "");
  var normMake = make.toLowerCase().replace(/[^a-z0-9]/g, "");
  
  for (var i = 1; i < data.length; i++) {
    var exPid = String(data[i][0] || "").trim().toUpperCase();
    var exName = String(data[i][1] || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
    var exMake = String(data[i][2] || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");

    if (exPid === pid.toUpperCase()) {
      throw new Error("Product with ID '" + pid + "' already exists.");
    }

    if (exName === normName && exMake === normMake) {
      var makeText = make ? " with Make '" + make + "'" : "";
      throw new Error("Product '" + name + "'" + makeText + " already exists in catalog.");
    }
  }
  
  sheet.appendRow([pid, name, make]);
  return {
    product_id: pid,
    standard_name: name,
    make: make
  };
}

// -------------------------------------------------------------
// Action: Update Product
// -------------------------------------------------------------
function updateProduct(payload) {
  var ss = getSpreadsheet();
  if (!ss) throw new Error("Google Spreadsheet not accessible.");
  ensureSheetsExist(ss);
  
  var pid = String(payload.product_id || "").trim();
  var newName = String(payload.standard_name || "").trim();
  var newMake = String(payload.make || "").trim();
  
  if (!pid || !newName) throw new Error("Product ID and Standard Product Name are required.");
  
  var sheet = ss.getSheetByName("Products");
  var data = sheet.getDataRange().getValues();

  var normName = newName.toLowerCase().replace(/[^a-z0-9]/g, "");
  var normMake = newMake.toLowerCase().replace(/[^a-z0-9]/g, "");
  
  var targetRowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    var exPid = String(data[i][0] || "").trim().toUpperCase();
    var exName = String(data[i][1] || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");
    var exMake = String(data[i][2] || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");

    if (exPid === pid.toUpperCase()) {
      targetRowIndex = i + 1; // 1-indexed sheet row
    } else if (exName === normName && exMake === normMake) {
      var makeText = newMake ? " with Make '" + newMake + "'" : "";
      throw new Error("Product '" + newName + "'" + makeText + " already exists in catalog.");
    }
  }

  if (targetRowIndex <= 1) {
    throw new Error("Product with ID '" + pid + "' not found.");
  }
  
  sheet.getRange(targetRowIndex, 2).setValue(newName);
  sheet.getRange(targetRowIndex, 3).setValue(newMake);
  
  return {
    success: true,
    message: "Product updated successfully",
    product_id: pid,
    standard_name: newName,
    make: newMake
  };
}

// -------------------------------------------------------------
// Action: Delete Product
// -------------------------------------------------------------
function deleteProduct(productId) {
  var ss = getSpreadsheet();
  if (!ss) throw new Error("Google Spreadsheet not accessible.");
  ensureSheetsExist(ss);
  
  var pid = String(productId || "").trim();
  if (!pid) throw new Error("Product ID is required.");
  
  var sheet = ss.getSheetByName("Products");
  var data = sheet.getDataRange().getValues();
  
  var targetRowIndex = -1;
  var deletedProdName = "";
  for (var i = 1; i < data.length; i++) {
    var exPid = String(data[i][0] || "").trim().toUpperCase();
    if (exPid === pid.toUpperCase()) {
      targetRowIndex = i + 1;
      deletedProdName = String(data[i][1] || "").trim().toUpperCase();
      break;
    }
  }

  if (targetRowIndex <= 1) {
    throw new Error("Product with ID '" + pid + "' not found.");
  }

  // 1. Delete product row from Products sheet
  sheet.deleteRow(targetRowIndex);

  // 2. Cascade delete all customer-product price records from Pricing sheet (and Customer_Prices if present)
  var pricingSheetNames = ["Pricing", "Customer_Prices", "Customer Prices", "Prices"];
  for (var s = 0; s < pricingSheetNames.length; s++) {
    var pSheet = ss.getSheetByName(pricingSheetNames[s]);
    if (pSheet) {
      var pData = pSheet.getDataRange().getValues();
      if (pData.length > 1) {
        var pHeader = pData[0] || [];
        var pidIdx = 1; // Default to column 2
        for (var h = 0; h < pHeader.length; h++) {
          var hName = String(pHeader[h] || "").trim().toLowerCase();
          if (hName === "product id" || hName === "product_id" || hName === "prod id" || hName === "pid") {
            pidIdx = h;
            break;
          }
        }
        for (var j = pData.length - 1; j >= 1; j--) {
          var rowVal = String(pData[j][pidIdx] || "").trim().toUpperCase();
          if (rowVal === pid.toUpperCase() || (deletedProdName && rowVal === deletedProdName)) {
            pSheet.deleteRow(j + 1);
          }
        }
      }
    }
  }

  return {
    success: true,
    message: "Product '" + pid + "' and all associated customer prices deleted successfully"
  };
}

// -------------------------------------------------------------
// Action: Create Customer
// -------------------------------------------------------------
function createCustomer(payload) {
  var ss = getSpreadsheet();
  if (!ss) throw new Error("Google Spreadsheet not accessible.");
  ensureSheetsExist(ss);
  
  var cid = String(payload.customer_id || "").trim();
  var name = String(payload.customer_name || "").trim();
  var gstin = String(payload.gstin || "").trim();
  
  if (!cid || !name) throw new Error("Customer ID and Customer Name are required.");
  
  var sheet = ss.getSheetByName("Customers");
  var data = sheet.getDataRange().getValues();
  var headerRow = data[0] || [];

  var cidIdx = 0, nameIdx = 1, gstinIdx = -1;
  for (var h = 0; h < headerRow.length; h++) {
    var hName = String(headerRow[h] || "").trim().toLowerCase();
    if (hName === "customer id" || hName === "customer_id" || hName === "id" || hName.indexOf("cust id") !== -1) {
      cidIdx = h;
    } else if (hName === "customer name" || hName === "customer_name" || hName === "name" || hName.indexOf("cust name") !== -1) {
      nameIdx = h;
    } else if (hName.indexOf("gst") !== -1 || hName.indexOf("tin") !== -1 || hName.indexOf("tax") !== -1) {
      gstinIdx = h;
    }
  }

  if (gstinIdx === -1) {
    gstinIdx = headerRow.length >= 7 ? 6 : 2;
  }

  var normName = name.toLowerCase().replace(/[^a-z0-9]/g, "");
  for (var i = 1; i < data.length; i++) {
    var exCid = String(data[i][cidIdx] || "").trim().toUpperCase();
    var exName = String(data[i][nameIdx] || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");

    if (exCid === cid.toUpperCase()) {
      throw new Error("Customer with ID '" + cid + "' already exists.");
    }
    if (exName === normName) {
      throw new Error("Customer '" + name + "' already exists in database.");
    }
  }
  
  var newRowLength = Math.max(3, gstinIdx + 1);
  var newRow = new Array(newRowLength).fill("");
  newRow[cidIdx] = cid;
  newRow[nameIdx] = name;
  newRow[gstinIdx] = gstin;

  // If sheet is 7-column layout, populate column 7 (index 6) as well
  if (newRowLength >= 7 || headerRow.length >= 7) {
    while (newRow.length < 7) newRow.push("");
    newRow[6] = gstin;
  }
  
  sheet.appendRow(newRow);
  return {
    customer_id: cid,
    customer_name: name,
    gstin: gstin
  };
}

// -------------------------------------------------------------
// Action: Update Customer
// -------------------------------------------------------------
function updateCustomer(payload) {
  var ss = getSpreadsheet();
  if (!ss) throw new Error("Google Spreadsheet not accessible.");
  ensureSheetsExist(ss);
  
  var cid = String(payload.customer_id || "").trim();
  var newName = String(payload.customer_name || "").trim();
  var newGstin = String(payload.gstin || "").trim();
  
  if (!cid || !newName) throw new Error("Customer ID and Customer Name are required.");
  
  var sheet = ss.getSheetByName("Customers");
  var data = sheet.getDataRange().getValues();
  var headerRow = data[0] || [];

  var cidIdx = 0, nameIdx = 1, gstinIdx = -1;
  for (var h = 0; h < headerRow.length; h++) {
    var hName = String(headerRow[h] || "").trim().toLowerCase();
    if (hName === "customer id" || hName === "customer_id" || hName === "id" || hName.indexOf("cust id") !== -1) {
      cidIdx = h;
    } else if (hName === "customer name" || hName === "customer_name" || hName === "name" || hName.indexOf("cust name") !== -1) {
      nameIdx = h;
    } else if (hName.indexOf("gst") !== -1 || hName.indexOf("tin") !== -1 || hName.indexOf("tax") !== -1) {
      gstinIdx = h;
    }
  }

  var targetRowIndex = -1;
  var targetRow = null;
  var normName = newName.toLowerCase().replace(/[^a-z0-9]/g, "");
  
  for (var i = 1; i < data.length; i++) {
    var exCid = String(data[i][cidIdx] || "").trim().toUpperCase();
    var exName = String(data[i][nameIdx] || "").trim().toLowerCase().replace(/[^a-z0-9]/g, "");

    if (exCid === cid.toUpperCase()) {
      targetRowIndex = i + 1;
      targetRow = data[i];
    } else if (exName === normName) {
      throw new Error("Customer '" + newName + "' already exists in database.");
    }
  }

  if (targetRowIndex <= 1) {
    throw new Error("Customer with ID '" + cid + "' not found.");
  }

  if (gstinIdx === -1) {
    if (targetRow) {
      if (targetRow[6] !== undefined && String(targetRow[6]).trim() !== "") {
        gstinIdx = 6;
      } else if (targetRow[2] !== undefined && String(targetRow[2]).trim() !== "") {
        gstinIdx = 2;
      }
    }
    if (gstinIdx === -1) {
      gstinIdx = (headerRow.length >= 7) ? 6 : 2;
    }
  }
  
  sheet.getRange(targetRowIndex, nameIdx + 1).setValue(newName);
  sheet.getRange(targetRowIndex, gstinIdx + 1).setValue(newGstin);
  
  // Dual-sync: update both column 3 (index 2) and column 7 (index 6) if row has 7+ columns
  if (targetRow && targetRow.length >= 7) {
    sheet.getRange(targetRowIndex, 7).setValue(newGstin);
    if (gstinIdx !== 2) {
      // Also update column 3 if it was used
      sheet.getRange(targetRowIndex, 3).setValue(newGstin);
    }
  }
  
  return {
    success: true,
    message: "Customer updated successfully",
    customer_id: cid,
    customer_name: newName,
    gstin: newGstin
  };
}

// -------------------------------------------------------------
// Action: Delete Customer
// -------------------------------------------------------------
function deleteCustomer(customerId) {
  var ss = getSpreadsheet();
  if (!ss) throw new Error("Google Spreadsheet not accessible.");
  ensureSheetsExist(ss);
  
  var cid = String(customerId || "").trim();
  if (!cid) throw new Error("Customer ID is required.");
  
  var sheet = ss.getSheetByName("Customers");
  var data = sheet.getDataRange().getValues();
  
  var headerRow = data[0] || [];
  var cidIdx = 0, nameIdx = 1;
  for (var h = 0; h < headerRow.length; h++) {
    var hName = String(headerRow[h] || "").trim().toLowerCase();
    if (hName === "customer id" || hName === "customer_id" || hName === "id" || hName.indexOf("cust id") !== -1) {
      cidIdx = h;
    } else if (hName === "customer name" || hName === "customer_name" || hName === "name" || hName.indexOf("cust name") !== -1) {
      nameIdx = h;
    }
  }

  var targetRowIndex = -1;
  var deletedCustName = "";
  for (var i = 1; i < data.length; i++) {
    var exCid = String(data[i][cidIdx] || "").trim().toUpperCase();
    if (exCid === cid.toUpperCase()) {
      targetRowIndex = i + 1;
      deletedCustName = String(data[i][nameIdx] || "").trim().toUpperCase();
      break;
    }
  }

  if (targetRowIndex <= 1) {
    throw new Error("Customer with ID '" + cid + "' not found.");
  }

  // 1. Delete customer row from Customers sheet
  sheet.deleteRow(targetRowIndex);

  // 2. Cascade delete all customer-product price records from Pricing sheet (and Customer_Prices if present)
  var pricingSheetNames = ["Pricing", "Customer_Prices", "Customer Prices", "Prices"];
  for (var s = 0; s < pricingSheetNames.length; s++) {
    var pSheet = ss.getSheetByName(pricingSheetNames[s]);
    if (pSheet) {
      var pData = pSheet.getDataRange().getValues();
      if (pData.length > 1) {
        var pHeader = pData[0] || [];
        var custColIdx = 0; // Default to column 1
        for (var h = 0; h < pHeader.length; h++) {
          var hName = String(pHeader[h] || "").trim().toLowerCase();
          if (hName === "customer id" || hName === "customer_id" || hName === "cust id" || hName === "cid") {
            custColIdx = h;
            break;
          }
        }
        for (var k = pData.length - 1; k >= 1; k--) {
          var rowVal = String(pData[k][custColIdx] || "").trim().toUpperCase();
          if (rowVal === cid.toUpperCase() || (deletedCustName && rowVal === deletedCustName)) {
            pSheet.deleteRow(k + 1);
          }
        }
      }
    }
  }

  return {
    success: true,
    message: "Customer '" + cid + "' and all associated product prices deleted successfully"
  };
}



// -------------------------------------------------------------
// Action: Upsert Customer Price
// -------------------------------------------------------------
function upsertPricing(payload) {
  var ss = getSpreadsheet();
  if (!ss) throw new Error("Google Spreadsheet not accessible.");
  ensureSheetsExist(ss);
  
  var cid = String(payload.customer_id || "").trim();
  var pid = String(payload.product_id || "").trim();
  var newPrice = parseFloat(payload.price);
  
  if (!cid || !pid || isNaN(newPrice)) {
    throw new Error("Customer ID, Product ID, and numeric price are required.");
  }
  
  var sheet = ss.getSheetByName("Pricing");
  var data = sheet.getDataRange().getValues();
  var todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  function formatSimpleDate(d) {
    if (!d) return todayStr;
    if (d instanceof Date) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    var str = String(d).trim();
    if (str.indexOf("GMT") !== -1 || str.indexOf("Standard Time") !== -1 || str.length > 20) {
      try {
        var parsed = new Date(str);
        if (!isNaN(parsed.getTime())) {
          return Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
      } catch (e) {}
    }
    if (str.indexOf("T") !== -1) str = str.split("T")[0];
    if (str.indexOf(" ") !== -1 && str.length > 10) str = str.split(" ")[0];
    return str || todayStr;
  }
  
  var existingRowIndex = -1;
  for (var i = 1; i < data.length; i++) {
    var rowCid = String(data[i][0]).trim().toUpperCase();
    var rowPid = String(data[i][1]).trim().toUpperCase();
    if (rowCid === cid.toUpperCase() && rowPid === pid.toUpperCase()) {
      existingRowIndex = i + 1; // 1-indexed sheet row
      break;
    }
  }
  
  var prevPrice = null;
  var historyStr = "";
  var isUpdate = false;
  
  if (existingRowIndex > 1) {
    isUpdate = true;
    var oldRow = data[existingRowIndex - 1];
    var oldPrice = parseFloat(oldRow[2]);
    var oldDate = formatSimpleDate(oldRow[4]);
    var oldHistory = String(oldRow[5] || "").trim();
    
    if (!isNaN(oldPrice) && oldPrice !== newPrice) {
      prevPrice = oldPrice;
      var formattedOldPrice = "₹" + oldPrice.toFixed(2).replace(/\d(?=(\d{3})+\.)/g, '$&,');
      var newEntry = formattedOldPrice + " (" + oldDate + ")";
      historyStr = oldHistory ? (newEntry + ", " + oldHistory) : newEntry;
    } else {
      prevPrice = (oldRow[3] !== "" && oldRow[3] !== null) ? parseFloat(oldRow[3]) : null;
      historyStr = oldHistory;
    }
    
    sheet.getRange(existingRowIndex, 3).setValue(newPrice);
    sheet.getRange(existingRowIndex, 4).setValue(prevPrice !== null ? prevPrice : "");
    sheet.getRange(existingRowIndex, 5).setValue(todayStr);
    sheet.getRange(existingRowIndex, 6).setValue(historyStr);
  } else {
    isUpdate = false;
    sheet.appendRow([cid, pid, newPrice, "", todayStr, ""]);
  }
  
  return {
    success: true,
    message: isUpdate ? "Pricing updated successfully" : "Pricing created successfully",
    data: {
      customer_id: cid,
      product_id: pid,
      price: newPrice,
      previous_price: prevPrice,
      last_updated: todayStr,
      price_history: historyStr
    }
  };
}

// Helper: Format Date as YYYY-MM-DD (removes GMT/timezone text)
function formatDateOnly(val) {
  if (!val) return "";
  if (Object.prototype.toString.call(val) === "[object Date]" || val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  var s = String(val).trim();
  if (s.indexOf("GMT") !== -1 || s.indexOf("Standard Time") !== -1) {
    try {
      var d = new Date(s);
      if (!isNaN(d.getTime())) {
        return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
    } catch (e) {}
  }
  if (s.indexOf("T") !== -1) {
    return s.split("T")[0];
  }
  return s;
}

// -------------------------------------------------------------
// Action: Save Quotation to 'Quotations' Sheet
// -------------------------------------------------------------
function saveQuotation(payload) {
  var ss = getSpreadsheet();
  if (!ss) throw new Error("Google Spreadsheet not accessible.");
  ensureSheetsExist(ss);
  
  if (!payload) throw new Error("Quotation payload is empty.");
  
  var quotationNo = String(payload.quotation_no || payload.quotationNo || "").trim();
  if (!quotationNo) {
    var now = new Date();
    quotationNo = "ARH-QT-" + Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyyMMdd-HHmmss") + "-" + Math.floor(100 + Math.random() * 900);
  }
  
  var quotationDate = formatDateOnly(payload.quotation_date || payload.quotationDate);
  var validUntil = formatDateOnly(payload.valid_until || payload.validUntil);
  var customerId = String(payload.customer_id || payload.customerId || "").trim();
  var customerName = String(payload.customer_name || payload.customerName || "").trim();
  var customerGstin = String(payload.customer_gstin || payload.customerGstin || "").trim();
  
  var subtotal = Number(payload.subtotal) || 0;
  var discountAmount = Number(payload.discount_amount !== undefined ? payload.discount_amount : payload.discountAmount) || 0;
  var taxableAmount = Number(payload.taxable_amount !== undefined ? payload.taxable_amount : payload.taxableAmount) || 0;
  var taxRate = Number(payload.tax_rate !== undefined ? payload.tax_rate : payload.taxRate) || 0;
  var taxAmount = Number(payload.tax_amount !== undefined ? payload.tax_amount : payload.taxAmount) || 0;
  var roundOff = Number(payload.round_off !== undefined ? payload.round_off : payload.roundOff) || 0;
  var grandTotal = Number(payload.grand_total !== undefined ? payload.grand_total : payload.grandTotal) || 0;
  var terms = String(payload.terms || "").trim();
  var savedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  
  var items = Array.isArray(payload.items) ? payload.items : [];
  if (items.length === 0) {
    throw new Error("Quotation must contain at least one line item.");
  }
  
  var sheet = ss.getSheetByName("Quotations");
  if (!sheet) {
    sheet = ss.insertSheet("Quotations");
    var headers = [
      "Quotation No", "Quotation Date", "Valid Till", "Customer ID", "Customer Name", "Customer GSTIN",
      "Item #", "Product ID", "Product Name", "Make / Brand", "UOM", "Qty", "Unit Rate", "Total Amount",
      "Remark", "Subtotal", "Discount", "Taxable Amount", "Tax Rate", "Tax Amount",
      "Round Off", "Grand Total", "Terms & Conditions", "Saved At", "Order Status", "Confirmed At"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f1f5f9");
  } else if (sheet.getLastRow() >= 1 && sheet.getLastColumn() >= 1) {
    // Automatically clean (₹) and (%) from existing header row if present
    var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var needsHeaderUpdate = false;
    for (var c = 0; c < currentHeaders.length; c++) {
      var h = String(currentHeaders[c] || "");
      if (h.indexOf("(₹)") !== -1 || h.indexOf("(%)") !== -1) {
        currentHeaders[c] = h.replace(/\s*\(₹\)/g, "").replace(/\s*\(%\)/g, "").trim();
        needsHeaderUpdate = true;
      }
    }
    if (needsHeaderUpdate) {
      sheet.getRange(1, 1, 1, currentHeaders.length).setValues([currentHeaders]);
    }
  }
  
  var orderStatus = String(payload.order_status || payload.orderStatus || "PENDING").trim();
  var confirmedAt = String(payload.confirmed_at || payload.confirmedAt || "").trim();

  // 1. Check if an existing quotation with this number is already CONFIRMED
  var data = sheet.getDataRange().getValues();
  if (data.length > 1) {
    for (var i = 1; i < data.length; i++) {
      var rowQNo = String(data[i][0] || "").trim().toUpperCase();
      var rowStatus = String(data[i][24] || "").trim().toUpperCase();
      if (rowQNo === quotationNo.toUpperCase() && rowStatus === "CONFIRMED") {
        throw new Error("Quotation '" + quotationNo + "' is already confirmed as an order and cannot be edited or modified.");
      }
    }
  }

  // 2. Delete previous draft rows for this Quotation No to prevent duplicates when updating
  if (data.length > 1) {
    for (var i = data.length - 1; i >= 1; i--) {
      var rowQNo = String(data[i][0] || "").trim();
      if (rowQNo.toUpperCase() === quotationNo.toUpperCase()) {
        sheet.deleteRow(i + 1);
      }
    }
  }
  
  // 2. Prepare new rows to append
  var rowsToAppend = [];
  for (var idx = 0; idx < items.length; idx++) {
    var item = items[idx];
    var itemNo = idx + 1;
    var productId = String(item.product_id || item.productId || "").trim();
    var productName = String(item.product_name || item.productName || item.standard_name || "").trim();
    var make = String(item.make || "").trim();
    var uom = String(item.uom || "Nos").trim();
    var qty = Number(item.qty) || 0;
    var rate = Number(item.rate) || 0;
    var discountVal = Number(item.discount_val) || 0;
    var discountType = String(item.discount_type || "percent");
    var discountedRate = item.discounted_rate !== undefined ? Number(item.discounted_rate) : (discountVal > 0 ? (discountType === 'fixed' ? Math.max(0, rate - discountVal) : Math.max(0, rate - (rate * discountVal)/100)) : rate);
    var total = (Number(item.total) !== undefined && Number(item.total) > 0) ? Number(item.total) : (qty * discountedRate);
    var remark = String(item.remark || "").trim();
    
    rowsToAppend.push([
      quotationNo,
      quotationDate,
      validUntil,
      customerId,
      customerName,
      customerGstin,
      itemNo,
      productId,
      productName,
      make,
      uom,
      qty,
      rate,
      total,
      remark,
      subtotal,
      discountAmount,
      taxableAmount,
      taxRate,
      taxAmount,
      roundOff,
      grandTotal,
      terms,
      savedAt,
      orderStatus,
      confirmedAt,
      discountVal,
      discountType,
      discountedRate
    ]);
  }
  
  if (rowsToAppend.length > 0) {
    var startRow = sheet.getLastRow() + 1;
    var numRows = rowsToAppend.length;
    var numCols = rowsToAppend[0].length;
    sheet.getRange(startRow, 1, numRows, numCols).setValues(rowsToAppend);
  }
  
  return {
    success: true,
    message: "Quotation '" + quotationNo + "' saved successfully to Google Sheets ('Quotations' sheet)",
    quotation_no: quotationNo,
    items_count: items.length,
    grand_total: grandTotal,
    saved_at: savedAt,
    order_status: orderStatus,
    confirmed_at: confirmedAt
  };
}

// -------------------------------------------------------------
// Action: Retrieve Saved Quotations
// -------------------------------------------------------------
function getQuotations() {
  var ss = getSpreadsheet();
  if (!ss) return [];
  ensureSheetsExist(ss);
  
  var sheet = ss.getSheetByName("Quotations");
  if (!sheet) return [];
  
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  
  var quotationsMap = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var qNo = String(row[0] || "").trim();
    if (!qNo) continue;
    
    if (!quotationsMap[qNo]) {
      quotationsMap[qNo] = {
        quotation_no: qNo,
        quotation_date: formatDateOnly(row[1]),
        valid_until: formatDateOnly(row[2]),
        customer_id: String(row[3] || ""),
        customer_name: String(row[4] || ""),
        customer_gstin: String(row[5] || ""),
        subtotal: Number(row[15]) || 0,
        discount_amount: Number(row[16]) || 0,
        taxable_amount: Number(row[17]) || 0,
        tax_rate: Number(row[18]) || 0,
        tax_amount: Number(row[19]) || 0,
        round_off: Number(row[20]) || 0,
        grand_total: Number(row[21]) || 0,
        terms: String(row[22] || ""),
        saved_at: formatDateOnly(row[23]),
        order_status: String(row[24] || "PENDING").trim(),
        confirmed_at: formatDateOnly(row[25]),
        items: []
      };
    }
    
    var itemRate = Number(row[12]) || 0;
    var itemTotal = Number(row[13]) || 0;
    var itemQty = Number(row[11]) || 1;
    var itemDiscVal = (row.length > 26 && row[26] !== "" && !isNaN(Number(row[26]))) ? Number(row[26]) : 0;
    var itemDiscType = (row.length > 27 && row[27]) ? String(row[27]) : "percent";
    var itemDiscRate = (row.length > 28 && row[28] !== "" && !isNaN(Number(row[28]))) ? Number(row[28]) : null;

    if (itemDiscRate === null || isNaN(itemDiscRate)) {
      if (itemDiscVal > 0) {
        itemDiscRate = itemDiscType === 'fixed' ? Math.max(0, itemRate - itemDiscVal) : Math.max(0, itemRate - (itemRate * itemDiscVal) / 100);
      } else if (itemQty > 0 && itemTotal > 0 && itemTotal < itemQty * itemRate) {
        itemDiscRate = Math.round((itemTotal / itemQty) * 100) / 100;
        itemDiscVal = Math.round(((itemRate - itemDiscRate) / itemRate) * 1000) / 10;
        itemDiscType = "percent";
      } else {
        itemDiscRate = itemRate;
      }
    }

    quotationsMap[qNo].items.push({
      item_no: row[6],
      product_id: row[7],
      product_name: row[8],
      make: row[9],
      uom: row[10],
      qty: itemQty,
      rate: itemRate,
      discount_val: itemDiscVal,
      discount_type: itemDiscType,
      discounted_rate: itemDiscRate,
      total: itemTotal,
      remark: row[14]
    });
  }
  
  var result = [];
  for (var key in quotationsMap) {
    result.push(quotationsMap[key]);
  }
  return result;
}

// -------------------------------------------------------------
// Action: Confirm Quotation Order & Update Customer Pricing
// -------------------------------------------------------------
function confirmQuotationOrder(payload) {
  var ss = getSpreadsheet();
  if (!ss) throw new Error("Google Spreadsheet not accessible.");
  ensureSheetsExist(ss);
  
  var quotationNo = String(payload.quotation_no || payload.quotationNo || "").trim();
  var customerId = String(payload.customer_id || payload.customerId || "").trim();
  var items = Array.isArray(payload.items) ? payload.items : [];
  
  if (!customerId) throw new Error("Customer ID is required to confirm order.");
  if (items.length === 0) throw new Error("No line items found to confirm.");

  // Check if this quotation was already confirmed
  var quoteSheet = ss.getSheetByName("Quotations");
  if (quoteSheet && quotationNo) {
    var qData = quoteSheet.getDataRange().getValues();
    if (qData.length > 1) {
      for (var r = 1; r < qData.length; r++) {
        var rQNo = String(qData[r][0] || "").trim().toUpperCase();
        var rStatus = String(qData[r][24] || "").trim().toUpperCase();
        if (rQNo === quotationNo.toUpperCase() && rStatus === "CONFIRMED") {
          throw new Error("Quotation '" + quotationNo + "' has already been confirmed as an order and cannot be re-confirmed.");
        }
      }
    }
  }
  
  // 1. Build map of products for name-to-ID fallback matching
  var prodSheet = ss.getSheetByName("Products");
  var prodData = prodSheet ? prodSheet.getDataRange().getValues() : [];
  var prodNameMap = {};
  if (prodData.length > 1) {
    for (var p = 1; p < prodData.length; p++) {
      var pId = String(prodData[p][0] || "").trim();
      var pName = String(prodData[p][1] || "").trim().toUpperCase();
      if (pId && pName) {
        prodNameMap[pName] = pId;
      }
    }
  }
  
  // 2. Loop through each item and update pricing with history
  var updatedItems = [];
  var skippedItems = [];
  
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    var pid = String(it.product_id || it.productId || "").trim();
    var pName = String(it.product_name || it.productName || it.standard_name || "").trim();
    var rate = parseFloat(it.discounted_rate !== undefined && it.discounted_rate !== null ? it.discounted_rate : (it.rate !== undefined ? it.rate : it.price));
    
    // Fallback: match by product name if product_id is empty
    if (!pid && pName && prodNameMap[pName.toUpperCase()]) {
      pid = prodNameMap[pName.toUpperCase()];
    }
    
    if (pid && !isNaN(rate) && rate > 0) {
      try {
        var upRes = upsertPricing({
          customer_id: customerId,
          product_id: pid,
          price: rate
        });
        updatedItems.push({
          product_id: pid,
          product_name: pName,
          price: rate,
          previous_price: upRes.data ? upRes.data.previous_price : null,
          history: upRes.data ? upRes.data.price_history : ""
        });
      } catch (e) {
        skippedItems.push({ product_id: pid, reason: e.message });
      }
    } else {
      skippedItems.push({ product_id: pid || pName, reason: "Missing product ID or invalid price" });
    }
  }
  
  // 3. Mark Quotations sheet status as CONFIRMED
  var confirmTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var quoteSheet = ss.getSheetByName("Quotations");
  if (quoteSheet && quotationNo) {
    var qData = quoteSheet.getDataRange().getValues();
    if (qData.length > 1) {
      var headers = qData[0];
      var statusColIdx = 24; // Column Y (0-indexed 24)
      var confirmedAtColIdx = 25; // Column Z (0-indexed 25)
      
      // Look for dynamic header indices if columns were shifted
      for (var h = 0; h < headers.length; h++) {
        var hName = String(headers[h] || "").trim().toLowerCase();
        if (hName === "order status" || hName === "status") statusColIdx = h;
        if (hName === "confirmed at" || hName === "order confirmed at") confirmedAtColIdx = h;
      }
      
      // Ensure headers exist if sheet had fewer columns
      if (headers.length <= statusColIdx) {
        quoteSheet.getRange(1, statusColIdx + 1).setValue("Order Status").setFontWeight("bold").setBackground("#f1f5f9");
      }
      if (headers.length <= confirmedAtColIdx) {
        quoteSheet.getRange(1, confirmedAtColIdx + 1).setValue("Confirmed At").setFontWeight("bold").setBackground("#f1f5f9");
      }
      
      for (var r = 1; r < qData.length; r++) {
        if (String(qData[r][0] || "").trim().toUpperCase() === quotationNo.toUpperCase()) {
          quoteSheet.getRange(r + 1, statusColIdx + 1).setValue("CONFIRMED");
          quoteSheet.getRange(r + 1, confirmedAtColIdx + 1).setValue(confirmTime);
        }
      }
    }
  }
  
  return {
    success: true,
    message: "Order confirmed successfully! Updated customer prices for " + updatedItems.length + " item(s) with price history.",
    quotation_no: quotationNo,
    customer_id: customerId,
    updated_count: updatedItems.length,
    updated_items: updatedItems,
    skipped_items: skippedItems,
    confirmed_at: confirmTime
  };
}

// -------------------------------------------------------------
// Helper: Levenshtein Distance Similarity
// -------------------------------------------------------------
function getLevenshteinSimilarity(s1, s2) {
  var m = s1.length, n = s2.length;
  if (m === 0 || n === 0) return 0;
  
  var dp = [];
  for (var i = 0; i <= m; i++) {
    dp[i] = [i];
  }
  for (var j = 0; j <= n; j++) {
    dp[0][j] = j;
  }
  
  for (var i = 1; i <= m; i++) {
    for (var j = 1; j <= n; j++) {
      if (s1[i - 1] === s2[j - 1]) {
        dp[i][j] = dp[i - 1][j - 1];
      } else {
        dp[i][j] = 1 + Math.min(dp[i - 1][j], dp[i][j - 1], dp[i - 1][j - 1]);
      }
    }
  }
  return 1 - (dp[m][n] / Math.max(m, n));
}

// -------------------------------------------------------------
// Helper: Ensure Worksheets & Headers Exist
// -------------------------------------------------------------
function ensureSheetsExist(ss) {
  var sheetsToCreate = [
    { name: "Customers", headers: ["Customer ID", "Customer Name", "Contact Person", "Phone", "Email", "Address", "GSTIN"] },
    { name: "Products", headers: ["Product ID", "Standard Product Name", "Make"] },
    { name: "Pricing", headers: ["Customer ID", "Product ID", "Customer Price", "Previous Price", "Last Updated", "Price History"] },
    { 
      name: "Quotations", 
      headers: [
        "Quotation No",
        "Quotation Date",
        "Valid Till",
        "Customer ID",
        "Customer Name",
        "Customer GSTIN",
        "Item #",
        "Product ID",
        "Product Name",
        "Make / Brand",
        "UOM",
        "Qty",
        "Unit Rate",
        "Total Amount",
        "Remark",
        "Subtotal",
        "Discount",
        "Taxable Amount",
        "Tax Rate",
        "Tax Amount",
        "Round Off",
        "Grand Total",
        "Terms & Conditions",
        "Saved At",
        "Order Status",
        "Confirmed At"
      ] 
    }
  ];
  
  for (var i = 0; i < sheetsToCreate.length; i++) {
    var item = sheetsToCreate[i];
    var sheet = ss.getSheetByName(item.name);
    if (!sheet) {
      sheet = ss.insertSheet(item.name);
      sheet.appendRow(item.headers);
      sheet.getRange(1, 1, 1, item.headers.length).setFontWeight("bold").setBackground("#f1f5f9");
    } else if (item.name === "Quotations" && sheet.getLastRow() >= 1 && sheet.getLastColumn() >= 1) {
      var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      var needsHeaderUpdate = false;
      for (var c = 0; c < currentHeaders.length; c++) {
        var h = String(currentHeaders[c] || "");
        if (h.indexOf("(₹)") !== -1 || h.indexOf("(%)") !== -1) {
          currentHeaders[c] = h.replace(/\s*\(₹\)/g, "").replace(/\s*\(%\)/g, "").trim();
          needsHeaderUpdate = true;
        }
      }
      if (needsHeaderUpdate) {
        sheet.getRange(1, 1, 1, currentHeaders.length).setValues([currentHeaders]);
      }
    }
  }
}
