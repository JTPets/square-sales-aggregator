/**************************************************************
 * 1) Create custom menu and handle setup (API key, email, etc.)
 **************************************************************/
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Square API')
    .addItem('Set API Key', 'setApiKey')
    .addItem('Set Email Address', 'setEmailAddress')
    .addSeparator()
    .addItem('Start Aggregated Sales Processing', 'startAggregatedSalesProcessing')
    .addSeparator()
    // Runs every 3 hours instead of daily
    .addItem('Set 3-Hour Timer', 'create3HourTrigger')
    .addToUi();
}

// Prompt for Square API Key
function setApiKey() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Set Square API Key',
    'Please enter your Square API access token:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() == ui.Button.OK) {
    var apiKey = response.getResponseText().trim();
    if (apiKey) {
      PropertiesService.getDocumentProperties().setProperty('SQUARE_ACCESS_TOKEN', apiKey);
      ui.alert('Success', 'Your Square API access token has been saved.', ui.ButtonSet.OK);
    } else {
      ui.alert('Error', 'No API key entered.', ui.ButtonSet.OK);
    }
  } else {
    ui.alert('Operation cancelled.');
  }
}

// Prompt for Notification Email
function setEmailAddress() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Set Notification Email',
    'Please enter your email address:',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() == ui.Button.OK) {
    var emailAddress = response.getResponseText().trim();
    if (emailAddress) {
      PropertiesService.getDocumentProperties().setProperty('NOTIFICATION_EMAIL', emailAddress);
      ui.alert('Success', 'Your email address has been saved.', ui.ButtonSet.OK);
    } else {
      ui.alert('Error', 'No email address entered.', ui.ButtonSet.OK);
    }
  } else {
    ui.alert('Operation cancelled.');
  }
}

/**************************************************************
 * 2) Main function to start the 91-day aggregated sales process
 **************************************************************/
function startAggregatedSalesProcessing() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = 'Sales-Aggregated';

  try {
    // Clear or create the sheet
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.clear();
    } else {
      sheet = ss.insertSheet(sheetName);
    }

    // Fetch & write 91-day aggregated sales
    fetchAndWriteAggregatedSales(sheet);

    // Send success email (if email is set)
    var docProps = PropertiesService.getDocumentProperties();
    var emailAddress = docProps.getProperty('NOTIFICATION_EMAIL');
    if (emailAddress) {
      MailApp.sendEmail({
        to: emailAddress,
        subject: "Square 91-Day Aggregated Sales - SUCCESS",
        body: "Successfully fetched and aggregated the 91-day sales from Square."
      });
    }
  } catch (error) {
    Logger.log("Error (startAggregatedSalesProcessing): " + error);
    // Send failure email
    var docProps = PropertiesService.getDocumentProperties();
    var emailAddress = docProps.getProperty('NOTIFICATION_EMAIL');
    if (emailAddress) {
      MailApp.sendEmail({
        to: emailAddress,
        subject: "Square 91-Day Aggregated Sales - FAILED",
        body: "The Square aggregated sales data refresh failed:\n" + error
      });
    }
    displayAlert("An error occurred: " + error.message);
  }
}

/**************************************************************
 * 3) Time-driven trigger creation for an automatic refresh 
 *    every 3 hours.
 **************************************************************/
function create3HourTrigger() {
  // Remove old triggers to avoid duplicates
  deleteExistingTriggers();

  // Create a time-based trigger to run every 3 hours
  ScriptApp.newTrigger('startAggregatedSalesProcessing')
    .timeBased()
    .everyHours(3)
    .create();

  SpreadsheetApp.getUi().alert("A timer has been set to run the script every 3 hours.");
}

function deleteExistingTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'startAggregatedSalesProcessing') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

/**************************************************************
 * 4) Fetch 91-day COMPLETED orders, then aggregate QTY & Revenue
 *    by catalog_object_id, with columns per location.
 **************************************************************/
function fetchAndWriteAggregatedSales(sheet) {
  // 1) Get location data (IDs + names)
  var locationMap = fetchLocationData(); // { locId: locName }
  var locationIds = Object.keys(locationMap);
  if (!locationIds.length) {
    displayAlert("No locations found for this merchant.");
    return;
  }

  // 2) Calculate date range: last 91 days (13 weeks) in RFC 3339
  var endDate = new Date();
  var startDate = new Date();
  startDate.setDate(endDate.getDate() - 91);

  var startDateRFC3339 = toRfc3339(startDate);
  var endDateRFC3339 = toRfc3339(endDate);

  // 3) Fetch COMPLETED orders (all locations) in that time range
  var allOrders = fetchCompletedOrdersWithinPeriod(
    startDateRFC3339,
    endDateRFC3339,
    locationIds
  );
  if (!allOrders.length) {
    displayAlert("No completed orders found in the past 91 days.");
    return;
  }

  /**
   * 4) Tally up quantities & revenue by catalog_object_id.
   *
   * We'll store Variation ID, Item Name, Variation Name, etc.
   *
   * itemTally[catalogObjectId] = {
   *   variationId: string,         // "ID-B"
   *   itemName: string,
   *   variationName: string,
   *   totals: { qty: number, revenue: number },
   *   locationSales: {
   *     locId: { qty: number, revenue: number }
   *   }
   * }
   */
  var itemTally = {};

  allOrders.forEach(function(order) {
    var orderLocationId = order.location_id;
    if (!order.line_items || !order.line_items.length) {
      return;
    }

    order.line_items.forEach(function(li) {
      // Use catalog_object_id as the Variation ID
      var catalogObjectId = li.catalog_object_id || "N/A";
      var itemName = li.name || 'Unknown Item';
      var variationName = li.variation_name || '';

      // Parse quantity
      var qty = parseFloat(li.quantity || "0");

      // For revenue, we use lineItem.total_money.amount (in cents).
      var revenueCents = 0;
      if (li.total_money && typeof li.total_money.amount !== 'undefined') {
        revenueCents = parseInt(li.total_money.amount, 10);
      }

      // Initialize aggregator if needed
      if (!itemTally[catalogObjectId]) {
        itemTally[catalogObjectId] = {
          variationId: catalogObjectId, // "Variation ID (ID-B)"
          itemName: itemName,
          variationName: variationName,
          totals: { qty: 0, revenue: 0 },
          locationSales: {}
        };
      }

      // Ensure sub-object for location is initialized
      if (!itemTally[catalogObjectId].locationSales[orderLocationId]) {
        itemTally[catalogObjectId].locationSales[orderLocationId] = { qty: 0, revenue: 0 };
      }

      // Update location-level tallies
      itemTally[catalogObjectId].locationSales[orderLocationId].qty += qty;
      itemTally[catalogObjectId].locationSales[orderLocationId].revenue += revenueCents;

      // Update total tallies
      itemTally[catalogObjectId].totals.qty += qty;
      itemTally[catalogObjectId].totals.revenue += revenueCents;
    });
  });

  /**
   * 5) Build the header row. 
   * We'll have:
   *   Variation ID (ID-B), Item Name, Variation Name,
   *   then for each location: "LocName QTY", "LocName $"
   *   and finally: "Total QTY (91 days)", "Total Revenue (91 days)"
   */
  var headerRow = [
    "Variation ID (ID-B)",
    "Item Name",
    "Variation Name",
  ];

  locationIds.forEach(function(locId) {
    var locName = locationMap[locId];
    headerRow.push(locName + " QTY");
    headerRow.push(locName + " $");
  });
  
  headerRow.push("Total QTY (91 days)");
  headerRow.push("Total Revenue (91 days)");

  sheet.appendRow(headerRow);

  /**
   * 6) Convert the itemTally into final rows.
   *
   * Each row:
   * [ variationId, itemName, variationName, loc1Qty, loc1Revenue, loc2Qty, loc2Revenue, ... totalQty, totalRevenue ]
   */
  var allRows = [];
  for (var catalogObjectId in itemTally) {
    if (!itemTally.hasOwnProperty(catalogObjectId)) {
      continue;
    }
    var data = itemTally[catalogObjectId];
    var rowData = [
      data.variationId,  // Variation ID (ID-B)
      data.itemName,
      data.variationName
    ];

    // For each location, push QTY & Revenue
    locationIds.forEach(function(locId) {
      var locSales = data.locationSales[locId] || { qty: 0, revenue: 0 };
      rowData.push(locSales.qty);
      // Convert cents to currency format
      rowData.push((locSales.revenue / 100).toFixed(2));
    });

    // Finally, total QTY & total Revenue
    rowData.push(data.totals.qty);
    rowData.push((data.totals.revenue / 100).toFixed(2));

    allRows.push(rowData);
  }

  // 7) Write all item rows at once
  if (allRows.length) {
    sheet
      .getRange(sheet.getLastRow() + 1, 1, allRows.length, headerRow.length)
      .setValues(allRows);
  }

  displayAlert(
    "Aggregated item-level sales for 91 days (13 weeks) has been written to '" +
    sheet.getName() +
    "'."
  );
}

/**************************************************************
 * 5) Get COMPLETED orders for all (or selected) locations/time
 **************************************************************/
function fetchCompletedOrdersWithinPeriod(startDate, endDate, locationIds) {
  var orders = [];
  var body = {
    location_ids: locationIds,
    limit: 50,
    query: {
      filter: {
        state_filter: { states: ['COMPLETED'] },
        date_time_filter: {
          closed_at: {
            start_at: startDate,
            end_at: endDate
          }
        }
      },
      sort: { sort_field: 'CLOSED_AT' }
    }
  };

  var url = 'https://connect.squareup.com/v2/orders/search';
  var cursor = null;

  do {
    if (cursor) {
      body.cursor = cursor;
    }
    var options = {
      method: 'POST',
      contentType: 'application/json',
      muteHttpExceptions: true,
      payload: JSON.stringify(body)
    };
    var response = makeApiRequest(url, options);
    var jsonData = JSON.parse(response.getContentText());
    if (jsonData && jsonData.orders) {
      orders = orders.concat(jsonData.orders);
    }
    cursor = jsonData.cursor || null;
  } while (cursor);

  return orders;
}

/**************************************************************
 * 6) Fetch location data (IDs & names) from /v2/locations
 **************************************************************/
function fetchLocationData() {
  var locationMap = {};
  var url = 'https://connect.squareup.com/v2/locations';
  var options = {
    method: 'GET',
    headers: {
      "Square-Version": "2023-10-18",
      "Content-Type": "application/json"
    },
    muteHttpExceptions: true
  };

  var response = makeApiRequest(url, options);
  if (response.getResponseCode() === 200) {
    var jsonData = JSON.parse(response.getContentText());
    if (Array.isArray(jsonData.locations)) {
      jsonData.locations.forEach(function(loc) {
        var locId = loc.id;
        var locName = loc.name || 'Unnamed';
        locationMap[locId] = locName;
      });
    }
  } else {
    Logger.log("Error retrieving locations: " + response.getContentText());
    displayAlert("Error retrieving locations. Check logs.");
  }
  return locationMap;
}

/**************************************************************
 * 7) Generic Helpers: makeApiRequest, displayAlert, date format
 **************************************************************/
function makeApiRequest(url, options) {
  var docProps = PropertiesService.getDocumentProperties();
  var accessToken = docProps.getProperty('SQUARE_ACCESS_TOKEN');
  if (!accessToken) {
    displayAlert('Square Access Token not set. Use "Set API Key" first.');
    throw new Error('Access token is missing.');
  }
  // Ensure we have headers
  if (!options.headers) {
    options.headers = {};
  }
  options.headers["Authorization"] = "Bearer " + accessToken;
  if (!options.headers["Square-Version"]) {
    options.headers["Square-Version"] = "2023-10-18";
  }

  var response = UrlFetchApp.fetch(url, options);
  var statusCode = response.getResponseCode();

  // 401 = invalid/expired token
  if (statusCode === 401) {
    var emailAddress = docProps.getProperty('NOTIFICATION_EMAIL');
    if (emailAddress) {
      MailApp.sendEmail({
        to: emailAddress,
        subject: "Square Aggregated Sales Failed - Invalid Access Token",
        body: "Your Square access token is invalid or expired. Please update it."
      });
    }
    throw new Error('Access token is invalid or expired.');
  } else if (statusCode >= 200 && statusCode < 300) {
    return response; // success
  } else {
    Logger.log('API request failed: ' + statusCode + ' -> ' + response.getContentText());
    throw new Error('API request failed with status code ' + statusCode);
  }
}

function displayAlert(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (e) {
    Logger.log("Alert: " + message);
  }
}

// Convert JS Date to RFC3339
function toRfc3339(dateObj) {
  var year = dateObj.getUTCFullYear();
  var month = padNumber(dateObj.getUTCMonth() + 1);
  var day = padNumber(dateObj.getUTCDate());
  var hours = padNumber(dateObj.getUTCHours());
  var minutes = padNumber(dateObj.getUTCMinutes());
  var seconds = padNumber(dateObj.getUTCSeconds());
  return (
    year + '-' + month + '-' + day + 'T' +
    hours + ':' + minutes + ':' + seconds + '.000Z'
  );
}

function padNumber(num) {
  return (num < 10 ? '0' : '') + num;
}
