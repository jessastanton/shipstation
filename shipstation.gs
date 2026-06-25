// ============================================================
// PROJECT SHIPSTATION — Logic (hosted on GitHub)
// Loaded and run by the Apps Script bootstrap (loadAndRun).
//
// NOTE: SS_KEY, SS_SECRET, and createDailyTrigger() do NOT live
// here — they stay in the Apps Script project. This file is
// safe to host in a PUBLIC repo because it contains no secrets.
// ============================================================

var RECIPIENTS = [
  'jim@kolikof.com',
  'paula@kolikof.com',
  'jessica@kolikof.com'
];

var SERVICE_MAP = {
  'fedex_2day':              'FedEx 2Day®',
  'fedex_ground':            'FedEx Ground',
  'fedex_home_delivery':     'FedEx Home Delivery®',
  'fedex_first_overnight':   'FedEx First Overnight®',
  'fedex_priority_overnight':'FedEx Priority Overnight®',
  'fedex_standard_overnight':'FedEx Standard Overnight®',
  'fedex_express_saver':     'FedEx Express Saver®',
  'ups_ground':              'UPS Ground'
};

var PACKAGE_MAP = {
  'fedex_medium_box_onerate': 'FedEx One Rate® Medium Box',
  'fedex_small_box_onerate':  'FedEx One Rate® Small Box',
  'fedex_large_box_onerate':  'FedEx One Rate® Large Box',
  'package':                  'Package'
};

// ------------------------------------------------------------
// MAIN — called by the bootstrap after this file is eval'd
// ------------------------------------------------------------
function runProjectShipstation() {
  var today   = getTodayPT();
  var dateStr = today.toLocaleDateString('en-US', { month: 'numeric', day: 'numeric', year: 'numeric' });
  var apiDate = Utilities.formatDate(today, 'America/Los_Angeles', 'yyyy-MM-dd');

  Logger.log('Running Project Shipstation for: ' + apiDate);

  // 1. Fetch shipments
  var shipments = fetchShipments(apiDate);
  if (!shipments || shipments.length === 0) {
    Logger.log('No shipments found for ' + apiDate);
    return;
  }

  // 2. Fetch order details
  var orderMap = fetchOrders(shipments);

  // 3. Build rows
  var rows = buildRows(shipments, orderMap);

  // 4. Write today's data to a single tab
  var sheetName = Utilities.formatDate(today, 'America/Los_Angeles', 'MMM d yyyy');
  writeToSheet(rows, sheetName);

  // 5. Export only that tab as Excel and email it
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ws = ss.getSheetByName(sheetName);
  var xlsxBlob = exportSheetAsXlsx(ss, ws);
  var fileName = 'PROJECT_SHIPSTATION_' + apiDate.replace(/-/g, '') + '.xlsx';
  xlsxBlob.setName(fileName);

  sendEmail(xlsxBlob, fileName, dateStr, rows.length);
  Logger.log('Email sent — ' + rows.length + ' shipments.');

  // 6. Wipe the sheet clean after sending
  wipeSheet(ss);
  Logger.log('Sheet wiped.');
}

// ------------------------------------------------------------
// HELPERS
// ------------------------------------------------------------

function getTodayPT() {
  var now = new Date();
  var ptString = now.toLocaleString('en-US', { timeZone: 'America/Los_Angeles' });
  return new Date(ptString);
}

function getSSAuth() {
  // SS_KEY and SS_SECRET are declared in the Apps Script bootstrap,
  // and are in scope here because this file is eval'd inside loadAndRun.
  return 'Basic ' + Utilities.base64Encode(SS_KEY + ':' + SS_SECRET);
}

function fetchShipments(apiDate) {
  var url = 'https://ssapi.shipstation.com/shipments'
    + '?shipDateStart=' + apiDate
    + '&shipDateEnd='   + apiDate
    + '&pageSize=500';

  var resp = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': getSSAuth() },
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    Logger.log('ShipStation API error: ' + resp.getContentText());
    return [];
  }

  var data = JSON.parse(resp.getContentText());
  return (data.shipments || []).filter(function(s) { return !s.voided; });
}

function fetchOrders(shipments) {
  var orderMap = {};
  shipments.forEach(function(s) {
    try {
      var url = 'https://ssapi.shipstation.com/orders/' + s.orderId;
      var resp = UrlFetchApp.fetch(url, {
        headers: { 'Authorization': getSSAuth() },
        muteHttpExceptions: true
      });
      if (resp.getResponseCode() === 200) {
        var o = JSON.parse(resp.getContentText());
        orderMap[o.orderId] = o;
      }
      Utilities.sleep(120);
    } catch(e) {
      Logger.log('Order fetch error for ' + s.orderId + ': ' + e);
    }
  });
  return orderMap;
}

function calcShippingCost(svcCode, pkgCode, state, actualCost) {
  var pkg = pkgCode || '';
  if (pkg.indexOf('onerate') !== -1) return 34.96;
  if (svcCode === 'fedex_2day' && pkg === 'package') return 100.0;
  if (svcCode === 'fedex_standard_overnight' && pkg === 'package') {
    return (['CA','NV','AZ'].indexOf(state) !== -1) ? 25.0 : 100.0;
  }
  if (svcCode === 'fedex_priority_overnight' && pkg === 'package') return 100.0;
  return actualCost || 0;
}

function buildRows(shipments, orderMap) {
  var today2  = getTodayPT();
  var dateFmt = (today2.getMonth()+1) + "/" + today2.getDate() + "/" + today2.getFullYear();

  // Group shipments by order number to combine duplicates
  var grouped = {};
  var orderOfFirst = [];
  shipments.forEach(function(s) {
    var key = s.orderNumber || String(s.orderId);
    if (!grouped[key]) { grouped[key] = []; orderOfFirst.push(key); }
    grouped[key].push(s);
  });

  var rows = [];
  orderOfFirst.forEach(function(key) {
    var group   = grouped[key];
    var s       = group[0];
    var o       = orderMap[s.orderId] || {};
    var svcCode = s.serviceCode || "";
    var pkgCode = s.packageCode || "";
    var state   = (s.shipTo || {}).state || "";
    var count   = group.length;

    var totalCost = 0;
    group.forEach(function(ship) {
      totalCost += calcShippingCost(svcCode, pkgCode, state, ship.shipmentCost);
    });

    var pkgDisplay = PACKAGE_MAP[pkgCode] || pkgCode;
    if (count > 1) pkgDisplay = count + " - " + pkgDisplay;

    rows.push({
      date:          dateFmt,
      orderNumber:   s.orderNumber || "",
      name:          (s.shipTo || {}).name || "",
      service:       SERVICE_MAP[svcCode] || svcCode,
      package:       pkgDisplay,
      state:         state,
      orderTotal:    o.orderTotal || 0,
      orderShipping: o.shippingAmount || 0,
      shippingCost:  totalCost,
      comments:      ""
    });
  });

  return rows;
}

function writeToSheet(rows, sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Clear all existing sheets, keep one and rename it
  var allSheets = ss.getSheets();
  if (allSheets.length > 1) {
    for (var i = 1; i < allSheets.length; i++) ss.deleteSheet(allSheets[i]);
  }
  var ws = allSheets[0];
  ws.setName(sheetName);
  ws.clearContents();
  ws.clearFormats();

  // Group by service, sort by order number within each group
  var groups = {};
  rows.forEach(function(r) {
    if (!groups[r.service]) groups[r.service] = [];
    groups[r.service].push(r);
  });
  Object.keys(groups).forEach(function(k) {
    groups[k].sort(function(a,b) { return String(a.orderNumber).localeCompare(String(b.orderNumber)); });
  });
  var sortedServices = Object.keys(groups).sort();

  var HEADERS    = ['Date - Shipped Date','Order - Number','Ship To - Name',
    'Shipment - Service','Package Type','State',
    'Order Total','Order Shipping','Shipping Cost','Comments'];
  var DARK_BLUE  = '#2F4F8F';
  var MED_BLUE   = '#4472C4';
  var LIGHT_BLUE = '#D9E1F2';
  var NAVY       = '#1F3864';
  var WHITE      = '#FFFFFF';

  var currentRow = 1;

  // Main header row
  var hdrRange = ws.getRange(currentRow, 1, 1, 10);
  hdrRange.setValues([HEADERS]);
  hdrRange.setBackground(DARK_BLUE).setFontColor(WHITE).setFontWeight('bold')
    .setFontFamily('Arial').setFontSize(10).setHorizontalAlignment('center');
  ws.setRowHeight(currentRow, 30);
  currentRow++;

  sortedServices.forEach(function(svcName) {
    var svcRows = groups[svcName];

    // Group header
    ws.getRange(currentRow, 1, 1, 10).merge()
      .setValue(svcName)
      .setBackground(MED_BLUE).setFontColor(WHITE).setFontWeight('bold')
      .setFontFamily('Arial').setFontSize(10);
    currentRow++;

    // Repeated column headers (no date in first column)
    var repeatHdrs = [''].concat(HEADERS.slice(1));
    ws.getRange(currentRow, 1, 1, 10).setValues([repeatHdrs])
      .setBackground(DARK_BLUE).setFontColor(WHITE).setFontWeight('bold')
      .setFontFamily('Arial').setFontSize(10).setHorizontalAlignment('center');
    currentRow++;

    var dataStart = currentRow;

    svcRows.forEach(function(r) {
      ws.getRange(currentRow, 1, 1, 10).setValues([[
        r.date, r.orderNumber, r.name, r.service, r.package,
        r.state, r.orderTotal, r.orderShipping, r.shippingCost, r.comments
      ]]).setFontFamily('Arial').setFontSize(10);
      ws.getRange(currentRow, 7, 1, 3).setNumberFormat('#,##0.00').setHorizontalAlignment('right');
      ws.getRange(currentRow, 2).setHorizontalAlignment('right');
      currentRow++;
    });

    var dataEnd = currentRow - 1;

    // Subtotal row
    ws.getRange(currentRow, 1, 1, 10).setValues([[
      '', '', 'SUBTOTAL — ' + svcName + ' (' + svcRows.length + ' orders)', '', '', '',
      '=SUM(G'+dataStart+':G'+dataEnd+')',
      '=SUM(H'+dataStart+':H'+dataEnd+')',
      '=SUM(I'+dataStart+':I'+dataEnd+')',
      ''
    ]]).setBackground(LIGHT_BLUE).setFontWeight('bold').setFontFamily('Arial').setFontSize(10);
    ws.getRange(currentRow, 7, 1, 3).setNumberFormat('#,##0.00').setHorizontalAlignment('right');
    currentRow++;

    // Blank spacer
    currentRow++;
  });

  // Grand Total row
  var lastData = currentRow - 2;
  ws.getRange(currentRow, 1, 1, 10).setValues([[
    '', '', 'GRAND TOTAL', '', '', '',
    '=SUMIF(A2:A'+lastData+',"<>"&"",G2:G'+lastData+')',
    '=SUMIF(A2:A'+lastData+',"<>"&"",H2:H'+lastData+')',
    '=SUMIF(A2:A'+lastData+',"<>"&"",I2:I'+lastData+')',
    ''
  ]]).setBackground(NAVY).setFontColor(WHITE).setFontWeight('bold')
    .setFontFamily('Arial').setFontSize(10);
  ws.getRange(currentRow, 7, 1, 3).setNumberFormat('#,##0.00').setHorizontalAlignment('right');

  ws.setFrozenRows(1);
  [100,100,170,155,225,56,90,100,90,175].forEach(function(w,i) { ws.setColumnWidth(i+1, w); });
}

function exportSheetAsXlsx(ss, ws) {
  var url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export'
    + '?format=xlsx'
    + '&gid=' + ws.getSheetId()
    + '&access_token=' + ScriptApp.getOAuthToken();
  return UrlFetchApp.fetch(url, { muteHttpExceptions: true }).getBlob();
}

function sendEmail(blob, fileName, dateStr, count) {
  var subject = 'Project Shipstation — ' + dateStr;
  var body    = 'Hi team,\n\nAttached is the Project Shipstation shipping report for ' + dateStr
    + '.\n\nTotal shipments: ' + count
    + '\n\nThis report is generated automatically each day at 5pm PT.\n\nKolikof';

  RECIPIENTS.forEach(function(email) {
    GmailApp.sendEmail(email, subject, body, {
      attachments: [blob.copyBlob().setName(fileName)],
      name: 'Kolikof Shipping Reports'
    });
  });
}

function wipeSheet(ss) {
  var allSheets = ss.getSheets();
  if (allSheets.length > 1) {
    for (var i = 1; i < allSheets.length; i++) ss.deleteSheet(allSheets[i]);
  }
  var ws = allSheets[0];
  ws.setName('Ready');
  ws.clearContents();
  ws.clearFormats();
  ws.getRange(1,1).setValue('Waiting for next run...');
}
