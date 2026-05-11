// ════════════════════════════════════════════════════════════════════
// CONFIGURATION
// ════════════════════════════════════════════════════════════════════
const TELEGRAM_TOKEN = '8797809778:AAG_leZ3D3oUU8cC6lGmpCbwPo-WWCS9mVg';
const PREP_GROUP_ID  = '-5238710771';
const STOCK_GROUP_ID = '-5136469511';

const ALL_SITES = [
  'SD Withington', 'SD Wythenshawe', 'SD Cheadle', 'SD Newcastle',
  'SD EH 12', 'SD EH 15', 'DC Withington', 'DC Cheadle', 'DC Oxford Road'
];

const LOG_HEADERS = [
  'Timestamp', 'Site', 'Item Name', 'Unit', 'Qty', 'Supplier',
  'Price (£)', 'Total (£)', 'Notes', 'Delivery Date',
  'Telegram Sent', 'Order ID', 'Date', 'Month-Year'
];

// One row per ORDER — used by Looker Studio and the HTML summary
const SUMMARY_HEADERS = [
  'Order ID', 'Site', 'Order Type', 'Submitted', 'Delivery Date',
  'Items', 'Total Value (£)', 'Prep Telegram', 'Stock Telegram',
  'Notes', 'Date', 'Month-Year'
];

// ════════════════════════════════════════════════════════════════════
// doGet — routes to product list OR weekly summary
// ════════════════════════════════════════════════════════════════════
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) || 'products';
  if (action === 'summary')   return handleSummaryGet(e);
  if (action === 'dashboard') return handleDashboardGet(e);
  return handleProductsGet(e);
}

// ── Products handler (original logic) ───────────────────────────────
function handleProductsGet(e) {
  try {
    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const priceWs = ss.getSheetByName('Price List');
    const rows    = priceWs.getDataRange().getValues();

    const SITE_COLS = {
      'SD Withington':  10,
      'SD Wythenshawe': 11,
      'SD Cheadle':     12,
      'SD Newcastle':   13,
      'SD EH 12':       14,
      'SD EH 15':       15,
      'DC Withington':  16,
      'DC Cheadle':     17,
      'DC Oxford Road': 18,
    };

    const CAT_DISPLAY = {
      'Raw':                          'Meat & Protein',
      'Fresh':                        'Fresh Produce',
      'Frozen':                       'Frozen',
      'Potted Sauces':                'Potted Sauces',
      'Sauces':                       'Sauces',
      'Packaging':                    'Packaging',
      'Drinks':                       'Drinks',
      'Dry Goods':                    'Dry Goods',
      'Dessert Toppings':             'Dessert Toppings',
      'Dessert Mix':                  'Dessert Mix',
      'Cleaning & Kitchen Equipment': 'Cleaning & Kitchen',
      'BOH':                          'Kitchen Equipment',
      'Other':                        'Other',
    };

    const siteData = {};
    for (const site of Object.keys(SITE_COLS)) {
      siteData[site] = { prep: {}, stock: {} };
    }

    for (let i = 3; i < rows.length; i++) {
      const row  = rows[i];
      const name = (row[0] || '').toString().trim();
      if (!name || name.startsWith('KEY')) continue;

      const unit      = (row[2] || '').toString().trim();
      const active    = (row[5] || '').toString().trim().toLowerCase();
      const category  = (row[7] || '').toString().trim();
      const orderType = (row[8] || '').toString().trim();

      if (active !== 'yes') continue;
      if (!orderType) continue;

      const displayCat = CAT_DISPLAY[category] || category || 'Other';
      const otLower    = orderType.toLowerCase();
      const inPrep     = otLower === 'prep'  || otLower === 'both';
      const inStock    = otLower === 'stock' || otLower === 'both';

      for (const [site, colIdx] of Object.entries(SITE_COLS)) {
        const siteVal = (row[colIdx] || '').toString().trim().toLowerCase();
        if (siteVal !== 'yes') continue;

        if (inPrep) {
          if (!siteData[site].prep[displayCat]) siteData[site].prep[displayCat] = [];
          siteData[site].prep[displayCat].push({ name, unit });
        }
        if (inStock) {
          if (!siteData[site].stock[displayCat]) siteData[site].stock[displayCat] = [];
          siteData[site].stock[displayCat].push({ name, unit });
        }
      }
    }

    const callback = e && e.parameter && e.parameter.callback;
    const json     = JSON.stringify({ ok: true, data: siteData });
    if (callback) {
      return ContentService.createTextOutput(`${callback}(${json})`).setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    const callback = e && e.parameter && e.parameter.callback;
    const errJson  = JSON.stringify({ ok: false, error: err.toString() });
    if (callback) {
      return ContentService.createTextOutput(`${callback}(${errJson})`).setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(errJson).setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Dashboard handler — per-site, custom date range, full history ────
function handleDashboardGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const days   = parseInt(params.days || '7', 10);
    const site   = (params.site || '').trim();
    const now    = new Date();
    let fromDate = null, toDate = null;
    if (params.from) fromDate = parseFlexDate(params.from);
    if (params.to)   { toDate = parseFlexDate(params.to); if (toDate) toDate.setHours(23,59,59,999); }
    if (!fromDate && days > 0) fromDate = new Date(now.getTime() - days * 24 * 60 * 60 * 1000);
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sumWs = ss.getSheetByName('Orders Summary');
    const logWs = ss.getSheetByName('Order Log');
    const sumData = sumWs ? sumWs.getDataRange().getValues() : [];
    let totalOrders = 0, totalItems = 0, totalValue = 0;
    const bySiteMap = {};
    ALL_SITES.forEach(s => { bySiteMap[s] = { orders:0, items:0, value:0 }; });
    const byMonth = {}, orders = [], failures = [];

    for (let i = 1; i < sumData.length; i++) {
      const row     = sumData[i];
      const rowSite = (row[1] || '').toString().trim();
      if (site && rowSite !== site) continue;
      const rowDate = parseDDMMYYYY((row[10] || '').toString());
      if (fromDate && rowDate && rowDate < fromDate) continue;
      if (toDate   && rowDate && rowDate > toDate)   continue;

      const orderId   = (row[0]  || '').toString();
      const orderType = (row[2]  || '').toString();
      const submitted = (row[3]  || '').toString();
      const delivDate = (row[4]  || '').toString();
      const items     = parseInt(row[5])   || 0;
      const value     = parseFloat(row[6]) || 0;
      const prepTg    = (row[7]  || '').toString();
      const stockTg   = (row[8]  || '').toString();
      const monthYear = (row[11] || '').toString();
      const hasFail   = prepTg.includes('❌') || stockTg.includes('❌');

      totalOrders++; totalItems += items; totalValue += value;
      if (bySiteMap[rowSite]) { bySiteMap[rowSite].orders++; bySiteMap[rowSite].items += items; bySiteMap[rowSite].value += value; }
      if (!byMonth[monthYear]) byMonth[monthYear] = { month: monthYear, orders:0, items:0, value:0 };
      byMonth[monthYear].orders++; byMonth[monthYear].items += items; byMonth[monthYear].value += value;
      orders.push({ orderId, site:rowSite, submitted, delivDate, type:orderType, items, value:Math.round(value*100)/100, prepTg, stockTg, hasFail });
      if (hasFail) failures.push({ orderId, site:rowSite, submitted, delivDate, prepTg, stockTg });
    }
    orders.sort((a,b) => b.submitted.localeCompare(a.submitted));

    const logData = logWs ? logWs.getDataRange().getValues() : [];
    const itemTotals = {};
    for (let i = 1; i < logData.length; i++) {
      const row     = logData[i];
      const rowSite = (row[1] || '').toString().trim();
      if (site && rowSite !== site) continue;
      const rowDate = parseDDMMYYYY((row[12] || '').toString());
      if (fromDate && rowDate && rowDate < fromDate) continue;
      if (toDate   && rowDate && rowDate > toDate)   continue;
      const name = (row[2] || '').toString().trim();
      const qty  = parseFloat(row[4]) || 0;
      if (!name || qty <= 0) continue;
      itemTotals[name] = (itemTotals[name] || 0) + qty;
    }
    const topItems = Object.entries(itemTotals).sort((a,b) => b[1]-a[1]).slice(0,15)
      .map(([name,qty]) => ({ name, qty: Math.round(qty*10)/10 }));

    // Supplier breakdown from Order Log
    const supplierMap2 = {};
    for (let i = 1; i < logData.length; i++) {
      const row     = logData[i];
      const rowSite = (row[1] || '').toString().trim();
      if (site && rowSite !== site) continue;
      const rowDate = parseDDMMYYYY((row[12] || '').toString());
      if (fromDate && rowDate && rowDate < fromDate) continue;
      if (toDate   && rowDate && rowDate > toDate)   continue;
      const sup   = (row[5]  || '').toString().trim() || 'Unknown';
      const total = parseFloat(row[7]) || 0;
      if (!supplierMap2[sup]) supplierMap2[sup] = { supplier: sup, rows: 0, spend: 0 };
      supplierMap2[sup].rows++;
      supplierMap2[sup].spend += total;
    }
    const grandSpend = Object.values(supplierMap2).reduce((s,x) => s + x.spend, 0);
    const bySupplier = Object.values(supplierMap2)
      .sort((a,b) => b.spend - a.spend)
      .map(s => ({ supplier: s.supplier, orders: s.rows,
                   spend: Math.round(s.spend*100)/100,
                   pct:   grandSpend > 0 ? Math.round((s.spend/grandSpend)*1000)/10 : 0 }));

    // Per-site with avg per order + % of total
    const grandValue = Object.values(bySiteMap).reduce((s,x) => s + x.value, 0);
    const bySite = ALL_SITES.filter(s => bySiteMap[s].orders > 0)
      .map(s => ({
        site:   s,
        orders: bySiteMap[s].orders,
        items:  bySiteMap[s].items,
        value:  Math.round(bySiteMap[s].value*100)/100,
        avg:    bySiteMap[s].orders > 0 ? Math.round((bySiteMap[s].value/bySiteMap[s].orders)*100)/100 : 0,
        pct:    grandValue > 0 ? Math.round((bySiteMap[s].value/grandValue)*1000)/10 : 0
      }))
      .sort((a,b) => b.value - a.value);

    const mOrd = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    const mParse = m => { const p = m.split('-'); return parseInt(p[1])*12 + mOrd.indexOf(p[0]); };
    const byMonthArr = Object.values(byMonth).sort((a,b) => mParse(b.month)-mParse(a.month))
      .map(m => ({ ...m, value:Math.round(m.value*100)/100 }));

    // Add vs prior month
    for (let i = 0; i < byMonthArr.length; i++) {
      const prev = byMonthArr[i + 1];
      byMonthArr[i].vsPrior = prev != null ? Math.round((byMonthArr[i].value - prev.value)*100)/100 : null;
    }

    // Recent credits
    const credWs2   = ss.getSheetByName('Credits');
    const credData  = credWs2 ? credWs2.getDataRange().getValues() : [];
    const credits   = [];
    for (let i = credData.length - 1; i >= 1 && credits.length < 50; i--) {
      const r = credData[i];
      credits.push({ time: r[0], site: r[1], orderRef: r[2], item: r[3],
                     qty: r[4], unit: r[5], price: r[6], total: r[7], reason: r[8] });
    }

    const json = JSON.stringify({
      ok:true, site:site||'all', days,
      stats:{ orders:totalOrders, items:totalItems, value:Math.round(totalValue*100)/100,
              avg: totalOrders ? Math.round((totalItems/totalOrders)*10)/10 : 0 },
      bySite, byMonth:byMonthArr, bySupplier,
      orders:orders.slice(0,150), totalCount:orders.length,
      failures, topItems, credits
    });
    const cb = params.callback;
    if (cb) return ContentService.createTextOutput(cb+'('+json+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    const json = JSON.stringify({ ok:false, error:err.toString() });
    const cb   = e && e.parameter && e.parameter.callback;
    if (cb) return ContentService.createTextOutput(cb+'('+json+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }
}

function parseFlexDate(str) {
  if (!str) return null;
  str = str.toString().trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    const [y,m,d] = str.split('-');
    return new Date(parseInt(y), parseInt(m)-1, parseInt(d));
  }
  return parseDDMMYYYY(str);
}

// ── Summary handler — last 7 days, per site, per order ──────────────
function handleSummaryGet(e) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const logWs = ss.getSheetByName('Order Log');
    const data  = logWs.getDataRange().getValues();

    const now     = new Date();
    const weekAgo = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);

    // siteOrders[site][orderId] = { orderId, time, delivDate, items, prepOk, stockOk, hasFailed }
    const siteOrders = {};
    ALL_SITES.forEach(s => { siteOrders[s] = {}; });

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      // Col M (index 12) = dd/MM/yyyy date string
      const rowDate = parseDDMMYYYY((row[12] || '').toString());
      if (!rowDate || rowDate < weekAgo) continue;

      const site      = (row[1]  || '').toString().trim();
      const orderId   = (row[11] || '').toString().trim();
      const timeStr   = (row[0]  || '').toString().trim();
      const delivDate = (row[9]  || '').toString().trim();
      const tgStatus  = (row[10] || '').toString();

      if (!siteOrders[site]) continue;

      if (!siteOrders[site][orderId]) {
        siteOrders[site][orderId] = {
          orderId, time: timeStr, delivDate,
          items: 0, prepOk: null, stockOk: null, hasFailed: false
        };
      }

      const o = siteOrders[site][orderId];
      o.items++;

      // Parse Telegram status per item — consistent across items in same order
      if (tgStatus.includes('✅ Prep'))  o.prepOk  = true;
      if (tgStatus.includes('❌ Prep'))  { o.prepOk  = false; o.hasFailed = true; }
      if (tgStatus.includes('✅ Stock')) o.stockOk = true;
      if (tgStatus.includes('❌ Stock')) { o.stockOk = false; o.hasFailed = true; }
    }

    // Convert to sorted arrays — most recent first
    const result = {};
    ALL_SITES.forEach(site => {
      const orders = Object.values(siteOrders[site]);
      if (!orders.length) return;
      orders.sort((a, b) => b.time.localeCompare(a.time));
      result[site] = orders;
    });

    const weekStart = Utilities.formatDate(weekAgo, Session.getScriptTimeZone(), 'dd MMM');
    const weekEnd   = Utilities.formatDate(now,     Session.getScriptTimeZone(), 'dd MMM yyyy');
    const json      = JSON.stringify({ ok: true, period: weekStart + ' – ' + weekEnd, sites: result });

    const callback = e && e.parameter && e.parameter.callback;
    if (callback) return ContentService.createTextOutput(`${callback}(${json})`).setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    const json     = JSON.stringify({ ok: false, error: err.toString() });
    const callback = e && e.parameter && e.parameter.callback;
    if (callback) return ContentService.createTextOutput(`${callback}(${json})`).setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }
}

function parseDDMMYYYY(str) {
  const parts = str.split('/');
  if (parts.length !== 3) return null;
  return new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
}

// ════════════════════════════════════════════════════════════════════
// doPost — handles order submissions
//
// KEY CHANGE: Telegram is sent FIRST so the delivery status is known
// before any row is written to the log. Every logged row gets an
// accurate "Telegram Sent" value — no placeholder, no retroactive update.
// ════════════════════════════════════════════════════════════════════
function doPost(e) {
  try {
    let payload;
    try {
      if (e.postData.type === 'application/json') {
        payload = JSON.parse(e.postData.contents);
      } else {
        const raw = decodeURIComponent(e.postData.contents.replace(/^data=/, ''));
        payload = JSON.parse(raw);
      }
    } catch(parseErr) {
      payload = JSON.parse(e.postData.contents);
    }

    if ((payload.action || '') === 'record_credit') return handleCreditPost(payload);

    const site      = (payload.site      || '').trim();
    const orderType = (payload.orderType || '').trim();
    const notes     = (payload.notes     || '').trim();
    const rawDate   = (payload.deliveryDate || '').trim();
    const delivDate = rawDate ? rawDate.split('-').reverse().join('/') : '';
    const items     = payload.items || [];

    if (!site)         return jsonResponse({ ok: false, error: 'Missing site' });
    if (!orderType)    return jsonResponse({ ok: false, error: 'Missing orderType' });
    if (!items.length) return jsonResponse({ ok: false, error: 'No items' });

    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const priceWs = ss.getSheetByName('Price List');
    const logWs   = ss.getSheetByName('Order Log');
    const siteWs  = getOrCreateSiteSheet(ss, site);

    // Sequential order ID
    const props   = PropertiesService.getScriptProperties();
    const lastNum = parseInt(props.getProperty('ORDER_COUNTER') || '0', 10);
    const nextNum = lastNum + 1;
    props.setProperty('ORDER_COUNTER', nextNum.toString());
    const orderId = 'ORD-' + nextNum.toString().padStart(4, '0');

    const stamp     = new Date();
    const timeStr   = Utilities.formatDate(stamp, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    const dateOnly  = Utilities.formatDate(stamp, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const monthYear = Utilities.formatDate(stamp, Session.getScriptTimeZone(), 'MMM-yyyy');

    // Build lookup maps from Price List (live)
    const prRows = priceWs.getDataRange().getValues();
    const priceMap = {}, supplierMap = {}, unitMap = {}, orderTypeMap = {};
    for (let i = 3; i < prRows.length; i++) {
      const n = (prRows[i][0] || '').toString().trim();
      if (!n || n.startsWith('KEY')) continue;
      priceMap[n]     = parseFloat(prRows[i][4]) || 0;
      supplierMap[n]  = prRows[i][3] || '';
      unitMap[n]      = prRows[i][2] || '';
      orderTypeMap[n] = (prRows[i][8] || '').toString().trim();
    }

    // Build categorised item arrays
    const prepItems = [], stockItems = [], allItems = [];
    items.forEach(item => {
      const name = (item.name || '').trim();
      const qty  = parseFloat(item.qty) || 0;
      if (!name || qty <= 0 || qty > 999) return;

      const price    = priceMap[name] || item.price || 0;
      const supplier = supplierMap[name] || '';
      const unit     = unitMap[name]     || item.unit || '';
      const total    = Math.round(price * qty * 100) / 100;
      const ot       = (orderTypeMap[name] || item.section || 'stock').toLowerCase();

      allItems.push({ name, unit, qty, price, supplier, total, ot });
      if (ot === 'prep'  || ot === 'both') prepItems.push({ name, unit, qty });
      if (ot === 'stock' || ot === 'both') stockItems.push({ name, unit, qty });
    });

    if (!allItems.length) return jsonResponse({ ok: false, error: 'No valid items to log' });

    // ── Send Telegram FIRST — status known before any row is written ──
    let prepOk = null, stockOk = null;
    if (prepItems.length > 0) {
      prepOk = sendTelegram(PREP_GROUP_ID, site, prepItems, notes, delivDate, orderId, timeStr, 'PREP');
    }
    if (stockItems.length > 0) {
      stockOk = sendTelegram(STOCK_GROUP_ID, site, stockItems, notes, delivDate, orderId, timeStr, 'STOCK');
    }

    // ── Log to master Order Log + per-site sheet ──────────────────────
    allItems.forEach(item => {
      const tgStatus = buildTelegramStatus(item.ot, prepOk, stockOk);
      const row = [
        timeStr, site, item.name, item.unit, item.qty, item.supplier,
        item.price, item.total, notes, delivDate,
        tgStatus, orderId, dateOnly, monthYear
      ];
      logWs.appendRow(row);
      siteWs.appendRow(row);
    });

    // ── One summary row per order (used by Looker Studio + HTML summary) ──
    const summaryWs  = getOrCreateSummarySheet(ss);
    const totalValue = allItems.reduce((sum, i) => sum + i.total, 0);
    summaryWs.appendRow([
      orderId, site, orderType, timeStr, delivDate,
      allItems.length,
      Math.round(totalValue * 100) / 100,
      prepOk  === true ? '✅ Sent' : prepOk  === false ? '❌ Failed' : '—',
      stockOk === true ? '✅ Sent' : stockOk === false ? '❌ Failed' : '—',
      notes, dateOnly, monthYear
    ]);

    return jsonResponse({
      ok: true,
      orderId,
      itemsLogged: allItems.length,
      telegram: { prep: prepOk, stock: stockOk }
    });

  } catch(err) {
    Logger.log('doPost error: ' + err);
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function buildTelegramStatus(ot, prepOk, stockOk) {
  const parts = [];
  if (ot === 'prep'  || ot === 'both') parts.push(prepOk  === true ? '✅ Prep'  : '❌ Prep');
  if (ot === 'stock' || ot === 'both') parts.push(stockOk === true ? '✅ Stock' : '❌ Stock');
  return parts.join(' | ');
}

// Creates the site sheet if it doesn't exist yet, with headers + formatting
function getOrCreateSiteSheet(ss, site) {
  let sheet = ss.getSheetByName(site);
  if (!sheet) {
    sheet = ss.insertSheet(site);
    sheet.appendRow(LOG_HEADERS);
    sheet.getRange(1, 1, 1, LOG_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Creates the Orders Summary sheet (one row per order, not per item)
function getOrCreateSummarySheet(ss) {
  let sheet = ss.getSheetByName('Orders Summary');
  if (!sheet) {
    sheet = ss.insertSheet('Orders Summary');
    sheet.appendRow(SUMMARY_HEADERS);
    sheet.getRange(1, 1, 1, SUMMARY_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#e8f0fe');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ════════════════════════════════════════════════════════════════════
// TELEGRAM SENDER
// ════════════════════════════════════════════════════════════════════
function sendTelegram(chatId, site, items, notes, deliveryDate, orderId, timeStr, label) {
  let msg  = `${label} ORDER\n`;
  msg     += `📍 ${site}\n`;
  msg     += `Ref: ${orderId} | ${timeStr}\n`;
  if (deliveryDate) msg += `🗓 Delivery: ${deliveryDate}\n`;
  msg     += `─────────────────────\n`;
  items.forEach(it => { msg += `${it.name} — ${it.qty}x (${it.unit})\n`; });
  msg     += `─────────────────────`;
  if (notes) msg += `\nNotes: ${notes}`;

  const url = `https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`;
  try {
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, text: msg }),
      muteHttpExceptions: true
    });
    const code = response.getResponseCode();
    if (code !== 200) {
      Logger.log('Telegram FAILED to ' + chatId + ' (' + label + '): HTTP ' + code + ' — ' + response.getContentText());
      return false;
    }
    Logger.log('Telegram OK to ' + chatId + ' (' + label + ')');
    return true;
  } catch(err) {
    Logger.log('Telegram EXCEPTION to ' + chatId + ' (' + label + '): ' + err);
    return false;
  }
}

// ════════════════════════════════════════════════════════════════════
// MONTHLY ARCHIVE
//
// Runs automatically on the 1st of each month at 2am (set up by
// setupTrigger). For every site that has a sheet, it creates an
// archive tab named "SD Withington — Apr-2026" containing that
// month's rows. The live site sheet is left untouched (full history).
// ════════════════════════════════════════════════════════════════════
function createMonthlyArchives() {
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const now        = new Date();
  const prevMonth  = new Date(now.getFullYear(), now.getMonth() - 1, 1);
  const monthLabel = Utilities.formatDate(prevMonth, Session.getScriptTimeZone(), 'MMM-yyyy');

  ALL_SITES.forEach(site => {
    const siteSheet = ss.getSheetByName(site);
    if (!siteSheet) return;

    const archiveName = site + ' — ' + monthLabel;
    if (ss.getSheetByName(archiveName)) return; // Already archived this month

    const allData  = siteSheet.getDataRange().getValues();
    if (allData.length <= 1) return; // Headers only, nothing to archive

    // Col N (index 13) = Month-Year e.g. "Apr-2026"
    const monthRows = allData.slice(1).filter(row => row[13] === monthLabel);
    if (!monthRows.length) return;

    const archive = ss.insertSheet(archiveName);
    archive.appendRow(LOG_HEADERS);
    monthRows.forEach(row => archive.appendRow(row));

    archive.getRange(1, 1, 1, LOG_HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#f3f3f3');
    archive.setFrozenRows(1);
    archive.autoResizeColumns(1, LOG_HEADERS.length);
  });

  Logger.log('Monthly archives created for ' + monthLabel);
}

// ════════════════════════════════════════════════════════════════════
// CREDIT RECORDER — logs items that couldn't be delivered
// ════════════════════════════════════════════════════════════════════
function handleCreditPost(payload) {
  try {
    const site     = (payload.site     || '').trim();
    const orderRef = (payload.orderRef || '').trim();
    const itemName = (payload.itemName || '').trim();
    const qty      = parseFloat(payload.qty)   || 0;
    const unit     = (payload.unit     || '').trim();
    const price    = parseFloat(payload.price) || 0;
    const reason   = (payload.reason   || 'Out of stock').trim();

    if (!site || !itemName || qty <= 0) {
      return jsonResponse({ ok: false, error: 'Missing required fields (site, item, qty).' });
    }

    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const stamp    = new Date();
    const timeStr  = Utilities.formatDate(stamp, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    const dateOnly = Utilities.formatDate(stamp, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const monthYr  = Utilities.formatDate(stamp, Session.getScriptTimeZone(), 'MMM-yyyy');
    const total    = Math.round(price * qty * 100) / 100;

    let credWs = ss.getSheetByName('Credits');
    if (!credWs) {
      credWs = ss.insertSheet('Credits');
      credWs.appendRow(['Timestamp','Site','Order Ref','Item Name','Qty','Unit','Price (£)','Total (£)','Reason','Date','Month-Year']);
      credWs.getRange(1,1,1,11).setFontWeight('bold').setBackground('#FFF3CD');
      credWs.setFrozenRows(1);
    }

    credWs.appendRow([timeStr, site, orderRef, itemName, qty, unit, price, total, reason, dateOnly, monthYr]);
    return jsonResponse({ ok: true });

  } catch(err) {
    Logger.log('handleCreditPost error: ' + err);
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

// ════════════════════════════════════════════════════════════════════
// PRICE LOGGER — auto-logs every price change in Price List col E
// ════════════════════════════════════════════════════════════════════
function onPriceEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== 'Price List') return;
  if (e.range.getColumn() > 5 || e.range.getColumn() + e.range.getNumColumns() - 1 < 5) return;
  if (e.range.getRow() < 4) return;

  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const histWs = ss.getSheetByName('Price History');
  const now    = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

  const numRows  = e.range.getNumRows();
  const startRow = e.range.getRow();

  for (let i = 0; i < numRows; i++) {
    const row  = startRow + i;
    if (row < 4) continue;
    const name = sheet.getRange(row, 1).getValue();
    if (!name || name.toString().startsWith('KEY')) continue;

    const newVal = parseFloat(sheet.getRange(row, 5).getValue()) || 0;
    const oldVal = (numRows === 1 && e.oldValue) ? parseFloat(e.oldValue) : null;
    const change = oldVal !== null ? Math.round((newVal - oldVal) * 10000) / 10000 : null;
    const pct    = (oldVal !== null && oldVal) ? ((change / oldVal) * 100).toFixed(2) + '%' : 'N/A';

    histWs.appendRow([
      name,
      oldVal !== null ? oldVal : '(previous unknown)',
      newVal,
      change !== null ? change : '(bulk edit)',
      pct,
      now
    ]);
  }
}

// ════════════════════════════════════════════════════════════════════
// HELPER
// ════════════════════════════════════════════════════════════════════
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════════
// SETUP — run ONCE after pasting this script and deploying as Web App
// ════════════════════════════════════════════════════════════════════
function setupTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Price edit trigger
  ScriptApp.newTrigger('onPriceEdit').forSpreadsheet(ss).onEdit().create();

  // Monthly archive — runs at 2am on the 1st of every month
  ScriptApp.newTrigger('createMonthlyArchives')
    .timeBased()
    .onMonthDay(1)
    .atHour(2)
    .create();

  Logger.log('✅ Triggers ready. Deploy as Web App for the form to work.');
}
