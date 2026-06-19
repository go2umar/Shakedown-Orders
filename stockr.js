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
  if (action === 'summary')      return handleSummaryGet(e);
  if (action === 'dashboard')    return handleDashboardGet(e);
  if (action === 'get_orders')   return handleGetOrders(e);
  if (action === 'recent_order') return handleRecentOrderGet(e);
  if (action === 'get_pin')      return handleGetPin(e);
  return handleProductsGet(e);
}

// ── Products handler (original logic) ---------------------──────────
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

    const siteData = {};
    for (const site of Object.keys(SITE_COLS)) {
      siteData[site] = { prep: {}, stock: {} };
    }

    for (let i = 3; i < rows.length; i++) {
      const row  = rows[i];
      const name = (row[0] || '').toString().trim();
      if (!name || name.startsWith('KEY')) continue;

      const pkgType   = (row[1] || '').toString().trim(); // Col B = Pkg Type (packaging description)
      const unit      = (row[2] || '').toString().trim(); // Col C = Unit (used for quantities)
      const active    = (row[5] || '').toString().trim().toLowerCase();
      const category  = (row[7] || '').toString().trim();
      const orderType = (row[8] || '').toString().trim();

      if (active !== 'yes') continue;
      if (!orderType) continue;

      const displayCat = category || 'Other';
      const otLower    = orderType.toLowerCase();
      const inPrep     = otLower === 'prep'  || otLower === 'both';
      const inStock    = otLower === 'stock' || otLower === 'both';

      for (const [site, colIdx] of Object.entries(SITE_COLS)) {
        const siteVal = (row[colIdx] || '').toString().trim().toLowerCase();
        if (siteVal !== 'yes') continue;

        if (inPrep) {
          if (!siteData[site].prep[displayCat]) siteData[site].prep[displayCat] = [];
          siteData[site].prep[displayCat].push({ name, unit, pkgType });
        }
        if (inStock) {
          if (!siteData[site].stock[displayCat]) siteData[site].stock[displayCat] = [];
          siteData[site].stock[displayCat].push({ name, unit, pkgType });
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
    const byMonth = {}, byMonthBySite = {}, orders = [], failures = [];

    for (let i = 1; i < sumData.length; i++) {
      const row     = sumData[i];
      // Skip corrupt/header rows — valid orders always start with ORD-
      if (!/^ORD-\d+$/.test((row[0] || '').toString().trim())) continue;
      const rowSite = (row[1] || '').toString().trim();
      if (site && rowSite !== site) continue;
      const rawRowDate = row[10];
      const rowDate = rawRowDate instanceof Date
        ? new Date(rawRowDate.getFullYear(), rawRowDate.getMonth(), rawRowDate.getDate())
        : parseDDMMYYYY((rawRowDate || '').toString().trim());
      if (fromDate && rowDate && rowDate < fromDate) continue;
      if (toDate   && rowDate && rowDate > toDate)   continue;

      const orderId   = (row[0]  || '').toString();
      const orderType = (row[2]  || '').toString();
      const submitted = (row[3]  || '').toString();
      const rawD4     = row[4];
      const delivDate = rawD4 instanceof Date ? Utilities.formatDate(rawD4, Session.getScriptTimeZone(), 'dd/MM/yyyy') : (rawD4||'').toString().trim();
      const items     = parseInt(row[5])   || 0;
      const value     = parseFloat(row[6]) || 0;
      const prepTg    = (row[7]  || '').toString();
      const stockTg   = (row[8]  || '').toString();
      const rawMonth  = row[11];
      const monthYear = rawMonth instanceof Date
        ? Utilities.formatDate(rawMonth, Session.getScriptTimeZone(), 'MMM-yyyy')
        : (rawMonth || '').toString().trim();
      const hasFail   = prepTg.includes('❌') || stockTg.includes('❌');

      totalOrders++; totalItems += items; totalValue += value;
      if (bySiteMap[rowSite]) { bySiteMap[rowSite].orders++; bySiteMap[rowSite].items += items; bySiteMap[rowSite].value += value; }
      if (!byMonth[monthYear]) byMonth[monthYear] = { month: monthYear, orders:0, items:0, value:0 };
      byMonth[monthYear].orders++; byMonth[monthYear].items += items; byMonth[monthYear].value += value;
      // Monthly by site (for per-store trend chart)
      if (!byMonthBySite[monthYear]) byMonthBySite[monthYear] = {};
      byMonthBySite[monthYear][rowSite] = Math.round(((byMonthBySite[monthYear][rowSite]||0) + value)*100)/100;
      orders.push({ orderId, site:rowSite, submitted, delivDate, type:orderType, items, value:Math.round(value*100)/100, prepTg, stockTg, hasFail });
      if (hasFail) failures.push({ orderId, site:rowSite, submitted, delivDate, prepTg, stockTg });
    }
    orders.sort((a,b) => b.submitted.localeCompare(a.submitted));

    // Use the same order IDs as Orders Summary — guarantees itemBreakdown
    // totals match exactly regardless of date format or credit timing.
    const filteredOrderIds = new Set(orders.map(o => o.orderId));

    // Parse col M (index 12) for the supplier loop (still needs date filtering)
    const parseLogDate = raw => raw instanceof Date
      ? new Date(raw.getFullYear(), raw.getMonth(), raw.getDate())
      : parseDDMMYYYY((raw || '').toString().trim());

    const logData = logWs ? logWs.getDataRange().getValues() : [];
    const itemTotals = {}, itemSpend = {};
    for (let i = 1; i < logData.length; i++) {
      const row     = logData[i];
      const orderId = (row[11] || '').toString().trim();
      if (!filteredOrderIds.has(orderId)) continue;
      const name  = (row[2] || '').toString().trim();
      const qty   = parseFloat(row[4]) || 0;
      const total = parseFloat(row[7]) || 0;
      if (!name || qty <= 0) continue;
      itemTotals[name] = (itemTotals[name] || 0) + qty;
      itemSpend[name]  = Math.round(((itemSpend[name] || 0) + total) * 100) / 100;
    }
    // Subtract batch credits for orders in the filtered set.
    // Recall credits are excluded — they already reduced Order Log qty to 0.
    const credWsDash = ss.getSheetByName('Credits');
    if (credWsDash) {
      const credRowsDash = credWsDash.getDataRange().getValues();
      for (let i = 1; i < credRowsDash.length; i++) {
        const r = credRowsDash[i];
        const reason  = (r[8] || '').toString().trim();
        if (reason === 'Order Recalled') continue;
        const credOid = (r[2] || '').toString().trim();
        if (!filteredOrderIds.has(credOid)) continue;
        const cName  = (r[3] || '').toString().trim();
        const cTotal = parseFloat(r[7]) || 0;
        const cQty   = parseFloat(r[4]) || 0;
        if (!cName || cTotal <= 0) continue;
        if (itemSpend[cName]  !== undefined) itemSpend[cName]  = Math.round((itemSpend[cName]  - cTotal) * 100) / 100;
        if (itemTotals[cName] !== undefined) itemTotals[cName] = Math.round((itemTotals[cName] - cQty)   * 100) / 100;
      }
    }

    // Build VAT and category lookup from Price List (col J = index 9, col H = index 7)
    const vatLookup = {}, catLookupDash = {};
    const priceWsDash = ss.getSheetByName('Price List');
    if (priceWsDash) {
      const prRowsDash = priceWsDash.getDataRange().getValues();
      for (let i = 3; i < prRowsDash.length; i++) {
        const n = (prRowsDash[i][0] || '').toString().trim();
        if (!n || n.startsWith('KEY')) continue;
        vatLookup[n]     = (prRowsDash[i][9] || '').toString().trim();
        catLookupDash[n] = (prRowsDash[i][7] || '').toString().trim() || 'Other';
      }
    }
    // Apply manual overrides (takes precedence over Price List)
    const overrideWsDash = ss.getSheetByName('Item Overrides');
    if (overrideWsDash) {
      const overrideRows = overrideWsDash.getDataRange().getValues();
      for (let i = 1; i < overrideRows.length; i++) {
        const n = (overrideRows[i][0] || '').toString().trim();
        if (!n) continue;
        const oCat = (overrideRows[i][1] || '').toString().trim();
        const oVat = (overrideRows[i][2] || '').toString().trim();
        if (oCat) catLookupDash[n] = oCat;
        if (oVat) vatLookup[n]     = oVat;
      }
    }

    // All items in period with VAT and category — for the cost breakdown table
    const itemBreakdown = Object.entries(itemTotals)
      .sort((a,b) => a[0].localeCompare(b[0]))
      .map(([name, qty]) => {
        const spend    = Math.round((itemSpend[name] || 0) * 100) / 100;
        const vat      = vatLookup[name] || '';
        const category = catLookupDash[name] || 'Other';
        const totalCost = vat.includes('20%') ? Math.round(spend * 1.2 * 100) / 100 : spend;
        return { name, qty: Math.round(qty * 10) / 10, spend, vat, totalCost, category };
      });

    // Include both qty and spend so the frontend can sort either way
    const topItems = Object.entries(itemTotals).sort((a,b) => b[1]-a[1]).slice(0,50)
      .map(([name,qty]) => ({
        name,
        qty:   Math.round(qty * 10) / 10,
        spend: Math.round((itemSpend[name] || 0) * 100) / 100
      }));

    // Supplier breakdown from Order Log
    const supplierMap2 = {};
    for (let i = 1; i < logData.length; i++) {
      const row     = logData[i];
      const rowSite = (row[1] || '').toString().trim();
      if (site && rowSite !== site) continue;
      const rowDate = parseLogDate(row[12]);
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

    // Unpriced items — Order Log rows where Price = 0 and name is not empty
    const unpriced = [];
    for (let i = 1; i < logData.length; i++) {
      const row      = logData[i];
      const rowSite  = (row[1] || '').toString().trim();
      if (site && rowSite !== site) continue;
      const itemName = (row[2] || '').toString().trim();
      const price    = parseFloat(row[6]) || 0;
      const ordId    = (row[11] || '').toString().trim();
      if (!itemName || price > 0) continue;
      // Only include items with no supplier (indicates custom/unlisted)
      const supplier = (row[5] || '').toString().trim();
      if (supplier) continue;
      const rawD12  = row[12];
      const dateStr = rawD12 instanceof Date
        ? Utilities.formatDate(rawD12, Session.getScriptTimeZone(), 'dd/MM/yyyy')
        : (rawD12 || '').toString();
      unpriced.push({ orderId: ordId, site: rowSite, date: dateStr,
                      item: itemName, qty: parseFloat(row[4])||0, unit: (row[3]||'').toString() });
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

    // All unique categories from Price List + overrides, in standard order then alpha, Other last
    const allCatSet = new Set(Object.values(catLookupDash));
    allCatSet.add('Other');
    const catOrderBase = ['Raw','Sauces','Potted Sauces','Fresh','Frozen','BOH','Packaging'];
    const categories = [
      ...catOrderBase.filter(c => allCatSet.has(c)),
      ...[...allCatSet].filter(c => !catOrderBase.includes(c) && c !== 'Other').sort(),
      'Other'
    ];

    const json = JSON.stringify({
      ok:true, site:site||'all', days,
      stats:{ orders:totalOrders, items:totalItems, value:Math.round(totalValue*100)/100,
              avg: totalOrders ? Math.round((totalItems/totalOrders)*10)/10 : 0 },
      bySite, byMonth:byMonthArr, bySupplier,
      orders:orders.slice(0,150), totalCount:orders.length,
      failures, topItems, credits, unpriced, byMonthBySite, itemBreakdown, categories
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

// ── Recent order — returns most recent order within last 30 min ──────
// Used by the HTML for cross-device recall bar on page load.
function handleRecentOrderGet(e) {
  try {
    const ss   = SpreadsheetApp.getActiveSpreadsheet();
    const tgWs = ss.getSheetByName('TG Messages');
    const cb   = e && e.parameter && e.parameter.callback;

    const resp = obj => {
      const json = JSON.stringify(obj);
      if (cb) return ContentService.createTextOutput(cb+'('+json+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
      return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
    };

    if (!tgWs) return resp({ ok: true, order: null });

    const data   = tgWs.getDataRange().getValues();
    const now    = new Date();
    const cutoff = new Date(now.getTime() - 30 * 60 * 1000);

    const orders = [];
    for (let i = 1; i < data.length; i++) {
      const orderId = (data[i][0]||'').toString().trim();
      const site    = (data[i][1]||'').toString().trim();
      const sentAt  = parseDDMMYYYYHHMM((data[i][4]||'').toString().trim());
      if (!orderId || !sentAt || sentAt < cutoff) continue;
      orders.push({ orderId, site, submittedAt: sentAt.getTime() });
    }
    // Newest first
    orders.sort((a, b) => b.submittedAt - a.submittedAt);

    return resp({ ok: true, orders });
  } catch(err) {
    const json = JSON.stringify({ ok: false, error: err.toString() });
    const cb   = e && e.parameter && e.parameter.callback;
    if (cb) return ContentService.createTextOutput(cb+'('+json+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }
}

function parseDDMMYYYYHHMM(str) {
  if (!str) return null;
  const parts = str.split(' ');
  if (parts.length < 2) return null;
  const [d, m, y] = parts[0].split('/');
  const [h, min]  = parts[1].split(':');
  if (!y || !m || !d || !h || !min) return null;
  return new Date(parseInt(y), parseInt(m)-1, parseInt(d), parseInt(h), parseInt(min));
}

// ── Order lookup — search by Order ID or by site + date ─────────────
function handleGetOrders(e) {
  try {
    const params  = (e && e.parameter) || {};
    const orderId = (params.orderId || '').trim();
    const site    = (params.site    || '').trim();
    const date    = (params.date    || '').trim(); // DD/MM/YYYY

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const logWs = ss.getSheetByName('Order Log');
    const sumWs = ss.getSheetByName('Orders Summary');
    const cb    = params.callback;

    const resp = obj => {
      const json = JSON.stringify(obj);
      if (cb) return ContentService.createTextOutput(cb+'('+json+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
      return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
    };

    if (!logWs) return resp({ ok: false, error: 'Order Log not found.' });

    // Search by Order ID — return every item in that order
    if (orderId) {
      const logData = logWs.getDataRange().getValues();
      const items   = [];
      for (let i = 1; i < logData.length; i++) {
        const row = logData[i];
        if ((row[11]||'').toString().trim() !== orderId) continue;
        const rawD9  = row[9];
        const rawD0 = row[0];
        const submittedFmt = rawD0 instanceof Date
          ? Utilities.formatDate(rawD0, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
          : (rawD0 || '').toString().trim();
        items.push({
          name:      (row[2] ||'').toString().trim(),
          unit:      (row[3] ||'').toString().trim(),
          qty:       parseFloat(row[4]) || 0,
          supplier:  (row[5] ||'').toString().trim(),
          price:     parseFloat(row[6]) || 0,
          total:     parseFloat(row[7]) || 0,
          tgStatus:  (row[10]||'').toString().trim(),
          site:      (row[1] ||'').toString().trim(),
          submitted: submittedFmt,
          delivDate: rawD9 instanceof Date ? Utilities.formatDate(rawD9, Session.getScriptTimeZone(), 'dd/MM/yyyy') : (rawD9||'').toString().trim()
        });
      }
      // Check Credits sheet — mark items that have been partially or fully credited.
      // Key by "name|price" so credits at the old price don't bleed onto rows
      // of the same item added later at a different price.
      // Credits are consumed row-by-row (in Order Log order) so that a re-added
      // item of the same name/price is not marked credited by the original credit.
      const credWs = ss.getSheetByName('Credits');
      if (credWs) {
        const credData  = credWs.getDataRange().getValues();
        const creditMap = {};
        for (let i = 1; i < credData.length; i++) {
          const r   = credData[i];
          const ref = (r[2] || '').toString().trim();
          if (ref !== orderId) continue;
          const n     = (r[3] || '').toString().trim();
          const price = (parseFloat(r[6]) || 0).toFixed(4);
          const q     = parseFloat(r[4]) || 0;
          const key   = n + '|' + price;
          creditMap[key] = (creditMap[key] || 0) + q;
        }
        // Consume credits in the same order the items appear in the log
        // so re-added rows don't inherit credits meant for the original row
        const remaining = Object.assign({}, creditMap);
        items.forEach(item => {
          const key   = item.name + '|' + (parseFloat(item.price) || 0).toFixed(4);
          const avail = remaining[key] || 0;
          const used  = Math.min(avail, item.qty);
          remaining[key]     = Math.max(0, avail - item.qty);
          item.creditedQty   = Math.round(used * 10) / 10;
          item.fullyCredited = used >= item.qty;
        });
      }

      // Mark removed items (qty=0) so the manager can see what was removed
      // but staff in recall-edit mode can filter them out client-side
      const orderSite = items[0] ? items[0].site : '';
      items.forEach(item => { if (item.qty === 0) item.removed = true; });
      return resp({ ok: true, mode: 'items', orderId, site: orderSite, items });
    }

    // Search by site + date — return list of matching orders
    if (!sumWs) return resp({ ok: false, error: 'Orders Summary not found. Run migrateHistoricalData() first.' });
    const sumData = sumWs.getDataRange().getValues();
    const orders  = [];
    for (let i = 1; i < sumData.length; i++) {
      const row     = sumData[i];
      const rowSite     = (row[1]||'').toString().trim();
      // Filter by delivery date (col E, index 4) — not the submission date
      const rawDelivDate = row[4];
      const rowDelivDate = rawDelivDate instanceof Date
        ? Utilities.formatDate(rawDelivDate, Session.getScriptTimeZone(), 'dd/MM/yyyy')
        : (rawDelivDate||'').toString().trim();
      if (site && rowSite !== site) continue;
      if (date && rowDelivDate !== date) continue;
      orders.push({
        orderId:   (row[0]||'').toString().trim(),
        site:      rowSite,
        type:      (row[2]||'').toString().trim(),
        submitted: (row[3]||'').toString().trim(),
        delivDate: (row[4] instanceof Date ? Utilities.formatDate(row[4], Session.getScriptTimeZone(), 'dd/MM/yyyy') : (row[4]||'').toString().trim()),
        items:     parseInt(row[5])   || 0,
        value:     parseFloat(row[6]) || 0,
        prepTg:    (row[7]||'').toString().trim(),
        stockTg:   (row[8]||'').toString().trim(),
        date:      rowDelivDate
      });
    }
    orders.sort((a, b) => b.submitted.localeCompare(a.submitted));
    return resp({ ok: true, mode: 'orders', orders });

  } catch(err) {
    const json = JSON.stringify({ ok: false, error: err.toString() });
    const cb   = e && e.parameter && e.parameter.callback;
    if (cb) return ContentService.createTextOutput(cb+'('+json+')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }
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

    if ((payload.action || '') === 'set_pin')               return handleSetPin(payload);
    if ((payload.action || '') === 'record_credit')        return handleCreditPost(payload);
    if ((payload.action || '') === 'batch_credit')         return handleBatchCredit(payload);
    if ((payload.action || '') === 'change_qty')           return handleChangeQty(payload);
    if ((payload.action || '') === 'reverse_credit')       return handleReverseCredit(payload);
    if ((payload.action || '') === 'set_item_price')       return handleSetItemPrice(payload);
    if ((payload.action || '') === 'add_to_order')         return handleAddToOrder(payload);
    if ((payload.action || '') === 'recall_order')         return handleRecallOrder(payload);
    if ((payload.action || '') === 'change_delivery_date') return handleChangeDeliveryDate(payload);
    if ((payload.action || '') === 'save_item_override')   return handleSaveItemOverride(payload);

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

    // Build lookup maps from Price List (active items only)
    const prRows = priceWs.getDataRange().getValues();
    const priceMap = {}, supplierMap = {}, unitMap = {}, orderTypeMap = {}, categoryMap = {};
    const inactiveNames = new Set(); // names that exist in Price List but are inactive
    for (let i = 3; i < prRows.length; i++) {
      const n = (prRows[i][0] || '').toString().trim();
      if (!n || n.startsWith('KEY')) continue;
      const activeFlag = (prRows[i][5] || '').toString().trim().toLowerCase();
      if (activeFlag !== 'yes') { inactiveNames.add(n); continue; }
      priceMap[n]     = parseFloat(prRows[i][4]) || 0;
      supplierMap[n]  = prRows[i][3] || '';
      unitMap[n]      = prRows[i][2] || '';
      orderTypeMap[n] = (prRows[i][8] || '').toString().trim();
      categoryMap[n]  = (prRows[i][7] || '').toString().trim() || 'Other';
    }

    // Build categorised item arrays
    const prepItems = [], stockItems = [], allItems = [];
    items.forEach(item => {
      const name = (item.name || '').trim();
      const qty  = parseFloat(item.qty) || 0;
      if (!name || qty <= 0 || qty > 999) return;
      // Reject items that exist in the Price List but are marked inactive
      if (inactiveNames.has(name) && !priceMap[name]) return;

      const price    = priceMap[name] || item.price || 0;
      const supplier = supplierMap[name] || '';
      const unit     = unitMap[name]     || item.unit || '';
      const total    = Math.round(price * qty * 100) / 100;
      const ot       = (orderTypeMap[name] || item.section || 'stock').toLowerCase();
      const category = categoryMap[name]  || 'Other';
      const note     = (item.note || '').toString().trim().slice(0, 120);

      allItems.push({ name, unit, qty, price, supplier, total, ot, category, note });
      if (ot === 'prep'  || ot === 'both') prepItems.push({ name, unit, qty, category, note });
      if (ot === 'stock' || ot === 'both') stockItems.push({ name, unit, qty, category, note });
    });

    if (!allItems.length) return jsonResponse({ ok: false, error: 'No valid items to log' });

    // ── Send Telegram FIRST — status known before any row is written ──
    let prepOk = null, stockOk = null, prepMsgId = null, stockMsgId = null;
    if (prepItems.length > 0) {
      const r = sendTelegram(PREP_GROUP_ID, site, prepItems, notes, delivDate, orderId, timeStr, 'PREP');
      prepOk = r.ok; prepMsgId = r.messageId;
    }
    if (stockItems.length > 0) {
      const r = sendTelegram(STOCK_GROUP_ID, site, stockItems, notes, delivDate, orderId, timeStr, 'STOCK');
      stockOk = r.ok; stockMsgId = r.messageId;
    }
    // Store message IDs so orders can be recalled within 30 min
    storeTelegramMsgIds(ss, orderId, site, prepMsgId, stockMsgId, timeStr);

    // ── Batch write to Order Log + site sheet (one setValues call each) ──
    // Much faster than appendRow per item — reduces N sheet ops to 2.
    const logRows = allItems.map(item => {
      const tgStatus = buildTelegramStatus(item.ot, prepOk, stockOk);
      return [
        timeStr, site, item.name, item.unit, item.qty, item.supplier,
        item.price, item.total, notes, delivDate,
        tgStatus, orderId, dateOnly, monthYear
      ];
    });
    const cols = LOG_HEADERS.length;
    logWs.getRange(logWs.getLastRow() + 1, 1, logRows.length, cols).setValues(logRows);
    siteWs.getRange(siteWs.getLastRow() + 1, 1, logRows.length, cols).setValues(logRows);

    // ── One summary row per order ---------------------─────────────────
    const summaryWs  = getOrCreateSummarySheet(ss);
    const totalValue = allItems.reduce((sum, i) => sum + i.total, 0);
    const sumRow     = [[
      orderId, site, orderType, timeStr, delivDate,
      allItems.length,
      Math.round(totalValue * 100) / 100,
      prepOk  === true ? '✅ Sent' : prepOk  === false ? '❌ Failed' : '—',
      stockOk === true ? '✅ Sent' : stockOk === false ? '❌ Failed' : '—',
      notes, dateOnly, monthYear
    ]];
    summaryWs.getRange(summaryWs.getLastRow() + 1, 1, 1, SUMMARY_HEADERS.length).setValues(sumRow);

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
function pluraliseUnit(unit, qty) {
  if (!unit) return '';
  const u = unit.trim();
  const n = parseFloat(qty) || 0;
  // Units that never pluralise
  if (/^(kg|g|ml|cl|l|ltr|ltrs|litre|litres|each|ea|%)$/i.test(u)) return u;
  // Already plural or qty is exactly 1
  if (n === 1) return u;
  const bl = u.toLowerCase();
  // Words ending in ch, sh, ss, x, z → +es
  if (/ch$|sh$|ss$|[xz]$/.test(bl)) return u + 'es';
  // Words ending in consonant+y → remove y, +ies (rarely applies to units but handle it)
  if (/[^aeiou]y$/i.test(u)) return u.slice(0, -1) + 'ies';
  // Default: +s
  return u + 's';
}

function buildTelegramText(site, items, notes, deliveryDate, orderId, timeStr, label) {
  const sortItems = arr => arr.slice().sort((a, b) => {
    if (site.startsWith('DC')) {
      const aClub = a.name.toLowerCase().startsWith('club');
      const bClub = b.name.toLowerCase().startsWith('club');
      if (aClub !== bClub) return aClub ? -1 : 1;
    }
    return a.name.localeCompare(b.name);
  });
  // *bold* markers are WhatsApp-compatible: plain text in Telegram, bold when pasted into WhatsApp
  let msg = `*📌 ${site.toUpperCase()}*\n`;
  if (label.includes('ADDITION')) msg += `*+ ADDITION*\n`;
  msg += `Ref: ${orderId} | ${timeStr}\n`;
  if (deliveryDate) msg += `Delivery: ${deliveryDate}\n`;
  msg += `---------------------\n`;
  const hasCategories = items.some(i => i.category);
  if (hasCategories) {
    const catMap = {};
    items.forEach(it => { const cat = it.category || 'Other'; if (!catMap[cat]) catMap[cat] = []; catMap[cat].push(it); });
    const CAT_ORDER = ['Raw', 'Sauces', 'Potted Sauces', 'Fresh', 'Frozen', 'BOH'];
    const sortedCats = Object.keys(catMap).sort((a, b) => {
      const ai = CAT_ORDER.indexOf(a), bi = CAT_ORDER.indexOf(b);
      if (ai !== -1 && bi !== -1) return ai - bi;
      if (ai !== -1) return -1; if (bi !== -1) return 1;
      return a.localeCompare(b);
    });
    sortedCats.forEach((cat, idx) => {
      if (idx > 0) msg += `\n`;
      msg += `*${cat}*\n`;
      sortItems(catMap[cat]).forEach(it => {
        const q = it.qty % 1 === 0 ? Math.round(it.qty) : it.qty;
        msg += `• ${it.name}  —  ${q} ${pluraliseUnit(it.unit, it.qty)}${it.note ? `  (${it.note})` : ''}${it.creditAnnotation ? `  *${it.creditAnnotation}*` : ''}\n`;
      });
    });
  } else {
    sortItems(items).forEach(it => {
      const q = it.qty % 1 === 0 ? Math.round(it.qty) : it.qty;
      msg += `• ${it.name}  —  ${q} ${pluraliseUnit(it.unit, it.qty)}${it.note ? `  (${it.note})` : ''}\n`;
    });
  }
  msg += `---------------------`;
  if (notes) msg += `\n📝 Notes: ${notes}`;
  return msg;
}

function sendTelegramText(chatId, text) {
  try {
    const res  = UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/sendMessage`, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, text }),
      muteHttpExceptions: true
    });
    const body = JSON.parse(res.getContentText());
    return { ok: body.ok === true, messageId: body.result ? body.result.message_id : null };
  } catch(e) { Logger.log('sendTelegramText error: ' + e); return { ok: false, messageId: null }; }
}

function editTelegramMessage(chatId, messageId, text) {
  try {
    const res  = UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/editMessageText`, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, message_id: parseInt(messageId), text }),
      muteHttpExceptions: true
    });
    const body = JSON.parse(res.getContentText());
    return body.ok === true;
  } catch(e) { Logger.log('editTelegramMessage error: ' + e); return false; }
}

function sendTelegram(chatId, site, items, notes, deliveryDate, orderId, timeStr, label) {
  const msg    = buildTelegramText(site, items, notes, deliveryDate, orderId, timeStr, label);
  const result = sendTelegramText(chatId, msg);
  Logger.log('Telegram ' + (result.ok ? 'OK' : 'FAILED') + ' to ' + chatId + ' (' + label + ') msgId=' + result.messageId);
  return result;
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
// ADD ITEM TO EXISTING ORDER
// Appends a new line to Order Log + site sheet, updates Orders Summary,
// and edits the existing Telegram message to include the new item.
// Falls back to sending a new message if the original can't be edited.
// ════════════════════════════════════════════════════════════════════
function handleAddToOrder(payload) {
  try {
    const orderId     = (payload.orderId  || '').trim();
    const site        = (payload.site     || '').trim();
    const itemName    = (payload.itemName || '').trim();
    const qty         = parseFloat(payload.qty) || 0;
    const manualPrice = parseFloat(payload.price) || 0;

    if (!orderId || !site || !itemName || qty <= 0) {
      return jsonResponse({ ok: false, error: 'Missing required fields.' });
    }

    const ss      = SpreadsheetApp.getActiveSpreadsheet();
    const priceWs = ss.getSheetByName('Price List');
    const logWs   = ss.getSheetByName('Order Log');
    const sumWs   = ss.getSheetByName('Orders Summary');

    // Look up item from Price List (active rows only)
    const prRows = priceWs.getDataRange().getValues();
    let price = manualPrice, supplier = '', unit = '', orderType = 'stock', category = 'Other';
    for (let i = 3; i < prRows.length; i++) {
      const n = (prRows[i][0] || '').toString().trim();
      if (n !== itemName) continue;
      const activeFlag = (prRows[i][5] || '').toString().trim().toLowerCase();
      if (activeFlag !== 'yes') continue;
      price     = parseFloat(prRows[i][4]) || manualPrice;
      supplier  = (prRows[i][3] || '').toString();
      unit      = (prRows[i][2] || '').toString().trim();
      orderType = (prRows[i][8] || '').toString().trim().toLowerCase();
      category  = (prRows[i][7] || '').toString().trim() || 'Other';
      break;
    }
    const total = Math.round(price * qty * 100) / 100;

    // Build category lookup for all items in this order
    const catLookup = {};
    for (let i = 3; i < prRows.length; i++) {
      const n = (prRows[i][0] || '').toString().trim();
      if (n) catLookup[n] = (prRows[i][7] || '').toString().trim() || 'Other';
    }

    // Build credit map for this order so credited items are excluded from Telegram
    // Key: "name|price.toFixed(4)" — matches the crediting logic in handleGetOrders
    // Mutable copy so credits are consumed row-by-row — prevents the same
    // credit amount from being applied to both an original row AND a
    // re-added row of the same item at the same price.
    const remainingCredits = Object.assign({}, buildCreditMap(ss, orderId));

    // Find original order's meta + read all existing items from Order Log
    const logData = logWs.getDataRange().getValues();
    let origTime = '', origDeliv = '', origNotes = '', origDate = '', origMonth = '';
    const prepItems = [], stockItems = [];
    for (let i = 1; i < logData.length; i++) {
      const row = logData[i];
      if ((row[11] || '').toString().trim() !== orderId) continue;
      if (!origTime) {
        origTime  = fmtTimestamp(row[0]);
        origNotes = (row[8]  || '').toString();
        const rd  = row[9];
        origDeliv = rd instanceof Date ? Utilities.formatDate(rd, Session.getScriptTimeZone(), 'dd/MM/yyyy') : (rd||'').toString();
        const rd2 = row[12];
        origDate  = rd2 instanceof Date ? Utilities.formatDate(rd2, Session.getScriptTimeZone(), 'dd/MM/yyyy') : (rd2||'').toString();
        const rd3 = row[13];
        origMonth = rd3 instanceof Date ? Utilities.formatDate(rd3, Session.getScriptTimeZone(), 'MMM-yyyy') : (rd3||'').toString();
      }
      const qty2 = parseFloat(row[4]) || 0;
      if (qty2 <= 0) continue; // removed during recall
      const name2  = (row[2] || '').toString().trim();
      const unit2  = (row[3] || '').toString().trim();
      const price2 = parseFloat(row[6]) || 0;
      const key2   = name2 + '|' + price2.toFixed(4);
      const avail  = remainingCredits[key2] || 0;
      const creditedQty = Math.min(avail, qty2);
      remainingCredits[key2] = Math.max(0, avail - qty2); // consume
      if (creditedQty >= qty2) continue; // fully credited — exclude from Telegram
      const activeQty = Math.round((qty2 - creditedQty) * 100) / 100;
      const tg    = (row[10] || '').toString();
      const item2 = { name: name2, unit: unit2, qty: activeQty, category: catLookup[name2] || 'Other' };
      if (tg.includes('Prep') && tg.includes('Stock')) { prepItems.push(item2); stockItems.push(item2); }
      else if (tg.includes('Prep'))  prepItems.push(item2);
      else                           stockItems.push(item2);
    }

    // Add new item to the relevant list(s) so the full updated message can be built
    const newItem = { name: itemName, unit, qty, category };
    const ot = orderType.toLowerCase();
    if (ot === 'prep'  || ot === 'both') prepItems.push(newItem);
    if (ot === 'stock' || ot === 'both') stockItems.push(newItem);

    // Get existing Telegram message IDs
    let oldPrepMsgId = null, oldStockMsgId = null;
    const tgWs = ss.getSheetByName('TG Messages');
    if (tgWs) {
      const tgData = tgWs.getDataRange().getValues();
      for (let i = 1; i < tgData.length; i++) {
        if ((tgData[i][0] || '').toString().trim() !== orderId) continue;
        oldPrepMsgId  = tgData[i][2] ? String(tgData[i][2]) : null;
        oldStockMsgId = tgData[i][3] ? String(tgData[i][3]) : null;
        break;
      }
    }

    // Edit existing message or fall back to sending a new one
    const addedTimeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    let prepOk = null, stockOk = null;
    let newPrepMsgId = oldPrepMsgId, newStockMsgId = oldStockMsgId;

    const editOrSend = (chatId, oldMsgId, items, label) => {
      const text = buildTelegramText(site, items, origNotes, origDeliv, orderId, addedTimeStr, label);
      if (oldMsgId) {
        // Always try to delete + resend so Telegram fires a new notification.
        // Only fall back to silent edit if deletion fails (message too old etc).
        const deleted = deleteTelegramMessage(chatId, oldMsgId);
        if (deleted) return sendTelegramText(chatId, text);
        const edited = editTelegramMessage(chatId, oldMsgId, text);
        return { ok: edited, messageId: edited ? parseInt(oldMsgId) : null };
      }
      return sendTelegramText(chatId, text);
    };

    if (prepItems.length > 0 && (ot === 'prep' || ot === 'both')) {
      const r = editOrSend(PREP_GROUP_ID, oldPrepMsgId, prepItems, 'PREP');
      prepOk = r.ok;
      if (r.messageId) newPrepMsgId = String(r.messageId);
    }
    if (stockItems.length > 0 && (ot === 'stock' || ot === 'both')) {
      const r = editOrSend(STOCK_GROUP_ID, oldStockMsgId, stockItems, 'STOCK');
      stockOk = r.ok;
      if (r.messageId) newStockMsgId = String(r.messageId);
    }

    // Update TG Messages if any message ID changed
    if (newPrepMsgId !== oldPrepMsgId || newStockMsgId !== oldStockMsgId) {
      storeTelegramMsgIds(ss, orderId, site, newPrepMsgId, newStockMsgId, addedTimeStr);
    }

    const tgStatus = buildTelegramStatus(orderType, prepOk, stockOk);

    // Append new item to Order Log and site sheet
    const siteWs = ss.getSheetByName(site);
    const newRow = [origTime, site, itemName, unit, qty, supplier, price, total, origNotes, origDeliv, tgStatus + ' (addition)', orderId, origDate, origMonth];
    logWs.appendRow(newRow);
    if (siteWs) siteWs.appendRow(newRow);

    // Update Orders Summary
    if (sumWs) {
      const sumData = sumWs.getDataRange().getValues();
      for (let i = 1; i < sumData.length; i++) {
        if ((sumData[i][0] || '').toString().trim() !== orderId) continue;
        sumWs.getRange(i + 1, 6).setValue((parseInt(sumData[i][5]) || 0) + 1);
        sumWs.getRange(i + 1, 7).setValue(Math.round(((parseFloat(sumData[i][6]) || 0) + total) * 100) / 100);
        break;
      }
    }

    return jsonResponse({ ok: true, price, total });
  } catch(err) {
    Logger.log('handleAddToOrder error: ' + err);
    return jsonResponse({ ok: false, error: err.toString() });
  }
}


// ════════════════════════════════════════════════════════════════════
// RECALL ORDER — deletes original Telegram message(s) and resends
// a complete updated order. Only works within 30 minutes.
// ════════════════════════════════════════════════════════════════════
function handleRecallOrder(payload) {
  try {
    const orderId     = (payload.orderId || '').trim();
    const site        = (payload.site    || '').trim();
    const modItems    = payload.modifiedItems || null; // [{name, unit, qty, tg}]
    if (!orderId || !site) return jsonResponse({ ok: false, error: 'Missing orderId or site.' });

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const logWs = ss.getSheetByName('Order Log');
    const tgWs  = ss.getSheetByName('TG Messages');

    // Build category lookup from Price List (active items only)
    const recallCatLookup = {};
    const priceWsR = ss.getSheetByName('Price List');
    if (priceWsR) {
      const prRowsR = priceWsR.getDataRange().getValues();
      for (let i = 3; i < prRowsR.length; i++) {
        const n = (prRowsR[i][0] || '').toString().trim();
        if (!n) continue;
        const activeFlag = (prRowsR[i][5] || '').toString().trim().toLowerCase();
        if (activeFlag !== 'yes') continue;
        recallCatLookup[n] = (prRowsR[i][7] || '').toString().trim() || 'Other';
      }
    }

    // Collect all current items for this order grouped by prep/stock
    const logData = logWs.getDataRange().getValues();
    const prepItems = [], stockItems = [], bothItems = [];
    let notes = '', delivDate = '', timeStr = '', origDate = '', origMonth = '';
    for (let i = 1; i < logData.length; i++) {
      const row = logData[i];
      if ((row[11] || '').toString().trim() !== orderId) continue;
      if (!notes)    notes    = (row[8] || '').toString();
      if (!timeStr)  timeStr  = (row[0] || '').toString();
      if (!delivDate) {
        const rd = row[9];
        delivDate = rd instanceof Date ? Utilities.formatDate(rd, Session.getScriptTimeZone(), 'dd/MM/yyyy') : (rd||'').toString();
      }
      if (!origDate) {
        const rd2 = row[12];
        origDate = rd2 instanceof Date ? Utilities.formatDate(rd2, Session.getScriptTimeZone(), 'dd/MM/yyyy') : (rd2||'').toString();
      }
      if (!origMonth) {
        const rd3 = row[13];
        origMonth = rd3 instanceof Date ? Utilities.formatDate(rd3, Session.getScriptTimeZone(), 'MMM-yyyy') : (rd3||'').toString();
      }
      const name = (row[2] || '').toString().trim();
      const unit = (row[3] || '').toString().trim();
      const qty  = parseFloat(row[4]) || 0;
      if (qty <= 0) continue; // skip removed items — don't include in Telegram
      const tg   = (row[10] || '').toString();
      const item = { name, unit, qty, category: recallCatLookup[name] || 'Other' };
      if (tg.includes('Prep') && tg.includes('Stock')) { prepItems.push(item); stockItems.push(item); }
      else if (tg.includes('Prep'))  prepItems.push(item);
      else                           stockItems.push(item);
    }

    // Read existing Telegram message IDs before replacing them
    let oldPrepMsgId = null, oldStockMsgId = null;
    if (tgWs) {
      const tgData = tgWs.getDataRange().getValues();
      for (let i = 1; i < tgData.length; i++) {
        if ((tgData[i][0] || '').toString().trim() !== orderId) continue;
        oldPrepMsgId  = tgData[i][2] ? String(tgData[i][2]) : null;
        oldStockMsgId = tgData[i][3] ? String(tgData[i][3]) : null;
        break;
      }
    }

    // modItems was explicitly provided (even as empty array) → use it as the new item list
    // modItems === null means no changes were sent → use original Order Log items
    let modMap = null; // declared here so it's accessible after the if block
    if (modItems !== null && modItems !== undefined) {
      prepItems.length = 0; stockItems.length = 0;
      modMap = {};
      modItems.forEach(m => { modMap[m.name] = parseFloat(m.qty) || 0; });
      modItems.forEach(item => {
        const mi = { name: item.name, unit: item.unit, qty: parseFloat(item.qty) || 0, category: recallCatLookup[item.name] || 'Other' };
        if (mi.qty <= 0) return; // removed items must not appear in the Telegram message
        const tg = (item.tg || '').toLowerCase();
        if (tg.includes('prep'))  prepItems.push(mi);
        if (tg.includes('stock')) stockItems.push(mi);
      });
      // Update Order Log: new qty for changed items, 0 for removed items
      // Auto-credit items whose qty was reduced or fully removed
      let credWs = ss.getSheetByName('Credits');
      if (!credWs) {
        credWs = ss.insertSheet('Credits');
        credWs.appendRow(['Timestamp','Site','Order Ref','Item Name','Qty','Unit','Price (£)','Total (£)','Reason','Date','Month-Year']);
        credWs.getRange(1,1,1,11).setFontWeight('bold').setBackground('#FFF3CD');
        credWs.setFrozenRows(1);
      }
      const creditStamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
      const creditDate  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
      const creditMonth = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM-yyyy');
      const logData2 = logWs.getDataRange().getValues();
      const existingNames = new Set();
      for (let i = 1; i < logData2.length; i++) {
        if ((logData2[i][11] || '').toString().trim() !== orderId) continue;
        const name    = (logData2[i][2] || '').toString().trim();
        const unit    = (logData2[i][3] || '').toString().trim();
        const origQty = parseFloat(logData2[i][4]) || 0;
        const price   = parseFloat(logData2[i][6]) || 0;
        existingNames.add(name);
        const newQty  = name in modMap ? modMap[name] : 0;
        logWs.getRange(i + 1, 5).setValue(newQty);
        logWs.getRange(i + 1, 8).setValue(Math.round(price * newQty * 100) / 100);
        const creditQty = Math.round((origQty - newQty) * 100) / 100;
        if (creditQty > 0) {
          credWs.appendRow([creditStamp, site, orderId, name, creditQty, unit, price, Math.round(price * creditQty * 100) / 100, 'Order Recalled', creditDate, creditMonth]);
        }
      }
      // Insert new items (not in original order) into Order Log and site sheet
      const priceWsNew = ss.getSheetByName('Price List');
      const prRowsNew  = priceWsNew ? priceWsNew.getDataRange().getValues() : [];
      const priceLu    = {};
      for (let i = 3; i < prRowsNew.length; i++) {
        const n = (prRowsNew[i][0] || '').toString().trim();
        if (!n) continue;
        const activeFlag = (prRowsNew[i][5] || '').toString().trim().toLowerCase();
        if (activeFlag !== 'yes') continue;
        priceLu[n] = { price: parseFloat(prRowsNew[i][4]) || 0, supplier: (prRowsNew[i][3] || '').toString(), unit: (prRowsNew[i][2] || '').toString().trim() };
      }
      const siteWsNew = ss.getSheetByName(site);
      modItems.forEach(item => {
        const name = (item.name || '').trim();
        if (!name || existingNames.has(name)) return;
        const qty = parseFloat(item.qty) || 0;
        if (qty <= 0) return;
        const pl       = priceLu[name] || {};
        const price    = pl.price    || 0;
        const unit     = pl.unit     || item.unit || '';
        const supplier = pl.supplier || '';
        const total    = Math.round(price * qty * 100) / 100;
        const newRow   = [creditStamp, site, name, unit, qty, supplier, price, total, notes, delivDate, (item.tg || '✅ Stock') + ' (recall addition)', orderId, origDate, origMonth];
        logWs.appendRow(newRow);
        if (siteWsNew) siteWsNew.appendRow(newRow);
      });
    }

    const newTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

    // Helper: delete old message, edit it if deletion fails, return new messageId
    const replaceOrEdit = (chatId, oldMsgId, newText) => {
      if (oldMsgId) {
        const deleted = deleteTelegramMessage(chatId, oldMsgId);
        if (deleted) {
          return sendTelegramText(chatId, newText);
        } else {
          const edited = editTelegramMessage(chatId, oldMsgId, newText);
          return { ok: edited, messageId: edited ? parseInt(oldMsgId) : null };
        }
      }
      return sendTelegramText(chatId, newText);
    };

    // If all items removed — delete original messages (edit with cancellation notice if undeletable)
    if (prepItems.length === 0 && stockItems.length === 0) {
      const cancelText = () => `*📌 ${site.toUpperCase()}*\n*↩ ORDER CANCELLED*\nRef: ${orderId}\nAll items removed.`;
      if (oldPrepMsgId) {
        const deleted = deleteTelegramMessage(PREP_GROUP_ID, oldPrepMsgId);
        if (!deleted) editTelegramMessage(PREP_GROUP_ID, oldPrepMsgId, cancelText());
      }
      if (oldStockMsgId) {
        const deleted = deleteTelegramMessage(STOCK_GROUP_ID, oldStockMsgId);
        if (!deleted) editTelegramMessage(STOCK_GROUP_ID, oldStockMsgId, cancelText());
      }
      storeTelegramMsgIds(ss, orderId, site, null, null, newTime);
      // Update Orders Summary to 0 items / £0 when entire order is cancelled
      const sumWsCancel = ss.getSheetByName('Orders Summary');
      if (sumWsCancel) {
        const sumData = sumWsCancel.getDataRange().getValues();
        for (let i = 1; i < sumData.length; i++) {
          if ((sumData[i][0] || '').toString().trim() !== orderId) continue;
          sumWsCancel.getRange(i + 1, 6).setValue(0);
          sumWsCancel.getRange(i + 1, 7).setValue(0);
          break;
        }
      }
      Logger.log('Recall: all items removed.');
      return jsonResponse({ ok: true });
    }

    // Build updated message texts and replace (delete+send) or edit in place
    let prepR  = { ok: false, messageId: null };
    let stockR = { ok: false, messageId: null };

    if (prepItems.length > 0) {
      const text = buildTelegramText(site, prepItems, notes, delivDate, orderId, newTime, 'PREP');
      prepR = replaceOrEdit(PREP_GROUP_ID, oldPrepMsgId, text);
    } else if (oldPrepMsgId) {
      deleteTelegramMessage(PREP_GROUP_ID, oldPrepMsgId);
    }

    if (stockItems.length > 0) {
      const text = buildTelegramText(site, stockItems, notes, delivDate, orderId, newTime, 'STOCK');
      stockR = replaceOrEdit(STOCK_GROUP_ID, oldStockMsgId, text);
    } else if (oldStockMsgId) {
      deleteTelegramMessage(STOCK_GROUP_ID, oldStockMsgId);
    }

    // Update stored message IDs
    storeTelegramMsgIds(ss, orderId, site, prepR.messageId, stockR.messageId, newTime);

    // ── Sync site sheet + Orders Summary when modifications were made ──
    if (modMap) {
      // Site sheet — mirror the Order Log changes
      const siteSheet = ss.getSheetByName(site);
      if (siteSheet) {
        const sData = siteSheet.getDataRange().getValues();
        for (let i = 1; i < sData.length; i++) {
          if ((sData[i][11] || '').toString().trim() !== orderId) continue;
          const name   = (sData[i][2] || '').toString().trim();
          const newQty = name in modMap ? modMap[name] : 0;
          const price  = parseFloat(sData[i][6]) || 0;
          siteSheet.getRange(i + 1, 5).setValue(newQty);
          siteSheet.getRange(i + 1, 8).setValue(Math.round(price * newQty * 100) / 100);
        }
      }

      // Orders Summary — recalculate items count and total value from updated Order Log
      const sumWs = ss.getSheetByName('Orders Summary');
      if (sumWs) {
        const freshLog = logWs.getDataRange().getValues();
        let newCount = 0, newValue = 0;
        for (let i = 1; i < freshLog.length; i++) {
          if ((freshLog[i][11] || '').toString().trim() !== orderId) continue;
          const qty   = parseFloat(freshLog[i][4]) || 0;
          const total = parseFloat(freshLog[i][7]) || 0;
          if (qty > 0) { newCount++; newValue += total; }
        }
        const sumData = sumWs.getDataRange().getValues();
        for (let i = 1; i < sumData.length; i++) {
          if ((sumData[i][0] || '').toString().trim() !== orderId) continue;
          sumWs.getRange(i + 1, 6).setValue(newCount);                       // Items
          sumWs.getRange(i + 1, 7).setValue(Math.round(newValue * 100) / 100); // Total Value (£)
          break;
        }
      }
    }

    return jsonResponse({ ok: true });
  } catch(err) {
    Logger.log('handleRecallOrder error: ' + err);
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

function deleteTelegramMessage(chatId, messageId) {
  try {
    const res  = UrlFetchApp.fetch(`https://api.telegram.org/bot${TELEGRAM_TOKEN}/deleteMessage`, {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ chat_id: chatId, message_id: parseInt(messageId) }),
      muteHttpExceptions: true
    });
    const body = JSON.parse(res.getContentText());
    return body.ok === true;
  } catch(e) { Logger.log('deleteTelegramMessage failed: ' + e); return false; }
}

function storeTelegramMsgIds(ss, orderId, site, prepMsgId, stockMsgId, timeStr) {
  let tgWs = ss.getSheetByName('TG Messages');
  if (!tgWs) {
    tgWs = ss.insertSheet('TG Messages');
    tgWs.appendRow(['Order ID','Site','Prep Msg ID','Stock Msg ID','Sent At']);
    tgWs.getRange(1,1,1,5).setFontWeight('bold').setBackground('#f3f3f3');
    tgWs.setFrozenRows(1);
  }
  // Remove old entry for this orderId if exists (we'll add fresh one)
  const data = tgWs.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if ((data[i][0]||'').toString().trim() === orderId) tgWs.deleteRow(i + 1);
  }
  tgWs.appendRow([orderId, site, prepMsgId||'', stockMsgId||'', timeStr]);
}

// ════════════════════════════════════════════════════════════════════
// SET ITEM PRICE — manager sets price for a custom/unpriced item
// Finds the row(s) in Order Log by Order ID + Item Name and updates
// Price (£) and Total (£) in place.
// ════════════════════════════════════════════════════════════════════
function handleSetItemPrice(payload) {
  try {
    const orderId  = (payload.orderId  || '').trim();
    const itemName = (payload.itemName || '').trim();
    const price    = parseFloat(payload.price) || 0;

    if (!orderId || !itemName) {
      return jsonResponse({ ok: false, error: 'Missing orderId or itemName.' });
    }

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const logWs = ss.getSheetByName('Order Log');
    const data  = logWs.getDataRange().getValues();
    let updated = 0;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if ((row[11] || '').toString().trim() === orderId &&
          (row[2]  || '').toString().trim() === itemName) {
        const qty   = parseFloat(row[4]) || 0;
        const total = Math.round(price * qty * 100) / 100;
        logWs.getRange(i + 1, 7).setValue(price); // Col G = Price (£)
        logWs.getRange(i + 1, 8).setValue(total); // Col H = Total (£)
        updated++;
      }
    }

    // Update the site-specific sheet too
    if (updated > 0) {
      const siteSheet = ss.getSheetByName(payload.site || '');
      if (siteSheet) {
        const sData = siteSheet.getDataRange().getValues();
        for (let i = 1; i < sData.length; i++) {
          const row = sData[i];
          if ((row[11] || '').toString().trim() === orderId &&
              (row[2]  || '').toString().trim() === itemName) {
            const qty   = parseFloat(row[4]) || 0;
            const total = Math.round(price * qty * 100) / 100;
            siteSheet.getRange(i + 1, 7).setValue(price);
            siteSheet.getRange(i + 1, 8).setValue(total);
          }
        }
      }

      // Recalculate Orders Summary total from the updated Order Log
      const sumWs = ss.getSheetByName('Orders Summary');
      if (sumWs) {
        const freshLog = logWs.getDataRange().getValues();
        let newValue = 0;
        for (let i = 1; i < freshLog.length; i++) {
          if ((freshLog[i][11] || '').toString().trim() !== orderId) continue;
          const qty   = parseFloat(freshLog[i][4]) || 0;
          const total = parseFloat(freshLog[i][7]) || 0;
          if (qty > 0) newValue += total;
        }
        const sumData = sumWs.getDataRange().getValues();
        for (let i = 1; i < sumData.length; i++) {
          if ((sumData[i][0] || '').toString().trim() !== orderId) continue;
          sumWs.getRange(i + 1, 7).setValue(Math.round(newValue * 100) / 100);
          break;
        }
      }
    }

    return jsonResponse({ ok: true, updated });

  } catch(err) {
    Logger.log('handleSetItemPrice error: ' + err);
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

// ════════════════════════════════════════════════════════════════════
// MANAGER PIN — stored in Script Properties so it syncs across devices
// ════════════════════════════════════════════════════════════════════
function handleGetPin(e) {
  const pin = PropertiesService.getScriptProperties().getProperty('manager_pin') || '1234';
  const cb  = e && e.parameter && e.parameter.callback;
  const json = JSON.stringify({ ok: true, pin });
  if (cb) return ContentService.createTextOutput(cb + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function handleSetPin(payload) {
  try {
    const currentPin = (payload.currentPin || '').toString().trim();
    const newPin     = (payload.newPin     || '').toString().trim();
    const stored     = PropertiesService.getScriptProperties().getProperty('manager_pin') || '1234';
    if (currentPin !== stored)      return jsonResponse({ ok: false, error: 'Incorrect current PIN' });
    if (!/^\d{4}$/.test(newPin))    return jsonResponse({ ok: false, error: 'PIN must be exactly 4 digits' });
    PropertiesService.getScriptProperties().setProperty('manager_pin', newPin);
    return jsonResponse({ ok: true });
  } catch(err) {
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

// ════════════════════════════════════════════════════════════════════
// BATCH CREDIT — logs multiple credits in one call, then rebuilds the
// Telegram message with credited items annotated inline.
// ════════════════════════════════════════════════════════════════════
function handleBatchCredit(payload) {
  try {
    const orderId = (payload.orderRef || '').trim();
    const site    = (payload.site     || '').trim();
    const credits = payload.credits   || [];

    if (!orderId || !site || !credits.length) {
      return jsonResponse({ ok: false, error: 'Missing required fields.' });
    }

    const ss       = SpreadsheetApp.getActiveSpreadsheet();
    const stamp    = new Date();
    const timeStr  = Utilities.formatDate(stamp, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    const dateOnly = Utilities.formatDate(stamp, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    const monthYr  = Utilities.formatDate(stamp, Session.getScriptTimeZone(), 'MMM-yyyy');

    // ── Log all credits to Credits sheet ─────────────────────────
    let credWs = ss.getSheetByName('Credits');
    if (!credWs) {
      credWs = ss.insertSheet('Credits');
      credWs.appendRow(['Timestamp','Site','Order ID','Item Name','Qty','Unit','Price (£)','Total (£)','Reason','Date','Month-Year']);
    }
    let totalCreditValue = 0;
    credits.forEach(c => {
      const qty   = parseFloat(c.qty)   || 0;
      const price = parseFloat(c.price) || 0;
      const total = Math.round(price * qty * 100) / 100;
      totalCreditValue += total;
      credWs.appendRow([timeStr, site, orderId, c.itemName, qty, c.unit, price, total, c.reason, dateOnly, monthYr]);
    });

    // ── Deduct credit total from Orders Summary ───────────────────
    const sumWs = ss.getSheetByName('Orders Summary');
    if (sumWs && totalCreditValue > 0) {
      const sumData = sumWs.getDataRange().getValues();
      for (let i = 1; i < sumData.length; i++) {
        if ((sumData[i][0] || '').toString().trim() !== orderId) continue;
        const current = parseFloat(sumData[i][6]) || 0;
        sumWs.getRange(i + 1, 7).setValue(Math.round(Math.max(0, current - totalCreditValue) * 100) / 100);
        break;
      }
    }

    // ── Build category lookup ─────────────────────────────────────
    const catLookup = {};
    const priceWs = ss.getSheetByName('Price List');
    if (priceWs) {
      const prRows = priceWs.getDataRange().getValues();
      for (let i = 3; i < prRows.length; i++) {
        const n = (prRows[i][0] || '').toString().trim();
        if (n) catLookup[n] = (prRows[i][7] || '').toString().trim() || 'Other';
      }
    }

    // ── Build credit + reason maps (all credits for this order) ──
    // Mutable copy consumed row-by-row so re-added items aren't double-credited
    const remainingCredits = Object.assign({}, buildCreditMap(ss, orderId));
    const reasonMap = {}; // name → most-recent reason
    const credData  = credWs.getDataRange().getValues();
    for (let i = 1; i < credData.length; i++) {
      if ((credData[i][2] || '').toString().trim() !== orderId) continue;
      reasonMap[(credData[i][3] || '').toString().trim()] = (credData[i][8] || '').toString().trim();
    }

    // ── Read all items from Order Log; annotate credited ones ─────
    const logWs = ss.getSheetByName('Order Log');
    if (!logWs) return jsonResponse({ ok: false, error: 'Order Log not found.' });
    const logData = logWs.getDataRange().getValues();
    let origTime = '', origDeliv = '', origNotes = '';
    const prepItems = [], stockItems = [];

    for (let i = 1; i < logData.length; i++) {
      const row = logData[i];
      if ((row[11] || '').toString().trim() !== orderId) continue;
      if (!origTime) {
        origTime  = fmtTimestamp(row[0]);
        origNotes = (row[8] || '').toString();
        const rd  = row[9];
        origDeliv = rd instanceof Date ? Utilities.formatDate(rd, Session.getScriptTimeZone(), 'dd/MM/yyyy') : (rd||'').toString();
      }
      const qty2 = parseFloat(row[4]) || 0;
      if (qty2 <= 0) continue; // removed during recall

      const name2  = (row[2]  || '').toString().trim();
      const unit2  = (row[3]  || '').toString().trim();
      const price2 = parseFloat(row[6]) || 0;
      const key2   = name2 + '|' + price2.toFixed(4);
      const avail  = remainingCredits[key2] || 0;
      const creditedQty = Math.min(avail, qty2);
      remainingCredits[key2] = Math.max(0, avail - qty2); // consume so re-added rows aren't affected
      const reason = reasonMap[name2] || '';
      const tg     = (row[10] || '').toString();

      // Build inline annotation for credited items
      let creditAnnotation = '';
      if (creditedQty >= qty2) {
        creditAnnotation = '✕ fully credited' + (reason ? ': ' + reason : '');
      } else if (creditedQty > 0) {
        creditAnnotation = '✕ ' + creditedQty + ' ' + unit2 + ' credited' + (reason ? ': ' + reason : '');
      }

      // Fully credited rows still appear in the message (with annotation); re-added rows show normally
      const displayQty = creditedQty >= qty2 ? qty2 : Math.max(0, Math.round((qty2 - creditedQty) * 100) / 100);
      const item2 = { name: name2, unit: unit2, qty: displayQty, category: catLookup[name2] || 'Other', creditAnnotation };
      if (tg.includes('Prep') && tg.includes('Stock')) { prepItems.push(item2); stockItems.push(item2); }
      else if (tg.includes('Prep'))  prepItems.push(item2);
      else                           stockItems.push(item2);
    }

    // ── Delete + resend Telegram ──────────────────────────────────
    let oldPrepMsgId = null, oldStockMsgId = null;
    const tgWs = ss.getSheetByName('TG Messages');
    if (tgWs) {
      const tgData = tgWs.getDataRange().getValues();
      for (let i = 1; i < tgData.length; i++) {
        if ((tgData[i][0] || '').toString().trim() !== orderId) continue;
        oldPrepMsgId  = tgData[i][2] ? String(tgData[i][2]) : null;
        oldStockMsgId = tgData[i][3] ? String(tgData[i][3]) : null;
        break;
      }
    }
    const updatedTimeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    let newPrepMsgId = oldPrepMsgId, newStockMsgId = oldStockMsgId;

    // Credits use edit-first (silent update, no new notification).
    // Fall back to delete+resend only if the edit fails, and to a fresh
    // send as a last resort so the group always gets the updated info.
    const editOrSend = (chatId, oldMsgId, items, label) => {
      const text = buildTelegramText(site, items, origNotes, origDeliv, orderId, origTime, label);
      if (oldMsgId) {
        const edited = editTelegramMessage(chatId, oldMsgId, text);
        if (edited) return { ok: true, messageId: parseInt(oldMsgId) };
        // Edit failed — delete old and resend
        const deleted = deleteTelegramMessage(chatId, oldMsgId);
        if (deleted) return sendTelegramText(chatId, text);
      }
      // No ID or all attempts failed — send new message
      return sendTelegramText(chatId, text);
    };

    if (prepItems.length > 0) {
      const r = editOrSend(PREP_GROUP_ID, oldPrepMsgId, prepItems, 'PREP');
      if (r.messageId) newPrepMsgId = String(r.messageId);
    }
    if (stockItems.length > 0) {
      const r = editOrSend(STOCK_GROUP_ID, oldStockMsgId, stockItems, 'STOCK');
      if (r.messageId) newStockMsgId = String(r.messageId);
    }
    if (newPrepMsgId !== oldPrepMsgId || newStockMsgId !== oldStockMsgId) {
      storeTelegramMsgIds(ss, orderId, site, newPrepMsgId, newStockMsgId, updatedTimeStr);
    }

    return jsonResponse({ ok: true, credited: credits.length });
  } catch(err) {
    Logger.log('handleBatchCredit error: ' + err);
    return jsonResponse({ ok: false, error: err.toString() });
  }
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
// REVERSE CREDIT — manager only
// Deletes all credit entries for an item in an order, restores the
// Orders Summary total, and rebuilds the Telegram message without
// the credit annotation.
// ════════════════════════════════════════════════════════════════════
function handleReverseCredit(payload) {
  try {
    const orderId  = (payload.orderRef || '').trim();
    const site     = (payload.site     || '').trim();
    const itemName = (payload.itemName || '').trim();
    const price    = parseFloat(payload.price) || 0;

    if (!orderId || !site || !itemName) {
      return jsonResponse({ ok: false, error: 'Missing required fields.' });
    }

    const ss     = SpreadsheetApp.getActiveSpreadsheet();
    const credWs = ss.getSheetByName('Credits');
    if (!credWs) return jsonResponse({ ok: false, error: 'Credits sheet not found.' });

    // Find and delete all matching credit rows (iterate bottom→top so indices stay valid)
    const credData = credWs.getDataRange().getValues();
    let reversedTotal = 0;
    for (let i = credData.length - 1; i >= 1; i--) {
      const r = credData[i];
      if ((r[2] || '').toString().trim() !== orderId)   continue;
      if ((r[3] || '').toString().trim() !== itemName)  continue;
      if (Math.abs((parseFloat(r[6]) || 0) - price) > 0.01) continue;
      reversedTotal += parseFloat(r[7]) || 0; // accumulate restored value
      credWs.deleteRow(i + 1);
    }

    // Add reversed value back to Orders Summary
    if (reversedTotal > 0) {
      const sumWs = ss.getSheetByName('Orders Summary');
      if (sumWs) {
        const sumData = sumWs.getDataRange().getValues();
        for (let i = 1; i < sumData.length; i++) {
          if ((sumData[i][0] || '').toString().trim() !== orderId) continue;
          const current = parseFloat(sumData[i][6]) || 0;
          sumWs.getRange(i + 1, 7).setValue(Math.round((current + reversedTotal) * 100) / 100);
          break;
        }
      }
    }

    // Rebuild Telegram (item now shows without credit annotation)
    const catLookup = {};
    const priceWs = ss.getSheetByName('Price List');
    if (priceWs) {
      const prRows = priceWs.getDataRange().getValues();
      for (let i = 3; i < prRows.length; i++) {
        const n = (prRows[i][0] || '').toString().trim();
        if (n) catLookup[n] = (prRows[i][7] || '').toString().trim() || 'Other';
      }
    }

    const remainingCredits = Object.assign({}, buildCreditMap(ss, orderId));
    const logWs = ss.getSheetByName('Order Log');
    if (!logWs) return jsonResponse({ ok: true }); // sheets updated, TG best-effort
    const logData = logWs.getDataRange().getValues();
    let origTime = '', origDeliv = '', origNotes = '';
    const prepItems = [], stockItems = [];

    for (let i = 1; i < logData.length; i++) {
      const row = logData[i];
      if ((row[11] || '').toString().trim() !== orderId) continue;
      if (!origTime) {
        origTime  = fmtTimestamp(row[0]);
        origNotes = (row[8] || '').toString();
        const rd  = row[9];
        origDeliv = rd instanceof Date ? Utilities.formatDate(rd, Session.getScriptTimeZone(), 'dd/MM/yyyy') : (rd||'').toString();
      }
      const qty2 = parseFloat(row[4]) || 0;
      if (qty2 <= 0) continue;
      const name2  = (row[2]  || '').toString().trim();
      const unit2  = (row[3]  || '').toString().trim();
      const price2 = parseFloat(row[6]) || 0;
      const key2   = name2 + '|' + price2.toFixed(4);
      const avail  = remainingCredits[key2] || 0;
      const credited = Math.min(avail, qty2);
      remainingCredits[key2] = Math.max(0, avail - qty2);
      if (credited >= qty2) continue;
      const activeQty = Math.round((qty2 - credited) * 100) / 100;
      const tg = (row[10] || '').toString();
      const item2 = { name: name2, unit: unit2, qty: activeQty, category: catLookup[name2] || 'Other' };
      if (tg.includes('Prep') && tg.includes('Stock')) { prepItems.push(item2); stockItems.push(item2); }
      else if (tg.includes('Prep')) prepItems.push(item2);
      else stockItems.push(item2);
    }

    let oldPrepMsgId = null, oldStockMsgId = null;
    const tgWs = ss.getSheetByName('TG Messages');
    if (tgWs) {
      const tgData = tgWs.getDataRange().getValues();
      for (let i = 1; i < tgData.length; i++) {
        if ((tgData[i][0] || '').toString().trim() !== orderId) continue;
        oldPrepMsgId  = tgData[i][2] ? String(tgData[i][2]) : null;
        oldStockMsgId = tgData[i][3] ? String(tgData[i][3]) : null;
        break;
      }
    }

    const updatedTimeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    let newPrepMsgId = oldPrepMsgId, newStockMsgId = oldStockMsgId;

    const editOrSend = (chatId, oldMsgId, items, label) => {
      const text = buildTelegramText(site, items, origNotes, origDeliv, orderId, origTime, label);
      if (oldMsgId) {
        const edited = editTelegramMessage(chatId, oldMsgId, text);
        if (edited) return { ok: true, messageId: parseInt(oldMsgId) };
        const deleted = deleteTelegramMessage(chatId, oldMsgId);
        if (deleted) return sendTelegramText(chatId, text);
      }
      return sendTelegramText(chatId, text);
    };

    if (prepItems.length > 0) {
      const r = editOrSend(PREP_GROUP_ID, oldPrepMsgId, prepItems, 'PREP');
      if (r.messageId) newPrepMsgId = String(r.messageId);
    }
    if (stockItems.length > 0) {
      const r = editOrSend(STOCK_GROUP_ID, oldStockMsgId, stockItems, 'STOCK');
      if (r.messageId) newStockMsgId = String(r.messageId);
    }
    if (newPrepMsgId !== oldPrepMsgId || newStockMsgId !== oldStockMsgId) {
      storeTelegramMsgIds(ss, orderId, site, newPrepMsgId, newStockMsgId, updatedTimeStr);
    }

    return jsonResponse({ ok: true, reversedTotal });
  } catch(err) {
    Logger.log('handleReverseCredit error: ' + err);
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

// ════════════════════════════════════════════════════════════════════
// CHANGE QTY — manager only
// Sets an item to a new quantity directly, adjusting the Order Log,
// site sheet, and Orders Summary. No credit entry is created — this
// is for correcting a wrong order qty, not recording a delivery issue.
// Telegram is rebuilt (delete+resend) so the group sees the correction.
// ════════════════════════════════════════════════════════════════════
function handleChangeQty(payload) {
  try {
    const orderId  = (payload.orderRef  || '').trim();
    const site     = (payload.site      || '').trim();
    const itemName = (payload.itemName  || '').trim();
    const newQty   = parseFloat(payload.newQty) || 0;
    const price    = parseFloat(payload.price)  || 0;

    if (!orderId || !site || !itemName || newQty <= 0) {
      return jsonResponse({ ok: false, error: 'Missing required fields.' });
    }

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const logWs = ss.getSheetByName('Order Log');
    if (!logWs) return jsonResponse({ ok: false, error: 'Order Log not found.' });

    // Find and update the row in Order Log
    const logData  = logWs.getDataRange().getValues();
    let origQty = -1, updatedSheetRow = -1;

    for (let i = 1; i < logData.length; i++) {
      const row = logData[i];
      if ((row[11] || '').toString().trim() !== orderId) continue;
      if ((row[2]  || '').toString().trim() !== itemName) continue;
      if (Math.abs((parseFloat(row[6]) || 0) - price) > 0.01) continue;
      origQty        = parseFloat(row[4]) || 0;
      updatedSheetRow = i + 1;
      break;
    }

    if (updatedSheetRow === -1) return jsonResponse({ ok: false, error: 'Item not found in order.' });

    const newTotal  = Math.round(price * newQty  * 100) / 100;
    const origTotal = Math.round(price * origQty * 100) / 100;
    const diff      = Math.round((newTotal - origTotal) * 100) / 100; // + means more cost

    logWs.getRange(updatedSheetRow, 5).setValue(newQty);
    logWs.getRange(updatedSheetRow, 8).setValue(newTotal);

    // Update site sheet
    const siteWs = ss.getSheetByName(site);
    if (siteWs) {
      const siteData = siteWs.getDataRange().getValues();
      for (let i = 1; i < siteData.length; i++) {
        if ((siteData[i][11] || '').toString().trim() !== orderId) continue;
        if ((siteData[i][2]  || '').toString().trim() !== itemName) continue;
        if (Math.abs((parseFloat(siteData[i][6]) || 0) - price) > 0.01) continue;
        siteWs.getRange(i + 1, 5).setValue(newQty);
        siteWs.getRange(i + 1, 8).setValue(newTotal);
        break;
      }
    }

    // Update Orders Summary total
    const sumWs = ss.getSheetByName('Orders Summary');
    if (sumWs) {
      const sumData = sumWs.getDataRange().getValues();
      for (let i = 1; i < sumData.length; i++) {
        if ((sumData[i][0] || '').toString().trim() !== orderId) continue;
        const current = parseFloat(sumData[i][6]) || 0;
        sumWs.getRange(i + 1, 7).setValue(Math.round(Math.max(0, current + diff) * 100) / 100);
        break;
      }
    }

    // ── Rebuild Telegram with updated qty ────────────────────────
    const catLookup = {};
    const priceWs = ss.getSheetByName('Price List');
    if (priceWs) {
      const prRows = priceWs.getDataRange().getValues();
      for (let i = 3; i < prRows.length; i++) {
        const n = (prRows[i][0] || '').toString().trim();
        if (n) catLookup[n] = (prRows[i][7] || '').toString().trim() || 'Other';
      }
    }

    const remainingCredits = Object.assign({}, buildCreditMap(ss, orderId));
    const logData2 = logWs.getDataRange().getValues(); // re-read after update
    let origTime = '', origDeliv = '', origNotes = '';
    const prepItems = [], stockItems = [];

    for (let i = 1; i < logData2.length; i++) {
      const row = logData2[i];
      if ((row[11] || '').toString().trim() !== orderId) continue;
      if (!origTime) {
        origTime  = fmtTimestamp(row[0]);
        origNotes = (row[8] || '').toString();
        const rd  = row[9];
        origDeliv = rd instanceof Date ? Utilities.formatDate(rd, Session.getScriptTimeZone(), 'dd/MM/yyyy') : (rd||'').toString();
      }
      const qty2 = parseFloat(row[4]) || 0;
      if (qty2 <= 0) continue;
      const name2  = (row[2]  || '').toString().trim();
      const unit2  = (row[3]  || '').toString().trim();
      const price2 = parseFloat(row[6]) || 0;
      const key2   = name2 + '|' + price2.toFixed(4);
      const avail  = remainingCredits[key2] || 0;
      const credited = Math.min(avail, qty2);
      remainingCredits[key2] = Math.max(0, avail - qty2);
      if (credited >= qty2) continue;
      const activeQty = Math.round((qty2 - credited) * 100) / 100;
      const tg = (row[10] || '').toString();
      const item2 = { name: name2, unit: unit2, qty: activeQty, category: catLookup[name2] || 'Other' };
      if (tg.includes('Prep') && tg.includes('Stock')) { prepItems.push(item2); stockItems.push(item2); }
      else if (tg.includes('Prep')) prepItems.push(item2);
      else stockItems.push(item2);
    }

    let oldPrepMsgId = null, oldStockMsgId = null;
    const tgWs = ss.getSheetByName('TG Messages');
    if (tgWs) {
      const tgData = tgWs.getDataRange().getValues();
      for (let i = 1; i < tgData.length; i++) {
        if ((tgData[i][0] || '').toString().trim() !== orderId) continue;
        oldPrepMsgId  = tgData[i][2] ? String(tgData[i][2]) : null;
        oldStockMsgId = tgData[i][3] ? String(tgData[i][3]) : null;
        break;
      }
    }

    const updatedTimeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    let newPrepMsgId = oldPrepMsgId, newStockMsgId = oldStockMsgId;

    const editOrSend = (chatId, oldMsgId, items, label) => {
      const text = buildTelegramText(site, items, origNotes, origDeliv, orderId, origTime, label);
      if (oldMsgId) {
        const deleted = deleteTelegramMessage(chatId, oldMsgId);
        if (deleted) return sendTelegramText(chatId, text);
        const edited = editTelegramMessage(chatId, oldMsgId, text);
        return { ok: edited, messageId: edited ? parseInt(oldMsgId) : null };
      }
      return sendTelegramText(chatId, text);
    };

    if (prepItems.length > 0) {
      const r = editOrSend(PREP_GROUP_ID, oldPrepMsgId, prepItems, 'PREP');
      if (r.messageId) newPrepMsgId = String(r.messageId);
    }
    if (stockItems.length > 0) {
      const r = editOrSend(STOCK_GROUP_ID, oldStockMsgId, stockItems, 'STOCK');
      if (r.messageId) newStockMsgId = String(r.messageId);
    }
    if (newPrepMsgId !== oldPrepMsgId || newStockMsgId !== oldStockMsgId) {
      storeTelegramMsgIds(ss, orderId, site, newPrepMsgId, newStockMsgId, updatedTimeStr);
    }

    return jsonResponse({ ok: true, newQty, newTotal });
  } catch(err) {
    Logger.log('handleChangeQty error: ' + err);
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

// ════════════════════════════════════════════════════════════════════
// CHANGE DELIVERY DATE — manager only
// Updates col J in Order Log + site sheet, then rebuilds the Telegram
// message(s) with the new date using the same delete+resend pattern
// as add_to_order so a fresh notification fires.
// ════════════════════════════════════════════════════════════════════
function handleChangeDeliveryDate(payload) {
  try {
    const orderId   = (payload.orderId   || '').trim();
    const delivDate = (payload.delivDate || '').trim(); // YYYY-MM-DD from HTML date input

    if (!orderId || !delivDate) {
      return jsonResponse({ ok: false, error: 'Missing orderId or delivDate.' });
    }

    // Convert YYYY-MM-DD → dd/MM/yyyy (display / Telegram) and a Date object (Sheets)
    const [y, m, d] = delivDate.split('-').map(Number);
    if (!y || !m || !d) return jsonResponse({ ok: false, error: 'Invalid date format.' });
    const dateObj  = new Date(y, m - 1, d);
    const delivFmt = d.toString().padStart(2,'0') + '/' + m.toString().padStart(2,'0') + '/' + y;

    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const logWs = ss.getSheetByName('Order Log');
    if (!logWs) return jsonResponse({ ok: false, error: 'Order Log not found.' });

    // Build category lookup from Price List
    const catLookup = {};
    const priceWs = ss.getSheetByName('Price List');
    if (priceWs) {
      const prRows = priceWs.getDataRange().getValues();
      for (let i = 3; i < prRows.length; i++) {
        const n = (prRows[i][0] || '').toString().trim();
        if (n) catLookup[n] = (prRows[i][7] || '').toString().trim() || 'Other';
      }
    }

    // Mutable credit copy — consumed row-by-row so re-added items aren't double-credited
    const remainingCredits = Object.assign({}, buildCreditMap(ss, orderId));

    // Read all rows for this order; collect items + meta
    const logData = logWs.getDataRange().getValues();
    let site = '', origTime = '', origNotes = '';
    const prepItems = [], stockItems = [];
    const logRowsToUpdate = [];

    for (let i = 1; i < logData.length; i++) {
      const row = logData[i];
      if ((row[11] || '').toString().trim() !== orderId) continue;
      logRowsToUpdate.push(i + 1);
      if (!site) {
        site       = (row[1]  || '').toString().trim();
        origTime   = fmtTimestamp(row[0]);
        origNotes  = (row[8]  || '').toString();
      }
      const qty2 = parseFloat(row[4]) || 0;
      if (qty2 <= 0) continue; // removed during recall
      const name2  = (row[2]  || '').toString().trim();
      const unit2  = (row[3]  || '').toString().trim();
      const price2 = parseFloat(row[6]) || 0;
      const key2   = name2 + '|' + price2.toFixed(4);
      const avail  = remainingCredits[key2] || 0;
      const creditedQty = Math.min(avail, qty2);
      remainingCredits[key2] = Math.max(0, avail - qty2); // consume
      if (creditedQty >= qty2) continue; // fully credited — exclude from Telegram
      const activeQty = Math.round((qty2 - creditedQty) * 100) / 100;
      const tg    = (row[10] || '').toString();
      const item2 = { name: name2, unit: unit2, qty: activeQty, category: catLookup[name2] || 'Other' };
      if (tg.includes('Prep') && tg.includes('Stock')) { prepItems.push(item2); stockItems.push(item2); }
      else if (tg.includes('Prep'))  prepItems.push(item2);
      else                           stockItems.push(item2);
    }

    // ── Update sheets ────────────────────────────────────────────────
    logRowsToUpdate.forEach(r => logWs.getRange(r, 10).setValue(dateObj));
    if (site) {
      const siteWs = ss.getSheetByName(site);
      if (siteWs) {
        const siteData = siteWs.getDataRange().getValues();
        for (let i = 1; i < siteData.length; i++) {
          if ((siteData[i][11] || '').toString().trim() === orderId)
            siteWs.getRange(i + 1, 10).setValue(dateObj);
        }
      }
      // Also update Orders Summary col E (Delivery Date)
      const sumWs = ss.getSheetByName('Orders Summary');
      if (sumWs) {
        const sumData = sumWs.getDataRange().getValues();
        for (let i = 1; i < sumData.length; i++) {
          if ((sumData[i][0] || '').toString().trim() === orderId) {
            sumWs.getRange(i + 1, 5).setValue(delivFmt);
            break;
          }
        }
      }
    }

    // ── Rebuild Telegram messages with new delivery date ─────────────
    let oldPrepMsgId = null, oldStockMsgId = null;
    const tgWs = ss.getSheetByName('TG Messages');
    if (tgWs) {
      const tgData = tgWs.getDataRange().getValues();
      for (let i = 1; i < tgData.length; i++) {
        if ((tgData[i][0] || '').toString().trim() !== orderId) continue;
        oldPrepMsgId  = tgData[i][2] ? String(tgData[i][2]) : null;
        oldStockMsgId = tgData[i][3] ? String(tgData[i][3]) : null;
        break;
      }
    }

    const updatedTimeStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
    let newPrepMsgId = oldPrepMsgId, newStockMsgId = oldStockMsgId;

    const editOrSend = (chatId, oldMsgId, items, label) => {
      const text = buildTelegramText(site, items, origNotes, delivFmt, orderId, origTime, label);
      if (oldMsgId) {
        const deleted = deleteTelegramMessage(chatId, oldMsgId);
        if (deleted) return sendTelegramText(chatId, text);
        const edited = editTelegramMessage(chatId, oldMsgId, text);
        return { ok: edited, messageId: edited ? parseInt(oldMsgId) : null };
      }
      return sendTelegramText(chatId, text);
    };

    let prepTgOk = null, stockTgOk = null;
    if (prepItems.length > 0) {
      const r = editOrSend(PREP_GROUP_ID, oldPrepMsgId, prepItems, 'PREP');
      prepTgOk = r.ok;
      if (r.messageId) newPrepMsgId = String(r.messageId);
    }
    if (stockItems.length > 0) {
      const r = editOrSend(STOCK_GROUP_ID, oldStockMsgId, stockItems, 'STOCK');
      stockTgOk = r.ok;
      if (r.messageId) newStockMsgId = String(r.messageId);
    }

    if (newPrepMsgId !== oldPrepMsgId || newStockMsgId !== oldStockMsgId) {
      storeTelegramMsgIds(ss, orderId, site, newPrepMsgId, newStockMsgId, updatedTimeStr);
    }

    return jsonResponse({ ok: true, updated: logRowsToUpdate.length, telegram: { prep: prepTgOk, stock: stockTgOk } });
  } catch(err) {
    Logger.log('handleChangeDeliveryDate error: ' + err);
    return jsonResponse({ ok: false, error: err.toString() });
  }
}

// ════════════════════════════════════════════════════════════════════
// HELPERS
// ════════════════════════════════════════════════════════════════════

// Formats a timestamp value from a sheet cell as "dd/MM/yy HH:mm".
// Handles both Date objects (Google Sheets auto-conversion) and strings.
function fmtTimestamp(raw) {
  if (raw instanceof Date) {
    return Utilities.formatDate(raw, Session.getScriptTimeZone(), 'dd/MM/yy HH:mm');
  }
  const s = (raw || '').toString().trim();
  // Normalise stored "dd/MM/yyyy HH:mm" → "dd/MM/yy HH:mm"
  const m = s.match(/^(\d{2}\/\d{2}\/)(\d{4})( \d{2}:\d{2})/);
  return m ? m[1] + m[2].slice(2) + m[3] : s;
}

// Returns a map of "name|price.toFixed(4)" → total credited qty for an order.
// Used when rebuilding Telegram messages so fully/partially credited items
// are excluded or reduced correctly.
function buildCreditMap(ss, orderId) {
  const map = {};
  const credWs = ss.getSheetByName('Credits');
  if (!credWs) return map;
  const credData = credWs.getDataRange().getValues();
  for (let i = 1; i < credData.length; i++) {
    if ((credData[i][2] || '').toString().trim() !== orderId) continue;
    const name  = (credData[i][3] || '').toString().trim();
    const price = parseFloat(credData[i][6]) || 0;
    const qty   = parseFloat(credData[i][4]) || 0;
    const key   = name + '|' + price.toFixed(4);
    map[key] = (map[key] || 0) + qty;
  }
  return map;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ════════════════════════════════════════════════════════════════════
// MIGRATE HISTORICAL DATA
//
// Run this ONCE from the Apps Script editor to backfill:
//   • Orders Summary  — one row per order
//   • Store sheets    — one tab per site with that site's order rows
//
// It reads entirely from the Order Log (source of truth) and
// rebuilds both destinations from scratch. Safe to re-run.
// ════════════════════════════════════════════════════════════════════
function migrateHistoricalData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const logWs = ss.getSheetByName('Order Log');
  if (!logWs) { Logger.log('Order Log not found.'); return; }

  const logData = logWs.getDataRange().getValues();
  if (logData.length < 2) { Logger.log('Order Log is empty.'); return; }

  // Helper: safely format any cell that might be a Date object
  function safeDate(val, fmt) {
    if (!val) return '';
    if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), fmt);
    return val.toString().trim();
  }

  // Group Order Log rows by Order ID
  const orderMap = {};  // orderId → order summary object
  const siteRows = {};  // site    → array of log rows

  for (let i = 1; i < logData.length; i++) {
    const row = logData[i];
    const orderId   = safeDate(row[11], 'dd/MM/yyyy'); // Col L — Order ID (text, not date)
    // Order ID is not a date — use toString directly
    const orderIdStr = (row[11] || '').toString().trim();
    if (!orderIdStr) continue;

    const site      = (row[1]  || '').toString().trim();
    const itemName  = (row[2]  || '').toString().trim();
    const unit      = (row[3]  || '').toString().trim();
    const qty       = parseFloat(row[4])  || 0;
    const supplier  = (row[5]  || '').toString().trim();
    const price     = parseFloat(row[6])  || 0;
    const total     = parseFloat(row[7])  || 0;
    const notes     = (row[8]  || '').toString().trim();
    const delivDate = safeDate(row[9],  'dd/MM/yyyy');
    const tgStatus  = (row[10] || '').toString().trim();
    const timeStr   = safeDate(row[0],  'dd/MM/yyyy HH:mm');
    const dateOnly  = safeDate(row[12], 'dd/MM/yyyy');
    const monthYear = safeDate(row[13], 'MMM-yyyy');

    if (!itemName) continue;

    // Build order summary
    if (!orderMap[orderIdStr]) {
      orderMap[orderIdStr] = {
        orderId: orderIdStr, site, submitted: timeStr,
        delivDate, notes, date: dateOnly, monthYear,
        prepTg: '', stockTg: '', items: 0, totalValue: 0
      };
    }
    const ord = orderMap[orderIdStr];
    ord.items++;
    ord.totalValue += total;
    if (tgStatus.includes('Prep')) {
      if (tgStatus.includes('❌')) ord.prepTg = '❌ Failed';
      else if (!ord.prepTg.includes('❌')) ord.prepTg = '✅ Sent';
    }
    if (tgStatus.includes('Stock')) {
      if (tgStatus.includes('❌')) ord.stockTg = '❌ Failed';
      else if (!ord.stockTg.includes('❌')) ord.stockTg = '✅ Sent';
    }

    // Collect site rows
    if (!siteRows[site]) siteRows[site] = [];
    siteRows[site].push([timeStr, site, itemName, unit, qty, supplier, price, total, notes, delivDate, tgStatus, orderIdStr, dateOnly, monthYear]);
  }

  // ── Rebuild Orders Summary ---------------------─────────────────
  let sumWs = ss.getSheetByName('Orders Summary');
  if (!sumWs) {
    sumWs = ss.insertSheet('Orders Summary');
  } else {
    const lr = sumWs.getLastRow();
    if (lr > 1) sumWs.getRange(2, 1, lr - 1, SUMMARY_HEADERS.length).clearContent();
  }
  // Ensure headers
  sumWs.getRange(1, 1, 1, SUMMARY_HEADERS.length).setValues([SUMMARY_HEADERS]).setFontWeight('bold').setBackground('#e8f0fe');
  sumWs.setFrozenRows(1);

  const summaryRows = Object.values(orderMap).map(ord => {
    const hasPrep  = ord.prepTg  !== '';
    const hasStock = ord.stockTg !== '';
    const type = hasPrep && hasStock ? 'both' : hasPrep ? 'prep' : 'stock';
    return [
      ord.orderId, ord.site, type, ord.submitted, ord.delivDate,
      ord.items, Math.round(ord.totalValue * 100) / 100,
      ord.prepTg  || '—', ord.stockTg || '—',
      ord.notes, ord.date, ord.monthYear
    ];
  });
  summaryRows.sort((a, b) => a[3].localeCompare(b[3]));

  if (summaryRows.length > 0) {
    sumWs.getRange(2, 1, summaryRows.length, SUMMARY_HEADERS.length).setValues(summaryRows);
    sumWs.autoResizeColumns(1, SUMMARY_HEADERS.length);
  }
  Logger.log('Orders Summary: ' + summaryRows.length + ' orders written.');

  // ── Rebuild site-specific sheets ---------------------───────────
  ALL_SITES.forEach(site => {
    const rows = siteRows[site];
    let sheet  = ss.getSheetByName(site);
    if (!sheet) {
      sheet = ss.insertSheet(site);
    } else {
      const lr = sheet.getLastRow();
      if (lr > 1) sheet.getRange(2, 1, lr - 1, LOG_HEADERS.length).clearContent();
    }
    sheet.getRange(1, 1, 1, LOG_HEADERS.length).setValues([LOG_HEADERS]).setFontWeight('bold').setBackground('#f3f3f3');
    sheet.setFrozenRows(1);

    if (rows && rows.length > 0) {
      sheet.getRange(2, 1, rows.length, LOG_HEADERS.length).setValues(rows);
      sheet.autoResizeColumns(1, LOG_HEADERS.length);
      Logger.log(site + ': ' + rows.length + ' rows written.');
    } else {
      Logger.log(site + ': no orders found.');
    }
  });

  Logger.log('Migration complete.');
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

// ════════════════════════════════════════════════════════════════════
// SAVE ITEM OVERRIDE — manager sets category / VAT treatment manually
// for items whose names no longer match the Price List.
// Stored in "Item Overrides" sheet; applied on every dashboard load.
// ════════════════════════════════════════════════════════════════════
function handleSaveItemOverride(payload) {
  try {
    const itemName = (payload.itemName || '').trim();
    const category = (payload.category || '').trim();
    const vat      = (payload.vat      || '').trim();

    if (!itemName) return jsonResponse({ ok: false, error: 'Missing itemName.' });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let ws = ss.getSheetByName('Item Overrides');
    if (!ws) {
      ws = ss.insertSheet('Item Overrides');
      ws.appendRow(['Item Name', 'Category', 'VAT Treatment']);
      ws.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#E8F0FE');
      ws.setFrozenRows(1);
    }

    const data = ws.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if ((data[i][0] || '').toString().trim() === itemName) {
        ws.getRange(i + 1, 2).setValue(category);
        ws.getRange(i + 1, 3).setValue(vat);
        return jsonResponse({ ok: true });
      }
    }
    ws.appendRow([itemName, category, vat]);
    return jsonResponse({ ok: true });

  } catch(err) {
    Logger.log('handleSaveItemOverride error: ' + err);
    return jsonResponse({ ok: false, error: err.toString() });
  }
}
