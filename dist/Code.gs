/**
 * Store App Prototype: a small handmade-goods shop that runs free on a Google account.
 * Google Sheet (data) + Apps Script web app + Google Drive (photos, payment screenshots) + Gmail.
 *
 * Install: paste this file as Code.gs, paste appsscript.json, run setup() once, deploy as a web app.
 * Customers: the /exec link.   Admin: /exec?page=admin (PIN).
 *
 * Money is whole rupees. Times are India time. The Sheet is never deleted from;
 * only payment screenshots are removed, from their own folder, after proofRetentionDays.
 */

var VERSION = '1.2';
var TZ = 'Asia/Kolkata';
var SH = { settings: 'Settings', dashboard: 'Dashboard', products: 'Products', orders: 'Orders', reviews: 'Reviews', log: 'Log' };
var PRODUCT_FOLDER = 'Store – Product photos';
var PROOF_FOLDER = 'Store – Payment screenshots (private)';

var PRODUCT_COLS = ['id', 'name', 'category', 'price', 'stock', 'description', 'photoIds', 'active', 'createdAt', 'updatedAt'];
var ORDER_COLS = ['id', 'createdAt', 'status', 'source', 'mode', 'payMethod', 'itemsText', 'subtotal', 'shipping', 'total',
  'name', 'phone', 'email', 'address', 'city', 'pin', 'consentAt', 'holdUntil',
  'proofRef', 'proofFileId', 'proofAt', 'proofDeletedAt', 'paid', 'paidVia', 'paidAt',
  'shipProvider', 'shipTracking', 'driverName', 'driverPhone', 'shipNote', 'shippedAt',
  'completedAt', 'cancelledAt', 'cancelReason', 'refundDue', 'refundedAt', 'flags',
  'payLinkEmailedAt', 'nudgedAt', 'stockHeld', 'itemsJson', 'token', 'reviewInviteAt'];
var REVIEW_COLS = ['id', 'createdAt', 'status', 'productId', 'productName', 'orderId', 'customerName', 'showAs', 'stars', 'text',
  'consentAt', 'checkedAt', 'reason', 'reply', 'repliedAt'];
var NUMBER_COLS = { price: 1, stock: 1, subtotal: 1, shipping: 1, total: 1, stars: 1 };
var DATE_COLS = { createdAt: 1, updatedAt: 1, consentAt: 1, holdUntil: 1, proofAt: 1, proofDeletedAt: 1, paidAt: 1,
  shippedAt: 1, completedAt: 1, cancelledAt: 1, refundedAt: 1, payLinkEmailedAt: 1, nudgedAt: 1, reviewInviteAt: 1, checkedAt: 1, repliedAt: 1 };

var DEFAULTS = [
  ['storeName', 'My Craft Store', 'Shown at the top of the shop and in emails'],
  ['tagline', 'Candles and crafts, made by hand in small batches', 'One line under the store name'],
  ['appUrl', '', 'The web app URL that ends in /exec (from Deploy)'],
  ['adminEmail', '', 'Where order alerts go'],
  ['adminPhone', '', 'Shown to customers for help, e.g. +91 98xxxxxxxx'],
  ['pickupAddress', '', 'Where customers collect in person'],
  ['upiId', '', 'UPI ID that receives payments'],
  ['upiName', '', 'Name shown in the UPI app'],
  ['shippingFee', 10, 'Rupees, charged on shipped orders below freeShippingFrom'],
  ['freeShippingFrom', 500, 'Shipped orders of this many rupees or more ship free'],
  ['shipAreaName', 'Delhi NCR', 'Shown on the Ship to me option'],
  ['shipCities', 'Delhi, Noida, Gurugram (Gurgaon), Faridabad', 'Cities customers can choose for shipping'],
  ['shipPinPrefixes', '110, 201, 122, 121', 'A shipping PIN code must start with one of these'],
  ['holdMinutes', 30, 'Minutes items are held while a customer pays by UPI'],
  ['pickupHoldHours', 72, 'Hours a pay-by-cash pickup order is kept aside'],
  ['proofRetentionDays', 90, 'Payment screenshots are moved to the Drive bin after this many days'],
  ['checkNudgeHours', 3, 'Email the admin again if a payment is still unchecked after this many hours (0 = off)'],
  ['categories', 'Candles, Gift boxes, Other crafts', 'Product categories, comma separated'],
  ['lowStockAt', 2, 'Show "Only N left" at or below this number'],
  ['adminAlerts', 'all', 'all = every order; changes = payments and reviews to check, and problems only; none'],
  ['orderPrefix', 'ORD-', 'Start of every order number'],
  ['consentText', 'I agree to {store} using my name, phone number, email and delivery address to process and deliver this order.', 'Tick box at checkout. {store} becomes the store name'],
  ['privacyText', 'DRAFT, replace with reviewed text. We collect your name, phone number, email and (for shipped orders) delivery address only to process, deliver and contact you about your order. Your details and any payment screenshot are stored in Google Sheets and Google Drive on the store owner\'s Google account. Payment screenshots are deleted after {days} days. If you write a review, it is shown with your first name and initial (or as \"a verified buyer\") after we check it, and we remove it if you ask. We do not sell or share your details, except with the delivery service that brings your order. To see, correct or delete your details, contact us at {phone}.', 'Shown from the checkout tick box'],
  ['reviewsOn', 'yes', 'yes = buyers are invited to review what they bought; no = reviews are switched off'],
  ['reviewMinForStars', 3, 'Show a product\'s average stars only once it has this many published reviews'],
  ['reviewConsentText', 'My review and the name I choose may be shown on the {store} shop. I can ask for it to be removed at any time.', 'Tick box on the review form'],
  ['reviewRules', 'Reviews are checked before they appear. We publish honest reviews, good or bad, without changing them. We do not publish reviews with abuse, personal details, or claims that a product treats or cures a health condition.', 'Shown on the review form and in the admin'],
  ['adminPin', '', 'PIN for the admin page (made by setup)'],
  ['productFolderId', '', 'Drive folder for product photos (found or made by setup)'],
  ['proofFolderId', '', 'Private Drive folder for payment screenshots (found or made by setup)']
];

/* ------------------------------------------------------------------ helpers */

function now_() { return new Date(); }
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function props_() { return PropertiesService.getScriptProperties(); }
function ms_(v) { if (!v) return 0; var d = (v instanceof Date) ? v : new Date(v); var t = d.getTime(); return isNaN(t) ? 0 : t; }
function fmt_(v) { var t = ms_(v); return t ? Utilities.formatDate(new Date(t), TZ, 'EEE d MMM, h:mm a') : ''; }
function rs_(n) { return 'Rs.' + Number(n || 0); }
function list_(s) { return String(s || '').split(',').map(function (x) { return x.trim(); }).filter(function (x) { return x; }); }
function norm_(s) { return String(s || '').toLowerCase().replace(/[‐-―\-]+/g, '-').replace(/\s+/g, ' ').trim(); }

/** Free text typed by a person: trimmed, one line, no leading formula characters, length capped. */
function text_(v, max) {
  var s = String(v == null ? '' : v).replace(/[\u0000-\u001f\u007f]+/g, ' ').replace(/\s+/g, ' ').trim();
  s = s.replace(/^[=+\-@\s]+/, '');
  return s.slice(0, max || 200);
}
function phone_(v) {
  var s = String(v == null ? '' : v).replace(/[\s\-().]/g, '');
  return /^\+?\d{8,15}$/.test(s) ? s : '';
}
function email_(v) {
  var s = String(v == null ? '' : v).trim().toLowerCase();
  return /^[a-z0-9][^\s@<>"',;]*@[^\s@<>"',;]+\.[a-z]{2,}$/.test(s) && s.length <= 120 ? s : '';
}
function token_() { return (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, ''); }

var SETTINGS_CACHE = null;
function settings_() {
  if (SETTINGS_CACHE) return SETTINGS_CACHE;
  var out = {};
  DEFAULTS.forEach(function (d) { out[d[0]] = d[1]; });
  var sh = ss_().getSheetByName(SH.settings);
  if (sh) {
    var v = sh.getDataRange().getValues();
    for (var r = 1; r < v.length; r++) { if (v[r][0] !== '' && v[r][0] != null) out[String(v[r][0]).trim()] = v[r][1]; }
  }
  ['shippingFee', 'freeShippingFrom', 'holdMinutes', 'pickupHoldHours', 'proofRetentionDays', 'checkNudgeHours', 'lowStockAt', 'reviewMinForStars'].forEach(function (k) {
    var n = Number(out[k]); out[k] = isNaN(n) || n < 0 ? Number(defaultOf_(k)) : n;
  });
  out.storeName = String(out.storeName || 'Store');
  SETTINGS_CACHE = out;
  return out;
}
function defaultOf_(k) { for (var i = 0; i < DEFAULTS.length; i++) if (DEFAULTS[i][0] === k) return DEFAULTS[i][1]; return ''; }
function setSetting_(key, value) {
  var sh = ss_().getSheetByName(SH.settings);
  var v = sh.getDataRange().getValues();
  for (var r = 1; r < v.length; r++) {
    if (String(v[r][0]).trim() === key) { sh.getRange(r + 1, 2).setNumberFormat('@').setValue(String(value)); SETTINGS_CACHE = null; return; }
  }
  var row = sh.getLastRow() + 1;
  sh.getRange(row, 2).setNumberFormat('@');
  sh.getRange(row, 1, 1, 3).setValues([[key, String(value), '']]);
  SETTINGS_CACHE = null;
}
function appUrl_() {
  var u = String(settings_().appUrl || '').trim();
  if (!u) { try { u = ScriptApp.getService().getUrl() || ''; } catch (e) { u = ''; } }
  return u;
}

/** Reads a tab into objects keyed by its header row. */
function read_(name) {
  var sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('The store is not set up yet. Run setup first.');
  var v = sh.getDataRange().getValues();
  var head = v[0], out = [];
  for (var r = 1; r < v.length; r++) {
    if (v[r][0] === '' || v[r][0] == null) continue;
    var o = { _row: r + 1 };
    for (var c = 0; c < head.length; c++) { if (head[c] !== '') o[head[c]] = v[r][c]; }
    out.push(o);
  }
  return out;
}
/** Writes one object back to its row, or adds a new row. Text columns are kept as text. */
function write_(name, o) {
  var sh = ss_().getSheetByName(name);
  var width = sh.getLastColumn();
  var head = sh.getRange(1, 1, 1, width).getValues()[0];
  var vals = [], formats = [];
  for (var c = 0; c < head.length; c++) {
    var k = head[c], x = o[k];
    if (x == null) x = '';
    if (NUMBER_COLS[k]) { vals.push(x === '' ? '' : Number(x)); formats.push('0'); }
    else if (DATE_COLS[k]) { vals.push(x === '' ? '' : new Date(ms_(x))); formats.push('d mmm yyyy, h:mm am/pm'); }
    else { vals.push(String(x)); formats.push('@'); }
  }
  var row = o._row || (sh.getLastRow() + 1);
  var range = sh.getRange(row, 1, 1, width);
  range.setNumberFormats([formats]);
  range.setValues([vals]);
  o._row = row;
  return o;
}
function log_(what, detail) {
  try {
    var sh = ss_().getSheetByName(SH.log);
    if (sh) sh.appendRow([now_(), String(what), String(detail == null ? '' : detail).slice(0, 500)]);
  } catch (e) { /* logging must never break an order */ }
}
function withLock_(fn) {
  var lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { var out = fn(); dash_(); return out; } finally { SpreadsheetApp.flush(); lock.releaseLock(); }
}
function flag_(o, text) {
  var cur = String(o.flags || '');
  if (cur.indexOf(text) < 0) o.flags = cur ? cur + ' | ' + text : text;
}
function items_(o) { try { return JSON.parse(o.itemsJson || '[]'); } catch (e) { return []; } }

/* -------------------------------------------------------------------- setup */

/** Safe to run again: keeps settings, PIN, folders and all data. */
function setup() {
  var ss = ss_();
  var st = ensureSheet_(ss, SH.settings, ['key', 'value', 'notes']);
  if (!ss.getSheetByName(SH.dashboard)) { var dsh = ss.insertSheet(SH.dashboard); dsh.setColumnWidth(1, 230); dsh.setColumnWidth(4, 230); }
  ensureSheet_(ss, SH.products, PRODUCT_COLS);
  ensureSheet_(ss, SH.orders, ORDER_COLS);
  ensureSheet_(ss, SH.reviews, REVIEW_COLS);
  ensureSheet_(ss, SH.log, ['when', 'what', 'detail']);

  var have = {};
  st.getDataRange().getValues().forEach(function (r) { have[String(r[0]).trim()] = true; });
  DEFAULTS.forEach(function (d) {
    if (have[d[0]]) return;
    var row = st.getLastRow() + 1;
    st.getRange(row, 2).setNumberFormat('@');
    st.getRange(row, 1, 1, 3).setValues([[d[0], String(d[1]), d[2]]]);
  });
  st.setColumnWidth(1, 170); st.setColumnWidth(2, 380); st.setColumnWidth(3, 460);
  SETTINGS_CACHE = null;
  var s = settings_();

  if (!String(s.adminEmail || '').trim()) { try { setSetting_('adminEmail', Session.getEffectiveUser().getEmail()); } catch (e) { } }
  if (!String(s.adminPin || '').trim()) setSetting_('adminPin', String(100000 + Math.floor(Math.random() * 900000)));

  var parent = parentFolder_();
  var pf = folderFromSetting_('productFolderId') || findOrCreateFolder_(parent, PRODUCT_FOLDER);
  setSetting_('productFolderId', pf.getId());
  var xf = folderFromSetting_('proofFolderId') || findOrCreateFolder_(parent, PROOF_FOLDER);
  setSetting_('proofFolderId', xf.getId());

  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'tick') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('tick').timeBased().everyMinutes(5).create();
  dash_();
  log_('setup', 'version ' + VERSION);
  SpreadsheetApp.flush();
  return 'Setup finished. Reload the Sheet, then use Store app > Show admin PIN and links.';
}
function ensureSheet_(ss, name, cols) {
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); }
  var width = sh.getLastColumn();
  var head = width ? sh.getRange(1, 1, 1, width).getValues()[0] : [];
  cols.forEach(function (c) {
    if (head.indexOf(c) < 0) { head.push(c); sh.getRange(1, head.length).setValue(c); }
  });
  sh.getRange(1, 1, 1, head.length).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}
function parentFolder_() {
  var it = DriveApp.getFileById(ss_().getId()).getParents();
  return it.hasNext() ? it.next() : DriveApp.getRootFolder();
}
function folderFromSetting_(key) {
  var id = String(settings_()[key] || '').trim();
  if (!id) return null;
  try { var f = DriveApp.getFolderById(id); return f.isTrashed() ? null : f; } catch (e) { return null; }
}
/** Finds a folder by name next to the Sheet (ignoring case, spacing and dash style), or makes it. */
function findOrCreateFolder_(parent, name) {
  var want = norm_(name), it = parent.getFolders();
  while (it.hasNext()) { var f = it.next(); if (norm_(f.getName()) === want) return f; }
  return parent.createFolder(name);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Store app')
    .addItem('Show admin PIN and links', 'showLinks')
    .addItem('Refresh dashboard', 'refreshDashboard')
    .addItem('Run clean-up now', 'tick')
    .addToUi();
}
function showLinks() {
  var s = settings_(), u = appUrl_();
  var left = '?'; try { left = MailApp.getRemainingDailyQuota(); } catch (e) { }
  SpreadsheetApp.getUi().alert('Store links',
    'Shop (customers): ' + (u || '(deploy the web app, then put its URL in Settings > appUrl)') +
    '\nAdmin: ' + (u ? u + '?page=admin' : '-') +
    '\nAdmin PIN: ' + s.adminPin +
    '\nEmails left today: ' + left, SpreadsheetApp.getUi().ButtonSet.OK);
}

/* ------------------------------------------------------------------ web app */

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!ss_().getSheetByName(SH.settings)) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:24px">This store is not set up yet. Run <b>setup</b> in the script editor.</p>');
  }
  var s = settings_();
  if (p.page === 'admin') {
    return page_(ADMIN_HTML, { store: s.storeName, order: text_(p.order, 40), tab: p.tab === 'reviews' ? 'reviews' : '' }, s.storeName + ' admin');
  }
  var boot = { shop: apiBoot(), order: null, review: null };
  if (p.order) { try { boot.order = apiGetOrder(String(p.order)); } catch (err) { boot.order = null; } }
  if (p.review) { boot.reviewAsked = true; try { boot.review = apiReviewInfo(String(p.review)); } catch (err) { boot.review = null; } }
  return page_(SHOP_HTML, boot, s.storeName);
}
function page_(html, boot, title) {
  var json = JSON.stringify(boot).replace(/</g, '\\u003c').replace(/\u2028|\u2029/g, ' ');
  return HtmlService.createHtmlOutput(html.replace('"__BOOT__"', function () { return json; }))
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* --------------------------------------------------------------- public API */

function publicProduct_(p) {
  return { id: String(p.id), name: String(p.name), category: String(p.category || ''), price: Number(p.price) || 0,
    stock: Math.max(0, Number(p.stock) || 0), description: String(p.description || ''), photos: list_(p.photoIds) };
}
function catalog_() {
  var s = settings_(), byProduct = {};
  if (reviewsOn_()) {
    reviews_().filter(function (r) { return String(r.status) === 'published'; })
      .sort(function (a, b) { return ms_(b.createdAt) - ms_(a.createdAt); })
      .forEach(function (r) { (byProduct[String(r.productId)] = byProduct[String(r.productId)] || []).push(r); });
  }
  return read_(SH.products).filter(function (p) { return String(p.active) === 'yes'; }).map(function (p) {
    var x = publicProduct_(p), list = byProduct[x.id] || [], sum = 0;
    list.forEach(function (r) { sum += Number(r.stars) || 0; });
    x.reviewCount = list.length;
    x.avgStars = list.length >= Math.max(1, s.reviewMinForStars) ? Math.round(sum / list.length * 10) / 10 : null;
    x.reviews = list.slice(0, 20).map(function (r) {
      return { stars: Number(r.stars) || 0, text: String(r.text || ''), by: String(r.showAs || 'A verified buyer'), date: day_(r.createdAt), reply: String(r.reply || '') };
    });
    return x;
  });
}
function day_(v) { var t = ms_(v); return t ? Utilities.formatDate(new Date(t), TZ, 'd MMM yyyy') : ''; }
function apiBoot() {
  var s = settings_();
  return {
    store: {
      name: s.storeName, tagline: String(s.tagline || ''), phone: String(s.adminPhone || ''),
      pickupAddress: String(s.pickupAddress || ''), shippingFee: s.shippingFee, freeFrom: s.freeShippingFrom,
      shipArea: String(s.shipAreaName || ''), cities: list_(s.shipCities), pinPrefixes: list_(s.shipPinPrefixes),
      holdMinutes: s.holdMinutes, pickupHoldHours: s.pickupHoldHours, lowStockAt: s.lowStockAt,
      categories: list_(s.categories), upiReady: !!String(s.upiId || '').trim(),
      consentText: String(s.consentText || '').replace(/\{store\}/g, s.storeName),
      reviewsOn: reviewsOn_(), reviewMinForStars: s.reviewMinForStars,
      privacyText: String(s.privacyText || '').replace(/\{store\}/g, s.storeName).replace(/\{days\}/g, s.proofRetentionDays).replace(/\{phone\}/g, String(s.adminPhone || 'the number above'))
    },
    products: catalog_()
  };
}

function findOrder_(orders, token) {
  token = String(token || '');
  if (token.length < 32) return null;
  for (var i = 0; i < orders.length; i++) if (String(orders[i].token) === token) return orders[i];
  return null;
}
function upiLink_(o) {
  var s = settings_();
  return 'upi://pay?pa=' + encodeURIComponent(String(s.upiId).trim()) + '&pn=' + encodeURIComponent(String(s.upiName || s.storeName)) +
    '&am=' + Number(o.total).toFixed(2) + '&cu=INR&tn=' + encodeURIComponent('Order ' + o.id);
}
function publicOrder_(o, message) {
  var s = settings_();
  var shipped = String(o.status) === 'shipped' || (String(o.status) === 'completed' && o.shippedAt);
  var expired = String(o.status) === 'cancelled' && String(o.cancelReason) === 'Not paid in time';
  return {
    token: String(o.token), id: String(o.id), status: String(o.status), mode: String(o.mode), payMethod: String(o.payMethod),
    items: items_(o), subtotal: Number(o.subtotal) || 0, shipping: Number(o.shipping) || 0, total: Number(o.total) || 0,
    name: String(o.name || ''), paid: String(o.paid) === 'yes', hasProof: !!(o.proofRef || o.proofFileId || o.proofAt),
    holdUntil: ms_(o.holdUntil), holdText: fmt_(o.holdUntil), canAddProof: String(o.status) === 'awaiting_payment' || String(o.status) === 'payment_check' || expired,
    upiId: String(s.upiId || ''), upiName: String(s.upiName || s.storeName), upiLink: upiLink_(o),
    pickupAddress: String(s.pickupAddress || ''), storePhone: String(s.adminPhone || ''),
    ship: shipped ? { provider: String(o.shipProvider || ''), tracking: String(o.shipTracking || ''), driverName: String(o.driverName || ''),
      driverPhone: String(o.driverPhone || ''), note: String(o.shipNote || ''), at: fmt_(o.shippedAt) } : null,
    cancelReason: String(o.cancelReason || ''), refundDue: String(o.refundDue) === 'yes', message: message || '',
    canReview: canReview_(o)
  };
}
function apiGetOrder(token) {
  var o = findOrder_(read_(SH.orders), token);
  if (!o) throw new Error('We could not find that order.');
  return publicOrder_(o);
}

function inShipArea_(city, pin) {
  var s = settings_();
  var cities = list_(s.shipCities).map(norm_), prefixes = list_(s.shipPinPrefixes);
  if (cities.indexOf(norm_(city)) < 0) return false;
  if (!/^\d{6}$/.test(pin)) return false;
  if (!prefixes.length) return true;
  return prefixes.some(function (p) { return pin.indexOf(p) === 0; });
}
function nextOrderId_(orders) {
  var prefix = String(settings_().orderPrefix || ''), max = 1000;
  orders.forEach(function (o) { var m = String(o.id).match(/(\d+)\s*$/); if (m) max = Math.max(max, Number(m[1])); });
  return prefix + (max + 1);
}
/** Checks a cart against the Sheet. Returns priced lines, or the problem. Prices always come from the Sheet. */
function priceCart_(cart, products) {
  if (!cart || !cart.length || cart.length > 30) return { error: 'Your cart is empty.' };
  var byId = {}, lines = [], seen = {}, problems = [];
  products.forEach(function (p) { byId[String(p.id)] = p; });
  for (var i = 0; i < cart.length; i++) {
    var id = String(cart[i].id), qty = Math.floor(Number(cart[i].qty));
    if (seen[id] || !(qty >= 1 && qty <= 50)) return { error: 'Something in your cart is not right. Please check it and try again.' };
    seen[id] = true;
    var p = byId[id];
    if (!p || String(p.active) !== 'yes') { problems.push((p ? p.name : 'An item') + ' is no longer available'); continue; }
    var left = Math.max(0, Number(p.stock) || 0);
    if (left < qty) { problems.push(left ? 'Only ' + left + ' left of ' + p.name : p.name + ' is sold out'); continue; }
    lines.push({ id: id, name: String(p.name), price: Number(p.price) || 0, qty: qty });
  }
  if (problems.length) return { error: problems.join('. ') + '. Please update your cart.', stock: true };
  var subtotal = 0;
  lines.forEach(function (l) { subtotal += l.price * l.qty; });
  return { lines: lines, subtotal: subtotal };
}
function moveStock_(lines, sign, products) {
  var byId = {};
  products.forEach(function (p) { byId[String(p.id)] = p; });
  lines.forEach(function (l) {
    var p = byId[String(l.id)];
    if (!p) return;
    p.stock = Math.max(0, (Number(p.stock) || 0) + sign * l.qty);
    p.updatedAt = now_();
    write_(SH.products, p);
  });
}
function releaseStock_(o) {
  if (String(o.stockHeld) !== 'yes') return;
  moveStock_(items_(o), +1, read_(SH.products));
  o.stockHeld = 'no';
}
function itemsText_(lines) { return lines.map(function (l) { return l.qty + ' x ' + l.name; }).join(', '); }

function apiPlaceOrder(p) {
  p = p || {};
  var s = settings_();
  var mode = p.mode === 'ship' ? 'ship' : 'pickup';
  var pay = p.pay === 'cash' ? 'cash' : 'upi';
  if (mode === 'ship' && pay === 'cash') return { ok: false, error: 'Shipped orders are paid by UPI. Cash is only for collecting in person.' };
  if (p.consent !== true) return { ok: false, error: 'Please tick the box to agree before placing your order.' };
  var name = text_(p.name, 80), phone = phone_(p.phone), email = email_(p.email);
  if (name.length < 2) return { ok: false, error: 'Please enter your full name.' };
  if (!phone) return { ok: false, error: 'Please enter a valid mobile number.' };
  if (!email) return { ok: false, error: 'Please enter a valid email address.' };
  var address = '', city = '', pin = '';
  if (mode === 'ship') {
    address = text_(p.address, 240); city = text_(p.city, 40); pin = String(p.pin || '').replace(/\D/g, '');
    if (address.length < 8) return { ok: false, error: 'Please enter your full delivery address.' };
    if (!inShipArea_(city, pin)) return { ok: false, error: 'Sorry, we deliver only within ' + s.shipAreaName + ' for now. Please check your city and PIN code, or choose Collect in person.' };
  } else if (!String(s.pickupAddress || '').trim()) {
    return { ok: false, error: 'Collecting in person is not available yet. Please choose Ship to me.' };
  }
  if (pay === 'upi' && !String(s.upiId || '').trim()) return { ok: false, error: 'UPI payment is not set up yet. Please contact us.' };

  return withLock_(function () {
    var products = read_(SH.products);
    var cart = priceCart_(p.items, products);
    if (cart.error) return { ok: false, error: cart.error, products: cart.stock ? catalog_() : null };
    var orders = read_(SH.orders), t = now_();
    var shipping = (mode === 'ship' && cart.subtotal < s.freeShippingFrom) ? s.shippingFee : 0;
    var o = {
      id: nextOrderId_(orders), createdAt: t, status: pay === 'upi' ? 'awaiting_payment' : 'cash_pickup', source: 'online',
      mode: mode, payMethod: pay, itemsText: itemsText_(cart.lines), subtotal: cart.subtotal, shipping: shipping, total: cart.subtotal + shipping,
      name: name, phone: phone, email: email, address: address, city: city, pin: pin, consentAt: t,
      holdUntil: new Date(t.getTime() + (pay === 'upi' ? s.holdMinutes * 60000 : s.pickupHoldHours * 3600000)),
      paid: 'no', stockHeld: 'yes', itemsJson: JSON.stringify(cart.lines), token: token_()
    };
    moveStock_(cart.lines, -1, products);
    write_(SH.orders, o);
    log_('order placed', o.id + ' ' + mode + ' ' + pay + ' ' + rs_(o.total));
    if (pay === 'cash') {
      mailCustomer_(o, 'Your order ' + o.id + ' is reserved', [
        'Thank you, ' + o.name + '. Your order is reserved.',
        orderLines_(o),
        'Collect from: ' + s.pickupAddress,
        'Please collect and pay ' + rs_(o.total) + ' in cash by ' + fmt_(o.holdUntil) + '. After that the items go back on sale.'
      ], 'View your order');
      alertAdmin_('all', 'New pickup order ' + o.id + ' (cash ' + rs_(o.total) + ')', [o.name + ', ' + o.phone, orderLines_(o), 'Kept aside until ' + fmt_(o.holdUntil) + '.'], o);
    }
    return { ok: true, order: publicOrder_(o) };
  });
}

function apiCancelOrder(token) {
  return withLock_(function () {
    var o = findOrder_(read_(SH.orders), token);
    if (!o) throw new Error('We could not find that order.');
    var st = String(o.status);
    if (st !== 'awaiting_payment' && st !== 'cash_pickup') throw new Error('This order can no longer be cancelled here. Please contact us.');
    releaseStock_(o);
    o.status = 'cancelled'; o.cancelledAt = now_(); o.cancelReason = 'Cancelled by customer'; o.holdUntil = '';
    write_(SH.orders, o);
    log_('order cancelled by customer', o.id);
    if (st === 'cash_pickup') alertAdmin_('all', 'Pickup order ' + o.id + ' was cancelled by the customer', [o.name + ', ' + o.phone, orderLines_(o)], o);
    return publicOrder_(o);
  });
}

/** Customer says they have paid: a 12-digit UPI reference number, a screenshot, or both. */
function apiSubmitProof(token, ref, imageDataUrl) {
  ref = String(ref || '').replace(/\D/g, '');
  var hasImage = !!imageDataUrl;
  if (ref && ref.length !== 12) throw new Error('A UPI reference number has exactly 12 digits. Please check it, or add a screenshot instead.');
  if (!ref && !hasImage) throw new Error('Please enter the 12-digit UPI reference number or add a screenshot of your payment.');
  var bytes = hasImage ? imageBytes_(imageDataUrl) : null;

  return withLock_(function () {
    var orders = read_(SH.orders);
    var o = findOrder_(orders, token);
    if (!o) throw new Error('We could not find that order.');
    var st = String(o.status), s = settings_();
    var expired = st === 'cancelled' && String(o.cancelReason) === 'Not paid in time';
    if (st !== 'awaiting_payment' && st !== 'payment_check' && !expired) throw new Error('This order does not need a payment confirmation.');

    var soldOut = false;
    if (expired) {
      var products = read_(SH.products);
      var again = priceCart_(items_(o).map(function (l) { return { id: l.id, qty: l.qty }; }), products);
      if (again.error) soldOut = true;
      else { moveStock_(items_(o), -1, products); o.stockHeld = 'yes'; o.cancelledAt = ''; o.cancelReason = ''; flag_(o, 'Paid after the hold ended'); }
    }
    if (bytes) {
      if (o.proofFileId) { try { DriveApp.getFileById(String(o.proofFileId)).setTrashed(true); } catch (e) { } }
      var folder = DriveApp.getFolderById(String(s.proofFolderId));
      var file = folder.createFile(Utilities.newBlob(bytes, 'image/jpeg', o.id + ' payment ' + Utilities.formatDate(now_(), TZ, 'yyyy-MM-dd HHmm') + '.jpg'));
      o.proofFileId = file.getId(); o.proofDeletedAt = '';
    }
    if (ref) {
      o.proofRef = ref;
      orders.forEach(function (x) { if (x !== o && String(x.proofRef) === ref) flag_(o, 'Same UPI reference as order ' + x.id); });
    }
    o.proofAt = now_(); o.holdUntil = '';
    var message;
    if (soldOut) {
      o.refundDue = 'yes';
      flag_(o, 'Paid after the hold ended and the items had sold. Check the payment and refund it');
      message = 'Your hold had ended and these items have since sold. We will check your payment and refund it. Sorry about this.';
    } else {
      o.status = 'payment_check';
      message = 'Your order is reserved, you\'ll receive an email when the order is confirmed.';
    }
    write_(SH.orders, o);
    log_('payment proof', o.id + (ref ? ' ref ' + ref : '') + (bytes ? ' screenshot' : '') + (soldOut ? ' SOLD OUT' : ''));
    var lines = [o.name + ', ' + o.phone, orderLines_(o), 'Amount to find in your UPI app: ' + rs_(o.total)];
    if (ref) lines.push('UPI reference number: ' + ref);
    if (o.flags) lines.push('Note: ' + o.flags);
    lines.push('Check the money in your UPI app before confirming. A screenshot can be faked.');
    alertAdmin_('changes', (soldOut ? 'Refund needed: ' : 'Payment to check: ') + o.id + ' ' + rs_(o.total), lines, o, bytes);
    return publicOrder_(o, message);
  });
}
function imageBytes_(dataUrl) {
  var m = String(dataUrl).match(/^data:image\/jpeg;base64,([A-Za-z0-9+\/=]+)$/);
  if (!m) throw new Error('That picture could not be read. Please try another one.');
  if (m[1].length > 2800000) throw new Error('That picture is too large. Please try another one.');
  return Utilities.base64Decode(m[1]);
}

/* -------------------------------------------------------------------- email */

function orderLines_(o) {
  var out = items_(o).map(function (l) { return l.qty + ' x ' + l.name + ' (' + rs_(l.price * l.qty) + ')'; });
  if (Number(o.shipping) > 0) out.push('Shipping ' + rs_(o.shipping));
  out.push('Total ' + rs_(o.total));
  return out.join('\n');
}
function esc_(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
function send_(to, subject, lines, button, bytes) {
  if (!to) return false;
  var s = settings_();
  try {
    if (MailApp.getRemainingDailyQuota() < 1) { log_('email NOT sent (daily limit reached)', to + ' | ' + subject); return false; }
    var html = '<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.5;color:#16211E;max-width:520px">' +
      lines.filter(function (l) { return l; }).map(function (l) { return '<p style="margin:0 0 12px">' + esc_(l).replace(/\n/g, '<br>') + '</p>'; }).join('') +
      (button ? '<p style="margin:18px 0"><a href="' + esc_(button.url) + '" style="background:#1E4A40;color:#ffffff;text-decoration:none;padding:12px 18px;border-radius:10px;display:inline-block;font-weight:bold">' + esc_(button.text) + '</a></p>' : '') +
      (bytes ? '<p style="margin:0 0 12px"><img src="cid:proof" alt="Payment screenshot" style="max-width:320px;border:1px solid #D5DEDA;border-radius:8px"></p>' : '') +
      '<p style="margin:18px 0 0;color:#55635F;font-size:13px">' + esc_(s.storeName) + (s.adminPhone ? ' · ' + esc_(s.adminPhone) : '') + '</p></div>';
    var msg = { to: to, subject: subject, htmlBody: html, body: lines.join('\n\n') + (button ? '\n\n' + button.text + ': ' + button.url : ''), name: s.storeName };
    if (s.adminEmail) msg.replyTo = String(s.adminEmail);
    if (bytes) msg.inlineImages = { proof: Utilities.newBlob(bytes, 'image/jpeg', 'payment.jpg') };
    MailApp.sendEmail(msg);
    return true;
  } catch (e) { log_('email failed', to + ' | ' + subject + ' | ' + e); return false; }
}
function mailCustomer_(o, subject, lines, buttonText, path) {
  var u = appUrl_();
  return send_(String(o.email || ''), subject, lines, u && buttonText ? { text: buttonText, url: u + (path || '?order=') + o.token } : null);
}
/** level 'changes' = things the admin must act on; 'all' = information only. */
function alertAdmin_(level, subject, lines, o, bytes, path) {
  var s = settings_(), mode = String(s.adminAlerts || 'all').toLowerCase();
  if (mode === 'none' || (mode === 'changes' && level === 'all')) return false;
  var u = appUrl_(), link = o ? '?page=admin&order=' + encodeURIComponent(o.id) : path;
  return send_(String(s.adminEmail || ''), subject, lines, u && link ? { text: 'Open in the store admin', url: u + link } : null, bytes);
}

/* ------------------------------------------------------------------ reviews */

function reviewsOn_() { return String(settings_().reviewsOn || 'yes').trim().toLowerCase() !== 'no'; }
function reviews_() { var sh = ss_().getSheetByName(SH.reviews); return sh ? read_(SH.reviews) : []; }
/** Only finished online orders with an email can be reviewed, once per product per order. */
function canReview_(o) {
  return reviewsOn_() && String(o.status) === 'completed' && String(o.source) === 'online';
}
/** "Asha Verma" -> "Asha V." */
function shortName_(name) {
  var parts = text_(name, 80).split(' ').filter(function (x) { return x; });
  if (!parts.length) return 'A verified buyer';
  return parts[0] + (parts.length > 1 ? ' ' + parts[parts.length - 1].charAt(0).toUpperCase() + '.' : '');
}
function apiReviewInfo(token) {
  var o = findOrder_(read_(SH.orders), token);
  if (!o) throw new Error('We could not find that order.');
  var s = settings_(), done = {};
  reviews_().forEach(function (r) { if (String(r.orderId) === String(o.id)) done[String(r.productId)] = String(r.status); });
  var photos = {};
  read_(SH.products).forEach(function (p) { photos[String(p.id)] = list_(p.photoIds)[0] || ''; });
  return {
    token: String(o.token), id: String(o.id), open: canReview_(o), status: String(o.status),
    shortName: shortName_(o.name), consentText: String(s.reviewConsentText || '').replace(/\{store\}/g, s.storeName),
    rules: String(s.reviewRules || ''),
    items: items_(o).map(function (l) { return { id: String(l.id), name: String(l.name), photo: photos[String(l.id)] || '', done: done[String(l.id)] || '' }; })
  };
}
/** A buyer's review. It waits for the admin's check before anyone else can see it, and is never edited. */
function apiSubmitReview(token, d) {
  d = d || {};
  var stars = Math.floor(Number(d.stars));
  if (!(stars >= 1 && stars <= 5)) throw new Error('Please choose from 1 to 5 stars.');
  if (d.consent !== true) throw new Error('Please tick the box to agree before sending your review.');
  var body = text_(d.text, 1000);
  return withLock_(function () {
    var o = findOrder_(read_(SH.orders), token);
    if (!o) throw new Error('We could not find that order.');
    if (!canReview_(o)) throw new Error('This order cannot be reviewed yet.');
    var line = null;
    items_(o).forEach(function (l) { if (String(l.id) === String(d.productId)) line = l; });
    if (!line) throw new Error('That item is not in this order.');
    var all = reviews_();
    all.forEach(function (r) { if (String(r.orderId) === String(o.id) && String(r.productId) === String(line.id)) throw new Error('You have already reviewed this item. Thank you.'); });
    var max = 0;
    all.forEach(function (r) { var m = String(r.id).match(/(\d+)$/); if (m) max = Math.max(max, Number(m[1])); });
    var t = now_();
    var r = { id: 'R' + (max + 1), createdAt: t, status: 'waiting', productId: String(line.id), productName: String(line.name), orderId: String(o.id),
      customerName: String(o.name || ''), showAs: d.anonymous === true ? 'A verified buyer' : shortName_(o.name), stars: stars, text: body, consentAt: t };
    write_(SH.reviews, r);
    log_('review received', r.id + ' ' + r.productName + ' ' + stars + ' stars');
    alertAdmin_('changes', 'New review to check: ' + r.productName + ' (' + stars + ' stars)', [
      r.customerName + ', order ' + o.id, stars + ' of 5 stars' + (body ? '\n"' + body + '"' : ''), 'It will appear in the shop only after you publish it.'
    ], null, null, '?page=admin&tab=reviews');
    return apiReviewInfo(token);
  });
}
function inviteReview_(o) {
  if (!canReview_(o) || !o.email || o.reviewInviteAt) return;
  o.reviewInviteAt = now_();
  mailCustomer_(o, 'How are you finding your order ' + o.id + '?', [
    'Hello ' + o.name + ', thank you for your order. We hope you are enjoying it.',
    'If you have a moment, we would love to hear what you think. Each piece is handmade, and your review helps others choose.',
    'This is optional. Reviews are checked before they appear, and only your first name and initial are shown.'
  ], 'Write a review', '?review=');
}
function adminReview_(r) {
  return { id: String(r.id), created: fmt_(r.createdAt), createdMs: ms_(r.createdAt), status: String(r.status), productId: String(r.productId),
    productName: String(r.productName), orderId: String(r.orderId), customerName: String(r.customerName || ''), showAs: String(r.showAs || ''),
    stars: Number(r.stars) || 0, text: String(r.text || ''), reason: String(r.reason || ''), reply: String(r.reply || ''), checked: fmt_(r.checkedAt) };
}
function reviewById_(id) {
  var list = reviews_();
  for (var i = 0; i < list.length; i++) if (String(list[i].id) === String(id)) return list[i];
  throw new Error('That review was not found.');
}
/** publish = true shows it; false keeps it hidden with a reason. Published reviews can later be hidden (e.g. the buyer asked). */
function adminReviewDecision(session, id, publish, reason) {
  return adminDo_(session, function () {
    var r = reviewById_(id);
    if (publish) {
      if (String(r.status) !== 'waiting') throw new Error('This review has already been dealt with.');
      r.status = 'published'; r.reason = '';
    } else {
      if (String(r.status) === 'hidden') throw new Error('This review is already hidden.');
      r.status = 'hidden'; r.reason = text_(reason, 120) || 'Not published';
    }
    r.checkedAt = now_();
    write_(SH.reviews, r);
    log_('review ' + r.status, r.id + (r.reason ? ' ' + r.reason : ''));
  });
}
function adminReviewReply(session, id, reply) {
  return adminDo_(session, function () {
    var r = reviewById_(id);
    r.reply = text_(reply, 600); r.repliedAt = r.reply ? now_() : '';
    write_(SH.reviews, r);
    log_('review reply', r.id);
  });
}

/* -------------------------------------------------------------------- admin */

function adminLogin(pin) {
  var props = props_(), t = now_().getTime();
  if (t < Number(props.getProperty('pinLockUntil') || 0)) return { ok: false, error: 'Too many wrong tries. Please try again in 15 minutes.' };
  if (String(pin || '').trim() !== String(settings_().adminPin).trim() || !String(settings_().adminPin).trim()) {
    var n = Number(props.getProperty('pinFails') || 0) + 1;
    if (n >= 5) { props.setProperty('pinLockUntil', String(t + 15 * 60000)); n = 0; }
    props.setProperty('pinFails', String(n));
    return { ok: false, error: 'That PIN is not right.' };
  }
  props.setProperty('pinFails', '0');
  var sessions = {};
  try { sessions = JSON.parse(props.getProperty('sessions') || '{}'); } catch (e) { sessions = {}; }
  Object.keys(sessions).forEach(function (k) { if (sessions[k] < t) delete sessions[k]; });
  var tok = token_();
  sessions[tok] = t + 30 * 86400000;
  props.setProperty('sessions', JSON.stringify(sessions));
  return { ok: true, session: tok, data: adminData_() };
}
function auth_(session) {
  var sessions = {};
  try { sessions = JSON.parse(props_().getProperty('sessions') || '{}'); } catch (e) { sessions = {}; }
  var exp = sessions[String(session || '')];
  if (!exp || exp < now_().getTime()) throw new Error('SIGNIN');
}
function adminOrder_(o) {
  return {
    id: String(o.id), created: fmt_(o.createdAt), createdMs: ms_(o.createdAt), status: String(o.status), source: String(o.source), mode: String(o.mode),
    payMethod: String(o.payMethod), items: items_(o), itemsText: String(o.itemsText || ''), subtotal: Number(o.subtotal) || 0,
    shipping: Number(o.shipping) || 0, total: Number(o.total) || 0, name: String(o.name || ''), phone: String(o.phone || ''), email: String(o.email || ''),
    address: [o.address, o.city, o.pin].filter(function (x) { return x; }).join(', '), hold: fmt_(o.holdUntil),
    proofRef: String(o.proofRef || ''), hasShot: !!o.proofFileId, shotDeleted: !!o.proofDeletedAt, proofAt: fmt_(o.proofAt),
    paid: String(o.paid) === 'yes', paidVia: String(o.paidVia || ''),
    shipProvider: String(o.shipProvider || ''), shipTracking: String(o.shipTracking || ''), driverName: String(o.driverName || ''),
    driverPhone: String(o.driverPhone || ''), shipNote: String(o.shipNote || ''), shippedAt: fmt_(o.shippedAt),
    cancelReason: String(o.cancelReason || ''), refundDue: String(o.refundDue) === 'yes', flags: String(o.flags || '')
  };
}
function adminData_() {
  var s = settings_();
  var orders = read_(SH.orders).map(adminOrder_).sort(function (a, b) { return b.createdMs - a.createdMs; });
  var open = [], done = [];
  orders.forEach(function (o) {
    var final = (o.status === 'completed' || o.status === 'cancelled') && !o.refundDue;
    if (!final) open.push(o); else if (done.length < 40) done.push(o);
  });
  var left = null; try { left = MailApp.getRemainingDailyQuota(); } catch (e) { }
  var missing = [];
  if (!String(s.upiId || '').trim()) missing.push('UPI ID');
  if (!String(s.pickupAddress || '').trim()) missing.push('pickup address');
  if (!String(s.adminPhone || '').trim()) missing.push('phone number');
  var revs = reviews_().map(adminReview_).sort(function (a, b) { return b.createdMs - a.createdMs; });
  return {
    reviews: revs.filter(function (r) { return r.status === 'waiting'; }).concat(revs.filter(function (r) { return r.status !== 'waiting'; }).slice(0, 60)),
    orders: open.concat(done),
    products: read_(SH.products).map(function (p) { var x = publicProduct_(p); x.active = String(p.active) === 'yes'; return x; }),
    store: { name: s.storeName, upiId: String(s.upiId || ''), upiName: String(s.upiName || s.storeName), categories: list_(s.categories), lowStockAt: s.lowStockAt, shopUrl: appUrl_(), emailsLeft: left, missing: missing, reviewsOn: reviewsOn_(), reviewRules: String(s.reviewRules || '') }
  };
}
function adminRefresh(session) { auth_(session); return adminData_(); }
function orderById_(id) {
  var list = read_(SH.orders);
  for (var i = 0; i < list.length; i++) if (String(list[i].id) === String(id)) return list[i];
  throw new Error('That order was not found.');
}
function adminDo_(session, fn) {
  auth_(session);
  return withLock_(function () { fn(); return adminData_(); });
}

function adminPaymentDecision(session, id, received) {
  return adminDo_(session, function () {
    var o = orderById_(id), s = settings_();
    if (String(o.status) !== 'payment_check') throw new Error('This payment has already been dealt with.');
    if (received) {
      o.status = 'confirmed'; o.paid = 'yes'; o.paidVia = 'upi'; o.paidAt = now_();
      write_(SH.orders, o);
      mailCustomer_(o, 'Order ' + o.id + ' confirmed', [
        'Thank you, ' + o.name + '. We have received your payment of ' + rs_(o.total) + ' and your order is confirmed.',
        orderLines_(o),
        String(o.mode) === 'ship' ? 'Delivery to: ' + [o.address, o.city, o.pin].join(', ') + '\nWe will email you again when it is on its way.'
          : 'Collect from: ' + s.pickupAddress + '\nYour items are kept aside for you.'
      ], 'View your order');
    } else {
      releaseStock_(o);
      o.status = 'cancelled'; o.cancelledAt = now_(); o.cancelReason = 'Payment not received';
      write_(SH.orders, o);
      mailCustomer_(o, 'We could not find your payment for order ' + o.id, [
        'Hello ' + o.name + ', we could not find a payment of ' + rs_(o.total) + ' for order ' + o.id + ', so the order has been cancelled and the items are back on sale.',
        'If you did pay, please reply to this email with your UPI reference number or a screenshot and we will sort it out.'
      ]);
    }
    log_('payment ' + (received ? 'received' : 'not received'), o.id);
  });
}

/** Mark as shipped, or change the delivery details later. All details are optional. */
function adminShip(session, id, d, notify) {
  d = d || {};
  return adminDo_(session, function () {
    var o = orderById_(id), first = String(o.status) === 'confirmed';
    if (String(o.mode) !== 'ship' || (!first && String(o.status) !== 'shipped')) throw new Error('This order cannot be marked as shipped.');
    var link = String(d.tracking || '').trim();
    if (link && !/^https?:\/\/[^\s<>"]{4,400}$/i.test(link)) throw new Error('The tracking link should start with https://');
    var dphone = String(d.driverPhone || '').trim();
    if (dphone && !phone_(dphone)) throw new Error('The driver phone number does not look right.');
    o.shipProvider = text_(d.provider, 60); o.shipTracking = link; o.driverName = text_(d.driverName, 60);
    o.driverPhone = phone_(dphone); o.shipNote = text_(d.note, 200);
    if (first) { o.status = 'shipped'; o.shippedAt = now_(); }
    write_(SH.orders, o);
    log_(first ? 'shipped' : 'delivery details changed', o.id);
    if (first || notify) {
      var lines = [(first ? 'Good news, ' + o.name + '. Your order ' + o.id + ' is on its way.' : 'The delivery details for your order ' + o.id + ' have changed.'), orderLines_(o)];
      var det = [];
      if (o.shipProvider) det.push('Delivery by: ' + o.shipProvider);
      if (o.driverName) det.push('Driver: ' + o.driverName + (o.driverPhone ? ', ' + o.driverPhone : ''));
      else if (o.driverPhone) det.push('Driver phone: ' + o.driverPhone);
      if (o.shipTracking) det.push('Track it here: ' + o.shipTracking);
      if (o.shipNote) det.push(o.shipNote);
      if (det.length) lines.push(det.join('\n'));
      lines.push('Delivery to: ' + [o.address, o.city, o.pin].join(', '));
      mailCustomer_(o, 'Order ' + o.id + (first ? ' is on its way' : ': delivery update'), lines, 'View your order');
    }
  });
}
function adminDelivered(session, id) {
  return adminDo_(session, function () {
    var o = orderById_(id);
    if (String(o.status) !== 'shipped') throw new Error('Only a shipped order can be marked as delivered.');
    o.status = 'completed'; o.completedAt = now_();
    inviteReview_(o);
    write_(SH.orders, o);
    log_('delivered', o.id);
  });
}
/** Pickup handed over. via is 'cash' or 'upi' for orders not yet paid. */
function adminCollected(session, id, via) {
  return adminDo_(session, function () {
    var o = orderById_(id), st = String(o.status);
    if (String(o.mode) !== 'pickup' || (st !== 'cash_pickup' && st !== 'confirmed')) throw new Error('This order cannot be marked as collected.');
    if (String(o.paid) !== 'yes') { o.paid = 'yes'; o.paidVia = via === 'upi' ? 'upi' : 'cash'; o.paidAt = now_(); }
    o.status = 'completed'; o.completedAt = now_(); o.holdUntil = ''; o.stockHeld = 'no';
    inviteReview_(o);
    write_(SH.orders, o);
    log_('collected', o.id + ' ' + o.paidVia);
  });
}
function adminCancel(session, id, reason) {
  return adminDo_(session, function () {
    var o = orderById_(id), st = String(o.status);
    if (st === 'cancelled' || st === 'completed' || st === 'shipped') throw new Error('This order can no longer be cancelled.');
    releaseStock_(o);
    o.status = 'cancelled'; o.cancelledAt = now_(); o.cancelReason = text_(reason, 160) || 'Cancelled by the store'; o.holdUntil = '';
    if (String(o.paid) === 'yes') o.refundDue = 'yes';
    write_(SH.orders, o);
    log_('cancelled by admin', o.id + ' ' + o.cancelReason);
    if (String(o.source) === 'online') {
      mailCustomer_(o, 'Order ' + o.id + ' has been cancelled', [
        'Hello ' + o.name + ', your order ' + o.id + ' has been cancelled.', 'Reason: ' + o.cancelReason,
        String(o.refundDue) === 'yes' ? 'Your payment of ' + rs_(o.total) + ' will be refunded to you by UPI.' : ''
      ]);
    }
  });
}
function adminRefunded(session, id) {
  return adminDo_(session, function () {
    var o = orderById_(id);
    if (String(o.refundDue) !== 'yes') throw new Error('No refund is due on this order.');
    o.refundDue = 'done'; o.refundedAt = now_();
    write_(SH.orders, o);
    log_('refunded', o.id);
  });
}
function adminProof(session, id) {
  auth_(session);
  var o = orderById_(id);
  if (!o.proofFileId) throw new Error('There is no screenshot for this order.');
  var blob = DriveApp.getFileById(String(o.proofFileId)).getBlob();
  return 'data:image/jpeg;base64,' + Utilities.base64Encode(blob.getBytes());
}

function adminSaveProduct(session, d) {
  d = d || {};
  return adminDo_(session, function () {
    var s = settings_(), products = read_(SH.products), p = null, t = now_();
    if (d.id) {
      for (var i = 0; i < products.length; i++) if (String(products[i].id) === String(d.id)) p = products[i];
      if (!p) throw new Error('That product was not found.');
    }
    var name = text_(d.name, 80), price = Math.round(Number(d.price)), stock = Math.floor(Number(d.stock));
    if (name.length < 2) throw new Error('Please give the product a name.');
    if (!(price >= 1 && price <= 1000000)) throw new Error('Please enter a price in rupees.');
    if (!(stock >= 0 && stock <= 9999)) throw new Error('Please enter how many you have.');
    var had = p ? list_(p.photoIds) : [];
    var keep = (d.keep || []).map(String).filter(function (id) { return had.indexOf(id) >= 0; });
    var add = d.add || [];
    if (keep.length + add.length > 4) throw new Error('A product can have up to 4 photos.');
    var folder = add.length ? DriveApp.getFolderById(String(s.productFolderId)) : null;
    add.forEach(function (dataUrl, n) {
      var f = folder.createFile(Utilities.newBlob(imageBytes_(dataUrl), 'image/jpeg', name + ' ' + (keep.length + n + 1) + '.jpg'));
      f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      keep.push(f.getId());
    });
    had.forEach(function (id) { if (keep.indexOf(id) < 0) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) { } } });
    if (!p) {
      var max = 100;
      products.forEach(function (x) { var m = String(x.id).match(/(\d+)$/); if (m) max = Math.max(max, Number(m[1])); });
      p = { id: 'P' + (max + 1), createdAt: t };
    }
    p.name = name; p.category = text_(d.category, 40); p.price = price; p.stock = stock; p.description = text_(d.description, 600);
    p.photoIds = keep.join(','); p.active = d.active === false ? 'no' : 'yes'; p.updatedAt = t;
    write_(SH.products, p);
    log_('product saved', p.id + ' ' + p.name);
  });
}
function adminSetStock(session, id, stock) {
  return adminDo_(session, function () {
    var n = Math.floor(Number(stock));
    if (!(n >= 0 && n <= 9999)) throw new Error('Stock must be a number from 0 up.');
    var products = read_(SH.products);
    for (var i = 0; i < products.length; i++) {
      if (String(products[i].id) === String(id)) { products[i].stock = n; products[i].updatedAt = now_(); write_(SH.products, products[i]); return; }
    }
    throw new Error('That product was not found.');
  });
}
/** A sale made face to face: stock goes down and the sale is recorded as completed. */
function adminInPersonSale(session, d) {
  d = d || {};
  return adminDo_(session, function () {
    var products = read_(SH.products);
    var cart = priceCart_(d.items, products);
    if (cart.error) throw new Error(cart.error);
    var t = now_(), via = d.via === 'upi' ? 'upi' : 'cash';
    var o = {
      id: nextOrderId_(read_(SH.orders)), createdAt: t, status: 'completed', source: 'inperson', mode: 'pickup', payMethod: via,
      itemsText: itemsText_(cart.lines), subtotal: cart.subtotal, shipping: 0, total: cart.subtotal,
      name: text_(d.name, 80), phone: phone_(d.phone), paid: 'yes', paidVia: via, paidAt: t, completedAt: t,
      stockHeld: 'no', itemsJson: JSON.stringify(cart.lines), token: token_()
    };
    moveStock_(cart.lines, -1, products);
    write_(SH.orders, o);
    log_('in-person sale', o.id + ' ' + via + ' ' + rs_(o.total));
  });
}
function custKey_(o) { var d = String(o.phone || '').replace(/\D/g, ''); return d.length >= 8 ? d.slice(-10) : ''; }
function monthOf_(v) { var t = ms_(v); return t ? Utilities.formatDate(new Date(t), TZ, 'yyyy-MM') : ''; }
/** Numbers for one month ('yyyy-MM'). A customer is counted by phone number, from paid orders. */
function statsFor_(orders, month) {
  var r = { month: month, orders: 0, sales: 0, upi: 0, cash: 0, shippingFees: 0, shipped: 0, pickup: 0, inPerson: 0, cancelled: 0,
    toCheck: 0, cashDue: 0, refundsDue: 0, items: 0, customers: 0, newCustomers: 0, returningCustomers: 0, top: [],
    reviewsIn: 0, reviewsPublished: 0, reviewsHidden: 0, reviewsWaiting: 0, avgStars: null };
  var starSum = 0;
  reviews_().forEach(function (rv) {
    if (monthOf_(rv.createdAt) !== month) return;
    r.reviewsIn++;
    var st = String(rv.status);
    if (st === 'published') { r.reviewsPublished++; starSum += Number(rv.stars) || 0; } else if (st === 'hidden') r.reviewsHidden++; else r.reviewsWaiting++;
  });
  if (r.reviewsPublished) r.avgStars = Math.round(starSum / r.reviewsPublished * 10) / 10;
  var count = {}, first = {}, seen = {};
  orders.forEach(function (o) {
    if (String(o.paid) !== 'yes' || String(o.status) === 'cancelled') return;
    var k = custKey_(o), m = monthOf_(o.createdAt);
    if (k && m && (!first[k] || m < first[k])) first[k] = m;
  });
  orders.forEach(function (o) {
    if (monthOf_(o.createdAt) !== month) return;
    var st = String(o.status);
    if (String(o.refundDue) === 'yes') r.refundsDue += Number(o.total) || 0;
    if (st === 'cancelled') { r.cancelled++; return; }
    r.orders++;
    if (st === 'payment_check') r.toCheck++;
    if (st === 'cash_pickup') r.cashDue += Number(o.total) || 0;
    if (String(o.paid) !== 'yes') return;
    var tot = Number(o.total) || 0;
    r.sales += tot; r.shippingFees += Number(o.shipping) || 0;
    if (String(o.paidVia) === 'cash') r.cash += tot; else r.upi += tot;
    if (String(o.source) === 'inperson') r.inPerson++; else if (String(o.mode) === 'ship') r.shipped++; else r.pickup++;
    items_(o).forEach(function (l) { r.items += l.qty; count[l.name] = (count[l.name] || 0) + l.qty; });
    var k = custKey_(o);
    if (k && !seen[k]) { seen[k] = true; r.customers++; if (first[k] === month) r.newCustomers++; else r.returningCustomers++; }
  });
  r.top = Object.keys(count).map(function (k) { return { name: k, qty: count[k] }; }).sort(function (a, b) { return b.qty - a.qty; }).slice(0, 5);
  return r;
}
function months_(orders) {
  var set = {}; set[monthOf_(now_())] = true;
  orders.forEach(function (o) { var m = monthOf_(o.createdAt); if (m) set[m] = true; });
  return Object.keys(set).sort().reverse();
}
function adminStats(session, month) {
  auth_(session);
  var orders = read_(SH.orders);
  month = /^\d{4}-\d{2}$/.test(String(month)) ? String(month) : monthOf_(now_());
  var r = statsFor_(orders, month), low = settings_().lowStockAt;
  r.lowStock = read_(SH.products).filter(function (p) { return String(p.active) === 'yes' && (Number(p.stock) || 0) <= low; })
    .map(function (p) { return { name: String(p.name), stock: Math.max(0, Number(p.stock) || 0) }; });
  return r;
}

/* ---------------------------------------------------------------- dashboard */

/** Rewrites the Dashboard tab. Pick a month in B3; the table below shows every month. */
function refreshDashboard_(picked) {
  var sh = ss_().getSheetByName(SH.dashboard);
  if (!sh) return;
  var orders = read_(SH.orders), months = months_(orders), thisMonth = monthOf_(now_());
  // A month picked in B3 is remembered (in M1) until the calendar month changes; then the dashboard returns to the current month.
  var note = sh.getRange(1, 13), month = thisMonth;
  if (picked) {
    var cur = sh.getRange(3, 2).getValue();
    month = (cur instanceof Date) ? monthOf_(cur) : String(cur || '').trim();
    note.setNumberFormat('@').setValue(month + '@' + thisMonth);
  } else {
    var mem = String(note.getValue() || '').split('@');
    if (mem.length === 2 && mem[1] === thisMonth) month = mem[0];
  }
  if (months.indexOf(month) < 0) month = thisMonth;
  var r = statsFor_(orders, month), s = settings_();
  var left = [
    [s.storeName + ': dashboard', ''], ['Updated ' + fmt_(now_()), ''], ['Month (pick one)', month], ['', ''],
    ['Money received (Rs.)', r.sales], ['  by UPI', r.upi], ['  by cash', r.cash], ['Orders', r.orders],
    ['  delivered', r.shipped], ['  collected in person', r.pickup], ['  in-person sales', r.inPerson], ['Items sold', r.items],
    ['Customers', r.customers], ['  new', r.newCustomers], ['  returning', r.returningCustomers],
    ['Delivery fees (Rs.)', r.shippingFees], ['Cancelled orders', r.cancelled], ['Payments waiting for a check', r.toCheck],
    ['Cash still due (Rs.)', r.cashDue], ['Refunds due (Rs.)', r.refundsDue],
    ['Reviews received', r.reviewsIn], ['  published', r.reviewsPublished], ['  not published', r.reviewsHidden], ['  waiting for a check', r.reviewsWaiting],
    ['Average stars (published)', r.avgStars == null ? '' : r.avgStars]
  ];
  var best = [['Best sellers this month', 'Sold']];
  r.top.forEach(function (t) { best.push([t.name, t.qty]); });
  while (best.length < 6) best.push(['', '']);
  var head = ['Month', 'Money received', 'UPI', 'Cash', 'Orders', 'Items sold', 'Customers', 'New', 'Returning', 'Delivery fees', 'Cancelled', 'Reviews'];
  var table = [['Month by month', '', '', '', '', '', '', '', '', '', '', ''], head];
  months.forEach(function (m) {
    var x = statsFor_(orders, m);
    table.push([m, x.sales, x.upi, x.cash, x.orders, x.items, x.customers, x.newCustomers, x.returningCustomers, x.shippingFees, x.cancelled, x.reviewsIn]);
  });
  sh.getRange(1, 1, Math.max(sh.getLastRow(), 1), 12).clearContent();
  sh.getRange(3, 2).setNumberFormat('@');
  sh.getRange(1, 1, left.length, 2).setValues(left);
  sh.getRange(5, 4, best.length, 2).setValues(best);
  var top = left.length + 2;
  sh.getRange(top, 1, table.length, 1).setNumberFormat('@');
  sh.getRange(top, 1, table.length, 12).setValues(table);
  sh.getRange(1, 1).setFontWeight('bold'); sh.getRange(5, 4, 1, 2).setFontWeight('bold'); sh.getRange(top, 1, 2, 12).setFontWeight('bold');
  sh.getRange(3, 2).setFontWeight('bold').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(months, true).setAllowInvalid(false).build());
}
function refreshDashboard() { refreshDashboard_(); }
/** Changing the month in Dashboard B3 redraws the dashboard. */
function onEdit(e) {
  try {
    var rg = e && e.range;
    if (rg && rg.getSheet().getName() === SH.dashboard && rg.getRow() === 3 && rg.getColumn() === 2) refreshDashboard_(true);
  } catch (err) { }
}
function dash_() { try { refreshDashboard_(); } catch (e) { log_('dashboard not refreshed', e); } }

/* -------------------------------------------------------------------- timer */

/** Runs every 5 minutes: releases unpaid holds, sends pay links, nudges the admin, clears old screenshots. */
function tick() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    var s = settings_(), t = now_().getTime();
    read_(SH.orders).forEach(function (o) {
      var st = String(o.status), hold = ms_(o.holdUntil);
      if (st === 'awaiting_payment') {
        if (hold && t > hold) {
          releaseStock_(o);
          o.status = 'cancelled'; o.cancelledAt = now_(); o.cancelReason = 'Not paid in time'; o.holdUntil = '';
          write_(SH.orders, o);
          log_('hold ended', o.id);
          mailCustomer_(o, 'Your order ' + o.id + ' was not completed', [
            'Hello ' + o.name + ', we did not get a payment confirmation for order ' + o.id + ' in time, so the items are back on sale.',
            'If you have already paid, open your order and add your UPI reference number or a screenshot, and we will check it.'
          ], 'Open your order');
        } else if (!o.payLinkEmailedAt && t - ms_(o.createdAt) > 3 * 60000) {
          o.payLinkEmailedAt = now_();
          write_(SH.orders, o);
          mailCustomer_(o, 'Finish your order ' + o.id, [
            'Hello ' + o.name + ', your items are held until ' + fmt_(o.holdUntil) + '.', orderLines_(o),
            'Pay ' + rs_(o.total) + ' by UPI, then add your UPI reference number or a screenshot to confirm.'
          ], 'Pay and confirm');
        }
      } else if (st === 'cash_pickup' && hold && t > hold) {
        releaseStock_(o);
        o.status = 'cancelled'; o.cancelledAt = now_(); o.cancelReason = 'Not collected within ' + s.pickupHoldHours + ' hours'; o.holdUntil = '';
        write_(SH.orders, o);
        log_('pickup hold ended', o.id);
        mailCustomer_(o, 'Your order ' + o.id + ' is no longer reserved', [
          'Hello ' + o.name + ', order ' + o.id + ' was not collected within ' + s.pickupHoldHours + ' hours, so the items are back on sale.',
          'You are welcome to order again at any time.'
        ]);
        alertAdmin_('changes', 'Pickup order ' + o.id + ' was not collected and has been released', [o.name + ', ' + o.phone, orderLines_(o)], o);
      } else if (st === 'payment_check' && s.checkNudgeHours > 0 && !o.nudgedAt && t - ms_(o.proofAt) > s.checkNudgeHours * 3600000) {
        o.nudgedAt = now_();
        write_(SH.orders, o);
        alertAdmin_('changes', 'Still waiting for your check: ' + o.id + ' ' + rs_(o.total), [o.name + ', ' + o.phone, orderLines_(o), 'The customer is waiting to hear whether the order is confirmed.'], o);
      }
    });
    dash_();
    var today = Utilities.formatDate(now_(), TZ, 'yyyy-MM-dd');
    if (props_().getProperty('lastPurge') !== today) { purgeProofs_(); props_().setProperty('lastPurge', today); }
  } finally { SpreadsheetApp.flush(); lock.releaseLock(); }
}
/** Deletes old payment screenshots. It only ever looks inside the payment screenshots folder. */
function purgeProofs_() {
  var s = settings_(), id = String(s.proofFolderId || '').trim();
  if (!id || !(s.proofRetentionDays > 0)) return 0;
  var cutoff = now_().getTime() - s.proofRetentionDays * 86400000;
  var gone = {}, n = 0, it = DriveApp.getFolderById(id).getFiles();
  while (it.hasNext()) {
    var f = it.next();
    if (f.getDateCreated().getTime() < cutoff) { gone[f.getId()] = true; f.setTrashed(true); n++; }
  }
  if (n) {
    read_(SH.orders).forEach(function (o) {
      if (o.proofFileId && gone[String(o.proofFileId)]) { o.proofFileId = ''; o.proofDeletedAt = now_(); write_(SH.orders, o); }
    });
    log_('old payment screenshots removed', n);
  }
  return n;
}

/* ------------------------------------------------------------ shop page */

var SHOP_HTML = String.raw`<!doctype html>
<html lang="en">
<head>
<base target="_top">
<meta charset="utf-8">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=Outfit:wght@500;600&display=swap">
<style>
*{box-sizing:border-box}html,body{margin:0}
body{background:#F3F6F4;color:#16211E;font-family:Inter,system-ui,sans-serif;font-size:15px;line-height:1.4;-webkit-text-size-adjust:100%}
#app{max-width:480px;margin:0 auto;min-height:100vh;display:flex;flex-direction:column}
h1,h2,h3,p{margin:0}
h1{font:600 22px/28px Outfit,sans-serif}h2{font:600 17px/22px Outfit,sans-serif}
.price{font-family:Outfit,sans-serif;font-weight:600}
.muted{color:#55635F}.small{font-size:13px;line-height:19px}.tiny{font-size:12px;line-height:17px}
.pad{padding:0 20px}.stack{display:flex;flex-direction:column;gap:12px}
.row{display:flex;align-items:center;gap:10px}.between{justify-content:space-between}.grow{flex:1;min-width:0}.wrap{flex-wrap:wrap}
.card{background:#fff;border:1px solid #E3E9E6;border-radius:14px;padding:14px}
.btn{min-height:52px;border-radius:14px;border:none;background:#1E4A40;color:#fff;font:600 16px Inter,sans-serif;padding:8px 16px;display:flex;align-items:center;justify-content:center;gap:8px;text-decoration:none;cursor:pointer;text-align:center;width:100%}
.btn.sec{background:#fff;color:#16211E;border:1px solid #B9C6C1}
.btn.sm{min-height:46px;font-size:14px;border-radius:12px}
.btn[disabled]{background:#D5DEDA;color:#3F4C48;cursor:default}
.btn.warn{background:#fff;color:#8E2A1E;border:1px solid #E2B8B0}
.btn.fit{width:auto}
.iconbtn{width:48px;height:48px;border-radius:24px;border:none;background:none;display:flex;align-items:center;justify-content:center;cursor:pointer;color:#16211E;position:relative;flex-shrink:0}
.iconbtn.ring{background:#fff;border:1px solid #D5DEDA}
.badge{position:absolute;top:-4px;right:-4px;min-width:20px;height:20px;border-radius:10px;background:#1E4A40;color:#fff;font-size:12px;font-weight:600;display:flex;align-items:center;justify-content:center;padding:0 5px}
.chip{min-height:44px;padding:0 16px;border-radius:22px;border:1px solid #D5DEDA;background:#fff;color:#16211E;font:500 14px Inter,sans-serif;cursor:pointer}
.chip.on{background:#1E4A40;color:#fff;border-color:#1E4A40}
.tag{display:inline-block;white-space:nowrap;padding:3px 8px;border-radius:8px;font-size:12px;font-weight:600;background:#E3EEE9;color:#1E4A40}
.tag.amber{background:#FDF1DC;color:#7A4206}.tag.red{background:#FBE9E6;color:#8E2A1E}
.field{display:flex;flex-direction:column;gap:6px;font-size:13px;font-weight:500}
.field input,.field select,.field textarea{height:48px;padding:0 14px;border-radius:12px;border:1px solid #B9C6C1;background:#fff;font:16px Inter,sans-serif;color:#16211E;width:100%}
.field textarea{height:auto;padding:12px 14px;resize:none}
.opt{min-height:60px;padding:10px 14px;border-radius:14px;border:2px solid #D5DEDA;background:#fff;color:#16211E;font-family:Inter,sans-serif;text-align:left;cursor:pointer;display:flex;flex-direction:column;justify-content:center;gap:2px;width:100%}
.opt.on{border-color:#1E4A40;background:#E3EEE9}.opt b{font-size:15px;font-weight:600}.opt span{font-size:13px;line-height:18px;color:#55635F}
.stepper{display:flex;align-items:center;border:1px solid #D5DEDA;border-radius:12px;height:46px;background:#fff;flex-shrink:0}
.stepper button{width:44px;height:44px;border:none;background:none;cursor:pointer;display:flex;align-items:center;justify-content:center;color:#16211E}
.stepper span{min-width:24px;text-align:center;font-weight:600;font-size:16px}
.bar{position:sticky;bottom:0;background:#fff;border-top:1px solid #D5DEDA;padding:12px 20px 16px;display:flex;flex-direction:column;gap:8px;margin-top:auto;z-index:5}
.note{padding:10px 12px;border-radius:12px;background:#FDF1DC;color:#7A4206;font-size:13px;line-height:19px}
.ph{background:#E3E9E6;display:flex;align-items:center;justify-content:center;color:#55635F;font-size:12px}
img.sq,.ph.sq{width:100%;aspect-ratio:1/1;object-fit:cover;display:block}
.thumb{width:64px;height:64px;border-radius:10px;object-fit:cover;flex-shrink:0}
.grid2{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}
.grid3{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px}
.sum{display:flex;justify-content:space-between;gap:12px}
#toast{position:fixed;left:16px;right:16px;bottom:100px;max-width:448px;margin:0 auto;background:#16211E;color:#fff;padding:12px 16px;border-radius:12px;font-size:14px;line-height:20px;display:none;z-index:70}
#busy{position:fixed;inset:0;background:rgba(243,246,244,.6);display:none;align-items:center;justify-content:center;z-index:80}
#busy div{width:36px;height:36px;border-radius:50%;border:4px solid #D5DEDA;border-top-color:#1E4A40;animation:sp 1s linear infinite}
@keyframes sp{to{transform:rotate(360deg)}}
.sheet{position:fixed;inset:0;background:rgba(22,33,30,.45);z-index:40;display:flex;align-items:flex-end;justify-content:center}
.sheet>div{background:#F3F6F4;width:100%;max-width:480px;max-height:92vh;overflow:auto;border-radius:18px 18px 0 0;padding:16px 20px 20px;display:flex;flex-direction:column;gap:12px}
a{color:#1E4A40}
.pre{white-space:pre-line}

</style>
</head>
<body>
<div id="app"></div>
<div id="toast" role="status"></div>
<div id="busy"><div></div></div>
<script>
/* QR code generator: qrcode-generator by Kazuhiko Arase, MIT licence */
var qrcode=function(){var t=function(t,r){var e=t,n=g[r],o=null,i=0,a=null,u=[],f={},c=function(t,r){o=function(t){for(var r=new Array(t),e=0;e<t;e+=1){r[e]=new Array(t);for(var n=0;n<t;n+=1)r[e][n]=null}return r}(i=4*e+17),l(0,0),l(i-7,0),l(0,i-7),s(),h(),d(t,r),e>=7&&v(t),null==a&&(a=p(e,n,u)),w(a,r)},l=function(t,r){for(var e=-1;e<=7;e+=1)if(!(t+e<=-1||i<=t+e))for(var n=-1;n<=7;n+=1)r+n<=-1||i<=r+n||(o[t+e][r+n]=0<=e&&e<=6&&(0==n||6==n)||0<=n&&n<=6&&(0==e||6==e)||2<=e&&e<=4&&2<=n&&n<=4)},h=function(){for(var t=8;t<i-8;t+=1)null==o[t][6]&&(o[t][6]=t%2==0);for(var r=8;r<i-8;r+=1)null==o[6][r]&&(o[6][r]=r%2==0)},s=function(){for(var t=B.getPatternPosition(e),r=0;r<t.length;r+=1)for(var n=0;n<t.length;n+=1){var i=t[r],a=t[n];if(null==o[i][a])for(var u=-2;u<=2;u+=1)for(var f=-2;f<=2;f+=1)o[i+u][a+f]=-2==u||2==u||-2==f||2==f||0==u&&0==f}},v=function(t){for(var r=B.getBCHTypeNumber(e),n=0;n<18;n+=1){var a=!t&&1==(r>>n&1);o[Math.floor(n/3)][n%3+i-8-3]=a}for(n=0;n<18;n+=1){a=!t&&1==(r>>n&1);o[n%3+i-8-3][Math.floor(n/3)]=a}},d=function(t,r){for(var e=n<<3|r,a=B.getBCHTypeInfo(e),u=0;u<15;u+=1){var f=!t&&1==(a>>u&1);u<6?o[u][8]=f:u<8?o[u+1][8]=f:o[i-15+u][8]=f}for(u=0;u<15;u+=1){f=!t&&1==(a>>u&1);u<8?o[8][i-u-1]=f:u<9?o[8][15-u-1+1]=f:o[8][15-u-1]=f}o[i-8][8]=!t},w=function(t,r){for(var e=-1,n=i-1,a=7,u=0,f=B.getMaskFunction(r),c=i-1;c>0;c-=2)for(6==c&&(c-=1);;){for(var g=0;g<2;g+=1)if(null==o[n][c-g]){var l=!1;u<t.length&&(l=1==(t[u]>>>a&1)),f(n,c-g)&&(l=!l),o[n][c-g]=l,-1==(a-=1)&&(u+=1,a=7)}if((n+=e)<0||i<=n){n-=e,e=-e;break}}},p=function(t,r,e){for(var n=A.getRSBlocks(t,r),o=b(),i=0;i<e.length;i+=1){var a=e[i];o.put(a.getMode(),4),o.put(a.getLength(),B.getLengthInBits(a.getMode(),t)),a.write(o)}var u=0;for(i=0;i<n.length;i+=1)u+=n[i].dataCount;if(o.getLengthInBits()>8*u)throw"code length overflow. ("+o.getLengthInBits()+">"+8*u+")";for(o.getLengthInBits()+4<=8*u&&o.put(0,4);o.getLengthInBits()%8!=0;)o.putBit(!1);for(;!(o.getLengthInBits()>=8*u||(o.put(236,8),o.getLengthInBits()>=8*u));)o.put(17,8);return function(t,r){for(var e=0,n=0,o=0,i=new Array(r.length),a=new Array(r.length),u=0;u<r.length;u+=1){var f=r[u].dataCount,c=r[u].totalCount-f;n=Math.max(n,f),o=Math.max(o,c),i[u]=new Array(f);for(var g=0;g<i[u].length;g+=1)i[u][g]=255&t.getBuffer()[g+e];e+=f;var l=B.getErrorCorrectPolynomial(c),h=k(i[u],l.getLength()-1).mod(l);for(a[u]=new Array(l.getLength()-1),g=0;g<a[u].length;g+=1){var s=g+h.getLength()-a[u].length;a[u][g]=s>=0?h.getAt(s):0}}var v=0;for(g=0;g<r.length;g+=1)v+=r[g].totalCount;var d=new Array(v),w=0;for(g=0;g<n;g+=1)for(u=0;u<r.length;u+=1)g<i[u].length&&(d[w]=i[u][g],w+=1);for(g=0;g<o;g+=1)for(u=0;u<r.length;u+=1)g<a[u].length&&(d[w]=a[u][g],w+=1);return d}(o,n)};f.addData=function(t,r){var e=null;switch(r=r||"Byte"){case"Numeric":e=M(t);break;case"Alphanumeric":e=x(t);break;case"Byte":e=m(t);break;case"Kanji":e=L(t);break;default:throw"mode:"+r}u.push(e),a=null},f.isDark=function(t,r){if(t<0||i<=t||r<0||i<=r)throw t+","+r;return o[t][r]},f.getModuleCount=function(){return i},f.make=function(){if(e<1){for(var t=1;t<40;t++){for(var r=A.getRSBlocks(t,n),o=b(),i=0;i<u.length;i++){var a=u[i];o.put(a.getMode(),4),o.put(a.getLength(),B.getLengthInBits(a.getMode(),t)),a.write(o)}var g=0;for(i=0;i<r.length;i++)g+=r[i].dataCount;if(o.getLengthInBits()<=8*g)break}e=t}c(!1,function(){for(var t=0,r=0,e=0;e<8;e+=1){c(!0,e);var n=B.getLostPoint(f);(0==e||t>n)&&(t=n,r=e)}return r}())},f.createTableTag=function(t,r){t=t||2;var e="";e+='<table style="',e+=" border-width: 0px; border-style: none;",e+=" border-collapse: collapse;",e+=" padding: 0px; margin: "+(r=void 0===r?4*t:r)+"px;",e+='">',e+="<tbody>";for(var n=0;n<f.getModuleCount();n+=1){e+="<tr>";for(var o=0;o<f.getModuleCount();o+=1)e+='<td style="',e+=" border-width: 0px; border-style: none;",e+=" border-collapse: collapse;",e+=" padding: 0px; margin: 0px;",e+=" width: "+t+"px;",e+=" height: "+t+"px;",e+=" background-color: ",e+=f.isDark(n,o)?"#000000":"#ffffff",e+=";",e+='"/>';e+="</tr>"}return e+="</tbody>",e+="</table>"},f.createSvgTag=function(t,r,e,n){var o={};"object"==typeof arguments[0]&&(t=(o=arguments[0]).cellSize,r=o.margin,e=o.alt,n=o.title),t=t||2,r=void 0===r?4*t:r,(e="string"==typeof e?{text:e}:e||{}).text=e.text||null,e.id=e.text?e.id||"qrcode-description":null,(n="string"==typeof n?{text:n}:n||{}).text=n.text||null,n.id=n.text?n.id||"qrcode-title":null;var i,a,u,c,g=f.getModuleCount()*t+2*r,l="";for(c="l"+t+",0 0,"+t+" -"+t+",0 0,-"+t+"z ",l+='<svg version="1.1" xmlns="http://www.w3.org/2000/svg"',l+=o.scalable?"":' width="'+g+'px" height="'+g+'px"',l+=' viewBox="0 0 '+g+" "+g+'" ',l+=' preserveAspectRatio="xMinYMin meet"',l+=n.text||e.text?' role="img" aria-labelledby="'+y([n.id,e.id].join(" ").trim())+'"':"",l+=">",l+=n.text?'<title id="'+y(n.id)+'">'+y(n.text)+"</title>":"",l+=e.text?'<description id="'+y(e.id)+'">'+y(e.text)+"</description>":"",l+='<rect width="100%" height="100%" fill="white" cx="0" cy="0"/>',l+='<path d="',a=0;a<f.getModuleCount();a+=1)for(u=a*t+r,i=0;i<f.getModuleCount();i+=1)f.isDark(a,i)&&(l+="M"+(i*t+r)+","+u+c);return l+='" stroke="transparent" fill="black"/>',l+="</svg>"},f.createDataURL=function(t,r){t=t||2,r=void 0===r?4*t:r;var e=f.getModuleCount()*t+2*r,n=r,o=e-r;return I(e,e,function(r,e){if(n<=r&&r<o&&n<=e&&e<o){var i=Math.floor((r-n)/t),a=Math.floor((e-n)/t);return f.isDark(a,i)?0:1}return 1})},f.createImgTag=function(t,r,e){t=t||2,r=void 0===r?4*t:r;var n=f.getModuleCount()*t+2*r,o="";return o+="<img",o+=' src="',o+=f.createDataURL(t,r),o+='"',o+=' width="',o+=n,o+='"',o+=' height="',o+=n,o+='"',e&&(o+=' alt="',o+=y(e),o+='"'),o+="/>"};var y=function(t){for(var r="",e=0;e<t.length;e+=1){var n=t.charAt(e);switch(n){case"<":r+="&lt;";break;case">":r+="&gt;";break;case"&":r+="&amp;";break;case'"':r+="&quot;";break;default:r+=n}}return r};return f.createASCII=function(t,r){if((t=t||1)<2)return function(t){t=void 0===t?2:t;var r,e,n,o,i,a=1*f.getModuleCount()+2*t,u=t,c=a-t,g={"██":"█","█ ":"▀"," █":"▄","  ":" "},l={"██":"▀","█ ":"▀"," █":" ","  ":" "},h="";for(r=0;r<a;r+=2){for(n=Math.floor((r-u)/1),o=Math.floor((r+1-u)/1),e=0;e<a;e+=1)i="█",u<=e&&e<c&&u<=r&&r<c&&f.isDark(n,Math.floor((e-u)/1))&&(i=" "),u<=e&&e<c&&u<=r+1&&r+1<c&&f.isDark(o,Math.floor((e-u)/1))?i+=" ":i+="█",h+=t<1&&r+1>=c?l[i]:g[i];h+="\n"}return a%2&&t>0?h.substring(0,h.length-a-1)+Array(a+1).join("▀"):h.substring(0,h.length-1)}(r);t-=1,r=void 0===r?2*t:r;var e,n,o,i,a=f.getModuleCount()*t+2*r,u=r,c=a-r,g=Array(t+1).join("██"),l=Array(t+1).join("  "),h="",s="";for(e=0;e<a;e+=1){for(o=Math.floor((e-u)/t),s="",n=0;n<a;n+=1)i=1,u<=n&&n<c&&u<=e&&e<c&&f.isDark(o,Math.floor((n-u)/t))&&(i=0),s+=i?g:l;for(o=0;o<t;o+=1)h+=s+"\n"}return h.substring(0,h.length-1)},f.renderTo2dContext=function(t,r){r=r||2;for(var e=f.getModuleCount(),n=0;n<e;n++)for(var o=0;o<e;o++)t.fillStyle=f.isDark(n,o)?"black":"white",t.fillRect(o*r,n*r,r,r)},f};t.stringToBytes=(t.stringToBytesFuncs={default:function(t){for(var r=[],e=0;e<t.length;e+=1){var n=t.charCodeAt(e);r.push(255&n)}return r}}).default,t.createStringToBytes=function(t,r){var e=function(){for(var e=S(t),n=function(){var t=e.read();if(-1==t)throw"eof";return t},o=0,i={};;){var a=e.read();if(-1==a)break;var u=n(),f=n()<<8|n();i[String.fromCharCode(a<<8|u)]=f,o+=1}if(o!=r)throw o+" != "+r;return i}(),n="?".charCodeAt(0);return function(t){for(var r=[],o=0;o<t.length;o+=1){var i=t.charCodeAt(o);if(i<128)r.push(i);else{var a=e[t.charAt(o)];"number"==typeof a?(255&a)==a?r.push(a):(r.push(a>>>8),r.push(255&a)):r.push(n)}}return r}};var r,e,n,o,i,a=1,u=2,f=4,c=8,g={L:1,M:0,Q:3,H:2},l=0,h=1,s=2,v=3,d=4,w=5,p=6,y=7,B=(r=[[],[6,18],[6,22],[6,26],[6,30],[6,34],[6,22,38],[6,24,42],[6,26,46],[6,28,50],[6,30,54],[6,32,58],[6,34,62],[6,26,46,66],[6,26,48,70],[6,26,50,74],[6,30,54,78],[6,30,56,82],[6,30,58,86],[6,34,62,90],[6,28,50,72,94],[6,26,50,74,98],[6,30,54,78,102],[6,28,54,80,106],[6,32,58,84,110],[6,30,58,86,114],[6,34,62,90,118],[6,26,50,74,98,122],[6,30,54,78,102,126],[6,26,52,78,104,130],[6,30,56,82,108,134],[6,34,60,86,112,138],[6,30,58,86,114,142],[6,34,62,90,118,146],[6,30,54,78,102,126,150],[6,24,50,76,102,128,154],[6,28,54,80,106,132,158],[6,32,58,84,110,136,162],[6,26,54,82,110,138,166],[6,30,58,86,114,142,170]],e=1335,n=7973,i=function(t){for(var r=0;0!=t;)r+=1,t>>>=1;return r},(o={}).getBCHTypeInfo=function(t){for(var r=t<<10;i(r)-i(e)>=0;)r^=e<<i(r)-i(e);return 21522^(t<<10|r)},o.getBCHTypeNumber=function(t){for(var r=t<<12;i(r)-i(n)>=0;)r^=n<<i(r)-i(n);return t<<12|r},o.getPatternPosition=function(t){return r[t-1]},o.getMaskFunction=function(t){switch(t){case l:return function(t,r){return(t+r)%2==0};case h:return function(t,r){return t%2==0};case s:return function(t,r){return r%3==0};case v:return function(t,r){return(t+r)%3==0};case d:return function(t,r){return(Math.floor(t/2)+Math.floor(r/3))%2==0};case w:return function(t,r){return t*r%2+t*r%3==0};case p:return function(t,r){return(t*r%2+t*r%3)%2==0};case y:return function(t,r){return(t*r%3+(t+r)%2)%2==0};default:throw"bad maskPattern:"+t}},o.getErrorCorrectPolynomial=function(t){for(var r=k([1],0),e=0;e<t;e+=1)r=r.multiply(k([1,C.gexp(e)],0));return r},o.getLengthInBits=function(t,r){if(1<=r&&r<10)switch(t){case a:return 10;case u:return 9;case f:case c:return 8;default:throw"mode:"+t}else if(r<27)switch(t){case a:return 12;case u:return 11;case f:return 16;case c:return 10;default:throw"mode:"+t}else{if(!(r<41))throw"type:"+r;switch(t){case a:return 14;case u:return 13;case f:return 16;case c:return 12;default:throw"mode:"+t}}},o.getLostPoint=function(t){for(var r=t.getModuleCount(),e=0,n=0;n<r;n+=1)for(var o=0;o<r;o+=1){for(var i=0,a=t.isDark(n,o),u=-1;u<=1;u+=1)if(!(n+u<0||r<=n+u))for(var f=-1;f<=1;f+=1)o+f<0||r<=o+f||0==u&&0==f||a==t.isDark(n+u,o+f)&&(i+=1);i>5&&(e+=3+i-5)}for(n=0;n<r-1;n+=1)for(o=0;o<r-1;o+=1){var c=0;t.isDark(n,o)&&(c+=1),t.isDark(n+1,o)&&(c+=1),t.isDark(n,o+1)&&(c+=1),t.isDark(n+1,o+1)&&(c+=1),0!=c&&4!=c||(e+=3)}for(n=0;n<r;n+=1)for(o=0;o<r-6;o+=1)t.isDark(n,o)&&!t.isDark(n,o+1)&&t.isDark(n,o+2)&&t.isDark(n,o+3)&&t.isDark(n,o+4)&&!t.isDark(n,o+5)&&t.isDark(n,o+6)&&(e+=40);for(o=0;o<r;o+=1)for(n=0;n<r-6;n+=1)t.isDark(n,o)&&!t.isDark(n+1,o)&&t.isDark(n+2,o)&&t.isDark(n+3,o)&&t.isDark(n+4,o)&&!t.isDark(n+5,o)&&t.isDark(n+6,o)&&(e+=40);var g=0;for(o=0;o<r;o+=1)for(n=0;n<r;n+=1)t.isDark(n,o)&&(g+=1);return e+=Math.abs(100*g/r/r-50)/5*10},o),C=function(){for(var t=new Array(256),r=new Array(256),e=0;e<8;e+=1)t[e]=1<<e;for(e=8;e<256;e+=1)t[e]=t[e-4]^t[e-5]^t[e-6]^t[e-8];for(e=0;e<255;e+=1)r[t[e]]=e;var n={glog:function(t){if(t<1)throw"glog("+t+")";return r[t]},gexp:function(r){for(;r<0;)r+=255;for(;r>=256;)r-=255;return t[r]}};return n}();function k(t,r){if(void 0===t.length)throw t.length+"/"+r;var e=function(){for(var e=0;e<t.length&&0==t[e];)e+=1;for(var n=new Array(t.length-e+r),o=0;o<t.length-e;o+=1)n[o]=t[o+e];return n}(),n={getAt:function(t){return e[t]},getLength:function(){return e.length},multiply:function(t){for(var r=new Array(n.getLength()+t.getLength()-1),e=0;e<n.getLength();e+=1)for(var o=0;o<t.getLength();o+=1)r[e+o]^=C.gexp(C.glog(n.getAt(e))+C.glog(t.getAt(o)));return k(r,0)},mod:function(t){if(n.getLength()-t.getLength()<0)return n;for(var r=C.glog(n.getAt(0))-C.glog(t.getAt(0)),e=new Array(n.getLength()),o=0;o<n.getLength();o+=1)e[o]=n.getAt(o);for(o=0;o<t.getLength();o+=1)e[o]^=C.gexp(C.glog(t.getAt(o))+r);return k(e,0).mod(t)}};return n}var A=function(){var t=[[1,26,19],[1,26,16],[1,26,13],[1,26,9],[1,44,34],[1,44,28],[1,44,22],[1,44,16],[1,70,55],[1,70,44],[2,35,17],[2,35,13],[1,100,80],[2,50,32],[2,50,24],[4,25,9],[1,134,108],[2,67,43],[2,33,15,2,34,16],[2,33,11,2,34,12],[2,86,68],[4,43,27],[4,43,19],[4,43,15],[2,98,78],[4,49,31],[2,32,14,4,33,15],[4,39,13,1,40,14],[2,121,97],[2,60,38,2,61,39],[4,40,18,2,41,19],[4,40,14,2,41,15],[2,146,116],[3,58,36,2,59,37],[4,36,16,4,37,17],[4,36,12,4,37,13],[2,86,68,2,87,69],[4,69,43,1,70,44],[6,43,19,2,44,20],[6,43,15,2,44,16],[4,101,81],[1,80,50,4,81,51],[4,50,22,4,51,23],[3,36,12,8,37,13],[2,116,92,2,117,93],[6,58,36,2,59,37],[4,46,20,6,47,21],[7,42,14,4,43,15],[4,133,107],[8,59,37,1,60,38],[8,44,20,4,45,21],[12,33,11,4,34,12],[3,145,115,1,146,116],[4,64,40,5,65,41],[11,36,16,5,37,17],[11,36,12,5,37,13],[5,109,87,1,110,88],[5,65,41,5,66,42],[5,54,24,7,55,25],[11,36,12,7,37,13],[5,122,98,1,123,99],[7,73,45,3,74,46],[15,43,19,2,44,20],[3,45,15,13,46,16],[1,135,107,5,136,108],[10,74,46,1,75,47],[1,50,22,15,51,23],[2,42,14,17,43,15],[5,150,120,1,151,121],[9,69,43,4,70,44],[17,50,22,1,51,23],[2,42,14,19,43,15],[3,141,113,4,142,114],[3,70,44,11,71,45],[17,47,21,4,48,22],[9,39,13,16,40,14],[3,135,107,5,136,108],[3,67,41,13,68,42],[15,54,24,5,55,25],[15,43,15,10,44,16],[4,144,116,4,145,117],[17,68,42],[17,50,22,6,51,23],[19,46,16,6,47,17],[2,139,111,7,140,112],[17,74,46],[7,54,24,16,55,25],[34,37,13],[4,151,121,5,152,122],[4,75,47,14,76,48],[11,54,24,14,55,25],[16,45,15,14,46,16],[6,147,117,4,148,118],[6,73,45,14,74,46],[11,54,24,16,55,25],[30,46,16,2,47,17],[8,132,106,4,133,107],[8,75,47,13,76,48],[7,54,24,22,55,25],[22,45,15,13,46,16],[10,142,114,2,143,115],[19,74,46,4,75,47],[28,50,22,6,51,23],[33,46,16,4,47,17],[8,152,122,4,153,123],[22,73,45,3,74,46],[8,53,23,26,54,24],[12,45,15,28,46,16],[3,147,117,10,148,118],[3,73,45,23,74,46],[4,54,24,31,55,25],[11,45,15,31,46,16],[7,146,116,7,147,117],[21,73,45,7,74,46],[1,53,23,37,54,24],[19,45,15,26,46,16],[5,145,115,10,146,116],[19,75,47,10,76,48],[15,54,24,25,55,25],[23,45,15,25,46,16],[13,145,115,3,146,116],[2,74,46,29,75,47],[42,54,24,1,55,25],[23,45,15,28,46,16],[17,145,115],[10,74,46,23,75,47],[10,54,24,35,55,25],[19,45,15,35,46,16],[17,145,115,1,146,116],[14,74,46,21,75,47],[29,54,24,19,55,25],[11,45,15,46,46,16],[13,145,115,6,146,116],[14,74,46,23,75,47],[44,54,24,7,55,25],[59,46,16,1,47,17],[12,151,121,7,152,122],[12,75,47,26,76,48],[39,54,24,14,55,25],[22,45,15,41,46,16],[6,151,121,14,152,122],[6,75,47,34,76,48],[46,54,24,10,55,25],[2,45,15,64,46,16],[17,152,122,4,153,123],[29,74,46,14,75,47],[49,54,24,10,55,25],[24,45,15,46,46,16],[4,152,122,18,153,123],[13,74,46,32,75,47],[48,54,24,14,55,25],[42,45,15,32,46,16],[20,147,117,4,148,118],[40,75,47,7,76,48],[43,54,24,22,55,25],[10,45,15,67,46,16],[19,148,118,6,149,119],[18,75,47,31,76,48],[34,54,24,34,55,25],[20,45,15,61,46,16]],r=function(t,r){var e={};return e.totalCount=t,e.dataCount=r,e},e={};return e.getRSBlocks=function(e,n){var o=function(r,e){switch(e){case g.L:return t[4*(r-1)+0];case g.M:return t[4*(r-1)+1];case g.Q:return t[4*(r-1)+2];case g.H:return t[4*(r-1)+3];default:return}}(e,n);if(void 0===o)throw"bad rs block @ typeNumber:"+e+"/errorCorrectionLevel:"+n;for(var i=o.length/3,a=[],u=0;u<i;u+=1)for(var f=o[3*u+0],c=o[3*u+1],l=o[3*u+2],h=0;h<f;h+=1)a.push(r(c,l));return a},e}(),b=function(){var t=[],r=0,e={getBuffer:function(){return t},getAt:function(r){var e=Math.floor(r/8);return 1==(t[e]>>>7-r%8&1)},put:function(t,r){for(var n=0;n<r;n+=1)e.putBit(1==(t>>>r-n-1&1))},getLengthInBits:function(){return r},putBit:function(e){var n=Math.floor(r/8);t.length<=n&&t.push(0),e&&(t[n]|=128>>>r%8),r+=1}};return e},M=function(t){var r=a,e=t,n={getMode:function(){return r},getLength:function(t){return e.length},write:function(t){for(var r=e,n=0;n+2<r.length;)t.put(o(r.substring(n,n+3)),10),n+=3;n<r.length&&(r.length-n==1?t.put(o(r.substring(n,n+1)),4):r.length-n==2&&t.put(o(r.substring(n,n+2)),7))}},o=function(t){for(var r=0,e=0;e<t.length;e+=1)r=10*r+i(t.charAt(e));return r},i=function(t){if("0"<=t&&t<="9")return t.charCodeAt(0)-"0".charCodeAt(0);throw"illegal char :"+t};return n},x=function(t){var r=u,e=t,n={getMode:function(){return r},getLength:function(t){return e.length},write:function(t){for(var r=e,n=0;n+1<r.length;)t.put(45*o(r.charAt(n))+o(r.charAt(n+1)),11),n+=2;n<r.length&&t.put(o(r.charAt(n)),6)}},o=function(t){if("0"<=t&&t<="9")return t.charCodeAt(0)-"0".charCodeAt(0);if("A"<=t&&t<="Z")return t.charCodeAt(0)-"A".charCodeAt(0)+10;switch(t){case" ":return 36;case"$":return 37;case"%":return 38;case"*":return 39;case"+":return 40;case"-":return 41;case".":return 42;case"/":return 43;case":":return 44;default:throw"illegal char :"+t}};return n},m=function(r){var e=f,n=t.stringToBytes(r),o={getMode:function(){return e},getLength:function(t){return n.length},write:function(t){for(var r=0;r<n.length;r+=1)t.put(n[r],8)}};return o},L=function(r){var e=c,n=t.stringToBytesFuncs.SJIS;if(!n)throw"sjis not supported.";!function(){var t=n("友");if(2!=t.length||38726!=(t[0]<<8|t[1]))throw"sjis not supported."}();var o=n(r),i={getMode:function(){return e},getLength:function(t){return~~(o.length/2)},write:function(t){for(var r=o,e=0;e+1<r.length;){var n=(255&r[e])<<8|255&r[e+1];if(33088<=n&&n<=40956)n-=33088;else{if(!(57408<=n&&n<=60351))throw"illegal char at "+(e+1)+"/"+n;n-=49472}n=192*(n>>>8&255)+(255&n),t.put(n,13),e+=2}if(e<r.length)throw"illegal char at "+(e+1)}};return i},D=function(){var t=[],r={writeByte:function(r){t.push(255&r)},writeShort:function(t){r.writeByte(t),r.writeByte(t>>>8)},writeBytes:function(t,e,n){e=e||0,n=n||t.length;for(var o=0;o<n;o+=1)r.writeByte(t[o+e])},writeString:function(t){for(var e=0;e<t.length;e+=1)r.writeByte(t.charCodeAt(e))},toByteArray:function(){return t},toString:function(){var r="";r+="[";for(var e=0;e<t.length;e+=1)e>0&&(r+=","),r+=t[e];return r+="]"}};return r},S=function(t){var r=t,e=0,n=0,o=0,i={read:function(){for(;o<8;){if(e>=r.length){if(0==o)return-1;throw"unexpected end of file./"+o}var t=r.charAt(e);if(e+=1,"="==t)return o=0,-1;t.match(/^\s$/)||(n=n<<6|a(t.charCodeAt(0)),o+=6)}var i=n>>>o-8&255;return o-=8,i}},a=function(t){if(65<=t&&t<=90)return t-65;if(97<=t&&t<=122)return t-97+26;if(48<=t&&t<=57)return t-48+52;if(43==t)return 62;if(47==t)return 63;throw"c:"+t};return i},I=function(t,r,e){for(var n=function(t,r){var e=t,n=r,o=new Array(t*r),i={setPixel:function(t,r,n){o[r*e+t]=n},write:function(t){t.writeString("GIF87a"),t.writeShort(e),t.writeShort(n),t.writeByte(128),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(255),t.writeByte(255),t.writeByte(255),t.writeString(","),t.writeShort(0),t.writeShort(0),t.writeShort(e),t.writeShort(n),t.writeByte(0);var r=a(2);t.writeByte(2);for(var o=0;r.length-o>255;)t.writeByte(255),t.writeBytes(r,o,255),o+=255;t.writeByte(r.length-o),t.writeBytes(r,o,r.length-o),t.writeByte(0),t.writeString(";")}},a=function(t){for(var r=1<<t,e=1+(1<<t),n=t+1,i=u(),a=0;a<r;a+=1)i.add(String.fromCharCode(a));i.add(String.fromCharCode(r)),i.add(String.fromCharCode(e));var f,c,g,l=D(),h=(f=l,c=0,g=0,{write:function(t,r){if(t>>>r!=0)throw"length over";for(;c+r>=8;)f.writeByte(255&(t<<c|g)),r-=8-c,t>>>=8-c,g=0,c=0;g|=t<<c,c+=r},flush:function(){c>0&&f.writeByte(g)}});h.write(r,n);var s=0,v=String.fromCharCode(o[s]);for(s+=1;s<o.length;){var d=String.fromCharCode(o[s]);s+=1,i.contains(v+d)?v+=d:(h.write(i.indexOf(v),n),i.size()<4095&&(i.size()==1<<n&&(n+=1),i.add(v+d)),v=d)}return h.write(i.indexOf(v),n),h.write(e,n),h.flush(),l.toByteArray()},u=function(){var t={},r=0,e={add:function(n){if(e.contains(n))throw"dup key:"+n;t[n]=r,r+=1},size:function(){return r},indexOf:function(r){return t[r]},contains:function(r){return void 0!==t[r]}};return e};return i}(t,r),o=0;o<r;o+=1)for(var i=0;i<t;i+=1)n.setPixel(i,o,e(i,o));var a=D();n.write(a);for(var u=function(){var t=0,r=0,e=0,n="",o={},i=function(t){n+=String.fromCharCode(a(63&t))},a=function(t){if(t<0);else{if(t<26)return 65+t;if(t<52)return t-26+97;if(t<62)return t-52+48;if(62==t)return 43;if(63==t)return 47}throw"n:"+t};return o.writeByte=function(n){for(t=t<<8|255&n,r+=8,e+=1;r>=6;)i(t>>>r-6),r-=6},o.flush=function(){if(r>0&&(i(t<<6-r),t=0,r=0),e%3!=0)for(var o=3-e%3,a=0;a<o;a+=1)n+="="},o.toString=function(){return n},o}(),f=a.toByteArray(),c=0;c<f.length;c+=1)u.writeByte(f[c]);return u.flush(),"data:image/gif;base64,"+u};return t}();qrcode.stringToBytesFuncs["UTF-8"]=function(t){return function(t){for(var r=[],e=0;e<t.length;e++){var n=t.charCodeAt(e);n<128?r.push(n):n<2048?r.push(192|n>>6,128|63&n):n<55296||n>=57344?r.push(224|n>>12,128|n>>6&63,128|63&n):(e++,n=65536+((1023&n)<<10|1023&t.charCodeAt(e)),r.push(240|n>>18,128|n>>12&63,128|n>>6&63,128|63&n))}return r}(t)},function(t){"function"==typeof define&&define.amd?define([],t):"object"==typeof exports&&(module.exports=t())}(function(){return qrcode});
</script>
<script>
function h(tag, attrs, kids) {
  var e = document.createElement(tag), k;
  if (attrs) for (k in attrs) {
    var v = attrs[k];
    if (v == null || v === false) continue;
    if (k === 'class') e.className = v;
    else if (k === 'text') e.textContent = v;
    else if (k === 'value') e.value = v;
    else if (k.slice(0, 2) === 'on') e.addEventListener(k.slice(2), v);
    else e.setAttribute(k, v === true ? '' : v);
  }
  (kids || []).forEach(function (c) {
    if (c == null || c === false) return;
    e.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
  });
  return e;
}
var IC = {
  back: 'M15 5l-7 7 7 7', plus: 'M12 5v14M5 12h14', minus: 'M5 12h14', x: 'M6 6l12 12M18 6L6 18', check: 'M5 12l5 5 9-10',
  bag: 'M6 8h12l-1 12H7L6 8z|M9 8V6a3 3 0 0 1 6 0v2', up: 'M12 16V5M7 10l5-5 5 5M5 19h14',
  cam: 'M4 8h3l2-3h6l2 3h3v11H4z|M12 16.5a3.5 3.5 0 1 0 0-7 3.5 3.5 0 0 0 0 7z', pic: 'M4 5h16v14H4z|M4 16l5-5 4 4 3-3 4 4'
};
function ic(name, size) {
  var ns = 'http://www.w3.org/2000/svg', s = document.createElementNS(ns, 'svg');
  var a = { width: size || 22, height: size || 22, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', 'stroke-width': 2, 'stroke-linecap': 'round', 'stroke-linejoin': 'round', 'aria-hidden': 'true' };
  for (var k in a) s.setAttribute(k, a[k]);
  IC[name].split('|').forEach(function (d) { var p = document.createElementNS(ns, 'path'); p.setAttribute('d', d); s.appendChild(p); });
  return s;
}
var toastTimer = null;
function toast(msg) {
  var t = document.getElementById('toast');
  t.textContent = msg; t.style.display = 'block';
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function () { t.style.display = 'none'; }, 5000);
}
function busy(on) { document.getElementById('busy').style.display = on ? 'flex' : 'none'; }
function call(fn, args, ok, opts) {
  opts = opts || {};
  if (!opts.quiet) busy(true);
  var r = google.script.run
    .withSuccessHandler(function (x) { if (!opts.quiet) busy(false); ok(x); })
    .withFailureHandler(function (e) {
      if (!opts.quiet) busy(false);
      var m = String((e && e.message) || '').replace(/^Error:\s*/, '');
      if (opts.fail && opts.fail(m)) return;
      if (!opts.quiet) toast(m && m.length < 220 ? m : 'Something went wrong. Please try again.');
    });
  r[fn].apply(r, args);
}
function rs(n) { return 'Rs.' + n; }
function photo(id, alt, cls, w) {
  if (!id) return h('div', { class: 'ph ' + cls, text: 'No photo yet' });
  var im = h('img', { class: cls, alt: alt, loading: 'lazy', src: 'https://lh3.googleusercontent.com/d/' + id + '=w' + (w || 600) });
  im.addEventListener('error', function () {
    if (im.getAttribute('data-f')) return;
    im.setAttribute('data-f', '1');
    im.src = 'https://drive.google.com/thumbnail?id=' + id + '&sz=w' + (w || 600);
  });
  return im;
}
/** Shrinks a picked photo on the phone. square=true crops the middle to a square. */
function shrink(file, maxSide, square, quality, cb) {
  var fr = new FileReader();
  fr.onerror = function () { toast('That picture could not be read.'); };
  fr.onload = function () {
    var im = new Image();
    im.onerror = function () { toast('That picture could not be read. Please try a JPG or PNG.'); };
    im.onload = function () {
      var sw = im.naturalWidth, sh = im.naturalHeight, sx = 0, sy = 0;
      if (square) { var m = Math.min(sw, sh); sx = (sw - m) / 2; sy = (sh - m) / 2; sw = sh = m; }
      var k = Math.min(1, maxSide / Math.max(sw, sh));
      var c = document.createElement('canvas');
      c.width = Math.round(sw * k); c.height = Math.round(sh * k);
      var x = c.getContext('2d');
      x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height);
      x.drawImage(im, sx, sy, sw, sh, 0, 0, c.width, c.height);
      cb(c.toDataURL('image/jpeg', quality));
    };
    im.src = fr.result;
  };
  fr.readAsDataURL(file);
}
function qrCanvas(text, size) {
  var c = h('canvas', { width: size, height: size, role: 'img', 'aria-label': 'UPI QR code' });
  c.style.cssText = 'width:' + size / 2 + 'px;height:' + size / 2 + 'px;border-radius:8px;background:#fff';
  try {
    var q = qrcode(0, 'M'); q.addData(text); q.make();
    var n = q.getModuleCount(), cell = Math.floor(size / (n + 6)), off = Math.floor((size - n * cell) / 2), x = c.getContext('2d');
    x.fillStyle = '#fff'; x.fillRect(0, 0, size, size); x.fillStyle = '#000';
    for (var r = 0; r < n; r++) for (var k = 0; k < n; k++) if (q.isDark(r, k)) x.fillRect(off + k * cell, off + r * cell, cell, cell);
    return c;
  } catch (e) { return null; }
}
function stepper(value, min, max, onChange, what) {
  var num = h('span', { 'aria-live': 'polite', text: String(value) });
  function set(v) { v = Math.max(min, Math.min(max, v)); if (v === value) return; value = v; num.textContent = String(v); onChange(v); }
  return h('div', { class: 'stepper', role: 'group', 'aria-label': what }, [
    h('button', { type: 'button', 'aria-label': 'One less', onclick: function () { set(value - 1); } }, [ic('minus', 16)]),
    num,
    h('button', { type: 'button', 'aria-label': 'One more', onclick: function () { set(value + 1); } }, [ic('plus', 16)])
  ]);
}
function field(label, input) { return h('label', { class: 'field' }, [label, input]); }
function keep(key, val) {
  try {
    if (val === undefined) return window.localStorage.getItem(key);
    if (val === null) window.localStorage.removeItem(key); else window.localStorage.setItem(key, val);
  } catch (e) { }
  return null;
}
function sheet(kids, onClose) {
  var s = h('div', { class: 'sheet', role: 'dialog', 'aria-modal': 'true' }, [h('div', null, kids)]);
  s.addEventListener('click', function (e) { if (e.target === s) close(); });
  function close() { if (s.parentNode) s.parentNode.removeChild(s); if (onClose) onClose(); }
  document.body.appendChild(s);
  return close;
}
function topbar(title, onBack, right) {
  return h('div', { class: 'row', style: 'padding:12px 12px 4px;gap:4px' }, [
    onBack ? h('button', { type: 'button', class: 'iconbtn', 'aria-label': 'Back', onclick: onBack }, [ic('back', 24)]) : h('div', { style: 'width:8px' }),
    h('h1', { class: 'grow', text: title }), right || null
  ]);
}

var BOOT = "__BOOT__";
var S = BOOT.shop.store, PRODUCTS = BOOT.shop.products;
var app = document.getElementById('app');
var st = {
  view: 'shop', cat: 'All', pid: null, qty: 1, shot: 0, cart: {}, order: BOOT.order, review: BOOT.review, rv: {},
  form: { mode: S.pickupAddress ? 'pickup' : 'ship', pay: 'upi', name: '', phone: '', email: '', address: '', city: '', pin: '', consent: false },
  proof: { ref: '', img: null }
};
var pollTimer = null;
try { st.cart = JSON.parse(keep('ll_cart') || '{}') || {}; } catch (e) { st.cart = {}; }

function product(id) { for (var i = 0; i < PRODUCTS.length; i++) if (PRODUCTS[i].id === id) return PRODUCTS[i]; return null; }
function lines() {
  var out = [];
  Object.keys(st.cart).forEach(function (id) {
    var p = product(id), q = Math.min(st.cart[id], p ? p.stock : 0);
    if (!p || q < 1) { delete st.cart[id]; return; }
    st.cart[id] = q;
    out.push({ p: p, qty: q });
  });
  return out;
}
function saveCart() { keep('ll_cart', JSON.stringify(st.cart)); }
function totals() {
  var sub = 0, n = 0;
  lines().forEach(function (l) { sub += l.p.price * l.qty; n += l.qty; });
  var ship = (st.form.mode === 'ship' && sub < S.freeFrom && sub > 0) ? S.shippingFee : 0;
  return { sub: sub, count: n, ship: ship, total: sub + ship };
}
function go(view, more) {
  st.view = view;
  if (more) for (var k in more) st[k] = more[k];
  render();
  window.scrollTo(0, 0);
}
function setOrder(o) {
  st.order = o;
  if (o && (o.status === 'awaiting_payment' || o.status === 'payment_check' || o.status === 'cash_pickup' || o.status === 'confirmed' || o.status === 'shipped')) keep('ll_order', o.token);
  else keep('ll_order', null);
}
function openOrder(o) { setOrder(o); go(o.status === 'awaiting_payment' ? 'pay' : 'status'); }
function render() {
  clearInterval(pollTimer); pollTimer = null;
  app.textContent = '';
  var v = { shop: viewShop, product: viewProduct, cart: viewCart, pay: viewPay, confirm: viewConfirm, status: viewStatus, review: viewReview }[st.view]();
  v.forEach(function (n) { if (n) app.appendChild(n); });
}
function cartButton() {
  var n = totals().count;
  return h('button', { type: 'button', class: 'iconbtn ring', 'aria-label': 'Open cart, ' + n + ' items', onclick: function () { go('cart'); } },
    [ic('bag'), n ? h('span', { class: 'badge', text: String(n) }) : null]);
}
function stockTag(p) {
  if (p.stock < 1) return h('span', { class: 'tag red', text: 'Sold out' });
  if (p.stock <= S.lowStockAt) return h('span', { class: 'tag amber', text: 'Only ' + p.stock + ' left' });
  return null;
}
function add(p, qty) {
  var have = st.cart[p.id] || 0;
  if (have + qty > p.stock) { toast(p.stock ? 'Only ' + p.stock + ' left of ' + p.name + '.' : p.name + ' is sold out.'); return false; }
  st.cart[p.id] = have + qty; saveCart(); return true;
}

/* ------------------------------------------------------------------- shop */
function viewShop() {
  var cats = ['All'].concat(S.categories.filter(function (c) { return PRODUCTS.some(function (p) { return p.category === c; }); }));
  var list = PRODUCTS.filter(function (p) { return st.cat === 'All' || p.category === st.cat; });
  var t = totals(), saved = keep('ll_order');
  return [
    h('div', { class: 'row between', style: 'padding:20px 20px 12px' }, [
      h('div', { class: 'grow' }, [h('div', { style: 'font:600 22px/28px Outfit,sans-serif', text: S.name }), h('div', { class: 'small muted', text: S.tagline })]),
      cartButton()
    ]),
    saved ? h('div', { class: 'pad', style: 'padding-bottom:12px' }, [h('button', { type: 'button', class: 'btn sec sm', text: 'You have an order in progress. View it',
      onclick: function () { call('apiGetOrder', [saved], openOrder, { fail: function () { keep('ll_order', null); render(); return false; } }); } })]) : null,
    cats.length > 2 ? h('div', { class: 'row wrap pad', style: 'gap:8px;padding-bottom:16px' }, cats.map(function (c) {
      return h('button', { type: 'button', class: 'chip' + (c === st.cat ? ' on' : ''), 'aria-pressed': String(c === st.cat), text: c, onclick: function () { go('shop', { cat: c }); } });
    })) : null,
    list.length ? h('div', { class: 'grid2 pad' }, list.map(function (p) {
      return h('div', { class: 'card', style: 'padding:0;overflow:hidden;display:flex;flex-direction:column' }, [
        h('button', { type: 'button', 'aria-label': 'View ' + p.name, style: 'border:none;padding:0;background:none;cursor:pointer;position:relative;display:block', onclick: function () { go('product', { pid: p.id, qty: 1, shot: 0 }); } }, [
          photo(p.photos[0], p.name, 'sq', 500),
          stockTag(p) ? h('span', { style: 'position:absolute;left:8px;top:8px' }, [stockTag(p)]) : null
        ]),
        h('div', { style: 'padding:10px 12px 12px;display:flex;flex-direction:column;gap:8px;flex:1' }, [
          h('div', { style: 'font-size:14px;line-height:19px;font-weight:500;min-height:38px', text: p.name }),
          p.avgStars != null ? h('div', { class: 'small', style: 'margin-top:-4px' }, [starRow(p.avgStars, 14), h('span', { class: 'muted', text: ' ' + p.avgStars + ' (' + p.reviewCount + ')' })]) : null,
          h('div', { class: 'row between', style: 'margin-top:auto' }, [
            h('div', { class: 'price', style: 'font-size:17px', text: rs(p.price) }),
            h('button', { type: 'button', class: 'iconbtn', style: 'width:44px;height:44px;background:#E3EEE9;color:#1E4A40', disabled: p.stock < 1, 'aria-label': 'Add ' + p.name + ' to cart',
              onclick: function () { if (add(p, 1)) { toast('Added to your cart.'); render(); } } }, [ic('plus', 20)])
          ])
        ])
      ]);
    })) : h('p', { class: 'pad muted', text: 'New pieces are on their way. Please check back soon.' }),
    h('p', { class: 'small muted', style: 'padding:16px 20px 20px', text: shipLine() + ' Pay by UPI; cash is accepted only when you collect in person.' }),
    t.count ? h('div', { class: 'bar' }, [h('button', { type: 'button', class: 'btn', style: 'justify-content:space-between', onclick: function () { go('cart'); } },
      [h('span', { text: t.count + (t.count === 1 ? ' item · ' : ' items · ') + rs(t.sub) }), h('span', { text: 'View cart' })])]) : null
  ];
}
function shipLine() {
  return 'Collect in person or have it delivered within ' + S.shipArea + '. Delivery is ' + rs(S.shippingFee) + ' on orders under ' + rs(S.freeFrom) + ' and free above that.';
}

/* ---------------------------------------------------------------- product */
function viewProduct() {
  var p = product(st.pid);
  if (!p) { st.view = 'shop'; return viewShop(); }
  var inCart = st.cart[p.id] || 0, max = Math.max(0, p.stock - inCart);
  if (st.qty > max) st.qty = Math.max(1, max);
  var addBtn = h('button', { type: 'button', class: 'btn', disabled: max < 1, text: max < 1 ? (p.stock < 1 ? 'Sold out' : 'All of these are in your cart') : 'Add to cart · ' + rs(p.price * st.qty),
    onclick: function () { if (add(p, st.qty)) go('cart'); } });
  return [
    topbar('', function () { go('shop'); }, cartButton()),
    h('div', { class: 'pad' }, [h('div', { style: 'border-radius:18px;overflow:hidden' }, [photo(p.photos[st.shot] || p.photos[0], p.name, 'sq', 900)])]),
    p.photos.length > 1 ? h('div', { class: 'row pad', style: 'gap:8px;padding-top:12px' }, p.photos.map(function (id, i) {
      var b = h('button', { type: 'button', 'aria-label': 'Photo ' + (i + 1), 'aria-pressed': String(i === st.shot), style: 'padding:0;background:none;border-radius:10px;overflow:hidden;width:60px;height:60px;cursor:pointer;border:2px solid ' + (i === st.shot ? '#1E4A40' : '#D5DEDA'),
        onclick: function () { go('product', { shot: i }); } }, [photo(id, '', 'sq', 200)]);
      return b;
    })) : null,
    h('div', { class: 'stack pad', style: 'padding-top:20px;gap:8px' }, [
      p.category ? h('div', { class: 'small muted', text: p.category }) : null,
      h('h1', { style: 'font-size:26px;line-height:32px', text: p.name }),
      h('div', { class: 'row' }, [h('div', { class: 'price', style: 'font-size:24px', text: rs(p.price) }), stockTag(p) || h('span', { class: 'tag', text: 'In stock' })]),
      p.avgStars != null ? h('a', { href: '#reviews', class: 'small row', style: 'gap:6px;text-decoration:none;color:#16211E;min-height:24px' }, [starRow(p.avgStars, 16), h('span', { text: p.avgStars + ' · ' + p.reviewCount + ' reviews' })]) : null,
      p.description ? h('p', { style: 'font-size:15px;line-height:23px;color:#2D3A36;margin-top:8px', text: p.description }) : null,
      h('p', { class: 'small muted', text: 'Each piece is made by hand, so colours and details differ a little from the photo.' })
    ]),
    h('div', { class: 'card small', style: 'margin:20px 20px 0' }, [h('div', { text: shipLine() })]),
    reviewsBlock(p),
    h('div', { style: 'height:20px' }),
    h('div', { class: 'bar', style: 'flex-direction:row;align-items:center;gap:12px' }, [
      max > 0 ? stepper(st.qty, 1, max, function (v) { st.qty = v; addBtn.textContent = 'Add to cart · ' + rs(p.price * v); }, 'Quantity') : null,
      addBtn
    ])
  ];
}

function starRow(n, size) {
  var full = Math.round(n), wrap = h('span', { role: 'img', 'aria-label': n + ' out of 5 stars', style: 'display:inline-flex;gap:1px;vertical-align:-2px' });
  for (var i = 1; i <= 5; i++) wrap.appendChild(starIcon(i <= full, size || 16));
  return wrap;
}
function starIcon(on, size) {
  var ns = 'http://www.w3.org/2000/svg', s = document.createElementNS(ns, 'svg'), p = document.createElementNS(ns, 'path');
  s.setAttribute('width', size); s.setAttribute('height', size); s.setAttribute('viewBox', '0 0 24 24'); s.setAttribute('aria-hidden', 'true');
  p.setAttribute('d', 'M12 3l2.7 5.6 6.1.9-4.4 4.3 1 6.1L12 17l-5.4 2.9 1-6.1-4.4-4.3 6.1-.9z');
  p.setAttribute('fill', on ? '#B8730A' : 'none'); p.setAttribute('stroke', on ? '#B8730A' : '#8A9692'); p.setAttribute('stroke-width', '1.6'); p.setAttribute('stroke-linejoin', 'round');
  s.appendChild(p); return s;
}
function reviewsBlock(p) {
  if (!S.reviewsOn) return null;
  var head = p.reviewCount ? (p.reviewCount === 1 ? '1 review' : p.reviewCount + ' reviews') : 'Reviews';
  return h('div', { id: 'reviews', class: 'stack pad', style: 'padding-top:24px;gap:10px' }, [
    h('h2', { text: head }),
    p.reviewCount ? null : h('p', { class: 'small muted', text: 'No reviews yet. Buyers are invited to review after their order is complete.' })
  ].concat(p.reviews.map(function (r) {
    return h('div', { class: 'card stack', style: 'gap:6px' }, [
      h('div', { class: 'row between' }, [starRow(r.stars, 16), h('span', { class: 'tiny muted', text: r.date })]),
      r.text ? h('p', { style: 'font-size:15px;line-height:22px', text: r.text }) : null,
      h('div', { class: 'tiny muted', text: r.by + ' · Verified purchase' }),
      r.reply ? h('div', { class: 'small', style: 'margin-top:4px;padding:10px 12px;border-radius:10px;background:#F3F6F4' }, [h('div', { style: 'font-weight:600', text: 'Reply from ' + S.name }), h('div', { text: r.reply })]) : null
    ]);
  })));
}

/* ----------------------------------------------------------------- review */
function viewReview() {
  var R = st.review;
  if (!R) return [topbar('Write a review', function () { go('shop'); }), h('p', { class: 'pad muted', text: 'We could not find that order. Please use the link in your email.' })];
  var out = [topbar('Write a review', function () { go('shop'); }),
    h('p', { class: 'small muted pad', text: 'Order ' + R.id + '. Review each item you bought. It is optional, and you can review one item and skip the others.' })];
  if (!R.open) {
    out.push(h('p', { class: 'pad', style: 'padding-top:12px', text: R.status === 'completed' ? 'Reviews are not open for this order.' : 'You can review this order once it has been delivered or collected.' }));
    return out;
  }
  R.items.forEach(function (it) {
    var f = st.rv[it.id] = st.rv[it.id] || { stars: 0, text: '', consent: false, anon: false };
    if (it.done) {
      out.push(h('div', { class: 'card row', style: 'margin:12px 20px 0' }, [photo(it.photo, it.name, 'thumb', 200),
        h('div', { class: 'grow' }, [h('div', { style: 'font-weight:600', text: it.name }), h('div', { class: 'small muted', text: it.done === 'published' ? 'Thank you. Your review is in the shop.' : 'Thank you. Your review will appear once it has been checked.' })])]));
      return;
    }
    var starsBox = h('div', { role: 'radiogroup', 'aria-label': 'Stars for ' + it.name, class: 'row', style: 'gap:2px' });
    function drawStars() {
      starsBox.textContent = '';
      for (var i = 1; i <= 5; i++) (function (n) {
        starsBox.appendChild(h('button', { type: 'button', role: 'radio', 'aria-checked': String(f.stars === n), 'aria-label': n + (n === 1 ? ' star' : ' stars'),
          style: 'width:44px;height:44px;border:none;background:none;padding:0;cursor:pointer;display:flex;align-items:center;justify-content:center',
          onclick: function () { f.stars = n; drawStars(); } }, [starIcon(n <= f.stars, 30)]));
      })(i);
    }
    drawStars();
    var ta = h('textarea', { rows: 4, maxlength: 1000, placeholder: 'What did you like? How did it look, smell, burn?', oninput: function (e) { f.text = e.target.value; } });
    ta.value = f.text;
    var cid = 'c-' + it.id, aid = 'a-' + it.id;
    var consent = h('input', { type: 'checkbox', id: cid, style: 'width:24px;height:24px;flex-shrink:0;accent-color:#1E4A40', onchange: function (e) { f.consent = e.target.checked; } });
    consent.checked = f.consent;
    var anon = h('input', { type: 'checkbox', id: aid, style: 'width:24px;height:24px;flex-shrink:0;accent-color:#1E4A40', onchange: function (e) { f.anon = e.target.checked; } });
    anon.checked = f.anon;
    out.push(h('div', { class: 'card stack', style: 'margin:12px 20px 0' }, [
      h('div', { class: 'row' }, [photo(it.photo, it.name, 'thumb', 200), h('div', { class: 'grow', style: 'font-weight:600', text: it.name })]),
      h('div', { class: 'small', style: 'font-weight:500', text: 'Your rating' }), starsBox,
      field('Your review (optional)', ta),
      h('div', { class: 'row', style: 'align-items:flex-start;gap:12px' }, [anon, h('label', { for: aid, class: 'small', text: 'Show my review as "A verified buyer" instead of "' + R.shortName + '"' })]),
      h('div', { class: 'row', style: 'align-items:flex-start;gap:12px' }, [consent, h('label', { for: cid, class: 'small', text: R.consentText })]),
      h('button', { type: 'button', class: 'btn', text: 'Send review', onclick: function () {
        if (!f.stars) return toast('Please choose from 1 to 5 stars.');
        if (!f.consent) return toast('Please tick the box to agree before sending your review.');
        call('apiSubmitReview', [R.token, { productId: it.id, stars: f.stars, text: f.text, consent: true, anonymous: f.anon }], function (x) {
          st.review = x; delete st.rv[it.id]; render(); toast('Thank you for your review.');
        });
      } })
    ]));
  });
  out.push(h('p', { class: 'tiny muted', style: 'padding:16px 20px 24px', text: R.rules }));
  out.push(h('div', { class: 'pad', style: 'padding-bottom:24px' }, [h('button', { type: 'button', class: 'btn sec', text: 'Back to the shop', onclick: function () { go('shop'); } })]));
  return out;
}

/* ------------------------------------------------------------------- cart */
function viewCart() {
  var ls = lines(), f = st.form;
  saveCart();
  if (!ls.length) {
    return [topbar('Your cart', function () { go('shop'); }), h('div', { class: 'stack pad', style: 'padding-top:24px' }, [
      h('p', { class: 'muted', text: 'Your cart is empty.' }), h('button', { type: 'button', class: 'btn sec', text: 'Back to the shop', onclick: function () { go('shop'); } })])];
  }
  if (f.mode === 'ship') f.pay = 'upi';
  if (f.mode === 'pickup' && !S.pickupAddress) f.mode = 'ship';
  var t = totals();
  function input(key, attrs) {
    attrs.value = f[key];
    attrs.oninput = function (e) { f[key] = e.target.value; };
    return h('input', attrs);
  }
  var city = h('select', { onchange: function (e) { f.city = e.target.value; } },
    [h('option', { value: '', text: 'Choose your city' })].concat(S.cities.map(function (c) { return h('option', { value: c, text: c, selected: c === f.city }); })));
  var consent = h('input', { type: 'checkbox', id: 'consent', style: 'width:24px;height:24px;flex-shrink:0;accent-color:#1E4A40', onchange: function (e) { f.consent = e.target.checked; } });
  consent.checked = f.consent;
  var cta = f.pay === 'upi' ? 'Continue to pay ' + rs(t.total) : 'Place order · ' + rs(t.total) + ' cash at pickup';
  return [
    topbar('Your cart', function () { go('shop'); }),
    h('div', { class: 'stack pad', style: 'padding-top:8px;gap:10px' }, ls.map(function (l) {
      return h('div', { class: 'card row', style: 'padding:12px;align-items:flex-start' }, [
        photo(l.p.photos[0], l.p.name, 'thumb', 200),
        h('div', { class: 'grow stack', style: 'gap:6px' }, [
          h('div', { class: 'row between', style: 'align-items:flex-start' }, [
            h('div', { style: 'font-weight:500;line-height:20px', text: l.p.name }),
            h('button', { type: 'button', class: 'iconbtn', style: 'width:44px;height:44px;margin:-10px -10px 0 0', 'aria-label': 'Remove ' + l.p.name, onclick: function () { delete st.cart[l.p.id]; render(); } }, [ic('x', 18)])
          ]),
          h('div', { class: 'row between' }, [
            stepper(l.qty, 1, l.p.stock, function (v) { st.cart[l.p.id] = v; render(); }, 'Quantity of ' + l.p.name),
            h('div', { class: 'price', style: 'font-size:17px', text: rs(l.p.price * l.qty) })
          ])
        ])
      ]);
    })),

    h('h2', { style: 'margin:24px 20px 10px', text: 'How would you like to get it?' }),
    h('div', { class: 'grid2 pad', style: 'gap:10px' }, [
      S.pickupAddress ? h('button', { type: 'button', class: 'opt' + (f.mode === 'pickup' ? ' on' : ''), 'aria-pressed': String(f.mode === 'pickup'), onclick: function () { f.mode = 'pickup'; render(); } },
        [h('b', { text: 'Collect in person' }), h('span', { text: 'Free. UPI or cash.' })]) : null,
      h('button', { type: 'button', class: 'opt' + (f.mode === 'ship' ? ' on' : ''), 'aria-pressed': String(f.mode === 'ship'), onclick: function () { f.mode = 'ship'; f.pay = 'upi'; render(); } },
        [h('b', { text: 'Deliver to me' }), h('span', { text: S.shipArea + ' only. ' + (t.sub < S.freeFrom ? rs(S.shippingFee) + ' delivery.' : 'Free delivery.') + ' UPI only.' })])
    ]),
    f.mode === 'pickup' ? h('div', { class: 'card small', style: 'margin:12px 20px 0' }, [
      h('div', { style: 'font-weight:600', text: 'Collect from' }), h('div', { text: S.pickupAddress }),
      h('div', { class: 'muted', style: 'margin-top:6px', text: 'We will email you when it is confirmed. Cash orders are kept aside for ' + S.pickupHoldHours + ' hours.' })
    ]) : h('div', { class: 'stack pad', style: 'padding-top:12px' }, [
      field('House / flat, street, area', input('address', { type: 'text', autocomplete: 'street-address', maxlength: 240 })),
      h('div', { class: 'grid2', style: 'gap:10px' }, [
        field('City', city),
        field('PIN code', input('pin', { type: 'text', inputmode: 'numeric', autocomplete: 'postal-code', maxlength: 6 }))
      ]),
      h('div', { class: 'small muted', text: 'We deliver only within ' + S.shipArea + ' for now.' })
    ]),

    h('h2', { style: 'margin:24px 20px 10px', text: 'Your details' }),
    h('div', { class: 'stack pad' }, [
      field('Full name', input('name', { type: 'text', autocomplete: 'name', maxlength: 80 })),
      field('Mobile number', input('phone', { type: 'tel', autocomplete: 'tel', maxlength: 16 })),
      field('Email, for your order updates', input('email', { type: 'email', autocomplete: 'email', maxlength: 120 }))
    ]),

    h('h2', { style: 'margin:24px 20px 10px', text: 'How will you pay?' }),
    h('div', { class: 'stack pad', style: 'gap:10px' }, [
      h('button', { type: 'button', class: 'opt' + (f.pay === 'upi' ? ' on' : ''), 'aria-pressed': String(f.pay === 'upi'), onclick: function () { f.pay = 'upi'; render(); } },
        [h('b', { text: 'UPI, pay now' }), h('span', { text: 'GPay, PhonePe, Paytm or any UPI app' })]),
      f.mode === 'pickup' ? h('button', { type: 'button', class: 'opt' + (f.pay === 'cash' ? ' on' : ''), 'aria-pressed': String(f.pay === 'cash'), onclick: function () { f.pay = 'cash'; render(); } },
        [h('b', { text: 'Cash, when I collect' }), h('span', { text: 'Pay at pickup' })])
        : h('div', { class: 'small muted', style: 'padding:12px 14px;border:1px dashed #B9C6C1;border-radius:14px', text: 'Cash is available only when you collect in person. Delivered orders are paid by UPI first; there is no cash on delivery.' })
    ]),

    h('div', { class: 'card stack', style: 'margin:24px 20px 0;gap:10px' }, [
      h('div', { class: 'sum' }, [h('span', { text: 'Items' }), h('span', { text: rs(t.sub) })]),
      h('div', { class: 'sum' }, [h('span', { text: f.mode === 'ship' ? 'Delivery' : 'Collect in person' }), h('span', { text: t.ship ? rs(t.ship) : 'Free' })]),
      t.ship ? h('div', { class: 'note', text: 'Add ' + rs(S.freeFrom - t.sub) + ' more for free delivery.' }) : null,
      h('div', { style: 'height:1px;background:#E3E9E6' }),
      h('div', { class: 'sum price', style: 'font-size:19px' }, [h('span', { text: 'Total' }), h('span', { text: rs(t.total) })])
    ]),

    h('div', { class: 'row pad', style: 'margin:16px 0 20px;align-items:flex-start;gap:12px' }, [
      consent,
      h('div', { class: 'small' }, [
        h('label', { for: 'consent', text: S.consentText + ' ' }),
        h('button', { type: 'button', style: 'border:none;background:none;padding:0;color:#1E4A40;text-decoration:underline;font:inherit;cursor:pointer;min-height:24px', text: 'Privacy notice',
          onclick: function () { var close = sheet([h('h2', { text: 'Privacy notice' }), h('p', { class: 'small pre', text: S.privacyText }), h('button', { type: 'button', class: 'btn sec', text: 'Close', onclick: function () { close(); } })]); } })
      ])
    ]),

    h('div', { class: 'bar' }, [
      h('button', { type: 'button', class: 'btn', text: cta, onclick: placeOrder }),
      h('div', { class: 'tiny muted', style: 'text-align:center', text: f.pay === 'upi' ? 'Your items are held for ' + S.holdMinutes + ' minutes while you pay.' : 'Your items are kept aside for ' + S.pickupHoldHours + ' hours.' })
    ])
  ];
}
function placeOrder() {
  var f = st.form;
  var pin = String(f.pin).replace(/\D/g, '');
  if (f.mode === 'ship') {
    if (String(f.address).trim().length < 8) return toast('Please enter your full delivery address.');
    if (!f.city) return toast('Please choose your city.');
    if (pin.length !== 6 || (S.pinPrefixes.length && !S.pinPrefixes.some(function (p) { return pin.indexOf(p) === 0; }))) return toast('That PIN code is outside ' + S.shipArea + '. We deliver only within ' + S.shipArea + ' for now.');
  }
  if (String(f.name).trim().length < 2) return toast('Please enter your full name.');
  if (!/^\+?\d{8,15}$/.test(String(f.phone).replace(/[\s\-().]/g, ''))) return toast('Please enter a valid mobile number.');
  if (!/^[^\s@]+@[^\s@]+\.[A-Za-z]{2,}$/.test(String(f.email).trim())) return toast('Please enter a valid email address.');
  if (!f.consent) return toast('Please tick the box to agree before placing your order.');
  var payload = { items: lines().map(function (l) { return { id: l.p.id, qty: l.qty }; }), mode: f.mode, pay: f.pay, name: f.name, phone: f.phone, email: f.email, address: f.address, city: f.city, pin: pin, consent: true };
  call('apiPlaceOrder', [payload], function (r) {
    if (!r.ok) { if (r.products) { PRODUCTS = r.products; render(); } toast(r.error); return; }
    r.order.items.forEach(function (l) { var p = product(l.id); if (p) p.stock = Math.max(0, p.stock - l.qty); });
    st.cart = {}; saveCart(); st.proof = { ref: '', img: null };
    openOrder(r.order);
  });
}

/* -------------------------------------------------------------------- pay */
function amountCard(o, sub) {
  return h('div', { style: 'margin:8px 20px 0;padding:18px;border-radius:14px;background:#1E4A40;color:#fff' }, [
    h('div', { class: 'small', style: 'color:#D6E6DF', text: 'Order ' + o.id + ' · ' + (o.mode === 'ship' ? 'Deliver to me' : 'Collect in person') }),
    h('div', { class: 'price', style: 'font-size:34px;line-height:40px', text: rs(o.total) }),
    h('div', { class: 'small', style: 'color:#D6E6DF', text: sub })
  ]);
}
function viewPay() {
  var o = st.order;
  var qr = qrCanvas(o.upiLink, 400);
  pollTimer = setInterval(function () { call('apiGetOrder', [o.token], function (x) { if (x.status !== 'awaiting_payment') openOrder(x); }, { quiet: true }); }, 60000);
  return [
    topbar('Pay for your order', null),
    amountCard(o, rs(o.subtotal) + ' items' + (o.shipping ? ' + ' + rs(o.shipping) + ' delivery' : '') + '. Held for you until ' + o.holdText + '.'),
    h('h2', { style: 'margin:20px 20px 12px', text: 'Step 1 of 2: pay ' + rs(o.total) + ' by UPI' }),
    h('div', { class: 'card stack', style: 'margin:0 20px;gap:14px' }, [
      h('a', { class: 'btn', href: o.upiLink, text: 'Open my UPI app' }),
      qr ? h('div', { class: 'small muted', style: 'text-align:center', text: 'or scan this code from another phone' }) : null,
      qr ? h('div', { style: 'display:flex;justify-content:center' }, [qr]) : null,
      h('div', { class: 'row between', style: 'padding:4px 4px 4px 14px;border-radius:12px;background:#F3F6F4' }, [
        h('div', { class: 'grow' }, [h('div', { class: 'tiny muted', text: 'UPI ID' }), h('div', { style: 'font-weight:600;word-break:break-all', text: o.upiId })]),
        h('button', { type: 'button', class: 'btn sec sm fit', text: 'Copy', onclick: function () {
          try { navigator.clipboard.writeText(o.upiId).then(function () { toast('UPI ID copied.'); }, function () { toast('Press and hold the UPI ID to copy it.'); }); } catch (e) { toast('Press and hold the UPI ID to copy it.'); }
        } })
      ])
    ]),
    h('div', { class: 'note', style: 'margin:16px 20px 0', text: 'Pay in your UPI app first, then come back here and tap Next. On the next screen you will confirm the payment with its reference number or a screenshot.' }),
    h('div', { style: 'padding:16px 20px 20px' }, [h('button', { type: 'button', class: 'btn warn sm', text: 'Cancel this order', onclick: function () {
      var close = sheet([h('h2', { text: 'Cancel order ' + o.id + '?' }), h('p', { class: 'small muted', text: 'Only cancel if you have not paid. The items go back on sale.' }),
        h('button', { type: 'button', class: 'btn warn', text: 'Yes, cancel the order', onclick: function () { close(); call('apiCancelOrder', [o.token], function (x) { setOrder(x); go('status'); }); } }),
        h('button', { type: 'button', class: 'btn sec', text: 'Keep my order', onclick: function () { close(); } })]);
    } })]),
    h('div', { class: 'bar' }, [
      h('button', { type: 'button', class: 'btn', text: 'Next: confirm my payment', onclick: function () { go('confirm'); } }),
      o.storePhone ? h('div', { class: 'tiny muted', style: 'text-align:center', text: 'Trouble paying? Call or WhatsApp ' + o.storePhone }) : null
    ])
  ];
}

/* ---------------------------------------------------------------- confirm */
function viewConfirm() {
  var o = st.order, pr = st.proof;
  var btn = h('button', { type: 'button', class: 'btn', text: 'I have paid ' + rs(o.total), onclick: function () {
    if (btn.disabled) return;
    call('apiSubmitProof', [o.token, pr.ref, pr.img], function (x) { st.proof = { ref: '', img: null }; setOrder(x); go('status'); });
  } });
  var hint = h('div', { class: 'tiny muted', style: 'text-align:center' });
  function check() {
    var n = pr.ref.length, ok = (n === 12) || (n === 0 && !!pr.img);
    btn.disabled = !ok;
    hint.textContent = ok ? 'Tap only after the money has left your account.' : (n > 0 && n !== 12 ? 'The reference number has 12 digits (' + n + ' so far).' : 'Enter the reference number or add a screenshot to continue.');
  }
  var shotBox = h('div');
  function drawShot() {
    shotBox.textContent = '';
    var pick = h('input', { type: 'file', accept: 'image/*', style: 'position:absolute;opacity:0;width:1px;height:1px', onchange: function (e) {
      var file = e.target.files && e.target.files[0];
      if (file) shrink(file, 1400, false, 0.72, function (d) { pr.img = d; drawShot(); check(); });
    } });
    if (pr.img) {
      shotBox.appendChild(h('div', { class: 'row between', style: 'padding:10px 12px;border-radius:12px;background:#E3EEE9' }, [
        h('img', { src: pr.img, alt: 'Your payment screenshot', style: 'width:52px;height:52px;object-fit:cover;border-radius:8px' }),
        h('div', { class: 'grow', style: 'font-weight:600;color:#1E4A40;font-size:14px', text: 'Screenshot added' }),
        h('label', { class: 'btn sec sm fit', style: 'position:relative' }, ['Change', pick])
      ]));
    } else {
      shotBox.appendChild(h('label', { class: 'btn sec', style: 'position:relative;min-height:84px;border:2px dashed #B9C6C1;background:#F3F6F4' }, [ic('up'), 'Choose screenshot', pick]));
    }
  }
  drawShot(); check();
  return [
    topbar('Confirm your payment', o.status === 'awaiting_payment' ? function () { go('pay'); } : function () { go('status'); }),
    h('div', { class: 'card row between', style: 'margin:8px 20px 0' }, [
      h('div', null, [h('div', { class: 'small muted', text: 'Order ' + o.id + ' · to ' + o.upiId }), h('div', { class: 'price', style: 'font-size:22px', text: rs(o.total) })])
    ]),
    h('h2', { style: 'margin:20px 20px 4px', text: 'Step 2 of 2: show us your payment' }),
    h('p', { class: 'small muted pad', text: 'Give either one. The reference number is easiest if you paid from another phone.' }),
    h('div', { class: 'card stack', style: 'margin:12px 20px 0' }, [
      field('12-digit UPI reference number (UTR / UPI transaction ID)', h('input', { type: 'text', inputmode: 'numeric', maxlength: 14, value: pr.ref, placeholder: 'e.g. 412345678901',
        oninput: function (e) { pr.ref = e.target.value.replace(/\D/g, '').slice(0, 12); if (e.target.value !== pr.ref) e.target.value = pr.ref; check(); } })),
      h('div', { class: 'small muted', style: 'text-align:center', text: 'or' }),
      shotBox
    ]),
    h('p', { class: 'small muted', style: 'padding:16px 20px 20px', text: 'We check each payment by hand before confirming. If no payment is found, the order is cancelled and the items go back on sale.' }),
    h('div', { class: 'bar' }, [btn, hint])
  ];
}

/* ----------------------------------------------------------------- status */
function viewStatus() {
  var o = st.order, s = o.status, title, body = [], extra = [];
  function p(text, cls) { return h('p', { class: cls || '', text: text }); }
  var pickup = h('div', { class: 'card small' }, [h('div', { style: 'font-weight:600', text: 'Collect from' }), h('div', { text: o.pickupAddress })]);
  if (s === 'awaiting_payment') { st.view = 'pay'; return viewPay(); }
  if (s === 'payment_check') {
    title = 'Order reserved';
    body = [p(o.message || 'Your order is reserved, you\'ll receive an email when the order is confirmed.'), p('We are checking your payment of ' + rs(o.total) + '. This page updates by itself.', 'small muted'),
      h('button', { type: 'button', class: 'btn sec sm', text: 'Change my payment proof', onclick: function () { go('confirm'); } })];
  } else if (s === 'confirmed') {
    title = 'Order confirmed';
    body = [p('Payment received. Thank you, ' + o.name + '.'), o.mode === 'ship' ? p('We are getting your order ready. We will email you when it is on its way.', 'small muted') : pickup];
  } else if (s === 'cash_pickup') {
    title = 'Order reserved';
    body = [p('Your order is reserved. Please collect it and pay ' + rs(o.total) + ' in cash by ' + o.holdText + '.'), pickup, p('After that the items go back on sale.', 'small muted')];
    extra = [h('button', { type: 'button', class: 'btn warn sm', text: 'Cancel this order', onclick: function () { call('apiCancelOrder', [o.token], function (x) { setOrder(x); render(); }); } })];
  } else if (s === 'shipped') {
    title = 'On its way';
    var d = o.ship || {}, rows = [];
    if (d.provider) rows.push(h('div', { text: 'Delivery by: ' + d.provider }));
    if (d.driverName) rows.push(h('div', { text: 'Driver: ' + d.driverName }));
    if (d.driverPhone) rows.push(h('div', null, ['Driver phone: ', h('a', { href: 'tel:' + d.driverPhone, text: d.driverPhone })]));
    if (d.note) rows.push(h('div', { text: d.note }));
    body = [p('Your order was sent out on ' + d.at + '.'), rows.length ? h('div', { class: 'card small stack', style: 'gap:6px' }, rows) : null,
      d.tracking ? h('a', { class: 'btn', href: d.tracking, target: '_blank', rel: 'noopener', text: 'Track my delivery' }) : null];
  } else if (s === 'completed') {
    title = 'Thank you';
    body = [p('This order is complete. We hope you love it.'),
      o.canReview ? h('button', { type: 'button', class: 'btn', text: 'Review what you bought', onclick: function () { call('apiReviewInfo', [o.token], function (x) { st.review = x; go('review'); }); } }) : null];
  } else {
    title = 'Order cancelled';
    body = [p(o.message || ('This order was cancelled.' + (o.cancelReason ? ' Reason: ' + o.cancelReason + '.' : ''))),
      o.refundDue ? p('Your payment will be refunded to you by UPI.', 'small muted') : null,
      o.canAddProof ? p('If you have already paid, add your payment proof and we will check it.', 'small muted') : null,
      o.canAddProof ? h('button', { type: 'button', class: 'btn sec sm', text: 'I have already paid', onclick: function () { go('confirm'); } }) : null];
  }
  if (s === 'payment_check' || s === 'confirmed' || s === 'cash_pickup' || s === 'shipped') {
    pollTimer = setInterval(function () { call('apiGetOrder', [o.token], function (x) { if (x.status !== o.status || JSON.stringify(x.ship) !== JSON.stringify(o.ship)) { setOrder(x); render(); } }, { quiet: true }); }, 60000);
  }
  return [
    topbar(title, null),
    h('div', { class: 'stack pad', style: 'padding-top:8px' }, body),
    h('div', { class: 'card stack', style: 'margin:16px 20px 0;gap:8px' }, [h('div', { class: 'small muted', text: 'Order ' + o.id })].concat(o.items.map(function (l) {
      return h('div', { class: 'sum' }, [h('span', { text: l.qty + ' × ' + l.name }), h('span', { text: rs(l.price * l.qty) })]);
    })).concat([
      o.shipping ? h('div', { class: 'sum' }, [h('span', { text: 'Delivery' }), h('span', { text: rs(o.shipping) })]) : null,
      h('div', { style: 'height:1px;background:#E3E9E6' }),
      h('div', { class: 'sum price', style: 'font-size:17px' }, [h('span', { text: 'Total' }), h('span', { text: rs(o.total) })])
    ])),
    h('div', { class: 'stack pad', style: 'padding-top:16px;padding-bottom:24px' }, extra.concat([
      h('button', { type: 'button', class: 'btn sec', text: 'Back to the shop', onclick: function () { call('apiBoot', [], function (b) { PRODUCTS = b.products; S = b.store; go('shop'); }, { fail: function () { go('shop'); return true; } }); } }),
      o.storePhone ? h('div', { class: 'tiny muted', style: 'text-align:center', text: 'Questions? Call or WhatsApp ' + o.storePhone }) : null
    ]))
  ];
}

if (BOOT.reviewAsked) st.view = 'review';
else if (st.order) { setOrder(st.order); st.view = st.order.status === 'awaiting_payment' ? 'pay' : 'status'; }
render();
</script>
</body>
</html>
`;

/* ----------------------------------------------------------- admin page */

var ADMIN_HTML = String.raw`<!doctype html>
<html lang="en">
<head>
<base target="_top">
<meta charset="utf-8">
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600&family=Outfit:wght@500;600&display=swap">
<style>
*{box-sizing:border-box}html,body{margin:0}
body{background:#F3F6F4;color:#16211E;font-family:Inter,system-ui,sans-serif;font-size:15px;line-height:1.4;-webkit-text-size-adjust:100%}
#app{max-width:480px;margin:0 auto;min-height:100vh;display:flex;flex-direction:column}
h1,h2,h3,p{margin:0}
h1{font:600 22px/28px Outfit,sans-serif}h2{font:600 17px/22px Outfit,sans-serif}
.price{font-family:Outfit,sans-serif;font-weight:600}
.muted{color:#55635F}.small{font-size:13px;line-height:19px}.tiny{font-size:12px;line-height:17px}
.pad{padding:0 20px}.stack{display:flex;flex-direction:column;gap:12px}
.row{display:flex;align-items:center;gap:10px}.between{justify-content:space-between}.grow{flex:1;min-width:0}.wrap{flex-wrap:wrap}
.card{background:#fff;border:1px solid #E3E9E6;border-radius:14px;padding:14px}
.btn{min-height:52px;border-radius:14px;border:none;background:#1E4A40;color:#fff;font:600 16px Inter,sans-serif;padding:8px 16px;display:flex;align-items:center;justify-content:center;gap:8px;text-decoration:none;cursor:pointer;text-align:center;width:100%}
.btn.sec{background:#fff;color:#16211E;border:1px solid #B9C6C1}
.btn.sm{min-height:46px;font-size:14px;border-radius:12px}
.btn[disabled]{background:#D5DEDA;color:#3F4C48;cursor:default}
.btn.warn{background:#fff;color:#8E2A1E;border:1px solid #E2B8B0}
.btn.fit{width:auto}
.iconbtn{width:48px;height:48px;border-radius:24px;border:none;background:none;display:flex;align-items:center;justify-content:center;cursor:pointer;color:#16211E;position:relative;flex-shrink:0}
.iconbtn.ring{background:#fff;border:1px solid #D5DEDA}
.badge{position:absolute;top:-4px;right:-4px;min-width:20px;height:20px;border-radius:10px;background:#1E4A40;color:#fff;font-size:12px;font-weight:600;display:flex;align-items:center;justify-content:center;padding:0 5px}
.chip{min-height:44px;padding:0 16px;border-radius:22px;border:1px solid #D5DEDA;background:#fff;color:#16211E;font:500 14px Inter,sans-serif;cursor:pointer}
.chip.on{background:#1E4A40;color:#fff;border-color:#1E4A40}
.tag{display:inline-block;white-space:nowrap;padding:3px 8px;border-radius:8px;font-size:12px;font-weight:600;background:#E3EEE9;color:#1E4A40}
.tag.amber{background:#FDF1DC;color:#7A4206}.tag.red{background:#FBE9E6;color:#8E2A1E}
.field{display:flex;flex-direction:column;gap:6px;font-size:13px;font-weight:500}
.field input,.field select,.field textarea{height:48px;padding:0 14px;border-radius:12px;border:1px solid #B9C6C1;background:#fff;font:16px Inter,sans-serif;color:#16211E;width:100%}
.field textarea{height:auto;padding:12px 14px;resize:none}
.opt{min-height:60px;padding:10px 14px;border-radius:14px;border:2px solid #D5DEDA;background:#fff;color:#16211E;font-family:Inter,sans-serif;text-align:left;cursor:pointer;display:flex;flex-direction:column;justify-content:center;gap:2px;width:100%}
.opt.on{border-color:#1E4A40;background:#E3EEE9}.opt b{font-size:15px;font-weight:600}.opt span{font-size:13px;line-height:18px;color:#55635F}
.stepper{display:flex;align-items:center;border:1px solid #D5DEDA;border-radius:12px;height:46px;background:#fff;flex-shrink:0}
.stepper button{width:44px;height:44px;border:none;background:none;cursor:pointer;display:flex;align-items:center;justify-content:center;color:#16211E}
.stepper span{min-width:24px;text-align:center;font-weight:600;font-size:16px}
.bar{position:sticky;bottom:0;background:#fff;border-top:1px solid #D5DEDA;padding:12px 20px 16px;display:flex;flex-direction:column;gap:8px;margin-top:auto;z-index:5}
.note{padding:10px 12px;border-radius:12px;background:#FDF1DC;color:#7A4206;font-size:13px;line-height:19px}
.ph{background:#E3E9E6;display:flex;align-items:center;justify-content:center;color:#55635F;font-size:12px}
img.sq,.ph.sq{width:100%;aspect-ratio:1/1;object-fit:cover;display:block}
.thumb{width:64px;height:64px;border-radius:10px;object-fit:cover;flex-shrink:0}
.grid2{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px}
.grid3{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px}
.sum{display:flex;justify-content:space-between;gap:12px}
#toast{position:fixed;left:16px;right:16px;bottom:100px;max-width:448px;margin:0 auto;background:#16211E;color:#fff;padding:12px 16px;border-radius:12px;font-size:14px;line-height:20px;display:none;z-index:70}
#busy{position:fixed;inset:0;background:rgba(243,246,244,.6);display:none;align-items:center;justify-content:center;z-index:80}
#busy div{width:36px;height:36px;border-radius:50%;border:4px solid #D5DEDA;border-top-color:#1E4A40;animation:sp 1s linear infinite}
@keyframes sp{to{transform:rotate(360deg)}}
.sheet{position:fixed;inset:0;background:rgba(22,33,30,.45);z-index:40;display:flex;align-items:flex-end;justify-content:center}
.sheet>div{background:#F3F6F4;width:100%;max-width:480px;max-height:92vh;overflow:auto;border-radius:18px 18px 0 0;padding:16px 20px 20px;display:flex;flex-direction:column;gap:12px}
a{color:#1E4A40}
.pre{white-space:pre-line}

.sec-h{margin:22px 20px 8px;font-size:13px;font-weight:600;color:#1E4A40}.sec-h.amber{color:#7A4206}
.nav{position:sticky;bottom:0;background:#fff;border-top:1px solid #D5DEDA;padding:8px 12px 14px;display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:2px;margin-top:auto;z-index:5}
.nav button{min-height:52px;border-radius:12px;border:none;background:none;font:500 13px Inter,sans-serif;color:#16211E;cursor:pointer;padding:0 2px}
.nav button.on{background:#E3EEE9;color:#1E4A40;font-weight:600}
.ocard{margin:0 20px 10px}.ocard.hot{border-color:#F0D9B0}
.tile{background:#fff;border:1px solid #E3E9E6;border-radius:14px;padding:12px}
.tile b{display:block;font:600 20px/26px Outfit,sans-serif}
</style>
</head>
<body>
<div id="app"></div>
<div id="toast" role="status"></div>
<div id="busy"><div></div></div>
<script>
/* QR code generator: qrcode-generator by Kazuhiko Arase, MIT licence */
var qrcode=function(){var t=function(t,r){var e=t,n=g[r],o=null,i=0,a=null,u=[],f={},c=function(t,r){o=function(t){for(var r=new Array(t),e=0;e<t;e+=1){r[e]=new Array(t);for(var n=0;n<t;n+=1)r[e][n]=null}return r}(i=4*e+17),l(0,0),l(i-7,0),l(0,i-7),s(),h(),d(t,r),e>=7&&v(t),null==a&&(a=p(e,n,u)),w(a,r)},l=function(t,r){for(var e=-1;e<=7;e+=1)if(!(t+e<=-1||i<=t+e))for(var n=-1;n<=7;n+=1)r+n<=-1||i<=r+n||(o[t+e][r+n]=0<=e&&e<=6&&(0==n||6==n)||0<=n&&n<=6&&(0==e||6==e)||2<=e&&e<=4&&2<=n&&n<=4)},h=function(){for(var t=8;t<i-8;t+=1)null==o[t][6]&&(o[t][6]=t%2==0);for(var r=8;r<i-8;r+=1)null==o[6][r]&&(o[6][r]=r%2==0)},s=function(){for(var t=B.getPatternPosition(e),r=0;r<t.length;r+=1)for(var n=0;n<t.length;n+=1){var i=t[r],a=t[n];if(null==o[i][a])for(var u=-2;u<=2;u+=1)for(var f=-2;f<=2;f+=1)o[i+u][a+f]=-2==u||2==u||-2==f||2==f||0==u&&0==f}},v=function(t){for(var r=B.getBCHTypeNumber(e),n=0;n<18;n+=1){var a=!t&&1==(r>>n&1);o[Math.floor(n/3)][n%3+i-8-3]=a}for(n=0;n<18;n+=1){a=!t&&1==(r>>n&1);o[n%3+i-8-3][Math.floor(n/3)]=a}},d=function(t,r){for(var e=n<<3|r,a=B.getBCHTypeInfo(e),u=0;u<15;u+=1){var f=!t&&1==(a>>u&1);u<6?o[u][8]=f:u<8?o[u+1][8]=f:o[i-15+u][8]=f}for(u=0;u<15;u+=1){f=!t&&1==(a>>u&1);u<8?o[8][i-u-1]=f:u<9?o[8][15-u-1+1]=f:o[8][15-u-1]=f}o[i-8][8]=!t},w=function(t,r){for(var e=-1,n=i-1,a=7,u=0,f=B.getMaskFunction(r),c=i-1;c>0;c-=2)for(6==c&&(c-=1);;){for(var g=0;g<2;g+=1)if(null==o[n][c-g]){var l=!1;u<t.length&&(l=1==(t[u]>>>a&1)),f(n,c-g)&&(l=!l),o[n][c-g]=l,-1==(a-=1)&&(u+=1,a=7)}if((n+=e)<0||i<=n){n-=e,e=-e;break}}},p=function(t,r,e){for(var n=A.getRSBlocks(t,r),o=b(),i=0;i<e.length;i+=1){var a=e[i];o.put(a.getMode(),4),o.put(a.getLength(),B.getLengthInBits(a.getMode(),t)),a.write(o)}var u=0;for(i=0;i<n.length;i+=1)u+=n[i].dataCount;if(o.getLengthInBits()>8*u)throw"code length overflow. ("+o.getLengthInBits()+">"+8*u+")";for(o.getLengthInBits()+4<=8*u&&o.put(0,4);o.getLengthInBits()%8!=0;)o.putBit(!1);for(;!(o.getLengthInBits()>=8*u||(o.put(236,8),o.getLengthInBits()>=8*u));)o.put(17,8);return function(t,r){for(var e=0,n=0,o=0,i=new Array(r.length),a=new Array(r.length),u=0;u<r.length;u+=1){var f=r[u].dataCount,c=r[u].totalCount-f;n=Math.max(n,f),o=Math.max(o,c),i[u]=new Array(f);for(var g=0;g<i[u].length;g+=1)i[u][g]=255&t.getBuffer()[g+e];e+=f;var l=B.getErrorCorrectPolynomial(c),h=k(i[u],l.getLength()-1).mod(l);for(a[u]=new Array(l.getLength()-1),g=0;g<a[u].length;g+=1){var s=g+h.getLength()-a[u].length;a[u][g]=s>=0?h.getAt(s):0}}var v=0;for(g=0;g<r.length;g+=1)v+=r[g].totalCount;var d=new Array(v),w=0;for(g=0;g<n;g+=1)for(u=0;u<r.length;u+=1)g<i[u].length&&(d[w]=i[u][g],w+=1);for(g=0;g<o;g+=1)for(u=0;u<r.length;u+=1)g<a[u].length&&(d[w]=a[u][g],w+=1);return d}(o,n)};f.addData=function(t,r){var e=null;switch(r=r||"Byte"){case"Numeric":e=M(t);break;case"Alphanumeric":e=x(t);break;case"Byte":e=m(t);break;case"Kanji":e=L(t);break;default:throw"mode:"+r}u.push(e),a=null},f.isDark=function(t,r){if(t<0||i<=t||r<0||i<=r)throw t+","+r;return o[t][r]},f.getModuleCount=function(){return i},f.make=function(){if(e<1){for(var t=1;t<40;t++){for(var r=A.getRSBlocks(t,n),o=b(),i=0;i<u.length;i++){var a=u[i];o.put(a.getMode(),4),o.put(a.getLength(),B.getLengthInBits(a.getMode(),t)),a.write(o)}var g=0;for(i=0;i<r.length;i++)g+=r[i].dataCount;if(o.getLengthInBits()<=8*g)break}e=t}c(!1,function(){for(var t=0,r=0,e=0;e<8;e+=1){c(!0,e);var n=B.getLostPoint(f);(0==e||t>n)&&(t=n,r=e)}return r}())},f.createTableTag=function(t,r){t=t||2;var e="";e+='<table style="',e+=" border-width: 0px; border-style: none;",e+=" border-collapse: collapse;",e+=" padding: 0px; margin: "+(r=void 0===r?4*t:r)+"px;",e+='">',e+="<tbody>";for(var n=0;n<f.getModuleCount();n+=1){e+="<tr>";for(var o=0;o<f.getModuleCount();o+=1)e+='<td style="',e+=" border-width: 0px; border-style: none;",e+=" border-collapse: collapse;",e+=" padding: 0px; margin: 0px;",e+=" width: "+t+"px;",e+=" height: "+t+"px;",e+=" background-color: ",e+=f.isDark(n,o)?"#000000":"#ffffff",e+=";",e+='"/>';e+="</tr>"}return e+="</tbody>",e+="</table>"},f.createSvgTag=function(t,r,e,n){var o={};"object"==typeof arguments[0]&&(t=(o=arguments[0]).cellSize,r=o.margin,e=o.alt,n=o.title),t=t||2,r=void 0===r?4*t:r,(e="string"==typeof e?{text:e}:e||{}).text=e.text||null,e.id=e.text?e.id||"qrcode-description":null,(n="string"==typeof n?{text:n}:n||{}).text=n.text||null,n.id=n.text?n.id||"qrcode-title":null;var i,a,u,c,g=f.getModuleCount()*t+2*r,l="";for(c="l"+t+",0 0,"+t+" -"+t+",0 0,-"+t+"z ",l+='<svg version="1.1" xmlns="http://www.w3.org/2000/svg"',l+=o.scalable?"":' width="'+g+'px" height="'+g+'px"',l+=' viewBox="0 0 '+g+" "+g+'" ',l+=' preserveAspectRatio="xMinYMin meet"',l+=n.text||e.text?' role="img" aria-labelledby="'+y([n.id,e.id].join(" ").trim())+'"':"",l+=">",l+=n.text?'<title id="'+y(n.id)+'">'+y(n.text)+"</title>":"",l+=e.text?'<description id="'+y(e.id)+'">'+y(e.text)+"</description>":"",l+='<rect width="100%" height="100%" fill="white" cx="0" cy="0"/>',l+='<path d="',a=0;a<f.getModuleCount();a+=1)for(u=a*t+r,i=0;i<f.getModuleCount();i+=1)f.isDark(a,i)&&(l+="M"+(i*t+r)+","+u+c);return l+='" stroke="transparent" fill="black"/>',l+="</svg>"},f.createDataURL=function(t,r){t=t||2,r=void 0===r?4*t:r;var e=f.getModuleCount()*t+2*r,n=r,o=e-r;return I(e,e,function(r,e){if(n<=r&&r<o&&n<=e&&e<o){var i=Math.floor((r-n)/t),a=Math.floor((e-n)/t);return f.isDark(a,i)?0:1}return 1})},f.createImgTag=function(t,r,e){t=t||2,r=void 0===r?4*t:r;var n=f.getModuleCount()*t+2*r,o="";return o+="<img",o+=' src="',o+=f.createDataURL(t,r),o+='"',o+=' width="',o+=n,o+='"',o+=' height="',o+=n,o+='"',e&&(o+=' alt="',o+=y(e),o+='"'),o+="/>"};var y=function(t){for(var r="",e=0;e<t.length;e+=1){var n=t.charAt(e);switch(n){case"<":r+="&lt;";break;case">":r+="&gt;";break;case"&":r+="&amp;";break;case'"':r+="&quot;";break;default:r+=n}}return r};return f.createASCII=function(t,r){if((t=t||1)<2)return function(t){t=void 0===t?2:t;var r,e,n,o,i,a=1*f.getModuleCount()+2*t,u=t,c=a-t,g={"██":"█","█ ":"▀"," █":"▄","  ":" "},l={"██":"▀","█ ":"▀"," █":" ","  ":" "},h="";for(r=0;r<a;r+=2){for(n=Math.floor((r-u)/1),o=Math.floor((r+1-u)/1),e=0;e<a;e+=1)i="█",u<=e&&e<c&&u<=r&&r<c&&f.isDark(n,Math.floor((e-u)/1))&&(i=" "),u<=e&&e<c&&u<=r+1&&r+1<c&&f.isDark(o,Math.floor((e-u)/1))?i+=" ":i+="█",h+=t<1&&r+1>=c?l[i]:g[i];h+="\n"}return a%2&&t>0?h.substring(0,h.length-a-1)+Array(a+1).join("▀"):h.substring(0,h.length-1)}(r);t-=1,r=void 0===r?2*t:r;var e,n,o,i,a=f.getModuleCount()*t+2*r,u=r,c=a-r,g=Array(t+1).join("██"),l=Array(t+1).join("  "),h="",s="";for(e=0;e<a;e+=1){for(o=Math.floor((e-u)/t),s="",n=0;n<a;n+=1)i=1,u<=n&&n<c&&u<=e&&e<c&&f.isDark(o,Math.floor((n-u)/t))&&(i=0),s+=i?g:l;for(o=0;o<t;o+=1)h+=s+"\n"}return h.substring(0,h.length-1)},f.renderTo2dContext=function(t,r){r=r||2;for(var e=f.getModuleCount(),n=0;n<e;n++)for(var o=0;o<e;o++)t.fillStyle=f.isDark(n,o)?"black":"white",t.fillRect(o*r,n*r,r,r)},f};t.stringToBytes=(t.stringToBytesFuncs={default:function(t){for(var r=[],e=0;e<t.length;e+=1){var n=t.charCodeAt(e);r.push(255&n)}return r}}).default,t.createStringToBytes=function(t,r){var e=function(){for(var e=S(t),n=function(){var t=e.read();if(-1==t)throw"eof";return t},o=0,i={};;){var a=e.read();if(-1==a)break;var u=n(),f=n()<<8|n();i[String.fromCharCode(a<<8|u)]=f,o+=1}if(o!=r)throw o+" != "+r;return i}(),n="?".charCodeAt(0);return function(t){for(var r=[],o=0;o<t.length;o+=1){var i=t.charCodeAt(o);if(i<128)r.push(i);else{var a=e[t.charAt(o)];"number"==typeof a?(255&a)==a?r.push(a):(r.push(a>>>8),r.push(255&a)):r.push(n)}}return r}};var r,e,n,o,i,a=1,u=2,f=4,c=8,g={L:1,M:0,Q:3,H:2},l=0,h=1,s=2,v=3,d=4,w=5,p=6,y=7,B=(r=[[],[6,18],[6,22],[6,26],[6,30],[6,34],[6,22,38],[6,24,42],[6,26,46],[6,28,50],[6,30,54],[6,32,58],[6,34,62],[6,26,46,66],[6,26,48,70],[6,26,50,74],[6,30,54,78],[6,30,56,82],[6,30,58,86],[6,34,62,90],[6,28,50,72,94],[6,26,50,74,98],[6,30,54,78,102],[6,28,54,80,106],[6,32,58,84,110],[6,30,58,86,114],[6,34,62,90,118],[6,26,50,74,98,122],[6,30,54,78,102,126],[6,26,52,78,104,130],[6,30,56,82,108,134],[6,34,60,86,112,138],[6,30,58,86,114,142],[6,34,62,90,118,146],[6,30,54,78,102,126,150],[6,24,50,76,102,128,154],[6,28,54,80,106,132,158],[6,32,58,84,110,136,162],[6,26,54,82,110,138,166],[6,30,58,86,114,142,170]],e=1335,n=7973,i=function(t){for(var r=0;0!=t;)r+=1,t>>>=1;return r},(o={}).getBCHTypeInfo=function(t){for(var r=t<<10;i(r)-i(e)>=0;)r^=e<<i(r)-i(e);return 21522^(t<<10|r)},o.getBCHTypeNumber=function(t){for(var r=t<<12;i(r)-i(n)>=0;)r^=n<<i(r)-i(n);return t<<12|r},o.getPatternPosition=function(t){return r[t-1]},o.getMaskFunction=function(t){switch(t){case l:return function(t,r){return(t+r)%2==0};case h:return function(t,r){return t%2==0};case s:return function(t,r){return r%3==0};case v:return function(t,r){return(t+r)%3==0};case d:return function(t,r){return(Math.floor(t/2)+Math.floor(r/3))%2==0};case w:return function(t,r){return t*r%2+t*r%3==0};case p:return function(t,r){return(t*r%2+t*r%3)%2==0};case y:return function(t,r){return(t*r%3+(t+r)%2)%2==0};default:throw"bad maskPattern:"+t}},o.getErrorCorrectPolynomial=function(t){for(var r=k([1],0),e=0;e<t;e+=1)r=r.multiply(k([1,C.gexp(e)],0));return r},o.getLengthInBits=function(t,r){if(1<=r&&r<10)switch(t){case a:return 10;case u:return 9;case f:case c:return 8;default:throw"mode:"+t}else if(r<27)switch(t){case a:return 12;case u:return 11;case f:return 16;case c:return 10;default:throw"mode:"+t}else{if(!(r<41))throw"type:"+r;switch(t){case a:return 14;case u:return 13;case f:return 16;case c:return 12;default:throw"mode:"+t}}},o.getLostPoint=function(t){for(var r=t.getModuleCount(),e=0,n=0;n<r;n+=1)for(var o=0;o<r;o+=1){for(var i=0,a=t.isDark(n,o),u=-1;u<=1;u+=1)if(!(n+u<0||r<=n+u))for(var f=-1;f<=1;f+=1)o+f<0||r<=o+f||0==u&&0==f||a==t.isDark(n+u,o+f)&&(i+=1);i>5&&(e+=3+i-5)}for(n=0;n<r-1;n+=1)for(o=0;o<r-1;o+=1){var c=0;t.isDark(n,o)&&(c+=1),t.isDark(n+1,o)&&(c+=1),t.isDark(n,o+1)&&(c+=1),t.isDark(n+1,o+1)&&(c+=1),0!=c&&4!=c||(e+=3)}for(n=0;n<r;n+=1)for(o=0;o<r-6;o+=1)t.isDark(n,o)&&!t.isDark(n,o+1)&&t.isDark(n,o+2)&&t.isDark(n,o+3)&&t.isDark(n,o+4)&&!t.isDark(n,o+5)&&t.isDark(n,o+6)&&(e+=40);for(o=0;o<r;o+=1)for(n=0;n<r-6;n+=1)t.isDark(n,o)&&!t.isDark(n+1,o)&&t.isDark(n+2,o)&&t.isDark(n+3,o)&&t.isDark(n+4,o)&&!t.isDark(n+5,o)&&t.isDark(n+6,o)&&(e+=40);var g=0;for(o=0;o<r;o+=1)for(n=0;n<r;n+=1)t.isDark(n,o)&&(g+=1);return e+=Math.abs(100*g/r/r-50)/5*10},o),C=function(){for(var t=new Array(256),r=new Array(256),e=0;e<8;e+=1)t[e]=1<<e;for(e=8;e<256;e+=1)t[e]=t[e-4]^t[e-5]^t[e-6]^t[e-8];for(e=0;e<255;e+=1)r[t[e]]=e;var n={glog:function(t){if(t<1)throw"glog("+t+")";return r[t]},gexp:function(r){for(;r<0;)r+=255;for(;r>=256;)r-=255;return t[r]}};return n}();function k(t,r){if(void 0===t.length)throw t.length+"/"+r;var e=function(){for(var e=0;e<t.length&&0==t[e];)e+=1;for(var n=new Array(t.length-e+r),o=0;o<t.length-e;o+=1)n[o]=t[o+e];return n}(),n={getAt:function(t){return e[t]},getLength:function(){return e.length},multiply:function(t){for(var r=new Array(n.getLength()+t.getLength()-1),e=0;e<n.getLength();e+=1)for(var o=0;o<t.getLength();o+=1)r[e+o]^=C.gexp(C.glog(n.getAt(e))+C.glog(t.getAt(o)));return k(r,0)},mod:function(t){if(n.getLength()-t.getLength()<0)return n;for(var r=C.glog(n.getAt(0))-C.glog(t.getAt(0)),e=new Array(n.getLength()),o=0;o<n.getLength();o+=1)e[o]=n.getAt(o);for(o=0;o<t.getLength();o+=1)e[o]^=C.gexp(C.glog(t.getAt(o))+r);return k(e,0).mod(t)}};return n}var A=function(){var t=[[1,26,19],[1,26,16],[1,26,13],[1,26,9],[1,44,34],[1,44,28],[1,44,22],[1,44,16],[1,70,55],[1,70,44],[2,35,17],[2,35,13],[1,100,80],[2,50,32],[2,50,24],[4,25,9],[1,134,108],[2,67,43],[2,33,15,2,34,16],[2,33,11,2,34,12],[2,86,68],[4,43,27],[4,43,19],[4,43,15],[2,98,78],[4,49,31],[2,32,14,4,33,15],[4,39,13,1,40,14],[2,121,97],[2,60,38,2,61,39],[4,40,18,2,41,19],[4,40,14,2,41,15],[2,146,116],[3,58,36,2,59,37],[4,36,16,4,37,17],[4,36,12,4,37,13],[2,86,68,2,87,69],[4,69,43,1,70,44],[6,43,19,2,44,20],[6,43,15,2,44,16],[4,101,81],[1,80,50,4,81,51],[4,50,22,4,51,23],[3,36,12,8,37,13],[2,116,92,2,117,93],[6,58,36,2,59,37],[4,46,20,6,47,21],[7,42,14,4,43,15],[4,133,107],[8,59,37,1,60,38],[8,44,20,4,45,21],[12,33,11,4,34,12],[3,145,115,1,146,116],[4,64,40,5,65,41],[11,36,16,5,37,17],[11,36,12,5,37,13],[5,109,87,1,110,88],[5,65,41,5,66,42],[5,54,24,7,55,25],[11,36,12,7,37,13],[5,122,98,1,123,99],[7,73,45,3,74,46],[15,43,19,2,44,20],[3,45,15,13,46,16],[1,135,107,5,136,108],[10,74,46,1,75,47],[1,50,22,15,51,23],[2,42,14,17,43,15],[5,150,120,1,151,121],[9,69,43,4,70,44],[17,50,22,1,51,23],[2,42,14,19,43,15],[3,141,113,4,142,114],[3,70,44,11,71,45],[17,47,21,4,48,22],[9,39,13,16,40,14],[3,135,107,5,136,108],[3,67,41,13,68,42],[15,54,24,5,55,25],[15,43,15,10,44,16],[4,144,116,4,145,117],[17,68,42],[17,50,22,6,51,23],[19,46,16,6,47,17],[2,139,111,7,140,112],[17,74,46],[7,54,24,16,55,25],[34,37,13],[4,151,121,5,152,122],[4,75,47,14,76,48],[11,54,24,14,55,25],[16,45,15,14,46,16],[6,147,117,4,148,118],[6,73,45,14,74,46],[11,54,24,16,55,25],[30,46,16,2,47,17],[8,132,106,4,133,107],[8,75,47,13,76,48],[7,54,24,22,55,25],[22,45,15,13,46,16],[10,142,114,2,143,115],[19,74,46,4,75,47],[28,50,22,6,51,23],[33,46,16,4,47,17],[8,152,122,4,153,123],[22,73,45,3,74,46],[8,53,23,26,54,24],[12,45,15,28,46,16],[3,147,117,10,148,118],[3,73,45,23,74,46],[4,54,24,31,55,25],[11,45,15,31,46,16],[7,146,116,7,147,117],[21,73,45,7,74,46],[1,53,23,37,54,24],[19,45,15,26,46,16],[5,145,115,10,146,116],[19,75,47,10,76,48],[15,54,24,25,55,25],[23,45,15,25,46,16],[13,145,115,3,146,116],[2,74,46,29,75,47],[42,54,24,1,55,25],[23,45,15,28,46,16],[17,145,115],[10,74,46,23,75,47],[10,54,24,35,55,25],[19,45,15,35,46,16],[17,145,115,1,146,116],[14,74,46,21,75,47],[29,54,24,19,55,25],[11,45,15,46,46,16],[13,145,115,6,146,116],[14,74,46,23,75,47],[44,54,24,7,55,25],[59,46,16,1,47,17],[12,151,121,7,152,122],[12,75,47,26,76,48],[39,54,24,14,55,25],[22,45,15,41,46,16],[6,151,121,14,152,122],[6,75,47,34,76,48],[46,54,24,10,55,25],[2,45,15,64,46,16],[17,152,122,4,153,123],[29,74,46,14,75,47],[49,54,24,10,55,25],[24,45,15,46,46,16],[4,152,122,18,153,123],[13,74,46,32,75,47],[48,54,24,14,55,25],[42,45,15,32,46,16],[20,147,117,4,148,118],[40,75,47,7,76,48],[43,54,24,22,55,25],[10,45,15,67,46,16],[19,148,118,6,149,119],[18,75,47,31,76,48],[34,54,24,34,55,25],[20,45,15,61,46,16]],r=function(t,r){var e={};return e.totalCount=t,e.dataCount=r,e},e={};return e.getRSBlocks=function(e,n){var o=function(r,e){switch(e){case g.L:return t[4*(r-1)+0];case g.M:return t[4*(r-1)+1];case g.Q:return t[4*(r-1)+2];case g.H:return t[4*(r-1)+3];default:return}}(e,n);if(void 0===o)throw"bad rs block @ typeNumber:"+e+"/errorCorrectionLevel:"+n;for(var i=o.length/3,a=[],u=0;u<i;u+=1)for(var f=o[3*u+0],c=o[3*u+1],l=o[3*u+2],h=0;h<f;h+=1)a.push(r(c,l));return a},e}(),b=function(){var t=[],r=0,e={getBuffer:function(){return t},getAt:function(r){var e=Math.floor(r/8);return 1==(t[e]>>>7-r%8&1)},put:function(t,r){for(var n=0;n<r;n+=1)e.putBit(1==(t>>>r-n-1&1))},getLengthInBits:function(){return r},putBit:function(e){var n=Math.floor(r/8);t.length<=n&&t.push(0),e&&(t[n]|=128>>>r%8),r+=1}};return e},M=function(t){var r=a,e=t,n={getMode:function(){return r},getLength:function(t){return e.length},write:function(t){for(var r=e,n=0;n+2<r.length;)t.put(o(r.substring(n,n+3)),10),n+=3;n<r.length&&(r.length-n==1?t.put(o(r.substring(n,n+1)),4):r.length-n==2&&t.put(o(r.substring(n,n+2)),7))}},o=function(t){for(var r=0,e=0;e<t.length;e+=1)r=10*r+i(t.charAt(e));return r},i=function(t){if("0"<=t&&t<="9")return t.charCodeAt(0)-"0".charCodeAt(0);throw"illegal char :"+t};return n},x=function(t){var r=u,e=t,n={getMode:function(){return r},getLength:function(t){return e.length},write:function(t){for(var r=e,n=0;n+1<r.length;)t.put(45*o(r.charAt(n))+o(r.charAt(n+1)),11),n+=2;n<r.length&&t.put(o(r.charAt(n)),6)}},o=function(t){if("0"<=t&&t<="9")return t.charCodeAt(0)-"0".charCodeAt(0);if("A"<=t&&t<="Z")return t.charCodeAt(0)-"A".charCodeAt(0)+10;switch(t){case" ":return 36;case"$":return 37;case"%":return 38;case"*":return 39;case"+":return 40;case"-":return 41;case".":return 42;case"/":return 43;case":":return 44;default:throw"illegal char :"+t}};return n},m=function(r){var e=f,n=t.stringToBytes(r),o={getMode:function(){return e},getLength:function(t){return n.length},write:function(t){for(var r=0;r<n.length;r+=1)t.put(n[r],8)}};return o},L=function(r){var e=c,n=t.stringToBytesFuncs.SJIS;if(!n)throw"sjis not supported.";!function(){var t=n("友");if(2!=t.length||38726!=(t[0]<<8|t[1]))throw"sjis not supported."}();var o=n(r),i={getMode:function(){return e},getLength:function(t){return~~(o.length/2)},write:function(t){for(var r=o,e=0;e+1<r.length;){var n=(255&r[e])<<8|255&r[e+1];if(33088<=n&&n<=40956)n-=33088;else{if(!(57408<=n&&n<=60351))throw"illegal char at "+(e+1)+"/"+n;n-=49472}n=192*(n>>>8&255)+(255&n),t.put(n,13),e+=2}if(e<r.length)throw"illegal char at "+(e+1)}};return i},D=function(){var t=[],r={writeByte:function(r){t.push(255&r)},writeShort:function(t){r.writeByte(t),r.writeByte(t>>>8)},writeBytes:function(t,e,n){e=e||0,n=n||t.length;for(var o=0;o<n;o+=1)r.writeByte(t[o+e])},writeString:function(t){for(var e=0;e<t.length;e+=1)r.writeByte(t.charCodeAt(e))},toByteArray:function(){return t},toString:function(){var r="";r+="[";for(var e=0;e<t.length;e+=1)e>0&&(r+=","),r+=t[e];return r+="]"}};return r},S=function(t){var r=t,e=0,n=0,o=0,i={read:function(){for(;o<8;){if(e>=r.length){if(0==o)return-1;throw"unexpected end of file./"+o}var t=r.charAt(e);if(e+=1,"="==t)return o=0,-1;t.match(/^\s$/)||(n=n<<6|a(t.charCodeAt(0)),o+=6)}var i=n>>>o-8&255;return o-=8,i}},a=function(t){if(65<=t&&t<=90)return t-65;if(97<=t&&t<=122)return t-97+26;if(48<=t&&t<=57)return t-48+52;if(43==t)return 62;if(47==t)return 63;throw"c:"+t};return i},I=function(t,r,e){for(var n=function(t,r){var e=t,n=r,o=new Array(t*r),i={setPixel:function(t,r,n){o[r*e+t]=n},write:function(t){t.writeString("GIF87a"),t.writeShort(e),t.writeShort(n),t.writeByte(128),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(0),t.writeByte(255),t.writeByte(255),t.writeByte(255),t.writeString(","),t.writeShort(0),t.writeShort(0),t.writeShort(e),t.writeShort(n),t.writeByte(0);var r=a(2);t.writeByte(2);for(var o=0;r.length-o>255;)t.writeByte(255),t.writeBytes(r,o,255),o+=255;t.writeByte(r.length-o),t.writeBytes(r,o,r.length-o),t.writeByte(0),t.writeString(";")}},a=function(t){for(var r=1<<t,e=1+(1<<t),n=t+1,i=u(),a=0;a<r;a+=1)i.add(String.fromCharCode(a));i.add(String.fromCharCode(r)),i.add(String.fromCharCode(e));var f,c,g,l=D(),h=(f=l,c=0,g=0,{write:function(t,r){if(t>>>r!=0)throw"length over";for(;c+r>=8;)f.writeByte(255&(t<<c|g)),r-=8-c,t>>>=8-c,g=0,c=0;g|=t<<c,c+=r},flush:function(){c>0&&f.writeByte(g)}});h.write(r,n);var s=0,v=String.fromCharCode(o[s]);for(s+=1;s<o.length;){var d=String.fromCharCode(o[s]);s+=1,i.contains(v+d)?v+=d:(h.write(i.indexOf(v),n),i.size()<4095&&(i.size()==1<<n&&(n+=1),i.add(v+d)),v=d)}return h.write(i.indexOf(v),n),h.write(e,n),h.flush(),l.toByteArray()},u=function(){var t={},r=0,e={add:function(n){if(e.contains(n))throw"dup key:"+n;t[n]=r,r+=1},size:function(){return r},indexOf:function(r){return t[r]},contains:function(r){return void 0!==t[r]}};return e};return i}(t,r),o=0;o<r;o+=1)for(var i=0;i<t;i+=1)n.setPixel(i,o,e(i,o));var a=D();n.write(a);for(var u=function(){var t=0,r=0,e=0,n="",o={},i=function(t){n+=String.fromCharCode(a(63&t))},a=function(t){if(t<0);else{if(t<26)return 65+t;if(t<52)return t-26+97;if(t<62)return t-52+48;if(62==t)return 43;if(63==t)return 47}throw"n:"+t};return o.writeByte=function(n){for(t=t<<8|255&n,r+=8,e+=1;r>=6;)i(t>>>r-6),r-=6},o.flush=function(){if(r>0&&(i(t<<6-r),t=0,r=0),e%3!=0)for(var o=3-e%3,a=0;a<o;a+=1)n+="="},o.toString=function(){return n},o}(),f=a.toByteArray(),c=0;c<f.length;c+=1)u.writeByte(f[c]);return u.flush(),"data:image/gif;base64,"+u};return t}();qrcode.stringToBytesFuncs["UTF-8"]=function(t){return function(t){for(var r=[],e=0;e<t.length;e++){var n=t.charCodeAt(e);n<128?r.push(n):n<2048?r.push(192|n>>6,128|63&n):n<55296||n>=57344?r.push(224|n>>12,128|n>>6&63,128|63&n):(e++,n=65536+((1023&n)<<10|1023&t.charCodeAt(e)),r.push(240|n>>18,128|n>>12&63,128|n>>6&63,128|63&n))}return r}(t)},function(t){"function"==typeof define&&define.amd?define([],t):"object"==typeof exports&&(module.exports=t())}(function(){return qrcode});
</script>
<script>
function h(tag, attrs, kids) {
  var e = document.createElement(tag), k;
  if (attrs) for (k in attrs) {
    var v = attrs[k];
    if (v == null || v === false) continue;
    if (k === 'class') e.className = v;
    else if (k === 'text') e.textContent = v;
    else if (k === 'value') e.value = v;
    else if (k.slice(0, 2) === 'on') e.addEventListener(k.slice(2), v);
    else e.setAttribute(k, v === true ? '' : v);
  }
  (kids || []).forEach(function (c) {
    if (c == null || c === false) return;
    e.appendChild(typeof c === 'string' ? document.createTextNode(c) : c);
  });
  return e;
}
var IC = {
  back: 'M15 5l-7 7 7 7', plus: 'M12 5v14M5 12h14', minus: 'M5 12h14', x: 'M6 6l12 12M18 6L6 18', check: 'M5 12l5 5 9-10',
  bag: 'M6 8h12l-1 12H7L6 8z|M9 8V6a3 3 0 0 1 6 0v2', up: 'M12 16V5M7 10l5-5 5 5M5 19h14',
  cam: 'M4 8h3l2-3h6l2 3h3v11H4z|M12 16.5a3.5 3.5 0 1 0 0-7 3.5 3.5 0 0 0 0 7z', pic: 'M4 5h16v14H4z|M4 16l5-5 4 4 3-3 4 4'
};
function ic(name, size) {
  var ns = 'http://www.w3.org/2000/svg', s = document.createElementNS(ns, 'svg');
  var a = { width: size || 22, height: size || 22, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', 'stroke-width': 2, 'stroke-linecap': 'round', 'stroke-linejoin': 'round', 'aria-hidden': 'true' };
  for (var k in a) s.setAttribute(k, a[k]);
  IC[name].split('|').forEach(function (d) { var p = document.createElementNS(ns, 'path'); p.setAttribute('d', d); s.appendChild(p); });
  return s;
}
var toastTimer = null;
function toast(msg) {
  var t = document.getElementById('toast');
  t.textContent = msg; t.style.display = 'block';
  clearTimeout(toastTimer);
  toastTimer = setTimeout(function () { t.style.display = 'none'; }, 5000);
}
function busy(on) { document.getElementById('busy').style.display = on ? 'flex' : 'none'; }
function call(fn, args, ok, opts) {
  opts = opts || {};
  if (!opts.quiet) busy(true);
  var r = google.script.run
    .withSuccessHandler(function (x) { if (!opts.quiet) busy(false); ok(x); })
    .withFailureHandler(function (e) {
      if (!opts.quiet) busy(false);
      var m = String((e && e.message) || '').replace(/^Error:\s*/, '');
      if (opts.fail && opts.fail(m)) return;
      if (!opts.quiet) toast(m && m.length < 220 ? m : 'Something went wrong. Please try again.');
    });
  r[fn].apply(r, args);
}
function rs(n) { return 'Rs.' + n; }
function photo(id, alt, cls, w) {
  if (!id) return h('div', { class: 'ph ' + cls, text: 'No photo yet' });
  var im = h('img', { class: cls, alt: alt, loading: 'lazy', src: 'https://lh3.googleusercontent.com/d/' + id + '=w' + (w || 600) });
  im.addEventListener('error', function () {
    if (im.getAttribute('data-f')) return;
    im.setAttribute('data-f', '1');
    im.src = 'https://drive.google.com/thumbnail?id=' + id + '&sz=w' + (w || 600);
  });
  return im;
}
/** Shrinks a picked photo on the phone. square=true crops the middle to a square. */
function shrink(file, maxSide, square, quality, cb) {
  var fr = new FileReader();
  fr.onerror = function () { toast('That picture could not be read.'); };
  fr.onload = function () {
    var im = new Image();
    im.onerror = function () { toast('That picture could not be read. Please try a JPG or PNG.'); };
    im.onload = function () {
      var sw = im.naturalWidth, sh = im.naturalHeight, sx = 0, sy = 0;
      if (square) { var m = Math.min(sw, sh); sx = (sw - m) / 2; sy = (sh - m) / 2; sw = sh = m; }
      var k = Math.min(1, maxSide / Math.max(sw, sh));
      var c = document.createElement('canvas');
      c.width = Math.round(sw * k); c.height = Math.round(sh * k);
      var x = c.getContext('2d');
      x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height);
      x.drawImage(im, sx, sy, sw, sh, 0, 0, c.width, c.height);
      cb(c.toDataURL('image/jpeg', quality));
    };
    im.src = fr.result;
  };
  fr.readAsDataURL(file);
}
function qrCanvas(text, size) {
  var c = h('canvas', { width: size, height: size, role: 'img', 'aria-label': 'UPI QR code' });
  c.style.cssText = 'width:' + size / 2 + 'px;height:' + size / 2 + 'px;border-radius:8px;background:#fff';
  try {
    var q = qrcode(0, 'M'); q.addData(text); q.make();
    var n = q.getModuleCount(), cell = Math.floor(size / (n + 6)), off = Math.floor((size - n * cell) / 2), x = c.getContext('2d');
    x.fillStyle = '#fff'; x.fillRect(0, 0, size, size); x.fillStyle = '#000';
    for (var r = 0; r < n; r++) for (var k = 0; k < n; k++) if (q.isDark(r, k)) x.fillRect(off + k * cell, off + r * cell, cell, cell);
    return c;
  } catch (e) { return null; }
}
function stepper(value, min, max, onChange, what) {
  var num = h('span', { 'aria-live': 'polite', text: String(value) });
  function set(v) { v = Math.max(min, Math.min(max, v)); if (v === value) return; value = v; num.textContent = String(v); onChange(v); }
  return h('div', { class: 'stepper', role: 'group', 'aria-label': what }, [
    h('button', { type: 'button', 'aria-label': 'One less', onclick: function () { set(value - 1); } }, [ic('minus', 16)]),
    num,
    h('button', { type: 'button', 'aria-label': 'One more', onclick: function () { set(value + 1); } }, [ic('plus', 16)])
  ]);
}
function field(label, input) { return h('label', { class: 'field' }, [label, input]); }
function keep(key, val) {
  try {
    if (val === undefined) return window.localStorage.getItem(key);
    if (val === null) window.localStorage.removeItem(key); else window.localStorage.setItem(key, val);
  } catch (e) { }
  return null;
}
function sheet(kids, onClose) {
  var s = h('div', { class: 'sheet', role: 'dialog', 'aria-modal': 'true' }, [h('div', null, kids)]);
  s.addEventListener('click', function (e) { if (e.target === s) close(); });
  function close() { if (s.parentNode) s.parentNode.removeChild(s); if (onClose) onClose(); }
  document.body.appendChild(s);
  return close;
}
function topbar(title, onBack, right) {
  return h('div', { class: 'row', style: 'padding:12px 12px 4px;gap:4px' }, [
    onBack ? h('button', { type: 'button', class: 'iconbtn', 'aria-label': 'Back', onclick: onBack }, [ic('back', 24)]) : h('div', { style: 'width:8px' }),
    h('h1', { class: 'grow', text: title }), right || null
  ]);
}

var BOOT = "__BOOT__";
var app = document.getElementById('app');
var session = keep('ll_admin') || '';
var D = null, tab = BOOT.tab || 'orders', focusOrder = BOOT.order || '', stats = null, stockTimers = {};

function signOut() { session = ''; keep('ll_admin', null); D = null; render(); }
function api(fn, args, ok, quiet) {
  call(fn, [session].concat(args || []), ok, { quiet: quiet, fail: function (m) { if (m.indexOf('SIGNIN') >= 0) { signOut(); toast('Please enter your PIN again.'); return true; } return false; } });
}
function act(fn, args, done) { api(fn, args, function (d) { D = d; render(); if (done) done(); }); }
function render() {
  app.textContent = '';
  var v = !D ? viewLogin() : (tab === 'products' ? viewProducts() : tab === 'reviews' ? viewReviews() : tab === 'sales' ? viewSales() : viewOrders());
  v.forEach(function (n) { if (n) app.appendChild(n); });
  if (D) app.appendChild(nav());
  if (D && focusOrder && tab === 'orders') { var el = document.getElementById('o-' + focusOrder); focusOrder = ''; if (el) el.scrollIntoView(); }
}
function upiLink(amount, note) {
  return 'upi://pay?pa=' + encodeURIComponent(D.store.upiId) + '&pn=' + encodeURIComponent(D.store.upiName) + '&am=' + Number(amount).toFixed(2) + '&cu=INR&tn=' + encodeURIComponent(note);
}

/* ------------------------------------------------------------------ login */
function viewLogin() {
  var pin = h('input', { type: 'password', inputmode: 'numeric', autocomplete: 'off', maxlength: 8, style: 'text-align:center;font-size:22px;letter-spacing:6px' });
  function go() {
    call('adminLogin', [pin.value], function (r) {
      if (!r.ok) { toast(r.error); pin.value = ''; return; }
      session = r.session; keep('ll_admin', session); D = r.data; render();
    });
  }
  pin.addEventListener('keydown', function (e) { if (e.key === 'Enter') go(); });
  if (session) setTimeout(function () { api('adminRefresh', [], function (d) { D = d; render(); }); }, 0);
  return [h('div', { class: 'stack pad', style: 'padding-top:72px' }, [
    h('div', { class: 'small muted', text: 'Store admin' }), h('h1', { style: 'font-size:26px;line-height:32px', text: BOOT.store }),
    field('Enter your PIN', pin), h('button', { type: 'button', class: 'btn', text: 'Open', onclick: go })
  ])];
}
function nav() {
  var todo = D.orders.filter(function (o) { return o.status === 'payment_check' || o.refundDue || (o.status === 'confirmed' && o.mode === 'ship'); }).length;
  var waiting = (D.reviews || []).filter(function (r) { return r.status === 'waiting'; }).length;
  function b(id, label) { return h('button', { type: 'button', class: tab === id ? 'on' : '', 'aria-current': tab === id ? 'page' : null, text: label, onclick: function () { tab = id; render(); window.scrollTo(0, 0); if (id !== 'sales') api('adminRefresh', [], function (d) { D = d; render(); }, true); } }); }
  return h('nav', { class: 'nav', 'aria-label': 'Admin sections' }, [b('orders', 'Orders' + (todo ? ' · ' + todo : '')), b('products', 'Products'), b('reviews', 'Reviews' + (waiting ? ' · ' + waiting : '')), b('sales', 'Insights')]);
}
function header(title, right) {
  return h('div', { class: 'row between', style: 'padding:20px 20px 8px' }, [
    h('div', null, [h('div', { class: 'small muted', text: D.store.name + ' admin' }), h('h1', { style: 'font-size:24px;line-height:30px', text: title })]), right
  ]);
}

/* ----------------------------------------------------------------- orders */
function contact(o) {
  if (!o.phone && !o.email) return null;
  var digits = o.phone.replace(/\D/g, ''); if (digits.length === 10) digits = '91' + digits;
  function a(text, href, blank) { return h('a', { class: 'btn sec sm', href: href, target: blank ? '_blank' : null, rel: blank ? 'noopener' : null, text: text }); }
  return h('div', { class: 'row' }, [o.phone ? a('Call', 'tel:' + o.phone) : null, o.phone ? a('WhatsApp', 'https://wa.me/' + digits, true) : null, o.email ? a('Email', 'mailto:' + o.email) : null]);
}
function head(o, tag) {
  return h('div', { class: 'row between', style: 'align-items:flex-start' }, [
    h('div', { class: 'grow stack', style: 'gap:3px' }, [
      h('div', { style: 'font-weight:600', text: o.id + (o.name ? ' · ' + o.name : '') }),
      h('div', { class: 'small muted', text: o.itemsText }),
      h('div', { class: 'small muted', text: (o.source === 'inperson' ? 'In-person sale' : o.mode === 'ship' ? 'Deliver' : 'Collect in person') + ' · ' + o.created })
    ]),
    h('div', { style: 'text-align:right;flex-shrink:0' }, [h('div', { class: 'price', style: 'font-size:18px', text: rs(o.total) }), tag || null])
  ]);
}
function proofBits(o) {
  var out = [];
  if (o.proofRef) out.push(h('div', { class: 'row between', style: 'padding:8px 8px 8px 12px;border-radius:12px;background:#F3F6F4' }, [
    h('div', null, [h('div', { class: 'tiny muted', text: 'UPI reference number' }), h('div', { class: 'price', style: 'font-size:19px;letter-spacing:1px', text: o.proofRef })]),
    h('button', { type: 'button', class: 'btn sec sm fit', text: 'Copy', onclick: function () { copy(o.proofRef); } })]));
  if (o.hasShot) out.push(h('button', { type: 'button', class: 'btn sec sm', text: 'View payment screenshot', onclick: function () {
    api('adminProof', [o.id], function (src) { var close = sheet([h('h2', { text: 'Payment screenshot · ' + o.id }), h('img', { src: src, alt: 'Payment screenshot', style: 'width:100%;border-radius:12px' }), h('button', { type: 'button', class: 'btn sec', text: 'Close', onclick: function () { close(); } })]); });
  } }));
  else if (o.shotDeleted) out.push(h('div', { class: 'tiny muted', text: 'Screenshot removed after the keeping period.' }));
  if (o.flags) out.push(h('div', { class: 'note', text: o.flags }));
  return out;
}
function copy(text) {
  try { navigator.clipboard.writeText(text).then(function () { toast('Copied.'); }, function () { toast(text); }); } catch (e) { toast(text); }
}
function card(o, hot, kids) { return h('div', { class: 'card stack ocard' + (hot ? ' hot' : ''), id: 'o-' + o.id }, kids); }
function two(a, b) { return h('div', { class: 'grid2', style: 'gap:10px' }, [a, b]); }
function button(text, cls, fn) { return h('button', { type: 'button', class: 'btn sm ' + (cls || ''), text: text, onclick: fn }); }
function cancelButton(o) { return button('Cancel order', 'warn', function () { cancelSheet(o); }); }

function viewOrders() {
  var by = { check: [], refund: [], pack: [], collect: [], waiting: [], sent: [], done: [] };
  D.orders.forEach(function (o) {
    if (o.refundDue) by.refund.push(o);
    else if (o.status === 'payment_check') by.check.push(o);
    else if (o.status === 'confirmed' && o.mode === 'ship') by.pack.push(o);
    else if (o.status === 'cash_pickup' || o.status === 'confirmed') by.collect.push(o);
    else if (o.status === 'awaiting_payment') by.waiting.push(o);
    else if (o.status === 'shipped') by.sent.push(o);
    else by.done.push(o);
  });
  var out = [header('Orders', h('button', { type: 'button', class: 'btn sm fit', onclick: saleSheet }, [ic('plus', 18), 'In-person sale']))];
  if (D.store.missing.length) out.push(h('div', { class: 'note', style: 'margin:4px 20px 0', text: 'Finish setting up: add your ' + D.store.missing.join(', ') + ' in the Sheet\'s Settings tab.' }));
  function section(title, list, amber, fn) {
    if (!list.length) return;
    out.push(h('h2', { class: 'sec-h' + (amber ? ' amber' : ''), text: title + ' · ' + list.length }));
    list.forEach(function (o) { out.push(fn(o)); });
  }
  section('Payment to check', by.check, true, function (o) {
    return card(o, true, [head(o), h('div', { class: 'small muted', text: 'Says paid ' + o.proofAt + (o.mode === 'ship' ? ' · ' + o.address : '') })].concat(proofBits(o)).concat([
      two(button('Payment received', '', function () { act('adminPaymentDecision', [o.id, true], function () { toast('Order confirmed. The customer has been emailed.'); }); }),
        button('Not received', 'sec', function () {
          var close = sheet([h('h2', { text: 'No payment for ' + o.id + '?' }), h('p', { class: 'small muted', text: 'The order is cancelled, the items go back on sale, and the customer is emailed to reply with proof if they did pay.' }),
            button('Yes, payment not received', 'warn', function () { close(); act('adminPaymentDecision', [o.id, false]); }), button('Go back', 'sec', function () { close(); })]);
        })),
      h('div', { class: 'tiny muted', text: 'Find ' + rs(o.total) + ' in your UPI app first. A screenshot can be faked.' }), contact(o)]));
  });
  section('Refunds to make', by.refund, true, function (o) {
    return card(o, true, [head(o, h('span', { class: 'tag red', text: 'Refund due' })), h('div', { class: 'small muted', text: o.cancelReason || 'Cancelled after payment' })].concat(proofBits(o)).concat([
      button('I have refunded ' + rs(o.total), '', function () { act('adminRefunded', [o.id]); }), contact(o)]));
  });
  section('To pack and send', by.pack, false, function (o) {
    return card(o, false, [head(o, h('span', { class: 'tag', text: 'Paid' })), h('div', { class: 'small', text: o.address }),
      two(button('Mark as sent', '', function () { shipSheet(o); }), button('Copy address', 'sec', function () { copy(o.name + ', ' + o.phone + ', ' + o.address); })), contact(o), cancelButton(o)]);
  });
  section('Waiting to be collected', by.collect, false, function (o) {
    var kids = [head(o, o.paid ? h('span', { class: 'tag', text: 'Paid' }) : h('span', { class: 'tag amber', text: 'Cash due' }))];
    if (!o.paid) kids.push(h('div', { class: 'small muted', text: 'Kept aside until ' + o.hold }));
    kids.push(o.paid ? button('Collected', '', function () { act('adminCollected', [o.id, '']); })
      : two(button('Collected · cash', '', function () { act('adminCollected', [o.id, 'cash']); }), button('Collected · UPI', 'sec', function () {
        var qr = qrCanvas(upiLink(o.total, 'Order ' + o.id), 400);
        var close = sheet([h('h2', { text: 'Collect ' + rs(o.total) + ' by UPI' }), qr ? h('div', { style: 'display:flex;justify-content:center' }, [qr]) : null, h('div', { class: 'small', style: 'text-align:center', text: 'UPI ID: ' + D.store.upiId }),
          button('Payment received · mark collected', '', function () { close(); act('adminCollected', [o.id, 'upi']); }), button('Close', 'sec', function () { close(); })]);
      })));
    kids.push(contact(o)); kids.push(cancelButton(o));
    return card(o, false, kids);
  });
  section('On the way', by.sent, false, function (o) {
    var det = [o.shipProvider, o.driverName, o.driverPhone].filter(function (x) { return x; }).join(' · ');
    return card(o, false, [head(o, h('span', { class: 'tag', text: 'Sent' })), h('div', { class: 'small muted', text: 'Sent ' + o.shippedAt + (det ? ' · ' + det : '') }),
      o.shipTracking ? h('a', { class: 'small', href: o.shipTracking, target: '_blank', rel: 'noopener', text: 'Open tracking link' }) : null,
      two(button('Mark delivered', '', function () { act('adminDelivered', [o.id]); }), button('Edit delivery details', 'sec', function () { shipSheet(o); })), contact(o)]);
  });
  section('Waiting for the customer to pay', by.waiting, false, function (o) {
    return card(o, false, [head(o), h('div', { class: 'small muted', text: 'Items held until ' + o.hold + '. Released by itself if not paid.' }), contact(o), cancelButton(o)]);
  });
  section('Recent', by.done, false, function (o) {
    return card(o, false, [head(o, h('span', { class: 'tag' + (o.status === 'cancelled' ? ' red' : ''), text: o.status === 'cancelled' ? 'Cancelled' : 'Done' })),
      o.status === 'cancelled' && o.cancelReason ? h('div', { class: 'small muted', text: o.cancelReason }) : null]);
  });
  if (out.length === 1 + (D.store.missing.length ? 1 : 0)) out.push(h('p', { class: 'pad muted', style: 'padding-top:16px', text: 'No orders yet. Share your shop link to get started.' }));
  out.push(h('div', { class: 'stack pad', style: 'padding-top:16px;padding-bottom:20px' }, [
    button('Refresh', 'sec', function () { act('adminRefresh', []); }),
    D.store.shopUrl ? button('Copy shop link', 'sec', function () { copy(D.store.shopUrl); }) : null,
    D.store.emailsLeft != null ? h('div', { class: 'tiny muted', style: 'text-align:center', text: 'Emails left today: ' + D.store.emailsLeft }) : null
  ]));
  return out;
}
function shipSheet(o) {
  var edit = o.status === 'shipped';
  function inp(val, attrs) { attrs.value = val; return h('input', attrs); }
  var provider = inp(o.shipProvider, { type: 'text', maxlength: 60, placeholder: 'Courier or delivery app' });
  var link = inp(o.shipTracking, { type: 'url', inputmode: 'url', maxlength: 400, placeholder: 'Paste the link from the delivery app' });
  var dname = inp(o.driverName, { type: 'text', maxlength: 60 });
  var dphone = inp(o.driverPhone, { type: 'tel', maxlength: 16 });
  var note = inp(o.shipNote, { type: 'text', maxlength: 200, placeholder: 'e.g. Arriving between 4 and 6 pm' });
  var tell = h('input', { type: 'checkbox', id: 'tell', style: 'width:24px;height:24px;accent-color:#1E4A40' });
  var close = sheet([
    h('h2', { text: (edit ? 'Delivery details · ' : 'Send order ') + o.id }),
    h('p', { class: 'small muted', text: 'All of these are optional. Whatever you fill in is emailed to the customer.' }),
    field('Delivery service', provider), field('Tracking link', link), field('Driver name', dname), field('Driver phone', dphone), field('Note for the customer', note),
    edit ? h('div', { class: 'row' }, [tell, h('label', { for: 'tell', class: 'small', text: 'Email the customer about this change' })]) : null,
    button(edit ? 'Save details' : 'Mark as sent and email the customer', '', function () {
      var d = { provider: provider.value, tracking: link.value.trim(), driverName: dname.value, driverPhone: dphone.value.trim(), note: note.value };
      if (d.tracking && !/^https?:\/\//i.test(d.tracking)) return toast('The tracking link should start with https://');
      close(); act('adminShip', [o.id, d, edit ? tell.checked : true], function () { toast(edit ? 'Details saved.' : 'Marked as sent.'); });
    }),
    button('Go back', 'sec', function () { close(); })
  ]);
}
function cancelSheet(o) {
  var why = h('select', null, ['Out of stock', 'Customer asked to cancel', 'Could not deliver', 'Not collected', 'Other'].map(function (x) { return h('option', { value: x, text: x }); }));
  var more = h('input', { type: 'text', maxlength: 120 });
  var close = sheet([
    h('h2', { text: 'Cancel order ' + o.id + '?' }),
    h('p', { class: 'small muted', text: 'The items go back on sale and the customer is emailed.' + (o.paid ? ' This order is paid, so it will move to Refunds to make.' : '') }),
    field('Reason', why), field('Anything to add (optional)', more),
    button('Yes, cancel the order', 'warn', function () { close(); act('adminCancel', [o.id, why.value + (more.value.trim() ? ': ' + more.value.trim() : '')]); }),
    button('Keep the order', 'sec', function () { close(); })
  ]);
}
function saleSheet() {
  var qty = {}, via = 'cash';
  var totalEl = h('div', { class: 'price', style: 'font-size:22px' }), qrBox = h('div');
  var name = h('input', { type: 'text', maxlength: 80 }), phone = h('input', { type: 'tel', maxlength: 16 });
  var list = D.products.filter(function (p) { return p.stock > 0; });
  function total() { var t = 0; list.forEach(function (p) { t += (qty[p.id] || 0) * p.price; }); return t; }
  function draw() {
    var t = total();
    totalEl.textContent = 'Total ' + rs(t);
    qrBox.textContent = '';
    if (via === 'upi' && t > 0 && D.store.upiId) { var qr = qrCanvas(upiLink(t, D.store.name), 400); if (qr) qrBox.appendChild(h('div', { style: 'display:flex;justify-content:center' }, [qr])); }
  }
  var cash = h('button', { type: 'button', class: 'opt on', onclick: function () { via = 'cash'; cash.className = 'opt on'; upi.className = 'opt'; draw(); } }, [h('b', { text: 'Cash' })]);
  var upi = h('button', { type: 'button', class: 'opt', onclick: function () { via = 'upi'; upi.className = 'opt on'; cash.className = 'opt'; draw(); } }, [h('b', { text: 'UPI' }), h('span', { text: 'Shows a QR code' })]);
  var close = sheet([
    h('h2', { text: 'In-person sale' }),
    list.length ? h('div', { class: 'stack', style: 'gap:8px' }, list.map(function (p) {
      return h('div', { class: 'card row', style: 'padding:10px' }, [photo(p.photos[0], p.name, 'thumb', 200),
        h('div', { class: 'grow' }, [h('div', { style: 'font-weight:500;font-size:14px', text: p.name }), h('div', { class: 'small muted', text: rs(p.price) + ' · ' + p.stock + ' in stock' })]),
        stepper(0, 0, p.stock, function (v) { qty[p.id] = v; draw(); }, 'Quantity of ' + p.name)]);
    })) : h('p', { class: 'muted', text: 'Nothing is in stock.' }),
    totalEl, two(cash, upi), qrBox,
    field('Customer name (optional)', name), field('Customer phone (optional)', phone),
    button('Record sale', '', function () {
      var items = list.filter(function (p) { return qty[p.id] > 0; }).map(function (p) { return { id: p.id, qty: qty[p.id] }; });
      if (!items.length) return toast('Choose at least one item.');
      close(); act('adminInPersonSale', [{ items: items, via: via, name: name.value, phone: phone.value }], function () { toast('Sale recorded. Stock updated.'); });
    }),
    button('Go back', 'sec', function () { close(); })
  ]);
  draw();
}

/* --------------------------------------------------------------- products */
function viewProducts() {
  var out = [header('Products', h('button', { type: 'button', class: 'btn sm fit', onclick: function () { productSheet(null); } }, [ic('plus', 18), 'Add product'])),
    h('p', { class: 'small muted pad', style: 'padding-bottom:12px', text: 'Tap + or − to change how many you have. Stock goes down by itself when someone orders.' })];
  if (!D.products.length) out.push(h('p', { class: 'pad muted', text: 'No products yet. Tap Add product to add your first one.' }));
  out.push(h('div', { class: 'stack pad', style: 'gap:10px;padding-bottom:20px' }, D.products.map(function (p) {
    var tag = h('span');
    function setTag() {
      var t = !p.active ? ['Hidden', 'red'] : p.stock < 1 ? ['Sold out', 'amber'] : p.stock <= D.store.lowStockAt ? ['Low stock', 'amber'] : ['On sale', ''];
      tag.className = 'tag ' + t[1]; tag.textContent = t[0];
    }
    setTag();
    return h('div', { class: 'card row', style: 'padding:12px' }, [
      photo(p.photos[0], p.name, 'thumb', 200),
      h('div', { class: 'grow stack', style: 'gap:4px' }, [
        h('div', { style: 'font-weight:600;font-size:14px;line-height:19px', text: p.name }),
        h('div', { class: 'row wrap', style: 'gap:4px 8px' }, [h('span', { text: rs(p.price) }), tag]),
        h('button', { type: 'button', style: 'border:none;background:none;padding:0;color:#1E4A40;font:600 13px Inter,sans-serif;text-align:left;min-height:28px;cursor:pointer', text: 'Edit', onclick: function () { productSheet(p); } })
      ]),
      stepper(p.stock, 0, 9999, function (v) {
        p.stock = v; setTag();
        clearTimeout(stockTimers[p.id]);
        stockTimers[p.id] = setTimeout(function () { api('adminSetStock', [p.id, v], function (d) { D = d; }, true); }, 900);
      }, 'Stock of ' + p.name)
    ]);
  })));
  return out;
}
function productSheet(p) {
  var photos = (p ? p.photos : []).map(function (id) { return { id: id }; });
  var cat = p ? p.category : (D.store.categories[0] || ''), stock = p ? p.stock : 1;
  var name = h('input', { type: 'text', maxlength: 80, value: p ? p.name : '' });
  var price = h('input', { type: 'text', inputmode: 'numeric', maxlength: 7, value: p ? String(p.price) : '' });
  var desc = h('textarea', { rows: 4, maxlength: 600, placeholder: 'What it is made of, the scent, the size, how long it burns' });
  desc.value = p ? p.description : '';
  var live = h('input', { type: 'checkbox', id: 'live', style: 'width:24px;height:24px;accent-color:#1E4A40' });
  live.checked = p ? p.active : true;
  var grid = h('div', { class: 'grid3' }), cats = h('div', { class: 'row wrap', style: 'gap:8px' });
  function picker(label, icon, capture) {
    var input = h('input', { type: 'file', accept: 'image/*', capture: capture ? 'environment' : null, style: 'position:absolute;opacity:0;width:1px;height:1px', onchange: function (e) {
      var f = e.target.files && e.target.files[0];
      if (f) shrink(f, 900, true, 0.8, function (d) { photos.push({ data: d }); drawPhotos(); });
    } });
    return h('label', { style: 'position:relative;aspect-ratio:1/1;border-radius:12px;border:2px dashed #B9C6C1;background:#fff;font-size:13px;font-weight:600;cursor:pointer;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:8px;text-align:center' }, [ic(icon, 24), label, input]);
  }
  function drawPhotos() {
    grid.textContent = '';
    photos.forEach(function (ph, i) {
      var im = ph.data ? h('img', { class: 'sq', src: ph.data, alt: 'New photo ' + (i + 1) }) : photo(ph.id, 'Photo ' + (i + 1), 'sq', 300);
      grid.appendChild(h('div', { style: 'position:relative;border-radius:12px;overflow:hidden' }, [im,
        h('button', { type: 'button', 'aria-label': 'Remove photo ' + (i + 1), style: 'position:absolute;top:0;right:0;width:44px;height:44px;border:none;background:none;cursor:pointer;display:flex;align-items:center;justify-content:center',
          onclick: function () { photos.splice(i, 1); drawPhotos(); } }, [h('span', { style: 'width:26px;height:26px;border-radius:13px;background:#16211E;color:#fff;display:flex;align-items:center;justify-content:center' }, [ic('x', 14)])])]));
    });
    if (photos.length < 4) { grid.appendChild(picker('Take photo', 'cam', true)); grid.appendChild(picker('From gallery', 'pic', false)); }
  }
  function drawCats() {
    cats.textContent = '';
    D.store.categories.forEach(function (c) { cats.appendChild(h('button', { type: 'button', class: 'chip' + (c === cat ? ' on' : ''), 'aria-pressed': String(c === cat), text: c, onclick: function () { cat = c; drawCats(); } })); });
  }
  drawPhotos(); drawCats();
  var close = sheet([
    h('h2', { text: p ? 'Edit product' : 'Add a product' }),
    h('div', { class: 'small', style: 'font-weight:600', text: 'Photos (up to 4, the first is the main photo)' }), grid,
    h('div', { class: 'tiny muted', text: 'Photos are cropped square and made smaller on your phone, so they upload quickly.' }),
    field('Name', name),
    h('div', { class: 'stack', style: 'gap:8px' }, [h('div', { class: 'small', style: 'font-weight:500', text: 'Category' }), cats]),
    h('div', { class: 'grid2', style: 'align-items:end' }, [field('Price (Rs.)', price),
      h('div', { class: 'stack', style: 'gap:6px' }, [h('div', { class: 'small', style: 'font-weight:500', text: 'How many do you have?' }), stepper(stock, 0, 9999, function (v) { stock = v; }, 'Stock')])]),
    field('Description (optional)', desc),
    h('div', { class: 'card row' }, [live, h('label', { for: 'live', class: 'grow' }, [h('div', { style: 'font-weight:600', text: 'Show in the shop' }), h('div', { class: 'small muted', text: 'Untick to hide it from customers' })])]),
    button('Save product', '', function () {
      var d = { id: p ? p.id : '', name: name.value, category: cat, price: Number(String(price.value).replace(/[^\d.]/g, '')), stock: stock, description: desc.value, active: live.checked,
        keep: photos.filter(function (x) { return x.id; }).map(function (x) { return x.id; }), add: photos.filter(function (x) { return x.data; }).map(function (x) { return x.data; }) };
      if (String(d.name).trim().length < 2) return toast('Please give the product a name.');
      if (!(d.price >= 1)) return toast('Please enter a price in rupees.');
      api('adminSaveProduct', [d], function (x) { close(); D = x; render(); toast('Product saved.'); });
    }),
    button('Go back', 'sec', function () { close(); })
  ]);
}

/* ---------------------------------------------------------------- reviews */
function stars(n) { var t = ''; for (var i = 1; i <= 5; i++) t += i <= n ? '★' : '☆'; return h('span', { role: 'img', 'aria-label': n + ' out of 5 stars', style: 'color:#B8730A;font-size:17px;letter-spacing:1px', text: t }); }
function viewReviews() {
  var out = [header('Reviews', null)];
  if (!D.store.reviewsOn) out.push(h('div', { class: 'note', style: 'margin:4px 20px 0', text: 'Reviews are switched off (Settings ▸ reviewsOn). Buyers are not invited and nothing is shown in the shop.' }));
  out.push(h('div', { class: 'card small', style: 'margin:4px 20px 0' }, [h('div', { style: 'font-weight:600;margin-bottom:4px', text: 'How to check a review' }),
    h('div', { text: 'Publish honest reviews, good or bad. You cannot change what a buyer wrote; reply instead. Don\'t publish abuse, personal details, spam, or claims that a product treats or cures a health condition.' })]));
  var waiting = D.reviews.filter(function (r) { return r.status === 'waiting'; }), rest = D.reviews.filter(function (r) { return r.status !== 'waiting'; });
  function bits(r) {
    return [
      h('div', { class: 'row between', style: 'align-items:flex-start' }, [h('div', { class: 'grow' }, [h('div', { style: 'font-weight:600', text: r.productName }), h('div', { class: 'small muted', text: r.customerName + ' · order ' + r.orderId + ' · ' + r.created })]), stars(r.stars)]),
      r.text ? h('p', { style: 'font-size:15px;line-height:22px', text: r.text }) : h('p', { class: 'small muted', text: 'Stars only, no text.' }),
      h('div', { class: 'tiny muted', text: 'Shown as: ' + r.showAs }),
      r.reply ? h('div', { class: 'small', style: 'padding:10px 12px;border-radius:10px;background:#F3F6F4' }, [h('div', { style: 'font-weight:600', text: 'Your reply' }), h('div', { text: r.reply })]) : null
    ];
  }
  out.push(h('h2', { class: 'sec-h amber', text: 'Reviews to check · ' + waiting.length }));
  if (!waiting.length) out.push(h('p', { class: 'small muted pad', text: 'Nothing to check.' }));
  waiting.forEach(function (r) {
    out.push(h('div', { class: 'card stack ocard hot' }, bits(r).concat([
      two(button('Publish', '', function () { act('adminReviewDecision', [r.id, true, ''], function () { toast('Published.'); }); }), button('Don\'t publish', 'sec', function () { hideSheet(r, false); })),
      button(r.reply ? 'Edit reply' : 'Reply', 'sec', function () { replySheet(r); })
    ])));
  });
  if (rest.length) out.push(h('h2', { class: 'sec-h', text: 'Checked' }));
  rest.forEach(function (r) {
    var pub = r.status === 'published';
    out.push(h('div', { class: 'card stack ocard' }, [h('span', { class: 'tag' + (pub ? '' : ' red'), style: 'align-self:flex-start', text: pub ? 'In the shop' : 'Not published · ' + r.reason })].concat(bits(r)).concat([
      pub ? two(button(r.reply ? 'Edit reply' : 'Reply', 'sec', function () { replySheet(r); }), button('Take down', 'warn', function () { hideSheet(r, true); })) : null
    ])));
  });
  out.push(h('div', { style: 'height:20px' }));
  return out;
}
function hideSheet(r, takeDown) {
  var reasons = takeDown ? ['Buyer asked to remove it', 'Abuse or offensive', 'Personal details', 'Health or cure claim', 'Other']
    : ['Health or cure claim', 'Abuse or offensive', 'Personal details', 'Not about the product', 'Spam or fake', 'Other'];
  var why = h('select', null, reasons.map(function (x) { return h('option', { value: x, text: x }); }));
  var close = sheet([
    h('h2', { text: takeDown ? 'Take this review down?' : 'Don\'t publish this review?' }),
    h('p', { class: 'small muted', text: 'It stays in your Sheet with the reason, but customers won\'t see it. Don\'t hide a review only because it is negative.' }),
    field('Reason', why),
    button(takeDown ? 'Take down' : 'Don\'t publish', 'warn', function () { close(); act('adminReviewDecision', [r.id, false, why.value]); }),
    button('Go back', 'sec', function () { close(); })
  ]);
}
function replySheet(r) {
  var ta = h('textarea', { rows: 4, maxlength: 600, placeholder: 'Thank the buyer, or answer their point' });
  ta.value = r.reply;
  var close = sheet([
    h('h2', { text: 'Reply to ' + r.showAs }),
    h('p', { class: 'small muted', text: 'Your reply is shown under the review once it is published. Leave it empty to remove your reply.' }),
    field('Your reply', ta),
    button('Save reply', '', function () { close(); act('adminReviewReply', [r.id, ta.value], function () { toast('Reply saved.'); }); }),
    button('Go back', 'sec', function () { close(); })
  ]);
}

/* ------------------------------------------------------------------ sales */
function viewSales() {
  var now = new Date(), month = (stats && stats.month) || (now.getFullYear() + '-' + ('0' + (now.getMonth() + 1)).slice(-2));
  var pick = h('input', { type: 'month', value: month, onchange: function (e) { if (e.target.value) load(e.target.value); } });
  function load(m) { api('adminStats', [m], function (r) { stats = r; render(); }); }
  if (!stats) setTimeout(function () { load(month); }, 0);
  var out = [header('Insights', null), h('div', { class: 'pad' }, [field('Month', pick)])];
  if (stats) {
    var r = stats;
    function tile(label, value) { return h('div', { class: 'tile' }, [h('span', { class: 'small muted', text: label }), h('b', { text: String(value) })]); }
    out.push(h('div', { class: 'grid2 pad', style: 'padding-top:16px;gap:10px' }, [
      tile('Money received', rs(r.sales)), tile('Orders', r.orders), tile('By UPI', rs(r.upi)), tile('By cash', rs(r.cash)),
      tile('Customers', r.customers), tile('Items sold', r.items), tile('New customers', r.newCustomers), tile('Returning customers', r.returningCustomers),
      tile('Delivered orders', r.shipped), tile('Collected orders', r.pickup), tile('In-person sales', r.inPerson), tile('Payments to check', r.toCheck),
      tile('Delivery fees', rs(r.shippingFees)), tile('Cancelled', r.cancelled), tile('Cash still due', rs(r.cashDue)), tile('Refunds due', rs(r.refundsDue)),
      tile('Reviews received', r.reviewsIn), tile('Average stars', r.avgStars == null ? '–' : r.avgStars + ' / 5'), tile('Reviews published', r.reviewsPublished), tile('Reviews waiting', r.reviewsWaiting)
    ]));
    if (r.top.length) out.push(h('div', { class: 'card stack', style: 'margin:16px 20px 0;gap:8px' }, [h('div', { style: 'font-weight:600', text: 'Best sellers' })].concat(r.top.map(function (t) {
      return h('div', { class: 'sum small' }, [h('span', { text: t.name }), h('span', { text: String(t.qty) })]);
    }))));
  }
  if (stats && stats.lowStock && stats.lowStock.length) out.push(h('div', { class: 'card stack', style: 'margin:16px 20px 0;gap:8px' }, [h('div', { style: 'font-weight:600', text: 'Running low right now' })].concat(stats.lowStock.map(function (t) {
    return h('div', { class: 'sum small' }, [h('span', { text: t.name }), h('span', { text: t.stock ? t.stock + ' left' : 'Sold out' })]);
  }))));
  out.push(h('p', { class: 'tiny muted pad', style: 'padding-top:16px', text: 'Customers are counted by phone number, from paid orders. The same numbers are on the Dashboard tab of your Google Sheet.' }));
  out.push(h('div', { class: 'pad', style: 'padding-top:20px;padding-bottom:20px' }, [button('Sign out', 'sec', signOut)]));
  return out;
}
render();
</script>
</body>
</html>
`;
