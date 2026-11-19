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
  ['shipAreaName', 'Bengaluru', 'Shown on the Ship to me option'],
  ['shipCities', 'Bengaluru, Bangalore', 'Cities customers can choose for shipping'],
  ['shipPinPrefixes', '560, 562', 'A shipping PIN code must start with one of these'],
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
