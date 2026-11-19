const { makeWorld } = require('./mock');
let pass = 0, fail = 0;
function ok(cond, label) { if (cond) pass++; else { fail++; console.log('FAIL: ' + label); } }
function throws(fn, re, label) { try { fn(); ok(false, label + ' (did not throw)'); } catch (e) { ok(re.test(String(e.message)), label + ' -> ' + e.message); } }
const JPG = 'data:image/jpeg;base64,' + Buffer.from('fake-jpeg-bytes').toString('base64');

const W = makeWorld();
// The owner made the screenshots folder by hand, with a plain hyphen and different capitals.
const handMade = W.makeFolder('store - payment Screenshots (Private)', W.home);
ok(/Setup finished/.test(W.run('setup')), 'setup runs');
ok(/Setup finished/.test(W.run('setup')), 'setup runs twice');
const S = () => Object.fromEntries(W.table('Settings').map((r) => [r.key, r.value]));
ok(S().proofFolderId === handMade.id, 'setup reuses the hand-made screenshots folder');
ok(Object.values(W.folders).filter((f) => /screenshots/i.test(f.name)).length === 1, 'no second screenshots folder');
ok(Object.values(W.folders).some((f) => f.name === 'Store – Product photos' && f.parent === W.home), 'product photo folder made next to the Sheet');
ok(/^\d{6}$/.test(String(S().adminPin)), 'admin PIN made');
ok(W.triggers.length === 1, 'one timer after two setups');
const pin = String(S().adminPin);
W.setting('upiId', 'store@upi'); W.setting('upiName', 'My Craft Store'); W.setting('pickupAddress', '12 Example Road'); W.setting('adminPhone', '+91 9800000000');
W.setting('appUrl', 'https://script.google.com/macros/s/TEST/exec');

// --- admin sign-in
ok(W.run('adminLogin', '000000').ok === false, 'wrong PIN refused');
const login = W.run('adminLogin', pin);
ok(login.ok && login.session.length >= 32, 'right PIN accepted');
const A = login.session;
throws(() => W.run('adminRefresh', 'nope'), /SIGNIN/, 'bad session refused');
throws(() => W.run('adminSetStock', '', 'P101', 5), /SIGNIN/, 'admin action without session refused');
for (let i = 0; i < 5; i++) W.run('adminLogin', '111111');
ok(/15 minutes/.test(W.run('adminLogin', pin).error || ''), 'locked after 5 wrong PINs');
W.advance(16);
ok(W.run('adminLogin', pin).ok, 'unlocked after 15 minutes');

// --- products
let d = W.run('adminSaveProduct', A, { name: 'Peony jar candle', category: 'Candles', price: 350, stock: 5, description: 'Soy wax', active: true, keep: [], add: [JPG, JPG] });
ok(d.products.length === 1 && d.products[0].photos.length === 2, 'product saved with two photos');
const p1 = d.products[0].id;
ok(W.files[d.products[0].photos[0]].sharing === 'anyone' && W.files[d.products[0].photos[0]].folder.name === 'Store – Product photos', 'product photo is link-viewable and in the photo folder');
d = W.run('adminSaveProduct', A, { name: 'Blue rose candle', category: 'Candles', price: 300, stock: 1, active: true, keep: [], add: [] });
const p2 = d.products[1].id;
d = W.run('adminSaveProduct', A, { name: 'Hidden thing', category: 'Other crafts', price: 100, stock: 3, active: false, keep: [], add: [] });
const p3 = d.products[2].id;
throws(() => W.run('adminSaveProduct', A, { name: 'x', price: 10, stock: 1 }), /name/, 'short product name refused');
throws(() => W.run('adminSaveProduct', A, { name: 'Free thing', price: 0, stock: 1 }), /price/, 'zero price refused');
throws(() => W.run('adminSaveProduct', A, { name: 'Many', price: 5, stock: 1, keep: [], add: [JPG, JPG, JPG, JPG, JPG] }), /4 photos/, 'more than 4 photos refused');
const firstPhoto = d.products[0].photos[0];
d = W.run('adminSaveProduct', A, { id: p1, name: 'Peony jar candle', category: 'Candles', price: 350, stock: 5, description: '=HYPERLINK("x")', active: true, keep: [d.products[0].photos[1]], add: [] });
ok(d.products[0].photos.length === 1 && W.files[firstPhoto].trashed, 'removed photo is binned');
ok(d.products[0].description.indexOf('=') !== 0, 'leading = stripped from description');
d = W.run('adminSetStock', A, p1, 6);
ok(d.products[0].stock === 6, 'stock set');
let boot = W.run('apiBoot');
ok(boot.products.length === 2 && !boot.products.some((p) => p.id === p3), 'hidden product not in the shop');
ok(boot.store.consentText.includes('My Craft Store') && !boot.store.consentText.includes('{store}'), 'consent text filled in');
ok(JSON.stringify(boot).indexOf(pin) < 0 && JSON.stringify(boot).indexOf('store@upi') < 0, 'shop data does not leak the PIN or UPI ID');

// --- order validation
const base = { items: [{ id: p1, qty: 1 }], mode: 'ship', pay: 'upi', name: 'Asha Verma', phone: '+91 98111 22333', email: 'asha@example.com', address: '45 MG Road, near the park', city: 'Bengaluru', pin: '560001', consent: true };
const place = (o) => W.run('apiPlaceOrder', Object.assign({}, base, o));
ok(/tick the box/.test(place({ consent: false }).error), 'no consent refused');
ok(/tick the box/.test(place({ consent: 'yes' }).error), 'consent must be true');
ok(/Bengaluru/.test(place({ pin: '400001' }).error), 'Mumbai PIN refused');
ok(/Bengaluru/.test(place({ city: 'Mumbai' }).error), 'city outside the list refused');
ok(/Bengaluru/.test(place({ pin: '2013' }).error), 'short PIN refused');
ok(/Cash is only/.test(place({ pay: 'cash' }).error), 'cash on delivery refused');
ok(/mobile/.test(place({ phone: '12' }).error), 'bad phone refused');
ok(/email/.test(place({ email: 'nope' }).error), 'bad email refused');
ok(/name/.test(place({ name: ' ' }).error), 'empty name refused');
ok(/address/.test(place({ address: 'x' }).error), 'short address refused');
ok(/empty/.test(place({ items: [] }).error), 'empty cart refused');
ok(/not right/.test(place({ items: [{ id: p1, qty: 0 }] }).error), 'zero quantity refused');
ok(/not right/.test(place({ items: [{ id: p1, qty: 1 }, { id: p1, qty: 1 }] }).error), 'duplicate line refused');
ok(/no longer available/.test(place({ items: [{ id: p3, qty: 1 }] }).error), 'hidden product cannot be ordered');
const tooMany = place({ items: [{ id: p2, qty: 2 }] });
ok(/Only 1 left/.test(tooMany.error) && tooMany.products.length === 2, 'over-stock refused with fresh stock');
ok(W.table('Orders').length === 0 && W.mails.length === 0, 'nothing saved or emailed by refused orders');

// --- UPI delivered order: Rs.10 below Rs.500
let r = place({});
ok(r.ok && r.order.status === 'awaiting_payment' && r.order.shipping === 10 && r.order.total === 360, 'order under 500 gets Rs.10 delivery');
const o1 = r.order;
ok(/^ORD-1001$/.test(o1.id), 'first order number');
ok(/am=360\.00/.test(o1.upiLink) && /pa=store%40upi/.test(o1.upiLink), 'UPI link has the amount and ID');
ok(W.run('apiBoot').products[0].stock === 5, 'stock reserved when the order is placed');
ok(W.mails.length === 0, 'no email yet for a UPI order');
let row = W.table('Orders')[0];
ok(row.phone === '+919811122333' && row.pin === '560001' && typeof row.pin === 'string', 'phone and PIN kept as text');
ok(row.consentAt instanceof Date || !!row.consentAt, 'consent time recorded');
throws(() => W.run('apiGetOrder', 'short'), /could not find/, 'bad token refused');
W.advance(4); W.run('tick');
ok(W.mails.length === 1 && /Finish your order/.test(W.mails[0].subject) && W.mails[0].to === 'asha@example.com', 'pay-link email after 3 minutes');
W.run('tick');
ok(W.mails.length === 1, 'pay-link email sent once');
throws(() => W.run('apiSubmitProof', o1.token, '', ''), /reference number or add a screenshot/, 'proof needs a reference or a screenshot');
throws(() => W.run('apiSubmitProof', o1.token, '12345', ''), /12 digits/, 'short reference refused');
throws(() => W.run('apiSubmitProof', o1.token, '', 'data:text/html;base64,AAAA'), /could not be read/, 'non-image proof refused');
r = W.run('apiSubmitProof', o1.token, '4123 4567 8901', '');
ok(r.status === 'payment_check' && /Your order is reserved, you'll receive an email when the order is confirmed/.test(r.message), 'reference accepted, reserved message shown');
ok(W.mails.length === 2 && /Payment to check: ORD-1001 Rs.360/.test(W.mails[1].subject) && W.mails[1].to === 'owner@example.com', 'admin emailed to check');
ok(/412345678901/.test(W.mails[1].body) && /page=admin&order=ORD-1001/.test(W.mails[1].body), 'admin email has the reference and a link');
r = W.run('apiSubmitProof', o1.token, '', JPG);
ok(r.status === 'payment_check' && W.mails.length === 3 && W.mails[2].inlineImages, 'screenshot can be added afterwards and is emailed inline');
const shot = W.table('Orders')[0].proofFileId;
ok(W.files[shot].folder === handMade && W.files[shot].sharing === 'private', 'screenshot stored privately in the screenshots folder');
ok(/^data:image\/jpeg;base64,/.test(W.run('adminProof', A, 'ORD-1001')), 'admin can view the screenshot');
throws(() => W.run('adminProof', 'bad', 'ORD-1001'), /SIGNIN/, 'screenshot needs sign-in');
throws(() => W.run('adminShip', A, 'ORD-1001', {}, true), /cannot be marked/, 'cannot ship before payment is confirmed');
W.advance(60); W.run('tick');
ok(W.table('Orders')[0].status === 'payment_check', 'hold does not expire while waiting for the check');
W.advance(180); W.run('tick'); W.run('tick');
ok(W.mails.filter((m) => /Still waiting for your check/.test(m.subject)).length === 1, 'admin nudged once after 3 hours');
d = W.run('adminPaymentDecision', A, 'ORD-1001', true);
ok(d.orders[0].status === 'confirmed' && d.orders[0].paid, 'payment received -> confirmed');
ok(/Order ORD-1001 confirmed/.test(W.mails[W.mails.length - 1].subject), 'customer emailed the confirmation');
throws(() => W.run('adminPaymentDecision', A, 'ORD-1001', true), /already been dealt with/, 'cannot decide twice');
throws(() => W.run('adminShip', A, 'ORD-1001', { tracking: 'javascript:alert(1)' }, true), /https/, 'bad tracking link refused');
throws(() => W.run('adminShip', A, 'ORD-1001', { driverPhone: 'abc' }, true), /driver phone/, 'bad driver phone refused');
let n = W.mails.length;
d = W.run('adminShip', A, 'ORD-1001', { provider: 'QuickDrop', tracking: 'https://example.com/track/abc', driverName: 'Ramesh', driverPhone: '9876543210', note: 'Arriving by 5 pm' }, true);
ok(d.orders[0].status === 'shipped' && d.orders[0].shipProvider === 'QuickDrop', 'marked as sent with porter details');
let m = W.mails[W.mails.length - 1];
ok(W.mails.length === n + 1 && /on its way/.test(m.subject) && /QuickDrop/.test(m.body) && /Ramesh, 9876543210/.test(m.body) && /example\.com\/track\/abc/.test(m.body), 'sent email carries the porter details');
let pub = W.run('apiGetOrder', o1.token);
ok(pub.status === 'shipped' && pub.ship.tracking === 'https://example.com/track/abc' && pub.ship.driverName === 'Ramesh', 'customer page shows delivery details');
n = W.mails.length;
W.run('adminShip', A, 'ORD-1001', { provider: 'QuickDrop', driverName: 'Suresh' }, false);
ok(W.mails.length === n && W.table('Orders')[0].driverName === 'Suresh', 'details edited without an email');
W.run('adminShip', A, 'ORD-1001', { provider: 'QuickDrop', driverName: 'Suresh' }, true);
ok(W.mails.length === n + 1 && /delivery update/.test(W.mails[n].subject), 'details edited with an email when asked');
throws(() => W.run('adminCancel', A, 'ORD-1001', 'x'), /no longer be cancelled/, 'shipped order cannot be cancelled');
n = W.mails.length;
d = W.run('adminDelivered', A, 'ORD-1001');
ok(W.mails.length === n + 1 && /How are you finding your order ORD-1001/.test(W.mails[n].subject) && W.mails[n].to === 'asha@example.com' && /\?review=/.test(W.mails[n].body), 'review invite emailed on delivery');
ok(d.orders.find((o) => o.id === 'ORD-1001').status === 'completed', 'marked delivered');
ok(W.run('apiBoot').products[0].stock === 5, 'stock stays sold after delivery');

// --- free delivery from Rs.500
r = place({ items: [{ id: p1, qty: 2 }] });
ok(r.ok && r.order.subtotal === 700 && r.order.shipping === 0 && r.order.total === 700, 'order of 500 or more ships free');
const o2 = r.order;
ok(W.run('apiBoot').products[0].stock === 3, 'two more reserved');
// hold runs out, then the customer pays late while stock is still there
W.advance(31); W.run('tick');
ok(W.table('Orders')[1].status === 'cancelled' && W.table('Orders')[1].cancelReason === 'Not paid in time', 'unpaid hold released after 30 minutes');
ok(W.run('apiBoot').products[0].stock === 5, 'stock returned when the hold ends');
ok(/was not completed/.test(W.mails[W.mails.length - 1].subject), 'customer told the hold ended');
r = W.run('apiSubmitProof', o2.token, '999988887777', '');
ok(r.status === 'payment_check' && W.run('apiBoot').products[0].stock === 3, 'late payment re-reserves the items when still in stock');
ok(/Paid after the hold ended/.test(W.table('Orders')[1].flags), 'late payment flagged for the admin');
d = W.run('adminPaymentDecision', A, o2.id, false);
ok(W.table('Orders')[1].status === 'cancelled' && W.run('apiBoot').products[0].stock === 5, 'not received -> cancelled and stock returned');
ok(/could not find your payment/.test(W.mails[W.mails.length - 1].subject), 'customer told payment was not found');

// late payment after the last piece sold
r = place({ items: [{ id: p2, qty: 1 }], mode: 'pickup' });
const o3 = r.order;
ok(r.ok && r.order.shipping === 0 && r.order.total === 300, 'pickup has no delivery fee');
ok(/sold out/.test(place({ items: [{ id: p2, qty: 1 }] }).error), 'last piece cannot be ordered twice');
W.advance(31); W.run('tick');
r = place({ items: [{ id: p2, qty: 1 }], mode: 'pickup', pay: 'cash', email: 'bina@example.com', name: 'Bina' });
const o4 = r.order;
ok(r.ok && r.order.status === 'cash_pickup', 'cash pickup order reserved at once');
ok(W.mails.slice(-2).some((x) => /is reserved/.test(x.subject) && x.to === 'bina@example.com') && W.mails.slice(-2).some((x) => /New pickup order/.test(x.subject)), 'cash pickup emails customer and admin');
ok(Math.round((r.order.holdUntil - W.clock) / 3600000) === 72, 'cash pickup held 72 hours');
r = W.run('apiSubmitProof', o3.token, '111122223333', JPG);
ok(r.status === 'cancelled' && r.refundDue && /have since sold/.test(r.message), 'late payment after sell-out goes to refund');
d = W.run('adminRefresh', A);
ok(d.orders.find((o) => o.id === o3.id).refundDue && /Refund needed/.test(W.mails[W.mails.length - 1].subject), 'admin sees the refund to make');
d = W.run('adminRefunded', A, o3.id);
ok(!d.orders.find((o) => o.id === o3.id).refundDue, 'refund marked as done');
// duplicate reference
r = place({ items: [{ id: p1, qty: 1 }] });
W.run('apiSubmitProof', r.order.token, '412345678901', '');
ok(/Same UPI reference as order ORD-1001/.test(W.table('Orders').find((x) => x.id === r.order.id).flags), 'reused reference number flagged');
d = W.run('adminPaymentDecision', A, r.order.id, true);
n = W.mails.length;
d = W.run('adminCancel', A, r.order.id, 'Out of stock');
ok(d.orders.find((o) => o.id === r.order.id).refundDue && /refunded to you/.test(W.mails[n].body), 'cancelling a paid order creates a refund and tells the customer');
ok(W.run('apiBoot').products[0].stock === 5, 'cancel returns stock');

// --- cash pickup: collect, expiry, customer cancel
n = W.mails.length;
d = W.run('adminCollected', A, o4.id, 'cash');
ok(W.mails.length === n + 1 && /How are you finding/.test(W.mails[n].subject) && W.mails[n].to === 'bina@example.com', 'review invite emailed on pickup');
ok(d.orders.find((o) => o.id === o4.id).status === 'completed' && d.orders.find((o) => o.id === o4.id).paidVia === 'cash', 'collected with cash');
W.run('adminSetStock', A, p2, 2);
r = place({ items: [{ id: p2, qty: 1 }], mode: 'pickup', pay: 'cash' });
const o5 = r.order;
W.advance(71 * 60); W.run('tick');
ok(W.table('Orders').find((x) => x.id === o5.id).status === 'cash_pickup', 'cash pickup still held at 71 hours');
W.advance(70); W.run('tick');
ok(W.table('Orders').find((x) => x.id === o5.id).status === 'cancelled' && W.run('apiBoot').products[1].stock === 2, 'cash pickup released after 72 hours');
ok(W.mails.slice(-2).some((x) => /no longer reserved/.test(x.subject)) && W.mails.slice(-2).some((x) => /was not collected/.test(x.subject)), 'both told about the release');
r = place({ items: [{ id: p2, qty: 1 }], mode: 'pickup', pay: 'cash' });
r = W.run('apiCancelOrder', r.order.token);
ok(r.status === 'cancelled' && W.run('apiBoot').products[1].stock === 2, 'customer can cancel a cash pickup');
throws(() => W.run('apiCancelOrder', o1.token), /no longer be cancelled/, 'customer cannot cancel a finished order');
// UPI pickup
r = place({ items: [{ id: p2, qty: 1 }], mode: 'pickup' });
W.run('apiSubmitProof', r.order.token, '', JPG);
W.run('adminPaymentDecision', A, r.order.id, true);
ok(/Collect from: 12 Example Road/.test(W.mails[W.mails.length - 1].body), 'paid pickup email has the address');
W.advance(100 * 60); W.run('tick');
ok(W.table('Orders').find((x) => x.id === r.order.id).status === 'confirmed', 'paid pickup is never auto-released');
d = W.run('adminCollected', A, r.order.id, '');
ok(d.orders.find((o) => o.id === r.order.id).paidVia === 'upi', 'paid pickup collected');

// --- in-person sale
n = W.mails.length;
d = W.run('adminInPersonSale', A, { items: [{ id: p1, qty: 2 }], via: 'cash', name: 'Walk in' });
const ips = d.orders.find((o) => o.source === 'inperson');
ok(d.products[0].stock === 3 && ips && ips.status === 'completed' && ips.paidVia === 'cash' && W.mails.length === n, 'in-person sale lowers stock, no email');
throws(() => W.run('adminInPersonSale', A, { items: [{ id: p1, qty: 9 }], via: 'cash' }), /Only 3 left/, 'in-person sale cannot exceed stock');

// --- sales
const month = W.run('adminStats', A, '');
ok(month.sales === 360 + 300 + 300 + 700 && month.cash === 300 + 700 && month.upi === 660, 'sales totals by UPI and cash: ' + JSON.stringify(month));
ok(month.shippingFees === 10 && month.inPerson === 1 && month.top[0].name === 'Peony jar candle', 'delivery fees, in-person count and best seller');

// --- insights: customers and the Sheet dashboard
ok(month.customers === 1 && month.newCustomers === 1 && month.returningCustomers === 0, 'customers counted by phone: ' + month.customers + '/' + month.newCustomers);
ok(Array.isArray(month.lowStock) && month.lowStock.some((x) => x.name === 'Blue rose candle'), 'low stock list');
const dash = () => W.sheets.Dashboard.rows;
const cell = (label) => (dash().find((r) => r && r[0] === label) || [])[1];
ok(cell('Money received (Rs.)') === month.sales && cell('  by UPI') === month.upi && cell('  by cash') === month.cash, 'Sheet dashboard matches the app');
ok(cell('Customers') === 1 && cell('Month (pick one)') === month.month, 'dashboard shows customers and the month');
ok(dash().some((r) => r && r[0] === month.month && r[1] === month.sales && r[4] === month.orders), 'month-by-month table has this month');
ok(dash().some((r) => r && r[3] === 'Peony jar candle'), 'dashboard lists best sellers');

// --- settings respected
W.setting('shippingFee', 40); W.setting('freeShippingFrom', 1000);
r = place({ items: [{ id: p1, qty: 2 }] });
ok(r.order.shipping === 40 && r.order.total === 740, 'changed delivery fee and threshold are used');
W.run('apiCancelOrder', r.order.token);
W.setting('adminAlerts', 'changes');
n = W.mails.length;
r = place({ items: [{ id: p1, qty: 1 }], mode: 'pickup', pay: 'cash' });
ok(W.mails.length === n + 1 && W.mails[n].to === 'asha@example.com', 'adminAlerts=changes skips the new-order email');
W.run('apiCancelOrder', r.order.token);
W.setting('upiId', '');
ok(/UPI payment is not set up/.test(place({}).error), 'UPI order refused when no UPI ID is set');
W.setting('upiId', 'store@upi');

// --- 90-day clean-up touches only the screenshots folder
const before = Object.values(W.files).filter((f) => !f.trashed);
const shots = before.filter((f) => f.folder === handMade).length, photos = before.filter((f) => f.folder !== handMade).length;
ok(shots >= 2 && photos >= 1, 'have screenshots and product photos to test with');
W.advance(70 * 24 * 60); W.run('tick');
ok(Object.values(W.files).filter((f) => !f.trashed && f.folder === handMade).length === shots, 'screenshots kept before 90 days');
const rowsBefore = W.table('Orders').length;
W.advance(21 * 24 * 60); W.run('tick');
ok(Object.values(W.files).filter((f) => !f.trashed && f.folder === handMade).length === 0, 'screenshots binned after 90 days');
ok(Object.values(W.files).filter((f) => !f.trashed && f.folder !== handMade).length === photos, 'product photos untouched by the clean-up');
ok(W.table('Orders').length === rowsBefore && W.table('Orders')[0].proofRef === '412345678901' && W.table('Orders')[0].proofFileId === '' && !!W.table('Orders')[0].proofDeletedAt, 'order rows kept, with the reference number; only the image link cleared');
throws(() => W.run('adminRefresh', A), /SIGNIN/, 'admin session ends after 30 days');
const A2 = W.run('adminLogin', pin).session;
ok(W.run('adminRefresh', A2).orders.find((o) => o.id === 'ORD-1001').shotDeleted, 'admin sees the screenshot was removed');

// --- next month: returning customers, and the month picker on the Sheet
const A3 = W.run('adminLogin', pin).session;
r = place({ items: [{ id: p1, qty: 1 }], mode: 'pickup', pay: 'cash' });
W.run('adminCollected', A3, r.order.id, 'cash');
r = place({ items: [{ id: p1, qty: 1 }], mode: 'pickup', pay: 'cash', phone: '9000011111', name: 'New Person', email: 'new@example.com' });
W.run('adminCollected', A3, r.order.id, 'cash');
const later = W.run('adminStats', A3, '');
ok(later.month !== month.month && later.customers === 2 && later.newCustomers === 1 && later.returningCustomers === 1, 'next month: one new and one returning customer ' + JSON.stringify([later.month, later.customers, later.newCustomers, later.returningCustomers]));
ok(cell('Month (pick one)') === later.month && cell('  returning') === 1, 'dashboard follows the current month');
ok(W.sheets.Dashboard.validation.list.includes(month.month) && W.sheets.Dashboard.validation.list.includes(later.month), 'month picker lists every month with orders');
W.sheets.Dashboard.rows[2][1] = month.month;
W.run('onEdit', { range: null });
W.ctx.onEdit({ range: W.sheets.Dashboard.getRange(3, 2) });
ok(cell('Month (pick one)') === month.month && cell('Money received (Rs.)') === month.sales, 'picking an older month in B3 redraws the dashboard');
W.run('adminSetStock', A3, p1, 4);
ok(cell('Month (pick one)') === month.month, 'the picked month is kept when the dashboard refreshes');
ok(dash().filter((r) => r && /^\d{4}-\d{2}$/.test(String(r[0]))).length >= 2, 'table has a row per month');

// --- reviews
const AR = W.run('adminLogin', pin).session;
W.run('adminSetStock', AR, p1, 20);
let info = W.run('apiReviewInfo', o1.token);
ok(info.open && info.items.length === 1 && info.items[0].id === p1 && info.shortName === 'Asha V.' && /My Craft Store/.test(info.consentText), 'review page for a delivered order');
const o2info = W.run('apiReviewInfo', o2.token);
ok(!o2info.open, 'cancelled order cannot be reviewed');
throws(() => W.run('apiSubmitReview', o2.token, { productId: p1, stars: 5, consent: true }), /cannot be reviewed/, 'review refused on a cancelled order');
throws(() => W.run('apiReviewInfo', 'x'), /could not find/, 'bad review link refused');
throws(() => W.run('apiSubmitReview', o1.token, { productId: p1, stars: 0, consent: true }), /1 to 5/, 'stars must be 1 to 5');
throws(() => W.run('apiSubmitReview', o1.token, { productId: p1, stars: 6, consent: true }), /1 to 5/, 'six stars refused');
throws(() => W.run('apiSubmitReview', o1.token, { productId: p1, stars: 5, consent: false }), /tick the box/, 'review needs consent');
throws(() => W.run('apiSubmitReview', o1.token, { productId: p2, stars: 5, consent: true }), /not in this order/, 'cannot review an item not bought');
n = W.mails.length;
info = W.run('apiSubmitReview', o1.token, { productId: p1, stars: 4, text: '=IMPORTRANGE("x") Lovely scent, burns evenly.', consent: true });
ok(info.items[0].done === 'waiting', 'review saved as waiting');
let rv = W.table('Reviews')[0];
ok(rv.status === 'waiting' && rv.showAs === 'Asha V.' && rv.stars === 4 && rv.text.indexOf('=') !== 0 && !!rv.consentAt && rv.orderId === 'ORD-1001', 'review row stored safely with consent time');
ok(W.mails.length === n + 1 && /New review to check/.test(W.mails[n].subject) && /tab=reviews/.test(W.mails[n].body), 'admin told about the new review');
throws(() => W.run('apiSubmitReview', o1.token, { productId: p1, stars: 5, consent: true }), /already reviewed/, 'one review per item per order');
ok(W.run('apiBoot').products.find((p) => p.id === p1).reviewCount === 0, 'waiting review not shown in the shop');
d = W.run('adminRefresh', AR);
ok(d.reviews.length === 1 && d.reviews[0].status === 'waiting' && d.reviews[0].customerName === 'Asha Verma', 'admin sees the review to check');
throws(() => W.run('adminReviewDecision', 'bad', rv.id, true, ''), /SIGNIN/, 'review decision needs sign-in');
d = W.run('adminReviewDecision', AR, rv.id, true, '');
throws(() => W.run('adminReviewDecision', AR, rv.id, true, ''), /already been dealt with/, 'cannot publish twice');
d = W.run('adminReviewReply', AR, rv.id, 'Thank you, Asha!');
boot = W.run('apiBoot');
let pr = boot.products.find((p) => p.id === p1);
ok(pr.reviewCount === 1 && pr.avgStars === null && pr.reviews[0].by === 'Asha V.' && pr.reviews[0].reply === 'Thank you, Asha!' && /Lovely scent/.test(pr.reviews[0].text), 'published review shown, no average below 3 reviews');
ok(JSON.stringify(pr.reviews).indexOf('Verma') < 0 && JSON.stringify(pr.reviews).indexOf('ORD-1001') < 0, 'shop never shows surname or order number');
ok(W.run('apiGetOrder', o1.token).canReview, 'order page offers a review');
// two more reviewers of the same product; one anonymous, one with a cure claim that is not published
function buyAndFinish(name, phone, email) {
  const r1 = place({ items: [{ id: p1, qty: 1 }], mode: 'pickup', pay: 'cash', name, phone, email });
  W.run('adminCollected', AR, r1.order.id, 'cash');
  return r1.order.token;
}
const t2 = buyAndFinish('Chitra Nair', '9811100001', 'chitra@example.com');
const t3 = buyAndFinish('Dev Rao', '9811100002', 'dev@example.com');
const t4 = buyAndFinish('Esha Paul', '9811100003', 'esha@example.com');
W.run('apiSubmitReview', t2, { productId: p1, stars: 5, text: 'Beautiful', consent: true, anonymous: true });
W.run('apiSubmitReview', t3, { productId: p1, stars: 2, text: 'Smaller than I expected', consent: true });
W.run('apiSubmitReview', t4, { productId: p1, stars: 5, text: 'This candle cured my anxiety', consent: true });
const rows = W.table('Reviews');
ok(rows[1].showAs === 'A verified buyer' && rows[1].customerName === 'Chitra Nair', 'anonymous review keeps the name private');
W.run('adminReviewDecision', AR, rows[1].id, true, '');
W.run('adminReviewDecision', AR, rows[2].id, true, '');
W.run('adminReviewDecision', AR, rows[3].id, false, 'Health or cure claim');
pr = W.run('apiBoot').products.find((p) => p.id === p1);
ok(pr.reviewCount === 3 && pr.avgStars === 3.7, 'average shown from 3 published reviews: ' + pr.avgStars);
ok(!pr.reviews.some((x) => /cured/.test(x.text)) && pr.reviews.some((x) => x.stars === 2), 'cure claim hidden, negative review published');
ok(W.table('Reviews')[3].status === 'hidden' && W.table('Reviews')[3].reason === 'Health or cure claim', 'hidden review kept with its reason');
W.run('adminReviewDecision', AR, rows[1].id, false, 'Buyer asked to remove it');
pr = W.run('apiBoot').products.find((p) => p.id === p1);
ok(pr.reviewCount === 2 && pr.avgStars === null, 'taken-down review removed and average hidden again');
throws(() => W.run('adminReviewDecision', AR, rows[1].id, false, 'x'), /already hidden/, 'cannot hide twice');
W.setting('reviewMinForStars', 2);
ok(W.run('apiBoot').products.find((p) => p.id === p1).avgStars === 3, 'star threshold is a setting');
W.setting('reviewMinForStars', 3);
W.setting('reviewsOn', 'no');
ok(W.run('apiBoot').products.find((p) => p.id === p1).reviewCount === 0 && !W.run('apiReviewInfo', t2).open, 'reviews can be switched off');
n = W.mails.length;
const t5 = buyAndFinish('Off Person', '9811100009', 'off@example.com');
ok(!W.mails.slice(n).some((m) => /How are you finding/.test(m.subject)), 'no invite while reviews are off');
W.setting('reviewsOn', 'yes');
const reviewMonth = W.run('adminStats', AR, '');
ok(reviewMonth.reviewsIn === 4 && reviewMonth.reviewsPublished === 2 && reviewMonth.reviewsHidden === 2 && reviewMonth.avgStars === 3, 'insights count reviews: ' + JSON.stringify([reviewMonth.reviewsIn, reviewMonth.reviewsPublished, reviewMonth.avgStars]));
// in-person sales are not invited
n = W.mails.length;
W.run('adminInPersonSale', AR, { items: [{ id: p1, qty: 1 }], via: 'cash', name: 'Walk in' });
ok(W.mails.length === n, 'in-person sale gets no review invite');

// --- email limit and pages
W.quota = 0; n = W.mails.length;
r = place({ items: [{ id: p1, qty: 1 }], mode: 'pickup', pay: 'cash' });
ok(r.ok && W.mails.length === n && W.table('Log').some((x) => /daily limit/.test(x.what)), 'order still works when the email limit is reached, and it is logged');
W.quota = 50;
let page = W.run('doGet', { parameter: {} });
ok(/My Craft Store/.test(page.title) && page.html.indexOf('"__BOOT__"') < 0 && /var BOOT = \{"shop"/.test(page.html), 'shop page is built with its data');
page = W.run('doGet', { parameter: { page: 'admin', order: '<script>x</script>' } });
ok(page.html.indexOf('<script>x') < 0 && page.html.indexOf(pin) < 0, 'admin page escapes the order parameter and has no PIN');
page = W.run('doGet', { parameter: { order: o1.token } });
ok(/"status":"completed"/.test(page.html), 'order link opens the order');
ok(!W.table('Orders').some((x) => Object.values(x).some((v) => typeof v === 'string' && /^=/.test(v))), 'no formulas in the Orders tab');
ok(W.table('Orders').every((x) => typeof x.total === 'number'), 'totals are numbers in the Sheet');

console.log(pass + ' passed, ' + fail + ' failed');
process.exit(fail ? 1 : 0);
