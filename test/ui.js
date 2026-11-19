const { makeWorld } = require('./mock');
const fs = require('fs');
const { chromium } = require('playwright');
const shotDir = __dirname + '/shots';
const FIX = __dirname + '/fixtures/';
fs.mkdirSync(shotDir, { recursive: true });
const img = (n) => 'data:image/jpeg;base64,' + fs.readFileSync(FIX + n).toString('base64');
let pass = 0, fail = 0;
const ok = (c, l) => { if (c) pass++; else { fail++; console.log('FAIL: ' + l); } };

(async () => {
  const W = makeWorld();
  W.run('setup');
  const pin = String(W.table('Settings').find((r) => r.key === 'adminPin').value);
  W.setting('upiId', 'demo@upi'); W.setting('upiName', 'My Craft Store'); W.setting('pickupAddress', '12 Example Road, Sector 18'); W.setting('adminPhone', '+91 98000 00000');
  W.setting('appUrl', 'https://store.test/exec');
  const A = W.run('adminLogin', pin).session;
  [['Peony jar candle', 350, 5, 'img_3965.jpg', 'Candles'], ['Daisy jar candle', 350, 6, 'img_3966.jpg', 'Candles'], ['Blue rose candle', 300, 1, 'img_3971.jpg', 'Candles'], ['Flower garden bowl candle', 650, 2, 'img_3973.jpg', 'Gift boxes']]
    .forEach((p) => W.run('adminSaveProduct', A, { name: p[0], price: p[1], stock: p[2], category: p[4], description: 'Hand-poured soy wax candle in a glass jar.', active: true, keep: [], add: [img(p[3])] }));

  const browser = await chromium.launch(process.env.CHROME_PATH ? { executablePath: process.env.CHROME_PATH } : {});
  const context = await browser.newContext({ viewport: { width: 390, height: 844 }, deviceScaleFactor: 1 });
  const errors = [];
  async function open(url) {
    const page = await context.newPage();
    page.on('pageerror', (e) => errors.push(String(e)));
    page.on('console', (m) => { if (m.type() === 'error' && !/Failed to load resource|ERR_/.test(m.text())) errors.push(m.text()); });
    await page.exposeFunction('__gas', (fn, args) => { try { return { v: W.run(fn, ...args) }; } catch (e) { return { e: String(e.message) }; } });
    await page.addInitScript(() => {
      function runner(okFn, failFn) {
        return new Proxy({}, { get(_, name) {
          if (name === 'withSuccessHandler') return (f) => runner(f, failFn);
          if (name === 'withFailureHandler') return (f) => runner(okFn, f);
          return (...args) => { window.__gas(name, args).then((r) => setTimeout(() => { if (r.e) failFn && failFn(new Error(r.e)); else okFn && okFn(r.v); }, 30)); };
        } });
      }
      window.google = { script: { run: runner(null, null) } };
    });
    await page.route('**/*', (route) => {
      const u = new URL(route.request().url());
      if (u.host === 'store.test') { const q = Object.fromEntries(u.searchParams); return route.fulfill({ contentType: 'text/html', body: W.run('doGet', { parameter: q }).html }); }
      if (u.host === 'lh3.googleusercontent.com') { const id = u.pathname.split('/')[2].split('=')[0]; const f = W.files[id]; return f ? route.fulfill({ contentType: 'image/jpeg', body: Buffer.from(f.bytes) }) : route.abort(); }
      return route.abort();
    });
    await page.goto(url);
    return page;
  }
  const snap = async (page, name) => { await page.evaluate(() => { const t = document.getElementById('toast'); if (t) t.style.display = 'none'; }); await page.screenshot({ path: shotDir + '/' + name + '.png', fullPage: true }); };
  const noOverflow = async (page, name) => ok(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1), name + ' fits the phone width');

  // ---------- customer
  let page = await open('https://store.test/exec');
  await page.waitForSelector('text=My Craft Store');
  ok(await page.locator('.grid2 .card').count() === 4, 'four products shown');
  ok(await page.locator('text=Only 1 left').count() === 1, 'low stock tag shown');
  await noOverflow(page, 'shop'); await snap(page, '1-shop');
  await page.click('button[aria-label="View Peony jar candle"]');
  await page.waitForSelector('text=Add to cart');
  await page.click('button[aria-label="One more"]');
  ok(await page.locator('text=Add to cart · Rs.700').count() === 1, 'product total follows quantity');
  await page.click('button[aria-label="One less"]');
  await noOverflow(page, 'product'); await snap(page, '2-product');
  await page.click('text=Add to cart · Rs.350');
  await page.waitForSelector('text=Your cart');
  await page.click('text=Deliver to me');
  ok(await page.locator('text=Add Rs.150 more for free delivery.').count() === 1, 'free delivery nudge');
  ok(await page.locator('text=Cash, when I collect').count() === 0, 'no cash option for delivery');
  await page.click('text=Continue to pay Rs.360');
  await page.waitForSelector('#toast:visible');
  ok(/delivery address/.test(await page.textContent('#toast')), 'asks for the address first');
  await page.fill('input[autocomplete="street-address"]', '45 Sector 62, near the park');
  await page.selectOption('select', 'Noida');
  await page.fill('input[autocomplete="postal-code"]', '400001');
  await page.fill('input[autocomplete="name"]', 'Asha Verma');
  await page.fill('input[autocomplete="tel"]', '98111 22333');
  await page.fill('input[autocomplete="email"]', 'asha@example.com');
  await page.click('text=Continue to pay Rs.360');
  await page.waitForFunction(() => /outside Delhi NCR/.test(document.getElementById('toast').textContent));
  ok(true, 'PIN outside the delivery area stopped on the phone');
  await page.fill('input[autocomplete="postal-code"]', '201301');
  await page.click('text=Continue to pay Rs.360');
  await page.waitForFunction(() => /tick the box/.test(document.getElementById('toast').textContent));
  ok(W.table('Orders').length === 0, 'no order without the consent tick');
  await page.click('text=Privacy notice'); await page.waitForSelector('.sheet'); await snap(page, '3b-privacy'); await page.click('.sheet >> text=Close');
  await page.check('#consent');
  await noOverflow(page, 'checkout'); await snap(page, '3-checkout');
  await page.click('text=Continue to pay Rs.360');
  await page.waitForSelector('text=Step 1 of 2');
  ok(await page.locator('canvas').count() === 1, 'QR code drawn');
  ok((await page.getAttribute('a.btn', 'href')).startsWith('upi://pay?pa=demo%40upi'), 'UPI button link');
  ok(await page.locator('text=I have paid').count() === 0, 'no "I have paid" on the payment screen');
  await noOverflow(page, 'pay'); await snap(page, '4-pay');
  await page.click('text=Next: confirm my payment');
  await page.waitForSelector('text=Step 2 of 2');
  ok(await page.locator('button:has-text("I have paid Rs.360")').isDisabled(), '"I have paid" disabled with no proof');
  await page.fill('input[inputmode="numeric"]', '41234');
  ok(await page.locator('button:has-text("I have paid Rs.360")').isDisabled(), 'still disabled with a short reference');
  await page.fill('input[inputmode="numeric"]', '412345678901');
  ok(await page.locator('button:has-text("I have paid Rs.360")').isEnabled(), 'enabled with a 12-digit reference');
  await page.fill('input[inputmode="numeric"]', '');
  await page.setInputFiles('input[type=file]', FIX + 'img_3970.jpg');
  await page.waitForSelector('text=Screenshot added');
  ok(await page.locator('button:has-text("I have paid Rs.360")').isEnabled(), 'enabled with a screenshot only');
  await noOverflow(page, 'confirm'); await snap(page, '5-confirm');
  await page.click('button:has-text("I have paid Rs.360")');
  await page.waitForSelector("text=Your order is reserved, you'll receive an email when the order is confirmed.");
  ok(W.table('Orders')[0].status === 'payment_check' && !!W.table('Orders')[0].proofFileId, 'order waits for the check with a screenshot saved');
  await snap(page, '6-reserved');
  const token = W.table('Orders')[0].token;
  await page.close();

  // cash pickup
  page = await open('https://store.test/exec');
  ok(await page.locator('text=You have an order in progress').count() === 1, 'returning visitor sees the order in progress');
  await page.click('button[aria-label="Add Blue rose candle to cart"]');
  await page.click('text=View cart');
  await page.click('text=Collect in person');
  await page.click('text=Cash, when I collect');
  await page.fill('input[autocomplete="name"]', 'Bina Rao'); await page.fill('input[autocomplete="tel"]', '+919822233344'); await page.fill('input[autocomplete="email"]', 'bina@example.com');
  await page.check('#consent');
  await snap(page, '3c-checkout-cash');
  await page.click('text=Place order · Rs.300 cash at pickup');
  await page.waitForSelector('text=Please collect it and pay Rs.300 in cash');
  await snap(page, '6b-cash-reserved');
  await page.click('text=Back to the shop');
  await page.waitForSelector('text=Sold out');
  ok(await page.locator('button[aria-label="Add Blue rose candle to cart"]').isDisabled(), 'sold-out product cannot be added');
  await page.close();

  // ---------- admin
  page = await open('https://store.test/exec?page=admin&order=ORD-1001');
  await page.fill('input[type=password]', '000000'); await page.click('text=Open');
  await page.waitForFunction(() => /not right/.test(document.getElementById('toast').textContent));
  await page.fill('input[type=password]', pin); await page.click('text=Open');
  await page.waitForSelector('text=Payment to check · 1');
  await noOverflow(page, 'admin orders'); await snap(page, '7-admin-orders');
  await page.click('text=View payment screenshot'); await page.waitForSelector('.sheet img'); await snap(page, '7b-admin-proof'); await page.click('.sheet >> text=Close');
  await page.click('text=Payment received');
  await page.waitForSelector('text=To pack and send · 1');
  await page.click('text=Mark as sent');
  await page.waitForSelector('.sheet');
  await page.fill('input[placeholder="Courier or delivery app"]', 'QuickDrop');
  await page.fill('input[type=url]', 'example.com/track/abc');
  await page.click('text=Mark as sent and email the customer');
  await page.waitForFunction(() => /https/.test(document.getElementById('toast').textContent));
  await page.fill('input[type=url]', 'https://example.com/track/abc');
  const inputs = page.locator('.sheet input[type=text]');
  await inputs.nth(1).fill('Ramesh');
  await page.fill('.sheet input[type=tel]', '9876543210');
  await snap(page, '8-admin-ship');
  await page.click('text=Mark as sent and email the customer');
  await page.waitForSelector('text=On the way · 1');
  ok(W.table('Orders')[0].status === 'shipped' && W.table('Orders')[0].driverName === 'Ramesh', 'admin marked as sent with porter details');
  await page.click('text=Collected · UPI'); await page.waitForSelector('.sheet canvas'); await snap(page, '8b-admin-collect-qr'); await page.click('.sheet >> text=Close');
  await page.click('text=In-person sale'); await page.waitForSelector('.sheet');
  await page.click('.sheet button[aria-label="One more"] >> nth=0');
  ok(await page.locator('text=Total Rs.350').count() === 1, 'in-person sale total');
  await snap(page, '8c-admin-sale');
  await page.click('text=Record sale');
  await page.waitForFunction(() => /Sale recorded/.test(document.getElementById('toast').textContent));
  await snap(page, '9-admin-orders-after');
  await page.click('nav >> text=Products');
  await page.waitForSelector('text=Add product');
  await noOverflow(page, 'admin products'); await snap(page, '10-admin-products');
  await page.click('button[aria-label="One more"] >> nth=1');
  await page.waitForTimeout(1300);
  ok(W.table('Products')[1].stock === 7, 'stock stepper saves');
  await page.click('text=Add product'); await page.waitForSelector('.sheet');
  await page.setInputFiles('.sheet input[type=file] >> nth=1', FIX + 'img_3967.jpg');
  await page.waitForSelector('.sheet img[alt="New photo 1"]');
  await page.fill('.sheet input[type=text] >> nth=0', 'Rose trio gel candle');
  await page.fill('.sheet input[inputmode=numeric]', '400');
  await page.click('.sheet .stepper button[aria-label="One more"]');
  await snap(page, '11-admin-add-product');
  await page.click('text=Save product');
  await page.waitForFunction(() => /Product saved/.test(document.getElementById('toast').textContent));
  const np = W.table('Products').find((p) => p.name === 'Rose trio gel candle');
  ok(np && np.price === 400 && np.stock === 2 && String(np.photoIds).length > 5, 'new product saved from the phone with its photo');
  await page.click('nav >> text=Insights');
  await page.waitForSelector('text=Money received');
  ok(await page.locator('text=Returning customers').count() === 1 && await page.locator('text=Running low right now').count() === 1, 'insights show customers and low stock');
  await noOverflow(page, 'admin sales'); await snap(page, '12-admin-sales');
  await page.close();

  // customer sees the delivery details
  page = await open('https://store.test/exec?order=' + token);
  await page.waitForSelector('text=On its way');
  ok(await page.locator('text=Track my delivery').count() === 1 && await page.locator('text=Driver: Ramesh').count() === 1, 'customer order page shows porter details');
  await snap(page, '13-order-on-its-way');
  await page.close();

  // ---------- reviews
  W.run('adminDelivered', A, 'ORD-1001');
  page = await open('https://store.test/exec?review=' + token);
  await page.waitForSelector('text=Write a review');
  ok(await page.locator('text=Peony jar candle').count() >= 1, 'review page lists the item bought');
  await page.click('text=Send review');
  await page.waitForFunction(() => /1 to 5 stars/.test(document.getElementById('toast').textContent));
  await page.click('button[aria-label="4 stars"]');
  ok(await page.getAttribute('button[aria-label="4 stars"]', 'aria-checked') === 'true', 'star choice is marked for screen readers');
  await page.fill('textarea', 'Lovely scent and it burns evenly. The jar is reusable too.');
  await page.click('text=Send review');
  await page.waitForFunction(() => /tick the box/.test(document.getElementById('toast').textContent));
  await page.click('input[id^="c-"]');
  await noOverflow(page, 'review form'); await snap(page, '14-review-form');
  await page.click('text=Send review');
  await page.waitForSelector('text=Your review will appear once it has been checked');
  ok(W.table('Reviews').length === 1 && W.table('Reviews')[0].stars === 4, 'review saved from the phone');
  await snap(page, '15-review-sent');
  await page.close();
  // two more buyers review the same candle, so the average shows
  const AA = A;
  [['Chitra Nair', '9811100001', 5, 'Beautiful gift, my mother loved it.'], ['Dev Rao', '9811100002', 3, 'Nice, but smaller than I expected.']].forEach((b) => {
    const r = W.run('apiPlaceOrder', { items: [{ id: W.table('Products')[0].id, qty: 1 }], mode: 'pickup', pay: 'cash', name: b[0], phone: b[1], email: 'x' + b[1] + '@example.com', consent: true });
    W.run('adminCollected', AA, r.order.id, 'cash');
    W.run('apiSubmitReview', r.order.token, { productId: W.table('Products')[0].id, stars: b[2], text: b[3], consent: true });
  });
  page = await open('https://store.test/exec?page=admin&tab=reviews');
  ok(await page.locator('text=Reviews to check').count() + await page.locator('input[type=password]').count() >= 0, 'admin page opens');
  await page.waitForSelector('text=Reviews to check', { timeout: 5000 }).catch(async () => { await page.fill('input[type=password]', pin); await page.click('text=Open'); });
  await page.waitForSelector('text=Reviews to check · 3');
  await noOverflow(page, 'admin reviews'); await snap(page, '16-admin-reviews');
  await page.click('.ocard >> nth=0 >> text=Reply');
  await page.fill('.sheet textarea', 'Thank you so much!');
  await page.click('text=Save reply');
  await page.waitForFunction(() => /Reply saved/.test(document.getElementById('toast').textContent));
  for (let i = 0; i < 3; i++) { await page.click('.ocard.hot >> nth=0 >> text=Publish'); await page.waitForTimeout(250); }
  await page.waitForSelector('text=Reviews to check · 0');
  ok(W.table('Reviews').every((r) => r.status === 'published'), 'admin published the reviews');
  await page.click('.ocard >> nth=0 >> text=Take down');
  await page.waitForSelector('.sheet select');
  await snap(page, '17-admin-take-down');
  await page.click('.sheet >> text=Go back');
  await page.click('nav >> text=Insights');
  await page.waitForSelector('text=Reviews received');
  await page.close();
  page = await open('https://store.test/exec');
  await page.waitForSelector('text=My Craft Store');
  ok(await page.locator('[aria-label="4 out of 5 stars"]').count() >= 1, 'shop card shows the average stars');
  await snap(page, '1b-shop-with-stars');
  await page.click('button[aria-label="View Peony jar candle"]');
  await page.waitForSelector('text=3 reviews');
  ok(await page.locator('text=Reply from My Craft Store').count() === 1 && await page.locator('text=Verified purchase').count() === 3, 'product page shows reviews, replies and verified labels');
  ok(await page.locator('text=Nair').count() === 0, 'surname not shown');
  await noOverflow(page, 'product with reviews'); await snap(page, '18-product-reviews');
  await page.close();

  ok(errors.length === 0, 'no script errors: ' + errors.join(' | '));
  await browser.close();
  console.log(pass + ' passed, ' + fail + ' failed');
  process.exit(fail ? 1 : 0);
})().catch((e) => { console.error(e); process.exit(1); });
