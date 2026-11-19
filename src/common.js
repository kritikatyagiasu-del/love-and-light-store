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
