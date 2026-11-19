// Builds dist/Code.gs from src/. Run: node build.js
const fs = require('fs');
const path = require('path');
const { minify_sync } = require('terser');
const src = (f) => fs.readFileSync(path.join(__dirname, 'src', f), 'utf8');

const qr = minify_sync(fs.readFileSync(path.join(__dirname, 'node_modules/qrcode-generator/dist/qrcode.js'), 'utf8')).code;
const css = src('common.css'), common = src('common.js');

function page(name) {
  const html = src(name)
    .replace('/*__CSS__*/', () => css)
    .replace('/*__QR__*/', () => '/* QR code generator: qrcode-generator by Kazuhiko Arase, MIT licence */\n' + qr)
    .replace('/*__COMMON__*/', () => common);
  if (html.includes('`') || html.includes('${')) throw new Error(name + ' contains a backtick or ${, which would break the template string');
  if ((html.match(/<\/script>/g) || []).length !== 2) throw new Error(name + ' has an unexpected </script>');
  if (/\\$/.test(html)) throw new Error(name + ' ends with a backslash');
  return html;
}
const out = src('server.js') +
  '\n/* ------------------------------------------------------------ shop page */\n\nvar SHOP_HTML = String.raw`' + page('shop.html') + '`;\n' +
  '\n/* ----------------------------------------------------------- admin page */\n\nvar ADMIN_HTML = String.raw`' + page('admin.html') + '`;\n';
fs.mkdirSync(path.join(__dirname, 'dist'), { recursive: true });
fs.writeFileSync(path.join(__dirname, 'dist', 'Code.gs'), out);
fs.writeFileSync(path.join(__dirname, 'dist', 'appsscript.json'), JSON.stringify({
  timeZone: 'Asia/Kolkata', dependencies: {}, exceptionLogging: 'STACKDRIVER', runtimeVersion: 'V8',
  webapp: { executeAs: 'USER_DEPLOYING', access: 'ANYONE_ANONYMOUS' }
}, null, 2) + '\n');
console.log('Code.gs', out.length, 'chars,', out.split('\n').length, 'lines');
