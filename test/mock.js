// In-memory stand-ins for the Google services Code.gs uses, so the app can run in Node.
const vm = require('vm');
const fs = require('fs');
const path = require('path');

function makeWorld() {
  const W = { mails: [], clock: Date.parse('2026-11-16T05:00:00Z'), files: {}, folders: {}, quota: 100, triggers: [], props: {}, n: 0 };
  const id = (p) => p + (++W.n).toString().padStart(4, '0') + 'xxxxxxxxxxxxxxxxxxxx';

  class Range {
    constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr: nr || 1, nc: nc || 1 }); }
    getValues() { const out = []; for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(this.sh.cell(this.r + i, this.c + j)); out.push(row); } return out; }
    setValues(v) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sh.put(this.r + i, this.c + j, v[i][j]); return this; }
    setValue(x) { this.sh.put(this.r, this.c, x); return this; }
    setNumberFormat(f) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sh.fmt[(this.r + i) + ':' + (this.c + j)] = f; return this; }
    setNumberFormats(f) { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) this.sh.fmt[(this.r + i) + ':' + (this.c + j)] = f[i][j]; return this; }
    setFontWeight() { return this; }
    getValue() { return this.sh.cell(this.r, this.c); }
    clearContent() { for (let i = 0; i < this.nr; i++) for (let j = 0; j < this.nc; j++) if (this.sh.rows[this.r + i - 1]) this.sh.rows[this.r + i - 1][this.c + j - 1] = ''; return this; }
    setDataValidation(v) { this.sh.validation = v; return this; }
    getSheet() { return this.sh; } getRow() { return this.r; } getColumn() { return this.c; }
  }
  class Sheet {
    constructor(name) { this.name = name; this.rows = []; this.fmt = {}; }
    cell(r, c) { const row = this.rows[r - 1]; const v = row ? row[c - 1] : ''; return v === undefined ? '' : v; }
    put(r, c, v) {
      while (this.rows.length < r) this.rows.push([]);
      const f = this.fmt[r + ':' + c];
      // Like Sheets: a text-formatted cell keeps what is typed; otherwise formulas and numbers are interpreted.
      if (typeof v === 'string' && f !== '@') {
        if (/^=/.test(v)) throw new Error('FORMULA written to ' + this.name + ' ' + r + ':' + c + ' -> ' + v);
        if (/^[+-]?\d+(\.\d+)?$/.test(v)) v = Number(v);
      }
      this.rows[r - 1][c - 1] = v;
    }
    getLastRow() { let n = this.rows.length; while (n > 0 && !(this.rows[n - 1] || []).some((x) => x !== '' && x != null)) n--; return n; }
    getLastColumn() { return this.rows.reduce((m, r) => Math.max(m, r.length), 0); }
    getDataRange() { return new Range(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
    getRange(r, c, nr, nc) { return new Range(this, r, c, nr, nc); }
    appendRow(v) { const r = this.getLastRow() + 1; v.forEach((x, j) => this.put(r, j + 1, x)); }
    setFrozenRows() { } setColumnWidth() { } getName() { return this.name; }
  }
  const sheets = {};
  const ss = {
    getId: () => 'SHEETFILE', getSheetByName: (n) => sheets[n] || null,
    insertSheet: (n) => (sheets[n] = new Sheet(n)),
  };
  W.sheets = sheets;

  function iter(list) { let i = 0; return { hasNext: () => i < list.length, next: () => list[i++] }; }
  function makeFolder(name, parent) {
    const f = { id: id('FOLDER'), name, parent, trashed: false, getId() { return f.id; }, getName() { return f.name; }, isTrashed() { return f.trashed; },
      getFolders() { return iter(Object.values(W.folders).filter((x) => x.parent === f && !x.trashed)); },
      createFolder(n) { return makeFolder(n, f); },
      getFiles() { return iter(Object.values(W.files).filter((x) => x.folder === f && !x.trashed)); },
      createFile(blob) {
        const file = { id: id('FILE'), name: blob.name, bytes: blob.bytes, folder: f, trashed: false, created: new Date(W.clock), sharing: 'private',
          getId() { return file.id; }, setTrashed(t) { file.trashed = t; return file; }, getDateCreated() { return file.created; },
          getBlob() { return { getBytes: () => file.bytes }; }, setSharing(a) { file.sharing = a; return file; }, getName() { return file.name; } };
        W.files[file.id] = file; return file;
      } };
    W.folders[f.id] = f; return f;
  }
  W.root = makeFolder('My Drive', null);
  W.home = makeFolder('My Store', W.root);
  W.makeFolder = makeFolder;

  const ctx = {
    console, JSON, Math, Date: class extends Date { constructor(...a) { if (a.length) super(...a); else super(W.clock); } static now() { return W.clock; } },
    SpreadsheetApp: { getActiveSpreadsheet: () => ss, flush() { }, newDataValidation: () => ({ list: null, requireValueInList(l) { this.list = l; return this; }, setAllowInvalid() { return this; }, build() { return this; } }), getUi: () => ({ createMenu: () => ({ addItem() { return this; }, addToUi() { } }), alert() { }, ButtonSet: { OK: 1 } }) },
    DriveApp: {
      Access: { ANYONE_WITH_LINK: 'anyone' }, Permission: { VIEW: 'view' }, getRootFolder: () => W.root,
      getFileById(i) { if (i === 'SHEETFILE') return { getParents: () => iter([W.home]) }; const f = W.files[i]; if (!f) throw new Error('no file ' + i); return f; },
      getFolderById(i) { const f = W.folders[i]; if (!f) throw new Error('no folder ' + i); return f; },
    },
    Utilities: {
      getUuid: () => 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, () => Math.floor(Math.random() * 16).toString(16)),
      base64Decode: (s) => Array.from(Buffer.from(s, 'base64')), base64Encode: (b) => Buffer.from(b).toString('base64'),
      newBlob: (bytes, type, name) => ({ bytes, type, name }),
      formatDate(d, tz, f) {
        const p = Object.fromEntries(new Intl.DateTimeFormat('en-GB', { timeZone: tz, year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hour12: false, weekday: 'short' }).formatToParts(d).map((x) => [x.type, x.value]));
        if (f === 'yyyy-MM') return p.year + '-' + p.month;
        if (f === 'yyyy-MM-dd') return p.year + '-' + p.month + '-' + p.day;
        if (f === 'd MMM yyyy') return new Intl.DateTimeFormat('en-GB', { timeZone: tz, day: 'numeric', month: 'short', year: 'numeric' }).format(d);
        if (f === 'yyyy-MM-dd HHmm') return p.year + '-' + p.month + '-' + p.day + ' ' + p.hour + p.minute;
        return new Intl.DateTimeFormat('en-GB', { timeZone: tz, weekday: 'short', day: 'numeric', month: 'short', hour: 'numeric', minute: '2-digit', hour12: true }).format(d);
      },
    },
    MailApp: { getRemainingDailyQuota: () => W.quota, sendEmail(m) { W.quota--; W.mails.push(m); } },
    LockService: { getScriptLock: () => ({ waitLock() { }, tryLock: () => true, releaseLock() { } }) },
    PropertiesService: { getScriptProperties: () => ({ getProperty: (k) => (k in W.props ? W.props[k] : null), setProperty(k, v) { W.props[k] = String(v); } }) },
    ScriptApp: {
      getProjectTriggers: () => W.triggers.slice(), deleteTrigger(t) { W.triggers = W.triggers.filter((x) => x !== t); },
      newTrigger: (fn) => ({ timeBased() { return this; }, everyMinutes() { return this; }, create() { W.triggers.push({ getHandlerFunction: () => fn }); } }),
      getService: () => ({ getUrl: () => 'https://script.google.com/macros/s/TEST/exec' }),
    },
    Session: { getEffectiveUser: () => ({ getEmail: () => 'owner@example.com' }) },
    HtmlService: { createHtmlOutput: (html) => ({ html, title: '', setTitle(t) { this.title = t; return this; }, addMetaTag() { return this; } }) },
  };
  vm.createContext(ctx);
  vm.runInContext(fs.readFileSync(path.join(__dirname, '..', 'dist', 'Code.gs'), 'utf8'), ctx, { filename: 'Code.gs' });
  W.ctx = ctx;
  W.run = (fn, ...args) => vm.runInContext(fn, ctx)(...JSON.parse(JSON.stringify(args)));
  W.advance = (minutes) => { W.clock += minutes * 60000; };
  W.table = (name) => { const sh = sheets[name]; const head = sh.rows[0]; return sh.rows.slice(1).filter((r) => r && r[0] !== '' && r[0] != null).map((r) => Object.fromEntries(head.map((k, i) => [k, r[i] === undefined ? '' : r[i]]))); };
  W.setting = (k, v) => { const sh = sheets.Settings; const row = sh.rows.findIndex((r) => r[0] === k); sh.rows[row][1] = v; vm.runInContext('SETTINGS_CACHE = null', ctx); };
  return W;
}
module.exports = { makeWorld };
