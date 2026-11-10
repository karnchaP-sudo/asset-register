/**
 * ทะเบียนสินทรัพย์ (Fixed Asset Register) – Google Sheets + Apps Script
 *
 * ชีตที่ใช้ (setup() สร้างให้):
 *   Assets      ทะเบียนสินทรัพย์ 1 แถวต่อ 1 รายการ
 *   Queue       รายการรอขึ้นทะเบียน / ไม่ขึ้นทะเบียน
 *   Categories  หมวดสินทรัพย์และรหัสบัญชี
 *   Settings    ค่าตั้งระบบ (key / value เป็น JSON)
 *   Log         บันทึกการแก้ไข
 * รูปถ่ายเก็บในโฟลเดอร์ Drive "รูปทรัพย์สิน" ข้างไฟล์ชีต
 */

const PRINT_PAGE_URL = 'https://karnchap-sudo.github.io/asset-register/print.html';

const ASSET_COLS = ['id','code','cat','name','sn','qty','unit','branch','custodian','vendor','invoice','acqDate','cost','salvage','life','status',
  'docNo','seqFrom','seqTo','refs','po','legacy','note','photos','json','updatedAt','updatedBy'];
const QUEUE_COLS = ['pid','status','cat','name','qty','unit','branch','vendor','invoice','acqDate','cost','refs','po','acc','accName','reason','note',
  'rejReason','rejDate','json','updatedAt','updatedBy'];
const CAT_COLS = ['code','acc','name','life','dep','supply'];
const NUM = {qty:1, cost:1, salvage:1, life:1, seqFrom:1, seqTo:1};
const DATE = {acqDate:1, rejDate:1};
const DEFAULT_CATS = [
  ['BD','1702100','ส่วนปรับปรุงอาคาร',10,true,false],
  ['OF','1704001','อุปกรณ์เครื่องใช้สำนักงาน',5,true,false],
  ['FF','1703001','เครื่องตกแต่งติดตั้ง',5,true,false],
  ['KT','1704000','เครื่องมือเครื่องใช้ในครัว',5,true,false],
  ['VH','1706000','ยานพาหนะ',5,true,false],
  ['SW','1801000','ซอฟต์แวร์และแอปพลิเคชั่น',3,true,false],
  ['CP','1707000','งานระหว่างก่อสร้าง',0,false,false],
  ['CS','1306000','วัสดุสิ้นเปลือง',0,false,true]
];

/* ---------------- web app ---------------- */
function doGet(){
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ทะเบียนสินทรัพย์')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

function onOpen(){
  SpreadsheetApp.getUi().createMenu('ทะเบียนสินทรัพย์')
    .addItem('ตั้งค่าชีตครั้งแรก', 'setup')
    .addToUi();
}

/* ---------------- setup ---------------- */
function setup(){
  const ss = SpreadsheetApp.getActive();
  sheet_('Assets', ASSET_COLS); migrate_('Assets', ASSET_COLS);
  sheet_('Queue', QUEUE_COLS);
  const cs = sheet_('Categories', CAT_COLS);
  if (cs.getLastRow() < 2) cs.getRange(2, 1, DEFAULT_CATS.length, CAT_COLS.length).setValues(DEFAULT_CATS);
  else {
    const have = cs.getRange(2, 1, cs.getLastRow() - 1, 1).getValues().map(r => String(r[0]));
    DEFAULT_CATS.filter(c => have.indexOf(c[0]) < 0).forEach(c => cs.appendRow(c));
  }
  const st = sheet_('Settings', ['key','value']);
  const defaults = {threshold:5000, branches:[], closed:[], skipped:[], label:{size:'50x30', density:3, type:1, task:'auto'}};
  const keys = st.getLastRow() > 1 ? st.getRange(2, 1, st.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  Object.keys(defaults).filter(k => keys.indexOf(k) < 0).forEach(k => st.appendRow([k, JSON.stringify(defaults[k])]));
  sheet_('Log', ['time','user','action','detail']);
  photoFolder_();
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (def && def.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(def);
  return 'ok';
}
function sheet_(name, cols){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getMaxColumns() < cols.length) sh.insertColumnsAfter(sh.getMaxColumns(), cols.length - sh.getMaxColumns());
  if (sh.getLastRow() === 0){
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).setNumberFormat('@');   // keep codes and dates as plain text
  }
  return sh;
}
/** Bring an existing sheet up to date when columns were added (e.g. 'sn'): rows are moved by header name. */
function migrate_(name, cols){
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() === 0) return;
  const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  if (head.length === cols.length && head.every((h, i) => h === cols[i])) return;
  const rows = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, head.length).getValues() : [];
  const moved = rows.map(r => cols.map(c => { const i = head.indexOf(c); return i < 0 ? '' : r[i]; }));
  if (sh.getMaxColumns() < cols.length) sh.insertColumnsAfter(sh.getMaxColumns(), cols.length - sh.getMaxColumns());
  sh.getRange(1, 1, sh.getMaxRows(), Math.max(cols.length, head.length)).clearContent();
  sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat('@');
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold');
  if (moved.length) sh.getRange(2, 1, moved.length, cols.length).setValues(moved);
}
function photoFolder_(){
  const props = PropertiesService.getDocumentProperties();
  const id = props.getProperty('photoFolder');
  if (id){ try { return DriveApp.getFolderById(id); } catch(e){} }
  const file = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  const parents = file.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const f = parent.createFolder('รูปทรัพย์สิน');
  props.setProperty('photoFolder', f.getId());
  return f;
}

/* ---------------- read ---------------- */
function rows_(name){
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return {sh, head: sh ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : [], rows: []};
  const v = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
  return {sh, head: v[0].map(String), rows: v.slice(1)};
}
function iso_(v){
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '');
}
function toObj_(head, r){
  let o = {};
  const ji = head.indexOf('json');
  if (ji >= 0 && r[ji]) { try { o = JSON.parse(r[ji]); } catch(e){} }
  head.forEach((h, i) => {
    if (!h || h === 'json' || h === 'updatedAt' || h === 'updatedBy') return;
    let v = r[i];
    if (h === 'refs') v = String(v || '').split(',').map(s => s.trim()).filter(String);
    else if (h === 'photos') { try { v = v ? JSON.parse(v) : []; } catch(e){ v = []; } }
    else if (NUM[h]) v = v === '' || v === null ? (h === 'qty' ? 1 : 0) : Number(v);
    else if (DATE[h]) v = iso_(v);
    else v = String(v === null || v === undefined ? '' : v);
    o[h] = v;
  });
  return o;
}
function settings_(){
  const t = rows_('Settings'), out = {};
  t.rows.forEach(r => { try { out[String(r[0])] = JSON.parse(r[1]); } catch(e){ out[String(r[0])] = r[1]; } });
  return out;
}
function getState(){
  const a = rows_('Assets'), q = rows_('Queue'), c = rows_('Categories');
  const assets = a.rows.filter(r => r[0]).map(r => toObj_(a.head, r));
  assets.forEach(x => { x.history = x.history || []; x.photos = x.photos || []; });
  const queue = q.rows.filter(r => r[0]).map(r => toObj_(q.head, r));
  const cats = c.rows.filter(r => r[0]).map(r => ({code:String(r[0]), acc:String(r[1]), name:String(r[2]), life:Number(r[3]) || 0,
    dep: r[4] === true || String(r[4]).toUpperCase() === 'TRUE', supply: r[5] === true || String(r[5]).toUpperCase() === 'TRUE'}));
  const s = settings_();
  const branches = (s.branches || []).slice();
  assets.forEach(x => { if (x.branch && branches.indexOf(x.branch) < 0) branches.push(x.branch); });
  return {
    cats, branches, threshold: Number(s.threshold) || 5000, closed: s.closed || [], skipped: s.skipped || [], label: s.label || {},
    assets,
    pending: queue.filter(x => x.status !== 'rejected'),
    rejected: queue.filter(x => x.status === 'rejected'),
    user: Session.getActiveUser().getEmail(),
    printUrl: PRINT_PAGE_URL
  };
}
function getPhotos(ids){
  const out = {};
  (ids || []).slice(0, 40).forEach(id => {
    try { const b = DriveApp.getFileById(id).getBlob(); out[id] = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes()); }
    catch(e){ out[id] = null; }
  });
  return out;
}

/* ---------------- write ---------------- */
function pad_(n){ return ('000' + n).slice(-4); }
function nextCode_(existing, cat, date){
  const prefix = cat + date.slice(2, 4) + date.slice(5, 7) + '/';
  let max = 0;
  existing.forEach(c => { if (c.indexOf(prefix) === 0) max = Math.max(max, parseInt(c.slice(prefix.length, prefix.length + 4), 10) || 0); });
  return prefix + pad_(max + 1);
}
/**
 * ch = {assets:[obj], deleteAssets:[id], queue:[obj], deleteQueue:[pid], settings:{key:value}, newIds:[id], log:[text]}
 * New assets get their code checked against the sheet, so two people saving at once never share a code.
 */
function saveChanges(ch){
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const user = Session.getActiveUser().getEmail() || '';
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const result = {codes:{}, photos:{}};
    migrate_('Assets', ASSET_COLS);
    const a = rows_('Assets');
    const idx = {}; a.rows.forEach((r, i) => { if (r[0]) idx[String(r[0])] = i; });
    const codes = a.rows.map(r => String(r[1]));
    const newIds = {}; (ch.newIds || []).forEach(id => newIds[id] = 1);
    const folder = (ch.assets || []).some(x => (x.photos || []).some(p => p.src && !p.driveId)) ? photoFolder_() : null;

    (ch.assets || []).forEach(x => {
      // photos: upload new ones to Drive, keep only ids in the sheet
      (x.photos || []).forEach(p => {
        if (p.src && !p.driveId && folder){
          const m = String(p.src).match(/^data:([^;]+);base64,(.*)$/);
          if (m){ const f = folder.createFile(Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], (x.code || x.id) + '_' + p.id + '.jpg')); p.driveId = f.getId(); result.photos[p.id] = p.driveId; }
        }
        delete p.src;
      });
      // code: re-check for new assets
      if (newIds[x.id] && !(x.id in idx)){
        if (x.docNo){
          let m = 0; a.rows.forEach(r => { if (String(r[ASSET_COLS.indexOf('docNo')]) === x.docNo) m = Math.max(m, Number(r[ASSET_COLS.indexOf('seqTo')]) || 0); });
          const n = Math.max(1, (Number(x.seqTo) || 1) - (Number(x.seqFrom) || 1) + 1);
          if (m >= (Number(x.seqFrom) || 1)){ x.seqFrom = m + 1; x.seqTo = m + n; x.code = x.docNo + '/' + pad_(x.seqFrom); result.codes[x.id] = {code:x.code, seqFrom:x.seqFrom, seqTo:x.seqTo}; }
        } else if (codes.indexOf(x.code) >= 0){
          x.code = nextCode_(codes, x.cat, x.acqDate); result.codes[x.id] = {code:x.code};
        }
      }
      codes.push(x.code);
      const row = ASSET_COLS.map(h => {
        if (h === 'json') return JSON.stringify(x);
        if (h === 'updatedAt') return now;
        if (h === 'updatedBy') return user;
        if (h === 'refs') return (x.refs || []).join(', ');
        if (h === 'photos') return JSON.stringify(x.photos || []);
        const v = x[h]; return v === undefined || v === null ? '' : String(v);
      });
      if (x.id in idx) a.rows[idx[x.id]] = row; else { idx[x.id] = a.rows.length; a.rows.push(row); }
    });
    const delA = {}; (ch.deleteAssets || []).forEach(id => delA[id] = 1);
    writeAll_(a.sh, ASSET_COLS, a.rows.filter(r => !delA[String(r[0])]));

    if ((ch.queue || []).length || (ch.deleteQueue || []).length){
      const q = rows_('Queue');
      const qi = {}; q.rows.forEach((r, i) => { if (r[0]) qi[String(r[0])] = i; });
      (ch.queue || []).forEach(x => {
        const row = QUEUE_COLS.map(h => {
          if (h === 'json') return JSON.stringify(x);
          if (h === 'updatedAt') return now;
          if (h === 'updatedBy') return user;
          if (h === 'refs') return (x.refs || []).join(', ');
          const v = x[h]; return v === undefined || v === null ? '' : String(v);
        });
        if (x.pid in qi) q.rows[qi[x.pid]] = row; else { qi[x.pid] = q.rows.length; q.rows.push(row); }
      });
      const delQ = {}; (ch.deleteQueue || []).forEach(id => delQ[id] = 1);
      writeAll_(q.sh, QUEUE_COLS, q.rows.filter(r => !delQ[String(r[0])]));
    }

    if (ch.settings){
      const st = rows_('Settings');
      Object.keys(ch.settings).forEach(k => {
        const i = st.rows.findIndex(r => String(r[0]) === k);
        const val = JSON.stringify(ch.settings[k]);
        if (i >= 0) st.sh.getRange(i + 2, 2).setValue(val); else st.sh.appendRow([k, val]);
      });
      if (ch.cats){
        const cs = SpreadsheetApp.getActive().getSheetByName('Categories');
        const vals = ch.cats.map(c => [c.code, c.acc, c.name, c.life, c.dep !== false, !!c.supply]);
        cs.getRange(2, 1, Math.max(cs.getLastRow() - 1, 1), CAT_COLS.length).clearContent();
        cs.getRange(2, 1, vals.length, CAT_COLS.length).setValues(vals);
      }
    }
    if ((ch.log || []).length){
      const lg = SpreadsheetApp.getActive().getSheetByName('Log');
      lg.getRange(lg.getLastRow() + 1, 1, ch.log.length, 4).setValues(ch.log.map(t => [now, user, 'save', String(t).slice(0, 500)]));
    }
    SpreadsheetApp.flush();
    return result;
  } finally {
    lock.releaseLock();
  }
}
function writeAll_(sh, cols, rows){
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, cols.length).clearContent();
  if (rows.length) sh.getRange(2, 1, rows.length, cols.length).setNumberFormat('@').setValues(rows.map(r => cols.map((_, i) => r[i] === undefined || r[i] === null ? '' : r[i])));
}
