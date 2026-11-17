// 탑클래스 단어관리 앱 전용 Apps Script (단어DB 시트에 연결)
const SS = SpreadsheetApp.getActiveSpreadsheet();
const COLS = {
  sets:     ['id','cat','type','name','unitName','rule','created','memo','count'],
  words:    ['id','set','unit','u','en','ko','q','kind','src','added'],
  students: ['id','name','cls','school','plans','memo','active','order'],
  records:  ['id','date','sid','set','from','to','count','method','words','result','note','created'],
  homework: ['date','sid','text','updated'],
  settings: ['key','value']
};

function doGet(e)  { return handle(e); }
function doPost(e) { return handle(e); }

function handle(e) {
  let body;
  try {
    body = e.postData ? JSON.parse(e.postData.contents) : e.parameter;
  } catch (err) { return respond({ error: 'parse error: ' + err.message }); }
  const action = body.action;
  try {
    let r;
    if      (action === 'readAll')       r = readAll();
    else if (action === 'readWords')     r = readWords();
    else if (action === 'readRecord')    r = readTable('records').find(x => x.id === String(body.id)) || { error: 'not found' };
    else if (action === 'saveStudent')   r = locked(() => upsert('students', body.row, ['id']));
    else if (action === 'deleteStudent') r = locked(() => removeWhere('students', ['id'], [body.id]));
    else if (action === 'saveRecords')   r = locked(() => { (body.rows || []).forEach(x => upsert('records', x, ['id'])); return { ok: true }; });
    else if (action === 'deleteRecord')  r = locked(() => removeWhere('records', ['id'], [body.id]));
    else if (action === 'saveHomework')  r = locked(() => upsert('homework', body.row, ['date','sid']));
    else if (action === 'saveSetting')   r = locked(() => upsert('settings', { key: body.key, value: body.value }, ['key']));
    else if (action === 'updateWord')    r = locked(() => updateWord(body));
    else if (action === 'importSet')     r = locked(() => importSet(body));
    else r = { error: 'unknown action' };
    return respond(r);
  } catch (err) { return respond({ error: err.message }); }
}

function respond(d) {
  return ContentService.createTextOutput(JSON.stringify(d)).setMimeType(ContentService.MimeType.JSON);
}

function locked(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function sh(name) {
  let s = SS.getSheetByName(name);
  if (!s) { s = SS.insertSheet(name); s.getRange(1, 1, 1, COLS[name].length).setValues([COLS[name]]); }
  return s;
}

function cellStr(v) {
  if (v instanceof Date) {
    return v.getFullYear() + '-' + String(v.getMonth() + 1).padStart(2, '0') + '-' + String(v.getDate()).padStart(2, '0');
  }
  return v === null || v === undefined ? '' : String(v);
}

function readTable(name) {
  const s = sh(name);
  const last = s.getLastRow();
  if (last < 2) return [];
  const cols = COLS[name];
  const vals = s.getRange(2, 1, last - 1, cols.length).getValues();
  return vals.filter(r => r[0] !== '' && r[0] !== null).map(r => {
    const o = {};
    cols.forEach((c, i) => o[c] = cellStr(r[i]));
    return o;
  });
}

function readAll() {
  const settings = {};
  readTable('settings').forEach(r => settings[r.key] = r.value);
  return {
    sets: readTable('sets'),
    students: readTable('students'),
    records: recentRecords(),
    homework: readTable('homework'),
    settings: settings
  };
}

// 45일 지난 기록은 단어 목록을 빼고 보냄 (다시 출력할 때 readRecord로 따로 받음)
function recentRecords() {
  const d = new Date(); d.setDate(d.getDate() - 45);
  const cut = cellStr(d);
  return readTable('records').map(r => { if (r.date < cut) r.words = ''; return r; });
}

function readWords() {
  const s = sh('words');
  const last = s.getLastRow();
  const ver = getSetting('wordsVer') || '1';
  if (last < 2) return { ver: ver, rows: [] };
  const vals = s.getRange(2, 1, last - 1, COLS.words.length).getValues();
  return { ver: ver, rows: vals.filter(r => r[0] !== '').map(r => r.map(cellStr)) };
}

function getSetting(key) {
  const r = readTable('settings').find(x => x.key === key);
  return r ? r.value : '';
}

function bumpWordsVer() {
  const v = String(Number(getSetting('wordsVer') || '1') + 1);
  upsert('settings', { key: 'wordsVer', value: v }, ['key']);
  return v;
}

function upsert(name, obj, keys) {
  const s = sh(name);
  const cols = COLS[name];
  const row = cols.map(c => obj[c] === undefined || obj[c] === null ? '' : String(obj[c]));
  const last = s.getLastRow();
  let idx = -1;
  if (last >= 2) {
    const vals = s.getRange(2, 1, last - 1, cols.length).getValues();
    const ki = keys.map(k => cols.indexOf(k));
    idx = vals.findIndex(r => ki.every((i, j) => cellStr(r[i]) === String(obj[keys[j]])));
  }
  const target = idx >= 0 ? s.getRange(idx + 2, 1, 1, cols.length) : s.getRange(last + 1, 1, 1, cols.length);
  target.setNumberFormat('@').setValues([row]);
  return { ok: true };
}

function removeWhere(name, keys, values) {
  const s = sh(name);
  const cols = COLS[name];
  const last = s.getLastRow();
  if (last < 2) return { ok: true };
  const vals = s.getRange(2, 1, last - 1, cols.length).getValues();
  const ki = keys.map(k => cols.indexOf(k));
  for (let i = vals.length - 1; i >= 0; i--) {
    if (ki.every((c, j) => cellStr(vals[i][c]) === String(values[j]))) s.deleteRow(i + 2);
  }
  return { ok: true };
}

function updateWord(body) {
  const s = sh('words');
  const f = s.getRange('A:A').createTextFinder(String(body.id)).matchEntireCell(true).findNext();
  if (!f) return { error: 'word not found' };
  const r = f.getRow();
  if (body.en !== undefined) s.getRange(r, 5).setNumberFormat('@').setValue(String(body.en));
  if (body.ko !== undefined) s.getRange(r, 6).setNumberFormat('@').setValue(String(body.ko));
  s.getRange(r, 9).setNumberFormat('@').setValue('수정');
  return { ok: true, ver: bumpWordsVer() };
}

// body.set = {id,cat,type,name,unitName,rule,memo}, body.words = [[unit,u,en,ko,q,kind,src,added], ...]
function importSet(body) {
  const set = body.set;
  if (!set || !set.id) return { error: 'no set' };
  const s = sh('words');
  const last = s.getLastRow();
  let maxN = 0, count = 0;
  if (last >= 2) {
    s.getRange(2, 1, last - 1, 2).getValues().forEach(r => {
      if (String(r[1]) === set.id) {
        count++;
        const n = Number(String(r[0]).split('-').pop());
        if (n > maxN) maxN = n;
      }
    });
  }
  const rows = (body.words || []).map(w => {
    maxN++;
    return [set.id + '-' + String(maxN).padStart(5, '0'), set.id, w[0], Number(w[1]), w[2], w[3], w[4] || '', w[5] || '', w[6] || '', w[7] || ''];
  });
  if (rows.length) {
    const rng = s.getRange(last + 1, 1, rows.length, COLS.words.length);
    rng.setNumberFormat('@');
    s.getRange(last + 1, 4, rows.length, 1).setNumberFormat('0');
    rng.setValues(rows);
  }
  const old = readTable('sets').find(x => x.id === set.id) || {};
  const merged = Object.assign({}, old, set, { count: count + rows.length, created: old.created || set.created || '' });
  upsert('sets', merged, ['id']);
  return { ok: true, added: rows.length, ver: bumpWordsVer() };
}
