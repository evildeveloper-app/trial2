// ═══════════════════════════════════════════
//  GOOGLE APPS SCRIPT BACKEND CODE REFERENCE
// ═══════════════════════════════════════════
const SCRIPT_CODE = `
function doGet(e) {
  try {
    // Write operations are sent as GET with body= param (avoids POST redirect issues)
    if (e.parameter.body) {
      var b = JSON.parse(e.parameter.body);
      if (!_checkAuth(b.k, b.action)) return R({error: 'unauthorized'});
      var fn = ACTIONS[b.action];
      if (!fn) return R({error: 'Unknown action: ' + b.action});
      return R(fn(b));
    }
    var action = e.parameter.action || 'getAll';
    if (!_checkAuth(e.parameter.k, action)) return R({error: 'unauthorized'});
    if (action === 'getAll') return R(getAllData());
    return R({error: 'Unknown action'});
  } catch(err) { return R({error: err.message}); }
}
function doPost(e) {
  try {
    var b = JSON.parse(e.postData.contents);
    if (!_checkAuth(b.k, b.action)) return R({error: 'unauthorized'});
    var fn = ACTIONS[b.action];
    if (!fn) return R({error: 'Unknown action'});
    return R(fn(b));
  } catch(err) { return R({error: err.message}); }
}
function R(d) { return ContentService.createTextOutput(JSON.stringify(d)).setMimeType(ContentService.MimeType.JSON); }
// ── Optional shared-password gate ──
// If Settings has an 'app_secret' value, EVERY request must carry the matching token
// (?k=...). If no secret is set, the endpoint stays open (backward compatible), so
// existing setups keep working until the owner turns a password on.
function _secretVal(){ var st = sheet('Settings',['Key','Value']); var r = findRow(st,'app_secret'); return r===-1 ? '' : String(st.getRange(r,2).getValue()||''); }
function _checkAuth(k, action){ var s = _secretVal(); if(!s) return true; if(action==='setSecret') return true; return String(k||'') === s; }
function sheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); var s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold'); s.setFrozenRows(1); }
  return s;
}
function rows(s) { return s.getLastRow() < 2 ? [] : s.getRange(2,1,s.getLastRow()-1,s.getLastColumn()).getValues(); }
var TX_HEADERS = ['ID','Date','Description','Amount','Type','Category','Account','Notes','MemberID','CreatedAt'];
var TX_ROLL_LIMIT = 10000; // rows per Transactions tab before rolling to a fresh tab
function txSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = [sheet('Transactions', TX_HEADERS)];
  ss.getSheets().forEach(function(sh){ if (sh.getName().indexOf('Transactions Archive') === 0) out.push(sh); });
  return out;
}
function txFind(id) {
  var list = txSheets();
  for (var i = 0; i < list.length; i++) { var r = findRow(list[i], id); if (r !== -1) return { s: list[i], row: r }; }
  return null;
}
function rolloverIfFull(tx) {
  // When the active Transactions tab passes the limit, rename it to an archive tab
  // and start a fresh Transactions tab. Reads merge every tab, so nothing is lost
  // and the app never notices.
  if (tx.getLastRow() <= TX_ROLL_LIMIT + 1) return;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var n = 1; while (ss.getSheetByName('Transactions Archive ' + n)) n++;
  tx.setName('Transactions Archive ' + n);
  var fresh = ss.insertSheet('Transactions', 0);
  fresh.getRange(1,1,1,TX_HEADERS.length).setValues([TX_HEADERS]).setFontWeight('bold');
  fresh.setFrozenRows(1);
}
function findRow(s, id) {
  var data = rows(s);
  for (var i=0;i<data.length;i++) { if (String(data[i][0]) === String(id)) return i+2; }
  return -1;
}
function fmtDate(d) {
  if (Object.prototype.toString.call(d) === '[object Date]') {
    // Format in the SPREADSHEET's own timezone, not UTC. Sheets stores a typed
    // date at local midnight; formatting it in UTC shifts it to the previous day
    // for users ahead of UTC (e.g. Australia). Using the sheet timezone keeps the
    // date the user entered exactly intact on every read.
    var tz;
    try { tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(); } catch (e) { tz = 'UTC'; }
    return Utilities.formatDate(d, tz || 'UTC', 'yyyy-MM-dd');
  }
  return d;
}
function adjustAccountBalance(accountName, amount, type, sign) {
  if (!accountName) return;
  var ac = sheet('Accounts',['ID','Name','Type','Balance']);
  var data = rows(ac);
  for (var i=0;i<data.length;i++) {
    if (data[i][1] === accountName) {
      var delta = (type === 'expense' ? -Number(amount) : Number(amount)) * sign;
      ac.getRange(i+2,4).setValue(Number(data[i][3]||0) + delta);
      return;
    }
  }
}
function getAllData() {
  var txList = txSheets();
  var ac = sheet('Accounts',['ID','Name','Type','Balance']);
  var mb = sheet('Members',['ID','Name','PIN','Color']);
  var rc = sheet('Recurring',['ID','Name','Amount','Type','Category','Account','Frequency','NextDate','MemberID','Active']);
  var bg = sheet('Budgets',['ID','Category','Limit']);
  var st = sheet('Settings',['Key','Value']);
  var settings = {}; rows(st).forEach(function(r){ if(r[0] && r[0]!=='app_secret') settings[r[0]]=r[1]; });
  return {
    transactions: txList.reduce(function(acc, sh){ return acc.concat(rows(sh)); }, []).filter(function(r){return r[0];}).map(function(r){ return {id:r[0],date:fmtDate(r[1]),description:r[2],amount:r[3],type:r[4],category:r[5],account:r[6],notes:r[7],memberId:r[8]||'',ts:r[9]||''}; }),
    accounts: rows(ac).filter(function(r){return r[0];}).map(function(r){ return {id:r[0],name:r[1],type:r[2],balance:r[3]}; }),
    members: rows(mb).filter(function(r){return r[0];}).map(function(r){ return {id:r[0],name:r[1],pin:String(r[2]),color:r[3]}; }),
    recurring: rows(rc).filter(function(r){return r[0];}).map(function(r){ return {id:r[0],name:r[1],amount:r[2],type:r[3],category:r[4],account:r[5],frequency:r[6],nextDate:fmtDate(r[7]),memberId:r[8]||'',active:r[9]}; }),
    budgets: rows(bg).filter(function(r){return r[0];}).map(function(r){ return {id:r[0],category:r[1],limit:r[2]}; }),
    settings: settings
  };
}
var ACTIONS = {
  addTransaction: function(b) {
    // Idempotent: if this id already exists (e.g. an offline write that was actually
    // received but whose reply got lost, then retried), don't append a duplicate.
    if (txFind(b.id)) return {success:true, id:b.id, dup:true};
    var tx = sheet('Transactions', TX_HEADERS);
    tx.appendRow([b.id, b.date, b.description, b.amount, b.type, b.category, b.account, b.notes||'', b.memberId||'', b.ts || Date.now()]);
    adjustAccountBalance(b.account, b.amount, b.type, 1);
    rolloverIfFull(tx);
    return {success:true, id:b.id};
  },
  updateTransaction: function(b) {
    var f = txFind(b.id);
    if (!f) return {error:'Not found'};
    var old = f.s.getRange(f.row,1,1,10).getValues()[0];
    // Conflict resolution for multiple people (up to 6+): each edit carries a timestamp
    // (col 10). If the incoming edit is OLDER than what's already stored, a slower/stale
    // device is trying to overwrite a newer edit — ignore it. Last edit wins, deterministically.
    var incomingTs = Number(b.ts||0), storedTs = Number(old[9]||0);
    if (incomingTs && storedTs && incomingTs < storedTs) return {success:true, stale:true};
    adjustAccountBalance(old[6], old[3], old[4], -1);
    // Bump the hidden timestamp (col 10) on every edit so the app always knows the
    // most-recently-touched record, independent of the calendar date.
    f.s.getRange(f.row,1,1,10).setValues([[b.id, b.date, b.description, b.amount, b.type, b.category, b.account, b.notes||'', b.memberId || old[8] || '', b.ts || Date.now()]]);
    adjustAccountBalance(b.account, b.amount, b.type, 1);
    return {success:true};
  },
  deleteTx: function(b) {
    var f = txFind(b.id);
    if (!f) return {error:'Not found'};
    var old = f.s.getRange(f.row,1,1,10).getValues()[0];
    adjustAccountBalance(old[6], old[3], old[4], -1);
    f.s.deleteRow(f.row);
    return {success:true};
  },
  addAccount: function(b) {
    var ac = sheet('Accounts',['ID','Name','Type','Balance']);
    ac.appendRow([b.id, b.name, b.type, Number(b.balance)||0]);
    return {success:true, id:b.id};
  },
  updateAccount: function(b) {
    var ac = sheet('Accounts',['ID','Name','Type','Balance']);
    var r = findRow(ac, b.id);
    if (r===-1) return {error:'Not found'};
    ac.getRange(r,1,1,4).setValues([[b.id, b.name, b.type, Number(b.balance)||0]]);
    return {success:true};
  },
  deleteAccount: function(b) {
    var ac = sheet('Accounts',['ID','Name','Type','Balance']);
    var r = findRow(ac, b.id);
    if (r!==-1) ac.deleteRow(r);
    return {success:true};
  },
  addMember: function(b) {
    var mb = sheet('Members',['ID','Name','PIN','Color']);
    mb.appendRow([b.id, b.name, b.pin, b.color]);
    return {success:true, id:b.id};
  },
  deleteMember: function(b) {
    var mb = sheet('Members',['ID','Name','PIN','Color']);
    var r = findRow(mb, b.id);
    if (r!==-1) mb.deleteRow(r);
    return {success:true};
  },
  addRecurring: function(b) {
    var rc = sheet('Recurring',['ID','Name','Amount','Type','Category','Account','Frequency','NextDate','MemberID','Active']);
    rc.appendRow([b.id, b.name, b.amount, b.type, b.category, b.account||'', b.frequency, b.nextDate, b.memberId||'', true]);
    return {success:true, id:b.id};
  },
  updateRecurring: function(b) {
    var rc = sheet('Recurring',['ID','Name','Amount','Type','Category','Account','Frequency','NextDate','MemberID','Active']);
    var r = findRow(rc, b.id);
    if (r===-1) return {error:'Not found'};
    var old = rc.getRange(r,1,1,10).getValues()[0];
    rc.getRange(r,1,1,10).setValues([[b.id, b.name, b.amount, b.type, b.category, b.account||'', b.frequency, b.nextDate, b.memberId || old[8] || '', b.active!==undefined?b.active:old[9]]]);
    return {success:true};
  },
  deleteRecurring: function(b) {
    var rc = sheet('Recurring',['ID','Name','Amount','Type','Category','Account','Frequency','NextDate','MemberID','Active']);
    var r = findRow(rc, b.id);
    if (r!==-1) rc.deleteRow(r);
    return {success:true};
  },
  setBudget: function(b) {
    var bg = sheet('Budgets',['ID','Category','Limit']);
    var r = findRow(bg, b.id);
    if (r===-1) bg.appendRow([b.id, b.category, Number(b.limit)||0]);
    else bg.getRange(r,1,1,3).setValues([[b.id, b.category, Number(b.limit)||0]]);
    return {success:true, id:b.id};
  },
  deleteBudget: function(b) {
    var bg = sheet('Budgets',['ID','Category','Limit']);
    var r = findRow(bg, b.id);
    if (r!==-1) bg.deleteRow(r);
    return {success:true};
  },
  updateSettings: function(b) {
    var st = sheet('Settings',['Key','Value']);
    var r = findRow(st, b.key);
    if (r===-1) st.appendRow([b.key, b.value]);
    else st.getRange(r,2).setValue(b.value);
    return {success:true};
  },
  // Set or change the shared password. Allowed with no proof when none is set yet;
  // to CHANGE an existing one, the caller must send the current password in b.current.
  setSecret: function(b) {
    var st = sheet('Settings',['Key','Value']);
    var r = findRow(st, 'app_secret');
    var cur = r===-1 ? '' : String(st.getRange(r,2).getValue()||'');
    if (cur && String(b.current||'') !== cur) return {error:'wrong password'};
    var val = String(b.secret||'');
    if (r===-1) st.appendRow(['app_secret', val]);
    else st.getRange(r,2).setValue(val);
    return {success:true};
  }
};
`;

// ═══════════════════════════════════════════
//  GLOBAL APPLICATION STATE
// ═══════════════════════════════════════════
// ── Versioning: MAJOR.MINOR · majors are big releases, +0.5 feature drops, +0.1 fixes
const APP_VERSION = '3.4';

const S = {
  txs: [], accts: [], incs: [], members: [], recurring: [], budgets: [],
  type: 'expense', recType: 'expense', cat: 'Food', recCat: 'Bills', filter: 'all', memberFilter: 'all', period: 'week', heroPeriod: 'month',
  cur: {code:'AUD',sym:'A$'}, name: 'Lucid User', mode: 'local', family: false, url: '',
  familyCode: '', familyName: '',
  member: null,
  editingTxId: null, editingAcctId: null, editingRecId: null, editingBudgetId: null,
  addDraft: null, _skipDraft: false,
  newMemberPin: '', memberPin: '', pinEntryVal: '', pinTarget: null,
  pollHandle: null, fbUid: null, fbFamilyCode: null, pendingJoinCode: false, simpleMode: false,
  barChart: null, donutChart: null, lineChart: null, lineChart2: null, dowChart: null, catPieChart: null, accent: '#6C5CE7',
  analyticsTab: 'overview', breakdownView: 'category',
  theme: 'system', onboardAccts: [], onboardingNext: 'enterApp',
  setupContext: 'solo', trackIncome: true, dueQueue: [], memberId: null
};

const CATS = [
  {n:'Food',e:'🍔'},{n:'Transport',e:'🚗'},{n:'Shopping',e:'🛍'},{n:'Health',e:'💊'},
  {n:'Entertainment',e:'🎬'},{n:'Bills',e:'📋'},{n:'Education',e:'📚'},{n:'Travel',e:'✈️'},
  {n:'Coffee',e:'☕'},{n:'Fitness',e:'💪'},{n:'Home',e:'🏠'},{n:'Other',e:'💡'}
];

const INC_CATS = [
  {n:'Salary',e:'💼'},{n:'Freelance',e:'💻'},{n:'Business',e:'🏢'},
  {n:'Investment',e:'📈'},{n:'Rental',e:'🏠'},{n:'Other',e:'✨'}
];
// ── Category list helpers (supports user-managed lists) ──
function getCatList(isIncome) {
  const key = isIncome ? 'prizm_all_inc_cats' : 'prizm_all_cats';
  try {
    const s = localStorage.getItem(key);
    if (s) { const arr = JSON.parse(s); if (Array.isArray(arr) && arr.length) return arr; }
  } catch(e) {}
  return (isIncome ? INC_CATS : CATS).map(c => ({...c}));
}

function _saveCatList(isIncome, list) {
  localStorage.setItem(isIncome ? 'prizm_all_inc_cats' : 'prizm_all_cats', JSON.stringify(list));
  schedulePrefSync(); // #2 · sync added/removed categories across devices
}

function addCat(type) {
  const isInc = type === 'inc';
  const emoji = (document.getElementById('new-'+type+'-emoji').value.trim()) || (isInc ? '💸' : '🏷️');
  const name = document.getElementById('new-'+type+'-name').value.trim();
  if (!name) return;
  const list = getCatList(isInc);
  if (!list.find(c => c.n === name)) { list.push({n: name, e: emoji}); _saveCatList(isInc, list); }
  document.getElementById('new-'+type+'-emoji').value = '';
  document.getElementById('new-'+type+'-name').value = '';
  renderCatManager();
}

function removeCat(type, name) {
  const isInc = type === 'inc';
  _saveCatList(isInc, getCatList(isInc).filter(c => c.n !== name));
  renderCatManager();
}

function updateCatEmoji(type, name, val) {
  const isInc = type === 'inc';
  const list = getCatList(isInc);
  const cat = list.find(c => c.n === name);
  if (cat && val.trim()) { cat.e = val.trim(); _saveCatList(isInc, list); } // _saveCatList syncs
}

const EMOJI_SUGGEST = {
  exp:['🍔','🚗','🛍','💊','🎬','📋','📚','✈️','☕','💪','🏠','🎮','🐶','👶','🎁','💡'],
  inc:['💼','💻','🏢','📈','🏠','🪙','🎁','✨']
};
function renderCatManager() {
  ['exp','inc'].forEach(function(type) {
    const isInc = type === 'inc';
    const el = document.getElementById('adv-'+type+'-cats');
    if (!el) return;
    el.innerHTML = getCatList(isInc).map(function(c) {
      const sn = esc(jsq(c.n));
      return `<div class="cat-row">
        <input class="finput cat-name-in" type="text" value="${esc(c.n)}" onchange="renameCat('${type}','${sn}',this.value)" title="Tap to rename">
        <input class="finput cat-emoji-in" type="text" maxlength="2" value="${esc(c.e)}" oninput="updateCatEmoji('${type}','${sn}',this.value)" title="Icon">
        <button onclick="removeCat('${type}','${sn}')" style="background:none;border:none;cursor:pointer;color:var(--text-3);font-size:20px;padding:4px 8px;flex-shrink:0;line-height:1" title="Remove">&times;</button>
      </div>`;
    }).join('') + `<div class="cat-suggest"><div class="cat-suggest-label">Suggested icons</div>` +
      EMOJI_SUGGEST[type].map(function(e){ return `<span onclick="document.getElementById('new-${type}-emoji').value='${e}'" title="Use this icon">${e}</span>`; }).join('') +
      `</div>`;
  });
}
function renameCat(type, oldName, newName) {
  newName = (newName || '').trim();
  const isInc = type === 'inc';
  if (!newName || newName === oldName) { renderCatManager(); return; }
  const list = getCatList(isInc);
  if (list.find(c => c.n === newName)) { toast('That category already exists'); renderCatManager(); return; }
  const c = list.find(x => x.n === oldName);
  if (!c) { renderCatManager(); return; }
  c.n = newName;
  _saveCatList(isInc, list);
  // Move every existing record over to the new name
  let changed = false;
  (S.txs||[]).forEach(t => { if (t.category === oldName) { t.category = newName; changed = true; } });
  (S.recurring||[]).forEach(r => { if (r.category === oldName) { r.category = newName; changed = true; } });
  (S.budgets||[]).forEach(b => { if (b.category === oldName) { b.category = newName; changed = true; } });
  if (changed) { saveLocalData(); if (S.mode === 'firebase') fbWriteAll().catch(e => console.warn(e)); }
  // Carry any custom emoji/name over to the new key
  [isInc ? 'prizm_inc_emojis' : 'prizm_cat_emojis', isInc ? 'prizm_inc_names' : 'prizm_cat_names'].forEach(k => {
    const m = safeParse(localStorage.getItem(k), {});
    if (m[oldName] !== undefined) { m[newName] = m[oldName]; delete m[oldName]; localStorage.setItem(k, JSON.stringify(m)); }
  });
  if (S.cat === oldName) S.cat = newName;
  if (S.recCat === oldName) S.recCat = newName;
  renderCatManager(); buildCatGrid(); buildRecCatGrid(); buildBudgetCatOptions(); renderAll();
  toast('✓ Renamed · existing entries updated');
}

// ═══ PERSONALIZE PAGE (avatar, name, look & feel) ═══
function _avatarPhoto() { try { return localStorage.getItem('prizm_avatar_photo') || ''; } catch(e) { return ''; } }
function _avatarColor() { try { return localStorage.getItem('prizm_avatar_color') || ''; } catch(e) { return ''; } }
// Render the user's avatar into any tile: photo wins, else color + initial.
function applyAvatarTo(el, initial, colorOverride) {
  if (!el) return;
  const photo = _avatarPhoto();
  if (photo) {
    el.textContent = '';
    el.style.background = 'center/cover no-repeat url("' + photo + '")';
    el.style.color = '';
  } else {
    // The user's explicitly chosen avatar color wins; fall back to the passed-in
    // color (e.g. a family member's assigned color) only when none was picked.
    const c = _avatarColor() || colorOverride;
    el.style.background = c || '';
    el.style.color = c ? contrastText(c) : '';
    el.textContent = (initial || 'P');
  }
}
function openPersonalize() {
  const p = document.getElementById('personalize-page');
  if (!p) return;
  p.classList.add('open');
  renderPersonalize();
}
function closePersonalize() {
  const p = document.getElementById('personalize-page');
  if (p) p.classList.remove('open');
  renderSettings(); // reflect any changes back onto the Settings page
}
function renderPersonalize() {
  const initial = ((S.family && S.member ? S.member.name : S.name) || 'P').charAt(0).toUpperCase();
  applyAvatarTo(document.getElementById('pp-avatar'), initial, S.family && S.member ? S.member.color : '');
  const rm = document.getElementById('pp-remove-photo');
  if (rm) rm.style.display = _avatarPhoto() ? '' : 'none';
  const avc = _avatarColor();
  document.querySelectorAll('#pp-color-row .color-dot').forEach(d => d.classList.toggle('active', d.dataset.avc === avc));
  const nameSub = document.getElementById('pref-name');
  if (nameSub) nameSub.textContent = S.name;
  // refresh look & feel selector states
  const savedTheme = localStorage.getItem('prizm_theme') || 'spectrum';
  document.querySelectorAll('#theme-selector .fc').forEach(el => el.classList.toggle('active', el.dataset.themeMode === savedTheme));
  const savedAccent = localStorage.getItem('prizm_accent') || '#6C5CE7';
  document.querySelectorAll('#accent-selector .color-dot').forEach(el => el.classList.toggle('active', el.id === 'accent-active-dot' || el.dataset.accent === savedAccent));
  updatePrefUI();
}
function pickAvatarPhoto(input) {
  const file = input.files && input.files[0];
  input.value = '';
  if (!file) return;
  const img = new Image();
  const url = URL.createObjectURL(file);
  img.onload = () => {
    try {
      // Center-crop to a 144px square, compress — keeps localStorage usage tiny.
      const SIZE = 144;
      const canvas = document.createElement('canvas');
      canvas.width = SIZE; canvas.height = SIZE;
      const ctx = canvas.getContext('2d');
      const side = Math.min(img.width, img.height);
      ctx.drawImage(img, (img.width-side)/2, (img.height-side)/2, side, side, 0, 0, SIZE, SIZE);
      const data = canvas.toDataURL('image/jpeg', 0.85);
      try { localStorage.setItem('prizm_avatar_photo', data); } catch(e) { toast('⚠️ Storage full — photo not saved'); return; }
      toast('✓ Photo updated');
      renderPersonalize(); renderSettings();
    } catch(e) { toast('Could not read that image'); console.warn(e); }
    finally { URL.revokeObjectURL(url); }
  };
  img.onerror = () => { URL.revokeObjectURL(url); toast('Could not read that image'); };
  img.src = url;
}
function removeAvatarPhoto() {
  try { localStorage.removeItem('prizm_avatar_photo'); } catch(e) {}
  toast('Photo removed');
  renderPersonalize(); renderSettings();
}
function setAvatarColor(c) {
  try {
    if (_avatarColor() === c) localStorage.removeItem('prizm_avatar_color'); // tap again to clear
    else localStorage.setItem('prizm_avatar_color', c);
  } catch(e) {}
  renderPersonalize(); renderSettings();
}

// Categories section is collapsible again — plain in-place expand, nothing nested.
function toggleAdvanced() {
  const panel = document.getElementById('advanced-panel');
  const chev = document.getElementById('adv-chev');
  if (!panel) return;
  const isOpen = panel.style.display !== 'none';
  panel.style.display = isOpen ? 'none' : 'block';
  if (chev) chev.style.transform = isOpen ? '' : 'rotate(180deg)';
  if (!isOpen) renderCatManager();
}


const ACCT_ICONS = { Checking:'🏦',Savings:'💰',Credit:'💳',Cash:'💵',Investment:'📈',Crypto:'🔮',Other:'📂' };
function getAcctIcon(type) { return ACCT_ICONS[type] || '📂'; }

const CAT_COLORS = {
  Food:'rgba(239,68,68,0.12)',Transport:'rgba(96,165,250,0.12)',Shopping:'rgba(245,158,11,0.12)',
  Health:'rgba(34,197,94,0.12)',Entertainment:'rgba(167,139,250,0.12)',Bills:'rgba(148,163,184,0.10)'
};

const CURRENCIES=[
  {code:'USD',sym:'$',name:'US Dollar'},{code:'EUR',sym:'€',name:'Euro'},{code:'GBP',sym:'£',name:'British Pound'},
  {code:'AUD',sym:'A$',name:'Australian Dollar'},{code:'INR',sym:'₹',name:'Indian Rupee'},{code:'CAD',sym:'C$',name:'Canadian Dollar'},
  {code:'JPY',sym:'¥',name:'Japanese Yen'},{code:'CNY',sym:'CN¥',name:'Chinese Yuan'},{code:'SGD',sym:'S$',name:'Singapore Dollar'},
  {code:'NZD',sym:'NZ$',name:'New Zealand Dollar'}
];

const MEMBER_COLORS = ['#F59E0B','#3B82F6','#22C55E','#EC4899','#A78BFA','#06B6D4','#F97316','#EF4444','#84CC16','#6C5CE7'];

// ═══════════════════════════════════════════
//  CORE UTILITIES (escaping, dates)
// ═══════════════════════════════════════════
// #9 · HTML-escape any user/synced free text before injecting into innerHTML.
const _ESC_MAP = {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'};
function esc(s){ return String(s == null ? '' : s).replace(/[&<>"']/g, c => _ESC_MAP[c]); }
// Crash-proof JSON reads · corrupted localStorage falls back instead of throwing
function safeParse(str, fallback){ try { const v = JSON.parse(str); return v == null ? fallback : v; } catch(e) { return fallback; } }
// Escape a string for use inside a single-quoted JS string in an inline handler.
function jsq(s){ return String(s == null ? '' : s).replace(/\\/g,'\\\\').replace(/'/g,"\\'"); }
// Is the UI currently on a dark surface? (spectrum with system-light renders light)
function uiIsDark(){ const t = document.documentElement.getAttribute('data-theme'); if (t === 'light') return false; if (t) return true; return !(window.matchMedia && window.matchMedia('(prefers-color-scheme: light)').matches); }
// Chart palette derived from the active accent so graphs match every theme
function _hex2hsl(hex){ hex=String(hex).replace('#',''); if(hex.length===3) hex=hex.split('').map(c=>c+c).join(''); const r=parseInt(hex.slice(0,2),16)/255,g=parseInt(hex.slice(2,4),16)/255,b=parseInt(hex.slice(4,6),16)/255; const mx=Math.max(r,g,b),mn=Math.min(r,g,b); let h=0,sat=0; const l=(mx+mn)/2; if(mx!==mn){ const d=mx-mn; sat=l>0.5?d/(2-mx-mn):d/(mx+mn); if(mx===r)h=(g-b)/d+(g<b?6:0); else if(mx===g)h=(b-r)/d+2; else h=(r-g)/d+4; h*=60; } return [h,sat*100,l*100]; }
function _hsl2hex(h,sat,l){ sat/=100;l/=100; const k=n=>(n+h/30)%12; const a=sat*Math.min(l,1-l); const f=n=>l-a*Math.max(-1,Math.min(k(n)-3,Math.min(9-k(n),1))); const to=x=>Math.round(255*x).toString(16).padStart(2,'0'); return '#'+to(f(0))+to(f(8))+to(f(4)); }
function themeChartColors(){
  const cs = getComputedStyle(document.documentElement);
  return {
    inc: cs.getPropertyValue('--green').trim() || '#22C55E',
    exp: cs.getPropertyValue('--red').trim() || '#EF4444',
    accent: cs.getPropertyValue('--accent').trim() || '#6C5CE7'
  };
}
function _hexA(hex, a){
  if (!/^#[0-9A-Fa-f]{6}$/.test(hex)) return hex;
  const r = parseInt(hex.slice(1,3),16), g = parseInt(hex.slice(3,5),16), b = parseInt(hex.slice(5,7),16);
  return 'rgba(' + r + ',' + g + ',' + b + ',' + a + ')';
}
function themePalette(n){ let base='#6C5CE7'; try{ base=getComputedStyle(document.documentElement).getPropertyValue('--accent').trim()||base; }catch(e){} const hsl=_hex2hsl(base); const out=[]; for(let i=0;i<n;i++){ out.push(_hsl2hex((hsl[0]+i*137.508)%360, Math.max(45,Math.min(85,hsl[1]||60)), Math.max(48,Math.min(64,hsl[2]||55)))); } return out; }

// #8 · parse a stored 'YYYY-MM-DD' string as a LOCAL date (not UTC midnight),
// so "this month" / period boundaries are consistent with new Date(y,m,1).
function parseDate(s){
  if (s instanceof Date) return s;
  if (!s) return new Date(NaN);
  const str = String(s).slice(0,10);
  const p = str.split('-');
  if (p.length === 3) { const d = new Date(+p[0], +p[1]-1, +p[2]); if (!isNaN(d)) return d; }
  return new Date(s);
}
function dateMs(s){ return parseDate(s).getTime(); }
// Hidden entry time: never shown in the UI, only used so same-day entries
// sort newest-first. Falls back to the timestamp embedded in the tx id.
function txTs(t){ if (t && t.ts) return Number(t.ts) || 0; const m = /^tx_(\d{13})/.exec((t && t.id) || ''); return m ? Number(m[1]) : 0; }
function sortTxs(arr){ arr.sort((a,b) => (dateMs(b.date) - dateMs(a.date)) || (txTs(b) - txTs(a))); return arr; }
function isoFromParts(y,m,d){ return y + '-' + String(m).padStart(2,'0') + '-' + String(d).padStart(2,'0'); }
function isoLocal(dt){ return isoFromParts(dt.getFullYear(), dt.getMonth()+1, dt.getDate()); }
function todayISO(){ return isoLocal(new Date()); }
function fmtDateShort(s){ const d = parseDate(s); return isNaN(d) ? esc(s) : d.toLocaleDateString('en-US',{month:'short',day:'numeric'}); }

// #3 · synthetic bucket for shared / joint family spending (excluded from per-person math).
const JOINT_ID = 'joint';
const JOINT_LABEL = 'Joint / Shared';

// #2 · Customization that must travel across devices (previously localStorage-only).
// Maps a synced "settings" key → the localStorage key that holds its value.
const SYNC_PREF_KEYS = {
  cat_names:'prizm_cat_names', inc_names:'prizm_inc_names',
  cat_emojis:'prizm_cat_emojis', inc_emojis:'prizm_inc_emojis',
  all_cats:'prizm_all_cats', all_inc_cats:'prizm_all_inc_cats',
  simple:'prizm_simple', compact:'prizm_compact', corner:'prizm_corner',
  noemoji:'prizm_noemoji', accent:'prizm_accent', theme:'prizm_theme',
  hero_period:'prizm_hero_period'
};
function gatherPrefs(){
  const o = {};
  for (const k in SYNC_PREF_KEYS){ const v = localStorage.getItem(SYNC_PREF_KEYS[k]); if (v !== null && v !== '') o[k] = v; }
  return o;
}
let _localPrefsDirtyUntil = 0; // suppress incoming-pref clobber right after a local change
function applyIncomingPrefs(settings){
  if (!settings) return;
  // Don't let a background sync overwrite a preference (theme, accent, simple mode…)
  // the user JUST changed on this device — otherwise the change appears to "switch
  // back" when a slightly-stale Sheet read lands right after. Cross-device prefs still
  // sync in, just after this short settle window (once the local change is pushed up).
  if (Date.now() < _localPrefsDirtyUntil) return;
  let changed = false;
  for (const k in SYNC_PREF_KEYS){
    if (!Object.prototype.hasOwnProperty.call(settings, k)) continue;
    const lk = SYNC_PREF_KEYS[k]; const v = settings[k];
    if (v === '' || v == null) { if (localStorage.getItem(lk) !== null) { localStorage.removeItem(lk); changed = true; } }
    else if (localStorage.getItem(lk) !== String(v)) { localStorage.setItem(lk, String(v)); changed = true; }
  }
  if (changed) reapplyPrefsToUI();
}
function reapplyPrefsToUI(){
  try { applyEmojiPref(); } catch(e) {}
  try { setCorner(localStorage.getItem('prizm_corner') || 'round', null); } catch(e) {}
  try { if (localStorage.getItem('prizm_compact')) document.body.setAttribute('data-compact','1'); else document.body.removeAttribute('data-compact'); } catch(e) {}
  try { const ac = localStorage.getItem('prizm_accent'); if (ac) setAccent(ac, false); } catch(e) {}
  try { applyTheme(localStorage.getItem('prizm_theme') || 'system'); } catch(e) {}
  try { S.simpleMode = localStorage.getItem('prizm_simple') === '1'; applySimpleMode(); } catch(e) {}
  try { S.heroPeriod = localStorage.getItem('prizm_hero_period') || 'month'; } catch(e) {}
  try { if (document.getElementById('cat-grid')) { buildCatGrid(); buildRecCatGrid(); buildBudgetCatOptions(); } } catch(e) {}
}
let _prefSyncTimer = null;
function flushPrefs(){
  if (S.mode === 'local') return;
  if (S.mode === 'firebase') { fbWriteAll().catch(e => console.warn(e)); return; }
  const p = gatherPrefs();
  Object.keys(SYNC_PREF_KEYS).forEach(k => { apiPost({ action:'updateSettings', key:k, value: p[k] !== undefined ? p[k] : '' }); });
}
function schedulePrefSync(){ _localPrefsDirtyUntil = Date.now() + 15000; clearTimeout(_prefSyncTimer); _prefSyncTimer = setTimeout(flushPrefs, 700); }

// ═══════════════════════════════════════════
//  APP INITIALIZATION
// ═══════════════════════════════════════════
// ─── Mobile keyboard: push overlay above keyboard when it opens ───
function _syncOverlayToKeyboard() {
  if (!window.visualViewport) return;
  const vvH = window.visualViewport.height;
  const kbH = Math.max(0, window.innerHeight - vvH - window.visualViewport.offsetTop);
  const kbVisible = kbH > 50;
  document.querySelectorAll('.overlay.open').forEach(ol => {
    ol.style.bottom = kbVisible ? kbH + 'px' : '';
  });
  // Constrain sheet height to visible area when keyboard is up
  document.querySelectorAll('.overlay.open .sheet').forEach(sh => {
    sh.style.maxHeight = kbVisible ? Math.floor(vvH * 0.96) + 'px' : '';
  });
}
if (window.visualViewport) {
  window.visualViewport.addEventListener('resize', _syncOverlayToKeyboard);
  window.visualViewport.addEventListener('scroll', _syncOverlayToKeyboard);
}
// Block background scroll when modal open, but allow scrolling inside .sheet
document.addEventListener('touchmove', function(e) {
  if (!document.body.classList.contains('modal-open')) return;
  if (e.target.closest && e.target.closest('.sheet')) return;
  e.preventDefault();
}, { passive: false });

window.addEventListener('DOMContentLoaded', () => {
  applyEmojiPref(); initSheetSwipe();
  const wv = document.getElementById('welcome-version'); if (wv) wv.textContent = 'v' + APP_VERSION;
  // Check first whether we're returning from a Google sign-in redirect. If so,
  // handleAuthRedirect() takes care of everything and we skip the normal boot().
  handleAuthRedirect().then(handled => {
    if (handled) return;
    try { boot(); } catch(e) {
      // Show setup screen as fallback so page is never blank
      const s = document.getElementById('setup-screen');
      if (s) { s.style.display = 'block'; showSetup('sp-welcome'); }
      console.error('Lucid boot error:', e);
    }
  }).catch(e => {
    console.error('Auth redirect check failed:', e);
    try { boot(); } catch(e2) {
      const s = document.getElementById('setup-screen');
      if (s) { s.style.display = 'block'; showSetup('sp-welcome'); }
      console.error('Lucid boot error:', e2);
    }
  });
});

// Recurring items shouldn't wait for a restart: re-check when the app regains
// focus and once an hour while it stays open.
document.addEventListener('visibilitychange', () => {
  if (!document.hidden && document.getElementById('shell') && document.getElementById('shell').style.display !== 'none') {
    try { checkRecurringDue(); } catch(e) {}
    // Polling pauses/backs off while hidden — catch up right away on return
    if (S.mode === 'sheets') { try { syncData(); } catch(e) {} }
  }
});
setInterval(() => {
  if (document.getElementById('shell') && document.getElementById('shell').style.display !== 'none') {
    try { checkRecurringDue(); } catch(e) {}
  }
}, 60*60*1000);

// Swipe down on any bottom sheet to dismiss it
function initSheetSwipe() {
  document.querySelectorAll('.overlay .sheet').forEach(sheet => {
    let startY = 0, dy = 0, dragging = false;
    sheet.addEventListener('touchstart', e => {
      if (sheet.scrollTop > 2) return;
      startY = e.touches[0].clientY; dy = 0; dragging = true;
      sheet.style.transition = 'none';
    }, { passive: true });
    sheet.addEventListener('touchmove', e => {
      if (!dragging) return;
      dy = e.touches[0].clientY - startY;
      if (dy <= 0 || sheet.scrollTop > 2) { sheet.style.transform = ''; return; }
      sheet.style.transform = 'translateY(' + dy + 'px)';
    }, { passive: true });
    sheet.addEventListener('touchend', () => {
      if (!dragging) return;
      dragging = false;
      sheet.style.transition = '';
      sheet.style.transform = '';
      const ov = sheet.closest('.overlay');
      if (dy > 90 && ov) closeOverlay(ov.id);
      dy = 0;
    });
  });
}

function boot() {
  injectManifest();
  const addDateEl = document.getElementById('add-date');
  // Use the LOCAL date string, not valueAsDate (which reads in UTC and shows the
  // wrong day for users ahead of/behind UTC — e.g. Australia would default to
  // yesterday in the evening).
  if (addDateEl) addDateEl.value = todayISO();
  updateGreeting();

  S.cur = CURRENCIES.find(c => c.code === (localStorage.getItem('prizm_cur') || 'AUD')) || {code:'AUD',sym:'A$'};
  S.name = localStorage.getItem('prizm_name') || 'Lucid User';

  // If the user chose their own accent, restore it first so the theme won't override it.
  if (localStorage.getItem('prizm_accent_custom') === '1') {
    setAccent(localStorage.getItem('prizm_accent') || '#6C5CE7', false, false);
  }
  applyTheme(localStorage.getItem('prizm_theme') || 'spectrum');

  const storedMode = localStorage.getItem('prizm_mode'); // 'local' | 'sheets'
  const isFamily = localStorage.getItem('prizm_family') === 'true';
  const url = localStorage.getItem('prizm_url');
  const memberId = localStorage.getItem('prizm_member_id');

  if (!storedMode) { document.getElementById('setup-screen').style.display = 'block'; showSetup('sp-welcome'); return; }

  S.mode = storedMode;
  S.family = isFamily;
  S.url = url || '';
  S.familyCode = localStorage.getItem('prizm_family_code') || '';
  S.familyName = localStorage.getItem('prizm_family_name') || '';

  loadLocalData(); // populate from cache immediately

  if (S.mode === 'local') { enterApp(); return; }
  if (S.mode === 'firebase') {
    // Firebase storage has been removed. Any device still flagged as 'firebase'
    // is sent back to setup to re-choose Local or Google Sheets. Its on-device
    // cache (already loaded above) is untouched, so nothing local is lost.
    localStorage.removeItem('prizm_mode');
    localStorage.removeItem('prizm_fb_uid');
    localStorage.removeItem('prizm_fb_family_code');
    document.getElementById('setup-screen').style.display = 'block';
    showSetup('sp-welcome');
    return;
  }
  // Sheets mode
  if (S.family && !memberId) {
    // Need to pick / create a profile before entering the app
    document.getElementById('setup-screen').style.display = 'block';
    loader(true);
    apiGet().then(d => {
      loader(false);
      if (!d || d.error) { toast('⚠️ Could not connect, using offline cache'); enterApp(); return; }
      applyRemoteData(d);
      showProfilePicker();
    });
    return;
  }

  if (memberId) {
    S.memberId = memberId;
    S.member = { id: memberId, name: localStorage.getItem('prizm_member_name')||S.name, color: localStorage.getItem('prizm_member_color')||'#FAFAFA', pin:'' };
  }

  enterApp();
  syncData().then(() => checkRecurringDue());
  startPolling();
}

function enterApp() {
  document.getElementById('setup-screen').style.display = 'none';
  document.getElementById('shell').style.display = 'flex';
  // Restore persisted preference flags
  S.simpleMode = localStorage.getItem('prizm_simple') === '1';
  S.trackIncome = localStorage.getItem('prizm_track_income') !== '0'; // default true
  S.heroPeriod = localStorage.getItem('prizm_hero_period') || 'month';
  if (localStorage.getItem('prizm_compact')) document.body.setAttribute('data-compact','1');
  const savedCorner = localStorage.getItem('prizm_corner') || 'round';
  setCorner(savedCorner, null);
  applySimpleMode();
  applyTrackIncome();
  buildCatGrid();
  buildRecCatGrid();
  buildBudgetCatOptions();
  applyMemberUI();
  const vEl = document.getElementById('app-version');
  if (vEl) vEl.textContent = 'Lucid v' + APP_VERSION;
  renderAll();
  checkRecurringDue();
}

function toggleAccentExtras() {
  const extras = document.getElementById('accent-extras');
  const btn = document.getElementById('accent-more-btn');
  if (!extras) return;
  const showing = extras.style.display !== 'none';
  extras.style.display = showing ? 'none' : 'contents';
  if (btn) btn.textContent = showing ? '···' : '✕';
}
// When the user picks a custom accent, re-skin the hero card from that accent —
// keeping the exact same style recipe as the stock gradient (light tint → accent →
// cool shift, same angle/stops, accent-tinted shadow). Signature-theme accents
// leave the theme's own hero gradient untouched.
function _applyAccentHero(hex) {
  const root = document.documentElement;
  const r = parseInt(hex.slice(1,3),16), g = parseInt(hex.slice(3,5),16), b = parseInt(hex.slice(5,7),16);
  const lum = (0.299*r + 0.587*g + 0.114*b) / 255;
  root.style.setProperty('--hero-grad',
    `linear-gradient(135deg, color-mix(in srgb, ${hex} 78%, #FFFFFF) 0%, ${hex} 40%, color-mix(in srgb, ${hex} 62%, #22D3EE) 130%)`);
  root.style.setProperty('--hero-txt', lum > 0.62 ? '#141216' : '#FFFFFF');
  root.style.setProperty('--hero-lab', lum > 0.62 ? 'rgba(20,18,22,.64)' : 'rgba(255,255,255,.78)');
  root.style.setProperty('--hero-shadow', `0 14px 44px rgba(${r},${g},${b},.38), 0 2px 8px rgba(0,0,0,.18)`);
  // De-clash the little income/spent dots: if the accent sits in green or red
  // territory, the matching dot would melt into (or fight) the card — swap it to
  // the hero text color so it stays meaningful without looking accidental.
  const hue = _hexHue(hex);
  root.style.setProperty('--hero-dot-inc', (hue >= 90 && hue <= 190) ? 'var(--hero-txt)' : 'var(--green)');
  root.style.setProperty('--hero-dot-exp', (hue <= 25 || hue >= 335) ? 'var(--hero-txt)' : 'var(--red)');
}
function _clearAccentHero() {
  const root = document.documentElement;
  ['--hero-grad','--hero-txt','--hero-lab','--hero-shadow','--hero-dot-inc','--hero-dot-exp']
    .forEach(p => root.style.removeProperty(p));
}
function _hexHue(hex) {
  const r = parseInt(hex.slice(1,3),16)/255, g = parseInt(hex.slice(3,5),16)/255, b = parseInt(hex.slice(5,7),16)/255;
  const mx = Math.max(r,g,b), mn = Math.min(r,g,b), d = mx - mn;
  if (!d) return 0;
  let h = mx === r ? ((g-b)/d) % 6 : mx === g ? (b-r)/d + 2 : (r-g)/d + 4;
  h = Math.round(h * 60); return h < 0 ? h + 360 : h;
}
function setAccent(hex, notify = true, userPick = false) {
  S.accent = hex;
  localStorage.setItem('prizm_accent', hex);
  if (userPick) localStorage.setItem('prizm_accent_custom', '1');
  document.documentElement.style.setProperty('--accent', hex);

  const r = parseInt(hex.slice(1,3),16), g = parseInt(hex.slice(3,5),16), b = parseInt(hex.slice(5,7),16);
  document.documentElement.style.setProperty('--accent-glow', `rgba(${r},${g},${b},0.22)`);
  document.documentElement.style.setProperty('--accent-glow2', `rgba(${r},${g},${b},0.08)`);
  document.documentElement.style.setProperty('--accent-soft', `rgba(${r},${g},${b},0.14)`);

  // Pick a readable text color (near-black or near-white) for content sitting on the accent color
  const luminance = (0.299*r + 0.587*g + 0.114*b) / 255;
  document.documentElement.style.setProperty('--on-accent', luminance > 0.6 ? '#0B0C0E' : '#FFFFFF');

  // Hero card only follows a user-chosen accent when the user opted in (v3.2:
  // accent-only is the default so themes keep their signature hero gradients).
  const _custom = userPick || localStorage.getItem('prizm_accent_custom') === '1';
  if (_custom && localStorage.getItem('prizm_accent_hero') === '1') _applyAccentHero(hex);
  else if (userPick) _clearAccentHero();

  // The always-visible "current" dot shows var(--accent), so it ALWAYS carries the
  // selection ring — theme signature accents match no swatch, and the swatches sit
  // inside the collapsible extras panel (F10).
  document.querySelectorAll('#accent-selector .color-dot').forEach(dot => {
    dot.classList.toggle('active', dot.id === 'accent-active-dot' || dot.getAttribute('data-accent') === hex);
  });

  if (notify) { toast('🎨 Accent Color Updated'); renderAll(); schedulePrefSync(); /* #2 */ }
}

// Each theme's signature accent (applied unless the user picked a custom one)
function useThemeAccent() {
  localStorage.removeItem('prizm_accent_custom');
  _clearAccentHero();
  setAccent(THEME_ACCENTS[S.theme] || '#6C5CE7', true, false);
  _updateAccentUI();
}
function toggleAccentHero() {
  const on = localStorage.getItem('prizm_accent_hero') === '1';
  localStorage.setItem('prizm_accent_hero', on ? '0' : '1');
  if (!on && localStorage.getItem('prizm_accent_custom') === '1') _applyAccentHero(S.accent);
  else _clearAccentHero();
  _updateAccentUI();
  toast(on ? 'Hero card follows the theme' : 'Hero card follows your accent');
  schedulePrefSync();
}
function _updateAccentUI() {
  const chip = document.getElementById('accent-default-chip');
  if (chip) chip.classList.toggle('active', localStorage.getItem('prizm_accent_custom') !== '1');
  const sub = document.getElementById('accent-hero-sub');
  if (sub) sub.textContent = localStorage.getItem('prizm_accent_hero') === '1' ? 'On' : 'Off';
}
const THEME_ACCENTS = { spectrum:'#6C5CE7', minimal:'#9AA2FF', fintech:'#C6F24E', aurora:'#C4B5FD', warm:'#F59E0B', light:'#6366F1', dark:'#6C5CE7', noir:'#6C5CE7' };
function applyTheme(mode, notify = false) {
  // mode: spectrum (default) | minimal | fintech | aurora | warm | light | dark | noir
  if (!mode || mode === 'system') mode = 'spectrum';
  S.theme = mode;
  localStorage.setItem('prizm_theme', mode);

  const root = document.documentElement;
  if (mode === 'spectrum') root.removeAttribute('data-theme'); // base :root = Spectrum
  else root.setAttribute('data-theme', mode);

  document.querySelectorAll('#theme-selector [data-theme-mode]').forEach(btn => {
    btn.classList.toggle('active', btn.getAttribute('data-theme-mode') === mode);
  });

  // Give each theme its signature accent unless the user has chosen their own.
  if (localStorage.getItem('prizm_accent_custom') !== '1') {
    _clearAccentHero(); // let the theme's own hero gradient show
    setAccent(THEME_ACCENTS[mode] || '#6C5CE7', false, false);
  }

  // Keep the browser/status-bar chrome in step with the theme
  setTimeout(() => {
    try {
      const mc = document.querySelector('meta[name="theme-color"]');
      if (mc) mc.content = getComputedStyle(document.body).backgroundColor;
    } catch(e) {}
  }, 60);

  if (notify) { toast('Theme updated'); renderAll(); schedulePrefSync(); /* #2 */ }
}

function injectManifest() {
  // The manifest (with icon) is already injected in <head>; never overwrite it
  // with an icon-less copy · that broke the installed-app icon on Android.
  try {
    const el = document.getElementById('manifest-placeholder');
    if (el && el.href) return;
  } catch(e) {}
}

// ═══════════════════════════════════════════
//  ONBOARDING / SETUP FLOW
// ═══════════════════════════════════════════
function chooseWho(ctx) {
  S.setupContext = ctx;
  S.family = (ctx === 'family');
  showSetup(ctx === 'family' ? 'sp-family-db' : 'sp-solo-db');
}

function quickStart() {
  // Instant local access · no account, no sync.
  localStorage.setItem('prizm_mode', 'local');
  localStorage.setItem('prizm_family', S.setupContext === 'family' ? 'true' : 'false');
  S.mode = 'local';
  S.family = (S.setupContext === 'family');
  loadLocalData();
  if (S.family) { S.onboardingNext = 'profilePicker'; showProfilePicker(); }
  else { S.onboardingNext = 'enterApp'; showSetup('sp-accounts'); }
}

// ── Sample data (for trying the app offline) ─────────────────────────────────
// Fills the app with realistic demo data in local mode, so the whole thing can be
// explored without signing in or typing anything. Replaces on-device data only.
function loadSampleData() {
  const ago  = n => isoLocal(new Date(Date.now() - n * 86400000));
  const soon = n => isoLocal(new Date(Date.now() + n * 86400000));
  const dom  = d => Number(String(d).slice(8, 10));
  S.name = 'Alex'; localStorage.setItem('prizm_name', 'Alex');
  S.accts = [
    { id: 'ac_s1', name: 'Everyday',    type: 'Checking', balance: 2413.80 },
    { id: 'ac_s2', name: 'Savings',     type: 'Savings',  balance: 8250.00 },
    { id: 'ac_s3', name: 'Credit Card', type: 'Credit',   balance: -486.20 }
  ];
  const T = (id, d, desc, amt, type, cat, acct) => ({ id: 'tx_sample_' + id, date: d, description: desc, amount: amt, type: type, category: cat, account: acct, notes: '', memberId: '', ts: Date.now() - id * 1000 });
  S.txs = [
    T(1, ago(0),  'Coffee',          5.50,   'expense', 'Coffee',    'Everyday'),
    T(2, ago(1),  'Woolworths',      82.30,  'expense', 'Food',      'Everyday'),
    T(3, ago(2),  'Uber',            18.40,  'expense', 'Transport', 'Credit Card'),
    T(4, ago(3),  'Gym membership',  22.00,  'expense', 'Fitness',   'Everyday'),
    T(5, ago(4),  'Netflix',         16.99,  'expense', 'Bills',     'Credit Card'),
    T(6, ago(5),  'Pharmacy',        34.10,  'expense', 'Health',    'Everyday'),
    T(7, ago(6),  'Dinner out',      56.00,  'expense', 'Food',      'Credit Card'),
    T(8, ago(7),  'Salary',          2400.00,'income',  'Salary',    'Everyday'),
    T(9, ago(8),  'Petrol',          70.00,  'expense', 'Transport', 'Everyday'),
    T(10, ago(10),'Amazon',          43.25,  'expense', 'Shopping',  'Credit Card'),
    T(11, ago(12),'Groceries',       91.60,  'expense', 'Food',      'Everyday'),
    T(12, ago(14),'Electricity',     120.00, 'expense', 'Bills',     'Everyday'),
    T(13, ago(16),'Movies',          28.00,  'expense', 'Shopping',  'Everyday'),
    T(14, ago(18),'Coffee',          4.80,   'expense', 'Coffee',    'Everyday'),
    T(15, ago(20),'Train fare',      12.00,  'expense', 'Transport', 'Everyday'),
    T(16, ago(24),'New shoes',       89.99,  'expense', 'Shopping',  'Credit Card'),
    T(17, ago(28),'Groceries',       76.40,  'expense', 'Food',      'Everyday'),
    T(18, ago(30),'Freelance job',   350.00, 'income',  'Freelance', 'Everyday'),
    T(19, ago(34),'Doctor visit',    60.00,  'expense', 'Health',    'Everyday'),
    T(20, ago(38),'Restaurant',      48.20,  'expense', 'Food',      'Credit Card'),
    T(21, ago(40),'Gym membership',  22.00,  'expense', 'Fitness',   'Everyday'),
    T(22, ago(44),'Internet',        65.00,  'expense', 'Bills',     'Everyday'),
    T(23, ago(48),'Groceries',       88.10,  'expense', 'Food',      'Everyday'),
    T(24, ago(52),'Salary',          2400.00,'income',  'Salary',    'Everyday')
  ];
  sortTxs(S.txs);
  S.budgets = [
    { id: 'bg_s1', category: 'Food',      limit: 600 },
    { id: 'bg_s2', category: 'Shopping',  limit: 300 },
    { id: 'bg_s3', category: 'Transport', limit: 150 }
  ];
  S.recurring = [
    { id: 'rec_s1', name: 'Rent',    amount: 1450,  type: 'expense', category: 'Bills',   account: 'Everyday',    frequency: 'monthly', nextDate: soon(4),  targetDay: dom(soon(4)),  memberId: '', active: true },
    { id: 'rec_s2', name: 'Netflix', amount: 16.99, type: 'expense', category: 'Bills',   account: 'Credit Card', frequency: 'monthly', nextDate: soon(9),  targetDay: dom(soon(9)),  memberId: '', active: true },
    { id: 'rec_s3', name: 'Gym',     amount: 22,    type: 'expense', category: 'Fitness', account: 'Everyday',    frequency: 'monthly', nextDate: soon(12), targetDay: dom(soon(12)), memberId: '', active: true },
    { id: 'rec_s4', name: 'Salary',  amount: 2400,  type: 'income',  category: 'Salary',  account: 'Everyday',    frequency: 'monthly', nextDate: soon(6),  targetDay: dom(soon(6)),  memberId: '', active: true }
  ];
  localStorage.setItem('prizm_goals', JSON.stringify([
    { id: 'goal_s1', name: 'Emergency fund', target: 5000, saved: 2000, emoji: '🛟' },
    { id: 'goal_s2', name: 'Bali holiday',   target: 2500, saved: 820,  emoji: '🏝️' }
  ]));
  S.mode = 'local'; S.family = false;
  localStorage.setItem('prizm_mode', 'local');
  localStorage.setItem('prizm_family', 'false');
  saveLocalData();
  document.getElementById('setup-screen').style.display = 'none';
  try { buildCatGrid(); buildRecCatGrid(); buildBudgetCatOptions(); } catch(e) {}
  enterApp();
  toast('✨ Sample data loaded');
}
function confirmLoadSample() {
  showCustomConfirm('Load sample data?',
    'This replaces the data on this device with demo transactions, budgets, bills and goals so you can explore. You can start fresh anytime from Settings.',
    'Load sample data', loadSampleData);
}


// Rough onboarding progress per page — drives the top progress bar fill (F22).
const SETUP_PROGRESS = {
  'sp-welcome':10, 'sp-who':20, 'sp-firebase':35, 'sp-solo-db':35, 'sp-family-db':35,
  'sp-auto-1':40, 'sp-auto-2':48, 'sp-auto-3':55, 'sp-sheets-solo':55, 'sp-sheets-family':55, 'sp-family-join':55,
  'sp-family-code':75, 'sp-profile-picker':80, 'sp-new-member':85,
  'sp-pin-entry':85, 'sp-accounts':90
};
function showSetup(id) {
  document.querySelectorAll('.setup-page').forEach(p => p.classList.remove('active'));
  document.getElementById(id).classList.add('active');
  const scr = document.getElementById('setup-screen');
  if (scr) scr.style.setProperty('--setup-prog', (SETUP_PROGRESS[id] || 10) + '%');
  if (id === 'sp-new-member') {
    S.newMemberPin = '';
    document.getElementById('new-member-name').value = '';
    updatePinDots('new-member-pin-dots', '');
    document.getElementById('new-member-save').disabled = true;
    const kb = document.getElementById('new-member-pin-kb'); if (kb) kb.value = '';
  }
  if (id === 'sp-accounts') {
    renderOnboardAcctsList();
    const nEl = document.getElementById('onboard-name');
    if (nEl && !nEl.value) nEl.value = (S.name && S.name !== 'Lucid User') ? S.name : '';
  }
}

// ═══════════════════════════════════════════
//  ONBOARDING: NAME YOUR CATEGORIES (one-time, skippable)
// ═══════════════════════════════════════════
function renderOnboardAcctsList() {
  const wrap = document.getElementById('onboard-accts-list');
  if (!S.onboardAccts.length) { wrap.innerHTML = ''; return; }
  wrap.innerHTML = S.onboardAccts.map((a, i) => `
    <div class="srow" style="cursor:default;padding:10px 14px">
      <div class="srow-ico">${getAcctIcon(a.type)}</div>
      <div class="srow-body"><div class="srow-title">${a.name}</div><div class="srow-sub">${a.type}${a.balance ? ' · ' + fmtMoney(a.balance) : ''}</div></div>
      <div class="icon-btn" style="width:32px;height:32px" onclick="removeOnboardAcct(${i})">✕</div>
    </div>`).join('');
}

function addOnboardAccount() {
  const name = document.getElementById('onboard-acct-name').value.trim();
  if (!name) { toast('Enter an account name'); return; }
  const type = document.getElementById('onboard-acct-type').value;
  const balance = parseFloat(document.getElementById('onboard-acct-bal').value) || 0;
  S.onboardAccts.push({ name, type, balance });
  document.getElementById('onboard-acct-name').value = '';
  document.getElementById('onboard-acct-bal').value = '';
  renderOnboardAcctsList();
}

function removeOnboardAcct(i) {
  S.onboardAccts.splice(i, 1);
  renderOnboardAcctsList();
}

async function finishOnboarding(skip) {
  const nEl = document.getElementById('onboard-name');
  const nm = nEl ? nEl.value.trim() : '';
  if (nm) { S.name = nm; localStorage.setItem('prizm_name', nm); apiPost({ action:'updateSettings', key:'name', value:nm }); }
  // Pick up a final account typed in the fields but not yet "added"
  if (!skip) {
    const name = document.getElementById('onboard-acct-name').value.trim();
    if (name) {
      const type = document.getElementById('onboard-acct-type').value;
      const balance = parseFloat(document.getElementById('onboard-acct-bal').value) || 0;
      S.onboardAccts.push({ name, type, balance });
    }
  }
  if (!skip && S.onboardAccts.length) {
    loader(true);
    for (let i = 0; i < S.onboardAccts.length; i++) {
      const a = S.onboardAccts[i];
      await apiPost({ action: 'addAccount', id: 'ac_' + Date.now() + '_' + i, name: a.name, type: a.type, balance: a.balance });
    }
    loader(false);
  }
  S.onboardAccts = [];
  if (S.onboardingNext === 'profilePicker') showProfilePicker();
  else enterApp();
}

function extractDeployId(url) {
  const m = url.match(/\/s\/([a-zA-Z0-9_-]+)\/exec/);
  return m ? m[1] : url;
}
function buildUrlFromCode(code) {
  code = code.trim();
  if (code.includes('script.google.com')) return code;
  return 'https://script.google.com/macros/s/' + code + '/exec';
}

// Friendly validation for the pasted Apps Script web app link
function normalizeScriptUrl(raw) {
  const url = (raw || '').trim();
  if (!url) return { error: 'Paste your web app link first.' };
  if (!url.includes('script.google.com')) return { error: "That doesn't look like a Google link. It should start with https://script.google.com" };
  if (!url.includes('/exec')) return { error: 'Almost there · copy the deployment\'s "Web app URL" (it ends in /exec), not the editor address.' };
  return { url };
}

async function connectPersonal() {
  const dEl = document.getElementById('personal-url');
  const rawUrl = dEl ? dEl.value : '';
  const { url, error } = normalizeScriptUrl(rawUrl);
  if (error) return toast('⚠️ ' + error);
  loader(true);
  S.url = url;
  const d = await apiGet();
  loader(false);
  if (!d || d.error) return toast('⚠️ Could not connect. Check that "Who has access" is set to Anyone and try again.');
  applyRemoteData(d);
  localStorage.setItem('prizm_url', url);
  localStorage.setItem('prizm_mode', 'sheets');
  localStorage.setItem('prizm_family', 'false');
  S.mode = 'sheets'; S.family = false;
  startPolling();
  toast('✅ Connected to your Sheet');
  S.onboardingNext = 'enterApp';
  showSetup('sp-accounts');
}

async function connectFamilyCreate() {
  const fuDEl = document.getElementById('family-url');
  const fnDEl = document.getElementById('family-name');
  const famName = (fnDEl ? fnDEl.value.trim() : '') || 'My Family';
  const { url, error } = normalizeScriptUrl(fuDEl ? fuDEl.value : '');
  if (error) return toast('⚠️ ' + error);
  loader(true);
  S.url = url;
  const d = await apiGet();
  loader(false);
  if (!d || d.error) return toast('⚠️ Could not connect. Check that "Who has access" is set to Anyone and try again.');
  applyRemoteData(d);
  S.familyName = famName;
  S.familyCode = extractDeployId(url);
  localStorage.setItem('prizm_url', url);
  localStorage.setItem('prizm_mode', 'sheets');
  localStorage.setItem('prizm_family', 'true');
  localStorage.setItem('prizm_family_code', S.familyCode);
  localStorage.setItem('prizm_family_name', famName);
  S.mode = 'sheets'; S.family = true;
  // Don't start polling here · user is still in the setup wizard.
  // Polling starts inside chooseMember() → enterApp() after profile selection.
  apiPost({action:'updateSettings', key:'familyName', value:famName});
  document.getElementById('family-code-val').textContent = S.familyCode;
  showSetup('sp-family-code');
}

function copyFamilyCode() {
  if (!S.familyCode) { toast('No family code yet · connect Google or a Sheet first'); return; }
  navigator.clipboard.writeText(S.familyCode).then(() => toast('📋 Family code copied')).catch(() => toast('⚠️ Could not copy'));
}

function finishFamilySetup() {
  S.onboardingNext = 'profilePicker';
  showProfilePicker();
}

async function connectFamilyJoin() {
  const code = document.getElementById('join-code').value.trim();
  if (!code) return toast('⚠️ Enter a code');
  const url = buildUrlFromCode(code);
  loader(true);
  S.url = url;
  const d = await apiGet();
  loader(false);
  if (!d || d.error) return toast('⚠️ Could not connect with that code');
  applyRemoteData(d);
  S.familyCode = extractDeployId(url);
  S.familyName = (d.settings && d.settings.familyName) || 'Family';
  localStorage.setItem('prizm_url', url);
  localStorage.setItem('prizm_mode', 'sheets');
  localStorage.setItem('prizm_family', 'true');
  localStorage.setItem('prizm_family_code', S.familyCode);
  localStorage.setItem('prizm_family_name', S.familyName);
  S.mode = 'sheets'; S.family = true;
  // Polling starts after profile selection (chooseMember → startPolling)
  showProfilePicker();
}

function skipSetup() {
  localStorage.setItem('prizm_mode', 'local');
  localStorage.setItem('prizm_family', 'false');
  S.mode = 'local'; S.family = false;
  loadLocalData();
  toast('📁 Offline ledger initialized');
  S.onboardingNext = 'enterApp';
  showSetup('sp-accounts');
}

function applyRemoteData(d) {
  try {
    const sig = JSON.stringify([d.transactions||[], d.accounts||[], d.members||[], d.recurring||[], d.budgets||[], d.settings||{}]);
    if (sig === S._lastRemoteSig) return false; // nothing changed · skip re-render
    S._lastRemoteSig = sig;
  } catch(e) {}
  S.txs = sortTxs(d.transactions || []);
  S.accts = d.accounts || [];
  S.members = d.members || [];
  S.recurring = d.recurring || [];
  S.budgets = d.budgets || [];
  if (d.settings) {
    if (d.settings.currency) { const c = CURRENCIES.find(x => x.code === d.settings.currency); if (c) { S.cur = c; localStorage.setItem('prizm_cur', c.code); } }
    if (d.settings.name) { S.name = d.settings.name; localStorage.setItem('prizm_name', S.name); }
    if (d.settings.familyName) S.familyName = d.settings.familyName;
    applyIncomingPrefs(d.settings); // #2 · sync custom categories, simple mode, theme, etc.
  }
  // Refresh the active member object in case name/colour changed remotely
  if (S.member && S.member.id) {
    const fresh = S.members.find(m => m.id === S.member.id);
    if (fresh) {
      S.member = { ...S.member, name: fresh.name, color: fresh.color };
      localStorage.setItem('prizm_member_name', fresh.name);
      localStorage.setItem('prizm_member_color', fresh.color);
    }
  }
  saveLocalData();
  return true;
}

// ═══════════════════════════════════════════
//  PROFILE PICKER & PIN HANDLING
// ═══════════════════════════════════════════
function showProfilePicker() {
  document.getElementById('profile-picker-title').textContent = S.members.length ? "Who's this?" : "Welcome!";
  const grid = document.getElementById('profile-grid');
  grid.innerHTML = S.members.map(m => `
    <div class="member-card" onclick="selectProfile('${esc(m.id)}')">
      <div class="member-avatar" style="background:${esc(m.color)}">${esc((m.name||'?').charAt(0).toUpperCase())}</div>
      <div class="member-card-name">${esc(m.name)}</div>
    </div>`).join('');
  document.getElementById('new-member-back').style.display = S.members.length ? 'flex' : 'none';
  showSetup('sp-profile-picker');
}

function selectProfile(id) {
  const m = S.members.find(x => x.id === id);
  if (!m) return;
  S.pinTarget = m;
  const av = document.getElementById('pin-entry-avatar');
  av.style.background = m.color;
  av.textContent = (m.name||'?').charAt(0).toUpperCase();
  document.getElementById('pin-entry-name').textContent = m.name;
  S.pinEntryVal = '';
  updatePinDots('pin-entry-dots', '');
  const kb = document.getElementById('pin-entry-kb'); if (kb) kb.value = '';
  showSetup('sp-pin-entry');
}

// #10 · PINs are a soft "stop a sibling peeking" lock, NOT real access control
// (client-side, only 10,000 combinations). We at least avoid storing them in
// plaintext: new PINs are salted-hashed; legacy plaintext PINs still verify.
async function hashPin(p){
  try {
    const buf = await crypto.subtle.digest('SHA-256', new TextEncoder().encode('prizm-pin:' + p));
    return 'h:' + Array.from(new Uint8Array(buf)).map(x => x.toString(16).padStart(2,'0')).join('');
  } catch(e) { return String(p); }
}
async function verifyPin(input, stored){
  if (stored == null) return false;
  const s = String(stored);
  if (s.indexOf('h:') === 0) return (await hashPin(input)) === s;
  return String(input) === s; // legacy plaintext
}

function pinEntryKey(k) {
  if (k === 'clear') S.pinEntryVal = '';
  else if (k === 'back') S.pinEntryVal = S.pinEntryVal.slice(0,-1);
  else if (S.pinEntryVal.length < 4) S.pinEntryVal += k;
  updatePinDots('pin-entry-dots', S.pinEntryVal);
  _pinEntryTryVerify();
}
function _pinEntryTryVerify() {
  if (S.pinEntryVal.length !== 4) return;
  const entered = S.pinEntryVal;
  setTimeout(async () => {
    if (await verifyPin(entered, S.pinTarget.pin)) {
      chooseMember(S.pinTarget);
    } else {
      toast('⚠️ Incorrect PIN');
      S.pinEntryVal = '';
      updatePinDots('pin-entry-dots', '');
      const kb = document.getElementById('pin-entry-kb'); if (kb) kb.value = '';
    }
  }, 150);
}
function pinKbInput(ctx, el) {
  const v = el.value.replace(/\D/g, '').slice(0, 4);
  el.value = v;
  if (ctx === 'new') {
    S.newMemberPin = v;
    updatePinDots('new-member-pin-dots', v);
    document.getElementById('new-member-save').disabled = v.length !== 4;
  } else if (ctx === 'member') {
    S.newMemberPin = v;
    updatePinDots('member-pin-dots', v);
    const name = document.getElementById('member-name').value.trim();
    document.getElementById('member-save-btn').disabled = !(name && v.length === 4);
  } else if (ctx === 'entry') {
    S.pinEntryVal = v;
    updatePinDots('pin-entry-dots', v);
    _pinEntryTryVerify();
  }
}

function chooseMember(m) {
  S.member = m;
  S.memberId = m.id;
  localStorage.setItem('prizm_member_id', m.id);
  localStorage.setItem('prizm_member_name', m.name);
  localStorage.setItem('prizm_member_color', m.color);
  enterApp();
  startPolling();
}

function switchProfile() {
  localStorage.removeItem('prizm_member_id');
  localStorage.removeItem('prizm_member_name');
  localStorage.removeItem('prizm_member_color');
  S.member = null;
  S.memberId = null;
  stopPolling();
  document.getElementById('shell').style.display = 'none';
  document.getElementById('setup-screen').style.display = 'block';
  // Refresh member list from remote so newly-joined members appear
  if (S.mode === 'sheets' && S.url) {
    loader(true);
    apiGet().then(d => {
      loader(false);
      if (d && !d.error) applyRemoteData(d);
      showProfilePicker();
    }).catch(() => { loader(false); showProfilePicker(); });
  } else if (S.mode === 'firebase') {
    loader(true);
    firebaseRead().then(d => { loader(false); if (d) applyRemoteData(d); showProfilePicker(); })
      .catch(() => { loader(false); showProfilePicker(); });
  } else {
    showProfilePicker();
  }
}

function pinKey(k) {
  if (k === 'clear') S.newMemberPin = '';
  else if (k === 'back') S.newMemberPin = S.newMemberPin.slice(0,-1);
  else if (S.newMemberPin.length < 4) S.newMemberPin += k;
  updatePinDots('new-member-pin-dots', S.newMemberPin);
  document.getElementById('new-member-save').disabled = S.newMemberPin.length !== 4;
}

async function saveNewMember() {
  const name = document.getElementById('new-member-name').value.trim();
  if (!name) return toast('⚠️ Enter your name');
  if (S.newMemberPin.length !== 4) return toast('⚠️ Choose a 4-digit PIN');
  const color = MEMBER_COLORS[S.members.length % MEMBER_COLORS.length];
  const id = 'mem_' + Date.now();
  const pinHash = await hashPin(S.newMemberPin); // #10
  const m = {id, name, pin: pinHash, color};
  loader(true);
  await apiPost({action:'addMember', id, name, pin:pinHash, color});
  loader(false);
  chooseMember(m);
}

function updatePinDots(containerId, val) {
  document.querySelectorAll('#'+containerId+' .pin-dot').forEach((d,i) => d.classList.toggle('filled', i < val.length));
}

function applyMemberUI() {
  document.getElementById('family-section').style.display = S.family ? 'block' : 'none';
  document.getElementById('switch-profile-btn').style.display = S.family ? 'flex' : 'none';
  document.getElementById('member-chip-row').style.display = S.family ? 'flex' : 'none';
  if (S.family) {
    document.getElementById('family-code-sub').textContent = S.familyCode ? ('Code: ' + S.familyCode) : (S.mode === 'local' ? 'Local family · connect a sync option to invite others' : 'Tap to copy & share');
  }
}

// ═══════════════════════════════════════════
//  DATA RECOVERY & OPTIMISTIC ENGINE
// ═══════════════════════════════════════════
// FIREBASE — REMOVED. Lucid stores data only on-device (local) or in the user's
// own Google Sheet. No Firebase SDK is loaded and no data ever leaves the user's
// Google account. These inert globals remain only so legacy guarded branches
// (S.mode==='firebase', which can no longer be set) stay reference-safe.
let fbApp=null,fbAuth=null,fbDb=null;
// ── Google Identity Services · popup token flow ──
// Immune to the firebaseapp.com redirect problem (third-party storage partitioning)
// because no middleman page is ever involved: Google's own popup returns the token
// directly to this page. Requires the OAuth Web client ID from the Firebase project
// (Firebase console → Authentication → Sign-in method → Google → Web SDK configuration).
const GOOGLE_CLIENT_ID='63943061976-05855nuk42jl0oidbcimtcnjkeq558hf.apps.googleusercontent.com';
function gisReady(){ return !!(window.google && google.accounts && google.accounts.oauth2 && GOOGLE_CLIENT_ID.indexOf('.apps.googleusercontent.com')>0 && GOOGLE_CLIENT_ID.indexOf('PASTE_')!==0); }
function _gisCancelled(e){ var m=String((e&&e.message)||e||''); return /popup_closed|access_denied|user_cancel|closed/i.test(m); }
function gisToken(scope){
  return new Promise(function(resolve,reject){
    try{
      var tc=google.accounts.oauth2.initTokenClient({
        client_id:GOOGLE_CLIENT_ID,
        scope:scope,
        callback:function(resp){ if(resp && resp.access_token) resolve(resp.access_token); else reject(new Error((resp && resp.error)||'no_token')); },
        error_callback:function(err){ reject(new Error((err && (err.type||err.message))||'popup_closed')); }
      });
      tc.requestAccessToken();
    }catch(e){ reject(e); }
  });
}
function initFirebase(){/* Firebase removed — no-op. */}
// Popup-first sign-in. signInWithRedirect silently fails on any domain other than
// the Firebase authDomain in modern browsers (third-party storage partitioning eats
// the redirect session → getRedirectResult() returns null → user lands back on the
// home screen as if nothing happened). signInWithPopup is immune to that AND hands
// back the OAuth token in the same page session, so auto-setup can run end-to-end
// without a reload. Redirect is kept only as a fallback for popup blockers.
function _popupFellBack(e){
  return !!(e && (e.code==='auth/popup-blocked' || e.code==='auth/operation-not-supported-in-this-environment' || e.code==='auth/web-storage-unsupported'));
}
function _popupCancelled(e){
  return !!(e && (e.code==='auth/popup-closed-by-user' || e.code==='auth/cancelled-popup-request' || e.code==='auth/user-cancelled'));
}
// Firebase removed: "Sign in with Google" now means "sign in & build your Google
// Sheet". Route straight to the GIS-powered Sheets auto-setup.
async function connectWithGoogle(){ return autoSetupSheet(S.setupContext==='family'); }

// _finishGoogleSignIn() removed with Firebase — the Google Sheets auto-setup
// (finishAutoSetupSheet) is now the only "Sign in with Google" destination.

// Firebase used a signInWithRedirect() fallback that needed post-reload handling.
// GIS uses popups only, so there is no redirect to recover — this is now a no-op
// that lets boot() run normally on every load.
async function handleAuthRedirect(){ return false; }
// ═══════════════════════════════════════════
//  AUTOMATIC SHEET SETUP (v3 flagship)
//  Google sign-in → create Sheet → install backend → deploy → autofill URL
// ═══════════════════════════════════════════
const AUTO_SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/script.projects',
  'https://www.googleapis.com/auth/script.deployments'
];
async function gApi(token, method, url, body) {
  const r = await fetch(url, { method, headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' }, body: body ? JSON.stringify(body) : undefined });
  const d = await r.json().catch(() => ({}));
  if (!r.ok) { const err = new Error((d.error && d.error.message) || ('HTTP ' + r.status)); err.status = r.status; throw err; }
  return d;
}
// ── Guided auto-setup wizard ──
function openAutoSetup(family){
  S.setupContext = family ? 'family' : 'solo'; S.family = !!family;
  wizSetStep(localStorage.getItem('prizm_asapi_on')==='true' ? 2 : 1);
}
function wizSetStep(n){ showSetup('sp-auto-' + n); }
function wizOpenSwitch(){ window.open('https://script.google.com/home/usersettings','_blank'); }
function wizSwitchDone(){ localStorage.setItem('prizm_asapi_on','true'); wizSetStep(2); toast('Great — now sign in'); }
async function autoSetupSheet(family) {
  if (location.protocol === 'file:') { showCustomAlert('Hosting required', 'Automatic setup needs the app served over http(s). Use the manual steps below, or host Lucid first.'); return; }
  if (gisReady()) {
    // New reliable path: token straight from Google's popup, no middleman page.
    try {
      loader(true); toast('Opening Google sign-in…');
      const token = await gisToken(AUTO_SCOPES.join(' '));
      await finishAutoSetupSheet(token, family);
    } catch(e) {
      loader(false);
      if (!_gisCancelled(e)) { toast('Sign-in failed'); console.warn(e); }
    }
    return;
  }
  // GIS is the only sign-in path now (Firebase removed). If it hasn't loaded yet,
  // ask the user to retry rather than falling back to anything.
  toast('Google sign-in is still loading — try again in a moment.');
}
async function finishAutoSetupSheet(token, family) {
  wizSetStep(3);
  let ssId = null, ssUrl = '';
  try {
    toast('Creating your Sheet…');
    const ss = await gApi(token, 'POST', 'https://sheets.googleapis.com/v4/spreadsheets', { properties: { title: family ? 'Lucid Family Ledger' : 'Lucid Ledger' } });
    ssId = ss.spreadsheetId; ssUrl = ss.spreadsheetUrl || '';
    toast('Installing the backend…');
    const proj = await gApi(token, 'POST', 'https://script.googleapis.com/v1/projects', { title: 'Lucid Backend', parentId: ssId });
    await gApi(token, 'PUT', 'https://script.googleapis.com/v1/projects/' + proj.scriptId + '/content', {
      files: [
        { name: 'Code', type: 'SERVER_JS', source: SCRIPT_CODE },
        { name: 'appsscript', type: 'JSON', source: JSON.stringify({ timeZone: 'Etc/GMT', exceptionLogging: 'STACKDRIVER', runtimeVersion: 'V8', webapp: { access: 'ANYONE_ANONYMOUS', executeAs: 'USER_DEPLOYING' }, oauthScopes: ['https://www.googleapis.com/auth/spreadsheets'] }) }
      ]
    });
    toast('Publishing…');
    const ver = await gApi(token, 'POST', 'https://script.googleapis.com/v1/projects/' + proj.scriptId + '/versions', { description: 'Lucid v' + APP_VERSION });
    const dep = await gApi(token, 'POST', 'https://script.googleapis.com/v1/projects/' + proj.scriptId + '/deployments', { versionNumber: ver.versionNumber, manifestFileName: 'appsscript', description: 'Lucid' });
    let execUrl = '';
    (dep.entryPoints || []).forEach(ep => { if (ep.webApp && ep.webApp.url) execUrl = ep.webApp.url; });
    if (!execUrl && dep.deploymentId) execUrl = 'https://script.google.com/macros/s/' + dep.deploymentId + '/exec';
    loader(false);
    // Autofill the URL so the user can see it landed
    const input = document.getElementById(family ? 'family-url' : 'personal-url');
    if (input) input.value = execUrl;
    showCustomConfirm('One tap to finish',
      'Google needs a one-time OK to run your new backend. A script.google.com tab will open. Choose your account and tap Allow, then come back here.',
      'Open & authorize',
      () => { window.open(execUrl, '_blank'); _pollAutoConnect(execUrl, family); });
  } catch(e) {
    loader(false); console.warn('auto setup:', e);
    const msg = String(e.message || '');
    if (e.status === 403 && /Apps Script API|has not been used|is disabled/i.test(msg)) {
      showCustomConfirm('One-time Google setting',
        'Google requires every user to enable "Google Apps Script API" once before apps can install scripts (their rule, ~10 seconds). Tap below to open the google.com settings page, turn the switch ON, then come back and run setup again.' + (ssUrl ? '\n\nYour Sheet was already created and is safe in your Drive.' : ''),
        'Open Google setting',
        () => { localStorage.removeItem('prizm_asapi_on'); wizSetStep(1); window.open('https://script.google.com/home/usersettings', '_blank'); });
    } else {
      showCustomAlert('Automatic setup hit a snag', msg + (ssUrl ? '\n\nYour Sheet was created and is safe in your Drive. You can finish with the manual steps below.' : '\n\nYou can use the manual steps below instead.'));
      showSetup(family ? 'sp-sheets-family' : 'sp-sheets-solo');
    }
  }
}
function _wiz3Spinner(spinning) {
  const sp = document.getElementById('wiz3-spinner');
  if (sp) sp.style.animationPlayState = spinning ? 'running' : 'paused';
}
// Runs _finishAutoConnect() in its OWN try/catch, deliberately kept separate
// from the fetch/JSON-parse retry loop below. Previously _finishAutoConnect()
// was called from inside the polling try block, so if it threw for any reason
// (bad data shape, a missing DOM node, apiPost throwing, etc.) that error was
// silently swallowed by the poll's catch(e){} — and because S._autoDone had
// already been set to true right before the throw, every future poll tick
// just returned immediately, permanently freezing the wizard on Step 3 with
// the spinner still animating and zero feedback. This wrapper fixes that:
// failures are logged, shown to the user, and left retryable.
function _safeFinishAutoConnect(url, family, d) {
  S._autoLastPayload = d; // cache so a retry can skip re-fetching
  try {
    _finishAutoConnect(url, family, d);
    return true;
  } catch (e) {
    console.error('Auto setup: authorization succeeded but finishing setup failed:', e);
    S._autoDone = false; // let polling / manual retry try again
    _wiz3Spinner(true);
    const el = document.getElementById('wiz3-status');
    if (el) el.textContent = 'Connected, but finishing setup hit a snag — tap "connect now" to retry.';
    showCustomAlert('Almost there',
      'Google authorized the connection, but Lucid hit an error while finishing setup:\n\n' + (e && e.message ? e.message : String(e)) +
      '\n\nTap "connect now" to retry — no need to reopen the Google tab or sign in again.');
    return false;
  }
}
function _pollAutoConnect(url, family) {
  S._autoExecUrl = url; S._autoFamily = !!family;
  toast('Waiting for authorization…');
  // Fresh Apps Script API deployments can take a few minutes to propagate,
  // so this window is intentionally generous rather than the previous 3 min.
  const deadline = Date.now() + 360000;
  const status = () => document.getElementById('wiz3-status');
  const tryOnce = async () => {
    if (S._autoDone) return;
    let d = null, reachErr = null;
    try {
      const r = await fetch(url + '?action=getAll&_ts=' + Date.now() + authK());
      d = await r.json(); // unauthorized returns HTML → json() throws → treated as "not yet"
    } catch(e) {
      reachErr = e; // network/CORS failure vs. "still needs authorization" — logged below
    }
    if (d && !d.error) {
      S._autoDone = true;
      if (_safeFinishAutoConnect(url, family, d)) return; // success — wizard has moved on
      // _safeFinishAutoConnect already reset S._autoDone and updated the UI on failure;
      // fall through to keep polling so a transient error can resolve itself.
    } else if (d && d.error) {
      // The engine responded with a real application error (not just "not
      // authorized yet" — that case returns HTML, which fails JSON parsing
      // and lands in the reachErr branch below instead). Retrying blindly
      // would never fix this, so surface it immediately instead of masking
      // it behind a generic "still waiting" message forever.
      console.error('Auto setup: engine returned an error:', d.error);
      _wiz3Spinner(false);
      const el = status(); if (el) el.textContent = 'The engine reported an error — see details below.';
      showCustomAlert('The engine replied with an error', String(d.error) + '\n\nSend this message to your helper if it keeps happening.');
      return; // stop polling; retrying an application error won't help
    } else if (reachErr) {
      console.warn('Auto setup: poll attempt could not reach the engine:', reachErr);
    }
    if (Date.now() < deadline) {
      const el = status();
      if (el && !S._autoDone) el.textContent = 'Still waiting — once you have tapped Allow on the Google tab, tap "connect now" below.';
      setTimeout(tryOnce, 4000);
    } else {
      _wiz3Spinner(false);
      const el = status(); if (el) el.textContent = 'Timed out — tap "connect now" below after allowing, or reopen the Google tab.';
      toast('Still waiting · tap "connect now" once you have allowed');
    }
  };
  S._autoDone = false;
  _wiz3Spinner(true);
  setTimeout(tryOnce, 5000);
}
async function wizTryConnectNow() {
  const url = S._autoExecUrl;
  if (!url) { showCustomAlert('Nothing to connect yet', 'Run the setup again from the start — the engine link from this session was lost (it is created fresh each run).'); return; }
  _wiz3Spinner(true);
  const el = document.getElementById('wiz3-status'); if (el) el.textContent = 'Checking…';
  loader(true);
  try {
    const r = await fetch(url + '?action=getAll&_ts=' + Date.now());
    const text = await r.text();
    loader(false);
    let d = null;
    try { d = JSON.parse(text); } catch(e) {}
    if (d && !d.error) { S._autoDone = true; _safeFinishAutoConnect(url, S._autoFamily, d); return; }
    if (d && d.error) {
      _wiz3Spinner(false);
      showCustomAlert('The engine replied with an error', String(d.error) + '\n\nSend this message to your helper if it keeps happening.');
    } else {
      if (el) el.textContent = 'Still waiting — once you have tapped Allow on the Google tab, tap "connect now" below.';
      showCustomAlert('Not authorized yet',
        'The engine is reachable but still asks for permission. Tap "Open the Google tab again", choose your account, tap Advanced \u2192 Go to Lucid Backend (unsafe) \u2192 Allow. Then come back and tap "connect now".');
    }
  } catch(e) {
    loader(false);
    console.warn('Auto setup: manual connect-now could not reach the engine:', e);
    if (el) el.textContent = 'Still waiting — once you have tapped Allow on the Google tab, tap "connect now" below.';
    showCustomAlert('Could not reach the engine',
      (e && e.message ? e.message : 'Network error') + '\n\nCheck your internet, then tap "connect now" again. If this repeats, run the setup once more from the start.');
  }
}
function wizOpenEngine() {
  if (S._autoExecUrl) window.open(S._autoExecUrl, '_blank');
  else showCustomAlert('Link not available', 'Run the setup again from the start to get a fresh engine link.');
}
function _finishAutoConnect(url, family, d) {
  applyRemoteData(d);
  S.url = url;
  localStorage.setItem('prizm_url', url);
  localStorage.setItem('prizm_mode', 'sheets');
  localStorage.setItem('prizm_family', family ? 'true' : 'false');
  S.mode = 'sheets'; S.family = !!family;
  if (family) {
    S.familyCode = extractDeployId(url);
    localStorage.setItem('prizm_family_code', S.familyCode);
    const fnEl = document.getElementById('family-name');
    S.familyName = ((fnEl && fnEl.value) || '').trim() || 'My Family';
    localStorage.setItem('prizm_family_name', S.familyName);
    apiPost({ action: 'updateSettings', key: 'familyName', value: S.familyName });
    const cv = document.getElementById('family-code-val'); if (cv) cv.textContent = S.familyCode;
    toast('✅ Sheet created & connected');
    showSetup('sp-family-code');
  } else {
    startPolling();
    toast('✅ Sheet created & connected');
    S.onboardingNext = 'enterApp';
    showSetup('sp-accounts');
  }
}

// ── Firebase data layer — REMOVED ───────────────────────────────────────────
// All cloud storage now lives in the user's own Google Sheet (S.mode==='sheets')
// or on-device (S.mode==='local'). These are inert stubs kept only so any legacy
// S.mode==='firebase' guard (which can no longer be true) stays reference-safe.
async function firebaseRead(){ return null; }
async function firebaseWrite(){ return false; }
async function fbWriteAll(){ /* no-op — Firebase removed */ }
async function signOutFirebase(){ /* no-op — Firebase removed */ }
function toggleCompact() {
  const on = document.body.hasAttribute('data-compact');
  if (on) { document.body.removeAttribute('data-compact'); localStorage.removeItem('prizm_compact'); }
  else { document.body.setAttribute('data-compact','1'); localStorage.setItem('prizm_compact','1'); }
  const sub = document.getElementById('compact-mode-sub');
  if (sub) sub.textContent = on ? 'Off' : 'On';
  const tog = document.getElementById('compact-mode-toggle');
  if (tog) tog.classList.toggle('on', !on);
  schedulePrefSync(); // #2
}
function setCorner(mode, el) {
  document.querySelectorAll('.corner-opt').forEach(b => b.classList.remove('active'));
  if (el) el.classList.add('active');
  localStorage.setItem('prizm_corner', mode);
  if (el) schedulePrefSync(); // #2 · only when user-initiated (el set), not on boot/incoming
  const root = document.documentElement;
  if (mode === 'sharp') {
    root.style.setProperty('--r-xs','2px'); root.style.setProperty('--r-sm','4px');
    root.style.setProperty('--r-md','6px'); root.style.setProperty('--r-lg','8px');
    root.style.setProperty('--r-xl','10px');
  } else if (mode === 'pill') {
    root.style.setProperty('--r-xs','12px'); root.style.setProperty('--r-sm','18px');
    root.style.setProperty('--r-md','24px'); root.style.setProperty('--r-lg','32px');
    root.style.setProperty('--r-xl','44px');
  } else {
    root.style.setProperty('--r-xs','6px'); root.style.setProperty('--r-sm','10px');
    root.style.setProperty('--r-md','14px'); root.style.setProperty('--r-lg','20px');
    root.style.setProperty('--r-xl','26px');
  }
}
function toggleEmoji() {
  const emojiOn = !document.body.hasAttribute('data-noemoji');
  if (emojiOn) { document.body.setAttribute('data-noemoji','1'); localStorage.setItem('prizm_noemoji','1'); }
  else { document.body.removeAttribute('data-noemoji'); localStorage.removeItem('prizm_noemoji'); }
  renderHome(); renderSettings();
  schedulePrefSync(); // #2
}
function applyEmojiPref() {
  // Emoji is default; Lucide is opt-in
  if (localStorage.getItem('prizm_noemoji') === '1') document.body.setAttribute('data-noemoji','1');
  else document.body.removeAttribute('data-noemoji');
}

function updateMemberPicker(rowId, selId, selectedId) {
  const row = document.getElementById(rowId);
  const sel = document.getElementById(selId);
  if (!row || !sel) return;
  const show = S.family && S.members && S.members.length > 0;
  row.style.display = show ? '' : 'none';
  if (!show) return;
  const selfId = S.memberId || (S.member && S.member.id) || (S.members[0] && S.members[0].id) || '';
  const current = (selectedId !== undefined && selectedId !== '' && selectedId !== null) ? selectedId : selfId;
  sel.innerHTML = S.members.map(m =>
    `<option value="${esc(m.id)}" ${m.id === current ? 'selected' : ''}>${esc(m.name)}</option>`
  ).join('') +
  `<option value="${JOINT_ID}" ${current === JOINT_ID ? 'selected' : ''}>👥 ${esc(JOINT_LABEL)}</option>`;
}

function setTrackIncome(on) {
  S.trackIncome = on;
  localStorage.setItem('prizm_track_income', on ? '1' : '0');
  applyTrackIncome();
  renderSettings();
  renderAll(); // refresh hero (spent-only vs balance) and the goals card
}
function applyTrackIncome() {
  const incBtn = document.getElementById('tb-inc');
  if (incBtn) incBtn.style.display = S.trackIncome ? '' : 'none';
  const ttp = document.getElementById('ttp');
  if (ttp) ttp.style.display = S.trackIncome ? '' : 'none';
  const rIncBtn = document.getElementById('rtb-inc');
  if (rIncBtn) rIncBtn.style.display = S.trackIncome ? '' : 'none';
  const rttp = document.getElementById('rttp');
  if (rttp) rttp.style.display = S.trackIncome ? '' : 'none';
  if (!S.trackIncome && S.type === 'income') setType('expense');
  const incStat = document.getElementById('hero-income-stat');
  if (incStat) incStat.style.display = S.trackIncome ? '' : 'none';
  const heroRow = document.getElementById('home-hero')?.querySelector?.('.hero-row');
  if (heroRow) { heroRow.style.display = 'grid'; heroRow.classList.toggle('single-stat', !S.trackIncome); }
  const incStatBox = document.getElementById('a-income')?.closest?.('.stat-box');
  if (incStatBox) incStatBox.style.display = S.trackIncome ? '' : 'none';
  const saveRateBox = document.getElementById('a-save-rate')?.closest?.('.stat-box');
  if (saveRateBox) saveRateBox.style.display = S.trackIncome ? '' : 'none';
  const incBreakTab = document.querySelector('.an-sub-tab[onclick*="income"]');
  if (incBreakTab) incBreakTab.style.display = S.trackIncome ? '' : 'none';
  const donutTitle = document.getElementById('donut-chart-title');
  if (donutTitle) donutTitle.textContent = S.trackIncome ? 'Income vs expenses' : 'Spending by category';
  const aNet = document.getElementById('a-net');
  if (aNet) aNet.style.display = S.trackIncome ? '' : 'none';
  const fcInc = document.getElementById('fc-income');
  if (fcInc) fcInc.style.display = S.trackIncome ? '' : 'none';
  const trackSub = document.getElementById('track-income-sub');
  if (trackSub) trackSub.textContent = S.trackIncome ? 'On' : 'Off';
  try { renderFilterChips(); } catch(e) {} // add/remove the Income chip
}

function setSimpleMode(on) {
  S.simpleMode = on;
  localStorage.setItem('prizm_simple', on ? '1' : '');
  applySimpleMode();
  renderSettings();
  schedulePrefSync(); // #2
}
function applySimpleMode() {
  const shell = document.getElementById('shell');
  const hide = id => { const el=document.getElementById(id); if(el) el.style.display='none'; };
  const show = id => { const el=document.getElementById(id); if(el) el.style.display=''; };
  const hero = document.getElementById('home-hero');
  const simpleHome = document.getElementById('simple-home');
  const fullHome = document.getElementById('full-home');
  const banner = document.getElementById('simple-exit-banner');
  if (S.simpleMode) {
    hide('nb-transactions'); hide('nb-analytics');
    if (shell) shell.setAttribute('data-simple','1');
    if (hero) hero.style.display = 'none';
    if (simpleHome) simpleHome.style.display = 'flex';
    if (fullHome) fullHome.style.display = 'none';
    if (banner) banner.style.display = 'flex';
    if (['page-transactions','page-analytics'].some(id => document.getElementById(id)?.classList.contains('active'))) goTab('home');
    renderSimpleHome();
  } else {
    show('nb-transactions'); show('nb-analytics');
    if (shell) shell.removeAttribute('data-simple');
    if (hero) hero.style.display = '';
    if (simpleHome) simpleHome.style.display = 'none';
    if (fullHome) fullHome.style.display = '';
    if (banner) banner.style.display = 'none';
  }
}

function renderSimpleHome() {
  if (!S.simpleMode) return;
  const now = new Date();
  const mStart = new Date(now.getFullYear(), now.getMonth(), 1).getTime();
  const mTxs = S.txs.filter(t => t.type==='expense' && dateMs(t.date) >= mStart);
  const mExp = mTxs.reduce((sum,t) => sum+parseFloat(t.amount||0), 0);
  const monthLbl = document.getElementById('simple-month-lbl');
  if (monthLbl) monthLbl.textContent = 'Spent · ' + now.toLocaleString('default',{month:'long'});
  const spentEl = document.getElementById('simple-spent-amt');
  if (spentEl) { spentEl.textContent = fmtMoney(mExp); fitHeroAmount(spentEl, 30, 20); }
  const lblEl = document.getElementById('simple-spent-lbl');
  if (lblEl) lblEl.textContent = mTxs.length
    ? (mTxs.length + (mTxs.length===1 ? ' expense' : ' expenses') + ' this month')
    : 'Nothing spent yet';
  const sod = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const wStart = sod - 6*86400000;
  let tToday = 0, tWeek = 0;
  S.txs.forEach(t => {
    if (t.type !== 'expense') return;
    const ts = dateMs(t.date), a = parseFloat(t.amount||0);
    if (ts >= sod) tToday += a;
    if (ts >= wStart) tWeek += a;
  });
  const tEl = document.getElementById('simple-today-amt'); if (tEl) tEl.textContent = fmtMoneyCompact(tToday);
  const wEl = document.getElementById('simple-week-amt'); if (wEl) wEl.textContent = fmtMoneyCompact(tWeek);
  const recent = [...S.txs].filter(t=>t.type==='expense').slice(0,5);
  const cEl = document.getElementById('simple-recent-count');
  if (cEl) cEl.textContent = recent.length ? 'last ' + recent.length : '';
  const recentEl = document.getElementById('simple-recent-txs');
  if (recentEl) {
    recentEl.innerHTML = recent.length
      ? recent.map(t => txHTML(t)).join('')
      : '<div class="empty" style="padding:26px 12px"><div class="empty-title">No expenses yet</div><div class="empty-sub">Tap the button above to add one.</div></div>';
  }
}

function loadLocalData() {
  const arr = k => { const v = safeParse(localStorage.getItem(k), []); return Array.isArray(v) ? v : []; };
  S.txs = arr('prizm_local_txs');
  sortTxs(S.txs); // newest-first, entry time breaks same-day ties
  S.accts = arr('prizm_local_accts');
  S.members = arr('prizm_local_members');
  S.recurring = arr('prizm_local_recurring');
  S.budgets = arr('prizm_local_budgets');
}

// ~5MB localStorage budget. Warn proactively near the ceiling, warn loudly (and
// repeatedly, throttled) when a write actually fails, and never let one failed
// setItem abort the remaining keys.
const LS_BUDGET = 5 * 1024 * 1024;
let _lastQuotaToast = 0;   // re-armed: warns at most once/min, not once per page load
let _nearQuotaWarned = false;
function _quotaToast(msg) {
  const now = Date.now();
  if (now - _lastQuotaToast < 60000) return;
  _lastQuotaToast = now;
  toast(msg);
}
function saveLocalData() {
  const pairs = [
    ['prizm_local_txs', S.txs],
    ['prizm_local_accts', S.accts],
    ['prizm_local_members', S.members],
    ['prizm_local_recurring', S.recurring],
    ['prizm_local_budgets', S.budgets]
  ];
  let failed = false, totalBytes = 0;
  pairs.forEach(([key, val]) => {
    // Each key in its own try: one quota failure must not silently skip the rest.
    try {
      const json = JSON.stringify(val);
      totalBytes += json.length + key.length;
      localStorage.setItem(key, json);
    } catch(e) {
      failed = true;
      console.warn('saveLocalData (' + key + '):', e);
    }
  });
  if (failed) {
    _quotaToast('⚠️ Device storage full — recent changes are NOT saved. Export your data now (Settings → Export CSV).');
  } else if (totalBytes > LS_BUDGET * 0.8 && !_nearQuotaWarned) {
    _nearQuotaWarned = true;
    toast('⚠️ Local storage is over 80% full · consider exporting/archiving old data');
  }
}

// ── Full data backup & restore (JSON) ───────────────────────────────────────
// A safety net independent of Google Sheets. Captures everything Lucid holds so a
// user can never permanently lose their history.
function _backupPayload() {
  return {
    format: 'prizm-backup',
    version: (typeof APP_VERSION !== 'undefined' ? APP_VERSION : '1'),
    exportedAt: new Date().toISOString(),
    transactions: S.txs || [],
    accounts: S.accts || [],
    members: S.members || [],
    recurring: S.recurring || [],
    budgets: S.budgets || [],
    categories: { exp: getCatList(false), inc: getCatList(true) },
    goals: safeParse(localStorage.getItem('prizm_goals'), []),
    settings: { name: S.name, currency: S.cur && S.cur.code, familyName: S.familyName || '', prefs: gatherPrefs() }
  };
}
function exportBackup() {
  try {
    const blob = new Blob([JSON.stringify(_backupPayload(), null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'prizm-backup-' + todayISO() + '.json';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
    toast('✅ Backup downloaded');
  } catch (e) { console.warn('exportBackup:', e); toast('⚠️ Could not create backup'); }
}
function importBackup(input) {
  const file = input && input.files && input.files[0];
  input.value = ''; // allow re-selecting the same file later
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function (e) {
    let data;
    try { data = JSON.parse(e.target.result); } catch (err) { toast('⚠️ That file isn\'t a valid backup'); return; }
    if (!data || data.format !== 'prizm-backup' || !Array.isArray(data.transactions)) { toast('⚠️ Not a Lucid backup file'); return; }
    showCustomConfirm('Restore this backup?',
      'This replaces the data on THIS device with the backup (' + (data.transactions.length) + ' transactions, saved ' +
      (data.exportedAt ? new Date(data.exportedAt).toLocaleDateString() : 'unknown date') + '). Your current data on this device will be overwritten.',
      'Restore',
      function () {
        try {
          S.txs = sortTxs(data.transactions || []);
          S.accts = data.accounts || [];
          S.members = data.members || [];
          S.recurring = data.recurring || [];
          S.budgets = data.budgets || [];
          if (data.categories) {
            if (Array.isArray(data.categories.exp)) _saveCatList(false, data.categories.exp);
            if (Array.isArray(data.categories.inc)) _saveCatList(true, data.categories.inc);
          }
          if (Array.isArray(data.goals)) localStorage.setItem('prizm_goals', JSON.stringify(data.goals));
          if (data.settings) {
            if (data.settings.currency) { const c = CURRENCIES.find(x => x.code === data.settings.currency); if (c) { S.cur = c; localStorage.setItem('prizm_cur', c.code); } }
            if (data.settings.name) { S.name = data.settings.name; localStorage.setItem('prizm_name', S.name); }
            if (data.settings.prefs) applyIncomingPrefs(data.settings.prefs);
          }
          saveLocalData();
          buildCatGrid(); buildRecCatGrid(); buildBudgetCatOptions();
          renderAll();
          if (S.mode === 'sheets') fbWriteAll(); // no-op unless firebase; sheets keeps device copy
          toast('✅ Backup restored');
        } catch (err) { console.warn('restore:', err); toast('⚠️ Restore failed'); }
      });
  };
  reader.readAsText(file);
}

let _lastGetAllText = null; // raw-text change detection · skips parse+apply on unchanged polls
async function apiGet(skipIfUnchanged) {
  if (!S.url) return null;
  try {
    const response = await fetch(S.url + '?action=getAll&_ts=' + Date.now() + authK(), { method: 'GET' });
    if (!response.ok) throw new Error('Network un-optimized');
    const txt = await response.text();
    const data = JSON.parse(txt);
    if (data && data.error === 'unauthorized') { _handleUnauthorized(); return null; }
    if (skipIfUnchanged && txt === _lastGetAllText) return null; // unchanged → no-op
    _lastGetAllText = txt;
    return data;
  } catch(e) { return null; }
}
// Password helpers (shared-secret lock). Kept simple: prompt, store, retry.
let _pwPrompting = false;
function _handleUnauthorized(){
  if (_pwPrompting) return; _pwPrompting = true;
  const p = window.prompt('This Lucid ledger is password-protected.\nEnter the password to continue:');
  _pwPrompting = false;
  if (p === null) return;
  try { localStorage.setItem('prizm_secret', p.trim()); } catch(e) {}
  toast('Checking password…');
  setTimeout(() => { try { syncData(); outboxFlush(); } catch(e) {} }, 300);
}
function setLucidPassword(){
  if (S.mode !== 'sheets') { toast('Connect a Google Sheet first'); return; }
  const current = _prizmSecret();
  let np = window.prompt(current
    ? 'Enter a NEW password (leave blank to remove the password):'
    : 'Set a password for this ledger.\nEveryone sharing this data will need it:');
  if (np === null) return;
  np = np.trim();
  apiPost({ action:'setSecret', secret: np, current: current }); // offline-safe via outbox
  try { localStorage.setItem('prizm_secret', np); } catch(e) {}
  toast(np ? '🔒 Password set — share it with your family' : 'Password removed');
}

async function apiPost(body) {
  // 1. Optimistic update: apply to local cache immediately
  handleLocalPost(body);
  renderAll();

  if (S.mode === 'local') return {success:true};
  if (S.mode==='firebase') { fbWriteAll().catch(e=>console.warn(e)); return {success:true}; }
  // 2. Durable outbox: persist the write, then try to flush it. If we're offline (or
  //    the request fails) the write STAYS queued in localStorage and is replayed, in
  //    order, on reconnect — so a write made offline is never lost. (Previously this
  //    was a fire-and-forget fetch whose result vanished when offline.)
  outboxEnqueue(body);
  outboxFlush();
  return {success: true};
}

function applyBalanceDelta(accountName, amount, type, sign) {
  const acct = S.accts.find(a => a.name === accountName);
  if (!acct) return;
  const delta = (type === 'expense' ? -Number(amount) : Number(amount)) * sign;
  acct.balance = Number(acct.balance || 0) + delta;
}

function handleLocalPost(b) {
  if (b.action === 'addTransaction') {
    const id = b.id || 'tx_' + Date.now();
    S.txs.unshift({ id, date: b.date, description: b.description, amount: Number(b.amount), type: b.type, category: b.category, account: b.account || '—', notes: b.notes || '', memberId: b.memberId || '', ts: b.ts || Date.now() });
    applyBalanceDelta(b.account, b.amount, b.type, 1);
    sortTxs(S.txs);
    saveLocalData(); return {success:true, id};
  }
  if (b.action === 'updateTransaction') {
    const tx = S.txs.find(t => t.id === b.id);
    if (tx) {
      applyBalanceDelta(tx.account, tx.amount, tx.type, -1);
      tx.date = b.date; tx.description = b.description; tx.amount = Number(b.amount); tx.type = b.type;
      tx.category = b.category; tx.account = b.account || '—'; tx.notes = b.notes || '';
      if (b.memberId !== undefined) tx.memberId = b.memberId;
      if (b.ts) tx.ts = b.ts; // stamp edit recency (hidden from UI)
      applyBalanceDelta(tx.account, tx.amount, tx.type, 1);
      sortTxs(S.txs);
      saveLocalData();
    }
    return {success:true};
  }
  if (b.action === 'deleteTx') {
    const idx = S.txs.findIndex(t => t.id === b.id);
    if (idx !== -1) {
      const tx = S.txs[idx];
      applyBalanceDelta(tx.account, tx.amount, tx.type, -1);
      S.txs.splice(idx,1); saveLocalData();
    }
    return {success:true};
  }
  if (b.action === 'addAccount') { S.accts.push({ id: b.id || 'ac_'+Date.now(), name: b.name, type: b.type, balance: Number(b.balance)||0 }); saveLocalData(); return {success:true}; }
  if (b.action === 'updateAccount') {
    const a = S.accts.find(x => x.id === b.id);
    if (a) {
      const oldName = a.name;
      a.name = b.name; a.type = b.type; a.balance = Number(b.balance)||0;
      if (oldName !== a.name) S.txs.forEach(t => { if (t.account === oldName) t.account = a.name; });
      saveLocalData();
    }
    return {success:true};
  }
  if (b.action === 'deleteAccount') {
    // Cascade: clear the account off transactions (they display "—", as the delete
    // dialog promises). Without this, orphaned name-strings silently re-attach to
    // any future account created with the same name.
    const deleted = S.accts.find(a => a.id === b.id);
    if (deleted && deleted.name) S.txs.forEach(t => { if (t.account === deleted.name) t.account = '—'; });
    S.accts = S.accts.filter(a => a.id !== b.id);
    saveLocalData();
    return {success:true};
  }
  if (b.action === 'addMember') { const id = b.id || 'mem_' + Date.now(); S.members.push({id, name:b.name, pin:b.pin, color:b.color}); saveLocalData(); return {success:true}; }
  if (b.action === 'deleteMember') { S.members = S.members.filter(m => m.id !== b.id); saveLocalData(); return {success:true}; }
  if (b.action === 'addRecurring') { const id = b.id || 'rec_' + Date.now(); S.recurring.push({id, name:b.name, amount:Number(b.amount), type:b.type, category:b.category, account:b.account||'', frequency:b.frequency, nextDate:b.nextDate, targetDay: Number(b.targetDay) || Number(String(b.nextDate||'').slice(8,10)) || undefined, memberId:b.memberId||'', active:true}); saveLocalData(); return {success:true}; }
  if (b.action === 'updateRecurring') {
    const r = S.recurring.find(x => x.id === b.id);
    if (r) { Object.assign(r, {name:b.name, amount:Number(b.amount), type:b.type, category:b.category, account:b.account||'', frequency:b.frequency, nextDate:b.nextDate, targetDay: b.targetDay!==undefined ? (Number(b.targetDay)||undefined) : r.targetDay, memberId: b.memberId!==undefined?b.memberId:r.memberId, active: b.active!==undefined?b.active:r.active}); saveLocalData(); }
    return {success:true};
  }
  if (b.action === 'deleteRecurring') { S.recurring = S.recurring.filter(r => r.id !== b.id); saveLocalData(); return {success:true}; }
  if (b.action === 'setBudget') {
    const id = b.id || 'bg_' + Date.now();
    let bgt = S.budgets.find(x => x.id === id) || S.budgets.find(x => x.category === b.category);
    if (bgt) { bgt.category = b.category; bgt.limit = Number(b.limit)||0; }
    else S.budgets.push({id, category:b.category, limit:Number(b.limit)||0});
    saveLocalData(); return {success:true};
  }
  if (b.action === 'deleteBudget') { S.budgets = S.budgets.filter(x => x.id !== b.id); saveLocalData(); return {success:true}; }
  if (b.action === 'updateSettings') {
    if (b.key === 'name') { S.name = b.value; localStorage.setItem('prizm_name', b.value); }
    if (b.key === 'currency') { const c = CURRENCIES.find(x => x.code === b.value); if (c) { S.cur = c; localStorage.setItem('prizm_cur', c.code); } }
    if (b.key === 'familyName') { S.familyName = b.value; localStorage.setItem('prizm_family_name', b.value); }
    else if (SYNC_PREF_KEYS[b.key]) { if (b.value === '' || b.value == null) localStorage.removeItem(SYNC_PREF_KEYS[b.key]); else localStorage.setItem(SYNC_PREF_KEYS[b.key], b.value); }
    return {success:true};
  }
  return {success:false};
}

async function syncDataWithLoader() {
  if (S.mode==='local'){renderAll();toast('📁 Local data refreshed');return;}
  if (S.mode==='firebase'){loader(true);const d=await firebaseRead();loader(false);if(!d){toast('⚠️ Offline cache');renderAll();return;}applyRemoteData(d);buildCatGrid();buildRecCatGrid();buildBudgetCatOptions();renderAll();toast('🔥 Firebase refreshed');return;}
  loader(true); const d = await apiGet(); loader(false);

  if (!d || d.error) { toast('⚠️ Using offline cache'); renderAll(); return; }

  applyRemoteData(d);
  buildCatGrid(); buildRecCatGrid(); buildBudgetCatOptions();
  renderAll();

  const t = new Date().toLocaleTimeString([], {hour: '2-digit', minute: '2-digit'});
  const lbl = document.getElementById('last-sync-lbl'); if (lbl) lbl.textContent = 'Last synced ' + t;
  toast('✅ Ledger synchronized');
  checkRecurringDue();
}

// Background headless sync (no loader)
async function syncData() {
  if (S.mode==='local') return;
  if (S.mode==='firebase') { const d=await firebaseRead(); if(d && applyRemoteData(d)) renderAll(); return; }
  // Don't overwrite local state while a write is in-flight, queued, or just landed
  if (_pendingWrites > 0 || outboxCount() > 0 || Date.now() - _lastWriteComplete < _SETTLE_MS) return;
  const d = await apiGet(true); if (!d||d.error) return; if (applyRemoteData(d)) renderAll();
}

let _fbListener = null;
let _fbListenerPath = null;
let _pendingWrites = 0;       // count of in-flight sheet writes
let _lastWriteComplete = 0;   // timestamp of last completed write (ms)
const _SETTLE_MS = 4000;      // don't sync from remote for 4s after a write lands

// ══════════════════════════════════════════════════════════════════
//  DURABLE WRITE OUTBOX  (offline-safe writes)
//  Every Sheet write is saved to localStorage BEFORE it's sent. A write made
//  while offline stays queued and is replayed, in order, when the connection
//  returns — so nothing is ever lost. A non-empty outbox counts as "pending"
//  so a background getAll can't clobber changes that haven't synced yet.
// ══════════════════════════════════════════════════════════════════
// Shared-password token helpers. If the user has set a password, it's sent with every
// request: as ?k=… on reads, and folded into the body on writes (backend checks both).
function _prizmSecret(){ try { return localStorage.getItem('prizm_secret') || ''; } catch(e){ return ''; } }
function authK(){ const s = _prizmSecret(); return s ? '&k=' + encodeURIComponent(s) : ''; }
function authBody(obj){ const s = _prizmSecret(); return s ? Object.assign({}, obj, { k: s }) : obj; }

const _OUTBOX_KEY = 'prizm_outbox';
let _outboxFlushing = false;
function outboxLoad() { try { return JSON.parse(localStorage.getItem(_OUTBOX_KEY) || '[]'); } catch(e) { return []; } }
function outboxSave(arr) { try { localStorage.setItem(_OUTBOX_KEY, JSON.stringify(arr)); } catch(e) {} }
function outboxCount() { return outboxLoad().length; }
function outboxEnqueue(body) {
  const q = outboxLoad();
  q.push({ qid: 'w_' + Date.now() + '_' + Math.random().toString(36).slice(2,7), body: body, ts: Date.now(), tries: 0 });
  outboxSave(q);
  updateSyncBadge();
}
async function outboxFlush() {
  if (_outboxFlushing) return;
  if (S.mode !== 'sheets' || !S.url) return;                                   // only the Sheets backend queues
  if (typeof navigator !== 'undefined' && navigator.onLine === false) return;  // offline → retry later
  _outboxFlushing = true;
  try {
    let q = outboxLoad();
    while (q.length) {
      const item = q[0];
      try {
        const qs = '?body=' + encodeURIComponent(JSON.stringify(authBody(item.body))) + '&_ts=' + Date.now();
        const r = await fetch(S.url + qs);
        if (r && !r.ok && r.status >= 500) throw new Error('HTTP ' + r.status); // server error → retry
        // Reached the server → drop this item and continue with the next.
        q = outboxLoad(); q.shift(); outboxSave(q);
        _lastWriteComplete = Date.now();
      } catch (e) {
        // Network failure / offline — keep the write and stop so order is preserved.
        item.tries = (item.tries || 0) + 1;
        q = outboxLoad();
        if (item.tries > 10) { console.error('Outbox: dropping a write after repeated failures', item, e); q.shift(); outboxSave(q); continue; }
        if (q[0]) { q[0].tries = item.tries; outboxSave(q); }
        break;
      }
    }
  } finally {
    _outboxFlushing = false;
    updateSyncBadge();
  }
}
function updateSyncBadge() {
  try {
    const n = outboxCount();
    const el = document.getElementById('last-sync-lbl');
    if (el) el.textContent = n > 0 ? (n + ' change' + (n > 1 ? 's' : '') + ' waiting to sync…') : 'Refresh from cloud';
  } catch(e) {}
}
// Replay triggers: reconnect, initial load, and a periodic backstop.
window.addEventListener('online', () => { try { toast('Back online — syncing…'); } catch(e) {} outboxFlush(); });
window.addEventListener('load', () => { setTimeout(outboxFlush, 3000); });
setInterval(() => { if (outboxCount()) outboxFlush(); }, 20000);

function startPolling() {
  stopPolling();
  if (S.mode === 'sheets') {
    // Adaptive poll: 8s while things are changing, backing off to 60s when idle,
    // and paused entirely while the tab is hidden (a full getAll every 8s forever
    // is needless load once the dataset grows). Any detected change snaps back to 8s.
    const token = {};
    S.pollToken = token;
    const POLL_MIN = 8000, POLL_MAX = 60000;
    let delay = POLL_MIN;
    const loop = async () => {
      if (S.pollToken !== token) return;
      if (!document.hidden) {
        const before = S._lastRemoteSig;
        try { await syncData(); } catch(e) {}
        delay = (S._lastRemoteSig === before) ? Math.min(Math.round(delay * 1.5), POLL_MAX) : POLL_MIN;
      }
      if (S.pollToken !== token) return;
      S.pollHandle = setTimeout(loop, delay);
    };
    S.pollHandle = setTimeout(loop, POLL_MIN);
  }
}
function stopPolling() {
  S.pollToken = null;
  if (S.pollHandle) { clearTimeout(S.pollHandle); S.pollHandle = null; }
}

// ═══════════════════════════════════════════
//  RECURRING TRANSACTIONS · DUE REVIEW
// ═══════════════════════════════════════════
function advanceDate(dateStr, freq, targetDay) {
  // #5 · month/year steps clamp to the last valid day of the target month
  // (e.g. Jan 31 → Feb 28/29) so recurring dates never overflow. The clamp is
  // computed FRESH each cycle from the item's original target day (targetDay),
  // so a "31st monthly" item clamped to Feb 28 goes back to Mar 31 instead of
  // drifting to the 28th forever.
  const base = (dateStr ? String(dateStr).slice(0,10) : todayISO());
  const parts = base.split('-').map(Number);
  let Y = parts[0], M = parts[1] - 1, D = parts[2];
  if (isNaN(Y) || isNaN(M) || isNaN(D)) { const n = new Date(); Y = n.getFullYear(); M = n.getMonth(); D = n.getDate(); }
  if (freq === 'daily')  return isoLocal(new Date(Y, M, D + 1));
  if (freq === 'weekly') return isoLocal(new Date(Y, M, D + 7));
  if (freq === 'yearly') { Y += 1; }
  else { M += 1; if (M > 11) { M = 0; Y += 1; } } // monthly
  const want = Number(targetDay) || D;
  const lastDay = new Date(Y, M + 1, 0).getDate();
  return isoFromParts(Y, M + 1, Math.min(want, lastDay));
}

function checkRecurringDue() {
  const shell = document.getElementById('shell');
  if (!shell || shell.style.display === 'none') return;
  const openOv = document.querySelector('.overlay.open');
  if (openOv && openOv.id !== 'due-overlay') return; // never interrupt an open sheet
  const today = todayISO();
  S.dueQueue = (S.recurring || []).filter(r => r.active !== false && r.nextDate && String(r.nextDate).slice(0,10) <= today);
  if (!S.dueQueue.length) return;
  renderDueList();
  if (!document.getElementById('due-overlay').classList.contains('open')) openOverlay('due-overlay');
}

function renderDueList() {
  const wrap = document.getElementById('due-list');
  if (!S.dueQueue.length) { closeOverlay('due-overlay'); return; }
  wrap.innerHTML = S.dueQueue.map((r,i) => {
    const isInc = r.type === 'income';
    return `<div class="due-card" style="flex-direction:column;align-items:stretch;gap:12px">
      <div style="display:flex;align-items:center;gap:12px">
        <div class="tx-ico" style="background:${getCatColor(r.category)}">${getCatIcon(r.category, isInc)}</div>
        <div class="due-info"><div class="due-name">${esc(r.name)}</div><div class="due-meta">${esc(r.frequency)} · due ${esc(r.nextDate)}</div></div>
        <div class="due-amt ${isInc?'income':''}">${isInc?'+':'-'}${fmtMoney(r.amount)}</div>
      </div>
      <div class="due-actions" style="justify-content:flex-end">
        <div class="due-btn" onclick="skipDue(${i})">Skip</div>
        <div class="due-btn confirm" onclick="confirmDue(${i})">Confirm</div>
      </div>
    </div>`;
  }).join('');
}

async function confirmDue(i) {
  const r = S.dueQueue[i];
  if (!r) return;
  const payload = {action:'addTransaction', id:'tx_'+Date.now(), ts: Date.now(), amount:r.amount, description:r.name, type:r.type, category:r.category, account: r.account || (S.accts[0] && S.accts[0].name) || '', date: r.nextDate || todayISO(), notes:'Recurring: '+r.name, memberId: r.memberId || (S.member ? S.member.id : '')};
  await apiPost(payload);
  const tday = r.targetDay || Number(String(r.nextDate||'').slice(8,10)) || undefined;
  const nd = advanceDate(r.nextDate, r.frequency, tday);
  await apiPost({action:'updateRecurring', id:r.id, name:r.name, amount:r.amount, type:r.type, category:r.category, account:r.account, frequency:r.frequency, nextDate: nd, targetDay: tday, memberId:r.memberId, active:true});
  // If it's still overdue (item was several periods behind), keep it in the queue
  if (nd <= todayISO()) { r.nextDate = nd; } else { S.dueQueue.splice(i,1); }
  renderDueList();
  if (!S.dueQueue.length) closeOverlay('due-overlay');
  toast('✅ Recurring entry posted');
}

async function skipDue(i) {
  const r = S.dueQueue[i];
  if (!r) return;
  const tday = r.targetDay || Number(String(r.nextDate||'').slice(8,10)) || undefined;
  const nd = advanceDate(r.nextDate, r.frequency, tday);
  await apiPost({action:'updateRecurring', id:r.id, name:r.name, amount:r.amount, type:r.type, category:r.category, account:r.account, frequency:r.frequency, nextDate: nd, targetDay: tday, memberId:r.memberId, active:true});
  if (nd <= todayISO()) { r.nextDate = nd; } else { S.dueQueue.splice(i,1); }
  renderDueList();
  if (!S.dueQueue.length) closeOverlay('due-overlay');
}

// ═══════════════════════════════════════════
//  SHARED HELPERS
// ═══════════════════════════════════════════
const _fmtCache = {};
// Shrink a big amount's font until it fits its box, so millions don't overflow
// the hero card. Keeps small values at full size.
function fitHeroAmount(el, maxPx, minPx) {
  if (!el) return;
  maxPx = maxPx || 44; minPx = minPx || 22;
  el.style.whiteSpace = 'nowrap';
  let size = maxPx, guard = 0;
  el.style.fontSize = size + 'px';
  while (el.scrollWidth > el.clientWidth && size > minPx && guard < 48) { size--; el.style.fontSize = size + 'px'; guard++; }
}
function fmtMoney(amount) {
  const k = S.cur.code;
  if (!_fmtCache[k]) { try { _fmtCache[k] = new Intl.NumberFormat('en-US', { style:'currency', currency:k }); } catch(e) { return (S.cur.sym||'$') + (Number(amount)||0).toFixed(2); } }
  return _fmtCache[k].format(amount || 0);
}
// Compact money for tight spaces (stat pills, chips, bars) so gigantic values never
// overflow/overlap. Full precision is still used in transaction rows. Above ~1e6 it
// renders like $1.2M / $100T; below that it's the normal formatted amount.
const _fmtCompactCache = {};
function fmtMoneyCompact(amount) {
  const n = Number(amount) || 0;
  if (Math.abs(n) < 1e6) return fmtMoney(n);
  const k = S.cur.code;
  try {
    if (!_fmtCompactCache[k]) _fmtCompactCache[k] = new Intl.NumberFormat('en-US', { style:'currency', currency:k, notation:'compact', maximumFractionDigits:1 });
    return _fmtCompactCache[k].format(n);
  } catch(e) { return fmtMoney(n); }
}
function updateGreeting() {
  const hr = new Date().getHours();
  const base = hr < 5 ? 'Good night' : hr < 12 ? 'Good morning' : hr < 18 ? 'Good afternoon' : 'Good evening';
  const who = (S.member && S.member.name) ? S.member.name : (S.name || 'there');
  const first = (who || 'there').trim().split(/\s+/)[0] || 'there';
  const displayName = (first === 'Lucid' || first === 'User') ? 'there' : first;
  const lbl = document.getElementById('home-greeting-label');
  const nm = document.getElementById('home-greeting-name');
  if (lbl) lbl.textContent = base + ',';
  if (nm) nm.textContent = displayName;
  const el = document.getElementById('home-greeting');
  if (el) el.textContent = base + ', ' + displayName;
}
function getCatIcon(name, isIncome) {
  if (document.body.hasAttribute('data-noemoji')) {
    // Minimal icon mode: theme-tinted initial instead of an emoji
    const ch = String(name || '?').charAt(0).toUpperCase();
    return '<span style="font-family:var(--head-font);font-weight:800;color:var(--accent)">' + esc(ch) + '</span>';
  }
  const customEmojis = safeParse(localStorage.getItem(isIncome ? 'prizm_inc_emojis' : 'prizm_cat_emojis'), {});
  if (customEmojis[name]) return customEmojis[name];
  let f = (isIncome ? INC_CATS : CATS).find(c => c.n === name);
  if (f) return f.e;
  // #22 · match against the user-managed category list, then fall back fuzzily
  try { f = getCatList(isIncome).find(c => c.n === name); if (f && f.e) return f.e; } catch(e) {}
  return fuzzyCatIcon(name, isIncome);
}
// #22 · keyword + deterministic fallback so custom / CSV-imported categories
// get a sensible (and visually distinct) icon instead of all sharing 💡 / ✨.
const FUZZY_ICON = [
  [/grocer|food|restaurant|dining|eat|meal|lunch|dinner|takeaway|snack/, '🍔'],
  [/transport|car|fuel|gas|petrol|uber|taxi|train|bus|metro|parking|flight/, '🚗'],
  [/shop|store|amazon|cloth|retail|mall/, '🛍'],
  [/health|pharm|medic|doctor|dental|hospital|clinic/, '💊'],
  [/entertain|movie|cinema|game|music|netflix|spotify|stream/, '🎬'],
  [/bill|util|electric|water|rent|mortgage|insur|phone|internet|subscription/, '📋'],
  [/educ|school|book|course|tuition|class|study/, '📚'],
  [/coffee|cafe|starbucks|tea/, '☕'],
  [/fit|gym|sport|workout|yoga/, '💪'],
  [/home|house|furnit|rent|repair/, '🏠'],
  [/travel|hotel|holiday|trip|vacation/, '✈️'],
  [/salary|payroll|wage|paycheck/, '💼'],
  [/freelance|contract|gig/, '💻'],
  [/business|biz|company/, '🏢'],
  [/invest|divid|stock|interest|crypto/, '📈'],
  [/gift|bonus|refund|cashback/, '🎁']
];
function fuzzyCatIcon(name, isIncome) {
  const n = String(name || '').toLowerCase();
  for (let i = 0; i < FUZZY_ICON.length; i++) { if (FUZZY_ICON[i][0].test(n)) return FUZZY_ICON[i][1]; }
  const pool = isIncome ? ['✨','💵','🪙','💹','🎁'] : ['🏷️','📦','🧾','🔖','🗂️','🧩','🎯'];
  let h = 0; for (let i = 0; i < n.length; i++) h = (h * 31 + n.charCodeAt(i)) >>> 0;
  return pool[n ? (h % pool.length) : 0];
}
// Icon tint now follows the active accent/theme for a cohesive, theme-aware look.
function getCatColor(name) { return 'var(--accent-soft)'; }

function openOverlay(id) {
  const el = document.getElementById(id);
  el.classList.add('open');
  document.body.classList.add('modal-open');
  _syncOverlayToKeyboard();
}
function closeOverlay(id) {
  if (id === 'add-overlay') {
    if (!S._skipDraft) { try { _captureAddDraft(); } catch(e) {} }
    S._skipDraft = false;
  }
  const el = document.getElementById(id);
  el.classList.remove('open');
  el.style.bottom = '';
  // Only remove modal-open if no other overlays are open
  if (!document.querySelector('.overlay.open')) {
    document.body.classList.remove('modal-open');
  }
}
function closeBg(e, id) { if (e.target === document.getElementById(id)) closeOverlay(id); }
function loader(on) { document.getElementById('loader').style.display = on ? 'flex' : 'none'; }

function showCustomConfirm(t, b, c, fn) {
  document.getElementById('dialog-title').textContent = t; document.getElementById('dialog-body').textContent = b;
  const cancel = document.getElementById('dialog-cancel-btn'); if (cancel) cancel.style.display = '';
  const btn = document.getElementById('dialog-action-btn'); btn.textContent = c;
  btn.className = 'dialog-btn dialog-btn-danger';
  btn.onclick = () => { fn(); closeCustomDialog(); };
  document.getElementById('custom-dialog-overlay').classList.add('open');
}
function closeCustomDialog() { document.getElementById('custom-dialog-overlay').classList.remove('open'); }

function showCustomAlert(title, body) {
  document.getElementById('dialog-title').textContent = title;
  document.getElementById('dialog-body').textContent = body;
  const c = document.getElementById('dialog-cancel-btn');
  if (c) c.style.display = 'none';
  const btn = document.getElementById('dialog-action-btn');
  btn.textContent = 'OK';
  btn.className = 'dialog-btn dialog-btn-cancel';
  btn.onclick = () => closeCustomDialog();
  document.getElementById('custom-dialog-overlay').classList.add('open');
}

function copyScript() {
  navigator.clipboard.writeText(SCRIPT_CODE).then(() => toast('📋 Copied to Clipboard')).catch(() => {
    const t = document.createElement('textarea'); t.value = SCRIPT_CODE; document.body.appendChild(t);
    t.select(); document.execCommand('copy'); document.body.removeChild(t); toast('📋 Copied to Clipboard');
  });
}

// ═══════════════════════════════════════════
//  ONBOARDING: "Need help with Deploy?" visual guide
// ═══════════════════════════════════════════
const DEPLOY_GUIDE_HTML = `
<svg viewBox="0 0 320 380" font-family="Inter, -apple-system, sans-serif">
  <!-- New deployment dialog -->
  <rect x="4" y="4" width="312" height="232" rx="14" fill="var(--bg-3)" stroke="var(--line-2)"/>
  <text x="20" y="30" font-size="13" font-weight="700" fill="var(--text-1)">New deployment</text>
  <circle cx="295" cy="22" r="9" fill="none" stroke="var(--line-3)"/>
  <text x="295" y="26" font-size="10" fill="var(--text-3)" text-anchor="middle">✕</text>

  <text x="20" y="55" font-size="10" letter-spacing="1" fill="var(--text-3)">SELECT TYPE</text>
  <rect x="20" y="62" width="140" height="32" rx="8" fill="var(--bg-4)" stroke="var(--line-2)"/>
  <text x="32" y="83" font-size="12" font-weight="600" fill="var(--text-1)">⚙️ Web app</text>

  <text x="20" y="115" font-size="10" letter-spacing="1" fill="var(--text-3)">EXECUTE AS</text>
  <rect x="20" y="122" width="180" height="32" rx="8" fill="var(--bg-4)" stroke="var(--line-2)"/>
  <text x="32" y="143" font-size="12" font-weight="600" fill="var(--text-1)">👤 Me</text>

  <text x="20" y="175" font-size="10" letter-spacing="1" fill="var(--text-3)">WHO HAS ACCESS</text>
  <rect x="20" y="182" width="140" height="34" rx="8" fill="var(--accent-soft)" stroke="var(--accent)" stroke-width="2"/>
  <text x="32" y="204" font-size="12" font-weight="700" fill="var(--text-1)">🌐 Anyone</text>
  <text x="172" y="204" font-size="11" font-weight="700" fill="var(--accent)">← set this</text>

  <rect x="208" y="226" width="92" height="34" rx="17" fill="var(--accent)"/>
  <text x="254" y="248" font-size="13" font-weight="800" fill="var(--on-accent)" text-anchor="middle">Deploy</text>

  <!-- Result card -->
  <rect x="4" y="252" width="312" height="120" rx="14" fill="var(--bg-3)" stroke="var(--line-2)"/>
  <text x="20" y="280" font-size="13" font-weight="700" fill="var(--green)">Deployment created ✓</text>
  <text x="20" y="305" font-size="10" letter-spacing="1" fill="var(--text-3)">WEB APP URL</text>
  <rect x="20" y="312" width="200" height="32" rx="8" fill="var(--bg-4)" stroke="var(--line-2)"/>
  <text x="30" y="332" font-size="10" font-family="monospace" fill="var(--text-2)">.../macros/s/AKfycb.../exec</text>
  <rect x="230" y="312" width="66" height="32" rx="8" fill="var(--accent)"/>
  <text x="263" y="332" font-size="11" font-weight="700" fill="var(--on-accent)" text-anchor="middle">📋 Copy</text>
  <text x="20" y="362" font-size="11" fill="var(--text-3)">← copy this link, then paste it into Lucid</text>
</svg>
<div class="deploy-guide-note">
  The only setting you need to change is <strong>Who has access</strong> · set it to <strong>Anyone</strong>. Leave everything else as-is, then tap <strong>Deploy</strong>.<br><br>
  The first time, Google may ask you to review permissions · that's normal. Tap <strong>Authorize access</strong>, choose your account, then continue.<br><br>
  Finally, copy the web app link Google gives you and paste it below.<br><br>
  Script missing or not working? <a onclick="copyScript()">Tap to re-copy the backend code</a>, paste it into the Apps Script editor (replacing everything there), and save.
</div>`;

function toggleDeployGuide(id) {
  const panel = document.getElementById(id);
  const btn = document.getElementById('toggle-' + id);
  if (!panel.dataset.loaded) {
    panel.querySelector('.deploy-guide-inner').innerHTML = DEPLOY_GUIDE_HTML;
    panel.dataset.loaded = '1';
  }
  const open = panel.classList.toggle('open');
  btn.classList.toggle('open', open);
}

function doExport() {
  if (!S.txs || S.txs.length === 0) return toast('⚠️ No data to export');
  const q = v => '"' + String(v == null ? '' : v).replace(/"/g,'""') + '"';
  let csv = 'Date,Type,Category,Description,Account,Amount,Notes,Member\n';
  S.txs.forEach(t => { const m = S.members.find(x=>x.id===t.memberId); csv += [q(t.date),q(t.type),q(t.category),q(t.description),q(t.account),Number(t.amount)||0,q(t.notes||''),q(m?m.name:'')].join(',') + '\n'; });
  const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
  a.download = `prizm_export_${todayISO()}.csv`; a.click(); toast('📥 Export complete');
}
let _csvParsed=null,_csvHeaders=[],_csvColumnMap={date:'',amount:'',description:'',category:'',type:'',notes:''};
function openImportCSV(){var ov=document.getElementById('csv-import-overlay');if(ov){ov.style.display='flex';var s1=document.getElementById('csv-step-1');var s2=document.getElementById('csv-step-2');if(s1)s1.style.display='';if(s2)s2.style.display='none';var fi=document.getElementById('csv-file-input');if(fi)fi.value='';}}
function closeImportCSV(){var ov=document.getElementById('csv-import-overlay');if(ov)ov.style.display='none';}
function csvBack(){document.getElementById('csv-step-1').style.display='';document.getElementById('csv-step-2').style.display='none';}
function handleCSVFile(input){var file=input.files[0];if(!file)return;var reader=new FileReader();reader.onload=function(e){try{var text=e.target.result;var lines=text.split('\n').map(function(l){return l.replace(/\r$/,'');}).filter(function(l){return l.trim();});if(lines.length<2){toast('CSV must have at least 2 rows');return;}function parseRow(row){var cols=[],cur='',inQ=false;for(var i=0;i<row.length;i++){var c=row[i];if(c==='"'){if(inQ&&row[i+1]==='"'){cur+='"';i++;}else inQ=!inQ;}else if(c===','&&!inQ){cols.push(cur.trim());cur='';}else cur+=c;}cols.push(cur.trim());return cols;}_csvHeaders=parseRow(lines[0]);_csvParsed=lines.slice(1).map(parseRow);function autoMap(kws){var k=_csvHeaders.findIndex(function(h){return kws.some(function(kw){return h.toLowerCase().indexOf(kw)>=0;});});return k>=0?k:'';}_csvColumnMap={date:autoMap(['date','time','when']),amount:autoMap(['amount','sum','total','debit','credit','value']),description:autoMap(['description','desc','merchant','name','memo','narrative']),category:autoMap(['category','cat','tag']),type:autoMap(['type','direction']),notes:autoMap(['note','memo','comment'])};document.getElementById('csv-step-1').style.display='none';document.getElementById('csv-step-2').style.display='';renderCSVMapping();renderCSVPreview();}catch(err){toast('Could not parse CSV');}};reader.readAsText(file);}
function renderCSVMapping(){
  var acctRow=document.getElementById('csv-acct-row'),acctSel=document.getElementById('csv-acct-select');
  if(acctRow&&acctSel){acctRow.style.display=S.accts&&S.accts.length?'flex':'none';if(S.accts&&S.accts.length){acctSel.innerHTML=S.accts.map(function(a){return'<option value="'+esc(a.name)+'">'+esc(a.name)+'</option>';}).join('');}}
  var memRow=document.getElementById('csv-member-row'),memSel=document.getElementById('csv-member-select');
  if(memRow&&memSel){var showM=S.family&&S.members&&S.members.length>0;memRow.style.display=showM?'flex':'none';if(showM){var selfId=S.memberId||(S.members[0]&&S.members[0].id)||'';memSel.innerHTML=S.members.map(function(m){return'<option value="'+esc(m.id)+'"'+(m.id===selfId?' selected':'')+'>'+esc(m.name)+'</option>';}).join('')+'<option value="'+JOINT_ID+'">👥 '+esc(JOINT_LABEL)+'</option>';}}
  var wrap=document.getElementById('csv-column-mapping');if(!wrap)return;var fields=[{key:'date',label:'Date *'},{key:'amount',label:'Amount *'},{key:'description',label:'Description'},{key:'category',label:'Category'},{key:'type',label:'Type (income/expense)'},{key:'notes',label:'Notes'}];wrap.innerHTML=fields.map(function(f){var sel='<select onchange=\"var v=this.value;_csvColumnMap[\''+f.key+'\']=(v===\'__none__\'?\'\':(parseInt(v)));renderCSVPreview()\" style=\"flex:1;background:var(--bg-3);border:1px solid var(--line-2);border-radius:var(--r-sm);padding:7px 10px;color:var(--text-1);font-size:13px\"><option value=\"__none__\">-- skip --</option>'+_csvHeaders.map(function(hh,i){return '<option value=\"'+i+'\"'+(_csvColumnMap[f.key]===i?' selected':'')+'>'+esc(hh)+'</option>';}).join('')+'</select>';return '<div style=\"display:flex;align-items:center;gap:10px\"><div style=\"width:110px;font-size:13px;font-weight:500;flex-shrink:0\">'+f.label+'</div>'+sel+'</div>';}).join('');}
function renderCSVPreview(){var wrap=document.getElementById('csv-preview');if(!wrap||!_csvParsed)return;var preview=_csvParsed.slice(0,5);function get(row,key){var idx=_csvColumnMap[key];return esc((idx!==''&&idx!==undefined&&idx!==null)?(row[idx]||'--'):'--');}wrap.innerHTML='<table style="width:100%;border-collapse:collapse;font-size:11px"><tr style="color:var(--text-3);border-bottom:1px solid var(--line-2)"><th style="text-align:left;padding:4px 6px">Date</th><th style="text-align:left;padding:4px 6px">Desc</th><th style="text-align:right;padding:4px 6px">Amount</th></tr>'+preview.map(function(row){return '<tr style="border-bottom:1px solid var(--line)"><td style="padding:5px 6px;color:var(--text-2)">'+get(row,'date')+'</td><td style="padding:5px 6px;max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">'+get(row,'description')+'</td><td style="padding:5px 6px;text-align:right;font-weight:600">'+get(row,'amount')+'</td></tr>';}).join('')+'</table>';}
// #4 · Robust CSV date parsing. Never blindly trusts new Date('06/07/2026')
// (which silently reads it as US M/D). Honors an explicit format choice, and in
// Auto mode infers DD/MM vs MM/DD from the whole column.
function parseCsvDate(raw, fmt){
  var s=String(raw||'').trim();
  if(!s) return '';
  var iso=s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if(iso) return iso[1]+'-'+iso[2].padStart(2,'0')+'-'+iso[3].padStart(2,'0');
  var parts=s.split(/[\/\-\.\s]+/).filter(Boolean);
  if(parts.length>=3 && /^\d+$/.test(parts[0]) && /^\d+$/.test(parts[1]) && /^\d+$/.test(parts[2])){
    var p0=parts[0],p1=parts[1],p2=parts[2];
    if(p0.length===4) return p0+'-'+p1.padStart(2,'0')+'-'+p2.padStart(2,'0'); // yyyy/mm/dd
    var yr=p2.length===4?p2:(parseInt(p2,10)>=70?'19'+p2.padStart(2,'0'):'20'+p2.padStart(2,'0'));
    var d,m;
    if(fmt==='mdy'){ m=p0; d=p1; } else { d=p0; m=p1; } // default dmy
    // Hard guard: if the chosen 'month' can't be a month but 'day' can, swap.
    if(parseInt(m,10)>12 && parseInt(d,10)<=12){ var t=m; m=d; d=t; }
    return yr+'-'+m.padStart(2,'0')+'-'+d.padStart(2,'0');
  }
  var dt=new Date(s); // named months e.g. "5 Jul 2026"
  return isNaN(dt) ? '' : isoLocal(dt);
}
function inferDateFmt(rows, idx){
  var dmy=0,mdy=0;
  for(var i=0;i<rows.length;i++){
    var s=String(rows[i][idx]||'').trim();
    if(/^\d{4}-/.test(s)) continue;
    var p=s.split(/[\/\-\.\s]+/).filter(Boolean);
    if(p.length<2||!/^\d+$/.test(p[0])||!/^\d+$/.test(p[1])||p[0].length===4) continue;
    var a=parseInt(p[0],10), b=parseInt(p[1],10);
    if(a>12&&b<=12) dmy++; else if(b>12&&a<=12) mdy++;
  }
  return mdy>dmy ? 'mdy' : 'dmy'; // tie → DD/MM (most of the world)
}
// #28 · description → category keyword rules for smarter imports
var CSV_CAT_RULES=[
  [/grocer|supermarket|restaurant|dining|food|meal|pizza|burger|kebab|sushi|mcdonald|kfc|subway|uber\s?\*?\s?eats|doordash|menulog|woolworths|coles|aldi|iga\b/i,'Food'],
  [/uber(?!\s?\*?\s?eats)|taxi|lyft|fuel|petrol|gas\s?station|shell|bp\b|caltex|train|bus|metro|parking|toll|opal|myki/i,'Transport'],
  [/amazon|ebay|kmart|target|myer|big\s?w|shop|store|retail|clothing|fashion/i,'Shopping'],
  [/pharmac|chemist|doctor|dental|medic|hospital|clinic|physio|optom/i,'Health'],
  [/netflix|spotify|disney|stan\b|binge|cinema|movie|theatre|game|steam|playstation|xbox/i,'Entertainment'],
  [/electric|energy|water\s?bill|gas\s?bill|internet|broadband|telstra|optus|vodafone|insurance|council|rates|subscription|rent\b|mortgage/i,'Bills'],
  [/school|tuition|course|udemy|coursera|textbook|university|tafe/i,'Education'],
  [/hotel|airbnb|flight|qantas|jetstar|virgin|booking\.com|expedia|hostel/i,'Travel'],
  [/coffee|cafe|starbucks|espresso|barista/i,'Coffee'],
  [/gym|fitness|anytime|yoga|pilates|sport/i,'Fitness'],
  [/bunnings|ikea|furniture|hardware|repair|plumb|electrician/i,'Home'],
  [/salary|payroll|wages?\b|paycheck|pay\s?run/i,'Salary'],
  [/dividend|interest\s?(paid|earned)|distribution/i,'Investment']
];
function guessCatFromDesc(desc){
  var d=String(desc||'');
  if(!d) return '';
  for(var i=0;i<CSV_CAT_RULES.length;i++){ if(CSV_CAT_RULES[i][0].test(d)) return CSV_CAT_RULES[i][1]; }
  return '';
}
function doImportCSV(){
  if(!_csvParsed)return;
  var dateIdx=_csvColumnMap.date,amtIdx=_csvColumnMap.amount;
  if(dateIdx===''||amtIdx===''||dateIdx===undefined||amtIdx===undefined){toast('Map Date and Amount first');return;}
  var acctSel=document.getElementById('csv-acct-select');
  var acctName=acctSel&&acctSel.value?acctSel.value:(S.accts.length>0?S.accts[0].name:'');
  var fmtSel=document.getElementById('csv-dateformat-select');
  var fmtChoice=fmtSel?fmtSel.value:'auto';
  var _csvDateFmt=(fmtChoice==='auto'||!fmtChoice)?inferDateFmt(_csvParsed,dateIdx):fmtChoice;
  var baseTs=Date.now();
  // #28 · no Type column but the amounts mix negatives and positives?
  // Use the sign: negative = expense, positive = income (standard bank exports).
  var hasTypeCol=(_csvColumnMap.type!==''&&_csvColumnMap.type!==undefined);
  var hasNeg=false,hasPos=false;
  if(!hasTypeCol){
    _csvParsed.forEach(function(row){
      var v=parseFloat(String(row[amtIdx]||'').replace(/[^0-9.\-]/g,''));
      if(!isNaN(v)&&v!==0){ if(v<0)hasNeg=true; else hasPos=true; }
    });
  }
  var signTyping=!hasTypeCol&&hasNeg&&hasPos&&S.trackIncome;
  // #28 · skip rows that already exist (re-importing the same bank export is common).
  // Occurrence-COUNTED, not a plain set: two genuinely identical purchases (same
  // day/amount/merchant — e.g. two coffees) are two rows in the export and both
  // must import. Each existing tx only "absorbs" one matching CSV row.
  var seen={};
  (S.txs||[]).forEach(function(t){ var k=String(t.date).slice(0,10)+'|'+t.type+'|'+Number(t.amount)+'|'+String(t.description||'').toLowerCase(); seen[k]=(seen[k]||0)+1; });
  var dup=0;
  var memSel=document.getElementById('csv-member-select');
  var memberId=(S.family&&memSel&&memSel.value)?memSel.value:(S.memberId||null);
  var imported=[];
  _csvParsed.forEach(function(row,i){
    var rawDate=row[dateIdx]||'';
    var rawAmt=row[amtIdx]||'';
    var dateStr=parseCsvDate(rawDate,_csvDateFmt);
    if(!dateStr)return;
    var amt=parseFloat(rawAmt.toString().replace(/[^0-9.\-]/g,''));
    if(isNaN(amt)||amt===0)return;
    var type='expense';
    if(hasTypeCol){
      var rt=(row[_csvColumnMap.type]||'').toLowerCase();
      if(/income|credit|cr\b|deposit|salary/.test(rt))type='income';
    } else if(signTyping){
      type=amt<0?'expense':'income';
    }
    var desc=(_csvColumnMap.description!==''&&_csvColumnMap.description!==undefined)?row[_csvColumnMap.description]||'':'Imported';
    var catRaw=(_csvColumnMap.category!==''&&_csvColumnMap.category!==undefined)?row[_csvColumnMap.category]||'':'';
    var cat=catRaw||guessCatFromDesc(desc)||'Other';
    var notes=(_csvColumnMap.notes!==''&&_csvColumnMap.notes!==undefined)?row[_csvColumnMap.notes]||'':'';
    var key=dateStr+'|'+type+'|'+Math.abs(amt)+'|'+String(desc).toLowerCase();
    if(seen[key]>0){seen[key]--;dup++;return;} // absorb one existing copy, keep the rest
    imported.push({
      id:'tx_'+(baseTs+i),
      date:dateStr,type:type,category:cat,description:desc,
      amount:Math.abs(amt),account:acctName,notes:notes,
      memberId:memberId,recurring:false
    });
  });
  if(!imported.length){toast(dup?('All '+dup+' rows already in your ledger'):'No valid rows found');return;}
  closeImportCSV();
  var dupNote=dup?(' · '+dup+' duplicate'+(dup===1?'':'s')+' skipped'):'';
  // Apply every row to the local cache first, then sync once (firebase) or with a
  // small concurrent pool (sheets) · much faster than sequential single requests.
  imported.forEach(function(tx){
    handleLocalPost({action:'addTransaction',id:tx.id,date:tx.date,type:tx.type,
      category:tx.category,description:tx.description,amount:tx.amount,
      account:tx.account,notes:tx.notes,memberId:tx.memberId});
  });
  renderAll();
  if(S.mode==='firebase'){
    fbWriteAll().catch(function(e){console.warn(e);});
    toast('✅ Imported '+imported.length+dupNote);
  } else if(S.mode==='sheets'){
    // Queue every imported row through the durable outbox so the import is
    // offline-safe too (replayed on reconnect) and can't be lost mid-sync.
    imported.forEach(function(tx){
      outboxEnqueue({action:'addTransaction',id:tx.id,date:tx.date,type:tx.type,category:tx.category,description:tx.description,amount:tx.amount,account:tx.account,notes:tx.notes,memberId:tx.memberId});
    });
    outboxFlush();
    toast('✅ Imported '+imported.length+dupNote+' · syncing…');
  } else {
    toast('✅ Imported '+imported.length+dupNote);
  }
}


// Destructive reset, deliberately hard to do by accident: the primary action of the
// first dialog IS downloading a CSV backup. Only after that does the real delete ask.
function deleteAndResetSheet() {
  showCustomConfirm('Back up your data first',
    'You are about to disconnect Lucid and wipe its data on this device. Download a CSV copy of every transaction first — one tap, and you will thank yourself later.',
    'Download CSV backup',
    () => {
      try { doExport(); } catch(e) { console.warn(e); }
      setTimeout(_confirmDeleteReset, 700);
    });
}
function _confirmDeleteReset() {
  showCustomConfirm('Delete and Reset Sheet?',
    'Lucid on this device resets and returns to the setup screen. The Sheet / cloud data itself is NOT deleted — this device just forgets it. Remove the Sheet from your Drive yourself if you want it gone for good.',
    'Delete and reset',
    async () => { stopPolling(); if (S.mode === 'firebase') await signOutFirebase(); localStorage.clear(); location.reload(); });
}
function reconnect() {
  const msg = S.mode === 'firebase'
    ? 'Sign out of Google on this device? Your data stays safe in the cloud.'
    : S.mode === 'sheets'
      ? 'Disconnect your Google Sheet? The Sheet itself is untouched — this device just forgets the link.'
      : 'Change how Lucid stores data? This resets Lucid on this device, including locally saved data. Export a CSV first if you need it.';
  showCustomConfirm('Change storage?', msg, 'Continue', () => {
    setTimeout(() => {
      showCustomConfirm('Are you sure?',
        'Lucid on this device will reset and return to the setup screen. Nothing in your Google account is deleted.',
        'Yes, continue',
        async () => { stopPolling(); if (S.mode === 'firebase') await signOutFirebase(); localStorage.clear(); location.reload(); });
    }, 180);
  });
}

function editName() {
  const n = prompt('Display name:', S.name);
  if (!n) return;
  S.name = n; localStorage.setItem('prizm_name', n);
  renderAll();
  apiPost({action: 'updateSettings', key: 'name', value: n});
}

function editCurrency() {
  const wrap = document.getElementById('currency-list');
  wrap.innerHTML = CURRENCIES.map(c => `
    <div class="srow" onclick="selectCurrency('${c.code}')">
      <div class="srow-ico" style="font-size:13px;font-weight:800">${esc(c.sym)}</div>
      <div class="srow-body"><div class="srow-title">${c.code}</div><div class="srow-sub">${esc(c.name||'')}</div></div>
      ${S.cur.code === c.code ? '<div style="color:var(--accent);font-weight:800">✓</div>' : ''}
    </div>`).join('');
  openOverlay('currency-overlay');
}
function selectCurrency(code) {
  const found = CURRENCIES.find(x => x.code === code);
  if (!found || found.code === S.cur.code) { closeOverlay('currency-overlay'); return; }
  const apply = () => {
    S.cur = found; localStorage.setItem('prizm_cur', found.code);
    closeOverlay('currency-overlay');
    renderAll();
    apiPost({action: 'updateSettings', key: 'currency', value: found.code});
    toast('✓ Currency updated');
  };
  // Display-only relabel — no conversion happens. Warn when history exists so mixed-
  // currency amounts aren't silently summed under one symbol.
  if (S.txs && S.txs.length) {
    showCustomConfirm('Change currency?',
      'This changes the display symbol only — your ' + S.txs.length + ' existing amounts are NOT converted. Totals will mix old ' + S.cur.code + ' amounts with new ' + found.code + ' entries under one symbol.',
      'Change anyway', apply);
  } else apply();
}

let _tt;
function toast(msg) {
  const el = document.getElementById('toast'); el.textContent = msg; el.classList.add('show');
  clearTimeout(_tt); _tt = setTimeout(() => el.classList.remove('show'), 2600);
}

// ═══════════════════════════════════════════
//  RENDERING ENGINE
// ═══════════════════════════════════════════
let _renderRaf = null;
function renderAll() {
  if (_renderRaf) return;
  _renderRaf = requestAnimationFrame(() => { _renderRaf = null; _doRenderAll(); });
}
function _doRenderAll() {
  const active = document.querySelector('.page.active')?.id || 'page-home';
  try { renderHome(); } catch(e) { console.warn('renderHome:', e); }
  try { renderUpcoming(); } catch(e) { console.warn('renderUpcoming:', e); }
  try { renderGoals(); } catch(e) { console.warn('renderGoals:', e); }
  try { renderMemberChips(); } catch(e) {}
  if (active === 'page-transactions') { try { renderAllTx(); } catch(e) { console.warn('renderAllTx:', e); } }
  if (active === 'page-settings') {
    try { renderSettings(); } catch(e) { console.warn('renderSettings:', e); }
    try { updatePrefUI(); } catch(e) {}
  }
  try { updateAcctSelect(); updateRecAcctSelect(); } catch(e) {}
  if (active === 'page-analytics') { try { renderAnalytics(); } catch(e) {} }
}

// ── Upcoming bills (Home) ────────────────────────────────────────────────────
// A cash-flow-foresight card: recurring EXPENSES due in the next 14 days, with a
// running total, so users see what's about to leave their account. Hides itself
// when there's nothing upcoming.
function _dueLabel(d) {
  const today = todayISO();
  const tmrw = isoLocal(new Date(Date.now() + 86400000));
  if (d <= today) return 'Due today';
  if (d === tmrw) return 'Due tomorrow';
  const days = Math.round((parseDate(d).getTime() - parseDate(today).getTime()) / 86400000);
  return 'In ' + days + ' days · ' + fmtDateShort(d);
}
function cycleUpcomingDays() {
  const opts = [7, 14, 30, 60];
  const cur = Number(localStorage.getItem('prizm_upcoming_days')) || 14;
  const next = opts[(opts.indexOf(cur) + 1) % opts.length] || 14;
  try { localStorage.setItem('prizm_upcoming_days', String(next)); } catch (e) {}
  try { toast('Upcoming: next ' + next + ' days'); } catch (e) {}
  renderUpcoming();
}
function renderUpcoming() {
  const el = document.getElementById('home-upcoming');
  if (!el) return;
  const today = todayISO();
  const _upDays = Number(localStorage.getItem('prizm_upcoming_days')) || 14;
  const horizon = isoLocal(new Date(Date.now() + _upDays * 86400000));
  const due = (S.recurring || [])
    .filter(r => r.active !== false && r.type === 'expense' && r.nextDate)
    .map(r => ({ r: r, d: String(r.nextDate).slice(0, 10) }))
    .filter(x => x.d >= today && x.d <= horizon)
    .sort((a, b) => a.d < b.d ? -1 : (a.d > b.d ? 1 : 0));
  if (!due.length) { el.style.display = 'none'; el.innerHTML = ''; return; }
  el.style.display = '';
  const total = due.reduce((s, x) => s + (Number(x.r.amount) || 0), 0);
  const rowsHtml = due.slice(0, 4).map(x => {
    const emoji = getCatEmoji(x.r.category, false);
    return '<div style="display:flex;align-items:center;gap:11px;padding:9px 0;border-top:1px solid var(--line)">' +
      '<div style="font-size:17px;flex-shrink:0">' + esc(emoji) + '</div>' +
      '<div style="flex:1;min-width:0">' +
        '<div style="font-size:13px;font-weight:600;color:var(--text-1);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">' + esc(x.r.name) + '</div>' +
        '<div style="font-size:11px;color:var(--text-3)">' + esc(_dueLabel(x.d)) + '</div>' +
      '</div>' +
      '<div style="font-size:13px;font-weight:700;color:var(--text-1);flex-shrink:0">' + fmtMoney(Number(x.r.amount) || 0) + '</div>' +
    '</div>';
  }).join('');
  const more = due.length > 4 ? '<div style="font-size:11px;color:var(--text-3);padding-top:8px">+' + (due.length - 4) + ' more</div>' : '';
  el.innerHTML = '<div style="margin:16px 16px 0;padding:16px;border-radius:var(--r-lg);background:var(--bg-2);border:1px solid var(--line-2)">' +
    '<div style="display:flex;align-items:baseline;justify-content:space-between;margin-bottom:4px">' +
      '<div style="font-size:13px;font-weight:700;color:var(--text-1)">Upcoming bills</div>' +
      '<div onclick="cycleUpcomingDays()" style="font-size:11px;color:var(--text-3);cursor:pointer" title="Tap to change">Next ' + _upDays + ' days &#8964;</div>' +
    '</div>' +
    '<div style="font-size:22px;font-weight:800;letter-spacing:-0.5px;color:var(--text-1);margin-bottom:6px">' + fmtMoney(total) + '</div>' +
    rowsHtml + more +
  '</div>';
}

// ── Savings goals ────────────────────────────────────────────────────────────
// Fully self-contained: its own localStorage key, its own card + modal. Never
// touches transactions, accounts, or sync.
function _goals() { return safeParse(localStorage.getItem('prizm_goals'), []); }
function _saveGoals(g) { try { localStorage.setItem('prizm_goals', JSON.stringify(g)); } catch(e) {} }
let _editingGoalId = null;
function renderGoals() {
  const el = document.getElementById('home-goals');
  if (!el) return;
  // Savings only make sense alongside income — show goals only when the user
  // tracks income or has recorded some. Otherwise keep the card hidden entirely.
  const showIncome = S.trackIncome || (S.txs || []).some(t => t.type === 'income');
  if (!showIncome) { el.style.display = 'none'; el.innerHTML = ''; return; }
  const goals = _goals();
  let rows = '';
  goals.forEach(g => {
    const t = Number(g.target) || 0, s = Number(g.saved) || 0;
    const pct = t > 0 ? Math.max(0, Math.min(100, Math.round(s / t * 100))) : 0;
    const gid = esc(jsq(g.id));
    rows += '<div style="padding:11px 0;border-top:1px solid var(--line)">' +
      '<div style="display:flex;align-items:center;gap:10px;margin-bottom:8px">' +
        '<div style="font-size:16px">' + esc(g.emoji || '🎯') + '</div>' +
        '<div onclick="openGoalModal(\'' + gid + '\')" style="cursor:pointer;flex:1;min-width:0">' +
          '<div style="font-size:13px;font-weight:600;color:var(--text-1);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">' + esc(g.name) + '</div>' +
          '<div style="font-size:11px;color:var(--text-3)">' + fmtMoney(s) + ' of ' + fmtMoney(t) + ' &middot; ' + pct + '%</div>' +
        '</div>' +
        '<button onclick="openGoalAdd(\'' + gid + '\')" style="flex-shrink:0;border:none;cursor:pointer;background:var(--accent-soft);color:var(--accent);font-size:12px;font-weight:700;padding:7px 13px;border-radius:999px">+ Add</button>' +
      '</div>' +
      '<div style="height:7px;border-radius:99px;background:var(--bg-4);overflow:hidden"><div style="height:100%;width:' + pct + '%;background:var(--accent);border-radius:99px;transition:width .3s"></div></div>' +
    '</div>';
  });
  el.style.display = '';
  el.innerHTML = '<div style="margin:16px 16px 0;padding:16px;border-radius:var(--r-lg);background:var(--bg-2);border:1px solid var(--line-2)">' +
    '<div style="display:flex;align-items:center;justify-content:space-between">' +
      '<div style="font-size:13px;font-weight:700;color:var(--text-1)">Savings goals</div>' +
      '<div onclick="openGoalModal()" style="cursor:pointer;font-size:12px;font-weight:600;color:var(--accent)">New goal</div>' +
    '</div>' +
    (goals.length ? rows : '<div style="font-size:12px;color:var(--text-3);padding-top:8px">Set a goal to start tracking your savings.</div>') +
  '</div>';
}
function openGoalModal(id) {
  _editingGoalId = id || null;
  const g = id ? _goals().find(x => x.id === id) : null;
  document.getElementById('goal-modal-title').textContent = g ? 'Edit goal' : 'New goal';
  document.getElementById('goal-name').value = g ? (g.name || '') : '';
  document.getElementById('goal-target').value = g ? (g.target || '') : '';
  document.getElementById('goal-saved').value = g ? (g.saved || '') : '';
  document.getElementById('goal-emoji').value = g ? (g.emoji || '') : '';
  document.getElementById('goal-delete-btn').style.display = g ? '' : 'none';
  openOverlay('goal-overlay');
}
function submitGoal() {
  const name = document.getElementById('goal-name').value.trim();
  if (!name) { toast('Enter a name'); return; }
  const target = parseFloat(document.getElementById('goal-target').value) || 0;
  const saved = parseFloat(document.getElementById('goal-saved').value) || 0;
  const emoji = document.getElementById('goal-emoji').value.trim() || '🎯';
  const goals = _goals();
  if (_editingGoalId) { const g = goals.find(x => x.id === _editingGoalId); if (g) { g.name = name; g.target = target; g.saved = saved; g.emoji = emoji; } }
  else { goals.push({ id: 'goal_' + Date.now(), name: name, target: target, saved: saved, emoji: emoji }); }
  _saveGoals(goals); _editingGoalId = null; closeOverlay('goal-overlay'); renderGoals();
  toast('✅ Goal saved');
}
function deleteGoal() {
  if (!_editingGoalId) return;
  const id = _editingGoalId;
  showCustomConfirm('Delete goal?', 'This removes the goal. Your transactions are not affected.', 'Delete', function () {
    _saveGoals(_goals().filter(x => x.id !== id)); _editingGoalId = null; closeOverlay('goal-overlay'); renderGoals();
  });
}
// Contribute to a goal by entering an amount to ADD — no mental maths needed.
let _contribGoalId = null;
function openGoalAdd(id) {
  _contribGoalId = id;
  const g = _goals().find(x => x.id === id);
  document.getElementById('goal-add-title').textContent = g ? ('Add to ' + g.name) : 'Add to goal';
  document.getElementById('goal-add-amt').value = '';
  openOverlay('goal-add-overlay');
  setTimeout(function(){ const a = document.getElementById('goal-add-amt'); if (a && window.innerWidth >= 640) a.focus(); }, 120);
}
function commitGoalAdd() {
  const amt = parseFloat(document.getElementById('goal-add-amt').value);
  if (!amt || amt <= 0) { toast('Enter an amount'); return; }
  const goals = _goals();
  const g = goals.find(x => x.id === _contribGoalId);
  if (g) { g.saved = (Number(g.saved) || 0) + amt; _saveGoals(goals); }
  _contribGoalId = null;
  closeOverlay('goal-add-overlay');
  renderGoals();
  toast('✅ Added to goal');
}

// ── Quick-add chips (add modal) ──────────────────────────────────────────────
// Tap a recent entry to prefill amount/description/category. New entries only;
// never shown while editing.
function renderQuickAdd() {
  const el = document.getElementById('quick-add-chips');
  if (!el) return;
  if (S.editingTxId) { el.style.display = 'none'; el.innerHTML = ''; return; }
  const seen = {}, picks = [];
  for (const t of (S.txs || [])) {
    const d = (t.description || '').trim();
    if (!d) continue;
    const key = d.toLowerCase();
    if (seen[key]) continue;
    seen[key] = 1;
    picks.push(t);
    if (picks.length >= 6) break;
  }
  if (!picks.length) { el.style.display = 'none'; el.innerHTML = ''; return; }
  el.style.display = 'block';
  const chips = picks.map(t => {
    const emoji = getCatEmoji(t.category, t.type === 'income');
    const amt = fmtMoney(Number(t.amount) || 0);
    return '<button type="button" onclick="quickFill(\'' + esc(jsq(t.id)) + '\')" style="cursor:pointer;display:inline-flex;align-items:center;gap:7px;padding:8px 12px;border-radius:999px;background:var(--bg-3);border:1px solid var(--line-2);font-size:12px;color:var(--text-1);white-space:nowrap;max-width:190px;overflow:hidden;font-family:inherit">' +
      '<span style="font-size:14px;flex-shrink:0">' + esc(emoji) + '</span>' +
      '<span style="overflow:hidden;text-overflow:ellipsis;font-weight:600">' + esc(t.description) + '</span>' +
      '<span style="color:var(--text-3);flex-shrink:0">' + amt + '</span>' +
    '</button>';
  }).join('');
  el.innerHTML = '<div style="font-size:11px;font-weight:700;letter-spacing:.3px;text-transform:uppercase;color:var(--text-3);margin-bottom:8px">Quick add</div>' +
    '<div style="display:flex;gap:8px;flex-wrap:wrap">' + chips + '</div>';
}
function quickFill(id) {
  const t = (S.txs || []).find(x => x.id === id);
  if (!t) return;
  setType(t.type === 'income' ? 'income' : 'expense');
  const amtEl = document.getElementById('add-amt');
  if (amtEl) amtEl.value = (Number(t.amount) || '').toString();
  const descEl = document.getElementById('add-desc');
  if (descEl) descEl.value = t.description || '';
  S.cat = t.category;
  try { buildCatGrid(); } catch(e) {}
}

function renderHome() {
  updateGreeting();
  const now = new Date();
  // Hero card totals reset over the duration the user picked (week/month/year/all)
  const hp = S.heroPeriod || 'month';
  let periodStart = 0, periodLbl = 'All time';
  if (hp === 'week') { periodStart = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime() - 6*86400000; periodLbl = 'Last 7 days'; }
  else if (hp === 'month') { periodStart = new Date(now.getFullYear(), now.getMonth(), 1).getTime(); periodLbl = now.toLocaleString('default', { month: 'long', year: 'numeric' }); }
  else if (hp === 'year') { periodStart = new Date(now.getFullYear(), 0, 1).getTime(); periodLbl = String(now.getFullYear()); }
  document.getElementById('hero-period').textContent = periodLbl;

  let mInc = 0, mExp = 0;
  S.txs.forEach(t => {
    if (dateMs(t.date) >= periodStart) { if (t.type === 'income') mInc += (parseFloat(t.amount)||0); else mExp += (parseFloat(t.amount)||0); }
  });
  const netWorth = S.accts.reduce((sum,a) => sum + parseFloat(a.balance||0), 0);
  // Income is "in play" if the user tracks income OR has ever recorded any. When it
  // isn't, net worth ≈ (–)spending, so the hero shows only Spent to avoid duplication.
  const showIncome = S.trackIncome || S.txs.some(t => t.type === 'income');
  const heroLabelEl = document.querySelector('#home-hero .hero-label');
  const heroRowEl   = document.querySelector('#home-hero .hero-row');
  const incStatEl   = document.getElementById('hero-income-stat');
  if (showIncome) {
    if (heroLabelEl) heroLabelEl.textContent = 'Balance';
    document.getElementById('hero-balance').textContent = fmtMoney(netWorth);
    if (heroRowEl) heroRowEl.style.display = '';
    if (incStatEl) incStatEl.style.display = '';
    document.getElementById('hero-income').textContent = fmtMoneyCompact(mInc);
    document.getElementById('hero-spent').textContent = fmtMoneyCompact(mExp);
  } else {
    // Spending-only view: the big number becomes what was spent this period.
    if (heroLabelEl) heroLabelEl.textContent = 'Spent';
    document.getElementById('hero-balance').textContent = fmtMoney(mExp);
    if (heroRowEl) heroRowEl.style.display = 'none';
  }
  fitHeroAmount(document.getElementById('hero-balance'), 44, 24);

  const chips = document.getElementById('chips-row');
  const acctSh = document.getElementById('acct-sh');
  if (!S.accts || !S.accts.length) {
    if (acctSh) acctSh.style.display = 'flex';
    chips.innerHTML = '<div style="color:var(--text-3);font-size:13px;padding:4px 2px;font-weight:500">No accounts yet · tap Add account above</div>';
  } else {
    if (acctSh) acctSh.style.display = 'none';
    chips.innerHTML = S.accts.map(a => {
      const ico = getAcctIcon(a.type);
      return `<div class="chip" onclick="openAccountAnalytics('${esc(jsq(a.name))}')"><div class="chip-ico">${ico}</div><div class="chip-name">${esc(a.name)}</div><div class="chip-bal">${fmtMoneyCompact(a.balance)}</div></div>`;
    }).join('');
  }

  renderBudgetSnapshot();

  let homeTxs = S.txs;
  if (S.family && S.memberFilter !== 'all') homeTxs = homeTxs.filter(t => t.memberId === S.memberFilter);
  renderTxList(homeTxs.slice(0, 5), 'home-txs', true);
  if (S.simpleMode) renderSimpleHome();
}

function renderBudgetSnapshot() {
  const wrap = document.getElementById('budget-snapshot');
  if (!S.budgets || !S.budgets.length) { wrap.innerHTML = ''; return; }
  const now = new Date();
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1).getTime();
  const spend = {};
  S.txs.forEach(t => { if (t.type === 'expense' && dateMs(t.date) >= monthStart) spend[t.category] = (spend[t.category]||0) + (parseFloat(t.amount)||0); });
  wrap.innerHTML = `<div class="sh"><h2>Budgets</h2><span class="sh-action" onclick="goTab('analytics')">View all</span></div>` +
    `<div class="chart-box" style="margin-bottom:14px">` +
    S.budgets.slice(0,3).map(b => {
      const spent = spend[b.category] || 0;
      const pct = b.limit > 0 ? Math.min(100, Math.round(spent/b.limit*100)) : 0;
      const over = spent > b.limit;
      return `<div class="budget-row">
        <div class="budget-top"><div class="budget-name">${getCatIcon(b.category,false)} ${esc(b.category)}</div><div class="budget-amt ${over?'over':''}">${fmtMoney(spent)} / ${fmtMoney(b.limit)}</div></div>
        <div class="budget-track"><div class="budget-fill ${over?'over':''}" style="width:${pct}%"></div></div>
      </div>`;
    }).join('') +
    `</div>`;
}

// Readable text color (near-black / white) for an arbitrary hex background — used
// for member badges so dark member colors don't get illegible black initials (F25).
function contrastText(hex) {
  try {
    const h = String(hex).replace('#','');
    const r = parseInt(h.slice(0,2),16), g = parseInt(h.slice(2,4),16), b = parseInt(h.slice(4,6),16);
    return ((0.299*r + 0.587*g + 0.114*b) / 255) > 0.6 ? '#0B0C0E' : '#FFFFFF';
  } catch(e) { return '#0B0C0E'; }
}
function txHTML(t) {
  const isInc = t.type === 'income';
  const colorClass = isInc ? 'income' : '';
  const dStr = fmtDateShort(t.date);
  let memberBadge = '';
  if (S.family && t.memberId) {
    if (t.memberId === JOINT_ID) {
      memberBadge = `<span class="tx-member" style="background:var(--accent);color:var(--on-accent)" title="${esc(JOINT_LABEL)}">👥</span>`;
    } else {
      const m = S.members.find(x => x.id === t.memberId);
      if (m) memberBadge = `<span class="tx-member" style="background:${esc(m.color)};color:${contrastText(m.color)}" title="${esc(m.name)}">${esc((m.name||'?').charAt(0).toUpperCase())}</span>`;
    }
  }
  return `
    <div class="tx-item" onclick="openEditTx('${esc(t.id)}')">
      <div class="tx-ico" style="background:${getCatColor(t.category)}">${getCatIcon(t.category, isInc)}</div>
      <div class="tx-body">
        <div class="tx-name"><span class="tx-text">${esc(t.description || t.category)}</span>${memberBadge}</div>
        <div class="tx-meta"><span class="tx-tag">${esc(t.category)}</span>${t.account && t.account !== '—' ? `<span style="color:var(--text-3)">${esc(t.account)}</span>` : ''}</div>
      </div>
      <div class="tx-right">
        <div class="tx-amt ${colorClass}">${isInc ? '+' : '-'}${fmtMoney(t.amount)}</div>
        <div class="tx-date">${dStr}</div>
      </div>
    </div>`;
}

function renderTxList(arr, containerId, isEmptyCompact = false) {
  const container = document.getElementById(containerId);
  if (!arr || !arr.length) {
    container.innerHTML = `<div class="empty"><div class="empty-icon">○</div><div class="empty-title">No transactions found</div><div class="empty-sub">${isEmptyCompact ? 'Tap + to record your first ledger entry.' : 'Adjust filters or add entries.'}</div></div>`;
    return;
  }
  container.innerHTML = arr.map(t => txHTML(t)).join('');
}

function renderMemberChips() {
  const row = document.getElementById('member-chip-row');
  if (!S.family || !S.members.length) { row.style.display = 'none'; row.innerHTML = ''; return; }
  row.style.display = 'flex';
  let html = `<div class="mchip ${S.memberFilter==='all'?'active':''}" onclick="setMemberFilter('all',this)">All</div>`;
  html += S.members.map(m => `<div class="mchip ${S.memberFilter===m.id?'active':''}" onclick="setMemberFilter('${esc(m.id)}',this)"><span class="mchip-dot" style="background:${esc(m.color)}"></span>${esc(m.name)}</div>`).join('');
  // #3 · Joint / Shared filter chip
  html += `<div class="mchip ${S.memberFilter===JOINT_ID?'active':''}" onclick="setMemberFilter('${JOINT_ID}',this)"><span class="mchip-dot" style="background:var(--accent);color:#fff;font-size:9px">👥</span>${esc(JOINT_LABEL)}</div>`;
  row.innerHTML = html;
}

function setMemberFilter(id, el) {
  S.memberFilter = id;
  document.querySelectorAll('.mchip').forEach(c => c.classList.remove('active'));
  if (el) el.classList.add('active');
  renderAll();
}

// Debounced wrapper for the search box: don't rebuild the whole list on every keystroke.
let _txSearchTimer = null;
function renderAllTxDebounced() {
  clearTimeout(_txSearchTimer);
  S._txRenderCap = TX_RENDER_CHUNK; // new query → reset pagination
  _txSearchTimer = setTimeout(renderAllTx, 180);
}
const TX_RENDER_CHUNK = 200; // rows rendered before the "Show more" tail
function showMoreTx() {
  S._txRenderCap = (S._txRenderCap || TX_RENDER_CHUNK) + TX_RENDER_CHUNK;
  renderAllTx();
}
// Build the ledger category-filter chips from the live category list so new,
// renamed, and custom categories always show (previously a fixed set of 7).
function renderFilterChips() {
  const row = document.getElementById('tx-filter-row');
  if (!row) return;
  const cats = getCatList(false).map(c => c.n);
  // Skip rebuilds when nothing changed (avoids resetting horizontal scroll).
  const sig = S.filter + '|' + (S.trackIncome ? '1' : '0') + '|' + cats.join('');
  if (row._sig === sig) return;
  row._sig = sig;
  let html = '<div class="fc' + (S.filter === 'all' ? ' active' : '') + '" onclick="setFilter(\'all\',this)">All</div>';
  cats.forEach(function (n) {
    html += '<div class="fc' + (S.filter === n ? ' active' : '') + '" onclick="setFilter(\'' + esc(jsq(n)) + '\',this)">' + esc(n) + '</div>';
  });
  if (S.trackIncome) {
    html += '<div class="fc' + (S.filter === 'income' ? ' active' : '') + '" id="fc-income" onclick="setFilter(\'income\',this)">Income</div>';
  }
  row.innerHTML = html;
}

function renderAllTx() {
  renderFilterChips();
  const searchEl = document.getElementById('search-in');
  const query = (searchEl ? searchEl.value : '').toLowerCase();
  let list = S.txs || [];
  if (S.filter === 'expense' || S.filter === 'income') list = list.filter(t => t.type === S.filter);
  else if (S.filter !== 'all') list = list.filter(t => t.category === S.filter || t.account === S.filter);
  if (S.family && S.memberFilter !== 'all') list = list.filter(t => t.memberId === S.memberFilter);
  if (query) list = list.filter(t => ((t.description||'') + ' ' + (t.category||'') + ' ' + (t.account||'') + ' ' + (t.notes||'')).toLowerCase().includes(query));

  document.getElementById('tx-count').textContent = `${list.length} items`;
  // Topbar right context: active filter, or the current month
  const txPeriodEl = document.getElementById('tx-topbar-period');
  if (txPeriodEl) txPeriodEl.textContent = (S.filter && S.filter !== 'all')
    ? ('Filtered · ' + S.filter)
    : new Date().toLocaleDateString('en-US', {month:'long', year:'numeric'});
  const el = document.getElementById('all-txs');

  if (!list.length) {
    el.innerHTML = '<div class="empty"><div class="empty-icon">○</div><div class="empty-title">No matching entries</div><div class="empty-sub">Adjust filters or search.</div></div>';
    return;
  }

  // Cap the DOM at S._txRenderCap rows (long-term datasets reach thousands); the
  // "Show more" tail extends in chunks instead of rebuilding an unbounded list.
  const cap = S._txRenderCap || TX_RENDER_CHUNK;
  const shown = list.length > cap ? list.slice(0, cap) : list;
  const groups = {};
  shown.forEach(t => {
    const k = parseDate(t.date).toLocaleDateString('en-US', {month: 'long', year: 'numeric'});
    (groups[k] = groups[k] || []).push(t);
  });
  let html = Object.entries(groups).map(([m, txs]) => `<div class="month-hd">${m}</div>${txs.map(t => txHTML(t)).join('')}`).join('');
  if (list.length > cap) html += `<button class="btn-ghost" style="width:100%;margin:12px 0" onclick="showMoreTx()">Show ${Math.min(TX_RENDER_CHUNK, list.length - cap)} more (${list.length - cap} remaining)</button>`;
  el.innerHTML = html;
}

function filterByAccount(name) {
  S.filter = name;
  goTab('transactions');
  document.querySelectorAll('.fc').forEach(c => {
    c.classList.remove('active');
    if (c.textContent === name) c.classList.add('active');
  });
  renderAllTx();
  toast('Showing ' + name + ' · tap All to clear');
}

// Tapping an account card on Home opens Analytics → Breakdown → Accounts and
// scrolls to / highlights that specific account.
function openAccountAnalytics(name) {
  goTab('analytics');
  const bdTab = Array.from(document.querySelectorAll('.an-tab')).find(t => /breakdown/i.test(t.textContent));
  setAnalyticsTab('breakdown', bdTab);
  const acctSub = Array.from(document.querySelectorAll('.an-sub-tab')).find(t => /account/i.test(t.textContent));
  if (acctSub) { try { setBreakdownView('account', acctSub); } catch(e) {} }
  setTimeout(() => {
    const wrap = document.getElementById('bd-acct-bars');
    if (!wrap) return;
    wrap.querySelectorAll('.cat-bar').forEach(r => {
      const nm = r.querySelector('.cbi-name');
      if (nm && nm.textContent === name) {
        r.scrollIntoView({ behavior:'smooth', block:'center' });
        r.style.transition = 'background .3s';
        r.style.background = 'var(--accent-soft)';
        setTimeout(() => { r.style.background = ''; }, 1600);
      }
    });
  }, 280);
}

// ═══════════════════════════════════════════
//  ANALYTICS SYSTEM
// ═══════════════════════════════════════════
let _chartRetries = 0;
function _showAnalyticsError() {
  const panel = document.getElementById('an-overview');
  if (panel && !document.getElementById('an-chart-err')) {
    const d = document.createElement('div');
    d.id = 'an-chart-err'; d.className = 'chart-box';
    d.innerHTML = '<div class="no-data" style="cursor:pointer" onclick="_chartRetries=0;renderAnalytics()">⚠️ Charts couldn\'t load · Chart.js may be blocked or you\'re offline.<br>Tap to retry.</div>';
    panel.insertBefore(d, panel.firstChild);
  }
}
function renderAnalytics() {
  // #23 · bounded retry with a visible error state instead of an infinite silent loop.
  if (typeof Chart === 'undefined') {
    if (_chartRetries++ < 20) { setTimeout(renderAnalytics, 300); return; }
    _showAnalyticsError(); return;
  }
  _chartRetries = 0;
  try {
    Chart.defaults.animation = false;
    Chart.defaults.font.family = getComputedStyle(document.body).fontFamily || 'Inter, sans-serif';
    Chart.defaults.color = getComputedStyle(document.documentElement).getPropertyValue('--text-2').trim() || '#9CA1AB';
  } catch(e) {}
  const _err = document.getElementById('an-chart-err'); if (_err) _err.remove();
  const { txs, cutOff, lbl, days } = getAnalyticsTxs();
  // Summary stats
  let tInc = 0, tExp = 0, catTotals = {}, incTotals = {}, acctTotals = {};
  txs.forEach(t => {
    const amt = (parseFloat(t.amount)||0);
    if (t.type === 'income') { tInc += amt; incTotals[t.category] = (incTotals[t.category]||0) + amt; }
    else { tExp += amt; catTotals[t.category] = (catTotals[t.category]||0) + amt; }
    if (t.account) acctTotals[t.account] = (acctTotals[t.account]||0) + (t.type==='expense'?amt:0);
  });
  // Topbar context line: selected period + its spend total
  const anSub = document.getElementById('an-topbar-sub');
  if (anSub) anSub.textContent = lbl + ' · ' + fmtMoney(tExp) + ' spent';
  const saveRate = tInc > 0 ? Math.round(((tInc - tExp) / tInc) * 100) : 0;
  const dailyAvg = days > 0 ? tExp / days : tExp;

  // Update stat cards
  document.getElementById('a-spent').textContent = fmtMoney(tExp);
  document.getElementById('a-txcount').textContent = txs.length + ' transactions';
  document.getElementById('a-income').textContent = fmtMoney(tInc);
  document.getElementById('a-net').textContent = 'Net ' + fmtMoney(tInc - tExp);
  const _sr = document.getElementById('a-save-rate');
  _sr.textContent = saveRate + '%';
  _sr.style.color = saveRate < 0 ? 'var(--red)' : '';
  document.getElementById('a-daily-avg').textContent = fmtMoney(dailyAvg);
  document.getElementById('a-period-lbl').textContent = lbl;
  try { renderInsights(txs, tExp, tInc, catTotals); } catch(e) { console.warn('insights:', e); }

  // Render active tab content
  const activeTab = S.analyticsTab || 'overview';
  renderOverviewTab(txs, tInc, tExp, catTotals);
  renderTrendsTab(txs, tInc, tExp, lbl, days);
  renderBreakdownTab(txs, tInc, tExp, catTotals, incTotals, acctTotals);

  // Family person tab visibility
  const personTab = document.querySelector('.an-sub-tab-person');
  if (personTab) personTab.style.display = S.family ? '' : 'none';
}

function getAnalyticsTxs() {
  const now = new Date(); let cutOff = 0; let cutEnd = Infinity; let lbl = ''; let days = 7;
  // Day-aligned boundary: exactly 7 calendar days including today. A rolling
  // now-minus-168h cutoff included/excluded boundary days depending on the time
  // of day you looked — numbers changed within the same date.
  const sod = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  if (S.period === 'week')  { cutOff = sod - 6*86400000; lbl = 'Last 7 days'; days = 7; }
  else if (S.period === 'month') { cutOff = new Date(now.getFullYear(), now.getMonth(), 1).getTime(); lbl = now.toLocaleDateString('en-US',{month:'long',year:'numeric'}); days = now.getDate(); }
  else if (S.period === '3m')   { cutOff = new Date(now.getFullYear(), now.getMonth()-3, now.getDate()).getTime(); lbl = 'Last 3 months'; days = 91; }
  else if (S.period === 'year') { cutOff = new Date(now.getFullYear(), 0, 1).getTime(); lbl = now.getFullYear()+''; days = Math.ceil((now - new Date(now.getFullYear(),0,1)) / 86400000); }
  else if (S.period === 'custom' && S.periodCustom) {
    const r = _customRange(S.periodCustom);
    cutOff = r[0]; cutEnd = r[1];
    lbl = _customPeriodLabel(S.periodCustom);
    days = Math.max(1, Math.round((Math.min(cutEnd, Date.now()) - cutOff) / 86400000));
  }
  else { cutOff = 0; lbl = 'All time'; days = Math.max(1, S.txs.length > 0 ? Math.ceil((now - new Date(S.txs[S.txs.length-1]?.date||now)) / 86400000) : 1); }
  let txs = S.txs.filter(t => { const ts = dateMs(t.date); return ts >= cutOff && ts < cutEnd; });
  if (S.family && S.memberFilter !== 'all') txs = txs.filter(t => t.memberId === S.memberFilter);
  return { txs, cutOff, lbl, days };
}

// Accurate, plain-language findings computed from the selected period
function renderInsights(txs, tExp, tInc, catTotals) {
  const box = document.getElementById('insights-box');
  const wrap = document.getElementById('insights-rows');
  if (!box || !wrap) return;
  const rows = [];
  const now = new Date();

  if (S.period === 'month') {
    const daysIn = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
    const elapsed = now.getDate();
    // Month-end projection (needs a few days of data to be meaningful)
    if (tExp > 0 && elapsed >= 3 && elapsed < daysIn) {
      rows.push({ i:'📈', t:'On pace to spend ' + fmtMoney(tExp / elapsed * daysIn) + ' this month', s: fmtMoney(tExp / elapsed) + ' per day so far' });
    }
    // Same-point comparison vs last month (day N vs day N · a fair comparison)
    const lmY = now.getMonth() === 0 ? now.getFullYear() - 1 : now.getFullYear();
    const lmM = now.getMonth() === 0 ? 11 : now.getMonth() - 1;
    const lmDays = new Date(lmY, lmM + 1, 0).getDate();
    const lmStart = new Date(lmY, lmM, 1).getTime();
    const lmCut = new Date(lmY, lmM, Math.min(elapsed, lmDays), 23, 59, 59).getTime();
    let lmExp = 0;
    let base = S.txs;
    if (S.family && S.memberFilter !== 'all') base = base.filter(t => t.memberId === S.memberFilter);
    base.forEach(t => { if (t.type === 'expense') { const ts = dateMs(t.date); if (ts >= lmStart && ts <= lmCut) lmExp += parseFloat(t.amount || 0); } });
    if (lmExp > 0) {
      const d = Math.round((tExp - lmExp) / lmExp * 100);
      rows.push({
        i: d > 5 ? '⬆️' : d < -5 ? '⬇️' : '➡️',
        t: Math.abs(d) <= 5 ? 'Spending about the same as last month' : ('Spending ' + Math.abs(d) + '% ' + (d > 0 ? 'more' : 'less') + ' than this time last month'),
        s: fmtMoney(lmExp) + ' by day ' + Math.min(elapsed, lmDays) + ' last month'
      });
    }
  }

  // Top payee in the period
  const per = {};
  txs.forEach(t => { if (t.type === 'expense' && t.description) per[t.description] = (per[t.description] || 0) + parseFloat(t.amount || 0); });
  const top = Object.entries(per).sort((a, b) => b[1] - a[1])[0];
  if (top && tExp > 0 && top[1] >= tExp * 0.1) {
    rows.push({ i:'🏷️', t: esc(top[0]) + ' is your biggest expense', s: fmtMoney(top[1]) + ' · ' + Math.round(top[1] / tExp * 100) + '% of period spending' });
  }

  // Largest category increase vs the previous period
  try {
    const { txs: prevTxs } = getPrevPeriodTxs();
    if (prevTxs.length) {
      const prevCat = {};
      prevTxs.forEach(t => { if (t.type === 'expense') prevCat[t.category] = (prevCat[t.category] || 0) + parseFloat(t.amount || 0); });
      let best = null;
      Object.keys(catTotals).forEach(c => { const inc = catTotals[c] - (prevCat[c] || 0); if (inc > 0 && (!best || inc > best.inc)) best = { c, inc }; });
      if (best && best.inc >= Math.max(20, tExp * 0.1)) {
        rows.push({ i: getCatIcon(best.c, false), t: esc(best.c) + ' is up ' + fmtMoney(best.inc) + ' vs the previous period', s: prevCat[best.c] ? ('was ' + fmtMoney(prevCat[best.c])) : 'new spending this period' });
      }
    }
  } catch(e) {}

  // Income coverage (only when tracking income)
  if (S.trackIncome && tInc > 0 && tExp > 0) {
    const ratio = tExp / tInc;
    if (ratio > 1) rows.push({ i:'⚠️', t:'Spent ' + Math.round((ratio - 1) * 100) + '% more than you earned', s: fmtMoney(tExp - tInc) + ' over income this period' });
    else if (ratio <= 0.7) rows.push({ i:'✅', t:'Kept ' + Math.round((1 - ratio) * 100) + '% of your income', s: fmtMoney(tInc - tExp) + ' saved this period' });
  }

  if (!rows.length) { box.style.display = 'none'; return; }
  box.style.display = '';
  wrap.innerHTML = rows.slice(0, 4).map(r => `
    <div class="ins-row" style="display:flex;gap:12px;align-items:flex-start;padding:11px 0;border-bottom:1px solid var(--line)">
      <div class="cbi" style="font-size:15px;flex-shrink:0">${r.i}</div>
      <div style="flex:1;min-width:0">
        <div style="font-size:13px;font-weight:600;line-height:1.45;overflow:hidden;text-overflow:ellipsis">${r.t}</div>
        <div style="font-size:12px;color:var(--text-3);margin-top:2px">${r.s}</div>
      </div>
    </div>`).join('');
}

function renderOverviewTab(txs, tInc, tExp, catTotals) {
  // Donut
  try {
    const dCtx = document.getElementById('donut-chart');
    if (S.donutChart) S.donutChart.destroy();
    const _noInc = !S.trackIncome;
    const donutBox = dCtx?.closest?.('.chart-box');
    if (donutBox) { donutBox.querySelector('.cb-title').textContent = _noInc ? 'Spending by category' : 'Income vs expenses'; }
    if (tExp === 0 && (tInc === 0 || _noInc)) {
      S.donutChart = new Chart(dCtx, { type: 'doughnut', data: { datasets: [{ data: [1], backgroundColor: [getComputedStyle(document.documentElement).getPropertyValue('--bg-4').trim() || '#1D2024'], borderWidth: 0 }] }, options: { responsive:true, maintainAspectRatio:false, cutout:'75%', plugins:{tooltip:{enabled:false}} }});
    } else if (_noInc) {
      const _cats = Object.entries(catTotals).sort((a,b) => b[1]-a[1]).slice(0,6);
      const _cols = themePalette(Math.max(_cats.length, 1));
      S._noIncCats = _cats; S._noIncCols = _cols;
      S.donutChart = new Chart(dCtx, { type: 'doughnut', data: { labels: _cats.map(c=>c[0]), datasets: [{ data: _cats.map(c=>c[1]), backgroundColor: _cols, borderWidth: 2, borderColor: getComputedStyle(document.documentElement).getPropertyValue('--bg-3').trim() || '#15171A' }] }, options: { responsive:true, maintainAspectRatio:false, cutout:'75%', plugins:{ legend:{display:false}, tooltip:{ callbacks:{ label: ctx => ' ' + ctx.label + ': ' + fmtMoney(ctx.raw) } } } }});
    } else {
      S.donutChart = new Chart(dCtx, { type: 'doughnut', data: { labels: ['Income','Expenses'], datasets: [{ data: [tInc, tExp], backgroundColor: [themeChartColors().inc, themeChartColors().exp], borderWidth: 2, borderColor: getComputedStyle(document.documentElement).getPropertyValue('--bg-3').trim() || '#15171A' }] }, options: { responsive:true, maintainAspectRatio:false, cutout:'75%', plugins:{ legend:{ display:false }, tooltip:{ callbacks:{ label: ctx => ' ' + fmtMoney(ctx.raw) } } } }});
    }
    const net = tInc - tExp;
    const total = tInc + tExp || 1;
    const legend = document.getElementById('donut-legend');
    if (legend) {
      const legendRows = _noInc
        ? (S._noIncCats || []).map((c, i) => ({ color:(S._noIncCols || [])[i] || '#EF4444', label:esc(c[0]), amount:c[1], pct: tExp > 0 ? Math.round(c[1]/tExp*100) : 0 }))
        : [
            { color:themeChartColors().inc, label:'Income', amount:tInc, pct:Math.round(tInc/total*100) },
            { color:themeChartColors().exp, label:'Expenses', amount:tExp, pct:Math.round(tExp/total*100) },
            { color:'var(--accent)', label:'Net', amount:net, pct: null }
          ];
      legend.innerHTML = legendRows.map(row => `<div class="donut-legend-row">
        <div style="width:10px;height:10px;border-radius:50%;background:${row.color};flex-shrink:0"></div>
        <div style="font-size:13px;color:var(--text-2);flex:1">${row.label}</div>
        <div style="font-size:13px;font-weight:700;color:${row.label==='Net'?(net>=0?'var(--green)':'var(--red)'):'var(--text-1)'}">${fmtMoney(row.amount)}</div>
        <div style="font-size:11px;color:var(--text-3)">${row.pct == null ? '' : row.pct + '%'}</div>
      </div>`).join('');
    }
    // Keep legacy IDs working if still referenced
    const dlInc = document.getElementById('dl-inc'); if(dlInc) dlInc.textContent = fmtMoney(tInc);
    const dlExp = document.getElementById('dl-exp'); if(dlExp) dlExp.textContent = fmtMoney(tExp);
    const netEl = document.getElementById('dl-net');
    if (netEl) { netEl.textContent = fmtMoney(net); netEl.style.color = net >= 0 ? 'var(--green)' : 'var(--red)'; }
  } catch(e) {}
  // Cat bars (top 8 for overview)
  renderCatBars('cat-bars', catTotals, tExp, 8);
  // Budget progress
  renderBudgetProgress();

  // ── Theme-aware colors ──
  const isDark = uiIsDark();
  const tickColor = isDark ? 'rgba(255,255,255,0.4)' : 'rgba(0,0,0,0.4)';
  const gridColor = isDark ? 'rgba(255,255,255,0.05)' : 'rgba(0,0,0,0.06)';
  const legendColor = isDark ? 'rgba(255,255,255,0.5)' : 'rgba(0,0,0,0.5)';

  // Member-filtered basis so the family filter applies to every chart
  let allTxs = S.txs;
  if (S.family && S.memberFilter !== 'all') allTxs = allTxs.filter(t => t.memberId === S.memberFilter);

  // ── Monthly trend · last 6 months ──
  const monthlyEl = document.getElementById('monthly-trend-chart');
  if (monthlyEl) {
    const months = [];
    const now = new Date();
    for (let i = 5; i >= 0; i--) {
      const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
      months.push({ label: d.toLocaleString('default',{month:'short'}), year: d.getFullYear(), month: d.getMonth() });
    }
    const expData = months.map(m => {
      return allTxs.filter(t => t.type==='expense' && parseDate(t.date).getFullYear()===m.year && parseDate(t.date).getMonth()===m.month)
        .reduce((s,t) => s+parseFloat(t.amount||0), 0);
    });
    const incData = months.map(m => {
      return allTxs.filter(t => t.type==='income' && parseDate(t.date).getFullYear()===m.year && parseDate(t.date).getMonth()===m.month)
        .reduce((s,t) => s+parseFloat(t.amount||0), 0);
    });
    if (S.lineChart2) { S.lineChart2.destroy(); S.lineChart2 = null; }
    S.lineChart2 = new Chart(monthlyEl, {
      type: 'line',
      data: {
        labels: months.map(m => m.label),
        datasets: [
          { label: 'Expenses', data: expData, borderColor: themeChartColors().exp, backgroundColor: _hexA(themeChartColors().exp, 0.08), tension: 0.4, fill: true, pointRadius: 3, pointBackgroundColor: themeChartColors().exp },
          ...(S.trackIncome ? [{ label: 'Income', data: incData, borderColor: themeChartColors().inc, backgroundColor: _hexA(themeChartColors().inc, 0.08), tension: 0.4, fill: true, pointRadius: 3, pointBackgroundColor: themeChartColors().inc }] : [])
        ]
      },
      options: {
        responsive: true, maintainAspectRatio: false, interaction: { mode:'index', intersect:false },
        plugins: { legend: { display: true, position:'bottom', labels:{ color: legendColor, font:{size:11}, boxWidth:10, padding:12 } } },
        scales: {
          x: { grid:{ display:false }, ticks:{ color:tickColor, font:{size:10} }, border:{display:false} },
          y: { grid:{ color:gridColor }, ticks:{ color:tickColor, font:{size:10}, callback: v => S.cur.sym+v.toLocaleString() }, border:{display:false} }
        }
      }
    });
  }

  // ── Spending by day of week ──
  const dowEl = document.getElementById('dow-chart');
  if (dowEl) {
    const days = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
    const totals = [0,0,0,0,0,0,0], counts = [0,0,0,0,0,0,0];
    allTxs.filter(t=>t.type==='expense').forEach(t => {
      const d = parseDate(t.date).getDay();
      totals[d] += parseFloat(t.amount||0); counts[d]++;
    });
    const avgs = totals.map((t,i) => counts[i] ? Math.round(t/counts[i]) : 0);
    if (S.dowChart) { S.dowChart.destroy(); S.dowChart = null; }
    const _acc = getComputedStyle(document.documentElement).getPropertyValue('--accent').trim() || '#6C5CE7';
    S.dowChart = new Chart(dowEl, {
      type: 'bar',
      data: {
        labels: days,
        datasets: [{ data: avgs, backgroundColor: avgs.map((_,i)=> i===new Date().getDay() ? _acc : _acc + '4D'), borderRadius: 6, borderSkipped: false }]
      },
      options: {
        responsive: true, maintainAspectRatio: false,
        plugins: { legend:{display:false} },
        scales: {
          x: { grid:{display:false}, ticks:{color:tickColor,font:{size:10}}, border:{display:false} },
          y: { grid:{color:gridColor}, ticks:{color:tickColor,font:{size:10}, callback: v=>S.cur.sym+v}, border:{display:false} }
        }
      }
    });
  }
}

function renderTrendsTab(txs, tInc, tExp, lbl, days) {
  if ((S.analyticsTab||'overview') !== 'trends') return;
  try {
    // Group by period bucket
    const now = new Date();
    let buckets = {};
    // Custom periods (a past week / month / year) reuse the same bucket shapes as
    // their live counterparts, just anchored on the custom range instead of today.
    const _cust = (S.period === 'custom' && S.periodCustom) ? S.periodCustom : null;
    const _range = _cust ? _customRange(_cust) : null;
    const _mode = _cust ? (_cust.type === 'lastweek' ? 'week' : (_cust.type === 'year' ? 'year' : 'month')) : S.period;
    const getKey = (d) => {
      if (_mode === 'week') return d.toLocaleDateString('en-US',{weekday:'short'});
      if (_mode === 'month') return d.getDate()+'';
      if (_mode === '3m') { const wk = Math.ceil(d.getDate()/7); return d.toLocaleDateString('en-US',{month:'short'})+' W'+wk; }
      if (_mode === 'year') return d.toLocaleDateString('en-US',{month:'short'});
      return d.toLocaleDateString('en-US',{month:'short',year:'2-digit'});
    };

    // Seed buckets in true chronological order (and with no gaps) — otherwise the
    // week chart ran Sun→Sat regardless of today, and month/year lines skipped
    // empty days/months, distorting the shape.
    if (_mode === 'week') {
      const base = _range ? new Date(_range[0]) : null;
      for (let i = 0; i < 7; i++) {
        const d = base ? new Date(base.getFullYear(), base.getMonth(), base.getDate() + i)
                       : new Date(now.getFullYear(), now.getMonth(), now.getDate() - (6 - i));
        buckets[d.toLocaleDateString('en-US',{weekday:'short'})] = {inc:0,exp:0};
      }
    } else if (_mode === 'month') {
      const _y = _range ? new Date(_range[0]).getFullYear() : now.getFullYear();
      const _m = _range ? new Date(_range[0]).getMonth() : now.getMonth();
      const _isCur = _y === now.getFullYear() && _m === now.getMonth();
      const _last = _isCur ? now.getDate() : new Date(_y, _m + 1, 0).getDate();
      for (let i = 1; i <= _last; i++) buckets[i+''] = {inc:0,exp:0};
    } else if (_mode === 'year') {
      const _y = _range ? new Date(_range[0]).getFullYear() : now.getFullYear();
      const _lastM = (_y === now.getFullYear()) ? now.getMonth() : 11;
      for (let m = 0; m <= _lastM; m++) buckets[new Date(_y, m, 1).toLocaleDateString('en-US',{month:'short'})] = {inc:0,exp:0};
    }

    [...txs].sort((a,b) => dateMs(a.date) - dateMs(b.date)).forEach(t => {
      const d = parseDate(t.date); const k = getKey(d);
      if (!buckets[k]) buckets[k] = {inc:0,exp:0};
      const amt = (parseFloat(t.amount)||0);
      if (t.type === 'income') buckets[k].inc += amt; else buckets[k].exp += amt;
    });

    const labels = Object.keys(buckets);
    const incData = labels.map(k => buckets[k].inc);
    const expData = labels.map(k => buckets[k].exp);

    const lCtx = document.getElementById('line-chart');
    if (S.lineChart) S.lineChart.destroy();
    const accent = getComputedStyle(document.documentElement).getPropertyValue('--accent').trim() || '#6C5CE7';
    const _dk = uiIsDark();
    const _tick = _dk ? 'rgba(255,255,255,0.45)' : 'rgba(0,0,0,0.45)';
    const _grid = _dk ? 'rgba(255,255,255,0.05)' : 'rgba(0,0,0,0.06)';
    S.lineChart = new Chart(lCtx, {
      type: 'line',
      data: {
        labels,
        datasets: [
          ...(S.trackIncome ? [{ label:'Income', data: incData, borderColor:themeChartColors().inc, backgroundColor:_hexA(themeChartColors().inc, 0.10), fill:true, tension:0.35, pointRadius:3, pointBackgroundColor:themeChartColors().inc, borderWidth:2 }] : []),
          { label:'Expenses', data: expData, borderColor:themeChartColors().exp, backgroundColor:_hexA(themeChartColors().exp, 0.08), fill:true, tension:0.35, pointRadius:3, pointBackgroundColor:themeChartColors().exp, borderWidth:2 }
        ]
      },
      options: {
        responsive:true, maintainAspectRatio:false,
        plugins: { legend:{ display:true, labels:{ color: getComputedStyle(document.documentElement).getPropertyValue('--text-2').trim()||'#9CA1AB', boxWidth:10, boxHeight:10, borderRadius:5, useBorderRadius:true, font:{size:11,weight:'600'} } }, tooltip:{ callbacks:{ label: ctx => ' ' + fmtMoney(ctx.raw) } } },
        scales: { x:{ ticks:{color:_tick,font:{size:10}}, grid:{color:_grid} }, y:{ ticks:{color:_tick,font:{size:10}, callback: v => fmtMoney(v)}, grid:{color:_grid} } }
      }
    });

    document.getElementById('a-trend-lbl').textContent = lbl;
    const trendsTitle = document.getElementById('trends-inc-title');
    if (trendsTitle) trendsTitle.textContent = S.trackIncome ? 'Income vs expenses' : 'Expenses over time';

    // Bar chart (daily spending)
    const bCtx = document.getElementById('bar-chart');
    if (S.barChart) S.barChart.destroy();
    const daily = {};
    [...txs].filter(t=>t.type==='expense').sort((a,b) => dateMs(a.date) - dateMs(b.date)).forEach(t => {
      const dStr = fmtDateShort(t.date);
      daily[dStr] = (daily[dStr]||0) + (parseFloat(t.amount)||0);
    });
    const dKeys = Object.keys(daily).slice(-14);
    S.barChart = new Chart(bCtx, {
      type:'bar',
      data:{ labels:dKeys, datasets:[{ data:dKeys.map(k=>daily[k]), backgroundColor: accent+'CC', borderRadius:5 }] },
      options:{ responsive:true, maintainAspectRatio:false, plugins:{ legend:{display:false}, tooltip:{callbacks:{label:ctx=>' '+fmtMoney(ctx.raw)}} }, scales:{ x:{ticks:{color:_tick,font:{size:10}},grid:{color:_grid}}, y:{ticks:{color:_tick,font:{size:10},callback:v=>fmtMoney(v)},grid:{color:_grid}} } }
    });

    // Period comparison
    renderPeriodComparison(txs);
  } catch(e) { console.error('trends error', e); }
}

function renderPeriodComparison(txs) {
  const wrap = document.getElementById('trend-compare');
  if (!wrap) return;
  // Get previous period txs
  const { txs: prevTxs } = getPrevPeriodTxs();
  const curr = { inc:0, exp:0 }; const prev = { inc:0, exp:0 };
  txs.forEach(t => { if(t.type==='income') curr.inc+=(parseFloat(t.amount)||0); else curr.exp+=(parseFloat(t.amount)||0); });
  prevTxs.forEach(t => { if(t.type==='income') prev.inc+=(parseFloat(t.amount)||0); else prev.exp+=(parseFloat(t.amount)||0); });

  const rows = [
    { label: 'Spending', curr: curr.exp, prev: prev.exp, lowerBetter: true },
    ...(S.trackIncome ? [
      { label: 'Income', curr: curr.inc, prev: prev.inc, lowerBetter: false },
      { label: 'Net savings', curr: curr.inc - curr.exp, prev: prev.inc - prev.exp, lowerBetter: false }
    ] : [])
  ];
  wrap.innerHTML = rows.map(r => {
    const delta = r.prev !== 0 ? ((r.curr - r.prev) / Math.abs(r.prev) * 100) : 0;
    const dir = delta > 2 ? 'up' : delta < -2 ? 'down' : 'neutral';
    const good = r.lowerBetter ? dir === 'down' : dir === 'up';
    const arrow = delta > 2 ? '↑' : delta < -2 ? '↓' : '→';
    const cls = (dir === 'neutral') ? 'neutral' : (good ? 'up' : 'down');
    return `<div class="trend-compare-row">
      <div class="tc-label">${r.label}</div>
      <div class="tc-vals">
        <div class="tc-curr">${fmtMoney(r.curr)}</div>
        ${r.prev !== 0 ? `<div class="tc-delta ${cls}">${arrow} ${Math.abs(Math.round(delta))}% vs prev</div>` : '<div class="tc-delta neutral">No prev data</div>'}
      </div>
    </div>`;
  }).join('');
}

function getPrevPeriodTxs() {
  const now = new Date(); let prevStart = 0; let prevEnd = 0;
  const sod = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  if (S.period === 'week') { prevEnd = sod - 6*86400000; prevStart = prevEnd - 7*86400000; }
  else if (S.period === 'month') { const m = new Date(now.getFullYear(), now.getMonth(), 1); prevEnd = m.getTime(); prevStart = new Date(now.getFullYear(), now.getMonth()-1, 1).getTime(); }
  else if (S.period === '3m') { prevEnd = new Date(now.getFullYear(), now.getMonth()-3, now.getDate()).getTime(); prevStart = new Date(now.getFullYear(), now.getMonth()-6, now.getDate()).getTime(); }
  else if (S.period === 'year') { prevEnd = new Date(now.getFullYear(), 0, 1).getTime(); prevStart = new Date(now.getFullYear()-1, 0, 1).getTime(); }
  else if (S.period === 'custom' && S.periodCustom) { const r = _customRange(S.periodCustom); prevEnd = r[0]; prevStart = r[0] - (r[1] - r[0]); }
  else { return { txs: [] }; }
  let txs = S.txs.filter(t => { const ts = dateMs(t.date); return ts >= prevStart && ts < prevEnd; });
  if (S.family && S.memberFilter !== 'all') txs = txs.filter(t => t.memberId === S.memberFilter);
  return { txs };
}

function renderBreakdownTab(txs, tInc, tExp, catTotals, incTotals, acctTotals) {
  // ── Category pie chart ──
  try {
    const pieEl = document.getElementById('cat-pie-chart');
    if (pieEl) {
      const cats = Object.entries(catTotals).sort((a,b)=>b[1]-a[1]).slice(0,8);
      const COLORS = themePalette(8);
      if (S.catPieChart) { S.catPieChart.destroy(); S.catPieChart = null; }
      if (cats.length === 0) {
        S.catPieChart = new Chart(pieEl, { type:'doughnut', data:{ datasets:[{ data:[1], backgroundColor:[getComputedStyle(document.documentElement).getPropertyValue('--bg-4').trim() || '#1D2024'], borderWidth:0 }] }, options:{ responsive:true, maintainAspectRatio:false, cutout:'72%', plugins:{tooltip:{enabled:false}} }});
      } else {
        S.catPieChart = new Chart(pieEl, {
          type:'doughnut',
          data:{ labels: cats.map(c=>c[0]), datasets:[{ data: cats.map(c=>c[1]), backgroundColor: cats.map((_,i)=>COLORS[i%COLORS.length]), borderWidth:2, borderColor: getComputedStyle(document.documentElement).getPropertyValue('--bg-2').trim()||'#0D0F11' }] },
          options:{ responsive:true, maintainAspectRatio:false, cutout:'72%', plugins:{ legend:{display:false}, tooltip:{ callbacks:{ label: ctx=>' '+fmtMoney(ctx.raw)+' ('+Math.round(ctx.raw/tExp*100)+'%)' } } } }
        });
      }
      const legend = document.getElementById('cat-pie-legend');
      if (legend) {
        legend.innerHTML = cats.map(([cat,amt],i) => `<div style="display:flex;align-items:center;gap:8px;padding:5px 0;border-bottom:1px solid var(--line)">
          <div style="width:10px;height:10px;border-radius:50%;background:${COLORS[i%COLORS.length]};flex-shrink:0"></div>
          <div style="font-size:13px;color:var(--text-2);flex:1">${esc(cat)}</div>
          <div style="font-size:13px;font-weight:700">${fmtMoney(amt)}</div>
          <div style="font-size:11px;color:var(--text-3)">${tExp>0?Math.round(amt/tExp*100):0}%</div>
        </div>`).join('');
      }
    }
  } catch(e) { console.warn('pie chart:', e); }

  if ((S.analyticsTab||'overview') !== 'breakdown') return;
  const view = S.breakdownView || 'category';
  // Quick facts on top · the pie moved to the bottom of the tab
  const statsEl = document.getElementById('bd-stats');
  if (statsEl) {
    const topCat = Object.entries(catTotals).sort((a,b) => b[1]-a[1])[0];
    const expTxs = txs.filter(t => t.type === 'expense');
    const biggest = expTxs.slice().sort((a,b) => (parseFloat(b.amount)||0)-(parseFloat(a.amount)||0))[0];
    const avgTx = expTxs.length ? tExp / expTxs.length : 0;
    statsEl.innerHTML = `
      <div class="stat-box"><div class="sb-label">Top category</div><div class="sb-val" style="font-size:17px">${topCat ? esc(topCat[0]) : '—'}</div><div class="sb-sub">${topCat ? fmtMoney(topCat[1]) : 'No spending yet'}</div></div>
      <div class="stat-box"><div class="sb-label">Largest expense</div><div class="sb-val" style="font-size:17px">${biggest ? fmtMoney(biggest.amount) : '—'}</div><div class="sb-sub">${biggest ? esc(biggest.description || biggest.category) : ''}</div></div>
      <div class="stat-box"><div class="sb-label">Avg expense</div><div class="sb-val" style="font-size:17px">${fmtMoney(avgTx)}</div><div class="sb-sub">${expTxs.length} expense${expTxs.length===1?'':'s'}</div></div>
      <div class="stat-box"><div class="sb-label">Categories used</div><div class="sb-val" style="font-size:17px">${Object.keys(catTotals).length}</div><div class="sb-sub">this period</div></div>`;
  }
  renderCatBars('bd-cat-bars', catTotals, tExp, 20);
  // Account bars
  const acctWrap = document.getElementById('bd-acct-bars');
  if (acctWrap) {
    const arr = Object.keys(acctTotals).map(k => ({name:k, amt:acctTotals[k]})).sort((a,b)=>b.amt-a.amt);
    if (!arr.length) { acctWrap.innerHTML = '<div class="no-data">No transactions with accounts</div>'; }
    else {
      const max = arr[0].amt;
      acctWrap.innerHTML = arr.map(a => {
        const pct = max > 0 ? Math.round(a.amt/max*100) : 0;
        const acctObj = (S.accts||[]).find(x=>x.name===a.name)||{};
        const icon = getAcctIcon(acctObj.type);
        return `<div class="cat-bar"><div class="cbi">${icon}</div><div class="cbi-info"><div class="cbi-top"><div class="cbi-name">${esc(a.name)}</div><div class="cbi-pct">${pct}%</div></div><div class="bar-track"><div class="bar-fill" style="width:${pct}%"></div></div></div><div class="cbi-amt">${fmtMoney(a.amt)}</div></div>`;
      }).join('');
    }
  }
  // Income breakdown
  const incWrap = document.getElementById('bd-inc-bars');
  if (incWrap) {
    renderCatBarsInto(incWrap, incTotals, tInc, 20, true);
  }
  // Member bars (family)
  const memWrap = document.getElementById('member-bars');
  if (memWrap) { renderMemberBars(txs, tExp, memWrap); }
}

function renderCatBars(elId, catObj, totalExp, limit) {
  const wrap = document.getElementById(elId);
  if (!wrap) return;
  renderCatBarsInto(wrap, catObj, totalExp, limit, false);
}

function renderCatBarsInto(wrap, catObj, total, limit, isIncome) {
  const arr = Object.keys(catObj).map(k => ({cat:k, amt:catObj[k]})).sort((a,b)=>b.amt-a.amt).slice(0, limit||8);
  if (!arr.length) { wrap.innerHTML = '<div class="no-data">No data yet</div>'; return; }
  wrap.innerHTML = arr.map(c => {
    const pct = total > 0 ? Math.round(c.amt/total*100) : 0;
    return `<div class="cat-bar"><div class="cbi">${getCatIcon(c.cat, isIncome)}</div><div class="cbi-info"><div class="cbi-top"><div class="cbi-name">${esc(c.cat)}</div><div class="cbi-pct">${pct}%</div></div><div class="bar-track"><div class="bar-fill${isIncome?' bar-fill-inc':''}"></div></div></div><div class="cbi-amt">${fmtMoney(c.amt)}</div></div>`;
  }).join('');
  // Animate bar widths after paint
  requestAnimationFrame(() => {
    wrap.querySelectorAll('.bar-fill').forEach((el,i) => {
      const pct = total > 0 ? Math.round(arr[i].amt/total*100) : 0;
      el.style.width = pct + '%';
    });
  });
}

function renderMemberBars(aTxs, tExp, wrapEl) {
  if (!wrapEl) wrapEl = document.getElementById('member-bars');
  if (!S.family || !S.members || !S.members.length) { if(wrapEl) wrapEl.innerHTML = '<div class="no-data">Family mode only</div>'; return; }
  const totals = {};
  aTxs.forEach(t => { if (t.type==='expense' && t.memberId) totals[t.memberId] = (totals[t.memberId]||0) + (parseFloat(t.amount)||0); });
  const arr = S.members.map(m => ({m, amt: totals[m.id]||0})).sort((a,b)=>b.amt-a.amt);
  // #3 · show Joint / Shared spending as its own bucket when present
  const jointAmt = totals[JOINT_ID] || 0;
  if (jointAmt > 0) arr.push({ m: { id: JOINT_ID, name: JOINT_LABEL, color: '#6C5CE7' }, amt: jointAmt });
  if (!arr.some(x => x.amt > 0)) { wrapEl.innerHTML = '<div class="no-data">No spending recorded yet</div>'; return; }
  wrapEl.innerHTML = arr.map(x => {
    const pct = tExp > 0 ? Math.round(x.amt/tExp*100) : 0;
    return `<div class="cat-bar"><div class="cbi" style="background:${esc(x.m.color)}22;color:${esc(x.m.color)}">${esc((x.m.name||'?').charAt(0).toUpperCase())}</div><div class="cbi-info"><div class="cbi-top"><div class="cbi-name">${esc(x.m.name)}</div><div class="cbi-pct">${pct}%</div></div><div class="bar-track"><div class="bar-fill" style="width:${pct}%;background:${esc(x.m.color)}"></div></div></div><div class="cbi-amt">${fmtMoney(x.amt)}</div></div>`;
  }).join('');
}

function renderBudgetProgress() {
  const box = document.getElementById('budget-progress-box');
  const wrap = document.getElementById('budget-progress');
  if (!S.budgets || !S.budgets.length) { if(box) box.style.display = 'none'; return; }
  if(box) box.style.display = 'block';
  const now = new Date();
  const monthStart = new Date(now.getFullYear(), now.getMonth(), 1).getTime();
  const spend = {};
  S.txs.forEach(t => { if (t.type==='expense' && dateMs(t.date) >= monthStart) spend[t.category] = (spend[t.category]||0) + (parseFloat(t.amount)||0); });
  wrap.innerHTML = S.budgets.map(b => {
    const spent = spend[b.category]||0;
    const pct = b.limit > 0 ? Math.min(100, Math.round(spent/b.limit*100)) : 0;
    const over = spent > b.limit;
    return `<div class="budget-row">
      <div class="budget-top"><div class="budget-name">${getCatIcon(b.category,false)} ${esc(b.category)}</div><div class="budget-amt ${over?'over':''}">${fmtMoney(spent)} / ${fmtMoney(b.limit)}</div></div>
      <div class="budget-track"><div class="budget-fill ${over?'over':''}" style="width:${pct}%"></div></div>
    </div>`;
  }).join('');
}

function setAnalyticsTab(tab, el) {
  S.analyticsTab = tab;
  document.querySelectorAll('.an-tab').forEach(t => t.classList.remove('active'));
  if (el) el.classList.add('active');
  document.querySelectorAll('.an-panel').forEach(p => p.style.display = 'none');
  const panel = document.getElementById('an-' + tab);
  if (panel) panel.style.display = '';
  renderAnalytics();
}

function setBreakdownView(view, el) {
  S.breakdownView = view;
  document.querySelectorAll('.an-sub-tab').forEach(t => t.classList.remove('active'));
  if (el) el.classList.add('active');
  ['bd-cat-panel','bd-acct-panel','bd-inc-panel','bd-person-panel'].forEach(id => {
    const el2 = document.getElementById(id); if(el2) el2.style.display = 'none';
  });
  const panelMap = { category:'bd-cat-panel', account:'bd-acct-panel', income:'bd-inc-panel', person:'bd-person-panel' };
  const p = document.getElementById(panelMap[view]); if(p) p.style.display = '';
  renderAnalytics();
}

function updateCharts() { /* no-op: charts rendered in renderAnalytics now */ }



// ═══════════════════════════════════════════
//  SETTINGS RENDERING
// ═══════════════════════════════════════════
function renderSettings() {
  const _el = id => document.getElementById(id);
  if (S.family && S.member) {
    if (_el('profile-name')) _el('profile-name').textContent = S.member.name;
    applyAvatarTo(_el('profile-av'), (S.member.name||'?').charAt(0).toUpperCase(), S.member.color);
    if (_el('profile-sub')) _el('profile-sub').textContent = (S.familyName || 'Family') + ' · ' + (S.mode === 'sheets' ? 'Sheets Connected ✓' : 'Local Memory Active');
  } else {
    if (_el('profile-name')) _el('profile-name').textContent = S.name;
    applyAvatarTo(_el('profile-av'), S.name.charAt(0).toUpperCase());
    if (_el('profile-sub')) _el('profile-sub').textContent = S.mode==='sheets'?'Sheets Connected ✓':S.mode==='firebase'?'🔥 Firebase Connected ✓':'Local Memory Active';
  }
  if (_el('pref-name')) _el('pref-name').textContent = S.name;
  const tbSub = document.getElementById('settings-topbar-sub');
  if (tbSub) tbSub.textContent = S.mode === 'sheets' ? 'Google Sheets · connected' : S.mode === 'firebase' ? 'Google · connected' : 'Stored on this device';
  const simpleSub = document.getElementById('simple-mode-sub');
  if (simpleSub) simpleSub.textContent = S.simpleMode ? 'On' : 'Off';
  const tiSub = document.getElementById('track-income-sub');
  if (tiSub) tiSub.textContent = S.trackIncome ? 'On' : 'Off';
  const dmTitle = document.getElementById('data-mode-title');
  const dmSub = document.getElementById('data-mode-sub');
  if (dmTitle && dmSub) {
    if (S.mode === 'firebase') { dmTitle.textContent = 'Google account'; dmSub.textContent = 'Connected · tap to disconnect'; }
    else if (S.mode === 'sheets') { dmTitle.textContent = 'Google Sheet'; dmSub.textContent = 'Connected · tap to disconnect'; }
    else { dmTitle.textContent = 'Connect cloud storage'; dmSub.textContent = 'Your data currently lives only on this device'; }
  }
  const emojiSub = document.getElementById('emoji-toggle-sub');
  if (emojiSub) emojiSub.textContent = document.body.hasAttribute('data-noemoji') ? 'Off (minimal initials)' : 'On';
  document.getElementById('pref-currency').textContent = `${S.cur.code} (${S.cur.sym})`;

  const acctWrap = document.getElementById('acct-settings-list');
  const _acctBtn = document.getElementById('acct-add-btn');
  if (!S.accts || !S.accts.length) {
    if (_acctBtn) _acctBtn.style.display = 'none';
    acctWrap.innerHTML = '<div class="srow" onclick="openAcctModal()"><div class="srow-ico">&#127974;</div><div class="srow-body"><div class="srow-title" style="font-size:13px">Add your first account</div></div><div class="srow-arrow" style="font-size:18px;font-weight:600;color:var(--accent)">+</div></div>';
  } else {
    if (_acctBtn) _acctBtn.style.display = '';
    acctWrap.innerHTML = S.accts.map(a => `
      <div class="srow" onclick="openAcctModal('${esc(a.id)}')">
        <div class="srow-ico">${getAcctIcon(a.type)}</div>
        <div class="srow-body"><div class="srow-title">${esc(a.name)}</div><div class="srow-sub">${esc(a.type)}</div></div>
        <div style="font-size:14px;font-weight:700;margin-right:8px">${fmtMoney(a.balance)}</div>
        <div class="srow-arrow"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M9 18l6-6-6-6"/></svg></div>
      </div>`).join('');
  }

  const recWrap = document.getElementById('income-list');
  const _recBtn = document.getElementById('rec-add-btn');
  if (!S.recurring || !S.recurring.length) {
    if (_recBtn) _recBtn.style.display = 'none';
    recWrap.innerHTML = '<div class="settings-group"><div class="srow" onclick="openRecurringModal()"><div class="srow-ico">&#128257;</div><div class="srow-body"><div class="srow-title" style="font-size:13px">Add a recurring bill or income</div></div><div class="srow-arrow" style="font-size:18px;font-weight:600;color:var(--accent)">+</div></div></div>';
  } else {
    if (_recBtn) _recBtn.style.display = '';
    recWrap.innerHTML = S.recurring.map(r => `
      <div class="srow" onclick="openRecurringModal('${esc(r.id)}')">
        <div class="srow-ico">${getCatIcon(r.category, r.type === 'income')}</div>
        <div class="srow-body"><div class="srow-title">${esc(r.name)}</div><div class="srow-sub">${esc(r.frequency)} · ${fmtMoney(r.amount)}</div></div>
        <div class="srow-arrow"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M9 18l6-6-6-6"/></svg></div>
      </div>`).join('');
  }

  const budgetWrap = document.getElementById('budget-list');
  const _budBtn = document.getElementById('budget-add-btn');
  if (!S.budgets || !S.budgets.length) {
    if (_budBtn) _budBtn.style.display = 'none';
    budgetWrap.innerHTML = '<div class="settings-group"><div class="srow" onclick="openBudgetModal()"><div class="srow-ico">&#127919;</div><div class="srow-body"><div class="srow-title" style="font-size:13px">Set a monthly budget</div></div><div class="srow-arrow" style="font-size:18px;font-weight:600;color:var(--accent)">+</div></div></div>';
  } else {
    if (_budBtn) _budBtn.style.display = '';
    budgetWrap.innerHTML = S.budgets.map(b => `
      <div class="srow" onclick="openBudgetModal('${esc(b.id)}')">
        <div class="srow-ico">${getCatIcon(b.category, false)}</div>
        <div class="srow-body"><div class="srow-title">${esc(b.category)}</div><div class="srow-sub">${fmtMoney(b.limit)} / month</div></div>
        <div class="srow-arrow"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M9 18l6-6-6-6"/></svg></div>
      </div>`).join('');
  }

  if (S.family) {
    const memWrap = document.getElementById('member-list');
    memWrap.innerHTML = (S.members||[]).map(m => `
      <div class="srow" style="cursor:default">
        <div class="member-avatar" style="width:36px;height:36px;font-size:14px;background:${esc(m.color)};flex-shrink:0">${esc((m.name||'?').charAt(0).toUpperCase())}</div>
        <div class="srow-body" style="margin-left:10px"><div class="srow-title">${esc(m.name)}${S.member && S.member.id===m.id ? ' (you)' : ''}</div></div>
        ${S.member && S.member.id===m.id ? '' : `<div class="del-ico" onclick="delMember('${esc(m.id)}', event)"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round"><path d="M3 6h18M19 6v14a2 2 0 01-2 2H7a2 2 0 01-2-2V6m3 0V4a2 2 0 012-2h4a2 2 0 012 2v2"/></svg></div>`}
      </div>`).join('');
  }

  renderCatManager(); // Categories section is always visible now
}

function delMember(id, e) {
  if (e) e.stopPropagation();
  showCustomConfirm('Remove Member', "Remove this family member? Their past transactions remain.", 'Remove', async () => {
    toast('🗑 Member removed');
    await apiPost({action:'deleteMember', id});
  });
}



// ═══════════════════════════════════════════
//  PREFERENCES & ACCOUNT SELECTS
// ═══════════════════════════════════════════
function updatePrefUI() {
  _updateAccentUI();
  const compSub = document.getElementById('compact-mode-sub');
  if (compSub) compSub.textContent = document.body.hasAttribute('data-compact') ? 'On' : 'Off';
  const compTog = document.getElementById('compact-mode-toggle');
  if (compTog) compTog.classList.toggle('on', document.body.hasAttribute('data-compact'));
  const savedCorner = localStorage.getItem('prizm_corner') || 'round';
  document.querySelectorAll('.corner-opt').forEach(b => b.classList.toggle('active', b.dataset.corner === savedCorner));
  document.querySelectorAll('#hero-period-setting .fc').forEach(b => b.classList.toggle('active', b.dataset.hp === (S.heroPeriod || 'month')));
  document.getElementById('pref-name').textContent = S.name;
  const simpleSub = document.getElementById('simple-mode-sub');
  if (simpleSub) simpleSub.textContent = S.simpleMode ? 'On' : 'Off';
  const tiSub = document.getElementById('track-income-sub');
  if (tiSub) tiSub.textContent = S.trackIncome ? 'On' : 'Off';
  const dmTitle = document.getElementById('data-mode-title');
  const dmSub = document.getElementById('data-mode-sub');
  if (dmTitle && dmSub) {
    if (S.mode === 'firebase') { dmTitle.textContent = 'Google account'; dmSub.textContent = 'Connected · tap to disconnect'; }
    else if (S.mode === 'sheets') { dmTitle.textContent = 'Google Sheet'; dmSub.textContent = 'Connected · tap to disconnect'; }
    else { dmTitle.textContent = 'Connect cloud storage'; dmSub.textContent = 'Your data currently lives only on this device'; }
  }
  const emojiSub = document.getElementById('emoji-toggle-sub');
  if (emojiSub) emojiSub.textContent = document.body.hasAttribute('data-noemoji') ? 'Off (minimal initials)' : 'On';
  document.getElementById('pref-currency').textContent = `${S.cur.code} (${S.cur.sym})`;
  // Show just the base symbol (strip country prefix like A$ → $)
  const baseSym = S.cur.sym.replace(/^[A-Za-z]+/, '');
  document.getElementById('add-symbol').textContent = baseSym || S.cur.sym;
}

// #24 · single shared builder. #6 · preserve the current selection across rebuilds
// so a background sync (renderAll) doesn't silently wipe the account the user picked
// while the Add/Edit modal is open.
function _buildAcctSelect(selectEl) {
  if (!selectEl) return;
  // Never rebuild while the user is interacting with this select's modal —
  // a background sync re-rendering the options resets the open picker on mobile.
  const ov = selectEl.closest('.overlay');
  if (ov && ov.classList.contains('open')) return;
  const prev = selectEl.value;
  if (!S.accts || !S.accts.length) selectEl.innerHTML = '<option value="">No accounts yet</option>';
  else selectEl.innerHTML = '<option value="">Select account</option>' + S.accts.map(a => `<option value="${esc(a.name)}">${getAcctIcon(a.type)} ${esc(a.name)}</option>`).join('');
  if (prev && Array.from(selectEl.options).some(o => o.value === prev)) selectEl.value = prev;
}
function updateAcctSelect() { _buildAcctSelect(document.getElementById('add-acct')); }
function updateRecAcctSelect() { _buildAcctSelect(document.getElementById('rec-acct')); }

// ═══════════════════════════════════════════
//  CATEGORY GRIDS
// ═══════════════════════════════════════════
function getCatEmoji(name, isIncome) {
  const customEmojis = safeParse(localStorage.getItem(isIncome ? 'prizm_inc_emojis' : 'prizm_cat_emojis'), {});
  if (customEmojis[name]) return customEmojis[name];
  // Custom categories store their chosen icon on the saved category list (.e),
  // which is where addCat/updateCatEmoji write it. Look there before falling back
  // to the built-in defaults, otherwise custom categories always showed 💡.
  const saved = getCatList(isIncome).find(c => c.n === name);
  if (saved && saved.e) return saved.e;
  const list = isIncome ? INC_CATS : CATS;
  const f = list.find(c => c.n === name);
  return f ? f.e : (isIncome ? '✨' : '💡');
}

function buildCatGrid() {
  const list = getCatList(S.type === 'income');
  const customNames = safeParse(localStorage.getItem(S.type === 'income' ? 'prizm_inc_names' : 'prizm_cat_names'), {});
  if (!list.find(c => c.n === S.cat) && !customNames[S.cat]) S.cat = list[0].n;
  document.getElementById('cat-grid').innerHTML = list.map(c => {
    const displayName = customNames[c.n] || c.n;
    const emoji = getCatEmoji(c.n, S.type === 'income');
    return `<div class="cat-btn ${c.n === S.cat ? 'sel' : ''}" onclick="selectCat('${esc(jsq(c.n))}', this)">
      <div class="cat-e">${esc(emoji)}</div><div class="cat-n">${esc(displayName)}</div>
    </div>`;
  }).join('') + `<div class="cat-btn" onclick="quickAddCat()" style="border-style:dashed;opacity:.85"><div class="cat-e">＋</div><div class="cat-n">New</div></div>`;
}

function quickAddCat() {
  const name = (prompt('New category name:') || '').trim();
  if (!name) return;
  const isInc = S.type === 'income';
  const list = getCatList(isInc);
  if (!list.find(c => c.n === name)) { list.push({ n: name, e: fuzzyCatIcon(name, isInc) }); _saveCatList(isInc, list); }
  S.cat = name;
  buildCatGrid();
}

function selectCat(name, el) {
  S.cat = name;
  document.querySelectorAll('#cat-grid .cat-btn').forEach(b => b.classList.remove('sel'));
  el.classList.add('sel');
}

function buildRecCatGrid() {
  const list = getCatList(S.recType === 'income');
  const customNames = safeParse(localStorage.getItem(S.recType === 'income' ? 'prizm_inc_names' : 'prizm_cat_names'), {});
  // Only re-pick when the current choice isn't valid for this type (e.g. after an
  // expense↔income toggle). Restore the last category chosen for this type if we can,
  // otherwise fall back to a sensible default — never silently snap back to Food.
  if (!list.find(c => c.n === S.recCat)) {
    const remembered = S.recType === 'income' ? S._lastRecCatInc : S._lastRecCatExp;
    const preferred = S.recType === 'income' ? 'Salary' : 'Bills';
    if (remembered && list.find(c => c.n === remembered)) S.recCat = remembered;
    else S.recCat = list.find(c => c.n === preferred) ? preferred : list[0].n;
  }
  document.getElementById('rec-cat-grid').innerHTML = list.map(c => {
    const displayName = customNames[c.n] || c.n;
    const emoji = getCatEmoji(c.n, S.recType === 'income');
    return `<div class="cat-btn ${c.n === S.recCat ? 'sel' : ''}" onclick="selectRecCat('${esc(jsq(c.n))}', this)">
      <div class="cat-e">${esc(emoji)}</div><div class="cat-n">${esc(displayName)}</div>
    </div>`;
  }).join('');
}

function selectRecCat(name, el) {
  S.recCat = name;
  if (S.recType === 'income') S._lastRecCatInc = name; else S._lastRecCatExp = name;
  document.querySelectorAll('#rec-cat-grid .cat-btn').forEach(b => b.classList.remove('sel'));
  el.classList.add('sel');
}

function buildBudgetCatOptions() {
  const sel = document.getElementById('budget-cat');
  if (!sel) return;
  const customNames = safeParse(localStorage.getItem('prizm_cat_names'), {});
  const list = getCatList(false).map(c => ({...c}));
  // Keep the category of the budget being edited selectable even if it was removed
  if (S.editingBudgetId) {
    const b = (S.budgets||[]).find(x => x.id === S.editingBudgetId);
    if (b && b.category && !list.find(c => c.n === b.category)) list.push({ n: b.category, e: getCatIcon(b.category, false) });
  }
  sel.innerHTML = list.map(c => `<option value="${esc(c.n)}">${esc(getCatEmoji(c.n, false) === '💡' && c.e ? c.e : getCatEmoji(c.n, false))} ${esc(customNames[c.n] || c.n)}</option>`).join('');
}

// ═══════════════════════════════════════════
//  NAVIGATION
// ═══════════════════════════════════════════
function goTab(tab) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nb').forEach(n => n.classList.remove('active'));
  document.getElementById('page-' + tab).classList.add('active');
  document.getElementById('nb-' + tab).classList.add('active');
  // Render the page being opened (other pages are skipped during background renders)
  if (tab === 'analytics') renderAnalytics();
  else if (tab === 'transactions') { try { renderAllTx(); } catch(e) {} }
  else if (tab === 'settings') { try { renderSettings(); updatePrefUI(); } catch(e) {} }
}

function setFilter(f, el) {
  S.filter = f;
  document.querySelectorAll('#page-transactions .fc').forEach(c => c.classList.remove('active'));
  if (el) el.classList.add('active');
  renderAllTx();
}

function setHeroPeriod(p) {
  S.heroPeriod = p;
  localStorage.setItem('prizm_hero_period', p);
  document.querySelectorAll('#hero-period-setting .fc').forEach(b => b.classList.toggle('active', b.dataset.hp === p));
  renderHome();
  schedulePrefSync();
}

const _MONTH_NAMES = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
function openPeriodPicker() {
  S._ppYear = S._ppYear || new Date().getFullYear();
  buildPeriodPicker();
  openOverlay('period-overlay');
}
function buildPeriodPicker() {
  const yEl = document.getElementById('pp-years'), mEl = document.getElementById('pp-months');
  if (!yEl || !mEl) return;
  const ys = new Set([new Date().getFullYear()]);
  S.txs.forEach(t => { const y = parseDate(t.date).getFullYear(); if (y > 1970) ys.add(y); });
  const years = [...ys].sort((a,b) => b - a).slice(0, 12);
  yEl.innerHTML = years.map(y =>
    '<div class="fc' + (y === S._ppYear ? ' active' : '') + '" onclick="S._ppYear=' + y + ';buildPeriodPicker()" style="font-size:12px;padding:6px 12px">' + y + '</div>'
  ).join('') + '<div class="fc" onclick="setCustomPeriod({type:\'year\',y:S._ppYear})" style="font-size:12px;padding:6px 12px;border-color:var(--accent);color:var(--accent)">Whole ' + S._ppYear + ' &rarr;</div>';
  const now = new Date();
  mEl.innerHTML = _MONTH_NAMES.map((n, i) => {
    const future = S._ppYear > now.getFullYear() || (S._ppYear === now.getFullYear() && i > now.getMonth());
    return '<div class="fc" style="font-size:12px;padding:9px 0;text-align:center;display:flex;align-items:center;justify-content:center;' + (future ? 'opacity:.35;pointer-events:none;' : '') + '" onclick="setCustomPeriod({type:\'month\',y:S._ppYear,m:' + i + '})">' + n + '</div>';
  }).join('');
}
function setCustomPeriod(p) {
  S.periodCustom = p; S.period = 'custom';
  document.querySelectorAll('#page-analytics .pb').forEach(c => c.classList.remove('active'));
  const pill = document.getElementById('pb-custom');
  if (pill) { pill.classList.add('active'); pill.textContent = _customPeriodLabel(p); }
  closeOverlay('period-overlay');
  renderAnalytics();
}
// Apply a user-picked From → To date range as the analytics period.
function applyCustomRange() {
  const fEl = document.getElementById('pp-from'), tEl = document.getElementById('pp-to');
  let from = fEl && fEl.value, to = tEl && tEl.value;
  if (!from || !to) { toast('Pick both a start and end date'); return; }
  if (from > to) { const tmp = from; from = to; to = tmp; } // tolerate reversed order
  setCustomPeriod({ type: 'range', from: from, to: to });
}
function _customPeriodLabel(p) {
  const now = new Date();
  if (p.type === 'lastweek') return 'Last week';
  if (p.type === 'lastmonth') return new Date(now.getFullYear(), now.getMonth() - 1, 1).toLocaleDateString('en-US', {month:'long', year:'numeric'});
  if (p.type === 'year') return String(p.y);
  if (p.type === 'range') return fmtDateShort(p.from) + ' – ' + fmtDateShort(p.to);
  return new Date(p.y, p.m, 1).toLocaleDateString('en-US', {month:'long', year:'numeric'});
}
function _customRange(p) {
  const now = new Date();
  if (p.type === 'range') {
    // Inclusive of the 'to' day: end boundary is the next midnight (getAnalyticsTxs
    // filters with ts < cutEnd).
    return [parseDate(p.from).getTime(), parseDate(p.to).getTime() + 86400000];
  }
  if (p.type === 'lastweek') {
    const day = (now.getDay() + 6) % 7;
    const thisMon = new Date(now.getFullYear(), now.getMonth(), now.getDate() - day).getTime();
    return [thisMon - 7 * 86400000, thisMon];
  }
  if (p.type === 'lastmonth') return [new Date(now.getFullYear(), now.getMonth() - 1, 1).getTime(), new Date(now.getFullYear(), now.getMonth(), 1).getTime()];
  if (p.type === 'year') return [new Date(p.y, 0, 1).getTime(), new Date(p.y + 1, 0, 1).getTime()];
  return [new Date(p.y, p.m, 1).getTime(), new Date(p.y, p.m + 1, 1).getTime()];
}
document.addEventListener('input', function(e){
  var t = e.target;
  if (t && t.classList && t.classList.contains('amt-input')) {
    t.style.width = Math.min(9, Math.max(4.2, (t.value || '').length + 1.2)) + 'ch';
  }
});
function setPeriod(p, el) {
  S.period = p;
  S.periodCustom = null;
  const _pill = document.getElementById('pb-custom');
  if (_pill) _pill.textContent = 'More…';
  document.querySelectorAll('#page-analytics .pb').forEach(c => c.classList.remove('active'));
  if (el) el.classList.add('active');
  renderAnalytics();
}

// ═══════════════════════════════════════════
//  TRANSACTION MODAL (Add / Edit)
// ═══════════════════════════════════════════
function setType(type) {
  S.type = type;
  document.getElementById('tb-exp').classList.toggle('active', type === 'expense');
  document.getElementById('tb-inc').classList.toggle('active', type === 'income');
  document.getElementById('ttp').classList.toggle('income', type === 'income');

  buildCatGrid();
}

function openAddModal() {
  S.editingTxId = null;
  const _amt = document.getElementById('add-amt'); if (_amt) _amt.style.width = '';
  document.getElementById('add-modal-title').textContent = 'New transaction';
  document.getElementById('del-tx-btn').style.display = 'none';
  updateAcctSelect();
  try { const bs = S.cur.sym.replace(/^[A-Za-z]+/, ''); document.getElementById('add-symbol').textContent = bs || S.cur.sym; } catch(e) {}
  const d = S.addDraft;
  if (d) {
    // Unsaved draft kept from last time · restore until saved or cancelled
    document.getElementById('add-amt').value = d.amt || '';
    document.getElementById('add-desc').value = d.desc || '';
    document.getElementById('add-notes').value = d.notes || '';
    document.getElementById('add-date').value = d.date || todayISO();
    setType(d.type || 'expense');
    if (d.cat) { S.cat = d.cat; buildCatGrid(); }
    if (d.acct) document.getElementById('add-acct').value = d.acct;
    else if (S.accts.length) document.getElementById('add-acct').value = S.accts[0].name;
  } else {
    document.getElementById('add-amt').value = '';
    document.getElementById('add-desc').value = '';
    document.getElementById('add-notes').value = '';
    document.getElementById('add-date').value = todayISO();
    if (S.accts.length) document.getElementById('add-acct').value = S.accts[0].name;
    setType('expense');
  }
  applyTrackIncome();
  updateMemberPicker('add-member-row','add-member-select');
  try { renderQuickAdd(); } catch(e) {}
  openOverlay('add-overlay');
  // Focus amount only on desktop · on mobile the keyboard opening immediately hides Save button
  setTimeout(() => {
    const amtEl = document.getElementById('add-amt');
    if (amtEl && window.innerWidth >= 640) amtEl.focus();
  }, 360);
}

function openEditTx(id) {
  const t = S.txs.find(x => x.id === id);
  if (!t) return;
  S.editingTxId = id;
  document.getElementById('add-modal-title').textContent = 'Edit transaction';
  document.getElementById('del-tx-btn').style.display = 'block';
  S.type = t.type;
  document.getElementById('tb-exp').classList.toggle('active', t.type === 'expense');
  document.getElementById('tb-inc').classList.toggle('active', t.type === 'income');
  document.getElementById('ttp').classList.toggle('income', t.type === 'income');

  S.cat = t.category;
  buildCatGrid();
  document.getElementById('add-amt').value = t.amount;
  document.getElementById('add-desc').value = t.description || '';
  document.getElementById('add-notes').value = t.notes || '';
  document.getElementById('add-date').value = t.date ? t.date.slice(0,10) : todayISO();
  updateAcctSelect();
  document.getElementById('add-acct').value = t.account || '';
  updateMemberPicker('add-member-row','add-member-select', t.memberId || '');
  openOverlay('add-overlay');
}

function _captureAddDraft() {
  if (S.editingTxId) return; // drafts only for new transactions
  const amt = document.getElementById('add-amt').value;
  const desc = document.getElementById('add-desc').value;
  const notes = document.getElementById('add-notes').value;
  if (!amt && !desc && !notes) { S.addDraft = null; return; }
  S.addDraft = {
    amt, desc, notes,
    date: document.getElementById('add-date').value,
    acct: document.getElementById('add-acct').value,
    cat: S.cat, type: S.type
  };
}
function cancelAddTx() {
  S.addDraft = null;
  S._skipDraft = true;
  closeOverlay('add-overlay');
}

async function submitTx() {
  // Mandatory account: must have at least one account before recording anything.
  if (!S.accts || S.accts.length === 0) {
    toast('Add an account first');
    closeOverlay('add-overlay');
    openAcctModal();
    return;
  }
  const amt = parseFloat((document.getElementById('add-amt').value || '').replace(/[^0-9.]/g, ''));
  if (!amt || amt <= 0) { toast('Enter a valid amount'); return; }
  const desc = document.getElementById('add-desc').value.trim();
  const acct = document.getElementById('add-acct').value;
  if (acct === '') { toast('Select an account'); return; }
  const date = document.getElementById('add-date').value || todayISO();
  const notes = document.getElementById('add-notes').value.trim();

  const editingId = S.editingTxId;
  const existing = editingId ? S.txs.find(x => x.id === editingId) : null;
  // Read the member picker BEFORE closing the overlay (hidden elements read wrong)
  const _memRow = document.getElementById('add-member-row');
  const _memSel = document.getElementById('add-member-select');
  const memVal = (_memRow && _memRow.style.display !== 'none' && _memSel) ? (_memSel.value || null) : null;
  // Saved · discard any kept draft, then optimistic close BEFORE the write
  S.addDraft = null; S._skipDraft = true;
  closeOverlay('add-overlay');

  if (editingId) {
    await apiPost({ action: 'updateTransaction', id: editingId, ts: Date.now(), type: S.type, amount: amt, category: S.cat, description: desc, account: acct, date, notes, memberId: memVal !== null ? memVal : (existing ? existing.memberId : null) });
    toast('✓ Transaction updated');
  } else {
    await apiPost({ action: 'addTransaction', id: 'tx_' + Date.now(), ts: Date.now(), type: S.type, amount: amt, category: S.cat, description: desc, account: acct, date, notes, memberId: memVal });
    toast(S.type === 'income' ? '✓ Income added' : '✓ Expense saved');
  }
}

function deleteTxFromModal() {
  if (!S.editingTxId) return;
  const id = S.editingTxId;
  showCustomConfirm('Delete Transaction', 'Remove this entry and adjust the account balance?', 'Delete', async () => {
    closeOverlay('add-overlay');
    toast('🗑 Transaction deleted');
    await apiPost({ action: 'deleteTx', id });
  });
}

// ═══════════════════════════════════════════
//  ACCOUNT MODAL (Add / Edit)
// ═══════════════════════════════════════════
function openAcctModal(id) {
  if (id) {
    const a = S.accts.find(x => x.id === id);
    if (!a) return;
    S.editingAcctId = id;
    document.getElementById('acct-modal-title').textContent = 'Edit account';
    document.getElementById('del-acct-btn').style.display = 'block';
    document.getElementById('acct-name').value = a.name;
    document.getElementById('acct-type').value = a.type;
    document.getElementById('acct-bal').value = a.balance;
  } else {
    S.editingAcctId = null;
    document.getElementById('acct-modal-title').textContent = 'Add account';
    document.getElementById('del-acct-btn').style.display = 'none';
    document.getElementById('acct-name').value = '';
    document.getElementById('acct-type').value = 'Checking';
    document.getElementById('acct-bal').value = '';
  }
  openOverlay('acct-overlay');
}

async function submitAcct() {
  const name = document.getElementById('acct-name').value.trim();
  if (!name) { toast('Enter an account name'); return; }
  const dup = (S.accts||[]).find(a => a.name.toLowerCase() === name.toLowerCase() && a.id !== S.editingAcctId);
  if (dup) { toast('An account with that name already exists'); return; }
  const type = document.getElementById('acct-type').value;
  const bal = parseFloat(document.getElementById('acct-bal').value) || 0;
  const editingId = S.editingAcctId;
  closeOverlay('acct-overlay');
  if (editingId) {
    await apiPost({ action: 'updateAccount', id: editingId, name, type, balance: bal });
    toast('✓ Account updated');
  } else {
    await apiPost({ action: 'addAccount', id: 'ac_' + Date.now(), name, type, balance: bal });
    toast('✓ Account added');
  }
}

function deleteAcctFromModal() {
  if (!S.editingAcctId) return;
  const id = S.editingAcctId;
  showCustomConfirm('Delete Account', 'Past transactions referencing this account will show "—". Continue?', 'Delete', async () => {
    closeOverlay('acct-overlay');
    toast('🗑 Account deleted');
    await apiPost({ action: 'deleteAccount', id });
  });
}

// ═══════════════════════════════════════════
//  RECURRING MODAL (Add / Edit)
// ═══════════════════════════════════════════
function setRecType(type) {
  S.recType = type;
  document.getElementById('rtb-exp').classList.toggle('active', type === 'expense');
  document.getElementById('rtb-inc').classList.toggle('active', type === 'income');
  document.getElementById('rttp').classList.toggle('income', type === 'income');
  buildRecCatGrid();
}

function openRecurringModal(id) {
  updateRecAcctSelect();
  if (id) {
    const r = S.recurring.find(x => x.id === id);
    if (!r) return;
    S.editingRecId = id;
    document.getElementById('recurring-modal-title').textContent = 'Edit recurring';
    document.getElementById('del-rec-btn').style.display = 'block';
    document.getElementById('rec-name').value = r.name;
    document.getElementById('rec-amt').value = r.amount;
    S.recType = r.type;
    S.recCat = r.category;
    document.getElementById('rec-acct').value = r.account || '';
    document.getElementById('rec-freq').value = r.frequency || 'monthly';
    document.getElementById('rec-next').value = r.nextDate ? r.nextDate.slice(0,10) : '';
  } else {
    S.editingRecId = null;
    document.getElementById('recurring-modal-title').textContent = 'Add recurring';
    document.getElementById('del-rec-btn').style.display = 'none';
    document.getElementById('rec-name').value = '';
    document.getElementById('rec-amt').value = '';
    S.recType = 'expense';
    S.recCat = CATS.find(c => c.n === 'Bills') ? 'Bills' : CATS[0].n; // recurring = usually a bill, not Food
    if (S.accts.length) document.getElementById('rec-acct').value = S.accts[0].name;
    document.getElementById('rec-freq').value = 'monthly';
    document.getElementById('rec-next').value = todayISO();
  }
  document.getElementById('rtb-exp').classList.toggle('active', S.recType === 'expense');
  document.getElementById('rtb-inc').classList.toggle('active', S.recType === 'income');
  document.getElementById('rttp').classList.toggle('income', S.recType === 'income');
  buildRecCatGrid();
  applyTrackIncome();
  openOverlay('recurring-overlay');
}

async function submitRecurring() {
  const name = document.getElementById('rec-name').value.trim();
  if (!name) { toast('Enter a name'); return; }
  const amt = parseFloat(document.getElementById('rec-amt').value);
  if (!amt || amt <= 0) { toast('Enter a valid amount'); return; }
  const acct = document.getElementById('rec-acct').value;
  const freq = document.getElementById('rec-freq').value;
  const next = document.getElementById('rec-next').value || todayISO();
  const editingId = S.editingRecId;
  const existing = editingId ? S.recurring.find(x => x.id === editingId) : null;
  closeOverlay('recurring-overlay');
  if (editingId) {
    await apiPost({ action: 'updateRecurring', id: editingId, name, type: S.recType, amount: amt, category: S.recCat, account: acct, frequency: freq, nextDate: next, memberId: existing ? existing.memberId : null, active: true });
    toast('✓ Updated');
  } else {
    await apiPost({ action: 'addRecurring', id: 'rec_' + Date.now(), name, type: S.recType, amount: amt, category: S.recCat, account: acct, frequency: freq, nextDate: next, memberId: S.member ? S.member.id : null, active: true });
    toast('✓ Added');
  }
  checkRecurringDue();
}

function deleteRecFromModal() {
  if (!S.editingRecId) return;
  const id = S.editingRecId;
  showCustomConfirm('Delete Recurring', 'Remove this recurring item?', 'Delete', async () => {
    closeOverlay('recurring-overlay');
    toast('🗑 Deleted');
    await apiPost({ action: 'deleteRecurring', id });
  });
}

// ═══════════════════════════════════════════
//  BUDGET MODAL (Add / Edit)
// ═══════════════════════════════════════════
function openBudgetModal(id) {
  buildBudgetCatOptions();
  if (id) {
    const b = S.budgets.find(x => x.id === id);
    if (!b) return;
    S.editingBudgetId = id;
    document.getElementById('budget-modal-title').textContent = 'Edit budget';
    document.getElementById('del-budget-btn').style.display = 'block';
    document.getElementById('budget-cat').value = b.category;
    document.getElementById('budget-limit').value = b.limit;
  } else {
    S.editingBudgetId = null;
    document.getElementById('budget-modal-title').textContent = 'Add budget';
    document.getElementById('del-budget-btn').style.display = 'none';
    document.getElementById('budget-limit').value = '';
  }
  openOverlay('budget-overlay');
}

async function submitBudget() {
  const cat = document.getElementById('budget-cat').value;
  const limit = parseFloat(document.getElementById('budget-limit').value);
  if (!limit || limit <= 0) { toast('Enter a valid limit'); return; }
  const wasEditing = !!S.editingBudgetId;
  // #7 · one budget per category: if a budget already exists for this category,
  // update it instead of minting a new id (which produced duplicate rows).
  let budgetId = S.editingBudgetId;
  if (!budgetId) { const existing = (S.budgets || []).find(b => b.category === cat); budgetId = existing ? existing.id : 'bg_' + Date.now(); }
  closeOverlay('budget-overlay');
  await apiPost({ action: 'setBudget', id: budgetId, category: cat, limit });
  toast(wasEditing ? '✓ Budget updated' : '✓ Budget set');
}

function deleteBudgetFromModal() {
  if (!S.editingBudgetId) return;
  const id = S.editingBudgetId;
  showCustomConfirm('Delete Budget', 'Remove this budget limit?', 'Delete', async () => {
    closeOverlay('budget-overlay');
    toast('🗑 Budget removed');
    await apiPost({ action: 'deleteBudget', id });
  });
}

// ═══════════════════════════════════════════
//  MEMBER MODAL (from Settings)
// ═══════════════════════════════════════════
function openMemberModal() {
  S.newMemberPin = '';
  document.getElementById('member-name').value = '';
  updatePinDots('member-pin-dots', '');
  const kb = document.getElementById('member-pin-kb'); if (kb) kb.value = '';
  document.getElementById('member-save-btn').disabled = true;
  openOverlay('member-overlay');
}

function memberPinKey(k) {
  if (k === 'check') { /* re-evaluate */ }
  else if (k === 'clear') S.newMemberPin = '';
  else if (k === 'back') S.newMemberPin = S.newMemberPin.slice(0, -1);
  else if (S.newMemberPin.length < 4) S.newMemberPin += k;
  updatePinDots('member-pin-dots', S.newMemberPin);
  const name = document.getElementById('member-name').value.trim();
  document.getElementById('member-save-btn').disabled = !(name && S.newMemberPin.length === 4);
}

async function submitMember() {
  const name = document.getElementById('member-name').value.trim();
  if (!name || S.newMemberPin.length !== 4) return;
  const usedColors = S.members.map(m => m.color);
  const color = MEMBER_COLORS.find(c => !usedColors.includes(c)) || MEMBER_COLORS[S.members.length % MEMBER_COLORS.length];
  const id = 'mem_' + Date.now();
  const pinHash = await hashPin(S.newMemberPin); // #10
  closeOverlay('member-overlay');
  await apiPost({ action: 'addMember', id, name, pin: pinHash, color });
  toast(`✓ ${name} added`);
}

// ═══════════════════════════════════════════
//  CATEGORY EDITOR · with emoji picker
// ═══════════════════════════════════════════
function populateCategoryEditor(side) {
  const customNames = safeParse(localStorage.getItem('prizm_cat_names'), {});
  const customEmojis = safeParse(localStorage.getItem('prizm_cat_emojis'), {});
  const incNames = safeParse(localStorage.getItem('prizm_inc_names'), {});
  const incEmojis = safeParse(localStorage.getItem('prizm_inc_emojis'), {});

  if (!side || side === 'exp') {
    const el = document.getElementById('onboard-exp-cats');
    if (el) el.innerHTML = CATS.slice(0,5).map(c => `
      <div style="display:flex;gap:8px;align-items:center">
        <input class="finput" type="text" maxlength="2" value="${esc(customEmojis[c.n] || c.e)}" oninput="saveCustomEmoji('exp','${esc(c.n)}',this.value)" style="width:52px;text-align:center;font-size:20px;padding:8px 4px;flex-shrink:0">
        <input class="finput" type="text" placeholder="${esc(c.n)}" value="${esc(customNames[c.n] || '')}" oninput="saveCustomName('exp','${esc(c.n)}',this.value)" style="flex:1">
      </div>`).join('');
  }

  if (!side || side === 'inc') {
    const el = document.getElementById('onboard-inc-cats');
    if (el) el.innerHTML = INC_CATS.slice(0,5).map(c => `
      <div style="display:flex;gap:8px;align-items:center">
        <input class="finput" type="text" maxlength="2" value="${esc(incEmojis[c.n] || c.e)}" oninput="saveCustomEmoji('inc','${esc(c.n)}',this.value)" style="width:52px;text-align:center;font-size:20px;padding:8px 4px;flex-shrink:0">
        <input class="finput" type="text" placeholder="${esc(c.n)}" value="${esc(incNames[c.n] || '')}" oninput="saveCustomName('inc','${esc(c.n)}',this.value)" style="flex:1">
      </div>`).join('');
  }
}

function saveCustomEmoji(type, key, val) {
  const storageKey = type === 'inc' ? 'prizm_inc_emojis' : 'prizm_cat_emojis';
  const obj = safeParse(localStorage.getItem(storageKey), {});
  if (val.trim()) obj[key] = val.trim();
  else delete obj[key];
  localStorage.setItem(storageKey, JSON.stringify(obj));
  schedulePrefSync(); // #2
}

function saveCustomName(type, key, val) {
  const storageKey = type === 'inc' ? 'prizm_inc_names' : 'prizm_cat_names';
  const obj = safeParse(localStorage.getItem(storageKey), {});
  if (val.trim()) obj[key] = val.trim();
  else delete obj[key];
  localStorage.setItem(storageKey, JSON.stringify(obj));
  schedulePrefSync(); // #2
}

function saveCategoryNames() {
  // Changes are persisted live via oninput handlers (saveCustomEmoji / saveCustomName)
  // Just advance to the accounts setup step
  showSetup('sp-accounts');
}
