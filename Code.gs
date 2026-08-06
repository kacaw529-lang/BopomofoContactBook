/************************************************************
 *  班級數位聯絡簿 — Google Apps Script 後端
 *  綁定於 Google 試算表（Extensions ▸ Apps Script）
 *  部署為「網頁應用程式」：執行身分＝我、存取對象＝任何人
 *  詳細步驟見 README.md
 ************************************************************/

var PHOTO_FOLDER = '聯絡簿照片';   // 照片存放的雲端硬碟資料夾名稱

/************************************************************
 *  ★ 教師寫入密碼（務必改成你自己的密碼）★
 *  作用：只有輸入正確密碼才能新增／修改／刪除功課、成語、
 *        連結、名單、範本與稽核紀錄。家長與學生的「數位作業
 *        繳交」不需密碼。防護在伺服器端執行，改前端也無法繞過。
 *  注意：請勿使用與 Google 帳號相同的密碼。
 ************************************************************/
var TEACHER_PASS = '改成你的密碼';

function checkPass(p){
  return String(p || '') === TEACHER_PASS && TEACHER_PASS !== '';
}

/* ---------- 共用 ---------- */
function ss(){ return SpreadsheetApp.getActiveSpreadsheet(); }
function ymd(v){
  if(v instanceof Date) return Utilities.formatDate(v, 'Asia/Taipei', 'yyyy-MM-dd');
  return String(v).trim();
}
// 寬容比對：支援乾淨字串(2026-07-11)、Date 物件、以及 Date 長字串(Sat Jul 11 2026 ...)
function dateMatch(cell, target){
  var d = null;
  if(cell instanceof Date){ d = cell; }
  else {
    var s = String(cell).trim();
    if(s === target) return true;                 // 乾淨字串直接相等
    if(/^\d{4}-\d{2}-\d{2}$/.test(s)) return false; // 是乾淨字串但不同日
    var p = new Date(s);                          // 嘗試解析 Date 長字串
    if(!isNaN(p.getTime())) d = p; else return false;
  }
  var zones = ['Asia/Taipei', ss().getSpreadsheetTimeZone(), 'GMT'];
  for(var i = 0; i < zones.length; i++){
    if(Utilities.formatDate(d, zones[i], 'yyyy-MM-dd') === target) return true;
  }
  return false;
}
// 正規化任何日期輸入為 yyyy-MM-dd 文字
function normDate(v){
  var s = String(v).trim();
  if(/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var p = (v instanceof Date) ? v : new Date(s);
  if(!isNaN(p.getTime())) return Utilities.formatDate(p, 'Asia/Taipei', 'yyyy-MM-dd');
  return s;
}
function sheet(name, header){
  var s = ss().getSheetByName(name);
  if(!s){ s = ss().insertSheet(name); s.appendRow(header); return s; }
  var width = Math.max(header.length, s.getLastColumn() || header.length);
  var firstRow = s.getRange(1, 1, 1, width).getValues()[0];
  var need = false;
  for(var i = 0; i < header.length; i++){ if(firstRow[i] !== header[i]){ need = true; break; } }
  if(need) s.getRange(1, 1, 1, header.length).setValues([header]);
  return s;
}
function jsonOut(obj){
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
function jsonp(cb, obj){
  return ContentService.createTextOutput(cb + '(' + JSON.stringify(obj) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/* ---------- 照片：data URL -> 雲端硬碟，回傳可顯示網址 ---------- */
function savePhoto(dataUrl, name){
  if(!dataUrl || dataUrl.indexOf('data:') !== 0) return dataUrl || '';
  try{
    var m = dataUrl.match(/^data:([^;]+);base64,(.+)$/);
    if(!m) return '';
    var blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], name + '.jpg');
    var folders = DriveApp.getFoldersByName(PHOTO_FOLDER);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(PHOTO_FOLDER);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000';
  }catch(err){ return ''; }
}

/* ---------- 讀取（前端用 JSONP 呼叫）---------- */
function doGet(e){
  var p = e.parameter || {};
  var cb = p.callback || 'callback';
  if(p.action === 'load'){
    return jsonp(cb, loadDate(p.date));
  }
  if(p.action === 'range'){
    return jsonp(cb, loadRange(p.from, p.to));
  }
  return jsonp(cb, {ok:true, msg:'聯絡簿後端運作中'});
}

function loadDate(date){
  var hw = readRows('功課', ['日期','id','內容','注音','數位'])
    .filter(function(r){ return dateMatch(r['日期'], date); })
    .map(function(r){
      var zh = {};
      try{ zh = r['注音'] ? JSON.parse(r['注音']) : {}; }catch(x){}
      return { id:r['id'], text:r['內容'], zhuyin:zh, digital: !!r['數位'] };
    });
  var idioms = readRows('成語', ['日期','id','成語','解釋','照片'])
    .filter(function(r){ return dateMatch(r['日期'], date); })
    .map(function(r){ return { id:r['id'], text:r['成語'], mean:r['解釋'], photo:r['照片'] }; });
  var links = readRows('連結', ['id','名稱','網址','圖示'])
    .map(function(r){ return { id:r['id'], title:r['名稱'], url:r['網址'], icon:r['圖示'] }; });
  var submissions = readRows('繳交', ['時間','日期','學生','完成項目','備註','照片'])
    .filter(function(r){ return dateMatch(r['日期'], date); })
    .map(function(r){
      var done = {};
      try{ done = r['完成項目'] ? JSON.parse(r['完成項目']) : {}; }catch(x){}
      return { student:r['學生'], done:done, note:r['備註'], photo:r['照片'],
               ts: r['時間'] ? new Date(r['時間']).getTime() : Date.now() };
    });
  var reminder = '';
  var rem = readRows('叮嚀', ['日期','內容']).filter(function(r){ return dateMatch(r['日期'], date); });
  if(rem.length) reminder = rem[rem.length-1]['內容'];
  var roster = readRows('名單', ['姓名']).map(function(r){ return r['姓名']; }).filter(String);
  var templates = readRows('範本', ['功課項目']).map(function(r){ return r['功課項目']; }).filter(String);
  var audit = {};
  readRows('稽核', ['日期','學生','完成項目']).filter(function(r){ return dateMatch(r['日期'], date); }).forEach(function(r){
    var done = {};
    try{ done = r['完成項目'] ? JSON.parse(r['完成項目']) : {}; }catch(x){}
    audit[r['學生']] = done;
  });
  return { homework:hw, idioms:idioms, reminder:reminder, links:links, submissions:submissions, roster:roster, audit:audit, templates:templates };
}

/* 日期區間：給「未補繳」儀表板用 */
function loadRange(from, to){
  var roster = readRows('名單', ['姓名']).map(function(r){ return r['姓名']; }).filter(String);
  var templates = readRows('範本', ['功課項目']).map(function(r){ return r['功課項目']; }).filter(String);
  var inRange = function(d){ return (!from || d >= from) && (!to || d <= to); };
  var byDate = {};
  readRows('功課', ['日期','id','內容','注音','數位']).forEach(function(r){
    var d = ymd(r['日期']); if(!inRange(d)) return;
    if(!byDate[d]) byDate[d] = { date:d, homework:[], audit:{} };
    byDate[d].homework.push({ id:r['id'], text:r['內容'], digital: !!r['數位'] });
  });
  readRows('稽核', ['日期','學生','完成項目']).forEach(function(r){
    var d = ymd(r['日期']); if(!inRange(d) || !byDate[d]) return;
    var done = {};
    try{ done = r['完成項目'] ? JSON.parse(r['完成項目']) : {}; }catch(x){}
    byDate[d].audit[r['學生']] = done;
  });
  var days = Object.keys(byDate).sort().map(function(d){ return byDate[d]; });
  return { roster:roster, days:days };
}

function readRows(name, header){
  var s = ss().getSheetByName(name);
  if(!s) { sheet(name, header); return []; }
  var values = s.getDataRange().getValues();
  if(values.length < 2) return [];
  var head = values[0];
  return values.slice(1).map(function(row){
    var o = {}; head.forEach(function(h,i){ o[h] = row[i]; }); return o;
  });
}

/* ---------- 寫入（前端用 POST 呼叫）---------- */
function doPost(e){
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try{
    var d = JSON.parse(e.postData.contents);
    if(d.date) d.date = normDate(d.date);

    // 學生／家長繳交：不需密碼（但只能新增自己的繳交紀錄）
    if(d.action === 'submit') return jsonOut(submit(d));

    // 以下皆為教師動作，必須通過密碼驗證
    if(!checkPass(d.pass)) return jsonOut({ok:false, code:'AUTH', msg:'密碼錯誤，無權限修改'});

    if(d.action === 'saveDay')       return jsonOut(saveDay(d));
    if(d.action === 'saveLinks')     return jsonOut(saveLinks(d));
    if(d.action === 'saveRoster')    return jsonOut(saveRoster(d));
    if(d.action === 'saveTemplates') return jsonOut(saveTemplates(d));
    if(d.action === 'saveAudit')     return jsonOut(saveAudit(d));
    return jsonOut({ok:false, msg:'未知動作'});
  }catch(err){
    return jsonOut({ok:false, msg:String(err)});
  }finally{
    lock.releaseLock();
  }
}

function deleteRowsByDate(name, header, date, ids){
  var s = sheet(name, header);
  SpreadsheetApp.flush(); // 先讓先前的寫入落定，避免讀到快取舊值
  var values = s.getDataRange().getValues();
  var idSet = {};
  (ids || []).forEach(function(x){ if(x) idSet[String(x)] = true; });
  for(var i = values.length - 1; i >= 1; i--){
    var hitDate = dateMatch(values[i][0], date);
    var hitId = values[i][1] != null && idSet[String(values[i][1])];
    if(hitDate || hitId) s.deleteRow(i + 1);
  }
  SpreadsheetApp.flush();
}

function saveDay(d){
  // 功課
  deleteRowsByDate('功課', ['日期','id','內容','注音','數位'], d.date, (d.homework||[]).map(function(h){return h.id;}));
  var hwS = sheet('功課', ['日期','id','內容','注音','數位']);
  hwS.getRange('A:A').setNumberFormat('@');
  (d.homework || []).forEach(function(h){
    hwS.appendRow([d.date, h.id, h.text, JSON.stringify(h.zhuyin || {}), h.digital ? '1' : '']);
  });
  // 成語（照片若為 data URL 先上傳硬碟）
  deleteRowsByDate('成語', ['日期','id','成語','解釋','照片'], d.date, (d.idioms||[]).map(function(i){return i.id;}));
  var idS = sheet('成語', ['日期','id','成語','解釋','照片']);
  idS.getRange('A:A').setNumberFormat('@');
  (d.idioms || []).forEach(function(i){
    var photo = savePhoto(i.photo, '成語_' + d.date + '_' + i.id);
    idS.appendRow([d.date, i.id, i.text, i.mean || '', photo]);
  });
  // 叮嚀
  deleteRowsByDate('叮嚀', ['日期','內容'], d.date);
  if(d.reminder){
    var rmS = sheet('叮嚀', ['日期','內容']);
    rmS.getRange('A:A').setNumberFormat('@');
    rmS.appendRow([d.date, d.reminder]);
  }
  return {ok:true};
}

function saveLinks(d){
  var s = sheet('連結', ['id','名稱','網址','圖示']);
  s.clear();
  s.appendRow(['id','名稱','網址','圖示']);
  (d.links || []).forEach(function(l){
    var icon = savePhoto(l.icon, '連結圖示_' + l.id);
    s.appendRow([l.id, l.title, l.url, icon]);
  });
  return {ok:true};
}

function submit(d){
  var s = sheet('繳交', ['時間','日期','學生','完成項目','備註','照片']);
  s.getRange('B:B').setNumberFormat('@');
  var photo = savePhoto(d.photo, '繳交_' + d.date + '_' + Date.now());
  s.appendRow([new Date(), d.date, d.student, JSON.stringify(d.done || {}), d.note || '', photo]);
  return {ok:true};
}

function saveTemplates(d){
  var s = sheet('範本', ['功課項目']);
  s.clear();
  s.appendRow(['功課項目']);
  (d.templates || []).forEach(function(t){ if(t) s.appendRow([t]); });
  return {ok:true};
}

function saveRoster(d){
  var s = sheet('名單', ['姓名']);
  s.clear();
  s.appendRow(['姓名']);
  (d.roster || []).forEach(function(n){ s.appendRow([n]); });
  return {ok:true};
}

function saveAudit(d){
  deleteRowsByDate('稽核', ['日期','學生','完成項目'], d.date);
  var s = sheet('稽核', ['日期','學生','完成項目']);
  s.getRange('A:A').setNumberFormat('@');
  var a = d.audit || {};
  Object.keys(a).forEach(function(stu){
    s.appendRow([d.date, stu, JSON.stringify(a[stu] || {})]);
  });
  return {ok:true};
}

/************************************************************
 *  selfTest — 一鍵自我診斷（平時不影響運作，出問題時執行）
 *  用法：Apps Script 編輯器選取 selfTest ▸ 執行 ▸ 看執行紀錄
 *  會檢查：試算表綁定、工作表結構、日期格式三形態比對、
 *          doGet 模擬回應。任一項 ✗ 即為病灶所在。
 ************************************************************/
function selfTest(){
  var today = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd');
  Logger.log('=== 聯絡簿自我診斷（' + today + '）===');

  // 1. 試算表綁定
  try{ Logger.log('1. 綁定試算表：✓ ' + ss().getName()); }
  catch(e){ Logger.log('1. 綁定試算表：✗ ' + e + '（專案沒綁到試算表，請從試算表▸擴充功能▸Apps Script 進入）'); return; }

  // 2. 工作表結構
  var expect = { '功課':['日期','id','內容','注音','數位'], '成語':['日期','id','成語','解釋','照片'],
                 '連結':['id','名稱','網址','圖示'], '繳交':['時間','日期','學生','完成項目','備註','照片'],
                 '叮嚀':['日期','內容'], '名單':['姓名'], '稽核':['日期','學生','完成項目'], '範本':['功課項目'] };
  for(var name in expect){
    var s = ss().getSheetByName(name);
    if(!s){ Logger.log('2. 工作表「' + name + '」：（不存在，首次寫入會自動建立）'); continue; }
    var head = s.getRange(1,1,1,expect[name].length).getValues()[0].join(',');
    Logger.log('2. 工作表「' + name + '」表頭：' + (head === expect[name].join(',') ? '✓' : '✗ 實際=[' + head + '] 應為=[' + expect[name].join(',') + ']'));
  }

  // 3. 日期三形態比對（歷史教訓：偽裝成日期的長字串）
  var cases = [
    ['乾淨字串', today],
    ['Date物件', new Date()],
    ['Date長字串', new Date().toString()]
  ];
  cases.forEach(function(c){
    Logger.log('3. dateMatch(' + c[0] + ')：' + (dateMatch(c[1], today) ? '✓' : '✗ ← 日期比對失敗，檢查 dateMatch'));
  });

  // 4. 功課表逐列健檢（型別與比對）
  var hw = ss().getSheetByName('功課');
  if(hw){
    var v = hw.getDataRange().getValues();
    for(var i = 1; i < Math.min(v.length, 6); i++){
      var cell = v[i][0];
      Logger.log('4. 功課列' + (i+1) + '：值=[' + cell + '] 型別=' + (cell instanceof Date ? 'Date' : 'text'));
    }
    if(v.length > 6) Logger.log('4. …（僅列出前5列）');
  }

  // 5. 模擬前端請求
  var res = doGet({parameter:{action:'load', date:today, callback:'t'}});
  var body = ''; try{ body = res.getContent(); }catch(e){ body = String(res); }
  Logger.log('5. doGet(今日) 回應前200字：' + body.slice(0, 200));
  Logger.log('=== 診斷結束：任何 ✗ 即為問題所在 ===');
}
