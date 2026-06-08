/*************************************************************
 * 姊妹成長銀行 2.0 Pro — Google Apps Script 後台 API
 * --------------------------------------------------------
 * 用法：
 * 1) 建立一個 Google 試算表
 * 2) 擴充功能 → Apps Script，把這份程式碼整個貼上
 * 3) 部署 → 新增部署作業 → 類型「網頁應用程式」
 *    - 執行身分：我（你的帳號）
 *    - 誰可以存取：「所有人」(Anyone)
 * 4) 複製「網頁應用程式網址」(.../exec) 貼到前端設定頁
 *
 * 資料模型：整份 state 以 JSON 存在 DATA!A1（單一真實來源）
 *           另外自動產生「孩子總覽 / 紀錄」分頁給人類閱讀（唯讀快照）
 *************************************************************/

const DATA_SHEET   = 'DATA';
const SHEET_KIDS   = '孩子總覽';
const SHEET_RECORD = '紀錄';

// 與前端一致的孩子名稱對照
const NAME_MAP = {
  weining:'惟甯', xinying:'昕穎', yaoting:'耀霆', yuxuan:'予璇'
};
// 與前端一致的等級表
const LEVELS = [
  {lv:1,name:'小小新手',xp:0},   {lv:2,name:'努力寶寶',xp:100},
  {lv:3,name:'任務小達人',xp:250},{lv:4,name:'自律小勇士',xp:500},
  {lv:5,name:'成長騎士',xp:800},  {lv:6,name:'姊妹隊長',xp:1200},
  {lv:7,name:'超級自律王',xp:1700},{lv:8,name:'家庭 MVP',xp:2300},
  {lv:9,name:'成長冠軍',xp:3000}, {lv:10,name:'傳說姊妹',xp:4000}
];

/* ===== 入口 ===== */
function doGet(e){
  const action = (e && e.parameter && e.parameter.action) || 'load';
  if (action === 'load') return json(loadState());
  if (action === 'ping') return json({ok:true, msg:'pong', time:Date.now()});
  return json({ok:false, error:'unknown action'});
}

function doPost(e){
  const lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch(err){ return json({ok:false, error:'busy'}); }
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'save') {
      const st = body.state || {};
      st.updatedAt = Date.now();
      saveState(st);
      try { writeMirror(st); } catch(mErr) { /* 鏡像失敗不影響主資料 */ }
      return json({ok:true, updatedAt: st.updatedAt});
    }
    return json({ok:false, error:'unknown action'});
  } catch(err) {
    return json({ok:false, error:String(err)});
  } finally {
    lock.releaseLock();
  }
}

/* ===== 工具 ===== */
function json(o){
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
function sheet(name){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

/* ===== 主資料讀寫（JSON blob） ===== */
function loadState(){
  const sh = sheet(DATA_SHEET);
  const v = sh.getRange('A1').getValue();
  if (!v) return {ok:true, state:null, updatedAt:0};
  try {
    const st = JSON.parse(v);
    return {ok:true, state:st, updatedAt: st.updatedAt || 0};
  } catch(e){
    return {ok:true, state:null, updatedAt:0};
  }
}
function saveState(st){
  sheet(DATA_SHEET).getRange('A1').setValue(JSON.stringify(st));
}

/* ===== 人類可讀鏡像（唯讀快照） ===== */
function levelOf(lt){
  let cur = LEVELS[0];
  for (const l of LEVELS) if (lt >= l.xp) cur = l;
  return cur;
}
function writeMirror(st){
  // 孩子總覽
  const k = sheet(SHEET_KIDS);
  k.clear();
  k.getRange(1,1,1,7).setValues([['孩子','等級','稱號','累積XP','可用XP','平板分鐘','徽章數']]);
  const kids = st.kids || {};
  const rows = [];
  Object.keys(kids).forEach(key=>{
    const c = kids[key];
    const L = levelOf(c.lifetime||0);
    rows.push([ NAME_MAP[key]||key, 'Lv.'+L.lv, L.name,
      c.lifetime||0, c.xp||0, c.tablet||0, (c.badges||[]).length ]);
  });
  if (rows.length) k.getRange(2,1,rows.length,7).setValues(rows);
  k.getRange(1,1,1,7).setFontWeight('bold');

  // 紀錄（最新 500 筆）
  const r = sheet(SHEET_RECORD);
  r.clear();
  r.getRange(1,1,1,8).setValues([['時間','孩子','加分者','類別','任務','XP','備註','狀態']]);
  const recs = (st.records || []).slice(0, 500);
  const statusMap = {pending:'待審核', approved:'已核准', rejected:'已拒絕', done:'完成'};
  const rr = recs.map(x=>[
    x.time ? new Date(x.time) : '',
    x.kid ? (NAME_MAP[x.kid]||x.kid) : '姊妹',
    x.parent||'', x.cat||'', x.task||'',
    x.xp||0, x.note||'', statusMap[x.status]||x.status||''
  ]);
  if (rr.length) {
    r.getRange(2,1,rr.length,8).setValues(rr);
    r.getRange(2,1,rr.length,1).setNumberFormat('yyyy/MM/dd HH:mm');
  }
  r.getRange(1,1,1,8).setFontWeight('bold');
}
