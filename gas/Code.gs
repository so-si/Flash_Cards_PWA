const SPREADSHEET_ID = '1IYyx3Peb6Jkiq8miyZaTOpfhoHKR5rGQ1qhwhLrdDXw';
const FLASH_SHEET_NAME = 'Flash';
const FLASH_HISTORY_SHEET_NAME = 'FlashHistory';
const FLASH_Q_STATS = 'FlashStatsQuestion';
const FLASH_S_STATS = 'FlashStatsSubject';
const FLASH_D_STATS = 'FlashStatsDaily';
const JST_TIMEZONE = 'Asia/Tokyo';
const FLASH_DAILY_GOAL = 50;

function doGet(e) {
  const p = e && e.parameter ? e.parameter : {};
  if (!p.action) return HtmlService.createHtmlOutput('Flash Cards API');
  const callback = normalizeCallback_(p.callback);
  try {
    verifyApiToken_(p.token);
    const payload = p.payload ? JSON.parse(p.payload) : {};
    let response;
    switch (p.action) {
      case 'getFlashCards': response = {success:true,cards:getFlashCards()}; break;
      case 'saveFlashResult': response = saveFlashResult(payload.cardNo,payload.subject,payload.userAnswer,payload.correctAnswer,payload.isCorrect,payload.eventId,payload.timestamp); break;
      case 'getFlashStats': response = {success:true,stats:getFlashStats()}; break;
      case 'getFlashDailyStatus': response = {success:true,status:getFlashDailyStatus()}; break;
      case 'getFlashQuestionStats': response = {success:true,stats:getFlashQuestionStats(payload.cardNo)}; break;
      case 'getFlashStatsBundle': response = {success:true,bundle:getFlashStatsBundle()}; break;
      case 'rebuildFlashStats': response = rebuildFlashStats(); break;
      default: throw new Error('不明なAPIアクションです。');
    }
    return createApiOutput_(callback,response);
  } catch (err) {
    return createApiOutput_(callback,{success:false,error:err && err.message ? err.message : String(err)});
  }
}

function normalizeCallback_(v){const s=normalizeText(v);if(!/^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(s))throw new Error('callbackが不正です。');return s;}
function verifyApiToken_(v){const expected=PropertiesService.getScriptProperties().getProperty('API_TOKEN');if(expected&&normalizeText(v)!==expected)throw new Error('APIトークンが不正です。');}
function createApiOutput_(cb,data){return ContentService.createTextOutput(cb+'('+JSON.stringify(data)+');').setMimeType(ContentService.MimeType.JAVASCRIPT);}
function ss_(){return SpreadsheetApp.openById(SPREADSHEET_ID);}
function normalizeText(v){return v==null?'':String(v).trim();}
function formatJstDate_(d){return Utilities.formatDate(d,JST_TIMEZONE,'yyyy-MM-dd');}
function normalizeFlashAnswer_(v){const s=normalizeText(v);if(s==='〇'||s==='○')return '○';if(s==='✕'||s==='×'||s.toLowerCase()==='x')return '×';return s;}
function normalizeStudyDate_(v,fallback){if(v instanceof Date&&!isNaN(v))return formatJstDate_(v);const s=normalizeText(v);if(/^\d{4}-\d{2}-\d{2}$/.test(s))return s;const d=s?new Date(s):new Date(fallback);return isNaN(d)?'':formatJstDate_(d);}

function getFlashSheet_(){const sh=ss_().getSheetByName(FLASH_SHEET_NAME);if(!sh)throw new Error('Flash シートが見つかりません。');return sh;}
function getFlashHistorySheet_(){const ss=ss_();let sh=ss.getSheetByName(FLASH_HISTORY_SHEET_NAME);if(!sh)sh=ss.insertSheet(FLASH_HISTORY_SHEET_NAME);if(sh.getLastRow()===0)sh.getRange(1,1,1,8).setValues([['日時','問','科目','回答','正答','正誤','EventID','学習日(JST)']]);return sh;}
function getStatsSheet_(name,headers){const ss=ss_();let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);if(sh.getLastRow()===0)sh.getRange(1,1,1,headers.length).setValues([headers]);return sh;}
function qStats_(){return getStatsSheet_(FLASH_Q_STATS,['問','回答数','正解数']);}
function sStats_(){return getStatsSheet_(FLASH_S_STATS,['科目','回答数','正解数']);}
function dStats_(){return getStatsSheet_(FLASH_D_STATS,['日付','回答数','正解数']);}

function getFlashCards(){
  const sh=getFlashSheet_(),lr=sh.getLastRow(),lc=sh.getLastColumn();if(lr<2)return[];
  const v=sh.getRange(1,1,lr,lc).getDisplayValues(),h=v[0].map(normalizeText);
  const i={no:h.indexOf('問'),subject:h.indexOf('科目'),precedent:h.indexOf('判例名'),question:h.indexOf('質問'),answer:h.indexOf('正答'),source:h.indexOf('出典'),explanation:h.indexOf('解説'),detailUrl:h.indexOf('詳細リンク')};
  ['no','subject','precedent','question','answer','source'].forEach(k=>{if(i[k]<0)throw new Error('Flash シートの列が不足しています: '+k);});
  return v.slice(1).filter(r=>normalizeText(r[i.question])).map(r=>({no:normalizeText(r[i.no]),subject:normalizeText(r[i.subject]),precedent:normalizeText(r[i.precedent]),question:normalizeText(r[i.question]),answer:normalizeFlashAnswer_(r[i.answer]),source:normalizeText(r[i.source]),explanation:i.explanation>=0?normalizeText(r[i.explanation]):'',detailUrl:i.detailUrl>=0?normalizeText(r[i.detailUrl]):''}));
}

function saveFlashResult(cardNo,subject,userAnswer,correctAnswer,isCorrect,eventId,eventTimestamp){
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const no=normalizeText(cardNo),sub=normalizeText(subject),ans=normalizeFlashAnswer_(userAnswer),correct=normalizeFlashAnswer_(correctAnswer),eid=normalizeText(eventId);
    if(!no||!ans||!correct)throw new Error('短答回答データが不正です。');
    const hist=getFlashHistorySheet_();
    if(eid&&flashEventIdExists_(hist,eid))return{success:true,duplicate:true};
    let dt=new Date();if(eventTimestamp){const p=new Date(eventTimestamp);if(!isNaN(p))dt=p;}
    const ok=isCorrect===true||String(isCorrect).toLowerCase()==='true',day=formatJstDate_(dt);
    hist.appendRow([dt,no,sub,ans,correct,ok,eid,day]);
    incrementStat_(qStats_(),no,ok);
    incrementStat_(sStats_(),sub||'未分類',ok);
    incrementStat_(dStats_(),day,ok);
    return{success:true,duplicate:false};
  }finally{lock.releaseLock();}
}
function flashEventIdExists_(sh,eid){if(!eid||sh.getLastRow()<2)return false;return Boolean(sh.getRange(2,7,sh.getLastRow()-1,1).createTextFinder(eid).matchEntireCell(true).findNext());}
function incrementStat_(sh,key,ok){const lr=sh.getLastRow();if(lr>=2){const vals=sh.getRange(2,1,lr-1,1).getDisplayValues();for(let i=0;i<vals.length;i++){if(normalizeText(vals[i][0])===key){const row=i+2;sh.getRange(row,2).setValue(Number(sh.getRange(row,2).getValue()||0)+1);if(ok)sh.getRange(row,3).setValue(Number(sh.getRange(row,3).getValue()||0)+1);return;}}}sh.appendRow([key,1,ok?1:0]);}

function ensureStatsBuilt_(){const q=qStats_(),s=sStats_(),d=dStats_();if(q.getLastRow()>1||s.getLastRow()>1||d.getLastRow()>1)return;const hist=getFlashHistorySheet_();if(hist.getLastRow()>1)rebuildFlashStats();}
function readStatMap_(sh){const out={};if(sh.getLastRow()<2)return out;sh.getRange(2,1,sh.getLastRow()-1,3).getValues().forEach(r=>{const k=normalizeText(r[0]);if(k)out[k]={total:Number(r[1]||0),correct:Number(r[2]||0)};});return out;}

function rebuildFlashStats(){
  const q=qStats_(),s=sStats_(),d=dStats_();[q,s,d].forEach(sh=>{if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).clearContent();});
  const qm={},sm={},dm={},hist=getFlashHistorySheet_(),lr=hist.getLastRow();
  if(lr>=2)hist.getRange(2,1,lr-1,8).getValues().forEach(r=>{const no=normalizeText(r[1]);if(!no)return;const sub=normalizeText(r[2])||'未分類',ok=r[5]===true||String(r[5]).toLowerCase()==='true',day=normalizeStudyDate_(r[7],r[0]);incMap_(qm,no,ok);incMap_(sm,sub,ok);if(day)incMap_(dm,day,ok);});
  writeMap_(q,qm);writeMap_(s,sm);writeMap_(d,dm);
  return{success:true,questions:Object.keys(qm).length,subjects:Object.keys(sm).length,days:Object.keys(dm).length};
}
function incMap_(m,k,ok){if(!m[k])m[k]={total:0,correct:0};m[k].total++;if(ok)m[k].correct++;}
function writeMap_(sh,m){const rows=Object.keys(m).sort().map(k=>[k,m[k].total,m[k].correct]);if(rows.length)sh.getRange(2,1,rows.length,3).setValues(rows);}

function getFlashStatsBundle(){
  ensureStatsBuilt_();
  const questions=readStatMap_(qStats_()),subjects=readStatMap_(sStats_()),dayMap=readStatMap_(dStats_());
  let total=0,correct=0;Object.keys(questions).forEach(k=>{total+=questions[k].total;correct+=questions[k].correct;});
  const today=formatJstDate_(new Date()),td=dayMap[today]||{total:0,correct:0};
  const days=Object.keys(dayMap).sort().slice(-30).map(date=>({date,total:dayMap[date].total,correct:dayMap[date].correct,accuracy:dayMap[date].total?dayMap[date].correct/dayMap[date].total:0}));
  return{stats:{total,correct,accuracy:total?correct/total:0},daily:{date:today,goal:FLASH_DAILY_GOAL,count:td.total,correct:td.correct,accuracy:td.total?td.correct/td.total:0,remaining:Math.max(FLASH_DAILY_GOAL-td.total,0),achieved:td.total>=FLASH_DAILY_GOAL},questions,subjects,days};
}
function getFlashStats(){return getFlashStatsBundle().stats;}
function getFlashQuestionStats(cardNo){const no=normalizeText(cardNo);if(!no)throw new Error('問番号が不正です。');ensureStatsBuilt_();const s=readStatMap_(qStats_())[no]||{total:0,correct:0};return{cardNo:no,total:s.total,correct:s.correct,accuracy:s.total?s.correct/s.total:0};}
function getFlashDailyStatus(){return getFlashStatsBundle().daily;}
