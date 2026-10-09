const SPREADSHEET_ID = '1IYyx3Peb6Jkiq8miyZaTOpfhoHKR5rGQ1qhwhLrdDXw';
const FLASH_SHEET_NAME = 'Flash';
const FLASH_HISTORY_SHEET_NAME = 'FlashHistory';
const FLASH_Q_STATE = 'FlashQuestionState';
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
      case 'setFlashReviewPriority': response = setFlashReviewPriority(payload.cardNo,payload.enabled); break;
      case 'saveFlashResult': response = saveFlashResult(payload.cardNo,payload.subject,payload.userAnswer,payload.correctAnswer,payload.isCorrect,payload.eventId,payload.timestamp); break;
      case 'saveFlashResultsBatch': response = saveFlashResultsBatch(payload.events); break;
      case 'getFlashStats': response = {success:true,stats:getFlashStats()}; break;
      case 'getFlashDailyStatus': response = {success:true,status:getFlashDailyStatus()}; break;
      case 'getFlashQuestionStats': response = {success:true,stats:getFlashQuestionStats(payload.cardNo)}; break;
      case 'getFlashStatsBundle': response = {success:true,bundle:getFlashStatsBundle()}; break;
      case 'rebuildFlashStats': response = rebuildFlashStats(); break;
      case 'repairFlashStats': response = repairFlashStats(); break;
      default: throw new Error('不明なAPIアクションです。');
    }
    return createApiOutput_(callback,response);
  } catch (err) {
    return createApiOutput_(callback,{success:false,error:err && err.message ? err.message : String(err)});
  }
}

function normalizeCallback_(v){const s=normalizeText(v);if(!/^[A-Za-z_$][0-9A-Za-z_$\\.]*$/.test(s))throw new Error('callbackが不正です。');return s;}
function verifyApiToken_(v){const expected=PropertiesService.getScriptProperties().getProperty('API_TOKEN');if(expected&&normalizeText(v)!==expected)throw new Error('APIトークンが不正です。');}
function createApiOutput_(cb,data){return ContentService.createTextOutput(cb+'('+JSON.stringify(data)+');').setMimeType(ContentService.MimeType.JAVASCRIPT);}
function ss_(){return SpreadsheetApp.openById(SPREADSHEET_ID);}
function normalizeText(v){return v==null?'':String(v).trim();}
function formatJstDate_(d){return Utilities.formatDate(d,JST_TIMEZONE,'yyyy-MM-dd');}
function normalizeFlashAnswer_(v){const s=normalizeText(v);if(s==='〇'||s==='○')return '○';if(s==='✕'||s==='×'||s.toLowerCase()==='x')return '×';return s;}
function normalizeStudyDate_(v,fallback){if(v instanceof Date&&!isNaN(v))return formatJstDate_(v);const s=normalizeText(v);if(/^\\d{4}-\\d{2}-\\d{2}$/.test(s))return s;const d=s?new Date(s):new Date(fallback);return isNaN(d)?'':formatJstDate_(d);}

function getFlashSheet_(){const sh=ss_().getSheetByName(FLASH_SHEET_NAME);if(!sh)throw new Error('Flash シートが見つかりません。');return sh;}
function getFlashHistorySheet_(){const ss=ss_();let sh=ss.getSheetByName(FLASH_HISTORY_SHEET_NAME);if(!sh)sh=ss.insertSheet(FLASH_HISTORY_SHEET_NAME);if(sh.getLastRow()===0)sh.getRange(1,1,1,8).setValues([['日時','問','科目','回答','正答','正誤','EventID','学習日(JST)']]);return sh;}
function getStatsSheet_(name,headers){const ss=ss_();let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);if(sh.getLastRow()===0)sh.getRange(1,1,1,headers.length).setValues([headers]);return sh;}
function qState_(){return getStatsSheet_(FLASH_Q_STATE,['問','出題数','正答数','最終回答日時','重点復習']);}
function sStats_(){return getStatsSheet_(FLASH_S_STATS,['科目','回答数','正解数']);}
function dStats_(){return getStatsSheet_(FLASH_D_STATS,['日付','回答数','正解数']);}

function isReviewPriority_(value){const s=normalizeText(value).toLowerCase();return value===true||value===1||['1','true','yes','y','○','〇','重点','high'].indexOf(s)>=0;}

function getFlashCards(){
  const sh=getFlashSheet_(),lr=sh.getLastRow(),lc=sh.getLastColumn();if(lr<2)return[];
  const v=sh.getRange(1,1,lr,lc).getDisplayValues(),h=v[0].map(normalizeText);
  const i={no:h.indexOf('問'),subject:h.indexOf('科目'),topic:h.indexOf('論点'),question:h.indexOf('質問'),answer:h.indexOf('正答'),source:h.indexOf('出典'),explanation:h.indexOf('解説'),detailUrl:h.indexOf('詳細リンク')};
  ['no','subject','topic','question','answer','source'].forEach(k=>{if(i[k]<0)throw new Error('Flash シートの列が不足しています: '+k);});
  return v.slice(1).filter(r=>normalizeText(r[i.question])).map(r=>({no:normalizeText(r[i.no]),subject:normalizeText(r[i.subject]),topic:normalizeText(r[i.topic]),question:normalizeText(r[i.question]),answer:normalizeFlashAnswer_(r[i.answer]),source:normalizeText(r[i.source]),explanation:i.explanation>=0?normalizeText(r[i.explanation]):'',detailUrl:i.detailUrl>=0?normalizeText(r[i.detailUrl]):''}));
}

function findQuestionStateRow_(sh,no){
  const lr=sh.getLastRow();if(lr<2)return 0;
  const vals=sh.getRange(2,1,lr-1,1).getDisplayValues().flat();
  const idx=vals.findIndex(v=>normalizeText(v)===no);
  return idx<0?0:idx+2;
}
function ensureQuestionStateRow_(sh,no){
  let row=findQuestionStateRow_(sh,no);
  if(row)return row;
  sh.appendRow([no,0,0,'','']);
  return sh.getLastRow();
}
function setFlashReviewPriority(cardNo,enabled){
  const no=normalizeText(cardNo);if(!no)throw new Error('問番号が不正です。');
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const sh=qState_(),row=ensureQuestionStateRow_(sh,no),flag=enabled===true||String(enabled).toLowerCase()==='true';
    sh.getRange(row,5).setValue(flag?'○':'');
    return{success:true,cardNo:no,reviewPriority:flag};
  }finally{lock.releaseLock();}
}
function updateQuestionState_(no,ok,dt){
  const sh=qState_(),row=ensureQuestionStateRow_(sh,no);
  sh.getRange(row,2).setValue(Number(sh.getRange(row,2).getValue()||0)+1);
  if(ok)sh.getRange(row,3).setValue(Number(sh.getRange(row,3).getValue()||0)+1);
  sh.getRange(row,4).setValue(dt);
}

function saveFlashResult(cardNo,subject,userAnswer,correctAnswer,isCorrect,eventId,eventTimestamp){
  const r=saveFlashResultsBatch([{cardNo,subject,userAnswer,correctAnswer,isCorrect,eventId,timestamp:eventTimestamp}]);
  return{success:true,duplicate:r.duplicates>0};
}
function saveFlashResultsBatch(events){
  if(!Array.isArray(events))throw new Error('回答データが不正です。');
  if(events.length>100)throw new Error('一度に同期できる回答は100件までです。');
  if(!events.length)return{success:true,saved:0,duplicates:0,eventIds:[]};
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const hist=getFlashHistorySheet_(),q=qState_(),ss=sStats_(),ds=dStats_();

    const existingIds=new Set();
    const histLast=hist.getLastRow();
    if(histLast>=2)hist.getRange(2,7,histLast-1,1).getDisplayValues().forEach(r=>{const id=normalizeText(r[0]);if(id)existingIds.add(id);});

    const qMap={};const qLast=q.getLastRow();
    if(qLast>=2)q.getRange(2,1,qLast-1,5).getValues().forEach((r,i)=>{
      const no=normalizeText(r[0]);if(no)qMap[no]={row:i+2,total:Number(r[1]||0),correct:Number(r[2]||0),last:r[3],priority:r[4]};
    });
    function readRowMap_(sh){
      const out={},lr=sh.getLastRow();
      if(lr>=2)sh.getRange(2,1,lr-1,3).getValues().forEach((r,i)=>{const k=normalizeText(r[0]);if(k)out[k]={row:i+2,total:Number(r[1]||0),correct:Number(r[2]||0)};});
      return out;
    }
    const sMap=readRowMap_(ss),dMap=readRowMap_(ds);
    const histRows=[],newQ=[],newS=[],newD=[],changedQ=new Set(),changedS=new Set(),changedD=new Set(),acceptedIds=[];
    let duplicates=0;

    events.forEach(ev=>{
      const no=normalizeText(ev&&ev.cardNo),sub=normalizeText(ev&&ev.subject)||'未分類',ans=normalizeFlashAnswer_(ev&&ev.userAnswer),correct=normalizeFlashAnswer_(ev&&ev.correctAnswer),eid=normalizeText(ev&&ev.eventId);
      if(!no||!ans||!correct)throw new Error('短答回答データが不正です。');
      if(eid&&existingIds.has(eid)){duplicates++;acceptedIds.push(eid);return;}
      let dt=new Date();if(ev&&ev.timestamp){const p=new Date(ev.timestamp);if(!isNaN(p))dt=p;}
      const ok=ev&&ev.isCorrect===true||String(ev&&ev.isCorrect).toLowerCase()==='true',day=formatJstDate_(dt);
      histRows.push([dt,no,sub,ans,correct,ok,eid,day]);
      if(eid){existingIds.add(eid);acceptedIds.push(eid);}

      let qr=qMap[no];
      if(!qr){qr=qMap[no]={row:0,total:0,correct:0,last:'',priority:''};newQ.push(no);}
      qr.total++;if(ok)qr.correct++;qr.last=dt;if(qr.row)changedQ.add(no);

      let sr=sMap[sub];
      if(!sr){sr=sMap[sub]={row:0,total:0,correct:0};newS.push(sub);}
      sr.total++;if(ok)sr.correct++;if(sr.row)changedS.add(sub);

      let dr=dMap[day];
      if(!dr){dr=dMap[day]={row:0,total:0,correct:0};newD.push(day);}
      dr.total++;if(ok)dr.correct++;if(dr.row)changedD.add(day);
    });

    if(histRows.length)hist.getRange(histLast+1,1,histRows.length,8).setValues(histRows);

    changedQ.forEach(no=>{const r=qMap[no];q.getRange(r.row,2,1,3).setValues([[r.total,r.correct,r.last]]);});
    if(newQ.length){const rows=newQ.map(no=>{const r=qMap[no];return[no,r.total,r.correct,r.last,r.priority||''];});q.getRange(q.getLastRow()+1,1,rows.length,5).setValues(rows);}

    changedS.forEach(k=>{const r=sMap[k];ss.getRange(r.row,2,1,2).setValues([[r.total,r.correct]]);});
    if(newS.length){const rows=newS.map(k=>[k,sMap[k].total,sMap[k].correct]);ss.getRange(ss.getLastRow()+1,1,rows.length,3).setValues(rows);}

    changedD.forEach(k=>{const r=dMap[k];ds.getRange(r.row,2,1,2).setValues([[r.total,r.correct]]);});
    if(newD.length){const rows=newD.map(k=>[k,dMap[k].total,dMap[k].correct]);ds.getRange(ds.getLastRow()+1,1,rows.length,3).setValues(rows);}

    return{success:true,saved:histRows.length,duplicates,eventIds:acceptedIds};
  }finally{lock.releaseLock();}
}
function flashEventIdExists_(sh,eid){if(!eid||sh.getLastRow()<2)return false;return Boolean(sh.getRange(2,7,sh.getLastRow()-1,1).createTextFinder(eid).matchEntireCell(true).findNext());}
function incrementStat_(sh,key,ok){const lr=sh.getLastRow();if(lr>=2){const vals=sh.getRange(2,1,lr-1,1).getDisplayValues();for(let i=0;i<vals.length;i++){if(normalizeText(vals[i][0])===key){const row=i+2;sh.getRange(row,2).setValue(Number(sh.getRange(row,2).getValue()||0)+1);if(ok)sh.getRange(row,3).setValue(Number(sh.getRange(row,3).getValue()||0)+1);return;}}}sh.appendRow([key,1,ok?1:0]);}

function statSheetTotal_(sh){if(sh.getLastRow()<2)return 0;return sh.getRange(2,2,sh.getLastRow()-1,1).getValues().reduce((n,r)=>n+Number(r[0]||0),0);}
function ensureStatsBuilt_(){
  const q=qState_(),hist=getFlashHistorySheet_();
  const historyCount=Math.max(hist.getLastRow()-1,0),questionCount=statSheetTotal_(q);
  if(historyCount!==questionCount)rebuildFlashStats();
}
function readStatMap_(sh){const out={};if(sh.getLastRow()<2)return out;sh.getRange(2,1,sh.getLastRow()-1,3).getValues().forEach(r=>{const k=normalizeText(r[0]);if(k)out[k]={total:Number(r[1]||0),correct:Number(r[2]||0)};});return out;}
function readQuestionStateMap_(){
  const sh=qState_(),out={};if(sh.getLastRow()<2)return out;
  sh.getRange(2,1,sh.getLastRow()-1,5).getValues().forEach(r=>{
    const k=normalizeText(r[0]);if(!k)return;
    const last=r[3] instanceof Date&&!isNaN(r[3])?r[3].toISOString():normalizeText(r[3]);
    out[k]={total:Number(r[1]||0),correct:Number(r[2]||0),lastAnsweredAt:last,reviewPriority:isReviewPriority_(r[4])};
  });
  return out;
}
function readDailyStatMap_(){const sh=dStats_(),out={};if(sh.getLastRow()<2)return out;sh.getRange(2,1,sh.getLastRow()-1,3).getValues().forEach(r=>{const k=normalizeStudyDate_(r[0],r[0]);if(k)out[k]={total:Number(r[1]||0),correct:Number(r[2]||0)};});return out;}

function rebuildFlashStats(){
  const q=qState_(),s=sStats_(),d=dStats_(),priority={};
  if(q.getLastRow()>=2)q.getRange(2,1,q.getLastRow()-1,5).getValues().forEach(r=>{const no=normalizeText(r[0]);if(no&&isReviewPriority_(r[4]))priority[no]='○';});
  [q,s,d].forEach(sh=>{if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).clearContent();});
  const qm={},sm={},dm={},hist=getFlashHistorySheet_(),lr=hist.getLastRow();
  if(lr>=2)hist.getRange(2,1,lr-1,8).getValues().forEach(r=>{
    const no=normalizeText(r[1]);if(!no)return;
    const sub=normalizeText(r[2])||'未分類',ok=r[5]===true||String(r[5]).toLowerCase()==='true',day=normalizeStudyDate_(r[7],r[0]);
    const dt=r[0] instanceof Date?r[0]:new Date(r[0]),valid=!isNaN(dt);
    if(!qm[no])qm[no]={total:0,correct:0,lastAnsweredAt:null};
    qm[no].total++;if(ok)qm[no].correct++;
    if(valid&&(!qm[no].lastAnsweredAt||dt.getTime()>qm[no].lastAnsweredAt.getTime()))qm[no].lastAnsweredAt=dt;
    incMap_(sm,sub,ok);if(day)incMap_(dm,day,ok);
  });
  const qRows=Object.keys(qm).sort((a,b)=>Number(a)-Number(b)).map(no=>[no,qm[no].total,qm[no].correct,qm[no].lastAnsweredAt||'',priority[no]||'']);
  Object.keys(priority).filter(no=>!qm[no]).sort((a,b)=>Number(a)-Number(b)).forEach(no=>qRows.push([no,0,0,'',priority[no]]));
  if(qRows.length)q.getRange(2,1,qRows.length,5).setValues(qRows);
  writeMap_(s,sm);writeMap_(d,dm);
  return{success:true,questions:Object.keys(qm).length,subjects:Object.keys(sm).length,days:Object.keys(dm).length,answers:Math.max(lr-1,0)};
}
function repairFlashStats(){ensureStatsBuilt_();return{success:true,bundle:getFlashStatsBundle()};}
function incMap_(m,k,ok){if(!m[k])m[k]={total:0,correct:0};m[k].total++;if(ok)m[k].correct++;}
function writeMap_(sh,m){const rows=Object.keys(m).sort().map(k=>[k,m[k].total,m[k].correct]);if(rows.length)sh.getRange(2,1,rows.length,3).setValues(rows);}

function getFlashStatsBundle(){
  ensureStatsBuilt_();
  const questions=readQuestionStateMap_(),subjects=readStatMap_(sStats_()),dayMap=readDailyStatMap_();
  let total=0,correct=0;Object.keys(questions).forEach(k=>{total+=questions[k].total;correct+=questions[k].correct;});
  const today=formatJstDate_(new Date()),td=dayMap[today]||{total:0,correct:0};
  const days=Object.keys(dayMap).sort().map(date=>({date,total:dayMap[date].total,correct:dayMap[date].correct,accuracy:dayMap[date].total?dayMap[date].correct/dayMap[date].total:0}));
  return{stats:{total,correct,accuracy:total?correct/total:0},daily:{date:today,goal:FLASH_DAILY_GOAL,count:td.total,correct:td.correct,accuracy:td.total?td.correct/td.total:0,remaining:Math.max(FLASH_DAILY_GOAL-td.total,0),achieved:td.total>=FLASH_DAILY_GOAL},questions,subjects,days};
}
function getFlashStats(){return getFlashStatsBundle().stats;}
function getFlashQuestionStats(cardNo){const no=normalizeText(cardNo);if(!no)throw new Error('問番号が不正です。');ensureStatsBuilt_();const s=readQuestionStateMap_()[no]||{total:0,correct:0};return{cardNo:no,total:s.total,correct:s.correct,accuracy:s.total?s.correct/s.total:0};}
function getFlashDailyStatus(){return getFlashStatsBundle().daily;}
