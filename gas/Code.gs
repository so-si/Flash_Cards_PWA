const SPREADSHEET_ID = '1IYyx3Peb6Jkiq8miyZaTOpfhoHKR5rGQ1qhwhLrdDXw';
const FLASH_SHEET_NAME = 'Flash';
const FLASH_HISTORY_SHEET_NAME = 'FlashHistory';
const FLASH_Q_STATE = 'FlashQuestionState';
const FLASH_S_STATS = 'FlashStatsSubject';
const FLASH_D_STATS = 'FlashStatsDaily';
const JST_TIMEZONE = 'Asia/Tokyo';
const FLASH_DAILY_GOAL = 100;

function doGet(e) {
  const p = e && e.parameter ? e.parameter : {};
  if (!p.action) return HtmlService.createHtmlOutput('Flash Cards API');
  const callback = normalizeCallback_(p.callback);
  try {
    verifyApiToken_(p.token);
    const payload = p.payload ? JSON.parse(p.payload) : {};
    let response;
    switch (p.action) {
      case 'getFlashCards': response = Object.assign({success:true},getFlashCardsPayload_()); break;
      case 'markFlashResultReview': response = markFlashResultReview(payload.eventId); break;
      case 'saveFlashResult': response = saveFlashResult(payload.cardNo,payload.subject,payload.userAnswer,payload.correctAnswer,payload.isCorrect,payload.eventId,payload.timestamp,payload.needsReview); break;
      case 'saveFlashResultsBatch': response = saveFlashResultsBatch(payload.events); break;
      case 'getFlashStats': response = {success:true,stats:getFlashStats()}; break;
      case 'getFlashDailyStatus': response = {success:true,status:getFlashDailyStatus()}; break;
      case 'getFlashQuestionStats': response = {success:true,stats:getFlashQuestionStats(payload.cardNo)}; break;
      case 'getFlashStatsBundle': response = {success:true,bundle:getFlashStatsBundle()}; break;
      case 'getFlashReviewQueue': response = {success:true,queue:getFlashReviewQueue()}; break;
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
function normalizeStudyDate_(v,fallback){
  if(v instanceof Date&&!isNaN(v.getTime()))return formatJstDate_(v);
  const raw=normalizeText(v),m=raw.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})$/);
  if(m)return m[1]+'-'+m[2].padStart(2,'0')+'-'+m[3].padStart(2,'0');
  const d=raw?new Date(raw):new Date(fallback);
  return isNaN(d.getTime())?'':formatJstDate_(d);
}

function getFlashSheet_(){const sh=ss_().getSheetByName(FLASH_SHEET_NAME);if(!sh)throw new Error('Flash シートが見つかりません。');return sh;}
function getFlashHistorySheet_(){const ss=ss_();let sh=ss.getSheetByName(FLASH_HISTORY_SHEET_NAME);if(!sh)sh=ss.insertSheet(FLASH_HISTORY_SHEET_NAME);if(sh.getMaxColumns()<9)sh.insertColumnsAfter(sh.getMaxColumns(),9-sh.getMaxColumns());if(sh.getLastRow()===0)sh.getRange(1,1,1,9).setValues([['日時','問','科目','回答','正答','正誤','EventID','学習日(JST)','要復習指定']]);else if(sh.getRange(1,9).getValue()!=='要復習指定')sh.getRange(1,9).setValue('要復習指定');return sh;}
function getStatsSheet_(name,headers){const ss=ss_();let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);if(sh.getLastRow()===0)sh.getRange(1,1,1,headers.length).setValues([headers]);return sh;}
function qState_(){const sh=getStatsSheet_(FLASH_Q_STATE,['問','出題数','正答数','最終回答日時','最終要復習']);if(sh.getRange(1,5).getValue()!=='最終要復習'){const n=sh.getLastRow()-1;if(n>0)sh.getRange(2,5,n,1).clearContent();sh.getRange(1,5).setValue('最終要復習');}return sh;}
function sStats_(){return getStatsSheet_(FLASH_S_STATS,['科目','回答数','正解数']);}
function dStats_(){return getStatsSheet_(FLASH_D_STATS,['日付','回答数','正解数']);}

function getFlashCardsPayload_(){
  const sh=getFlashSheet_(),lr=sh.getLastRow(),lc=sh.getLastColumn();
  if(lr<1)return {cards:[],version:''};
  const v=sh.getRange(1,1,lr,lc).getDisplayValues();
  const hasVersion=normalizeText(v[0][0]).replace(/^\uFEFF/,'')==='#VERSION';
  const version=hasVersion?normalizeText(v[0][1]):'';
  const headerIndex=hasVersion?1:0;
  if(lr<=headerIndex+1)return {cards:[],version:version};
  const h=v[headerIndex].map(normalizeText);
  const i={no:h.indexOf('問'),subject:h.indexOf('科目'),topic:h.indexOf('論点'),question:h.indexOf('質問'),answer:h.indexOf('正答'),source:h.indexOf('出典'),explanation:h.indexOf('解説'),detailUrl:h.indexOf('詳細リンク')};
  ['no','subject','topic','question','answer','source'].forEach(k=>{if(i[k]<0)throw new Error('Flash シートの列が不足しています: '+k);});
  const cards=v.slice(headerIndex+1).filter(r=>normalizeText(r[i.question])).map(r=>({no:normalizeText(r[i.no]),subject:normalizeText(r[i.subject]),topic:normalizeText(r[i.topic]),question:normalizeText(r[i.question]),answer:normalizeFlashAnswer_(r[i.answer]),source:normalizeText(r[i.source]),explanation:i.explanation>=0?normalizeText(r[i.explanation]):'',detailUrl:i.detailUrl>=0?normalizeText(r[i.detailUrl]):''}));
  return {cards:cards,version:version};
}
function getFlashCards(){return getFlashCardsPayload_().cards;}

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
function updateQuestionState_(no,ok,dt){
  const sh=qState_(),row=ensureQuestionStateRow_(sh,no);
  sh.getRange(row,2).setValue(Number(sh.getRange(row,2).getValue()||0)+1);
  if(ok)sh.getRange(row,3).setValue(Number(sh.getRange(row,3).getValue()||0)+1);
  sh.getRange(row,4).setValue(dt);
}

function saveFlashResult(cardNo,subject,userAnswer,correctAnswer,isCorrect,eventId,eventTimestamp,needsReview){
  const r=saveFlashResultsBatch([{cardNo,subject,userAnswer,correctAnswer,isCorrect,eventId,timestamp:eventTimestamp,needsReview}]);
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
      const no=normalizeText(r[0]);if(no)qMap[no]={row:i+2,total:Number(r[1]||0),correct:Number(r[2]||0),last:r[3],needsReview:r[4]===true};
    });
    function readRowMap_(sh){
      const out={},lr=sh.getLastRow();
      if(lr>=2)sh.getRange(2,1,lr-1,3).getValues().forEach((r,i)=>{const k=normalizeText(r[0]);if(k)out[k]={row:i+2,total:Number(r[1]||0),correct:Number(r[2]||0)};});
      return out;
    }
    const sMap=readRowMap_(ss),dMap={};
    const dLast=ds.getLastRow();
    if(dLast>=2)ds.getRange(2,1,dLast-1,3).getValues().forEach((r,i)=>{
      const day=normalizeStudyDate_(r[0],r[0]);
      if(day&&!dMap[day])dMap[day]={row:i+2,total:Number(r[1]||0),correct:Number(r[2]||0)};
    });
    const histRows=[],newQ=[],newS=[],newD=[],changedQ=new Set(),changedS=new Set(),changedD=new Set(),acceptedIds=[];
    let duplicates=0;

    events.forEach(ev=>{
      const no=normalizeText(ev&&ev.cardNo),sub=normalizeText(ev&&ev.subject)||'未分類',ans=normalizeFlashAnswer_(ev&&ev.userAnswer),correct=normalizeFlashAnswer_(ev&&ev.correctAnswer),eid=normalizeText(ev&&ev.eventId);
      if(!no||!ans||!correct)throw new Error('短答回答データが不正です。');
      if(eid&&existingIds.has(eid)){duplicates++;acceptedIds.push(eid);return;}
      let dt=new Date();if(ev&&ev.timestamp){const p=new Date(ev.timestamp);if(!isNaN(p))dt=p;}
      const ok=ev&&ev.isCorrect===true||String(ev&&ev.isCorrect).toLowerCase()==='true',day=formatJstDate_(dt);
      histRows.push([dt,no,sub,ans,correct,ok,eid,day,ev.needsReview===true]);
      if(eid){existingIds.add(eid);acceptedIds.push(eid);}

      let qr=qMap[no];
      if(!qr){qr=qMap[no]={row:0,total:0,correct:0,last:'',needsReview:false};newQ.push(no);}
      qr.total++;if(ok)qr.correct++;qr.last=dt;qr.needsReview=!ok||ev.needsReview===true;if(qr.row)changedQ.add(no);

      let sr=sMap[sub];
      if(!sr){sr=sMap[sub]={row:0,total:0,correct:0};newS.push(sub);}
      sr.total++;if(ok)sr.correct++;if(sr.row)changedS.add(sub);

      let dr=dMap[day];
      if(!dr){dr=dMap[day]={row:0,total:0,correct:0};newD.push(day);}
      dr.total++;if(ok)dr.correct++;if(dr.row)changedD.add(day);
    });

    if(histRows.length)hist.getRange(histLast+1,1,histRows.length,9).setValues(histRows);

    changedQ.forEach(no=>{const r=qMap[no];q.getRange(r.row,2,1,4).setValues([[r.total,r.correct,r.last,r.needsReview]]);});
    if(newQ.length){const rows=newQ.map(no=>{const r=qMap[no];return[no,r.total,r.correct,r.last,r.needsReview];});q.getRange(q.getLastRow()+1,1,rows.length,5).setValues(rows);}

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
  const q=qState_(),hist=getFlashHistorySheet_(),daily=dStats_();
  const historyCount=Math.max(hist.getLastRow()-1,0),questionCount=statSheetTotal_(q);
  const dailyCount=statSheetTotal_(daily);
  let badDaily=historyCount!==dailyCount;
  if(!badDaily&&daily.getLastRow()>2){
    const seen=new Set();
    daily.getRange(2,1,daily.getLastRow()-1,1).getValues().forEach(r=>{
      const key=normalizeStudyDate_(r[0],r[0]);
      if(!key||seen.has(key))badDaily=true;
      if(key)seen.add(key);
    });
  }
  if(historyCount!==questionCount||badDaily){
    const lock=LockService.getScriptLock();lock.waitLock(20000);
    try{
      const actualHistory=Math.max(getFlashHistorySheet_().getLastRow()-1,0);
      const ds=dStats_(),found=new Set();
      let duplicate=false;
      if(ds.getLastRow()>1)ds.getRange(2,1,ds.getLastRow()-1,1).getValues().forEach(r=>{
        const k=normalizeStudyDate_(r[0],r[0]);
        if(!k||found.has(k))duplicate=true;
        if(k)found.add(k);
      });
      if(actualHistory!==statSheetTotal_(qState_())||actualHistory!==statSheetTotal_(ds)||duplicate)rebuildFlashStats();
    }finally{lock.releaseLock();}
  }
}
function readStatMap_(sh){const out={};if(sh.getLastRow()<2)return out;sh.getRange(2,1,sh.getLastRow()-1,3).getValues().forEach(r=>{const k=normalizeText(r[0]);if(k)out[k]={total:Number(r[1]||0),correct:Number(r[2]||0)};});return out;}
function readQuestionStateMap_(){
  const sh=qState_(),out={};if(sh.getLastRow()<2)return out;
  sh.getRange(2,1,sh.getLastRow()-1,5).getValues().forEach(r=>{
    const k=normalizeText(r[0]);if(!k)return;
    const last=r[3] instanceof Date&&!isNaN(r[3])?r[3].toISOString():normalizeText(r[3]);
    out[k]={total:Number(r[1]||0),correct:Number(r[2]||0),lastAnsweredAt:last,needsReview:r[4]===true};
  });
  return out;
}
function readDailyStatMap_(){
  const sh=dStats_(),out={};
  if(sh.getLastRow()<2)return out;
  sh.getRange(2,1,sh.getLastRow()-1,3).getValues().forEach(r=>{
    const k=normalizeStudyDate_(r[0],r[0]);
    if(k){if(!out[k])out[k]={total:0,correct:0};out[k].total+=Number(r[1]||0);out[k].correct+=Number(r[2]||0);}
  });
  return out;
}

function rebuildFlashStats(){
  const q=qState_(),s=sStats_(),d=dStats_();
  [q,s,d].forEach(sh=>{if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).clearContent();});
  const qm={},sm={},dm={},hist=getFlashHistorySheet_(),lr=hist.getLastRow();
  if(lr>=2)hist.getRange(2,1,lr-1,9).getValues().forEach(r=>{
    const no=normalizeText(r[1]);if(!no)return;
    const sub=normalizeText(r[2])||'未分類',ok=r[5]===true||String(r[5]).toLowerCase()==='true',day=normalizeStudyDate_(r[7],r[0]);
    const dt=r[0] instanceof Date?r[0]:new Date(r[0]),valid=!isNaN(dt);
    if(!qm[no])qm[no]={total:0,correct:0,lastAnsweredAt:null,needsReview:false};
    qm[no].total++;if(ok)qm[no].correct++;
    if(valid&&(!qm[no].lastAnsweredAt||dt.getTime()>qm[no].lastAnsweredAt.getTime())){qm[no].lastAnsweredAt=dt;qm[no].needsReview=!ok||r[8]===true;}
    incMap_(sm,sub,ok);if(day)incMap_(dm,day,ok);
  });
  const qRows=Object.keys(qm).sort((a,b)=>Number(a)-Number(b)).map(no=>[no,qm[no].total,qm[no].correct,qm[no].lastAnsweredAt||'',qm[no].needsReview]);
  if(qRows.length)q.getRange(2,1,qRows.length,5).setValues(qRows);
  writeMap_(s,sm);writeMap_(d,dm);
  return{success:true,questions:Object.keys(qm).length,subjects:Object.keys(sm).length,days:Object.keys(dm).length,answers:Math.max(lr-1,0)};
}
function repairFlashStats(){
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{return{success:true,repair:rebuildFlashStats()};}
  finally{lock.releaseLock();}
}
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

function getFlashReviewQueue(){
 const today=formatJstDate_(new Date()),cutoff=new Date();cutoff.setDate(cutoff.getDate()-8);
 const cutoffDay=formatJstDate_(cutoff),wrongByDay={},h=getFlashHistorySheet_(),lr=h.getLastRow();
 if(lr>=2)h.getRange(2,1,lr-1,9).getValues().forEach(r=>{
 const day=normalizeStudyDate_(r[7],r[0]),no=normalizeText(r[1]),correct=r[5]===true||String(r[5]).toLowerCase()==='true';
 if(no&&(!correct||r[8]===true)&&day>=cutoffDay&&day<today){if(!wrongByDay[day])wrongByDay[day]=[];wrongByDay[day].push(no);}
 });
 Object.keys(wrongByDay).forEach(day=>wrongByDay[day]=[...new Set(wrongByDay[day])]);
 return {wrongByDay,asOf:new Date().toISOString()};
}

function markFlashResultReview(eventId){
 const id=normalizeText(eventId);if(!id)throw new Error('EventIDがありません');
 const lock=LockService.getScriptLock();lock.waitLock(20000);
 try{
 const h=getFlashHistorySheet_(),lr=h.getLastRow();
 if(lr<2)throw new Error('回答の同期待ちです');
 const ids=h.getRange(2,7,lr-1,1).getDisplayValues().flat(),at=ids.lastIndexOf(id);
 if(at<0)throw new Error('回答の同期待ちです');
 const row=at+2,no=normalizeText(h.getRange(row,2).getValue());
 h.getRange(row,9).setValue(true);
 const state=qState_(),sr=ensureQuestionStateRow_(state,no);
 const latest=h.getRange(row,1).getValue(),stored=state.getRange(sr,4).getValue();
 if(!stored||new Date(latest).getTime()>=new Date(stored).getTime())state.getRange(sr,5).setValue(true);
 return {success:true};
 }finally{lock.releaseLock()}
}
