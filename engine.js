/* MBBS Recall Pro v5 core engine. Dependency-free; works offline. */
(function(root){
'use strict';
const SUBJECTS=['Medicine','Surgery','Obstetrics','Gynecology','Pediatrics','Ophthalmology','ENT','Orthopedics','Dermatology','Anesthesia','Radiology','Psychiatry'];
const DEFAULT_INTERVALS={Regular:{fail:1,partial:2,pass:[3,7,14,30,45]},'Must-Do':{fail:1,partial:2,pass:[2,5,12,25,40]},Late:{fail:1,partial:2,pass:[2,5,10,20,35]}};
const today=()=>{const d=new Date();return iso(d)};
function iso(d){return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`}
function date(s){if(typeof s!=='string'||!/^\d{4}-\d{2}-\d{2}$/.test(s))return null;const [y,m,d]=s.split('-').map(Number);const x=new Date(y,m-1,d,12);return x.getFullYear()===y&&x.getMonth()===m-1&&x.getDate()===d?x:null}
function validDate(s){return !!date(s)}
function addDays(s,n){const d=date(s);if(!d)return '';d.setDate(d.getDate()+Number(n));return iso(d)}
function diffDays(a,b){const da=date(a),db=date(b);return da&&db?Math.round((da-db)/86400000):NaN}
function normResult(s){const x=String(s??'').trim().toUpperCase();if(['PASS','P','PASSED'].includes(x))return 'PASS';if(['FAIL','F','FAILED'].includes(x))return 'FAIL';if(['PARTIAL','PART','HALF'].includes(x))return 'PARTIAL';return ''}
function normPriority(s){const x=String(s??'').trim().toLowerCase().replace(/[\s_-]/g,'');if(x==='regular')return 'Regular';if(x==='mustdo')return 'Must-Do';if(x==='late')return 'Late';return ''}
function safeNum(v){const n=Number(v);return Number.isFinite(n)?n:0}
function compute(state, todayStr=today()){
 const cards=Array.isArray(state.master)?state.master:[], revisions=Array.isArray(state.revisions)?state.revisions:[];
 const idCounts=new Map(); cards.forEach(c=>idCounts.set(String(c.id),(idCounts.get(String(c.id))||0)+1));
 const cardById=new Map(); cards.forEach(c=>{if(!cardById.has(String(c.id)))cardById.set(String(c.id),c)});
 const revRows=revisions.map((r,i)=>({...r,_i:i,_key:(validDate(r.logDate)&&Number.isInteger(Number(r.impId))&&normResult(r.result)?Math.floor(date(r.logDate).getTime()/86400000)*100000+i+2:0),_result:normResult(r.result)}));
 const reviewRank=new Map();const rankCount=new Map();[...revRows].filter(r=>r._key).sort((a,b)=>a._key-b._key).forEach(r=>{const k=String(r.impId),n=(rankCount.get(k)||0)+1;rankCount.set(k,n);reviewRank.set(r._i,n)});
 const revById=new Map(); revRows.forEach(r=>{const k=String(r.impId);if(!revById.has(k))revById.set(k,[]);revById.get(k).push(r)});
 const revisionComputed=revRows.map(r=>{const issues=[];if(r.logDate&&!validDate(r.logDate))issues.push('Invalid date');if(r.impId!==''&&r.impId!=null&&!cardById.has(String(r.impId)))issues.push('Unknown IMP ID');if(r.result&&!r._result)issues.push('Invalid result');if(r.logDate&&r.impId!==''&&r.impId!=null&&r.result&&!r._key)issues.push('Revision ignored until fixed');const c=cardById.get(String(r.impId));return {...r,subjectAuto:c?.subject||'',unitAuto:c?.unit||'',questionAuto:c?.question||'',reviewNo:reviewRank.get(r._i)||'',resultNorm:r._result,key:r._key||'',dataQuality:issues.join(' | '),valid:!!r._key&&!!c};});
 const computed=cards.map(c=>{
  const issues=[]; const id=Number(c.id); if(!Number.isInteger(id)||id<1)issues.push('Invalid ID');
  if(c.entryDate&&!validDate(c.entryDate))issues.push('Invalid entry date');
  if(!String(c.subject||'').trim())issues.push('Missing subject'); if(!String(c.unit||'').trim())issues.push('Missing unit');
  if(c.subject&&c.unit&&state.units&&Array.isArray(state.units[c.subject])&&!state.units[c.subject].includes(c.unit))issues.push("Unit doesn't match subject");
  if(!String(c.question||'').trim())issues.push('Missing question'); if(!normPriority(c.priority))issues.push('Invalid priority'); if((idCounts.get(String(c.id))||0)>1)issues.push('Duplicate ID');
  const dq=issues.map(x=>'⚠️ '+x).join(' | ');
  const validRevs=(revById.get(String(c.id))||[]).filter(r=>r._key).sort((a,b)=>a._key-b._key);
  const last=validRevs[validRevs.length-1]||null, lastKey=last?last._key:0;
  const lastFailKey=validRevs.filter(r=>r._result==='FAIL').reduce((m,r)=>Math.max(m,r._key),0);
  let streak=0;if(last&&last._result!=='FAIL')streak=new Set(validRevs.filter(r=>r._result==='PASS'&&r._key>lastFailKey&&r._key<=lastKey).map(r=>r.logDate)).size;
  const stage=validRevs.length?Math.max(1,Math.min(5,streak)):0;
  const priority=normPriority(c.priority);let rawDue='';
  if(!dq&&priority){if(!last)rawDue=addDays(c.entryDate||todayStr,1+(id%5));else {const it=(state.intervals||DEFAULT_INTERVALS)[priority]||DEFAULT_INTERVALS[priority];const days=last._result==='FAIL'?safeNum(it.fail):last._result==='PARTIAL'?safeNum(it.partial):safeNum((it.pass||DEFAULT_INTERVALS[priority].pass)[Math.max(0,Math.min(4,stage-1))]);rawDue=addDays(last.logDate,days)}}
  let nextDue='';const exam=state.examDate&&validDate(state.examDate)?state.examDate:'';
  if(!dq&&rawDue){if(!exam)nextDue=rawDue;else {const cutoff=addDays(exam,-1);if(last&&last.logDate>=cutoff&&last._result==='PASS')nextDue='';else nextDue=rawDue>cutoff?cutoff:rawDue}}
  const status=dq?dq.slice(0,80):last&&!last._result?'⚠️ Invalid revision result':!nextDue?'✓ Done':nextDue<=todayStr?`DUE${diffDays(todayStr,nextDue)>0?` (${diffDays(todayStr,nextDue)}d overdue)`:''}`:`Due ${formatShort(nextDue)}${exam&&rawDue>addDays(exam,-1)?' ⚠️ capped':''}`;
  const recent=[...validRevs].sort((a,b)=>b._key-a._key).slice(0,3).map(r=>r._result);
  const attempts=validRevs.length,passes=validRevs.filter(r=>r._result==='PASS').length,fails=validRevs.filter(r=>r._result==='FAIL').length;
  return {...c,id,priorityNormLabel:priority,priorityNormNum:priority==='Regular'?1:priority==='Must-Do'?2:priority==='Late'?3:'',lastRevised:last?.logDate||'',lastResult:last?._result||'',totalAttempts:attempts,passes,fails,streak,stage,rawDue,nextDue,status,last1:recent[0]||'',last2:recent[1]||'',last3:recent[2]||'',trend:recent.join(' → '),failRate:attempts?fails/attempts:0,daysLeft:nextDue?diffDays(nextDue,todayStr):'',dataQuality:dq,lastKey,lastFailKey,validRevs,due:!!nextDue&&nextDue<=todayStr,new:!last};
 });
 const validMaster=computed.filter(c=>!c.dataQuality), due=computed.filter(c=>!c.dataQuality&&c.nextDue&&c.nextDue<=todayStr), week=computed.filter(c=>!c.dataQuality&&c.nextDue&&c.nextDue<=addDays(todayStr,7));
 const totalRevs=revisionComputed.filter(r=>r.logDate&&r.impId!==''&&r.result).length;
 const validResults=revisionComputed.filter(r=>r.resultNorm).length, passCount=revisionComputed.filter(r=>r.resultNorm==='PASS').length;
 const subjects=SUBJECTS.map(subject=>{const a=computed.filter(c=>c.subject===subject);return {subject,total:a.length,dueToday:a.filter(c=>c.due).length,mastered:a.filter(c=>c.stage>=4).length,inProgress:a.filter(c=>c.stage>=2&&c.stage<=3).length,new:a.filter(c=>c.new).length,fails:a.reduce((n,c)=>n+c.fails,0),progress:a.length?Math.round(a.filter(c=>c.stage>=4).length/a.length*100):0}});
 const failRows=computed.filter(c=>c.fails>0).sort((a,b)=>b.fails-a.fails||b.failRate-a.failRate).slice(0,20);
 return {cards:computed,revisions:revisionComputed,due,week,subjects,failRows,dashboard:{totalIMPs:cards.length,dueToday:due.length,dueNext7:week.length,neverRevised:computed.filter(c=>c.new).length,mastered:computed.filter(c=>c.stage>=4).length,inProgress:computed.filter(c=>c.stage>=2&&c.stage<=3).length,totalRevs,overallPassRate:validResults?passCount/validResults:0,dataIssues:computed.filter(c=>c.dataQuality).length,revisionIssues:revisionComputed.filter(r=>r.dataQuality).length,capacityMasterWarn:cards.length>=950,capacityMasterFull:cards.length>=1000,capacityRevWarn:revisions.length>=9500,capacityRevFull:revisions.length>=9999},today:todayStr};
}
function formatShort(s){const d=date(s);return d?`${String(d.getDate()).padStart(2,'0')}-${['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][d.getMonth()]}`:s}
function validateCard(c, state, excludeId){const errors=[];if(!Number.isInteger(Number(c.id))||Number(c.id)<1)errors.push('ID must be a positive whole number.');if(!c.entryDate||!validDate(c.entryDate))errors.push('Entry date must be a real date.');if(!SUBJECTS.includes(c.subject))errors.push('Choose a subject from the workbook subject list.');if(!c.unit?.trim())errors.push('Unit is required.');if(!c.question?.trim())errors.push('Question is required.');if(!normPriority(c.priority))errors.push('Priority must be Regular, Must-Do or Late.');if(state.units?.[c.subject]&&!state.units[c.subject].includes(c.unit))errors.push('Choose a configured unit or add it under Unit Management first.');if(state.master.some(x=>Number(x.id)===Number(c.id)&&Number(x.id)!==Number(excludeId)))errors.push('That ID is already used.');return errors}
function validateRevision(r,state,excludeIndex){const errors=[];if(!validDate(r.logDate))errors.push('Revision date must be a real date.');if(!Number.isInteger(Number(r.impId))||!state.master.some(c=>Number(c.id)===Number(r.impId)))errors.push('Choose an existing IMP ID.');if(!normResult(r.result))errors.push('Result must be PASS, FAIL or PARTIAL.');return errors}
function csvEscape(v){let s=String(v??'');if(/^[\s\u0000-\u0020]*[=+@\-]/.test(s))s="'"+s;return '"'+s.replace(/"/g,'""')+'"'}
function toCSV(rows,headers){return [headers.map(csvEscape).join(','),...rows.map(r=>headers.map(h=>csvEscape(r[h])).join(','))].join('\r\n')}
function parseCSV(text){const rows=[];let row=[],v='',q=false;for(let i=0;i<text.length;i++){const c=text[i];if(q){if(c==='"'&&text[i+1]==='"'){v+='"';i++}else if(c==='"')q=false;else v+=c}else if(c==='"')q=true;else if(c===','){row.push(v);v=''}else if(c==='\n'){row.push(v.replace(/\r$/,''));rows.push(row);row=[];v=''}else v+=c}if(v.length||row.length){row.push(v.replace(/\r$/,''));rows.push(row)}return rows.filter(r=>r.some(x=>String(x).trim()!==''))}
const api={SUBJECTS,DEFAULT_INTERVALS,today,iso,date,validDate,addDays,diffDays,normResult,normPriority,compute,validateCard,validateRevision,csvEscape,toCSV,parseCSV,formatShort};
if(typeof module!=='undefined'&&module.exports)module.exports=api;root.MBBSRecallEngine=api;
})(typeof globalThis!=='undefined'?globalThis:this);
