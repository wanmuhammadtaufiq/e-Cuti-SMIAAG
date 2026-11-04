// Ujian aliran server Code.gs menggunakan simulasi Apps Script (tiada Sheet/Drive sebenar). Jalankan: node tests/server.test.js
const crypto=require('crypto'),fs=require('fs'),vm=require('vm');
const sheets={};
function mkSheet(name){const rows=[];return {rows,
 getRange(r,c,nr=1,nc=1){return{
  getValues:()=>{const o=[];for(let i=0;i<nr;i++){o.push([]);for(let j=0;j<nc;j++)o[i].push((rows[r-1+i]||[])[c-1+j]??"")}return o},
  setValues:v=>{v.forEach((row,i)=>row.forEach((x,j)=>{rows[r-1+i]=rows[r-1+i]||[];rows[r-1+i][c-1+j]=x}));return {setFontWeight(){}}},
  setValue:x=>{rows[r-1]=rows[r-1]||[];rows[r-1][c-1]=x},
  clearContent:()=>{for(let i=0;i<nr;i++)rows[r-1+i]=[]},
  setFontWeight:()=>({})}},
 getLastColumn(){return Math.max(1,...rows.map(r=>r.length))},
 getLastRow(){let n=rows.length;while(n>0&&!(rows[n-1]||[]).some(x=>x!==""&&x!==undefined))n--;return n},
 getDataRange(){const lr=Math.max(1,this.getLastRow()),lc=this.getLastColumn();return this.getRange(1,1,lr,lc)},
 appendRow(v){rows.length=this.getLastRow();rows.push(v.map(x=>typeof x==='string'&&/^\d{4}-\d{2}-\d{2}$/.test(x)?new Date(x+"T00:00:00+08:00"):(typeof x==='string'&&x.startsWith("'")?x.slice(1):x)))},
 deleteRow(i){rows.splice(i-1,1)},setFrozenRows(){}}}
const ss={getSheetByName:n=>sheets[n],insertSheet:n=>(sheets[n]=mkSheet(n))};
const cache={};const files={};
const ctx={console,JSON,Date,Math,parseInt,Array,String,
 SpreadsheetApp:{getActiveSpreadsheet:()=>ss,openById:()=>ss,flush(){}},
 Utilities:{getUuid:()=>crypto.randomUUID(),DigestAlgorithm:{SHA_256:1},Charset:{UTF_8:1},
  computeDigest:(a,s)=>[...crypto.createHash('sha256').update(s).digest()].map(b=>b>127?b-256:b),
  formatDate:(d,tz,f)=>{const p=Object.fromEntries(new Intl.DateTimeFormat('en-GB',{timeZone:tz,year:'numeric',month:'2-digit',day:'2-digit',hour:'numeric',minute:'2-digit',hour12:true}).formatToParts(d).map(x=>[x.type,x.value]));
   return f==="yyyy-MM-dd"?`${p.year}-${p.month}-${p.day}`:`${p.day}/${p.month}/${p.year} ${p.hour}.${p.minute}${p.dayPeriod}`},
  base64Decode:s=>Buffer.from(s,'base64'),base64Encode:b=>Buffer.from(b).toString('base64'),newBlob:(d,m,n)=>({d,m,n})},
 PropertiesService:{getScriptProperties:()=>({getProperty:k=>cache['p'+k],setProperty:(k,v)=>cache['p'+k]=v})},
 CacheService:{getScriptCache:()=>({get:k=>cache[k]??null,put:(k,v)=>cache[k]=v,remove:k=>delete cache[k]})},
 LockService:{getScriptLock:()=>({tryLock:()=>true,releaseLock(){}})},
 DriveApp:{getFoldersByName:()=>({hasNext:()=>false}),createFolder:()=>({createFile:b=>{const id='f'+Object.keys(files).length;files[id]=b;return{getId:()=>id}}}),
  getFileById:id=>({getBlob:()=>({getContentType:()=>files[id].m,getBytes:()=>files[id].d})})},
 MailApp:{sendEmail(){}},Logger:{log(){}},
 ContentService:{createTextOutput:s=>({setMimeType(){return JSON.parse(s)}}),MimeType:{JSON:1}}};
vm.createContext(ctx);vm.runInContext(fs.readFileSync(''+require('path').join(__dirname,'..','apps-script','Code.gs')+'','utf8'),ctx);
const post=p=>ctx.doPost({postData:{contents:JSON.stringify(p)}});
ctx.setupAdminPertama();
sheets.DATA_GURU=mkSheet('DATA_GURU');sheets.DATA_GURU.rows.push(["NAMA_GURU","BIL"],["Ali Bin Abu",10]);
let r=post({action:"LOGIN",emel:"admin@smiaag.edu.my",password:"TukarSaya123"});console.log('login',r.success,r.mestiTukarPassword);
let tok=r.token;
console.log('blocked',post({action:"GET_USERS",token:tok}).mustChangePassword);
console.log('wrongold',post({action:"CHANGE_PASSWORD",token:tok,oldPassword:"x",newPassword:"Baru1234"}).message);
console.log('chg',post({action:"CHANGE_PASSWORD",token:tok,oldPassword:"TukarSaya123",newPassword:"Baru1234"}).success);
tok=post({action:"LOGIN",emel:"admin@smiaag.edu.my",password:"Baru1234"}).token;
const guru=[];for(let i=0;i<4;i++){post({action:"ADD_USER",token:tok,emel:`g${i}@x.my`,nama:i?`Guru ${i}`:"Ali Bin Abu",role:"GURU",tempPassword:"TukarSaya123"});
 let t=post({action:"LOGIN",emel:`g${i}@x.my`,password:"TukarSaya123"}).token;post({action:"CHANGE_PASSWORD",token:t,oldPassword:"TukarSaya123",newPassword:"abcdef1"});
 guru.push(post({action:"LOGIN",emel:`g${i}@x.my`,password:"abcdef1"}).token);}
console.log('public submit',post({action:"SUBMIT_LEAVE",application:{}}).message);
const app=(s,e)=>({action:"SUBMIT_LEAVE",application:{startDate:s,endDate:e,reasonCategory:"Kursus / Bengkel",reasonDetails:"x",fileBase64:Buffer.from("PDF").toString('base64'),fileMimeType:"application/pdf",fileName:"a.pdf",relief:[{day:"Isnin",time:"8",class:"=cmd",replacement:"Guru 1"}]}});
for(let i=0;i<4;i++){const t=post({...app("2026-11-09","2026-11-10"),token:guru[i]});console.log('sub',i,t.success,t.id||t.message)}
console.log('overlap',post({...app("2026-11-10","2026-11-11"),token:guru[0]}).message);
console.log('sunday only',post({...app("2026-11-08","2026-11-08"),token:guru[3]}).message);
post({action:"ADD_UNLIMITED_DATE",token:tok,dates:["2026-11-09","2026-11-10","bad"]});
console.log('after unlimited',post({...app("2026-11-09","2026-11-10"),token:guru[3]}).id);
let g=post({action:"GET_LEAVE_APPLICATIONS",token:guru[1]});console.log('guru sees',g.data.map(a=>a.id+':'+a.status+':'+(a.teacherEmail||'-')),g.usage['2026-11-09']);
const id0=g.data[0]?.id;
console.log('approve',post({action:"APPROVE_LEAVE",token:tok,appId:"CT-2026-001",newStatus:"APPROVED"}).success, post({action:"APPROVE_LEAVE",token:tok,appId:"CT-2026-001",newStatus:"APPROVED"}).message);
let pub=post({action:"GET_LEAVE_APPLICATIONS",token:"PUBLIC_GHOST"});console.log('public',JSON.stringify(pub.data));
console.log('att other',post({action:"GET_ATTACHMENT",token:guru[1],appId:"CT-2026-001"}).message,'att own',post({action:"GET_ATTACHMENT",token:guru[0],appId:"CT-2026-001"}).base64);
let ad=post({action:"GET_LEAVE_APPLICATIONS",token:tok});console.log('admin row',JSON.stringify(ad.data[0]));
console.log('dataguru guru',JSON.stringify(post({action:"GET_DATA_GURU",token:guru[0]}).data),JSON.stringify(post({action:"GET_DATA_GURU",token:guru[1]}).data));
// role change invalidates
post({action:"ADD_USER",token:tok,emel:"a2@x.my",nama:"A2",role:"ADMIN",tempPassword:"TukarSaya123"});
let a2=post({action:"LOGIN",emel:"a2@x.my",password:"TukarSaya123"}).token;post({action:"CHANGE_PASSWORD",token:a2,oldPassword:"TukarSaya123",newPassword:"zzzzzz1"});a2=post({action:"LOGIN",emel:"a2@x.my",password:"zzzzzz1"}).token;
post({action:"EDIT_USER",token:tok,targetEmel:"a2@x.my",emel:"a2@x.my",nama:"A2",role:"GURU"});
console.log('demoted',post({action:"GET_USERS",token:a2}).message);
console.log('logout',post({action:"LOGOUT",token:guru[2]}).success,post({action:"GET_DATA_GURU",token:guru[2]}).message);
console.log('sheet relief',sheets.PermohonanCuti.rows[1][10]);
// --- Sabtu ---
// 14/11/2026 = Sabtu
console.log('sabtu biasa',post({...app("2026-11-14","2026-11-14"),token:guru[1]}).message);
console.log('jumaat-isnin hari', (()=>{const r=post({...app("2026-11-13","2026-11-16"),token:guru[1]});return r.id})(), sheets.PermohonanCuti.rows.at(-1)[6]);
console.log('tanda bukan sabtu',post({action:"ADD_UNLIMITED_DATE",token:tok,jenis:"SABTU",dates:["2026-11-20"]}).message);
console.log('tanda sabtu',JSON.stringify(post({action:"ADD_UNLIMITED_DATE",token:tok,jenis:"SABTU",dates:["2026-11-21","2026-11-20"]}).added));
console.log('sabtu program',post({...app("2026-11-21","2026-11-21"),token:guru[3]}).id||'ok');
console.log('list',JSON.stringify(post({action:"GET_UNLIMITED_DATES"})));
console.log('buang',post({action:"REMOVE_UNLIMITED_DATE",token:tok,jenis:"SABTU",date:"2026-11-21"}).success, JSON.stringify(post({action:"GET_UNLIMITED_DATES"}).sabtu));
