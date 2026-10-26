const SHEET_NAME = 'ทะเบียนครุภัณฑ์';
const HEADERS = ['ลำดับที่','รหัส','รายการ','รุ่น/ยี่ห้อ','วันเดือนปีที่ได้รับ','ปีที่ได้รับ','สถานที่ติดตั้ง','ประเภทครุภัณฑ์','จำนวน','ราคา','สถานะ','หมายเหตุ','ภาพประกอบ','วันที่สร้าง','วันที่แก้ไข'];
const TYPES = ['ครุภัณฑ์สำนักงาน','ครุภัณฑ์คอมพิวเตอร์','ครุภัณฑ์ยานพาหนะและขนส่ง','ครุภัณฑ์ไฟฟ้าและวิทยุ','ครุภัณฑ์การศึกษา','ครุภัณฑ์วิทยาศาสตร์การแพทย์','ประเภทอื่น ๆ'];
const STATUSES = ['ใช้งานได้','ชำรุด','รอซ่อม','ส่งซ่อม','เสื่อมสภาพ','รอจำหน่าย','จำหน่าย','สูญหาย','ไม่พบตัว','อื่น ๆ'];

function doGet() {
  setupSheet_();
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ทะเบียนครุภัณฑ์ สำนักงานพัฒนาชุมชนจังหวัดพังงา')
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}
function setupSheet_() {
  const p=PropertiesService.getScriptProperties();
  let id=p.getProperty('SPREADSHEET_ID');
  let ss;
  if(id){ try { ss=SpreadsheetApp.openById(id); } catch(e) {} }
  if(!ss){ ss=SpreadsheetApp.create('ทะเบียนครุภัณฑ์ สำนักงานพัฒนาชุมชนจังหวัดพังงา'); p.setProperty('SPREADSHEET_ID',ss.getId()); }
  let sh=ss.getSheetByName(SHEET_NAME);
  if(!sh) sh=ss.insertSheet(SHEET_NAME);
  if(sh.getLastRow()===0){ sh.getRange(1,1,1,HEADERS.length).setValues([HEADERS]); sh.setFrozenRows(1); sh.getRange(1,1,1,HEADERS.length).setFontWeight('bold').setBackground('#173f35').setFontColor('#ffffff'); }
  return sh;
}
function getData() {
  const sh=setupSheet_(), values=sh.getDataRange().getValues();
  if(values.length<2) return {records:[],types:TYPES,statuses:STATUSES,spreadsheetUrl:SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')).getUrl()};
  const records=values.slice(1).filter(r=>r[1]!=='' || r[2]!=='').map((r,i)=>({
    row:i+2,seq:r[0],code:String(r[1]||''),item:String(r[2]||''),brand:String(r[3]||''),
    received:r[4] instanceof Date?Utilities.formatDate(r[4],Session.getScriptTimeZone(),'yyyy-MM-dd'):String(r[4]||''),
    year:String(r[5]||''),location:String(r[6]||''),type:String(r[7]||''),quantity:Number(r[8]||0),price:Number(r[9]||0),
    status:String(r[10]||''),note:String(r[11]||''),image:String(r[12]||''),created:String(r[13]||''),updated:String(r[14]||'')
  }));
  records.sort((a,b)=>thaiCompare_(a.item,b.item) || (dateValue_(b.received)-dateValue_(a.received)));
  records.forEach((r,i)=>r.seq=i+1);
  return {records,types:TYPES,statuses:STATUSES,spreadsheetUrl:SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')).getUrl()};
}
function saveRecord(payload, token) {
  requireAdmin_(token);
  const sh=setupSheet_(), data=normalize_(payload);
  if(!data.code || !data.item) throw new Error('กรุณากรอกรหัสและรายการครุภัณฑ์');
  const values=sh.getDataRange().getValues();
  let row=-1;
  for(let i=1;i<values.length;i++) if(String(values[i][1])===data.code){row=i+1;break;}
  const now=new Date();
  const existingCreated=row>0?values[row-1][13]:now;
  const out=[0,data.code,data.item,data.brand,data.received?new Date(data.received+'T00:00:00'): '',data.year,data.location,data.type,data.quantity,data.price,data.status,data.note,data.image,existingCreated,now];
  if(row>0) sh.getRange(row,1,1,HEADERS.length).setValues([out]);
  else sh.appendRow(out);
  sh.getRange(2,5,Math.max(sh.getLastRow()-1,1),1).setNumberFormat('dd/MM/yyyy');
  return {ok:true,message:row>0?'แก้ไขข้อมูลเรียบร้อย':'บันทึกข้อมูลเรียบร้อย'};
}
function deleteRecord(code, token) {
  requireAdmin_(token);
  const sh=setupSheet_(), values=sh.getDataRange().getValues();
  for(let i=1;i<values.length;i++) if(String(values[i][1])===String(code)){sh.deleteRow(i+1);return {ok:true};}
  throw new Error('ไม่พบรหัสครุภัณฑ์นี้');
}
function verifyAdmin(password) {
  const saved=PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if(!saved) throw new Error('ผู้ดูแลระบบยังไม่ได้ตั้งรหัสผ่านใน Script Properties');
  if(String(password)!==saved) throw new Error('รหัสผ่านไม่ถูกต้อง');
  const token=Utilities.getUuid();
  CacheService.getScriptCache().put('admin_'+token,'1',1800);
  return {token:token};
}
function requireAdmin_(token) {
  if(!token || CacheService.getScriptCache().get('admin_'+token)!=='1') throw new Error('สิทธิ์ผู้ดูแลระบบหมดอายุ กรุณาเข้าสู่ระบบใหม่');
}
function uploadImage(fileName,mimeType,dataUrl,token) {
  requireAdmin_(token);
  if(!/^image\/(png|jpeg|jpg|gif|webp)$/.test(String(mimeType))) throw new Error('รองรับไฟล์ภาพ PNG, JPG, GIF หรือ WEBP เท่านั้น');
  if(!String(dataUrl).startsWith('data:'+mimeType+';base64,')) throw new Error('รูปแบบไฟล์ไม่ถูกต้อง');
  const base64=String(dataUrl).split(',')[1]||'';
  if(base64.length>7*1024*1024) throw new Error('ไฟล์ภาพมีขนาดใหญ่เกินกำหนด');
  const blob=Utilities.newBlob(Utilities.base64Decode(base64),mimeType,String(fileName||'ภาพครุภัณฑ์').replace(/[^\wก-๙.-]/g,'_').slice(0,100));
  const name='ภาพประกอบทะเบียนครุภัณฑ์ พังงา', fs=DriveApp.getFoldersByName(name), folder=fs.hasNext()?fs.next():DriveApp.createFolder(name);
  const file=folder.createFile(blob); return {url:file.getUrl(),name:file.getName()};
}
function normalize_(p) {
  p=p||{};
  return {code:String(p.code||'').trim(),item:String(p.item||'').trim(),brand:String(p.brand||'').trim(),
    received:String(p.received||''),year:String(p.year||'').trim(),location:String(p.location||'').trim(),
    type:String(p.type||'').trim(),quantity:Math.max(0,Number(p.quantity||0)),price:Math.max(0,Number(p.price||0)),
    status:String(p.status||'').trim(),note:String(p.note||'').trim(),image:String(p.image||'').trim()};
}
function thaiCompare_(a,b){ return String(a).localeCompare(String(b),'th'); }
function dateValue_(s){ if(!s)return 0; const d=new Date(s); return isNaN(d)?0:d.getTime(); }
function exportReport(format, filters) {
  const data=getData().records, f=filters||{};
  const filtered=data.filter(r=>(!f.type||r.type===f.type)&&(!f.status||r.status===f.status)&&(!f.query||[r.code,r.item,r.brand,r.location].join(' ').toLowerCase().includes(String(f.query).toLowerCase())));
  const rows=[HEADERS.slice(0,13),...filtered.map((r,i)=>[i+1,r.code,r.item,r.brand,r.received,r.year,r.location,r.type,r.quantity,r.price,r.status,r.note,r.image])];
  const ss=SpreadsheetApp.create('รายงานทะเบียนครุภัณฑ์ชั่วคราว');
  const sh=ss.getSheets()[0]; sh.setName('รายงาน'); sh.getRange(1,1,rows.length,13).setValues(rows);
  sh.getRange(1,1,1,13).setFontWeight('bold').setBackground('#173f35').setFontColor('#ffffff');
  sh.setFrozenRows(1); sh.autoResizeColumns(1,13);
  const mime=format==='pdf'?'application/pdf':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const url='https://docs.google.com/spreadsheets/d/'+ss.getId()+'/export?format='+(format==='pdf'?'pdf':'xlsx')+'&portrait=false&fitw=true&sheetnames=false&printtitle=false&pagenumbers=true&gridlines=true';
  const response=UrlFetchApp.fetch(url,{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true});
  if(response.getResponseCode()!==200){DriveApp.getFileById(ss.getId()).setTrashed(true);throw new Error('ส่งออกรายงานไม่สำเร็จ กรุณาตรวจสอบสิทธิ์การอนุญาตของสคริปต์');}
  const blob=response.getBlob().setName('รายงานทะเบียนครุภัณฑ์.'+(format==='pdf'?'pdf':'xlsx'));
  const file=DriveApp.createFile(blob);
  DriveApp.getFileById(ss.getId()).setTrashed(true);
  return {url:file.getUrl(),name:file.getName()};
}
