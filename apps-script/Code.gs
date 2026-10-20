/* Daily Support Desk. Deploy as USER_ACCESSING, never anonymous. */
const FIELDS = ['id','date','customer','email','phone','orderId','channel','issue','action','agent','status','awb','followUpDate','callingDate','address','sourceSheet','sourceRow','sourceKey','raw','version','updatedAt','updatedBy','deletedAt'];
function config_() {
  const p = PropertiesService.getScriptProperties();
  const admins = (p.getProperty('ADMIN_EMAILS') || '').split(',').map(x=>x.trim().toLowerCase()).filter(Boolean);
  if (!admins.length || !p.getProperty('MASTER_SHEET_ID')) throw Error('Set ADMIN_EMAILS and MASTER_SHEET_ID in Script Properties.');
  return {p,admins,sheetId:p.getProperty('MASTER_SHEET_ID')};
}
function actor_() {
  const c=config_(), email=Session.getActiveUser().getEmail().trim().toLowerCase();
  if (!email || !c.admins.includes(email)) throw Error('Access denied. Sign in with an authorized admin Google account.');
  return email;
}
function lock_(fn) { const lock=LockService.getScriptLock(); lock.waitLock(30000); try{return fn();} finally{lock.releaseLock();} }
function table_() {
  const book=SpreadsheetApp.openById(config_().sheetId), sheet=book.getSheetByName('Tasks');
  if(!sheet)throw Error('Run initialize once from the Apps Script editor.');
  const h=sheet.getRange(1,1,1,FIELDS.length).getValues()[0];
  if(h.join('|')!==FIELDS.join('|'))throw Error('Tasks columns changed. Restore the expected headers before saving.');
  return sheet;
}
function records_(sheet) {
  if(sheet.getLastRow()<2)return [];
  const zone=sheet.getParent().getSpreadsheetTimeZone();
  return sheet.getRange(2,1,sheet.getLastRow()-1,FIELDS.length).getValues().map((r,i)=>({row:i+2,task:Object.fromEntries(FIELDS.map((f,j)=>{
    const v=r[j];
    // google.script.run cannot serialize Date objects, including nested ones.
    if(v instanceof Date)return [f,['date','followUpDate','callingDate'].includes(f)?Utilities.formatDate(v,zone,'yyyy-MM-dd'):v.toISOString()];
    return [f,v==null?'':v];
  }))}));
}
function text_(value,max) { if(value==null)return ''; const t=String(value).trim(); if(t.length>max)throw Error('Field exceeds '+max+' characters.');return t; }
function date_(v) { const s=text_(v,10); if(!s)return ''; if(!/^\d{4}-\d{2}-\d{2}$/.test(s))throw Error('Date must be YYYY-MM-DD.');const d=new Date(s+'T00:00:00Z');if(!Number.isFinite(d.getTime())||d.toISOString().slice(0,10)!==s)throw Error('Invalid calendar date.');return s; }
function clean_(input) {
  const t={};FIELDS.forEach(f=>t[f]=text_(input[f],f==='raw'?40000:4000));
  t.date=date_(input.date);t.followUpDate=date_(input.followUpDate);t.callingDate=date_(input.callingDate);
  if(!t.customer || !t.issue)throw Error('Customer and issue are required. Incomplete legacy rows must be reviewed first.');
  if(t.email&&!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(t.email))throw Error('Invalid email.');
  if(t.agent&&!['Akshay','Vaishnavi'].includes(t.agent))throw Error('Unknown agent.');
  if(!t.channel)throw Error('Channel is required.');
  if(t.raw){JSON.parse(t.raw);}
  return t;
}
// Prefix user-controlled formula-like text to prevent spreadsheet formula execution.
function safeCell_(value) {return typeof value==='string'&&/^[=+@-]/.test(value)?"'"+value:value;}
function values_(task){return FIELDS.map(f=>safeCell_(task[f] == null?'':task[f]));}
function doGet(){actor_();return HtmlService.createHtmlOutputFromFile('Index').setTitle('Daily Support Desk').addMetaTag('viewport','width=device-width, initial-scale=1');}
function initialize(){actor_();return lock_(()=>{const b=SpreadsheetApp.openById(config_().sheetId);let s=b.getSheetByName('Tasks');if(s&&s.getLastRow()) {table_();return 'Already initialized';}if(!s)s=b.insertSheet('Tasks');s.getRange(1,1,1,FIELDS.length).setValues([FIELDS]);s.setFrozenRows(1);s.getRange(1,1,1,FIELDS.length).setFontWeight('bold');return 'Master Tasks tab initialized';});}
function listTasks(includeTrash){actor_();return records_(table_()).map(x=>x.task).filter(t=>includeTrash?!!t.deletedAt:!t.deletedAt);}
function saveTask(input){const actor=actor_();return lock_(()=>{
  const s=table_(),all=records_(s),existing=input.id?all.find(x=>x.task.id===input.id):null;
  if(input.id&&!existing)throw Error('Record not found. Refresh.');
  if(existing&&(Number(existing.task.version)!==Number(input.version)||existing.task.deletedAt))throw Error('Conflict: record changed or is in Trash. Refresh before editing.');
  const t=clean_(input);t.id=existing?existing.task.id:Utilities.getUuid();t.version=existing?Number(existing.task.version)+1:1;t.updatedAt=new Date().toISOString();t.updatedBy=actor;t.deletedAt='';
  // Provenance is server-preserved on edits; users cannot rewrite import identity.
  for(const f of ['sourceKey','sourceSheet','sourceRow','raw'])t[f]=existing?existing.task[f]:'';
  const row=existing?existing.row:s.getLastRow()+1;s.getRange(row,1,1,FIELDS.length).setValues([values_(t)]);return t;
});}
function setTrash(id,version,restore){const actor=actor_();return lock_(()=>{const s=table_(),found=records_(s).find(x=>x.task.id===id);if(!found||Number(found.task.version)!==Number(version))throw Error('Conflict: refresh the record.');const t=found.task;t.deletedAt=restore?'':new Date().toISOString();t.version=Number(t.version)+1;t.updatedAt=new Date().toISOString();t.updatedBy=actor;s.getRange(found.row,1,1,FIELDS.length).setValues([values_(t)]);return {saved:true};});}
function importTasks(rows){const actor=actor_();if(!Array.isArray(rows)||!rows.length||rows.length>500)throw Error('Import 1–500 reviewed records per batch.');return lock_(()=>{const s=table_(),keys=new Set(records_(s).map(x=>x.task.sourceKey).filter(Boolean));let skipped=0;const items=[];
  // Validate the entire batch before writing anything.
  rows.forEach(r=>{const t=clean_(r);if(!/^[a-f0-9]{64}$/.test(t.sourceKey))throw Error('Missing stable source key. Use the supplied import preparation tool.');if(keys.has(t.sourceKey)){skipped++;return;}keys.add(t.sourceKey);t.id=Utilities.getUuid();t.version=1;t.deletedAt='';t.updatedAt=new Date().toISOString();t.updatedBy=actor;items.push(values_(t));});
  if(items.length)s.getRange(s.getLastRow()+1,1,items.length,FIELDS.length).setValues(items);return {imported:items.length,skipped};});}
function backupNow(){actor_();return lock_(()=>backup_());}
function backup_(){const c=config_(),folderId=c.p.getProperty('BACKUP_FOLDER_ID');if(!folderId)throw Error('Set BACKUP_FOLDER_ID first.');const stamp=Utilities.formatDate(new Date(),'Asia/Kolkata','yyyy-MM-dd_HHmmss');const folder=DriveApp.getFolderById(folderId);const copy=DriveApp.getFileById(c.sheetId).makeCopy('Support_Backup_'+stamp,folder);c.p.setProperty('LAST_BACKUP_AT',new Date().toISOString());return {name:copy.getName(),url:copy.getUrl()};}
// Trailing underscore keeps this trigger function inaccessible to google.script.run.
function scheduledBackup_(){lock_(()=>backup_());}
function installWeeklyBackup(){actor_();if(!config_().p.getProperty('BACKUP_FOLDER_ID'))throw Error('Set BACKUP_FOLDER_ID first.');const existing=ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='scheduledBackup_');if(!existing.length)ScriptApp.newTrigger('scheduledBackup_').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(23).inTimezone('Asia/Kolkata').create();return 'Weekly backup configured for Sunday, approximately 11 PM India time.';}
function getStatus(){const email=actor_(),p=config_().p;return {email,role:'Admin',lastBackup:p.getProperty('LAST_BACKUP_AT')||null,storage:'Google Sheet',weeklyBackupInstalled:ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==='scheduledBackup_')};}

// Only already-trashed records can be permanently removed by an authenticated admin.
function deletePermanently(id,version){
  actor_();
  if(typeof id!=='string'||!id||!Number.isInteger(version)||version<1)throw Error('Invalid record reference. Refresh and retry.');
  return lock_(()=>{
    const sheet=table_(),found=records_(sheet).find(x=>x.task.id===id);
    if(!found)throw Error('Record no longer exists. Refresh the list.');
    if(Number(found.task.version)!==version)throw Error('Conflict: record changed. Refresh before deleting.');
    if(!found.task.deletedAt)throw Error('Move the record to Trash before permanently deleting it.');
    sheet.deleteRow(found.row);
    return {deleted:true,id};
  });
}
