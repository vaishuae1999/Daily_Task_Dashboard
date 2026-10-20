import ExcelJS from 'exceljs';
import {createHash} from 'node:crypto';
import fs from 'node:fs/promises';
import path from 'node:path';
const [input,output='data/import']=process.argv.slice(2);
if(!input)throw Error('Usage: npm run prepare-import -- /path/to/source.xlsx data/import');
const norm=s=>String(s||'').toLowerCase().replace(/[^a-z0-9]/g,'');
const aliases={orderId:['orderid','ordid','waybillnoordid','orderidawbno'],date:['date','dateofcontact','dateofrto','delivereddate','dateofcalling'],customer:['customername','custname','customerid','customeridname'],email:['email','customeremail'],phone:['contactno'],issue:['issue','commentdm','issueactionpoints','issuereasonforcancellation','feedback','customerfeedback','complaint','issuecomment'],action:['actionpoints'],agent:['doneby','answeredby'],status:['status'],channel:['source'],awb:['awb','awbno'],followUpDate:['dateoffollowup'],callingDate:['callingdate','dateofcalling'],address:['completeaddressafterconfirmation']};
const defaultChannels={'Bitespeed_Whatsapp':'WhatsApp','Emails':'Email','InstagramFb DM & CMTS':'Social (unspecified)','Present website data RTO Calls':'RTO Call','All RTO Calls':'RTO Call','RTO List sent by Yash':'RTO Call','cancelled orders':'Cancellation Call','Giveaway contest':'Giveaway Call','High Risk Order Confirmation ca':'Confirmation Call','Orders More than 5000':'High-value order','remainders':'Reminder','Proffesional Service':'Professional Service','dashboard exceptions  NDR':'NDR','Amazon comments customer data':'Amazon Comment','call enquiriescomplaints':'Call','Fb_commentsdms':'Facebook (unspecified)'};
function text(cell){if(cell.value instanceof Date)return cell.value.toISOString().slice(0,10);return String(cell.text??'').trim();}
function date(value){if(!value||value==='--')return '';let s=value;if(/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(s)){const [d,m,y]=s.split('/');s=`${y}-${m.padStart(2,'0')}-${d.padStart(2,'0')}`;}if(!/^\d{4}-\d{2}-\d{2}$/.test(s))throw Error('Unrecognized date: '+value);const d=new Date(s+'T00:00:00Z');if(!Number.isFinite(d.getTime())||d.toISOString().slice(0,10)!==s)throw Error('Invalid date: '+value);return s;}
const w=new ExcelJS.Workbook();await w.xlsx.readFile(input);const staged=[],counts=[];
for(const ws of w.worksheets){const headers=[];ws.getRow(1).eachCell({includeEmpty:true},(c,i)=>headers[i-1]=text(c));const headerNorm=headers.map(norm);const mapping={};for(const [key,names] of Object.entries(aliases))mapping[key]=headerNorm.findIndex(h=>names.includes(h));const headerFound=mapping.customer>=0;let count=0;
ws.eachRow((row,rowNumber)=>{if(headerFound&&rowNumber===1)return;const raw={};row.eachCell((c,i)=>{const v=text(c);if(v)raw[(headers[i-1]||'Column')+' ['+i+']']=v;});if(!Object.keys(raw).length)return;count++;const t={};for(const k of Object.keys(aliases))t[k]=mapping[k]>=0?text(row.getCell(mapping[k]+1)):'';t.channel=t.channel||defaultChannels[ws.name]||'Unclassified';t.sourceSheet=ws.name;t.sourceRow=String(rowNumber);t.raw=JSON.stringify(raw);t.sourceKey=createHash('sha256').update(JSON.stringify([ws.name,rowNumber,raw])).digest('hex');const problems=[];
if(!headerFound)problems.push('No recognized header row: manual mapping required');
for(const k of ['customer','issue'])if(!t[k]||t[k]==='--')problems.push('Missing '+k);
for(const k of ['date','callingDate','followUpDate'])try{t[k]=date(t[k]);}catch(e){problems.push(e.message);}
if(t.agent&&!['Akshay','Vaishnavi'].includes(t.agent))problems.push('Unknown agent');
if(t.email&&!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(t.email))problems.push('Invalid email');
if(t.channel.includes('unspecified')||t.channel==='Unclassified')problems.push('Confirm channel');
if(t.raw.length>40000||Object.values(t).some(v=>String(v).length>4000&&v!==t.raw))problems.push('Field exceeds backend size limit');
staged.push({task:t,problems});});counts.push({sheet:ws.name,sourceRows:count});}
// Flag possible cross-tab duplicates for human review; do not silently delete them.
const groups=new Map();for(const entry of staged){const t=entry.task;if(!t.orderId||t.orderId==='--'||!t.customer||!t.issue)continue;const key=JSON.stringify([t.orderId.replace(/^#\s*/,''),norm(t.customer),norm(t.issue)]);if(!groups.has(key))groups.set(key,[]);groups.get(key).push(entry);}
for(const group of groups.values())if(new Set(group.map(e=>e.task.sourceSheet)).size>1)group.forEach(e=>e.problems.push('Possible cross-tab duplicate: compare source records'));
const ready=staged.filter(e=>!e.problems.length).map(e=>e.task),review=staged.filter(e=>e.problems.length);await fs.mkdir(output,{recursive:true});
// Fresh output folder prevents old batches from an earlier run being mistaken for current data.
const old=await fs.readdir(output);if(old.some(n=>/^ready-\d+\.json$|^review\.json$|^summary\.json$/.test(n)))throw Error('Output contains a previous import. Choose a new directory.');
for(let i=0;i<ready.length;i+=500)await fs.writeFile(path.join(output,`ready-${String(i/500+1).padStart(3,'0')}.json`),JSON.stringify(ready.slice(i,i+500),null,2));
await fs.writeFile(path.join(output,'review.json'),JSON.stringify(review,null,2));const summary={sheets:counts,totalRows:staged.length,ready:ready.length,needsReview:review.length,notes:['Source status preserved; no automatic Done conversion.','Blank dates and agents remain blank.','No records uploaded. Ready files must be reviewed before import.','Source keys prevent reimport of the same snapshot rows. Changed/reordered source rows require reconciliation.']};await fs.writeFile(path.join(output,'summary.json'),JSON.stringify(summary,null,2));console.log(JSON.stringify(summary,null,2));
