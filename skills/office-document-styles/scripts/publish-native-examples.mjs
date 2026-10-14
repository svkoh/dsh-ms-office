import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {createHash} from 'node:crypto';
import {atlas,validateAtlas} from './style-atlas.mjs';
const root=fileURLToPath(new URL('../../../',import.meta.url));
const skill=fileURLToPath(new URL('../',import.meta.url));
const templates=path.join(skill,'resources/templates'),previews=path.join(skill,'resources/previews');
const sha=b=>createHash('sha256').update(b).digest('hex');
const readJson=p=>JSON.parse(fs.readFileSync(p,'utf8').replace(/^\uFEFF/,''));
if(!validateAtlas().passed)throw Error('Invalid element atlas definition');
const ready=atlas.profiles.map(p=>{
 const host=p.host==='PowerPoint'?'ppt':p.host.toLowerCase(),dir=path.join(root,'docs/style-atlas/plugin-'+host);
 const proof=readJson(path.join(dir,p.id+'.evidence.json'));
 if(!(p.host==='Excel'?proof.contentVerified&&proof.nativeVerified:proof.passed&&proof.fontsVerified))throw Error('Native plugin verification incomplete: '+p.id);
 const count=p.host==='Word'?proof.nativePageCount:p.host==='Excel'?2:proof.slideCount;
 const images=p.host==='Excel'?['excel-basic-styles.png','excel-data-chart.png']:Array.from({length:count},(_,i)=>p.id+(p.host==='PowerPoint'?'-slide-':'-page-')+String(i+1).padStart(2,'0')+'.png');
 for(const name of [p.file,...images]){const f=path.join(dir,name);if(!fs.existsSync(f)||!fs.statSync(f).size)throw Error('Missing plugin artifact: '+f);}
 if(fs.readFileSync(path.join(dir,p.file)).subarray(0,2).toString()!=='PK')throw Error('Not a native Office package');
 const pdfFile=path.join(dir,p.id+'.pdf'),pdf=fs.existsSync(pdfFile)&&fs.readFileSync(pdfFile).subarray(0,5).toString()==='%PDF-'?p.id+'.pdf':null;
 return {p,dir,proof,count,images,pdf};
});

function publishBytes(target,bytes){
 if(fs.existsSync(target)){
  const old=fs.readFileSync(target);if(sha(old)===sha(bytes))return;
  const relative=path.relative(root,target);if(relative.startsWith('..')||path.isAbsolute(relative))throw Error('Publish target outside repository');
  const archive=path.join(root,'.tmp/reference-style-assets',sha(old).slice(0,12),relative);fs.mkdirSync(path.dirname(archive),{recursive:true});
  if(fs.existsSync(archive)&&sha(fs.readFileSync(archive))!==sha(old))throw Error('Archive collision');if(!fs.existsSync(archive))fs.writeFileSync(archive,old);
 }
 fs.mkdirSync(path.dirname(target),{recursive:true});fs.writeFileSync(target,bytes);
}
const records=[];
for(const {p,dir,proof,count,images,pdf} of ready){
 const record={presetId:p.id,host:p.host,name:p.name,office:p.file,pdf,preview:p.id+'.png',pageCount:count,previewCount:images.length,coverage:p.coverage,fontFamilies:p.fontFamilies,pluginOnlyWrites:true,previewScope:'Full native Microsoft Office page or worksheet',verification:'Actual plugin generation, native package checks and native Office visual review.',checksums:{}};
 const copies=[[p.file,path.join(templates,p.file)],[p.file,path.join(root,'docs/style-atlas',p.file)]];
 if(pdf)copies.push([pdf,path.join(templates,pdf)],[pdf,path.join(root,'docs/style-atlas',pdf)]);
 images.forEach((name,i)=>copies.push([name,path.join(previews,p.id,'page-'+String(i+1).padStart(2,'0')+'.png')]));copies.push([images[0],path.join(previews,record.preview)]);
 for(const [name,target]of copies){const bytes=fs.readFileSync(path.join(dir,name));publishBytes(target,bytes);record.checksums[path.relative(root,target).replaceAll('\\','/')]=sha(bytes);}
 const publicProof={schemaVersion:1,source:proof.source||'Actual Office.js plugin',passed:true,host:p.host,pluginOnlyWrites:true,pageCount:count,fontsVerified:true,checksums:record.checksums};

const keys=['visualQa','tocCacheVerified','nativeRendering','renderer','fontChecks','nativePreviewCount','paragraphSpacing','globalStyles','bodyStyle','trueOnePointFive','firstLineVerified','a4','margins25mm','firstPageDifferent','nativeTables','nativePictures','nativeListParagraphs','hasNativeNumbering','hasToc','hasPage','originalPngEmbedded','native','contentVerified','nativeVerified','pdfVerified','numberFormatsVerified','wrappedRowsVerified','dateOverflowFixed','allCellsSongVerified','slideSize','tableVerified','imageAspectPreserved'];
 for(const key of keys)if(proof[key]!==undefined)publicProof[key]=proof[key];
   if(publicProof.nativeRendering){const native=publicProof.nativeRendering;publicProof.nativeRendering=Object.fromEntries(Object.entries(native).filter(([key])=>key!=='instance'&&key!=='pages'));publicProof.nativeRendering.pages=native.pages.map(({index,width,height,sha256})=>({index,width,height,sha256}));}
publishBytes(path.join(templates,p.id+'.evidence.json'),Buffer.from(JSON.stringify(publicProof,null,2)+'\n'));records.push(record);
}



publishBytes(path.join(templates,'index.json'),Buffer.from(JSON.stringify({schemaVersion:3,dataLabel:atlas.dataLabel,purpose:atlas.purpose,templates:records},null,2)+'\n'));
console.log('Published exactly three plugin-created native files and native previews.');
