'use strict';
const assert=require('node:assert/strict');
const zlib=require('node:zlib');
require('../xlsx-import.js');
function zip(files){
 const locals=[],centrals=[];let offset=0;
 for(const [name,content] of Object.entries(files)){
  const filename=Buffer.from(name),raw=Buffer.from(content),packed=zlib.deflateRawSync(raw);
  const local=Buffer.alloc(30);local.writeUInt32LE(0x04034b50,0);local.writeUInt16LE(20,4);local.writeUInt16LE(8,8);local.writeUInt32LE(packed.length,18);local.writeUInt32LE(raw.length,22);local.writeUInt16LE(filename.length,26);
  locals.push(local,filename,packed);
  const central=Buffer.alloc(46);central.writeUInt32LE(0x02014b50,0);central.writeUInt16LE(20,4);central.writeUInt16LE(20,6);central.writeUInt16LE(8,10);central.writeUInt32LE(packed.length,20);central.writeUInt32LE(raw.length,24);central.writeUInt16LE(filename.length,28);central.writeUInt32LE(offset,42);
  centrals.push(central,filename);offset+=local.length+filename.length+packed.length;
 }
 const centralSize=centrals.reduce((n,b)=>n+b.length,0),count=Object.keys(files).length,eocd=Buffer.alloc(22);eocd.writeUInt32LE(0x06054b50,0);eocd.writeUInt16LE(count,8);eocd.writeUInt16LE(count,10);eocd.writeUInt32LE(centralSize,12);eocd.writeUInt32LE(offset,16);
 return new Uint8Array(Buffer.concat([...locals,...centrals,eocd]));
}
const ss=(r,c,v)=>`<c r="${c}${r}" t="inlineStr"><is><t>${String(v).replace(/&/g,'&amp;').replace(/</g,'&lt;')}</t></is></c>`;
const num=(r,c,v)=>`<c r="${c}${r}"><v>${v}</v></c>`;
const sheet=rows=>`<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${rows.map((cells,i)=>`<row r="${i+1}">${cells.join('')}</row>`).join('')}</sheetData></worksheet>`;
async function main(){
 const subjects=['Medicine','Surgery','Obstetrics','Gynecology','Pediatrics','Ophthalmology','ENT','Orthopedics','Dermatology','Anesthesia','Radiology','Psychiatry'];
 const master=sheet([[],[],[num(3,'A',7),ss(3,'B','2026-10-01'),ss(3,'C','Medicine'),ss(3,'D','Cardiology'),ss(3,'E','Test question'),ss(3,'F','Key point'),ss(3,'G','Source'),ss(3,'H','Note'),ss(3,'I','Must-Do')]]);
 const revisions=sheet([[ss(1,'A','Log Date')],[ss(2,'A','2026-10-02'),num(2,'B',7),ss(2,'C','Pass'),ss(2,'D','Reviewed')]]);
 const unitRows=[subjects.map((s,i)=>ss(1,String.fromCharCode(65+i),s)),subjects.map((s,i)=>ss(2,String.fromCharCode(65+i),i===0?'Cardiology':'Unit A'))];
 const units=sheet(unitRows);
 const intRows=[];for(let r=1;r<=11;r++){let cells=[];if(r===5){cells=[ss(5,'A','Fail'),num(5,'B',1),num(5,'C',1),num(5,'D',1)]}if(r===6){cells=[ss(6,'A','Partial'),num(6,'B',2),num(6,'C',2),num(6,'D',2)]}if(r>=7&&r<=11){cells=[num(r,'B',r-4),num(r,'C',r-5),num(r,'D',r-5)]}intRows.push(cells)}
 const intervals=sheet(intRows),dashboard=sheet([[],[],[ss(3,'A','Exam Date')]]);
 const names=['Dashboard','Master Entry','Revision Log','Units Config','Settings - Intervals'],targets=['worksheets/sheet1.xml','worksheets/sheet2.xml','worksheets/sheet3.xml','worksheets/sheet4.xml','worksheets/sheet5.xml'];
 const workbook=`<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${names.map((n,i)=>`<sheet name="${n}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')}</sheets></workbook>`;
 const rels=`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${targets.map((t,i)=>`<Relationship Id="rId${i+1}" Target="${t}" Type="worksheet"/>`).join('')}</Relationships>`;
 const fixture=zip({'xl/workbook.xml':workbook,'xl/_rels/workbook.xml.rels':rels,'xl/worksheets/sheet1.xml':dashboard,'xl/worksheets/sheet2.xml':master,'xl/worksheets/sheet3.xml':revisions,'xl/worksheets/sheet4.xml':units,'xl/worksheets/sheet5.xml':intervals});
 const result=await MBBSXLSXImporter.parse(fixture);
 assert.equal(result.counts.master,1);assert.equal(result.counts.revisions,1);assert.equal(result.counts.subjects,12);assert.equal(result.master[0].id,7);assert.equal(result.master[0].subject,'Medicine');assert.equal(result.master[0].unit,'Cardiology');assert.equal(result.master[0].priority,'Must-Do');assert.equal(result.revisions[0].result,'Pass');assert.equal(result.intervals.Regular.pass[0],3);assert.equal(result.intervals['Must-Do'].pass[0],2);assert.equal(result.units.Medicine[0],'Cardiology');assert.equal(result.examDate,'');assert.equal(MBBSXLSXImporter.dateValue('46296'),'2026-10-01');
 console.log('PASS — synthetic .xlsx import: 12 assertions');
}
main().catch(e=>{console.error(e);process.exit(1)});
