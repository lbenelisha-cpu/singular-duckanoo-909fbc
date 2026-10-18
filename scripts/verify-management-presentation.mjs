import assert from 'node:assert/strict'
import fs from 'node:fs/promises'
import {DOMParser,XMLSerializer} from '@xmldom/xmldom'
import JSZip from 'jszip'
import {buildPresentationData,packagingSize} from '../src/management/presentationData.js'
import {generateTemplatePresentation} from '../src/management/templateExport.js'
globalThis.DOMParser=DOMParser;globalThis.XMLSerializer=XMLSerializer
const from='2026-07-01',to='2026-09-30',base={from,to}
const empty=buildPresentationData(base)
assert.equal(empty.total,null);assert.equal(empty.copq,null);assert.equal(empty.periodSafety.firstAid,null)
assert.equal(packagingSize({routingGroup:'42-P-04'}),null)
assert.equal(packagingSize({desc:'PRODUCT 10 LT'}),'10')
const rows=[];for(let month=7;month<=9;month++)for(const [i,size] of ['1','5','10','20'].entries())rows.push({facility:'42',date:`2026-0${month}-01`,packSize:size,qty:(i+1)*1000,shift:{key:['morning','evening','night'][i%3]}})
rows.push({facility:'43',date:'2026-07-01',qty:999999})
const supplements={months:Object.fromEntries([7,8,9].map(m=>[`2026-0${m}`,{plan1:1000,plan5:2000,plan1020:9000,fms1:1100,fms5:2100,fms1020:7000,cost:20000,rft:100,copq:0,siteCopq:0}])),project:{include:true,date:'2026-10'}}
const model=buildPresentationData({...base,rows,supplements});assert.equal(model.total,30000);assert.equal(model.costRate,2);assert.equal(model.copq,0);assert.equal(model.days,3)
assert.throws(()=>buildPresentationData({...base,supplements:{months:{'2026-07':{rft:101}}}}),/לא תקין/)
assert.equal(buildPresentationData({...base,from:'2026-07-02',rows,supplements}).costRate,null)
const result=await generateTemplatePresentation(model,{templateBuffer:await fs.readFile('public/presentation/facility42-template.pptx')})
assert.equal(result.slideCount,20)
const zip=await JSZip.loadAsync(result.bytes),pres=await zip.file('ppt/presentation.xml').async('string')
assert.equal((pres.match(/<p:sldId /g)||[]).length,20)
for(const name of Object.keys(zip.files).filter(x=>/^ppt\/slides\/slide\d+.xml$/.test(x))){const doc=new DOMParser().parseFromString(await zip.file(name).async('string'),'application/xml');for(const table of Array.from(doc.getElementsByTagNameNS('http://schemas.openxmlformats.org/drawingml/2006/main','tbl')))assert.equal(table.getElementsByTagNameNS('*','extLst').length,0)}
const chart=await zip.file('ppt/charts/chart100.xml').async('string');assert.match(chart,/10000/);assert.doesNotMatch(chart,/739112/)
const sparse=await generateTemplatePresentation(empty,{templateBuffer:await fs.readFile('public/presentation/facility42-template.pptx')});assert.equal(sparse.slideCount,18)
await fs.writeFile('../build/IML-generated-test.pptx',result.bytes)
console.log('PASS: missing vs zero, facility/date/pack filtering, partial coverage, validation, template export, native charts and Office table ordering')
