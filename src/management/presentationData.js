// Null means missing; zero is a reported value. Never derive RFT from deviation-only data.
export const SIZES = ['1','5','10','20']
export const MONTH_NAMES = ['ינואר','פברואר','מרץ','אפריל','מאי','יוני','יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר']
export const finite = v => v === '' || v == null || typeof v === 'boolean' || !Number.isFinite(Number(v)) ? null : Number(v)
export const sumComplete = values => values.length && values.every(v=>v!==null) ? values.reduce((a,b)=>a+b,0) : null
export const sumKnown = values => values.filter(v=>v!==null).reduce((a,b)=>a+b,0)
export const divide = (a,b) => a!==null && b!==null && b>0 ? a/b : null
export function dateKey(value) { if(typeof value==='string' && /^\d{4}-\d{2}-\d{2}/.test(value))return value.slice(0,10); const d=new Date(value);return value!=null&&!Number.isNaN(+d)?`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`:'' }
export function monthsBetween(from,to) { if(!/^\d{4}-\d{2}-\d{2}$/.test(from||'')||!/^\d{4}-\d{2}-\d{2}$/.test(to||'')||from>to)throw new Error('יש לבחור תאריכי התחלה וסיום תקינים'); if(from.slice(0,4)!==to.slice(0,4))throw new Error('המצגת משווה חודשים בתוך שנה אחת. יש לבחור תקופה באותה שנה.'); const out=[];for(let m=Number(from.slice(5,7));m<=Number(to.slice(5,7));m++)out.push(`${to.slice(0,4)}-${String(m).padStart(2,'0')}`);return out }
export function packagingSize(row) {
  const explicit=String(row.packSize??'').replace(/\s/g,'');if(SIZES.includes(explicit))return explicit
  const text=`${row.desc||''} ${row.routingDescription||''} ${row.prodLineTool||''} ${row.prodLine||''} ${row.routingGroup||''}`.toUpperCase()
  const matches=[...text.matchAll(/(?:^|[^\d])(20|10|5|1)\s*(?:LT|LITER|LIT|L\b|ליטר)/g)]
  if(matches.length)return matches.at(-1)[1]
  if(/42-P-02|LQ-P-1(?:\D|$)/.test(text))return '1'
  if(/42-P-03|LQ-P-5(?:\D|$)/.test(text))return '5'
  // Combined 10/20 resource is ambiguous without a pack description.
  return null
}
export function facility42(row){return ['42','1542'].includes(String(row.facility||'')) || /^42-P-0[234]$/.test(String(row.routingGroup||''))}
const groupValue=(groups,id,field) => {const value=groups?.[id]?.[field];return finite(value)}
export function buildPresentationData({rows=[],history={},supplements={},from,to,isFacility42=facility42,safety={}}) {
  validateSupplements(supplements)
  const selectedKeys=monthsBetween(from,to),year=to.slice(0,4),ytdKeys=monthsBetween(`${year}-01-01`,to)
  const scoped=rows.filter(r=>isFacility42(r)&&dateKey(r.productionDay||r.date).startsWith(year))
  const periods=supplements.periods||{},annual=supplements.annualByMonth?.[to.slice(0,7)]||(supplements.annual?.[year]?.through===to.slice(0,7)?supplements.annual[year]:{})
  const byKey=new Map()
  for(const key of ytdKeys){
    const manual=supplements.months?.[key]||{},historyPlan=history.planActual?.[key]?.['42'],costRec=history.contractor42?.[key]
    const start=key===from.slice(0,7)?from:`${key}-01`,end=key===to.slice(0,7)?to:`${key}-31`
    const data=scoped.filter(r=>dateKey(r.productionDay||r.date).slice(0,7)===key&&dateKey(r.productionDay||r.date)>=start&&dateKey(r.productionDay||r.date)<=end)
    const hasRows=data.length>0, map=new Map(), packs=Object.fromEntries(SIZES.map(s=>[s,0])), shifts=[0,0,0];let unmapped=0,unknownShift=0
    for(const row of data){const qty=finite(row.qty);if(qty===null)continue;const day=dateKey(row.productionDay||row.date);map.set(day,(map.get(day)||0)+qty);const size=packagingSize(row);if(size)packs[size]+=qty;else unmapped+=qty;const shift=typeof row.shift==='object'?row.shift?.key:row.shift;const idx=['morning','evening','night'].indexOf(shift);if(idx>=0)shifts[idx]+=qty;else unknownShift+=qty}
    // Manual quantity is only a fallback for a month without loaded quantity rows.
    const manualPacks=SIZES.map(s=>finite(manual[`pack${s}`])),manualTotal=sumComplete(manualPacks)
    const total=hasRows?sumKnown(data.map(r=>finite(r.qty))):manualTotal
    const packValues=hasRows?(unmapped===0?SIZES.map(s=>packs[s]):SIZES.map(()=>null)):manualPacks
    const daily=[...map].filter(([,qty])=>qty>0).sort(([a],[b])=>a.localeCompare(b)).map(([date,qty])=>({date,qty}))
    const days=hasRows?daily.length:finite(manual.days),shiftValues=hasRows&&unknownShift===0?shifts:[0,1,2].map(i=>finite(manual[`shift${i+1}`]))
    const planGroups=['42-P-02','42-P-03','42-P-04'].map((id,i)=>finite(manual[['plan1','plan5','plan1020'][i]])??groupValue(historyPlan?.groups,id,'plan'))
    const plan=sumComplete(planGroups)??finite(manual.plan)??finite(historyPlan?.plan)
    const fmsGroups=['42-P-02','42-P-03','42-P-04'].map((id,i)=>finite(manual[['fms1','fms5','fms1020'][i]])??groupValue(historyPlan?.groups,id,'actual'))
    const fmsActual=sumComplete(fmsGroups)??finite(historyPlan?.actual)
    const cost=finite(manual.cost)??finite(costRec?.cost),rft=finite(manual.rft),copq=finite(manual.copq),siteCopq=finite(manual.siteCopq)
    const dailyTarget=divide(plan,days),met=dailyTarget!==null&&daily.length?daily.filter(d=>d.qty>=dailyTarget).length:null
    const partial=(start!==`${key}-01`)||(key===to.slice(0,7)&&Number(to.slice(8))!==new Date(Number(year),Number(key.slice(5)),0).getDate())
    const missing=[];if(total===null)missing.push('תפוקה');if(unmapped!==0)missing.push('מיפוי גדלי אריזה');if(plan===null)missing.push('תכנון FMS');if(cost===null)missing.push('עלות קבלן');if(rft===null)missing.push('RFT');if(copq===null)missing.push('COPQ');if(shiftValues.some(v=>v===null))missing.push('משמרות')
    byKey.set(key,{key,label:MONTH_NAMES[Number(key.slice(5))-1],total,packs:packValues,days,daily,shifts:shiftValues,plan,planGroups,fmsActual,fmsGroups,cost,costRate:partial?null:divide(cost,total),rft,copq,siteCopq,dailyTarget,met,pct:divide(total,plan),missing,partial,quantitySource:hasRows?'קובץ כמויות':manualTotal!==null?'הזנה ידנית מאושרת':'חסר',unmapped})
  }
  const selected=selectedKeys.map(k=>byKey.get(k)), ytd=ytdKeys.map(k=>byKey.get(k)), total=sumComplete(selected.map(m=>m.total)),plan=sumComplete(selected.map(m=>m.plan)),cost=sumComplete(selected.map(m=>m.cost)),days=sumComplete(selected.map(m=>m.days)),copq=sumComplete(selected.map(m=>m.copq))
  const period=periods[`${from}_${to}`]||{},annualRft=finite(annual.rft),annualCopq=finite(annual.copq)??sumComplete(ytd.map(m=>m.copq)),annualSiteCopq=finite(annual.siteCopq)??sumComplete(ytd.map(m=>m.siteCopq))
  const selectedRfts=selected.map(m=>m.rft),periodRft=finite(period.rft)??(selectedRfts.every(v=>v!==null)&&new Set(selectedRfts).size===1?selectedRfts[0]:null)
  const safetyFields=['fatal','irreversible','lostDays','medical','firstAid','nearMiss','unsafe'];const resolveSafety=scope=>Object.fromEntries(safetyFields.map(k=>[k,finite(period[scope]?.[k])??finite(safety[scope]?.[k])]))
  const periodSafety=resolveSafety('period'),annualSafety=Object.fromEntries(safetyFields.map(k=>[k,finite(annual.safety?.[k])??finite(safety.annual?.[k])]))
  // Existing safety summary defaults empty months to zero. Treat that as missing, not reported zero.
  if(!safety.months?.length&&!period.period)for(const k of safetyFields)periodSafety[k]=null
  if(!safety.months?.length&&!safety.annualManualEnabled&&!annual.safety)for(const k of safetyFields)annualSafety[k]=null
  if(safety.records){for(const k of safetyFields){periodSafety[k]=finite(period.period?.[k])??sumComplete(selectedKeys.map(month=>finite(safety.records[month]?.[k])));annualSafety[k]=finite(annual.safety?.[k])??(safety.annualManualEnabled?finite(safety.annual?.[k]):sumComplete(ytdKeys.map(month=>finite(safety.records[month]?.[k]))))}}
  return {from,to,year,selected,ytd,total,plan,cost,days,avgDaily:divide(total,days),costRate:selected.some(m=>m.partial)?null:divide(cost,total),pct:divide(total,plan),copq,annualRft,periodRft,annualCopq,annualSiteCopq,copqShare:divide(annualCopq,annualSiteCopq),periodSafety,annualSafety,period,project:supplements.project||{},missing:selected.flatMap(m=>m.missing.map(label=>`${m.label}: ${label}`)),isPartial:selected.some(m=>m.partial)}
}

export function validateSupplements(data) {
 const inspect=(obj,path='')=>{for(const [key,value] of Object.entries(obj||{})){if(value&&typeof value==='object'){inspect(value,`${path}${key}.`);continue}if(/^(plan|plan1|plan5|plan1020|fms1|fms5|fms1020|cost|rft|copq|siteCopq|pack1|pack5|pack10|pack20|days|shift1|shift2|shift3|fatal|irreversible|lostDays|medical|firstAid|nearMiss|unsafe)$/.test(key)&&value!==''&&value!=null){const n=finite(value);if(n===null||n<0||(key==='rft'&&n>100))throw new Error(`נתון לא תקין: ${path}${key}`)}}};inspect(data)
}
