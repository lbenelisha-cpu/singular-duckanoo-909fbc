const clean = v => String(v ?? '').trim().replace(/["'׳״]/g,'').replace(/\s+/g,' ')
const number = (value, label) => {
  if (value === '' || value == null || typeof value === 'boolean') throw new Error(`${label}: חסר ערך מספרי מחושב. יש לחשב ולשמור את הקובץ ב-Excel.`)
  const n = Number(String(value).replace(/,/g,''))
  if (!Number.isFinite(n)) throw new Error(`${label}: ערך לא תקין`)
  return n
}
const same = (a,b) => Math.abs(a-b) < 0.02
function totalColumn(matrix, name) {
  const header = matrix.find(row => clean(row[0]) === 'תאריך' && row.some(v=>clean(v)==='סהכ'))
  if (!header) throw new Error(`${name}: לא נמצאה עמודת סה״כ`)
  return header.findIndex(v=>clean(v)==='סהכ')
}
function sideValue(matrix, label, name) {
  for (const row of matrix) for (let i=1;i<row.length;i++) if(clean(row[i])===label) {
    const value = row.slice(i+1).find(v=>v!=='' && v!=null)
    return number(value,`${name}: ${label}`)
  }
  return null
}
export function parseContractorWorkbook(wb, fileName, XLSX) {
  const recognized = wb.SheetNames.some(name=>clean(name).startsWith('מינימום'))
  if (!recognized) return null // Keep the existing simple-table importer available.
  const yearMatch = String(fileName).match(/20\d{2}/)
  if (!yearMatch) throw new Error('יש לציין את שנת החשבון בשם הקובץ')
  const year = Number(yearMatch[0]), months = new Map()
  for (const name of wb.SheetNames) {
    const label = clean(name)
    const kind = label.startsWith('מינימום') ? 'minimum' : /רגי\s+מלא/.test(label) ? 'regie' : null
    const match = name.match(/(\d{1,2})[.\/-](20\d{2})/)
    if (!kind || !match) continue
    const month=Number(match[1]), sheetYear=Number(match[2])
    // The supplied 2026 workbook includes stale September-December 2025 tabs.
    if (sheetYear!==year || month<1 || month>12 || (year===2026 && month>8)) continue
    const key=`${year}-${String(month).padStart(2,'0')}`
    if (!months.has(key)) months.set(key,{})
    if (months.get(key)[kind]) throw new Error(`${key}: יותר מגיליון ${kind} אחד`)
    months.get(key)[kind]={name,matrix:XLSX.utils.sheet_to_json(wb.Sheets[name],{header:1,defval:'',raw:true})}
  }
  const result=[]
  for (const [month,pair] of [...months].sort(([a],[b])=>a.localeCompare(b))) {
    if (!pair.minimum || !pair.regie) throw new Error(`${month}: חסר גיליון מינימום או רג׳י מלא`)
    const min=pair.minimum, reg=pair.regie
    const costCol=totalColumn(min.matrix,min.name), qtyCol=totalColumn(reg.matrix,reg.name)
    const costs=min.matrix.filter(r=>clean(r[0])==='סכום כולל')
    if(costs.length!==1)throw new Error(`${min.name}: סיכום תשלום חסר או לא חד משמעי`)
    const cost=number(costs[0][costCol],`${min.name}: סכום כולל`)
    const topCost=min.matrix.find(r=>clean(r[0])==='סכום כולל יומי')
    const components = ["תשלום עבור תפוקות","תשלום מינימום","תשלום עבור רגי"].map(label => min.matrix.find(r=>clean(r[0])===label))
    if(components.some(r=>!r) || !same(components.reduce((sum,r)=>sum+number(r[costCol],min.name),0),cost)) throw new Error(`${min.name}: הסכום הסופי אינו תואם לרכיביו`)
    const lines={}, shifts=[0,0,0];let shift=-1,items=0
    for(const row of reg.matrix) {
      const label=clean(row[0]);if(label==='סכום כולל יומי')break
      const shiftMatch=label.match(/^משמרת ([123])$/);if(shiftMatch)shift=Number(shiftMatch[1])-1
      if(/^(1|5|10|20)\s*L$/i.test(label) && shift>=0){const n=number(row[qtyCol],`${reg.name}: ${label}`);lines[label]=(lines[label]||0)+n;shifts[shift]+=n;items++}
    }
    if(items!==12)throw new Error(`${reg.name}: לא זוהו 12 שורות אריזה בשלוש משמרות`)
    const detailQty=shifts.reduce((a,b)=>a+b,0)
    const summaryQty=sideValue(reg.matrix,'סהכ תוצרת ארוזה',reg.name)
    const packaged=summaryQty ?? detailQty
    if(packaged<=0)throw new Error(`${reg.name}: כמות האריזה אינה חיובית; לא ניתן לחשב עלות לליטר`)
    const shiftCost=min.matrix.filter(r=>clean(r[0])==='סכום למשמרת').slice(0,3).map(r=>number(r[costCol],min.name))
    const matchingShiftCosts = shiftCost.length===3 && same(shiftCost.reduce((a,b)=>a+b,0),cost)
    // Avoid presenting shift quantities that contradict the authoritative packaging total.
    result.push({month,facility:'42',packaged,cost,cost_per_unit:cost/packaged,lines:same(packaged,detailQty)?lines:{},shift_qty:same(packaged,detailQty)?shifts:[],shift_cost:matchingShiftCosts?shiftCost:[],source_label:`${fileName} | תשלום: ${min.name} | ליטרים: ${reg.name}${summaryQty!=null?' / סה״כ תוצרת ארוזה':''}`})
  }
  if(!result.length)throw new Error('לא נמצאו זוגות גיליונות חודשיים מתאימים לשנת הקובץ')
  return result
}
