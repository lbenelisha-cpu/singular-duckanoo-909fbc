import pptxGenBundleUrl from '../vendor/pptxgen.bundle.js?url'

const PPTX_CDN = 'https://cdn.jsdelivr.net/gh/gitbrent/pptxgenjs@4.0.1/dist/pptxgen.bundle.js'

async function getPptxGenJS() {
  if (window.PptxGenJS) return window.PptxGenJS
  await new Promise((resolve, reject) => {
    const existing = document.querySelector('script[data-iml-pptxgen]')
    if (existing) {
      existing.addEventListener('load', resolve, { once: true })
      existing.addEventListener('error', () => reject(new Error('לא ניתן לטעון את מנוע PowerPoint')), { once: true })
      return
    }
    const script = document.createElement('script')
    script.src = pptxGenBundleUrl || PPTX_CDN
    script.async = true
    script.dataset.imlPptxgen = '1'
    script.onload = resolve
    script.onerror = () => reject(new Error('לא ניתן לטעון את מנוע PowerPoint'))
    document.head.appendChild(script)
  })
  if (!window.PptxGenJS) throw new Error('מנוע PowerPoint לא נטען')
  return window.PptxGenJS
}

const NAVY = '0B2038'
const TEAL = '0F766E'
const AQUA = '20B8B2'
const SKY = '37A8E8'
const PURPLE = '8B5CF6'
const ORANGE = 'F59E0B'
const RED = 'D9534F'
const GREEN = '19A974'
const LIGHT = 'F4F8FB'
const MUTED = '6B7C8F'
const WHITE = 'FFFFFF'
const DARK = '14283E'

const num = value => Number(value || 0)
const fmt = value => Math.round(num(value)).toLocaleString('he-IL')
const money = value => `₪${fmt(value)}`
const pct = value => Number.isFinite(Number(value)) ? `${Number(value).toFixed(1)}%` : '—'

const rtl = (extra = {}) => ({ fontFace: 'Arial', lang: 'he-IL', rtlMode: true, align: 'right', valign: 'mid', color: DARK, ...extra })
const centered = (extra = {}) => ({ fontFace: 'Arial', lang: 'he-IL', rtlMode: true, align: 'center', valign: 'mid', color: DARK, ...extra })

function addHeader(slide, title, subtitle = '') {
  slide.background = { color: LIGHT }
  slide.addShape('rect', { x: 0, y: 0, w: 13.333, h: 0.16, fill: { color: AQUA }, line: { color: AQUA } })
  slide.addText(title, { x: 5.15, y: 0.42, w: 7.55, h: 0.68, fontSize: 35, bold: true, color: NAVY, fontFace: 'Arial', rtlMode: true, lang: 'he-IL', align: 'right', margin: 0 })
  if (subtitle) slide.addText(subtitle, { x: 5.0, y: 1.12, w: 7.7, h: 0.36, fontSize: 16, color: MUTED, fontFace: 'Arial', rtlMode: true, lang: 'he-IL', align: 'right', margin: 0 })
  slide.addText('IML CONTROL', { x: 0.65, y: 0.52, w: 2.2, h: 0.32, fontSize: 15, bold: true, color: TEAL, fontFace: 'Arial', margin: 0 })
}

function addFooter(slide, text = 'IML CONTROL · תקציר מנהלים') {
  slide.addText(text, { x: 0.55, y: 7.08, w: 6.0, h: 0.18, fontSize: 9, color: MUTED, fontFace: 'Arial', align: 'left', margin: 0 })
}

function addKpi(slide, { x, y, w, h = 1.05, label, value, note = '', accent = TEAL }) {
  slide.addShape('roundRect', { x, y, w, h, rectRadius: 0.08, fill: { color: WHITE, transparency: 3 }, line: { color: 'DCE6EE', transparency: 35, width: 1 } })
  slide.addShape('rect', { x: x + w - 0.07, y: y + 0.08, w: 0.045, h: h - 0.16, fill: { color: accent }, line: { color: accent } })
  slide.addText(label, { x: x + 0.12, y: y + 0.13, w: w - 0.28, h: 0.24, fontSize: 9, color: MUTED, fontFace: 'Arial', rtlMode: true, lang: 'he-IL', align: 'right', margin: 0 })
  slide.addText(String(value), { x: x + 0.12, y: y + 0.38, w: w - 0.28, h: 0.38, fontSize: 20, bold: true, color: DARK, fontFace: 'Arial', rtlMode: true, lang: 'he-IL', align: 'right', margin: 0 })
  if (note) slide.addText(note, { x: x + 0.12, y: y + 0.79, w: w - 0.28, h: 0.17, fontSize: 7.5, color: MUTED, fontFace: 'Arial', rtlMode: true, lang: 'he-IL', align: 'right', margin: 0 })
}

function addSectionTitle(slide, title, x = 0.65, y = 1.55, w = 12.0) {
  slide.addText(title, { x, y, w, h: 0.42, fontSize: 24, bold: true, color: NAVY, fontFace: 'Arial', rtlMode: true, lang: 'he-IL', align: 'right', margin: 0 })
}

function addMonthlyBars(slide, rows, { x = 0.8, y = 2.2, w = 11.8, h = 3.8 } = {}) {
  const data = rows.slice(-12)
  if (!data.length) {
    slide.addText('אין נתונים חודשיים לתקופה שנבחרה', { x, y: y + 1.2, w, h: 0.4, fontSize: 18, color: 'C9D6E1', ...centered() })
    return
  }
  const max = Math.max(1, ...data.flatMap(r => [num(r.plan), num(r.actual)]))
  const groupW = w / data.length
  const barW = Math.min(0.23, groupW * 0.28)
  const chartTop = y + 0.25
  const chartH = h - 0.75
  slide.addShape('line', { x, y: chartTop + chartH, w, h: 0, line: { color: '70869A', transparency: 35, width: 1 } })
  data.forEach((r, i) => {
    const baseX = x + i * groupW + groupW * 0.5
    const planH = Math.max(0.04, num(r.plan) / max * chartH)
    const actH = Math.max(0.04, num(r.actual) / max * chartH)
    slide.addShape('rect', { x: baseX - barW - 0.03, y: chartTop + chartH - planH, w: barW, h: planH, fill: { color: '9FB3C4' }, line: { color: '9FB3C4' } })
    slide.addShape('rect', { x: baseX + 0.03, y: chartTop + chartH - actH, w: barW, h: actH, fill: { color: AQUA }, line: { color: AQUA } })
    slide.addText(String(r.label || r.key || '').replace(/^.*\s/, ''), { x: x + i * groupW, y: chartTop + chartH + 0.08, w: groupW, h: 0.18, fontSize: 7.5, color: 'C6D3DD', align: 'center', margin: 0, fontFace: 'Arial' })
    if (num(r.plan)) slide.addText(pct(r.pct), { x: x + i * groupW, y: chartTop + chartH - Math.max(planH, actH) - 0.24, w: groupW, h: 0.17, fontSize: 9, bold: true, color: NAVY, align: 'center', margin: 0, fontFace: 'Arial' })
  })
  slide.addText('תכנון', { x: x + 8.8, y: y + h - 0.18, w: 0.8, h: 0.2, fontSize: 8, color: 'C6D3DD', ...rtl() })
  slide.addShape('rect', { x: x + 9.55, y: y + h - 0.14, w: 0.12, h: 0.12, fill: { color: '9FB3C4' }, line: { color: '9FB3C4' } })
  slide.addText('ביצוע', { x: x + 10.2, y: y + h - 0.18, w: 0.8, h: 0.2, fontSize: 8, color: 'C6D3DD', ...rtl() })
  slide.addShape('rect', { x: x + 10.95, y: y + h - 0.14, w: 0.12, h: 0.12, fill: { color: AQUA }, line: { color: AQUA } })
}

function addAnnualCards(slide, rows, y = 2.1) {
  const colors = [SKY, PURPLE, AQUA]
  rows.forEach((r, i) => {
    const x = 0.75 + i * 4.15
    slide.addShape('roundRect', { x, y, w: 3.75, h: 2.2, fill: { color: WHITE, transparency: 2 }, line: { color: 'DCE6EE', transparency: 40 } })
    slide.addText(String(r.year), { x: x + 0.25, y: y + 0.2, w: 3.25, h: 0.35, fontSize: 20, bold: true, color: colors[i] || AQUA, align: 'center', margin: 0, fontFace: 'Arial' })
    slide.addText(fmt(r.actual), { x: x + 0.25, y: y + 0.65, w: 3.25, h: 0.42, fontSize: 25, bold: true, color: DARK, align: 'center', margin: 0, fontFace: 'Arial' })
    slide.addText('ביצוע בפועל', { x: x + 0.25, y: y + 1.08, w: 3.25, h: 0.24, fontSize: 9, color: MUTED, ...centered() })
    slide.addText(`תכנון ${fmt(r.plan)} · ${r.plan ? pct(r.pct) : 'ללא תכנון'}`, { x: x + 0.25, y: y + 1.42, w: 3.25, h: 0.28, fontSize: 10, bold: true, color: r.pct >= 100 ? GREEN : r.pct >= 90 ? ORANGE : RED, ...centered() })
    if (num(r.costPerUnit)) slide.addText(`עלות/יחידת תפוקה ₪${num(r.costPerUnit).toFixed(3)}`, { x: x + 0.25, y: y + 1.78, w: 3.25, h: 0.22, fontSize: 9, color: MUTED, ...centered() })
  })
}

function addRankList(slide, title, rows, { x, y, w, valueKey = 'qty', labelFn }) {
  slide.addShape('roundRect', { x, y, w, h: 3.9, fill: { color: WHITE, transparency: 2 }, line: { color: 'DCE6EE', transparency: 45 } })
  slide.addText(title, { x: x + 0.25, y: y + 0.22, w: w - 0.5, h: 0.34, fontSize: 16, bold: true, color: DARK, ...rtl() })
  rows.slice(0, 6).forEach((r, i) => {
    const yy = y + 0.72 + i * 0.48
    slide.addText(labelFn(r, i), { x: x + 1.35, y: yy, w: w - 1.6, h: 0.24, fontSize: 9.5, bold: true, color: DARK, ...rtl() })
    slide.addText(fmt(r[valueKey]), { x: x + 0.15, y: yy, w: 1.1, h: 0.24, fontSize: 10.5, bold: true, color: TEAL, align: 'left', margin: 0, fontFace: 'Arial' })
    if (i < Math.min(rows.length, 6) - 1) slide.addShape('line', { x: x + 0.2, y: yy + 0.34, w: w - 0.4, h: 0, line: { color: 'E6EDF2', width: 0.6 } })
  })
}

export async function exportManagementPresentation({ summary, from, to }) {
  const PptxGenJS = await getPptxGenJS()
  const pptx = new PptxGenJS()
  pptx.author = 'IML CONTROL'; pptx.company = 'ADAMA'; pptx.subject = 'Facility 42 Management Summary'
  pptx.title = 'סיכום מתקן 42'; pptx.lang = 'he-IL'
  pptx.theme = { headFontFace: 'Arial', bodyFontFace: 'Arial', lang: 'he-IL' }
  pptx.defineLayout({ name: 'IML_WIDE', width: 13.333, height: 7.5 }); pptx.layout = 'IML_WIDE'

  const facilities = summary.logicalFacilities?.length ? summary.logicalFacilities.join(', ') : '42'
  const period = `${from || 'תחילת הנתונים'} עד ${to || 'סוף הנתונים'}`
  const monthly = summary.monthlyTrend || []
  const total = num(summary.total), days = num(summary.days)
  const plan = num(summary.fmsPlan), actual = num(summary.fmsActual)
  const planPct = plan ? actual / plan * 100 : 0
  const lineGroups = ['42-P-02','42-P-03','42-P-04']
  const lineLabels = {'42-P-02':'1 ליטר','42-P-03':'5 ליטר','42-P-04':'10/20 ליטר'}
  const groupTotal = g => monthly.reduce((s,m)=>s+num(m.groups?.[g]?.actual),0)
  const groupPlan = g => monthly.reduce((s,m)=>s+num(m.groups?.[g]?.plan),0)
  const activityDays = Math.max(1, days)
  const dailyTarget = plan / activityDays
  const avgDaily = num(summary.avgDaily)
  const addTitleFooter = (slide, title, subtitle='') => { addHeader(slide,title,subtitle); addFooter(slide,`IML CONTROL · מתקן 42 · ${period}`) }
  const addNoSource = (slide, text, y=3.0) => slide.addText(text,{x:1.0,y,w:11.3,h:0.75,fontSize:18,bold:true,color:MUTED,...centered()})
  const addTable = (slide, headers, rows, x=0.75, y=2.0, widths=null) => {
    const ws=widths||headers.map(()=>11.8/headers.length), rh=0.48
    let xx=x; headers.forEach((h,i)=>{slide.addText(h,{x:xx,y,w:ws[i],h:rh,fontSize:11,bold:true,color:WHITE,fill:{color:TEAL},line:{color:'DCE6EE'},...centered()});xx+=ws[i]})
    rows.forEach((row,r)=>{xx=x;row.forEach((v,i)=>{slide.addText(String(v),{x:xx,y:y+rh*(r+1),w:ws[i],h:rh,fontSize:10,fill:{color:r%2?'F7FAFC':'FFFFFF'},line:{color:'DCE6EE'},...centered()});xx+=ws[i]})})
  }
  const addLineChart = (slide, rows, series, x=0.9, y=2.15, w=11.4, h=3.6) => {
    if(!rows.length){addNoSource(slide,'אין נתונים חודשיים לתקופה שנבחרה');return}
    const vals=rows.flatMap(r=>series.map(s=>num(s.value(r)))); const max=Math.max(1,...vals)
    slide.addShape('line',{x,y:y+h,w,h:0,line:{color:'A9B7C4',width:1}})
    series.forEach((s,si)=>{let prev=null;rows.forEach((r,i)=>{const px=x+(rows.length===1?w/2:i*w/(rows.length-1));const py=y+h-(num(s.value(r))/max*h);slide.addShape('ellipse',{x:px-0.045,y:py-0.045,w:0.09,h:0.09,fill:{color:s.color},line:{color:s.color}});if(prev)slide.addShape('line',{x:prev.x,y:prev.y,w:px-prev.x,h:py-prev.y,line:{color:s.color,width:2}});prev={x:px,y:py};slide.addText(String(r.label||r.key),{x:px-0.45,y:y+h+0.1,w:0.9,h:0.2,fontSize:8,color:MUTED,align:'center',margin:0})})})
    series.forEach((s,i)=>{slide.addShape('rect',{x:9.1+i*1.55,y:6.15,w:0.12,h:0.12,fill:{color:s.color},line:{color:s.color}});slide.addText(s.label,{x:9.25+i*1.55,y:6.09,w:1.2,h:0.2,fontSize:9,color:MUTED,...rtl()})})
  }
  const tryAddLogo = async slide => {
    try { const res=await fetch('/icons/adama-mark-rotate.gif'); if(!res.ok) throw new Error('logo'); const blob=await res.blob(); const data=await new Promise((resolve,reject)=>{const fr=new FileReader();fr.onload=()=>resolve(fr.result);fr.onerror=reject;fr.readAsDataURL(blob)}); slide.addImage({data,x:5.25,y:0.38,w:2.85,h:2.85}) } catch { slide.addText('ADAMA',{x:5.2,y:0.9,w:2.9,h:0.7,fontSize:38,bold:true,color:TEAL,align:'center',margin:0}) }
  }

  // 1 — cover
  let slide=pptx.addSlide(); slide.background={color:'F7FAFC'}
  slide.addShape('rect',{x:0,y:0,w:13.333,h:0.14,fill:{color:AQUA},line:{color:AQUA}})
  await tryAddLogo(slide)
  slide.addText('סיכום מתקן 42',{x:2.0,y:3.0,w:9.3,h:0.75,fontSize:42,bold:true,color:NAVY,...centered()})
  slide.addText(`התקופה: ${period}`,{x:2.0,y:3.72,w:9.3,h:0.38,fontSize:18,color:MUTED,...centered()})
  addKpi(slide,{x:0.7,y:4.65,w:2.8,label:'תפוקה בפועל',value:fmt(total),note:`${days} ימי פעילות`,accent:AQUA})
  addKpi(slide,{x:3.75,y:4.65,w:2.8,label:'FMS מול תכנון',value:plan?pct(planPct):'—',note:plan?`${fmt(actual)} / ${fmt(plan)}`:'אין תכנון',accent:GREEN})
  addKpi(slide,{x:6.8,y:4.65,w:2.8,label:'שנה מול שנה',value:summary.previousActual?`${summary.yoyPct>=0?'+':''}${pct(summary.yoyPct)}`:'—',note:`מול ${summary.currentYear-1}`,accent:ORANGE})
  addKpi(slide,{x:9.85,y:4.65,w:2.8,label:'עלות קבלן / ליטר',value:summary.contractorCostPerUnit?`₪${num(summary.contractorCostPerUnit).toFixed(3)}`:'—',note:summary.contractorCostPerUnit?`${summary.contractorMonths} חודשים`:'אין נתון',accent:PURPLE})
  addFooter(slide,`IML CONTROL · מתקן 42 · ${period}`)

  // 2 — agenda
  slide=pptx.addSlide(); addTitleFooter(slide,'על סדר היום',`סיכום מתקן 42 · ${period}`)
  const agenda=['תמונת מצב בטיחות','תפוקות אריזה בתקופה','עמידה ביעד היומי','השוואה ומגמות חודשיות','תמהיל אריזה לפי קווים','ניתוח משמרות ועלויות','מדדי איכות RFT / חריגות','תובנות והמלצות']
  agenda.forEach((t,i)=>{const col=i%2,row=Math.floor(i/2),x=0.8+col*6.1,y=1.6+row*1.15;slide.addText(String(i+1).padStart(2,'0'),{x,y,w:0.7,h:0.42,fontSize:18,bold:true,color:AQUA,align:'center',margin:0});slide.addText(t,{x:x+0.85,y,w:4.9,h:0.42,fontSize:18,bold:true,...rtl()})})

  // 3 — safety
  slide=pptx.addSlide(); addTitleFooter(slide,'תמונת מצב בטיחות – מתקן 42','הנתונים יוזנו אוטומטית לאחר חיבור מקור בטיחות לאפליקציה')
  addNoSource(slide,'מקור נתוני בטיחות עדיין לא מחובר ל-IML CONTROL. השקף מוכן לחיבור אוטומטי ללא הזנת מספרים ידנית.',2.5)

  // 4 — production summary
  slide=pptx.addSlide(); addTitleFooter(slide,'סיכום תפוקות אריזה',`כמויות לפי קווי אריזה · ${period}`)
  const prodRows=lineGroups.map(g=>[lineLabels[g],fmt(groupTotal(g)),total?pct(groupTotal(g)/total*100):'—'])
  prodRows.push(['סה״כ',fmt(total),'100%']); addTable(slide,['קו אריזה','ליטרים','%'],prodRows,2.0,2.0,[3.1,3.1,3.1])
  slide.addText(`ממוצע ליום פעילות: ${fmt(avgDaily)} ליטר · שיא יומי: ${fmt(summary.peakDaily)}`,{x:1.3,y:5.55,w:10.7,h:0.45,fontSize:17,bold:true,color:NAVY,...centered()})

  // 5 — FMS by line
  slide=pptx.addSlide(); addTitleFooter(slide,'תכנון FMS מול ביצוע',`לפי קווי האריזה · ${period}`)
  const fmsRows=lineGroups.map(g=>{const p=groupPlan(g),a=groupTotal(g);return[lineLabels[g],fmt(p),fmt(a),p?pct(a/p*100):'—']}); fmsRows.push(['סה״כ',fmt(plan),fmt(actual),plan?pct(planPct):'—']); addTable(slide,['קו','תכנון','ביצוע','%'],fmsRows,1.1,2.0,[2.6,2.6,2.6,2.6])

  // 6 — daily target overview
  slide=pptx.addSlide(); addTitleFooter(slide,'עמידה ביעד היומי',`היעד נגזר מתכנון FMS בתקופה שנבחרה`)
  addKpi(slide,{x:0.8,y:2.0,w:2.7,label:'יעד יומי נגזר',value:fmt(dailyTarget),note:plan?`${fmt(plan)} ÷ ${days} ימי פעילות`:'אין תכנון',accent:TEAL})
  addKpi(slide,{x:3.8,y:2.0,w:2.7,label:'ממוצע יומי',value:fmt(avgDaily),note:dailyTarget?pct(avgDaily/dailyTarget*100):'',accent:SKY})
  addKpi(slide,{x:6.8,y:2.0,w:2.7,label:'שיא יומי',value:fmt(summary.peakDaily),note:'מקובץ הכמויות',accent:GREEN})
  addKpi(slide,{x:9.8,y:2.0,w:2.7,label:'עמידה מול FMS',value:plan?pct(planPct):'—',note:`ביצוע ${fmt(actual)}`,accent:ORANGE})
  addLineChart(slide,monthly,[{label:'ממוצע/ביצוע',color:AQUA,value:r=>r.actual},{label:'תכנון',color:'9FB3C4',value:r=>r.plan}],1.0,3.55,11.2,2.0)

  // 7 — monthly comparison table
  slide=pptx.addSlide(); addTitleFooter(slide,'השוואה לכל חודשי התקופה','תכנון, ביצוע, עמידה ועלות')
  const monthRows=monthly.slice(-10).map(r=>[r.label||r.key,fmt(r.actual),fmt(r.plan),r.plan?pct(r.pct):'—',r.costPerUnit?`₪${num(r.costPerUnit).toFixed(3)}`:'—']); addTable(slide,['חודש','תפוקה','תכנון FMS','% עמידה','עלות/ליטר'],monthRows,0.75,1.75,[2.0,2.35,2.35,2.0,2.35])

  // 8 — key trends
  slide=pptx.addSlide(); addTitleFooter(slide,'מגמות מרכזיות',`לפי התקופה שנבחרה`)
  addKpi(slide,{x:0.8,y:1.8,w:2.8,label:'חודש שיא',value:summary.peakMonth?.label||'—',note:summary.peakMonth?fmt(summary.peakMonth.actual):'',accent:GREEN})
  addKpi(slide,{x:3.85,y:1.8,w:2.8,label:'חודש חלש',value:summary.weakMonth?.label||'—',note:summary.weakMonth?fmt(summary.weakMonth.actual):'',accent:ORANGE})
  addKpi(slide,{x:6.9,y:1.8,w:2.8,label:'עמידה מיטבית FMS',value:summary.bestPlanMonth?.label||'—',note:summary.bestPlanMonth?.plan?pct(summary.bestPlanMonth.pct):'',accent:AQUA})
  addKpi(slide,{x:9.95,y:1.8,w:2.6,label:'עלות ממוצעת',value:summary.contractorCostPerUnit?`₪${num(summary.contractorCostPerUnit).toFixed(3)}`:'—',note:'עלות קבלן / תפוקה',accent:PURPLE})
  addLineChart(slide,monthly,[{label:'ביצוע',color:AQUA,value:r=>r.actual},{label:'תכנון',color:ORANGE,value:r=>r.plan}],1.0,3.45,11.2,2.15)

  // 9 — plan vs actual chart
  slide=pptx.addSlide(); addTitleFooter(slide,'תכנון FMS מול ביצוע – מגמה חודשית','תכנון חודשי לפי מקור FMS וביצוע מקובץ הכמויות')
  addMonthlyBars(slide,monthly,{x:0.75,y:1.7,w:11.85,h:4.7})

  // 10 — packaging mix
  slide=pptx.addSlide(); addTitleFooter(slide,'תמהיל אריזה לפי קווים',`ליטרים לפי גודל אריזה · ${period}`)
  const mixRows=lineGroups.map(g=>[lineLabels[g],fmt(groupTotal(g)),total?pct(groupTotal(g)/total*100):'—',fmt(groupPlan(g)),groupPlan(g)?pct(groupTotal(g)/groupPlan(g)*100):'—']); addTable(slide,['קו','ביצוע','% מהתפוקה','תכנון','% מול תכנון'],mixRows,0.8,2.0,[2.0,2.25,2.25,2.25,2.25])
  addRankList(slide,'מוצרים מובילים בתקופה',summary.topMaterials||[],{x:1.2,y:4.25,w:10.9,labelFn:(r,i)=>`#${i+1} · ${r.desc||r.material}`})

  // 11 — shifts
  slide=pptx.addSlide(); addTitleFooter(slide,'ניתוח משמרות – תפוקה ועלות','מוכן לחיבור אוטומטי לנתוני משמרות')
  const shiftSource=summary.latestContractorRecord?.shift_qty||[]
  if(shiftSource.length){const rows=shiftSource.map((v,i)=>[`משמרת ${i+1}`,fmt(v),summary.latestContractorRecord?.shift_cost?.[i]?money(summary.latestContractorRecord.shift_cost[i]):'—']);addTable(slide,['משמרת','תפוקה','עלות'],rows,2.0,2.1,[3.1,3.1,3.1])}else addNoSource(slide,'אין כרגע נתוני משמרות מובנים בתקופה שנבחרה. השקף יוזן אוטומטית לאחר חיבור מקור המשמרות.',2.7)

  // 12 — production cost
  slide=pptx.addSlide(); addTitleFooter(slide,'עלות ייצור לליטר',`עלות קבלן ÷ תפוקה מקובץ הכמויות`)
  addKpi(slide,{x:1.0,y:1.8,w:3.4,label:'סה״כ עלות קבלן',value:summary.contractorCost?money(summary.contractorCost):'—',note:`${summary.contractorMonths||0} חודשים`,accent:PURPLE})
  addKpi(slide,{x:4.95,y:1.8,w:3.4,label:'תפוקה בתקופה',value:fmt(total),note:'מקור: קובץ כמויות',accent:AQUA})
  addKpi(slide,{x:8.9,y:1.8,w:3.4,label:'עלות לליטר',value:summary.contractorCostPerUnit?`₪${num(summary.contractorCostPerUnit).toFixed(3)}`:'—',note:'מחושב אוטומטית',accent:TEAL})
  addLineChart(slide,monthly.filter(r=>r.costPerUnit),[{label:'₪ לליטר',color:PURPLE,value:r=>r.costPerUnit}],1.0,3.55,11.2,2.0)

  // 13 — RFT
  slide=pptx.addSlide(); addTitleFooter(slide,'מדד R.F.T – איכות מהפעם הראשונה','היעד והחישוב יוצגו רק ממקור איכות מאומת')
  if(summary.hasReliableRft){addKpi(slide,{x:4.2,y:2.1,w:4.9,h:1.5,label:'RFT בתקופה',value:pct(summary.rft),note:'מקור מאומת',accent:GREEN})}else{addNoSource(slide,`RFT אינו מחושב מנתוני חריגות בלבד. קיימים ${fmt(summary.qualityLots)} לוטים/רשומות איכות, אך נדרש מקור First Pass מלא.`,2.6)}

  // 14 — COPQ
  slide=pptx.addSlide(); addTitleFooter(slide,'עלות אי-איכות (COPQ) – מתקן 42','מוכן לחיבור אוטומטי לדוח COPQ')
  addNoSource(slide,'מקור COPQ עדיין לא מחובר לאפליקציה. לאחר חיבורו יוצגו כאן עלות חודשית, מצטבר, חלק יחסי ומגמה.',2.7)

  // 15 — data governance / deviations
  slide=pptx.addSlide(); addTitleFooter(slide,'איכות, חריגות וממשל נתונים','מקורות הנתונים המשמשים את המצגת')
  addKpi(slide,{x:0.8,y:1.9,w:3.6,label:'לוטים / רשומות איכות',value:fmt(summary.qualityLots),note:'Inspection Lots / החלטות / חריגות',accent:SKY})
  addKpi(slide,{x:4.85,y:1.9,w:3.6,label:'לוטים חריגים',value:fmt(summary.qualityBadLots),note:'לפי הנתונים הזמינים',accent:RED})
  addKpi(slide,{x:8.9,y:1.9,w:3.6,label:'מקור תפוקה',value:'קובץ כמויות',note:'מקור יחיד לתפוקה וביצוע',accent:AQUA})
  slide.addText('כלל ניהולי: חשבון הקבלן משמש לעלות בלבד; כל תפוקה, ביצוע, תמהיל ומכנה לעלות נלקחים מקובץ הכמויות.',{x:1.1,y:4.15,w:11.1,h:0.75,fontSize:18,bold:true,color:NAVY,...centered()})

  // 16 — insights
  slide=pptx.addSlide(); addTitleFooter(slide,'תובנות, חריגים והמלצות','סיכום אוטומטי מבוסס נתוני IML CONTROL')
  const insights=summary.insights||[]; if(!insights.length)addNoSource(slide,'אין תובנות אוטומטיות זמינות לטווח שנבחר.',2.8)
  insights.slice(0,6).forEach((ins,i)=>{const col=i%2,row=Math.floor(i/2),x=0.75+col*6.0,y=1.65+row*1.55,color=ins.state==='good'?GREEN:ins.state==='risk'?RED:ORANGE;slide.addShape('roundRect',{x,y,w:5.65,h:1.2,fill:{color:WHITE},line:{color:'DDE6EC'}});slide.addShape('rect',{x:x+5.52,y:y+0.08,w:0.05,h:1.04,fill:{color},line:{color}});slide.addText(ins.title,{x:x+0.25,y:y+0.15,w:5.0,h:0.3,fontSize:14,bold:true,...rtl()});slide.addText(ins.text,{x:x+0.25,y:y+0.5,w:5.0,h:0.48,fontSize:10,color:MUTED,...rtl()})})

  // 17 — closing
  slide=pptx.addSlide(); slide.background={color:'F7FAFC'}; await tryAddLogo(slide)
  slide.addText('תודה רבה!',{x:2.0,y:3.0,w:9.3,h:0.7,fontSize:42,bold:true,color:NAVY,...centered()})
  slide.addText(`מתקן 42 · ${period}`,{x:2.0,y:3.8,w:9.3,h:0.4,fontSize:20,color:TEAL,...centered()})
  slide.addText('IML CONTROL · נתונים אוטומטיים לקבלת החלטות',{x:2.0,y:5.2,w:9.3,h:0.35,fontSize:15,color:MUTED,...centered()})

  const safePeriod=`${from||'start'}_${to||'end'}`.replace(/[^0-9A-Za-z_-]/g,'_')
  await pptx.writeFile({fileName:`IML_Facility42_Summary_${safePeriod}.pptx`})
}

