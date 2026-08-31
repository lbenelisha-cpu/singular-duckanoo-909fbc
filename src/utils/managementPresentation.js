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
    script.src = PPTX_CDN
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
  slide.background = { color: NAVY }
  slide.addShape('rect', { x: 0, y: 0, w: 13.333, h: 0.16, fill: { color: AQUA }, line: { color: AQUA } })
  slide.addText(title, { x: 6.45, y: 0.45, w: 6.25, h: 0.55, fontSize: 24, bold: true, color: WHITE, fontFace: 'Arial', rtlMode: true, lang: 'he-IL', align: 'right', margin: 0 })
  if (subtitle) slide.addText(subtitle, { x: 6.0, y: 1.02, w: 6.7, h: 0.35, fontSize: 10.5, color: 'B8C7D4', fontFace: 'Arial', rtlMode: true, lang: 'he-IL', align: 'right', margin: 0 })
}

function addFooter(slide, text = 'IML CONTROL · תקציר מנהלים') {
  slide.addText(text, { x: 0.55, y: 7.08, w: 6.0, h: 0.18, fontSize: 8, color: 'A9BBC8', fontFace: 'Arial', align: 'left', margin: 0 })
}

function addKpi(slide, { x, y, w, h = 1.05, label, value, note = '', accent = TEAL }) {
  slide.addShape('roundRect', { x, y, w, h, rectRadius: 0.08, fill: { color: WHITE, transparency: 3 }, line: { color: 'DCE6EE', transparency: 35, width: 1 } })
  slide.addShape('rect', { x: x + w - 0.07, y: y + 0.08, w: 0.045, h: h - 0.16, fill: { color: accent }, line: { color: accent } })
  slide.addText(label, { x: x + 0.12, y: y + 0.13, w: w - 0.28, h: 0.24, fontSize: 9, color: MUTED, fontFace: 'Arial', rtlMode: true, lang: 'he-IL', align: 'right', margin: 0 })
  slide.addText(String(value), { x: x + 0.12, y: y + 0.38, w: w - 0.28, h: 0.38, fontSize: 20, bold: true, color: DARK, fontFace: 'Arial', rtlMode: true, lang: 'he-IL', align: 'right', margin: 0 })
  if (note) slide.addText(note, { x: x + 0.12, y: y + 0.79, w: w - 0.28, h: 0.17, fontSize: 7.5, color: MUTED, fontFace: 'Arial', rtlMode: true, lang: 'he-IL', align: 'right', margin: 0 })
}

function addSectionTitle(slide, title, x = 0.65, y = 1.55, w = 12.0) {
  slide.addText(title, { x, y, w, h: 0.36, fontSize: 17, bold: true, color: WHITE, fontFace: 'Arial', rtlMode: true, lang: 'he-IL', align: 'right', margin: 0 })
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
    if (num(r.plan)) slide.addText(pct(r.pct), { x: x + i * groupW, y: chartTop + chartH - Math.max(planH, actH) - 0.24, w: groupW, h: 0.17, fontSize: 7.5, bold: true, color: WHITE, align: 'center', margin: 0, fontFace: 'Arial' })
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
    if (num(r.costPerUnit)) slide.addText(`עלות/ליטר ₪${num(r.costPerUnit).toFixed(3)}`, { x: x + 0.25, y: y + 1.78, w: 3.25, h: 0.22, fontSize: 9, color: MUTED, ...centered() })
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
  pptx.layout = 'LAYOUT_WIDE'
  pptx.author = 'IML CONTROL'
  pptx.company = 'ADAMA'
  pptx.subject = 'Management Summary'
  pptx.title = 'תקציר מנהלים IML CONTROL'
  pptx.lang = 'he-IL'
  pptx.theme = {
    headFontFace: 'Arial', bodyFontFace: 'Arial', lang: 'he-IL'
  }
  pptx.defineLayout({ name: 'IML_WIDE', width: 13.333, height: 7.5 })
  pptx.layout = 'IML_WIDE'

  const facilities = summary.logicalFacilities?.length ? summary.logicalFacilities.join(', ') : 'כל המתקנים'
  const period = `${from || 'תחילת הנתונים'} עד ${to || 'סוף הנתונים'}`

  // 1 — cover
  let slide = pptx.addSlide()
  slide.background = { color: NAVY }
  slide.addShape('rect', { x: 0, y: 0, w: 13.333, h: 0.16, fill: { color: AQUA }, line: { color: AQUA } })
  slide.addText('IML CONTROL', { x: 0.75, y: 0.65, w: 3.0, h: 0.4, fontSize: 18, bold: true, color: WHITE, fontFace: 'Arial', margin: 0 })
  slide.addText('תקציר מנהלים', { x: 5.4, y: 1.65, w: 6.8, h: 0.85, fontSize: 36, bold: true, color: WHITE, fontFace: 'Arial', rtlMode: true, lang: 'he-IL', align: 'right', margin: 0 })
  slide.addText(`${period} · מתקנים ${facilities}`, { x: 5.0, y: 2.55, w: 7.2, h: 0.42, fontSize: 14, color: 'C5D2DD', fontFace: 'Arial', rtlMode: true, lang: 'he-IL', align: 'right', margin: 0 })
  addKpi(slide, { x: 0.85, y: 4.65, w: 2.75, label: 'תפוקה בפועל', value: fmt(summary.total), note: `${summary.days} ימי פעילות`, accent: AQUA })
  addKpi(slide, { x: 3.85, y: 4.65, w: 2.75, label: 'FMS מול תכנון', value: summary.fmsPlan ? pct(summary.fmsActual / summary.fmsPlan * 100) : '—', note: `${fmt(summary.fmsActual)} / ${fmt(summary.fmsPlan)}`, accent: GREEN })
  addKpi(slide, { x: 6.85, y: 4.65, w: 2.75, label: 'שנה מול שנה', value: summary.previousActual ? `${summary.yoyPct >= 0 ? '+' : ''}${pct(summary.yoyPct)}` : '—', note: `מול ${summary.currentYear - 1}`, accent: ORANGE })
  addKpi(slide, { x: 9.85, y: 4.65, w: 2.75, label: 'עלות קבלן / ליטר', value: summary.contractorCostPerUnit ? `₪${summary.contractorCostPerUnit.toFixed(3)}` : '—', note: summary.contractorCostPerUnit ? `${summary.contractorMonths} חודשים` : 'אין נתון בתקופה', accent: PURPLE })

  // 2 — executive snapshot
  slide = pptx.addSlide(); addHeader(slide, 'תמונת מצב ניהולית', `${period} · מתקנים ${facilities}`)
  addSectionTitle(slide, 'מדדי מפתח')
  addKpi(slide, { x: 0.75, y: 2.05, w: 2.85, label: 'ממוצע ליום פעילות', value: fmt(summary.avgDaily), note: `שיא יומי ${fmt(summary.peakDaily)}`, accent: SKY })
  addKpi(slide, { x: 3.85, y: 2.05, w: 2.85, label: 'תכנון FMS', value: fmt(summary.fmsPlan), note: `ביצוע ${fmt(summary.fmsActual)}`, accent: TEAL })
  addKpi(slide, { x: 6.95, y: 2.05, w: 2.85, label: 'פער מול FMS', value: summary.fmsPlan ? `${summary.fmsActual - summary.fmsPlan >= 0 ? '+' : ''}${fmt(summary.fmsActual - summary.fmsPlan)}` : '—', note: summary.fmsPlan ? pct(summary.fmsActual / summary.fmsPlan * 100) : '', accent: summary.fmsActual >= summary.fmsPlan ? GREEN : ORANGE })
  addKpi(slide, { x: 10.05, y: 2.05, w: 2.55, label: 'RFT', value: summary.hasReliableRft ? pct(summary.rft) : '—', note: summary.hasReliableRft ? 'מקור מאומת' : 'ממתין למקור איכות מאומת', accent: ORANGE })
  slide.addText('תובנות אוטומטיות', { x: 6.6, y: 3.65, w: 6.0, h: 0.35, fontSize: 16, bold: true, color: WHITE, ...rtl() })
  summary.insights.slice(0, 4).forEach((ins, i) => {
    const y = 4.12 + i * 0.62
    const color = ins.state === 'good' ? GREEN : ins.state === 'risk' ? RED : ORANGE
    slide.addShape('roundRect', { x: 6.55, y, w: 6.05, h: 0.5, fill: { color: 'F7F9FB' }, line: { color: 'E0E7ED' } })
    slide.addShape('rect', { x: 12.48, y: y + 0.05, w: 0.055, h: 0.4, fill: { color }, line: { color } })
    slide.addText(ins.title, { x: 9.35, y: y + 0.08, w: 3.0, h: 0.2, fontSize: 10, bold: true, color: DARK, ...rtl() })
    slide.addText(ins.text, { x: 6.75, y: y + 0.27, w: 5.6, h: 0.17, fontSize: 7.8, color: MUTED, ...rtl() })
  })
  addRankList(slide, 'מוצרים מובילים', summary.topMaterials || [], { x: 0.75, y: 3.65, w: 5.5, labelFn: (r, i) => `#${i + 1} · ${r.desc || r.material}` })
  addFooter(slide)

  // 3 — monthly plan vs actual
  slide = pptx.addSlide(); addHeader(slide, 'תכנון FMS מול ביצוע', 'מגמה חודשית לפי התקופה שנבחרה')
  addMonthlyBars(slide, summary.monthlyTrend || [], { x: 0.75, y: 1.75, w: 11.85, h: 4.55 })
  addFooter(slide)

  // 4 — annual trend
  slide = pptx.addSlide(); addHeader(slide, 'מגמה רב־שנתית 2024–2026', 'השוואת אותם חודשים בין שלוש השנים')
  addAnnualCards(slide, summary.annualRows || [], 2.0)
  if (summary.peakMonth || summary.weakMonth || summary.bestPlanMonth) {
    addKpi(slide, { x: 1.15, y: 5.05, w: 3.4, label: 'חודש שיא', value: summary.peakMonth?.label || '—', note: summary.peakMonth ? fmt(summary.peakMonth.actual) : '', accent: GREEN })
    addKpi(slide, { x: 4.95, y: 5.05, w: 3.4, label: 'חודש חלש', value: summary.weakMonth?.label || '—', note: summary.weakMonth ? fmt(summary.weakMonth.actual) : '', accent: ORANGE })
    addKpi(slide, { x: 8.75, y: 5.05, w: 3.4, label: 'עמידה מיטבית ב-FMS', value: summary.bestPlanMonth?.label || '—', note: summary.bestPlanMonth?.plan ? pct(summary.bestPlanMonth.pct) : '', accent: AQUA })
  }
  addFooter(slide)

  // 5 — facilities and products
  slide = pptx.addSlide(); addHeader(slide, 'תמהיל תפוקה', 'מתקנים ומוצרים מובילים בתקופה')
  addRankList(slide, 'תפוקה לפי מתקן ניהולי', summary.facilityRows || [], { x: 0.75, y: 1.75, w: 5.85, labelFn: r => `מתקן ${r.facility}` })
  addRankList(slide, 'מוצרים מובילים', summary.topMaterials || [], { x: 6.75, y: 1.75, w: 5.85, labelFn: (r, i) => `#${i + 1} · ${r.desc || r.material}` })
  addFooter(slide)

  // 6 — cost
  slide = pptx.addSlide(); addHeader(slide, 'עלויות ויעילות', facilities === '42' ? 'חשבונות קבלן מתקן 42' : 'עלות קבלן זמינה למתקן 42 בלבד')
  addKpi(slide, { x: 0.75, y: 1.8, w: 3.6, label: 'סה״כ תשלום לקבלן', value: summary.contractorCost ? money(summary.contractorCost) : '—', note: summary.contractorCost ? `${summary.contractorMonths} חודשים` : 'אין נתונים בתקופה', accent: TEAL })
  addKpi(slide, { x: 4.85, y: 1.8, w: 3.6, label: 'תוצרת בחשבונות', value: summary.contractorPackaged ? fmt(summary.contractorPackaged) : '—', note: 'ליטר / יחידות מדווחות', accent: SKY })
  addKpi(slide, { x: 8.95, y: 1.8, w: 3.6, label: 'עלות משוקללת', value: summary.contractorCostPerUnit ? `₪${summary.contractorCostPerUnit.toFixed(3)}` : '—', note: 'לליטר / יחידה', accent: PURPLE })
  if (summary.annualRows?.some(r => num(r.costPerUnit))) {
    slide.addText('עלות משוקללת לפי שנה', { x: 6.4, y: 3.25, w: 6.1, h: 0.35, fontSize: 16, bold: true, color: WHITE, ...rtl() })
    const vals = summary.annualRows.map(r => num(r.costPerUnit))
    const max = Math.max(0.01, ...vals)
    summary.annualRows.forEach((r, i) => {
      const y = 3.85 + i * 0.75
      slide.addText(String(r.year), { x: 1.1, y, w: 1.0, h: 0.28, fontSize: 12, bold: true, color: WHITE, align: 'left', margin: 0, fontFace: 'Arial' })
      slide.addShape('roundRect', { x: 2.2, y: y + 0.02, w: 8.2, h: 0.25, fill: { color: '29445C' }, line: { color: '29445C' } })
      if (num(r.costPerUnit)) slide.addShape('roundRect', { x: 2.2, y: y + 0.02, w: Math.max(0.12, 8.2 * num(r.costPerUnit) / max), h: 0.25, fill: { color: [SKY, PURPLE, AQUA][i] || AQUA }, line: { color: [SKY, PURPLE, AQUA][i] || AQUA } })
      slide.addText(num(r.costPerUnit) ? `₪${num(r.costPerUnit).toFixed(3)}` : '—', { x: 10.65, y: y - 0.02, w: 1.3, h: 0.3, fontSize: 12, bold: true, color: WHITE, align: 'left', margin: 0, fontFace: 'Arial' })
    })
  } else {
    slide.addText('אין נתוני עלות קבלן זמינים בתקופה שנבחרה', { x: 1.0, y: 4.15, w: 11.2, h: 0.5, fontSize: 20, color: 'C9D6E1', ...centered() })
  }
  addFooter(slide)

  // 7 — quality
  slide = pptx.addSlide(); addHeader(slide, 'איכות ומגמות', 'RFT, לוטים וחריגות בתקופה')
  addKpi(slide, { x: 0.85, y: 2.0, w: 3.6, label: 'RFT', value: summary.hasReliableRft ? pct(summary.rft) : '—', note: summary.hasReliableRft ? 'מקור מאומת' : 'נדרש מקור RFT/UD מלא', accent: ORANGE })
  addKpi(slide, { x: 4.85, y: 2.0, w: 3.6, label: 'לוטים / רשומות איכות', value: fmt(summary.qualityLots), note: 'Inspection Lots / החלטות / חריגות', accent: SKY })
  addKpi(slide, { x: 8.85, y: 2.0, w: 3.6, label: 'לוטים חריגים', value: fmt(summary.qualityBadLots), note: 'לפי נתוני החריגות הזמינים', accent: RED })
  slide.addShape('roundRect', { x: 1.0, y: 4.0, w: 11.3, h: 1.55, fill: { color: 'F6F8FA' }, line: { color: 'DFE7ED' } })
  slide.addText(summary.hasReliableRft ? 'RFT מחושב ממקור מאומת.' : 'RFT אינו מוצג עד לחיבור מקור מלא ואמין של כל Inspection Lots והחלטת השימוש/First Pass. המערכת אינה מייצרת אחוז איכות מנתוני חריגות בלבד.', { x: 1.35, y: 4.4, w: 10.6, h: 0.75, fontSize: 16, color: DARK, bold: !summary.hasReliableRft, ...centered() })
  addFooter(slide)

  // 8 — insights / close
  slide = pptx.addSlide(); addHeader(slide, 'תובנות והמלצות', 'סיכום אוטומטי מבוסס נתוני IML CONTROL')
  const insights = summary.insights || []
  insights.slice(0, 6).forEach((ins, i) => {
    const col = i % 2
    const row = Math.floor(i / 2)
    const x = col ? 6.7 : 0.75
    const y = 1.75 + row * 1.45
    const color = ins.state === 'good' ? GREEN : ins.state === 'risk' ? RED : ORANGE
    slide.addShape('roundRect', { x, y, w: 5.85, h: 1.1, fill: { color: WHITE }, line: { color: 'DDE6EC' } })
    slide.addShape('rect', { x: x + 5.72, y: y + 0.08, w: 0.055, h: 0.94, fill: { color }, line: { color } })
    slide.addText(ins.title, { x: x + 0.3, y: y + 0.16, w: 5.15, h: 0.28, fontSize: 13, bold: true, color: DARK, ...rtl() })
    slide.addText(ins.text, { x: x + 0.3, y: y + 0.48, w: 5.15, h: 0.42, fontSize: 9, color: MUTED, breakLine: false, ...rtl() })
  })
  slide.addText('המצגת הופקה אוטומטית מתוך תקציר המנהלים של IML CONTROL', { x: 1.0, y: 6.35, w: 11.3, h: 0.35, fontSize: 12, color: 'B8C7D4', ...centered() })
  addFooter(slide)

  const safePeriod = `${from || 'start'}_${to || 'end'}`.replace(/[^0-9A-Za-z_-]/g, '_')
  await pptx.writeFile({ fileName: `IML_Management_Summary_${safePeriod}.pptx` })
}
