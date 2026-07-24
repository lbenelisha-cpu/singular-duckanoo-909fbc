import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import {
  Upload, Database, Factory, FlaskConical, CalendarDays, Search, CheckCircle2,
  AlertTriangle, Clock3, X, BarChart3, Download, Trash2, Save, Target,
  Gauge, CalendarCheck, BellRing, TrendingUp, FileSpreadsheet
} from 'lucide-react'
import './styles.css'

const LEGACY_DAILY_TARGETS = {
  '1519': 80000, '1521': 60000, '1523': 40000, '1524': 6000,
  '1525': 130000, '1528': 80000, '1540': 18000, '1541': 210000,
  '1542': 60000, '1543': 60000,
}
const FACILITY_ALIASES = {
  '1519': ['1519', '19', '19-F-01', '19-F-02'],
  '1521': ['1521', '21'],
  '1523': ['1523', '43', '43-P-A', '43-P-B'],
  '1524': ['1524', '24'],
  '1525': ['1525', '25'],
  '1528': ['1528', '28'],
  '1540': ['1540', '40'],
  '1541': ['1541', '41'],
  '1542': ['1542', '42-P-01', 'T42A'],
  '1543': ['1543', '42-P-03', 'T42B'],
}
const PRIMARY_FACILITIES = ['1519', '1541', '1540', '1525', '1523', '1528', '1524', '1542', '1543']
const DEFAULT_FACILITIES = PRIMARY_FACILITIES
const STORAGE_KEY = 'iml-control-center-sprint7'

const normalize = (v) => String(v ?? '').trim()
const normKey = (v) => normalize(v).toLowerCase().replace(/[\s_\-./()]+/g, '')
const num = (v) => {
  const n = Number(String(v ?? '').replace(/,/g, '').replace(/\s/g, ''))
  return Number.isFinite(n) ? n : 0
}
const excelDate = (v) => {
  if (v === '' || v === null || v === undefined) return null
  if (v instanceof Date && !Number.isNaN(v.getTime())) return v.getFullYear() <= 1900 ? null : new Date(v)
  if (typeof v === 'number') {
    if (v < 1) return null
    const d = XLSX.SSF.parse_date_code(v)
    return d ? new Date(d.y, d.m - 1, d.d, d.H || 0, d.M || 0, d.S || 0) : null
  }
  const text = normalize(v)
  const match = text.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/)
  if (match) {
    let year = Number(match[3]); if (year < 100) year += 2000
    return new Date(year, Number(match[2]) - 1, Number(match[1]), Number(match[4] || 0), Number(match[5] || 0), Number(match[6] || 0))
  }
  const d = new Date(text)
  return Number.isNaN(d.getTime()) || d.getFullYear() <= 1900 ? null : d
}
const excelTime = (v) => {
  if (v === '' || v === null || v === undefined) return null
  if (v instanceof Date && !Number.isNaN(v.getTime())) return { h: v.getHours(), m: v.getMinutes(), s: v.getSeconds() }
  if (typeof v === 'number') {
    const seconds = Math.round((v % 1) * 86400) % 86400
    return { h: Math.floor(seconds / 3600), m: Math.floor((seconds % 3600) / 60), s: seconds % 60 }
  }
  const match = normalize(v).match(/(\d{1,2}):(\d{2})(?::(\d{2}))?/)
  return match ? { h: Number(match[1]), m: Number(match[2]), s: Number(match[3] || 0) } : null
}
const combineExcelDateTime = (dateValue, timeValue, fallbackValue = '') => {
  const date = excelDate(dateValue) || excelDate(fallbackValue)
  if (!date) return null
  const time = excelTime(timeValue)
  const result = new Date(date)
  if (time) result.setHours(time.h, time.m, time.s, 0)
  return result
}
const iso = (d) => {
  if (!d) return ''
  const value = new Date(d)
  if (Number.isNaN(value.getTime()) || value.getFullYear() <= 1900) return ''
  return `${value.getFullYear()}-${String(value.getMonth() + 1).padStart(2, '0')}-${String(value.getDate()).padStart(2, '0')}`
}
const monthKey = (d) => iso(d).slice(0, 7)
const fmt = (n) => Math.round(n || 0).toLocaleString('he-IL')
const pctFmt = (n) => `${Math.round(n || 0)}%`
const getField = (row, names) => {
  const map = new Map(Object.keys(row || {}).map(k => [normKey(k), row[k]]))
  for (const name of names) {
    const value = map.get(normKey(name))
    if (value !== undefined) return value
  }
  return ''
}
const canonicalFacility = (value) => {
  const clean = normalize(value).toUpperCase()
  for (const [id, aliases] of Object.entries(FACILITY_ALIASES)) {
    if (aliases.some(alias => normalize(alias).toUpperCase() === clean)) return id
  }
  const digits = clean.match(/15\d{2}/)?.[0]
  return digits || clean
}
const matchesDateRange = (date, from, to) => {
  const value = iso(date)
  if (!from && !to) return true
  return Boolean(value) && (!from || value >= from) && (!to || value <= to)
}
const isWorkday = (date) => ![5, 6].includes(date.getDay()) // Israel default: Sunday–Thursday
const daysInMonth = (key) => {
  const [year, month] = key.split('-').map(Number)
  return new Date(year, month, 0).getDate()
}
const workdayCount = (key, startDay = 1, endDay = daysInMonth(key)) => {
  const [year, month] = key.split('-').map(Number)
  let count = 0
  for (let day = startDay; day <= endDay; day++) if (isWorkday(new Date(year, month - 1, day))) count++
  return count
}
const parseMonth = (value, fallbackDate = null) => {
  if (value instanceof Date) return monthKey(value)
  if (typeof value === 'number') {
    const d = excelDate(value); if (d) return monthKey(d)
  }
  const text = normalize(value)
  let m = text.match(/(20\d{2})[.\/-](\d{1,2})/)
  if (m) return `${m[1]}-${String(Number(m[2])).padStart(2, '0')}`
  m = text.match(/(\d{1,2})[.\/-](20\d{2})/)
  if (m) return `${m[2]}-${String(Number(m[1])).padStart(2, '0')}`
  return fallbackDate ? monthKey(fallbackDate) : ''
}

async function readWorkbook(file) {
  const buf = await file.arrayBuffer()
  const wb = XLSX.read(buf, { type: 'array', cellDates: true })
  let rows = []
  for (const sheetName of wb.SheetNames) {
    const sheetRows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '', raw: true })
    if (sheetRows.length) rows = rows.concat(sheetRows.map(row => ({ ...row, __sheet: sheetName })))
  }
  return rows
}
function classifyFile(rows) {
  // Some workbooks start with a title/summary sheet. Inspect headers across
  // a sample of rows instead of relying only on the first row.
  const keys = [...new Set(rows.slice(0, 250).flatMap(row => Object.keys(row || {}).map(normKey)))]
  const hasAny = (...terms) => terms.some(term => keys.some(k => k.includes(normKey(term))))
  const hasAll = (...groups) => groups.every(group => group.some(term => hasAny(term)))

  const targetLike = hasAll(
    ['monthly target', 'monthly plan', 'יעד חודשי', 'תוכנית חודשית', 'target', 'plan'],
    ['facility', 'מתקן', 'resource', 'משאב']
  )
  if (targetLike) return 'targets'

  const productionLike = hasAny('actual finish time', 'delivered quantity', 'confirmed yield quantity') &&
    hasAny('storage location', 'order', 'batch')
  if (productionLike) return 'production'

  const deviationLike = hasAny('rejected characteristics', 'qa status', 'ud remarks', 'restricted - recycling')
  if (deviationLike) return 'deviations'

  const qualityLike = hasAny(
    'inspection lot', 'inspection lot #', 'inspection lot storage location',
    'master insp characteristic', 'master insp charactristic',
    'result status', 'qa approval', 'start date of inspection', 'end date of inspection'
  )
  if (qualityLike) return 'quality'

  return 'unknown'
}

export default function App() {
  const [production, setProduction] = useState([])
  const [quality, setQuality] = useState([])
  const [deviations, setDeviations] = useState([])
  const [targets, setTargets] = useState([])
  const [status, setStatus] = useState('ממתין לטעינת קבצים')
  const [busy, setBusy] = useState(false)
  const [query, setQuery] = useState('')
  const [from, setFrom] = useState('')
  const [to, setTo] = useState('')
  const [selectedFacilities, setSelectedFacilities] = useState([])
  const [activeTab, setActiveTab] = useState('production')
  const [planningMonth, setPlanningMonth] = useState('')
  const [additionalFacilities, setAdditionalFacilities] = useState([])
  const [facilityToAdd, setFacilityToAdd] = useState('')
  const [periodYear, setPeriodYear] = useState('')
  const [periodQuarter, setPeriodQuarter] = useState('')

  useEffect(() => {
    try {
      const saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || 'null')
      if (saved) {
        setProduction(saved.production || []); setQuality(saved.quality || [])
        setDeviations(saved.deviations || []); setTargets(saved.targets || [])
        setStatus('הנתונים האחרונים שוחזרו מהדפדפן')
      }
    } catch (e) { console.warn('Could not restore saved data', e) }
  }, [])

  useEffect(() => {
    if (!production.length && !quality.length && !deviations.length && !targets.length) return
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify({ production, quality, deviations, targets, savedAt: new Date().toISOString() }))
    } catch (e) { setStatus('הנתונים נטענו, אך לא ניתן היה לשמור אותם בדפדפן') }
  }, [production, quality, deviations, targets])

  const handleFiles = async (files) => {
    if (!files.length) return
    setBusy(true)
    try {
      const loaded = []
      for (const file of files) {
        setStatus(`קורא את ${file.name}...`)
        const rows = await readWorkbook(file)
        const kind = classifyFile(rows)
        if (kind === 'production') setProduction(rows)
        if (kind === 'quality') setQuality(rows)
        if (kind === 'deviations') setDeviations(rows)
        if (kind === 'targets') {
          const fallbackMonth = parseMonth(file.name, new Date())
          const parsed = rows.map(r => ({
            facility: canonicalFacility(getField(r, ['Facility', 'מתקן', 'Storage Location', 'Resource', 'משאב'])),
            month: parseMonth(getField(r, ['Month', 'חודש', 'Target Month', 'Plan Month']), fallbackMonth),
            activity: normalize(getField(r, ['Activity', 'Type', 'סוג פעילות', 'סוג', 'Production/Packaging'])) || 'אריזה',
            target: num(getField(r, ['Monthly Target', 'Monthly Plan', 'Plan', 'יעד חודשי', 'תוכנית חודשית', 'Target'])),
            capacity: num(getField(r, ['Capacity', 'קיבולת', 'Monthly Capacity', 'קיבולת חודשית'])),
            notes: normalize(getField(r, ['Notes', 'Remarks', 'הערות'])),
          })).filter(r => r.facility && r.target > 0)
          setTargets(parsed)
          if (parsed[0]?.month) setPlanningMonth(parsed[0].month)
        }
        loaded.push(`${file.name}: ${fmt(rows.length)} רשומות${kind === 'unknown' ? ' (סוג לא זוהה)' : ''}`)
      }
      setStatus(`נטען בהצלחה — ${loaded.join(' | ')}`)
    } catch (e) {
      console.error(e); setStatus(`שגיאה בקריאת קובץ: ${e.message}`)
    } finally { setBusy(false) }
  }

  const prod = useMemo(() => production.map(r => {
    const finish = combineExcelDateTime(
      getField(r, ['Actual finish date', 'Actual Finish Date']),
      getField(r, ['Actual Finish Time', 'Actual finish time']),
      getField(r, ['Release date (actual)', 'Time Stamp'])
    )
    return {
      facility: canonicalFacility(getField(r, ['Storage Location', 'Storage location'])),
      date: finish,
      qty: num(getField(r, ['Delivered quantity (GMEIN)', 'Confirmed Yield Quantity (GMEIN)', 'Delivered quantity'])),
      order: normalize(getField(r, ['Order', 'Process Order', 'Work Order'])),
      batch: normalize(getField(r, ['Batch', 'Batch Number'])),
      material: normalize(getField(r, ['Material'])),
      desc: normalize(getField(r, ['Material description', 'Material Description'])),
      orderType: normalize(getField(r, ['Order Type'])),
      hour: finish ? finish.getHours() : null,
    }
  }).filter(r => r.facility), [production])

  const qualityRows = useMemo(() => quality.map(r => ({
    facility: canonicalFacility(getField(r, ['Inspection Lot Storage Location', 'Process Order Storage Location', 'Storage Location', 'Facility', 'Production Line'])),
    date: excelDate(getField(r, ['Date of Lot Creation', 'Start Date of Inspection', 'Process Order Confirmed Release Date', 'End Date of Inspection', 'Inspection Lot UD Date', 'Process Order Delivered Date'])),
    batch: normalize(getField(r, ['Batch', 'Batch Number'])), material: normalize(getField(r, ['Material', 'Material #'])),
    order: normalize(getField(r, ['Process Order', 'Process Order #', 'Order'])), status: normalize(getField(r, ['Result Status', 'QA Approval', 'Status'])),
    inspectionLot: normalize(getField(r, ['Inspection Lot', 'Inspection Lot #'])),
  })), [quality])

  const deviationRows = useMemo(() => deviations.map(r => ({
    facility: canonicalFacility(getField(r, ['Facility', 'Production Line', 'Storage Location'])),
    date: excelDate(getField(r, ['Date of Lot Creation', 'Inspection Lot UD Date', 'Process Order Delivered Date', 'Start Date of Inspection'])),
    batch: normalize(getField(r, ['Batch', 'Batch Number'])), material: normalize(getField(r, ['Material'])),
    status: normalize(getField(r, ['QA Status', 'Status'])), rejected: normalize(getField(r, ['Rejected characteristics', 'Rejected characteristics '])),
    remarks: normalize(getField(r, ['UD Remarks', 'Remarks'])),
  })), [deviations])

  const dataMonths = useMemo(() => [...new Set(prod.map(r => monthKey(r.date)).filter(Boolean))].sort(), [prod])
  const targetMonths = useMemo(() => [...new Set(targets.map(r => r.month).filter(Boolean))].sort(), [targets])
  const availableMonths = useMemo(() => [...new Set([...targetMonths, ...dataMonths])].sort().reverse(), [targetMonths, dataMonths])
  useEffect(() => { if (!planningMonth && availableMonths.length) setPlanningMonth(availableMonths[0]) }, [availableMonths, planningMonth])

  const dateBounds = useMemo(() => {
    const ds = prod.map(r => r.date).filter(Boolean).sort((a, b) => a - b)
    return { min: iso(ds[0]), max: iso(ds.at(-1)) }
  }, [prod])

  const baseFiltered = useMemo(() => {
    const q = query.trim().toLowerCase()
    return prod.filter(r => matchesDateRange(r.date, from, to) && (!q || [r.facility, r.order, r.batch, r.material, r.desc].some(v => v.toLowerCase().includes(q))))
  }, [prod, from, to, query])
  const filtered = useMemo(() => baseFiltered.filter(r => !selectedFacilities.length || selectedFacilities.includes(r.facility)), [baseFiltered, selectedFacilities])

  const discoveredFacilities = useMemo(() => [...new Set([...targets.map(t => t.facility), ...prod.map(r => r.facility)].filter(Boolean))].sort(), [targets, prod])
  const optionalFacilities = useMemo(() => discoveredFacilities.filter(id => !PRIMARY_FACILITIES.includes(id) && !additionalFacilities.includes(id)), [discoveredFacilities, additionalFacilities])
  const facilities = useMemo(() => [...PRIMARY_FACILITIES, ...additionalFacilities], [additionalFacilities])
  const availableYears = useMemo(() => [...new Set(prod.map(r => r.date?.getFullYear()).filter(Boolean))].sort((a,b) => b-a), [prod])

  const planningRows = useMemo(() => {
    if (!planningMonth) return []
    const [year, month] = planningMonth.split('-').map(Number)
    const monthRows = prod.filter(r => monthKey(r.date) === planningMonth)
    const latestDataDate = monthRows.map(r => r.date).filter(Boolean).sort((a,b) => b-a)[0]
    const now = new Date()
    const isCurrent = now.getFullYear() === year && now.getMonth() + 1 === month
    const isPast = new Date(year, month, 0) < new Date(now.getFullYear(), now.getMonth(), 1)
    const asOfDay = isCurrent ? Math.min(now.getDate(), daysInMonth(planningMonth)) : isPast ? daysInMonth(planningMonth) : (latestDataDate?.getDate() || 0)
    const elapsedWorkdays = workdayCount(planningMonth, 1, Math.max(0, asOfDay))
    const totalWorkdays = workdayCount(planningMonth)
    const remainingWorkdays = Math.max(0, totalWorkdays - elapsedWorkdays)
    const monthTargets = targets.filter(t => t.month === planningMonth)
    const ids = facilities
    return ids.map(id => {
      let rows = monthRows.filter(r => r.facility === id)
      if (id === '1542') rows = rows.filter(r => r.orderType.toUpperCase().includes('ZFIN'))
      const actual = rows.reduce((s, r) => s + r.qty, 0)
      const targetRow = monthTargets.find(t => t.facility === id)
      const target = targetRow?.target || 0
      const dailyMap = new Map()
      rows.forEach(r => { const d = iso(r.date); if (d) dailyMap.set(d, (dailyMap.get(d) || 0) + r.qty) })
      const dailyValues = [...dailyMap.values()]
      const actualDays = dailyValues.length
      const average = actualDays ? actual / actualDays : 0
      const recent = [...dailyMap.entries()].sort((a,b) => a[0].localeCompare(b[0])).slice(-7).map(([,v]) => v)
      const recentAverage = recent.length ? recent.reduce((a,b) => a+b, 0) / recent.length : average
      const provenMax = dailyValues.length ? Math.max(...dailyValues) : (targetRow?.capacity ? targetRow.capacity / Math.max(1, totalWorkdays) : LEGACY_DAILY_TARGETS[id] || 0)
      const remaining = Math.max(0, target - actual)
      const requiredDaily = remainingWorkdays ? remaining / remainingWorkdays : remaining
      const forecast = actual + recentAverage * remainingWorkdays
      const capacityForecast = actual + provenMax * remainingWorkdays
      let state = 'no-target', label = 'ללא יעד'
      if (target > 0 && actual >= target) { state = 'achieved'; label = 'היעד הושג' }
      else if (target > 0 && remainingWorkdays === 0) { state = 'risk'; label = 'היעד לא הושג' }
      else if (target > 0 && provenMax > 0 && requiredDaily > provenMax) { state = 'risk'; label = 'לא בר־השגה' }
      else if (target > 0 && forecast < target) { state = 'warning'; label = 'נדרש שיפור קצב' }
      else if (target > 0) { state = 'good'; label = 'במסלול ליעד' }
      return {
        id, activity: targetRow?.activity || 'אריזה', target, capacity: targetRow?.capacity || 0, actual,
        pct: target ? actual / target * 100 : 0, remaining, requiredDaily, average, recentAverage, provenMax,
        forecast, capacityForecast, elapsedWorkdays, remainingWorkdays, totalWorkdays,
        orders: new Set(rows.map(r => r.order).filter(Boolean)).size, state, label,
      }
    })
  }, [planningMonth, prod, targets, facilities])

  const facilityStats = useMemo(() => facilities.map(id => {
    let rows = baseFiltered.filter(r => r.facility === id)
    if (id === '1542') rows = rows.filter(r => r.orderType.toUpperCase().includes('ZFIN'))
    const actual = rows.reduce((s, r) => s + r.qty, 0)
    const plan = planningRows.find(x => x.id === id)
    return { id, actual, target: plan?.target || 0, pct: plan?.target ? actual / plan.target * 100 : 0, orders: new Set(rows.map(r => r.order).filter(Boolean)).size, state: plan?.state || 'no-target', forecast: plan?.forecast || 0 }
  }), [facilities, baseFiltered, planningRows])

  const toggleFacility = (id) => setSelectedFacilities(current => current.includes(id) ? current.filter(x => x !== id) : [...current, id])
  const allFacilitiesSelected = selectedFacilities.length === facilities.length
  const toggleAllFacilities = () => setSelectedFacilities(allFacilitiesSelected ? [] : [...facilities])

  const total = filtered.reduce((s, x) => s + x.qty, 0)
  const activeFacilities = facilityStats.filter(x => x.actual > 0).length
  const morningQty = filtered.filter(r => r.hour !== null && r.hour >= 7 && r.hour < 19).reduce((s, r) => s + r.qty, 0)
  const nightQty = filtered.filter(r => r.hour !== null && (r.hour >= 19 || r.hour < 7)).reduce((s, r) => s + r.qty, 0)
  const allQualityBad = qualityRows.filter(r => { const st = r.status.toLowerCase(); return st && !['accepted', 'תקין', 'pass', 'approved', 'מאושר'].some(x => st.includes(x)) })
  const allOpenDeviations = deviationRows.filter(r => { const st = r.status.toLowerCase(); return !st || !['approved', 'closed', 'מאושר', 'סגור'].some(x => st.includes(x)) })
  const qualityBad = useMemo(() => allQualityBad.filter(r => (!selectedFacilities.length || selectedFacilities.includes(r.facility)) && matchesDateRange(r.date, from, to)), [allQualityBad, selectedFacilities, from, to])
  const openDeviations = useMemo(() => allOpenDeviations.filter(r => (!selectedFacilities.length || selectedFacilities.includes(r.facility)) && matchesDateRange(r.date, from, to)), [allOpenDeviations, selectedFacilities, from, to])

  const alerts = useMemo(() => planningRows.filter(r => ['risk', 'warning'].includes(r.state)).sort((a,b) => ({ risk:0, warning:1 }[a.state] - { risk:0, warning:1 }[b.state])), [planningRows])
  const achievedCount = planningRows.filter(r => ['achieved', 'good'].includes(r.state)).length
  const riskCount = planningRows.filter(r => r.state === 'risk').length
  const warningCount = planningRows.filter(r => r.state === 'warning').length
  const targetTotal = planningRows.reduce((s,r) => s + r.target, 0)
  const targetActual = planningRows.reduce((s,r) => s + r.actual, 0)
  const targetForecast = planningRows.reduce((s,r) => s + r.forecast, 0)

  const exportWorkbook = () => {
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filtered.map(r => ({ Date: iso(r.date), Time: r.date ? r.date.toLocaleTimeString('he-IL', {hour:'2-digit',minute:'2-digit'}) : '', Facility: r.facility, Order: r.order, Batch: r.batch, Material: r.material, Description: r.desc, Quantity: r.qty }))), 'Production')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(planningRows.map(r => ({ Month: planningMonth, Facility: r.id, Activity: r.activity, MonthlyTarget: r.target, Actual: r.actual, Remaining: r.remaining, RequiredDaily: r.requiredDaily, RecentAverage: r.recentAverage, ProvenMax: r.provenMax, Forecast: r.forecast, Status: r.label }))), 'Planning')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(qualityBad.map(r => ({ Date: iso(r.date), Facility: r.facility, InspectionLot: r.inspectionLot, Order: r.order, Batch: r.batch, Material: r.material, Status: r.status }))), 'Quality')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(openDeviations.map(r => ({ Date: iso(r.date), Facility: r.facility, Batch: r.batch, Material: r.material, Status: r.status, Rejected: r.rejected, Remarks: r.remarks }))), 'Deviations')
    XLSX.writeFile(wb, `IML_Sprint7_Report_${new Date().toISOString().slice(0,10)}.xlsx`)
  }
  const downloadTargetTemplate = () => {
    const rows = DEFAULT_FACILITIES.map(id => ({ 'חודש': planningMonth || monthKey(new Date()), 'מתקן': id, 'סוג פעילות': 'אריזה', 'יעד חודשי': '', 'קיבולת חודשית': '', 'הערות': '' }))
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'יעדים חודשיים'); XLSX.writeFile(wb, 'IML_Monthly_Targets_Template.xlsx')
  }
  const clearAllData = () => {
    if (!window.confirm('למחוק את כל הנתונים והיעדים השמורים בדפדפן?')) return
    setProduction([]); setQuality([]); setDeviations([]); setTargets([]); localStorage.removeItem(STORAGE_KEY)
    setStatus('כל הנתונים נמחקו'); setFrom(''); setTo(''); setQuery(''); setSelectedFacilities([]); setPlanningMonth(''); setAdditionalFacilities([]); setFacilityToAdd(''); setPeriodYear(''); setPeriodQuarter('')
  }
  const dailyTrend = useMemo(() => {
    const map = new Map(); filtered.forEach(r => { const d = iso(r.date); if (d) map.set(d, (map.get(d) || 0) + r.qty) })
    return [...map.entries()].sort((a,b) => a[0].localeCompare(b[0])).slice(-14)
  }, [filtered])
  const maxDaily = Math.max(1, ...dailyTrend.map(([,v]) => v))
  const setQuickRange = (daysBack) => {
    if (!dateBounds.max) return
    const end = new Date(`${dateBounds.max}T12:00:00`), start = new Date(end); start.setDate(end.getDate() - daysBack + 1)
    setFrom(iso(start)); setTo(dateBounds.max); setPeriodYear(''); setPeriodQuarter('')
  }
  const applyYearFilter = (yearValue) => {
    setPeriodYear(yearValue); setPeriodQuarter('')
    if (!yearValue) { setFrom(''); setTo(''); return }
    setFrom(`${yearValue}-01-01`); setTo(`${yearValue}-12-31`)
  }
  const applyQuarterFilter = (quarterValue) => {
    setPeriodQuarter(quarterValue)
    if (!periodYear || !quarterValue) return
    const q = Number(quarterValue), startMonth = (q - 1) * 3 + 1, endMonth = startMonth + 2
    const lastDay = new Date(Number(periodYear), endMonth, 0).getDate()
    setFrom(`${periodYear}-${String(startMonth).padStart(2,'0')}-01`)
    setTo(`${periodYear}-${String(endMonth).padStart(2,'0')}-${String(lastDay).padStart(2,'0')}`)
  }
  const addFacility = () => {
    if (!facilityToAdd || additionalFacilities.includes(facilityToAdd)) return
    setAdditionalFacilities(current => [...current, facilityToAdd]); setFacilityToAdd('')
  }
  const removeAdditionalFacility = (id) => {
    setAdditionalFacilities(current => current.filter(x => x !== id))
    setSelectedFacilities(current => current.filter(x => x !== id))
  }

  return <div className="dashboard" dir="rtl">
    <aside className="side">
      <div className="brand">IML<span>CONTROL</span></div>
      <div className="side-stat"><Database/><div><b>{fmt(production.length)}</b><small>רשומות תפוקה</small></div></div>
      <div className="side-stat"><Target/><div><b>{targets.length}</b><small>יעדים חודשיים</small></div></div>
      <div className="side-stat"><FlaskConical/><div><b>{fmt(quality.length + deviations.length)}</b><small>רשומות איכות</small></div></div>
      <div className="side-note">Sprint 7: יעד חודשי, קצב דינמי, תחזית והתראות. הנתונים נשמרים בדפדפן בלבד.</div>
    </aside>

    <main className="main">
      <header className="header">
        <div><h1>מרכז שליטה למתקני אריזה</h1><p>תפוקה, איכות, יעדים חודשיים ותחזית סוף חודש</p></div>
        <div className="header-actions">
          <button className="action secondary" onClick={downloadTargetTemplate}><FileSpreadsheet size={18}/> תבנית יעדים</button>
          <button className="action secondary" onClick={exportWorkbook} disabled={!production.length}><Download size={18}/> יצוא Excel</button>
          <button className="action danger" onClick={clearAllData} disabled={!production.length && !quality.length && !deviations.length && !targets.length}><Trash2 size={18}/> מחיקה</button>
          <label className={`upload ${busy ? 'disabled' : ''}`}><Upload size={19}/>{busy ? 'טוען...' : 'טעינת Excel'}<input type="file" multiple accept=".xlsx,.xls" disabled={busy} onChange={e => handleFiles([...e.target.files])}/></label>
        </div>
      </header>

      <div className="load-status"><CheckCircle2 size={18}/>{status}</div>

      <section className="planning-toolbar">
        <div><Target/><span>חודש תכנון</span><select value={planningMonth} onChange={e => setPlanningMonth(e.target.value)}>{!availableMonths.length && <option value="">אין נתונים</option>}{availableMonths.map(m => <option key={m} value={m}>{m}</option>)}</select></div>
        <small>היעד היומי אינו מוזן: הוא מחושב מחדש בכל יום לפי יתרת היעד וימי העבודה שנותרו.</small>
      </section>

      <section className="executive-strip">
        <Executive icon={<Target/>} title="ביצוע מול יעד" value={targetTotal ? pctFmt(targetActual / targetTotal * 100) : '—'} sub={`${fmt(targetActual)} מתוך ${fmt(targetTotal)}`}/>
        <Executive icon={<TrendingUp/>} title="תחזית סוף חודש" value={targetTotal ? pctFmt(targetForecast / targetTotal * 100) : '—'} sub={fmt(targetForecast)}/>
        <Executive icon={<CheckCircle2/>} title="במסלול / הושג" value={achievedCount} sub="מתקנים" good/>
        <Executive icon={<AlertTriangle/>} title="בסיכון" value={warningCount} sub="נדרש שיפור" warn/>
        <Executive icon={<BellRing/>} title="לא בר־השגה" value={riskCount} sub="דורש פעולה" bad/>
      </section>

      <section className="filters">
        <label><Search size={17}/><input value={query} onChange={e => setQuery(e.target.value)} placeholder="חיפוש מתקן, Order, Batch או חומר"/></label>
        <label><CalendarDays size={17}/>שנה<select value={periodYear} onChange={e => applyYearFilter(e.target.value)}><option value="">כל השנים</option>{availableYears.map(y => <option key={y} value={y}>{y}</option>)}</select></label>
        <label><CalendarDays size={17}/>רבעון<select value={periodQuarter} disabled={!periodYear} onChange={e => applyQuarterFilter(e.target.value)}><option value="">כל השנה</option><option value="1">רבעון 1</option><option value="2">רבעון 2</option><option value="3">רבעון 3</option><option value="4">רבעון 4</option></select></label>
        <label><CalendarDays size={17}/>מתאריך<input type="date" min={dateBounds.min} max={dateBounds.max} value={from} onChange={e => { setFrom(e.target.value); setPeriodYear(''); setPeriodQuarter('') }}/></label>
        <label><CalendarDays size={17}/>עד תאריך<input type="date" min={dateBounds.min} max={dateBounds.max} value={to} onChange={e => { setTo(e.target.value); setPeriodYear(''); setPeriodQuarter('') }}/></label>
        <button onClick={() => setQuickRange(1)}>יום אחרון</button><button onClick={() => setQuickRange(2)}>יומיים</button><button onClick={() => setQuickRange(30)}>30 יום</button>
        <button onClick={() => { setFrom(''); setTo(''); setQuery(''); setSelectedFacilities([]); setPeriodYear(''); setPeriodQuarter('') }}>נקה</button>
      </section>

      <section className="extra-facilities">
        <div><Factory size={18}/><strong>מתקנים נוספים</strong><span>ברירת המחדל מציגה רק את מתקני הליבה.</span></div>
        <div className="extra-facility-actions"><select value={facilityToAdd} onChange={e => setFacilityToAdd(e.target.value)}><option value="">בחר מתקן נוסף</option>{optionalFacilities.map(id => <option key={id} value={id}>{id}</option>)}</select><button onClick={addFacility} disabled={!facilityToAdd}>הוסף מתקן</button></div>
        {!!additionalFacilities.length && <div className="extra-facility-chips">{additionalFacilities.map(id => <button key={id} onClick={() => removeAdditionalFacility(id)}>{id}<X size={14}/></button>)}</div>}
      </section>

      {selectedFacilities.length > 0 && <div className="selection"><div className="selection-info"><span>מתקנים נבחרים:</span><div className="selection-chips">{selectedFacilities.map(id => <button className="selection-chip" key={id} onClick={() => toggleFacility(id)}>{id}<X size={14}/></button>)}</div></div><button className="clear-selection" onClick={() => setSelectedFacilities([])}><X size={16}/> הסר הכול</button></div>}

      <section className="summary-grid">
        <Summary title="סה״כ תפוקה מסוננת" value={fmt(total)} sub="ליטר / ק״ג לפי הקובץ"/>
        <Summary title="מתקנים פעילים" value={activeFacilities} sub={`מתוך ${facilities.length} מתקנים`}/>
        <Summary title="משמרת בוקר" value={fmt(morningQty)} sub="07:00–19:00"/>
        <Summary title="משמרת לילה" value={fmt(nightQty)} sub="19:00–07:00"/>
        <Summary title="חריגות פתוחות" value={openDeviations.length} sub="לפי קובץ החריגות" warn/>
        <Summary title="איכות לא תקינה" value={qualityBad.length} sub="לפי קובץ האיכות" warn/>
      </section>

      <div className="section-title facility-title"><div className="section-title-text"><Gauge/><div><h2>תחזית חודשית לפי מתקן</h2><p>יעד חודשי, קצב נדרש, קצב אחרון, שיא מוכח ותחזית</p></div></div><button className="select-all-facilities" onClick={toggleAllFacilities}><CheckCircle2 size={17}/>{allFacilitiesSelected ? 'ביטול בחירת הכול' : 'בחירת כל המתקנים'}</button></div>
      <section className="forecast-grid">
        {planningRows.map(row => <ForecastCard key={row.id} {...row} selected={selectedFacilities.includes(row.id)} onClick={() => toggleFacility(row.id)}/>) }
        {!planningRows.length && <div className="empty wide-empty">טען קובץ יעדים חודשי כדי להציג תחזית.</div>}
      </section>

      <section className="alert-panel">
        <div className="panel-head"><div><BellRing/><h2>מה דורש תשומת לב היום?</h2></div><span>{alerts.length} התראות</span></div>
        <div className="alert-list">
          {alerts.map(r => <div className={`alert-item ${r.state}`} key={r.id}><div className="alert-symbol">{r.state === 'risk' ? '!' : '⚠'}</div><div><strong>מתקן {r.id} — {r.label}</strong><p>{r.state === 'risk' ? `נדרש ${fmt(r.requiredDaily)} ליום, אך השיא המוכח הוא ${fmt(r.provenMax)}.` : `התחזית היא ${fmt(r.forecast)} מול יעד ${fmt(r.target)}. נדרש קצב של ${fmt(r.requiredDaily)} ליום.`}</p></div></div>)}
          {!alerts.length && <div className="empty">אין התראות תכנון לחודש הנבחר.</div>}
        </div>
      </section>

      <section className="daily-management">
        <div className="panel-head"><div><CalendarCheck/><h2>Daily Management</h2></div><span>{planningMonth}</span></div>
        <div className="table-wrap"><table><thead><tr><th>מתקן</th><th>פעילות</th><th>יעד חודשי</th><th>בפועל</th><th>% ביצוע</th><th>נותר</th><th>ימי עבודה נותרו</th><th>נדרש ליום</th><th>ממוצע 7 ימים</th><th>שיא מוכח</th><th>תחזית</th><th>סטטוס</th></tr></thead><tbody>
          {planningRows.map(r => <tr key={r.id}><td><b>{r.id}</b></td><td>{r.activity}</td><td>{fmt(r.target)}</td><td>{fmt(r.actual)}</td><td>{pctFmt(r.pct)}</td><td>{fmt(r.remaining)}</td><td>{r.remainingWorkdays}</td><td>{fmt(r.requiredDaily)}</td><td>{fmt(r.recentAverage)}</td><td>{fmt(r.provenMax)}</td><td>{fmt(r.forecast)}</td><td><StatusBadge state={r.state} label={r.label}/></td></tr>)}
          {!planningRows.length && <tr><td colSpan="12" className="empty">אין יעדים לחודש הנבחר</td></tr>}
        </tbody></table></div>
      </section>

      <section className="trend-card">
        <div className="trend-head"><div><h2>מגמת תפוקה יומית</h2><p>14 הימים האחרונים בטווח הנבחר</p></div><Save size={20}/></div>
        <div className="trend-bars">{dailyTrend.map(([date, value]) => <div className="trend-item" key={date} title={`${date}: ${fmt(value)}`}><div className="trend-value">{fmt(value)}</div><div className="trend-track"><i style={{height: `${Math.max(5, value / maxDaily * 100)}%`}}/></div><small>{date.slice(5)}</small></div>)}{!dailyTrend.length && <div className="empty trend-empty">טען קובץ תפוקות להצגת מגמה</div>}</div>
      </section>

      <div className="section-title facility-title"><div className="section-title-text"><Factory/><div><h2>ביצועים לפי מתקן בטווח המסונן</h2><p>לחיצה על כרטיס מסננת תפוקה, איכות וחריגות</p></div></div></div>
      <section className="facility-grid">{facilityStats.map(x => <Facility key={x.id} {...x} selected={selectedFacilities.includes(x.id)} onClick={() => toggleFacility(x.id)}/>)}</section>

      <section className="tabs">
        <button className={activeTab === 'production' ? 'active' : ''} onClick={() => setActiveTab('production')}><BarChart3 size={16}/> תפוקה</button>
        <button className={activeTab === 'quality' ? 'active' : ''} onClick={() => setActiveTab('quality')}><FlaskConical size={16}/> איכות ({qualityBad.length})</button>
        <button className={activeTab === 'deviations' ? 'active' : ''} onClick={() => setActiveTab('deviations')}><AlertTriangle size={16}/> מנות חריגות ({openDeviations.length})</button>
      </section>
      {activeTab === 'production' && <section className="details"><h2>רשומות תפוקה אחרונות</h2><div className="table-wrap"><table><thead><tr><th>תאריך</th><th>שעה</th><th>מתקן</th><th>הזמנה</th><th>Batch</th><th>חומר</th><th>כמות</th></tr></thead><tbody>{filtered.slice(-200).reverse().map((r, i) => <tr key={i}><td>{iso(r.date)}</td><td>{r.date ? r.date.toLocaleTimeString('he-IL', {hour:'2-digit',minute:'2-digit'}) : ''}</td><td>{r.facility}</td><td>{r.order}</td><td>{r.batch}</td><td>{r.desc || r.material}</td><td>{fmt(r.qty)}</td></tr>)}{!filtered.length && <tr><td colSpan="7" className="empty">טען קובץ תפוקות כדי להציג נתונים</td></tr>}</tbody></table></div></section>}
      {activeTab === 'quality' && <section className="details"><h2>תוצאות איכות לא תקינות</h2><div className="table-wrap"><table><thead><tr><th>תאריך</th><th>מתקן</th><th>Inspection Lot</th><th>Order</th><th>Batch</th><th>Material</th><th>סטטוס</th></tr></thead><tbody>{qualityBad.slice(0,300).map((r,i) => <tr key={i}><td>{iso(r.date)}</td><td>{r.facility}</td><td>{r.inspectionLot}</td><td>{r.order}</td><td>{r.batch}</td><td>{r.material}</td><td><span className="status-bad">{r.status || 'ללא סטטוס'}</span></td></tr>)}{!qualityBad.length && <tr><td colSpan="7" className="empty">לא נמצאו תוצאות איכות לא תקינות</td></tr>}</tbody></table></div></section>}
      {activeTab === 'deviations' && <section className="details"><h2>מנות חריגות פתוחות</h2><div className="table-wrap"><table><thead><tr><th>תאריך</th><th>מתקן</th><th>Batch</th><th>Material</th><th>סטטוס</th><th>מאפיינים שנדחו</th><th>הערות</th></tr></thead><tbody>{openDeviations.slice(0,300).map((r,i) => <tr key={i}><td>{iso(r.date)}</td><td>{r.facility}</td><td>{r.batch}</td><td>{r.material}</td><td><span className="status-bad">{r.status || 'פתוח'}</span></td><td>{r.rejected}</td><td>{r.remarks}</td></tr>)}{!openDeviations.length && <tr><td colSpan="7" className="empty">לא נמצאו מנות חריגות פתוחות</td></tr>}</tbody></table></div></section>}
    </main>
  </div>
}

function Summary({ title, value, sub, warn }) { return <div className={`summary ${warn ? 'warn' : ''}`}><span>{title}</span><b>{value}</b><small>{sub}</small></div> }
function Executive({ icon, title, value, sub, good, warn, bad }) { return <div className={`executive ${good?'good':''} ${warn?'warn':''} ${bad?'bad':''}`}><div className="executive-icon">{icon}</div><div><span>{title}</span><b>{value}</b><small>{sub}</small></div></div> }
function StatusBadge({ state, label }) { return <span className={`status-pill ${state}`}>{label}</span> }
function ForecastCard({ id, target, actual, pct, remaining, requiredDaily, recentAverage, provenMax, forecast, remainingWorkdays, state, label, selected, onClick }) {
  return <article className={`forecast-card ${state} ${selected ? 'selected' : ''}`} onClick={onClick} role="button" tabIndex="0">
    <div className="forecast-head"><div><small>מתקן</small><h3>{id}</h3></div><StatusBadge state={state} label={label}/></div>
    <div className="forecast-main"><div><span>ביצוע</span><b>{pctFmt(pct)}</b></div><div><span>תחזית</span><b>{target ? pctFmt(forecast / target * 100) : '—'}</b></div></div>
    <div className="bar"><i style={{width:`${Math.min(100, pct)}%`}}/></div>
    <div className="forecast-values"><span>יעד<strong>{fmt(target)}</strong></span><span>בפועל<strong>{fmt(actual)}</strong></span><span>נותר<strong>{fmt(remaining)}</strong></span></div>
    <div className="forecast-metrics"><div><span>ימים נותרו</span><b>{remainingWorkdays}</b></div><div><span>נדרש ליום</span><b>{fmt(requiredDaily)}</b></div><div><span>7 ימים</span><b>{fmt(recentAverage)}</b></div><div><span>שיא מוכח</span><b>{fmt(provenMax)}</b></div></div>
  </article>
}
function Facility({ id, actual, target, pct, orders, selected, onClick }) {
  const state = !target ? 'no-target' : pct >= 100 ? 'good' : pct >= 75 ? 'warning' : 'risk'
  return <article className={`facility ${state} ${selected ? 'selected' : ''}`} onClick={onClick} role="button" tabIndex="0"><div className="facility-top"><div><small>מתקן</small><h3>{id}</h3></div><b>{target ? pctFmt(pct) : '—'}</b></div><div className="bar"><i style={{width:`${Math.min(100,pct)}%`}}/></div><div className="facility-numbers"><span>בפועל<strong>{fmt(actual)}</strong></span><span>יעד חודשי<strong>{fmt(target)}</strong></span><span>הזמנות<strong>{orders}</strong></span></div></article>
}