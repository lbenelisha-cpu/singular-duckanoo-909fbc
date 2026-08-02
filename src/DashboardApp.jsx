import { useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx'
import {
  Upload, Database, Factory, FlaskConical, CalendarDays, Search, CheckCircle2,
  AlertTriangle, Clock3, X, BarChart3, Download, Trash2, Save, Target,
  Gauge, CalendarCheck, BellRing, TrendingUp, FileSpreadsheet, ShieldCheck, RefreshCw, ClipboardList, Activity, LogOut, UserCircle, Cloud, WifiOff, ArrowLeft, HeartPulse
} from 'lucide-react'
import { loadAllCloudDatasets, uploadCloudDataset, deleteAllCloudDatasets, getCloudHealth } from './cloudData'
import { supabase } from './supabase'
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
const RESOURCE_LABELS = {
  '1542|LQ-P-1': { station: 'P-02', line: '1 ליטר' },
  '1542|LQ-P-5': { station: 'P-03', line: '5 ליטר' },
  '1542|LQ-P-10': { station: 'P-04', line: '10/20 ליטר' },
}
const STORAGE_KEY = 'iml-control-center-sprint7'
const DB_NAME = 'iml-control-center-db'
const DB_STORE = 'dashboard-state'
const DB_KEY = 'sprint7'

const openDashboardDb = () => new Promise((resolve, reject) => {
  const request = indexedDB.open(DB_NAME, 1)
  request.onupgradeneeded = () => {
    const db = request.result
    if (!db.objectStoreNames.contains(DB_STORE)) db.createObjectStore(DB_STORE)
  }
  request.onsuccess = () => resolve(request.result)
  request.onerror = () => reject(request.error)
})
const idbGet = async () => {
  const db = await openDashboardDb()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, 'readonly')
    const req = tx.objectStore(DB_STORE).get(DB_KEY)
    req.onsuccess = () => resolve(req.result || null)
    req.onerror = () => reject(req.error)
    tx.oncomplete = () => db.close()
  })
}
const idbSet = async (value) => {
  const db = await openDashboardDb()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, 'readwrite')
    tx.objectStore(DB_STORE).put(value, DB_KEY)
    tx.oncomplete = () => { db.close(); resolve() }
    tx.onerror = () => { db.close(); reject(tx.error) }
  })
}
const idbClear = async () => {
  const db = await openDashboardDb()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, 'readwrite')
    tx.objectStore(DB_STORE).delete(DB_KEY)
    tx.oncomplete = () => { db.close(); resolve() }
    tx.onerror = () => { db.close(); reject(tx.error) }
  })
}

const normalize = (v) => String(v ?? '').trim()
const normalizeRouting = (v) => normalize(v).toUpperCase()
const resourceMeta = (facility, routingGroup) => RESOURCE_LABELS[`${facility}|${normalizeRouting(routingGroup)}`] || {}
const planningName = (row) => row.routingGroup ? `מתקן ${row.facility} · ${row.station || row.routingGroup}${row.lineName ? ` · ${row.lineName}` : ''}` : `מתקן ${row.facility}`
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

const shiftInfo = (date) => {
  if (!date) return { key: 'unknown', label: 'ללא שעה', order: 9, workDate: '' }
  const d = new Date(date)
  const hour = d.getHours()
  let key = 'morning', label = 'בוקר', order = 0
  if (hour >= 15 && hour < 23) { key = 'evening'; label = 'ערב'; order = 1 }
  else if (hour >= 23 || hour < 7) { key = 'night'; label = 'לילה'; order = 2 }
  const work = new Date(d)
  if (hour < 7) work.setDate(work.getDate() - 1)
  return { key, label, order, workDate: `${work.getFullYear()}-${String(work.getMonth()+1).padStart(2,'0')}-${String(work.getDate()).padStart(2,'0')}` }
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
  const wb = XLSX.read(buf, { type: 'array', cellDates: true, dense: true })
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

export default function DashboardApp({ currentUser, userRole = 'viewer', isGuest = false, onSignOut }) {
  const canManageData = ['admin', 'manager'].includes(userRole)
  const canDeleteData = userRole === 'admin'

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
  const [dataMeta, setDataMeta] = useState({ production:null, quality:null, deviations:null, targets:null })
  const [selectedBatch, setSelectedBatch] = useState('')
  const [cloudState, setCloudState] = useState({ mode:'connecting', lastSync:null, message:'מתחבר למסד המשותף...', latencyMs:null, live:false })
  const [uploadProgress, setUploadProgress] = useState(null)

  useEffect(() => {
    let active = true
    ;(async () => {
      try {
        setCloudState({ mode:'connecting', lastSync:null, message:'קורא את הנתונים המשותפים מ־Supabase...' })
        const cloud = await loadAllCloudDatasets(kind => {
          if (active) setStatus(`מסנכרן ${kind} מהענן...`)
        })
        if (!active) return
        const reviveQuality = (cloud.quality?.rows || []).map(row => row?.__compactQuality && row.date ? { ...row, date:new Date(row.date) } : row)
        setProduction(cloud.production?.rows || [])
        setQuality(reviveQuality)
        setDeviations(cloud.deviations?.rows || [])
        setTargets(cloud.targets?.rows || [])
        const meta = {
          production: cloud.production?.meta || null,
          quality: cloud.quality?.meta || null,
          deviations: cloud.deviations?.meta || null,
          targets: cloud.targets?.meta || null,
        }
        setDataMeta(meta)
        const lastSync = [meta.production,meta.quality,meta.deviations,meta.targets].map(x=>x?.loadedAt).filter(Boolean).sort().at(-1) || new Date().toISOString()
        const health = await getCloudHealth().catch(() => null)
        setCloudState({ mode:'cloud', lastSync, message:health?.versioned ? 'מחובר למסד הנתונים המשותף — מנוע גרסאות פעיל' : (health?.schemaMessage || 'מחובר לענן במצב תאימות'), latencyMs:health?.latencyMs ?? null, live:true })
        setStatus('הנתונים נטענו מ־Supabase וזמינים לכל המשתמשים')
      } catch (cloudError) {
        console.warn('Cloud restore failed; using browser cache', cloudError)
        try {
          const saved = await idbGet()
          if (!active) return
          if (saved) {
            setProduction(saved.production || []); setQuality(saved.quality || [])
            setDeviations(saved.deviations || []); setTargets(saved.targets || [])
            setDataMeta(saved.dataMeta || { production:null, quality:null, deviations:null, targets:null })
            setStatus('אין חיבור לענן — מוצג גיבוי מקומי מהדפדפן')
          } else setStatus('אין חיבור לענן ולא נמצא גיבוי מקומי')
        } catch (cacheError) { console.warn('Could not restore IndexedDB data', cacheError) }
        setCloudState({ mode:'offline', lastSync:null, message:cloudError?.message || 'אין חיבור למסד המשותף', latencyMs:null, live:false })
      }
    })()
    return () => { active = false }
  }, [])

  useEffect(() => {
    if (!supabase) return
    let refreshTimer
    const channel = supabase
      .channel('iml-data-sources-live')
      .on('postgres_changes', { event: '*', schema: 'public', table: 'iml_data_sources' }, () => {
        clearTimeout(refreshTimer)
        refreshTimer = setTimeout(async () => {
          try {
            setStatus('התקבל עדכון חדש מהענן — מסנכרן...')
            const cloud = await loadAllCloudDatasets()
            const reviveQuality = (cloud.quality?.rows || []).map(row => row?.__compactQuality && row.date ? { ...row, date:new Date(row.date) } : row)
            setProduction(cloud.production?.rows || [])
            setQuality(reviveQuality)
            setDeviations(cloud.deviations?.rows || [])
            setTargets(cloud.targets?.rows || [])
            const meta = { production:cloud.production?.meta||null, quality:cloud.quality?.meta||null, deviations:cloud.deviations?.meta||null, targets:cloud.targets?.meta||null }
            setDataMeta(meta)
            const lastSync = [meta.production,meta.quality,meta.deviations,meta.targets].map(x=>x?.loadedAt).filter(Boolean).sort().at(-1) || new Date().toISOString()
            setCloudState(current => ({ ...current, mode:'cloud', live:true, lastSync, message:'עדכון חי התקבל מ־Supabase' }))
            setStatus('הנתונים עודכנו אוטומטית וזמינים לכל המשתמשים')
          } catch (error) {
            console.warn('Realtime refresh failed', error)
            setCloudState(current => ({ ...current, live:false, message:'מחובר לענן, אך העדכון החי נכשל' }))
          }
        }, 700)
      })
      .subscribe(status => {
        if (status === 'SUBSCRIBED') setCloudState(current => ({ ...current, live:true }))
        if (status === 'CHANNEL_ERROR' || status === 'TIMED_OUT') setCloudState(current => ({ ...current, live:false }))
      })
    return () => { clearTimeout(refreshTimer); supabase.removeChannel(channel) }
  }, [])

  useEffect(() => {
    if (!production.length && !quality.length && !deviations.length && !targets.length) return
    const timer = setTimeout(() => {
      idbSet({ production, quality, deviations, targets, dataMeta, savedAt: new Date().toISOString() })
        .catch(e => { console.error(e); setStatus('הנתונים בענן, אך יצירת גיבוי מקומי נכשלה') })
    }, 800)
    return () => clearTimeout(timer)
  }, [production, quality, deviations, targets, dataMeta])

  const validateRows = (kind, rows) => {
    const sample = rows.slice(0, 250)
    const present = (...names) => sample.some(r => names.some(n => getField(r, [n]) !== ''))
    const checks = {
      production: [
        ['מתקן / Storage Location', present('Storage Location', 'Storage location')],
        ['כמות', present('Delivered quantity (GMEIN)', 'Confirmed Yield Quantity (GMEIN)', 'Delivered quantity')],
        ['Order או Batch', present('Order', 'Process Order', 'Batch', 'Batch Number')],
      ],
      quality: [
        ['Inspection Lot או Batch', present('Inspection Lot', 'Inspection Lot #', 'Batch', 'Batch Number')],
        ['מאפיין או סטטוס איכות', present('Master Insp Charactristic', 'Master Inspection Characteristic', 'Result Status', 'QA Approval')],
      ],
      deviations: [
        ['Batch', present('Batch', 'Batch Number')],
        ['סטטוס / מאפיין חריג', present('QA Status', 'Rejected characteristics', 'UD Remarks')],
      ],
      targets: [
        ['מתקן / Storage Location', present('Storage Location', 'Facility', 'מתקן')],
        ['יעד חודשי', present('Monthly Target', 'Monthly Plan', 'Plan', 'יעד חודשי', 'תוכנית חודשית', 'Target')],
      ],
    }
    return (checks[kind] || []).filter(([, ok]) => !ok).map(([label]) => label)
  }

  const loadFiles = async (files, forcedKind = '') => {
    if (!files.length) return
    setBusy(true)
    try {
      const loaded = []
      for (const file of files) {
        setStatus(`קורא את ${file.name}...`)
        const rows = await readWorkbook(file)
        const detected = classifyFile(rows)
        const kind = forcedKind || detected
        const missing = validateRows(kind, rows)
        if (missing.length) throw new Error(`${file.name}: חסרות עמודות חובה — ${missing.join(', ')}`)
        let storedCount = rows.length
        let rowsForCloud = rows
        if (kind === 'production') {
          const compact = rows.map(r => ({
            __compactProduction: true,
            facility: canonicalFacility(getField(r, ['Storage Location', 'Storage location'])),
            finishDate: combineExcelDateTime(
              getField(r, ['Actual finish date', 'Actual Finish Date']),
              getField(r, ['Actual Finish Time', 'Actual finish time']),
              getField(r, ['Release date (actual)', 'Time Stamp'])
            )?.toISOString?.() || '',
            qty: num(getField(r, ['Delivered quantity (GMEIN)', 'Confirmed Yield Quantity (GMEIN)', 'Delivered quantity'])),
            order: normalize(getField(r, ['Order', 'Process Order', 'Work Order'])),
            batch: normalize(getField(r, ['Batch', 'Batch Number'])),
            material: normalize(getField(r, ['Material', 'Material #', 'Material Number', 'Material No.', 'מקט', 'מק"ט', 'מק״ט'])),
            desc: normalize(getField(r, ['Material description', 'Material Description'])),
            orderType: normalize(getField(r, ['Order Type'])),
            routingGroup: normalizeRouting(getField(r, ['Routing group', 'Routing Group', 'RoutingGroup'])),
          })).filter(r => r.facility && (r.qty || r.order || r.batch))
          storedCount = compact.length
          rowsForCloud = compact
        }
        else if (kind === 'quality') {
          const compact = rows.map(r => ({
            __compactQuality: true,
            facility: canonicalFacility(getField(r, ['Inspection Lot Storage Location', 'Process Order Storage Location', 'Storage Location', 'Facility', 'Production Line'])),
            date: combineExcelDateTime(
              getField(r, ['Sample Date', 'Sampling Date', 'Date of Sample', 'Date of Sampling', 'תאריך דגימה', 'Start Date of Inspection', 'Date of Lot Creation', 'Process Order Confirmed Release Date', 'End Date of Inspection', 'Inspection Lot UD Date', 'Process Order Delivered Date']),
              getField(r, ['Sample Time', 'Sampling Time', 'Time of Sample', 'Time of Sampling', 'שעת דגימה', 'Inspection Time', 'Start Time of Inspection', 'Time']),
              getField(r, ['Sample Date Time', 'Sampling Date Time', 'Sample Datetime', 'Sampling Datetime', 'תאריך ושעת דגימה'])
            ),
            batch: normalize(getField(r, ['Batch', 'Batch Number'])), material: normalize(getField(r, ['Material', 'Material #', 'Material Number', 'Material No.', 'מקט', 'מק"ט', 'מק״ט'])),
            order: normalize(getField(r, ['Process Order', 'Process Order #', 'Order'])), status: normalize(getField(r, ['Result Status', 'QA Approval', 'Status'])),
            approval: normalize(getField(r, ['QA Approval'])), inspectionLot: normalize(getField(r, ['Inspection Lot', 'Inspection Lot #'])),
            characteristic: normalize(getField(r, ['Master Insp Charactristic', 'Master Inspection Characteristic'])),
            value: normalize(getField(r, ['Arithmetic Mean of Valid Measured Values'])), lower: normalize(getField(r, ['Lower Specif Limit', 'Lower Spec Limit'])),
            upper: normalize(getField(r, ['Upper Specif Limit', 'Upper Spec Limit'])), unit: normalize(getField(r, ['Unit of Measurement'])),
            line: normalize(getField(r, ['Production Line'])), remarks: normalize(getField(r, ['Charactristic Remarks', 'Characteristic Remarks', 'Batch Remarks'])),
            qualitative: normalize(getField(r, ['Qualitative'])),
          })).filter(r => r.batch || r.inspectionLot)
          storedCount = compact.length; rowsForCloud = compact
        } else if (kind === 'deviations') {}
        else if (kind === 'targets') {
          const fallbackMonth = parseMonth(file.name, new Date())
          const parsed = rows.map(r => {
            const facility = canonicalFacility(getField(r, ['Storage Location', 'Facility', 'מתקן']))
            const routingGroup = normalizeRouting(getField(r, ['Routing group', 'Routing Group', 'RoutingGroup', 'קבוצת ניתוב', 'משאב']))
            const mapped = resourceMeta(facility, routingGroup)
            return {
              facility,
              routingGroup,
              station: normalize(getField(r, ['Station', 'Resource', 'Work Center', 'תחנה'])) || mapped.station || '',
              lineName: normalize(getField(r, ['Line', 'Line Description', 'Packaging Type', 'סוג אריזה', 'קו'])) || mapped.line || '',
              month: parseMonth(getField(r, ['Month', 'חודש', 'Target Month', 'Plan Month']), fallbackMonth),
              activity: normalize(getField(r, ['Activity', 'Type', 'סוג פעילות', 'סוג', 'Production/Packaging'])) || 'אריזה',
              target: num(getField(r, ['Monthly Target', 'Monthly Plan', 'Plan', 'יעד חודשי', 'תוכנית חודשית', 'Target'])),
              capacity: num(getField(r, ['Capacity', 'קיבולת', 'Monthly Capacity', 'קיבולת חודשית'])),
              notes: normalize(getField(r, ['Notes', 'Remarks', 'הערות'])),
            }
          }).filter(r => r.facility && r.target > 0)
          if (!parsed.length) throw new Error(`${file.name}: לא נמצאו יעדים חודשיים תקינים`)
          storedCount = parsed.length; rowsForCloud = parsed; if (parsed[0]?.month) setPlanningMonth(parsed[0].month)
        } else throw new Error(`${file.name}: סוג הקובץ לא זוהה. השתמש באזור הטעינה המתאים במרכז הנתונים.`)
        const facilitiesFound = new Set(rows.slice(0, 5000).map(r => canonicalFacility(getField(r, ['Storage Location','Inspection Lot Storage Location','Process Order Storage Location','Facility','Production Line','מתקן']))).filter(Boolean)).size
        const nextMeta = { fileName:file.name, rows:storedCount, rawRows:rows.length, loadedAt:new Date().toISOString(), facilities:facilitiesFound, valid:true, source:'cloud' }
        setStatus(`מעלה את ${file.name} למסד המשותף...`)
        setUploadProgress({ fileName:file.name, kind, phase:'prepare', percent:0, message:'מכין את הנתונים' })
        const savedMeta = await uploadCloudDataset(kind, rowsForCloud, nextMeta, currentUser, progress => {
          setUploadProgress({ fileName:file.name, kind, ...progress })
          setStatus(`${file.name}: ${progress.message} (${progress.percent}%)`)
        })
        if (kind === 'production') setProduction(rowsForCloud)
        else if (kind === 'quality') setQuality(rowsForCloud)
        else if (kind === 'deviations') setDeviations(rowsForCloud)
        else if (kind === 'targets') setTargets(rowsForCloud)
        setDataMeta(current => ({ ...current, [kind]: savedMeta }))
        setCloudState({ mode:'cloud', lastSync:savedMeta.loadedAt, message:'מחובר ומסונכרן עם Supabase', latencyMs:cloudState.latencyMs, live:true })
        loaded.push(`${file.name}: ${fmt(storedCount)} רשומות בענן`)
      }
      setStatus(`הטעינה לענן הושלמה — ${loaded.join(' | ')}`)
    } catch (e) {
      console.error('Cloud upload failed', e)
      const technical = [e?.message, e?.details, e?.hint, e?.code].filter(Boolean).join(' | ')
      setStatus(`שגיאה בטעינה לענן: ${technical || 'שגיאה לא ידועה'}`)
      setCloudState(current => ({
        ...current,
        mode: current.lastSync ? 'cloud' : current.mode,
        live: current.live,
        message: `החיבור לענן פעיל, אך העלאת הקובץ נכשלה: ${technical || 'שגיאה לא ידועה'}`,
      }))
    }
    finally { setBusy(false); setTimeout(() => setUploadProgress(null), 1800) }
  }

  const handleFiles = (files) => loadFiles(files)


  const prod = useMemo(() => production.map(r => {
    if (r?.__compactProduction) {
      return {
        facility: canonicalFacility(r.facility),
        date: r.finishDate ? new Date(r.finishDate) : null,
        qty: num(r.qty),
        order: normalize(r.order),
        batch: normalize(r.batch),
        material: normalize(r.material),
        desc: normalize(r.desc),
        orderType: normalize(r.orderType),
        routingGroup: normalizeRouting(r.routingGroup),
      }
    }
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
      material: normalize(getField(r, ['Material', 'Material #', 'Material Number', 'Material No.', 'מקט', 'מק"ט', 'מק״ט'])),
      desc: normalize(getField(r, ['Material description', 'Material Description'])),
      orderType: normalize(getField(r, ['Order Type'])),
      routingGroup: normalizeRouting(getField(r, ['Routing group', 'Routing Group'])),
      routingDescription: normalize(getField(r, ['Description', 'Routing Description'])),
      hour: finish ? finish.getHours() : null,
      shift: shiftInfo(finish),
    }
  }).filter(r => r.facility), [production])

  const qualityRows = useMemo(() => quality.map(r => r.__compactQuality ? r : ({
    facility: canonicalFacility(getField(r, ['Inspection Lot Storage Location', 'Process Order Storage Location', 'Storage Location', 'Facility', 'Production Line'])),
    date: combineExcelDateTime(
      getField(r, ['Sample Date', 'Sampling Date', 'Date of Sample', 'Date of Sampling', 'תאריך דגימה', 'Start Date of Inspection', 'Date of Lot Creation', 'Process Order Confirmed Release Date', 'End Date of Inspection', 'Inspection Lot UD Date', 'Process Order Delivered Date']),
      getField(r, ['Sample Time', 'Sampling Time', 'Time of Sample', 'Time of Sampling', 'שעת דגימה', 'Inspection Time', 'Start Time of Inspection', 'Time']),
      getField(r, ['Sample Date Time', 'Sampling Date Time', 'Sample Datetime', 'Sampling Datetime', 'תאריך ושעת דגימה'])
    ),
    batch: normalize(getField(r, ['Batch', 'Batch Number'])), material: normalize(getField(r, ['Material', 'Material #', 'Material Number', 'Material No.', 'מקט', 'מק"ט', 'מק״ט'])),
    order: normalize(getField(r, ['Process Order', 'Process Order #', 'Order'])), status: normalize(getField(r, ['Result Status', 'QA Approval', 'Status'])),
    inspectionLot: normalize(getField(r, ['Inspection Lot', 'Inspection Lot #'])),
  })), [quality])

  const deviationRows = useMemo(() => deviations.map(r => ({
    facility: canonicalFacility(getField(r, ['Facility', 'Production Line', 'Storage Location'])),
    date: excelDate(getField(r, ['Date of Lot Creation', 'Inspection Lot UD Date', 'Process Order Delivered Date', 'Start Date of Inspection'])),
    batch: normalize(getField(r, ['Batch', 'Batch Number'])), material: normalize(getField(r, ['Material', 'Material #', 'Material Number', 'Material No.', 'מקט', 'מק"ט', 'מק״ט'])),
    inspectionLot: normalize(getField(r, ['Inspection Lot', 'Inspection Lot #'])),
    status: normalize(getField(r, ['QA Status', 'Status'])),
    rejectedCount: num(getField(r, ['Rejected characteristics', 'Rejected characteristics '])),
    remarks: normalize(getField(r, ['UD Remarks', 'Remarks', 'Batch Remarks'])),
    udCode: normalize(getField(r, ['UD Code'])),
  })), [deviations])

  const rejectedQualityByBatch = useMemo(() => {
    const map = new Map()
    qualityRows.forEach(r => {
      const status = normalize(r.status).toLowerCase()
      const isRejected = ['rejection', 'rejected', 'fail', 'failed', 'פסול', 'לא תקין', 'חריג'].some(x => status.includes(x))
      if (!isRejected || !r.batch || !r.characteristic) return
      const item = {
        characteristic: r.characteristic,
        value: r.value,
        lower: r.lower,
        upper: r.upper,
        unit: r.unit,
        remarks: r.remarks,
        qualitative: r.qualitative,
        inspectionLot: r.inspectionLot,
        date: r.date,
      }
      const current = map.get(r.batch) || []
      const duplicate = current.some(x => x.characteristic === item.characteristic && x.value === item.value && x.inspectionLot === item.inspectionLot)
      if (!duplicate) current.push(item)
      map.set(r.batch, current)
    })
    return map
  }, [qualityRows])

  const approvedQualityByBatch = useMemo(() => {
    const map = new Map()
    qualityRows.forEach(r => {
      const status = normalize(r.status || r.approval).toLowerCase()
      const isRejected = ['rejection', 'rejected', 'fail', 'failed', 'פסול', 'לא תקין', 'חריג'].some(x => status.includes(x))
      if (isRejected || !r.batch || !r.characteristic) return
      const item = {
        characteristic: r.characteristic,
        value: r.value,
        lower: r.lower,
        upper: r.upper,
        unit: r.unit,
        remarks: r.remarks,
        qualitative: r.qualitative,
        inspectionLot: r.inspectionLot,
        date: r.date,
      }
      const current = map.get(r.batch) || []
      const duplicate = current.some(x => x.characteristic === item.characteristic && x.value === item.value && x.inspectionLot === item.inspectionLot)
      if (!duplicate) current.push(item)
      map.set(r.batch, current)
    })
    return map
  }, [qualityRows])

  const materialByBatch = useMemo(() => {
    const map = new Map()
    prod.forEach(r => { if (r.batch && r.material && !map.has(r.batch)) map.set(r.batch, r.material) })
    qualityRows.forEach(r => { if (r.batch && r.material && !map.has(r.batch)) map.set(r.batch, r.material) })
    return map
  }, [prod, qualityRows])

  const enrichedDeviationRows = useMemo(() => deviationRows.map(r => {
    const batchQuality = qualityRows.filter(q => q.batch && q.batch === r.batch)
    const lotQuality = r.inspectionLot ? batchQuality.filter(q => q.inspectionLot === r.inspectionLot) : []
    const matchedQuality = lotQuality.length ? lotQuality : batchQuality
    const sampleDate = matchedQuality.map(q => q.date).filter(Boolean).sort((a,b) => new Date(b) - new Date(a))[0] || null
    return {
      ...r,
      sampleDate,
      material: r.material || materialByBatch.get(r.batch) || '',
      rejectedCharacteristics: rejectedQualityByBatch.get(r.batch) || [],
      approvedCharacteristics: approvedQualityByBatch.get(r.batch) || [],
    }
  }), [deviationRows, qualityRows, materialByBatch, rejectedQualityByBatch, approvedQualityByBatch])

  const batchIndex = useMemo(() => {
    const map = new Map()
    const ensure = (batch) => {
      const key = normalize(batch)
      if (!key) return null
      if (!map.has(key)) map.set(key, { batch:key, production:[], quality:[], deviations:[] })
      return map.get(key)
    }
    prod.forEach(row => { const item=ensure(row.batch); if (item) item.production.push(row) })
    qualityRows.forEach(row => { const item=ensure(row.batch); if (item) item.quality.push(row) })
    enrichedDeviationRows.forEach(row => { const item=ensure(row.batch); if (item) item.deviations.push(row) })
    return map
  }, [prod, qualityRows, enrichedDeviationRows])

  const selectedBatchData = selectedBatch ? batchIndex.get(selectedBatch) : null
  const openBatchCard = (batch) => { if (batch) setSelectedBatch(normalize(batch)) }

  const dataMonths = useMemo(() => [...new Set(prod.map(r => monthKey(r.date)).filter(Boolean))].sort(), [prod])
  const targetMonths = useMemo(() => [...new Set(targets.map(r => r.month).filter(Boolean))].sort(), [targets])
  const availableMonths = useMemo(() => [...new Set([...targetMonths, ...dataMonths])].sort().reverse(), [targetMonths, dataMonths])
  useEffect(() => { if (!planningMonth && availableMonths.length) setPlanningMonth(availableMonths[0]) }, [availableMonths, planningMonth])

  const dateBounds = useMemo(() => {
    // The date selector controls every dataset, therefore its available range
    // must include production, quality and deviations (not production only).
    const ds = [...prod, ...qualityRows, ...deviationRows]
      .map(r => r.date)
      .filter(Boolean)
      .sort((a, b) => new Date(a) - new Date(b))
    return { min: iso(ds[0]), max: iso(ds.at(-1)) }
  }, [prod, qualityRows, deviationRows])

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

    return facilities.flatMap(facility => {
      const facilityTargets = monthTargets.filter(t => t.facility === facility)
      const targetRows = facilityTargets.length ? facilityTargets : [null]
      return targetRows.map((targetRow, index) => {
        const routingGroup = normalizeRouting(targetRow?.routingGroup)
        const mapped = resourceMeta(facility, routingGroup)
        let rows = monthRows.filter(r => r.facility === facility)
        if (facility === '1542') rows = rows.filter(r => r.orderType.toUpperCase().includes('ZFIN'))
        if (routingGroup) rows = rows.filter(r => normalizeRouting(r.routingGroup) === routingGroup)
        const actual = rows.reduce((sum, r) => sum + r.qty, 0)
        const target = targetRow?.target || 0
        const dailyMap = new Map()
        rows.forEach(r => { const date = iso(r.date); if (date) dailyMap.set(date, (dailyMap.get(date) || 0) + r.qty) })
        const dailyValues = [...dailyMap.values()]
        const actualDays = dailyValues.length
        const average = actualDays ? actual / actualDays : 0
        const recent = [...dailyMap.entries()].sort((a,b) => a[0].localeCompare(b[0])).slice(-7).map(([,value]) => value)
        const recentAverage = recent.length ? recent.reduce((a,b) => a+b, 0) / recent.length : average
        const provenMax = dailyValues.length ? Math.max(...dailyValues) : (targetRow?.capacity ? targetRow.capacity / Math.max(1, totalWorkdays) : (!routingGroup ? LEGACY_DAILY_TARGETS[facility] || 0 : 0))
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
          id: `${facility}::${routingGroup || `ALL-${index}`}`,
          facility,
          routingGroup,
          station: targetRow?.station || mapped.station || '',
          lineName: targetRow?.lineName || mapped.line || rows.find(r => r.routingDescription)?.routingDescription || '',
          activity: targetRow?.activity || 'אריזה', target, capacity: targetRow?.capacity || 0, actual,
          pct: target ? actual / target * 100 : 0, remaining, requiredDaily, average, recentAverage, provenMax,
          forecast, capacityForecast, elapsedWorkdays, remainingWorkdays, totalWorkdays,
          orders: new Set(rows.map(r => r.order).filter(Boolean)).size, state, label,
        }
      })
    })
  }, [planningMonth, prod, targets, facilities])


  const facilityStats = useMemo(() => facilities.map(id => {
    let rows = baseFiltered.filter(r => r.facility === id)
    if (id === '1542') rows = rows.filter(r => r.orderType.toUpperCase().includes('ZFIN'))
    const actual = rows.reduce((s, r) => s + r.qty, 0)
    const plans = planningRows.filter(x => x.facility === id)
    const target = plans.reduce((sum, row) => sum + row.target, 0)
    const forecast = plans.reduce((sum, row) => sum + row.forecast, 0)
    const rank = { risk: 3, warning: 2, 'no-target': 1, good: 0, achieved: 0 }
    const state = plans.sort((a,b) => (rank[b.state] || 0) - (rank[a.state] || 0))[0]?.state || 'no-target'
    return { id, actual, target, pct: target ? actual / target * 100 : 0, orders: new Set(rows.map(r => r.order).filter(Boolean)).size, state, forecast }
  }), [facilities, baseFiltered, planningRows])


  const toggleFacility = (id) => setSelectedFacilities(current => current.includes(id) ? current.filter(x => x !== id) : [...current, id])
  const allFacilitiesSelected = selectedFacilities.length === facilities.length
  const toggleAllFacilities = () => setSelectedFacilities(allFacilitiesSelected ? [] : [...facilities])

  const total = filtered.reduce((s, x) => s + x.qty, 0)
  const activeFacilities = facilityStats.filter(x => x.actual > 0).length
  const morningQty = filtered.filter(r => r.shift?.key === 'morning').reduce((s, r) => s + r.qty, 0)
  const eveningQty = filtered.filter(r => r.shift?.key === 'evening').reduce((s, r) => s + r.qty, 0)
  const nightQty = filtered.filter(r => r.shift?.key === 'night').reduce((s, r) => s + r.qty, 0)

  const shiftAnalysis = useMemo(() => {
    const batchShift = new Map()
    filtered.forEach(r => { if (r.batch && r.shift?.key !== 'unknown') batchShift.set(r.batch, r.shift) })
    const rowsByShift = new Map(['morning','evening','night'].map(key => [key, []]))
    filtered.forEach(r => { if (rowsByShift.has(r.shift?.key)) rowsByShift.get(r.shift.key).push(r) })
    const labels = { morning:'בוקר', evening:'ערב', night:'לילה' }
    const hours = { morning:'07:00–15:00', evening:'15:00–23:00', night:'23:00–07:00' }
    return ['morning','evening','night'].map(key => {
      const rows = rowsByShift.get(key).slice().sort((a,b) => (a.date?.getTime?.() || 0) - (b.date?.getTime?.() || 0))
      const groups = new Map()
      rows.forEach(r => {
        const groupKey = `${r.shift.workDate}|${r.facility}|${r.routingGroup || 'NO-ROUTING'}`
        const list = groups.get(groupKey) || []; list.push(r); groups.set(groupKey, list)
      })
      const changeovers = []
      groups.forEach(list => {
        list.sort((a,b) => a.date - b.date)
        for (let i=1;i<list.length;i++) {
          const prev=list[i-1], curr=list[i]
          if (prev.material && curr.material && prev.material !== curr.material) {
            const minutes = Math.max(0, Math.round((curr.date - prev.date)/60000))
            changeovers.push({ facility:curr.facility, routingGroup:curr.routingGroup, at:curr.date, fromMaterial:prev.material, fromDesc:prev.desc, toMaterial:curr.material, toDesc:curr.desc, minutes })
          }
        }
      })
      const batches = new Set(rows.map(r=>r.batch).filter(Boolean))
      const shiftDeviations = enrichedDeviationRows.filter(d => { const st = normalize(d.status).toLowerCase(); const isOpen = !st || !['approved','closed','מאושר','סגור'].some(x => st.includes(x)); return isOpen && d.batch && batches.has(d.batch) })
      const total = rows.reduce((sum,r)=>sum+r.qty,0)
      return {
        key, label:labels[key], hours:hours[key], total,
        orders:new Set(rows.map(r=>r.order).filter(Boolean)).size,
        batches:batches.size,
        materials:new Set(rows.map(r=>r.material).filter(Boolean)).size,
        changeovers,
        avgChangeover:changeovers.length ? changeovers.reduce((sum,c)=>sum+c.minutes,0)/changeovers.length : 0,
        totalChangeover:changeovers.reduce((sum,c)=>sum+c.minutes,0),
        deviations:shiftDeviations.length,
        avgPerHour:total/8,
        share: total ? total / Math.max(1, filtered.reduce((sum,r)=>sum+r.qty,0))*100 : 0,
      }
    })
  }, [filtered, enrichedDeviationRows])
  // A single date/facility scope is shared by production, quality and deviations.
  // This prevents quality/deviation cards from continuing to show rows outside
  // the selected period (for example when choosing "יום אחרון").
  const filteredQualityRows = useMemo(() => qualityRows.filter(r =>
    (!selectedFacilities.length || selectedFacilities.includes(r.facility)) &&
    matchesDateRange(r.date, from, to)
  ), [qualityRows, selectedFacilities, from, to])

  const filteredDeviationRows = useMemo(() => enrichedDeviationRows.filter(r =>
    (!selectedFacilities.length || selectedFacilities.includes(r.facility)) &&
    matchesDateRange(r.sampleDate || r.date, from, to)
  ), [enrichedDeviationRows, selectedFacilities, from, to])

  const qualityBad = useMemo(() => filteredQualityRows.filter(r => {
    const st = normalize(r.status).toLowerCase()
    return st && !['accepted', 'תקין', 'pass', 'approved', 'מאושר'].some(x => st.includes(x))
  }).sort((a,b) => (b.date?.getTime?.() || new Date(b.date || 0).getTime()) - (a.date?.getTime?.() || new Date(a.date || 0).getTime())), [filteredQualityRows])

  const openDeviations = useMemo(() => filteredDeviationRows.filter(r => {
    const st = normalize(r.status).toLowerCase()
    return !st || !['approved', 'closed', 'מאושר', 'סגור'].some(x => st.includes(x))
  }).sort((a,b) => new Date(b.sampleDate || b.date || 0) - new Date(a.sampleDate || a.date || 0)), [filteredDeviationRows])

  const alerts = useMemo(() => planningRows.filter(r => ['risk', 'warning'].includes(r.state)).sort((a,b) => ({ risk:0, warning:1 }[a.state] - { risk:0, warning:1 }[b.state])), [planningRows])
  const achievedCount = planningRows.filter(r => ['achieved', 'good'].includes(r.state)).length
  const riskCount = planningRows.filter(r => r.state === 'risk').length
  const warningCount = planningRows.filter(r => r.state === 'warning').length
  const targetTotal = planningRows.reduce((s,r) => s + r.target, 0)
  const targetActual = planningRows.reduce((s,r) => s + r.actual, 0)
  const targetForecast = planningRows.reduce((s,r) => s + r.forecast, 0)
  const uniqueOrders = useMemo(() => new Set(filtered.map(r => r.order).filter(Boolean)).size, [filtered])
  const uniqueBatches = useMemo(() => new Set(filtered.map(r => r.batch).filter(Boolean)).size, [filtered])
  const managerInsights = useMemo(() => {
    const insights = []
    const highestRisk = planningRows.filter(r => r.state === 'risk').sort((a,b) => (b.requiredDaily - b.provenMax) - (a.requiredDaily - a.provenMax))[0]
    const warning = planningRows.filter(r => r.state === 'warning').sort((a,b) => a.forecast / Math.max(1,a.target) - b.forecast / Math.max(1,b.target))[0]
    const best = planningRows.filter(r => r.target > 0).sort((a,b) => (b.forecast / b.target) - (a.forecast / a.target))[0]
    if (highestRisk) insights.push({ state:'risk', title:`${planningName(highestRisk)} דורש פעולה`, text:`הקצב הנדרש הוא ${fmt(highestRisk.requiredDaily)} ליום לעומת שיא מוכח של ${fmt(highestRisk.provenMax)}.` })
    if (warning) insights.push({ state:'warning', title:`${planningName(warning)} נמצא בסיכון`, text:`התחזית היא ${fmt(warning.forecast)} מול יעד של ${fmt(warning.target)}. נדרש שיפור בקצב היומי.` })
    if (best && ['good','achieved'].includes(best.state)) insights.push({ state:'good', title:`${planningName(best)} מוביל`, text:`התחזית הנוכחית היא ${pctFmt(best.forecast / best.target * 100)} מהיעד החודשי.` })
    if (openDeviations.length) insights.push({ state:'risk', title:`${openDeviations.length} חריגות איכות פתוחות`, text:'לחיצה על הכרטיס תפתח את פירוט המנות ומאפייני החריגה.' })
    if (!insights.length) insights.push({ state:'good', title:'אין מוקדים חריגים', text:'לפי הנתונים הטעונים, המתקנים נמצאים במסלול תקין ואין חריגות פתוחות.' })
    return insights.slice(0,4)
  }, [planningRows, openDeviations.length])


  const controlTowerFacilities = useMemo(() => {
    const deviationByFacility = new Map()
    openDeviations.forEach(row => {
      const facility = canonicalFacility(row.facility)
      if (facility) deviationByFacility.set(facility, (deviationByFacility.get(facility) || 0) + 1)
    })
    return facilities.map(facility => {
      const rows = planningRows.filter(row => row.facility === facility)
      const target = rows.reduce((sum, row) => sum + row.target, 0)
      const actual = rows.reduce((sum, row) => sum + row.actual, 0)
      const forecast = rows.reduce((sum, row) => sum + row.forecast, 0)
      const requiredDaily = rows.reduce((sum, row) => sum + row.requiredDaily, 0)
      const orders = rows.reduce((sum, row) => sum + row.orders, 0)
      const deviationsCount = deviationByFacility.get(facility) || 0
      const forecastPct = target ? forecast / target * 100 : 0
      const actualPct = target ? actual / target * 100 : 0
      const planScore = target ? Math.min(100, forecastPct) : 70
      const qualityScore = Math.max(0, 100 - deviationsCount * 8)
      const healthScore = Math.max(0, Math.min(100, Math.round(planScore * 0.72 + qualityScore * 0.28)))
      const state = !target ? 'no-target' : forecastPct >= 100 ? 'good' : forecastPct >= 90 ? 'warning' : 'risk'
      return { facility, target, actual, forecast, requiredDaily, orders, deviationsCount, forecastPct, actualPct, healthScore, state, gap: forecast - target }
    }).filter(row => row.target > 0 || row.actual > 0 || row.deviationsCount > 0)
      .sort((a,b) => ({risk:0,warning:1,good:2,'no-target':3}[a.state] - {risk:0,warning:1,good:2,'no-target':3}[b.state]) || a.facility.localeCompare(b.facility))
  }, [facilities, planningRows, openDeviations])

  const controlTowerTrend = useMemo(() => {
    const byDay = new Map()
    prod.filter(row => monthKey(row.date) === planningMonth).forEach(row => {
      const key = iso(row.date)
      if (key) byDay.set(key, (byDay.get(key) || 0) + row.qty)
    })
    return [...byDay.entries()].sort((a,b) => a[0].localeCompare(b[0])).slice(-7)
  }, [prod, planningMonth])

  const jumpToDetails = (tab) => {
    setActiveTab(tab)
    window.setTimeout(() => document.getElementById('details-section')?.scrollIntoView({ behavior:'smooth', block:'start' }), 50)
  }

  const exportWorkbook = () => {
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filtered.map(r => ({ Date: iso(r.date), Time: r.date ? r.date.toLocaleTimeString('he-IL', {hour:'2-digit',minute:'2-digit'}) : '', Facility: r.facility, Order: r.order, Batch: r.batch, Material: r.material, Description: r.desc, RoutingGroup: r.routingGroup, RoutingDescription: r.routingDescription, Quantity: r.qty }))), 'Production')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(planningRows.map(r => ({ Month: planningMonth, Facility: r.facility, RoutingGroup: r.routingGroup, Station: r.station, Line: r.lineName, Activity: r.activity, MonthlyTarget: r.target, Actual: r.actual, Remaining: r.remaining, RequiredDaily: r.requiredDaily, RecentAverage: r.recentAverage, ProvenMax: r.provenMax, Forecast: r.forecast, Status: r.label }))), 'Planning')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(qualityBad.map(r => ({ Date: iso(r.date), Facility: r.facility, InspectionLot: r.inspectionLot, Order: r.order, Batch: r.batch, Material: r.material, Status: r.status }))), 'Quality')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(openDeviations.map(r => ({ Date: iso(r.date), Facility: r.facility, Batch: r.batch, Material: r.material, Status: r.status, RejectedCount: r.rejectedCount, RejectedCharacteristics: r.rejectedCharacteristics.map(x => `${x.characteristic}: ${x.value || x.qualitative || '-'} (${[x.lower, x.upper].filter(v => v !== '').join('–')} ${x.unit || ''})`).join(' | '), ApprovedCharacteristics: r.approvedCharacteristics.map(x => `${x.characteristic}: ${x.value || x.qualitative || '-'} (${[x.lower, x.upper].filter(v => v !== '').join('–')} ${x.unit || ''})`).join(' | '), Remarks: r.remarks }))), 'Deviations')
    XLSX.writeFile(wb, `IML_Sprint8_Resource_Report_${new Date().toISOString().slice(0,10)}.xlsx`)
  }
  const downloadTargetTemplate = () => {
    const month = planningMonth || monthKey(new Date())
    const rows = DEFAULT_FACILITIES.flatMap(facility => facility === '1542'
      ? [
          { 'חודש': month, 'Storage Location': facility, 'Routing group': 'LQ-P-1', 'תחנה': 'P-02', 'סוג אריזה': '1 ליטר', 'סוג פעילות': 'אריזה', 'יעד חודשי': '', 'קיבולת חודשית': '', 'שיטת תחזית': 'ממוצע 7 ימי עבודה אחרונים', 'ימי עבודה בחודש': '', 'הערות': '' },
          { 'חודש': month, 'Storage Location': facility, 'Routing group': 'LQ-P-5', 'תחנה': 'P-03', 'סוג אריזה': '5 ליטר', 'סוג פעילות': 'אריזה', 'יעד חודשי': '', 'קיבולת חודשית': '', 'שיטת תחזית': 'ממוצע 7 ימי עבודה אחרונים', 'ימי עבודה בחודש': '', 'הערות': '' },
          { 'חודש': month, 'Storage Location': facility, 'Routing group': 'LQ-P-10', 'תחנה': 'P-04', 'סוג אריזה': '10/20 ליטר', 'סוג פעילות': 'אריזה', 'יעד חודשי': '', 'קיבולת חודשית': '', 'שיטת תחזית': 'ממוצע 7 ימי עבודה אחרונים', 'ימי עבודה בחודש': '', 'הערות': '' },
        ]
      : [{ 'חודש': month, 'Storage Location': facility, 'Routing group': '', 'תחנה': '', 'סוג אריזה': '', 'סוג פעילות': 'אריזה', 'יעד חודשי': '', 'קיבולת חודשית': '', 'שיטת תחזית': 'ממוצע 7 ימי עבודה אחרונים', 'ימי עבודה בחודש': '', 'הערות': '' }])
    const wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'יעדים חודשיים'); XLSX.writeFile(wb, 'IML_Monthly_Targets_Template.xlsx')
  }
  const clearAllData = async () => {
    if (!window.confirm('למחוק את כל הנתונים המשותפים מהענן? הפעולה תשפיע על כל המשתמשים.')) return
    setBusy(true)
    try {
      await deleteAllCloudDatasets(currentUser)
      setProduction([]); setQuality([]); setDeviations([]); setTargets([]); setDataMeta({ production:null, quality:null, deviations:null, targets:null }); localStorage.removeItem(STORAGE_KEY); await idbClear().catch(console.warn)
      setStatus('כל הנתונים נמחקו מהענן'); setCloudState({ mode:'cloud', lastSync:new Date().toISOString(), message:'מחובר למסד המשותף', latencyMs:cloudState.latencyMs, live:true }); setFrom(''); setTo(''); setQuery(''); setSelectedFacilities([]); setPlanningMonth(''); setAdditionalFacilities([]); setFacilityToAdd(''); setPeriodYear(''); setPeriodQuarter('')
    } catch (error) { setStatus(`מחיקת הנתונים מהענן נכשלה: ${error.message}`) }
    finally { setBusy(false); setTimeout(() => setUploadProgress(null), 1800) }
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
    <aside className="side filter-side">
      <div className="brand branded"><img src="/icons/mark-128.png" alt="IML"/><div>IML<span>CONTROL</span></div></div>
      <div className="side-filter-title"><Search size={18}/><strong>חיפוש וסינון</strong></div>
      <label className="side-field"><span>Quick Search</span><div><Search size={16}/><input value={query} onChange={e => setQuery(e.target.value)} placeholder="Order, Batch, חומר..."/></div></label>
      <label className="side-field"><span>חודש תכנון</span><select value={planningMonth} onChange={e => setPlanningMonth(e.target.value)}>{!availableMonths.length && <option value="">אין נתונים</option>}{availableMonths.map(m => <option key={m} value={m}>{m}</option>)}</select></label>
      <label className="side-field"><span>מתאריך</span><input type="date" min={dateBounds.min} max={dateBounds.max} value={from} onChange={e => { setFrom(e.target.value); setPeriodYear(''); setPeriodQuarter('') }}/></label>
      <label className="side-field"><span>עד תאריך</span><input type="date" min={dateBounds.min} max={dateBounds.max} value={to} onChange={e => { setTo(e.target.value); setPeriodYear(''); setPeriodQuarter('') }}/></label>
      <label className="side-field"><span>מתקן</span><select value={selectedFacilities.length === 1 ? selectedFacilities[0] : ''} onChange={e => setSelectedFacilities(e.target.value ? [e.target.value] : [])}><option value="">כל המתקנים</option>{facilities.map(id => <option key={id} value={id}>{id}</option>)}</select></label>
      <div className="side-quick-ranges"><button onClick={() => setQuickRange(1)}>יום</button><button onClick={() => setQuickRange(2)}>יומיים</button><button onClick={() => setQuickRange(30)}>30 יום</button></div>
      <button className="side-clear" onClick={() => { setFrom(''); setTo(''); setQuery(''); setSelectedFacilities([]); setPeriodYear(''); setPeriodQuarter('') }}><X size={16}/> ניקוי מסננים</button>
      <div className="side-live-stats"><div><Database/><span><b>{fmt(production.length)}</b><small>תפוקה</small></span></div><div><FlaskConical/><span><b>{fmt(quality.length + deviations.length)}</b><small>איכות</small></span></div></div>
      <div className="side-note">Sprint 11.2.5 · {userRole === 'admin' ? 'Admin' : userRole === 'manager' ? 'Manager' : 'Viewer'}</div>
    </aside>

    <main className="main">
      <header className="header">
        <div><h1>חדר בקרה — מתקני אריזה</h1><p>Sprint 11.2.5 — Control Tower & Roles</p></div>
        <div className="header-actions">
          <div className="user-session"><img className="user-brand-avatar" src="/icons/mark-64.png" alt="IML"/><span><b>{isGuest ? 'אורח' : (currentUser?.email || 'משתמש')}</b><small>{isGuest ? 'צפייה בלבד' : userRole === 'admin' ? 'מנהל מערכת' : userRole === 'manager' ? 'מנהל מתקן' : 'צפייה בלבד'}</small></span></div>
          <button className="action secondary" onClick={downloadTargetTemplate}><FileSpreadsheet size={18}/> תבנית יעדים</button>
          <button className="action secondary" onClick={exportWorkbook} disabled={!production.length}><Download size={18}/> יצוא Excel</button>
          {canDeleteData && <button className="action danger" onClick={clearAllData} disabled={!production.length && !quality.length && !deviations.length && !targets.length}><Trash2 size={18}/> מחיקה</button>}
          {canManageData && <label className={`upload ${busy ? 'disabled' : ''}`}><Upload size={19}/>{busy ? 'טוען...' : 'טעינת Excel'}<input type="file" multiple accept=".xlsx,.xls" disabled={busy} onChange={e => handleFiles([...e.target.files])}/></label>}
          <button className="action secondary" onClick={onSignOut}><LogOut size={18}/> יציאה</button>
        </div>
      </header>

      <div className="load-status"><CheckCircle2 size={18}/>{status}</div>
      {uploadProgress && <section className="upload-progress-card"><div className="upload-progress-head"><strong>{uploadProgress.fileName}</strong><span>{uploadProgress.percent}%</span></div><div className="upload-progress-track"><div style={{width:`${uploadProgress.percent}%`}}/></div><small>{uploadProgress.message}</small></section>}

      <section className={`cloud-status ${cloudState.mode}`}>
        <div className="cloud-status-icon">{cloudState.mode === 'offline' ? <WifiOff/> : <Cloud/>}</div>
        <div><strong>{cloudState.mode === 'cloud' ? 'מערכת פעילה בענן' : cloudState.mode === 'connecting' ? 'מתחבר לענן' : cloudState.mode === 'offline' ? 'מצב מקומי זמני' : 'שגיאת סנכרון'}</strong><span>{cloudState.message}</span></div>
        <div className="cloud-status-meta"><small>מקור נתונים</small><b>{cloudState.mode === 'cloud' ? 'Supabase' : 'Browser Cache'}</b><small>{cloudState.live ? '● עדכון חי פעיל' : '○ עדכון חי לא פעיל'}</small>{cloudState.latencyMs != null && <small>זמן תגובה: {cloudState.latencyMs}ms</small>}{cloudState.lastSync && <small>סנכרון: {new Date(cloudState.lastSync).toLocaleString('he-IL')}</small>}</div>
      </section>


      <section className="control-tower-hero">
        <div className="control-tower-title"><div><span className="tower-kicker">PACKAGING CONTROL TOWER</span><h2>תמונת מצב ניהולית בזמן אמת</h2><p>יעדים, תחזית חודשית, איכות והזמנות — במבט אחד</p></div><div className={`tower-online ${cloudState.mode === 'cloud' ? 'online' : ''}`}><span></span>{cloudState.mode === 'cloud' ? 'ONLINE' : 'OFFLINE'}</div></div>
        <div className="tower-kpis">
          <button onClick={() => document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'})}><Target/><span>יעד חודשי כולל</span><b>{fmt(targetTotal)}</b><small>{planningMonth || 'ללא חודש נבחר'}</small></button>
          <button onClick={() => document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'})}><Database/><span>בוצע החודש</span><b>{fmt(targetActual)}</b><small>{targetTotal ? pctFmt(targetActual / targetTotal * 100) : '—'} מהיעד</small></button>
          <button onClick={() => document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'})}><TrendingUp/><span>תחזית סוף חודש</span><b>{fmt(targetForecast)}</b><small>{targetTotal ? pctFmt(targetForecast / targetTotal * 100) : '—'} מהיעד</small></button>
          <button onClick={() => document.getElementById('alerts-section')?.scrollIntoView({behavior:'smooth'})}><Gauge/><span>קצב יומי נדרש</span><b>{fmt(planningRows.reduce((sum,row)=>sum+row.requiredDaily,0))}</b><small>לכל המתקנים</small></button>
          <button className={targetForecast >= targetTotal && targetTotal ? 'good' : 'bad'} onClick={() => document.getElementById('alerts-section')?.scrollIntoView({behavior:'smooth'})}><AlertTriangle/><span>פער צפוי</span><b>{targetTotal ? fmt(targetForecast-targetTotal) : '—'}</b><small>{targetForecast >= targetTotal ? 'מעל היעד' : 'נדרש להגביר קצב'}</small></button>
        </div>
      </section>

      <section className="tower-facility-section">
        <div className="panel-head"><div><Factory/><h2>סקירת מתקנים</h2></div><span>{controlTowerFacilities.length} מתקנים במעקב</span></div>
        <div className="tower-facility-grid">
          {controlTowerFacilities.map(row => <button key={row.facility} className={`tower-facility-card ${row.state}`} onClick={() => { setSelectedFacilities([row.facility]); document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'}) }}>
            <div className="tower-facility-head"><div><i></i><strong>{row.facility}</strong></div><span>{row.state === 'good' ? 'תקין' : row.state === 'warning' ? 'דורש תשומת לב' : row.state === 'risk' ? 'בסיכון' : 'ללא יעד'}</span></div>
            <div className="tower-health"><div><HeartPulse/><span>Health Score</span></div><b>{row.healthScore}<small>/100</small></b></div>
            <div className="tower-progress"><i style={{width:`${Math.min(100,row.actualPct)}%`}}/></div>
            <dl><div><dt>יעד חודשי</dt><dd>{fmt(row.target)}</dd></div><div><dt>בוצע</dt><dd>{fmt(row.actual)}</dd></div><div><dt>תחזית</dt><dd>{fmt(row.forecast)}</dd></div><div><dt>פער צפוי</dt><dd className={row.gap >= 0 ? 'positive' : 'negative'}>{row.gap >= 0 ? '+' : ''}{fmt(row.gap)}</dd></div><div><dt>קצב נדרש</dt><dd>{fmt(row.requiredDaily)}</dd></div><div><dt>חריגות פתוחות</dt><dd>{row.deviationsCount}</dd></div></dl>
            <span className="tower-enter">לפרטים מלאים <ArrowLeft size={16}/></span>
          </button>)}
          {!controlTowerFacilities.length && <div className="empty wide-empty">טען יעדים חודשיים ונתוני תפוקה כדי להפעיל את חדר הבקרה.</div>}
        </div>
      </section>

      <section className="tower-lower-grid">
        <article className="tower-alerts-card"><div className="panel-head"><div><BellRing/><h2>התראות אחרונות</h2></div><button onClick={() => document.getElementById('alerts-section')?.scrollIntoView({behavior:'smooth'})}>הצג הכול</button></div><div className="tower-alert-list">{managerInsights.map((item,index)=><button key={index} className={item.state} onClick={() => item.title.includes('חריגות') ? jumpToDetails('deviations') : document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'})}><i>{item.state==='risk'?'!':item.state==='warning'?'⚠':'✓'}</i><div><strong>{item.title}</strong><span>{item.text}</span></div></button>)}</div></article>
        <article className="tower-chart-card"><div className="panel-head"><div><BarChart3/><h2>ביצוע יומי — 7 ימים אחרונים</h2></div><span>{planningMonth}</span></div><div className="tower-mini-chart">{controlTowerTrend.map(([day,value])=>{const max=Math.max(1,...controlTowerTrend.map(x=>x[1])); return <div key={day}><b>{fmt(value)}</b><span><i style={{height:`${Math.max(8,value/max*100)}%`}}/></span><small>{day.slice(5)}</small></div>})}{!controlTowerTrend.length&&<div className="empty">אין נתונים להצגת מגמה</div>}</div></article>
      </section>

      <section className="data-center">
        <div className="panel-head"><div><ShieldCheck/><h2>מרכז נתונים</h2></div><span>4 מקורות מידע</span></div>
        <p className="data-center-help">כל קובץ נבדק בדפדפן ולאחר מכן נשמר ב־Supabase. מרגע שהטעינה מסתיימת, אותו מידע זמין לכל המשתמשים המחוברים.</p>
        <div className="data-source-grid">
          <DataSource title="תפוקות" icon={<Factory/>} meta={dataMeta.production} count={production.length} acceptLabel="טען קובץ תפוקות" busy={busy} onFiles={files => loadFiles(files, 'production')} canManage={canManageData}/>
          <DataSource title="תוצאות איכות" icon={<FlaskConical/>} meta={dataMeta.quality} count={quality.length} acceptLabel="טען תוצאות איכות" busy={busy} onFiles={files => loadFiles(files, 'quality')} canManage={canManageData}/>
          <DataSource title="חריגות איכות" icon={<AlertTriangle/>} meta={dataMeta.deviations} count={deviations.length} acceptLabel="טען קובץ חריגות" busy={busy} onFiles={files => loadFiles(files, 'deviations')} canManage={canManageData}/>
          <DataSource title="יעדים חודשיים" icon={<Target/>} meta={dataMeta.targets} count={targets.length} acceptLabel="טען קובץ יעדים" busy={busy} onFiles={files => loadFiles(files, 'targets')} canManage={canManageData}/>
        </div>
      </section>

      <section className="planning-toolbar">
        <div><Target/><span>חודש תכנון</span><select value={planningMonth} onChange={e => setPlanningMonth(e.target.value)}>{!availableMonths.length && <option value="">אין נתונים</option>}{availableMonths.map(m => <option key={m} value={m}>{m}</option>)}</select></div>
        <small>היעד היומי אינו מוזן: הוא מחושב מחדש בכל יום לפי יתרת היעד וימי העבודה שנותרו.</small>
      </section>

      <section className="executive-strip executive-six">
        <Executive icon={<Target/>} title="ביצוע מול יעד" value={targetTotal ? pctFmt(targetActual / targetTotal * 100) : '—'} sub={`${fmt(targetActual)} מתוך ${fmt(targetTotal)}`} onClick={() => document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'})}/>
        <Executive icon={<TrendingUp/>} title="תחזית סוף חודש" value={targetTotal ? pctFmt(targetForecast / targetTotal * 100) : '—'} sub={fmt(targetForecast)} onClick={() => document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'})}/>
        <Executive icon={<CheckCircle2/>} title="במסלול / הושג" value={achievedCount} sub="מתקנים" good onClick={() => document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'})}/>
        <Executive icon={<AlertTriangle/>} title="בסיכון" value={warningCount} sub="נדרש שיפור" warn onClick={() => document.getElementById('alerts-section')?.scrollIntoView({behavior:'smooth'})}/>
        <Executive icon={<BellRing/>} title="חריגות פתוחות" value={openDeviations.length} sub="מנות איכות" bad onClick={() => jumpToDetails('deviations')}/>
        <Executive icon={<ClipboardList/>} title="Orders בטווח" value={uniqueOrders} sub={`${uniqueBatches} מנות`} onClick={() => jumpToDetails('production')}/>
      </section>

      <section className="manager-brief">
        <div className="panel-head"><div><Activity/><h2>תמונת מצב ניהולית</h2></div><span>מבוסס על הנתונים הטעונים</span></div>
        <div className="manager-insights">{managerInsights.map((item,i) => <button key={i} className={`manager-insight ${item.state}`} onClick={() => item.title.includes('חריגות') ? jumpToDetails('deviations') : document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'})}><strong>{item.title}</strong><span>{item.text}</span></button>)}</div>
      </section>

      <section className="filters legacy-filters">
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
        <Summary title="משמרת בוקר" value={fmt(morningQty)} sub="07:00–15:00"/>
        <Summary title="משמרת ערב" value={fmt(eveningQty)} sub="15:00–23:00"/>
        <Summary title="משמרת לילה" value={fmt(nightQty)} sub="23:00–07:00"/>
        <Summary title="חריגות פתוחות" value={openDeviations.length} sub="לפי קובץ החריגות" warn/>
        <Summary title="איכות לא תקינה" value={qualityBad.length} sub="לפי קובץ האיכות" warn/>
      </section>

      <div className="section-title facility-title"><div className="section-title-text"><Gauge/><div><h2>תחזית חודשית לפי מתקן</h2><p>יעד חודשי, קצב נדרש, קצב אחרון, שיא מוכח ותחזית</p></div></div><button className="select-all-facilities" onClick={toggleAllFacilities}><CheckCircle2 size={17}/>{allFacilitiesSelected ? 'ביטול בחירת הכול' : 'בחירת כל המתקנים'}</button></div>
      <section className="forecast-grid" id="planning-section">
        {planningRows.map(row => <ForecastCard key={row.id} {...row} selected={selectedFacilities.includes(row.facility)} onClick={() => toggleFacility(row.facility)}/>) }
        {!planningRows.length && <div className="empty wide-empty">טען קובץ יעדים חודשי כדי להציג תחזית.</div>}
      </section>

      <section className="alert-panel" id="alerts-section">
        <div className="panel-head"><div><BellRing/><h2>מה דורש תשומת לב היום?</h2></div><span>{alerts.length} התראות</span></div>
        <div className="alert-list">
          {alerts.map(r => <div className={`alert-item ${r.state}`} key={r.id}><div className="alert-symbol">{r.state === 'risk' ? '!' : '⚠'}</div><div><strong>{planningName(r)} — {r.label}</strong><p>{r.state === 'risk' ? `נדרש ${fmt(r.requiredDaily)} ליום, אך השיא המוכח הוא ${fmt(r.provenMax)}.` : `התחזית היא ${fmt(r.forecast)} מול יעד ${fmt(r.target)}. נדרש קצב של ${fmt(r.requiredDaily)} ליום.`}</p></div></div>)}
          {!alerts.length && <div className="empty">אין התראות תכנון לחודש הנבחר.</div>}
        </div>
      </section>

      <section className="daily-management">
        <div className="panel-head"><div><CalendarCheck/><h2>Daily Management</h2></div><span>{planningMonth}</span></div>
        <div className="table-wrap"><table><thead><tr><th>מתקן</th><th>Routing group</th><th>תחנה / קו</th><th>פעילות</th><th>יעד חודשי</th><th>בפועל</th><th>% ביצוע</th><th>נותר</th><th>ימי עבודה נותרו</th><th>נדרש ליום</th><th>ממוצע 7 ימים</th><th>שיא מוכח</th><th>תחזית</th><th>סטטוס</th></tr></thead><tbody>
          {planningRows.map(r => <tr key={r.id}><td><b>{r.facility}</b></td><td>{r.routingGroup || 'כל המתקן'}</td><td>{[r.station, r.lineName].filter(Boolean).join(' · ') || '—'}</td><td>{r.activity}</td><td>{fmt(r.target)}</td><td>{fmt(r.actual)}</td><td>{pctFmt(r.pct)}</td><td>{fmt(r.remaining)}</td><td>{r.remainingWorkdays}</td><td>{fmt(r.requiredDaily)}</td><td>{fmt(r.recentAverage)}</td><td>{fmt(r.provenMax)}</td><td>{fmt(r.forecast)}</td><td><StatusBadge state={r.state} label={r.label}/></td></tr>)}
          {!planningRows.length && <tr><td colSpan="14" className="empty">אין יעדים לחודש הנבחר</td></tr>}
        </tbody></table></div>
      </section>

      <section className="trend-card">
        <div className="trend-head"><div><h2>מגמת תפוקה יומית</h2><p>14 הימים האחרונים בטווח הנבחר</p></div><Save size={20}/></div>
        <div className="trend-bars">{dailyTrend.map(([date, value]) => <div className="trend-item" key={date} title={`${date}: ${fmt(value)}`}><div className="trend-value">{fmt(value)}</div><div className="trend-track"><i style={{height: `${Math.max(5, value / maxDaily * 100)}%`}}/></div><small>{date.slice(5)}</small></div>)}{!dailyTrend.length && <div className="empty trend-empty">טען קובץ תפוקות להצגת מגמה</div>}</div>
      </section>

      <div className="section-title facility-title"><div className="section-title-text"><Factory/><div><h2>ביצועים לפי מתקן בטווח המסונן</h2><p>לחיצה על כרטיס מסננת תפוקה, איכות וחריגות</p></div></div></div>
      <section className="facility-grid">{facilityStats.map(x => <Facility key={x.id} {...x} selected={selectedFacilities.includes(x.id)} onClick={() => toggleFacility(x.id)}/>)}</section>

      <section className="tabs" id="details-section">
        <button className={activeTab === 'production' ? 'active' : ''} onClick={() => setActiveTab('production')}><BarChart3 size={16}/> תפוקה</button>
        <button className={activeTab === 'shifts' ? 'active' : ''} onClick={() => setActiveTab('shifts')}><Clock3 size={16}/> ניתוח משמרות</button>
        <button className={activeTab === 'quality' ? 'active' : ''} onClick={() => setActiveTab('quality')}><FlaskConical size={16}/> איכות ({qualityBad.length})</button>
        <button className={activeTab === 'deviations' ? 'active' : ''} onClick={() => setActiveTab('deviations')}><AlertTriangle size={16}/> מנות חריגות ({openDeviations.length})</button>
      </section>
      {activeTab === 'production' && <section className="details"><h2>רשומות תפוקה אחרונות</h2><div className="table-wrap"><table><thead><tr><th>תאריך</th><th>שעה</th><th>מתקן</th><th>Routing group</th><th>הזמנה</th><th>Batch</th><th>מק״ט חומר</th><th>תיאור חומר</th><th>כמות</th></tr></thead><tbody>{filtered.slice(-200).reverse().map((r, i) => <tr key={i}><td>{iso(r.date)}</td><td>{r.date ? r.date.toLocaleTimeString('he-IL', {hour:'2-digit',minute:'2-digit'}) : ''}</td><td>{r.facility}</td><td>{r.routingGroup || '—'}</td><td>{r.order}</td><td>{r.batch ? <button type="button" className="batch-link" onClick={() => openBatchCard(r.batch)}>{r.batch}</button> : '—'}</td><td>{r.material || '—'}</td><td>{r.desc || '—'}</td><td>{fmt(r.qty)}</td></tr>)}{!filtered.length && <tr><td colSpan="9" className="empty">טען קובץ תפוקות כדי להציג נתונים</td></tr>}</tbody></table></div></section>}
      {activeTab === 'shifts' && <section className="details shift-intelligence"><h2>ניתוח משמרות — בוקר, ערב ולילה</h2><p className="details-note">החלפות מוצר מזוהות לפי שינוי מק״ט באותו מתקן ו-Routing group. זמן המעבר הוא פער זמן משוער בין שני דיווחים עוקבים.</p><div className="shift-card-grid">{shiftAnalysis.map(item => <article className={`shift-analysis-card shift-${item.key}`} key={item.key}><div className="shift-card-head"><div><h3>{item.label}</h3><span>{item.hours}</span></div><strong>{fmt(item.total)}</strong></div><div className="shift-metrics"><div><span>קצב ממוצע לשעה</span><b>{fmt(item.avgPerHour)}</b></div><div><span>Orders</span><b>{item.orders}</b></div><div><span>Batch</span><b>{item.batches}</b></div><div><span>מק״טים</span><b>{item.materials}</b></div><div><span>החלפות מוצר</span><b>{item.changeovers.length}</b></div><div><span>חריגות איכות</span><b>{item.deviations}</b></div><div><span>ממוצע מעבר</span><b>{fmt(item.avgChangeover)} דק׳</b></div><div><span>תרומה לתפוקה</span><b>{pctFmt(item.share)}</b></div></div></article>)}</div><h3 className="shift-subtitle">פירוט החלפות מוצר</h3><div className="table-wrap"><table><thead><tr><th>משמרת</th><th>שעה</th><th>מתקן</th><th>Routing group</th><th>מוצר קודם</th><th>מוצר חדש</th><th>פער דיווח משוער</th></tr></thead><tbody>{shiftAnalysis.flatMap(item => item.changeovers.map((c,i) => <tr key={`${item.key}-${i}`}><td>{item.label}</td><td>{c.at?.toLocaleTimeString('he-IL',{hour:'2-digit',minute:'2-digit'})}</td><td>{c.facility}</td><td>{c.routingGroup || '—'}</td><td>{c.fromMaterial}{c.fromDesc ? ` · ${c.fromDesc}` : ''}</td><td>{c.toMaterial}{c.toDesc ? ` · ${c.toDesc}` : ''}</td><td>{fmt(c.minutes)} דקות</td></tr>))}{!shiftAnalysis.some(item => item.changeovers.length) && <tr><td colSpan="7" className="empty">לא זוהו החלפות מוצר בטווח שנבחר</td></tr>}</tbody></table></div></section>}
      {activeTab === 'quality' && <section className="details"><h2>תוצאות איכות לא תקינות</h2><div className="table-wrap"><table><thead><tr><th>תאריך דגימה</th><th>שעת דגימה</th><th>מתקן</th><th>Inspection Lot</th><th>Order</th><th>Batch</th><th>מק״ט חומר</th><th>סטטוס</th></tr></thead><tbody>{qualityBad.slice(0,300).map((r,i) => <tr key={i}><td>{iso(r.date)}</td><td>{r.date ? new Date(r.date).toLocaleTimeString('he-IL',{hour:'2-digit',minute:'2-digit'}) : '—'}</td><td>{r.facility}</td><td>{r.inspectionLot}</td><td>{r.order}</td><td>{r.batch ? <button type="button" className="batch-link" onClick={() => openBatchCard(r.batch)}>{r.batch}</button> : '—'}</td><td>{r.material}</td><td><span className="status-bad">{r.status || 'ללא סטטוס'}</span></td></tr>)}{!qualityBad.length && <tr><td colSpan="8" className="empty">לא נמצאו תוצאות איכות לא תקינות</td></tr>}</tbody></table></div></section>}
      {activeTab === 'deviations' && <section className="details"><h2>מנות חריגות פתוחות</h2><p className="details-note">לכל מנה מוצגים מאפייני החריגה ולצדם המאפיינים התקינים שנמשכו מקובץ תוצאות האיכות לפי Batch.</p><div className="table-wrap"><table><thead><tr><th>תאריך חריגה</th><th>תאריך דגימה</th><th>שעת דגימה</th><th>מתקן</th><th>Batch</th><th>מק״ט חומר</th><th>סטטוס</th><th>מאפייני החריגה</th><th>מאפיינים תקינים</th><th>הערות</th></tr></thead><tbody>{openDeviations.slice(0,300).map((r,i) => <tr key={i}><td>{iso(r.date)}</td><td>{iso(r.sampleDate) || '—'}</td><td>{r.sampleDate ? new Date(r.sampleDate).toLocaleTimeString('he-IL',{hour:'2-digit',minute:'2-digit'}) : '—'}</td><td>{r.facility}</td><td>{r.batch ? <button type="button" className="batch-link" onClick={() => openBatchCard(r.batch)}>{r.batch}</button> : '—'}</td><td>{r.material}</td><td><span className="status-bad">{r.status || 'פתוח'}</span>{r.udCode && <small className="ud-code">{r.udCode}</small>}</td><td className="deviation-characteristics"><div className="characteristics-count bad-count">{r.rejectedCharacteristics.length} חריגים</div>{r.rejectedCharacteristics.length ? r.rejectedCharacteristics.map((c,j) => <div className="deviation-characteristic" key={`${c.characteristic}-${j}`}><strong>{c.characteristic}</strong><span>תוצאה: <b>{c.value || c.qualitative || '—'}{c.unit ? ` ${c.unit}` : ''}</b></span><span>מפרט: {c.lower !== '' || c.upper !== '' ? `${c.lower || '—'} עד ${c.upper || '—'}${c.unit ? ` ${c.unit}` : ''}` : '—'}</span>{c.remarks && c.remarks !== 'N/A' && <small>{c.remarks}</small>}</div>) : <span className="no-characteristics">לא נמצאו פרטי מאפיינים חריגים בקובץ האיכות{r.rejectedCount ? ` (בקובץ החריגות מופיע מספר: ${r.rejectedCount})` : ''}</span>}</td><td className="deviation-characteristics valid-characteristics"><div className="characteristics-count good-count">{r.approvedCharacteristics.length} תקינים</div>{r.approvedCharacteristics.length ? r.approvedCharacteristics.map((c,j) => <div className="deviation-characteristic valid-characteristic" key={`${c.characteristic}-${j}`}><strong>{c.characteristic}</strong><span>תוצאה: <b>{c.value || c.qualitative || '—'}{c.unit ? ` ${c.unit}` : ''}</b></span><span>מפרט: {c.lower !== '' || c.upper !== '' ? `${c.lower || '—'} עד ${c.upper || '—'}${c.unit ? ` ${c.unit}` : ''}` : '—'}</span>{c.remarks && c.remarks !== 'N/A' && <small>{c.remarks}</small>}</div>) : <span className="no-characteristics">לא נמצאו מאפיינים תקינים למנה בקובץ האיכות</span>}</td><td>{r.remarks || '—'}</td></tr>)}{!openDeviations.length && <tr><td colSpan="10" className="empty">לא נמצאו מנות חריגות פתוחות</td></tr>}</tbody></table></div></section>}
    </main>
    {selectedBatchData && <BatchControlCard data={selectedBatchData} onClose={() => setSelectedBatch('')}/>}
  </div>
}

function BatchControlCard({ data, onClose }) {
  const [qualityFilter, setQualityFilter] = useState('all')
  const productionRows = data.production || []
  const qualityRows = data.quality || []
  const deviationRows = data.deviations || []
  const firstProduction = productionRows[0] || {}
  const firstQuality = qualityRows[0] || {}
  const facilities = [...new Set([...productionRows.map(r=>r.facility), ...qualityRows.map(r=>r.facility), ...deviationRows.map(r=>r.facility)].filter(Boolean))]
  const orders = [...new Set([...productionRows.map(r=>r.order), ...qualityRows.map(r=>r.order)].filter(Boolean))]
  const materials = [...new Set([...productionRows.map(r=>r.material), ...qualityRows.map(r=>r.material), ...deviationRows.map(r=>r.material)].filter(Boolean))]
  const descriptions = [...new Set(productionRows.map(r=>r.desc).filter(Boolean))]
  const routingGroups = [...new Set(productionRows.map(r=>r.routingGroup).filter(Boolean))]
  const totalQty = productionRows.reduce((sum,r)=>sum+num(r.qty),0)
  const prodDates = productionRows.map(r=>r.date).filter(Boolean).sort((a,b)=>a-b)
  const qualityDates = qualityRows.map(r=>r.date).filter(Boolean).sort((a,b)=>a-b)
  const allQuality = qualityRows.filter(r=>r.characteristic).map((r,index) => {
    const status = normalize(r.status || r.approval).toLowerCase()
    const rejected = ['rejection','rejected','fail','failed','פסול','לא תקין','חריג'].some(x=>status.includes(x))
    return {...r, rejected, _key:`${r.characteristic}-${r.inspectionLot}-${index}`}
  })
  const badCount = allQuality.filter(r=>r.rejected).length
  const goodCount = allQuality.length-badCount
  const shownQuality = allQuality.filter(r=>qualityFilter==='all' || (qualityFilter==='bad' ? r.rejected : !r.rejected))
  const qaApprovals = [...new Set(qualityRows.map(r=>r.approval || r.status).filter(Boolean))]
  const inspectionLots = [...new Set(qualityRows.map(r=>r.inspectionLot).filter(Boolean))]
  const qualityPct = allQuality.length ? Math.round(goodCount/allQuality.length*100) : 0
  const released = qaApprovals.some(v => ['approved','released','מאושר','שוחרר'].some(x=>normalize(v).toLowerCase().includes(x)))
  const steps = [
    {label:'ייצור / אריזה', done:productionRows.length>0, date:prodDates[0]},
    {label:'בדיקות איכות', done:qualityRows.length>0, date:qualityDates[0]},
    {label:'החלטת QA', done:qaApprovals.length>0, date:qualityDates.at(-1)},
    {label:'שחרור', done:released, date:released ? qualityDates.at(-1) : null},
  ]
  const exportBatch = () => {
    const wb = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(productionRows.map(r=>({Date:iso(r.date),Facility:r.facility,RoutingGroup:r.routingGroup,Order:r.order,Batch:r.batch,Material:r.material,Description:r.desc,Quantity:r.qty}))), 'Production')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(allQuality.map(r=>({Date:iso(r.date),InspectionLot:r.inspectionLot,Material:r.material||materials.join(', '),Characteristic:r.characteristic,Result:r.value||r.qualitative,Lower:r.lower,Upper:r.upper,Unit:r.unit,Status:r.rejected?'חריג':'תקין',Remarks:r.remarks}))), 'Quality')
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(deviationRows.map(r=>({Date:iso(r.date),Facility:r.facility,Material:r.material||materials.join(', '),Status:r.status,UDCode:r.udCode,Remarks:r.remarks}))), 'Deviations')
    XLSX.writeFile(wb, `Batch_${data.batch}.xlsx`)
  }
  return <div className="batch-modal-backdrop" onMouseDown={e=>{if(e.target===e.currentTarget) onClose()}}>
    <section className="batch-control-card" role="dialog" aria-modal="true" aria-label={`כרטיס מנה ${data.batch}`}>
      <header className="batch-card-head"><div><span className="batch-eyebrow">BATCH CONTROL CENTER</span><h2>כרטיס מנה — {data.batch}</h2><p>{descriptions.join(' · ') || materials.join(', ') || 'ללא תיאור חומר'}</p></div><div className="batch-head-actions"><button type="button" className="batch-export" onClick={exportBatch}><Download size={17}/> ייצוא מנה</button><button type="button" className="batch-close" onClick={onClose} aria-label="סגירה"><X/></button></div></header>
      <div className="batch-summary-grid">
        <BatchMetric label="Batch" value={data.batch}/><BatchMetric label="Order" value={orders.join(', ') || '—'}/><BatchMetric label="מק״ט חומר" value={materials.join(', ') || '—'}/><BatchMetric label="מתקן" value={facilities.join(', ') || '—'}/><BatchMetric label="Routing group" value={routingGroups.join(', ') || '—'}/><BatchMetric label="כמות ארוזה" value={fmt(totalQty)}/><BatchMetric label="Inspection Lot" value={inspectionLots.join(', ') || '—'}/><BatchMetric label="QA" value={qaApprovals.join(', ') || 'טרם התקבלה החלטה'}/>
      </div>
      <div className="batch-kpi-row"><div className="batch-kpi good"><span>מאפיינים תקינים</span><b>{goodCount}</b></div><div className="batch-kpi bad"><span>מאפיינים חריגים</span><b>{badCount}</b></div><div className="batch-kpi"><span>אחוז הצלחה</span><b>{qualityPct}%</b></div><div className="batch-kpi"><span>חריגות פתוחות</span><b>{deviationRows.length}</b></div></div>
      <div className="batch-card-body">
        <section className="batch-panel quality-panel"><div className="batch-panel-title"><div><h3>תוצאות איכות</h3><p>כל המאפיינים שנמצאו למנה בקובץ האיכות</p></div><div className="quality-filter"><button className={qualityFilter==='all'?'active':''} onClick={()=>setQualityFilter('all')}>הכול {allQuality.length}</button><button className={qualityFilter==='good'?'active good':''} onClick={()=>setQualityFilter('good')}>תקינים {goodCount}</button><button className={qualityFilter==='bad'?'active bad':''} onClick={()=>setQualityFilter('bad')}>חריגים {badCount}</button></div></div><div className="table-wrap batch-quality-table"><table><thead><tr><th>מק״ט חומר</th><th>מאפיין</th><th>תוצאה</th><th>גבול תחתון</th><th>גבול עליון</th><th>יחידה</th><th>סטטוס</th><th>הערה</th></tr></thead><tbody>{shownQuality.map(r=><tr key={r._key} className={r.rejected?'quality-row-bad':''}><td>{r.material || materials.join(', ') || '—'}</td><td><b>{r.characteristic}</b></td><td>{r.value||r.qualitative||'—'}</td><td>{r.lower||'—'}</td><td>{r.upper||'—'}</td><td>{r.unit||'—'}</td><td><span className={`quality-status ${r.rejected?'bad':'good'}`}>{r.rejected?'חריג':'תקין'}</span></td><td>{r.remarks||'—'}</td></tr>)}{!shownQuality.length&&<tr><td colSpan="8" className="empty">לא נמצאו תוצאות במסנן שנבחר</td></tr>}</tbody></table></div></section>
        <aside className="batch-side-column"><section className="batch-panel"><div className="batch-panel-title"><div><h3>Timeline</h3><p>מצב התקדמות המנה</p></div></div><div className="batch-timeline">{steps.map((step,i)=><div className={`timeline-step ${step.done?'done':''}`} key={step.label}><i>{step.done?<CheckCircle2 size={18}/>:<Clock3 size={18}/>}</i><div><b>{step.label}</b><span>{step.date?`${iso(step.date)} ${new Date(step.date).toLocaleTimeString('he-IL',{hour:'2-digit',minute:'2-digit'})}`:'טרם הושלם'}</span></div>{i<steps.length-1&&<em/>}</div>)}</div></section><section className="batch-panel"><div className="batch-panel-title"><div><h3>חריגות והערות</h3><p>{deviationRows.length} רשומות מקושרות למנה</p></div></div><div className="batch-deviations">{deviationRows.map((r,i)=><article key={i}><div><span className="quality-status bad">{r.status||'חריגה'}</span>{r.udCode&&<b>{r.udCode}</b>}</div><p><b>מק״ט חומר: {r.material || materials.join(', ') || '—'}</b></p><p>{r.remarks||'לא הוזנה הערה'}</p><small>{iso(r.date)||'ללא תאריך'}</small></article>)}{!deviationRows.length&&<div className="batch-empty-good"><CheckCircle2/> לא נמצאו חריגות למנה</div>}</div></section></aside>
      </div>
    </section>
  </div>
}
function BatchMetric({label,value}) { return <div className="batch-metric"><span>{label}</span><b>{value}</b></div> }

function DataSource({ title, icon, meta, count, acceptLabel, busy, onFiles, canManage }) {
  const loaded = Boolean(meta || count)
  const loadedAt = meta?.loadedAt ? new Date(meta.loadedAt).toLocaleString('he-IL') : 'טרם נטען'
  return <article className={`data-source ${loaded ? 'ready' : ''}`}>
    <div className="data-source-head"><div className="data-source-icon">{icon}</div><div><h3>{title}</h3><span>{loaded ? 'תקין וזמין' : 'ממתין לקובץ'}</span></div></div>
    <div className="data-source-count"><b>{fmt(count)}</b><span>רשומות פעילות</span></div>
    <div className="data-source-meta"><small title={meta?.fileName || ''}>{meta?.fileName || 'לא נבחר קובץ'}</small><small>{loadedAt}</small>{meta?.source === 'cloud' && <small className="cloud-source-label">מקור: Supabase{meta?.loadedBy ? ` · ${meta.loadedBy}` : ''}</small>}{meta?.facilities ? <small>{meta.facilities} מתקנים זוהו במדגם</small> : null}</div>
    {canManage ? <label className={`source-upload ${busy ? 'disabled' : ''}`}><RefreshCw size={16}/>{acceptLabel}<input type="file" accept=".xlsx,.xls" disabled={busy} onChange={e => { const files=[...e.target.files]; e.target.value=''; onFiles(files) }}/></label> : <div className="viewer-lock"><ShieldCheck size={16}/> צפייה בלבד</div>}
  </article>
}

function Summary({ title, value, sub, warn }) { return <div className={`summary ${warn ? 'warn' : ''}`}><span>{title}</span><b>{value}</b><small>{sub}</small></div> }
function Executive({ icon, title, value, sub, good, warn, bad, onClick }) { return <button type="button" className={`executive ${good?'good':''} ${warn?'warn':''} ${bad?'bad':''} ${onClick?'clickable':''}`} onClick={onClick}><div className="executive-icon">{icon}</div><div><span>{title}</span><b>{value}</b><small>{sub}</small></div></button> }
function StatusBadge({ state, label }) { return <span className={`status-pill ${state}`}>{label}</span> }
function ForecastCard({ facility, routingGroup, station, lineName, target, actual, pct, remaining, requiredDaily, recentAverage, provenMax, forecast, remainingWorkdays, state, label, selected, onClick }) {
  return <article className={`forecast-card ${state} ${selected ? 'selected' : ''}`} onClick={onClick} role="button" tabIndex="0">
    <div className="forecast-head"><div><small>מתקן</small><h3>{facility}</h3>{routingGroup && <div className="forecast-resource"><b>{station || routingGroup}</b><span>{lineName || routingGroup}</span><small>{routingGroup}</small></div>}</div><StatusBadge state={state} label={label}/></div>
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