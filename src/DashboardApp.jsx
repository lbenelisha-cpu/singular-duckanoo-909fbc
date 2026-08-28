import { useEffect, useMemo, useState } from 'react'
import {
  Upload, Database, Factory, FlaskConical, CalendarDays, Search, CheckCircle2,
  AlertTriangle, Clock3, X, BarChart3, Download, Trash2, Save, Target,
  Gauge, CalendarCheck, BellRing, TrendingUp, FileSpreadsheet, ShieldCheck, RefreshCw, ClipboardList, Activity, LogOut, UserCircle, Cloud, WifiOff, ArrowLeft, HeartPulse, Printer, PanelRightClose, PanelRightOpen, Maximize2, Minimize2, Home, ChevronLeft, Settings2, Volume2, VolumeX
} from 'lucide-react'
import { loadCloudDatasetOnce, getCloudDatasetMeta, uploadCloudDataset, uploadCloudDatasetIncremental, deleteAllCloudDatasets, getCloudHealth, saveActiveTargetWorkbook, loadActiveTargetWorkbook, saveMonthlyTargetDataset, loadAllMonthlyTargetDatasets, saveMonthlyTargetWorkbook, loadMonthlyTargetWorkbook } from './cloudData'
import { supabase } from './supabase'
import { buildResourceRows } from './resourceEngine'
import { productionMappingKey, stationFamily } from './mappingEngine'
import { prodLineInfo, isExcludedProdLine, excelFacilityLabel } from './prodLineMapping'
import * as XLSXCore from 'xlsx'
import './styles.css'

// Use the styled browser build when available, but always fall back to the
// project's existing xlsx dependency so the app can never white-screen if
// the external script is blocked or slow to load.
const XLSX = window.XLSX || XLSXCore

const LEGACY_DAILY_TARGETS = {
  '1519': 80000, '1521': 60000, '1523': 40000, '1524': 6000,
  '1525': 130000, '1528': 80000, '1540': 18000, '1541': 210000,
  '1542': 60000, '1543': 60000,
}
const FACILITY_ALIASES = {
  '1519': ['1519', '19', '19-F-01', '19-F-02'],
  '1521': ['1521', '21'],
  '1523': ['1523', '23'],
  '1524': ['1524', '24'],
  '1525': ['1525', '25'],
  '1528': ['1528', '28'],
  '1540': ['1540', '40'],
  '1541': ['1541', '41'],
  '1542': ['1542', '42-P-01', 'T42A'],
  '1543': ['1543', '43', '43-P-A', '43-P-B'],
  '1142': ['1142'],
  '1123': ['1123'],
}
const PRIMARY_FACILITIES = ['1519', '1541', '1540', '1525', '1523', '1521', '1528', '1524', '1542', '1543', '1142', '1123']
const DEFAULT_FACILITIES = PRIMARY_FACILITIES
const RESOURCE_LABELS = {
  '1542|LQ-P-1': { station: 'P-02', line: '1 ליטר' },
  '1542|LQ-P-5': { station: 'P-03', line: '5 ליטר' },
  '1542|LQ-P-10': { station: 'P-04', line: '10/20 ליטר' },
}
const STORAGE_KEY = 'iml-control-center-sprint7'
const DB_NAME = 'iml-control-center-db'
const DB_STORE = 'dashboard-state'
const DB_KEY = 'sprint1182-build2-batch-material'
const TARGET_FILE_KEY = 'latest-monthly-target-workbook'
const APP_VERSION = '11.9.36'
const BUILD_LABEL = 'Sprint 11.9.36 — PROD LINE Fast Mapping'
const VERSION_CHECK_INTERVAL_MS = 5 * 60 * 1000

const FACILITY_COLOR_PALETTE = ['#E8F3FF','#E9F8EF','#FFF3D9','#F4EAFF','#FFE9EC','#E7F7F7','#F1F1F1','#FFF0E5','#EAF0FF','#F6F0E8','#E8F8FF','#FDEBFF']
// Stable, collision-free colors for the facilities used by IML CONTROL.
// This keeps the same facility color in the table, summary chips, print and Excel export.
const FACILITY_FIXED_COLORS = {
  '1123':'#FDEBFF',
  '1142':'#E8F8FF',
  '1519':'#E9F8EF',
  '1521':'#FFF3D9',
  '1523':'#F4EAFF',
  '1524':'#F1F1F1',
  '1525':'#FFE9EC',
  '1528':'#DFF4FF',
  '1540':'#FFF0E5',
  '1541':'#EAF0FF',
  '1542':'#E6F8ED',
  '1543':'#FFF0C9',
}
const facilityColorFor = facility => {
  const text = String(facility || '—').trim()
  if (FACILITY_FIXED_COLORS[text]) return FACILITY_FIXED_COLORS[text]
  // Fallback for a future facility not yet in the fixed map.
  let hash = 0
  for (let i=0;i<text.length;i++) hash = ((hash << 5) - hash + text.charCodeAt(i)) | 0
  return FACILITY_COLOR_PALETTE[Math.abs(hash) % FACILITY_COLOR_PALETTE.length]
}
const isoDate = value => {
  if (!value) return ''

  const date = value instanceof Date ? value : new Date(value)
  if (Number.isNaN(date.getTime())) return ''

  const pad = number => String(number).padStart(2, '0')

  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`
}
const MAPPING_STORAGE_KEY = 'iml-product-mappings-sprint1181'
const MAPPING_TIMELINE_KEY = 'iml-product-mapping-timeline-sprint1181'
const readLocalJson = (key, fallback = []) => {
  try { return JSON.parse(localStorage.getItem(key) || '') || fallback } catch { return fallback }
}

const initialYesterdayDate = () => { const date = new Date(); date.setDate(date.getDate() - 1); return isoDate(date) }
const initialToDate = () => initialYesterdayDate()
const initialFromDate = () => initialYesterdayDate()

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
const idbGetKey = async (key) => {
  const db = await openDashboardDb()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, 'readonly')
    const req = tx.objectStore(DB_STORE).get(key)
    req.onsuccess = () => resolve(req.result || null)
    req.onerror = () => reject(req.error)
    tx.oncomplete = () => db.close()
  })
}
const idbSetKey = async (key, value) => {
  const db = await openDashboardDb()
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, 'readwrite')
    tx.objectStore(DB_STORE).put(value, key)
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
const batchMaterialKey = (batch, material) => {
  const b = normalize(batch)
  const m = normalize(material)
  return b && m ? `${b}|${m}` : ''
}
const normalizeRouting = (v) => normalize(v).toUpperCase()
const resourceMeta = (facility, routingGroup) => RESOURCE_LABELS[`${facility}|${normalizeRouting(routingGroup)}`] || {}
const planningName = (row) => row.resource || (row.routingGroup ? `מתקן ${row.facility} · ${row.station || row.routingGroup}${row.lineName ? ` · ${row.lineName}` : ''}` : `מתקן ${row.facility}`)
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
const localDateTimeString = date => {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return ''

  const pad = value => String(value).padStart(2, '0')

  return [
    date.getFullYear(),
    pad(date.getMonth() + 1),
    pad(date.getDate()),
  ].join('-') + 'T' + [
    pad(date.getHours()),
    pad(date.getMinutes()),
    pad(date.getSeconds()),
  ].join(':')
}
const localDateOnlyString = value => {
  const date = excelDate(value)
  if (!date) return ''
  const pad = number => String(number).padStart(2, '0')
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`
}

const productionDateFromDay = day => {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(day || ''))) return null
  const [year, month, date] = day.split('-').map(Number)
  return new Date(year, month - 1, date, 12, 0, 0, 0)
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

// Production rows mapped to facility 1542 must actually belong to one of the
// approved Facility 42 packaging routings. Storage aliases such as T42A are
// logistical locations and, on their own, are not proof that the row was
// produced/packed on Facility 42. This prevents rows such as M85NOCST from
// inflating Facility 42 and the recent-production table.
const isFacility42PackagingRoute = (routingGroup, routingDescription = '') => {
  const route = normalize(`${routingGroup || ''} ${routingDescription || ''}`).toUpperCase()
  return /(^|\s)LQ-P-(1|5|10)(\s|$)/.test(route) ||
    route.includes('42-P-02') || route.includes('42-P-03') || route.includes('42-P-04') ||
    route.includes('LIQUID 1 LITER') || route.includes('LIQUID 5 LITER') || route.includes('LIQUID 10/20 LITER')
}

const productionFacility = (facilityValue, routingGroup, routingDescription = '') => {
  const facility = canonicalFacility(facilityValue)
  if (facility !== '1542') return facility
  return isFacility42PackagingRoute(routingGroup, routingDescription) ? '1542' : ''
}

const productionAssignment = (facilityValue, routingGroup, routingDescription = '', prodLineValue = '') => {
  if (isExcludedProdLine(prodLineValue)) return { facility:'', mapping:null }
  const mapping = prodLineInfo(prodLineValue)
  if (mapping?.facility) return { facility:mapping.facility, mapping }
  return { facility:productionFacility(facilityValue, routingGroup, routingDescription), mapping:null }
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

const TARGET_MONTHS = { jan:1, january:1, feb:2, february:2, mar:3, march:3, apr:4, april:4, may:5, jun:6, june:6, jul:7, july:7, aug:8, august:8, sep:9, sept:9, september:9, oct:10, october:10, nov:11, november:11, dec:12, december:12 }
const targetMonthFromTitle = (value, fallbackDate = new Date()) => {
  const text = normalize(value)
  const english = text.match(/(?:production\s+report\s+)?([a-z]{3,9})\s+(20\d{2})/i)
  if (english) { const month = TARGET_MONTHS[english[1].toLowerCase()]; if (month) return `${english[2]}-${String(month).padStart(2,'0')}` }
  return parseMonth(text) || monthKey(fallbackDate)
}
const targetFacilityIds = (resource) => {
  const text = normalize(resource)
  const upperText = text.toUpperCase()
  // Approved business mappings. EC(23) is intentionally mapped to station 1523.
  if (/^EC\s*\(23\)/i.test(text)) return ['1523']
  if (/^EC\s*\(25\)/i.test(text)) return ['1525']
  // Bromacil is produced under station 1540. Product membership comes from the DATA sheet.
  if (/^BROMACIL\s*\(25\s*,\s*40\)/i.test(text)) return ['1540']
  if (/^SHAKED\s+ISO\s+42$/i.test(text)) return ['1142']
  if (/^SHAKED\s+ISO\s+23$/i.test(text)) return ['1123']
  if (/^LQ\s*43\b/i.test(text)) return ['1543']
  if (/^SC\s*\(28\)/i.test(text)) return ['1528']
  if (/WG\s*SMALL\s+PACKS?\s*\(19\)/i.test(text) || /^WG\s*\(19\)/i.test(text)) return ['1519']
  if (/LQ\s*(1|5|10\s*\/\s*20)\s*(LT|L)/i.test(text)) return ['1542']
  // Facility 24 target names do not contain a parenthesized station number,
  // so they need an explicit mapping to storage location 1524.
  if (/^24F(?:128)?$/i.test(text)) return ['1524']

  const raw = (text.match(/\(([^)]+)\)/)?.[1] || (text.match(/\b(19|23|24|25|28|40|41|42|43)\b/g) || []).join(','))
  const plantIds = [...new Set(String(raw).match(/19|23|24|25|28|40|41|42|43/g) || [])]
  const map = { '19':'1519','23':'1523','24':'1524','25':'1525','28':'1528','40':'1540','41':'1541','42':'1542','43':'1543' }
  return plantIds.map(id => map[id]).filter(Boolean)
}
const targetDescriptionTokens = (resource) => {
  const text = normalize(resource).toUpperCase()
  if (/LQ\s*1\s*(LT|L)\b/.test(text)) return ['1L','1 L','1LT','1 LT']
  if (/LQ\s*5\s*(LT|L)\b/.test(text)) return ['5L','5 L','5LT','5 LT']
  if (/LQ\s*10\s*\/\s*20/.test(text)) return ['10L','10 L','20L','20 L','10/20']
  if (/LQ\s*43/.test(text)) return []
  const clean = text.replace(/\([^)]*\)/g,'').replace(/\bSMALL\s+PACKS?\b/g,'SMALL PACK').trim()
  const generic = new Set(['EC','SC','WG','CS','LQ','24F'])
  return !clean || generic.has(clean) ? [] : [clean]
}
const APPROVED_TARGET_RESOURCES = new Set([
  'EC (23)','SHAKED ISO 42','SHAKED ISO 23','LQ 1LT (42)','LQ 5 LT (42)','LQ 10/20 LT (42)','LQ 43','SC (28)','WG (19)','WG SMALL PACKS (19)',
  '24F128','24F','EC (25)','DIURON (40)','TOLUREX (40)','CS (25,40)','BROMACIL (25,40)','GALIGAN (25,40)',
  'PROPA PREMIX (25,40)','FLUOROCHLORIDON (25,40)','SAFLUFENACIL TECH (25,40)','METAZACHLOR (41)',
  'ATRALONE (41)','NANA (41)','D. DAMASCONE (41)'
])
const isApprovedTargetResource = value => APPROVED_TARGET_RESOURCES.has(normalize(value).toUpperCase())

const parseTargetNumber = (value) => {
  if (value === null || value === undefined || value === '' || /^\s*-+\s*$/.test(String(value)) || /DIV\/0/i.test(String(value))) return 0
  const text = String(value).trim(); const negative = /^\(.*\)$/.test(text)
  const n = Number(text.replace(/[(),%\s]/g,'').replace(/,/g,''))
  return Number.isFinite(n) ? (negative ? -n : n) : 0
}

// Sprint 11.9.0 Trial 4 — normalize legacy monthly targets.
// SUM targets are expressed in thousands (t / m³), while production rows are L / kg.
// New uploads are already multiplied by 1000; old cloud/cache rows are normalized here once in memory.
const normalizeStoredTargetRow = row => {
  const scaleLegacy = value => {
    const n = Number(value) || 0
    return n > 0 && n < 10000 ? n * 1000 : n
  }
  return { ...row, target:scaleLegacy(row?.target), capacity:scaleLegacy(row?.capacity) }
}
const isLegacyCombinedTarget = value => {
  const text = normalize(value).toUpperCase().replace(/\s+/g,' ')
  return /EC\s*\(\s*(23\s*,\s*25|25\s*,\s*23)\s*\)/.test(text) ||
    /SHAKED\s+ISO\s*\(\s*(42\s*\+\s*23|42\s*\+\s*43)\s*\)/.test(text)
}
const normalizeStoredTargets = rows => (rows || [])
  .map(normalizeStoredTargetRow)
  .filter(row => !isLegacyCombinedTarget(row?.resource || row?.lineName || ''))


const stableDateKey = value => {
  if (!value) return ''
  const date = value instanceof Date ? value : new Date(value)
  return Number.isNaN(date.getTime()) ? normalize(value) : date.toISOString()
}

// Stable row identities used to prevent duplicate records when a file is loaded again.
const productionRowKey = row => [
  normalize(row?.facility), normalize(row?.productionDay), stableDateKey(row?.finishDate || row?.date),
  normalize(row?.order), normalize(row?.batch), normalize(row?.material), normalize(row?.routingGroup), normalize(row?.prodLine),
  normalize(row?.orderType), String(Number(row?.qty) || 0), String(Number(row?.plannedQty) || 0)
].join('|')

const qualityBusinessRowKey = row => [
  normalize(row?.inspectionLot), normalize(row?.sampleNo), normalize(row?.operationActivity),
  normalize(row?.characteristic), normalize(row?.material), normalize(row?.batch)
].join('|')
const qualityLegacyRowKey = row => [
  normalize(row?.inspectionLot), normalize(row?.batch), normalize(row?.material),
  normalize(row?.characteristic), stableDateKey(row?.date), normalize(row?.value), normalize(row?.qualitative)
].join('|')
const qualityRowKey = row => (normalize(row?.sampleNo) || normalize(row?.operationActivity)) ? qualityBusinessRowKey(row) : qualityLegacyRowKey(row)

const deviationRawRowKey = row => [
  normalize(getField(row, ['Inspection Lot','Inspection Lot #'])),
  normalize(getField(row, ['Batch','Batch Number'])),
  normalize(getField(row, ['Material #','Material Number','Material No.','מקט','מק"ט','מק״ט','Material'])),
  normalize(getField(row, ['UD Code','Usage Decision','Usage decision','החלטת שימוש'])),
  stableDateKey(excelDate(getField(row, ['Inspection Lot UD Date','Date of Lot Creation','Process Order Delivered Date','Start Date of Inspection'])))
].join('|')

const dedupeRows = (rows, keyFn) => {
  const seen = new Set()
  const output = []
  for (const row of rows || []) {
    const key = keyFn(row)
    if (!key || seen.has(key)) continue
    seen.add(key)
    output.push(row)
  }
  return output
}

async function filterNewRows(existingRows, incomingRows, keyFn, onProgress) {
  const keys = new Set()
  const total = existingRows.length + incomingRows.length
  let done = 0
  for (let i = 0; i < existingRows.length; i += 4000) {
    existingRows.slice(i, i + 4000).forEach(row => keys.add(keyFn(row)))
    done += Math.min(4000, existingRows.length - i)
    onProgress?.(done, total)
    await new Promise(resolve => setTimeout(resolve, 0))
  }
  const fresh = []
  for (let i = 0; i < incomingRows.length; i += 4000) {
    incomingRows.slice(i, i + 4000).forEach(row => {
      const key = keyFn(row)
      if (!key || keys.has(key)) return
      keys.add(key)
      fresh.push(row)
    })
    done += Math.min(4000, incomingRows.length - i)
    onProgress?.(done, total)
    await new Promise(resolve => setTimeout(resolve, 0))
  }
  return fresh
}


async function filterNewQualityRows(existingRows, incomingRows, onProgress) {
  const businessKeys = new Set(), legacyKeys = new Set()
  const total = existingRows.length + incomingRows.length
  let done = 0
  for (let i = 0; i < existingRows.length; i += 4000) {
    existingRows.slice(i, i + 4000).forEach(row => {
      if (normalize(row?.sampleNo) || normalize(row?.operationActivity)) businessKeys.add(qualityBusinessRowKey(row))
      legacyKeys.add(qualityLegacyRowKey(row))
    })
    done += Math.min(4000, existingRows.length - i); onProgress?.(done,total); await new Promise(r=>setTimeout(r,0))
  }
  const fresh=[]
  for (let i = 0; i < incomingRows.length; i += 4000) {
    incomingRows.slice(i, i + 4000).forEach(row => {
      const bk=qualityBusinessRowKey(row), lk=qualityLegacyRowKey(row)
      if (businessKeys.has(bk) || legacyKeys.has(lk)) return
      businessKeys.add(bk); legacyKeys.add(lk); fresh.push(row)
    })
    done += Math.min(4000, incomingRows.length - i); onProgress?.(done,total); await new Promise(r=>setTimeout(r,0))
  }
  return fresh
}

async function readTargetWorkbook(file) {
  const buf = await file.arrayBuffer()
  const wb = XLSX.read(buf, { type:'array', cellDates:true, dense:true })
  const output = []

  // The DATA sheet is the product master. Build Family -> Item Code dynamically.
  // Packaging facilities 1542 and 1519 intentionally keep their existing routing logic.
  const materialsByFamily = new Map()
  for (const sheetName of wb.SheetNames) {
    const matrix = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header:1, defval:'', raw:true })
    if (!matrix.length) continue
    const headerIndex = matrix.findIndex(row => {
      const keys = row.map(normKey)
      return keys.includes(normKey('Item Code')) && keys.includes(normKey('Family'))
    })
    if (headerIndex < 0) continue
    const header = matrix[headerIndex].map(normalize)
    const itemCodeIndex = header.findIndex(value => normKey(value) === normKey('Item Code'))
    const familyIndex = header.findIndex(value => normKey(value) === normKey('Family'))
    if (itemCodeIndex < 0 || familyIndex < 0) continue
    matrix.slice(headerIndex + 1).forEach(row => {
      const material = normalize(row[itemCodeIndex])
      const family = normalize(row[familyIndex]).toUpperCase()
      if (!material || !family || family === '#N/A') return
      const list = materialsByFamily.get(family) || new Set()
      list.add(material)
      materialsByFamily.set(family, list)
    })
  }

  for (const sheetName of wb.SheetNames) {
    const matrix = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header:1, defval:'', raw:true })
    if (!matrix.length) continue
    const titleText = matrix.slice(0,3).flat().map(normalize).join(' ')
    const fallbackMonth = targetMonthFromTitle(`${titleText} ${file.name}`, new Date())
    const headerIndex = matrix.findIndex(row => { const keys=row.map(normKey); return keys.some(k=>k==='capacity') && keys.some(k=>k==='plan') })
    if (headerIndex < 0) continue
    const header = matrix[headerIndex].map((v,i)=>normalize(v)||(i===0?'Resource':`Column_${i+1}`))
    for (const row of matrix.slice(headerIndex+1)) {
      const resource=normalize(row[0]); if(!resource||/^total$/i.test(resource)) continue
      const familyName = resource.replace(/\([^)]*\)/g, '').trim().toUpperCase()
      const protectedPackaging = /LQ\s*(1|5|10\s*\/\s*20)\s*(LT|L)/i.test(resource) || /WG/i.test(resource)
      const familyMaterials = protectedPackaging ? [] : [...(materialsByFamily.get(familyName) || [])]
      const item={__sheet:sheetName,Resource:resource,Month:fallbackMonth,__family:familyName,__materials:familyMaterials}
      header.forEach((key,i)=>{item[key]=row[i]??''}); output.push(item)
    }
  }
  return output
}

async function readWorkbook(file) {
  const buf = await file.arrayBuffer()
  const wb = XLSX.read(buf, { type: 'array', cellDates: false, dense: true })
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


const displayDatasetName = (kind, month = '') => {
  if (kind === 'production') return 'קובץ כמויות יצור ואריזה'
  if (kind === 'quality') return 'קובץ איכות תוצאות'
  if (kind === 'deviations') return 'קובץ איכות החלטת שימוש'
  if (kind === 'targets') {
    const mm = String(month || '').match(/-(\d{2})$/)?.[1] || String(new Date().getMonth()+1).padStart(2,'0')
    return `יעדים לחודש ${mm}`
  }
  return 'קובץ נתונים'
}

const normalizeDatasetMeta = (kind, meta, targetMonth = '') => {
  if (!meta) return null
  const month = kind === 'targets'
    ? (String(targetMonth || meta.month || '').match(/\d{4}-\d{2}/)?.[0] || '')
    : ''
  return {
    ...meta,
    fileName: displayDatasetName(kind, month),
    originalFileName: meta.originalFileName || meta.fileName || '',
  }
}



// Global table tools — generic search/filter/sort controls for every data table.
// Implemented at the DOM layer so existing calculation logic and row renderers stay untouched.
const useUniversalTableTools = () => {
  useEffect(() => {
    const normalizeTableText = value => String(value ?? '').trim().toLowerCase()

    const cellSortValue = cell => {
      const raw = String(cell?.innerText || '').trim()
      const numeric = Number(raw.replace(/,/g, '').replace(/[^0-9.\-]/g, ''))
      if (raw && Number.isFinite(numeric) && /\d/.test(raw)) return { numeric:true, value:numeric }
      const dateMatch = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/)
      if (dateMatch) return { numeric:true, value:new Date(raw).getTime() }
      return { numeric:false, value:raw.toLocaleLowerCase('he-IL') }
    }

    const applyView = table => {
      if (!table?.tBodies?.[0]) return
      const toolbar = table.parentElement?.querySelector(':scope > .smart-table-toolbar')
      if (!toolbar) return
      const globalQuery = normalizeTableText(toolbar.querySelector('[data-smart-global]')?.value)
      const columnIndex = Number(toolbar.querySelector('[data-smart-column]')?.value ?? -1)
      const columnQuery = normalizeTableText(toolbar.querySelector('[data-smart-value]')?.value)
      let visible = 0
      const rows = [...table.tBodies[0].rows]
      rows.forEach(row => {
        if (row.classList.contains('smart-empty-row')) return
        const cells = [...row.cells]
        const allText = normalizeTableText(cells.map(cell => cell.innerText).join(' '))
        const selectedText = columnIndex >= 0 ? normalizeTableText(cells[columnIndex]?.innerText) : allText
        const matchesGlobal = !globalQuery || allText.includes(globalQuery)
        const matchesColumn = !columnQuery || selectedText.includes(columnQuery)
        const show = matchesGlobal && matchesColumn
        row.style.display = show ? '' : 'none'
        if (show) visible += 1
      })
      const count = toolbar.querySelector('[data-smart-count]')
      if (count) count.textContent = `${visible} רשומות`
      const sumBadge = toolbar.querySelector('[data-smart-sum]')
      const sumColumn = Number(table.dataset.smartSumColumn ?? -1)
      if (sumBadge && sumColumn >= 0) {
        const total = rows.reduce((sum,row) => {
          if (row.style.display === 'none' || row.classList.contains('smart-empty-row')) return sum
          const raw = String(row.cells[sumColumn]?.innerText || '').replace(/,/g,'').replace(/[^0-9.\-]/g,'')
          const value = Number(raw)
          return sum + (Number.isFinite(value) ? value : 0)
        }, 0)
        const formatter = new Intl.NumberFormat('he-IL',{maximumFractionDigits:2})
        sumBadge.textContent = `סה״כ כמות: ${formatter.format(total)}`

        // Optional dynamic breakdown: sums the currently visible rows by a selected
        // facility/station/resource column. This follows all smart-table filters.
        const groupColumn = Number(table.dataset.smartGroupColumn ?? -1)
        const groupWrap = toolbar.querySelector('[data-smart-group-summary]')
        if (groupWrap && groupColumn >= 0) {
          const groups = new Map()
          rows.forEach(row => {
            if (row.style.display === 'none' || row.classList.contains('smart-empty-row')) return
            const routingLabel = String(row.cells[groupColumn]?.innerText || '').trim() || 'ללא שיוך'
            const raw = String(row.cells[sumColumn]?.innerText || '').replace(/,/g,'').replace(/[^0-9.\-]/g,'')
            const value = Number(raw)
            const facility = String(row.dataset.facility || '').trim()
            const isFacility42 = ['1542','1142'].includes(facility)
            const useFacilitySummary = table.dataset.smartFacilitySummary === '1' && facility && !isFacility42
            const label = useFacilitySummary ? `מתקן ${facility}` : routingLabel
            const current = groups.get(label) || { total:0, facilities:new Set() }
            current.total += Number.isFinite(value) ? value : 0
            if (facility) current.facilities.add(facility)
            groups.set(label, current)
          })
          groupWrap.innerHTML = [...groups.entries()]
            .sort((a,b) => String(a[0]).localeCompare(String(b[0]), 'he', {numeric:true, sensitivity:'base'}))
            .map(([label,info]) => {
              const facility = [...info.facilities][0] || ''
              const bg = facility ? facilityColorFor(facility) : ''
              const style = bg ? ` style="background:${bg};border-color:${bg}"` : ''
              return `<span class="smart-group-chip"${style}><span>${label}</span><b>${formatter.format(info.total)}</b></span>`
            })
            .join('')
          groupWrap.style.display = groups.size ? '' : 'none'
        }
      }
    }

    const sortTable = (table, columnIndex, direction) => {
      const body = table?.tBodies?.[0]
      if (!body || columnIndex < 0 || !direction) return
      const rows = [...body.rows].filter(row => row.cells.length > columnIndex)
      rows.sort((a,b) => {
        const av = cellSortValue(a.cells[columnIndex])
        const bv = cellSortValue(b.cells[columnIndex])
        let result
        if (av.numeric && bv.numeric) result = av.value - bv.value
        else result = String(av.value).localeCompare(String(bv.value), 'he', { numeric:true, sensitivity:'base' })
        return direction === 'desc' ? -result : result
      })
      rows.forEach(row => body.appendChild(row))
      applyView(table)
    }

    const enhanceTable = table => {
      if (!table || table.dataset.smartTableReady === '1' || table.dataset.noSmartTable === '1') return
      const headerCells = [...(table.tHead?.rows?.[0]?.cells || [])]
      if (!headerCells.length || !table.tBodies?.length) return
      table.dataset.smartTableReady = '1'
      table.classList.add('smart-table-enabled')
      const wrap = table.parentElement
      if (!wrap) return

      const toolbar = document.createElement('div')
      toolbar.className = 'smart-table-toolbar'
      const options = headerCells.map((cell,index) => `<option value="${index}">${String(cell.innerText || `עמודה ${index+1}`).trim()}</option>`).join('')
      toolbar.innerHTML = `
        <div class="smart-table-search"><span>⌕</span><input data-smart-global type="search" placeholder="חיפוש בכל הטבלה..." autocomplete="off" /></div>
        <label class="smart-table-column"><span>עמודה</span><select data-smart-column><option value="-1">כל העמודות</option>${options}</select></label>
        <div class="smart-table-search smart-table-value"><span>≡</span><input data-smart-value type="search" placeholder="ערך לסינון, למשל A1" autocomplete="off" /></div>
        <label class="smart-table-sort"><span>מיון</span><select data-smart-sort><option value="">ללא</option><option value="asc">עולה ↑</option><option value="desc">יורד ↓</option></select></label>
        <button type="button" class="smart-table-clear">נקה סינון</button>
        ${table.dataset.smartSumColumn != null ? '<b data-smart-sum>סה״כ כמות: 0</b>' : ''}
        ${table.dataset.smartGroupColumn != null ? '<div class="smart-group-summary" data-smart-group-summary></div>' : ''}
        <b data-smart-count>${table.tBodies[0].rows.length} רשומות</b>
      `
      wrap.insertBefore(toolbar, table)

      const onFilter = () => applyView(table)
      toolbar.querySelector('[data-smart-global]')?.addEventListener('input', onFilter)
      toolbar.querySelector('[data-smart-column]')?.addEventListener('change', () => {
        applyView(table)
        const direction = toolbar.querySelector('[data-smart-sort]')?.value
        const index = Number(toolbar.querySelector('[data-smart-column]')?.value ?? -1)
        if (direction && index >= 0) sortTable(table,index,direction)
      })
      toolbar.querySelector('[data-smart-value]')?.addEventListener('input', onFilter)
      toolbar.querySelector('[data-smart-sort]')?.addEventListener('change', event => {
        const index = Number(toolbar.querySelector('[data-smart-column]')?.value ?? -1)
        if (index >= 0 && event.target.value) sortTable(table,index,event.target.value)
      })
      toolbar.querySelector('.smart-table-clear')?.addEventListener('click', () => {
        const global = toolbar.querySelector('[data-smart-global]')
        const value = toolbar.querySelector('[data-smart-value]')
        const column = toolbar.querySelector('[data-smart-column]')
        const sort = toolbar.querySelector('[data-smart-sort]')
        if (global) global.value = ''
        if (value) value.value = ''
        if (column) column.value = '-1'
        if (sort) sort.value = ''
        applyView(table)
      })

      headerCells.forEach((cell,index) => {
        if (cell.querySelector('button')) return
        cell.classList.add('smart-sortable-header')
        cell.title = 'לחץ למיון עולה / יורד'
        cell.addEventListener('click', event => {
          if (event.target.closest('input,select,button,a')) return
          const next = cell.dataset.sortDirection === 'asc' ? 'desc' : 'asc'
          headerCells.forEach(other => { if (other !== cell) delete other.dataset.sortDirection })
          cell.dataset.sortDirection = next
          sortTable(table,index,next)
        })
      })
      applyView(table)
    }

    const scan = () => document.querySelectorAll('.dashboard table').forEach(enhanceTable)
    let scheduled = false
    const scheduleScan = () => {
      if (scheduled) return
      scheduled = true
      requestAnimationFrame(() => {
        scheduled = false
        scan()
        document.querySelectorAll('.dashboard table[data-smart-table-ready="1"]').forEach(applyView)
      })
    }
    scan()
    const observer = new MutationObserver(scheduleScan)
    observer.observe(document.body, { childList:true, subtree:true })
    return () => observer.disconnect()
  }, [])
}


const autoFitExcelSheet = worksheet => {
  if (!worksheet?.['!ref']) return worksheet
  const range = XLSX.utils.decode_range(worksheet['!ref'])
  const cols = []
  for (let c = range.s.c; c <= range.e.c; c += 1) {
    let maxLen = 0
    for (let r = range.s.r; r <= range.e.r; r += 1) {
      const cell = worksheet[XLSX.utils.encode_cell({ r, c })]
      const value = cell?.w ?? cell?.v ?? ''
      const len = String(value ?? '').split(/\r?\n/).reduce((m, line) => Math.max(m, line.length), 0)
      maxLen = Math.max(maxLen, len)
    }
    cols.push({ wch: Math.min(100, Math.max(10, maxLen + 3)) })
  }
  worksheet['!cols'] = cols
  return worksheet
}

const appendAutoFitJsonSheet = (workbook, rows, name) => {
  const worksheet = XLSX.utils.json_to_sheet(rows)
  autoFitExcelSheet(worksheet)
  // Excel sheet view: right-to-left for Hebrew/IML exports.
  worksheet['!views'] = [{ rightToLeft:true }]
  XLSX.utils.book_append_sheet(workbook, worksheet, name)
  return worksheet
}


const UI_SOUNDS_STORAGE_KEY = 'iml-ui-sounds-enabled'

const playUiTone = (type = 'click') => {
  try {
    const AudioCtx = window.AudioContext || window.webkitAudioContext
    if (!AudioCtx) return
    const ctx = new AudioCtx()
    const now = ctx.currentTime
    const gain = ctx.createGain()
    gain.connect(ctx.destination)
    const presets = {
      click:{ f1:520, f2:660, duration:.07, volume:.035, wave:'sine' },
      nav:{ f1:440, f2:620, duration:.09, volume:.04, wave:'sine' },
      select:{ f1:620, f2:760, duration:.08, volume:.035, wave:'triangle' },
      success:{ f1:660, f2:880, duration:.16, volume:.045, wave:'sine' },
      export:{ f1:560, f2:820, duration:.18, volume:.045, wave:'triangle' },
      print:{ f1:480, f2:700, duration:.12, volume:.04, wave:'square' },
      refresh:{ f1:420, f2:640, duration:.15, volume:.04, wave:'sine' },
      warning:{ f1:330, f2:260, duration:.18, volume:.045, wave:'sawtooth' },
      close:{ f1:520, f2:360, duration:.12, volume:.04, wave:'sine' },
    }
    const p = presets[type] || presets.click
    const osc = ctx.createOscillator()
    osc.type = p.wave
    osc.frequency.setValueAtTime(p.f1, now)
    osc.frequency.exponentialRampToValueAtTime(Math.max(40, p.f2), now + p.duration)
    gain.gain.setValueAtTime(0.0001, now)
    gain.gain.exponentialRampToValueAtTime(p.volume, now + .015)
    gain.gain.exponentialRampToValueAtTime(0.0001, now + p.duration)
    osc.connect(gain)
    osc.start(now)
    osc.stop(now + p.duration + .02)
    osc.onended = () => ctx.close().catch(() => {})
  } catch {}
}

const speakBye = () => {
  try {
    if (!('speechSynthesis' in window)) return
    window.speechSynthesis.cancel()
    const utterance = new SpeechSynthesisUtterance('Bye')
    utterance.lang = 'en-US'
    utterance.rate = 1
    utterance.pitch = 1
    utterance.volume = .75
    window.speechSynthesis.speak(utterance)
  } catch {}
}

const classifyButtonSound = button => {
  const text = String(button?.innerText || button?.textContent || '').trim().toLowerCase()
  const aria = String(button?.getAttribute?.('aria-label') || '').trim().toLowerCase()
  const title = String(button?.getAttribute?.('title') || '').trim().toLowerCase()
  const combined = `${text} ${aria} ${title}`
  if (/יציאה|התנתק|sign out|logout|exit/.test(combined)) return 'exit'
  if (/ייצוא|יצוא|excel|download/.test(combined)) return 'export'
  if (/הדפס|print/.test(combined)) return 'print'
  if (/רענן|refresh|עדכון/.test(combined)) return 'refresh'
  if (/מחק|מחיקה|delete|אזהרה|warning/.test(combined)) return 'warning'
  if (/שמור|אישור|אשר|save|approve|העלה|טעינת/.test(combined)) return 'success'
  if (/מתקן|בחר|סינון|filter|select/.test(combined)) return 'select'
  if (/דף|סקירה|ניהול|איכות|תחזית|מגמה|פרטים|כניסה|בית|home/.test(combined)) return 'nav'
  if (/סגור|נקה|×|✕/.test(combined)) return 'close'
  return 'click'
}

const useUiSounds = enabled => {
  useEffect(() => {
    const onClick = event => {
      if (!enabled) return
      const button = event.target?.closest?.('button,[role="button"],a')
      if (!button || button.dataset?.noUiSound === '1') return
      const type = classifyButtonSound(button)
      if (type === 'exit') {
        playUiTone('close')
        speakBye()
      } else {
        playUiTone(type)
      }
    }
    document.addEventListener('click', onClick, true)
    return () => document.removeEventListener('click', onClick, true)
  }, [enabled])
}

export default function DashboardApp({ currentUser, userRole = 'viewer', isGuest = false, onSignOut, onRequestAdminLogin }) {
  useUniversalTableTools()

  const canManageData = ['admin', 'manager'].includes(userRole)
  const canDeleteData = userRole === 'admin'
  const [uiSoundsEnabled, setUiSoundsEnabled] = useState(() => localStorage.getItem(UI_SOUNDS_STORAGE_KEY) !== '0')
  useUiSounds(uiSoundsEnabled)
  useEffect(() => { localStorage.setItem(UI_SOUNDS_STORAGE_KEY, uiSoundsEnabled ? '1' : '0') }, [uiSoundsEnabled])

  const [production, setProduction] = useState([])
  const [quality, setQuality] = useState([])
  const [deviations, setDeviations] = useState([])
  const [targets, setTargets] = useState([])
  const [status, setStatus] = useState('ממתין לטעינת קבצים')
  const [busy, setBusy] = useState(false)
  const [query, setQuery] = useState('')
  const [from, setFrom] = useState(initialFromDate)
  const [to, setTo] = useState(initialToDate)
  const [selectedFacilities, setSelectedFacilities] = useState([])
  const [activeTab, setActiveTab] = useState('production')
  const [productionSort, setProductionSort] = useState({ key:'date', direction:'desc' })
  const [quantityVarianceRow, setQuantityVarianceRow] = useState(null)
  const [facilityViewMode, setFacilityViewMode] = useState('relevant')
  const [showFacilityOverview, setShowFacilityOverview] = useState(true)
  const [showMonthlyForecast, setShowMonthlyForecast] = useState(true)
  const [simulatorOnlyIssues, setSimulatorOnlyIssues] = useState(true)
  const [manualMappings, setManualMappings] = useState(() => readLocalJson(MAPPING_STORAGE_KEY, []))
  const [mappingTimeline, setMappingTimeline] = useState(() => readLocalJson(MAPPING_TIMELINE_KEY, []))
  const [mappingDialog, setMappingDialog] = useState(null)
  const [mappingTargetKey, setMappingTargetKey] = useState('')
  const [mappingMessage, setMappingMessage] = useState('')
  const [targetMappings, setTargetMappings] = useState([])
  const [targetMappingDialog, setTargetMappingDialog] = useState(null)
  const [targetMappingFamily, setTargetMappingFamily] = useState('')
  const [planningMonth, setPlanningMonth] = useState('')
  const [additionalFacilities, setAdditionalFacilities] = useState([])
  const [facilityToAdd, setFacilityToAdd] = useState('')
  // Dashboard scope rule: regular KPIs, tables, quality and planning only use
  // facilities that are currently available in the main facility picker.
  // A discovered facility starts contributing only after a manager explicitly
  // adds it to the picker. This prevents unrelated storage locations in SAP
  // exports from inflating IML CONTROL totals.
  const facilities = useMemo(() => [...PRIMARY_FACILITIES, ...additionalFacilities], [additionalFacilities])
  const selectableFacilitySet = useMemo(() => new Set(facilities.map(String)), [facilities])
  const [dailyAdditionalFacilities, setDailyAdditionalFacilities] = useState(() => readLocalJson('iml-daily-additional-facilities', []))
  const [dailyFacilityToAdd, setDailyFacilityToAdd] = useState('')
  const [periodYear, setPeriodYear] = useState('')
  const [periodQuarter, setPeriodQuarter] = useState('')
  const [dataMeta, setDataMeta] = useState({ production:null, quality:null, deviations:null, targets:null })
  const [selectedBatch, setSelectedBatch] = useState('')
  const [selectedBatchMaterial, setSelectedBatchMaterial] = useState('')
  const [selectedResource, setSelectedResource] = useState(null)
  const [cloudState, setCloudState] = useState({ mode:'connecting', lastSync:null, message:'מתחבר למסד המשותף...', latencyMs:null, live:false })
  const [uploadProgress, setUploadProgress] = useState(null)
  const [perfStats, setPerformance] = useState({ startedAt:0, cache:'MISS', queries:0, rows:0, loadMs:0, phase:'אתחול' })
  const [sidebarCollapsed, setSidebarCollapsed] = useState(() => localStorage.getItem('iml-ui-sidebar-collapsed') === '1')
  const [managementMode, setManagementMode] = useState(() => localStorage.getItem('iml-ui-management-mode') === '1')
  const [diagnosticsOpen, setDiagnosticsOpen] = useState(false)
  const [showHome, setShowHome] = useState(true)
  const [facilityPickerOpen, setFacilityPickerOpen] = useState(false)
  const [showDataCenter, setShowDataCenter] = useState(false)
  const [dailyEventFormOpen, setDailyEventFormOpen] = useState(false)
  const [dailyEventHistoryOpen, setDailyEventHistoryOpen] = useState(false)
  const [dailyEventType, setDailyEventType] = useState('')
  const [dailyEventFacility, setDailyEventFacility] = useState('')
  const [dailyEventSeverity, setDailyEventSeverity] = useState('')
  const [dailyEventText, setDailyEventText] = useState('')
  const [dailyEventDate, setDailyEventDate] = useState('')
  const [dailyEvents, setDailyEvents] = useState(() => readLocalJson('iml-daily-events', []))
  const [dailyReportHistory, setDailyReportHistory] = useState(() => readLocalJson('iml-daily-report-history', []))
  const [dailyCloudReady, setDailyCloudReady] = useState(false)
  const [availableUpdate, setAvailableUpdate] = useState(null)

  useEffect(() => { localStorage.setItem('iml-ui-sidebar-collapsed', sidebarCollapsed ? '1' : '0') }, [sidebarCollapsed])
  useEffect(() => { localStorage.setItem('iml-ui-management-mode', managementMode ? '1' : '0') }, [managementMode])
  useEffect(() => { localStorage.setItem('iml-daily-additional-facilities', JSON.stringify(dailyAdditionalFacilities)) }, [dailyAdditionalFacilities])
  useEffect(() => { localStorage.setItem('iml-daily-report-history', JSON.stringify(dailyReportHistory.slice(-5000))) }, [dailyReportHistory])
  useEffect(() => { localStorage.setItem('iml-daily-events', JSON.stringify(dailyEvents.slice(-3000))) }, [dailyEvents])

  useEffect(() => {
    if (!supabase) return
    let active = true
    const normalizeCloudEvent = row => ({
      id: row.external_id || row.id,
      date: row.event_date,
      type: row.event_type || '',
      facility: String(row.facility || ''),
      severity: row.severity || '',
      description: row.description || '',
      createdAt: row.created_at || '',
      createdBy: row.created_by || '',
    })
    const normalizeCloudReport = row => ({
      id: row.id,
      importedAt: row.imported_at || '',
      fileName: row.file_name || '',
      reportDate: row.report_date || '',
      material: row.material || '',
      facility: String(row.facility || ''),
      line: row.line || '',
      description: row.description || '',
      batch: row.batch || '',
      machineStatus: row.machine_status || '',
      quantity: num(row.quantity),
      facilityTotal: num(row.facility_total),
      notes: row.notes || '',
    })
    const syncDailyCloud = async () => {
      try {
        const localEvents = readLocalJson('iml-daily-events', [])
        const localReports = readLocalJson('iml-daily-report-history', [])
        if (currentUser?.id && localEvents.length) {
          const payload = localEvents.slice(-3000).filter(e => e?.id && e?.date).map(e => ({
            external_id:String(e.id), event_date:e.date, event_type:e.type || '', facility:String(e.facility || ''),
            severity:e.severity || '', description:e.description || '', created_at:e.createdAt || new Date().toISOString(),
            created_by:e.createdBy || currentUser?.email || '', created_by_id:currentUser?.id || null,
          }))
          if (payload.length) await supabase.from('iml_daily_events').upsert(payload, { onConflict:'external_id', ignoreDuplicates:true })
        }
        if (currentUser?.id && localReports.length) {
          const payload = localReports.slice(-5000).filter(r => r?.reportDate && r?.material).map((r, index) => ({
            client_key: r.clientKey || `${reportDateToIso(r.reportDate) || r.reportDate}|${r.facility || ''}|${r.material || ''}|${r.batch || ''}|${r.line || ''}|${r.fileName || ''}|${r.importedAt || ''}|${index}`,
            imported_at:r.importedAt || new Date().toISOString(), file_name:r.fileName || '', report_date:reportDateToIso(r.reportDate) || r.reportDate,
            material:r.material || '', facility:String(r.facility || ''), line:r.line || '', description:r.description || '', batch:r.batch || '',
            machine_status:r.machineStatus || '', quantity:num(r.quantity), facility_total:num(r.facilityTotal), notes:r.notes || '',
            created_by:currentUser?.email || '', created_by_id:currentUser?.id || null,
          }))
          if (payload.length) await supabase.from('iml_daily_report_rows').upsert(payload, { onConflict:'client_key', ignoreDuplicates:true })
        }
        const [{ data:eventRows, error:eventError }, { data:reportRows, error:reportError }] = await Promise.all([
          supabase.from('iml_daily_events').select('*').order('event_date', { ascending:false }).order('created_at', { ascending:false }).limit(5000),
          supabase.from('iml_daily_report_rows').select('*').order('report_date', { ascending:false }).order('imported_at', { ascending:false }).limit(10000),
        ])
        if (eventError) throw eventError
        if (reportError) throw reportError
        if (!active) return
        setDailyEvents((eventRows || []).map(normalizeCloudEvent))
        setDailyReportHistory((reportRows || []).map(normalizeCloudReport))
        setDailyCloudReady(true)
      } catch (error) {
        console.warn('Daily cloud sync failed; using browser cache', error)
        if (active) setDailyCloudReady(false)
      }
    }
    syncDailyCloud()
    return () => { active = false }
  }, [currentUser?.id])
  useEffect(() => {
    if (!canManageData || sessionStorage.getItem('iml-open-data-center-after-login') !== '1') return
    sessionStorage.removeItem('iml-open-data-center-after-login')
    setShowDataCenter(true)
    window.setTimeout(() => document.getElementById('data-center-section')?.scrollIntoView({behavior:'smooth', block:'start'}), 120)
  }, [canManageData])

  useEffect(() => {
    let active = true

    const currentBundlePath = () => {
      const scripts = [...document.querySelectorAll('script[type="module"][src]')]
      const script = scripts.find(node => /\/assets\/[^/]+\.js(?:\?|$)/.test(node.src)) || scripts[0]
      if (!script?.src) return ''
      try { return new URL(script.src, window.location.origin).pathname } catch { return script.src }
    }

    const latestBundlePath = async () => {
      const response = await fetch(`/index.html?update-check=${Date.now()}`, {
        cache:'no-store',
        headers:{ 'Cache-Control':'no-cache, no-store, must-revalidate' },
      })
      if (!response.ok) return ''
      const html = await response.text()
      const doc = new DOMParser().parseFromString(html, 'text/html')
      const scripts = [...doc.querySelectorAll('script[type="module"][src]')]
      const script = scripts.find(node => /\/assets\/[^/]+\.js(?:\?|$)/.test(node.getAttribute('src') || '')) || scripts[0]
      const src = script?.getAttribute('src') || ''
      if (!src) return ''
      try { return new URL(src, window.location.origin).pathname } catch { return src }
    }

    const readReleaseInfo = async () => {
      try {
        const response = await fetch(`/version.json?ts=${Date.now()}`, { cache:'no-store' })
        return response.ok ? await response.json() : null
      } catch { return null }
    }

    const checkForUpdate = async () => {
      try {
        const currentAsset = currentBundlePath()
        const latestAsset = await latestBundlePath()
        if (!active || !currentAsset || !latestAsset) return

        if (currentAsset !== latestAsset) {
          const release = await readReleaseInfo()
          if (!active) return
          setAvailableUpdate({
            ...(release || {}),
            version:normalize(release?.version) || 'חדשה',
            currentAsset,
            latestAsset,
          })
        } else {
          setAvailableUpdate(null)
        }
      } catch (error) {
        console.debug('Build update check skipped', error)
      }
    }

    const onVisibility = () => {
      if (document.visibilityState === 'visible') checkForUpdate()
    }

    checkForUpdate()
    const timer = window.setInterval(checkForUpdate, VERSION_CHECK_INTERVAL_MS)
    document.addEventListener('visibilitychange', onVisibility)
    window.addEventListener('focus', checkForUpdate)

    return () => {
      active = false
      window.clearInterval(timer)
      document.removeEventListener('visibilitychange', onVisibility)
      window.removeEventListener('focus', checkForUpdate)
    }
  }, [])

  const refreshApplication = async () => {
    try {
      const registrations = await navigator.serviceWorker?.getRegistrations?.()
      await Promise.all((registrations || []).map(async registration => {
        try { await registration.update() } catch {}
        try { await registration.unregister() } catch {}
      }))
    } catch {}

    const url = new URL(window.location.href)
    url.searchParams.set('refresh', String(Date.now()))
    window.location.replace(url.toString())
  }

  useEffect(() => { localStorage.setItem(MAPPING_STORAGE_KEY, JSON.stringify(manualMappings)) }, [manualMappings])
  useEffect(() => { localStorage.setItem(MAPPING_TIMELINE_KEY, JSON.stringify(mappingTimeline.slice(0,500))) }, [mappingTimeline])

  useEffect(() => {
    if (!supabase || !currentUser?.id) return
    let active = true
    supabase.from('iml_product_mappings').select('*').order('created_at', { ascending:false })
      .then(({ data, error }) => {
        if (!active || error || !data?.length) return
        setManualMappings(data.map(row => ({
          id:row.id, rowKey:row.row_key, material:row.material || '', description:row.description || '',
          family:row.facility_family, targetKey:row.target_key, targetResource:row.target_resource,
          status:row.status, active:row.active !== false, createdBy:row.created_by_email || '',
          createdAt:row.created_at, approvedAt:row.approved_at || null, cloud:true,
        })))
      })
    return () => { active = false }
  }, [currentUser?.id])

  useEffect(() => {
    if (!supabase || !currentUser?.id) return
    let active = true
    supabase.from('iml_target_mappings').select('*').eq('active', true).order('updated_at', { ascending:false })
      .then(({ data, error }) => {
        if (!active || error) return
        setTargetMappings((data || []).map(row => ({
          id:row.id, resource:row.target_resource, month:row.target_month || '', family:row.facility_family,
          createdBy:row.created_by_email || '', createdAt:row.created_at, updatedAt:row.updated_at, cloud:true,
        })))
      })
    return () => { active = false }
  }, [currentUser?.id])

  useEffect(() => {
    let active = true
    const startedAt = performance.now()
    const kinds = ['production', 'quality', 'deviations', 'targets']

    const applyDataset = (kind, dataset) => {
      const rows = dataset?.rows || []
      if (kind === 'production') {
  setProduction(
    dedupeRows(rows, productionRowKey).map(row => {
      const sourceDate = row.finishDate || row.date

      return {
        ...row,
        date: sourceDate
          ? new Date(sourceDate)
          : null,
      }
    })
  )
}
      if (kind === 'quality') setQuality(dedupeRows(rows.map(row => row?.__compactQuality && row.date ? { ...row, date:new Date(row.date) } : row), qualityRowKey))
      if (kind === 'deviations') setDeviations(dedupeRows(rows, deviationRawRowKey))
      if (kind === 'targets') setTargets(normalizeStoredTargets(rows))
      setDataMeta(current => ({ ...current, [kind]:normalizeDatasetMeta(kind, dataset?.meta || null, kind === 'targets' ? (rows?.[0]?.month || planningMonth) : '') }))
      return rows.length
    }

    ;(async () => {
      let cached = null
      try {
        cached = await idbGet()
        if (active && cached) {
          setProduction(
  dedupeRows((cached.production || []), productionRowKey).map(row => {
    const sourceDate = row.finishDate || row.date

    return {
      ...row,
      date: sourceDate
        ? sourceDate instanceof Date
          ? sourceDate
          : new Date(sourceDate)
        : null,
    }
  })
)
          setQuality(dedupeRows(cached.quality || [], qualityRowKey))
          setDeviations(dedupeRows(cached.deviations || [], deviationRawRowKey))
          setTargets(normalizeStoredTargets(cached.targets || []))
          setDataMeta({
            production: normalizeDatasetMeta('production', cached.dataMeta?.production),
            quality: normalizeDatasetMeta('quality', cached.dataMeta?.quality),
            deviations: normalizeDatasetMeta('deviations', cached.dataMeta?.deviations),
            targets: normalizeDatasetMeta('targets', cached.dataMeta?.targets, (cached.targets || [])[0]?.month || planningMonth),
          })
          setStatus('מוצג מטמון מקומי — בודק עדכונים מהשרת...')
          setCloudState({ mode:'connecting', lastSync:cached.savedAt || null, message:'הדשבורד זמין; בודק גרסאות חדשות ברקע...', latencyMs:null, live:false })
          setPerformance(current => ({ ...current, cache:'HIT', rows:(cached.production?.length||0)+(cached.quality?.length||0)+(cached.deviations?.length||0)+(cached.targets?.length||0), phase:'בדיקת גרסאות' }))
        }
      } catch (error) { console.warn('Cache restore failed', error) }

      try {
        const remoteMeta = {}
        await Promise.all(kinds.map(async kind => { remoteMeta[kind] = await getCloudDatasetMeta(kind) }))
        if (!active) return
        setPerformance(current => ({ ...current, queries:current.queries + 4 }))

        const changed = kind => {
          const remote = remoteMeta[kind]
          const local = cached?.dataMeta?.[kind]
          if (!remote) return !local
          const remoteId = remote.active_version_id || remote.updated_at || remote.loaded_at
          const localId = local?.versionId || local?.loadedAt
          return !cached || !local || String(remoteId || '') !== String(localId || '')
        }

        let loadedRows = 0
        if (changed('production')) {
          setStatus('טוען נתוני ייצור מעודכנים...')
          loadedRows += applyDataset('production', await loadCloudDatasetOnce('production'))
          setPerformance(current => ({ ...current, queries:current.queries + 1, phase:'הדשבורד זמין' }))
          await new Promise(resolve => setTimeout(resolve, 0))
        }

        for (const kind of ['targets', 'quality', 'deviations']) {
          if (!active) return
          if (!changed(kind)) continue
          setStatus(`טוען ${kind} ברקע...`)
          loadedRows += applyDataset(kind, await loadCloudDatasetOnce(kind))
          setPerformance(current => ({ ...current, queries:current.queries + 1, phase:`נטען ${kind}` }))
          await new Promise(resolve => setTimeout(resolve, 0))
        }

        // Sprint 11.9.34: the monthly archive is the source of truth for targets.
        // It contains August, September, October... together, so changing the
        // planning month never makes an older month disappear.
        try {
          const monthlyArchive = await loadAllMonthlyTargetDatasets()
          if (monthlyArchive?.rows?.length) {
            setTargets(normalizeStoredTargets(monthlyArchive.rows))
            setDataMeta(current => ({ ...current, targets:normalizeDatasetMeta('targets', monthlyArchive.meta, monthlyArchive.months?.at(-1)?.month || planningMonth) }))
            loadedRows += monthlyArchive.rows.length
          }
        } catch (monthlyTargetError) {
          console.warn('Monthly target archive unavailable; active target dataset remains as fallback', monthlyTargetError)
        }

        // Keep the exact target workbook synchronized as well as the normalized rows.
        // This solves the case where computer B sees the new target cards but still
        // downloads an older local template.
        if (remoteMeta.targets) {
          try {
            const targetMonthForWorkbook = planningMonth || monthKey(new Date())
            const cloudWorkbook = await loadMonthlyTargetWorkbook(targetMonthForWorkbook).catch(() => null) || await loadActiveTargetWorkbook()
            if (cloudWorkbook?.bytes) await idbSetKey(TARGET_FILE_KEY, cloudWorkbook)
          } catch (workbookError) {
            console.warn('Target workbook background sync skipped', workbookError)
          }
        }

        const health = await getCloudHealth().catch(() => null)
        if (!active) return
        const elapsed = Math.round(performance.now() - startedAt)
        const lastSync = Object.values(remoteMeta).map(x => x?.loaded_at || x?.updated_at).filter(Boolean).sort().at(-1) || new Date().toISOString()
        setCloudState({ mode:'cloud', lastSync, message:'מחובר לענן — טעינה חכמה של 7 ימים כברירת מחדל', latencyMs:health?.latencyMs ?? null, live:true })
        setStatus(cached && loadedRows === 0 ? 'הנתונים במטמון מעודכנים — לא נדרשה הורדה מחדש' : 'הנתונים המעודכנים נטענו בהצלחה')
        setPerformance(current => ({ ...current, loadMs:elapsed, rows:current.rows + loadedRows, phase:'הושלם' }))
      } catch (cloudError) {
        console.warn('Smart cloud restore failed', cloudError)
        if (!active) return
        setCloudState({ mode:cached ? 'offline' : 'error', lastSync:cached?.savedAt || null, message:cached ? 'השרת אינו זמין — מוצג מטמון מקומי' : (cloudError?.message || 'טעינת הנתונים נכשלה'), latencyMs:null, live:false })
        setStatus(cached ? 'מוצג גיבוי מקומי; הסנכרון לענן נכשל' : 'לא נמצאו נתונים זמינים')
      }
    })()
    return () => { active = false }
  }, [])

  useEffect(() => {
    if (!supabase) return
    let refreshTimer
    const channel = supabase.channel('iml-data-sources-live')
      .on('postgres_changes', { event:'*', schema:'public', table:'iml_data_sources' }, payload => {
        clearTimeout(refreshTimer)
        refreshTimer = setTimeout(async () => {
          const kind = payload?.new?.kind || payload?.old?.kind
          if (!['production','quality','deviations','targets'].includes(kind)) return
          try {
            setStatus(`התקבל עדכון ${kind} — מסנכרן רק את הקובץ שהשתנה...`)
            const dataset = await loadCloudDatasetOnce(kind)
            const rows = dataset?.rows || []
            if (kind === 'production') {
  setProduction(
    dedupeRows(rows, productionRowKey).map(row => {
      const sourceDate = row.finishDate || row.date

      return {
        ...row,
        date: sourceDate
          ? new Date(sourceDate)
          : null,
      }
    })
  )
}
            if (kind === 'quality') setQuality(dedupeRows(rows.map(row => row?.__compactQuality && row.date ? { ...row, date:new Date(row.date) } : row), qualityRowKey))
            if (kind === 'deviations') setDeviations(dedupeRows(rows, deviationRawRowKey))
            if (kind === 'targets') {
              try {
                const monthlyArchive = await loadAllMonthlyTargetDatasets()
                setTargets(normalizeStoredTargets(monthlyArchive?.rows?.length ? monthlyArchive.rows : rows))
              } catch (monthlyTargetError) {
                console.warn('Monthly target live sync fallback', monthlyTargetError)
                setTargets(normalizeStoredTargets(rows))
              }
              try {
                const cloudWorkbook = await loadMonthlyTargetWorkbook(planningMonth || monthKey(new Date())).catch(() => null) || await loadActiveTargetWorkbook()
                if (cloudWorkbook?.bytes) await idbSetKey(TARGET_FILE_KEY, cloudWorkbook)
              } catch (workbookError) {
                console.warn('Live target workbook sync skipped', workbookError)
              }
            }
            setDataMeta(current => ({ ...current, [kind]:dataset?.meta || null }))
            setCloudState(current => ({ ...current, mode:'cloud', live:true, lastSync:new Date().toISOString(), message:`עודכן ${kind} בלבד` }))
            setPerformance(current => ({ ...current, queries:current.queries + 1, rows:rows.length, phase:'עדכון חי' }))
            setStatus(`העדכון של ${kind} הושלם`)
          } catch (error) {
            console.warn('Realtime refresh failed', error)
            setCloudState(current => ({ ...current, live:false, message:'מחובר לענן, אך העדכון החי נכשל' }))
          }
        }, 700)
      })
      .on('postgres_changes', { event:'*', schema:'public', table:'iml_monthly_targets' }, () => {
        clearTimeout(refreshTimer)
        refreshTimer = setTimeout(async () => {
          try {
            const monthlyArchive = await loadAllMonthlyTargetDatasets()
            if (monthlyArchive?.rows?.length) {
              setTargets(normalizeStoredTargets(monthlyArchive.rows))
              setDataMeta(current => ({ ...current, targets:normalizeDatasetMeta('targets', monthlyArchive.meta, planningMonth) }))
              setStatus('היסטוריית היעדים החודשית עודכנה מהענן')
            }
          } catch (error) { console.warn('Monthly target realtime refresh failed', error) }
        }, 500)
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
        ['מתקן / Storage Location / PROD LINE', present('Storage Location', 'Storage location', 'PROD LINE', 'Prod Line', 'Production Line')],
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
        ['משאב', present('Resource', 'משאב', 'Storage Location', 'Facility', 'מתקן')],
        ['תוכנית חודשית', present('Plan', 'Monthly Target', 'Monthly Plan', 'יעד חודשי', 'תוכנית חודשית', 'Target')],
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
        const initialDisplayName = forcedKind ? displayDatasetName(forcedKind, forcedKind === 'targets' ? planningMonth : '') : file.name
        setStatus(`קורא את ${initialDisplayName}...`)
        let targetWorkbookOriginal = null
        if (forcedKind === 'targets') {
          const originalBuffer = await file.arrayBuffer()
          targetWorkbookOriginal = {
            name: file.name,
            type: file.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            lastModified: file.lastModified || Date.now(),
            savedAt: new Date().toISOString(),
            bytes: originalBuffer,
          }
          await idbSetKey(TARGET_FILE_KEY, targetWorkbookOriginal)
        }
        const rows = forcedKind === 'targets' ? await readTargetWorkbook(file) : await readWorkbook(file)
        const detected = forcedKind === 'targets' ? 'targets' : classifyFile(rows)
        const kind = forcedKind || detected
        const missing = validateRows(kind, rows)
        if (missing.length) throw new Error(`${file.name}: חסרות עמודות חובה — ${missing.join(', ')}`)
        let storedCount = rows.length
        let rowsForCloud = rows
        let lastFileUniqueRows = rows.length
        if (kind === 'production') {
          const compact = dedupeRows(rows.map(r => {
            const prodLine = normalize(getField(r, ['PROD LINE', 'Prod Line', 'Production Line', 'Production line']))
            const routingGroup = normalizeRouting(getField(r, ['Routing group', 'Routing Group', 'RoutingGroup']))
            const routingDescription = normalize(getField(r, ['Description', 'Routing Description']))
            const assignment = productionAssignment(getField(r, ['Storage Location', 'Storage location']), routingGroup, routingDescription, prodLine)
            return {
            __compactProduction: true,
            facility: assignment.facility,
            prodLine,
            mappedResource: assignment.mapping?.resource || '',
            prodLineTool: assignment.mapping?.tool || '',
            productionDay: localDateOnlyString(getField(r, ['Actual finish date', 'Actual Finish Date'])),
            finishDate: localDateTimeString(
  combineExcelDateTime(
    getField(r, ['Actual finish date', 'Actual Finish Date']),
    getField(r, ['Actual Finish Time', 'Actual finish time']),
    getField(r, ['Release date (actual)', 'Time Stamp'])
  )
),
            qty: num(getField(r, ['Delivered quantity (GMEIN)', 'Confirmed Yield Quantity (GMEIN)', 'Delivered quantity'])),
            plannedQty: num(getField(r, ['Order quantity (GMEIN)', 'Order Quantity (GMEIN)', 'Order quantity', 'Planned quantity', 'Planned Quantity'])),
            order: normalize(getField(r, ['Order', 'Process Order', 'Work Order'])),
            batch: normalize(getField(r, ['Batch', 'Batch Number'])),
            material: normalize(getField(r, ['Material', 'Material #', 'Material Number', 'Material No.', 'מקט', 'מק"ט', 'מק״ט'])),
            desc: normalize(getField(r, ['Material description', 'Material Description'])),
            orderType: normalize(getField(r, ['Order Type'])),
            routingGroup,
            routingDescription,
          }}).filter(r => r.facility && (r.qty || r.order || r.batch)), productionRowKey)
          storedCount = compact.length
          rowsForCloud = compact
          lastFileUniqueRows = compact.length
        }
        else if (kind === 'quality') {
          const compact = dedupeRows(rows.map(r => ({
            __compactQuality: true,
            facility: canonicalFacility(getField(r, ['Inspection Lot Storage Location', 'Process Order Storage Location', 'Storage Location', 'Facility', 'Production Line'])),
            date: combineExcelDateTime(
              getField(r, ['Sample Date', 'Sampling Date', 'Date of Sample', 'Date of Sampling', 'תאריך דגימה', 'Start Date of Inspection', 'Date of Lot Creation', 'Process Order Confirmed Release Date', 'End Date of Inspection', 'Inspection Lot UD Date', 'Process Order Delivered Date']),
              getField(r, ['Sample Time', 'Sampling Time', 'Time of Sample', 'Time of Sampling', 'שעת דגימה', 'Inspection Time', 'Start Time of Inspection', 'Time']),
              getField(r, ['Sample Date Time', 'Sampling Date Time', 'Sample Datetime', 'Sampling Datetime', 'תאריך ושעת דגימה'])
            ),
            batch: normalize(getField(r, ['Batch', 'Batch Number'])), material: normalize(getField(r, [
  'Material #',
  'Material Number',
  'Material No.',
  'מקט',
  'מק"ט',
  'מק״ט',
  'Material'
])),
            order: normalize(getField(r, ['Process Order', 'Process Order #', 'Order'])), status: normalize(getField(r, ['Result Status', 'QA Approval', 'Status'])),
            approval: normalize(getField(r, ['QA Approval'])), inspectionLot: normalize(getField(r, ['Inspection Lot', 'Inspection Lot #'])),
            sampleNo: normalize(getField(r, ['Sample #', 'Sample Number', 'Sample'])),
            operationActivity: normalize(getField(r, ['Operation Activity', 'Operation activity', 'Operation'])),
            operationText: normalize(getField(r, ['Operation short text', 'Operation Short Text'])),
            characteristic: normalize(getField(r, ['Master Insp Charactristic', 'Master Inspection Characteristic'])),
            value: normalize(getField(r, ['Arithmetic Mean of Valid Measured Values'])), lower: normalize(getField(r, ['Lower Specif Limit', 'Lower Spec Limit'])),
            upper: normalize(getField(r, ['Upper Specif Limit', 'Upper Spec Limit'])), unit: normalize(getField(r, ['Unit of Measurement'])),
            line: normalize(getField(r, ['Production Line'])), remarks: normalize(getField(r, ['Charactristic Remarks', 'Characteristic Remarks', 'Batch Remarks'])),
            qualitative: normalize(getField(r, ['Qualitative'])),
            udCode: normalize(getField(r, ['UD Code', 'Usage Decision', 'Usage decision', 'החלטת שימוש'])),
          })).filter(r => r.batch || r.inspectionLot), qualityRowKey)
          lastFileUniqueRows = compact.length
          setStatus(`בודק אילו רשומות איכות חדשות קיימות ב-${displayDatasetName('quality')}...`)
          setUploadProgress({ fileName:displayDatasetName('quality'), kind, phase:'dedupe', percent:0, message:'משווה מול נתוני האיכות הקיימים' })
          const fresh = await filterNewQualityRows(quality, compact, (completed,total) => {
            const percent = total ? Math.round(completed / total * 100) : 100
            setUploadProgress({ fileName:displayDatasetName('quality'), kind, phase:'dedupe', percent, message:'מסנן רשומות שכבר קיימות' })
          })
          storedCount = fresh.length; rowsForCloud = fresh
          if (!fresh.length) {
            loaded.push(`${displayDatasetName('quality')}: לא נמצאו רשומות איכות חדשות`)
            setStatus(`${displayDatasetName('quality')}: כל ${fmt(compact.length)} הרשומות כבר קיימות — לא בוצעה העלאה`)
            continue
          }
        } else if (kind === 'deviations') {
          rowsForCloud = dedupeRows(rows, deviationRawRowKey)
          storedCount = rowsForCloud.length
        }
        else if (kind === 'targets') {
          const fallbackMonth = targetMonthFromTitle(file.name, new Date())
          const parsed = rows.flatMap(r => {
            const sourceResource = normalize(getField(r, ['Resource','משאב','1','Product','Line']))
            const makeTarget = (resource, overrides = {}) => {
              const facilities = targetFacilityIds(resource)
              const facility = facilities[0] || canonicalFacility(getField(r, ['Storage Location','Facility','מתקן']))
              return {
                resource, facility, facilities: facilities.length ? facilities : (facility ? [facility] : []),
                facilityLabel:(resource.match(/\(([^)]+)\)/)?.[1]||'').trim(), descriptionTokens:targetDescriptionTokens(resource),
                productFamily:normalize(r.__family), materials:Array.isArray(r.__materials) ? r.__materials.map(normalize).filter(Boolean) : [],
                routingGroup:normalizeRouting(getField(r,['Routing group','Routing Group','RoutingGroup','קבוצת ניתוב'])),
                station:normalize(getField(r,['Station','Work Center','תחנה']))||facility, lineName:resource,
                month:parseMonth(getField(r,['Month','חודש','Target Month','Plan Month']),fallbackMonth)||fallbackMonth,
                activity:normalize(getField(r,['Activity','Type','סוג פעילות','Production/Packaging']))||'ייצור / אריזה',
                capacity:parseTargetNumber(getField(r,['Capacity','קיבולת','Monthly Capacity','קיבולת חודשית'])) * 1000,
                target:parseTargetNumber(getField(r,['Plan','Monthly Target','Monthly Plan','יעד חודשי','תוכנית חודשית','Target'])) * 1000,
                fileProduction:parseTargetNumber(getField(r,['Production','ייצור'])), fileAchievement:parseTargetNumber(getField(r,['% Achievement','Achievement'])),
                requiredPerDay:parseTargetNumber(getField(r,['Req. t/d','Required t/d'])), lastDay:parseTargetNumber(getField(r,['Last day'])),
                adjustedRequiredPerDay:parseTargetNumber(getField(r,['Adjusted Req. t/d'])), actualPerDay:parseTargetNumber(getField(r,['Actual t/d'])),
                ratePercent:parseTargetNumber(getField(r,['%Rate','Rate'])), adherence:parseTargetNumber(getField(r,['New Adherence'])),
                recyclingPlan:parseTargetNumber(getField(r,['Recycling Plan'])), recycled:parseTargetNumber(getField(r,['Recycled'])),
                forPacking:parseTargetNumber(getField(r,['For Packing'])), restrictedRecycling:parseTargetNumber(getField(r,['Restricted - recycling'])),
                restrictedDisposal:parseTargetNumber(getField(r,['Restricted – risk for disposal','Restricted - risk for disposal'])),
                average:parseTargetNumber(getField(r,['AVERAGE','Average'])), notes:normalize(getField(r,['Notes','Remarks','הערות'])),
                ...overrides
              }
            }
            if (/^SHAKED\s+ISO\s*\(42\+23\)$/i.test(sourceResource)) {
              return [
                makeTarget('Shaked iso 42', { facility:'1142', facilities:['1142'], station:'1142', capacity:0, target:0 }),
                makeTarget('Shaked iso 23', { facility:'1123', facilities:['1123'], station:'1123', capacity:0, target:0 })
              ]
            }
            return [makeTarget(sourceResource)]
          }).filter(r => r.resource && (r.target > 0 || r.capacity > 0))
            .map(targetRow => {
              // Dynamic Targets v1: the monthly workbook is the source of truth.
              // New rows are never blocked by a hard-coded resource whitelist.
              // If the row name does not contain a known station, infer station(s)
              // from current production using the DATA-sheet material list.
              if (targetRow.facilities?.length) return targetRow
              const materialSet = new Set((targetRow.materials || []).map(normalize).filter(Boolean))
              if (!materialSet.size) return { ...targetRow, mappingStatus:'requires-mapping', mappingReason:'לא נמצאו מק״טים למשפחה בגיליון DATA' }

              // Dynamic Mapping V3: infer only when the evidence is unambiguous.
              // Group 11xx/15xx aliases into the same production family, ignore the two
              // dedicated bulk streams (1142+999 and 1119+777), and never guess when
              // the same target materials are reported in more than one station family.
              const evidence = production.filter(prodRow => {
                if (!materialSet.has(normalize(prodRow.material))) return false
                const station = normalize(prodRow.facility)
                const markerText = normalize(`${prodRow.desc || ''} ${prodRow.routingDescription || ''}`)
                if (station === '1142' && /(^|\D)999(\D|$)/.test(markerText)) return false
                if (station === '1119' && /(^|\D)777(\D|$)/.test(markerText)) return false
                return true
              })
              const familyCounts = new Map()
              evidence.forEach(prodRow => {
                const family = stationFamily(prodRow.facility)
                if (!family) return
                familyCounts.set(family, (familyCounts.get(family) || 0) + 1)
              })
              const candidates = [...familyCounts.entries()].sort((a,b) => b[1]-a[1])
              if (candidates.length !== 1) {
                return {
                  ...targetRow,
                  mappingStatus:'requires-mapping',
                  mappingCandidates:candidates.map(([facility, rows]) => ({ facility, rows })),
                  mappingReason:candidates.length ? 'נמצאה תפוקה ביותר ממשפחת תחנה אחת — נדרש אישור מנהל' : 'לא נמצאה תפוקה תואמת למק״טים בקובץ הכמויות'
                }
              }
              const inferredFacility = candidates[0][0]
              return {
                ...targetRow,
                facility:inferredFacility,
                facilities:[inferredFacility],
                station:inferredFacility,
                mappingStatus:'inferred',
                mappingConfidence:'high',
                mappingReason:`זוהה אוטומטית לפי מק״טים בגיליון DATA ותפוקה במשפחת תחנה ${inferredFacility}`
              }
            })
          if(!parsed.length) throw new Error(`${file.name}: לא נמצאו שורות יעד תקינות. ודא שקיימות עמודות Resource, Capacity ו-Plan.`)
          storedCount=parsed.length; rowsForCloud=parsed; if(parsed[0]?.month) setPlanningMonth(parsed[0].month)
        } else throw new Error(`${file.name}: סוג הקובץ לא זוהה. השתמש באזור הטעינה המתאים במרכז הנתונים.`)
        const facilitiesFound = new Set(rows.slice(0, 5000).map(r => canonicalFacility(getField(r, ['Storage Location','Inspection Lot Storage Location','Process Order Storage Location','Facility','Production Line','מתקן']))).filter(Boolean)).size
        const displayName = displayDatasetName(kind, kind === 'targets' ? rowsForCloud?.[0]?.month : '')
        const nextMeta = { fileName:displayName, originalFileName:file.name, rows:storedCount, rawRows:rows.length, lastFileRows:rows.length, lastFileUniqueRows, loadedAt:new Date().toISOString(), facilities:facilitiesFound, valid:true, source:'cloud' }
        setStatus(`מעלה את ${displayName} למסד המשותף...`)
        setUploadProgress({ fileName:displayName, kind, phase:'prepare', percent:0, message:'מכין את הנתונים' })
        const progressHandler = progress => {
          setUploadProgress({ fileName:displayName, kind, ...progress })
          setStatus(`${displayName}: ${progress.message} (${progress.percent}%)`)
        }
        const savedMeta = kind === 'quality'
          ? await uploadCloudDatasetIncremental(kind, rowsForCloud, { ...nextMeta, existingRows:quality.length }, currentUser, progressHandler)
          : await uploadCloudDataset(kind, rowsForCloud, nextMeta, currentUser, progressHandler)
        if (kind === 'production') setProduction(rowsForCloud)
        else if (kind === 'quality') setQuality(current => dedupeRows([...current, ...rowsForCloud], qualityRowKey))
        else if (kind === 'deviations') setDeviations(rowsForCloud)
        else if (kind === 'targets') {
          const targetMonth = rowsForCloud?.[0]?.month || planningMonth
          try {
            await saveMonthlyTargetDataset(targetMonth, rowsForCloud, { ...nextMeta, ...savedMeta, originalFileName:file.name }, currentUser)
            const monthlyArchive = await loadAllMonthlyTargetDatasets()
            setTargets(normalizeStoredTargets(monthlyArchive?.rows?.length ? monthlyArchive.rows : rowsForCloud))
          } catch (monthlyTargetError) {
            console.warn('Monthly target archive save failed', monthlyTargetError)
            throw monthlyTargetError
          }
          if (targetWorkbookOriginal) {
            try {
              const workbookPayload = { ...targetWorkbookOriginal, versionId:savedMeta?.versionId || '' }
              const monthlyWorkbookMeta = await saveMonthlyTargetWorkbook(targetMonth, workbookPayload, currentUser)
              // Keep the legacy singleton updated as a compatibility fallback.
              await saveActiveTargetWorkbook(workbookPayload, currentUser).catch(() => null)
              await idbSetKey(TARGET_FILE_KEY, {
                ...targetWorkbookOriginal,
                month:targetMonth,
                targetVersionId:monthlyWorkbookMeta?.target_version_id || savedMeta?.versionId || '',
                savedAt:monthlyWorkbookMeta?.updated_at || targetWorkbookOriginal.savedAt,
                source:'cloud-monthly',
              })
            } catch (workbookError) {
              console.warn('Target workbook monthly cloud sync failed', workbookError)
              loaded.push(`אזהרה: יעדי ${targetMonth} נשמרו, אך סנכרון קובץ ה-Excel המלא נכשל (${workbookError?.message || 'שגיאה'})`)
            }
          }
        }
        setDataMeta(current => ({ ...current, [kind]: savedMeta }))
        setCloudState({ mode:'cloud', lastSync:savedMeta.loadedAt, message:'מחובר ומסונכרן עם Supabase', latencyMs:cloudState.latencyMs, live:true })
        loaded.push(`${displayName}: ${fmt(storedCount)} רשומות בענן`)
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
      const assignment = productionAssignment(r.facility, r.routingGroup, r.routingDescription, r.prodLine)
      return {
        facility: assignment.facility,
        prodLine: normalize(r.prodLine),
        mappedResource: normalize(r.mappedResource || assignment.mapping?.resource),
        prodLineTool: normalize(r.prodLineTool || assignment.mapping?.tool),
        productionDay: normalize(r.productionDay) || iso(r.finishDate),
        date: productionDateFromDay(r.productionDay) || (r.finishDate ? new Date(r.finishDate) : null),
        qty: num(r.qty),
        plannedQty: num(r.plannedQty),
        order: normalize(r.order),
        batch: normalize(r.batch),
        material: normalize(r.material),
        desc: normalize(r.desc),
        orderType: normalize(r.orderType),
        routingGroup: normalizeRouting(r.routingGroup),
        routingDescription: normalize(r.routingDescription),
        hour: r.finishDate ? new Date(r.finishDate).getHours() : null,
        shift: shiftInfo(r.finishDate ? new Date(r.finishDate) : null),
      }
    }
    const finish = combineExcelDateTime(
      getField(r, ['Actual finish date', 'Actual Finish Date']),
      getField(r, ['Actual Finish Time', 'Actual finish time']),
      getField(r, ['Release date (actual)', 'Time Stamp'])
    )
    const prodLine = normalize(getField(r, ['PROD LINE', 'Prod Line', 'Production Line', 'Production line']))
    const routingGroup = normalizeRouting(getField(r, ['Routing group', 'Routing Group', 'RoutingGroup']))
    const routingDescription = normalize(getField(r, ['Description', 'Routing Description']))
    const assignment = productionAssignment(getField(r, ['Storage Location', 'Storage location']), routingGroup, routingDescription, prodLine)
    return {
      facility: assignment.facility,
      prodLine,
      mappedResource: assignment.mapping?.resource || '',
      prodLineTool: assignment.mapping?.tool || '',
      productionDay: localDateOnlyString(getField(r, ['Actual finish date', 'Actual Finish Date'])),
      date: productionDateFromDay(localDateOnlyString(getField(r, ['Actual finish date', 'Actual Finish Date']))) || finish,
      qty: num(getField(r, ['Delivered quantity (GMEIN)', 'Confirmed Yield Quantity (GMEIN)', 'Delivered quantity'])),
      plannedQty: num(getField(r, ['Order quantity (GMEIN)', 'Order Quantity (GMEIN)', 'Order quantity', 'Planned quantity', 'Planned Quantity'])),
      order: normalize(getField(r, ['Order', 'Process Order', 'Work Order'])),
      batch: normalize(getField(r, ['Batch', 'Batch Number'])),
      material: normalize(getField(r, [
  'Material #',
  'Material Number',
  'Material No.',
  'מקט',
  'מק"ט',
  'מק״ט',
  'Material'
])),
      desc: normalize(getField(r, ['Material description', 'Material Description'])),
      orderType: normalize(getField(r, ['Order Type'])),
      routingGroup,
      routingDescription,
      hour: finish ? finish.getHours() : null,
      shift: shiftInfo(finish),
    }
  }).filter(r => r.facility), [production])
const materialByBatchDescription = useMemo(() => {
  const map = new Map()

  prod.forEach(row => {
    const batch = normalize(row.batch)
    const material = normalize(row.material)
    const desc = normalize(row.desc).toLowerCase()

    if (!batch || !material || !desc) return

    map.set(`${batch}|${desc}`, material)
  })

  return map
}, [prod])
  const qualityRows = useMemo(() => quality.map(r => {
  if (r.__compactQuality) {
    const repairedMaterial =
      materialByBatchDescription.get(
        `${normalize(r.batch)}|${normalize(r.material).toLowerCase()}`
      )

    return {
      ...r,
      material: repairedMaterial || r.material
    }
  }

  return ({
    facility: canonicalFacility(getField(r, ['Inspection Lot Storage Location', 'Process Order Storage Location', 'Storage Location', 'Facility', 'Production Line'])),
    date: combineExcelDateTime(
      getField(r, ['Sample Date', 'Sampling Date', 'Date of Sample', 'Date of Sampling', 'תאריך דגימה', 'Start Date of Inspection', 'Date of Lot Creation', 'Process Order Confirmed Release Date', 'End Date of Inspection', 'Inspection Lot UD Date', 'Process Order Delivered Date']),
      getField(r, ['Sample Time', 'Sampling Time', 'Time of Sample', 'Time of Sampling', 'שעת דגימה', 'Inspection Time', 'Start Time of Inspection', 'Time']),
      getField(r, ['Sample Date Time', 'Sampling Date Time', 'Sample Datetime', 'Sampling Datetime', 'תאריך ושעת דגימה'])
    ),
    batch: normalize(getField(r, ['Batch', 'Batch Number'])), material: normalize(getField(r, [
  'Material #',
  'Material Number',
  'Material No.',
  'מקט',
  'מק"ט',
  'מק״ט',
  'Material'
])),
 order: normalize(getField(r, ['Process Order', 'Process Order #', 'Order'])),
status: normalize(getField(r, ['Result Status', 'QA Approval', 'Status'])),
inspectionLot: normalize(getField(r, ['Inspection Lot', 'Inspection Lot #'])),
udCode: normalize(getField(r, ['UD Code', 'Usage Decision', 'Usage decision', 'החלטת שימוש']))
  })
}), [quality, materialByBatchDescription])

  const deviationRows = useMemo(() => deviations.map(r => ({
    facility: canonicalFacility(getField(r, ['Facility', 'Production Line', 'Storage Location'])),
    date: excelDate(getField(r, ['Date of Lot Creation', 'Inspection Lot UD Date', 'Process Order Delivered Date', 'Start Date of Inspection'])),
    batch: normalize(getField(r, ['Batch', 'Batch Number'])),
material: normalize(getField(r, [
  'Material #',
  'Material Number',
  'Material No.',
  'מקט',
  'מק"ט',
  'מק״ט',
  'Material'
])),
    inspectionLot: normalize(getField(r, ['Inspection Lot', 'Inspection Lot #'])),
    status: normalize(getField(r, ['QA Status', 'Status'])),
    rejectedCount: num(getField(r, ['Rejected characteristics', 'Rejected characteristics '])),
    remarks: normalize(getField(r, ['UD Remarks', 'Remarks', 'Batch Remarks'])),
    udCode: normalize(getField(r, ['UD Code'])),
  })), [deviations])

  // Normalized rows that are allowed to enter the dashboard. Raw uploaded rows
  // remain in the cloud dataset, but unrelated facilities are excluded from all
  // regular user-facing calculations until they are explicitly added to the picker.
  const dashboardProd = useMemo(() => prod.filter(row => selectableFacilitySet.has(String(row.facility || ''))), [prod, selectableFacilitySet])
  const dashboardQualityRows = useMemo(() => qualityRows.filter(row => selectableFacilitySet.has(String(row.facility || ''))), [qualityRows, selectableFacilitySet])
  const dashboardDeviationRows = useMemo(() => deviationRows.filter(row => selectableFacilitySet.has(String(row.facility || ''))), [deviationRows, selectableFacilitySet])

  // Sprint 11.9.36 performance: index detailed quality characteristics only for
  // deviation rows currently needed plus the Batch card currently opened. The quality
  // dataset can exceed 800K result rows, so building full arrays for every Batch during
  // initial render is unnecessary and can make Chrome report "Page Unresponsive".
  const selectedQualityKey = useMemo(() => {
    const batch = normalize(selectedBatch)
    if (!batch) return ''
    const requestedMaterial = normalize(selectedBatchMaterial)
    if (requestedMaterial) return batchMaterialKey(batch, requestedMaterial)
    const materials = [...new Set(dashboardProd.filter(row => normalize(row.batch) === batch).map(row => normalize(row.material)).filter(Boolean))]
    return materials.length === 1 ? batchMaterialKey(batch, materials[0]) : ''
  }, [selectedBatch, selectedBatchMaterial, dashboardProd])

  const qualityDetailKeys = useMemo(() => {
    const keys = new Set(dashboardDeviationRows.map(row => batchMaterialKey(row.batch, row.material)).filter(Boolean))
    if (selectedQualityKey) keys.add(selectedQualityKey)
    return keys
  }, [dashboardDeviationRows, selectedQualityKey])

  const qualityIndex = useMemo(() => {
    const byBatch = new Map()
    const byBatchMaterial = new Map()
    const rejected = new Map()
    const approved = new Map()
    const latestByBatchMaterial = new Map()
    const latestByBatchMaterialLot = new Map()
    const seenRejected = new Map()
    const seenApproved = new Map()

    const addCharacteristic = (target, seenMap, key, item) => {
      if (!key || !item.characteristic) return
      const signature = `${item.characteristic}|${item.value}|${item.inspectionLot}`
      let seen = seenMap.get(key)
      if (!seen) { seen = new Set(); seenMap.set(key, seen) }
      if (seen.has(signature)) return
      seen.add(signature)
      const list = target.get(key) || []
      list.push(item)
      target.set(key, list)
    }

    dashboardQualityRows.forEach(row => {
      const key = batchMaterialKey(row.batch, row.material)
      if (!key || !qualityDetailKeys.has(key)) return

      const list = byBatchMaterial.get(key) || []
      list.push(row)
      byBatchMaterial.set(key, list)

      const timestamp = row.date ? new Date(row.date).getTime() : 0
      if (timestamp > (latestByBatchMaterial.get(key)?.timestamp || 0)) {
        latestByBatchMaterial.set(key, { timestamp, date: row.date })
      }
      if (row.inspectionLot) {
        const lotKey = `${key}|${normalize(row.inspectionLot)}`
        if (timestamp > (latestByBatchMaterialLot.get(lotKey)?.timestamp || 0)) {
          latestByBatchMaterialLot.set(lotKey, { timestamp, date: row.date })
        }
      }
      const status = normalize(row.status || row.approval).toLowerCase()
      const isRejected = ['rejection', 'rejected', 'fail', 'failed', 'פסול', 'לא תקין', 'חריג'].some(x => status.includes(x))
      const item = {
        characteristic: row.characteristic,
        value: row.value,
        lower: row.lower,
        upper: row.upper,
        unit: row.unit,
        remarks: row.remarks,
        qualitative: row.qualitative,
        inspectionLot: row.inspectionLot,
        date: row.date,
      }
      addCharacteristic(isRejected ? rejected : approved, isRejected ? seenRejected : seenApproved, key, item)
    })

    return {
    byBatch,
    byBatchMaterial,
    rejected,
    approved,
    latestByBatchMaterial,
    latestByBatchMaterialLot,
}
  }, [dashboardQualityRows, qualityDetailKeys])

  const enrichedDeviationRows = useMemo(() => dashboardDeviationRows.map(row => {
    const key = batchMaterialKey(row.batch, row.material)
    const lotKey = key && row.inspectionLot ? `${key}|${normalize(row.inspectionLot)}` : ''
    const sampleDate = key
      ? ((lotKey && qualityIndex.latestByBatchMaterialLot.get(lotKey)?.date) || qualityIndex.latestByBatchMaterial.get(key)?.date || null)
      : null
    return {
      ...row,
      sampleDate,
      rejectedCharacteristics: key ? (qualityIndex.rejected.get(key) || []) : [],
      approvedCharacteristics: key ? (qualityIndex.approved.get(key) || []) : [],
    }
  }), [dashboardDeviationRows, qualityIndex])

  // Plant rule: a quality record is uniquely identified by exact Batch + Material.
  // Order, facility, routing group and inspection lot remain display fields only.
  const selectedBatchData = useMemo(() => {
    const batch = normalize(selectedBatch)
    const requestedMaterial = normalize(selectedBatchMaterial)
    if (!batch) return null

    const batchProductionRows = dashboardProd.filter(row => normalize(row.batch) === batch)
    const batchMaterials = [...new Set(batchProductionRows.map(row => normalize(row.material)).filter(Boolean))]
    const material = requestedMaterial || (batchMaterials.length === 1 ? batchMaterials[0] : '')
    const key = batchMaterialKey(batch, material)
    const productionRows = material
      ? batchProductionRows.filter(row => normalize(row.material) === material)
      : batchProductionRows

    const qualityForBatchMaterial = key ? (qualityIndex.byBatchMaterial.get(key) || []) : []
    const deviationForBatchMaterial = key
      ? enrichedDeviationRows.filter(row => batchMaterialKey(row.batch, row.material) === key)
      : []

    return {
      batch,
      material,
      production: productionRows,
      quality: qualityForBatchMaterial,
      deviations: deviationForBatchMaterial,
    }
  }, [selectedBatch, selectedBatchMaterial, dashboardProd, qualityIndex, enrichedDeviationRows])

  const openBatchCard = (batch, material = '') => {
    if (!batch) return
    setSelectedBatch(normalize(batch))
    setSelectedBatchMaterial(normalize(material))
  }

  const dataMonths = useMemo(() => [...new Set(dashboardProd.map(r => monthKey(r.date)).filter(Boolean))].sort(), [dashboardProd])
  const targetMonths = useMemo(() => [...new Set(targets.map(r => r.month).filter(Boolean))].sort(), [targets])
  const availableMonths = useMemo(() => [...new Set([...targetMonths, ...dataMonths])].sort().reverse(), [targetMonths, dataMonths])
  useEffect(() => { if (!planningMonth && availableMonths.length) setPlanningMonth(availableMonths[0]) }, [availableMonths, planningMonth])

  const dateBounds = useMemo(() => {
    let minTime = Infinity
    let maxTime = -Infinity
    const scan = rows => rows.forEach(row => {
      const time = row.date ? new Date(row.date).getTime() : NaN
      if (!Number.isFinite(time)) return
      if (time < minTime) minTime = time
      if (time > maxTime) maxTime = time
    })
    scan(dashboardProd); scan(dashboardQualityRows); scan(dashboardDeviationRows)
    return {
      min: Number.isFinite(minTime) ? iso(new Date(minTime)) : '',
      max: Number.isFinite(maxTime) ? iso(new Date(maxTime)) : '',
    }
  }, [dashboardProd, dashboardQualityRows, dashboardDeviationRows])

  const baseFiltered = useMemo(() => {
    const q = query.trim().toLowerCase()
    return dashboardProd.filter(r => {
      const day = r.productionDay || iso(r.date)
      const inRange = (!from || day >= from) && (!to || day <= to)
      return inRange && (!q || [r.facility, r.prodLine, r.prodLineTool, r.order, r.batch, r.material, r.desc].some(v => String(v || '').toLowerCase().includes(q)))
    })
  }, [dashboardProd, from, to, query])
  const filtered = useMemo(() => baseFiltered.filter(r => !selectedFacilities.length || selectedFacilities.includes(r.facility)), [baseFiltered, selectedFacilities])

  // Dedicated Facility 42 material balance. This view is intentionally independent
  // of Shaked ISO 42 targets/cards: bulk input = station 1142 + description marker 999;
  // packed output = the approved 1L/5L/10-20L Facility 42 packaging routes.
  const facility42Balance = useMemo(() => {
    const inRange = prod.filter(r => {
      const day = r.productionDay || iso(r.date)
      return (!from || day >= from) && (!to || day <= to)
    })
    // Use SAP Delivered quantity only (stored in r.qty). Keep bulk and packed output
    // mutually exclusive: bulk = 1142 + marker 999; packed = 1542 + ZFIN + approved LQ routes.
    const bulkRows = inRange.filter(r =>
      normalize(r.facility) === '1142' && /999/.test(normalize(r.desc))
    )
    const packedRows = inRange.filter(r =>
      normalize(r.facility) === '1542' &&
      normalize(r.orderType).toUpperCase() === 'ZFIN' &&
      isFacility42PackagingRoute(r.routingGroup, r.routingDescription)
    )
    const routeBucket = row => {
      const route = normalize(`${row.routingGroup || ''} ${row.routingDescription || ''}`).toUpperCase()
      if (/(^|\s)LQ-P-1(\s|$)/.test(route) || route.includes('42-P-02') || route.includes('LIQUID 1 LITER')) return '1L'
      if (/(^|\s)LQ-P-5(\s|$)/.test(route) || route.includes('42-P-03') || route.includes('LIQUID 5 LITER')) return '5L'
      return '10/20L'
    }
    // Residues are reported to 1542 as ZSEM and explicitly carry a 200L/1000L
    // packaging description. They are output from the bulk and must therefore
    // be deducted from the balance, but must NOT be counted as line packaging.
    const residuePattern = /(^|\\D)(200|1000)\\s*(L|LT|LTR|LITER|LITRE)(\\D|$)/i
    const residueRows = inRange.filter(r =>
      normalize(r.facility) === '1542' &&
      normalize(r.orderType).toUpperCase() === 'ZSEM' &&
      residuePattern.test(normalize(`${r.desc || ''} ${r.routingDescription || ''}`))
    )
    const bulk = bulkRows.reduce((sum,r)=>sum+num(r.qty),0)
    const byLine = {'1L':0,'5L':0,'10/20L':0}
    packedRows.forEach(r => { byLine[routeBucket(r)] += num(r.qty) })
    const packed = Object.values(byLine).reduce((a,b)=>a+b,0)
    const residues = residueRows.reduce((sum,r)=>sum+num(r.qty),0)
    const balance = bulk - packed - residues
    const utilization = bulk > 0 ? (packed + residues) / bulk * 100 : 0
    return { bulkRows, packedRows, residueRows, bulk, byLine, packed, residues, balance, utilization }
  }, [prod, from, to])

  // Dedicated Facility 19 material balance. Bulk is identified by the approved
  // business marker 777 in the material description. Bulk is reported strictly
  // at station 1119. Packaged output is reported at station 1519 and is split
  // between normal WG production and small packs. Keeping 1119 and 1519 separate
  // prevents bulk rows from being mixed with packaged production.
  const facility19Balance = useMemo(() => {
    const inRange = prod.filter(r => {
      const day = r.productionDay || iso(r.date)
      return (!from || day >= from) && (!to || day <= to)
    })
    const has777 = row => /(^|\D)777(\D|$)/.test(normalize(`${row.desc || ''} ${row.routingDescription || ''}`))
    const stationFamily19 = row => {
      const digits = normalize(row.facility).replace(/\D/g, '')
      return digits.length >= 2 && digits.slice(-2) === '19'
    }
    const isSmallPack = row => {
      const route = normalize(`${row.routingGroup || ''} ${row.routingDescription || ''} ${row.resource || ''} ${row.line || ''}`).toUpperCase()
      return ['19PWG-01','19PWG-05','19PWG-15'].some(code => route.includes(code))
    }
    const bulkRows = inRange.filter(r => normalize(r.facility) === '1119' && has777(r))
    const packedRows = inRange.filter(r => normalize(r.facility) === '1519' && !has777(r))
    const byType = {'WG':0,'SMALL PACKS':0}
    packedRows.forEach(r => { byType[isSmallPack(r) ? 'SMALL PACKS' : 'WG'] += num(r.qty) })
    const bulk = bulkRows.reduce((sum,r)=>sum+num(r.qty),0)
    const packed = Object.values(byType).reduce((a,b)=>a+b,0)
    const balance = bulk - packed
    const utilization = bulk > 0 ? packed / bulk * 100 : 0
    return { bulkRows, packedRows, bulk, byType, packed, balance, utilization, isSmallPack }
  }, [prod, from, to])


  const udByBatchMaterial = useMemo(() => {
    const exact = new Map()
    const byBatch = new Map()
    const put = (row) => {
      const batch = normalize(row?.batch)
      const material = normalize(row?.material)
      const ud = normalize(row?.udCode || row?.usageDecision)
      if (!batch || !ud) return
      if (material) exact.set(`${batch}|${material}`, ud)
      if (!byBatch.has(batch)) byBatch.set(batch, ud)
    }
    dashboardQualityRows.forEach(put)
    dashboardDeviationRows.forEach(put)
    return { exact, byBatch }
  }, [dashboardQualityRows, dashboardDeviationRows])

  const productionUsageDecision = row => {
    const batch = normalize(row?.batch)
    const material = normalize(row?.material)
    if (!batch) return '—'
    return udByBatchMaterial.exact.get(`${batch}|${material}`) || udByBatchMaterial.byBatch.get(batch) || '—'
  }

  const productionSortValue = (row, key) => {
    if (key === 'date') return row?.date ? new Date(row.date).getTime() : 0
    if (key === 'ud') return productionUsageDecision(row)
    if (key === 'facility') return normalize(row?.facility)
    if (key === 'routingGroup') return normalize(row?.routingGroup)
    if (key === 'order') return normalize(row?.order)
    if (key === 'batch') return normalize(row?.batch)
    if (key === 'material') return normalize(row?.material)
    if (key === 'desc') return normalize(row?.desc)
    if (key === 'qty') return Number(row?.qty) || 0
    return ''
  }

  const sortedRecentProduction = useMemo(() => {
    const rows = filtered.slice()
    const { key, direction } = productionSort
    const factor = direction === 'asc' ? 1 : -1
    rows.sort((a,b) => {
      const av = productionSortValue(a,key), bv = productionSortValue(b,key)
      if (typeof av === 'number' && typeof bv === 'number') return (av-bv)*factor
      return String(av).localeCompare(String(bv), 'he', { numeric:true, sensitivity:'base' }) * factor
    })
    return rows.slice(0,200)
  }, [filtered, productionSort, udByBatchMaterial])

  const toggleProductionSort = key => setProductionSort(current => ({
    key,
    direction: current.key === key && current.direction === 'asc' ? 'desc' : 'asc'
  }))
  const productionSortArrow = key => productionSort.key === key ? (productionSort.direction === 'asc' ? ' ↑' : ' ↓') : ' ↕'


  const discoveredFacilities = useMemo(() => [...new Set([...targets.map(t => t.facility), ...prod.map(r => r.facility)].filter(Boolean))].sort(), [targets, prod])
  const optionalFacilities = useMemo(() => discoveredFacilities.filter(id => !PRIMARY_FACILITIES.includes(id) && !additionalFacilities.includes(id)), [discoveredFacilities, additionalFacilities])
  // Daily Management default: never show 11xx stations automatically.
  // They remain available in the manual "+ add facility" list.
  const dailyCoreFacilities = useMemo(
    () => PRIMARY_FACILITIES.filter(id => !String(id).startsWith('11')),
    []
  )
  const dailyFacilities = useMemo(
    () => [...new Set([...dailyCoreFacilities, ...dailyAdditionalFacilities])],
    [dailyCoreFacilities, dailyAdditionalFacilities]
  )
  const dailyOptionalFacilities = useMemo(
    () => discoveredFacilities.filter(id => !dailyFacilities.includes(id)),
    [discoveredFacilities, dailyFacilities]
  )
  const availableYears = useMemo(() => [...new Set(dashboardProd.map(r => r.date?.getFullYear()).filter(Boolean))].sort((a,b) => b-a), [dashboardProd])

  const targetsWithAdminMappings = useMemo(() => targets.map(target => {
    const resourceKey = normalize(target.resource).toUpperCase()
    // Repair Facility 24 targets that were uploaded before the explicit 1524
    // mapping existed. This also fixes already-saved cloud target versions.
    const normalizedTarget = /^24F(?:128)?$/.test(resourceKey)
      ? { ...target, facility:'1524', facilities:['1524'], station:'1524', mappingStatus:'business-approved', mappingReason:'24F / 24F128 משויכים למתקן 1524' }
      : target
    const saved = targetMappings.find(mapping => normalize(mapping.resource).toUpperCase() === resourceKey && (!mapping.month || mapping.month === target.month))
    if (!saved?.family) return normalizedTarget
    return { ...normalizedTarget, facility:saved.family, facilities:[saved.family], station:saved.family, mappingStatus:'manual-approved', mappingReason:`מיפוי מנהל מאושר למשפחת תחנה ${saved.family}` }
  }), [targets, targetMappings])

  const dashboardTargetsWithAdminMappings = useMemo(() => targetsWithAdminMappings.filter(target => {
    const ids = (target.facilities || [target.facility]).map(String).filter(Boolean)
    return ids.some(id => selectableFacilitySet.has(id))
  }), [targetsWithAdminMappings, selectableFacilitySet])

  const planningRows = useMemo(() => buildResourceRows({
    production: dashboardProd,
    targets:dashboardTargetsWithAdminMappings,
    planningMonth,
    fallbackFacilities: facilities,
    manualMappings,
  }), [planningMonth, dashboardProd, dashboardTargetsWithAdminMappings, facilities, manualMappings])


  const mappingTargets = useMemo(() => planningRows.map(row => ({
    key: row.targetKey,
    resource: row.resource,
    facility: row.facility,
    facilities: row.facilities,
    target: row.target,
    actual: row.actual,
    pct: row.pct,
  })), [planningRows])

  const addMappingEvent = (action, mapping, note = '') => {
    setMappingTimeline(current => [{
      id:`${Date.now()}-${Math.random()}`,
      at:new Date().toISOString(), action, note,
      mappingId:mapping?.id || '', material:mapping?.material || '',
      family:mapping?.family || '', targetResource:mapping?.targetResource || '',
      user:currentUser?.email || currentUser?.user_metadata?.full_name || 'מנהל',
    }, ...current].slice(0,500))
  }

  const syncMappingToCloud = async mapping => {
    if (!supabase || !currentUser?.id) return null
    const payload = {
      row_key:mapping.rowKey, material:mapping.material || '', description:mapping.description || '',
      facility_family:mapping.family, target_key:mapping.targetKey, target_resource:mapping.targetResource,
      status:mapping.status, active:mapping.active !== false, created_by:currentUser.id,
      created_by_email:currentUser.email || '', approved_by:mapping.status === 'approved' ? currentUser.id : null,
      approved_at:mapping.status === 'approved' ? new Date().toISOString() : null,
      updated_at:new Date().toISOString(),
    }
    if (mapping.id?.length === 36) payload.id = mapping.id
    const { data, error } = await supabase.from('iml_product_mappings').upsert(payload).select().maybeSingle()
    if (error) return null
    return data
  }

  const openMappingDialog = item => {
    const candidates = mappingTargets.filter(target => (target.facilities || []).map(stationFamily).includes(item.family))
    setMappingDialog({ ...item, candidates })
    setMappingTargetKey(candidates[0]?.key || '')
    setMappingMessage('')
  }

  const savePendingMapping = async () => {
    if (!mappingDialog || !mappingTargetKey) return
    const target = mappingTargets.find(item => item.key === mappingTargetKey)
    if (!target) return
    const row = mappingDialog.row
    const mapping = {
      id:globalThis.crypto?.randomUUID?.() || `${Date.now()}-${Math.random()}`,
      rowKey:productionMappingKey(row), material:row.material || '', description:row.desc || '',
      family:mappingDialog.family, targetKey:mappingTargetKey, targetResource:target.resource,
      status:'pending', active:true, createdBy:currentUser?.email || '', createdAt:new Date().toISOString(),
    }
    setManualMappings(current => [mapping, ...current.filter(item => item.rowKey !== mapping.rowKey || item.status === 'rejected')])
    addMappingEvent('created', mapping, 'נוצר שיוך ממתין לאישור')
    const cloudRow = await syncMappingToCloud(mapping)
    if (cloudRow?.id) setManualMappings(current => current.map(item => item.id === mapping.id ? { ...item, id:cloudRow.id, cloud:true } : item))
    setMappingDialog(null)
    setMappingMessage('השיוך נשמר כממתין לאישור')
  }

  const updateMappingStatus = async (mapping, status) => {
    const updated = { ...mapping, status, active:status !== 'rejected', approvedAt:status === 'approved' ? new Date().toISOString() : mapping.approvedAt }
    setManualMappings(current => current.map(item => item.id === mapping.id ? updated : item))
    addMappingEvent(status, updated, status === 'approved' ? 'השיוך אושר והשפיע על הדשבורד' : 'השיוך נדחה')
    await syncMappingToCloud(updated)
  }

  const rollbackMapping = async mapping => {
    const updated = { ...mapping, active:false, status:'rejected' }
    setManualMappings(current => current.map(item => item.id === mapping.id ? updated : item))
    addMappingEvent('rollback', updated, 'השיוך בוטל והחישוב חזר למצב הקודם')
    await syncMappingToCloud(updated)
  }


  const availableTargetFamilies = useMemo(() => [...new Set(dashboardProd.map(row => stationFamily(row.facility)).filter(Boolean))].sort(), [dashboardProd])
  const unmappedPlanningTargets = useMemo(() => planningRows.filter(row => row.mappingStatus === 'requires-mapping'), [planningRows])

  const openTargetMappingDialog = row => {
    setTargetMappingDialog(row)
    const firstCandidate = row.mappingCandidates?.[0]?.facility || availableTargetFamilies[0] || ''
    setTargetMappingFamily(firstCandidate)
    setMappingMessage('')
  }

  const saveTargetMapping = async () => {
    if (!targetMappingDialog || !targetMappingFamily || !supabase || !currentUser?.id) return
    const payload = {
      target_resource:targetMappingDialog.resource, target_month:planningMonth || '', facility_family:targetMappingFamily,
      active:true, created_by:currentUser.id, created_by_email:currentUser.email || '', updated_at:new Date().toISOString(),
    }
    const { data, error } = await supabase.from('iml_target_mappings').upsert(payload, { onConflict:'target_resource,target_month' }).select().maybeSingle()
    if (error) { setMappingMessage(`שמירת המיפוי נכשלה: ${error.message}`); return }
    const mapped = { id:data?.id, resource:targetMappingDialog.resource, month:planningMonth || '', family:targetMappingFamily, createdBy:currentUser.email || '', createdAt:data?.created_at, updatedAt:data?.updated_at, cloud:true }
    setTargetMappings(current => [mapped, ...current.filter(item => !(normalize(item.resource).toUpperCase() === normalize(mapped.resource).toUpperCase() && item.month === mapped.month))])
    setTargetMappingDialog(null)
    setMappingMessage(`המיפוי של ${targetMappingDialog.resource} נשמר ואושר לכל המשתמשים`)
  }

  const removeTargetMapping = async mapping => {
    if (!mapping?.id || !supabase) return
    const { error } = await supabase.from('iml_target_mappings').update({ active:false, updated_at:new Date().toISOString() }).eq('id', mapping.id)
    if (error) { setMappingMessage(`ביטול המיפוי נכשל: ${error.message}`); return }
    setTargetMappings(current => current.filter(item => item.id !== mapping.id))
    setMappingMessage(`המיפוי של ${mapping.resource} בוטל`)
  }

  const activeTargetFamilies = useMemo(() => new Set(
    mappingTargets.flatMap(target => (target.facilities || [target.facility]).map(stationFamily)).filter(Boolean)
  ), [mappingTargets])

  const isDedicatedBulkMappingRow = row => {
    const facility = normalize(row?.facility)
    const description = normalize(row?.desc).toUpperCase()
    return (facility === '1142' && description.includes('999')) || (facility === '1119' && description.includes('777'))
  }

  const mappingSimulation = useMemo(() => {
    const assignments = new Map()
    planningRows.forEach(targetRow => {
      ;(targetRow.productionRows || []).forEach(row => {
        const list = assignments.get(row) || []
        list.push({
          resource: targetRow.resource,
          facility: targetRow.facility,
          station: targetRow.station,
          routingGroup: targetRow.routingGroup,
        })
        assignments.set(row, list)
      })
    })

    const rows = filtered.map((row, index) => {
      const matches = assignments.get(row) || []
      const rowKey = productionMappingKey(row)
      const approvedManual = manualMappings.find(mapping => mapping.active !== false && mapping.status === 'approved' && mapping.rowKey === rowKey)
      const pendingManual = manualMappings.find(mapping => mapping.active !== false && mapping.status === 'pending' && mapping.rowKey === rowKey)
      const stationDigits = String(row.facility || '').replace(/\D/g, '')
      const suffix = stationDigits.length >= 2 ? stationDigits.slice(-2) : ''
      const family = suffix ? `15${suffix}` : row.facility || ''
      const dedicatedBulk = isDedicatedBulkMappingRow(row)
      const relevantToActiveTarget = activeTargetFamilies.has(stationFamily(family))
      const rawStatus = matches.length === 0 ? 'unmatched' : matches.length > 1 ? 'duplicate' : 'matched'
      const status = dedicatedBulk ? 'ignored-bulk' : rawStatus
      const actionable = !dedicatedBulk && relevantToActiveTarget && rawStatus !== 'matched'
      return {
        key: `${row.order || ''}-${row.batch || ''}-${index}`,
        row,
        matches,
        status, rawStatus, actionable, relevantToActiveTarget, dedicatedBulk,
        family,
        assignedResource: matches.map(item => item.resource).filter(Boolean).join(' | '),
        approvedManual, pendingManual,
        explanation: dedicatedBulk ? (normalize(row.facility) === '1142' ? 'באלק מתקן 42 — מטופל רק במאזן 42 (1142 + 999)' : 'באלק מתקן 19 — מטופל רק במאזן 19 (1119 + 777)') : rawStatus === 'matched'
          ? `שויך ליעד אחד: ${matches[0].resource || matches[0].facility || 'ללא שם'}`
          : rawStatus === 'duplicate'
            ? `שויך ל-${matches.length} יעדים — חשש לספירה כפולה`
            : 'לא נמצא יעד תואם לפי כללי המנוע הנוכחיים',
      }
    })

    const summary = rows.reduce((acc, item) => {
      acc.rows += 1
      acc.quantity += Number(item.row.qty) || 0
      if (item.status === 'matched') { acc.matched += 1; acc.matchedQty += Number(item.row.qty) || 0 }
      if (item.status === 'unmatched') { acc.unmatched += 1; acc.unmatchedQty += Number(item.row.qty) || 0 }
      if (item.status === 'duplicate') { acc.duplicate += 1; acc.duplicateQty += Number(item.row.qty) || 0 }
      if (item.status === 'ignored-bulk') { acc.ignoredBulk += 1; acc.ignoredBulkQty += Number(item.row.qty) || 0 }
      if (item.actionable) { acc.actionable += 1; acc.actionableQty += Number(item.row.qty) || 0 }
      return acc
    }, { rows:0, quantity:0, matched:0, matchedQty:0, unmatched:0, unmatchedQty:0, duplicate:0, duplicateQty:0, ignoredBulk:0, ignoredBulkQty:0, actionable:0, actionableQty:0 })

    return { rows, summary }
  }, [filtered, planningRows, manualMappings, activeTargetFamilies])

  const visibleMappingSimulation = useMemo(() => (
    simulatorOnlyIssues
      ? mappingSimulation.rows.filter(item => item.actionable)
      : mappingSimulation.rows.filter(item => item.status !== 'ignored-bulk')
  ), [mappingSimulation, simulatorOnlyIssues])

  const exportMappingSimulation = () => {
    const rows = mappingSimulation.rows.map(item => ({
      Date: item.row.productionDay || iso(item.row.date),
      Time: item.row.date ? new Date(item.row.date).toLocaleTimeString('he-IL', {hour:'2-digit', minute:'2-digit'}) : '',
      StorageLocation: item.row.facility,
      FacilityFamily: item.family,
      ProdLine: item.row.prodLine || '',
      Tool: item.row.prodLineTool || '',
      RoutingGroup: item.row.routingGroup,
      OrderType: item.row.orderType,
      Order: item.row.order,
      Batch: item.row.batch,
      Material: item.row.material,
      Description: item.row.desc,
      Quantity: item.row.qty,
      MatchStatus: item.status,
      RelevantToActiveTarget: item.relevantToActiveTarget ? 'YES' : 'NO',
      ActionRequired: item.actionable ? 'YES' : 'NO',
      AssignedResource: item.assignedResource,
      Explanation: item.explanation,
    }))
    const wb = XLSX.utils.book_new()
    appendAutoFitJsonSheet(wb, rows, 'Mapping Simulator')
    XLSX.writeFile(wb, `IML_Production_Mapping_Simulator_${new Date().toISOString().slice(0,10)}.xlsx`)
  }

  const facilityStats = useMemo(() => facilities.map(id => {
    let rows = baseFiltered.filter(r => r.facility === id)
    if (id === '1542') rows = rows.filter(r => r.orderType.toUpperCase().includes('ZFIN'))
    const actual = rows.reduce((s, r) => s + r.qty, 0)
    const plans = planningRows.filter(x => (x.facilities || [x.facility]).includes(id))
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
  const filteredQualityRows = useMemo(() => dashboardQualityRows.filter(r =>
    (!selectedFacilities.length || selectedFacilities.includes(r.facility)) &&
    matchesDateRange(r.date, from, to)
  ), [dashboardQualityRows, selectedFacilities, from, to])

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

  const alerts = useMemo(() => {
    if (!selectedFacilities.length) return []
    return planningRows.filter(r => ['risk', 'warning'].includes(r.state) && (r.facilities || [r.facility]).some(id => selectedFacilities.includes(String(id)))).sort((a,b) => ({ risk:0, warning:1 }[a.state] - { risk:0, warning:1 }[b.state]))
  }, [planningRows, selectedFacilities])
  const achievedCount = planningRows.filter(r => ['achieved', 'good'].includes(r.state)).length
  const riskCount = planningRows.filter(r => r.state === 'risk').length
  const warningCount = planningRows.filter(r => r.state === 'warning').length
  const scopedPlanningRows = useMemo(() => {
    if (!selectedFacilities.length) return planningRows
    return planningRows.filter(row => {
      const rowFacilities = (row.facilities || [row.facility]).map(String)
      return rowFacilities.some(id => selectedFacilities.includes(id))
    })
  }, [planningRows, selectedFacilities])
  const targetTotal = scopedPlanningRows.reduce((s,r) => s + r.target, 0)
  const targetActual = scopedPlanningRows.reduce((s,r) => s + r.actual, 0)
  const targetForecast = scopedPlanningRows.reduce((s,r) => s + r.forecast, 0)
  const targetRequiredDaily = scopedPlanningRows.reduce((sum,row) => sum + row.requiredDaily, 0)
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
    openDeviations.forEach(row => { const facility=canonicalFacility(row.facility); if(facility) deviationByFacility.set(facility,(deviationByFacility.get(facility)||0)+1) })
    return planningRows.map(row => {
      const facilitiesForRow=row.facilities||[row.facility].filter(Boolean)
      const deviationsCount=facilitiesForRow.reduce((sum,id)=>sum+(deviationByFacility.get(id)||0),0)
      const forecastPct=row.target?row.forecast/row.target*100:0, actualPct=row.target?row.actual/row.target*100:0
      const planScore=row.target?Math.min(100,forecastPct):70, qualityScore=Math.max(0,100-deviationsCount*8)
      const healthScore=Math.max(0,Math.min(100,Math.round(planScore*.72+qualityScore*.28)))
      const state=!row.target?'no-target':forecastPct>=100?'good':forecastPct>=90?'warning':'risk'
      return { ...row, facility:row.resource||row.facility, facilityIds:facilitiesForRow, deviationsCount, forecastPct, actualPct, healthScore, state, gap:row.forecast-row.target }
    }).filter(row=>row.target>0||row.actual>0||row.deviationsCount>0)
      .sort((a,b)=>({risk:0,warning:1,good:2,'no-target':3}[a.state]-{risk:0,warning:1,good:2,'no-target':3}[b.state])||a.facility.localeCompare(b.facility))
  }, [planningRows, openDeviations])

  const facilityViewMatches = row => {
    if (facilityViewMode === 'target') return Number(row.target || 0) > 0
    if (facilityViewMode === 'active') return Number(row.actual || 0) > 0
    if (facilityViewMode === 'risk') return row.state === 'risk' || row.state === 'warning'
    if (facilityViewMode === 'all') return true
    return Number(row.target || 0) > 0 || Number(row.actual || 0) > 0
  }
  const visibleControlTowerFacilities = useMemo(() => controlTowerFacilities.filter(facilityViewMatches), [controlTowerFacilities, facilityViewMode])
  const visiblePlanningRows = useMemo(() => planningRows.filter(facilityViewMatches), [planningRows, facilityViewMode])
  const dailyPlanningRows = useMemo(() => planningRows.filter(row => (row.facilities || [row.facility]).some(id => dailyFacilities.includes(String(id)))), [planningRows, dailyFacilities])
  const addDailyFacility = () => {
    if (!dailyFacilityToAdd || dailyAdditionalFacilities.includes(dailyFacilityToAdd)) return
    setDailyAdditionalFacilities(current => [...current, dailyFacilityToAdd])
    setDailyFacilityToAdd('')
  }
  const removeDailyFacility = id => setDailyAdditionalFacilities(current => current.filter(item => item !== id))
  const selectedControlTowerFacilities = useMemo(() => selectedFacilities.length ? visibleControlTowerFacilities.filter(row => (row.facilityIds || [row.facility]).some(id => selectedFacilities.includes(String(id)))) : [], [visibleControlTowerFacilities, selectedFacilities])
  const selectedForecastRows = useMemo(() => selectedFacilities.length ? visiblePlanningRows.filter(row => (row.facilities || [row.facility]).some(id => selectedFacilities.includes(String(id)))) : [], [visiblePlanningRows, selectedFacilities])
  const selectedFacilityStats = useMemo(() => selectedFacilities.length ? facilityStats.filter(row => selectedFacilities.includes(String(row.id))) : [], [facilityStats, selectedFacilities])
  const reportDateToIso = value => {
    const text = String(value || '').trim()
    const match = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/)
    if (!match) return /^\d{4}-\d{2}-\d{2}$/.test(text) ? text : ''
    let year = Number(match[3]); if (year < 100) year += 2000
    return `${year}-${String(Number(match[2])).padStart(2,'0')}-${String(Number(match[1])).padStart(2,'0')}`
  }
  const selectedSingleReportDate = from && to && from === to ? from : ''
  const resetDailyEventForm = () => {
    setDailyEventType('')
    setDailyEventFacility(selectedFacilities.length === 1 ? selectedFacilities[0] : '')
    setDailyEventSeverity('')
    setDailyEventText('')
    setDailyEventDate(selectedSingleReportDate || isoDate(new Date()))
  }
  const openDailyEventForm = () => {
    resetDailyEventForm()
    setDailyEventFormOpen(true)
  }
  const saveDailyEvent = async () => {
    const eventDate = dailyEventDate || selectedSingleReportDate || isoDate(new Date())
    if (!eventDate || !dailyEventType || !dailyEventFacility || !dailyEventSeverity || !dailyEventText.trim()) {
      setStatus('יש למלא תאריך, סוג אירוע, מתקן, חומרה ותיאור לפני השמירה.')
      return
    }
    if (!supabase || !currentUser?.id) {
      setStatus('כדי לשמור אירוע משותף לכל המחשבים יש להיכנס כמנהל או מנהל מתקן.')
      onRequestAdminLogin?.()
      return
    }
    const entry = {
      id: `event-${Date.now()}-${Math.random().toString(36).slice(2,8)}`,
      date: eventDate,
      type: dailyEventType,
      facility: String(dailyEventFacility),
      severity: dailyEventSeverity,
      description: dailyEventText.trim(),
      createdAt: new Date().toISOString(),
      createdBy: currentUser?.email || '',
    }
    try {
      setBusy(true)
      const { data, error } = await supabase.from('iml_daily_events').insert({
        external_id:entry.id, event_date:entry.date, event_type:entry.type, facility:entry.facility,
        severity:entry.severity, description:entry.description, created_at:entry.createdAt,
        created_by:entry.createdBy, created_by_id:currentUser.id,
      }).select().single()
      if (error) throw error
      const saved = { ...entry, id:data?.external_id || entry.id, createdAt:data?.created_at || entry.createdAt }
      setDailyEvents(current => [saved, ...current.filter(item => item.id !== saved.id)])
      setDailyEventFormOpen(false)
      resetDailyEventForm()
      setDailyCloudReady(true)
      setStatus(`האירוע נשמר בענן בהצלחה לתאריך ${new Date(`${eventDate}T12:00:00`).toLocaleDateString('he-IL')} וזמין מכל מחשב.`)
    } catch (error) {
      console.error(error)
      setStatus(`שמירת האירוע בענן נכשלה: ${error?.message || 'שגיאה לא ידועה'}`)
    } finally {
      setBusy(false)
    }
  }
  const visibleDailyEvents = useMemo(() => dailyEvents.filter(event => {
    if (from && event.date < from) return false
    if (to && event.date > to) return false
    if (selectedFacilities.length && !selectedFacilities.includes(String(event.facility || ''))) return false
    return true
  }).sort((a,b) => String(b.date).localeCompare(String(a.date)) || String(b.createdAt).localeCompare(String(a.createdAt))), [dailyEvents, from, to, selectedFacilities])
  const savedDailyEventsForSelection = useMemo(() => dailyEvents.filter(event => {
    if (!selectedSingleReportDate || event.date !== selectedSingleReportDate) return false
    return !selectedFacilities.length || selectedFacilities.includes(String(event.facility || ''))
  }), [dailyEvents, selectedSingleReportDate, selectedFacilities])
  const savedDailyReportRowsForSelection = useMemo(() => dailyReportHistory.filter(row => {
    if (!selectedSingleReportDate || reportDateToIso(row.reportDate) !== selectedSingleReportDate) return false
    return !selectedFacilities.length || selectedFacilities.includes(String(row.facility || ''))
  }), [dailyReportHistory, selectedSingleReportDate, selectedFacilities])
  const monthlyDailyReportHistory = useMemo(() => {
    const parseMonthKey = value => {
      const text = String(value || '').trim()
      const m = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/)
      if (!m) return ''
      let year = Number(m[3]); if (year < 100) year += 2000
      return `${year}-${String(Number(m[2])).padStart(2,'0')}`
    }
    const map = new Map()
    dailyReportHistory.forEach(row => {
      const month = parseMonthKey(row.reportDate) || String(row.importedAt || '').slice(0,7)
      const facility = String(row.facility || 'ללא מתקן')
      const key = `${month}|${facility}`
      const current = map.get(key) || { month, facility, files:new Set(), rows:0, quantity:0, notes:0, machineStatuses:0 }
      current.files.add(row.fileName || row.reportDate || key)
      current.rows += 1
      current.quantity += num(row.quantity)
      if (String(row.notes || '').trim()) current.notes += 1
      if (String(row.machineStatus || '').trim()) current.machineStatuses += 1
      map.set(key, current)
    })
    return [...map.values()].map(item => ({ ...item, reports:item.files.size }))
      .sort((a,b) => b.month.localeCompare(a.month) || String(b.facility).localeCompare(String(a.facility),'he',{numeric:true}))
  }, [dailyReportHistory])
  const facilityViewFilters = <div className="facility-view-filters">
    {[['relevant','יעד + פעילות'],['target','עם יעד'],['active','פעילים'],['risk','בסיכון'],['all','הכול']].map(([value,label]) =>
      <button key={value} type="button" className={facilityViewMode === value ? 'active' : ''} onClick={() => setFacilityViewMode(value)}>{label}</button>
    )}
  </div>

  const controlTowerTrend = useMemo(() => {
    const byDay = new Map()
    dashboardProd.filter(row => monthKey(row.date) === planningMonth).forEach(row => {
      const key = iso(row.date)
      if (key) byDay.set(key, (byDay.get(key) || 0) + row.qty)
    })
    return [...byDay.entries()].sort((a,b) => a[0].localeCompare(b[0])).slice(-7)
  }, [dashboardProd, planningMonth])

  const jumpToDetails = (tab) => {
    setActiveTab(tab)
    window.setTimeout(() => document.getElementById('details-section')?.scrollIntoView({ behavior:'smooth', block:'start' }), 50)
  }

  const facilityColor = facilityColorFor
  const facilitySortDesc = (a,b) => {
    const an = Number(String(a.facility || '').replace(/\D/g,'')), bn = Number(String(b.facility || '').replace(/\D/g,''))
    if (Number.isFinite(an) && Number.isFinite(bn) && an !== bn) return bn-an
    return String(b.facility || '').localeCompare(String(a.facility || ''), 'he')
  }
  const xmlEscape = value => String(value ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&apos;')
  const buildDailyFormulationsWorksheet = productionRows => {
    const rows = [...productionRows].sort((a,b) => facilitySortDesc(a,b) || String(a.routingGroup||'').localeCompare(String(b.routingGroup||''), 'he', {numeric:true}) || String(a.material||'').localeCompare(String(b.material||''), 'he', {numeric:true}))
    const grouped = new Map()
    rows.forEach(row => {
      const facility = String(row.facility || '—')
      const list = grouped.get(facility) || []
      list.push(row)
      grouped.set(facility, list)
    })
    const displayDate = from && to && from === to
      ? new Date(`${from}T12:00:00`).toLocaleDateString('he-IL', {day:'2-digit',month:'2-digit',year:'2-digit'})
      : (from || to || new Date().toLocaleDateString('he-IL', {day:'2-digit',month:'2-digit',year:'2-digit'}))
    const eventText = savedDailyEventsForSelection.length
      ? savedDailyEventsForSelection.map(event => `אירוע ${event.type} · מתקן ${event.facility} · חומרה ${event.severity} - ${event.description}`).join(' | ')
      : 'לא נשמרו אירועים לתאריך הדוח שנבחר.'
    const cell = (value='', style='formBody', extra='') => {
      const numeric = typeof value === 'number' && Number.isFinite(value)
      return `<Cell ss:StyleID="${style}"${extra}><Data ss:Type="${numeric ? 'Number' : 'String'}">${xmlEscape(value)}</Data></Cell>`
    }
    const outputRows = []
    let excelRow = 8
    grouped.forEach((groupRows, facility) => {
      const total = groupRows.reduce((sum,row) => sum + num(row.qty), 0)
      const span = Math.max(0, groupRows.length - 1)
      groupRows.forEach((row,index) => {
        const cells = []
        cells.push(cell(row.material || '', 'formPlain', ' ss:Index="2"'))
        if (index === 0) {
          cells.push(cell(facility, 'formGroup', span ? ` ss:MergeDown="${span}"` : ''))
          cells.push(cell(row.prodLineTool || row.prodLine || row.routingGroup || '', 'formBody'))
          cells.push(cell(row.desc || '', 'formBody'))
          cells.push(cell(row.batch || '', 'formBody'))
          cells.push(cell(row.machineStatus || '', 'formStatus'))
          cells.push(cell(num(row.qty), 'formBody'))
          cells.push(cell(total, 'formTotal', span ? ` ss:MergeDown="${span}"` : ''))
          cells.push(cell(row.notes || '', 'formNotes'))
          cells.push(cell('', 'formExtra'))
        } else {
          cells.push(cell(row.prodLineTool || row.prodLine || row.routingGroup || '', 'formBody', ' ss:Index="4"'))
          cells.push(cell(row.desc || '', 'formBody'))
          cells.push(cell(row.batch || '', 'formBody'))
          cells.push(cell(row.machineStatus || '', 'formStatus'))
          cells.push(cell(num(row.qty), 'formBody'))
          cells.push(cell(row.notes || '', 'formNotes', ' ss:Index="10"'))
          cells.push(cell('', 'formExtra'))
        }
        outputRows.push(`<Row ss:AutoFitHeight="1">${cells.join('')}</Row>`)
        excelRow += 1
      })
    })
    if (!rows.length) outputRows.push(`<Row>${cell('', 'formPlain', ' ss:Index="2"')}${cell('אין נתוני תפוקה בטווח שנבחר','formBody',' ss:MergeAcross="7"')}</Row>`)
    return `<Worksheet ss:Name="דיווח יומי פורמולציות" ss:RightToLeft="1"><Table x:FullColumns="1" x:FullRows="1"><Column ss:Index="2" ss:Width="92"/><Column ss:Width="135"/><Column ss:Width="75"/><Column ss:Width="250"/><Column ss:Width="105"/><Column ss:Width="155"/><Column ss:Width="88"/><Column ss:Width="95"/><Column ss:Width="110"/><Column ss:Width="275"/><Row ss:Index="5" ss:Height="23"><Cell ss:Index="3" ss:MergeAcross="7" ss:StyleID="formEvent"><Data ss:Type="String">${xmlEscape(eventText)}</Data></Cell></Row><Row ss:Index="7" ss:Height="22">${cell('מקט','formHeader',' ss:Index="2"')}${cell(displayDate,'formDateHeader')}${cell('קו יצור','formHeader')}${cell('חומר','formHeader')}${cell('מספר אצווה','formHeader')}${cell('סטטוס מכונה','formHeader')}${cell('תפוקה','formHeader')}${cell('סה"כ תפוקה','formHeader')}${cell('הערות','formHeader')}</Row>${outputRows.join('')}</Table><WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel"><DisplayRightToLeft/><Selected/><FreezePanes/><FrozenNoSplit/><SplitHorizontal>7</SplitHorizontal><TopRowBottomPane>7</TopRowBottomPane><ProtectObjects>False</ProtectObjects><ProtectScenarios>False</ProtectScenarios></WorksheetOptions></Worksheet>`
  }

  const exportStyledExcel = (sheets, filename, productionRowsForTemplate = []) => {
    // Self-contained Office Open XML writer. This keeps the report as a real
    // .xlsx file while preserving fills, borders, RTL and alignment without
    // adding any npm/CDN dependency (important for stable Netlify builds).
    const esc = value => String(value ?? '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;')
    const colName = n => { let s=''; for (let x=n+1; x>0; x=Math.floor((x-1)/26)) s=String.fromCharCode(65+((x-1)%26))+s; return s }
    const excelTextLength = value => Math.max(...String(value ?? '').split(/\r?\n/).map(line => [...line].length), 0)
    const autoWidth = (values, { min=9, max=55, pad=3 } = {}) => Math.min(max, Math.max(min, ...values.map(v => excelTextLength(v) + pad)))
    const autoRowHeight = (values, widths, { min=20, max=72 } = {}) => {
      let lines = 1
      values.forEach((value,index) => {
        const width = Math.max(6, widths[index] || 12)
        const wrapped = String(value ?? '').split(/\r?\n/).reduce((sum,line) => sum + Math.max(1, Math.ceil([...line].length / Math.max(5, width - 2))), 0)
        lines = Math.max(lines, wrapped)
      })
      return Math.min(max, Math.max(min, 18 * lines))
    }
    const cellXml = (r,c,value,style=5) => {
      if (value === null || value === undefined || value === '') return ''
      const ref = `${colName(c)}${r+1}`
      if (typeof value === 'number' && Number.isFinite(value)) return `<c r="${ref}" s="${style}" t="n"><v>${value}</v></c>`
      return `<c r="${ref}" s="${style}" t="inlineStr"><is><t xml:space="preserve">${esc(value)}</t></is></c>`
    }
    const crcTable = (() => {
      const t = new Uint32Array(256)
      for (let n=0;n<256;n++) { let c=n; for (let k=0;k<8;k++) c=(c&1)?(0xEDB88320^(c>>>1)):(c>>>1); t[n]=c>>>0 }
      return t
    })()
    const crc32 = bytes => { let c=0xFFFFFFFF; for (const b of bytes) c=crcTable[(c^b)&255]^(c>>>8); return (c^0xFFFFFFFF)>>>0 }
    const u16 = n => [n&255,(n>>>8)&255]
    const u32 = n => [n&255,(n>>>8)&255,(n>>>16)&255,(n>>>24)&255]
    const encoder = new TextEncoder()
    const zipStore = entries => {
      const locals=[], centrals=[]; let offset=0
      entries.forEach(({name,text}) => {
        const nb=encoder.encode(name), data=encoder.encode(text), crc=crc32(data)
        const local = new Uint8Array([0x50,0x4b,0x03,0x04,...u16(20),...u16(0),...u16(0),...u16(0),...u16(0),...u32(crc),...u32(data.length),...u32(data.length),...u16(nb.length),...u16(0),...nb,...data])
        locals.push(local)
        const central = new Uint8Array([0x50,0x4b,0x01,0x02,...u16(20),...u16(20),...u16(0),...u16(0),...u16(0),...u16(0),...u32(crc),...u32(data.length),...u32(data.length),...u16(nb.length),...u16(0),...u16(0),...u16(0),...u16(0),...u32(0),...u32(offset),...nb])
        centrals.push(central); offset += local.length
      })
      const centralSize=centrals.reduce((n,a)=>n+a.length,0)
      const end = new Uint8Array([0x50,0x4b,0x05,0x06,...u16(0),...u16(0),...u16(entries.length),...u16(entries.length),...u32(centralSize),...u32(offset),...u16(0)])
      return new Blob([...locals,...centrals,end], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})
    }
    const safeSheetName = (name, used) => {
      const base=String(name||'Sheet').replace(/[\\/?*\[\]:]/g,' ').trim().slice(0,31)||'Sheet'
      let candidate=base,n=2; while(used.has(candidate)) candidate=`${base.slice(0,Math.max(1,28-String(n).length))} ${n++}`
      used.add(candidate); return candidate
    }
    const usedNames=new Set()
    const xmlSheets=[]

    const rows=[...productionRowsForTemplate].sort((a,b)=>facilitySortDesc(a,b)||String(a.routingGroup||'').localeCompare(String(b.routingGroup||''),'he',{numeric:true})||String(a.material||'').localeCompare(String(b.material||''),'he',{numeric:true}))
    const grouped=new Map()
    rows.forEach(row=>{const f=String(row.facility||'—'); const arr=grouped.get(f)||[]; arr.push(row); grouped.set(f,arr)})
    const displayDate=from&&to&&from===to?new Date(`${from}T12:00:00`).toLocaleDateString('he-IL',{day:'2-digit',month:'2-digit',year:'2-digit'}):(from||to||new Date().toLocaleDateString('he-IL',{day:'2-digit',month:'2-digit',year:'2-digit'}))
    const eventText=savedDailyEventsForSelection.length?savedDailyEventsForSelection.map(event=>`אירוע ${event.type} · מתקן ${event.facility} · חומרה ${event.severity} - ${event.description}`).join(' | '):'לא נשמרו אירועים לתאריך הדוח שנבחר.'
    const dailyHeaders=['מקט',displayDate,'קו יצור','חומר','מספר אצווה','סטטוס מכונה','תפוקה','סה"כ תפוקה','הערות']
    const dailyColumnValues = dailyHeaders.map((header,index) => [header, ...rows.map(row => {
      const facility=String(row.facility||'—')
      const total=(grouped.get(facility)||[]).reduce((sum,item)=>sum+num(item.qty),0)
      return [row.material||'',facility,row.prodLineTool||row.prodLine||row.routingGroup||'',row.desc||'',row.batch||'','',num(row.qty),total,''][index]
    })])
    const dailyWidths = dailyColumnValues.map((values,index) => autoWidth(values, index===3 ? {min:18,max:48,pad:3} : index===8 ? {min:12,max:34,pad:3} : {min:10,max:24,pad:3}))
    const dailyRows=[]
    dailyRows.push(`<row r="2" ht="28" customHeight="1">${cellXml(1,1,'דוח יומי – מתקנים נבחרים',1)}</row>`)
    dailyRows.push(`<row r="3" ht="22" customHeight="1">${cellXml(2,1,`מספר מתקנים בדוח: ${grouped.size}`,2)}</row>`)
    dailyRows.push(`<row r="5" ht="30" customHeight="1">${cellXml(4,1,eventText,3)}</row>`)
    dailyRows.push(`<row r="7" ht="${autoRowHeight(dailyHeaders,dailyWidths,{min:24,max:42})}" customHeight="1">${dailyHeaders.map((h,i)=>cellXml(6,i+1,h,4)).join('')}</row>`)
    const merges=['B2:J2','B3:J3','B5:J5']
    let rr=7
    grouped.forEach((groupRows,facility)=>{
      const total=groupRows.reduce((sum,row)=>sum+num(row.qty),0), start=rr
      groupRows.forEach((row,index)=>{
        const cells=[
          cellXml(rr,1,row.material||'',5),
          index===0?cellXml(rr,2,facility,6):'',
          cellXml(rr,3,row.prodLineTool||row.prodLine||row.routingGroup||'',5),
          cellXml(rr,4,row.desc||'',5),
          cellXml(rr,5,row.batch||'',5),
          '', // סטטוס מכונה — intentionally blank for manual edit
          cellXml(rr,7,num(row.qty),7),
          index===0?cellXml(rr,8,total,7):'',
          ''  // הערות — intentionally blank for manual edit
        ]
        const rowValues=[row.material||'',facility,row.prodLineTool||row.prodLine||row.routingGroup||'',row.desc||'',row.batch||'','',num(row.qty),total,'']
        const rowHeight=autoRowHeight(rowValues,dailyWidths,{min:20,max:72})
        dailyRows.push(`<row r="${rr+1}" ht="${rowHeight}" customHeight="1">${cells.join('')}</row>`); rr++
      })
      if(groupRows.length>1){merges.push(`C${start+1}:C${rr}`); merges.push(`I${start+1}:I${rr}`)}
    })
    if(!rows.length) dailyRows.push(`<row r="8" ht="22" customHeight="1">${cellXml(7,1,'אין נתוני תפוקה בטווח שנבחר',5)}</row>`)
    const dailySheet=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0" rightToLeft="1"><pane ySplit="7" topLeftCell="A8" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>${dailyWidths.map((width,index)=>`<col min="${index+2}" max="${index+2}" width="${width}" customWidth="1" bestFit="1"/>`).join('')}</cols><sheetData>${dailyRows.join('')}</sheetData><mergeCells count="${merges.length}">${merges.map(ref=>`<mergeCell ref="${ref}"/>`).join('')}</mergeCells><pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.2" footer="0.2"/></worksheet>`
    xmlSheets.push({name:safeSheetName('דיווח יומי פורמולציות',usedNames),xml:dailySheet})

    sheets.forEach(sh=>{
      const cols=sh.columns||[], body=sh.rows||[]
      const sheetRows=[]
      sheetRows.push(`<row r="1" ht="24" customHeight="1">${cols.map((c,i)=>cellXml(0,i,c.label,4)).join('')}</row>`)
      const genericWidths=cols.map(c=>autoWidth([c.label,...body.map(r=>r[c.key])],{min:10,max:55,pad:3}))
      body.forEach((r,ri)=>{ const vals=cols.map(c=>r[c.key]??''); const h=autoRowHeight(vals,genericWidths,{min:20,max:72}); sheetRows.push(`<row r="${ri+2}" ht="${h}" customHeight="1">${cols.map((c,ci)=>cellXml(ri+1,ci,r[c.key]??'',typeof r[c.key]==='number'?7:5)).join('')}</row>`) })
      const widths=genericWidths.map((w,i)=>`<col min="${i+1}" max="${i+1}" width="${w}" customWidth="1" bestFit="1"/>`).join('')
      const ref=cols.length?`A1:${colName(cols.length-1)}${Math.max(1,body.length+1)}`:'A1:A1'
      const xml=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0" rightToLeft="1"/></sheetViews><cols>${widths}</cols><sheetData>${sheetRows.join('')}</sheetData>${cols.length?`<autoFilter ref="${ref}"/>`:''}<pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.2" footer="0.2"/></worksheet>`
      xmlSheets.push({name:safeSheetName(sh.name,usedNames),xml})
    })

    const styles=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="5"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="15"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Segoe UI"/></font><font><b/><sz val="10"/><color rgb="FF0B2F4A"/><name val="Segoe UI"/></font><font><sz val="10"/><color rgb="FF1F2937"/><name val="Segoe UI"/></font></fonts><fills count="6"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF0B2F4A"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDCEAF3"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFDE7A8"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFEAF3F8"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFB8C6D1"/></left><right style="thin"><color rgb="FFB8C6D1"/></right><top style="thin"><color rgb="FFB8C6D1"/></top><bottom style="thin"><color rgb="FFB8C6D1"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="8"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" readingOrder="2"/></xf><xf numFmtId="0" fontId="3" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" readingOrder="2"/></xf><xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1" readingOrder="2"/></xf><xf numFmtId="0" fontId="2" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1" readingOrder="2"/></xf><xf numFmtId="0" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1" readingOrder="2"/></xf><xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1" readingOrder="2"/></xf><xf numFmtId="3" fontId="4" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1" applyNumberFormat="1"><alignment horizontal="center" vertical="center" readingOrder="2"/></xf></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`
    const workbookSheets=xmlSheets.map((sh,i)=>`<sheet name="${esc(sh.name)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join('')
    const workbook=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets>${workbookSheets}</sheets></workbook>`
    const wbRels=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${xmlSheets.map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join('')}<Relationship Id="rId${xmlSheets.length+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`
    const types=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${xmlSheets.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}</Types>`
    const rootRels=`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`
    const entries=[{name:'[Content_Types].xml',text:types},{name:'_rels/.rels',text:rootRels},{name:'xl/workbook.xml',text:workbook},{name:'xl/_rels/workbook.xml.rels',text:wbRels},{name:'xl/styles.xml',text:styles},...xmlSheets.map((sh,i)=>({name:`xl/worksheets/sheet${i+1}.xml`,text:sh.xml}))]
    const blob=zipStore(entries)
    const outName=String(filename||'IML_Facility_Report.xlsx').replace(/\.(xml|xls|xlsx)$/i,'.xlsx')
    const url=URL.createObjectURL(blob), a=document.createElement('a'); a.href=url; a.download=outName; document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(url),1000)
  }


  const handleDailyReportRoundtripFile = async file => {
    try {
      setBusy(true)
      const buf = await file.arrayBuffer()
      const wb = XLSX.read(buf, { type:'array', cellDates:true })
      const sheetName = wb.SheetNames.find(name => /דיווח יומי פורמולציות/i.test(name)) || wb.SheetNames[0]
      const ws = wb.Sheets[sheetName]
      const matrix = XLSX.utils.sheet_to_json(ws, { header:1, defval:'', raw:false })
      const headerRowIndex = matrix.findIndex(row => row.some(cell => String(cell).trim() === 'מקט') && row.some(cell => /תפוקה/.test(String(cell))))
      if (headerRowIndex < 0) throw new Error('לא נמצאה שורת כותרות של דיווח יומי פורמולציות')

      const header = matrix[headerRowIndex].map(v => String(v || '').trim())
      const idx = label => header.findIndex(h => h === label)
      const materialIdx = idx('מקט')
      const lineIdx = idx('קו יצור')
      const descIdx = idx('חומר')
      const batchIdx = header.findIndex(h => /אצווה|Batch/i.test(h))
      const machineIdx = idx('סטטוס מכונה')
      const qtyIdx = idx('תפוקה')
      const totalIdx = idx('סה"כ תפוקה')
      const notesIdx = idx('הערות')
      const reportDate = header.find(h => /^\d{1,2}[./-]\d{1,2}[./-]\d{2,4}$/.test(h)) || ''
      const facilityIdx = header.indexOf(reportDate)

      let lastFacility = ''
      const imported = []
      matrix.slice(headerRowIndex + 1).forEach(row => {
        const material = String(row[materialIdx] || '').trim()
        if (!material) return
        const facilityCell = facilityIdx >= 0 ? String(row[facilityIdx] || '').trim() : ''
        if (facilityCell) lastFacility = facilityCell
        imported.push({
          importedAt:new Date().toISOString(),
          fileName:file.name,
          reportDate,
          material,
          facility:lastFacility,
          line:lineIdx >= 0 ? String(row[lineIdx] || '').trim() : '',
          description:descIdx >= 0 ? String(row[descIdx] || '').trim() : '',
          batch:batchIdx >= 0 ? String(row[batchIdx] || '').trim() : '',
          machineStatus:machineIdx >= 0 ? String(row[machineIdx] || '').trim() : '',
          quantity:qtyIdx >= 0 ? num(row[qtyIdx]) : 0,
          facilityTotal:totalIdx >= 0 ? num(row[totalIdx]) : 0,
          notes:notesIdx >= 0 ? String(row[notesIdx] || '').trim() : '',
        })
      })
      if (!imported.length) throw new Error('לא נמצאו רשומות להחזרה לאפליקציה')
      if (!supabase || !currentUser?.id) throw new Error('נדרשת כניסת מנהל כדי לשמור דוח ערוך בענן')
      const reportIso = reportDateToIso(reportDate)
      if (!reportIso) throw new Error('לא זוהה תאריך תקין בכותרת הדוח')
      const uploadToken = `${file.name}-${Date.now()}`
      const payload = imported.map((row, index) => ({
        client_key:`${reportIso}|${row.facility || ''}|${row.material || ''}|${row.batch || ''}|${row.line || ''}|${uploadToken}|${index}`,
        imported_at:row.importedAt, file_name:row.fileName, report_date:reportIso, material:row.material,
        facility:String(row.facility || ''), line:row.line, description:row.description, batch:row.batch,
        machine_status:row.machineStatus, quantity:num(row.quantity), facility_total:num(row.facilityTotal), notes:row.notes,
        created_by:currentUser?.email || '', created_by_id:currentUser.id,
      }))
      const { data:savedRows, error:saveError } = await supabase.from('iml_daily_report_rows').insert(payload).select('*')
      if (saveError) throw saveError
      const cloudRows = (savedRows || []).map(row => ({
        id:row.id, importedAt:row.imported_at, fileName:row.file_name, reportDate:row.report_date,
        material:row.material, facility:String(row.facility || ''), line:row.line || '', description:row.description || '',
        batch:row.batch || '', machineStatus:row.machine_status || '', quantity:num(row.quantity), facilityTotal:num(row.facility_total), notes:row.notes || '',
      }))
      setDailyReportHistory(current => [...cloudRows, ...current])
      setDailyCloudReady(true)
      setStatus(`הדוח הערוך נטען חזרה: ${fmt(cloudRows.length)} רשומות נשמרו ב־Supabase וזמינות מכל מחשב`)
    } catch (error) {
      console.error(error)
      setStatus(`טעינת הדוח הערוך נכשלה: ${error?.message || 'שגיאה לא ידועה'}`)
    } finally {
      setBusy(false)
    }
  }

  const downloadSavedDailyReport = () => {
    if (!selectedSingleReportDate) {
      setStatus('כדי להוריד דוח שמור יש לבחור יום יחיד וזהה בשדות מתאריך ועד תאריך')
      return
    }
    if (!savedDailyReportRowsForSelection.length) {
      setStatus(`לא נמצא דוח ערוך שמור לתאריך ${selectedSingleReportDate}${selectedFacilities.length ? ` ולמתקנים ${selectedFacilities.join(', ')}` : ''}`)
      return
    }
    const restoredRows = savedDailyReportRowsForSelection.map(row => ({
      date:new Date(`${selectedSingleReportDate}T12:00:00`), productionDay:selectedSingleReportDate,
      facility:String(row.facility || ''), material:String(row.material || ''),
      routingGroup:String(row.line || ''), desc:String(row.description || ''), batch:String(row.batch || ''),
      machineStatus:String(row.machineStatus || ''), qty:num(row.quantity), notes:String(row.notes || '')
    }))
    exportStyledExcel([
      { name:'דוח שמור', columns:[
        {key:'Date',label:'תאריך'},{key:'Facility',label:'מתקן'},{key:'Material',label:'מק״ט'},
        {key:'Description',label:'חומר'},{key:'Batch',label:'מספר אצווה'},
        {key:'MachineStatus',label:'סטטוס מכונה'},{key:'Quantity',label:'תפוקה'},{key:'Notes',label:'הערות'}
      ], rows:savedDailyReportRowsForSelection.map(row => ({
        Date:selectedSingleReportDate, Facility:row.facility, Material:row.material,
        Description:row.description, Batch:row.batch, MachineStatus:row.machineStatus,
        Quantity:num(row.quantity), Notes:row.notes
      })) }
    ], `IML_Daily_Report_Saved_${selectedSingleReportDate}.xlsx`, restoredRows)
    setStatus(`הדוח הערוך השמור לתאריך ${selectedSingleReportDate} הורד בהצלחה`)
  }

  const exportWorkbook = () => {
    const productionRows = [...filtered].sort(facilitySortDesc)
    const totals = productionRows.reduce((m,r) => { const f=String(r.facility||'—'); m.set(f,(m.get(f)||0)+num(r.qty)); return m }, new Map())
    const summaryRows = [...totals.entries()].sort((a,b)=>facilitySortDesc({facility:a[0]},{facility:b[0]})).map(([facility,total]) => ({ __facility:facility, Facility:facility, Records:productionRows.filter(r=>String(r.facility||'—')===facility).length, TotalQuantity:total }))
    const facilityToolMap = new Map()
    productionRows.filter(r => ['1523','1528'].includes(String(r.facility || ''))).forEach(r => {
      const facility = excelFacilityLabel(r)
      const tool = r.prodLineTool || r.prodLine || 'לא סווג'
      const key = `${facility}|${tool}|${r.prodLine || ''}`
      const current = facilityToolMap.get(key) || { __facility:String(r.facility || ''), Facility:facility, Tool:tool, ProdLine:r.prodLine || '', Records:0, TotalQuantity:0 }
      current.Records += 1
      current.TotalQuantity += num(r.qty)
      facilityToolMap.set(key, current)
    })
    const facilityToolRows = [...facilityToolMap.values()].sort((a,b) => String(a.Facility).localeCompare(String(b.Facility),'he',{numeric:true}) || String(a.Tool).localeCompare(String(b.Tool),'he',{numeric:true}))
    exportStyledExcel([
      { name:'Production', columns:[['Date','תאריך'],['Time','שעה'],['Facility','מתקן'],['ProdLine','PROD LINE'],['Tool','כלי / אזור'],['FacilityTotal','סה״כ מתקן'],['Order','הזמנה'],['Batch','Batch'],['Material','מק״ט חומר'],['Description','תיאור חומר'],['RoutingGroup','מתקן / תחנה'],['Quantity','כמות'],['GroupTotal','סה״כ']].map(([key,label])=>({key,label})), rows:productionRows.map((r,index,rows)=>{ const facility=String(r.facility||'—'); const firstInFacility=index===0 || String(rows[index-1]?.facility||'—')!==facility; return { __facility:facility, Date:iso(r.date), Time:r.date?r.date.toLocaleTimeString('he-IL',{hour:'2-digit',minute:'2-digit'}):'', Facility:r.facility, ProdLine:r.prodLine || '', Tool:r.prodLineTool || '', FacilityTotal:totals.get(facility)||0, Order:r.order, Batch:r.batch, Material:r.material, Description:r.desc, RoutingGroup:r.routingGroup, Quantity:r.qty, GroupTotal:firstInFacility ? (totals.get(facility)||0) : '' } }) },
      { name:'סיכום מתקנים', columns:[{key:'Facility',label:'מתקן'},{key:'Records',label:'מספר רשומות'},{key:'TotalQuantity',label:'סה״כ כמות'}], rows:summaryRows },
      { name:'פירוט 23-28', columns:[{key:'Facility',label:'מתקן'},{key:'Tool',label:'כלי / אזור'},{key:'ProdLine',label:'PROD LINE'},{key:'Records',label:'מספר רשומות'},{key:'TotalQuantity',label:'סה״כ כמות'}], rows:facilityToolRows },
      { name:'Planning', columns:[{key:'Month',label:'Month'},{key:'Facility',label:'Facility'},{key:'Station',label:'Station'},{key:'MonthlyTarget',label:'Monthly Target'},{key:'Actual',label:'Actual'},{key:'Remaining',label:'Remaining'},{key:'Forecast',label:'Forecast'},{key:'Status',label:'Status'}], rows:planningRows.map(r=>({__facility:String(r.facility||'—'),Month:planningMonth,Facility:r.facility,Station:r.station,MonthlyTarget:r.target,Actual:r.actual,Remaining:r.remaining,Forecast:r.forecast,Status:r.label})) },
      { name:'Quality', columns:[{key:'Date',label:'Date'},{key:'Facility',label:'Facility'},{key:'InspectionLot',label:'Inspection Lot'},{key:'Order',label:'Order'},{key:'Batch',label:'Batch'},{key:'Material',label:'Material'},{key:'Status',label:'Status'}], rows:qualityBad.map(r=>({__facility:String(r.facility||'—'),Date:iso(r.date),Facility:r.facility,InspectionLot:r.inspectionLot,Order:r.order,Batch:r.batch,Material:r.material,Status:r.status})) },
      { name:'חריגות איכות', columns:[{key:'Date',label:'Date'},{key:'Facility',label:'Facility'},{key:'Batch',label:'Batch'},{key:'Material',label:'Material'},{key:'Status',label:'Status'},{key:'Remarks',label:'Remarks'}], rows:openDeviations.map(r=>({__facility:String(r.facility||'—'),Date:iso(r.date),Facility:r.facility,Batch:r.batch,Material:r.material,Status:r.status,Remarks:r.remarks})) }
    ], `IML_Facility_Report_${from && to ? (from === to ? from : `${from}_to_${to}`) : (from || to || new Date().toISOString().slice(0,10))}.xlsx`, productionRows)
  }

  const downloadTargetWorkbook = async () => {
    try {
      // Cloud is the source of truth. Every computer should receive the exact
      // workbook that was last uploaded, including newly added target rows.
      let stored = null
      try {
        const requestedTargetMonth = planningMonth || monthKey(new Date())
        stored = await loadMonthlyTargetWorkbook(requestedTargetMonth)
        if (!stored?.bytes) stored = await loadActiveTargetWorkbook()
        if (stored?.bytes) await idbSetKey(TARGET_FILE_KEY, stored)
      } catch (cloudWorkbookError) {
        console.warn('Cloud monthly target workbook unavailable; using local fallback', cloudWorkbookError)
      }

      if (!stored?.bytes) stored = await idbGetKey(TARGET_FILE_KEY)
      if (stored?.bytes) {
        const blob = new Blob([stored.bytes], { type: stored.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
        const url = URL.createObjectURL(blob)
        const link = document.createElement('a')
        link.href = url
        link.download = stored.name || `IML_Monthly_Targets_${planningMonth || monthKey(new Date())}.xlsx`
        document.body.appendChild(link)
        link.click()
        link.remove()
        setTimeout(() => URL.revokeObjectURL(url), 1500)
        setStatus(stored.source === 'cloud'
          ? `הורד קובץ היעדים של ${planningMonth || monthKey(new Date())} מהענן — זהה בכל המחשבים`
          : 'הורד קובץ היעדים המקומי. הענן לא היה זמין ולכן נעשה שימוש בגיבוי המקומי.')
        return
      }

      // First installation fallback: master workbook bundled with the application.
      const response = await fetch('/templates/IML_Targets_Master.xlsx', { cache:'no-store' })
      if (!response.ok) throw new Error('קובץ תבנית היעדים המלא לא נמצא באפליקציה')
      const blob = await response.blob()
      const url = URL.createObjectURL(blob)
      const link = document.createElement('a')
      link.href = url
      link.download = '26 דוח ביצוע SAP Aug.xlsx'
      document.body.appendChild(link)
      link.click()
      link.remove()
      setTimeout(() => URL.revokeObjectURL(url), 1500)
      setStatus('לא נמצא עדיין קובץ יעדים פעיל בענן — הורדה תבנית ראשונית מהאפליקציה')
    } catch (error) {
      console.error('Target workbook download failed', error)
      setStatus(`הורדת קובץ היעדים נכשלה: ${error?.message || 'שגיאה לא ידועה'}`)
    }
  }

  const openPrintReport = (title, subtitle, headers, rows) => {
    const popup = window.open('', '_blank', 'width=1400,height=900')
    if (!popup) { setStatus('הדפדפן חסם חלון הדפסה. יש לאפשר חלונות קופצים לאתר.'); return }
    const esc = value => String(value ?? '').replace(/[&<>\"]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;'}[ch]))
    const facilityRows = rows.filter(row => row._facility)
    const facilityIds = [...new Set(facilityRows.map(row => String(row._facility)))]
    const showFacilityColors = facilityIds.length > 1
    const facilitySummary = facilityIds.map(facility => {
      const group = facilityRows.filter(row => String(row._facility) === facility)
      const hasQty = group.some(row => Number.isFinite(Number(row._qtyRaw)))
      const qty = group.reduce((sum, row) => sum + (Number.isFinite(Number(row._qtyRaw)) ? Number(row._qtyRaw) : 0), 0)
      return { facility, count: group.length, qty, hasQty, color: facilityColorFor(facility) }
    }).sort((a,b) => String(b.facility).localeCompare(String(a.facility), 'he', {numeric:true}))
    const facilitySummaryHtml = showFacilityColors ? `<div class="facility-summary">${facilitySummary.map(item => `<div class="facility-chip" style="background:${item.color};border-color:${item.color}"><b>מתקן ${esc(item.facility)}</b><span>${item.hasQty ? `סה״כ ${esc(fmt(item.qty))}` : `${item.count} רשומות`}</span></div>`).join('')}</div>` : ''
    const bodyRows = rows.map((row, index) => {
      const facility = row._facility ? String(row._facility) : ''
      const colorStyle = showFacilityColors && facility ? ` style="--facility-bg:${facilityColorFor(facility)}"` : ''
      const facilityClass = showFacilityColors && facility ? ' facility-row' : ''
      return `<tr class="${esc(row._state || '')}${facilityClass}"${colorStyle}><td class="index">${index+1}</td>${headers.map(h => `<td>${esc(row[h.key])}</td>`).join('')}</tr>`
    }).join('')
    popup.document.write(`<!doctype html><html dir="rtl"><head><meta charset="utf-8"><title>${esc(title)}</title><style>
      @page{size:A4 landscape;margin:9mm}*{box-sizing:border-box}body{font-family:Arial,"Segoe UI",sans-serif;color:#16324a;margin:0;background:#f4f8fb}.sheet{background:white;padding:22px}.head{display:flex;justify-content:space-between;align-items:flex-end;border-bottom:4px solid #159b83;padding-bottom:14px;margin-bottom:18px}.head h1{margin:0;font-size:28px}.head p{margin:5px 0 0;color:#64748b}.brand{font-weight:800;color:#159b83;font-size:19px}.summary{display:flex;gap:10px;margin:0 0 12px;flex-wrap:wrap}.chip{background:#eaf7f3;border:1px solid #c7eadf;padding:8px 12px;border-radius:12px;font-weight:700}.facility-summary{display:flex;gap:8px;flex-wrap:wrap;margin:0 0 16px}.facility-chip{display:flex;gap:8px;align-items:center;border:1px solid;padding:8px 12px;border-radius:12px}.facility-chip b{font-size:12px}.facility-chip span{font-size:11px;font-weight:700}.facility-row td{background:var(--facility-bg)!important;-webkit-print-color-adjust:exact;print-color-adjust:exact}table{width:100%;border-collapse:separate;border-spacing:0;font-size:11px;overflow:hidden;border:1px solid #dbe5ec;border-radius:12px}th{background:#173b57;color:white;padding:9px 6px}td{padding:8px 6px;border-bottom:1px solid #e6edf2;text-align:center}tr:nth-child(even) td{background:#f8fbfd}tr.good td,tr.achieved td{background:#effaf5}tr.warning td{background:#fff8e6}tr.risk td{background:#fff0f0}.index{font-weight:700;color:#64748b}.footer{margin-top:12px;color:#64748b;font-size:10px}@media print{*{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}body{background:white}.sheet{padding:0}.no-print{display:none}}
    </style></head><body><div class="sheet"><div class="head"><div><h1>${esc(title)}</h1><p>${esc(subtitle)}</p></div><div class="brand">IML CONTROL</div></div><div class="summary"><div class="chip">חודש: ${esc(planningMonth || '—')}</div><div class="chip">תאריך הדפסה: ${esc(new Date().toLocaleString('he-IL'))}</div><div class="chip">רשומות: ${rows.length}</div></div>${facilitySummaryHtml}<table><thead><tr><th>#</th>${headers.map(h=>`<th>${esc(h.label)}</th>`).join('')}</tr></thead><tbody>${bodyRows}</tbody></table><div class="footer">דוח ניהולי — IML CONTROL</div></div><script>window.onload=()=>setTimeout(()=>window.print(),250)<\/script></body></html>`)
    popup.document.close()
  }

  const planningDisplayStation = row => /^EC\s*\(23\)/i.test(normalize(row?.resource)) ? '1523' : (row?.station || row?.facility || '—')

  const printDailyManagement = () => openPrintReport(
    'Daily Management — ניהול יומי',
    'יעדים, ביצוע, תחזית וקצב נדרש לכל משאב פעיל',
    [
      {key:'resource',label:'משאב יעד'},{key:'station',label:'תחנה'},{key:'line',label:'קו / משאב'},
      {key:'target',label:'יעד חודשי'},{key:'actual',label:'בפועל'},{key:'pct',label:'% ביצוע'},
      {key:'remaining',label:'נותר'},{key:'days',label:'ימים נותרו'},{key:'required',label:'נדרש ליום'},
      {key:'avg7',label:'ממוצע 7 ימים'},{key:'forecast',label:'תחזית'},{key:'status',label:'סטטוס'}
    ],
    dailyPlanningRows.map(r => ({_state:r.state,resource:r.resource||r.facility,station:planningDisplayStation(r),line:r.lineName||r.routingGroup||'—',target:fmt(r.target),actual:fmt(r.actual),pct:pctFmt(r.pct),remaining:fmt(r.remaining),days:r.remainingWorkdays,required:fmt(r.requiredDaily),avg7:fmt(r.recentAverage),forecast:fmt(r.forecast),status:r.label}))
  )


  const selectedFacilityLabel = () => selectedFacilities.length ? selectedFacilities.join(', ') : 'כל המתקנים'
  const rowMatchesSelectedFacilities = row => {
    if (!selectedFacilities.length) return true
    const ids = row.facilityIds || row.facilities || [row.station, row.facility].filter(Boolean)
    return ids.some(id => selectedFacilities.includes(String(id)))
  }
  const printFacilityOverview = () => {
    const rows = visibleControlTowerFacilities.filter(rowMatchesSelectedFacilities)
    openPrintReport('סקירת מתקנים', `מתקנים: ${selectedFacilityLabel()}`, [
      {key:'resource',label:'משאב / סימניה'},{key:'station',label:'תחנה'},{key:'target',label:'יעד'},{key:'actual',label:'בוצע'},
      {key:'forecast',label:'תחזית'},{key:'gap',label:'פער צפוי'},{key:'required',label:'קצב נדרש'},{key:'health',label:'Health Score'},{key:'deviations',label:'חריגות'}
    ], rows.map(r => ({_state:r.state,resource:r.resource||r.facility,station:(r.facilityIds||r.facilities||[]).join(', ')||r.station||'—',target:fmt(r.target),actual:fmt(r.actual),forecast:fmt(r.forecast),gap:fmt(r.gap),required:fmt(r.requiredDaily),health:r.healthScore,deviations:r.deviationsCount})))
  }
  const printMonthlyForecast = () => {
    const rows = visiblePlanningRows.filter(rowMatchesSelectedFacilities)
    openPrintReport('תחזית חודשית לפי מתקן', `מתקנים: ${selectedFacilityLabel()}`, [
      {key:'resource',label:'משאב / סימניה'},{key:'station',label:'תחנה'},{key:'target',label:'יעד חודשי'},{key:'actual',label:'בוצע'},
      {key:'pct',label:'% ביצוע'},{key:'remaining',label:'נותר'},{key:'required',label:'נדרש ליום'},{key:'avg7',label:'ממוצע 7 ימים'},{key:'max',label:'שיא מוכח'},{key:'forecast',label:'תחזית'},{key:'status',label:'סטטוס'}
    ], rows.map(r => ({_state:r.state,resource:r.resource||r.facility,station:planningDisplayStation(r),target:fmt(r.target),actual:fmt(r.actual),pct:pctFmt(r.pct),remaining:fmt(r.remaining),required:fmt(r.requiredDaily),avg7:fmt(r.recentAverage),max:fmt(r.provenMax),forecast:fmt(r.forecast),status:r.label})))
  }
  const printFacilityPerformance = () => {
    const rows = facilityStats.filter(r => !selectedFacilities.length || selectedFacilities.includes(r.id))
    openPrintReport('ביצועים לפי מתקן בטווח המסונן', `מתקנים: ${selectedFacilityLabel()} · ${from || '—'} עד ${to || '—'}`, [
      {key:'facility',label:'מתקן'},{key:'qty',label:'כמות'},{key:'orders',label:'Orders'},{key:'batches',label:'מנות'}
    ], rows.map(r => ({_facility:r.id,_qtyRaw:num(r.actual),facility:r.id,qty:fmt(r.actual),orders:r.orders||0,batches:new Set(baseFiltered.filter(x => x.facility === r.id).map(x => x.batch).filter(Boolean)).size})))
  }
  const printRecentProduction = () => {
    const rows = sortedRecentProduction.filter(r => !selectedFacilities.length || selectedFacilities.includes(r.facility))
    const totalQty = rows.reduce((sum, row) => sum + num(row.qty), 0)
    const stationTotals = [...rows.reduce((map, row) => {
      const station = normalize(row.routingGroup || row.station || row.facility || 'ללא תחנה') || 'ללא תחנה'
      map.set(station, (map.get(station) || 0) + num(row.qty))
      return map
    }, new Map()).entries()].sort((a,b) => b[1] - a[1])
    const weighingSummary = `סה״כ כמות: ${fmt(totalQty)}${stationTotals.length ? ' · ' + stationTotals.map(([station, qty]) => `${station}: ${fmt(qty)}`).join(' · ') : ''}`
    openPrintReport('רשומות תפוקה אחרונות', `מתקנים: ${selectedFacilityLabel()} · ${from || '—'} עד ${to || '—'} · ${weighingSummary}`, [
      {key:'date',label:'תאריך'},{key:'ud',label:'החלטת שימוש (UD)'},{key:'facility',label:'משאב יעד'},{key:'routing',label:'מתקן / תחנה'},
      {key:'order',label:'הזמנה'},{key:'batch',label:'Batch'},{key:'material',label:'מק״ט חומר'},{key:'desc',label:'תיאור חומר'},{key:'planned',label:'כמות מתוכננת'},{key:'qty',label:'כמות בפועל'},{key:'gap',label:'פער'}
    ], rows.map(r => ({_facility:r.facility,_qtyRaw:num(r.qty),date:iso(r.date),ud:productionUsageDecision(r),facility:r.facility,routing:r.routingGroup||'—',order:r.order||'—',batch:r.batch||'—',material:r.material||'—',desc:r.desc||'—',planned:r.plannedQty?fmt(r.plannedQty):'—',qty:fmt(r.qty),gap:r.plannedQty?fmt(num(r.qty)-num(r.plannedQty)):'—'})))
  }

  const printMonthlyTargets = () => {
    const rows = targets.filter(t => !planningMonth || t.month === planningMonth)
    openPrintReport('קובץ יעדים חודשי', 'תצוגת יעדים ושיוכי תחנות לצורך בקרה ועדכון', [
      {key:'resource',label:'משאב / יעד'},{key:'station',label:'תחנה'},{key:'facilities',label:'מתקנים משויכים'},
      {key:'target',label:'יעד חודשי'},{key:'capacity',label:'קיבולת'},{key:'notes',label:'הערות'}
    ], rows.map(t => ({resource:t.resource||'—',station:t.station||t.facility||'—',facilities:(t.facilities||[t.facility]).filter(Boolean).join(', '),target:fmt(t.target),capacity:fmt(t.capacity),notes:t.notes||''})))
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

  const openHomeArea = (area) => {
    setShowHome(false)
    if (area === 'quality') setActiveTab('quality')
    if (area === 'production' || area === 'recent') setActiveTab('production')
    if (area === 'admin') setActiveTab(canManageData ? 'mapping-center' : 'production')
    window.setTimeout(() => {
      const targetId = area === 'planning' ? 'planning-section' : area === 'daily' ? 'daily-management-section' : area === 'quality' || area === 'recent' || area === 'admin' ? 'details-section' : 'control-tower-section'
      document.getElementById(targetId)?.scrollIntoView({ behavior:'smooth', block:'start' })
    }, 80)
  }

  const updateBanner = availableUpdate ? <div role="alert" style={{margin:'10px 18px',padding:'12px 16px',borderRadius:12,background:'#fff3cd',border:'1px solid #f0c36d',display:'flex',alignItems:'center',justifyContent:'space-between',gap:14,boxShadow:'0 4px 14px rgba(15,35,55,.08)',direction:'rtl'}}>
    <div style={{display:'flex',alignItems:'center',gap:10}}><RefreshCw size={19}/><div><strong style={{display:'block'}}>קיים עדכון חדש ל-IML CONTROL</strong><small>{availableUpdate.version && availableUpdate.version !== 'חדשה' ? `גרסה חדשה ${availableUpdate.version} · ` : ''}נמצא Build חדש בשרת · נדרש רענון אפליקציה</small></div></div>
    <button type="button" onClick={refreshApplication} style={{border:0,borderRadius:9,padding:'9px 15px',fontWeight:800,cursor:'pointer',background:'#0f8f7d',color:'#fff',whiteSpace:'nowrap'}}>רענן עכשיו</button>
  </div> : null

  if (showHome) return <div className="command-home" dir="rtl">
    <header className="command-home-header">
      <div className="command-home-brand"><img src="/icons/adama-mark-128.png" alt="IML"/><div><strong>חדר בקרה — מתקני אריזה</strong><span>COMMAND CENTER</span></div></div>
      <button type="button" className="command-home-sound-toggle" data-no-ui-sound="1" onClick={() => { setUiSoundsEnabled(v => !v); playUiTone(uiSoundsEnabled ? 'close' : 'success') }} title={uiSoundsEnabled ? 'כיבוי צלילים' : 'הפעלת צלילים'}>{uiSoundsEnabled ? <Volume2 size={18}/> : <VolumeX size={18}/>}</button><div className="command-home-user"><Home size={19}/><b>דף ראשי</b><span></span><div><strong>{isGuest ? 'אורח' : (currentUser?.email || 'משתמש')}</strong><small>{isGuest ? 'צפייה בלבד' : userRole === 'admin' ? 'מנהל מערכת' : userRole === 'manager' ? 'מנהל מתקן' : 'צפייה בלבד'}</small></div></div>
    </header>
    {updateBanner}
    <main className="command-home-main">
      <section className="command-home-status">
        <article><div className={`home-status-icon cloud ${cloudState.mode}`}><Cloud/></div><span>מצב מערכת</span><b>{cloudState.mode === 'cloud' ? 'מחובר לענן' : cloudState.mode === 'connecting' ? 'מתחבר...' : 'מצב מקומי'}</b><small>{cloudState.lastSync ? `עדכון אחרון ${new Date(cloudState.lastSync).toLocaleString('he-IL')}` : 'ממתין לסנכרון'}</small></article>
        <article><Database/><span>כמויות</span><b>{fmt(production.length)}</b><small>רשומות</small></article>
        <article><FlaskConical/><span>איכות</span><b>{fmt(quality.length)}</b><small>תוצאות</small></article>
        <article><AlertTriangle/><span>מנות חריגות</span><b>{fmt(openDeviations.length)}</b><small>פתוחות בטווח</small></article>
        <article className="home-range"><CalendarDays/><span>טווח תאריכים נוכחי</span><b>{from || '—'} — {to || '—'}</b><small>{planningMonth || 'ללא חודש תכנון'}</small></article>
      </section>
      <section className="command-home-welcome"><span>IML CONTROL</span><h1>ברוך הבא לחדר הבקרה</h1><p>בחר אזור לצפייה ולניתוח נתונים</p></section>
      <section className="command-home-cards">
        <button className="home-nav-card teal" onClick={()=>openHomeArea('production')}><div className="home-nav-icon"><BarChart3/></div><h2>סקירת מתקנים</h2><p>סקירה כוללת של כל המתקנים, ביצועים מול יעדים ותחזית חודשית</p><span>כניסה לאזור <ChevronLeft/></span></button>
        <button className="home-nav-card green" onClick={()=>openHomeArea('daily')}><div className="home-nav-icon"><ClipboardList/></div><h2>ניהול יומי</h2><p>ניהול ותפעול יומי של מתקנים, קווים, תקלות והערות</p><span>כניסה לאזור <ChevronLeft/></span></button>
        <button className="home-nav-card purple" onClick={()=>openHomeArea('planning')}><div className="home-nav-icon"><Factory/></div><h2>תחזית חודשית לפי מתקן</h2><p>תחזית ביצועים חודשית לפי מתקן עם ניתוח קצבים ועמידה ביעדים</p><span>כניסה לאזור <ChevronLeft/></span></button>
        <button className="home-nav-card orange" onClick={()=>openHomeArea('quality')}><div className="home-nav-icon"><FlaskConical/></div><h2>איכות</h2><p>סקירת איכות, מנות חריגות, תוצאות מעבדה וכרטיסי מנה</p><span>כניסה לאזור <ChevronLeft/></span></button>
        <button className="home-nav-card blue" onClick={()=>openHomeArea('recent')}><div className="home-nav-icon"><ClipboardList/></div><h2>רשומות תפוקה אחרונות</h2><p>רשומות התפוקה האחרונות ומעקב אחר המנות והחומרים שיוצרו</p><span>כניסה לאזור <ChevronLeft/></span></button>
      </section>
      <section className="command-home-bottom">
        <article className="home-attention"><div className="home-bottom-title"><BellRing/><h3>דורש תשומת לב</h3></div>{openDeviations.length ? <div className="home-alert-row"><AlertTriangle/><b>{openDeviations.length} מנות חריגות פתוחות</b><span>מומלץ לעבור למסך האיכות לבדיקה</span><button onClick={()=>openHomeArea('quality')}>לצפייה <ChevronLeft/></button></div> : <div className="home-ok"><CheckCircle2/> אין כרגע חריגות פתוחות בטווח שנבחר</div>}</article>
        <article className="home-quick"><div className="home-bottom-title"><Activity/><h3>פעולות מהירות</h3></div><div className="home-quick-grid"><button onClick={()=>openHomeArea('production')}><BarChart3/> סקירת מתקנים</button><button onClick={()=>openHomeArea('quality')}><FlaskConical/> איכות</button><button onClick={()=>openHomeArea('planning')}><Target/> יעדים ותחזית</button>{canManageData && <button onClick={()=>openHomeArea('admin')}><Settings2/> מרכז ניהול</button>}</div></article>
      </section>
    </main>
    <footer className="command-home-footer"><span>IML CONTROL © 2026</span><b>{BUILD_LABEL}</b><span className={cloudState.mode === 'cloud' ? 'online' : ''}>● {cloudState.mode === 'cloud' ? 'ONLINE' : 'OFFLINE'}</span></footer>
  </div>

  return <div className={`dashboard ${sidebarCollapsed ? 'sidebar-collapsed' : ''} ${managementMode ? 'management-mode' : ''}`} dir="rtl">
    <aside className="side filter-side">
      <button type="button" className="side-collapse-button" onClick={() => setSidebarCollapsed(v => !v)} title={sidebarCollapsed ? 'פתיחת מסננים' : 'כיווץ מסננים'}>{sidebarCollapsed ? <PanelRightOpen size={18}/> : <PanelRightClose size={18}/>}</button>
      <div className="brand branded"><img src="/icons/adama-mark-128.png" alt="IML"/><div>IML<span>CONTROL</span></div></div>
      <div className="side-filter-title"><Search size={18}/><strong>חיפוש וסינון</strong></div>
      <label className="side-field"><span>Quick Search</span><div><Search size={16}/><input value={query} onChange={e => setQuery(e.target.value)} placeholder="Order, Batch, חומר..."/></div></label>
      <label className="side-field"><span>חודש תכנון</span><select value={planningMonth} onChange={e => setPlanningMonth(e.target.value)}>{!availableMonths.length && <option value="">אין נתונים</option>}{availableMonths.map(m => <option key={m} value={m}>{m}</option>)}</select></label>
      <label className="side-field"><span>מתאריך</span><input type="date" min={dateBounds.min} max={dateBounds.max} value={from} onChange={e => { setFrom(e.target.value); setPeriodYear(''); setPeriodQuarter('') }}/></label>
      <label className="side-field"><span>עד תאריך</span><input type="date" min={dateBounds.min} max={dateBounds.max} value={to} onChange={e => { setTo(e.target.value); setPeriodYear(''); setPeriodQuarter('') }}/></label>
      <div className="side-field facility-multi-field"><span>מתקנים</span><button type="button" className={`facility-multi-trigger ${selectedFacilities.length ? 'active' : ''}`} onClick={() => setFacilityPickerOpen(v => !v)}><Factory size={16}/><b>{selectedFacilities.length ? `${selectedFacilities.length} מתקנים נבחרו` : 'לא נבחר מתקן'}</b><span>⌄</span></button>{facilityPickerOpen && <div className="facility-multi-menu"><div className="facility-multi-menu-head"><strong>בחירת מתקנים</strong><button type="button" onClick={() => setSelectedFacilities([])}>נקה הכול</button></div>{facilities.map(id => <label key={id}><input type="checkbox" checked={selectedFacilities.includes(id)} onChange={() => toggleFacility(id)}/><span>{id}</span></label>)}</div>}</div>
      <div className="side-quick-ranges"><button onClick={() => setQuickRange(1)}>יום</button><button onClick={() => setQuickRange(2)}>יומיים</button><button onClick={() => setQuickRange(30)}>30 יום</button></div>
      <button className={`side-clear ${showDataCenter ? 'active' : ''}`} type="button" onClick={() => { if (!canManageData) { sessionStorage.setItem('iml-open-data-center-after-login','1'); onRequestAdminLogin?.(); return } setShowDataCenter(v => !v); window.setTimeout(() => document.getElementById('data-center-section')?.scrollIntoView({behavior:'smooth', block:'start'}), 60) }}><Database size={16}/> מרכז נתונים</button>
      <button className="side-clear" onClick={() => { setFrom(''); setTo(''); setQuery(''); setSelectedFacilities([]); setPeriodYear(''); setPeriodQuarter('') }}><X size={16}/> ניקוי מסננים</button>
      <div className="side-live-stats"><div><Database/><span><b>{fmt(production.length)}</b><small>תפוקה</small></span></div><div><FlaskConical/><span><b>{fmt(quality.length + deviations.length)}</b><small>איכות</small></span></div></div>
      <div className="side-note">{BUILD_LABEL} · {userRole === 'admin' ? 'Admin' : userRole === 'manager' ? 'Manager' : 'Viewer'}</div>
    </aside>

    <main className="main" id="control-tower-section">
      <header className="header">
        <div><h1>חדר בקרה — מתקני אריזה</h1><p>{BUILD_LABEL}</p></div>
        <div className="header-actions">
          <div className="ui-refresh-controls"><button type="button" className="action ui-control home-return" onClick={() => setShowHome(true)}><Home size={18}/><span>דף ראשי</span></button>
            <button type="button" className="action ui-control" onClick={() => setSidebarCollapsed(v => !v)}>{sidebarCollapsed ? <PanelRightOpen size={18}/> : <PanelRightClose size={18}/>}<span>{sidebarCollapsed ? 'פתח מסננים' : 'כווץ מסננים'}</span></button>
            <button type="button" className={`action ui-control ${managementMode ? 'active' : ''}`} onClick={() => setManagementMode(v => !v)}>{managementMode ? <Minimize2 size={18}/> : <Maximize2 size={18}/>}<span>{managementMode ? 'יציאה ממצב ניהולי' : 'מצב ניהולי'}</span></button>
          </div>
          <button type="button" className={`action ui-control ${uiSoundsEnabled ? 'active' : ''}`} data-no-ui-sound="1" onClick={() => { setUiSoundsEnabled(v => !v); playUiTone(uiSoundsEnabled ? 'close' : 'success') }} title={uiSoundsEnabled ? 'כיבוי צלילי ממשק' : 'הפעלת צלילי ממשק'}>{uiSoundsEnabled ? <Volume2 size={18}/> : <VolumeX size={18}/>}<span>{uiSoundsEnabled ? 'צלילים פעילים' : 'צלילים כבויים'}</span></button><div className="user-session"><img className="user-brand-avatar" src="/icons/adama-mark-64.png" alt="IML"/><span><b>{isGuest ? 'אורח' : (currentUser?.email || 'משתמש')}</b><small>{isGuest ? 'צפייה בלבד' : userRole === 'admin' ? 'מנהל מערכת' : userRole === 'manager' ? 'מנהל מתקן' : 'צפייה בלבד'}</small></span></div>
          <button className="action secondary" onClick={printMonthlyTargets} disabled={!targets.length}><Printer size={18}/> הדפסת יעדים</button>
          <button className="action secondary" onClick={downloadTargetWorkbook}><FileSpreadsheet size={18}/> הורדת תבנית יעדים</button>
          <div className="daily-event-export-controls">
            <button type="button" className="action secondary event-action-button" onClick={openDailyEventForm}><BellRing size={18}/> הוספת אירוע</button>
            <button type="button" className="action secondary event-action-button" onClick={() => setDailyEventHistoryOpen(true)}><ClipboardList size={18}/> היסטוריית אירועים{visibleDailyEvents.length ? ` (${visibleDailyEvents.length})` : ''}</button>
          </div>
          <button className="action secondary" onClick={exportWorkbook} disabled={!production.length}><Download size={18}/> יצוא Excel</button>
          <label className={`action secondary ${busy ? 'disabled' : ''}`} style={{cursor:'pointer'}}><Upload size={18}/> טעינת דוח ערוך<input type="file" accept=".xlsx,.xls" hidden disabled={busy} onChange={e=>{const file=e.target.files?.[0]; if(file) handleDailyReportRoundtripFile(file); e.target.value=''}}/></label>
          <button className="action secondary" type="button" onClick={downloadSavedDailyReport} disabled={!selectedSingleReportDate}><Download size={18}/> הורדת דוח שמור</button>
          {canDeleteData && <button className="action danger" onClick={clearAllData} disabled={!production.length && !quality.length && !deviations.length && !targets.length}><Trash2 size={18}/> מחיקה</button>}
          {canManageData ? <label className={`upload ${busy ? 'disabled' : ''}`}><Upload size={19}/>{busy ? 'טוען...' : 'טעינת Excel'}<input type="file" multiple accept=".xlsx,.xls" disabled={busy} onChange={e => handleFiles([...e.target.files])}/></label> : <button className="action upload" type="button" onClick={onRequestAdminLogin}><Upload size={19}/> טעינת Excel</button>}
          {canManageData ? <button className="action secondary" onClick={onSignOut}><LogOut size={18}/> יציאת מנהל</button> : <button className="action secondary" onClick={onRequestAdminLogin}><ShieldCheck size={18}/> כניסת מנהל</button>}
        </div>
      </header>
      {dailyEventFormOpen && <div className="daily-event-modal-backdrop" onMouseDown={event => { if (event.target === event.currentTarget) setDailyEventFormOpen(false) }}>
        <section className="daily-event-modal" dir="rtl" role="dialog" aria-modal="true" aria-label="הוספת אירוע">
          <div className="daily-event-modal-head"><div><small>דיווח יומי</small><h2>הוספת אירוע</h2><p>מלא את פרטי האירוע ושמור. ניתן לשמור יותר מאירוע אחד לאותו יום.</p></div><button type="button" className="daily-event-close" onClick={() => setDailyEventFormOpen(false)} aria-label="סגור"><X size={20}/></button></div>
          <div className="daily-event-form-grid">
            <label><span>תאריך אירוע</span><input type="date" value={dailyEventDate} onChange={e=>setDailyEventDate(e.target.value)}/></label>
            <label><span>סוג אירוע</span><select value={dailyEventType} onChange={e=>setDailyEventType(e.target.value)}><option value="">בחר סוג אירוע...</option><option>שפך</option><option>בטיחות</option><option>איכות</option><option>סביבה</option><option>תפעולי</option><option>אחר</option></select></label>
            <label><span>מתקן</span><select value={dailyEventFacility} onChange={e=>setDailyEventFacility(e.target.value)}><option value="">בחר מתקן...</option>{facilities.map(id=><option key={id} value={id}>{id}</option>)}</select></label>
            <label><span>חומרת אירוע</span><select value={dailyEventSeverity} onChange={e=>setDailyEventSeverity(e.target.value)}><option value="">בחר חומרה...</option><option>נמוכה</option><option>בינונית</option><option>גבוהה</option><option>קריטית</option></select></label>
            <label className="daily-event-description"><span>תיאור האירוע</span><textarea rows="5" value={dailyEventText} onChange={e=>setDailyEventText(e.target.value)} placeholder="תאר בקצרה מה קרה, היכן ומה הפעולה שבוצעה..."/></label>
          </div>
          <div className="daily-event-modal-actions"><button type="button" className="action secondary" onClick={() => setDailyEventFormOpen(false)}>ביטול</button><button type="button" className="action event-save-button" onClick={saveDailyEvent}><Save size={18}/> שמור אירוע</button></div>
        </section>
      </div>}

      {dailyEventHistoryOpen && <div className="daily-event-modal-backdrop" onMouseDown={event => { if (event.target === event.currentTarget) setDailyEventHistoryOpen(false) }}>
        <section className="daily-event-modal daily-event-history-modal" dir="rtl" role="dialog" aria-modal="true" aria-label="היסטוריית אירועים">
          <div className="daily-event-modal-head"><div><small>היסטוריה לפי המסננים הפעילים</small><h2>היסטוריית אירועים</h2><p>{from || 'ללא תאריך התחלה'} עד {to || 'ללא תאריך סיום'}{selectedFacilities.length ? ` · מתקנים ${selectedFacilities.join(', ')}` : ' · כל המתקנים'}</p></div><button type="button" className="daily-event-close" onClick={() => setDailyEventHistoryOpen(false)} aria-label="סגור"><X size={20}/></button></div>
          <div className="daily-event-history-list">{visibleDailyEvents.length ? visibleDailyEvents.map(event => <article key={event.id} className={`daily-event-history-card severity-${event.severity}`}><div className="daily-event-history-top"><strong>{event.type}</strong><span>{new Date(`${event.date}T12:00:00`).toLocaleDateString('he-IL')}</span></div><div className="daily-event-history-meta"><span>מתקן {event.facility}</span><b>{event.severity}</b></div><p>{event.description}</p><small>{event.createdAt ? `נשמר ${new Date(event.createdAt).toLocaleString('he-IL')}` : ''}</small></article>) : <div className="daily-event-history-empty">לא נמצאו אירועים בטווח ובמתקנים שנבחרו.</div>}</div>
          <div className="daily-event-modal-actions"><button type="button" className="action secondary" onClick={() => setDailyEventHistoryOpen(false)}>סגור</button><button type="button" className="action event-save-button" onClick={() => { setDailyEventHistoryOpen(false); openDailyEventForm() }}><BellRing size={18}/> אירוע חדש</button></div>
        </section>
      </div>}

      {updateBanner}

      <section className="top-status-bar" aria-label="סטטוס מערכת">
        <div className={`top-status-item cloud ${cloudState.mode}`}><span className="status-dot"></span><small>ענן</small><b>{cloudState.mode === 'cloud' ? 'מחובר' : cloudState.mode === 'connecting' ? 'מתחבר' : 'מקומי'}</b></div>
        <div className="top-status-item"><small>עדכון אחרון</small><b>{cloudState.lastSync ? new Date(cloudState.lastSync).toLocaleTimeString('he-IL', {hour:'2-digit', minute:'2-digit'}) : '—'}</b></div>
        <div className="top-status-item"><small>חודש פעיל</small><b>{planningMonth || '—'}</b></div>
        <div className="top-status-item"><small>תפוקה</small><b>{fmt(dashboardProd.length)}</b></div>
        <div className="top-status-item"><small>איכות</small><b>{fmt(dashboardQualityRows.length + dashboardDeviationRows.length)}</b></div>
        <div className="top-status-item range"><small>טווח</small><b>{from || '—'} → {to || '—'}</b></div>
        {canManageData && <button type="button" className={`diagnostics-toggle ${diagnosticsOpen ? 'active' : ''}`} onClick={() => setDiagnosticsOpen(v => !v)} title="אבחון מערכת"><Activity size={16}/><span>אבחון</span></button>}
      </section>

      {(busy || cloudState.mode !== 'cloud' || /נכשל|שגיאה|ממתין|מתחבר|בודק|טוען|לא זמין/.test(status)) && <div className={`load-status compact ${cloudState.mode === 'error' || /נכשל|שגיאה/.test(status) ? 'error' : ''}`}><CheckCircle2 size={17}/>{status}</div>}
      {uploadProgress && <section className="upload-progress-card"><div className="upload-progress-head"><strong>{uploadProgress.fileName}</strong><span>{uploadProgress.percent}%</span></div><div className="upload-progress-track"><div style={{width:`${uploadProgress.percent}%`}}/></div><small>{uploadProgress.message}</small></section>}

      {canManageData && diagnosticsOpen && <section className="diagnostics-panel">
        <div className="diagnostics-panel-head"><div><Activity size={18}/><strong>אבחון מערכת</strong></div><button type="button" onClick={() => setDiagnosticsOpen(false)}><X size={16}/> סגור</button></div>
        <div className="diagnostics-grid">
          <article><small>Cache</small><b>{perfStats.cache}</b></article>
          <article><small>Queries</small><b>{perfStats.queries}</b></article>
          <article><small>Production</small><b>{production.length.toLocaleString()}</b></article>
          <article><small>Quality</small><b>{quality.length.toLocaleString()}</b></article>
          <article><small>In range</small><b>{(filtered.length + filteredQualityRows.length + filteredDeviationRows.length).toLocaleString()}</b></article>
          <article><small>Load</small><b>{perfStats.loadMs ? `${perfStats.loadMs}ms` : perfStats.phase}</b></article>
          <article><small>Data source</small><b>{cloudState.mode === 'cloud' ? 'Supabase' : 'Browser Cache'}</b></article>
          <article><small>Realtime</small><b>{cloudState.live ? 'פעיל' : 'לא פעיל'}</b></article>
          <article><small>Latency</small><b>{cloudState.latencyMs != null ? `${cloudState.latencyMs}ms` : '—'}</b></article>
          <article className="wide"><small>Range</small><b>{from || '—'} → {to || '—'}</b></article>
          <article className="wide"><small>Last sync</small><b>{cloudState.lastSync ? new Date(cloudState.lastSync).toLocaleString('he-IL') : '—'}</b></article>
        </div>
      </section>}

      {cloudState.mode !== 'cloud' && <section className={`cloud-status ${cloudState.mode}`}>
        <div className="cloud-status-icon">{cloudState.mode === 'offline' ? <WifiOff/> : <Cloud/>}</div>
        <div><strong>{cloudState.mode === 'connecting' ? 'מתחבר לענן' : cloudState.mode === 'offline' ? 'מצב מקומי זמני' : 'שגיאת סנכרון'}</strong><span>{cloudState.message}</span></div>
        <div className="cloud-status-meta"><small>מקור נתונים</small><b>{cloudState.mode === 'offline' ? 'Browser Cache' : 'Supabase'}</b><small>{cloudState.live ? '● עדכון חי פעיל' : '○ עדכון חי לא פעיל'}</small>{cloudState.latencyMs != null && <small>זמן תגובה: {cloudState.latencyMs}ms</small>}{cloudState.lastSync && <small>סנכרון: {new Date(cloudState.lastSync).toLocaleString('he-IL')}</small>}</div>
      </section>}


      <section className="control-tower-hero">
        <div className="control-tower-title"><div><span className="tower-kicker">PACKAGING CONTROL TOWER</span><h2>תמונת מצב ניהולית בזמן אמת</h2><p>יעדים, תחזית חודשית, איכות והזמנות — במבט אחד</p></div><div className={`tower-online ${cloudState.mode === 'cloud' ? 'online' : ''}`}><span></span>{cloudState.mode === 'cloud' ? 'ONLINE' : 'OFFLINE'}</div></div>
        <div className="tower-kpis">
          <button onClick={() => document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'})}><Target/><span>יעד חודשי כולל</span><b>{fmt(targetTotal)}</b><small>{planningMonth || 'ללא חודש נבחר'}</small></button>
          <button onClick={() => document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'})}><Database/><span>בוצע החודש</span><b>{fmt(targetActual)}</b><small>{targetTotal ? pctFmt(targetActual / targetTotal * 100) : '—'} מהיעד</small></button>
          <button onClick={() => document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'})}><TrendingUp/><span>תחזית סוף חודש</span><b>{fmt(targetForecast)}</b><small>{targetTotal ? pctFmt(targetForecast / targetTotal * 100) : '—'} מהיעד</small></button>
          <button onClick={() => document.getElementById('alerts-section')?.scrollIntoView({behavior:'smooth'})}><Gauge/><span>קצב יומי נדרש</span><b>{fmt(targetRequiredDaily)}</b><small>{selectedFacilities.length ? `ל-${selectedFacilities.length} מתקנים שנבחרו` : 'לכל המתקנים'}</small></button>
          <button className={targetForecast >= targetTotal && targetTotal ? 'good' : 'bad'} onClick={() => document.getElementById('alerts-section')?.scrollIntoView({behavior:'smooth'})}><AlertTriangle/><span>פער צפוי</span><b>{targetTotal ? fmt(targetForecast-targetTotal) : '—'}</b><small>{targetForecast >= targetTotal ? 'מעל היעד' : 'נדרש להגביר קצב'}</small></button>
        </div>
      </section>

      <section className="tower-facility-section">
        <div className="panel-head facility-panel-head"><div><Factory/><h2>סקירת מתקנים</h2></div><div className="facility-head-actions">{facilityViewFilters}<span>{selectedControlTowerFacilities.length} מתקנים נבחרים</span><button type="button" className="section-print-btn" onClick={printFacilityOverview}><Printer size={16}/> הדפסה</button><button type="button" className="section-toggle" onClick={() => setShowFacilityOverview(v => !v)}>{showFacilityOverview ? 'הסתר' : 'הצג'}</button></div></div>
        {showFacilityOverview && selectedFacilities.length > 0 && <div className="tower-facility-grid">
          {selectedFacilities.includes('1542') && <Facility42BalanceOverviewCard balance={facility42Balance} onClick={() => jumpToDetails('bulk-balance')}/>}
          {selectedFacilities.includes('1519') && <Facility19BalanceOverviewCard balance={facility19Balance} onClick={() => jumpToDetails('bulk-balance-19')}/>}
          {selectedControlTowerFacilities.map(row => <button key={row.facility} className={`tower-facility-card ${row.state}`} onClick={() => setSelectedResource(row)}>
            <div className="tower-facility-head"><div><i></i><strong>{row.facility}</strong></div><span>{row.state === 'good' ? 'תקין' : row.state === 'warning' ? 'דורש תשומת לב' : row.state === 'risk' ? 'בסיכון' : 'ללא יעד'}</span></div>
            <div className="tower-health"><div><HeartPulse/><span>Health Score</span></div><b>{row.healthScore}<small>/100</small></b></div>
            <div className="tower-progress"><i style={{width:`${Math.min(100,row.actualPct)}%`}}/></div>
            <dl><div><dt>יעד חודשי</dt><dd>{fmt(row.target)}</dd></div><div><dt>בוצע</dt><dd>{fmt(row.actual)}</dd></div><div><dt>תחזית</dt><dd>{fmt(row.forecast)}</dd></div><div><dt>פער צפוי</dt><dd className={row.gap >= 0 ? 'positive' : 'negative'}>{row.gap >= 0 ? '+' : ''}{fmt(row.gap)}</dd></div><div><dt>קצב נדרש</dt><dd>{fmt(row.requiredDaily)}</dd></div><div><dt>חריגות פתוחות</dt><dd>{row.deviationsCount}</dd></div></dl>
            <span className="tower-enter">לפרטים מלאים <ArrowLeft size={16}/></span>
          </button>)}
          {!selectedControlTowerFacilities.length && !selectedFacilities.includes('1542') && !selectedFacilities.includes('1519') && <div className="empty wide-empty">אין נתונים למתקנים שנבחרו.</div>}
        </div>}
        {showFacilityOverview && !selectedFacilities.length && <div className="empty wide-empty">בחר מתקן בתפריט חיפוש וסינון להצגת סקירת מתקנים.</div>}
      </section>

      <section className="tower-lower-grid">
        <article className="tower-alerts-card"><div className="panel-head"><div><BellRing/><h2>התראות אחרונות</h2></div><button onClick={() => document.getElementById('alerts-section')?.scrollIntoView({behavior:'smooth'})}>הצג הכול</button></div><div className="tower-alert-list">{managerInsights.map((item,index)=><button key={index} className={item.state} onClick={() => item.title.includes('חריגות') ? jumpToDetails('deviations') : document.getElementById('planning-section')?.scrollIntoView({behavior:'smooth'})}><i>{item.state==='risk'?'!':item.state==='warning'?'⚠':'✓'}</i><div><strong>{item.title}</strong><span>{item.text}</span></div></button>)}</div></article>
        <article className="tower-chart-card"><div className="panel-head"><div><BarChart3/><h2>ביצוע יומי — 7 ימים אחרונים</h2></div><span>{planningMonth}</span></div><div className="tower-mini-chart">{controlTowerTrend.map(([day,value])=>{const max=Math.max(1,...controlTowerTrend.map(x=>x[1])); return <div key={day}><b>{fmt(value)}</b><span><i style={{height:`${Math.max(8,value/max*100)}%`}}/></span><small>{day.slice(5)}</small></div>})}{!controlTowerTrend.length&&<div className="empty">אין נתונים להצגת מגמה</div>}</div></article>
      </section>

      {showDataCenter && <section className="data-center" id="data-center-section">
        <div className="panel-head"><div><ShieldCheck/><h2>מרכז נתונים</h2></div><span>4 מקורות מידע</span></div>
        <p className="data-center-help">כל קובץ נבדק בדפדפן ולאחר מכן נשמר ב־Supabase. מרגע שהטעינה מסתיימת, אותו מידע זמין לכל המשתמשים המחוברים.</p>
        <div className="data-source-grid">
          <DataSource title="תפוקות" icon={<Factory/>} meta={dataMeta.production} count={production.length} acceptLabel="טען קובץ תפוקות" busy={busy} onFiles={files => loadFiles(files, 'production')} canManage={canManageData}/>
          <DataSource title="תוצאות איכות" icon={<FlaskConical/>} meta={dataMeta.quality} count={quality.length} rows={quality} showYearBreakdown acceptLabel="הוסף תוצאות איכות חדשות" busy={busy} onFiles={files => loadFiles(files, 'quality')} canManage={canManageData}/>
          <DataSource title="חריגות איכות" icon={<AlertTriangle/>} meta={dataMeta.deviations} count={deviations.length} acceptLabel="טען קובץ חריגות" busy={busy} onFiles={files => loadFiles(files, 'deviations')} canManage={canManageData}/>
          <DataSource title="יעדים חודשיים" icon={<Target/>} meta={dataMeta.targets} count={targets.length} acceptLabel="טען קובץ יעדים" busy={busy} onFiles={files => loadFiles(files, 'targets')} canManage={canManageData}/>
        </div>
      </section>}

      {!!dailyReportHistory.length && <section className="details" id="daily-report-history-section">
        <div className="details-title-row"><div><h2>היסטוריית דוחות יומיים שהוחזרו לאפליקציה</h2><p className="details-note">סיכום חודשי לפי מתקן מתוך דוחות Excel שנערכו ידנית ונטענו חזרה · {dailyCloudReady ? 'שמירה משותפת ב־Supabase' : 'מטמון מקומי עד לחיבור לענן'}.</p></div><span className="production-record-count">{dailyReportHistory.length} רשומות</span></div>
        <div className="table-wrap"><table><thead><tr><th>חודש</th><th>מתקן</th><th>דוחות</th><th>רשומות</th><th>סה״כ תפוקה</th><th>סטטוס מכונה</th><th>הערות</th></tr></thead><tbody>{monthlyDailyReportHistory.map(row=><tr key={`${row.month}-${row.facility}`}><td>{row.month}</td><td>{row.facility}</td><td>{row.reports}</td><td>{row.rows}</td><td>{fmt(row.quantity)}</td><td>{row.machineStatuses}</td><td>{row.notes}</td></tr>)}</tbody></table></div>
      </section>}

      <section className="planning-toolbar">
        <div><Target/><span>חודש תכנון</span><select value={planningMonth} onChange={e => setPlanningMonth(e.target.value)}>{!availableMonths.length && <option value="">אין נתונים</option>}{availableMonths.map(m => <option key={m} value={m}>{m}</option>)}</select></div>
        <small>היעד היומי אינו מוזן: הוא מחושב מחדש בכל יום לפי יתרת היעד והימים שנותרו בחודש.</small>
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
        <div><Factory size={18}/><strong>מתקנים נוספים</strong><span>רק מתקנים שמופיעים בבורר המתקנים נכנסים לחישובי הדשבורד. מתקן נוסף ייכלל רק לאחר הוספה מפורשת.</span></div>
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

      <div className="section-title facility-title"><div className="section-title-text"><Gauge/><div><h2>תחזית חודשית לפי מתקן</h2><p>יעד חודשי, קצב נדרש, קצב אחרון, שיא מוכח ותחזית</p></div></div><div className="facility-head-actions">{facilityViewFilters}<button type="button" className="section-print-btn" onClick={printMonthlyForecast}><Printer size={16}/> הדפסה</button><button type="button" className="section-toggle" onClick={() => setShowMonthlyForecast(v => !v)}>{showMonthlyForecast ? 'הסתר' : 'הצג'}</button><button className="select-all-facilities" onClick={toggleAllFacilities}><CheckCircle2 size={17}/>{allFacilitiesSelected ? 'ביטול בחירת הכול' : 'בחירת כל המתקנים'}</button></div></div>
      {showMonthlyForecast && selectedFacilities.length > 0 && <section className="forecast-grid" id="planning-section">
        {selectedFacilities.includes('1542') && <Facility42BalanceCard balance={facility42Balance}/>}
        {selectedFacilities.includes('1519') && <Facility19BalanceCard balance={facility19Balance}/>}
        {selectedForecastRows.map(row => <ForecastCard key={row.id} {...row} selected={(row.facilities || [row.facility]).some(id => selectedFacilities.includes(id))} onClick={() => setSelectedFacilities(row.facilities || [row.facility])}/>) }
        {!selectedForecastRows.length && !selectedFacilities.includes('1542') && !selectedFacilities.includes('1519') && <div className="empty wide-empty">אין נתוני תחזית למתקנים שנבחרו.</div>}
      </section>}
      {showMonthlyForecast && !selectedFacilities.length && <div className="empty wide-empty">בחר מתקן להצגת תחזית חודשית.</div>}

      <section className="alert-panel" id="alerts-section">
        <div className="panel-head"><div><BellRing/><h2>מה דורש תשומת לב היום?</h2></div><span>{alerts.length} התראות</span></div>
        <div className="alert-list">
          {alerts.map(r => <div className={`alert-item ${r.state}`} key={r.id}><div className="alert-symbol">{r.state === 'risk' ? '!' : '⚠'}</div><div><strong>{planningName(r)} — {r.label}</strong><p>{r.state === 'risk' ? `נדרש ${fmt(r.requiredDaily)} ליום, אך השיא המוכח הוא ${fmt(r.provenMax)}.` : `התחזית היא ${fmt(r.forecast)} מול יעד ${fmt(r.target)}. נדרש קצב של ${fmt(r.requiredDaily)} ליום.`}</p></div></div>)}
          {!alerts.length && <div className="empty">{selectedFacilities.length ? 'אין התראות למתקנים שנבחרו.' : 'בחר מתקן להצגת הערות והתראות.'}</div>}
        </div>
      </section>

      <section className="daily-management" id="daily-management-section">
        <div className="panel-head"><div><CalendarCheck/><h2>Daily Management</h2></div><div className="daily-print-actions"><span>{planningMonth}</span><button type="button" className="action secondary daily-print-btn" onClick={printDailyManagement} disabled={!dailyPlanningRows.length}><Printer size={17}/> הדפסה צבעונית</button></div></div>
        <div className="extra-facilities">
          <div><Factory size={18}/><strong>מתקנים ב-Daily Management</strong><span>מתקני הליבה מוצגים תמיד. תחנות 11 אינן מוצגות כברירת מחדל, אך זמינות להוספה ידנית יחד עם כל מתקן נוסף שמופיע בנתונים.</span></div>
          <div className="extra-facility-actions"><select value={dailyFacilityToAdd} onChange={e => setDailyFacilityToAdd(e.target.value)}><option value="">בחר מתקן נוסף</option>{dailyOptionalFacilities.map(id => <option key={id} value={id}>{id}</option>)}</select><button onClick={addDailyFacility} disabled={!dailyFacilityToAdd}>+ הוסף מתקן</button></div>
          {!!dailyAdditionalFacilities.length && <div className="extra-facility-chips">{dailyAdditionalFacilities.map(id => <button key={id} onClick={() => removeDailyFacility(id)}>{id}<X size={14}/></button>)}</div>}
        </div>
        <div className="table-wrap"><table><thead><tr><th>משאב יעד</th><th>מתקן / תחנה</th><th>תחנה / קו</th><th>פעילות</th><th>יעד חודשי</th><th>בפועל</th><th>% ביצוע</th><th>נותר</th><th>ימים נותרו</th><th>נדרש ליום</th><th>ממוצע 7 ימים</th><th>שיא מוכח</th><th>תחזית</th><th>סטטוס</th></tr></thead><tbody>
          {dailyPlanningRows.map(r => <tr key={r.id}><td><b>{r.resource || r.facility}</b></td><td>{planningDisplayStation(r)}</td><td>{[planningDisplayStation(r), r.lineName].filter(Boolean).join(' · ') || '—'}</td><td>{r.activity}</td><td>{fmt(r.target)}</td><td>{fmt(r.actual)}</td><td>{pctFmt(r.pct)}</td><td>{fmt(r.remaining)}</td><td>{r.remainingWorkdays}</td><td>{fmt(r.requiredDaily)}</td><td>{fmt(r.recentAverage)}</td><td>{fmt(r.provenMax)}</td><td>{fmt(r.forecast)}</td><td><StatusBadge state={r.state} label={r.label}/></td></tr>)}
          {!dailyPlanningRows.length && <tr><td colSpan="14" className="empty">אין נתוני Daily Management למתקנים שנבחרו</td></tr>}
        </tbody></table></div>
      </section>

      <section className="trend-card">
        <div className="trend-head"><div><h2>מגמת תפוקה יומית</h2><p>14 הימים האחרונים בטווח הנבחר</p></div><Save size={20}/></div>
        <div className="trend-bars">{dailyTrend.map(([date, value]) => <div className="trend-item" key={date} title={`${date}: ${fmt(value)}`}><div className="trend-value">{fmt(value)}</div><div className="trend-track"><i style={{height: `${Math.max(5, value / maxDaily * 100)}%`}}/></div><small>{date.slice(5)}</small></div>)}{!dailyTrend.length && <div className="empty trend-empty">טען קובץ תפוקות להצגת מגמה</div>}</div>
      </section>

      <div className="section-title facility-title"><div className="section-title-text"><Factory/><div><h2>ביצועים לפי מתקן בטווח המסונן</h2><p>לחיצה על כרטיס מסננת תפוקה, איכות וחריגות</p></div></div><button type="button" className="section-print-btn" onClick={printFacilityPerformance}><Printer size={16}/> הדפסה</button></div>
      <section className="facility-grid">{selectedFacilityStats.map(x => <Facility key={x.id} {...x} selected={selectedFacilities.includes(x.id)} onClick={() => toggleFacility(x.id)}/>) }{!selectedFacilities.length && <div className="empty wide-empty">בחר מתקן להצגת ביצועים בטווח המסונן.</div>}</section>

      <section className="tabs" id="details-section">
        <button className={activeTab === 'production' ? 'active' : ''} onClick={() => setActiveTab('production')}><BarChart3 size={16}/> תפוקה</button>
        <button className={activeTab === 'shifts' ? 'active' : ''} onClick={() => setActiveTab('shifts')}><Clock3 size={16}/> ניתוח משמרות</button>
        <button className={activeTab === 'bulk-balance' ? 'active' : ''} onClick={() => setActiveTab('bulk-balance')}><Activity size={16}/> מאזן מתקן 42</button>
        <button className={activeTab === 'bulk-balance-19' ? 'active' : ''} onClick={() => setActiveTab('bulk-balance-19')}><Activity size={16}/> מאזן מתקן 19</button>
        <button className={activeTab === 'quality' ? 'active' : ''} onClick={() => setActiveTab('quality')}><FlaskConical size={16}/> איכות ({qualityBad.length})</button>
        <button className={activeTab === 'deviations' ? 'active' : ''} onClick={() => setActiveTab('deviations')}><AlertTriangle size={16}/> מנות חריגות ({openDeviations.length})</button>
        {canManageData && <button className={activeTab === 'mapping-simulator' ? 'active' : ''} onClick={() => setActiveTab('mapping-simulator')}><ClipboardList size={16}/> סימולטור שיוך ({mappingSimulation.summary.actionable})</button>}
        {canManageData && <button className={activeTab === 'mapping-center' ? 'active' : ''} onClick={() => setActiveTab('mapping-center')}><ShieldCheck size={16}/> מרכז מיפויים ({manualMappings.filter(item => item.active !== false && item.status === 'pending').length})</button>}
      </section>
      {activeTab === 'bulk-balance' && <section className="details facility42-balance">
        <div className="details-title-row"><div><h2>מאזן תשומות מול תפוקות — מתקן 42</h2><p className="details-note">כרטיסיה עצמאית ללא יעד. תשומה: תחנה 1142 ותיאור המכיל 999. תפוקה: כל האריזות 1, 5, 10/20 ליטר במתקן 42. שאריות: דיווחי 200/1000 ליטר בתחנה 1542.</p></div><span className="production-record-count">{from || 'תחילת נתונים'} — {to || 'היום'}</span></div>
        <div className="balance-kpi-grid">
          <article><span>באלק שיוצר</span><b>{fmt(facility42Balance.bulk)}</b><small>ליטר · {facility42Balance.bulkRows.length} רשומות</small></article>
          <article><span>אריזה 1 ליטר</span><b>{fmt(facility42Balance.byLine['1L'])}</b><small>ליטר</small></article>
          <article><span>אריזה 5 ליטר</span><b>{fmt(facility42Balance.byLine['5L'])}</b><small>ליטר</small></article>
          <article><span>אריזה 10/20 ליטר</span><b>{fmt(facility42Balance.byLine['10/20L'])}</b><small>ליטר</small></article>
          <article className="balance-total"><span>סה״כ נארז</span><b>{fmt(facility42Balance.packed)}</b><small>ליטר</small></article>
          <article><span>שאריות 200/1000 ליטר</span><b>{fmt(facility42Balance.residues)}</b><small>ליטר · {facility42Balance.residueRows.length} רשומות</small></article>
          <article className={facility42Balance.balance < 0 ? 'balance-negative' : 'balance-positive'}><span>יתרה: באלק − אריזה − שאריות</span><b>{fmt(facility42Balance.balance)}</b><small>ליטר</small></article>
          <article><span>% ניצול באלק</span><b>{facility42Balance.bulk ? pctFmt(facility42Balance.utilization) : '—'}</b><small>(אריזה + שאריות) ÷ באלק</small></article>
        </div>
        <div className="balance-note"><AlertTriangle size={18}/><span>בהשוואה יומית ייתכן פער תזמון: באלק שיוצר ביום מסוים יכול להיארז ביום אחר. לכן המאזן החודשי מייצג טוב יותר את התהליך.</span></div>
        <h3 className="shift-subtitle">פירוט לפי סוג דיווח</h3>
        <div className="table-wrap"><table><thead><tr><th>סוג</th><th>מקור</th><th>כמות</th><th>רשומות</th><th>Batch</th><th>Orders</th></tr></thead><tbody>
          <tr><td><b>באלק</b></td><td>1142 + תיאור 999</td><td><b>{fmt(facility42Balance.bulk)}</b></td><td>{facility42Balance.bulkRows.length}</td><td>{new Set(facility42Balance.bulkRows.map(r=>r.batch).filter(Boolean)).size}</td><td>{new Set(facility42Balance.bulkRows.map(r=>r.order).filter(Boolean)).size}</td></tr>
          {['1L','5L','10/20L'].map(line => { const rows=facility42Balance.packedRows.filter(r => { const route=normalize(`${r.routingGroup||''} ${r.routingDescription||''}`).toUpperCase(); return line==='1L' ? (/(^|\s)LQ-P-1(\s|$)/.test(route)||route.includes('42-P-02')||route.includes('LIQUID 1 LITER')) : line==='5L' ? (/(^|\s)LQ-P-5(\s|$)/.test(route)||route.includes('42-P-03')||route.includes('LIQUID 5 LITER')) : !(/(^|\s)LQ-P-(1|5)(\s|$)/.test(route)||route.includes('42-P-02')||route.includes('42-P-03')||route.includes('LIQUID 1 LITER')||route.includes('LIQUID 5 LITER')); }); return <tr key={line}><td><b>אריזה {line}</b></td><td>1542 + ZFIN + Routing</td><td><b>{fmt(facility42Balance.byLine[line])}</b></td><td>{rows.length}</td><td>{new Set(rows.map(r=>r.batch).filter(Boolean)).size}</td><td>{new Set(rows.map(r=>r.order).filter(Boolean)).size}</td></tr> })}
        </tbody></table></div>
      </section>}
      {activeTab === 'bulk-balance-19' && <section className="details facility42-balance facility19-balance">
        <div className="details-title-row"><div><h2>מאזן תשומות מול תפוקות — מתקן 19</h2><p className="details-note">כרטיסיה עצמאית ללא יעד. תשומה: באלק מתחנה 1119 שתיאורו מכיל 777. תפוקה: דיווחי תחנה 1519, בחלוקה ל-WG ולאריזות קטנות.</p></div><span className="production-record-count">{from || 'תחילת נתונים'} — {to || 'היום'}</span></div>
        <div className="balance-kpi-grid">
          <article><span>באלק שיוצר</span><b>{fmt(facility19Balance.bulk)}</b><small>ק״ג · {facility19Balance.bulkRows.length} רשומות</small></article>
          <article><span>מנות ייצור WG</span><b>{fmt(facility19Balance.byType['WG'])}</b><small>ק״ג</small></article>
          <article><span>אריזות קטנות</span><b>{fmt(facility19Balance.byType['SMALL PACKS'])}</b><small>ק״ג</small></article>
          <article className="balance-total"><span>סה״כ תפוקה ארוזה</span><b>{fmt(facility19Balance.packed)}</b><small>ק״ג</small></article>
          <article className={facility19Balance.balance < 0 ? 'balance-negative' : 'balance-positive'}><span>יתרת באלק מול תפוקה</span><b>{fmt(facility19Balance.balance)}</b><small>ק״ג</small></article>
          <article><span>% ניצול באלק</span><b>{facility19Balance.bulk ? pctFmt(facility19Balance.utilization) : '—'}</b><small>תפוקה ÷ באלק</small></article>
        </div>
        <div className="balance-note"><AlertTriangle size={18}/><span>בהשוואה יומית ייתכן פער תזמון בין ייצור הבאלק לבין האריזה. המאזן החודשי מייצג טוב יותר את התהליך.</span></div>
        <h3 className="shift-subtitle">פירוט לפי סוג דיווח</h3>
        <div className="table-wrap"><table><thead><tr><th>סוג</th><th>מקור</th><th>כמות</th><th>רשומות</th><th>Batch</th><th>Orders</th></tr></thead><tbody>
          <tr><td><b>באלק</b></td><td>1119 + תיאור 777</td><td><b>{fmt(facility19Balance.bulk)}</b></td><td>{facility19Balance.bulkRows.length}</td><td>{new Set(facility19Balance.bulkRows.map(r=>r.batch).filter(Boolean)).size}</td><td>{new Set(facility19Balance.bulkRows.map(r=>r.order).filter(Boolean)).size}</td></tr>
          {['WG','SMALL PACKS'].map(type => { const rows=facility19Balance.packedRows.filter(r => (facility19Balance.isSmallPack(r) ? 'SMALL PACKS' : 'WG') === type); return <tr key={type}><td><b>{type === 'WG' ? 'מנות ייצור WG' : 'אריזות קטנות'}</b></td><td>1519 + {type === 'WG' ? 'WG רגיל' : '19PWG-01/05/15'}</td><td><b>{fmt(facility19Balance.byType[type])}</b></td><td>{rows.length}</td><td>{new Set(rows.map(r=>r.batch).filter(Boolean)).size}</td><td>{new Set(rows.map(r=>r.order).filter(Boolean)).size}</td></tr> })}
        </tbody></table></div>
      </section>}
      {activeTab === 'production' && <section className="details"><div className="details-title-row"><h2>רשומות תפוקה אחרונות</h2><div className="details-title-actions"><span className="details-note">לחיצה על כותרת עמודה ממיינת מקטן לגדול / מהגדול לקטן</span><button type="button" className="section-print-btn" onClick={printRecentProduction}><Printer size={16}/> הדפסה</button></div></div><div className="table-wrap"><table className="sortable-production-table" data-smart-sum-column="8" data-smart-group-column="3" data-smart-facility-summary="1"><thead><tr><th><button type="button" onClick={()=>toggleProductionSort('date')}>תאריך{productionSortArrow('date')}</button></th><th><button type="button" onClick={()=>toggleProductionSort('ud')}>החלטת שימוש (UD){productionSortArrow('ud')}</button></th><th><button type="button" onClick={()=>toggleProductionSort('facility')}>משאב יעד{productionSortArrow('facility')}</button></th><th><button type="button" onClick={()=>toggleProductionSort('routingGroup')}>מתקן / תחנה{productionSortArrow('routingGroup')}</button></th><th><button type="button" onClick={()=>toggleProductionSort('order')}>הזמנה{productionSortArrow('order')}</button></th><th><button type="button" onClick={()=>toggleProductionSort('batch')}>Batch{productionSortArrow('batch')}</button></th><th><button type="button" onClick={()=>toggleProductionSort('material')}>מק״ט חומר{productionSortArrow('material')}</button></th><th><button type="button" onClick={()=>toggleProductionSort('desc')}>תיאור חומר{productionSortArrow('desc')}</button></th><th><button type="button" onClick={()=>toggleProductionSort('qty')}>כמות{productionSortArrow('qty')}</button></th></tr></thead><tbody>{sortedRecentProduction.map((r, i) => <tr key={`${r.order}-${r.batch}-${i}`} data-facility={r.facility || ''} style={{backgroundColor: new Set(sortedRecentProduction.map(x => x.facility).filter(Boolean)).size > 1 ? facilityColor(r.facility) : undefined}}><td>{iso(r.date)}</td><td>{productionUsageDecision(r)}</td><td>{r.facility}</td><td>{r.routingGroup || '—'}</td><td>{r.order}</td><td>{r.batch ? <button type="button" className="batch-link" onClick={() => openBatchCard(r.batch, r.material)}>{r.batch}</button> : '—'}</td><td>{r.material || '—'}</td><td>{r.desc || '—'}</td><td><button type="button" className={`qty-variance-btn ${num(r.plannedQty)>0 && Math.abs(num(r.qty)-num(r.plannedQty))>0.0001 ? 'has-variance' : ''}`} onClick={()=>setQuantityVarianceRow(r)} title={num(r.plannedQty)>0 ? 'לחץ להצגת כמות מתוכננת, בפועל והפער' : 'לא נמצאה כמות מתוכננת לרשומה'}>{fmt(r.qty)}</button></td></tr>)}{!sortedRecentProduction.length && <tr className="smart-empty-row"><td colSpan="9" className="empty">אין רשומות להצגה</td></tr>}</tbody></table></div></section>}
      {activeTab === 'mapping-simulator' && canManageData && <section className="details mapping-simulator">
        <div className="mapping-simulator-head"><div><h2>סימולטור שיוך תפוקה</h2><p className="details-note">המסך מתמקד בחריגים שרלוונטיים ליעדים הפעילים. באלק 1142+999 ובאלק 1119+777 מוחרגים אוטומטית ומטופלים רק במאזני 42 ו-19.</p></div><div className="mapping-simulator-actions"><label><input type="checkbox" checked={simulatorOnlyIssues} onChange={event => setSimulatorOnlyIssues(event.target.checked)}/> הצג רק בעיות</label><button type="button" onClick={exportMappingSimulation}><Download size={16}/> ייצוא סימולציה</button></div></div>
        {mappingMessage && <div className="mapping-message">{mappingMessage}</div>}
        <div className="mapping-summary-grid"><article><span>רשומות בטווח</span><b>{fmt(mappingSimulation.summary.rows)}</b><small>{fmt(mappingSimulation.summary.quantity)} כמות</small></article><article className="mapping-ok"><span>שויכו תקין</span><b>{fmt(mappingSimulation.summary.matched)}</b><small>{fmt(mappingSimulation.summary.matchedQty)} כמות</small></article><article className="mapping-duplicate"><span>דורשות החלטה</span><b>{fmt(mappingSimulation.summary.actionable)}</b><small>{fmt(mappingSimulation.summary.actionableQty)} כמות</small></article><article><span>באלק מוחרג</span><b>{fmt(mappingSimulation.summary.ignoredBulk)}</b><small>999 / 777 · מטופל במאזנים</small></article></div>
        <div className="table-wrap mapping-table-wrap"><table><thead><tr><th>סטטוס</th><th>תאריך</th><th>תחנה</th><th>משפחה</th><th>Order</th><th>Batch</th><th>מק״ט</th><th>תיאור מוצר</th><th>כמות</th><th>יעד</th><th>מקור</th><th>פעולה</th></tr></thead><tbody>{visibleMappingSimulation.slice(0,1000).map(item => <tr key={item.key} className={`mapping-row mapping-${item.status}`}><td><span className={`mapping-status mapping-status-${item.status}`}>{item.status === 'matched' ? 'תקין' : item.status === 'duplicate' ? 'כפול' : item.status === 'ignored-bulk' ? 'באלק מוחרג' : 'לא שויך'}</span></td><td>{item.row.productionDay || iso(item.row.date)}</td><td>{item.row.facility || '—'}</td><td>{item.family || '—'}</td><td>{item.row.order || '—'}</td><td>{item.row.batch || '—'}</td><td>{item.row.material || '—'}</td><td>{item.row.desc || '—'}</td><td>{fmt(item.row.qty)}</td><td>{item.assignedResource || item.pendingManual?.targetResource || '—'}</td><td>{item.approvedManual ? 'מיפוי ידני מאושר' : item.pendingManual ? 'ממתין לאישור' : item.explanation}</td><td>{item.actionable && item.family !== '1542' && !item.pendingManual ? <button className="mapping-assign-btn" onClick={() => openMappingDialog(item)}>שייך</button> : item.pendingManual ? <span className="mapping-pending-chip">ממתין</span> : '—'}</td></tr>)}{!visibleMappingSimulation.length && <tr><td colSpan="12" className="empty">לא נמצאו בעיות שיוך בטווח שנבחר</td></tr>}</tbody></table></div>
      </section>}

      {activeTab === 'mapping-center' && canManageData && <section className="details mapping-center">
        <div className="mapping-simulator-head"><div><h2>מרכז בקרת מיפויים</h2><p className="details-note">אישור, דחייה וביטול של שיוכים. שיוך משפיע על הדשבורד רק לאחר אישור.</p></div></div>
        {mappingMessage && <div className="mapping-message">{mappingMessage}</div>}
        <div className="mapping-summary-grid"><article className="mapping-duplicate"><span>יעדים דורשי מיפוי</span><b>{unmappedPlanningTargets.length}</b></article><article className="mapping-ok"><span>מיפויי יעד מאושרים</span><b>{targetMappings.length}</b></article><article className="mapping-duplicate"><span>מיפויי תפוקה ממתינים</span><b>{manualMappings.filter(item => item.active !== false && item.status === 'pending').length}</b></article><article className="mapping-ok"><span>מיפויי תפוקה מאושרים</span><b>{manualMappings.filter(item => item.active !== false && item.status === 'approved').length}</b></article></div>
        <h3 className="shift-subtitle">יעדים שדורשים מיפוי</h3>
        <div className="table-wrap"><table><thead><tr><th>יעד</th><th>חודש</th><th>סיבה</th><th>מועמדים</th><th>פעולה</th></tr></thead><tbody>{unmappedPlanningTargets.map(row => <tr key={`unmapped-${row.id}`}><td><b>{row.resource}</b></td><td>{planningMonth}</td><td>{row.mappingReason || 'לא נמצאה התאמה חד-משמעית'}</td><td>{row.mappingCandidates?.length ? row.mappingCandidates.map(item => `${item.facility} (${item.rows})`).join(', ') : '—'}</td><td><button className="mapping-assign-btn" onClick={() => openTargetMappingDialog(row)}>הגדר מיפוי</button></td></tr>)}{!unmappedPlanningTargets.length && <tr><td colSpan="5" className="empty">כל היעדים משויכים ✓</td></tr>}</tbody></table></div>
        <h3 className="shift-subtitle">מיפויי יעד מאושרים בענן</h3>
        <div className="table-wrap"><table><thead><tr><th>יעד</th><th>חודש</th><th>משפחת תחנה</th><th>נוצר ע״י</th><th>פעולה</th></tr></thead><tbody>{targetMappings.map(mapping => <tr key={`target-map-${mapping.id || mapping.resource}`}><td><b>{mapping.resource}</b></td><td>{mapping.month || 'כל החודשים'}</td><td>{mapping.family}</td><td>{mapping.createdBy || 'מנהל'}</td><td><button className="danger" onClick={() => removeTargetMapping(mapping)}>בטל</button></td></tr>)}{!targetMappings.length && <tr><td colSpan="5" className="empty">אין מיפויי יעד ידניים</td></tr>}</tbody></table></div>
        <h3 className="shift-subtitle">מיפויי רשומות תפוקה</h3>
        <div className="table-wrap"><table><thead><tr><th>סטטוס</th><th>מתקן</th><th>מק״ט</th><th>תיאור</th><th>יעד</th><th>נוצר ע״י</th><th>תאריך</th><th>פעולות</th></tr></thead><tbody>{manualMappings.map(mapping => <tr key={mapping.id}><td><span className={`mapping-status mapping-status-${mapping.status === 'approved' ? 'matched' : mapping.status === 'pending' ? 'duplicate' : 'unmatched'}`}>{mapping.status === 'approved' ? 'מאושר' : mapping.status === 'pending' ? 'ממתין' : 'נדחה'}</span></td><td>{mapping.family}</td><td>{mapping.material || '—'}</td><td>{mapping.description || '—'}</td><td>{mapping.targetResource}</td><td>{mapping.createdBy || 'מנהל'}</td><td>{mapping.createdAt ? new Date(mapping.createdAt).toLocaleString('he-IL') : '—'}</td><td><div className="mapping-row-actions">{mapping.status === 'pending' && <><button onClick={() => updateMappingStatus(mapping,'approved')}>אשר</button><button className="danger" onClick={() => updateMappingStatus(mapping,'rejected')}>דחה</button></>}{mapping.status === 'approved' && mapping.active !== false && <button className="danger" onClick={() => rollbackMapping(mapping)}>בטל שיוך</button>}</div></td></tr>)}{!manualMappings.length && <tr><td colSpan="8" className="empty">עדיין לא נוצרו מיפויים ידניים</td></tr>}</tbody></table></div>
        <div className="mapping-timeline"><h3>Timeline החלטות</h3>{mappingTimeline.slice(0,30).map(event => <div key={event.id}><time>{new Date(event.at).toLocaleString('he-IL')}</time><b>{event.action}</b><span>{event.material || '—'} → {event.targetResource || '—'}</span><small>{event.user} · {event.note}</small></div>)}</div>
      </section>}

      {targetMappingDialog && <div className="mapping-dialog-backdrop"><div className="mapping-dialog"><button className="mapping-dialog-close" onClick={() => setTargetMappingDialog(null)}><X/></button><h2>הגדרת מיפוי ליעד</h2><div className="mapping-source-card"><b>{targetMappingDialog.resource}</b><span>{targetMappingDialog.mappingReason || 'נדרש מיפוי מנהל'}</span><small>חודש {planningMonth} · יעד {fmt(targetMappingDialog.target)}</small></div><label>בחר משפחת תחנה<select value={targetMappingFamily} onChange={event => setTargetMappingFamily(event.target.value)}><option value="">בחר תחנה...</option>{availableTargetFamilies.map(family => <option key={family} value={family}>{family}</option>)}</select></label>{targetMappingDialog.mappingCandidates?.length > 0 && <div className="mapping-impact"><h3>ראיות שנמצאו</h3>{targetMappingDialog.mappingCandidates.map(item => <div key={item.facility}><span>משפחת תחנה {item.facility}</span><b>{item.rows} רשומות</b></div>)}</div>}<p className="details-note">לאחר שמירה המיפוי נשמר בענן ומשפיע על כל המחשבים. מאזן 42 (1142+999) ומאזן 19 (1119+777) נשארים כללים ייעודיים ואינם משתנים כאן.</p><button className="mapping-save-btn" disabled={!targetMappingFamily} onClick={saveTargetMapping}><Save size={17}/> שמור ואשר מיפוי</button></div></div>}

      {mappingDialog && <div className="mapping-dialog-backdrop"><div className="mapping-dialog"><button className="mapping-dialog-close" onClick={() => setMappingDialog(null)}><X/></button><h2>שיוך חומר ליעד</h2><div className="mapping-source-card"><b>{mappingDialog.row.material || 'ללא מק״ט'}</b><span>{mappingDialog.row.desc || 'ללא תיאור'}</span><small>תחנה {mappingDialog.row.facility} · משפחה {mappingDialog.family} · כמות {fmt(mappingDialog.row.qty)}</small></div><label>בחר יעד<select value={mappingTargetKey} onChange={event => setMappingTargetKey(event.target.value)}><option value="">בחר יעד...</option>{mappingDialog.candidates.map(target => <option key={target.key} value={target.key}>{target.resource} — יעד {fmt(target.target)} — בוצע {fmt(target.actual)}</option>)}</select></label>{(() => { const target=mappingTargets.find(item => item.key===mappingTargetKey); if(!target) return null; const impacted=mappingSimulation.rows.filter(item => item.family===mappingDialog.family && productionMappingKey(item.row)===productionMappingKey(mappingDialog.row)); const qty=impacted.reduce((sum,item)=>sum+(Number(item.row.qty)||0),0); const before=target.actual||0; const after=before+qty; return <div className="mapping-impact"><h3>סימולציית השפעה</h3><div><span>כמות שתשויך</span><b>{fmt(qty)}</b></div><div><span>ביצוע לפני</span><b>{fmt(before)}</b></div><div><span>ביצוע אחרי</span><b>{fmt(after)}</b></div><div><span>עמידה ביעד</span><b>{target.target ? `${Math.round(before/target.target*100)}% → ${Math.round(after/target.target*100)}%` : 'ללא יעד'}</b></div><p>השינוי לא ישפיע על הדשבורד עד לאישור במרכז המיפויים.</p></div> })()}<button className="mapping-save-btn" disabled={!mappingTargetKey} onClick={savePendingMapping}><Save size={17}/> שמור כממתין לאישור</button></div></div>}

      {activeTab === 'shifts' && <section className="details shift-intelligence"><h2>ניתוח משמרות — בוקר, ערב ולילה</h2><p className="details-note">החלפות מוצר מזוהות לפי שינוי מק״ט באותו מתקן ו-Routing group. זמן המעבר הוא פער זמן משוער בין שני דיווחים עוקבים.</p><div className="shift-card-grid">{shiftAnalysis.map(item => <article className={`shift-analysis-card shift-${item.key}`} key={item.key}><div className="shift-card-head"><div><h3>{item.label}</h3><span>{item.hours}</span></div><strong>{fmt(item.total)}</strong></div><div className="shift-metrics"><div><span>קצב ממוצע לשעה</span><b>{fmt(item.avgPerHour)}</b></div><div><span>Orders</span><b>{item.orders}</b></div><div><span>Batch</span><b>{item.batches}</b></div><div><span>מק״טים</span><b>{item.materials}</b></div><div><span>החלפות מוצר</span><b>{item.changeovers.length}</b></div><div><span>חריגות איכות</span><b>{item.deviations}</b></div><div><span>ממוצע מעבר</span><b>{fmt(item.avgChangeover)} דק׳</b></div><div><span>תרומה לתפוקה</span><b>{pctFmt(item.share)}</b></div></div></article>)}</div><h3 className="shift-subtitle">פירוט החלפות מוצר</h3><div className="table-wrap"><table><thead><tr><th>משמרת</th><th>שעה</th><th>משאב יעד</th><th>מתקן / תחנה</th><th>מוצר קודם</th><th>מוצר חדש</th><th>פער דיווח משוער</th></tr></thead><tbody>{shiftAnalysis.flatMap(item => item.changeovers.map((c,i) => <tr key={`${item.key}-${i}`}><td>{item.label}</td><td>{c.at?.toLocaleTimeString('he-IL',{hour:'2-digit',minute:'2-digit'})}</td><td>{c.facility}</td><td>{c.routingGroup || '—'}</td><td>{c.fromMaterial}{c.fromDesc ? ` · ${c.fromDesc}` : ''}</td><td>{c.toMaterial}{c.toDesc ? ` · ${c.toDesc}` : ''}</td><td>{fmt(c.minutes)} דקות</td></tr>))}{!shiftAnalysis.some(item => item.changeovers.length) && <tr><td colSpan="7" className="empty">לא זוהו החלפות מוצר בטווח שנבחר</td></tr>}</tbody></table></div></section>}
      {activeTab === 'quality' && <section className="details"><h2>תוצאות איכות לא תקינות</h2><div className="table-wrap"><table><thead><tr><th>תאריך דגימה</th><th>שעת דגימה</th><th>מתקן</th><th>Inspection Lot</th><th>Order</th><th>Batch</th><th>מק״ט חומר</th><th>סטטוס</th></tr></thead><tbody>{qualityBad.slice(0,300).map((r,i) => <tr key={i}><td>{iso(r.date)}</td><td>{r.date ? new Date(r.date).toLocaleTimeString('he-IL',{hour:'2-digit',minute:'2-digit'}) : '—'}</td><td>{r.facility}</td><td>{r.inspectionLot}</td><td>{r.order}</td><td>{r.batch ? <button type="button" className="batch-link" onClick={() => openBatchCard(r.batch, r.material)}>{r.batch}</button> : '—'}</td><td>{r.material}</td><td><span className="status-bad">{r.status || 'ללא סטטוס'}</span></td></tr>)}{!qualityBad.length && <tr><td colSpan="8" className="empty">לא נמצאו תוצאות איכות לא תקינות</td></tr>}</tbody></table></div></section>}
      {activeTab === 'deviations' && <section className="details"><h2>מנות חריגות פתוחות</h2><p className="details-note">לכל מנה מוצגים מאפייני החריגה ולצדם המאפיינים התקינים שנמשכו מקובץ תוצאות האיכות לפי Batch + מק״ט.</p><div className="table-wrap"><table><thead><tr><th>תאריך חריגה</th><th>תאריך דגימה</th><th>שעת דגימה</th><th>מתקן</th><th>Batch</th><th>מק״ט חומר</th><th>סטטוס</th><th>מאפייני החריגה</th><th>מאפיינים תקינים</th><th>הערות</th></tr></thead><tbody>{openDeviations.slice(0,300).map((r,i) => <tr key={i}><td>{iso(r.date)}</td><td>{iso(r.sampleDate) || '—'}</td><td>{r.sampleDate ? new Date(r.sampleDate).toLocaleTimeString('he-IL',{hour:'2-digit',minute:'2-digit'}) : '—'}</td><td>{r.facility}</td><td>{r.batch ? <button type="button" className="batch-link" onClick={() => openBatchCard(r.batch, r.material)}>{r.batch}</button> : '—'}</td><td>{r.material}</td><td><span className="status-bad">{r.status || 'פתוח'}</span>{r.udCode && <small className="ud-code">{r.udCode}</small>}</td><td className="deviation-characteristics"><div className="characteristics-count bad-count">{r.rejectedCharacteristics.length} חריגים</div>{r.rejectedCharacteristics.length ? r.rejectedCharacteristics.map((c,j) => <div className="deviation-characteristic" key={`${c.characteristic}-${j}`}><strong>{c.characteristic}</strong><span>תוצאה: <b>{c.value || c.qualitative || '—'}{c.unit ? ` ${c.unit}` : ''}</b></span><span>מפרט: {c.lower !== '' || c.upper !== '' ? `${c.lower || '—'} עד ${c.upper || '—'}${c.unit ? ` ${c.unit}` : ''}` : '—'}</span>{c.remarks && c.remarks !== 'N/A' && <small>{c.remarks}</small>}</div>) : <span className="no-characteristics">לא נמצאו פרטי מאפיינים חריגים בקובץ האיכות{r.rejectedCount ? ` (בקובץ החריגות מופיע מספר: ${r.rejectedCount})` : ''}</span>}</td><td className="deviation-characteristics valid-characteristics"><div className="characteristics-count good-count">{r.approvedCharacteristics.length} תקינים</div>{r.approvedCharacteristics.length ? r.approvedCharacteristics.map((c,j) => <div className="deviation-characteristic valid-characteristic" key={`${c.characteristic}-${j}`}><strong>{c.characteristic}</strong><span>תוצאה: <b>{c.value || c.qualitative || '—'}{c.unit ? ` ${c.unit}` : ''}</b></span><span>מפרט: {c.lower !== '' || c.upper !== '' ? `${c.lower || '—'} עד ${c.upper || '—'}${c.unit ? ` ${c.unit}` : ''}` : '—'}</span>{c.remarks && c.remarks !== 'N/A' && <small>{c.remarks}</small>}</div>) : <span className="no-characteristics">לא נמצאו מאפיינים תקינים למנה בקובץ האיכות</span>}</td><td>{r.remarks || '—'}</td></tr>)}{!openDeviations.length && <tr><td colSpan="10" className="empty">לא נמצאו מנות חריגות פתוחות</td></tr>}</tbody></table></div></section>}
    </main>
    {quantityVarianceRow && <div className="quantity-variance-backdrop" onMouseDown={e=>{if(e.target===e.currentTarget)setQuantityVarianceRow(null)}}><section className="quantity-variance-modal" role="dialog" aria-modal="true"><div className="quantity-variance-head"><div><small>בקרת כמות</small><h3>{quantityVarianceRow.material || 'רשומת תפוקה'}</h3><p>{quantityVarianceRow.desc || ''}</p></div><button type="button" onClick={()=>setQuantityVarianceRow(null)} aria-label="סגירה"><X size={20}/></button></div><div className="quantity-variance-grid"><div><span>כמות מתוכננת</span><b>{num(quantityVarianceRow.plannedQty)>0 ? fmt(quantityVarianceRow.plannedQty) : 'לא זמין'}</b></div><div><span>כמות בפועל</span><b>{fmt(quantityVarianceRow.qty)}</b></div><div className={num(quantityVarianceRow.plannedQty)>0 && Math.abs(num(quantityVarianceRow.qty)-num(quantityVarianceRow.plannedQty))>0.0001 ? 'variance-alert' : 'variance-ok'}><span>פער</span><b>{num(quantityVarianceRow.plannedQty)>0 ? fmt(num(quantityVarianceRow.qty)-num(quantityVarianceRow.plannedQty)) : '—'}</b></div><div><span>סטייה %</span><b>{num(quantityVarianceRow.plannedQty)>0 ? `${((num(quantityVarianceRow.qty)-num(quantityVarianceRow.plannedQty))/num(quantityVarianceRow.plannedQty)*100).toFixed(1)}%` : '—'}</b></div></div><div className="quantity-variance-meta"><span>Order <b>{quantityVarianceRow.order || '—'}</b></span><span>Batch <b>{quantityVarianceRow.batch || '—'}</b></span><span>מתקן <b>{quantityVarianceRow.facility || '—'}</b></span><span>Routing <b>{quantityVarianceRow.routingGroup || '—'}</b></span></div></section></div>}
    {selectedResource && <ResourceDetailModal resource={selectedResource} onClose={() => setSelectedResource(null)} onOpenBatch={(batch, material='') => { setSelectedResource(null); openBatchCard(batch, material) }}/>}
    {selectedBatchData && <BatchControlCard data={selectedBatchData} onClose={() => { setSelectedBatch(''); setSelectedBatchMaterial('') }}/>}
  </div>
}

function ResourceDetailModal({ resource, onClose, onOpenBatch }) {
  useEffect(() => {
    const previous = document.body.style.overflow
    document.body.style.overflow = 'hidden'
    const close = event => { if (event.key === 'Escape') onClose() }
    window.addEventListener('keydown', close)
    return () => { document.body.style.overflow = previous; window.removeEventListener('keydown', close) }
  }, [onClose])
  const rows = resource.productionRows || []
  const batches = [...new Set(rows.map(row => row.batch).filter(Boolean))]
  const orders = [...new Set(rows.map(row => row.order).filter(Boolean))]
  const materials = [...rows.reduce((map,row) => {
    const key = normalize(row.material) || 'ללא מק״ט'
    const current = map.get(key) || { material:key, description:row.desc || '', qty:0, batches:new Set(), orders:new Set(), rows:0 }
    current.qty += Number(row.qty) || 0
    current.rows += 1
    if (row.batch) current.batches.add(row.batch)
    if (row.order) current.orders.add(row.order)
    if (!current.description && row.desc) current.description = row.desc
    map.set(key,current)
    return map
  }, new Map()).values()].sort((a,b)=>b.qty-a.qty)
  const calculatedActual = rows.reduce((sum,row)=>sum+(Number(row.qty)||0),0)
  const progress = resource.target ? Math.min(100, resource.actual / resource.target * 100) : 0
  const printRows = () => {
    const popup = window.open('', '_blank', 'width=1400,height=900')
    if (!popup) return
    const esc = value => String(value ?? '').replace(/[&<>\"]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;'}[ch]))
    const body = rows.slice().sort((a,b)=>String(b.productionDay||'').localeCompare(String(a.productionDay||''))).map((row,index)=>`<tr><td>${index+1}</td><td>${esc(row.productionDay || iso(row.date) || '—')}</td><td>${esc(row.facility||'—')}</td><td>${esc(row.routingGroup||'—')}</td><td>${esc(row.orderType||'—')}</td><td>${esc(row.order||'—')}</td><td>${esc(row.batch||'—')}</td><td>${esc(row.material||'—')}</td><td>${esc(row.desc||'—')}</td><td>${esc(fmt(row.qty))}</td></tr>`).join('')
    popup.document.write(`<!doctype html><html dir="rtl"><head><meta charset="utf-8"><title>${esc(resource.resource)}</title><style>@page{size:A4 landscape;margin:8mm}body{font-family:Arial,sans-serif;color:#173b57}.head{display:flex;justify-content:space-between;border-bottom:4px solid #159b83;padding-bottom:12px;margin-bottom:14px}.head h1{margin:0}.meta{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:12px}.meta span{background:#eaf7f3;padding:7px 10px;border-radius:10px;font-weight:700}table{width:100%;border-collapse:collapse;font-size:10px}th{background:#173b57;color:#fff;padding:7px}td{padding:6px;border-bottom:1px solid #dce7ec;text-align:center}tr:nth-child(even) td{background:#f7fafc}</style></head><body><div class="head"><div><h1>${esc(resource.resource)}</h1><p>${esc([resource.station&&`Station ${resource.station}`,resource.description,resource.lineName].filter(Boolean).join(' · '))}</p></div><b>IML CONTROL</b></div><div class="meta"><span>יעד ${esc(fmt(resource.target))}</span><span>בוצע ${esc(fmt(resource.actual))}</span><span>תחזית ${esc(fmt(resource.forecast))}</span><span>מק״טים ${materials.length}</span><span>מנות ${batches.length}</span></div><table><thead><tr><th>#</th><th>תאריך</th><th>תחנה</th><th>Routing</th><th>Order Type</th><th>Order</th><th>Batch</th><th>מק״ט</th><th>תיאור</th><th>כמות</th></tr></thead><tbody>${body}</tbody></table><script>window.onload=()=>setTimeout(()=>window.print(),200)<\/script></body></html>`)
    popup.document.close()
  }

  const exportRows = () => {
    const wb = XLSX.utils.book_new()
    appendAutoFitJsonSheet(wb, materials.map(item => ({
      Material:item.material, Description:item.description, Quantity:item.qty, Batches:item.batches.size, Orders:item.orders.size, Rows:item.rows,
    })), 'Materials')
    appendAutoFitJsonSheet(wb, rows.map(row => ({
      Date:row.productionDay || iso(row.date), Facility:row.facility, ProdLine:row.prodLine || '', Tool:row.prodLineTool || '', RoutingGroup:row.routingGroup, OrderType:row.orderType,
      Order:row.order, Batch:row.batch, Material:row.material, Description:row.desc, Quantity:row.qty,
    })), 'Calculation Rows')
    XLSX.writeFile(wb, `IML_${String(resource.resource || 'resource').replace(/[^a-zA-Z0-9_-]+/g,'_')}_${planningSafeMonth(resource)}.xlsx`)
  }
  return <div className="resource-modal-backdrop" onMouseDown={event => { if (event.target === event.currentTarget) onClose() }}>
    <section className="resource-detail-modal resource-detail-modal-audit" role="dialog" aria-modal="true" aria-label={`פירוט ${resource.resource}`}>
      <header><div><span>RESOURCE CONTROL CENTER</span><h2>{resource.resource}</h2><p>{[resource.station && `Station ${resource.station}`, resource.description, resource.lineName].filter(Boolean).join(' · ')}</p></div><div className="resource-head-actions"><button className="resource-export-btn" onClick={printRows}><Printer size={17}/> הדפסת סימניה</button><button className="resource-export-btn" onClick={exportRows}><FileSpreadsheet size={17}/> ייצוא חישוב</button><button onClick={onClose} aria-label="סגירה"><X/></button></div></header>
      <div className={`resource-status-banner ${resource.state}`}><strong>{resource.label}</strong><span>תחזית {fmt(resource.forecast)} מול יעד {fmt(resource.target)}</span></div>
      <div className="resource-detail-kpis">
        <BatchMetric label="יעד חודשי" value={fmt(resource.target)}/><BatchMetric label="בוצע" value={fmt(resource.actual)}/><BatchMetric label="סכום רשומות" value={fmt(calculatedActual)}/><BatchMetric label="מק״טים" value={materials.length}/><BatchMetric label="מנות" value={batches.length}/><BatchMetric label="Orders" value={orders.length}/><BatchMetric label="תחזית" value={fmt(resource.forecast)}/><BatchMetric label="קצב נדרש" value={fmt(resource.requiredDaily)}/>
      </div>
      <div className="resource-detail-progress"><div><span>התקדמות חודשית</span><b>{Math.round(progress)}%</b></div><i><em style={{width:`${progress}%`}}/></i></div>
      <section className="resource-audit-section">
        <div className="resource-audit-title"><div><h3>מק״טים שנכנסו לחישוב</h3><p>הסכום בטבלה הוא בדיוק מקור נתון ה״בוצע״ בכרטיס.</p></div><span>{materials.length} מק״טים · {fmt(calculatedActual)}</span></div>
        <div className="table-wrap resource-audit-table"><table><thead><tr><th>מק״ט</th><th>תיאור</th><th>כמות בחישוב</th><th>מנות</th><th>Orders</th><th>רשומות</th></tr></thead><tbody>{materials.map(item=><tr key={item.material}><td><b>{item.material}</b></td><td>{item.description || '—'}</td><td><b>{fmt(item.qty)}</b></td><td>{item.batches.size}</td><td>{item.orders.size}</td><td>{item.rows}</td></tr>)}{!materials.length&&<tr><td colSpan="6" className="empty">לא נמצאו מק״טים בחישוב.</td></tr>}</tbody></table></div>
      </section>
      <section className="resource-audit-section">
        <div className="resource-audit-title"><div><h3>פירוט מלא — מנות ורשומות תפוקה</h3><p>כל שורה שהמנוע כלל בחישוב המשאב.</p></div><span>{rows.length} רשומות</span></div>
        <div className="table-wrap resource-audit-table resource-row-audit"><table><thead><tr><th>תאריך</th><th>תחנה</th><th>Routing</th><th>Order Type</th><th>Order</th><th>Batch</th><th>מק״ט</th><th>תיאור</th><th>כמות</th></tr></thead><tbody>{rows.slice().sort((a,b)=>String(b.productionDay||'').localeCompare(String(a.productionDay||''))).map((row,index)=><tr key={`${row.order}-${row.batch}-${row.material}-${index}`}><td>{row.productionDay || iso(row.date) || '—'}</td><td>{row.facility || '—'}</td><td>{row.routingGroup || '—'}</td><td>{row.orderType || '—'}</td><td>{row.order || '—'}</td><td>{row.batch ? <button type="button" className="batch-link" onClick={()=>onOpenBatch(row.batch,row.material)}>{row.batch}</button> : '—'}</td><td>{row.material || '—'}</td><td>{row.desc || '—'}</td><td><b>{fmt(row.qty)}</b></td></tr>)}{!rows.length&&<tr><td colSpan="9" className="empty">לא נמצאו רשומות בחישוב.</td></tr>}</tbody></table></div>
      </section>
    </section>
  </div>
}

const planningSafeMonth = resource => String(resource?.month || new Date().toISOString().slice(0,7))

const BATCH_MODAL_WIDE_STYLES = `
.batch-modal-backdrop{padding:10px!important;align-items:stretch!important;justify-content:stretch!important}
.batch-control-card.batch-control-card-wide{width:calc(100vw - 20px)!important;max-width:none!important;height:calc(100vh - 20px)!important;max-height:none!important;margin:0!important;border-radius:22px!important;display:flex!important;flex-direction:column!important;overflow:hidden!important}
.batch-control-card-wide .batch-card-head{flex:0 0 auto}
.batch-control-card-wide .batch-summary-grid{grid-template-columns:repeat(8,minmax(130px,1fr))!important;gap:10px!important;padding-inline:18px!important}
.batch-control-card-wide .batch-kpi-row{grid-template-columns:repeat(4,minmax(170px,1fr))!important;padding-inline:18px!important}
.batch-control-card-wide .batch-card-body.batch-card-body-wide{display:block!important;flex:1 1 auto!important;min-height:0!important;padding:0 18px 18px!important;overflow:hidden!important}
.batch-workspace-tabs{display:flex;gap:10px;align-items:center;flex-wrap:wrap;padding:10px 0 12px;border-bottom:1px solid #dce7ec;margin-bottom:12px}
.batch-workspace-tabs button{border:1px solid #cbdbe2;background:#fff;color:#24445a;border-radius:12px;padding:10px 18px;font-weight:800;cursor:pointer;display:flex;align-items:center;gap:8px}
.batch-workspace-tabs button.active{background:#0d766f;color:#fff;border-color:#0d766f;box-shadow:0 8px 18px rgba(13,118,111,.18)}
.batch-workspace-tabs button.bad.active{background:#c93636;border-color:#c93636}
.batch-workspace{height:calc(100% - 58px);min-height:0;overflow:auto}
.batch-workspace .batch-panel{height:auto;min-height:100%;margin:0!important}
.batch-workspace .table-wrap{max-height:none!important;height:auto!important;overflow:auto!important}
.batch-workspace table{width:100%!important;min-width:1380px!important;table-layout:auto!important}
.batch-workspace th,.batch-workspace td{padding:12px 14px!important;white-space:nowrap;vertical-align:top}
.batch-workspace th:nth-child(2),.batch-workspace td:nth-child(2){min-width:300px;white-space:normal}
.batch-workspace th:last-child,.batch-workspace td:last-child{min-width:260px;white-space:normal}
.batch-deviation-table th:nth-child(7),.batch-deviation-table td:nth-child(7){min-width:360px;white-space:normal}
.batch-deviation-table th:nth-child(8),.batch-deviation-table td:nth-child(8){min-width:320px;white-space:normal}
.batch-deviation-table .deviation-characteristic{margin-bottom:8px;padding:8px 10px;border-radius:10px;background:#fff6f6;border:1px solid #ffd7d7}
.batch-deviation-table .deviation-characteristic strong{display:block;margin-bottom:4px}
.batch-support-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:14px}
.batch-support-grid .batch-panel{min-height:0!important}
@media(max-width:1200px){.batch-control-card-wide .batch-summary-grid{grid-template-columns:repeat(4,minmax(130px,1fr))!important}.batch-support-grid{grid-template-columns:1fr}.batch-workspace table{min-width:1100px!important}}
`;

function BatchControlCard({ data, onClose }) {
  const [qualityFilter, setQualityFilter] = useState('all')
  const [activeBatchTab, setActiveBatchTab] = useState('quality')
  useEffect(() => {
    const previous = document.body.style.overflow
    document.body.style.overflow = 'hidden'
    const onKey = (event) => { if (event.key === 'Escape') onClose() }
    window.addEventListener('keydown', onKey)
    return () => { document.body.style.overflow = previous; window.removeEventListener('keydown', onKey) }
  }, [onClose])
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
    appendAutoFitJsonSheet(wb, productionRows.map(r=>({Date:iso(r.date),Facility:r.facility,RoutingGroup:r.routingGroup,Order:r.order,Batch:r.batch,Material:r.material,Description:r.desc,Quantity:r.qty})), 'Production')
    appendAutoFitJsonSheet(wb, allQuality.map(r=>({Date:iso(r.date),InspectionLot:r.inspectionLot,Material:r.material||materials.join(', '),Characteristic:r.characteristic,Result:r.value||r.qualitative,Lower:r.lower,Upper:r.upper,Unit:r.unit,Status:r.rejected?'חריג':'תקין',Remarks:r.remarks})), 'Quality')
    appendAutoFitJsonSheet(wb, deviationRows.map(r=>({Date:iso(r.date),Facility:r.facility,Material:r.material||materials.join(', '),Status:r.status,UDCode:r.udCode,Remarks:r.remarks})), 'חריגות איכות')
    XLSX.writeFile(wb, `Batch_${data.batch}.xlsx`)
  }
  return <div className="batch-modal-backdrop" onMouseDown={e=>{if(e.target===e.currentTarget) onClose()}}>
    <style>{BATCH_MODAL_WIDE_STYLES}</style>
    <section className="batch-control-card batch-control-card-wide" role="dialog" aria-modal="true" aria-label={`כרטיס מנה ${data.batch}`}>
      <header className="batch-card-head"><div><span className="batch-eyebrow">BATCH CONTROL CENTER</span><h2>כרטיס מנה — {data.batch}</h2><p>{descriptions.join(' · ') || materials.join(', ') || 'ללא תיאור חומר'}</p></div><div className="batch-head-actions"><button type="button" className="batch-export" onClick={exportBatch}><Download size={17}/> ייצוא מנה</button><button type="button" className="batch-close" onClick={onClose} aria-label="סגירה"><X/></button></div></header>
      <div className="batch-summary-grid">
        <BatchMetric label="Batch" value={data.batch}/><BatchMetric label="Order" value={orders.join(', ') || '—'}/><BatchMetric label="מק״ט חומר" value={materials.join(', ') || '—'}/><BatchMetric label="מתקן" value={facilities.join(', ') || '—'}/><BatchMetric label="Routing group" value={routingGroups.join(', ') || '—'}/><BatchMetric label="כמות ארוזה" value={fmt(totalQty)}/><BatchMetric label="Inspection Lot" value={inspectionLots.join(', ') || '—'}/><BatchMetric label="QA" value={qaApprovals.join(', ') || 'טרם התקבלה החלטה'}/>
      </div>
      <div className="batch-kpi-row"><div className="batch-kpi good"><span>מאפיינים תקינים</span><b>{goodCount}</b></div><div className="batch-kpi bad"><span>מאפיינים חריגים</span><b>{badCount}</b></div><div className="batch-kpi"><span>אחוז הצלחה</span><b>{qualityPct}%</b></div><div className="batch-kpi"><span>חריגות פתוחות</span><b>{deviationRows.length}</b></div></div>
      <div className="batch-card-body batch-card-body-wide">
        <nav className="batch-workspace-tabs" aria-label="לשוניות כרטיס מנה">
          <button type="button" className={activeBatchTab==='quality'?'active':''} onClick={()=>setActiveBatchTab('quality')}><FlaskConical size={17}/> תוצאות איכות ({allQuality.length})</button>
          <button type="button" className={`bad ${activeBatchTab==='deviations'?'active':''}`} onClick={()=>setActiveBatchTab('deviations')}><AlertTriangle size={17}/> חריגים ({deviationRows.length})</button>
          <button type="button" className={activeBatchTab==='production'?'active':''} onClick={()=>setActiveBatchTab('production')}><Factory size={17}/> ייצור ואריזה ({productionRows.length})</button>
          <button type="button" className={activeBatchTab==='timeline'?'active':''} onClick={()=>setActiveBatchTab('timeline')}><Clock3 size={17}/> Timeline והערות</button>
        </nav>
        <div className="batch-workspace">
          {activeBatchTab==='quality' && <section className="batch-panel quality-panel"><div className="batch-panel-title"><div><h3>תוצאות איכות</h3><p>תצוגה רחבה של כל המאפיינים שנמצאו למנה בקובץ האיכות</p></div><div className="quality-filter"><button className={qualityFilter==='all'?'active':''} onClick={()=>setQualityFilter('all')}>הכול {allQuality.length}</button><button className={qualityFilter==='good'?'active good':''} onClick={()=>setQualityFilter('good')}>תקינים {goodCount}</button><button className={qualityFilter==='bad'?'active bad':''} onClick={()=>setQualityFilter('bad')}>חריגים {badCount}</button></div></div><div className="table-wrap batch-quality-table"><table><thead><tr><th>מק״ט חומר</th><th>מאפיין</th><th>תוצאה</th><th>גבול תחתון</th><th>גבול עליון</th><th>יחידה</th><th>סטטוס</th><th>הערה</th></tr></thead><tbody>{shownQuality.map(r=><tr key={r._key} className={r.rejected?'quality-row-bad':''}><td>{r.material || materials.join(', ') || '—'}</td><td><b>{r.characteristic}</b></td><td>{r.value||r.qualitative||'—'}</td><td>{r.lower||'—'}</td><td>{r.upper||'—'}</td><td>{r.unit||'—'}</td><td><span className={`quality-status ${r.rejected?'bad':'good'}`}>{r.rejected?'חריג':'תקין'}</span></td><td>{r.remarks||'—'}</td></tr>)}{!shownQuality.length&&<tr><td colSpan="8" className="empty">לא נמצאו תוצאות במסנן שנבחר</td></tr>}</tbody></table></div></section>}

          {activeBatchTab==='deviations' && <section className="batch-panel deviation-panel"><div className="batch-panel-title"><div><h3>חריגים</h3><p>תצוגה רחבה של כל החריגות, המאפיינים החריגים וההערות המקושרות למנה</p></div><span className="production-record-count">{deviationRows.length} חריגים</span></div><div className="table-wrap batch-deviation-table"><table><thead><tr><th>תאריך חריגה</th><th>תאריך דגימה</th><th>שעת דגימה</th><th>מתקן</th><th>סטטוס</th><th>UD Code</th><th>מאפייני חריגה</th><th>הערות</th></tr></thead><tbody>{deviationRows.map((r,i)=><tr key={`${r.batch}-${r.inspectionLot||i}-${i}`} className="quality-row-bad"><td>{iso(r.date)||'—'}</td><td>{iso(r.sampleDate)||'—'}</td><td>{r.sampleDate?new Date(r.sampleDate).toLocaleTimeString('he-IL',{hour:'2-digit',minute:'2-digit'}):'—'}</td><td>{r.facility||facilities.join(', ')||'—'}</td><td><span className="quality-status bad">{r.status||'פתוח'}</span></td><td>{r.udCode||'—'}</td><td>{r.rejectedCharacteristics?.length ? r.rejectedCharacteristics.map((c,j)=><div className="deviation-characteristic" key={`${c.characteristic}-${j}`}><strong>{c.characteristic}</strong><span>תוצאה: <b>{c.value||c.qualitative||'—'}{c.unit?` ${c.unit}`:''}</b></span><span>מפרט: {c.lower!==''||c.upper!==''?`${c.lower||'—'} עד ${c.upper||'—'}${c.unit?` ${c.unit}`:''}`:'—'}</span>{c.remarks&&c.remarks!=='N/A'&&<small>{c.remarks}</small>}</div>) : <span className="no-characteristics">לא נמצאו פרטי מאפיינים חריגים בקובץ האיכות{r.rejectedCount?` (בקובץ החריגות מופיע מספר: ${r.rejectedCount})`:''}</span>}</td><td>{r.remarks||'—'}</td></tr>)}{!deviationRows.length&&<tr><td colSpan="8" className="empty">לא נמצאו חריגות למנה</td></tr>}</tbody></table></div></section>}

          {activeBatchTab==='production' && <section className="batch-panel production-panel"><div className="batch-panel-title"><div><h3>נתוני ייצור ואריזה</h3><p>כל דיווחי התפוקה המקושרים למנה</p></div><span className="production-record-count">{productionRows.length} דיווחים</span></div><div className="table-wrap batch-production-table"><table><thead><tr><th>תאריך</th><th>שעה</th><th>משאב יעד</th><th>מתקן / תחנה</th><th>Order</th><th>מק״ט</th><th>תיאור</th><th>כמות</th><th>משמרת</th></tr></thead><tbody>{productionRows.slice().sort((a,b)=>new Date(b.date||0)-new Date(a.date||0)).map((r,index)=><tr key={`${r.order}-${r.date?.getTime?.()||index}-${index}`}><td>{iso(r.date)||'—'}</td><td>{r.date?new Date(r.date).toLocaleTimeString('he-IL',{hour:'2-digit',minute:'2-digit'}):'—'}</td><td>{r.facility||'—'}</td><td>{r.routingGroup||'—'}</td><td>{r.order||'—'}</td><td>{r.material||'—'}</td><td>{r.desc||'—'}</td><td><b>{fmt(r.qty)}</b></td><td>{shiftInfo(r.date).label}</td></tr>)}{!productionRows.length&&<tr><td colSpan="9" className="empty">לא נמצאו דיווחי ייצור או אריזה למנה</td></tr>}</tbody></table></div></section>}

          {activeBatchTab==='timeline' && <div className="batch-support-grid"><section className="batch-panel"><div className="batch-panel-title"><div><h3>Timeline</h3><p>מצב התקדמות המנה</p></div></div><div className="batch-timeline">{steps.map((step,i)=><div className={`timeline-step ${step.done?'done':''}`} key={step.label}><i>{step.done?<CheckCircle2 size={18}/>:<Clock3 size={18}/>}</i><div><b>{step.label}</b><span>{step.date?`${iso(step.date)} ${new Date(step.date).toLocaleTimeString('he-IL',{hour:'2-digit',minute:'2-digit'})}`:'טרם הושלם'}</span></div>{i<steps.length-1&&<em/>}</div>)}</div></section><section className="batch-panel"><div className="batch-panel-title"><div><h3>חריגות והערות</h3><p>{deviationRows.length} רשומות מקושרות למנה</p></div></div><div className="batch-deviations">{deviationRows.map((r,i)=><article key={i}><div><span className="quality-status bad">{r.status||'חריגה'}</span>{r.udCode&&<b>{r.udCode}</b>}</div><p><b>מק״ט חומר: {r.material || materials.join(', ') || '—'}</b></p><p>{r.remarks||'לא הוזנה הערה'}</p><small>{iso(r.date)||'ללא תאריך'}</small></article>)}{!deviationRows.length&&<div className="batch-empty-good"><CheckCircle2/> לא נמצאו חריגות למנה</div>}</div></section></div>}
        </div>
      </div>
    </section>
  </div>
}
function BatchMetric({label,value}) { return <div className="batch-metric"><span>{label}</span><b>{value}</b></div> }

function DataSource({ title, icon, meta, count, rows = [], showYearBreakdown = false, acceptLabel, busy, onFiles, canManage }) {
  const [breakdownOpen, setBreakdownOpen] = useState(false)
  const [selectedYear, setSelectedYear] = useState(null)
  const loaded = Boolean(meta || count)
  const loadedAt = meta?.loadedAt ? new Date(meta.loadedAt).toLocaleString('he-IL') : 'טרם נטען'
  const yearBreakdown = useMemo(() => {
    if (!showYearBreakdown) return []
    const counts = new Map()
    let noDate = 0
    rows.forEach(row => {
      const raw = row?.date
      const d = raw instanceof Date ? raw : (raw ? new Date(raw) : null)
      const year = d && Number.isFinite(d.getTime()) ? d.getFullYear() : null
      if (year && year >= 2000 && year <= 2100) counts.set(year, (counts.get(year) || 0) + 1)
      else noDate += 1
    })
    const result = [...counts.entries()].sort((a,b) => b[0]-a[0]).map(([year,value]) => ({ year:String(year), value }))
    if (noDate) result.push({ year:'ללא תאריך', value:noDate })
    return result
  }, [rows, showYearBreakdown])
  const monthBreakdown = useMemo(() => {
    if (!showYearBreakdown || !selectedYear) return []
    const months = Array.from({ length: 12 }, (_, i) => ({ month: i + 1, value: 0 }))
    rows.forEach(row => {
      const raw = row?.date
      const d = raw instanceof Date ? raw : (raw ? new Date(raw) : null)
      if (!d || !Number.isFinite(d.getTime()) || d.getFullYear() !== Number(selectedYear)) return
      months[d.getMonth()].value += 1
    })
    return months
  }, [rows, showYearBreakdown, selectedYear])
  const selectedYearTotal = monthBreakdown.reduce((sum, item) => sum + item.value, 0)
  const monthNames = ['ינואר','פברואר','מרץ','אפריל','מאי','יוני','יולי','אוגוסט','ספטמבר','אוקטובר','נובמבר','דצמבר']
  return <>
    <article className={`data-source ${loaded ? 'ready' : ''}`}>
      <div className="data-source-head"><div className="data-source-icon">{icon}</div><div><h3>{title}</h3><span>{loaded ? 'תקין וזמין' : 'ממתין לקובץ'}</span></div></div>
      <div className="data-source-count"><b>{fmt(count)}</b><span>רשומות ייחודיות במאגר</span></div>{meta?.lastFileRows ? <div className="data-source-last-file"><b>{fmt(meta.lastFileRows)}</b><span>רשומות בקובץ האחרון</span>{meta?.lastFileUniqueRows != null && <small>ייחודיות בקובץ: {fmt(meta.lastFileUniqueRows)}</small>}</div> : null}
      {showYearBreakdown && loaded && <button type="button" className="source-breakdown-btn" onClick={()=>setBreakdownOpen(true)}>פירוט מאגר לפי שנה</button>}
      <div className="data-source-meta"><small title={meta?.fileName || ''}>{meta?.fileName || 'לא נבחר קובץ'}</small><small>{loadedAt}</small>{meta?.source === 'cloud' && <small className="cloud-source-label">מקור: Supabase{meta?.loadedBy ? ` · ${meta.loadedBy}` : ''}</small>}{meta?.facilities ? <small>{meta.facilities} מתקנים זוהו במדגם</small> : null}</div>
      {canManage ? <label className={`source-upload ${busy ? 'disabled' : ''}`}><RefreshCw size={16}/>{acceptLabel}<input type="file" accept=".xlsx,.xls" disabled={busy} onChange={e => { const files=[...e.target.files]; e.target.value=''; onFiles(files) }}/></label> : <div className="viewer-lock"><ShieldCheck size={16}/> צפייה בלבד</div>}
    </article>
    {breakdownOpen && <div className="year-breakdown-backdrop" onMouseDown={e=>{if(e.target===e.currentTarget)setBreakdownOpen(false)}}>
      <section className="year-breakdown-modal" role="dialog" aria-modal="true" aria-label="פירוט מאגר איכות לפי שנה">
        <header><div><small>QUALITY DATABASE</small><h3>פירוט מאגר תוצאות איכות</h3><p>פילוח הרשומות הייחודיות שנמצאות כרגע באפליקציה לפי שנת תאריך הבדיקה.</p></div><button type="button" onClick={()=>setBreakdownOpen(false)} aria-label="סגירה"><X size={20}/></button></header>
        <div className="year-breakdown-total"><span>סה״כ רשומות ייחודיות</span><b>{fmt(count)}</b></div>
        {!selectedYear ? <>
          <div className="year-breakdown-list">{yearBreakdown.map(item => <button type="button" className={item.year === 'ללא תאריך' ? 'year-breakdown-row disabled' : 'year-breakdown-row'} key={item.year} disabled={item.year === 'ללא תאריך'} onClick={()=>item.year !== 'ללא תאריך' && setSelectedYear(item.year)}><span>{item.year}</span><b>{fmt(item.value)}</b><em>{count ? `${(item.value/count*100).toFixed(1)}%` : '0%'}</em></button>)}</div>
          <div className="year-breakdown-hint">לחץ על שנה כדי לראות פירוט חודשי</div>
        </> : <div className="month-audit">
          <div className="month-audit-head"><button type="button" onClick={()=>setSelectedYear(null)}>← חזרה לשנים</button><div><small>MONTHLY AUDIT</small><h4>פירוט חודשי — {selectedYear}</h4></div></div>
          <div className="month-audit-total"><span>סה״כ בשנת {selectedYear}</span><b>{fmt(selectedYearTotal)}</b></div>
          <div className="month-audit-list">{monthBreakdown.map(item => <div key={item.month}><span>{monthNames[item.month-1]}</span><b>{fmt(item.value)}</b><em>{selectedYearTotal ? `${(item.value/selectedYearTotal*100).toFixed(1)}%` : '0%'}</em></div>)}</div>
        </div>}
        <footer><small>הפילוח אינו מוחק או משנה נתונים ב-Supabase — הוא כלי בדיקה בלבד.</small></footer>
      </section>
    </div>}
  </>
}

function Summary({ title, value, sub, warn }) { return <div className={`summary ${warn ? 'warn' : ''}`}><span>{title}</span><b>{value}</b><small>{sub}</small></div> }
function Executive({ icon, title, value, sub, good, warn, bad, onClick }) { return <button type="button" className={`executive ${good?'good':''} ${warn?'warn':''} ${bad?'bad':''} ${onClick?'clickable':''}`} onClick={onClick}><div className="executive-icon">{icon}</div><div><span>{title}</span><b>{value}</b><small>{sub}</small></div></button> }
function StatusBadge({ state, label }) { return <span className={`status-pill ${state}`}>{label}</span> }

function Facility42BalanceOverviewCard({ balance, onClick }) {
  const diffClass = balance.balance > 0 ? 'positive' : balance.balance < 0 ? 'negative' : 'neutral'
  return <button type="button" className={`tower-facility-card facility42-overview-balance ${diffClass}`} onClick={onClick}>
    <div className="tower-facility-head"><div><i></i><strong>מאזן מתקן 42</strong></div><span>ללא יעד</span></div>
    <div className="facility42-overview-main"><div><span>באלק 1142 + 999</span><b>{fmt(balance.bulk)}</b></div><div><span>סה״כ ארוז ZFIN</span><b>{fmt(balance.packed)}</b></div></div>
    <div className="facility42-overview-lines"><span>1L <b>{fmt(balance.byLine['1L'])}</b></span><span>5L <b>{fmt(balance.byLine['5L'])}</b></span><span>10/20L <b>{fmt(balance.byLine['10/20L'])}</b></span></div>
    <dl><div><dt>יתרת באלק</dt><dd className={balance.balance >= 0 ? 'positive' : 'negative'}>{balance.balance >= 0 ? '+' : ''}{fmt(balance.balance)}</dd></div><div><dt>ניצול באלק</dt><dd>{balance.bulk ? pctFmt(balance.utilization) : '—'}</dd></div></dl>
    <span className="tower-enter">לפירוט המאזן <ArrowLeft size={16}/></span>
  </button>
}

function Facility42BalanceCard({ balance }) {
  const diffClass = balance.balance > 0 ? 'positive' : balance.balance < 0 ? 'negative' : 'neutral'
  return <article className={`forecast-card facility42-balance-card ${diffClass}`}>
    <div className="forecast-head"><div><small>מאזן תשומות / תפוקות</small><h3>מאזן מתקן 42</h3><div className="forecast-resource"><b>1142 + 999</b><span>מול אריזה 1L / 5L / 10–20L</span></div></div><span className="status-badge no-target">ללא יעד</span></div>
    <div className="balance-card-main"><div><span>באלק שיוצר</span><b>{fmt(balance.bulk)}</b></div><div><span>סה״כ נארז</span><b>{fmt(balance.packed)}</b></div></div>
    <div className="balance-card-lines"><span>1L<strong>{fmt(balance.byLine['1L'])}</strong></span><span>5L<strong>{fmt(balance.byLine['5L'])}</strong></span><span>10/20L<strong>{fmt(balance.byLine['10/20L'])}</strong></span></div>
    <div className="balance-card-footer"><div><span>יתרה</span><b>{balance.balance > 0 ? '+' : ''}{fmt(balance.balance)}</b></div><div><span>ניצול באלק</span><b>{pctFmt(balance.utilization)}</b></div></div>
    <small className="balance-card-note">המאזן מחושב לפי טווח התאריכים שנבחר. פער יומי יכול לנבוע מאריזה ביום שונה מיום ייצור הבאלק.</small>
  </article>
}
function Facility19BalanceOverviewCard({ balance, onClick }) {
  const diffClass = balance.balance > 0 ? 'positive' : balance.balance < 0 ? 'negative' : 'neutral'
  return <button type="button" className={`tower-facility-card facility42-overview-balance facility19-overview-balance ${diffClass}`} onClick={onClick}>
    <div className="tower-facility-head"><div><i></i><strong>מאזן מתקן 19</strong></div><span>ללא יעד</span></div>
    <div className="facility42-overview-main"><div><span>באלק 1119 + 777</span><b>{fmt(balance.bulk)}</b></div><div><span>סה״כ תפוקה 1519</span><b>{fmt(balance.packed)}</b></div></div>
    <div className="facility42-overview-lines facility19-overview-lines"><span>WG <b>{fmt(balance.byType['WG'])}</b></span><span>Small Packs <b>{fmt(balance.byType['SMALL PACKS'])}</b></span></div>
    <dl><div><dt>יתרת באלק</dt><dd className={balance.balance >= 0 ? 'positive' : 'negative'}>{balance.balance >= 0 ? '+' : ''}{fmt(balance.balance)}</dd></div><div><dt>ניצול באלק</dt><dd>{balance.bulk ? pctFmt(balance.utilization) : '—'}</dd></div></dl>
    <span className="tower-enter">לפירוט המאזן <ArrowLeft size={16}/></span>
  </button>
}

function Facility19BalanceCard({ balance }) {
  const diffClass = balance.balance > 0 ? 'positive' : balance.balance < 0 ? 'negative' : 'neutral'
  return <article className={`forecast-card facility42-balance-card facility19-balance-card ${diffClass}`}>
    <div className="forecast-head"><div><small>מאזן תשומות / תפוקות</small><h3>מאזן מתקן 19</h3><div className="forecast-resource"><b>1119 + 777</b><span>מול WG / SMALL PACKS בתחנה 1519</span></div></div><span className="status-badge no-target">ללא יעד</span></div>
    <div className="balance-card-main"><div><span>באלק שיוצר</span><b>{fmt(balance.bulk)}</b></div><div><span>סה״כ תפוקה</span><b>{fmt(balance.packed)}</b></div></div>
    <div className="balance-card-lines facility19-balance-lines"><span>WG<strong>{fmt(balance.byType['WG'])}</strong></span><span>SMALL PACKS<strong>{fmt(balance.byType['SMALL PACKS'])}</strong></span></div>
    <div className="balance-card-footer"><div><span>יתרה</span><b>{balance.balance > 0 ? '+' : ''}{fmt(balance.balance)}</b></div><div><span>ניצול באלק</span><b>{balance.bulk ? pctFmt(balance.utilization) : '—'}</b></div></div>
    <small className="balance-card-note">המאזן מחושב לפי טווח התאריכים שנבחר. פער יומי יכול לנבוע מתזמון שונה בין ייצור הבאלק לאריזה.</small>
  </article>
}

function ForecastCard({ facility, facilities, resource, packagingType, routingGroup, station, lineName, target, actual, pct, remaining, requiredDaily, recentAverage, provenMax, forecast, remainingWorkdays, state, label, selected, onClick }) {
  return <article className={`forecast-card ${state} ${selected ? 'selected' : ''}`} onClick={onClick} role="button" tabIndex="0">
    <div className="forecast-head"><div><small>משאב / מתקן</small><h3>{resource || facility}</h3>{(facilities || []).includes('1542') && packagingType && <div className="forecast-packaging-type"><span>קו אריזה</span><b>{packagingType}</b></div>}{(routingGroup || station) && <div className="forecast-resource"><b>{station || routingGroup}</b><span>{lineName || routingGroup}</span>{routingGroup && <small>{routingGroup}</small>}</div>}</div><StatusBadge state={state} label={label}/></div>
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
