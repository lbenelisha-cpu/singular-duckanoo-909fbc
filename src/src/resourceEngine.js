import { mappingMatchesTarget, stationFamily, targetMappingKey } from './mappingEngine.js'
const text = value => String(value ?? '').trim()
const upper = value => text(value).toUpperCase()
const isoDate = value => {
  if (!value) return ''
  const date = value instanceof Date ? value : new Date(value)
  if (Number.isNaN(date.getTime())) return ''
  return `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`
}
const monthOf = value => isoDate(value).slice(0, 7)
const isWorkday = date => true // All calendar days are operating days, including Friday and Saturday
const daysInMonth = key => {
  const [year, month] = key.split('-').map(Number)
  return new Date(year, month, 0).getDate()
}
const workdayCount = (key, startDay = 1, endDay = daysInMonth(key)) => {
  const first = Math.max(1, Number(startDay) || 1)
  const last = Math.min(daysInMonth(key), Number(endDay) || daysInMonth(key))
  return last >= first ? last - first + 1 : 0
}

const resourceCode = row => upper(`${row.routingGroup || ''} ${row.routingDescription || ''} ${row.resource || ''} ${row.line || ''}`)
const isSmallPack19 = row => upper(row.mappedResource) === 'WG SMALL PACKS (19)' || ['19PWG-01','19PWG-05','19PWG-15','19-P-03','19-P-04','19-P-05','FCLT-19'].some(code => resourceCode(row).includes(code))

const matchProductionToTarget = (row, target, manualMappings = []) => {
  // Approved business rules below take precedence over manual/legacy mappings.
  // This prevents old mappings from leaking 10L rows into 1L or unrelated stations into product families.
  const manualDecision = mappingMatchesTarget(row, target, manualMappings)

  const targetName = upper(target.resource)
  const station = upper(row.facility)
  const route = resourceCode(row)
  const directResource = upper(row.mappedResource)

  // Sprint 11.9.36 — PROD LINE first-pass mapping. For resources whose PROD LINE
  // uniquely identifies a dashboard card, avoid repeated description/material
  // tests. Facility 25 remains on the material rules below because it has several
  // product-family target cards in addition to EC (25).
  const directMappedCards = new Set([
    'LQ 1LT (42)','LQ 5 LT (42)','LQ 10/20 LT (42)','LQ 43',
    'WG (19)','WG SMALL PACKS (19)','EC (23)','24F','24F128','SC (28)',
    'DIURON (40)','TOLUREX (40)'
  ])
  if (directResource && directMappedCards.has(directResource)) {
    const isSameCard = targetName === directResource
    if (directResource.startsWith('LQ ') && station === '1542') {
      return isSameCard && upper(row.orderType).includes('ZFIN')
    }
    if (directResource.startsWith('WG')) {
      const bulkMarkerText = upper(`${row.desc || ''} ${row.routingDescription || ''}`)
      return isSameCard && !/(^|\D)777(\D|$)/.test(bulkMarkerText)
    }
    return isSameCard
  }

  // Sprint 11.9.0 — business rules approved for monthly targets.
  // Specific line/resource rules always win over broad station rules.
  if (/^SC\s*\(28\)/.test(targetName)) return station === '1528'

  // Correct facility mapping for EC, Shaked ISO and LQ43.
  const isLq43Resource = route.includes('43-P-A') || route.includes('43-P-B') ||
    /(^|\s)LQ-P-(1|5|10)(\s|$)/.test(route)
  if (/^LQ\s*43\b/.test(targetName)) return station === '1543' &&
    upper(row.orderType).includes('ZFIN') && isLq43Resource
  if (/^EC\s*\(23\)/.test(targetName)) return station === '1523'
  if (/^EC\s*\(25\)/.test(targetName)) return station === '1525'
  if (/^SHAKED\s+ISO\s+42$/.test(targetName)) {
    // Station 1142 rows marked 999 are Facility 42 bulk input and belong ONLY
    // to the dedicated Facility 42 balance card, never to Shaked ISO 42.
    const bulkMarkerText = upper(`${row.desc || ''} ${row.routingDescription || ''}`)
    return station === '1142' && !/(^|\D)999(\D|$)/.test(bulkMarkerText)
  }
  if (/^SHAKED\s+ISO\s+23$/.test(targetName)) return station === '1123'
  if (/^PILOT\s*\(1521\)$/.test(targetName)) return station === '1521'

  // Facility 24 has two separate monthly target cards. Prefer the DATA-sheet
  // material lists when available so the same production is not counted in
  // both cards; keep the description marker as a legacy fallback.
  if (/^24F(?:128)?$/.test(targetName)) {
    if (station !== '1524') return false
    const material = upper(row.material)
    const isLrhPack24 = material === '10000007617'
    // Approved reassignment: all LRH Pack (24) production belongs to 24F128.
    // It must be excluded from 24F so the quantity is counted exactly once.
    if (targetName === '24F128' && isLrhPack24) return true
    if (targetName === '24F' && isLrhPack24) return false
    const targetMaterials = new Set((target.materials || []).map(upper).filter(Boolean))
    if (targetMaterials.size) return targetMaterials.has(material)
    const productionText = upper(`${row.desc || ''} ${row.routingDescription || ''} ${row.routingGroup || ''}`)
    return targetName === '24F128' ? productionText.includes('24F128') : !productionText.includes('24F128')
  }

  // Facility 42: identify packaging line by the actual SAP Routing group.
  // Keep the previous 42-P-* aliases and Description text as fallbacks for older files.
  const isFacility42Zfin = station === '1542' && upper(row.orderType).includes('ZFIN')
  // Exact routing matching is critical: `LQ-P-1` must NOT match `LQ-P-10`.
  const lq1 = /(^|\s)LQ-P-1(\s|$)/.test(route) || route.includes('42-P-02') || route.includes('LIQUID 1 LITER')
  const lq5 = /(^|\s)LQ-P-5(\s|$)/.test(route) || route.includes('42-P-03') || route.includes('LIQUID 5 LITER')
  const lq1020 = /(^|\s)LQ-P-10(\s|$)/.test(route) || route.includes('42-P-04') || route.includes('LIQUID 10/20 LITER')
  if (/LQ\s*1\s*(LT|L)\b/.test(targetName)) return isFacility42Zfin && lq1
  if (/LQ\s*5\s*(LT|L)\b/.test(targetName)) return isFacility42Zfin && lq5
  if (/LQ\s*10\s*\/\s*20/.test(targetName)) return isFacility42Zfin && lq1020

  // Facility 19: rows marked 777 are bulk input and belong only to the dedicated
  // Facility 19 material-balance view. Never count them again as packed WG output.
  const facility19MarkerText = upper(`${row.desc || ''} ${row.routingDescription || ''}`)
  const isFacility19Bulk = /(^|\D)777(\D|$)/.test(facility19MarkerText)
  if (/WG\s*SMALL\s+PACKS?\s*\(19\)/.test(targetName)) return station === '1519' && !isFacility19Bulk && isSmallPack19(row)
  if (/^WG\s*\(19\)/.test(targetName)) return station === '1519' && !isFacility19Bulk && !isSmallPack19(row)

  // CS (25,40): approved product rule. Only material 20000000722 belongs to CS.
  // Keep the station scope explicit as 1525/1540 so legacy 1125/1140 aliases cannot leak in.
  if (/^CS\s*\(25\s*,\s*40\)/.test(targetName)) {
    return station === '1525' && upper(row.material) === '20000000722'
  }

  // Approved exact-material rules from the production quantities audit (12/08/2026).
  if (/^METAZACHLOR\b/.test(targetName)) {
    return station === '1541' && upper(row.material) === '30000000097'
  }
  if (/^SAFLUFENACIL\s+TECH\b/.test(targetName)) {
    const facilities = target.facilities?.length ? target.facilities.map(upper) : [upper(target.facility)].filter(Boolean)
    return facilities.includes(station) && upper(row.material) === '10000015999'
  }

  // Bromacil: production is reported under station 1540.
  // Prefer the DATA-sheet Item Code list when it is present on the target row.
  // The BRMC fallback keeps legacy/cached target rows working until the DATA mapping
  // is refreshed. Both ZFIN and ZSEM are valid for this production family.
  if (/^BROMACIL\b/.test(targetName)) {
    if (station !== '1540') return false
    const orderType = upper(row.orderType)
    if (orderType && !orderType.includes('ZFIN') && !orderType.includes('ZSEM')) return false

    const targetMaterials = new Set((target.materials || []).map(upper).filter(Boolean))
    if (targetMaterials.size) return targetMaterials.has(upper(row.material))

    const bromacilText = upper(`${row.desc || ''} ${row.routingDescription || ''}`)
    return /(^|\s)BRMC\d*/.test(bromacilText) || bromacilText.includes('BROMACIL')
  }

  // Sprint 11.9.18 — approved production audit overrides (20/08/2026).
  // Exact material assignment wins over the reporting storage location for target-card mapping.
  const approvedMaterialFacility = {
    '30000006846':'1525',
    '10000001434':'1541',
    '20000000716':'1523',
    '10000015999':'1528',
    '20000005829':'1525',
    '20000000692':'1525',
    '10000001198':'1523',
    '10000011346':'1523',
    '10000015919':'1540',
    '10000015939':'1540',
    '10000014392':'1540',
    '10000015938':'1540',
    '10000015930':'1540',
    '20000007617':'1524',
    '10000007617':'1524',
    '10000014393':'1524',
    '20000000246':'1524',
    '20000001538':'1524',
    '10000001477':'1524',
    '50000000089':'1541',
    '10000012624':'1541',
  }
  // Cleaning material: never count as production/packaging.
  if (upper(row.material) === 'CL10000013819') return false
  const approvedFacility = approvedMaterialFacility[upper(row.material)]
  if (approvedFacility) {
    const facilities = target.facilities?.length ? target.facilities.map(upper) : [upper(target.facility)].filter(Boolean)
    if (facilities.includes(approvedFacility)) return true
    // Also recognize standard card names when the target workbook uses a line/family label.
    if (approvedFacility === '1528' && /^SC\s*\(28\)/.test(targetName)) return true
    if (approvedFacility === '1523' && /^EC\s*\(23\)/.test(targetName)) return true
    if (approvedFacility === '1525' && /^EC\s*\(25\)/.test(targetName)) return true
    return false
  }

  // For all remaining resources, an explicit manual mapping may still override the generic DATA/family rule.
  if (manualDecision !== null) return manualDecision

  // Production families: DATA sheet Item Codes are the source of truth.
  // 1542 and 1519 are handled above and are deliberately excluded from this mechanism.
  const targetMaterials = new Set((target.materials || []).map(upper).filter(Boolean))
  if (targetMaterials.size) {
    const facilities = target.facilities?.length ? target.facilities : [target.facility].filter(Boolean)
    if (!facilities.map(upper).includes(station)) return false
    return targetMaterials.has(upper(row.material))
  }

  const facilities = target.facilities?.length ? target.facilities : [target.facility].filter(Boolean)
  const rowFamily = stationFamily(row.facility)
  const targetFamilies = facilities.map(stationFamily)
  if (!targetFamilies.includes(rowFamily)) return false

  if (rowFamily === '1542' && !upper(row.orderType).includes('ZFIN')) return false
  if (target.routingGroup && upper(row.routingGroup) !== upper(target.routingGroup)) return false
  const tokens = (target.descriptionTokens || []).map(upper).filter(Boolean)
  if (!tokens.length) return true
  const haystack = upper(`${row.routingDescription || ''} ${row.desc || ''}`)
  return tokens.some(token => haystack.includes(token))
}


const packagingTypeForTarget = target => {
  const resource = upper(target.resource)
  const tokens = (target.descriptionTokens || []).map(upper)
  if (/LQ\s*1\s*(LT|L)\b/.test(resource) || tokens.some(t => ['1L','1 L','1LT','1 LT'].includes(t))) return '1 ליטר'
  if (/LQ\s*5\s*(LT|L)\b/.test(resource) || tokens.some(t => ['5L','5 L','5LT','5 LT'].includes(t))) return '5 ליטר'
  if (/LQ\s*10\s*\/\s*20/.test(resource) || tokens.some(t => ['10L','10 L','20L','20 L','10/20'].includes(t))) return '10/20 ליטר'
  return ''
}

export function buildResourceRows({ production = [], targets = [], planningMonth = '', fallbackFacilities = [], manualMappings = [], now = new Date() }) {
  if (!planningMonth) return []
  const [year, month] = planningMonth.split('-').map(Number)
  const monthRows = production.filter(row => {
  const sourceDate = row.productionDay || row.date || row.finishDate
  return monthOf(sourceDate) === planningMonth
})

const latestDate = monthRows.reduce((latest, row) => {
  const sourceDate = row.productionDay || row.date || row.finishDate
  const date = sourceDate instanceof Date ? sourceDate : new Date(sourceDate)

  return Number.isNaN(date.getTime()) || (latest && latest >= date)
    ? latest
    : date
}, null)
  const currentMonth = now.getFullYear() === year && now.getMonth() + 1 === month
  const pastMonth = new Date(year, month, 0) < new Date(now.getFullYear(), now.getMonth(), 1)
  const asOfDay = currentMonth ? Math.min(now.getDate(), daysInMonth(planningMonth)) : pastMonth ? daysInMonth(planningMonth) : (latestDate?.getDate() || 0)
  const elapsedWorkdays = workdayCount(planningMonth, 1, Math.max(0, asOfDay))
  const totalWorkdays = workdayCount(planningMonth)
  const remainingWorkdays = Math.max(0, totalWorkdays - elapsedWorkdays)
  // The selected month's approved target rows are the sole source of planning
  // cards. Production that has no target must not create a card or borrow a
  // neighbouring target by row order / partial matching.
  const sourceRows = targets.filter(target => target.month === planningMonth)

  return sourceRows.map((targetRow, index) => {
    const rows = monthRows.filter(row => matchProductionToTarget(row, targetRow, manualMappings))
    const dailyMap = new Map()
    let actual = 0
    const orders = new Set()
    const batches = new Set()
    rows.forEach(row => {
      actual += Number(row.qty) || 0
      if (row.order) orders.add(row.order)
      if (row.batch) batches.add(row.batch)
      const day = row.productionDay || isoDate(row.date || row.finishDate)
      if (day) dailyMap.set(day, (dailyMap.get(day) || 0) + (Number(row.qty) || 0))
    })
    const dailyEntries = [...dailyMap.entries()].sort((a,b) => a[0].localeCompare(b[0]))
    const dailyValues = dailyEntries.map(([, value]) => value)
    const average = dailyValues.length ? actual / dailyValues.length : 0
    const recent = dailyValues.slice(-7)
    const recentAverage = recent.length ? recent.reduce((sum, value) => sum + value, 0) / recent.length : average
    const target = Number(targetRow.target) || 0
    const capacity = Number(targetRow.capacity) || 0
    const provenMax = dailyValues.length ? Math.max(...dailyValues) : (capacity ? capacity / Math.max(1, totalWorkdays) : 0)
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
    const facilities = targetRow.facilities?.length ? targetRow.facilities : [targetRow.facility].filter(Boolean)
    return {
      id:`${targetRow.resource || targetRow.facility || index}::${index}`, targetKey:targetMappingKey(targetRow),
      resource:targetRow.resource || `מתקן ${targetRow.facility}`,
      facility:facilities.join(','), facilities,
      routingGroup:targetRow.routingGroup || '', station:/^EC\s*\(23\)/i.test(String(targetRow.resource || '').trim()) ? '1523' : (targetRow.station || facilities.join(',')),
      description:(targetRow.descriptionTokens || []).join(' / '), packagingType:packagingTypeForTarget(targetRow), lineName:targetRow.lineName || targetRow.resource || '', activity:targetRow.activity || 'ייצור / אריזה',
      target, capacity, actual, pct:target ? actual / target * 100 : 0, remaining, requiredDaily,
      average, recentAverage, provenMax, forecast, capacityForecast, elapsedWorkdays, remainingWorkdays, totalWorkdays,
      orders:orders.size, batches:batches.size, productionRows:rows, dailyEntries, state, label,
      mappingStatus:targetRow.mappingStatus || '',
      mappingReason:targetRow.mappingReason || '',
      mappingCandidates:targetRow.mappingCandidates || [],
      mappingConfidence:targetRow.mappingConfidence || '',
      notes:targetRow.notes || '', recyclingPlan:Number(targetRow.recyclingPlan)||0, recycled:Number(targetRow.recycled)||0,
      forPacking:Number(targetRow.forPacking)||0, restrictedRecycling:Number(targetRow.restrictedRecycling)||0,
      restrictedDisposal:Number(targetRow.restrictedDisposal)||0,
    }
  })
}
