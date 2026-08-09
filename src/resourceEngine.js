import { mappingMatchesTarget, stationFamily, targetMappingKey } from './mappingEngine'
const text = value => String(value ?? '').trim()
const upper = value => text(value).toUpperCase()
const isoDate = value => {
  if (!value) return ''
  const date = value instanceof Date ? value : new Date(value)
  if (Number.isNaN(date.getTime())) return ''
  return `${date.getFullYear()}-${String(date.getMonth()+1).padStart(2,'0')}-${String(date.getDate()).padStart(2,'0')}`
}
const monthOf = value => isoDate(value).slice(0, 7)
const isWorkday = date => ![5, 6].includes(date.getDay())
const daysInMonth = key => {
  const [year, month] = key.split('-').map(Number)
  return new Date(year, month, 0).getDate()
}
const workdayCount = (key, startDay = 1, endDay = daysInMonth(key)) => {
  const [year, month] = key.split('-').map(Number)
  let count = 0
  for (let day = startDay; day <= endDay; day += 1) {
    if (isWorkday(new Date(year, month - 1, day))) count += 1
  }
  return count
}

const resourceCode = row => upper(`${row.routingGroup || ''} ${row.routingDescription || ''} ${row.resource || ''} ${row.line || ''}`)
const isSmallPack19 = row => ['19PWG-01','19PWG-05','19PWG-15'].some(code => resourceCode(row).includes(code))

const matchProductionToTarget = (row, target, manualMappings = []) => {
  const manualDecision = mappingMatchesTarget(row, target, manualMappings)
  if (manualDecision !== null) return manualDecision

  const targetName = upper(target.resource)
  const station = upper(row.facility)
  const route = resourceCode(row)

  // Sprint 11.9.0 — business rules approved for monthly targets.
  // Specific line/resource rules always win over broad station rules.
  if (/^SC\s*\(28\)/.test(targetName)) return station === '1528'

  // LQ 43 has priority over the broad EC(23) station rule to avoid double counting.
  const isLq43Resource = route.includes('43-P-A') || route.includes('43-P-B')
  if (/^LQ\s*43\b/.test(targetName)) return isLq43Resource
  if (/^EC\s*\(23\)/.test(targetName)) return station === '1523' && !isLq43Resource

  // Facility 42: identify packaging line by the actual SAP Routing group.
  // Keep the previous 42-P-* aliases and Description text as fallbacks for older files.
  const isFacility42Zfin = station === '1542' && upper(row.orderType).includes('ZFIN')
  const lq1 = route.includes('LQ-P-1') || route.includes('42-P-02') || route.includes('LIQUID 1 LITER')
  const lq5 = route.includes('LQ-P-5') || route.includes('42-P-03') || route.includes('LIQUID 5 LITER')
  const lq1020 = route.includes('LQ-P-10') || route.includes('42-P-04') || route.includes('LIQUID 10/20 LITER')
  if (/LQ\s*1\s*(LT|L)\b/.test(targetName)) return isFacility42Zfin && lq1
  if (/LQ\s*5\s*(LT|L)\b/.test(targetName)) return isFacility42Zfin && lq5
  if (/LQ\s*10\s*\/\s*20/.test(targetName)) return isFacility42Zfin && lq1020

  if (/WG\s*SMALL\s+PACKS?\s*\(19\)/.test(targetName)) return station === '1519' && isSmallPack19(row)
  if (/^WG\s*\(19\)/.test(targetName)) return station === '1519' && !isSmallPack19(row)

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
  const monthTargets = targets.filter(target => target.month === planningMonth)
  const sourceRows = monthTargets.length ? monthTargets : fallbackFacilities.map(facility => ({ facility, facilities:[facility], resource:`מתקן ${facility}`, target:0, capacity:0, descriptionTokens:[] }))

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
      routingGroup:targetRow.routingGroup || '', station:targetRow.station || facilities.join(','),
      description:(targetRow.descriptionTokens || []).join(' / '), packagingType:packagingTypeForTarget(targetRow), lineName:targetRow.lineName || targetRow.resource || '', activity:targetRow.activity || 'ייצור / אריזה',
      target, capacity, actual, pct:target ? actual / target * 100 : 0, remaining, requiredDaily,
      average, recentAverage, provenMax, forecast, capacityForecast, elapsedWorkdays, remainingWorkdays, totalWorkdays,
      orders:orders.size, batches:batches.size, productionRows:rows, dailyEntries, state, label,
      notes:targetRow.notes || '', recyclingPlan:Number(targetRow.recyclingPlan)||0, recycled:Number(targetRow.recycled)||0,
      forPacking:Number(targetRow.forPacking)||0, restrictedRecycling:Number(targetRow.restrictedRecycling)||0,
      restrictedDisposal:Number(targetRow.restrictedDisposal)||0,
    }
  })
}
