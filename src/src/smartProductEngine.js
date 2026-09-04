const clean = value => String(value ?? '').trim()
const upper = value => clean(value).toUpperCase()
const compact = value => upper(value)
  .replace(/[^A-Z0-9\u0590-\u05FF]+/g, ' ')
  .replace(/\s+/g, ' ')
  .trim()

const STOP_WORDS = new Set(['THE','AND','FOR','OF','PACK','PACKING','GENERAL','PRODUCTION','ISO','TECH','TECHNICAL','CRUDE','BATCHES'])

export const stationFamily = station => {
  const digits = clean(station).replace(/\D/g, '')
  if (digits.length < 2) return clean(station)
  const prefix = digits.slice(0, 2)
  const suffix = digits.slice(-2)
  if (['11','12','15'].includes(prefix)) return `15${suffix}`
  return digits
}

const targetTokens = resource => compact(resource)
  .replace(/\([^)]*\)/g, ' ')
  .split(' ')
  .filter(token => token.length >= 2 && !STOP_WORDS.has(token))

const catalogText = entry => compact([
  entry.family,
  entry.segmentationClass,
  entry.type,
  entry.itemDescription,
  entry.res,
].join(' '))

const hasToken = (haystack, token) => {
  if (!token) return false
  if (haystack.includes(token)) return true
  const aliases = {
    TOLUREX: ['CHLOROTOLURON','CLTN'],
    CHLOROTOLURON: ['TOLUREX','CLTN'],
    DIURON: ['DRON'],
    METAZACHLOR: ['METZACHLOR','MTZC'],
    METZACHLOR: ['METAZACHLOR','MTZC'],
    ATRALONE: ['ATRL'],
    GALIGAN: ['OXYFLUORFEN','OXFF'],
    OXYFLUORFEN: ['GALIGAN','OXFF'],
    PROPAQUIZAFOP: ['PPQZ'],
    FLUOROCHLORIDON: ['FLUROCHLORIDONE','FLUOROCHLORIDONE'],
  }
  return (aliases[token] || []).some(alias => haystack.includes(alias))
}

export const buildCatalogIndex = targets => {
  const byMaterial = new Map()
  ;(targets || []).filter(row => row?.__productCatalog).forEach(entry => {
    ;[entry.itemCode, entry.oldBulk].map(clean).filter(Boolean).forEach(code => byMaterial.set(code, entry))
  })
  return byMaterial
}

export const scoreProductTarget = (row, target, catalogEntry) => {
  const family = stationFamily(row.facility)
  const facilities = target.facilities?.length ? target.facilities : [target.facility].filter(Boolean)
  if (!facilities.includes(family)) return -1

  const resource = compact(target.resource)
  const rowText = compact(`${row.desc || ''} ${row.routingDescription || ''}`)
  const catalog = catalogEntry ? catalogText(catalogEntry) : ''
  const haystack = `${catalog} ${rowText}`.trim()
  const tokens = targetTokens(target.resource)

  let score = facilities.length === 1 ? 60 : 50
  if (catalogEntry) score += 35

  let tokenHits = 0
  tokens.forEach(token => {
    if (hasToken(haystack, token)) tokenHits += 1
  })
  score += tokenHits * 20

  const resourceWords = resource.split(' ')
  if (resourceWords.includes('EC') && /\bEC\b/.test(haystack)) score += 25
  if (resourceWords.includes('SC') && /\bSC\b/.test(haystack)) score += 25
  if (resourceWords.includes('WG') && /\bWG\b/.test(haystack)) score += 25
  if (resourceWords.includes('WP') && /\bWP\b/.test(haystack)) score += 25
  if (resourceWords.includes('CS') && /\bCS\b/.test(haystack)) score += 25

  return score
}

export const assignProductionTargets = ({ production = [], targetRows = [], catalogRows = [] }) => {
  const catalogIndex = buildCatalogIndex(catalogRows)
  const assignments = new Map(targetRows.map(target => [target, []]))
  const unmatched = []

  production.forEach(row => {
    const family = stationFamily(row.facility)
    const candidates = targetRows.filter(target => {
      const facilities = target.facilities?.length ? target.facilities : [target.facility].filter(Boolean)
      return facilities.includes(family)
    })
    if (!candidates.length) { unmatched.push(row); return }

    const material = clean(row.material)
    const catalogEntry = catalogIndex.get(material)
    let best = null
    let bestScore = -1
    candidates.forEach(target => {
      const score = scoreProductTarget(row, target, catalogEntry)
      if (score > bestScore) { best = target; bestScore = score }
    })

    if (best && bestScore >= 60) assignments.get(best).push(row)
    else unmatched.push(row)
  })

  return { assignments, unmatched }
}
