const clean = value => String(value ?? '').trim()
const upper = value => clean(value).toUpperCase().replace(/\s+/g, ' ')

export const stationFamily = station => {
  const digits = clean(station).replace(/\D/g, '')
  if (digits.length < 2) return clean(station)
  return `15${digits.slice(-2)}`
}

export const targetMappingKey = target => {
  const facilities = target?.facilities?.length
    ? target.facilities
    : [target?.facility].filter(Boolean)
  return [
    upper(target?.resource),
    facilities.map(stationFamily).sort().join(','),
    upper(target?.routingGroup),
  ].join('|')
}

export const productionMappingKey = row => {
  const family = stationFamily(row?.facility)
  const material = upper(row?.material)
  const description = upper(row?.desc)
  return `${family}|${material || description}`
}

export const approvedMappingForRow = (row, mappings = []) => {
  const rowKey = productionMappingKey(row)
  return mappings.find(mapping =>
    mapping &&
    mapping.active !== false &&
    mapping.status === 'approved' &&
    mapping.rowKey === rowKey
  ) || null
}

export const mappingMatchesTarget = (row, target, mappings = []) => {
  const mapping = approvedMappingForRow(row, mappings)
  if (!mapping) return null
  return mapping.targetKey === targetMappingKey(target)
}

export const normalizedDescription = value => upper(value)
