export const uploadDay = (now = new Date()) => new Intl.DateTimeFormat('en-CA', { timeZone:'Asia/Jerusalem', year:'numeric', month:'2-digit', day:'2-digit' }).format(now)
const canonical = value => {
  if (value instanceof Date) return value.toISOString()
  if (Array.isArray(value)) return value.map(canonical)
  if (value && typeof value === 'object') return Object.fromEntries(Object.keys(value).sort().filter(k => k !== '__sheet' && !(k === 'date' && value.__compactProduction)).map(k => [k, canonical(value[k])]))
  return value
}
// Compare groups so multiple legitimate results with the same business identity are preserved.
const group = (rows, key) => {
  const result = new Map()
  for (const row of rows) {
    const id = key(row)
    if (!result.has(id)) result.set(id, new Set())
    result.get(id).add(JSON.stringify(canonical(row)))
  }
  return new Map([...result].map(([id, values]) => [id, JSON.stringify([...values].sort())]))
}
export function calculateUploadStats(before, after, key, previous, now = new Date()) {
  const day = uploadDay(now)
  const oldRows = group(before, key), newRows = group(after, key)
  const sameDay = previous?.day === day
  const addedKeys = new Set(sameDay ? previous.addedKeys : [])
  const updatedKeys = new Set(sameDay ? previous.updatedKeys : [])
  let added = 0, updated = 0, removed = 0
  for (const [id, value] of newRows) {
    if (!oldRows.has(id)) { added++; addedKeys.add(id) }
    else if (oldRows.get(id) !== value) { updated++; updatedKeys.add(id) }
  }
  for (const id of oldRows.keys()) if (!newRows.has(id)) removed++
  return { day, addedKeys:[...addedKeys], updatedKeys:[...updatedKeys], addedToday:addedKeys.size, updatedToday:updatedKeys.size, lastAdded:added, lastUpdated:updated, lastRemoved:removed, total:after.length }
}
