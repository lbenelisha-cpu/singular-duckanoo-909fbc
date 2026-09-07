// Delivered quantity is the cumulative output of a process order, not a daily movement.
// Keep source/cloud snapshots untouched; derive daily increments before any date filter.
export function productionDailyQuantities(rows) {
  const groups = new Map(), result = []
  for (const row of rows) {
    const day = row.productionDay
    if (!row.order || !/^\d{4}-\d{2}-\d{2}$/.test(day || '')) {
      result.push({ ...row, cumulativeQty:row.qty, quantityBasis:'unresolved' })
      continue
    }
    const key = JSON.stringify([row.order,row.batch,row.material,row.facility].map(v=>String(v ?? '').trim()))
    if (!groups.has(key)) groups.set(key, new Map())
    const days = groups.get(key)
    const previous = days.get(day)
    // Equal timestamps keep the first row: imports place the newest snapshot first.
    if (!previous || String(row.snapshotTime || '') > String(previous.snapshotTime || '')) days.set(day,row)
  }
  for (const days of groups.values()) {
    let previous = 0
    for (const [day,row] of [...days].sort(([a],[b])=>a.localeCompare(b))) {
      const cumulative = Number(row.qty) || 0
      result.push({ ...row, productionDay:day, cumulativeQty:cumulative, previousCumulativeQty:previous, qty:cumulative-previous, quantityBasis:'daily-difference' })
      // Negative differences are corrections, not new production; do not clamp them to zero.
      previous = cumulative
    }
  }
  return result
}
