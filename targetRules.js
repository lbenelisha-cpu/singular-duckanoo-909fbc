const normalizeName = value => String(value ?? '').trim().replace(/\s+/g, ' ')

const APPROVED_TARGET_RESOURCES = new Set([
  'EC (23)','SHAKED ISO 42','SHAKED ISO 23','LQ 1LT (42)','LQ 5 LT (42)','LQ 10/20 LT (42)','LQ 43','SC (28)','WG (19)','WG SMALL PACKS (19)',
  '24F128','24F','EC (25)','DIURON (40)','TOLUREX (40)','CS (25,40)','BROMACIL (25,40)','GALIGAN (25,40)',
  'PROPA PREMIX (25,40)','FLUOROCHLORIDON (25,40)','SAFLUFENACIL TECH (25,40)','METAZACHLOR (41)',
  'ATRALONE (41)','NANA (41)','D. DAMASCONE (41)'
])

export const isApprovedTargetResource = value => APPROVED_TARGET_RESOURCES.has(normalizeName(value).toUpperCase())

export const parseTargetNumber = value => {
  if (value === null || value === undefined || value === '' || /^\s*-+\s*$/.test(String(value)) || /DIV\/0/i.test(String(value))) return 0
  const text = String(value).trim()
  const negative = /^\(.*\)$/.test(text)
  const number = Number(text.replace(/[(),%\s]/g, '').replace(/,/g, ''))
  return Number.isFinite(number) ? (negative ? -number : number) : 0
}

export const importedTargetValues = (planValue, capacityValue) => {
  const capacity = parseTargetNumber(capacityValue) * 1000
  const plan = parseTargetNumber(planValue) * 1000
  return { capacity, target:plan > 0 ? plan : capacity }
}
