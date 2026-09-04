import assert from 'node:assert/strict'
import XLSX from 'xlsx'
import { buildResourceRows } from '../src/resourceEngine.js'
import { importedTargetValues, isApprovedTargetResource } from '../src/targetRules.js'

const workbookPath = process.argv[2]
if (!workbookPath) throw new Error('Usage: node scripts/verify-september-targets.mjs <targets.xlsx>')

const workbook = XLSX.readFile(workbookPath, { raw:true })
const sheet = workbook.Sheets.Sum
assert.ok(sheet, 'Missing Sum sheet')
const matrix = XLSX.utils.sheet_to_json(sheet, { header:1, defval:'', raw:true })
const headerIndex = matrix.findIndex(row => row.some(value => String(value).trim().toLowerCase() === 'capacity') && row.some(value => String(value).trim().toLowerCase() === 'plan'))
assert.ok(headerIndex >= 0, 'Missing Capacity/Plan header')
const header = matrix[headerIndex].map(value => String(value).trim().toLowerCase())
const capacityIndex = header.indexOf('capacity')
const planIndex = header.indexOf('plan')
const source = new Map(matrix.slice(headerIndex + 1).map(row => [String(row[0]).trim().toUpperCase(), row]))

const imported = name => {
  const row = source.get(name.toUpperCase())
  assert.ok(row, `Missing source row: ${name}`)
  return importedTargetValues(row[planIndex], row[capacityIndex])
}

assert.equal(imported('CS (25,40)').target, 8000)
assert.equal(imported('Saflufenacil Tech (25,40)').target, 1000)
assert.equal(imported('Galigan (25,40)').target, 45000)
assert.equal(isApprovedTargetResource('Galigan  iso (42)'), false)
assert.equal(isApprovedTargetResource('LINURON'), false)
assert.equal(isApprovedTargetResource('Phenol Oxime (41)'), false)
assert.equal(isApprovedTargetResource('D. Damascone (41)'), true)

const rows = buildResourceRows({
  production:[{ date:new Date('2026-09-03'), productionDay:'2026-09-03', facility:'1541', material:'PHENOL', desc:'Phenol Oxime', qty:123 }],
  targets:[{ month:'2026-09', resource:'D. Damascone (41)', facility:'1541', facilities:['1541'], target:20000, capacity:40000, materials:['DAMASCONE'] }],
  planningMonth:'2026-09',
  fallbackFacilities:['1541'],
  now:new Date('2026-09-04'),
})
assert.deepEqual(rows.map(row => row.resource), ['D. Damascone (41)'])
console.log('September target rules verified')
