import fs from 'node:fs'
import path from 'node:path'

const projectRoot = process.cwd()
const filePath = path.join(projectRoot, 'src', 'DashboardApp.jsx')

if (!fs.existsSync(filePath)) {
  console.error('ERROR: src/DashboardApp.jsx was not found. Run this command from the project root.')
  process.exit(1)
}

let source = fs.readFileSync(filePath, 'utf8').replace(/\r\n/g, '\n')
const original = source

const backupPath = filePath.replace(/\.jsx$/, `.BEFORE_BATCH_MATERIAL_${Date.now()}.jsx`)
fs.copyFileSync(filePath, backupPath)

const fail = message => {
  console.error(`ERROR: ${message}`)
  console.error(`Backup created at: ${backupPath}`)
  process.exit(1)
}

// 1) State: Batch + Material are the only identity fields used by Batch Control Center.
const selectedBatchState = "  const [selectedBatch, setSelectedBatch] = useState('')"
if (!source.includes('selectedBatchMaterial')) {
  if (!source.includes(selectedBatchState)) fail('selectedBatch state was not found')
  source = source.replace(
    selectedBatchState,
    `${selectedBatchState}\n  const [selectedBatchMaterial, setSelectedBatchMaterial] = useState('')`
  )
}

// 2) Add a direct Batch+Material quality index while preserving the existing byBatch index.
if (!source.includes('const byBatchMaterial = new Map()')) {
  const marker = '    const byBatch = new Map()\n'
  if (!source.includes(marker)) fail('qualityIndex byBatch map was not found')
  source = source.replace(marker, `${marker}    const byBatchMaterial = new Map()\n`)
}

if (!source.includes('const batchMaterialRows = byBatchMaterial.get(batchMaterialKey)')) {
  const marker = `      const list = byBatch.get(batch) || []\n      list.push(row)\n      byBatch.set(batch, list)`
  if (!source.includes(marker)) fail('qualityIndex byBatch population block was not found')
  const replacement = `${marker}\n\n      // Factory rule: Batch + Material is the unique quality identity.\n      const rowMaterial = normalize(row.material)\n      if (rowMaterial) {\n        const batchMaterialKey = \`${'${batch}'}|${'${rowMaterial}'}\`\n        const batchMaterialRows = byBatchMaterial.get(batchMaterialKey) || []\n        batchMaterialRows.push(row)\n        byBatchMaterial.set(batchMaterialKey, batchMaterialRows)\n      }`
  source = source.replace(marker, replacement)
}

// Add the new index to qualityIndex return value.
source = source.replace(
  'return { byBatch, rejected, approved, material, latestByBatch, latestByBatchLot }',
  'return { byBatch, byBatchMaterial, rejected, approved, material, latestByBatch, latestByBatchLot }'
)
if (!source.includes('return { byBatch, byBatchMaterial,')) {
  fail('qualityIndex return object could not be updated')
}

// 3) Replace Batch Control lookup logic with exact Batch+Material only.
const selectedStart = source.indexOf('  const selectedBatchData = useMemo(() => {')
const openStart = source.indexOf('  const openBatchCard', selectedStart)
if (selectedStart < 0 || openStart < 0) fail('selectedBatchData/openBatchCard block was not found')

const selectedBlock = `  const selectedBatchData = useMemo(() => {\n    const batch = normalize(selectedBatch)\n    const material = normalize(selectedBatchMaterial)\n    if (!batch || !material) return null\n\n    const batchMaterialKey = \`${'${batch}'}|${'${material}'}\`\n    return {\n      batch,\n      production: prod.filter(row =>\n        normalize(row.batch) === batch && normalize(row.material) === material\n      ),\n      quality: qualityIndex.byBatchMaterial.get(batchMaterialKey) || [],\n      deviations: enrichedDeviationRows.filter(row =>\n        normalize(row.batch) === batch && normalize(row.material) === material\n      ),\n    }\n  }, [selectedBatch, selectedBatchMaterial, prod, qualityIndex, enrichedDeviationRows])\n\n`
source = source.slice(0, selectedStart) + selectedBlock + source.slice(openStart)

// 4) Replace openBatchCard itself, ignoring Order/Inspection Lot/Routing/etc.
const openBlockStart = source.indexOf('  const openBatchCard', selectedStart)
const dataMonthsStart = source.indexOf('  const dataMonths', openBlockStart)
if (openBlockStart < 0 || dataMonthsStart < 0) fail('openBatchCard boundary was not found')
const openBlock = `  const openBatchCard = (batch, material = '') => {\n    if (!batch || !material) return\n    setSelectedBatch(normalize(batch))\n    setSelectedBatchMaterial(normalize(material))\n  }\n\n`
source = source.slice(0, openBlockStart) + openBlock + source.slice(dataMonthsStart)

// 5) Every visible Batch button passes Batch + Material only.
source = source.replaceAll('openBatchCard(r.batch, r.material, r.order)', 'openBatchCard(r.batch, r.material)')
source = source.replaceAll('openBatchCard(r.batch)', 'openBatchCard(r.batch, r.material)')

// 6) Resource Control Center keeps the exact same visual display, but carries Material behind the Batch button.
source = source.replace(
  'onOpenBatch={batch => { setSelectedResource(null); openBatchCard(batch) }}',
  'onOpenBatch={(batch, material) => { setSelectedResource(null); openBatchCard(batch, material) }}'
)
source = source.replace(
  'onOpenBatch={(batch, material) => { setSelectedResource(null); openBatchCard(batch, material, undefined) }}',
  'onOpenBatch={(batch, material) => { setSelectedResource(null); openBatchCard(batch, material) }}'
)

const batchesLine = "  const batches = [...new Set(rows.map(row => row.batch).filter(Boolean))]"
if (source.includes(batchesLine) && !source.includes('const batchItems = [...new Map(')) {
  source = source.replace(
    batchesLine,
    "  const batchItems = [...new Map(rows.filter(row => row.batch && row.material).map(row => [`${normalize(row.batch)}|${normalize(row.material)}`, { batch:row.batch, material:row.material }])).values()]\n  const batches = [...new Set(batchItems.map(item => item.batch))]"
  )
}
source = source.replace(
  '{batches.slice(0,24).map(batch => <button key={batch} onClick={() => onOpenBatch(batch)}>{batch}<ArrowLeft size={15}/></button>)}',
  '{batchItems.slice(0,24).map(item => <button key={`${item.batch}|${item.material}`} onClick={() => onOpenBatch(item.batch, item.material)}>{item.batch}<ArrowLeft size={15}/></button>)}'
)

// 7) Clear both identity fields when closing the Batch card.
source = source.replaceAll(
  "onClose={() => setSelectedBatch('')}",
  "onClose={() => { setSelectedBatch(''); setSelectedBatchMaterial('') }}"
)

// 8) Remove temporary debug logs used during diagnosis, if present.
source = source.replace(/\n\s*console\.log\(\s*['\"]QUALITY['\"][\s\S]*?normalize\(row\.order\)\s*\)\s*/m, '\n')
source = source.replace(/\n\s*console\.log\(['\"]Production Materials['\"][\s\S]*?console\.log\(['\"]Quality Candidates['\"][^\n]*\n/m, '\n')

if (source === original) {
  console.log('No changes were needed. The Batch+Material patch appears to already be installed.')
  process.exit(0)
}

fs.writeFileSync(filePath, source, 'utf8')
console.log('PATCH_OK')
console.log('Changed only: src/DashboardApp.jsx')
console.log('Rule installed: Batch + Material (exact match)')
console.log(`Backup: ${backupPath}`)
