import * as XLSX from 'xlsx'

const HEADER_COLS = [
  'Snag No.', 'Date Raised', 'Area / Floor', 'Zone / Grid',
  'Discipline', 'Contractor', 'Item Description', 'Priority',
  'Assigned To', 'Target Date', 'Completed Date', 'Days Overdue',
  'Status', 'CM Inspector', 'Evidence / Photo Ref.',
  'Corrective Action Required', 'Remarks',
]

const COL_WIDTHS = [10, 12, 16, 16, 18, 20, 40, 10, 18, 12, 12, 12, 18, 14, 18, 40, 24]

const STATUS_COLORS = {
  'Open':               'FFFCE4E1',
  'In Progress':        'FFE3F0FF',
  'On Hold':            'FFF3E5FF',
  'Pending Inspection': 'FFFFF9E5',
  'Approved':           'FFE8F5E9',
  'Completed':          'FFE8F5E9',
  'Closed':             'FFF5F5F5',
  'Rejected':           'FFFCE4EC',
  'Rework':             'FFFCE4EC',
}

const PRIORITY_COLORS = {
  'Critical': 'FFFFCCCC',
  'High':     'FFFFF9C4',
  'Medium':   'FFE3F2FD',
  'Low':      'FFE8F5E9',
}

function daysBetween(targetDate, completedDate) {
  if (!targetDate) return 0
  const target = new Date(targetDate)
  const end = completedDate ? new Date(completedDate) : new Date()
  const diff = Math.floor((end - target) / (1000 * 60 * 60 * 24))
  return diff > 0 ? diff : 0
}

function headerStyle(bgHex) {
  return {
    font: { bold: true, color: { rgb: 'FFFFFFFF' }, name: 'Arial', sz: 10 },
    fill: { fgColor: { rgb: bgHex || 'FF1A2744' }, patternType: 'solid' },
    alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
    border: {
      top:    { style: 'thin', color: { rgb: 'FF9CA3AF' } },
      bottom: { style: 'thin', color: { rgb: 'FF9CA3AF' } },
      left:   { style: 'thin', color: { rgb: 'FF9CA3AF' } },
      right:  { style: 'thin', color: { rgb: 'FF9CA3AF' } },
    },
  }
}

function cellStyle(bgHex) {
  return {
    font: { name: 'Arial', sz: 9 },
    fill: bgHex ? { fgColor: { rgb: bgHex }, patternType: 'solid' } : undefined,
    alignment: { vertical: 'top', wrapText: true },
    border: {
      top:    { style: 'hair', color: { rgb: 'FFE5E7EB' } },
      bottom: { style: 'hair', color: { rgb: 'FFE5E7EB' } },
      left:   { style: 'hair', color: { rgb: 'FFE5E7EB' } },
      right:  { style: 'hair', color: { rgb: 'FFE5E7EB' } },
    },
  }
}

export function exportToExcel(snags, lang = 'en') {
  const wb = XLSX.utils.book_new()

  // ── Sheet 1: Snag List ─────────────────────────────────────────
  const wsData = []

  // Title rows
  wsData.push([{ v: 'P25 TRITON PROJECT — SNAG LIST / PUNCH LIST TRACKER', t: 's' }])
  wsData.push([
    { v: 'Project', t: 's' }, { v: 'Phoenix P25 Triton', t: 's' },
    null, null, null, null, null, null,
    { v: 'Owner', t: 's' }, { v: 'Phoenix Group', t: 's' },
  ])
  wsData.push([
    { v: 'Exported', t: 's' }, { v: new Date().toLocaleDateString('en-GB'), t: 's' },
    null, null, null, null, null, null,
    { v: 'Location', t: 's' }, { v: 'Hyderabad, India', t: 's' },
  ])
  wsData.push([
    { v: 'Prepared By', t: 's' }, { v: 'CM Team — JLCM', t: 's' },
    null, null, null, null, null, null,
    { v: 'Document Status', t: 's' }, { v: 'Working Tracker', t: 's' },
  ])
  wsData.push([]) // blank row

  // Header row
  wsData.push(HEADER_COLS.map(h => ({ v: h, t: 's' })))

  // Data rows
  snags.forEach(snag => {
    const overdue = daysBetween(snag.targetDate, snag.completedDate)
    wsData.push([
      { v: snag.id,                    t: 'n' },
      { v: snag.dateRaised,            t: 's' },
      { v: snag.area || '',            t: 's' },
      { v: snag.zone || '',            t: 's' },
      { v: snag.discipline || '',      t: 's' },
      { v: snag.contractor || '',      t: 's' },
      { v: snag.description || '',     t: 's' },
      { v: snag.priority || '',        t: 's' },
      { v: snag.assignedTo || '',      t: 's' },
      { v: snag.targetDate || '',      t: 's' },
      { v: snag.completedDate || '',   t: 's' },
      { v: overdue,                    t: 'n' },
      { v: snag.status || '',          t: 's' },
      { v: snag.cmInspector || '',     t: 's' },
      { v: snag.photoRef || '',        t: 's' },
      { v: snag.correctiveAction || '', t: 's' },
      { v: snag.remarks || '',         t: 's' },
    ])
  })

  const ws = XLSX.utils.aoa_to_sheet(wsData)

  // Column widths
  ws['!cols'] = COL_WIDTHS.map(w => ({ wch: w }))

  // Row heights
  ws['!rows'] = [
    { hpt: 28 }, // title
    { hpt: 16 }, { hpt: 16 }, { hpt: 16 }, { hpt: 8 }, // meta rows
    { hpt: 36 }, // header
    ...snags.map(() => ({ hpt: 48 })),
  ]

  // Merge title row across all columns
  ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: HEADER_COLS.length - 1 } }]

  // Apply styles
  const titleCell = ws['A1']
  if (titleCell) {
    titleCell.s = {
      font: { bold: true, sz: 13, color: { rgb: 'FFFFFFFF' }, name: 'Arial' },
      fill: { fgColor: { rgb: 'FF1A2744' }, patternType: 'solid' },
      alignment: { horizontal: 'center', vertical: 'center' },
    }
  }

  // Style header row (row index 5 = Excel row 6)
  HEADER_COLS.forEach((_, ci) => {
    const addr = XLSX.utils.encode_cell({ r: 5, c: ci })
    if (ws[addr]) ws[addr].s = headerStyle('FF1A2744')
  })

  // Style data rows
  snags.forEach((snag, ri) => {
    const rowIdx = 6 + ri
    const statusBg = STATUS_COLORS[snag.status]
    const priorityBg = PRIORITY_COLORS[snag.priority]

    HEADER_COLS.forEach((_, ci) => {
      const addr = XLSX.utils.encode_cell({ r: rowIdx, c: ci })
      if (!ws[addr]) ws[addr] = { v: '', t: 's' }
      // Status col (12) gets status color, priority col (7) gets priority color
      const bg = ci === 12 ? statusBg : ci === 7 ? priorityBg : (rowIdx % 2 === 0 ? 'FFFAFAFA' : undefined)
      ws[addr].s = cellStyle(bg)
    })
  })

  // ── Sheet 2: Dashboard ────────────────────────────────────────
  const total   = snags.length
  const open    = snags.filter(s => ['Open','In Progress','On Hold'].includes(s.status)).length
  const closed  = snags.filter(s => ['Completed','Closed','Approved'].includes(s.status)).length
  const critical = snags.filter(s => s.priority === 'Critical').length
  const high    = snags.filter(s => s.priority === 'High').length
  const medium  = snags.filter(s => s.priority === 'Medium').length
  const overdue = snags.filter(s =>
    !['Closed','Completed','Approved'].includes(s.status) &&
    s.targetDate && new Date(s.targetDate) < new Date()
  ).length
  const pending = snags.filter(s => s.status === 'Pending Inspection').length

  const statusCount = snags.reduce((a, s) => ({ ...a, [s.status]: (a[s.status]||0)+1 }), {})
  const contractorCount = snags.reduce((a, s) => ({ ...a, [s.contractor]: (a[s.contractor]||0)+1 }), {})

  const dashData = [
    ['P25 Snag List Dashboard'],
    [],
    ['SUMMARY', '', '', 'STATUS BREAKDOWN', '', '', 'PRIORITY BREAKDOWN'],
    ['Total Snags',   total,   '', 'Open',               statusCount['Open']||0,               '', 'Critical', critical],
    ['Open / Active', open,    '', 'In Progress',        statusCount['In Progress']||0,        '', 'High',     high],
    ['Closed',        closed,  '', 'On Hold',            statusCount['On Hold']||0,            '', 'Medium',   medium],
    ['Critical',      critical,'', 'Pending Inspection', statusCount['Pending Inspection']||0, '', 'Low',      total-critical-high-medium],
    ['High Priority', high,    '', 'Approved',           statusCount['Approved']||0],
    ['Overdue',       overdue, '', 'Completed',          statusCount['Completed']||0],
    ['Pending Insp.', pending, '', 'Closed',             statusCount['Closed']||0],
    [],
    ['CONTRACTOR BREAKDOWN'],
    ...Object.entries(contractorCount).map(([c, n]) => [c, n]),
  ]

  const wsDash = XLSX.utils.aoa_to_sheet(dashData)
  wsDash['!cols'] = [{ wch: 18 }, { wch: 10 }, { wch: 4 }, { wch: 20 }, { wch: 10 }, { wch: 4 }, { wch: 12 }, { wch: 10 }]

  // ── Sheet 3: Activity Log ─────────────────────────────────────
  const logHeader = ['Snag No.', 'Area', 'Contractor', 'Date', 'Action', 'By']
  const logRows = []
  snags.forEach(snag => {
    (snag.activityLog || []).forEach(log => {
      logRows.push([snag.id, snag.area, snag.contractor, log.date, log.action, log.by])
    })
  })
  const logData = [logHeader, ...logRows]
  const wsLog = XLSX.utils.aoa_to_sheet(logData)
  wsLog['!cols'] = [{ wch: 10 }, { wch: 16 }, { wch: 20 }, { wch: 12 }, { wch: 40 }, { wch: 16 }]

  // Add sheets
  XLSX.utils.book_append_sheet(wb, ws, 'P25 Snag List')
  XLSX.utils.book_append_sheet(wb, wsDash, 'Dashboard')
  XLSX.utils.book_append_sheet(wb, wsLog, 'Activity Log')

  // Export
  const dateStr = new Date().toISOString().split('T')[0].replace(/-/g, '')
  XLSX.writeFile(wb, `P25_PunchList_Export_${dateStr}.xlsx`)
}
