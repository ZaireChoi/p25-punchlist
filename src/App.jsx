import { useState, useEffect, useCallback } from 'react'
import { loadSnags, saveSnags, onSnagsUpdated } from './lib/storage.js'
import { exportToExcel } from './lib/excel.js'

// ─── TRANSLATIONS ─────────────────────────────────────────────────
const T = {
  en: {
    projectTag: 'P25 TRITON · HYDERABAD',
    appTitle: 'Punch List Manager',
    dashboard: 'Dashboard', list: 'List', newSnag: 'New Snag',
    totalSnags: 'Total Snags', inProgress: 'In Progress',
    critical: 'Critical', overdue: 'Overdue',
    pendingAlert: (n) => `${n} Pending Inspection`,
    pendingAlertSub: 'Contractor complete → CM inspection required',
    view: 'View →',
    statusBreakdown: 'Status Breakdown', recentSnags: 'Recent Snags',
    snagList: 'Snag List', allStatus: 'All Status',
    allPriority: 'All Priority', allContractor: 'All Contractor',
    reset: '✕ Reset', noSnags: 'No matching snags found',
    addSnag: 'Register New Snag',
    areaFloor: 'Area / Floor *', zoneGrid: 'Zone / Grid',
    discipline: 'Discipline *', contractor: 'Contractor *',
    priority: 'Priority', assignedTo: 'Assigned To',
    targetDate: 'Target Date *', photoRef: 'Photo Reference',
    description: 'Snag Description *',
    descPlaceholder: 'Describe the defect in detail...',
    correctiveAction: 'Corrective Action Required',
    correctivePlaceholder: 'Required corrective action...',
    remarks: 'Remarks',
    submitSnag: '📋 Register Snag',
    select: 'Select...',
    backToList: '← Back to List',
    snagNo: 'Snag #',
    statusChange: 'Change Status',
    roleLabel: (r) => `${r} ▾`,
    activityLog: 'Activity Log',
    exportExcel: '📥 Export Excel',
    exportAll: 'Export All',
    emailTo: 'To:', emailSubject: 'Subject:',
    toastRegistered: (id) => `Snag #${id} registered!`,
    toastStatus: (s) => `Status updated: ${s}`,
    toastExported: 'Excel exported!',
    targetDateLabel: 'Target Date', cmInspector: 'CM Inspector',
    photoRefLabel: 'Photo Ref.', completedDate: 'Completed Date',
    overdueTag: 'Overdue', daysOverdue: (n) => `${n}d overdue`,
    emailNotice: '※ Email notification: configure SMTP in Settings',
    settings: 'Settings',
    settingsTitle: 'Settings / 설정',
    smtpNote: 'Email notifications require EmailJS configuration.',
    emailJsId: 'EmailJS Service ID',
    emailJsTemplate: 'EmailJS Template ID',
    emailJsKey: 'EmailJS Public Key',
    save: 'Save',
    saved: 'Saved!',
    clearData: '🗑 Clear All Data',
    clearConfirm: 'Delete ALL snag data? This cannot be undone.',
    importJson: '📂 Import JSON',
    exportJson: '📤 Export JSON',
    vendorEmails: 'Contractor Emails',
    importExcel: '📂 Import Excel',
    filterByDiscipline: 'All Discipline',
    discipline2: 'Discipline',
    cmInspectorLabel: 'CM Inspector',
  },
  ko: {
    projectTag: 'P25 TRITON · 하이데라바드',
    appTitle: '펀치리스트 매니저',
    dashboard: '대시보드', list: '목록', newSnag: '새 스냅',
    totalSnags: '전체', inProgress: '진행 중',
    critical: 'Critical', overdue: '기한초과',
    pendingAlert: (n) => `검사 대기 ${n}건`,
    pendingAlertSub: '업체 완료 → CM 검사 필요',
    view: '보기 →',
    statusBreakdown: '상태별 현황', recentSnags: '최근 등록',
    snagList: '스냅 목록', allStatus: '전체 상태',
    allPriority: '전체 우선순위', allContractor: '전체 업체',
    reset: '✕ 초기화', noSnags: '해당하는 스냅이 없습니다',
    addSnag: '새 스냅 등록',
    areaFloor: '구역 / 층 *', zoneGrid: 'Zone / Grid',
    discipline: '공종 *', contractor: '업체 *',
    priority: '우선순위', assignedTo: '담당자',
    targetDate: '목표일 *', photoRef: '사진 참조번호',
    description: '스냅 설명 *',
    descPlaceholder: '결함 내용을 상세히 기입...',
    correctiveAction: '시정 조치 사항',
    correctivePlaceholder: '필요한 시정 조치 내용...',
    remarks: '비고',
    submitSnag: '📋 스냅 등록',
    select: '선택...',
    backToList: '← 목록으로',
    snagNo: '스냅 #',
    statusChange: '상태 변경',
    roleLabel: (r) => `${r} ▾`,
    activityLog: '활동 이력',
    exportExcel: '📥 Excel 내보내기',
    exportAll: '전체 내보내기',
    emailTo: '받는 사람:', emailSubject: '제목:',
    toastRegistered: (id) => `스냅 #${id} 등록 완료!`,
    toastStatus: (s) => `상태 변경: ${s}`,
    toastExported: 'Excel 내보내기 완료!',
    targetDateLabel: '목표일', cmInspector: 'CM 검사관',
    photoRefLabel: '사진 참조', completedDate: '완료일',
    overdueTag: '기한초과', daysOverdue: (n) => `${n}일 초과`,
    emailNotice: '※ 이메일 알림: 설정에서 EmailJS 구성 필요',
    settings: '설정',
    settingsTitle: 'Settings / 설정',
    smtpNote: '이메일 알림은 EmailJS 설정이 필요합니다.',
    emailJsId: 'EmailJS Service ID',
    emailJsTemplate: 'EmailJS Template ID',
    emailJsKey: 'EmailJS Public Key',
    save: '저장',
    saved: '저장됨!',
    clearData: '🗑 전체 데이터 초기화',
    clearConfirm: '모든 스냅 데이터를 삭제하시겠습니까? 복구 불가합니다.',
    importJson: '📂 JSON 불러오기',
    exportJson: '📤 JSON 내보내기',
    vendorEmails: '업체 이메일 설정',
    importExcel: '📂 Excel 가져오기',
    filterByDiscipline: '전체 공종',
    discipline2: '공종',
    cmInspectorLabel: 'CM 검사관',
  },
}

// ─── CONSTANTS ────────────────────────────────────────────────────
const STORAGE_KEY = 'p25v5_snags'
const SETTINGS_KEY = 'p25v5_settings'

const DEFAULT_CONTRACTOR_EMAILS = {
  'ALUFIT': 'pm@alufit-project.com',
  'L&T': 'pm@lnt-project.com',
  'HEMASRI': 'site@hemasri-project.com',
  'MP Waterproofing': 'eng@mpwp-project.com',
  'JSSL': 'pm@jssl-project.com',
  'HÖRMANN': 'pm@hormann-project.com',
  'Client/Commercial': 'client@phoenixgroup.com',
  'All Contractors': 'contractors@jlcm.com',
  'Other': 'site@jlcm.com',
  'TBD': 'cm-team@jlcm.com',
}

const PRIORITIES = ['Critical', 'High', 'Medium', 'Low']
const STATUSES = ['Open','In Progress','On Hold','Pending Inspection','Approved','Completed','Closed','Rejected','Rework']
const DISCIPLINES = ['Civil / Finishing','Housekeeping','Painting','Waterproofing','Façade','Fireproofing','Finishing','Fire Door','EHS','MEP','Logistics','QA/QC']
const CONTRACTORS = ['ALUFIT','L&T','HEMASRI','MP Waterproofing','JSSL','HÖRMANN','Client/Commercial','All Contractors','Other','TBD']
const AREAS = ['B1 / Ramp','Basement','Typical Floor','B6','Façade','Upper Floors','CAF Area','Fire Door Area','Site-wide','Other']
const CM_INSPECTORS = ['CM Civil','CM Finishing','CM Façade','CM QA/QC','CM EHS','CM Lead','CM MEP']

const INIT_SNAGS = [
  { id:1, dateRaised:'2026-05-12', area:'B1 / Ramp', zone:'North Ramp', discipline:'Civil / Finishing', contractor:'HEMASRI', description:'VDF dummy joint sealant not completed before panel / subsequent works', priority:'High', assignedTo:'Site Engineer', targetDate:'2026-05-14', completedDate:'', status:'Open', cmInspector:'CM Civil', photoRef:'Photo-001', correctiveAction:'Complete joint cleaning, backer rod if required, sealant application, and submit photo evidence.', remarks:'Access may be restricted after panel installation.', activityLog:[{date:'2026-05-12',action:'Snag raised',by:'CM Team'}] },
  { id:2, dateRaised:'2026-05-12', area:'Basement', zone:'Lift Lobby', discipline:'Housekeeping', contractor:'L&T', description:'Desnagging and housekeeping not completed before follow-on works', priority:'Medium', assignedTo:'L&T Supervisor', targetDate:'2026-05-17', completedDate:'', status:'In Progress', cmInspector:'CM Finishing', photoRef:'Photo-002', correctiveAction:'Clear debris, remove obstructions, and release work front floor-wise.', remarks:'Coordinate with façade and MEP teams.', activityLog:[{date:'2026-05-12',action:'Snag raised',by:'CM Team'},{date:'2026-05-13',action:'Status → In Progress',by:'L&T Supervisor'}] },
  { id:3, dateRaised:'2026-05-13', area:'Typical Floor', zone:'Shaft Area', discipline:'Painting', contractor:'HEMASRI', description:'Shaft paint touch-up / incomplete coverage observed', priority:'Medium', assignedTo:'Painting Supervisor', targetDate:'2026-05-20', completedDate:'', status:'Open', cmInspector:'CM Finishing', photoRef:'Photo-003', correctiveAction:'Complete surface preparation, primer/paint touch-up, and request inspection.', remarks:'Check ventilation and access safety.', activityLog:[{date:'2026-05-13',action:'Snag raised',by:'CM Team'}] },
  { id:4, dateRaised:'2026-05-13', area:'B6', zone:'Waterproofing Area', discipline:'Waterproofing', contractor:'MP Waterproofing', description:'Waterproofing work front readiness and inspection sequence to be confirmed', priority:'High', assignedTo:'Waterproofing Engineer', targetDate:'2026-05-14', completedDate:'', status:'Open', cmInspector:'CM Civil', photoRef:'Photo-004', correctiveAction:'Confirm substrate readiness, material approval, MIR/WIR sequence, and submit inspection request.', remarks:'Critical before monsoon exposure.', activityLog:[{date:'2026-05-13',action:'Snag raised',by:'CM Team'}] },
  { id:5, dateRaised:'2026-05-14', area:'Façade', zone:'East Elevation', discipline:'Façade', contractor:'ALUFIT', description:'Panel vertical shifting status not marked floor-wise / elevation-wise', priority:'High', assignedTo:'Façade Manager', targetDate:'2026-05-20', completedDate:'', status:'In Progress', cmInspector:'CM Façade', photoRef:'Tracker-001', correctiveAction:'Submit floor-wise and elevation-wise delivery, vertical shifting, and installation status.', remarks:'Required for recovery monitoring.', activityLog:[{date:'2026-05-14',action:'Snag raised',by:'CM Team'},{date:'2026-05-15',action:'Status → In Progress',by:'ALUFIT PM'}] },
  { id:6, dateRaised:'2026-05-14', area:'Upper Floors', zone:'Fireproofing Zone', discipline:'Fireproofing', contractor:'JSSL', description:'Fireproofing spray work area requires inspection record and photo evidence', priority:'Medium', assignedTo:'JSSL Supervisor', targetDate:'2026-05-25', completedDate:'', status:'Open', cmInspector:'CM QA/QC', photoRef:'Photo-005', correctiveAction:'Submit WIR, thickness check record if applicable, and before/after photos.', remarks:'Coordinate access and protection.', activityLog:[{date:'2026-05-14',action:'Snag raised',by:'CM Team'}] },
  { id:7, dateRaised:'2026-05-15', area:'CAF Area', zone:'Staircase / Lift Lobby', discipline:'Finishing', contractor:'Client/Commercial', description:'CAF work front risk due to contractor withdrawal / demobilization', priority:'Critical', assignedTo:'Client / Commercial', targetDate:'2026-05-22', completedDate:'', status:'On Hold', cmInspector:'CM Lead', photoRef:'Issue-CAF-001', correctiveAction:'Confirm replacement strategy, scope split, and immediate mobilization plan.', remarks:'Potential Fire NOC / OC impact.', activityLog:[{date:'2026-05-15',action:'Snag raised',by:'CM Team'}] },
  { id:8, dateRaised:'2026-05-15', area:'Façade', zone:'Material Staging', discipline:'Logistics', contractor:'ALUFIT', description:'Bending panel / fastener delivery status requires recovery confirmation', priority:'Critical', assignedTo:'ALUFIT PM', targetDate:'2026-05-13', completedDate:'', status:'Open', cmInspector:'CM Façade', photoRef:'Delivery-001', correctiveAction:'Confirm production, delivery, staging, and installation recovery action with measurable weekly targets.', remarks:'Critical path risk.', activityLog:[{date:'2026-05-15',action:'Snag raised',by:'CM Team'}] },
]

const PCOLOR = {
  Critical: { bg:'#fee2e2', text:'#991b1b', dot:'#ef4444' },
  High:     { bg:'#fef3c7', text:'#92400e', dot:'#f59e0b' },
  Medium:   { bg:'#dbeafe', text:'#1e40af', dot:'#3b82f6' },
  Low:      { bg:'#dcfce7', text:'#166534', dot:'#22c55e' },
}

const SCOLOR = {
  'Open':               { bg:'#fee2e2', text:'#991b1b' },
  'In Progress':        { bg:'#dbeafe', text:'#1e40af' },
  'On Hold':            { bg:'#ede9fe', text:'#5b21b6' },
  'Pending Inspection': { bg:'#fef3c7', text:'#92400e' },
  'Approved':           { bg:'#dcfce7', text:'#166534' },
  'Completed':          { bg:'#dcfce7', text:'#166534' },
  'Closed':             { bg:'#e5e7eb', text:'#374151' },
  'Rejected':           { bg:'#fce7f3', text:'#9d174d' },
  'Rework':             { bg:'#fce7f3', text:'#9d174d' },
}

const TRANSITIONS = {
  CM: {
    'Open':               ['In Progress','On Hold'],
    'In Progress':        ['On Hold','Pending Inspection'],
    'On Hold':            ['Open','In Progress'],
    'Pending Inspection': ['Approved','Closed','Rejected','Rework'],
    'Approved':           ['Closed'],
    'Rejected':           ['Open'],
    'Rework':             ['Open'],
  },
  Contractor: {
    'Open':        ['In Progress'],
    'In Progress': ['Pending Inspection'],
    'Rework':      ['In Progress'],
  },
  Approver: {
    'Pending Inspection': ['Approved','Closed','Rejected','Rework'],
    'Approved':           ['Closed'],
  },
}

function daysOverdue(targetDate, completedDate, status) {
  if (['Closed','Completed','Approved'].includes(status)) return 0
  if (!targetDate) return 0
  const diff = Math.floor((new Date() - new Date(targetDate)) / 86400000)
  return diff > 0 ? diff : 0
}

// ─── SUB-COMPONENTS ───────────────────────────────────────────────
function Badge({ label, bg, text, sm }) {
  return <span style={{ background:bg, color:text, padding:sm?'2px 8px':'4px 12px', borderRadius:20, fontSize:sm?10:11, fontWeight:700, whiteSpace:'nowrap' }}>{label}</span>
}

function InfoRow({ label, value }) {
  if (!value) return null
  return (
    <div style={{ display:'flex', justifyContent:'space-between', padding:'7px 0', borderBottom:'1px solid #f3f4f6' }}>
      <span style={{ fontSize:12, color:'#9ca3af' }}>{label}</span>
      <span style={{ fontSize:13, fontWeight:600, color:'#374151', textAlign:'right', maxWidth:'62%' }}>{value}</span>
    </div>
  )
}

function SnagCard({ snag, onClick, t }) {
  const pc = PCOLOR[snag.priority] || {}
  const sc = SCOLOR[snag.status] || {}
  const od = daysOverdue(snag.targetDate, snag.completedDate, snag.status)
  return (
    <div onClick={onClick} style={{ background:'white', borderRadius:12, padding:14, marginBottom:10, cursor:'pointer', borderLeft:`4px solid ${pc.dot||'#9ca3af'}`, boxShadow:'0 1px 4px rgba(0,0,0,0.07)' }}>
      <div style={{ display:'flex', justifyContent:'space-between', marginBottom:5, alignItems:'center' }}>
        <span style={{ fontSize:11, color:'#9ca3af', fontWeight:600 }}>#{snag.id} · {snag.dateRaised}</span>
        <Badge label={snag.status} bg={sc.bg} text={sc.text} sm />
      </div>
      <div style={{ fontSize:14, fontWeight:600, color:'#1a2744', marginBottom:5, lineHeight:1.4 }}>
        {snag.description.slice(0,85)}{snag.description.length>85?'...':''}
      </div>
      <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center' }}>
        <span style={{ fontSize:11, color:'#6b7280' }}>{snag.area} · {snag.contractor}</span>
        <div style={{ display:'flex', gap:5, alignItems:'center' }}>
          {od > 0 && <span style={{ fontSize:10, background:'#fee2e2', color:'#991b1b', padding:'2px 6px', borderRadius:4, fontWeight:700 }}>{t.daysOverdue(od)}</span>}
          <Badge label={snag.priority} bg={pc.bg} text={pc.text} sm />
        </div>
      </div>
    </div>
  )
}

// ─── SETTINGS PANEL ───────────────────────────────────────────────
function SettingsPanel({ t, lang, onClose }) {
  const stored = JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}')
  const [emails, setEmails] = useState(stored.contractorEmails || DEFAULT_CONTRACTOR_EMAILS)
  const [saved, setSaved] = useState(false)

  function handleSave() {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify({ contractorEmails: emails }))
    setSaved(true)
    setTimeout(() => setSaved(false), 2000)
  }

  function handleClearData() {
    if (window.confirm(t.clearConfirm)) {
      localStorage.removeItem(STORAGE_KEY)
      window.location.reload()
    }
  }

  return (
    <div style={{ position:'fixed', inset:0, background:'rgba(0,0,0,0.55)', zIndex:300, display:'flex', alignItems:'flex-end' }}>
      <div style={{ background:'white', borderRadius:'20px 20px 0 0', padding:20, width:'100%', maxHeight:'85vh', overflow:'auto', boxSizing:'border-box' }}>
        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:16 }}>
          <div style={{ fontWeight:700, fontSize:16, color:'#1a2744' }}>{t.settingsTitle}</div>
          <button onClick={onClose} style={{ background:'#f3f4f6', border:'none', borderRadius:8, padding:'6px 12px', cursor:'pointer', fontWeight:700 }}>✕</button>
        </div>

        <div style={{ fontWeight:700, fontSize:13, color:'#374151', marginBottom:10 }}>{t.vendorEmails}</div>
        {CONTRACTORS.map(c => (
          <div key={c} style={{ marginBottom:8 }}>
            <label style={{ display:'block', fontSize:11, fontWeight:600, color:'#9ca3af', marginBottom:3 }}>{c}</label>
            <input value={emails[c]||''} onChange={e => setEmails(p => ({ ...p, [c]: e.target.value }))} style={{ width:'100%', padding:'8px 10px', borderRadius:7, border:'1px solid #d1d5db', fontSize:13, boxSizing:'border-box' }} />
          </div>
        ))}

        <button onClick={handleSave} style={{ width:'100%', padding:12, background:'#1a2744', color:'white', border:'none', borderRadius:10, fontWeight:700, fontSize:14, cursor:'pointer', marginTop:12 }}>
          {saved ? t.saved : t.save}
        </button>
        <button onClick={handleClearData} style={{ width:'100%', padding:12, background:'#fee2e2', color:'#991b1b', border:'none', borderRadius:10, fontWeight:700, fontSize:13, cursor:'pointer', marginTop:8 }}>
          {t.clearData}
        </button>
      </div>
    </div>
  )
}

// ─── MAIN APP ─────────────────────────────────────────────────────
export default function App() {
  const [lang, setLang] = useState('en')
  const t = T[lang]

  const [snags, setSnags] = useState([])
  const [view, setView] = useState('dashboard')
  const [sel, setSel] = useState(null)
  const [role, setRole] = useState('CM')
  const [filter, setFilter] = useState({ status:'', priority:'', contractor:'', discipline:'' })
  const [toast, setToast] = useState(null)
  const [roleMenu, setRoleMenu] = useState(false)
  const [showSettings, setShowSettings] = useState(false)
  const [form, setForm] = useState({
    area:'', zone:'', discipline:'', contractor:'', description:'',
    priority:'High', assignedTo:'', targetDate:'', remarks:'', correctiveAction:'', photoRef:'', cmInspector:'',
  })

  useEffect(() => {
    const saved = loadSnags(STORAGE_KEY)
    setSnags(saved || INIT_SNAGS)
    if (!saved) saveSnags(STORAGE_KEY, INIT_SNAGS)

    const unsub = onSnagsUpdated((data) => setSnags(data))
    return unsub
  }, [])

  const save = useCallback((data) => {
    setSnags(data)
    saveSnags(STORAGE_KEY, data)
  }, [])

  function showToast(msg, type = 'ok') {
    setToast({ msg, type })
    setTimeout(() => setToast(null), 3000)
  }

  function handleStatusChange(snag, ns) {
    const now = new Date().toISOString().split('T')[0]
    const updated = {
      ...snag,
      status: ns,
      completedDate: ['Completed','Closed','Approved'].includes(ns) ? now : snag.completedDate,
      activityLog: [...(snag.activityLog||[]), { date:now, action:`Status: ${snag.status} → ${ns}`, by:role }],
    }
    save(snags.map(s => s.id === snag.id ? updated : s))
    setSel(updated)
    showToast(t.toastStatus(ns))
  }

  function handleAdd(e) {
    e.preventDefault()
    const newId = Math.max(...snags.map(s => s.id), 0) + 1
    const now = new Date().toISOString().split('T')[0]
    const snag = { id:newId, dateRaised:now, ...form, status:'Open', completedDate:'', activityLog:[{ date:now, action:'Snag raised', by:role }] }
    save([...snags, snag])
    showToast(t.toastRegistered(newId))
    setForm({ area:'', zone:'', discipline:'', contractor:'', description:'', priority:'High', assignedTo:'', targetDate:'', remarks:'', correctiveAction:'', photoRef:'', cmInspector:'' })
    setView('list')
  }

  function handleExport() {
    exportToExcel(snags, lang)
    showToast(t.toastExported)
  }

  function handleExportJson() {
    const blob = new Blob([JSON.stringify(snags, null, 2)], { type:'application/json' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = `P25_PunchList_${new Date().toISOString().split('T')[0]}.json`
    a.click()
    URL.revokeObjectURL(url)
  }

  function handleImportJson(e) {
    const file = e.target.files[0]
    if (!file) return
    const reader = new FileReader()
    reader.onload = (ev) => {
      try {
        const data = JSON.parse(ev.target.result)
        if (Array.isArray(data)) { save(data); showToast('Imported!') }
      } catch { showToast('Import failed', 'error') }
    }
    reader.readAsText(file)
    e.target.value = ''
  }

  const today = new Date()
  const stats = {
    total:   snags.length,
    open:    snags.filter(s => ['Open','In Progress','On Hold'].includes(s.status)).length,
    critical: snags.filter(s => s.priority === 'Critical').length,
    overdue:  snags.filter(s => daysOverdue(s.targetDate, s.completedDate, s.status) > 0).length,
    pending:  snags.filter(s => s.status === 'Pending Inspection').length,
  }

  const filtered = snags
    .filter(s => {
      if (filter.status     && s.status     !== filter.status)     return false
      if (filter.priority   && s.priority   !== filter.priority)   return false
      if (filter.contractor && s.contractor !== filter.contractor) return false
      if (filter.discipline && s.discipline !== filter.discipline) return false
      return true
    })
    .sort((a, b) => ({ Critical:0, High:1, Medium:2, Low:3 }[a.priority] - { Critical:0, High:1, Medium:2, Low:3 }[b.priority]))

  const avail = sel ? (TRANSITIONS[role]?.[sel.status] || []) : []

  return (
    <div style={{ fontFamily:"'Segoe UI',system-ui,sans-serif", maxWidth:480, margin:'0 auto', minHeight:'100vh', background:'#f1f5f9', paddingBottom:72 }} onClick={() => setRoleMenu(false)}>

      {/* HEADER */}
      <div style={{ background:'#1a2744', color:'white', padding:'12px 16px', position:'sticky', top:0, zIndex:100, boxShadow:'0 2px 12px rgba(0,0,0,0.4)' }}>
        <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center' }}>
          <div>
            <div style={{ fontSize:10, color:'#f59e0b', fontWeight:700, letterSpacing:1.5 }}>{t.projectTag}</div>
            <div style={{ fontSize:16, fontWeight:700 }}>{t.appTitle}</div>
          </div>
          <div style={{ display:'flex', gap:7, alignItems:'center' }}>
            <button onClick={() => setLang(l => l==='en'?'ko':'en')} style={{ background:'rgba(255,255,255,0.12)', border:'1px solid rgba(255,255,255,0.3)', color:'white', padding:'5px 10px', borderRadius:16, fontSize:11, fontWeight:700, cursor:'pointer' }}>
              {lang==='en'?'한국어':'English'}
            </button>
            <button onClick={() => setShowSettings(true)} style={{ background:'rgba(255,255,255,0.12)', border:'1px solid rgba(255,255,255,0.3)', color:'white', padding:'5px 10px', borderRadius:16, fontSize:11, fontWeight:700, cursor:'pointer' }}>⚙</button>
            <div style={{ position:'relative' }} onClick={e => e.stopPropagation()}>
              <button onClick={() => setRoleMenu(!roleMenu)} style={{ background:'rgba(245,158,11,0.2)', border:'1.5px solid #f59e0b', color:'#f59e0b', padding:'5px 12px', borderRadius:16, fontSize:11, fontWeight:700, cursor:'pointer' }}>
                {t.roleLabel(role)}
              </button>
              {roleMenu && (
                <div style={{ position:'absolute', right:0, top:34, background:'#1e3a5f', borderRadius:10, overflow:'hidden', zIndex:200, boxShadow:'0 4px 20px rgba(0,0,0,0.5)', minWidth:120 }}>
                  {['CM','Contractor','Approver'].map(r => (
                    <button key={r} onClick={() => { setRole(r); setRoleMenu(false) }} style={{ display:'block', width:'100%', padding:'11px 14px', background:role===r?'rgba(245,158,11,0.25)':'transparent', color:'white', border:'none', fontSize:13, fontWeight:600, cursor:'pointer', textAlign:'left' }}>{r}</button>
                  ))}
                </div>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* TOAST */}
      {toast && (
        <div style={{ position:'fixed', top:68, left:'50%', transform:'translateX(-50%)', background:toast.type==='error'?'#ef4444':'#10b981', color:'white', padding:'10px 20px', borderRadius:24, fontSize:13, fontWeight:700, zIndex:400, boxShadow:'0 4px 16px rgba(0,0,0,0.25)', whiteSpace:'nowrap' }}>
          {toast.msg}
        </div>
      )}

      {/* SETTINGS */}
      {showSettings && <SettingsPanel t={t} lang={lang} onClose={() => setShowSettings(false)} />}

      {/* ══ DASHBOARD ══ */}
      {view==='dashboard' && (
        <div style={{ padding:16 }}>
          <div style={{ fontSize:12, color:'#9ca3af', marginBottom:2 }}>
            {today.toLocaleDateString(lang==='ko'?'ko-KR':'en-US', { weekday:'long', year:'numeric', month:'long', day:'numeric' })}
          </div>
          <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:14 }}>
            <div style={{ fontSize:20, fontWeight:700, color:'#1a2744' }}>{t.dashboard}</div>
            <button onClick={handleExport} style={{ background:'#1a2744', color:'#f59e0b', border:'none', padding:'8px 14px', borderRadius:9, fontSize:12, fontWeight:700, cursor:'pointer' }}>
              {t.exportExcel}
            </button>
          </div>

          <div style={{ display:'grid', gridTemplateColumns:'1fr 1fr', gap:10, marginBottom:14 }}>
            {[
              { label:t.totalSnags,  val:stats.total,   bg:'#e8ecf4', tc:'#1a2744' },
              { label:t.inProgress,  val:stats.open,    bg:'#dbeafe', tc:'#1e40af' },
              { label:t.critical,    val:stats.critical,bg:'#fee2e2', tc:'#991b1b' },
              { label:t.overdue,     val:stats.overdue, bg:'#fef3c7', tc:'#92400e' },
            ].map(item => (
              <div key={item.label} style={{ background:item.bg, borderRadius:12, padding:'14px 16px' }}>
                <div style={{ fontSize:11, fontWeight:700, color:item.tc, opacity:0.75, marginBottom:4 }}>{item.label}</div>
                <div style={{ fontSize:34, fontWeight:800, color:item.tc }}>{item.val}</div>
              </div>
            ))}
          </div>

          {stats.pending > 0 && (
            <div style={{ background:'#fff7ed', border:'1.5px solid #f59e0b', borderRadius:12, padding:'12px 14px', marginBottom:14, display:'flex', gap:12, alignItems:'center' }}>
              <div style={{ fontSize:24 }}>🔔</div>
              <div style={{ flex:1 }}>
                <div style={{ fontWeight:700, color:'#92400e', fontSize:14 }}>{t.pendingAlert(stats.pending)}</div>
                <div style={{ fontSize:12, color:'#b45309' }}>{t.pendingAlertSub}</div>
              </div>
              <button onClick={() => { setFilter(f=>({...f,status:'Pending Inspection'})); setView('list') }} style={{ background:'#f59e0b', color:'#1a2744', border:'none', padding:'7px 12px', borderRadius:8, fontSize:12, fontWeight:700, cursor:'pointer', whiteSpace:'nowrap' }}>
                {t.view}
              </button>
            </div>
          )}

          <div style={{ background:'white', borderRadius:12, padding:16, marginBottom:14 }}>
            <div style={{ fontWeight:700, color:'#1a2744', fontSize:14, marginBottom:12 }}>{t.statusBreakdown}</div>
            {Object.entries(snags.reduce((a,s)=>({...a,[s.status]:(a[s.status]||0)+1}),{})).map(([st,cnt]) => (
              <div key={st} style={{ marginBottom:8 }}>
                <div style={{ display:'flex', justifyContent:'space-between', marginBottom:3 }}>
                  <span style={{ fontSize:12, color:'#374151' }}>{st}</span>
                  <span style={{ fontSize:12, fontWeight:700, color:(SCOLOR[st]||{}).text||'#374151' }}>{cnt}</span>
                </div>
                <div style={{ background:'#f3f4f6', borderRadius:4, height:5 }}>
                  <div style={{ background:(SCOLOR[st]||{}).text||'#6b7280', height:5, borderRadius:4, width:`${Math.round((cnt/Math.max(stats.total,1))*100)}%` }} />
                </div>
              </div>
            ))}
          </div>

          <div style={{ fontWeight:700, color:'#1a2744', fontSize:14, marginBottom:10 }}>{t.recentSnags}</div>
          {[...snags].reverse().slice(0,5).map(s => <SnagCard key={s.id} snag={s} t={t} onClick={() => { setSel(s); setView('detail') }} />)}
        </div>
      )}

      {/* ══ LIST ══ */}
      {view==='list' && (
        <div>
          <div style={{ padding:'16px 16px 0' }}>
            <div style={{ display:'flex', justifyContent:'space-between', alignItems:'center', marginBottom:12 }}>
              <div style={{ fontSize:20, fontWeight:700, color:'#1a2744' }}>
                {t.snagList} <span style={{ fontSize:14, color:'#9ca3af', fontWeight:400 }}>({filtered.length})</span>
              </div>
              <button onClick={handleExport} style={{ background:'#1a2744', color:'#f59e0b', border:'none', padding:'7px 12px', borderRadius:8, fontSize:11, fontWeight:700, cursor:'pointer' }}>
                {t.exportExcel}
              </button>
            </div>
          </div>
          <div style={{ padding:'0 16px 10px', display:'flex', gap:7, overflowX:'auto' }}>
            {[
              { key:'status',     label:t.allStatus,     opts:STATUSES },
              { key:'priority',   label:t.allPriority,   opts:PRIORITIES },
              { key:'contractor', label:t.allContractor, opts:CONTRACTORS },
              { key:'discipline', label:t.filterByDiscipline, opts:DISCIPLINES },
            ].map(f => (
              <select key={f.key} value={filter[f.key]} onChange={e => setFilter(p=>({...p,[f.key]:e.target.value}))} style={{ fontSize:11, padding:'7px 9px', borderRadius:8, border:'1px solid #d1d5db', background:'white', flex:'none', cursor:'pointer' }}>
                <option value="">{f.label}</option>
                {f.opts.map(o => <option key={o}>{o}</option>)}
              </select>
            ))}
            {Object.values(filter).some(Boolean) && (
              <button onClick={() => setFilter({ status:'',priority:'',contractor:'',discipline:'' })} style={{ fontSize:11, padding:'7px 9px', borderRadius:8, border:'1px solid #ef4444', background:'#fee2e2', color:'#991b1b', cursor:'pointer', flex:'none', fontWeight:700 }}>
                {t.reset}
              </button>
            )}
          </div>
          <div style={{ padding:'0 16px' }}>
            {filtered.length===0
              ? <div style={{ textAlign:'center', padding:40, color:'#9ca3af' }}>{t.noSnags}</div>
              : filtered.map(s => <SnagCard key={s.id} snag={s} t={t} onClick={() => { setSel(s); setView('detail') }} />)
            }
          </div>
        </div>
      )}

      {/* ══ ADD ══ */}
      {view==='add' && (
        <div style={{ padding:16 }}>
          <div style={{ fontSize:20, fontWeight:700, color:'#1a2744', marginBottom:16 }}>{t.addSnag}</div>
          <form onSubmit={handleAdd}>
            {[
              { label:t.areaFloor,  key:'area',       type:'select', opts:AREAS,         req:true },
              { label:t.zoneGrid,   key:'zone',        type:'text',   req:false },
              { label:t.discipline, key:'discipline',  type:'select', opts:DISCIPLINES,   req:true },
              { label:t.contractor, key:'contractor',  type:'select', opts:CONTRACTORS,   req:true },
              { label:t.priority,   key:'priority',    type:'select', opts:PRIORITIES,    req:false },
              { label:t.assignedTo, key:'assignedTo',  type:'text',   req:false },
              { label:t.cmInspectorLabel, key:'cmInspector', type:'select', opts:CM_INSPECTORS, req:false },
              { label:t.targetDate, key:'targetDate',  type:'date',   req:true },
              { label:t.photoRef,   key:'photoRef',    type:'text',   req:false },
            ].map(f => (
              <div key={f.key} style={{ marginBottom:12 }}>
                <label style={{ display:'block', fontSize:12, fontWeight:700, color:'#374151', marginBottom:4 }}>{f.label}</label>
                {f.type==='select' ? (
                  <select required={f.req} value={form[f.key]} onChange={e => setForm(p=>({...p,[f.key]:e.target.value}))} style={{ width:'100%', padding:'10px 12px', borderRadius:8, border:'1px solid #d1d5db', fontSize:14, background:'white', boxSizing:'border-box' }}>
                    <option value="">{t.select}</option>
                    {f.opts.map(o => <option key={o}>{o}</option>)}
                  </select>
                ) : (
                  <input type={f.type} required={f.req} value={form[f.key]} onChange={e => setForm(p=>({...p,[f.key]:e.target.value}))} style={{ width:'100%', padding:'10px 12px', borderRadius:8, border:'1px solid #d1d5db', fontSize:14, boxSizing:'border-box' }} />
                )}
              </div>
            ))}

            <div style={{ marginBottom:12 }}>
              <label style={{ display:'block', fontSize:12, fontWeight:700, color:'#374151', marginBottom:4 }}>{t.description}</label>
              <textarea required value={form.description} onChange={e => setForm(p=>({...p,description:e.target.value}))} rows={3} placeholder={t.descPlaceholder} style={{ width:'100%', padding:'10px 12px', borderRadius:8, border:'1px solid #d1d5db', fontSize:14, resize:'vertical', boxSizing:'border-box' }} />
            </div>

            <div style={{ marginBottom:12 }}>
              <label style={{ display:'block', fontSize:12, fontWeight:700, color:'#374151', marginBottom:4 }}>{t.correctiveAction}</label>
              <textarea value={form.correctiveAction} onChange={e => setForm(p=>({...p,correctiveAction:e.target.value}))} rows={3} placeholder={t.correctivePlaceholder} style={{ width:'100%', padding:'10px 12px', borderRadius:8, border:'1px solid #d1d5db', fontSize:14, resize:'vertical', boxSizing:'border-box' }} />
            </div>

            <div style={{ marginBottom:20 }}>
              <label style={{ display:'block', fontSize:12, fontWeight:700, color:'#374151', marginBottom:4 }}>{t.remarks}</label>
              <textarea value={form.remarks} onChange={e => setForm(p=>({...p,remarks:e.target.value}))} rows={2} style={{ width:'100%', padding:'10px 12px', borderRadius:8, border:'1px solid #d1d5db', fontSize:14, resize:'vertical', boxSizing:'border-box' }} />
            </div>

            <button type="submit" style={{ width:'100%', padding:14, background:'#1a2744', color:'white', border:'none', borderRadius:12, fontSize:15, fontWeight:700, cursor:'pointer' }}>
              {t.submitSnag}
            </button>
          </form>
        </div>
      )}

      {/* ══ DETAIL ══ */}
      {view==='detail' && sel && (
        <div style={{ padding:16 }}>
          <button onClick={() => setView('list')} style={{ background:'none', border:'none', color:'#6b7280', cursor:'pointer', padding:'0 0 12px', fontSize:14, fontWeight:600 }}>
            {t.backToList}
          </button>

          <div style={{ display:'flex', justifyContent:'space-between', alignItems:'flex-start', marginBottom:12 }}>
            <div>
              <div style={{ fontSize:11, color:'#9ca3af', fontWeight:600 }}>{t.snagNo}{sel.id} · {sel.dateRaised}</div>
              <div style={{ fontSize:17, fontWeight:700, color:'#1a2744', lineHeight:1.3, maxWidth:220 }}>{sel.area} / {sel.zone}</div>
            </div>
            <div style={{ display:'flex', flexDirection:'column', gap:5, alignItems:'flex-end' }}>
              <Badge label={sel.status} bg={(SCOLOR[sel.status]||{}).bg} text={(SCOLOR[sel.status]||{}).text} />
              <Badge label={sel.priority} bg={(PCOLOR[sel.priority]||{}).bg} text={(PCOLOR[sel.priority]||{}).text} />
              {daysOverdue(sel.targetDate, sel.completedDate, sel.status) > 0 && (
                <span style={{ fontSize:10, background:'#fee2e2', color:'#991b1b', padding:'2px 8px', borderRadius:12, fontWeight:700 }}>
                  {t.daysOverdue(daysOverdue(sel.targetDate, sel.completedDate, sel.status))}
                </span>
              )}
            </div>
          </div>

          {avail.length > 0 && (
            <div style={{ background:'#f0f7ff', borderRadius:12, padding:14, marginBottom:14 }}>
              <div style={{ fontSize:11, fontWeight:700, color:'#1e40af', marginBottom:8 }}>
                {t.statusChange} · {role}
              </div>
              <div style={{ display:'flex', flexWrap:'wrap', gap:8 }}>
                {avail.map(ns => (
                  <button key={ns} onClick={() => handleStatusChange(sel, ns)} style={{
                    padding:'9px 16px', borderRadius:8, border:'none', fontSize:13, fontWeight:700, cursor:'pointer',
                    background: ['Approved','Closed','Completed'].includes(ns)?'#10b981':['Rejected','Rework'].includes(ns)?'#ef4444':ns==='Pending Inspection'?'#f59e0b':'#3b82f6',
                    color:'white',
                  }}>→ {ns}</button>
                ))}
              </div>
            </div>
          )}

          <div style={{ background:'white', borderRadius:12, padding:16, marginBottom:10 }}>
            <InfoRow label={t.contractor}       value={sel.contractor} />
            <InfoRow label={t.discipline2}      value={sel.discipline} />
            <InfoRow label={t.assignedTo}       value={sel.assignedTo} />
            <InfoRow label={t.targetDateLabel}  value={sel.targetDate} />
            <InfoRow label={t.cmInspectorLabel} value={sel.cmInspector} />
            <InfoRow label={t.photoRefLabel}    value={sel.photoRef} />
            {sel.completedDate && <InfoRow label={t.completedDate} value={sel.completedDate} />}
          </div>

          <div style={{ background:'white', borderRadius:12, padding:16, marginBottom:10 }}>
            <div style={{ fontSize:11, fontWeight:700, color:'#9ca3af', marginBottom:6 }}>{t.description.replace(' *','')}</div>
            <div style={{ fontSize:14, color:'#374151', lineHeight:1.7 }}>{sel.description}</div>
          </div>

          {sel.correctiveAction && (
            <div style={{ background:'#f0fdf4', borderRadius:12, padding:16, marginBottom:10 }}>
              <div style={{ fontSize:11, fontWeight:700, color:'#166534', marginBottom:6 }}>{t.correctiveAction}</div>
              <div style={{ fontSize:14, color:'#374151', lineHeight:1.7 }}>{sel.correctiveAction}</div>
            </div>
          )}

          {sel.remarks && (
            <div style={{ background:'#fffbeb', borderRadius:12, padding:16, marginBottom:10 }}>
              <div style={{ fontSize:11, fontWeight:700, color:'#92400e', marginBottom:6 }}>{t.remarks}</div>
              <div style={{ fontSize:14, color:'#374151', lineHeight:1.7 }}>{sel.remarks}</div>
            </div>
          )}

          <div style={{ background:'white', borderRadius:12, padding:16, marginBottom:10 }}>
            <div style={{ fontSize:13, fontWeight:700, color:'#1a2744', marginBottom:12 }}>{t.activityLog}</div>
            {(sel.activityLog||[]).map((log,i) => (
              <div key={i} style={{ display:'flex', gap:10, marginBottom:8 }}>
                <div style={{ width:8, height:8, borderRadius:'50%', background:'#f59e0b', marginTop:5, flexShrink:0 }} />
                <div>
                  <div style={{ fontSize:13, color:'#374151' }}>{log.action}</div>
                  <div style={{ fontSize:11, color:'#9ca3af' }}>{log.date} · {log.by}</div>
                </div>
              </div>
            ))}
          </div>
        </div>
      )}

      {/* BOTTOM NAV */}
      <div style={{ position:'fixed', bottom:0, left:'50%', transform:'translateX(-50%)', width:'100%', maxWidth:480, background:'white', borderTop:'1px solid #e5e7eb', display:'flex', zIndex:100 }}>
        {[
          { id:'dashboard', icon:'📊', label:t.dashboard },
          { id:'list',      icon:'📋', label:t.list,    badge:stats.pending },
          { id:'add',       icon:'➕', label:t.newSnag },
        ].map(tab => (
          <button key={tab.id} onClick={() => setView(tab.id)} style={{ flex:1, padding:'11px 0', background:'none', border:'none', cursor:'pointer', display:'flex', flexDirection:'column', alignItems:'center', gap:3, borderTop:view===tab.id?'2.5px solid #f59e0b':'2.5px solid transparent' }}>
            <div style={{ position:'relative' }}>
              <span style={{ fontSize:20 }}>{tab.icon}</span>
              {tab.badge > 0 && (
                <span style={{ position:'absolute', top:-4, right:-10, background:'#ef4444', color:'white', borderRadius:'50%', width:16, height:16, fontSize:10, fontWeight:700, display:'flex', alignItems:'center', justifyContent:'center' }}>{tab.badge}</span>
              )}
            </div>
            <span style={{ fontSize:10, fontWeight:700, color:view===tab.id?'#f59e0b':'#9ca3af' }}>{tab.label}</span>
          </button>
        ))}
      </div>
    </div>
  )
}
