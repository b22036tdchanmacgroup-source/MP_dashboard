// ============================================================
// ★ 여기에 Apps Script 배포 URL을 입력하세요
// ============================================================
const DEFAULT_SCRIPT_URL = '';
// ============================================================

const today = new Date();
const todayStr = `${today.getFullYear()}.${String(today.getMonth()+1).padStart(2,'0')}.${String(today.getDate()).padStart(2,'0')}`;

let projects = [];
let routineData = {};
let teamMemos = [];
let ceoMemos = [];
let teamInfo = []; // 신규: 팀원 상세 정보
let isOfflineMode = false;

// 페이징 관련
let currentProjectPage = 0;
const PROJECTS_PER_PAGE = 4;

// 팀원 및 상태 필터
let activeFilterMember = null;
let activeFilterStatus = null;

// ── 공통 상태 산정 유틸 ────────────────────────────────────
function getComputedStatus(p) {
    const st = p['상태'] || p['status'] || '진행';
    const prg = parseInt(p['진행률'] || p['progress'] || 0);
    
    if (st === '보류') return '보류';
    if (st === '완료' || prg >= 100) return '완료';
    
    const startD = fmtDate(p['착수일'] || p['startDate'] || '');
    const planD  = fmtDate(p['계획종료일'] || p['planDate'] || '');
    let planPct  = null;
    if (startD && planD) {
        const s = new Date(startD), e = new Date(planD), now = new Date();
        if (!isNaN(s) && !isNaN(e) && e > s) {
            planPct = Math.min(100, Math.max(0, Math.round((now-s)/(e-s)*100)));
        }
    }
    if (planPct !== null && (planPct - prg) >= 10) return '지연';
    return '정상';
}

// ── 초기화 ───────────────────────────────────────────────
window.onload = () => {
    document.getElementById('currentDateDisplay').innerText = todayStr;

    const savedUrl = localStorage.getItem('sheets_script_url');
    if (savedUrl) {
        fetchFromSheets(savedUrl);
    } else if (DEFAULT_SCRIPT_URL) {
        fetchFromSheets(DEFAULT_SCRIPT_URL);
    } else {
        hideLoading();
        document.getElementById('urlSetupOverlay').style.display = 'flex';
    }
};

// ── 구글시트에서 데이터 가져오기 ─────────────────────────
async function fetchFromSheets(url) {
    const scriptUrl = url || localStorage.getItem('sheets_script_url') || DEFAULT_SCRIPT_URL;
    if (!scriptUrl) { useOfflineMode(); return; }

    showLoading('구글시트 데이터 불러오는 중...');
    try {
        const res = await fetch(scriptUrl);
        const json = await res.json();
        if (!json.ok) throw new Error(json.error);

        const d = json.data;
        projects   = d.projects  || [];
        routineData= d.routines  || {};
        teamInfo   = d.teamInfo  || []; // 데이터 수신
        teamMemos  = d.teamMemos || [];
        ceoMemos   = d.ceoMemos  || [];
        isOfflineMode = false;
        hideOfflineBanner();
        const now = new Date();
        const syncTime = `${now.getFullYear()}.${String(now.getMonth()+1).padStart(2,'0')}.${String(now.getDate()).padStart(2,'0')} ${String(now.getHours()).padStart(2,'0')}:${String(now.getMinutes()).padStart(2,'0')}`;
        setStatus(`마지막 동기화: ${syncTime}`);
    } catch(e) {
        console.warn('시트 로드 실패, 오프라인 모드로 전환:', e);
        loadFromLocalStorage();
        showOfflineBanner();
        setStatus('⚠️ 오프라인 (로컬 데이터)');
    }
    hideLoading();
    render();
}

// ── 오프라인(localStorage) ────────────────────────────────
function loadFromLocalStorage() {
    isOfflineMode = true;
    projects    = JSON.parse(localStorage.getItem('biz_p_v11') || '[]');
    routineData = JSON.parse(localStorage.getItem('routine_v11') || '{}');
    teamMemos   = JSON.parse(localStorage.getItem('team_m_v11') || '[]');
    ceoMemos    = JSON.parse(localStorage.getItem('ceo_m_v11')  || '[]');

    // KPI 데모값 (구글시트 미연결 시 샘플 표시)
    if (!routineData['KPI_매출']) {
        routineData['KPI_매출']      = '128억';
        routineData['KPI_매출증감']  = '+6.2% YoY';
        routineData['KPI_점유율']    = '24.8%';
        routineData['KPI_점유율증감']= '+1.4%p';
    }
}

function useOfflineMode() {
    document.getElementById('urlSetupOverlay').style.display = 'none';
    loadFromLocalStorage();
    showOfflineBanner();
    hideLoading();
    setStatus('오프라인 모드');
    render();
}

// ── URL 설정 ─────────────────────────────────────────────
function openUrlSetup() {
    const input = document.getElementById('scriptUrlInput');
    const saved = localStorage.getItem('sheets_script_url') || '';
    if (saved && input) input.value = saved;
    document.getElementById('urlSetupOverlay').style.display = 'flex';
}
function closeUrlSetup() { document.getElementById('urlSetupOverlay').style.display = 'none'; }

function saveScriptUrl() {
    const url = document.getElementById('scriptUrlInput').value.trim();
    if (!url || !url.startsWith('https://script.google.com')) {
        alert('올바른 Apps Script URL을 입력해주세요.\nhttps://script.google.com/macros/s/... 형태여야 합니다.');
        return;
    }
    localStorage.setItem('sheets_script_url', url);
    document.getElementById('urlSetupOverlay').style.display = 'none';
    fetchFromSheets(url);
}

// ── 렌더링 ───────────────────────────────────────────────
function render() {
    renderProjects();
    renderTeamMemos();
    renderCeoMemos();
    renderRoutines();
    renderReportTitle();
    renderQuickLinks();
}


function renderReportTitle() {
    const title = routineData['보고제목'] || '';
    const el = document.getElementById('reportTitleDisplay');
    if (title.trim()) {
        el.textContent = title;
        el.classList.remove('hidden');
    } else {
        el.classList.add('hidden');
    }
}

// ── 헤더 바로가기 아이콘 렌더링 ──────────────────────────────
// 구글시트 [상시업무] 탭에 아래 키-값으로 입력하세요:
//   바로가기1_아이콘  |  🏛
//   바로가기1_제목    |  나라장터
//   바로가기1_URL     |  https://www.g2b.go.kr
//   바로가기2_아이콘  |  📊
//   바로가기2_제목    |  ERP
//   바로가기2_URL     |  https://erp.company.com
//   (바로가기1 ~ 바로가기5 까지 지원)
function renderQuickLinks() {
    const container = document.getElementById('quickLinks');
    if (!container) return;
    container.innerHTML = '';
    for (let i = 1; i <= 5; i++) {
        const icon  = (routineData['바로가기' + i + '_아이콘'] || '').trim();
        const label = (routineData['바로가기' + i + '_제목']   || '').trim();
        const url   = (routineData['바로가기' + i + '_URL']    || '').trim();
        if (!label) continue;
        const btn = document.createElement('button');
        btn.className = 'flex items-center gap-1.5 px-3 py-2 rounded-lg bg-white/10 hover:bg-white/20 border border-white/15 transition-all text-xs font-bold';
        btn.title = label + (url ? ' → ' + url : '');
        btn.innerHTML = '<span>' + (icon || '🔗') + '</span>'
                      + '<span class="opacity-90 whitespace-nowrap">' + label + '</span>';
        if (url) {
            btn.addEventListener('click', function(e) {
                e.stopPropagation();
                window.open(url, '_blank', 'popup=yes,width=1200,height=800,menubar=no,toolbar=no,location=yes,status=no');
            });
        }
        container.appendChild(btn);
    }
}

// ── 날짜 포맷 헬퍼 ───────────────────────────────────────
function fmtDate(val) {
    if (!val) return '';
    // Date 객체 또는 날짜 문자열 모두 처리
    const d = new Date(val);
    if (isNaN(d.getTime())) return String(val); // 파싱 불가면 원본 반환
    const y = d.getFullYear();
    const m = String(d.getMonth()+1).padStart(2,'0');
    const day = String(d.getDate()).padStart(2,'0');
    return `${y}.${m}.${day}`;
}

// ── 프로젝트 아이디(문자열)로 과업 인덱스 찾기 (3단계 매칭) ────────────────
function resolveProjectIndex(idTrimmed) {
    if (!idTrimmed) return -1;
    // 1. 완전 일치
    let idx = projects.findIndex(p => (p['업무명'] || p['title'] || '').trim() === idTrimmed);
    if (idx < 0) {
        // 2. 포함 일치
        idx = projects.findIndex(p => {
            const t = (p['업무명'] || p['title'] || '').trim();
            return t.includes(idTrimmed) || idTrimmed.includes(t);
        });
    }
    if (idx < 0) {
        // 3. 키워드 분리 매칭
        const keywords = idTrimmed.split(/[\s\-\/\(\)·,]+/).filter(w => w.length >= 2);
        if (keywords.length > 0) {
            let bestScore = 0;
            projects.forEach((p, i) => {
                const t = (p['업무명'] || p['title'] || '').trim();
                const score = keywords.filter(kw => t.includes(kw)).length;
                if (score > bestScore) { bestScore = score; idx = i; }
            });
            if (bestScore === 0) idx = -1;
        }
    }
    return idx;
}

// ── 관련된 지시/협의사항 건수 카운트 ──────────────────────
// 원본 배열(projects)의 인덱스 기준 매칭, focusProject와 완벽히 동일한 조건
function getAssociatedMemoCount(origProjIdx) {
    return ceoMemos.filter(m => {
        if (!m.projectId) return false;
        const resolvedIdx = resolveProjectIndex(m.projectId.trim());
        return resolvedIdx === origProjIdx;
    }).length;
}

// ── 프로젝트 컴팩트 요약 리스트 렌더링 ────────────────────────
function renderProjectSummaryList(projs) {
    const listWrap = document.getElementById('projectSummaryListWrap');
    const list = document.getElementById('projectSummaryList');
    if (!list || !listWrap) return;
    
    list.innerHTML = '';
    
    if (projs.length === 0) {
        listWrap.style.display = 'none';
        return;
    }
    
    listWrap.style.display = 'block';

    projs.forEach((p) => {
        const title    = p['업무명'] || p['title'] || '(제목 없음)';
        const status   = p['상태'] || p['status'] || '진행';
        const imp      = p['우선순위'] || p['importance'] || '중';
        const progress = parseInt(p['진행률'] || p['progress'] || 0);
        
        let statusCls = '';
        if (status === '완료') statusCls = 'bg-emerald-100 text-emerald-700';
        else if (status === '보류') statusCls = 'bg-slate-100 text-slate-500';
        else if (status === '지연') statusCls = 'bg-red-100 text-red-700';
        else statusCls = 'bg-blue-100 text-blue-700';

        const barColor = progress >= 100 ? '#10b981' : (status === '지연' ? '#ef4444' : '#1E3A8A');
        
        const origIdx = projects.indexOf(p);
        const memoCount = getAssociatedMemoCount(origIdx);

        const row = document.createElement('div');
        row.className = 'compact-card flex flex-col justify-between bg-white rounded-lg px-3 py-2.5 shadow-sm border border-slate-100 hover:border-[#1E3A8A]/30 transition-all cursor-pointer group min-h-[60px] gap-2';
        
        if (origIdx === window.activeProjectIdx) {
            row.classList.add('active-card');
        }

        row.onclick = () => {
            window.activeProjectIdx = origIdx;
            document.querySelectorAll('.compact-card, .project-card').forEach(n => n.classList.remove('active-card'));
            row.classList.add('active-card');
            isDualMode ? renderExpandedDetail(origIdx) : openDetailPopup(origIdx);
        };

        row.innerHTML = `
            <div class="flex items-start justify-between gap-1">
                <div class="flex items-center gap-1.5 min-w-0">
                    <span class="badge-${imp} text-[9px] font-bold px-1 py-0.5 rounded-sm shrink-0 leading-none">${imp}</span>
                    <span class="px-1.5 py-0.5 rounded-sm text-[9px] font-black uppercase tracking-wider ${statusCls} shrink-0 leading-none">${status}</span>
                    <span class="text-xs font-bold text-slate-700 truncate group-hover:text-[#1E3A8A] transition-colors" title="${title}">${title}</span>
                </div>
                ${memoCount > 0 ? `<div class="shrink-0 flex items-center justify-center bg-rose-50 border border-rose-200 px-1.5 py-0.5 rounded-full" title="관련 지시/협의사항">
                    <span class="text-[9px]">💬</span><span class="text-[10px] font-black text-rose-600 ml-0.5">${memoCount}</span>
                </div>` : ''}
            </div>
            <div class="flex items-center gap-2">
                <div class="flex-1 h-1 bg-slate-100 rounded-full overflow-hidden">
                    <div class="h-full rounded-full transition-all duration-500" style="width:${progress}%; background:${barColor}"></div>
                </div>
                <span class="text-[10px] font-black text-slate-800 w-6 text-right shrink-0">${progress}%</span>
            </div>
        `;
        list.appendChild(row);
    });
}

function toggleSummaryList() {
    const listWrap = document.getElementById('projectSummaryListWrap');
    const btn = document.getElementById('summaryToggleBtn');
    if (!listWrap || !btn) return;
    
    if (listWrap.style.display === 'none') {
        listWrap.style.display = 'block';
        btn.innerHTML = '컴팩트 뷰 숨기기 <span>▲</span>';
    } else {
        listWrap.style.display = 'none';
        btn.innerHTML = '컴팩트 뷰 펼치기 <span>▼</span>';
    }
}

function toggleStatusFilter(status) {
    activeFilterStatus = (activeFilterStatus === status) ? null : status;
    updateStatusFilterUI();
    currentProjectPage = 0;
    renderProjects();
}

function updateStatusFilterUI() {
    ['정상','지연','보류','완료'].forEach(st => {
        const btn = document.getElementById('btn-filter-' + st);
        if (btn) {
            if (activeFilterStatus) {
                if (activeFilterStatus === st) {
                    btn.classList.add('ring-2', 'ring-offset-1', 'ring-slate-400');
                    btn.style.opacity = '1';
                } else {
                    btn.classList.remove('ring-2', 'ring-offset-1', 'ring-slate-400');
                    btn.style.opacity = '0.3';
                }
            } else {
                btn.classList.remove('ring-2', 'ring-offset-1', 'ring-slate-400');
                btn.style.opacity = '1';
            }
        }
    });
}

function renderProjects() {
    const grid = document.getElementById('projectGrid');
    grid.innerHTML = '';

    // 팀원 및 상태 필터 적용
    const filteredProjects = projects.filter(p => {
        let matchMember = true;
        if (activeFilterMember) {
            const pm = (p['PM'] || '').trim();
            const assignee = (p['담당자'] || '').trim();
            matchMember = (pm === activeFilterMember || assignee === activeFilterMember);
        }
        let matchStatus = true;
        if (activeFilterStatus) {
            matchStatus = (getComputedStatus(p) === activeFilterStatus);
        }
        return matchMember && matchStatus;
    });

    if (filteredProjects.length === 0) {
        grid.innerHTML = `<div class="col-span-2 flex flex-col items-center justify-center py-20 text-slate-300">
            <p class="text-5xl mb-3">📋</p><p class="font-bold text-sm">해당 조건의 과업이 없습니다.</p></div>`;
        updateSummary([]);
        return;
    }

    updateSummary(projects); // 전체 요약은 필터 무관하게 표시
    renderProjectSummaryList(filteredProjects); // 요약 컴팩트 리스트 렌더링

    filteredProjects.forEach((p) => {
        const progress = parseInt(p['진행률'] || p['progress'] || 0);
        const status   = p['상태'] || p['status'] || '진행';
        const imp      = p['우선순위'] || p['importance'] || '중';
        const title    = p['업무명'] || p['title'] || '(제목 없음)';
        const desc     = p['과업목표'] || p['goals'] || p['내용'] || '';
        const pm       = p['PM'] || p['pmText'] || '';
        const assignee = p['담당자'] || '';
        const collab   = p['협업팀'] || p['collab'] || '';
        const issue    = p['이슈사항'] || p['issue'] || '';
        const startD   = fmtDate(p['착수일'] || p['startDate'] || '');
        const planD    = fmtDate(p['계획종료일'] || p['planDate'] || '');
        
        // 계획 공정 및 지연 판정
        let planPct = null;
        if (startD && planD) {
            const s = new Date(startD), e = new Date(planD), now = new Date();
            if (!isNaN(s) && !isNaN(e) && e > s)
                planPct = Math.min(100, Math.max(0, Math.round((now - s) / (e - s) * 100)));
        }
        const lag      = planPct !== null ? planPct - progress : 0;
        const isSevere = lag >= 10;
        const isLate   = lag > 0;

        // D-day
        let dday = '';
        if (planD) {
            const e = new Date(planD), now = new Date();
            now.setHours(0,0,0,0); e.setHours(0,0,0,0);
            const diff = Math.round((e - now) / 86400000);
            dday = diff > 0 ? 'D-' + diff : diff === 0 ? 'D-day' : 'D+' + Math.abs(diff);
        }

        let statusCls = '';
        if (status === '완료') statusCls = 'bg-emerald-100 text-emerald-700';
        else if (status === '보류') statusCls = 'bg-slate-100 text-slate-500';
        else if (status === '지연' || isSevere) statusCls = 'bg-red-100 text-red-700';
        else statusCls = 'bg-blue-100 text-blue-700';

        const barColor = progress >= 100 ? '#10b981' : isSevere ? '#ef4444' : isLate ? '#f97316' : '#1E3A8A';

        const card = document.createElement('div');
        card.className = 'project-card';
        // 상태별 하단 구분선 색상
        if (status === '완료') card.style.borderBottomColor = '#10b981';
        else if (status === '보류') card.style.borderBottomColor = '#94a3b8';
        else if (status === '지연' || isSevere) card.style.borderBottomColor = '#ef4444';
        else card.style.borderBottomColor = '#1E3A8A';
        const origIdx = projects.indexOf(p); // 필터 후에도 원본 배열 인덱스 유지
        
        if (origIdx === window.activeProjectIdx) {
            card.classList.add('active-card');
        }

        card.onclick = () => {
            window.activeProjectIdx = origIdx;
            document.querySelectorAll('.compact-card, .project-card').forEach(n => n.classList.remove('active-card'));
            card.classList.add('active-card');
            
            isDualMode ? renderExpandedDetail(origIdx) : openDetailPopup(origIdx);
        };
        card.innerHTML = `
            <div class="flex justify-between items-start">
                <div class="flex items-center gap-2">
                    <span class="px-3 py-1 rounded-full text-xs font-black uppercase tracking-wider ${statusCls}">${status}</span>
                    <span class="badge-${imp} text-xs font-bold px-2 py-0.5 rounded-full">${imp}</span>
                </div>
                <div class="text-right">
                    <span class="text-xs font-bold text-slate-500">PM ${pm}${assignee ? ` / co-worker ${assignee}` : ''}</span>
                    ${collab ? `<div class="text-[11px] text-slate-400 mt-0.5">🤝 ${collab}</div>` : ''}
                </div>
            </div>
            <div class="title mt-0.5 leading-tight">${title}</div>
            <div class="desc text-slate-500" style="font-size:13px; max-height:2.8em; overflow:hidden; display:-webkit-box; -webkit-line-clamp:2; -webkit-box-orient:vertical;">${desc}</div>
            ${issue && issue.trim() ? `
                <div class="bg-red-50 rounded-lg p-1.5 border border-red-100 mt-0.5">
                    <div class="flex items-center gap-1 font-black text-red-600 text-[11px] mb-0.5">
                        <span>⚠️ 이슈사항</span>
                    </div>
                    <ul class="text-xs text-red-700 leading-snug space-y-0.5" style="margin:0; padding-left:12px; list-style-type:disc;">
                        ${issue.trim().split('\n').filter(l => l.trim()).slice(0, 4).map(l => `<li>${l.trim()}</li>`).join('')}
                    </ul>
                </div>` : ''}
            <div class="mt-auto pt-1.5 border-t border-slate-100">
                <div class="flex justify-between items-center mb-1.5">
                    <div class="flex items-center gap-1.5">
                        <span class="text-[11px] font-bold text-slate-400">FINISH</span>
                        <span class="text-xs font-black text-slate-600">${planD || '-'}</span>
                        ${dday ? `<span class="bg-slate-800 text-white text-[11px] px-1.5 py-0.5 rounded font-black">${dday}</span>` : ''}
                    </div>
                    ${isSevere ? `<span class="text-[11px] font-black text-red-600 animate-pulse">계획대비 ${lag}%p↓</span>` : ''}
                </div>
                <div class="flex items-center gap-2">
                    <div class="flex-1 relative" style="height:8px;">
                        <div style="position:absolute;inset:0;background:#f1f5f9;border-radius:4px;overflow:hidden;">
                            <div style="height:100%;width:${progress}%;background:${barColor};transition:width 0.6s ease;"></div>
                        </div>
                        ${planPct !== null ? `<div title="계획공정 ${planPct}%" style="position:absolute;top:-3px;bottom:-3px;left:calc(${planPct}% - 1px);width:2px;background:#334155;border-radius:1px;z-index:2;"></div>` : ''}
                    </div>
                    <span class="text-xs font-black text-slate-800 shrink-0">${progress}%</span>
                </div>
            </div>
        `;
        grid.appendChild(card);
    });
}

function prevProjectPage() {
    if (currentProjectPage > 0) {
        currentProjectPage--;
        renderProjects();
    }
}

function nextProjectPage() {
    const totalPages = Math.ceil(projects.length / PROJECTS_PER_PAGE);
    if (currentProjectPage < totalPages - 1) {
        currentProjectPage++;
        renderProjects();
    }
}


function updateSummary(projs) {
    let ok=0, late=0, hold=0, done=0;
    const teamMap = {}; // 팀원별 데이터 집계

    projs.forEach(p => {
        const st  = p['상태'] || p['status'] || '진행';
        const prg = parseInt(p['진행률'] || 0);
        const pm  = (p['PM'] || '').trim();
        const assignee = (p['담당자'] || '').trim();
        
        // 팀원 추출 (PM과 담당자 모두 포함)
        const members = new Set();
        if (pm) members.add(pm);
        if (assignee) members.add(assignee);

        let currentStatus = '정상';
        if (st === '보류') { hold++; currentStatus = '보류'; }
        else if (st === '완료' || prg >= 100) { done++; currentStatus = '완료'; }
        else {
            const startD = fmtDate(p['착수일'] || '');
            const planD  = fmtDate(p['계획종료일'] || '');
            let planPct  = null;
            if (startD && planD) {
                const s = new Date(startD), e = new Date(planD), now = new Date();
                if (!isNaN(s) && !isNaN(e) && e > s)
                    planPct = Math.min(100, Math.max(0, Math.round((now-s)/(e-s)*100)));
            }
            if (planPct !== null && (planPct - prg) >= 10) { currentStatus = '지연'; }
            else currentStatus = '정상';
        }
        if (currentStatus === '지연') late++;
        else ok++;

        // 팀원별 프로젝트 및 참여율 집합 구성
        members.forEach(m => {
            if (!teamMap[m]) teamMap[m] = { name: m, projects: [], stats: { 정상:0, 지연:0, 보류:0, 완료:0 } };
            teamMap[m].stats[currentStatus]++;
            
            // 참여율 데이터 확인 (시트에 '참여율' 컬럼이 있는 경우 우선 사용)
            let ratio = p['참여율'] || p['ratio'] || '';
            teamMap[m].projects.push({ title: p['업무명'] || '과업', status: currentStatus, ratio: ratio });
        });
    });

    const set = (id,v) => { const el=document.getElementById(id); if(el) el.textContent=v; };
    set('sum-total', projs.length);
    set('sum-ok',    ok);
    set('sum-late',  late);
    set('sum-hold',  hold);
    set('sum-done',  done);

    renderTeamInvolvement(teamMap);
}

const MEMBER_ORDER = ['김우진', '임민경', '최선아', '김윤재', '이미영', '국혜림'];
const MEMBER_INFO = {
    '김우진': { role: '팀장', avatar: 'assets/img/avatars/김우진.png' },
    '임민경': { role: '책임', avatar: 'assets/img/avatars/임민경.png' },
    '최선아': { role: '책임', avatar: 'assets/img/avatars/최선아.png' },
    '김윤재': { role: '선임', avatar: 'assets/img/avatars/김윤재.png' },
    '이미영': { role: '선임', avatar: 'assets/img/avatars/이미영.png' },
    '국혜림': { role: '선임', avatar: 'assets/img/avatars/국혜림.png' }
};

function renderTeamInvolvement(teamMap) {
    const grid = document.getElementById('teamInvolvementGrid');
    if (!grid) return;
    grid.innerHTML = '';

    MEMBER_ORDER.forEach(name => {
        // teamInfo(시트 탭)에서 데이터 찾기, 없으면 기존 routineData(키-값) 검색
        const tInfo = teamInfo.find(t => t['이름'] === name) || {};
        
        const role = tInfo['직함'] || routineData[name + '_직함'] || MEMBER_INFO[name]?.role || '책임';
        let mainR = tInfo['주요비중'] || routineData[name + '_주요비중'] || '0';
        let routR = tInfo['상시비중'] || routineData[name + '_상시비중'] || '0';

        if (mainR && !mainR.toString().includes('%')) mainR += '%';
        if (routR && !routR.toString().includes('%')) routR += '%';

        const tasksStr = tInfo['과업리스트'] || routineData[name + '_과업리스트'] || '';
        
        const m = teamMap[name] || { name: name, projects: [] };
        const info = MEMBER_INFO[name] || { avatar: 'assets/img/avatars/mem1.png' };
        
        // 자동 집계 프로젝트 + 수동 추가 과업 통합
        const autoTags = m.projects.map(p => p.title);
        const manualTags = tasksStr ? tasksStr.split(',').map(t => t.trim()).filter(t => t !== '') : [];
        const allTasks = [...new Set([...autoTags, ...manualTags])]; // 중복 제거 통함

        const popupHtml = allTasks.map(t => `
            <div class="flex items-center justify-between text-[11px] mb-1">
                <span class="truncate pr-2 flex items-center gap-1"><span class="w-1.5 h-1.5 rounded-full bg-emerald-500"></span> ${t}</span>
            </div>
        `).join('') || '<div class="text-[11px] text-slate-400 text-center">진행 중인 과업 없음</div>';

        const card = document.createElement('div');
        card.className = 'member-wide-card';
        // 필터 활성 상태 반영
        if (activeFilterMember === name) card.classList.add('filter-active');
        else if (activeFilterMember && activeFilterMember !== name) card.classList.add('filter-dim');

        card.innerHTML = `
            <div class="member-avatar" style="background-image:url('${info.avatar}');"></div>
            <div class="member-role-tag">${role}</div>
            <div class="text-center w-full">
                <div class="text-xl font-black text-slate-800 mb-1">${name}</div>
                <div class="flex justify-center items-center gap-2 mb-2">
                    <div class="flex flex-col items-center px-2 border-r border-slate-200">
                        <span class="text-[11px] text-slate-400 font-bold uppercase">Main</span>
                        <span class="text-xs font-black text-slate-700">${mainR}</span>
                    </div>
                    <div class="flex flex-col items-center px-1">
                        <span class="text-[11px] text-slate-400 font-bold uppercase">Routine</span>
                        <span class="text-xs font-black text-slate-700">${routR}</span>
                    </div>
                </div>
            </div>
            <div class="member-ratio-popup">
                <div class="font-black text-sm mb-3 border-b pb-2 flex justify-between">
                    <span>📊 ${name} 업무 리스트</span>
                </div>
                <div class="space-y-1">
                    ${popupHtml}
                </div>
            </div>
        `;

        // 팀원 카드 클릭 → 필터 토글 (renderProjects → updateSummary → renderTeamInvolvement 순으로 재렌더)
        card.addEventListener('click', (e) => {
            e.stopPropagation();
            activeFilterMember = (activeFilterMember === name) ? null : name;
            currentProjectPage = 0;
            updateFilterBadge();
            renderProjects(); // updateSummary → renderTeamInvolvement 포함
        });

        grid.appendChild(card);
    });
}

function updateFilterBadge() {
    const badge = document.getElementById('filterBadge');
    const nameEl = document.getElementById('filterBadgeName');
    if (!badge) return;
    if (activeFilterMember) {
        nameEl.textContent = activeFilterMember + ' 필터';
        badge.style.display = 'flex';
    } else {
        badge.style.display = 'none';
    }
}

function clearMemberFilter() {
    activeFilterMember = null;
    currentProjectPage = 0;
    updateFilterBadge();
    renderProjects(); // updateSummary → renderTeamInvolvement 포함하여 dim/active 자동 초기화
}

// 클릭 외부 시 팝업 닫기 (필터는 유지)
document.addEventListener('click', () => {
    document.querySelectorAll('.member-wide-card').forEach(c => c.classList.remove('active'));
});



function renderTeamMemos() {
    const list = document.getElementById('teamMemoList');
    list.innerHTML = '';
    if (teamMemos.length === 0) {
        list.innerHTML = `<p class="text-xs text-gray-300 text-center mt-4">구글시트 <strong>팀공유사항</strong> 탭에 입력하세요</p>`;
        return;
    }
    const frag = document.createDocumentFragment();
    teamMemos.forEach(m => {
        const hasStatus = m.status && m.status.trim() && m.status !== '담당자 확인';
        const item = document.createElement('div');
        item.className = 'px-3 py-2.5 border-b border-slate-100 flex flex-col gap-1.5';
        item.innerHTML = `
            <div class="flex items-center gap-1.5 flex-wrap min-w-0">
                ${m.projectId ? `<span class="text-xs font-black text-[#1E3A8A] bg-[#1E3A8A]/10 px-2 py-0.5 rounded shrink-0">[${m.projectId}]</span>` : ''}
                <span class="text-sm font-bold text-slate-700 flex-1 min-w-0 truncate">${m.author || ''}</span>
                ${hasStatus ? `<span class="text-xs font-bold bg-slate-100 text-slate-600 px-2 py-0.5 rounded shrink-0">[${m.status}]</span>` : ''}
            </div>
            <p class="text-sm text-slate-600 leading-relaxed whitespace-pre-wrap">${m.text || ''}</p>`;
        frag.appendChild(item);
    });
    list.appendChild(frag);
}

function renderCeoMemos() {
    const list = document.getElementById('ceoMemoList');
    list.innerHTML = '';
    if (ceoMemos.length === 0) {
        list.innerHTML = `<p class="text-xs text-rose-200 text-center mt-4">구글시트 <strong>사장님지시사항</strong> 탭에 입력하세요</p>`;
        return;
    }
    // 시트 드롭다운 값과 일치하도록 모든 상태값 색상 매핑
    // 시트에 없는 값도 자동으로 색상이 지정되도록 동적 처리
    const statusColor = {
        // 기존 값
        '담당자 확인': 'bg-slate-600',
        '반영 예정':   'bg-blue-600',
        '별도 보고':   'bg-purple-600',
        '협의 필요':   'bg-rose-500',
        // 시트 드롭다운 추가값 — 필요시 여기에 계속 추가
        '추진 예정':   'bg-green-600',
        '계획 검토':   'bg-yellow-500',
        '완료':        'bg-emerald-600',
        '보류':        'bg-gray-500',
        '긴급':        'bg-red-600',
    };
    // 위에 없는 값은 텍스트 길이 기반으로 색상 자동 배정
    function getStatusColor(status) {
        if (statusColor[status]) return statusColor[status];
        const colors = ['bg-slate-500','bg-teal-600','bg-cyan-600','bg-indigo-500','bg-rose-500'];
        return colors[status.length % colors.length];
    }

    const sorted = [...ceoMemos].reverse();
    const frag2 = document.createDocumentFragment();
    sorted.forEach((m, i) => {
        const sc = getStatusColor(m.status || '');
        const hasStatus = m.status && m.status.trim();
        const isNewest = i === 0;
        const hasProject = m.projectId && m.projectId.trim();
        const hoverCls = hasProject ? (isNewest ? 'hover:bg-[#E11D48]/20' : 'hover:bg-orange-50') : '';
        const itemCls = `${hoverCls} transition-colors rounded-lg ${isNewest
            ? 'bg-[#E11D48]/10 p-3 border-2 border-[#E11D48]/40 space-y-1.5'
            : 'bg-white p-2.5 border border-orange-100 space-y-1'}`;
        const item = document.createElement('div');
        item.className = itemCls;
        if (hasProject) {
            item.style.cursor = 'pointer';
            item.addEventListener('click', () => focusProject(m.projectId));
        }
        item.innerHTML = `
            ${isNewest ? '<div class="text-[11px] font-black text-[#E11D48] mb-1">🔔 최신 지시사항</div>' : ''}
            <div class="flex items-center gap-1 flex-wrap">
                ${hasProject ? `<span class="text-[11px] font-bold bg-orange-100 text-orange-700 px-1.5 py-0.5 rounded">🔗 ${m.projectId}</span>` : ''}
                ${hasStatus ? `<span class="text-[11px] font-bold ${sc} text-white px-2 py-0.5 rounded ml-auto">${m.status}</span>` : ''}
            </div>
            <p class="${isNewest ? 'text-sm' : 'text-xs'} text-slate-700 leading-relaxed whitespace-pre-wrap font-medium">${m.text || ''}</p>`;
        frag2.appendChild(item);
    });
    list.appendChild(frag2);
}

// ── 지시사항 클릭 → 프로젝트 카드 포커스 ────────────────────
function focusProject(id) {
    if (!id) return;
    const idTrimmed = id.trim();

    // 공통 매칭 함수 사용
    let idx = resolveProjectIndex(idTrimmed);

    if (idx < 0) {
        showToast('해당 과업을 찾을 수 없습니다: ' + idTrimmed);
        return;
    }

    // 필터 활성 시 해제
    if (activeFilterMember) {
        activeFilterMember = null;
        updateFilterBadge();
        renderProjects();
    }

    // 스크롤 방식: 전체 렌더된 카드 목록에서 해당 카드를 찾아 스크롤
    setTimeout(() => {
        const cards = document.querySelectorAll('#projectGrid .project-card');
        // filteredProjects 내에서의 순서를 찾기 (필터 해제 후이므로 projects와 동일)
        const cardEl = cards[idx];
        if (cardEl) {
            cardEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
            cardEl.classList.add('highlight-card');
            setTimeout(() => cardEl.classList.remove('highlight-card'), 2200);
        }
    }, 80);
}

// ── 간단한 토스트 알림 ──────────────────────────────────────
function showToast(msg) {
    let toast = document.getElementById('toastMsg');
    if (!toast) {
        toast = document.createElement('div');
        toast.id = 'toastMsg';
        toast.style.cssText = 'position:fixed;bottom:24px;left:50%;transform:translateX(-50%);background:#334155;color:white;padding:10px 20px;border-radius:10px;font-size:13px;font-weight:600;z-index:999;opacity:0;transition:opacity 0.3s;pointer-events:none;';
        document.body.appendChild(toast);
    }
    toast.textContent = msg;
    toast.style.opacity = '1';
    setTimeout(() => { toast.style.opacity = '0'; }, 2500);
}

function renderRoutines() {
    const keys = ['경영현황','동종업계','지재권','CCP','팀지원','바론사업','ERP','법인관리'];
    keys.forEach(k => {
        const el = document.getElementById('routine-' + k);
        if (el) el.textContent = routineData[k] || '(내용 없음)';
        const attEl = document.getElementById('routine-att-' + k);
        if (attEl) renderRoutineAtt(attEl, routineData[k + '_파일'] || '');
    });
}

function renderRoutineAtt(container, raw) {
    container.innerHTML = '';
    if (!raw.trim()) return;
    raw.split('\n').forEach(line => {
        line = line.trim(); if (!line) return;
        const parts = line.split('|');
        const name = parts[0].trim();
        const url  = (parts[1] || '').trim();
        const finalUrl = url || (name.startsWith('http') ? name : '');
        const a = document.createElement('a');
        a.textContent = '\uD83D\uDCC4 ' + name;
        a.className = 'inline-flex items-center text-xs font-semibold text-[#1E3A8A] bg-slate-100 hover:bg-[#1E3A8A] hover:text-white px-2.5 py-1 rounded-full border border-slate-200 transition-all cursor-pointer mr-1 mb-1';
        if (finalUrl) {
            a.title = finalUrl;
            a.addEventListener('click', function(e) {
                e.stopPropagation(); e.preventDefault();
                window.open(finalUrl, '_blank', 'noopener,noreferrer');
            });
        } else { a.style.opacity='0.5'; a.style.cursor='default'; }
        container.appendChild(a);
    });
}

// ── 모달 ─────────────────────────────────────────────────
function openModal(idx) {
    const p = projects[idx];
    const pm       = p['PM'] || p['pmText'] || '';
    const assignee = p['담당자'] || '';
    const collab   = p['협업팀'] || p['collab'] || '';
    const startD   = fmtDate(p['착수일'] || p['startDate'] || '');
    const planD    = fmtDate(p['계획종료일'] || p['planDate'] || '');

    document.getElementById('modalTitle').innerText   = (p['업무명'] || p['title'] || '').replace(/\n/g,' ');
    document.getElementById('mPmBadge').textContent       = pm || '-';
    document.getElementById('mAssigneeBadge').textContent = assignee || '-';
    document.getElementById('mCollabBadge').textContent   = collab || '-';
    document.getElementById('mStartBadge').textContent    = startD || '-';
    document.getElementById('mPlanBadge').textContent     = planD || '-';
    const durBadge = document.getElementById('mDurBadge');
    if (startD && planD) {
        const _s = new Date(startD), _e = new Date(planD), _n = new Date();
        _s.setHours(0,0,0,0); _e.setHours(0,0,0,0); _n.setHours(0,0,0,0);
        const total   = Math.round((_e - _s) / 86400000);
        const elapsed = Math.max(0, Math.min(total, Math.round((_n - _s) / 86400000)));
        const remain  = Math.max(0, Math.round((_e - _n) / 86400000));
        document.getElementById('mDurTotal').textContent   = '총 ' + total + '일';
        document.getElementById('mDurElapsed').textContent = '진행 ' + elapsed + '일';
        document.getElementById('mDurRemain').textContent  = '잔여 ' + remain + '일';
        durBadge.style.display = 'flex';
    } else {
        durBadge.style.display = 'none';
    }
    document.getElementById('mPm').value       = pm;
    document.getElementById('mAssignee').value = assignee;
    const mCollabEl = document.getElementById('mCollab'); if(mCollabEl) mCollabEl.value = collab;
    document.getElementById('mGoals').value    = p['과업내용'] || p['goals'] || '';
    document.getElementById('mIssue').value    = p['이슈사항'] || p['issue'] || '';
    document.getElementById('mSeminars').value = p['회의이력'] || p['seminars'] || '';

    // 주요 URL 버튼 (3-2) — 구글시트 프로젝트탭 "주요URL" 컬럼: 제목|URL (줄바꿈으로 여러 개)
    const urlBox = document.getElementById('mUrlLinks');
    urlBox.innerHTML = '';
    const rawUrls = (p['주요URL'] || p['주요 URL'] || '').trim();
    if (rawUrls) {
        rawUrls.split('\n').forEach(line => {
            line = line.trim(); if (!line) return;
            const parts = line.split('|');
            const name = parts[0].trim();
            const url  = (parts[1] || '').trim();
            const finalUrl = url || (name.startsWith('http') ? name : '');
            const btn = document.createElement('button');
            btn.innerHTML = '🔗 ' + name;
            btn.className = 'text-[11px] font-bold bg-amber-400 hover:bg-amber-300 text-amber-900 px-3 py-1.5 rounded-lg border border-amber-300 transition-all shadow-sm';
            if (finalUrl) btn.addEventListener('click', e => { e.stopPropagation(); window.open(finalUrl,'_blank','noopener,noreferrer'); });
            urlBox.appendChild(btn);
        });
    }

    // 붙임파일/하이퍼링크 렌더링
    // 구글시트에서 '붙임파일' 컬럼: "파일명|URL" 형식으로 여러 개는 줄바꿈(\n)으로 구분
    const attBox = document.getElementById('mAttachments');
    attBox.innerHTML = '';
    const rawAtt = p['붙임파일'] || p['attachments'] || '';
    if (rawAtt.trim()) {
        const label = document.createElement('span');
        label.className = 'text-xs font-bold text-slate-400 mr-1 shrink-0';
        label.textContent = '📎 붙임파일:';
        attBox.appendChild(label);
        rawAtt.split('\n').forEach(line => {
            line = line.trim();
            if (!line) return;
            const parts = line.split('|');
            const name = parts[0].trim();
            const url  = (parts[1] || '').trim();

            const a = document.createElement('a');
            a.textContent = '📄 ' + name;
            a.className = 'inline-flex items-center gap-1 text-xs font-semibold text-[#1E3A8A] bg-slate-100 hover:bg-[#1E3A8A] hover:text-white px-3 py-1.5 rounded-full border border-slate-200 transition-all cursor-pointer';

            const finalUrl = url || (name.startsWith('http') ? name : '');

            if (finalUrl) {
                a.title = finalUrl;
                // ★ 카드 전체 onclick이 링크를 가로채지 못하도록 전파 차단 후 새 탭 열기
                a.addEventListener('click', e => {
                    e.stopPropagation();
                    e.preventDefault();
                    window.open(finalUrl, '_blank', 'noopener,noreferrer');
                });
            } else {
                a.style.opacity = '0.5';
                a.style.cursor = 'default';
            }

            attBox.appendChild(a);
        });
    }

    document.getElementById('modalOverlay').style.display = 'flex';
}
function closeModal() { document.getElementById('modalOverlay').style.display = 'none'; }

function openRoutineModal() { document.getElementById('routineModalOverlay').style.display = 'flex'; }
function closeRoutineModal() { document.getElementById('routineModalOverlay').style.display = 'none'; }

function switchTab(name, btn) {
    document.querySelectorAll('.tab-panel').forEach(el => el.classList.add('hidden'));
    document.querySelectorAll('.tab-btn').forEach(el => el.classList.remove('active'));
    document.getElementById(`tab-${name}`).classList.remove('hidden');
    btn.classList.add('active');
}

// ── KPI 요약바 렌더링 ───────────────────────────────────
function renderKpiBar() {
    // ─ D-day (2Q 마감: 6월 30일) ─
    const qEnd = new Date(today.getFullYear(), 5, 30);
    qEnd.setHours(23,59,59,0);
    const dDiff = Math.ceil((qEnd - today) / 86400000);
    const ddEl = document.getElementById('kpi-dday-count');
    if (ddEl) ddEl.textContent = dDiff > 0 ? 'D-' + dDiff : dDiff === 0 ? 'D-day' : 'D+' + Math.abs(dDiff);
    const ddDateEl = document.getElementById('kpi-dday-date');
    if (ddDateEl) ddDateEl.textContent = today.getFullYear() + '.06.30 마감';

    // ─ 과업 통계 계산 ─
    let total = projects.length, ok = 0, late = 0, hold = 0, done = 0, totalPrg = 0;
    let lateProjs = [];
    const collabSet = new Set();

    projects.forEach(p => {
        const st  = p['상태'] || p['status'] || '진행';
        const prg = parseInt(p['진행률'] || 0);
        totalPrg += prg;

        // 협업팀 수집
        const collab = (p['협업팀'] || p['collab'] || '').trim();
        if (collab) collab.split(/[,·\/]/).forEach(t => { const c = t.trim(); if (c) collabSet.add(c); });

        if (st === '보류')               { hold++; return; }
        if (st === '완료' || prg >= 100) { done++; return; }

        const startD = fmtDate(p['착수일'] || '');
        const planD  = fmtDate(p['계획종료일'] || '');
        let planPct  = null;
        if (startD && planD) {
            const s = new Date(startD), e = new Date(planD), now = new Date();
            if (!isNaN(s) && !isNaN(e) && e > s)
                planPct = Math.min(100, Math.max(0, Math.round((now-s)/(e-s)*100)));
        }
        const lag = planPct !== null ? planPct - prg : 0;
        if (lag >= 10) {
            late++;
            lateProjs.push({ name: p['업무명'] || p['title'] || '과업', lag });
        } else ok++;
    });

    const avgPrg  = total > 0 ? Math.round(totalPrg / total) : 0;
    const latePct = total > 0 ? Math.round(late / total * 100) : 0;
    const collabs = [...collabSet];

    // ── ① 과업 평균 진행률 ──
    const el_tp = document.getElementById('kpi-task-pct');
    if (el_tp) el_tp.textContent = total > 0 ? avgPrg + '%' : '—';

    // 스택 바 (ok/late/hold/done 비율)
    if (total > 0) {
        const setW = (id, n) => { const el = document.getElementById(id); if (el) el.style.width = Math.round(n/total*100) + '%'; };
        setW('ksb-ok',   ok);
        setW('ksb-late', late);
        setW('ksb-hold', hold);
        setW('ksb-done', done);
    }

    const pillsEl = document.getElementById('kpi-task-pills');
    if (pillsEl) {
        pillsEl.innerHTML = total > 0
            ? `<span class="kpi-pill" style="background:rgba(52,211,153,0.18);color:#6ee7b7">정상 ${ok}</span>`
            + `<span class="kpi-pill" style="background:rgba(249,115,22,0.18);color:#fdba74">지연 ${late}</span>`
            + `<span class="kpi-pill" style="background:rgba(148,163,184,0.18);color:#cbd5e1">보류 ${hold}</span>`
            + `<span class="kpi-pill" style="background:rgba(96,165,250,0.18);color:#93c5fd">완료 ${done}</span>`
            : '<span style="font-size:9px;color:rgba(255,255,255,0.3)">데이터 없음</span>';
    }

    // ── ② 지연 비율 ──
    const delayCntEl = document.getElementById('kpi-delay-count');
    if (delayCntEl) {
        delayCntEl.textContent = late;
        delayCntEl.style.color = late === 0 ? '#34d399' : late <= 2 ? '#fbbf24' : '#f87171';
    }
    const delayDetailEl = document.getElementById('kpi-delay-detail');
    if (delayDetailEl) {
        delayDetailEl.textContent = total > 0
            ? (late === 0 ? '전체 과업 정상 진행 중' : `전체 ${total}건 중 ${latePct}% 지연`)
            : '—';
    }

    // 반원 게이지 (지연 비율)
    const gaugeArc = document.getElementById('kpi-gauge-arc');
    const gaugePct = document.getElementById('kpi-gauge-pct');
    if (gaugeArc) {
        const arcLen = 69.1;
        const offset = arcLen * (1 - latePct / 100);
        gaugeArc.style.strokeDashoffset = offset;
        gaugeArc.style.stroke = latePct === 0 ? '#34d399' : latePct <= 20 ? '#fbbf24' : '#f87171';
    }
    if (gaugePct) gaugePct.textContent = latePct + '%';

    // 지연 과업 칩 (최대 2개)
    const dchipsEl = document.getElementById('kpi-delay-chips');
    if (dchipsEl) {
        dchipsEl.innerHTML = '';
        lateProjs.slice(0, 3).forEach(sp => {
            const chip = document.createElement('div');
            chip.style.cssText = 'background:rgba(248,113,113,0.12);border:1px solid rgba(248,113,113,0.3);border-radius:4px;padding:2px 7px;font-size:9px;font-weight:700;color:#fca5a5;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:140px;flex-shrink:0';
            chip.title = sp.name;
            chip.textContent = `▼${sp.lag}%p · ${sp.name.length > 10 ? sp.name.slice(0,9)+'…' : sp.name}`;
            dchipsEl.appendChild(chip);
        });
        if (lateProjs.length > 3) {
            const more = document.createElement('div');
            more.style.cssText = 'font-size:9px;color:rgba(255,255,255,0.3);font-weight:600;white-space:nowrap;flex-shrink:0';
            more.textContent = `+${lateProjs.length - 3}건`;
            dchipsEl.appendChild(more);
        }
    }

    // ── ③ 협업팀 커버리지 ──
    const collabNumEl = document.getElementById('kpi-collab-num');
    if (collabNumEl) collabNumEl.textContent = collabs.length || '—';

    const collabTagsEl = document.getElementById('kpi-collab-tags');
    if (collabTagsEl) {
        collabTagsEl.innerHTML = '';
        collabs.slice(0, 6).forEach(name => {
            const tag = document.createElement('span');
            tag.style.cssText = 'background:rgba(147,197,253,0.12);border:1px solid rgba(147,197,253,0.2);border-radius:4px;padding:2px 5px;font-size:9px;font-weight:700;color:#93c5fd;white-space:nowrap';
            tag.textContent = name.length > 6 ? name.slice(0,5)+'…' : name;
            tag.title = name;
            collabTagsEl.appendChild(tag);
        });
        if (collabs.length > 6) {
            const more = document.createElement('span');
            more.style.cssText = 'font-size:9px;color:rgba(255,255,255,0.3);font-weight:600;padding:2px';
            more.textContent = `+${collabs.length - 6}`;
            collabTagsEl.appendChild(more);
        }
        if (collabs.length === 0) {
            collabTagsEl.innerHTML = '<span style="font-size:9px;color:rgba(255,255,255,0.3)">협업팀 정보 없음</span>';
        }
    }
}

// ── 유틸 ─────────────────────────────────────────────────
function showLoading(msg) {
    document.getElementById('loadingOverlay').style.display = 'flex';
    document.querySelector('#loadingOverlay p').textContent = msg || '로딩 중...';
}
function hideLoading() { document.getElementById('loadingOverlay').style.display = 'none'; }
function setStatus(msg) { document.getElementById('syncStatus').textContent = msg; }

// ── 확장 모드 (Dual Monitor Support) ─────────────────────────────
let isDualMode = false;

function toggleDualMode() {
    isDualMode = !isDualMode;
    document.body.classList.toggle('dual-mode', isDualMode);
    const btn     = document.getElementById('dualModeBtn');
    const btnText = document.getElementById('dualModeBtnText');
    if (isDualMode) {
        btn.classList.add('active');
        btnText.textContent = '일반 모드';
        showExpandedPlaceholder();
    } else {
        btn.classList.remove('active');
        btnText.textContent = '확장 모드';
        showExpandedPlaceholder();
    }
}

function showExpandedPlaceholder() {
    const panel = document.getElementById('expandedPanel');
    if (!panel) return;
    panel.innerHTML = `<div style="display:flex;flex-direction:column;align-items:center;justify-content:center;height:100%;color:#94a3b8;padding:32px;text-align:center">
        <div style="font-size:52px;margin-bottom:16px;opacity:0.4">🖥️</div>
        <p style="font-size:14px;font-weight:700;color:#64748b;margin-bottom:6px">프로젝트 카드를 클릭하세요</p>
        <p style="font-size:12px">선택한 과업의 상세 내용이 이 영역에 표시됩니다</p>
    </div>`;
}

// ── 상세 내용 공통 렌더러 (확장 패널 & 팝업 공유) ──────────────
function _renderDetailContent(idx, container) {
    const p = projects[idx];
    if (!p || !container) return;

    const pm       = p['PM'] || p['pmText'] || '';
    const assignee = p['담당자'] || '';
    const collab   = p['협업팀'] || p['collab'] || '';
    const startD   = fmtDate(p['착수일'] || p['startDate'] || '');
    const planD    = fmtDate(p['계획종료일'] || p['planDate'] || '');
    const title    = p['업무명'] || p['title'] || '(제목 없음)';
    const goals    = p['과업내용'] || p['goals'] || '';
    const issue    = p['이슈사항'] || p['issue'] || '';
    const seminars = p['회의이력'] || p['seminars'] || '';
    const status   = p['상태'] || p['status'] || '진행';
    const imp      = p['우선순위'] || p['importance'] || '중';
    const progress = parseInt(p['진행률'] || 0);

    // 기간 및 계획공정 계산
    let durHtml = '';
    let detailPlanPct = null;
    if (startD && planD) {
        const _s = new Date(startD), _e = new Date(planD), _n = new Date();
        _s.setHours(0,0,0,0); _e.setHours(0,0,0,0); _n.setHours(0,0,0,0);
        const totalD   = Math.round((_e - _s) / 86400000);
        const elapsedD = Math.max(0, Math.min(totalD, Math.round((_n - _s) / 86400000)));
        const remainD  = Math.max(0, Math.round((_e - _n) / 86400000));
        if (_e > _s) detailPlanPct = Math.min(100, Math.max(0, Math.round((_n - _s) / (_e - _s) * 100)));
        durHtml = `<div style="background:rgba(0,0,0,0.22);border:1px solid rgba(255,255,255,0.18);border-radius:8px;padding:5px 12px;display:flex;align-items:center;gap:8px">
            <span style="font-size:11px;font-weight:700;opacity:0.5;text-transform:uppercase;letter-spacing:0.05em">기간</span>
            <span style="font-size:13px;font-weight:900;color:#fff">총 ${totalD}일</span>
            <span style="opacity:0.3">|</span>
            <span style="font-size:11px;font-weight:700;color:#7dd3fc">진행 ${elapsedD}일</span>
            <span style="opacity:0.3">/</span>
            <span style="font-size:11px;font-weight:700;color:#fde68a">잔여 ${remainD}일</span>
        </div>`;
    }
    const detailLag = detailPlanPct !== null ? detailPlanPct - progress : 0;
    const detailIsLate = detailLag > 0;
    const detailIsSevere = detailLag >= 10;

    const barBg = progress >= 100 ? '#10b981' : detailIsSevere ? '#ef4444' : detailIsLate ? '#f97316' : status === '보류' ? '#94a3b8' : '#1E3A8A';

    const esc = s => s ? s.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;') : '';
    const issueEscaped   = esc(issue);
    const goalsEscaped   = esc(goals);
    const seminarsEscaped= esc(seminars);
    const rawAtt = (p['붙임파일'] || p['attachments'] || '').trim();

    container.innerHTML = `
    <div style="padding:22px;display:flex;flex-direction:column;gap:18px;min-height:100%;box-sizing:border-box">

        <!-- ① 헤더 그라디언트 배너 -->
        <div style="background:linear-gradient(135deg,#1E3A8A 0%,#172554 100%);border-radius:12px;padding:20px 22px;color:white;flex-shrink:0">
            <h2 style="font-size:18px;font-weight:900;line-height:1.2;margin-bottom:12px;letter-spacing:-0.3px">${esc(title)}</h2>
            <!-- 인적 정보 -->
            <div style="display:flex;flex-wrap:wrap;gap:7px;margin-bottom:9px;font-size:12px">
                <div style="background:rgba(255,255,255,0.14);border-radius:8px;padding:4px 11px;display:flex;gap:5px;align-items:center">
                    <span style="opacity:0.55">PM</span><strong>${pm || '-'}</strong>
                </div>
                <div style="background:rgba(255,255,255,0.14);border-radius:8px;padding:4px 11px;display:flex;gap:5px;align-items:center">
                    <span style="opacity:0.55">co-worker</span><strong>${assignee || '-'}</strong>
                </div>
                <div style="background:rgba(255,255,255,0.14);border-radius:8px;padding:4px 11px;display:flex;gap:5px;align-items:center">
                    <span style="opacity:0.55">협업팀</span><strong>${collab || '-'}</strong>
                </div>
                <span class="badge-${imp}" style="border-radius:999px;padding:3px 9px;font-size:10px;font-weight:700;align-self:center">${imp}</span>
                <span class="status-${status}" style="border-radius:999px;padding:3px 10px;font-size:11px;font-weight:700;align-self:center">${status}</span>
            </div>
            <!-- 일정 정보 -->
            <div style="display:flex;flex-wrap:wrap;gap:7px;margin-bottom:13px;font-size:12px">
                <div style="background:rgba(0,0,0,0.2);border:1px solid rgba(255,255,255,0.18);border-radius:8px;padding:4px 12px;display:flex;gap:7px;align-items:center">
                    <span style="font-size:11px;font-weight:700;opacity:0.5;text-transform:uppercase;letter-spacing:0.05em">착수</span>
                    <span style="font-weight:900;font-size:13px;color:#6ee7b7">${startD || '-'}</span>
                </div>
                <div style="background:rgba(0,0,0,0.2);border:1px solid rgba(255,255,255,0.18);border-radius:8px;padding:4px 12px;display:flex;gap:7px;align-items:center">
                    <span style="font-size:11px;font-weight:700;opacity:0.5;text-transform:uppercase;letter-spacing:0.05em">완료예정</span>
                    <span style="font-weight:900;font-size:13px;color:#fdba74">${planD || '-'}</span>
                </div>
                ${durHtml}
            </div>
            <!-- 진행률 바 (계획공정 마커 포함) -->
            <div>
                <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:5px">
                    <div style="display:flex;align-items:center;gap:8px">
                        <span style="font-size:11px;font-weight:700;opacity:0.65">진행률</span>
                        ${detailIsSevere ? `<span style="font-size:11px;font-weight:800;color:#fca5a5;animation:pulse 2s infinite">계획대비 ${detailLag}%p↓</span>` : ''}
                    </div>
                    <span style="font-size:20px;font-weight:900">${progress}%</span>
                </div>
                <div style="position:relative;height:10px;">
                    <div style="position:absolute;inset:0;background:rgba(255,255,255,0.15);border-radius:5px;overflow:hidden;">
                        <div style="height:100%;width:${progress}%;background:${barBg};border-radius:5px;transition:width 0.9s ease"></div>
                    </div>
                    ${detailPlanPct !== null ? `<div title="계획공정 ${detailPlanPct}%" style="position:absolute;top:-3px;bottom:-3px;left:calc(${detailPlanPct}% - 1px);width:2px;background:rgba(255,255,255,0.7);border-radius:1px;z-index:2"></div>` : ''}
                </div>
            </div>
            <!-- 주요 URL 버튼 영역 -->
            <div id="exp-urlLinks" style="display:flex;flex-wrap:wrap;gap:6px;margin-top:12px"></div>
        </div>

        <!-- ② 본문 2컬럼 그리드 -->
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;flex:1;min-height:0">
            <!-- 왼쪽: 과업내용 + 이슈 -->
            <div style="display:flex;flex-direction:column;gap:12px;min-height:0">
                <div style="display:flex;flex-direction:column;flex:1;min-height:0">
                    <div class="exp-label">과업 목표 및 내용</div>
                    <div class="exp-textbox" style="min-height:120px">${goalsEscaped || '(내용 없음)'}</div>
                </div>
                ${issueEscaped ? `<div style="display:flex;flex-direction:column;flex-shrink:0">
                    <div class="exp-label">⚠ 이슈사항</div>
                    <div class="exp-textbox" style="white-space:pre-wrap">${issueEscaped}</div>
                </div>` : ''}
            </div>
            <!-- 오른쪽: 회의 이력 -->
            <div style="display:flex;flex-direction:column;min-height:0">
                <div class="exp-label">세미나 및 회의 이력</div>
                <div class="exp-textbox" style="min-height:200px;overflow-y:auto">${seminarsEscaped || '(내용 없음)'}</div>
            </div>
        </div>

        <!-- ③ 붙임파일 -->
        ${rawAtt ? `<div id="exp-attachments" style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;padding-top:12px;border-top:1px solid #dde3eb;flex-shrink:0">
            <span style="font-size:11px;font-weight:700;color:#94a3b8">📎 붙임파일:</span>
        </div>` : ''}

        <!-- ④ 하단 닫기 버튼 -->
        <div style="display:flex;justify-content:flex-end;padding-top:14px;border-top:1px solid #e2e8f0;margin-top:4px;flex-shrink:0">
            <button onclick="closeDetailPopup()" style="padding:8px 28px;border-radius:10px;background:#334155;color:white;font-size:13px;font-weight:700;border:none;cursor:pointer;transition:background 0.15s" onmouseover="this.style.background='#1e293b'" onmouseout="this.style.background='#334155'">닫기</button>
        </div>

    </div>`;

    // 주요 URL 버튼 (DOM으로 안전하게 생성)
    const urlBox = container.querySelector('#exp-urlLinks');
    if (urlBox) {
        const rawUrls = (p['주요URL'] || p['주요 URL'] || '').trim();
        if (rawUrls) {
            rawUrls.split('\n').forEach(line => {
                line = line.trim(); if (!line) return;
                const parts = line.split('|');
                const name = parts[0].trim();
                const url  = (parts[1] || '').trim();
                const finalUrl = url || (name.startsWith('http') ? name : '');
                const btn = document.createElement('button');
                btn.textContent = '🔗 ' + name;
                btn.className = 'text-xs font-bold bg-amber-400 hover:bg-amber-300 text-amber-900 px-3 py-1.5 rounded-lg border border-amber-300 transition-all shadow-sm';
                if (finalUrl) btn.addEventListener('click', e => { e.stopPropagation(); window.open(finalUrl, '_blank', 'noopener,noreferrer'); });
                urlBox.appendChild(btn);
            });
        }
    }

    // 붙임파일 버튼 (DOM으로 안전하게 생성)
    const attBox = container.querySelector('#exp-attachments');
    if (attBox && rawAtt) {
        rawAtt.split('\n').forEach(line => {
            line = line.trim(); if (!line) return;
            const parts = line.split('|');
            const name = parts[0].trim();
            const url  = (parts[1] || '').trim();
            const finalUrl = url || (name.startsWith('http') ? name : '');
            const btn = document.createElement('button');
            btn.textContent = '📄 ' + name;
            btn.className = 'text-xs font-semibold text-[#1E3A8A] bg-slate-100 hover:bg-[#1E3A8A] hover:text-white px-3 py-1.5 rounded-full border border-slate-200 transition-all cursor-pointer';
            if (finalUrl) btn.addEventListener('click', e => { e.stopPropagation(); window.open(finalUrl, '_blank', 'noopener,noreferrer'); });
            else { btn.style.opacity = '0.5'; btn.style.cursor = 'default'; }
            attBox.appendChild(btn);
        });
    }
}

function renderExpandedDetail(idx) {
    const panel = document.getElementById('expandedPanel');
    if (!panel) return;
    _renderDetailContent(idx, panel);
}

// ── 일반 모드 팝업 (확장 패널과 동일한 리치 뷰) ────────────────
function openDetailPopup(idx) {
    const body = document.getElementById('detailPopupBody');
    if (!body) return;
    _renderDetailContent(idx, body);
    document.getElementById('detailPopupOverlay').style.display = 'flex';
}

function closeDetailPopup() {
    document.getElementById('detailPopupOverlay').style.display = 'none';
}

// ── 오프라인 배너 ─────────────────────────────────────────
// ── ESC 키로 모달 닫기 ──────────────────────────────────────
document.addEventListener('keydown', (e) => {
    if (e.key !== 'Escape') return;
    const detail  = document.getElementById('detailPopupOverlay');
    const routine = document.getElementById('routineModalOverlay');
    const modal   = document.getElementById('modalOverlay');
    const urlSetup= document.getElementById('urlSetupOverlay');
    if (detail  && detail.style.display  === 'flex') { closeDetailPopup();  return; }
    if (routine && routine.style.display === 'flex') { closeRoutineModal(); return; }
    if (modal   && modal.style.display   === 'flex') { closeModal();        return; }
    if (urlSetup&& urlSetup.style.display=== 'flex') { closeUrlSetup();     return; }
});

function showOfflineBanner() {
    const el = document.getElementById('offlineBanner');
    if (el) el.style.display = 'flex';
}
function hideOfflineBanner() {
    const el = document.getElementById('offlineBanner');
    if (el) el.style.display = 'none';
}

// ── 확장 모드 패널 리사이징 (Draggable Splitter) ──────────────────
document.addEventListener('DOMContentLoaded', () => {
    const resizer = document.getElementById('panelDragResizer');
    const dualOuterWrap = document.getElementById('dualOuterWrap');
    if (!resizer || !dualOuterWrap) return;

    let isResizing = false;

    resizer.addEventListener('mousedown', (e) => {
        if (!document.body.classList.contains('dual-mode')) return;
        isResizing = true;
        resizer.classList.add('active');
        document.body.style.cursor = 'col-resize';
        // 텍스트 선택 방지
        document.body.style.userSelect = 'none';
    });

    document.addEventListener('mousemove', (e) => {
        if (!isResizing) return;
        
        // 브라우저 전체 폭 기준으로 분할 비율 계산
        const w = window.innerWidth;
        // 마우스 X좌표가 차지하는 비율 계산 (최소 20%, 최대 80% 제한)
        let newLeftWidth = (e.clientX / w) * 100;
        if (newLeftWidth < 20) newLeftWidth = 20;
        if (newLeftWidth > 80) newLeftWidth = 80;

        const newRightWidth = 100 - newLeftWidth;

        // CSS 변수에 할당 (style.css에서 var(--left-width) 등으로 받음)
        dualOuterWrap.style.setProperty('--left-width', newLeftWidth + 'fr');
        dualOuterWrap.style.setProperty('--right-width', newRightWidth + 'fr');
    });

    document.addEventListener('mouseup', () => {
        if (isResizing) {
            isResizing = false;
            resizer.classList.remove('active');
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
        }
    });
});
