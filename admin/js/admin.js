/**
 * OfficeMaker 관리자 모듈
 */

// ============ 전역 상태 ============
const state = {
    jwtToken: '',
    userInfo: null,
    templates: [],
    currentTemplate: null,
    currentSlide: null,
    slides: [],
    objects: [],
    selectedObject: null,
    selectedObjects: [],
    clipboard: [],
    // 마퀴(영역) 선택
    isMarqueeSelecting: false,
    marqueeStart: null,
    // 멀티 드래그
    multiDragOffsets: [],
    fonts: [],
    slideMeta: {
        content_type: 'body',
        layout: '',
        has_title: false,
        has_governance: false,
        description_count: 0,
    },
    isDragging: false,
    isResizing: false,
    dragOffset: { x: 0, y: 0 },
    resizeDir: '',
    resizeStart: { x: 0, y: 0, w: 0, h: 0, ox: 0, oy: 0 },
    // 도형 드로잉 모드
    isDrawing: false,
    drawShapeType: null,
    drawChartType: null,
    drawStart: null,
    // 격자/눈금자 (순환: 50 → 25 → 12 → 0(OFF))
    gridLevel: 0,  // 0=50px, 1=25px, 2=12px, 3=OFF
    gridSize: 50,
    // 레이아웃 프리셋 변경 감지
    _previousLayout: '',
    // 최근 사용 색상 (최대 5개)
    recentColors: [],
    // 스킬 관리
    activeTab: 'templates', // 'templates' | 'skills'
    skills: [],
    currentSkill: null,
};

// ============ 슬라이드 크기 관리 ============
function getCanvasSize(slideSize) {
    const sizes = {
        '16:9': { w: 960, h: 540 },
        '4:3':  { w: 960, h: 720 },
        'A4':   { w: 960, h: 665 },
    };
    return sizes[slideSize] || sizes['16:9'];
}

function getTemplateCanvasSize() {
    if (!state.currentTemplate) return { w: 960, h: 540 };
    return getCanvasSize(state.currentTemplate.slide_size || '16:9');
}

// ============ 초기화 ============
const _SOLUTION = window.__SOLUTION_NAME__ || 'OfficeMaker';

$(document).ready(function () {
    // 솔루션명 동적 반영
    document.title = _SOLUTION + ' - 관리자';
    $('.app-title-admin').text(_SOLUTION + ' 관리자');

    // URL에서 JWT 추출
    const pathParts = window.location.pathname.split('/');
    // /admin 경로에서 JWT를 URL에서 가져옴
    const urlParams = new URLSearchParams(window.location.search);
    state.jwtToken = urlParams.get('token') || pathParts[1] || '';

    if (!state.jwtToken || state.jwtToken === 'admin') {
        // JWT가 없으면 로컬스토리지에서 확인 (임시 - 실제로는 URL에서 가져와야 함)
        state.jwtToken = localStorage.getItem('officemaker_jwt') || '';
        if (!state.jwtToken) {
            showLoginPrompt();
            return;
        }
    }

    verifyToken();
});

function showLoginPrompt() {
    const token = prompt('관리자 JWT 토큰을 입력하세요:');
    if (token) {
        state.jwtToken = token;
        verifyToken();
    }
}

async function verifyToken() {
    try {
        const res = await apiGet('/api/auth/verify/' + state.jwtToken);
        if (!res.is_admin) {
            alert('관리자 권한이 필요합니다.');
            return;
        }
        state.userInfo = res.user;
        $('#userInfo').text(`${res.user.nm} (${res.user.dp})`);
        init();
    } catch (e) {
        alert('인증 실패: ' + (e.message || '토큰이 유효하지 않습니다'));
        showLoginPrompt();
    }
}

async function init() {
    initShapePicker();
    await loadFonts();
    await loadTemplates();
}

// ============ API 헬퍼 ============
function apiUrl(path) {
    if (path.startsWith('/api/auth/')) return path;
    return '/' + state.jwtToken + path;
}

async function apiGet(path) {
    const res = await fetch(apiUrl(path));
    if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.detail || 'API 오류');
    }
    return res.json();
}

async function apiPost(path, data) {
    const res = await fetch(apiUrl(path), {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
    });
    if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.detail || 'API 오류');
    }
    return res.json();
}

async function apiPut(path, data) {
    const res = await fetch(apiUrl(path), {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data),
    });
    if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.detail || 'API 오류');
    }
    return res.json();
}

async function apiDelete(path) {
    const res = await fetch(apiUrl(path), { method: 'DELETE' });
    if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.detail || 'API 오류');
    }
    return res.json();
}

async function apiUpload(path, formData) {
    const res = await fetch(apiUrl(path), {
        method: 'POST',
        body: formData,
    });
    if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.detail || 'API 오류');
    }
    return res.json();
}

// ============ 폰트 로드 ============
async function loadFonts() {
    try {
        const res = await apiGet('/api/fonts');
        state.fonts = res.fonts || [];
        updateFontSelector();
        loadWebFonts(state.fonts);
    } catch (e) {
        // 기본 폰트 사용
        state.fonts = [
            { name: 'Arial', family: 'Arial' },
            { name: '맑은 고딕', family: '맑은 고딕' },
        ];
        updateFontSelector();
    }
}

function updateFontSelector() {
    const select = $('#propFont');
    select.empty();
    const defaultFonts = [
        { name: 'Arial', family: 'Arial' },
        { name: '맑은 고딕', family: '맑은 고딕' },
    ];
    const allFonts = [...defaultFonts, ...state.fonts];
    const seen = new Set();
    allFonts.forEach(f => {
        if (!seen.has(f.family)) {
            seen.add(f.family);
            select.append(`<option value="${f.family}" style="font-family:'${f.family}',sans-serif">${f.name}</option>`);
        }
    });
}

// ============ 웹 폰트 동적 로딩 ============
function loadWebFonts(fonts) {
    fonts.forEach(function(font) {
        if (!font.url) return;
        const linkId = 'webfont-' + font.family.replace(/\s+/g, '-');
        if (document.getElementById(linkId)) return;

        if (font.url.includes('fonts.googleapis.com') || font.url.endsWith('.css')) {
            _loadFontCSS(linkId, font.url, font.family);
        } else {
            const fontFace = new FontFace(font.family, 'url(' + font.url + ')');
            fontFace.load().then(function(loaded) {
                document.fonts.add(loaded);
            }).catch(function(err) {
                console.warn('Font load failed:', font.family, err);
            });
        }
    });
}

function _loadFontCSS(linkId, url, family, retries) {
    if (retries === undefined) retries = 2;
    const link = document.createElement('link');
    link.id = linkId;
    link.rel = 'stylesheet';
    link.href = url;
    link.crossOrigin = 'anonymous';
    link.onerror = function() {
        console.warn('Font CSS load failed:', family, '(retries left:', retries, ')');
        link.remove();
        if (retries > 0) {
            setTimeout(function() { _loadFontCSS(linkId, url, family, retries - 1); }, 2000);
        }
    };
    document.head.appendChild(link);
}

function removeWebFont(family) {
    const linkId = 'webfont-' + family.replace(/\s+/g, '-');
    const el = document.getElementById(linkId);
    if (el) el.remove();
}

// ============ 폰트 관리 모달 ============
const FONT_PRESETS = {
    pretendard: {
        name: 'Pretendard',
        family: 'Pretendard',
        url: 'https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/static/pretendard.min.css'
    },
    notosanskr: {
        name: 'Noto Sans KR',
        family: 'Noto Sans KR',
        url: 'https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;600;700&display=swap'
    },
    nanumgothic: {
        name: '나눔고딕',
        family: 'NanumGothic',
        url: 'https://fonts.googleapis.com/css2?family=NanumGothic:wght@400;700;800&display=swap'
    },
    nanummyeongjo: {
        name: '나눔명조',
        family: 'NanumMyeongjo',
        url: 'https://fonts.googleapis.com/css2?family=NanumMyeongjo:wght@400;700;800&display=swap'
    },
    spoqahansansneo: {
        name: 'Spoqa Han Sans Neo',
        family: 'Spoqa Han Sans Neo',
        url: 'https://spoqa.github.io/spoqa-han-sans/css/SpoqaHanSansNeo.css'
    },
    ibmplexsanskr: {
        name: 'IBM Plex Sans KR',
        family: 'IBM Plex Sans KR',
        url: 'https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+KR:wght@300;400;500;600;700&display=swap'
    },
};

function showFontManageModal() {
    renderFontList();
    $('#newFontName').val('');
    $('#newFontFamily').val('');
    $('#newFontUrl').val('');
    $('#fontPreset').val('');
    $('#fontPreviewArea').hide();
    $('#fontManageModal').show();
}

function renderFontList() {
    const container = $('#fontListContainer');
    container.empty();

    if (state.fonts.length === 0) {
        container.html('<div style="padding:16px;text-align:center;color:#999;font-size:13px;">등록된 폰트가 없습니다</div>');
        return;
    }

    state.fonts.forEach(function(font) {
        const badge = font.url
            ? '<span class="font-badge web">웹폰트</span>'
            : '<span class="font-badge system">시스템</span>';
        container.append(`
            <div class="font-list-item">
                <div>
                    <span style="font-size:13px;font-weight:500;font-family:'${escapeHtml(font.family)}',sans-serif;">${escapeHtml(font.name)}</span>
                    <span style="font-size:11px;color:#888;margin-left:6px;">(${escapeHtml(font.family)})</span>
                    ${badge}
                </div>
                <button class="btn" style="padding:2px 10px;font-size:11px;color:#dc3545;border-color:#dc3545;"
                        onclick="deleteFontItem('${font._id}', '${escapeHtml(font.name)}', '${escapeHtml(font.family)}')">삭제</button>
            </div>
        `);
    });
}

function applyFontPreset(key) {
    if (!key || !FONT_PRESETS[key]) {
        $('#fontPreviewArea').hide();
        return;
    }
    const preset = FONT_PRESETS[key];
    $('#newFontName').val(preset.name);
    $('#newFontFamily').val(preset.family);
    $('#newFontUrl').val(preset.url);

    // 프리뷰용 폰트 로딩
    loadWebFonts([preset]);
    $('#fontPreviewText').css('fontFamily', "'" + preset.family + "', sans-serif");
    $('#fontPreviewArea').show();
}

async function addNewFont() {
    const name = $('#newFontName').val().trim();
    const family = $('#newFontFamily').val().trim();
    const url = $('#newFontUrl').val().trim();

    if (!name || !family) {
        showToast('이름과 Font Family는 필수입니다', 'error');
        return;
    }

    try {
        await apiPost('/api/admin/fonts', { name, family, url });
        showToast('폰트가 추가되었습니다', 'success');
        await loadFonts();
        renderFontList();
        $('#newFontName').val('');
        $('#newFontFamily').val('');
        $('#newFontUrl').val('');
        $('#fontPreset').val('');
        $('#fontPreviewArea').hide();
    } catch (e) {
        showToast('폰트 추가 실패: ' + e.message, 'error');
    }
}

async function deleteFontItem(fontId, fontName, fontFamily) {
    if (!confirm('"' + fontName + '" 폰트를 삭제하시겠습니까?')) return;

    try {
        await apiDelete('/api/admin/fonts/' + fontId);
        removeWebFont(fontFamily);
        showToast('폰트가 삭제되었습니다', 'success');
        await loadFonts();
        renderFontList();
    } catch (e) {
        showToast('삭제 실패: ' + e.message, 'error');
    }
}

// ============ 폰트 일괄 변경 ============

function showBulkFontModal() {
    if (!state.currentTemplate) {
        showToast('먼저 템플릿을 선택하세요', 'error');
        return;
    }
    if (state.slides.length === 0) {
        showToast('슬라이드가 없습니다', 'error');
        return;
    }

    // 사용 중인 폰트 수집
    const usedFonts = collectUsedFonts();

    // 소스 폰트 드롭다운
    const fromSelect = $('#bulkFontFrom');
    fromSelect.empty().append('<option value="">전체 폰트</option>');
    usedFonts.forEach(function(f) {
        fromSelect.append('<option value="' + f.family + '">' + f.family + ' (' + f.count + '개)</option>');
    });

    // 타겟 폰트 드롭다운
    const toSelect = $('#bulkFontTo');
    toSelect.empty();
    const defaultFonts = [
        { name: 'Arial', family: 'Arial' },
        { name: '맑은 고딕', family: '맑은 고딕' },
    ];
    const allFonts = [...defaultFonts, ...state.fonts];
    const seen = new Set();
    allFonts.forEach(function(f) {
        if (!seen.has(f.family)) {
            seen.add(f.family);
            toSelect.append('<option value="' + f.family + '" style="font-family:\'' + f.family + '\',sans-serif">' + f.name + '</option>');
        }
    });

    updateBulkFontPreview();
    fromSelect.off('change.bulkfont').on('change.bulkfont', updateBulkFontPreview);
    $('#bulkFontModal').show();
}

function collectUsedFonts() {
    const fontMap = {};
    state.slides.forEach(function(slide) {
        (slide.objects || []).forEach(function(obj) {
            if (obj.obj_type === 'text' && obj.text_style && obj.text_style.font_family) {
                var family = obj.text_style.font_family;
                fontMap[family] = (fontMap[family] || 0) + 1;
            }
        });
    });
    return Object.keys(fontMap).sort().map(function(family) {
        return { family: family, count: fontMap[family] };
    });
}

function updateBulkFontPreview() {
    var fromFont = $('#bulkFontFrom').val();
    var objectCount = 0, slideCount = 0;
    state.slides.forEach(function(slide) {
        var affected = false;
        (slide.objects || []).forEach(function(obj) {
            if (obj.obj_type === 'text' && obj.text_style) {
                if (!fromFont || obj.text_style.font_family === fromFont) {
                    objectCount++;
                    affected = true;
                }
            }
        });
        if (affected) slideCount++;
    });
    $('#bulkFontAffectedCount').text(objectCount);
    $('#bulkFontSlideCount').text(slideCount);
}

async function applyBulkFontChange() {
    var fromFont = $('#bulkFontFrom').val() || null;
    var toFont = $('#bulkFontTo').val();
    if (!toFont) {
        showToast('변경할 폰트를 선택하세요', 'error');
        return;
    }
    var fromLabel = fromFont || '전체 폰트';
    if (!confirm('"' + fromLabel + '" → "' + toFont + '" 변경하시겠습니까?')) return;

    try {
        var res = await apiPut('/api/admin/templates/' + state.currentTemplate._id + '/bulk-font', {
            from_font: fromFont,
            to_font: toFont
        });
        showToast(res.updated_count + '개 텍스트의 폰트가 변경되었습니다', 'success');
        closeModal('bulkFontModal');
        await selectTemplate(state.currentTemplate._id);
    } catch (e) {
        showToast('폰트 변경 실패: ' + e.message, 'error');
    }
}


// ============ 프롬프트 관리 모달 ============
let _promptList = [];
let _editingPromptId = null;
let _availableModels = [];
let _promptCurrentPage = 1;
const _promptPageSize = 5;

async function showPromptManageModal() {
    await Promise.all([loadPromptList(), loadAvailableModels()]);
    showPromptList();
    $('#promptManageModal').show();
}

async function loadAvailableModels() {
    try {
        const res = await apiGet('/api/admin/prompts/models');
        _availableModels = res.models || [];
    } catch (e) {
        _availableModels = [];
    }
}

async function loadPromptList() {
    try {
        const res = await apiGet('/api/admin/prompts');
        _promptList = res.prompts || [];
        _promptCurrentPage = 1;
        renderPromptList();
    } catch (e) {
        showToast('프롬프트 로드 실패: ' + e.message, 'error');
    }
}

function _getModelDisplayName(modelId) {
    if (!modelId) return '-';
    const found = _availableModels.find(function(m) { return m.id === modelId; });
    return found ? found.name : modelId;
}

function renderPromptList() {
    const container = $('#promptListContainer');
    container.empty();

    if (_promptList.length === 0) {
        container.html('<div style="padding:16px;text-align:center;color:#999;font-size:13px;">등록된 프롬프트가 없습니다</div>');
        return;
    }

    const totalPages = Math.ceil(_promptList.length / _promptPageSize);
    if (_promptCurrentPage > totalPages) _promptCurrentPage = totalPages;
    if (_promptCurrentPage < 1) _promptCurrentPage = 1;
    const startIdx = (_promptCurrentPage - 1) * _promptPageSize;
    const pageItems = _promptList.slice(startIdx, startIdx + _promptPageSize);

    pageItems.forEach(function(p) {
        const updatedAt = p.updated_at ? new Date(p.updated_at).toLocaleString('ko-KR') : '';
        const modelName = _getModelDisplayName(p.model);
        container.append(`
            <div class="prompt-list-item" onclick="editPrompt('${p._id}')">
                <div style="flex:1;min-width:0;">
                    <div style="font-size:14px;font-weight:500;color:#333;">${escapeHtml(p.name)}</div>
                    <div style="font-size:11px;color:#999;margin-top:2px;">${escapeHtml(p.description || '')}</div>
                    <div style="font-size:11px;color:#bbb;margin-top:2px;">Key: ${escapeHtml(p.key)} | 수정: ${updatedAt}</div>
                </div>
                <div style="display:flex;align-items:center;gap:8px;margin-left:12px;flex-shrink:0;">
                    <span style="font-size:11px;padding:3px 8px;border-radius:10px;background:#eef2ff;color:#4338ca;font-weight:500;white-space:nowrap;">${escapeHtml(modelName)}</span>
                    <span style="color:#2d5a8e;font-size:12px;white-space:nowrap;">편집 →</span>
                </div>
            </div>
        `);
    });

    // Pagination
    if (totalPages > 1) {
        let paginationHtml = '<div class="prompt-pagination">';
        paginationHtml += `<button class="prompt-page-btn ${_promptCurrentPage === 1 ? 'disabled' : ''}" onclick="_promptGoPage(${_promptCurrentPage - 1})" ${_promptCurrentPage === 1 ? 'disabled' : ''}>
            <svg width="12" height="12" viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"><path d="M7.5 2.5L4 6l3.5 3.5"/></svg>
        </button>`;
        for (let i = 1; i <= totalPages; i++) {
            paginationHtml += `<button class="prompt-page-btn ${i === _promptCurrentPage ? 'active' : ''}" onclick="_promptGoPage(${i})">${i}</button>`;
        }
        paginationHtml += `<button class="prompt-page-btn ${_promptCurrentPage === totalPages ? 'disabled' : ''}" onclick="_promptGoPage(${_promptCurrentPage + 1})" ${_promptCurrentPage === totalPages ? 'disabled' : ''}>
            <svg width="12" height="12" viewBox="0 0 12 12" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round"><path d="M4.5 2.5L8 6l-3.5 3.5"/></svg>
        </button>`;
        paginationHtml += `<span class="prompt-page-info">${_promptCurrentPage} / ${totalPages}</span>`;
        paginationHtml += '</div>';
        container.after().parent().find('.prompt-pagination').remove();
        container.after(paginationHtml);
    } else {
        container.parent().find('.prompt-pagination').remove();
    }
}

function _promptGoPage(page) {
    const totalPages = Math.ceil(_promptList.length / _promptPageSize);
    if (page < 1 || page > totalPages) return;
    _promptCurrentPage = page;
    renderPromptList();
}

function showPromptList() {
    $('#promptListView').show();
    $('#promptEditView').hide().css('display', 'none');
    _editingPromptId = null;
}

function editPrompt(promptId) {
    const p = _promptList.find(item => item._id === promptId);
    if (!p) return;

    _editingPromptId = promptId;
    $('#promptEditName').text(p.name);
    $('#promptEditDesc').text(p.description || '');
    $('#promptEditContent').val(p.content);

    // 모델 셀렉트 박스 구성
    const select = $('#promptEditModel');
    select.empty();
    _availableModels.forEach(function(m) {
        const selected = m.id === (p.model || '') ? ' selected' : '';
        select.append(`<option value="${escapeHtml(m.id)}"${selected}>${escapeHtml(m.name)} - ${escapeHtml(m.description)}</option>`);
    });

    $('#promptListView').hide();
    $('#promptEditView').css('display', 'flex').show();
}

async function savePrompt() {
    if (!_editingPromptId) return;
    const content = $('#promptEditContent').val();
    const model = $('#promptEditModel').val();

    try {
        await apiPut('/api/admin/prompts/' + _editingPromptId, { content, model });
        showToast('프롬프트가 저장되었습니다', 'success');
        await loadPromptList();
        showPromptList();
    } catch (e) {
        showToast('저장 실패: ' + e.message, 'error');
    }
}

async function resetPrompt() {
    if (!_editingPromptId) return;
    if (!confirm('프롬프트를 기본값으로 복원하시겠습니까? 현재 수정 내용이 사라집니다.')) return;

    try {
        const res = await apiPost('/api/admin/prompts/' + _editingPromptId + '/reset', {});
        $('#promptEditContent').val(res.content);
        if (res.model) {
            $('#promptEditModel').val(res.model);
        }
        showToast('기본값으로 복원되었습니다', 'success');
        await loadPromptList();
    } catch (e) {
        showToast('복원 실패: ' + e.message, 'error');
    }
}


// ============ 템플릿 관리 ============
async function loadTemplates() {
    try {
        const res = await apiGet('/api/admin/templates');
        state.templates = res.templates || [];
        renderTemplateList();
    } catch (e) {
        showToast('템플릿 로드 실패', 'error');
    }
}

function renderTemplateList() {
    const list = $('#templateList');
    list.empty();

    if (state.templates.length === 0) {
        list.html('<div style="padding:12px;text-align:center;color:#999;font-size:12px;">템플릿이 없습니다</div>');
        updateTemplateComboText();
        return;
    }

    state.templates.forEach(t => {
        const active = state.currentTemplate && state.currentTemplate._id === t._id ? 'active' : '';
        const published = t.is_published;
        const statusBadge = published
            ? '<span class="template-status-badge published" title="사용자 공개 중">공개</span>'
            : '<span class="template-status-badge draft" title="비공개 (작업 중)">비공개</span>';
        list.append(`
            <div class="template-option ${active}" onclick="selectTemplate('${t._id}')" oncontextmenu="showTemplateContextMenu(event, '${t._id}')">
                <div class="option-title-row">
                    <span class="option-title">${escapeHtml(t.name)}</span>
                    ${statusBadge}
                </div>
                <div class="option-meta">${t.description ? escapeHtml(t.description) : ''}</div>
            </div>
        `);
    });

    updateTemplateComboText();
}

function updateTemplateComboText() {
    if (state.currentTemplate) {
        $('#templateComboText').text(state.currentTemplate.name);
    } else {
        $('#templateComboText').text('템플릿을 선택하세요');
    }
}

function toggleTemplateDropdown() {
    const dd = $('#templateComboDropdown');
    dd.toggleClass('open');
}

// 콤보박스 외부 클릭 시 닫기
$(document).on('click', function(e) {
    if (!$(e.target).closest('#templateCombo').length) {
        $('#templateComboDropdown').removeClass('open');
    }
});

function showNewTemplateModal() {
    $('#newTemplateName').val('');
    $('#newTemplateDesc').val('');
    $('#importPptxFile').val('');
    $('#importFileLabel').text('클릭하여 .pptx 파일을 선택하세요');
    $('#importFileDropZone').css('border-color', '#ccc');
    $('#importProgress').hide();
    $('#importResult').hide();
    $('#btnCreateTemplate').prop('disabled', false).text('생성');
    // 슬라이드 크기 선택 초기화 (16:9 기본)
    $('.size-option').removeClass('selected');
    $('.size-option[data-size="16:9"]').addClass('selected');
    $('#newTemplateModal').show();
}

// 슬라이드 크기 선택 클릭 핸들러
$(document).on('click', '.size-option', function() {
    $('.size-option').removeClass('selected');
    $(this).addClass('selected');
});

function closeModal(id) {
    $('#' + id).hide();
}

function onImportFileSelected(event) {
    const file = event.target.files[0];
    if (!file) {
        $('#importFileLabel').text('클릭하여 .pptx 파일을 선택하세요');
        $('#importFileDropZone').css('border-color', '#ccc');
        return;
    }

    if (!file.name.toLowerCase().endsWith('.pptx')) {
        showToast('.pptx 파일만 가져올 수 있습니다', 'error');
        event.target.value = '';
        return;
    }

    $('#importFileLabel').html('<strong>' + file.name + '</strong><br><span style="font-size:12px;color:#666;">' + (file.size / 1024 / 1024).toFixed(1) + ' MB</span>');
    $('#importFileDropZone').css('border-color', '#2d5a8e');

    // 파일명에서 템플릿 이름 자동 채우기 (비어있는 경우)
    if (!$('#newTemplateName').val().trim()) {
        const nameWithoutExt = file.name.replace(/\.pptx$/i, '');
        $('#newTemplateName').val(nameWithoutExt);
    }
}

async function createTemplate() {
    const name = $('#newTemplateName').val().trim();
    if (!name) {
        alert('템플릿 이름을 입력하세요');
        return;
    }

    const fileInput = document.getElementById('importPptxFile');
    const file = fileInput.files && fileInput.files[0];

    if (file) {
        // PPTX 파일이 있으면 import 방식
        $('#btnCreateTemplate').prop('disabled', true).text('처리 중...');
        $('#importProgress').show();
        $('#importResult').hide();
        $('#importProgressText').text('PPTX 파일 분석 중...');

        const formData = new FormData();
        formData.append('file', file);
        formData.append('template_name', name);

        try {
            const res = await apiUpload('/api/admin/templates/import-pptx', formData);

            $('#importProgress').hide();
            $('#importResult').show();

            const classInfo = (res.classification || []).map(c => {
                const typeNames = {
                    'title_slide': '타이틀',
                    'toc': '목차',
                    'section_divider': '간지',
                    'body': '본문',
                    'closing': '마무리'
                };
                return (typeNames[c.content_type] || c.content_type);
            });

            $('#importResultText').html(
                '<strong>✓ 템플릿 생성 완료!</strong><br>' +
                '슬라이드 ' + res.slides_count + '개 분석 완료<br>' +
                '<span style="font-size:12px;color:#555;">구성: ' + classInfo.join(' → ') + '</span>'
            );

            await loadTemplates();

            setTimeout(() => {
                closeModal('newTemplateModal');
                selectTemplate(res.template_id);
                showToast('PPTX에서 템플릿이 생성되었습니다 (' + res.slides_count + '개 슬라이드)', 'success');
            }, 1500);

        } catch (e) {
            $('#importProgress').hide();
            showToast('PPTX 가져오기 실패: ' + e.message, 'error');
        } finally {
            $('#btnCreateTemplate').prop('disabled', false).text('생성');
        }
    } else {
        // 빈 템플릿 생성
        try {
            const selectedSize = $('.size-option.selected').data('size') || '16:9';
            const res = await apiPost('/api/admin/templates', {
                name: name,
                description: $('#newTemplateDesc').val().trim(),
                slide_size: selectedSize,
            });
            closeModal('newTemplateModal');
            showToast('템플릿이 생성되었습니다', 'success');
            await loadTemplates();
            selectTemplate(res.template._id);
        } catch (e) {
            showToast('생성 실패: ' + e.message, 'error');
        }
    }
}

// PPTX 드래그 앤 드롭
$(document).on('dragover', '#importFileDropZone', function(e) {
    e.preventDefault();
    $(this).addClass('drag-over');
});
$(document).on('dragleave', '#importFileDropZone', function(e) {
    e.preventDefault();
    $(this).removeClass('drag-over');
});
$(document).on('drop', '#importFileDropZone', function(e) {
    e.preventDefault();
    $(this).removeClass('drag-over');
    const files = e.originalEvent.dataTransfer.files;
    if (files.length > 0) {
        const fileInput = document.getElementById('importPptxFile');
        fileInput.files = files;
        onImportFileSelected({ target: fileInput });
    }
});

async function selectTemplate(templateId) {
    if (state.isDrawing) cancelDrawMode();
    $('#templateComboDropdown').removeClass('open');
    try {
        const res = await apiGet('/api/admin/templates/' + templateId);
        state.currentTemplate = res.template;
        state.slides = res.slides || [];
        state.currentSlide = null;
        state.objects = [];
        state.selectedObject = null;
        state.selectedObjects = [];
        removeGroupBoundingBox();

        renderTemplateList();
        renderSlideList();
        showEditor();

        $('#slideSection').show();
        $('#btnDeleteTemplate').show();
        $('#btnBulkFont').show();
        $('#btnPublishTemplate').show();
        updatePublishButton();
        updateTemplateNameBar();

        // 캔버스 초기화
        clearCanvas();
        updateBgUI();
        scaleCanvas();

        if (state.slides.length > 0) {
            selectSlide(state.slides[0]._id);
        }
    } catch (e) {
        showToast('템플릿 로드 실패', 'error');
    }
}

function updateBgUI() {
    // 슬라이드별 배경 우선, 없으면 템플릿 배경
    const slideBg = state.currentSlide && (typeof state.currentSlide === 'object')
        ? state.currentSlide.background_image : null;
    const templateBg = state.currentTemplate ? state.currentTemplate.background_image : null;
    const bgImage = slideBg || templateBg;

    if (bgImage) {
        $('#canvasBg').css('background-image', `url(${bgImage})`);
        $('#btnRemoveBg').show();
    } else {
        $('#canvasBg').css('background-image', 'none');
        $('#btnRemoveBg').hide();
    }

    // 슬라이드 배경 해제 버튼 표시
    if (slideBg) {
        $('#btnRemoveSlideBg').show();
    } else {
        $('#btnRemoveSlideBg').hide();
    }
}

async function deleteCurrentTemplate() {
    if (!state.currentTemplate) return;
    if (!confirm(`"${state.currentTemplate.name}" 템플릿을 삭제하시겠습니까?\n관련 슬라이드도 모두 삭제됩니다.`)) return;

    try {
        await apiDelete('/api/admin/templates/' + state.currentTemplate._id);
        state.currentTemplate = null;
        state.slides = [];
        state.currentSlide = null;
        state.objects = [];
        showToast('템플릿이 삭제되었습니다', 'success');
        await loadTemplates();
        hideEditor();
    } catch (e) {
        showToast('삭제 실패', 'error');
    }
}

// ============ 템플릿 게시/이름 수정 ============

function updatePublishButton() {
    if (!state.currentTemplate) return;
    const btn = $('#btnPublishTemplate');
    if (state.currentTemplate.is_published) {
        btn.text('사용자 비공개').removeClass('btn-publish').addClass('btn-unpublish');
        btn.attr('title', '사용자에게 비공개 처리');
    } else {
        btn.text('사용자 적용').removeClass('btn-unpublish').addClass('btn-publish');
        btn.attr('title', '사용자에게 공개 (템플릿 선택 가능)');
    }
}

async function togglePublishTemplate() {
    if (!state.currentTemplate) return;
    const isPublished = state.currentTemplate.is_published;
    const action = isPublished ? '비공개' : '사용자 적용(공개)';
    if (!confirm(`이 템플릿을 ${action} 처리하시겠습니까?`)) return;

    try {
        const res = await apiPut('/api/admin/templates/' + state.currentTemplate._id + '/publish');
        state.currentTemplate.is_published = res.is_published;
        // 목록 데이터도 갱신
        const tpl = state.templates.find(t => t._id === state.currentTemplate._id);
        if (tpl) tpl.is_published = res.is_published;
        updatePublishButton();
        renderTemplateList();
        const msg = res.is_published ? '사용자에게 공개되었습니다' : '비공개 처리되었습니다';
        showToast(msg, 'success');
    } catch (e) {
        showToast('상태 변경 실패', 'error');
    }
}

// 템플릿 이름 바 표시/숨김
function updateTemplateNameBar() {
    if (!state.currentTemplate) {
        $('#templateNameBar').hide();
        return;
    }
    $('#templateNameBar').show();
    $('#templateNameLabel').text(state.currentTemplate.name);
    // 슬라이드 크기 배지 표시
    const slideSize = state.currentTemplate.slide_size || '16:9';
    const sz = getCanvasSize(slideSize);
    $('#templateSizeBadge').text(slideSize + ' (' + sz.w + 'x' + sz.h + ')');
    $('#templateNameDisplay').show();
    $('#templateNameEdit').hide();
}

function startEditTemplateName() {
    if (!state.currentTemplate) return;
    $('#templateNameDisplay').hide();
    $('#templateNameEdit').show();
    $('#templateNameInput').val(state.currentTemplate.name).focus().select();
}

function cancelEditTemplateName() {
    $('#templateNameEdit').hide();
    $('#templateNameDisplay').show();
}

function handleTemplateNameKey(e) {
    if (e.key === 'Enter') {
        e.preventDefault();
        saveTemplateName();
    } else if (e.key === 'Escape') {
        cancelEditTemplateName();
    }
}

async function saveTemplateName() {
    if (!state.currentTemplate) return;
    const newName = $('#templateNameInput').val().trim();
    if (!newName || newName === state.currentTemplate.name) {
        cancelEditTemplateName();
        return;
    }

    try {
        await apiPut('/api/admin/templates/' + state.currentTemplate._id, { name: newName });
        state.currentTemplate.name = newName;
        const tpl = state.templates.find(t => t._id === state.currentTemplate._id);
        if (tpl) tpl.name = newName;
        renderTemplateList();
        updateTemplateComboText();
        updateTemplateNameBar();
        showToast('템플릿 이름이 변경되었습니다', 'success');
    } catch (e) {
        showToast('이름 변경 실패', 'error');
    }
}

function showTemplateContextMenu(e, templateId) {
    e.preventDefault();
    e.stopPropagation();
    $('.template-context-menu').remove();

    const tpl = state.templates.find(t => t._id === templateId);
    if (!tpl) return;

    const publishLabel = tpl.is_published ? '비공개 전환' : '사용자 적용(공개)';
    const menu = $(`
        <div class="template-context-menu" style="position:fixed;left:${e.clientX}px;top:${e.clientY}px;z-index:10000;">
            <div class="ctx-item" onclick="startEditTemplateNameById('${templateId}')">이름 수정</div>
            <div class="ctx-item" onclick="contextTogglePublish('${templateId}')">${publishLabel}</div>
        </div>
    `);
    $('body').append(menu);
    $(document).one('mousedown', function () {
        $('.template-context-menu').remove();
    });
}

function startEditTemplateNameById(templateId) {
    $('.template-context-menu').remove();
    // 먼저 해당 템플릿 선택
    if (!state.currentTemplate || state.currentTemplate._id !== templateId) {
        selectTemplate(templateId).then(() => {
            startEditTemplateName();
        });
    } else {
        startEditTemplateName();
    }
}

async function contextTogglePublish(templateId) {
    $('.template-context-menu').remove();
    const tpl = state.templates.find(t => t._id === templateId);
    if (!tpl) return;

    try {
        const res = await apiPut('/api/admin/templates/' + templateId + '/publish');
        tpl.is_published = res.is_published;
        if (state.currentTemplate && state.currentTemplate._id === templateId) {
            state.currentTemplate.is_published = res.is_published;
            updatePublishButton();
        }
        renderTemplateList();
        const msg = res.is_published ? '사용자에게 공개되었습니다' : '비공개 처리되었습니다';
        showToast(msg, 'success');
    } catch (e) {
        showToast('상태 변경 실패', 'error');
    }
}

// ============ 슬라이드 관리 ============
function renderSlideList() {
    const list = $('#slideList');
    list.empty();

    const sz = getTemplateCanvasSize();
    const aspectRatio = sz.w + '/' + sz.h;

    state.slides.forEach((s, idx) => {
        const active = state.currentSlide && state.currentSlide._id === s._id ? 'active' : '';
        const meta = s.slide_meta || {};
        const typeLabel = { title_slide: '타이틀', toc: '목차', section_divider: '간지', body: '본문', closing: '마무리' }[meta.content_type] || '본문';

        list.append(`
            <div class="sidebar-item slide-item ${active}"
                 draggable="true"
                 data-slide-id="${s._id}"
                 data-slide-idx="${idx}"
                 onclick="selectSlide('${s._id}')">
                <div class="slide-mini-preview" id="miniPreview_${s._id}" style="aspect-ratio:${aspectRatio};"></div>
                <div class="item-title">슬라이드 ${idx + 1}</div>
                <div class="item-meta">${typeLabel} | 오브젝트 ${(s.objects || []).length}개</div>
            </div>
        `);
    });

    // 썸네일 렌더링
    state.slides.forEach(s => renderSlideThumbnail(s));

    // 드래그&드롭 이벤트 바인딩
    bindSlideDragEvents();
}

// ============ 슬라이드 드래그&드롭 순서 변경 ============
let dragSrcIdx = null;

function bindSlideDragEvents() {
    const items = document.querySelectorAll('#slideList .slide-item');

    items.forEach(item => {
        item.addEventListener('dragstart', function (e) {
            dragSrcIdx = parseInt(this.dataset.slideIdx);
            this.classList.add('dragging');
            e.dataTransfer.effectAllowed = 'move';
        });

        item.addEventListener('dragend', function () {
            this.classList.remove('dragging');
            document.querySelectorAll('.slide-item.drag-over').forEach(el => el.classList.remove('drag-over'));
            dragSrcIdx = null;
        });

        item.addEventListener('dragover', function (e) {
            e.preventDefault();
            e.dataTransfer.dropEffect = 'move';
            // 자기 자신 제외
            const targetIdx = parseInt(this.dataset.slideIdx);
            if (targetIdx !== dragSrcIdx) {
                document.querySelectorAll('.slide-item.drag-over').forEach(el => el.classList.remove('drag-over'));
                this.classList.add('drag-over');
            }
        });

        item.addEventListener('dragleave', function () {
            this.classList.remove('drag-over');
        });

        item.addEventListener('drop', function (e) {
            e.preventDefault();
            this.classList.remove('drag-over');
            const targetIdx = parseInt(this.dataset.slideIdx);

            if (dragSrcIdx === null || dragSrcIdx === targetIdx) return;

            // 배열에서 순서 변경
            const moved = state.slides.splice(dragSrcIdx, 1)[0];
            state.slides.splice(targetIdx, 0, moved);

            // order 필드 업데이트 후 서버 반영
            updateSlideOrders();
            renderSlideList();
        });
    });
}

async function updateSlideOrders() {
    // 각 슬라이드의 order를 새 위치에 맞게 업데이트
    for (let i = 0; i < state.slides.length; i++) {
        const slide = state.slides[i];
        const newOrder = i + 1;
        if (slide.order !== newOrder) {
            slide.order = newOrder;
            try {
                await apiPut('/api/admin/slides/' + slide._id, { order: newOrder });
            } catch (e) {
                // 에러 무시 - 다음에 저장 시 반영
            }
        }
    }
    showToast('슬라이드 순서가 변경되었습니다', 'success');
}

function renderSlideThumbnail(slide) {
    const container = document.getElementById('miniPreview_' + slide._id);
    if (!container) return;

    container.innerHTML = '';

    // 배경 이미지 (슬라이드별 우선)
    const bgImage = slide.background_image || (state.currentTemplate && state.currentTemplate.background_image);
    if (bgImage) {
        container.style.backgroundImage = `url(${bgImage})`;
        container.style.backgroundSize = 'cover';
        container.style.backgroundPosition = 'center';
    } else {
        container.style.backgroundImage = 'none';
        container.style.background = '#f0f0f0';
    }

    const objects = slide.objects || [];
    if (objects.length === 0) return;

    // 썸네일 크기 기준 스케일 계산 (캔버스 크기 → 썸네일 크기)
    const sz = getTemplateCanvasSize();
    const thumbW = container.clientWidth || 160;
    const thumbH = container.clientHeight || 90;
    const scaleX = thumbW / sz.w;
    const scaleY = thumbH / sz.h;

    objects.forEach(obj => {
        const el = document.createElement('div');
        el.style.position = 'absolute';
        el.style.left = (obj.x * scaleX) + 'px';
        el.style.top = (obj.y * scaleY) + 'px';
        el.style.width = (obj.width * scaleX) + 'px';
        el.style.height = (obj.height * scaleY) + 'px';
        el.style.overflow = 'hidden';
        el.style.pointerEvents = 'none';

        if (obj.obj_type === 'image' && obj.image_url) {
            const img = document.createElement('img');
            img.src = obj.image_url;
            img.style.width = '100%';
            img.style.height = '100%';
            img.style.objectFit = obj.image_fit || 'contain';
            el.appendChild(img);
        } else if (obj.obj_type === 'shape') {
            el.style.overflow = 'visible';
            const svgStr = createShapeSVG(obj);
            // viewBox 유지, 크기를 썸네일에 맞게 조정
            el.innerHTML = svgStr
                .replace(/width="[^"]*"/, `width="${obj.width * scaleX}"`)
                .replace(/height="[^"]*"/, `height="${obj.height * scaleY}"`);
        } else if (obj.obj_type === 'text') {
            const style = obj.text_style || {};
            const fontSize = Math.max(1, Math.round((style.font_size || 16) * scaleX));
            el.style.fontSize = fontSize + 'px';
            el.style.lineHeight = '1.2';
            el.style.fontFamily = style.font_family || 'Arial';
            el.style.color = style.color || '#000';
            el.style.fontWeight = style.bold ? 'bold' : 'normal';
            el.style.fontStyle = style.italic ? 'italic' : 'normal';
            el.style.textAlign = style.align || 'left';
            el.textContent = obj.text_content || '';
        }

        container.appendChild(el);
    });
}

async function addSlide() {
    if (!state.currentTemplate) return;
    try {
        const res = await apiPost('/api/admin/slides', {
            template_id: state.currentTemplate._id,
            objects: [],
            slide_meta: { content_type: 'body', layout: '', has_title: false, has_governance: false, description_count: 0 },
        });
        state.slides.push(res.slide);
        renderSlideList();
        selectSlide(res.slide._id);
        showToast('슬라이드가 추가되었습니다', 'success');
    } catch (e) {
        showToast('추가 실패', 'error');
    }
}

function selectSlide(slideId) {
    if (state.isDrawing) cancelDrawMode();
    const slide = state.slides.find(s => s._id === slideId);
    if (!slide) return;

    state.currentSlide = slide;
    state.objects = (slide.objects || []).map(obj => ({ ...obj }));
    state.selectedObject = null;
    state.selectedObjects = [];
    removeGroupBoundingBox();
    state.slideMeta = slide.slide_meta || {
        content_type: 'body',
        layout: '',
        has_title: false,
        has_governance: false,
        description_count: 0,
    };
    state._previousLayout = state.slideMeta.layout || '';

    renderSlideList();
    renderCanvas();
    showPropertiesPanel();
    updateSlideMetaUI();
    $('#btnDeleteSlide').show();
}

/**
 * 부제목/설명 오브젝트를 텍스트 내용의 번호순으로 정렬
 * 예: "1.부제목", "3.부제목", "2.부제목" → 1, 2, 3 순
 */
function sortObjectsByRole(objects) {
    // 텍스트에서 선행 번호 추출 (예: "1.부제목" → 1, "2. 설명" → 2)
    function extractNum(text) {
        const m = (text || '').match(/^(\d+)/);
        return m ? parseInt(m[1]) : 9999;
    }

    // 역할별로 분리
    const subtitles = [];
    const descriptions = [];
    const others = [];

    objects.forEach(obj => {
        if (obj.role === 'subtitle') subtitles.push(obj);
        else if (obj.role === 'description') descriptions.push(obj);
        else others.push(obj);
    });

    // 번호순 정렬
    subtitles.sort((a, b) => extractNum(a.text_content) - extractNum(b.text_content));
    descriptions.sort((a, b) => extractNum(a.text_content) - extractNum(b.text_content));

    // 재조합: 기타 오브젝트 → 부제목 (순서대로) → 설명 (순서대로)
    return [...others, ...subtitles, ...descriptions];
}

async function saveSlide() {
    if (!state.currentSlide) {
        showToast('저장할 슬라이드를 선택하세요', 'error');
        return;
    }

    // 현재 오브젝트에서 텍스트 내용 동기화
    state.objects.forEach(obj => {
        if (obj.obj_type === 'text') {
            const el = document.querySelector(`[data-obj-id="${obj.obj_id}"] .text-content`);
            if (el) {
                obj.text_content = el.innerText || el.textContent;
            }
        }
    });

    // 본문 슬라이드의 부제목/설명을 번호순으로 정렬
    const sortedObjects = sortObjectsByRole(state.objects);
    state.objects = sortedObjects;

    // 오브젝트 role에서 메타 자동 계산
    var auto = calcMetaFromObjects(state.objects);
    state.slideMeta.has_title = auto.has_title;
    state.slideMeta.has_governance = auto.has_governance;
    state.slideMeta.description_count = auto.description_count;

    try {
        await apiPut('/api/admin/slides/' + state.currentSlide._id, {
            objects: sortedObjects,
            slide_meta: state.slideMeta,
        });

        // 로컬 상태 업데이트
        const idx = state.slides.findIndex(s => s._id === state.currentSlide._id);
        if (idx >= 0) {
            state.slides[idx].objects = [...state.objects];
            state.slides[idx].slide_meta = { ...state.slideMeta };
        }

        showToast('슬라이드가 저장되었습니다', 'success');
        renderSlideList();
    } catch (e) {
        showToast('저장 실패: ' + e.message, 'error');
    }
}

async function deleteCurrentSlide() {
    if (!state.currentSlide) return;
    if (!confirm('이 슬라이드를 삭제하시겠습니까?')) return;

    try {
        await apiDelete('/api/admin/slides/' + state.currentSlide._id);
        state.slides = state.slides.filter(s => s._id !== state.currentSlide._id);
        state.currentSlide = null;
        state.objects = [];
        clearCanvas();
        renderSlideList();
        $('#btnDeleteSlide').hide();
        showToast('슬라이드가 삭제되었습니다', 'success');

        if (state.slides.length > 0) {
            selectSlide(state.slides[0]._id);
        }
    } catch (e) {
        showToast('삭제 실패', 'error');
    }
}

// ============ 캔버스 & 오브젝트 ============
function showEditor() {
    $('#emptyState').hide();
    $('#canvas').show();
    $('#toolbar').show();
    scaleCanvas();
    updateGridVisibility();
}

function hideEditor() {
    $('#emptyState').show();
    $('#canvas').hide();
    $('#toolbar').hide();
    $('#slideSection').hide();
    $('#propertiesPanel').hide();
    $('#btnDeleteTemplate').hide();
    $('#btnBulkFont').hide();
    $('#btnRemoveBg').hide();
    $('#btnPublishTemplate').hide();
    $('#templateNameBar').hide();
}

function scaleCanvas() {
    const wrapper = document.getElementById('canvasWrapper');
    const canvas = document.getElementById('canvas');
    if (!wrapper || !canvas || canvas.style.display === 'none') return;

    const sz = getTemplateCanvasSize();
    canvas.style.width = sz.w + 'px';
    canvas.style.height = sz.h + 'px';

    const wrapperW = wrapper.clientWidth - 40;
    const wrapperH = wrapper.clientHeight - 40;
    const scaleX = wrapperW / sz.w;
    const scaleY = wrapperH / sz.h;
    const scale = Math.min(scaleX, scaleY, 1.2);

    canvas.style.transform = `scale(${scale})`;
}

$(window).on('resize', function () {
    scaleCanvas();
});

// ============ 격자 ============
const GRID_LEVELS = [
    { size: 50, cls: 'grid-50', label: '격자 50' },
    { size: 25, cls: 'grid-25', label: '격자 25' },
    { size: 12, cls: 'grid-12', label: '격자 12' },
    { size: 0,  cls: '',        label: '격자 OFF' },
];

function toggleGrid() {
    state.gridLevel = (state.gridLevel + 1) % GRID_LEVELS.length;
    const level = GRID_LEVELS[state.gridLevel];
    state.gridSize = level.size;
    updateGridVisibility();
}

function updateGridVisibility() {
    const grid = $('#canvasGrid');
    const btn = $('#btnGridToggle');
    const level = GRID_LEVELS[state.gridLevel];

    grid.removeClass('visible grid-50 grid-25 grid-12');
    if (level.size > 0) {
        grid.addClass('visible ' + level.cls);
        btn.addClass('grid-active');
    } else {
        btn.removeClass('grid-active');
    }
    // 버튼 라벨 업데이트
    btn.find('.grid-label').text(level.label);
}

function showPropertiesPanel() {
    $('#propertiesPanel').show();
}

function clientToCanvasCoords(clientX, clientY) {
    const canvas = document.getElementById('canvas');
    const rect = canvas.getBoundingClientRect();
    const sz = getTemplateCanvasSize();
    const scaleX = rect.width / sz.w;
    const scaleY = rect.height / sz.h;
    return {
        x: Math.round(Math.max(0, Math.min((clientX - rect.left) / scaleX, sz.w))),
        y: Math.round(Math.max(0, Math.min((clientY - rect.top) / scaleY, sz.h))),
    };
}

function clearCanvas() {
    $('#canvas .canvas-object').remove();
}

function renderCanvas() {
    clearCanvas();
    const canvas = $('#canvas');

    // 배경 이미지 (슬라이드별 우선)
    updateBgUI();

    state.objects.forEach(obj => {
        const el = createObjectElement(obj);
        canvas.append(el);
    });
}

function createObjectElement(obj) {
    const typeClass = obj.obj_type === 'image' ? 'obj-image' :
                      obj.obj_type === 'image_area' ? 'obj-image-area' :
                      obj.obj_type === 'shape' ? 'obj-shape' :
                      obj.obj_type === 'table' ? 'obj-table' :
                      obj.obj_type === 'chart' ? 'obj-chart' : 'obj-text';
    const zIndex = obj.z_index !== undefined ? obj.z_index : 10;
    const div = $('<div>')
        .addClass('canvas-object')
        .addClass(typeClass)
        .attr('data-obj-id', obj.obj_id)
        .css({
            left: obj.x + 'px',
            top: obj.y + 'px',
            width: obj.width + 'px',
            height: obj.height + 'px',
            zIndex: zIndex,
        });

    if (obj.obj_type === 'image' && obj.image_url) {
        const fit = obj.image_fit || 'contain';
        div.append(`<img src="${obj.image_url}" alt="image" style="object-fit:${fit};">`);
    } else if (obj.obj_type === 'image_area') {
        div.append(`<div class="image-area-placeholder">
            <svg width="32" height="32" viewBox="0 0 32 32" fill="none" stroke="currentColor" stroke-width="1.5" opacity="0.5">
                <rect x="2" y="2" width="28" height="28" rx="2" stroke-dasharray="4 3"/>
                <path d="M6 24l7-9 5 6 4-5 4 8" stroke-linecap="round" stroke-linejoin="round"/>
                <circle cx="11" cy="11" r="3"/>
            </svg>
            <span>이미지 영역</span>
        </div>`);
    } else if (obj.obj_type === 'shape') {
        div.append(createShapeSVG(obj));
    } else if (obj.obj_type === 'text') {
        const style = obj.text_style || {};
        const textDiv = $('<div>')
            .addClass('text-content')
            .attr('contenteditable', 'true')
            .css({
                fontFamily: style.font_family || 'Arial',
                fontSize: (style.font_size || 16) + 'px',
                color: style.color || '#000000',
                fontWeight: style.bold ? 'bold' : 'normal',
                fontStyle: style.italic ? 'italic' : 'normal',
                textAlign: style.align || 'left',
            })
            .text(obj.text_content || '텍스트를 입력하세요');
        div.append(textDiv);
    } else if (obj.obj_type === 'table') {
        div.addClass('obj-table');
        const wrapper = $('<div>').addClass('table-wrapper').css({
            width: '100%', height: '100%', overflow: 'hidden',
        });
        wrapper.html(createTableHTML(obj));
        div.append(wrapper);
        // Bind table events after appending to DOM
        setTimeout(() => _bindTableEvents(div, obj), 0);
    } else if (obj.obj_type === 'chart') {
        div.addClass('obj-chart');
        const wrapper = $('<div>').addClass('chart-wrapper').css({
            width: '100%', height: '100%',
        });
        div.append(wrapper);
        // Render chart after DOM append
        setTimeout(() => renderChartInElement(wrapper[0], obj), 50);
    }

    // 리사이즈 핸들
    div.append('<div class="resize-handle se"></div>');
    div.append('<div class="resize-handle sw"></div>');
    div.append('<div class="resize-handle ne"></div>');
    div.append('<div class="resize-handle nw"></div>');

    // 삭제 버튼
    div.append('<button class="delete-btn" onclick="deleteObject(\'' + obj.obj_id + '\')">×</button>');

    // 이벤트 바인딩
    div.on('mousedown', function (e) {
        if ($(e.target).hasClass('resize-handle')) {
            startResize(e, obj.obj_id, e.target);
        } else if (!$(e.target).hasClass('delete-btn')) {
            const isTextContent = obj.obj_type === 'text' && $(e.target).hasClass('text-content');

            // Alt+클릭: 오브젝트 복사 후 복사본 드래그
            if (e.altKey && !isTextContent) {
                const clonedId = duplicateObject(obj.obj_id);
                if (clonedId) {
                    selectObject(clonedId);
                    startDrag(e, clonedId);
                    return;
                }
            }

            // Ctrl+클릭: 멀티셀렉트 토글
            if ((e.ctrlKey || e.metaKey) && !isTextContent) {
                toggleInSelection(obj.obj_id);
                if (isObjectSelected(obj.obj_id) && state.selectedObjects.length > 1) {
                    startGroupDrag(e);
                } else if (state.selectedObjects.length === 1) {
                    startDrag(e, state.selectedObjects[0].obj_id);
                }
                return;
            }

            // 일반 클릭
            if (!isObjectSelected(obj.obj_id)) {
                selectObjectSingle(obj.obj_id);
            } else if (state.selectedObjects.length > 1) {
                // 이미 멀티셀렉트 상태에서 선택된 객체 클릭: 주 선택만 변경
                state.selectedObject = state.objects.find(o => o.obj_id === obj.obj_id);
                updatePropertiesPanel();
            }

            // 드래그 시작
            if (!isTextContent) {
                if (state.selectedObjects.length > 1) {
                    startGroupDrag(e);
                } else {
                    startDrag(e, obj.obj_id);
                }
            }
        }
    });

    return div;
}

function generateObjId() {
    return 'obj_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
}

function getNextZIndex() {
    if (state.objects.length === 0) return 10;
    const maxZ = Math.max(...state.objects.map(o => o.z_index !== undefined ? o.z_index : 10));
    return maxZ + 1;
}

function duplicateObject(objId) {
    const src = state.objects.find(o => o.obj_id === objId);
    if (!src) return null;

    const clone = JSON.parse(JSON.stringify(src));
    clone.obj_id = generateObjId();
    clone.z_index = getNextZIndex();

    state.objects.push(clone);
    const el = createObjectElement(clone);
    $('#canvas').append(el);
    return clone.obj_id;
}

// ============ 복사/붙여넣기 (Ctrl+C / Ctrl+V) ============

function copySelectedObjects() {
    if (state.selectedObjects.length === 0) return;
    state.clipboard = state.selectedObjects.map(obj => JSON.parse(JSON.stringify(obj)));
    const count = state.clipboard.length;
    showToast(`${count}개 오브젝트가 복사되었습니다`, 'info');
}

function pasteObjects() {
    if (state.clipboard.length === 0) return;
    if (!state.currentSlide) {
        showToast('먼저 슬라이드를 선택하세요', 'error');
        return;
    }

    const newIds = [];
    const offset = 20;

    const sz = getTemplateCanvasSize();
    state.clipboard.forEach(src => {
        const clone = JSON.parse(JSON.stringify(src));
        clone.obj_id = generateObjId();
        clone.z_index = getNextZIndex();
        clone.x = Math.min((clone.x || 0) + offset, sz.w - (clone.width || 100));
        clone.y = Math.min((clone.y || 0) + offset, sz.h - (clone.height || 50));

        state.objects.push(clone);
        const el = createObjectElement(clone);
        $('#canvas').append(el);
        newIds.push(clone.obj_id);
    });

    // 붙여넣은 오브젝트들을 선택 상태로
    setSelection(newIds);

    // 다음 붙여넣기 시 추가 오프셋이 적용되도록 clipboard 위치 업데이트
    state.clipboard = state.clipboard.map(obj => {
        const updated = JSON.parse(JSON.stringify(obj));
        updated.x = Math.min((updated.x || 0) + offset, sz.w - (updated.width || 100));
        updated.y = Math.min((updated.y || 0) + offset, sz.h - (updated.height || 50));
        return updated;
    });

    showToast(`${newIds.length}개 오브젝트가 붙여넣어졌습니다`, 'success');
}

function addTextObject() {
    if (!state.currentSlide) {
        showToast('먼저 슬라이드를 선택하세요', 'error');
        return;
    }

    const obj = {
        obj_id: generateObjId(),
        obj_type: 'text',
        x: 100,
        y: 100,
        width: 300,
        height: 60,
        z_index: getNextZIndex(),
        text_content: '텍스트를 입력하세요',
        text_style: {
            font_family: 'Arial',
            font_size: 16,
            color: '#000000',
            bold: false,
            italic: false,
            align: 'left',
        },
        role: null,
        placeholder: null,
    };

    state.objects.push(obj);
    const el = createObjectElement(obj);
    $('#canvas').append(el);
    selectObject(obj.obj_id);
}

function triggerImageUpload() {
    if (!state.currentSlide) {
        showToast('먼저 슬라이드를 선택하세요', 'error');
        return;
    }
    $('#imageUploadInput').click();
}

async function handleImageUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const formData = new FormData();
    formData.append('file', file);

    try {
        const res = await apiUpload('/api/admin/slides/upload-image', formData);

        const obj = {
            obj_id: generateObjId(),
            obj_type: 'image',
            x: 100,
            y: 100,
            width: 200,
            height: 150,
            z_index: getNextZIndex(),
            image_url: res.image_url,
            image_fit: 'contain',
            role: null,
            placeholder: null,
        };

        state.objects.push(obj);
        const el = createObjectElement(obj);
        $('#canvas').append(el);
        selectObject(obj.obj_id);
        showToast('이미지가 추가되었습니다', 'success');
    } catch (e) {
        showToast('이미지 업로드 실패', 'error');
    }

    event.target.value = '';
}

function addImageAreaObject() {
    if (!state.currentSlide) {
        showToast('먼저 슬라이드를 선택하세요', 'error');
        return;
    }

    const obj = {
        obj_id: generateObjId(),
        obj_type: 'image_area',
        x: 100,
        y: 100,
        width: 400,
        height: 300,
        z_index: getNextZIndex(),
        image_fit: 'cover',
        role: null,
        placeholder: null,
    };

    state.objects.push(obj);
    const el = createObjectElement(obj);
    $('#canvas').append(el);
    selectObject(obj.obj_id);
    showToast('이미지 영역이 추가되었습니다', 'success');
}

function uploadBackground() {
    if (!state.currentTemplate) return;
    $('#bgUploadInput').click();
}

async function handleBgUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const formData = new FormData();
    formData.append('file', file);

    try {
        const res = await apiUpload('/api/admin/templates/' + state.currentTemplate._id + '/background', formData);
        state.currentTemplate.background_image = res.image_url;
        updateBgUI();
        showToast('배경 이미지가 설정되었습니다', 'success');
    } catch (e) {
        showToast('배경 업로드 실패', 'error');
    }

    event.target.value = '';
}

async function removeBackground() {
    if (!state.currentTemplate) return;
    if (!confirm('배경 이미지를 제거하시겠습니까?')) return;

    try {
        await apiPut('/api/admin/templates/' + state.currentTemplate._id, {
            background_image: null,
        });
        state.currentTemplate.background_image = null;
        updateBgUI();
        showToast('배경 이미지가 제거되었습니다', 'success');
    } catch (e) {
        showToast('배경 제거 실패', 'error');
    }
}

async function setImageAsBackground() {
    if (!state.selectedObject || state.selectedObject.obj_type !== 'image') return;
    if (!state.currentSlide) return;

    const imageUrl = state.selectedObject.image_url;
    if (!imageUrl) return;

    const slideId = typeof state.currentSlide === 'object' ? state.currentSlide._id : state.currentSlide;

    try {
        await apiPut('/api/admin/slides/' + slideId, {
            background_image: imageUrl,
        });

        // 슬라이드 로컬 상태 업데이트
        const slide = state.slides.find(s => s._id === slideId);
        if (slide) slide.background_image = imageUrl;
        if (typeof state.currentSlide === 'object') state.currentSlide.background_image = imageUrl;

        updateBgUI();

        // 이미지 오브젝트 제거
        const objId = state.selectedObject.obj_id;
        state.objects = state.objects.filter(o => o.obj_id !== objId);
        state.selectedObjects = state.selectedObjects.filter(o => o.obj_id !== objId);
        $(`[data-obj-id="${objId}"]`).remove();
        state.selectedObject = null;
        removeGroupBoundingBox();
        $('#objProperties').hide();
        $('#imageProperties').hide();
        $('#imageAreaProperties').hide();
        $('#textProperties').hide();
        $('#shapeProperties').hide();

        if (slide) renderSlideThumbnail(slide);

        showToast('이 슬라이드의 배경이 설정되었습니다', 'success');
    } catch (e) {
        showToast('배경 설정 실패', 'error');
    }
}

async function removeSlideBackground() {
    if (!state.currentSlide) return;
    const slideId = typeof state.currentSlide === 'object' ? state.currentSlide._id : state.currentSlide;

    try {
        await apiPut('/api/admin/slides/' + slideId, {
            background_image: '',
        });

        const slide = state.slides.find(s => s._id === slideId);
        if (slide) delete slide.background_image;
        if (typeof state.currentSlide === 'object') delete state.currentSlide.background_image;

        updateBgUI();
        if (slide) renderSlideThumbnail(slide);

        showToast('슬라이드 배경이 해제되었습니다', 'success');
    } catch (e) {
        showToast('배경 해제 실패', 'error');
    }
}

function deleteObject(objId) {
    state.objects = state.objects.filter(o => o.obj_id !== objId);
    state.selectedObjects = state.selectedObjects.filter(o => o.obj_id !== objId);
    $(`[data-obj-id="${objId}"]`).remove();

    if (state.selectedObject && state.selectedObject.obj_id === objId) {
        state.selectedObject = state.selectedObjects.length > 0
            ? state.selectedObjects[state.selectedObjects.length - 1]
            : null;
    }
    updateGroupBoundingBox();
    updatePropertiesPanel();
}

function selectObject(objId) {
    selectObjectSingle(objId);
}

// ============ 멀티셀렉트 헬퍼 ============

function isObjectSelected(objId) {
    return state.selectedObjects.some(o => o.obj_id === objId);
}

function clearSelection() {
    state.selectedObjects = [];
    state.selectedObject = null;
    $('.canvas-object').removeClass('selected');
    removeGroupBoundingBox();
    $('#objProperties').hide();
    $('#imageProperties').hide();
    $('#imageAreaProperties').hide();
    $('#textProperties').hide();
    $('#shapeProperties').hide();
}

function selectObjectSingle(objId) {
    const obj = state.objects.find(o => o.obj_id === objId);
    if (!obj) return;

    state.selectedObjects = [obj];
    state.selectedObject = obj;

    $('.canvas-object').removeClass('selected');
    $(`[data-obj-id="${objId}"]`).addClass('selected');
    removeGroupBoundingBox();

    updatePropertiesPanel();
}

function addToSelection(objId) {
    const obj = state.objects.find(o => o.obj_id === objId);
    if (!obj) return;
    if (isObjectSelected(objId)) return;

    state.selectedObjects.push(obj);
    state.selectedObject = obj;
    $(`[data-obj-id="${objId}"]`).addClass('selected');
    updateGroupBoundingBox();
    updatePropertiesPanel();
}

function removeFromSelection(objId) {
    state.selectedObjects = state.selectedObjects.filter(o => o.obj_id !== objId);
    $(`[data-obj-id="${objId}"]`).removeClass('selected');

    if (state.selectedObjects.length > 0) {
        state.selectedObject = state.selectedObjects[state.selectedObjects.length - 1];
    } else {
        state.selectedObject = null;
    }
    updateGroupBoundingBox();
    updatePropertiesPanel();
}

function toggleInSelection(objId) {
    if (isObjectSelected(objId)) {
        removeFromSelection(objId);
    } else {
        addToSelection(objId);
    }
}

function setSelection(objIds) {
    state.selectedObjects = [];
    state.selectedObject = null;
    $('.canvas-object').removeClass('selected');

    objIds.forEach(id => {
        const obj = state.objects.find(o => o.obj_id === id);
        if (obj) {
            state.selectedObjects.push(obj);
            $(`[data-obj-id="${id}"]`).addClass('selected');
        }
    });

    if (state.selectedObjects.length > 0) {
        state.selectedObject = state.selectedObjects[state.selectedObjects.length - 1];
    }
    updateGroupBoundingBox();
    updatePropertiesPanel();
}

function rectsIntersect(ax, ay, aw, ah, bx, by, bw, bh) {
    return !(ax + aw <= bx || bx + bw <= ax || ay + ah <= by || by + bh <= ay);
}

function updateGroupBoundingBox() {
    removeGroupBoundingBox();
    if (state.selectedObjects.length <= 1) return;

    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    state.selectedObjects.forEach(obj => {
        minX = Math.min(minX, obj.x);
        minY = Math.min(minY, obj.y);
        maxX = Math.max(maxX, obj.x + obj.width);
        maxY = Math.max(maxY, obj.y + obj.height);
    });

    const bbox = $('<div id="groupBoundingBox"></div>').css({
        position: 'absolute',
        left: (minX - 3) + 'px',
        top: (minY - 3) + 'px',
        width: (maxX - minX + 6) + 'px',
        height: (maxY - minY + 6) + 'px',
        border: '2px dashed #ff6b35',
        background: 'transparent',
        zIndex: 9,
        pointerEvents: 'none',
        borderRadius: '2px',
    });

    const badge = $('<div class="group-count-badge"></div>')
        .text(state.selectedObjects.length + '개 선택')
        .css({
            position: 'absolute',
            top: '-22px',
            left: '0',
            background: '#ff6b35',
            color: 'white',
            padding: '1px 8px',
            borderRadius: '3px',
            fontSize: '11px',
            fontWeight: '500',
            whiteSpace: 'nowrap',
            pointerEvents: 'none',
        });

    bbox.append(badge);
    $('#canvas').append(bbox);
}

function removeGroupBoundingBox() {
    $('#groupBoundingBox').remove();
}

function deleteSelectedObjects() {
    if (state.selectedObjects.length === 0) return;
    const idsToDelete = state.selectedObjects.map(o => o.obj_id);
    idsToDelete.forEach(id => {
        state.objects = state.objects.filter(o => o.obj_id !== id);
        $(`[data-obj-id="${id}"]`).remove();
    });
    clearSelection();
}

function startGroupDrag(e) {
    state.isDragging = true;
    const canvasRect = document.getElementById('canvas').getBoundingClientRect();
    const sz = getTemplateCanvasSize();
    const scale = sz.w / canvasRect.width;

    state.multiDragOffsets = state.selectedObjects.map(obj => ({
        obj_id: obj.obj_id,
        offsetX: (e.clientX - canvasRect.left) * scale - obj.x,
        offsetY: (e.clientY - canvasRect.top) * scale - obj.y,
    }));
    e.preventDefault();
}

// ============ 드래그 & 리사이즈 ============
function startDrag(e, objId) {
    state.isDragging = true;
    const obj = state.objects.find(o => o.obj_id === objId);
    if (!obj) return;

    const canvasRect = document.getElementById('canvas').getBoundingClientRect();
    state.dragOffset = {
        x: e.clientX - canvasRect.left - obj.x,
        y: e.clientY - canvasRect.top - obj.y,
    };

    e.preventDefault();
}

function startResize(e, objId, handleEl) {
    state.isResizing = true;
    const obj = state.objects.find(o => o.obj_id === objId);
    if (!obj) return;

    selectObject(objId);

    const classList = handleEl.className;
    if (classList.includes('se')) state.resizeDir = 'se';
    else if (classList.includes('sw')) state.resizeDir = 'sw';
    else if (classList.includes('ne')) state.resizeDir = 'ne';
    else if (classList.includes('nw')) state.resizeDir = 'nw';

    state.resizeStart = {
        x: e.clientX,
        y: e.clientY,
        w: obj.width,
        h: obj.height,
        ox: obj.x,
        oy: obj.y,
    };

    e.preventDefault();
    e.stopPropagation();
}

$(document).on('mousemove', function (e) {
    // 도형 드로잉 중: 고스트 요소 업데이트
    if (state.isDrawing && state.drawStart) {
        const coords = clientToCanvasCoords(e.clientX, e.clientY);
        const ghost = $('#drawGhost');
        if (ghost.length === 0) return;

        const x1 = state.drawStart.x;
        const y1 = state.drawStart.y;
        const left = Math.min(x1, coords.x);
        const top = Math.min(y1, coords.y);
        const width = Math.abs(coords.x - x1);
        const height = Math.abs(coords.y - y1);

        ghost.css({ left: left + 'px', top: top + 'px', width: width + 'px', height: height + 'px' });

        const isLine = (state.drawShapeType === 'line' || state.drawShapeType === 'arrow');
        const previewH = isLine ? Math.max(height, 4) : height;
        if (width > 0 && previewH > 0) {
            const ghostObj = {
                width: width, height: previewH, obj_id: 'ghost',
                shape_style: {
                    shape_type: state.drawShapeType,
                    fill_color: isLine ? 'transparent' : '#4A90D9',
                    fill_opacity: isLine ? 0 : 0.3,
                    stroke_color: '#2d5a8e',
                    stroke_width: 2,
                    stroke_dash: 'dashed',
                    border_radius: state.drawShapeType === 'rounded_rectangle' ? 12 : 0,
                    arrow_head: state.drawShapeType === 'arrow' ? 'end' : 'none',
                },
            };
            ghost.html(createShapeSVG(ghostObj));
        }
        return;
    }

    // 마퀴(영역) 선택 드래그 중
    if (state.isMarqueeSelecting && state.marqueeStart) {
        const coords = clientToCanvasCoords(e.clientX, e.clientY);
        const marquee = $('#marqueeSelect');
        if (marquee.length === 0) return;

        const x1 = state.marqueeStart.x;
        const y1 = state.marqueeStart.y;
        const left = Math.min(x1, coords.x);
        const top = Math.min(y1, coords.y);
        const width = Math.abs(coords.x - x1);
        const height = Math.abs(coords.y - y1);

        marquee.css({ left: left + 'px', top: top + 'px', width: width + 'px', height: height + 'px' });

        // 실시간 교차 객체 하이라이트
        $('.canvas-object').removeClass('marquee-hover');
        state.objects.forEach(obj => {
            if (rectsIntersect(left, top, width, height, obj.x, obj.y, obj.width, obj.height)) {
                $(`[data-obj-id="${obj.obj_id}"]`).addClass('marquee-hover');
            }
        });
        return;
    }

    if (state.isDragging && state.selectedObjects.length > 0) {
        const canvasRect = document.getElementById('canvas').getBoundingClientRect();
        const sz = getTemplateCanvasSize();
        const scale = sz.w / canvasRect.width;

        if (state.multiDragOffsets.length > 0) {
            // 그룹 드래그: 모든 선택 객체 이동
            state.multiDragOffsets.forEach(info => {
                const obj = state.objects.find(o => o.obj_id === info.obj_id);
                if (!obj) return;

                let x = (e.clientX - canvasRect.left) * scale - info.offsetX;
                let y = (e.clientY - canvasRect.top) * scale - info.offsetY;

                x = Math.max(0, Math.min(x, sz.w - obj.width));
                y = Math.max(0, Math.min(y, sz.h - obj.height));

                obj.x = Math.round(x);
                obj.y = Math.round(y);
                $(`[data-obj-id="${obj.obj_id}"]`).css({ left: obj.x + 'px', top: obj.y + 'px' });
            });
            updatePositionInputs();
            updateGroupBoundingBox();
        } else if (state.selectedObject) {
            // 단일 드래그
            let x = e.clientX - canvasRect.left - state.dragOffset.x;
            let y = e.clientY - canvasRect.top - state.dragOffset.y;

            x = Math.max(0, Math.min(x, sz.w - state.selectedObject.width));
            y = Math.max(0, Math.min(y, sz.h - state.selectedObject.height));

            state.selectedObject.x = Math.round(x);
            state.selectedObject.y = Math.round(y);

            const el = $(`[data-obj-id="${state.selectedObject.obj_id}"]`);
            el.css({ left: x + 'px', top: y + 'px' });

            updatePositionInputs();
        }
    }

    if (state.isResizing && state.selectedObject) {
        const dx = e.clientX - state.resizeStart.x;
        const dy = e.clientY - state.resizeStart.y;
        const obj = state.selectedObject;

        let newW = state.resizeStart.w;
        let newH = state.resizeStart.h;
        let newX = state.resizeStart.ox;
        let newY = state.resizeStart.oy;

        // 라인/화살표는 최소 높이를 작게 허용
        const isLine = obj.obj_type === 'shape' &&
            (obj.shape_style?.shape_type === 'line' || obj.shape_style?.shape_type === 'arrow');
        const minH = isLine ? 2 : 30;

        switch (state.resizeDir) {
            case 'se':
                newW = Math.max(50, state.resizeStart.w + dx);
                newH = Math.max(minH, state.resizeStart.h + dy);
                break;
            case 'sw':
                newW = Math.max(50, state.resizeStart.w - dx);
                newH = Math.max(minH, state.resizeStart.h + dy);
                newX = state.resizeStart.ox + dx;
                break;
            case 'ne':
                newW = Math.max(50, state.resizeStart.w + dx);
                newH = Math.max(minH, state.resizeStart.h - dy);
                newY = state.resizeStart.oy + dy;
                break;
            case 'nw':
                newW = Math.max(50, state.resizeStart.w - dx);
                newH = Math.max(minH, state.resizeStart.h - dy);
                newX = state.resizeStart.ox + dx;
                newY = state.resizeStart.oy + dy;
                break;
        }

        obj.x = Math.round(newX);
        obj.y = Math.round(newY);
        obj.width = Math.round(newW);
        obj.height = Math.round(newH);

        const el = $(`[data-obj-id="${obj.obj_id}"]`);
        el.css({ left: newX + 'px', top: newY + 'px', width: newW + 'px', height: newH + 'px' });

        // 도형 리사이즈 시 SVG 재렌더
        if (obj.obj_type === 'shape') {
            el.find('svg').remove();
            el.prepend(createShapeSVG(obj));
        }

        updatePositionInputs();
    }
});

$(document).on('mouseup', function (e) {
    // 도형 드로잉 완료
    if (state.isDrawing && state.drawStart) {
        const coords = clientToCanvasCoords(e.clientX, e.clientY);
        const x1 = state.drawStart.x;
        const y1 = state.drawStart.y;

        let left = Math.min(x1, coords.x);
        let top = Math.min(y1, coords.y);
        let width = Math.abs(coords.x - x1);
        let height = Math.abs(coords.y - y1);

        $('#drawGhost').remove();

        // 최소 크기 임계값: 너무 작으면 기본 크기로 생성
        const MIN_SIZE = 20;
        if (width < MIN_SIZE && height < MIN_SIZE) {
            width = 200;
            const isLine = (state.drawShapeType === 'line' || state.drawShapeType === 'arrow');
            height = isLine ? 4 : 150;
            const sz2 = getTemplateCanvasSize();
            left = Math.max(0, Math.min(left - width / 2, sz2.w - width));
            top = Math.max(0, Math.min(top - height / 2, sz2.h - height));
        }

        // 캔버스 경계 제한
        const szBound = getTemplateCanvasSize();
        left = Math.max(0, left);
        top = Math.max(0, top);
        width = Math.min(width, szBound.w - left);
        height = Math.min(height, szBound.h - top);

        if (state.drawChartType) {
            createChartAtRect(state.drawChartType, left, top, width, height);
            state.drawChartType = null;
        } else {
            createShapeAtRect(state.drawShapeType, left, top, width, height);
        }
        cancelDrawMode();
        return;
    }

    // 마퀴 선택 완료
    if (state.isMarqueeSelecting) {
        const coords = clientToCanvasCoords(e.clientX, e.clientY);
        const marquee = $('#marqueeSelect');

        const x1 = state.marqueeStart.x;
        const y1 = state.marqueeStart.y;
        const left = Math.min(x1, coords.x);
        const top = Math.min(y1, coords.y);
        const width = Math.abs(coords.x - x1);
        const height = Math.abs(coords.y - y1);

        marquee.remove();
        $('.canvas-object').removeClass('marquee-hover');

        state.isMarqueeSelecting = false;
        state.marqueeStart = null;

        // 3px 미만 이동이면 단순 클릭 = 선택 해제
        if (width < 3 && height < 3) {
            if (!e.ctrlKey && !e.metaKey) clearSelection();
            return;
        }

        // 교차하는 객체 찾기
        const selectedIds = [];
        state.objects.forEach(obj => {
            if (rectsIntersect(left, top, width, height, obj.x, obj.y, obj.width, obj.height)) {
                selectedIds.push(obj.obj_id);
            }
        });

        if (e.ctrlKey || e.metaKey) {
            selectedIds.forEach(id => addToSelection(id));
        } else {
            setSelection(selectedIds);
        }
        return;
    }

    state.isDragging = false;
    if (state.isResizing && state.selectedObject && state.selectedObject.obj_type === 'chart') {
        const el = $(`[data-obj-id="${state.selectedObject.obj_id}"]`);
        setTimeout(() => renderChartInElement(el.find('.chart-wrapper')[0], state.selectedObject), 50);
    }
    state.isResizing = false;
    state.multiDragOffsets = [];
});

// 키보드 이벤트 처리
$(document).on('keydown', function (e) {
    // ESC: 드로잉 모드 취소 또는 선택 해제
    if (e.key === 'Escape') {
        if (state.isDrawing) {
            cancelDrawMode();
        } else if (state.selectedObjects.length > 0) {
            clearSelection();
        }
        e.preventDefault();
        return;
    }

    // 텍스트 편집 중이면 이후 단축키 무시
    const active = document.activeElement;
    const isEditing = active && (active.isContentEditable || active.tagName === 'INPUT' || active.tagName === 'TEXTAREA');

    // Delete: 선택된 객체 삭제
    if (e.key === 'Delete' && state.selectedObjects.length > 0 && !isEditing) {
        e.preventDefault();
        deleteSelectedObjects();
        return;
    }

    // Ctrl+A: 전체 선택
    if ((e.ctrlKey || e.metaKey) && e.key === 'a' && !isEditing) {
        e.preventDefault();
        setSelection(state.objects.map(o => o.obj_id));
        return;
    }

    // Ctrl+C: 선택된 오브젝트 복사
    if ((e.ctrlKey || e.metaKey) && e.key === 'c' && !isEditing) {
        if (state.selectedObjects.length > 0) {
            e.preventDefault();
            copySelectedObjects();
            return;
        }
    }

    // Ctrl+V: 복사된 오브젝트 붙여넣기
    if ((e.ctrlKey || e.metaKey) && e.key === 'v' && !isEditing) {
        if (state.clipboard.length > 0) {
            e.preventDefault();
            pasteObjects();
            return;
        }
    }

    // 화살표 위/아래: 오브젝트 미선택 시 슬라이드 탐색
    if (state.selectedObjects.length === 0 && !isEditing) {
        if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
            if (!state.currentSlide || !state.slides || state.slides.length === 0) return;
            e.preventDefault();
            var curIdx = state.slides.findIndex(function(s) { return s._id === state.currentSlide._id; });
            if (curIdx === -1) return;
            var newIdx = e.key === 'ArrowUp' ? curIdx - 1 : curIdx + 1;
            if (newIdx < 0 || newIdx >= state.slides.length) return;
            selectSlide(state.slides[newIdx]._id);
            // 선택된 슬라이드가 보이도록 스크롤
            var slideEl = document.querySelectorAll('#slideList .slide-item')[newIdx];
            if (slideEl) slideEl.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
            return;
        }
    }

    // 화살표 키: 선택된 모든 오브젝트 이동
    if (state.selectedObjects.length === 0) return;
    if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].indexOf(e.key) === -1) return;
    if (isEditing) return;

    e.preventDefault();
    const step = e.shiftKey ? 10 : 1;

    const sz = getTemplateCanvasSize();
    state.selectedObjects.forEach(obj => {
        switch (e.key) {
            case 'ArrowLeft':  obj.x = Math.max(0, obj.x - step); break;
            case 'ArrowRight': obj.x = Math.min(sz.w - obj.width, obj.x + step); break;
            case 'ArrowUp':    obj.y = Math.max(0, obj.y - step); break;
            case 'ArrowDown':  obj.y = Math.min(sz.h - obj.height, obj.y + step); break;
        }
        obj.x = Math.round(obj.x);
        obj.y = Math.round(obj.y);
        $(`[data-obj-id="${obj.obj_id}"]`).css({ left: obj.x + 'px', top: obj.y + 'px' });
    });

    updatePositionInputs();
    updateGroupBoundingBox();
});

// 캔버스 외부 클릭 시 드로잉 모드 취소
$(document).on('mousedown', function (e) {
    if (state.isDrawing && !$(e.target).closest('#canvas').length) {
        cancelDrawMode();
    }
});

// 캔버스 클릭 (선택 해제 또는 도형 드로잉 시작)
$('#canvas').on('mousedown', function (e) {
    if (e.target !== this && e.target.id !== 'canvasBg' && e.target.id !== 'canvasGrid') return;

    if (state.isDrawing) {
        // 드로잉 모드: 드래그 시작
        e.preventDefault();
        const coords = clientToCanvasCoords(e.clientX, e.clientY);
        state.drawStart = { x: coords.x, y: coords.y };

        const ghost = $('<div id="drawGhost"></div>').css({
            position: 'absolute',
            left: coords.x + 'px',
            top: coords.y + 'px',
            width: '0px',
            height: '0px',
            zIndex: 50,
            pointerEvents: 'none',
        });
        $('#canvas').append(ghost);
    } else {
        // 일반 모드: 마퀴(영역) 선택 시작
        e.preventDefault();
        const coords = clientToCanvasCoords(e.clientX, e.clientY);
        state.isMarqueeSelecting = true;
        state.marqueeStart = { x: coords.x, y: coords.y };

        const marquee = $('<div id="marqueeSelect"></div>').css({
            position: 'absolute',
            left: coords.x + 'px',
            top: coords.y + 'px',
            width: '0px',
            height: '0px',
            border: '1px dashed #2d5a8e',
            background: 'rgba(45, 90, 142, 0.08)',
            zIndex: 50,
            pointerEvents: 'none',
        });
        $('#canvas').append(marquee);

        // Ctrl 미누름 시 기존 선택 해제
        if (!e.ctrlKey && !e.metaKey) {
            clearSelection();
        }
    }
});

// ============ 속성 패널 ============
function updatePropertiesPanel() {
    const count = state.selectedObjects.length;

    if (count === 0) {
        $('#objProperties').hide();
        $('#imageProperties').hide();
        $('#imageAreaProperties').hide();
        $('#textProperties').hide();
        $('#shapeProperties').hide();
        $('#tableProperties').hide();
        $('#chartProperties').hide();
        $('.multi-select-info').remove();
        return;
    }

    if (count === 1) {
        // 단일 선택: 기존 동작
        $('.multi-select-info').remove();
        const obj = state.selectedObject;
        if (!obj) return;

        $('#objProperties').show();
        $('#propRole').val(obj.role || '');
        $('#propPlaceholder').val(obj.placeholder || '').prop('disabled', false);
        $('#propX').val(Math.round(obj.x));
        $('#propY').val(Math.round(obj.y));
        $('#propW').val(Math.round(obj.width));
        $('#propH').val(Math.round(obj.height));
        $('#propZIndex').val(obj.z_index !== undefined ? obj.z_index : 10);
        // W/H 편집 가능
        $('#propW').prop('disabled', false);
        $('#propH').prop('disabled', false);

        if (obj.obj_type === 'image') {
            $('#imageProperties').show();
            $('#imageAreaProperties').hide();
            $('#textProperties').hide();
            $('#shapeProperties').hide();
            $('#tableProperties').hide();
            $('#chartProperties').hide();
            $('#propImageFit').val(obj.image_fit || 'contain');
        } else if (obj.obj_type === 'image_area') {
            $('#imageProperties').hide();
            $('#imageAreaProperties').show();
            $('#textProperties').hide();
            $('#shapeProperties').hide();
            $('#tableProperties').hide();
            $('#chartProperties').hide();
            $('#propImageAreaFit').val(obj.image_fit || 'cover');
        } else if (obj.obj_type === 'text') {
            $('#imageProperties').hide();
            $('#imageAreaProperties').hide();
            $('#textProperties').show();
            $('#shapeProperties').hide();
            $('#tableProperties').hide();
            $('#chartProperties').hide();
            const style = obj.text_style || {};
            $('#propFont').val(style.font_family || 'Arial');
            $('#propFontSize').val(style.font_size || 16);
            $('#propColor').val(style.color || '#000000');
            $('#propBold').toggleClass('active', !!style.bold).removeClass('mixed');
            $('#propItalic').toggleClass('active', !!style.italic).removeClass('mixed');
            $('#propAlign').val(style.align || 'left');
            $('#propTextContent').val(obj.text_content || '').prop('disabled', false);
        } else if (obj.obj_type === 'shape') {
            $('#imageProperties').hide();
            $('#imageAreaProperties').hide();
            $('#textProperties').hide();
            $('#shapeProperties').show();
            $('#tableProperties').hide();
            $('#chartProperties').hide();
            const s = obj.shape_style || {};
            $('#propShapeType').val(s.shape_type || 'rectangle');
            $('#propFillColor').val(s.fill_color || '#4A90D9');
            const opVal = Math.round((s.fill_opacity !== undefined ? s.fill_opacity : 1) * 100);
            $('#propFillOpacity').val(opVal);
            $('#opacityVal').text(opVal);
            $('#propStrokeColor').val(s.stroke_color || '#333333');
            $('#propStrokeWidth').val(s.stroke_width !== undefined ? s.stroke_width : 2);
            $('#propStrokeDash').val(s.stroke_dash || 'solid');
            $('#propBorderRadius').val(s.border_radius || 0);
            $('#propArrowHead').val(s.arrow_head || 'end');
            updateShapePropertyVisibility();
        } else if (obj.obj_type === 'table') {
            $('#imageProperties').hide();
            $('#imageAreaProperties').hide();
            $('#textProperties').hide();
            $('#shapeProperties').hide();
            $('#tableProperties').show();
            $('#chartProperties').hide();
            const ts = obj.table_style || {};
            $('#propTableRows').val(ts.rows || 3);
            $('#propTableCols').val(ts.cols || 3);
            $('#propTableHeaderRow').prop('checked', ts.header_row !== false);
            $('#propTableBandedRows').prop('checked', ts.banded_rows || false);
            $('#propTableHeaderBg').val(ts.header_bg_color || '#4472C4');
            $('#propTableHeaderText').val(ts.header_text_color || '#FFFFFF');
            $('#propTableBorderColor').val(ts.border_color || '#BFBFBF');
            $('#propTableFontSize').val(ts.font_size || 11);
        } else if (obj.obj_type === 'chart') {
            $('#imageProperties').hide();
            $('#imageAreaProperties').hide();
            $('#textProperties').hide();
            $('#shapeProperties').hide();
            $('#tableProperties').hide();
            $('#chartProperties').show();
            const cs = obj.chart_style || {};
            $('#propChartType').val(cs.chart_type || 'bar');
            $('#propChartTitle').val(cs.title || '');
            $('#propChartLegend').prop('checked', cs.show_legend !== false);
            $('#propChartGrid').prop('checked', cs.show_grid !== false);
        } else {
            $('#imageProperties').hide();
            $('#imageAreaProperties').hide();
            $('#textProperties').hide();
            $('#shapeProperties').hide();
            $('#tableProperties').hide();
            $('#chartProperties').hide();
        }
        return;
    }

    // 멀티셀렉트 모드
    $('#objProperties').show();

    // 멀티선택 안내 배지
    if ($('.multi-select-info').length === 0) {
        $('#objProperties').prepend(
            '<div class="multi-select-info">' +
            '<span style="font-size:14px;">&#9776;</span> ' +
            '<span class="multi-select-count">' + count + '개 객체 선택됨</span>' +
            '</div>'
        );
    } else {
        $('.multi-select-count').text(count + '개 객체 선택됨');
    }

    // 그룹 바운딩박스 좌표
    let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
    state.selectedObjects.forEach(obj => {
        minX = Math.min(minX, obj.x);
        minY = Math.min(minY, obj.y);
        maxX = Math.max(maxX, obj.x + obj.width);
        maxY = Math.max(maxY, obj.y + obj.height);
    });
    $('#propX').val(Math.round(minX));
    $('#propY').val(Math.round(minY));
    $('#propW').val(Math.round(maxX - minX)).prop('disabled', false);
    $('#propH').val(Math.round(maxY - minY)).prop('disabled', false);

    // z_index: 공통값 또는 혼합
    const zValues = [...new Set(state.selectedObjects.map(o => o.z_index !== undefined ? o.z_index : 10))];
    $('#propZIndex').val(zValues.length === 1 ? zValues[0] : '');

    // Role: 공통값 또는 혼합
    const roles = [...new Set(state.selectedObjects.map(o => o.role || ''))];
    $('#propRole').val(roles.length === 1 ? roles[0] : '');
    $('#propPlaceholder').val('').prop('disabled', true);

    // 타입별 패널
    const types = [...new Set(state.selectedObjects.map(o => o.obj_type))];

    if (types.length === 1 && types[0] === 'text') {
        $('#imageProperties').hide();
        $('#imageAreaProperties').hide();
        $('#textProperties').show();
        $('#shapeProperties').hide();
        $('#tableProperties').hide();
        $('#chartProperties').hide();
        updateMultiTextPanel();
    } else if (types.length === 1 && types[0] === 'shape') {
        $('#imageProperties').hide();
        $('#imageAreaProperties').hide();
        $('#textProperties').hide();
        $('#shapeProperties').show();
        $('#tableProperties').hide();
        $('#chartProperties').hide();
        updateMultiShapePanel();
    } else if (types.length === 1 && types[0] === 'image') {
        $('#imageProperties').show();
        $('#imageAreaProperties').hide();
        $('#textProperties').hide();
        $('#shapeProperties').hide();
        $('#tableProperties').hide();
        $('#chartProperties').hide();
        const fits = [...new Set(state.selectedObjects.map(o => o.image_fit || 'contain'))];
        $('#propImageFit').val(fits.length === 1 ? fits[0] : 'contain');
    } else if (types.length === 1 && types[0] === 'image_area') {
        $('#imageProperties').hide();
        $('#imageAreaProperties').show();
        $('#textProperties').hide();
        $('#shapeProperties').hide();
        $('#tableProperties').hide();
        $('#chartProperties').hide();
        const fits = [...new Set(state.selectedObjects.map(o => o.image_fit || 'cover'))];
        $('#propImageAreaFit').val(fits.length === 1 ? fits[0] : 'cover');
    } else {
        // 혼합 타입: 타입별 패널 숨김
        $('#imageProperties').hide();
        $('#imageAreaProperties').hide();
        $('#textProperties').hide();
        $('#shapeProperties').hide();
        $('#tableProperties').hide();
        $('#chartProperties').hide();
    }
}

function getCommonValue(objects, getter) {
    const values = objects.map(getter);
    const unique = [...new Set(values.map(v => JSON.stringify(v)))];
    if (unique.length === 1) return JSON.parse(unique[0]);
    return '__mixed__';
}

function updateMultiTextPanel() {
    const textObjs = state.selectedObjects.filter(o => o.obj_type === 'text');
    const fontFamily = getCommonValue(textObjs, o => (o.text_style || {}).font_family || 'Arial');
    const fontSize = getCommonValue(textObjs, o => (o.text_style || {}).font_size || 16);
    const color = getCommonValue(textObjs, o => (o.text_style || {}).color || '#000000');
    const bold = getCommonValue(textObjs, o => !!(o.text_style || {}).bold);
    const italic = getCommonValue(textObjs, o => !!(o.text_style || {}).italic);
    const align = getCommonValue(textObjs, o => (o.text_style || {}).align || 'left');

    $('#propFont').val(fontFamily === '__mixed__' ? '' : fontFamily);
    $('#propFontSize').val(fontSize === '__mixed__' ? '' : fontSize);
    $('#propFontSize').attr('placeholder', fontSize === '__mixed__' ? '혼합' : '');
    $('#propColor').val(color === '__mixed__' ? '#000000' : color);
    $('#propBold').toggleClass('active', bold === true).toggleClass('mixed', bold === '__mixed__');
    $('#propItalic').toggleClass('active', italic === true).toggleClass('mixed', italic === '__mixed__');
    $('#propAlign').val(align === '__mixed__' ? '' : align);
    $('#propTextContent').val('').attr('placeholder', textObjs.length + '개 텍스트 선택됨').prop('disabled', true);
}

function updateMultiShapePanel() {
    const shapeObjs = state.selectedObjects.filter(o => o.obj_type === 'shape');
    const fillColor = getCommonValue(shapeObjs, o => (o.shape_style || {}).fill_color || '#4A90D9');
    const fillOpacity = getCommonValue(shapeObjs, o => {
        const op = (o.shape_style || {}).fill_opacity;
        return op !== undefined ? op : 1;
    });
    const strokeColor = getCommonValue(shapeObjs, o => (o.shape_style || {}).stroke_color || '#333333');
    const strokeWidth = getCommonValue(shapeObjs, o => {
        const sw = (o.shape_style || {}).stroke_width;
        return sw !== undefined ? sw : 2;
    });
    const strokeDash = getCommonValue(shapeObjs, o => (o.shape_style || {}).stroke_dash || 'solid');
    const shapeType = getCommonValue(shapeObjs, o => (o.shape_style || {}).shape_type || 'rectangle');

    $('#propShapeType').val(shapeType === '__mixed__' ? '' : shapeType);
    $('#propFillColor').val(fillColor === '__mixed__' ? '#4A90D9' : fillColor);
    const opVal = fillOpacity === '__mixed__' ? 100 : Math.round(fillOpacity * 100);
    $('#propFillOpacity').val(opVal);
    $('#opacityVal').text(opVal);
    $('#propStrokeColor').val(strokeColor === '__mixed__' ? '#333333' : strokeColor);
    $('#propStrokeWidth').val(strokeWidth === '__mixed__' ? '' : strokeWidth);
    $('#propStrokeDash').val(strokeDash === '__mixed__' ? '' : strokeDash);
    updateShapePropertyVisibility();
}

function updatePositionInputs() {
    if (state.selectedObjects.length > 1) {
        // 멀티셀렉트: 그룹 바운딩박스 좌표
        let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
        state.selectedObjects.forEach(obj => {
            minX = Math.min(minX, obj.x);
            minY = Math.min(minY, obj.y);
            maxX = Math.max(maxX, obj.x + obj.width);
            maxY = Math.max(maxY, obj.y + obj.height);
        });
        $('#propX').val(Math.round(minX));
        $('#propY').val(Math.round(minY));
        $('#propW').val(Math.round(maxX - minX));
        $('#propH').val(Math.round(maxY - minY));
        return;
    }
    if (!state.selectedObject) return;
    $('#propX').val(Math.round(state.selectedObject.x));
    $('#propY').val(Math.round(state.selectedObject.y));
    $('#propW').val(Math.round(state.selectedObject.width));
    $('#propH').val(Math.round(state.selectedObject.height));
}

function updateObjProperty(prop, value) {
    if (state.selectedObjects.length === 0) return;
    state.selectedObjects.forEach(obj => {
        obj[prop] = value;
    });
    if (prop === 'role') {
        updateSlideMetaUI();
        // 날짜 역할 선택 시 오늘 날짜 자동 입력
        if (value === 'date') {
            const today = new Date();
            const y = today.getFullYear();
            const m = String(today.getMonth() + 1).padStart(2, '0');
            const d = String(today.getDate()).padStart(2, '0');
            const dateStr = `${y}.${m}.${d}`;
            state.selectedObjects.forEach(o => {
                if (o.obj_type === 'text') {
                    o.text_content = dateStr;
                    $(`.canvas-object[data-obj-id="${o.obj_id}"] .text-content`).text(dateStr);
                }
            });
            if (state.selectedObjects.length === 1) {
                $('#propTextContent').val(dateStr);
            }
        }
    }
}

function updateObjPosition() {
    if (state.selectedObjects.length > 1) {
        // 멀티셀렉트: 현재 바운딩박스 계산
        let oldMinX = Infinity, oldMinY = Infinity, oldMaxX = -Infinity, oldMaxY = -Infinity;
        state.selectedObjects.forEach(obj => {
            oldMinX = Math.min(oldMinX, obj.x);
            oldMinY = Math.min(oldMinY, obj.y);
            oldMaxX = Math.max(oldMaxX, obj.x + obj.width);
            oldMaxY = Math.max(oldMaxY, obj.y + obj.height);
        });
        const oldBBoxW = oldMaxX - oldMinX;
        const oldBBoxH = oldMaxY - oldMinY;

        const newX = parseInt($('#propX').val()) || 0;
        const newY = parseInt($('#propY').val()) || 0;
        const newW = parseInt($('#propW').val()) || oldBBoxW;
        const newH = parseInt($('#propH').val()) || oldBBoxH;

        // X/Y 이동 델타
        const deltaX = newX - oldMinX;
        const deltaY = newY - oldMinY;

        // W/H 비례 스케일 팩터
        const scaleX = (oldBBoxW > 0) ? newW / oldBBoxW : 1;
        const scaleY = (oldBBoxH > 0) ? newH / oldBBoxH : 1;

        state.selectedObjects.forEach(obj => {
            // 바운딩박스 원점(oldMinX, oldMinY) 기준으로 상대 위치 스케일링
            const relX = obj.x - oldMinX;
            const relY = obj.y - oldMinY;

            obj.x = Math.round(newX + relX * scaleX);
            obj.y = Math.round(newY + relY * scaleY);
            obj.width = Math.max(1, Math.round(obj.width * scaleX));
            obj.height = Math.max(1, Math.round(obj.height * scaleY));

            // 캔버스 범위 제한
            const szClamp = getTemplateCanvasSize();
            obj.x = Math.max(0, Math.min(szClamp.w - obj.width, obj.x));
            obj.y = Math.max(0, Math.min(szClamp.h - obj.height, obj.y));

            const el = $(`[data-obj-id="${obj.obj_id}"]`);
            el.css({
                left: obj.x + 'px',
                top: obj.y + 'px',
                width: obj.width + 'px',
                height: obj.height + 'px',
            });
            // 도형 SVG 재렌더
            if (obj.obj_type === 'shape') {
                el.find('svg').remove();
                el.prepend(createShapeSVG(obj));
            }
        });
        updateGroupBoundingBox();
        renderSlideThumbnail(state.slides.find(s => s._id === state.currentSlide));
        return;
    }

    if (!state.selectedObject) return;
    state.selectedObject.x = parseInt($('#propX').val()) || 0;
    state.selectedObject.y = parseInt($('#propY').val()) || 0;
    state.selectedObject.width = parseInt($('#propW').val()) || 100;
    state.selectedObject.height = parseInt($('#propH').val()) || 50;

    const el = $(`[data-obj-id="${state.selectedObject.obj_id}"]`);
    el.css({
        left: state.selectedObject.x + 'px',
        top: state.selectedObject.y + 'px',
        width: state.selectedObject.width + 'px',
        height: state.selectedObject.height + 'px',
    });
    // 도형 SVG 재렌더
    if (state.selectedObject.obj_type === 'shape') {
        el.find('svg').remove();
        el.prepend(createShapeSVG(state.selectedObject));
    }
}

function updateImageFit(value) {
    const imgObjs = state.selectedObjects.filter(o => o.obj_type === 'image');
    if (imgObjs.length === 0) return;
    imgObjs.forEach(obj => {
        obj.image_fit = value;
        $(`[data-obj-id="${obj.obj_id}"] img`).css('object-fit', value);
    });
    renderSlideThumbnail(state.slides.find(s => s._id === state.currentSlide));
}

function updateImageAreaFit(value) {
    if (!state.selectedObject || state.selectedObject.obj_type !== 'image_area') return;
    state.selectedObject.image_fit = value;
    // No visual change needed for placeholder
}

// ============ 최근 사용 색상 ============

function addRecentColor(color) {
    if (!color) return;
    const c = color.toUpperCase();
    // 중복 제거
    state.recentColors = state.recentColors.filter(rc => rc.toUpperCase() !== c);
    // 맨 앞에 추가
    state.recentColors.unshift(color);
    // 최대 5개 유지
    if (state.recentColors.length > 5) {
        state.recentColors = state.recentColors.slice(0, 5);
    }
    renderRecentColors();
}

function renderRecentColors() {
    $('.recent-colors').each(function () {
        const $container = $(this);
        const targetId = $container.data('target');
        $container.empty();
        if (state.recentColors.length === 0) return;
        state.recentColors.forEach(function (color) {
            const $swatch = $('<span>')
                .addClass('recent-color-swatch')
                .css('background-color', color)
                .attr('title', color)
                .on('click', function () {
                    // 색상 적용
                    $('#' + targetId).val(color);
                    // 해당 input의 onchange 트리거
                    if (targetId === 'propColor') {
                        updateTextStyle('color', color);
                    } else if (targetId === 'propFillColor') {
                        updateShapeStyle('fill_color', color);
                    } else if (targetId === 'propStrokeColor') {
                        updateShapeStyle('stroke_color', color);
                    }
                    addRecentColor(color);
                });
            $container.append($swatch);
        });
    });
}

function updateTextStyle(prop, value) {
    // 멀티셀렉트: 선택된 모든 텍스트 객체에 적용
    const textObjs = state.selectedObjects.filter(o => o.obj_type === 'text');
    if (textObjs.length === 0) return;

    textObjs.forEach(obj => {
        if (!obj.text_style) obj.text_style = {};
        obj.text_style[prop] = value;

        const textEl = $(`[data-obj-id="${obj.obj_id}"] .text-content`);
        switch (prop) {
            case 'font_family': textEl.css('fontFamily', value); break;
            case 'font_size': textEl.css('fontSize', value + 'px'); break;
            case 'color': textEl.css('color', value); break;
            case 'bold': textEl.css('fontWeight', value ? 'bold' : 'normal'); break;
            case 'italic': textEl.css('fontStyle', value ? 'italic' : 'normal'); break;
            case 'align': textEl.css('textAlign', value); break;
        }
    });

    // 색상 변경 시 최근 색상에 추가
    if (prop === 'color') {
        addRecentColor(value);
    }
}

function toggleBold() {
    const btn = $('#propBold');
    const newVal = !btn.hasClass('active');
    btn.toggleClass('active', newVal).removeClass('mixed');
    updateTextStyle('bold', newVal);
}

function toggleItalic() {
    const btn = $('#propItalic');
    const newVal = !btn.hasClass('active');
    btn.toggleClass('active', newVal).removeClass('mixed');
    updateTextStyle('italic', newVal);
}

function updateTextContent(value) {
    if (!state.selectedObject || state.selectedObject.obj_type !== 'text') return;
    state.selectedObject.text_content = value;
    $(`[data-obj-id="${state.selectedObject.obj_id}"] .text-content`).text(value);
}

// ============ 레이어(z-index) 관리 ============

function bringForward() {
    if (state.selectedObjects.length === 0) return;
    // 현재 선택된 객체들보다 z_index가 바로 위인 객체 찾기
    const selectedIds = new Set(state.selectedObjects.map(o => o.obj_id));
    const maxSelectedZ = Math.max(...state.selectedObjects.map(o => o.z_index !== undefined ? o.z_index : 10));

    // 선택되지 않은 객체 중 바로 위 z_index 가진 것들
    const above = state.objects.filter(o => !selectedIds.has(o.obj_id) && (o.z_index !== undefined ? o.z_index : 10) > maxSelectedZ);
    if (above.length === 0) return; // 이미 맨 앞

    const targetZ = Math.min(...above.map(o => o.z_index !== undefined ? o.z_index : 10));

    // 선택된 객체: targetZ 값으로, target 객체: 선택 객체의 원래 z로 교환
    state.selectedObjects.forEach(obj => {
        obj.z_index = targetZ + 1;
        $(`[data-obj-id="${obj.obj_id}"]`).css('zIndex', obj.z_index);
    });
    updatePropertiesPanel();
}

function sendBackward() {
    if (state.selectedObjects.length === 0) return;
    const selectedIds = new Set(state.selectedObjects.map(o => o.obj_id));
    const minSelectedZ = Math.min(...state.selectedObjects.map(o => o.z_index !== undefined ? o.z_index : 10));

    const below = state.objects.filter(o => !selectedIds.has(o.obj_id) && (o.z_index !== undefined ? o.z_index : 10) < minSelectedZ);
    if (below.length === 0) return; // 이미 맨 뒤

    const targetZ = Math.max(...below.map(o => o.z_index !== undefined ? o.z_index : 10));

    state.selectedObjects.forEach(obj => {
        obj.z_index = targetZ - 1;
        if (obj.z_index < 1) obj.z_index = 1;
        $(`[data-obj-id="${obj.obj_id}"]`).css('zIndex', obj.z_index);
    });
    updatePropertiesPanel();
}

function bringToFront() {
    if (state.selectedObjects.length === 0) return;
    const maxZ = Math.max(...state.objects.map(o => o.z_index !== undefined ? o.z_index : 10));
    state.selectedObjects.forEach((obj, i) => {
        obj.z_index = maxZ + 1 + i;
        $(`[data-obj-id="${obj.obj_id}"]`).css('zIndex', obj.z_index);
    });
    updatePropertiesPanel();
}

function sendToBack() {
    if (state.selectedObjects.length === 0) return;
    const minZ = Math.min(...state.objects.map(o => o.z_index !== undefined ? o.z_index : 10));
    const newBase = Math.max(1, minZ - state.selectedObjects.length);
    state.selectedObjects.forEach((obj, i) => {
        obj.z_index = newBase + i;
        $(`[data-obj-id="${obj.obj_id}"]`).css('zIndex', obj.z_index);
    });
    updatePropertiesPanel();
}

function updateZIndex(value) {
    const z = parseInt(value);
    if (isNaN(z) || z < 1) return;
    state.selectedObjects.forEach(obj => {
        obj.z_index = z;
        $(`[data-obj-id="${obj.obj_id}"]`).css('zIndex', z);
    });
}

// ============ 도형 기능 ============

function cancelDrawMode() {
    state.isDrawing = false;
    state.drawShapeType = null;
    state.drawChartType = null;
    state.drawStart = null;
    $('#drawGhost').remove();
    $('#canvas').removeClass('draw-mode');
    $('.shape-dropdown-wrapper .btn').removeClass('draw-mode-active');
}

function createShapeAtRect(shapeType, x, y, width, height) {
    const isLine = (shapeType === 'line' || shapeType === 'arrow');
    const obj = {
        obj_id: generateObjId(),
        obj_type: 'shape',
        x: x,
        y: y,
        width: width,
        height: isLine ? Math.max(height, 4) : height,
        z_index: getNextZIndex(),
        shape_style: {
            shape_type: shapeType,
            fill_color: isLine ? 'transparent' : '#4A90D9',
            fill_opacity: isLine ? 0 : 1.0,
            stroke_color: '#333333',
            stroke_width: 2,
            stroke_dash: 'solid',
            border_radius: shapeType === 'rounded_rectangle' ? 12 : 0,
            arrow_head: shapeType === 'arrow' ? 'end' : 'none',
        },
        role: null,
        placeholder: null,
    };

    state.objects.push(obj);
    const el = createObjectElement(obj);
    $('#canvas').append(el);
    selectObject(obj.obj_id);
}

function toggleShapeDropdown() {
    const dd = $('#shapeDropdown');
    if (dd.is(':visible')) {
        dd.hide();
        return;
    }
    dd.show();
    setTimeout(function() {
        $(document).on('mousedown.shapepicker', function (e) {
            if (!$(e.target).closest('.shape-dropdown-wrapper').length) {
                dd.hide();
                $(document).off('mousedown.shapepicker');
            }
        });
    }, 0);
}

function addShapeObject(shapeType) {
    if (!state.currentSlide) {
        showToast('먼저 슬라이드를 선택하세요', 'error');
        return;
    }
    $('#shapeDropdown').hide();

    // 드로잉 모드 진입
    state.isDrawing = true;
    state.drawShapeType = shapeType;
    state.drawStart = null;

    $('#canvas').addClass('draw-mode');
    $('.shape-dropdown-wrapper .btn').addClass('draw-mode-active');
    showToast('캔버스에서 드래그하여 도형을 그리세요 (ESC: 취소)', 'info');
}

// ============ 도형 헬퍼 함수 ============
function _regPolygonPts(cx, cy, r, n, rotDeg) {
    var pts = [], rad = (rotDeg - 90) * Math.PI / 180;
    for (var i = 0; i < n; i++) {
        var a = rad + 2 * Math.PI * i / n;
        pts.push((cx + r * Math.cos(a)).toFixed(1) + ',' + (cy + r * Math.sin(a)).toFixed(1));
    }
    return pts.join(' ');
}

function _starPts(cx, cy, outerR, innerR, n, rotDeg) {
    var pts = [], rad = (rotDeg - 90) * Math.PI / 180;
    for (var i = 0; i < n * 2; i++) {
        var r = i % 2 === 0 ? outerR : innerR;
        var a = rad + Math.PI * i / n;
        pts.push((cx + r * Math.cos(a)).toFixed(1) + ',' + (cy + r * Math.sin(a)).toFixed(1));
    }
    return pts.join(' ');
}

// ============ 도형 아이콘 SVG (24x24 picker용) ============
function _getShapeIconSVG(type) {
    var s = 'fill="none" stroke="#444" stroke-width="1.2"';
    var sf = 'fill="#444" stroke="none"';
    switch (type) {
        // 직사각형
        case 'rectangle': return '<svg viewBox="0 0 24 24"><rect x="2" y="5" width="20" height="14" '+s+'/></svg>';
        case 'rounded_rectangle': return '<svg viewBox="0 0 24 24"><rect x="2" y="5" width="20" height="14" rx="3" '+s+'/></svg>';
        case 'snip_1_rect': return '<svg viewBox="0 0 24 24"><path d="M2 5h16l4 4v10H2z" '+s+'/></svg>';
        case 'snip_2_diag_rect': return '<svg viewBox="0 0 24 24"><path d="M6 5h16l-4 4v10H2V9z" '+s+'/></svg>';
        case 'round_1_rect': return '<svg viewBox="0 0 24 24"><path d="M2 5h16a4 4 0 0 1 4 4v10H2z" '+s+'/></svg>';
        case 'round_2_diag_rect': return '<svg viewBox="0 0 24 24"><path d="M6 5h16v10a4 4 0 0 1-4 4H2V9a4 4 0 0 1 4-4z" '+s+'/></svg>';
        // 기본 도형
        case 'ellipse': return '<svg viewBox="0 0 24 24"><ellipse cx="12" cy="12" rx="10" ry="7" '+s+'/></svg>';
        case 'triangle': return '<svg viewBox="0 0 24 24"><polygon points="12,2 22,20 2,20" '+s+'/></svg>';
        case 'right_triangle': return '<svg viewBox="0 0 24 24"><polygon points="2,20 22,20 2,4" '+s+'/></svg>';
        case 'parallelogram': return '<svg viewBox="0 0 24 24"><polygon points="6,5 22,5 18,19 2,19" '+s+'/></svg>';
        case 'trapezoid': return '<svg viewBox="0 0 24 24"><polygon points="6,5 18,5 22,19 2,19" '+s+'/></svg>';
        case 'diamond': return '<svg viewBox="0 0 24 24"><polygon points="12,2 22,12 12,22 2,12" '+s+'/></svg>';
        case 'pentagon': return '<svg viewBox="0 0 24 24"><polygon points="'+_regPolygonPts(12,12,10,5,0)+'" '+s+'/></svg>';
        case 'hexagon': return '<svg viewBox="0 0 24 24"><polygon points="'+_regPolygonPts(12,12,10,6,0)+'" '+s+'/></svg>';
        case 'heptagon': return '<svg viewBox="0 0 24 24"><polygon points="'+_regPolygonPts(12,12,10,7,0)+'" '+s+'/></svg>';
        case 'octagon': return '<svg viewBox="0 0 24 24"><polygon points="'+_regPolygonPts(12,12,10,8,0)+'" '+s+'/></svg>';
        case 'decagon': return '<svg viewBox="0 0 24 24"><polygon points="'+_regPolygonPts(12,12,10,10,0)+'" '+s+'/></svg>';
        case 'dodecagon': return '<svg viewBox="0 0 24 24"><polygon points="'+_regPolygonPts(12,12,10,12,0)+'" '+s+'/></svg>';
        case 'cross': return '<svg viewBox="0 0 24 24"><polygon points="8,2 16,2 16,8 22,8 22,16 16,16 16,22 8,22 8,16 2,16 2,8 8,8" '+s+'/></svg>';
        case 'donut': return '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" '+s+'/><circle cx="12" cy="12" r="5" '+s+'/></svg>';
        case 'no_smoking': return '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" '+s+'/><line x1="5" y1="5" x2="19" y2="19" stroke="#444" stroke-width="1.2"/></svg>';
        case 'block_arc': return '<svg viewBox="0 0 24 24"><path d="M4 20A10 10 0 0 1 20 20" '+s+'/></svg>';
        case 'heart': return '<svg viewBox="0 0 24 24"><path d="M12 21C12 21 3 14 3 8.5C3 5.4 5.4 3 8.5 3C10.2 3 11.8 3.8 12 5C12.2 3.8 13.8 3 15.5 3C18.6 3 21 5.4 21 8.5C21 14 12 21 12 21Z" '+s+'/></svg>';
        case 'lightning_bolt': return '<svg viewBox="0 0 24 24"><polygon points="13,2 6,13 11,13 10,22 18,10 13,10" '+s+'/></svg>';
        case 'sun': return '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="5" '+s+'/><g stroke="#444" stroke-width="1.2"><line x1="12" y1="1" x2="12" y2="4"/><line x1="12" y1="20" x2="12" y2="23"/><line x1="1" y1="12" x2="4" y2="12"/><line x1="20" y1="12" x2="23" y2="12"/><line x1="4.2" y1="4.2" x2="6.3" y2="6.3"/><line x1="17.7" y1="17.7" x2="19.8" y2="19.8"/><line x1="4.2" y1="19.8" x2="6.3" y2="17.7"/><line x1="17.7" y1="6.3" x2="19.8" y2="4.2"/></g></svg>';
        case 'moon': return '<svg viewBox="0 0 24 24"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z" '+s+'/></svg>';
        case 'cloud': return '<svg viewBox="0 0 24 24"><path d="M6 19a4 4 0 0 1-.8-7.9A5.5 5.5 0 0 1 16 6.5h.5A4 4 0 0 1 20 13a3 3 0 0 1-2 5H6z" '+s+'/></svg>';
        case 'smiley_face': return '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" '+s+'/><path d="M8 14s1.5 2 4 2 4-2 4-2" '+s+'/><circle cx="9" cy="9" r="1" '+sf+'/><circle cx="15" cy="9" r="1" '+sf+'/></svg>';
        case 'folded_corner': return '<svg viewBox="0 0 24 24"><path d="M2 2h20v14l-6 6H2z" '+s+'/><path d="M16 22v-6h6" '+s+'/></svg>';
        case 'frame': return '<svg viewBox="0 0 24 24"><rect x="2" y="2" width="20" height="20" '+s+'/><rect x="5" y="5" width="14" height="14" '+s+'/></svg>';
        case 'teardrop': return '<svg viewBox="0 0 24 24"><path d="M12 2C12 2 4 10 4 15a8 8 0 0 0 16 0c0-5-8-13-8-13z" '+s+'/></svg>';
        case 'plaque': return '<svg viewBox="0 0 24 24"><path d="M4 2h16a0 0 0 0 1 0 0v0a4 4 0 0 0-4 4v12a4 4 0 0 0 4 4v0H4v0a4 4 0 0 0 4-4V6a4 4 0 0 0-4-4z" '+s+'/></svg>';
        case 'brace_pair': return '<svg viewBox="0 0 24 24"><path d="M8 2C5 2 4 4 4 6v4c0 1-1 2-2 2 1 0 2 1 2 2v4c0 2 1 4 4 4" '+s+'/><path d="M16 2c3 0 4 2 4 4v4c0 1 1 2 2 2-1 0-2 1-2 2v4c0 2-1 4-4 4" '+s+'/></svg>';
        case 'bracket_pair': return '<svg viewBox="0 0 24 24"><path d="M8 2H4v20h4" '+s+'/><path d="M16 2h4v20h-4" '+s+'/></svg>';
        // 블록 화살표
        case 'right_arrow': return '<svg viewBox="0 0 24 24"><polygon points="2,8 15,8 15,3 22,12 15,21 15,16 2,16" '+s+'/></svg>';
        case 'left_arrow': return '<svg viewBox="0 0 24 24"><polygon points="22,8 9,8 9,3 2,12 9,21 9,16 22,16" '+s+'/></svg>';
        case 'up_arrow': return '<svg viewBox="0 0 24 24"><polygon points="8,22 8,9 3,9 12,2 21,9 16,9 16,22" '+s+'/></svg>';
        case 'down_arrow': return '<svg viewBox="0 0 24 24"><polygon points="8,2 8,15 3,15 12,22 21,15 16,15 16,2" '+s+'/></svg>';
        case 'left_right_arrow': return '<svg viewBox="0 0 24 24"><polygon points="6,3 6,8 18,8 18,3 23,12 18,21 18,16 6,16 6,21 1,12" '+s+'/></svg>';
        case 'up_down_arrow': return '<svg viewBox="0 0 24 24"><polygon points="3,6 8,6 8,18 3,18 12,23 21,18 16,18 16,6 21,6 12,1" '+s+'/></svg>';
        case 'quad_arrow': return '<svg viewBox="0 0 24 24"><polygon points="12,1 15,5 13,5 13,9 17,9 17,5 21,5 23,12 21,19 17,19 17,15 13,15 13,19 15,19 12,23 9,19 11,19 11,15 7,15 7,19 3,19 1,12 3,5 7,5 7,9 11,9 11,5 9,5" '+s+'/></svg>';
        case 'notched_right_arrow': return '<svg viewBox="0 0 24 24"><polygon points="2,8 15,8 15,3 22,12 15,21 15,16 2,16 5,12" '+s+'/></svg>';
        case 'chevron': return '<svg viewBox="0 0 24 24"><polygon points="2,4 16,4 22,12 16,20 2,20 8,12" '+s+'/></svg>';
        case 'home_plate': return '<svg viewBox="0 0 24 24"><polygon points="2,4 18,4 22,12 18,20 2,20" '+s+'/></svg>';
        case 'striped_right_arrow': return '<svg viewBox="0 0 24 24"><polygon points="8,8 15,8 15,3 22,12 15,21 15,16 8,16" '+s+'/><line x1="5" y1="8" x2="5" y2="16" stroke="#444" stroke-width="1.2"/><line x1="3" y1="8" x2="3" y2="16" stroke="#444" stroke-width="1.2"/></svg>';
        case 'bent_arrow': return '<svg viewBox="0 0 24 24"><path d="M4 20V10h8V5l6 7-6 7v-5H8v6z" '+s+'/></svg>';
        case 'u_turn_arrow': return '<svg viewBox="0 0 24 24"><path d="M6 20V10a6 6 0 0 1 12 0v4h4l-6 6-6-6h4v-4a2 2 0 0 0-4 0v10z" '+s+'/></svg>';
        case 'circular_arrow': return '<svg viewBox="0 0 24 24"><path d="M20 12a8 8 0 1 1-3-6.3" '+s+'/><polygon points="20,2 22,8 16,8" '+sf+'/></svg>';
        // 수학
        case 'math_plus': return '<svg viewBox="0 0 24 24"><line x1="12" y1="4" x2="12" y2="20" stroke="#444" stroke-width="2"/><line x1="4" y1="12" x2="20" y2="12" stroke="#444" stroke-width="2"/></svg>';
        case 'math_minus': return '<svg viewBox="0 0 24 24"><line x1="4" y1="12" x2="20" y2="12" stroke="#444" stroke-width="2"/></svg>';
        case 'math_multiply': return '<svg viewBox="0 0 24 24"><line x1="6" y1="6" x2="18" y2="18" stroke="#444" stroke-width="2"/><line x1="18" y1="6" x2="6" y2="18" stroke="#444" stroke-width="2"/></svg>';
        case 'math_divide': return '<svg viewBox="0 0 24 24"><line x1="4" y1="12" x2="20" y2="12" stroke="#444" stroke-width="2"/><circle cx="12" cy="7" r="1.5" '+sf+'/><circle cx="12" cy="17" r="1.5" '+sf+'/></svg>';
        case 'math_equal': return '<svg viewBox="0 0 24 24"><line x1="4" y1="9" x2="20" y2="9" stroke="#444" stroke-width="2"/><line x1="4" y1="15" x2="20" y2="15" stroke="#444" stroke-width="2"/></svg>';
        case 'math_not_equal': return '<svg viewBox="0 0 24 24"><line x1="4" y1="9" x2="20" y2="9" stroke="#444" stroke-width="2"/><line x1="4" y1="15" x2="20" y2="15" stroke="#444" stroke-width="2"/><line x1="16" y1="4" x2="8" y2="20" stroke="#444" stroke-width="2"/></svg>';
        // 별 및 현수막
        case 'star_4_point': return '<svg viewBox="0 0 24 24"><polygon points="'+_starPts(12,12,10,4,4,0)+'" '+s+'/></svg>';
        case 'star_5_point': return '<svg viewBox="0 0 24 24"><polygon points="'+_starPts(12,12,10,4,5,0)+'" '+s+'/></svg>';
        case 'star_6_point': return '<svg viewBox="0 0 24 24"><polygon points="'+_starPts(12,12,10,5,6,0)+'" '+s+'/></svg>';
        case 'star_8_point': return '<svg viewBox="0 0 24 24"><polygon points="'+_starPts(12,12,10,5,8,0)+'" '+s+'/></svg>';
        case 'star_10_point': return '<svg viewBox="0 0 24 24"><polygon points="'+_starPts(12,12,10,6,10,0)+'" '+s+'/></svg>';
        case 'star_12_point': return '<svg viewBox="0 0 24 24"><polygon points="'+_starPts(12,12,10,6,12,0)+'" '+s+'/></svg>';
        case 'star_16_point': return '<svg viewBox="0 0 24 24"><polygon points="'+_starPts(12,12,10,7,16,0)+'" '+s+'/></svg>';
        case 'star_24_point': return '<svg viewBox="0 0 24 24"><polygon points="'+_starPts(12,12,10,7,24,0)+'" '+s+'/></svg>';
        case 'star_32_point': return '<svg viewBox="0 0 24 24"><polygon points="'+_starPts(12,12,10,8,32,0)+'" '+s+'/></svg>';
        case 'explosion_1': return '<svg viewBox="0 0 24 24"><polygon points="12,2 14,8 20,4 16,10 22,12 16,14 20,20 14,16 12,22 10,16 4,20 8,14 2,12 8,10 4,4 10,8" '+s+'/></svg>';
        case 'explosion_2': return '<svg viewBox="0 0 24 24"><polygon points="12,1 13,7 18,3 15,9 22,8 17,12 22,16 15,15 18,21 13,17 12,23 11,17 6,21 9,15 2,16 7,12 2,8 9,9 6,3 11,7" '+s+'/></svg>';
        case 'wave': return '<svg viewBox="0 0 24 24"><path d="M2 8c4-6 6 6 10 0s6 6 10 0v8c-4 6-6-6-10 0s-6-6-10 0z" '+s+'/></svg>';
        case 'double_wave': return '<svg viewBox="0 0 24 24"><path d="M2 6c3-4 5 4 10 0s7 4 10 0" '+s+'/><path d="M2 18c3-4 5 4 10 0s7 4 10 0" '+s+'/></svg>';
        case 'ribbon': return '<svg viewBox="0 0 24 24"><path d="M2 6h20v12H2z" '+s+'/><path d="M2 6l3 3-3 3" '+s+'/><path d="M22 6l-3 3 3 3" '+s+'/></svg>';
        // 설명선
        case 'wedge_rect_callout': return '<svg viewBox="0 0 24 24"><path d="M2 3h20v13H14l-2 5-2-5H2z" '+s+'/></svg>';
        case 'wedge_round_rect_callout': return '<svg viewBox="0 0 24 24"><path d="M5 3h14a3 3 0 0 1 3 3v7a3 3 0 0 1-3 3H14l-2 5-2-5H5a3 3 0 0 1-3-3V6a3 3 0 0 1 3-3z" '+s+'/></svg>';
        case 'wedge_ellipse_callout': return '<svg viewBox="0 0 24 24"><ellipse cx="12" cy="10" rx="10" ry="7" '+s+'/><path d="M10 16l0 5 4-4" '+s+'/></svg>';
        case 'cloud_callout': return '<svg viewBox="0 0 24 24"><path d="M6 16a4 4 0 0 1-.8-7.9A5.5 5.5 0 0 1 16 4h.5A4 4 0 0 1 20 10a3 3 0 0 1-2 5H6z" '+s+'/><circle cx="8" cy="19" r="1" '+sf+'/><circle cx="6" cy="21" r="0.7" '+sf+'/></svg>';
        case 'border_callout_1': return '<svg viewBox="0 0 24 24"><rect x="2" y="2" width="20" height="14" '+s+'/><line x1="8" y1="16" x2="8" y2="22" stroke="#444" stroke-width="1.2"/></svg>';
        case 'border_callout_2': return '<svg viewBox="0 0 24 24"><rect x="2" y="2" width="20" height="14" '+s+'/><polyline points="8,16 8,19 12,22" '+s+'/></svg>';
        case 'border_callout_3': return '<svg viewBox="0 0 24 24"><rect x="2" y="2" width="20" height="14" '+s+'/><polyline points="8,16 10,19 6,22" '+s+'/></svg>';
        // 선
        case 'line': return '<svg viewBox="0 0 24 24"><line x1="2" y1="18" x2="22" y2="6" stroke="#444" stroke-width="1.2"/></svg>';
        case 'arrow': return '<svg viewBox="0 0 24 24"><line x1="2" y1="18" x2="19" y2="6" stroke="#444" stroke-width="1.2"/><polygon points="22,4 16,6 19,9" fill="#444"/></svg>';
        default: return '<svg viewBox="0 0 24 24"><rect x="2" y="5" width="20" height="14" '+s+'/></svg>';
    }
}

// ============ 도형 카탈로그 ============
var SHAPE_CATEGORIES = [
    { name: '직사각형', shapes: [
        { id: 'rectangle', label: '사각형' },
        { id: 'rounded_rectangle', label: '둥근 사각형' },
        { id: 'snip_1_rect', label: '모서리 1개 잘림' },
        { id: 'snip_2_diag_rect', label: '대각 2개 잘림' },
        { id: 'round_1_rect', label: '모서리 1개 둥근' },
        { id: 'round_2_diag_rect', label: '대각 2개 둥근' },
    ]},
    { name: '기본 도형', shapes: [
        { id: 'ellipse', label: '원/타원' },
        { id: 'triangle', label: '삼각형' },
        { id: 'right_triangle', label: '직각 삼각형' },
        { id: 'parallelogram', label: '평행사변형' },
        { id: 'trapezoid', label: '사다리꼴' },
        { id: 'diamond', label: '마름모' },
        { id: 'pentagon', label: '오각형' },
        { id: 'hexagon', label: '육각형' },
        { id: 'heptagon', label: '칠각형' },
        { id: 'octagon', label: '팔각형' },
        { id: 'decagon', label: '십각형' },
        { id: 'dodecagon', label: '십이각형' },
        { id: 'cross', label: '십자형' },
        { id: 'donut', label: '도넛' },
        { id: 'no_smoking', label: '금지' },
        { id: 'block_arc', label: '호' },
        { id: 'heart', label: '하트' },
        { id: 'lightning_bolt', label: '번개' },
        { id: 'sun', label: '태양' },
        { id: 'moon', label: '달' },
        { id: 'cloud', label: '구름' },
        { id: 'smiley_face', label: '웃는 얼굴' },
        { id: 'folded_corner', label: '접힌 모서리' },
        { id: 'frame', label: '프레임' },
        { id: 'teardrop', label: '물방울' },
        { id: 'plaque', label: '명판' },
        { id: 'brace_pair', label: '중괄호 쌍' },
        { id: 'bracket_pair', label: '대괄호 쌍' },
    ]},
    { name: '블록 화살표', shapes: [
        { id: 'right_arrow', label: '오른쪽 화살표' },
        { id: 'left_arrow', label: '왼쪽 화살표' },
        { id: 'up_arrow', label: '위쪽 화살표' },
        { id: 'down_arrow', label: '아래쪽 화살표' },
        { id: 'left_right_arrow', label: '좌우 화살표' },
        { id: 'up_down_arrow', label: '상하 화살표' },
        { id: 'quad_arrow', label: '사방 화살표' },
        { id: 'notched_right_arrow', label: '노치 오른쪽 화살표' },
        { id: 'chevron', label: '쉐브론' },
        { id: 'home_plate', label: '홈 플레이트' },
        { id: 'striped_right_arrow', label: '줄무늬 오른쪽 화살표' },
        { id: 'bent_arrow', label: '꺾인 화살표' },
        { id: 'u_turn_arrow', label: '유턴 화살표' },
        { id: 'circular_arrow', label: '원형 화살표' },
    ]},
    { name: '수학', shapes: [
        { id: 'math_plus', label: '더하기' },
        { id: 'math_minus', label: '빼기' },
        { id: 'math_multiply', label: '곱하기' },
        { id: 'math_divide', label: '나누기' },
        { id: 'math_equal', label: '등호' },
        { id: 'math_not_equal', label: '부등호' },
    ]},
    { name: '별 및 현수막', shapes: [
        { id: 'star_4_point', label: '4각별' },
        { id: 'star_5_point', label: '5각별' },
        { id: 'star_6_point', label: '6각별' },
        { id: 'star_8_point', label: '8각별' },
        { id: 'star_10_point', label: '10각별' },
        { id: 'star_12_point', label: '12각별' },
        { id: 'star_16_point', label: '16각별' },
        { id: 'star_24_point', label: '24각별' },
        { id: 'star_32_point', label: '32각별' },
        { id: 'explosion_1', label: '폭발 1' },
        { id: 'explosion_2', label: '폭발 2' },
        { id: 'wave', label: '물결' },
        { id: 'double_wave', label: '이중 물결' },
        { id: 'ribbon', label: '리본' },
    ]},
    { name: '설명선', shapes: [
        { id: 'wedge_rect_callout', label: '사각 설명선' },
        { id: 'wedge_round_rect_callout', label: '둥근 사각 설명선' },
        { id: 'wedge_ellipse_callout', label: '타원 설명선' },
        { id: 'cloud_callout', label: '구름 설명선' },
        { id: 'border_callout_1', label: '테두리 설명선 1' },
        { id: 'border_callout_2', label: '테두리 설명선 2' },
        { id: 'border_callout_3', label: '테두리 설명선 3' },
    ]},
    { name: '선', shapes: [
        { id: 'line', label: '직선' },
        { id: 'arrow', label: '화살표' },
    ]},
];

function initShapePicker() {
    var container = document.getElementById('shapeDropdown');
    if (!container) return;
    var html = '';
    SHAPE_CATEGORIES.forEach(function(cat) {
        html += '<div class="shape-picker-category">' + cat.name + '</div>';
        html += '<div class="shape-picker-grid">';
        cat.shapes.forEach(function(shape) {
            html += '<button type="button" class="shape-picker-item" title="' + shape.label + '" onclick="addShapeObject(\'' + shape.id + '\')">';
            html += _getShapeIconSVG(shape.id);
            html += '</button>';
        });
        html += '</div>';
    });
    container.innerHTML = html;
}

function createShapeSVG(obj) {
    const s = obj.shape_style || {};
    const w = obj.width;
    const h = obj.height;
    const fill = s.fill_opacity === 0 || s.fill_color === 'transparent' ? 'none' : (s.fill_color || '#4A90D9');
    const fillOpacity = s.fill_opacity !== undefined ? s.fill_opacity : 1;
    const stroke = s.stroke_color || '#333333';
    const strokeW = s.stroke_width !== undefined ? s.stroke_width : 2;
    const dashMap = { dashed: '8,4', dotted: '2,4' };
    const dashAttr = dashMap[s.stroke_dash] ? `stroke-dasharray="${dashMap[s.stroke_dash]}"` : '';
    const half = strokeW / 2;
    const commonFill = `fill="${fill}" fill-opacity="${fillOpacity}"`;
    const commonStroke = `stroke="${stroke}" stroke-width="${strokeW}" ${dashAttr}`;

    let inner = '';

    switch (s.shape_type) {
        case 'rounded_rectangle': {
            const rx = s.border_radius || 12;
            inner = `<rect x="${half}" y="${half}" width="${Math.max(0, w - strokeW)}" height="${Math.max(0, h - strokeW)}"
                      rx="${rx}" ry="${rx}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'ellipse':
            inner = `<ellipse cx="${w / 2}" cy="${h / 2}" rx="${Math.max(0, w / 2 - half)}" ry="${Math.max(0, h / 2 - half)}"
                      ${commonFill} ${commonStroke}/>`;
            break;
        case 'line':
            inner = `<line x1="0" y1="${h / 2}" x2="${w}" y2="${h / 2}"
                      stroke="${stroke}" stroke-width="${strokeW}" ${dashAttr}/>`;
            break;
        case 'arrow': {
            const mid = 'arrow_' + obj.obj_id.replace(/[^a-zA-Z0-9]/g, '_');
            const markerStart = s.arrow_head === 'both' ?
                `<marker id="${mid}_s" markerWidth="10" markerHeight="7" refX="0" refY="3.5" orient="auto-start-reverse">
                    <polygon points="10 0, 10 7, 0 3.5" fill="${stroke}"/>
                </marker>` : '';
            inner = `<defs>
                <marker id="${mid}" markerWidth="10" markerHeight="7" refX="10" refY="3.5" orient="auto">
                    <polygon points="0 0, 10 3.5, 0 7" fill="${stroke}"/>
                </marker>${markerStart}
            </defs>
            <line x1="0" y1="${h / 2}" x2="${w}" y2="${h / 2}"
                  stroke="${stroke}" stroke-width="${strokeW}" ${dashAttr}
                  marker-end="url(#${mid})"
                  ${s.arrow_head === 'both' ? `marker-start="url(#${mid}_s)"` : ''}/>`;
            break;
        }
        // 직사각형 변형
        case 'snip_1_rect': {
            var c = Math.min(w, h) * 0.2;
            inner = `<polygon points="${half},${half} ${w - c - half},${half} ${w - half},${c + half} ${w - half},${h - half} ${half},${h - half}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'snip_2_diag_rect': {
            var c = Math.min(w, h) * 0.2;
            inner = `<polygon points="${c + half},${half} ${w - half},${half} ${w - half},${h - c - half} ${w - c - half},${h - half} ${half},${h - half} ${half},${c + half}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'round_1_rect': {
            var r = Math.min(w, h) * 0.2;
            inner = `<path d="M${half},${half} H${w - r - half} A${r},${r} 0 0 1 ${w - half},${r + half} V${h - half} H${half} Z" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'round_2_diag_rect': {
            var r = Math.min(w, h) * 0.2;
            inner = `<path d="M${r + half},${half} H${w - half} V${h - r - half} A${r},${r} 0 0 1 ${w - r - half},${h - half} H${half} V${r + half} A${r},${r} 0 0 1 ${r + half},${half} Z" ${commonFill} ${commonStroke}/>`;
            break;
        }
        // 기본 도형
        case 'triangle':
            inner = `<polygon points="${w / 2},${half} ${w - half},${h - half} ${half},${h - half}" ${commonFill} ${commonStroke}/>`;
            break;
        case 'right_triangle':
            inner = `<polygon points="${half},${h - half} ${w - half},${h - half} ${half},${half}" ${commonFill} ${commonStroke}/>`;
            break;
        case 'parallelogram': {
            var off = w * 0.2;
            inner = `<polygon points="${off},${half} ${w - half},${half} ${w - off},${h - half} ${half},${h - half}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'trapezoid': {
            var off = w * 0.2;
            inner = `<polygon points="${off},${half} ${w - off},${half} ${w - half},${h - half} ${half},${h - half}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'diamond':
            inner = `<polygon points="${w / 2},${half} ${w - half},${h / 2} ${w / 2},${h - half} ${half},${h / 2}" ${commonFill} ${commonStroke}/>`;
            break;
        case 'pentagon':
            inner = `<polygon points="${_regPolygonPts(w / 2, h / 2, Math.min(w, h) / 2 - half, 5, 0)}" ${commonFill} ${commonStroke}/>`;
            break;
        case 'hexagon':
            inner = `<polygon points="${_regPolygonPts(w / 2, h / 2, Math.min(w, h) / 2 - half, 6, 0)}" ${commonFill} ${commonStroke}/>`;
            break;
        case 'heptagon':
            inner = `<polygon points="${_regPolygonPts(w / 2, h / 2, Math.min(w, h) / 2 - half, 7, 0)}" ${commonFill} ${commonStroke}/>`;
            break;
        case 'octagon':
            inner = `<polygon points="${_regPolygonPts(w / 2, h / 2, Math.min(w, h) / 2 - half, 8, 0)}" ${commonFill} ${commonStroke}/>`;
            break;
        case 'decagon':
            inner = `<polygon points="${_regPolygonPts(w / 2, h / 2, Math.min(w, h) / 2 - half, 10, 0)}" ${commonFill} ${commonStroke}/>`;
            break;
        case 'dodecagon':
            inner = `<polygon points="${_regPolygonPts(w / 2, h / 2, Math.min(w, h) / 2 - half, 12, 0)}" ${commonFill} ${commonStroke}/>`;
            break;
        case 'cross': {
            var arm = w * 0.3;
            var armH = h * 0.3;
            inner = `<polygon points="${arm},${half} ${w - arm},${half} ${w - arm},${armH} ${w - half},${armH} ${w - half},${h - armH} ${w - arm},${h - armH} ${w - arm},${h - half} ${arm},${h - half} ${arm},${h - armH} ${half},${h - armH} ${half},${armH} ${arm},${armH}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'donut': {
            var cx = w / 2, cy = h / 2;
            var rx1 = w / 2 - half, ry1 = h / 2 - half;
            var rx2 = rx1 * 0.5, ry2 = ry1 * 0.5;
            inner = `<path d="M${cx + rx1},${cy} A${rx1},${ry1} 0 1 0 ${cx - rx1},${cy} A${rx1},${ry1} 0 1 0 ${cx + rx1},${cy} Z M${cx + rx2},${cy} A${rx2},${ry2} 0 1 1 ${cx - rx2},${cy} A${rx2},${ry2} 0 1 1 ${cx + rx2},${cy} Z" fill-rule="evenodd" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'no_smoking': {
            var cx = w / 2, cy = h / 2;
            var rx = w / 2 - half, ry = h / 2 - half;
            inner = `<ellipse cx="${cx}" cy="${cy}" rx="${rx}" ry="${ry}" ${commonFill} ${commonStroke}/>`;
            var dx = rx * 0.707, dy = ry * 0.707;
            inner += `<line x1="${cx - dx}" y1="${cy - dy}" x2="${cx + dx}" y2="${cy + dy}" stroke="${stroke}" stroke-width="${strokeW}" ${dashAttr}/>`;
            break;
        }
        case 'block_arc': {
            inner = `<path d="M${half},${h - half} A${w / 2 - half},${h / 2 - half} 0 0 1 ${w - half},${h - half}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'heart': {
            var cx = w / 2, top = h * 0.2;
            inner = `<path d="M${cx},${h - half} C${cx},${h * 0.65} ${half},${h * 0.45} ${half},${top + h * 0.1} A${w * 0.25},${h * 0.2} 0 0 1 ${cx},${top} A${w * 0.25},${h * 0.2} 0 0 1 ${w - half},${top + h * 0.1} C${w - half},${h * 0.45} ${cx},${h * 0.65} ${cx},${h - half} Z" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'lightning_bolt': {
            inner = `<polygon points="${w * 0.55},${half} ${w * 0.25},${h * 0.45} ${w * 0.45},${h * 0.45} ${w * 0.4},${h - half} ${w * 0.75},${h * 0.5} ${w * 0.55},${h * 0.5} ${w * 0.6},${half}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'sun': {
            var cx = w / 2, cy = h / 2;
            var r = Math.min(w, h) * 0.22;
            inner = `<circle cx="${cx}" cy="${cy}" r="${r}" ${commonFill} ${commonStroke}/>`;
            var rayR = Math.min(w, h) * 0.45;
            for (var i = 0; i < 8; i++) {
                var a = i * Math.PI / 4;
                inner += `<line x1="${cx + r * Math.cos(a)}" y1="${cy + r * Math.sin(a)}" x2="${cx + rayR * Math.cos(a)}" y2="${cy + rayR * Math.sin(a)}" stroke="${stroke}" stroke-width="${strokeW}" ${dashAttr}/>`;
            }
            break;
        }
        case 'moon': {
            inner = `<path d="M${w * 0.7},${half} A${w * 0.4},${h * 0.45} 0 1 0 ${w * 0.7},${h - half} A${w * 0.25},${h * 0.35} 0 0 1 ${w * 0.7},${half} Z" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'cloud': {
            inner = `<path d="M${w * 0.25},${h * 0.75} A${w * 0.17},${w * 0.17} 0 0 1 ${w * 0.12},${h * 0.45} A${w * 0.22},${w * 0.22} 0 0 1 ${w * 0.38},${h * 0.22} A${w * 0.2},${w * 0.2} 0 0 1 ${w * 0.68},${h * 0.2} A${w * 0.17},${w * 0.17} 0 0 1 ${w * 0.87},${h * 0.42} A${w * 0.15},${w * 0.15} 0 0 1 ${w * 0.82},${h * 0.72} Z" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'smiley_face': {
            var cx = w / 2, cy = h / 2;
            var rx = w / 2 - half, ry = h / 2 - half;
            inner = `<ellipse cx="${cx}" cy="${cy}" rx="${rx}" ry="${ry}" ${commonFill} ${commonStroke}/>`;
            inner += `<circle cx="${cx - rx * 0.3}" cy="${cy - ry * 0.2}" r="${Math.min(rx, ry) * 0.08}" fill="${stroke}"/>`;
            inner += `<circle cx="${cx + rx * 0.3}" cy="${cy - ry * 0.2}" r="${Math.min(rx, ry) * 0.08}" fill="${stroke}"/>`;
            inner += `<path d="M${cx - rx * 0.35},${cy + ry * 0.2} Q${cx},${cy + ry * 0.55} ${cx + rx * 0.35},${cy + ry * 0.2}" fill="none" stroke="${stroke}" stroke-width="${strokeW}" ${dashAttr}/>`;
            break;
        }
        case 'folded_corner': {
            var fold = Math.min(w, h) * 0.2;
            inner = `<polygon points="${half},${half} ${w - half},${half} ${w - half},${h - fold - half} ${w - fold - half},${h - half} ${half},${h - half}" ${commonFill} ${commonStroke}/>`;
            inner += `<polygon points="${w - fold - half},${h - half} ${w - fold - half},${h - fold - half} ${w - half},${h - fold - half}" fill="#e0e0e0" stroke="${stroke}" stroke-width="${Math.max(1, strokeW * 0.7)}"/>`;
            break;
        }
        case 'frame': {
            var bw = Math.min(w, h) * 0.12;
            inner = `<rect x="${half}" y="${half}" width="${w - strokeW}" height="${h - strokeW}" ${commonFill} ${commonStroke}/>`;
            inner += `<rect x="${bw}" y="${bw}" width="${w - bw * 2}" height="${h - bw * 2}" fill="white" fill-opacity="0.8" stroke="${stroke}" stroke-width="${Math.max(1, strokeW * 0.7)}"/>`;
            break;
        }
        case 'teardrop': {
            inner = `<path d="M${w / 2},${half} C${w / 2},${half} ${w - half},${h * 0.35} ${w - half},${h * 0.6} A${w / 2 - half},${h * 0.4 - half} 0 1 1 ${half},${h * 0.6} C${half},${h * 0.35} ${w / 2},${half} ${w / 2},${half} Z" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'plaque': {
            var r = Math.min(w, h) * 0.12;
            inner = `<path d="M${half},${r + half} A${r},${r} 0 0 1 ${r + half},${half} H${w - r - half} A${r},${r} 0 0 1 ${w - half},${r + half} V${h - r - half} A${r},${r} 0 0 1 ${w - r - half},${h - half} H${r + half} A${r},${r} 0 0 1 ${half},${h - r - half} Z" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'brace_pair': {
            var q = Math.min(w, h) * 0.15;
            inner = `<path d="M${w * 0.15},${half} Q${half},${half} ${half},${q + half} V${h / 2 - q} Q${half},${h / 2} ${half - q * 0.3},${h / 2} Q${half},${h / 2} ${half},${h / 2 + q} V${h - q - half} Q${half},${h - half} ${w * 0.15},${h - half}" fill="none" ${commonStroke}/>`;
            inner += `<path d="M${w * 0.85},${half} Q${w - half},${half} ${w - half},${q + half} V${h / 2 - q} Q${w - half},${h / 2} ${w - half + q * 0.3},${h / 2} Q${w - half},${h / 2} ${w - half},${h / 2 + q} V${h - q - half} Q${w - half},${h - half} ${w * 0.85},${h - half}" fill="none" ${commonStroke}/>`;
            break;
        }
        case 'bracket_pair': {
            inner = `<path d="M${w * 0.15},${half} H${half} V${h - half} H${w * 0.15}" fill="none" ${commonStroke}/>`;
            inner += `<path d="M${w * 0.85},${half} H${w - half} V${h - half} H${w * 0.85}" fill="none" ${commonStroke}/>`;
            break;
        }
        // 블록 화살표
        case 'right_arrow': {
            var shaft = h * 0.3;
            inner = `<polygon points="${half},${h / 2 - shaft} ${w * 0.65},${h / 2 - shaft} ${w * 0.65},${half} ${w - half},${h / 2} ${w * 0.65},${h - half} ${w * 0.65},${h / 2 + shaft} ${half},${h / 2 + shaft}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'left_arrow': {
            var shaft = h * 0.3;
            inner = `<polygon points="${w - half},${h / 2 - shaft} ${w * 0.35},${h / 2 - shaft} ${w * 0.35},${half} ${half},${h / 2} ${w * 0.35},${h - half} ${w * 0.35},${h / 2 + shaft} ${w - half},${h / 2 + shaft}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'up_arrow': {
            var shaft = w * 0.3;
            inner = `<polygon points="${w / 2 - shaft},${h - half} ${w / 2 - shaft},${h * 0.35} ${half},${h * 0.35} ${w / 2},${half} ${w - half},${h * 0.35} ${w / 2 + shaft},${h * 0.35} ${w / 2 + shaft},${h - half}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'down_arrow': {
            var shaft = w * 0.3;
            inner = `<polygon points="${w / 2 - shaft},${half} ${w / 2 - shaft},${h * 0.65} ${half},${h * 0.65} ${w / 2},${h - half} ${w - half},${h * 0.65} ${w / 2 + shaft},${h * 0.65} ${w / 2 + shaft},${half}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'left_right_arrow': {
            var shaft = h * 0.3, head = w * 0.2;
            inner = `<polygon points="${half},${h / 2} ${head},${half} ${head},${h / 2 - shaft} ${w - head},${h / 2 - shaft} ${w - head},${half} ${w - half},${h / 2} ${w - head},${h - half} ${w - head},${h / 2 + shaft} ${head},${h / 2 + shaft} ${head},${h - half}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'up_down_arrow': {
            var shaft = w * 0.3, head = h * 0.2;
            inner = `<polygon points="${w / 2},${half} ${w - half},${head} ${w / 2 + shaft},${head} ${w / 2 + shaft},${h - head} ${w - half},${h - head} ${w / 2},${h - half} ${half},${h - head} ${w / 2 - shaft},${h - head} ${w / 2 - shaft},${head} ${half},${head}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'quad_arrow': {
            var s2 = Math.min(w, h) * 0.12, head = Math.min(w, h) * 0.25;
            var cx = w / 2, cy = h / 2;
            inner = `<polygon points="${cx},${half} ${cx + head},${head + half} ${cx + s2},${head + half} ${cx + s2},${cy - s2} ${cx + head + s2},${cy - s2} ${cx + head + s2},${cy - head} ${w - half},${cy} ${cx + head + s2},${cy + head} ${cx + head + s2},${cy + s2} ${cx + s2},${cy + s2} ${cx + s2},${cy + head + s2} ${cx + head},${cy + head + s2} ${cx},${h - half} ${cx - head},${cy + head + s2} ${cx - s2},${cy + head + s2} ${cx - s2},${cy + s2} ${cx - head - s2},${cy + s2} ${cx - head - s2},${cy + head} ${half},${cy} ${cx - head - s2},${cy - head} ${cx - head - s2},${cy - s2} ${cx - s2},${cy - s2} ${cx - s2},${head + half} ${cx - head},${head + half}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'notched_right_arrow': {
            var shaft = h * 0.3, notch = w * 0.08;
            inner = `<polygon points="${half},${h / 2 - shaft} ${w * 0.65},${h / 2 - shaft} ${w * 0.65},${half} ${w - half},${h / 2} ${w * 0.65},${h - half} ${w * 0.65},${h / 2 + shaft} ${half},${h / 2 + shaft} ${notch + half},${h / 2}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'chevron': {
            var notch = w * 0.25;
            inner = `<polygon points="${half},${half} ${w - notch},${half} ${w - half},${h / 2} ${w - notch},${h - half} ${half},${h - half} ${notch},${h / 2}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'home_plate': {
            var point = w * 0.15;
            inner = `<polygon points="${half},${half} ${w - point},${half} ${w - half},${h / 2} ${w - point},${h - half} ${half},${h - half}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'striped_right_arrow': {
            var shaft = h * 0.3;
            inner = `<polygon points="${w * 0.2},${h / 2 - shaft} ${w * 0.65},${h / 2 - shaft} ${w * 0.65},${half} ${w - half},${h / 2} ${w * 0.65},${h - half} ${w * 0.65},${h / 2 + shaft} ${w * 0.2},${h / 2 + shaft}" ${commonFill} ${commonStroke}/>`;
            var stripeGap = w * 0.04;
            for (var si = 0; si < 3; si++) {
                var sx = half + si * (stripeGap + strokeW);
                inner += `<line x1="${sx}" y1="${h / 2 - shaft}" x2="${sx}" y2="${h / 2 + shaft}" stroke="${stroke}" stroke-width="${strokeW}"/>`;
            }
            break;
        }
        case 'bent_arrow': {
            var shaft = h * 0.12;
            inner = `<path d="M${half},${h - half} V${h * 0.35} H${w * 0.55} V${h * 0.15} L${w - half},${h * 0.35} L${w * 0.55},${h * 0.55} V${h * 0.35 + shaft * 2} H${half + shaft * 3} V${h - half} Z" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'u_turn_arrow': {
            var shaftW = w * 0.12;
            inner = `<path d="M${half},${h - half} V${h * 0.3} A${w * 0.3},${h * 0.25} 0 0 1 ${w * 0.6 + half},${h * 0.3} V${h * 0.5} H${w * 0.8} L${w * 0.6},${h * 0.7} L${w * 0.4},${h * 0.5} H${w * 0.6 + half - shaftW} V${h * 0.3} A${w * 0.18},${h * 0.13} 0 0 0 ${half + shaftW},${h * 0.3} V${h - half} Z" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'circular_arrow': {
            var cx = w / 2, cy = h / 2;
            var rx = w * 0.4, ry = h * 0.4;
            var rx2 = w * 0.28, ry2 = h * 0.28;
            inner = `<path d="M${cx + rx},${cy} A${rx},${ry} 0 1 0 ${cx},${cy - ry} L${cx},${cy - ry - h * 0.08} A${rx + w * 0.08},${ry + h * 0.08} 0 1 1 ${cx + rx + w * 0.08},${cy}" ${commonFill} ${commonStroke}/>`;
            inner += `<polygon points="${cx + rx + w * 0.12},${cy - h * 0.08} ${cx + rx + w * 0.12},${cy + h * 0.08} ${cx + rx - w * 0.02},${cy}" fill="${fill === 'none' ? stroke : fill}" stroke="${stroke}" stroke-width="${strokeW}"/>`;
            break;
        }
        // 수학
        case 'math_plus': {
            var arm = Math.min(w, h) * 0.18;
            inner = `<polygon points="${w / 2 - arm},${half} ${w / 2 + arm},${half} ${w / 2 + arm},${h / 2 - arm} ${w - half},${h / 2 - arm} ${w - half},${h / 2 + arm} ${w / 2 + arm},${h / 2 + arm} ${w / 2 + arm},${h - half} ${w / 2 - arm},${h - half} ${w / 2 - arm},${h / 2 + arm} ${half},${h / 2 + arm} ${half},${h / 2 - arm} ${w / 2 - arm},${h / 2 - arm}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'math_minus':
            inner = `<rect x="${half}" y="${h * 0.35}" width="${w - strokeW}" height="${h * 0.3}" ${commonFill} ${commonStroke}/>`;
            break;
        case 'math_multiply': {
            var cx = w / 2, cy = h / 2, arm = Math.min(w, h) * 0.08, len = Math.min(w, h) * 0.4;
            inner = `<path d="M${cx - arm},${half} L${cx + arm},${half} L${cx + arm},${cy - arm - len * 0.3} L${w - half},${cy - arm} L${w - half},${cy + arm} L${cx + arm + len * 0.3},${cy + arm} L${w - half - arm},${h - half - arm} L${cx + arm},${h - half} L${cx - arm},${h - half} L${cx - arm},${cy + arm + len * 0.3} L${half},${cy + arm} L${half},${cy - arm} L${cx - arm - len * 0.3},${cy - arm} Z" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'math_divide': {
            var dotR = Math.min(w, h) * 0.08;
            inner = `<rect x="${half}" y="${h * 0.4}" width="${w - strokeW}" height="${h * 0.2}" ${commonFill} ${commonStroke}/>`;
            inner += `<circle cx="${w / 2}" cy="${h * 0.2}" r="${dotR}" ${commonFill} ${commonStroke}/>`;
            inner += `<circle cx="${w / 2}" cy="${h * 0.8}" r="${dotR}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'math_equal':
            inner = `<rect x="${w * 0.1}" y="${h * 0.28}" width="${w * 0.8}" height="${h * 0.12}" ${commonFill} ${commonStroke}/>`;
            inner += `<rect x="${w * 0.1}" y="${h * 0.6}" width="${w * 0.8}" height="${h * 0.12}" ${commonFill} ${commonStroke}/>`;
            break;
        case 'math_not_equal':
            inner = `<rect x="${w * 0.1}" y="${h * 0.28}" width="${w * 0.8}" height="${h * 0.12}" ${commonFill} ${commonStroke}/>`;
            inner += `<rect x="${w * 0.1}" y="${h * 0.6}" width="${w * 0.8}" height="${h * 0.12}" ${commonFill} ${commonStroke}/>`;
            inner += `<line x1="${w * 0.65}" y1="${h * 0.15}" x2="${w * 0.35}" y2="${h * 0.85}" stroke="${stroke}" stroke-width="${strokeW + 1}" ${dashAttr}/>`;
            break;
        // 별
        case 'star_4_point':
            inner = `<polygon points="${_starPts(w / 2, h / 2, Math.min(w, h) / 2 - half, Math.min(w, h) * 0.18, 4, 0)}" ${commonFill} ${commonStroke}/>`;
            break;
        case 'star_5_point':
            inner = `<polygon points="${_starPts(w / 2, h / 2, Math.min(w, h) / 2 - half, Math.min(w, h) * 0.18, 5, 0)}" ${commonFill} ${commonStroke}/>`;
            break;
        case 'star_6_point':
            inner = `<polygon points="${_starPts(w / 2, h / 2, Math.min(w, h) / 2 - half, Math.min(w, h) * 0.22, 6, 0)}" ${commonFill} ${commonStroke}/>`;
            break;
        case 'star_8_point':
            inner = `<polygon points="${_starPts(w / 2, h / 2, Math.min(w, h) / 2 - half, Math.min(w, h) * 0.22, 8, 0)}" ${commonFill} ${commonStroke}/>`;
            break;
        case 'star_10_point':
            inner = `<polygon points="${_starPts(w / 2, h / 2, Math.min(w, h) / 2 - half, Math.min(w, h) * 0.25, 10, 0)}" ${commonFill} ${commonStroke}/>`;
            break;
        case 'star_12_point':
            inner = `<polygon points="${_starPts(w / 2, h / 2, Math.min(w, h) / 2 - half, Math.min(w, h) * 0.25, 12, 0)}" ${commonFill} ${commonStroke}/>`;
            break;
        case 'star_16_point':
            inner = `<polygon points="${_starPts(w / 2, h / 2, Math.min(w, h) / 2 - half, Math.min(w, h) * 0.3, 16, 0)}" ${commonFill} ${commonStroke}/>`;
            break;
        case 'star_24_point':
            inner = `<polygon points="${_starPts(w / 2, h / 2, Math.min(w, h) / 2 - half, Math.min(w, h) * 0.32, 24, 0)}" ${commonFill} ${commonStroke}/>`;
            break;
        case 'star_32_point':
            inner = `<polygon points="${_starPts(w / 2, h / 2, Math.min(w, h) / 2 - half, Math.min(w, h) * 0.35, 32, 0)}" ${commonFill} ${commonStroke}/>`;
            break;
        case 'explosion_1': {
            inner = `<polygon points="${w * 0.5},${half} ${w * 0.58},${h * 0.3} ${w * 0.85},${h * 0.15} ${w * 0.68},${h * 0.38} ${w - half},${h * 0.5} ${w * 0.68},${h * 0.6} ${w * 0.85},${h * 0.85} ${w * 0.58},${h * 0.68} ${w * 0.5},${h - half} ${w * 0.42},${h * 0.68} ${w * 0.15},${h * 0.85} ${w * 0.32},${h * 0.6} ${half},${h * 0.5} ${w * 0.32},${h * 0.38} ${w * 0.15},${h * 0.15} ${w * 0.42},${h * 0.3}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'explosion_2': {
            inner = `<polygon points="${w * 0.5},${half} ${w * 0.54},${h * 0.25} ${w * 0.72},${h * 0.1} ${w * 0.62},${h * 0.3} ${w * 0.9},${h * 0.25} ${w * 0.7},${h * 0.42} ${w - half},${h * 0.5} ${w * 0.72},${h * 0.58} ${w * 0.9},${h * 0.75} ${w * 0.62},${h * 0.65} ${w * 0.72},${h * 0.9} ${w * 0.54},${h * 0.72} ${w * 0.5},${h - half} ${w * 0.46},${h * 0.72} ${w * 0.28},${h * 0.9} ${w * 0.38},${h * 0.65} ${w * 0.1},${h * 0.75} ${w * 0.28},${h * 0.58} ${half},${h * 0.5} ${w * 0.3},${h * 0.42} ${w * 0.1},${h * 0.25} ${w * 0.38},${h * 0.3} ${w * 0.28},${h * 0.1} ${w * 0.46},${h * 0.25}" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'wave': {
            inner = `<path d="M${half},${h * 0.35} C${w * 0.2},${half} ${w * 0.35},${half} ${w * 0.5},${h * 0.35} C${w * 0.65},${h * 0.5} ${w * 0.8},${h * 0.5} ${w - half},${h * 0.35} V${h * 0.65} C${w * 0.8},${h - half} ${w * 0.65},${h - half} ${w * 0.5},${h * 0.65} C${w * 0.35},${h * 0.5} ${w * 0.2},${h * 0.5} ${half},${h * 0.65} Z" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'double_wave': {
            inner = `<path d="M${half},${h * 0.3} C${w * 0.15},${h * 0.15} ${w * 0.35},${h * 0.15} ${w * 0.5},${h * 0.3} S${w * 0.85},${h * 0.45} ${w - half},${h * 0.3}" fill="none" ${commonStroke}/>`;
            inner += `<path d="M${half},${h * 0.7} C${w * 0.15},${h * 0.55} ${w * 0.35},${h * 0.55} ${w * 0.5},${h * 0.7} S${w * 0.85},${h * 0.85} ${w - half},${h * 0.7}" fill="none" ${commonStroke}/>`;
            break;
        }
        case 'ribbon': {
            inner = `<path d="M${half},${h * 0.3} L${w * 0.12},${half} H${w * 0.88} L${w - half},${h * 0.3} L${w * 0.88},${h * 0.5} V${h - half} H${w * 0.12} V${h * 0.5} Z" ${commonFill} ${commonStroke}/>`;
            break;
        }
        // 설명선
        case 'wedge_rect_callout': {
            inner = `<path d="M${half},${half} H${w - half} V${h * 0.65} H${w * 0.55} L${w * 0.4},${h - half} L${w * 0.35},${h * 0.65} H${half} Z" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'wedge_round_rect_callout': {
            var r = Math.min(w, h) * 0.08;
            inner = `<path d="M${r + half},${half} H${w - r - half} A${r},${r} 0 0 1 ${w - half},${r + half} V${h * 0.65 - r} A${r},${r} 0 0 1 ${w - r - half},${h * 0.65} H${w * 0.55} L${w * 0.4},${h - half} L${w * 0.35},${h * 0.65} H${r + half} A${r},${r} 0 0 1 ${half},${h * 0.65 - r} V${r + half} A${r},${r} 0 0 1 ${r + half},${half} Z" ${commonFill} ${commonStroke}/>`;
            break;
        }
        case 'wedge_ellipse_callout': {
            inner = `<ellipse cx="${w / 2}" cy="${h * 0.4}" rx="${w / 2 - half}" ry="${h * 0.35}" ${commonFill} ${commonStroke}/>`;
            inner += `<path d="M${w * 0.42},${h * 0.7} L${w * 0.35},${h - half} L${w * 0.52},${h * 0.72}" fill="${fill === 'none' ? 'none' : fill}" stroke="${stroke}" stroke-width="${strokeW}" ${dashAttr}/>`;
            break;
        }
        case 'cloud_callout': {
            inner = `<path d="M${w * 0.25},${h * 0.65} A${w * 0.17},${w * 0.15} 0 0 1 ${w * 0.12},${h * 0.38} A${w * 0.2},${w * 0.2} 0 0 1 ${w * 0.38},${h * 0.18} A${w * 0.18},${w * 0.18} 0 0 1 ${w * 0.65},${h * 0.15} A${w * 0.15},${w * 0.15} 0 0 1 ${w * 0.85},${h * 0.35} A${w * 0.13},${w * 0.13} 0 0 1 ${w * 0.8},${h * 0.62} Z" ${commonFill} ${commonStroke}/>`;
            inner += `<circle cx="${w * 0.3}" cy="${h * 0.78}" r="${Math.min(w, h) * 0.04}" fill="${fill === 'none' ? 'none' : fill}" stroke="${stroke}" stroke-width="${strokeW}"/>`;
            inner += `<circle cx="${w * 0.22}" cy="${h * 0.88}" r="${Math.min(w, h) * 0.025}" fill="${fill === 'none' ? 'none' : fill}" stroke="${stroke}" stroke-width="${strokeW}"/>`;
            break;
        }
        case 'border_callout_1': {
            inner = `<rect x="${half}" y="${half}" width="${w - strokeW}" height="${h * 0.65}" ${commonFill} ${commonStroke}/>`;
            inner += `<line x1="${w * 0.3}" y1="${h * 0.65}" x2="${w * 0.3}" y2="${h - half}" stroke="${stroke}" stroke-width="${strokeW}" ${dashAttr}/>`;
            break;
        }
        case 'border_callout_2': {
            inner = `<rect x="${half}" y="${half}" width="${w - strokeW}" height="${h * 0.65}" ${commonFill} ${commonStroke}/>`;
            inner += `<polyline points="${w * 0.3},${h * 0.65} ${w * 0.3},${h * 0.8} ${w * 0.45},${h - half}" fill="none" stroke="${stroke}" stroke-width="${strokeW}" ${dashAttr}/>`;
            break;
        }
        case 'border_callout_3': {
            inner = `<rect x="${half}" y="${half}" width="${w - strokeW}" height="${h * 0.65}" ${commonFill} ${commonStroke}/>`;
            inner += `<polyline points="${w * 0.3},${h * 0.65} ${w * 0.4},${h * 0.8} ${w * 0.2},${h - half}" fill="none" stroke="${stroke}" stroke-width="${strokeW}" ${dashAttr}/>`;
            break;
        }
        default: // rectangle
            inner = `<rect x="${half}" y="${half}" width="${Math.max(0, w - strokeW)}" height="${Math.max(0, h - strokeW)}"
                      ${commonFill} ${commonStroke}/>`;
            break;
    }

    return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" style="overflow:visible;">${inner}</svg>`;
}

function updateShapeStyle(prop, value) {
    // 멀티셀렉트: 선택된 모든 도형 객체에 적용
    const shapeObjs = state.selectedObjects.filter(o => o.obj_type === 'shape');
    if (shapeObjs.length === 0) return;

    shapeObjs.forEach(obj => {
        if (!obj.shape_style) obj.shape_style = {};
        obj.shape_style[prop] = value;

        // 도형 유형 변경 시 라인 관련 기본값 조정
        if (prop === 'shape_type') {
            const isLine = (value === 'line' || value === 'arrow');
            if (isLine) {
                obj.shape_style.fill_opacity = 0;
                obj.shape_style.fill_color = 'transparent';
                if (obj.height > 20) {
                    obj.height = 4;
                    $(`[data-obj-id="${obj.obj_id}"]`).css('height', '4px');
                }
            }
            if (value === 'arrow') {
                obj.shape_style.arrow_head = obj.shape_style.arrow_head || 'end';
            }
            if (value === 'rounded_rectangle') {
                obj.shape_style.border_radius = obj.shape_style.border_radius || 12;
            }
        }

        // SVG 재렌더
        const el = $(`[data-obj-id="${obj.obj_id}"]`);
        el.find('svg').remove();
        el.prepend(createShapeSVG(obj));
    });

    updateShapePropertyVisibility();

    // 색상 변경 시 최근 색상에 추가
    if (prop === 'fill_color' || prop === 'stroke_color') {
        addRecentColor(value);
    }
}

function updateShapePropertyVisibility() {
    const s = state.selectedObject?.shape_style || {};
    const isLine = s.shape_type === 'line' || s.shape_type === 'arrow';
    $('#shapeFillRow').toggle(!isLine);
    $('#borderRadiusGroup').toggle(s.shape_type === 'rounded_rectangle');
    $('#arrowHeadGroup').toggle(s.shape_type === 'arrow');

    // 도형 유형 변경 시 라인이 아닌 블록 화살표는 fill 표시
    var blockArrows = ['right_arrow','left_arrow','up_arrow','down_arrow','left_right_arrow','up_down_arrow','quad_arrow','notched_right_arrow','chevron','home_plate','striped_right_arrow','bent_arrow','u_turn_arrow','circular_arrow'];
    if (blockArrows.indexOf(s.shape_type) >= 0) {
        $('#shapeFillRow').show();
    }
}

// ============ 슬라이드 메타 ============
function calcMetaFromObjects(objects) {
    var hasTitle = false, hasGovernance = false, descCount = 0;
    (objects || []).forEach(function(obj) {
        if (obj.role === 'title') hasTitle = true;
        if (obj.role === 'governance') hasGovernance = true;
        if (obj.role === 'description') descCount++;
    });
    return { has_title: hasTitle, has_governance: hasGovernance, description_count: descCount };
}

function updateSlideMetaUI() {
    $('#metaContentType').val(state.slideMeta.content_type || 'body');
    $('#metaLayout').val(state.slideMeta.layout || '');
    _showMetaTypeGuide(state.slideMeta.content_type || 'body');

    // 자동 계산된 메타 정보 표시
    var auto = calcMetaFromObjects(state.objects);
    var parts = [];
    if (auto.has_title) parts.push('제목');
    if (auto.has_governance) parts.push('거버넌스');
    if (auto.description_count > 0) parts.push('설명 ' + auto.description_count + '개');
    if (parts.length > 0) {
        $('#metaAutoSummary').text('자동 감지: ' + parts.join(' · '));
        $('#metaAutoInfo').show();
    } else {
        $('#metaAutoInfo').hide();
    }
}

function updateSlideMeta() {
    const newLayout = $('#metaLayout').val();
    const oldLayout = state._previousLayout || '';

    var auto = calcMetaFromObjects(state.objects);
    state.slideMeta = {
        content_type: $('#metaContentType').val(),
        layout: newLayout,
        has_title: auto.has_title,
        has_governance: auto.has_governance,
        description_count: auto.description_count,
    };
    _showMetaTypeGuide(state.slideMeta.content_type);
    updateSlideMetaUI();

    // 레이아웃이 변경되었으면 프리셋 적용
    if (newLayout && newLayout !== oldLayout) {
        applyLayoutPreset(newLayout);
    }
}

// ============ 레이아웃 프리셋 ============

function _makeTextObj({ x, y, width, height, text, font_size, color, bold, align, role, placeholder }) {
    return {
        obj_id: generateObjId(),
        obj_type: 'text',
        x: x,
        y: y,
        width: width,
        height: height,
        text_content: text || '텍스트를 입력하세요',
        text_style: {
            font_family: 'Arial',
            font_size: font_size || 16,
            color: color || '#000000',
            bold: bold || false,
            italic: false,
            align: align || 'left',
        },
        role: role || null,
        placeholder: placeholder || null,
    };
}

function generateLayoutPreset(layout) {
    switch (layout) {
        case 'cover':
            return {
                content_type: 'title_slide',
                has_title: true,
                has_governance: false,
                description_count: 0,
                objects: [
                    _makeTextObj({ x: 80, y: 180, width: 800, height: 80, text: '프레젠테이션 제목',
                        font_size: 36, bold: true, align: 'center', color: '#222222',
                        role: 'title', placeholder: 'main_title' }),
                    _makeTextObj({ x: 80, y: 280, width: 800, height: 50, text: '부제목 또는 작성자',
                        font_size: 18, align: 'center', color: '#666666',
                        role: 'subtitle', placeholder: 'sub_title' }),
                ],
            };

        case 'numbered_list':
            return {
                content_type: 'toc',
                has_title: true,
                has_governance: false,
                description_count: 5,
                objects: [
                    _makeTextObj({ x: 60, y: 30, width: 840, height: 50, text: '목차',
                        font_size: 28, bold: true, color: '#222222',
                        role: 'title', placeholder: 'main_title' }),
                    ...[0, 1, 2, 3, 4].map(i => _makeTextObj({
                        x: 80, y: 110 + (i * 60), width: 800, height: 40,
                        text: '0' + (i + 1) + '. 목차 항목 ' + (i + 1),
                        font_size: 16, color: '#333333',
                        role: 'description', placeholder: 'toc_item_' + i,
                    })),
                ],
            };

        case 'divider':
            return {
                content_type: 'section_divider',
                has_title: true,
                has_governance: true,
                description_count: 0,
                objects: [
                    _makeTextObj({ x: 80, y: 180, width: 800, height: 40, text: '01',
                        font_size: 14, bold: true, align: 'center', color: '#888888',
                        role: 'governance', placeholder: 'section_num' }),
                    _makeTextObj({ x: 80, y: 230, width: 800, height: 70, text: '섹션 제목',
                        font_size: 32, bold: true, align: 'center', color: '#222222',
                        role: 'title', placeholder: 'section_title' }),
                    _makeTextObj({ x: 80, y: 310, width: 800, height: 40, text: '섹션 부제목',
                        font_size: 16, align: 'center', color: '#666666',
                        role: 'subtitle', placeholder: 'section_subtitle' }),
                ],
            };

        case 'single_column':
            return {
                content_type: 'body',
                has_title: true,
                has_governance: true,
                description_count: 3,
                objects: [
                    _makeTextObj({ x: 60, y: 30, width: 840, height: 50, text: '슬라이드 제목',
                        font_size: 24, bold: true, color: '#222222',
                        role: 'title', placeholder: 'main_title' }),
                    _makeTextObj({ x: 60, y: 80, width: 840, height: 30, text: '거버넌스 요약문',
                        font_size: 12, color: '#888888',
                        role: 'governance', placeholder: 'governance_text' }),
                    ...[0, 1, 2].map(i => [
                        _makeTextObj({ x: 60, y: 130 + (i * 100), width: 840, height: 30,
                            text: (i + 1) + '. 부제목',
                            font_size: 16, bold: true, color: '#2d5a8e',
                            role: 'subtitle', placeholder: 'subtitle_' + i }),
                        _makeTextObj({ x: 60, y: 165 + (i * 100), width: 840, height: 45,
                            text: '설명 텍스트 ' + (i + 1),
                            font_size: 13, color: '#444444',
                            role: 'description', placeholder: 'desc_' + i }),
                    ]).flat(),
                ],
            };

        case 'two_column': {
            const cols = [
                { x: 60, items: [0, 1] },
                { x: 500, items: [2, 3] },
            ];
            const pairs = [];
            cols.forEach(function (col) {
                col.items.forEach(function (idx, row) {
                    pairs.push(_makeTextObj({
                        x: col.x, y: 130 + (row * 110), width: 400, height: 30,
                        text: (idx + 1) + '. 부제목',
                        font_size: 15, bold: true, color: '#2d5a8e',
                        role: 'subtitle', placeholder: 'subtitle_' + idx,
                    }));
                    pairs.push(_makeTextObj({
                        x: col.x, y: 165 + (row * 110), width: 400, height: 55,
                        text: '설명 텍스트 ' + (idx + 1),
                        font_size: 12, color: '#444444',
                        role: 'description', placeholder: 'desc_' + idx,
                    }));
                });
            });
            return {
                content_type: 'body',
                has_title: true,
                has_governance: true,
                description_count: 4,
                objects: [
                    _makeTextObj({ x: 60, y: 30, width: 840, height: 50, text: '슬라이드 제목',
                        font_size: 24, bold: true, color: '#222222',
                        role: 'title', placeholder: 'main_title' }),
                    _makeTextObj({ x: 60, y: 80, width: 840, height: 30, text: '거버넌스 요약문',
                        font_size: 12, color: '#888888',
                        role: 'governance', placeholder: 'governance_text' }),
                    ...pairs,
                ],
            };
        }

        case 'grid': {
            const cells = [
                { x: 60, y: 130 },
                { x: 500, y: 130 },
                { x: 60, y: 290 },
                { x: 500, y: 290 },
            ];
            const cardObjs = [];
            cells.forEach(function (cell, i) {
                cardObjs.push(_makeTextObj({
                    x: cell.x, y: cell.y, width: 400, height: 28,
                    text: (i + 1) + '. 키워드',
                    font_size: 15, bold: true, color: '#2d5a8e',
                    role: 'subtitle', placeholder: 'subtitle_' + i,
                }));
                cardObjs.push(_makeTextObj({
                    x: cell.x, y: cell.y + 32, width: 400, height: 60,
                    text: '설명 텍스트 ' + (i + 1),
                    font_size: 12, color: '#444444',
                    role: 'description', placeholder: 'desc_' + i,
                }));
            });
            return {
                content_type: 'body',
                has_title: true,
                has_governance: true,
                description_count: 4,
                objects: [
                    _makeTextObj({ x: 60, y: 30, width: 840, height: 50, text: '슬라이드 제목',
                        font_size: 24, bold: true, color: '#222222',
                        role: 'title', placeholder: 'main_title' }),
                    _makeTextObj({ x: 60, y: 80, width: 840, height: 30, text: '거버넌스 요약문',
                        font_size: 12, color: '#888888',
                        role: 'governance', placeholder: 'governance_text' }),
                    ...cardObjs,
                ],
            };
        }

        case 'closing':
            return {
                content_type: 'closing',
                has_title: true,
                has_governance: false,
                description_count: 0,
                objects: [
                    _makeTextObj({ x: 80, y: 190, width: 800, height: 80, text: '감사합니다',
                        font_size: 36, bold: true, align: 'center', color: '#222222',
                        role: 'title', placeholder: 'closing_title' }),
                    _makeTextObj({ x: 80, y: 290, width: 800, height: 40, text: '연락처 또는 메시지',
                        font_size: 16, align: 'center', color: '#666666',
                        role: 'subtitle', placeholder: 'closing_subtitle' }),
                ],
            };

        default:
            return null;
    }
}

function getLayoutDisplayName(layout) {
    var names = {
        cover: '커버 (타이틀용)',
        numbered_list: '번호 리스트 (목차용)',
        divider: '디바이더 (간지용)',
        single_column: '단일 컬럼',
        two_column: '2단 컬럼',
        grid: '카드 그리드',
        closing: '마무리',
    };
    return names[layout] || layout;
}

function applyLayoutPreset(layout) {
    if (!layout) return;

    // 기존 오브젝트가 있으면 확인
    if (state.objects.length > 0) {
        if (!confirm('레이아웃 프리셋을 적용하면 기존 오브젝트가 모두 삭제됩니다.\n계속하시겠습니까?')) {
            // 드롭다운을 이전 값으로 복원
            $('#metaLayout').val(state._previousLayout || '');
            state.slideMeta.layout = state._previousLayout || '';
            return;
        }
    }

    var presetData = generateLayoutPreset(layout);
    if (!presetData) return;

    // 오브젝트 초기화 및 프리셋 적용
    state.objects = presetData.objects;
    state.selectedObject = null;
    state.selectedObjects = [];
    removeGroupBoundingBox();

    // 슬라이드 메타 자동 업데이트 (has_title 등은 오브젝트에서 자동 계산)
    var auto = calcMetaFromObjects(state.objects);
    state.slideMeta.content_type = presetData.content_type;
    state.slideMeta.has_title = auto.has_title;
    state.slideMeta.has_governance = auto.has_governance;
    state.slideMeta.description_count = auto.description_count;
    state.slideMeta.layout = layout;

    // 이전 레이아웃 갱신
    state._previousLayout = layout;

    // UI 갱신
    updateSlideMetaUI();
    renderCanvas();

    // 속성 패널 초기화 (선택된 오브젝트 없음)
    $('#objProperties').hide();
    $('#textProperties').hide();
    $('#shapeProperties').hide();

    showToast(getLayoutDisplayName(layout) + ' 프리셋이 적용되었습니다', 'success');
}

function _showMetaTypeGuide(contentType) {
    const guides = {
        title_slide: '<strong>타이틀 슬라이드</strong><br>• "프레젠테이션 제목" → 필드 유형: <strong>제목</strong><br>• "부제목/작성자" → 필드 유형: <strong>부제목</strong>',
        toc: '<strong>목차 슬라이드</strong><br>• "목차" 텍스트 → 필드 유형: <strong>제목</strong><br>• 각 목차 항목 (1, 2, 3...) → 필드 유형: <strong>설명</strong><br>• 항목 수만큼 설명 텍스트 오브젝트를 배치하세요',
        section_divider: '<strong>섹션 간지</strong><br>• "섹션 제목" → 필드 유형: <strong>제목</strong><br>• "섹션 번호" (01, 02...) → 필드 유형: <strong>거버넌스</strong><br>• "섹션 부제목" → 필드 유형: <strong>부제목</strong>',
        body: '<strong>본문 슬라이드</strong><br>• "슬라이드 제목" → 필드 유형: <strong>제목</strong><br>• "거버넌스/요약문" → 필드 유형: <strong>거버넌스</strong><br>• 각 핵심 키워드 → 필드 유형: <strong>부제목</strong><br>• 각 상세 설명 → 필드 유형: <strong>설명</strong><br>• 부제목과 설명은 순서대로 1:1 매핑됩니다',
        closing: '<strong>마무리 슬라이드</strong><br>• "감사합니다" 등 → 필드 유형: <strong>제목</strong><br>• "연락처/메시지" → 필드 유형: <strong>부제목</strong> 또는 <strong>설명</strong>',
        image_slide: '<strong>이미지 슬라이드</strong><br>• 사용자가 업로드한 이미지가 <strong>이미지 영역</strong> 오브젝트에 자동 배치됩니다<br>• 툴바의 "이미지 영역" 버튼으로 영역을 추가하세요',
    };
    const guide = guides[contentType];
    if (guide) {
        $('#metaTypeGuide').html(guide).show();
    } else {
        $('#metaTypeGuide').hide();
    }
}

// ============ 사용 현황 대시보드 ============

const dashState = {
    currentSection: 'overview',
    charts: {},
    users: { page: 1, total: 0 },
    projects: { page: 1, total: 0 },
    activity: { page: 1, total: 0 },
};

function showDashboard() {
    $('#dashboardOverlay').show();
    switchDashboardSection('overview');
}

function closeDashboard() {
    $('#dashboardOverlay').hide();
    Object.values(dashState.charts).forEach(c => { try { c.destroy(); } catch(e) {} });
    dashState.charts = {};
}

function switchDashboardSection(section) {
    dashState.currentSection = section;
    $('.dashboard-nav-item').removeClass('active');
    $(`.dashboard-nav-item[data-section="${section}"]`).addClass('active');
    $('.dashboard-section').removeClass('active');
    $(`#dashSection-${section}`).addClass('active');
    switch (section) {
        case 'overview': loadDashboardOverview(); break;
        case 'users': loadDashboardUsers(); break;
        case 'projects': loadDashboardProjects(); break;
        case 'activity': loadDashboardActivity(); break;
    }
}

// ---- 전체 현황 ----
async function loadDashboardOverview() {
    try {
        const data = await apiGet('/api/admin/dashboard/overview');
        renderOverviewStats(data);
        renderOverviewCharts(data);
    } catch (e) {
        showToast('대시보드 로딩 실패', 'error');
    }
}

const STAT_COLORS = ['#4f7cac', '#4caf50', '#ff9800', '#9c27b0', '#1976d2', '#26a69a'];
const STAT_BGS = ['#e8f0fe', '#e8f5e9', '#fff3e0', '#f3e5f5', '#e3f2fd', '#e0f2f1'];
const STAT_ICONS = ['&#128101;', '&#128196;', '&#128197;', '&#9889;', '&#128203;', '&#128206;'];

function renderOverviewStats(data) {
    const cards = [
        { label: '전체 사용자', value: data.total_users, sub: `활성 사용자 ${data.active_users}명`, idx: 0 },
        { label: '전체 프로젝트', value: data.total_projects, sub: `생성완료 ${data.generated_projects}건`, idx: 1 },
        { label: '이번 달 프로젝트', value: data.month_projects, sub: `오늘 ${data.today_projects}건`, idx: 2 },
        { label: 'AI 생성 건수', value: data.total_generations, sub: `슬라이드 ${data.generation_breakdown.slides} | 엑셀 ${data.generation_breakdown.excel} | Word ${data.generation_breakdown.docx}`, idx: 3 },
        { label: '템플릿', value: data.total_templates, sub: `공개 ${data.published_templates}개`, idx: 4 },
        { label: '총 리소스', value: data.total_resources, sub: '등록된 전체 리소스', idx: 5 },
    ];
    let html = '';
    cards.forEach(c => {
        html += `<div class="stat-card">
            <div class="stat-card-header">
                <div class="stat-label">${c.label}</div>
                <div class="stat-icon" style="background:${STAT_BGS[c.idx]};color:${STAT_COLORS[c.idx]};">${STAT_ICONS[c.idx]}</div>
            </div>
            <div class="stat-value" style="color:${STAT_COLORS[c.idx]};">${c.value.toLocaleString()}</div>
            <div class="stat-sub">${c.sub}</div>
        </div>`;
    });
    $('#overviewStatCards').html(html);
}

function renderOverviewCharts(data) {
    Object.values(dashState.charts).forEach(c => { try { c.destroy(); } catch(e) {} });
    dashState.charts = {};

    // 일별 추이
    const dLabels = data.daily_trend.map(d => d._id.slice(5));
    const dValues = data.daily_trend.map(d => d.count);
    dashState.charts.daily = new Chart(document.getElementById('chartDailyTrend'), {
        type: 'line',
        data: {
            labels: dLabels,
            datasets: [{
                label: '프로젝트',
                data: dValues,
                borderColor: '#4f7cac',
                backgroundColor: 'rgba(79,124,172,0.08)',
                fill: true,
                tension: 0.35,
                pointRadius: 3,
                pointBackgroundColor: '#4f7cac',
                borderWidth: 2,
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } }, x: { ticks: { maxTicksLimit: 15, font: { size: 11 } } } }
        }
    });

    // 프로젝트 유형
    const typeMap = { slide: '슬라이드', excel: '엑셀', onlyoffice_pptx: 'OO PPTX', onlyoffice_xlsx: 'OO XLSX', onlyoffice_docx: 'OO DOCX' };
    const tColors = ['#4f7cac', '#4caf50', '#ff9800', '#9c27b0', '#1976d2'];
    dashState.charts.type = new Chart(document.getElementById('chartProjectType'), {
        type: 'doughnut',
        data: {
            labels: data.project_type_distribution.map(d => typeMap[d._id] || d._id),
            datasets: [{ data: data.project_type_distribution.map(d => d.count), backgroundColor: tColors.slice(0, data.project_type_distribution.length), borderWidth: 0 }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { position: 'bottom', labels: { font: { size: 11 }, padding: 12 } } },
            cutout: '60%'
        }
    });

    // 리소스 유형
    const resMap = { file: '파일', text: '텍스트', web: '웹 검색', url: 'URL' };
    const rColors = ['#1976d2', '#4caf50', '#ff9800', '#9c27b0'];
    dashState.charts.resource = new Chart(document.getElementById('chartResourceType'), {
        type: 'doughnut',
        data: {
            labels: data.resource_type_distribution.map(d => resMap[d._id] || d._id),
            datasets: [{ data: data.resource_type_distribution.map(d => d.count), backgroundColor: rColors.slice(0, data.resource_type_distribution.length), borderWidth: 0 }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { position: 'bottom', labels: { font: { size: 11 }, padding: 12 } } },
            cutout: '60%'
        }
    });

    // 생성 현황
    const gb = data.generation_breakdown;
    dashState.charts.gen = new Chart(document.getElementById('chartGenerationBreakdown'), {
        type: 'bar',
        data: {
            labels: ['슬라이드', '엑셀', 'Word'],
            datasets: [{ label: '생성 건수', data: [gb.slides, gb.excel, gb.docx], backgroundColor: ['#4f7cac', '#4caf50', '#ff9800'], borderRadius: 6, barThickness: 40 }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } } }
        }
    });
}

// ---- 사용자 현황 ----
async function loadDashboardUsers(page) {
    dashState.users.page = page || 1;
    const search = $('#dashUserSearch').val() || '';
    try {
        const data = await apiGet(`/api/admin/dashboard/users?search=${encodeURIComponent(search)}&page=${dashState.users.page}&limit=20`);
        dashState.users.total = data.total;
        renderDashUserList(data.users);
        renderDashPagination('dashUserPagination', data.total, dashState.users.page, 20, 'loadDashboardUsers');
    } catch (e) {
        showToast('사용자 목록 로딩 실패', 'error');
    }
}

function renderDashUserList(users) {
    if (!users.length) {
        $('#dashUserList').html('<div class="dashboard-empty">검색 결과가 없습니다</div>');
        return;
    }
    const avatarColors = ['#4f7cac', '#4caf50', '#ff9800', '#9c27b0', '#1976d2', '#e91e63', '#00897b', '#5c6bc0'];
    let html = `<div class="dashboard-list-header">
        <div style="width:48px;"></div>
        <div style="flex:1;">이름</div>
        <div style="flex:1;">부서</div>
        <div style="flex:1.2;">이메일</div>
        <div style="width:80px;text-align:center;">프로젝트</div>
        <div style="width:80px;text-align:center;">역할</div>
        <div style="width:120px;">최근 활동</div>
    </div>`;
    users.forEach((u, i) => {
        const initial = (u.nm || '?')[0];
        const color = avatarColors[i % avatarColors.length];
        const roleBadge = u.role === 'admin'
            ? '<span class="dash-badge role-admin">관리자</span>'
            : '<span class="dash-badge role-user">일반</span>';
        const lastAct = u.last_activity ? new Date(u.last_activity).toLocaleDateString('ko-KR') : '-';
        html += `<div class="dashboard-list-item clickable" onclick="showUserDetail('${escapeHtml(u.ky)}')">
            <div class="dash-avatar" style="background:${color}22;color:${color};">${escapeHtml(initial)}</div>
            <div style="flex:1;margin-left:12px;font-weight:500;">${escapeHtml(u.nm)}</div>
            <div style="flex:1;color:#666;font-size:12px;">${escapeHtml(u.dp || '-')}</div>
            <div style="flex:1.2;color:#888;font-size:12px;">${escapeHtml(u.em || '-')}</div>
            <div style="width:80px;text-align:center;font-weight:600;color:#4f7cac;">${u.project_count}</div>
            <div style="width:80px;text-align:center;">${roleBadge}</div>
            <div style="width:120px;font-size:12px;color:#888;">${lastAct}</div>
        </div>`;
    });
    $('#dashUserList').html(html);
}

// ---- 프로젝트 현황 ----
async function loadDashboardProjects(page) {
    dashState.projects.page = page || 1;
    const search = $('#dashProjectSearch').val() || '';
    const typeFilter = $('#dashProjectTypeFilter').val() || '';
    const statusFilter = $('#dashProjectStatusFilter').val() || '';
    try {
        const data = await apiGet(`/api/admin/dashboard/projects?search=${encodeURIComponent(search)}&project_type=${typeFilter}&status=${statusFilter}&page=${dashState.projects.page}&limit=20`);
        dashState.projects.total = data.total;
        renderDashProjectList(data.projects);
        renderDashPagination('dashProjectPagination', data.total, dashState.projects.page, 20, 'loadDashboardProjects');
    } catch (e) {
        showToast('프로젝트 목록 로딩 실패', 'error');
    }
}

function renderDashProjectList(projects) {
    if (!projects.length) {
        $('#dashProjectList').html('<div class="dashboard-empty">검색 결과가 없습니다</div>');
        return;
    }
    const typeMap = { slide: '슬라이드', excel: '엑셀', onlyoffice_pptx: 'OO PPTX', onlyoffice_xlsx: 'OO XLSX', onlyoffice_docx: 'OO DOCX' };
    const typeCls = { slide: 'type-slide', excel: 'type-excel', onlyoffice_pptx: 'type-oo', onlyoffice_xlsx: 'type-oo', onlyoffice_docx: 'type-docx' };
    const statusMap = { draft: '초안', preparing: '준비중', generating: '생성중', generated: '생성완료' };
    const statusCls = { draft: 'status-draft', preparing: 'status-preparing', generating: 'status-generating', generated: 'status-generated' };

    let html = `<div class="dashboard-list-header">
        <div style="flex:2;">프로젝트명</div>
        <div style="flex:1;">작성자</div>
        <div style="width:100px;text-align:center;">유형</div>
        <div style="width:80px;text-align:center;">상태</div>
        <div style="width:60px;text-align:center;">슬라이드</div>
        <div style="width:60px;text-align:center;">리소스</div>
        <div style="width:110px;">생성일</div>
    </div>`;
    projects.forEach(p => {
        const pt = p.project_type || 'slide';
        const st = p.status || 'draft';
        const createdAt = p.created_at ? new Date(p.created_at).toLocaleDateString('ko-KR') : '-';
        html += `<div class="dashboard-list-item">
            <div style="flex:2;">
                <div style="font-weight:500;">${escapeHtml(p.name || '제목 없음')}</div>
                <div style="font-size:11px;color:#aaa;margin-top:2px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:300px;">${escapeHtml((p.description || '').substring(0, 80))}</div>
            </div>
            <div style="flex:1;font-size:12px;">
                <div style="font-weight:500;">${escapeHtml(p._user_name || '-')}</div>
                <div style="color:#999;font-size:11px;">${escapeHtml(p._user_dept || '')}</div>
            </div>
            <div style="width:100px;text-align:center;"><span class="dash-badge ${typeCls[pt] || 'type-slide'}">${typeMap[pt] || pt}</span></div>
            <div style="width:80px;text-align:center;"><span class="dash-badge ${statusCls[st] || 'status-draft'}">${statusMap[st] || st}</span></div>
            <div style="width:60px;text-align:center;font-weight:600;color:#4f7cac;">${p._slide_count || 0}</div>
            <div style="width:60px;text-align:center;color:#666;">${p._resource_count || 0}</div>
            <div style="width:110px;font-size:12px;color:#888;">${createdAt}</div>
        </div>`;
    });
    $('#dashProjectList').html(html);
}

// ---- 활동 로그 ----
async function loadDashboardActivity(page) {
    dashState.activity.page = page || 1;
    const userSearch = $('#dashActivityUserSearch').val() || '';
    const dateFrom = $('#dashActivityDateFrom').val() || '';
    const dateTo = $('#dashActivityDateTo').val() || '';
    try {
        const data = await apiGet(`/api/admin/dashboard/activity?user_search=${encodeURIComponent(userSearch)}&date_from=${dateFrom}&date_to=${dateTo}&page=${dashState.activity.page}&limit=30`);
        dashState.activity.total = data.total;
        renderDashActivityList(data.activities);
        renderDashPagination('dashActivityPagination', data.total, dashState.activity.page, 30, 'loadDashboardActivity');
    } catch (e) {
        showToast('활동 로그 로딩 실패', 'error');
    }
}

function renderDashActivityList(activities) {
    if (!activities.length) {
        $('#dashActivityList').html('<div class="dashboard-empty">활동 기록이 없습니다</div>');
        return;
    }
    const typeMap = { slide: '슬라이드', excel: '엑셀', onlyoffice_pptx: 'OO PPTX', onlyoffice_xlsx: 'OO XLSX', onlyoffice_docx: 'OO DOCX' };
    const typeCls = { slide: 'type-slide', excel: 'type-excel', onlyoffice_pptx: 'type-oo', onlyoffice_xlsx: 'type-oo', onlyoffice_docx: 'type-docx' };
    const avatarColors = ['#4f7cac', '#4caf50', '#ff9800', '#9c27b0', '#1976d2', '#e91e63', '#00897b', '#5c6bc0'];

    let html = '';
    activities.forEach((a, i) => {
        const initial = (a.user_name || '?')[0];
        const color = avatarColors[i % avatarColors.length];
        const pt = a.project_type || 'slide';
        const updatedAt = a.updated_at
            ? new Date(a.updated_at).toLocaleString('ko-KR', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' })
            : '-';
        html += `<div class="dashboard-list-item" style="gap:12px;">
            <div class="dash-avatar" style="background:${color}22;color:${color};">${escapeHtml(initial)}</div>
            <div style="min-width:100px;">
                <div style="font-weight:500;font-size:13px;">${escapeHtml(a.user_name)}</div>
                <div style="font-size:11px;color:#999;">${escapeHtml(a.user_dept)}</div>
            </div>
            <div style="flex:1;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">
                <span style="color:#333;">${escapeHtml(a.project_name || '제목 없음')}</span>
            </div>
            <div style="width:80px;text-align:center;"><span class="dash-badge ${typeCls[pt] || 'type-slide'}">${typeMap[pt] || pt}</span></div>
            <div style="width:130px;text-align:right;font-size:12px;color:#888;">${updatedAt}</div>
        </div>`;
    });
    $('#dashActivityList').html(html);
}

// ---- 사용자 상세 대시보드 ----
async function showUserDetail(userKey) {
    // 사용자 상세 섹션으로 전환
    dashState.currentSection = 'userDetail';
    $('.dashboard-nav-item').removeClass('active');
    $('.dashboard-section').removeClass('active');
    $('#dashSection-userDetail').addClass('active');

    // 로딩 표시
    $('#userDetailProfile').html('<div class="dashboard-loading">사용자 정보를 불러오는 중...</div>');
    $('#userDetailStats').html('');
    $('#userDetailProjects').html('');

    try {
        const data = await apiGet(`/api/admin/dashboard/user/${encodeURIComponent(userKey)}`);
        renderUserDetailProfile(data);
        renderUserDetailStats(data);
        renderUserDetailCharts(data);
        renderUserDetailProjects(data.recent_projects);
    } catch (e) {
        showToast('사용자 정보 로딩 실패', 'error');
    }
}

function renderUserDetailProfile(data) {
    const u = data.user;
    const avatarColors = ['#4f7cac', '#4caf50', '#ff9800', '#9c27b0', '#1976d2', '#e91e63'];
    const color = avatarColors[Math.abs(hashCode(u.ky)) % avatarColors.length];
    const initial = (u.nm || '?')[0];
    const roleBadge = u.role === 'admin'
        ? '<span class="dash-badge role-admin" style="margin-left:8px;">관리자</span>'
        : '<span class="dash-badge role-user" style="margin-left:8px;">일반</span>';

    const html = `
        <div class="user-detail-avatar" style="background:${color}22;color:${color};">${escapeHtml(initial)}</div>
        <div class="user-detail-info">
            <div class="user-detail-name">${escapeHtml(u.nm)}${roleBadge}</div>
            <div class="user-detail-meta">
                <span>&#128188; ${escapeHtml(u.dp || '-')}</span>
                <span>&#9993; ${escapeHtml(u.em || '-')}</span>
                <span>&#128273; ${escapeHtml(u.ky)}</span>
            </div>
        </div>`;
    $('#userDetailProfile').html(html);
}

function hashCode(str) {
    let hash = 0;
    for (let i = 0; i < str.length; i++) {
        hash = ((hash << 5) - hash) + str.charCodeAt(i);
        hash |= 0;
    }
    return hash;
}

function renderUserDetailStats(data) {
    const gb = data.generation_breakdown;
    const cards = [
        { label: '전체 프로젝트', value: data.total_projects, sub: `이번 달 ${data.month_projects}건`, idx: 0 },
        { label: '생성 완료', value: data.generated_projects, sub: `${data.total_projects ? Math.round(data.generated_projects / data.total_projects * 100) : 0}% 완료율`, idx: 1 },
        { label: 'AI 생성', value: data.total_generations, sub: `슬라이드 ${gb.slides} | 엑셀 ${gb.excel} | Word ${gb.docx}`, idx: 3 },
        { label: '리소스', value: data.total_resources, sub: '등록된 리소스', idx: 5 },
    ];
    let html = '';
    cards.forEach(c => {
        html += `<div class="stat-card">
            <div class="stat-card-header">
                <div class="stat-label">${c.label}</div>
                <div class="stat-icon" style="background:${STAT_BGS[c.idx]};color:${STAT_COLORS[c.idx]};">${STAT_ICONS[c.idx]}</div>
            </div>
            <div class="stat-value" style="color:${STAT_COLORS[c.idx]};">${c.value.toLocaleString()}</div>
            <div class="stat-sub">${c.sub}</div>
        </div>`;
    });
    $('#userDetailStats').html(html);
}

function renderUserDetailCharts(data) {
    // 기존 사용자 차트 정리
    ['userDaily', 'userType', 'userStatus', 'userResource'].forEach(k => {
        if (dashState.charts[k]) { try { dashState.charts[k].destroy(); } catch(e) {} delete dashState.charts[k]; }
    });

    // 일별 추이
    const dLabels = data.daily_trend.map(d => d._id.slice(5));
    const dValues = data.daily_trend.map(d => d.count);
    dashState.charts.userDaily = new Chart(document.getElementById('userChartDailyTrend'), {
        type: 'line',
        data: {
            labels: dLabels,
            datasets: [{
                label: '프로젝트',
                data: dValues,
                borderColor: '#4f7cac',
                backgroundColor: 'rgba(79,124,172,0.08)',
                fill: true, tension: 0.35, pointRadius: 3, pointBackgroundColor: '#4f7cac', borderWidth: 2,
            }]
        },
        options: {
            responsive: true, maintainAspectRatio: false,
            plugins: { legend: { display: false } },
            scales: { y: { beginAtZero: true, ticks: { stepSize: 1 } }, x: { ticks: { maxTicksLimit: 15, font: { size: 11 } } } }
        }
    });

    // 프로젝트 유형
    const typeMap = { slide: '슬라이드', excel: '엑셀', onlyoffice_pptx: 'OO PPTX', onlyoffice_xlsx: 'OO XLSX', onlyoffice_docx: 'OO DOCX' };
    const tColors = ['#4f7cac', '#4caf50', '#ff9800', '#9c27b0', '#1976d2'];
    if (data.project_type_distribution.length) {
        dashState.charts.userType = new Chart(document.getElementById('userChartProjectType'), {
            type: 'doughnut',
            data: {
                labels: data.project_type_distribution.map(d => typeMap[d._id] || d._id),
                datasets: [{ data: data.project_type_distribution.map(d => d.count), backgroundColor: tColors, borderWidth: 0 }]
            },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { font: { size: 11 }, padding: 12 } } }, cutout: '60%' }
        });
    } else {
        document.getElementById('userChartProjectType').parentElement.innerHTML = '<div class="dashboard-empty" style="padding:60px 0;">데이터 없음</div>';
    }

    // 프로젝트 상태
    const statusMap = { draft: '초안', preparing: '준비중', generating: '생성중', generated: '생성완료' };
    const sColors = ['#bdbdbd', '#ffa726', '#42a5f5', '#66bb6a'];
    if (data.project_status_distribution.length) {
        dashState.charts.userStatus = new Chart(document.getElementById('userChartProjectStatus'), {
            type: 'doughnut',
            data: {
                labels: data.project_status_distribution.map(d => statusMap[d._id] || d._id),
                datasets: [{ data: data.project_status_distribution.map(d => d.count), backgroundColor: sColors, borderWidth: 0 }]
            },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { font: { size: 11 }, padding: 12 } } }, cutout: '60%' }
        });
    } else {
        document.getElementById('userChartProjectStatus').parentElement.innerHTML = '<div class="dashboard-empty" style="padding:60px 0;">데이터 없음</div>';
    }

    // 리소스 유형
    const resMap = { file: '파일', text: '텍스트', web: '웹 검색', url: 'URL' };
    const rColors = ['#1976d2', '#4caf50', '#ff9800', '#9c27b0'];
    if (data.resource_type_distribution.length) {
        dashState.charts.userResource = new Chart(document.getElementById('userChartResourceType'), {
            type: 'doughnut',
            data: {
                labels: data.resource_type_distribution.map(d => resMap[d._id] || d._id),
                datasets: [{ data: data.resource_type_distribution.map(d => d.count), backgroundColor: rColors, borderWidth: 0 }]
            },
            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { font: { size: 11 }, padding: 12 } } }, cutout: '60%' }
        });
    } else {
        document.getElementById('userChartResourceType').parentElement.innerHTML = '<div class="dashboard-empty" style="padding:60px 0;">데이터 없음</div>';
    }
}

function renderUserDetailProjects(projects) {
    if (!projects.length) {
        $('#userDetailProjects').html('<div class="dashboard-empty">프로젝트가 없습니다</div>');
        return;
    }
    const typeMap = { slide: '슬라이드', excel: '엑셀', onlyoffice_pptx: 'OO PPTX', onlyoffice_xlsx: 'OO XLSX', onlyoffice_docx: 'OO DOCX' };
    const typeCls = { slide: 'type-slide', excel: 'type-excel', onlyoffice_pptx: 'type-oo', onlyoffice_xlsx: 'type-oo', onlyoffice_docx: 'type-docx' };
    const statusMap = { draft: '초안', preparing: '준비중', generating: '생성중', generated: '생성완료' };
    const statusCls = { draft: 'status-draft', preparing: 'status-preparing', generating: 'status-generating', generated: 'status-generated' };

    let html = `<div class="dashboard-list-header">
        <div style="flex:2;">프로젝트명</div>
        <div style="width:100px;text-align:center;">유형</div>
        <div style="width:80px;text-align:center;">상태</div>
        <div style="width:60px;text-align:center;">슬라이드</div>
        <div style="width:120px;">생성일</div>
        <div style="width:120px;">최근 수정</div>
    </div>`;
    projects.forEach(p => {
        const pt = p.project_type || 'slide';
        const st = p.status || 'draft';
        const createdAt = p.created_at ? new Date(p.created_at).toLocaleDateString('ko-KR') : '-';
        const updatedAt = p.updated_at ? new Date(p.updated_at).toLocaleDateString('ko-KR') : '-';
        html += `<div class="dashboard-list-item">
            <div style="flex:2;font-weight:500;">${escapeHtml(p.name || '제목 없음')}</div>
            <div style="width:100px;text-align:center;"><span class="dash-badge ${typeCls[pt] || 'type-slide'}">${typeMap[pt] || pt}</span></div>
            <div style="width:80px;text-align:center;"><span class="dash-badge ${statusCls[st] || 'status-draft'}">${statusMap[st] || st}</span></div>
            <div style="width:60px;text-align:center;font-weight:600;color:#4f7cac;">${p.slide_count || 0}</div>
            <div style="width:120px;font-size:12px;color:#888;">${createdAt}</div>
            <div style="width:120px;font-size:12px;color:#888;">${updatedAt}</div>
        </div>`;
    });
    $('#userDetailProjects').html(html);
}

// ---- 페이지네이션 ----
function renderDashPagination(containerId, total, currentPage, pageSize, fnName) {
    const totalPages = Math.ceil(total / pageSize);
    if (totalPages <= 1) {
        $(`#${containerId}`).html(total > 0 ? `<span class="dash-page-info">총 ${total}건</span>` : '');
        return;
    }
    let html = '';
    if (currentPage > 1) html += `<button onclick="${fnName}(${currentPage - 1})">&#8249; 이전</button>`;
    const startP = Math.max(1, currentPage - 3);
    const endP = Math.min(totalPages, startP + 6);
    for (let i = startP; i <= endP; i++) {
        html += `<button class="${i === currentPage ? 'active' : ''}" onclick="${fnName}(${i})">${i}</button>`;
    }
    if (currentPage < totalPages) html += `<button onclick="${fnName}(${currentPage + 1})">다음 &#8250;</button>`;
    html += `<span class="dash-page-info">총 ${total}건</span>`;
    $(`#${containerId}`).html(html);
}

// =====================================================
// TABLE & CHART OBJECTS
// =====================================================

// -- Chart dropdown --
function toggleChartDropdown() {
    const dd = $('#chartDropdown');
    // Close other dropdowns
    $('#shapeDropdown').hide();
    dd.toggle();
}

// Close chart dropdown on outside click (add to existing document click handler if present)
$(document).on('click', function(e) {
    if (!$(e.target).closest('.shape-dropdown-wrapper').length) {
        $('#chartDropdown').hide();
    }
});

// -- Add Table Object --
function addTableObject() {
    if (!state.currentSlide) {
        showToast('먼저 슬라이드를 선택하세요', 'error');
        return;
    }

    const rows = 3, cols = 3;
    const data = [];
    for (let r = 0; r < rows; r++) {
        const row = [];
        for (let c = 0; c < cols; c++) {
            row.push(r === 0 ? `항목 ${c + 1}` : '');
        }
        data.push(row);
    }

    const obj = {
        obj_id: generateObjId(),
        obj_type: 'table',
        x: 100,
        y: 100,
        width: 400,
        height: 200,
        z_index: getNextZIndex(),
        table_style: {
            rows: rows,
            cols: cols,
            data: data,
            header_row: true,
            banded_rows: false,
            banded_cols: false,
            cell_styles: {},
            border_color: '#BFBFBF',
            border_width: 1,
            header_bg_color: '#4472C4',
            header_text_color: '#FFFFFF',
            font_family: 'Arial',
            font_size: 11,
            merged_cells: [],
        },
        role: 'table',
        placeholder: null,
    };

    state.objects.push(obj);
    const el = createObjectElement(obj);
    $('#canvas').append(el);
    selectObject(obj.obj_id);
    showToast('표가 추가되었습니다', 'success');
}

// -- Add Chart Object --
function addChartObject(chartType) {
    $('#chartDropdown').hide();

    if (!state.currentSlide) {
        showToast('먼저 슬라이드를 선택하세요', 'error');
        return;
    }

    // Enter drawing mode like shapes
    state.isDrawing = true;
    state.drawShapeType = null;
    state.drawChartType = chartType || 'bar';
    state.drawStart = null;
    $('#canvas').addClass('draw-mode');
    showToast('캔버스에서 차트 영역을 드래그하세요 (ESC: 취소)', 'info');
}

function createChartAtRect(chartType, x, y, width, height) {
    const sampleData = {
        labels: ['항목1', '항목2', '항목3', '항목4'],
        datasets: [{
            label: '시리즈 1',
            data: [65, 59, 80, 81],
            backgroundColor: ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000'],
        }]
    };

    const obj = {
        obj_id: generateObjId(),
        obj_type: 'chart',
        x: x,
        y: y,
        width: Math.max(width, 200),
        height: Math.max(height, 150),
        z_index: getNextZIndex(),
        chart_style: {
            chart_type: chartType,
            chart_data: sampleData,
            title: '',
            show_legend: true,
            show_grid: true,
            color_scheme: ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'],
        },
        role: 'chart',
        placeholder: null,
    };

    state.objects.push(obj);
    renderCanvas();
    selectObject(obj.obj_id);
    showToast('차트가 추가되었습니다', 'success');
}

// -- Render Table in Canvas --
function createTableHTML(obj) {
    const style = obj.table_style || {};
    const rows = style.rows || 3;
    const cols = style.cols || 3;
    const data = style.data || [];
    const headerRow = style.header_row !== false;
    const bandedRows = style.banded_rows || false;
    const headerBg = style.header_bg_color || '#4472C4';
    const headerText = style.header_text_color || '#FFFFFF';
    const borderColor = style.border_color || '#BFBFBF';
    const fontSize = style.font_size || 11;
    const cellStyles = style.cell_styles || {};
    const mergedCells = style.merged_cells || [];

    // Build merged cell lookup
    const mergedLookup = {};
    mergedCells.forEach(m => {
        for (let r = m.start_row; r <= m.end_row; r++) {
            for (let c = m.start_col; c <= m.end_col; c++) {
                if (r === m.start_row && c === m.start_col) {
                    mergedLookup[`${r}_${c}`] = {
                        rowspan: m.end_row - m.start_row + 1,
                        colspan: m.end_col - m.start_col + 1
                    };
                } else {
                    mergedLookup[`${r}_${c}`] = 'skip';
                }
            }
        }
    });

    let html = `<table style="border-collapse:collapse;width:100%;height:100%;table-layout:fixed;border:1px solid ${borderColor};">`;
    for (let r = 0; r < rows; r++) {
        html += '<tr>';
        for (let c = 0; c < cols; c++) {
            const key = `${r}_${c}`;
            const mergeInfo = mergedLookup[key];
            if (mergeInfo === 'skip') continue;

            const cellText = (r < data.length && c < data[r].length) ? data[r][c] : '';
            const cs = cellStyles[key] || {};

            let bgColor = cs.bg_color || '';
            let textColor = cs.text_color || '#000000';
            let bold = cs.bold || false;
            let textAlign = cs.text_align || 'left';

            if (headerRow && r === 0) {
                bgColor = bgColor || headerBg;
                textColor = cs.text_color || headerText;
                bold = cs.bold !== false;
            } else if (bandedRows && r % 2 === 0 && (!headerRow || r > 0)) {
                bgColor = bgColor || '#D9E2F3';
            }

            let attrs = `contenteditable="true" data-row="${r}" data-col="${c}"`;
            let styles = `border:1px solid ${borderColor};padding:2px 4px;font-size:${fontSize}px;`;
            styles += `text-align:${textAlign};`;
            if (bgColor) styles += `background:${bgColor};`;
            if (textColor) styles += `color:${textColor};`;
            if (bold) styles += `font-weight:bold;`;

            if (mergeInfo && typeof mergeInfo === 'object') {
                attrs += ` rowspan="${mergeInfo.rowspan}" colspan="${mergeInfo.colspan}"`;
            }

            html += `<td ${attrs} style="${styles}">${cellText}</td>`;
        }
        html += '</tr>';
    }
    html += '</table>';
    return html;
}

// -- Render Chart in Canvas --
const chartInstances = {};

function renderChartInElement(el, obj) {
    const chartStyle = obj.chart_style || {};
    const chartType = chartStyle.chart_type || 'bar';
    const chartData = chartStyle.chart_data || {};

    // Destroy previous instance if exists
    if (chartInstances[obj.obj_id]) {
        chartInstances[obj.obj_id].destroy();
        delete chartInstances[obj.obj_id];
    }

    const container = $(el);
    container.empty();

    if (!chartData.labels || !chartData.datasets || chartData.datasets.length === 0) {
        container.html('<div class="chart-placeholder"><svg xmlns="http://www.w3.org/2000/svg" width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.5"><rect x="3" y="3" width="18" height="18" rx="2"/><line x1="8" y1="15" x2="8" y2="9"/><line x1="12" y1="15" x2="12" y2="6"/><line x1="16" y1="15" x2="16" y2="12"/></svg><span>차트 영역</span></div>');
        return;
    }

    const canvas = document.createElement('canvas');
    container.append(canvas);

    // Map chart type for Chart.js
    let cjsType = chartType;
    if (chartType === 'area') cjsType = 'line';

    const datasets = (chartData.datasets || []).map((ds, i) => {
        const colors = chartStyle.color_scheme || ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
        const result = { ...ds };
        if (chartType === 'pie' || chartType === 'doughnut') {
            result.backgroundColor = (ds.data || []).map((_, j) => colors[j % colors.length]);
        } else {
            result.backgroundColor = ds.backgroundColor || colors[i % colors.length];
            result.borderColor = ds.borderColor || colors[i % colors.length];
        }
        if (chartType === 'area') {
            result.fill = true;
            result.backgroundColor = (colors[i % colors.length]) + '40';
        }
        return result;
    });

    const config = {
        type: cjsType,
        data: {
            labels: chartData.labels || [],
            datasets: datasets,
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            animation: false,
            plugins: {
                legend: { display: chartStyle.show_legend !== false },
                title: chartStyle.title ? { display: true, text: chartStyle.title } : { display: false },
            },
        }
    };

    if (chartType !== 'pie' && chartType !== 'doughnut' && chartType !== 'radar') {
        config.options.scales = {
            x: { display: chartStyle.show_grid !== false },
            y: { display: chartStyle.show_grid !== false, beginAtZero: true },
        };
    }

    try {
        chartInstances[obj.obj_id] = new Chart(canvas, config);
    } catch (e) {
        container.html('<div class="chart-placeholder"><span>차트 렌더링 오류</span></div>');
    }
}

// -- Table Context Menu --
let tableCtxState = { objId: null, row: -1, col: -1, selectedCells: [] };

function showTableContextMenu(e, objId) {
    e.preventDefault();
    e.stopPropagation();

    const cell = $(e.target).closest('td, th');
    if (!cell.length) return;

    tableCtxState.objId = objId;
    tableCtxState.row = parseInt(cell.attr('data-row'));
    tableCtxState.col = parseInt(cell.attr('data-col'));

    // Collect selected cells
    const tableEl = cell.closest('table');
    tableCtxState.selectedCells = [];
    tableEl.find('.cell-selected').each(function() {
        tableCtxState.selectedCells.push({
            row: parseInt($(this).attr('data-row')),
            col: parseInt($(this).attr('data-col'))
        });
    });
    if (tableCtxState.selectedCells.length === 0) {
        tableCtxState.selectedCells = [{ row: tableCtxState.row, col: tableCtxState.col }];
    }

    const menu = $('#tableContextMenu');
    menu.css({ left: e.clientX + 'px', top: e.clientY + 'px', display: 'block' });

    // Ensure menu doesn't go off screen
    setTimeout(() => {
        const rect = menu[0].getBoundingClientRect();
        if (rect.right > window.innerWidth) menu.css('left', (e.clientX - rect.width) + 'px');
        if (rect.bottom > window.innerHeight) menu.css('top', (e.clientY - rect.height) + 'px');
    }, 0);
}

$(document).on('click', function(e) {
    if (!$(e.target).closest('#tableContextMenu').length) {
        $('#tableContextMenu').hide();
    }
});

function tableCtxAction(action, value) {
    const obj = state.objects.find(o => o.obj_id === tableCtxState.objId);
    if (!obj || obj.obj_type !== 'table') return;

    const ts = obj.table_style;
    const r = tableCtxState.row;
    const c = tableCtxState.col;

    switch (action) {
        case 'insertRowAbove':
            ts.rows++;
            ts.data.splice(r, 0, new Array(ts.cols).fill(''));
            _shiftCellStyles(ts, 'row', r, 1);
            break;
        case 'insertRowBelow':
            ts.rows++;
            ts.data.splice(r + 1, 0, new Array(ts.cols).fill(''));
            _shiftCellStyles(ts, 'row', r + 1, 1);
            break;
        case 'insertColLeft':
            ts.cols++;
            ts.data.forEach(row => row.splice(c, 0, ''));
            _shiftCellStyles(ts, 'col', c, 1);
            break;
        case 'insertColRight':
            ts.cols++;
            ts.data.forEach(row => row.splice(c + 1, 0, ''));
            _shiftCellStyles(ts, 'col', c + 1, 1);
            break;
        case 'deleteRow':
            if (ts.rows <= 1) return;
            ts.rows--;
            ts.data.splice(r, 1);
            _shiftCellStyles(ts, 'row', r, -1);
            break;
        case 'deleteCol':
            if (ts.cols <= 1) return;
            ts.cols--;
            ts.data.forEach(row => row.splice(c, 1));
            _shiftCellStyles(ts, 'col', c, -1);
            break;
        case 'mergeCells': {
            const cells = tableCtxState.selectedCells;
            if (cells.length < 2) { showToast('2개 이상의 셀을 선택하세요', 'error'); return; }
            const minR = Math.min(...cells.map(c => c.row));
            const maxR = Math.max(...cells.map(c => c.row));
            const minC = Math.min(...cells.map(c => c.col));
            const maxC = Math.max(...cells.map(c => c.col));
            ts.merged_cells = ts.merged_cells || [];
            ts.merged_cells.push({ start_row: minR, start_col: minC, end_row: maxR, end_col: maxC });
            break;
        }
        case 'splitCell':
            ts.merged_cells = (ts.merged_cells || []).filter(m =>
                !(r >= m.start_row && r <= m.end_row && c >= m.start_col && c <= m.end_col)
            );
            break;
        case 'cellBgColor':
            tableCtxState.selectedCells.forEach(sc => {
                const key = `${sc.row}_${sc.col}`;
                ts.cell_styles[key] = ts.cell_styles[key] || {};
                ts.cell_styles[key].bg_color = value;
            });
            break;
        case 'cellTextColor':
            tableCtxState.selectedCells.forEach(sc => {
                const key = `${sc.row}_${sc.col}`;
                ts.cell_styles[key] = ts.cell_styles[key] || {};
                ts.cell_styles[key].text_color = value;
            });
            break;
        case 'alignLeft':
        case 'alignCenter':
        case 'alignRight': {
            const align = action.replace('align', '').toLowerCase();
            tableCtxState.selectedCells.forEach(sc => {
                const key = `${sc.row}_${sc.col}`;
                ts.cell_styles[key] = ts.cell_styles[key] || {};
                ts.cell_styles[key].text_align = align;
            });
            break;
        }
        case 'toggleBold':
            tableCtxState.selectedCells.forEach(sc => {
                const key = `${sc.row}_${sc.col}`;
                ts.cell_styles[key] = ts.cell_styles[key] || {};
                ts.cell_styles[key].bold = !ts.cell_styles[key].bold;
            });
            break;
    }

    // Re-render table
    const el = $(`[data-obj-id="${obj.obj_id}"]`);
    el.find('.table-wrapper').html(createTableHTML(obj));
    _bindTableEvents(el, obj);
    updatePropertiesPanel();
    $('#tableContextMenu').hide();
}

function _shiftCellStyles(ts, dimension, index, delta) {
    const newStyles = {};
    for (const [key, val] of Object.entries(ts.cell_styles || {})) {
        const [r, c] = key.split('_').map(Number);
        let nr = r, nc = c;
        if (dimension === 'row') {
            if (delta > 0 && r >= index) nr = r + delta;
            else if (delta < 0 && r === index) continue;
            else if (delta < 0 && r > index) nr = r + delta;
        } else {
            if (delta > 0 && c >= index) nc = c + delta;
            else if (delta < 0 && c === index) continue;
            else if (delta < 0 && c > index) nc = c + delta;
        }
        newStyles[`${nr}_${nc}`] = val;
    }
    ts.cell_styles = newStyles;

    // Shift merged cells
    ts.merged_cells = (ts.merged_cells || []).map(m => {
        const nm = { ...m };
        if (dimension === 'row') {
            if (delta > 0) {
                if (nm.start_row >= index) nm.start_row += delta;
                if (nm.end_row >= index) nm.end_row += delta;
            } else {
                if (nm.start_row > index) nm.start_row += delta;
                if (nm.end_row > index) nm.end_row += delta;
            }
        } else {
            if (delta > 0) {
                if (nm.start_col >= index) nm.start_col += delta;
                if (nm.end_col >= index) nm.end_col += delta;
            } else {
                if (nm.start_col > index) nm.start_col += delta;
                if (nm.end_col > index) nm.end_col += delta;
            }
        }
        return nm;
    }).filter(m => m.start_row >= 0 && m.start_col >= 0 && m.end_row < ts.rows && m.end_col < ts.cols);
}

// -- Table Cell Selection (for merge) --
function _bindTableEvents(el, obj) {
    const tableEl = el.find('table');
    let isSelecting = false;
    let startCell = null;

    // Cell editing -> sync to data
    tableEl.find('td').on('blur', function() {
        const r = parseInt($(this).attr('data-row'));
        const c = parseInt($(this).attr('data-col'));
        const text = $(this).text();
        const ts = obj.table_style;
        while (ts.data.length <= r) ts.data.push([]);
        while (ts.data[r].length <= c) ts.data[r].push('');
        ts.data[r][c] = text;
    });

    // Cell click -> prevent object drag when editing
    tableEl.find('td').on('mousedown', function(e) {
        e.stopPropagation();
        if (e.button === 2) return; // right click handled by context menu

        // Clear previous selection
        tableEl.find('.cell-selected').removeClass('cell-selected');
        isSelecting = true;
        startCell = { row: parseInt($(this).attr('data-row')), col: parseInt($(this).attr('data-col')) };
        $(this).addClass('cell-selected');
    });

    tableEl.find('td').on('mouseover', function() {
        if (!isSelecting) return;
        const endRow = parseInt($(this).attr('data-row'));
        const endCol = parseInt($(this).attr('data-col'));
        tableEl.find('.cell-selected').removeClass('cell-selected');
        const minR = Math.min(startCell.row, endRow);
        const maxR = Math.max(startCell.row, endRow);
        const minC = Math.min(startCell.col, endCol);
        const maxC = Math.max(startCell.col, endCol);
        tableEl.find('td').each(function() {
            const r = parseInt($(this).attr('data-row'));
            const c = parseInt($(this).attr('data-col'));
            if (r >= minR && r <= maxR && c >= minC && c <= maxC) {
                $(this).addClass('cell-selected');
            }
        });
    });

    $(document).on('mouseup.tablesel', function() {
        isSelecting = false;
    });

    // Right-click context menu
    tableEl.on('contextmenu', function(e) {
        showTableContextMenu(e, obj.obj_id);
        return false;
    });
}

// -- Table Properties --
function updateTableDimension(dim, value) {
    if (!state.selectedObject || state.selectedObject.obj_type !== 'table') return;
    const ts = state.selectedObject.table_style;
    const oldVal = ts[dim === 'rows' ? 'rows' : 'cols'];
    ts[dim] = value;

    if (dim === 'rows') {
        while (ts.data.length < value) ts.data.push(new Array(ts.cols).fill(''));
        if (ts.data.length > value) ts.data.length = value;
    } else {
        ts.data.forEach(row => {
            while (row.length < value) row.push('');
            if (row.length > value) row.length = value;
        });
    }

    const el = $(`[data-obj-id="${state.selectedObject.obj_id}"]`);
    el.find('.table-wrapper').html(createTableHTML(state.selectedObject));
    _bindTableEvents(el, state.selectedObject);
}

function updateTableOption(key, value) {
    if (!state.selectedObject || state.selectedObject.obj_type !== 'table') return;
    state.selectedObject.table_style[key] = value;
    const el = $(`[data-obj-id="${state.selectedObject.obj_id}"]`);
    el.find('.table-wrapper').html(createTableHTML(state.selectedObject));
    _bindTableEvents(el, state.selectedObject);
}

function applyTablePreset(preset) {
    if (!state.selectedObject || state.selectedObject.obj_type !== 'table') return;
    const ts = state.selectedObject.table_style;

    const presets = {
        default: { header_bg_color: '#4472C4', header_text_color: '#FFFFFF', border_color: '#BFBFBF', header_row: true, banded_rows: false },
        dark: { header_bg_color: '#333333', header_text_color: '#FFFFFF', border_color: '#555555', header_row: true, banded_rows: true },
        green: { header_bg_color: '#70AD47', header_text_color: '#FFFFFF', border_color: '#BFBFBF', header_row: true, banded_rows: false },
        orange: { header_bg_color: '#ED7D31', header_text_color: '#FFFFFF', border_color: '#BFBFBF', header_row: true, banded_rows: false },
        plain: { header_bg_color: '#FFFFFF', header_text_color: '#000000', border_color: '#999999', header_row: false, banded_rows: false },
    };

    const p = presets[preset] || presets.default;
    Object.assign(ts, p);
    ts.cell_styles = {}; // Reset per-cell overrides

    const el = $(`[data-obj-id="${state.selectedObject.obj_id}"]`);
    el.find('.table-wrapper').html(createTableHTML(state.selectedObject));
    _bindTableEvents(el, state.selectedObject);
    updatePropertiesPanel();
    showToast('테이블 스타일이 적용되었습니다', 'success');
}

// -- Chart Properties --
function updateChartOption(key, value) {
    if (!state.selectedObject || state.selectedObject.obj_type !== 'chart') return;
    state.selectedObject.chart_style[key] = value;
    const el = $(`[data-obj-id="${state.selectedObject.obj_id}"]`);
    renderChartInElement(el.find('.chart-wrapper')[0], state.selectedObject);
}

// -- Chart Data Editor Modal --
function openChartDataEditor() {
    if (!state.selectedObject || state.selectedObject.obj_type !== 'chart') return;
    const cs = state.selectedObject.chart_style;
    const cd = cs.chart_data || {};

    $('#chartTypeSelect').val(cs.chart_type || 'bar');
    $('#chartLabelsInput').val((cd.labels || []).join(', '));

    const container = $('#chartDatasetsContainer');
    container.empty();
    (cd.datasets || []).forEach((ds, i) => {
        container.append(`
            <div class="dataset-row" data-idx="${i}" style="border:1px solid #e5e7eb;border-radius:6px;padding:8px;margin-top:8px;">
                <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">
                    <input type="text" class="ds-label" value="${ds.label || 'Series ' + (i+1)}" placeholder="시리즈 이름" style="width:60%;padding:6px 8px;border:1px solid #ddd;border-radius:6px;font-size:13px;">
                    <button class="btn btn-danger btn-sm" onclick="removeChartDataset(${i})" style="padding:2px 8px;">삭제</button>
                </div>
                <input type="text" class="ds-data" value="${(ds.data || []).join(', ')}" placeholder="값 (쉼표 구분)" style="width:100%;padding:6px 8px;border:1px solid #ddd;border-radius:6px;font-size:13px;">
            </div>
        `);
    });

    $('#chartDataModal').show();
}

function closeChartDataEditor() {
    $('#chartDataModal').hide();
}

function addChartDataset() {
    const container = $('#chartDatasetsContainer');
    const idx = container.children().length;
    container.append(`
        <div class="dataset-row" data-idx="${idx}" style="border:1px solid #e5e7eb;border-radius:6px;padding:8px;margin-top:8px;">
            <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:4px;">
                <input type="text" class="ds-label" value="시리즈 ${idx+1}" placeholder="시리즈 이름" style="width:60%;padding:6px 8px;border:1px solid #ddd;border-radius:6px;font-size:13px;">
                <button class="btn btn-danger btn-sm" onclick="$(this).closest('.dataset-row').remove()" style="padding:2px 8px;">삭제</button>
            </div>
            <input type="text" class="ds-data" value="" placeholder="값 (쉼표 구분)" style="width:100%;padding:6px 8px;border:1px solid #ddd;border-radius:6px;font-size:13px;">
        </div>
    `);
}

function removeChartDataset(idx) {
    $(`#chartDatasetsContainer .dataset-row[data-idx="${idx}"]`).remove();
}

function applyChartData() {
    if (!state.selectedObject || state.selectedObject.obj_type !== 'chart') return;

    const labels = $('#chartLabelsInput').val().split(',').map(s => s.trim()).filter(Boolean);
    const datasets = [];
    const colors = state.selectedObject.chart_style.color_scheme || ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];

    $('#chartDatasetsContainer .dataset-row').each(function(i) {
        const label = $(this).find('.ds-label').val() || 'Series ' + (i + 1);
        const data = $(this).find('.ds-data').val().split(',').map(s => parseFloat(s.trim()) || 0);
        datasets.push({
            label: label,
            data: data,
            backgroundColor: colors[i % colors.length],
        });
    });

    const chartType = $('#chartTypeSelect').val();
    state.selectedObject.chart_style.chart_type = chartType;
    state.selectedObject.chart_style.chart_data = { labels, datasets };

    // Update chart type dropdown in properties panel
    $('#propChartType').val(chartType);

    const el = $(`[data-obj-id="${state.selectedObject.obj_id}"]`);
    renderChartInElement(el.find('.chart-wrapper')[0], state.selectedObject);
    closeChartDataEditor();
    showToast('차트 데이터가 적용되었습니다', 'success');
}

// ============ 유틸리티 ============
function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function showToast(message, type) {
    const toast = $(`<div class="toast ${type || ''}">${escapeHtml(message)}</div>`);
    $('body').append(toast);
    setTimeout(() => toast.remove(), 3000);
}

// ============ 스킬 관리 (HTML 리포트) ============

function switchAdminTab(tab) {
    state.activeTab = tab;

    // 탭 버튼 활성화
    $('.sidebar-tab').removeClass('active');
    $(`.sidebar-tab[data-tab="${tab}"]`).addClass('active');

    if (tab === 'templates') {
        // 템플릿 탭 활성화
        $('#tabContentTemplates').addClass('active').show();
        $('#tabContentSkills').removeClass('active').hide();

        // 캔버스/툴바 영역 복원
        $('#canvasWrapper').show();
        $('#toolbar').show();
        $('#skillEditorWrapper').hide();

        // 템플릿 관련 헤더 버튼 표시
        $('#btnGridToggle').show();
        $('#btnBulkFont').css('display', state.currentTemplate ? '' : 'none');

        // 현재 템플릿이 선택되어 있으면 에디터 표시
        if (state.currentTemplate) {
            showEditor();
            $('#propertiesPanel').show();
            $('#templateNameBar').show();
            $('#btnDeleteTemplate').show();
            $('#btnPublishTemplate').show();
        } else {
            hideEditor();
        }
    } else if (tab === 'skills') {
        // 스킬 탭 활성화
        $('#tabContentTemplates').removeClass('active').hide();
        $('#tabContentSkills').addClass('active').show();

        // 캔버스/툴바/속성패널 숨기기
        $('#canvasWrapper').hide();
        $('#toolbar').hide();
        $('#propertiesPanel').hide();
        $('#templateNameBar').hide();
        $('#btnDeleteTemplate').hide();
        $('#btnPublishTemplate').hide();
        $('#btnBulkFont').hide();
        $('#btnGridToggle').hide();

        // 스킬 에디터 표시
        $('#skillEditorWrapper').show();

        // 스킬 목록 로드
        loadSkills();
    }
}

async function loadSkills() {
    try {
        const res = await apiGet('/api/admin/skills');
        state.skills = res.skills || [];
        renderSkillList();
    } catch (e) {
        showToast('스킬 로드 실패', 'error');
    }
}

function renderSkillList() {
    const list = $('#skillList');
    list.empty();

    if (state.skills.length === 0) {
        list.html('<div style="padding:12px;text-align:center;color:#999;font-size:12px;">등록된 스킬이 없습니다</div>');
        updateSkillComboText();
        return;
    }

    state.skills.forEach(s => {
        const active = state.currentSkill && state.currentSkill._id === s._id ? 'active' : '';
        const published = s.is_published;
        const statusBadge = published
            ? '<span class="skill-status-badge published" title="공개 중">공개</span>'
            : '<span class="skill-status-badge draft" title="비공개">비공개</span>';
        list.append(`
            <div class="template-option ${active}" onclick="selectSkill('${s._id}')">
                <div class="option-title-row">
                    <span class="option-title">${escapeHtml(s.title)}</span>
                    ${statusBadge}
                </div>
                <div class="option-meta">${s.description ? escapeHtml(s.description) : ''}</div>
            </div>
        `);
    });

    updateSkillComboText();
}

function updateSkillComboText() {
    if (state.currentSkill) {
        $('#skillComboText').text(state.currentSkill.title);
    } else {
        $('#skillComboText').text('스킬을 선택하세요');
    }
}

function toggleSkillDropdown() {
    const dd = $('#skillComboDropdown');
    dd.toggleClass('open');
}

// 스킬 콤보박스 외부 클릭 시 닫기
$(document).on('click', function(e) {
    if (!$(e.target).closest('#skillCombo').length) {
        $('#skillComboDropdown').removeClass('open');
    }
});

async function selectSkill(skillId) {
    $('#skillComboDropdown').removeClass('open');
    try {
        const res = await apiGet('/api/admin/skills/' + skillId);
        state.currentSkill = res.skill || res;

        renderSkillList();
        showSkillEditor(state.currentSkill);
    } catch (e) {
        showToast('스킬 로드 실패', 'error');
    }
}

function createNewSkill() {
    state.currentSkill = null;
    renderSkillList();
    updateSkillComboText();

    // 에디터 필드 초기화
    $('#skillTitle').val('');
    $('#skillDescription').val('');
    $('#skillPrompt').val('');
    $('#skillPageCount').val('auto');
    $('#skillTheme').val('corporate');

    // UI 상태
    $('#skillEditorTitle').text('새 스킬');
    $('#btnSkillPublish').hide();
    $('#btnSkillDelete').hide();

    // 에디터 표시
    $('#skillEditorEmpty').hide();
    $('#skillEditorForm').show();
}

function showSkillEditor(skill) {
    // 필드 채우기
    $('#skillTitle').val(skill.title || '');
    $('#skillDescription').val(skill.description || '');
    $('#skillPrompt').val(skill.skill_prompt || '');
    $('#skillPageCount').val(skill.page_count_default || 'auto');
    $('#skillTheme').val(skill.theme || 'corporate');

    // UI 상태
    $('#skillEditorTitle').text(escapeHtml(skill.title || '스킬 편집'));

    // 공개/비공개 버튼
    $('#btnSkillPublish').show();
    updateSkillPublishButton(skill.is_published);

    // 삭제 버튼
    $('#btnSkillDelete').show();

    // 에디터 표시
    $('#skillEditorEmpty').hide();
    $('#skillEditorForm').show();
}

function updateSkillPublishButton(isPublished) {
    const btn = $('#btnSkillPublish');
    btn.removeClass('btn-publish btn-unpublish');
    if (isPublished) {
        btn.addClass('btn-unpublish').text('비공개 전환');
    } else {
        btn.addClass('btn-publish').text('공개 전환');
    }
}

async function saveSkill() {
    const title = $('#skillTitle').val().trim();
    if (!title) {
        showToast('스킬 제목을 입력하세요', 'error');
        return;
    }

    const data = {
        title: title,
        description: $('#skillDescription').val().trim(),
        skill_prompt: $('#skillPrompt').val(),
        page_count_default: $('#skillPageCount').val(),
        theme: $('#skillTheme').val(),
    };

    try {
        if (state.currentSkill && state.currentSkill._id) {
            // 수정
            const res = await apiPut('/api/admin/skills/' + state.currentSkill._id, data);
            state.currentSkill = res.skill || res;
            showToast('스킬이 저장되었습니다', 'success');
        } else {
            // 새 스킬 생성
            data.is_published = false;
            const res = await apiPost('/api/admin/skills', data);
            state.currentSkill = res.skill || res;
            showToast('새 스킬이 생성되었습니다', 'success');
        }

        // 목록 갱신
        await loadSkills();

        // 에디터 업데이트
        showSkillEditor(state.currentSkill);
    } catch (e) {
        showToast('스킬 저장 실패: ' + (e.message || ''), 'error');
    }
}

async function deleteSkill() {
    if (!state.currentSkill || !state.currentSkill._id) return;
    if (!confirm(`"${state.currentSkill.title}" 스킬을 삭제하시겠습니까?`)) return;

    try {
        await apiDelete('/api/admin/skills/' + state.currentSkill._id);
        state.currentSkill = null;
        showToast('스킬이 삭제되었습니다', 'success');

        // 에디터 초기화
        $('#skillEditorForm').hide();
        $('#skillEditorEmpty').show();

        // 목록 갱신
        await loadSkills();
    } catch (e) {
        showToast('스킬 삭제 실패', 'error');
    }
}

async function toggleSkillPublish() {
    if (!state.currentSkill || !state.currentSkill._id) return;

    try {
        const res = await apiPut('/api/admin/skills/' + state.currentSkill._id + '/publish', {});
        state.currentSkill.is_published = res.is_published;

        updateSkillPublishButton(res.is_published);
        showToast(res.is_published ? '스킬이 공개되었습니다' : '스킬이 비공개로 전환되었습니다', 'success');

        // 목록 갱신
        await loadSkills();
    } catch (e) {
        showToast('공개 상태 변경 실패', 'error');
    }
}
