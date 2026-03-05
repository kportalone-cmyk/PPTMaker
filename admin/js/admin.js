/**
 * PPTMaker 관리자 모듈
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
    drawStart: null,
    // 격자/눈금자
    showGrid: true,
    gridSize: 50,  // 50px 간격
    // 레이아웃 프리셋 변경 감지
    _previousLayout: '',
};

// ============ 초기화 ============
$(document).ready(function () {
    // URL에서 JWT 추출
    const pathParts = window.location.pathname.split('/');
    // /admin 경로에서 JWT를 URL에서 가져옴
    const urlParams = new URLSearchParams(window.location.search);
    state.jwtToken = urlParams.get('token') || pathParts[1] || '';

    if (!state.jwtToken || state.jwtToken === 'admin') {
        // JWT가 없으면 로컬스토리지에서 확인 (임시 - 실제로는 URL에서 가져와야 함)
        state.jwtToken = localStorage.getItem('pptmaker_jwt') || '';
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
            const link = document.createElement('link');
            link.id = linkId;
            link.rel = 'stylesheet';
            link.href = font.url;
            document.head.appendChild(link);
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
        list.append(`
            <div class="template-option ${active}" onclick="selectTemplate('${t._id}')">
                <div class="option-title">${escapeHtml(t.name)}</div>
                <div class="option-meta">${escapeHtml(t.template_type || '')} ${t.description ? '· ' + escapeHtml(t.description) : ''}</div>
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
    $('#newTemplateModal').show();
}

function closeModal(id) {
    $('#' + id).hide();
}

async function createTemplate() {
    const name = $('#newTemplateName').val().trim();
    if (!name) {
        alert('템플릿 이름을 입력하세요');
        return;
    }
    try {
        const res = await apiPost('/api/admin/templates', {
            name: name,
            description: $('#newTemplateDesc').val().trim(),
        });
        closeModal('newTemplateModal');
        showToast('템플릿이 생성되었습니다', 'success');
        await loadTemplates();
        selectTemplate(res.template._id);
    } catch (e) {
        showToast('생성 실패: ' + e.message, 'error');
    }
}

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

        renderTemplateList();
        renderSlideList();
        showEditor();

        $('#slideSection').show();
        $('#btnDeleteTemplate').show();

        // 캔버스 초기화
        clearCanvas();
        updateBgUI();

        if (state.slides.length > 0) {
            selectSlide(state.slides[0]._id);
        }
    } catch (e) {
        showToast('템플릿 로드 실패', 'error');
    }
}

function updateBgUI() {
    if (state.currentTemplate && state.currentTemplate.background_image) {
        $('#canvasBg').css('background-image', `url(${state.currentTemplate.background_image})`);
        $('#btnRemoveBg').show();
    } else {
        $('#canvasBg').css('background-image', 'none');
        $('#btnRemoveBg').hide();
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

// ============ 슬라이드 관리 ============
function renderSlideList() {
    const list = $('#slideList');
    list.empty();

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
                <div class="slide-mini-preview" id="miniPreview_${s._id}"></div>
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

    // 배경 이미지
    if (state.currentTemplate && state.currentTemplate.background_image) {
        container.style.backgroundImage = `url(${state.currentTemplate.background_image})`;
        container.style.backgroundSize = 'cover';
        container.style.backgroundPosition = 'center';
    } else {
        container.style.backgroundImage = 'none';
        container.style.background = '#f0f0f0';
    }

    const objects = slide.objects || [];
    if (objects.length === 0) return;

    // 썸네일 크기 기준 스케일 계산 (캔버스 960x540 → 썸네일 크기)
    const thumbW = container.clientWidth || 160;
    const thumbH = container.clientHeight || 90;
    const scaleX = thumbW / 960;
    const scaleY = thumbH / 540;

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
            img.style.objectFit = 'contain';
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
    $('#btnRemoveBg').hide();
}

function scaleCanvas() {
    const wrapper = document.getElementById('canvasWrapper');
    const canvas = document.getElementById('canvas');
    if (!wrapper || !canvas || canvas.style.display === 'none') return;

    const wrapperW = wrapper.clientWidth - 40;
    const wrapperH = wrapper.clientHeight - 40;
    const scaleX = wrapperW / 960;
    const scaleY = wrapperH / 540;
    const scale = Math.min(scaleX, scaleY, 1.2);

    canvas.style.transform = `scale(${scale})`;
}

$(window).on('resize', function () {
    scaleCanvas();
});

// ============ 격자 ============
function toggleGrid() {
    state.showGrid = !state.showGrid;
    updateGridVisibility();
}

function updateGridVisibility() {
    const grid = $('#canvasGrid');
    const btn = $('#btnGridToggle');

    if (state.showGrid) {
        grid.addClass('visible');
        btn.addClass('grid-active');
    } else {
        grid.removeClass('visible');
        btn.removeClass('grid-active');
    }
}

function showPropertiesPanel() {
    $('#propertiesPanel').show();
}

function clientToCanvasCoords(clientX, clientY) {
    const canvas = document.getElementById('canvas');
    const rect = canvas.getBoundingClientRect();
    const scaleX = rect.width / 960;
    const scaleY = rect.height / 540;
    return {
        x: Math.round(Math.max(0, Math.min((clientX - rect.left) / scaleX, 960))),
        y: Math.round(Math.max(0, Math.min((clientY - rect.top) / scaleY, 540))),
    };
}

function clearCanvas() {
    $('#canvas .canvas-object').remove();
}

function renderCanvas() {
    clearCanvas();
    const canvas = $('#canvas');

    // 배경 이미지
    if (state.currentTemplate && state.currentTemplate.background_image) {
        $('#canvasBg').css('background-image', `url(${state.currentTemplate.background_image})`);
    }

    state.objects.forEach(obj => {
        const el = createObjectElement(obj);
        canvas.append(el);
    });
}

function createObjectElement(obj) {
    const typeClass = obj.obj_type === 'image' ? 'obj-image' :
                      obj.obj_type === 'shape' ? 'obj-shape' : 'obj-text';
    const div = $('<div>')
        .addClass('canvas-object')
        .addClass(typeClass)
        .attr('data-obj-id', obj.obj_id)
        .css({
            left: obj.x + 'px',
            top: obj.y + 'px',
            width: obj.width + 'px',
            height: obj.height + 'px',
        });

    if (obj.obj_type === 'image' && obj.image_url) {
        div.append(`<img src="${obj.image_url}" alt="image">`);
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
            selectObject(obj.obj_id);
            if (obj.obj_type !== 'text' || !$(e.target).hasClass('text-content')) {
                startDrag(e, obj.obj_id);
            }
        }
    });

    return div;
}

function generateObjId() {
    return 'obj_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
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
            image_url: res.image_url,
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

function deleteObject(objId) {
    state.objects = state.objects.filter(o => o.obj_id !== objId);
    $(`[data-obj-id="${objId}"]`).remove();
    state.selectedObject = null;
    $('#objProperties').hide();
    $('#textProperties').hide();
    $('#shapeProperties').hide();
}

function selectObject(objId) {
    state.selectedObject = state.objects.find(o => o.obj_id === objId) || null;

    // UI 업데이트
    $('.canvas-object').removeClass('selected');
    $(`[data-obj-id="${objId}"]`).addClass('selected');

    if (state.selectedObject) {
        updatePropertiesPanel();
    }
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

    if (state.isDragging && state.selectedObject) {
        const canvasRect = document.getElementById('canvas').getBoundingClientRect();
        let x = e.clientX - canvasRect.left - state.dragOffset.x;
        let y = e.clientY - canvasRect.top - state.dragOffset.y;

        // 캔버스 경계 제한
        x = Math.max(0, Math.min(x, 960 - state.selectedObject.width));
        y = Math.max(0, Math.min(y, 540 - state.selectedObject.height));

        state.selectedObject.x = Math.round(x);
        state.selectedObject.y = Math.round(y);

        const el = $(`[data-obj-id="${state.selectedObject.obj_id}"]`);
        el.css({ left: x + 'px', top: y + 'px' });

        updatePositionInputs();
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
            left = Math.max(0, Math.min(left - width / 2, 960 - width));
            top = Math.max(0, Math.min(top - height / 2, 540 - height));
        }

        // 캔버스 경계 제한
        left = Math.max(0, left);
        top = Math.max(0, top);
        width = Math.min(width, 960 - left);
        height = Math.min(height, 540 - top);

        createShapeAtRect(state.drawShapeType, left, top, width, height);
        cancelDrawMode();
        return;
    }

    state.isDragging = false;
    state.isResizing = false;
});

// ESC 키로 드로잉 모드 취소
$(document).on('keydown', function (e) {
    if (e.key === 'Escape' && state.isDrawing) {
        cancelDrawMode();
        e.preventDefault();
    }
});

// 캔버스 외부 클릭 시 드로잉 모드 취소
$(document).on('mousedown', function (e) {
    if (state.isDrawing && !$(e.target).closest('#canvas').length) {
        cancelDrawMode();
    }
});

// 캔버스 클릭 (선택 해제 또는 도형 드로잉 시작)
$('#canvas').on('mousedown', function (e) {
    if (e.target !== this && e.target.id !== 'canvasBg') return;

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
        // 일반 모드: 선택 해제
        $('.canvas-object').removeClass('selected');
        state.selectedObject = null;
        $('#objProperties').hide();
        $('#textProperties').hide();
        $('#shapeProperties').hide();
    }
});

// ============ 속성 패널 ============
function updatePropertiesPanel() {
    const obj = state.selectedObject;
    if (!obj) return;

    $('#objProperties').show();
    $('#propRole').val(obj.role || '');
    $('#propPlaceholder').val(obj.placeholder || '');
    $('#propX').val(Math.round(obj.x));
    $('#propY').val(Math.round(obj.y));
    $('#propW').val(Math.round(obj.width));
    $('#propH').val(Math.round(obj.height));

    if (obj.obj_type === 'text') {
        $('#textProperties').show();
        $('#shapeProperties').hide();
        const style = obj.text_style || {};
        $('#propFont').val(style.font_family || 'Arial');
        $('#propFontSize').val(style.font_size || 16);
        $('#propColor').val(style.color || '#000000');
        $('#propBold').toggleClass('active', !!style.bold);
        $('#propItalic').toggleClass('active', !!style.italic);
        $('#propAlign').val(style.align || 'left');
        $('#propTextContent').val(obj.text_content || '');
    } else if (obj.obj_type === 'shape') {
        $('#textProperties').hide();
        $('#shapeProperties').show();
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
    } else {
        $('#textProperties').hide();
        $('#shapeProperties').hide();
    }
}

function updatePositionInputs() {
    if (!state.selectedObject) return;
    $('#propX').val(Math.round(state.selectedObject.x));
    $('#propY').val(Math.round(state.selectedObject.y));
    $('#propW').val(Math.round(state.selectedObject.width));
    $('#propH').val(Math.round(state.selectedObject.height));
}

function updateObjProperty(prop, value) {
    if (!state.selectedObject) return;
    state.selectedObject[prop] = value;
}

function updateObjPosition() {
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
}

function updateTextStyle(prop, value) {
    if (!state.selectedObject || state.selectedObject.obj_type !== 'text') return;
    if (!state.selectedObject.text_style) {
        state.selectedObject.text_style = {};
    }
    state.selectedObject.text_style[prop] = value;

    // 캔버스 오브젝트에 스타일 반영
    const textEl = $(`[data-obj-id="${state.selectedObject.obj_id}"] .text-content`);
    switch (prop) {
        case 'font_family': textEl.css('fontFamily', value); break;
        case 'font_size': textEl.css('fontSize', value + 'px'); break;
        case 'color': textEl.css('color', value); break;
        case 'bold': textEl.css('fontWeight', value ? 'bold' : 'normal'); break;
        case 'italic': textEl.css('fontStyle', value ? 'italic' : 'normal'); break;
        case 'align': textEl.css('textAlign', value); break;
    }
}

function toggleBold() {
    const btn = $('#propBold');
    const newVal = !btn.hasClass('active');
    btn.toggleClass('active', newVal);
    updateTextStyle('bold', newVal);
}

function toggleItalic() {
    const btn = $('#propItalic');
    const newVal = !btn.hasClass('active');
    btn.toggleClass('active', newVal);
    updateTextStyle('italic', newVal);
}

function updateTextContent(value) {
    if (!state.selectedObject || state.selectedObject.obj_type !== 'text') return;
    state.selectedObject.text_content = value;
    $(`[data-obj-id="${state.selectedObject.obj_id}"] .text-content`).text(value);
}

// ============ 도형 기능 ============

function cancelDrawMode() {
    state.isDrawing = false;
    state.drawShapeType = null;
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
    dd.toggle();
    if (dd.is(':visible')) {
        setTimeout(() => {
            $(document).one('click', function (e) {
                if (!$(e.target).closest('.shape-dropdown-wrapper').length) {
                    dd.hide();
                }
            });
        }, 0);
    }
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

    let inner = '';

    switch (s.shape_type) {
        case 'rounded_rectangle': {
            const rx = s.border_radius || 12;
            inner = `<rect x="${half}" y="${half}" width="${Math.max(0, w - strokeW)}" height="${Math.max(0, h - strokeW)}"
                      rx="${rx}" ry="${rx}"
                      fill="${fill}" fill-opacity="${fillOpacity}"
                      stroke="${stroke}" stroke-width="${strokeW}" ${dashAttr}/>`;
            break;
        }
        case 'ellipse':
            inner = `<ellipse cx="${w / 2}" cy="${h / 2}" rx="${Math.max(0, w / 2 - half)}" ry="${Math.max(0, h / 2 - half)}"
                      fill="${fill}" fill-opacity="${fillOpacity}"
                      stroke="${stroke}" stroke-width="${strokeW}" ${dashAttr}/>`;
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
        default: // rectangle
            inner = `<rect x="${half}" y="${half}" width="${Math.max(0, w - strokeW)}" height="${Math.max(0, h - strokeW)}"
                      fill="${fill}" fill-opacity="${fillOpacity}"
                      stroke="${stroke}" stroke-width="${strokeW}" ${dashAttr}/>`;
            break;
    }

    return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" style="overflow:visible;">${inner}</svg>`;
}

function updateShapeStyle(prop, value) {
    if (!state.selectedObject || state.selectedObject.obj_type !== 'shape') return;
    if (!state.selectedObject.shape_style) {
        state.selectedObject.shape_style = {};
    }
    state.selectedObject.shape_style[prop] = value;

    // 도형 유형 변경 시 라인 관련 기본값 조정
    if (prop === 'shape_type') {
        const isLine = (value === 'line' || value === 'arrow');
        if (isLine) {
            state.selectedObject.shape_style.fill_opacity = 0;
            state.selectedObject.shape_style.fill_color = 'transparent';
            if (state.selectedObject.height > 20) {
                state.selectedObject.height = 4;
                const el = $(`[data-obj-id="${state.selectedObject.obj_id}"]`);
                el.css('height', '4px');
            }
        }
        if (value === 'arrow') {
            state.selectedObject.shape_style.arrow_head = state.selectedObject.shape_style.arrow_head || 'end';
        }
        if (value === 'rounded_rectangle') {
            state.selectedObject.shape_style.border_radius = state.selectedObject.shape_style.border_radius || 12;
        }
    }

    // SVG 재렌더
    const el = $(`[data-obj-id="${state.selectedObject.obj_id}"]`);
    el.find('svg').remove();
    el.prepend(createShapeSVG(state.selectedObject));

    updateShapePropertyVisibility();
}

function updateShapePropertyVisibility() {
    const s = state.selectedObject?.shape_style || {};
    const isLine = s.shape_type === 'line' || s.shape_type === 'arrow';
    $('#shapeFillRow').toggle(!isLine);
    $('#borderRadiusGroup').toggle(s.shape_type === 'rounded_rectangle');
    $('#arrowHeadGroup').toggle(s.shape_type === 'arrow');
}

// ============ 슬라이드 메타 ============
function updateSlideMetaUI() {
    $('#metaContentType').val(state.slideMeta.content_type || 'body');
    $('#metaLayout').val(state.slideMeta.layout || '');
    $('#metaHasTitle').prop('checked', state.slideMeta.has_title || false);
    $('#metaHasGovernance').prop('checked', state.slideMeta.has_governance || false);
    $('#metaDescCount').val(state.slideMeta.description_count || 0);
    _showMetaTypeGuide(state.slideMeta.content_type || 'body');
}

function updateSlideMeta() {
    const newLayout = $('#metaLayout').val();
    const oldLayout = state._previousLayout || '';

    state.slideMeta = {
        content_type: $('#metaContentType').val(),
        layout: newLayout,
        has_title: $('#metaHasTitle').is(':checked'),
        has_governance: $('#metaHasGovernance').is(':checked'),
        description_count: parseInt($('#metaDescCount').val()) || 0,
    };
    _showMetaTypeGuide(state.slideMeta.content_type);

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

    // 슬라이드 메타 자동 업데이트
    state.slideMeta.content_type = presetData.content_type;
    state.slideMeta.has_title = presetData.has_title;
    state.slideMeta.has_governance = presetData.has_governance;
    state.slideMeta.description_count = presetData.description_count;
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
    };
    const guide = guides[contentType];
    if (guide) {
        $('#metaTypeGuide').html(guide).show();
    } else {
        $('#metaTypeGuide').hide();
    }
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
