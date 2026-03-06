/**
 * PPTMaker - Kimi K2 Slides Style Frontend
 */

// ============ 전역 상태 ============
const state = {
    jwtToken: '',
    userInfo: null,
    isAdmin: false,
    lang: 'ko',
    projects: [],
    currentProject: null,
    resources: [],
    generatedSlides: [],
    currentSlideIndex: 0,
    templates: [],
    selectedTemplateId: null,
    searchResults: [],
    searchHighlightIndex: -1,
    lastWebSearchResult: null,
    sidebarCollapsed: false,
    editMode: false,
    editSelectedObj: null,
    editDragging: false,
    editResizing: false,
    editDragOffset: { x: 0, y: 0 },
    editResizeDir: '',
    editResizeStart: null,
    editDirtySlides: new Set(),
};

let _animationCancelled = false;
let _isAnimating = false;

// 탭 전환 시 애니메이션 즉시 완료 처리
document.addEventListener('visibilitychange', () => {
    if (!document.hidden && _isAnimating && !_animationCancelled) {
        _animationCancelled = true;
    }
});

// ============ 다국어 사전 ============
const I18N = {
    ko: {
        appTitle: 'PPTMaker',
        appSubtitle: '기업용 파워포인트 자동 생성 솔루션',
        login: '로그인',
        logout: '로그아웃',
        userSearch: '사용자 검색',
        userSearchPlaceholder: '이름을 입력하세요',
        password: '비밀번호',
        passwordPlaceholder: '비밀번호를 입력하세요',
        noSearchResult: '검색 결과가 없습니다',
        myProjects: '내 프로젝트',
        newProject: '새 프로젝트 생성',
        newProjectTitle: '새 프로젝트',
        projectName: '프로젝트 이름',
        projectNamePlaceholder: '프로젝트 이름을 입력하세요',
        description: '설명',
        descriptionPlaceholder: '프로젝트 설명을 입력하세요',
        create: '생성',
        cancel: '취소',
        save: '저장',
        edit: '수정',
        delete: '삭제',
        uploadFile: '파일',
        addText: '텍스트',
        webSearch: '검색',
        noResources: '리소스가 없습니다',
        selectTemplate: '템플릿 선택',
        selectTemplatePlaceholder: '템플릿을 선택하세요',
        outputLang: '출력 언어',
        defaultLang: '기본 언어',
        instructions: '사용자 지침',
        instructionsPlaceholder: '프레젠테이션에 대한 지침을 입력하세요...',
        slideCountAuto: '자동',
        generateBtn: '생성',
        presentationView: '프레젠테이션',
        downloadPptx: 'PPTX',
        downloadPdf: 'PDF',
        copyShareLink: '공유',
        editSlide: '편집',
        editSave: '저장',
        editReplaceImage: '이미지 교체',
        editUnsavedChanges: '저장하지 않은 변경사항이 있습니다. 저장하시겠습니까?',
        editSaved: '변경사항이 저장되었습니다',
        noSlides: '아래에서 지침을 입력하고 생성 버튼을 눌러주세요',
        prev: '이전',
        next: '다음',
        textResourceTitle: '텍스트 리소스 추가',
        textResTitle: '제목',
        textResTitlePlaceholder: '리소스 제목',
        textResContent: '내용',
        textResContentPlaceholder: '텍스트를 붙여넣으세요...',
        add: '추가',
        webSearchTitle: '웹 검색',
        searchQuery: '검색어',
        searchQueryPlaceholder: '검색할 내용을 입력하세요',
        searchBtn: '검색',
        searchResult: '검색 결과',
        close: '닫기',
        editProject: '프로젝트 수정',
        sharedPresentation: '공유 프레젠테이션',
        statusDraft: '초안',
        statusPreparing: '준비 중',
        statusGenerated: '생성 완료',
        noDesc: '설명 없음',
        typeFile: '파일',
        typeText: '텍스트',
        typeWeb: '웹 검색',
        welcomeTitle: '무엇을 만들어볼까요?',
        welcomeDesc: '프로젝트를 선택하거나 새로 만들어 AI 프레젠테이션을 시작하세요',
        msgSelectUser: '사용자를 선택하세요',
        msgEnterPassword: '비밀번호를 입력하세요',
        msgSelectTemplate: '템플릿을 선택하세요',
        msgAddResources: '먼저 리소스를 추가하세요',
        msgEnterContent: '내용을 입력하세요',
        msgEnterQuery: '검색어를 입력하세요',
        msgEnterProjectName: '프로젝트 이름을 입력하세요',
        msgGenerating: '슬라이드 생성 중...',
        msgUploading: '파일 업로드 중...',
        msgLoadingProject: '프로젝트 로드 중...',
        msgProcessing: '처리 중...',
        msgSlidesGenerated: '개 슬라이드가 생성되었습니다',
        msgAllComplete: '모든 슬라이드 생성이 완료되었습니다!',
        msgShareCopied: '공유 링크가 복사되었습니다',
        msgProjectCreated: '프로젝트가 생성되었습니다',
        msgProjectDeleted: '프로젝트가 삭제되었습니다',
        msgUpdated: '수정되었습니다',
        msgDeleted: '삭제되었습니다',
        msgTextAdded: '텍스트가 추가되었습니다',
        msgFileUploaded: '파일이 업로드되었습니다',
        msgSearchAdded: '검색 결과가 리소스로 추가되었습니다',
        msgConfirmDelete: '이 프로젝트를 삭제하시겠습니까?',
        msgConfirmDeleteRes: '이 리소스를 삭제하시겠습니까?',
        searching: '검색 중...',
        adminPage: '관리자 페이지',
        authChecking: '접속정보 확인중입니다',
        authCheckingSub: '잠시만 기다려주세요...',
        authErrorTitle: '인증정보가 올바르지 않습니다',
        authErrorDesc: '접속 링크가 만료되었거나 잘못된 인증정보입니다.\n관리자에게 문의하거나 다시 시도해주세요.',
    },
    en: {
        appTitle: 'PPTMaker',
        appSubtitle: 'Enterprise PowerPoint Auto-Generation',
        login: 'Login',
        logout: 'Logout',
        userSearch: 'User Search',
        userSearchPlaceholder: 'Enter name',
        password: 'Password',
        passwordPlaceholder: 'Enter password',
        noSearchResult: 'No results found',
        myProjects: 'My Projects',
        newProject: 'New Project',
        newProjectTitle: 'New Project',
        projectName: 'Project Name',
        projectNamePlaceholder: 'Enter project name',
        description: 'Description',
        descriptionPlaceholder: 'Enter description',
        create: 'Create',
        cancel: 'Cancel',
        save: 'Save',
        edit: 'Edit',
        delete: 'Delete',
        uploadFile: 'File',
        addText: 'Text',
        webSearch: 'Search',
        noResources: 'No resources',
        selectTemplate: 'Select Template',
        selectTemplatePlaceholder: 'Choose a template',
        outputLang: 'Output Language',
        defaultLang: 'Default',
        instructions: 'Instructions',
        instructionsPlaceholder: 'Enter instructions for your presentation...',
        slideCountAuto: 'Auto',
        generateBtn: 'Generate',
        presentationView: 'Present',
        downloadPptx: 'PPTX',
        downloadPdf: 'PDF',
        copyShareLink: 'Share',
        editSlide: 'Edit',
        editSave: 'Save',
        editReplaceImage: 'Replace',
        editUnsavedChanges: 'You have unsaved changes. Save now?',
        editSaved: 'Changes saved',
        noSlides: 'Enter instructions below and click Generate',
        prev: 'Prev',
        next: 'Next',
        textResourceTitle: 'Add Text Resource',
        textResTitle: 'Title',
        textResTitlePlaceholder: 'Resource title',
        textResContent: 'Content',
        textResContentPlaceholder: 'Paste your text here...',
        add: 'Add',
        webSearchTitle: 'Web Search',
        searchQuery: 'Query',
        searchQueryPlaceholder: 'Enter search query',
        searchBtn: 'Search',
        searchResult: 'Results',
        close: 'Close',
        editProject: 'Edit Project',
        sharedPresentation: 'Shared Presentation',
        statusDraft: 'Draft',
        statusPreparing: 'Preparing',
        statusGenerated: 'Generated',
        noDesc: 'No description',
        typeFile: 'File',
        typeText: 'Text',
        typeWeb: 'Web',
        welcomeTitle: 'What would you like to create?',
        welcomeDesc: 'Select a project or create a new one to start your AI presentation',
        msgSelectUser: 'Please select a user',
        msgEnterPassword: 'Please enter password',
        msgSelectTemplate: 'Please select a template',
        msgAddResources: 'Please add resources first',
        msgEnterContent: 'Please enter content',
        msgEnterQuery: 'Please enter a search query',
        msgEnterProjectName: 'Please enter project name',
        msgGenerating: 'Generating slides...',
        msgUploading: 'Uploading...',
        msgLoadingProject: 'Loading project...',
        msgProcessing: 'Processing...',
        msgSlidesGenerated: ' slides generated',
        msgAllComplete: 'All slides generated!',
        msgShareCopied: 'Share link copied',
        msgProjectCreated: 'Project created',
        msgProjectDeleted: 'Project deleted',
        msgUpdated: 'Updated',
        msgDeleted: 'Deleted',
        msgTextAdded: 'Text added',
        msgFileUploaded: 'File uploaded',
        msgSearchAdded: 'Search result added',
        msgConfirmDelete: 'Delete this project?',
        msgConfirmDeleteRes: 'Delete this resource?',
        searching: 'Searching...',
        adminPage: 'Admin',
        authChecking: 'Verifying access...',
        authCheckingSub: 'Please wait a moment...',
        authErrorTitle: 'Invalid authentication',
        authErrorDesc: 'The access link has expired or is invalid.\nPlease contact your administrator or try again.',
    },
    ja: {
        appTitle: 'PPTMaker',
        appSubtitle: '企業向けPPT自動生成',
        login: 'ログイン',
        logout: 'ログアウト',
        userSearch: 'ユーザー検索',
        userSearchPlaceholder: '名前を入力',
        password: 'パスワード',
        passwordPlaceholder: 'パスワードを入力',
        noSearchResult: '結果なし',
        myProjects: 'プロジェクト',
        newProject: '新規作成',
        newProjectTitle: '新規プロジェクト',
        projectName: 'プロジェクト名',
        projectNamePlaceholder: 'プロジェクト名を入力',
        description: '説明',
        descriptionPlaceholder: '説明を入力',
        create: '作成',
        cancel: 'キャンセル',
        save: '保存',
        edit: '編集',
        delete: '削除',
        uploadFile: 'ファイル',
        addText: 'テキスト',
        webSearch: '検索',
        noResources: 'リソースなし',
        selectTemplate: 'テンプレート選択',
        selectTemplatePlaceholder: 'テンプレートを選択',
        outputLang: '出力言語',
        defaultLang: 'デフォルト',
        instructions: '指示',
        instructionsPlaceholder: 'プレゼンテーションの指示を入力...',
        slideCountAuto: '自動',
        generateBtn: '生成',
        presentationView: 'プレゼン',
        downloadPptx: 'PPTX',
        downloadPdf: 'PDF',
        copyShareLink: '共有',
        editSlide: '編集',
        editSave: '保存',
        editReplaceImage: '画像変更',
        editUnsavedChanges: '保存されていない変更があります。保存しますか？',
        editSaved: '変更が保存されました',
        noSlides: '指示を入力し生成ボタンを押してください',
        prev: '前',
        next: '次',
        textResourceTitle: 'テキスト追加',
        textResTitle: 'タイトル',
        textResTitlePlaceholder: 'タイトル',
        textResContent: '内容',
        textResContentPlaceholder: 'テキストを貼り付け...',
        add: '追加',
        webSearchTitle: 'ウェブ検索',
        searchQuery: '検索ワード',
        searchQueryPlaceholder: '検索内容を入力',
        searchBtn: '検索',
        searchResult: '検索結果',
        close: '閉じる',
        editProject: 'プロジェクト編集',
        sharedPresentation: '共有プレゼン',
        statusDraft: '下書き',
        statusPreparing: '準備中',
        statusGenerated: '完了',
        noDesc: '説明なし',
        typeFile: 'ファイル',
        typeText: 'テキスト',
        typeWeb: 'ウェブ',
        welcomeTitle: '何を作りますか？',
        welcomeDesc: 'プロジェクトを選択または新規作成してAIプレゼンを開始',
        msgSelectUser: 'ユーザーを選択してください',
        msgEnterPassword: 'パスワードを入力してください',
        msgSelectTemplate: 'テンプレートを選択してください',
        msgAddResources: 'リソースを追加してください',
        msgEnterContent: '内容を入力してください',
        msgEnterQuery: '検索ワードを入力してください',
        msgEnterProjectName: 'プロジェクト名を入力してください',
        msgGenerating: 'スライド生成中...',
        msgUploading: 'アップロード中...',
        msgLoadingProject: '読み込み中...',
        msgProcessing: '処理中...',
        msgSlidesGenerated: '個のスライドが生成されました',
        msgAllComplete: '全スライド生成完了！',
        msgShareCopied: '共有リンクをコピーしました',
        msgProjectCreated: '作成しました',
        msgProjectDeleted: '削除しました',
        msgUpdated: '更新しました',
        msgDeleted: '削除しました',
        msgTextAdded: 'テキストを追加しました',
        msgFileUploaded: 'アップロードしました',
        msgSearchAdded: '検索結果を追加しました',
        msgConfirmDelete: 'このプロジェクトを削除しますか？',
        msgConfirmDeleteRes: 'このリソースを削除しますか？',
        searching: '検索中...',
        adminPage: '管理者ページ',
        authChecking: 'アクセス情報を確認中です',
        authCheckingSub: 'しばらくお待ちください...',
        authErrorTitle: '認証情報が正しくありません',
        authErrorDesc: 'アクセスリンクの有効期限が切れているか、無効です。\n管理者にお問い合わせください。',
    },
    zh: {
        appTitle: 'PPTMaker',
        appSubtitle: '企业级PPT自动生成',
        login: '登录',
        logout: '退出',
        userSearch: '搜索用户',
        userSearchPlaceholder: '输入姓名',
        password: '密码',
        passwordPlaceholder: '输入密码',
        noSearchResult: '未找到结果',
        myProjects: '我的项目',
        newProject: '新建项目',
        newProjectTitle: '新项目',
        projectName: '项目名称',
        projectNamePlaceholder: '输入项目名称',
        description: '描述',
        descriptionPlaceholder: '输入描述',
        create: '创建',
        cancel: '取消',
        save: '保存',
        edit: '编辑',
        delete: '删除',
        uploadFile: '文件',
        addText: '文本',
        webSearch: '搜索',
        noResources: '无资源',
        selectTemplate: '选择模板',
        selectTemplatePlaceholder: '选择模板',
        outputLang: '输出语言',
        defaultLang: '默认',
        instructions: '指令',
        instructionsPlaceholder: '输入演示文稿的指令...',
        slideCountAuto: '自动',
        generateBtn: '生成',
        presentationView: '演示',
        downloadPptx: 'PPTX',
        downloadPdf: 'PDF',
        copyShareLink: '分享',
        editSlide: '编辑',
        editSave: '保存',
        editReplaceImage: '替换图片',
        editUnsavedChanges: '有未保存的更改。是否保存？',
        editSaved: '更改已保存',
        noSlides: '请在下方输入指令并点击生成',
        prev: '上一页',
        next: '下一页',
        textResourceTitle: '添加文本',
        textResTitle: '标题',
        textResTitlePlaceholder: '标题',
        textResContent: '内容',
        textResContentPlaceholder: '粘贴文本...',
        add: '添加',
        webSearchTitle: '网页搜索',
        searchQuery: '搜索词',
        searchQueryPlaceholder: '输入搜索内容',
        searchBtn: '搜索',
        searchResult: '搜索结果',
        close: '关闭',
        editProject: '编辑项目',
        sharedPresentation: '共享演示',
        statusDraft: '草稿',
        statusPreparing: '准备中',
        statusGenerated: '已完成',
        noDesc: '无描述',
        typeFile: '文件',
        typeText: '文本',
        typeWeb: '网页',
        welcomeTitle: '想要创建什么？',
        welcomeDesc: '选择项目或新建以开始AI演示',
        msgSelectUser: '请选择用户',
        msgEnterPassword: '请输入密码',
        msgSelectTemplate: '请选择模板',
        msgAddResources: '请先添加资源',
        msgEnterContent: '请输入内容',
        msgEnterQuery: '请输入搜索词',
        msgEnterProjectName: '请输入项目名称',
        msgGenerating: '正在生成...',
        msgUploading: '上传中...',
        msgLoadingProject: '加载中...',
        msgProcessing: '处理中...',
        msgSlidesGenerated: '个幻灯片已生成',
        msgAllComplete: '全部生成完成！',
        msgShareCopied: '链接已复制',
        msgProjectCreated: '已创建',
        msgProjectDeleted: '已删除',
        msgUpdated: '已更新',
        msgDeleted: '已删除',
        msgTextAdded: '文本已添加',
        msgFileUploaded: '已上传',
        msgSearchAdded: '搜索结果已添加',
        msgConfirmDelete: '确定删除此项目？',
        msgConfirmDeleteRes: '确定删除此资源？',
        searching: '搜索中...',
        adminPage: '管理后台',
        authChecking: '正在验证访问信息',
        authCheckingSub: '请稍候...',
        authErrorTitle: '认证信息无效',
        authErrorDesc: '访问链接已过期或无效。\n请联系管理员或重试。',
    },
};

/** 번역 */
function t(key) {
    const dict = I18N[state.lang] || I18N['ko'];
    return dict[key] || I18N['ko'][key] || key;
}

/** URL에서 언어 코드 파싱 */
function parseLangFromUrl() {
    const pathParts = window.location.pathname.split('/').filter(Boolean);
    const last = pathParts[pathParts.length - 1];
    if (last && last.length >= 2 && last.length <= 5 && /^[a-z]{2,5}$/i.test(last)) {
        const langCode = last.toLowerCase();
        if (I18N[langCode]) return langCode;
    }
    return '';
}

/** UI 전체 다국어 적용 */
function applyI18n() {
    // 로그인
    $('.login-logo h1').text(t('appTitle'));
    $('.login-subtitle').text(t('appSubtitle'));
    $('label[for="loginUserSearch"]').text(t('userSearch'));
    $('#loginUserSearch').attr('placeholder', t('userSearchPlaceholder'));
    $('label[for="loginPassword"]').text(t('password'));
    $('#loginPassword').attr('placeholder', t('passwordPlaceholder'));
    $('.btn-login').text(t('login'));

    // 인증 화면
    $('.i18n-authChecking').text(t('authChecking'));
    $('.i18n-authCheckingSub').text(t('authCheckingSub'));
    $('.i18n-authErrorTitle').text(t('authErrorTitle'));
    $('.i18n-authErrorDesc').html(t('authErrorDesc').replace(/\n/g, '<br>'));

    // 사이드바
    $('.i18n-newProject').text(t('newProject'));
    $('.i18n-myProjects').text(t('myProjects'));

    // 관리자
    $('.i18n-adminPage').text(t('adminPage'));

    // 빈 상태
    $('.i18n-welcomeTitle').text(t('welcomeTitle'));
    $('.i18n-welcomeDesc').text(t('welcomeDesc'));

    // 워크스페이스
    $('.i18n-uploadFile').text(t('uploadFile'));
    $('.i18n-addText').text(t('addText'));
    $('.i18n-webSearch').text(t('webSearch'));
    $('.i18n-noSlides').text(t('noSlides'));
    $('.i18n-presentationView').text(t('presentationView'));
    $('.i18n-copyShareLink').text(t('copyShareLink'));
    $('.i18n-generateBtn').text(t('generateBtn'));
    $('.i18n-slideCountAuto').text(t('slideCountAuto'));
    $('#instructionsInput').attr('placeholder', t('instructionsPlaceholder'));

    // 모달
    $('.i18n-newProjectTitle').text(t('newProjectTitle'));
    $('.i18n-editProject').text(t('editProject'));
    $('.i18n-projectName').text(t('projectName'));
    $('.i18n-description').text(t('description'));
    $('.i18n-cancel').text(t('cancel'));
    $('.i18n-create').text(t('create'));
    $('.i18n-save').text(t('save'));
    $('.i18n-add').text(t('add'));
    $('.i18n-close').text(t('close'));
    $('.i18n-textResourceTitle').text(t('textResourceTitle'));
    $('.i18n-textResTitle').text(t('textResTitle'));
    $('.i18n-textResContent').text(t('textResContent'));
    $('.i18n-webSearchTitle').text(t('webSearchTitle'));
    $('.i18n-searchQuery').text(t('searchQuery'));
    $('.i18n-searchBtn').text(t('searchBtn'));
    $('.i18n-searchResult').text(t('searchResult'));

    // placeholder
    $('#newProjectName').attr('placeholder', t('projectNamePlaceholder'));
    $('#newProjectDesc').attr('placeholder', t('descriptionPlaceholder'));
    $('#textResTitle').attr('placeholder', t('textResTitlePlaceholder'));
    $('#textResContent').attr('placeholder', t('textResContentPlaceholder'));
    $('#webSearchQuery').attr('placeholder', t('searchQueryPlaceholder'));

    document.title = t('appTitle') + ' - ' + t('appSubtitle');
}

// ============ 초기화 ============
$(document).ready(function () {
    const detectedLang = parseLangFromUrl();
    if (detectedLang) state.lang = detectedLang;
    applyI18n();

    const pathParts = window.location.pathname.split('/').filter(Boolean);
    let jwtSegment = pathParts[0] || '';
    if (pathParts.length >= 2 && detectedLang && pathParts[pathParts.length - 1].toLowerCase() === detectedLang) {
        jwtSegment = pathParts[0];
    }

    if (jwtSegment && jwtSegment !== 'app' && jwtSegment !== 'shared' && jwtSegment.length > 20) {
        state.jwtToken = jwtSegment;
        verifyAndInit();
    } else if (jwtSegment === 'shared') {
        loadSharedPresentation(pathParts[1]);
    } else {
        showLogin();
    }

    initUserSearch();
});

function hideAllViews() {
    $('#loginView').hide();
    $('#appView').hide();
    $('#authCheckingView').hide();
    $('#authErrorView').hide();
}

function showLogin() {
    hideAllViews();
    $('#loginView').css('display', 'flex');
}

function showApp() {
    hideAllViews();
    $('#appView').css('display', 'flex');
}

function showAuthChecking() {
    hideAllViews();
    $('#authCheckingView').css('display', 'flex');
}

function showAuthError() {
    hideAllViews();
    $('#authErrorView').css('display', 'flex');
}

// ============ 사용자 검색 ============
function initUserSearch() {
    let debounceTimer = null;

    $('#loginUserSearch').on('input', function () {
        const query = $(this).val().trim();
        if (query.length < 1) {
            $('#userSearchDropdown').hide();
            state.searchResults = [];
            return;
        }
        clearTimeout(debounceTimer);
        debounceTimer = setTimeout(() => searchUsers(query), 300);
    });

    $('#loginUserSearch').on('keydown', function (e) {
        const dropdown = $('#userSearchDropdown');
        if (!dropdown.is(':visible')) return;
        const items = dropdown.find('.user-search-item');
        const maxIdx = items.length - 1;

        if (e.key === 'ArrowDown') {
            e.preventDefault();
            state.searchHighlightIndex = Math.min(state.searchHighlightIndex + 1, maxIdx);
            highlightSearchItem(items);
        } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            state.searchHighlightIndex = Math.max(state.searchHighlightIndex - 1, 0);
            highlightSearchItem(items);
        } else if (e.key === 'Enter') {
            e.preventDefault();
            if (state.searchHighlightIndex >= 0 && state.searchHighlightIndex <= maxIdx) {
                selectSearchUser(state.searchHighlightIndex);
            }
        } else if (e.key === 'Escape') {
            dropdown.hide();
        }
    });

    $(document).on('click', function (e) {
        if (!$(e.target).closest('.user-search-wrapper').length) {
            $('#userSearchDropdown').hide();
        }
    });
}

async function searchUsers(query) {
    try {
        const res = await fetch('/api/auth/search-users', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name: query }),
        });
        const data = await res.json();
        state.searchResults = data.users || [];
        state.searchHighlightIndex = -1;
        renderSearchDropdown();
    } catch (e) {
        console.error('Search failed:', e);
    }
}

function renderSearchDropdown() {
    const dropdown = $('#userSearchDropdown');
    dropdown.empty();
    if (state.searchResults.length === 0) {
        dropdown.html(`<div style="padding:12px;text-align:center;color:#9ca3af;font-size:13px;">${t('noSearchResult')}</div>`);
        dropdown.show();
        return;
    }
    state.searchResults.forEach((user, idx) => {
        dropdown.append(`
            <div class="user-search-item" data-index="${idx}" onclick="selectSearchUser(${idx})">
                <div class="name">${escapeHtml(user.nm)}</div>
                <div class="dept">${escapeHtml(user.dp || '')} · ${escapeHtml(user.em || '')}</div>
            </div>
        `);
    });
    dropdown.show();
}

function highlightSearchItem(items) {
    items.removeClass('highlighted');
    if (state.searchHighlightIndex >= 0) {
        $(items[state.searchHighlightIndex]).addClass('highlighted');
        items[state.searchHighlightIndex].scrollIntoView({ block: 'nearest' });
    }
}

function selectSearchUser(index) {
    const user = state.searchResults[index];
    if (!user) return;
    $('#loginUserSearch').val(user.nm);
    $('#loginUserKey').val(user.ky);
    $('#selectedUserName').text(user.nm);
    $('#selectedUserDept').text(user.dp || '');
    $('#selectedUserInfo').show();
    $('#userSearchDropdown').hide();
}

// ============ 로그인/인증 ============
async function doLogin() {
    const userKey = $('#loginUserKey').val();
    const password = $('#loginPassword').val();

    if (!userKey) { showToast(t('msgSelectUser'), 'error'); return; }
    if (!password) { showToast(t('msgEnterPassword'), 'error'); return; }

    try {
        const res = await fetch('/api/auth/login', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ user_key: userKey, password: password }),
        });
        if (!res.ok) {
            const err = await res.json();
            throw new Error(err.detail || 'Login failed');
        }
        const data = await res.json();
        state.jwtToken = data.token;
        state.userInfo = data.user;
        state.isAdmin = (data.user.role || '').toLowerCase() === 'admin';
        window.history.replaceState({}, '', '/' + state.jwtToken + '/' + state.lang);
        showApp();
        initApp();
    } catch (e) {
        showToast(e.message, 'error');
    }
}

async function verifyAndInit() {
    showAuthChecking();
    try {
        const res = await fetch('/api/auth/verify/' + state.jwtToken);
        if (!res.ok) throw new Error('Token expired');
        const data = await res.json();
        state.userInfo = data.user;
        state.isAdmin = data.is_admin || false;
        showApp();
        initApp();
    } catch (e) {
        showAuthError();
    }
}

function logout() {
    state.jwtToken = '';
    state.userInfo = null;
    state.currentProject = null;
    window.history.replaceState({}, '', '/');
    showLogin();
}

async function initApp() {
    // 사용자 정보 표시
    const nm = state.userInfo.nm || '';
    $('#sidebarUserName').text(nm);
    $('#sidebarUserDept').text(state.userInfo.dp || '');
    $('#userAvatar').text(nm.charAt(0).toUpperCase());

    // 관리자 버튼 표시/숨김
    if (state.isAdmin) {
        $('#btnAdminHeader').show();
        $('#btnAdminMain').show();
    } else {
        $('#btnAdminHeader').hide();
        $('#btnAdminMain').hide();
    }

    applyI18n();
    await loadProjects();
    await loadTemplates();
    await loadSupportedLangs();
    await loadAndApplyWebFonts();

    // 빈 상태 표시
    showEmptyState();

    // 슬라이드 드래그 & 드롭 초기화
    initSlideDragDrop('#slideThumbList', '.slide-thumb-v');
    initSlideDragDrop('#gridContainer', '.grid-thumb');
}

function goToAdmin() {
    window.open('/admin?token=' + state.jwtToken, '_blank');
}

// ============ 웹 폰트 로딩 ============
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

async function loadAndApplyWebFonts() {
    try {
        const res = await apiGet('/api/fonts');
        state.fonts = res.fonts || [];
        loadWebFonts(state.fonts);
    } catch (e) {
        console.warn('Failed to load fonts:', e);
    }
}

// ============ API 헬퍼 ============
function apiUrl(path) {
    return '/' + state.jwtToken + path;
}

async function apiGet(path) {
    const res = await fetch(apiUrl(path));
    if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.detail || 'API Error');
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
        throw new Error(err.detail || 'API Error');
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
        throw new Error(err.detail || 'API Error');
    }
    return res.json();
}

async function apiDelete(path) {
    const res = await fetch(apiUrl(path), { method: 'DELETE' });
    if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.detail || 'API Error');
    }
    return res.json();
}

async function apiUpload(path, formData) {
    const res = await fetch(apiUrl(path), { method: 'POST', body: formData });
    if (!res.ok) {
        const err = await res.json().catch(() => ({}));
        throw new Error(err.detail || 'API Error');
    }
    return res.json();
}

// ============ 사이드바 ============
function toggleSidebar() {
    state.sidebarCollapsed = !state.sidebarCollapsed;
    $('#sidebar').toggleClass('collapsed', state.sidebarCollapsed);
    $('#sidebarOpenBtn').toggle(state.sidebarCollapsed);
}

function showEmptyState() {
    $('#emptyState').show();
    $('#projectWorkspace').hide();
    if (state.isAdmin) $('#btnAdminMain').show();
    renderRecentProjects();
}

function renderRecentProjects() {
    const grid = $('#recentProjectsGrid');
    grid.empty();
    const recent = state.projects.slice(0, 6);
    if (recent.length === 0) return;

    recent.forEach(p => {
        const statusLabel = { draft: t('statusDraft'), preparing: t('statusPreparing'), generated: t('statusGenerated') }[p.status] || t('statusDraft');
        const date = p.created_at ? new Date(p.created_at).toLocaleDateString() : '';
        grid.append(`
            <div class="recent-card" onclick="openProject('${p._id}')">
                <div class="rc-title">${escapeHtml(p.name)}</div>
                <div class="rc-desc">${escapeHtml(p.description || t('noDesc'))}</div>
                <div class="rc-meta">
                    <span class="rc-date">${date}</span>
                    <span class="rc-status ${p.status || 'draft'}">${statusLabel}</span>
                </div>
            </div>
        `);
    });
}

// ============ 프로젝트 관리 ============
async function loadProjects() {
    try {
        const res = await apiGet('/api/projects');
        state.projects = res.projects || [];
        renderProjectList();
    } catch (e) {
        showToast(t('msgLoadingProject'), 'error');
    }
}

function renderProjectList() {
    const list = $('#projectList');
    list.empty();

    if (state.projects.length === 0) {
        list.html(`<div style="padding:16px 12px;text-align:center;color:var(--sidebar-text);font-size:12px;opacity:0.6;">${t('noResources')}</div>`);
        return;
    }

    state.projects.forEach(p => {
        const isActive = state.currentProject && state.currentProject._id === p._id;
        const date = p.created_at ? new Date(p.created_at).toLocaleDateString() : '';
        list.append(`
            <div class="project-item ${isActive ? 'active' : ''}" onclick="openProject('${p._id}')">
                <div class="proj-icon">
                    <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2"><rect x="2" y="2" width="12" height="12" rx="2"/><path d="M5 6h6M5 9h4"/></svg>
                </div>
                <div class="proj-info">
                    <div class="proj-name">${escapeHtml(p.name)}</div>
                    <div class="proj-date">${date}</div>
                </div>
                <div class="proj-status-dot ${p.status || 'draft'}"></div>
            </div>
        `);
    });
}

function showNewProjectModal() {
    $('#newProjectName').val('');
    $('#newProjectDesc').val('');
    $('#newProjectModal').show();
}

async function createProject() {
    const name = $('#newProjectName').val().trim();
    if (!name) { showToast(t('msgEnterProjectName'), 'error'); return; }

    try {
        const res = await apiPost('/api/projects', {
            name: name,
            description: $('#newProjectDesc').val().trim(),
        });
        closeModal('newProjectModal');
        showToast(t('msgProjectCreated'), 'success');
        await loadProjects();
        // 새로 만든 프로젝트 바로 열기
        if (res.project_id) {
            openProject(res.project_id);
        } else {
            renderRecentProjects();
        }
    } catch (e) {
        showToast(e.message, 'error');
    }
}

async function openProject(projectId) {
    try {
        showLoading(t('msgLoadingProject'));
        const res = await apiGet('/api/projects/' + projectId);
        state.currentProject = res.project;
        state.resources = res.resources || [];
        state.generatedSlides = res.generated_slides || [];
        state.currentSlideIndex = 0;

        hideLoading();
        renderProjectWorkspace();
        renderProjectList(); // 사이드바 활성 상태 업데이트
    } catch (e) {
        hideLoading();
        showToast(t('msgLoadingProject'), 'error');
    }
}

function renderProjectWorkspace() {
    $('#emptyState').hide();
    $('#btnAdminMain').hide();
    $('#projectWorkspace').css('display', 'flex');

    // 헤더
    $('#wsProjectTitle').text(state.currentProject.name);
    const statusLabel = { draft: t('statusDraft'), preparing: t('statusPreparing'), generated: t('statusGenerated') }[state.currentProject.status] || t('statusDraft');
    $('#wsProjectStatus').text(statusLabel).attr('class', 'ws-status ' + (state.currentProject.status || 'draft'));

    // 리소스 칩
    renderResourceChips();

    // 슬라이드 영역
    renderSlideArea();

    // 지침 복원
    if (state.currentProject.instructions) {
        $('#instructionsInput').val(state.currentProject.instructions);
        autoResizeTextarea(document.getElementById('instructionsInput'));
    } else {
        $('#instructionsInput').val('');
    }

    // 템플릿 선택 복원
    state.selectedTemplateId = state.currentProject.template_id || null;
    updateTemplateButtonLabel();
}

function showEditProjectModal() {
    if (!state.currentProject) return;
    $('#editProjectName').val(state.currentProject.name);
    $('#editProjectDesc').val(state.currentProject.description || '');
    $('#editProjectModal').show();
}

async function updateProject() {
    try {
        await apiPut('/api/projects/' + state.currentProject._id, {
            name: $('#editProjectName').val().trim(),
            description: $('#editProjectDesc').val().trim(),
        });
        state.currentProject.name = $('#editProjectName').val().trim();
        state.currentProject.description = $('#editProjectDesc').val().trim();
        $('#wsProjectTitle').text(state.currentProject.name);
        closeModal('editProjectModal');
        showToast(t('msgUpdated'), 'success');
        await loadProjects();
    } catch (e) {
        showToast(e.message, 'error');
    }
}

async function deleteProject() {
    if (!confirm(t('msgConfirmDelete'))) return;
    try {
        await apiDelete('/api/projects/' + state.currentProject._id);
        showToast(t('msgProjectDeleted'), 'success');
        state.currentProject = null;
        await loadProjects();
        showEmptyState();
    } catch (e) {
        showToast(e.message, 'error');
    }
}

// ============ 프로젝트 초기화 ============
function showResetConfirm() {
    if (!state.currentProject) return;
    $('#resetConfirmModal').show();
}

async function executeProjectReset() {
    closeModal('resetConfirmModal');
    if (!state.currentProject) return;
    try {
        showLoading(t('msgResetting') || '초기화 중...');
        await apiPost('/api/projects/' + state.currentProject._id + '/reset', {});
        state.generatedSlides = [];
        state.currentProject.status = 'draft';
        state.currentProject.template_id = null;
        state.selectedTemplateId = null;
        updateTemplateButtonLabel();
        hideLoading();
        renderProjectWorkspace();
        showToast(t('msgResetDone') || '프로젝트가 초기화되었습니다', 'success');
    } catch (e) {
        hideLoading();
        showToast(e.message, 'error');
    }
}

// ============ 리소스 관리 ============
function renderResourceChips() {
    const count = state.resources.length;
    const summary = $('#resourceSummary');
    const countText = $('#resourceCountText');
    const list = $('#resourceDropdownList');

    // 요약 버튼 표시/숨김
    if (count > 0) {
        countText.text(count + '개 리소스');
        summary.show();
    } else {
        summary.hide();
        $('#resourceDropdown').hide();
        summary.removeClass('open');
    }

    // 드롭다운 리스트 렌더링
    list.empty();
    const icons = { file: '📎', text: '📝', web: '🔍' };
    const typeLabels = { file: 'File', text: 'Text', web: 'Web' };

    state.resources.forEach(r => {
        const icon = icons[r.resource_type] || '📄';
        const iconClass = r.resource_type || 'file';
        const typeLabel = typeLabels[r.resource_type] || r.resource_type;
        list.append(`
            <div class="resource-dd-item" onclick="showResourceContent('${r._id}')">
                <div class="resource-dd-icon ${iconClass}">${icon}</div>
                <div class="resource-dd-info">
                    <div class="resource-dd-title">${escapeHtml(r.title || r.resource_type)}</div>
                    <div class="resource-dd-type">${typeLabel}</div>
                </div>
                <button class="resource-dd-remove" onclick="event.stopPropagation();deleteResource('${r._id}')" title="Remove">&times;</button>
            </div>
        `);
    });
}

function toggleResourceDropdown() {
    const dropdown = $('#resourceDropdown');
    const summary = $('#resourceSummary');
    const isOpen = dropdown.is(':visible');
    if (isOpen) {
        dropdown.hide();
        summary.removeClass('open');
    } else {
        dropdown.show();
        summary.addClass('open');
    }
}

// 드롭다운 외부 클릭 시 닫기 (모달 내부 클릭은 제외)
$(document).on('click', function(e) {
    if ($(e.target).closest('.modal, .modal-overlay').length) return;
    if (!$(e.target).closest('#resourceSummary, #resourceDropdown').length) {
        $('#resourceDropdown').hide();
        $('#resourceSummary').removeClass('open');
    }
});

function showTextResourceModal() {
    $('#textResTitle').val('');
    $('#textResContent').val('');
    $('#textResourceModal').show();
}

async function addTextResource() {
    const title = $('#textResTitle').val().trim();
    const content = $('#textResContent').val().trim();
    if (!content) { showToast(t('msgEnterContent'), 'error'); return; }

    try {
        const res = await apiPost('/api/resources/text', {
            project_id: state.currentProject._id,
            resource_type: 'text',
            title: title || 'Text',
            content: content,
        });
        state.resources.push(res.resource);
        renderResourceChips();
        closeModal('textResourceModal');
        showToast(t('msgTextAdded'), 'success');
    } catch (e) {
        showToast(e.message, 'error');
    }
}

function triggerFileUpload() {
    $('#fileUploadInput').click();
}

async function handleFileUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const formData = new FormData();
    formData.append('file', file);
    formData.append('project_id', state.currentProject._id);
    formData.append('title', file.name);

    try {
        showLoading(t('msgUploading'));
        const res = await apiUpload('/api/resources/file', formData);
        state.resources.push(res.resource);
        renderResourceChips();
        hideLoading();
        showToast(t('msgFileUploaded'), 'success');
    } catch (e) {
        hideLoading();
        showToast(e.message, 'error');
    }
    event.target.value = '';
}

function showWebSearchModal() {
    $('#webSearchQuery').val('');
    $('#webSearchResult').hide();
    state.lastWebSearchResult = null;
    $('#webSearchModal').show();
}

async function doWebSearch() {
    const query = $('#webSearchQuery').val().trim();
    if (!query) { showToast(t('msgEnterQuery'), 'error'); return; }

    try {
        $('#btnWebSearch').prop('disabled', true).find('span').text(t('searching'));
        const res = await apiPost('/api/resources/web-search', {
            project_id: state.currentProject._id,
            query: query,
        });

        const resources = res.resources || [];
        if (resources.length === 0) {
            showToast('검색 결과가 없습니다', 'error');
            return;
        }

        resources.forEach(function (r) { state.resources.push(r); });
        state.lastWebSearchResult = resources[0];
        renderResourceChips();

        // 검색 결과 요약 표시
        var summaryLines = resources.map(function (r, i) {
            var title = r.title || ('페이지 ' + (i + 1));
            var url = r.source_url || '';
            var preview = (r.content || '').substring(0, 200);
            return '[ ' + (i + 1) + '. ' + title + ' ]\n' + (url ? url + '\n' : '') + preview + '...\n';
        });
        $('#webSearchResultContent').text(summaryLines.join('\n'));
        $('#webSearchResult').show();
        showToast(res.count + '개 페이지가 리소스로 추가되었습니다', 'success');
    } catch (e) {
        showToast(e.message, 'error');
    } finally {
        $('#btnWebSearch').prop('disabled', false).find('span').text(t('searchBtn'));
    }
}

async function showResourceContent(resourceId) {
    const resource = state.resources.find(r => r._id === resourceId);
    if (!resource) return;

    const icons = { file: '📎', text: '📝', web: '🔍' };
    const iconClasses = { file: 'file', text: 'text', web: 'web' };
    const icon = icons[resource.resource_type] || '📄';
    const iconClass = iconClasses[resource.resource_type] || 'file';

    $('#resourceContentIcon').attr('class', 'chip-icon ' + iconClass).text(icon);
    $('#resourceContentTitle').text(resource.title || resource.resource_type);

    // 이미 content가 로컬에 있으면 사용, 없으면 서버에서 조회
    let content = resource.content;
    if (!content && content !== '') {
        try {
            const res = await apiGet('/api/resources/content/' + resourceId);
            content = res.content || '';
            // 로컬 캐시 업데이트
            resource.content = content;
        } catch (e) {
            content = '';
        }
    }

    const viewer = $('#resourceContentBody');
    if (content && content.trim()) {
        if (resource.resource_type === 'text') {
            // 사용자 직접 입력 텍스트는 그대로 표시
            viewer.html('<pre style="white-space:pre-wrap;font-family:inherit;margin:0;">' + escapeHtml(content) + '</pre>');
        } else {
            // 파일/검색 결과는 마크다운 렌더링
            viewer.html(simpleMarkdownToHtml(content));
        }
    } else {
        viewer.html('<div class="empty-content">내용이 없습니다</div>');
    }

    $('#resourceContentModal').show();
}

function simpleMarkdownToHtml(md) {
    if (!md) return '';
    let html = escapeHtml(md);

    // 수평선
    html = html.replace(/^---$/gm, '<hr>');

    // 헤딩 (#### → h4, ### → h3, ## → h2, # → h1)
    html = html.replace(/^#### (.+)$/gm, '<h4>$1</h4>');
    html = html.replace(/^### (.+)$/gm, '<h3>$1</h3>');
    html = html.replace(/^## (.+)$/gm, '<h2>$1</h2>');
    html = html.replace(/^# (.+)$/gm, '<h1>$1</h1>');

    // 볼드
    html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');

    // 이탤릭
    html = html.replace(/\*(.+?)\*/g, '<em>$1</em>');

    // 마크다운 테이블 변환
    html = html.replace(/((?:^\|.+\|$\n?)+)/gm, function(tableBlock) {
        const lines = tableBlock.trim().split('\n');
        if (lines.length < 2) return tableBlock;

        let tableHtml = '<table>';
        lines.forEach((line, idx) => {
            const cells = line.split('|').filter((_, i, arr) => i > 0 && i < arr.length - 1);
            // 구분선 행 (| --- | --- |) 건너뜀
            if (cells.every(c => /^\s*-{2,}\s*$/.test(c))) return;

            const tag = idx === 0 ? 'th' : 'td';
            const wrap = idx === 0 ? 'thead' : (idx === 2 || (idx === 1 && !lines[1].match(/^\|[\s-|]+\|$/)) ? '' : '');
            tableHtml += '<tr>';
            cells.forEach(c => {
                tableHtml += `<${tag}>${c.trim()}</${tag}>`;
            });
            tableHtml += '</tr>';
        });
        tableHtml += '</table>';
        return tableHtml;
    });

    // 리스트
    html = html.replace(/^- (.+)$/gm, '<li>$1</li>');
    html = html.replace(/(<li>.*<\/li>\n?)+/g, '<ul>$&</ul>');

    // 번호 리스트
    html = html.replace(/^\d+\.\s+(.+)$/gm, '<li>$1</li>');

    // 링크
    html = html.replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2" target="_blank">$1</a>');

    // URL 자동 링크 (http로 시작하는 것)
    html = html.replace(/(^|[^"'>])(https?:\/\/[^\s<]+)/g, '$1<a href="$2" target="_blank">$2</a>');

    // 줄바꿈
    html = html.replace(/\n\n/g, '</p><p>');
    html = '<p>' + html + '</p>';

    // 빈 p 태그 정리
    html = html.replace(/<p>\s*<\/p>/g, '');
    html = html.replace(/<p>\s*(<h[1-4]>)/g, '$1');
    html = html.replace(/(<\/h[1-4]>)\s*<\/p>/g, '$1');
    html = html.replace(/<p>\s*(<table>)/g, '$1');
    html = html.replace(/(<\/table>)\s*<\/p>/g, '$1');
    html = html.replace(/<p>\s*(<ul>)/g, '$1');
    html = html.replace(/(<\/ul>)\s*<\/p>/g, '$1');
    html = html.replace(/<p>\s*(<hr>)/g, '$1');
    html = html.replace(/(<hr>)\s*<\/p>/g, '$1');

    return html;
}

async function deleteResource(resourceId) {
    if (!confirm(t('msgConfirmDeleteRes'))) return;
    try {
        await apiDelete('/api/resources/' + resourceId);
        state.resources = state.resources.filter(r => r._id !== resourceId);
        renderResourceChips();
        showToast(t('msgDeleted'), 'success');
    } catch (e) {
        showToast(e.message, 'error');
    }
}

// ============ 템플릿 / 언어 ============
async function loadTemplates() {
    try {
        const res = await apiGet('/api/templates');
        state.templates = res.templates || [];
        updateTemplateButtonLabel();
    } catch (e) {
        console.error('Template load failed:', e);
    }
}

function updateTemplateButtonLabel() {
    if (state.selectedTemplateId) {
        const tmpl = state.templates.find(t => t._id === state.selectedTemplateId);
        if (tmpl) {
            $('#templateSelectLabel').text(tmpl.name);
            $('#templateSelectBtn').addClass('has-value');
            return;
        }
    }
    $('#templateSelectLabel').text(t('selectTemplatePlaceholder') || '템플릿을 선택하세요');
    $('#templateSelectBtn').removeClass('has-value');
}

function openTemplatePickerModal() {
    const grid = $('#templatePickerGrid');
    grid.empty();

    if (state.templates.length === 0) {
        grid.append('<div style="text-align:center;padding:40px;color:var(--text-muted);">등록된 템플릿이 없습니다</div>');
        $('#templatePickerModal').show();
        return;
    }

    state.templates.forEach(tmpl => {
        const isActive = state.selectedTemplateId === tmpl._id;
        const card = $(`
            <div class="template-picker-card ${isActive ? 'active' : ''}" onclick="selectTemplateFromPicker('${tmpl._id}')">
                <div class="template-picker-thumb"></div>
                <div class="template-picker-info">
                    <div class="template-picker-name">${escapeHtml(tmpl.name)}</div>
                    <div class="template-picker-count">${tmpl.slide_count || 0}개 슬라이드</div>
                </div>
            </div>
        `);

        // 첫 번째 슬라이드 썸네일 렌더링
        const thumbContainer = card.find('.template-picker-thumb');
        if (tmpl.first_slide) {
            const slideData = { ...tmpl.first_slide };
            // 슬라이드 배경이 없으면 템플릿 배경 사용
            if (!slideData.background_image && tmpl.background_image) {
                slideData.background_image = tmpl.background_image;
            }
            renderSlideToContainer(thumbContainer, slideData, 240, 135);
        }

        grid.append(card);
    });

    $('#templatePickerModal').show();
}

function selectTemplateFromPicker(templateId) {
    state.selectedTemplateId = templateId;
    updateTemplateButtonLabel();
    closeModal('templatePickerModal');
}

async function loadSupportedLangs() {
    try {
        const res = await fetch('/api/config/langs');
        if (!res.ok) return;
        const data = await res.json();
        const select = $('#langSelect');
        select.empty().append(`<option value="">${t('defaultLang')}</option>`);
        (data.langs || []).forEach(l => {
            const selected = l.code === data.default ? ' selected' : '';
            select.append(`<option value="${l.code}"${selected}>${l.label}</option>`);
        });
    } catch (e) {
        console.error('Lang load failed:', e);
    }
}

// ============ PPT 생성 ============
async function generatePPT() {
    const templateId = state.selectedTemplateId;
    const instructions = $('#instructionsInput').val().trim();
    const lang = $('#langSelect').val();
    const slideCount = $('#slideCountSelect').val();

    if (!templateId) { showToast(t('msgSelectTemplate'), 'error'); return; }
    if (state.resources.length === 0) { showToast(t('msgAddResources'), 'error'); return; }

    $('#btnGenerate').prop('disabled', true);
    _animationCancelled = true;
    state.generatedSlides = [];
    state.currentSlideIndex = 0;

    // 로딩 오버레이 대신 슬라이드 프리뷰 영역 표시 + 아웃라인 탭
    $('#slideEmpty').hide();
    $('#slidePreview').css('display', 'flex');
    $('#wsSlideTools').css('display', 'flex');
    switchPanelTab('outline');

    // 스트리밍 프로그레스 UI 표시 (내부 스크롤 모드)
    $('#slideTextList').addClass('streaming-active').empty().append(`
        <div id="streamingProgress" class="streaming-progress">
            <div class="streaming-progress-header">
                <div class="streaming-spinner-sm"></div>
                <span id="streamingStatus">AI가 슬라이드를 설계하고 있습니다...</span>
            </div>
            <div id="streamingContent" class="streaming-content"></div>
        </div>
    `);

    // 이전 생성 결과 초기화 (썸네일, 카운터)
    $('#slideThumbnails').empty();
    $('#slideThumbList').empty();
    $('#slideCounter').text('0 / 0');
    $('#slideCounterInline').text('0 / 0');

    // 캔버스 영역에 로딩 애니메이션 표시
    $('#previewCanvas .preview-obj').remove();
    $('#previewBg').css('background-image', 'none');
    $('#canvasLoadingOverlay').remove();
    $('#previewCanvas').append(`
        <div class="canvas-loading-overlay" id="canvasLoadingOverlay">
            <div class="canvas-loading-animation">
                <div class="canvas-loading-ring"></div>
                <div class="loading-slide"></div>
                <div class="loading-slide"></div>
                <div class="loading-slide"></div>
            </div>
            <div class="canvas-loading-text">
                <div class="canvas-loading-title">슬라이드를 만들고 있습니다</div>
                <div class="canvas-loading-sub">AI가 최적의 레이아웃을 설계하는 중<span class="canvas-loading-dots"><span>.</span><span>.</span><span>.</span></span></div>
            </div>
        </div>
    `);

    try {
        const response = await fetch(`/${state.jwtToken}/api/generate/stream`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                project_id: state.currentProject._id,
                template_id: templateId,
                instructions: instructions,
                lang: lang,
                slide_count: slideCount,
            }),
        });

        if (!response.ok) {
            let errMsg = '생성 실패';
            try { const err = await response.json(); errMsg = err.detail || errMsg; } catch (_) {}
            throw new Error(errMsg);
        }

        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;

            buffer += decoder.decode(value, { stream: true });
            const parts = buffer.split('\n\n');
            buffer = parts.pop() || '';

            for (const part of parts) {
                for (const line of part.split('\n')) {
                    if (!line.startsWith('data: ')) continue;
                    try {
                        const data = JSON.parse(line.slice(6));
                        _handleStreamEvent(data);
                    } catch (e) {
                        console.warn('SSE parse error:', e);
                    }
                }
            }
        }
    } catch (e) {
        showToast(e.message || '슬라이드 생성 중 오류가 발생했습니다.', 'error');
        console.error('Generate stream error:', e);
        if (state.generatedSlides.length === 0) {
            $('#slideEmpty').show();
            $('#slidePreview').hide();
            $('#wsSlideTools').hide();
        }
    } finally {
        $('#canvasLoadingOverlay').remove();
        $('#btnGenerate').prop('disabled', false);
        $('#streamingProgress').remove();
        $('#slideTextList').removeClass('streaming-active');
    }
}

function _handleStreamEvent(data) {
    switch (data.event) {
        case 'template_analysis': {
            // 템플릿 타입 분석 결과 표시
            const types = data.types || {};
            const typeOrder = ['title_slide', 'toc', 'section_divider', 'body', 'closing'];
            let analysisHtml = '<div class="template-analysis">';
            analysisHtml += '<div class="analysis-title">템플릿 분석</div>';
            typeOrder.forEach(key => {
                const info = types[key];
                if (!info) return;
                const icon = info.available ? '✓' : '✗';
                const cls = info.available ? 'available' : 'missing';
                analysisHtml += `<div class="analysis-item ${cls}">`;
                analysisHtml += `<span class="analysis-icon">${icon}</span>`;
                analysisHtml += `<span class="analysis-label">${escapeHtml(info.label)}</span>`;
                if (info.available && info.count > 0) {
                    analysisHtml += `<span class="analysis-count">${info.count}개</span>`;
                }
                analysisHtml += '</div>';
            });
            analysisHtml += '</div>';

            // 스트리밍 프로그레스 앞에 삽입
            const progress = $('#streamingProgress');
            if (progress.length) {
                progress.before(analysisHtml);
            }
            break;
        }

        case 'start':
            $('#streamingStatus').text(data.message || 'AI가 슬라이드를 설계하고 있습니다...');
            break;

        case 'delta': {
            const el = document.getElementById('streamingContent');
            if (el) {
                const text = data.text || '';
                el.textContent += text;
                el.scrollTop = el.scrollHeight;
            }
            break;
        }

        case 'parsing':
            $('#streamingStatus').text(data.message || '슬라이드를 구성하고 있습니다...');
            $('#streamingContent').css('opacity', '0.4');
            break;

        case 'outline': {
            // 아웃라인 데이터 표시
            const outlineSlides = data.slides || [];
            const typeLabelsMap = {
                title: 'Cover', toc: 'Contents',
                section: 'Chapter', content: 'Body', closing: 'Closing',
            };

            let outlineHtml = '<div class="outline-summary">';
            outlineHtml += '<div class="outline-summary-title">아웃라인</div>';
            outlineSlides.forEach((s, i) => {
                const badge = typeLabelsMap[s.type] || s.type;
                outlineHtml += `<div class="outline-summary-item">`;
                outlineHtml += `<span class="outline-summary-num">${i + 1}</span>`;
                outlineHtml += `<span class="outline-summary-text">${escapeHtml(s.title || '슬라이드 ' + (i + 1))}</span>`;
                outlineHtml += `<span class="slide-type-badge">${badge}</span>`;
                outlineHtml += '</div>';
            });
            outlineHtml += '</div>';

            // 스트리밍 프로그레스를 아웃라인으로 교체
            $('#streamingProgress').replaceWith(outlineHtml);
            break;
        }

        case 'slide': {
            // 첫 슬라이드 도착 시 로딩 오버레이 제거
            $('#canvasLoadingOverlay').remove();

            const slide = data.slide;
            state.generatedSlides.push(slide);
            const idx = state.generatedSlides.length - 1;
            state.currentSlideIndex = idx;

            // 캔버스에 현재 슬라이드 렌더링
            renderSlideAtIndex(idx);
            renderSlideThumbnails();
            renderSlideThumbList();
            updateSlideNav();
            break;
        }

        case 'complete':
            state.currentSlideIndex = 0;
            // 슬라이드 프리뷰 영역만 표시 (썸네일은 애니메이션에서 하나씩 추가)
            $('#slideEmpty').hide();
            $('#slidePreview').css('display', 'flex');
            $('#wsSlideTools').css('display', 'flex');
            renderSlideTextPanel();
            // Slide 탭으로 전환 후 타이핑 애니메이션 시작
            animateSlideGeneration();
            break;

        case 'error':
            showToast(data.message || '생성 중 오류가 발생했습니다.', 'error');
            break;
    }
}

function _appendSlideOutline(slide, index) {
    const typeLabels = {
        title_slide: 'Cover', toc: 'Contents',
        section_divider: 'Chapter', body: '', closing: 'Closing',
    };

    const textObjs = (slide.objects || []).filter(o => o.obj_type === 'text');
    let titleText = '';
    const sections = [];

    // 타이틀 추출
    textObjs.forEach(obj => {
        let text = (obj.generated_text || '').trim();
        if (!text) {
            const fb = (obj.text_content || '').trim();
            if (fb && fb !== '텍스트를 입력하세요' && fb !== 'Enter text') text = fb;
        }
        if (!text) return;
        const role = obj.role || obj._auto_role || '';
        const fontSize = (obj.text_style || {}).font_size || 16;
        if (!titleText && (role === 'title' || fontSize >= 24)) titleText = text;
    });

    // items 사용
    (slide.items || []).forEach(item => {
        sections.push({ header: item.heading || '', body: item.detail || '' });
    });

    const meta = slide.slide_meta || {};
    const typeLabel = typeLabels[meta.content_type || ''] || '';
    const badgeHtml = typeLabel ? `<span class="slide-type-badge">${typeLabel}</span>` : '';

    let sectionsHtml = '';
    sections.forEach(sec => {
        if (sec.header || sec.body) {
            sectionsHtml += '<div class="outline-section">';
            if (sec.header) sectionsHtml += `<div class="outline-section-title">${escapeHtml(sec.header)}</div>`;
            if (sec.body) sectionsHtml += `<div class="outline-section-body">${escapeHtml(sec.body)}</div>`;
            sectionsHtml += '</div>';
        }
    });

    const el = $(`
        <div class="slide-text-item stream-in" onclick="goToSlide(${index})" data-slide-idx="${index}">
            <div class="slide-text-item-header">
                <div class="slide-text-item-num">${index + 1}</div>
                <div class="slide-text-item-title">${escapeHtml(titleText || '슬라이드 ' + (index + 1))}</div>
                ${badgeHtml}
            </div>
            ${sectionsHtml}
        </div>
    `);

    // 스트리밍 프로그레스 바로 앞에 삽입
    const progress = $('#streamingProgress');
    if (progress.length) {
        progress.before(el);
    } else {
        $('#slideTextList').append(el);
    }

    // 새 슬라이드로 스크롤
    el[0].scrollIntoView({ block: 'nearest', behavior: 'smooth' });
}

// ============ 슬라이드 프리뷰 ============
function renderSlideArea() {
    if (state.generatedSlides.length === 0) {
        $('#slideEmpty').show();
        $('#slidePreview').hide();
        $('#wsSlideTools').hide();
        // 썸네일/캔버스/아웃라인 등 잔여 콘텐츠 제거
        $('#slideThumbnails').empty();
        $('#slideThumbList').empty();
        $('#slideTextList').empty();
        $('#previewCanvas').find('.preview-obj').remove();
        $('#slideCounter').text('0 / 0');
        $('#slideCounterInline').text('0 / 0');
        state.currentSlideIndex = 0;
        return;
    }

    $('#slideEmpty').hide();
    $('#slidePreview').css('display', 'flex');
    $('#wsSlideTools').css('display', 'flex');
    renderSlideThumbList();
    renderSlideTextPanel();
    renderSlideAtIndex(state.currentSlideIndex);
    renderSlideThumbnails();
    updateSlideNav();
}

function renderSlideAtIndex(index) {
    const slide = state.generatedSlides[index];
    if (!slide) return;

    const canvas = $('#previewCanvas');
    canvas.find('.preview-obj').remove();

    if (slide.background_image) {
        $('#previewBg').css('background-image', `url(${slide.background_image})`);
    } else {
        $('#previewBg').css('background-image', 'none');
    }

    const canvasW = canvas.width();
    const canvasH = canvas.height();
    const scaleX = canvasW / 960;
    const scaleY = canvasH / 540;

    let descIndex = 0; // description 오브젝트 인덱스 카운터
    let subIndex = 0;  // subtitle 오브젝트 인덱스 카운터
    const items = slide.items || [];

    (slide.objects || []).forEach(obj => {
        const div = $('<div>').addClass('preview-obj').css({
            position: 'absolute',
            left: (obj.x * scaleX) + 'px',
            top: (obj.y * scaleY) + 'px',
            width: (obj.width * scaleX) + 'px',
            height: (obj.height * scaleY) + 'px',
            zIndex: 10,
        });

        if (obj.obj_type === 'image' && obj.image_url) {
            const imgFit = obj.image_fit || 'contain';
            div.append(`<img src="${obj.image_url}" style="width:100%;height:100%;object-fit:${imgFit};">`);
        } else if (obj.obj_type === 'text') {
            const style = obj.text_style || {};
            const text = obj.generated_text || obj.text_content || '';
            const scaledFontSize = (style.font_size || 16) * Math.min(scaleX, scaleY);
            const role = obj.role || obj._auto_role || '';

            div.css({
                fontFamily: style.font_family || 'Inter, Arial, sans-serif',
                fontSize: scaledFontSize + 'px',
                color: style.color || '#000',
                fontWeight: style.bold ? 'bold' : 'normal',
                fontStyle: style.italic ? 'italic' : 'normal',
                textAlign: style.align || 'left',
                padding: (8 * scaleX) + 'px',
                overflow: 'hidden',
                wordWrap: 'break-word',
            });
            div.attr('data-text', text);

            // subtitle 역할: items[].heading을 직접 매핑
            if (role === 'subtitle' && items.length > 0 && subIndex < items.length) {
                const headingText = items[subIndex].heading || '';
                div.css('whiteSpace', 'pre-wrap');
                div.text(headingText);
                subIndex++;
            // description 역할: items[].detail만 표시
            } else if (role === 'description' && items.length > 0 && descIndex < items.length) {
                const item = items[descIndex];
                const detailText = item.detail || '';
                div.css('whiteSpace', 'pre-wrap');
                div.text(detailText);
                descIndex++;
                // 텍스트가 넘칠 경우 높이 자동 확장
                div.css({ height: 'auto', minHeight: (obj.height * scaleY) + 'px', overflow: 'visible' });
            } else if ((role === 'subtitle' || role === 'description') && items.length > 0) {
                // items 소진된 초과 subtitle/description 오브젝트는 렌더링하지 않음
                return;
            } else {
                div.css('whiteSpace', 'pre-wrap');
                div.text(text);
                // description 역할은 텍스트 넘침 시 높이 자동 확장
                if (role === 'description') {
                    div.css({ height: 'auto', minHeight: (obj.height * scaleY) + 'px', overflow: 'visible' });
                }
            }
        }

        canvas.append(div);
    });
}

function renderSlideToContainer(container, slide, thumbW, thumbH) {
    const bgStyle = slide.background_image ? `background-image:url(${slide.background_image});background-size:cover;background-position:center;` : '';
    container.attr('style', (container.attr('style') || '') + bgStyle);

    const scaleX = thumbW / 960;
    const scaleY = thumbH / 540;
    let thumbDescIdx = 0;
    let thumbSubIdx = 0;
    const thumbItems = slide.items || [];

    (slide.objects || []).forEach(obj => {
        const div = $('<div>').css({
            position: 'absolute',
            left: (obj.x * scaleX) + 'px',
            top: (obj.y * scaleY) + 'px',
            width: (obj.width * scaleX) + 'px',
            height: (obj.height * scaleY) + 'px',
            overflow: 'hidden',
            pointerEvents: 'none',
        });

        if (obj.obj_type === 'image' && obj.image_url) {
            const imgFit = obj.image_fit || 'contain';
            div.append(`<img src="${obj.image_url}" style="width:100%;height:100%;object-fit:${imgFit};">`);
        } else if (obj.obj_type === 'shape') {
            // 도형은 축소 렌더링 생략 (너무 작음)
        } else if (obj.obj_type === 'text') {
            const style = obj.text_style || {};
            const text = obj.generated_text || obj.text_content || '';
            const scaledFontSize = Math.max(1, (style.font_size || 16) * Math.min(scaleX, scaleY));
            const role = obj.role || obj._auto_role || '';

            div.css({
                fontFamily: style.font_family || 'Arial, sans-serif',
                fontSize: scaledFontSize + 'px',
                lineHeight: '1.3',
                color: style.color || '#000',
                fontWeight: style.bold ? 'bold' : 'normal',
                fontStyle: style.italic ? 'italic' : 'normal',
                textAlign: style.align || 'left',
                wordWrap: 'break-word',
            });

            // subtitle: items[].heading 직접 매핑
            if (role === 'subtitle' && thumbItems.length > 0 && thumbSubIdx < thumbItems.length) {
                div.css('whiteSpace', 'pre-wrap');
                div.text(thumbItems[thumbSubIdx].heading || '');
                thumbSubIdx++;
            // description: items[].detail만 렌더링
            } else if (role === 'description' && thumbItems.length > 0 && thumbDescIdx < thumbItems.length) {
                div.css('whiteSpace', 'pre-wrap');
                div.text(thumbItems[thumbDescIdx].detail || '');
                thumbDescIdx++;
            } else if ((role === 'subtitle' || role === 'description') && thumbItems.length > 0) {
                // items 소진된 초과 subtitle/description 오브젝트는 렌더링하지 않음
                return;
            } else {
                div.css('whiteSpace', 'pre-wrap');
                div.text(text);
            }
        }

        container.append(div);
    });
}

function renderSlideThumbnails() {
    const container = $('#slideThumbnails');
    container.empty();
    state.generatedSlides.forEach((slide, i) => {
        const isActive = i === state.currentSlideIndex;
        const thumbEl = $(`
            <div class="slide-thumb ${isActive ? 'active' : ''}" onclick="goToSlide(${i})">
                <div class="slide-thumb-inner"></div>
            </div>
        `);
        renderSlideToContainer(thumbEl.find('.slide-thumb-inner'), slide, 64, 36);
        container.append(thumbEl);
    });
}

function renderSlideThumbList() {
    const list = $('#slideThumbList');
    list.empty();
    state.generatedSlides.forEach((slide, i) => {
        const isActive = i === state.currentSlideIndex;
        const thumbEl = $(`
            <div class="slide-thumb-v ${isActive ? 'active' : ''}" draggable="true" onclick="goToSlide(${i})" data-slide-idx="${i}">
                <div class="slide-thumb-v-num">${i + 1}</div>
                <div class="slide-thumb-v-inner"></div>
            </div>
        `);
        renderSlideToContainer(thumbEl.find('.slide-thumb-v-inner'), slide, 256, 144);
        list.append(thumbEl);
    });
}

function _appendSingleThumbnail(i) {
    const slide = state.generatedSlides[i];
    if (!slide) return;
    const container = $('#slideThumbnails');
    // 이전 active 해제
    container.find('.slide-thumb').removeClass('active');
    const thumbEl = $(`
        <div class="slide-thumb active" onclick="goToSlide(${i})">
            <div class="slide-thumb-inner"></div>
        </div>
    `);
    renderSlideToContainer(thumbEl.find('.slide-thumb-inner'), slide, 64, 36);
    container.append(thumbEl);
    // 새 썸네일이 보이도록 스크롤
    thumbEl[0].scrollIntoView({ inline: 'nearest', behavior: 'smooth' });
}

function _appendSingleThumbV(i) {
    const slide = state.generatedSlides[i];
    if (!slide) return;
    const list = $('#slideThumbList');
    // 이전 active 해제
    list.find('.slide-thumb-v').removeClass('active');
    const thumbEl = $(`
        <div class="slide-thumb-v active" onclick="goToSlide(${i})" data-slide-idx="${i}">
            <div class="slide-thumb-v-num">${i + 1}</div>
            <div class="slide-thumb-v-inner"></div>
        </div>
    `);
    renderSlideToContainer(thumbEl.find('.slide-thumb-v-inner'), slide, 256, 144);
    list.append(thumbEl);
    thumbEl[0].scrollIntoView({ block: 'nearest', behavior: 'smooth' });
}

function renderSlideTextPanel() {
    const list = $('#slideTextList');
    list.empty();

    // 슬라이드 유형 매핑
    const typeLabels = {
        title_slide: 'Cover',
        toc: 'Contents',
        section_divider: 'Chapter',
        body: '',
        closing: 'Closing',
    };

    state.generatedSlides.forEach((slide, i) => {
        const isActive = i === state.currentSlideIndex;
        const textObjs = (slide.objects || []).filter(o => o.obj_type === 'text');

        // 역할별로 텍스트 오브젝트 분류
        let titleText = '';
        const sections = [];

        // 먼저 title 텍스트 추출
        textObjs.forEach(obj => {
            let text = (obj.generated_text || '').trim();
            if (!text) {
                const fallback = (obj.text_content || '').trim();
                if (fallback && fallback !== '텍스트를 입력하세요' && fallback !== 'Enter text') {
                    text = fallback;
                }
            }
            if (!text) return;
            const role = obj.role || obj._auto_role || '';
            const fontSize = (obj.text_style || {}).font_size || 16;
            if (!titleText && (role === 'title' || fontSize >= 24)) {
                titleText = text;
            }
        });

        // items 배열이 있으면 구조화된 데이터 사용 (우선)
        const slideItems = slide.items || [];
        if (slideItems.length > 0) {
            slideItems.forEach(item => {
                sections.push({
                    header: item.heading || '',
                    body: item.detail || ''
                });
            });
        } else {
            // items 없으면 기존 텍스트 파싱 로직 (하위호환)
            textObjs.forEach(obj => {
                let text = (obj.generated_text || '').trim();
                if (!text) {
                    const fallback = (obj.text_content || '').trim();
                    if (fallback && fallback !== '텍스트를 입력하세요' && fallback !== 'Enter text') {
                        text = fallback;
                    }
                }
                if (!text) return;
                const role = obj.role || obj._auto_role || '';
                const fontSize = (obj.text_style || {}).font_size || 16;

                if (role === 'title' || fontSize >= 24) {
                    // 이미 추출됨
                } else if (role === 'governance') {
                    // 거버넌스 텍스트는 생략
                } else if (role === 'subtitle' || (fontSize >= 16 && fontSize < 24 && (obj.text_style || {}).bold)) {
                    sections.push({ header: text, body: '' });
                } else {
                    // 본문 텍스트 - "heading\ndetail" 형식 파싱 시도
                    const blocks = text.split('\n\n');
                    if (blocks.length > 1) {
                        // 여러 블록이 있으면 각 블록을 heading/detail로 분리
                        blocks.forEach(block => {
                            const lines = block.split('\n');
                            if (lines.length >= 2) {
                                sections.push({ header: lines[0], body: lines.slice(1).join('\n') });
                            } else if (lines[0]) {
                                sections.push({ header: '', body: lines[0] });
                            }
                        });
                    } else if (sections.length > 0 && !sections[sections.length - 1].body) {
                        sections[sections.length - 1].body = text;
                    } else {
                        sections.push({ header: '', body: text });
                    }
                }
            });
        }

        // 슬라이드 유형 뱃지
        const meta = slide.slide_meta || {};
        const contentType = meta.content_type || '';
        const typeLabel = typeLabels[contentType] || '';
        const badgeHtml = typeLabel ? `<span class="slide-type-badge">${typeLabel}</span>` : '';

        // 섹션 HTML 생성
        let sectionsHtml = '';
        sections.forEach(sec => {
            if (sec.header || sec.body) {
                sectionsHtml += '<div class="outline-section">';
                if (sec.header) {
                    sectionsHtml += `<div class="outline-section-title">${escapeHtml(sec.header)}</div>`;
                }
                if (sec.body) {
                    sectionsHtml += `<div class="outline-section-body">${escapeHtml(sec.body)}</div>`;
                }
                sectionsHtml += '</div>';
            }
        });

        list.append(`
            <div class="slide-text-item ${isActive ? 'active' : ''}" onclick="goToSlide(${i})" data-slide-idx="${i}">
                <div class="slide-text-item-header">
                    <div class="slide-text-item-num">${i + 1}</div>
                    <div class="slide-text-item-title">${escapeHtml(titleText || '슬라이드 ' + (i + 1))}</div>
                    ${badgeHtml}
                </div>
                ${sectionsHtml}
            </div>
        `);
    });
}

// ============ 패널 탭 / 그리드 뷰 ============

function switchPanelTab(tab) {
    $('.panel-tab').removeClass('active');
    $(`.panel-tab[data-tab="${tab}"]`).addClass('active');
    $('.panel-tab-content').removeClass('active');
    if (tab === 'slide') {
        $('#panelTabSlide').addClass('active');
    } else {
        $('#panelTabOutline').addClass('active');
    }
}

function toggleGridView() {
    const overlay = $('#gridOverlay');
    if (overlay.is(':visible')) {
        overlay.hide();
    } else {
        renderGridView();
        overlay.show();
    }
}

function renderGridView() {
    const container = $('#gridContainer');
    container.empty();
    const total = state.generatedSlides.length;
    $('#gridSlideCount').text(`${state.currentSlideIndex + 1} / ${total}`);

    state.generatedSlides.forEach((slide, i) => {
        const isActive = i === state.currentSlideIndex;
        const texts = (slide.objects || [])
            .filter(o => o.obj_type === 'text')
            .map(o => o.generated_text || o.text_content || '')
            .filter(t => t.trim());
        const titleText = texts[0] ? texts[0].substring(0, 40) : '슬라이드 ' + (i + 1);

        const gridEl = $(`
            <div class="grid-thumb ${isActive ? 'active' : ''}" draggable="true" data-slide-idx="${i}" onclick="goToSlide(${i}); toggleGridView();">
                <div class="grid-thumb-inner">
                    <div class="grid-thumb-num">${i + 1}</div>
                </div>
                <div class="grid-thumb-title">${escapeHtml(titleText)}</div>
            </div>
        `);
        renderSlideToContainer(gridEl.find('.grid-thumb-inner'), slide, 320, 180);
        container.append(gridEl);
    });
}

function updateSlideNav() {
    const total = state.generatedSlides.length;
    const current = state.currentSlideIndex + 1;
    $('#slideCounter').text(`${current} / ${total}`);
    $('#slideCounterInline').text(`${current} / ${total}`);
    $('#btnPrevSlide').prop('disabled', state.currentSlideIndex === 0);
    $('#btnNextSlide').prop('disabled', state.currentSlideIndex >= total - 1);

    // 하단 썸네일 활성 상태 업데이트
    $('.slide-thumb').removeClass('active');
    $(`.slide-thumb:eq(${state.currentSlideIndex})`).addClass('active');

    // 좌측 세로 썸네일 활성 상태 업데이트
    $('.slide-thumb-v').removeClass('active');
    $(`.slide-thumb-v[data-slide-idx="${state.currentSlideIndex}"]`).addClass('active');
    const activeThumb = $(`.slide-thumb-v[data-slide-idx="${state.currentSlideIndex}"]`);
    if (activeThumb.length) {
        activeThumb[0].scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    }

    // 좌측 아웃라인 패널 활성 상태 업데이트
    $('.slide-text-item').removeClass('active');
    $(`.slide-text-item[data-slide-idx="${state.currentSlideIndex}"]`).addClass('active');
    const activeItem = $(`.slide-text-item[data-slide-idx="${state.currentSlideIndex}"]`);
    if (activeItem.length) {
        activeItem[0].scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    }
}

function goToSlide(index) {
    _animationCancelled = true;
    if (state.editMode) collectEditedText();
    state.currentSlideIndex = index;
    if (state.editMode) {
        renderSlideAtIndexEditable(index);
        _updateEditSlideCounter();
    } else {
        renderSlideAtIndex(index);
    }
    updateSlideNav();
}

function prevSlide() {
    if (state.currentSlideIndex > 0) {
        _animationCancelled = true;
        if (state.editMode) collectEditedText();
        state.currentSlideIndex--;
        if (state.editMode) {
            renderSlideAtIndexEditable(state.currentSlideIndex);
            _updateEditSlideCounter();
        } else {
            renderSlideAtIndex(state.currentSlideIndex);
        }
        updateSlideNav();
    }
}

function nextSlide() {
    if (state.currentSlideIndex < state.generatedSlides.length - 1) {
        _animationCancelled = true;
        if (state.editMode) collectEditedText();
        state.currentSlideIndex++;
        if (state.editMode) {
            renderSlideAtIndexEditable(state.currentSlideIndex);
            _updateEditSlideCounter();
        } else {
            renderSlideAtIndex(state.currentSlideIndex);
        }
        updateSlideNav();
    }
}

// ============ 타이핑 애니메이션 ============
async function animateSlideGeneration() {
    _animationCancelled = false;
    _isAnimating = true;
    switchPanelTab('slide');

    // 애니메이션 시작 시 썸네일/카운터 초기화 (슬라이드 완성 후 하나씩 추가)
    $('#slideThumbnails').empty();
    $('#slideThumbList').empty();
    $('#slideCounter').text(`0 / ${state.generatedSlides.length}`);
    $('#slideCounterInline').text(`0 / ${state.generatedSlides.length}`);

    for (let i = 0; i < state.generatedSlides.length; i++) {
        if (_animationCancelled) break;

        state.currentSlideIndex = i;
        renderSlideAtIndex(i);

        // 텍스트가 있는 모든 프리뷰 오브젝트 수집
        const textEls = [];
        $('#previewCanvas .preview-obj').each(function () {
            const $el = $(this);
            const plainText = $el.text().trim();
            if (plainText) {
                textEls.push({
                    $el: $el,
                    savedHtml: $el.html(),
                    plainText: plainText,
                });
            }
        });

        // 모든 텍스트 요소 비우기
        textEls.forEach(item => item.$el.empty());

        // 각 텍스트 요소를 하나씩 타이핑
        for (const item of textEls) {
            if (_animationCancelled) break;

            const $el = item.$el;
            const cursor = $('<span class="typing-cursor"></span>');
            $el.append(cursor);

            const text = item.plainText;
            for (let k = 0; k < text.length; k++) {
                if (_animationCancelled) break;
                const ch = text[k];
                if (ch === '\n') {
                    cursor.before($('<br>')[0]);
                } else {
                    cursor.before(document.createTextNode(ch));
                }
                if (k % 2 === 0) await sleep(30);
            }
            cursor.remove();

            // 원래 포맷된 HTML 복원
            if (!_animationCancelled) {
                $el.html(item.savedHtml);
            }
        }

        // 타이핑 완료 후 해당 슬라이드 썸네일 추가
        if (!_animationCancelled) {
            _appendSingleThumbnail(i);
            _appendSingleThumbV(i);
            $('#slideCounter').text(`${i + 1} / ${state.generatedSlides.length}`);
            $('#slideCounterInline').text(`${i + 1} / ${state.generatedSlides.length}`);
        }

        // 슬라이드 간 전환 딜레이
        if (!_animationCancelled && i < state.generatedSlides.length - 1) {
            await sleep(600);
        }
    }

    _isAnimating = false;

    // 애니메이션 완료/취소 후 처리
    if (_animationCancelled) {
        // 취소 시: 아직 추가 안 된 썸네일을 모두 채우기
        const addedCount = $('#slideThumbList .slide-thumb-v').length;
        for (let j = addedCount; j < state.generatedSlides.length; j++) {
            _appendSingleThumbnail(j);
            _appendSingleThumbV(j);
        }
        $('#slideCounter').text(`${state.currentSlideIndex + 1} / ${state.generatedSlides.length}`);
        $('#slideCounterInline').text(`${state.currentSlideIndex + 1} / ${state.generatedSlides.length}`);
        updateSlideNav();
    } else {
        // 정상 완료: 1페이지로 이동 (썸네일 재빌드 없이 active만 변경)
        state.currentSlideIndex = 0;
        renderSlideAtIndex(0);
        updateSlideNav();
        requestAnimationFrame(() => {
            $('#slideThumbList').scrollTop(0);
            $('#slideArea').scrollTop(0);
            $('.slide-canvas-area').scrollTop(0);
        });
        showToast(t('msgAllComplete'), 'success');
    }
}

function sleep(ms) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// ============ 슬라이드 편집 모드 ============

let _editImageReplaceObjIdx = -1;

function toggleEditMode() {
    if (state.editMode) {
        exitEditMode();
    } else {
        enterEditMode();
    }
}

function enterEditMode() {
    if (state.generatedSlides.length === 0) return;
    state.editMode = true;
    state.editTool = 'select';
    $('#previewCanvas').addClass('edit-mode');
    $('#btnEditToggle').addClass('active');
    $('#btnEditSave').show();
    // 하단 썸네일 숨기고 편집 도구 모음 표시
    $('.slide-nav').hide();
    $('#editBottomToolbar').show();
    _updateEditSlideCounter();
    _populateEditFontSelector();
    renderSlideAtIndexEditable(state.currentSlideIndex);
}

function exitEditMode() {
    collectEditedText();
    if (state.editDirtySlides.size > 0) {
        if (confirm(t('editUnsavedChanges'))) {
            saveCurrentSlide();
            return; // saveCurrentSlide이 _exitEditModeClean 호출
        }
    }
    _exitEditModeClean();
}

function editGetScale() {
    const canvas = document.getElementById('previewCanvas');
    const rect = canvas.getBoundingClientRect();
    return { x: rect.width / 960, y: rect.height / 540 };
}

function renderSlideAtIndexEditable(index) {
    const slide = state.generatedSlides[index];
    if (!slide) return;

    const canvas = $('#previewCanvas');
    canvas.find('.preview-obj').remove();

    if (slide.background_image) {
        $('#previewBg').css('background-image', `url(${slide.background_image})`);
    } else {
        $('#previewBg').css('background-image', 'none');
    }

    const canvasW = canvas.width();
    const canvasH = canvas.height();
    const scaleX = canvasW / 960;
    const scaleY = canvasH / 540;

    let descIndex = 0;
    let subIndex = 0;
    const items = slide.items || [];

    (slide.objects || []).forEach((obj, objIdx) => {
        const role = obj.role || obj._auto_role || '';
        let currentItemIdx = -1;
        let displayText = '';
        let isHidden = false;

        if (obj.obj_type === 'text') {
            if (role === 'subtitle' && items.length > 0) {
                if (subIndex < items.length) {
                    displayText = items[subIndex].heading || '';
                    currentItemIdx = subIndex;
                    subIndex++;
                } else {
                    isHidden = true;
                    subIndex++;
                }
            } else if (role === 'description' && items.length > 0) {
                if (descIndex < items.length) {
                    displayText = items[descIndex].detail || '';
                    currentItemIdx = descIndex;
                    descIndex++;
                } else {
                    isHidden = true;
                    descIndex++;
                }
            } else {
                displayText = obj.generated_text || obj.text_content || '';
            }
        }

        if (isHidden) return;

        const div = $('<div>').addClass('preview-obj').css({
            position: 'absolute',
            left: (obj.x * scaleX) + 'px',
            top: (obj.y * scaleY) + 'px',
            width: (obj.width * scaleX) + 'px',
            height: (obj.height * scaleY) + 'px',
            zIndex: 10,
        });
        div.attr('data-obj-idx', objIdx);
        div.attr('data-obj-type', obj.obj_type);
        div.attr('data-role', role);
        if (currentItemIdx >= 0) div.attr('data-item-idx', currentItemIdx);

        if (obj.obj_type === 'image' && obj.image_url) {
            const imgFit = obj.image_fit || 'contain';
            div.append(`<img src="${obj.image_url}" style="width:100%;height:100%;object-fit:${imgFit};pointer-events:none;">`);
            div.append(`<button class="edit-img-replace-btn" onclick="triggerEditImageReplace(${objIdx})">${t('editReplaceImage')}</button>`);
        } else if (obj.obj_type === 'shape') {
            const svg = _createEditShapeSVG(obj);
            div.css('pointerEvents', 'all').html(svg);
        } else if (obj.obj_type === 'text') {
            const style = obj.text_style || {};
            const scaledFontSize = (style.font_size || 16) * Math.min(scaleX, scaleY);
            const decoParts = [];
            if (style.underline) decoParts.push('underline');
            if (style.strikethrough) decoParts.push('line-through');

            const textDiv = $('<div>').addClass('edit-text-content').css({
                fontFamily: style.font_family || 'Inter, Arial, sans-serif',
                fontSize: scaledFontSize + 'px',
                color: style.color || '#000',
                fontWeight: style.bold ? 'bold' : 'normal',
                fontStyle: style.italic ? 'italic' : 'normal',
                textAlign: style.align || 'left',
                textDecoration: decoParts.length ? decoParts.join(' ') : 'none',
                padding: (8 * scaleX) + 'px',
                whiteSpace: 'pre-wrap',
                wordWrap: 'break-word',
                outline: 'none',
                width: '100%',
                height: '100%',
                boxSizing: 'border-box',
            }).attr('contenteditable', 'true')
              .text(displayText);

            textDiv.on('input', function () {
                markSlideDirty(index);
            });

            textDiv.on('mousedown', function (e) {
                if (div.hasClass('edit-selected')) {
                    e.stopPropagation();
                }
            });

            div.append(textDiv);

            if (role === 'description') {
                div.css({ height: 'auto', minHeight: (obj.height * scaleY) + 'px', overflow: 'visible' });
            }
        }

        // Resize handles + delete button
        div.append('<div class="edit-resize-handle se"></div>');
        div.append('<div class="edit-resize-handle sw"></div>');
        div.append('<div class="edit-resize-handle ne"></div>');
        div.append('<div class="edit-resize-handle nw"></div>');
        div.append(`<button class="edit-obj-delete-btn" onclick="deleteEditSelectedObject()" title="Delete">×</button>`);

        div.on('mousedown', function (e) {
            if ($(e.target).hasClass('edit-resize-handle')) {
                startEditResize(e, objIdx, e.target);
            } else if ($(e.target).hasClass('edit-img-replace-btn')) {
                // let button handler work
            } else if ($(e.target).hasClass('edit-text-content')) {
                editSelectObject(objIdx);
            } else {
                editSelectObject(objIdx);
                startEditDrag(e, objIdx);
            }
        });

        canvas.append(div);
    });
}

// ---- Selection ----
function editSelectObject(objIdx) {
    const slide = state.generatedSlides[state.currentSlideIndex];
    if (!slide) return;
    state.editSelectedObj = { objIdx: objIdx, obj: slide.objects[objIdx] };
    $('#previewCanvas .preview-obj').removeClass('edit-selected');
    $(`#previewCanvas .preview-obj[data-obj-idx="${objIdx}"]`).addClass('edit-selected');

    const obj = slide.objects[objIdx];
    if (obj && obj.obj_type === 'text') {
        showEditTextToolbar(objIdx);
    } else {
        $('#editTextToolbar').hide();
    }
}

function editDeselectAll() {
    state.editSelectedObj = null;
    $('#previewCanvas .preview-obj').removeClass('edit-selected');
    $('#editTextToolbar').hide();
}

// ---- Drag ----
function startEditDrag(e, objIdx) {
    state.editDragging = true;
    const slide = state.generatedSlides[state.currentSlideIndex];
    const obj = slide.objects[objIdx];
    if (!obj) return;
    const canvasRect = document.getElementById('previewCanvas').getBoundingClientRect();
    const scale = editGetScale();
    state.editDragOffset = {
        x: ((e.clientX - canvasRect.left) / scale.x) - obj.x,
        y: ((e.clientY - canvasRect.top) / scale.y) - obj.y,
    };
    e.preventDefault();
}

// ---- Resize ----
function startEditResize(e, objIdx, handleEl) {
    state.editResizing = true;
    editSelectObject(objIdx);
    const obj = state.editSelectedObj.obj;
    const cls = handleEl.className;
    if (cls.includes(' se')) state.editResizeDir = 'se';
    else if (cls.includes(' sw')) state.editResizeDir = 'sw';
    else if (cls.includes(' ne')) state.editResizeDir = 'ne';
    else if (cls.includes(' nw')) state.editResizeDir = 'nw';
    const scale = editGetScale();
    state.editResizeStart = {
        mouseX: e.clientX, mouseY: e.clientY,
        w: obj.width, h: obj.height, ox: obj.x, oy: obj.y,
        scaleX: scale.x, scaleY: scale.y,
    };
    e.preventDefault();
    e.stopPropagation();
}

function handleEditResize(e) {
    const s = state.editResizeStart;
    const obj = state.editSelectedObj.obj;
    const dx = (e.clientX - s.mouseX) / s.scaleX;
    const dy = (e.clientY - s.mouseY) / s.scaleY;
    let newW = s.w, newH = s.h, newX = s.ox, newY = s.oy;
    const minW = 50, minH = 30;
    switch (state.editResizeDir) {
        case 'se': newW = Math.max(minW, s.w + dx); newH = Math.max(minH, s.h + dy); break;
        case 'sw': newW = Math.max(minW, s.w - dx); newH = Math.max(minH, s.h + dy); newX = s.ox + (s.w - newW); break;
        case 'ne': newW = Math.max(minW, s.w + dx); newH = Math.max(minH, s.h - dy); newY = s.oy + (s.h - newH); break;
        case 'nw': newW = Math.max(minW, s.w - dx); newH = Math.max(minH, s.h - dy); newX = s.ox + (s.w - newW); newY = s.oy + (s.h - newH); break;
    }
    // Boundary constraint
    newX = Math.max(0, newX); newY = Math.max(0, newY);
    if (newX + newW > 960) newW = 960 - newX;
    if (newY + newH > 540) newH = 540 - newY;

    obj.x = Math.round(newX); obj.y = Math.round(newY);
    obj.width = Math.round(newW); obj.height = Math.round(newH);

    const scale = editGetScale();
    const el = $(`#previewCanvas .preview-obj[data-obj-idx="${state.editSelectedObj.objIdx}"]`);
    el.css({
        left: (newX * scale.x) + 'px', top: (newY * scale.y) + 'px',
        width: (newW * scale.x) + 'px', height: (newH * scale.y) + 'px',
    });
    if (obj.obj_type === 'text') {
        const style = obj.text_style || {};
        el.find('.edit-text-content').css('fontSize', ((style.font_size || 16) * Math.min(scale.x, scale.y)) + 'px');
    }
    markSlideDirty(state.currentSlideIndex);
}

// ---- Global mouse handlers for edit mode ----
$(document).on('mousemove', function (e) {
    if (!state.editMode) return;
    if (state.editDragging && state.editSelectedObj) {
        const canvasRect = document.getElementById('previewCanvas').getBoundingClientRect();
        const scale = editGetScale();
        const obj = state.editSelectedObj.obj;
        let x = ((e.clientX - canvasRect.left) / scale.x) - state.editDragOffset.x;
        let y = ((e.clientY - canvasRect.top) / scale.y) - state.editDragOffset.y;
        x = Math.max(0, Math.min(x, 960 - obj.width));
        y = Math.max(0, Math.min(y, 540 - obj.height));
        obj.x = Math.round(x); obj.y = Math.round(y);
        const el = $(`#previewCanvas .preview-obj[data-obj-idx="${state.editSelectedObj.objIdx}"]`);
        el.css({ left: (x * scale.x) + 'px', top: (y * scale.y) + 'px' });
        markSlideDirty(state.currentSlideIndex);
    }
    if (state.editResizing && state.editSelectedObj) {
        handleEditResize(e);
    }
});

$(document).on('mouseup', function () {
    if (!state.editMode) return;
    state.editDragging = false;
    state.editResizing = false;
});

// Canvas background click → deselect or insert text
$('#previewCanvas').on('mousedown', function (e) {
    if (!state.editMode) return;
    if (e.target === this || e.target.id === 'previewBg') {
        if (state.editTool === 'text') {
            // 클릭 위치에 텍스트 오브젝트 삽입
            const canvasRect = this.getBoundingClientRect();
            const scale = editGetScale();
            const x = Math.round((e.clientX - canvasRect.left) / scale.x);
            const y = Math.round((e.clientY - canvasRect.top) / scale.y);
            const slide = state.generatedSlides[state.currentSlideIndex];
            if (slide) {
                slide.objects.push({
                    obj_id: 'obj_' + Date.now(),
                    obj_type: 'text',
                    x: Math.max(0, x - 150), y: Math.max(0, y - 25),
                    width: 300, height: 50,
                    text_content: '', generated_text: '',
                    text_style: { font_family: 'Inter', font_size: 16, color: '#000000', bold: false, italic: false, align: 'left' },
                });
                markSlideDirty(state.currentSlideIndex);
                renderSlideAtIndexEditable(state.currentSlideIndex);
                setEditTool('select');
            }
        } else {
            editDeselectAll();
        }
    }
});

// ESC → deselect or exit edit mode, Delete → 오브젝트 삭제
$(document).on('keydown', function (e) {
    if (!state.editMode) return;
    if (e.key === 'Escape') {
        if (state.editSelectedObj) {
            editDeselectAll();
        } else {
            toggleEditMode();
        }
        e.preventDefault();
    } else if (e.key === 'Delete' && state.editSelectedObj) {
        // 텍스트 편집 중이면 삭제키는 텍스트 삭제로 동작
        const active = document.activeElement;
        if (active && active.getAttribute('contenteditable') === 'true') return;
        deleteEditSelectedObject();
        e.preventDefault();
    }
});

// Window resize → re-render editable canvas
$(window).on('resize', function () {
    if (state.editMode) {
        collectEditedText();
        renderSlideAtIndexEditable(state.currentSlideIndex);
    }
});

// ---- Text Toolbar ----
function showEditTextToolbar(objIdx) {
    const el = $(`#previewCanvas .preview-obj[data-obj-idx="${objIdx}"]`);
    if (!el.length) return;
    const rect = el[0].getBoundingClientRect();
    const toolbar = $('#editTextToolbar');
    const tbHeight = 40;
    let top = rect.top - tbHeight - 8;
    if (top < 4) top = rect.bottom + 4;
    // 좌측 넘침 방지
    let left = rect.left;
    const tbWidth = toolbar.outerWidth() || 460;
    if (left + tbWidth > window.innerWidth - 8) left = window.innerWidth - tbWidth - 8;
    if (left < 4) left = 4;
    toolbar.css({ left: left + 'px', top: top + 'px' }).show();

    const obj = state.editSelectedObj.obj;
    const style = obj.text_style || {};
    // 폰트 색상
    const color = style.color || '#000000';
    $('#editColorBar').css('background', color);
    $('#editFontColorPicker').val(color);
    // 폰트 패밀리
    $('#editFontFamily').val(style.font_family || 'Inter');
    // 폰트 사이즈
    $('#editFontSize').val(String(style.font_size || 16));
    // 토글 버튼 상태
    $('#editBoldBtn').toggleClass('active', !!style.bold);
    $('#editItalicBtn').toggleClass('active', !!style.italic);
    $('#editUnderlineBtn').toggleClass('active', !!style.underline);
    $('#editStrikeBtn').toggleClass('active', !!style.strikethrough);
    // 정렬
    $('#editAlignLeftBtn').toggleClass('active', style.align === 'left' || !style.align);
    $('#editAlignCenterBtn').toggleClass('active', style.align === 'center');
    $('#editAlignRightBtn').toggleClass('active', style.align === 'right');
}

function _populateEditFontSelector() {
    const sel = $('#editFontFamily');
    sel.empty();
    const defaults = [
        { name: 'Inter', family: 'Inter' },
        { name: 'Arial', family: 'Arial' },
        { name: 'Georgia', family: 'Georgia' },
        { name: 'Courier New', family: 'Courier New' },
    ];
    // 서버 폰트 + 기본 폰트
    const allFonts = (state.fonts && state.fonts.length) ? state.fonts : defaults;
    allFonts.forEach(f => {
        sel.append($('<option>').val(f.family).text(f.name || f.family));
    });
}

function _updateEditSlideCounter() {
    $('#editSlideCounter').text(`${state.currentSlideIndex + 1} / ${state.generatedSlides.length}`);
}

// ---- Color Picker ----
const _editColorPalette = [
    '#000000','#434343','#666666','#999999','#b7b7b7','#cccccc','#d9d9d9','#ffffff',
    '#980000','#ff0000','#ff9900','#ffff00','#00ff00','#00ffff','#4a86e8','#0000ff',
    '#9900ff','#ff00ff','#e6b8af','#f4cccc','#fce5cd','#fff2cc','#d9ead3','#d0e0e3',
    '#c9daf8','#cfe2f3','#d9d2e9','#ead1dc','#dd7e6b','#ea9999','#f9cb9c','#ffe599',
    '#b6d7a8','#a2c4c9','#a4c2f4','#9fc5e8','#b4a7d6','#d5a6bd','#cc4125','#e06666',
    '#f6b26b','#ffd966','#93c47d','#76a5af','#6d9eeb','#6fa8dc','#8e7cc3','#c27ba0',
];

function _initEditColorGrid() {
    const grid = $('#editColorGrid');
    if (grid.children().length > 0) return;
    _editColorPalette.forEach(c => {
        grid.append(
            $('<div>').addClass('edit-color-swatch')
                .css('background', c)
                .attr('title', c)
                .on('click', function () {
                    setEditFontColor(c);
                    $('#editColorPicker').hide();
                })
        );
    });
}

function toggleEditColorPicker() {
    const picker = $('#editColorPicker');
    if (picker.is(':visible')) {
        picker.hide();
    } else {
        _initEditColorGrid();
        if (state.editSelectedObj) {
            const style = state.editSelectedObj.obj.text_style || {};
            $('#editColorCustomInput').val(style.color || '#000000');
        }
        picker.show();
    }
}

// 팔레트 외부 클릭 시 닫기
$(document).on('mousedown', function (e) {
    if ($('#editColorPicker').is(':visible') && !$(e.target).closest('.edit-tb-color-wrap').length) {
        $('#editColorPicker').hide();
    }
});

// ---- Text Style Functions ----
function setEditFontColor(color) {
    if (!state.editSelectedObj || state.editSelectedObj.obj.obj_type !== 'text') return;
    const obj = state.editSelectedObj.obj;
    if (!obj.text_style) obj.text_style = {};
    obj.text_style.color = color;
    const el = $(`#previewCanvas .preview-obj[data-obj-idx="${state.editSelectedObj.objIdx}"] .edit-text-content`);
    el.css('color', color);
    $('#editColorBar').css('background', color);
    markSlideDirty(state.currentSlideIndex);
}

function setEditFontFamily(family) {
    if (!state.editSelectedObj || state.editSelectedObj.obj.obj_type !== 'text') return;
    const obj = state.editSelectedObj.obj;
    if (!obj.text_style) obj.text_style = {};
    obj.text_style.font_family = family;
    const el = $(`#previewCanvas .preview-obj[data-obj-idx="${state.editSelectedObj.objIdx}"] .edit-text-content`);
    el.css('fontFamily', family);
    markSlideDirty(state.currentSlideIndex);
}

function setEditFontSize(size) {
    if (!state.editSelectedObj || state.editSelectedObj.obj.obj_type !== 'text') return;
    const obj = state.editSelectedObj.obj;
    if (!obj.text_style) obj.text_style = {};
    obj.text_style.font_size = parseInt(size) || 16;
    const scale = editGetScale();
    const scaledSize = obj.text_style.font_size * Math.min(scale.x, scale.y);
    const el = $(`#previewCanvas .preview-obj[data-obj-idx="${state.editSelectedObj.objIdx}"] .edit-text-content`);
    el.css('fontSize', scaledSize + 'px');
    markSlideDirty(state.currentSlideIndex);
}

function toggleEditBold() {
    if (!state.editSelectedObj || state.editSelectedObj.obj.obj_type !== 'text') return;
    const obj = state.editSelectedObj.obj;
    if (!obj.text_style) obj.text_style = {};
    obj.text_style.bold = !obj.text_style.bold;
    const el = $(`#previewCanvas .preview-obj[data-obj-idx="${state.editSelectedObj.objIdx}"] .edit-text-content`);
    el.css('fontWeight', obj.text_style.bold ? 'bold' : 'normal');
    $('#editBoldBtn').toggleClass('active', obj.text_style.bold);
    markSlideDirty(state.currentSlideIndex);
}

function toggleEditItalic() {
    if (!state.editSelectedObj || state.editSelectedObj.obj.obj_type !== 'text') return;
    const obj = state.editSelectedObj.obj;
    if (!obj.text_style) obj.text_style = {};
    obj.text_style.italic = !obj.text_style.italic;
    const el = $(`#previewCanvas .preview-obj[data-obj-idx="${state.editSelectedObj.objIdx}"] .edit-text-content`);
    el.css('fontStyle', obj.text_style.italic ? 'italic' : 'normal');
    $('#editItalicBtn').toggleClass('active', obj.text_style.italic);
    markSlideDirty(state.currentSlideIndex);
}

function toggleEditUnderline() {
    if (!state.editSelectedObj || state.editSelectedObj.obj.obj_type !== 'text') return;
    const obj = state.editSelectedObj.obj;
    if (!obj.text_style) obj.text_style = {};
    obj.text_style.underline = !obj.text_style.underline;
    const el = $(`#previewCanvas .preview-obj[data-obj-idx="${state.editSelectedObj.objIdx}"] .edit-text-content`);
    _applyTextDecoration(el, obj.text_style);
    $('#editUnderlineBtn').toggleClass('active', obj.text_style.underline);
    markSlideDirty(state.currentSlideIndex);
}

function toggleEditStrikethrough() {
    if (!state.editSelectedObj || state.editSelectedObj.obj.obj_type !== 'text') return;
    const obj = state.editSelectedObj.obj;
    if (!obj.text_style) obj.text_style = {};
    obj.text_style.strikethrough = !obj.text_style.strikethrough;
    const el = $(`#previewCanvas .preview-obj[data-obj-idx="${state.editSelectedObj.objIdx}"] .edit-text-content`);
    _applyTextDecoration(el, obj.text_style);
    $('#editStrikeBtn').toggleClass('active', obj.text_style.strikethrough);
    markSlideDirty(state.currentSlideIndex);
}

function _applyTextDecoration(el, style) {
    const parts = [];
    if (style.underline) parts.push('underline');
    if (style.strikethrough) parts.push('line-through');
    el.css('textDecoration', parts.length ? parts.join(' ') : 'none');
}

function setEditAlign(align) {
    if (!state.editSelectedObj || state.editSelectedObj.obj.obj_type !== 'text') return;
    const obj = state.editSelectedObj.obj;
    if (!obj.text_style) obj.text_style = {};
    obj.text_style.align = align;
    const el = $(`#previewCanvas .preview-obj[data-obj-idx="${state.editSelectedObj.objIdx}"] .edit-text-content`);
    el.css('textAlign', align);
    $('#editAlignLeftBtn').toggleClass('active', align === 'left');
    $('#editAlignCenterBtn').toggleClass('active', align === 'center');
    $('#editAlignRightBtn').toggleClass('active', align === 'right');
    markSlideDirty(state.currentSlideIndex);
}

// ---- Bottom Toolbar Functions ----
function setEditTool(tool) {
    state.editTool = tool;
    $('.edit-bt-btn').removeClass('active');
    $(`#editTool${tool.charAt(0).toUpperCase() + tool.slice(1)}`).addClass('active');
    $('#editShapeMenu').hide();
    if (tool === 'select') {
        $('#previewCanvas').css('cursor', 'default');
    } else if (tool === 'text') {
        $('#previewCanvas').css('cursor', 'crosshair');
    }
}

function toggleEditShapeMenu() {
    $('#editShapeMenu').toggle();
}

function insertEditShape(shapeType) {
    $('#editShapeMenu').hide();
    const slide = state.generatedSlides[state.currentSlideIndex];
    if (!slide) return;
    const newObj = {
        obj_id: 'obj_' + Date.now(),
        obj_type: 'shape',
        x: 330, y: 180, width: 300, height: 180,
        shape_style: {
            shape_type: shapeType,
            fill_color: '#4A90D9',
            fill_opacity: 0.2,
            stroke_color: '#2d5a8e',
            stroke_width: 2,
            stroke_dash: 'solid',
            border_radius: shapeType === 'rounded_rectangle' ? 12 : 0,
            arrow_head: shapeType === 'arrow' ? 'end' : 'none',
        },
    };
    if (shapeType === 'line' || shapeType === 'arrow') {
        newObj.height = 4;
        newObj.width = 300;
        newObj.y = 270;
    }
    slide.objects.push(newObj);
    markSlideDirty(state.currentSlideIndex);
    renderSlideAtIndexEditable(state.currentSlideIndex);
    setEditTool('select');
}

async function handleEditInsertImage(event) {
    const file = event.target.files[0];
    if (!file) return;
    const slide = state.generatedSlides[state.currentSlideIndex];
    if (!slide) return;
    const formData = new FormData();
    formData.append('file', file);
    try {
        const resp = await fetch(`/${state.jwtToken}/api/generate/upload-image`, {
            method: 'POST', body: formData,
        });
        const data = await resp.json();
        if (data.image_url) {
            slide.objects.push({
                obj_id: 'obj_' + Date.now(),
                obj_type: 'image',
                x: 330, y: 150, width: 300, height: 240,
                image_url: data.image_url,
            });
            markSlideDirty(state.currentSlideIndex);
            renderSlideAtIndexEditable(state.currentSlideIndex);
        }
    } catch (e) {
        showToast('Image upload failed', 'error');
    }
    event.target.value = '';
}

// ---- Shape SVG Helper ----
function _createEditShapeSVG(obj) {
    const s = obj.shape_style || {};
    const w = obj.width, h = obj.height;
    const sw = s.stroke_width || 2;
    const fill = s.fill_color || '#4A90D9';
    const opacity = s.fill_opacity !== undefined ? s.fill_opacity : 0.2;
    const stroke = s.stroke_color || '#2d5a8e';
    const dash = s.stroke_dash === 'dashed' ? `stroke-dasharray="${sw * 3} ${sw * 2}"` : (s.stroke_dash === 'dotted' ? `stroke-dasharray="${sw} ${sw}"` : '');
    const half = sw / 2;
    const st = s.shape_type || 'rectangle';
    let inner = '';
    if (st === 'ellipse') {
        inner = `<ellipse cx="${w / 2}" cy="${h / 2}" rx="${w / 2 - half}" ry="${h / 2 - half}" fill="${fill}" fill-opacity="${opacity}" stroke="${stroke}" stroke-width="${sw}" ${dash}/>`;
    } else if (st === 'rounded_rectangle') {
        const r = s.border_radius || 12;
        inner = `<rect x="${half}" y="${half}" width="${w - sw}" height="${h - sw}" rx="${r}" fill="${fill}" fill-opacity="${opacity}" stroke="${stroke}" stroke-width="${sw}" ${dash}/>`;
    } else if (st === 'line' || st === 'arrow') {
        const mid = h / 2;
        let marker = '';
        if (st === 'arrow') {
            marker = `<defs><marker id="ah_${obj.obj_id}" markerWidth="8" markerHeight="6" refX="7" refY="3" orient="auto"><path d="M0,0 L8,3 L0,6" fill="${stroke}"/></marker></defs>`;
            inner = `${marker}<line x1="${half}" y1="${mid}" x2="${w - half}" y2="${mid}" stroke="${stroke}" stroke-width="${sw}" ${dash} marker-end="url(#ah_${obj.obj_id})"/>`;
        } else {
            inner = `<line x1="${half}" y1="${mid}" x2="${w - half}" y2="${mid}" stroke="${stroke}" stroke-width="${sw}" ${dash}/>`;
        }
    } else {
        inner = `<rect x="${half}" y="${half}" width="${w - sw}" height="${h - sw}" fill="${fill}" fill-opacity="${opacity}" stroke="${stroke}" stroke-width="${sw}" ${dash}/>`;
    }
    return `<svg width="100%" height="100%" viewBox="0 0 ${w} ${h}" preserveAspectRatio="none" style="pointer-events:none;">${inner}</svg>`;
}

// ---- Image Replace ----
function triggerEditImageReplace(objIdx) {
    _editImageReplaceObjIdx = objIdx;
    $('#editImageInput').click();
}

async function handleEditImageReplace(event) {
    const file = event.target.files[0];
    if (!file || _editImageReplaceObjIdx < 0) return;
    const slide = state.generatedSlides[state.currentSlideIndex];
    const obj = slide.objects[_editImageReplaceObjIdx];
    if (!obj || obj.obj_type !== 'image') return;

    const formData = new FormData();
    formData.append('file', file);
    try {
        const resp = await fetch(`/${state.jwtToken}/api/generate/upload-image`, {
            method: 'POST',
            body: formData,
        });
        const data = await resp.json();
        if (data.image_url) {
            obj.image_url = data.image_url;
            $(`#previewCanvas .preview-obj[data-obj-idx="${_editImageReplaceObjIdx}"] img`).attr('src', data.image_url);
            markSlideDirty(state.currentSlideIndex);
        }
    } catch (e) {
        showToast('Image upload failed', 'error');
    }
    event.target.value = '';
    _editImageReplaceObjIdx = -1;
}

// ---- Delete Object ----
function deleteEditSelectedObject() {
    if (!state.editSelectedObj) return;
    const slide = state.generatedSlides[state.currentSlideIndex];
    if (!slide) return;
    const objIdx = state.editSelectedObj.objIdx;
    slide.objects.splice(objIdx, 1);
    state.editSelectedObj = null;
    $('#editTextToolbar').hide();
    markSlideDirty(state.currentSlideIndex);
    renderSlideAtIndexEditable(state.currentSlideIndex);
}

// ---- Save ----
function markSlideDirty(index) {
    const slide = state.generatedSlides[index];
    if (slide && slide._id) {
        state.editDirtySlides.add(slide._id);
    }
}

function collectEditedText() {
    const slide = state.generatedSlides[state.currentSlideIndex];
    if (!slide) return;
    $('#previewCanvas .preview-obj').each(function () {
        const el = $(this);
        const objIdx = parseInt(el.attr('data-obj-idx'));
        const objType = el.attr('data-obj-type');
        const role = el.attr('data-role');
        const itemIdx = parseInt(el.attr('data-item-idx'));

        if (objType === 'text') {
            const textEl = el.find('.edit-text-content');
            if (!textEl.length) return;
            const newText = textEl.text();
            if (role === 'subtitle' && !isNaN(itemIdx) && slide.items && slide.items[itemIdx]) {
                slide.items[itemIdx].heading = newText;
            } else if (role === 'description' && !isNaN(itemIdx) && slide.items && slide.items[itemIdx]) {
                slide.items[itemIdx].detail = newText;
            } else if (!isNaN(objIdx) && slide.objects[objIdx]) {
                slide.objects[objIdx].generated_text = newText;
            }
        }
    });
}

async function saveCurrentSlide() {
    collectEditedText();

    // 현재 슬라이드도 dirty에 추가
    const curSlide = state.generatedSlides[state.currentSlideIndex];
    if (curSlide && curSlide._id) {
        state.editDirtySlides.add(curSlide._id);
    }

    // 모든 dirty 슬라이드 저장
    const dirtyIds = [...state.editDirtySlides];
    if (dirtyIds.length === 0) {
        // 변경 없어도 편집 모드 종료
        _exitEditModeClean();
        return;
    }

    let failCount = 0;
    for (const slideId of dirtyIds) {
        const slide = state.generatedSlides.find(s => s._id === slideId);
        if (!slide) continue;
        try {
            const resp = await fetch(`/${state.jwtToken}/api/generate/slides/${slideId}`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ objects: slide.objects, items: slide.items || [] }),
            });
            if (!resp.ok) throw new Error();
            state.editDirtySlides.delete(slideId);
        } catch (e) {
            failCount++;
        }
    }

    if (failCount > 0) {
        showToast(`${failCount}개 슬라이드 저장 실패`, 'error');
    } else {
        showToast(t('editSaved'), 'success');
    }

    // 저장 완료 후 조회 모드로 전환
    _exitEditModeClean();
}

function _exitEditModeClean() {
    state.editMode = false;
    state.editSelectedObj = null;
    state.editDirtySlides.clear();
    $('#previewCanvas').removeClass('edit-mode');
    $('#btnEditToggle').removeClass('active');
    $('#btnEditSave').hide();
    $('#editTextToolbar').hide();
    $('#editBottomToolbar').hide();
    $('#editShapeMenu').hide();
    $('.slide-nav').show();
    renderSlideAtIndex(state.currentSlideIndex);
    renderSlideThumbnails();
    renderSlideThumbList();
}

// ============ 프레젠테이션 모드 ============
let presentationIndex = 0;
let _presActivePanel = 'A'; // 'A' 또는 'B' - 현재 활성 패널

function startPresentation() {
    if (state.generatedSlides.length === 0) {
        showToast(t('noSlides'), 'error');
        return;
    }
    presentationIndex = 0;
    _presActivePanel = 'A';
    $('#presentationSlideA').addClass('active');
    $('#presentationSlideB').removeClass('active');
    $('#presentationMode').show();
    renderPresentationSlide(0, 'A');

    $(document).on('keydown.presentation', function (e) {
        if (e.key === 'ArrowRight' || e.key === ' ') {
            e.preventDefault();
            if (presentationIndex < state.generatedSlides.length - 1) {
                presentationIndex++;
                transitionPresentationSlide(presentationIndex);
            }
        } else if (e.key === 'ArrowLeft') {
            e.preventDefault();
            if (presentationIndex > 0) {
                presentationIndex--;
                transitionPresentationSlide(presentationIndex);
            }
        } else if (e.key === 'Escape') {
            exitPresentation();
        }
    });
}

let _presTransitioning = false;

function transitionPresentationSlide(index) {
    if (_presTransitioning) return;
    _presTransitioning = true;

    // 다음 패널에 새 슬라이드를 미리 렌더링한 뒤 크로스페이드
    const nextPanel = _presActivePanel === 'A' ? 'B' : 'A';
    renderPresentationSlide(index, nextPanel);

    // 크로스페이드: 새 패널 활성화, 기존 패널 비활성화
    $(`#presentationSlide${nextPanel}`).addClass('active');
    $(`#presentationSlide${_presActivePanel}`).removeClass('active');
    _presActivePanel = nextPanel;

    setTimeout(() => { _presTransitioning = false; }, 350);
}

function renderPresentationSlide(index, panel) {
    const slide = state.generatedSlides[index];
    if (!slide) return;

    const container = $(`#presentationSlide${panel}`);
    container.find('.preview-obj').remove();

    if (slide.background_image) {
        $(`#presentationBg${panel}`).css({ 'background-image': `url(${slide.background_image})`, 'background': '' });
    } else {
        $(`#presentationBg${panel}`).css({ 'background-image': 'none', 'background': 'white' });
    }

    const containerW = container.width();
    const containerH = container.height();
    const scaleX = containerW / 960;
    const scaleY = containerH / 540;

    let presDescIdx = 0;
    let presSubIdx = 0;
    const presItems = slide.items || [];

    (slide.objects || []).forEach(obj => {
        const div = $('<div>').addClass('preview-obj').css({
            position: 'absolute',
            left: (obj.x * scaleX) + 'px',
            top: (obj.y * scaleY) + 'px',
            width: (obj.width * scaleX) + 'px',
            height: (obj.height * scaleY) + 'px',
            zIndex: 10,
        });

        if (obj.obj_type === 'image' && obj.image_url) {
            const imgFit = obj.image_fit || 'contain';
            div.append(`<img src="${obj.image_url}" style="width:100%;height:100%;object-fit:${imgFit};">`);
        } else if (obj.obj_type === 'text') {
            const style = obj.text_style || {};
            const text = obj.generated_text || obj.text_content || '';
            const scaledFontSize = (style.font_size || 16) * Math.min(scaleX, scaleY);
            const role = obj.role || obj._auto_role || '';
            div.css({
                fontFamily: style.font_family || 'Inter, Arial, sans-serif',
                fontSize: scaledFontSize + 'px',
                color: style.color || '#000',
                fontWeight: style.bold ? 'bold' : 'normal',
                fontStyle: style.italic ? 'italic' : 'normal',
                textAlign: style.align || 'left',
                padding: (8 * scaleX) + 'px',
                overflow: 'hidden',
                wordWrap: 'break-word',
                whiteSpace: 'pre-wrap',
            });

            // subtitle 역할: items[].heading 직접 매핑
            if (role === 'subtitle' && presItems.length > 0 && presSubIdx < presItems.length) {
                div.text(presItems[presSubIdx].heading || '');
                presSubIdx++;
            // description 역할: item.detail만 표시
            } else if (role === 'description' && presItems.length > 0 && presDescIdx < presItems.length) {
                const item = presItems[presDescIdx];
                div.text(item.detail || '');
                presDescIdx++;
                // 텍스트 넘침 시 높이 자동 확장
                div.css({ height: 'auto', minHeight: (obj.height * scaleY) + 'px', overflow: 'visible' });
            } else {
                div.text(text);
            }
        }

        container.append(div);
    });
}

function exitPresentation() {
    $('#presentationMode').hide();
    $(document).off('keydown.presentation');
}

// ============ 다운로드 & 공유 ============
function downloadPPTX() {
    if (!state.currentProject || state.generatedSlides.length === 0) {
        showToast(t('noSlides'), 'error');
        return;
    }
    window.open(apiUrl('/api/generate/' + state.currentProject._id + '/download/pptx'), '_blank');
}

function downloadPDF() {
    if (!state.currentProject || state.generatedSlides.length === 0) {
        showToast(t('noSlides'), 'error');
        return;
    }
    window.open(apiUrl('/api/generate/' + state.currentProject._id + '/download/pdf'), '_blank');
}

async function copyShareLink() {
    if (!state.currentProject) return;
    try {
        const res = await apiGet('/api/generate/' + state.currentProject._id + '/share-link');
        const shareUrl = window.location.origin + '/shared/' + res.share_token;
        await navigator.clipboard.writeText(shareUrl);
        showToast(t('msgShareCopied'), 'success');
    } catch (e) {
        showToast(e.message, 'error');
    }
}

// ============ 공유 프레젠테이션 ============
async function loadSharedPresentation(shareToken) {
    try {
        // 공유 페이지용 폰트 로딩 (인증 불필요)
        try {
            const fontRes = await fetch('/api/fonts/public');
            if (fontRes.ok) {
                const fontData = await fontRes.json();
                loadWebFonts(fontData.fonts || []);
            }
        } catch (e) { /* 폰트 로딩 실패 무시 */ }

        const res = await fetch('/api/shared/' + shareToken + '/slides');
        if (!res.ok) throw new Error('Not found');
        const data = await res.json();
        state.generatedSlides = data.slides || [];
        state.currentSlideIndex = 0;

        document.title = data.project_name + ' - PPTMaker';
        showApp();
        $('#sidebar').hide();
        $('#sidebarUserName').text(t('sharedPresentation'));

        // 바로 프레젠테이션 모드
        if (state.generatedSlides.length > 0) {
            startPresentation();
        }
    } catch (e) {
        alert('Shared presentation not found.');
    }
}

// ============ UI 헬퍼 ============
function autoResizeTextarea(el) {
    el.style.height = 'auto';
    el.style.height = Math.min(el.scrollHeight, 120) + 'px';
}

function escapeHtml(text) {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// ============ 슬라이드 드래그 & 드롭 순서 변경 ============
var _dragSlideIdx = null;

function initSlideDragDrop(containerSelector, itemSelector) {
    var container = $(containerSelector)[0];
    if (!container) return;

    container.addEventListener('dragstart', function (e) {
        var item = e.target.closest(itemSelector);
        if (!item) return;
        _dragSlideIdx = parseInt(item.dataset.slideIdx);
        item.classList.add('dragging');
        e.dataTransfer.effectAllowed = 'move';
        e.dataTransfer.setData('text/plain', _dragSlideIdx);
    });

    container.addEventListener('dragover', function (e) {
        e.preventDefault();
        e.dataTransfer.dropEffect = 'move';
        var item = e.target.closest(itemSelector);
        if (!item || parseInt(item.dataset.slideIdx) === _dragSlideIdx) return;

        // 드롭 위치 인디케이터
        $(containerSelector).find(itemSelector).removeClass('drag-over-top drag-over-bottom');
        var rect = item.getBoundingClientRect();
        var isVertical = containerSelector === '#slideThumbList';
        if (isVertical) {
            var midY = rect.top + rect.height / 2;
            if (e.clientY < midY) {
                item.classList.add('drag-over-top');
            } else {
                item.classList.add('drag-over-bottom');
            }
        } else {
            var midX = rect.left + rect.width / 2;
            if (e.clientX < midX) {
                item.classList.add('drag-over-top');
            } else {
                item.classList.add('drag-over-bottom');
            }
        }
    });

    container.addEventListener('dragleave', function (e) {
        var item = e.target.closest(itemSelector);
        if (item) {
            item.classList.remove('drag-over-top', 'drag-over-bottom');
        }
    });

    container.addEventListener('drop', function (e) {
        e.preventDefault();
        $(containerSelector).find(itemSelector).removeClass('drag-over-top drag-over-bottom dragging');
        var item = e.target.closest(itemSelector);
        if (!item || _dragSlideIdx === null) return;

        var toIdx = parseInt(item.dataset.slideIdx);
        if (toIdx === _dragSlideIdx) return;

        // 상/하 또는 좌/우 위치에 따라 삽입 위치 결정
        var rect = item.getBoundingClientRect();
        var isVertical = containerSelector === '#slideThumbList';
        var insertAfter;
        if (isVertical) {
            insertAfter = e.clientY >= rect.top + rect.height / 2;
        } else {
            insertAfter = e.clientX >= rect.left + rect.width / 2;
        }

        var finalTo = insertAfter ? toIdx : toIdx;
        if (_dragSlideIdx < toIdx && !insertAfter) finalTo = toIdx - 1;
        if (_dragSlideIdx > toIdx && insertAfter) finalTo = toIdx + 1;

        reorderSlides(_dragSlideIdx, finalTo);
        _dragSlideIdx = null;
    });

    container.addEventListener('dragend', function () {
        $(containerSelector).find(itemSelector).removeClass('dragging drag-over-top drag-over-bottom');
        _dragSlideIdx = null;
    });
}

function reorderSlides(fromIndex, toIndex) {
    if (fromIndex === toIndex) return;
    if (fromIndex < 0 || fromIndex >= state.generatedSlides.length) return;
    if (toIndex < 0 || toIndex >= state.generatedSlides.length) return;

    // 현재 보고 있는 슬라이드 추적
    var currentSlideId = state.generatedSlides[state.currentSlideIndex]
        ? state.generatedSlides[state.currentSlideIndex]._id
        : null;

    // 배열에서 이동
    var moved = state.generatedSlides.splice(fromIndex, 1)[0];
    state.generatedSlides.splice(toIndex, 0, moved);

    // order 필드 갱신
    state.generatedSlides.forEach(function (s, i) {
        s.order = i + 1;
    });

    // 현재 슬라이드 인덱스 복원
    if (currentSlideId) {
        for (var i = 0; i < state.generatedSlides.length; i++) {
            if (state.generatedSlides[i]._id === currentSlideId) {
                state.currentSlideIndex = i;
                break;
            }
        }
    }

    // UI 전체 갱신
    renderSlideThumbList();
    renderSlideThumbnails();
    renderSlideTextPanel();
    if ($('#gridOverlay').is(':visible')) {
        renderGridView();
    }
    updateSlideNav();

    // 백엔드 저장
    saveSlideOrder();
}

function saveSlideOrder() {
    var slideIds = state.generatedSlides.map(function (s) { return s._id; });
    apiPut('/api/generate/' + state.currentProject._id + '/reorder', {
        slide_ids: slideIds,
    }).catch(function (e) {
        showToast('순서 저장 실패: ' + e.message, 'error');
    });
}

function showToast(message, type) {
    const toast = $(`<div class="toast ${type || ''}">${escapeHtml(message)}</div>`);
    $('#toastContainer').append(toast);
    setTimeout(() => toast.remove(), 3500);
}

function closeModal(id) {
    $('#' + id).hide();
}

function showLoading(text) {
    $('#loadingText').text(text || t('msgProcessing'));
    $('#loadingOverlay').show();
}

function hideLoading() {
    $('#loadingOverlay').hide();
}
