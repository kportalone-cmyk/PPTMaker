/**
 * OfficeMaker - Kimi K2 Slides Style Frontend
 */

// ============ 솔루션명 ============
const _SOLUTION = window.__SOLUTION_NAME__ || 'OfficeMaker';

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
    manualMode: false,
    manualTemplateSlides: [],
    editMode: false,
    editSelectedObj: null,
    editDragging: false,
    editResizing: false,
    editDragOffset: { x: 0, y: 0 },
    editResizeDir: '',
    editResizeStart: null,
    editDirtySlides: new Set(),
    editingSlideId: null,  // 현재 편집 중인 슬라이드 ID (Lock 추적용)
    // 협업 상태
    collaborators: [],
    sharedProjects: [],
    isSharedView: false,
    activeLocks: {},
    onlineUsers: [],
    isCollabProject: false,
    collabRole: null,
    pollInterval: null,
    heartbeatInterval: null,
    lastSlideTimestamps: {},
    editAutoSaveInterval: null,
    projectPage: 0,         // 프로젝트 목록 현재 페이지 (0-based)
    projectPageSize: 10,    // 페이지당 프로젝트 수
    // 엑셀 상태
    selectedProjectType: 'slide',
    generatedExcel: null,
    univerAPI: null,
    // OnlyOffice 상태
    onlyofficeDoc: null,
    onlyofficeEditor: null,
    generatedDocx: null,
    // DOCX 양식 템플릿
    docxTemplateId: null,
};

let _animationCancelled = false;
let _isAnimating = false;
let _isGenerating = false;
let _abortController = null;
let _streamReader = null;

// ============ 슬라이드 사이즈 헬퍼 ============
function getSlideCanvasSize(slideSize) {
    const sizes = {
        '16:9': { w: 960, h: 540 },
        '4:3':  { w: 960, h: 720 },
        'A4':   { w: 960, h: 679 },
    };
    return sizes[slideSize] || sizes['16:9'];
}

function getCurrentSlideSize() {
    // Try to get from selected template or current project's template
    if (state._templateSlideSize) return getSlideCanvasSize(state._templateSlideSize);
    return { w: 960, h: 540 };
}

function getThumbDimensions(baseWidth, slideSize) {
    const sz = slideSize ? getSlideCanvasSize(slideSize) : getCurrentSlideSize();
    const ratio = sz.h / sz.w;
    return { w: baseWidth, h: Math.round(baseWidth * ratio) };
}

function updateSlideCanvasAspect() {
    const sz = getCurrentSlideSize();
    const ratio = sz.w + ' / ' + sz.h;
    $('#previewCanvas').css({ 'aspect-ratio': ratio, 'max-width': sz.w + 'px' });
    // 좌측 썸네일 aspect-ratio 업데이트
    $('.slide-thumb-v-inner').css('aspect-ratio', ratio);
    $('.grid-thumb-inner').css('aspect-ratio', ratio);
    // 템플릿 피커 썸네일
    $('.template-picker-thumb').css('aspect-ratio', ratio);
    $('.template-slide-thumb').css('aspect-ratio', ratio);
}

// 탭 전환 시 애니메이션 즉시 완료 처리
document.addEventListener('visibilitychange', () => {
    if (!document.hidden && _isAnimating && !_animationCancelled) {
        _animationCancelled = true;
    }
});

// ============ 다국어 사전 ============
const I18N = {
    ko: {
        appTitle: (window.__SOLUTION_NAME__ || 'OfficeMaker'),
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
        addUrl: 'URL',
        urlModalTitle: 'URL 추가',
        urlLabel: 'URL 목록',
        urlPlaceholder: 'URL을 한 줄에 하나씩 입력하세요',
        urlHint: 'YouTube URL은 자동으로 자막을 추출합니다',
        msgUrlAdded: 'URL 리소스가 추가되었습니다',
        msgUrlCollecting: 'URL 콘텐츠 수집 중...',
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
        noSlides: '지침을 입력하고 프레젠테이션을 생성하세요',
        emptyStep1: '리소스를 추가하세요',
        emptyStep2: '지침을 입력하세요',
        emptyStep3: '생성 버튼을 누르세요',
        addSlideBtn: '슬라이드 추가',
        addSlide: '슬라이드 추가',
        selectTemplateSlide: '슬라이드 레이아웃 선택',
        generateSlideText: 'AI 생성',
        slideInstructionPlaceholder: '이 슬라이드에 대한 지침을 입력하세요... (예: 제목을 바꿔줘, 항목을 추가해줘)',
        msgSlideAdded: '슬라이드가 추가되었습니다',
        msgSlideDeleted: '슬라이드가 삭제되었습니다',
        msgSlideTextGenerated: '텍스트가 생성되었습니다',
        msgSelectTemplate: '먼저 템플릿을 선택하세요',
        msgGeneratingText: '텍스트 생성 중...',
        aiModifyRequest: 'AI 수정 요청',
        editingNow: '편집 중',
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
        statusGenerating: '생성 중',
        statusGenerated: '생성 완료',
        statusStopped: '중단됨',
        noDesc: '설명 없음',
        projectType: '프로젝트 유형',
        typeSlide: '슬라이드',
        typeExcel: '엑셀',
        generateExcel: '엑셀 생성',
        modifyExcel: '엑셀 수정',
        excelModifyPlaceholder: '수정할 내용을 입력하세요 (예: 3월 데이터 삭제, 새 열 추가, 차트 타입 변경)',
        excelUploading: '엑셀 파일 업로드 중...',
        excelUploaded: '엑셀 파일이 업로드되었습니다',
        downloadXlsx: 'XLSX 다운로드',
        excelGenerating: 'AI가 데이터를 구조화하고 있습니다...',
        excelSearching: '인터넷에서 자료를 검색하고 있습니다...',
        excelSearchDone: '검색 완료! 데이터를 생성합니다...',
        excelPreparing: '리소스를 분석하고 있습니다...',
        excelStreaming: '데이터를 생성하고 있습니다...',
        excelStreamingDetail: '{sheets}개 시트 · {rows}개 행 수집',
        excelFinalizing: '데이터를 정리하고 있습니다...',
        msgExcelGenerated: '엑셀 데이터가 생성되었습니다',
        msgNoExcelData: '생성된 엑셀 데이터가 없습니다',
        excelCharts: '차트',
        typeWord: '워드',
        generateWord: '문서 생성',
        wordModifyPlaceholder: '수정할 내용을 입력하세요 (예: 섹션 추가, 내용 보강, 형식 변경)',
        wordGenerating: 'AI가 문서를 작성하고 있습니다...',
        wordSearching: '인터넷에서 자료를 검색하고 있습니다...',
        wordSearchDone: '검색 완료! 문서를 작성합니다...',
        wordPreparing: '리소스를 분석하고 있습니다...',
        wordStreaming: '문서를 작성하고 있습니다...',
        wordFinalizing: '문서를 정리하고 있습니다...',
        msgWordGenerated: '문서가 생성되었습니다',
        msgNoWordData: '생성된 문서가 없습니다',
        downloadDocx: 'DOCX 다운로드',
        resourceHint: '리소스를 추가하면 참고하여 생성하고, 없으면 웹 검색으로 자료를 수집합니다.',
        msgNeedInstructions: '어떤 내용을 어떻게 작성할지 지침을 입력해주세요',
        msgEnterInstructions: '지침을 입력하세요 (리소스가 없으면 인터넷 검색으로 자료를 수집합니다)',
        typeFile: '파일',
        typeText: '텍스트',
        typeWeb: '웹 검색',
        welcomeTitle: '무엇을 만들어볼까요?',
        welcomeDesc: '프로젝트를 선택하거나 새로 만들어 AI 문서 작업을 시작하세요',
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
        msgConfirmDeleteAllRes: '모든 리소스를 삭제하시겠습니까?',
        msgAllResDeleted: '모든 리소스가 삭제되었습니다',
        deleteAll: '전체 삭제',
        searching: '검색 중...',
        stopBtn: '중단',
        restartBtn: '재생성',
        msgStopped: '생성이 중단되었습니다',
        msgResetting: '초기화 중...',
        msgResetDone: '프로젝트가 초기화되었습니다',
        adminPage: '관리자 페이지',
        authChecking: '접속정보 확인중입니다',
        authCheckingSub: '잠시만 기다려주세요...',
        authErrorTitle: '인증정보가 올바르지 않습니다',
        authErrorDesc: '접속 링크가 만료되었거나 잘못된 인증정보입니다.\n관리자에게 문의하거나 다시 시도해주세요.',
        translateBtn: '언어 전환',
        translateTitle: '언어 전환',
        translateDesc: '현재 슬라이드를 선택한 언어로 번역합니다.<br>원본 프로젝트는 그대로 보존됩니다.',
        translateNoSlides: '생성된 슬라이드가 없습니다',
        translateInProgress: '다른 작업이 진행 중입니다',
        translatingTo: '{lang}로 번역 중입니다...',
        translatingProgress: '{lang}로 번역 중... ({current}/{total})',
        translatingCanvas: '{lang}로 번역 중',
        translatingCanvasSub: '슬라이드 텍스트를 변환하는 중',
        translateSlideComplete: '슬라이드 {n} 번역 완료',
        translateComplete: '{lang} 번역 완료 ({total}개 슬라이드)',
        translateDone: '{lang}로 번역 완료! 새 프로젝트: {name}',
        translateError: '번역 중 오류가 발생했습니다.',
        translateFailed: '번역 실패',
    },
    en: {
        appTitle: (window.__SOLUTION_NAME__ || 'OfficeMaker'),
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
        addUrl: 'URL',
        urlModalTitle: 'Add URLs',
        urlLabel: 'URL List',
        urlPlaceholder: 'Enter one URL per line',
        urlHint: 'YouTube URLs will automatically extract subtitles',
        msgUrlAdded: 'URL resources added',
        msgUrlCollecting: 'Collecting URL content...',
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
        noSlides: 'Enter instructions and generate your presentation',
        emptyStep1: 'Add resources',
        emptyStep2: 'Enter instructions',
        emptyStep3: 'Click Generate',
        addSlideBtn: 'Add Slide',
        addSlide: 'Add Slide',
        selectTemplateSlide: 'Select Slide Layout',
        generateSlideText: 'AI Generate',
        slideInstructionPlaceholder: 'Enter instructions for this slide... (e.g., change the title, add an item)',
        msgSlideAdded: 'Slide added',
        msgSlideDeleted: 'Slide deleted',
        msgSlideTextGenerated: 'Text generated',
        msgSelectTemplate: 'Please select a template first',
        msgGeneratingText: 'Generating text...',
        aiModifyRequest: 'AI Modify',
        editingNow: 'Editing',
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
        statusGenerating: 'Generating',
        statusGenerated: 'Generated',
        statusStopped: 'Stopped',
        noDesc: 'No description',
        projectType: 'Project Type',
        typeSlide: 'Slides',
        typeExcel: 'Excel',
        generateExcel: 'Generate Excel',
        modifyExcel: 'Modify Excel',
        excelModifyPlaceholder: 'Enter modification instructions (e.g., delete March data, add new column, change chart type)',
        excelUploading: 'Uploading Excel file...',
        excelUploaded: 'Excel file uploaded successfully',
        downloadXlsx: 'Download XLSX',
        excelGenerating: 'AI is structuring data...',
        excelSearching: 'Searching the internet for data...',
        excelSearchDone: 'Search complete! Generating data...',
        excelPreparing: 'Analyzing resources...',
        excelStreaming: 'Generating data...',
        excelStreamingDetail: '{sheets} sheets · {rows} rows collected',
        excelFinalizing: 'Finalizing data...',
        msgExcelGenerated: 'Excel data has been generated',
        msgNoExcelData: 'No generated Excel data',
        excelCharts: 'Charts',
        typeWord: 'Word',
        generateWord: 'Generate Word',
        wordModifyPlaceholder: 'Enter modification instructions (e.g., add section, enhance content, change format)',
        wordGenerating: 'AI is writing the document...',
        wordSearching: 'Searching the internet for data...',
        wordSearchDone: 'Search complete! Writing document...',
        wordPreparing: 'Analyzing resources...',
        wordStreaming: 'Writing document...',
        wordFinalizing: 'Finalizing document...',
        msgWordGenerated: 'Document has been generated',
        msgNoWordData: 'No generated document data',
        downloadDocx: 'Download DOCX',
        resourceHint: 'Add resources to use as reference, or leave empty to generate from web search.',
        msgNeedInstructions: 'Please enter instructions on what content to create and how',
        msgEnterInstructions: 'Enter instructions (internet search will be used if no resources)',
        typeFile: 'File',
        typeText: 'Text',
        typeWeb: 'Web',
        welcomeTitle: 'What would you like to create?',
        welcomeDesc: 'Select a project or create a new one to start AI document creation',
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
        msgConfirmDeleteAllRes: 'Delete all resources?',
        msgAllResDeleted: 'All resources deleted',
        deleteAll: 'Delete All',
        searching: 'Searching...',
        stopBtn: 'Stop',
        restartBtn: 'Regenerate',
        msgStopped: 'Generation stopped',
        msgResetting: 'Resetting...',
        msgResetDone: 'Project has been reset',
        adminPage: 'Admin',
        authChecking: 'Verifying access...',
        authCheckingSub: 'Please wait a moment...',
        authErrorTitle: 'Invalid authentication',
        authErrorDesc: 'The access link has expired or is invalid.\nPlease contact your administrator or try again.',
        translateBtn: 'Translate',
        translateTitle: 'Translate Slides',
        translateDesc: 'Translate current slides to a selected language.<br>The original project will be preserved.',
        translateNoSlides: 'No generated slides',
        translateInProgress: 'Another operation is in progress',
        translatingTo: 'Translating to {lang}...',
        translatingProgress: 'Translating to {lang}... ({current}/{total})',
        translatingCanvas: 'Translating to {lang}',
        translatingCanvasSub: 'Converting slide text',
        translateSlideComplete: 'Slide {n} translated',
        translateComplete: '{lang} translation complete ({total} slides)',
        translateDone: 'Translated to {lang}! New project: {name}',
        translateError: 'An error occurred during translation.',
        translateFailed: 'Translation failed',
    },
    ja: {
        appTitle: (window.__SOLUTION_NAME__ || 'OfficeMaker'),
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
        addUrl: 'URL',
        urlModalTitle: 'URL追加',
        urlLabel: 'URLリスト',
        urlPlaceholder: '1行に1つのURLを入力してください',
        urlHint: 'YouTube URLは自動的に字幕を抽出します',
        msgUrlAdded: 'URLリソースが追加されました',
        msgUrlCollecting: 'URLコンテンツ収集中...',
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
        noSlides: '指示を入力してプレゼンテーションを生成',
        emptyStep1: 'リソースを追加',
        emptyStep2: '指示を入力',
        emptyStep3: '生成ボタンを押す',
        addSlideBtn: 'スライド追加',
        addSlide: 'スライド追加',
        selectTemplateSlide: 'スライドレイアウト選択',
        generateSlideText: 'AI生成',
        slideInstructionPlaceholder: 'このスライドの指示を入力... (例: タイトルを変更、項目を追加)',
        msgSlideAdded: 'スライドを追加しました',
        msgSlideDeleted: 'スライドを削除しました',
        msgSlideTextGenerated: 'テキストを生成しました',
        msgSelectTemplate: 'テンプレートを選択してください',
        msgGeneratingText: 'テキスト生成中...',
        aiModifyRequest: 'AI修正',
        editingNow: '編集中',
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
        statusGenerating: '生成中',
        statusGenerated: '完了',
        statusStopped: '中断',
        noDesc: '説明なし',
        projectType: 'プロジェクトタイプ',
        typeSlide: 'スライド',
        typeExcel: 'エクセル',
        generateExcel: 'エクセル生成',
        modifyExcel: 'エクセル修正',
        excelModifyPlaceholder: '修正内容を入力してください (例: 3月データ削除、新しい列追加、チャートタイプ変更)',
        excelUploading: 'エクセルファイルをアップロード中...',
        excelUploaded: 'エクセルファイルがアップロードされました',
        downloadXlsx: 'XLSXダウンロード',
        excelGenerating: 'AIがデータを構造化しています...',
        excelSearching: 'インターネットで資料を検索中...',
        excelSearchDone: '検索完了！データを生成中...',
        excelPreparing: 'リソースを分析中...',
        excelStreaming: 'データを生成中...',
        excelStreamingDetail: '{sheets}シート · {rows}行取得',
        excelFinalizing: 'データを整理中...',
        msgExcelGenerated: 'エクセルデータが生成されました',
        msgNoExcelData: '生成されたデータがありません',
        excelCharts: 'チャート',
        typeWord: 'ワード',
        generateWord: 'ドキュメント生成',
        downloadDocx: 'DOCXダウンロード',
        resourceHint: 'リソースを追加すると参考にして生成し、なければウェブ検索で資料を収集します。',
        msgNeedInstructions: 'どのような内容をどのように作成するか指示を入力してください',
        msgEnterInstructions: '指示を入力してください（リソースがない場合はインターネット検索を使用）',
        typeFile: 'ファイル',
        typeText: 'テキスト',
        typeWeb: 'ウェブ',
        welcomeTitle: '何を作りますか？',
        welcomeDesc: 'プロジェクトを選択または新規作成してAIドキュメント作成を開始',
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
        msgConfirmDeleteAllRes: '全リソースを削除しますか？',
        msgAllResDeleted: '全リソースを削除しました',
        deleteAll: '全削除',
        searching: '検索中...',
        stopBtn: '中断',
        restartBtn: '再生成',
        msgStopped: '生成が中断されました',
        msgResetting: '初期化中...',
        msgResetDone: 'プロジェクトが初期化されました',
        adminPage: '管理者ページ',
        authChecking: 'アクセス情報を確認中です',
        authCheckingSub: 'しばらくお待ちください...',
        authErrorTitle: '認証情報が正しくありません',
        authErrorDesc: 'アクセスリンクの有効期限が切れているか、無効です。\n管理者にお問い合わせください。',
        translateBtn: '言語変換',
        translateTitle: '言語変換',
        translateDesc: '現在のスライドを選択した言語に翻訳します。<br>元のプロジェクトはそのまま保存されます。',
        translateNoSlides: '生成されたスライドがありません',
        translateInProgress: '他の作業が進行中です',
        translatingTo: '{lang}に翻訳中...',
        translatingProgress: '{lang}に翻訳中... ({current}/{total})',
        translatingCanvas: '{lang}に翻訳中',
        translatingCanvasSub: 'スライドテキストを変換中',
        translateSlideComplete: 'スライド {n} 翻訳完了',
        translateComplete: '{lang} 翻訳完了 ({total}スライド)',
        translateDone: '{lang}に翻訳完了！新規プロジェクト: {name}',
        translateError: '翻訳中にエラーが発生しました。',
        translateFailed: '翻訳失敗',
    },
    zh: {
        appTitle: (window.__SOLUTION_NAME__ || 'OfficeMaker'),
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
        addUrl: 'URL',
        urlModalTitle: '添加URL',
        urlLabel: 'URL列表',
        urlPlaceholder: '每行输入一个URL',
        urlHint: 'YouTube URL会自动提取字幕',
        msgUrlAdded: 'URL资源已添加',
        msgUrlCollecting: '正在收集URL内容...',
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
        noSlides: '输入指令并生成演示文稿',
        emptyStep1: '添加资源',
        emptyStep2: '输入指令',
        emptyStep3: '点击生成',
        addSlideBtn: '添加幻灯片',
        addSlide: '添加幻灯片',
        selectTemplateSlide: '选择幻灯片布局',
        generateSlideText: 'AI生成',
        slideInstructionPlaceholder: '输入此幻灯片的指令... (例如: 修改标题, 添加项目)',
        msgSlideAdded: '幻灯片已添加',
        msgSlideDeleted: '幻灯片已删除',
        msgSlideTextGenerated: '文本已生成',
        msgSelectTemplate: '请先选择模板',
        msgGeneratingText: '正在生成文本...',
        aiModifyRequest: 'AI修改',
        editingNow: '编辑中',
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
        statusGenerating: '生成中',
        statusGenerated: '已完成',
        statusStopped: '已停止',
        noDesc: '无描述',
        projectType: '项目类型',
        typeSlide: '幻灯片',
        typeExcel: '电子表格',
        generateExcel: '生成表格',
        modifyExcel: '修改表格',
        excelModifyPlaceholder: '请输入修改内容 (例如: 删除3月数据, 添加新列, 更改图表类型)',
        excelUploading: '正在上传Excel文件...',
        excelUploaded: 'Excel文件上传成功',
        downloadXlsx: '下载XLSX',
        excelGenerating: 'AI正在整理数据...',
        excelSearching: '正在从互联网搜索资料...',
        excelSearchDone: '搜索完成！正在生成数据...',
        excelPreparing: '正在分析资源...',
        excelStreaming: '正在生成数据...',
        excelStreamingDetail: '{sheets}个工作表 · {rows}行已收集',
        excelFinalizing: '正在整理数据...',
        msgExcelGenerated: '表格数据已生成',
        msgNoExcelData: '没有生成的数据',
        excelCharts: '图表',
        typeWord: 'Word',
        generateWord: '生成文档',
        downloadDocx: '下载DOCX',
        resourceHint: '添加资源将作为参考生成，未添加则通过网络搜索收集资料。',
        msgNeedInstructions: '请输入要创建什么内容以及如何创建的指示',
        msgEnterInstructions: '请输入指令（如无资源将通过网络搜索收集资料）',
        typeFile: '文件',
        typeText: '文本',
        typeWeb: '网页',
        welcomeTitle: '想要创建什么？',
        welcomeDesc: '选择项目或新建以开始AI文档创作',
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
        msgConfirmDeleteAllRes: '确定删除所有资源？',
        msgAllResDeleted: '所有资源已删除',
        deleteAll: '全部删除',
        searching: '搜索中...',
        stopBtn: '停止',
        restartBtn: '重新生成',
        msgStopped: '生成已停止',
        msgResetting: '正在初始化...',
        msgResetDone: '项目已初始化',
        adminPage: '管理后台',
        authChecking: '正在验证访问信息',
        authCheckingSub: '请稍候...',
        authErrorTitle: '认证信息无效',
        authErrorDesc: '访问链接已过期或无效。\n请联系管理员或重试。',
        translateBtn: '语言转换',
        translateTitle: '语言转换',
        translateDesc: '将当前幻灯片翻译为所选语言。<br>原项目将保留不变。',
        translateNoSlides: '没有已生成的幻灯片',
        translateInProgress: '其他操作正在进行中',
        translatingTo: '正在翻译为{lang}...',
        translatingProgress: '正在翻译为{lang}... ({current}/{total})',
        translatingCanvas: '正在翻译为{lang}',
        translatingCanvasSub: '正在转换幻灯片文本',
        translateSlideComplete: '幻灯片 {n} 翻译完成',
        translateComplete: '{lang} 翻译完成 ({total}个幻灯片)',
        translateDone: '已翻译为{lang}！新项目: {name}',
        translateError: '翻译过程中出现错误。',
        translateFailed: '翻译失败',
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
    $('.i18n-translateBtn').text(t('translateBtn'));
    // 버튼 상태에 따라 텍스트 변경
    if (_isGenerating) {
        // 생성 중이면 중단 버튼 유지
    } else if (state.generatedSlides.length > 0) {
        $('.i18n-generateBtn').text(t('restartBtn'));
    } else {
        $('.i18n-generateBtn').text(t('generateBtn'));
    }
    $('.i18n-slideCountAuto').text(t('slideCountAuto'));
    $('.i18n-addSlideBtn').text(t('addSlideBtn'));
    $('#instructionsInput').attr('placeholder', t('instructionsPlaceholder'));
    $('#slideInstructionInput').attr('placeholder', t('slideInstructionPlaceholder'));

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

    // 리소스/엑셀
    $('.i18n-resourceHint').text(t('resourceHint'));
    $('.i18n-excelCharts').text(t('excelCharts'));
    $('.i18n-addUrl').text(t('addUrl'));

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
    // 솔루션명 동적 반영
    document.title = _SOLUTION;
    $('.app-title').text(_SOLUTION);
    $('.app-title-sidebar').text(_SOLUTION);

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

function goToMain() {
    state.currentProject = null;
    showEmptyState();
    loadProjects();
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
    if (retries === undefined) retries = 1;
    const link = document.createElement('link');
    link.id = linkId;
    link.rel = 'stylesheet';
    link.href = url;
    link.onerror = function() {
        console.warn('Font CSS load failed:', family, '(retries left:', retries, ')');
        link.remove();
        if (retries > 0) {
            setTimeout(function() { _loadFontCSS(linkId, url, family, retries - 1); }, 2000);
        }
    };
    document.head.appendChild(link);
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
    // 현재 활성 탭에 따라 렌더링
    const activeTab = $('.home-tab.active').data('tab') || 'recent';
    switchHomeTab(activeTab);
}

function switchHomeTab(tabName) {
    // 탭 버튼 활성화
    $('.home-tab').removeClass('active');
    $(`.home-tab[data-tab="${tabName}"]`).addClass('active');

    // 탭 카운트 배지 업데이트
    _updateHomeTabCounts();

    if (tabName === 'recent') {
        renderRecentProjects();
    } else if (tabName === 'shared') {
        renderSharedProjects();
    } else if (tabName === 'myshared') {
        renderMySharedProjects();
    }
}

function _updateHomeTabCounts() {
    const sharedCount = state.sharedProjects.length;
    if (sharedCount > 0) {
        $('#sharedTabCount').text(sharedCount).show();
    } else {
        $('#sharedTabCount').hide();
    }
    // 내가 공유한 파일: 협업자가 1명 이상인 내 프로젝트
    const mySharedCount = state.projects.filter(p => (p._collab_count || 0) > 0).length;
    if (mySharedCount > 0) {
        $('#mySharedTabCount').text(mySharedCount).show();
    } else {
        $('#mySharedTabCount').hide();
    }
}

function renderRecentProjects() {
    const grid = $('#recentProjectsGrid');
    grid.empty();
    // 내 프로젝트 + 공유 프로젝트를 합쳐서 최근순 9개
    const all = [
        ...state.projects.map(p => ({ ...p, _isShared: false })),
        ...state.sharedProjects.map(p => ({ ...p, _isShared: true })),
    ].sort((a, b) => new Date(b.created_at || 0) - new Date(a.created_at || 0)).slice(0, 9);

    if (all.length === 0) {
        grid.html(`
            <div class="home-empty-tab" style="grid-column: 1 / -1;">
                <svg width="48" height="48" fill="none" stroke="currentColor" stroke-width="1.5"><rect x="8" y="8" width="32" height="32" rx="6"/><path d="M18 20h12M18 26h8M18 32h10"/></svg>
                <p>${t('noRecentProjects', '아직 프로젝트가 없습니다')}<br><span style="font-size:12px;opacity:0.7;">${t('createNewProject', '새 프로젝트를 만들어보세요')}</span></p>
            </div>
        `);
        return;
    }

    all.forEach(p => _appendProjectCard(grid, p));
}

function renderSharedProjects() {
    const grid = $('#recentProjectsGrid');
    grid.empty();

    const shared = state.sharedProjects
        .map(p => ({ ...p, _isShared: true }))
        .sort((a, b) => new Date(b.created_at || 0) - new Date(a.created_at || 0));

    if (shared.length === 0) {
        grid.html(`
            <div class="home-empty-tab" style="grid-column: 1 / -1;">
                <svg width="48" height="48" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M16 40V24a8 8 0 0116 0v16"/><circle cx="16" cy="42" r="3"/><circle cx="32" cy="42" r="3"/><circle cx="24" cy="18" r="3"/></svg>
                <p>${t('noSharedProjects', '공유받은 파일이 없습니다')}<br><span style="font-size:12px;opacity:0.7;">${t('sharedProjectsDesc', '다른 사용자가 공유한 프로젝트가 여기에 표시됩니다')}</span></p>
            </div>
        `);
        return;
    }

    shared.forEach(p => _appendProjectCard(grid, p));
}

function renderMySharedProjects() {
    const grid = $('#recentProjectsGrid');
    grid.empty();

    const myShared = state.projects
        .filter(p => (p._collab_count || 0) > 0)
        .map(p => ({ ...p, _isShared: false, _isMyShared: true }))
        .sort((a, b) => new Date(b.created_at || 0) - new Date(a.created_at || 0));

    if (myShared.length === 0) {
        grid.html(`
            <div class="home-empty-tab" style="grid-column: 1 / -1;">
                <svg width="48" height="48" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M26 10l10 10-10 10"/><path d="M36 20H17a9 9 0 00-9 9v4"/></svg>
                <p>${t('noMySharedProjects', '공유한 파일이 없습니다')}<br><span style="font-size:12px;opacity:0.7;">${t('mySharedProjectsDesc', '프로젝트에 협업자를 추가하면 여기에 표시됩니다')}</span></p>
            </div>
        `);
        return;
    }

    myShared.forEach(p => _appendProjectCard(grid, p));
}

function _appendProjectCard(grid, p) {
    const statusLabel = { draft: t('statusDraft'), preparing: t('statusPreparing'), generating: t('statusGenerating'), generated: t('statusGenerated'), stop_requested: t('statusStopped'), stopped: t('statusStopped') }[p.status] || t('statusDraft');
    const date = p.created_at ? new Date(p.created_at).toLocaleDateString() : '';
    const collabCount = p._collab_count || 0;
    const collabIndicator = collabCount > 0
        ? `<span class="rc-collab" onclick="event.stopPropagation();openProject('${p._id}');setTimeout(showCollabModal,300)" title="${t('collaboration','협업')} (${collabCount})"><svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="4.5" cy="4.5" r="2.2"/><circle cx="9" cy="5.5" r="1.8"/><path d="M.5 10.5c0-2 1.5-3.5 4-3.5s2.5.8 3 1.5"/></svg> ${collabCount}</span>`
        : '';
    const sharedBadge = p._isShared
        ? `<span class="rc-shared-role">${p._collab_role === 'editor' ? t('editor','편집자') : t('viewer','뷰어')}</span>`
        : '';
    const typeBadge = _getProjectTypeBadge(p.project_type);
    const deleteBtn = p._isShared ? '' : `<button class="rc-delete-btn" onclick="deleteProjectById('${p._id}', event)" title="${t('delete','삭제')}"><svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M3 4h8M5 4V3a1 1 0 011-1h2a1 1 0 011 1v1M10 4v7a1 1 0 01-1 1H5a1 1 0 01-1-1V4"/></svg></button>`;
    // 공유 프로젝트: 소유자 정보 / 내가 공유한 프로젝트: 협업자 수 표시
    let subInfo = '';
    if (p._isShared && p._owner_name) {
        subInfo = `<div class="rc-shared-by"><svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="1.5"><circle cx="6" cy="4" r="2.5"/><path d="M1.5 12c0-2.5 2-4.5 4.5-4.5s4.5 2 4.5 4.5"/></svg>${escapeHtml(p._owner_name)}</div>`;
    } else if (p._isMyShared && collabCount > 0) {
        subInfo = `<div class="rc-shared-by"><svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M7 7.5a2.5 2.5 0 10-5 0"/><circle cx="4.5" cy="3" r="2"/><path d="M10 8.5a2 2 0 10-4 0"/><circle cx="8" cy="4" r="1.5"/></svg>${collabCount}${t('people', '명')}${t('sharedWith', '에게 공유중')}</div>`;
    }
    const typeClass = _getProjectTypeClass(p.project_type);
    grid.append(`
        <div class="recent-card ${typeClass}" onclick="openProject('${p._id}')">
            ${deleteBtn}
            <div class="rc-title">${typeBadge}${escapeHtml(p.name)}${collabIndicator}</div>
            <div class="rc-desc">${escapeHtml(p.description || t('noDesc'))}</div>
            ${subInfo}
            <div class="rc-meta">
                <span class="rc-date">${date}${sharedBadge}</span>
                <span class="rc-status ${p.status || 'draft'}">${statusLabel}</span>
            </div>
        </div>
    `);
}

// ============ 프로젝트 관리 ============
async function loadProjects() {
    try {
        const res = await apiGet('/api/projects');
        state.projects = res.projects || [];
        state.sharedProjects = res.shared_projects || [];
        renderProjectList();
        _updateHomeTabCounts();
    } catch (e) {
        showToast(t('msgLoadingProject'), 'error');
    }
}

function renderProjectList(skipAutoNav) {
    const list = $('#projectList');
    list.empty();

    // 내 프로젝트 + 공유 프로젝트를 하나로 합침
    const allProjects = [
        ...state.projects.map(p => ({ ...p, _isShared: false })),
        ...state.sharedProjects.map(p => ({ ...p, _isShared: true })),
    ];

    if (allProjects.length === 0) {
        list.html(`<div style="padding:16px 12px;text-align:center;color:var(--sidebar-text);font-size:12px;opacity:0.6;">${t('noResources')}</div>`);
        return;
    }

    const pageSize = state.projectPageSize;
    const totalPages = Math.ceil(allProjects.length / pageSize);

    // 페이지 범위 보정
    if (state.projectPage >= totalPages) state.projectPage = totalPages - 1;
    if (state.projectPage < 0) state.projectPage = 0;

    const start = state.projectPage * pageSize;
    const pageItems = allProjects.slice(start, start + pageSize);

    // 현재 프로젝트가 포함된 페이지로 이동 (새 프로젝트 열 때, 수동 페이지 이동 시 제외)
    if (!skipAutoNav && state.currentProject) {
        const globalIdx = allProjects.findIndex(p => p._id === state.currentProject._id);
        if (globalIdx >= 0) {
            const targetPage = Math.floor(globalIdx / pageSize);
            if (targetPage !== state.projectPage) {
                state.projectPage = targetPage;
                renderProjectList();
                return;
            }
        }
    }

    let lastWasShared = false;

    pageItems.forEach(p => {
        // 공유 프로젝트 섹션 라벨 (첫 번째 공유 프로젝트 앞에 한 번만)
        if (p._isShared && !lastWasShared) {
            list.append(`<div class="sidebar-section-label" style="margin-top:12px;padding:0 12px;font-size:10px;text-transform:uppercase;letter-spacing:0.05em;color:var(--sidebar-text);opacity:0.5;">공유됨</div>`);
            lastWasShared = true;
        }

        const isActive = state.currentProject && state.currentProject._id === p._id;
        const date = p.created_at ? new Date(p.created_at).toLocaleDateString() : '';

        const typeClass = _getProjectTypeClass(p.project_type);
        const projTypeIcon = _getProjectTypeIcon(p.project_type);

        if (p._isShared) {
            const roleLabel = p._collab_role === 'editor' ? '편집자' : '뷰어';
            list.append(`
                <div class="project-item ${isActive ? 'active' : ''}" onclick="openProject('${p._id}')">
                    <div class="proj-icon ${typeClass}">
                        ${projTypeIcon}
                    </div>
                    <div class="proj-info">
                        <div class="proj-name">${escapeHtml(p.name)}</div>
                        <div class="proj-date">${date} · ${roleLabel}</div>
                    </div>
                    <div class="proj-status-dot ${p.status || 'draft'}"></div>
                </div>
            `);
        } else {
            const collabCount = p._collab_count || 0;
            const collabBadge = collabCount > 0
                ? `<span class="proj-collab-badge" onclick="event.stopPropagation();openProject('${p._id}');setTimeout(showCollabModal,300)" title="${t('collaboration','협업')} (${collabCount})"><svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="1.8"><circle cx="4.5" cy="4.5" r="2.2"/><circle cx="9" cy="5.5" r="1.8"/><path d="M.5 10.5c0-2 1.5-3.5 4-3.5s2.5.8 3 1.5"/></svg><span>${collabCount}</span></span>`
                : '';
            list.append(`
                <div class="project-item ${isActive ? 'active' : ''}" onclick="openProject('${p._id}')">
                    <div class="proj-icon ${typeClass}">
                        ${projTypeIcon}
                    </div>
                    <div class="proj-info">
                        <div class="proj-name">${escapeHtml(p.name)}${collabBadge}</div>
                        <div class="proj-date">${date}</div>
                    </div>
                    <button class="proj-delete-btn" onclick="deleteProjectById('${p._id}', event)" title="${t('delete','삭제')}">
                        <svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="1.8"><path d="M3 4h8M5 4V3a1 1 0 011-1h2a1 1 0 011 1v1M10 4v7a1 1 0 01-1 1H5a1 1 0 01-1-1V4"/></svg>
                    </button>
                    <div class="proj-status-dot ${p.status || 'draft'}"></div>
                </div>
            `);
        }
    });

    // 페이징 UI (2페이지 이상일 때만)
    if (totalPages > 1) {
        let paginationHtml = '<div class="project-pagination">';
        paginationHtml += `<button class="pp-btn ${state.projectPage === 0 ? 'disabled' : ''}" onclick="goProjectPage(${state.projectPage - 1})" ${state.projectPage === 0 ? 'disabled' : ''}><svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2"><path d="M8 2L3 6l5 4"/></svg></button>`;

        // 페이지 번호 (최대 5개 표시)
        let startPage = Math.max(0, state.projectPage - 2);
        let endPage = Math.min(totalPages - 1, startPage + 4);
        if (endPage - startPage < 4) startPage = Math.max(0, endPage - 4);

        for (let i = startPage; i <= endPage; i++) {
            paginationHtml += `<button class="pp-num ${i === state.projectPage ? 'active' : ''}" onclick="goProjectPage(${i})">${i + 1}</button>`;
        }

        paginationHtml += `<button class="pp-btn ${state.projectPage === totalPages - 1 ? 'disabled' : ''}" onclick="goProjectPage(${state.projectPage + 1})" ${state.projectPage === totalPages - 1 ? 'disabled' : ''}><svg width="12" height="12" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 2l5 4-5 4"/></svg></button>`;
        paginationHtml += '</div>';
        list.append(paginationHtml);
    }
}

function goProjectPage(page) {
    state.projectPage = page;
    renderProjectList(true);
}

function showNewProjectModal() {
    $('#newProjectName').val('');
    $('#newProjectDesc').val('');
    state.selectedProjectType = 'slide';
    $('.type-option').removeClass('active');
    $('[data-type="slide"]').addClass('active');
    $('#newProjectModal').show();
}

function selectProjectType(type) {
    state.selectedProjectType = type;
    $('.type-option').removeClass('active');
    $(`[data-type="${type}"]`).addClass('active');
}

async function createProject() {
    const name = $('#newProjectName').val().trim();
    if (!name) { showToast(t('msgEnterProjectName'), 'error'); return; }

    try {
        const res = await apiPost('/api/projects', {
            name: name,
            description: $('#newProjectDesc').val().trim(),
            project_type: state.selectedProjectType,
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

        // 이전 프로젝트의 폴링 중지 및 OnlyOffice 에디터 정리
        stopCollabPolling();
        destroyOnlyOfficeEditor();

        const res = await apiGet('/api/projects/' + projectId);
        state.currentProject = res.project;
        state.resources = res.resources || [];
        state.generatedSlides = res.generated_slides || [];
        state.generatedExcel = res.generated_excel || null;
        state.onlyofficeDoc = res.onlyoffice_doc || null;
        state.generatedDocx = res.generated_docx || null;
        state.currentSlideIndex = 0;
        state.collabRole = res.project._collab_role || 'owner';

        hideLoading();
        renderProjectWorkspace();
        renderProjectList();

        // 배경/오브젝트 이미지 백그라운드 프리로드
        _preloadSlideImages(state.generatedSlides);

        // 협업 초기화
        await initCollaboration(projectId);
    } catch (e) {
        hideLoading();
        showToast(t('msgLoadingProject'), 'error');
    }
}

// ---- 프로젝트 제목 인라인 편집 ----
var _titleSaving = false;

function startTitleEdit() {
    if ($('#titleEditInput').length) return; // 이미 편집 중
    var $title = $('#wsProjectTitle');
    var currentName = state.currentProject.name;
    var $input = $('<input type="text" id="titleEditInput" class="ws-title-edit">');
    $input.val(currentName);
    $title.hide().after($input);
    $input.focus();
    $input[0].select();

    $input.on('keydown', function(e) {
        if (e.key === 'Enter') {
            e.preventDefault();
            saveTitleEdit($input, currentName);
        }
        if (e.key === 'Escape') {
            e.preventDefault();
            cancelTitleEdit($input);
        }
    });
    $input.on('blur', function() {
        if (!_titleSaving) {
            saveTitleEdit($input, currentName);
        }
    });
}

async function saveTitleEdit($input, originalName) {
    if (_titleSaving) return;
    var newName = $input.val().trim();
    if (!newName || newName === originalName) {
        cancelTitleEdit($input);
        return;
    }
    _titleSaving = true;
    try {
        await apiPut('/api/projects/' + state.currentProject._id, { name: newName });
        state.currentProject.name = newName;
        $('#wsProjectTitle').text(newName).show();
        $input.remove();
        showToast(t('msgUpdated'), 'success');
        await loadProjects();
    } catch (e) {
        showToast(e.message, 'error');
        cancelTitleEdit($input);
    } finally {
        _titleSaving = false;
    }
}

function cancelTitleEdit($input) {
    $('#wsProjectTitle').show();
    $input.remove();
}

function renderProjectWorkspace() {
    $('#emptyState').hide();
    $('#btnAdminMain').hide();
    $('#projectWorkspace').css('display', 'flex');

    // 헤더
    $('#wsProjectTitle').text(state.currentProject.name);
    const statusLabel = { draft: t('statusDraft'), preparing: t('statusPreparing'), generating: t('statusGenerating'), generated: t('statusGenerated'), stop_requested: t('statusStopped'), stopped: t('statusStopped') }[state.currentProject.status] || t('statusDraft');
    $('#wsProjectStatus').text(statusLabel).attr('class', 'ws-status ' + (state.currentProject.status || 'draft'));

    // 제목 클릭 인라인 편집
    $('#wsProjectTitle').off('click').on('click', function() {
        startTitleEdit();
    });

    // 리소스 칩
    renderResourceChips();

    const projectType = state.currentProject.project_type || 'slide';
    const isExcel = (projectType === 'excel');
    const isOnlyOffice = projectType.startsWith('onlyoffice_');
    const isWord = (projectType === 'word');

    // 모든 워크스페이스 숨기기
    $('#excelWorkspace').hide();
    $('#onlyofficeWorkspace').hide();
    $('#wordWorkspace').hide();
    $('#btnModifyExcel').hide();
    $('#btnDocxTemplate').hide();
    _destroyExcelCharts();
    // 전체화면 상태 해제
    $('#appView').removeClass('canvas-fullscreen');
    $('.canvas-fullscreen-exit').remove();
    $('#btnCanvasFullscreen').hide();

    if (isOnlyOffice) {
        // OnlyOffice 모드
        destroyOnlyOfficeEditor();
        $('#slideEmpty').hide();
        $('#slidePreview').hide();
        $('#onlyofficeWorkspace').show();
        $('#templateSelectBtn').hide();
        $('#slideCountSelect').hide();
        $('#docxPageCountSelect').hide();
        $('#btnAddSlide').hide();

        const ooLabels = {
            'onlyoffice_pptx': 'PPT 생성',
            'onlyoffice_xlsx': 'Excel 생성',
            'onlyoffice_docx': 'Word 생성',
        };
        const ooTitles = {
            'onlyoffice_pptx': '프레젠테이션',
            'onlyoffice_xlsx': '스프레드시트',
            'onlyoffice_docx': '문서',
        };
        const ooPlaceholders = {
            'onlyoffice_pptx': '프레젠테이션 내용에 대한 지침을 입력하세요...',
            'onlyoffice_xlsx': '스프레드시트에 정리할 내용에 대한 지침을 입력하세요...',
            'onlyoffice_docx': '문서에 작성할 내용에 대한 지침을 입력하세요...',
        };
        $('#btnGenerate span').text(ooLabels[projectType] || '생성');
        $('#onlyofficeTitle').text(ooTitles[projectType] || '문서');
        $('#instructionsInput').attr('placeholder', ooPlaceholders[projectType] || '지침을 입력하세요...');

        // onlyoffice_pptx는 템플릿 선택 + 슬라이드 수 필요
        if (projectType === 'onlyoffice_pptx') {
            $('#templateSelectBtn').show();
            $('#slideCountSelect').show();
        }
        // onlyoffice_docx는 페이지 수 선택 표시
        if (projectType === 'onlyoffice_docx') {
            $('#docxPageCountSelect').show();
            $('#btnDocxTemplate').show();
        }

        // OnlyOffice 워크스페이스 초기화
        initOnlyOfficeWorkspace();
    } else if (isExcel) {
        // 엑셀 모드
        $('#slideEmpty').hide();
        $('#slidePreview').hide();
        $('#excelWorkspace').show();
        $('#templateSelectBtn').hide();
        $('#slideCountSelect').hide();
        $('#docxPageCountSelect').hide();
        $('#btnAddSlide').hide();
        $('#btnGenerate span').text(t('generateExcel'));
        $('#instructionsInput').attr('placeholder', '엑셀에 정리할 내용에 대한 지침을 입력하세요...');
        $('#btnCanvasFullscreen').show();
        initExcelWorkspace();
    } else if (isWord) {
        // 워드 모드
        $('#slideEmpty').hide();
        $('#slidePreview').hide();
        $('#wordWorkspace').show();
        $('#templateSelectBtn').hide();
        $('#slideCountSelect').hide();
        $('#docxPageCountSelect').show();
        $('#btnDocxTemplate').show();
        $('#btnAddSlide').hide();
        $('#btnGenerate span').text(t('generateWord'));
        $('#instructionsInput').attr('placeholder', '문서에 작성할 내용에 대한 지침을 입력하세요...');
        $('#btnCanvasFullscreen').show();
        initWordWorkspace();
    } else {
        // 슬라이드 모드
        $('#templateSelectBtn').show();
        $('#slideCountSelect').show();
        $('#docxPageCountSelect').hide();
        $('#btnAddSlide').show();
        $('#btnGenerate span').text(t('generateBtn'));
        $('#instructionsInput').attr('placeholder', t('instructionsPlaceholder'));
        $('#btnCanvasFullscreen').show();
        renderSlideArea();
    }

    // 지침 복원
    if (state.currentProject.instructions) {
        $('#instructionsInput').val(state.currentProject.instructions);
        autoResizeTextarea(document.getElementById('instructionsInput'));
    } else {
        $('#instructionsInput').val('');
    }

    if (!isExcel && !isOnlyOffice && !isWord) {
        // 템플릿 선택 복원
        state.selectedTemplateId = state.currentProject.template_id || null;
        // 템플릿 slide_size 복원
        const _tmpl = state.templates.find(t => t._id === state.selectedTemplateId);
        state._templateSlideSize = (_tmpl && _tmpl.slide_size) || '16:9';
        updateSlideCanvasAspect();
        updateTemplateButtonLabel();

        // 수동 모드 상태 복원
        state.manualMode = !!state.currentProject.manual_mode;
        updateManualModeUI();
    } else if (projectType === 'onlyoffice_pptx') {
        state.selectedTemplateId = state.currentProject.template_id || null;
        const _tmplOo = state.templates.find(t => t._id === state.selectedTemplateId);
        state._templateSlideSize = (_tmplOo && _tmplOo.slide_size) || '16:9';
        updateSlideCanvasAspect();
        updateTemplateButtonLabel();
    }

    // 협업 UI 업데이트
    updateCollabUI();

    // docx 템플릿 로드
    if (projectType === 'word' || projectType === 'onlyoffice_docx') {
        loadDocxTemplate();
    } else {
        state.docxTemplateId = null;
    }
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

async function deleteProjectById(projectId, event) {
    if (event) { event.stopPropagation(); event.preventDefault(); }
    if (!confirm(t('msgConfirmDelete'))) return;
    try {
        await apiDelete('/api/projects/' + projectId);
        showToast(t('msgProjectDeleted'), 'success');
        if (state.currentProject && state.currentProject._id === projectId) {
            state.currentProject = null;
            showEmptyState();
        }
        await loadProjects();
        renderRecentProjects();
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
        state.generatedExcel = null;
        state.onlyofficeDoc = null;
        state.generatedDocx = null;
        state.currentProject.status = 'draft';
        state.currentProject.template_id = null;
        state.selectedTemplateId = null;
        state._templateSlideSize = '16:9';
        updateSlideCanvasAspect();
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
    const icons = { file: '📎', text: '📝', web: '🔍', url: '🔗', youtube: '▶️', image: '🖼️' };
    const typeLabels = { file: 'File', text: 'Text', web: 'Web', url: 'URL', youtube: 'YouTube', image: 'Image' };

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

    // 전체 삭제 버튼 표시
    const footer = $('#resourceDdFooter');
    if (count > 0) {
        $('#deleteAllLabel').text(t('deleteAll'));
        footer.show();
    } else {
        footer.hide();
    }
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

function triggerImageResourceUpload() {
    $('#imageResourceUploadInput').click();
}

async function handleImageResourceUpload(event) {
    const files = event.target.files;
    if (!files || files.length === 0) return;

    // 이미지 파일만 필터링
    const imageFiles = Array.from(files).filter(f => f.type.startsWith('image/'));
    if (imageFiles.length === 0) return;

    // 분석 옵션 모달 표시
    _pendingImageFiles = imageFiles;
    const fileNames = imageFiles.map(f => f.name).join(', ');
    $('#imageAnalyzeFileNames').text(imageFiles.length > 1 ? imageFiles.length + '개 이미지' : imageFiles[0].name);
    $('#imageAnalyzeCheckbox').prop('checked', true);
    $('#imageAnalyzeModal').show();

    event.target.value = '';
}

let _pendingImageFiles = [];

async function confirmImageUpload() {
    const analyze = $('#imageAnalyzeCheckbox').is(':checked');
    $('#imageAnalyzeModal').hide();

    const files = _pendingImageFiles;
    _pendingImageFiles = [];
    if (!files.length) return;

    const projectId = state.currentProject._id;
    let successCount = 0;

    showLoading(analyze ? '이미지 내용 분석 중...' : '이미지 업로드 중...');

    for (let i = 0; i < files.length; i++) {
        const file = files[i];

        const formData = new FormData();
        formData.append('file', file);
        formData.append('project_id', projectId);
        formData.append('title', file.name);
        formData.append('analyze', analyze ? '1' : '0');

        try {
            const res = await apiUpload('/api/resources/image', formData);
            if (res.resource) {
                state.resources.push(res.resource);
                successCount++;
            }
        } catch (e) {
            console.error('Image upload failed:', file.name, e);
        }
    }

    hideLoading();
    renderResourceChips();

    if (successCount > 0) {
        showToast(successCount + '개 이미지가 추가되었습니다', 'success');
    } else {
        showToast('이미지 업로드 실패', 'error');
    }
}

function cancelImageUpload() {
    _pendingImageFiles = [];
    $('#imageAnalyzeModal').hide();
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

function showURLModal() {
    $('#urlInput').val('');
    $('#urlResourceModal').show();
}

async function addURLResources() {
    const text = $('#urlInput').val().trim();
    if (!text) { showToast(t('msgEnterContent'), 'error'); return; }

    const urls = text.split('\n').map(u => u.trim()).filter(u => u);
    if (urls.length === 0) { showToast(t('msgEnterContent'), 'error'); return; }

    try {
        $('#btnAddUrls').prop('disabled', true);
        showLoading(t('msgUrlCollecting'));
        const res = await apiPost('/api/resources/urls', {
            project_id: state.currentProject._id,
            urls: urls,
        });

        const resources = res.resources || [];
        resources.forEach(r => state.resources.push(r));
        renderResourceChips();
        closeModal('urlResourceModal');

        const errCount = (res.errors || []).length;
        if (resources.length > 0) {
            showToast(resources.length + '개 ' + t('msgUrlAdded') + (errCount > 0 ? ` (${errCount}개 실패)` : ''), 'success');
        } else if (errCount > 0) {
            showToast('URL 수집 실패: ' + res.errors[0].error, 'error');
        }
    } catch (e) {
        showToast(e.message, 'error');
    } finally {
        $('#btnAddUrls').prop('disabled', false);
        hideLoading();
    }
}

async function showResourceContent(resourceId) {
    const resource = state.resources.find(r => r._id === resourceId);
    if (!resource) return;

    const icons = { file: '📎', text: '📝', web: '🔍', image: '🖼️' };
    const iconClasses = { file: 'file', text: 'text', web: 'web', url: 'url', youtube: 'youtube', image: 'image' };
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
    if (resource.resource_type === 'image' && resource.file_path) {
        let html = '';
        if (content && content.trim()) {
            html = '<div class="image-preview-layout">'
                + '<div class="image-preview-thumb"><img src="' + resource.file_path + '"></div>'
                + '<div class="image-desc-section">'
                + '<div class="image-desc-header"><svg width="14" height="14" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M12 20h9"/><path d="M16.5 3.5a2.121 2.121 0 013 3L7 19l-4 1 1-4L16.5 3.5z"/></svg> AI 이미지 분석</div>'
                + '<div class="image-desc-body">' + simpleMarkdownToHtml(content) + '</div>'
                + '</div></div>';
        } else {
            html = '<div style="text-align:center;"><img src="' + resource.file_path + '" style="max-width:100%;max-height:400px;border-radius:8px;"></div>';
        }
        viewer.html(html);
    } else if (content && content.trim()) {
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

    const modalCard = $('#resourceContentModal .modal-card');
    if (resource.resource_type === 'image' && content && content.trim()) {
        modalCard.css('max-width', '960px');
        viewer.css({'max-height': 'none', 'overflow-y': 'visible'});
    } else {
        modalCard.css('max-width', '720px');
        viewer.css({'max-height': '60vh', 'overflow-y': 'auto'});
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

async function deleteAllResources() {
    if (!state.currentProject) return;
    if (!confirm(t('msgConfirmDeleteAllRes'))) return;
    try {
        await apiDelete('/api/resources/all/' + state.currentProject._id);
        state.resources = [];
        renderResourceChips();
        showToast(t('msgAllResDeleted'), 'success');
    } catch (e) {
        showToast(e.message, 'error');
    }
}

// ============ 수동 생성 모드 ============

async function openAddSlideModal() {
    if (!state.currentProject) return;

    // 템플릿 미선택 시 선택 유도
    if (!state.selectedTemplateId) {
        showToast(t('msgSelectTemplate'), 'error');
        openTemplatePickerModal();
        return;
    }

    showTemplateSlidePickerModal();
}

function updateManualModeUI() {
    // 슬라이드가 있으면 AI 지침 바 표시 (viewer 제외)
    if (state.generatedSlides.length > 0 && state.collabRole !== 'viewer') {
        $('#slideInstructionBar').show();
    } else {
        $('#slideInstructionBar').hide();
    }
    // 좌측 패널 + 추가 버튼 갱신
    renderSlideThumbList();
}

async function showTemplateSlidePickerModal() {
    if (!state.selectedTemplateId) {
        showToast(t('msgSelectTemplate'), 'error');
        return;
    }

    const grid = $('#templateSlideGrid');
    grid.empty().html('<div style="text-align:center;padding:20px;color:var(--text-tertiary);">Loading...</div>');
    $('#templateSlidePickerModal').show();

    try {
        const res = await apiGet('/api/templates/' + state.selectedTemplateId + '/slides');
        const slides = res.slides || [];
        state.manualTemplateSlides = slides;

        grid.empty();
        if (slides.length === 0) {
            grid.html('<div style="text-align:center;padding:20px;color:var(--text-tertiary);">템플릿에 슬라이드가 없습니다</div>');
            return;
        }

        slides.forEach(slide => {
            const item = $('<div>').addClass('template-slide-item').attr('data-slide-id', slide._id);
            const thumbContainer = $('<div>').addClass('template-slide-thumb');
            const _td240 = getThumbDimensions(240);
            renderSlideToContainer(thumbContainer, slide, _td240.w, _td240.h);
            item.append(thumbContainer);

            const meta = slide.slide_meta || {};
            const typeLabel = {
                title_slide: 'Cover', toc: 'Contents',
                section_divider: 'Chapter', body: 'Body', closing: 'Closing'
            }[meta.content_type] || meta.content_type || 'Slide';
            item.append(`<div class="template-slide-label">${typeLabel}</div>`);

            item.on('click', () => addManualSlide(slide._id));
            grid.append(item);
        });
    } catch (e) {
        grid.html('<div style="text-align:center;padding:20px;color:var(--danger);">' + e.message + '</div>');
    }
}

async function addManualSlide(templateSlideId) {
    if (!state.currentProject) return;

    // 삽입 위치 결정: 선택된 슬라이드 다음 또는 맨 끝
    let insertAfterOrder = null;
    if (state.currentSlideIndex >= 0 && state.currentSlideIndex < state.generatedSlides.length) {
        const selectedSlide = state.generatedSlides[state.currentSlideIndex];
        if (selectedSlide && selectedSlide.order != null) {
            insertAfterOrder = selectedSlide.order;
        }
    }

    try {
        showLoading(t('msgSlideAdded'));
        const res = await apiPost('/api/generate/manual-slide', {
            project_id: state.currentProject._id,
            template_slide_id: templateSlideId,
            insert_after_order: insertAfterOrder,
        });

        // 삽입 후 전체 슬라이드 재로딩 (order shift 반영)
        const slidesRes = await apiGet('/api/generate/' + state.currentProject._id + '/slides');
        state.generatedSlides = slidesRes.slides || [];

        // 새로 추가된 슬라이드로 이동
        const newSlideIndex = state.generatedSlides.findIndex(s => s._id === res.slide._id);
        state.currentSlideIndex = newSlideIndex >= 0 ? newSlideIndex : state.generatedSlides.length - 1;

        // 프로젝트 상태 업데이트
        state.currentProject.status = 'generated';

        // 협업 타임스탬프 캐시 갱신 (불필요한 재로딩 방지)
        state.generatedSlides.forEach(s => {
            if (s._id && s.updated_at) {
                state.lastSlideTimestamps[s._id] = s.updated_at;
            }
        });

        // UI 표시
        $('#slideEmpty').hide();
        $('#slidePreview').show();
        renderSlideThumbList();
        goToSlide(state.currentSlideIndex);
        updateManualModeUI();

        closeModal('templateSlidePickerModal');
        hideLoading();
        showToast(t('msgSlideAdded'), 'success');

        // 추가 후 자동으로 편집 모드 진입
        if (!state.editMode) {
            enterEditMode();
        }
    } catch (e) {
        hideLoading();
        showToast(e.message, 'error');
    }
}

function handleSlideInstructionKey(e) {
    if (e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        generateSlideText();
    }
}

function _collectCurrentSlideContent(slide) {
    // 현재 슬라이드의 텍스트 내용을 수집하여 LLM에 전달
    if (!slide) return {};
    const contents = {};
    const items = slide.items || [];
    let subIdx = 0, descIdx = 0;

    (slide.objects || []).forEach(obj => {
        if (obj.obj_type !== 'text') return;
        const role = obj.role || obj._auto_role || '';
        const placeholder = obj._placeholder || obj.placeholder || role || obj.obj_type;

        if (role === 'subtitle' && items.length > 0) {
            if (subIdx < items.length) {
                contents[placeholder + '_' + subIdx] = items[subIdx].heading || '';
                subIdx++;
            }
        } else if (role === 'description' && items.length > 0) {
            if (descIdx < items.length) {
                contents[placeholder + '_' + descIdx] = items[descIdx].detail || '';
                descIdx++;
            }
        } else {
            contents[placeholder] = obj.generated_text || obj.text_content || '';
        }
    });

    return { contents, items };
}

async function generateSlideText() {
    if (!state.currentProject || state.generatedSlides.length === 0) return;

    const slide = state.generatedSlides[state.currentSlideIndex];
    if (!slide) return;

    const instruction = $('#slideInstructionInput').val().trim();
    if (!instruction) {
        $('#slideInstructionInput').focus();
        return;
    }

    // 편집 모드에서 수정된 텍스트 수집
    if (state.editMode) collectEditedText();

    // 현재 슬라이드의 기존 내용 수집
    const currentContent = _collectCurrentSlideContent(slide);
    const oldItemsCount = (slide.items || []).length;

    try {
        // 로딩 UI 표시
        $('#btnSlideAI').prop('disabled', true).html(
            '<svg class="spin-icon" width="13" height="13" viewBox="0 0 16 16" fill="none" stroke="currentColor" stroke-width="2"><path d="M8 1a7 7 0 106.93 6"/></svg> ' +
            t('msgGeneratingText')
        ).css('opacity', '0.7');
        $('#slideInstructionInput').prop('disabled', true);

        const res = await apiPost('/api/generate/slide-text', {
            project_id: state.currentProject._id,
            slide_id: slide._id,
            instruction: instruction,
            template_slide_id: slide.template_slide_id || '',
            current_content: currentContent,
        });

        let finalObjects = res.objects;
        let finalItems = res.items;
        const newItemsCount = (finalItems || []).length;

        // 항목 수가 달라졌으면 적합한 템플릿 슬라이드로 자동 전환
        if (newItemsCount > 0 && newItemsCount !== oldItemsCount) {
            try {
                // contents 재구성 (switch-template-slide에 전달)
                const contentsForSwitch = {};
                (finalObjects || []).forEach(obj => {
                    if (obj.obj_type === 'text') {
                        const ph = obj.placeholder || obj._auto_placeholder || obj.role || '';
                        if (ph && obj.generated_text) {
                            contentsForSwitch[ph] = obj.generated_text;
                        }
                    }
                });

                const switchRes = await apiPost('/api/generate/switch-template-slide', {
                    slide_id: slide._id,
                    items_count: newItemsCount,
                    contents: contentsForSwitch,
                    items: finalItems,
                });

                if (switchRes.switched && switchRes.slide) {
                    finalObjects = switchRes.slide.objects;
                    finalItems = switchRes.slide.items;
                    slide.template_slide_id = switchRes.new_template_slide_id;
                    slide.background_image = switchRes.slide.background_image;
                }
            } catch (switchErr) {
                // 템플릿 전환 실패해도 텍스트 변경은 유지
                console.log('Template switch skipped:', switchErr.message);
            }
        }

        // 슬라이드 데이터 업데이트
        slide.objects = finalObjects;
        slide.items = finalItems;

        // 타이핑 애니메이션으로 캔버스에 표시
        if (state.editMode) {
            _exitEditModeClean();
        }
        await _animateSlideTextUpdate(state.currentSlideIndex);

        // 자동 저장
        try {
            await fetch(`/${state.jwtToken}/api/generate/slides/${slide._id}`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ objects: slide.objects, items: slide.items || [] }),
            });
        } catch (saveErr) {
            console.log('Auto-save failed:', saveErr);
        }

        // 썸네일 업데이트
        renderSlideThumbList();
        renderSlideThumbnails();
        renderSlideTextPanel();

        // 입력 초기화
        $('#slideInstructionInput').val('');
        autoResizeTextarea(document.getElementById('slideInstructionInput'));

        showToast(t('msgSlideTextGenerated'), 'success');
    } catch (e) {
        showToast(e.message, 'error');
    } finally {
        // 버튼 원래 상태로 복원
        $('#btnSlideAI').prop('disabled', false).html(
            '<svg width="13" height="13" fill="none" stroke="currentColor" stroke-width="2"><path d="M11 1L3 6l8 5V1z"/></svg> ' +
            t('aiModifyRequest', 'AI 수정 요청')
        ).css('opacity', '');
        $('#slideInstructionInput').prop('disabled', false);
    }
}

async function _animateSlideTextUpdate(index) {
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
    const _sz = getCurrentSlideSize();
    const scaleX = canvasW / _sz.w;
    const scaleY = canvasH / _sz.h;

    let descIndex = 0;
    let subIndex = 0;
    const items = slide.items || [];

    // 모든 오브젝트를 먼저 배치 (텍스트는 비워둠)
    const textEls = [];
    (slide.objects || []).forEach(obj => {
        const div = $('<div>').addClass('preview-obj').css({
            position: 'absolute',
            left: (obj.x * scaleX) + 'px',
            top: (obj.y * scaleY) + 'px',
            width: (obj.width * scaleX) + 'px',
            zIndex: obj.z_index !== undefined ? obj.z_index : 10,
        });

        if (obj.obj_type === 'image' && obj.image_url) {
            div.css('height', (obj.height * scaleY) + 'px');
            const imgFit = obj.image_fit || 'contain';
            div.append(`<img src="${obj.image_url}" style="width:100%;height:100%;object-fit:${imgFit};">`);
            canvas.append(div);
        } else if (obj.obj_type === 'shape') {
            div.css('height', (obj.height * scaleY) + 'px');
            div.html(_createEditShapeSVG(obj));
            canvas.append(div);
        } else if (obj.obj_type === 'text') {
            const style = obj.text_style || {};
            const role = obj.role || '';
            let text = obj.generated_text || obj.text_content || '';

            if (role === 'subtitle' && items.length > 0 && subIndex < items.length) {
                text = items[subIndex].heading || '';
                subIndex++;
            } else if (role === 'description' && items.length > 0 && descIndex < items.length) {
                text = items[descIndex].detail || '';
                descIndex++;
            }

            const fontSize = (style.font_size || 16) * scaleX;
            div.css({
                fontSize: fontSize + 'px',
                fontFamily: style.font_family || 'Inter, sans-serif',
                fontWeight: style.bold ? '700' : '400',
                fontStyle: style.italic ? 'italic' : 'normal',
                color: style.color || '#000',
                textAlign: style.align || 'left',
                lineHeight: '1.4',
                overflow: 'hidden',
                height: 'auto',
                textDecoration: [
                    style.underline ? 'underline' : '',
                    style.strikethrough ? 'line-through' : ''
                ].filter(Boolean).join(' ') || 'none',
            });

            if (role === 'number' || role === 'governance') {
                div.css('height', (obj.height * scaleY) + 'px');
            }

            canvas.append(div);
            if (text) {
                textEls.push({ $el: div, text: text });
            }
        }
    });

    // 텍스트 타이핑 애니메이션
    for (const item of textEls) {
        const $el = item.$el;
        const text = item.text;
        const cursor = $('<span style="display:inline-block;width:2px;height:1em;background:var(--accent);animation:blink 0.6s infinite;vertical-align:text-bottom;margin-left:1px;"></span>');
        $el.empty().append(cursor);

        const speed = Math.max(10, Math.min(30, 600 / text.length));
        for (let k = 0; k < text.length; k++) {
            const ch = text[k];
            if (ch === '\n') {
                cursor.before($('<br>')[0]);
            } else {
                cursor.before(document.createTextNode(ch));
            }
            if (k % 2 === 0) await new Promise(r => setTimeout(r, speed));
        }
        cursor.remove();
    }
}

async function deleteManualSlide(slideId) {
    if (!confirm(t('msgConfirmDeleteRes'))) return;

    try {
        await apiDelete('/api/generate/manual-slide/' + slideId);
        state.generatedSlides = state.generatedSlides.filter(s => s._id !== slideId);

        if (state.generatedSlides.length === 0) {
            state.currentSlideIndex = 0;
            $('#previewCanvas').find('.preview-obj,.edit-obj-wrap').remove();
            $('#previewBg').css('background-image', 'none');
        } else {
            if (state.currentSlideIndex >= state.generatedSlides.length) {
                state.currentSlideIndex = state.generatedSlides.length - 1;
            }
            goToSlide(state.currentSlideIndex);
        }
        renderSlideThumbList();
        showToast(t('msgSlideDeleted'), 'success');
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
            const _tmplSz = tmpl.slide_size || '16:9';
            const _tdPicker = getThumbDimensions(240, _tmplSz);
            // 썸네일 aspect-ratio를 개별 템플릿의 slide_size에 맞춤
            const _szPicker = getSlideCanvasSize(_tmplSz);
            thumbContainer.css('aspect-ratio', _szPicker.w + ' / ' + _szPicker.h);
            renderSlideToContainer(thumbContainer, slideData, _tdPicker.w, _tdPicker.h, _tmplSz);
        }

        grid.append(card);
    });

    $('#templatePickerModal').show();
}

function selectTemplateFromPicker(templateId) {
    state.selectedTemplateId = templateId;
    // 선택한 템플릿의 slide_size 저장
    const tmpl = state.templates.find(t => t._id === templateId);
    state._templateSlideSize = (tmpl && tmpl.slide_size) || '16:9';
    updateSlideCanvasAspect();
    updateTemplateButtonLabel();
    closeModal('templatePickerModal');

    // 프로젝트에 선택된 템플릿 저장 (다시 열 때 복원용)
    if (state.currentProject && state.currentProject._id) {
        state.currentProject.template_id = templateId;
        apiPut('/api/projects/' + state.currentProject._id, { template_id: templateId }).catch(() => {});
    }
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

// ============ 이미지 슬라이드 생성 ============
async function generateImageSlides(templateId) {
    _isGenerating = true;
    _animationCancelled = true;
    state.generatedSlides = [];
    state.currentSlideIndex = 0;

    _showStopButton();

    $('#slideEmpty').hide();
    $('#slidePreview').css('display', 'flex');
    $('#wsSlideTools').css('display', 'flex');
    _setSlideToolsDisabled(true);

    // 로딩 UI
    $('#slideThumbnails').empty();
    $('#slideThumbList').empty();
    $('#slideCounter').text('0 / 0');
    $('#slideCounterInline').text('0 / 0');
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
                <div class="canvas-loading-title">이미지 슬라이드를 만들고 있습니다</div>
                <div class="canvas-loading-sub">이미지를 배치하는 중<span class="canvas-loading-dots"><span>.</span><span>.</span><span>.</span></span></div>
            </div>
        </div>
    `);

    try {
        const res = await apiPost('/api/generate/image-slides', {
            project_id: state.currentProject._id,
            template_id: templateId,
            instructions: '',
            lang: $('#langSelect').val(),
            slide_count: 'auto',
        });

        const slides = res.slides || [];
        state.generatedSlides = slides;

        if (slides.length > 0) {
            state.currentSlideIndex = 0;
            renderSlideThumbnails();
            renderSlideThumbList();
            renderSlideTextPanel();
            navigateSlide(0);
            showToast(slides.length + '개 이미지 슬라이드가 생성되었습니다', 'success');
        } else {
            showToast('이미지 슬라이드를 생성할 수 없습니다', 'error');
            $('#slideEmpty').show();
            $('#slidePreview').hide();
            $('#wsSlideTools').hide();
        }
    } catch (e) {
        showToast(e.message || '이미지 슬라이드 생성 실패', 'error');
        if (state.generatedSlides.length === 0) {
            $('#slideEmpty').show();
            $('#slidePreview').hide();
            $('#wsSlideTools').hide();
        }
    } finally {
        _isGenerating = false;
        $('#canvasLoadingOverlay').remove();
        _showGenerateOrRestartButton();
        _setSlideToolsDisabled(false);
    }
}

// ============ PPT 생성 ============
async function generatePPT() {
    const templateId = state.selectedTemplateId;
    const instructions = $('#instructionsInput').val().trim();
    const lang = $('#langSelect').val();
    const slideCount = $('#slideCountSelect').val();

    if (!templateId) { showToast(t('msgSelectTemplate'), 'error'); return; }
    const hasImageResources = state.resources.some(r => r.resource_type === 'image');
    const hasNonImageResources = state.resources.some(r => r.resource_type !== 'image');
    if (!instructions && !hasImageResources) { showToast(t('msgNeedInstructions'), 'error'); return; }
    if (state.resources.length === 0 && !instructions) { showToast(t('msgEnterInstructions', '지침을 입력하세요'), 'error'); return; }

    // 이미지 리소스만 있고 지침 없음 → 이미지 슬라이드 전용 생성
    if (hasImageResources && !instructions && !hasNonImageResources) {
        return generateImageSlides(templateId);
    }

    _isGenerating = true;
    _animationCancelled = true;
    state.generatedSlides = [];
    state.currentSlideIndex = 0;
    state._contentQueue = [];
    state._generationComplete = false;

    // 생성 중 → 중단 버튼으로 전환
    _showStopButton();

    // 로딩 오버레이 대신 슬라이드 프리뷰 영역 표시 + 아웃라인 탭
    $('#slideEmpty').hide();
    $('#slidePreview').css('display', 'flex');
    $('#wsSlideTools').css('display', 'flex');

    // 생성 중 슬라이드 도구 버튼 비활성화
    _setSlideToolsDisabled(true);
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

    _abortController = new AbortController();

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
            signal: _abortController.signal,
        });

        if (!response.ok) {
            let errMsg = '생성 실패';
            try { const err = await response.json(); errMsg = err.detail || errMsg; } catch (_) {}
            throw new Error(errMsg);
        }

        _streamReader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';

        while (true) {
            const { done, value } = await _streamReader.read();
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
        if (e.name === 'AbortError') {
            // 사용자가 중단한 경우 - 정상 처리
            console.log('Generation aborted by user');
        } else {
            showToast(e.message || '슬라이드 생성 중 오류가 발생했습니다.', 'error');
            console.error('Generate stream error:', e);
        }
        if (state.generatedSlides.length === 0) {
            $('#slideEmpty').show();
            $('#slidePreview').hide();
            $('#wsSlideTools').hide();
        }
    } finally {
        _isGenerating = false;
        _streamReader = null;
        _abortController = null;
        $('#canvasLoadingOverlay').remove();
        $('#streamingProgress').remove();
        $('#slideTextList').removeClass('streaming-active');
        _showGenerateOrRestartButton();
        // 생성 완료 → 슬라이드 도구 버튼 활성화
        _setSlideToolsDisabled(false);
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

        case 'slides_skeleton': {
            // 스켈레톤 데이터 저장 (썸네일은 slide_content 도착 시 한 장씩 추가)
            const skeletonSlides = data.slides || [];
            state.generatedSlides = skeletonSlides;
            state.currentSlideIndex = 0;

            // 로딩 오버레이 제거
            $('#canvasLoadingOverlay').remove();

            // 슬라이드 프리뷰 영역 표시
            $('#slideEmpty').hide();
            $('#slidePreview').css('display', 'flex');
            $('#wsSlideTools').css('display', 'flex');

            // 첫 슬라이드 캔버스 렌더링 (배경만 보임)
            renderSlideAtIndex(0);
            updateSlideNav();

            // 썸네일 초기화 (한 장씩 추가될 예정)
            $('#slideThumbnails').empty();
            $('#slideThumbList').empty();

            // 모든 배경 이미지 프리로드 시작
            _preloadSlideImages(skeletonSlides);

            $('#slideCounter').text(`0 / ${skeletonSlides.length}`);
            $('#slideCounterInline').text(`0 / ${skeletonSlides.length}`);
            break;
        }

        case 'slide_content': {
            const idx = data.index;
            const fullSlide = data.slide;

            // 해당 슬라이드를 완전한 데이터로 교체
            state.generatedSlides[idx] = fullSlide;

            // 콘텐츠 큐에 추가
            if (!state._contentQueue) state._contentQueue = [];
            state._contentQueue.push(idx);

            // 첫 번째 콘텐츠 도착 시 타이핑 시작
            if (!_isAnimating && state._contentQueue.length === 1) {
                _processContentQueue();
            }
            break;
        }

        case 'complete':
            renderSlideTextPanel();
            if (!_isAnimating) {
                // 타이핑이 이미 끝난 경우 최종 처리
                state.currentSlideIndex = 0;
                renderSlideAtIndex(0);
                renderSlideThumbnails();
                renderSlideThumbList();
                updateSlideNav();
                _preloadSlideImages(state.generatedSlides);
                showToast(t('msgAllComplete'), 'success');
            } else {
                // 아직 타이핑 중이면 완료 플래그 설정
                state._generationComplete = true;
            }
            break;

        case 'stopped':
            showToast(data.message || t('msgStopped'), 'info');
            if (state.generatedSlides.length > 0) {
                state.currentSlideIndex = 0;
                renderSlideAtIndex(0);
                renderSlideThumbnails();
                renderSlideThumbList();
                updateSlideNav();
                renderSlideTextPanel();
            }
            break;

        case 'error':
            showToast(data.message || '생성 중 오류가 발생했습니다.', 'error');
            break;
    }
}


async function stopGeneration() {
    if (!_isGenerating) return;

    // 중복 클릭 방지
    $('#btnStop').prop('disabled', true);

    try {
        // 1. 서버에 중단 요청 (MongoDB 상태 변경)
        await fetch(`/${state.jwtToken}/api/generate/stop/${state.currentProject._id}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
        });
    } catch (e) {
        console.warn('Stop request failed:', e);
    }

    // 2. 프론트엔드 연결 즉시 끊기
    if (_abortController) {
        _abortController.abort();
    }
    if (_streamReader) {
        try { await _streamReader.cancel(); } catch (_) {}
    }
}


function _setSlideToolsDisabled(disabled) {
    $('#wsSlideTools .slide-tool-btn').prop('disabled', disabled);
    if (disabled) {
        $('#wsSlideTools .slide-tool-btn').addClass('disabled');
    } else {
        $('#wsSlideTools .slide-tool-btn').removeClass('disabled');
    }
}


function _showStopButton() {
    const $btn = $('#btnGenerate');
    $btn.attr('id', 'btnStop')
        .attr('onclick', 'stopGeneration()')
        .removeClass('btn-generate')
        .addClass('btn-stop')
        .prop('disabled', false)
        .html(`
            <svg width="16" height="16" fill="currentColor" viewBox="0 0 16 16">
                <rect x="3" y="3" width="10" height="10" rx="1.5"/>
            </svg>
            <span>${t('stopBtn')}</span>
        `);
    // 생성/수정 중에는 수정 버튼 숨김
    $('#btnModifyExcel').hide();
}


function _showGenerateOrRestartButton() {
    const $btn = $('#btnStop').length ? $('#btnStop') : $('#btnGenerate');
    const projectType = state.currentProject ? state.currentProject.project_type : 'slide';
    const isExcel = projectType === 'excel';
    const isOnlyOffice = projectType && projectType.startsWith('onlyoffice_');

    let hasContent, genLabel, restartLabel;
    if (isOnlyOffice) {
        hasContent = !!state.onlyofficeDoc;
        const ooLabels = { 'onlyoffice_pptx': 'PPT 생성', 'onlyoffice_xlsx': 'Excel 생성', 'onlyoffice_docx': 'Word 생성' };
        genLabel = ooLabels[projectType] || '생성';
        restartLabel = genLabel;
    } else if (isExcel) {
        hasContent = !!state.generatedExcel;
        genLabel = t('generateExcel');
        restartLabel = t('generateExcel');
    } else if (projectType === 'word') {
        hasContent = !!state.generatedDocx;
        genLabel = t('generateWord');
        restartLabel = t('generateWord');
    } else {
        hasContent = state.generatedSlides.length > 0;
        genLabel = t('generateBtn');
        restartLabel = t('restartBtn');
    }

    $btn.attr('id', 'btnGenerate')
        .attr('onclick', 'handleGenerate()')
        .removeClass('btn-stop')
        .addClass('btn-generate')
        .prop('disabled', false)
        .html(hasContent
            ? `<svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
                <path d="M1 4v6h6"/><path d="M3.51 15a9 9 0 1 0 2.13-9.36L1 10"/>
               </svg>
               <span class="i18n-generateBtn">${restartLabel}</span>`
            : `<svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2"><path d="M13 3L5 8.5l8 5.5V3z"/></svg>
               <span class="i18n-generateBtn">${genLabel}</span>`
        );

    // 엑셀 수정 버튼 표시/숨김
    if (isExcel && hasContent) {
        $('#btnModifyExcel').show().find('span').text(t('modifyExcel'));
    } else {
        $('#btnModifyExcel').hide();
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
        $('#slideInstructionBar').hide();
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
    if (state.collabRole !== 'viewer') {
        $('#slideInstructionBar').show();
    }
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
    canvas.find('.live-edit-banner').remove();

    // 다른 사용자가 편집 중인 슬라이드이면 실시간 편집 배너 표시
    if (state.isCollabProject && slide._id && state.activeLocks[slide._id]) {
        const lock = state.activeLocks[slide._id];
        if (lock.user_key !== (state.userInfo && state.userInfo.ky)) {
            canvas.append(`<div class="live-edit-banner"><span class="live-dot"></span>${escapeHtml(lock.user_name)} ${t('editingNow','편집 중')}</div>`);
        }
    }

    if (slide.background_image) {
        $('#previewBg').css('background-image', `url(${slide.background_image})`);
    } else {
        $('#previewBg').css('background-image', 'none');
    }

    const canvasW = canvas.width();
    const canvasH = canvas.height();
    const _sz = getCurrentSlideSize();
    const scaleX = canvasW / _sz.w;
    const scaleY = canvasH / _sz.h;

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
            zIndex: obj.z_index !== undefined ? obj.z_index : 10,
        });

        if (obj.obj_type === 'image' && obj.image_url) {
            const imgFit = obj.image_fit || 'contain';
            div.append(`<img src="${obj.image_url}" style="width:100%;height:100%;object-fit:${imgFit};">`);
        } else if (obj.obj_type === 'shape') {
            div.html(_createEditShapeSVG(obj));
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
                boxSizing: 'border-box',
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
            } else if ((role === 'subtitle' || role === 'description') && items.length > 0) {
                // items 소진된 초과 subtitle/description 오브젝트는 렌더링하지 않음
                return;
            } else {
                div.css('whiteSpace', 'pre-wrap');
                div.text(text);
            }
            // 텍스트가 넘칠 경우 높이 자동 확장 (number/governance 제외)
            if (role !== 'number' && role !== 'governance') {
                div.css({ height: 'auto', minHeight: (obj.height * scaleY) + 'px', overflow: 'visible' });
            }
        }

        canvas.append(div);
    });
}

function renderSlideToContainer(container, slide, thumbW, thumbH, slideSize) {
    const bgStyle = slide.background_image ? `background-image:url(${slide.background_image});background-size:cover;background-position:center;` : '';
    container.attr('style', (container.attr('style') || '') + bgStyle);

    const _sz = slideSize ? getSlideCanvasSize(slideSize) : getCurrentSlideSize();
    const scaleX = thumbW / _sz.w;
    const scaleY = thumbH / _sz.h;
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
            div.html(_createEditShapeSVG(obj));
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
        { const _td64 = getThumbDimensions(64); renderSlideToContainer(thumbEl.find('.slide-thumb-inner'), slide, _td64.w, _td64.h); }
        container.append(thumbEl);
    });
}

function renderSlideThumbList() {
    const list = $('#slideThumbList');
    list.empty();
    state.generatedSlides.forEach((slide, i) => {
        const isActive = i === state.currentSlideIndex;
        const thumbWrap = $('<div>').addClass('slide-thumb-v-wrap');
        const thumbEl = $(`
            <div class="slide-thumb-v ${isActive ? 'active' : ''}" draggable="true" tabindex="0" onclick="goToSlide(${i})" data-slide-idx="${i}" data-slide-id="${slide._id || ''}">
                <div class="slide-thumb-v-num">${i + 1}</div>
                <div class="slide-thumb-v-inner"></div>
            </div>
        `);
        { const _td256 = getThumbDimensions(256); renderSlideToContainer(thumbEl.find('.slide-thumb-v-inner'), slide, _td256.w, _td256.h); }

        // 협업: Lock 배지 표시
        if (state.isCollabProject && slide._id && state.activeLocks[slide._id]) {
            const lock = state.activeLocks[slide._id];
            if (lock.user_key !== (state.userInfo && state.userInfo.ky)) {
                thumbEl.append(`<div class="lock-badge"><svg width="10" height="10" fill="currentColor" viewBox="0 0 16 16"><path d="M8 1a3 3 0 0 0-3 3v2H4a1 1 0 0 0-1 1v6a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1V7a1 1 0 0 0-1-1h-1V4a3 3 0 0 0-3-3zm-2 3a2 2 0 1 1 4 0v2H6V4z"/></svg><span>${escapeHtml(lock.user_name)}</span></div>`);
            }
        }

        thumbWrap.append(thumbEl);

        // 슬라이드 삭제 버튼 (owner/editor + 슬라이드 생성 상태)
        const canModifySlides = (state.collabRole === 'owner' || state.collabRole === 'editor')
            && state.currentProject && state.currentProject.status === 'generated';
        if (canModifySlides) {
            thumbWrap.append(`<button class="slide-thumb-delete" onclick="event.stopPropagation();deleteManualSlide('${slide._id}')" title="삭제">&times;</button>`);
        }
        list.append(thumbWrap);
    });

    // 슬라이드 추가 버튼 (owner/editor + 슬라이드 1개 이상)
    const canAddSlide = (state.collabRole === 'owner' || state.collabRole === 'editor')
        && state.currentProject && state.generatedSlides.length > 0;
    if (canAddSlide) {
        const addBtn = $(`
            <div class="slide-add-btn" onclick="showTemplateSlidePickerModal()">
                <svg width="24" height="24" fill="none" stroke="currentColor" stroke-width="1.5"><path d="M12 6v12M6 12h12"/></svg>
                <span>${t('addSlide')}</span>
            </div>
        `);
        list.append(addBtn);
    }

    // 슬라이드 카운터 업데이트
    const total = state.generatedSlides.length;
    const current = total > 0 ? state.currentSlideIndex + 1 : 0;
    $('#slideCounterInline').text(current + ' / ' + total);
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
    { const _td64 = getThumbDimensions(64); renderSlideToContainer(thumbEl.find('.slide-thumb-inner'), slide, _td64.w, _td64.h); }
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
        <div class="slide-thumb-v active" tabindex="0" onclick="goToSlide(${i})" data-slide-idx="${i}">
            <div class="slide-thumb-v-num">${i + 1}</div>
            <div class="slide-thumb-v-inner"></div>
        </div>
    `);
    { const _td256 = getThumbDimensions(256); renderSlideToContainer(thumbEl.find('.slide-thumb-v-inner'), slide, _td256.w, _td256.h); }
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

        // 먼저 title 텍스트 추출 (number/governance 역할은 제외)
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
            if (role === 'number' || role === 'governance') return;
            const fontSize = (obj.text_style || {}).font_size || 16;
            if (!titleText && (role === 'title' || role === 'subtitle' || fontSize >= 24)) {
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

                if (role === 'governance' || role === 'number') {
                    // 거버넌스/번호 텍스트는 생략
                } else if (role === 'title' || (!role && fontSize >= 24)) {
                    // 이미 추출됨 (명시적 subtitle 역할은 제외)
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
        { const _td320 = getThumbDimensions(320); renderSlideToContainer(gridEl.find('.grid-thumb-inner'), slide, _td320.w, _td320.h); }
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
    const prevIndex = state.currentSlideIndex;
    // 다른 슬라이드로 이동 시 편집 모드 해제
    if (state.editMode && index !== prevIndex) {
        exitEditMode();
    }
    state.currentSlideIndex = index;
    renderSlideAtIndex(index);
    updateSlideNav();
}

function prevSlide() {
    if (state.currentSlideIndex > 0) {
        _animationCancelled = true;
        if (state.editMode) exitEditMode();
        state.currentSlideIndex--;
        renderSlideAtIndex(state.currentSlideIndex);
        updateSlideNav();
    }
}

function nextSlide() {
    if (state.currentSlideIndex < state.generatedSlides.length - 1) {
        _animationCancelled = true;
        if (state.editMode) exitEditMode();
        state.currentSlideIndex++;
        renderSlideAtIndex(state.currentSlideIndex);
        updateSlideNav();
    }
}

// ============ 콘텐츠 큐 처리 (스켈레톤 → 텍스트 타이핑) ============
async function _processContentQueue() {
    _isAnimating = true;
    _animationCancelled = false;
    switchPanelTab('slide');

    while (state._contentQueue && state._contentQueue.length > 0) {
        if (_animationCancelled) break;

        const idx = state._contentQueue.shift();
        const slide = state.generatedSlides[idx];
        if (!slide) continue;

        state.currentSlideIndex = idx;
        updateSlideNav();

        // 배경 이미지는 스켈레톤 단계에서 이미 프리로드됨 (캐시 히트)
        const imgWaits = [];
        if (slide.background_image) imgWaits.push(_waitForImage(slide.background_image));
        (slide.objects || []).forEach(obj => {
            if (obj.obj_type === 'image' && obj.image_url) imgWaits.push(_waitForImage(obj.image_url));
        });
        if (imgWaits.length > 0) await Promise.all(imgWaits);

        // 슬라이드 렌더링 (텍스트 포함)
        renderSlideAtIndex(idx);

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

            if (!_animationCancelled) {
                $el.html(item.savedHtml);
            }
        }

        // 타이핑 완료 후 해당 슬라이드 썸네일 추가 (한 장씩)
        if (!_animationCancelled) {
            _appendSingleThumbnail(idx);
            _appendSingleThumbV(idx);
            $('#slideCounter').text(`${idx + 1} / ${state.generatedSlides.length}`);
            $('#slideCounterInline').text(`${idx + 1} / ${state.generatedSlides.length}`);
        }

        // 슬라이드 간 전환 딜레이
        if (!_animationCancelled && state._contentQueue.length > 0) {
            await sleep(600);
        }
    }

    _isAnimating = false;

    // 애니메이션 취소 시 아직 추가 안 된 썸네일을 모두 채우기
    if (_animationCancelled) {
        const addedCount = $('#slideThumbList .slide-thumb-v').length;
        for (let j = addedCount; j < state.generatedSlides.length; j++) {
            if (state.generatedSlides[j] && state.generatedSlides[j].items && state.generatedSlides[j].items.length > 0) {
                _appendSingleThumbnail(j);
                _appendSingleThumbV(j);
            }
        }
        $('#slideCounter').text(`${state.currentSlideIndex + 1} / ${state.generatedSlides.length}`);
        $('#slideCounterInline').text(`${state.currentSlideIndex + 1} / ${state.generatedSlides.length}`);
        updateSlideNav();
    }

    // 생성 완료 후 최종 처리
    if (state._generationComplete) {
        state._generationComplete = false;
        state.currentSlideIndex = 0;
        renderSlideAtIndex(0);
        renderSlideThumbnails();
        renderSlideThumbList();
        updateSlideNav();
        renderSlideTextPanel();
        requestAnimationFrame(() => {
            $('#slideThumbList').scrollTop(0);
            $('#slideArea').scrollTop(0);
            $('.slide-canvas-area').scrollTop(0);
        });
        showToast(t('msgAllComplete'), 'success');
    }
}

function _updateThumbnailAtIndex(idx) {
    const slide = state.generatedSlides[idx];
    if (!slide) return;

    // 하단 가로 썸네일 업데이트
    const thumbs = $('#slideThumbnails .slide-thumb');
    if (thumbs.length > idx) {
        const $inner = $(thumbs[idx]).find('.slide-thumb-inner');
        $inner.empty().removeAttr('style');
        { const _td64u = getThumbDimensions(64); renderSlideToContainer($inner, slide, _td64u.w, _td64u.h); }
    }

    // 좌측 세로 썸네일 업데이트
    const $thumbV = $(`#slideThumbList .slide-thumb-v[data-slide-idx="${idx}"]`);
    if ($thumbV.length) {
        const $innerV = $thumbV.find('.slide-thumb-v-inner');
        $innerV.empty().removeAttr('style');
        { const _td256u = getThumbDimensions(256); renderSlideToContainer($innerV, slide, _td256u.w, _td256u.h); }
    }
}

// ============ 타이핑 애니메이션 (레거시) ============
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

        // 배경 + 오브젝트 이미지 로드 완료까지 캔버스 내용 숨김
        const slide = state.generatedSlides[i];
        const imgWaits = [];
        if (slide?.background_image) imgWaits.push(_waitForImage(slide.background_image));
        (slide?.objects || []).forEach(obj => {
            if (obj.obj_type === 'image' && obj.image_url) imgWaits.push(_waitForImage(obj.image_url));
        });

        if (imgWaits.length > 0) {
            // 이미지 로딩 중에는 캔버스 내용을 보이지 않게 처리
            $('#previewCanvas .preview-obj').css('visibility', 'hidden');
            await Promise.all(imgWaits);
        }

        // 이미지 로드 완료 후 슬라이드 렌더링
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

async function enterEditMode() {
    if (state.generatedSlides.length === 0) return;

    // 협업 프로젝트: viewer는 편집 불가
    if (state.isCollabProject && state.collabRole === 'viewer') {
        showToast('뷰어는 편집할 수 없습니다', 'error');
        return;
    }

    // 협업 프로젝트: Lock 획득
    if (state.isCollabProject) {
        const slide = state.generatedSlides[state.currentSlideIndex];
        if (slide && slide._id) {
            const existingLock = state.activeLocks[slide._id];
            if (existingLock && existingLock.user_key !== state.userInfo.ky) {
                showToast(`${existingLock.user_name}이(가) 편집 중입니다`, 'error');
                return;
            }
            try {
                await apiPost('/api/projects/' + state.currentProject._id + '/slides/' + slide._id + '/lock', {});
            } catch (e) {
                showToast(e.message, 'error');
                return;
            }
        }
    }

    // 편집 중인 슬라이드 ID 기록
    const editSlide = state.generatedSlides[state.currentSlideIndex];
    state.editingSlideId = editSlide ? editSlide._id : null;

    state.editMode = true;
    state.editTool = 'select';
    $('#previewCanvas').addClass('edit-mode');
    $('#btnEditToggle').addClass('active');
    $('#btnEditSave').show();
    // 하단 썸네일 숨기고 편집 도구 모음 표시
    $('.slide-nav').hide();
    $('#editBottomToolbar').show();
    // AI 지침 입력 바 유지 (viewer 제외)
    if (state.collabRole !== 'viewer') {
        $('#slideInstructionBar').show();
    }
    _updateEditSlideCounter();
    _populateEditFontSelector();
    renderSlideAtIndexEditable(state.currentSlideIndex);

    // 협업 프로젝트: 편집 중 자동 저장 시작 (3초마다)
    _startEditAutoSave();
}

let _lastAutoSaveSnapshot = null;

function _startEditAutoSave() {
    _stopEditAutoSave();
    if (!state.isCollabProject) return;
    // 초기 스냅샷 저장
    const slide = state.generatedSlides[state.currentSlideIndex];
    _lastAutoSaveSnapshot = slide ? JSON.stringify({ objects: slide.objects, items: slide.items || [] }) : null;

    state.editAutoSaveInterval = setInterval(() => {
        if (!state.editMode || !state.editingSlideId) return;
        collectEditedText();
        const slide = state.generatedSlides[state.currentSlideIndex];
        if (!slide || slide._id !== state.editingSlideId) return;

        // 변경 감지: 이전 스냅샷과 비교하여 변경된 경우만 저장
        const current = JSON.stringify({ objects: slide.objects, items: slide.items || [] });
        if (current === _lastAutoSaveSnapshot) return;
        _lastAutoSaveSnapshot = current;

        fetch(`/${state.jwtToken}/api/generate/slides/${slide._id}`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: current,
        }).catch(() => {});
    }, 5000);
}

function _stopEditAutoSave() {
    if (state.editAutoSaveInterval) {
        clearInterval(state.editAutoSaveInterval);
        state.editAutoSaveInterval = null;
    }
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
    const _sz = getCurrentSlideSize();
    return { x: rect.width / _sz.w, y: rect.height / _sz.h };
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
    const _sz = getCurrentSlideSize();
    const scaleX = canvasW / _sz.w;
    const scaleY = canvasH / _sz.h;

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
            zIndex: obj.z_index !== undefined ? obj.z_index : 10,
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

    // 이동 단축키 힌트 표시
    _showMoveHint(objIdx);
}

function editDeselectAll() {
    state.editSelectedObj = null;
    $('#previewCanvas .preview-obj').removeClass('edit-selected');
    $('#editTextToolbar').hide();
    _removeMoveHint();
}

function _showMoveHint(objIdx) {
    _removeMoveHint();
    const isMac = /Mac|iPhone|iPad/.test(navigator.platform);
    const modKey = isMac ? '⌘' : 'Ctrl';
    const el = $(`#previewCanvas .preview-obj[data-obj-idx="${objIdx}"]`);
    if (el.length === 0) return;
    el.append(`
        <div class="edit-move-hint">
            <span class="edit-move-hint-key">${modKey}</span>
            <span class="edit-move-hint-plus">+</span>
            <span class="edit-move-hint-key">↑↓←→</span>
            <span class="edit-move-hint-label">${t('moveHint', '이동')}</span>
        </div>
    `);
}

function _removeMoveHint() {
    $('.edit-move-hint').remove();
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
    const _szR = getCurrentSlideSize();
    newX = Math.max(0, newX); newY = Math.max(0, newY);
    if (newX + newW > _szR.w) newW = _szR.w - newX;
    if (newY + newH > _szR.h) newH = _szR.h - newY;

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
        const _szD = getCurrentSlideSize();
        x = Math.max(0, Math.min(x, _szD.w - obj.width));
        y = Math.max(0, Math.min(y, _szD.h - obj.height));
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
    } else if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key) && state.editSelectedObj) {
        // Ctrl(Win/Linux) / Cmd(Mac) + 방향키: 텍스트 편집 중이라도 오브젝트 이동
        // 방향키만: 텍스트 편집 중이면 커서 이동으로 동작
        const active = document.activeElement;
        const isEditing = active && active.getAttribute('contenteditable') === 'true';
        const isMoveKey = /Mac|iPhone|iPad/.test(navigator.platform) ? e.metaKey : e.ctrlKey;
        if (isEditing && !isMoveKey) return;
        e.preventDefault();

        const obj = state.editSelectedObj.obj;
        const step = e.shiftKey ? 10 : 1;
        const _sz = getCurrentSlideSize();

        if (e.key === 'ArrowLeft')  obj.x = Math.max(0, obj.x - step);
        if (e.key === 'ArrowRight') obj.x = Math.min(_sz.w - obj.width, obj.x + step);
        if (e.key === 'ArrowUp')    obj.y = Math.max(0, obj.y - step);
        if (e.key === 'ArrowDown')  obj.y = Math.min(_sz.h - obj.height, obj.y + step);

        const scale = editGetScale();
        const el = $(`#previewCanvas .preview-obj[data-obj-idx="${state.editSelectedObj.objIdx}"]`);
        el.css({ left: (obj.x * scale.x) + 'px', top: (obj.y * scale.y) + 'px' });
        markSlideDirty(state.currentSlideIndex);
    }
});

// 좌측 썸네일 키보드 ↑↓ 네비게이션
$('#slideThumbList').on('keydown', '.slide-thumb-v', function (e) {
    if (e.key !== 'ArrowUp' && e.key !== 'ArrowDown') return;
    e.preventDefault();
    const idx = parseInt($(this).attr('data-slide-idx'), 10);
    const next = e.key === 'ArrowUp' ? idx - 1 : idx + 1;
    if (next < 0 || next >= state.generatedSlides.length) return;
    goToSlide(next);
    const $target = $(`#slideThumbList .slide-thumb-v[data-slide-idx="${next}"]`);
    if ($target.length) {
        $target.focus();
        $target[0].scrollIntoView({ block: 'nearest', behavior: 'smooth' });
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

// ---- Shape Catalog ----
var SHAPE_CATEGORIES = [
    { name: '직사각형', shapes: ['rectangle','rounded_rectangle','snip_1_rect','snip_2_diag_rect','round_1_rect','round_2_diag_rect'] },
    { name: '기본 도형', shapes: ['ellipse','triangle','right_triangle','parallelogram','trapezoid','diamond','pentagon','hexagon','heptagon','octagon','decagon','dodecagon','cross','donut','no_smoking','block_arc','heart','lightning_bolt','sun','moon','cloud','smiley_face','folded_corner','frame','teardrop','plaque','brace_pair','bracket_pair'] },
    { name: '블록 화살표', shapes: ['right_arrow','left_arrow','up_arrow','down_arrow','left_right_arrow','up_down_arrow','quad_arrow','notched_right_arrow','chevron','home_plate','striped_right_arrow','bent_arrow','u_turn_arrow','circular_arrow'] },
    { name: '수학', shapes: ['math_plus','math_minus','math_multiply','math_divide','math_equal','math_not_equal'] },
    { name: '별 및 현수막', shapes: ['star_4_point','star_5_point','star_6_point','star_8_point','star_10_point','star_12_point','star_16_point','star_24_point','star_32_point','explosion_1','explosion_2','wave','double_wave','ribbon'] },
    { name: '설명선', shapes: ['wedge_rect_callout','wedge_round_rect_callout','wedge_ellipse_callout','cloud_callout','border_callout_1','border_callout_2','border_callout_3'] },
    { name: '선', shapes: ['line','arrow'] },
];

function _regPolygonPts(cx, cy, r, n, rotDeg) {
    var pts = [];
    var startAngle = (rotDeg || -90) * Math.PI / 180;
    for (var i = 0; i < n; i++) {
        var a = startAngle + (2 * Math.PI * i / n);
        pts.push((cx + r * Math.cos(a)).toFixed(1) + ',' + (cy + r * Math.sin(a)).toFixed(1));
    }
    return pts.join(' ');
}

function _starPts(cx, cy, outerR, innerR, points) {
    var pts = [];
    var startAngle = -Math.PI / 2;
    for (var i = 0; i < points * 2; i++) {
        var a = startAngle + (Math.PI * i / points);
        var r = (i % 2 === 0) ? outerR : innerR;
        pts.push((cx + r * Math.cos(a)).toFixed(1) + ',' + (cy + r * Math.sin(a)).toFixed(1));
    }
    return pts.join(' ');
}

function _getShapeIconSVG(type) {
    var s = 'stroke="#444" stroke-width="1.2" fill="none"';
    switch(type) {
        // 직사각형
        case 'rectangle': return '<svg viewBox="0 0 24 24"><rect x="3" y="5" width="18" height="14" '+s+'/></svg>';
        case 'rounded_rectangle': return '<svg viewBox="0 0 24 24"><rect x="3" y="5" width="18" height="14" rx="4" '+s+'/></svg>';
        case 'snip_1_rect': return '<svg viewBox="0 0 24 24"><path d="M3 5h14l4 4v10H3z" '+s+'/></svg>';
        case 'snip_2_diag_rect': return '<svg viewBox="0 0 24 24"><path d="M7 5h14v10l-4 4H3V9z" '+s+'/></svg>';
        case 'round_1_rect': return '<svg viewBox="0 0 24 24"><path d="M3 5h14a4 4 0 014 4v10H3z" '+s+'/></svg>';
        case 'round_2_diag_rect': return '<svg viewBox="0 0 24 24"><path d="M7 5h14v10a4 4 0 01-4 4H3V9a4 4 0 014-4z" '+s+'/></svg>';
        // 기본 도형
        case 'ellipse': return '<svg viewBox="0 0 24 24"><ellipse cx="12" cy="12" rx="9" ry="7" '+s+'/></svg>';
        case 'triangle': return '<svg viewBox="0 0 24 24"><polygon points="12,3 22,21 2,21" '+s+'/></svg>';
        case 'right_triangle': return '<svg viewBox="0 0 24 24"><polygon points="3,21 3,3 21,21" '+s+'/></svg>';
        case 'parallelogram': return '<svg viewBox="0 0 24 24"><polygon points="7,5 22,5 17,19 2,19" '+s+'/></svg>';
        case 'trapezoid': return '<svg viewBox="0 0 24 24"><polygon points="6,5 18,5 22,19 2,19" '+s+'/></svg>';
        case 'diamond': return '<svg viewBox="0 0 24 24"><polygon points="12,2 22,12 12,22 2,12" '+s+'/></svg>';
        case 'pentagon': return '<svg viewBox="0 0 24 24"><polygon points="'+_regPolygonPts(12,12,10,5)+'" '+s+'/></svg>';
        case 'hexagon': return '<svg viewBox="0 0 24 24"><polygon points="'+_regPolygonPts(12,12,10,6)+'" '+s+'/></svg>';
        case 'heptagon': return '<svg viewBox="0 0 24 24"><polygon points="'+_regPolygonPts(12,12,10,7)+'" '+s+'/></svg>';
        case 'octagon': return '<svg viewBox="0 0 24 24"><polygon points="'+_regPolygonPts(12,12,10,8)+'" '+s+'/></svg>';
        case 'decagon': return '<svg viewBox="0 0 24 24"><polygon points="'+_regPolygonPts(12,12,10,10)+'" '+s+'/></svg>';
        case 'dodecagon': return '<svg viewBox="0 0 24 24"><polygon points="'+_regPolygonPts(12,12,10,12)+'" '+s+'/></svg>';
        case 'cross': return '<svg viewBox="0 0 24 24"><polygon points="8,2 16,2 16,8 22,8 22,16 16,16 16,22 8,22 8,16 2,16 2,8 8,8" '+s+'/></svg>';
        case 'donut': return '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" '+s+'/><circle cx="12" cy="12" r="5" '+s+'/></svg>';
        case 'no_smoking': return '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" '+s+'/><line x1="5" y1="5" x2="19" y2="19" '+s+'/></svg>';
        case 'block_arc': return '<svg viewBox="0 0 24 24"><path d="M4 20A10 10 0 0122 12h-4a6 6 0 00-12 0v8z" '+s+'/></svg>';
        case 'heart': return '<svg viewBox="0 0 24 24"><path d="M12 21C12 21 3 14 3 8.5C3 5.4 5.4 3 8.5 3c1.7 0 3.3.8 3.5 2 .2-1.2 1.8-2 3.5-2C18.6 3 21 5.4 21 8.5 21 14 12 21 12 21z" '+s+'/></svg>';
        case 'lightning_bolt': return '<svg viewBox="0 0 24 24"><polygon points="13,2 4,14 11,14 10,22 20,10 13,10" '+s+'/></svg>';
        case 'sun': return '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="5" '+s+'/><line x1="12" y1="2" x2="12" y2="5" '+s+'/><line x1="12" y1="19" x2="12" y2="22" '+s+'/><line x1="2" y1="12" x2="5" y2="12" '+s+'/><line x1="19" y1="12" x2="22" y2="12" '+s+'/><line x1="4.9" y1="4.9" x2="6.8" y2="6.8" '+s+'/><line x1="17.2" y1="17.2" x2="19.1" y2="19.1" '+s+'/><line x1="4.9" y1="19.1" x2="6.8" y2="17.2" '+s+'/><line x1="17.2" y1="6.8" x2="19.1" y2="4.9" '+s+'/></svg>';
        case 'moon': return '<svg viewBox="0 0 24 24"><path d="M21 12.79A9 9 0 1111.21 3 7 7 0 0021 12.79z" '+s+'/></svg>';
        case 'cloud': return '<svg viewBox="0 0 24 24"><path d="M6 19a4 4 0 01-.8-7.9A5.5 5.5 0 0117 9.6 3.5 3.5 0 0120 13a3.5 3.5 0 01-3.5 3.5H6z" '+s+'/></svg>';
        case 'smiley_face': return '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="10" '+s+'/><path d="M8 14s1.5 2 4 2 4-2 4-2" '+s+'/><circle cx="9" cy="9" r="1" fill="#444" stroke="none"/><circle cx="15" cy="9" r="1" fill="#444" stroke="none"/></svg>';
        case 'folded_corner': return '<svg viewBox="0 0 24 24"><path d="M3 3h18v13l-5 5H3z" '+s+'/><path d="M16 16v5l5-5z" '+s+'/></svg>';
        case 'frame': return '<svg viewBox="0 0 24 24"><rect x="2" y="2" width="20" height="20" '+s+'/><rect x="5" y="5" width="14" height="14" '+s+'/></svg>';
        case 'teardrop': return '<svg viewBox="0 0 24 24"><path d="M12 3C12 3 5 11 5 15a7 7 0 0014 0c0-4-7-12-7-12z" '+s+'/></svg>';
        case 'plaque': return '<svg viewBox="0 0 24 24"><path d="M4 4c2 2 2 4 0 6s-2 4 0 6c0 2 2 4 4 4 2-2 4-2 6 0s4 2 6 0c2 0 4-2 4-4-2-2-2-4 0-6s2-4 0-6c0-2-2-4-4-4-2 2-4 2-6 0s-4-2-6 0C6 0 4 2 4 4z" '+s+'/></svg>';
        case 'brace_pair': return '<svg viewBox="0 0 24 24"><path d="M8 3c-3 0-4 1-4 4v2c0 2-1 3-2 3 1 0 2 1 2 3v2c0 3 1 4 4 4" '+s+'/><path d="M16 3c3 0 4 1 4 4v2c0 2 1 3 2 3-1 0-2 1-2 3v2c0 3-1 4-4 4" '+s+'/></svg>';
        case 'bracket_pair': return '<svg viewBox="0 0 24 24"><path d="M8 3H5v18h3" '+s+'/><path d="M16 3h3v18h-3" '+s+'/></svg>';
        // 블록 화살표
        case 'right_arrow': return '<svg viewBox="0 0 24 24"><polygon points="2,7 15,7 15,3 22,12 15,21 15,17 2,17" '+s+'/></svg>';
        case 'left_arrow': return '<svg viewBox="0 0 24 24"><polygon points="22,7 9,7 9,3 2,12 9,21 9,17 22,17" '+s+'/></svg>';
        case 'up_arrow': return '<svg viewBox="0 0 24 24"><polygon points="7,22 7,9 3,9 12,2 21,9 17,9 17,22" '+s+'/></svg>';
        case 'down_arrow': return '<svg viewBox="0 0 24 24"><polygon points="7,2 7,15 3,15 12,22 21,15 17,15 17,2" '+s+'/></svg>';
        case 'left_right_arrow': return '<svg viewBox="0 0 24 24"><polygon points="7,3 7,7 17,7 17,3 22,12 17,21 17,17 7,17 7,21 2,12" '+s+'/></svg>';
        case 'up_down_arrow': return '<svg viewBox="0 0 24 24"><polygon points="3,7 12,2 21,7 17,7 17,17 21,17 12,22 3,17 7,17 7,7" '+s+'/></svg>';
        case 'quad_arrow': return '<svg viewBox="0 0 24 24"><polygon points="12,1 16,5 14,5 14,10 19,10 19,8 23,12 19,16 19,14 14,14 14,19 16,19 12,23 8,19 10,19 10,14 5,14 5,16 1,12 5,8 5,10 10,10 10,5 8,5" '+s+'/></svg>';
        case 'notched_right_arrow': return '<svg viewBox="0 0 24 24"><polygon points="2,7 15,7 15,3 22,12 15,21 15,17 2,17 5,12" '+s+'/></svg>';
        case 'chevron': return '<svg viewBox="0 0 24 24"><polygon points="2,3 17,3 22,12 17,21 2,21 7,12" '+s+'/></svg>';
        case 'home_plate': return '<svg viewBox="0 0 24 24"><polygon points="2,3 17,3 22,12 17,21 2,21" '+s+'/></svg>';
        case 'striped_right_arrow': return '<svg viewBox="0 0 24 24"><polygon points="8,7 15,7 15,3 22,12 15,21 15,17 8,17" '+s+'/><line x1="4" y1="7" x2="4" y2="17" '+s+'/><line x1="6" y1="7" x2="6" y2="17" '+s+'/></svg>';
        case 'bent_arrow': return '<svg viewBox="0 0 24 24"><path d="M3 21V10h9V5l6 7-6 7v-5H7v7z" '+s+'/></svg>';
        case 'u_turn_arrow': return '<svg viewBox="0 0 24 24"><path d="M6 21V10a6 6 0 0112 0v4h3l-5 6-5-6h3v-4a2 2 0 00-4 0v11z" '+s+'/></svg>';
        case 'circular_arrow': return '<svg viewBox="0 0 24 24"><path d="M12 4a8 8 0 017.6 5.5" '+s+'/><path d="M20 4v6h-6" '+s+'/><path d="M12 20a8 8 0 01-7.6-5.5" '+s+'/></svg>';
        // 수학
        case 'math_plus': return '<svg viewBox="0 0 24 24"><line x1="12" y1="4" x2="12" y2="20" '+s+'/><line x1="4" y1="12" x2="20" y2="12" '+s+'/></svg>';
        case 'math_minus': return '<svg viewBox="0 0 24 24"><line x1="4" y1="12" x2="20" y2="12" '+s+'/></svg>';
        case 'math_multiply': return '<svg viewBox="0 0 24 24"><line x1="5" y1="5" x2="19" y2="19" '+s+'/><line x1="19" y1="5" x2="5" y2="19" '+s+'/></svg>';
        case 'math_divide': return '<svg viewBox="0 0 24 24"><line x1="4" y1="12" x2="20" y2="12" '+s+'/><circle cx="12" cy="7" r="1.5" fill="#444" stroke="none"/><circle cx="12" cy="17" r="1.5" fill="#444" stroke="none"/></svg>';
        case 'math_equal': return '<svg viewBox="0 0 24 24"><line x1="4" y1="9" x2="20" y2="9" '+s+'/><line x1="4" y1="15" x2="20" y2="15" '+s+'/></svg>';
        case 'math_not_equal': return '<svg viewBox="0 0 24 24"><line x1="4" y1="9" x2="20" y2="9" '+s+'/><line x1="4" y1="15" x2="20" y2="15" '+s+'/><line x1="16" y1="4" x2="8" y2="20" '+s+'/></svg>';
        // 별 및 현수막
        case 'star_4_point': return '<svg viewBox="0 0 24 24"><polygon points="'+_starPts(12,12,10,3.8,4)+'" '+s+'/></svg>';
        case 'star_5_point': return '<svg viewBox="0 0 24 24"><polygon points="'+_starPts(12,12,10,3.8,5)+'" '+s+'/></svg>';
        case 'star_6_point': return '<svg viewBox="0 0 24 24"><polygon points="'+_starPts(12,12,10,5,6)+'" '+s+'/></svg>';
        case 'star_8_point': return '<svg viewBox="0 0 24 24"><polygon points="'+_starPts(12,12,10,5,8)+'" '+s+'/></svg>';
        case 'star_10_point': return '<svg viewBox="0 0 24 24"><polygon points="'+_starPts(12,12,10,6,10)+'" '+s+'/></svg>';
        case 'star_12_point': return '<svg viewBox="0 0 24 24"><polygon points="'+_starPts(12,12,10,6,12)+'" '+s+'/></svg>';
        case 'star_16_point': return '<svg viewBox="0 0 24 24"><polygon points="'+_starPts(12,12,10,7,16)+'" '+s+'/></svg>';
        case 'star_24_point': return '<svg viewBox="0 0 24 24"><polygon points="'+_starPts(12,12,10,7.5,24)+'" '+s+'/></svg>';
        case 'star_32_point': return '<svg viewBox="0 0 24 24"><polygon points="'+_starPts(12,12,10,8,32)+'" '+s+'/></svg>';
        case 'explosion_1': return '<svg viewBox="0 0 24 24"><polygon points="12,2 14,8 20,4 16,10 23,12 16,14 20,20 14,16 12,22 10,16 4,20 8,14 1,12 8,10 4,4 10,8" '+s+'/></svg>';
        case 'explosion_2': return '<svg viewBox="0 0 24 24"><polygon points="12,1 13,7 18,3 15,9 22,8 17,12 23,16 16,15 18,21 13,17 12,23 11,17 6,21 8,15 1,16 7,12 2,8 9,9 6,3 11,7" '+s+'/></svg>';
        case 'wave': return '<svg viewBox="0 0 24 24"><path d="M2 8c3-4 6-4 10 0s7 4 10 0v10c-3 4-6 4-10 0s-7-4-10 0z" '+s+'/></svg>';
        case 'double_wave': return '<svg viewBox="0 0 24 24"><path d="M2 6c3-3 6-3 10 0s7 3 10 0" '+s+'/><path d="M2 12c3-3 6-3 10 0s7 3 10 0" '+s+'/><path d="M2 18c3-3 6-3 10 0s7 3 10 0" '+s+'/></svg>';
        case 'ribbon': return '<svg viewBox="0 0 24 24"><path d="M2 7h20v10H2z" '+s+'/><path d="M5 7V4l-3 3M19 7V4l3 3M5 17v3l-3-3M19 17v3l3-3" '+s+'/></svg>';
        // 설명선
        case 'wedge_rect_callout': return '<svg viewBox="0 0 24 24"><path d="M2 3h20v13H14l-2 5-2-5H2z" '+s+'/></svg>';
        case 'wedge_round_rect_callout': return '<svg viewBox="0 0 24 24"><path d="M5 3h14a3 3 0 013 3v7a3 3 0 01-3 3H14l-2 5-2-5H5a3 3 0 01-3-3V6a3 3 0 013-3z" '+s+'/></svg>';
        case 'wedge_ellipse_callout': return '<svg viewBox="0 0 24 24"><ellipse cx="12" cy="9" rx="10" ry="7" '+s+'/><path d="M10 15l-1 6 4-5" '+s+'/></svg>';
        case 'cloud_callout': return '<svg viewBox="0 0 24 24"><path d="M6 16a4 4 0 01-.8-7.9A5.5 5.5 0 0117 6.6 3.5 3.5 0 0120 10a3.5 3.5 0 01-3.5 3.5H6z" '+s+'/><circle cx="8" cy="19" r="1.5" '+s+'/><circle cx="5" cy="21" r="1" '+s+'/></svg>';
        case 'border_callout_1': return '<svg viewBox="0 0 24 24"><rect x="2" y="2" width="20" height="14" '+s+'/><line x1="8" y1="16" x2="6" y2="22" '+s+'/></svg>';
        case 'border_callout_2': return '<svg viewBox="0 0 24 24"><rect x="2" y="2" width="20" height="14" '+s+'/><polyline points="8,16 10,19 6,22" '+s+'/></svg>';
        case 'border_callout_3': return '<svg viewBox="0 0 24 24"><rect x="2" y="2" width="20" height="14" '+s+'/><polyline points="8,16 12,18 8,20 6,22" '+s+'/></svg>';
        // 선
        case 'line': return '<svg viewBox="0 0 24 24"><line x1="3" y1="19" x2="21" y2="5" '+s+'/></svg>';
        case 'arrow': return '<svg viewBox="0 0 24 24"><line x1="3" y1="12" x2="18" y2="12" '+s+'/><path d="M15 8l4 4-4 4" '+s+'/></svg>';
        default: return '<svg viewBox="0 0 24 24"><rect x="3" y="5" width="18" height="14" '+s+'/></svg>';
    }
}

function _buildShapePickerHTML() {
    var html = '';
    SHAPE_CATEGORIES.forEach(function(cat) {
        html += '<div class="shape-picker-category">' + cat.name + '</div>';
        html += '<div class="shape-picker-grid">';
        cat.shapes.forEach(function(shapeType) {
            html += '<button class="shape-picker-item" onclick="insertEditShape(\'' + shapeType + '\')" title="' + shapeType + '">';
            html += _getShapeIconSVG(shapeType);
            html += '</button>';
        });
        html += '</div>';
    });
    return html;
}

// ============ 편집 모드 격자 ============
const EDIT_GRID_LEVELS = [
    { size: 0,  cls: '',        label: 'OFF' },
    { size: 50, cls: 'grid-50', label: '50' },
    { size: 25, cls: 'grid-25', label: '25' },
    { size: 12, cls: 'grid-12', label: '12' },
];
var _editGridLevel = 0;

function toggleEditGrid() {
    _editGridLevel = (_editGridLevel + 1) % EDIT_GRID_LEVELS.length;
    const level = EDIT_GRID_LEVELS[_editGridLevel];
    const grid = $('#editCanvasGrid');
    const btn = $('#btnEditGrid');

    grid.removeClass('visible grid-50 grid-25 grid-12');
    if (level.size > 0) {
        grid.addClass('visible ' + level.cls);
        btn.addClass('grid-on');
    } else {
        btn.removeClass('grid-on');
    }
    btn.find('.edit-grid-label').text(level.label);
}

var _shapePickerInitialized = false;

function toggleEditShapeMenu() {
    var $menu = $('#editShapeMenu');
    if (!_shapePickerInitialized) {
        $menu.html(_buildShapePickerHTML());
        _shapePickerInitialized = true;
    }
    $menu.toggle();
    if ($menu.is(':visible')) {
        setTimeout(function() {
            $(document).one('click', function(e) {
                if (!$(e.target).closest('#editShapeMenu, #editToolShape').length) {
                    $menu.hide();
                }
            });
        }, 0);
    }
}

function insertEditShape(shapeType) {
    $('#editShapeMenu').hide();
    const slide = state.generatedSlides[state.currentSlideIndex];
    if (!slide) return;
    var isLine = (shapeType === 'line' || shapeType === 'arrow');
    var isBlockArrow = ['right_arrow','left_arrow','up_arrow','down_arrow','left_right_arrow','up_down_arrow','quad_arrow','notched_right_arrow','chevron','home_plate','striped_right_arrow','bent_arrow','u_turn_arrow','circular_arrow'].indexOf(shapeType) >= 0;
    const newObj = {
        obj_id: 'obj_' + Date.now(),
        obj_type: 'shape',
        x: 330, y: 180, width: isBlockArrow ? 300 : 300, height: isLine ? 4 : (isBlockArrow ? 180 : 180),
        shape_style: {
            shape_type: shapeType,
            fill_color: isLine ? 'transparent' : '#4A90D9',
            fill_opacity: isLine ? 0 : 0.2,
            stroke_color: '#2d5a8e',
            stroke_width: 2,
            stroke_dash: 'solid',
            border_radius: shapeType === 'rounded_rectangle' ? 12 : 0,
            arrow_head: shapeType === 'arrow' ? 'end' : 'none',
        },
    };
    if (isLine) {
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
    const fa = `fill="${fill}" fill-opacity="${opacity}" stroke="${stroke}" stroke-width="${sw}" ${dash}`;
    const cx = w / 2, cy = h / 2;
    let inner = '';

    switch(st) {
        case 'ellipse':
            inner = `<ellipse cx="${cx}" cy="${cy}" rx="${cx - half}" ry="${cy - half}" ${fa}/>`;
            break;
        case 'rounded_rectangle': {
            const r = s.border_radius || 12;
            inner = `<rect x="${half}" y="${half}" width="${w - sw}" height="${h - sw}" rx="${r}" ${fa}/>`;
            break;
        }
        case 'snip_1_rect': {
            const sn = Math.min(w, h) * 0.2;
            inner = `<path d="M${half} ${half}h${w - sw - sn}l${sn} ${sn}v${h - sw - sn}H${half}z" ${fa}/>`;
            break;
        }
        case 'snip_2_diag_rect': {
            const sn = Math.min(w, h) * 0.2;
            inner = `<path d="M${half + sn} ${half}h${w - sw - sn}v${h - sw - sn}l-${sn} ${sn}H${half}V${half + sn}z" ${fa}/>`;
            break;
        }
        case 'round_1_rect': {
            const rr = Math.min(w, h) * 0.25;
            inner = `<path d="M${half} ${half}h${w - sw - rr}a${rr} ${rr} 0 01${rr} ${rr}v${h - sw - rr}H${half}z" ${fa}/>`;
            break;
        }
        case 'round_2_diag_rect': {
            const rr = Math.min(w, h) * 0.25;
            inner = `<path d="M${half + rr} ${half}h${w - sw - rr}v${h - sw - rr}a${rr} ${rr} 0 01-${rr} ${rr}H${half}V${half + rr}a${rr} ${rr} 0 01${rr}-${rr}z" ${fa}/>`;
            break;
        }
        case 'triangle':
            inner = `<polygon points="${cx},${half} ${w - half},${h - half} ${half},${h - half}" ${fa}/>`;
            break;
        case 'right_triangle':
            inner = `<polygon points="${half},${h - half} ${half},${half} ${w - half},${h - half}" ${fa}/>`;
            break;
        case 'parallelogram': {
            const off = w * 0.2;
            inner = `<polygon points="${off},${half} ${w - half},${half} ${w - off},${h - half} ${half},${h - half}" ${fa}/>`;
            break;
        }
        case 'trapezoid': {
            const off = w * 0.15;
            inner = `<polygon points="${off},${half} ${w - off},${half} ${w - half},${h - half} ${half},${h - half}" ${fa}/>`;
            break;
        }
        case 'diamond':
            inner = `<polygon points="${cx},${half} ${w - half},${cy} ${cx},${h - half} ${half},${cy}" ${fa}/>`;
            break;
        case 'pentagon':
            inner = `<polygon points="${_regPolygonPts(cx, cy, Math.min(cx, cy) - half, 5)}" ${fa}/>`;
            break;
        case 'hexagon':
            inner = `<polygon points="${_regPolygonPts(cx, cy, Math.min(cx, cy) - half, 6)}" ${fa}/>`;
            break;
        case 'heptagon':
            inner = `<polygon points="${_regPolygonPts(cx, cy, Math.min(cx, cy) - half, 7)}" ${fa}/>`;
            break;
        case 'octagon':
            inner = `<polygon points="${_regPolygonPts(cx, cy, Math.min(cx, cy) - half, 8)}" ${fa}/>`;
            break;
        case 'decagon':
            inner = `<polygon points="${_regPolygonPts(cx, cy, Math.min(cx, cy) - half, 10)}" ${fa}/>`;
            break;
        case 'dodecagon':
            inner = `<polygon points="${_regPolygonPts(cx, cy, Math.min(cx, cy) - half, 12)}" ${fa}/>`;
            break;
        case 'cross': {
            const t3 = w / 3, t3h = h / 3;
            inner = `<polygon points="${t3},${half} ${2*t3},${half} ${2*t3},${t3h} ${w-half},${t3h} ${w-half},${2*t3h} ${2*t3},${2*t3h} ${2*t3},${h-half} ${t3},${h-half} ${t3},${2*t3h} ${half},${2*t3h} ${half},${t3h} ${t3},${t3h}" ${fa}/>`;
            break;
        }
        case 'donut': {
            const or = Math.min(cx, cy) - half;
            const ir = or * 0.5;
            inner = `<circle cx="${cx}" cy="${cy}" r="${or}" ${fa}/><circle cx="${cx}" cy="${cy}" r="${ir}" fill="white" fill-opacity="1" stroke="${stroke}" stroke-width="${sw}" ${dash}/>`;
            break;
        }
        case 'no_smoking': {
            const r = Math.min(cx, cy) - half;
            inner = `<circle cx="${cx}" cy="${cy}" r="${r}" ${fa}/><line x1="${cx - r * 0.7}" y1="${cy - r * 0.7}" x2="${cx + r * 0.7}" y2="${cy + r * 0.7}" stroke="${stroke}" stroke-width="${sw}" ${dash}/>`;
            break;
        }
        case 'block_arc': {
            const r = Math.min(w, h) * 0.45;
            inner = `<path d="M${half} ${h - half}A${r} ${r} 0 01${w - half} ${cy}h-${w * 0.2}a${r * 0.6} ${r * 0.6} 0 00-${w * 0.6 - sw} ${cy - half - (h - half - cy)}v${h * 0.35}z" ${fa}/>`;
            break;
        }
        case 'heart': {
            const hw = w * 0.5, hh = h * 0.3;
            inner = `<path d="M${cx} ${h * 0.85}C${cx} ${h * 0.85} ${half} ${h * 0.55} ${half} ${hh + half}C${half} ${half} ${cx * 0.5} ${half} ${cx} ${h * 0.35}C${cx + cx * 0.5} ${half} ${w - half} ${half} ${w - half} ${hh + half}C${w - half} ${h * 0.55} ${cx} ${h * 0.85} ${cx} ${h * 0.85}z" ${fa}/>`;
            break;
        }
        case 'lightning_bolt':
            inner = `<polygon points="${w*0.55},${half} ${w*0.2},${cy} ${w*0.5},${cy} ${w*0.4},${h-half} ${w*0.85},${cy} ${w*0.55},${cy}" ${fa}/>`;
            break;
        case 'sun': {
            const sr = Math.min(cx, cy) * 0.4;
            const or2 = Math.min(cx, cy) - half;
            inner = `<circle cx="${cx}" cy="${cy}" r="${sr}" ${fa}/>`;
            for (var i = 0; i < 12; i++) {
                var a = Math.PI * 2 * i / 12;
                inner += `<line x1="${cx + sr * 1.2 * Math.cos(a)}" y1="${cy + sr * 1.2 * Math.sin(a)}" x2="${cx + or2 * Math.cos(a)}" y2="${cy + or2 * Math.sin(a)}" stroke="${stroke}" stroke-width="${sw}" ${dash}/>`;
            }
            break;
        }
        case 'moon':
            inner = `<path d="M${w * 0.7} ${half}a${cy} ${cy} 0 10 0 ${h - sw}A${cx * 0.7} ${cy} 0 01${w * 0.7} ${half}z" ${fa}/>`;
            break;
        case 'cloud':
            inner = `<path d="M${w*0.25} ${h*0.7}a${w*0.15} ${h*0.15} 0 01${w*0.02}-${h*0.3} ${w*0.2} ${h*0.2} 0 01${w*0.38}-${h*0.12} ${w*0.15} ${h*0.12} 0 01${w*0.28} ${h*0.05} ${w*0.12} ${h*0.12} 0 01${w*0.02} ${h*0.22}z" ${fa}/>`;
            break;
        case 'smiley_face': {
            const fr = Math.min(cx, cy) - half;
            inner = `<circle cx="${cx}" cy="${cy}" r="${fr}" ${fa}/>`;
            inner += `<circle cx="${cx - fr * 0.3}" cy="${cy - fr * 0.2}" r="${fr * 0.08}" fill="${stroke}" stroke="none"/>`;
            inner += `<circle cx="${cx + fr * 0.3}" cy="${cy - fr * 0.2}" r="${fr * 0.08}" fill="${stroke}" stroke="none"/>`;
            inner += `<path d="M${cx - fr * 0.35} ${cy + fr * 0.15}q${fr * 0.35} ${fr * 0.4} ${fr * 0.7} 0" fill="none" stroke="${stroke}" stroke-width="${sw}"/>`;
            break;
        }
        case 'folded_corner': {
            const fc = Math.min(w, h) * 0.2;
            inner = `<path d="M${half} ${half}h${w - sw}v${h - sw - fc}l-${fc} ${fc}H${half}z" ${fa}/>`;
            inner += `<path d="M${w - half - fc} ${h - half}v-${fc}h${fc}z" fill="none" stroke="${stroke}" stroke-width="${sw}"/>`;
            break;
        }
        case 'frame': {
            const ft = Math.min(w, h) * 0.12;
            inner = `<rect x="${half}" y="${half}" width="${w - sw}" height="${h - sw}" ${fa}/>`;
            inner += `<rect x="${ft}" y="${ft}" width="${w - ft * 2}" height="${h - ft * 2}" fill="white" fill-opacity="1" stroke="${stroke}" stroke-width="${sw}"/>`;
            break;
        }
        case 'teardrop':
            inner = `<path d="M${cx} ${half}C${cx} ${half} ${half} ${cy * 0.8} ${half} ${cy + cy * 0.3}a${cx - half} ${cy * 0.6} 0 00${w - sw} 0C${w - half} ${cy * 0.8} ${cx} ${half} ${cx} ${half}z" ${fa}/>`;
            break;
        case 'plaque': {
            const pr = Math.min(w, h) * 0.15;
            inner = `<path d="M${half + pr} ${half}h${w - sw - 2 * pr}a${pr} ${pr} 0 00${pr} ${pr}v${h - sw - 2 * pr}a${pr} ${pr} 0 00-${pr} ${pr}H${half + pr}a${pr} ${pr} 0 00-${pr}-${pr}V${half + pr}a${pr} ${pr} 0 00${pr}-${pr}z" ${fa}/>`;
            break;
        }
        case 'brace_pair': {
            const bw = w * 0.1;
            inner = `<path d="M${bw + half} ${half}c-${bw} 0-${bw} ${h * 0.1}-${bw} ${h * 0.25}s0 ${h * 0.15}-${bw * 0.5} ${h * 0.25} ${bw * 0.5} ${h * 0.1} ${bw * 0.5} ${h * 0.25}s0 ${h * 0.15} ${bw} ${h * 0.25}" fill="none" stroke="${stroke}" stroke-width="${sw}" ${dash}/>`;
            inner += `<path d="M${w - bw - half} ${half}c${bw} 0 ${bw} ${h * 0.1} ${bw} ${h * 0.25}s0 ${h * 0.15} ${bw * 0.5} ${h * 0.25}-${bw * 0.5} ${h * 0.1}-${bw * 0.5} ${h * 0.25}s0 ${h * 0.15}-${bw} ${h * 0.25}" fill="none" stroke="${stroke}" stroke-width="${sw}" ${dash}/>`;
            break;
        }
        case 'bracket_pair': {
            const bkw = w * 0.1;
            inner = `<path d="M${bkw + half} ${half}H${half}v${h - sw}h${bkw}" fill="none" stroke="${stroke}" stroke-width="${sw}" ${dash}/>`;
            inner += `<path d="M${w - bkw - half} ${half}H${w - half}v${h - sw}h-${bkw}" fill="none" stroke="${stroke}" stroke-width="${sw}" ${dash}/>`;
            break;
        }
        // 블록 화살표
        case 'right_arrow': {
            const sh = h * 0.3, ah = h * 0.5, aw = w * 0.65;
            inner = `<polygon points="${half},${sh} ${aw},${sh} ${aw},${half} ${w-half},${cy} ${aw},${h-half} ${aw},${h-sh} ${half},${h-sh}" ${fa}/>`;
            break;
        }
        case 'left_arrow': {
            const sh = h * 0.3, aw = w * 0.35;
            inner = `<polygon points="${w-half},${sh} ${aw},${sh} ${aw},${half} ${half},${cy} ${aw},${h-half} ${aw},${h-sh} ${w-half},${h-sh}" ${fa}/>`;
            break;
        }
        case 'up_arrow': {
            const sw2 = w * 0.3, ah2 = h * 0.35;
            inner = `<polygon points="${sw2},${h-half} ${sw2},${ah2} ${half},${ah2} ${cx},${half} ${w-half},${ah2} ${w-sw2},${ah2} ${w-sw2},${h-half}" ${fa}/>`;
            break;
        }
        case 'down_arrow': {
            const sw2 = w * 0.3, ah2 = h * 0.65;
            inner = `<polygon points="${sw2},${half} ${sw2},${ah2} ${half},${ah2} ${cx},${h-half} ${w-half},${ah2} ${w-sw2},${ah2} ${w-sw2},${half}" ${fa}/>`;
            break;
        }
        case 'left_right_arrow': {
            const sh3 = h * 0.3, aw3 = w * 0.25;
            inner = `<polygon points="${half},${cy} ${aw3},${half} ${aw3},${sh3} ${w-aw3},${sh3} ${w-aw3},${half} ${w-half},${cy} ${w-aw3},${h-half} ${w-aw3},${h-sh3} ${aw3},${h-sh3} ${aw3},${h-half}" ${fa}/>`;
            break;
        }
        case 'up_down_arrow': {
            const sw3 = w * 0.3, ah3 = h * 0.25;
            inner = `<polygon points="${cx},${half} ${w-half},${ah3} ${w-sw3},${ah3} ${w-sw3},${h-ah3} ${w-half},${h-ah3} ${cx},${h-half} ${half},${h-ah3} ${sw3},${h-ah3} ${sw3},${ah3} ${half},${ah3}" ${fa}/>`;
            break;
        }
        case 'quad_arrow': {
            const qs = Math.min(w, h) * 0.15, qe = Math.min(w, h) * 0.35;
            inner = `<polygon points="${cx},${half} ${cx+qe},${qe} ${cx+qs},${qe} ${cx+qs},${cy-qs} ${cx+qe},${cy-qs} ${cx+qe},${cy-qe} ${w-half},${cy} ${cx+qe},${cy+qe} ${cx+qe},${cy+qs} ${cx+qs},${cy+qs} ${cx+qs},${cy+qe} ${cx+qe},${cy+qe} ${cx},${h-half} ${cx-qe},${cy+qe} ${cx-qs},${cy+qe} ${cx-qs},${cy+qs} ${cx-qe},${cy+qs} ${cx-qe},${cy+qe} ${half},${cy} ${cx-qe},${cy-qe} ${cx-qe},${cy-qs} ${cx-qs},${cy-qs} ${cx-qs},${cy-qe} ${cx-qe},${cy-qe}" ${fa}/>`;
            break;
        }
        case 'notched_right_arrow': {
            const sh4 = h * 0.3, aw4 = w * 0.65, notch = w * 0.1;
            inner = `<polygon points="${half},${sh4} ${aw4},${sh4} ${aw4},${half} ${w-half},${cy} ${aw4},${h-half} ${aw4},${h-sh4} ${half},${h-sh4} ${notch},${cy}" ${fa}/>`;
            break;
        }
        case 'chevron': {
            const cv = w * 0.25;
            inner = `<polygon points="${half},${half} ${w - cv},${half} ${w - half},${cy} ${w - cv},${h - half} ${half},${h - half} ${cv},${cy}" ${fa}/>`;
            break;
        }
        case 'home_plate': {
            const hp = w * 0.2;
            inner = `<polygon points="${half},${half} ${w - hp},${half} ${w - half},${cy} ${w - hp},${h - half} ${half},${h - half}" ${fa}/>`;
            break;
        }
        case 'striped_right_arrow': {
            const sh5 = h * 0.3, aw5 = w * 0.65, strL = w * 0.12;
            inner = `<polygon points="${strL + w * 0.05},${sh5} ${aw5},${sh5} ${aw5},${half} ${w-half},${cy} ${aw5},${h-half} ${aw5},${h-sh5} ${strL + w * 0.05},${h-sh5}" ${fa}/>`;
            inner += `<line x1="${strL * 0.4}" y1="${sh5}" x2="${strL * 0.4}" y2="${h-sh5}" stroke="${stroke}" stroke-width="${sw}"/>`;
            inner += `<line x1="${strL * 0.8}" y1="${sh5}" x2="${strL * 0.8}" y2="${h-sh5}" stroke="${stroke}" stroke-width="${sw}"/>`;
            break;
        }
        case 'bent_arrow': {
            const bh = h * 0.45, baw = w * 0.4;
            inner = `<path d="M${half} ${h-half}V${bh}h${w*0.4}V${bh*0.6}l-${w*0.12} 0 ${w*0.22}-${bh*0.6+half-half} ${w*0.22} ${bh*0.6}-${w*0.12} 0V${bh+h*0.08}H${half+w*0.15}V${h-half}z" ${fa}/>`;
            break;
        }
        case 'u_turn_arrow': {
            const ur = w * 0.2;
            inner = `<path d="M${w*0.25} ${h-half}V${h*0.35}a${ur} ${ur} 0 01${w*0.5} 0v${h*0.2}h${w*0.12}l-${w*0.18} ${h*0.2}-${w*0.18}-${h*0.2}h${w*0.12}v-${h*0.2}a${ur*0.5} ${ur*0.5} 0 00-${w*0.2} 0v${h*0.45}z" ${fa}/>`;
            break;
        }
        case 'circular_arrow': {
            const cr = Math.min(cx, cy) * 0.7;
            inner = `<path d="M${cx + cr} ${cy}a${cr} ${cr} 0 11-${cr * 0.3}-${cr * 0.95}" fill="none" stroke="${stroke}" stroke-width="${sw}" ${dash}/>`;
            inner += `<polygon points="${cx + cr * 0.15},${cy - cr * 0.6} ${cx + cr * 0.15},${cy - cr * 1.2} ${cx + cr * 0.6},${cy - cr * 0.9}" fill="${fill}" fill-opacity="${opacity}" stroke="${stroke}" stroke-width="${sw * 0.5}"/>`;
            break;
        }
        // 수학
        case 'math_plus': {
            const mw = w * 0.25, mh = h * 0.25;
            inner = `<polygon points="${cx-mw},${half} ${cx+mw},${half} ${cx+mw},${cy-mh} ${w-half},${cy-mh} ${w-half},${cy+mh} ${cx+mw},${cy+mh} ${cx+mw},${h-half} ${cx-mw},${h-half} ${cx-mw},${cy+mh} ${half},${cy+mh} ${half},${cy-mh} ${cx-mw},${cy-mh}" ${fa}/>`;
            break;
        }
        case 'math_minus':
            inner = `<rect x="${half}" y="${cy - h * 0.15}" width="${w - sw}" height="${h * 0.3}" ${fa}/>`;
            break;
        case 'math_multiply': {
            const mt = Math.min(w, h) * 0.12;
            inner = `<line x1="${half + mt}" y1="${half + mt}" x2="${w - half - mt}" y2="${h - half - mt}" stroke="${stroke}" stroke-width="${sw * 2}" ${dash}/>`;
            inner += `<line x1="${w - half - mt}" y1="${half + mt}" x2="${half + mt}" y2="${h - half - mt}" stroke="${stroke}" stroke-width="${sw * 2}" ${dash}/>`;
            break;
        }
        case 'math_divide': {
            const dr = Math.min(w, h) * 0.08;
            inner = `<rect x="${half}" y="${cy - h * 0.08}" width="${w - sw}" height="${h * 0.16}" ${fa}/>`;
            inner += `<circle cx="${cx}" cy="${h * 0.22}" r="${dr}" fill="${fill}" fill-opacity="${opacity}" stroke="${stroke}" stroke-width="${sw}"/>`;
            inner += `<circle cx="${cx}" cy="${h * 0.78}" r="${dr}" fill="${fill}" fill-opacity="${opacity}" stroke="${stroke}" stroke-width="${sw}"/>`;
            break;
        }
        case 'math_equal':
            inner = `<rect x="${half}" y="${cy - h * 0.25}" width="${w - sw}" height="${h * 0.15}" ${fa}/>`;
            inner += `<rect x="${half}" y="${cy + h * 0.1}" width="${w - sw}" height="${h * 0.15}" ${fa}/>`;
            break;
        case 'math_not_equal':
            inner = `<rect x="${half}" y="${cy - h * 0.25}" width="${w - sw}" height="${h * 0.12}" ${fa}/>`;
            inner += `<rect x="${half}" y="${cy + h * 0.13}" width="${w - sw}" height="${h * 0.12}" ${fa}/>`;
            inner += `<line x1="${w * 0.65}" y1="${half}" x2="${w * 0.35}" y2="${h - half}" stroke="${stroke}" stroke-width="${sw * 1.5}"/>`;
            break;
        // 별 및 현수막
        case 'star_4_point':
            inner = `<polygon points="${_starPts(cx, cy, Math.min(cx, cy) - half, (Math.min(cx, cy) - half) * 0.38, 4)}" ${fa}/>`;
            break;
        case 'star_5_point':
            inner = `<polygon points="${_starPts(cx, cy, Math.min(cx, cy) - half, (Math.min(cx, cy) - half) * 0.38, 5)}" ${fa}/>`;
            break;
        case 'star_6_point':
            inner = `<polygon points="${_starPts(cx, cy, Math.min(cx, cy) - half, (Math.min(cx, cy) - half) * 0.5, 6)}" ${fa}/>`;
            break;
        case 'star_8_point':
            inner = `<polygon points="${_starPts(cx, cy, Math.min(cx, cy) - half, (Math.min(cx, cy) - half) * 0.5, 8)}" ${fa}/>`;
            break;
        case 'star_10_point':
            inner = `<polygon points="${_starPts(cx, cy, Math.min(cx, cy) - half, (Math.min(cx, cy) - half) * 0.6, 10)}" ${fa}/>`;
            break;
        case 'star_12_point':
            inner = `<polygon points="${_starPts(cx, cy, Math.min(cx, cy) - half, (Math.min(cx, cy) - half) * 0.6, 12)}" ${fa}/>`;
            break;
        case 'star_16_point':
            inner = `<polygon points="${_starPts(cx, cy, Math.min(cx, cy) - half, (Math.min(cx, cy) - half) * 0.7, 16)}" ${fa}/>`;
            break;
        case 'star_24_point':
            inner = `<polygon points="${_starPts(cx, cy, Math.min(cx, cy) - half, (Math.min(cx, cy) - half) * 0.75, 24)}" ${fa}/>`;
            break;
        case 'star_32_point':
            inner = `<polygon points="${_starPts(cx, cy, Math.min(cx, cy) - half, (Math.min(cx, cy) - half) * 0.8, 32)}" ${fa}/>`;
            break;
        case 'explosion_1': {
            var pts1 = [];
            for (var i1 = 0; i1 < 8; i1++) {
                var a1 = Math.PI * 2 * i1 / 8 - Math.PI / 2;
                var or1 = Math.min(cx, cy) - half;
                var ir1 = or1 * (0.45 + Math.random() * 0.15);
                pts1.push((cx + or1 * Math.cos(a1)).toFixed(1) + ',' + (cy + or1 * Math.sin(a1)).toFixed(1));
                var a1b = a1 + Math.PI / 8;
                pts1.push((cx + ir1 * Math.cos(a1b)).toFixed(1) + ',' + (cy + ir1 * Math.sin(a1b)).toFixed(1));
            }
            inner = `<polygon points="${pts1.join(' ')}" ${fa}/>`;
            break;
        }
        case 'explosion_2': {
            var pts2 = [];
            for (var i2 = 0; i2 < 12; i2++) {
                var a2 = Math.PI * 2 * i2 / 12 - Math.PI / 2;
                var or2 = Math.min(cx, cy) - half;
                var ir2 = or2 * (0.5 + (i2 % 3) * 0.05);
                pts2.push((cx + or2 * Math.cos(a2)).toFixed(1) + ',' + (cy + or2 * Math.sin(a2)).toFixed(1));
                var a2b = a2 + Math.PI / 12;
                pts2.push((cx + ir2 * Math.cos(a2b)).toFixed(1) + ',' + (cy + ir2 * Math.sin(a2b)).toFixed(1));
            }
            inner = `<polygon points="${pts2.join(' ')}" ${fa}/>`;
            break;
        }
        case 'wave':
            inner = `<path d="M${half} ${h*0.3}c${w*0.15}-${h*0.25} ${w*0.35}-${h*0.25} ${cx} 0s${w*0.35} ${h*0.25} ${cx} 0v${h*0.4}c-${w*0.15} ${h*0.25}-${w*0.35} ${h*0.25}-${cx} 0s-${w*0.35}-${h*0.25}-${cx} 0z" ${fa}/>`;
            break;
        case 'double_wave':
            inner = `<path d="M${half} ${h*0.2}c${w*0.15}-${h*0.15} ${w*0.35}-${h*0.15} ${cx} 0s${w*0.35} ${h*0.15} ${cx} 0" fill="none" stroke="${stroke}" stroke-width="${sw}" ${dash}/>`;
            inner += `<path d="M${half} ${cy}c${w*0.15}-${h*0.15} ${w*0.35}-${h*0.15} ${cx} 0s${w*0.35} ${h*0.15} ${cx} 0" fill="none" stroke="${stroke}" stroke-width="${sw}" ${dash}/>`;
            inner += `<path d="M${half} ${h*0.8}c${w*0.15}-${h*0.15} ${w*0.35}-${h*0.15} ${cx} 0s${w*0.35} ${h*0.15} ${cx} 0" fill="none" stroke="${stroke}" stroke-width="${sw}" ${dash}/>`;
            break;
        case 'ribbon': {
            const ry = h * 0.15;
            inner = `<path d="M${half} ${ry + half}h${w - sw}v${h - sw - 2 * ry}H${half}z" ${fa}/>`;
            inner += `<path d="M${half} ${ry + half}L${w*0.1} ${half}l${w*0.05} ${ry}" fill="none" stroke="${stroke}" stroke-width="${sw}"/>`;
            inner += `<path d="M${w-half} ${ry + half}L${w*0.9} ${half}l-${w*0.05} ${ry}" fill="none" stroke="${stroke}" stroke-width="${sw}"/>`;
            inner += `<path d="M${half} ${h - ry - half}L${w*0.1} ${h-half}l${w*0.05}-${ry}" fill="none" stroke="${stroke}" stroke-width="${sw}"/>`;
            inner += `<path d="M${w-half} ${h - ry - half}L${w*0.9} ${h-half}l-${w*0.05}-${ry}" fill="none" stroke="${stroke}" stroke-width="${sw}"/>`;
            break;
        }
        // 설명선
        case 'wedge_rect_callout': {
            const ct = h * 0.7;
            inner = `<path d="M${half} ${half}h${w - sw}v${ct}H${cx + w * 0.08}l-${w * 0.08} ${h - half - ct - half}l-${w * 0.08} -${h - half - ct - half}H${half}z" ${fa}/>`;
            break;
        }
        case 'wedge_round_rect_callout': {
            const rrc = Math.min(w, h) * 0.08;
            const cth = h * 0.7;
            inner = `<path d="M${half + rrc} ${half}h${w - sw - 2*rrc}a${rrc} ${rrc} 0 01${rrc} ${rrc}v${cth - rrc}H${cx + w*0.08}l-${w*0.08} ${h - half - cth - half}l-${w*0.08}-${h - half - cth - half}H${half}V${half + rrc}a${rrc} ${rrc} 0 01${rrc}-${rrc}z" ${fa}/>`;
            break;
        }
        case 'wedge_ellipse_callout': {
            const erx = (w - sw) / 2, ery = (h * 0.65 - sw) / 2;
            inner = `<ellipse cx="${cx}" cy="${ery + half}" rx="${erx}" ry="${ery}" ${fa}/>`;
            inner += `<path d="M${cx - w*0.05} ${ery*2}l-${w*0.03} ${h*0.3} ${w*0.12}-${h*0.22}z" fill="${fill}" fill-opacity="${opacity}" stroke="${stroke}" stroke-width="${sw * 0.5}"/>`;
            break;
        }
        case 'cloud_callout':
            inner = `<path d="M${w*0.25} ${h*0.6}a${w*0.12} ${h*0.12} 0 01${w*0.02}-${h*0.25} ${w*0.18} ${h*0.18} 0 01${w*0.35}-${h*0.1} ${w*0.13} ${h*0.1} 0 01${w*0.25} ${h*0.04} ${w*0.1} ${h*0.1} 0 01${w*0.02} ${h*0.18}z" ${fa}/>`;
            inner += `<circle cx="${w*0.3}" cy="${h*0.72}" r="${Math.min(w,h)*0.04}" fill="${fill}" fill-opacity="${opacity}" stroke="${stroke}" stroke-width="${sw}"/>`;
            inner += `<circle cx="${w*0.22}" cy="${h*0.82}" r="${Math.min(w,h)*0.025}" fill="${fill}" fill-opacity="${opacity}" stroke="${stroke}" stroke-width="${sw * 0.7}"/>`;
            break;
        case 'border_callout_1':
            inner = `<rect x="${half}" y="${half}" width="${w - sw}" height="${h * 0.65}" ${fa}/>`;
            inner += `<line x1="${w * 0.3}" y1="${h * 0.65 + half}" x2="${w * 0.2}" y2="${h - half}" stroke="${stroke}" stroke-width="${sw}"/>`;
            break;
        case 'border_callout_2':
            inner = `<rect x="${half}" y="${half}" width="${w - sw}" height="${h * 0.65}" ${fa}/>`;
            inner += `<polyline points="${w*0.3},${h*0.65+half} ${w*0.35},${h*0.8} ${w*0.2},${h-half}" fill="none" stroke="${stroke}" stroke-width="${sw}"/>`;
            break;
        case 'border_callout_3':
            inner = `<rect x="${half}" y="${half}" width="${w - sw}" height="${h * 0.65}" ${fa}/>`;
            inner += `<polyline points="${w*0.3},${h*0.65+half} ${w*0.4},${h*0.75} ${w*0.25},${h*0.85} ${w*0.2},${h-half}" fill="none" stroke="${stroke}" stroke-width="${sw}"/>`;
            break;
        case 'line': {
            const midL = h / 2;
            inner = `<line x1="${half}" y1="${midL}" x2="${w - half}" y2="${midL}" stroke="${stroke}" stroke-width="${sw}" ${dash}/>`;
            break;
        }
        case 'arrow': {
            const midA = h / 2;
            const markerId = 'ah_' + obj.obj_id;
            inner = `<defs><marker id="${markerId}" markerWidth="8" markerHeight="6" refX="7" refY="3" orient="auto"><path d="M0,0 L8,3 L0,6" fill="${stroke}"/></marker></defs>`;
            inner += `<line x1="${half}" y1="${midA}" x2="${w - half}" y2="${midA}" stroke="${stroke}" stroke-width="${sw}" ${dash} marker-end="url(#${markerId})"/>`;
            break;
        }
        default: // rectangle
            inner = `<rect x="${half}" y="${half}" width="${w - sw}" height="${h - sw}" ${fa}/>`;
            break;
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
            const newText = textEl[0].innerText;
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
    // 자동 저장 중지
    _stopEditAutoSave();

    // 협업 프로젝트: Lock 해제 (편집 시작 시 기록한 슬라이드 ID 사용)
    if (state.isCollabProject && state.currentProject && state.editingSlideId) {
        apiDelete('/api/projects/' + state.currentProject._id + '/slides/' + state.editingSlideId + '/lock').catch(() => {});
    }

    state.editMode = false;
    state.editSelectedObj = null;
    state.editingSlideId = null;
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
var presentationIndex = 0;
var _presActivePanel = 'A'; // 'A' 또는 'B' - 현재 활성 패널
var _presDirection = 'forward'; // 전환 방향: 'forward' 또는 'backward'

async function startPresentation() {
    if (state.generatedSlides.length === 0) {
        showToast(t('noSlides'), 'error');
        return;
    }

    // 모든 배경 이미지 + 오브젝트 이미지를 사전 로딩
    const imageUrls = new Set();
    state.generatedSlides.forEach(slide => {
        if (slide.background_image) imageUrls.add(slide.background_image);
        (slide.objects || []).forEach(obj => {
            if (obj.obj_type === 'image' && obj.image_url) imageUrls.add(obj.image_url);
        });
    });

    if (imageUrls.size > 0) {
        await _preloadImages([...imageUrls]);
    }

    presentationIndex = 0;
    _presActivePanel = 'A';
    $('#presentationSlideA').addClass('active');
    $('#presentationSlideB').removeClass('active');
    $('#presentationMode').show();

    // 공유 URL 접속 시 닫기 버튼 숨김
    if (state.isSharedView) {
        $('#presentationMode .presentation-close').hide();
    } else {
        $('#presentationMode .presentation-close').show();
    }

    _presDirection = 'forward';
    renderPresentationSlide(0, 'A');
    _buildPresCarousel();
    _updatePresNav();
    _updatePresProgress();

    $(document).on('keydown.presentation', function (e) {
        if (e.key === 'ArrowRight' || e.key === ' ') {
            e.preventDefault();
            presNavNext();
        } else if (e.key === 'ArrowLeft') {
            e.preventDefault();
            presNavPrev();
        } else if (e.key === 'Escape' && !state.isSharedView) {
            exitPresentation();
        } else if (e.key === 'f' || e.key === 'F') {
            togglePresFullscreen();
        }
    });
}


function _preloadImages(urls) {
    return Promise.all(urls.map(url =>
        new Promise(resolve => {
            const img = new Image();
            img.onload = resolve;
            img.onerror = resolve; // 실패해도 계속 진행
            img.src = url;
            // 10초 타임아웃
            setTimeout(resolve, 10000);
        })
    ));
}


function _waitForImage(url) {
    return new Promise(resolve => {
        const img = new Image();
        img.onload = resolve;
        img.onerror = resolve;
        img.src = url;
        // 이미 캐시되어 있으면 즉시 resolve
        if (img.complete) { resolve(); return; }
        // 5초 타임아웃
        setTimeout(resolve, 5000);
    });
}


function _preloadSlideImages(slides) {
    if (!slides || slides.length === 0) return;
    const urls = new Set();
    slides.forEach(slide => {
        if (slide.background_image) urls.add(slide.background_image);
        (slide.objects || []).forEach(obj => {
            if (obj.obj_type === 'image' && obj.image_url) urls.add(obj.image_url);
        });
    });
    if (urls.size > 0) _preloadImages([...urls]);
}

var _presTransitioning = false;

function transitionPresentationSlide(index) {
    if (_presTransitioning) return;
    _presTransitioning = true;

    var nextPanel = _presActivePanel === 'A' ? 'B' : 'A';
    var $current = $('#presentationSlide' + _presActivePanel);
    var $next = $('#presentationSlide' + nextPanel);
    var dir = _presDirection;
    var exitClass = dir === 'forward' ? 'pres-exit-forward' : 'pres-exit-backward';
    var enterClass = dir === 'forward' ? 'pres-enter-forward' : 'pres-enter-backward';

    // 1) 새 패널에 슬라이드 렌더링 (아직 보이지 않음)
    renderPresentationSlide(index, nextPanel);

    // 2) 새 패널 시작 위치 설정 (transition 없이)
    $next.css('transition', 'none').addClass(enterClass).removeClass('active');
    // 강제 reflow로 시작 위치 적용
    $next[0].offsetHeight;

    // 3) transition 복원 후 동시 전환
    $next.css('transition', '');
    requestAnimationFrame(function() {
        $next.removeClass(enterClass).addClass('active');
        $current.addClass(exitClass).removeClass('active');
    });

    _presActivePanel = nextPanel;
    _updatePresProgress();

    setTimeout(function() {
        // 전환 완료: exit 클래스 제거
        $current.removeClass(exitClass);
        _presTransitioning = false;
    }, 1400);
}

function renderPresentationSlide(index, panel) {
    const slide = state.generatedSlides[index];
    if (!slide) return;

    const container = $(`#presentationSlide${panel}`);
    container.find('.preview-obj').remove();

    if (slide.background_image) {
        $(`#presentationBg${panel}`).css({ 'background-image': `url(${slide.background_image})`, 'background-color': 'transparent' });
    } else {
        $(`#presentationBg${panel}`).css({ 'background-image': 'none', 'background-color': 'white' });
    }

    const containerW = container.width();
    const containerH = container.height();
    const _sz = getCurrentSlideSize();
    const scaleX = containerW / _sz.w;
    const scaleY = containerH / _sz.h;

    let presDescIdx = 0;
    let presSubIdx = 0;
    const presItems = slide.items || [];

    var animIdx = 0; // 애니메이션 stagger 인덱스

    (slide.objects || []).forEach(function(obj) {
        var div = $('<div>').addClass('preview-obj').css({
            position: 'absolute',
            left: (obj.x * scaleX) + 'px',
            top: (obj.y * scaleY) + 'px',
            width: (obj.width * scaleX) + 'px',
            height: (obj.height * scaleY) + 'px',
            zIndex: obj.z_index !== undefined ? obj.z_index : 10,
        });

        var role = (obj.role || obj._auto_role || '');

        if (obj.obj_type === 'image' && obj.image_url) {
            var imgFit = obj.image_fit || 'contain';
            div.append('<img src="' + obj.image_url + '" style="width:100%;height:100%;object-fit:' + imgFit + ';">');
            div.addClass('pres-obj-animate pres-anim-image pres-delay-' + Math.min(animIdx, 8));
            animIdx++;
        } else if (obj.obj_type === 'text') {
            var style = obj.text_style || {};
            var text = obj.generated_text || obj.text_content || '';
            var scaledFontSize = (style.font_size || 16) * Math.min(scaleX, scaleY);
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
                var item = presItems[presDescIdx];
                div.text(item.detail || '');
                presDescIdx++;
                // 텍스트 넘침 시 높이 자동 확장
                div.css({ height: 'auto', minHeight: (obj.height * scaleY) + 'px', overflow: 'visible' });
            } else {
                div.text(text);
            }

            // 역할별 애니메이션 클래스 부여
            var animClass = role === 'title' ? 'pres-anim-title' : '';
            div.addClass('pres-obj-animate ' + animClass + ' pres-delay-' + Math.min(animIdx, 8));
            animIdx++;
        }

        container.append(div);
    });
}

function _buildPresCarousel() {
    const $carousel = $('#presNavCarousel');
    $carousel.empty();
    state.generatedSlides.forEach((slide, i) => {
        const bgStyle = slide.background_image
            ? `background-image:url(${slide.background_image})`
            : 'background:#e5e7eb';
        $carousel.append(
            `<div class="pres-thumb${i === presentationIndex ? ' active' : ''}" data-idx="${i}" style="${bgStyle}" onclick="presNavGoTo(${i})"><span class="pres-thumb-num">${i + 1}</span></div>`
        );
    });
}

function _updatePresNav() {
    const total = state.generatedSlides.length;
    const current = presentationIndex + 1;
    $('#presNavCurrent').text(current);
    $('#presNavTotal').text(total);
    $('#presNavPrev').prop('disabled', presentationIndex <= 0);
    $('#presNavNext').prop('disabled', presentationIndex >= total - 1);

    // 썸네일 active 표시
    $('#presNavCarousel .pres-thumb').removeClass('active');
    $('#presNavCarousel .pres-thumb[data-idx="' + presentationIndex + '"]').addClass('active');

    // 캐러셀 스크롤 — active 썸네일이 보이도록
    _scrollCarouselToActive();
}

function _scrollCarouselToActive() {
    const $wrap = $('.pres-nav-carousel-wrap');
    const $carousel = $('#presNavCarousel');
    const $active = $carousel.find('.pres-thumb.active');
    if (!$active.length) return;

    const wrapW = $wrap.width();
    const thumbL = $active[0].offsetLeft;
    const thumbW = $active.outerWidth();
    // 활성 썸네일을 중앙에 위치
    let offset = thumbL - (wrapW / 2) + (thumbW / 2);
    const maxOffset = $carousel[0].scrollWidth - wrapW;
    offset = Math.max(0, Math.min(offset, maxOffset));
    $carousel.css('transform', 'translateX(' + (-offset) + 'px)');
}

function presNavGoTo(idx) {
    if (idx === presentationIndex || idx < 0 || idx >= state.generatedSlides.length) return;
    _presDirection = idx > presentationIndex ? 'forward' : 'backward';
    presentationIndex = idx;
    transitionPresentationSlide(presentationIndex);
    _updatePresNav();
}

function _updatePresProgress() {
    var total = state.generatedSlides.length;
    var pct = total <= 1 ? 100 : ((presentationIndex / (total - 1)) * 100);
    $('#presProgressBar').css('width', pct + '%');
}

function presNavPrev() {
    if (presentationIndex > 0) {
        presNavGoTo(presentationIndex - 1);
    }
}

function presNavNext() {
    if (presentationIndex < state.generatedSlides.length - 1) {
        presNavGoTo(presentationIndex + 1);
    }
}

function togglePresFullscreen() {
    const presEl = document.getElementById('presentationMode');
    if (document.fullscreenElement || document.webkitFullscreenElement || document.msFullscreenElement) {
        if (document.exitFullscreen) document.exitFullscreen();
        else if (document.webkitExitFullscreen) document.webkitExitFullscreen();
        else if (document.msExitFullscreen) document.msExitFullscreen();
    } else {
        if (presEl.requestFullscreen) presEl.requestFullscreen().catch(() => {});
        else if (presEl.webkitRequestFullscreen) presEl.webkitRequestFullscreen();
        else if (presEl.msRequestFullscreen) presEl.msRequestFullscreen();
    }
}

function exitPresentation() {
    $('#presentationMode').hide();
    $(document).off('keydown.presentation');
    // 전체화면 해제
    if (document.fullscreenElement || document.webkitFullscreenElement || document.msFullscreenElement) {
        if (document.exitFullscreen) document.exitFullscreen();
        else if (document.webkitExitFullscreen) document.webkitExitFullscreen();
        else if (document.msExitFullscreen) document.msExitFullscreen();
    }
}

// 전체화면 변경 시 아이콘 업데이트 + 비공유 뷰 종료 처리
function _onFullscreenChange() {
    const isFS = !!(document.fullscreenElement || document.webkitFullscreenElement || document.msFullscreenElement);
    const $icon = $('#presFullscreenIcon');
    if (isFS) {
        // 축소 아이콘
        $icon.html('<path d="M5 2v3H2M11 2v3h3M5 14v-3H2M11 14v-3h3"/>');
    } else {
        // 확대 아이콘
        $icon.html('<path d="M2 5V2h3M14 5V2h-3M2 11v3h3M14 11v3h-3"/>');
        // 비공유 뷰에서 전체화면 해제 시 프레젠테이션 종료
        if ($('#presentationMode').is(':visible') && !state.isSharedView) {
            exitPresentation();
        }
    }
}
document.addEventListener('fullscreenchange', _onFullscreenChange);
document.addEventListener('webkitfullscreenchange', _onFullscreenChange);

// ============ 다운로드 & 공유 ============

async function _downloadFile(url, defaultFilename) {
    // 다운로드 오버레이 표시
    let $overlay = $('#downloadOverlay');
    if (!$overlay.length) {
        $('body').append(`
            <div id="downloadOverlay" style="position:fixed;inset:0;z-index:99999;display:flex;align-items:center;justify-content:center;background:rgba(0,0,0,.35);backdrop-filter:blur(4px);opacity:0;transition:opacity .2s">
                <div style="background:#fff;border-radius:16px;padding:32px 48px;display:flex;flex-direction:column;align-items:center;gap:16px;box-shadow:0 20px 60px rgba(0,0,0,.2)">
                    <div class="dl-spinner" style="width:40px;height:40px;border:3px solid #e5e7eb;border-top-color:#6366f1;border-radius:50%;animation:dlSpin .7s linear infinite"></div>
                    <div style="font-size:15px;font-weight:600;color:#1f2937" id="downloadOverlayMsg">문서를 준비하고 있습니다...</div>
                    <div style="font-size:12px;color:#9ca3af">잠시만 기다려 주세요</div>
                </div>
            </div>
            <style>@keyframes dlSpin{to{transform:rotate(360deg)}}</style>
        `);
        $overlay = $('#downloadOverlay');
    }
    $('#downloadOverlayMsg').text('문서를 준비하고 있습니다...');
    $overlay.show();
    requestAnimationFrame(() => $overlay.css('opacity', '1'));

    try {
        const res = await fetch(url);
        if (!res.ok) {
            const err = await res.json().catch(() => ({}));
            throw new Error(err.detail || `다운로드 실패 (${res.status})`);
        }

        // Content-Disposition에서 파일명 추출
        const cd = res.headers.get('Content-Disposition') || '';
        let filename = defaultFilename;
        const match = cd.match(/filename\*?=(?:UTF-8''|"?)([^";]+)/i);
        if (match) filename = decodeURIComponent(match[1].replace(/"/g, ''));

        $('#downloadOverlayMsg').text('다운로드 중...');
        const blob = await res.blob();
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 100);
        showToast('다운로드가 완료되었습니다.', 'success');
    } catch (e) {
        console.error('[Download]', e);
        showToast(e.message || '다운로드 실패', 'error');
    } finally {
        $overlay.css('opacity', '0');
        setTimeout(() => $overlay.hide(), 200);
    }
}

function downloadPPTX() {
    if (!state.currentProject || state.generatedSlides.length === 0) {
        showToast(t('noSlides'), 'error');
        return;
    }
    const url = apiUrl('/api/generate/' + state.currentProject._id + '/download/pptx');
    _downloadFile(url, (state.currentProject.name || 'presentation') + '.pptx');
}

function downloadPDF() {
    if (!state.currentProject || state.generatedSlides.length === 0) {
        showToast(t('noSlides'), 'error');
        return;
    }
    const url = apiUrl('/api/generate/' + state.currentProject._id + '/download/pdf');
    _downloadFile(url, (state.currentProject.name || 'presentation') + '.pdf');
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

// ============ 언어 전환 (번역) ============
function showTranslateModal() {
    if (!state.currentProject || state.generatedSlides.length === 0) {
        showToast(t('translateNoSlides', '생성된 슬라이드가 없습니다'), 'error');
        return;
    }

    const langInfo = {
        ko: { name: '한국어', flag: '\uD83C\uDDF0\uD83C\uDDF7' },
        en: { name: 'English', flag: '\uD83C\uDDFA\uD83C\uDDF8' },
        ja: { name: '日本語', flag: '\uD83C\uDDEF\uD83C\uDDF5' },
        zh: { name: '中文', flag: '\uD83C\uDDE8\uD83C\uDDF3' },
    };

    const list = $('#translateLangList');
    list.empty();

    // Get available languages from the lang selector
    const availableLangs = [];
    $('#langSelect option').each(function() {
        const code = $(this).val();
        if (code && langInfo[code]) {
            availableLangs.push(code);
        }
    });

    // If no langs from selector, use defaults
    if (availableLangs.length === 0) {
        ['ko', 'en', 'ja', 'zh'].forEach(c => availableLangs.push(c));
    }

    availableLangs.forEach(code => {
        const info = langInfo[code] || { name: code, flag: '\uD83C\uDF10' };
        list.append(`
            <div class="translate-lang-item" onclick="startTranslation('${code}')">
                <span class="lang-flag">${info.flag}</span>
                <div class="lang-info">
                    <div class="lang-name">${info.name}</div>
                    <div class="lang-code">${code.toUpperCase()}</div>
                </div>
                <span class="lang-arrow">\u2192</span>
            </div>
        `);
    });

    // Update modal text with i18n
    $('#translateModal .i18n-translateTitle').text(t('translateTitle', '언어 전환'));
    $('#translateModal .i18n-translateDesc').html(t('translateDesc', '현재 슬라이드를 선택한 언어로 번역합니다.<br>원본 프로젝트는 그대로 보존됩니다.'));

    $('#translateModal').show();
}

async function startTranslation(targetLang) {
    closeModal('translateModal');

    if (_isGenerating) {
        showToast(t('translateInProgress', '다른 작업이 진행 중입니다'), 'error');
        return;
    }

    const langNames = { ko: '한국어', en: 'English', ja: '日本語', zh: '中文' };
    const langName = langNames[targetLang] || targetLang;

    _isGenerating = true;
    _showStopButton();
    _setSlideToolsDisabled(true);

    // --- Phase 1 Setup: 기존 아웃라인 유지, 번역 중 상태 표시 ---
    switchPanelTab('outline');

    // 기존 아웃라인 항목에 translating 클래스 추가 (스피너 표시)
    $('#slideTextList .slide-text-item').each(function() {
        $(this).addClass('translating');
    });

    // 상단 상태 배너 추가
    const translatingMsg = t('translatingTo', '{lang}로 번역 중입니다...').replace('{lang}', langName);
    $('#slideTextList').prepend(`
        <div id="translateStatusBanner" style="padding:10px 14px;display:flex;align-items:center;gap:8px;border-bottom:1px solid var(--border);">
            <div class="streaming-spinner-sm"></div>
            <span id="translateStatusText" style="font-size:13px;font-weight:600;color:var(--accent);">${translatingMsg}</span>
        </div>
    `);

    // 캔버스: 진행률 바 표시
    $('#canvasLoadingOverlay').remove();
    $('#translateProgressBar').remove();
    $('#previewCanvas').append(`<div class="translate-progress-bar" id="translateProgressBar" style="width:0%"></div>`);

    _abortController = new AbortController();
    let newProjectId = null;
    let newProjectName = null;
    let totalSlides = state.generatedSlides.length;
    const translatedSlides = new Array(totalSlides).fill(null);
    let completedCount = 0;

    try {
        const response = await fetch(`/${state.jwtToken}/api/generate/translate/stream`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                project_id: state.currentProject._id,
                target_lang: targetLang,
            }),
            signal: _abortController.signal,
        });

        if (!response.ok) {
            let errMsg = t('translateFailed', '번역 실패');
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

                        if (data.event === 'start') {
                            newProjectId = data.new_project_id;
                            newProjectName = data.new_project_name;
                            totalSlides = data.total_slides || totalSlides;

                        } else if (data.event === 'slide_progress') {
                            // 텍스트 없는 슬라이드 - 원본 그대로 복사
                            const idx = data.index;
                            if (idx != null && idx < state.generatedSlides.length) {
                                translatedSlides[idx] = state.generatedSlides[idx];
                                completedCount++;
                                _updateTranslateOutlineItem(idx, state.generatedSlides[idx]);
                                _updateTranslateProgress(completedCount, totalSlides, langName);
                            }

                        } else if (data.event === 'slide_complete') {
                            const idx = data.index;
                            if (idx != null) {
                                translatedSlides[idx] = data.slide;
                                completedCount++;
                                // Phase 1: 아웃라인 항목을 번역된 내용으로 교체
                                _updateTranslateOutlineItem(idx, data.slide);
                                _updateTranslateProgress(completedCount, totalSlides, langName);
                            }

                        } else if (data.event === 'complete') {
                            newProjectId = data.new_project_id;
                            newProjectName = data.new_project_name;

                        } else if (data.event === 'error') {
                            showToast(data.message || t('translateError', '번역 중 오류가 발생했습니다.'), 'error');
                        }
                    } catch (e) {
                        console.warn('Translate SSE parse error:', e);
                    }
                }
            }
        }

        // --- Phase 2: 슬라이드 프리뷰 순차 애니메이션 ---
        if (newProjectId && translatedSlides.filter(Boolean).length > 0) {
            $('#translateProgressBar').css('width', '100%');
            const animatingMsg = t('translateAnimating', '번역 완료! 슬라이드 미리보기 업데이트 중...');
            $('#translateStatusText').text(animatingMsg);

            // 번역된 데이터를 임시로 적용하여 프리뷰 표시
            const originalSlides = [...state.generatedSlides];
            const validTranslated = translatedSlides.map((s, i) => s || originalSlides[i]);
            state.generatedSlides = validTranslated;

            // 각 슬라이드를 순차적으로 캔버스에 표시
            await _animateTranslatedSlides(validTranslated);

            // 원본 복원 후 새 프로젝트로 전환
            state.generatedSlides = originalSlides;

            const doneMsg = t('translateDone', '{lang}로 번역 완료! 새 프로젝트: {name}')
                .replace('{lang}', langName)
                .replace('{name}', newProjectName);
            showToast(doneMsg, 'success');

            await loadProjects();
            await openProject(newProjectId);
        }

    } catch (e) {
        if (e.name === 'AbortError') {
            console.log('Translation aborted by user');
        } else {
            showToast(e.message || t('translateError', '번역 중 오류가 발생했습니다.'), 'error');
            console.error('Translate stream error:', e);
        }
    } finally {
        _isGenerating = false;
        _streamReader = null;
        _abortController = null;
        $('#translateProgressBar').remove();
        $('#translateStatusBanner').remove();
        $('#slideTextList .slide-text-item').removeClass('translating translate-done');
        _showGenerateOrRestartButton();
        _setSlideToolsDisabled(false);
    }
}

// 번역 진행률 업데이트
function _updateTranslateProgress(completedCount, totalSlides, langName) {
    const pct = Math.round((completedCount / totalSlides) * 100);
    $('#translateProgressBar').css('width', pct + '%');
    const progressMsg = t('translatingProgress', '{lang}로 번역 중... ({current}/{total})')
        .replace('{lang}', langName)
        .replace('{current}', completedCount)
        .replace('{total}', totalSlides);
    $('#translateStatusText').text(progressMsg);
}

// Phase 1: 아웃라인 항목을 번역된 내용으로 in-place 교체
function _updateTranslateOutlineItem(index, slide) {
    const item = $(`.slide-text-item[data-slide-idx="${index}"]`);
    if (item.length === 0) return;

    item.removeClass('translating');

    const typeLabels = {
        title_slide: 'Cover', toc: 'Contents',
        section_divider: 'Chapter', body: '', closing: 'Closing',
    };

    const textObjs = (slide.objects || []).filter(o => o.obj_type === 'text');
    let titleText = '';
    const sections = [];

    // 제목 추출
    textObjs.forEach(obj => {
        let text = (obj.generated_text || '').trim();
        if (!text) {
            const fb = (obj.text_content || '').trim();
            if (fb && fb !== '텍스트를 입력하세요' && fb !== 'Enter text') text = fb;
        }
        if (!text) return;
        const role = obj.role || obj._auto_role || '';
        if (role === 'number' || role === 'governance') return;
        const fontSize = (obj.text_style || {}).font_size || 16;
        if (!titleText && (role === 'title' || role === 'subtitle' || fontSize >= 24)) {
            titleText = text;
        }
    });

    // 섹션 추출
    const slideItems = slide.items || [];
    if (slideItems.length > 0) {
        slideItems.forEach(si => {
            sections.push({ header: si.heading || '', body: si.detail || '' });
        });
    } else {
        textObjs.forEach(obj => {
            let text = (obj.generated_text || '').trim();
            if (!text) {
                const fb = (obj.text_content || '').trim();
                if (fb && fb !== '텍스트를 입력하세요' && fb !== 'Enter text') text = fb;
            }
            if (!text) return;
            const role = obj.role || obj._auto_role || '';
            const fontSize = (obj.text_style || {}).font_size || 16;
            if (role === 'governance' || role === 'number') return;
            if (role === 'title' || (!role && fontSize >= 24)) return;
            if (role === 'subtitle' || (fontSize >= 16 && fontSize < 24 && (obj.text_style || {}).bold)) {
                sections.push({ header: text, body: '' });
            } else {
                const blocks = text.split('\n\n');
                if (blocks.length > 1) {
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

    item.html(`
        <div class="slide-text-item-header">
            <div class="slide-text-item-num">${index + 1}</div>
            <div class="slide-text-item-title">${escapeHtml(titleText || t('slideN', '슬라이드 {n}').replace('{n}', index + 1))}</div>
            ${badgeHtml}
        </div>
        ${sectionsHtml}
    `);

    // 완료 하이라이트 애니메이션
    item.addClass('translate-done');
    setTimeout(() => item.removeClass('translate-done'), 700);

    // 해당 항목으로 스크롤
    item[0].scrollIntoView({ block: 'nearest', behavior: 'smooth' });
}

// Phase 2: 번역된 슬라이드를 순차적으로 캔버스에 표시
async function _animateTranslatedSlides(slides) {
    const delay = slides.length > 10 ? 300 : (slides.length > 5 ? 500 : 700);

    for (let i = 0; i < slides.length; i++) {
        if (!_isGenerating) break;

        state.currentSlideIndex = i;
        renderSlideAtIndex(i);
        updateSlideNav();
        renderSlideThumbnails();

        // 썸네일 목록에서도 활성 상태 업데이트
        $('.slide-thumb-item').removeClass('active');
        $(`.slide-thumb-item[data-idx="${i}"]`).addClass('active');

        // 아웃라인에서도 활성 상태 업데이트
        $('.slide-text-item').removeClass('active');
        $(`.slide-text-item[data-slide-idx="${i}"]`).addClass('active');

        // 캔버스 flash 애니메이션
        const canvas = $('#previewCanvas');
        canvas.addClass('translate-slide-flash');

        await sleep(delay);
        canvas.removeClass('translate-slide-flash');

        if (i < slides.length - 1) {
            await sleep(80);
        }
    }
}

// ============ 공유 프레젠테이션 ============
async function loadSharedPresentation(shareToken) {
    state.isSharedView = true;
    try {
        // 공유 페이지용 폰트 로딩 (인증 불필요)
        try {
            const fontRes = await fetch('/api/fonts/public');
            if (fontRes.ok) {
                const fontData = await fontRes.json();
                loadWebFonts(fontData.fonts || []);
            }
        } catch (e) { /* 폰트 로딩 실패 무시 */ }

        // 프로젝트 타입 조회
        const infoRes = await fetch('/api/shared/' + shareToken + '/info');
        if (!infoRes.ok) throw new Error('Not found');
        const info = await infoRes.json();
        const projectType = info.project_type || 'slide';

        if (projectType === 'excel') {
            await loadSharedExcel(shareToken, info.project_name);
        } else if (projectType === 'word') {
            await loadSharedWord(shareToken, info.project_name);
        } else {
            // 슬라이드 (기본)
            const res = await fetch('/api/shared/' + shareToken + '/slides');
            if (!res.ok) throw new Error('Not found');
            const data = await res.json();
            state.generatedSlides = data.slides || [];
            state.currentSlideIndex = 0;

            // 공유 프레젠테이션 템플릿 slide_size 설정
            if (data.template && data.template.slide_size) {
                state._templateSlideSize = data.template.slide_size;
            } else {
                state._templateSlideSize = '16:9';
            }
            updateSlideCanvasAspect();

            document.title = data.project_name + ' - ' + (window.__SOLUTION_NAME__ || 'OfficeMaker');
            showApp();
            $('#sidebar').hide();
            $('#sidebarUserName').text(t('sharedPresentation'));

            if (state.generatedSlides.length > 0) {
                startPresentation();
            }
        }
    } catch (e) {
        alert('Shared content not found.');
    }
}

async function loadSharedExcel(shareToken, projectName) {
    const res = await fetch('/api/shared/' + shareToken + '/excel');
    if (!res.ok) throw new Error('Not found');
    const data = await res.json();

    document.title = (projectName || data.project_name) + ' - ' + (window.__SOLUTION_NAME__ || 'OfficeMaker');
    showApp();
    $('#sidebar').hide();
    $('#sidebarUserName').text(t('sharedPresentation'));

    // 빈 상태 숨기고 프로젝트 워크스페이스 표시
    $('#emptyState').hide();
    $('#projectWorkspace').css('display', 'flex');
    $('.workspace-header').hide();

    // 모든 워크스페이스 숨기기, 엑셀만 표시
    $('#slidePreview').hide();
    $('#slideEmpty').hide();
    $('#wordWorkspace').hide();
    $('#onlyofficeWorkspace').hide();
    $('#excelWorkspace').show();
    $('#wsSlideTools').hide();
    $('#inputBar').hide();
    // 공유 모드에서는 편집 관련 버튼 숨기기
    $('#btnUploadExcel').hide();
    $('#btnDownloadXlsx').hide();
    $('#btnShareExcel').hide();
    $('#btnModifyExcel').hide();
    $('#excelTitle').text(data.excel.meta && data.excel.meta.title ? data.excel.meta.title : (projectName || '스프레드시트'));

    // 엑셀 툴바 숨기기 (공유 읽기 전용)
    $('.excel-toolbar').hide();

    // Univer 로드 및 데이터 표시
    await loadUniverScripts();
    populateUniverFromData(data.excel);
    renderExcelCharts(data.excel);
}

async function loadSharedWord(shareToken, projectName) {
    const res = await fetch('/api/shared/' + shareToken + '/docx');
    if (!res.ok) throw new Error('Not found');
    const data = await res.json();

    document.title = (projectName || data.project_name) + ' - ' + (window.__SOLUTION_NAME__ || 'OfficeMaker');
    showApp();
    $('#sidebar').hide();
    $('#sidebarUserName').text(t('sharedPresentation'));

    // 빈 상태 숨기고 프로젝트 워크스페이스 표시
    $('#emptyState').hide();
    $('#projectWorkspace').css('display', 'flex');
    $('.workspace-header').hide();

    // 모든 워크스페이스 숨기기, 워드만 표시
    $('#slidePreview').hide();
    $('#slideEmpty').hide();
    $('#excelWorkspace').hide();
    $('#onlyofficeWorkspace').hide();
    $('#wordWorkspace').show();
    $('#wsSlideTools').hide();
    $('#inputBar').hide();
    // 공유 모드에서는 편집 관련 버튼 숨기기
    $('#btnDownloadDocx').hide();
    $('#btnShareWord').hide();
    $('#wordTitle').text(data.docx.meta && data.docx.meta.title ? data.docx.meta.title : (projectName || '문서'));

    // CKEditor 로드 및 읽기 전용으로 표시
    await loadCKEditorScript();
    if (_ckEditorInstance) {
        await _ckEditorInstance.destroy();
        _ckEditorInstance = null;
    }
    const container = document.getElementById('ckeditorContainer');
    if (container) container.innerHTML = '<div id="ckeditorEditor"></div>';

    const {
        ClassicEditor, Essentials,
        Bold, Italic, Underline, Strikethrough,
        Font, Alignment, List, Link,
        Table, TableToolbar, TableProperties, TableCellProperties,
        Indent, IndentBlock, BlockQuote, Heading,
        Paragraph, GeneralHtmlSupport
    } = CKEDITOR;

    _ckEditorInstance = await ClassicEditor.create(
        document.getElementById('ckeditorEditor'),
        {
            plugins: [
                Essentials, Bold, Italic, Underline, Strikethrough,
                Font, Alignment, List, Link,
                Table, TableToolbar, TableProperties, TableCellProperties,
                Indent, IndentBlock, BlockQuote, Heading, Paragraph,
                GeneralHtmlSupport
            ],
            toolbar: [],
            htmlSupport: {
                allow: [
                    { name: 'img', attributes: true, styles: true, classes: true },
                    { name: 'div', attributes: true, styles: true, classes: true },
                ]
            },
            language: 'ko',
            licenseKey: 'GPL'
        }
    );

    // 차트 렌더링을 위해 ECharts 로드
    if (typeof echarts === 'undefined') {
        try { await loadEChartsScript(); } catch(e) {}
    }
    const html = _docxSectionsToHtml(data.docx);
    _ckEditorInstance.setData(html);
    _ckEditorInstance.enableReadOnlyMode('shared');

    // CKEditor 툴바 숨기기 (공유 읽기 전용)
    const ckToolbar = document.querySelector('.ck-toolbar');
    if (ckToolbar) ckToolbar.style.display = 'none';

    // 워드 툴바도 숨기기
    $('.word-toolbar').hide();
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
    }).then(function () {
        // 순서 변경 후 타임스탬프 캐시 무효화 → 다음 폴링에서 본인 변경을 재로드하지 않도록
        // 서버가 모든 슬라이드의 updated_at을 갱신하므로 캐시를 초기화
        slideIds.forEach(function (id) {
            delete state.lastSlideTimestamps[id];
        });
    }).catch(function (e) {
        showToast('순서 저장 실패: ' + e.message, 'error');
    });
}

// ============ 협업 기능 ============

async function initCollaboration(projectId) {
    try {
        const res = await apiGet('/api/projects/' + projectId + '/collaborators');
        state.collaborators = res.collaborators || [];
        state.isCollabProject = (
            state.collaborators.length > 0 || state.collabRole !== 'owner'
        );

        if (state.isCollabProject) {
            // 슬라이드 타임스탬프 초기화
            state.lastSlideTimestamps = {};
            state.generatedSlides.forEach(s => {
                if (s._id && s.updated_at) {
                    state.lastSlideTimestamps[s._id] = s.updated_at;
                }
            });
            startCollabPolling(projectId);
            updateCollabUI();
        }
    } catch (e) {
        state.isCollabProject = false;
    }
}

function startCollabPolling(projectId) {
    stopCollabPolling();
    sendHeartbeat(projectId);
    pollCollabStatus(projectId);
    state.pollInterval = setInterval(() => pollCollabStatus(projectId), 5000);
    state.heartbeatInterval = setInterval(() => sendHeartbeat(projectId), 30000);
}

function stopCollabPolling() {
    if (state.pollInterval) {
        clearInterval(state.pollInterval);
        state.pollInterval = null;
    }
    if (state.heartbeatInterval) {
        clearInterval(state.heartbeatInterval);
        state.heartbeatInterval = null;
    }
    state.activeLocks = {};
    state.onlineUsers = [];
}

async function pollCollabStatus(projectId) {
    try {
        const res = await apiGet('/api/projects/' + projectId + '/collab-status');

        // Lock 상태 업데이트
        const newLocks = {};
        (res.locks || []).forEach(lock => {
            newLocks[lock.slide_id] = lock;
        });
        state.activeLocks = newLocks;
        state.onlineUsers = res.online_users || [];

        // 슬라이드 구조 변경 감지 (추가/삭제)
        const serverSlideIds = new Set(Object.keys(res.slide_timestamps || {}));
        const localSlideIds = new Set(state.generatedSlides.map(s => s._id).filter(Boolean));
        const addedIds = [...serverSlideIds].filter(id => !localSlideIds.has(id));
        const removedIds = [...localSlideIds].filter(id => !serverSlideIds.has(id));

        // 기존 슬라이드 내용 변경 감지 (본인이 편집 중인 슬라이드는 제외)
        const changedSlides = [];
        Object.entries(res.slide_timestamps || {}).forEach(([slideId, ts]) => {
            const prev = state.lastSlideTimestamps[slideId];
            if (prev && prev !== ts) {
                // 본인이 편집 중인 슬라이드는 자동저장에 의한 변경이므로 제외
                if (state.editMode && state.editingSlideId === slideId) return;
                changedSlides.push(slideId);
            }
        });

        // 타임스탬프 캐시 전체 갱신
        state.lastSlideTimestamps = { ...(res.slide_timestamps || {}) };

        // 구조 변경 시 전체 슬라이드 재로딩
        if (addedIds.length > 0 || removedIds.length > 0) {
            await reloadAllSlides();
        } else if (changedSlides.length > 0) {
            await reloadChangedSlides(changedSlides);
        }

        updateLockIndicators();
        updateOnlinePresence();
    } catch (e) {
        // 폴링 실패 시 무시 (다음 주기에 재시도)
    }
}

async function sendHeartbeat(projectId) {
    try {
        const body = {};
        // 편집 모드일 때만 editing_slide_id 전송 → 서버가 해당 Lock만 갱신
        if (state.editMode && state.editingSlideId) {
            body.editing_slide_id = state.editingSlideId;
        }
        await apiPost('/api/projects/' + projectId + '/heartbeat', body);
    } catch (e) { /* silent */ }
}

async function reloadChangedSlides(slideIds) {
    try {
        // 변경된 슬라이드만 델타 API로 가져옴 (전체 로딩 방지)
        const res = await apiPost('/api/generate/' + state.currentProject._id + '/slides/delta', {
            slide_ids: slideIds,
        });
        const deltaSlides = res.slides || [];

        // 순서 변경 감지: 하나라도 order가 다르면 전체 재로딩으로 전환
        const orderChanged = deltaSlides.some(updated => {
            const existing = state.generatedSlides.find(s => s._id === updated._id);
            return existing && existing.order !== updated.order;
        });

        if (orderChanged) {
            // 순서 변경은 전체 재로딩이 안전 (편집 중 슬라이드 누락 방지)
            await reloadAllSlides();
            return;
        }

        deltaSlides.forEach(updated => {
            const idx = state.generatedSlides.findIndex(s => s._id === updated._id);
            if (idx !== -1) {
                state.generatedSlides[idx] = updated;
                if (idx === state.currentSlideIndex && !state.editMode) {
                    renderSlideAtIndex(idx);
                } else if (idx === state.currentSlideIndex && state.editMode) {
                    showToast('다른 사용자가 이 슬라이드를 수정했습니다', 'warning');
                }
            }
        });

        renderSlideThumbnails();
        renderSlideThumbList();
    } catch (e) { /* silent */ }
}

async function reloadAllSlides() {
    if (!state.currentProject) return;
    try {
        const res = await apiGet('/api/generate/' + state.currentProject._id + '/slides');
        const newSlides = res.slides || [];

        // 현재 보고 있는 슬라이드 _id 기억 (선택 유지)
        const currentSlideId = (state.generatedSlides[state.currentSlideIndex] || {})._id || null;

        state.generatedSlides = newSlides;

        // 같은 슬라이드를 유지하거나 가장 가까운 인덱스로 이동
        if (currentSlideId) {
            const newIndex = newSlides.findIndex(s => s._id === currentSlideId);
            state.currentSlideIndex = newIndex >= 0 ? newIndex :
                Math.min(state.currentSlideIndex, newSlides.length - 1);
        } else {
            state.currentSlideIndex = Math.min(state.currentSlideIndex, newSlides.length - 1);
        }
        if (state.currentSlideIndex < 0) state.currentSlideIndex = 0;

        // UI 갱신
        renderSlideThumbnails();
        renderSlideThumbList();
        renderSlideTextPanel();
        if (!state.editMode) {
            renderSlideAtIndex(state.currentSlideIndex);
        }
        updateSlideNav();

        // 빈 상태 처리
        if (newSlides.length === 0) {
            $('#slidePreview').hide();
            $('#slideEmpty').show();
        } else {
            $('#slideEmpty').hide();
            $('#slidePreview').show();
        }
    } catch (e) { /* silent */ }
}

function updateLockIndicators() {
    $('.slide-thumb-v').each(function() {
        const slideId = $(this).data('slide-id');
        const lock = state.activeLocks[slideId];
        $(this).find('.lock-badge').remove();

        if (lock && lock.user_key !== (state.userInfo && state.userInfo.ky)) {
            $(this).append(`<div class="lock-badge"><svg width="10" height="10" fill="currentColor" viewBox="0 0 16 16"><path d="M8 1a3 3 0 0 0-3 3v2H4a1 1 0 0 0-1 1v6a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1V7a1 1 0 0 0-1-1h-1V4a3 3 0 0 0-3-3zm-2 3a2 2 0 1 1 4 0v2H6V4z"/></svg><span>${escapeHtml(lock.user_name)}</span></div>`);
        }
    });
}

function updateOnlinePresence() {
    const $bar = $('#collabPresenceBar');
    if (!state.isCollabProject) { $bar.hide(); return; }
    $bar.show();
    const avatars = state.onlineUsers.map(u => {
        const isSelf = state.userInfo && u.user_key === state.userInfo.ky;
        return `<div class="presence-avatar ${isSelf ? 'self' : ''}" title="${escapeHtml(u.user_name)}">${escapeHtml(u.user_name.charAt(0))}</div>`;
    }).join('');
    $bar.html(avatars);
}

function updateCollabUI() {
    const isOwner = state.collabRole === 'owner';
    const isEditor = state.collabRole === 'editor';
    const isViewer = state.collabRole === 'viewer';

    // 협업 관리 버튼 (owner만 + 협업 프로젝트일 때)
    if (state.isCollabProject || isOwner) {
        $('#btnCollabManage').show();
        const count = state.collaborators.length;
        $('#collabCountLabel').text(count > 0 ? `협업 (${count})` : '협업');
    } else {
        $('#btnCollabManage').hide();
    }

    // 히스토리 버튼 (협업 프로젝트일 때)
    if (state.isCollabProject) {
        $('#btnHistory').show();
    } else {
        $('#btnHistory').hide();
    }

    // viewer: 편집/삭제/리셋/생성 관련 버튼 숨김
    if (isViewer) {
        $('#btnEditToggle').hide();
        $('#btnEditSave').hide();
        $('.resource-add-btns').hide();
        $('#inputBar').hide();
        $('#btnResetProject').hide();
        $('#btnDeleteProject').hide();
        $('#slideInstructionBar').hide();
    } else if (isEditor) {
        $('#btnEditToggle').show();
        // editor는 슬라이드 AI 텍스트 생성 가능
        $('#slideInstructionBar').show();
        // editor는 리소스/생성 관련은 숨김
        $('.resource-add-btns').hide();
        $('#inputBar').hide();
        $('#btnResetProject').hide();
        $('#btnDeleteProject').hide();
    } else {
        // owner: 모두 표시
        $('#btnEditToggle').show();
        $('.resource-add-btns').show();
        $('#inputBar').show();
    }
}

// ── 협업자 관리 모달 ──

let _collabSearchTimer = null;

async function showCollabModal() {
    if (!state.currentProject) return;
    try {
        const res = await apiGet('/api/projects/' + state.currentProject._id + '/collaborators');
        state.collaborators = res.collaborators || [];
    } catch (e) {}
    renderCollabList();

    // owner만 추가 UI 표시
    if (state.collabRole === 'owner') {
        $('#collabAddRow').show();
    } else {
        $('#collabAddRow').hide();
    }

    $('#collabUserSearch').val('');
    $('#collabUserKey').val('');
    $('#collabSearchDropdown').hide();
    $('#collabModal').show();
}

function renderCollabList() {
    const $list = $('#collabList');
    $list.empty();

    if (state.collaborators.length === 0) {
        $list.html('<div style="padding:16px;text-align:center;color:var(--text-tertiary);font-size:13px;">아직 협업자가 없습니다</div>');
        return;
    }

    const canManage = state.collabRole === 'owner';
    state.collaborators.forEach(c => {
        $list.append(`
            <div class="collab-item">
                <div class="collab-avatar">${escapeHtml((c.user_name || '?').charAt(0))}</div>
                <div class="collab-info">
                    <div class="collab-name">${escapeHtml(c.user_name || '')}</div>
                    <div class="collab-dept">${escapeHtml(c.user_dept || '')}</div>
                </div>
                ${canManage ? `
                <select class="collab-role-select" onchange="updateCollabRole('${c.user_key}', this.value)">
                    <option value="editor" ${c.role === 'editor' ? 'selected' : ''}>편집자</option>
                    <option value="viewer" ${c.role === 'viewer' ? 'selected' : ''}>뷰어</option>
                </select>
                <button class="collab-remove-btn" onclick="removeCollaborator('${c.user_key}')" title="제거">&times;</button>
                ` : `<span class="collab-role-label" style="font-size:12px;color:var(--text-secondary);">${c.role === 'editor' ? '편집자' : '뷰어'}</span>`}
            </div>
        `);
    });
}

let _collabSearchIndex = -1;

function initCollabSearch() {
    _collabSearchIndex = -1;

    $('#collabUserSearch').off('input').on('input', function() {
        const query = $(this).val().trim();
        clearTimeout(_collabSearchTimer);
        _collabSearchIndex = -1;
        if (query.length < 1) {
            $('#collabSearchDropdown').hide();
            return;
        }
        _collabSearchTimer = setTimeout(async () => {
            try {
                const res = await fetch('/api/auth/search-users', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ name: query }),
                });
                const data = await res.json();
                const users = data.users || [];
                renderCollabSearchDropdown(users);
            } catch (e) {}
        }, 300);
    });

    $('#collabUserSearch').off('keydown').on('keydown', function(e) {
        const dd = $('#collabSearchDropdown');
        if (!dd.is(':visible')) return;
        const items = dd.find('.search-item');
        if (items.length === 0) return;

        if (e.key === 'ArrowDown') {
            e.preventDefault();
            _collabSearchIndex = Math.min(_collabSearchIndex + 1, items.length - 1);
            highlightCollabItem(items);
        } else if (e.key === 'ArrowUp') {
            e.preventDefault();
            _collabSearchIndex = Math.max(_collabSearchIndex - 1, 0);
            highlightCollabItem(items);
        } else if (e.key === 'Enter') {
            e.preventDefault();
            // 아무 항목도 선택 안 된 상태에서 Enter 시 첫 번째 항목 선택
            if (_collabSearchIndex < 0 && items.length > 0) {
                _collabSearchIndex = 0;
            }
            if (_collabSearchIndex >= 0 && _collabSearchIndex < items.length) {
                selectCollabUser(_collabSearchIndex);
            }
        } else if (e.key === 'Escape') {
            dd.hide();
            _collabSearchIndex = -1;
        }
    });

    // 이벤트 위임 방식으로 드롭다운 항목 클릭 처리 (동적 생성 요소 대응)
    $('#collabSearchDropdown').off('mousedown.collabSelect click.collabSelect')
        .on('mousedown.collabSelect', '.search-item', function(e) {
            e.preventDefault();
            e.stopPropagation();
        })
        .on('click.collabSelect', '.search-item', function(e) {
            e.preventDefault();
            e.stopPropagation();
            const idx = parseInt($(this).data('index'), 10);
            selectCollabUser(idx);
        });

    // 드롭다운 외부 클릭 시 닫기
    $(document).off('click.collabSearch').on('click.collabSearch', function(e) {
        if (!$(e.target).closest('.collab-search-wrap').length) {
            $('#collabSearchDropdown').hide();
            _collabSearchIndex = -1;
        }
    });
}

function highlightCollabItem(items) {
    items.removeClass('highlighted');
    if (_collabSearchIndex >= 0 && _collabSearchIndex < items.length) {
        const $target = $(items[_collabSearchIndex]);
        $target.addClass('highlighted');
        // 스크롤 자동 이동
        const dd = $('#collabSearchDropdown');
        const itemTop = $target.position().top;
        const itemHeight = $target.outerHeight();
        const ddHeight = dd.height();
        const scrollTop = dd.scrollTop();
        if (itemTop + itemHeight > ddHeight) {
            dd.scrollTop(scrollTop + itemTop + itemHeight - ddHeight + 4);
        } else if (itemTop < 0) {
            dd.scrollTop(scrollTop + itemTop - 4);
        }
    }
}

function renderCollabSearchDropdown(users) {
    const dd = $('#collabSearchDropdown');
    dd.empty();
    _collabSearchIndex = -1;
    if (users.length === 0) {
        dd.html('<div class="search-no-result">검색 결과가 없습니다</div>');
        dd.data('users', []);
        dd.show();
        return;
    }
    users.forEach((u, i) => {
        const initial = (u.nm || '?').charAt(0);
        dd.append(`
            <div class="search-item" data-index="${i}">
                <div class="search-avatar">${escapeHtml(initial)}</div>
                <div class="search-info">
                    <div class="search-name">${escapeHtml(u.nm)}</div>
                    <div class="search-meta">
                        ${u.dp ? `<span class="search-dept">${escapeHtml(u.dp)}</span>` : ''}
                        ${u.em ? `<span class="search-email">${escapeHtml(u.em)}</span>` : ''}
                    </div>
                </div>
            </div>
        `);
    });
    // 이벤트 위임은 initCollabSearch()에서 처리
    dd.data('users', users);
    dd.show();
}

function selectCollabUser(index) {
    const users = $('#collabSearchDropdown').data('users') || [];
    const user = users[index];
    if (!user) return;
    $('#collabUserSearch').val(user.nm);
    $('#collabUserKey').val(user.ky);
    $('#collabSearchDropdown').hide();
    _collabSearchIndex = -1;
    addCollaborator();
}

async function addCollaborator() {
    const userKey = $('#collabUserKey').val();
    const role = $('#collabRoleSelect').val();
    if (!userKey) { showToast('사용자를 선택하세요', 'error'); return; }
    try {
        await apiPost('/api/projects/' + state.currentProject._id + '/collaborators', {
            user_key: userKey, role: role
        });
        const res = await apiGet('/api/projects/' + state.currentProject._id + '/collaborators');
        state.collaborators = res.collaborators || [];
        renderCollabList();
        showToast('협업자가 추가되었습니다', 'success');
        $('#collabUserSearch').val('');
        $('#collabUserKey').val('');

        // 폴링 시작 (아직 안 하고 있으면)
        state.isCollabProject = true;
        if (!state.pollInterval) startCollabPolling(state.currentProject._id);
        updateCollabUI();
    } catch (e) {
        showToast(e.message, 'error');
    }
}

async function updateCollabRole(userKey, newRole) {
    try {
        await apiPut('/api/projects/' + state.currentProject._id + '/collaborators/' + userKey, {
            role: newRole
        });
        showToast('역할이 변경되었습니다', 'success');
    } catch (e) {
        showToast(e.message, 'error');
    }
}

async function removeCollaborator(userKey) {
    if (!confirm('협업자를 제거하시겠습니까?')) return;
    try {
        await apiDelete('/api/projects/' + state.currentProject._id + '/collaborators/' + userKey);
        const res = await apiGet('/api/projects/' + state.currentProject._id + '/collaborators');
        state.collaborators = res.collaborators || [];
        renderCollabList();
        showToast('협업자가 제거되었습니다', 'success');

        // 협업자가 0이면 폴링 중지
        if (state.collaborators.length === 0 && state.collabRole === 'owner') {
            state.isCollabProject = false;
            stopCollabPolling();
        }
        updateCollabUI();
    } catch (e) {
        showToast(e.message, 'error');
    }
}

// ── 변경 이력 ──

async function showHistoryPanel() {
    if (!state.currentProject) return;
    try {
        const res = await apiGet('/api/projects/' + state.currentProject._id + '/history');
        renderHistoryList(res.history || []);
        $('#historyPanel').show();
    } catch (e) {
        showToast(e.message, 'error');
    }
}

function renderHistoryList(history) {
    const $list = $('#historyList');
    $list.empty();
    if (history.length === 0) {
        $list.html('<div style="padding:24px;text-align:center;color:var(--text-tertiary);font-size:13px;">변경 이력이 없습니다</div>');
        return;
    }
    history.forEach(h => {
        const time = h.created_at ? new Date(h.created_at).toLocaleString() : '';
        const canRevert = state.collabRole === 'owner' && h.action !== 'revert';
        $list.append(`
            <div class="history-item">
                <div class="history-item-header">
                    <span>${escapeHtml(h.user_name || '')}</span>
                    <span>${time}</span>
                </div>
                <div class="history-item-desc">${escapeHtml(h.description || h.action || '')}</div>
                ${canRevert ? `<button class="history-revert-btn" onclick="revertHistory('${h._id}')">되돌리기</button>` : ''}
            </div>
        `);
    });
}

async function revertHistory(historyId) {
    if (!confirm('이 시점으로 슬라이드를 되돌리겠습니까?')) return;
    try {
        const res = await apiPost('/api/projects/' + state.currentProject._id + '/history/' + historyId + '/revert', {});
        if (res.slide_id) {
            await reloadChangedSlides([res.slide_id]);
        }
        showToast('되돌리기 완료', 'success');
        // 이력 새로고침
        await showHistoryPanel();
    } catch (e) {
        showToast(e.message, 'error');
    }
}

// ── 페이지 종료 시 정리 (pagehide 사용 - unload Permissions Policy 위반 방지) ──

window.addEventListener('pagehide', function() {
    stopCollabPolling();
    if (state.onlyofficeEditor) {
        try { state.onlyofficeEditor.destroyEditor(); } catch (e) { }
        state.onlyofficeEditor = null;
    }
});

// 협업 검색 초기화 (DOM ready 시)
$(function() {
    initCollabSearch();
});

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


// ============ 엑셀 워크스페이스 ============

let _univerScriptsLoaded = false;
let _univerScriptsLoading = false;

// Progressive Excel streaming state
let _excelStreamBuffer = '';
let _excelSheetStates = []; // [{name, columnsSet, parsedRowCount}] per sheet
let _isExcelModifying = false; // 수정 모드 플래그 (progressive parsing 스킵)

async function loadUniverScripts() {
    if (_univerScriptsLoaded) return;
    if (_univerScriptsLoading) {
        // 이미 로딩 중이면 완료될 때까지 대기
        while (_univerScriptsLoading) {
            await new Promise(r => setTimeout(r, 100));
        }
        return;
    }
    _univerScriptsLoading = true;

    const scripts = [
        'https://unpkg.com/react@18.3.1/umd/react.production.min.js',
        'https://unpkg.com/react-dom@18.3.1/umd/react-dom.production.min.js',
        'https://unpkg.com/rxjs/dist/bundles/rxjs.umd.min.js',
        'https://unpkg.com/@univerjs/presets/lib/umd/index.js',
        'https://unpkg.com/@univerjs/preset-sheets-core/lib/umd/index.js',
        'https://unpkg.com/@univerjs/preset-sheets-core/lib/umd/locales/en-US.js',
    ];

    for (const src of scripts) {
        // 이미 로드된 스크립트 건너뛰기
        if (document.querySelector(`script[src="${src}"]`)) continue;
        await new Promise((resolve, reject) => {
            const s = document.createElement('script');
            s.src = src;
            s.onload = resolve;
            s.onerror = reject;
            document.head.appendChild(s);
        });
    }

    // CSS 로드
    if (!document.querySelector('link[href*="preset-sheets-core"]')) {
        const link = document.createElement('link');
        link.rel = 'stylesheet';
        link.href = 'https://unpkg.com/@univerjs/preset-sheets-core/lib/index.css';
        document.head.appendChild(link);
    }

    _univerScriptsLoaded = true;
    _univerScriptsLoading = false;
}

async function initExcelWorkspace() {
    try {
        await loadUniverScripts();

        // 기존 생성 데이터가 있으면 스타일 포함하여 채우기
        if (state.generatedExcel && state.generatedExcel.sheets) {
            populateUniverFromData(state.generatedExcel);
            renderExcelCharts(state.generatedExcel);
            $('#btnDownloadXlsx').show();
            $('#btnShareExcel').show();
            $('#btnModifyExcel').show().find('span').text(t('modifyExcel'));
            if (state.generatedExcel.meta && state.generatedExcel.meta.title) {
                $('#excelTitle').text(state.generatedExcel.meta.title);
            }
            $('#instructionsInput').attr('placeholder', t('excelModifyPlaceholder'));
        } else {
            _destroyExcelCharts();
            initUniver();
            $('#btnDownloadXlsx').hide();
            $('#btnShareExcel').hide();
            $('#btnModifyExcel').hide();
            $('#excelTitle').text(t('typeExcel'));
        }
    } catch (e) {
        console.error('[Excel] Univer 로딩 실패:', e);
        $('#univerContainer').html('<div style="padding:40px;text-align:center;color:#999;">스프레드시트 컴포넌트를 불러오지 못했습니다.</div>');
    }
}

/**
 * Univer 인스턴스 초기화
 * @param {Object} [workbookData] - 스타일 포함 워크북 데이터. 없으면 빈 워크북 생성
 */
function initUniver(workbookData) {
    let container = document.getElementById('univerContainer');
    if (!container) return;

    // 이전 인스턴스 안전 정리
    if (state.univerAPI) {
        const oldAPI = state.univerAPI;
        state.univerAPI = null;

        // 새 컨테이너 먼저 준비 (DOM 교체 전에 dispose하면 React contains 에러 발생)
        const fresh = container.cloneNode(false);
        container.parentNode.replaceChild(fresh, container);
        container = fresh;

        // DOM 분리 후 dispose (에러가 나도 old DOM에서만 발생, 새 DOM에 영향 없음)
        try { oldAPI.dispose(); } catch(e) {}
    } else {
        container.innerHTML = '';
    }

    const { createUniver } = UniverPresets;
    const { LocaleType, mergeLocales } = UniverCore;
    const { UniverSheetsCorePreset } = UniverPresetSheetsCore;

    const { univerAPI } = createUniver({
        locale: LocaleType.EN_US,
        locales: {
            [LocaleType.EN_US]: mergeLocales(UniverPresetSheetsCoreEnUS),
        },
        presets: [
            UniverSheetsCorePreset({ container }),
        ],
    });

    state.univerAPI = univerAPI;
    const defaultWb = workbookData || {
        name: 'Workbook',
        sheetOrder: ['sheet1'],
        sheets: {
            sheet1: {
                id: 'sheet1',
                name: 'Sheet1',
                defaultColumnWidth: 100,
                defaultRowHeight: 24,
                rowCount: 200,
                columnCount: 26,
            },
        },
    };
    univerAPI.createWorkbook(defaultWb);
}

/**
 * 엑셀 데이터로 스타일 포함 워크북 생성
 * (다운로드 XLSX와 동일한 스타일: 파란 헤더, 교차 행 색상, 테두리)
 */
function populateUniverFromData(excelData) {
    if (!excelData) return;
    const workbookData = _buildStyledWorkbookData(excelData);
    if (!workbookData) return;
    initUniver(workbookData);
}

// 공통 셀 스타일 정의 (다운로드 XLSX와 동일)
const _EXCEL_STYLES = {
    header: {
        bg: { rgb: '#4472C4' },
        cl: { rgb: '#FFFFFF' },
        bl: 1,
        fs: 11,
        ff: 'Arial',
        ht: 2,
        vt: 2,
        tb: 3,
        bd: {
            t: { s: 1, cl: { rgb: '#D9D9D9' } },
            b: { s: 1, cl: { rgb: '#D9D9D9' } },
            l: { s: 1, cl: { rgb: '#D9D9D9' } },
            r: { s: 1, cl: { rgb: '#D9D9D9' } },
        },
    },
    cell: {
        vt: 2,
        tb: 3,
        bd: {
            t: { s: 1, cl: { rgb: '#D9D9D9' } },
            b: { s: 1, cl: { rgb: '#D9D9D9' } },
            l: { s: 1, cl: { rgb: '#D9D9D9' } },
            r: { s: 1, cl: { rgb: '#D9D9D9' } },
        },
    },
    cellAlt: {
        bg: { rgb: '#F2F7FB' },
        vt: 2,
        tb: 3,
        bd: {
            t: { s: 1, cl: { rgb: '#D9D9D9' } },
            b: { s: 1, cl: { rgb: '#D9D9D9' } },
            l: { s: 1, cl: { rgb: '#D9D9D9' } },
            r: { s: 1, cl: { rgb: '#D9D9D9' } },
        },
    },
};

/**
 * 엑셀 데이터 → Univer IWorkbookData 변환 (스타일, 열 너비, 시트 탭 포함)
 */
function _buildStyledWorkbookData(excelData) {
    const sheets = excelData.sheets || [];
    if (sheets.length === 0) return null;

    const sheetConfigs = {};
    const sheetOrder = [];

    for (let si = 0; si < sheets.length; si++) {
        const sheetData = sheets[si];
        const columns = sheetData.columns || [];
        const rows = sheetData.rows || [];
        const cellData = {};

        // 헤더 행 (row 0)
        cellData[0] = {};
        for (let c = 0; c < columns.length; c++) {
            cellData[0][c] = { v: columns[c], s: _EXCEL_STYLES.header };
        }

        // 데이터 행
        for (let r = 0; r < rows.length; r++) {
            const rowIdx = r + 1;
            cellData[rowIdx] = {};
            const style = (rowIdx % 2 === 0) ? _EXCEL_STYLES.cellAlt : _EXCEL_STYLES.cell;
            for (let c = 0; c < rows[r].length; c++) {
                cellData[rowIdx][c] = { v: rows[r][c], s: style };
            }
        }

        // 열 너비 자동 계산 (한글 2배 너비)
        const columnData = {};
        const allRows = [columns, ...rows];
        for (let c = 0; c < columns.length; c++) {
            let maxLen = 0;
            for (let r = 0; r < Math.min(allRows.length, 50); r++) {
                if (allRows[r] && allRows[r][c] != null) {
                    const val = String(allRows[r][c]);
                    const charLen = [...val].reduce((sum, ch) => sum + (ch.charCodeAt(0) > 127 ? 2 : 1), 0);
                    maxLen = Math.max(maxLen, charLen);
                }
            }
            columnData[c] = { w: Math.min(Math.max((maxLen + 4) * 8, 80), 480) };
        }

        const sheetId = `sheet_${si}`;
        sheetOrder.push(sheetId);
        sheetConfigs[sheetId] = {
            id: sheetId,
            name: (sheetData.name || `Sheet${si + 1}`).substring(0, 31),
            cellData: cellData,
            columnData: columnData,
            defaultColumnWidth: 100,
            defaultRowHeight: 24,
            rowCount: Math.max(rows.length + 100, 200),
            columnCount: Math.max(columns.length + 10, 26),
        };
    }

    return {
        name: excelData.meta?.title || 'Workbook',
        sheetOrder: sheetOrder,
        sheets: sheetConfigs,
    };
}

/**
 * 스트리밍 중 수신된 텍스트를 점진적으로 파싱하여 Univer 시트에 데이터 추가
 * 다중 시트 지원: "sheets" 배열 내 각 시트 블록을 개별 파싱
 */
function _tryProgressiveExcelPopulate() {
    if (!state.univerAPI) return;

    const workbook = state.univerAPI.getActiveWorkbook();
    if (!workbook) return;

    const text = _excelStreamBuffer;

    // "sheets": [ 위치 찾기
    const sheetsArrayMatch = text.match(/"sheets"\s*:\s*\[/);
    if (!sheetsArrayMatch) return;
    const sheetsArrayStart = text.indexOf(sheetsArrayMatch[0]) + sheetsArrayMatch[0].length;
    const sheetsText = text.substring(sheetsArrayStart);

    // 각 시트 오브젝트의 시작 위치 찾기 ({ ... "name": ... })
    const sheetStarts = [];
    let searchFrom = 0;
    while (true) {
        const nameIdx = sheetsText.indexOf('"name"', searchFrom);
        if (nameIdx < 0) break;
        // "name" 앞의 { 찾기
        let braceIdx = nameIdx;
        while (braceIdx > searchFrom && sheetsText[braceIdx] !== '{') braceIdx--;
        if (sheetsText[braceIdx] === '{') {
            sheetStarts.push(braceIdx);
        }
        searchFrom = nameIdx + 6;
    }

    if (sheetStarts.length === 0) return;

    // 각 시트 블록 처리
    for (let si = 0; si < sheetStarts.length; si++) {
        const blockStart = sheetStarts[si];
        const blockEnd = si < sheetStarts.length - 1 ? sheetStarts[si + 1] : sheetsText.length;
        const sheetBlock = sheetsText.substring(blockStart, blockEnd);

        // 새 시트 감지 → 상태 초기화 + Univer 시트 탭 생성
        if (si >= _excelSheetStates.length) {
            _excelSheetStates.push({ name: '', columnsSet: false, parsedRowCount: 0 });

            // 첫 번째 시트는 이미 존재, 두 번째부터 새 시트 생성
            if (si > 0) {
                try {
                    const tempName = 'Sheet' + (si + 1);
                    const newSheet = workbook.create(tempName);
                    if (newSheet && newSheet.activate) {
                        newSheet.activate();
                    }
                } catch (e) {
                    console.log('[Excel] 시트 생성:', e.message);
                }
            }
        }

        const ss = _excelSheetStates[si];

        // 현재 시트가 아닌 이전 시트는 스킵 (이미 처리 완료)
        if (si < _excelSheetStates.length - 1) continue;

        // 활성 시트 가져오기
        const sheet = workbook.getActiveSheet();
        if (!sheet) continue;

        // 시트 이름 추출
        if (!ss.name) {
            const nameMatch = sheetBlock.match(/"name"\s*:\s*"([^"]+)"/);
            if (nameMatch) {
                ss.name = nameMatch[1];
                try { sheet.setName(ss.name); } catch (e) {}
            }
        }

        // 컬럼 헤더 추출 + 스타일 + 열 너비 적용
        if (!ss.columnsSet) {
            const colMatch = sheetBlock.match(/"columns"\s*:\s*\[([^\]]*)\]/);
            if (colMatch) {
                try {
                    const columns = JSON.parse('[' + colMatch[1] + ']');
                    if (columns.length > 0) {
                        for (let c = 0; c < columns.length; c++) {
                            try { sheet.getRange(0, c, 1, 1).setValue(columns[c]); } catch (e2) {}
                            // 열 너비 설정 (헤더 텍스트 기반, 최소 100px)
                            try {
                                const hdrLen = [...String(columns[c])].reduce((s, ch) => s + (ch.charCodeAt(0) > 127 ? 2 : 1), 0);
                                const w = Math.min(Math.max((hdrLen + 4) * 8, 100), 480);
                                sheet.setColumnWidth(c, w);
                            } catch (e2) {}
                        }
                        try {
                            sheet.getRange(0, 0, 1, columns.length)
                                .setFontWeight('bold').setFontColor('#FFFFFF').setFontSize(11);
                        } catch (e2) {}
                        try {
                            sheet.setRowDefaultStyle(0, { bg: { rgb: '#4472C4' }, ht: 2, vt: 2 });
                        } catch (e2) {}
                        ss.columnsSet = true;
                    }
                } catch (e) {}
            }
        }

        // rows 데이터를 점진적으로 추출
        if (ss.columnsSet) {
            const rowsMatch = sheetBlock.match(/"rows"\s*:\s*\[/);
            if (rowsMatch) {
                const rowsStartIdx = sheetBlock.indexOf(rowsMatch[0]) + rowsMatch[0].length;
                const rowsText = sheetBlock.substring(rowsStartIdx);
                const rows = _extractCompleteRows(rowsText);

                if (rows.length > ss.parsedRowCount) {
                    for (let i = ss.parsedRowCount; i < rows.length; i++) {
                        const rowIdx = i + 1;
                        for (let c = 0; c < rows[i].length; c++) {
                            try { sheet.getRange(rowIdx, c, 1, 1).setValue(rows[i][c]); } catch (e) {}
                        }
                        if (rowIdx % 2 === 0) {
                            try { sheet.setRowDefaultStyle(rowIdx, { bg: { rgb: '#F2F7FB' } }); } catch (e) {}
                        }
                    }
                    ss.parsedRowCount = rows.length;
                }
            }
        }
    }
}

/**
 * 부분 텍스트에서 완성된 JSON 배열([...])을 추출
 * "rows": [ 뒤의 텍스트를 받아 각 [...] 행을 파싱
 */
function _extractCompleteRows(text) {
    const rows = [];
    let depth = 0;
    let startIdx = -1;
    let inString = false;
    let escapeNext = false;

    for (let i = 0; i < text.length; i++) {
        const ch = text[i];

        if (escapeNext) {
            escapeNext = false;
            continue;
        }

        if (ch === '\\' && inString) {
            escapeNext = true;
            continue;
        }

        if (ch === '"') {
            inString = !inString;
            continue;
        }

        if (inString) continue;

        if (ch === '[') {
            if (depth === 0) startIdx = i;
            depth++;
        } else if (ch === ']') {
            depth--;
            if (depth === 0 && startIdx >= 0) {
                // 완성된 행 배열 발견
                const rowStr = text.substring(startIdx, i + 1);
                try {
                    const row = JSON.parse(rowStr);
                    if (Array.isArray(row)) {
                        // 문자열 숫자를 실제 숫자로 변환
                        for (let c = 0; c < row.length; c++) {
                            if (typeof row[c] === 'string') {
                                const num = Number(row[c]);
                                if (!isNaN(num) && row[c].trim() !== '') {
                                    row[c] = num;
                                }
                            }
                        }
                        rows.push(row);
                    }
                } catch (e) {}
                startIdx = -1;
            }
            if (depth < 0) break; // rows 배열 종료
        }
    }

    return rows;
}

function handleGenerate() {
    const type = state.currentProject ? state.currentProject.project_type : 'slide';
    if (type === 'excel') {
        generateExcel();
    } else if (type === 'onlyoffice_pptx') {
        generateOnlyOfficePptx();
    } else if (type === 'onlyoffice_xlsx') {
        generateOnlyOfficeXlsx();
    } else if (type === 'onlyoffice_docx') {
        generateOnlyOfficeDocx();
    } else if (type === 'word') {
        if (state.generatedDocx && state.generatedDocx.sections && state.generatedDocx.sections.length > 0) {
            modifyWord();
        } else {
            generateWord();
        }
    } else {
        generatePPT();
    }
}

async function generateExcel() {
    const instructions = $('#instructionsInput').val().trim();
    const lang = $('#langSelect').val();

    // 리소스가 없으면 지침(instructions)이 필수 (인터넷 검색에 사용)
    if (state.resources.length === 0 && !instructions) {
        showToast(t('msgEnterInstructions'), 'error');
        return;
    }

    _isGenerating = true;
    state.generatedExcel = null;
    $('#btnDownloadXlsx').hide();

    // 기존 차트 정리
    _destroyExcelCharts();

    // 생성 중 상태 표시
    _showStopButton();
    _showExcelProgress(t('excelPreparing'), '');

    // Progressive parsing 상태 초기화
    _excelStreamBuffer = '';
    _excelSheetStates = [];

    // Univer를 빈 시트로 즉시 초기화 (실시간 데이터 표시용)
    initUniver();

    _abortController = new AbortController();

    try {
        const response = await fetch(`/${state.jwtToken}/api/generate/excel/stream`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                project_id: state.currentProject._id,
                instructions: instructions,
                lang: lang,
            }),
            signal: _abortController.signal,
        });

        if (!response.ok) {
            const err = await response.json().catch(() => ({}));
            throw new Error(err.detail || 'Generation failed');
        }

        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;

            buffer += decoder.decode(value, { stream: true });
            const lines = buffer.split('\n');
            buffer = lines.pop() || '';

            for (const line of lines) {
                if (!line.startsWith('data: ')) continue;
                try {
                    const evt = JSON.parse(line.slice(6));
                    _handleExcelSSEEvent(evt);
                } catch (e) { /* ignore parse errors */ }
            }
        }

        // 남은 버퍼 처리
        if (buffer.startsWith('data: ')) {
            try {
                const evt = JSON.parse(buffer.slice(6));
                _handleExcelSSEEvent(evt);
            } catch (e) {}
        }

    } catch (e) {
        if (e.name !== 'AbortError') {
            console.error('[Excel] 생성 실패:', e);
            showToast(e.message || '엑셀 생성 실패', 'error');
        }
    } finally {
        _isGenerating = false;
        _abortController = null;
        _hideExcelProgress();
        _showGenerateOrRestartButton();
    }
}

function _detectChartRequest(instruction) {
    // 차트/그래프 키워드 확인
    const chartKeywords = ['차트', 'chart', '그래프', 'graph'];
    const isChartRequest = chartKeywords.some(kw => instruction.toLowerCase().includes(kw));
    if (!isChartRequest) return null;

    // 생성 키워드 확인 (삭제/수정 요청은 LLM으로)
    const createKeywords = ['생성', '만들', '추가', '그려', '그리', 'create', 'add', 'generate', 'make', 'draw'];
    const isCreate = createKeywords.some(kw => instruction.toLowerCase().includes(kw));
    if (!isCreate) return null;

    // 차트 타입 감지
    const chartTypeMap = [
        [['막대', 'bar', '바 차트', '바차트', '세로 막대'], 'bar'],
        [['선형', '라인', 'line', '꺾은선', '선 그래프', '추세'], 'line'],
        [['원형', '파이', 'pie', '원 그래프'], 'pie'],
        [['영역', 'area'], 'area'],
        [['산점', 'scatter', '산포'], 'scatter'],
        [['도넛', 'doughnut', '도너츠'], 'doughnut'],
        [['방사', '레이더', 'radar', '레이다'], 'radar'],
    ];

    let chartType = 'bar'; // 기본값
    const lowerInst = instruction.toLowerCase();
    for (const [keywords, type] of chartTypeMap) {
        if (keywords.some(kw => lowerInst.includes(kw))) {
            chartType = type;
            break;
        }
    }

    // 시트 인덱스 감지
    let sheetIndex = 0;
    // "1번 시트", "2번시트", "첫번째 시트" 등
    const sheetPatterns = [
        { regex: /(\d+)\s*번\s*시트/, handler: m => parseInt(m[1]) - 1 },
        { regex: /시트\s*(\d+)/, handler: m => parseInt(m[1]) - 1 },
        { regex: /sheet\s*(\d+)/i, handler: m => parseInt(m[1]) - 1 },
        { regex: /첫\s*번째/, handler: () => 0 },
        { regex: /두\s*번째/, handler: () => 1 },
        { regex: /세\s*번째/, handler: () => 2 },
    ];
    for (const { regex, handler } of sheetPatterns) {
        const match = instruction.match(regex);
        if (match) {
            sheetIndex = handler(match);
            break;
        }
    }

    // 활성 시트 인덱스 fallback
    if (sheetIndex === 0 && state.univerAPI) {
        try {
            const wb = state.univerAPI.getActiveWorkbook();
            if (wb) {
                const activeSheet = wb.getActiveSheet();
                if (activeSheet) {
                    const sheetId = activeSheet.getSheetId();
                    const m = sheetId && sheetId.match(/sheet_(\d+)/);
                    if (m) sheetIndex = parseInt(m[1], 10);
                }
            }
        } catch (e) {}
    }

    // 차트 제목 추출 (따옴표로 감싼 제목)
    let title = null;
    const titleMatch = instruction.match(/[""']([^""']+)[""']/);
    if (titleMatch) title = titleMatch[1];

    return { sheetIndex, chartType, title };
}

async function _generateChartDirect(chartReq) {
    _isGenerating = true;
    _showStopButton();
    _showExcelProgress('차트를 생성하고 있습니다...', '');

    try {
        const response = await fetch(`/${state.jwtToken}/api/generate/excel/chart`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                project_id: state.currentProject._id,
                sheet_index: chartReq.sheetIndex,
                chart_type: chartReq.chartType,
                title: chartReq.title,
            }),
        });

        if (!response.ok) {
            const err = await response.json().catch(() => ({}));
            throw new Error(err.detail || '차트 생성 실패');
        }

        const result = await response.json();
        if (result.success && result.excel) {
            state.generatedExcel = result.excel;

            // Univer 리로드 + 차트 렌더링
            await loadUniverScripts();
            populateUniverFromData(state.generatedExcel);
            renderExcelCharts(state.generatedExcel);

            // 차트 이미지가 있으면 차트 영역에 이미지도 표시
            if (result.chart_image_url) {
                _appendChartImage(result.chart_image_url);
            }

            $('#btnDownloadXlsx').show();
            $('#btnShareExcel').show();
            if (state.generatedExcel.meta && state.generatedExcel.meta.title) {
                $('#excelTitle').text(state.generatedExcel.meta.title);
            }

            showToast('차트가 생성되었습니다', 'success');
        }
    } catch (e) {
        console.error('[Excel] 차트 생성 실패:', e);
        showToast(e.message || '차트 생성 실패', 'error');
    } finally {
        _isGenerating = false;
        _hideExcelProgress();
        _showGenerateOrRestartButton();
        $('#instructionsInput').val('');
        autoResizeTextarea(document.getElementById('instructionsInput'));
    }
}

function _appendChartImage(imageUrl) {
    // 차트 이미지를 차트 컨테이너에 추가 표시
    const container = document.getElementById('excelChartsContainer');
    const grid = document.getElementById('excelChartsGrid');
    if (!container || !grid) return;

    const card = document.createElement('div');
    card.className = 'excel-chart-card';
    card.style.cssText = 'text-align:center;';

    const img = document.createElement('img');
    img.src = imageUrl;
    img.style.cssText = 'max-width:100%;height:auto;border-radius:6px;';
    img.alt = 'Generated Chart';

    card.appendChild(img);
    grid.appendChild(card);
    $(container).show();
}

async function modifyExcel() {
    const instruction = $('#instructionsInput').val().trim();
    if (!instruction) {
        showToast(t('msgEnterInstructions'), 'error');
        return;
    }

    if (!state.generatedExcel || !state.generatedExcel.sheets) {
        showToast('수정할 엑셀 데이터가 없습니다.', 'error');
        return;
    }

    // 차트 생성 요청 감지 → LLM 없이 직접 생성
    const chartReq = _detectChartRequest(instruction);
    if (chartReq) {
        await _generateChartDirect(chartReq);
        return;
    }

    _isGenerating = true;
    _isExcelModifying = true;
    _destroyExcelCharts();
    _showStopButton();
    _showExcelProgress(t('excelModifying') || 'AI가 데이터를 수정하고 있습니다...', '');

    _excelStreamBuffer = '';
    _excelSheetStates = [];
    // 수정 모드: Univer 초기화 안 함 - 기존 화면 유지

    _abortController = new AbortController();

    try {
        // 현재 활성 시트 인덱스 탐지
        let targetSheetIndex = null;
        if (state.univerAPI) {
            try {
                const wb = state.univerAPI.getActiveWorkbook();
                if (wb) {
                    const activeSheet = wb.getActiveSheet();
                    if (activeSheet) {
                        const sheetId = activeSheet.getSheetId();
                        // sheetId 형식: "sheet_0", "sheet_1", ...
                        const match = sheetId && sheetId.match(/sheet_(\d+)/);
                        if (match) targetSheetIndex = parseInt(match[1], 10);
                    }
                }
            } catch (e) {
                console.warn('[Excel] 활성 시트 인덱스 탐지 실패:', e);
            }
        }

        const response = await fetch(`/${state.jwtToken}/api/generate/excel/modify/stream`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                project_id: state.currentProject._id,
                instruction: instruction,
                current_data: {
                    meta: state.generatedExcel.meta || {},
                    sheets: state.generatedExcel.sheets || [],
                },
                lang: $('#langSelect').val(),
                target_sheet_index: targetSheetIndex,
            }),
            signal: _abortController.signal,
        });

        if (!response.ok) {
            const err = await response.json().catch(() => ({}));
            throw new Error(err.detail || 'Modification failed');
        }

        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;

            buffer += decoder.decode(value, { stream: true });
            const lines = buffer.split('\n');
            buffer = lines.pop() || '';

            for (const line of lines) {
                if (!line.startsWith('data: ')) continue;
                try {
                    const evt = JSON.parse(line.slice(6));
                    _handleExcelSSEEvent(evt);
                } catch (e) {}
            }
        }

        if (buffer.startsWith('data: ')) {
            try {
                const evt = JSON.parse(buffer.slice(6));
                _handleExcelSSEEvent(evt);
            } catch (e) {}
        }

    } catch (e) {
        if (e.name !== 'AbortError') {
            console.error('[Excel] 수정 실패:', e);
            showToast(e.message || '엑셀 수정 실패', 'error');
        }
    } finally {
        _isGenerating = false;
        _isExcelModifying = false;
        _abortController = null;
        _hideExcelProgress();
        _showGenerateOrRestartButton();
    }
}

async function uploadExcelFile(event) {
    const file = event.target.files[0];
    if (!file) return;

    const ext = file.name.split('.').pop().toLowerCase();
    if (!['xlsx', 'xls'].includes(ext)) {
        showToast('xlsx 또는 xls 파일만 업로드 가능합니다', 'error');
        event.target.value = '';
        return;
    }

    if (!state.currentProject) {
        showToast('프로젝트를 먼저 선택하세요', 'error');
        event.target.value = '';
        return;
    }

    _destroyExcelCharts();
    _showExcelProgress(t('excelUploading') || '엑셀 파일 업로드 중...', '');

    const formData = new FormData();
    formData.append('file', file);
    formData.append('project_id', state.currentProject._id);

    try {
        const response = await fetch(`/${state.jwtToken}/api/generate/excel/upload`, {
            method: 'POST',
            body: formData,
        });

        if (!response.ok) {
            const err = await response.json().catch(() => ({}));
            throw new Error(err.detail || '업로드 실패');
        }

        const result = await response.json();
        if (result.success && result.excel) {
            state.generatedExcel = result.excel;

            // Univer에 데이터 표시
            await loadUniverScripts();
            populateUniverFromData(state.generatedExcel);
            renderExcelCharts(state.generatedExcel);

            // UI 업데이트
            $('#btnDownloadXlsx').show();
            $('#btnShareExcel').show();
            $('#btnModifyExcel').show().find('span').text(t('modifyExcel'));
            if (state.generatedExcel.meta && state.generatedExcel.meta.title) {
                $('#excelTitle').text(state.generatedExcel.meta.title);
            }

            // placeholder를 수정 모드로 변경
            $('#instructionsInput').val('').attr('placeholder', t('excelModifyPlaceholder'));
            autoResizeTextarea(document.getElementById('instructionsInput'));

            _showGenerateOrRestartButton();
            showToast(t('excelUploaded') || '엑셀 파일이 업로드되었습니다', 'success');
        }
    } catch (e) {
        console.error('[Excel] 업로드 실패:', e);
        showToast(e.message || '엑셀 업로드 실패', 'error');
    } finally {
        _hideExcelProgress();
        event.target.value = '';
    }
}

function _handleExcelSSEEvent(evt) {
    const eventType = evt.event;

    switch (eventType) {
        case 'start':
            _showExcelProgress(t('excelPreparing'), '');
            break;

        case 'searching':
            // 인터넷 검색 중 상태 표시
            _showExcelProgress(t('excelSearching'), '');
            if (!_isExcelModifying) {
                const oldAPI = state.univerAPI;
                state.univerAPI = null;
                // 컨테이너 먼저 교체 후 dispose (React contains 에러 방지)
                const oldC = document.getElementById('univerContainer');
                if (oldC) {
                    const freshC = oldC.cloneNode(false);
                    oldC.parentNode.replaceChild(freshC, oldC);
                    freshC.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;height:100%;flex-direction:column;gap:16px;color:#888;">' +
                        '<div class="dot-loading"><span></span><span></span><span></span></div>' +
                        '<p>' + t('excelSearching') + '</p></div>';
                }
                if (oldAPI) {
                    try { oldAPI.dispose(); } catch(e) {}
                }
            }
            break;

        case 'search_done':
            // 검색 완료 → Univer 다시 초기화하여 실시간 스트리밍 준비
            _showExcelProgress(t('excelSearchDone'), '');
            _excelStreamBuffer = '';
            _excelSheetStates = [];
            if (!_isExcelModifying) {
                initUniver();
            }
            break;

        case 'delta':
            _excelStreamBuffer += (evt.text || '');
            if (_isExcelModifying) {
                // 수정 모드: progressive parsing 스킵, 기존 화면 유지
                _updateExcelStreamProgress();
            } else {
                // 신규 생성: AI 스트리밍 텍스트 → 실시간 파싱하여 Univer에 점진적으로 표시
                _tryProgressiveExcelPopulate();
                _updateExcelStreamProgress();
            }
            break;

        case 'parsing':
            _showExcelProgress(t('excelFinalizing'), '');
            break;

        case 'excel_data':
            // 최종 데이터 수신 → 스타일 포함 워크북으로 갱신
            _hideExcelProgress();
            state.generatedExcel = evt.excel;
            state.currentProject.status = 'generated';
            $('#wsProjectStatus').text(t('statusGenerated')).attr('class', 'ws-status generated');

            // 최종 데이터로 스타일 포함 워크북 재생성
            populateUniverFromData(state.generatedExcel);

            // 차트 렌더링
            renderExcelCharts(state.generatedExcel);

            // 다운로드/수정 버튼 표시
            $('#btnDownloadXlsx').show();
            $('#btnShareExcel').show();
            $('#btnModifyExcel').show().find('span').text(t('modifyExcel'));
            if (state.generatedExcel.meta && state.generatedExcel.meta.title) {
                $('#excelTitle').text(state.generatedExcel.meta.title);
            }
            $('#instructionsInput').val('').attr('placeholder', t('excelModifyPlaceholder'));
            autoResizeTextarea(document.getElementById('instructionsInput'));
            break;

        case 'complete':
            _hideExcelProgress();
            showToast(t('msgExcelGenerated'), 'success');
            break;

        case 'stopped':
            _hideExcelProgress();
            showToast(evt.message || t('statusStopped'), 'info');
            state.currentProject.status = 'stopped';
            $('#wsProjectStatus').text(t('statusStopped')).attr('class', 'ws-status stopped');
            break;

        case 'error':
            _hideExcelProgress();
            showToast(evt.message || '생성 실패', 'error');
            break;
    }
}

// ============ 엑셀 진행 상태 UI ============

function _showExcelProgress(msg, detail) {
    $('#excelProgressMsg').text(msg);
    $('#excelProgressDetail').text(detail || '');
    $('#excelProgressBar').removeClass('determinate');
    $('#excelProgressOverlay').show();
}

function _hideExcelProgress() {
    $('#excelProgressOverlay').hide();
}

function _updateExcelStreamProgress() {
    const sheetCount = _excelSheetStates.length || 0;
    let totalRows = 0;
    for (const ss of _excelSheetStates) {
        totalRows += (ss.parsedRowCount || 0);
    }

    if (sheetCount === 0 && totalRows === 0) {
        _showExcelProgress(t('excelGenerating'), '');
    } else {
        const detail = t('excelStreamingDetail')
            .replace('{sheets}', sheetCount)
            .replace('{rows}', totalRows);
        _showExcelProgress(t('excelStreaming'), detail);
    }
}

function downloadXLSX() {
    if (!state.currentProject || !state.generatedExcel) {
        showToast(t('msgNoExcelData'), 'error');
        return;
    }
    const url = `/${state.jwtToken}/api/generate/${state.currentProject._id}/download/xlsx`;
    _downloadFile(url, (state.currentProject.name || 'spreadsheet') + '.xlsx');
}


// ============ 워드 문서 (CKEditor 5) ============

let _ckEditorInstance = null;
let _ckEditorScriptLoaded = false;
let _ckEditorCssLoaded = false;

async function loadCKEditorScript() {
    if (_ckEditorScriptLoaded) return;

    // CKEditor 5 CSS 로드
    if (!_ckEditorCssLoaded) {
        const link = document.createElement('link');
        link.rel = 'stylesheet';
        link.href = 'https://cdn.jsdelivr.net/npm/ckeditor5@44.1.0/dist/browser/ckeditor5.css';
        document.head.appendChild(link);
        _ckEditorCssLoaded = true;
    }

    // CKEditor 5 JS 로드
    await new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/npm/ckeditor5@44.1.0/dist/browser/ckeditor5.umd.js';
        script.onload = () => resolve();
        script.onerror = () => reject(new Error('CKEditor 5 스크립트 로드 실패'));
        document.head.appendChild(script);
    });

    // 한국어 번역 로드
    await new Promise((resolve) => {
        const script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/npm/ckeditor5@44.1.0/dist/translations/ko.umd.js';
        script.onload = () => resolve();
        script.onerror = () => { console.warn('[CKEditor] 한국어 번역 로드 실패, 기본 언어 사용'); resolve(); };
        document.head.appendChild(script);
    });

    _ckEditorScriptLoaded = true;
}

function _showWordProgress(msg, detail) {
    $('#wordProgressOverlay').show();
    $('#wordProgressMsg').text(msg);
    $('#wordProgressDetail').text(detail || '');
}

function _hideWordProgress() {
    $('#wordProgressOverlay').hide();
    $('#wordProgressBar').css('width', '0%');
}

// ====== 워드 실시간 스트리밍 프리뷰 ======

let _wordStreamBuffer = '';
let _wordDeltaCount = 0;
let _wordRequestedTotal = 0; // 사용자가 요청한 전체 섹션 수
let _wordStreamRenderTimer = null;
const _wordChartCache = new Map(); // 차트 이미지 캐시 (JSON → dataURL)
let _pendingChartConfigs = []; // 현재 렌더 사이클의 대기 차트 configs

function _showWordStreamingPreview() {
    _wordStreamBuffer = '';
    _wordChartCache.clear();
    _pendingChartConfigs = [];
    $('#ckeditorContainer').hide();
    $('#wordStreamingPreview').show().scrollTop(0);
    $('#wordStreamingContent').html('');
}

function _hideWordStreamingPreview() {
    if (_wordStreamRenderTimer) {
        clearTimeout(_wordStreamRenderTimer);
        _wordStreamRenderTimer = null;
    }
    $('#wordStreamingPreview').hide();
    $('#ckeditorContainer').show();
}

function _appendWordStreamDelta(text) {
    _wordStreamBuffer += text;
    // 렌더링 쓰로틀: 80ms 간격
    if (!_wordStreamRenderTimer) {
        _wordStreamRenderTimer = setTimeout(() => {
            _wordStreamRenderTimer = null;
            _renderWordStreamingContent();
        }, 80);
    }
}

function _renderWordStreamingContent() {
    const cleaned = _cleanStreamingJsonToReadable(_wordStreamBuffer);
    const html = _streamingMarkdownToHtml(cleaned);
    const container = document.getElementById('wordStreamingContent');
    if (!container) return;
    container.innerHTML = html + '<span class="word-typing-cursor"></span>';

    // 대기 중인 차트 렌더링
    _renderInlineStreamCharts(container);

    // 자동 스크롤
    const preview = document.getElementById('wordStreamingPreview');
    if (preview) {
        preview.scrollTop = preview.scrollHeight;
    }
}

function _renderInlineStreamCharts(container) {
    const pendingDivs = container.querySelectorAll('.docx-stream-chart-pending');
    if (!pendingDivs.length) return;

    // ECharts 미로드 → 로딩 트리거 후 다음 사이클에서 렌더링
    if (typeof echarts === 'undefined') {
        loadEChartsScript().catch(() => {});
        return;
    }

    pendingDivs.forEach(div => {
        const idx = parseInt(div.getAttribute('data-chart-idx'));
        if (isNaN(idx) || !_pendingChartConfigs[idx]) return;

        const jsonStr = _pendingChartConfigs[idx];

        try {
            const config = JSON.parse(jsonStr);
            const option = _buildEChartsOptionFromChartJson(config);

            // 오프스크린 div로 렌더링 후 이미지 추출
            const offDiv = document.createElement('div');
            offDiv.style.cssText = 'width:900px;height:450px;position:absolute;left:-9999px;top:-9999px;';
            document.body.appendChild(offDiv);

            const chart = echarts.init(offDiv);
            chart.setOption(option);

            const dataUrl = chart.getDataURL({ type: 'png', pixelRatio: 2, backgroundColor: '#fff' });
            chart.dispose();
            document.body.removeChild(offDiv);

            _wordChartCache.set(jsonStr, dataUrl);

            const safeJson = jsonStr.replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
            div.innerHTML = `<img src="${dataUrl}" style="max-width:100%;border-radius:8px;" data-chart-json="${safeJson}" />`;
            div.className = 'docx-stream-chart';
        } catch (e) {
            console.error('[Word Stream Chart] render error:', e);
        }
    });
}

function _cleanStreamingJsonToReadable(raw) {
    let text = raw;

    // 차트 블록 마커 보존 (```chart → @@CHARTBLOCK@@)
    text = text.replace(/```chart/g, '@@CHARTBLOCK@@');

    // JSON 코드 블록 마커 제거 (차트 외)
    text = text.replace(/```json\s*/g, '');
    text = text.replace(/```\s*/g, '');

    // meta 블록에서 제목과 설명 추출
    let metaTitle = '';
    let metaDesc = '';
    const metaTitleMatch = text.match(/"title"\s*:\s*"((?:[^"\\]|\\.)*)"/);
    if (metaTitleMatch) {
        metaTitle = _unescJsonStr(metaTitleMatch[1]);
    }
    const metaDescMatch = text.match(/"description"\s*:\s*"((?:[^"\\]|\\.)*)"/);
    if (metaDescMatch) {
        metaDesc = _unescJsonStr(metaDescMatch[1]);
    }

    // sections 배열에서 섹션 추출
    const sectionsIdx = text.indexOf('"sections"');
    if (sectionsIdx < 0) {
        // sections 시작 전 → 제목만 표시
        let result = '';
        if (metaTitle) result += '# ' + metaTitle + '\n\n';
        if (metaDesc) result += metaDesc + '\n\n---\n\n';
        return result || '';
    }

    const sectionsText = text.substring(sectionsIdx);
    let result = '';
    if (metaTitle) result += '# ' + metaTitle + '\n\n';
    if (metaDesc) result += metaDesc + '\n\n---\n\n';

    // 각 섹션 오브젝트 추출 (완전한 것 + 진행 중인 것)
    const sectionBlocks = _extractSectionBlocks(sectionsText);
    for (const block of sectionBlocks) {
        if (block.title) {
            const level = block.level || 2;
            result += '#'.repeat(Math.min(level, 4)) + ' ' + block.title + '\n\n';
        }
        if (block.content) {
            result += block.content + '\n\n';
        }
    }

    return result.trim();
}

function _extractSectionBlocks(text) {
    const blocks = [];
    let pos = 0;

    while (pos < text.length) {
        // 다음 "title" 키 찾기
        const titleKeyIdx = text.indexOf('"title"', pos);
        if (titleKeyIdx < 0) break;

        // title 값 추출
        const titleValStart = text.indexOf(':', titleKeyIdx + 7);
        if (titleValStart < 0) break;

        const titleVal = _extractJsonStringAt(text, titleValStart + 1);
        if (!titleVal) {
            // title 값이 아직 완성되지 않음
            pos = titleKeyIdx + 7;
            continue;
        }

        // level 추출
        let level = 2;
        const afterTitle = text.substring(titleVal.endIdx);
        const levelMatch = afterTitle.match(/"level"\s*:\s*(\d+)/);
        if (levelMatch) level = parseInt(levelMatch[1]);

        // content 추출
        const contentKeyIdx = text.indexOf('"content"', titleVal.endIdx);
        let content = '';
        let nextPos = titleVal.endIdx + 50;

        const nextTitleIdx = text.indexOf('"title"', titleVal.endIdx + 1);
        if (contentKeyIdx >= 0 && (nextTitleIdx < 0 || contentKeyIdx < nextTitleIdx)) {
            const contentValStart = text.indexOf(':', contentKeyIdx + 9);
            if (contentValStart >= 0) {
                const contentVal = _extractJsonStringAt(text, contentValStart + 1);
                if (contentVal) {
                    content = contentVal.value;
                    nextPos = contentVal.endIdx;
                } else {
                    // content가 아직 스트리밍 중 - 부분 추출
                    content = _extractPartialJsonString(text, contentValStart + 1);
                    nextPos = text.length;
                }
            }
        }

        blocks.push({
            title: titleVal.value,
            level: level,
            content: content,
        });

        pos = nextPos;
    }

    return blocks;
}

function _extractJsonStringAt(text, fromIdx) {
    // fromIdx 이후 첫 번째 따옴표 찾기
    let i = fromIdx;
    while (i < text.length && text[i] !== '"') i++;
    if (i >= text.length) return null;

    i++; // 여는 따옴표 건너뛰기
    let value = '';
    while (i < text.length) {
        if (text[i] === '\\' && i + 1 < text.length) {
            value += text[i] + text[i + 1];
            i += 2;
        } else if (text[i] === '"') {
            return { value: _unescJsonStr(value), endIdx: i + 1 };
        } else {
            value += text[i];
            i++;
        }
    }
    return null; // 닫는 따옴표 없음 (아직 스트리밍 중)
}

function _extractPartialJsonString(text, fromIdx) {
    let i = fromIdx;
    while (i < text.length && text[i] !== '"') i++;
    if (i >= text.length) return '';

    i++; // 여는 따옴표 건너뛰기
    let value = '';
    while (i < text.length) {
        if (text[i] === '\\' && i + 1 < text.length) {
            value += text[i] + text[i + 1];
            i += 2;
        } else if (text[i] === '"') {
            return _unescJsonStr(value);
        } else {
            value += text[i];
            i++;
        }
    }
    // 닫는 따옴표 없이 끝남 - 부분 콘텐츠 반환
    return _unescJsonStr(value);
}

function _unescJsonStr(s) {
    return s.replace(/\\n/g, '\n')
            .replace(/\\t/g, '\t')
            .replace(/\\"/g, '"')
            .replace(/\\\\/g, '\\');
}

function _streamingMarkdownToHtml(md) {
    if (!md) return '';
    _pendingChartConfigs = []; // 대기 차트 초기화
    let html = '';
    const lines = md.split('\n');
    let inList = false;
    let inOrderedList = false;
    let inTable = false;
    let tableLines = [];

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const trimmed = line.trim();

        // 테이블 처리
        if (trimmed.startsWith('|') && trimmed.includes('|', 1)) {
            if (!inTable) {
                inTable = true;
                tableLines = [];
                if (inList) { html += '</ul>'; inList = false; }
                if (inOrderedList) { html += '</ol>'; inOrderedList = false; }
            }
            tableLines.push(trimmed);
            continue;
        } else if (inTable) {
            html += _renderStyledTable(tableLines);
            inTable = false;
            tableLines = [];
        }

        // 빈 줄
        if (!trimmed) {
            if (inList) { html += '</ul>'; inList = false; }
            if (inOrderedList) { html += '</ol>'; inOrderedList = false; }
            continue;
        }

        // 차트 블록 (@@CHARTBLOCK@@ 마커)
        if (trimmed === '@@CHARTBLOCK@@') {
            if (inList) { html += '</ul>'; inList = false; }
            if (inOrderedList) { html += '</ol>'; inOrderedList = false; }

            // JSON 수집 (중괄호 매칭)
            let jsonLines = [];
            let braceCount = 0;
            let started = false;
            let j = i + 1;

            while (j < lines.length) {
                const nextLine = lines[j].trim();
                if (!nextLine && !started) { j++; continue; }
                if (nextLine.includes('{') && !started) started = true;
                if (started) {
                    jsonLines.push(lines[j]);
                    braceCount += (nextLine.match(/{/g) || []).length;
                    braceCount -= (nextLine.match(/}/g) || []).length;
                    if (braceCount <= 0) { j++; break; }
                }
                j++;
            }

            const jsonStr = jsonLines.join('\n').trim();

            if (braceCount <= 0 && jsonStr) {
                // 완성된 JSON → 캐시 확인 후 이미지 또는 대기
                if (_wordChartCache.has(jsonStr)) {
                    const safeJson2 = jsonStr.replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
                    html += `<div class="docx-stream-chart"><img src="${_wordChartCache.get(jsonStr)}" data-chart-json="${safeJson2}" /></div>`;
                } else {
                    const idx = _pendingChartConfigs.length;
                    _pendingChartConfigs.push(jsonStr);
                    html += `<div class="docx-stream-chart-pending" data-chart-idx="${idx}"><div class="docx-chart-loading"><div class="dot-loading"><span></span><span></span><span></span></div><p>차트를 생성하고 있습니다...</p></div></div>`;
                }
            } else if (started) {
                // 아직 스트리밍 중인 JSON → 로딩 표시
                html += `<div class="docx-chart-loading"><div class="dot-loading"><span></span><span></span><span></span></div><p>차트를 생성하고 있습니다...</p></div>`;
            }

            i = j - 1;
            continue;
        }

        // 제목 (# ## ### ####)
        const headingMatch = trimmed.match(/^(#{1,4})\s+(.+)$/);
        if (headingMatch) {
            if (inList) { html += '</ul>'; inList = false; }
            if (inOrderedList) { html += '</ol>'; inOrderedList = false; }
            const level = headingMatch[1].length;
            html += `<h${level}>${_inlineFormatStyled(headingMatch[2])}</h${level}>`;
            continue;
        }

        // 구분선
        if (/^[-*_]{3,}\s*$/.test(trimmed)) {
            if (inList) { html += '</ul>'; inList = false; }
            if (inOrderedList) { html += '</ol>'; inOrderedList = false; }
            html += '<hr/>';
            continue;
        }

        // 불릿 리스트
        if (trimmed.startsWith('- ') || trimmed.startsWith('* ')) {
            if (inOrderedList) { html += '</ol>'; inOrderedList = false; }
            if (!inList) { html += '<ul>'; inList = true; }
            html += `<li>${_inlineFormatStyled(trimmed.substring(2))}</li>`;
            continue;
        }

        // 번호 리스트
        const numMatch = trimmed.match(/^\d+\.\s+(.+)$/);
        if (numMatch) {
            if (inList) { html += '</ul>'; inList = false; }
            if (!inOrderedList) { html += '<ol>'; inOrderedList = true; }
            html += `<li>${_inlineFormatStyled(numMatch[1])}</li>`;
            continue;
        }

        // 인용문
        if (trimmed.startsWith('> ')) {
            if (inList) { html += '</ul>'; inList = false; }
            if (inOrderedList) { html += '</ol>'; inOrderedList = false; }
            html += `<blockquote><p>${_inlineFormatStyled(trimmed.substring(2))}</p></blockquote>`;
            continue;
        }

        // 일반 텍스트
        if (inList) { html += '</ul>'; inList = false; }
        if (inOrderedList) { html += '</ol>'; inOrderedList = false; }
        html += `<p>${_inlineFormatStyled(trimmed)}</p>`;
    }

    if (inList) html += '</ul>';
    if (inOrderedList) html += '</ol>';
    if (inTable) html += _renderStyledTable(tableLines);

    return html;
}

function _inlineFormatStyled(text) {
    text = text.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
    text = text.replace(/\*(.+?)\*/g, '<em>$1</em>');
    text = text.replace(/~~(.+?)~~/g, '<s>$1</s>');
    text = text.replace(/`(.+?)`/g, '<code style="background:#f1f5f9;padding:1px 5px;border-radius:3px;font-size:0.9em;color:#6366f1;">$1</code>');
    return text;
}

function _renderStyledTable(lines) {
    if (lines.length < 2) return '';
    const dataLines = lines.filter(line => {
        const cells = line.replace(/^\||\|$/g, '').split('|').map(c => c.trim());
        return !cells.every(c => /^[-:]+$/.test(c));
    });
    if (dataLines.length === 0) return '';

    let html = '<table>';
    html += '<thead><tr>';
    const headerCells = dataLines[0].replace(/^\||\|$/g, '').split('|').map(c => c.trim());
    headerCells.forEach(cell => {
        html += `<th>${_inlineFormatStyled(cell)}</th>`;
    });
    html += '</tr></thead><tbody>';
    for (let i = 1; i < dataLines.length; i++) {
        const cells = dataLines[i].replace(/^\||\|$/g, '').split('|').map(c => c.trim());
        html += '<tr>';
        cells.forEach(cell => {
            html += `<td>${_inlineFormatStyled(cell)}</td>`;
        });
        html += '</tr>';
    }
    html += '</tbody></table>';
    return html;
}


async function initWordWorkspace() {
    try {
        // CKEditor 5 스크립트 동적 로딩
        await loadCKEditorScript();

        // 기존 CKEditor 인스턴스 정리
        if (_ckEditorInstance) {
            await _ckEditorInstance.destroy();
            _ckEditorInstance = null;
        }

        // 컨테이너에 편집 영역 생성
        const container = document.getElementById('ckeditorContainer');
        if (!container) return;
        container.innerHTML = '<div id="ckeditorEditor"></div>';

        // CKEditor 5 UMD 모듈에서 필요한 클래스 추출
        const {
            ClassicEditor, Essentials,
            Bold, Italic, Underline, Strikethrough,
            Font, Alignment, List, Link,
            Table, TableToolbar, TableProperties, TableCellProperties,
            Indent, IndentBlock, BlockQuote, Heading,
            Undo, SourceEditing, FindAndReplace,
            WordCount, RemoveFormat, Paragraph,
            GeneralHtmlSupport
        } = CKEDITOR;

        // CKEditor 5 초기화
        _ckEditorInstance = await ClassicEditor.create(
            document.getElementById('ckeditorEditor'),
            {
                plugins: [
                    Essentials, Bold, Italic, Underline, Strikethrough,
                    Font, Alignment, List, Link,
                    Table, TableToolbar, TableProperties, TableCellProperties,
                    Indent, IndentBlock, BlockQuote, Heading,
                    Undo, SourceEditing, FindAndReplace,
                    WordCount, RemoveFormat, Paragraph,
                    GeneralHtmlSupport
                ],
                toolbar: {
                    items: [
                        'undo', 'redo', '|',
                        'heading', '|',
                        'bold', 'italic', 'underline', 'strikethrough', '|',
                        'fontColor', 'fontBackgroundColor', '|',
                        'alignment', '|',
                        'bulletedList', 'numberedList', '|',
                        'outdent', 'indent', '|',
                        'insertTable', 'link', '|',
                        'removeFormat', 'sourceEditing', 'findAndReplace'
                    ],
                    shouldNotGroupWhenFull: false
                },
                htmlSupport: {
                    allow: [
                        { name: 'img', attributes: true, styles: true, classes: true },
                        { name: 'div', attributes: true, styles: true, classes: true },
                    ]
                },
                table: {
                    contentToolbar: [
                        'tableColumn', 'tableRow', 'mergeTableCells',
                        'tableProperties', 'tableCellProperties'
                    ]
                },
                heading: {
                    options: [
                        { model: 'paragraph', title: 'Paragraph', class: 'ck-heading_paragraph' },
                        { model: 'heading1', view: 'h1', title: 'Heading 1', class: 'ck-heading_heading1' },
                        { model: 'heading2', view: 'h2', title: 'Heading 2', class: 'ck-heading_heading2' },
                        { model: 'heading3', view: 'h3', title: 'Heading 3', class: 'ck-heading_heading3' },
                        { model: 'heading4', view: 'h4', title: 'Heading 4', class: 'ck-heading_heading4' },
                        { model: 'heading5', view: 'h5', title: 'Heading 5', class: 'ck-heading_heading5' },
                        { model: 'heading6', view: 'h6', title: 'Heading 6', class: 'ck-heading_heading6' }
                    ]
                },
                language: 'ko',
                licenseKey: 'GPL'
            }
        );

        // 기존 생성 데이터가 있으면 로드
        if (state.generatedDocx && state.generatedDocx.sections) {
            // 차트 렌더링을 위해 ECharts 로드
            if (typeof echarts === 'undefined') {
                try { await loadEChartsScript(); } catch(e) {}
            }
            const html = _docxSectionsToHtml(state.generatedDocx);
            _ckEditorInstance.setData(html);
            $('#btnDownloadDocx').show();
            $('#btnShareWord').show();
            $('#btnRewriteWord').css('display', 'flex');
            if (state.generatedDocx.meta && state.generatedDocx.meta.title) {
                $('#wordTitle').text(state.generatedDocx.meta.title);
            }
            $('#instructionsInput').attr('placeholder', t('wordModifyPlaceholder'));
        } else {
            $('#btnDownloadDocx').hide();
            $('#btnShareWord').hide();
            $('#btnRewriteWord').hide();
            $('#wordTitle').text(t('typeWord'));
        }
    } catch (e) {
        console.error('[Word] CKEditor 5 초기화 실패:', e);
        $('#ckeditorContainer').html('<div style="padding:40px;text-align:center;color:#999;">에디터를 불러오지 못했습니다.</div>');
    }
}

function _docxSectionsToHtml(docxData) {
    let html = '';
    const meta = docxData.meta || {};

    // 문서 제목
    if (meta.title) {
        html += `<h1>${_escapeHtmlWord(meta.title)}</h1>`;
    }
    // 문서 설명
    if (meta.description) {
        html += `<p>${_escapeHtmlWord(meta.description)}</p><hr/>`;
    }

    // 섹션 처리
    const sections = docxData.sections || [];
    sections.forEach(section => {
        const level = Math.min(Math.max(section.level || 1, 1), 6);
        if (section.title) {
            html += `<h${level}>${_escapeHtmlWord(section.title)}</h${level}>`;
        }
        if (section.content) {
            html += _markdownToHtml(section.content);
        }
    });

    return html;
}

function _escapeHtmlWord(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function _markdownToHtml(md) {
    let html = '';
    const lines = md.split('\n');
    let inList = false;
    let inOrderedList = false;
    let inTable = false;
    let tableLines = [];

    for (let i = 0; i < lines.length; i++) {
        const line = lines[i];
        const trimmed = line.trim();

        // 차트 코드블록 처리 (```chart ... ```)
        if (trimmed === '```chart') {
            if (inList) { html += '</ul>'; inList = false; }
            if (inOrderedList) { html += '</ol>'; inOrderedList = false; }
            if (inTable) { html += _renderTableHtml(tableLines); inTable = false; tableLines = []; }

            let jsonLines = [];
            let braceCount = 0;
            let started = false;
            let j = i + 1;
            while (j < lines.length) {
                const nextLine = lines[j].trim();
                if (nextLine === '```') { j++; break; }
                if (!nextLine && !started) { j++; continue; }
                if (nextLine.includes('{') && !started) started = true;
                if (started) {
                    jsonLines.push(lines[j]);
                    braceCount += (nextLine.match(/{/g) || []).length;
                    braceCount -= (nextLine.match(/}/g) || []).length;
                    if (braceCount <= 0) { j++; break; }
                }
                j++;
            }
            // 닫는 ``` 건너뛰기
            if (j < lines.length && lines[j].trim() === '```') j++;

            const jsonStr = jsonLines.join('\n').trim();
            if (jsonStr) {
                try {
                    const config = JSON.parse(jsonStr);
                    const dataUrl = _renderChartToImage(config);
                    if (dataUrl) {
                        const safeJson3 = jsonStr.replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
                        html += `<div class="docx-stream-chart"><img src="${dataUrl}" style="max-width:100%;" data-chart-json="${safeJson3}" /></div>`;
                    }
                } catch (e) {
                    // JSON 파싱 실패 시 무시
                }
            }
            i = j - 1;
            continue;
        }

        // 일반 코드블록 스킵 (```xxx ... ```)
        if (trimmed.startsWith('```')) {
            if (inList) { html += '</ul>'; inList = false; }
            if (inOrderedList) { html += '</ol>'; inOrderedList = false; }
            let j = i + 1;
            while (j < lines.length && lines[j].trim() !== '```') j++;
            if (j < lines.length) j++; // 닫는 ``` 건너뛰기
            i = j - 1;
            continue;
        }

        // 테이블 처리
        if (trimmed.startsWith('|') && trimmed.includes('|', 1)) {
            if (!inTable) {
                inTable = true;
                tableLines = [];
                if (inList) { html += '</ul>'; inList = false; }
                if (inOrderedList) { html += '</ol>'; inOrderedList = false; }
            }
            tableLines.push(trimmed);
            continue;
        } else if (inTable) {
            html += _renderTableHtml(tableLines);
            inTable = false;
            tableLines = [];
        }

        // 빈 줄
        if (!trimmed) {
            if (inList) { html += '</ul>'; inList = false; }
            if (inOrderedList) { html += '</ol>'; inOrderedList = false; }
            continue;
        }

        // 마크다운 헤딩 (### 제목)
        const headingMatch = trimmed.match(/^(#{1,6})\s+(.+)$/);
        if (headingMatch) {
            if (inList) { html += '</ul>'; inList = false; }
            if (inOrderedList) { html += '</ol>'; inOrderedList = false; }
            const level = headingMatch[1].length;
            html += `<h${level}>${_inlineFormat(headingMatch[2])}</h${level}>`;
            continue;
        }

        // 구분선 (---, ***, ___)
        if (/^[-*_]{3,}\s*$/.test(trimmed)) {
            if (inList) { html += '</ul>'; inList = false; }
            if (inOrderedList) { html += '</ol>'; inOrderedList = false; }
            html += '<hr/>';
            continue;
        }

        // 불릿 리스트
        if (trimmed.startsWith('- ') || trimmed.startsWith('* ')) {
            if (inOrderedList) { html += '</ol>'; inOrderedList = false; }
            if (!inList) { html += '<ul>'; inList = true; }
            html += `<li>${_inlineFormat(trimmed.substring(2))}</li>`;
            continue;
        }

        // 번호 리스트
        const numMatch = trimmed.match(/^\d+\.\s+(.+)$/);
        if (numMatch) {
            if (inList) { html += '</ul>'; inList = false; }
            if (!inOrderedList) { html += '<ol>'; inOrderedList = true; }
            html += `<li>${_inlineFormat(numMatch[1])}</li>`;
            continue;
        }

        // 인용문
        if (trimmed.startsWith('> ')) {
            if (inList) { html += '</ul>'; inList = false; }
            if (inOrderedList) { html += '</ol>'; inOrderedList = false; }
            html += `<blockquote>${_inlineFormat(trimmed.substring(2))}</blockquote>`;
            continue;
        }

        // 일반 텍스트
        if (inList) { html += '</ul>'; inList = false; }
        if (inOrderedList) { html += '</ol>'; inOrderedList = false; }
        html += `<p>${_inlineFormat(trimmed)}</p>`;
    }

    if (inList) html += '</ul>';
    if (inOrderedList) html += '</ol>';
    if (inTable) html += _renderTableHtml(tableLines);

    return html;
}

function _inlineFormat(text) {
    // Bold: **text**
    text = text.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
    // Italic: *text*
    text = text.replace(/\*(.+?)\*/g, '<em>$1</em>');
    // Strikethrough: ~~text~~
    text = text.replace(/~~(.+?)~~/g, '<s>$1</s>');
    return text;
}

function _renderTableHtml(lines) {
    if (lines.length < 2) return '';
    // Filter out separator lines (---|---)
    const dataLines = lines.filter(line => {
        const cells = line.replace(/^\||\|$/g, '').split('|').map(c => c.trim());
        return !cells.every(c => /^[-:]+$/.test(c));
    });
    if (dataLines.length === 0) return '';

    let html = '<figure class="table"><table>';
    html += '<thead><tr>';
    const headerCells = dataLines[0].replace(/^\||\|$/g, '').split('|').map(c => c.trim());
    headerCells.forEach(cell => {
        html += `<th>${_inlineFormat(cell)}</th>`;
    });
    html += '</tr></thead><tbody>';
    for (let i = 1; i < dataLines.length; i++) {
        const cells = dataLines[i].replace(/^\||\|$/g, '').split('|').map(c => c.trim());
        html += '<tr>';
        cells.forEach(cell => {
            html += `<td>${_inlineFormat(cell)}</td>`;
        });
        html += '</tr>';
    }
    html += '</tbody></table></figure>';
    return html;
}

function _renderChartToImage(config) {
    if (typeof echarts === 'undefined') return null;
    try {
        const option = _buildEChartsOptionFromChartJson(config);
        const offDiv = document.createElement('div');
        offDiv.style.cssText = 'width:900px;height:450px;position:absolute;left:-9999px;top:-9999px;';
        document.body.appendChild(offDiv);
        const chart = echarts.init(offDiv);
        chart.setOption(option);
        const dataUrl = chart.getDataURL({ type: 'png', pixelRatio: 2, backgroundColor: '#fff' });
        chart.dispose();
        document.body.removeChild(offDiv);
        return dataUrl;
    } catch (e) {
        console.error('[Chart Render] error:', e);
        return null;
    }
}

function _htmlToDocxSections(html) {
    // Parse TinyMCE HTML back to sections format for saving
    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');
    const body = doc.body;

    const sections = [];
    let currentSection = null;
    const meta = {};

    const children = Array.from(body.children);

    for (const el of children) {
        const tag = el.tagName.toLowerCase();

        // h1-h6 = section headers
        if (/^h[1-6]$/.test(tag)) {
            const level = parseInt(tag.charAt(1));
            // First h1 might be the document title
            if (level === 1 && sections.length === 0 && !meta.title) {
                meta.title = el.textContent.trim();
                continue;
            }
            // Save previous section
            if (currentSection) {
                sections.push(currentSection);
            }
            currentSection = {
                title: el.textContent.trim(),
                level: level,
                content: '',
            };
        } else {
            // Content elements
            if (!currentSection) {
                // Content before first heading - check if it's description
                if (tag === 'p' && sections.length === 0 && !meta.description) {
                    const text = el.textContent.trim();
                    if (text && el.querySelector('hr') === null) {
                        meta.description = text;
                        continue;
                    }
                }
                if (tag === 'hr') continue;
                currentSection = { title: '', level: 1, content: '' };
            }
            currentSection.content += _elementToMarkdown(el) + '\n';
        }
    }

    if (currentSection) {
        sections.push(currentSection);
    }

    return { sections, meta };
}

function _elementToMarkdown(el) {
    const tag = el.tagName.toLowerCase();

    // 차트 이미지 복원 (data-chart-json 속성이 있는 img)
    const chartImg = el.querySelector('img[data-chart-json]');
    if (chartImg) {
        const jsonStr = chartImg.getAttribute('data-chart-json');
        if (jsonStr) {
            try {
                const parsed = JSON.parse(jsonStr);
                return '```chart\n' + JSON.stringify(parsed, null, 2) + '\n```';
            } catch (e) {}
        }
    }

    // CKEditor 5는 테이블을 <figure class="table">로 감쌈
    if (tag === 'figure') {
        const innerTable = el.querySelector('table');
        if (innerTable) return _elementToMarkdown(innerTable);
        return el.textContent.trim();
    }

    if (tag === 'p') {
        // p 태그 내부에 차트 이미지가 있으면 처리
        const pChartImg = el.querySelector('img[data-chart-json]');
        if (pChartImg) {
            const jsonStr = pChartImg.getAttribute('data-chart-json');
            if (jsonStr) {
                try {
                    const parsed = JSON.parse(jsonStr);
                    return '```chart\n' + JSON.stringify(parsed, null, 2) + '\n```';
                } catch (e) {}
            }
        }
        return el.innerHTML
            .replace(/<strong>(.*?)<\/strong>/g, '**$1**')
            .replace(/<b>(.*?)<\/b>/g, '**$1**')
            .replace(/<em>(.*?)<\/em>/g, '*$1*')
            .replace(/<i>(.*?)<\/i>/g, '*$1*')
            .replace(/<s>(.*?)<\/s>/g, '~~$1~~')
            .replace(/<[^>]+>/g, '')
            .trim();
    }

    if (tag === 'ul') {
        return Array.from(el.querySelectorAll(':scope > li')).map(li => `- ${li.textContent.trim()}`).join('\n');
    }
    if (tag === 'ol') {
        return Array.from(el.querySelectorAll(':scope > li')).map((li, i) => `${i + 1}. ${li.textContent.trim()}`).join('\n');
    }
    if (tag === 'blockquote') {
        return `> ${el.textContent.trim()}`;
    }
    if (tag === 'table') {
        const rows = Array.from(el.querySelectorAll('tr'));
        if (rows.length === 0) return '';
        let md = '';
        rows.forEach((row, idx) => {
            const cells = Array.from(row.querySelectorAll('th, td')).map(c => c.textContent.trim());
            md += '| ' + cells.join(' | ') + ' |\n';
            if (idx === 0) {
                md += '| ' + cells.map(() => '---').join(' | ') + ' |\n';
            }
        });
        return md.trim();
    }
    if (tag === 'hr') return '';

    return el.textContent.trim();
}

// ============ DOCX 템플릿 관리 ============

async function loadDocxTemplate() {
    if (!state.currentProject) return;
    try {
        const res = await apiGet(`/api/resources/docx-template/${state.currentProject._id}`);
        if (res.template) {
            state.docxTemplateId = res.template._id;
            _updateDocxTemplateButton(res.template.original_filename);
        } else {
            state.docxTemplateId = null;
            _updateDocxTemplateButton(null);
        }
    } catch(e) {
        state.docxTemplateId = null;
        _updateDocxTemplateButton(null);
    }
}

function _updateDocxTemplateButton(filename) {
    const $btn = $('#btnDocxTemplate');
    if (filename) {
        // 파일명 15자 제한
        const short = filename.length > 15 ? filename.substring(0, 12) + '...' : filename;
        $btn.addClass('has-template');
        $('#docxTemplateLabel').html(
            `<span style="max-width:100px;overflow:hidden;text-overflow:ellipsis" title="${filename}">${short}</span>` +
            `<span class="template-remove" onclick="event.stopPropagation();removeDocxTemplate()">✕</span>`
        );
    } else {
        $btn.removeClass('has-template');
        $('#docxTemplateLabel').text('양식 업로드');
    }
}

function handleDocxTemplate() {
    if (state.docxTemplateId) {
        // 이미 템플릿이 있으면 교체 확인
        if (confirm('기존 양식을 교체하시겠습니까?')) {
            $('#docxTemplateFileInput').click();
        }
    } else {
        $('#docxTemplateFileInput').click();
    }
}

async function removeDocxTemplate() {
    if (!state.docxTemplateId) return;
    try {
        await apiDelete(`/api/resources/docx-template/${state.docxTemplateId}`);
        state.docxTemplateId = null;
        _updateDocxTemplateButton(null);
        showToast('양식이 제거되었습니다.', 'success');
    } catch(e) {
        showToast('양식 제거 실패', 'error');
    }
}

// 파일 인풋 변경 핸들러 - 초기화 시점에 등록
$(document).on('change', '#docxTemplateFileInput', async function() {
    const file = this.files[0];
    if (!file) return;

    if (!file.name.toLowerCase().endsWith('.docx')) {
        showToast('.docx 파일만 업로드 가능합니다.', 'error');
        this.value = '';
        return;
    }

    const formData = new FormData();
    formData.append('project_id', state.currentProject._id);
    formData.append('file', file);

    try {
        // 기존 템플릿 삭제
        if (state.docxTemplateId) {
            await apiDelete(`/api/resources/docx-template/${state.docxTemplateId}`);
        }

        const res = await apiUpload('/api/resources/docx-template', formData);
        state.docxTemplateId = res.template._id;
        _updateDocxTemplateButton(res.template.original_filename);
        showToast('양식 템플릿이 등록되었습니다.', 'success');
    } catch(e) {
        showToast(e.message || '양식 업로드 실패', 'error');
    }

    this.value = '';
});

async function generateWord() {
    const instructions = $('#instructionsInput').val().trim();
    const lang = $('#langSelect').val();

    if (state.resources.length === 0 && !instructions) {
        showToast(t('msgEnterInstructions'), 'error');
        return;
    }

    _isGenerating = true;
    _wordDeltaCount = 0;
    _wordRequestedTotal = parseInt($('#docxPageCountSelect').val()) || 0; // 요청한 전체 섹션 수
    state.generatedDocx = null;
    $('#btnDownloadDocx').hide();

    _showStopButton();
    _showWordProgress(t('wordPreparing'), '');

    // 스트리밍 프리뷰 표시 (CKEditor 대신)
    _showWordStreamingPreview();

    _abortController = new AbortController();

    try {
        const response = await fetch(`/${state.jwtToken}/api/generate/docx/stream`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                project_id: state.currentProject._id,
                instructions: instructions,
                lang: lang,
                section_count: $('#docxPageCountSelect').val(),
                docx_template_id: state.docxTemplateId || null,
            }),
            signal: _abortController.signal,
        });

        if (!response.ok) {
            const err = await response.json().catch(() => ({}));
            throw new Error(err.detail || 'Generation failed');
        }

        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;

            buffer += decoder.decode(value, { stream: true });
            const lineArr = buffer.split('\n');
            buffer = lineArr.pop() || '';

            for (const line of lineArr) {
                if (!line.startsWith('data: ')) continue;
                try {
                    const evt = JSON.parse(line.slice(6));
                    _handleWordSSEEvent(evt);
                } catch (e) { /* ignore */ }
            }
        }

        if (buffer.startsWith('data: ')) {
            try {
                const evt = JSON.parse(buffer.slice(6));
                _handleWordSSEEvent(evt);
            } catch (e) {}
        }

    } catch (e) {
        if (e.name !== 'AbortError') {
            console.error('[Word] 생성 실패:', e);
            showToast(e.message || '문서 생성 실패', 'error');
            _hideWordStreamingPreview();
        } else {
            // 중단 시: 현재 프리뷰 그대로 유지, 커서만 제거 (재파싱 금지)
            if (_wordStreamRenderTimer) {
                clearTimeout(_wordStreamRenderTimer);
                _wordStreamRenderTimer = null;
            }
            const cursor = document.querySelector('#wordStreamingContent .word-typing-cursor');
            if (cursor) cursor.remove();
            showToast(t('msgStopped'), 'info');
        }
    } finally {
        _isGenerating = false;
        _abortController = null;
        _hideWordProgress();
        _showGenerateOrRestartButton();
    }
}

async function modifyWord() {
    const instructions = $('#instructionsInput').val().trim();
    if (!instructions) {
        showToast(t('msgEnterInstructions'), 'error');
        return;
    }

    const lang = $('#langSelect').val();

    // CKEditor에서 현재 HTML 가져와서 sections로 변환
    let currentData = state.generatedDocx || {};
    if (_ckEditorInstance) {
        const parsed = _htmlToDocxSections(_ckEditorInstance.getData());
        currentData = {
            sections: parsed.sections,
            meta: parsed.meta || currentData.meta || {},
        };
    }

    _isGenerating = true;
    _showStopButton();
    _showWordProgress(t('wordStreaming'), '');

    // 스트리밍 프리뷰 표시
    _showWordStreamingPreview();

    _abortController = new AbortController();

    try {
        const response = await fetch(`/${state.jwtToken}/api/generate/docx/modify/stream`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                project_id: state.currentProject._id,
                instruction: instructions,
                current_data: currentData,
                lang: lang,
            }),
            signal: _abortController.signal,
        });

        if (!response.ok) {
            const err = await response.json().catch(() => ({}));
            throw new Error(err.detail || 'Modification failed');
        }

        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;

            buffer += decoder.decode(value, { stream: true });
            const lineArr = buffer.split('\n');
            buffer = lineArr.pop() || '';

            for (const line of lineArr) {
                if (!line.startsWith('data: ')) continue;
                try {
                    const evt = JSON.parse(line.slice(6));
                    _handleWordSSEEvent(evt);
                } catch (e) { /* ignore */ }
            }
        }

        if (buffer.startsWith('data: ')) {
            try {
                const evt = JSON.parse(buffer.slice(6));
                _handleWordSSEEvent(evt);
            } catch (e) {}
        }

    } catch (e) {
        if (e.name !== 'AbortError') {
            console.error('[Word] 수정 실패:', e);
            showToast(e.message || '문서 수정 실패', 'error');
            _hideWordStreamingPreview();
        } else {
            // 중단 시: 현재 프리뷰 그대로 유지, 커서만 제거
            if (_wordStreamRenderTimer) {
                clearTimeout(_wordStreamRenderTimer);
                _wordStreamRenderTimer = null;
            }
            const cursor = document.querySelector('#wordStreamingContent .word-typing-cursor');
            if (cursor) cursor.remove();
            showToast(t('msgStopped'), 'info');
        }
    } finally {
        _isGenerating = false;
        _abortController = null;
        _hideWordProgress();
        _showGenerateOrRestartButton();
    }
}

function _handleWordSSEEvent(evt) {
    const eventType = evt.event;

    switch (eventType) {
        case 'start':
            _showWordProgress(t('wordPreparing'), '');
            break;

        case 'searching':
            _showWordProgress(t('wordSearching'), '');
            break;

        case 'search_done':
            _showWordProgress(t('wordSearchDone'), '');
            break;

        case 'delta':
            // 실시간 스트리밍 텍스트 표시
            if (evt.text) {
                _appendWordStreamDelta(evt.text);
            }
            // 섹션 진행률 표시 (10 delta마다 체크)
            _wordDeltaCount = (_wordDeltaCount || 0) + 1;
            if (_wordDeltaCount % 10 === 0) {
                const wSecIdx = _wordStreamBuffer.indexOf('"sections"');
                if (wSecIdx >= 0) {
                    const wSecText = _wordStreamBuffer.substring(wSecIdx);
                    const wContentCount = (wSecText.match(/"content"\s*:/g) || []).length;
                    const wTotalFromBuffer = (wSecText.match(/"title"\s*:/g) || []).length;
                    // 요청한 전체 섹션 수 우선, 없으면 버퍼에서 추정
                    const wTotal = _wordRequestedTotal > 0 ? Math.max(_wordRequestedTotal, wContentCount) : wTotalFromBuffer;
                    if (wTotal > 0) {
                        _showWordProgress(`섹션 작성 중 (${wContentCount}/${wTotal})`, '');
                    }
                }
            }
            break;

        case 'parsing':
            // 타이머만 정리, 프리뷰 재파싱 하지 않음 (내용 깨짐 방지)
            if (_wordStreamRenderTimer) {
                clearTimeout(_wordStreamRenderTimer);
                _wordStreamRenderTimer = null;
            }
            _showWordProgress(t('wordFinalizing'), '');
            break;

        case 'docx_data':
            _hideWordProgress();
            state.generatedDocx = evt.docx;
            state.currentProject.status = 'generated';
            $('#wsProjectStatus').text(t('statusGenerated')).attr('class', 'ws-status generated');

            // 스트리밍 프리뷰 HTML을 그대로 CKEditor에 이식 (재변환 금지)
            if (_wordStreamRenderTimer) {
                clearTimeout(_wordStreamRenderTimer);
                _wordStreamRenderTimer = null;
            }
            _renderWordStreamingContent(); // 마지막 버퍼 반영
            {
                const streamHtml = $('#wordStreamingContent').html() || '';
                // 커서 제거
                const cleanHtml = streamHtml.replace(/<span class="word-typing-cursor"><\/span>/g, '');
                $('#wordStreamingPreview').hide();
                $('#ckeditorContainer').show();
                if (_ckEditorInstance && cleanHtml) {
                    _ckEditorInstance.setData(cleanHtml);
                }
            }

            $('#btnDownloadDocx').show();
            $('#btnShareWord').show();
            $('#btnRewriteWord').css('display', 'flex');
            if (state.generatedDocx.meta && state.generatedDocx.meta.title) {
                $('#wordTitle').text(state.generatedDocx.meta.title);
            }
            $('#instructionsInput').val('').attr('placeholder', t('wordModifyPlaceholder'));
            autoResizeTextarea(document.getElementById('instructionsInput'));
            break;

        case 'complete':
            _hideWordProgress();
            // 프리뷰가 아직 보이면 그대로 CKEditor로 이식
            if ($('#wordStreamingPreview').is(':visible')) {
                if (_wordStreamRenderTimer) {
                    clearTimeout(_wordStreamRenderTimer);
                    _wordStreamRenderTimer = null;
                }
                _renderWordStreamingContent();
                const streamHtml = ($('#wordStreamingContent').html() || '').replace(/<span class="word-typing-cursor"><\/span>/g, '');
                $('#wordStreamingPreview').hide();
                $('#ckeditorContainer').show();
                if (_ckEditorInstance && streamHtml) {
                    _ckEditorInstance.setData(streamHtml);
                }
            }
            showToast(t('msgWordGenerated'), 'success');
            break;

        case 'stopped':
            _hideWordProgress();
            // 중단 시: 현재 프리뷰 그대로 유지, 커서만 제거 (재파싱 금지)
            if (_wordStreamRenderTimer) {
                clearTimeout(_wordStreamRenderTimer);
                _wordStreamRenderTimer = null;
            }
            const stoppedCursor = document.querySelector('#wordStreamingContent .word-typing-cursor');
            if (stoppedCursor) stoppedCursor.remove();
            state.currentProject.status = 'stopped';
            $('#wsProjectStatus').text(t('statusStopped')).attr('class', 'ws-status stopped');
            showToast(t('msgStopped'), 'info');
            break;

        case 'error':
            _hideWordProgress();
            _hideWordStreamingPreview();
            showToast(evt.message || '생성 실패', 'error');
            break;
    }
}

function downloadDOCX() {
    if (!state.currentProject || !state.generatedDocx) {
        showToast(t('msgNoWordData'), 'error');
        return;
    }

    // 먼저 CKEditor에서 편집된 내용을 저장
    if (_ckEditorInstance) {
        const parsed = _htmlToDocxSections(_ckEditorInstance.getData());
        _saveAndDownloadDocx(parsed);
    } else {
        const url = `/${state.jwtToken}/api/generate/${state.currentProject._id}/download/docx`;
        _downloadFile(url, (state.currentProject.name || 'document') + '.docx');
    }
}

async function _saveAndDownloadDocx(parsed) {
    try {
        await fetch(`/${state.jwtToken}/api/generate/${state.currentProject._id}/docx`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                sections: parsed.sections,
                meta: parsed.meta || state.generatedDocx.meta || {},
            }),
        });
        const url = `/${state.jwtToken}/api/generate/${state.currentProject._id}/download/docx`;
        await _downloadFile(url, (state.currentProject.name || 'document') + '.docx');
    } catch (e) {
        console.error('[Word] 저장 후 다운로드 실패:', e);
        showToast('다운로드 실패', 'error');
    }
}

// ============ OnlyOffice 통합 ============

function _getProjectTypeBadge(projectType) {
    const badges = {
        'slide': '<span class="rc-type-badge ppt">PPT</span>',
        'excel': '<span class="rc-type-badge excel">Excel</span>',
        'onlyoffice_pptx': '<span class="rc-type-badge oo-pptx">OO PPT</span>',
        'onlyoffice_xlsx': '<span class="rc-type-badge oo-xlsx">OO Excel</span>',
        'onlyoffice_docx': '<span class="rc-type-badge oo-docx">OO Word</span>',
        'word': '<span class="rc-type-badge word">Word</span>',
    };
    return badges[projectType] || '<span class="rc-type-badge ppt">PPT</span>';
}

function _getProjectTypeClass(projectType) {
    const map = {
        'slide': 'ptype-ppt',
        'onlyoffice_pptx': 'ptype-ppt',
        'excel': 'ptype-excel',
        'onlyoffice_xlsx': 'ptype-excel',
        'word': 'ptype-word',
        'onlyoffice_docx': 'ptype-word',
    };
    return map[projectType] || 'ptype-ppt';
}

function _getProjectTypeIcon(projectType) {
    const icons = {
        'slide': '<svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="2" y="3" width="12" height="10" rx="2"/><path d="M7 7l3 2-3 2V7z"/></svg>',
        'onlyoffice_pptx': '<svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="2" y="3" width="12" height="10" rx="2"/><path d="M7 7l3 2-3 2V7z"/></svg>',
        'excel': '<svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="2" y="2" width="12" height="12" rx="2"/><line x1="2" y1="6" x2="14" y2="6"/><line x1="2" y1="10" x2="14" y2="10"/><line x1="6" y1="2" x2="6" y2="14"/><line x1="10" y1="2" x2="10" y2="14"/></svg>',
        'onlyoffice_xlsx': '<svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="2" y="2" width="12" height="12" rx="2"/><line x1="2" y1="6" x2="14" y2="6"/><line x1="2" y1="10" x2="14" y2="10"/><line x1="6" y1="2" x2="6" y2="14"/><line x1="10" y1="2" x2="10" y2="14"/></svg>',
        'word': '<svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="3" y="2" width="10" height="12" rx="2"/><path d="M6 5h4M6 8h4M6 11h2"/></svg>',
        'onlyoffice_docx': '<svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="3" y="2" width="10" height="12" rx="2"/><path d="M6 5h4M6 8h4M6 11h2"/></svg>',
    };
    return icons[projectType] || '<svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="1.8"><rect x="2" y="3" width="12" height="10" rx="2"/><path d="M7 7l3 2-3 2V7z"/></svg>';
}

let _onlyofficeScriptLoaded = false;

async function loadOnlyOfficeScript(onlyofficeUrl) {
    if (_onlyofficeScriptLoaded) return;
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = `${onlyofficeUrl}/web-apps/apps/api/documents/api.js`;
        script.onload = () => { _onlyofficeScriptLoaded = true; resolve(); };
        script.onerror = () => reject(new Error('OnlyOffice api.js 로드 실패'));
        document.head.appendChild(script);
    });
}

function destroyOnlyOfficeEditor() {
    if (state.onlyofficeEditor) {
        try { state.onlyofficeEditor.destroyEditor(); } catch (e) { }
        state.onlyofficeEditor = null;
    }
    $('#onlyofficeContainer').empty();
    // 전체화면 해제
    if ($('#appView').hasClass('oo-fullscreen')) {
        $('#appView').removeClass('oo-fullscreen');
        $('.oo-fullscreen-exit').remove();
    }
    $('#btnFullscreenOO').hide();
}

async function openOnlyOfficeEditor(projectId) {
    try {
        const res = await apiGet(`/api/onlyoffice/${projectId}/config`);
        const config = res.config;
        const onlyofficeUrl = res.onlyoffice_url;

        await loadOnlyOfficeScript(onlyofficeUrl);
        destroyOnlyOfficeEditor();

        // 에디터 컨테이너 준비
        $('#onlyofficeContainer').html('<div id="ooEditorDiv" style="width:100%;height:100%;"></div>');

        // 이벤트 핸들러 추가
        config.events = {
            onAppReady: function () {
                console.log('[OnlyOffice] 에디터 준비 완료');
            },
            onDocumentReady: function () {
                console.log('[OnlyOffice] 문서 로드 완료');
                $('#btnDownloadOO').show();
                $('#btnFullscreenOO').show();
            },
            onError: function (event) {
                console.error('[OnlyOffice] 에디터 오류:', event.data);
                showToast('OnlyOffice 에디터 오류가 발생했습니다', 'error');
            },
        };

        state.onlyofficeEditor = new DocsAPI.DocEditor('ooEditorDiv', config);

        // OnlyOffice SDK iframe에 unload 권한 부여 (Permissions Policy 위반 방지)
        setTimeout(function() {
            var ooIframe = document.querySelector('#ooEditorDiv iframe');
            if (ooIframe) {
                ooIframe.setAttribute('allow', 'clipboard-read; clipboard-write; unload');
            }
        }, 500);
    } catch (e) {
        console.error('[OnlyOffice] 에디터 열기 실패:', e);
        showToast('OnlyOffice 에디터를 열 수 없습니다: ' + e.message, 'error');
    }
}

// ============ AI 텍스트 리라이트 (선택 영역 수정) ============

let _rewriteSelectedText = '';
let _rewriteResultText = '';
let _rewriteSource = ''; // 'ckeditor' | 'manual'

function _getDocumentContextText() {
    // CKEditor: 전체 텍스트 추출
    if (_ckEditorInstance && $('#wordWorkspace').is(':visible')) {
        try {
            var root = _ckEditorInstance.model.document.getRoot();
            var text = '';
            for (var child of root.getChildren()) {
                var walker = child.getChildren ? child.getChildren() : [];
                for (var node of walker) {
                    if (node.is('$text') || node.is('$textProxy')) text += node.data;
                }
                text += '\n';
            }
            return text.trim();
        } catch (e) {
            // 폴백: HTML에서 텍스트 추출
            try {
                var html = _ckEditorInstance.getData();
                var tmp = document.createElement('div');
                tmp.innerHTML = html;
                return (tmp.textContent || tmp.innerText || '').trim();
            } catch (e2) {}
        }
    }
    // OnlyOffice: generated_docx 데이터에서 추출
    if (state.generatedDocx && state.generatedDocx.sections) {
        var parts = [];
        if (state.generatedDocx.meta) {
            if (state.generatedDocx.meta.title) parts.push(state.generatedDocx.meta.title);
            if (state.generatedDocx.meta.description) parts.push(state.generatedDocx.meta.description);
        }
        state.generatedDocx.sections.forEach(function(sec) {
            if (sec.title) parts.push(sec.title);
            if (sec.content) parts.push(sec.content);
        });
        return parts.join('\n\n');
    }
    return '';
}

function openRewriteModal() {
    console.log('[Rewrite] openRewriteModal 호출됨');
    try {
        // 모달 초기화
        _rewriteSelectedText = '';
        _rewriteResultText = '';
        _rewriteSource = '';

        _resetRewriteModalUI();

        // 1) CKEditor (일반 워드)
        if (_ckEditorInstance && $('#wordWorkspace').is(':visible')) {
            _rewriteSource = 'ckeditor';
            var sel = _ckEditorInstance.model.document.selection;
            var range = sel.getFirstRange();
            if (range && !range.isCollapsed) {
                var selectedText = '';
                for (var item of range.getItems()) {
                    if (item.is('$textProxy')) selectedText += item.data;
                }
                if (selectedText.trim()) {
                    _rewriteSelectedText = selectedText;
                    $('#rewriteSelectedText').text(selectedText);
                    _showRewriteModalNow();
                    return;
                }
            }
            _showRewriteManualInput('텍스트를 선택한 후 다시 시도하거나, 수정할 텍스트를 직접 입력하세요.');
            _showRewriteModalNow();
            return;
        }

        // 2) 그 외 수동 입력
        _rewriteSource = 'manual';
        _showRewriteManualInput('');
        _showRewriteModalNow();
    } catch (e) {
        console.error('[Rewrite] openRewriteModal 에러:', e);
        showToast('AI 수정 기능 오류: ' + e.message, 'error');
    }
}

function _resetRewriteModalUI() {
    if ($('#rewriteSelectedText').is('textarea')) {
        $('#rewriteSelectedText').replaceWith(
            '<div id="rewriteSelectedText" style="background:var(--surface-2);border:1px solid var(--border);border-radius:8px;padding:10px 12px;font-size:13px;color:var(--text);max-height:120px;overflow-y:auto;white-space:pre-wrap;"></div>'
        );
    }
    $('#rewriteSelectedText').html('<span style="color:#94a3b8;">선택된 텍스트를 가져오는 중...</span>');
    $('#rewriteInstructions').val('');
    $('#rewritePreview').hide();
    $('#rewriteResultText').text('');
    $('#btnRewriteApply').hide();
    $('#btnRewriteSubmit').show().prop('disabled', false).text('AI 수정 시작');
}

function _showRewriteModalNow() {
    $('#rewriteModal').css({'display': 'flex', 'z-index': '99999'});
    setTimeout(function() { $('#rewriteInstructions').focus(); }, 300);
}

function _showRewriteManualInput(msg) {
    var placeholder = msg || '수정할 텍스트를 여기에 붙여넣으세요.';
    $('#rewriteSelectedText').replaceWith(
        '<textarea id="rewriteSelectedText" rows="4" placeholder="' + placeholder + '" ' +
        'style="width:100%;background:var(--surface-2);border:1px solid var(--border);border-radius:8px;padding:10px 12px;font-size:13px;color:var(--text);max-height:120px;resize:vertical;outline:none;box-sizing:border-box;" ' +
        'onfocus="this.style.borderColor=\'#6366f1\'" onblur="this.style.borderColor=\'\'"></textarea>'
    );
}

function closeRewriteModal() {
    $('#rewriteModal').css('display', 'none');
    _rewriteSelectedText = '';
    _rewriteResultText = '';
}

async function submitRewrite() {
    var instructions = $('#rewriteInstructions').val().trim();
    if (!instructions) {
        showToast('수정 지시사항을 입력하세요.', 'warning');
        $('#rewriteInstructions').focus();
        return;
    }

    if (!_rewriteSelectedText && $('#rewriteSelectedText').is('textarea')) {
        _rewriteSelectedText = $('#rewriteSelectedText').val().trim();
    }
    if (!_rewriteSelectedText) {
        showToast('수정할 텍스트를 입력하세요.', 'warning');
        return;
    }

    var lang = $('#langSelect').val() || 'ko';
    $('#btnRewriteSubmit').prop('disabled', true).text('수정 중...');
    $('#btnRewriteApply').hide();
    $('#rewritePreview').show();
    $('#rewriteResultText').text('');
    _rewriteResultText = '';

    // 전체 문서 내용을 문맥으로 전달
    var contextText = _getDocumentContextText();

    try {
        var response = await fetch('/' + state.jwtToken + '/api/generate/rewrite/stream', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                project_id: state.currentProject ? state.currentProject._id : '',
                selected_text: _rewriteSelectedText,
                instructions: instructions,
                lang: lang,
                context_text: contextText,
            }),
        });

        if (!response.ok) {
            var err = await response.json().catch(function() { return {}; });
            throw new Error(err.detail || 'Rewrite failed');
        }

        var reader = response.body.getReader();
        var decoder = new TextDecoder();
        var buffer = '';

        while (true) {
            var chunk = await reader.read();
            if (chunk.done) break;
            buffer += decoder.decode(chunk.value, { stream: true });
            var lines = buffer.split('\n');
            buffer = lines.pop();

            for (var i = 0; i < lines.length; i++) {
                var line = lines[i];
                if (!line.startsWith('data: ')) continue;
                var evt;
                try { evt = JSON.parse(line.substring(6)); } catch(e) { continue; }

                if (evt.event === 'delta') {
                    _rewriteResultText += (evt.text || '');
                    $('#rewriteResultText').text(_rewriteResultText);
                    var el = document.getElementById('rewriteResultText');
                    if (el) el.scrollTop = el.scrollHeight;
                } else if (evt.event === 'done') {
                    _rewriteResultText = evt.text || _rewriteResultText;
                    $('#rewriteResultText').text(_rewriteResultText);
                } else if (evt.event === 'error') {
                    showToast(evt.message || '수정 실패', 'error');
                }
            }
        }

        if (_rewriteResultText.trim()) {
            $('#btnRewriteSubmit').text('다시 수정').prop('disabled', false);
            $('#btnRewriteApply').show();
        } else {
            $('#btnRewriteSubmit').text('AI 수정 시작').prop('disabled', false);
        }
    } catch (e) {
        console.error('[Rewrite] Error:', e);
        showToast('텍스트 수정 중 오류가 발생했습니다: ' + e.message, 'error');
        $('#btnRewriteSubmit').text('AI 수정 시작').prop('disabled', false);
    }
}

function applyRewrite() {
    if (!_rewriteResultText.trim()) {
        showToast('적용할 수정 결과가 없습니다.', 'warning');
        return;
    }

    // CKEditor (일반 워드)
    if (_rewriteSource === 'ckeditor' && _ckEditorInstance) {
        try {
            var editor = _ckEditorInstance;
            var html = _markdownToSimpleHtml(_rewriteResultText);
            var viewFragment = editor.data.processor.toView(html);
            var modelFragment = editor.data.toModel(viewFragment);
            editor.model.change(function(writer) {
                editor.model.insertContent(modelFragment);
            });
            showToast('텍스트가 수정되었습니다.', 'success');
            closeRewriteModal();
        } catch (e) {
            console.error('[Rewrite] CKEditor 적용 실패:', e);
            _copyToClipboard(_rewriteResultText);
            showToast('결과가 클립보드에 복사되었습니다. Ctrl+V로 붙여넣으세요.', 'info');
            closeRewriteModal();
        }
        return;
    }

    // 폴백: 클립보드 복사
    _copyToClipboard(_rewriteResultText);
    showToast('결과가 클립보드에 복사되었습니다. Ctrl+V로 붙여넣으세요.', 'info');
    closeRewriteModal();
}

function _copyToClipboard(text) {
    try {
        navigator.clipboard.writeText(text);
    } catch (e) {
        var ta = document.createElement('textarea');
        ta.value = text;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        document.body.removeChild(ta);
    }
}

function _markdownToSimpleHtml(md) {
    // 간단한 마크다운 → HTML 변환
    let html = md;
    // 빈줄로 문단 구분
    html = html.replace(/\n{2,}/g, '</p><p>');
    // 굵게
    html = html.replace(/\*\*(.+?)\*\*/g, '<b>$1</b>');
    // 기울임
    html = html.replace(/\*(.+?)\*/g, '<i>$1</i>');
    // 불릿 리스트
    html = html.replace(/^[-*]\s+(.+)$/gm, '<li>$1</li>');
    html = html.replace(/(<li>.*<\/li>\n?)+/g, '<ul>$&</ul>');
    // 번호 리스트
    html = html.replace(/^\d+\.\s+(.+)$/gm, '<li>$1</li>');
    // 줄바꿈
    html = html.replace(/\n/g, '<br>');
    // 정리
    html = '<p>' + html + '</p>';
    html = html.replace(/<p><\/p>/g, '');
    return html;
}

async function initOnlyOfficeWorkspace() {
    if (state.onlyofficeDoc) {
        await openOnlyOfficeEditor(state.currentProject._id);
    } else {
        // 문서가 없을 때 빈 상태 표시
        const projectType = state.currentProject ? state.currentProject.project_type : '';
        const icons = {
            'onlyoffice_pptx': '<svg width="48" height="48" fill="none" viewBox="0 0 24 24"><rect x="2" y="3" width="20" height="14" rx="2" stroke="#a78bfa" stroke-width="1.5"/><path d="M8 21h8M12 17v4" stroke="#a78bfa" stroke-width="1.5" stroke-linecap="round"/><path d="M7 8h4M7 11h6" stroke="#c4b5fd" stroke-width="1.5" stroke-linecap="round"/><rect x="14" y="7" width="4" height="5" rx=".5" fill="#ede9fe"/></svg>',
            'onlyoffice_xlsx': '<svg width="48" height="48" fill="none" viewBox="0 0 24 24"><rect x="3" y="3" width="18" height="18" rx="2" stroke="#34d399" stroke-width="1.5"/><path d="M3 9h18M3 15h18M9 3v18M15 3v18" stroke="#6ee7b7" stroke-width="1"/><rect x="9.5" y="9.5" width="5" height="5" rx=".5" fill="#d1fae5"/></svg>',
            'onlyoffice_docx': '<svg width="48" height="48" fill="none" viewBox="0 0 24 24"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z" stroke="#60a5fa" stroke-width="1.5"/><path d="M14 2v6h6" stroke="#93c5fd" stroke-width="1.5"/><path d="M8 13h8M8 17h5" stroke="#bfdbfe" stroke-width="1.5" stroke-linecap="round"/></svg>',
        };
        const labels = {
            'onlyoffice_pptx': { title: '프레젠테이션', desc: '지침을 입력하고 생성 버튼을 클릭하세요', color: '#8b5cf6', bg: 'rgba(139,92,246,.08)', bgHover: 'rgba(139,92,246,.14)' },
            'onlyoffice_xlsx': { title: '스프레드시트', desc: '지침을 입력하고 생성 버튼을 클릭하세요', color: '#10b981', bg: 'rgba(16,185,129,.08)', bgHover: 'rgba(16,185,129,.14)' },
            'onlyoffice_docx': { title: '문서', desc: '지침을 입력하고 Word 생성 버튼을 클릭하세요', color: '#3b82f6', bg: 'rgba(59,130,246,.08)', bgHover: 'rgba(59,130,246,.14)' },
        };
        const icon = icons[projectType] || icons['onlyoffice_docx'];
        const info = labels[projectType] || labels['onlyoffice_docx'];
        $('#onlyofficeContainer').html(`
            <div class="oo-empty-state">
                <div class="oo-empty-icon" style="background:${info.bg}">${icon}</div>
                <div class="oo-empty-title" style="color:${info.color}">${info.title}</div>
                <div class="oo-empty-desc">${info.desc}</div>
            </div>
        `);
    }
}

function updateFullscreenInputHeight() {
    const inputBar = document.querySelector('.input-bar');
    if (inputBar) {
        const h = inputBar.offsetHeight;
        document.documentElement.style.setProperty('--fullscreen-input-h', h + 'px');
    }
}

function toggleCanvasFullscreen() {
    const $app = $('#appView');
    const isFullscreen = $app.hasClass('canvas-fullscreen');
    const enterIcon = '<path d="M8 3H5a2 2 0 00-2 2v3m18 0V5a2 2 0 00-2-2h-3m0 18h3a2 2 0 002-2v-3M3 16v3a2 2 0 002 2h3"/>';
    const exitIcon = '<path d="M4 14h2a2 2 0 012 2v2m8-4h2a2 2 0 002-2V8m-8 16v-2a2 2 0 012-2h2M4 10V8a2 2 0 012-2h2"/>';

    if (isFullscreen) {
        $app.removeClass('canvas-fullscreen');
        $('.canvas-fullscreen-exit').remove();
        $('#btnCanvasFullscreen').attr('title', '전체화면');
        $('#icoCanvasFullscreen').html(enterIcon);
    } else {
        $app.addClass('canvas-fullscreen');
        $('body').append('<button class="canvas-fullscreen-exit" onclick="toggleCanvasFullscreen()" title="전체화면 해제"><svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">' + exitIcon + '</svg></button>');
        $('#btnCanvasFullscreen').attr('title', '전체화면 해제');
        $('#icoCanvasFullscreen').html(exitIcon);
        setTimeout(updateFullscreenInputHeight, 50);
    }

    setTimeout(() => { window.dispatchEvent(new Event('resize')); }, 100);
}

function toggleOnlyOfficeFullscreen() {
    const $app = $('#appView');
    const isFullscreen = $app.hasClass('oo-fullscreen');

    if (isFullscreen) {
        $app.removeClass('oo-fullscreen');
        $('.oo-fullscreen-exit').remove();
        $('#btnFullscreenOO').attr('title', '전체화면');
        $('#icoFullscreenOO').html('<path d="M8 3H5a2 2 0 00-2 2v3m18 0V5a2 2 0 00-2-2h-3m0 18h3a2 2 0 002-2v-3M3 16v3a2 2 0 002 2h3"/>');
    } else {
        $app.addClass('oo-fullscreen');
        $('body').append('<button class="oo-fullscreen-exit" onclick="toggleOnlyOfficeFullscreen()" title="전체화면 해제"><svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path d="M4 14h2a2 2 0 012 2v2m8-4h2a2 2 0 002-2V8m-8 16v-2a2 2 0 012-2h2M4 10V8a2 2 0 012-2h2"/></svg></button>');
        $('#btnFullscreenOO').attr('title', '전체화면 해제');
        $('#icoFullscreenOO').html('<path d="M4 14h2a2 2 0 012 2v2m8-4h2a2 2 0 002-2V8m-8 16v-2a2 2 0 012-2h2M4 10V8a2 2 0 012-2h2"/>');
        setTimeout(updateFullscreenInputHeight, 50);
    }

    if (state.onlyofficeEditor) {
        setTimeout(() => { window.dispatchEvent(new Event('resize')); }, 100);
    }
}

function downloadOnlyOfficeDoc() {
    if (!state.currentProject || !state.onlyofficeDoc) {
        showToast('다운로드할 문서가 없습니다', 'error');
        return;
    }
    const ext = { 'onlyoffice_pptx': '.pptx', 'onlyoffice_xlsx': '.xlsx', 'onlyoffice_docx': '.docx' }[state.currentProject.project_type] || '.docx';
    const url = `/${state.jwtToken}/api/generate/${state.currentProject._id}/download/onlyoffice`;
    _downloadFile(url, (state.currentProject.name || 'document') + ext);
}

// OnlyOffice SSE 진행 상태 표시
let _ooGenTimerInterval = null;
let _ooGenStartTime = 0;

function _showOnlyOfficeProgress(msg) {
    $('#onlyofficeProgressOverlay').show();
    $('#onlyofficeProgressMsg').text(msg || '준비 중...');
}

function _hideOnlyOfficeProgress() {
    $('#onlyofficeProgressOverlay').hide();
    $('#onlyofficeProgressBar').css('width', '0%');
    $('#ooGenContent').empty().show();
    _ooGenClearItems();
    _ooGenStopTimer();
    // 모든 step 초기화
    $('.oo-gen-step').removeClass('active done');
    $('.oo-gen-step-line').removeClass('done');
}

function _ooGenSetStep(stepName) {
    const order = ['search', 'generate', 'parse', 'file', 'editor'];
    const idx = order.indexOf(stepName);
    if (idx < 0) return;
    const $steps = $('.oo-gen-step');
    const $lines = $('.oo-gen-step-line');
    $steps.each(function (i) {
        const $s = $(this);
        if (i < idx) {
            $s.removeClass('active').addClass('done');
        } else if (i === idx) {
            $s.removeClass('done').addClass('active');
        } else {
            $s.removeClass('active done');
        }
    });
    $lines.each(function (i) {
        $(this).toggleClass('done', i < idx);
    });
}

function _ooGenStartTimer() {
    _ooGenStartTime = Date.now();
    _ooGenStopTimer();
    _ooGenUpdateTimer();
    _ooGenTimerInterval = setInterval(_ooGenUpdateTimer, 1000);
}
function _ooGenStopTimer() {
    if (_ooGenTimerInterval) { clearInterval(_ooGenTimerInterval); _ooGenTimerInterval = null; }
}
function _ooGenUpdateTimer() {
    const elapsed = Math.floor((Date.now() - _ooGenStartTime) / 1000);
    const m = Math.floor(elapsed / 60);
    const s = elapsed % 60;
    $('#ooGenTimer').text(`${m}:${s.toString().padStart(2, '0')}`);
}

function _ooGenAppendContent(text) {
    const el = document.getElementById('ooGenContent');
    if (!el) return;
    el.textContent += text;
    el.scrollTop = el.scrollHeight;
}

function _ooGenUpdateSmartSummary(fullText, projectType) {
    const el = document.getElementById('ooGenContent');
    if (!el) return;
    let itemCount = 0, lastTitle = '', charCount = fullText.length;
    if (projectType === 'onlyoffice_pptx') {
        itemCount = (fullText.match(/"template_index"/g) || []).length;
        const titles = fullText.match(/"title"\s*:\s*\{\s*"text"\s*:\s*"([^"]{1,80})"/g);
        if (titles && titles.length) { const m = titles[titles.length - 1].match(/"text"\s*:\s*"([^"]{1,80})"/); if (m) lastTitle = m[1]; }
        el.textContent = `AI가 프레젠테이션을 설계하고 있습니다...\n\n${itemCount > 0 ? `슬라이드 ${itemCount}개 생성 중` : '구조 분석 중'}${lastTitle ? ` — "${lastTitle}"` : ''}\n\n${(charCount / 1000).toFixed(1)}K 문자 수신`;
    } else if (projectType === 'onlyoffice_xlsx') {
        itemCount = (fullText.match(/"name"\s*:/g) || []).length;
        const names = fullText.match(/"name"\s*:\s*"([^"]{1,60})"/g);
        if (names && names.length) { const m = names[names.length - 1].match(/"name"\s*:\s*"([^"]{1,60})"/); if (m) lastTitle = m[1]; }
        el.textContent = `AI가 데이터를 구조화하고 있습니다...\n\n${itemCount > 0 ? `시트 ${itemCount}개` : '구조 분석 중'}${lastTitle ? ` — "${lastTitle}"` : ''}\n\n${(charCount / 1000).toFixed(1)}K 문자 수신`;
    } else if (projectType === 'onlyoffice_docx') {
        itemCount = (fullText.match(/"title"\s*:/g) || []).length;
        const titles = fullText.match(/"title"\s*:\s*"([^"]{1,80})"/g);
        if (titles && titles.length) { const m = titles[titles.length - 1].match(/"title"\s*:\s*"([^"]{1,80})"/); if (m) lastTitle = m[1]; }
        el.textContent = `AI가 문서를 작성하고 있습니다...\n\n${itemCount > 0 ? `섹션 ${itemCount}개 생성 중` : '구조 분석 중'}${lastTitle ? ` — "${lastTitle}"` : ''}\n\n${(charCount / 1000).toFixed(1)}K 문자 수신`;
    }
}

function _ooGenAddItem(current, total, title, type, detail) {
    const container = $('#ooGenItems');
    container.show();
    container.find('.oo-gen-item.active').removeClass('active').addClass('done');
    const typeLabel = type === 'slide' ? '슬라이드' : type === 'sheet' ? '시트' : '섹션';
    const detailHtml = detail ? `<span class="oo-gen-item-detail">${detail}</span>` : '';
    container.append(`<div class="oo-gen-item active" data-index="${current}"><span class="oo-gen-item-icon"></span><span class="oo-gen-item-text">${typeLabel} ${current}/${total}: ${title}</span>${detailHtml}</div>`);
    container[0].scrollTop = container[0].scrollHeight;
}

function _ooGenClearItems() {
    $('#ooGenItems').empty().hide();
}

async function _onlyofficeStreamGenerate(url, body) {
    _showOnlyOfficeProgress('준비 중...');
    _ooGenStartTimer();
    $('#ooGenContent').empty();
    _ooGenClearItems();
    _showStopButton();

    let projectType = 'onlyoffice_pptx';
    if (url.includes('/xlsx/')) projectType = 'onlyoffice_xlsx';
    else if (url.includes('/docx/')) projectType = 'onlyoffice_docx';

    state.currentProject.status = 'generating';
    $('#wsProjectStatus').text(t('statusGenerating')).attr('class', 'ws-status generating');

    try {
        const response = await fetch(`/${state.jwtToken}${url}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body),
        });

        if (!response.ok) {
            let errMsg = `HTTP ${response.status}`;
            try {
                const errBody = await response.json();
                errMsg = errBody.detail || errBody.message || errMsg;
            } catch {}
            _hideOnlyOfficeProgress();
            _showGenerateOrRestartButton();
            state.currentProject.status = 'draft';
            $('#wsProjectStatus').text(t('statusDraft')).attr('class', 'ws-status draft');
            showToast(errMsg, 'error');
            return;
        }

        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';
        let llmText = '';

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            buffer += decoder.decode(value, { stream: true });
            const lines = buffer.split('\n');
            buffer = lines.pop();

            for (const line of lines) {
                if (!line.startsWith('data: ')) continue;
                let evt;
                try { evt = JSON.parse(line.substring(6)); } catch { continue; }

                switch (evt.event) {
                    case 'searching':
                        _ooGenSetStep('search');
                        _showOnlyOfficeProgress('인터넷에서 자료를 검색하고 있습니다...');
                        break;
                    case 'search_done':
                        _ooGenSetStep('search');
                        _showOnlyOfficeProgress('검색 완료! 콘텐츠를 생성합니다...');
                        if (evt.result_count) {
                            _ooGenAppendContent(`[검색 완료] ${evt.result_count}건의 자료를 찾았습니다.\n\n`);
                        }
                        break;
                    case 'start':
                        _ooGenSetStep('generate');
                        _showOnlyOfficeProgress(evt.message || 'AI가 콘텐츠를 생성하고 있습니다...');
                        break;
                    case 'delta':
                        llmText += evt.text || '';
                        _ooGenUpdateSmartSummary(llmText, projectType);
                        const progress = Math.min(llmText.length / 5000 * 70, 70);
                        $('#onlyofficeProgressBar').css('width', progress + '%');
                        break;
                    case 'parsing':
                        _ooGenSetStep('parse');
                        _showOnlyOfficeProgress(evt.message || '콘텐츠를 구성하고 있습니다...');
                        $('#onlyofficeProgressBar').css('width', '75%');
                        $('#ooGenContent').hide();
                        break;
                    case 'item_progress':
                        _ooGenSetStep('parse');
                        const ipType = evt.type === 'slide' ? '슬라이드' : evt.type === 'sheet' ? '시트' : '섹션';
                        _showOnlyOfficeProgress(`${ipType} ${evt.current}/${evt.total} 처리 중: ${evt.title}`);
                        _ooGenAddItem(evt.current, evt.total, evt.title, evt.type, evt.detail);
                        const ipProgress = 70 + (evt.current / evt.total) * 15;
                        $('#onlyofficeProgressBar').css('width', ipProgress + '%');
                        break;
                    case 'file_creating':
                        _ooGenSetStep('file');
                        _showOnlyOfficeProgress(evt.message || '파일을 생성하고 있습니다...');
                        $('#onlyofficeProgressBar').css('width', '85%');
                        break;
                    case 'onlyoffice_ready':
                        _ooGenSetStep('editor');
                        _showOnlyOfficeProgress('OnlyOffice 에디터를 열고 있습니다...');
                        $('#onlyofficeProgressBar').css('width', '95%');
                        state.onlyofficeDoc = evt.document || null;
                        break;
                    case 'complete':
                        // Mark remaining items as done
                        $('#ooGenItems .oo-gen-item.active').removeClass('active').addClass('done');
                        _hideOnlyOfficeProgress();
                        _showGenerateOrRestartButton();
                        state.currentProject.status = 'generated';
                        $('#wsProjectStatus').text(t('statusGenerated')).attr('class', 'ws-status generated');
                        showToast('생성 완료!', 'success');
                        // 에디터 열기
                        if (state.onlyofficeDoc) {
                            await openOnlyOfficeEditor(state.currentProject._id);
                        }
                        break;
                    case 'stopped':
                        _hideOnlyOfficeProgress();
                        _showGenerateOrRestartButton();
                        state.currentProject.status = 'stopped';
                        $('#wsProjectStatus').text(t('statusStopped')).attr('class', 'ws-status stopped');
                        showToast(evt.message || '중단됨', 'info');
                        break;
                    case 'error':
                        _hideOnlyOfficeProgress();
                        _showGenerateOrRestartButton();
                        showToast(evt.message || '생성 실패', 'error');
                        break;
                }
            }
        }

        // 스트림이 끝났는데 complete/error 이벤트 없이 종료된 경우
        if ($('#onlyofficeProgressOverlay').is(':visible')) {
            _hideOnlyOfficeProgress();
            _showGenerateOrRestartButton();
        }
    } catch (e) {
        _hideOnlyOfficeProgress();
        _showGenerateOrRestartButton();
        showToast('생성 중 오류: ' + e.message, 'error');
    }
}

async function generateOnlyOfficePptx() {
    const templateId = state.selectedTemplateId;
    if (!templateId) {
        showToast(t('msgSelectTemplate'), 'error');
        return;
    }
    const instructions = $('#instructionsInput').val().trim();
    const lang = $('#langSelect').val();
    const slideCount = $('#slideCountSelect').val();

    await _onlyofficeStreamGenerate('/api/generate/onlyoffice/pptx/stream', {
        project_id: state.currentProject._id,
        template_id: templateId,
        instructions,
        lang,
        slide_count: slideCount,
    });
}

async function generateOnlyOfficeXlsx() {
    const instructions = $('#instructionsInput').val().trim();
    const lang = $('#langSelect').val();

    if (state.resources.length === 0 && !instructions) {
        showToast('리소스 또는 지침을 입력하세요.', 'error');
        return;
    }

    await _onlyofficeStreamGenerate('/api/generate/onlyoffice/xlsx/stream', {
        project_id: state.currentProject._id,
        instructions,
        lang,
    });
}

async function generateOnlyOfficeDocx() {
    const instructions = $('#instructionsInput').val().trim();
    const lang = $('#langSelect').val();

    if (state.resources.length === 0 && !instructions) {
        showToast('리소스 또는 지침을 입력하세요.', 'error');
        return;
    }

    _isGenerating = true;
    _showStopButton();
    state.currentProject.status = 'generating';
    $('#wsProjectStatus').text(t('statusGenerating')).attr('class', 'ws-status generating');

    // 기존 에디터 숨기고 HTML 미리보기로 스트리밍
    destroyOnlyOfficeEditor();
    _showOoDocxHtmlPreview();

    const sectionCount = $('#docxPageCountSelect').val();
    await _onlyofficeDocxHtmlStream('/api/generate/onlyoffice/docx/stream', {
        project_id: state.currentProject._id,
        instructions,
        lang,
        section_count: sectionCount,
        docx_template_id: state.docxTemplateId || null,
    });
}

// ============ OnlyOffice 실시간 스트리밍 (Word) ============

let _ooConnector = null;
let _ooDocReadyResolve = null;

async function _openOnlyOfficeForStreaming(projectId) {
    const res = await apiGet(`/api/onlyoffice/${projectId}/config`);
    const config = res.config;
    const onlyofficeUrl = res.onlyoffice_url;

    await loadOnlyOfficeScript(onlyofficeUrl);
    destroyOnlyOfficeEditor();
    _ooConnector = null;

    // 에디터 컨테이너 + 하단 플로팅 진행 표시
    $('#onlyofficeContainer').html(
        '<div id="ooEditorDiv" style="width:100%;height:100%;"></div>' +
        '<div id="ooStreamOverlay" style="position:absolute;bottom:60px;left:50%;transform:translateX(-50%);' +
        'background:rgba(30,30,50,0.85);color:#fff;padding:10px 24px;border-radius:24px;font-size:13px;' +
        'display:flex;align-items:center;gap:10px;z-index:10000;box-shadow:0 4px 20px rgba(0,0,0,0.3);' +
        'backdrop-filter:blur(8px);pointer-events:none;">' +
        '<div style="width:16px;height:16px;border:2px solid rgba(255,255,255,0.3);border-top-color:#fff;' +
        'border-radius:50%;animation:spin 0.8s linear infinite;"></div>' +
        '<span id="ooStreamMsg">AI가 문서를 작성하고 있습니다...</span>' +
        '<span id="ooStreamTimer" style="opacity:0.6;margin-left:4px;">0:00</span>' +
        '</div>'
    );

    const docReadyPromise = new Promise(resolve => { _ooDocReadyResolve = resolve; });

    config.events = {
        onAppReady: function() {
            console.log('[OnlyOffice-Stream] 에디터 준비 완료');
        },
        onDocumentReady: function() {
            console.log('[OnlyOffice-Stream] 문서 로드 완료');
            // Connector 생성 시도
            try {
                _ooConnector = state.onlyofficeEditor.createConnector();
                console.log('[OnlyOffice-Stream] Connector 생성 완료');
            } catch (e) {
                console.warn('[OnlyOffice-Stream] Connector 생성 실패 (버전 미지원):', e);
                _ooConnector = null;
            }
            if (_ooDocReadyResolve) {
                _ooDocReadyResolve();
                _ooDocReadyResolve = null;
            }
        },
        onError: function(event) {
            console.error('[OnlyOffice-Stream] 에디터 오류:', event.data);
            if (_ooDocReadyResolve) {
                _ooDocReadyResolve();
                _ooDocReadyResolve = null;
            }
        }
    };

    state.onlyofficeEditor = new DocsAPI.DocEditor('ooEditorDiv', config);

    // iframe 권한 설정
    setTimeout(function() {
        var ooIframe = document.querySelector('#ooEditorDiv iframe');
        if (ooIframe) {
            ooIframe.setAttribute('allow', 'clipboard-read; clipboard-write; unload');
        }
    }, 500);

    // 에디터 로드 대기
    await docReadyPromise;
}


let _ooStreamTimerInterval = null;
let _ooStreamStartTime = 0;

function _startStreamTimer() {
    _ooStreamStartTime = Date.now();
    if (_ooStreamTimerInterval) clearInterval(_ooStreamTimerInterval);
    _updateStreamTimer();
    _ooStreamTimerInterval = setInterval(_updateStreamTimer, 1000);
}

function _stopStreamTimer() {
    if (_ooStreamTimerInterval) { clearInterval(_ooStreamTimerInterval); _ooStreamTimerInterval = null; }
}

function _updateStreamTimer() {
    const elapsed = Math.floor((Date.now() - _ooStreamStartTime) / 1000);
    const m = Math.floor(elapsed / 60);
    const s = elapsed % 60;
    const el = document.getElementById('ooStreamTimer');
    if (el) el.textContent = `${m}:${s.toString().padStart(2, '0')}`;
}


async function _onlyofficeDocxStreamWithEditor(url, body) {
    _startStreamTimer();

    try {
        const response = await fetch(`/${state.jwtToken}${url}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body),
        });

        if (!response.ok) {
            let errMsg = `HTTP ${response.status}`;
            try {
                const errBody = await response.json();
                errMsg = errBody.detail || errBody.message || errMsg;
            } catch {}
            _isGenerating = false;
            _stopStreamTimer();
            $('#ooStreamOverlay').remove();
            _showGenerateOrRestartButton();
            state.currentProject.status = 'draft';
            $('#wsProjectStatus').text(t('statusDraft')).attr('class', 'ws-status draft');
            showToast(errMsg, 'error');
            return;
        }

        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';
        let sectionCount = 0;
        const _ooStreamRequestedTotal = parseInt(body.slide_count || body.section_count) || 0; // 요청한 전체 수

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            buffer += decoder.decode(value, { stream: true });
            const lines = buffer.split('\n');
            buffer = lines.pop();

            for (const line of lines) {
                if (!line.startsWith('data: ')) continue;
                let evt;
                try { evt = JSON.parse(line.substring(6)); } catch { continue; }

                switch (evt.event) {
                    case 'searching':
                        $('#ooStreamMsg').text('인터넷에서 자료를 검색하고 있습니다...');
                        break;
                    case 'search_done':
                        $('#ooStreamMsg').text('검색 완료! 문서를 작성합니다...');
                        break;
                    case 'start':
                        $('#ooStreamMsg').text('AI가 문서를 작성하고 있습니다...');
                        break;
                    case 'delta':
                        // 진행 상황만 표시 (section_ready에서 실제 삽입)
                        break;
                    case 'meta_ready':
                        // 메타 데이터는 섹션 리로드 시 함께 반영됨
                        break;
                    case 'section_ready':
                        sectionCount++;
                        const section = evt.section;
                        const _ooTotal = _ooStreamRequestedTotal > 0 ? Math.max(_ooStreamRequestedTotal, sectionCount) : sectionCount;
                        $('#ooStreamMsg').text(`섹션 작성 중 (${sectionCount}/${_ooTotal}): ${section.title || ''}`);
                        // 1. 커넥터로 실시간 삽입 시도
                        if (_ooConnector) {
                            _insertSectionViaConnector(section, sectionCount === 1);
                        }
                        // 2. 서버사이드: 문서 업데이트 후 에디터 리로드
                        if (evt.document_updated) {
                            state.onlyofficeDoc = evt.document_updated;
                            _debounceEditorReload();
                        }
                        break;
                    case 'file_creating':
                        $('#ooStreamMsg').text('최종 문서를 정리하고 있습니다...');
                        break;
                    case 'file_updated':
                        state.onlyofficeDoc = evt.document || null;
                        break;
                    case 'complete':
                        _isGenerating = false;
                        // 대기 중인 리로드 타이머 취소
                        if (_editorReloadTimer) { clearTimeout(_editorReloadTimer); _editorReloadTimer = null; }
                        _editorReloadPending = false;
                        _stopStreamTimer();
                        $('#ooStreamOverlay').fadeOut(300, function() { $(this).remove(); });
                        _showGenerateOrRestartButton();
                        state.currentProject.status = 'generated';
                        $('#wsProjectStatus').text(t('statusGenerated')).attr('class', 'ws-status generated');
                        showToast('생성 완료!', 'success');
                        // 최종 스타일 적용된 문서로 에디터 새로고침
                        if (state.onlyofficeDoc) {
                            // 리로드 진행 중이면 완료 후 최종 리로드
                            var _finalReload = async () => {
                                if (_editorReloading) {
                                    setTimeout(_finalReload, 500);
                                    return;
                                }
                                await openOnlyOfficeEditor(state.currentProject._id);
                            };
                            setTimeout(_finalReload, 300);
                        }
                        break;
                    case 'stopped':
                        _isGenerating = false;
                        if (_editorReloadTimer) { clearTimeout(_editorReloadTimer); _editorReloadTimer = null; }
                        _editorReloadPending = false;
                        _stopStreamTimer();
                        $('#ooStreamOverlay').fadeOut(300, function() { $(this).remove(); });
                        _showGenerateOrRestartButton();
                        state.currentProject.status = 'stopped';
                        $('#wsProjectStatus').text(t('statusStopped')).attr('class', 'ws-status stopped');
                        showToast(evt.message || '중단됨', 'info');
                        break;
                    case 'error':
                        _isGenerating = false;
                        _stopStreamTimer();
                        $('#ooStreamOverlay').fadeOut(300, function() { $(this).remove(); });
                        _showGenerateOrRestartButton();
                        showToast(evt.message || '생성 실패', 'error');
                        break;
                }
            }
        }

        // 스트림이 끝났는데 complete/error 없이 종료된 경우
        if ($('#ooStreamOverlay').is(':visible')) {
            _isGenerating = false;
            _stopStreamTimer();
            $('#ooStreamOverlay').remove();
            _showGenerateOrRestartButton();
        }
    } catch (e) {
        _isGenerating = false;
        _stopStreamTimer();
        $('#ooStreamOverlay').remove();
        _showGenerateOrRestartButton();
        showToast('생성 오류: ' + e.message, 'error');
    }
}


function _insertMetaViaConnector(meta) {
    if (!_ooConnector) return;
    try {
        var paras = [];
        if (meta.title) {
            paras.push({ text: meta.title, bold: true, size: 44, color: [26, 26, 46], font: '맑은 고딕', align: 'center' });
        }
        if (meta.description) {
            paras.push({ text: meta.description, size: 20, italic: true, color: [100, 116, 139], font: '맑은 고딕', align: 'center' });
            paras.push({ text: '' });
        }
        _pushParagraphsToEditor(paras);
    } catch (e) {
        console.warn('[OO-Stream] meta insert failed:', e);
    }
}


function _insertSectionViaConnector(section, isFirst) {
    if (!_ooConnector) return;
    try {
        var paras = _parseSectionToParagraphs(section);
        _pushParagraphsToEditor(paras);
    } catch (e) {
        console.warn('[OO-Stream] section insert failed:', e);
    }
}

function _parseSectionToParagraphs(section) {
    var result = [];
    var title = section.title || '';
    var level = section.level || 1;
    var content = section.content || '';

    // 섹션 제목
    if (title) {
        var fontSize, color;
        if (level <= 1) { fontSize = 40; color = [26, 26, 46]; }
        else if (level === 2) { fontSize = 30; color = [30, 41, 59]; }
        else if (level === 3) { fontSize = 26; color = [51, 65, 81]; }
        else { fontSize = 22; color = [71, 85, 105]; }
        result.push({ text: title, bold: true, size: fontSize, color: color, font: '맑은 고딕' });
    }

    // 본문 파싱
    if (content) {
        var lines = content.split('\n');
        for (var i = 0; i < lines.length; i++) {
            var line = lines[i].trim();
            if (!line) continue;

            // 마크다운 헤딩
            var hMatch = line.match(/^(#{1,4})\s+(.+)/);
            if (hMatch) {
                var hLevel = hMatch[1].length;
                var hText = hMatch[2].replace(/\*\*/g, '');
                var fs, cl;
                if (hLevel === 1) { fs = 36; cl = [26, 26, 46]; }
                else if (hLevel === 2) { fs = 28; cl = [30, 41, 59]; }
                else if (hLevel === 3) { fs = 24; cl = [51, 65, 81]; }
                else { fs = 22; cl = [71, 85, 105]; }
                result.push({ text: hText, bold: true, size: fs, color: cl, font: '맑은 고딕' });
                continue;
            }

            // 인용문
            if (line.charAt(0) === '>') {
                result.push({ text: line.replace(/^>\s*/, ''), italic: true, color: [79, 70, 229], font: '맑은 고딕' });
                continue;
            }

            // 불릿 리스트
            var bm = line.match(/^[-*]\s+(.+)/);
            if (bm) {
                result.push({ text: '• ' + bm[1].replace(/\*\*/g, '').replace(/\*/g, ''), font: '맑은 고딕' });
                continue;
            }

            // 숫자 리스트
            var nm = line.match(/^(\d+)\.\s+(.+)/);
            if (nm) {
                result.push({ text: nm[1] + '. ' + nm[2].replace(/\*\*/g, '').replace(/\*/g, ''), font: '맑은 고딕' });
                continue;
            }

            // 테이블/구분선 스킵
            if (line.charAt(0) === '|') continue;
            if (line.match(/^[-=]{3,}$/)) continue;

            // 일반 텍스트
            var plain = line.replace(/\*\*(.+?)\*\*/g, '$1').replace(/\*(.+?)\*/g, '$1').replace(/`(.+?)`/g, '$1');
            result.push({ text: plain, font: '맑은 고딕' });
        }
    }

    // 빈 줄 구분
    result.push({ text: '' });
    return result;
}


function _pushParagraphsToEditor(paras) {
    if (!_ooConnector || !paras.length) return;
    try {
        var json = JSON.stringify(paras);
        // Unicode escape: 모든 비-ASCII 문자와 특수문자를 \\uXXXX로 변환
        // atob/btoa 없이 안전하게 데이터를 함수 본문에 삽입
        var safe = '';
        for (var i = 0; i < json.length; i++) {
            var ch = json.charAt(i);
            var c = json.charCodeAt(i);
            if (c > 127) {
                safe += '\\u' + c.toString(16).padStart(4, '0');
            } else if (ch === '\\') {
                safe += '\\\\';
            } else if (ch === '"') {
                safe += '\\"';
            } else if (ch === '\n') {
                safe += '\\n';
            } else if (ch === '\r') {
                safe += '\\r';
            } else {
                safe += ch;
            }
        }

        var fnBody =
            'var _ps = JSON.parse("' + safe + '");' +
            'var oDoc = Api.GetDocument();' +
            'for (var i = 0; i < _ps.length; i++) {' +
            '  var p = _ps[i];' +
            '  var oPara = Api.CreateParagraph();' +
            '  if (p.align) oPara.SetJc(p.align);' +
            '  if (p.text) {' +
            '    var oRun = oPara.AddText(p.text);' +
            '    if (p.bold) oRun.SetBold(true);' +
            '    if (p.italic) oRun.SetItalic(true);' +
            '    if (p.size) oRun.SetFontSize(p.size);' +
            '    if (p.color) oRun.SetColor(p.color[0], p.color[1], p.color[2]);' +
            '    if (p.font) oRun.SetFontFamily(p.font);' +
            '  }' +
            '  oDoc.Push(oPara);' +
            '}';

        _ooConnector.callCommand(new Function(fnBody), true);
        console.log('[OO-Stream] Pushed', paras.length, 'paragraphs to editor');
        return true;
    } catch (e) {
        console.warn('[OO-Stream] pushParagraphs failed:', e);
        return false;
    }
}

let _editorReloadTimer = null;
let _editorReloading = false;
let _editorReloadPending = false;
let _lastReloadSectionCount = 0;

function _debounceEditorReload() {
    // 리로드 진행 중이면 대기열에 추가
    if (_editorReloading) {
        _editorReloadPending = true;
        return;
    }
    if (_editorReloadTimer) clearTimeout(_editorReloadTimer);
    _editorReloadTimer = setTimeout(async () => {
        _editorReloadTimer = null;
        _editorReloading = true;
        _editorReloadPending = false;
        try {
            console.log('[OO-Stream] 에디터 리로드 시작...');

            // 기존 에디터 위에 로딩 오버레이 표시 (깜빡임 방지)
            var $container = $('#onlyofficeContainer');
            $container.append('<div id="ooReloadOverlay" style="position:absolute;top:0;left:0;right:0;bottom:0;background:rgba(255,255,255,0.7);z-index:9999;display:flex;align-items:center;justify-content:center;"><div style="background:rgba(30,30,50,0.85);color:#fff;padding:10px 20px;border-radius:20px;font-size:13px;">문서 업데이트 중...</div></div>');

            await openOnlyOfficeEditor(state.currentProject._id);
            console.log('[OO-Stream] 에디터 리로드 완료');

            // 로딩 오버레이 제거
            $('#ooReloadOverlay').remove();

            // 스크롤을 문서 끝으로 이동 시도
            _scrollEditorToBottom();

            // 리로드 후 플로팅 진행 오버레이 복원
            if ($('#ooStreamOverlay').length === 0 && _isGenerating) {
                var elapsed = _ooStreamStartTime ? Math.floor((Date.now() - _ooStreamStartTime) / 1000) : 0;
                var m = Math.floor(elapsed / 60);
                var s = elapsed % 60;
                $('#onlyofficeContainer').append(
                    '<div id="ooStreamOverlay" style="position:absolute;bottom:60px;left:50%;transform:translateX(-50%);' +
                    'background:rgba(30,30,50,0.85);color:#fff;padding:10px 24px;border-radius:24px;font-size:13px;' +
                    'display:flex;align-items:center;gap:10px;z-index:10000;box-shadow:0 4px 20px rgba(0,0,0,0.3);' +
                    'backdrop-filter:blur(8px);pointer-events:none;">' +
                    '<div style="width:16px;height:16px;border:2px solid rgba(255,255,255,0.3);border-top-color:#fff;' +
                    'border-radius:50%;animation:spin 0.8s linear infinite;"></div>' +
                    '<span id="ooStreamMsg">AI가 문서를 작성하고 있습니다...</span>' +
                    '<span id="ooStreamTimer" style="opacity:0.6;margin-left:4px;">' + m + ':' + s.toString().padStart(2, '0') + '</span>' +
                    '</div>'
                );
            }
        } catch (e) {
            console.warn('[OO-Stream] Editor reload failed:', e);
            $('#ooReloadOverlay').remove();
        } finally {
            _editorReloading = false;
            // 대기 중인 리로드가 있으면 다시 실행
            if (_editorReloadPending) {
                _editorReloadPending = false;
                _debounceEditorReload();
            }
        }
    }, 3000); // 3초 디바운스 - 여러 섹션을 묶어서 한 번에 리로드
}


function _scrollEditorToBottom() {
    // 에디터 로드 후 문서 끝으로 스크롤
    setTimeout(function() {
        try {
            // 방법 1: Ctrl+End 키 이벤트 시뮬레이션
            var iframe = document.querySelector('#ooEditorDiv iframe');
            if (iframe && iframe.contentDocument) {
                var body = iframe.contentDocument.body || iframe.contentDocument.querySelector('.doc-body');
                if (body) body.scrollTop = body.scrollHeight;
            }
        } catch (e) {
            // cross-origin iframe - 접근 불가
        }

        // 방법 2: connector로 커서를 문서 끝으로 이동
        if (_ooConnector) {
            try {
                _ooConnector.callCommand(function() {
                    var oDoc = Api.GetDocument();
                    var cnt = oDoc.GetElementsCount();
                    if (cnt > 0) {
                        var last = oDoc.GetElement(cnt - 1);
                        last.MoveCursorToElement(false);
                    }
                }, false);
            } catch (e) {}
        }
    }, 1500);
}

// ============ OnlyOffice DOCX HTML 미리보기 스트리밍 ============

let _ooDocxStreamBuffer = '';
let _ooDocxStreamRenderTimer = null;
let _ooDocxChartCache = new Map();
let _ooDocxPendingCharts = [];
let _ooDocxTimerInterval = null;
let _ooDocxStartTime = 0;

function _showOoDocxHtmlPreview() {
    _ooDocxStreamBuffer = '';
    _ooDocxChartCache.clear();
    _ooDocxPendingCharts = [];

    // 에디터 컨테이너를 HTML 미리보기로 교체
    $('#onlyofficeContainer').html(
        '<div id="ooDocxPreviewWrap" class="oo-docx-preview-wrap">' +
            '<div class="oo-docx-preview-header">' +
                '<div class="oo-docx-preview-spinner"></div>' +
                '<span id="ooDocxPreviewMsg">AI가 문서를 작성하고 있습니다...</span>' +
                '<span id="ooDocxPreviewTimer" class="oo-docx-preview-timer">0:00</span>' +
            '</div>' +
            '<div id="ooDocxPreviewScroll" class="oo-docx-preview-scroll">' +
                '<div class="word-streaming-content" id="ooDocxPreviewContent"></div>' +
            '</div>' +
        '</div>'
    );

    // 타이머 시작
    _ooDocxStartTime = Date.now();
    if (_ooDocxTimerInterval) clearInterval(_ooDocxTimerInterval);
    _ooDocxTimerInterval = setInterval(_updateOoDocxTimer, 1000);
}

function _hideOoDocxHtmlPreview() {
    if (_ooDocxStreamRenderTimer) {
        clearTimeout(_ooDocxStreamRenderTimer);
        _ooDocxStreamRenderTimer = null;
    }
    if (_ooDocxTimerInterval) {
        clearInterval(_ooDocxTimerInterval);
        _ooDocxTimerInterval = null;
    }
}

function _updateOoDocxTimer() {
    const elapsed = Math.floor((Date.now() - _ooDocxStartTime) / 1000);
    const m = Math.floor(elapsed / 60);
    const s = elapsed % 60;
    const el = document.getElementById('ooDocxPreviewTimer');
    if (el) el.textContent = `${m}:${s.toString().padStart(2, '0')}`;
}

function _appendOoDocxDelta(text) {
    _ooDocxStreamBuffer += text;
    if (!_ooDocxStreamRenderTimer) {
        _ooDocxStreamRenderTimer = setTimeout(() => {
            _ooDocxStreamRenderTimer = null;
            _renderOoDocxPreview();
        }, 80);
    }
}

function _renderOoDocxPreview() {
    // Word 모듈의 기존 렌더링 함수 재사용
    const cleaned = _cleanStreamingJsonToReadable(_ooDocxStreamBuffer);

    // _streamingMarkdownToHtml는 _wordChartCache(const Map)와 _pendingChartConfigs(let)를 사용
    // OO용 차트 캐시를 임시로 주입하고, 호출 후 복원
    const savedCacheEntries = new Map(_wordChartCache);
    const savedPending = _pendingChartConfigs;
    _wordChartCache.clear();
    _ooDocxChartCache.forEach((v, k) => _wordChartCache.set(k, v));
    // _streamingMarkdownToHtml 내부에서 _pendingChartConfigs = [] 재할당됨
    _pendingChartConfigs = _ooDocxPendingCharts;

    const html = _streamingMarkdownToHtml(cleaned);

    // 함수 내부에서 _pendingChartConfigs가 새 배열로 교체되었으므로 캡처
    _ooDocxPendingCharts = _pendingChartConfigs;
    // word 쪽 원래 상태 복원
    _pendingChartConfigs = savedPending;
    _ooDocxChartCache.clear();
    _wordChartCache.forEach((v, k) => _ooDocxChartCache.set(k, v));
    _wordChartCache.clear();
    savedCacheEntries.forEach((v, k) => _wordChartCache.set(k, v));

    const container = document.getElementById('ooDocxPreviewContent');
    if (!container) return;
    container.innerHTML = html + '<span class="word-typing-cursor"></span>';

    // 차트 렌더링
    _renderOoDocxInlineCharts(container);

    // 자동 스크롤
    const scroll = document.getElementById('ooDocxPreviewScroll');
    if (scroll) scroll.scrollTop = scroll.scrollHeight;
}

function _renderOoDocxInlineCharts(container) {
    const pendingDivs = container.querySelectorAll('.docx-stream-chart-pending');
    if (!pendingDivs.length) return;

    if (typeof echarts === 'undefined') {
        loadEChartsScript().catch(() => {});
        return;
    }

    pendingDivs.forEach(div => {
        const idx = parseInt(div.getAttribute('data-chart-idx'));
        if (isNaN(idx) || !_ooDocxPendingCharts[idx]) return;
        const jsonStr = _ooDocxPendingCharts[idx];
        try {
            const config = JSON.parse(jsonStr);
            const option = _buildEChartsOptionFromChartJson(config);
            const offDiv = document.createElement('div');
            offDiv.style.cssText = 'width:900px;height:450px;position:absolute;left:-9999px;top:-9999px;';
            document.body.appendChild(offDiv);
            const chart = echarts.init(offDiv);
            chart.setOption(option);
            const dataUrl = chart.getDataURL({ type: 'png', pixelRatio: 2, backgroundColor: '#fff' });
            chart.dispose();
            document.body.removeChild(offDiv);
            _ooDocxChartCache.set(jsonStr, dataUrl);
            const safeJson = jsonStr.replace(/&/g,'&amp;').replace(/"/g,'&quot;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
            div.innerHTML = `<img src="${dataUrl}" style="max-width:100%;border-radius:8px;" data-chart-json="${safeJson}" />`;
            div.className = 'docx-stream-chart';
        } catch (e) {
            console.error('[OO-DOCX Preview] chart render error:', e);
        }
    });
}

function _showOoDocxLoadingOverlay(msg) {
    // 미리보기 위에 로딩 오버레이 표시
    if ($('#ooDocxLoadingOverlay').length) {
        $('#ooDocxLoadingMsg').text(msg);
        return;
    }
    $('#ooDocxPreviewWrap').append(
        '<div id="ooDocxLoadingOverlay" class="oo-docx-loading-overlay">' +
            '<div class="oo-docx-loading-card">' +
                '<div class="oo-docx-loading-pulse"></div>' +
                '<div class="oo-docx-loading-icon">' +
                    '<svg width="32" height="32" fill="none" stroke="currentColor" stroke-width="1.5" viewBox="0 0 24 24">' +
                        '<path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>' +
                        '<polyline points="14 2 14 8 20 8"/>' +
                        '<line x1="16" y1="13" x2="8" y2="13"/>' +
                        '<line x1="16" y1="17" x2="8" y2="17"/>' +
                        '<polyline points="10 9 9 9 8 9"/>' +
                    '</svg>' +
                '</div>' +
                '<p id="ooDocxLoadingMsg" class="oo-docx-loading-msg">' + (msg || '') + '</p>' +
                '<div class="oo-docx-loading-dots"><span></span><span></span><span></span></div>' +
            '</div>' +
        '</div>'
    );
}

function _updateOoDocxLoadingMsg(msg) {
    $('#ooDocxLoadingMsg').text(msg);
}

function _removeOoDocxLoadingOverlay() {
    $('#ooDocxLoadingOverlay').remove();
}

async function _loadOoEditorBehindPreview(projectId) {
    // 미리보기 뒤에 에디터를 로드하고, onDocumentReady에서 미리보기 제거
    try {
        const res = await apiGet(`/api/onlyoffice/${projectId}/config`);
        const config = res.config;
        const onlyofficeUrl = res.onlyoffice_url;

        await loadOnlyOfficeScript(onlyofficeUrl);

        // 기존 에디터 정리
        if (state.onlyofficeEditor) {
            try { state.onlyofficeEditor.destroyEditor(); } catch (e) {}
            state.onlyofficeEditor = null;
        }

        // 미리보기 뒤에 에디터 div 삽입 (z-index로 미리보기 아래 위치)
        if (!$('#ooEditorDiv').length) {
            $('#onlyofficeContainer').prepend('<div id="ooEditorDiv" style="width:100%;height:100%;position:absolute;top:0;left:0;z-index:0;"></div>');
        }
        // 미리보기를 위에 표시
        $('#ooDocxPreviewWrap').css({ position: 'relative', 'z-index': '10' });

        config.events = {
            onAppReady: function() {
                console.log('[OO-DocxPreview] 에디터 준비 완료');
            },
            onDocumentReady: function() {
                console.log('[OO-DocxPreview] 문서 로드 완료 → 미리보기 제거');
                $('#btnDownloadOO').show();
                $('#btnFullscreenOO').show();
                // 미리보기 + 로딩 오버레이 제거, 에디터 표시
                $('#ooDocxPreviewWrap').fadeOut(400, function() {
                    $(this).remove();
                    // 에디터 div를 정상 위치로
                    $('#ooEditorDiv').css({ position: '', top: '', left: '', 'z-index': '' });
                });
                showToast('문서 생성 완료!', 'success');
            },
            onError: function(event) {
                console.error('[OO-DocxPreview] 에디터 오류:', event.data);
                _removeOoDocxLoadingOverlay();
                $('#ooDocxPreviewWrap').remove();
                $('#ooEditorDiv').css({ position: '', top: '', left: '', 'z-index': '' });
                showToast('OnlyOffice 에디터 오류가 발생했습니다', 'error');
            },
        };

        state.onlyofficeEditor = new DocsAPI.DocEditor('ooEditorDiv', config);

        setTimeout(function() {
            var ooIframe = document.querySelector('#ooEditorDiv iframe');
            if (ooIframe) ooIframe.setAttribute('allow', 'clipboard-read; clipboard-write; unload');
        }, 500);

    } catch (e) {
        console.error('[OO-DocxPreview] 에디터 로드 실패:', e);
        _removeOoDocxLoadingOverlay();
        $('#ooDocxPreviewWrap').remove();
        showToast('에디터 로드 실패: ' + e.message, 'error');
    }
}

async function _onlyofficeDocxHtmlStream(url, body) {
    try {
        const response = await fetch(`/${state.jwtToken}${url}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body),
        });

        if (!response.ok) {
            let errMsg = `HTTP ${response.status}`;
            try {
                const errBody = await response.json();
                errMsg = errBody.detail || errBody.message || errMsg;
            } catch {}
            _isGenerating = false;
            _hideOoDocxHtmlPreview();
            _showGenerateOrRestartButton();
            state.currentProject.status = 'draft';
            $('#wsProjectStatus').text(t('statusDraft')).attr('class', 'ws-status draft');
            showToast(errMsg, 'error');
            return;
        }

        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';
        let sectionCount = 0;
        const _ooDocxRequestedTotal = parseInt(body.section_count) || 0; // 사용자가 요청한 전체 섹션 수
        let _ooDocxTotalSections = 0; // delta에서 지속 추적
        let _ooDocxDeltaCount = 0;

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            buffer += decoder.decode(value, { stream: true });
            const lines = buffer.split('\n');
            buffer = lines.pop();

            for (const line of lines) {
                if (!line.startsWith('data: ')) continue;
                let evt;
                try { evt = JSON.parse(line.substring(6)); } catch { continue; }

                switch (evt.event) {
                    case 'searching':
                        $('#ooDocxPreviewMsg').text('인터넷에서 자료를 검색하고 있습니다...');
                        break;
                    case 'search_done':
                        $('#ooDocxPreviewMsg').text('검색 완료! 문서를 작성합니다...');
                        break;
                    case 'start':
                        $('#ooDocxPreviewMsg').text('AI가 문서를 작성하고 있습니다...');
                        break;
                    case 'delta':
                        _appendOoDocxDelta(evt.text || '');
                        // 5 delta마다 전체 섹션 수 추정 업데이트
                        _ooDocxDeltaCount++;
                        if (_ooDocxDeltaCount % 5 === 0) {
                            const _si = _ooDocxStreamBuffer.indexOf('"sections"');
                            if (_si >= 0) {
                                const _st = _ooDocxStreamBuffer.substring(_si);
                                _ooDocxTotalSections = (_st.match(/"title"\s*:/g) || []).length;
                            }
                        }
                        break;
                    case 'meta_ready':
                        // meta는 스트리밍 텍스트에서 자동 추출되어 표시됨
                        break;
                    case 'section_ready':
                        sectionCount++;
                        const section = evt.section;
                        // 요청한 전체 섹션 수를 기준으로, 없으면 버퍼에서 추정
                        let total = _ooDocxRequestedTotal;
                        if (!total) {
                            const _si2 = _ooDocxStreamBuffer.indexOf('"sections"');
                            if (_si2 >= 0) {
                                _ooDocxTotalSections = (_ooDocxStreamBuffer.substring(_si2).match(/"title"\s*:/g) || []).length;
                            }
                            total = Math.max(_ooDocxTotalSections, sectionCount);
                        } else {
                            total = Math.max(total, sectionCount);
                        }
                        $('#ooDocxPreviewMsg').text(`섹션 작성 중 (${sectionCount}/${total}): ${section.title || ''}`);
                        break;
                    case 'file_creating':
                        // 마지막 렌더링 강제 실행
                        if (_ooDocxStreamRenderTimer) {
                            clearTimeout(_ooDocxStreamRenderTimer);
                            _ooDocxStreamRenderTimer = null;
                        }
                        _renderOoDocxPreview();
                        // 커서 제거 (작성 완료)
                        $('.word-typing-cursor', '#ooDocxPreviewContent').remove();
                        // 로딩 오버레이 표시
                        _showOoDocxLoadingOverlay('워드 문서를 생성하고 있습니다...');
                        break;
                    case 'file_updated':
                        state.onlyofficeDoc = evt.document || null;
                        break;
                    case 'complete':
                        _isGenerating = false;
                        _hideOoDocxHtmlPreview();
                        _showGenerateOrRestartButton();
                        state.currentProject.status = 'generated';
                        $('#wsProjectStatus').text(t('statusGenerated')).attr('class', 'ws-status generated');

                        if (state.onlyofficeDoc) {
                            _updateOoDocxLoadingMsg('에디터에 문서를 불러오고 있습니다...');
                            // 에디터를 미리보기 아래에 미리 로드, onDocumentReady에서 전환
                            _loadOoEditorBehindPreview(state.currentProject._id);
                        } else {
                            _removeOoDocxLoadingOverlay();
                            $('#ooDocxPreviewWrap').remove();
                            showToast('문서 생성 완료!', 'success');
                        }
                        break;
                    case 'stopped':
                        _isGenerating = false;
                        _hideOoDocxHtmlPreview();
                        _showGenerateOrRestartButton();
                        state.currentProject.status = 'stopped';
                        $('#wsProjectStatus').text(t('statusStopped')).attr('class', 'ws-status stopped');
                        // 중단 시에도 부분 생성된 문서가 있으면 에디터 열기
                        if (state.onlyofficeDoc) {
                            _showOoDocxLoadingOverlay('문서를 불러오고 있습니다...');
                            _loadOoEditorBehindPreview(state.currentProject._id);
                        } else {
                            _removeOoDocxLoadingOverlay();
                            $('#ooDocxPreviewWrap').fadeOut(300, function() { $(this).remove(); });
                        }
                        showToast(evt.message || '중단됨', 'info');
                        break;
                    case 'error':
                        _isGenerating = false;
                        _hideOoDocxHtmlPreview();
                        _showGenerateOrRestartButton();
                        $('#ooDocxPreviewWrap').fadeOut(300, function() { $(this).remove(); });
                        showToast(evt.message || '생성 실패', 'error');
                        break;
                }
            }
        }

        // 스트림이 끝났는데 complete/error 없이 종료된 경우
        if ($('#ooDocxPreviewWrap').is(':visible') && _isGenerating) {
            _isGenerating = false;
            _hideOoDocxHtmlPreview();
            _showGenerateOrRestartButton();
            $('#ooDocxPreviewWrap').remove();
        }
    } catch (e) {
        _isGenerating = false;
        _hideOoDocxHtmlPreview();
        _showGenerateOrRestartButton();
        $('#ooDocxPreviewWrap').remove();
        showToast('생성 오류: ' + e.message, 'error');
    }
}


// ============ 차트 렌더링 (Apache ECharts) ============

let _echartsLoaded = false;
let _echartsLoading = false;
let _activeExcelCharts = [];

async function loadEChartsScript() {
    if (_echartsLoaded) return;
    if (_echartsLoading) {
        while (_echartsLoading) await new Promise(r => setTimeout(r, 100));
        return;
    }
    _echartsLoading = true;
    try {
        await new Promise((resolve, reject) => {
            const s = document.createElement('script');
            s.src = 'https://cdn.jsdelivr.net/npm/echarts@5.6.0/dist/echarts.min.js';
            s.onload = resolve;
            s.onerror = reject;
            document.head.appendChild(s);
        });
        _echartsLoaded = true;
    } finally {
        _echartsLoading = false;
    }
}

function _destroyExcelCharts() {
    _activeExcelCharts.forEach(c => { try { c.dispose(); } catch(e) {} });
    _activeExcelCharts = [];
    const grid = document.getElementById('excelChartsGrid');
    if (grid) grid.innerHTML = '';
    $('#excelChartsContainer').hide();
    $('#excelResizer').hide();
    // 리사이저로 변경된 높이 초기화
    const _uniC = document.getElementById('univerContainer');
    const _chC = document.getElementById('excelChartsContainer');
    if (_uniC) { _uniC.style.flex = ''; _uniC.style.height = ''; }
    if (_chC) { _chC.style.height = ''; _chC.style.minHeight = ''; _chC.style.maxHeight = ''; }
}

async function renderExcelCharts(excelData) {
    if (!excelData || !excelData.sheets) return;

    let hasCharts = false;
    for (const sheet of excelData.sheets) {
        if (sheet.charts && sheet.charts.length > 0) {
            hasCharts = true;
            break;
        }
    }
    if (!hasCharts) {
        _destroyExcelCharts();
        return;
    }

    await loadEChartsScript();
    _destroyExcelCharts();

    const grid = document.getElementById('excelChartsGrid');
    if (!grid) return;

    const multiSheet = excelData.sheets.filter(s => s.charts && s.charts.length > 0).length > 1;

    for (const sheet of excelData.sheets) {
        const charts = sheet.charts || [];
        if (charts.length === 0) continue;

        for (const chartDef of charts) {
            const option = _buildEChartsOption(chartDef, sheet);
            if (!option) continue;

            const card = document.createElement('div');
            card.className = 'excel-chart-card';

            if (multiSheet) {
                const label = document.createElement('div');
                label.style.cssText = 'font-size:11px;color:#888;margin-bottom:4px;text-align:center;';
                label.textContent = sheet.name;
                card.appendChild(label);
            }

            const chartDiv = document.createElement('div');
            chartDiv.className = 'echarts-container';
            card.appendChild(chartDiv);
            grid.appendChild(card);

            const instance = echarts.init(chartDiv);
            instance.setOption(option);
            _activeExcelCharts.push(instance);
        }
    }

    $('#excelChartsContainer').show();
    $('#excelResizer').show();
    _initExcelResizer();

    // 컨테이너가 보인 후 차트 리사이즈 (레이아웃 반영)
    requestAnimationFrame(() => {
        _activeExcelCharts.forEach(c => { try { c.resize(); } catch(e) {} });
    });

    // 리사이즈 대응
    if (!window._echartsResizeHandler) {
        window._echartsResizeHandler = () => {
            _activeExcelCharts.forEach(c => { try { c.resize(); } catch(e) {} });
        };
        window.addEventListener('resize', window._echartsResizeHandler);
    }
}

/* ============ 엑셀/차트 리사이저 드래그 ============ */
function _initExcelResizer() {
    const resizer = document.getElementById('excelResizer');
    const univerEl = document.getElementById('univerContainer');
    const chartsEl = document.getElementById('excelChartsContainer');
    if (!resizer || !univerEl || !chartsEl || resizer._resizerInit) return;
    resizer._resizerInit = true;

    let startY = 0, startUH = 0, startCH = 0, dragging = false;

    function onMouseDown(e) {
        e.preventDefault();
        dragging = true;
        startY = e.clientY || (e.touches && e.touches[0].clientY) || 0;
        startUH = univerEl.getBoundingClientRect().height;
        startCH = chartsEl.getBoundingClientRect().height;
        resizer.classList.add('active');
        document.body.style.cursor = 'row-resize';
        document.body.style.userSelect = 'none';
    }

    function onMouseMove(e) {
        if (!dragging) return;
        const clientY = e.clientY || (e.touches && e.touches[0].clientY) || 0;
        const delta = clientY - startY;
        const minU = 120, minC = 120;
        let newUH = startUH + delta;
        let newCH = startCH - delta;
        if (newUH < minU) { newUH = minU; newCH = startUH + startCH - minU; }
        if (newCH < minC) { newCH = minC; newUH = startUH + startCH - minC; }
        univerEl.style.flex = 'none';
        univerEl.style.height = newUH + 'px';
        chartsEl.style.height = newCH + 'px';
        chartsEl.style.minHeight = '0';
        chartsEl.style.maxHeight = 'none';
    }

    function onMouseUp() {
        if (!dragging) return;
        dragging = false;
        resizer.classList.remove('active');
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
        // 차트 리사이즈 반영
        if (_activeExcelCharts && _activeExcelCharts.length) {
            _activeExcelCharts.forEach(c => { try { c.resize(); } catch(e) {} });
        }
    }

    resizer.addEventListener('mousedown', onMouseDown);
    resizer.addEventListener('touchstart', onMouseDown, { passive: false });
    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('touchmove', onMouseMove, { passive: false });
    document.addEventListener('mouseup', onMouseUp);
    document.addEventListener('touchend', onMouseUp);
}

const _CHART_PALETTE = ['#5470c6', '#91cc75', '#fac858', '#ee6666', '#73c0de', '#3ba272', '#fc8452', '#9a60b4'];

function _buildEChartsOption(chartDef, sheetData) {
    const rows = sheetData.rows || [];
    const columns = sheetData.columns || [];
    const dr = chartDef.data_range || {};
    const labelsCol = dr.labels_column || 0;
    const rowStart = dr.row_start || 0;
    const rowEnd = (dr.row_end != null) ? dr.row_end : rows.length - 1;

    const labels = [];
    for (let r = rowStart; r <= rowEnd && r < rows.length; r++) {
        labels.push(rows[r][labelsCol] != null ? String(rows[r][labelsCol]) : '');
    }

    const seriesDefs = dr.series || [];
    const isPie = chartDef.type === 'pie' || chartDef.type === 'doughnut';

    if (isPie) {
        // Pie/Doughnut
        const sDef = seriesDefs[0];
        if (!sDef) return null;
        const pieData = [];
        for (let r = rowStart; r <= rowEnd && r < rows.length; r++) {
            const v = rows[r][sDef.column];
            pieData.push({
                name: labels[r - rowStart] || '',
                value: typeof v === 'number' ? v : (parseFloat(v) || 0),
            });
        }
        return {
            color: _CHART_PALETTE,
            title: {
                text: chartDef.title || '',
                left: 'center',
                textStyle: { fontSize: 14, fontWeight: 600 },
            },
            tooltip: { trigger: 'item', formatter: '{b}: {c} ({d}%)' },
            legend: {
                show: chartDef.options?.show_legend !== false,
                bottom: 4,
                type: 'scroll',
                textStyle: { fontSize: 11 },
            },
            series: [{
                type: 'pie',
                radius: chartDef.type === 'doughnut' ? ['35%', '65%'] : '65%',
                center: ['50%', '48%'],
                data: pieData,
                label: { show: true, formatter: '{b}: {d}%' },
                emphasis: {
                    itemStyle: { shadowBlur: 10, shadowOffsetX: 0, shadowColor: 'rgba(0,0,0,0.2)' },
                },
                itemStyle: { borderRadius: chartDef.type === 'doughnut' ? 6 : 0, borderColor: '#fff', borderWidth: 2 },
            }],
        };
    }

    // Bar / Line / Area / Scatter / Radar
    const series = [];
    for (let si = 0; si < seriesDefs.length; si++) {
        const sDef = seriesDefs[si];
        const data = [];
        for (let r = rowStart; r <= rowEnd && r < rows.length; r++) {
            const v = rows[r][sDef.column];
            data.push(typeof v === 'number' ? v : (parseFloat(v) || 0));
        }
        const seriesName = sDef.name || columns[sDef.column] || ('Series ' + (si + 1));

        if (chartDef.type === 'radar') {
            series.push({ type: 'radar', name: seriesName, data: [{ value: data, name: seriesName }] });
        } else {
            const s = {
                type: chartDef.type === 'area' ? 'line' : (chartDef.type === 'scatter' ? 'scatter' : chartDef.type || 'bar'),
                name: seriesName,
                data: data,
                smooth: chartDef.type === 'line' || chartDef.type === 'area',
            };
            if (chartDef.type === 'area') {
                s.areaStyle = { opacity: 0.3 };
            }
            if (chartDef.type === 'bar') {
                s.itemStyle = { borderRadius: [4, 4, 0, 0] };
                if (chartDef.options?.stacked) s.stack = 'total';
            }
            series.push(s);
        }
    }

    if (chartDef.type === 'radar') {
        const maxVals = [];
        for (let i = 0; i < labels.length; i++) {
            let mx = 0;
            series.forEach(s => { if (s.data[0] && s.data[0].value[i] > mx) mx = s.data[0].value[i]; });
            maxVals.push(mx);
        }
        return {
            color: _CHART_PALETTE,
            title: { text: chartDef.title || '', left: 'center', textStyle: { fontSize: 14, fontWeight: 600 } },
            tooltip: {},
            legend: { show: chartDef.options?.show_legend !== false, bottom: 4, type: 'scroll', textStyle: { fontSize: 11 } },
            radar: { indicator: labels.map((l, i) => ({ name: l, max: Math.ceil(maxVals[i] * 1.2) || 100 })) },
            series: series,
        };
    }

    return {
        color: _CHART_PALETTE,
        title: {
            text: chartDef.title || '',
            left: 'center',
            textStyle: { fontSize: 14, fontWeight: 600 },
        },
        tooltip: {
            trigger: 'axis',
            axisPointer: { type: chartDef.type === 'bar' ? 'shadow' : 'line' },
        },
        legend: {
            show: chartDef.options?.show_legend !== false,
            bottom: 4,
            type: 'scroll',
            textStyle: { fontSize: 11 },
        },
        grid: { left: '3%', right: '4%', bottom: '18%', top: chartDef.title ? '15%' : '8%', containLabel: true },
        xAxis: {
            type: 'category',
            data: labels,
            axisLabel: { fontSize: 11, rotate: labels.length > 8 ? 30 : 0 },
            axisTick: { alignWithLabel: true },
        },
        yAxis: {
            type: 'value',
            splitLine: { lineStyle: { color: 'rgba(0,0,0,0.06)' } },
        },
        series: series,
    };
}

/**
 * 워드 스트리밍 차트 JSON → ECharts 옵션 변환
 * (Chart.js 호환 JSON 형식: {type, data:{labels, datasets}, options})
 */
function _buildEChartsOptionFromChartJson(config) {
    const chartType = config.type || 'bar';
    const labels = config.data?.labels || [];
    const datasets = config.data?.datasets || [];
    const chartTitle = config.options?.title || config.title || '';
    const titleText = typeof chartTitle === 'string' ? chartTitle : (chartTitle.text || '');
    const isPie = chartType === 'pie' || chartType === 'doughnut';

    if (isPie) {
        const ds = datasets[0] || {};
        const pieData = labels.map((l, i) => ({
            name: l,
            value: (ds.data || [])[i] || 0,
        }));
        return {
            animation: false,
            color: _CHART_PALETTE,
            title: { text: titleText, left: 'center', top: 12, textStyle: { fontSize: 16, fontFamily: "'Malgun Gothic', sans-serif" } },
            tooltip: { trigger: 'item', formatter: '{b}: {c} ({d}%)' },
            legend: { bottom: 12, type: 'scroll' },
            series: [{
                type: 'pie',
                radius: chartType === 'doughnut' ? ['35%', '65%'] : '65%',
                center: ['50%', '52%'],
                data: pieData,
                label: { show: true, formatter: '{b}: {d}%' },
                itemStyle: { borderRadius: chartType === 'doughnut' ? 6 : 0, borderColor: '#fff', borderWidth: 2 },
                emphasis: { itemStyle: { shadowBlur: 10, shadowColor: 'rgba(0,0,0,0.2)' } },
            }],
        };
    }

    const series = datasets.map(ds => {
        const s = {
            type: chartType === 'area' ? 'line' : chartType,
            name: ds.label || '',
            data: ds.data || [],
            smooth: chartType === 'line' || chartType === 'area',
        };
        if (chartType === 'area') s.areaStyle = { opacity: 0.3 };
        if (chartType === 'bar') s.itemStyle = { borderRadius: [4, 4, 0, 0] };
        return s;
    });

    return {
        animation: false,
        color: _CHART_PALETTE,
        title: { text: titleText, left: 'center', top: 12, textStyle: { fontSize: 16, fontFamily: "'Malgun Gothic', sans-serif" } },
        tooltip: { trigger: 'axis', axisPointer: { type: chartType === 'bar' ? 'shadow' : 'line' } },
        legend: { bottom: 12, type: 'scroll' },
        grid: { left: '3%', right: '4%', bottom: '18%', top: titleText ? '18%' : '10%', containLabel: true },
        xAxis: { type: 'category', data: labels, axisLabel: { fontSize: 11 } },
        yAxis: { type: 'value', splitLine: { lineStyle: { color: 'rgba(0,0,0,0.06)' } } },
        series: series,
    };
}
