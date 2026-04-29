const STORAGE_KEY = "eduscript-ai-studio-state-v1";
const DRIVE_SETTINGS_KEY = "eduscript-ai-drive-settings-v1";
const DRIVE_BACKUP_PREFIX = "eduscript-ai-backup-";
const DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.file";
const GOOGLE_IDENTITY_SCRIPT = "https://accounts.google.com/gsi/client";
const DEFAULT_CORE_RATIO = 0.85;
const CHECKPOINT_INTERVAL = 5;

const PROFESSIONAL_COURSE_STANDARD = {
  flow: [
    "需求分析與未指定欄位標註",
    "官方來源 / syllabus 對齊",
    "年度藍圖與依賴 timetable",
    "Lecture / PPT slide spec",
    "CA Lab resources、steps、deliverables、rubric",
    "Assessment goal、format、weight、rubric",
    "QA gate、accessibility gate、版本快照",
    "Markdown / JSON / PPTX / Course Pack 匯出",
  ],
  metadataFields: [
    "course_id",
    "course_title",
    "lecture_id",
    "module_id",
    "duration_minutes",
    "slide_target",
    "template_id",
    "difficulty",
    "prerequisites",
    "resource_profile",
    "official_alignment",
    "locale",
    "version",
    "model_name",
    "prompt_hash",
    "source_snapshot_id",
    "review_status",
  ],
  qaGates: [
    "Schema completeness",
    "時間與頁數對齊",
    "Lecture / Lab / Assessment 依賴一致",
    "技術步驟可驗收",
    "教學可用性與 checkpoint 覆蓋",
    "Accessibility：alt text、唯一標題、閱讀順序、對比",
    "版本、來源、review status 完整",
  ],
  slideSections: [
    "開場與情境",
    "學習目標與成功證據",
    "先修銜接",
    "概念地圖",
    "Terminology / CLI / 物件邊界",
    "標準流程",
    "Demo walkthrough",
    "Checkpoint",
    "常見錯誤",
    "Guided exercise",
    "Best practice / Lab bridge",
    "總結與下週預告",
  ],
};

let driveAccessToken = "";
let driveTokenClient = null;
let googleIdentityScriptPromise = null;
let driveAutoBackupTimer = null;

const bloomMap = {
  remember: {
    label: "記憶",
    verb: "列出、辨認、回想",
    strategy: "用定義卡、關鍵詞與快速提問建立共同語言。",
  },
  understand: {
    label: "理解",
    verb: "解釋、摘要、舉例",
    strategy: "用生活類比與概念圖讓學生說出自己的理解。",
  },
  apply: {
    label: "應用",
    verb: "操作、套用、解題",
    strategy: "把新概念放入真實題境，讓學生完成可觀察任務。",
  },
  analyze: {
    label: "分析",
    verb: "比較、拆解、判斷關係",
    strategy: "用對照表、因果鏈與錯誤辨析拉高思考層次。",
  },
  evaluate: {
    label: "評鑑",
    verb: "評估、辯護、提出證據",
    strategy: "要求學生基於準則作判斷，並說明證據品質。",
  },
  create: {
    label: "創造",
    verb: "設計、建構、產出",
    strategy: "讓學生整合知識，生成作品、模型或教學說明。",
  },
};

const gagneEvents = [
  { event: "引起動機", weight: 0.08, bloom: "remember" },
  { event: "提示目標", weight: 0.08, bloom: "understand" },
  { event: "喚起舊知", weight: 0.1, bloom: "remember" },
  { event: "呈現內容", weight: 0.22, bloom: "understand" },
  { event: "提供引導", weight: 0.12, bloom: "apply" },
  { event: "引發表現", weight: 0.14, bloom: "analyze" },
  { event: "提供回饋", weight: 0.1, bloom: "evaluate" },
  { event: "評量學習", weight: 0.08, bloom: "evaluate" },
  { event: "促進保留", weight: 0.08, bloom: "create" },
];

const pptTemplateCatalog = {
  title: {
    label: "封面",
    event: "引起動機",
    bloom: "understand",
    weight: 0.04,
    layout: "置中標題 + 底部 metadata bar",
    visual: "主視覺、課程關鍵字條、清楚課程定位",
    bodyBudget: 30,
    notesBudget: 100,
  },
  prerequisite: {
    label: "先備橋接",
    event: "喚起舊知",
    bloom: "understand",
    weight: 0.05,
    layout: "三欄：已具備 / 今天接上 / 今天不深入",
    visual: "bridge diagram 或 readiness checklist",
    bodyBudget: 70,
    notesBudget: 140,
  },
  objectives: {
    label: "學習目標",
    event: "提示目標",
    bloom: "understand",
    weight: 0.05,
    layout: "左側 outcomes，右側 Bloom tags",
    visual: "target ladder / checklist",
    bodyBudget: 80,
    notesBudget: 140,
  },
  agenda: {
    label: "議程",
    event: "提示目標",
    bloom: "remember",
    weight: 0.05,
    layout: "水平時間軸",
    visual: "timeline / process SmartArt",
    bodyBudget: 55,
    notesBudget: 120,
  },
  content: {
    label: "內容講解",
    event: "呈現內容",
    bloom: "understand",
    weight: 0.1,
    layout: "左概念右重點",
    visual: "diagram、2-column cards、concept map",
    bodyBudget: 90,
    notesBudget: 180,
  },
  example: {
    label: "範例",
    event: "提供引導",
    bloom: "apply",
    weight: 0.09,
    layout: "左情境右例子",
    visual: "annotated example / before-after",
    bodyBudget: 70,
    notesBudget: 180,
  },
  demo: {
    label: "Demo",
    event: "提供引導",
    bloom: "apply",
    weight: 0.13,
    layout: "上步驟下輸出",
    visual: "terminal flow、YAML snippet、expected output",
    bodyBudget: 65,
    notesBudget: 240,
  },
  exercise: {
    label: "練習",
    event: "引發表現",
    bloom: "apply",
    weight: 0.09,
    layout: "三段式：任務 / 限制 / 驗收",
    visual: "task card、checklist、timer",
    bodyBudget: 65,
    notesBudget: 160,
  },
  comparison: {
    label: "比較",
    event: "呈現內容",
    bloom: "analyze",
    weight: 0.09,
    layout: "comparison matrix / Venn",
    visual: "左右雙欄角色差異圖",
    bodyBudget: 80,
    notesBudget: 180,
  },
  pitfalls: {
    label: "易錯點",
    event: "提供回饋",
    bloom: "analyze",
    weight: 0.08,
    layout: "錯誤現象 / 先查哪裡 / 修正方向",
    visual: "diagnose ladder：get、describe、logs、events、exec",
    bodyBudget: 85,
    notesBudget: 190,
  },
  assessment: {
    label: "評量",
    event: "評量學習",
    bloom: "evaluate",
    weight: 0.07,
    layout: "問題在 slide，答案鍵在 notes",
    visual: "quiz card、rubric、acceptance criteria",
    bodyBudget: 65,
    notesBudget: 190,
  },
  summary: {
    label: "總結",
    event: "促進保留",
    bloom: "evaluate",
    weight: 0.04,
    layout: "三卡片 + next step",
    visual: "3-key recap",
    bodyBudget: 45,
    notesBudget: 110,
  },
  references: {
    label: "參考資料",
    event: "促進保留",
    bloom: "remember",
    weight: 0.02,
    layout: "official / vendor / exam 分組",
    visual: "極簡文字與 source priority",
    bodyBudget: 80,
    notesBudget: 80,
  },
};

const wpmProfiles = {
  informative: { label: "新概念解說", wpm: 132 },
  persuasive: { label: "引起動機", wpm: 154 },
  cta: { label: "作業指派", wpm: 146 },
  acoustic: { label: "大型教室", wpm: 124 },
};

const state = {
  annualPlan: null,
  slides: [],
  script: "",
  questions: [],
  materialPages: [],
  materialMeta: null,
  slideJson: [],
  versions: [],
  messages: [],
  auditLog: [],
  publishedRevision: null,
  studentQa: {
    question: "",
    answer: "",
    mode: "未提問",
    sources: [],
  },
  qaMetrics: {
    total: 0,
    grounded: 0,
    refused: 0,
    helpful: 0,
    needsTeacher: 0,
  },
  drive: {
    clientId: "",
    connected: false,
    busy: false,
    status: "未連接 Google Drive",
    lastBackup: null,
    lastBackupAt: null,
    lastLocalChange: null,
    pendingReason: "",
    autoBackup: false,
    backups: [],
  },
  gamma: {
    configured: false,
    exportAs: "pptx",
    lastGeneration: null,
    status: "未設定 Gamma API Key，可先匯出 Gamma-ready prompt。",
  },
  assessmentBank: null,
  interviewAnswers: "",
  role: "teacher",
  lastLessonInputs: null,
  budget: null,
  ai: {
    checked: false,
    enabled: false,
    provider: "",
    model: "",
    message: "AI 未檢查",
    busy: false,
    lastCheckedAt: null,
  },
};

const dom = {};

document.addEventListener("DOMContentLoaded", async () => {
  bindDom();
  bindEvents();
  restoreDriveSettings();
  await loadAppConfig();
  restoreState();
  await checkAiHealth();
  if (!state.slides.length && state.ai.enabled) {
    await generateLesson();
  }
  renderAll();
});

function bindDom() {
  dom.navItems = document.querySelectorAll(".nav-item");
  dom.views = document.querySelectorAll(".view");
  dom.annualModule = document.getElementById("annualModuleInput");
  dom.annualAudience = document.getElementById("annualAudienceInput");
  dom.annualWeeks = document.getElementById("annualWeeksInput");
  dom.annualLectureHours = document.getElementById("annualLectureHoursInput");
  dom.annualLabHours = document.getElementById("annualLabHoursInput");
  dom.annualAssessmentHours = document.getElementById("annualAssessmentHoursInput");
  dom.annualSlidesPerHour = document.getElementById("annualSlidesPerHourInput");
  dom.annualContext = document.getElementById("annualContextInput");
  dom.annualWeeklyList = document.getElementById("annualWeeklyListInput");
  dom.annualOfficialRefs = document.getElementById("annualOfficialRefsInput");
  dom.annualResourceConstraints = document.getElementById("annualResourceConstraintsInput");
  dom.annualLectureTopics = document.getElementById("annualLectureTopicsInput");
  dom.annualLabSpec = document.getElementById("annualLabSpecInput");
  dom.annualAssessmentSpec = document.getElementById("annualAssessmentSpecInput");
  dom.annualMetrics = document.getElementById("annualMetrics");
  dom.annualNote = document.getElementById("annualNote");
  dom.annualProfessionalStandard = document.getElementById("annualProfessionalStandard");
  dom.annualLectureStatus = document.getElementById("annualLectureStatus");
  dom.annualLecturePlan = document.getElementById("annualLecturePlan");
  dom.annualTimetableStatus = document.getElementById("annualTimetableStatus");
  dom.annualTimetable = document.getElementById("annualTimetable");
  dom.annualLabPlan = document.getElementById("annualLabPlan");
  dom.annualAssessmentPlan = document.getElementById("annualAssessmentPlan");
  dom.annualContentTitle = document.getElementById("annualContentTitle");
  dom.annualContentOutput = document.getElementById("annualContentOutput");
  dom.lessonForm = document.getElementById("lessonForm");
  dom.topic = document.getElementById("topicInput");
  dom.subject = document.getElementById("subjectInput");
  dom.audience = document.getElementById("audienceInput");
  dom.duration = document.getElementById("durationInput");
  dom.style = document.getElementById("styleInput");
  dom.objective = document.getElementById("objectiveInput");
  dom.context = document.getElementById("contextInput");
  dom.bloomChecks = document.querySelectorAll("#bloomChecks input");
  dom.questionList = document.getElementById("questionList");
  dom.questionAnswer = document.getElementById("questionAnswerInput");
  dom.questionAnswerStatus = document.getElementById("questionAnswerStatus");
  dom.timeline = document.getElementById("timeline");
  dom.slideGrid = document.getElementById("slideGrid");
  dom.slideSelect = document.getElementById("slideSelect");
  dom.slideFeedback = document.getElementById("slideFeedbackInput");
  dom.slideCount = document.getElementById("slideCount");
  dom.durationStatus = document.getElementById("durationStatus");
  dom.scriptWordStatus = document.getElementById("scriptWordStatus");
  dom.versionStatus = document.getElementById("versionStatus");
  dom.publishStatus = document.getElementById("publishStatus");
  dom.groundedRateStatus = document.getElementById("groundedRateStatus");
  dom.roleSelect = document.getElementById("roleSelect");
  dom.materialFile = document.getElementById("materialFileInput");
  dom.materialStatus = document.getElementById("materialStatus");
  dom.materialText = document.getElementById("materialTextInput");
  dom.startPage = document.getElementById("startPageInput");
  dom.scriptMinutes = document.getElementById("scriptMinutesInput");
  dom.coreMinutes = document.getElementById("coreMinutesInput");
  dom.wpmProfile = document.getElementById("wpmProfileInput");
  dom.pace = document.getElementById("paceInput");
  dom.timeBudget = document.getElementById("timeBudget");
  dom.wpmStatus = document.getElementById("wpmStatus");
  dom.coreMinutesStatus = document.getElementById("coreMinutesStatus");
  dom.targetWordsStatus = document.getElementById("targetWordsStatus");
  dom.scriptGoalStatus = document.getElementById("scriptGoalStatus");
  dom.scriptOutput = document.getElementById("scriptOutput");
  dom.assistantContext = document.getElementById("assistantContextInput");
  dom.assistantQuestion = document.getElementById("assistantQuestionInput");
  dom.chatLog = document.getElementById("chatLog");
  dom.publishedSummary = document.getElementById("publishedSummary");
  dom.studentQuestion = document.getElementById("studentQuestionInput");
  dom.studentAnswer = document.getElementById("studentAnswer");
  dom.sourceList = document.getElementById("sourceList");
  dom.versionList = document.getElementById("versionList");
  dom.auditLog = document.getElementById("auditLog");
  dom.governanceMetrics = document.getElementById("governanceMetrics");
  dom.driveClientId = document.getElementById("driveClientIdInput");
  dom.autoDriveBackup = document.getElementById("autoDriveBackupInput");
  dom.driveStatus = document.getElementById("driveStatus");
  dom.driveSyncMeta = document.getElementById("driveSyncMeta");
  dom.driveBackupList = document.getElementById("driveBackupList");
  dom.gammaStatus = document.getElementById("gammaStatus");
  dom.gammaResult = document.getElementById("gammaResult");
  dom.compareBox = document.getElementById("compareBox");
  dom.aiStatus = document.getElementById("aiStatus");
}

function bindEvents() {
  dom.navItems.forEach((item) => {
    item.addEventListener("click", () => switchView(item.dataset.view));
  });

  dom.lessonForm.addEventListener("submit", async (event) => {
    event.preventDefault();
    await generateLesson();
    switchView("builder");
  });

  document.getElementById("askQuestionsBtn").addEventListener("click", renderQuestions);
  document.getElementById("applyQuestionAnswersBtn").addEventListener("click", regenerateLessonFromInterviewAnswers);
  document.getElementById("sendSlidesToScriptBtn").addEventListener("click", sendSlidesToScriptMaterial);
  document.getElementById("generateAnnualPlanBtn").addEventListener("click", generateAnnualPlan);
  document.getElementById("exportAnnualMdBtn").addEventListener("click", exportAnnualMarkdown);
  document.getElementById("exportAnnualJsonBtn").addEventListener("click", exportAnnualJson);
  document.getElementById("generateAllLabsBtn").addEventListener("click", generateAllLabContent);
  document.getElementById("generateAssessmentBankBtn").addEventListener("click", generateAssessmentBank);
  document.getElementById("copyAnnualContentBtn").addEventListener("click", copyAnnualGeneratedContent);
  document.getElementById("regenerateSlideBtn").addEventListener("click", regenerateSelectedSlide);
  document.getElementById("loadDemoBtn").addEventListener("click", loadDemoProject);
  document.getElementById("saveVersionBtn").addEventListener("click", saveVersion);
  document.getElementById("publishLessonBtn").addEventListener("click", publishLesson);
  document.getElementById("exportJsonBtn").addEventListener("click", exportProjectJson);
  document.getElementById("exportProjectJsonBtn").addEventListener("click", exportProjectJson);
  document.getElementById("importProjectJsonBtn").addEventListener("click", () => document.getElementById("importProjectInput").click());
  document.getElementById("importProjectInput").addEventListener("change", importProjectJson);
  document.getElementById("exportLessonMdBtn").addEventListener("click", exportLessonMarkdown);
  document.getElementById("exportMarkdownBtn").addEventListener("click", exportLessonMarkdown);
  document.getElementById("exportPptxBtn").addEventListener("click", exportPptx);
  document.getElementById("exportCoursePackBtn").addEventListener("click", exportCoursePack);
  document.getElementById("exportGammaDeckBtn").addEventListener("click", exportGammaDeck);
  document.getElementById("copyPromptBtn").addEventListener("click", copyPrompt);
  document.getElementById("clearProjectBtn").addEventListener("click", clearProject);
  dom.aiStatus.addEventListener("click", (event) => {
    if (event.target.closest("[data-ai-refresh]")) {
      refreshAiHealth();
    }
  });
  document.getElementById("connectDriveBtn").addEventListener("click", connectGoogleDrive);
  document.getElementById("backupDriveBtn").addEventListener("click", () => backupToGoogleDrive());
  document.getElementById("listDriveBackupsBtn").addEventListener("click", () => listGoogleDriveBackups(true));
  document.getElementById("restoreLatestDriveBtn").addEventListener("click", restoreLatestGoogleDriveBackup);
  dom.driveClientId.addEventListener("change", () => {
    state.drive.clientId = clean(dom.driveClientId.value);
    persistDriveSettings();
    renderDrivePanel();
  });
  dom.autoDriveBackup.addEventListener("change", () => {
    state.drive.autoBackup = dom.autoDriveBackup.checked;
    persistDriveSettings();
    if (state.drive.autoBackup && state.drive.pendingReason && driveAccessToken) {
      scheduleDriveAutoBackup(state.drive.pendingReason);
    }
    renderDrivePanel();
  });
  dom.roleSelect.addEventListener("change", () => setRole(dom.roleSelect.value));
  dom.questionAnswer.addEventListener("change", () => {
    state.interviewAnswers = clean(dom.questionAnswer.value);
    renderQuestions();
    markDriveBackupNeeded("AI 追問回答更新");
    persistState();
  });

  dom.materialFile.addEventListener("change", handleMaterialUpload);
  document.getElementById("generateScriptBtn").addEventListener("click", generateScript);
  document.getElementById("shortenScriptBtn").addEventListener("click", () => reviseScript("shorten"));
  document.getElementById("expandScriptBtn").addEventListener("click", () => reviseScript("expand"));
  document.getElementById("completeScriptBtn").addEventListener("click", completeScriptToTarget);
  document.getElementById("addLectureDepthBtn").addEventListener("click", addCoreLectureDepth);
  dom.scriptOutput.addEventListener("input", () => {
    state.script = dom.scriptOutput.value;
    persistState();
    renderScriptGoal();
    renderStatus();
  });

  document.getElementById("sendAssistantBtn").addEventListener("click", sendAssistantMessage);
  document.getElementById("askStudentBtn").addEventListener("click", askPublishedLesson);
  document.getElementById("feedbackHelpfulBtn").addEventListener("click", () => recordStudentFeedback("helpful"));
  document.getElementById("feedbackNeedsTeacherBtn").addEventListener("click", () => recordStudentFeedback("needs_teacher"));
  dom.assistantQuestion.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      sendAssistantMessage();
    }
  });
  document.querySelectorAll(".quick-prompts button").forEach((button) => {
    button.addEventListener("click", () => {
      dom.assistantQuestion.value = button.dataset.prompt;
      sendAssistantMessage();
    });
  });

  dom.scriptMinutes.addEventListener("change", () => {
    syncDefaultCoreMinutes();
    state.budget = null;
    renderTimeBudget();
  });
  dom.coreMinutes.addEventListener("change", () => {
    dom.coreMinutes.dataset.autoCore = "false";
    state.budget = null;
    renderTimeBudget();
  });
  [dom.wpmProfile, dom.pace].forEach((input) => {
    input.addEventListener("change", () => {
      state.budget = null;
      renderTimeBudget();
    });
  });
}

function switchView(view) {
  dom.navItems.forEach((item) => item.classList.toggle("active", item.dataset.view === view));
  dom.views.forEach((panel) => panel.classList.toggle("active", panel.dataset.viewPanel === view));
}

function setRole(role) {
  state.role = role;
  if (dom.roleSelect.value !== role) dom.roleSelect.value = role;
  logAudit("角色切換", `目前角色：${roleLabel(role)}`);
  renderAnnualPlan();
  applyRolePermissions();
  if (role === "student") switchView("student");
  persistState();
}

function roleLabel(role) {
  const labels = {
    teacher: "教師",
    ta: "助教",
    student: "學生",
    admin: "管理者",
  };
  return labels[role] || "教師";
}

function applyRolePermissions() {
  if (!dom.roleSelect) return;
  dom.roleSelect.value = state.role || "teacher";
  const canEdit = ["teacher", "admin"].includes(state.role);
  const canAssist = ["teacher", "ta", "admin"].includes(state.role);
  const canPublish = ["teacher", "admin"].includes(state.role);

  setDisabled(["generateAnnualPlanBtn", "exportAnnualMdBtn", "exportAnnualJsonBtn", "generateAllLabsBtn", "generateAssessmentBankBtn", "copyAnnualContentBtn", "generateLessonBtn", "regenerateSlideBtn", "sendSlidesToScriptBtn", "applyQuestionAnswersBtn", "generateScriptBtn", "shortenScriptBtn", "expandScriptBtn", "completeScriptBtn", "addLectureDepthBtn", "saveVersionBtn", "exportJsonBtn", "exportProjectJsonBtn", "importProjectJsonBtn", "exportLessonMdBtn", "exportMarkdownBtn", "exportPptxBtn", "exportCoursePackBtn", "exportGammaDeckBtn", "copyPromptBtn"], !canEdit);
  setDisabled(["connectDriveBtn", "backupDriveBtn", "listDriveBackupsBtn", "restoreLatestDriveBtn"], !canEdit || state.drive.busy);
  setDisabled(["publishLessonBtn"], !canPublish);
  setDisabled(["sendAssistantBtn"], !canAssist);
}

function setDisabled(ids, disabled) {
  ids.forEach((id) => {
    const element = document.getElementById(id);
    if (element) element.disabled = disabled;
  });
}

async function checkAiHealth() {
  if (window.location.protocol === "file:") {
    state.ai = {
      checked: true,
      enabled: false,
      provider: "",
      model: "",
      message: "AI 未連線",
      busy: false,
      lastCheckedAt: new Date().toISOString(),
    };
    renderAiStatus();
    return;
  }

  try {
    const response = await fetch("/api/health");
    const data = await response.json();
    state.ai = {
      checked: true,
      enabled: Boolean(data.aiEnabled),
      provider: data.provider || "local",
      model: data.model || "",
      openAiCompatible: data.openAiCompatible || null,
      geminiModelTier: data.geminiModelTier || null,
      geminiGeneration: data.geminiGeneration || null,
      message: data.aiEnabled ? `${formatAiProviderName(data.provider)} 已連線` : "AI 未設定",
      busy: false,
      lastCheckedAt: new Date().toISOString(),
    };
  } catch {
    state.ai = {
      checked: true,
      enabled: false,
      provider: "",
      model: "",
      message: "伺服器未連線",
      busy: false,
      lastCheckedAt: new Date().toISOString(),
    };
  }
  renderAiStatus();
}

async function refreshAiHealth() {
  state.ai.busy = true;
  state.ai.message = "重新檢查 AI";
  renderAiStatus();
  await checkAiHealth();
}

async function loadAppConfig() {
  if (window.location.protocol === "file:") return;
  try {
    const response = await fetch("/api/config");
    if (!response.ok) return;
    const data = await response.json();
    if (data.googleDriveClientId && !state.drive.clientId) {
      state.drive.clientId = data.googleDriveClientId;
      state.drive.status = "已載入本機 Google Drive Client ID";
      persistDriveSettings();
      renderDrivePanel();
    }
    state.gamma.configured = Boolean(data.gammaConfigured);
    state.gamma.exportAs = data.gammaExportAs || "pptx";
    state.gamma.status = state.gamma.configured
      ? `Gamma API 已設定，將嘗試直接生成 ${state.gamma.exportAs.toUpperCase()}。`
      : "未設定 Gamma API Key，可先匯出 Gamma-ready prompt。";
    renderGammaPanel();
  } catch {
    // Optional local config only; the app still works if unavailable.
  }
}

async function requestAi(type, payload) {
  return requestAiWithRetry(type, payload, getAiRetryCount(type));
}

async function requestAiOnce(type, payload) {
  if (window.location.protocol === "file:") {
    throw new Error("AI 生成需要以 node server.js 啟動後使用 http://localhost:4173。");
  }

  if (!state.ai.checked) {
    await checkAiHealth();
  }

  if (!state.ai.enabled) {
    throw new Error("所有生成位置都必須使用已連線的 AI。請在 Codespaces Secrets / .env 設定 AI_PROVIDER=openai-compatible、OPENAI_COMPAT_BASE_URL、OPENAI_COMPAT_MODEL 與 OPENAI_COMPAT_API_KEY，然後重新啟動 server。");
  }

  const response = await fetch(`/api/ai/${type}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const raw = await response.text().catch(() => "");
    let error = {};
    try {
      error = raw ? JSON.parse(raw) : {};
    } catch {
      error = {};
    }
    const detail = error.error || error.message || clean(raw.replace(/<[^>]+>/g, " ")).slice(0, 220);
    if (response.status === 502 && !detail) {
      throw new Error("Codespaces / server 回傳 502。通常是 server 未重啟、port proxy timeout，或單次 AI request 太大。請查看 terminal log。");
    }
    const aiError = new Error(detail || `AI request failed: ${response.status}`);
    aiError.status = response.status;
    if (error.diagnostics) aiError.diagnostics = error.diagnostics;
    throw aiError;
  }

  return response.json();
}

function getLessonInputs() {
  const selectedBloom = Array.from(dom.bloomChecks)
    .filter((input) => input.checked)
    .map((input) => input.value);

  return {
    topic: clean(dom.topic.value) || "未命名課題",
    subject: clean(dom.subject.value) || "跨學科",
    audience: clean(dom.audience.value) || "學生",
    duration: clamp(Number(dom.duration.value) || 45, 10, 180),
    style: dom.style.value,
    objective: clean(dom.objective.value),
    context: clean(dom.context.value),
    interviewAnswers: clean(dom.questionAnswer?.value || state.interviewAnswers),
    bloom: selectedBloom.length ? selectedBloom : ["remember", "understand", "analyze"],
  };
}

function getAnnualInputs() {
  return {
    moduleTitle: clean(dom.annualModule.value) || "進階 Linux、Kubernetes CKA/CKAD 與雲端 EKS",
    audience: clean(dom.annualAudience.value) || "高級文憑 / IVE 雲端與系統管理學生",
    weeks: clamp(Number(dom.annualWeeks.value) || 30, 6, 52),
    lectureHours: clamp(Number(dom.annualLectureHours.value) || 13, 1, 80),
    labHours: clamp(Number(dom.annualLabHours.value) || 18, 1, 80),
    assessmentHours: clamp(Number(dom.annualAssessmentHours.value) || 6, 1, 40),
    slidesPerHour: clamp(Number(dom.annualSlidesPerHour.value) || 16, 8, 35),
    context: clean(dom.annualContext.value),
    weeklyList: splitPlanLines(dom.annualWeeklyList?.value || ""),
    officialRefs: splitPlanItems(dom.annualOfficialRefs?.value || ""),
    resourceConstraints: splitPlanItems(dom.annualResourceConstraints?.value || ""),
    lectureTopics: splitPlanItems(dom.annualLectureTopics.value),
    labSpec: splitPlanLines(dom.annualLabSpec.value),
    assessmentSpec: splitPlanLines(dom.annualAssessmentSpec.value),
  };
}

async function generateAnnualPlan() {
  const inputs = getAnnualInputs();
  const seedPlan = buildAnnualPlan(inputs);
  setAiBusy(true, "AI 生成全年規劃中");
  try {
    const aiPlan = await requestAi("annual-plan", { inputs, seedPlan });
    state.annualPlan = mergeAnnualAiPlan(seedPlan, aiPlan);
    state.assessmentBank = null;
    logAudit("年度規劃", `AI 生成 ${inputs.moduleTitle} 全年 Lecture / Lab / Assessment 藍圖`);
    renderAnnualPlan();
    markDriveBackupNeeded("AI 年度規劃");
    persistState();
  } catch (error) {
    console.warn(error);
    dom.annualNote.textContent = `AI 年度規劃生成失敗：${error.message}`;
    if (dom.compareBox) dom.compareBox.textContent = `AI 年度規劃生成失敗：${error.message}`;
  } finally {
    setAiBusy(false);
  }
}

function buildAnnualPlan(inputs) {
  const weeklyItems = parseWeeklyList(inputs.weeklyList, inputs);
  const weeklyLectures = weeklyItems.filter((item) => item.type === "Lecture/PPT");
  const weeklyLabs = weeklyItems.filter((item) => item.type === "CA Lab");
  const weeklyAssessments = weeklyItems.filter((item) => item.type === "Assessment");
  const lectureCount = Math.max(1, weeklyLectures.length || Math.ceil(inputs.lectureHours));
  const lectureHoursEach = inputs.lectureHours / lectureCount;
  const labSpecs = weeklyLabs.length ? weeklyLabs.map(formatWeeklyItemAsSpec) : inputs.labSpec.length ? inputs.labSpec : defaultLabSpecs();
  const labHoursEach = inputs.labHours / labSpecs.length;
  const lectureTopics = buildLectureTopics(inputs.lectureTopics, lectureCount, weeklyLectures);
  const lectureUnits = lectureTopics.map((topic, index) => {
    const weeklyItem = weeklyLectures[index] || null;
    return buildLectureUnit(topic, index, lectureCount, weeklyItem?.hours || lectureHoursEach, inputs, weeklyItem);
  });
  const labs = labSpecs.map((line, index) => buildLabUnit(line, index, labHoursEach));
  labs.forEach((lab, index) => {
    const weeklyItem = weeklyLabs[index];
    if (weeklyItem) {
      lab.week = weeklyItem.week;
      lab.sourceWeeklyItem = weeklyItem.raw;
    }
  });
  const assessments = buildAssessmentPlan(inputs, weeklyAssessments);
  const timetable = weeklyItems.length
    ? buildTimetableFromWeeklyItems({ weeklyItems, lectureUnits, labs, assessments, inputs })
    : buildAnnualTimetable({ lectureUnits, labs, assessments, inputs });
  const pptSlides = lectureUnits.reduce((sum, unit) => sum + unit.pptSlides, 0);

  return {
    id: cryptoId(),
    generatedAt: new Date().toISOString(),
    inputs,
    weeklyItems,
    metrics: {
      totalHours: roundOne(inputs.lectureHours + inputs.labHours + inputs.assessmentHours),
      lectureHours: inputs.lectureHours,
      labHours: inputs.labHours,
      assessmentHours: inputs.assessmentHours,
      lectureUnits: lectureUnits.length,
      pptSlides,
      labCount: labs.length,
      assessmentCount: assessments.length,
      recordingHours: inputs.lectureHours,
    },
    pptConsolidation: [
      "CKA 與 CKAD 共用的 Kubernetes 架構、kubectl 基礎、YAML 結構只保留一次。",
      "CKA 重點放在 cluster admin、networking、storage、troubleshooting；CKAD 重點放在 workload、config、probes、jobs。",
      "每 1 小時 lecture 拆成 1 個錄影 batch 與 1 份 PPT deck，方便重錄與後期剪輯。",
    ],
    professionalStandard: buildProfessionalStandard(inputs, weeklyItems),
    metadataContract: buildMetadataContract(inputs),
    qaGates: buildQaGates(inputs),
    accessibilityChecklist: buildAccessibilityChecklist(),
    lectureUnits,
    labs,
    assessments,
    timetable,
    generatedContent: {
      title: "Lab / Assessment 內容生成區",
      markdown: "",
      type: "",
      index: null,
      updatedAt: null,
    },
    assessmentBank: null,
    complianceNotes: [
      "CA 筆試題目應使用教師自建、公開授權或 AI 生成後人工審核的原創題；避免使用未授權題庫。",
      "EA Skill Test 保持 no hint；所有 API / Service endpoint 必須 public，並在 rubric 中列明驗收方法。",
      "所有 AI 生成教材需保留教師審核紀錄與版本備份。",
      "若使用者貼入現有每週清單，系統會保留週次脈絡，再補齊 lecture metadata、PPT slide spec、Lab/Assessment rubric 與 QA gate。",
    ],
  };
}

function mergeAnnualAiPlan(seedPlan, aiPlan) {
  const plan = structuredCloneSafe(seedPlan);
  plan.aiGenerated = {
    provider: state.ai?.provider || "ai",
    generatedAt: new Date().toISOString(),
    summary: aiPlan?.summary || "",
  };

  if (Array.isArray(aiPlan?.lectureUnits)) {
    plan.lectureUnits = plan.lectureUnits.map((unit, index) => mergeLectureAiUnit(unit, aiPlan.lectureUnits[index], plan.inputs, index));
  }

  if (Array.isArray(aiPlan?.labs)) {
    plan.labs = plan.labs.map((lab, index) => ({
      ...lab,
      ...(aiPlan.labs[index] || {}),
      id: lab.id,
      week: lab.week,
    }));
  }

  if (Array.isArray(aiPlan?.assessments)) {
    plan.assessments = plan.assessments.map((assessment, index) => ({
      ...assessment,
      ...(aiPlan.assessments[index] || {}),
      type: assessment.type,
      week: assessment.week,
      hours: assessment.hours,
    }));
  }

  if (Array.isArray(aiPlan?.pptConsolidation) && aiPlan.pptConsolidation.length) {
    plan.pptConsolidation = aiPlan.pptConsolidation;
  }
  if (Array.isArray(aiPlan?.qaGates) && aiPlan.qaGates.length) {
    plan.qaGates = aiPlan.qaGates.map((name, index) => ({
      id: `QA-${String(index + 1).padStart(2, "0")}`,
      name,
      passRule: "AI generated gate + teacher review",
      blocking: index < 2 ? "export" : "publish",
    }));
  }
  if (Array.isArray(aiPlan?.accessibilityChecklist) && aiPlan.accessibilityChecklist.length) {
    plan.accessibilityChecklist = aiPlan.accessibilityChecklist;
  }
  if (Array.isArray(aiPlan?.complianceNotes) && aiPlan.complianceNotes.length) {
    plan.complianceNotes = aiPlan.complianceNotes;
  }

  recalculateAnnualMetrics(plan);
  syncAnnualLectureTimetable(plan);
  return plan;
}

function mergeLectureAiUnit(unit, aiUnit, inputs, index) {
  if (!aiUnit) return unit;
  const merged = {
    ...unit,
    title: clean(aiUnit.title) || unit.title,
    subtopics: Array.isArray(aiUnit.subtopics) && aiUnit.subtopics.length ? aiUnit.subtopics.map(clean).filter(Boolean) : unit.subtopics,
    outcomes: Array.isArray(aiUnit.outcomes) && aiUnit.outcomes.length ? aiUnit.outcomes : unit.outcomes,
    pptFocus: Array.isArray(aiUnit.pptFocus) && aiUnit.pptFocus.length ? aiUnit.pptFocus : unit.pptFocus,
    recordingCue: aiUnit.recordingCue || unit.recordingCue,
    duplicateCleanup: aiUnit.duplicateCleanup || unit.duplicateCleanup,
    slideSpec: Array.isArray(aiUnit.slideSpec) && aiUnit.slideSpec.length ? aiUnit.slideSpec : unit.slideSpec,
    pptxChecklist: Array.isArray(aiUnit.pptxChecklist) && aiUnit.pptxChecklist.length ? aiUnit.pptxChecklist : unit.pptxChecklist,
    qaChecklist: Array.isArray(aiUnit.qaChecklist) && aiUnit.qaChecklist.length ? aiUnit.qaChecklist : unit.qaChecklist,
  };
  updateLectureDerivedFields(merged, inputs, index, { preserveAiChecklist: true, preserveAiContent: true });
  return merged;
}

function mergeLectureAiSummary(unit, aiUnit, inputs, index) {
  if (!aiUnit) return unit;
  const merged = {
    ...unit,
    title: clean(aiUnit.title) || unit.title,
    subtopics: Array.isArray(aiUnit.subtopics) && aiUnit.subtopics.length ? aiUnit.subtopics.map(clean).filter(Boolean) : unit.subtopics,
    videoMinutes: Number(aiUnit.teachingMinutes) || unit.videoMinutes,
    pptSlides: Number(aiUnit.slideTarget) || unit.pptSlides,
    templateId: clean(aiUnit.templateId) || unit.templateId,
    outcomes: Array.isArray(aiUnit.outcomes) && aiUnit.outcomes.length ? aiUnit.outcomes : unit.outcomes,
    pptFocus: Array.isArray(aiUnit.pptFocus) && aiUnit.pptFocus.length ? aiUnit.pptFocus : unit.pptFocus,
    recordingCue: aiUnit.recordingCue || unit.recordingCue,
    duplicateCleanup: aiUnit.duplicateCleanup || unit.duplicateCleanup,
    qaChecklist: Array.isArray(aiUnit.qaChecklist) && aiUnit.qaChecklist.length ? aiUnit.qaChecklist : unit.qaChecklist,
  };
  merged.hours = roundOne(merged.videoMinutes / 60);
  updateLectureDerivedFields(merged, inputs, index, { preserveAiChecklist: true, preserveAiContent: true });
  return merged;
}

function buildLectureAiRefreshPayload(unit) {
  return {
    id: unit.id,
    number: unit.number,
    week: unit.week,
    title: unit.title,
    subtopics: unit.subtopics || [],
    teachingMinutes: unit.videoMinutes,
    slideTarget: unit.pptSlides,
    templateId: unit.templateId,
    moduleId: unit.metadata?.module_id || "",
    difficulty: unit.metadata?.difficulty || "",
    resourceProfile: unit.metadata?.resource_profile || [],
    officialAlignment: unit.metadata?.official_alignment || [],
    outcomes: unit.outcomes || [],
    pptFocus: unit.pptFocus || [],
    recordingCue: unit.recordingCue || "",
    duplicateCleanup: unit.duplicateCleanup || "",
    qaChecklist: unit.qaChecklist || [],
  };
}

function buildLectureAiInputs(inputs) {
  return {
    moduleTitle: inputs.moduleTitle,
    audience: inputs.audience,
    context: inputs.context,
    slidesPerHour: inputs.slidesPerHour,
    resourceConstraints: inputs.resourceConstraints || [],
    officialRefs: inputs.officialRefs || [],
  };
}

function getLectureChecklistRanges(slideTarget, chunkSize = 10) {
  const total = clamp(Number(slideTarget) || 10, 1, 80);
  const ranges = [];
  for (let start = 1; start <= total; start += chunkSize) {
    ranges.push({ startSlide: start, endSlide: Math.min(total, start + chunkSize - 1) });
  }
  return ranges;
}

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function getAiRetryCount(type) {
  const heavyAiJobs = new Set([
    "annual-plan",
    "lesson",
    "script",
    "script-revision",
    "assessment-bank",
    "lecture-pptx",
    "lecture-pptx-summary",
    "lecture-pptx-checklist",
    "lab-content",
    "assessment-content",
  ]);
  return heavyAiJobs.has(type) ? 2 : 1;
}

function isRetryableAiError(error) {
  return /failed to fetch|network|timeout|timed out|502|503|504|econnreset|socket/i.test(error?.message || "");
}

async function requestAiWithRetry(type, payload, retries = 1) {
  let lastError = null;
  for (let attempt = 0; attempt <= retries; attempt += 1) {
    try {
      return await requestAiOnce(type, payload);
    } catch (error) {
      lastError = error;
      const retryable = isRetryableAiError(error);
      if (!retryable || attempt === retries) break;
      await wait(900 * (attempt + 1));
    }
  }
  throw await enrichAiError(lastError || new Error("AI request failed after retry."), type);
}

async function enrichAiError(error, type) {
  const aiError = error instanceof Error ? error : new Error(String(error || "AI request failed."));
  if (aiError.aiDiagnosticsFormatted) return aiError;

  const diagnostics = aiError.diagnostics || await fetchAiDiagnostics(type, aiError.message);
  const formatted = formatAiDiagnostics(diagnostics);
  if (!formatted) return aiError;

  const enriched = new Error(`${aiError.message}\n${formatted}`);
  enriched.status = aiError.status;
  enriched.diagnostics = diagnostics;
  enriched.aiDiagnosticsFormatted = true;
  return enriched;
}

async function fetchAiDiagnostics(type, message) {
  if (window.location.protocol === "file:") return null;
  try {
    const params = new URLSearchParams({
      source: type || "",
      error: message || "",
    });
    const response = await fetch(`/api/ai/diagnostics?${params.toString()}`);
    if (!response.ok) return null;
    return response.json();
  } catch {
    return null;
  }
}

function formatAiDiagnostics(diagnostics) {
  if (!diagnostics || typeof diagnostics !== "object") return "";
  const checks = Array.isArray(diagnostics.checks) ? diagnostics.checks : [];
  const failedChecks = checks.filter((check) => !check.ok);
  const highlights = failedChecks.length
    ? failedChecks
    : checks.filter((check) => ["provider", "token policy", "last error"].includes(check.name));
  const parts = [];

  if (highlights.length) {
    parts.push(`AI diagnostics: ${highlights.map(formatAiDiagnosticCheck).join(" | ")}`);
  }
  if (diagnostics.probe) {
    const probe = diagnostics.probe;
    parts.push(`AI live probe: ${probe.ok ? "OK" : "FAIL"} ${probe.detail || ""}${probe.hint ? ` (${probe.hint})` : ""}`.trim());
  }
  const firstHint = checks.map((check) => check.hint).find(Boolean);
  if (firstHint) parts.push(`Fix hint: ${firstHint}`);

  return parts.join("\n");
}

function formatAiDiagnosticCheck(check) {
  const status = check.ok ? "OK" : "FAIL";
  const detail = check.detail ? `: ${check.detail}` : "";
  const hint = check.hint ? ` (${check.hint})` : "";
  return `${status} ${check.name}${detail}${hint}`;
}

function buildLectureTopics(customTopics, count, weeklyLectures = []) {
  const defaults = [
    "進階 Linux 銜接：systemd、networking、package、logs",
    "Container runtime、OCI、image 與 registry 基礎",
    "Kubernetes architecture：control plane、node、etcd、scheduler",
    "kubectl、YAML、namespace 與 cluster context",
    "Workloads：Pod、Deployment、ReplicaSet、Job、CronJob",
    "Service、Ingress、DNS 與 NetworkPolicy",
    "ConfigMap、Secret、Volume、PV/PVC 與 StorageClass",
    "Security：RBAC、ServiceAccount、SecurityContext",
    "Troubleshooting：logs、events、describe、resource pressure",
    "CKA 衝刺：cluster admin、upgrade、backup、networking drill",
    "CKAD 衝刺：application design、probes、config、observability",
    "AWS Academy EKS：managed control plane、node group、IAM integration",
    "Isakei、Rancher 與期末 Skill Test briefing",
  ];
  if (weeklyLectures.length) {
    return Array.from({ length: count }, (_, index) => weeklyLectures[index]?.title || defaults[index % defaults.length]);
  }
  const source = customTopics.length ? customTopics : defaults;
  return Array.from({ length: count }, (_, index) => source[index] || defaults[index % defaults.length]);
}

function buildLectureUnit(topic, index, count, hours, inputs, weeklyItem = null) {
  const week = weeklyItem?.week || Math.max(1, Math.round(((index + 1) / count) * inputs.weeks));
  const pptSlides = Math.max(10, Math.round(hours * inputs.slidesPerHour));
  const focus = inferLectureFocus(topic);
  const templateId = inferLectureTemplateId(topic, pptSlides);
  const resourceProfile = inferResourceProfile(topic, inputs.resourceConstraints);
  const officialAlignment = inferOfficialAlignment(topic, inputs.officialRefs);
  const subtopics = inferLectureSubtopics(topic, weeklyItem, focus);
  return {
    id: `L${index + 1}`,
    number: index + 1,
    week,
    title: topic,
    subtopics,
    hours: roundOne(hours),
    videoMinutes: Math.round(hours * 60),
    pptSlides,
    templateId,
    deckName: `Deck ${index + 1}: ${topic}`,
    focus,
    metadata: {
      course_id: slugifyFilename(inputs.moduleTitle || "course").replace(/\.pptx$/i, ""),
      course_title: inputs.moduleTitle,
      lecture_id: `L${index + 1}`,
      module_id: inferModuleId(topic, index),
      duration_minutes: Math.round(hours * 60),
      slide_target: pptSlides,
      template_id: templateId,
      difficulty: inferDifficulty(topic),
      prerequisites: index === 0 ? ["未指定"] : [`L${index}`],
      resource_profile: resourceProfile,
      official_alignment: officialAlignment,
      locale: "zh-TW",
      version: "0.1.0",
      model_name: state.ai?.model || "ai-required",
      prompt_hash: `ai-seed:${hashString(`${inputs.moduleTitle}|${topic}|${week}|${pptSlides}`)}`,
      source_snapshot_id: buildSourceSnapshotId(inputs.officialRefs),
      review_status: "draft",
    },
    slideSpec: buildLectureSlideSpec(topic, pptSlides, subtopics),
    pptxChecklist: buildPptxGenerationChecklist({
      topic,
      subtopics,
      slideTarget: pptSlides,
      minutes: Math.round(hours * 60),
      templateId,
      focus,
      resourceProfile,
      officialAlignment,
    }),
    qaChecklist: buildLectureQaChecklist(),
    outcomes: [
      `學生能說明 ${focus.core} 的用途與限制。`,
      `學生能以 kubectl / YAML 完成一個可驗收的 ${focus.task} 任務。`,
      "學生能分辨 CKA 管理員視角與 CKAD 開發者視角的差異。",
    ],
    pptFocus: [
      "概念模型與指令流程",
      "CLI demo checkpoint",
      "常見錯誤與 troubleshooting cue",
      "CKA/CKAD 對應題型",
    ],
    recordingCue: `${Math.round(hours * 60)} 分鐘影片，建議拆成 3 段：概念、demo、exam drill。`,
    duplicateCleanup: "若與前一 deck 重複，只保留 exam angle、demo 差異與 troubleshooting 變體。",
    sourceWeeklyItem: weeklyItem?.raw || "",
  };
}

function parseWeeklyList(lines, inputs) {
  return (lines || [])
    .map((line, index) => parseWeeklyLine(line, index, inputs))
    .filter(Boolean);
}

function parseWeeklyLine(line, index, inputs) {
  const raw = clean(line);
  if (!raw) return null;
  const weekMatch =
    raw.match(/(?:week|wk|w)\s*0?(\d{1,2})\b/i) ||
    raw.match(/第\s*0?(\d{1,2})\s*[週周]/) ||
    raw.match(/^0?(\d{1,2})[\s).、:：-]/);
  const week = weekMatch ? clamp(Number(weekMatch[1]), 1, inputs.weeks || 52) : clamp(index + 1, 1, inputs.weeks || 52);
  const title = raw
    .replace(/^(?:week|wk|w)\s*0?\d{1,2}\s*[-:：).、]?\s*/i, "")
    .replace(/^第\s*0?\d{1,2}\s*[週周]\s*[-:：).、]?\s*/, "")
    .replace(/^0?\d{1,2}[\s).、:：-]+/, "")
    .trim() || raw;
  const hoursMatch = raw.match(/(\d+(?:\.\d+)?)\s*(?:h|hr|hrs|hour|hours|小時)/i);
  const hours = hoursMatch ? Number(hoursMatch[1]) : null;
  const type = classifyWeeklyItem(title);
  return {
    raw,
    week,
    type,
    title: stripWeeklyTypePrefix(title),
    hours,
  };
}

function stripWeeklyTypePrefix(title) {
  return String(title || "")
    .replace(/^(lecture|ppt|lab|ca lab|assessment|quiz|exam|test|project|buffer)\s*\d*\s*[-:：).、]?\s*/i, "")
    .replace(/^(課堂|講課|實驗|評核|測驗|期末|專題|緩衝)\s*\d*\s*[-:：).、]?\s*/, "")
    .trim();
}

function classifyWeeklyItem(title) {
  const lower = String(title || "").toLowerCase();
  const hasAssessmentCue = /assessment|quiz|exam|test|skill test|checkpoint|rubric|評核|考核|測驗|筆試|期末|口試|驗收|ea\b|ca\b/.test(lower);
  const hasLabCue = /lab|實驗|hands-on|workshop/.test(lower);
  const startsAsLab = /^(ca\s*)?lab\b|^實驗/.test(lower);
  if (hasAssessmentCue && !startsAsLab) return "Assessment";
  if (hasLabCue) return "CA Lab";
  if (hasAssessmentCue) return "Assessment";
  if (/project|capstone|buffer|catch-up|revision|緩衝|專題|複習週/.test(lower)) return "Project / Buffer";
  return "Lecture/PPT";
}

function formatWeeklyItemAsSpec(item) {
  return `${item.type} ${item.week}：${item.title}`;
}

function inferLectureSubtopics(topic, weeklyItem, focus) {
  const raw = [weeklyItem?.raw, topic].filter(Boolean).join("\n");
  const candidates = raw
    .split(/\r?\n|[；;]|、|,(?=\s*\S)/)
    .map((item) => clean(item)
      .replace(/^(?:week|wk|w)\s*\d+\s*[-:：).、]?\s*/i, "")
      .replace(/^第\s*\d+\s*[週周]\s*[-:：).、]?\s*/, "")
      .replace(/^(lecture|ppt|topic|subtopic|大題目|子題目)\s*\d*\s*[-:：).、]?\s*/i, ""))
    .filter((item) => item && item !== topic);
  const unique = Array.from(new Set(candidates)).slice(0, 8);
  if (unique.length >= 3) return unique;
  return [
    `${focus.core} 概念框架`,
    `${focus.task} 操作流程`,
    "CLI / YAML demo walkthrough",
    "常見錯誤與 troubleshooting 判斷",
    "Lab bridge 與可驗收成果",
  ];
}

function normalizeLectureSubtopics(value, fallback = []) {
  const items = String(value || "")
    .split(/\r?\n|[；;]|、/)
    .map((item) => clean(item).replace(/^[-*]\s*/, ""))
    .filter(Boolean);
  return items.length ? Array.from(new Set(items)).slice(0, 12) : fallback;
}

function calculateLectureSlideTarget(minutes, slidesPerHour) {
  return clamp(Math.round((Number(minutes) || 60) / 60 * (Number(slidesPerHour) || 12)), 6, 80);
}

async function updateAnnualLectureFromCard(index) {
  const plan = state.annualPlan;
  const unit = plan?.lectureUnits?.[index];
  if (!plan || !unit) return;
  const titleInput = dom.annualLecturePlan.querySelector(`[data-lecture-title="${index}"]`);
  const subtopicsInput = dom.annualLecturePlan.querySelector(`[data-lecture-subtopics="${index}"]`);
  const minutesInput = dom.annualLecturePlan.querySelector(`[data-lecture-minutes="${index}"]`);
  const title = clean(titleInput?.value) || unit.title;
  const minutes = clamp(Number(minutesInput?.value) || unit.videoMinutes || 60, 10, 360);
  const subtopics = normalizeLectureSubtopics(subtopicsInput?.value, unit.subtopics || []);

  unit.title = title;
  unit.subtopics = subtopics;
  unit.videoMinutes = minutes;
  unit.hours = roundOne(minutes / 60);
  unit.pptSlides = calculateLectureSlideTarget(minutes, plan.inputs.slidesPerHour);
  updateLectureDerivedFields(unit, plan.inputs, index);
  unit.aiRefresh = {
    state: "running",
    message: `AI 正在根據「${unit.title}」與 ${unit.videoMinutes} 分鐘重新生成中間教學重點、PPT spec 與逐頁清單。`,
    updatedAt: new Date().toISOString(),
  };
  recalculateAnnualMetrics(plan);
  syncAnnualLectureTimetable(plan);
  renderAnnualPlan();
  persistState();

  setAiBusy(true, "AI 更新 Lecture/PPT 清單中");
  let summaryDone = false;
  try {
    const compactInputs = buildLectureAiInputs(plan.inputs);
    const summary = await requestAi("lecture-pptx-summary", {
      unit: buildLectureAiRefreshPayload(unit),
      inputs: compactInputs,
    });
    summaryDone = true;
    Object.assign(unit, mergeLectureAiSummary(unit, summary, plan.inputs, index));
    unit.aiRefresh = {
      state: "running",
      message: `AI 已更新中間教學重點，正在分段生成 ${unit.pptSlides} 頁專業 PPTX 清單。`,
      updatedAt: new Date().toISOString(),
    };
    renderAnnualPlan();
    persistState();

    const slideSpec = [];
    const pptxChecklist = [];
    const ranges = getLectureChecklistRanges(unit.pptSlides, 3);
    for (const range of ranges) {
      unit.aiRefresh = {
        state: "running",
        message: `AI 正在生成 PPTX 清單 ${range.startSlide}-${range.endSlide} / ${unit.pptSlides}。`,
        updatedAt: new Date().toISOString(),
      };
      renderAnnualPlan();
      const chunk = await requestAiWithRetry("lecture-pptx-checklist", {
        unit: buildLectureAiRefreshPayload(unit),
        inputs: compactInputs,
        ...range,
      }, 1);
      slideSpec.push(...(Array.isArray(chunk.slideSpec) ? chunk.slideSpec : []));
      pptxChecklist.push(...(Array.isArray(chunk.pptxChecklist) ? chunk.pptxChecklist : []));
    }
    if (!slideSpec.length || !pptxChecklist.length) {
      throw new Error("AI 未回傳逐頁 PPTX 清單。");
    }
    unit.slideSpec = slideSpec.sort((a, b) => Number(a.slide_no || 0) - Number(b.slide_no || 0));
    unit.pptxChecklist = pptxChecklist.sort((a, b) => Number(a.slide_no || 0) - Number(b.slide_no || 0));
    unit.aiRefresh = {
      state: "done",
      message: `AI 已更新：${unit.title} / ${unit.videoMinutes} 分鐘 / ${unit.pptSlides} slides`,
      updatedAt: new Date().toISOString(),
    };
  } catch (error) {
    console.warn(error);
    unit.aiRefresh = {
      state: summaryDone ? "partial" : "error",
      message: `AI Lecture/PPT 清單更新失敗：${error.message}`,
      updatedAt: new Date().toISOString(),
    };
    if (dom.compareBox) dom.compareBox.textContent = unit.aiRefresh.message;
    logAudit("Lecture/PPT 編輯", unit.aiRefresh.message);
  } finally {
    setAiBusy(false);
  }
  recalculateAnnualMetrics(plan);
  syncAnnualLectureTimetable(plan);
  if (unit.aiRefresh.state === "done" || unit.aiRefresh.state === "partial") {
    logAudit("Lecture/PPT 編輯", `AI 更新 ${unit.id}：${unit.title} / ${unit.videoMinutes} 分鐘 / ${unit.pptSlides} slides`);
  }
  renderAnnualPlan();
  markDriveBackupNeeded("AI Lecture/PPT 清單編輯");
  persistState();
}

function updateLectureDerivedFields(unit, inputs, index, options = {}) {
  const focus = inferLectureFocus(`${unit.title} ${(unit.subtopics || []).join(" ")}`);
  const templateId = inferLectureTemplateId(unit.title, unit.pptSlides);
  const resourceProfile = inferResourceProfile(`${unit.title} ${unit.subtopics.join(" ")}`, inputs.resourceConstraints);
  const officialAlignment = inferOfficialAlignment(`${unit.title} ${unit.subtopics.join(" ")}`, inputs.officialRefs);
  const firstSubtopic = unit.subtopics?.[0] || focus.task;
  const secondSubtopic = unit.subtopics?.[1] || focus.practice || focus.task;
  const localRecordingCue = `${unit.videoMinutes} 分鐘教學，建議拆成 ${Math.max(2, Math.ceil(unit.videoMinutes / 20))} 段：概念、demo、checkpoint、lab bridge。`;
  const localOutcomes = [
    `學生能說明 ${focus.core} 的用途、限制與適用情境。`,
    `學生能完成與「${firstSubtopic}」相關的可驗收實作。`,
    `學生能根據「${secondSubtopic}」的證據說明成功、失敗與修正方向。`,
  ];
  const localPptFocus = [
    `${unit.title} 的概念框架與工作流程`,
    `${firstSubtopic} 的逐步 demo / command walkthrough`,
    `${secondSubtopic} 的 checkpoint、錯誤辨識與 troubleshooting`,
    "Lab bridge、QA gate、accessibility 與版本 metadata",
  ];
  unit.focus = focus;
  unit.templateId = templateId;
  unit.deckName = `Deck ${unit.number || index + 1}: ${unit.title}`;
  if (!options.preserveAiContent || !unit.recordingCue) {
    unit.recordingCue = localRecordingCue;
  }
  if (!options.preserveAiContent || !Array.isArray(unit.outcomes) || !unit.outcomes.length) {
    unit.outcomes = localOutcomes;
  }
  if (!options.preserveAiContent || !Array.isArray(unit.pptFocus) || !unit.pptFocus.length) {
    unit.pptFocus = localPptFocus;
  }
  unit.metadata = {
    ...(unit.metadata || {}),
    course_id: slugifyFilename(inputs.moduleTitle || "course").replace(/\.pptx$/i, ""),
    course_title: inputs.moduleTitle,
    lecture_id: unit.id || `L${index + 1}`,
    module_id: inferModuleId(unit.title, index),
    duration_minutes: unit.videoMinutes,
    slide_target: unit.pptSlides,
    template_id: templateId,
    difficulty: inferDifficulty(unit.title),
    prerequisites: index === 0 ? ["未指定"] : [`L${index}`],
    resource_profile: resourceProfile,
    official_alignment: officialAlignment,
    locale: "zh-TW",
    version: unit.metadata?.version || "0.1.0",
    model_name: state.ai?.model || "AI-required",
    prompt_hash: `AI-seed:${hashString(`${inputs.moduleTitle}|${unit.title}|${unit.subtopics.join("|")}|${unit.videoMinutes}|${unit.pptSlides}`)}`,
    source_snapshot_id: buildSourceSnapshotId(inputs.officialRefs),
    review_status: "draft",
  };
  const localSlideSpec = buildLectureSlideSpec(unit.title, unit.pptSlides, unit.subtopics);
  const localPptxChecklist = buildPptxGenerationChecklist({
    topic: unit.title,
    subtopics: unit.subtopics,
    slideTarget: unit.pptSlides,
    minutes: unit.videoMinutes,
    templateId,
    focus,
    resourceProfile,
    officialAlignment,
  });
  if (!options.preserveAiChecklist || !Array.isArray(unit.slideSpec) || !unit.slideSpec.length) {
    unit.slideSpec = localSlideSpec;
  }
  if (!options.preserveAiChecklist || !Array.isArray(unit.pptxChecklist) || !unit.pptxChecklist.length) {
    unit.pptxChecklist = localPptxChecklist;
  }
  unit.qaChecklist = buildLectureQaChecklist();
}

function recalculateAnnualMetrics(plan) {
  const lectureHours = roundOne((plan.lectureUnits || []).reduce((sum, unit) => sum + Number(unit.hours || 0), 0));
  const pptSlides = (plan.lectureUnits || []).reduce((sum, unit) => sum + Number(unit.pptSlides || 0), 0);
  plan.metrics.lectureHours = lectureHours;
  plan.metrics.pptSlides = pptSlides;
  plan.metrics.lectureUnits = plan.lectureUnits?.length || 0;
  plan.metrics.recordingHours = lectureHours;
  plan.metrics.totalHours = roundOne(lectureHours + Number(plan.metrics.labHours || 0) + Number(plan.metrics.assessmentHours || 0));
  plan.inputs.lectureHours = lectureHours;
  if (dom.annualLectureHours) dom.annualLectureHours.value = String(lectureHours);
}

function syncAnnualLectureTimetable(plan) {
  (plan.lectureUnits || []).forEach((unit) => {
    const row = (plan.timetable || []).find((item) => item.type === "Lecture/PPT" && item.id === unit.id);
    if (!row) return;
    row.week = unit.week;
    row.title = unit.title;
    row.hours = unit.hours;
    row.output = `${unit.deckName} / ${unit.pptSlides} slides / ${unit.videoMinutes} min video`;
    row.dependency = unit.metadata?.prerequisites?.join("、") || "依週次清單排序";
  });
}

function buildTimetableFromWeeklyItems({ weeklyItems, lectureUnits, labs, assessments, inputs }) {
  const lectureQueue = [...lectureUnits];
  const labQueue = [...labs];
  const assessmentQueue = [...assessments];
  const rows = weeklyItems.map((item) => {
    if (item.type === "Lecture/PPT") {
      const unit = lectureQueue.shift();
      return {
        week: item.week,
        type: "Lecture/PPT",
        id: unit?.id || `W${item.week}`,
        title: unit?.title || item.title,
        hours: unit?.hours || roundOne(inputs.lectureHours / Math.max(1, lectureUnits.length)),
        output: unit ? `${unit.deckName} / ${unit.pptSlides} slides / ${unit.videoMinutes} min video` : "PPT slide spec + speaker notes",
        owner: "Lecturer",
        dependency: unit?.metadata?.prerequisites?.join("、") || "依週次清單排序",
      };
    }

    if (item.type === "CA Lab") {
      const lab = labQueue.shift();
      return {
        week: item.week,
        type: "CA Lab",
        id: lab?.id || `Lab W${item.week}`,
        title: lab?.title || item.title,
        hours: lab?.hours || roundOne(inputs.labHours / Math.max(1, labs.length)),
        output: lab?.outcome || "steps、resources、deliverables、rubric、test cases",
        owner: "Teacher / TA",
        dependency: lab?.sourceWeeklyItem || "對齊前置 lecture 與 lab bridge",
      };
    }

    if (item.type === "Assessment") {
      const assessment = assessmentQueue.shift();
      return {
        week: item.week,
        type: "Assessment",
        id: assessment?.type || `A-W${item.week}`,
        title: assessment?.title || item.title,
        hours: assessment?.hours || roundOne(inputs.assessmentHours / Math.max(1, assessments.length)),
        output: assessment?.deliverables?.join("、") || "goal、format、weight、rubric、evidence",
        owner: assessment?.type === "EA" ? "#Cyrus / Lecturer" : "Lecturer / TA",
        dependency: assessment?.sourceWeeklyItem || "QA gate 通過後發布",
      };
    }

    return {
      week: item.week,
      type: item.type,
      id: `W${item.week}`,
      title: item.title,
      hours: item.hours || 0,
      output: "Buffer / repair pass / capstone integration",
      owner: "Lecturer",
      dependency: "用於補課、QA 修復或專題整合",
    };
  });

  return rows.sort((a, b) => a.week - b.week || typeOrder(a.type) - typeOrder(b.type));
}

function buildProfessionalStandard(inputs, weeklyItems) {
  const resourceConstraints = Array.isArray(inputs.resourceConstraints) ? inputs.resourceConstraints : [];
  const officialRefs = Array.isArray(inputs.officialRefs) ? inputs.officialRefs : [];
  return {
    summary: weeklyItems.length
      ? `已讀取 ${weeklyItems.length} 條每週清單，並重整為 Lecture/PPT、CA Lab、Assessment 與 QA gate。`
      : "未提供每週清單；使用預設年度課程包骨架生成。",
    flow: PROFESSIONAL_COURSE_STANDARD.flow,
    slideTemplate: PROFESSIONAL_COURSE_STANDARD.slideSections,
    weeklyBreakdown: {
      lecture: weeklyItems.filter((item) => item.type === "Lecture/PPT").length,
      lab: weeklyItems.filter((item) => item.type === "CA Lab").length,
      assessment: weeklyItems.filter((item) => item.type === "Assessment").length,
      buffer: weeklyItems.filter((item) => item.type === "Project / Buffer").length,
    },
    resourceProfile: resourceConstraints.length ? resourceConstraints : ["未指定"],
    officialAlignment: officialRefs.length ? officialRefs : ["未指定"],
  };
}

function buildMetadataContract(inputs) {
  const officialRefs = Array.isArray(inputs.officialRefs) ? inputs.officialRefs : [];
  return PROFESSIONAL_COURSE_STANDARD.metadataFields.map((field) => ({
    field,
    required: true,
    source: field.includes("source") || field.includes("official") ? "官方來源 / snapshot" : "年度規劃生成器",
    status: field === "official_alignment" && !officialRefs.length ? "needs_input" : "draft",
  }));
}

function buildQaGates() {
  return PROFESSIONAL_COURSE_STANDARD.qaGates.map((name, index) => ({
    id: `QA-${String(index + 1).padStart(2, "0")}`,
    name,
    passRule: index < 2 ? "自動檢查" : "自動檢查 + 教師審核",
    blocking: index !== 0 ? "publish" : "export",
  }));
}

function buildAccessibilityChecklist() {
  return [
    "每張投影片必須有唯一標題",
    "圖表、流程圖與截圖需要描述性 alt text",
    "閱讀順序需與視覺順序一致",
    "一般文字對比至少 4.5:1，大字至少 3:1",
    "正文避免低於 18pt；每頁保留足夠留白",
    "Lecture 結尾要有 Lab bridge 或 assessment handoff",
  ];
}

function inferLectureTemplateId(topic, slides) {
  const lower = String(topic || "").toLowerCase();
  if (/eks|aws|cloud/.test(lower)) return `LT-CLOUD-${slides}`;
  if (/troubleshoot|debug|故障|排錯/.test(lower)) return `LT-TROUBLE-${slides}`;
  if (/review|skill|exam|cka|ckad|複習|衝刺/.test(lower)) return `LT-REVIEW-${slides}`;
  if (/demo|yaml|kubectl|ansible|helm|kustomize/.test(lower)) return `LT-DEMO-${slides}`;
  return `LT-CORE-${slides}`;
}

function inferModuleId(topic, index) {
  const lower = String(topic || "").toLowerCase();
  if (/linux|shell|systemd|bash/.test(lower)) return "M01-linux-foundation";
  if (/ansible|automation|git/.test(lower)) return "M02-automation";
  if (/container|oci|image|registry/.test(lower)) return "M03-container";
  if (/kubernetes|kubectl|workload|service|ingress|storage|rbac|cka|ckad/.test(lower)) return "M04-kubernetes";
  if (/eks|aws|rancher|isakei|cloud/.test(lower)) return "M05-cloud-capstone";
  return `M${String(index + 1).padStart(2, "0")}-course-module`;
}

function inferDifficulty(topic) {
  const lower = String(topic || "").toLowerCase();
  if (/intro|導論|入門|foundation|基礎/.test(lower)) return "beginner";
  if (/troubleshoot|security|rbac|eks|cka|ckad|skill/.test(lower)) return "advanced";
  return "intermediate";
}

function inferResourceProfile(topic, resourceConstraints = []) {
  const lower = String(topic || "").toLowerCase();
  const resources = [];
  if (/linux|shell|systemd|bash|vm/.test(lower)) resources.push("vm");
  if (/ansible/.test(lower)) resources.push("ansible");
  if (/kubernetes|kubectl|minikube|cka|ckad|workload|service|ingress|storage/.test(lower)) resources.push("minikube", "kubectl");
  if (/eks|aws/.test(lower)) resources.push("aws-eks");
  if (/rancher/.test(lower)) resources.push("rancher");
  resourceConstraints.forEach((item) => {
    const normalized = clean(item).toLowerCase();
    if (normalized && !resources.includes(normalized)) resources.push(clean(item));
  });
  return resources.length ? resources : ["未指定"];
}

function inferOfficialAlignment(topic, officialRefs = []) {
  const lower = String(topic || "").toLowerCase();
  const refs = officialRefs.filter((ref) => {
    const refLower = ref.toLowerCase();
    if (/cka|cluster|troubleshoot|storage|network/.test(lower) && refLower.includes("cka")) return true;
    if (/ckad|workload|deployment|job|probe|config/.test(lower) && refLower.includes("ckad")) return true;
    if (/eks|aws/.test(lower) && refLower.includes("eks")) return true;
    if (/ansible/.test(lower) && refLower.includes("ansible")) return true;
    if (/minikube|kubectl|kubernetes/.test(lower) && refLower.includes("minikube")) return true;
    return false;
  });
  return refs.length ? refs : officialRefs.length ? officialRefs.slice(0, 3) : ["未指定"];
}

function buildSourceSnapshotId(refs = []) {
  const stamp = new Date().toISOString().slice(0, 10);
  const signature = refs.length ? hashString(refs.join("|")) : "unspecified";
  return `src-${stamp}-${signature}`;
}

function buildLectureSlideSpec(topic, slideTarget, subtopics = []) {
  const sections = PROFESSIONAL_COURSE_STANDARD.slideSections;
  const count = clamp(Number(slideTarget) || sections.length, 6, 80);
  return Array.from({ length: count }, (_, index) => {
    const section = sections[index % sections.length];
    const subtopic = subtopics[index % Math.max(1, subtopics.length)] || topic;
    return {
      slide_no: index + 1,
      section,
      subtopic,
      purpose: `${topic}｜${subtopic}｜${section}`,
      renderer_hint: section.includes("Demo") ? "demo command / YAML walkthrough" : section.includes("錯誤") ? "misconception matrix" : "clear 16:9 teaching slide",
      required_notes: index >= 6 ? "speaker notes must include answer key, checkpoint, and fallback" : "speaker notes must connect to learning objective",
      repeatsEveryHour: count > sections.length,
    };
  });
}

function buildPptxGenerationChecklist({ topic, subtopics, slideTarget, minutes, templateId, focus, resourceProfile, officialAlignment }) {
  const resolvedFocus = focus || inferLectureFocus(topic);
  const slideSpec = buildLectureSlideSpec(topic, slideTarget, subtopics);
  const minutesPerSlide = Math.max(1, Math.round((Number(minutes) || 60) / Math.max(1, slideSpec.length)));
  return slideSpec.map((slide) => ({
    slide_no: slide.slide_no,
    title: `${slide.slide_no}. ${slide.section}`,
    big_topic: topic,
    subtopic: slide.subtopic,
    teaching_minutes: minutesPerSlide,
    template_id: templateId,
    visible_text: [
      `${slide.subtopic} 的核心概念`,
      `本頁要學生完成的判斷：${resolvedFocus.task}`,
      "成功證據：能說出原因、步驟與驗收方式",
    ],
    visual_direction: slide.section.includes("Demo")
      ? "左側命令 / YAML，右側預期輸出與檢查點"
      : slide.section.includes("錯誤")
        ? "常見錯誤對照表：症狀、原因、修正、驗收"
        : "流程圖、概念地圖或 checklist，避免純文字塞滿",
    speaker_notes: [
      `用 ${minutesPerSlide} 分鐘講清楚本頁與「${topic}」的關係。`,
      `先講情境，再講 ${slide.subtopic}，最後用一個問題檢查學生是否能應用。`,
      "保留 answer key、fallback demo path 與下一頁轉場。",
    ],
    lab_bridge: slide.slide_no === slideSpec.length
      ? "收束到下一個 Lab / Assessment deliverable，明確說出學生要交甚麼證據。"
      : "指出本頁如何支援後續 Lab 或 checkpoint。",
    qa_gate: [
      "唯一 slide title",
      "visible text 不超過 4 bullets",
      "speaker notes 包含答案鍵",
      "alt text / reading order 可檢查",
    ],
    resource_profile: resourceProfile,
    official_alignment: officialAlignment,
  }));
}

function buildLectureQaChecklist() {
  return [
    "slide_target 是否與 duration 對齊",
    "每頁是否有唯一標題與 teaching purpose",
    "Demo / checkpoint / lab bridge 是否齊全",
    "metadata 是否含版本、來源快照與 review_status",
    "speaker notes 是否包含答案鍵與 fallback",
  ];
}

function buildAnnualTimetable({ lectureUnits, labs, assessments, inputs }) {
  const rows = [];

  lectureUnits.forEach((unit) => {
    rows.push({
      week: unit.week,
      type: "Lecture/PPT",
      id: unit.id,
      title: unit.title,
      hours: unit.hours,
      output: `${unit.deckName} / ${unit.pptSlides} slides / ${unit.videoMinutes} min video`,
      owner: "Lecturer",
      dependency: "先完成 PPT prompt，再錄製影片",
    });
  });

  labs.forEach((lab, index) => {
    const week = Math.min(inputs.weeks, Math.max(2, Math.round(((index + 1) / (labs.length + 1)) * inputs.weeks)));
    lab.week = week;
    rows.push({
      week,
      type: "CA Lab",
      id: lab.id,
      title: lab.title,
      hours: lab.hours,
      output: lab.outcome,
      owner: "Teacher / TA",
      dependency: index === 0 ? "Lecture L1-L2" : `完成 ${labs[index - 1]?.id || "previous Lab"}`,
    });
  });

  assessments.forEach((assessment, index) => {
    const week = index === 0 ? Math.max(4, Math.round(inputs.weeks * 0.45)) : Math.max(6, Math.round(inputs.weeks * 0.9));
    assessment.week = week;
    rows.push({
      week,
      type: "Assessment",
      id: assessment.type,
      title: assessment.title,
      hours: assessment.hours,
      output: assessment.deliverables.join("、"),
      owner: assessment.type === "EA" ? "#Cyrus / Lecturer" : "Lecturer / TA",
      dependency: assessment.type === "EA" ? "所有核心 Lab + public endpoint 準備" : "Lab checkpoint 與筆試題庫審核",
    });
  });

  return rows.sort((a, b) => a.week - b.week || typeOrder(a.type) - typeOrder(b.type));
}

function typeOrder(type) {
  if (type === "Lecture/PPT") return 1;
  if (type === "CA Lab") return 2;
  return 3;
}

function inferLectureFocus(topic) {
  const lower = String(topic || "").toLowerCase();
  const hasKubernetes = /kubernetes|kubectl|minikube|cka|ckad|eks|rancher|pod|deployment|ingress|cluster|namespace|workload/.test(lower);
  if (/ckad/.test(lower)) return { core: "CKAD application workload", task: "workload deployment", practice: "YAML 與 workload 驗收" };
  if (/cka/.test(lower)) return { core: "CKA cluster administration", task: "cluster operation", practice: "cluster 故障排查" };
  if (/eks|aws/.test(lower)) return { core: "EKS managed Kubernetes", task: "cloud cluster", practice: "managed node / endpoint 驗收" };
  if (/rancher/.test(lower)) return { core: "Rancher enterprise management", task: "multi-cluster operation", practice: "cluster policy 與存取控制" };
  if (hasKubernetes && /security|rbac/.test(lower)) return { core: "Kubernetes security control", task: "RBAC policy", practice: "least privilege 驗證" };
  if (hasKubernetes && /service|ingress|network/.test(lower)) return { core: "Kubernetes networking", task: "public service exposure", practice: "Service / Ingress 驗收" };
  if (/system admin|sysadmin|system administration|systemd|daemon|service|process|journalctl|monitor|linux|shell|bash|package|logs/.test(lower)) {
    if (/process|monitor|top|ps|kill|signal/.test(lower) && /systemd|daemon|service/.test(lower)) {
      return { core: "Linux process and service administration", task: "process monitoring plus service lifecycle control", practice: "ps/top/systemctl/journalctl 證據鏈" };
    }
    if (/process|monitor|top|ps|kill|signal/.test(lower)) {
      return { core: "Linux process monitoring and control", task: "process inspection and signal handling", practice: "process 狀態、PID、signal 與資源使用證據" };
    }
    if (/systemd|daemon|service/.test(lower)) {
      return { core: "Linux service and daemon management", task: "service lifecycle control", practice: "systemctl / journalctl 驗收" };
    }
    return { core: "advanced Linux system administration", task: "server operation workflow", practice: "logs、services、packages 與網絡檢查" };
  }
  if (/security|rbac/.test(lower)) return { core: "security and access control", task: "policy design", practice: "權限與審核證據" };
  return { core: "course module concept", task: "hands-on task", practice: "可驗收成果與反思" };
}

function buildLabUnit(line, index, hours) {
  const title = line.replace(/^Lab\s*\d+\s*[：:]\s*/i, "").trim() || `Lab ${index + 1}`;
  const lower = title.toLowerCase();
  const lab = {
    id: `Lab ${index + 1}`,
    number: index + 1,
    title,
    hours: roundOne(hours),
    environment: "Local VM / Minikube",
    deliverables: ["實驗記錄", "截圖證據", "Git / YAML artifact", "短答反思"],
    rubric: ["可重現", "指令與 YAML 正確", "故障排查合理", "安全與資源設定清楚"],
  };

  if (lower.includes("ubuntu") || lower.includes("vm")) {
    lab.environment = "Ubuntu Server VM，4GB RAM minimum，8GB RAM preferred";
    lab.outcome = "學生完成 VM、SSH、套件、網絡與 baseline hardening。";
  } else if (lower.includes("ansible")) {
    lab.environment = "Ubuntu controller + target nodes";
    lab.outcome = "學生研究並提交 Ansible playbook，自動部署 Kubernetes prerequisites / cluster。";
  } else if (lower.includes("minikube") && (lower.includes("ckad") || lower.includes("ckd"))) {
    lab.environment = "Minikube + kubectl + YAML";
    lab.outcome = "學生完成 CKAD workload、config、probe、job、resource request/limit drill。";
  } else if (lower.includes("minikube") && lower.includes("cka")) {
    lab.environment = "Minikube single-node cluster";
    lab.outcome = "學生完成 CKA 類型的 cluster admin、troubleshooting、service exposure drill。";
  } else if (lower.includes("eks") || lower.includes("aws")) {
    lab.environment = "AWS Academy / EKS";
    lab.outcome = "學生完成 EKS cluster walkthrough、node group、IAM 與 public endpoint 驗收。";
  } else if (lower.includes("rancher") || lower.includes("isakei")) {
    lab.environment = "Isakei environment + optional Rancher";
    lab.outcome = "學生完成 Isakei 作業基線，延伸比較 Rancher enterprise Kubernetes management。";
  } else {
    lab.outcome = "學生完成可驗收的 Kubernetes hands-on artifact。";
  }
  return lab;
}

function buildAssessmentPlan(inputs, weeklyAssessments = []) {
  const caHours = roundOne(inputs.assessmentHours * 0.45);
  const eaHours = roundOne(inputs.assessmentHours * 0.55);
  const defaults = [
    {
      type: "CA",
      title: "Continuous Assessment：筆試 + Lab checkpoint",
      hours: caHours,
      weight: "建議 40%",
      deliverables: ["短測 / quiz", "Lab evidence pack", "YAML / command log", "oral check 或 reflection"],
      rules: ["題目需為原創或授權來源", "AI 生成題需教師審核答案與難度", "Lab checkpoint 必須可重現"],
    },
    {
      type: "EA",
      title: "End Assessment：#Cyrus Isakei 作業",
      hours: roundOne(eaHours / 2),
      weight: "建議 30%",
      deliverables: ["Isakei artifact", "部署說明", "公開 endpoint", "測試證據"],
      rules: ["由 #Cyrus 負責作業規格", "需列明 public endpoint 驗收 URL", "提交版本需有 timestamp"],
    },
    {
      type: "EA",
      title: "End Assessment：No-hint Skill Test",
      hours: roundOne(eaHours / 2),
      weight: "建議 30%",
      deliverables: ["現場技能操作", "public API / service endpoint", "故障排查記錄"],
      rules: ["no hint", "endpoint must be public", "rubric 預先公開但測試題不提示"],
    },
  ];
  weeklyAssessments.forEach((item, index) => {
    if (defaults[index]) {
      defaults[index].title = item.title;
      defaults[index].week = item.week;
      defaults[index].sourceWeeklyItem = item.raw;
      if (item.hours) defaults[index].hours = item.hours;
    } else {
      defaults.push({
        type: item.title.toLowerCase().includes("ea") || item.title.includes("期末") ? "EA" : "CA",
        title: item.title,
        week: item.week,
        hours: item.hours || roundOne(inputs.assessmentHours / Math.max(1, weeklyAssessments.length)),
        weight: "未指定",
        deliverables: ["assessment brief", "evidence pack", "rubric", "teacher review record"],
        rules: ["需定義 goal、format、weight_percent", "需通過 QA gate 才可發布", "需保存 source snapshot 與版本"],
        sourceWeeklyItem: item.raw,
      });
    }
  });
  return defaults;
}

function defaultLabSpecs() {
  return [
    "Lab 1：Ubuntu Server VM，4GB/8GB RAM 資源規格",
    "Lab 2：Ansible 自動部署 Kubernetes 叢集",
    "Lab 3：Minikube CKA 衝刺",
    "Lab 4：Minikube CKAD 衝刺，修正 CKD 筆誤",
    "Lab 5：AWS Academy EKS",
    "Lab 6：Isakei 實驗；Rancher optional extension",
  ];
}

function splitPlanLines(value) {
  return String(value || "")
    .split(/\r?\n/)
    .map((item) => clean(item))
    .filter(Boolean);
}

function splitPlanItems(value) {
  return String(value || "")
    .split(/\r?\n|、|；|;/)
    .map((item) => clean(item))
    .filter(Boolean);
}

async function generateLesson() {
  const inputs = getLessonInputs();
  state.interviewAnswers = inputs.interviewAnswers || "";
  state.lastLessonInputs = inputs;
  setAiBusy(true, "生成教材中");
  let generated = false;

  try {
    const aiLesson = await requestAi("lesson", { inputs });
    if (aiLesson?.slides?.length) {
      state.questions = Array.isArray(aiLesson.questions) ? aiLesson.questions : [];
      state.slides = normalizeAiSlides(aiLesson.slides, inputs);
      logAudit("教材生成", `${formatAiProviderName(state.ai.provider)} 生成 ${state.slides.length} 頁教材草稿`);
      generated = true;
    } else {
      throw new Error("AI 未回傳有效 slides。");
    }
  } catch (error) {
    console.warn(error);
    if (dom.compareBox) dom.compareBox.textContent = `AI 教材生成失敗：${error.message}`;
    logAudit("教材生成", `AI 生成失敗：${error.message}`);
  } finally {
    setAiBusy(false);
  }

  if (!generated) {
    renderStatus();
    persistState();
    return;
  }

  dom.duration.value = inputs.duration;
  annotateSlidesWithSourceRefs(inputs);
  dom.materialText.value = buildMaterialFromSlides();
  dom.assistantContext.value = buildAssistantContext();
  renderQuestions();
  renderAll();
  markDriveBackupNeeded("教材生成");
  persistState();
}

function normalizeAiSlides(slides, inputs) {
  const fallbackPlan = buildPptDeckPlan(inputs);
  const fallbackMinutes = distributeMinutes(inputs.duration, slides.map(() => 1 / Math.max(slides.length, 1)));
  return slides.map((slide, index) => {
    const fallback = fallbackPlan[index] || fallbackPlan[index % fallbackPlan.length];
    const slideType = clean(slide.slideType || slide.type) || fallback?.slideType || "content";
    const template = pptTemplateCatalog[slideType] || pptTemplateCatalog.content;
    const bloomKey = inferBloomKey(slide.bloom, inputs, template, index);
    const bloom = bloomMap[bloomKey] || bloomMap.understand;
    const title = clean(slide.title) || fallback?.title || buildPptSlideTitle(slideType, inputs, index);
    const minutes = Number(slide.minutes) || fallback?.minutes || fallbackMinutes[index] || 5;
    const activity = clean(slide.activity) || buildPptTemplateActivity(slideType, inputs, bloom);
    const suggestedLayout = clean(slide.suggestedLayout) || template.layout;
    const suggestedVisual = clean(slide.suggestedVisual) || template.visual || inferPptVisualPrompt(title, inputs.topic, template.event);
    const factCheckPoints = Array.isArray(slide.factCheckPoints) && slide.factCheckPoints.length
      ? slide.factCheckPoints
      : buildPptFactCheckPoints(slideType, inputs);
    const speakerNotes = clean(slide.speakerNotes) || buildPptSpeakerNotes(slideType, inputs, title, activity);
    return {
      id: cryptoId(),
      number: index + 1,
      title,
      event: clean(slide.event) || template.event,
      bloom: clean(slide.bloom) || bloom.label,
      bloomKey,
      minutes,
      activity,
      slideType,
      suggestedLayout,
      suggestedVisual,
      speakerNotes,
      factCheckPoints,
      notes: buildPptSlidePrompt({
        title,
        event: clean(slide.event) || template.event,
        inputs,
        bloom,
        minutes,
        activity,
        sourceNotes: clean(slide.notes) || buildPptSlideSourceNotes(slideType, inputs, title, activity),
        slideType,
        template,
        suggestedLayout,
        suggestedVisual,
        factCheckPoints,
        speakerNotes,
      }),
    };
  });
}

function buildSlideFromDeckPlan(plan, inputs, index) {
  const template = pptTemplateCatalog[plan.slideType] || pptTemplateCatalog.content;
  const bloomKey = inferBloomKey(template.bloom, inputs, template, index);
  const bloom = bloomMap[bloomKey] || bloomMap.understand;
  const title = plan.title || buildPptSlideTitle(plan.slideType, inputs, index);
  const activity = buildPptTemplateActivity(plan.slideType, inputs, bloom);
  const factCheckPoints = buildPptFactCheckPoints(plan.slideType, inputs);
  const speakerNotes = buildPptSpeakerNotes(plan.slideType, inputs, title, activity);
  const suggestedVisual = plan.suggestedVisual || template.visual || inferPptVisualPrompt(title, inputs.topic, template.event);
  return {
    id: cryptoId(),
    number: index + 1,
    title,
    event: template.event,
    bloom: bloom.label,
    bloomKey,
    minutes: plan.minutes,
    activity,
    slideType: plan.slideType,
    suggestedLayout: plan.suggestedLayout || template.layout,
    suggestedVisual,
    speakerNotes,
    factCheckPoints,
    notes: buildPptSlidePrompt({
      title,
      event: template.event,
      inputs,
      bloom,
      minutes: plan.minutes,
      activity,
      sourceNotes: buildPptSlideSourceNotes(plan.slideType, inputs, title, activity),
      slideType: plan.slideType,
      template,
      suggestedLayout: plan.suggestedLayout || template.layout,
      suggestedVisual,
      factCheckPoints,
      speakerNotes,
    }),
  };
}

function buildPptDeckPlan(inputs) {
  const examOrTech = isExamOrTechnicalCourse(inputs);
  const selected = new Set(inputs.bloom || []);
  const longDeck = inputs.duration >= 50 || examOrTech;
  const sequence = longDeck
    ? ["title", "prerequisite", "objectives", "agenda", "content", "content", "example", "demo", "comparison", "pitfalls", "exercise", "assessment", "summary", "references"]
    : ["title", "objectives", "agenda", "content", "example", selected.has("apply") ? "demo" : "content", selected.has("apply") ? "exercise" : "assessment", "summary"];

  if (!sequence.includes("comparison") && (selected.has("analyze") || selected.has("evaluate"))) {
    sequence.splice(Math.max(4, sequence.length - 2), 0, "comparison");
  }
  if (!sequence.includes("assessment") && (selected.has("evaluate") || examOrTech)) {
    sequence.splice(Math.max(5, sequence.length - 1), 0, "assessment");
  }

  const weights = sequence.map((type) => pptTemplateCatalog[type]?.weight || 0.08);
  const minutes = distributeMinutes(inputs.duration, normalizeWeights(weights));
  return sequence.map((slideType, index) => {
    const template = pptTemplateCatalog[slideType] || pptTemplateCatalog.content;
    return {
      slideType,
      title: buildPptSlideTitle(slideType, inputs, index),
      minutes: minutes[index],
      suggestedLayout: template.layout,
      suggestedVisual: template.visual,
    };
  });
}

function normalizeWeights(weights) {
  const total = weights.reduce((sum, value) => sum + Number(value || 0), 0) || 1;
  return weights.map((value) => Number(value || 0) / total);
}

function isExamOrTechnicalCourse(inputs) {
  const text = `${inputs.topic} ${inputs.subject} ${inputs.context} ${inputs.objective} ${inputs.interviewAnswers}`.toLowerCase();
  return /(cka|ckad|kubernetes|kubectl|yaml|eks|linux|ansible|minikube|rancher|exam|assessment|skill test|lab)/i.test(text);
}

function inferBloomKey(value, inputs, template, index) {
  const normalized = String(value || "").toLowerCase();
  const byLabel = Object.keys(bloomMap).find((key) => bloomMap[key].label === value);
  if (byLabel) return byLabel;
  const direct = Object.keys(bloomMap).find((key) => normalized.includes(key));
  if (direct) return direct;
  if (inputs.bloom?.includes(template.bloom)) return template.bloom;
  return inputs.bloom?.[index % inputs.bloom.length] || template.bloom || "understand";
}

function buildPptSlideTitle(slideType, inputs, index) {
  const titles = {
    title: inputs.topic,
    prerequisite: "從先備能力到本課任務",
    objectives: "本節完成後，你應能",
    agenda: `${inputs.duration} 分鐘學習路線`,
    content: index < 5 ? `${inputs.topic} 的用途與限制` : "核心概念與操作邊界",
    example: "Worked Example：從需求到可驗收結果",
    demo: "Demo：用 YAML 與 kubectl 交付任務",
    exercise: "課內練習：完成一個可驗收任務",
    comparison: "工具相同，責任不同",
    pitfalls: "常見錯誤與 Troubleshooting 切點",
    assessment: "Mini Assessment：用考試視角驗收",
    summary: "三個 Takeaways 與下一步",
    references: "Official References 與版本查核",
  };
  return titles[slideType] || buildSlideTitle(pptTemplateCatalog[slideType]?.event || "呈現內容", inputs.topic, index);
}

function buildPptTemplateActivity(slideType, inputs, bloom) {
  const actions = {
    title: `用一句話說出你期待這堂 ${inputs.topic} 解決的實務問題。`,
    prerequisite: "快速標記自己已掌握、需要補強、今天暫不深入的能力。",
    objectives: `把每個 learning objective 轉成可交付或可驗收的 ${bloom.verb} 任務。`,
    agenda: "用時間軸確認本課的講解、demo、練習與評量位置。",
    content: `用「定義、用途、限制、常見誤區」拆解 ${inputs.topic}。`,
    example: "讀一個 worked example，指出情境、做法、結果與誤解。",
    demo: "跟著命令與 YAML 觀察 expected output，完成驗收檢查點。",
    exercise: "學生在限定時間內完成任務、交出 evidence 與驗收結果。",
    comparison: "用角色視角比較 CKA / CKAD 或管理員 / 開發者的成功標準。",
    pitfalls: "把錯誤現象對應到第一個要查的 command 或 evidence。",
    assessment: "完成 3 題 mini assessment 或 1 個 performance task。",
    summary: "用 exit ticket 寫出今天最重要的一個操作判斷。",
    references: "標記課後要查的 official source 與版本更新位置。",
  };
  return actions[slideType] || buildActivity(pptTemplateCatalog[slideType]?.event || "呈現內容", inputs, bloom);
}

function buildPptSlideSourceNotes(slideType, inputs, title, activity) {
  const objectives = splitPlanItems(inputs.objective).slice(0, 4);
  return [
    `course_json.title: ${inputs.topic}`,
    `course_json.subject_domain: ${inputs.subject}`,
    `course_json.audience_profile: ${inputs.audience}`,
    `course_json.duration_min: ${inputs.duration}`,
    `course_json.style: ${inputs.style}`,
    `course_json.objectives: ${objectives.length ? objectives.join("；") : inputs.objective || "需由教師確認"}`,
    `course_json.prerequisites: ${inputs.context || "未提供；請在 speaker notes 標示假設"}`,
    `course_json.teacher_interview_answers: ${inputs.interviewAnswers || "尚未提供"}`,
    `slide_goal: ${title}`,
    `linked_assessment: ${activity}`,
  ].join("\n");
}

function buildPptSpeakerNotes(slideType, inputs, title, activity) {
  return [
    `講者備忘稿重點：先說明「${title}」如何服務本課 learning objective，再連到可驗收任務。`,
    `轉場語：這頁不是孤立概念，下一步要把它變成學生能操作、能驗收、能排錯的行動。`,
    `互動提示：${activity}`,
    inputs.interviewAnswers ? `教師補充已納入：${inputs.interviewAnswers}` : "若班級背景不明，請以先備橋接頁做快速診斷。",
  ].join("\n");
}

function buildPptFactCheckPoints(slideType, inputs) {
  const points = [];
  const text = `${inputs.topic} ${inputs.subject} ${inputs.objective}`.toLowerCase();
  if (/(cka|ckad)/i.test(text)) points.push("CKA / CKAD 考試定位、版本與 domain 以 Linux Foundation 官方頁為準。");
  if (/(kubernetes|kubectl|yaml|pod|service|ingress)/i.test(text)) points.push("kubectl、YAML 與 Kubernetes resource 術語以 Kubernetes 官方文件為準。");
  if (/(eks|aws)/i.test(text)) points.push("EKS、IAM、node group 與 public endpoint 細節以 AWS 官方文件為準。");
  if (/(ansible)/i.test(text)) points.push("Ansible inventory、managed nodes 與 playbook 用語以 Ansible 官方文件為準。");
  if (["assessment", "exercise", "demo", "pitfalls"].includes(slideType)) points.push("驗收條件必須可觀察、可重做、可截圖或可用 command 證明。");
  return points.length ? points : ["如涉及版本、考試權重或校內政策，請教師在發布前查核。"];
}

function buildPptInterviewQuestions(inputs) {
  const base = [
    `這份 deck 最終要學生交出甚麼 evidence：截圖、YAML、kubectl output、反思，還是 LMS quiz？`,
    `哪些先備知識只是快速橋接，哪些必須在 PPT 內重新教？`,
    `這堂「${inputs.topic}」最重要的 assessment touchpoint 是 demo、exercise 還是 mini test？`,
    "有沒有必須使用的 official source、考試版本、工具版本或校本 rubric？",
    "你想把 references 放在最後一頁，還是放到 appendix / speaker notes？",
  ];
  return base.slice(0, 5);
}

function annotateSlidesWithSourceRefs(inputs) {
  if (!state.materialPages.length) return;
  state.slides = state.slides.map((slide) => {
    const keywords = extractKeywords(`${inputs.topic} ${slide.title} ${slide.activity} ${slide.notes}`);
    const refs = state.materialPages
      .map((page, index) => ({
        type: "material",
        number: page.number || index + 1,
        title: page.title || `教材片段 ${index + 1}`,
        score: scoreMaterialPage(page, keywords, index),
      }))
      .filter((ref) => ref.score > 0)
      .sort((a, b) => b.score - a.score)
      .slice(0, 2)
      .map(({ score, ...ref }) => ref);
    return { ...slide, sourceRefs: refs };
  });
}

function renderQuestions() {
  const inputs = getLessonInputs();
  const questions = state.questions.length
    ? state.questions
    : [
    `這堂「${inputs.topic}」結束後，學生要交出哪一種可觀察成果？`,
    `你最擔心學生在哪個先備概念上出錯？`,
    `本課需要更偏向考試答題、實驗探究，還是生活應用？`,
    `如果只能保留 ${Math.max(3, Math.round(inputs.duration * 0.2))} 分鐘互動，你會想放在課堂前段、中段還是收尾？`,
    `有沒有校本課綱、108 課綱核心素養或 LMS 評量欄位需要對齊？`,
  ];

  dom.questionList.innerHTML = questions
    .map(
      (question, index) => `
        <div class="question-item">
          <span>Q${index + 1}</span>
          <strong>${escapeHtml(question)}</strong>
        </div>
      `,
    )
    .join("");

  if (dom.questionAnswer && dom.questionAnswer.value !== state.interviewAnswers) {
    dom.questionAnswer.value = state.interviewAnswers || "";
  }
  if (dom.questionAnswerStatus) {
    dom.questionAnswerStatus.textContent = state.interviewAnswers
      ? "已記錄你的補充；按「根據回答再生成」會把答案納入教材與 PPT Prompt。"
      : "回答上面的追問後，可再生成更貼近你班級、評核與教材目標的版本。";
  }
}

async function regenerateLessonFromInterviewAnswers() {
  const answers = clean(dom.questionAnswer.value);
  if (!answers) {
    dom.questionAnswerStatus.textContent = "請先在回答欄輸入你的補充，例如學生背景、評核重點、想避開或加強的內容。";
    return;
  }
  state.interviewAnswers = answers;
  dom.questionAnswerStatus.textContent = "正在根據你的回答重新生成教材...";
  logAudit("AI 追問回答", answers.slice(0, 140));
  await generateLesson();
  dom.questionAnswerStatus.textContent = "已根據你的回答重新生成教材。";
}

function renderAll() {
  renderAnnualPlan();
  renderTimeline();
  renderSlides();
  renderTimeBudget();
  renderScriptGoal();
  renderScript();
  renderChat();
  renderPublishedQa();
  renderVersions();
  renderAuditLog();
  renderGovernanceMetrics();
  renderDrivePanel();
  renderGammaPanel();
  renderStatus();
  renderAiStatus();
  applyRolePermissions();
}

function getLectureRefreshState(unit) {
  return unit?.aiRefresh && ["running", "done", "partial", "error"].includes(unit.aiRefresh.state) ? unit.aiRefresh : null;
}

function renderLectureAiRefreshNotice(unit) {
  const refresh = getLectureRefreshState(unit);
  if (!refresh) return "";
  const label = refresh.state === "running" ? "AI 生成中" : refresh.state === "done" ? "AI 已更新" : refresh.state === "partial" ? "AI 部分完成" : "AI 生成失敗";
  return `<div class="lecture-ai-notice ${escapeHtml(refresh.state)}"><strong>${label}</strong><span>${escapeHtml(refresh.message || "")}</span></div>`;
}

function getLectureRecordingCue(unit) {
  const refresh = getLectureRefreshState(unit);
  if (refresh?.state === "running") {
    return "AI 正在重新生成教學段落、學習成果、PPT spec 與 QA gate。";
  }
  if (refresh?.state === "error") {
    return refresh.message || "AI 未能完成更新，請檢查 API key / Secrets / server 狀態。";
  }
  return unit.recordingCue || "AI 尚未生成 recording cue。";
}

function renderLectureOutcomeList(unit) {
  const refresh = getLectureRefreshState(unit);
  if (refresh?.state === "running") {
    return "<ul><li>AI 正在根據新的大題目、子題目與教學時間重新生成學習成果。</li></ul>";
  }
  if (refresh?.state === "error") {
    return "<ul><li>AI 未完成更新；請確認 Codespaces Secrets 已授權此 repo，並重啟 server 後再按「更新清單」。</li></ul>";
  }
  const outcomes = Array.isArray(unit.outcomes) && unit.outcomes.length ? unit.outcomes : ["AI 尚未回傳學習成果。"];
  return `<ul>${outcomes.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`;
}

function getLecturePptSpecSummary(unit) {
  const refresh = getLectureRefreshState(unit);
  if (refresh?.state === "running") return "PPT spec：AI 正在重寫逐頁 purpose / visible text / speaker notes...";
  if (refresh?.state === "partial") return "PPT spec：AI 已更新中間教學重點，但逐頁 PPTX 清單未完整完成。";
  if (refresh?.state === "error") return "PPT spec：AI 未完成，暫停使用舊清單，避免輸出錯誤主題。";
  return `PPT spec：${(unit.slideSpec || []).slice(0, 4).map((item) => item.section).join(" → ")}${unit.slideSpec?.length > 4 ? "..." : ""}`;
}

function getLectureQaSummary(unit) {
  const refresh = getLectureRefreshState(unit);
  if (refresh?.state === "running") return "QA：AI 正在對齊 duration、slide target、teaching purpose 與 lab bridge。";
  if (refresh?.state === "partial") return "QA：中間教學重點已更新；請重試以補齊逐頁 PPTX 清單。";
  if (refresh?.state === "error") return "QA：請先修復 AI 連線後重新生成。";
  return `QA：${(unit.qaChecklist || []).slice(0, 3).join("；")}`;
}

function renderAnnualPlan() {
  if (!dom.annualMetrics) return;
  const plan = state.annualPlan;
  if (!plan) {
    dom.annualMetrics.innerHTML = emptyText("按「生成全年規劃」建立 Lecture、Lab 與評核藍圖");
    dom.annualNote.textContent = "建議先確認 lecture 小時、Lab 小時與評核小時，再生成全年課程包。";
    dom.annualProfessionalStandard.innerHTML = emptyText("貼入現有每週清單後，這裡會顯示專業結構、metadata contract 與 QA gate。");
    dom.annualLectureStatus.textContent = "尚未生成";
    dom.annualLecturePlan.innerHTML = emptyText("尚未生成 Lecture / PPT 清單");
    dom.annualTimetableStatus.textContent = "尚未生成";
    dom.annualTimetable.innerHTML = emptyText("尚未生成 Timetable");
    dom.annualLabPlan.innerHTML = emptyText("尚未生成 CA Lab Series");
    dom.annualAssessmentPlan.innerHTML = emptyText("尚未生成 Assessment 規劃");
    dom.annualContentTitle.textContent = "Lab / Assessment 內容生成區";
    dom.annualContentOutput.value = "";
    return;
  }

  const metrics = plan.metrics;
  dom.annualMetrics.innerHTML = [
    annualMetric("總小時", `${metrics.totalHours}h`, "Lecture + Lab + Assessment"),
    annualMetric("Lecture", `${metrics.lectureHours}h`, `${metrics.lectureUnits} 個錄影 / PPT batch`),
    annualMetric("PPT 頁數", `${metrics.pptSlides}`, "估算，可按每小時頁數調整"),
    annualMetric("CA Lab", `${metrics.labCount}`, `${metrics.labHours}h hands-on`),
    annualMetric("Assessments", `${metrics.assessmentCount}`, `${metrics.assessmentHours}h planning / testing`),
    annualMetric("影片錄製", `${metrics.recordingHours}h`, "建議逐 deck 錄製"),
  ].join("");

  dom.annualNote.innerHTML = `
    <strong>${escapeHtml(plan.inputs.moduleTitle)}</strong>
    <span>${escapeHtml(plan.inputs.context || "全年課程包已建立。")}</span>
    <small>${escapeHtml(plan.pptConsolidation.join(" "))}</small>
  `;
  dom.annualProfessionalStandard.innerHTML = renderProfessionalStandard(plan);
  dom.annualLectureStatus.textContent = `${metrics.lectureUnits} decks / ${metrics.recordingHours}h video`;
  dom.annualTimetableStatus.textContent = `${plan.timetable?.length || 0} items across ${plan.inputs.weeks} weeks`;
  const canEditAnnual = ["teacher", "admin"].includes(state.role);

  dom.annualLecturePlan.innerHTML = plan.lectureUnits.map((unit, index) => `
    <article class="annual-card">
      <div class="annual-card-head">
        <div>
          <span>${escapeHtml(unit.id)}｜Week ${unit.week}｜${unit.hours}h｜${unit.pptSlides} slides</span>
          <strong>${escapeHtml(unit.title)}</strong>
        </div>
        <div class="card-actions">
          <button class="action-button ghost" type="button" data-lecture-refresh="${index}" ${canEditAnnual ? "" : "disabled"}>更新清單</button>
          <button class="action-button ghost" type="button" data-annual-lecture="${index}" ${canEditAnnual ? "" : "disabled"}>送到 PPT 流程</button>
        </div>
      </div>
      <div class="lecture-edit-grid">
        <label>
          大題目
          <input type="text" data-lecture-title="${index}" value="${escapeHtml(unit.title)}" ${canEditAnnual ? "" : "disabled"} />
        </label>
        <label>
          教學時間（分鐘）
          <input type="number" min="10" max="360" step="5" data-lecture-minutes="${index}" value="${escapeHtml(unit.videoMinutes || Math.round((unit.hours || 1) * 60))}" ${canEditAnnual ? "" : "disabled"} />
        </label>
        <label class="wide-field">
          子題目
          <textarea rows="3" data-lecture-subtopics="${index}" ${canEditAnnual ? "" : "disabled"}>${escapeHtml((unit.subtopics?.length ? unit.subtopics : inferLectureSubtopics(unit.title, null, unit.focus || inferLectureFocus(unit.title))).join("\n"))}</textarea>
        </label>
      </div>
      ${renderLectureAiRefreshNotice(unit)}
      <p>${escapeHtml(getLectureRecordingCue(unit))}</p>
      <div class="spec-chip-row">
        <span>${escapeHtml(unit.templateId || "LT-CORE")}</span>
        <span>${escapeHtml(unit.metadata?.module_id || "module")}</span>
        <span>${escapeHtml(unit.metadata?.difficulty || "intermediate")}</span>
        <span>${escapeHtml((unit.metadata?.resource_profile || []).slice(0, 3).join(" / "))}</span>
      </div>
      ${renderLectureOutcomeList(unit)}
      <small>${escapeHtml(getLecturePptSpecSummary(unit))}</small>
      <small>${escapeHtml(getLectureQaSummary(unit))}</small>
      <small>${escapeHtml(unit.duplicateCleanup)}</small>
      ${renderLecturePptxChecklist(unit)}
    </article>
  `).join("");

  dom.annualTimetable.innerHTML = renderAnnualTimetable(plan.timetable || []);

  dom.annualLabPlan.innerHTML = plan.labs.map((lab, index) => `
    <article class="annual-card">
      <div class="annual-card-head">
        <div>
          <span>${escapeHtml(lab.id)}｜Week ${lab.week || "-"}｜${lab.hours}h｜${escapeHtml(lab.environment)}</span>
          <strong>${escapeHtml(lab.title)}</strong>
        </div>
        <button class="action-button ghost" type="button" data-lab-content="${index}" ${canEditAnnual ? "" : "disabled"}>生成內容</button>
      </div>
      <p>${escapeHtml(lab.outcome)}</p>
      <small>交付：${escapeHtml(lab.deliverables.join("、"))}</small>
      ${lab.generatedContent ? "<small>狀態：已生成學生版、教師版、evidence pack 與 rubric</small>" : ""}
    </article>
  `).join("");

  dom.annualAssessmentPlan.innerHTML = plan.assessments.map((assessment, index) => `
    <article class="annual-card">
      <div class="annual-card-head">
        <div>
          <span>${escapeHtml(assessment.type)}｜Week ${assessment.week || "-"}｜${assessment.hours}h｜${escapeHtml(assessment.weight)}</span>
          <strong>${escapeHtml(assessment.title)}</strong>
        </div>
        <button class="action-button ghost" type="button" data-assessment-content="${index}" ${canEditAnnual ? "" : "disabled"}>生成內容</button>
      </div>
      <p>交付：${escapeHtml(assessment.deliverables.join("、"))}</p>
      <small>${escapeHtml(assessment.rules.join("；"))}</small>
      ${assessment.generatedContent ? "<small>狀態：已生成題庫、答案鍵、scenario task 與 rubric</small>" : ""}
    </article>
  `).join("");

  dom.annualContentTitle.textContent = plan.generatedContent?.title || "Lab / Assessment 內容生成區";
  dom.annualContentOutput.value = plan.generatedContent?.markdown || "";

  dom.annualLecturePlan.querySelectorAll("[data-annual-lecture]").forEach((button) => {
    button.addEventListener("click", () => sendAnnualLectureToBuilder(Number(button.dataset.annualLecture)));
  });
  dom.annualLecturePlan.querySelectorAll("[data-lecture-refresh]").forEach((button) => {
    button.addEventListener("click", () => updateAnnualLectureFromCard(Number(button.dataset.lectureRefresh)));
  });
  dom.annualLabPlan.querySelectorAll("[data-lab-content]").forEach((button) => {
    button.addEventListener("click", () => generateLabContent(Number(button.dataset.labContent)));
  });
  dom.annualAssessmentPlan.querySelectorAll("[data-assessment-content]").forEach((button) => {
    button.addEventListener("click", () => generateAssessmentContent(Number(button.dataset.assessmentContent)));
  });
}

function annualMetric(label, value, hint) {
  return `<div><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong><small>${escapeHtml(hint)}</small></div>`;
}

function renderLecturePptxChecklist(unit) {
  const checklist = getLecturePptxChecklist(unit);
  return `
    <details class="pptx-checklist">
      <summary>專業 PPTX 生成清單｜${escapeHtml(checklist.length)} slides</summary>
      <div class="pptx-checklist-grid">
        ${checklist.map((slide) => `
          <section class="pptx-slide-row">
            <div>
              <span>S${escapeHtml(slide.slide_no)}｜${escapeHtml(slide.teaching_minutes)} min｜${escapeHtml(slide.template_id)}</span>
              <strong>${escapeHtml(slide.title)}</strong>
              <small>大題目：${escapeHtml(slide.big_topic)}</small>
              <small>子題目：${escapeHtml(slide.subtopic)}</small>
            </div>
            <div>
              <span>Visible Text</span>
              <ul>${(slide.visible_text || []).map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>
            </div>
            <div>
              <span>Visual / Notes / QA</span>
              <p>${escapeHtml(slide.visual_direction)}</p>
              <small>${escapeHtml((slide.speaker_notes || []).join(" "))}</small>
              <small>QA：${escapeHtml((slide.qa_gate || []).join("；"))}</small>
            </div>
          </section>
        `).join("")}
      </div>
    </details>
  `;
}

function getLecturePptxChecklist(unit) {
  if (Array.isArray(unit.pptxChecklist) && unit.pptxChecklist.length === Number(unit.pptSlides || unit.pptxChecklist.length)) {
    return unit.pptxChecklist;
  }
  const focus = unit.focus || inferLectureFocus(unit.title);
  const subtopics = unit.subtopics?.length ? unit.subtopics : inferLectureSubtopics(unit.title, null, focus);
  return buildPptxGenerationChecklist({
    topic: unit.title,
    subtopics,
    slideTarget: unit.pptSlides || unit.slideSpec?.length || 10,
    minutes: unit.videoMinutes || Math.round((unit.hours || 1) * 60),
    templateId: unit.templateId || unit.metadata?.template_id || "LT-CORE",
    focus,
    resourceProfile: unit.metadata?.resource_profile || ["未指定"],
    officialAlignment: unit.metadata?.official_alignment || ["未指定"],
  });
}

function renderProfessionalStandard(plan) {
  const standard = plan.professionalStandard || buildProfessionalStandard(plan.inputs || {}, plan.weeklyItems || []);
  const weekly = standard.weeklyBreakdown || {};
  const metadata = (plan.metadataContract || []).slice(0, 8);
  const qaGates = plan.qaGates || [];
  return `
    <div class="professional-summary">
      <strong>${escapeHtml(standard.summary)}</strong>
      <span>Lecture ${weekly.lecture || 0} / Lab ${weekly.lab || 0} / Assessment ${weekly.assessment || 0} / Buffer ${weekly.buffer || 0}</span>
    </div>
    <div class="professional-grid">
      <section>
        <span>Pipeline</span>
        <ul>${(standard.flow || []).slice(0, 6).map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>
      </section>
      <section>
        <span>Metadata Contract</span>
        <ul>${metadata.map((item) => `<li>${escapeHtml(item.field)}｜${escapeHtml(item.status)}</li>`).join("")}</ul>
      </section>
      <section>
        <span>QA Gate</span>
        <ul>${qaGates.slice(0, 6).map((item) => `<li>${escapeHtml(item.id)}｜${escapeHtml(item.name)}</li>`).join("")}</ul>
      </section>
      <section>
        <span>Accessibility</span>
        <ul>${(plan.accessibilityChecklist || []).slice(0, 6).map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>
      </section>
    </div>
  `;
}

function renderAnnualTimetable(rows) {
  if (!rows.length) return emptyText("尚未生成 Timetable");
  return `
    <div class="timetable-head">
      <span>Week</span>
      <span>Type</span>
      <span>Item</span>
      <span>Hours</span>
      <span>Output / Dependency</span>
    </div>
    ${rows
      .map(
        (row) => `
          <div class="timetable-row">
            <strong>W${escapeHtml(row.week)}</strong>
            <span class="type-chip ${typeChipClass(row.type)}">${escapeHtml(row.type)}</span>
            <div>
              <strong>${escapeHtml(row.id)}｜${escapeHtml(row.title)}</strong>
              <small>${escapeHtml(row.owner)}</small>
            </div>
            <span>${escapeHtml(row.hours)}h</span>
            <div>
              <span>${escapeHtml(row.output)}</span>
              <small>${escapeHtml(row.dependency)}</small>
            </div>
          </div>
        `,
      )
      .join("")}
  `;
}

function typeChipClass(type) {
  if (type === "CA Lab") return "lab";
  if (type === "Assessment") return "assessment";
  if (type === "Project / Buffer") return "buffer";
  return "lecture";
}

async function generateLabContent(index) {
  const plan = state.annualPlan;
  const lab = plan?.labs?.[index];
  if (!plan || !lab) return;
  setAiBusy(true, "AI 生成 Lab 內容中");
  try {
    const result = await requestAi("lab-content", { lab, plan, index });
    const markdown = result?.markdown;
    if (!markdown) throw new Error("AI 未回傳 Lab markdown。");
    plan.generatedContent = {
      title: result.title || `${lab.id}｜${lab.title}`,
      markdown,
      type: "lab",
      index,
      updatedAt: new Date().toISOString(),
    };
    lab.generatedContent = markdown;
    logAudit("Lab 內容生成", `AI 已生成 ${lab.id} instructions / steps / rubric`);
    renderAnnualPlan();
    markDriveBackupNeeded("AI Lab 內容生成");
    persistState();
  } catch (error) {
    console.warn(error);
    dom.annualContentOutput.value = `AI Lab 內容生成失敗：${error.message}`;
  } finally {
    setAiBusy(false);
  }
}

async function generateAssessmentContent(index) {
  const plan = state.annualPlan;
  const assessment = plan?.assessments?.[index];
  if (!plan || !assessment) return;
  setAiBusy(true, "AI 生成 Assessment 內容中");
  try {
    const result = await requestAi("assessment-content", { assessment, plan, index });
    const markdown = result?.markdown;
    if (!markdown) throw new Error("AI 未回傳 Assessment markdown。");
    plan.generatedContent = {
      title: result.title || `${assessment.type}｜${assessment.title}`,
      markdown,
      type: "assessment",
      index,
      updatedAt: new Date().toISOString(),
    };
    assessment.generatedContent = markdown;
    logAudit("Assessment 內容生成", `AI 已生成 ${assessment.type} assessment brief / rubric`);
    renderAnnualPlan();
    markDriveBackupNeeded("AI Assessment 內容生成");
    persistState();
  } catch (error) {
    console.warn(error);
    dom.annualContentOutput.value = `AI Assessment 內容生成失敗：${error.message}`;
  } finally {
    setAiBusy(false);
  }
}

async function generateAllLabContent() {
  if (!state.annualPlan) await generateAnnualPlan();
  const plan = state.annualPlan;
  if (!plan?.labs?.length) return;
  setAiBusy(true, "AI 批量生成全部 Lab 中");
  try {
    const results = [];
    for (let index = 0; index < plan.labs.length; index += 1) {
      const lab = plan.labs[index];
      const result = await requestAi("lab-content", { lab, plan, index });
      if (!result?.markdown) throw new Error(`${lab.id} 未回傳 markdown。`);
      lab.generatedContent = result.markdown;
      results.push(result.markdown);
    }
    const markdown = results.join("\n\n---\n\n");
    plan.generatedContent = {
      title: "全部 CA Lab Series｜AI 完整學生版 + 教師版",
      markdown,
      type: "lab-batch",
      index: null,
      updatedAt: new Date().toISOString(),
    };
    logAudit("Lab 批量生成", `AI 已生成 ${plan.labs.length} 個 Lab instructions / evidence / rubric`);
    renderAnnualPlan();
    markDriveBackupNeeded("AI Lab 批量生成");
    persistState();
  } catch (error) {
    console.warn(error);
    dom.annualContentOutput.value = `AI Lab 批量生成失敗：${error.message}`;
  } finally {
    setAiBusy(false);
  }
}

async function generateAssessmentBank() {
  if (!state.annualPlan) await generateAnnualPlan();
  const plan = state.annualPlan;
  if (!plan?.assessments?.length) return;
  setAiBusy(true, "AI 生成 Assessment 題庫中");
  try {
    const result = await requestAi("assessment-bank", { plan });
    const markdown = result?.markdown;
    if (!markdown) throw new Error("AI 未回傳 Assessment bank markdown。");
    state.assessmentBank = {
      id: cryptoId(),
      title: result.title || `${plan.inputs.moduleTitle}｜Assessment 題庫與 Rubric`,
      markdown,
      generatedAt: new Date().toISOString(),
    };
    plan.assessmentBank = structuredCloneSafe(state.assessmentBank);
    if (Array.isArray(result.assessmentContents)) {
      result.assessmentContents.forEach((item, index) => {
        if (plan.assessments[index] && item?.markdown) {
          plan.assessments[index].generatedContent = item.markdown;
        }
      });
    }
    plan.generatedContent = {
      title: state.assessmentBank.title,
      markdown,
      type: "assessment-bank",
      index: null,
      updatedAt: new Date().toISOString(),
    };
    logAudit("Assessment 題庫", `AI 已生成 ${plan.assessments.length} 個評核項題庫、答案鍵與 Rubric`);
    renderAnnualPlan();
    markDriveBackupNeeded("AI Assessment 題庫生成");
    persistState();
  } catch (error) {
    console.warn(error);
    dom.annualContentOutput.value = `AI Assessment 題庫生成失敗：${error.message}`;
  } finally {
    setAiBusy(false);
  }
}

async function copyAnnualGeneratedContent() {
  const text = dom.annualContentOutput?.value || "";
  if (!text) return;
  try {
    await navigator.clipboard.writeText(text);
    logAudit("複製", "Lab / Assessment 生成內容已複製");
  } catch {
    // Clipboard can be unavailable in some browser contexts.
  }
}

function buildLabContentMarkdown(lab, plan, index) {
  const prevLecture = plan.lectureUnits
    .filter((unit) => unit.week <= (lab.week || plan.inputs.weeks))
    .slice(-1)[0] || plan.lectureUnits[index] || plan.lectureUnits[0];
  const checklist = inferLabChecklist(lab);
  const expectedOutputs = inferLabExpectedOutputs(lab);
  const troubleshooting = inferLabTroubleshooting(lab);
  const answerKey = inferLabAnswerKey(lab);
  const evidenceRows = [
    ["Environment", "VM / cluster / cloud setting screenshot", "能重現學生使用的環境"],
    ["Artifact", "YAML / playbook / command log / Git link", "證明不是只靠截圖交功課"],
    ["Verification", "kubectl output / endpoint response / logs", "證明任務真的運作"],
    ["Reflection", "80-120 字錯誤與修正記錄", "證明學生知道自己做了什麼"],
  ];
  return `# ${lab.id}: ${lab.title}

## Timetable

- Week: ${lab.week || "-"}
- Hours: ${lab.hours}
- Environment: ${lab.environment}
- Related lecture: ${prevLecture?.id || "N/A"} ${prevLecture?.title || ""}

## Lab Profile

- Mode: guided hands-on + evidence-based assessment
- Student version: 可以直接派給學生
- Teacher version: 下面包含 answer key、common errors 與 marking rubric
- Safety note: 若使用 cloud / public endpoint，必須在提交前列明 cleanup 方法與費用風險

## Learning Outcomes

1. 完成可重現的 hands-on artifact，並能解釋每個主要指令或 YAML 欄位的用途。
2. 以截圖、command log、YAML / playbook 證明結果。
3. 用 80-120 字反思一個錯誤、排查過程與修正方法。

## Student Brief

你需要根據課堂內容完成 ${lab.title}。提交內容必須足夠讓老師或助教在另一台環境重現結果。

## Step-by-step Tasks

${checklist.map((item, itemIndex) => `${itemIndex + 1}. ${item}`).join("\n")}

## Expected Output

${expectedOutputs.map((item) => `- ${item}`).join("\n")}

## Deliverables

${lab.deliverables.map((item) => `- ${item}`).join("\n")}

## Evidence Pack Table

| Evidence | What to submit | Why it matters |
| --- | --- | --- |
${evidenceRows.map((row) => `| ${row[0]} | ${row[1]} | ${row[2]} |`).join("\n")}

## Acceptance Criteria

- 指令、YAML 或 playbook 可以重新執行。
- 截圖能清楚顯示 cluster / workload / endpoint 狀態。
- 若有 public endpoint，必須列出 URL、測試方法與預期 response。
- 反思能指出至少一個實際遇到的 error 或 limitation。

## Marking Rubric

| Criteria | Weight | Evidence |
| --- | ---: | --- |
| Correctness | 30% | artifact 能完成 lab objective |
| Reproducibility | 25% | 另一台環境可根據步驟重做 |
| Troubleshooting evidence | 20% | 有 command、logs、events 或 endpoint 檢查 |
| Explanation quality | 15% | 能解釋核心指令 / YAML / cloud setting |
| Safety and cleanup | 10% | 有資源、權限、費用或 cleanup 意識 |

## Common Errors And Hints

${troubleshooting.map((item) => `- ${item}`).join("\n")}

## Teacher Answer Key / Checking Notes

${answerKey.map((item) => `- ${item}`).join("\n")}

## Extension / Stretch Task

- 把本 Lab 的成果改成一個 no-hint mini task，讓同學交換環境重做。
- 把 evidence pack 整理成 1 頁 PDF 或 README，模擬真實 workplace handover。

## Teacher Notes

- 先檢查學生是否理解資源限制，特別是 VM RAM、CPU、storage。
- 不直接給完整答案；只提示如何閱讀 error、events、logs。
- 若學生使用 AI 生成指令，必須要求他解釋每個 flag / field。
- 教師發布前需試跑一次所有驗收命令，並確認不含未授權題庫或外部答案。
`;
}

function inferLabChecklist(lab) {
  const text = `${lab.title} ${lab.environment} ${lab.outcome}`.toLowerCase();
  if (text.includes("ubuntu") || text.includes("vm")) {
    return [
      "建立 Ubuntu Server VM，記錄 CPU/RAM/disk 設定，建議 4GB minimum / 8GB preferred。",
      "設定 SSH、package update、hostname、time sync 與基本 firewall policy。",
      "安裝 Kubernetes 後續 lab 需要的 prerequisite tools。",
      "提交 VM readiness screenshot 與 resource usage evidence。",
    ];
  }
  if (text.includes("ansible")) {
    return [
      "建立 inventory，清楚標示 control node 與 target nodes。",
      "撰寫 playbook 安裝 prerequisites、container runtime、kubectl/kubeadm 或指定工具。",
      "執行 playbook 並保留 idempotency 證據。",
      "提交 playbook、inventory、執行 log 與 troubleshooting notes。",
    ];
  }
  if (text.includes("ckad")) {
    return [
      "以 YAML 建立 workload、config、secret、probe、resource request/limit。",
      "使用 kubectl 驗證 rollout、logs、describe 與 events。",
      "完成至少一題 job / cronjob 或 multi-container pattern。",
      "整理 CKAD 題型對應表與常見錯誤。",
    ];
  }
  if (text.includes("cka")) {
    return [
      "建立 Minikube cluster 並驗證 node、namespace、storage 或 networking 狀態。",
      "完成 service exposure、troubleshooting、resource inspection drill。",
      "用 kubectl explain / describe / logs 查找問題。",
      "提交 command log 與 no-hint 解題反思。",
    ];
  }
  if (text.includes("eks") || text.includes("aws")) {
    return [
      "使用 AWS Academy learning path 建立或 walkthrough EKS 相關資源。",
      "識別 managed control plane、node group、IAM integration 的角色。",
      "部署 sample workload 並公開 service endpoint。",
      "提交 endpoint 測試、成本/安全注意事項與 cleanup record。",
    ];
  }
  return [
    "準備指定環境並記錄版本。",
    "完成核心 Kubernetes artifact。",
    "驗證狀態、endpoint 或 logs。",
    "提交 evidence pack 與短答反思。",
  ];
}

function inferLabExpectedOutputs(lab) {
  const text = `${lab.title} ${lab.environment} ${lab.outcome}`.toLowerCase();
  if (text.includes("ubuntu") || text.includes("vm")) {
    return [
      "`hostnamectl`、`free -h`、`df -h`、`ip addr` 或等效截圖能證明 VM ready。",
      "SSH 可以連線，package update 完成，並列出 CPU/RAM/disk 設定。",
      "學生能說明 4GB minimum / 8GB preferred 對 Kubernetes lab 的影響。",
    ];
  }
  if (text.includes("ansible")) {
    return [
      "Inventory 可清楚分辨 control node 與 managed nodes。",
      "Playbook 至少第二次執行能顯示 idempotent 結果，沒有大量 unintended changes。",
      "提交 playbook、執行 log、failed task 排查記錄與修正版本。",
    ];
  }
  if (text.includes("ckad")) {
    return [
      "`kubectl get` / `describe` / `logs` 可證明 workload 正常。",
      "YAML 包含正確 image、ports、env/config、probe 或 resource setting。",
      "學生能指出 CKAD 視角下的 application acceptance criteria。",
    ];
  }
  if (text.includes("cka")) {
    return [
      "`kubectl get nodes`、namespace、service 或 storage 狀態能被驗證。",
      "Troubleshooting 記錄包含 describe/events/logs 或 node condition。",
      "學生能指出 CKA 視角下 cluster administrator 要負責的 evidence。",
    ];
  }
  if (text.includes("eks") || text.includes("aws")) {
    return [
      "能說明 managed control plane、node group、IAM 或 service exposure 的角色。",
      "Public endpoint 可由評核員測試，並附 expected response。",
      "提交 cleanup record，避免雲端資源持續收費。",
    ];
  }
  return [
    "核心 artifact 可被重新執行或重新檢查。",
    "狀態、log、endpoint 或截圖能證明任務完成。",
    "反思能說明一個錯誤、原因與修正方法。",
  ];
}

function inferLabTroubleshooting(lab) {
  const text = `${lab.title} ${lab.environment} ${lab.outcome}`.toLowerCase();
  const base = [
    "先讀 error message，再決定查 Linux、container runtime、Kubernetes API 還是 application layer。",
    "不要只交最後成功截圖；保留至少一段排查證據。",
  ];
  if (text.includes("minikube") || text.includes("kubernetes") || text.includes("cka") || text.includes("ckad")) {
    return [
      ...base,
      "Pod Pending：先查 resource request、node capacity、PVC 或 image pull。",
      "Service 無法訪問：先查 selector、port/targetPort、endpoint 與 firewall / ingress。",
      "YAML apply 失敗：先用 `kubectl explain` 或官方 docs 查欄位，不要猜欄位名。",
    ];
  }
  if (text.includes("ansible")) {
    return [
      ...base,
      "SSH failed：先查 inventory host、user、key、sudo 權限與 network。",
      "Task changed every run：檢查 module 是否 idempotent，避免用 shell 硬寫不可重入命令。",
    ];
  }
  if (text.includes("eks") || text.includes("aws")) {
    return [
      ...base,
      "Endpoint 不通：先查 security group、service type、load balancer、subnet 與 app health。",
      "權限失敗：先查 IAM role、policy、AWS Academy lab limitation 與 region。",
    ];
  }
  return base;
}

function inferLabAnswerKey(lab) {
  const text = `${lab.title} ${lab.environment} ${lab.outcome}`.toLowerCase();
  if (text.includes("ubuntu") || text.includes("vm")) {
    return [
      "Accept：VM specs、SSH、package baseline、network evidence 全部齊全。",
      "Reject / resubmit：只有桌面截圖，沒有 command output 或 SSH evidence。",
      "教師抽查：請學生解釋為什麼 Server 版比 Desktop 版更適合節省資源。",
    ];
  }
  if (text.includes("ansible")) {
    return [
      "Accept：inventory + playbook + log + second run evidence，且學生能解釋 idempotency。",
      "Reject / resubmit：只提交網上 playbook，不能說明每個 task 目的。",
      "教師抽查：改一個 target host，要求學生說明 playbook 如何調整。",
    ];
  }
  if (text.includes("ckad")) {
    return [
      "Accept：workload 可運作，YAML 可重用，學生能解釋 probes/config/resource setting。",
      "Reject / resubmit：只用 imperative command 建立，沒有提交 YAML 或驗收證據。",
      "教師抽查：要求學生用同一 YAML 改 namespace 或 image tag。",
    ];
  }
  if (text.includes("cka")) {
    return [
      "Accept：cluster/service/troubleshooting evidence 清楚，排查路徑合理。",
      "Reject / resubmit：只看到最終狀態，沒有 events/logs/describe 等排查記錄。",
      "教師抽查：提供一個 NotReady / Pending 現象，請學生說出第一個檢查命令。",
    ];
  }
  return [
    "Accept：artifact、驗收 evidence、反思三者一致。",
    "Reject / resubmit：結果不可重現或無法解釋主要步驟。",
    "教師抽查：要求學生用自己的話說明最重要的一個判斷。",
  ];
}

function buildAssessmentContentMarkdown(assessment, plan) {
  const isEa = assessment.type === "EA";
  const questionSet = buildAssessmentQuestionSet(assessment, plan);
  const rubric = buildAssessmentRubricRows(assessment);
  return `# ${assessment.type}: ${assessment.title}

## Timetable

- Week: ${assessment.week || "-"}
- Hours: ${assessment.hours}
- Weight: ${assessment.weight}
- Related course: ${plan.inputs.moduleTitle}

## Assessment Brief

${isEa
    ? "學生需要在 no-hint 條件下完成技能實作測驗。所有 API / Service endpoint 必須公開，並能由評核員直接驗收。"
    : "學生需要完成筆試與 Lab checkpoint，證明自己能理解核心概念並完成可重現的 hands-on artifact。"}

## Assessment Blueprint

| Area | Coverage | Evidence |
| --- | --- | --- |
| Linux prerequisite | VM readiness、SSH、logs、network、resource awareness | command output / short answer |
| Kubernetes concept | control plane、node、resource model、desired state | MC / short answer |
| kubectl / YAML | workload、service、config、troubleshooting | practical task / YAML artifact |
| Cloud / endpoint | EKS / public service / IAM awareness | endpoint response / explanation |
| Reflection | error diagnosis、limitation、cleanup | written reflection / oral check |

## Student Deliverables

${assessment.deliverables.map((item) => `- ${item}`).join("\n")}

## Task Design

${isEa
    ? [
        "1. 提供一個未完整部署的 Kubernetes scenario。",
        "2. 學生需自行判斷 required resources、YAML、service exposure 與 troubleshooting path。",
        "3. 不提供 hint；只公布時間、交付格式、rubric 與 endpoint requirement。",
        "4. 評核員以 public endpoint、kubectl evidence、logs 與 oral check 驗收。",
      ].join("\n")
    : [
        "1. 筆試題目覆蓋 Linux prerequisite、Kubernetes architecture、kubectl / YAML、workload、networking、storage、security。",
        "2. Lab checkpoint 以 evidence pack 評分，不只看截圖，也看可重現性。",
        "3. 題目可由 AI 生成草稿，但教師必須審核答案、難度與授課覆蓋度。",
      ].join("\n")}

## Question Bank

### MC / Concept Check

${questionSet.mcq.map((item, itemIndex) => `${itemIndex + 1}. ${item.question}\n   - A. ${item.options[0]}\n   - B. ${item.options[1]}\n   - C. ${item.options[2]}\n   - D. ${item.options[3]}\n   - Answer: ${item.answer}\n   - Rationale: ${item.rationale}`).join("\n\n")}

### Short Answer

${questionSet.short.map((item, itemIndex) => `${itemIndex + 1}. ${item.question}\n   - Marking points: ${item.points.join("；")}`).join("\n\n")}

### Scenario / Practical Task

${questionSet.practical.map((item, itemIndex) => `${itemIndex + 1}. ${item.task}\n   - Student evidence: ${item.evidence.join("、")}\n   - Teacher check: ${item.check}`).join("\n\n")}

## Rubric

| Criteria | Weight | Excellent | Pass | Resubmit |
| --- | ---: | --- | --- | --- |
${rubric.map((row) => `| ${row.criteria} | ${row.weight} | ${row.excellent} | ${row.pass} | ${row.resubmit} |`).join("\n")}

## Rules

${assessment.rules.map((item) => `- ${item}`).join("\n")}

## Sample Feedback Phrases

- High performance: 你的 evidence 足夠完整，而且能清楚解釋為什麼這個 output 代表任務完成。
- Borderline pass: 結果大致正確，但排查紀錄不足；請補上 command log 或 endpoint 驗收。
- Resubmit: 目前未能證明結果可重現，請補交 artifact、驗收方法和錯誤修正記錄。

## Teacher Checklist

- 確認題目沒有依賴未授課或未授權材料。
- 確認評分標準在評核前公開。
- 確認答案與驗收命令已由教師試跑。
- 保存版本、rubric、sample answer 與 moderation notes。
- 不使用未授權商業題庫；AI 生成題目必須由教師審核與改寫。
`;
}

function buildAssessmentBankMarkdown(plan) {
  const sections = plan.assessments
    .map((assessment, index) => buildAssessmentContentMarkdown(assessment, plan, index))
    .join("\n\n---\n\n");
  return `# ${plan.inputs.moduleTitle}｜Assessment 題庫與 Rubric

生成時間：${new Date().toLocaleString("zh-Hant")}
對象：${plan.inputs.audience}

## 使用原則

- 題庫只作教師出題草稿，不直接公開完整答案。
- 所有題目需由教師審核授課覆蓋度、難度、答案與語言清晰度。
- 不使用未授權題庫；AI 生成題必須改寫成校本情境。
- Practical / Skill Test 可公開 rubric，但不公開提示與完整解題步驟。

${sections}
`;
}

function buildAssessmentQuestionSet(assessment, plan) {
  const text = `${assessment.title} ${plan.inputs.moduleTitle} ${plan.inputs.context}`.toLowerCase();
  const endpointFocus = /eks|aws|endpoint|skill test|isakei/.test(text);
  const ckaFocus = /cka|cluster|troubleshooting|admin/.test(text);
  const ckadFocus = /ckad|yaml|workload|deployment|pod/.test(text);
  return {
    mcq: [
      {
        question: "在 Kubernetes troubleshooting 中，哪一項 evidence 最能幫助判斷 Pod 為何未能正常啟動？",
        options: ["只看最終截圖", "`kubectl describe` 的 events 與 `kubectl logs`", "只提交 YAML 檔名", "只說明自己已重開機"],
        answer: "B",
        rationale: "events 和 logs 能指出 scheduling、image、runtime 或 application 層錯誤。",
      },
      {
        question: endpointFocus
          ? "若 service endpoint 必須公開，學生最少應提交哪一類證據？"
          : "CKA 與 CKAD 題型最核心的差異是什麼？",
        options: endpointFocus
          ? ["桌面截圖", "URL、測試方法、expected response 與安全/cleanup 說明", "只提交課堂筆記", "只交 Git commit message"]
          : ["CKA 側重 cluster/admin，CKAD 側重 application/workload", "CKA 不需要 kubectl", "CKAD 不需要 YAML", "兩者完全相同"],
        answer: "B",
        rationale: endpointFocus
          ? "公開 endpoint 需要可由評核員重複驗收。"
          : "兩者都用 Kubernetes，但成功標準和責任範圍不同。",
      },
      {
        question: ckadFocus
          ? "為什麼 assessment 要求提交 YAML artifact，而不只提交成功截圖？"
          : "為什麼 Linux prerequisite 會影響 Kubernetes lab 成功率？",
        options: ckadFocus
          ? ["方便重現和審核 desired state", "YAML 只是裝飾", "截圖比 artifact 更可重現", "因為不用解釋指令"]
          : ["因為 node、runtime、network、storage 問題會令 cluster 或 workload 失敗", "Linux 與 Kubernetes 無關", "只要有瀏覽器即可", "只要學生記得答案即可"],
        answer: "A",
        rationale: ckadFocus
          ? "YAML 是 Kubernetes desired state 的核心證據。"
          : "Kubernetes 依賴穩定的 Linux 與 network 基礎。",
      },
    ],
    short: [
      {
        question: "請列出你會如何驗收一個 Kubernetes practical task 已完成。",
        points: ["artifact 可重現", "kubectl output 或 endpoint response 清楚", "學生能解釋主要欄位或指令", "有錯誤排查或限制說明"],
      },
      {
        question: ckaFocus
          ? "Node NotReady 時，你會先查哪三類證據？"
          : "提交 Lab evidence pack 時，哪些內容不能只靠截圖代替？",
        points: ckaFocus
          ? ["node condition", "kubelet / container runtime 狀態", "events、logs、network 或 resource pressure"]
          : ["YAML / playbook artifact", "command log", "驗收方法", "錯誤與修正反思"],
      },
      {
        question: "如何判斷一題 practical assessment 是否過度依賴提示？",
        points: ["題目是否已暴露解題路徑", "rubric 是否只公開成功標準", "學生是否需要自行選擇 command / resource", "是否能用 evidence 驗收而非跟步驟打勾"],
      },
    ],
    practical: [
      {
        task: endpointFocus
          ? "部署一個 sample workload，公開 service endpoint，並提交 URL、expected response 與 cleanup record。"
          : "根據指定情境建立或修復一個 Kubernetes resource，並提交 YAML、kubectl output 與排查記錄。",
        evidence: ["YAML / command log", "kubectl get/describe/logs", "endpoint 或狀態截圖", "短答解釋"],
        check: "評核員需要能在另一台環境或同一 cluster 重新驗收結果。",
      },
      {
        task: "教師提供一個常見錯誤情境，學生需在 no-hint 條件下提出第一個檢查命令與修正方向。",
        evidence: ["錯誤現象", "第一個檢查命令", "修正假設", "最終驗收 output"],
        check: "重點不是背答案，而是排查路徑是否合理。",
      },
    ],
  };
}

function buildAssessmentRubricRows(assessment) {
  const isEa = assessment.type === "EA";
  return [
    {
      criteria: "Correctness",
      weight: isEa ? "35%" : "30%",
      excellent: "結果完全符合 task，且能解釋關鍵設定",
      pass: "主要結果正確，少量說明不足",
      resubmit: "結果不可驗收或與題目要求不符",
    },
    {
      criteria: "Reproducibility",
      weight: isEa ? "20%" : "25%",
      excellent: "artifact、步驟與環境資料可重做",
      pass: "大致可重做，但缺少部分環境或版本資料",
      resubmit: "只有截圖，不能重現",
    },
    {
      criteria: "Troubleshooting evidence",
      weight: "20%",
      excellent: "有清楚 error、假設、command evidence 與修正",
      pass: "有部分排查紀錄，但推理不完整",
      resubmit: "沒有排查 evidence",
    },
    {
      criteria: "Explanation / reflection",
      weight: isEa ? "15%" : "15%",
      excellent: "能連接概念、操作與 assessment criteria",
      pass: "能描述做了什麼，但原因較薄弱",
      resubmit: "不能解釋主要步驟",
    },
    {
      criteria: "Security / cleanup",
      weight: "10%",
      excellent: "endpoint、權限、費用與 cleanup 都有交代",
      pass: "有基本安全意識",
      resubmit: "忽略公開 endpoint、權限或雲端費用風險",
    },
  ];
}

async function sendAnnualLectureToBuilder(index) {
  const unit = state.annualPlan?.lectureUnits?.[index];
  if (!unit) return;
  const inputs = {
    topic: unit.title,
    subject: state.annualPlan.inputs.moduleTitle,
    audience: state.annualPlan.inputs.audience,
    duration: clamp(unit.videoMinutes, 10, 180),
    style: "考試導向",
    objective: unit.outcomes.join("；"),
    context: [
      state.annualPlan.inputs.context,
      `PPT deck：${unit.deckName}`,
      `Template：${unit.templateId || unit.metadata?.template_id || "LT-CORE"}`,
      `PPT focus：${unit.pptFocus.join("、")}`,
      `Subtopics：${(unit.subtopics || []).join("、")}`,
      `Metadata：${JSON.stringify(unit.metadata || {}, null, 2)}`,
      `Detailed PPTX checklist：${JSON.stringify(getLecturePptxChecklist(unit), null, 2)}`,
      `QA checklist：${(unit.qaChecklist || []).join("；")}`,
      `Duplicate cleanup：${unit.duplicateCleanup}`,
    ].filter(Boolean).join("\n"),
    bloom: ["understand", "apply", "analyze", "evaluate"],
  };

  setFormInputs(inputs);
  dom.scriptMinutes.value = String(Math.max(10, Math.round(unit.videoMinutes)));
  syncDefaultCoreMinutes(true);
  switchView("builder");
  await generateLesson();
  logAudit("年度規劃", `${unit.id} 已送到教材共創 / PPT 流程`);
  markDriveBackupNeeded("年度 lecture 轉 PPT");
  persistState();
}

function renderTimeline() {
  if (!state.slides.length) {
    dom.timeline.innerHTML = emptyText("尚未生成教材");
    return;
  }

  dom.timeline.innerHTML = state.slides
    .map(
      (slide) => `
        <div class="timeline-item">
          <div class="timeline-minutes">${formatNumber(slide.minutes)}</div>
          <div>
            <strong>${escapeHtml(slide.event)}</strong>
            <span>${escapeHtml(slide.bloom)} · ${escapeHtml(slide.activity)}</span>
          </div>
        </div>
      `,
    )
    .join("");
}

function renderSlides() {
  if (!state.slides.length) {
    dom.slideGrid.innerHTML = emptyText("尚未生成教材");
    dom.slideSelect.innerHTML = "";
    return;
  }

  const currentValue = dom.slideSelect.value || "0";
  dom.slideSelect.innerHTML = state.slides
    .map((slide, index) => `<option value="${index}">第 ${slide.number} 頁</option>`)
    .join("");
  dom.slideSelect.value = Number(currentValue) < state.slides.length ? currentValue : "0";

  dom.slideGrid.innerHTML = state.slides
    .map(
      (slide, index) => `
        <article class="slide-card ${String(index) === dom.slideSelect.value ? "active" : ""}" data-slide-index="${index}">
          <div class="slide-title-row">
            <div>
              <h3>${escapeHtml(slide.title)}</h3>
              <div class="slide-meta">
                <span>${escapeHtml(slide.event)}</span>
                ${slide.slideType ? `<span>${escapeHtml(slide.slideType)}</span>` : ""}
                <span>${escapeHtml(slide.bloom)}</span>
                <span>${formatNumber(slide.minutes)} 分</span>
                <span>PPT Prompt</span>
                ${slide.sourceRefs?.length ? `<span>來源 ${slide.sourceRefs.length}</span>` : ""}
              </div>
            </div>
            <span class="slide-number">${slide.number}</span>
          </div>
          <textarea data-slide-notes="${index}" rows="10" aria-label="第 ${slide.number} 頁 PPT prompt">${escapeHtml(slide.notes)}</textarea>
        </article>
      `,
    )
    .join("");

  document.querySelectorAll("[data-slide-index]").forEach((card) => {
    card.addEventListener("click", (event) => {
      if (event.target.closest("textarea")) return;
      dom.slideSelect.value = card.dataset.slideIndex;
      renderSlides();
    });
  });

  document.querySelectorAll("[data-slide-notes]").forEach((textarea) => {
    textarea.addEventListener("input", (event) => {
      const index = Number(event.target.dataset.slideNotes);
      state.slides[index].notes = event.target.value;
      dom.assistantContext.value = buildAssistantContext();
      persistState();
    });
  });

  dom.slideSelect.onchange = renderSlides;
}

async function regenerateSelectedSlide() {
  if (!state.slides.length) return;
  const index = Number(dom.slideSelect.value || 0);
  const slide = state.slides[index];
  const feedback = clean(dom.slideFeedback.value) || "請讓內容更清楚、更適合學生理解";
  setAiBusy(true, "AI 重生成投影片中");
  try {
    const revised = await requestAi("slide-revision", { slide, feedback, inputs: state.lastLessonInputs || getLessonInputs() });
    slide.title = revised.title || slide.title;
    slide.activity = revised.activity || feedback;
    slide.notes = revised.notes || slide.notes;
    slide.speakerNotes = revised.speakerNotes || slide.speakerNotes;
    slide.suggestedLayout = revised.suggestedLayout || slide.suggestedLayout;
    slide.suggestedVisual = revised.suggestedVisual || slide.suggestedVisual;
    slide.factCheckPoints = Array.isArray(revised.factCheckPoints) ? revised.factCheckPoints : slide.factCheckPoints;
    logAudit("局部修改", `AI 已依教師意見更新第 ${slide.number} 頁：${feedback}`);
    dom.slideFeedback.value = "";
    dom.assistantContext.value = buildAssistantContext();
    renderSlides();
    markDriveBackupNeeded("AI 局部修改");
    persistState();
  } catch (error) {
    console.warn(error);
    if (dom.compareBox) dom.compareBox.textContent = `AI 投影片重生成失敗：${error.message}`;
  } finally {
    setAiBusy(false);
  }
}

async function handleMaterialUpload(event) {
  const file = event.target.files[0];
  if (!file) return;
  setMaterialStatus(`正在解析 ${file.name}...`, true);

  try {
    const parsed = await parseMaterialFile(file);
    if (parsed?.text) {
      state.materialPages = parsed.pages || [];
      state.slideJson = Array.isArray(parsed.slideJson) ? parsed.slideJson : [];
      state.materialMeta = {
        filename: parsed.filename || file.name,
        type: parsed.type || file.name.split(".").pop(),
        warning: parsed.warning || "",
        slideJsonCount: state.slideJson.length,
      };
      dom.materialText.value = parsed.text;
      logAudit("教材解析", `${state.materialMeta.filename} 解析為 ${state.materialPages.length || 1} 個片段${state.slideJson.length ? `，並建立 ${state.slideJson.length} 頁 slide_json` : ""}`);
      setMaterialStatus(
        `已解析 ${state.materialMeta.filename}：${state.materialPages.length || 1} 個片段${state.slideJson.length ? `，${state.slideJson.length} 頁 clean slide_json` : ""}${state.materialMeta.warning ? `。${state.materialMeta.warning}` : ""}`,
        true,
      );
      markDriveBackupNeeded("教材解析");
      persistState();
      return;
    }
  } catch (error) {
    console.warn(error);
  }

  if (isPlainTextFile(file)) {
    const text = await readFileAsText(file);
    state.materialPages = textToPages(text);
    state.slideJson = [];
    state.materialMeta = { filename: file.name, type: "text", warning: "" };
    dom.materialText.value = text;
    logAudit("教材解析", `${file.name} 讀入為 ${state.materialPages.length || 1} 個文字片段`);
    setMaterialStatus(`已讀入 ${file.name}：${state.materialPages.length || 1} 個文字片段`, true);
    markDriveBackupNeeded("教材解析");
    persistState();
    return;
  }

  setMaterialStatus("此檔案需要以 node server.js 啟動後才能解析。", false);
}

async function generateScript() {
  const material = clean(dom.materialText.value) || buildMaterialFromSlides();
  const inputs = state.lastLessonInputs || getLessonInputs();
  const startPage = clamp(Number(dom.startPage.value) || 1, 1, 999);
  const minutes = clamp(Number(dom.scriptMinutes.value) || 60, 3, 180);
  const budget = calculateBudget(minutes, getConfiguredCoreMinutes(minutes));
  const wpm = calculateWpm();
  const targetWords = Math.round(budget.core * wpm);
  const slideJson = getCleanSlideJsonForLecture(material, startPage, inputs);
  const scriptPages = slideJsonToScriptPages(slideJson, inputs);
  const focusedMaterial = JSON.stringify(slideJson, null, 2);
  const courseJson = buildCourseJsonForScript(inputs, { minutes, budget, wpm, targetWords });
  const teacherInterview = buildTeacherInterviewForScript(inputs);

  state.budget = { ...budget, wpm, targetWords };
  setAiBusy(true, "生成講稿中");

  try {
    const aiScript = await requestAi("script", {
      inputs,
      material: focusedMaterial,
      slideJson,
      courseJson,
      teacherInterview,
      scriptPages,
      startPage,
      minutes,
      budget,
      wpm,
      targetWords,
    });
    if (aiScript?.script || aiScript?.teacherScriptPages?.length) {
      state.script = composeFormalLectureScript(aiScript, {
        inputs,
        scriptPages,
        fragments: scriptPages.map((page) => `第 ${page.number} 頁：${page.title}\n${page.text}`),
        focusedMaterial,
        startPage,
        minutes,
        budget,
        wpm,
        targetWords,
      });
      const actualWords = countCoreLectureWords(state.script);
      logAudit("講稿生成", `${formatAiProviderName(state.ai.provider)} 依 ${scriptPages.length} 頁 PPT 生成逐頁教師口語稿（核心講授 ${actualWords} 字）`);
      renderScript();
      renderTimeBudget();
      markDriveBackupNeeded("講稿生成");
      persistState();
      return;
    }
  } catch (error) {
    console.warn(error);
    if (dom.compareBox) dom.compareBox.textContent = `AI 講稿生成失敗：${error.message}`;
    logAudit("講稿生成", `AI 生成失敗：${error.message}`);
  } finally {
    setAiBusy(false);
  }
}

async function reviseScript(mode) {
  if (!state.script) {
    await generateScript();
    return;
  }
  setAiBusy(true, "AI 修訂講稿中");
  try {
    const result = await requestAi("script-revision", {
      script: state.script,
      context: getScriptTargetContext(),
      mode,
    });
    if (!result?.script) throw new Error("AI 未回傳修訂講稿。");
    state.script = result.script;
    logAudit("講稿修訂", `AI 已${mode === "shorten" ? "精簡" : "擴寫"}講稿`);
    renderScript();
    markDriveBackupNeeded("AI 講稿修訂");
    persistState();
  } catch (error) {
    console.warn(error);
    if (dom.compareBox) dom.compareBox.textContent = `AI 講稿修訂失敗：${error.message}`;
  } finally {
    setAiBusy(false);
  }
}

function getScriptPagesForLecture(material, startPage, inputs) {
  return slideJsonToScriptPages(getCleanSlideJsonForLecture(material, startPage, inputs), inputs);
}

function getCleanSlideJsonForLecture(material, startPage, inputs) {
  const source = state.slideJson.length
    ? state.slideJson.map(normalizeSlideJsonFromParsedPpt)
    : state.materialPages.length
      ? state.materialPages.map((page, index) => pageToSlideJson(page, index, "ppt"))
      : state.slides.length
        ? state.slides.map((slide, index) => appSlideToCleanSlideJson(slide, index))
        : textToPages(material).map((page, index) => pageToSlideJson(page, index, "text"));
  const cleanedPages = source
    .map((slide, index) => normalizeCleanSlideJson(slide, index, inputs))
    .filter((slide) => slide.slide_title || slide.slide_body || slide.speaker_notes);
  const startIndex = clamp(startPage - 1, 0, Math.max(0, cleanedPages.length - 1));
  return cleanedPages.slice(startIndex);
}

function normalizeSlideJsonFromParsedPpt(slide) {
  return {
    slide_no: slide.slide_no,
    slide_title: slide.slide_title,
    slide_subtitle: slide.slide_subtitle,
    slide_body: slide.slide_body,
    visual_description: slide.visual_description,
    speaker_notes: slide.speaker_notes,
    source_type: slide.source_type || "原教材內容",
    extracted_from: slide.extracted_from || "ppt",
  };
}

function pageToSlideJson(page, index, extractedFrom = "ppt") {
  const rawText = sanitizeLectureSourceText(page.text || "");
  const speakerNotes = extractSpeakerNotesFromCleanText(rawText);
  const bodyWithoutNotes = removeSpeakerNotesFromText(rawText);
  const lines = bodyWithoutNotes.split(/\n+/).map(clean).filter(Boolean);
  const title = sanitizeLectureTitle(page.title || lines[0], index);
  const subtitle = lines[1] && lines[1].length <= 120 && lines[1] !== title ? lines[1] : "";
  const bodyStart = subtitle ? 2 : 1;
  return {
    slide_no: Number(page.number) || index + 1,
    slide_title: title,
    slide_subtitle: subtitle,
    slide_body: lines.slice(bodyStart).join("\n"),
    visual_description: "",
    speaker_notes: speakerNotes,
    source_type: "原教材內容",
    extracted_from: extractedFrom,
  };
}

function appSlideToCleanSlideJson(slide, index) {
  return {
    slide_no: slide.number || index + 1,
    slide_title: slide.title || `投影片 ${index + 1}`,
    slide_subtitle: slide.event || "",
    slide_body: extractVisibleSlideBodyFromAppSlide(slide),
    visual_description: slide.suggestedVisual || firstPromptField(slide.notes, "suggested_visual") || "",
    speaker_notes: slide.speakerNotes || firstPromptField(slide.notes, "speaker_notes") || "",
    source_type: "原教材內容",
    extracted_from: "app_clean_slide",
  };
}

function normalizeCleanSlideJson(slide, index, inputs) {
  const title = sanitizeLectureTitle(slide.slide_title || `投影片 ${index + 1}`, index);
  const subtitle = sanitizeLectureSourceText(slide.slide_subtitle || "");
  const body = sanitizeLectureSourceText(slide.slide_body || "");
  const speakerNotes = sanitizeLectureSourceText(slide.speaker_notes || "");
  const visualDescription = sanitizeLectureSourceText(slide.visual_description || "");
  return {
    slide_no: Number(slide.slide_no) || index + 1,
    slide_title: title,
    slide_subtitle: subtitle,
    slide_body: body,
    visual_description: visualDescription,
    speaker_notes: speakerNotes,
    source_type: "原教材內容",
    extracted_from: slide.extracted_from || "ppt",
    pageType: inferLecturePageType(`${title}\n${subtitle}\n${body}\n${speakerNotes}\n${visualDescription}\n${inputs.topic}`),
    keyPoints: extractScriptKeyPoints([subtitle, body, speakerNotes].filter(Boolean).join("\n"), title),
    sourceQuality: inferSourceQuality(`${body}\n${speakerNotes}`),
  };
}

function slideJsonToScriptPages(slideJson, inputs) {
  return slideJson.map((slide, index) => {
    const text = [
      slide.slide_subtitle,
      slide.slide_body,
      slide.visual_description && `視覺描述：${slide.visual_description}`,
      slide.speaker_notes && `speaker_notes：${slide.speaker_notes}`,
    ].filter(Boolean).join("\n");
    const page = {
      number: slide.slide_no || index + 1,
      title: slide.slide_title || `投影片 ${index + 1}`,
      text,
    };
    const normalized = sanitizeScriptPage(page, index, inputs);
    normalized.pageType = slide.pageType || normalized.pageType;
    normalized.keyPoints = slide.keyPoints?.length ? slide.keyPoints : normalized.keyPoints;
    normalized.sourceQuality = slide.sourceQuality || normalized.sourceQuality;
    normalized.slideJson = slide;
    return normalized;
  });
}

function sanitizeScriptPage(page, index, inputs) {
  const rawText = sanitizeLectureSourceText(page.text || "");
  const rawTitle = clean(page.title) || firstClientLine(rawText) || `投影片 ${index + 1}`;
  const title = sanitizeLectureTitle(rawTitle, index);
  return {
    number: Number(page.number) || index + 1,
    title,
    text: rawText,
    pageType: inferLecturePageType(`${title}\n${rawText}\n${inputs.topic}`),
    keyPoints: extractScriptKeyPoints(rawText, title),
    sourceQuality: inferSourceQuality(rawText),
  };
}

function extractVisibleSlideBodyFromAppSlide(slide) {
  const promptBody = extractPromptFieldLinesClient(slide.notes, "slide_body");
  const candidates = [
    ...promptBody,
    slide.activity,
  ].filter(Boolean);
  if (candidates.length) return candidates.join("\n");
  return sanitizeLectureSourceText(slide.notes || "")
    .split(/\n+/)
    .filter((line) => !/prompt|compiler|layout|visual|course_json|hard|輸出格式|硬性限制/i.test(line))
    .slice(0, 5)
    .join("\n");
}

function extractPromptFieldLinesClient(text, field) {
  const lines = String(text || "").split(/\r?\n/);
  const normalizedField = field.toLowerCase();
  const normalizeKey = (line) => line.trim().toLowerCase().replace(/^\d+\.\s*/, "");
  const start = lines.findIndex((line) => normalizeKey(line).startsWith(`${normalizedField}:`));
  if (start === -1) return [];
  const first = lines[start].trim().replace(/^\d+\.\s*/, "").split(":").slice(1).join(":").trim();
  const output = first ? [first] : [];
  for (let index = start + 1; index < lines.length; index += 1) {
    const line = lines[index].trim();
    if (!line) continue;
    if (/^\d+\.\s*[a-z_ ]+:/i.test(line) || /^[a-z_ ]+:/i.test(line) || /^【/.test(line)) break;
    output.push(line.replace(/^[-*]\s*/, ""));
    if (output.length >= 5) break;
  }
  return output;
}

function firstPromptField(text, field) {
  return extractPromptFieldLinesClient(text, field)[0] || "";
}

function extractSpeakerNotesFromCleanText(text) {
  const match = String(text || "").match(/(?:講者備註|speaker_notes|speaker notes|notes)[:：]\s*([\s\S]*)/i);
  return match ? sanitizeLectureSourceText(match[1]) : "";
}

function removeSpeakerNotesFromText(text) {
  return String(text || "").replace(/(?:講者備註|speaker_notes|speaker notes|notes)[:：]\s*[\s\S]*$/i, "").trim();
}

function buildCourseJsonForScript(inputs, meta = {}) {
  return {
    title: inputs.topic,
    subject_domain: inputs.subject,
    audience_profile: inputs.audience,
    duration_min: meta.minutes || inputs.duration,
    core_teaching_min: meta.budget?.core || "",
    style: inputs.style,
    objectives: splitPlanItems(inputs.objective),
    prerequisites: splitPlanItems(inputs.context),
    teacher_interview_answers: inputs.interviewAnswers || "",
    suggested_wpm: meta.wpm || "",
    target_words_reference: meta.targetWords || "",
    core_teaching_ratio: meta.minutes ? roundOne((meta.budget?.core || 0) / meta.minutes) : DEFAULT_CORE_RATIO,
    checkpoint_policy: `只在重點頁加入互動問題，約每 ${CHECKPOINT_INTERVAL} 頁 1 次`,
  };
}

function buildTeacherInterviewForScript(inputs) {
  return [
    `課題：${inputs.topic}`,
    `科目：${inputs.subject}`,
    `對象：${inputs.audience}`,
    `學習目標：${inputs.objective || "未提供"}`,
    `先備知識與班情：${inputs.context || "未提供"}`,
    `教師追問回答：${inputs.interviewAnswers || "尚未提供"}`,
  ].join("\n");
}

function sanitizeLectureSourceText(text) {
  return String(text || "")
    .replace(/\r/g, "\n")
    .split(/\n+/)
    .map((line) => line.trim())
    .filter(Boolean)
    .filter((line) => !/^講者備註[:：]?\s*(\d+[\s、,.;：:]*){1,12}$/u.test(line))
    .filter((line) => !/^(\d+[\s、,.;：:]*){1,12}$/u.test(line))
    .filter((line) => !/^(PPT Slide Compiler Prompt|Revision Request|輸出格式|硬性限制|本頁素材|建議講者備忘稿|需查核事項)$/i.test(line))
    .filter((line) => !/^(slide_no|slide_type|slide_goal|teaching_event|bloom_level|time_budget_min|layout_preference|visual_preference|linked_assessment)\s*:/i.test(line))
    .filter((line) => !/^(1\. slide_title|2\. slide_subtitle|3\. slide_body|4\. speaker_notes|5\. suggested_visual|6\. suggested_layout|7\. presenter_cues|8\. fact_check_points)\s*:?/i.test(line))
    .filter((line) => !/^[-*]\s*(使用 16:9|每頁|不要杜撰|資訊性圖像|slide_body|speaker_notes|Alt text|decorative)/i.test(line))
    .join("\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function sanitizeLectureTitle(title, index) {
  const cleaned = clean(title)
    .replace(/^講者備註[:：]?\s*/u, "")
    .replace(/^第\s*\d+\s*頁[:：]?\s*/u, "")
    .replace(/^Slide\s*\d+[:：]?\s*/i, "")
    .trim();
  if (!cleaned || /^(\d+[\s、,.;：:]*){1,12}$/u.test(cleaned)) return `投影片 ${index + 1}`;
  return cleaned.slice(0, 90);
}

function inferSourceQuality(text) {
  if (/需教師確認|teacher confirm|待查|最新|版本|考試權重|校內政策|薪資|市場/i.test(text)) return "需教師確認";
  if (/推定補充|assumption|未提供|資訊不足/i.test(text)) return "推定補充";
  return "原教材內容";
}

function inferLecturePageType(text) {
  const value = String(text || "").toLowerCase();
  if (/troubleshooting|debug|error|failed|notready|pending|crashloop|排錯|故障|錯誤|events|logs|describe/.test(value)) return "troubleshooting";
  if (/assessment|acceptance|rubric|checkpoint|skill test|evidence|驗收|評核|評量|測驗|交付|endpoint/.test(value)) return "assessment";
  if (/demo|walkthrough|kubectl|yaml|terminal|command|示範|演示|操作|執行/.test(value)) return "demo";
  if (/lab|實驗|hands-on|task/.test(value)) return "lab";
  return "content";
}

function extractScriptKeyPoints(text, title) {
  const lines = String(text || "")
    .split(/\n+/)
    .map((line) => line.replace(/^[-*]\s*/, "").trim())
    .filter((line) => line.length >= 4)
    .filter((line) => line !== title)
    .filter((line) => !/^(你是一位|請只用|course_json|workflow|style|objectives|prerequisites)/i.test(line));
  return [...new Set(lines)].slice(0, 4);
}

function pagesToScriptSource(pages) {
  return pages
    .map((page) => [
      `第 ${page.number} 頁：${page.title}`,
      `頁面類型：${page.pageType}`,
      page.text,
    ].filter(Boolean).join("\n"))
    .join("\n\n---\n\n");
}

function composeFormalLectureScript(aiScript, context) {
  const pageScripts = mergeTeacherScriptPages(aiScript, context);
  const handout = sanitizeGeneratedSection(aiScript?.selfStudyHandout) || buildStudentSelfStudyHandout(context, pageScripts);
  const generationLog = buildScriptGenerationLog(aiScript, context);
  return [
    `# ${context.inputs.topic}｜逐頁授課稿`,
    "",
    `> 正式授課稿以 ${context.scriptPages.length} 頁 PPT 為準。教師口語講稿只放逐頁授課內容；Prompt 與生成紀錄放在最後一區，避免干擾授課。`,
    "",
    "## 一、教師口語講稿",
    "",
    pageScripts.map(renderTeacherPageScript).join("\n\n"),
    "",
    "## 二、學生自學補充講義",
    "",
    handout,
    "",
    "## 三、生成紀錄 / Prompt",
    "",
    generationLog,
  ].join("\n");
}

function mergeTeacherScriptPages(aiScript, context) {
  const aiPages = Array.isArray(aiScript?.teacherScriptPages) ? aiScript.teacherScriptPages : [];
  const checkpointIndexes = getCheckpointIndexSet(context.scriptPages);
  return context.scriptPages.map((page, index) => {
    const aiPage = aiPages.find((item) => Number(item.pageNumber) === Number(page.number))
      || aiPages.find((item) => clean(item.title) && clean(item.title) === page.title);
    const local = buildLocalTeacherPageScript(page, context, index, checkpointIndexes);
    const shouldUseCheckpoint = checkpointIndexes.has(index);
    if (!aiPage) return local;
    return {
      ...local,
      sourceTags: normalizeSourceTags(aiPage.sourceTags, page),
      teachingPurpose: sanitizeGeneratedSection(aiPage.teachingPurpose) || local.teachingPurpose,
      spokenScript: sanitizeSpokenScript(aiPage.spokenScript) || local.spokenScript,
      checkpoint: shouldUseCheckpoint ? sanitizeGeneratedSection(aiPage.checkpoint) || local.checkpoint : "",
      transition: sanitizeGeneratedSection(aiPage.transition) || local.transition,
    };
  });
}

function buildLocalTeacherPageScript(page, context, index, checkpointIndexes = null) {
  const keyPoints = page.keyPoints.length ? page.keyPoints : [page.title, context.inputs.objective || context.inputs.topic].filter(Boolean);
  const checkpoint = (checkpointIndexes ? checkpointIndexes.has(index) : shouldIncludeCheckpoint(page, index, context.scriptPages.length))
    ? buildPageCheckpoint(page, context, keyPoints)
    : "";
  return {
    pageNumber: page.number,
    title: page.title,
    pageType: page.pageType,
    sourceTags: normalizeSourceTags([page.sourceQuality, "推定補充"], page),
    teachingPurpose: buildTeachingPurpose(page, context, keyPoints),
    spokenScript: buildPageSpokenScript(page, context, keyPoints, index),
    checkpoint,
    transition: buildPageTransition(page, context, index),
  };
}

function renderTeacherPageScript(page) {
  const lines = [
    `### 第 ${page.pageNumber} 頁：${page.title}`,
    "",
    `- 內容來源標記：${page.sourceTags.join("、")}`,
    "",
    "#### 本頁教學目的",
    page.teachingPurpose,
    "",
    "#### 教師口語講稿",
    page.spokenScript,
    "",
  ];
  if (clean(page.checkpoint)) {
    lines.push("#### 互動問題 / checkpoint", page.checkpoint, "");
  }
  lines.push("#### 轉場語", page.transition);
  return lines.join("\n");
}

function shouldIncludeCheckpoint(page, index, totalPages) {
  const type = page?.pageType || "";
  if (["demo", "troubleshooting", "assessment", "lab"].includes(type)) return true;
  if (totalPages <= CHECKPOINT_INTERVAL) return index === totalPages - 1;
  return (index + 1) % CHECKPOINT_INTERVAL === 0;
}

function getCheckpointIndexSet(pages = []) {
  const total = pages.length;
  if (!total) return new Set();
  const targetCount = Math.max(1, Math.ceil(total / CHECKPOINT_INTERVAL));
  const candidates = [];
  pages.forEach((page, index) => {
    if (["demo", "troubleshooting", "assessment", "lab"].includes(page.pageType)) candidates.push(index);
  });
  pages.forEach((_, index) => {
    if ((index + 1) % CHECKPOINT_INTERVAL === 0) candidates.push(index);
  });
  candidates.push(total - 1);
  return new Set([...new Set(candidates)].slice(0, targetCount));
}

function normalizeSourceTags(tags, page) {
  const allowed = ["原教材內容", "推定補充", "需教師確認"];
  const source = Array.isArray(tags) ? tags : [tags].filter(Boolean);
  const normalized = source
    .map((tag) => allowed.find((item) => String(tag || "").includes(item)))
    .filter(Boolean);
  if (page?.sourceQuality && !normalized.includes(page.sourceQuality)) normalized.unshift(page.sourceQuality);
  if (!normalized.includes("原教材內容")) normalized.unshift("原教材內容");
  if (page?.text?.length < 80 && !normalized.includes("需教師確認")) normalized.push("需教師確認");
  return [...new Set(normalized)].slice(0, 3);
}

function buildTeachingPurpose(page, context, keyPoints) {
  const firstPoint = keyPoints[0] || page.title;
  if (page.pageType === "demo") {
    return `【原教材內容】本頁用來把「${page.title}」轉成可跟做的示範流程。【推定補充】學生要看見操作順序、預期輸出，以及失敗時如何退回檢查。`;
  }
  if (page.pageType === "troubleshooting") {
    return `【原教材內容】本頁聚焦錯誤現象與排查入口。【推定補充】學生要學會先找第一個 command 或 evidence，而不是直接猜答案。`;
  }
  if (page.pageType === "assessment") {
    return `【原教材內容】本頁把學習成果轉成可驗收條件。【推定補充】驗收要能觀察、重做、截圖，或用 command output 證明。`;
  }
  return `【原教材內容】本頁先處理「${firstPoint.slice(0, 80)}」。【推定補充】教師要把投影片文字轉成一個清楚判斷，讓學生知道這頁和「${context.inputs.topic}」的關係。`;
}

function buildPageSpokenScript(page, context, keyPoints, index) {
  const point = keyPoints[0] || page.title;
  const second = keyPoints[1] || context.inputs.objective || "本課的學習目標";
  const opener = [
    "同學，這一頁我們先不要急着抄字，先看它想解決什麼問題。",
    "來到這一頁，我希望大家把焦點放在判斷方法，而不是單一答案。",
    "這頁的重點不是背名詞，而是看清楚它在整個任務中的位置。",
    "請大家先看標題，再看畫面中最能證明結果的那一個線索。",
  ][index % 4];

  if (page.pageType === "demo") {
    return `【原教材內容】${opener} 這個 demo 先由需求開始，再看要輸入哪個指令或 YAML，最後用 output 驗收。你要留意「${point.slice(0, 70)}」這個線索。【推定補充】正常情況下，我會先示範最小可行步驟，再請大家預測畫面會出現什麼。如果輸出不如預期，不要重做整個 lab，先回到上一個 command、log 或 YAML 欄位，找出是哪一層開始偏離。`;
  }
  if (page.pageType === "troubleshooting") {
    return `【原教材內容】${opener} 如果你看到錯誤，第一步不是猜原因，而是把現象變成 evidence。針對「${point.slice(0, 70)}」，先問：我應該看 status、events、logs，還是 endpoint response？【推定補充】例如 Kubernetes 問題通常先用 get 看狀態，再用 describe 看 events，必要時才進 logs。這樣排查才有次序，也比較接近 CKA/CKAD 的實戰要求。`;
  }
  if (page.pageType === "assessment") {
    return `【原教材內容】${opener} 這頁關心的是怎樣證明你真的完成任務，而不是只說「我做了」。請把「${point.slice(0, 70)}」改寫成可驗收句子。【推定補充】好的驗收條件一定能被重做，例如提交 YAML、command output、截圖、endpoint 測試或短答解釋。若別人在另一台環境無法確認，你的答案就還未足夠。`;
  }
  return `【原教材內容】${opener} 這頁提到「${point.slice(0, 80)}」，我們要把它接回「${second.slice(0, 70)}」。【推定補充】你可以把它理解成一個中間橋樑：先知道概念是什麼，再知道它在實作或評核中怎樣被看見。等一下我會請你用一句話說出本頁最重要的判斷，這比抄完整段文字更重要。`;
}

function buildPageCheckpoint(page, context, keyPoints) {
  if (page.pageType === "demo") {
    return `【推定補充】請學生指出 demo 的三個 evidence：輸入了什麼、預期輸出是什麼、失敗時第一個 fallback 檢查點是什麼。`;
  }
  if (page.pageType === "troubleshooting") {
    return "【推定補充】請學生用 30 秒回答：這個錯誤現象第一個要查的 command / evidence 是哪一個？為什麼？";
  }
  if (page.pageType === "assessment") {
    return "【推定補充】請學生把本頁要求改寫成一條 acceptance criterion，必須包含可觀察 evidence。";
  }
  return `【原教材內容】請學生用一句話說明「${(keyPoints[0] || page.title).slice(0, 60)}」的作用；【推定補充】再請一位同學補充它和本課任務的關係。`;
}

function buildPageTransition(page, context, index) {
  if (page.pageType === "demo") {
    return "完成示範流程後，下一頁我們會把剛才看到的 output 轉成學生自己要完成的任務。";
  }
  if (page.pageType === "troubleshooting") {
    return "有了排查入口後，下一步要看這個證據如何變成修正動作或評核要求。";
  }
  if (page.pageType === "assessment") {
    return "確認驗收條件後，我們就能回到前面的概念，檢查自己是否真的掌握。";
  }
  return index === context.scriptPages.length - 1
    ? "這頁收束後，我們會整理今天的自學補充和課後檢查。"
    : "掌握這一頁後，下一頁會把概念推進到更具體的操作、例子或驗收。";
}

function sanitizeGeneratedSection(value) {
  const text = sanitizeLectureSourceText(value || "");
  return text.replace(/^(本頁教學目的|教師口語講稿|互動問題\s*\/\s*checkpoint|轉場語)[:：]?\s*/u, "").trim();
}

function sanitizeSpokenScript(value) {
  const text = sanitizeGeneratedSection(value);
  if (!text || /PPT Slide Compiler Prompt|輸出格式|硬性限制/i.test(text)) return "";
  return text;
}

function buildStudentSelfStudyHandout(context, pageScripts) {
  const takeaways = pageScripts.slice(0, 8).map((page) => `- 第 ${page.pageNumber} 頁：${page.title}。自學時先讀本頁目的，再用一句話整理本頁核心判斷。`).join("\n");
  const checkpointPages = pageScripts.filter((page) => clean(page.checkpoint)).map((page) => `第 ${page.pageNumber} 頁`).join("、") || "本輪未設定";
  const technicalCue = context.scriptPages.some((page) => ["demo", "troubleshooting", "assessment"].includes(page.pageType))
    ? "\n\n### 技術課自學提醒\n- Demo 頁：記錄操作流程、預期輸出和 fallback。\n- Troubleshooting 頁：把錯誤現象對應到第一個 command 或 evidence。\n- 驗收頁：提交可觀察、可重做、可截圖或可用 command 證明的 evidence。"
    : "";
  return [
    "### 自學閱讀路線",
    "這份補充講義給學生課後自學使用，不取代教師逐頁口語講稿。閱讀時不要只看投影片文字，請把每頁都轉成一個「我能不能用自己的話說明」的判斷。",
    "",
    takeaways || `- 先重讀「${context.inputs.topic}」的核心概念，再完成重點頁 checkpoint。`,
    "",
    `### 重點互動頁\n${checkpointPages}`,
    technicalCue,
    "",
    "### 課後自我檢查",
    "1. 我能否用自己的話說出每頁最重要的一個判斷？",
    "2. 若有 demo，我能否說出預期輸出與失敗時第一個檢查點？",
    "3. 若有評核或 lab，我能否提交可重做的 evidence，而不是只交一張截圖？",
  ].join("\n");
}

function buildScriptGenerationLog(aiScript, context) {
  const mode = aiScript ? `${formatAiProviderName(state.ai.provider)} structured output` : "AI unavailable";
  const notes = Array.isArray(aiScript?.teachingNotes) ? aiScript.teachingNotes : [];
  return [
    `- 生成模式：${mode}`,
    `- PPT 頁數：${context.scriptPages.length}`,
    `- 起始頁：${context.startPage}`,
    `- 目標分鐘：${context.minutes}`,
    `- 核心講授分鐘：${formatNumber(context.budget.core)} 分；目標字數只計算教師口語核心講授。`,
    `- 講稿規則：正式區只放逐頁教師口語稿；互動 checkpoint 只放重點頁，約每 ${CHECKPOINT_INTERVAL} 頁 1 次；學生自學補充放第二區；Prompt / log 放第三區。`,
    `- 來源標記：原教材內容 / 推定補充 / 需教師確認。`,
    notes.length ? `- AI 課前提醒：${notes.join("；")}` : "- AI 課前提醒：無",
    "",
    "### Prompt 摘要",
    "系統要求每頁固定輸出：本頁教學目的、教師口語講稿、轉場語；互動問題 / checkpoint 只在重點頁輸出。Demo、Troubleshooting、驗收條件頁需加入對應實務元素。",
  ].join("\n");
}

async function completeScriptToTarget() {
  if (!state.script) {
    await generateScript();
  }
  const context = getScriptTargetContext();
  setAiBusy(true, "AI 補足講稿字數中");
  try {
    const result = await requestAi("script-revision", { script: state.script, context, mode: "complete_to_target_words" });
    if (!result?.script) throw new Error("AI 未回傳補足講稿。");
    state.script = result.script;
    logAudit("講稿字數達標", `AI 已補到 ${countCoreLectureWords(state.script)}/${context.targetWords} 字`);
    renderScript();
    markDriveBackupNeeded("AI 講稿字數達標");
    persistState();
  } catch (error) {
    console.warn(error);
    if (dom.compareBox) dom.compareBox.textContent = `AI 補足講稿失敗：${error.message}`;
  } finally {
    setAiBusy(false);
  }
}

async function addCoreLectureDepth() {
  if (!state.script) {
    await generateScript();
    return;
  }
  const context = getScriptTargetContext();
  setAiBusy(true, "AI 補強核心講授中");
  try {
    const result = await requestAi("script-revision", { script: state.script, context, mode: "add_core_lecture_depth" });
    if (!result?.script) throw new Error("AI 未回傳核心補強講稿。");
    state.script = result.script;
    logAudit("講稿補核心", `AI 補強後共 ${countCoreLectureWords(state.script)} 字`);
    renderScript();
    markDriveBackupNeeded("AI 講稿補核心");
    persistState();
  } catch (error) {
    console.warn(error);
    if (dom.compareBox) dom.compareBox.textContent = `AI 補強核心講授失敗：${error.message}`;
  } finally {
    setAiBusy(false);
  }
}

function getScriptTargetContext() {
  const inputs = state.lastLessonInputs || getLessonInputs();
  const minutes = clamp(Number(dom.scriptMinutes.value) || 60, 3, 180);
  const budget = calculateBudget(minutes, getConfiguredCoreMinutes(minutes));
  const wpm = calculateWpm();
  const targetWords = Math.round(budget.core * wpm);
  const material = clean(dom.materialText.value) || buildMaterialFromSlides();
  const startPage = clamp(Number(dom.startPage.value) || 1, 1, 999);
  const scriptPages = getScriptPagesForLecture(material, startPage, inputs);
  return {
    inputs,
    scriptPages,
    fragments: scriptPages.map((page) => `第 ${page.number} 頁：${page.title}\n${page.text}`),
    focusedMaterial: pagesToScriptSource(scriptPages) || material,
    startPage,
    minutes,
    budget,
    wpm,
    targetWords,
  };
}

function renderScriptGoal() {
  if (!dom.scriptGoalStatus) return;
  const context = getScriptTargetContext();
  const words = countCoreLectureWords(state.script);
  const percent = context.targetWords ? Math.min(100, Math.round((words / context.targetWords) * 100)) : 0;
  const gap = Math.max(0, context.targetWords - words);
  const status = words >= context.targetWords
    ? "核心講授字數已達標；學生自學補充與生成紀錄不會計入。"
    : `目前只計算教師核心講授，不計學生自學補充與生成紀錄；尚差約 ${gap} 字。`;
  dom.scriptGoalStatus.innerHTML = `
    <div class="target-meter-row">
      <div>
        <span>核心講授字數達標器</span>
        <strong>${escapeHtml(words)} / ${escapeHtml(context.targetWords)} 字</strong>
      </div>
      <span>${escapeHtml(percent)}%</span>
    </div>
    <div class="target-bar"><span style="width:${percent}%"></span></div>
    <small>${escapeHtml(status)}</small>
  `;
}

function buildScriptTargetExpansionBlock(context, index) {
  const fragments = context.fragments.length ? context.fragments : textToPages(context.focusedMaterial).map((page) => page.text).slice(0, 4);
  const source = fragments[index % Math.max(fragments.length, 1)] || context.inputs.context || context.inputs.objective || context.inputs.topic;
  const evidenceCue = inferScriptEvidenceCue(context.inputs);
  return `【核心講授補足 ${index + 1}｜教師可直接講授】\n
這一段用來把「${context.inputs.topic}」講得更完整。請先記住一個原則：學生不是只需要知道名詞，而是需要知道名詞背後的判斷方法。當你看到教材或投影片出現一個概念時，請先問三件事。第一，這個概念解決什麼問題？第二，它依賴哪些先備條件？第三，如果結果不符合預期，我可以用什麼 evidence 證明問題在哪一層？\n
本段素材重點是：${String(source).replace(/\s+/g, " ").slice(0, 520)}\n
以技術課堂來說，最有價值的學習不是把指令背下來，而是能把「情境、操作、證據、解釋」串成一條線。例如學生做完一個 Kubernetes 任務後，不應只提交截圖，而要說明使用了哪個 YAML、哪個 kubectl 指令、看到什麼 output、這個 output 如何證明任務完成。這樣的寫法可以直接服務 CA Lab、CKA/CKAD 練習和期末 Skill Test。\n
對學生而言，本課的最低要求是能跟着步驟完成；較高要求是能解釋為什麼這樣做；最高要求是遇到錯誤時能提出排查假設。請用以下問題自我檢查：如果我只能保留三句話，我會如何解釋這頁？如果 demo 失敗，我第一個會查什麼？如果這題變成評核，老師應該如何用 ${evidenceCue} 驗收？`;
}

function buildWorkedExampleDepthBlock(context) {
  const topic = context.inputs.topic;
  return `【Worked Example｜從課堂講解轉成學生可理解的步驟】\n
假設你現在要完成一個與「${topic}」相關的任務，不要先急着找答案。第一步是重寫任務要求，把它改成可驗收句子，例如「我需要建立某個 resource，並用指定 command 證明它已經運作」。第二步是列出你要看的證據，例如 status、events、logs、endpoint response、YAML spec 或版本資訊。第三步才是執行操作。這個順序能避免學生只照抄命令，卻不知道命令成功代表什麼。\n
如果這堂課之後要接 CA Lab，請把本段轉成 Lab evidence checklist。如果之後要接 Assessment，請把它轉成 marking rubric：正確性、可重現性、排錯證據、解釋能力和安全意識。這樣講稿不只是口頭稿，也能變成學生課後閱讀和做功課的路線圖。`;
}

function inferScriptEvidenceCue(inputs) {
  const text = `${inputs.topic} ${inputs.subject} ${inputs.context}`.toLowerCase();
  if (/kubernetes|cka|ckad|kubectl|yaml/.test(text)) return "kubectl output、YAML、logs 與 endpoint";
  if (/linux|vm|server/.test(text)) return "terminal output、service status 與系統設定截圖";
  if (/aws|eks|cloud/.test(text)) return "AWS console 記錄、endpoint 測試與 cleanup record";
  return "截圖、短答、操作紀錄與反思";
}

function ensureCompleteLectureScript(script, context) {
  const minimumWords = Math.max(600, Math.round(context.targetWords * 0.92));
  let output = normalizeLectureScriptForStudents(script, context);

  if (countWords(output) >= minimumWords) {
    return appendScriptWordBudgetNote(output, context);
  }

  const expansion = buildSelfStudyExpansion(context, minimumWords - countWords(output));
  output = `${output}\n\n${expansion}`;

  while (countWords(output) < minimumWords) {
    output = `${output}\n\n${buildAdditionalDeepeningBlock(context, countWords(output), minimumWords)}`;
  }

  return appendScriptWordBudgetNote(output, context);
}

function normalizeLectureScriptForStudents(script, context) {
  const normalized = clean(script);
  const header = [
    `# ${context.inputs.topic}｜完整課堂講稿`,
    "",
    `目標時間：${context.minutes} 分鐘｜核心講授：${formatNumber(context.budget.core)} 分鐘｜目標字數：約 ${context.targetWords}`,
    "",
    "這份講稿設計成可由教師口頭講授，也可讓學生課後自行閱讀。內容會把 PPT / PPT Prompt 轉成完整講義，而不是只保留版面提示。",
    "",
    "## Executive Summary",
    `本堂課會圍繞「${context.inputs.topic}」建立可操作的理解：先連接先備知識，再拆解核心概念，最後用 checkpoint、demo 或評核任務確認學生能否遷移。`,
    "",
    "## Assumptions / 需教師確認",
    context.inputs.interviewAnswers
      ? `教師追問回答已納入：${context.inputs.interviewAnswers}`
      : "尚未提供教師追問回答；若教材資訊不足，以下延伸會以「推定補充」處理。",
    "",
    "## Slide-by-Slide Full Spoken Script",
  ].join("\n");

  if (!normalized) return header;
  if (normalized.startsWith("# ")) return normalized;
  return `${header}\n\n${normalized}`;
}

function appendScriptWordBudgetNote(script, context) {
  const words = countWords(script);
  const note = `\n\n【字數與使用方式】\n目前講稿約 ${words} 字；目標核心講授字數約 ${context.targetWords} 字。若課堂時間不足，可先刪減「延伸閱讀」或「自我檢查」段落；若學生自學，建議完整閱讀並完成每段 checkpoint。`;
  const withoutOldNote = String(script || "").replace(/\n\n【字數與使用方式】[\s\S]*$/u, "");
  return `${withoutOldNote}${note}`;
}

function buildSelfStudyExpansion(context, deficit) {
  const fragments = context.fragments.length ? context.fragments : textToPages(context.focusedMaterial).map((page) => page.text).slice(0, 6);
  const usable = fragments.length ? fragments : [
    `${context.inputs.topic} 需要學生把 Linux 基礎、Kubernetes resource model 與實際操作連接起來。`,
  ];
  const blocks = [
    `【完整自學補充講義｜補足核心講授】\n以下內容用來把短講稿擴展成學生可以自行閱讀的完整版本。它不是額外離題補充，而是把投影片 prompt 中的重點轉成可理解的課堂文字。`,
  ];

  usable.slice(0, 8).forEach((fragment, index) => {
    blocks.push(buildFragmentTeachingBlock(fragment, index, context));
  });

  blocks.push(buildCommandAndYamlBridge(context));
  blocks.push(buildCommonMistakesBlock(context));
  blocks.push(buildStudentSelfCheckBlock(context));

  let output = blocks.join("\n\n");
  while (countWords(output) < Math.min(deficit, context.targetWords * 0.65)) {
    output = `${output}\n\n${buildAdditionalDeepeningBlock(context, countWords(output), deficit)}`;
  }
  return output;
}

function buildFragmentTeachingBlock(fragment, index, context) {
  const topic = context.inputs.topic;
  const title = firstClientLine(fragment) || `教材片段 ${index + 1}`;
  const cleanFragment = String(fragment || "").replace(/\s+/g, " ").slice(0, 900);
  return `【核心段落 ${index + 1}：${title}】

這一段要讓學生理解的不是單一指令，而是「為什麼這個步驟在 ${topic} 中必要」。請先把它看成一條因果鏈：底層環境是否穩定，會影響 Kubernetes 元件能否啟動；Kubernetes 元件是否正常，會影響 Pod、Service、Ingress 等資源能否被建立；而資源能否被建立，最後會影響應用是否真的可以對外提供服務。

教材素材重點：${cleanFragment}

學生閱讀時要特別留意三件事。第一，這裡出現的每個名詞都應該能連回一個實際檢查方法，例如 node 狀態、事件、log、YAML 欄位或 service endpoint。第二，當你看到錯誤時，不要只背答案，而要問：錯誤是在 Linux 層、container runtime 層、Kubernetes API 層，還是 application workload 層？第三，CKA 與 CKAD 的差別不是誰比較高級，而是責任範圍不同：CKA 關注 cluster 能不能穩定運作，CKAD 關注應用能不能以 Kubernetes-native 的方式被描述和部署。

Checkpoint：請用自己的話寫下本段最重要的一個「原因」和一個「結果」。如果你只能抄名詞，代表你還沒有真正掌握。`;
}

function buildCommandAndYamlBridge(context) {
  return `【從 Linux 指令到 Kubernetes YAML 的橋樑】

學習 ${context.inputs.topic} 時，最容易卡住的地方，是從 Linux 管理員熟悉的命令式思維，轉向 Kubernetes 的宣告式思維。在 Linux 中，你可能會輸入 systemctl start nginx、檢查 journalctl、查看 df -h 或 ip addr。這些操作的重點是「我現在叫系統做一件事」。但在 Kubernetes 中，你更多時候是寫一份 YAML，描述你想要的最終狀態，例如需要幾個 replicas、使用哪個 image、開放哪個 port、是否需要 readinessProbe。

這種轉換非常重要。因為 Kubernetes 的控制器會不斷把現實狀態拉回你宣告的狀態。如果 Pod 掛掉，Deployment 會嘗試重建；如果 Service selector 不對，流量就找不到目標 Pod；如果 resource request 過高，scheduler 可能無法安排到 node。學生需要理解：YAML 不是普通設定檔，而是你交給 Kubernetes API 的「合約」。

因此，在讀每一頁 PPT 時，都要問三個問題。第一，這個畫面對應哪一個 resource？第二，這個 resource 的 spec 描述了什麼 desired state？第三，當 desired state 和 actual state 不一致時，我可以用哪個 kubectl 指令找到差異？這三個問題就是從 Linux 管理員走向 Kubernetes 專家的橋樑。`;
}

function buildCommonMistakesBlock(context) {
  return `【常見錯誤與排查路徑】

學生在這堂課常見的第一個錯誤，是把 Kubernetes 問題全部當成 YAML 打錯。其實很多 cluster 問題來自 Linux 層，例如 swap 未關閉、DNS 解析失敗、時間不同步、磁碟滿了、container runtime 未正常啟動，或者 firewall rules 阻擋了必要流量。當 node 狀態是 NotReady 時，不應該立刻改 Deployment，而應該先檢查 node、kubelet、container runtime 和 network。

第二個錯誤，是只看 kubectl get 的表面結果。kubectl get 只能告訴你目前狀態，不能完整告訴你原因。要追原因，就要用 kubectl describe 看 events，用 kubectl logs 看 container output，用 kubectl get yaml 看實際 spec，用 kubectl explain 查欄位意義。對 CKA/CKAD 來說，這些排查指令比背大量答案更重要。

第三個錯誤，是忽略 public endpoint 的驗收。在期末 Skill Test 或 EKS lab 中，部署成功不代表任務完成。你必須證明外部可以訪問，並說明 endpoint、port、service type、ingress rule 或 load balancer 的角色。如果 endpoint 不能公開，就要能說明是 security group、service selector、ingress controller、DNS 還是 application 本身造成問題。`;
}

function buildStudentSelfCheckBlock(context) {
  return `【學生自我檢查】

讀完這堂課後，請不用看答案完成以下檢查。第一，請列出安裝或運行 Kubernetes 前，一台 Linux VM 至少要檢查的五件事，並說明每一件事如果出錯會造成什麼後果。第二，請解釋 CKA 和 CKAD 的責任分界：哪一些問題通常由 cluster administrator 處理，哪一些問題通常由 application developer 處理。第三，請用一個例子說明命令式操作和宣告式 YAML 的差別。第四，請寫出你排查一個 Pod 無法啟動時會採用的四個指令或步驟。

如果你可以完成以上四項，代表你已經不只是聽過名詞，而是能把 Linux、Kubernetes resource、kubectl 排查和 assessment 要求連接起來。這也是本課的真正目標：不只是生成一份 PPT 或背一份講稿，而是讓你能在真實 lab、CKA/CKAD 練習和期末 skill test 中獨立判斷下一步。`;
}

function buildAdditionalDeepeningBlock(context, currentWords, targetWords) {
  return `【延伸閱讀補足｜${currentWords}/${targetWords}】

為了讓本堂內容可以真正支持自學，這裡再補一層推理。學習 Kubernetes 時，學生很容易把注意力放在工具名稱，例如 kubeadm、kubectl、Minikube、EKS、Rancher。但工具只是入口，真正要建立的是判斷框架。當你看到一個任務時，先判斷它屬於環境準備、cluster 管理、application deployment、network exposure、storage binding、security policy 還是 troubleshooting。分類完成後，再選擇對應工具。

例如，如果任務要求你公開一個服務，你要先確認 Pod label 是否正確，再確認 Service selector 是否能找到 Pod，然後確認 port、targetPort、nodePort 或 ingress rule 是否配合。這個過程不是背指令，而是逐層驗證資料流。又例如，如果任務要求你修復 node 問題，你要先看 node condition，再看 kubelet 和 container runtime，而不是直接改 application YAML。

這種逐層思考方式，就是從初學者走向 CKA/CKAD 水平的關鍵。請把每個 lab 都當成一次排查訓練：先觀察現象，再提出假設，接著用指令驗證，最後記錄證據。`;
}

function renderScript() {
  dom.scriptOutput.value = state.script || "";
  renderScriptGoal();
  renderStatus();
}

function renderTimeBudget() {
  const minutes = clamp(Number(dom.scriptMinutes.value) || 60, 3, 180);
  syncDefaultCoreMinutes();
  const coreMinutes = getConfiguredCoreMinutes(minutes);
  if (dom.coreMinutes && Number(dom.coreMinutes.value) !== coreMinutes) {
    dom.coreMinutes.value = formatNumber(coreMinutes);
  }
  const budget = state.budget && Number(state.budget.total) === minutes && Number(state.budget.core) === coreMinutes
    ? state.budget
    : calculateBudget(minutes, coreMinutes);
  const wpm = calculateWpm();
  const targetWords = Math.round(budget.core * wpm);

  dom.timeBudget.innerHTML = [
    ["開場", budget.opening],
    ["核心講授", budget.core],
    ["問答反思", budget.qa],
    ["切換緩衝", budget.buffer],
  ]
    .map(
      ([label, value]) => {
        const percent = minutes ? Math.max(0, Math.min(100, Math.round((value / minutes) * 100))) : 0;
        return `
        <div class="budget-row">
          <div class="budget-label"><span>${label}</span><strong>${formatNumber(value)} 分</strong></div>
          <div class="budget-bar"><span style="width:${percent}%"></span></div>
        </div>
      `;
      },
    )
    .join("");

  dom.wpmStatus.textContent = String(wpm);
  dom.coreMinutesStatus.textContent = `${formatNumber(budget.core)} 分`;
  dom.targetWordsStatus.textContent = String(targetWords);
  renderScriptGoal();
}

async function sendAssistantMessage() {
  const question = clean(dom.assistantQuestion.value);
  if (!question) return;
  const context = clean(dom.assistantContext.value) || buildAssistantContext();
  state.messages.push({ role: "user", text: question });
  dom.assistantQuestion.value = "";
  renderChat();

  setAiBusy(true, "助理回應中");
  try {
    const aiReply = await requestAi("assistant", { context, question });
    if (aiReply?.answer) {
      const checks = Array.isArray(aiReply.checks) && aiReply.checks.length
        ? `\n\n需查核：\n${aiReply.checks.map((item) => `- ${item}`).join("\n")}`
        : "";
      const nextMove = aiReply.nextMove ? `\n\n下一步：${aiReply.nextMove}` : "";
      state.messages.push({ role: "assistant", text: `${aiReply.answer}${checks}${nextMove}` });
      logAudit("即時助理", `${formatAiProviderName(state.ai.provider)} 生成課堂回應`);
    } else {
      throw new Error("AI 未回傳有效課堂回應。");
    }
  } catch (error) {
    console.warn(error);
    state.messages.push({ role: "assistant", text: `AI 回應生成失敗：${error.message}` });
    logAudit("即時助理", `AI 生成失敗：${error.message}`);
  } finally {
    setAiBusy(false);
  }

  renderChat();
  persistState();
}

function renderChat() {
  if (!state.messages.length) {
    dom.chatLog.innerHTML = emptyText("尚未有回應紀錄");
    return;
  }

  dom.chatLog.innerHTML = state.messages
    .map(
      (message) => `
        <div class="message ${message.role === "user" ? "user" : "assistant"}">
          <strong>${message.role === "user" ? "問題" : "助理"}</strong>
          <p>${escapeHtml(message.text)}</p>
        </div>
      `,
    )
    .join("");
  dom.chatLog.scrollTop = dom.chatLog.scrollHeight;
}

function renderPublishedQa() {
  if (!dom.publishedSummary || !dom.studentAnswer || !dom.sourceList) return;

  if (!state.publishedRevision) {
    dom.publishedSummary.innerHTML = "尚未發布教材。教師按「發布教材」後，學生端才會根據已發布版本回答。";
  } else {
    const revision = state.publishedRevision;
    const materialCount = revision.materialPages?.length || 0;
    dom.publishedSummary.innerHTML = `
      <strong>${escapeHtml(revision.inputs.topic || "已發布教材")}</strong><br />
      發布時間：${escapeHtml(new Date(revision.publishedAt).toLocaleString("zh-Hant"))}<br />
      投影片：${revision.slides.length} 頁｜教材片段：${materialCount}｜模式：教材有據
    `;
  }

  dom.studentQuestion.value = state.studentQa.question || "";
  dom.studentAnswer.innerHTML = state.studentQa.answer
    ? `<strong>${escapeHtml(state.studentQa.mode)}</strong>\n${escapeHtml(state.studentQa.answer)}`
    : "學生提問後，答案會在這裡顯示。";
  renderStudentSources(state.studentQa.sources || []);
}

function renderStudentSources(sources) {
  if (!sources.length) {
    dom.sourceList.innerHTML = emptyText("尚未有來源引用");
    return;
  }

  dom.sourceList.innerHTML = sources
    .map(
      (source) => `
        <article class="source-item">
          <strong>${escapeHtml(source.label)}</strong>
          <span>${escapeHtml(source.preview)}</span>
          <span>chunk: ${escapeHtml(source.id || "n/a")}｜hash: ${escapeHtml((source.sourceHash || "").slice(0, 12))}｜confidence: ${Math.round((source.confidence || 0) * 100)}%</span>
        </article>
      `,
    )
    .join("");
}

async function askPublishedLesson() {
  const question = clean(dom.studentQuestion.value);
  if (!question) return;

  if (!state.publishedRevision) {
    state.studentQa = {
      question,
      mode: "尚未發布",
      answer: "目前沒有已發布教材。請教師先發布教材版本，學生端才可提問。",
      sources: [],
    };
    renderPublishedQa();
    return;
  }

  const sources = retrievePublishedSources(question, state.publishedRevision);
  const supported = sources.length && sources[0].score >= 3.5;
  const citationSources = sources.slice(0, 4).map((source) => ({
    id: source.id,
    type: source.type,
    label: source.label,
    preview: source.preview,
    confidence: source.confidence,
    sourceHash: source.sourceHash,
    text: String(source.text || "").slice(0, 1600),
  }));

  setAiBusy(true, "AI 生成學生問答中");
  try {
    const result = await requestAi("student-qa", {
      question,
      supported,
      sources: citationSources,
      revisionMeta: {
        topic: state.publishedRevision.inputs?.topic || "",
        level: state.publishedRevision.inputs?.level || "",
        publishedAt: state.publishedRevision.publishedAt,
        slideCount: state.publishedRevision.slides?.length || 0,
      },
    });
    state.studentQa = {
      question,
      mode: supported ? "AI 教材有據" : "AI 拒答 / 需老師補充",
      answer: [result.answer, result.nextMove ? `下一步：${result.nextMove}` : "", ...(result.checks || []).map((item) => `檢查：${item}`)]
        .filter(Boolean)
        .join("\n"),
      sources: supported ? citationSources : [],
    };
    updateQaMetrics(supported);
    logAudit("學生問答", `${state.studentQa.mode}：${question.slice(0, 80)}`);
    renderPublishedQa();
    renderGovernanceMetrics();
    persistState();
  } catch (error) {
    state.studentQa = {
      question,
      mode: "AI 生成失敗",
      answer: `AI 學生問答生成失敗：${error.message}`,
      sources: [],
    };
    logAudit("學生問答", `AI 生成失敗：${error.message}`);
    renderPublishedQa();
  } finally {
    setAiBusy(false);
  }
}

function retrievePublishedSources(question, revision) {
  const index = revision.citationIndex?.length ? revision.citationIndex : buildCitationIndex(revision);
  const queryVector = makeSearchVector(question);

  return index
    .map((source) => {
      const lexical = scoreSource(source.text, extractKeywords(question));
      const semantic = cosineSimilarity(queryVector, source.vector);
      const score = lexical + semantic * 10;
      return {
        ...source,
        score,
        confidence: Math.min(0.98, Number((score / 18).toFixed(2))),
      };
    })
    .filter((source) => source.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, 6);
}

function buildCitationIndex(revision) {
  const slideChunks = (revision.slides || []).map((slide) => {
    const text = `${slide.title}\n${slide.event}\n${slide.bloom}\n${slide.activity}\n${slide.notes}`;
    return createCitationChunk({
      type: "slide",
      number: slide.number || 0,
      label: `投影片 ${slide.number || ""}：${slide.title}`,
      text,
      preview: `${slide.event || ""} ${slide.bloom || ""} ${String(slide.notes || "").slice(0, 180)}`,
    });
  });
  const materialChunks = (revision.materialPages || []).map((page) => createCitationChunk({
    type: "material",
    number: page.number || 0,
    label: `教材片段 ${page.number || ""}：${page.title}`,
    text: page.text || "",
    preview: String(page.text || "").slice(0, 220),
  }));
  return [...slideChunks, ...materialChunks].filter((chunk) => chunk.text.trim());
}

function createCitationChunk({ type, number, label, text, preview }) {
  const normalizedText = String(text || "");
  const sourceHash = hashString(`${type}:${number}:${normalizedText}`);
  return {
    id: `${type}_${number}_${sourceHash.slice(0, 8)}`,
    type,
    number,
    label,
    text: normalizedText,
    preview,
    sourceHash,
    vector: makeSearchVector(normalizedText),
  };
}

function scoreSource(text, keywords) {
  const normalized = String(text || "").toLowerCase();
  return keywords.reduce((score, keyword) => {
    if (!keyword) return score;
    const lowered = keyword.toLowerCase();
    return score + (normalized.includes(lowered) ? Math.max(1, Math.min(4, lowered.length)) : 0);
  }, 0);
}

function makeSearchVector(text) {
  const vector = {};
  tokenizeForSearch(text).forEach((token) => {
    vector[token] = (vector[token] || 0) + 1;
  });
  return vector;
}

function tokenizeForSearch(text) {
  const value = String(text || "").toLowerCase();
  const words = value.match(/[a-z0-9]{2,}|[\u3400-\u9fff]{2,}/g) || [];
  const tokens = [];
  words.forEach((word) => {
    if (/^[\u3400-\u9fff]+$/.test(word)) {
      for (let index = 0; index < word.length - 1; index += 1) {
        tokens.push(word.slice(index, index + 2));
      }
      for (let index = 0; index < word.length - 2; index += 2) {
        tokens.push(word.slice(index, index + 3));
      }
    } else {
      tokens.push(word);
    }
  });
  return tokens;
}

function cosineSimilarity(a, b) {
  let dot = 0;
  let aMag = 0;
  let bMag = 0;
  Object.entries(a || {}).forEach(([token, value]) => {
    dot += value * (b[token] || 0);
    aMag += value * value;
  });
  Object.values(b || {}).forEach((value) => {
    bMag += value * value;
  });
  if (!aMag || !bMag) return 0;
  return dot / (Math.sqrt(aMag) * Math.sqrt(bMag));
}

function hashString(value) {
  let hash = 2166136261;
  const text = String(value || "");
  for (let index = 0; index < text.length; index += 1) {
    hash ^= text.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

function recordStudentFeedback(label) {
  if (!state.studentQa.question) return;
  if (label === "helpful") state.qaMetrics.helpful += 1;
  if (label === "needs_teacher") state.qaMetrics.needsTeacher += 1;
  logAudit("學生回饋", `${label}｜${state.studentQa.question.slice(0, 80)}`);
  dom.studentAnswer.innerHTML = `<strong>${escapeHtml(state.studentQa.mode)}</strong>\n${escapeHtml(state.studentQa.answer)}\n\n已記錄回饋：${label === "helpful" ? "有幫助" : "需要老師"}`;
  renderGovernanceMetrics();
  persistState();
}

function updateQaMetrics(supported) {
  state.qaMetrics.total += 1;
  if (supported) {
    state.qaMetrics.grounded += 1;
  } else {
    state.qaMetrics.refused += 1;
  }
}

function saveVersion() {
  const inputs = state.lastLessonInputs || getLessonInputs();
  const version = {
    id: cryptoId(),
    name: `${inputs.topic}｜${new Date().toLocaleString("zh-Hant")}`,
    createdAt: new Date().toISOString(),
    inputs,
    annualPlan: structuredCloneSafe(state.annualPlan),
    slides: structuredCloneSafe(state.slides),
    script: state.script,
    assessmentBank: structuredCloneSafe(state.assessmentBank),
    materialPages: structuredCloneSafe(state.materialPages),
    slideJson: structuredCloneSafe(state.slideJson),
    auditLog: structuredCloneSafe(state.auditLog),
  };
  state.versions.unshift(version);
  state.versions = state.versions.slice(0, 12);
  logAudit("版本保存", `${version.name} 已保存`);
  renderVersions();
  renderStatus();
  markDriveBackupNeeded("版本保存");
  persistState();
}

function publishLesson() {
  const inputs = state.lastLessonInputs || getLessonInputs();
  annotateSlidesWithSourceRefs(inputs);
  logAudit("發布", `${inputs.topic} 已發布，學生端只會根據此版本回答`);
  const revision = {
    id: cryptoId(),
    status: "published",
    publishedAt: new Date().toISOString(),
    inputs: structuredCloneSafe(inputs),
    slides: structuredCloneSafe(state.slides),
    script: state.script,
    materialPages: structuredCloneSafe(state.materialPages),
    materialMeta: structuredCloneSafe(state.materialMeta),
    auditLog: structuredCloneSafe(state.auditLog),
  };
  revision.citationIndex = buildCitationIndex(revision);
  state.publishedRevision = revision;
  renderPublishedQa();
  markDriveBackupNeeded("教材發布");
  persistState();
}

function renderVersions() {
  if (!state.versions.length) {
    dom.versionList.innerHTML = emptyText("尚未儲存版本");
    dom.compareBox.textContent = "儲存版本後，可在這裡對照目前草稿與歷史快照。";
    return;
  }

  dom.versionList.innerHTML = state.versions
    .map(
      (version, index) => {
        const stats = getVersionStats(version);
        return `
        <article class="version-item">
          <strong>${escapeHtml(version.name)}</strong>
          <span>${stats.slideCount} 頁教材 · ${stats.scriptWords} 字核心講授 · ${stats.generatedLabs}/${stats.labCount} Lab · ${stats.generatedAssessments}/${stats.assessmentCount} 評核</span>
          <div class="version-actions">
            <button type="button" data-restore-version="${index}">還原</button>
            <button type="button" data-compare-version="${index}">比較</button>
          </div>
        </article>
      `;
      },
    )
    .join("");

  document.querySelectorAll("[data-restore-version]").forEach((button) => {
    button.addEventListener("click", () => restoreVersion(Number(button.dataset.restoreVersion)));
  });
  document.querySelectorAll("[data-compare-version]").forEach((button) => {
    button.addEventListener("click", () => compareVersion(Number(button.dataset.compareVersion)));
  });
}

function renderAuditLog() {
  if (!dom.auditLog) return;
  if (!state.auditLog.length) {
    dom.auditLog.innerHTML = emptyText("尚未有生成紀錄");
    return;
  }

  dom.auditLog.innerHTML = state.auditLog
    .slice(0, 20)
    .map(
      (entry) => `
        <article class="audit-item">
          <strong>${escapeHtml(entry.action)}</strong>
          <span>${escapeHtml(new Date(entry.at).toLocaleString("zh-Hant"))}</span>
          <span>${escapeHtml(entry.detail)}</span>
        </article>
      `,
    )
    .join("");
}

function renderGovernanceMetrics() {
  if (!dom.governanceMetrics) return;
  const metrics = getGovernanceMetrics();
  dom.governanceMetrics.innerHTML = [
    metricCard("角色", roleLabel(state.role), "RBAC 模擬模式"),
    metricCard("發布版本", state.publishedRevision ? "已發布" : "草稿", state.publishedRevision ? new Date(state.publishedRevision.publishedAt).toLocaleString("zh-Hant") : "學生端不可用"),
    metricCard("Citation Chunks", String(metrics.chunkCount), "發布教材可檢索來源數"),
    metricCard("QA 有據率", `${metrics.groundedRate}%`, `${state.qaMetrics.grounded}/${state.qaMetrics.total} grounded`),
    metricCard("拒答率", `${metrics.refusalRate}%`, `${state.qaMetrics.refused}/${state.qaMetrics.total} refused`),
    metricCard("需要老師", String(state.qaMetrics.needsTeacher), "學生回饋需人工介入"),
  ].join("");
}

function metricCard(label, value, hint) {
  return `<article class="governance-card"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong><small>${escapeHtml(hint)}</small></article>`;
}

function getGovernanceMetrics() {
  const total = state.qaMetrics.total || 0;
  const groundedRate = total ? Math.round((state.qaMetrics.grounded / total) * 100) : 0;
  const refusalRate = total ? Math.round((state.qaMetrics.refused / total) * 100) : 0;
  const chunkCount = state.publishedRevision?.citationIndex?.length || 0;
  return { groundedRate, refusalRate, chunkCount };
}

function logAudit(action, detail) {
  state.auditLog.unshift({
    id: cryptoId(),
    at: new Date().toISOString(),
    action,
    detail,
  });
  state.auditLog = state.auditLog.slice(0, 80);
  renderAuditLog();
}

function restoreVersion(index) {
  const version = state.versions[index];
  if (!version) return;
  state.lastLessonInputs = version.inputs;
  if ("annualPlan" in version) {
    state.annualPlan = version.annualPlan ? structuredCloneSafe(version.annualPlan) : null;
  }
  state.slides = structuredCloneSafe(version.slides);
  state.script = version.script;
  state.assessmentBank = version.assessmentBank ? structuredCloneSafe(version.assessmentBank) : null;
  state.materialPages = structuredCloneSafe(version.materialPages || state.materialPages || []);
  state.slideJson = structuredCloneSafe(version.slideJson || state.slideJson || []);
  setFormInputs(version.inputs);
  dom.assistantContext.value = buildAssistantContext();
  renderAll();
  persistState();
}

function compareVersion(index) {
  const version = state.versions[index];
  if (!version) return;
  const current = buildComparableSnapshot({
    inputs: state.lastLessonInputs || getLessonInputs(),
    annualPlan: state.annualPlan,
    slides: state.slides,
    script: state.script,
    assessmentBank: state.assessmentBank,
  });
  const previous = buildComparableSnapshot(version);
  const slideDiff = diffLists(current.slideTitles, previous.slideTitles);
  const changedSlides = countChangedSlides(current.slides, previous.slides);
  const labDiff = diffLists(current.labTitles, previous.labTitles);
  const assessmentDiff = diffLists(current.assessmentTitles, previous.assessmentTitles);
  const riskNotes = buildVersionRiskNotes(current, previous, changedSlides);

  dom.compareBox.innerHTML = `
    <div class="compare-heading">
      <div>
        <span>Version Compare</span>
        <strong>${escapeHtml(version.name)}</strong>
      </div>
      <small>${escapeHtml(new Date(version.createdAt || Date.now()).toLocaleString("zh-Hant"))}</small>
    </div>
    <div class="compare-metrics">
      ${compareMetric("PPT 頁數", current.stats.slideCount, previous.stats.slideCount)}
      ${compareMetric("核心講授字數", current.stats.scriptWords, previous.stats.scriptWords)}
      ${compareMetric("已生成 Lab", current.stats.generatedLabs, previous.stats.generatedLabs)}
      ${compareMetric("Assessment 內容", current.stats.generatedAssessments, previous.stats.generatedAssessments)}
    </div>
    <div class="compare-sections">
      ${compareSection("新增 PPT 標題", slideDiff.added)}
      ${compareSection("移除 PPT 標題", slideDiff.removed)}
      ${compareSection("Lab 變化", [...labDiff.added.map((item) => `新增：${item}`), ...labDiff.removed.map((item) => `移除：${item}`)])}
      ${compareSection("Assessment 變化", [...assessmentDiff.added.map((item) => `新增：${item}`), ...assessmentDiff.removed.map((item) => `移除：${item}`)])}
      ${compareSection("需要教師留意", riskNotes)}
    </div>
    <p class="compare-footnote">內容相同標題但 prompt / notes 改動：${escapeHtml(changedSlides)} 頁。比較後如要保留目前草稿，請按「儲存版本」。</p>
  `;
}

function getVersionStats(version) {
  const annualPlan = version?.annualPlan || {};
  const labs = Array.isArray(annualPlan.labs) ? annualPlan.labs : [];
  const assessments = Array.isArray(annualPlan.assessments) ? annualPlan.assessments : [];
  return {
    slideCount: version?.slides?.length || 0,
    scriptWords: countCoreLectureWords(version?.script || ""),
    labCount: labs.length,
    generatedLabs: labs.filter((lab) => lab.generatedContent).length,
    assessmentCount: assessments.length,
    generatedAssessments: assessments.filter((item) => item.generatedContent).length,
  };
}

function buildComparableSnapshot(source) {
  const annualPlan = source.annualPlan || {};
  const labs = Array.isArray(annualPlan.labs) ? annualPlan.labs : [];
  const assessments = Array.isArray(annualPlan.assessments) ? annualPlan.assessments : [];
  const slides = Array.isArray(source.slides) ? source.slides : [];
  return {
    inputs: source.inputs || {},
    annualPlan,
    slides,
    slideTitles: slides.map((slide) => slide.title).filter(Boolean),
    labTitles: labs.map((lab) => `${lab.id || ""} ${lab.title || ""}`.trim()).filter(Boolean),
    assessmentTitles: assessments.map((item) => `${item.type || ""} ${item.title || ""}`.trim()).filter(Boolean),
    stats: getVersionStats(source),
  };
}

function diffLists(current, previous) {
  return {
    added: current.filter((item) => !previous.includes(item)),
    removed: previous.filter((item) => !current.includes(item)),
  };
}

function countChangedSlides(currentSlides, previousSlides) {
  const previousByTitle = new Map(previousSlides.map((slide) => [slide.title, hashString(`${slide.notes || ""}\n${slide.activity || ""}`)]));
  return currentSlides.filter((slide) => {
    const previousHash = previousByTitle.get(slide.title);
    return previousHash && previousHash !== hashString(`${slide.notes || ""}\n${slide.activity || ""}`);
  }).length;
}

function buildVersionRiskNotes(current, previous, changedSlides) {
  const notes = [];
  if (current.stats.scriptWords < previous.stats.scriptWords) notes.push("目前核心講授字數較舊版本短，請確認是否仍達到目標字數。");
  if (current.stats.generatedLabs < previous.stats.generatedLabs) notes.push("目前已生成 Lab 數量較少，可能漏了部分 Lab brief / rubric。");
  if (current.stats.generatedAssessments < previous.stats.generatedAssessments) notes.push("目前 Assessment 題庫或 rubric 較舊版本少。");
  if (changedSlides > 0) notes.push(`${changedSlides} 頁 PPT 標題相同但 prompt / notes 已改，建議抽查品質。`);
  if (!notes.length) notes.push("未見明顯風險；可按需要保存目前版本。");
  return notes;
}

function compareMetric(label, current, previous) {
  const diff = Number(current || 0) - Number(previous || 0);
  const sign = diff > 0 ? "+" : "";
  const tone = diff > 0 ? "positive" : diff < 0 ? "negative" : "neutral";
  return `<div class="compare-metric ${tone}"><span>${escapeHtml(label)}</span><strong>${escapeHtml(current)}</strong><small>${escapeHtml(sign + diff)} vs 版本</small></div>`;
}

function compareSection(title, items) {
  const content = items.length
    ? `<ul>${items.slice(0, 8).map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>`
    : "<p>無</p>";
  return `<section><strong>${escapeHtml(title)}</strong>${content}</section>`;
}

function exportProjectJson() {
  logAudit("匯出", "完整專案匯出為 JSON");
  const payload = buildProjectPayload("manual-json");
  persistState();
  downloadFile("eduscript-ai-project.json", JSON.stringify(payload, null, 2), "application/json;charset=utf-8");
}

function exportAnnualJson() {
  if (!state.annualPlan) generateAnnualPlan();
  logAudit("匯出", "全年課程包匯出為 JSON");
  persistState();
  downloadFile("eduscript-ai-year-plan.json", JSON.stringify(state.annualPlan, null, 2), "application/json;charset=utf-8");
}

function exportAnnualMarkdown() {
  if (!state.annualPlan) generateAnnualPlan();
  logAudit("匯出", "全年課程包匯出為 Markdown");
  persistState();
  downloadFile("eduscript-ai-year-plan.md", buildAnnualMarkdown(state.annualPlan), "text/markdown;charset=utf-8");
}

async function importProjectJson(event) {
  const file = event.target.files?.[0];
  event.target.value = "";
  if (!file) return;

  try {
    const text = await readFileAsText(file);
    const payload = JSON.parse(text);
    applyProjectPayload(payload, `本機備份 ${file.name}`);
    dom.compareBox.textContent = `已還原備份：${file.name}`;
  } catch (error) {
    dom.compareBox.textContent = `匯入失敗：${error.message}`;
  }
}

function exportLessonMarkdown() {
  logAudit("匯出", "教材大綱與講稿匯出為 Markdown");
  persistState();
  downloadFile("eduscript-ai-lesson.md", buildMarkdown(), "text/markdown;charset=utf-8");
}

async function exportPptx() {
  if (window.location.protocol === "file:") {
    dom.compareBox.textContent = "PPTX 匯出需要以 node server.js 啟動後使用 http://localhost:4173。";
    return;
  }

  setAiBusy(true, "匯出 PPTX 中");
  try {
    const response = await fetch("/api/export-pptx", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        inputs: state.lastLessonInputs || getLessonInputs(),
        slides: state.slides,
        script: state.script,
      }),
    });

    if (!response.ok) {
      const error = await response.json().catch(() => ({}));
      throw new Error(error.error || `PPTX export failed: ${response.status}`);
    }

    const payload = await response.json();
    downloadBase64File(payload.filename || "eduscript-ai-lesson.pptx", payload.data, payload.mimeType);
    logAudit("匯出", `PowerPoint 已匯出：${payload.filename || "eduscript-ai-lesson.pptx"}`);
    dom.compareBox.textContent = "PPTX 已產生。簡報內含 AI-Assisted 標記，請教師完成最後審核。";
    persistState();
  } catch (error) {
    dom.compareBox.textContent = `PPTX 匯出失敗：${error.message}`;
  } finally {
    setAiBusy(false);
  }
}

async function exportCoursePack() {
  if (window.location.protocol === "file:") {
    dom.compareBox.textContent = "Course Pack ZIP 需要以 node server.js 啟動後使用 http://localhost:4173。";
    return;
  }

  if (!state.annualPlan) generateAnnualPlan();

  setAiBusy(true, "整理 Course Pack");
  try {
    const project = buildProjectPayload("course-pack");
    const response = await fetch("/api/export-course-pack", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        project,
        annualMarkdown: buildAnnualMarkdown(state.annualPlan),
        lessonMarkdown: buildMarkdown(),
      }),
    });

    if (!response.ok) {
      const error = await response.json().catch(() => ({}));
      throw new Error(error.error || `Course Pack export failed: ${response.status}`);
    }

    const payload = await response.json();
    downloadBase64File(payload.filename || "eduscript-ai-course-pack.zip", payload.data, payload.mimeType);
    logAudit("匯出", `Course Pack ZIP 已匯出：${payload.filename || "eduscript-ai-course-pack.zip"}`);
    dom.compareBox.textContent = "Course Pack ZIP 已產生，內含年度規劃、教材大綱、講稿、PPTX、Lab / Assessment 摘要與完整 JSON 備份。";
    persistState();
  } catch (error) {
    dom.compareBox.textContent = `Course Pack 匯出失敗：${error.message}`;
  } finally {
    setAiBusy(false);
  }
}

async function exportGammaDeck() {
  if (!state.slides.length) {
    await generateLesson();
  }

  const inputText = buildGammaDeckText();
  const additionalInstructions = buildGammaAdditionalInstructions();
  const filename = `${slugifyFilename((state.lastLessonInputs || getLessonInputs()).topic || "eduscript-gamma-deck")}-gamma-prompt.md`;

  if (window.location.protocol === "file:") {
    downloadFile(filename, inputText, "text/markdown;charset=utf-8");
    state.gamma.status = "目前是 file:// 模式，已匯出 Gamma-ready prompt。";
    renderGammaPanel();
    dom.compareBox.textContent = "Gamma API 需要以 node server.js 開啟 APP；目前已下載可貼入 Gamma 的 prompt。";
    return;
  }

  setAiBusy(true, "正在送出 Gamma PPT");
  state.gamma.status = "正在送出 Gamma 生成工作...";
  renderGammaPanel();

  try {
    const response = await fetch("/api/gamma/generate", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        inputText,
        additionalInstructions,
        numCards: state.slides.length || undefined,
        exportAs: state.gamma.exportAs || "pptx",
        audience: (state.lastLessonInputs || getLessonInputs()).audience,
      }),
    });

    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      if (response.status === 503) {
        downloadFile(filename, inputText, "text/markdown;charset=utf-8");
        state.gamma.configured = false;
        state.gamma.status = "未設定 Gamma API Key，已匯出 Gamma-ready prompt。";
        state.gamma.lastGeneration = null;
        renderGammaPanel();
        dom.compareBox.textContent = `${payload.error || "Gamma API Key 未設定"}\n\n已下載 ${filename}，之後可貼入 Gamma 生成 PPT。`;
        return;
      }
      throw new Error(payload.error || `Gamma request failed: ${response.status}`);
    }

    const finalPayload = await pollGammaGeneration(payload);
    state.gamma.configured = true;
    state.gamma.lastGeneration = finalPayload;
    state.gamma.status = buildGammaStatus(finalPayload);
    renderGammaPanel();
    dom.compareBox.textContent = formatGammaResult(finalPayload);
    logAudit("匯出", `Gamma PPT 生成工作：${finalPayload.generationId || finalPayload.id || "submitted"}`);
    persistState();
  } catch (error) {
    downloadFile(filename, inputText, "text/markdown;charset=utf-8");
    state.gamma.status = `Gamma 生成失敗，已匯出 prompt：${error.message}`;
    renderGammaPanel();
    dom.compareBox.textContent = `Gamma 生成失敗：${error.message}\n\n已下載 ${filename}，可以先手動貼入 Gamma。`;
  } finally {
    setAiBusy(false);
  }
}

async function pollGammaGeneration(initialPayload) {
  const generationId = initialPayload.generationId || initialPayload.id;
  if (!generationId) return initialPayload;

  let latest = initialPayload;
  for (let attempt = 0; attempt < 8; attempt += 1) {
    if (["completed", "failed"].includes(String(latest.status || "").toLowerCase())) {
      return latest;
    }
    await wait(5000);
    const response = await fetch(`/api/gamma/generation/${encodeURIComponent(generationId)}`);
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      throw new Error(payload.error || `Gamma status failed: ${response.status}`);
    }
    latest = { ...latest, ...payload };
    state.gamma.status = buildGammaStatus(latest);
    renderGammaPanel();
  }
  return latest;
}

function buildGammaDeckText() {
  const inputs = state.lastLessonInputs || getLessonInputs();
  const annual = state.annualPlan;
  const slides = state.slides.length ? state.slides : [];
  const overview = [
    `# ${inputs.topic || "EduScript lesson"} - Gamma PPT Deck Brief`,
    "",
    "## Course JSON",
    "```json",
    JSON.stringify(buildCourseJsonForPpt(inputs), null, 2),
    "```",
    "",
    `Subject: ${inputs.subject || annual?.inputs?.moduleTitle || "N/A"}`,
    `Audience: ${inputs.audience || annual?.inputs?.audience || "N/A"}`,
    `Duration: ${inputs.duration || "N/A"} minutes`,
    `Objective: ${inputs.objective || "N/A"}`,
    "",
    "Create one presentation card per section below. Use the slide prompt as the source of truth for visual layout, visible content, checkpoints, and teaching flow.",
    "Every card must follow the reusable template fields: slide_title, slide_subtitle, slide_body, speaker_notes, suggested_visual, suggested_layout, presenter_cues, fact_check_points.",
    "Keep slide_body concise; move answer keys, fallback lines, and demo caveats to speaker_notes.",
  ];

  if (annual?.metrics) {
    overview.push(
      "",
      "## Academic-Year Context",
      `Lecture hours: ${annual.metrics.lectureHours}`,
      `CA Lab hours: ${annual.metrics.labHours}`,
      `Assessment hours: ${annual.metrics.assessmentHours}`,
      `PPT slides planned: ${annual.metrics.pptSlides}`,
    );
  }

  const slideText = slides.map((slide, index) => {
    return [
      "",
      "---",
      "",
      `## Card ${index + 1}: ${slide.title || `Slide ${index + 1}`}`,
      `Slide type: ${slide.slideType || "content"}`,
      `Timing: ${slide.minutes || ""} minutes`,
      `Teaching event: ${slide.event || ""}`,
      `Bloom level: ${slide.bloom || ""}`,
      `Suggested layout: ${slide.suggestedLayout || ""}`,
      `Suggested visual: ${slide.suggestedVisual || ""}`,
      "",
      "### Visible Content And Layout Prompt",
      slide.notes || slide.activity || "",
      "",
      "### Teacher Checkpoint",
      slide.activity || "Add one quick concept check question.",
      "",
      "### Speaker Notes",
      slide.speakerNotes || "Put answer key, transitions, demo fallback, and teacher-only caveats here.",
      "",
      "### Fact Check Points",
      (slide.factCheckPoints || []).map((item) => `- ${item}`).join("\n") || "- Check versions and official sources before publishing.",
    ].join("\n");
  });

  if (state.script) {
    slideText.push(
      "",
      "---",
      "",
      "## Speaker Notes Reference",
      "Use this only as teaching-note context. Do not place all script text on slides.",
      state.script.slice(0, 12000),
    );
  }

  return [...overview, ...slideText].join("\n");
}

function buildGammaAdditionalInstructions() {
  return [
    "Generate a polished 16:9 professional teaching presentation in Traditional Chinese.",
    "Preserve one card per Card section when possible.",
    "Keep slide text concise, but make diagrams/checklists/demos clear enough for students to follow.",
    "Use cloud, Linux terminal, YAML, Kubernetes architecture, lab evidence, and assessment visuals where relevant.",
    "Return an export file if exportAs is enabled.",
  ].join(" ");
}

function buildGammaStatus(payload) {
  const status = String(payload.status || "submitted");
  const exportUrl = payload.exportUrl || "";
  const gammaUrl = payload.gammaUrl || payload.url || "";
  if (status === "completed" && (exportUrl || gammaUrl)) {
    return `Gamma 已完成：${exportUrl ? "PPTX/PDF 匯出連結已生成" : "Gamma 文件已生成"}。`;
  }
  if (status === "failed") return "Gamma 生成失敗，請查看下方錯誤。";
  return `Gamma 生成工作已送出：${status}`;
}

function formatGammaResult(payload) {
  const lines = [
    "Gamma PPT 生成結果",
    `Generation ID: ${payload.generationId || payload.id || "N/A"}`,
    `Status: ${payload.status || "submitted"}`,
  ];
  if (payload.gammaUrl || payload.url) lines.push(`Gamma URL: ${payload.gammaUrl || payload.url}`);
  if (payload.exportUrl) lines.push(`Export URL: ${payload.exportUrl}`);
  if (payload.credits) {
    lines.push(`Credits: deducted ${payload.credits.deducted ?? "?"}, remaining ${payload.credits.remaining ?? "?"}`);
  }
  if (!payload.exportUrl && !payload.gammaUrl && payload.status !== "completed") {
    lines.push("Gamma 仍在處理中；稍後可用 Generation ID 查狀態。");
  }
  return lines.join("\n");
}

function renderGammaPanel() {
  if (!dom.gammaStatus) return;
  dom.gammaStatus.textContent = state.gamma.status || "未設定 Gamma API Key，可先匯出 Gamma-ready prompt。";
  dom.gammaStatus.classList.toggle("strong", Boolean(state.gamma.configured));
  dom.gammaStatus.classList.toggle("error", !state.gamma.configured);
  if (dom.gammaResult) {
    dom.gammaResult.textContent = state.gamma.lastGeneration
      ? formatGammaResult(state.gamma.lastGeneration)
      : "按「Gamma PPT」會先使用目前每頁 PPT Prompt 建立 Gamma deck brief；未設定 API Key 時會下載 Markdown prompt。";
  }
}

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function copyPrompt() {
  const prompt = buildPrompt();
  try {
    await navigator.clipboard.writeText(prompt);
    dom.compareBox.textContent = "Prompt 已複製。";
  } catch {
    dom.compareBox.textContent = prompt;
  }
  logAudit("匯出", "AI Prompt 已複製或顯示");
  persistState();
}

function buildProjectPayload(source = "manual") {
  const inputs = state.lastLessonInputs || getLessonInputs();
  return {
    schema: "eduscript-ai-project",
    schemaVersion: 5,
    app: "EduScript AI Studio",
    source,
    exportedAt: new Date().toISOString(),
    backupLabel: `${inputs.topic || "未命名教材"}｜${inputs.subject || "未分類"}`,
    annualPlan: state.annualPlan,
    inputs,
    slides: state.slides,
    questions: state.questions,
    materialMeta: state.materialMeta,
    materialPages: state.materialPages,
    slideJson: state.slideJson,
    materialText: dom.materialText.value,
    assistantContext: dom.assistantContext.value,
    script: state.script,
    versions: state.versions,
    messages: state.messages,
    auditLog: state.auditLog,
    publishedRevision: state.publishedRevision,
    studentQa: state.studentQa,
    qaMetrics: state.qaMetrics,
    assessmentBank: state.assessmentBank,
    gamma: state.gamma,
    interviewAnswers: state.interviewAnswers,
    role: state.role,
  };
}

function applyProjectPayload(payload, sourceLabel = "備份") {
  if (!payload || typeof payload !== "object") {
    throw new Error("備份檔格式不正確");
  }
  if (!Array.isArray(payload.slides) && !payload.script && !payload.inputs) {
    throw new Error("找不到教材內容");
  }

  state.annualPlan = payload.annualPlan || null;
  state.slides = payload.slides || [];
  state.script = payload.script || "";
  state.questions = payload.questions || [];
  state.materialPages = payload.materialPages || [];
  state.materialMeta = payload.materialMeta || null;
  state.slideJson = payload.slideJson || payload.materialMeta?.slideJson || [];
  state.versions = payload.versions || [];
  state.messages = payload.messages || state.messages || [];
  state.auditLog = payload.auditLog || [];
  state.publishedRevision = payload.publishedRevision || null;
  state.assessmentBank = payload.assessmentBank || payload.annualPlan?.assessmentBank || null;
  state.studentQa = payload.studentQa || {
    question: "",
    answer: "",
    mode: "未提問",
    sources: [],
  };
  state.qaMetrics = payload.qaMetrics || {
    total: 0,
    grounded: 0,
    refused: 0,
    helpful: 0,
    needsTeacher: 0,
  };
  state.gamma.lastGeneration = payload.gamma?.lastGeneration || state.gamma.lastGeneration;
  state.interviewAnswers = payload.interviewAnswers || payload.inputs?.interviewAnswers || "";
  state.role = payload.role || state.role || "teacher";
  state.lastLessonInputs = payload.inputs || payload.lastLessonInputs || state.lastLessonInputs || getLessonInputs();

  setFormInputs(state.lastLessonInputs);
  if (dom.questionAnswer) dom.questionAnswer.value = state.interviewAnswers;
  dom.materialText.value = payload.materialText || buildMaterialFromSlides();
  dom.assistantContext.value = payload.assistantContext || buildAssistantContext();
  if (state.materialMeta) {
    setMaterialStatus(`已載入 ${state.materialMeta.filename}：${state.materialPages.length || 1} 個片段`, true);
  } else {
    setMaterialStatus("可貼上教材或上傳檔案，生成講稿時會優先使用。");
  }

  logAudit("還原備份", `${sourceLabel} 已套用`);
  renderAll();
  persistState();
}

async function connectGoogleDrive() {
  const clientId = clean(dom.driveClientId.value);
  if (!clientId) {
    setDriveStatus("請先填入 Google OAuth Client ID。", false);
    return;
  }
  if (window.location.protocol === "file:") {
    setDriveStatus("Google Drive 登入需要用 http://localhost:4173 開啟；請先執行 node server.js。", false);
    return;
  }

  try {
    setDriveBusy(true, "正在連接 Google Drive...");
    await requestDriveToken(clientId);
    state.drive.clientId = clientId;
    state.drive.connected = true;
    setDriveStatus("已連接 Google Drive。", true);
    persistDriveSettings();
    logAudit("雲端備份", "Google Drive 已連接");
    renderDrivePanel();
  } catch (error) {
    state.drive.connected = false;
    setDriveStatus(`Google Drive 連接失敗：${error.message}`, false);
  } finally {
    setDriveBusy(false);
  }
}

async function backupToGoogleDrive(options = {}) {
  try {
    await ensureDriveAccess({ interactive: !options.automatic });
    setDriveBusy(true, options.automatic ? "正在自動備份到 Google Drive..." : "正在備份到 Google Drive...");
    const reason = options.reason || state.drive.pendingReason || "手動備份";
    const payload = buildProjectPayload(options.automatic ? "google-drive-auto" : "google-drive");
    const metadata = {
      name: buildDriveBackupFilename(payload),
      mimeType: "application/json",
      description: "EduScript AI Studio project backup",
      appProperties: {
        app: "eduscript-ai-studio",
        schemaVersion: String(payload.schemaVersion || 1),
      },
    };
    const file = await uploadDriveJson(metadata, payload);
    state.drive.lastBackup = file;
    state.drive.lastBackupAt = new Date().toISOString();
    state.drive.pendingReason = "";
    setDriveStatus(`已備份到 Google Drive：${file.name}`, true);
    logAudit("雲端備份", `${options.automatic ? "自動" : "手動"}備份完成：${file.name}（${reason}）`);
    persistDriveSettings();
    await listGoogleDriveBackups(false);
    setDriveStatus(`已備份到 Google Drive：${file.name}`, true);
  } catch (error) {
    setDriveStatus(`${options.automatic ? "自動" : ""}備份失敗：${error.message}`, state.drive.connected);
  } finally {
    setDriveBusy(false);
  }
}

async function listGoogleDriveBackups(showStatus = true) {
  try {
    await ensureDriveAccess();
    setDriveBusy(true, "正在讀取 Google Drive 備份...");
    const query = [
      `name contains '${DRIVE_BACKUP_PREFIX}'`,
      "mimeType = 'application/json'",
      "trashed = false",
    ].join(" and ");
    const url = new URL("https://www.googleapis.com/drive/v3/files");
    url.searchParams.set("q", query);
    url.searchParams.set("pageSize", "10");
    url.searchParams.set("orderBy", "modifiedTime desc");
    url.searchParams.set("fields", "files(id,name,modifiedTime,size,webViewLink)");
    const data = await driveFetch(url.toString());
    state.drive.backups = data.files || [];
    if (showStatus) {
      setDriveStatus(`找到 ${state.drive.backups.length} 個 Google Drive 備份。`, true);
    }
    renderDrivePanel();
    return state.drive.backups;
  } catch (error) {
    setDriveStatus(`讀取備份失敗：${error.message}`, false);
    return [];
  } finally {
    setDriveBusy(false);
  }
}

async function restoreLatestGoogleDriveBackup() {
  const backups = state.drive.backups.length ? state.drive.backups : await listGoogleDriveBackups(false);
  if (!backups.length) {
    setDriveStatus("Google Drive 未找到可還原的備份。", false);
    return;
  }
  await restoreGoogleDriveBackup(backups[0].id, backups[0].name);
}

async function restoreGoogleDriveBackup(fileId, filename = "Google Drive 備份") {
  const ok = window.confirm(`確定要用「${filename}」覆蓋目前工作台？`);
  if (!ok) return;

  try {
    await ensureDriveAccess();
    setDriveBusy(true, "正在還原 Google Drive 備份...");
    const response = await fetch(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(fileId)}?alt=media`, {
      headers: {
        Authorization: `Bearer ${driveAccessToken}`,
      },
    });
    if (!response.ok) throw new Error(await driveErrorMessage(response));
    const payload = await response.json();
    applyProjectPayload(payload, `Google Drive：${filename}`);
    setDriveStatus(`已還原：${filename}`, true);
    dom.compareBox.textContent = `已從 Google Drive 還原：${filename}`;
  } catch (error) {
    setDriveStatus(`還原失敗：${error.message}`, false);
  } finally {
    setDriveBusy(false);
  }
}

async function ensureDriveAccess(options = {}) {
  const interactive = options.interactive !== false;
  const clientId = clean(dom.driveClientId.value || state.drive.clientId);
  if (!clientId) throw new Error("請先填入 Google OAuth Client ID。");
  if (!driveAccessToken && !interactive) throw new Error("請先按「連接 Drive」取得授權。");
  if (!driveAccessToken) await connectGoogleDrive();
  if (!driveAccessToken) throw new Error("尚未取得 Google Drive 授權。");
}

async function requestDriveToken(clientId) {
  await loadGoogleIdentityScript();
  return new Promise((resolve, reject) => {
    driveTokenClient = window.google.accounts.oauth2.initTokenClient({
      client_id: clientId,
      scope: DRIVE_SCOPE,
      callback: (response) => {
        if (response.error) {
          reject(new Error(response.error_description || response.error));
          return;
        }
        driveAccessToken = response.access_token;
        resolve(response);
      },
    });
    driveTokenClient.requestAccessToken({ prompt: driveAccessToken ? "" : "consent" });
  });
}

function loadGoogleIdentityScript() {
  if (window.google?.accounts?.oauth2) return Promise.resolve();
  if (googleIdentityScriptPromise) return googleIdentityScriptPromise;
  googleIdentityScriptPromise = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = GOOGLE_IDENTITY_SCRIPT;
    script.async = true;
    script.defer = true;
    script.onload = resolve;
    script.onerror = () => reject(new Error("無法載入 Google Identity Services。"));
    document.head.appendChild(script);
  });
  return googleIdentityScriptPromise;
}

async function uploadDriveJson(metadata, payload) {
  const boundary = `eduscript_${Date.now()}`;
  const body = [
    `--${boundary}`,
    "Content-Type: application/json; charset=UTF-8",
    "",
    JSON.stringify(metadata),
    `--${boundary}`,
    "Content-Type: application/json; charset=UTF-8",
    "",
    JSON.stringify(payload, null, 2),
    `--${boundary}--`,
  ].join("\r\n");

  const response = await fetch("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,modifiedTime,size,webViewLink", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${driveAccessToken}`,
      "Content-Type": `multipart/related; boundary=${boundary}`,
    },
    body,
  });
  if (!response.ok) throw new Error(await driveErrorMessage(response));
  return response.json();
}

async function driveFetch(url, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: {
      ...(options.headers || {}),
      Authorization: `Bearer ${driveAccessToken}`,
    },
  });
  if (!response.ok) throw new Error(await driveErrorMessage(response));
  return response.json();
}

async function driveErrorMessage(response) {
  const payload = await response.json().catch(() => null);
  return payload?.error?.message || `Google Drive API ${response.status}`;
}

function buildDriveBackupFilename(payload) {
  const topic = slugifyFilename(payload.inputs?.topic || "lesson");
  const stamp = new Date().toISOString().replace(/[:.]/g, "-");
  return `${DRIVE_BACKUP_PREFIX}${topic}-${stamp}.json`;
}

function renderDrivePanel() {
  if (!dom.driveStatus) return;
  dom.driveClientId.value = state.drive.clientId || "";
  dom.autoDriveBackup.checked = Boolean(state.drive.autoBackup);
  dom.driveStatus.textContent = state.drive.status || "未連接 Google Drive";
  dom.driveStatus.classList.toggle("strong", Boolean(state.drive.connected));
  dom.driveStatus.classList.toggle("error", !state.drive.connected && state.drive.status !== "未連接 Google Drive");
  dom.driveSyncMeta.innerHTML = buildDriveSyncMeta();

  if (!state.drive.backups.length) {
    dom.driveBackupList.innerHTML = emptyText("尚未列出雲端備份");
    return;
  }

  dom.driveBackupList.innerHTML = state.drive.backups
    .map(
      (file) => `
        <div class="drive-backup-item">
          <div>
            <strong>${escapeHtml(file.name)}</strong>
            <span>${escapeHtml(new Date(file.modifiedTime).toLocaleString("zh-Hant"))}${file.size ? ` · ${formatBytes(file.size)}` : ""}</span>
          </div>
          <button class="action-button ghost" type="button" data-drive-restore="${escapeHtml(file.id)}">還原</button>
        </div>
      `,
    )
    .join("");

  dom.driveBackupList.querySelectorAll("[data-drive-restore]").forEach((button) => {
    button.addEventListener("click", () => {
      const file = state.drive.backups.find((backup) => backup.id === button.dataset.driveRestore);
      restoreGoogleDriveBackup(button.dataset.driveRestore, file?.name || "Google Drive 備份");
    });
  });
}

function buildDriveSyncMeta() {
  const status = state.drive.pendingReason
    ? `待備份：${state.drive.pendingReason}`
    : state.drive.lastBackupAt
      ? "已同步"
      : "尚未備份";
  return [
    ["備份狀態", status],
    ["最近本機更新", state.drive.lastLocalChange ? new Date(state.drive.lastLocalChange).toLocaleString("zh-Hant") : "未記錄"],
    ["最近雲端備份", state.drive.lastBackupAt ? new Date(state.drive.lastBackupAt).toLocaleString("zh-Hant") : "未備份"],
  ]
    .map(([label, value]) => `<div><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`)
    .join("");
}

function markDriveBackupNeeded(reason) {
  state.drive.pendingReason = reason;
  state.drive.lastLocalChange = new Date().toISOString();
  if (state.drive.autoBackup && driveAccessToken && !state.drive.busy) {
    scheduleDriveAutoBackup(reason);
  } else if (state.drive.connected || state.drive.clientId) {
    setDriveStatus(`有未備份更新：${reason}`, state.drive.connected);
  }
  persistDriveSettings();
  renderDrivePanel();
}

function scheduleDriveAutoBackup(reason) {
  clearTimeout(driveAutoBackupTimer);
  setDriveStatus(`已排程自動備份：${reason}`, true);
  driveAutoBackupTimer = setTimeout(() => {
    backupToGoogleDrive({ automatic: true, reason });
  }, 1500);
}

function setDriveBusy(busy, message = "") {
  state.drive.busy = busy;
  if (message) state.drive.status = message;
  renderDrivePanel();
  applyRolePermissions();
}

function setDriveStatus(message, connected) {
  state.drive.status = message;
  if (typeof connected === "boolean") state.drive.connected = connected;
  renderDrivePanel();
}

function persistDriveSettings() {
  localStorage.setItem(
    DRIVE_SETTINGS_KEY,
    JSON.stringify({
      clientId: state.drive.clientId,
      lastBackup: state.drive.lastBackup,
      lastBackupAt: state.drive.lastBackupAt,
      lastLocalChange: state.drive.lastLocalChange,
      pendingReason: state.drive.pendingReason,
      autoBackup: state.drive.autoBackup,
    }),
  );
}

function restoreDriveSettings() {
  try {
    const raw = localStorage.getItem(DRIVE_SETTINGS_KEY);
    if (!raw) return;
    const payload = JSON.parse(raw);
    state.drive.clientId = payload.clientId || "";
    state.drive.lastBackup = payload.lastBackup || null;
    state.drive.lastBackupAt = payload.lastBackupAt || null;
    state.drive.lastLocalChange = payload.lastLocalChange || null;
    state.drive.pendingReason = payload.pendingReason || "";
    state.drive.autoBackup = Boolean(payload.autoBackup);
    state.drive.status = state.drive.clientId ? "可連接 Google Drive" : "未連接 Google Drive";
  } catch {
    localStorage.removeItem(DRIVE_SETTINGS_KEY);
  }
}

function clearProject() {
  const ok = window.confirm("確定要清除本機工作台內容？");
  if (!ok) return;
  localStorage.removeItem(STORAGE_KEY);
  state.slides = [];
  state.annualPlan = null;
  state.script = "";
  state.questions = [];
  state.materialPages = [];
  state.materialMeta = null;
  state.slideJson = [];
  state.versions = [];
  state.messages = [];
  state.auditLog = [];
  state.publishedRevision = null;
  state.assessmentBank = null;
  state.studentQa = {
    question: "",
    answer: "",
    mode: "未提問",
    sources: [],
  };
  state.qaMetrics = {
    total: 0,
    grounded: 0,
    refused: 0,
    helpful: 0,
    needsTeacher: 0,
  };
  state.interviewAnswers = "";
  if (dom.questionAnswer) dom.questionAnswer.value = "";
  state.role = "teacher";
  state.lastLessonInputs = null;
  state.budget = null;
  syncDefaultCoreMinutes(true);
  generateLesson();
}

async function loadDemoProject() {
  setFormInputs({
    topic: "光合作用與能量轉換",
    subject: "高中生物",
    audience: "中四學生",
    duration: 45,
    style: "探究式",
    objective: "學生能解釋光反應與暗反應的角色，並分析光合作用如何連接生態系能量流。",
    context: "學生已學過細胞構造，但對能量轉換與反應位置仍容易混淆。",
    bloom: ["remember", "understand", "analyze", "evaluate"],
  });
  await generateLesson();
  dom.startPage.value = "4";
  dom.scriptMinutes.value = "18";
  syncDefaultCoreMinutes(true);
  await generateScript();
  saveVersion();
}

function renderStatus() {
  dom.slideCount.textContent = String(state.slides.length);
  dom.durationStatus.textContent = String((state.lastLessonInputs || getLessonInputs()).duration);
  dom.scriptWordStatus.textContent = String(countCoreLectureWords(state.script));
  dom.versionStatus.textContent = String(state.versions.length);
  dom.publishStatus.textContent = state.publishedRevision ? "已發布" : "草稿";
  dom.groundedRateStatus.textContent = `${getGovernanceMetrics().groundedRate}%`;
}

function renderAiStatus() {
  if (!dom.aiStatus) return;
  const status = state.ai.busy ? "checking" : state.ai.enabled ? "online" : state.ai.checked ? "fallback" : "checking";
  const providerName = formatAiProviderName(state.ai.provider);
  const label = state.ai.busy ? state.ai.message : state.ai.enabled ? providerName : state.ai.message || "AI 未設定";
  const detail = state.ai.enabled
    ? buildAiStatusDetail()
    : state.ai.checked
      ? "所有生成位置必須連線 AI"
      : "檢查中";
  const checkedAt = state.ai.lastCheckedAt
    ? new Date(state.ai.lastCheckedAt).toLocaleTimeString("zh-Hant", { hour: "2-digit", minute: "2-digit" })
    : "--:--";

  dom.aiStatus.classList.toggle("ai-online", status === "online");
  dom.aiStatus.classList.toggle("ai-fallback", status === "fallback");
  dom.aiStatus.classList.toggle("ai-checking", status === "checking");
  dom.aiStatus.innerHTML = `
    <span class="ai-status-light" aria-hidden="true"></span>
    <div class="ai-status-copy">
      <span>AI 狀態燈</span>
      <strong>${escapeHtml(label)}</strong>
      <small>${escapeHtml(detail)} · ${escapeHtml(checkedAt)}</small>
    </div>
    <button class="ai-refresh" type="button" data-ai-refresh aria-label="重新檢查 AI 狀態" title="重新檢查 AI 狀態">↻</button>
  `;
}

function setAiBusy(busy, message = "") {
  state.ai.busy = busy;
  if (message) state.ai.message = message;
  if (!busy && state.ai.enabled) {
    state.ai.message = `${formatAiProviderName(state.ai.provider)} 已連線`;
  }
  if (!busy && !state.ai.enabled && state.ai.checked) {
    state.ai.message = "AI 未設定";
  }
  renderAiStatus();
}

function buildAiStatusDetail() {
  const parts = [state.ai.model || "server model"];
  if (state.ai.provider === "openai-compatible" && state.ai.openAiCompatible) {
    parts.push(state.ai.openAiCompatible.baseUrl || "OpenAI-compatible");
    parts.push(`temp ${state.ai.openAiCompatible.temperature}`);
    parts.push(`max ${state.ai.openAiCompatible.maxTokens}`);
  }
  if (state.ai.provider === "gemini" && state.ai.geminiModelTier) {
    parts.push(state.ai.geminiModelTier.tier === "high" ? "高階模型" : state.ai.geminiModelTier.tier === "fast" ? "快速模型" : "模型等級待確認");
  }
  if (state.ai.provider === "gemini" && state.ai.geminiGeneration) {
    parts.push(`temp ${state.ai.geminiGeneration.temperature}`);
    parts.push(`script max ${state.ai.geminiGeneration.scriptMaxOutputTokens}`);
  }
  return parts.join(" · ");
}

function formatAiProviderName(provider) {
  const normalized = String(provider || "").toLowerCase();
  if (normalized === "gemini") return "Gemini";
  if (normalized === "openai-compatible") return "Qwen / OpenAI-compatible";
  return "AI";
}

function buildSlideTitle(event, topic, index) {
  const titleMap = {
    引起動機: `${topic} 的第一個衝突問題`,
    提示目標: `本課要完成的學習任務`,
    喚起舊知: `從已知概念連到 ${topic}`,
    呈現內容: `${topic} 的核心模型`,
    提供引導: `教師示範與思考路徑`,
    引發表現: `學生分析任務`,
    提供回饋: `常見誤解與修正`,
    評量學習: `Exit Ticket 與短答評量`,
    促進保留: `課後遷移與下一課預告`,
  };
  return titleMap[event] || `${index + 1}. ${topic}`;
}

function buildActivity(event, inputs, bloom) {
  const activityMap = {
    引起動機: `以反直覺問題開場，快速收集學生原始想法。`,
    提示目標: `把目標轉成學生可自評的 ${bloom.verb} 任務。`,
    喚起舊知: `用一題快速診斷題找出先備概念落差。`,
    呈現內容: `用圖解、例子與分段說明拆解 ${inputs.topic}。`,
    提供引導: `示範如何從題目線索推到正確概念。`,
    引發表現: `安排同儕討論，要求學生交出一個分析結果。`,
    提供回饋: `回收學生答案，針對誤解給出糾正性回饋。`,
    評量學習: `用短答或選擇題確認能否遷移。`,
    促進保留: `將本課連到下一課或生活情境。`,
  };
  return activityMap[event] || bloom.strategy;
}

function buildSlideNotes(event, inputs, bloom, minutes) {
  return [
    `教學目的：${event}，預計 ${formatNumber(minutes)} 分鐘。`,
    `學生任務：以「${bloom.verb}」處理「${inputs.topic}」。`,
    `教師講法：用${inputs.style}語氣，先點出 ${inputs.context || "學生目前的理解狀態"}。`,
    inputs.interviewAnswers ? `教師追問回答：${inputs.interviewAnswers}` : "",
    `內容重點：${bloom.strategy}`,
    `檢核方式：請學生用一句話說明本頁最重要的概念，教師即時標記需補救的答案。`,
  ].filter(Boolean).join("\n");
}

function buildPptSlidePrompt({
  title,
  event,
  inputs,
  bloom,
  minutes,
  activity,
  sourceNotes,
  slideType = "content",
  template = pptTemplateCatalog.content,
  suggestedLayout = "",
  suggestedVisual = "",
  factCheckPoints = [],
  speakerNotes = "",
}) {
  const visual = suggestedVisual || inferPptVisualPrompt(title, inputs.topic, event);
  const courseJson = buildCourseJsonForPpt(inputs);
  return `PPT Slide Compiler Prompt

你是一位資深教學設計師、技術講師與 PowerPoint 編輯。請只用繁體中文輸出，保留必要英文技術名詞、kubectl 指令、YAML 欄位與產品名稱。

【課程中繼資料 course_json】
${JSON.stringify(courseJson, null, 2)}

【本頁設定】
- slide_no: auto
- slide_type: ${slideType} / ${template.label || "內容講解"}
- slide_goal: ${title}
- teaching_event: ${event}
- bloom_level: ${bloom.label}
- time_budget_min: ${formatNumber(minutes)}
- style: ${inputs.style}
- layout_preference: ${suggestedLayout || template.layout}
- visual_preference: ${visual}
- linked_assessment: ${activity}

【輸出格式】
1. slide_title:
2. slide_subtitle:
3. slide_body:
4. speaker_notes:
5. suggested_visual:
6. suggested_layout:
7. presenter_cues:
8. fact_check_points:

【硬性限制】
- slide_body 請控制在 ${template.bodyBudget || 80} 字內；最多 4 個 bullets。
- speaker_notes 請控制在 ${template.notesBudget || 160} 字內；答案鍵、轉場語、demo fallback 放 notes，不要塞進 slide_body。
- 每張投影片必須有唯一標題、清楚閱讀順序、足夠留白、18pt 以上字級、sans serif 字型與高對比。
- 資訊性圖像要提供 alt text；裝飾性圖像請標示 decorative。
- 不要杜撰版本、考試權重、CLI 指令或 YAML 欄位。
- 若是 exam-oriented 或技術實作頁，務必標出常考點、易錯點或驗收條件。
- 若內容資訊不足，請標示「推定補充」或「需教師確認」。

【本頁素材】
${sourceNotes}

【建議講者備忘稿】
${speakerNotes}

【需查核事項】
${factCheckPoints.map((item) => `- ${item}`).join("\n")}`;
}

function buildCourseJsonForPpt(inputs) {
  return {
    title: inputs.topic,
    subject_domain: inputs.subject,
    audience_profile: inputs.audience,
    duration_min: inputs.duration,
    style: inputs.style,
    objectives: splitPlanItems(inputs.objective).slice(0, 6),
    prerequisites: splitPlanItems(inputs.context).slice(0, 8),
    bloom_levels: inputs.bloom || [],
    teacher_interview_answers: inputs.interviewAnswers || "",
    source_completeness: inputs.context ? "provided" : "partial_or_missing",
    workflow: "course interview -> course_json -> template catalog -> slide prompt compiler",
  };
}

function inferPptVisualPrompt(title, topic, event) {
  const text = `${title} ${topic} ${event}`.toLowerCase();
  if (text.includes("linux") || text.includes("ubuntu")) return "畫一個 VM / server readiness checklist，加上 CPU、RAM、network、SSH 狀態標記。";
  if (text.includes("ansible")) return "畫 control node 指向多個 target nodes 的 automation flow，旁邊放 playbook snippet。";
  if (text.includes("architecture") || text.includes("control plane")) return "畫 Kubernetes control plane、worker node、etcd、scheduler、API server 的架構圖。";
  if (text.includes("kubectl") || text.includes("yaml")) return "用 terminal command + YAML block 的雙欄 layout，標出 command 對應的 resource。";
  if (text.includes("service") || text.includes("ingress") || text.includes("network")) return "畫 client 到 Ingress / Service / Pod 的 traffic path，標出 public endpoint。";
  if (text.includes("storage") || text.includes("pvc")) return "畫 Pod、PVC、PV、StorageClass 的 binding relationship。";
  if (text.includes("security") || text.includes("rbac")) return "畫 Subject、Role、RoleBinding、Namespace 的 permission map。";
  if (text.includes("troubleshooting")) return "畫 diagnose ladder：get、describe、logs、events、exec。";
  if (text.includes("eks") || text.includes("aws")) return "畫 AWS EKS managed control plane、node group、IAM、public service 的 cloud diagram。";
  if (text.includes("rancher")) return "畫 multi-cluster dashboard view，突出 enterprise management、policy、monitoring。";
  return "畫一個由 concept、demo、student checkpoint 組成的三段式教學流程。";
}

function updatePptPromptWithFeedback(slide, feedback, additions) {
  return `${slide.notes}

Revision Request

教師修改意見：${feedback}
PPT 修改方向：
${additions.length ? additions.map((item) => `- ${item}`).join("\n") : "- 調整 visible text、diagram、checkpoint，使投影片更清楚、更可教。"}
- 保持這頁是 PPT prompt，不要改成逐字講稿。
- 如需講稿，請先生成 / 匯出 PPT，再到「進度講稿」頁用 PPT 文字生成每堂講稿。`;
}

function buildCoreScript(fragments, inputs) {
  const source = fragments.length
    ? fragments
    : [
        `第一步，建立 ${inputs.topic} 的核心定義。`,
        `第二步，說明它和學生既有知識的關係。`,
        `第三步，加入比較、判斷與應用。`,
      ];

  return source
    .slice(0, 5)
    .map((fragment, index) => {
      const cue = index === 0 ? "先建立共同語言" : index === 1 ? "再處理最容易混淆的地方" : "最後把概念放回題境";
      return `${index + 1}. ${cue}：${fragment}\n教師口語提示：這裡不要只要求學生記住答案，要追問「為什麼這個線索可以支持你的判斷」。`;
    })
    .join("\n\n");
}

function buildAssistantContext() {
  const inputs = state.lastLessonInputs || getLessonInputs();
  const slideSummary = state.slides
    .slice(0, 5)
    .map((slide) => `第 ${slide.number} 頁 ${slide.title}`)
    .join("；");
  return `課題：${inputs.topic}。對象：${inputs.audience}。學習目標：${inputs.objective}。教師追問回答：${inputs.interviewAnswers || "尚未提供"}。目前教材：${slideSummary}`;
}

function buildMaterialFromSlides() {
  return state.slides
    .map((slide) => `第 ${slide.number} 頁：${slide.title}\n${slide.notes}`)
    .join("\n\n");
}

function sendSlidesToScriptMaterial() {
  if (!state.slides.length) return;
  dom.materialText.value = buildMaterialFromSlides();
  dom.assistantContext.value = buildAssistantContext();
  switchView("script");
  setMaterialStatus("已把目前 PPT prompts 放入講稿素材。你也可以改為上傳真正生成後的 PPTX，再生成每堂講稿。", true);
  logAudit("講稿素材", "目前 PPT prompt 已送到進度講稿頁");
  persistState();
}

async function parseMaterialFile(file) {
  if (window.location.protocol === "file:") {
    return null;
  }

  const data = await readFileAsDataUrl(file);
  const response = await fetch("/api/parse-material", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      filename: file.name,
      mimeType: file.type,
      data,
    }),
  });

  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    throw new Error(error.error || `Material parse failed: ${response.status}`);
  }

  return response.json();
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(reader.error || new Error("File read failed"));
    reader.readAsDataURL(file);
  });
}

function readFileAsText(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(reader.error || new Error("File read failed"));
    reader.readAsText(file, "UTF-8");
  });
}

function isPlainTextFile(file) {
  return /\.(txt|md|csv|json)$/i.test(file.name) || file.type.startsWith("text/");
}

function setMaterialStatus(message, strong = false) {
  if (!dom.materialStatus) return;
  dom.materialStatus.textContent = message;
  dom.materialStatus.classList.toggle("strong", strong);
}

function textToPages(text) {
  const blocks = String(text || "")
    .split(/\n{2,}|第\s*\d+\s*頁|Slide\s*\d+/i)
    .map((block) => clean(block))
    .filter(Boolean);
  const source = blocks.length ? blocks : chunkTextForClient(text, 1000);
  return source.map((block, index) => ({
    number: index + 1,
    title: firstClientLine(block) || `文字片段 ${index + 1}`,
    text: block,
  }));
}

function materialFragments(material, startPage, inputs) {
  const pages = state.materialPages.length ? state.materialPages : textToPages(material);
  if (!pages.length) return [];

  const startIndex = clamp(startPage - 1, 0, Math.max(0, pages.length - 1));
  const localWindow = pages.slice(startIndex, startIndex + 12);
  const keywords = extractKeywords(`${inputs.topic} ${inputs.objective} ${inputs.context}`);
  const scored = localWindow.map((page, offset) => ({
    page,
    index: startIndex + offset,
    score: scoreMaterialPage(page, keywords, offset),
  }));

  return scored
    .sort((a, b) => b.score - a.score)
    .slice(0, 5)
    .sort((a, b) => a.index - b.index)
    .map(({ page }) => `第 ${page.number} 頁：${page.title}\n${page.text.slice(0, 650)}`);
}

function scoreMaterialPage(page, keywords, offset) {
  const text = `${page.title} ${page.text}`.toLowerCase();
  const keywordScore = keywords.reduce((sum, keyword) => sum + (text.includes(keyword.toLowerCase()) ? 3 : 0), 0);
  const proximityScore = Math.max(0, 8 - offset);
  const densityScore = Math.min(4, Math.round(page.text.length / 180));
  return keywordScore + proximityScore + densityScore;
}

function extractKeywords(text) {
  const value = String(text || "");
  const raw = value
    .replace(/[，。！？、；：「」『』（）()[\]{}]/g, " ")
    .split(/\s+/)
    .map((word) => word.trim())
    .filter((word) => word.length >= 2);
  const cjkTokens = [];
  const cjkSequences = value.match(/[\u3400-\u9fff]{2,}/g) || [];
  cjkSequences.forEach((sequence) => {
    for (let index = 0; index < sequence.length - 1; index += 1) {
      cjkTokens.push(sequence.slice(index, index + 2));
    }
    for (let index = 0; index < sequence.length - 2; index += 2) {
      cjkTokens.push(sequence.slice(index, index + 3));
    }
  });
  return Array.from(new Set([...raw, ...cjkTokens])).slice(0, 24);
}

function chunkTextForClient(text, maxLength) {
  const normalized = String(text || "").replace(/\s+/g, " ").trim();
  const chunks = [];
  for (let index = 0; index < normalized.length; index += maxLength) {
    chunks.push(normalized.slice(index, index + maxLength));
  }
  return chunks;
}

function firstClientLine(text) {
  return String(text || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .find(Boolean)
    ?.slice(0, 80) || "";
}

function calculateBudget(total, coreOverride = null) {
  const core = getValidCoreMinutes(total, coreOverride);
  const remaining = Math.max(0, total - core);
  return {
    total,
    opening: remaining * 0.4,
    core,
    qa: remaining * 0.4,
    buffer: remaining * 0.2,
    coreRatio: total ? core / total : DEFAULT_CORE_RATIO,
  };
}

function getConfiguredCoreMinutes(total) {
  return getValidCoreMinutes(total, dom.coreMinutes?.value);
}

function getValidCoreMinutes(total, value) {
  const fallback = roundOne(total * DEFAULT_CORE_RATIO);
  const number = Number(value);
  const core = Number.isFinite(number) && number > 0 ? number : fallback;
  return roundOne(clamp(core, 1, Math.max(1, total)));
}

function syncDefaultCoreMinutes(force = false) {
  if (!dom.coreMinutes || !dom.scriptMinutes) return;
  const total = clamp(Number(dom.scriptMinutes.value) || 60, 3, 180);
  const defaultCore = roundOne(total * DEFAULT_CORE_RATIO);
  if (force || !dom.coreMinutes.value || dom.coreMinutes.dataset.autoCore !== "false") {
    dom.coreMinutes.value = formatNumber(defaultCore);
    dom.coreMinutes.dataset.autoCore = "true";
  }
}

function calculateWpm() {
  const base = wpmProfiles[dom.wpmProfile.value]?.wpm || 132;
  if (dom.pace.value === "slow") return base - 12;
  if (dom.pace.value === "fast") return base + 10;
  return base;
}

function distributeMinutes(total, weights) {
  return weights.map((weight) => Math.max(1, Number((total * weight).toFixed(1))));
}

function buildMarkdown() {
  const inputs = state.lastLessonInputs || getLessonInputs();
  const slides = state.slides
    .map(
      (slide) => `## 第 ${slide.number} 頁：${slide.title}

- 教學事件：${slide.event}
- Bloom 層次：${slide.bloom}
- 分鐘：${formatNumber(slide.minutes)}

${slide.notes}`,
    )
    .join("\n\n");
  const audit = state.auditLog
    .slice(0, 12)
    .map((entry) => `- ${new Date(entry.at).toLocaleString("zh-Hant")}｜${entry.action}｜${entry.detail}`)
    .join("\n");

  return `# ${inputs.topic}

科目：${inputs.subject}
對象：${inputs.audience}
分鐘：${inputs.duration}
風格：${inputs.style}

## 學習目標

${inputs.objective}

## 教材大綱

${slides}

## 講稿

${state.script || "尚未生成講稿"}

## AI 生成透明度

本內容由 AI AI 協助生成，仍需教師完成最後審核。
${state.publishedRevision ? `\n發布版本：${state.publishedRevision.id}｜${new Date(state.publishedRevision.publishedAt).toLocaleString("zh-Hant")}\n` : ""}

${audit || "尚未有生成紀錄"}
`;
}

function buildAnnualMarkdown(plan) {
  if (!plan) return "# 全年課程包\n\n尚未生成年度規劃。\n";
  const metrics = plan.metrics;
  const lectures = plan.lectureUnits.map((unit) => `### ${unit.id}. ${unit.title}

- Week：${unit.week}
- Lecture / video：${unit.hours} 小時（${unit.videoMinutes} 分鐘）
- PPT：${unit.pptSlides} 頁，${unit.deckName}
- Template：${unit.templateId || unit.metadata?.template_id || "LT-CORE"}
- Metadata：${Object.entries(unit.metadata || {}).map(([key, value]) => `${key}=${Array.isArray(value) ? value.join("/") : value}`).join("；")}
- Recording cue：${unit.recordingCue}
- Learning outcomes：
${unit.outcomes.map((item) => `  - ${item}`).join("\n")}
- PPT focus：${unit.pptFocus.join("、")}
- PPT slide spec：
${(unit.slideSpec || []).map((item) => `  - S${item.slide_no} ${item.section}：${item.purpose}`).join("\n")}
- 專業 PPTX 生成清單：
${getLecturePptxChecklist(unit).map((slide) => `  - S${slide.slide_no} ${slide.title}｜大題目：${slide.big_topic}｜子題目：${slide.subtopic}｜${slide.teaching_minutes} min
    - Visible text：${slide.visible_text.join("；")}
    - Visual：${slide.visual_direction}
    - Speaker notes：${slide.speaker_notes.join(" ")}
    - QA：${slide.qa_gate.join("；")}`).join("\n")}
- QA checklist：
${(unit.qaChecklist || []).map((item) => `  - ${item}`).join("\n")}
- Dedup：${unit.duplicateCleanup}`).join("\n\n");

  const timetable = (plan.timetable || []).map((item) => `| W${item.week} | ${item.type} | ${item.id} | ${item.title} | ${item.hours}h | ${item.output} |`).join("\n");

  const labs = plan.labs.map((lab) => `### ${lab.id}. ${lab.title}

- 小時：${lab.hours}
- Week：${lab.week || "-"}
- 環境：${lab.environment}
- 產出：${lab.outcome}
- 交付：${lab.deliverables.join("、")}
- Rubric：${lab.rubric.join("、")}
${lab.generatedContent ? `\n#### 已生成 Lab 內容\n\n${lab.generatedContent}` : ""}`).join("\n\n");

  const assessments = plan.assessments.map((item) => `### ${item.type}. ${item.title}

- 小時：${item.hours}
- Week：${item.week || "-"}
- 權重：${item.weight}
- 交付：${item.deliverables.join("、")}
- 規則：${item.rules.join("；")}
${item.generatedContent ? `\n#### 已生成 Assessment 內容\n\n${item.generatedContent}` : ""}`).join("\n\n");

  return `# ${plan.inputs.moduleTitle}

生成時間：${new Date(plan.generatedAt).toLocaleString("zh-Hant")}
對象：${plan.inputs.audience}
學年週數：${plan.inputs.weeks}

## 專業與全面架構

${plan.professionalStandard?.summary || "未提供每週清單；使用預設骨架。"}

### 生成 Pipeline

${(plan.professionalStandard?.flow || []).map((item) => `- ${item}`).join("\n")}

### Metadata Contract

| Field | Required | Source | Status |
| --- | --- | --- | --- |
${(plan.metadataContract || []).map((item) => `| ${item.field} | ${item.required ? "yes" : "no"} | ${item.source} | ${item.status} |`).join("\n")}

### QA Gate

| ID | Name | Pass Rule | Blocking |
| --- | --- | --- | --- |
${(plan.qaGates || []).map((item) => `| ${item.id} | ${item.name} | ${item.passRule} | ${item.blocking} |`).join("\n")}

### Accessibility Checklist

${(plan.accessibilityChecklist || []).map((item) => `- ${item}`).join("\n")}

## 小時與交付總覽

- 總小時：${metrics.totalHours}
- Lecture / video：${metrics.lectureHours} 小時，${metrics.lectureUnits} 個 PPT / recording batch
- 預估 PPT：${metrics.pptSlides} 頁
- CA Lab：${metrics.labHours} 小時，${metrics.labCount} 個 lab
- Assessment：${metrics.assessmentHours} 小時，${metrics.assessmentCount} 個評核項

## PPT 去重策略

${plan.pptConsolidation.map((item) => `- ${item}`).join("\n")}

## Timetable

| Week | Type | ID | Item | Hours | Output |
| --- | --- | --- | --- | --- | --- |
${timetable || "| - | - | - | 尚未生成 | - | - |"}

## Lecture & PPT

${lectures}

## CA Lab Series

${labs}

## Assessments

${assessments}

${plan.assessmentBank?.markdown ? `## Assessment 題庫總表\n\n${plan.assessmentBank.markdown}` : ""}

## 合規與品質提醒

${plan.complianceNotes.map((item) => `- ${item}`).join("\n")}
`;
}

function buildPrompt() {
  const inputs = state.lastLessonInputs || getLessonInputs();
  return [
    "你是一位教育科技產品中的教學設計 AI。",
    `請為「${inputs.subject}」的「${inputs.topic}」設計 ${inputs.duration} 分鐘教材。`,
    `學生對象：${inputs.audience}`,
    `學習目標：${inputs.objective}`,
    `班情：${inputs.context}`,
    `教師追問回答：${inputs.interviewAnswers || "尚未提供"}`,
    `風格：${inputs.style}`,
    `必須整合 Bloom 層次：${inputs.bloom.map((key) => bloomMap[key].label).join("、")}`,
    "請輸出投影片大綱、講者稿、互動問題、評量方式與可修改版本。",
  ].join("\n");
}

function setFormInputs(inputs) {
  dom.topic.value = inputs.topic || "";
  dom.subject.value = inputs.subject || "";
  dom.audience.value = inputs.audience || "";
  dom.duration.value = inputs.duration || 45;
  dom.style.value = inputs.style || "清晰嚴謹";
  dom.objective.value = inputs.objective || "";
  dom.context.value = inputs.context || "";
  state.interviewAnswers = inputs.interviewAnswers || state.interviewAnswers || "";
  if (dom.questionAnswer) dom.questionAnswer.value = state.interviewAnswers;
  dom.bloomChecks.forEach((checkbox) => {
    checkbox.checked = (inputs.bloom || []).includes(checkbox.value);
  });
}

function persistState() {
  const payload = {
    annualPlan: state.annualPlan,
    slides: state.slides,
    script: state.script,
    questions: state.questions,
    materialPages: state.materialPages,
    slideJson: state.slideJson,
    materialMeta: state.materialMeta,
    versions: state.versions,
    messages: state.messages,
    auditLog: state.auditLog,
    publishedRevision: state.publishedRevision,
    studentQa: state.studentQa,
    qaMetrics: state.qaMetrics,
    assessmentBank: state.assessmentBank,
    interviewAnswers: state.interviewAnswers,
    role: state.role,
    lastLessonInputs: state.lastLessonInputs,
    materialText: dom.materialText.value,
    assistantContext: dom.assistantContext.value,
  };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
  renderStatus();
}

function restoreState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    const payload = JSON.parse(raw);
    state.annualPlan = payload.annualPlan || null;
    state.slides = payload.slides || [];
    state.script = payload.script || "";
    state.questions = payload.questions || [];
    state.materialPages = payload.materialPages || [];
    state.materialMeta = payload.materialMeta || null;
    state.slideJson = payload.slideJson || [];
    state.versions = payload.versions || [];
    state.messages = payload.messages || [];
    state.auditLog = payload.auditLog || [];
    state.publishedRevision = payload.publishedRevision || null;
    state.assessmentBank = payload.assessmentBank || payload.annualPlan?.assessmentBank || null;
    state.studentQa = payload.studentQa || state.studentQa;
    state.qaMetrics = payload.qaMetrics || state.qaMetrics;
    state.gamma.lastGeneration = payload.gamma?.lastGeneration || state.gamma.lastGeneration;
    state.interviewAnswers = payload.interviewAnswers || payload.lastLessonInputs?.interviewAnswers || "";
    state.role = payload.role || "teacher";
    state.lastLessonInputs = payload.lastLessonInputs || null;
    if (state.lastLessonInputs) setFormInputs(state.lastLessonInputs);
    if (dom.questionAnswer) dom.questionAnswer.value = state.interviewAnswers;
    dom.materialText.value = payload.materialText || buildMaterialFromSlides();
    dom.assistantContext.value = payload.assistantContext || buildAssistantContext();
    if (state.materialMeta) {
      setMaterialStatus(`已載入 ${state.materialMeta.filename}：${state.materialPages.length || 1} 個片段`, true);
    }
  } catch {
    localStorage.removeItem(STORAGE_KEY);
  }
}

function downloadFile(filename, content, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function downloadBase64File(filename, base64, type) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let index = 0; index < binary.length; index += 1) {
    bytes[index] = binary.charCodeAt(index);
  }
  const blob = new Blob([bytes], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function slugifyFilename(value) {
  return String(value || "lesson")
    .trim()
    .replace(/[\\/:*?"<>|]+/g, "-")
    .replace(/\s+/g, "-")
    .slice(0, 80) || "lesson";
}

function formatBytes(value) {
  const bytes = Number(value || 0);
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function countWords(text) {
  if (!text) return 0;
  const cjk = text.match(/[\u3400-\u9fff]/g)?.length || 0;
  const words = text
    .replace(/[\u3400-\u9fff]/g, " ")
    .match(/[A-Za-z0-9]+/g)?.length || 0;
  return cjk + words;
}

function countCoreLectureWords(script) {
  const text = String(script || "");
  if (!text.trim()) return 0;
  const teacherSection = extractBetween(text, "## 一、教師口語講稿", "## 二、學生自學補充講義")
    || extractBetween(text, "# 教師口語講稿", "# 講稿品質檢查")
    || text;
  const spokenBlocks = Array.from(teacherSection.matchAll(/####\s*教師口語講稿\s*\n([\s\S]*?)(?=\n####\s|\n###\s*第|\n##\s|$)/g))
    .map((match) => match[1]);
  const coreExpansion = Array.from(teacherSection.matchAll(/###\s*核心講授[^\n]*\n([\s\S]*?)(?=\n###\s|##\s|$)/g))
    .map((match) => match[1]);
  const source = spokenBlocks.length || coreExpansion.length
    ? [...spokenBlocks, ...coreExpansion].join("\n")
    : teacherSection;
  const coreOnly = source
    .split(/\r?\n/)
    .filter((line) => {
      const value = line.trim();
      if (!value) return false;
      if (/^#{1,6}\s*/.test(value)) return false;
      if (/^[-*]\s*內容來源標記[:：]/u.test(value)) return false;
      if (/^(來源|內容來源標記)[:：]/u.test(value)) return false;
      return true;
    })
    .join("\n");
  return countWords(coreOnly);
}

function extractBetween(text, startMarker, endMarker) {
  const start = text.indexOf(startMarker);
  if (start === -1) return "";
  const afterStart = start + startMarker.length;
  const end = text.indexOf(endMarker, afterStart);
  return text.slice(afterStart, end === -1 ? undefined : end);
}

function clean(value) {
  return String(value || "").trim();
}

function escapeHtml(value) {
  return String(value || "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function formatNumber(value) {
  return Number(value).toFixed(value % 1 === 0 ? 0 : 1);
}

function roundOne(value) {
  return Math.round(Number(value || 0) * 10) / 10;
}

function emptyText(text) {
  return `<div class="question-item"><span>${escapeHtml(text)}</span></div>`;
}

function cryptoId() {
  if (window.crypto?.randomUUID) return window.crypto.randomUUID();
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function structuredCloneSafe(value) {
  if (value == null) return value;
  if (window.structuredClone) return window.structuredClone(value);
  return JSON.parse(JSON.stringify(value));
}
