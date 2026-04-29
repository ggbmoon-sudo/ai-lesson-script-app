const STORAGE_KEY = "eduscript-ai-studio-state-v1";
const DRIVE_SETTINGS_KEY = "eduscript-ai-drive-settings-v1";
const DRIVE_BACKUP_PREFIX = "eduscript-ai-backup-";
const DRIVE_SCOPE = "https://www.googleapis.com/auth/drive.file";
const GOOGLE_IDENTITY_SCRIPT = "https://accounts.google.com/gsi/client";

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
  interviewAnswers: "",
  role: "teacher",
  lastLessonInputs: null,
  budget: null,
  ai: {
    checked: false,
    enabled: false,
    provider: "",
    model: "",
    message: "本機規則",
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
  if (!state.slides.length) {
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
  dom.annualLectureTopics = document.getElementById("annualLectureTopicsInput");
  dom.annualLabSpec = document.getElementById("annualLabSpecInput");
  dom.annualAssessmentSpec = document.getElementById("annualAssessmentSpecInput");
  dom.annualMetrics = document.getElementById("annualMetrics");
  dom.annualNote = document.getElementById("annualNote");
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
  dom.wpmProfile = document.getElementById("wpmProfileInput");
  dom.pace = document.getElementById("paceInput");
  dom.timeBudget = document.getElementById("timeBudget");
  dom.wpmStatus = document.getElementById("wpmStatus");
  dom.coreMinutesStatus = document.getElementById("coreMinutesStatus");
  dom.targetWordsStatus = document.getElementById("targetWordsStatus");
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
  dom.scriptOutput.addEventListener("input", () => {
    state.script = dom.scriptOutput.value;
    persistState();
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

  [dom.scriptMinutes, dom.wpmProfile, dom.pace].forEach((input) => {
    input.addEventListener("change", renderTimeBudget);
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

  setDisabled(["generateAnnualPlanBtn", "exportAnnualMdBtn", "exportAnnualJsonBtn", "copyAnnualContentBtn", "generateLessonBtn", "regenerateSlideBtn", "sendSlidesToScriptBtn", "applyQuestionAnswersBtn", "generateScriptBtn", "shortenScriptBtn", "expandScriptBtn", "saveVersionBtn", "exportJsonBtn", "exportProjectJsonBtn", "importProjectJsonBtn", "exportLessonMdBtn", "exportMarkdownBtn", "exportPptxBtn", "exportCoursePackBtn", "exportGammaDeckBtn", "copyPromptBtn"], !canEdit);
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
      provider: "local",
      model: "",
      message: "本機規則",
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
      message: data.aiEnabled ? `${formatAiProviderName(data.provider)} 已連線` : "本機規則",
      busy: false,
      lastCheckedAt: new Date().toISOString(),
    };
  } catch {
    state.ai = {
      checked: true,
      enabled: false,
      provider: "local",
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
  if (!state.ai.enabled || window.location.protocol === "file:") {
    return null;
  }

  const response = await fetch(`/api/ai/${type}`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    throw new Error(error.error || `AI request failed: ${response.status}`);
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
    lectureTopics: splitPlanItems(dom.annualLectureTopics.value),
    labSpec: splitPlanLines(dom.annualLabSpec.value),
    assessmentSpec: splitPlanLines(dom.annualAssessmentSpec.value),
  };
}

function generateAnnualPlan() {
  const inputs = getAnnualInputs();
  state.annualPlan = buildAnnualPlan(inputs);
  logAudit("年度規劃", `${inputs.moduleTitle} 已生成全年 Lecture / Lab / Assessment 藍圖`);
  renderAnnualPlan();
  markDriveBackupNeeded("年度規劃");
  persistState();
}

function buildAnnualPlan(inputs) {
  const lectureCount = Math.max(1, Math.ceil(inputs.lectureHours));
  const lectureHoursEach = inputs.lectureHours / lectureCount;
  const labSpecs = inputs.labSpec.length ? inputs.labSpec : defaultLabSpecs();
  const labHoursEach = inputs.labHours / labSpecs.length;
  const lectureTopics = buildLectureTopics(inputs.lectureTopics, lectureCount);
  const lectureUnits = lectureTopics.map((topic, index) => buildLectureUnit(topic, index, lectureCount, lectureHoursEach, inputs));
  const labs = labSpecs.map((line, index) => buildLabUnit(line, index, labHoursEach));
  const assessments = buildAssessmentPlan(inputs);
  const timetable = buildAnnualTimetable({ lectureUnits, labs, assessments, inputs });
  const pptSlides = lectureUnits.reduce((sum, unit) => sum + unit.pptSlides, 0);

  return {
    id: cryptoId(),
    generatedAt: new Date().toISOString(),
    inputs,
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
    complianceNotes: [
      "CA 筆試題目應使用教師自建、公開授權或 AI 生成後人工審核的原創題；避免使用未授權題庫。",
      "EA Skill Test 保持 no hint；所有 API / Service endpoint 必須 public，並在 rubric 中列明驗收方法。",
      "所有 AI 生成教材需保留教師審核紀錄與版本備份。",
    ],
  };
}

function buildLectureTopics(customTopics, count) {
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
  const source = customTopics.length ? customTopics : defaults;
  return Array.from({ length: count }, (_, index) => source[index] || defaults[index % defaults.length]);
}

function buildLectureUnit(topic, index, count, hours, inputs) {
  const week = Math.max(1, Math.round(((index + 1) / count) * inputs.weeks));
  const pptSlides = Math.max(10, Math.round(hours * inputs.slidesPerHour));
  const focus = inferLectureFocus(topic);
  return {
    id: `L${index + 1}`,
    number: index + 1,
    week,
    title: topic,
    hours: roundOne(hours),
    videoMinutes: Math.round(hours * 60),
    pptSlides,
    deckName: `Deck ${index + 1}: ${topic}`,
    focus,
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
  };
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
  const lower = topic.toLowerCase();
  if (lower.includes("ckad")) return { core: "CKAD application workload", task: "workload deployment" };
  if (lower.includes("cka")) return { core: "CKA cluster administration", task: "cluster operation" };
  if (lower.includes("eks") || lower.includes("aws")) return { core: "EKS managed Kubernetes", task: "cloud cluster" };
  if (lower.includes("rancher")) return { core: "Rancher enterprise management", task: "multi-cluster operation" };
  if (lower.includes("linux")) return { core: "advanced Linux operation", task: "server preparation" };
  if (lower.includes("security") || lower.includes("rbac")) return { core: "Kubernetes security control", task: "RBAC policy" };
  if (lower.includes("service") || lower.includes("ingress")) return { core: "Kubernetes networking", task: "public service exposure" };
  return { core: "Kubernetes core concept", task: "hands-on Kubernetes" };
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

function buildAssessmentPlan(inputs) {
  const caHours = roundOne(inputs.assessmentHours * 0.45);
  const eaHours = roundOne(inputs.assessmentHours * 0.55);
  return [
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

  try {
    const aiLesson = await requestAi("lesson", { inputs });
    if (aiLesson?.slides?.length) {
      state.questions = Array.isArray(aiLesson.questions) ? aiLesson.questions : [];
      state.slides = normalizeAiSlides(aiLesson.slides, inputs);
      logAudit("教材生成", `${formatAiProviderName(state.ai.provider)} 生成 ${state.slides.length} 頁教材草稿`);
    } else {
      generateLessonLocal(inputs);
      logAudit("教材生成", `本機規則生成 ${state.slides.length} 頁教材草稿`);
    }
  } catch (error) {
    console.warn(error);
    generateLessonLocal(inputs);
    logAudit("教材生成", `AI 不可用，改用本機規則生成 ${state.slides.length} 頁教材草稿`);
  } finally {
    setAiBusy(false);
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

function generateLessonLocal(inputs) {
  const deckPlan = buildPptDeckPlan(inputs);
  state.questions = buildPptInterviewQuestions(inputs);
  state.slides = deckPlan.map((plan, index) => buildSlideFromDeckPlan(plan, inputs, index));
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

function renderAnnualPlan() {
  if (!dom.annualMetrics) return;
  const plan = state.annualPlan;
  if (!plan) {
    dom.annualMetrics.innerHTML = emptyText("按「生成全年規劃」建立 Lecture、Lab 與評核藍圖");
    dom.annualNote.textContent = "建議先確認 lecture 小時、Lab 小時與評核小時，再生成全年課程包。";
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
        <button class="action-button ghost" type="button" data-annual-lecture="${index}" ${canEditAnnual ? "" : "disabled"}>送到 PPT 流程</button>
      </div>
      <p>${escapeHtml(unit.recordingCue)}</p>
      <ul>${unit.outcomes.map((item) => `<li>${escapeHtml(item)}</li>`).join("")}</ul>
      <small>${escapeHtml(unit.duplicateCleanup)}</small>
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
    </article>
  `).join("");

  dom.annualContentTitle.textContent = plan.generatedContent?.title || "Lab / Assessment 內容生成區";
  dom.annualContentOutput.value = plan.generatedContent?.markdown || "";

  dom.annualLecturePlan.querySelectorAll("[data-annual-lecture]").forEach((button) => {
    button.addEventListener("click", () => sendAnnualLectureToBuilder(Number(button.dataset.annualLecture)));
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
  return "lecture";
}

function generateLabContent(index) {
  const plan = state.annualPlan;
  const lab = plan?.labs?.[index];
  if (!plan || !lab) return;
  const markdown = buildLabContentMarkdown(lab, plan, index);
  plan.generatedContent = {
    title: `${lab.id}｜${lab.title}`,
    markdown,
    type: "lab",
    index,
    updatedAt: new Date().toISOString(),
  };
  lab.generatedContent = markdown;
  logAudit("Lab 內容生成", `${lab.id} 已生成 instructions / steps / rubric`);
  renderAnnualPlan();
  markDriveBackupNeeded("Lab 內容生成");
  persistState();
}

function generateAssessmentContent(index) {
  const plan = state.annualPlan;
  const assessment = plan?.assessments?.[index];
  if (!plan || !assessment) return;
  const markdown = buildAssessmentContentMarkdown(assessment, plan, index);
  plan.generatedContent = {
    title: `${assessment.type}｜${assessment.title}`,
    markdown,
    type: "assessment",
    index,
    updatedAt: new Date().toISOString(),
  };
  assessment.generatedContent = markdown;
  logAudit("Assessment 內容生成", `${assessment.type} 已生成 assessment brief / rubric`);
  renderAnnualPlan();
  markDriveBackupNeeded("Assessment 內容生成");
  persistState();
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
  return `# ${lab.id}: ${lab.title}

## Timetable

- Week: ${lab.week || "-"}
- Hours: ${lab.hours}
- Environment: ${lab.environment}
- Related lecture: ${prevLecture?.id || "N/A"} ${prevLecture?.title || ""}

## Learning Outcomes

1. 完成可重現的 hands-on artifact，並能解釋每個主要指令或 YAML 欄位的用途。
2. 以截圖、command log、YAML / playbook 證明結果。
3. 用 80-120 字反思一個錯誤、排查過程與修正方法。

## Student Brief

你需要根據課堂內容完成 ${lab.title}。提交內容必須足夠讓老師或助教在另一台環境重現結果。

## Step-by-step Tasks

${checklist.map((item, itemIndex) => `${itemIndex + 1}. ${item}`).join("\n")}

## Deliverables

${lab.deliverables.map((item) => `- ${item}`).join("\n")}

## Acceptance Criteria

- 指令、YAML 或 playbook 可以重新執行。
- 截圖能清楚顯示 cluster / workload / endpoint 狀態。
- 若有 public endpoint，必須列出 URL、測試方法與預期 response。
- 反思能指出至少一個實際遇到的 error 或 limitation。

## Marking Rubric

${lab.rubric.map((item) => `- ${item}: 25%`).join("\n")}

## Teacher Notes

- 先檢查學生是否理解資源限制，特別是 VM RAM、CPU、storage。
- 不直接給完整答案；只提示如何閱讀 error、events、logs。
- 若學生使用 AI 生成指令，必須要求他解釋每個 flag / field。
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

function buildAssessmentContentMarkdown(assessment, plan) {
  const isEa = assessment.type === "EA";
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

## Rubric

- Correctness: 35%
- Reproducibility: 25%
- Troubleshooting evidence: 20%
- Explanation / reflection: 10%
- Security and cleanup awareness: 10%

## Rules

${assessment.rules.map((item) => `- ${item}`).join("\n")}

## Teacher Checklist

- 確認題目沒有依賴未授課或未授權材料。
- 確認評分標準在評核前公開。
- 確認答案與驗收命令已由教師試跑。
- 保存版本、rubric、sample answer 與 moderation notes。
`;
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
      `PPT focus：${unit.pptFocus.join("、")}`,
      `Duplicate cleanup：${unit.duplicateCleanup}`,
    ].filter(Boolean).join("\n"),
    bloom: ["understand", "apply", "analyze", "evaluate"],
  };

  setFormInputs(inputs);
  dom.scriptMinutes.value = String(Math.max(10, Math.round(unit.videoMinutes)));
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

function regenerateSelectedSlide() {
  if (!state.slides.length) return;
  const index = Number(dom.slideSelect.value || 0);
  const slide = state.slides[index];
  const feedback = clean(dom.slideFeedback.value) || "請讓內容更清楚、更適合學生理解";
  const additions = [];

  if (feedback.includes("生活") || feedback.includes("例")) {
    additions.push(`生活化例子：把「${state.lastLessonInputs.topic}」連到學生每天會遇到的情境，再請學生指出其中的關鍵變因。`);
  }
  if (feedback.includes("互動") || feedback.includes("問")) {
    additions.push("互動提問：請學生先個人想 30 秒，再與旁邊同學交換答案，最後用一句話回報。");
  }
  if (feedback.includes("考") || feedback.includes("試")) {
    additions.push("考試提示：把這頁概念轉成一題短答題，並標示得分關鍵詞。");
  }
  if (feedback.includes("短") || feedback.includes("精簡")) {
    slide.notes = slide.notes
      .split("\n")
      .filter((line) => line.trim())
      .slice(0, 4)
      .join("\n");
  }

  slide.notes = updatePptPromptWithFeedback(slide, feedback, additions);
  slide.activity = feedback;
  logAudit("局部修改", `第 ${slide.number} 頁依教師意見更新：${feedback}`);
  dom.slideFeedback.value = "";
  dom.assistantContext.value = buildAssistantContext();
  renderSlides();
  markDriveBackupNeeded("局部修改");
  persistState();
}

async function handleMaterialUpload(event) {
  const file = event.target.files[0];
  if (!file) return;
  setMaterialStatus(`正在解析 ${file.name}...`, true);

  try {
    const parsed = await parseMaterialFile(file);
    if (parsed?.text) {
      state.materialPages = parsed.pages || [];
      state.materialMeta = {
        filename: parsed.filename || file.name,
        type: parsed.type || file.name.split(".").pop(),
        warning: parsed.warning || "",
      };
      dom.materialText.value = parsed.text;
      logAudit("教材解析", `${state.materialMeta.filename} 解析為 ${state.materialPages.length || 1} 個片段`);
      setMaterialStatus(
        `已解析 ${state.materialMeta.filename}：${state.materialPages.length || 1} 個片段${state.materialMeta.warning ? `。${state.materialMeta.warning}` : ""}`,
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
  const budget = calculateBudget(minutes);
  const wpm = calculateWpm();
  const targetWords = Math.round(budget.core * wpm);
  const fragments = materialFragments(material, startPage, inputs);
  const focusedMaterial = fragments.length ? fragments.join("\n\n") : material;

  state.budget = { ...budget, wpm, targetWords };
  setAiBusy(true, "生成講稿中");

  try {
    const aiScript = await requestAi("script", {
      inputs,
      material: focusedMaterial,
      startPage,
      minutes,
      budget,
      wpm,
      targetWords,
    });
    if (aiScript?.script) {
      const notes = Array.isArray(aiScript.teachingNotes) && aiScript.teachingNotes.length
        ? `\n\n【教師課前提醒】\n${aiScript.teachingNotes.map((note) => `- ${note}`).join("\n")}`
        : "";
      state.script = ensureCompleteLectureScript(`${aiScript.script}${notes}`, {
        inputs,
        fragments,
        focusedMaterial,
        startPage,
        minutes,
        budget,
        wpm,
        targetWords,
      });
      const actualWords = countWords(state.script);
      const detail = actualWords < Math.round(targetWords * 0.9)
        ? `仍低於目標，已補成 ${actualWords}/${targetWords} 字`
        : `達到 ${actualWords}/${targetWords} 字`;
      logAudit("講稿生成", `${formatAiProviderName(state.ai.provider)} 依第 ${startPage} 頁與 ${minutes} 分鐘設定生成完整講稿（${detail}）`);
      renderScript();
      renderTimeBudget();
      markDriveBackupNeeded("講稿生成");
      persistState();
      return;
    }
  } catch (error) {
    console.warn(error);
  } finally {
    setAiBusy(false);
  }

  const recap = state.slides
    .slice(0, Math.min(startPage - 1, state.slides.length))
    .slice(-3)
    .map((slide) => slide.title)
    .join("、");

  state.script = [
    `【開場與前情提要｜約 ${formatNumber(budget.opening)} 分鐘】`,
    `各位同學，我們先用一分鐘把上一段內容接回來。上一堂課最重要的線索是：${recap || "上一段的核心概念與今日主題的連接"}。今天我們會從第 ${startPage} 頁開始，把「${inputs.topic}」推進到可以解釋、比較，並能回答真實情境問題的程度。`,
    "",
    `【核心講授｜約 ${formatNumber(budget.core)} 分鐘｜目標 ${targetWords} 字】`,
    buildCoreScript(fragments, inputs),
    "",
    `【停頓互動｜約 ${formatNumber(budget.qa)} 分鐘】`,
    `請先不要急著抄答案。用 30 秒寫下：如果你要向一位未學過 ${inputs.subject} 的朋友解釋「${inputs.topic}」，你會先講哪一個關鍵字？接著請兩位同學分享，我會把答案分成「概念正確」「需要修正」「可以再深化」三類。`,
    "",
    `【收束與緩衝｜約 ${formatNumber(budget.buffer)} 分鐘】`,
    `最後我們整理三句話：第一，今天的核心概念是什麼；第二，它和上一堂課如何連接；第三，下一次你看到類似題目時要先找哪個線索。請把這三句寫在筆記最下方，作為本節 exit ticket。`,
  ].join("\n");
  state.script = ensureCompleteLectureScript(state.script, {
    inputs,
    fragments,
    focusedMaterial,
    startPage,
    minutes,
    budget,
    wpm,
    targetWords,
  });

  logAudit("講稿生成", `本機規則依第 ${startPage} 頁與 ${minutes} 分鐘設定生成完整講稿（${countWords(state.script)}/${targetWords} 字）`);
  renderScript();
  renderTimeBudget();
  markDriveBackupNeeded("講稿生成");
  persistState();
}

function reviseScript(mode) {
  if (!state.script) {
    generateScript();
    return;
  }
  if (mode === "shorten") {
    const paragraphs = state.script.split("\n").filter((line) => line.trim());
    state.script = paragraphs.filter((_, index) => index % 3 !== 2).join("\n");
  } else {
    state.script = `${state.script}\n\n【補充層次】\n如果學生反應良好，可以加入一個延伸問題：請比較本課概念在理想條件與真實環境中的差異，並說明哪一個限制最容易影響結論。`;
  }
  renderScript();
  markDriveBackupNeeded("講稿修訂");
  persistState();
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
  return script.includes("【字數與使用方式】") ? script : `${script}${note}`;
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
  renderStatus();
}

function renderTimeBudget() {
  const minutes = clamp(Number(dom.scriptMinutes.value) || 60, 3, 180);
  const budget = state.budget && Number(state.budget.total) === minutes ? state.budget : calculateBudget(minutes);
  const wpm = state.budget && Number(state.budget.total) === minutes ? state.budget.wpm : calculateWpm();
  const targetWords = Math.round(budget.core * wpm);

  dom.timeBudget.innerHTML = [
    ["開場", budget.opening, 10],
    ["核心講授", budget.core, 65],
    ["問答反思", budget.qa, 20],
    ["切換緩衝", budget.buffer, 5],
  ]
    .map(
      ([label, value, percent]) => `
        <div class="budget-row">
          <div class="budget-label"><span>${label}</span><strong>${formatNumber(value)} 分</strong></div>
          <div class="budget-bar"><span style="width:${percent}%"></span></div>
        </div>
      `,
    )
    .join("");

  dom.wpmStatus.textContent = String(wpm);
  dom.coreMinutesStatus.textContent = `${formatNumber(budget.core)} 分`;
  dom.targetWordsStatus.textContent = String(targetWords);
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
      state.messages.push({ role: "assistant", text: buildAssistantResponse(question, context) });
      logAudit("即時助理", "本機規則生成課堂回應");
    }
  } catch (error) {
    console.warn(error);
    state.messages.push({ role: "assistant", text: buildAssistantResponse(question, context) });
    logAudit("即時助理", "AI 不可用，改用本機規則生成課堂回應");
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

function askPublishedLesson() {
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
  state.studentQa = {
    question,
    mode: supported ? "教材有據" : "尚未在教材找到",
    answer: supported ? buildGroundedStudentAnswer(question, sources) : "已發布教材中未找到足夠依據回答這個問題。請改問與目前教材更直接相關的問題，或請老師補充。",
    sources: supported ? sources.slice(0, 4) : [],
  };
  updateQaMetrics(supported);
  logAudit("學生問答", `${state.studentQa.mode}：${question.slice(0, 80)}`);
  renderPublishedQa();
  renderGovernanceMetrics();
  persistState();
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

function buildGroundedStudentAnswer(question, sources) {
  const top = sources.slice(0, 3);
  const keyPoints = top
    .map((source, index) => `${index + 1}. ${source.preview.replace(/\s+/g, " ").slice(0, 110)}（信心 ${Math.round((source.confidence || 0) * 100)}%）`)
    .join("\n");
  return [
    `根據已發布教材，這題可以先從 ${top[0].label} 理解。`,
    "",
    keyPoints,
    "",
    `簡短回答：${question} 的答案需要回到上面來源中的概念與例子；如果你想確認細節，請先打開來源片段，再對照老師發布的投影片。`,
  ].join("\n");
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
    slides: structuredCloneSafe(state.slides),
    script: state.script,
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
      (version, index) => `
        <article class="version-item">
          <strong>${escapeHtml(version.name)}</strong>
          <span>${version.slides.length} 頁教材 · ${countWords(version.script)} 字講稿</span>
          <div class="version-actions">
            <button type="button" data-restore-version="${index}">還原</button>
            <button type="button" data-compare-version="${index}">比較</button>
          </div>
        </article>
      `,
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
  state.slides = structuredCloneSafe(version.slides);
  state.script = version.script;
  setFormInputs(version.inputs);
  dom.assistantContext.value = buildAssistantContext();
  renderAll();
  persistState();
}

function compareVersion(index) {
  const version = state.versions[index];
  if (!version) return;
  const currentTitles = state.slides.map((slide) => slide.title);
  const previousTitles = version.slides.map((slide) => slide.title);
  const added = currentTitles.filter((title) => !previousTitles.includes(title));
  const removed = previousTitles.filter((title) => !currentTitles.includes(title));

  dom.compareBox.textContent = [
    `版本：${version.name}`,
    `目前頁數：${state.slides.length}｜版本頁數：${version.slides.length}`,
    `目前講稿字數：${countWords(state.script)}｜版本講稿字數：${countWords(version.script)}`,
    "",
    `新增標題：${added.length ? added.join("、") : "無"}`,
    `移除標題：${removed.length ? removed.join("、") : "無"}`,
  ].join("\n");
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
    schemaVersion: 3,
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
    materialText: dom.materialText.value,
    assistantContext: dom.assistantContext.value,
    script: state.script,
    versions: state.versions,
    messages: state.messages,
    auditLog: state.auditLog,
    publishedRevision: state.publishedRevision,
    studentQa: state.studentQa,
    qaMetrics: state.qaMetrics,
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
  state.versions = payload.versions || [];
  state.messages = payload.messages || state.messages || [];
  state.auditLog = payload.auditLog || [];
  state.publishedRevision = payload.publishedRevision || null;
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
  state.versions = [];
  state.messages = [];
  state.auditLog = [];
  state.publishedRevision = null;
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
  await generateScript();
  saveVersion();
}

function renderStatus() {
  dom.slideCount.textContent = String(state.slides.length);
  dom.durationStatus.textContent = String((state.lastLessonInputs || getLessonInputs()).duration);
  dom.scriptWordStatus.textContent = String(countWords(state.script));
  dom.versionStatus.textContent = String(state.versions.length);
  dom.publishStatus.textContent = state.publishedRevision ? "已發布" : "草稿";
  dom.groundedRateStatus.textContent = `${getGovernanceMetrics().groundedRate}%`;
}

function renderAiStatus() {
  if (!dom.aiStatus) return;
  const status = state.ai.busy ? "checking" : state.ai.enabled ? "online" : state.ai.checked ? "fallback" : "checking";
  const providerName = formatAiProviderName(state.ai.provider);
  const label = state.ai.busy ? state.ai.message : state.ai.enabled ? providerName : state.ai.message || "本機規則";
  const detail = state.ai.enabled
    ? state.ai.model || "server model"
    : state.ai.checked
      ? "未使用雲端 AI"
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
    state.ai.message = "本機規則";
  }
  renderAiStatus();
}

function formatAiProviderName(provider) {
  const normalized = String(provider || "").toLowerCase();
  if (normalized === "gemini") return "Gemini";
  if (normalized === "openai") return "OpenAI";
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

function buildAssistantResponse(question, context) {
  const inputs = state.lastLessonInputs || getLessonInputs();
  if (question.includes("查核") || question.toLowerCase().includes("fact")) {
    return [
      `可先把「${inputs.topic}」相關說法拆成三類：定義、數據、推論。`,
      "1. 定義：確認用詞是否和教材或課綱一致。",
      "2. 數據：標出年份、來源與適用範圍，避免把研究結論說成普遍定律。",
      "3. 推論：提醒學生目前是根據哪些證據作判斷。",
      "課堂措辭：這個說法我先標記為待查證，我們用資料來源和推論鏈來確認。",
    ].join("\n");
  }

  if (question.includes("互動") || question.includes("活動")) {
    return `安排 3 分鐘快速互動：先讓學生個人寫下「${inputs.topic}」的一個關鍵因果關係，再兩人互換檢查，最後請一組說出他們如何判斷答案可靠。`;
  }

  if (question.toLowerCase().includes("exit")) {
    return `Exit Ticket：請用不超過 40 字回答「${inputs.topic} 中最重要的一個因果關係是什麼？請指出原因與結果。」`;
  }

  return [
    `可這樣回應：${question}`,
    "",
    `先用一句話定義問題，再連回目前課堂脈絡：「${context.slice(0, 120)}」。`,
    `接著給學生一個判斷準則：如果答案能同時說明概念、證據與限制，就代表理解已經到達 ${inputs.style} 課堂期待的層次。`,
  ].join("\n");
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

function calculateBudget(total) {
  return {
    total,
    opening: total * 0.1,
    core: total * 0.65,
    qa: total * 0.2,
    buffer: total * 0.05,
  };
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

本內容由 AI 或本機規則協助生成，仍需教師完成最後審核。
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
- Recording cue：${unit.recordingCue}
- Learning outcomes：
${unit.outcomes.map((item) => `  - ${item}`).join("\n")}
- PPT focus：${unit.pptFocus.join("、")}
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
    materialMeta: state.materialMeta,
    versions: state.versions,
    messages: state.messages,
    auditLog: state.auditLog,
    publishedRevision: state.publishedRevision,
    studentQa: state.studentQa,
    qaMetrics: state.qaMetrics,
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
    state.versions = payload.versions || [];
    state.messages = payload.messages || [];
    state.auditLog = payload.auditLog || [];
    state.publishedRevision = payload.publishedRevision || null;
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
  if (window.structuredClone) return window.structuredClone(value);
  return JSON.parse(JSON.stringify(value));
}
