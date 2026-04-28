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

  setDisabled(["generateAnnualPlanBtn", "exportAnnualMdBtn", "exportAnnualJsonBtn", "copyAnnualContentBtn", "generateLessonBtn", "regenerateSlideBtn", "sendSlidesToScriptBtn", "generateScriptBtn", "shortenScriptBtn", "expandScriptBtn", "saveVersionBtn", "exportJsonBtn", "exportProjectJsonBtn", "importProjectJsonBtn", "exportLessonMdBtn", "exportMarkdownBtn", "exportPptxBtn", "exportCoursePackBtn", "copyPromptBtn"], !canEdit);
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
  state.questions = [];
  const minutes = distributeMinutes(inputs.duration, gagneEvents.map((item) => item.weight));

  state.slides = gagneEvents.map((item, index) => {
    const bloomKey = inputs.bloom.includes(item.bloom)
      ? item.bloom
      : inputs.bloom[index % inputs.bloom.length];
    const bloom = bloomMap[bloomKey];
    const title = buildSlideTitle(item.event, inputs.topic, index);
    return {
      id: cryptoId(),
      number: index + 1,
      title,
      event: item.event,
      bloom: bloom.label,
      bloomKey,
      minutes: minutes[index],
      activity: buildActivity(item.event, inputs, bloom),
      notes: buildPptSlidePrompt({
        title,
        event: item.event,
        inputs,
        bloom,
        minutes: minutes[index],
        activity: buildActivity(item.event, inputs, bloom),
        sourceNotes: buildSlideNotes(item.event, inputs, bloom, minutes[index]),
      }),
    };
  });
}

function normalizeAiSlides(slides, inputs) {
  const fallbackMinutes = distributeMinutes(inputs.duration, slides.map(() => 1 / slides.length));
  return slides.map((slide, index) => {
    const bloomKey = Object.keys(bloomMap).find((key) => bloomMap[key].label === slide.bloom) || inputs.bloom[index % inputs.bloom.length] || "understand";
    const bloom = bloomMap[bloomKey] || bloomMap.understand;
    return {
      id: cryptoId(),
      number: index + 1,
      title: clean(slide.title) || buildSlideTitle(slide.event, inputs.topic, index),
      event: clean(slide.event) || gagneEvents[index % gagneEvents.length].event,
      bloom: clean(slide.bloom) || bloom.label,
      bloomKey,
      minutes: Number(slide.minutes) || fallbackMinutes[index],
      activity: clean(slide.activity) || buildActivity(gagneEvents[index % gagneEvents.length].event, inputs, bloom),
      notes: buildPptSlidePrompt({
        title: clean(slide.title) || buildSlideTitle(slide.event, inputs.topic, index),
        event: clean(slide.event) || gagneEvents[index % gagneEvents.length].event,
        inputs,
        bloom,
        minutes: Number(slide.minutes) || fallbackMinutes[index],
        activity: clean(slide.activity) || buildActivity(gagneEvents[index % gagneEvents.length].event, inputs, bloom),
        sourceNotes: clean(slide.notes) || buildSlideNotes(gagneEvents[index % gagneEvents.length].event, inputs, bloom, fallbackMinutes[index]),
      }),
    };
  });
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
  const minutes = clamp(Number(dom.scriptMinutes.value) || 20, 3, 120);
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
      state.script = `${aiScript.script}${notes}`;
      logAudit("講稿生成", `${formatAiProviderName(state.ai.provider)} 依第 ${startPage} 頁與 ${minutes} 分鐘設定生成講稿`);
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

  logAudit("講稿生成", `本機規則依第 ${startPage} 頁與 ${minutes} 分鐘設定生成講稿`);
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

function renderScript() {
  dom.scriptOutput.value = state.script || "";
  renderStatus();
}

function renderTimeBudget() {
  const minutes = clamp(Number(dom.scriptMinutes.value) || 20, 3, 120);
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
  state.role = payload.role || state.role || "teacher";
  state.lastLessonInputs = payload.inputs || payload.lastLessonInputs || state.lastLessonInputs || getLessonInputs();

  setFormInputs(state.lastLessonInputs);
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
    `內容重點：${bloom.strategy}`,
    `檢核方式：請學生用一句話說明本頁最重要的概念，教師即時標記需補救的答案。`,
  ].join("\n");
}

function buildPptSlidePrompt({ title, event, inputs, bloom, minutes, activity, sourceNotes }) {
  const visual = inferPptVisualPrompt(title, inputs.topic, event);
  return `PPT 生成 Prompt

目標：為「${inputs.topic}」製作一頁 ${formatNumber(minutes)} 分鐘教學投影片，對象是 ${inputs.audience}。

頁面標題：${title}
教學事件：${event}
Bloom 層次：${bloom.label}

版面設計：
- 使用 16:9 professional training deck layout。
- 左側放核心概念或流程，右側放 demo / diagram / checklist。
- 每頁只放 1 個主訊息，不要堆滿文字。

投影片可見文字：
- 主標題：${title}
- 3 個重點 bullet，每點不超過 16 個中文字或 10 個英文詞。
- 1 個學生任務或 checkpoint：${activity}

視覺元素：
- ${visual}
- 使用清楚的 icon、流程箭頭、terminal / YAML snippet 或 architecture diagram。
- 避免裝飾性圖片；畫面要能支援教師解釋與學生重溫。

內容素材：
${sourceNotes}

互動 / 評核提示：
- 加入 1 個 quick check question。
- 若是 CKA/CKAD 相關頁，標示 exam angle 與常見錯誤。

輸出要求：
- 產生可直接放入 PowerPoint 的 slide content。
- 不要寫成逐字講稿；講稿會在「進度講稿」頁用生成後的 PPT 再產生。`;
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
  return `課題：${inputs.topic}。對象：${inputs.audience}。學習目標：${inputs.objective}。目前教材：${slideSummary}`;
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
    state.role = payload.role || "teacher";
    state.lastLessonInputs = payload.lastLessonInputs || null;
    if (state.lastLessonInputs) setFormInputs(state.lastLessonInputs);
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
