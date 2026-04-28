const STORAGE_KEY = "eduscript-ai-studio-state-v1";

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
  slides: [],
  script: "",
  versions: [],
  messages: [],
  lastLessonInputs: null,
  budget: null,
};

const dom = {};

document.addEventListener("DOMContentLoaded", () => {
  bindDom();
  bindEvents();
  restoreState();
  if (!state.slides.length) {
    generateLesson();
  }
  renderAll();
});

function bindDom() {
  dom.navItems = document.querySelectorAll(".nav-item");
  dom.views = document.querySelectorAll(".view");
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
  dom.materialFile = document.getElementById("materialFileInput");
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
  dom.versionList = document.getElementById("versionList");
  dom.compareBox = document.getElementById("compareBox");
}

function bindEvents() {
  dom.navItems.forEach((item) => {
    item.addEventListener("click", () => switchView(item.dataset.view));
  });

  dom.lessonForm.addEventListener("submit", (event) => {
    event.preventDefault();
    generateLesson();
    switchView("builder");
  });

  document.getElementById("askQuestionsBtn").addEventListener("click", renderQuestions);
  document.getElementById("regenerateSlideBtn").addEventListener("click", regenerateSelectedSlide);
  document.getElementById("loadDemoBtn").addEventListener("click", loadDemoProject);
  document.getElementById("saveVersionBtn").addEventListener("click", saveVersion);
  document.getElementById("exportJsonBtn").addEventListener("click", exportProjectJson);
  document.getElementById("exportProjectJsonBtn").addEventListener("click", exportProjectJson);
  document.getElementById("exportLessonMdBtn").addEventListener("click", exportLessonMarkdown);
  document.getElementById("exportMarkdownBtn").addEventListener("click", exportLessonMarkdown);
  document.getElementById("copyPromptBtn").addEventListener("click", copyPrompt);
  document.getElementById("clearProjectBtn").addEventListener("click", clearProject);

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

function generateLesson() {
  const inputs = getLessonInputs();
  state.lastLessonInputs = inputs;
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
      notes: buildSlideNotes(item.event, inputs, bloom, minutes[index]),
    };
  });

  dom.duration.value = inputs.duration;
  dom.materialText.value = state.slides
    .map((slide) => `第 ${slide.number} 頁：${slide.title}\n${slide.notes}`)
    .join("\n\n");
  dom.assistantContext.value = buildAssistantContext();
  renderQuestions();
  renderAll();
  persistState();
}

function renderQuestions() {
  const inputs = getLessonInputs();
  const questions = [
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
  renderTimeline();
  renderSlides();
  renderTimeBudget();
  renderScript();
  renderChat();
  renderVersions();
  renderStatus();
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
              </div>
            </div>
            <span class="slide-number">${slide.number}</span>
          </div>
          <textarea data-slide-notes="${index}" rows="7">${escapeHtml(slide.notes)}</textarea>
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

  slide.notes = `${slide.notes}\n\n依據修改意見：「${feedback}」\n${additions.length ? additions.join("\n") : "調整方向：降低抽象詞彙，增加明確步驟與教師口語提示。"}`;
  slide.activity = feedback;
  dom.slideFeedback.value = "";
  dom.assistantContext.value = buildAssistantContext();
  renderSlides();
  persistState();
}

function handleMaterialUpload(event) {
  const file = event.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = () => {
    dom.materialText.value = String(reader.result || "");
  };
  reader.readAsText(file, "UTF-8");
}

function generateScript() {
  const material = clean(dom.materialText.value) || buildMaterialFromSlides();
  const inputs = state.lastLessonInputs || getLessonInputs();
  const startPage = clamp(Number(dom.startPage.value) || 1, 1, 999);
  const minutes = clamp(Number(dom.scriptMinutes.value) || 20, 3, 120);
  const budget = calculateBudget(minutes);
  const wpm = calculateWpm();
  const targetWords = Math.round(budget.core * wpm);
  const fragments = materialFragments(material, startPage);

  const recap = state.slides
    .slice(0, Math.min(startPage - 1, state.slides.length))
    .slice(-3)
    .map((slide) => slide.title)
    .join("、");

  state.budget = { ...budget, wpm, targetWords };
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

  renderScript();
  renderTimeBudget();
  persistState();
}

function reviseScript(mode) {
  if (!state.script) generateScript();
  if (mode === "shorten") {
    const paragraphs = state.script.split("\n").filter((line) => line.trim());
    state.script = paragraphs.filter((_, index) => index % 3 !== 2).join("\n");
  } else {
    state.script = `${state.script}\n\n【補充層次】\n如果學生反應良好，可以加入一個延伸問題：請比較本課概念在理想條件與真實環境中的差異，並說明哪一個限制最容易影響結論。`;
  }
  renderScript();
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

function sendAssistantMessage() {
  const question = clean(dom.assistantQuestion.value);
  if (!question) return;
  const context = clean(dom.assistantContext.value) || buildAssistantContext();
  state.messages.push({ role: "user", text: question });
  state.messages.push({ role: "assistant", text: buildAssistantResponse(question, context) });
  dom.assistantQuestion.value = "";
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
  renderVersions();
  renderStatus();
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
  const payload = {
    exportedAt: new Date().toISOString(),
    inputs: state.lastLessonInputs || getLessonInputs(),
    slides: state.slides,
    script: state.script,
    versions: state.versions,
  };
  downloadFile("eduscript-ai-project.json", JSON.stringify(payload, null, 2), "application/json;charset=utf-8");
}

function exportLessonMarkdown() {
  downloadFile("eduscript-ai-lesson.md", buildMarkdown(), "text/markdown;charset=utf-8");
}

async function copyPrompt() {
  const prompt = buildPrompt();
  try {
    await navigator.clipboard.writeText(prompt);
    dom.compareBox.textContent = "Prompt 已複製。";
  } catch {
    dom.compareBox.textContent = prompt;
  }
}

function clearProject() {
  const ok = window.confirm("確定要清除本機工作台內容？");
  if (!ok) return;
  localStorage.removeItem(STORAGE_KEY);
  state.slides = [];
  state.script = "";
  state.versions = [];
  state.messages = [];
  state.lastLessonInputs = null;
  generateLesson();
}

function loadDemoProject() {
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
  generateLesson();
  dom.startPage.value = "4";
  dom.scriptMinutes.value = "18";
  generateScript();
  saveVersion();
}

function renderStatus() {
  dom.slideCount.textContent = String(state.slides.length);
  dom.durationStatus.textContent = String((state.lastLessonInputs || getLessonInputs()).duration);
  dom.scriptWordStatus.textContent = String(countWords(state.script));
  dom.versionStatus.textContent = String(state.versions.length);
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

function materialFragments(material, startPage) {
  const blocks = material
    .split(/\n{2,}|第\s*\d+\s*頁|Slide\s*\d+/i)
    .map((text) => clean(text))
    .filter(Boolean);
  const start = Math.max(0, Math.min(blocks.length - 1, startPage - 1));
  return blocks.slice(start, start + 5).map((block) => block.slice(0, 180));
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
    slides: state.slides,
    script: state.script,
    versions: state.versions,
    messages: state.messages,
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
    state.slides = payload.slides || [];
    state.script = payload.script || "";
    state.versions = payload.versions || [];
    state.messages = payload.messages || [];
    state.lastLessonInputs = payload.lastLessonInputs || null;
    if (state.lastLessonInputs) setFormInputs(state.lastLessonInputs);
    dom.materialText.value = payload.materialText || buildMaterialFromSlides();
    dom.assistantContext.value = payload.assistantContext || buildAssistantContext();
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
