const http = require("node:http");
const fs = require("node:fs");
const path = require("node:path");
const zlib = require("node:zlib");
const { URL } = require("node:url");

loadEnvFile(path.join(__dirname, ".env"));

const PORT = Number(process.env.PORT || 4173);
const HOST = process.env.HOST || "0.0.0.0";
const AI_PROVIDER = normalizeProvider(process.env.AI_PROVIDER || "openai-compatible");
const OPENAI_COMPAT_BASE_URL = normalizeBaseUrl(process.env.OPENAI_COMPAT_BASE_URL || process.env.OPENAI_BASE_URL || process.env.AI_BASE_URL || "https://api.newcoin.top");
const OPENAI_COMPAT_CHAT_URL = process.env.OPENAI_COMPAT_CHAT_URL || "";
const OPENAI_COMPAT_API_KEY = process.env.OPENAI_COMPAT_API_KEY || process.env.OPENAI_API_KEY || process.env.AI_API_KEY || "";
const OPENAI_COMPAT_MODEL = process.env.OPENAI_COMPAT_MODEL || process.env.AI_MODEL || "qwen3.6-plus";
const OPENAI_COMPAT_TEMPERATURE = parseBoundedNumber(process.env.OPENAI_COMPAT_TEMPERATURE, 0.25, 0, 1);
const OPENAI_COMPAT_MAX_TOKENS = parsePositiveInteger(process.env.OPENAI_COMPAT_MAX_TOKENS, 16384);
const OPENAI_COMPAT_SCRIPT_MAX_TOKENS = parsePositiveInteger(process.env.OPENAI_COMPAT_SCRIPT_MAX_TOKENS, OPENAI_COMPAT_MAX_TOKENS);
const GEMINI_MODEL = process.env.GEMINI_MODEL || "gemini-3-pro-preview";
const GEMINI_API_KEY = process.env.GEMINI_API_KEY || process.env.GOOGLE_API_KEY || "";
const GEMINI_THINKING_LEVEL = normalizeThinkingLevel(process.env.GEMINI_THINKING_LEVEL || "high");
const GEMINI_THINKING_BUDGET = parseThinkingBudget(process.env.GEMINI_THINKING_BUDGET);
const GEMINI_TEMPERATURE = parseBoundedNumber(process.env.GEMINI_TEMPERATURE, 0.25, 0, 1);
const GEMINI_MAX_OUTPUT_TOKENS = parsePositiveInteger(process.env.GEMINI_MAX_OUTPUT_TOKENS, 32768);
const GEMINI_SCRIPT_MAX_OUTPUT_TOKENS = parsePositiveInteger(process.env.GEMINI_SCRIPT_MAX_OUTPUT_TOKENS, GEMINI_MAX_OUTPUT_TOKENS);
const GOOGLE_DRIVE_CLIENT_ID = process.env.GOOGLE_DRIVE_CLIENT_ID || "";
const PUBLIC_BASE_URL = process.env.PUBLIC_BASE_URL || process.env.RENDER_EXTERNAL_URL || "";
const GAMMA_API_KEY = process.env.GAMMA_API_KEY || "";
const GAMMA_API_BASE_URL = (process.env.GAMMA_API_BASE_URL || "https://public-api.gamma.app/v1.0").replace(/\/+$/, "");
const GAMMA_EXPORT_AS = normalizeGammaExport(process.env.GAMMA_EXPORT_AS || "pptx");
const GAMMA_TEXT_MODE = normalizeGammaTextMode(process.env.GAMMA_TEXT_MODE || "generate");
const GAMMA_THEME_ID = process.env.GAMMA_THEME_ID || "";
const GAMMA_FOLDER_IDS = parseGammaFolderIds(process.env.GAMMA_FOLDER_IDS || process.env.GAMMA_FOLDER_ID || "");
const ROOT = __dirname;

const AI_INSTRUCTIONS =
  "You are an expert teaching-material assistant. Return only valid JSON that matches the requested schema. Write in Traditional Chinese unless the source content asks otherwise.";

const mimeTypes = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".md": "text/markdown; charset=utf-8",
  ".txt": "text/plain; charset=utf-8",
};

const lessonSchema = {
  type: "object",
  additionalProperties: false,
  required: ["questions", "slides"],
  properties: {
    questions: {
      type: "array",
      minItems: 3,
      maxItems: 6,
      items: { type: "string" },
    },
    slides: {
      type: "array",
      minItems: 8,
      maxItems: 16,
      items: {
        type: "object",
        additionalProperties: false,
        required: ["title", "event", "bloom", "minutes", "activity", "notes"],
        properties: {
          title: { type: "string" },
          slideType: { type: "string" },
          event: { type: "string" },
          bloom: { type: "string" },
          minutes: { type: "number" },
          activity: { type: "string" },
          notes: { type: "string" },
          suggestedLayout: { type: "string" },
          suggestedVisual: { type: "string" },
          speakerNotes: { type: "string" },
          factCheckPoints: {
            type: "array",
            minItems: 0,
            maxItems: 5,
            items: { type: "string" },
          },
        },
      },
    },
  },
};

const scriptSchema = {
  type: "object",
  additionalProperties: false,
  required: ["teacherScriptPages", "selfStudyHandout", "generationLog", "script", "teachingNotes"],
  properties: {
    teacherScriptPages: {
      type: "array",
      minItems: 1,
      maxItems: 160,
      items: {
        type: "object",
        additionalProperties: false,
        required: ["pageNumber", "title", "sourceTags", "teachingPurpose", "spokenScript", "transition"],
        properties: {
          pageNumber: { type: "number" },
          title: { type: "string" },
          sourceTags: {
            type: "array",
            minItems: 1,
            maxItems: 3,
            items: { type: "string", enum: ["原教材內容", "推定補充", "需教師確認"] },
          },
          teachingPurpose: { type: "string" },
          spokenScript: { type: "string" },
          checkpoint: { type: "string" },
          transition: { type: "string" },
        },
      },
    },
    selfStudyHandout: { type: "string" },
    generationLog: { type: "string" },
    script: { type: "string" },
    teachingNotes: {
      type: "array",
      minItems: 2,
      maxItems: 6,
      items: { type: "string" },
    },
  },
};

const assistantSchema = {
  type: "object",
  additionalProperties: false,
  required: ["answer", "checks", "nextMove"],
  properties: {
    answer: { type: "string" },
    checks: {
      type: "array",
      minItems: 0,
      maxItems: 5,
      items: { type: "string" },
    },
    nextMove: { type: "string" },
  },
};

const stringArraySchema = {
  type: "array",
  minItems: 0,
  maxItems: 20,
  items: { type: "string" },
};

const pptxChecklistItemSchema = {
  type: "object",
  additionalProperties: false,
  required: ["slide_no", "title", "big_topic", "subtopic", "teaching_minutes", "template_id", "visible_text", "visual_direction", "speaker_notes", "lab_bridge", "qa_gate"],
  properties: {
    slide_no: { type: "number" },
    title: { type: "string" },
    big_topic: { type: "string" },
    subtopic: { type: "string" },
    teaching_minutes: { type: "number" },
    template_id: { type: "string" },
    visible_text: stringArraySchema,
    visual_direction: { type: "string" },
    speaker_notes: stringArraySchema,
    lab_bridge: { type: "string" },
    qa_gate: stringArraySchema,
    resource_profile: stringArraySchema,
    official_alignment: stringArraySchema,
  },
};

const slideSpecItemSchema = {
  type: "object",
  additionalProperties: false,
  required: ["slide_no", "section", "subtopic", "purpose", "renderer_hint", "required_notes"],
  properties: {
    slide_no: { type: "number" },
    section: { type: "string" },
    subtopic: { type: "string" },
    purpose: { type: "string" },
    renderer_hint: { type: "string" },
    required_notes: { type: "string" },
    repeatsEveryHour: { type: "boolean" },
  },
};

const lecturePptxSchema = {
  type: "object",
  additionalProperties: false,
  required: ["title", "subtopics", "teachingMinutes", "slideTarget", "templateId", "outcomes", "pptFocus", "recordingCue", "duplicateCleanup", "slideSpec", "pptxChecklist", "qaChecklist"],
  properties: {
    title: { type: "string" },
    subtopics: stringArraySchema,
    teachingMinutes: { type: "number" },
    slideTarget: { type: "number" },
    templateId: { type: "string" },
    outcomes: stringArraySchema,
    pptFocus: stringArraySchema,
    recordingCue: { type: "string" },
    duplicateCleanup: { type: "string" },
    slideSpec: {
      type: "array",
      minItems: 1,
      maxItems: 80,
      items: slideSpecItemSchema,
    },
    pptxChecklist: {
      type: "array",
      minItems: 1,
      maxItems: 80,
      items: pptxChecklistItemSchema,
    },
    qaChecklist: stringArraySchema,
  },
};

const lecturePptxSummarySchema = {
  type: "object",
  additionalProperties: false,
  required: ["title", "subtopics", "teachingMinutes", "slideTarget", "templateId", "outcomes", "pptFocus", "recordingCue", "duplicateCleanup", "qaChecklist"],
  properties: {
    title: { type: "string" },
    subtopics: stringArraySchema,
    teachingMinutes: { type: "number" },
    slideTarget: { type: "number" },
    templateId: { type: "string" },
    outcomes: stringArraySchema,
    pptFocus: stringArraySchema,
    recordingCue: { type: "string" },
    duplicateCleanup: { type: "string" },
    qaChecklist: stringArraySchema,
  },
};

const lecturePptxChecklistChunkSchema = {
  type: "object",
  additionalProperties: false,
  required: ["slideSpec", "pptxChecklist"],
  properties: {
    slideSpec: {
      type: "array",
      minItems: 1,
      maxItems: 12,
      items: slideSpecItemSchema,
    },
    pptxChecklist: {
      type: "array",
      minItems: 1,
      maxItems: 12,
      items: pptxChecklistItemSchema,
    },
  },
};

const annualPlanSchema = {
  type: "object",
  additionalProperties: false,
  required: ["summary", "lectureUnits", "labs", "assessments", "pptConsolidation", "qaGates", "accessibilityChecklist", "complianceNotes"],
  properties: {
    summary: { type: "string" },
    lectureUnits: {
      type: "array",
      minItems: 1,
      maxItems: 60,
      items: {
        type: "object",
        additionalProperties: false,
        required: ["id", "title", "subtopics", "outcomes", "pptFocus", "recordingCue", "duplicateCleanup", "slideSpec", "pptxChecklist", "qaChecklist"],
        properties: {
          id: { type: "string" },
          title: { type: "string" },
          subtopics: stringArraySchema,
          outcomes: stringArraySchema,
          pptFocus: stringArraySchema,
          recordingCue: { type: "string" },
          duplicateCleanup: { type: "string" },
          slideSpec: { type: "array", minItems: 1, maxItems: 80, items: slideSpecItemSchema },
          pptxChecklist: { type: "array", minItems: 1, maxItems: 80, items: pptxChecklistItemSchema },
          qaChecklist: stringArraySchema,
        },
      },
    },
    labs: {
      type: "array",
      minItems: 0,
      maxItems: 30,
      items: {
        type: "object",
        additionalProperties: false,
        required: ["id", "title", "environment", "outcome", "deliverables", "rubric"],
        properties: {
          id: { type: "string" },
          title: { type: "string" },
          environment: { type: "string" },
          outcome: { type: "string" },
          deliverables: stringArraySchema,
          rubric: stringArraySchema,
        },
      },
    },
    assessments: {
      type: "array",
      minItems: 0,
      maxItems: 20,
      items: {
        type: "object",
        additionalProperties: false,
        required: ["type", "title", "weight", "deliverables", "rules"],
        properties: {
          type: { type: "string" },
          title: { type: "string" },
          weight: { type: "string" },
          deliverables: stringArraySchema,
          rules: stringArraySchema,
        },
      },
    },
    pptConsolidation: stringArraySchema,
    qaGates: stringArraySchema,
    accessibilityChecklist: stringArraySchema,
    complianceNotes: stringArraySchema,
  },
};

const markdownContentSchema = {
  type: "object",
  additionalProperties: false,
  required: ["title", "markdown"],
  properties: {
    title: { type: "string" },
    markdown: { type: "string" },
  },
};

const assessmentBankSchema = {
  type: "object",
  additionalProperties: false,
  required: ["title", "markdown", "assessmentContents"],
  properties: {
    title: { type: "string" },
    markdown: { type: "string" },
    assessmentContents: {
      type: "array",
      minItems: 0,
      maxItems: 20,
      items: markdownContentSchema,
    },
  },
};

const scriptRevisionSchema = {
  type: "object",
  additionalProperties: false,
  required: ["script"],
  properties: {
    script: { type: "string" },
  },
};

const slideRevisionSchema = {
  type: "object",
  additionalProperties: false,
  required: ["title", "activity", "notes", "speakerNotes", "suggestedLayout", "suggestedVisual", "factCheckPoints"],
  properties: {
    title: { type: "string" },
    activity: { type: "string" },
    notes: { type: "string" },
    speakerNotes: { type: "string" },
    suggestedLayout: { type: "string" },
    suggestedVisual: { type: "string" },
    factCheckPoints: stringArraySchema,
  },
};

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, `http://${req.headers.host}`);

    if (req.method === "GET" && url.pathname === "/api/health") {
      const provider = resolveAiProvider();
      return sendJson(res, 200, {
        ok: true,
        aiEnabled: Boolean(provider),
        configuredProvider: AI_PROVIDER,
        provider: provider?.name || "local",
        model: provider?.model || "",
        openAiCompatible: provider?.name === "openai-compatible"
          ? {
              baseUrl: OPENAI_COMPAT_BASE_URL,
              temperature: OPENAI_COMPAT_TEMPERATURE,
              maxTokens: OPENAI_COMPAT_MAX_TOKENS,
              scriptMaxTokens: OPENAI_COMPAT_SCRIPT_MAX_TOKENS,
            }
          : null,
        thinking: provider?.name === "gemini" ? getGeminiThinkingStatus() : null,
        geminiModelTier: provider?.name === "gemini" ? classifyGeminiModel(GEMINI_MODEL) : null,
        geminiGeneration: provider?.name === "gemini"
          ? {
              temperature: GEMINI_TEMPERATURE,
              maxOutputTokens: GEMINI_MAX_OUTPUT_TOKENS,
              scriptMaxOutputTokens: GEMINI_SCRIPT_MAX_OUTPUT_TOKENS,
            }
          : null,
        availableProviders: {
          openaiCompatible: Boolean(OPENAI_COMPAT_API_KEY),
          gemini: Boolean(GEMINI_API_KEY),
        },
        environment: process.env.NODE_ENV || "development",
        publicBaseUrl: PUBLIC_BASE_URL,
        googleDriveConfigured: Boolean(GOOGLE_DRIVE_CLIENT_ID),
        gammaConfigured: Boolean(GAMMA_API_KEY),
        gammaExportAs: GAMMA_EXPORT_AS,
      });
    }

    if (req.method === "GET" && url.pathname === "/api/config") {
      return sendJson(res, 200, {
        googleDriveClientId: GOOGLE_DRIVE_CLIENT_ID,
        publicBaseUrl: PUBLIC_BASE_URL,
        gammaConfigured: Boolean(GAMMA_API_KEY),
        gammaExportAs: GAMMA_EXPORT_AS,
      });
    }

    if (req.method === "POST" && url.pathname.startsWith("/api/ai/")) {
      return await handleAiRequest(req, res, url.pathname);
    }

    if (req.method === "POST" && url.pathname === "/api/parse-material") {
      return await handleMaterialParse(req, res);
    }

    if (req.method === "POST" && url.pathname === "/api/export-pptx") {
      return await handlePptxExport(req, res);
    }

    if (req.method === "POST" && url.pathname === "/api/export-course-pack") {
      return await handleCoursePackExport(req, res);
    }

    if (req.method === "POST" && url.pathname === "/api/gamma/generate") {
      return await handleGammaGenerate(req, res);
    }

    if (req.method === "GET" && url.pathname.startsWith("/api/gamma/generation/")) {
      return await handleGammaStatus(req, res, url.pathname);
    }

    if (req.method !== "GET") {
      return sendJson(res, 405, { error: "Method not allowed" });
    }

    return serveStatic(url.pathname, res);
  } catch (error) {
    return sendJson(res, 500, { error: error.message || "Internal server error" });
  }
});

server.listen(PORT, HOST, () => {
  const provider = resolveAiProvider();
  const mode = provider ? `AI enabled (${provider.name}: ${provider.model})` : "AI provider required but not configured";
  const url = PUBLIC_BASE_URL || `http://localhost:${PORT}`;
  console.log(`EduScript AI Studio running at ${url} - ${mode}`);
});

async function handleAiRequest(req, res, pathname) {
  const provider = resolveAiProvider();
  if (!provider) {
    return sendJson(res, 503, {
      error: "AI generation is required. Set AI_PROVIDER=openai-compatible, OPENAI_COMPAT_BASE_URL, OPENAI_COMPAT_MODEL and OPENAI_COMPAT_API_KEY, then restart the server.",
    });
  }

  const body = await readJsonBody(req);

  if (pathname === "/api/ai/lesson") {
    const result = await createStructuredResponse({
      provider,
      schemaName: "lesson_plan",
      schema: lessonSchema,
      input: buildLessonPrompt(body),
    });
    return sendJson(res, 200, result);
  }

  if (pathname === "/api/ai/script") {
    const result = await createStructuredResponse({
      provider,
      schemaName: "lesson_script",
      schema: scriptSchema,
      input: buildScriptPrompt(body),
    });
    return sendJson(res, 200, result);
  }

  if (pathname === "/api/ai/assistant") {
    const result = await createStructuredResponse({
      provider,
      schemaName: "classroom_assistant",
      schema: assistantSchema,
      input: buildAssistantPrompt(body),
    });
    return sendJson(res, 200, result);
  }

  if (pathname === "/api/ai/student-qa") {
    const result = await createStructuredResponse({
      provider,
      schemaName: "student_grounded_qa",
      schema: assistantSchema,
      input: buildStudentQaPrompt(body),
    });
    return sendJson(res, 200, result);
  }

  if (pathname === "/api/ai/annual-plan") {
    const result = await createStructuredResponse({
      provider,
      schemaName: "annual_course_plan",
      schema: annualPlanSchema,
      input: buildAnnualPlanPrompt(body),
    });
    return sendJson(res, 200, result);
  }

  if (pathname === "/api/ai/lecture-pptx") {
    const result = await createStructuredResponse({
      provider,
      schemaName: "lecture_pptx_generation_plan",
      schema: lecturePptxSchema,
      input: buildLecturePptxPrompt(body),
    });
    return sendJson(res, 200, result);
  }

  if (pathname === "/api/ai/lecture-pptx-summary") {
    const result = await createStructuredResponse({
      provider,
      schemaName: "lecture_pptx_summary",
      schema: lecturePptxSummarySchema,
      input: buildLecturePptxSummaryPrompt(body),
    });
    return sendJson(res, 200, result);
  }

  if (pathname === "/api/ai/lecture-pptx-checklist") {
    const result = await createStructuredResponse({
      provider,
      schemaName: "lecture_pptx_checklist_chunk",
      schema: lecturePptxChecklistChunkSchema,
      input: buildLecturePptxChecklistPrompt(body),
    });
    return sendJson(res, 200, result);
  }

  if (pathname === "/api/ai/lab-content") {
    const result = await createStructuredResponse({
      provider,
      schemaName: "lab_content_markdown",
      schema: markdownContentSchema,
      input: buildLabContentPrompt(body),
    });
    return sendJson(res, 200, result);
  }

  if (pathname === "/api/ai/assessment-content") {
    const result = await createStructuredResponse({
      provider,
      schemaName: "assessment_content_markdown",
      schema: markdownContentSchema,
      input: buildAssessmentContentPrompt(body),
    });
    return sendJson(res, 200, result);
  }

  if (pathname === "/api/ai/assessment-bank") {
    const result = await createStructuredResponse({
      provider,
      schemaName: "assessment_bank_markdown",
      schema: assessmentBankSchema,
      input: buildAssessmentBankPrompt(body),
    });
    return sendJson(res, 200, result);
  }

  if (pathname === "/api/ai/script-revision") {
    const result = await createStructuredResponse({
      provider,
      schemaName: "script_revision",
      schema: scriptRevisionSchema,
      input: buildScriptRevisionPrompt(body),
    });
    return sendJson(res, 200, result);
  }

  if (pathname === "/api/ai/slide-revision") {
    const result = await createStructuredResponse({
      provider,
      schemaName: "slide_revision",
      schema: slideRevisionSchema,
      input: buildSlideRevisionPrompt(body),
    });
    return sendJson(res, 200, result);
  }

  return sendJson(res, 404, { error: "Unknown AI endpoint" });
}

async function handleMaterialParse(req, res) {
  const body = await readJsonBody(req, 30_000_000);
  const filename = String(body.filename || "material").trim();
  const mimeType = String(body.mimeType || "").trim();
  const extension = path.extname(filename).toLowerCase();
  const base64 = String(body.data || "").includes(",")
    ? String(body.data).split(",").pop()
    : String(body.data || "");

  if (!base64) {
    return sendJson(res, 400, { error: "Missing file data" });
  }

  const buffer = Buffer.from(base64, "base64");
  const parsed = parseMaterialBuffer(buffer, filename, extension, mimeType);
  return sendJson(res, 200, parsed);
}

async function handlePptxExport(req, res) {
  const body = await readJsonBody(req, 10_000_000);
  const inputs = body.inputs || {};
  const slides = Array.isArray(body.slides) ? body.slides : [];
  const script = String(body.script || "");

  if (!slides.length) {
    return sendJson(res, 400, { error: "No slides to export" });
  }

  const filename = safeFilename(`${inputs.topic || "eduscript-ai-lesson"}.pptx`);
  const deck = createPptxDeck({ inputs, slides, script });
  return sendJson(res, 200, {
    filename,
    mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    data: deck.toString("base64"),
  });
}

async function handleCoursePackExport(req, res) {
  const body = await readJsonBody(req, 20_000_000);
  const project = body.project || {};
  const inputs = project.inputs || body.inputs || {};
  const annualPlan = project.annualPlan || null;
  const slides = Array.isArray(project.slides) ? project.slides : [];
  const script = String(project.script || "");
  const annualMarkdown = String(body.annualMarkdown || "# 全年課程包\n\n尚未生成年度規劃。\n");
  const lessonMarkdown = String(body.lessonMarkdown || "# 教材大綱\n\n尚未生成教材。\n");
  const labMarkdown = buildLabsMarkdown(annualPlan);
  const assessmentMarkdown = buildAssessmentsMarkdown(annualPlan);
  const filename = safeArchiveFilename(`${inputs.topic || inputs.moduleTitle || "eduscript-ai-course-pack"}.zip`);

  const entries = [
    { name: "README.md", data: buildCoursePackReadme({ inputs, annualPlan, slides, script }) },
    { name: "01-year-plan.md", data: annualMarkdown },
    { name: "02-lesson-outline-and-script.md", data: lessonMarkdown },
    { name: "03-ca-lab-series.md", data: labMarkdown },
    { name: "04-assessments.md", data: assessmentMarkdown },
    { name: "project-backup.json", data: JSON.stringify(project, null, 2) },
  ];

  if (slides.length) {
    entries.push({
      name: "slides/lesson-deck.pptx",
      data: createPptxDeck({ inputs, slides, script }),
    });

    slides.forEach((slide, index) => {
      entries.push({
        name: `slides/slide-${String(index + 1).padStart(2, "0")}.md`,
        data: buildSlideMarkdown(slide, index + 1),
      });
    });
  }

  const archive = createZip(entries);
  return sendJson(res, 200, {
    filename,
    mimeType: "application/zip",
    data: archive.toString("base64"),
  });
}

async function handleGammaGenerate(req, res) {
  const body = await readJsonBody(req, 2_000_000);
  const inputText = String(body.inputText || "").trim();

  if (!inputText) {
    return sendJson(res, 400, { error: "Missing Gamma inputText" });
  }

  if (!GAMMA_API_KEY) {
    return sendJson(res, 503, {
      error: "GAMMA_API_KEY is not set. Export the Gamma-ready prompt and paste it into Gamma manually for now.",
      gammaConfigured: false,
    });
  }

  const gammaPayload = compactObject({
    inputText: inputText.slice(0, 400_000),
    additionalInstructions: String(body.additionalInstructions || "").trim().slice(0, 5_000) || undefined,
    textMode: normalizeGammaTextMode(body.textMode || GAMMA_TEXT_MODE),
    format: "presentation",
    numCards: normalizePositiveInteger(body.numCards),
    cardSplit: normalizeGammaCardSplit(body.cardSplit || "auto"),
    exportAs: normalizeGammaExport(body.exportAs || GAMMA_EXPORT_AS),
    themeId: String(body.themeId || GAMMA_THEME_ID || "").trim() || undefined,
    folderIds: Array.isArray(body.folderIds) && body.folderIds.length ? body.folderIds : GAMMA_FOLDER_IDS.length ? GAMMA_FOLDER_IDS : undefined,
    textOptions: compactObject({
      amount: body.textAmount || "auto",
      tone: body.tone || "professional",
      audience: body.audience || undefined,
      language: body.language || undefined,
    }),
    cardOptions: compactObject({
      dimensions: body.dimensions || "16x9",
    }),
  });

  const payload = await gammaFetch("/generations", {
    method: "POST",
    body: JSON.stringify(gammaPayload),
  });

  return sendJson(res, 200, {
    ...payload,
    gammaConfigured: true,
    request: {
      format: gammaPayload.format,
      numCards: gammaPayload.numCards || null,
      exportAs: gammaPayload.exportAs || null,
      textMode: gammaPayload.textMode,
    },
  });
}

async function handleGammaStatus(req, res, pathname) {
  const generationId = decodeURIComponent(pathname.split("/").pop() || "").trim();

  if (!generationId) {
    return sendJson(res, 400, { error: "Missing Gamma generation id" });
  }

  if (!GAMMA_API_KEY) {
    return sendJson(res, 503, { error: "GAMMA_API_KEY is not set.", gammaConfigured: false });
  }

  const payload = await gammaFetch(`/generations/${encodeURIComponent(generationId)}`, {
    method: "GET",
  });
  return sendJson(res, 200, payload);
}

async function gammaFetch(endpoint, options = {}) {
  const response = await fetch(`${GAMMA_API_BASE_URL}${endpoint}`, {
    ...options,
    headers: {
      "Content-Type": "application/json",
      "X-API-KEY": GAMMA_API_KEY,
      ...(options.headers || {}),
    },
  });
  const payload = await response.json().catch(() => ({}));

  if (!response.ok) {
    const detail = payload.error?.message || payload.message || payload.error || response.statusText;
    throw new Error(`Gamma API request failed: ${detail}`);
  }

  return payload;
}

function buildCoursePackReadme({ inputs, annualPlan, slides, script }) {
  const metrics = annualPlan?.metrics;
  const title = inputs.topic || inputs.moduleTitle || annualPlan?.inputs?.moduleTitle || "EduScript AI Course Pack";
  return `# ${title}

This ZIP was generated by EduScript AI Studio.

## Contents

- 01-year-plan.md: academic-year Lecture / Lab / Assessment plan
- 02-lesson-outline-and-script.md: lesson outline, slides, script, and AI transparency notes
- 03-ca-lab-series.md: CA Lab Series summary
- 04-assessments.md: CA / EA assessment planning
- slides/lesson-deck.pptx: generated PowerPoint deck
- slides/slide-XX.md: editable slide notes
- project-backup.json: full project backup for re-import

## Summary

- Subject: ${inputs.subject || annualPlan?.inputs?.moduleTitle || "N/A"}
- Audience: ${inputs.audience || annualPlan?.inputs?.audience || "N/A"}
- Slides: ${slides.length}
- Script words: ${countTextWords(script)}
- Lecture hours: ${metrics?.lectureHours || "N/A"}
- CA Lab hours: ${metrics?.labHours || "N/A"}
- Assessment hours: ${metrics?.assessmentHours || "N/A"}

AI-assisted content should be reviewed by the teacher before delivery.
`;
}

function buildLabsMarkdown(annualPlan) {
  const labs = Array.isArray(annualPlan?.labs) ? annualPlan.labs : [];
  if (!labs.length) return "# CA Lab Series\n\n尚未生成 Lab Series。\n";

  return `# CA Lab Series

${labs
  .map(
    (lab) => `## ${lab.id}. ${lab.title}

- 小時：${lab.hours}
- 環境：${lab.environment}
- 產出：${lab.outcome}
- 交付：${(lab.deliverables || []).join("、")}
- Rubric：${(lab.rubric || []).join("、")}
${lab.generatedContent ? `\n### Generated Content\n\n${lab.generatedContent}` : ""}
`,
  )
  .join("\n")}`;
}

function buildAssessmentsMarkdown(annualPlan) {
  const assessments = Array.isArray(annualPlan?.assessments) ? annualPlan.assessments : [];
  if (!assessments.length) return "# Assessments\n\n尚未生成評核規劃。\n";

  return `# Assessments

${assessments
  .map(
    (item) => `## ${item.type}. ${item.title}

- 小時：${item.hours}
- 權重：${item.weight}
- 交付：${(item.deliverables || []).join("、")}
- 規則：${(item.rules || []).join("；")}
${item.generatedContent ? `\n### Generated Content\n\n${item.generatedContent}` : ""}
`,
  )
  .join("\n")}`;
}

function buildSlideMarkdown(slide, fallbackNumber) {
  const number = slide.number || fallbackNumber;
  return `# Slide ${number}: ${slide.title || "Untitled"}

- Event: ${slide.event || "N/A"}
- Bloom: ${slide.bloom || "N/A"}
- Minutes: ${slide.minutes || "N/A"}

## Activity

${slide.activity || "N/A"}

## Notes

${slide.notes || "N/A"}
`;
}

function countTextWords(text) {
  const normalized = String(text || "").trim();
  if (!normalized) return 0;
  const latinWords = normalized.match(/[A-Za-z0-9]+(?:[-'][A-Za-z0-9]+)*/g) || [];
  const cjkChars = normalized.match(/[\u3400-\u9fff]/g) || [];
  return latinWords.length + cjkChars.length;
}

function parseMaterialBuffer(buffer, filename, extension, mimeType) {
  if ([".pptx", ".docx", ".xlsx"].includes(extension)) {
    return parseOpenXmlMaterial(buffer, filename, extension);
  }

  if (extension === ".pdf" || mimeType === "application/pdf") {
    return parsePdfMaterial(buffer, filename);
  }

  const text = buffer.toString("utf8");
  return {
    filename,
    type: extension.replace(".", "") || "text",
    pages: chunkText(text, 1200).map((chunk, index) => ({
      number: index + 1,
      title: `文字片段 ${index + 1}`,
      text: chunk,
    })),
    text,
    warning: "",
  };
}

function parseOpenXmlMaterial(buffer, filename, extension) {
  const entries = readZipEntries(buffer);

  if (extension === ".pptx") {
    const slides = Array.from(entries.keys())
      .filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name))
      .sort(compareNumberedPath);

    const pages = slides.map((name, index) => {
      const slideText = extractXmlText(entries.get(name).toString("utf8"));
      const notesName = `ppt/notesSlides/notesSlide${index + 1}.xml`;
      const notesText = entries.has(notesName)
        ? sanitizeParsedNotes(extractXmlText(entries.get(notesName).toString("utf8")))
        : "";
      const slideJson = buildCleanSlideJsonFromPpt({ slideText, notesText, index });
      const text = [
        slideJson.slide_title,
        slideJson.slide_subtitle,
        slideJson.slide_body,
        slideJson.speaker_notes && `講者備註：${slideJson.speaker_notes}`,
      ].filter(Boolean).join("\n");
      return {
        number: index + 1,
        title: slideJson.slide_title || `投影片 ${index + 1}`,
        text,
        slideJson,
      };
    });
    const slideJson = pages.map((page) => page.slideJson);

    return {
      filename,
      type: "pptx",
      pages,
      slideJson,
      text: pagesToText(pages),
      warning: pages.length ? "" : "未能在 PPTX 找到投影片文字。",
    };
  }

  if (extension === ".xlsx") {
    return parseXlsxMaterial(entries, filename);
  }

  const documentXml = entries.get("word/document.xml");
  if (!documentXml) {
    return {
      filename,
      type: "docx",
      pages: [],
      text: "",
      warning: "未能在 DOCX 找到 word/document.xml。",
    };
  }

  const paragraphs = extractDocxParagraphs(documentXml.toString("utf8"));
  const chunks = chunkText(paragraphs.join("\n"), 1200);
  const pages = chunks.map((chunk, index) => ({
    number: index + 1,
    title: firstLine(chunk) || `文件片段 ${index + 1}`,
    text: chunk,
  }));

  return {
    filename,
    type: "docx",
    pages,
    text: pagesToText(pages),
    warning: "",
  };
}

function sanitizeParsedNotes(text) {
  const lines = String(text || "")
    .split(/\r?\n|(?<=\D)\s+(?=\d{1,2}\s*$)/)
    .map((line) => line.trim())
    .filter(Boolean)
    .filter((line) => !/^(\d+[\s、,.;：:]*){1,12}$/u.test(line))
    .filter((line) => !/^(slide|notes?|speaker notes?|講者備註|備註)\s*\d*$/i.test(line));
  const cleaned = lines.join("\n").replace(/\s{2,}/g, " ").trim();
  if (cleaned.length < 8) return "";
  return cleaned;
}

function buildCleanSlideJsonFromPpt({ slideText, notesText, index }) {
  const lines = String(slideText || "")
    .split(/\r?\n+/)
    .map((line) => sanitizeParsedSlideLine(line))
    .filter(Boolean);
  const title = sanitizeParsedSlideLine(lines[0]) || `投影片 ${index + 1}`;
  const subtitleCandidate = lines[1] && lines[1].length <= 120 && !/^[-*•]/.test(lines[1])
    ? lines[1]
    : "";
  const bodyStart = subtitleCandidate ? 2 : 1;
  const body = lines.slice(bodyStart).join("\n").trim();
  return {
    slide_no: index + 1,
    slide_title: title,
    slide_subtitle: subtitleCandidate,
    slide_body: body,
    visual_description: "",
    speaker_notes: sanitizeParsedNotes(notesText),
    source_type: "原教材內容",
    extracted_from: "ppt",
  };
}

function sanitizeParsedSlideLine(line) {
  return String(line || "")
    .replace(/\s+/g, " ")
    .replace(/^講者備註[:：]?\s*/u, "")
    .trim();
}

function parsePdfMaterial(buffer, filename) {
  const raw = buffer.toString("latin1");
  const pageSegments = raw.split(/\/Type\s*\/Page\b/g).slice(1);
  const extracted = pageSegments.length
    ? pageSegments.map(extractPdfText).filter(Boolean)
    : [extractPdfText(raw)].filter(Boolean);
  const chunks = extracted.length ? extracted : chunkText(extractPdfText(raw), 1200);
  const pages = chunks.map((chunk, index) => ({
    number: index + 1,
    title: firstLine(chunk) || `PDF 片段 ${index + 1}`,
    text: chunk,
  }));

  return {
    filename,
    type: "pdf",
    pages,
    text: pagesToText(pages),
    warning: "PDF 解析為基礎文字抽取；掃描圖像型 PDF 需要 OCR 才能完整讀取。",
  };
}

function parseXlsxMaterial(entries, filename) {
  const sharedStrings = parseSharedStrings(entries.get("xl/sharedStrings.xml")?.toString("utf8") || "");
  const workbook = entries.get("xl/workbook.xml")?.toString("utf8") || "";
  const sheetNames = parseWorkbookSheetNames(workbook);
  const worksheetPaths = Array.from(entries.keys())
    .filter((name) => /^xl\/worksheets\/sheet\d+\.xml$/i.test(name))
    .sort(compareNumberedPath);

  const pages = worksheetPaths.map((name, index) => {
    const rows = parseWorksheetRows(entries.get(name).toString("utf8"), sharedStrings);
    const sheetName = sheetNames[index] || `工作表 ${index + 1}`;
    const textRows = rows
      .slice(0, 120)
      .map((row) => row.map((cell) => cell.value).filter(Boolean).join(" | "))
      .filter(Boolean);
    const text = textRows.join("\n");
    return {
      number: index + 1,
      title: sheetName,
      text: text || `${sheetName} 沒有可讀文字。`,
      rows,
    };
  });

  return {
    filename,
    type: "xlsx",
    pages,
    text: pagesToText(pages),
    warning: pages.length ? "" : "未能在 XLSX 找到工作表文字。",
  };
}

function parseSharedStrings(xml) {
  if (!xml) return [];
  return (xml.match(/<si\b[\s\S]*?<\/si>/g) || []).map((item) => extractXmlText(item));
}

function parseWorkbookSheetNames(xml) {
  const names = [];
  const regex = /<sheet\b[^>]*name="([^"]+)"/g;
  let match;
  while ((match = regex.exec(xml))) {
    names.push(decodeXml(match[1]));
  }
  return names;
}

function parseWorksheetRows(xml, sharedStrings) {
  const rows = [];
  const rowMatches = xml.match(/<row\b[\s\S]*?<\/row>/g) || [];
  for (const rowXml of rowMatches) {
    const cells = [];
    const cellRegex = /<c\b([^>]*)>([\s\S]*?)<\/c>/g;
    let match;
    while ((match = cellRegex.exec(rowXml))) {
      const attrs = match[1];
      const body = match[2];
      const ref = attrs.match(/\br="([^"]+)"/)?.[1] || "";
      const type = attrs.match(/\bt="([^"]+)"/)?.[1] || "";
      const rawValue = body.match(/<v>([\s\S]*?)<\/v>/)?.[1] || "";
      const inlineText = body.match(/<is\b[\s\S]*?<\/is>/)?.[0] || "";
      let value = "";

      if (type === "s") {
        value = sharedStrings[Number(rawValue)] || "";
      } else if (type === "inlineStr") {
        value = extractXmlText(inlineText);
      } else {
        value = decodeXml(rawValue);
      }

      cells.push({
        ref,
        column: ref.replace(/\d+/g, ""),
        value: normalizeWhitespace(value),
      });
    }
    if (cells.some((cell) => cell.value)) {
      rows.push(cells);
    }
  }
  return rows;
}

function createPptxDeck({ inputs, slides, script }) {
  const normalizedSlides = slides.map((slide, index) => ({
    number: index + 1,
    title: String(slide.title || `第 ${index + 1} 頁`),
    event: String(slide.event || ""),
    bloom: String(slide.bloom || ""),
    minutes: String(slide.minutes || ""),
    activity: String(slide.activity || ""),
    notes: String(slide.notes || ""),
    slideType: String(slide.slideType || slide.type || "content"),
    suggestedLayout: String(slide.suggestedLayout || ""),
    suggestedVisual: String(slide.suggestedVisual || ""),
    speakerNotes: String(slide.speakerNotes || ""),
    factCheckPoints: Array.isArray(slide.factCheckPoints) ? slide.factCheckPoints.map(String) : [],
  }));

  const entries = new Map();
  entries.set("[Content_Types].xml", contentTypesXml(normalizedSlides.length));
  entries.set("_rels/.rels", rootRelsXml());
  entries.set("docProps/core.xml", corePropsXml(inputs));
  entries.set("docProps/app.xml", appPropsXml(normalizedSlides.length));
  entries.set("ppt/presentation.xml", presentationXml(normalizedSlides.length));
  entries.set("ppt/_rels/presentation.xml.rels", presentationRelsXml(normalizedSlides.length));
  entries.set("ppt/theme/theme1.xml", themeXml());
  entries.set("ppt/slideMasters/slideMaster1.xml", slideMasterXml());
  entries.set("ppt/slideMasters/_rels/slideMaster1.xml.rels", slideMasterRelsXml());
  entries.set("ppt/slideLayouts/slideLayout1.xml", slideLayoutXml());
  entries.set("ppt/slideLayouts/_rels/slideLayout1.xml.rels", slideLayoutRelsXml());

  normalizedSlides.forEach((slide) => {
    entries.set(`ppt/slides/slide${slide.number}.xml`, slideXml(slide, inputs, script));
    entries.set(`ppt/slides/_rels/slide${slide.number}.xml.rels`, slideRelsXml(slide.number));
    entries.set(`ppt/notesSlides/notesSlide${slide.number}.xml`, notesSlideXml(slide, inputs, script));
    entries.set(`ppt/notesSlides/_rels/notesSlide${slide.number}.xml.rels`, notesSlideRelsXml(slide.number));
  });

  return createZip(Array.from(entries, ([name, content]) => ({
    name,
    data: Buffer.from(content, "utf8"),
  })));
}

function contentTypesXml(slideCount) {
  const slideOverrides = Array.from({ length: slideCount }, (_, index) =>
    `<Override PartName="/ppt/slides/slide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
  ).join("");
  const notesOverrides = Array.from({ length: slideCount }, (_, index) =>
    `<Override PartName="/ppt/notesSlides/notesSlide${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`,
  ).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  ${slideOverrides}
  ${notesOverrides}
</Types>`;
}

function rootRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
}

function corePropsXml(inputs) {
  const created = new Date().toISOString();
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>${xmlEscape(inputs.topic || "EduScript AI Lesson")}</dc:title>
  <dc:subject>${xmlEscape(inputs.subject || "AI-assisted lesson")}</dc:subject>
  <dc:creator>EduScript AI Studio</dc:creator>
  <cp:keywords>AI-assisted, lesson plan, teaching script</cp:keywords>
  <dcterms:created xsi:type="dcterms:W3CDTF">${created}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${created}</dcterms:modified>
</cp:coreProperties>`;
}

function appPropsXml(slideCount) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>EduScript AI Studio</Application>
  <PresentationFormat>On-screen Show (16:9)</PresentationFormat>
  <Slides>${slideCount}</Slides>
  <Company>EduScript AI Studio</Company>
</Properties>`;
}

function presentationXml(slideCount) {
  const slideIds = Array.from({ length: slideCount }, (_, index) =>
    `<p:sldId id="${256 + index}" r:id="rId${index + 2}"/>`,
  ).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
  <p:sldIdLst>${slideIds}</p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>
  <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`;
}

function presentationRelsXml(slideCount) {
  const slideRels = Array.from({ length: slideCount }, (_, index) =>
    `<Relationship Id="rId${index + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${index + 1}.xml"/>`,
  ).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  ${slideRels}
</Relationships>`;
}

function slideXml(slide, inputs, script) {
  const title = slide.title;
  const meta = [slide.slideType, slide.event, slide.bloom, slide.minutes && `${slide.minutes} 分鐘`].filter(Boolean).join(" / ");
  const body = buildPptxSlideBody(slide);
  const visual = slide.suggestedVisual || firstMatchingLine(slide.notes, "suggested_visual") || firstMatchingLine(slide.notes, "visual_preference");
  const layout = slide.suggestedLayout || firstMatchingLine(slide.notes, "suggested_layout") || firstMatchingLine(slide.notes, "layout_preference");
  const footer = `AI-Assisted Generation / Human review required / ${inputs.subject || ""}`;
  const scriptCue = firstLine(script).slice(0, 100);

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      ${groupShapeXml()}
      ${textShapeXml(2, "title", 560000, 420000, 11080000, 720000, [title], 3200, true, "173B63")}
      ${textShapeXml(3, "meta", 620000, 1180000, 10860000, 360000, [meta], 1500, false, "5E6F73")}
      ${textShapeXml(4, "body", 720000, 1700000, 6600000, 3600000, body, 1650, false, "1F2A2E")}
      ${textShapeXml(5, "visual", 7600000, 1700000, 3600000, 2400000, [`視覺：${visual || "diagram / checklist / terminal snippet"}`, `版面：${layout || "內建版型，清楚閱讀順序"}`], 1400, false, "285F63")}
      ${textShapeXml(7, "accessibility", 7600000, 4300000, 3600000, 900000, ["Alt text: describe insight, not decoration.", "Notes contain answer key and fallback."], 1050, false, "667276")}
      ${textShapeXml(8, "script cue", 720000, 5820000, 8400000, 380000, [scriptCue ? `講稿提示：${scriptCue}` : "講稿提示：答案鍵與轉場語見 speaker notes"], 1100, false, "667276")}
      ${textShapeXml(9, "footer", 9200000, 6220000, 2500000, 260000, [footer], 850, false, "8A8F91")}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
}

function buildPptxSlideBody(slide) {
  const promptBullets = extractPromptFieldLines(slide.notes, "slide_body").slice(0, 4);
  const body = [
    slide.activity && `任務：${slide.activity}`,
    ...promptBullets,
  ].filter(Boolean);
  if (body.length >= 3) return body.slice(0, 5);

  const fallback = slide.notes
    .split(/\r?\n/)
    .map((line) => normalizeWhitespace(line).replace(/^[-*]\s*/, ""))
    .filter((line) => line && !/^PPT Slide Compiler Prompt|你是一位|【|```|course_json|slide_no|slide_type|slide_goal|輸出格式|硬性限制/i.test(line))
    .filter((line) => line.length <= 90)
    .slice(0, 4);
  return [...body, ...fallback].slice(0, 5);
}

function extractPromptFieldLines(text, field) {
  const lines = String(text || "").split(/\r?\n/);
  const normalizedField = field.toLowerCase();
  const normalizePromptKey = (line) => line.trim().toLowerCase().replace(/^\d+\.\s*/, "");
  const start = lines.findIndex((line) => normalizePromptKey(line).startsWith(`${normalizedField}:`));
  if (start === -1) return [];
  const first = lines[start].trim().replace(/^\d+\.\s*/, "").split(":").slice(1).join(":").trim();
  const output = first ? [first] : [];
  for (let index = start + 1; index < lines.length; index += 1) {
    const line = lines[index].trim();
    if (!line) continue;
    if (/^\d+\.\s*[a-z_ ]+:/i.test(line) || /^[a-z_ ]+:/i.test(line)) break;
    output.push(line.replace(/^[-*]\s*/, ""));
    if (output.length >= 4) break;
  }
  return output;
}

function firstMatchingLine(text, pattern) {
  const match = String(text || "").split(/\r?\n/).find((line) => line.toLowerCase().includes(String(pattern).toLowerCase()));
  return match ? match.replace(/^[-*\s]*/, "").replace(/^[^:：]+[:：]\s*/, "").slice(0, 120) : "";
}

function notesSlideXml(slide, inputs, script) {
  const notes = buildPptxSpeakerNotes(slide, inputs, script);
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      ${groupShapeXml()}
      ${textShapeXml(2, "notes title", 450000, 420000, 6000000, 520000, [`Speaker Notes｜${slide.title}`], 1900, true, "173B63")}
      ${textShapeXml(3, "notes body", 520000, 1050000, 5900000, 7200000, notes, 1050, false, "1F2A2E")}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:notes>`;
}

function buildPptxSpeakerNotes(slide, inputs, script) {
  const explicit = slide.speakerNotes || firstMatchingLine(slide.notes, "speaker_notes");
  const factChecks = (slide.factCheckPoints?.length ? slide.factCheckPoints : extractPromptFieldLines(slide.notes, "fact_check_points")).slice(0, 5);
  const notes = [
    explicit || "先說明本頁和 learning objective 的關係，再引導學生完成可驗收任務。",
    slide.activity ? `課堂互動：${slide.activity}` : "",
    slide.suggestedLayout ? `版面提示：${slide.suggestedLayout}` : "",
    slide.suggestedVisual ? `視覺提示：${slide.suggestedVisual}` : "",
    ...factChecks.map((item) => `查核：${item}`),
    firstLine(script) ? `講稿連結：${firstLine(script).slice(0, 160)}` : "",
    "AI-assisted content. Teacher review required before delivery.",
  ].filter(Boolean);
  return notes.slice(0, 12);
}

function groupShapeXml() {
  return `<p:nvGrpSpPr>
  <p:cNvPr id="1" name=""/>
  <p:cNvGrpSpPr/>
  <p:nvPr/>
</p:nvGrpSpPr>
<p:grpSpPr>
  <a:xfrm>
    <a:off x="0" y="0"/>
    <a:ext cx="0" cy="0"/>
    <a:chOff x="0" y="0"/>
    <a:chExt cx="0" cy="0"/>
  </a:xfrm>
</p:grpSpPr>`;
}

function textShapeXml(id, name, x, y, cx, cy, lines, size, bold, color) {
  const paragraphs = lines.length ? lines : [""];
  return `<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="${id}" name="${xmlEscape(name)}"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:noFill/>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" anchor="t"/>
    <a:lstStyle/>
    ${paragraphs.map((line) => paragraphXml(line, size, bold, color)).join("")}
  </p:txBody>
</p:sp>`;
}

function paragraphXml(line, size, bold, color) {
  return `<a:p>
  <a:r>
    <a:rPr lang="zh-TW" sz="${size}"${bold ? ' b="1"' : ""}>
      <a:solidFill><a:srgbClr val="${color}"/></a:solidFill>
    </a:rPr>
    <a:t>${xmlEscape(line)}</a:t>
  </a:r>
  <a:endParaRPr lang="zh-TW" sz="${size}"/>
</a:p>`;
}

function slideRelsXml(slideNumber) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide${slideNumber}.xml"/>
</Relationships>`;
}

function notesSlideRelsXml(slideNumber) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide${slideNumber}.xml"/>
</Relationships>`;
}

function slideMasterXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld><p:spTree>${groupShapeXml()}</p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
  <p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles>
</p:sldMaster>`;
}

function slideMasterRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;
}

function slideLayoutXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree>${groupShapeXml()}</p:spTree></p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`;
}

function slideLayoutRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;
}

function themeXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="EduScript">
  <a:themeElements>
    <a:clrScheme name="EduScript">
      <a:dk1><a:srgbClr val="1F2A2E"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="173B63"/></a:dk2>
      <a:lt2><a:srgbClr val="F6F7F4"/></a:lt2>
      <a:accent1><a:srgbClr val="157A6E"/></a:accent1>
      <a:accent2><a:srgbClr val="C98217"/></a:accent2>
      <a:accent3><a:srgbClr val="7163A8"/></a:accent3>
      <a:accent4><a:srgbClr val="C9563A"/></a:accent4>
      <a:accent5><a:srgbClr val="667276"/></a:accent5>
      <a:accent6><a:srgbClr val="D7DFD9"/></a:accent6>
      <a:hlink><a:srgbClr val="157A6E"/></a:hlink>
      <a:folHlink><a:srgbClr val="7163A8"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="EduScript">
      <a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface="Microsoft JhengHei"/></a:majorFont>
      <a:minorFont><a:latin typeface="Aptos"/><a:ea typeface="Microsoft JhengHei"/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="EduScript"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>
  </a:themeElements>
</a:theme>`;
}


async function createStructuredResponse({ provider, schemaName, schema, input }) {
  if (provider.name === "gemini") {
    return createGeminiStructuredResponse({ schemaName, schema, input });
  }
  if (provider.name === "openai-compatible") {
    return createOpenAiCompatibleStructuredResponse({ schemaName, schema, input });
  }
  throw new Error(`Unsupported AI provider: ${provider.name}`);
}

async function createOpenAiCompatibleStructuredResponse({ schemaName, schema, input }) {
  const endpoint = getOpenAiCompatibleChatEndpoint();
  const maxTokens = schemaName === "lesson_script" ? OPENAI_COMPAT_SCRIPT_MAX_TOKENS : OPENAI_COMPAT_MAX_TOKENS;
  const requestBody = {
    model: OPENAI_COMPAT_MODEL,
    messages: [
      {
        role: "system",
        content: [
          AI_INSTRUCTIONS,
          "Return JSON only. Do not include markdown fences, commentary, or keys outside the requested structure.",
          `Schema name: ${schemaName}`,
          `JSON schema: ${JSON.stringify(schema)}`,
        ].join("\n\n"),
      },
      { role: "user", content: input },
    ],
    temperature: OPENAI_COMPAT_TEMPERATURE,
    max_tokens: maxTokens,
    stream: maxTokens > 4096,
    response_format: { type: "json_object" },
  };

  let payload = await postOpenAiCompatibleChat(endpoint, requestBody);
  if (payload.retryWithoutResponseFormat) {
    const fallbackBody = { ...requestBody };
    delete fallbackBody.response_format;
    payload = await postOpenAiCompatibleChat(endpoint, fallbackBody);
  }

  const text = extractOpenAiCompatibleText(payload);
  if (!text) {
    throw new Error("OpenAI-compatible response did not include message content.");
  }

  return parseStructuredJson(text, "OpenAI-compatible");
}

async function postOpenAiCompatibleChat(endpoint, body) {
  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${OPENAI_COMPAT_API_KEY}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(body),
  });

  if (body.stream) {
    const raw = await response.text().catch(() => "");
    if (!response.ok) {
      const detail = parseOpenAiCompatibleError(raw) || response.statusText;
      if (body.response_format && /response_format|json_schema|json_object|unsupported|not support/i.test(detail)) {
        return { retryWithoutResponseFormat: true };
      }
      throw new Error(`OpenAI-compatible request failed: ${detail}`);
    }
    return parseOpenAiCompatibleStream(raw);
  }

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    const detail = payload.error?.message || payload.error?.status || payload.message || response.statusText;
    if (body.response_format && /response_format|json_schema|json_object|unsupported|not support/i.test(detail)) {
      return { retryWithoutResponseFormat: true };
    }
    throw new Error(`OpenAI-compatible request failed: ${detail}`);
  }
  return payload;
}

function parseOpenAiCompatibleError(raw) {
  try {
    const payload = JSON.parse(raw || "{}");
    return payload.error?.message || payload.error?.status || payload.message || "";
  } catch {
    return String(raw || "").replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim().slice(0, 500);
  }
}

function parseOpenAiCompatibleStream(raw) {
  const chunks = [];
  for (const line of String(raw || "").split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed.startsWith("data:")) continue;
    const data = trimmed.slice("data:".length).trim();
    if (!data || data === "[DONE]") continue;
    try {
      const payload = JSON.parse(data);
      const choice = payload.choices?.[0] || {};
      const deltaContent = choice.delta?.content;
      const messageContent = choice.message?.content;
      const textContent = choice.text;
      if (typeof deltaContent === "string") chunks.push(deltaContent);
      if (typeof messageContent === "string") chunks.push(messageContent);
      if (typeof textContent === "string") chunks.push(textContent);
    } catch {
      // Ignore non-JSON stream keepalive lines.
    }
  }

  if (!chunks.length) {
    try {
      return JSON.parse(raw);
    } catch {
      return { choices: [{ message: { content: String(raw || "") } }] };
    }
  }

  return { choices: [{ message: { content: chunks.join("") } }] };
}

function getOpenAiCompatibleChatEndpoint() {
  if (OPENAI_COMPAT_CHAT_URL) return OPENAI_COMPAT_CHAT_URL;
  return OPENAI_COMPAT_BASE_URL.endsWith("/v1")
    ? `${OPENAI_COMPAT_BASE_URL}/chat/completions`
    : `${OPENAI_COMPAT_BASE_URL}/v1/chat/completions`;
}

function extractOpenAiCompatibleText(payload) {
  const content = payload?.choices?.[0]?.message?.content || payload?.choices?.[0]?.text || "";
  if (typeof content === "string") return content;
  if (Array.isArray(content)) {
    return content
      .map((part) => part?.text || part?.content || "")
      .filter(Boolean)
      .join("\n");
  }
  return "";
}

async function createGeminiStructuredResponse({ schemaName, schema, input }) {
  const modelPath = GEMINI_MODEL.startsWith("models/") ? GEMINI_MODEL.slice("models/".length) : GEMINI_MODEL;
  const endpoint = `https://generativelanguage.googleapis.com/v1beta/models/${encodeURIComponent(modelPath)}:generateContent`;
  const response = await fetch(endpoint, {
    method: "POST",
    headers: {
      "x-goog-api-key": GEMINI_API_KEY,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      systemInstruction: {
        parts: [{ text: AI_INSTRUCTIONS }],
      },
      contents: [
        {
          role: "user",
          parts: [{ text: input }],
        },
      ],
      generationConfig: buildGeminiGenerationConfig(schema, schemaName),
    }),
  });

  const payload = await response.json().catch(() => ({}));

  if (!response.ok) {
    const detail = payload.error?.message || payload.error?.status || response.statusText;
    throw new Error(`Gemini request failed: ${detail}`);
  }

  const text = extractGeminiText(payload);
  if (!text) {
    throw new Error("Gemini response did not include candidate text.");
  }

  return parseStructuredJson(text, "Gemini");
}

function buildAnnualPlanPrompt({ inputs, seedPlan }) {
  return `你是資深課程架構師、PPTX 製作總監與教學 QA 審查員。請使用 AI 重新生成一份專業、全面、可執行的全年課程包規劃。

要求：
1. 必須根據使用者的每週清單、官方來源、資源限制與 seedPlan 重新整理。
2. lectureUnits 每一項都要有大題目、子題目、outcomes、pptFocus、recordingCue、duplicateCleanup。
3. 每一個 lecture 必須輸出 slideSpec 與 pptxChecklist；pptxChecklist 要逐頁列出 visible_text、visual_direction、speaker_notes、lab_bridge、qa_gate。
4. labs 必須有 environment、outcome、deliverables、rubric。
5. assessments 必須有 type、title、weight、deliverables、rules。
6. qaGates 與 accessibilityChecklist 必須能阻擋不合格教材發布。
7. 請使用繁體中文；技術名詞可保留英文。

使用者輸入：
${JSON.stringify(inputs || {}, null, 2)}

目前本機 seed plan：
${JSON.stringify(seedPlan || {}, null, 2)}`;
}

function buildLecturePptxPrompt({ unit, inputs }) {
  return `你是專業 PPTX 教學設計師。請針對這個 Lecture 生成可直接交給 PPTX renderer / Gamma / PowerPoint 製作的詳細逐頁清單。

必須做到：
1. 使用使用者輸入的大題目、子題目與教學時間。
2. slideTarget 必須與教學時間和每小時頁數對齊。
3. pptxChecklist 必須逐頁包含 visible_text、visual_direction、speaker_notes、lab_bridge、qa_gate。
4. 每頁 visible_text 不超過 4 個 bullets；speaker_notes 要有答案鍵、checkpoint 或 fallback。
5. 加入 accessibility：唯一標題、alt text、reading order、contrast check。
6. outcomes、pptFocus、recordingCue 必須重新根據大題目、子題目、教學時間生成；如果現有資料與使用者輸入衝突，必須以使用者輸入為準。
7. 不可沿用與新大題目無關的舊內容；例如題目是 Linux / System Admin 時，不可保留 Kubernetes / CKA / CKAD 內容，除非使用者輸入明確要求。
8. 使用繁體中文；技術命令、YAML、產品名可保留英文。

課程輸入：
${JSON.stringify(inputs || {}, null, 2)}

Lecture 目前資料：
${JSON.stringify(unit || {}, null, 2)}`;
}

function buildLecturePptxSummaryPrompt({ unit, inputs }) {
  return `你是專業課程架構師與 PPTX 教學設計師。請快速重新生成這個 Lecture 的核心教學結構，只輸出 summary 欄位，不要輸出逐頁 checklist。

必須做到：
1. 以使用者最新大題目、子題目、教學時間、slideTarget 為準。
2. outcomes 必須具體、可觀察、可評核，不能沿用舊主題。
3. pptFocus 必須說明這堂課 PPT 應如何組織：概念、demo、checkpoint、troubleshooting、lab bridge。
4. recordingCue 必須依 teachingMinutes 拆成合理段落。
5. duplicateCleanup 必須指出若與前後 deck 重複，應保留什麼差異。
6. qaChecklist 必須能阻擋錯 topic、錯 duration、缺 speaker notes、缺 lab bridge 的 PPT。
7. 不可保留與新大題目無關的 Kubernetes / CKA / CKAD 內容，除非使用者輸入明確要求。
8. 使用繁體中文；技術命令、產品名可保留英文。

課程輸入：
${JSON.stringify(inputs || {}, null, 2)}

Lecture 最新資料：
${JSON.stringify(unit || {}, null, 2)}`;
}

function buildLecturePptxChecklistPrompt({ unit, inputs, startSlide, endSlide }) {
  return `你是 PPTX 製作總監。請只為指定 slide range 生成專業逐頁清單，避免一次輸出整份 deck。

Slide range：${startSlide} 到 ${endSlide}

必須做到：
1. slideSpec 與 pptxChecklist 只包含上述 slide_no 範圍。
2. 每一頁都必須對齊大題目、子題目、教學時間與 summary outcomes。
3. 每頁 visible_text 不超過 4 bullets。
4. speaker_notes 必須含 teaching purpose、講解重點、checkpoint / answer key / fallback。
5. visual_direction 必須能直接交給 PPTX/Gamma renderer。
6. lab_bridge 與 qa_gate 必須逐頁具體。
7. 使用繁體中文；命令、Linux service 名稱、產品名可保留英文。

課程輸入：
${JSON.stringify(inputs || {}, null, 2)}

Lecture summary：
${JSON.stringify(unit || {}, null, 2)}`;
}

function buildLabContentPrompt({ lab, plan, index }) {
  return `你是 CA Lab 設計師與技術助教。請用 AI 生成完整 Lab markdown，不可只列大綱。

輸出 markdown 必須包含：
- Student brief
- Learning goal
- Environment / resources
- Prerequisites
- Step-by-step instructions
- Expected outputs
- Evidence pack
- Troubleshooting guide
- Teacher answer key
- Rubric
- QA / safety / cleanup checklist

Lab index: ${index}
Lab:
${JSON.stringify(lab || {}, null, 2)}

Annual plan context:
${JSON.stringify(plan || {}, null, 2)}`;
}

function buildAssessmentContentPrompt({ assessment, plan, index }) {
  return `你是 assessment designer、moderator 與 rubric reviewer。請用 AI 生成完整 assessment markdown。

輸出 markdown 必須包含：
- Assessment brief
- Goal and assessed outcomes
- Format and time
- Deliverables
- Rules / integrity notes
- Question set or practical task set
- Model answer / answer key for teacher
- Evidence requirements
- Rubric table
- Moderation / QA checklist

Assessment index: ${index}
Assessment:
${JSON.stringify(assessment || {}, null, 2)}

Annual plan context:
${JSON.stringify(plan || {}, null, 2)}`;
}

function buildAssessmentBankPrompt({ plan }) {
  return `你是全年課程 assessment bank architect。請用 AI 為整個 course package 生成 assessment 題庫與 rubric markdown。

必須包含：
- 使用原則
- CA / EA 結構
- 每個 assessment 的題目、答案鍵、rubric、moderation notes
- practical skill test 的 no-hint 設計
- evidence pack 與 public endpoint 驗收規則
- 教師審核 checklist

Annual plan:
${JSON.stringify(plan || {}, null, 2)}`;
}

function buildScriptRevisionPrompt({ script, context, mode }) {
  return `你是資深教師講稿編修員。請用 AI 依指定模式修訂講稿，回傳完整 script。

模式：${mode || "expand"}

要求：
1. 保留原有章節結構。
2. 補強教師口語稿、轉場語、checkpoint、demo fallback。
3. 不輸出 prompt 解釋，只輸出修訂後完整講稿。
4. 使用繁體中文；技術名詞可保留英文。

Context:
${JSON.stringify(context || {}, null, 2)}

Current script:
${script || ""}`;
}

function buildSlideRevisionPrompt({ slide, feedback, inputs }) {
  return `你是專業 PPT prompt editor。請用 AI 根據教師意見重寫單頁 slide 的 PPT 生成 prompt。

要求：
1. 保留 slide 的教學目標，但改善 visible text、activity、speaker notes、layout、visual。
2. notes 必須是可直接給 PPTX/Gamma renderer 使用的 prompt。
3. 必須包含 slide_body、speaker_notes、suggested_layout、visual_preference、fact_check_points 等明確欄位。
4. 使用繁體中文；技術名詞可保留英文。

課程資料：
${JSON.stringify(inputs || {}, null, 2)}

原 slide：
${JSON.stringify(slide || {}, null, 2)}

教師修改意見：
${feedback || ""}`;
}

function buildLessonPrompt({ inputs }) {
  return `請生成一份專業 PPT deck 生成規格。請把「課程訪談資料」先正規化成 course_json，再用教學設計規則與投影片模板庫編譯成逐頁 PPT Prompt。不要直接把訪談文字改寫成鬆散投影片。

課題：${inputs.topic}
科目：${inputs.subject}
對象：${inputs.audience}
分鐘：${inputs.duration}
風格：${inputs.style}
學習目標：${inputs.objective}
先備知識與班情：${inputs.context}
教師已回答的 AI 追問：${inputs.interviewAnswers || "尚未提供"}
Bloom 層次：${(inputs.bloom || []).join(", ")}

建議流程：
1. 先建立 course_json：title、subject_domain、audience_profile、duration_min、style、objectives[]、prerequisites[]、bloom_levels[]、external_refs[]、source_completeness。
2. 用 backward design 對齊 learning objectives、assessment touchpoints 與 instructional strategies。
3. 依 slide template catalog 選擇頁型：title、prerequisite、objectives、agenda、content、example、demo、exercise、comparison、pitfalls、assessment、summary、references。
4. 若時長約 60 分鐘、技術實作或 exam-oriented，主 deck 建議 12-14 頁；必須包含 demo、exercise、assessment、comparison 或 pitfalls。
5. 每頁 notes 必須是可交給 Gamma / PPT AI 的單頁 prompt，包含 slide_title、slide_subtitle、slide_body、speaker_notes、suggested_visual、suggested_layout、presenter_cues、fact_check_points。

投影片可讀性與可及性限制：
- 每頁必須有唯一標題。
- slide_body 最多 4 個 bullets，避免大段文字；答案鍵、轉場語、fallback 放 speakerNotes。
- 建議使用 16:9、內建版型思維、足夠留白、18pt 以上字級、sans serif 字型、高對比。
- 資訊性圖像需提供 alt text；裝飾性圖像請標示 decorative。
- 若涉及 CKA/CKAD、kubectl、YAML、EKS、Ansible、Rancher，factCheckPoints 必須提醒以 official docs / exam pages 查核。

要求：
1. slides 每頁都要包含 slideType、suggestedLayout、suggestedVisual、speakerNotes、factCheckPoints。
2. event 仍需對應 Gagne 教學事件，bloom 需對應可觀察 learning outcome。
3. 若「教師已回答的 AI 追問」有內容，必須明顯反映在投影片重點、活動、評核與 examples 之中。
4. questions 是 AI 還需要追問教師的關鍵問題；不要重複已經被教師回答的問題。`;
}

function buildScriptPrompt({ inputs, material, slideJson = [], courseJson = null, teacherInterview = "", scriptPages = [], startPage, minutes, budget, wpm }) {
  const cleanSlideJson = Array.isArray(slideJson) && slideJson.length
    ? slideJson
    : buildFallbackSlideJsonFromScriptPages(scriptPages, material, startPage);
  const course = courseJson || buildCourseJsonForScriptPrompt({ inputs, minutes, budget, wpm });
  return `你是一位資深技術講師與教學設計師。請根據「完整 PPT slide_json」與「課程訪談資料」生成一份正式教師口語講稿。這份講稿要給老師直接上課使用，不是 prompt 匯出紀錄，也不是學生補充講義。

輸入資料：
【課程資訊】
${JSON.stringify(course, null, 2)}

【完整 PPT 解析 slide_json】
${JSON.stringify(cleanSlideJson, null, 2)}

【課程訪談資料】
${teacherInterview || inputs.interviewAnswers || "尚未提供"}

硬性要求：
1. 每一頁 PPT 都必須生成講稿，不可只挑部分頁面；teacherScriptPages 長度必須等於 slide_json 長度。
2. 完全對齊 PPT 頁序與頁數，不要新增不存在的投影片，不要省略任何一頁。
3. 每頁固定輸出：teachingPurpose、spokenScript、transition。checkpoint 只在重點頁輸出；非重點頁請留空字串或省略。
4. spokenScript 要像老師真的會講的話，不要只是複製投影片文字。每頁約 120-180 字，Demo / Troubleshooting 頁可較長；總字數只以核心講授分鐘估算。
5. sourceTags 只能使用「原教材內容」「推定補充」「需教師確認」三種文字；來自 PPT 用「原教材內容」；AI 合理延伸用「推定補充」；需要老師確認用「需教師確認」。
6. 禁止輸出內部 prompt、PPT Slide Compiler Prompt、debug log、版本紀錄、講者備註：1/2/3 這類解析殘留。
7. Demo 頁必須包含操作流程、預期輸出、驗收條件、失敗 fallback。
8. Troubleshooting 頁必須把錯誤現象對應到第一個要查的 command 或 evidence。
9. 互動問題 / checkpoint 只選重點頁，約每 5 頁 1 次；優先給 Demo、Troubleshooting、Assessment 或概念轉折頁，不要每頁硬塞問題。
10. generationLog 最後必須做自我檢查：是否每頁都有講稿、是否漏頁、checkpoint 是否只在重點頁、是否有重複模板句、是否有 prompt/debug log 殘留、哪些內容屬推定補充、哪些內容需教師確認。
11. script 欄位請輸出 Markdown，格式為：
# 教師口語講稿
## 第 1 頁：{slide_title}
來源：原教材內容 / 推定補充 / 需教師確認
### 本頁教學目的
...
### 教師口語講稿
...
### 互動問題 / Checkpoint（只在重點頁輸出）
...
### 轉場語
...
# 講稿品質檢查`;
}

function buildFallbackSlideJsonFromScriptPages(scriptPages = [], material = "", startPage = 1) {
  if (Array.isArray(scriptPages) && scriptPages.length) {
    return scriptPages.map((page, index) => ({
      slide_no: Number(page.number) || index + 1,
      slide_title: page.title || `投影片 ${index + 1}`,
      slide_subtitle: "",
      slide_body: String(page.text || "").slice(0, 1800),
      visual_description: "",
      speaker_notes: "",
      source_type: "原教材內容",
      extracted_from: "ppt",
    }));
  }
  return [{
    slide_no: Number(startPage) || 1,
    slide_title: "教材頁面",
    slide_subtitle: "",
    slide_body: String(material || "").slice(0, 1800),
    visual_description: "",
    speaker_notes: "",
    source_type: "原教材內容",
    extracted_from: "ppt",
  }];
}

function buildCourseJsonForScriptPrompt({ inputs, minutes, budget, wpm }) {
  return {
    title: inputs.topic,
    subject_domain: inputs.subject,
    audience_profile: inputs.audience,
    duration_min: minutes,
    core_teaching_min: budget?.core || "",
    style: inputs.style,
    objectives: inputs.objective,
    prerequisites: inputs.context,
    teacher_interview_answers: inputs.interviewAnswers || "",
    suggested_wpm: wpm,
    checkpoint_policy: "只在重點頁加入互動問題，約每 5 頁 1 次，不要每頁都輸出 checkpoint。",
  };
}

function buildAssistantPrompt({ context, question }) {
  return `目前課堂脈絡：
${context}

教師或學生問題：
${question}

請生成課堂即時助理回應：
1. answer 可直接口頭回答。
2. checks 列出需要事實查核或來源確認的點。
3. nextMove 給教師下一步課堂操作。`;
}

function buildStudentQaPrompt({ question, supported, sources = [], revisionMeta = {} }) {
  return `你是課程學生端問答助理。請只根據已發布教材來源回答；不可自行補充教材外內容。

問題：
${question || ""}

發布版本資料：
${JSON.stringify(revisionMeta || {}, null, 2)}

檢索是否達到支持門檻：${supported ? "yes" : "no"}

已發布教材來源：
${JSON.stringify(sources || [], null, 2)}

要求：
1. 如果 supported 為 false 或 sources 不足，answer 必須清楚拒答，說明目前教材未提供足夠依據，並請學生找老師補充或換一個更貼近教材的問題。
2. 如果 supported 為 true，answer 必須用學生容易理解的繁體中文回答，並點名引用來源 label；不得引入來源以外的新事實。
3. checks 列出答案需要教師確認或來源限制。
4. nextMove 給學生下一步應查看哪個來源或如何追問。`;
}

function extractGeminiText(payload) {
  const chunks = [];
  for (const candidate of payload.candidates || []) {
    for (const part of candidate.content?.parts || []) {
      if (part.text) {
        chunks.push(part.text);
      }
    }
  }
  return chunks.join("\n");
}

function parseStructuredJson(text, providerName) {
  const trimmed = String(text || "")
    .trim()
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/\s*```$/i, "")
    .trim();

  try {
    return JSON.parse(trimmed);
  } catch (error) {
    const objectStart = trimmed.indexOf("{");
    const objectEnd = trimmed.lastIndexOf("}");
    if (objectStart >= 0 && objectEnd > objectStart) {
      try {
        return JSON.parse(trimmed.slice(objectStart, objectEnd + 1));
      } catch {
        // Fall through to the original parse error for clearer diagnostics.
      }
    }
    throw new Error(`${providerName} response was not valid JSON: ${error.message}`);
  }
}

function toGeminiJsonSchema(schema) {
  if (Array.isArray(schema)) {
    return schema.map(toGeminiJsonSchema);
  }

  if (!schema || typeof schema !== "object") {
    return schema;
  }

  const allowedKeys = new Set([
    "$defs",
    "anyOf",
    "description",
    "enum",
    "items",
    "oneOf",
    "properties",
    "required",
    "type",
  ]);
  const result = {};

  for (const [key, value] of Object.entries(schema)) {
    if (!allowedKeys.has(key)) continue;

    if (key === "properties" || key === "$defs") {
      result[key] = Object.fromEntries(
        Object.entries(value || {}).map(([name, childSchema]) => [name, toGeminiJsonSchema(childSchema)]),
      );
    } else {
      result[key] = toGeminiJsonSchema(value);
    }
  }

  return result;
}

function buildGeminiGenerationConfig(schema, schemaName = "") {
  const config = {
    responseMimeType: "application/json",
    responseJsonSchema: toGeminiJsonSchema(schema),
    temperature: GEMINI_TEMPERATURE,
    maxOutputTokens: schemaName === "lesson_script" ? GEMINI_SCRIPT_MAX_OUTPUT_TOKENS : GEMINI_MAX_OUTPUT_TOKENS,
  };
  const thinkingConfig = buildGeminiThinkingConfig();
  if (thinkingConfig) {
    config.thinkingConfig = thinkingConfig;
  }
  return config;
}

function buildGeminiThinkingConfig() {
  const model = GEMINI_MODEL.toLowerCase();

  if (model.includes("gemini-3")) {
    return { thinkingLevel: GEMINI_THINKING_LEVEL };
  }

  if (GEMINI_THINKING_BUDGET !== null) {
    return { thinkingBudget: GEMINI_THINKING_BUDGET };
  }

  return undefined;
}

function getGeminiThinkingStatus() {
  const config = buildGeminiThinkingConfig();
  if (!config) {
    return { mode: "model-default" };
  }
  if (Object.hasOwn(config, "thinkingLevel")) {
    return { mode: "thinkingLevel", value: config.thinkingLevel };
  }
  return { mode: "thinkingBudget", value: config.thinkingBudget };
}

function classifyGeminiModel(model) {
  const value = String(model || "").toLowerCase();
  if (value.includes("flash") || value.includes("lite")) {
    return {
      tier: "fast",
      warning: "目前 Gemini model 偏向快速/低成本；完整技術講稿建議改用 Pro / Thinking 類模型。",
    };
  }
  if (value.includes("pro") || value.includes("thinking")) {
    return {
      tier: "high",
      warning: "",
    };
  }
  return {
    tier: "unknown",
    warning: "未能判斷 Gemini model 等級；請確認不是 Flash / Lite 類模型。",
  };
}

function resolveAiProvider() {
  if (AI_PROVIDER === "openai-compatible") {
    return OPENAI_COMPAT_API_KEY ? { name: "openai-compatible", model: OPENAI_COMPAT_MODEL } : null;
  }

  if (AI_PROVIDER === "gemini") {
    return GEMINI_API_KEY ? { name: "gemini", model: GEMINI_MODEL } : null;
  }

  if (AI_PROVIDER === "auto") {
    if (OPENAI_COMPAT_API_KEY) return { name: "openai-compatible", model: OPENAI_COMPAT_MODEL };
    if (GEMINI_API_KEY) return { name: "gemini", model: GEMINI_MODEL };
  }

  return null;
}

function normalizeProvider(value) {
  const provider = String(value || "auto").trim().toLowerCase();
  if (["openai-compatible", "openai_compatible", "openai", "qwen", "newcoin"].includes(provider)) {
    return "openai-compatible";
  }
  if (["auto", "gemini"].includes(provider)) return provider;
  return "openai-compatible";
}

function normalizeBaseUrl(value) {
  return String(value || "").trim().replace(/\/+$/, "");
}

function normalizeGammaExport(value) {
  const exportAs = String(value || "pptx").trim().toLowerCase();
  return ["pptx", "pdf", "png"].includes(exportAs) ? exportAs : "pptx";
}

function normalizeGammaTextMode(value) {
  const textMode = String(value || "generate").trim().toLowerCase();
  return ["generate", "condense", "preserve"].includes(textMode) ? textMode : "generate";
}

function normalizeGammaCardSplit(value) {
  const cardSplit = String(value || "auto").trim();
  return ["auto", "inputTextBreaks"].includes(cardSplit) ? cardSplit : "auto";
}

function normalizePositiveInteger(value) {
  const number = Number(value);
  if (!Number.isFinite(number) || number <= 0) return undefined;
  return Math.round(number);
}

function parseGammaFolderIds(value) {
  return String(value || "")
    .split(",")
    .map((item) => item.trim())
    .filter(Boolean)
    .slice(0, 10);
}

function compactObject(object) {
  return Object.fromEntries(
    Object.entries(object || {}).filter(([, value]) => {
      if (value === undefined || value === null || value === "") return false;
      if (Array.isArray(value)) return value.length > 0;
      if (typeof value === "object") return Object.keys(value).length > 0;
      return true;
    }),
  );
}

function normalizeThinkingLevel(value) {
  const level = String(value || "high").trim().toLowerCase();
  return ["minimal", "low", "medium", "high"].includes(level) ? level : "high";
}

function parseThinkingBudget(value) {
  if (value === undefined || value === null || String(value).trim() === "") {
    return null;
  }
  const budget = Number(value);
  return Number.isFinite(budget) ? Math.trunc(budget) : null;
}

function parsePositiveInteger(value, fallback) {
  const number = Number(value);
  if (!Number.isFinite(number) || number <= 0) return fallback;
  return Math.round(number);
}

function parseBoundedNumber(value, fallback, min, max) {
  const number = Number(value);
  if (!Number.isFinite(number)) return fallback;
  return Math.min(max, Math.max(min, number));
}

function createZip(entries) {
  const centralRecords = [];
  const fileRecords = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBuffer = Buffer.from(entry.name, "utf8");
    const data = Buffer.isBuffer(entry.data) ? entry.data : Buffer.from(String(entry.data), "utf8");
    const compressed = zlib.deflateRawSync(data);
    const crc = crc32(data);
    const localHeader = Buffer.alloc(30);

    localHeader.writeUInt32LE(0x04034b50, 0);
    localHeader.writeUInt16LE(20, 4);
    localHeader.writeUInt16LE(0x0800, 6);
    localHeader.writeUInt16LE(8, 8);
    localHeader.writeUInt16LE(0, 10);
    localHeader.writeUInt16LE(0, 12);
    localHeader.writeUInt32LE(crc, 14);
    localHeader.writeUInt32LE(compressed.length, 18);
    localHeader.writeUInt32LE(data.length, 22);
    localHeader.writeUInt16LE(nameBuffer.length, 26);
    localHeader.writeUInt16LE(0, 28);

    fileRecords.push(localHeader, nameBuffer, compressed);

    const centralHeader = Buffer.alloc(46);
    centralHeader.writeUInt32LE(0x02014b50, 0);
    centralHeader.writeUInt16LE(20, 4);
    centralHeader.writeUInt16LE(20, 6);
    centralHeader.writeUInt16LE(0x0800, 8);
    centralHeader.writeUInt16LE(8, 10);
    centralHeader.writeUInt16LE(0, 12);
    centralHeader.writeUInt16LE(0, 14);
    centralHeader.writeUInt32LE(crc, 16);
    centralHeader.writeUInt32LE(compressed.length, 20);
    centralHeader.writeUInt32LE(data.length, 24);
    centralHeader.writeUInt16LE(nameBuffer.length, 28);
    centralHeader.writeUInt16LE(0, 30);
    centralHeader.writeUInt16LE(0, 32);
    centralHeader.writeUInt16LE(0, 34);
    centralHeader.writeUInt16LE(0, 36);
    centralHeader.writeUInt32LE(0, 38);
    centralHeader.writeUInt32LE(offset, 42);
    centralRecords.push(centralHeader, nameBuffer);

    offset += localHeader.length + nameBuffer.length + compressed.length;
  }

  const centralDirectory = Buffer.concat(centralRecords);
  const endRecord = Buffer.alloc(22);
  endRecord.writeUInt32LE(0x06054b50, 0);
  endRecord.writeUInt16LE(0, 4);
  endRecord.writeUInt16LE(0, 6);
  endRecord.writeUInt16LE(entries.length, 8);
  endRecord.writeUInt16LE(entries.length, 10);
  endRecord.writeUInt32LE(centralDirectory.length, 12);
  endRecord.writeUInt32LE(offset, 16);
  endRecord.writeUInt16LE(0, 20);

  return Buffer.concat([...fileRecords, centralDirectory, endRecord]);
}

function crc32(buffer) {
  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc = (crc >>> 8) ^ crcTable[(crc ^ byte) & 0xff];
  }
  return (crc ^ 0xffffffff) >>> 0;
}

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let index = 0; index < 256; index += 1) {
    let value = index;
    for (let bit = 0; bit < 8; bit += 1) {
      value = value & 1 ? 0xedb88320 ^ (value >>> 1) : value >>> 1;
    }
    table[index] = value >>> 0;
  }
  return table;
})();

function readZipEntries(buffer) {
  const eocdOffset = findEndOfCentralDirectory(buffer);
  if (eocdOffset === -1) {
    throw new Error("Invalid ZIP/OpenXML file.");
  }

  const entryCount = buffer.readUInt16LE(eocdOffset + 10);
  const centralDirectoryOffset = buffer.readUInt32LE(eocdOffset + 16);
  const entries = new Map();
  let offset = centralDirectoryOffset;

  for (let index = 0; index < entryCount; index += 1) {
    if (buffer.readUInt32LE(offset) !== 0x02014b50) break;

    const compression = buffer.readUInt16LE(offset + 10);
    const compressedSize = buffer.readUInt32LE(offset + 20);
    const fileNameLength = buffer.readUInt16LE(offset + 28);
    const extraLength = buffer.readUInt16LE(offset + 30);
    const commentLength = buffer.readUInt16LE(offset + 32);
    const localHeaderOffset = buffer.readUInt32LE(offset + 42);
    const name = buffer.slice(offset + 46, offset + 46 + fileNameLength).toString("utf8");

    const localNameLength = buffer.readUInt16LE(localHeaderOffset + 26);
    const localExtraLength = buffer.readUInt16LE(localHeaderOffset + 28);
    const dataStart = localHeaderOffset + 30 + localNameLength + localExtraLength;
    const compressed = buffer.slice(dataStart, dataStart + compressedSize);

    if (!name.endsWith("/")) {
      entries.set(name, inflateZipEntry(compressed, compression));
    }

    offset += 46 + fileNameLength + extraLength + commentLength;
  }

  return entries;
}

function findEndOfCentralDirectory(buffer) {
  const signature = 0x06054b50;
  const minOffset = Math.max(0, buffer.length - 66_000);
  for (let offset = buffer.length - 22; offset >= minOffset; offset -= 1) {
    if (buffer.readUInt32LE(offset) === signature) {
      return offset;
    }
  }
  return -1;
}

function inflateZipEntry(buffer, compression) {
  if (compression === 0) return buffer;
  if (compression === 8) return zlib.inflateRawSync(buffer);
  throw new Error(`Unsupported ZIP compression method: ${compression}`);
}

function compareNumberedPath(a, b) {
  return extractFirstNumber(a) - extractFirstNumber(b);
}

function extractFirstNumber(value) {
  return Number(String(value).match(/\d+/)?.[0] || 0);
}

function extractXmlText(xml) {
  const chunks = [];
  const regex = /<(?:[a-z]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[a-z]+:)?t>/g;
  let match;
  while ((match = regex.exec(xml))) {
    chunks.push(decodeXml(match[1]));
  }
  return normalizeWhitespace(chunks.join("\n"));
}

function extractDocxParagraphs(xml) {
  const paragraphs = xml.match(/<w:p\b[\s\S]*?<\/w:p>/g) || [];
  const text = paragraphs.map(extractXmlText).filter(Boolean);
  return text.length ? text : [extractXmlText(xml)].filter(Boolean);
}

function decodeXml(value) {
  return String(value)
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'");
}

function extractPdfText(raw) {
  const chunks = [];
  let match;

  const stringRegex = /\((?:\\.|[^\\)])*\)\s*Tj/g;
  while ((match = stringRegex.exec(raw))) {
    chunks.push(decodePdfLiteral(match[0].replace(/\s*Tj$/, "")));
  }

  const arrayRegex = /\[((?:.|\n|\r)*?)\]\s*TJ/g;
  while ((match = arrayRegex.exec(raw))) {
    const inner = match[1];
    const literals = inner.match(/\((?:\\.|[^\\)])*\)/g) || [];
    chunks.push(literals.map(decodePdfLiteral).join(""));
  }

  const hexRegex = /<([0-9A-Fa-f\s]+)>\s*Tj/g;
  while ((match = hexRegex.exec(raw))) {
    chunks.push(decodePdfHex(match[1]));
  }

  return normalizeWhitespace(chunks.join("\n"));
}

function decodePdfLiteral(value) {
  const literal = String(value).replace(/^\(/, "").replace(/\)$/, "");
  return literal
    .replace(/\\([nrtbf\\()])/g, (_, escaped) => {
      const map = { n: "\n", r: "\r", t: "\t", b: "\b", f: "\f", "\\": "\\", "(": "(", ")": ")" };
      return map[escaped] || escaped;
    })
    .replace(/\\([0-7]{1,3})/g, (_, octal) => String.fromCharCode(parseInt(octal, 8)));
}

function decodePdfHex(value) {
  const hex = String(value).replace(/\s+/g, "");
  if (!hex) return "";
  const bytes = Buffer.from(hex.length % 2 ? `${hex}0` : hex, "hex");
  if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    const chars = [];
    for (let index = 2; index + 1 < bytes.length; index += 2) {
      chars.push(String.fromCharCode(bytes.readUInt16BE(index)));
    }
    return chars.join("");
  }
  return bytes.toString("latin1");
}

function chunkText(text, maxLength) {
  const normalized = normalizeWhitespace(text);
  if (!normalized) return [];
  const chunks = [];
  for (let index = 0; index < normalized.length; index += maxLength) {
    chunks.push(normalized.slice(index, index + maxLength));
  }
  return chunks;
}

function pagesToText(pages) {
  return pages
    .map((page) => `第 ${page.number} 頁：${page.title}\n${page.text}`)
    .join("\n\n");
}

function firstLine(text) {
  return String(text || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .find(Boolean)
    ?.slice(0, 80) || "";
}

function normalizeWhitespace(text) {
  return String(text || "")
    .replace(/\u0000/g, "")
    .replace(/[ \t]+/g, " ")
    .replace(/\s*\n\s*/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function xmlEscape(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function safeFilename(value) {
  const cleaned = String(value || "eduscript-ai-lesson.pptx")
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, "-")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .slice(0, 120)
    .replace(/^-|-$/g, "");
  if (!cleaned || cleaned === ".pptx") return "eduscript-ai-lesson.pptx";
  return cleaned.toLowerCase().endsWith(".pptx") ? cleaned : `${cleaned || "eduscript-ai-lesson"}.pptx`;
}

function safeArchiveFilename(value) {
  const cleaned = String(value || "eduscript-ai-course-pack.zip")
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, "-")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .slice(0, 120)
    .replace(/^-|-$/g, "");
  if (!cleaned || cleaned === ".zip") return "eduscript-ai-course-pack.zip";
  return cleaned.toLowerCase().endsWith(".zip") ? cleaned : `${cleaned || "eduscript-ai-course-pack"}.zip`;
}

function serveStatic(pathname, res) {
  const safePath = pathname === "/" ? "/index.html" : pathname;
  const filePath = path.normalize(path.join(ROOT, safePath));

  if (!filePath.startsWith(ROOT)) {
    return sendJson(res, 403, { error: "Forbidden" });
  }

  fs.readFile(filePath, (error, data) => {
    if (error) {
      return sendJson(res, 404, { error: "Not found" });
    }
    const contentType = mimeTypes[path.extname(filePath)] || "application/octet-stream";
    res.writeHead(200, {
      "Content-Type": contentType,
      "Cache-Control": "no-store",
    });
    res.end(data);
  });
}

function readJsonBody(req, maxLength = 1_000_000) {
  return new Promise((resolve, reject) => {
    let raw = "";
    req.on("data", (chunk) => {
      raw += chunk;
      if (raw.length > maxLength) {
        req.destroy();
        reject(new Error("Request body too large"));
      }
    });
    req.on("end", () => {
      try {
        resolve(raw ? JSON.parse(raw) : {});
      } catch {
        reject(new Error("Invalid JSON body"));
      }
    });
    req.on("error", reject);
  });
}

function sendJson(res, status, data) {
  res.writeHead(status, { "Content-Type": "application/json; charset=utf-8" });
  res.end(JSON.stringify(data));
}

function loadEnvFile(filePath) {
  if (!fs.existsSync(filePath)) return;
  const lines = fs.readFileSync(filePath, "utf8").split(/\r?\n/);
  for (const line of lines) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith("#")) continue;
    const separator = trimmed.indexOf("=");
    if (separator === -1) continue;
    const key = trimmed.slice(0, separator).trim();
    const rawValue = trimmed.slice(separator + 1).trim();
    const value = rawValue.replace(/^["']|["']$/g, "");
    if (key && process.env[key] === undefined) {
      process.env[key] = value;
    }
  }
}
