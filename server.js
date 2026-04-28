const http = require("node:http");
const fs = require("node:fs");
const path = require("node:path");
const zlib = require("node:zlib");
const { URL } = require("node:url");

loadEnvFile(path.join(__dirname, ".env"));

const PORT = Number(process.env.PORT || 4173);
const HOST = process.env.HOST || "0.0.0.0";
const AI_PROVIDER = normalizeProvider(process.env.AI_PROVIDER || "auto");
const OPENAI_MODEL = process.env.OPENAI_MODEL || "gpt-5.2";
const OPENAI_API_KEY = process.env.OPENAI_API_KEY || "";
const GEMINI_MODEL = process.env.GEMINI_MODEL || "gemini-3-flash-preview";
const GEMINI_API_KEY = process.env.GEMINI_API_KEY || process.env.GOOGLE_API_KEY || "";
const GOOGLE_DRIVE_CLIENT_ID = process.env.GOOGLE_DRIVE_CLIENT_ID || "";
const PUBLIC_BASE_URL = process.env.PUBLIC_BASE_URL || process.env.RENDER_EXTERNAL_URL || "";
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
      minItems: 6,
      maxItems: 12,
      items: {
        type: "object",
        additionalProperties: false,
        required: ["title", "event", "bloom", "minutes", "activity", "notes"],
        properties: {
          title: { type: "string" },
          event: { type: "string" },
          bloom: { type: "string" },
          minutes: { type: "number" },
          activity: { type: "string" },
          notes: { type: "string" },
        },
      },
    },
  },
};

const scriptSchema = {
  type: "object",
  additionalProperties: false,
  required: ["script", "teachingNotes"],
  properties: {
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
        availableProviders: {
          openai: Boolean(OPENAI_API_KEY),
          gemini: Boolean(GEMINI_API_KEY),
        },
        environment: process.env.NODE_ENV || "development",
        publicBaseUrl: PUBLIC_BASE_URL,
        googleDriveConfigured: Boolean(GOOGLE_DRIVE_CLIENT_ID),
      });
    }

    if (req.method === "GET" && url.pathname === "/api/config") {
      return sendJson(res, 200, {
        googleDriveClientId: GOOGLE_DRIVE_CLIENT_ID,
        publicBaseUrl: PUBLIC_BASE_URL,
      });
    }

    if (req.method === "POST" && url.pathname.startsWith("/api/ai/")) {
      return handleAiRequest(req, res, url.pathname);
    }

    if (req.method === "POST" && url.pathname === "/api/parse-material") {
      return handleMaterialParse(req, res);
    }

    if (req.method === "POST" && url.pathname === "/api/export-pptx") {
      return handlePptxExport(req, res);
    }

    if (req.method === "POST" && url.pathname === "/api/export-course-pack") {
      return handleCoursePackExport(req, res);
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
  const mode = provider ? `AI enabled (${provider.name}: ${provider.model})` : "local fallback only";
  const url = PUBLIC_BASE_URL || `http://localhost:${PORT}`;
  console.log(`EduScript AI Studio running at ${url} - ${mode}`);
});

async function handleAiRequest(req, res, pathname) {
  const provider = resolveAiProvider();
  if (!provider) {
    return sendJson(res, 503, {
      error: "No AI provider key is set. Add OPENAI_API_KEY or GEMINI_API_KEY, or use the frontend local fallback generation.",
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
        ? extractXmlText(entries.get(notesName).toString("utf8"))
        : "";
      const text = [slideText, notesText && `講者備註：${notesText}`].filter(Boolean).join("\n");
      return {
        number: index + 1,
        title: firstLine(text) || `投影片 ${index + 1}`,
        text,
      };
    });

    return {
      filename,
      type: "pptx",
      pages,
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
    entries.set(`ppt/slides/_rels/slide${slide.number}.xml.rels`, slideRelsXml());
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
  const meta = [slide.event, slide.bloom, slide.minutes && `${slide.minutes} 分鐘`].filter(Boolean).join(" / ");
  const body = [
    slide.activity && `活動：${slide.activity}`,
    ...slide.notes.split(/\r?\n/).filter(Boolean).slice(0, 7),
  ].filter(Boolean);
  const footer = `AI-Assisted Generation / Human review required / ${inputs.subject || ""}`;
  const scriptCue = firstLine(script).slice(0, 100);

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      ${groupShapeXml()}
      ${textShapeXml(2, "title", 560000, 420000, 11080000, 720000, [title], 3200, true, "173B63")}
      ${textShapeXml(3, "meta", 620000, 1180000, 10860000, 360000, [meta], 1500, false, "5E6F73")}
      ${textShapeXml(4, "body", 720000, 1780000, 10600000, 3900000, body, 1700, false, "1F2A2E")}
      ${textShapeXml(5, "script cue", 720000, 5820000, 8400000, 380000, [scriptCue ? `講稿提示：${scriptCue}` : ""], 1100, false, "667276")}
      ${textShapeXml(6, "footer", 9200000, 6220000, 2500000, 260000, [footer], 850, false, "8A8F91")}
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
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

function slideRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
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
    return createGeminiStructuredResponse({ schema, input });
  }

  return createOpenAiStructuredResponse({ schemaName, schema, input });
}

async function createOpenAiStructuredResponse({ schemaName, schema, input }) {
  const response = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${OPENAI_API_KEY}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: OPENAI_MODEL,
      instructions: AI_INSTRUCTIONS,
      input,
      text: {
        format: {
          type: "json_schema",
          name: schemaName,
          strict: true,
          schema,
        },
      },
    }),
  });

  const payload = await response.json().catch(() => ({}));

  if (!response.ok) {
    const detail = payload.error?.message || response.statusText;
    throw new Error(`OpenAI request failed: ${detail}`);
  }

  const text = payload.output_text || extractOutputText(payload);
  if (!text) {
    throw new Error("OpenAI response did not include output_text.");
  }

  return parseStructuredJson(text, "OpenAI");
}

async function createGeminiStructuredResponse({ schema, input }) {
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
      generationConfig: {
        responseMimeType: "application/json",
        responseJsonSchema: toGeminiJsonSchema(schema),
      },
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

function buildLessonPrompt({ inputs }) {
  return `請生成一份專業教材草稿。

課題：${inputs.topic}
科目：${inputs.subject}
對象：${inputs.audience}
分鐘：${inputs.duration}
風格：${inputs.style}
學習目標：${inputs.objective}
先備知識與班情：${inputs.context}
Bloom 層次：${(inputs.bloom || []).join(", ")}

要求：
1. 用逆向設計思維先對齊成果與評量。
2. 投影片要對應 Gagne 九大教學事件。
3. 每頁 notes 必須包含教師可直接使用的講法、互動提示與檢核方式。
4. questions 是 AI 還需要追問教師的關鍵問題。`;
}

function buildScriptPrompt({ inputs, material, startPage, minutes, budget, wpm, targetWords }) {
  const minimumWords = Math.round(Number(targetWords || 0) * 0.92);
  const maximumWords = Math.round(Number(targetWords || 0) * 1.12);
  return `請根據教材與目前進度生成「完整講義式課堂講稿」。

課題：${inputs.topic}
科目：${inputs.subject}
對象：${inputs.audience}
起始頁：${startPage}
目標分鐘：${minutes}
核心講授分鐘：${budget.core}
建議 WPM：${wpm}
核心講授目標字數：約 ${targetWords}
最低可接受字數：${minimumWords}
最高建議字數：${maximumWords}
風格：${inputs.style}

教材內容：
${String(material || "").slice(0, 12000)}

要求：
1. 這不是摘要，不是投影片 prompt，也不是只給老師看的提示；請寫成學生自行閱讀也能理解該堂內容的完整講義式講稿。
2. script 字數必須接近核心講授目標字數，最少 ${minimumWords} 字。不要只產生 1000-2000 字短稿。
3. 請用清楚段落展開：前情提要、核心概念、操作步驟、例子/比喻、命令或 YAML 解讀、常見錯誤、checkpoint、自學總結。
4. 若教材是 PPT Prompt，請把 prompt 轉化成可讀講義內容，不要重複「版面設計」「視覺元素」等 prompt 指令。
5. 不要大量使用「請教師補充」佔位；除非是薪資、最新市場數據、學校政策等必須查證的資料，其他技術概念要直接解釋。
6. 講稿要可口語朗讀，也要可直接發給學生閱讀。
7. teachingNotes 只提供教師課前提醒，不要把主要內容放在 teachingNotes。`;
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

function extractOutputText(payload) {
  const chunks = [];
  for (const item of payload.output || []) {
    for (const content of item.content || []) {
      if (content.type === "output_text" && content.text) {
        chunks.push(content.text);
      }
    }
  }
  return chunks.join("\n");
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

function resolveAiProvider() {
  if (AI_PROVIDER === "openai") {
    return OPENAI_API_KEY ? { name: "openai", model: OPENAI_MODEL } : null;
  }

  if (AI_PROVIDER === "gemini") {
    return GEMINI_API_KEY ? { name: "gemini", model: GEMINI_MODEL } : null;
  }

  if (OPENAI_API_KEY) {
    return { name: "openai", model: OPENAI_MODEL };
  }

  if (GEMINI_API_KEY) {
    return { name: "gemini", model: GEMINI_MODEL };
  }

  return null;
}

function normalizeProvider(value) {
  const provider = String(value || "auto").trim().toLowerCase();
  return ["auto", "openai", "gemini"].includes(provider) ? provider : "auto";
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
    res.writeHead(200, { "Content-Type": contentType });
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
