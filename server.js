const http = require("node:http");
const fs = require("node:fs");
const path = require("node:path");
const zlib = require("node:zlib");
const { URL } = require("node:url");

loadEnvFile(path.join(__dirname, ".env"));

const PORT = Number(process.env.PORT || 4173);
const MODEL = process.env.OPENAI_MODEL || "gpt-5.2";
const OPENAI_API_KEY = process.env.OPENAI_API_KEY || "";
const ROOT = __dirname;

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
      return sendJson(res, 200, {
        ok: true,
        aiEnabled: Boolean(OPENAI_API_KEY),
        model: MODEL,
      });
    }

    if (req.method === "POST" && url.pathname.startsWith("/api/ai/")) {
      return handleAiRequest(req, res, url.pathname);
    }

    if (req.method === "POST" && url.pathname === "/api/parse-material") {
      return handleMaterialParse(req, res);
    }

    if (req.method !== "GET") {
      return sendJson(res, 405, { error: "Method not allowed" });
    }

    return serveStatic(url.pathname, res);
  } catch (error) {
    return sendJson(res, 500, { error: error.message || "Internal server error" });
  }
});

server.listen(PORT, () => {
  const mode = OPENAI_API_KEY ? `AI enabled (${MODEL})` : "local fallback only";
  console.log(`EduScript AI Studio running at http://localhost:${PORT} - ${mode}`);
});

async function handleAiRequest(req, res, pathname) {
  if (!OPENAI_API_KEY) {
    return sendJson(res, 503, {
      error: "OPENAI_API_KEY is not set. The frontend will use local fallback generation.",
    });
  }

  const body = await readJsonBody(req);

  if (pathname === "/api/ai/lesson") {
    const result = await createStructuredResponse({
      schemaName: "lesson_plan",
      schema: lessonSchema,
      input: buildLessonPrompt(body),
    });
    return sendJson(res, 200, result);
  }

  if (pathname === "/api/ai/script") {
    const result = await createStructuredResponse({
      schemaName: "lesson_script",
      schema: scriptSchema,
      input: buildScriptPrompt(body),
    });
    return sendJson(res, 200, result);
  }

  if (pathname === "/api/ai/assistant") {
    const result = await createStructuredResponse({
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

function parseMaterialBuffer(buffer, filename, extension, mimeType) {
  if ([".pptx", ".docx"].includes(extension)) {
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

async function createStructuredResponse({ schemaName, schema, input }) {
  const response = await fetch("https://api.openai.com/v1/responses", {
    method: "POST",
    headers: {
      Authorization: `Bearer ${OPENAI_API_KEY}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      model: MODEL,
      instructions:
        "你是資深教學設計師與教育科技產品 AI。請用繁體中文回答，內容要可直接給教師使用，避免空泛形容詞。",
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

  return JSON.parse(text);
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
  return `請根據教材與目前進度生成接續講稿。

課題：${inputs.topic}
科目：${inputs.subject}
對象：${inputs.audience}
起始頁：${startPage}
目標分鐘：${minutes}
核心講授分鐘：${budget.core}
建議 WPM：${wpm}
核心講授目標字數：約 ${targetWords}
風格：${inputs.style}

教材內容：
${String(material || "").slice(0, 12000)}

要求：
1. 開頭要包含 1 至 2 分鐘前情提要。
2. 講稿必須可口語朗讀，並加入停頓與互動提示。
3. 若教材內容不足，請明確標記教師需補資料的位置。
4. teachingNotes 提供教師課前提醒。`;
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
  const regex = /<(?:a|w):t\b[^>]*>([\s\S]*?)<\/(?:a|w):t>/g;
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
