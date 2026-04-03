const express = require("express");
const AdmZip = require("adm-zip");
const { DOMParser, XMLSerializer } = require("xmldom");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const app = express();
app.set("trust proxy", true);
app.use(express.json({ limit: "20mb" }));

const PORT = process.env.PORT || 3000;
const TEMPLATE_DIR = path.join(__dirname, "templates");
const OUTPUT_DIR = "/tmp/hwpx-output";

const TEMPLATE_MAP = {
  nipa: "nipa.hwpx",
  webtoon: "webtoon.hwpx",
  kipo: "kipo.hwpx"
};

const DEFAULT_SPLIT_MODE = "auto"; // auto | force | off
const DEFAULT_MAX_CHARS_SINGLE = 45000;
const DEFAULT_MAX_CHARS_PER_PART = 22000;
const MAX_WAIT_MS = 25000;
const JOB_TTL_MS = 1000 * 60 * 60 * 6; // 6 hours

const jobs = new Map();

if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

function log(step, data = "") {
  console.log(`🟢 [${step}]`, data || "");
}

function makeError(code, message, detail = undefined) {
  const err = new Error(message);
  err.code = code;
  err.detail = detail;
  return err;
}

function getTemplatePath(templateId) {
  const fileName = TEMPLATE_MAP[templateId];
  if (!fileName) return null;
  return path.join(TEMPLATE_DIR, fileName);
}

function getBaseUrl(req) {
  const proto = req.headers["x-forwarded-proto"] || req.protocol || "https";
  const host = req.headers["x-forwarded-host"] || req.get("host");
  return `${proto}://${host}`;
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function cleanupOldJobs() {
  const now = Date.now();

  for (const [jobId, job] of jobs.entries()) {
    if (now - job.createdAt > JOB_TTL_MS) {
      jobs.delete(jobId);
    }
  }

  try {
    const files = fs.readdirSync(OUTPUT_DIR);
    for (const file of files) {
      const fullPath = path.join(OUTPUT_DIR, file);
      const stat = fs.statSync(fullPath);
      if (now - stat.mtimeMs > JOB_TTL_MS) {
        fs.unlinkSync(fullPath);
      }
    }
  } catch (err) {
    console.error("cleanup error:", err.message);
  }
}

setInterval(cleanupOldJobs, 1000 * 60 * 30);

// ===== 입력 전처리 =====
function sanitizeForXml(text = "") {
  return text
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .normalize("NFKC")
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, "");
}

function flattenCodeFences(text = "") {
  const lines = text.split("\n");
  const output = [];
  let inFence = false;

  for (const line of lines) {
    if (/^\s*```/.test(line)) {
      inFence = !inFence;
      continue;
    }

    if (inFence) {
      output.push(`    ${line}`);
    } else {
      output.push(line);
    }
  }

  return output.join("\n");
}

function preprocessContent(raw = "") {
  return sanitizeForXml(flattenCodeFences(raw))
    .replace(/\t/g, "    ")
    .trim();
}

// ===== 제목/섹션 파싱 =====
function normalizeTitle(text = "") {
  return text.replace(/[\s\.\-ⅠⅡⅢIV0-9()\[\]_:]/g, "").toLowerCase();
}

function parseHeading(line) {
  const md = line.match(/^\s{0,3}#{1,6}\s+(.+?)\s*$/);
  if (md) return md[1].trim();

  const numbered = line.match(/^\s{0,3}\d+(?:\.\d+)*[.)]?\s+(.+?)\s*$/);
  if (numbered) return numbered[1].trim();

  return null;
}

function splitSectionsOrdered(text) {
  const safe = preprocessContent(text);
  if (!safe) return [];

  const lines = safe.split("\n");
  const sections = [];
  let currentTitle = null;
  let currentLines = [];

  for (const line of lines) {
    const heading = parseHeading(line);

    if (heading) {
      if (currentTitle) {
        sections.push({
          title: currentTitle,
          bodyLines: currentLines.slice()
        });
      }
      currentTitle = heading;
      currentLines = [];
      continue;
    }

    if (currentTitle) {
      currentLines.push(line);
    }
  }

  if (currentTitle) {
    sections.push({
      title: currentTitle,
      bodyLines: currentLines.slice()
    });
  }

  // 제목이 하나도 없으면 전체를 하나의 임시 섹션으로 둠
  if (sections.length === 0) {
    return [
      {
        title: "본문",
        bodyLines: safe.split("\n")
      }
    ];
  }

  return sections;
}

function sectionToMarkdown(section) {
  return [`# ${section.title}`, ...section.bodyLines, ""].join("\n");
}

function sectionsToMarkdown(sections) {
  return sections.map(sectionToMarkdown).join("\n").trim();
}

function makePartTitle(sections, index) {
  if (!sections.length) return `Part ${index}`;
  if (sections.length === 1) return sections[0].title;
  return `${sections[0].title} ~ ${sections[sections.length - 1].title}`;
}

function planParts(content, splitMode, maxCharsSingle, maxCharsPerPart) {
  const orderedSections = splitSectionsOrdered(content);
  const fullMarkdown = sectionsToMarkdown(orderedSections);
  const totalLength = fullMarkdown.length;

  const shouldSplit =
    splitMode === "force" ||
    (splitMode === "auto" && totalLength > maxCharsSingle);

  if (!shouldSplit) {
    return [
      {
        partNumber: 1,
        title: "전체",
        content: fullMarkdown
      }
    ];
  }

  const parts = [];
  let current = [];
  let currentLen = 0;

  for (const section of orderedSections) {
    const sectionMarkdown = sectionToMarkdown(section);
    const sectionLen = sectionMarkdown.length;

    if (current.length > 0 && currentLen + sectionLen > maxCharsPerPart) {
      parts.push(current);
      current = [];
      currentLen = 0;
    }

    current.push(section);
    currentLen += sectionLen;
  }

  if (current.length > 0) {
    parts.push(current);
  }

  return parts.map((group, idx) => ({
    partNumber: idx + 1,
    title: makePartTitle(group, idx + 1),
    content: sectionsToMarkdown(group)
  }));
}

// ===== 템플릿 분석 =====
function extractSections(doc) {
  log("템플릿 분석 시작");

  const nodes = doc.getElementsByTagName("hp:p");
  const sections = [];

  for (let i = 0; i < nodes.length; i++) {
    const t = nodes[i].getElementsByTagName("hp:t")[0];
    if (!t) continue;

    const text = (t.textContent || "").trim();

    if (text.length > 0 && text.length < 80) {
      sections.push({
        title: text,
        norm: normalizeTitle(text),
        node: nodes[i],
        index: i
      });
    }
  }

  log("템플릿 섹션 수", sections.length);
  return sections;
}

// ===== 블록 파싱 =====
function isTable(line) {
  return /^\s*\|/.test(line);
}

function parseBlocks(lines) {
  const blocks = [];
  let buffer = [];
  let mode = null;

  for (const line of lines) {
    const nextMode = isTable(line) ? "table" : "text";

    if (mode !== nextMode) {
      if (buffer.length) {
        blocks.push({ type: mode, data: buffer.slice() });
      }
      buffer = [];
      mode = nextMode;
    }

    buffer.push(line);
  }

  if (buffer.length) {
    blocks.push({ type: mode, data: buffer.slice() });
  }

  return blocks;
}

function parseMarkdownTable(lines) {
  return lines
    .filter((line) => !/^\s*\|?[-:\s|]+\|?\s*$/.test(line))
    .map((line) =>
      line
        .split("|")
        .map((cell) => cell.trim())
        .filter(Boolean)
    )
    .filter((row) => row.length > 0);
}

// ===== 섹션 매칭 =====
function matchSections(templateSections, orderedInputSections) {
  log("섹션 매핑 시작");

  const mapping = new Map();
  const unmatched = [];

  for (const inputSection of orderedInputSections) {
    const normInput = normalizeTitle(inputSection.title);

    let best = null;
    let score = 0;

    for (const templateSection of templateSections) {
      if (
        templateSection.norm.includes(normInput) ||
        normInput.includes(templateSection.norm)
      ) {
        const s = Math.min(templateSection.norm.length, normInput.length);
        if (s > score) {
          best = templateSection;
          score = s;
        }
      }
    }

    if (!best) {
      unmatched.push(inputSection.title);
      log("매핑 실패", inputSection.title);
      continue;
    }

    const existing = mapping.get(best.index) || [];
    if (existing.length > 0) {
      existing.push("");
      existing.push(`[${inputSection.title}]`);
    }
    existing.push(...inputSection.bodyLines);

    mapping.set(best.index, existing);
    log("매핑 성공", `${inputSection.title} → ${best.title} (nodeIndex=${best.index})`);
  }

  return { mapping, unmatched };
}

// ===== 렌더 =====
function cloneParagraph(template, text) {
  const newP = template.cloneNode(true);
  const tNode = newP.getElementsByTagName("hp:t")[0];
  if (tNode) {
    tNode.textContent = text;
  }
  return newP;
}

function createTable(doc, data) {
  const tbl = doc.createElement("hp:tbl");

  for (const row of data) {
    const tr = doc.createElement("hp:tr");

    for (const cell of row) {
      const tc = doc.createElement("hp:tc");
      const p = doc.createElement("hp:p");
      const run = doc.createElement("hp:run");
      const t = doc.createElement("hp:t");

      t.appendChild(doc.createTextNode(cell));
      run.appendChild(t);
      p.appendChild(run);
      tc.appendChild(p);
      tr.appendChild(tc);
    }

    tbl.appendChild(tr);
  }

  return tbl;
}

function render(doc, templateSections, mapping) {
  log("렌더링 시작");

  const sectionByNodeIndex = new Map(
    templateSections.map((section) => [section.index, section])
  );

  for (const [nodeIndex, lines] of mapping.entries()) {
    const section = sectionByNodeIndex.get(Number(nodeIndex));

    if (!section) {
      throw makeError(
        "SECTION_LOOKUP_MISSING",
        `템플릿 섹션 조회 실패: ${nodeIndex}`
      );
    }

    const anchor = section.node;
    let cursor = anchor;

    log("렌더링 섹션", `${section.title} (nodeIndex=${nodeIndex})`);

    const blocks = parseBlocks(lines);

    for (const block of blocks) {
      if (block.type === "text") {
        for (const line of block.data) {
          if (!line.trim()) continue;

          const p = cloneParagraph(anchor, line);
          cursor.parentNode.insertBefore(p, cursor.nextSibling);
          cursor = p;
        }
      }

      if (block.type === "table") {
        const data = parseMarkdownTable(block.data);
        if (!data.length) continue;

        log("표 생성", `${data.length} rows`);
        const tbl = createTable(doc, data);
        cursor.parentNode.insertBefore(tbl, cursor.nextSibling);
        cursor = tbl;
      }
    }
  }

  return doc;
}

// ===== HWPX 생성 =====
function createOutputFileName(templateId, partNumber) {
  const suffix = crypto.randomUUID();
  return `${templateId}-part${partNumber}-${suffix}.hwpx`;
}

function generateHwpxFile({ templateId, content, baseUrl, partNumber, title }) {
  const templatePath = getTemplatePath(templateId);
  if (!templatePath || !fs.existsSync(templatePath)) {
    throw makeError("INVALID_TEMPLATE", `유효한 템플릿이 없음: ${templateId}`);
  }

  const zip = new AdmZip(templatePath);
  const entry = zip.getEntry("Contents/section0.xml");

  if (!entry) {
    throw makeError("MISSING_SECTION_XML", "section0.xml 없음");
  }

  const xml = entry.getData().toString("utf-8");
  const doc = new DOMParser().parseFromString(xml, "text/xml");

  const templateSections = extractSections(doc);
  const orderedInputSections = splitSectionsOrdered(content);
  const { mapping, unmatched } = matchSections(templateSections, orderedInputSections);

  if (mapping.size === 0) {
    throw makeError("NO_SECTION_MAPPING", "섹션 매핑 실패");
  }

  const updated = render(doc, templateSections, mapping);
  const newXml = new XMLSerializer().serializeToString(updated);

  zip.updateFile("Contents/section0.xml", Buffer.from(newXml, "utf-8"));

  const buffer = zip.toBuffer();
  if (!buffer || buffer.length < 2000) {
    throw makeError("EMPTY_OUTPUT", "파일 생성 실패");
  }

  const filename = createOutputFileName(templateId, partNumber);
  const filePath = path.join(OUTPUT_DIR, filename);
  fs.writeFileSync(filePath, buffer);

  return {
    partNumber,
    title,
    filename,
    downloadUrl: `${baseUrl}/download/${encodeURIComponent(filename)}`,
    unmatchedHeadings: unmatched
  };
}

// ===== Job 관리 =====
function createJobRecord({ templateId, content, splitMode, maxCharsSingle, maxCharsPerPart, baseUrl }) {
  const jobId = crypto.randomUUID();
  const now = Date.now();

  const job = {
    jobId,
    createdAt: now,
    updatedAt: now,
    status: "queued",
    progress: 0,
    step: "queued",
    templateId,
    content,
    splitMode,
    maxCharsSingle,
    maxCharsPerPart,
    baseUrl,
    split: false,
    totalParts: 0,
    completedParts: 0,
    files: [],
    errorCode: null,
    errorMessage: null
  };

  jobs.set(jobId, job);
  return job;
}

function updateJob(jobId, patch) {
  const job = jobs.get(jobId);
  if (!job) return null;

  Object.assign(job, patch, { updatedAt: Date.now() });
  jobs.set(jobId, job);
  return job;
}

function publicJob(job) {
  if (!job) return null;

  return {
    success: true,
    jobId: job.jobId,
    status: job.status,
    progress: job.progress,
    step: job.step,
    templateId: job.templateId,
    split: job.split,
    totalParts: job.totalParts,
    completedParts: job.completedParts,
    files: job.files,
    errorCode: job.errorCode,
    errorMessage: job.errorMessage
  };
}

async function processJob(jobId) {
  const job = jobs.get(jobId);
  if (!job) return;

  try {
    updateJob(jobId, {
      status: "running",
      step: "preprocessing",
      progress: 5
    });

    const parts = planParts(
      job.content,
      job.splitMode,
      job.maxCharsSingle,
      job.maxCharsPerPart
    );

    updateJob(jobId, {
      step: "planned",
      progress: 15,
      split: parts.length > 1,
      totalParts: parts.length
    });

    const files = [];

    for (let i = 0; i < parts.length; i++) {
      const part = parts[i];

      updateJob(jobId, {
        step: `rendering_part_${part.partNumber}`,
        progress: 15 + Math.round((i / parts.length) * 75)
      });

      const result = generateHwpxFile({
        templateId: job.templateId,
        content: part.content,
        baseUrl: job.baseUrl,
        partNumber: part.partNumber,
        title: part.title
      });

      files.push(result);

      updateJob(jobId, {
        completedParts: i + 1,
        files,
        progress: 15 + Math.round(((i + 1) / parts.length) * 75)
      });
    }

    updateJob(jobId, {
      status: "done",
      step: "done",
      progress: 100,
      files
    });
  } catch (err) {
    console.error("❌ JOB ERROR:", err.message);

    updateJob(jobId, {
      status: "failed",
      step: "failed",
      progress: 100,
      errorCode: err.code || "GENERATION_FAILED",
      errorMessage: err.message
    });
  }
}

// ===== API =====
app.get("/", (req, res) => {
  res.json({
    success: true,
    message: "HWPX Generator API is running"
  });
});

app.get("/health", (req, res) => {
  res.json({
    success: true,
    ok: true
  });
});

app.get("/templates", (req, res) => {
  res.json({
    success: true,
    templates: [
      { id: "nipa", label: "NIPA", fileName: "nipa.hwpx" },
      { id: "webtoon", label: "웹툰", fileName: "webtoon.hwpx" },
      { id: "kipo", label: "지식재산처", fileName: "kipo.hwpx" }
    ]
  });
});

app.post("/jobs", (req, res) => {
  try {
    const {
      templateId,
      content,
      splitMode = DEFAULT_SPLIT_MODE,
      maxCharsSingle = DEFAULT_MAX_CHARS_SINGLE,
      maxCharsPerPart = DEFAULT_MAX_CHARS_PER_PART
    } = req.body || {};

    if (!templateId) {
      throw makeError("MISSING_TEMPLATE_ID", "templateId 없음");
    }

    if (!Object.prototype.hasOwnProperty.call(TEMPLATE_MAP, templateId)) {
      throw makeError("INVALID_TEMPLATE", `유효한 템플릿이 없음: ${templateId}`);
    }

    if (!content || typeof content !== "string") {
      throw makeError("MISSING_CONTENT", "content 없음");
    }

    if (!["auto", "force", "off"].includes(splitMode)) {
      throw makeError("INVALID_SPLIT_MODE", "splitMode는 auto, force, off 중 하나여야 함");
    }

    const baseUrl = getBaseUrl(req);
    const job = createJobRecord({
      templateId,
      content,
      splitMode,
      maxCharsSingle: Number(maxCharsSingle) || DEFAULT_MAX_CHARS_SINGLE,
      maxCharsPerPart: Number(maxCharsPerPart) || DEFAULT_MAX_CHARS_PER_PART,
      baseUrl
    });

    setImmediate(() => {
      processJob(job.jobId).catch((err) => {
        console.error("processJob fatal:", err.message);
      });
    });

    res.status(202).json({
      success: true,
      jobId: job.jobId,
      status: job.status,
      pollUrl: `${baseUrl}/jobs/${job.jobId}`
    });
  } catch (err) {
    res.status(400).json({
      success: false,
      errorCode: err.code || "BAD_REQUEST",
      errorMessage: err.message
    });
  }
});

app.get("/jobs/:jobId", async (req, res) => {
  const { jobId } = req.params;
  const waitMs = Math.min(Number(req.query.waitMs || 0), MAX_WAIT_MS);

  const start = Date.now();

  while (true) {
    const job = jobs.get(jobId);

    if (!job) {
      return res.status(404).json({
        success: false,
        errorCode: "JOB_NOT_FOUND",
        errorMessage: "jobId를 찾을 수 없음"
      });
    }

    if (job.status === "done" || job.status === "failed") {
      return res.json(publicJob(job));
    }

    if (!waitMs || Date.now() - start >= waitMs) {
      return res.json(publicJob(job));
    }

    await sleep(700);
  }
});

app.get("/download/:fileName", (req, res) => {
  const decodedFileName = decodeURIComponent(req.params.fileName);
  const fileName = path.basename(decodedFileName);
  const filePath = path.join(OUTPUT_DIR, fileName);

  if (!fs.existsSync(filePath)) {
    return res.status(404).json({
      success: false,
      errorCode: "FILE_NOT_FOUND",
      errorMessage: "파일이 없음"
    });
  }

  res.download(filePath, fileName);
});

app.listen(PORT, () => {
  console.log(`🚀 HWPX GENERATOR RUNNING ON ${PORT}`);
});
