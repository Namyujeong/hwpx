const express = require("express");
const AdmZip = require("adm-zip");
const { DOMParser, XMLSerializer } = require("xmldom");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const app = express();
app.use(express.json({ limit: "20mb" }));

const PORT = process.env.PORT || 3000;
const BASE_URL = process.env.BASE_URL || `http://localhost:${PORT}`;
const TEMPLATE_DIR = path.join(__dirname, "templates");
const OUTPUT_DIR = "/tmp/hwpx-output";

if (!fs.existsSync(OUTPUT_DIR)) {
  fs.mkdirSync(OUTPUT_DIR, { recursive: true });
}

// 실제 templates 폴더 안 파일명 기준
const TEMPLATE_MAP = {
  nipa: "nipa.hwpx",
  webtoon: "webtoon.hwpx",
  "지식재산처": "지식재산처.hwpx"
};

// ===== 로그 =====
function log(step, data = "") {
  console.log(`🟢 [${step}]`, data || "");
}

// ===== 텍스트 정규화 =====
function normalize(text = "") {
  return text.replace(/[\s\.\-ⅠⅡⅢIV0-9]/g, "").toLowerCase();
}

// ===== 템플릿 분석 =====
function extractSections(doc) {
  log("템플릿 분석 시작");

  const nodes = doc.getElementsByTagName("hp:p");
  const sections = [];

  for (let i = 0; i < nodes.length; i++) {
    const t = nodes[i].getElementsByTagName("hp:t")[0];
    if (!t) continue;

    const text = t.textContent.trim();

    if (text.length > 0 && text.length < 50) {
      sections.push({
        title: text,
        norm: normalize(text),
        node: nodes[i],
        index: i
      });
    }
  }

  log("템플릿 섹션 수", sections.length);
  return sections;
}

// ===== 입력 파싱 =====
function splitSections(text) {
  log("본문 파싱 시작");

  const sections = {};
  const lines = text.split("\n");
  let current = null;

  lines.forEach((line) => {
    if (/^#+\s|^\d+\.\s/.test(line)) {
      current = line
        .replace(/^#+\s*/, "")
        .replace(/^\d+\.\s*/, "")
        .trim();

      if (current) sections[current] = [];
    } else if (current) {
      sections[current].push(line);
    }
  });

  log("본문 섹션 수", Object.keys(sections).length);
  return sections;
}

// ===== 블록 파싱 =====
function isTable(line) {
  return line.trim().startsWith("|");
}

function parseBlocks(lines) {
  const blocks = [];
  let buffer = [];
  let mode = null;

  lines.forEach((line) => {
    const nextMode = isTable(line) ? "table" : "text";

    if (mode !== nextMode) {
      if (buffer.length) {
        blocks.push({ type: mode, data: buffer });
      }
      buffer = [];
      mode = nextMode;
    }

    buffer.push(line);
  });

  if (buffer.length) {
    blocks.push({ type: mode, data: buffer });
  }

  return blocks;
}

// ===== 표 변환 =====
function parseMarkdownTable(lines) {
  return lines
    .filter((l) => !/^\s*\|?[-:\s|]+\|?\s*$/.test(l))
    .map((line) =>
      line
        .split("|")
        .map((c) => c.trim())
        .filter(Boolean)
    )
    .filter((row) => row.length > 0);
}

// ===== 매칭 =====
function matchSection(templateSections, inputSections) {
  log("섹션 매핑 시작");

  const map = {};

  Object.keys(inputSections).forEach((inputTitle) => {
    const normInput = normalize(inputTitle);

    let best = null;
    let score = 0;

    templateSections.forEach((t) => {
      if (t.norm.includes(normInput) || normInput.includes(t.norm)) {
        const s = Math.min(t.norm.length, normInput.length);
        if (s > score) {
          best = t;
          score = s;
        }
      }
    });

    if (best) {
      map[best.index] = inputSections[inputTitle];
      log("매핑 성공", `${inputTitle} → ${best.title}`);
    } else {
      log("매핑 실패", inputTitle);
    }
  });

  return map;
}

// ===== 렌더 =====
function cloneParagraph(template, text) {
  const newP = template.cloneNode(true);
  const tNode = newP.getElementsByTagName("hp:t")[0];
  if (tNode) tNode.textContent = text;
  return newP;
}

function createTable(doc, data) {
  const tbl = doc.createElement("hp:tbl");

  data.forEach((row) => {
    const tr = doc.createElement("hp:tr");

    row.forEach((cell) => {
      const tc = doc.createElement("hp:tc");
      const p = doc.createElement("hp:p");
      const run = doc.createElement("hp:run");
      const t = doc.createElement("hp:t");

      t.appendChild(doc.createTextNode(cell));
      run.appendChild(t);
      p.appendChild(run);
      tc.appendChild(p);
      tr.appendChild(tc);
    });

    tbl.appendChild(tr);
  });

  return tbl;
}

function render(doc, templateSections, mapping) {
  log("렌더링 시작");

  Object.entries(mapping).forEach(([index, lines]) => {
    const section = templateSections[index];
    const anchor = section.node;
    let cursor = anchor;

    log("렌더링 섹션", section.title);

    const blocks = parseBlocks(lines);

    blocks.forEach((block) => {
      if (block.type === "text") {
        block.data.forEach((line) => {
          if (!line.trim()) return;
          const p = cloneParagraph(anchor, line);
          cursor.parentNode.insertBefore(p, cursor.nextSibling);
          cursor = p;
        });
      }

      if (block.type === "table") {
        const data = parseMarkdownTable(block.data);
        if (!data.length) return;

        log("표 생성", `${data.length} rows`);
        const tbl = createTable(doc, data);
        cursor.parentNode.insertBefore(tbl, cursor.nextSibling);
        cursor = tbl;
      }
    });
  });

  return doc;
}

// ===== 유틸 =====
function getTemplatePath(templateId) {
  const fileName = TEMPLATE_MAP[templateId];
  if (!fileName) return null;
  return path.join(TEMPLATE_DIR, fileName);
}

// ===== 기본 확인 =====
app.get("/", (req, res) => {
  res.json({
    success: true,
    message: "HWPX Generator API is running"
  });
});

app.get("/health", (req, res) => {
  res.json({ ok: true });
});

app.get("/templates", (req, res) => {
  const templates = Object.entries(TEMPLATE_MAP).map(([id, fileName]) => ({
    id,
    fileName
  }));

  res.json({
    success: true,
    templates
  });
});

// ===== 생성 API =====
app.post("/generate-hwpx", (req, res) => {
  try {
    log("요청 시작");

    const { templateId, content } = req.body;

    if (!templateId) {
      throw new Error("templateId 없음");
    }

    if (!content || typeof content !== "string") {
      throw new Error("content 없음");
    }

    const templatePath = getTemplatePath(templateId);
    if (!templatePath || !fs.existsSync(templatePath)) {
      throw new Error(`유효한 템플릿이 없음: ${templateId}`);
    }

    const zip = new AdmZip(templatePath);
    log("템플릿 로드 완료", templatePath);

    const entry = zip.getEntry("Contents/section0.xml");
    if (!entry) {
      throw new Error("section0.xml 없음");
    }

    const xml = entry.getData().toString("utf-8");
    const doc = new DOMParser().parseFromString(xml, "text/xml");

    const templateSections = extractSections(doc);
    const inputSections = splitSections(content);
    const mapping = matchSection(templateSections, inputSections);

    if (Object.keys(mapping).length === 0) {
      throw new Error("섹션 매핑 실패");
    }

    const updated = render(doc, templateSections, mapping);
    const newXml = new XMLSerializer().serializeToString(updated);

    zip.updateFile("Contents/section0.xml", Buffer.from(newXml, "utf-8"));

    const buffer = zip.toBuffer();
    if (!buffer || buffer.length < 2000) {
      throw new Error("파일 생성 실패");
    }

    const fileId = crypto.randomUUID();
    const safeTemplateId = encodeURIComponent(templateId);
    const outputFileName = `${safeTemplateId}-${fileId}.hwpx`;
    const outputPath = path.join(OUTPUT_DIR, outputFileName);

    fs.writeFileSync(outputPath, buffer);

    const downloadUrl = `${BASE_URL}/download/${outputFileName}`;

    log("파일 생성 성공", outputFileName);

    res.json({
      success: true,
      filename: outputFileName,
      downloadUrl
    });
  } catch (err) {
    console.error("❌ ERROR:", err.message);

    res.status(500).json({
      success: false,
      error: err.message
    });
  }
});

// ===== 다운로드 =====
app.get("/download/:fileName", (req, res) => {
  const fileName = path.basename(req.params.fileName);
  const filePath = path.join(OUTPUT_DIR, fileName);

  if (!fs.existsSync(filePath)) {
    return res.status(404).json({
      success: false,
      error: "파일이 없음"
    });
  }

  res.download(filePath, fileName);
});

app.listen(PORT, () => {
  console.log(`🚀 HWPX GENERATOR RUNNING ON ${PORT}`);
});
