const express = require("express");
const AdmZip = require("adm-zip");
const { DOMParser, XMLSerializer } = require("xmldom");

const app = express();
app.use(express.json({ limit: "20mb" }));

// ===== 로그 유틸 =====
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
  let sections = [];

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

  lines.forEach(line => {
    if (/^#+\s|^\d+\./.test(line)) {
      current = line.replace(/^#+\s*/, "").replace(/^\d+\.\s*/, "").trim();
      sections[current] = [];
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
  let blocks = [];
  let buffer = [];
  let mode = null;

  lines.forEach(line => {

    if (isTable(line)) {
      if (mode !== "table") {
        if (buffer.length) blocks.push({ type: mode, data: buffer });
        buffer = [];
        mode = "table";
      }
      buffer.push(line);
    } else {
      if (mode !== "text") {
        if (buffer.length) blocks.push({ type: mode, data: buffer });
        buffer = [];
        mode = "text";
      }
      buffer.push(line);
    }

  });

  if (buffer.length) blocks.push({ type: mode, data: buffer });

  return blocks;
}

// ===== 표 변환 =====
function parseMarkdownTable(lines) {
  return lines
    .filter(l => !l.includes("---"))
    .map(line =>
      line.split("|")
        .map(c => c.trim())
        .filter(Boolean)
    );
}

// ===== 매칭 =====
function matchSection(templateSections, inputSections) {
  log("섹션 매핑 시작");

  const map = {};

  Object.keys(inputSections).forEach(inputTitle => {
    const normInput = normalize(inputTitle);

    let best = null;
    let score = 0;

    templateSections.forEach(t => {
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
function cloneParagraph(doc, template, text) {
  const newP = template.cloneNode(true);
  const tNode = newP.getElementsByTagName("hp:t")[0];
  if (tNode) tNode.textContent = text;
  return newP;
}

function createTable(doc, data) {
  const tbl = doc.createElement("hp:tbl");

  data.forEach(row => {
    const tr = doc.createElement("hp:tr");

    row.forEach(cell => {
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

    log("렌더링 섹션", section.title);

    let cursor = anchor;
    const blocks = parseBlocks(lines);

    blocks.forEach(block => {

      if (block.type === "text") {
        block.data.forEach(line => {
          if (!line.trim()) return;

          const p = cloneParagraph(doc, anchor, line);
          cursor.parentNode.insertBefore(p, cursor.nextSibling);
          cursor = p;
        });
      }

      if (block.type === "table") {
        const data = parseMarkdownTable(block.data);

        log("표 생성", data.length + " rows");

        const tbl = createTable(doc, data);
        cursor.parentNode.insertBefore(tbl, cursor.nextSibling);
        cursor = tbl;
      }

    });

  });

  return doc;
}

// ===== API =====
app.post("/generate-hwpx", (req, res) => {
  try {
    log("요청 시작");

    const { content } = req.body;

    if (!content) {
      throw new Error("content 없음");
    }

    // 파일 존재 체크
    const zip = new AdmZip("template.hwpx");
    log("템플릿 로드 완료");

    const entry = zip.getEntry("Contents/section0.xml");

    if (!entry) {
      throw new Error("section0.xml 없음");
    }

    const xml = entry.getData().toString("utf-8");
    log("XML 로드 완료");

    const doc = new DOMParser().parseFromString(xml, "text/xml");

    const templateSections = extractSections(doc);
    const inputSections = splitSections(content);

    const mapping = matchSection(templateSections, inputSections);

    if (Object.keys(mapping).length === 0) {
      throw new Error("섹션 매핑 실패");
    }

    const updated = render(doc, templateSections, mapping);

    const newXml = new XMLSerializer().serializeToString(updated);

    zip.updateFile("Contents/section0.xml", Buffer.from(newXml));

    const buffer = zip.toBuffer();

    // 🔥 파일 검증
    if (!buffer || buffer.length < 2000) {
      throw new Error("파일 생성 실패 (용량 이상)");
    }

    log("파일 생성 성공", buffer.length + " bytes");

    res.json({
      success: true,
      file: buffer.toString("base64"),
      filename: "result.hwpx"
    });

  } catch (err) {
    console.error("❌ ERROR:", err.message);

    res.status(500).json({
      success: false,
      error: err.message
    });
  }
});

// ===== 서버 =====
app.listen(process.env.PORT || 3000, () => {
  console.log("🚀 DEBUG SERVER RUNNING");
});
