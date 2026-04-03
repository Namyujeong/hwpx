const express = require("express");
const AdmZip = require("adm-zip");
const { DOMParser, XMLSerializer } = require("xmldom");

const app = express();
app.use(express.json({ limit: "20mb" }));

// ===== 1. 텍스트 정규화 =====
function normalize(text = "") {
  return text
    .replace(/[\s\.\-ⅠⅡⅢIV0-9]/g, "")
    .toLowerCase();
}

// ===== 2. 템플릿 구조 분석 =====
function extractSections(doc) {
  const nodes = doc.getElementsByTagName("hp:p");

  let sections = [];

  for (let i = 0; i < nodes.length; i++) {
    const t = nodes[i].getElementsByTagName("hp:t")[0];
    if (!t) continue;

    const text = t.textContent.trim();

    if (text.length < 30) { // 제목 후보
      sections.push({
        title: text,
        norm: normalize(text),
        node: nodes[i],
        index: i
      });
    }
  }

  // 경계 설정
  for (let i = 0; i < sections.length; i++) {
    sections[i].endIndex =
      i < sections.length - 1 ? sections[i + 1].index : nodes.length;
  }

  return sections;
}

// ===== 3. 입력 파싱 =====
function splitSections(text) {
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

  return sections;
}

// ===== 4. Markdown 파싱 =====
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

function parseMarkdownTable(lines) {
  return lines
    .filter(l => !l.includes("---"))
    .map(line =>
      line.split("|")
        .map(c => c.trim())
        .filter(Boolean)
    );
}

// ===== 5. 매칭 =====
function matchSection(templateSections, inputSections) {
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
    }
  });

  return map;
}

// ===== 6. 렌더링 =====
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
  const nodes = doc.getElementsByTagName("hp:p");

  Object.entries(mapping).forEach(([index, lines]) => {

    const section = templateSections[index];
    const anchor = section.node;

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
    const { content } = req.body;

    const zip = new AdmZip("template.hwpx");
    const entry = zip.getEntry("Contents/section0.xml");

    const xml = entry.getData().toString("utf-8");

    const doc = new DOMParser().parseFromString(xml, "text/xml");

    const templateSections = extractSections(doc);
    const inputSections = splitSections(content);

    const mapping = matchSection(templateSections, inputSections);

    const updated = render(doc, templateSections, mapping);

    const newXml = new XMLSerializer().serializeToString(updated);

    zip.updateFile("Contents/section0.xml", Buffer.from(newXml));

    const buffer = zip.toBuffer();

    res.json({
      file: buffer.toString("base64"),
      filename: "result.hwpx"
    });

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "생성 실패" });
  }
});

// ===== 서버 =====
app.listen(process.env.PORT || 3000, () => {
  console.log("🚀 universal engine running");
});
