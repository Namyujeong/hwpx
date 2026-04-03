const express = require("express");
const AdmZip = require("adm-zip");
const { DOMParser, XMLSerializer } = require("xmldom");

const app = express();
app.use(express.json({ limit: "10mb" }));

// ===== 입력 파싱 =====
function splitSections(text) {
  const sections = {};
  const lines = text.split("\n");

  let current = null;

  lines.forEach(line => {
    if (/^\d+\./.test(line)) {
      current = line.replace(/^\d+\.\s*/, "").trim();
      sections[current] = [];
    } else if (current) {
      sections[current].push(line);
    }
  });

  return sections;
}

// ===== 블록 분리 =====
function parseBlocks(lines) {
  let blocks = [];
  let tableBuffer = [];

  lines.forEach(line => {
    if (line.includes("\t")) {
      tableBuffer.push(line);
    } else {
      if (tableBuffer.length) {
        blocks.push({ type: "table", data: tableBuffer });
        tableBuffer = [];
      }
      blocks.push({ type: "text", data: line });
    }
  });

  if (tableBuffer.length) {
    blocks.push({ type: "table", data: tableBuffer });
  }

  return blocks;
}

// ===== 스타일 복제 문단 =====
function cloneParagraphWithText(doc, templateP, text) {
  const newP = templateP.cloneNode(true);

  const tNode = newP.getElementsByTagName("hp:t")[0];
  if (tNode) {
    tNode.textContent = text;
  }

  return newP;
}

// ===== 표 관련 =====
function findNextTable(node) {
  let current = node.nextSibling;

  while (current) {
    if (current.nodeName === "hp:tbl") return current;
    current = current.nextSibling;
  }

  return null;
}

function appendRowsToTable(tableNode, tableData) {
  const rowTemplate = tableNode.getElementsByTagName("hp:tr")[0];
  if (!rowTemplate) return;

  tableData.forEach(row => {
    const newRow = rowTemplate.cloneNode(true);
    const cells = newRow.getElementsByTagName("hp:t");

    row.split("\t").forEach((val, i) => {
      if (cells[i]) cells[i].textContent = val;
    });

    tableNode.appendChild(newRow);
  });
}

// ===== 새 표 생성 =====
function createTable(doc, tableData) {
  const tbl = doc.createElement("hp:tbl");

  tableData.forEach(row => {
    const tr = doc.createElement("hp:tr");
    const cells = row.split("\t");

    cells.forEach(cell => {
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

// ===== 핵심 삽입 =====
function injectContent(doc, sections) {
  const textNodes = doc.getElementsByTagName("hp:t");

  for (let i = 0; i < textNodes.length; i++) {
    const node = textNodes[i];
    const text = node.textContent;

    Object.keys(sections).forEach(title => {

      if (!text.includes(title)) return;

      const parentP = node.parentNode.parentNode; // hp:p

      const lines = sections[title];
      if (!lines || lines.length === 0) return;

      const blocks = parseBlocks(lines);

      let insertAfter = parentP;

      blocks.forEach(block => {

        // ===== 문단 (스타일 복제) =====
        if (block.type === "text") {
          const newP = cloneParagraphWithText(doc, parentP, block.data);

          insertAfter.parentNode.insertBefore(newP, insertAfter.nextSibling);
          insertAfter = newP;
        }

        // ===== 표 =====
        if (block.type === "table") {

          const nextTable = findNextTable(insertAfter);

          // 1️⃣ 템플릿 표 → 스타일 유지
          if (nextTable) {
            appendRowsToTable(nextTable, block.data);
            insertAfter = nextTable;
          }
          // 2️⃣ 없으면 생성
          else {
            const newTable = createTable(doc, block.data);

            insertAfter.parentNode.insertBefore(newTable, insertAfter.nextSibling);
            insertAfter = newTable;
          }
        }

      });

    });
  }

  return doc;
}

// ===== API =====
app.post("/generate-hwpx", (req, res) => {
  try {
    const { content } = req.body;

    if (!content) {
      return res.status(400).json({ error: "content 없음" });
    }

    const zip = new AdmZip("template.hwpx");
    const entry = zip.getEntry("Contents/section0.xml");

    if (!entry) {
      return res.status(500).json({ error: "section0.xml 없음" });
    }

    const xml = entry.getData().toString("utf-8");

    const doc = new DOMParser().parseFromString(xml, "text/xml");

    const sections = splitSections(content);

    const updatedDoc = injectContent(doc, sections);

    const newXml = new XMLSerializer().serializeToString(updatedDoc);

    zip.updateFile("Contents/section0.xml", Buffer.from(newXml));

    const buffer = zip.toBuffer();

    res.json({
      file: buffer.toString("base64"),
      filename: "result.hwpx"
    });

  } catch (err) {
    console.error("❌ 오류:", err);
    res.status(500).json({ error: "HWPX 생성 실패" });
  }
});

// ===== 서버 =====
const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log("🚀 server running");
});
