const express = require("express");
const AdmZip = require("adm-zip");

const app = express();
app.use(express.json({ limit: "10mb" }));

// ===== 유틸 =====
function escapeXml(str = "") {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

// ===== 템플릿 분석 =====
function extractTextNodes(xml) {
  const matches = [...xml.matchAll(/<hp:t>(.*?)<\/hp:t>/g)];
  return matches.map((m, i) => ({
    index: i,
    text: m[1]
  }));
}

function extractTables(xml) {
  return [...xml.matchAll(/<hp:tbl[\s\S]*?<\/hp:tbl>/g)]
    .map(m => m[0]);
}

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

// ===== 표 처리 =====
function extractRowTemplate(tableXml) {
  const match = tableXml.match(/<hp:tr[\s\S]*?<\/hp:tr>/);
  return match ? match[0] : null;
}

function fillRow(rowTemplate, rowData) {
  let i = 0;
  return rowTemplate.replace(/<hp:t>.*?<\/hp:t>/g, () => {
    const val = rowData[i++] || "";
    return `<hp:t>${escapeXml(val)}</hp:t>`;
  });
}

// 기존 표에 행 추가 (스타일 유지)
function buildTableFromTemplate(tableXml, tableData) {
  const rowTemplate = extractRowTemplate(tableXml);
  if (!rowTemplate) return tableXml;

  const newRows = tableData.map(row => {
    const cells = row.split("\t");
    return fillRow(rowTemplate, cells);
  }).join("");

  return tableXml.replace(
    rowTemplate,
    rowTemplate + newRows
  );
}

// 새 표 생성 (템플릿에 없을 경우)
function createTableXml(tableData) {
  let xml = `<hp:tbl>`;

  tableData.forEach(row => {
    const cells = row.split("\t");

    xml += `<hp:tr>`;

    cells.forEach(cell => {
      xml += `
<hp:tc>
  <hp:p>
    <hp:run>
      <hp:t>${escapeXml(cell)}</hp:t>
    </hp:run>
  </hp:p>
</hp:tc>`;
    });

    xml += `</hp:tr>`;
  });

  xml += `</hp:tbl>`;
  return xml;
}

// ===== 위치 찾기 =====
function findPositions(nodes, sections) {
  const map = {};

  nodes.forEach((node, i) => {
    Object.keys(sections).forEach(title => {
      if (node.text.includes(title)) {
        map[title] = i;
      }
    });
  });

  return map;
}

// ===== 핵심 삽입 =====
function injectContentSafe(xml, nodes, sections, positions) {
  let idx = 0;

  const tables = extractTables(xml);
  let tableCursor = 0;

  return xml.replace(/<hp:t>(.*?)<\/hp:t>/g, (match) => {
    let result = match;

    Object.entries(positions).forEach(([title, pos]) => {
      if (idx === pos) {

        // 공란 처리
        if (!sections[title] || sections[title].length === 0) return;

        const blocks = parseBlocks(sections[title]);
        let sectionXml = "";

        blocks.forEach(block => {

          // ===== 문단 =====
          if (block.type === "text") {
            sectionXml += `
<hp:p>
  <hp:run>
    <hp:t>${escapeXml(block.data)}</hp:t>
  </hp:run>
</hp:p>`;
          }

          // ===== 표 =====
          if (block.type === "table") {

            let tableXml;

            // 1️⃣ 템플릿 표 사용
            if (tables[tableCursor]) {
              tableXml = buildTableFromTemplate(
                tables[tableCursor],
                block.data
              );
              tableCursor++;
            }
            // 2️⃣ 없으면 새로 생성
            else {
              tableXml = createTableXml(block.data);
            }

            sectionXml += tableXml;
          }

        });

        result = match + sectionXml;
      }
    });

    idx++;
    return result;
  });
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

    let xml = entry.getData().toString("utf-8");

    const nodes = extractTextNodes(xml);
    const sections = splitSections(content);
    const positions = findPositions(nodes, sections);

    const newXml = injectContentSafe(xml, nodes, sections, positions);

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
