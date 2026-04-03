const express = require("express");
const AdmZip = require("adm-zip");

const app = express();
app.use(express.json({ limit: "10mb" }));

// ===== 유틸 =====
function escapeXml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// ===== 1. 템플릿 구조 분석 =====
function extractTextNodes(xml) {
  const matches = [...xml.matchAll(/<hp:t>(.*?)<\/hp:t>/g)];
  return matches.map((m, i) => ({
    index: i,
    text: m[1]
  }));
}

// ===== 2. 사용자 입력 파싱 =====
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

// ===== 3. 표 / 문단 분리 =====
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

// ===== 4. 스타일 유지 문단 생성 =====
function cloneParagraphTemplate(xml) {
  const match = xml.match(/<hp:p[\s\S]*?<\/hp:p>/);
  return match ? match[0] : null;
}

function buildParagraphs(templateP, text) {
  return text.split("\n").map(line => {
    return templateP.replace(
      /<hp:t>.*?<\/hp:t>/,
      `<hp:t>${escapeXml(line)}</hp:t>`
    );
  }).join("");
}

// ===== 5. 표 처리 (스타일 유지 핵심) =====
function extractTable(xml) {
  const match = xml.match(/<hp:tbl[\s\S]*?<\/hp:tbl>/);
  return match ? match[0] : null;
}

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

function buildTable(tableXml, tableData) {
  const rowTemplate = extractRowTemplate(tableXml);

  const newRows = tableData.map(row => {
    const cells = row.split("\t");
    return fillRow(rowTemplate, cells);
  }).join("");

  return tableXml.replace("</hp:tbl>", `${newRows}</hp:tbl>`);
}

// ===== 6. 제목 위치 찾기 =====
function findPositions(templateNodes, sections) {
  const map = {};

  templateNodes.forEach((node, i) => {
    Object.keys(sections).forEach(title => {
      if (node.text.includes(title)) {
        map[title] = i;
      }
    });
  });

  return map;
}

// ===== 7. XML 삽입 =====
function injectContent(xml, nodes, sections, positions) {
  let idx = 0;

  return xml.replace(/<hp:t>(.*?)<\/hp:t>/g, (match) => {
    let result = match;

    Object.entries(positions).forEach(([title, pos]) => {
      if (idx === pos) {
        const lines = sections[title];
        const blocks = parseBlocks(lines);

        let sectionXml = "";

        blocks.forEach(block => {
          if (block.type === "text") {
            const pTemplate = cloneParagraphTemplate(xml);
            sectionXml += buildParagraphs(pTemplate, block.data);
          }

          if (block.type === "table") {
            const tableTemplate = extractTable(xml);
            if (tableTemplate) {
              sectionXml += buildTable(tableTemplate, block.data);
            }
          }
        });

        result += sectionXml;
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

    const zip = new AdmZip("template.hwpx");
    const entry = zip.getEntry("Contents/section0.xml");

    let xml = entry.getData().toString("utf-8");

    const templateNodes = extractTextNodes(xml);
    const sections = splitSections(content);
    const positions = findPositions(templateNodes, sections);

    xml = injectContent(xml, templateNodes, sections, positions);

    zip.updateFile("Contents/section0.xml", Buffer.from(xml));

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
const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log("🚀 server running");
});
