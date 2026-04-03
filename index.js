const express = require("express");
const fs = require("fs");
const AdmZip = require("adm-zip");

const app = express();
app.use(express.json({ limit: "10mb" }));

// ✅ 헬스 체크
app.get("/", (req, res) => {
  res.send("ok");
});

// ✅ 텍스트 → 블록 파싱 (문단 / 표 구분)
function parseContent(text) {
  const lines = text.split("\n");

  let blocks = [];
  let tableBuffer = [];

  for (let line of lines) {
    line = line.trim();

    if (!line) continue;

    if (line.includes("\t")) {
      tableBuffer.push(line);
    } else {
      if (tableBuffer.length > 0) {
        blocks.push({ type: "table", data: tableBuffer });
        tableBuffer = [];
      }
      blocks.push({ type: "text", data: line });
    }
  }

  if (tableBuffer.length > 0) {
    blocks.push({ type: "table", data: tableBuffer });
  }

  return blocks;
}

// ✅ 텍스트 → HWPX 문단 XML
function textToXml(text) {
  return `
<hp:p>
  <hp:run>
    <hp:t>${escapeXml(text)}</hp:t>
  </hp:run>
</hp:p>`;
}

// ✅ 표 → HWPX XML
function tableToXml(rows) {
  let xml = `<hp:tbl>`;

  rows.forEach(row => {
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

// ✅ 전체 XML 생성
function buildXml(blocks) {
  return blocks.map(block => {
    if (block.type === "text") return textToXml(block.data);
    if (block.type === "table") return tableToXml(block.data);
  }).join("\n");
}

// ✅ XML escape (필수)
function escapeXml(str) {
  return str
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

// ✅ HWPX 생성 API
app.post("/generate-hwpx", (req, res) => {
  try {
    const { content } = req.body;

    if (!content) {
      return res.status(400).json({ error: "content 없음" });
    }

    // 🔥 템플릿 로드
    const zip = new AdmZip("template.hwpx");

    // 🔥 본문 XML 가져오기
    const entry = zip.getEntry("Contents/section0.xml");
    let xml = entry.getData().toString("utf-8");

    // 🔥 텍스트 → 구조 변환
    const blocks = parseContent(content);
    const generatedXml = buildXml(blocks);

    // 🔥 템플릿 치환
    xml = xml.replace("{{CONTENT}}", generatedXml);

    // 🔥 다시 삽입
    zip.updateFile("Contents/section0.xml", Buffer.from(xml));

    // 🔥 결과 생성
    const buffer = zip.toBuffer();

    // ✅ GPT Actions 대응 (JSON + Base64)
    res.json({
      file: buffer.toString("base64"),
      filename: "result.hwpx"
    });

  } catch (error) {
    console.error("❌ 오류:", error);
    res.status(500).json({ error: "HWPX 생성 실패" });
  }
});

// ✅ 서버 실행
const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`🚀 Server running on port ${port}`);
});
