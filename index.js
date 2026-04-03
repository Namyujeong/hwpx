const express = require("express");
const fs = require("fs");
const app = express();

app.use(express.json());

// 헬스 체크
app.get("/", (req, res) => {
  res.send("ok");
});

// 🔥 HWPX 생성 API (핵심)
app.post("/generate-hwpx", (req, res) => {
  console.log("요청 받음:", req.body);

  try {
    // 👉 실제로는 여기서 HWPX 생성 로직 들어감
    const filePath = "result.hwpx";

    // 테스트용: 파일이 있다고 가정
    const file = fs.readFileSync(filePath);

    res.json({
      file: file.toString("base64"),
      filename: "result.hwpx"
    });

  } catch (error) {
    console.error(error);
    res.status(500).json({
      error: "파일 생성 실패"
    });
  }
});

// 서버 실행
const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
