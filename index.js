const express = require("express");
const app = express();

// JSON 요청 처리
app.use(express.json());

// 1️⃣ 기본 확인 (헬스체크)
app.get("/", (req, res) => {
  res.send("ok");
});

// 2️⃣ 렌더 작업 생성 API
app.post("/createRenderJob", (req, res) => {
  console.log("요청 받음:", req.body);

  res.json({
    job_id: "job-123",
    status: "processing"
  });
});

// 3️⃣ 렌더 상태 조회 API
app.get("/getRenderJobStatus", (req, res) => {
  res.json({
    status: "completed",
    file_url: "https://example.com/sample.hwpx"
  });
});

// 서버 실행 (Render 필수 설정)
const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
