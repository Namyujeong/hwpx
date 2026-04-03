const express = require("express");
const app = express();

// 기본 확인용
app.get("/", (req, res) => {
  res.send("ok");
});

const port = process.env.PORT || 3000;
app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
