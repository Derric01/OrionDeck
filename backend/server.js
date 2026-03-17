const path = require("path");
require("dotenv").config();
const express = require("express");
const cors = require("cors");

const uploadRouter = require("./routes/upload");
const reportRouter = require("./routes/report");
const chatRouter = require("./routes/chat");

const app = express();
const PORT = process.env.PORT || 3001;

app.use(cors({ origin: "*" }));
app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ extended: true, limit: "50mb" }));

// Routes
app.use("/upload", uploadRouter);
app.use("/", reportRouter);
app.use("/chat", chatRouter);

// Serve built frontend in production (single-repo deploy)
if (process.env.NODE_ENV === "production") {
  const distDir = path.join(__dirname, "..", "frontend", "dist");
  app.use(express.static(distDir));
  app.get("*", (req, res) => {
    res.sendFile(path.join(distDir, "index.html"));
  });
}

// Health check
app.get("/health", (req, res) => {
  res.json({ status: "ok", service: "Braind Portfolio Reporting Engine", version: "1.0.0" });
});

app.listen(PORT, () => {
  console.log(`Braind Portfolio Engine running on http://localhost:${PORT}`);
});

module.exports = app;
