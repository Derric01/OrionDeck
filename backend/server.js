const path = require("path");
require("dotenv").config({ path: path.join(__dirname, ".env") });
const express = require("express");
const cors = require("cors");

const uploadRouter = require("./routes/upload");
const reportRouter = require("./routes/report");
const chatRouter = require("./routes/chat");
const { loadWalkthroughWorkbook } = require("./utils/walkthrough");

const app = express();
const PORT = process.env.PORT || 3001;

app.use(cors({ origin: "*" }));
app.use(express.json({ limit: "50mb" }));
app.use(express.urlencoded({ extended: true, limit: "50mb" }));

// Load the quarterly walkthrough workbook once at startup (read-only reference model)
try {
  loadWalkthroughWorkbook(path.join(__dirname, "Quarterly_Presentation_Walkthrough.xlsx"));
} catch (e) {
  console.warn("Walkthrough workbook load warning:", e.message);
}

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
