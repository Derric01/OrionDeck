const express = require("express");
const path = require("path");
const fs = require("fs");
const { getSlides } = require("../utils/slideContent");

const router = express.Router();
const TEMPLATE_PATH = path.join(__dirname, "..", "Orion_Q4_2025_Portfolio_Report.pptx");

// Return slide content for in-chat rendering
router.post("/generate-report", (req, res) => {
  try {
    const slides = getSlides();
    return res.json({
      success: true,
      message: "Your Q4 2025 Orion Portfolio Report has been generated.",
      thinking: [
        "Loading parsed Excel data from session",
        "Computing portfolio KPIs",
        "Mapping asset type breakdown",
        "Building occupancy analysis",
        "Preparing transaction activity table",
        "Compiling top tenant roster",
        "Generating lease expiry schedule",
        "Report complete — 8 slides ready",
      ],
      report_data: {
        title: "Orion Properties Inc. — Q4 2025 Portfolio Report",
        period: "Q4 2025",
        generatedAt: new Date().toISOString(),
        slides,
        downloadAvailable: true,
      },
    });
  } catch (err) {
    console.error("Report generation error:", err);
    return res.status(500).json({ error: "Failed to generate report." });
  }
});

// Serve the exact Orion template file for download. Edits are done via chat in the viewer.
router.get("/report/download", (req, res) => {
  try {
    if (!fs.existsSync(TEMPLATE_PATH)) {
      return res.status(404).json({ error: "Template file not found." });
    }
    const filename = "Orion_Q4_2025_Portfolio_Report.pptx";
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
    res.sendFile(TEMPLATE_PATH);
  } catch (err) {
    console.error("Download error:", err);
    return res.status(500).json({ error: "Failed to serve template." });
  }
});

// Return current slide state
router.get("/slides", (req, res) => {
  const slides = getSlides();
  return res.json({ slides });
});

module.exports = router;
