const express = require("express");
const path = require("path");
const fs = require("fs");
const { getSlides } = require("../utils/slideContent");
const { getPortfolioContext } = require("../utils/chatEngine");

const router = express.Router();
const TEMPLATE_PATH = path.join(__dirname, "..", "Orion_Q4_2025_Portfolio_Report.pptx");

// Return slide content for in-chat rendering
router.post("/generate-report", (req, res) => {
  try {
    const slides = getSlides();
    const ctx = getPortfolioContext();
    if (!ctx || !ctx.summary) {
      return res.status(400).json({
        success: false,
        error: "No uploaded Excel is loaded. Please upload Orion_Q4_2025_Raw_Data.xlsx first.",
        thinking: ["Missing parsed workbook in session", "Awaiting Excel upload"],
      });
    }
    const s = ctx?.summary || {};
    const occ = s.occupancyPct || "—";
    const abr = s.totalABR != null ? `$${(Number(s.totalABR) / 1e6).toFixed(1)}M` : "—";
    const walt = s.waltYears != null ? `${Number(s.waltYears).toFixed(2)} yrs` : "—";
    const props = s.operatingProperties ?? ctx?.properties?.length ?? "—";
    const leases = s.totalLeases ?? ctx?.leases?.length ?? "—";
    const txs = s.totalTransactions ?? ctx?.transactions?.length ?? "—";
    return res.json({
      success: true,
      message: "Your Q4 2025 Orion Portfolio Report has been generated.",
      thinking: [
        `Loading parsed workbook from session (${props} properties, ${leases} leases, ${txs} transactions)`,
        `KPIs ready: Occupancy ${occ} | ABR ${abr} | WALT ${walt}`,
        "Building slide 1 KPI cards from lease roll",
        "Compiling top tenant roster from ABR weighting",
        "Mapping asset type / industry / geography breakdowns",
        "Computing lease expiry schedule from lease end dates",
        "Building dispositions table from Transactions sheet",
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
