const express = require("express");
const multer = require("multer");
const path = require("path");
const { v4: uuidv4 } = require("uuid");
const { parsePortfolioExcel } = require("../utils/parseExcel");
const { setPortfolioContext } = require("../utils/chatEngine");
const { resetSlides, setSlidesFromPortfolioData } = require("../utils/slideContent");

const router = express.Router();

const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, path.join(__dirname, "../uploads"));
  },
  filename: (req, file, cb) => {
    const id = uuidv4();
    cb(null, `${id}-${file.originalname}`);
  },
});

const upload = multer({
  storage,
  fileFilter: (req, file, cb) => {
    if (
      file.mimetype === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
      file.mimetype === "application/vnd.ms-excel" ||
      file.originalname.endsWith(".xlsx") ||
      file.originalname.endsWith(".xls")
    ) {
      cb(null, true);
    } else {
      cb(new Error("Only Excel files are accepted (.xlsx, .xls)"));
    }
  },
});

// Simulate staged thinking process
function buildThinkingSteps(parsedData) {
  const summary = parsedData?.summary || {};
  const sheetCount = parsedData?.raw?.summary ? 4 : 0;
  const propsCount = parsedData?.properties?.length || 0;
  const leasesCount = parsedData?.leases?.length || 0;
  const txCount = parsedData?.transactions?.length || 0;
  const occ = summary.occupancyPct || "—";
  const abr = summary.totalABR != null ? `$${(Number(summary.totalABR) / 1e6).toFixed(1)}M` : "—";
  const walt = summary.waltYears != null ? `${Number(summary.waltYears).toFixed(2)} yrs` : "—";
  return [
    { step: 1, label: "File uploaded successfully", delay: 300 },
    { step: 2, label: `Parsing Excel — detected ${sheetCount} sheets`, delay: 600 },
    { step: 3, label: `Sheet names: Summary, Properties, Leases, Transactions`, delay: 900 },
    { step: 4, label: `Properties sheet: ${propsCount} records loaded`, delay: 1200 },
    { step: 5, label: `Leases sheet: ${leasesCount} lease records parsed`, delay: 1500 },
    { step: 6, label: `Transactions sheet: ${txCount} records loaded`, delay: 1800 },
    { step: 7, label: "Calculating portfolio KPIs from lease roll", delay: 2100 },
    { step: 8, label: `KPIs computed: Occupancy ${occ} | ABR ${abr} | WALT ${walt}`, delay: 2400 },
    { step: 9, label: "Mapping KPIs to presentation template (8 slides)", delay: 2700 },
    { step: 10, label: "Generating report slides from Orion template", delay: 3000 },
    { step: 11, label: "Slides rendered and ready for display", delay: 3300 },
  ];
}

router.post("/", upload.single("excel"), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: "No file uploaded or invalid file type." });
  }

  try {
    // Parse the uploaded Excel
    let parsedData = null;
    try {
      parsedData = parsePortfolioExcel(req.file.path);
    } catch (parseErr) {
      console.warn("Excel parse error:", parseErr.message);
      return res.status(400).json({
        success: false,
        error:
          parseErr.message ||
          "Failed to parse Excel. Please ensure you are using the Orion_Q4_2025_Raw_Data.xlsx template.",
      });
    }

    // Set context for chat engine (full Excel knowledge for Q&A)
    if (parsedData) {
      setPortfolioContext(parsedData);
      setSlidesFromPortfolioData(parsedData);
    } else {
      resetSlides();
    }

    const thinking = buildThinkingSteps(parsedData);
    const fileId = path.basename(req.file.filename).split("-")[0];

    return res.json({
      success: true,
      fileId,
      filename: req.file.originalname,
      thinking,
      message: "Excel file parsed successfully. Generating Q4 2025 portfolio report...",
      parsedSummary: parsedData?.summary || null,
    });
  } catch (err) {
    console.error("Upload error:", err);
    return res.status(500).json({ error: "Failed to process uploaded file." });
  }
});

module.exports = router;
