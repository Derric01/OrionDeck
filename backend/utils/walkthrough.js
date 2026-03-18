const path = require("path");
const fs = require("fs");
const xlsx = require("xlsx");

let walkthroughContext = null;

function loadWalkthroughWorkbook(filePath) {
  const abs = path.isAbsolute(filePath) ? filePath : path.join(process.cwd(), filePath);
  if (!fs.existsSync(abs)) {
    walkthroughContext = null;
    return null;
  }

  const wb = xlsx.readFile(abs);
  const sheets = {};
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    // Keep as a 2D grid for fidelity; cap columns by worksheet itself.
    const grid = xlsx.utils.sheet_to_json(ws, { header: 1, defval: "" });
    sheets[name] = grid;
  }

  walkthroughContext = {
    file: path.basename(abs),
    sheetNames: wb.SheetNames,
    sheets,
  };
  return walkthroughContext;
}

function setWalkthroughContext(ctx) {
  walkthroughContext = ctx;
}

function getWalkthroughContext() {
  return walkthroughContext;
}

module.exports = { loadWalkthroughWorkbook, setWalkthroughContext, getWalkthroughContext };

