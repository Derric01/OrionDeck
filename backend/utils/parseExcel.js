const xlsx = require("xlsx");

function parsePortfolioExcel(filePath) {
  const wb = xlsx.readFile(filePath);

  const result = {
    summary: {},
    properties: [],
    leases: [],
    transactions: [],
    raw: {},
  };

  // Parse Summary sheet
  if (wb.SheetNames.includes("Summary")) {
    const ws = wb.Sheets["Summary"];
    result.raw.summary = xlsx.utils.sheet_to_json(ws, { header: 1, defval: "" });
  }

  // Parse Properties sheet
  if (wb.SheetNames.includes("Properties")) {
    const ws = wb.Sheets["Properties"];
    const rows = xlsx.utils.sheet_to_json(ws, { header: 1, defval: "" });
    // Header row is row index 3 (0-indexed)
    const headers = rows[3];
    const requiredProps = [
      "Property ID",
      "Property Name",
      "Address",
      "City",
      "State",
      "Asset Type",
      "Status",
      "Year Built",
      "Year Renovated",
      "Rentable SF",
      "# Buildings",
      "# Floors",
      "Parking Spaces",
      "Construction Type",
      "Zoning",
      "Acquisition Date",
      "Acquisition Price",
      "Book Value",
    ];
    if (!headers || !requiredProps.every((h) => headers.includes(h))) {
      throw new Error(
        "Invalid Properties sheet structure. Please use the Orion_Q4_2025_Raw_Data.xlsx template (missing or renamed columns in 'Properties')."
      );
    }
    result.properties = rows.slice(4).filter((r) => r[0]).map((row) => {
      const obj = {};
      headers.forEach((h, i) => {
        if (h) obj[h] = row[i];
      });
      return obj;
    });

    // Walkthrough alignment: "Operating Properties" counts only rows where
    // Status = Operating. This prevents off-template rows from inflating KPIs.
    result.properties = result.properties.filter((p) => {
      const status = String(p?.["Status"] ?? "").trim().toLowerCase();
      if (!status) return false;
      return status === "operating" || status.includes("operating");
    });
  }

  // Parse Leases sheet
  if (wb.SheetNames.includes("Leases")) {
    const ws = wb.Sheets["Leases"];
    const rows = xlsx.utils.sheet_to_json(ws, { header: 1, defval: "" });
    const headers = rows[3];
    const requiredLeases = [
      "Lease ID",
      "Property ID",
      "Property Name",
      "City",
      "State",
      "Asset Type",
      "Tenant Name",
      "Credit Rating",
      "Inv. Grade?",
      "Leased SF",
      "Annual Base Rent",
      "Rent / SF",
      "Lease Type",
      "Commencement",
      "Expiration",
      "Remaining Term (yrs)",
      "Annual Escalation",
      "Renewal Options",
    ];
    if (!headers || !requiredLeases.every((h) => headers.includes(h))) {
      throw new Error(
        "Invalid Leases sheet structure. Please use the Orion_Q4_2025_Raw_Data.xlsx template (missing or renamed columns in 'Leases')."
      );
    }
    result.leases = rows.slice(4).filter((r) => r[0]).map((row) => {
      const obj = {};
      headers.forEach((h, i) => {
        if (h) obj[h] = row[i];
      });
      return obj;
    });
  }

  // Parse Transactions sheet
  if (wb.SheetNames.includes("Transactions")) {
    const ws = wb.Sheets["Transactions"];
    const rows = xlsx.utils.sheet_to_json(ws, { header: 1, defval: "" });
    const headers = rows[3];
    const requiredTx = [
      "Transaction ID",
      "Property Name",
      "City",
      "State",
      "Asset Type",
      "Type",
      "Transaction Date",
      "Quarter",
      "Rentable SF",
      "Gross Price",
      "Net Proceeds",
      "Book Value",
      "Gain / (Loss)",
      "Occ % at Sale",
      "Notes",
    ];
    if (!headers || !requiredTx.every((h) => headers.includes(h))) {
      throw new Error(
        "Invalid Transactions sheet structure. Please use the Orion_Q4_2025_Raw_Data.xlsx template (missing or renamed columns in 'Transactions')."
      );
    }
    result.transactions = rows
      .slice(4)
      .filter((r) => r[0] && r[0] !== "Q4 2025 TOTAL — 3 Dispositions")
      .map((row) => {
        const obj = {};
        headers.forEach((h, i) => {
          if (h) obj[h] = row[i];
        });
        return obj;
      });
  }

  // Derive key metrics from parsed data
  result.summary = deriveMetrics(result);

  return result;
}

function deriveMetrics(data) {
  const leases = data.leases;
  const properties = data.properties;
  const transactions = data.transactions;

  const totalABR = leases.reduce((sum, l) => {
    const val = parseFloat(String(l["Annual Base Rent"] || "0").replace(/[$,]/g, ""));
    return sum + (isNaN(val) ? 0 : val);
  }, 0);

  const leasedSF = leases.reduce((sum, l) => {
    const val = parseFloat(String(l["Leased SF"] || "0").replace(/[,]/g, ""));
    return sum + (isNaN(val) ? 0 : val);
  }, 0);

  const totalSF = properties.reduce((sum, p) => {
    const val = parseFloat(String(p["Rentable SF"] || "0").replace(/[,]/g, ""));
    return sum + (isNaN(val) ? 0 : val);
  }, 0);

  const igLeases = leases.filter(l => l["Inv. Grade?"] === "Yes");
  const igABR = igLeases.reduce((sum, l) => {
    const val = parseFloat(String(l["Annual Base Rent"] || "0").replace(/[$,]/g, ""));
    return sum + (isNaN(val) ? 0 : val);
  }, 0);

  // WALT: weighted average lease term (years). Use "Remaining Term (Yrs)" or "Expiry Date"
  let waltYears = null;
  const abrCol = "Annual Base Rent";
  const termCol = "Remaining Term (yrs)";
  const expiryCol = "Expiration";
  if (leases.length && totalABR > 0) {
    let weightedSum = 0;
    for (const l of leases) {
      const abr = parseFloat(String(l[abrCol] || "0").replace(/[$,]/g, "")) || 0;
      let years = 0;
      if (termCol && l[termCol] != null) years = parseFloat(String(l[termCol]).replace(/[^0-9.]/g, "")) || 0;
      else if (expiryCol && l[expiryCol]) {
        const d = new Date(l[expiryCol]);
        if (!isNaN(d.getTime())) years = Math.max(0, (d.getTime() - Date.now()) / (365.25 * 24 * 60 * 60 * 1000));
      }
      weightedSum += years * abr;
    }
    waltYears = weightedSum / totalABR;
  }

  return {
    operatingProperties: properties.length,
    totalRentableSF: totalSF,
    occupiedSF: leasedSF,
    occupancyPct: totalSF > 0 ? ((leasedSF / totalSF) * 100).toFixed(1) + "%" : "N/A",
    totalABR: totalABR,
    abrPerSF: leasedSF > 0 ? (totalABR / leasedSF).toFixed(2) : "N/A",
    waltYears: waltYears != null ? waltYears.toFixed(2) : null,
    igABR: igABR,
    igPct: totalABR > 0 ? ((igABR / totalABR) * 100).toFixed(1) + "%" : "N/A",
    totalLeases: leases.length,
    totalTransactions: transactions.length,
  };
}

module.exports = { parsePortfolioExcel };
