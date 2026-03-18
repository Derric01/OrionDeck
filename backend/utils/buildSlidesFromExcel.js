// Builds the 8 Orion template slides from parsed Excel.
// Output structure matches slideContent.js baseSlides (cover, portfolioHighlights, composition, assetPerformance, expiry, dispositions, outlook, appendix).

function fmtNum(n) {
  if (n == null || isNaN(n)) return "—";
  return Math.round(n).toLocaleString();
}
function fmtMoney(n) {
  if (n == null || isNaN(n)) return "—";
  const val = Math.round(n);
  if (val >= 1e6) return "$" + (val / 1e6).toFixed(1) + "M";
  return "$" + val.toLocaleString();
}
function fmtSF(n) {
  if (n == null || isNaN(n)) return "—";
  const val = Math.round(n);
  if (val >= 1e6) return (val / 1e6).toFixed(2) + "M SF";
  if (val >= 1e3) return (val / 1e3).toFixed(0) + "K SF";
  return val.toLocaleString() + " SF";
}

function excelDateToJsDate(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  if (typeof value === "number" && isFinite(value)) {
    // Excel's day 0 is 1899-12-30 in JS date serial terms
    const ms = Math.round((value - 25569) * 86400 * 1000);
    const d = new Date(ms);
    return isNaN(d.getTime()) ? null : d;
  }
  if (typeof value === "string") {
    const s = value.trim();
    // Handle common non-ISO formats found in uploaded workbooks, e.g. 31/03/2035 (DD/MM/YYYY)
    const m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) {
      const a = parseInt(m[1], 10);
      const b = parseInt(m[2], 10);
      const y = parseInt(m[3], 10);
      // Disambiguate: if day > 12, treat as DD/MM; if month > 12, treat as MM/DD; else default to DD/MM.
      const day = a > 12 ? a : b > 12 ? b : a;
      const month = a > 12 ? b : b > 12 ? a : b;
      const d2 = new Date(Date.UTC(y, month - 1, day));
      return isNaN(d2.getTime()) ? null : new Date(y, month - 1, day);
    }
  }
  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

function monthYear(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return "—";
  const months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  return `${months[d.getMonth()]} ${d.getFullYear()}`;
}

function safePctString(v) {
  if (v == null) return "—";
  const s = String(v).trim();
  if (!s) return "—";
  if (/%$/.test(s)) return s;
  const n = parseNumber(s);
  if (!isFinite(n)) return "—";
  return n.toFixed(1) + "%";
}

function findKey(obj, regexes) {
  if (!obj) return null;
  const keys = Object.keys(obj);
  for (const r of regexes) {
    const k = keys.find((x) => r.test(String(x)));
    if (k) return k;
  }
  return null;
}

function parseNumber(v) {
  if (v == null) return 0;
  if (typeof v === "number" && isFinite(v)) return v;
  const s = String(v).trim();
  if (!s) return 0;
  const cleaned = s.replace(/[$,%\s]/g, "").replace(/,/g, "");
  const n = parseFloat(cleaned);
  return isNaN(n) ? 0 : n;
}

function groupSum(rows, groupKey, valueKey) {
  const out = new Map();
  for (const r of rows || []) {
    const k = (r?.[groupKey] ?? "").toString().trim() || "—";
    const v = parseNumber(r?.[valueKey]);
    out.set(k, (out.get(k) || 0) + v);
  }
  return out;
}

function buildSlidesFromParsedData(parsedData, baseSlides) {
  if (!baseSlides || baseSlides.length < 8) throw new Error("baseSlides (8 slides) required");
  const summary = parsedData?.summary || {};
  const properties = parsedData?.properties || [];
  const leases = parsedData?.leases || [];
  const transactions = parsedData?.transactions || [];

  const totalABR = Number(summary.totalABR) || 0;
  const totalSF = Number(summary.totalRentableSF) || 0;
  const occupiedSF = Number(summary.occupiedSF) || 0;
  const occPct = summary.occupancyPct || (totalSF > 0 ? ((occupiedSF / totalSF) * 100).toFixed(1) + "%" : "—");
  const walt = summary.waltYears != null ? String(summary.waltYears) : (totalABR > 0 ? "—" : "—");
  const numProps = summary.operatingProperties ?? properties.length ?? 0;
  const igPct = summary.igPct || "—";
  const abrPerSF = summary.abrPerSF != null ? String(summary.abrPerSF) : "—";

  const abrCol = "Annual Base Rent";
  const sfCol = findKey(leases[0], [/leased\s*sf/i, /\bsf\b/i, /square/i]) || "Leased SF";
  const assetTypeKey = findKey(properties[0], [/asset\s*type/i, /\btype\b/i]) || "Asset Type";
  const tenantCol = findKey(leases[0], [/tenant/i, /lessee/i, /customer/i, /company/i]) || "Tenant Name";
  const ratingCol = findKey(leases[0], [/credit\s*rating/i, /\brating\b/i, /grade/i, /credit/i]);
  const leasePropertyCol = "Property ID";
  const propertyNameCol = findKey(properties[0], [/property/i, /asset/i, /building/i, /site/i, /name/i]);
  const propertyStateCol = findKey(properties[0], [/^\s*state\s*$/i, /state/i, /\bst\b/i]);
  const leaseIndustryCol = findKey(leases[0], [/industry/i, /sector/i]);
  const invGradeCol = findKey(leases[0], [/inv\.\s*grade/i, /investment\s*grade/i, /inv\s*grade/i]);
  const expiryCol = "Expiration";
  const remainingTermCol = "Remaining Term (yrs)";

  // ─── Slide 1: Cover (KPI cards sourced from uploaded Excel summary) ─────────
  const coverSlide = {
    ...baseSlides[0],
    content: {
      ...baseSlides[0].content,
      kpis: [
        { cardLabel: "PROPERTIES", value: fmtNum(numProps), subLabel: "Operating" },
        { cardLabel: "OCCUPANCY", value: occPct || "—", subLabel: "% rentable SF" },
        { cardLabel: "WALT", value: walt && walt !== "—" ? `${parseNumber(walt).toFixed(2)} yrs` : "—", subLabel: "wtd. avg. lease term" },
        { cardLabel: "ANNUALISED BASE RENT", value: fmtMoney(totalABR), subLabel: "ABR, uploaded workbook" },
        { cardLabel: "ABR / SF", value: abrPerSF !== "—" && abrPerSF !== "N/A" ? `$${parseNumber(abrPerSF).toFixed(2)}` : "—", subLabel: "active leases" },
        { cardLabel: "INV-GRADE TENANCY", value: igPct || "—", subLabel: "% of ABR" },
      ],
      // Keep highlights text from template unless the workbook carries quarter flags.
      highlightsSectionTitle: baseSlides[0].content.highlightsSectionTitle,
      highlights: baseSlides[0].content.highlights,
    },
  };

  // Top 10 tenants by ABR for Slide 2
  const topLeases = [...leases]
    .map(l => ({
      name: (tenantCol && l[tenantCol]) ? String(l[tenantCol]) : "—",
      creditRating: (ratingCol && l[ratingCol]) ? String(l[ratingCol]) : "—",
      abr: parseFloat(String(l[abrCol] || "0").replace(/[$,]/g, "")) || 0,
    }))
    .filter(l => l.abr > 0)
    .sort((a, b) => b.abr - a.abr)
    .slice(0, 10);
  const top10WithPct = topLeases.map((t, i) => ({
    rank: i + 1,
    name: t.name,
    creditRating: t.creditRating,
    pctABR: totalABR > 0 ? ((t.abr / totalABR) * 100).toFixed(1) + "%" : "—",
  }));
  const combinedTop10Pct = topLeases.length && totalABR > 0
    ? ((topLeases.reduce((s, t) => s + t.abr, 0) / totalABR) * 100).toFixed(1) + "%"
    : "60.5%";

  // ─── Slide 2: Portfolio Highlights (data-driven from uploaded Excel) ─────
  const highlightsSlide = {
    ...baseSlides[1],
    content: {
      metrics: {
        operatingProperties: fmtNum(numProps),
        totalLeasableSF: fmtSF(totalSF),
        abr: fmtMoney(totalABR),
        occupancy: occPct || "—",
        walt: walt && walt !== "—" ? `${parseNumber(walt).toFixed(2)} years` : "—",
        igTenancy: igPct && igPct !== "N/A" ? `${igPct} of ABR` : "—",
      },
      top10Tenants: top10WithPct,
      combinedNote:
        top10WithPct.length
          ? `Combined Top-10 tenants contribute ${combinedTop10Pct} of ABR and ${igPct || "—"} are investment grade.`
          : baseSlides[1].content.combinedNote,
      sourceTables: [
        {
          title: "Leases (Uploaded Excel — full table)",
          sheet: "Leases",
          rows: leases,
        },
      ],
    },
  };

  // Asset type aggregation for Slide 3
  const propByType = {};
  for (const p of properties) {
    const type = (p?.[assetTypeKey] ?? "Other").toString().trim() || "Other";
    if (!propByType[type]) propByType[type] = { count: 0, sf: 0, abr: 0 };
    propByType[type].count += 1;
    propByType[type].sf += parseNumber(p?.["Rentable SF"]);
  }
  for (const l of leases) {
    // Prefer explicit asset type on lease rows; if missing, bucket as Other.
    const leaseAssetTypeKey = findKey(l, [/asset\s*type/i, /\btype\b/i]);
    let type = leaseAssetTypeKey ? String(l[leaseAssetTypeKey] ?? "").trim() : "";
    if (!type) type = "Other";
    if (!propByType[type]) propByType[type] = { count: 0, sf: 0, abr: 0 };
    propByType[type].abr += parseNumber(l?.[abrCol]);
  }
  const assetTypes = Object.entries(propByType)
    .map(([type, v]) => ({ type, count: v.count, sf: v.sf, abr: v.abr }))
    .filter((a) => a.count > 0 || a.sf > 0 || a.abr > 0)
    .sort((a, b) => b.abr - a.abr);
  const totalABRAsset = assetTypes.reduce((s, a) => s + a.abr, 0) || totalABR;

  const compositionRows = assetTypes
    .filter((a) => a.abr > 0)
    .slice(0, 12)
    .map((a) => [
      a.type,
      fmtNum(a.count),
      fmtSF(a.sf),
      totalABRAsset > 0 ? ((a.abr / totalABRAsset) * 100).toFixed(1) + "%" : "—",
    ]);

  const compositionTotals = [
    "Total",
    fmtNum(numProps),
    fmtSF(totalSF),
    "100%",
  ];

  const igSplit = {
    investmentGrade: igPct && igPct !== "N/A" ? igPct : "—",
    nonInvestmentGrade:
      igPct && igPct !== "N/A"
        ? (100 - parseNumber(igPct)).toFixed(1) + "%"
        : "—",
  };

  const industryBreakdown = (() => {
    if (!leaseIndustryCol) return [];
    const byIndustry = groupSum(leases, leaseIndustryCol, abrCol);
    const entries = [...byIndustry.entries()].filter(([, v]) => v > 0).sort((a, b) => b[1] - a[1]);
    const top = entries.slice(0, 10).map(([name, v]) => ({
      name,
      pct: totalABR > 0 ? ((v / totalABR) * 100).toFixed(1) + "%" : "—",
    }));
    const remainder = entries.slice(10).reduce((s, [, v]) => s + v, 0);
    if (remainder > 0) {
      top.push({ name: "Other", pct: totalABR > 0 ? ((remainder / totalABR) * 100).toFixed(1) + "%" : "—" });
    }
    return top;
  })();

  const geographicBreakdown = (() => {
    if (!propertyStateCol || !propertyNameCol || !leasePropertyCol) return [];
    const propState = new Map(
      properties.map((p) => [String(p?.[propertyNameCol] ?? "").trim(), String(p?.[propertyStateCol] ?? "").trim()])
    );
    // Sum ABR by state via lease.property -> properties.state join
    const abrByState = new Map();
    for (const l of leases) {
      const prop = String(l?.[leasePropertyCol] ?? "").trim();
      const state = prop ? (propState.get(prop) || "") : "";
      if (!state) continue;
      abrByState.set(state, (abrByState.get(state) || 0) + parseNumber(l?.[abrCol]));
    }
    const entries = [...abrByState.entries()].filter(([, v]) => v > 0).sort((a, b) => b[1] - a[1]);
    const top = entries.slice(0, 10).map(([state, v]) => ({
      state,
      pct: totalABR > 0 ? ((v / totalABR) * 100).toFixed(1) + "%" : "—",
    }));
    const remainder = entries.slice(10).reduce((s, [, v]) => s + v, 0);
    if (remainder > 0) {
      top.push({
        state: `Other states (${Math.max(entries.length - 10, 0)} total)`,
        pct: totalABR > 0 ? ((remainder / totalABR) * 100).toFixed(1) + "%" : "—",
      });
    }
    return top;
  })();

  // ─── Slide 3: Composition (data-driven where workbook supports it) ──────
  const compositionSlide = {
    ...baseSlides[2],
    content: {
      ...baseSlides[2].content,
      igSplit,
      assetTable: {
        headers: ["Asset Type", "Properties", "Square Feet", "% of ABR"],
        rows: compositionRows.length ? compositionRows : baseSlides[2].content.assetTable.rows,
        totals: compositionRows.length ? compositionTotals : baseSlides[2].content.assetTable.totals,
      },
      industryBreakdown: industryBreakdown.length ? industryBreakdown : baseSlides[2].content.industryBreakdown,
      geographicBreakdown: geographicBreakdown.length ? geographicBreakdown : baseSlides[2].content.geographicBreakdown,
    },
  };

  // ─── Slide 4: Asset-by-Asset Performance (top 10 properties by ABR) ──────
  const assetPerfRows = (() => {
    if (!leasePropertyCol) return [];
    const abrByProp = groupSum(leases, leasePropertyCol, abrCol);
    const leasedSfByProp = sfCol ? groupSum(leases, leasePropertyCol, sfCol) : new Map();

    const propMeta = (() => {
      if (!propertyNameCol) return new Map();
      const m = new Map();
      for (const p of properties) {
        const id = String(p?.["Property ID"] ?? "").trim();
        if (!id) continue;
        m.set(id, {
          name: String(p?.["Property Name"] ?? "").trim() || id,
          type: String(p?.[assetTypeKey] ?? "").trim() || "—",
          rentableSF: parseNumber(p?.["Rentable SF"]),
        });
      }
      return m;
    })();

    const entries = [...abrByProp.entries()].filter(([, v]) => v > 0).sort((a, b) => b[1] - a[1]).slice(0, 10);

    const waltByProp = (() => {
      const out = new Map();
      if (!leases.length) return out;
      for (const l of leases) {
        const prop = String(l?.[leasePropertyCol] ?? "").trim();
        if (!prop) continue;
        const abr = parseNumber(l?.[abrCol]);
        if (!abr) continue;
        let years = 0;
        if (remainingTermCol && l?.[remainingTermCol] != null) years = parseNumber(l?.[remainingTermCol]);
        else if (expiryCol && l?.[expiryCol]) {
          const d = excelDateToJsDate(l?.[expiryCol]);
          if (d) years = Math.max(0, (d.getTime() - Date.now()) / (365.25 * 24 * 60 * 60 * 1000));
        }
        const prev = out.get(prop) || { weighted: 0, abr: 0 };
        out.set(prop, { weighted: prev.weighted + years * abr, abr: prev.abr + abr });
      }
      return out;
    })();

    const expiryByProp = (() => {
      const out = new Map();
      if (!expiryCol) return out;
      for (const l of leases) {
        const prop = String(l?.[leasePropertyCol] ?? "").trim();
        if (!prop) continue;
        const d = excelDateToJsDate(l?.[expiryCol]);
        if (!d) continue;
        const prev = out.get(prop);
        if (!prev) out.set(prop, d);
        else if (d > prev) out.set(prop, d);
      }
      return out;
    })();

    const creditByProp = (() => {
      const out = new Map();
      for (const l of leases) {
        const prop = String(l?.[leasePropertyCol] ?? "").trim();
        if (!prop) continue;
        const ig = invGradeCol ? String(l?.[invGradeCol] ?? "").trim().toLowerCase() : "";
        if (ig === "yes" || ig === "y" || ig === "true") out.set(prop, "IG");
        else if (ratingCol && l?.[ratingCol]) out.set(prop, String(l[ratingCol]));
      }
      return out;
    })();

    return entries.map(([prop, abr]) => {
      const leased = leasedSfByProp.get(prop) || 0;
      const meta = propMeta.get(prop) || { name: prop, type: "—", rentableSF: 0 };
      const occ = meta.rentableSF > 0 ? ((leased / meta.rentableSF) * 100).toFixed(0) + "%" : "—";
      const w = waltByProp.get(prop);
      const waltProp = w && w.abr > 0 ? (w.weighted / w.abr).toFixed(1) + "y" : "—";
      const exp = expiryByProp.get(prop);
      return {
        property: meta.name || prop,
        type: meta.type || "—",
        sf: meta.rentableSF ? fmtNum(meta.rentableSF) : (leased ? fmtNum(leased) : "—"),
        occupancy: occ,
        abr: fmtMoney(abr),
        walt: waltProp,
        leaseExpiry: exp ? monthYear(exp) : "—",
        credit: creditByProp.get(prop) || "—",
      };
    });
  })();

  const assetPerfSlide = {
    ...baseSlides[3],
    content: {
      headers: baseSlides[3].content.headers,
      rows: assetPerfRows.length ? assetPerfRows : baseSlides[3].content.rows,
      sourceTables: [
        {
          title: "Properties (Uploaded Excel — full table)",
          sheet: "Properties",
          rows: properties,
        },
      ],
    },
  };

  // ─── Slide 5: Lease Expiry Schedule (computed from lease expiry dates) ───
  const expirySlide = (() => {
    if (!expiryCol) return { ...baseSlides[4] };
    const buckets = new Map(); // year -> { leases, sf, abr }
    for (const l of leases) {
      const d = excelDateToJsDate(l?.[expiryCol]);
      if (!d) continue;
      const year = d.getFullYear();
      const bucket = year >= 2031 ? "2031+" : String(year);
      const prev = buckets.get(bucket) || { leases: 0, sf: 0, abr: 0 };
      prev.leases += 1;
      prev.sf += parseNumber(l?.[sfCol]);
      prev.abr += parseNumber(l?.[abrCol]);
      buckets.set(bucket, prev);
    }
    const years = [];
    for (let y = new Date().getFullYear(); y <= 2030; y += 1) years.push(String(y));
    // Ensure we always show 2026-2030 + 2031+ if available; otherwise use what's present.
    const ordered = ["2026", "2027", "2028", "2029", "2030", "2031+"].filter((k) => buckets.has(k) || k === "2031+");
    const schedule = ordered
      .map((k) => ({ year: k, ...(buckets.get(k) || { leases: 0, sf: 0, abr: 0 }) }))
      .filter((r) => r.leases > 0 || r.sf > 0 || r.abr > 0);

    const totalAbrExp = schedule.reduce((s, r) => s + r.abr, 0);
    const totalSfExp = schedule.reduce((s, r) => s + r.sf, 0);
    const totalLeasesExp = schedule.reduce((s, r) => s + r.leases, 0);

    const mapped = schedule.map((r) => ({
      year: r.year,
      leases: r.leases,
      sfExpiring: fmtSF(r.sf),
      abrExpiring: fmtMoney(r.abr),
      pctPortfolio: totalABR > 0 ? ((r.abr / totalABR) * 100).toFixed(1) + "%" : "—",
    }));

    const longTerm = schedule.find((r) => r.year === "2031+");
    const keyNote =
      longTerm && totalABR > 0
        ? `${((longTerm.abr / totalABR) * 100).toFixed(1)}% of ABR expires in 2031 or later, indicating limited near-term rollover risk.`
        : baseSlides[4].content.keyNote;

    return {
      ...baseSlides[4],
      content: {
        headers: baseSlides[4].content.headers,
        schedule: mapped.length ? mapped : baseSlides[4].content.schedule,
        totalRow: {
          year: "Total",
          leases: totalLeasesExp || baseSlides[4].content.totalRow.leases,
          sfExpiring: totalSfExp > 0 ? fmtSF(totalSfExp) : baseSlides[4].content.totalRow.sfExpiring,
          abrExpiring: totalAbrExp > 0 ? fmtMoney(totalAbrExp) : baseSlides[4].content.totalRow.abrExpiring,
          pctPortfolio: "100%",
        },
        keyNote,
        sourceTables: [
          {
            title: "Lease expiry inputs (from Leases sheet)",
            sheet: "Leases",
            rows: leases.map((l) => {
              const out = {};
              if (tenantCol) out["Tenant"] = l[tenantCol];
              if (leasePropertyCol) out["Property"] = l[leasePropertyCol];
              out["Annual Base Rent"] = l[abrCol];
              out["Leased SF"] = l[sfCol];
              if (expiryCol) out["Expiry"] = l[expiryCol];
              if (remainingTermCol) out["Remaining Term"] = l[remainingTermCol];
              return out;
            }),
          },
        ],
      },
    };
  })();

  // ─── Slide 6: Dispositions (from Transactions sheet when present) ───────
  const dispositionsSlide = (() => {
    if (!transactions.length) return { ...baseSlides[5] };
    const priceCol = "Gross Price";
    const sfTxCol = "Rentable SF";
    const occTxCol = "Occ % at Sale";
    const assetTypeTxCol = findKey(transactions[0], [/asset\s*type/i, /\btype\b/i]);
    const regionCol = findKey(transactions[0], [/region/i, /market/i]);
    const reasonCol = findKey(transactions[0], [/reason/i, /strateg/i, /rationale/i, /note/i]);
    const propCol = findKey(transactions[0], [/property/i, /asset/i, /building/i, /site/i, /name/i]) || Object.keys(transactions[0] || {})[0];

    const grossProceeds = priceCol ? transactions.reduce((s, t) => s + parseNumber(t?.[priceCol]), 0) : 0;

    const mappedTx = transactions.slice(0, 12).map((t) => ({
      property: String(t?.[propCol] ?? "—"),
      region: regionCol ? String(t?.[regionCol] ?? "—") : "—",
      assetType: assetTypeTxCol ? String(t?.[assetTypeTxCol] ?? "—") : "—",
      sf: sfTxCol ? fmtNum(parseNumber(t?.[sfTxCol])) : "—",
      price: priceCol ? fmtMoney(parseNumber(t?.[priceCol])) : "—",
      occupancy: occTxCol ? safePctString(t?.[occTxCol]) : "—",
      strategicReason: reasonCol ? String(t?.[reasonCol] ?? "—") : "—",
    }));

    return {
      ...baseSlides[5],
      content: {
        metrics: {
          q4PropertiesSold: fmtNum(transactions.length),
          q4GrossProceeds: grossProceeds ? fmtMoney(grossProceeds) : baseSlides[5].content.metrics.q4GrossProceeds,
          fy2025Sold: baseSlides[5].content.metrics.fy2025Sold,
          fy2025Proceeds: baseSlides[5].content.metrics.fy2025Proceeds,
          carryingCostSavings: baseSlides[5].content.metrics.carryingCostSavings,
          capexAvoided: baseSlides[5].content.metrics.capexAvoided,
        },
        transactions: mappedTx.length ? mappedTx : baseSlides[5].content.transactions,
        fullYearNote: baseSlides[5].content.fullYearNote,
        sourceTables: [
          {
            title: "Transactions (Uploaded Excel — full table)",
            sheet: "Transactions",
            rows: transactions,
          },
        ],
      },
    };
  })();

  // ─── Slide 7 & 8: Outlook and Appendix (use template) ───────────────────
  const outlookSlide = { ...baseSlides[6] };
  const appendixSlide = {
    ...baseSlides[7],
    content: {
      ...baseSlides[7].content,
      sourceTables: [
        {
          title: "Summary sheet (raw grid)",
          sheet: "Summary",
          rows: Array.isArray(parsedData?.raw?.summary)
            ? parsedData.raw.summary.map((r, idx) => ({ row: idx + 1, values: r }))
            : [],
        },
      ],
    },
  };

  return [
    coverSlide,
    highlightsSlide,
    compositionSlide,
    assetPerfSlide,
    expirySlide,
    dispositionsSlide,
    outlookSlide,
    appendixSlide,
  ];
}

module.exports = { buildSlidesFromParsedData };
