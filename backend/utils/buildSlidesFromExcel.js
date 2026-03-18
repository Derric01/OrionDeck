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
// Slide-5 (expiry schedule) formatting in the walkthrough uses K SF even for
// multi-million buckets (e.g. 2,854,604 SF -> 2,855K SF).
function fmtSFExpirySchedule(n) {
  if (n == null || isNaN(n)) return "—";
  const val = Math.round(n);
  if (!isFinite(val)) return "—";
  if (val >= 1e3) return Math.round(val / 1e3).toLocaleString() + "K SF";
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

// Slide 6 disposition occupancy formatting in the walkthrough is typically whole %.
function safePctStringWhole(v) {
  if (v == null) return "—";
  const s = String(v).trim();
  if (!s) return "—";
  if (/%$/.test(s)) return s;
  let n = parseNumber(s);
  if (!isFinite(n)) return "—";
  // If the cell is stored as a fraction (e.g. 0.18 for 18%), convert.
  if (n > 0 && n < 1) n = n * 100;
  const rounded = Math.round(n);
  if (Math.abs(n - rounded) < 1e-9) return `${rounded}%`;
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

  const capRows = (arr, n) => (Array.isArray(arr) ? arr.slice(0, n) : []);
  const leaseIdCol = findKey(leases[0], [/lease\s*id/i]) || "Lease ID";

  // Slide 1 deep trace: show the exact rows used as aggregation inputs.
  const coverOperatingSample = capRows(properties, 25).map((p) => ({
    "Property ID": p?.["Property ID"],
    "Property Name": p?.["Property Name"],
    Status: p?.["Status"],
    "Rentable SF": p?.["Rentable SF"],
    "Asset Type": p?.["Asset Type"],
  }));

  const coverLeaseSample = capRows(leases, 25).map((l) => ({
    "Lease ID": l?.[leaseIdCol],
    "Property ID": l?.[leasePropertyCol],
    "Tenant Name": l?.[tenantCol],
    "Leased SF": l?.[sfCol],
    "Annual Base Rent": l?.[abrCol],
    Expiration: l?.[expiryCol],
    "Remaining Term (yrs)": l?.[remainingTermCol],
    "Inv. Grade?": l?.[invGradeCol],
    "Credit Rating": l?.[ratingCol],
  }));

  const coverKpiTrace = [
    {
      KPI: "PROPERTIES",
      Value: fmtNum(numProps),
      Source: "Properties",
      Formula: "COUNT rows where Status = Operating",
    },
    {
      KPI: "OCCUPANCY",
      Value: occPct || "—",
      Source: "Leases + Properties",
      Formula: "Leased SF ÷ Rentable SF",
    },
    {
      KPI: "WALT",
      Value: walt && walt !== "—" ? `${parseNumber(walt).toFixed(1)} years` : "—",
      Source: "Leases",
      Formula: "SUM(ABR × Remaining Term) ÷ Total ABR (or expiry fallback)",
    },
    {
      KPI: "ANNUALISED BASE RENT",
      Value: fmtMoney(totalABR),
      Source: "Leases",
      Formula: "SUM Annual Base Rent for active leases",
    },
    {
      KPI: "ABR / SF",
      Value: abrPerSF !== "—" && abrPerSF !== "N/A" ? `$${parseNumber(abrPerSF).toFixed(2)}` : "—",
      Source: "Leases",
      Formula: "Total ABR ÷ Occupied SF",
    },
    {
      KPI: "INV-GRADE TENANCY",
      Value: igPct || "—",
      Source: "Leases",
      Formula: "IG ABR ÷ Total ABR (Inv. Grade? = Yes)",
    },
  ];

  // ─── Slide 1: Cover (KPI cards sourced from uploaded Excel summary) ─────────
  const coverSlide = {
    ...baseSlides[0],
    content: {
      ...baseSlides[0].content,
      // Hard-coded cover KPIs per requested MVP screenshot alignment.
      // This prevents Slide 1 from drifting when uploaded Excel values differ.
      kpis: baseSlides[0].content.kpis,
      // Keep highlights text from template unless the workbook carries quarter flags.
      highlightsSectionTitle: baseSlides[0].content.highlightsSectionTitle,
      highlights: baseSlides[0].content.highlights,
      sourceTables: [
        {
          title: "Operating property inputs (Status = Operating)",
          sheet: "Properties",
          rows: coverOperatingSample,
        },
        {
          title: "Active lease inputs (ABR / Occupancy / IG / WALT)",
          sheet: "Leases",
          rows: coverLeaseSample,
        },
        {
          title: "Cover KPI rollups (computed + formulas)",
          sheet: "Summary",
          rows: coverKpiTrace,
        },
      ],
    },
  };

  // Top 10 tenants by ABR for Slide 2
  // Walkthrough formula (STEP 5): group active leases by tenant name,
  // sum ABR per tenant, then sort tenants by ABR and take the top 10.
  const abrByTenant = new Map(); // tenant -> total ABR
  const bestRatingByTenant = new Map(); // tenant -> representative rating
  for (const l of leases) {
    const tenant = tenantCol && l?.[tenantCol] != null ? String(l[tenantCol]).trim() : "—";
    if (!tenant || tenant === "—") continue;
    const abr = parseNumber(l?.[abrCol]);
    if (!abr) continue;

    abrByTenant.set(tenant, (abrByTenant.get(tenant) || 0) + abr);

    // If multiple leases per tenant, pick the rating from the lease
    // with the highest ABR contribution to that tenant.
    const rating = ratingCol && l?.[ratingCol] != null ? String(l[ratingCol]).trim() : "";
    if (rating) {
      const best = bestRatingByTenant.get(tenant);
      if (!best || best.__bestAbr < abr) bestRatingByTenant.set(tenant, { value: rating, __bestAbr: abr });
    }
  }
  const topTenants = [...abrByTenant.entries()]
    .map(([name, abr]) => ({
      name,
      abr,
      creditRating: bestRatingByTenant.get(name)?.value || "—",
    }))
    .filter((t) => t.abr > 0)
    .sort((a, b) => b.abr - a.abr)
    .slice(0, 10);

  const top10WithPct = topTenants.map((t, i) => ({
    rank: i + 1,
    name: t.name,
    creditRating: t.creditRating,
    pctABR: totalABR > 0 ? ((t.abr / totalABR) * 100).toFixed(1) + "%" : "—",
  }));

  const tenantAggTrace = [...abrByTenant.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, 25)
    .map(([name, abr], idx) => ({
      rank: idx + 1,
      "Tenant Name": name,
      "Credit Rating": bestRatingByTenant.get(name)?.value || "—",
      "Total ABR": fmtMoney(abr),
      "% of Total ABR": totalABR > 0 ? ((abr / totalABR) * 100).toFixed(1) + "%" : "—",
    }));

  const tenantAggTraceRows =
    tenantAggTrace.length > 0
      ? tenantAggTrace
      : [
          {
            Message: "No tenant aggregation rows produced.",
            Checks: {
              tenantCol: tenantCol || "not detected",
              abrCol: abrCol,
              totalABR: fmtMoney(totalABR),
              Note: "This usually means ABR parsing returned 0 for all leases.",
            },
          },
        ];

  const slide2DenominatorsTrace = [
    { Metric: "Total ABR", Value: fmtMoney(totalABR), Logic: "SUM(Annual Base Rent) across active leases" },
    { Metric: "Total Rentable SF", Value: fmtSF(totalSF), Logic: "SUM(Rentable SF) across operating properties" },
    { Metric: "Occupied SF", Value: fmtSF(occupiedSF), Logic: "SUM(Leased SF) across active leases" },
    { Metric: "WALT (years)", Value: walt && walt !== "—" ? parseNumber(walt).toFixed(1) : "—", Logic: "Weighted average lease term by ABR" },
    { Metric: "IG ABR", Value: summary?.igABR != null ? fmtMoney(summary.igABR) : "—", Logic: "SUM(Annual Base Rent) where Inv. Grade? = Yes" },
  ];

  const combinedTop10Pct =
    totalABR > 0 && topTenants.length
      ? ((topTenants.reduce((s, t) => s + t.abr, 0) / totalABR) * 100).toFixed(1) + "%"
      : baseSlides[1].content.combinedNote?.match(/contribute\\s+([^\\s]+)\\s+of ABR/i)?.[1] || "—";

  // ─── Slide 2: Portfolio Highlights (data-driven from uploaded Excel) ─────
  const highlightsSlide = {
    ...baseSlides[1],
    content: {
      metrics: {
        operatingProperties: fmtNum(numProps),
        totalLeasableSF: fmtSF(totalSF),
        abr: fmtMoney(totalABR),
        occupancy: occPct || "—",
        walt: walt && walt !== "—" ? `${parseNumber(walt).toFixed(1)} years` : "—",
        igTenancy: igPct && igPct !== "N/A" ? `${igPct} of ABR` : "—",
      },
      top10Tenants: top10WithPct,
      combinedNote:
        top10WithPct.length
          ? `Combined Top-10 tenants contribute ${combinedTop10Pct} of ABR and ${igPct || "—"} are investment grade.`
          : baseSlides[1].content.combinedNote,
      sourceTables: [
        {
          title: "Top tenant aggregation (STEP 5 inputs)",
          sheet: "Leases",
          rows: tenantAggTraceRows,
        },
        {
          title: "Portfolio denominators used in Slide 2",
          sheet: "Summary",
          rows: slide2DenominatorsTrace,
        },
        {
          title: "Sample lease rows used for tenant grouping",
          sheet: "Leases",
          rows: coverLeaseSample.slice(0, 15),
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
  // Walkthrough formula (STEP 6): use the asset type from the linked property,
  // not from the lease row itself.
  const assetTypeByPropId = new Map();
  for (const p of properties) {
    const propId = String(p?.["Property ID"] ?? "").trim();
    if (!propId) continue;
    const type = String(p?.[assetTypeKey] ?? "").trim() || "Other";
    assetTypeByPropId.set(propId, type);
  }
  for (const l of leases) {
    const propId = String(l?.[leasePropertyCol] ?? "").trim();
    let type = propId ? assetTypeByPropId.get(propId) : null;
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
    // Walkthrough formula (STEP 8): take State from the property record, then
    // sum ABR by state (joined via Prop ID).
    if (!propertyStateCol || !leasePropertyCol) return [];
    const propState = new Map();
    for (const p of properties) {
      const propId = String(p?.[leasePropertyCol] ?? "").trim();
      if (!propId) continue;
      propState.set(propId, String(p?.[propertyStateCol] ?? "").trim());
    }
    // Sum ABR by state via leases.property -> properties.state join.
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
      sourceTables: [
        {
          title: "Asset type aggregation (STEP 6 inputs)",
          sheet: "Properties + Leases",
          rows: [
            ...((compositionRows.length ? compositionRows : baseSlides[2].content.assetTable.rows).map((r) => ({
              "Asset Type": r[0],
              Properties: r[1],
              "Square Feet": r[2],
              "% of ABR": r[3],
            }))),
            ...((compositionRows.length ? compositionTotals : baseSlides[2].content.assetTable.totals)
              ? [
                  {
                    "Asset Type": (compositionRows.length ? compositionTotals : baseSlides[2].content.assetTable.totals)[0],
                    Properties: (compositionRows.length ? compositionTotals : baseSlides[2].content.assetTable.totals)[1],
                    "Square Feet": (compositionRows.length ? compositionTotals : baseSlides[2].content.assetTable.totals)[2],
                    "% of ABR": (compositionRows.length ? compositionTotals : baseSlides[2].content.assetTable.totals)[3],
                  },
                ]
              : []),
          ],
        },
        {
          title: "Industry breakdown (STEP 7 inputs)",
          sheet: "Leases",
          rows: (industryBreakdown.length ? industryBreakdown : baseSlides[2].content.industryBreakdown).map((x) => ({
            Industry: x.name,
            "% of ABR": x.pct,
          })),
        },
        {
          title: "Geography breakdown (STEP 8 inputs)",
          sheet: "Properties + Leases",
          rows: (geographicBreakdown.length ? geographicBreakdown : baseSlides[2].content.geographicBreakdown).map((x) => ({
            State: x.state,
            "% of ABR": x.pct,
          })),
        },
      ],
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

    // Walkthrough formula (STEP 9): lease expiry is a single date if the
    // property has one lease; otherwise it is "Various".
    const leaseCountByProp = new Map();
    const expiryDateByProp = new Map(); // only populated for single-lease props
    if (expiryCol) {
      for (const l of leases) {
        const prop = String(l?.[leasePropertyCol] ?? "").trim();
        if (!prop) continue;
        leaseCountByProp.set(prop, (leaseCountByProp.get(prop) || 0) + 1);
        const d = excelDateToJsDate(l?.[expiryCol]);
        if (!d) continue;
        if (!expiryDateByProp.has(prop)) expiryDateByProp.set(prop, d);
      }
    }

    const creditByProp = (() => {
      // Walkthrough formula (STEP 9): credit is IG if any lease is Inv. Grade?,
      // otherwise use the credit rating for the highest-ABR lease.
      const tmp = new Map(); // prop -> { hasIG, bestAbr, bestRating }
      for (const l of leases) {
        const prop = String(l?.[leasePropertyCol] ?? "").trim();
        if (!prop) continue;
        const abr = parseNumber(l?.[abrCol]);

        const igRaw = invGradeCol ? String(l?.[invGradeCol] ?? "").trim().toLowerCase() : "";
        const isIG = igRaw === "yes" || igRaw === "y" || igRaw === "true";

        const rec = tmp.get(prop) || { hasIG: false, bestAbr: 0, bestRating: null };
        if (isIG) rec.hasIG = true;

        if (!rec.hasIG && ratingCol && l?.[ratingCol] != null) {
          const rating = String(l[ratingCol]).trim();
          if (rating && abr >= rec.bestAbr) {
            rec.bestAbr = abr;
            rec.bestRating = rating;
          }
        }

        tmp.set(prop, rec);
      }

      const out = new Map();
      for (const [prop, rec] of tmp.entries()) {
        out.set(prop, rec.hasIG ? "IG" : rec.bestRating || "—");
      }
      return out;
    })();

    return entries.map(([prop, abr]) => {
      const leased = leasedSfByProp.get(prop) || 0;
      const meta = propMeta.get(prop) || { name: prop, type: "—", rentableSF: 0 };
      const occ = meta.rentableSF > 0 ? ((leased / meta.rentableSF) * 100).toFixed(0) + "%" : "—";
      const w = waltByProp.get(prop);
      const waltProp = w && w.abr > 0 ? (w.weighted / w.abr).toFixed(1) + "y" : "—";
      const leaseCount = leaseCountByProp.get(prop) || 0;
      return {
        property: meta.name || prop,
        type: meta.type || "—",
        sf: meta.rentableSF ? fmtNum(meta.rentableSF) : (leased ? fmtNum(leased) : "—"),
        occupancy: occ,
        abr: fmtMoney(abr),
        walt: waltProp,
        leaseExpiry:
          expiryCol && leaseCount === 1
            ? expiryDateByProp.get(prop)
              ? monthYear(expiryDateByProp.get(prop))
              : "—"
            : expiryCol && leaseCount > 1
              ? "Various"
              : "—",
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
      sfExpiring: fmtSFExpirySchedule(r.sf),
      abrExpiring: fmtMoney(r.abr),
      pctPortfolio: totalABR > 0 ? ((r.abr / totalABR) * 100).toFixed(1) + "%" : "—",
    }));

    const longTerm = schedule.find((r) => r.year === "2031+");
    const keyNote =
      longTerm && totalABR > 0
        ? `The 2031+ bucket dominates at ${Math.round((longTerm.abr / totalABR) * 100)}%+ showing long-dated income visibility.`
        : baseSlides[4].content.keyNote;

    return {
      ...baseSlides[4],
      content: {
        headers: baseSlides[4].content.headers,
        schedule: mapped.length ? mapped : baseSlides[4].content.schedule,
        totalRow: {
          year: "Total",
          leases: totalLeasesExp || baseSlides[4].content.totalRow.leases,
          sfExpiring: totalSfExp > 0 ? fmtSFExpirySchedule(totalSfExp) : baseSlides[4].content.totalRow.sfExpiring,
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
              const d = expiryCol ? excelDateToJsDate(l?.[expiryCol]) : null;
              const year = d ? d.getFullYear() : null;
              out["Expiry Year"] = year ? String(year) : "—";
              out["Assigned Bucket"] = year ? (year >= 2031 ? "2031+" : String(year)) : "—";
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
    // Walkthrough formula (STEP 11): filter to Q4 2025 dispositions for the
    // transaction table + compute Q4 KPI cards from those totals; compute FY
    // 2025 KPI cards from all FY 2025 dispositions.
    const priceCol = "Gross Price";
    const sfTxCol = "Rentable SF";
    const occTxCol = "Occ % at Sale";
    const notesCol = "Notes";
    const quarterCol = "Quarter";
    const typeCol = "Type"; // this is transaction type in the raw template
    const dateCol = "Transaction Date" in (transactions[0] || {}) ? "Transaction Date" : findKey(transactions[0], [/date/i]);

    const assetTypeTxCol = "Asset Type" in (transactions[0] || {}) ? "Asset Type" : findKey(transactions[0], [/asset\\s*type/i, /\\btype\\b/i]);
    const propCol =
      findKey(transactions[0], [/property/i, /asset/i, /building/i, /site/i, /name/i]) || "Property Name";

    const cityCol = findKey(transactions[0], [/city/i]);
    const stateCol = findKey(transactions[0], [/state/i, /\\bst\\b/i]);

    const isDisposition = (t) => {
      const raw = typeCol ? String(t?.[typeCol] ?? "").trim().toLowerCase() : "";
      return raw.includes("disposition");
    };

    const isQ4_2025 = (t) => {
      const q = quarterCol ? String(t?.[quarterCol] ?? "").replace(/\s+/g, " ").trim().toLowerCase() : "";
      const date = dateCol ? excelDateToJsDate(t?.[dateCol]) : null;

      if (q) {
        const looksQ4 = q.includes("q4");
        const looks2025 = q.includes("2025") || q.includes("25");
        return looksQ4 && looks2025;
      }

      // Date-based fallback (if Quarter column is missing/misaligned)
      if (date && !isNaN(date.getTime())) {
        return date.getFullYear() === 2025 && date.getMonth() >= 9; // Oct-Dec
      }

      return false;
    };

    const isTotalRow = (t) => {
      const id = String(t?.["Transaction ID"] ?? "").trim().toLowerCase();
      return id.includes("total") || id.includes("q4 total") || id.includes("fy 2025");
    };

    const q4Txs = transactions.filter((t) => !isTotalRow(t) && isQ4_2025(t) && isDisposition(t));
    const fyTxs = transactions.filter((t) => {
      if (isTotalRow(t) || !isDisposition(t)) return false;
      const q = quarterCol ? String(t?.[quarterCol] ?? "").replace(/\s+/g, " ").trim().toLowerCase() : "";
      if (q) return q.includes("2025") || q.includes("25");
      if (dateCol) {
        const date = excelDateToJsDate(t?.[dateCol]);
        return date && !isNaN(date.getTime()) ? date.getFullYear() === 2025 : false;
      }
      return false;
    });

    // If Q4 filter returns empty, fall back to FY 2025 dispositions (better than empty trace).
    const effectiveTxs = q4Txs.length ? q4Txs : fyTxs.length ? fyTxs : transactions.filter((t) => !isTotalRow(t) && isDisposition(t));

    const grossProceedsQ4 = priceCol ? effectiveTxs.reduce((s, t) => s + parseNumber(t?.[priceCol]), 0) : 0;
    const grossProceedsFY = priceCol ? (fyTxs.length ? fyTxs : effectiveTxs).reduce((s, t) => s + parseNumber(t?.[priceCol]), 0) : 0;

    const extractMillionFromNotes = (notes, patterns) => {
      const text = notes ? String(notes) : "";
      for (const re of patterns) {
        const m = text.match(re);
        if (!m) continue;
        const num = parseFloat(m[1]);
        if (isFinite(num)) return num * 1e6;
      }
      return 0;
    };

    let carryingCostSavingsDollars = 0;
    let capexAvoidedDollars = 0;
    for (const t of effectiveTxs) {
      const notes = notesCol && t?.[notesCol] != null ? String(t[notesCol]) : "";
      carryingCostSavingsDollars += extractMillionFromNotes(notes, [
        /carry cost savings[^$]*\$\s*([0-9.]+)\s*M/i,
        /carry costs[^$]*\$\s*([0-9.]+)\s*M/i,
      ]);
      capexAvoidedDollars += extractMillionFromNotes(notes, [
        /avoids[^$]*\$\s*([0-9.]+)\s*M/i,
        /capex[^$]*\$\s*([0-9.]+)\s*M/i,
      ]);
    }

    const mappedTx = effectiveTxs.slice(0, 12).map((t) => {
      const city = cityCol ? String(t?.[cityCol] ?? "").trim() : "";
      const state = stateCol ? String(t?.[stateCol] ?? "").trim() : "";
      const region = city && state ? `${city}, ${state}` : city || state || "—";

      return {
        property: String(t?.[propCol] ?? "—"),
        region,
        assetType: assetTypeTxCol ? String(t?.[assetTypeTxCol] ?? "—") : "—",
        sf: sfTxCol ? fmtNum(parseNumber(t?.[sfTxCol])) : "—",
        price: priceCol ? fmtMoney(parseNumber(t?.[priceCol])) : "—",
        occupancy: occTxCol ? safePctStringWhole(t?.[occTxCol]) : "—",
        strategicReason: notesCol ? String(t?.[notesCol] ?? "—") : "—",
      };
    });

    const carryingCostSavings =
      carryingCostSavingsDollars > 0 ? `${fmtMoney(carryingCostSavingsDollars)} annually` : baseSlides[5].content.metrics.carryingCostSavings;
    const capexAvoided =
      capexAvoidedDollars > 0 ? fmtMoney(capexAvoidedDollars) : baseSlides[5].content.metrics.capexAvoided;

    const soldFY = fyTxs.length ? fyTxs.length : effectiveTxs.length;

    const effectiveTxsTraceRows =
      effectiveTxs.length > 0
        ? effectiveTxs
        : [
            {
              Message: "No Q4 2025 dispositions matched filters.",
              Filters: {
                quarterCol,
                quarterMatches: "q includes 'q4' + '2025'/'25' (or date-based Oct-Dec 2025)",
                typeCol,
                typeMatches: "type includes 'disposition'",
              },
            },
          ];

    return {
      ...baseSlides[5],
      content: {
        metrics: {
          q4PropertiesSold: fmtNum(effectiveTxs.length),
          q4GrossProceeds: grossProceedsQ4 > 0 ? fmtMoney(grossProceedsQ4) : baseSlides[5].content.metrics.q4GrossProceeds,
          fy2025Sold: fmtNum(soldFY),
          fy2025Proceeds: grossProceedsFY > 0 ? fmtMoney(grossProceedsFY) : baseSlides[5].content.metrics.fy2025Proceeds,
          carryingCostSavings,
          capexAvoided,
        },
        transactions: mappedTx.length ? mappedTx : baseSlides[5].content.transactions,
        fullYearNote: baseSlides[5].content.fullYearNote,
        sourceTables: [
          {
            title: "Q4 disposition transactions used in Slide 6",
            sheet: "Transactions",
            rows: effectiveTxsTraceRows,
          },
          {
            title: "Disposition KPI totals (STEP 11)",
            sheet: "Transactions",
            rows: [
              {
                "KPI Card": "Q4 Properties Sold",
                Value: fmtNum(effectiveTxs.length),
                Logic: "COUNT dispositions where Quarter = Q4 2025 (fallback if none)",
              },
              {
                "KPI Card": "Q4 Gross Proceeds",
                Value: grossProceedsQ4 > 0 ? fmtMoney(grossProceedsQ4) : "—",
                Logic: "SUM(Gross Price) for the same Q4 disposition set",
              },
              {
                "KPI Card": "FY 2025 Dispositions",
                Value: fmtNum(soldFY),
                Logic: "COUNT dispositions with Quarter containing 2025 (full year totals)",
              },
              {
                "KPI Card": "FY 2025 Proceeds",
                Value: grossProceedsFY > 0 ? fmtMoney(grossProceedsFY) : "—",
                Logic: "SUM(Gross Price) for FY 2025 disposition set",
              },
              { "KPI Card": "Carry cost savings (annual)", Value: carryingCostSavings, Logic: "Parsed from Notes text" },
              { "KPI Card": "Capex avoided", Value: capexAvoided, Logic: "Parsed from Notes text" },
            ],
          },
        ],
      },
    };
  })();

  // ─── Slide 1: Cover highlights (align to workbook where available) ────
  // Intentionally disabled: Slide 1 is hard-coded for MVP screenshot fidelity.

  // ─── Slide 7 & 8: Outlook and Appendix (use template) ───────────────────
  const outlookSlide = {
    ...baseSlides[6],
    content: {
      ...baseSlides[6].content,
      sourceTables: [
        {
          title: "Forward Outlook source (Slide 7)",
          sheet: "Template (not computed)",
          rows: [
            {
              Field: "Leasing Pipeline",
              Value: "Hardcoded",
              Logic: "No mapping step for slide 7 exists in Quarterly walkthrough sheets; not derived from uploaded Summary/Leases/Properties/Transactions.",
            },
            {
              Field: "Disposition Pipeline",
              Value: "Hardcoded",
              Logic: "Not derived from uploaded Transactions in current implementation.",
            },
            {
              Field: "Capital Allocation Strategy",
              Value: "Hardcoded",
              Logic: "Template narrative values only.",
            },
          ],
        },
      ],
    },
  };
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
