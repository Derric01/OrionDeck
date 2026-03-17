// Slide-by-Slide Data — Orion Q4 2025 Portfolio Report Template
// Matches the exact structure for the report viewer and download template.

const baseSlides = [
  // Slide 1 — Cover / Portfolio Snapshot (matches PPTX template layout)
  {
    id: 1,
    title: "Cover / Portfolio Snapshot",
    type: "cover",
    content: {
      title: "Orion Properties Inc.",
      reportTitle: "Q4 2025 Quarterly Portfolio Report",
      dateLine: "March 2026 NYSE: ONL",
      kpis: [
        { cardLabel: "PROPERTIES", value: "58", subLabel: "Operating" },
        { cardLabel: "OCCUPANCY", value: "78.7%", subLabel: "% rentable SF" },
        { cardLabel: "WALT", value: "5.7 yrs", subLabel: "wtd. avg. lease term" },
        { cardLabel: "ANNUALISED BASE RENT", value: "$111.3M", subLabel: "ABR, Dec 31 2025" },
        { cardLabel: "ABR / SF", value: "$16.52", subLabel: "active leases." },
        { cardLabel: "INV-GRADE TENANCY", value: "66.7%", subLabel: "% of ABR" },
      ],
      highlightsSectionTitle: "Q4 2025 HIGHLIGHTS",
      highlights: [
        { main: "183,000 SF leased", sub: "New & renewed leases in Q4" },
        { main: "$32.0M in dispositions", sub: "3 properties sold in Q4" },
        { main: "$80.7M FY 2025", sub: "10 properties sold full year" },
        { main: "$0.02/sh", sub: "Q4 2025 dividend declared" },
      ],
    },
  },
  // Slide 2 — Portfolio Highlights
  {
    id: 2,
    title: "Portfolio Highlights",
    type: "portfolioHighlights",
    content: {
      metrics: {
        operatingProperties: "58",
        totalLeasableSF: "6.74M SF",
        abr: "$111.3M",
        occupancy: "78.7%",
        walt: "5.7 years",
        igTenancy: "66.7% of ABR",
      },
      top10Tenants: [
        { rank: 1, name: "US Federal Agency", creditRating: "AA+", pctABR: "17.8%" },
        { rank: 2, name: "Investment-Grade A", creditRating: "A-", pctABR: "10.0%" },
        { rank: 3, name: "Investment-Grade B", creditRating: "BB", pctABR: "7.1%" },
        { rank: 4, name: "Investment-Grade C", creditRating: "A-", pctABR: "4.4%" },
        { rank: 5, name: "Investment-Grade D", creditRating: "BBB", pctABR: "4.1%" },
        { rank: 6, name: "Investment-Grade E", creditRating: "BBB", pctABR: "3.7%" },
        { rank: 7, name: "Investment-Grade F", creditRating: "BB+", pctABR: "3.5%" },
        { rank: 8, name: "Investment-Grade G", creditRating: "A", pctABR: "3.4%" },
        { rank: 9, name: "Investment-Grade H", creditRating: "BB", pctABR: "3.3%" },
        { rank: 10, name: "Investment-Grade I", creditRating: "BBB+", pctABR: "3.2%" },
      ],
      combinedNote: "Combined Top-10 tenants contribute 60.5% of ABR and 66.7% are investment grade.",
    },
  },
  // Slide 3 — Portfolio Composition
  {
    id: 3,
    title: "Portfolio Composition",
    type: "composition",
    content: {
      igSplit: { investmentGrade: "66.7%", nonInvestmentGrade: "33.3%" },
      assetTable: {
        headers: ["Asset Type", "Properties", "Square Feet", "% of ABR"],
        rows: [
          ["Traditional Office", "28", "3.08M SF", "43.2%"],
          ["Dedicated Use / Corporate", "14", "1.68M SF", "35.8%"],
          ["Medical Office / Life Sciences", "9", "0.96M SF", "12.4%"],
          ["Government / Public Sector", "4", "0.62M SF", "5.3%"],
          ["Flex / Industrial", "3", "0.40M SF", "3.3%"],
        ],
        totals: ["Total", "58", "6.74M SF", "100%"],
      },
      industryBreakdown: [
        { name: "Government & Public Services", pct: "18.3%" },
        { name: "Healthcare Equipment & Services", pct: "13.7%" },
        { name: "Capital Goods", pct: "10.8%" },
        { name: "Financial Institutions", pct: "10%" },
        { name: "Software & Services", pct: "9%" },
        { name: "Materials", pct: "7.7%" },
        { name: "Telecom Services", pct: "6.7%" },
        { name: "Commercial & Professional Services", pct: "5%" },
        { name: "Consumer Durables & Apparel", pct: "4.1%" },
        { name: "Transportation", pct: "4%" },
        { name: "Other", pct: "10.7%" },
      ],
      geographicBreakdown: [
        { state: "Texas", pct: "18.9%" },
        { state: "New Jersey", pct: "13.4%" },
        { state: "New York", pct: "10%" },
        { state: "Kentucky", pct: "9.5%" },
        { state: "Colorado", pct: "6%" },
        { state: "California", pct: "5.4%" },
        { state: "Maryland", pct: "4.4%" },
        { state: "Virginia", pct: "4.2%" },
        { state: "Georgia", pct: "4.2%" },
        { state: "Tennessee", pct: "4.2%" },
        { state: "Other states (18 total)", pct: "19.8%" },
      ],
    },
  },
  // Slide 4 — Asset-by-Asset Performance
  {
    id: 4,
    title: "Asset-by-Asset Performance",
    type: "assetPerformance",
    content: {
      headers: ["Property", "Type", "SF", "Occupancy", "ABR", "WALT", "Lease Expiry", "Credit"],
      rows: [
        { property: "Lake Cook Corp. Center – Deerfield, IL", type: "Dedicated Use", sf: "574,604", occupancy: "100%", abr: "$3.1M", walt: "6.0y", leaseExpiry: "Dec 2031", credit: "IG" },
        { property: "McLean Gov't Centre – McLean, VA", type: "Government", sf: "186,000", occupancy: "100%", abr: "$3.3M", walt: "6.7y", leaseExpiry: "Sep 2032", credit: "AA+" },
        { property: "Houston Corporate – Houston, TX", type: "Dedicated Use", sf: "284,000", occupancy: "100%", abr: "$3.2M", walt: "5.1y", leaseExpiry: "Jan 2031", credit: "BBB+" },
        { property: "Alpharetta Campus – GA", type: "Dedicated Use", sf: "218,000", occupancy: "100%", abr: "$2.4M", walt: "5.5y", leaseExpiry: "Jun 2031", credit: "BB+" },
        { property: "Albany Gov't Services – NY", type: "Government", sf: "168,000", occupancy: "100%", abr: "$4.2M", walt: "7.5y", leaseExpiry: "Jun 2033", credit: "AA+" },
        { property: "NJ Princeton Research – Plainsboro", type: "Medical Office", sf: "128,000", occupancy: "100%", abr: "$3.6M", walt: "5.0y", leaseExpiry: "Dec 2030", credit: "A+" },
        { property: "Parsippany Campus – NJ", type: "Traditional", sf: "166,000", occupancy: "62%", abr: "$1.8M", walt: "7.8y", leaseExpiry: "Various", credit: "IG" },
        { property: "Buffalo Campus – NY", type: "Dedicated Use", sf: "160,000", occupancy: "100%", abr: "$2.1M", walt: "2.3y", leaseExpiry: "Apr 2028", credit: "BBB+" },
        { property: "Louisville Medical – KY", type: "Medical Office", sf: "142,000", occupancy: "100%", abr: "$3.6M", walt: "4.0y", leaseExpiry: "Dec 2029", credit: "AA-" },
        { property: "Phoenix Medical – AZ", type: "Medical Office", sf: "92,000", occupancy: "100%", abr: "$2.1M", walt: "9.3y", leaseExpiry: "Mar 2036", credit: "A" },
      ],
    },
  },
  // Slide 5 — Lease Expiry Schedule
  {
    id: 5,
    title: "Lease Expiry Schedule",
    type: "expiry",
    content: {
      headers: ["Year", "Leases", "SF Expiring", "ABR Expiring", "% Portfolio"],
      schedule: [
        { year: "2026", leases: 8, sfExpiring: "680K SF", abrExpiring: "$6.2M", pctPortfolio: "5.4%" },
        { year: "2027", leases: 11, sfExpiring: "940K SF", abrExpiring: "$8.8M", pctPortfolio: "7.7%" },
        { year: "2028", leases: 9, sfExpiring: "820K SF", abrExpiring: "$10.2M", pctPortfolio: "8.9%" },
        { year: "2029", leases: 7, sfExpiring: "610K SF", abrExpiring: "$7.4M", pctPortfolio: "6.5%" },
        { year: "2030", leases: 6, sfExpiring: "580K SF", abrExpiring: "$6.8M", pctPortfolio: "6.0%" },
        { year: "2031+", leases: 22, sfExpiring: "4,970K SF", abrExpiring: "$74.5M", pctPortfolio: "65.4%" },
      ],
      totalRow: { year: "Total", leases: 63, sfExpiring: "7,600K SF", abrExpiring: "$111.3M", pctPortfolio: "100%" },
      keyNote: "65.4% of ABR expires in 2031 or later, indicating strong long-term income visibility and limited near-term rollover risk.",
    },
  },
  // Slide 6 — Q4 2025 Disposition Activity
  {
    id: 6,
    title: "Q4 2025 Disposition Activity",
    type: "dispositions",
    content: {
      metrics: {
        q4PropertiesSold: "3",
        q4GrossProceeds: "$32.0M",
        fy2025Sold: "10",
        fy2025Proceeds: "$80.7M",
        carryingCostSavings: "$12.4M annually",
        capexAvoided: "~$95M",
      },
      transactions: [
        { property: "Southeast US", region: "Traditional Office", assetType: "Traditional Office", sf: "127,000", price: "$11.8M", occupancy: "0%", strategicReason: "Vacant cost elimination" },
        { property: "Midwest US", region: "Traditional Office", assetType: "Traditional Office", sf: "98,000", price: "$9.4M", occupancy: "18%", strategicReason: "Anchor lease expiry 2026" },
        { property: "Sun Belt", region: "Traditional Office", assetType: "Traditional Office", sf: "102,000", price: "$10.8M", occupancy: "78%", strategicReason: "Non-core asset sale" },
      ],
      fullYearNote: "Selling weaker office assets reduces costs and accelerates portfolio transformation.",
    },
  },
  // Slide 7 — Forward Outlook
  {
    id: 7,
    title: "Forward Outlook",
    type: "outlook",
    content: {
      leasingPipeline: {
        sfUnderNegotiation: "380K SF",
        propertyCount: 8,
        recentActivity: [
          "3-year, 160,000 SF extension in Buffalo",
          "10.5-year, 23,000 SF lease in Phoenix",
        ],
      },
      dispositionPipeline: {
        sfCompleted: "516K SF",
        note: "Additional property sales completed after quarter end. More traditional office assets expected to be sold in 2026.",
      },
      capitalAllocation: {
        dedicatedUsePct: "35.8% of ABR",
        targetPct: "50%+ of ABR",
        priorities: [
          "Corporate HQs, life science facilities, and medical office properties",
          "Longer leases and higher tenant switching costs",
        ],
      },
    },
  },
  // Slide 8 — Data Pipeline Appendix
  {
    id: 8,
    title: "Data Pipeline Appendix",
    type: "appendix",
    content: {
      step1: {
        title: "RAW DATA INGESTION",
        bigNumber: "2,204",
        bigLabel: "raw data fields",
        detail:
          "61 property records + 81 individual lease records + 3 transaction records loaded from your property management system and CoStar API.",
      },
      step2: {
        title: "AGGREGATION ENGINE",
        bigNumber: "1ms",
        bigLabel: "aggregation time",
        detail:
          "WALT, occupancy %, ABR, investment-grade %, industry classification, geographic breakdown, and lease expiry schedule all calculated from individual lease records.",
      },
      step3: {
        title: "REPORT GENERATION",
        bigNumber: "8 slides",
        bigLabel: "institutional quality output",
        detail:
          "Finished presentation generated with all KPI cards, tables, charts, and narrative sections. Every number traces back to an individual lease record.",
      },
      footer:
        "BRAIND — The Operating System for Commercial Real Estate • Every number in this report is traceable to its source data.",
    },
  },
];

let currentSlides = JSON.parse(JSON.stringify(baseSlides));

function getSlides() {
  return currentSlides;
}

function resetSlides() {
  currentSlides = JSON.parse(JSON.stringify(baseSlides));
}

function updateSlide(slideId, field, value) {
  const slide = currentSlides.find((s) => s.id === parseInt(slideId));
  if (!slide) return { success: false, message: `Slide ${slideId} not found` };

  if (slide.type === "cover" && slide.content.kpis) {
    const kpi = slide.content.kpis.find(
      (k) => (k.cardLabel || k.label || "").toLowerCase().includes(field.toLowerCase())
    );
    if (kpi) {
      kpi.value = value;
      return { success: true, slide };
    }
  }

  if (slide.type === "portfolioHighlights" && slide.content.metrics && field in slide.content.metrics) {
    slide.content.metrics[field] = value;
    return { success: true, slide };
  }

  if (slide.content && field in slide.content) {
    slide.content[field] = value;
    return { success: true, slide };
  }

  return { success: false, message: `Field '${field}' not found in slide ${slideId}` };
}

function addNote(slideId, note) {
  const slide = currentSlides.find((s) => s.id === parseInt(slideId));
  if (!slide) return { success: false, message: `Slide ${slideId} not found` };

  if (!slide.content.notes) slide.content.notes = [];
  slide.content.notes.push(note);
  return { success: true, slide };
}

function setSlidesFromPortfolioData(parsedData) {
  const { buildSlidesFromParsedData } = require("./buildSlidesFromExcel");
  currentSlides = buildSlidesFromParsedData(parsedData, baseSlides);
}

module.exports = { getSlides, resetSlides, setSlidesFromPortfolioData, updateSlide, addNote, baseSlides };
