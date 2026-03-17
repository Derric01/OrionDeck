const PptxGenJS = require("pptxgenjs");

// ─── Theme ────────────────────────────────────────────────────────────────────
const T = {
  bg: "0F1117",
  headerBg: "1A1F2E",
  cardBg: "1E2740",
  cardBg2: "16213A",
  accent: "6366F1",
  accentLight: "818CF8",
  accentGlow: "A5B4FC",
  text: "F1F5F9",
  textMuted: "94A3B8",
  textDim: "64748B",
  border: "2D3A52",
  green: "10B981",
  red: "EF4444",
  amber: "F59E0B",
  white: "FFFFFF",
};

const SLIDE_W = 13.33;
const SLIDE_H = 7.5;
const CONTENT_X = 0.35;
const CONTENT_Y = 0.9;
const CONTENT_W = SLIDE_W - 0.7;

// ─── Header decoration on every slide ────────────────────────────────────────
function addSlideChrome(pptx, pSlide, slide, total) {
  // Dark header bar
  pSlide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: SLIDE_W, h: 0.7,
    fill: { color: T.headerBg },
    line: { color: T.border, width: 0.75 },
  });
  // Accent left bar
  pSlide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 0, w: 0.06, h: SLIDE_H,
    fill: { color: T.accent },
  });
  // Slide title in header
  pSlide.addText(slide.title.toUpperCase(), {
    x: 0.2, y: 0.16, w: 10, h: 0.42,
    fontSize: 13, bold: true, color: T.accentLight,
    charSpacing: 1.5,
  });
  // Slide counter
  pSlide.addText(`${slide.id} / ${total}`, {
    x: 12.3, y: 0.22, w: 0.9, h: 0.3,
    fontSize: 9, color: T.textDim, align: "right",
  });
  // Branding
  pSlide.addText("OrionDeck  |  Q4 2025", {
    x: 10.2, y: 0.22, w: 2, h: 0.3,
    fontSize: 8, color: T.textDim, align: "right",
  });
}

// ─── Slide builders ───────────────────────────────────────────────────────────

function buildTitle(pptx, pSlide, content) {
  pSlide.background = { color: T.bg };
  // Large gradient-ish decoration rectangle
  pSlide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 2.5, w: SLIDE_W, h: 3.2,
    fill: { color: T.headerBg },
  });
  pSlide.addShape(pptx.ShapeType.rect, {
    x: 0, y: 2.5, w: 0.08, h: 3.2,
    fill: { color: T.accent },
  });

  pSlide.addText(content.company, {
    x: 0.5, y: 1.1, w: SLIDE_W - 1, h: 1.3,
    fontSize: 46, bold: true, color: T.text, align: "center",
  });
  pSlide.addText(content.subtitle || "Portfolio Performance Report", {
    x: 0.5, y: 2.55, w: SLIDE_W - 1, h: 0.7,
    fontSize: 22, color: T.accentLight, align: "center",
  });
  pSlide.addText(content.period, {
    x: 0.5, y: 3.35, w: SLIDE_W - 1, h: 0.55,
    fontSize: 18, color: T.textMuted, align: "center",
  });
  pSlide.addText(content.date, {
    x: 0.5, y: 3.95, w: SLIDE_W - 1, h: 0.4,
    fontSize: 13, color: T.textDim, align: "center",
  });
  pSlide.addText(content.preparedBy, {
    x: 0.5, y: 5.8, w: SLIDE_W - 1, h: 0.35,
    fontSize: 10, color: T.textDim, align: "center", italic: true,
  });
}

function buildKPI(pptx, pSlide, content) {
  const kpis = content.kpis || [];
  const cols = 3;
  const cardW = (CONTENT_W - (cols - 1) * 0.2) / cols;
  const cardH = 1.55;
  const rows = Math.ceil(kpis.length / cols);

  // Centre vertically
  const totalH = rows * cardH + (rows - 1) * 0.2;
  const startY = CONTENT_Y + (SLIDE_H - CONTENT_Y - totalH) / 2 - 0.15;

  kpis.forEach((kpi, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = CONTENT_X + col * (cardW + 0.2);
    const y = startY + row * (cardH + 0.2);

    pSlide.addShape(pptx.ShapeType.roundRect, {
      x, y, w: cardW, h: cardH,
      fill: { color: T.cardBg },
      line: { color: T.border, width: 0.75 },
      rectRadius: 0.1,
    });
    // Accent top stripe
    pSlide.addShape(pptx.ShapeType.rect, {
      x, y, w: cardW, h: 0.06,
      fill: { color: T.accent },
    });

    const valStr = `${kpi.value}${kpi.unit && kpi.unit !== "" ? " " + kpi.unit : ""}`;
    pSlide.addText(valStr, {
      x: x + 0.15, y: y + 0.2, w: cardW - 0.3, h: 0.82,
      fontSize: 28, bold: true, color: T.accentGlow,
    });
    pSlide.addText(kpi.label, {
      x: x + 0.15, y: y + 1.08, w: cardW - 0.3, h: 0.38,
      fontSize: 10, color: T.textMuted, wrap: true,
    });
  });
}

function buildTable(pptx, pSlide, content) {
  const { headers, rows, totals } = content;

  const headerRow = headers.map((h) => ({
    text: h,
    options: { bold: true, color: T.accentLight, fill: { color: T.cardBg }, fontSize: 10 },
  }));

  const dataRows = rows.map((row, ri) =>
    row.map((cell) => ({
      text: String(cell),
      options: { color: T.text, fill: { color: ri % 2 === 0 ? T.cardBg2 : T.cardBg }, fontSize: 10 },
    }))
  );

  let allRows = [headerRow, ...dataRows];

  if (totals) {
    allRows.push(
      totals.map((cell) => ({
        text: String(cell),
        options: { bold: true, color: T.accentLight, fill: { color: "0D1526" }, fontSize: 10 },
      }))
    );
  }

  pSlide.addTable(allRows, {
    x: CONTENT_X,
    y: CONTENT_Y,
    w: CONTENT_W,
    h: SLIDE_H - CONTENT_Y - 0.3,
    border: { type: "solid", color: T.border, pt: 0.5 },
    valign: "middle",
    colW: headers.map((_, i) => {
      // First column wider
      if (i === 0) return CONTENT_W * 0.28;
      return (CONTENT_W * 0.72) / (headers.length - 1);
    }),
  });
}

function buildMetrics(pptx, pSlide, content) {
  const { occupancy, walt, notes } = content;

  // ── Left: Occupancy ──
  pSlide.addText("OCCUPANCY BY ASSET TYPE", {
    x: CONTENT_X, y: CONTENT_Y, w: 6, h: 0.3,
    fontSize: 9, bold: true, color: T.accentLight, charSpacing: 1,
  });

  const occEntries = Object.entries(occupancy);
  const occCardW = 1.72;
  const occCardH = 1.05;
  occEntries.forEach(([key, val], i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = CONTENT_X + col * (occCardW + 0.18);
    const y = 1.32 + row * (occCardH + 0.15);
    const is100 = val === "100.0%" || val === "100%";

    pSlide.addShape(pptx.ShapeType.roundRect, {
      x, y, w: occCardW, h: occCardH,
      fill: { color: T.cardBg },
      line: { color: is100 ? T.green : T.amber, width: 0.75 },
      rectRadius: 0.08,
    });
    pSlide.addText(val, {
      x: x + 0.08, y: y + 0.1, w: occCardW - 0.16, h: 0.6,
      fontSize: 22, bold: true, color: is100 ? T.green : T.amber, align: "center",
    });
    pSlide.addText(key.replace(/([A-Z])/g, " $1").replace(/^./, (s) => s.toUpperCase()).trim(), {
      x: x + 0.06, y: y + 0.68, w: occCardW - 0.12, h: 0.3,
      fontSize: 8, color: T.textMuted, align: "center", wrap: true,
    });
  });

  // ── Right: WALT ──
  pSlide.addText("WALT BY ASSET TYPE", {
    x: 6.8, y: CONTENT_Y, w: 6.2, h: 0.3,
    fontSize: 9, bold: true, color: T.accentLight, charSpacing: 1,
  });

  const maxWalt = 8;
  walt.byAssetType.forEach((item, i) => {
    const y = 1.32 + i * 0.82;
    const barMaxW = 4.5;
    const barW = Math.max(0.15, (parseFloat(item.walt) / maxWalt) * barMaxW);

    pSlide.addText(item.type, {
      x: 6.8, y: y + 0.2, w: 2.3, h: 0.35,
      fontSize: 10, color: T.text,
    });
    // Bar track
    pSlide.addShape(pptx.ShapeType.rect, {
      x: 9.25, y: y + 0.25, w: barMaxW, h: 0.22,
      fill: { color: T.cardBg2 }, line: { color: T.border, width: 0.5 },
    });
    // Bar fill
    pSlide.addShape(pptx.ShapeType.rect, {
      x: 9.25, y: y + 0.25, w: barW, h: 0.22,
      fill: { color: T.accent },
    });
    pSlide.addText(item.walt, {
      x: 13.85, y: y + 0.2, w: 0.9, h: 0.35,
      fontSize: 10, bold: true, color: T.accentLight, align: "right",
    });
  });

  if (notes) {
    const notesY = 5.65;
    pSlide.addShape(pptx.ShapeType.rect, {
      x: CONTENT_X, y: notesY - 0.1, w: CONTENT_W, h: 0.04,
      fill: { color: T.border },
    });
    notes.forEach((note, i) => {
      pSlide.addText(`• ${note}`, {
        x: CONTENT_X, y: notesY + i * 0.3, w: CONTENT_W, h: 0.28,
        fontSize: 8.5, color: T.textMuted, italic: true,
      });
    });
  }
}

function buildTransactions(pptx, pSlide, content) {
  const { summary, transactions: txs } = content;

  // Summary metric cards
  const sumEntries = Object.entries(summary).slice(0, 6);
  const cardW = (CONTENT_W - (sumEntries.length - 1) * 0.15) / sumEntries.length;
  sumEntries.forEach(([key, val], i) => {
    const x = CONTENT_X + i * (cardW + 0.15);
    pSlide.addShape(pptx.ShapeType.roundRect, {
      x, y: CONTENT_Y, w: cardW, h: 1.0,
      fill: { color: T.cardBg },
      line: { color: T.border, width: 0.75 },
      rectRadius: 0.08,
    });
    pSlide.addText(String(val), {
      x: x + 0.1, y: CONTENT_Y + 0.08, w: cardW - 0.2, h: 0.56,
      fontSize: 20, bold: true, color: T.accentLight, align: "center",
    });
    pSlide.addText(
      key.replace(/([A-Z])/g, " $1").replace(/Fy/g, "FY").trim(),
      {
        x: x + 0.05, y: CONTENT_Y + 0.64, w: cardW - 0.1, h: 0.28,
        fontSize: 8, color: T.textMuted, align: "center",
      }
    );
  });

  // Transaction table
  const tHeaders = ["Property", "Location", "Date", "SF", "Gross Price", "Net Proceeds", "Gain/(Loss)"];
  const headerRow = tHeaders.map((h) => ({
    text: h,
    options: { bold: true, color: T.accentLight, fill: { color: T.cardBg }, fontSize: 9 },
  }));
  const dataRows = txs.map((tx, ri) => {
    const isLoss = String(tx.gainLoss).startsWith("(");
    return [
      tx.property, tx.city, tx.date, tx.sf,
      tx.grossPrice, tx.netProceeds, tx.gainLoss,
    ].map((cell, ci) => ({
      text: String(cell),
      options: {
        color: ci === 6 ? (isLoss ? T.red : T.green) : T.text,
        fill: { color: ri % 2 === 0 ? T.cardBg2 : T.cardBg },
        fontSize: 9,
      },
    }));
  });

  pSlide.addTable([headerRow, ...dataRows], {
    x: CONTENT_X,
    y: 2.1,
    w: CONTENT_W,
    h: 3.8,
    border: { type: "solid", color: T.border, pt: 0.5 },
    valign: "middle",
    colW: [3.5, 1.4, 0.95, 0.85, 1.3, 1.3, 1.3].map(
      (v) => v * (CONTENT_W / 10.6)
    ),
  });

  // Notes row
  pSlide.addText("All dispositions: Traditional Office assets. Net proceeds earmarked for dedicated-use acquisitions.", {
    x: CONTENT_X, y: 6.25, w: CONTENT_W, h: 0.28,
    fontSize: 8.5, color: T.textDim, italic: true,
  });
}

function buildTenants(pptx, pSlide, content) {
  const tHeaders = ["Tenant", "Asset Type", "Rating", "ABR", "SF", "Expiry", "Rem. Term"];
  const headerRow = tHeaders.map((h) => ({
    text: h,
    options: { bold: true, color: T.accentLight, fill: { color: T.cardBg }, fontSize: 10 },
  }));

  const ratingColor = (r) => {
    if (!r) return T.textMuted;
    const u = r.toUpperCase();
    if (u.startsWith("AA") || u.startsWith("AAA")) return T.green;
    if (u.startsWith("A")) return "#34D399";
    if (u.startsWith("BBB")) return T.amber;
    return T.red;
  };

  const dataRows = content.tenants.map((t, ri) =>
    [t.name, t.assetType, t.creditRating, t.abr, t.leasedSF, t.expiry, t.remainingTerm].map(
      (cell, ci) => ({
        text: String(cell),
        options: {
          color: ci === 2 ? ratingColor(t.creditRating) : T.text,
          bold: ci === 2,
          fill: { color: ri % 2 === 0 ? T.cardBg2 : T.cardBg },
          fontSize: 10,
        },
      })
    )
  );

  pSlide.addTable([headerRow, ...dataRows], {
    x: CONTENT_X,
    y: CONTENT_Y,
    w: CONTENT_W,
    h: SLIDE_H - CONTENT_Y - 0.35,
    border: { type: "solid", color: T.border, pt: 0.5 },
    valign: "middle",
    colW: [2.8, 1.5, 0.9, 1.3, 1.1, 1.2, 1.0].map(
      (v) => v * (CONTENT_W / 9.8)
    ),
  });
}

function buildExpiry(pptx, pSlide, content) {
  const { schedule, notes } = content;

  const tHeaders = ["Period", "# Leases", "SF", "ABR", "% of ABR"];
  const headerRow = tHeaders.map((h) => ({
    text: h,
    options: { bold: true, color: T.accentLight, fill: { color: T.cardBg }, fontSize: 11 },
  }));

  const dataRows = schedule.map((row, ri) => {
    const isLast = row.period === "2031+";
    return [row.period, String(row.leases), row.sf, row.abr, row.pctOfABR].map((cell) => ({
      text: cell,
      options: {
        color: isLast ? T.accentGlow : T.text,
        bold: isLast,
        fill: { color: isLast ? T.cardBg : ri % 2 === 0 ? T.cardBg2 : T.cardBg },
        fontSize: 11,
      },
    }));
  });

  pSlide.addTable([headerRow, ...dataRows], {
    x: CONTENT_X,
    y: CONTENT_Y,
    w: CONTENT_W,
    h: 3.9,
    border: { type: "solid", color: T.border, pt: 0.5 },
    valign: "middle",
  });

  if (notes) {
    const notesY = 5.25;
    pSlide.addShape(pptx.ShapeType.rect, {
      x: CONTENT_X, y: notesY - 0.12, w: CONTENT_W, h: 0.04,
      fill: { color: T.accent },
    });
    notes.forEach((note, i) => {
      pSlide.addText(`— ${note}`, {
        x: CONTENT_X, y: notesY + 0.05 + i * 0.4, w: CONTENT_W, h: 0.36,
        fontSize: 10.5, color: T.textMuted,
      });
    });
  }
}

function buildHighlights(pptx, pSlide, content) {
  const sections = content.highlights || [];
  const cols = sections.length;
  const colW = (CONTENT_W - (cols - 1) * 0.2) / cols;
  const colH = SLIDE_H - CONTENT_Y - 0.3;

  sections.forEach((section, i) => {
    const x = CONTENT_X + i * (colW + 0.2);
    const y = CONTENT_Y;

    // Card
    pSlide.addShape(pptx.ShapeType.roundRect, {
      x, y, w: colW, h: colH,
      fill: { color: T.cardBg },
      line: { color: T.border, width: 0.75 },
      rectRadius: 0.12,
    });
    // Accent top stripe
    pSlide.addShape(pptx.ShapeType.rect, {
      x, y, w: colW, h: 0.07,
      fill: { color: T.accent },
    });

    // Category
    pSlide.addText(section.category, {
      x: x + 0.18, y: y + 0.18, w: colW - 0.36, h: 0.5,
      fontSize: 14, bold: true, color: T.accentLight,
    });

    // Divider
    pSlide.addShape(pptx.ShapeType.rect, {
      x: x + 0.18, y: y + 0.72, w: colW - 0.36, h: 0.03,
      fill: { color: T.border },
    });

    // Points
    section.points.forEach((pt, j) => {
      pSlide.addText(`• ${pt}`, {
        x: x + 0.18,
        y: y + 0.88 + j * 1.2,
        w: colW - 0.36,
        h: 1.1,
        fontSize: 11,
        color: T.text,
        valign: "top",
        wrap: true,
      });
    });
  });
}

// ─── Main export ─────────────────────────────────────────────────────────────

async function generatePPTXBuffer(slides) {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.title = "Orion Properties Inc. — Q4 2025 Portfolio Report";
  pptx.subject = "Portfolio Performance Report";
  pptx.author = "OrionDeck AI";

  for (const slide of slides) {
    const pSlide = pptx.addSlide();
    pSlide.background = { color: T.bg };

    const { type, content } = slide;

    if (type === "title") {
      buildTitle(pptx, pSlide, content);
    } else {
      addSlideChrome(pptx, pSlide, slide, slides.length);
      if (type === "kpi")          buildKPI(pptx, pSlide, content);
      else if (type === "table")   buildTable(pptx, pSlide, content);
      else if (type === "metrics") buildMetrics(pptx, pSlide, content);
      else if (type === "transactions") buildTransactions(pptx, pSlide, content);
      else if (type === "tenants") buildTenants(pptx, pSlide, content);
      else if (type === "expiry")  buildExpiry(pptx, pSlide, content);
      else if (type === "highlights") buildHighlights(pptx, pSlide, content);
    }
  }

  return pptx.write({ outputType: "nodebuffer" });
}

module.exports = { generatePPTXBuffer };
