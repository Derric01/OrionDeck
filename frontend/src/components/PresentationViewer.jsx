import { getDownloadUrl } from "../api/client";

export default function PresentationViewer({ slides }) {
  if (!slides || slides.length === 0) return null;

  return (
    <div className="presentation-viewer">
      <div className="presentation-header">
        <span className="pres-label">PRESENTATION</span>
        <a href={getDownloadUrl()} target="_blank" rel="noreferrer" className="download-btn">
          ↓ Download PPTX
        </a>
      </div>
      <div className="slides-list">
        {slides.map((slide) => (
          <SlideBlock key={slide.id} slide={slide} />
        ))}
      </div>
    </div>
  );
}

function SlideBlock({ slide }) {
  return (
    <div className="slide-block">
      <div className="slide-number">Slide {slide.id}</div>
      <div className="slide-content">
        <h3 className="slide-title">{slide.title}</h3>
        <SlideBody slide={slide} />
        {slide.content?.notes && slide.content.notes.length > 0 && (
          <div className="slide-notes">
            <div className="notes-label">NOTES</div>
            {slide.content.notes.map((note, i) => (
              <div key={i} className="note-item">— {note}</div>
            ))}
          </div>
        )}
      </div>
    </div>
  );
}

function SlideBody({ slide }) {
  const { type, content } = slide;
  if (!content) return null;

  // Slide 1 — Cover / Portfolio Snapshot (template layout: left = title + KPIs, right = Q4 highlights)
  if (type === "cover") {
    return (
      <div className="slide-cover">
        <div className="cover-left">
          <div className="cover-title">{content.title}</div>
          <div className="cover-report-title">{content.reportTitle}</div>
          <div className="cover-date-line">{content.dateLine}</div>
          <div className="cover-kpi-grid">
            {(content.kpis || []).map((kpi, i) => (
              <div key={i} className="cover-kpi-card">
                <div className="cover-kpi-value">{kpi.value}</div>
                <div className="cover-kpi-card-label">{kpi.cardLabel}</div>
                <div className="cover-kpi-sub">{kpi.subLabel}</div>
              </div>
            ))}
          </div>
        </div>
        <div className="cover-right">
          <div className="cover-highlights-title">{content.highlightsSectionTitle}</div>
          <div className="cover-highlights-list">
            {(content.highlights || []).map((h, i) => (
              <div key={i} className="cover-highlight-card">
                <div className="cover-highlight-main">{h.main}</div>
                <div className="cover-highlight-sub">{h.sub}</div>
              </div>
            ))}
          </div>
        </div>
      </div>
    );
  }

  // Slide 2 — Portfolio Highlights
  if (type === "portfolioHighlights") {
    const m = content.metrics || {};
    return (
      <div className="portfolio-highlights-body">
        <div className="ph-metrics">
          <div className="ph-metric"><strong>Operating Properties:</strong> {m.operatingProperties}</div>
          <div className="ph-metric"><strong>Total Leasable Area:</strong> {m.totalLeasableSF}</div>
          <div className="ph-metric"><strong>Annualised Base Rent:</strong> {m.abr}</div>
          <div className="ph-metric"><strong>Occupancy:</strong> {m.occupancy}</div>
          <div className="ph-metric"><strong>WALT:</strong> {m.walt}</div>
          <div className="ph-metric"><strong>Investment-Grade Tenancy:</strong> {m.igTenancy}</div>
        </div>
        <div className="metrics-sub-title">Top-10 Tenant Credit Profile</div>
        <div className="slide-table-wrap">
          <table className="slide-table">
            <thead>
              <tr>
                <th>Rank</th>
                <th>Tenant</th>
                <th>Credit Rating</th>
                <th>% ABR</th>
              </tr>
            </thead>
            <tbody>
              {(content.top10Tenants || []).map((t, i) => (
                <tr key={i}>
                  <td>{t.rank}</td>
                  <td>{t.name}</td>
                  <td><span className={`rating-badge ${getRatingClass(t.creditRating)}`}>{t.creditRating}</span></td>
                  <td><strong>{t.pctABR}</strong></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        <p className="combined-note">{content.combinedNote}</p>
      </div>
    );
  }

  // Slide 3 — Portfolio Composition
  if (type === "composition") {
    const ig = content.igSplit || {};
    const at = content.assetTable || {};
    return (
      <div className="composition-body">
        <div className="comp-section">
          <div className="metrics-sub-title">Investment-Grade Tenancy Split</div>
          <div className="ig-split">
            <span>Investment Grade: <strong>{ig.investmentGrade}</strong></span>
            <span>Non-Investment Grade: <strong>{ig.nonInvestmentGrade}</strong></span>
          </div>
        </div>
        <div className="comp-section">
          <div className="metrics-sub-title">Asset Type Distribution</div>
          <div className="slide-table-wrap">
            <table className="slide-table">
              <thead>
                <tr>
                  {(at.headers || []).map((h, i) => <th key={i}>{h}</th>)}
                </tr>
              </thead>
              <tbody>
                {(at.rows || []).map((row, i) => (
                  <tr key={i}>
                    {row.map((cell, j) => <td key={j}>{cell}</td>)}
                  </tr>
                ))}
                {at.totals && (
                  <tr className="totals-row">
                    {at.totals.map((cell, j) => <td key={j}><strong>{cell}</strong></td>)}
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
        <div className="comp-section two-cols">
          <div>
            <div className="metrics-sub-title">Industry Breakdown (% of ABR)</div>
            <ul className="breakdown-list">
              {(content.industryBreakdown || []).map((item, i) => (
                <li key={i}>{item.name} — <strong>{item.pct}</strong></li>
              ))}
            </ul>
          </div>
          <div>
            <div className="metrics-sub-title">Geographic Breakdown (% of ABR)</div>
            <ul className="breakdown-list">
              {(content.geographicBreakdown || []).map((item, i) => (
                <li key={i}>{item.state} — <strong>{item.pct}</strong></li>
              ))}
            </ul>
          </div>
        </div>
      </div>
    );
  }

  // Slide 4 — Asset-by-Asset Performance
  if (type === "assetPerformance") {
    const headers = content.headers || [];
    const rows = content.rows || [];
    return (
      <div className="slide-table-wrap">
        <table className="slide-table asset-perf-table">
          <thead>
            <tr>
              {headers.map((h, i) => <th key={i}>{h}</th>)}
            </tr>
          </thead>
          <tbody>
            {rows.map((row, i) => (
              <tr key={i}>
                <td>{row.property}</td>
                <td>{row.type}</td>
                <td>{row.sf}</td>
                <td>{row.occupancy}</td>
                <td>{row.abr}</td>
                <td>{row.walt}</td>
                <td>{row.leaseExpiry}</td>
                <td><span className={`rating-badge ${getRatingClass(row.credit)}`}>{row.credit}</span></td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
  }

  // Slide 5 — Lease Expiry Schedule
  if (type === "expiry") {
    const headers = content.headers || ["Year", "Leases", "SF Expiring", "ABR Expiring", "% Portfolio"];
    const schedule = content.schedule || [];
    const totalRow = content.totalRow;
    return (
      <div className="expiry-body">
        <div className="slide-table-wrap">
          <table className="slide-table">
            <thead>
              <tr>
                {headers.map((h, i) => <th key={i}>{h}</th>)}
              </tr>
            </thead>
            <tbody>
              {schedule.map((row, i) => (
                <tr key={i} className={row.year === "2031+" ? "highlight-row" : ""}>
                  <td><strong>{row.year}</strong></td>
                  <td>{row.leases}</td>
                  <td>{row.sfExpiring}</td>
                  <td>{row.abrExpiring}</td>
                  <td>
                    <div className="expiry-bar-wrap">
                      <div className="expiry-bar" style={{ width: row.pctPortfolio }} />
                      <span>{row.pctPortfolio}</span>
                    </div>
                  </td>
                </tr>
              ))}
              {totalRow && (
                <tr className="totals-row">
                  <td><strong>{totalRow.year}</strong></td>
                  <td>{totalRow.leases}</td>
                  <td>{totalRow.sfExpiring}</td>
                  <td>{totalRow.abrExpiring}</td>
                  <td>{totalRow.pctPortfolio}</td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
        {content.keyNote && <p className="key-note">{content.keyNote}</p>}
      </div>
    );
  }

  // Slide 6 — Q4 2025 Disposition Activity
  if (type === "dispositions") {
    const m = content.metrics || {};
    return (
      <div className="dispositions-body">
        <div className="ph-metrics">
          <div className="ph-metric"><strong>Properties sold in Q4:</strong> {m.q4PropertiesSold}</div>
          <div className="ph-metric"><strong>Q4 gross proceeds:</strong> {m.q4GrossProceeds}</div>
          <div className="ph-metric"><strong>FY 2025 properties sold:</strong> {m.fy2025Sold}</div>
          <div className="ph-metric"><strong>FY 2025 proceeds:</strong> {m.fy2025Proceeds}</div>
          <div className="ph-metric"><strong>Estimated carrying cost savings:</strong> {m.carryingCostSavings}</div>
          <div className="ph-metric"><strong>Capex avoided:</strong> {m.capexAvoided}</div>
        </div>
        <div className="metrics-sub-title">Transaction table</div>
        <div className="slide-table-wrap">
          <table className="slide-table">
            <thead>
              <tr>
                <th>Property</th>
                <th>Region</th>
                <th>Asset Type</th>
                <th>SF</th>
                <th>Price</th>
                <th>Occupancy</th>
                <th>Strategic Reason</th>
              </tr>
            </thead>
            <tbody>
              {(content.transactions || []).map((tx, i) => (
                <tr key={i}>
                  <td>{tx.property}</td>
                  <td>{tx.region}</td>
                  <td>{tx.assetType}</td>
                  <td>{tx.sf}</td>
                  <td>{tx.price}</td>
                  <td>{tx.occupancy}</td>
                  <td>{tx.strategicReason}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
        {content.fullYearNote && <p className="key-note">{content.fullYearNote}</p>}
      </div>
    );
  }

  // Slide 7 — Forward Outlook
  if (type === "outlook") {
    const lp = content.leasingPipeline || {};
    const dp = content.dispositionPipeline || {};
    const ca = content.capitalAllocation || {};
    return (
      <div className="outlook-body">
        <div className="outlook-section">
          <div className="metrics-sub-title">Leasing Pipeline</div>
          <p><strong>{lp.sfUnderNegotiation}</strong> under active negotiation across {lp.propertyCount} properties.</p>
          <p>Recent leasing activity: {(lp.recentActivity || []).map((a, i) => <span key={i}>{a}{i < (lp.recentActivity?.length - 1) ? "; " : ""}</span>)}</p>
        </div>
        <div className="outlook-section">
          <div className="metrics-sub-title">Disposition Pipeline</div>
          <p><strong>{dp.sfCompleted}</strong> additional property sales completed after quarter end. {dp.note}</p>
        </div>
        <div className="outlook-section">
          <div className="metrics-sub-title">Capital Allocation Strategy</div>
          <p>Dedicated Use assets currently represent <strong>{ca.dedicatedUsePct}</strong>. Strategic target is <strong>{ca.targetPct}</strong>. Future investment will prioritize {(ca.priorities || []).join("; ")}.</p>
        </div>
      </div>
    );
  }

  // Slide 8 — Data Pipeline Appendix
  if (type === "appendix") {
    const s1 = content.step1 || {};
    const s2 = content.step2 || {};
    const s3 = content.step3 || {};
    return (
      <div className="appendix-body">
        <div className="appendix-title">HOW THIS REPORT WAS GENERATED</div>
        <div className="appendix-subtitle">
          Braind ingested your raw lease and property data and produced every table, chart, and metric in this presentation automatically.
        </div>
        <div className="appendix-grid">
          <div className="appendix-card">
            <div className="appendix-card-top">
              <div className="appendix-step">1</div>
              <div className="appendix-card-title">{s1.title}</div>
            </div>
            <div className="appendix-big">{s1.bigNumber}</div>
            <div className="appendix-big-label">{s1.bigLabel}</div>
            <div className="appendix-detail">{s1.detail}</div>
          </div>

          <div className="appendix-card">
            <div className="appendix-card-top">
              <div className="appendix-step">2</div>
              <div className="appendix-card-title">{s2.title}</div>
            </div>
            <div className="appendix-big">{s2.bigNumber}</div>
            <div className="appendix-big-label">{s2.bigLabel}</div>
            <div className="appendix-detail">{s2.detail}</div>
          </div>

          <div className="appendix-card">
            <div className="appendix-card-top">
              <div className="appendix-step">3</div>
              <div className="appendix-card-title">{s3.title}</div>
            </div>
            <div className="appendix-big">{s3.bigNumber}</div>
            <div className="appendix-big-label">{s3.bigLabel}</div>
            <div className="appendix-detail">{s3.detail}</div>
          </div>
        </div>
        {content.footer && <div className="appendix-footer">{content.footer}</div>}
      </div>
    );
  }

  // Legacy / fallback
  if (type === "kpi") {
    return (
      <div className="kpi-grid">
        {(content.kpis || []).map((kpi, i) => (
          <div key={i} className="kpi-item">
            <div className="kpi-value">{kpi.value}{kpi.unit ? <span className="kpi-unit"> {kpi.unit}</span> : null}</div>
            <div className="kpi-label">{kpi.label}</div>
          </div>
        ))}
      </div>
    );
  }

  if (type === "table") {
    return (
      <div className="slide-table-wrap">
        <table className="slide-table">
          <thead>
            <tr>
              {(content.headers || []).map((h, i) => <th key={i}>{h}</th>)}
            </tr>
          </thead>
          <tbody>
            {(content.rows || []).map((row, i) => (
              <tr key={i}>
                {row.map((cell, j) => <td key={j}>{cell}</td>)}
              </tr>
            ))}
            {content.totals && (
              <tr className="totals-row">
                {content.totals.map((cell, j) => <td key={j}>{cell}</td>)}
              </tr>
            )}
          </tbody>
        </table>
      </div>
    );
  }

  return <pre style={{ fontSize: "11px", overflow: "auto" }}>{JSON.stringify(content, null, 2)}</pre>;
}

function getRatingClass(rating) {
  if (!rating) return "";
  const r = String(rating).toUpperCase();
  if (r.startsWith("AA") || r.startsWith("AAA")) return "rating-ig-high";
  if (r.startsWith("A")) return "rating-ig-mid";
  if (r.startsWith("BBB")) return "rating-ig-low";
  return "rating-sub";
}
