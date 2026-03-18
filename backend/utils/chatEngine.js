const OpenAI = require("openai");
const { getSlides, updateSlide, addNote } = require("./slideContent");
const { getWalkthroughContext } = require("./walkthrough");

let openai = null;
function getOpenAIClient() {
  if (openai) return openai;
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    const err = new Error("Missing OPENAI_API_KEY");
    err.code = "MISSING_OPENAI_API_KEY";
    throw err;
  }
  openai = new OpenAI({ apiKey });
  return openai;
}

let portfolioContext = null;
function setPortfolioContext(ctx) {
  portfolioContext = ctx;
}
function getPortfolioContext() {
  return portfolioContext;
}

// ─── System prompt for non-streaming completion ──────────────────────────────

function buildSystemPrompt(slides, portfolioContext) {
  const walkthrough = getWalkthroughContext();
  const walkthroughBlock = walkthrough
    ? (() => {
        // Keep this compact to control token burn while still grounding the model.
        const cap = (grid, maxRows = 60) => (Array.isArray(grid) ? grid.slice(0, maxRows) : null);
        const pick = (name) => cap(walkthrough.sheets?.[name]);
        const summary = pick("Portfolio_Summary");
        const topTenants = pick("STEP5_Top_Tenants");
        const assetType = pick("STEP6_Asset_Type");
        const industry = pick("STEP7_Industry");
        const geography = pick("STEP8_Geography");
        const leaseExpiry = pick("STEP10_Lease_Expiry");
        const dispositions = pick("STEP11_Dispositions");

        return `

QUARTERLY PRESENTATION WALKTHROUGH (read-only reference model used to verify calculations and mapping):
- This workbook represents the institutional calculation model and slide mapping steps.
- Use it to explain *how* each KPI/table is derived and to sanity-check mapping decisions.
- Do NOT expose file names/paths; only reference sheet names and columns generically.
Available sheets: ${JSON.stringify(walkthrough.sheetNames)}
Portfolio_Summary (preview): ${JSON.stringify(summary)}
STEP5_Top_Tenants (preview): ${JSON.stringify(topTenants)}
STEP6_Asset_Type (preview): ${JSON.stringify(assetType)}
STEP7_Industry (preview): ${JSON.stringify(industry)}
STEP8_Geography (preview): ${JSON.stringify(geography)}
STEP10_Lease_Expiry (preview): ${JSON.stringify(leaseExpiry)}
STEP11_Dispositions (preview): ${JSON.stringify(dispositions)}`;
      })()
    : "";

  const portfolioBlock = portfolioContext
    ? (() => {
        const summary = JSON.stringify(portfolioContext.summary || {}, null, 2);
        const props = (portfolioContext.properties || []).slice(0, 80);
        const leas = (portfolioContext.leases || []).slice(0, 100);
        const txs = portfolioContext.transactions || [];
        return `

PORTFOLIO DATA FROM UPLOADED EXCEL (this workbook is the single source of truth for all KPIs and slide numbers — never expose its file name or any backend file paths to the user):
Summary (KPIs computed from the lease roll): ${summary}
Properties (sample, ${(portfolioContext.properties || []).length} total): ${JSON.stringify(props, null, 2)}
Leases (sample, ${(portfolioContext.leases || []).length} total): ${JSON.stringify(leas, null, 2)}
Transactions: ${JSON.stringify(txs, null, 2)}
You should behave like an institutional-grade analytics engine with full knowledge of:
- The raw data workbook that contains properties, leases, and transactions.
- The internal calculation walkthrough workbook that shows how each KPI is derived step-by-step.
- The final presentation deck that the user sees in the viewer.
Never reveal internal workbook or deck file names, paths, environment variables, keys, or any implementation details about the backend. Describe everything in generic, user-facing terms only.
When users ask how a number was calculated, explain it using the same logic as the walkthrough Excel:
- Occupancy % = total leased SF / total rentable SF (from leases and properties).
- ABR and ABR / SF roll up from individual lease ‘Annual Base Rent’ and area.
- Investment-grade % of ABR comes from the ‘Inv. Grade?’ / credit-rating fields on each lease.
- WALT is a weighted average lease term by ABR, using remaining term or expiry dates.
Always give clear, step-by-step explanations and tie answers back to specific sheets, columns, and slide elements.`;
      })()
    : "";

  return `You are an expert institutional real-estate financial analyst AI for OrionDeck — an AI-powered portfolio reporting platform.
You are analyzing the Q4 2025 Portfolio Report for Orion Properties Inc., a diversified REIT.
${portfolioBlock}
${walkthroughBlock}

CURRENT LIVE SLIDE DATA (JSON — use for modifications and slide-specific questions):
${JSON.stringify(slides, null, 2)}

REPORT VIEWER = OFFICIAL SLIDE DECK:
- The on-screen report viewer is the live, data-driven version of the eight-slide Q4 2025 portfolio report deck.
- Treat these 8 slides as identical in structure and numbers to the user-facing slide deck — if the user mentions "the deck", "the PowerPoint", specific slides/pages, or "in the chat report", they are referring to this same slide JSON.
- Every KPI card, table cell, and percentage in the viewer is generated from the uploaded Excel using the same logic as the internal calculation model; never assume there is a discrepancy between the viewer and the deck unless the user explicitly says there is one.

YOUR CAPABILITIES:
1. Answer ANY question about the portfolio with precision — cite numbers, compute derived metrics, and explain trends and causal drivers.
2. Explain, in natural conversational language, how the report is generated end-to-end:
   - From raw Excel rows → derived KPIs → slide content → final PPTX.
   - Reference specific sheets, columns, and slide numbers when helpful.
3. Use the provided conversation history to interpret pronouns and references (e.g., "earlier", "slide 6", "the previous calculation"). Do not contradict the live slide JSON.
4. Modify slide content exactly as the user requests (change values, update text, update KPIs).
5. Add notes to any slide.
6. Provide strategic insight and commentary as a senior REIT / portfolio analyst (risk, opportunities, credit quality, lease rollover, capital allocation, disposition strategy).

MODIFICATION RULES:
- When a user says "change X to Y" or "update X to Y" or "set X to Y" in slide N, extract:
  - slideId: the slide number (1-8)
  - field: the exact text/value to find and replace
  - value: the new value to set
- For "add a note", extract slideId and noteText.
- Be smart about partial matches — "89.2%" is in slide 8 as "89.2% Investment-Grade ABR — institutional-quality credit base", so field = "89.2%" will do a substring replacement.

RESPONSE FORMAT — you MUST respond with valid JSON only, no extra text:
{
  "thinking": [
    "Concise internal reasoning step 1",
    "Concise internal reasoning step 2",
    "...(3-6 steps total)"
  ],
  "message": "Your response to the user in markdown. Be precise, data-driven, and professional.",
  "modification": null
}

OR if a modification is needed:
{
  "thinking": ["..."],
  "message": "Confirmation of what was updated.",
  "modification": {
    "type": "change",
    "slideId": 8,
    "field": "89.2%",
    "value": "90%"
  }
}

OR for adding a note:
{
  "thinking": ["..."],
  "message": "Note added.",
  "modification": {
    "type": "note",
    "slideId": 2,
    "noteText": "The note content here"
  }
}

-IMPORTANT:
- thinking steps should be SHORT (5-12 words each), like internal monologue.
- message formatting (user-friendly):
  - Do NOT use Markdown headings (no lines starting with #, ##, ###).
  - Do NOT use numbered lists.
  - Use 2–4 short paragraphs with bold labels.
  - If you describe transactions, show at most 3 rows in a compact markdown table with columns: Property, SF, Price, Occ%, Strategic Reason.
- Focus on the specific question the user asked (e.g., \"Explain slide 4\"), not every possible metric at once.
- Never hallucinate numbers — only use what is in the slide data and portfolio data above, which already embed the full Excel and walkthrough logic. If the user asks for something that is not directly derivable, say so and offer the closest reliable proxy or explain what additional data would be needed.
- For modifications: always include the modification object even if you mention it in the message.`;
}

// ─── Value replacement (exact + substring) ────────────────────────────────────

function replaceValueInObject(node, oldValue, newValue) {
  let changed = false;

  const normalise = (v) =>
    String(v).trim().replace(/\s+(?=%)/g, "").replace(/\s+(?=[A-Za-z])/g, " ");

  const normOld = normalise(oldValue);

  const applyReplacement = (str) => {
    const normStr = normalise(str);
    if (normStr === normOld) return { result: newValue, matched: true };
    if (normStr.includes(normOld)) {
      return { result: str.replace(oldValue, newValue), matched: true };
    }
    return { matched: false };
  };

  const visit = (obj) => {
    if (Array.isArray(obj)) {
      obj.forEach((item, idx) => {
        if (typeof item === "string") {
          const { result, matched } = applyReplacement(item);
          if (matched) { obj[idx] = result; changed = true; }
        } else if (typeof item === "object" && item !== null) {
          visit(item);
        }
      });
      return;
    }
    if (typeof obj === "object" && obj !== null) {
      Object.keys(obj).forEach((key) => {
        const val = obj[key];
        if (typeof val === "string") {
          const { result, matched } = applyReplacement(val);
          if (matched) { obj[key] = result; changed = true; }
        } else if (typeof val === "object" && val !== null) {
          visit(val);
        }
      });
    }
  };

  visit(node);
  return { changed };
}

// ─── Apply modification to live slide state ────────────────────────────────────

function applyModification(mod) {
  const slides = getSlides();
  if (!mod) return null;

  if (mod.type === "note") {
    addNote(mod.slideId, mod.noteText);
    return getSlides();
  }

  if (mod.type === "change") {
    // Try structured KPI update first
    const result = updateSlide(mod.slideId, mod.field, mod.value);
    if (!result.success) {
      // Fallback: substring-aware replacement on the target slide
      const slide = slides.find((s) => s.id === parseInt(mod.slideId));
      if (slide) {
        const r = replaceValueInObject(slide.content, mod.field, mod.value);
        // If not found on target slide, scan all slides
        if (!r.changed) {
          slides.forEach((s) => replaceValueInObject(s.content, mod.field, mod.value));
        }
      }
    }
    return getSlides();
  }

  return null;
}

// ─── Main non-streaming entry point ───────────────────────────────────────────

async function processMessage(userMessage, historyMessages = []) {
  const slides = getSlides();
  const context = getPortfolioContext();
  const systemPrompt = buildSystemPrompt(slides, context);

  try {
    const client = getOpenAIClient();
    const historyForModel = (historyMessages || [])
      .filter((m) => m && typeof m.content === "string" && m.content.trim())
      .slice(-12)
      .map((m) => ({
        role: m.role === "assistant" ? "assistant" : "user",
        content: m.content,
      }));

    const completion = await client.chat.completions.create({
      model: "gpt-4o",
      messages: [
        { role: "system", content: systemPrompt },
        ...historyForModel,
        { role: "user", content: userMessage },
      ],
      response_format: { type: "json_object" },
      temperature: 0.25,
      max_tokens: 2048,
    });

    const raw = completion.choices[0]?.message?.content || "{}";
    let parsed;
    try {
      parsed = JSON.parse(raw);
    } catch {
      parsed = {
        thinking: ["Model returned non-JSON response"],
        message: raw || "I processed your request.",
        modification: null,
      };
    }

    const updatedSlides = applyModification(parsed.modification || null);

    return {
      message: parsed.message || "Done.",
      thinking: parsed.thinking || [],
      slides: updatedSlides,
      action: parsed.modification
        ? { type: "slide_updated", modification: parsed.modification }
        : null,
    };
  } catch (err) {
    console.error("OpenAI error:", err.message);
    return {
      message:
        err.code === "MISSING_OPENAI_API_KEY"
          ? "Server is missing `OPENAI_API_KEY`. Set it in environment variables (recommended) or a local `.env` file and restart."
          : `I encountered an error: ${err.message}. Please check your API key and try again.`,
      thinking: ["Error occurred during AI processing"],
      slides: null,
      action: null,
    };
  }
}

module.exports = { processMessage, setPortfolioContext, getPortfolioContext };
