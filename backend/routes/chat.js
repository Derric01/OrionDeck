const express = require("express");
const { processMessage, processMessageStream, getPortfolioContext } = require("../utils/chatEngine");
const { getSlides } = require("../utils/slideContent");

const router = express.Router();

router.post("/", async (req, res) => {
  const { message, history } = req.body;

  if (!message || typeof message !== "string") {
    return res.status(400).json({ error: "message is required" });
  }

  try {
    const context = getPortfolioContext();
    const response = await processMessage(message, Array.isArray(history) ? history : []);
    const slides = response.action ? getSlides() : null;

    return res.json({
      message: response.message,
      thinking: response.thinking || [],
      action: response.action || null,
      slides,
      hasData: !!context,
    });
  } catch (err) {
    console.error("Chat error:", err);
    return res.status(500).json({ error: "Chat processing failed." });
  }
});

// ─── SSE streaming route for real-time "thinking" ─────────────────────────
router.post("/stream", async (req, res) => {
  const { message, history } = req.body || {};

  if (!message || typeof message !== "string") {
    return res.status(400).json({ error: "message is required" });
  }

  // SSE headers
  res.setHeader("Content-Type", "text/event-stream; charset=utf-8");
  res.setHeader("Cache-Control", "no-cache, no-transform");
  res.setHeader("Connection", "keep-alive");
  // For some proxies (nginx) you may need: res.flushHeaders()
  res.flushHeaders?.();

  let closed = false;
  res.on("close", () => {
    closed = true;
  });
  let finalSent = false;

  const send = (event, data) => {
    if (closed || res.writableEnded || res.destroyed) return;
    try {
      const payload = typeof data === "string" ? data : JSON.stringify(data);
      res.write(`event: ${event}\n`);
      // Keep data on one line to simplify client parsing
      res.write(`data: ${payload}\n\n`);
    } catch {
      // If the client disconnected mid-write, ignore.
    }
  };

  try {
    const context = getPortfolioContext();
    if (!context) {
      send("error", { error: "No uploaded Excel is loaded. Upload first." });
      return res.end();
    }

    let response = null;
    response = await processMessageStream(
      message,
      Array.isArray(history) ? history : [],
      {
        onThinkingStep: ({ step, label }) => send("thinking_step", { step, label }),
        onFinal: (finalRes) => {
          // finalRes already includes slides + action
          // Hard guard: only send slide state when an actual slide modification happened.
          if (!finalRes?.action) finalRes.slides = null;
          if (!finalSent) {
            finalSent = true;
            send("final", finalRes);
          }
        },
      }
    );

    // Safety: if onFinal didn't fire for some reason, send here.
    if (!finalSent && !closed && response) {
      finalSent = true;
      if (!response?.action) response.slides = null;
      send("final", response);
    }

    if (!closed && !res.writableEnded) res.end();
  } catch (err) {
    console.error("Chat stream error:", err);
    if (!closed && !finalSent) {
      send("error", { error: "Chat processing failed." });
      finalSent = true;
    }
    if (!closed && !res.writableEnded) res.end();
  }
});

module.exports = router;
