const express = require("express");
const { processMessage, getPortfolioContext } = require("../utils/chatEngine");
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

module.exports = router;
