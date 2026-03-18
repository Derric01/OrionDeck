import axios from "axios";

const BASE_URL = import.meta?.env?.VITE_API_BASE_URL || "";

export const api = axios.create({
  baseURL: BASE_URL,
  timeout: 60000,
});

export async function uploadExcel(file, onUploadProgress) {
  const formData = new FormData();
  formData.append("excel", file);
  const res = await api.post("/upload", formData, {
    headers: { "Content-Type": "multipart/form-data" },
    onUploadProgress,
  });
  return res.data;
}

export async function generateReport() {
  const res = await api.post("/generate-report");
  return res.data;
}

export async function sendChat(message, history = []) {
  const res = await api.post("/chat", { message, history });
  return res.data;
}

// Stream real-time thinking steps + final result via SSE.
export async function sendChatStream(message, history = [], { onThinkingStep, onFinal } = {}) {
  const url = `${BASE_URL}/chat/stream`;
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ message, history }),
  });

  if (!res.ok) {
    const txt = await res.text().catch(() => "");
    throw new Error(`Chat stream failed (${res.status}): ${txt || res.statusText}`);
  }

  const reader = res.body?.getReader();
  if (!reader) throw new Error("Missing response body reader for chat stream.");

  const decoder = new TextDecoder("utf-8");
  let buffer = "";
  let finalPayload = null;

  const dispatchEventBlock = (block) => {
    // Very small SSE parser (event + data only)
    const lines = block.split(/\r?\n/).filter(Boolean);
    let eventName = "";
    const dataLines = [];
    for (const line of lines) {
      if (line.startsWith("event:")) eventName = line.slice("event:".length).trim();
      else if (line.startsWith("data:")) dataLines.push(line.slice("data:".length).trim());
    }
    const dataStr = dataLines.join("\n");
    if (!eventName) return;
    if (eventName === "thinking_step") {
      const payload = JSON.parse(dataStr);
      onThinkingStep?.(payload);
      return;
    }
    if (eventName === "final") {
      finalPayload = JSON.parse(dataStr);
      onFinal?.(finalPayload);
      return;
    }
    if (eventName === "error") {
      const payload = JSON.parse(dataStr || "{}");
      throw new Error(payload?.error || "Unknown stream error");
    }
  };

  while (true) {
    const { value, done } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });

    const blocks = buffer.split(/\n\n/);
    buffer = blocks.pop() || "";
    for (const b of blocks) {
      if (!b.trim()) continue;
      dispatchEventBlock(b);
      if (finalPayload) return finalPayload;
    }
  }

  return finalPayload;
}

export async function getSlides() {
  const res = await api.get("/slides");
  return res.data;
}

export function getDownloadUrl() {
  return `${BASE_URL}/report/download`;
}
