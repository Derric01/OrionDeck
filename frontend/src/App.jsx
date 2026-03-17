import { useState, useCallback } from "react";
import ChatWindow from "./components/ChatWindow";
import ChatInput from "./components/ChatInput";
import ThinkingPanel from "./components/ThinkingPanel";
import { uploadExcel, generateReport, sendChat } from "./api/client";
import "./App.css";

let msgIdCounter = 0;
const newId = () => `msg-${++msgIdCounter}-${Date.now()}`;
const timestamp = () =>
  new Date().toLocaleTimeString("en-US", { hour: "2-digit", minute: "2-digit" });

export default function App() {
  const [messages, setMessages] = useState([]);
  const [thinkingSteps, setThinkingSteps] = useState([]);
  const [isThinking, setIsThinking] = useState(false);
  const [panelCollapsed, setPanelCollapsed] = useState(false);
  const [chatState, setChatState] = useState("idle");
  const [isLoading, setIsLoading] = useState(false);
  const [reportSlides, setReportSlides] = useState(null);

  const addMessage = useCallback((role, content, extras = {}) => {
    const msg = { id: newId(), role, content, timestamp: timestamp(), ...extras };
    setMessages((prev) => [...prev, msg]);
    return msg.id;
  }, []);

  const replaceTyping = useCallback((tempId, content, extras = {}) => {
    setMessages((prev) =>
      prev.map((m) =>
        m.id === tempId
          ? { ...m, isTyping: false, content, timestamp: timestamp(), ...extras }
          : m
      )
    );
  }, []);

  const handleSend = async (text) => {
    if (isLoading) return;
    addMessage("user", text);

    // If idle and user asks for quarterly report
    if (
      chatState === "idle" &&
      (text.toLowerCase().includes("quarterly") ||
        text.toLowerCase().includes("generat") ||
        text.toLowerCase().includes("q4") ||
        text.toLowerCase().includes("report") ||
        text.toLowerCase().includes("portfolio"))
    ) {
      setChatState("awaiting_upload");
      setThinkingSteps([{ step: 1, label: "Parsed user intent: portfolio report generation", delay: 0 }]);
      setIsThinking(false);
      addMessage(
        "agent",
        "Sure, I can generate the Q4 2025 Portfolio Report.\n\nThe report uses the **Orion Q4 2025 template**; you only need to upload **Orion_Q4_2025_Raw_Data.xlsx** (Summary, Properties, Leases, Transactions sheets).\n\nClick the **⊞** button below or use Upload Excel to attach your file.",
        { thinking: [] }
      );
      return;
    }

    await handleChatMessage(text);
  };

  const handleChatMessage = async (text) => {
    setIsLoading(true);

    const typingId = newId();
    setMessages((prev) => [...prev, { id: typingId, role: "agent", isTyping: true }]);

    // Reset thinking; backend returns full list of steps
    setThinkingSteps([]);
    setIsThinking(true);

    try {
      const res = await sendChat(text);

      const thinkSteps = (res.thinking || []).map((label, i) => ({
        step: i + 1,
        label,
        delay: 0,
      }));
      setThinkingSteps(thinkSteps);
      setIsThinking(false);

      const updatedSlides = res.slides || null;
      if (updatedSlides) setReportSlides(updatedSlides);

      replaceTyping(typingId, res.message, {
        thinking: res.thinking || [],
        slides: updatedSlides,
        action: res.action || null,
      });
    } catch (err) {
      setIsThinking(false);
      replaceTyping(
        typingId,
        `An error occurred: ${err.message}. Ensure the backend is running on port 3001.`,
        { thinking: [] }
      );
    }

    setIsLoading(false);
  };

  const handleUpload = async (file) => {
    if (isLoading) return;
    setIsLoading(true);
    setIsThinking(true);

    addMessage("user", `Uploading: ${file.name}`);
    const typingId = newId();
    setMessages((prev) => [...prev, { id: typingId, role: "agent", isTyping: true }]);

    setThinkingSteps([
      { step: 1, label: "File received — validating format", delay: 0 },
      { step: 2, label: "Reading Excel workbook structure", delay: 400 },
    ]);

    try {
      const uploadRes = await uploadExcel(file);
      if (!uploadRes.success) throw new Error(uploadRes.error || "Upload failed");

      const uploadThinking = uploadRes.thinking || [];
      setThinkingSteps(uploadThinking.map((s) => ({ step: s.step, label: s.label, delay: s.delay })));

      replaceTyping(
        typingId,
        `Excel file **${uploadRes.filename}** received and parsed.\n\nProcessing ${uploadRes.parsedSummary?.totalLeases || "89+"} lease records across ${uploadRes.parsedSummary?.operatingProperties || "61"} properties...\n\nGenerating presentation...`,
        { thinking: uploadRes.thinking?.map((s) => s.label) }
      );

      await new Promise((r) => setTimeout(r, 1200));
      const reportRes = await generateReport();

      const reportThinking = (reportRes.thinking || []).map((label, i) => ({
        step: uploadThinking.length + i + 1,
        label,
        delay: uploadThinking.length * 300 + i * 300,
      }));

      setThinkingSteps((prev) => [...prev, ...reportThinking]);
      setIsThinking(false);

      const slides = reportRes.report_data?.slides || [];
      setReportSlides(slides);
      setChatState("ready");

      addMessage(
        "agent",
        `Your **Q4 2025 Portfolio Report** has been generated. Below are all 8 slides.\n\nYou can now:\n- Ask questions about the data\n- Modify any slide: *"Change slide 2 occupancy to 92%"*\n- Add notes: *"Add a note about tenant diversification"*\n- Download the PPTX using the button above the slides`,
        { thinking: reportRes.thinking, slides }
      );
    } catch (err) {
      console.error("Upload/generate error:", err);
      setIsThinking(false);
      replaceTyping(
        typingId,
        `Error processing the file: ${err.message}\n\nEnsure:\n- The file is a valid Excel (.xlsx)\n- The backend is running on port 3001`,
        { thinking: [] }
      );
    }

    setIsLoading(false);
  };

  return (
    <div className="app-root">
      <header className="app-header">
        <div className="header-left">
          <span className="logo-mark">
            <img
              src="/Screenshot 2026-01-24 151951.png"
              alt="Braind logo"
              style={{ height: 28, width: 28, objectFit: "contain" }}
            />
          </span>
          <div>
            <div className="app-name">Braind</div>
            <div className="app-tagline">Portfolio Reporting Engine</div>
          </div>
        </div>
        <div className="header-right">
          <span className="header-badge">Q4 2025</span>
          <span className="header-status">
            {chatState === "ready" ? "● Report Active" : "○ Awaiting Data"}
          </span>
        </div>
      </header>

      <div className="app-body">
        <div className="chat-panel">
          <ChatWindow messages={messages} />
          <ChatInput
            onSend={handleSend}
            onUpload={handleUpload}
            disabled={isLoading}
            awaitingUpload={chatState === "awaiting_upload"}
          />
        </div>

        <ThinkingPanel
          steps={thinkingSteps}
          isThinking={isThinking}
          collapsed={panelCollapsed}
          onToggle={() => setPanelCollapsed((v) => !v)}
        />
      </div>
    </div>
  );
}
