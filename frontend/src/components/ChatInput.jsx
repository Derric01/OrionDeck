import { useState, useRef } from "react";
import FileUpload from "./FileUpload";

export default function ChatInput({ onSend, onUpload, disabled, awaitingUpload }) {
  const [message, setMessage] = useState("");
  const [showUpload, setShowUpload] = useState(false);
  const textareaRef = useRef(null);

  const handleSend = () => {
    const trimmed = message.trim();
    if (!trimmed || disabled) return;
    onSend(trimmed);
    setMessage("");
    if (textareaRef.current) {
      textareaRef.current.style.height = "auto";
    }
  };

  const handleKeyDown = (e) => {
    if (e.key === "Enter" && !e.shiftKey) {
      e.preventDefault();
      handleSend();
    }
  };

  const handleTextareaChange = (e) => {
    setMessage(e.target.value);
    // Auto-resize
    const ta = textareaRef.current;
    if (ta) {
      ta.style.height = "auto";
      ta.style.height = Math.min(ta.scrollHeight, 160) + "px";
    }
  };

  const handleUpload = (file) => {
    setShowUpload(false);
    onUpload(file);
  };

  return (
    <div className="chat-input-area">
      {showUpload && (
        <div className="upload-overlay">
          <FileUpload onUpload={handleUpload} disabled={disabled} />
          <button className="close-upload" onClick={() => setShowUpload(false)}>✕ Cancel</button>
        </div>
      )}

      {awaitingUpload && !showUpload && (
        <div className="awaiting-upload-hint">
          <span>Agent is waiting for your Excel file.</span>
          <button className="upload-trigger-btn" onClick={() => setShowUpload(true)}>
            Upload Excel
          </button>
        </div>
      )}

      <div className="input-row">
        <button
          className="attach-btn"
          onClick={() => setShowUpload((v) => !v)}
          title="Upload Excel file"
          disabled={disabled}
        >
          ⊞
        </button>
        <textarea
          ref={textareaRef}
          className="chat-textarea"
          value={message}
          onChange={handleTextareaChange}
          onKeyDown={handleKeyDown}
          placeholder="Ask about the portfolio or modify the report..."
          disabled={disabled}
          rows={1}
        />
        <button
          className={`send-btn ${disabled ? "disabled" : ""}`}
          onClick={handleSend}
          disabled={disabled || !message.trim()}
        >
          →
        </button>
      </div>
      <div className="input-hint">Press Enter to send · Shift+Enter for new line · ⊞ to upload Excel</div>
    </div>
  );
}
