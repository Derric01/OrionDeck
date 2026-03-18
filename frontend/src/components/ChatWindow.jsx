import { useEffect, useRef } from "react";
import PresentationViewer from "./PresentationViewer";

function renderMessage(text) {
  if (!text) return null;

  // Convert markdown-like content to JSX
  const lines = text.split("\n");
  const result = [];
  let tableBuffer = [];
  let inTable = false;

  const flushTable = () => {
    if (tableBuffer.length < 2) {
      tableBuffer.forEach((l, i) => result.push(<div key={`t-${i}`}>{l}</div>));
      tableBuffer = [];
      inTable = false;
      return;
    }

    const rows = tableBuffer.filter((l) => l.startsWith("|") && !l.match(/^\|[-| ]+\|$/));
    const tableEl = (
      <div key={`table-${result.length}`} className="msg-table-wrap">
        <table className="msg-table">
          <thead>
            <tr>
              {rows[0]
                .split("|")
                .filter((c) => c.trim())
                .map((cell, i) => (
                  <th key={i}>{cell.trim()}</th>
                ))}
            </tr>
          </thead>
          <tbody>
            {rows.slice(1).map((row, ri) => (
              <tr key={ri}>
                {row
                  .split("|")
                  .filter((c) => c.trim())
                  .map((cell, ci) => (
                    <td key={ci}>{renderInline(cell.trim())}</td>
                  ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    );
    result.push(tableEl);
    tableBuffer = [];
    inTable = false;
  };

  lines.forEach((line, i) => {
    if (line.startsWith("|")) {
      inTable = true;
      tableBuffer.push(line);
      return;
    }

    if (inTable) flushTable();

    if (line.startsWith("**") && line.endsWith("**") && line.length > 4) {
      result.push(<div key={i} className="msg-bold-line">{line.slice(2, -2)}</div>);
    } else if (line.startsWith("- ")) {
      result.push(<div key={i} className="msg-list-item">• {renderInline(line.slice(2))}</div>);
    } else if (line.trim() === "") {
      result.push(<div key={i} className="msg-spacer" />);
    } else {
      result.push(<div key={i} className="msg-line">{renderInline(line)}</div>);
    }
  });

  if (inTable) flushTable();

  return result;
}

function renderInline(text) {
  // Bold: **text**
  const parts = text.split(/(\*\*[^*]+\*\*)/g);
  return parts.map((part, i) => {
    if (part.startsWith("**") && part.endsWith("**")) {
      return <strong key={i}>{part.slice(2, -2)}</strong>;
    }
    return part;
  });
}

function Message({ msg }) {
  const isAgent = msg.role === "agent";

  return (
    <div className={`message ${isAgent ? "agent-message" : "user-message"}`}>
      {isAgent && (
        <div className="agent-avatar">
          <img src="/Screenshot 2026-01-24 151951.png" alt="Braind logo" />
        </div>
      )}
      <div className="message-bubble">
        {msg.isTyping ? (
          <div className="typing-indicator">
            <span /><span /><span />
          </div>
        ) : (
          <>
            {msg.slides && <PresentationViewer slides={msg.slides} />}
            <div className="message-text">{renderMessage(msg.content)}</div>
            {msg.thinking && msg.thinking.length > 0 && (
              <div className="inline-thinking">
                <div className="inline-thinking-toggle">
                  {msg.thinking.length} reasoning steps
                </div>
              </div>
            )}
            {msg.timestamp && (
              <div className="message-time">{msg.timestamp}</div>
            )}
          </>
        )}
      </div>
    </div>
  );
}

export default function ChatWindow({ messages }) {
  const bottomRef = useRef(null);

  useEffect(() => {
    bottomRef.current?.scrollIntoView({ behavior: "smooth" });
  }, [messages]);

  return (
    <div className="chat-window">
      {messages.length === 0 && (
        <div className="chat-empty">
          <div className="empty-icon">◎</div>
          <h2>Braind Portfolio Engine</h2>
          <p>Your AI-powered financial reporting assistant.</p>
          <p className="empty-hint">Start by saying <em>"Generate a Q4 portfolio reporting for me"</em></p>
        </div>
      )}
      <div className="messages-list">
        {messages.map((msg) => (
          <Message key={msg.id} msg={msg} />
        ))}
        <div ref={bottomRef} />
      </div>
    </div>
  );
}
