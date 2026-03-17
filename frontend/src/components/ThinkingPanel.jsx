import { useEffect, useRef } from "react";

export default function ThinkingPanel({ steps, isThinking, collapsed, onToggle }) {
  const listRef = useRef(null);

  // Auto-scroll to bottom as new steps arrive
  useEffect(() => {
    if (listRef.current) {
      listRef.current.scrollTop = listRef.current.scrollHeight;
    }
  }, [steps]);

  return (
    <div className={`thinking-panel ${collapsed ? "collapsed" : ""}`}>
      <div className="thinking-panel-header" onClick={onToggle}>
        <div className="thinking-panel-title">
          <span className="thinking-icon">{isThinking ? "◉" : "○"}</span>
          <span>Agent Thinking</span>
          {isThinking && <span className="thinking-badge">Live</span>}
          {!isThinking && steps.length > 0 && (
            <span className="thinking-count">{steps.length} steps</span>
          )}
        </div>
        <span className="collapse-btn">{collapsed ? "▶" : "◀"}</span>
      </div>

      {!collapsed && (
        <div className="thinking-panel-body" ref={listRef}>
          {steps.length === 0 && !isThinking && (
            <div className="thinking-empty">
              <span>Waiting for input...</span>
              <p>Upload an Excel file or ask a question to see the agent's reasoning process in real-time.</p>
            </div>
          )}

          <ul className="thinking-steps">
            {steps.map((step, i) => (
              <li key={`${step.step}-${i}`} className="thinking-step thinking-step-enter">
                <span className="step-num">{step.step || i + 1}</span>
                <span className="step-label">{step.label}</span>
                <span className="step-check">✓</span>
              </li>
            ))}
          </ul>

          {isThinking && (
            <div className="thinking-spinner">
              <span className="dot-1">·</span>
              <span className="dot-2">·</span>
              <span className="dot-3">·</span>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
