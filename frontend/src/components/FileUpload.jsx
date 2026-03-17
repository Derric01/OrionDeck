import { useState, useRef } from "react";

export default function FileUpload({ onUpload, disabled }) {
  const [dragging, setDragging] = useState(false);
  const [fileName, setFileName] = useState(null);
  const inputRef = useRef(null);

  const handleFile = (file) => {
    if (!file) return;
    if (!file.name.match(/\.(xlsx|xls)$/i)) {
      alert("Please upload an Excel file (.xlsx or .xls)");
      return;
    }
    setFileName(file.name);
    onUpload(file);
  };

  const handleDrop = (e) => {
    e.preventDefault();
    setDragging(false);
    const file = e.dataTransfer.files[0];
    handleFile(file);
  };

  const handleChange = (e) => {
    const file = e.target.files[0];
    handleFile(file);
  };

  return (
    <div
      className={`file-upload ${dragging ? "dragging" : ""} ${disabled ? "disabled" : ""}`}
      onDragOver={(e) => { e.preventDefault(); if (!disabled) setDragging(true); }}
      onDragLeave={() => setDragging(false)}
      onDrop={disabled ? undefined : handleDrop}
      onClick={() => !disabled && inputRef.current?.click()}
    >
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls"
        onChange={handleChange}
        style={{ display: "none" }}
        disabled={disabled}
      />
      <div className="file-upload-inner">
        <div className="upload-icon">⊞</div>
        {fileName ? (
          <p className="upload-filename">{fileName}</p>
        ) : (
          <>
            <p className="upload-label">Drop Excel file here or click to browse</p>
            <p className="upload-hint">.xlsx or .xls — Q4 raw data</p>
          </>
        )}
      </div>
    </div>
  );
}
