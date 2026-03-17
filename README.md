# Braind Portfolio Reporting Engine

AI-powered chatbot that accepts Excel uploads, generates portfolio reports, and renders them inline in a chat UI.

## Quick Start

### 1. Install dependencies
```bash
# Root (backend deps)
npm install

# Frontend
cd frontend && npm install && cd ..
```

### 2. Start both servers

**Terminal 1 — Backend (port 3001)**
```bash
node backend/server.js
```

**Terminal 2 — Frontend (port 5173)**
```bash
cd frontend && npm run dev
```

Open http://localhost:5173

## Demo Flow

1. Type: `"Generate a Q4 portfolio reporting for me"`
2. Agent prompts for upload
3. Upload `backend/Orion_Q4_2025_Raw_Data.xlsx`
4. Watch the Agent Thinking panel live-update
5. 8 slides appear inline in the chat
6. Ask questions: `"What is the occupancy rate?"`
7. Modify: `"Change slide 2 occupancy to 92%"`
8. Download PPTX from the button above the slides

## Chat Commands

| Query | Response |
|---|---|
| "What is the occupancy rate?" | Portfolio + by asset type breakdown |
| "Which tenant contributes the most rent?" | Top tenant by ABR |
| "What is the WALT?" | Portfolio WALT + by asset type |
| "Show me the Q4 transactions" | Disposition activity table |
| "Change slide 2 occupancy to 92%" | Live update slide content |
| "Add a note about tenant diversification" | Appends note to slide |
| "What is the ABR?" | Full ABR breakdown |

## Architecture

```
OrionDeck/
├── backend/
│   ├── server.js              Express API server
│   ├── routes/
│   │   ├── upload.js          POST /upload — multer + Excel parse
│   │   ├── report.js          POST /generate-report, GET /slides, GET /report/download
│   │   └── chat.js            POST /chat
│   └── utils/
│       ├── slideContent.js    Slide data (8 slides, mutable state)
│       ├── parseExcel.js      xlsx parser
│       └── chatEngine.js      Rule-based Q&A + slide modification engine
└── frontend/
    └── src/
        ├── App.jsx             State orchestrator
        ├── api/client.js       Axios API layer
        └── components/
            ├── ChatWindow.jsx  Message renderer with markdown tables
            ├── ChatInput.jsx   Input + file upload trigger
            ├── ThinkingPanel.jsx Animated agent reasoning sidebar
            ├── FileUpload.jsx  Drag & drop Excel uploader
            └── PresentationViewer.jsx 8-slide inline renderer
```
