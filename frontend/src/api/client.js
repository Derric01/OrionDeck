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

export async function sendChat(message) {
  const res = await api.post("/chat", { message });
  return res.data;
}

export async function getSlides() {
  const res = await api.get("/slides");
  return res.data;
}

export function getDownloadUrl() {
  return `${BASE_URL}/report/download`;
}
