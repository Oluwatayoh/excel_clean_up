"use client";

import { useState } from "react";

export default function Home() {
  const [file, setFile] = useState<File | null>(null);
  const [message, setMessage] = useState<string>("");

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files[0]) {
      setFile(e.target.files[0]);
    }
  };

  const handleSubmit = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();

    if (!file) {
      setMessage("Please select a file.");
      return;
    }

    const formData = new FormData();
    formData.append("excelFile", file);

    try {
      const response = await fetch("/api/upload", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const result = await response.json();
        setMessage(result.message || "An error occurred.");
        return;
      }

      // Trigger file download
      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "output.xlsx";
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);

      setMessage("File processed and downloaded successfully.");
    } catch (error) {
      console.error("Error:", error);
      setMessage("An error occurred while uploading the file.");
    }
  };

  return (
    <div style={{ padding: "20px" }}>
      <h1>Upload Excel and Download Processed File</h1>
      <form onSubmit={handleSubmit}>
        <input type="file" accept=".xlsx" onChange={handleFileChange} />
        <button type="submit">Process and Download</button>
      </form>
      {message && <p>{message}</p>}
    </div>
  );
}



