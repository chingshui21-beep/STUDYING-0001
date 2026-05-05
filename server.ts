import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import axios from "axios";

async function startServer() {
  const app = express();
  const PORT = 3000;

  app.use(express.json());

  // API Route to fetch Google Sheet content
  app.get("/api/fetch-gsheet", async (req, res) => {
    const { url } = req.query;
    if (!url || typeof url !== "string") {
      return res.status(400).json({ error: "Missing URL parameter" });
    }

    try {
      // Regex to extract file ID from Google Sheet URL
      const match = url.match(/\/d\/([\w-]+)/);
      const spreadsheetId = match ? match[1] : url;

      // Construct export URL (forcing CSV format)
      const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?format=csv`;
      
      const response = await axios.get(exportUrl);
      res.send(response.data);
    } catch (error: any) {
      console.error("Error fetching Google Sheet:", error.message);
      res.status(500).json({ error: "Failed to fetch Google Sheet. Please make sure the sheet is public (Anyone with the link can view)." });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
