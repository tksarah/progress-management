import express from "express";
import cors from "cors";
import { getSettings, resolveExcelPath, saveSettings } from "./config.js";
import {
  createProgress,
  ensureWorkbook,
  getSummary,
  listProgress,
  updateProgress
} from "./excelStore.js";

const app = express();
const port = process.env.PORT || 3001;

app.use(cors());
app.use(express.json());

app.get("/api/health", async (_request, response) => {
  const settings = await getSettings();
  response.json({ status: "ok", excelFilePath: settings.excelFilePath });
});

app.get("/api/settings", async (_request, response, next) => {
  try {
    const settings = await getSettings();
    await ensureWorkbook(settings.excelFilePath);
    response.json(settings);
  } catch (error) {
    next(error);
  }
});

app.put("/api/settings/excel-path", async (request, response, next) => {
  try {
    const inputPath = String(request.body.excelFilePath ?? "").trim();

    if (!inputPath) {
      response.status(400).json({ message: "excelFilePath は必須です。" });
      return;
    }

    const excelFilePath = resolveExcelPath(inputPath);
    await ensureWorkbook(excelFilePath);
    await saveSettings({ excelFilePath });
    response.json({ excelFilePath });
  } catch (error) {
    next(error);
  }
});

app.get("/api/progress", async (request, response, next) => {
  try {
    const settings = await getSettings();
    const items = await listProgress(settings.excelFilePath);
    const query = String(request.query.q ?? "").trim().toLowerCase();
    const status = String(request.query.status ?? "").trim();

    const filteredItems = items.filter((item) => {
      const matchesStatus = status ? item.status === status : true;
      const matchesQuery = query
        ? [
            item.kpiNumber,
            item.assignee,
            item.customer,
            item.content,
            item.nextAction
          ].join(" ").toLowerCase().includes(query)
        : true;

      return matchesStatus && matchesQuery;
    });

    const summary = await getSummary(settings.excelFilePath);

    response.json({
      items: filteredItems,
      total: filteredItems.length,
      summary,
      excelFilePath: settings.excelFilePath
    });
  } catch (error) {
    next(error);
  }
});

app.post("/api/progress", async (request, response, next) => {
  try {
    const settings = await getSettings();
    const created = await createProgress(settings.excelFilePath, request.body);
    response.status(201).json(created);
  } catch (error) {
    next(error);
  }
});

app.put("/api/progress/:id", async (request, response, next) => {
  try {
    const settings = await getSettings();
    const updated = await updateProgress(settings.excelFilePath, request.params.id, request.body);
    response.json(updated);
  } catch (error) {
    next(error);
  }
});

app.use((error, _request, response, _next) => {
  if (error.code === "VERSION_CONFLICT") {
    response.status(409).json({ message: error.message });
    return;
  }

  response.status(500).json({ message: error.message || "サーバーエラーが発生しました。" });
});

async function start() {
  const settings = await getSettings();
  await ensureWorkbook(settings.excelFilePath);

  app.listen(port, () => {
    console.log(`Backend listening on http://localhost:${port}`);
  });
}

start().catch((error) => {
  console.error(error);
  process.exit(1);
});
