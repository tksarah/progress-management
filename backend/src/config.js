import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const dataDirectory = path.resolve(__dirname, "../data");
const settingsPath = path.join(dataDirectory, "settings.json");
const defaultExcelPath = path.join(dataDirectory, "progress.xlsx");

const defaultSettings = {
  excelFilePath: defaultExcelPath
};

export async function ensureDataDirectory() {
  await fs.mkdir(dataDirectory, { recursive: true });
}

export async function getSettings() {
  await ensureDataDirectory();

  try {
    const content = await fs.readFile(settingsPath, "utf-8");
    const parsed = JSON.parse(content);

    return {
      ...defaultSettings,
      ...parsed
    };
  } catch (error) {
    if (error.code !== "ENOENT") {
      throw error;
    }

    await saveSettings(defaultSettings);
    return defaultSettings;
  }
}

export async function saveSettings(nextSettings) {
  await ensureDataDirectory();
  await fs.writeFile(settingsPath, JSON.stringify(nextSettings, null, 2), "utf-8");
}

export function resolveExcelPath(inputPath) {
  return path.resolve(inputPath);
}
