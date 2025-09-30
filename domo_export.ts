import "dotenv/config";
import fs from "node:fs/promises";
import path from "node:path";
import { chromium, Download } from "playwright";
import * as XLSX from "xlsx";
import { sendWithGmailAPI } from "./mailer_gmail.ts";

// ---------- config ----------
const DOWNLOAD_DIR = path.resolve("exporter/downloads");
const EMAIL_SUBJECT = process.env.EMAIL_SUBJECT ?? "DOMO export";
const EMAIL_BODY = process.env.EMAIL_BODY ?? "Attached CSV exports.";
const EMAIL_USER = process.env.EMAIL_USER!;
const EMAIL_TO   = process.env.EMAIL_TO!;

// ---------- helpers ----------
function safeFilename(name: string, finalExt = ".xlsx") {
  // Use the browser-suggested filename, but sanitize:
  // - strip control chars
  // - remove slashes and illegal chars
  // - trim and collapse whitespace
  // - strip trailing dots (Windows-incompatible)
  // - ensure single final extension
  let base = name;

  // Remove path-like pieces and control chars
  base = base.replace(/[\\/:*?"<>|\u0000-\u001F]/g, " ");
  // Collapse whitespace
  base = base.replace(/\s+/g, " ").trim();
  // Strip trailing dots
  base = base.replace(/\.+$/g, "");
  // If it already ends with .xlsx/.xls, drop it so we control extension
  base = base.replace(/\.(xlsx?|csv)$/i, "");
  // Guard length
  if (base.length === 0) base = "export";
  if (base.length > 180) base = base.slice(0, 180);

  return base + finalExt;
}

async function ensureDir(p: string) {
  await fs.mkdir(p, { recursive: true });
}

async function saveDownload(download: Download, ext = ".xlsx") {
  const suggested = download.suggestedFilename();
  const safe = safeFilename(suggested, ext);
  const target = path.join(DOWNLOAD_DIR, safe);
  await ensureDir(path.dirname(target));
  await download.saveAs(target); // robust now that path is clean and exists
  return target;
}

function xlsxToCsvFiles(xlsxPath: string) {
  const wb = XLSX.readFile(xlsxPath);
  const outFiles: string[] = [];
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const csv = XLSX.utils.sheet_to_csv(ws);
    const base = path.basename(xlsxPath).replace(/\.(xlsx?)$/i, "");
    const safeSheet = sheetName.replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, " ").trim();
    const out = path.join(DOWNLOAD_DIR, `${base} - ${safeSheet}.csv`);
    fs.writeFile(out, csv);
    outFiles.push(out);
  }
  return outFiles;
}

async function emailCsvs(csvPaths: string[]) {
  if (csvPaths.length === 0) return;

  const attachments = csvPaths.map(p => ({ filename: path.basename(p), path: p }));
  await sendWithGmailAPI({
    from: EMAIL_USER,
    to: EMAIL_TO,
    subject: EMAIL_SUBJECT,
    text: EMAIL_BODY,
    attachments
  });
}

// ---------- main ----------
(async () => {
  await ensureDir(DOWNLOAD_DIR);

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({
    acceptDownloads: true
  });
  const page = await context.newPage();

  // --- Login to DOMO (use your existing flow/selectors here) ---
  // NOTE: This is intentionally generic; keep your working login + navigation.
  // The only CHANGE you need is the download handler block below.
  await page.goto(process.env.DOMO_BASE_URL!, { waitUntil: "domcontentloaded" });

  // Example (replace with your current steps):
  // await page.click('text=Sign in');
  // await page.fill('[name="username"]', process.env.DOMO_USERNAME!);
  // await page.fill('[name="password"]', process.env.DOMO_PASSWORD!);
  // await Promise.all([
  //   page.waitForNavigation(),
  //   page.click('button:has-text("Log in")')
  // ]);
  // Navigate to the card/report and trigger export…

  // --- Robust download capture: works for one or many downloads ---
  const capturedXlsx: string[] = [];
  // Wherever you click the “Export” button that triggers a download,
  // wrap it with waitForEvent('download'):
  //
  // Example for one export:
  // const [download] = await Promise.all([
  //   page.waitForEvent("download"),
  //   page.click('button:has-text("Export")')
  // ]);
  // capturedXlsx.push(await saveDownload(download, ".xlsx"));
  //
  // Example for multiple exports - call your export actions in a loop and
  // wait/save each download:
  //
  // for (const locator of exportButtons) {
  //   const [download] = await Promise.all([
  //     page.waitForEvent("download"),
  //     locator.click()
  //   ]);
  //   capturedXlsx.push(await saveDownload(download, ".xlsx"));
  // }

  // -----
  // If you already had working export code, keep it. Just replace your old
  // `download.saveAs("exporter/downloads/whatever.xlsx")` with:
  // `capturedXlsx.push(await saveDownload(download, ".xlsx"));`
  // -----

  // Convert each XLSX to CSV(s)
  const csvs: string[] = [];
  for (const xPath of capturedXlsx) {
    const out = xlsxToCsvFiles(xPath);
    csvs.push(...out);
  }

  // Send via Gmail API (OAuth)
  await emailCsvs(csvs);

  await browser.close();
})().catch((err) => {
  console.error("Worker failed");
  console.error(err?.message || err);
  process.exit(1);
});
