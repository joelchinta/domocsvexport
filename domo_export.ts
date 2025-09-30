import "dotenv/config";
import fs from "node:fs/promises";
import path from "node:path";
import { chromium } from "playwright";
import type { BrowserContext, Download, Page } from "playwright";
import * as XLSX from "xlsx";
import { sendWithGmailAPI } from "./mailer_gmail.ts";

// ---------- config ----------
const DOWNLOAD_DIR = path.resolve("exporter/downloads");
const DOMO_BASE_URL = process.env.DOMO_BASE_URL!;
const EMAIL_SUBJECT = process.env.EMAIL_SUBJECT ?? "DOMO export";
const EMAIL_BODY = process.env.EMAIL_BODY ?? "Attached CSV exports.";
const EMAIL_USER = process.env.EMAIL_USER!;
const EMAIL_TO   = process.env.EMAIL_TO!;
// Optional tuning knobs:
// - DOMO_EXPORT_URLS: comma/newline/semicolon separated list of download endpoints.
// - DOMO_CARD_IDS: comma/newline/semicolon separated DOMO card IDs (auto-builds export URLs).
// - EXPORT_TIMEOUT_MS: override per-download timeout (defaults to 3 minutes).
const EXPORT_TIMEOUT_MS = Number(process.env.EXPORT_TIMEOUT_MS ?? 3 * 60_000);

type ExportTarget = {
  label: string;
  url: string;
};

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

function parseList(input: string | undefined) {
  if (!input) return [] as string[];
  return input
    .split(/[;,\n]/)
    .map((item) => item.trim())
    .filter(Boolean);
}

function buildExportTargets(baseUrl: string) {
  const targets: ExportTarget[] = [];

  const explicitUrls = parseList(process.env.DOMO_EXPORT_URLS);
  for (const url of explicitUrls) {
    const resolved = new URL(url, baseUrl).toString();
    targets.push({ label: url, url: resolved });
  }

  const cardIds = parseList(process.env.DOMO_CARD_IDS);
  for (const id of cardIds) {
    const trimmed = id.replace(/[^a-z0-9-]/gi, "");
    if (!trimmed) continue;
    const resolved = new URL(`/card/export/${trimmed}?format=xlsx`, baseUrl).toString();
    targets.push({ label: `card:${trimmed}`, url: resolved });
  }

  if (!targets.length) {
    throw new Error("No export targets configured. Set DOMO_EXPORT_URLS or DOMO_CARD_IDS.");
  }

  return targets;
}

async function loginToDomo(page: Page) {
  const loginUrl = new URL("/auth/index", DOMO_BASE_URL).toString();
  console.log(`Navigating to login page: ${loginUrl}`);
  await page.goto(loginUrl, { waitUntil: "domcontentloaded" });

  const username = process.env.DOMO_USERNAME!;
  const password = process.env.DOMO_PASSWORD!;

  const usernameSelector = 'input[name="username"], input[type="email"], input#username, input#email, input#userId';
  const passwordSelector = 'input[name="password"], input[type="password"], input#password, input#pass';
  const submitSelector = 'button[type="submit"], button[data-qa="login-button"], button:has-text("Log in"), button:has-text("Sign in")';

  const usernameCandidates = page.locator(usernameSelector);
  if ((await usernameCandidates.count()) === 0) {
    console.log("Login form not detected. Assuming SSO/auth already satisfied.");
    await ensureLandingPage(page);
    return;
  }

  const usernameInput = usernameCandidates.first();

  console.log("Filling username and password");
  await usernameInput.fill(username);

  const passwordCandidates = page.locator(passwordSelector);
  if ((await passwordCandidates.count()) === 0) {
    throw new Error("Password field not found on DOMO login page.");
  }
  const passwordInput = passwordCandidates.first();
  await passwordInput.fill(password);

  console.log("Submitting login form");
  const submitButton = page.locator(submitSelector).first();
  if ((await submitButton.count()) === 0) {
    throw new Error("Login submit button not found on DOMO login page.");
  }
  await Promise.all([
    page.waitForNavigation({ waitUntil: "networkidle" }),
    submitButton.click()
  ]);

  console.log("Login complete");
  await ensureLandingPage(page);
}

async function ensureLandingPage(page: Page) {
  try {
    await page.goto(DOMO_BASE_URL, { waitUntil: "networkidle" });
  } catch (err) {
    const message = (err as Error)?.message ?? String(err);
    if (!/ERR_ABORTED/.test(message)) {
      throw err;
    }
  }
}

async function triggerDownload(context: BrowserContext, target: ExportTarget) {
  console.log(`Triggering export for ${target.label}`);
  const page = await context.newPage();
  const downloadPromise = page.waitForEvent("download", { timeout: EXPORT_TIMEOUT_MS });

  try {
    await page.goto(target.url, { waitUntil: "domcontentloaded" });
  } catch (err) {
    const message = (err as Error)?.message ?? String(err);
    // Downloads often abort navigation with net::ERR_ABORTED – treat as expected
    if (!/ERR_ABORTED/.test(message)) {
      throw err;
    }
  }

  const download = await downloadPromise;
  await page.close();
  return download;
}

async function xlsxToCsvFiles(xlsxPath: string) {
  const wb = XLSX.readFile(xlsxPath);
  const outFiles: string[] = [];
  for (const sheetName of wb.SheetNames) {
    const ws = wb.Sheets[sheetName];
    const csv = XLSX.utils.sheet_to_csv(ws);
    const base = path.basename(xlsxPath).replace(/\.(xlsx?)$/i, "");
    const safeSheet = sheetName.replace(/[\\/:*?"<>|]/g, "_").replace(/\s+/g, " ").trim();
    const out = path.join(DOWNLOAD_DIR, `${base} - ${safeSheet}.csv`);
    await fs.writeFile(out, csv);
    outFiles.push(out);
  }
  return outFiles;
}

async function emailCsvs(csvPaths: string[]) {
  const attachments = csvPaths.map(p => ({ filename: path.basename(p), path: p }));
  const body = attachments.length
    ? EMAIL_BODY
    : `${EMAIL_BODY}\n\n(No CSV files were generated in this run.)`;
  await sendWithGmailAPI({
    from: EMAIL_USER,
    to: EMAIL_TO,
    subject: EMAIL_SUBJECT,
    text: body,
    attachments
  });
}

// ---------- main ----------
(async () => {
  await ensureDir(DOWNLOAD_DIR);

  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ acceptDownloads: true });
  const page = await context.newPage();

  await loginToDomo(page);

  await page.close();

  const targets = buildExportTargets(DOMO_BASE_URL);
  const downloadPaths: string[] = [];

  for (const target of targets) {
    const download = await triggerDownload(context, target);
    const suggested = download.suggestedFilename();
    const ext = path.extname(suggested) || ".xlsx";
    const saved = await saveDownload(download, ext);
    console.log(`Saved ${suggested} -> ${saved}`);
    downloadPaths.push(saved);
  }

  console.log(`Saved ${downloadPaths.length} download(s) to ${DOWNLOAD_DIR}`);

  const csvs: string[] = [];
  for (const filePath of downloadPaths) {
    if (/\.xlsx?$/i.test(filePath)) {
      const out = await xlsxToCsvFiles(filePath);
      csvs.push(...out);
    } else if (/\.csv$/i.test(filePath)) {
      csvs.push(filePath);
    } else {
      console.warn(`Skipping CSV conversion for unsupported file type: ${filePath}`);
    }
  }

  if (!csvs.length) {
    console.warn("No CSV exports generated.");
  }

  await emailCsvs(csvs);

  await browser.close();
})().catch((err) => {
  console.error("Worker failed");
  console.error(err?.message || err);
  process.exit(1);
});
