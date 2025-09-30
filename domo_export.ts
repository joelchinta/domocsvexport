import { chromium } from 'playwright';
import * as dotenv from 'dotenv';
import * as XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';
import * as nodemailer from 'nodemailer';

// Load environment variables
dotenv.config();

// ---------- Config and guards ----------
const REQ_ENV = [
  'DOMO_BASE_URL', 'DOMO_USERNAME', 'DOMO_PASSWORD', 'DOMO_VENDOR_ID',
  'EMAIL_USER', 'EMAIL_TO'
] as const;

for (const k of REQ_ENV) {
  if (!process.env[k]) throw new Error(`Missing required environment variable: ${k}`);
}

const DEBUG = process.env.DEBUG === '1';

// Email credential modes
const hasBasicSmtpCreds =
  Boolean(process.env.EMAIL_HOST && process.env.EMAIL_PORT && process.env.EMAIL_PASS);

const hasGoogleOAuthCreds =
  Boolean(process.env.GOOGLE_CLIENT_ID && process.env.GOOGLE_CLIENT_SECRET && process.env.GOOGLE_REFRESH_TOKEN);

if (!hasBasicSmtpCreds && !hasGoogleOAuthCreds) {
  throw new Error('Missing email credentials: provide either EMAIL_HOST/EMAIL_PORT/EMAIL_PASS or GOOGLE_CLIENT_ID/GOOGLE_CLIENT_SECRET/GOOGLE_REFRESH_TOKEN');
}

// Optional overrides
const ENV_START = (process.env.DOMO_START_DATE || '').trim(); // YYYY-MM-DD
const ENV_END   = (process.env.DOMO_END_DATE || '').trim();   // YYYY-MM-DD
const EXPORT_DIR = path.resolve('./exporter/downloads');
const EXPORT_SELECTOR = process.env.DOMO_EXPORT_SELECTOR || '.dt-button.buttons-excel';

// ---------- Utils ----------
function logDebug(msg: string) {
  if (DEBUG) console.log(msg);
}

function ensureDirectoryExists(filePath: string) {
  const dir = path.dirname(filePath);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function formatTemplate(template: string, startDate: string, endDate: string) {
  return template.replaceAll('{startDate}', startDate).replaceAll('{endDate}', endDate);
}

async function getGmailAccessToken(): Promise<string> {
  const params = new URLSearchParams({
    client_id: process.env.GOOGLE_CLIENT_ID!,
    client_secret: process.env.GOOGLE_CLIENT_SECRET!,
    refresh_token: process.env.GOOGLE_REFRESH_TOKEN!,
    grant_type: 'refresh_token'
  });

  const res = await fetch('https://oauth2.googleapis.com/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params.toString()
  });

  if (!res.ok) {
    // Avoid echoing entire response body in logs
    throw new Error(`Failed to fetch Gmail access token: ${res.status} ${res.statusText}`);
  }

  const json = await res.json();
  const token = json?.access_token as string | undefined;
  if (!token) throw new Error('Gmail token response missing access_token');
  return token;
}

async function createEmailTransporter() {
  if (hasGoogleOAuthCreds) {
    const accessToken = await getGmailAccessToken();
    return nodemailer.createTransport({
      service: 'gmail',
      auth: {
        type: 'OAuth2',
        user: process.env.EMAIL_USER,
        clientId: process.env.GOOGLE_CLIENT_ID,
        clientSecret: process.env.GOOGLE_CLIENT_SECRET,
        refreshToken: process.env.GOOGLE_REFRESH_TOKEN,
        accessToken
      }
    });
  }

  const port = Number.parseInt(process.env.EMAIL_PORT!, 10);
  return nodemailer.createTransport({
    host: process.env.EMAIL_HOST,
    port,
    secure: port === 465,
    auth: {
      user: process.env.EMAIL_USER,
      pass: process.env.EMAIL_PASS
    }
  });
}

async function sendReportEmail(csvPath: string, startDate: string, endDate: string) {
  const transporter = await createEmailTransporter();

  // Only verify and print diagnostics in DEBUG
  if (DEBUG) {
    try {
      await transporter.verify();
      logDebug('SMTP verified');
    } catch {
      logDebug('SMTP verify failed, continuing to send');
    }
  }

  const recipients = process.env.EMAIL_TO!
    .split(',')
    .map(a => a.trim())
    .filter(Boolean);

  if (!recipients.length) throw new Error('EMAIL_TO must contain at least one recipient');

  const subjectTemplate = process.env.EMAIL_SUBJECT || 'Domo Export Report ({startDate} to {endDate})';
  const bodyTemplate = process.env.EMAIL_BODY || 'Please find attached the Domo export report for the period {startDate} to {endDate}.';

  const subject = formatTemplate(subjectTemplate, startDate, endDate);
  const body = formatTemplate(bodyTemplate, startDate, endDate);

  await transporter.sendMail({
    from: process.env.EMAIL_USER,
    to: recipients,
    subject,
    text: body,
    attachments: [{ filename: path.basename(csvPath), path: csvPath }]
  });

  return recipients.length;
}

function isValidISODate(s: string) {
  return /^\d{4}-\d{2}-\d{2}$/.test(s);
}

async function getPreviousMondayAndSunday(): Promise<[string, string]> {
  // Local timezone logic. Override with DOMO_START_DATE/DOMO_END_DATE if provided.
  if (isValidISODate(ENV_START) && isValidISODate(ENV_END)) {
    return [ENV_START, ENV_END];
  }

  const today = new Date();
  // JS: Sunday=0...Saturday=6
  const day = today.getDay();
  // Go to this week's Monday
  const thisMonday = new Date(today);
  const deltaToMonday = ((day + 6) % 7); // 0 if Monday, 1 if Tuesday, ... 6 if Sunday
  thisMonday.setDate(today.getDate() - deltaToMonday);
  // Last week's Monday
  const lastMonday = new Date(thisMonday);
  lastMonday.setDate(thisMonday.getDate() - 7);
  // Last week's Sunday
  const lastSunday = new Date(lastMonday);
  lastSunday.setDate(lastMonday.getDate() + 6);

  const fmt = (d: Date) => d.toISOString().slice(0, 10);
  return [fmt(lastMonday), fmt(lastSunday)];
}

async function main() {
  const browser = await chromium.launch({ headless: true });

  try {
    const context = await browser.newContext({ acceptDownloads: true });
    const page = await context.newPage();

    // Login
    await page.goto(`${process.env.DOMO_BASE_URL}/session/login`, { waitUntil: 'domcontentloaded' });
    await page.locator('input[name="username"]').fill(process.env.DOMO_USERNAME!);
    await page.locator('input[name="password"]').fill(process.env.DOMO_PASSWORD!);
    await page.getByRole('button', { name: 'Sign in' }).click();

    // Wait for navigation post-login. If the app does not navigate, fall back to a short wait.
    await Promise.race([
      page.waitForLoadState('networkidle', { timeout: 10_000 }),
      page.waitForTimeout(2_000)
    ]);

    // Date range
    const [startDate, endDate] = await getPreviousMondayAndSunday();
    logDebug(`Using date range ${startDate} to ${endDate}`);

    // Navigate to slip page
    const url = `${process.env.DOMO_BASE_URL}/transactions/slipSingle/${process.env.DOMO_VENDOR_ID}/${startDate}/${endDate}/1`;
    await page.goto(url, { waitUntil: 'networkidle' });

    // Wait for export control
    await page.waitForSelector(EXPORT_SELECTOR, { timeout: 30_000 });

    // Prepare download
    const downloadPromise = page.waitForEvent('download');

    // Trigger export
    await page.click(EXPORT_SELECTOR);

    // Save XLSX
    const download = await downloadPromise;
    const suggested = download.suggestedFilename();
    const baseName = suggested && suggested.toLowerCase().endsWith('.xlsx')
      ? suggested
      : `domo_export_${startDate}_to_${endDate}.xlsx`;
    const xlsxPath = path.join(EXPORT_DIR, baseName);
    ensureDirectoryExists(xlsxPath);
    await download.saveAs(xlsxPath);

    // Convert to CSV
    const workbook = (XLSX as any).default ? (XLSX as any).default.readFile(xlsxPath) : XLSX.readFile(xlsxPath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const csvContent = XLSX.utils.sheet_to_csv(worksheet);
    const csvPath = path.join(EXPORT_DIR, path.basename(baseName).replace(/\.xlsx$/i, '.csv'));
    fs.writeFileSync(csvPath, csvContent);

    // Clean xlsx to minimize artifacts
    try { fs.unlinkSync(xlsxPath); } catch {}

    // Email report
    const recipientCount = await sendReportEmail(csvPath, startDate, endDate);

    // Minimal logs
    console.log('Export completed');
    console.log('Email sent');

    // Extra diagnostics only if DEBUG
    logDebug(`File: ${path.basename(csvPath)}`);
    logDebug(`Recipients: ${recipientCount}`);
    logDebug(`Date range: ${startDate} to ${endDate}`);

  } catch (err: any) {
    // Do not dump objects that might include sensitive info
    const msg = err?.message || String(err);
    console.error('Worker failed');
    if (DEBUG && err?.stack) console.error(err.stack);
    else console.error(msg);
    throw err;
  } finally {
    await browser.close();
  }
}

main().catch(() => process.exit(1));
