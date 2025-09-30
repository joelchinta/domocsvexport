import { chromium } from 'playwright';
import * as dotenv from 'dotenv';
import * as XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';
import * as nodemailer from 'nodemailer';

dotenv.config();

// ---------- Config ----------
const DEBUG = process.env.DEBUG === '1';
const EXPORT_DIR = path.resolve('./exporter/downloads');
const EXPORT_SELECTOR = process.env.DOMO_EXPORT_SELECTOR || '.dt-button.buttons-excel';

function logDebug(msg: string) { if (DEBUG) console.log(msg); }
function isNonEmpty(v: string | undefined | null) { return typeof v === 'string' && v.trim() !== ''; }

function requireEnv(keys: string[]) {
  const missing = keys.filter(k => !isNonEmpty(process.env[k]));
  if (missing.length) throw new Error(`Missing required env: ${missing.join(', ')}`);
}

function ensureDirectoryExists(filePath: string) {
  const dir = path.dirname(filePath);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function formatTemplate(t: string, startDate: string, endDate: string) {
  return t.replaceAll('{startDate}', startDate).replaceAll('{endDate}', endDate);
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

  if (!res.ok) throw new Error(`Failed to fetch Gmail access token: ${res.status} ${res.statusText}`);

  const json = await res.json();
  const token = json?.access_token as string | undefined;
  if (!token) throw new Error('Gmail token response missing access_token');
  return token;
}

async function createEmailTransporter() {
  const hasSMTP = isNonEmpty(process.env.EMAIL_HOST) && isNonEmpty(process.env.EMAIL_PORT) && isNonEmpty(process.env.EMAIL_PASS);
  const hasOAuth = isNonEmpty(process.env.GOOGLE_CLIENT_ID) && isNonEmpty(process.env.GOOGLE_CLIENT_SECRET) && isNonEmpty(process.env.GOOGLE_REFRESH_TOKEN);

  if (!hasSMTP && !hasOAuth) {
    throw new Error('Email creds missing. Provide either SMTP (EMAIL_HOST, EMAIL_PORT, EMAIL_PASS) or Google OAuth (GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, GOOGLE_REFRESH_TOKEN).');
  }

  if (hasOAuth) {
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
    auth: { user: process.env.EMAIL_USER, pass: process.env.EMAIL_PASS }
  });
}

async function sendReportEmail(csvPath: string, startDate: string, endDate: string) {
  const transporter = await createEmailTransporter();

  if (DEBUG) {
    try { await transporter.verify(); logDebug('SMTP verified'); }
    catch { logDebug('SMTP verify failed. Will attempt send anyway.'); }
  }

  const recipients = process.env.EMAIL_TO!
    .split(',')
    .map(a => a.trim())
    .filter(Boolean);

  if (recipients.length === 0) throw new Error('EMAIL_TO must contain at least one recipient');

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

function isValidISO(s: string) { return /^\d{4}-\d{2}-\d{2}$/.test(s); }

async function getPreviousMondayAndSunday(): Promise<[string, string]> {
  const envStart = (process.env.DOMO_START_DATE || '').trim();
  const envEnd = (process.env.DOMO_END_DATE || '').trim();
  if (isValidISO(envStart) && isValidISO(envEnd)) return [envStart, envEnd];

  const today = new Date();
  const day = today.getDay(); // 0 Sun .. 6 Sat
  const thisMonday = new Date(today);
  const deltaToMonday = ((day + 6) % 7);
  thisMonday.setDate(today.getDate() - deltaToMonday);
  const lastMonday = new Date(thisMonday);
  lastMonday.setDate(thisMonday.getDate() - 7);
  const lastSunday = new Date(lastMonday);
  lastSunday.setDate(lastMonday.getDate() + 6);
  const fmt = (d: Date) => d.toISOString().slice(0, 10);
  return [fmt(lastMonday), fmt(lastSunday)];
}

async function main() {
  // Validate env first, inside main so errors are visible
  requireEnv(['DOMO_BASE_URL', 'DOMO_USERNAME', 'DOMO_PASSWORD', 'DOMO_VENDOR_ID', 'EMAIL_USER', 'EMAIL_TO']);

  // Diagnostics
  logDebug(`hasSMTP=${isNonEmpty(process.env.EMAIL_HOST) && isNonEmpty(process.env.EMAIL_PORT) && isNonEmpty(process.env.EMAIL_PASS)}`);
  logDebug(`hasOAuth=${isNonEmpty(process.env.GOOGLE_CLIENT_ID) && isNonEmpty(process.env.GOOGLE_CLIENT_SECRET) && isNonEmpty(process.env.GOOGLE_REFRESH_TOKEN)}`);

  const browser = await chromium.launch({ headless: true });
  try {
    const context = await browser.newContext({ acceptDownloads: true });
    const page = await context.newPage();

    await page.goto(`${process.env.DOMO_BASE_URL}/session/login`, { waitUntil: 'domcontentloaded' });
    await page.locator('input[name="username"]').fill(process.env.DOMO_USERNAME!);
    await page.locator('input[name="password"]').fill(process.env.DOMO_PASSWORD!);
    await page.getByRole('button', { name: 'Sign in' }).click();

    await Promise.race([
      page.waitForLoadState('networkidle', { timeout: 10_000 }),
      page.waitForTimeout(2_000)
    ]);

    const [startDate, endDate] = await getPreviousMondayAndSunday();
    logDebug(`Using date range ${startDate} to ${endDate}`);

    const url = `${process.env.DOMO_BASE_URL}/transactions/slipSingle/${process.env.DOMO_VENDOR_ID}/${startDate}/${endDate}/1`;
    await page.goto(url, { waitUntil: 'networkidle' });

    await page.waitForSelector(EXPORT_SELECTOR, { timeout: 30_000 });

    const downloadPromise = page.waitForEvent('download');
    await page.click(EXPORT_SELECTOR);
    const download = await downloadPromise;

    const suggested = download.suggestedFilename();
    const baseName = suggested && suggested.toLowerCase().endsWith('.xlsx')
      ? suggested
      : `domo_export_${startDate}_to_${endDate}.xlsx`;
    const xlsxPath = path.join(EXPORT_DIR, baseName);
    ensureDirectoryExists(xlsxPath);
    await download.saveAs(xlsxPath);

    const workbook = (XLSX as any).default ? (XLSX as any).default.readFile(xlsxPath) : XLSX.readFile(xlsxPath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const csvContent = XLSX.utils.sheet_to_csv(worksheet);
    const csvPath = path.join(EXPORT_DIR, path.basename(baseName).replace(/\.xlsx$/i, '.csv'));
    fs.writeFileSync(csvPath, csvContent);

    try { fs.unlinkSync(xlsxPath); } catch {}

    const count = await sendReportEmail(csvPath, startDate, endDate);

    console.log('Export completed');
    console.log('Email sent');
    logDebug(`Recipients: ${count}`);
    logDebug(`File: ${path.basename(csvPath)}`);
    logDebug(`Date range: ${startDate} to ${endDate}`);

  } catch (err: any) {
    console.error('Worker failed');
    if (DEBUG && err?.stack) console.error(err.stack);
    else console.error(err?.message || String(err));
    throw err;
  } finally {
    await browser.close();
  }
}

main().catch(() => process.exit(1));
