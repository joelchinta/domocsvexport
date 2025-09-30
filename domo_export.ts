import { chromium } from 'playwright';
import * as dotenv from 'dotenv';
import * as XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';
import * as nodemailer from 'nodemailer';

// Load environment variables
dotenv.config();

// Validate required environment variables
const alwaysRequiredEnv = [
    'DOMO_BASE_URL', 'DOMO_USERNAME', 'DOMO_PASSWORD', 'DOMO_VENDOR_ID',
    'EMAIL_USER', 'EMAIL_TO'
];
for (const envVar of alwaysRequiredEnv) {
    if (!process.env[envVar]) {
        throw new Error(`Missing required environment variable: ${envVar}`);
    }
}

const hasBasicSmtpCreds = Boolean(
    process.env.EMAIL_HOST && process.env.EMAIL_PORT && process.env.EMAIL_PASS
);

const hasGoogleOAuthCreds = Boolean(
    process.env.GOOGLE_CLIENT_ID && process.env.GOOGLE_CLIENT_SECRET && process.env.GOOGLE_REFRESH_TOKEN
);

if (!hasBasicSmtpCreds && !hasGoogleOAuthCreds) {
    throw new Error('Missing email credentials: provide either EMAIL_HOST/EMAIL_PORT/EMAIL_PASS or GOOGLE_CLIENT_ID/GOOGLE_CLIENT_SECRET/GOOGLE_REFRESH_TOKEN');
}

const XLSXLib: typeof XLSX & { default?: typeof XLSX } = (XLSX as any).default ?? (XLSX as any);

function ensureDirectoryExists(filePath: string) {
    const directory = path.dirname(filePath);
    if (!fs.existsSync(directory)) {
        fs.mkdirSync(directory, { recursive: true });
    }
}

function formatTemplate(template: string, startDate: string, endDate: string) {
    return template
        .replaceAll('{startDate}', startDate)
        .replaceAll('{endDate}', endDate);
}

async function getGmailAccessToken(): Promise<string> {
    const params = new URLSearchParams({
        client_id: process.env.GOOGLE_CLIENT_ID!,
        client_secret: process.env.GOOGLE_CLIENT_SECRET!,
        refresh_token: process.env.GOOGLE_REFRESH_TOKEN!,
        grant_type: 'refresh_token'
    });

    const response = await fetch('https://oauth2.googleapis.com/token', {
        method: 'POST',
        headers: {
            'Content-Type': 'application/x-www-form-urlencoded'
        },
        body: params.toString()
    });

    if (!response.ok) {
        const errorBody = await response.text().catch(() => '<no body>');
        throw new Error(`Failed to fetch Gmail access token: ${response.status} ${response.statusText} - ${errorBody}`);
    }

    const tokenPayload = await response.json();
    const accessToken = tokenPayload.access_token as string | undefined;

    if (!accessToken) {
        throw new Error('Gmail token response did not include access_token');
    }

    return accessToken;
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
            },
        });
    }

    return nodemailer.createTransport({
        host: process.env.EMAIL_HOST,
        port: parseInt(process.env.EMAIL_PORT!, 10),
        secure: process.env.EMAIL_PORT === '465',
        auth: {
            user: process.env.EMAIL_USER,
            pass: process.env.EMAIL_PASS,
        },
    });
}

async function sendReportEmail(csvPath: string, startDate: string, endDate: string) {
    const transporter = await createEmailTransporter();

    await transporter.verify();

    const recipients = process.env.EMAIL_TO!
        .split(',')
        .map(address => address.trim())
        .filter(Boolean);

    if (!recipients.length) {
        throw new Error('EMAIL_TO must contain at least one valid recipient address.');
    }

    const subjectTemplate = process.env.EMAIL_SUBJECT || 'Domo Export Report ({startDate} to {endDate})';
    const bodyTemplate = process.env.EMAIL_BODY || 'Please find attached the Domo export report for the period {startDate} to {endDate}.';

    const subject = formatTemplate(subjectTemplate, startDate, endDate);
    const body = formatTemplate(bodyTemplate, startDate, endDate);

    await transporter.sendMail({
        from: process.env.EMAIL_USER,
        to: recipients,
        subject,
        text: body,
        attachments: [
            {
                filename: path.basename(csvPath),
                path: csvPath
            }
        ]
    });

    return recipients.length;
}

async function getPreviousMondayAndSunday(): Promise<[string, string]> {
    const today = new Date();
    const lastMonday = new Date(today);
    // First, go back to this week's Monday
    lastMonday.setDate(today.getDate() - today.getDay() + 1);
    // Then go back one more week
    lastMonday.setDate(lastMonday.getDate() - 7);
    
    const lastSunday = new Date(lastMonday);
    lastSunday.setDate(lastMonday.getDate() + 6);

    // Format dates as YYYY-MM-DD
    const formatDate = (date: Date) => {
        return date.toISOString().split('T')[0];
    };

    return [formatDate(lastMonday), formatDate(lastSunday)];
}

async function main() {
    const browser = await chromium.launch({
        headless: true
    });

    try {
        const context = await browser.newContext();
        const page = await context.newPage();

        // Login
        await page.goto(`${process.env.DOMO_BASE_URL}/session/login`);
        await page.locator('input[name="username"]').click();
        await page.locator('input[name="username"]').fill(process.env.DOMO_USERNAME!);
        await page.locator('input[name="password"]').fill(process.env.DOMO_PASSWORD!);
        await page.getByRole('button', { name: 'Sign in' }).click();

        // Wait for login to complete
        await page.waitForTimeout(2000);

        // Get date range
        const [startDate, endDate] = await getPreviousMondayAndSunday();

        // Navigate directly to the transaction slip URL
        const url = `${process.env.DOMO_BASE_URL}/transactions/slipSingle/${process.env.DOMO_VENDOR_ID}/${startDate}/${endDate}/1`;
        await page.goto(url);

        // Wait for the page to load
        await page.waitForLoadState('networkidle');

        // Wait for table to load and export button to be visible
        await page.waitForSelector('.dt-button.buttons-excel');

        // Setup download promise before clicking the export button
        const downloadPromise = page.waitForEvent('download');
        
        // Click the Excel export button
        await page.click('.dt-button.buttons-excel');

        // Wait for the download to start and save the file
        const download = await downloadPromise;
        const xlsxPath = `./exporter/downloads/domo_export_${startDate}_to_${endDate}.xlsx`;
        ensureDirectoryExists(xlsxPath);
        await download.saveAs(xlsxPath);

        // Convert XLSX to CSV
        const workbook = XLSXLib.readFile(xlsxPath);
        const csvPath = `./exporter/downloads/domo_export_${startDate}_to_${endDate}.csv`;
        
        // Get the first worksheet
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        
        // Convert to CSV and save
        const csvContent = XLSXLib.utils.sheet_to_csv(worksheet);
        fs.writeFileSync(csvPath, csvContent);

        // Delete the XLSX file if you don't need it
        fs.unlinkSync(xlsxPath);

        // Send email with CSV attachment
        const recipientCount = await sendReportEmail(csvPath, startDate, endDate);

        console.log('Export completed successfully!');
        console.log(`Date range: ${startDate} to ${endDate}`);
        console.log(`File saved as: ${path.basename(csvPath)}`);
        console.log(`Email sent to ${recipientCount} recipient(s).`);

    } catch (error) {
        console.error('An error occurred:', error);
        throw error;
    } finally {
        await browser.close();
    }
}

main().catch(console.error);
