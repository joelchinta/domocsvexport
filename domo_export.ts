import { chromium } from 'playwright';
import * as dotenv from 'dotenv';
import ExcelJS from 'exceljs';
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
    lastMonday.setDate(today.getDa
