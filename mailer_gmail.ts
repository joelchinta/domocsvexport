import { google } from "@googleapis/gmail";
import fs from "node:fs/promises";
import path from "node:path";

type SendOptions = {
  from: string;         // EMAIL_USER
  to: string;           // EMAIL_TO (comma or array supported)
  subject: string;      // EMAIL_SUBJECT
  text: string;         // EMAIL_BODY
  attachments?: { filename: string; path: string }[];
};

function toBase64Url(input: Buffer | string) {
  const b64 = (Buffer.isBuffer(input) ? input : Buffer.from(input))
    .toString("base64");
  return b64.replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

async function buildRawMessage(opts: SendOptions) {
  const boundary = "mixed_" + Math.random().toString(36).slice(2);
  const lines: string[] = [];

  lines.push(`From: ${opts.from}`);
  lines.push(`To: ${opts.to}`);
  lines.push(`Subject: ${opts.subject}`);
  lines.push(`MIME-Version: 1.0`);

  if (opts.attachments?.length) {
    lines.push(`Content-Type: multipart/mixed; boundary="${boundary}"`, "", `--${boundary}`);
    lines.push(`Content-Type: text/plain; charset="utf-8"`, "", opts.text, "");

    for (const att of opts.attachments) {
      const content = await fs.readFile(att.path);
      const b64 = content.toString("base64").replace(/(.{76})/g, "$1\n");
      const mime = att.filename.toLowerCase().endsWith(".csv")
        ? "text/csv"
        : "application/octet-stream";
      lines.push(`--${boundary}`);
      lines.push(`Content-Type: ${mime}; name="${att.filename}"`);
      lines.push(`Content-Transfer-Encoding: base64`);
      lines.push(`Content-Disposition: attachment; filename="${att.filename}"`, "", b64, "");
    }

    lines.push(`--${boundary}--`);
  } else {
    lines.push(`Content-Type: text/plain; charset="utf-8"`, "", opts.text);
  }

  const raw = lines.join("\r\n");
  return toBase64Url(raw);
}

export async function sendWithGmailAPI({
  from,
  to,
  subject,
  text,
  attachments = []
}: SendOptions) {
  const auth = new google.auth.OAuth2({
    clientId: process.env.GOOGLE_CLIENT_ID!,
    clientSecret: process.env.GOOGLE_CLIENT_SECRET!
  });
  auth.setCredentials({
    refresh_token: process.env.GOOGLE_REFRESH_TOKEN!
  });

  const gmail = google.gmail({ version: "v1", auth });

  const raw = await buildRawMessage({ from, to, subject, text, attachments });
  const res = await gmail.users.messages.send({
    userId: "me",
    requestBody: { raw }
  });

  return res.data;
}
