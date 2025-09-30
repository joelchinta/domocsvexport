// Minimal, helpful pre-run validation + visibility in logs
const must = [
  "DOMO_BASE_URL",
  "DOMO_USERNAME",
  "DOMO_PASSWORD",
  "EMAIL_USER",
  "EMAIL_TO",
  "GOOGLE_CLIENT_ID",
  "GOOGLE_CLIENT_SECRET",
  "GOOGLE_REFRESH_TOKEN"
];

console.log("[required env]");
let ok = true;
for (const k of must) {
  const val = process.env[k];
  if (!val) {
    console.log(`❌ ${k}: MISSING`);
    ok = false;
  } else {
    console.log(`✅ ${k}: set (${String(val).length} chars)`);
  }
}

if (!ok) {
  console.error("Preflight failed: missing env.");
  process.exit(1);
} else {
  console.log("Preflight passed.");
}
