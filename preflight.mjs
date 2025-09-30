// OAuth only
const req = ['DOMO_BASE_URL','DOMO_USERNAME','DOMO_PASSWORD','DOMO_VENDOR_ID','EMAIL_USER','EMAIL_TO'];
const google = ['GOOGLE_CLIENT_ID','GOOGLE_CLIENT_SECRET','GOOGLE_REFRESH_TOKEN'];

function show(keys, label){
  console.log(`[${label}]`);
  for(const k of keys){
    const v = process.env[k];
    console.log(`${k}:`, v && v.trim().length ? `set (${v.trim().length} chars)` : 'NOT SET');
  }
}

show(req, 'required');
show(google, 'google');

const missing = req.filter(k => !process.env[k]?.trim());
const hasGoogle = google.every(k => process.env[k]?.trim());

if (missing.length) {
  console.error(`Missing required env: ${missing.join(', ')}`);
  process.exit(2);
}
if (!hasGoogle) {
  console.error('Google OAuth not configured. Set GOOGLE_CLIENT_ID, GOOGLE_CLIENT_SECRET, GOOGLE_REFRESH_TOKEN.');
  process.exit(2);
}

// simple sanity
if (!/^[^@]+@[^@]+\.[^@]+$/.test(process.env.EMAIL_USER)) {
  console.error('EMAIL_USER does not look like an email address');
  process.exit(2);
}

console.log('Preflight passed.');
