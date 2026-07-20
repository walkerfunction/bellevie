/**
 * Belle Vie Studios — Inventory AI Auto-fill
 * ─────────────────────────────────────────
 * Install: Google Sheet → Extensions → Apps Script → paste this → Save
 *
 * One-time setup:
 *   1. Run setApiKey() once with your Anthropic key
 *   2. Add an onEdit trigger: Triggers → + Add Trigger → onEdit → On edit
 *
 * How it works:
 *   Jewelry maker pastes a Google Drive photo link into column A of any
 *   collection sheet. This script detects it, sends the image to Claude,
 *   and fills in Name (B), Price (C), Description (D), Sold = No (E).
 */

const ANTHROPIC_URL = 'https://api.anthropic.com/v1/messages';

// ── TRIGGER ──────────────────────────────────────────────────────────────────
function onEdit(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  // Skip _Config, _Template, and any internal sheets (underscore prefix)
  if (sheetName.startsWith('_')) return;

  // Only fire when column A (Image Link) is edited in a data row (row 2+)
  if (e.range.getColumn() !== 1 || e.range.getRow() < 2) return;

  const imageLink = (e.value || '').trim();
  if (!imageLink.includes('drive.google.com')) return;

  // Skip if Name is already filled (avoids re-triggering on manual edits)
  const row = e.range.getRow();
  if (sheet.getRange(row, 2).getValue()) return;

  // Show loading state so the jewelry maker knows it's working
  sheet.getRange(row, 2).setValue('⏳ Analyzing photo…');
  SpreadsheetApp.flush();

  try {
    const details = analyzeImage(imageLink, sheetName);
    sheet.getRange(row, 2).setValue(details.name);
    sheet.getRange(row, 3).setValue(details.price);
    sheet.getRange(row, 4).setValue(details.description);
    sheet.getRange(row, 5).setValue('No');
  } catch (err) {
    sheet.getRange(row, 2).setValue('⚠️ ' + err.message);
  }
}

// ── IMAGE ANALYSIS ────────────────────────────────────────────────────────────
function analyzeImage(driveUrl, collectionName) {
  const fileId = extractDriveFileId(driveUrl);

  // Fetch image bytes from Drive (file must be shared "Anyone with link can view")
  const imgRes = UrlFetchApp.fetch(
    `https://drive.google.com/uc?export=view&id=${fileId}`,
    { muteHttpExceptions: true }
  );

  if (imgRes.getResponseCode() !== 200) {
    throw new Error('Cannot access photo. Make sure it is shared publicly in Drive.');
  }

  const base64 = Utilities.base64Encode(imgRes.getContent());
  const mime   = imgRes.getHeaders()['Content-Type'] || 'image/jpeg';

  const prompt = `You are cataloguing handcrafted items for Belle Vie Studios, a small artisan shop.
This item belongs to the "${collectionName}" collection.

Examine the image and return ONLY a JSON object with these fields:
- name: short elegant product name (4-8 words)
- price: recommended retail price in USD for a handcrafted artisan piece, e.g. "$48.00"
- description: 2 warm sentences highlighting what makes this piece special

Respond with valid JSON only, no markdown, no explanation.
Example: {"name":"Crescent Moon Ceramic Pendant","price":"$42.00","description":"Hand-shaped..."}`;

  const payload = {
    model: 'claude-opus-4-8',
    max_tokens: 300,
    messages: [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: mime, data: base64 } },
        { type: 'text', text: prompt }
      ]
    }]
  };

  const key = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!key) throw new Error('API key not set. Run setApiKey() first.');

  const res = UrlFetchApp.fetch(ANTHROPIC_URL, {
    method: 'post',
    headers: {
      'x-api-key': key,
      'anthropic-version': '2023-06-01',
      'content-type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const body = JSON.parse(res.getContentText());
  if (body.error) throw new Error(body.error.message);

  return JSON.parse(body.content[0].text);
}

// ── HELPERS ───────────────────────────────────────────────────────────────────
function extractDriveFileId(url) {
  const patterns = [
    /\/file\/d\/([a-zA-Z0-9_-]+)/,
    /\/d\/([a-zA-Z0-9_-]+)/,
    /[?&]id=([a-zA-Z0-9_-]+)/
  ];
  for (const p of patterns) {
    const m = url.match(p);
    if (m) return m[1];
  }
  throw new Error('Could not read Drive file ID from URL.');
}

// ── SETUP (run once manually) ─────────────────────────────────────────────────
function setApiKey() {
  // Replace the string below with your actual Anthropic API key, run once, then delete the key from here
  PropertiesService.getScriptProperties().setProperty('ANTHROPIC_API_KEY', 'sk-ant-REPLACE_ME');
  Logger.log('API key saved.');
}
