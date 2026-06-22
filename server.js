/**
 * server.js — Twilio Missed-Call Text-Back System
 * CRM: Google Sheets | AI: Google Gemini | SMS: Twilio
 *
 * Required environment variables:
 *   GOOGLE_SERVICE_ACCOUNT_EMAIL  – Service account client_email
 *   GOOGLE_PRIVATE_KEY            – Service account private_key (with literal \n)
 *   GOOGLE_SHEET_ID               – Spreadsheet ID from the sheet URL
 *   TWILIO_ACCOUNT_SID            – Twilio Account SID
 *   TWILIO_AUTH_TOKEN             – Twilio Auth Token
 *   TWILIO_PHONE_NUMBER           – Your Twilio number in E.164 format (+1...)
 *   GEMINI_API_KEY                – Google AI Studio API key
 *   BOOKING_LINK                  – Your calendar / booking URL
 *   PORT                          – (optional) HTTP port, defaults to 3000
 */

"use strict";

// ─── Core deps ────────────────────────────────────────────────────────────────
const express = require("express");
const { google } = require("googleapis");
const twilio = require("twilio");
const { GoogleGenAI } = require("@google/genai");

// ─── Config ───────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
const SHEET_ID = process.env.GOOGLE_SHEET_ID;
const SHEET_TAB = "CRM"; // Name of the tab inside your spreadsheet
const BOOKING_LINK = process.env.BOOKING_LINK || "https://example.com/book";
const RATE_LIMIT_MINUTES = 15;

// Google Sheets column indices (0-based) — matches the header row order:
//   A=DateTime  B=PhoneNumber  C=Status  D=LastAIResponse
const COL = { DATE: 0, PHONE: 1, STATUS: 2, AI_RESPONSE: 3 };

// ─── System prompt — customise this for your business ─────────────────────────
const SYSTEM_PROMPT = `
You are a friendly assistant for [Clawsmedia].

Services & Pricing:
- [Service 1]: $[Price]
- [Service 2]: $[Price]
- [Service 3]: $[Price]

FAQs:
- Hours: Monday–Friday 9 am–6 pm, Saturday 10 am–4 pm.
- Location: [Your address or "fully remote"].
- Booking: Customers can book at ${BOOKING_LINK}.
- Cancellations: 24-hour notice required for a full refund.
- Payments: We accept all major credit cards and cash.

Instructions:
- Reply in a warm, concise, professional tone.
- Your reply MUST be 160 characters or fewer (one SMS).
- Never make up information not listed above.
- If unsure, say: "Great question! Call us at [+16478702542] for details."
`.trim();

// ─── Clients ──────────────────────────────────────────────────────────────────
const twilioClient = twilio(
  process.env.TWILIO_ACCOUNT_SID,
  process.env.TWILIO_AUTH_TOKEN
);

const genai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

// Google Sheets auth via Service Account JWT
const sheetsAuth = new google.auth.JWT({
  email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
  key: (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n"),
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});
const sheets = google.sheets({ version: "v4", auth: sheetsAuth });

// ─── Google Sheets helpers ────────────────────────────────────────────────────

/**
 * Fetch all data rows (skips the header row).
 * @returns {Promise<string[][]>} 2-D array of cell values
 */
async function getAllRows() {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_TAB}!A2:D`,
  });
  return res.data.values || [];
}

/**
 * Append a new row at the bottom of the sheet.
 */
async function appendRow(rowData) {
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: `${SHEET_TAB}!A:D`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values: [rowData] },
  });
}

/**
 * Update a single cell by 1-based row number and 0-based column index.
 * Column 0 → A, 1 → B, 2 → C, 3 → D
 */
async function updateCell(sheetRowNumber, colIndex, value) {
  const colLetter = String.fromCharCode(65 + colIndex); // 0→A, 1→B …
  const range = `${SHEET_TAB}!${colLetter}${sheetRowNumber}`;
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range,
    valueInputOption: "USER_ENTERED",
    requestBody: { values: [[value]] },
  });
}

/**
 * Find the most recent row for a given E.164 phone number.
 * Returns { rowIndex (0-based in array), sheetRow (1-based + header offset),
 *           row (string[]) } or null if not found.
 */
async function findLatestRowForPhone(phone) {
  const rows = await getAllRows();
  // Iterate in reverse so we get the most recent entry first
  for (let i = rows.length - 1; i >= 0; i--) {
    if (rows[i][COL.PHONE] === phone) {
      return {
        rowIndex: i,
        sheetRow: i + 2, // +1 for 0-based→1-based, +1 for header row
        row: rows[i],
      };
    }
  }
  return null;
}

// ─── AI helper ────────────────────────────────────────────────────────────────

/**
 * Generate a ≤160-char SMS reply using Gemini.
 * @param {string} userMessage
 * @returns {Promise<string>}
 */
async function generateAIReply(userMessage) {
  const model = "gemini-4.1";
  const result = await genai.models.generateContent({
    model,
    config: { systemInstruction: SYSTEM_PROMPT },
    contents: [{ role: "user", parts: [{ text: userMessage }] }],
  });

  let reply = result.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || "";

  // Hard-cap at 160 chars as a safety net
  if (reply.length > 160) {
    reply = reply.substring(0, 157) + "...";
  }
  return reply;
}

// ─── Twilio signature validation middleware ───────────────────────────────────

/**
 * Validates that requests genuinely come from Twilio.
 * Set TWILIO_WEBHOOK_VALIDATE=true in production.
 */
function twilioValidate(req, res, next) {
  if (process.env.TWILIO_WEBHOOK_VALIDATE !== "true") {
    return next(); // Skip in development
  }
  const valid = twilio.validateRequest(
    process.env.TWILIO_AUTH_TOKEN,
    req.headers["x-twilio-signature"] || "",
    `${req.protocol}://${req.get("host")}${req.originalUrl}`,
    req.body
  );
  if (!valid) {
    console.warn("⚠️  Invalid Twilio signature — request rejected");
    return res.status(403).send("Forbidden");
  }
  next();
}

// ─── Express app ──────────────────────────────────────────────────────────────
const app = express();

// Twilio sends URL-encoded POST bodies
app.use(express.urlencoded({ extended: false }));
app.use(express.json());

// ─── Health check ─────────────────────────────────────────────────────────────
app.get("/health", (_req, res) => res.json({ status: "ok" }));

// ─── POST /webhook/missed-call ────────────────────────────────────────────────
app.post("/webhook/missed-call", twilioValidate, async (req, res) => {
  // Twilio sends the caller's number in the "From" (or "Called") field
  const callerPhone = req.body.From || req.body.Called;

  if (!callerPhone) {
    console.warn("missed-call webhook: no phone number in payload");
    return res.status(400).send("Missing caller phone");
  }

  console.log(`📞 Missed call from ${callerPhone}`);

  try {
    // ── Rate-limit check ──────────────────────────────────────────────────────
    const rows = await getAllRows();
    const now = Date.now();
    const rateLimitMs = RATE_LIMIT_MINUTES * 60 * 1000;

    const recentEntry = rows.find((row) => {
      if (row[COL.PHONE] !== callerPhone) return false;
      const entryTime = new Date(row[COL.DATE]).getTime();
      return now - entryTime < rateLimitMs;
    });

    if (recentEntry) {
      console.log(
        `⏭️  Rate-limited: ${callerPhone} already contacted within ${RATE_LIMIT_MINUTES} min`
      );
      return res.status(200).send("Rate limited — no action taken");
    }

    // ── Send initial SMS ──────────────────────────────────────────────────────
    const message =
      `Hey, sorry we missed you! You can book a slot instantly here: ${BOOKING_LINK}` +
      ` or reply to this text with any questions.`;

    await twilioClient.messages.create({
      body: message,
      from: process.env.TWILIO_PHONE_NUMBER,
      to: callerPhone,
    });
    console.log(`✉️  Initial SMS sent to ${callerPhone}`);

    // ── Log to Google Sheet ───────────────────────────────────────────────────
    const dateTime = new Date().toISOString();
    await appendRow([dateTime, callerPhone, "Initial Text Sent", "N/A"]);
    console.log(`📋 Row appended to sheet for ${callerPhone}`);

    return res.status(200).send("OK");
  } catch (err) {
    console.error("Error in /webhook/missed-call:", err);
    return res.status(500).send("Internal Server Error");
  }
});

// ─── POST /webhook/incoming-sms ───────────────────────────────────────────────
app.post("/webhook/incoming-sms", twilioValidate, async (req, res) => {
  const senderPhone = req.body.From;
  const incomingText = (req.body.Body || "").trim();

  if (!senderPhone || !incomingText) {
    console.warn("incoming-sms webhook: missing From or Body");
    return res.status(400).send("Missing required fields");
  }

  console.log(`💬 SMS from ${senderPhone}: "${incomingText}"`);

  try {
    // ── Update status in sheet ────────────────────────────────────────────────
    const entry = await findLatestRowForPhone(senderPhone);

    if (entry) {
      await updateCell(entry.sheetRow, COL.STATUS, "Engaged");
      console.log(`📋 Status → "Engaged" for row ${entry.sheetRow}`);
    } else {
      // Unsolicited text from a number not in our sheet — log it anyway
      const dateTime = new Date().toISOString();
      await appendRow([dateTime, senderPhone, "Engaged", "N/A"]);
      console.log(`📋 New unsolicited contact ${senderPhone} — row created`);
    }

    // ── Generate AI reply ─────────────────────────────────────────────────────
    const aiReply = await generateAIReply(incomingText);
    console.log(`🤖 AI reply (${aiReply.length} chars): "${aiReply}"`);

    // ── Send AI reply via SMS ─────────────────────────────────────────────────
    await twilioClient.messages.create({
      body: aiReply,
      from: process.env.TWILIO_PHONE_NUMBER,
      to: senderPhone,
    });
    console.log(`✉️  AI reply sent to ${senderPhone}`);

    // ── Persist AI response in sheet ──────────────────────────────────────────
    if (entry) {
      await updateCell(entry.sheetRow, COL.AI_RESPONSE, aiReply);
    } else {
      // Re-fetch to get the row we just appended
      const newEntry = await findLatestRowForPhone(senderPhone);
      if (newEntry) {
        await updateCell(newEntry.sheetRow, COL.AI_RESPONSE, aiReply);
      }
    }
    console.log(`📋 Last AI Response updated for ${senderPhone}`);

    // Twilio expects an empty 200 (or a TwiML response) — we reply via REST API
    return res.status(200).send("OK");
  } catch (err) {
    console.error("Error in /webhook/incoming-sms:", err);
    return res.status(500).send("Internal Server Error");
  }
});

// ─── 404 fallback ─────────────────────────────────────────────────────────────
app.use((_req, res) => res.status(404).send("Not Found"));

// ─── Start ────────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`🚀 Server listening on port ${PORT}`);
  console.log(`   POST /webhook/missed-call`);
  console.log(`   POST /webhook/incoming-sms`);
});

module.exports = app; // for testing
