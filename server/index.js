/**
 * index.js — watsonx Dashboard backend proxy
 *
 * Exposes a single endpoint:
 *   POST /api/chat
 *   Body: { question: string, tickets: Array }
 *   Response: { answer: string }
 *
 * Security controls
 * ─────────────────
 * • Binds to 127.0.0.1 only — never 0.0.0.0 (IBM policy).
 * • API key is read from .env — never hardcoded.
 * • CORS is restricted to the local Vite dev server origin only.
 * • All user inputs are validated for type and length before use.
 * • Error responses return generic messages — no stack traces to client.
 * • Detailed errors are logged server-side only.
 * • Request body size is capped at 2 MB to prevent payload abuse.
 */

'use strict';

require('dotenv').config();

const express = require('express');
const cors    = require('cors');
const { askWatsonx } = require('./watsonx');

// ── Constants ────────────────────────────────────────────────────────────────

const PORT            = parseInt(process.env.PORT || '3001', 10);
const HOST            = '127.0.0.1';                 // localhost only — IBM policy
const ALLOWED_ORIGIN  = 'http://localhost:5173';     // Vite default dev port
const MAX_QUESTION_LEN = 1000;                        // characters
const MAX_TICKETS      = 5000;                        // rows

// ── App setup ────────────────────────────────────────────────────────────────

const app = express();

// Body size cap — prevents oversized payloads
app.use(express.json({ limit: '2mb' }));

// CORS — only the local Vite dev server is allowed
app.use(cors({
  origin: ALLOWED_ORIGIN,
  methods: ['POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type']
}));

// ── Health check ─────────────────────────────────────────────────────────────

app.get('/health', (_req, res) => {
  res.json({ status: 'ok' });
});

// ── POST /api/chat ────────────────────────────────────────────────────────────

app.post('/api/chat', async (req, res) => {
  try {
    const { question, tickets } = req.body;

    // ── Input validation ──────────────────────────────────────────────────

    if (typeof question !== 'string' || question.trim().length === 0) {
      return res.status(400).json({ error: 'question must be a non-empty string.' });
    }

    if (question.trim().length > MAX_QUESTION_LEN) {
      return res.status(400).json({
        error: `question must be ${MAX_QUESTION_LEN} characters or fewer.`
      });
    }

    if (!Array.isArray(tickets)) {
      return res.status(400).json({ error: 'tickets must be an array.' });
    }

    if (tickets.length === 0) {
      return res.status(400).json({
        error: 'No ticket data provided. Upload an Excel file and apply your filters first.'
      });
    }

    if (tickets.length > MAX_TICKETS) {
      return res.status(400).json({
        error: `Too many tickets in payload (max ${MAX_TICKETS}).`
      });
    }

    // ── Call watsonx.ai ────────────────────────────────────────────────────

    const answer = await askWatsonx(question.trim(), tickets);

    return res.json({ answer });

  } catch (err) {
    // Log detailed error server-side; send only a generic message to client
    console.error('[/api/chat] Error:', err.message);
    return res.status(500).json({
      error: 'The AI service encountered an error. Please try again shortly.'
    });
  }
});

// ── 404 catch-all ─────────────────────────────────────────────────────────────

app.use((_req, res) => {
  res.status(404).json({ error: 'Not found.' });
});

// ── Start ─────────────────────────────────────────────────────────────────────

app.listen(PORT, HOST, () => {
  console.log(`[server] Listening on http://${HOST}:${PORT}`);
  console.log(`[server] CORS allowed origin: ${ALLOWED_ORIGIN}`);
  console.log(`[server] Model: ${process.env.WATSONX_MODEL_ID || 'ibm/granite-13b-instruct-v2'}`);
});
