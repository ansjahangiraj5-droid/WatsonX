'use strict';

require('dotenv').config();

const express = require('express');
const cors    = require('cors');
const { askGemini } = require('./gemini');

// ── Constants ────────────────────────────────────────────────────────────────
const PORT            = parseInt(process.env.PORT || '3001', 10);
const HOST            = '127.0.0.1';                 
const ALLOWED_ORIGIN  = ['http://localhost:5173', 'http://127.0.0.1:5173','https://watson-x.vercel.app/'];     
const MAX_QUESTION_LEN = 1000;                        
const MAX_TICKETS      = 5000;                        

// ── App setup ────────────────────────────────────────────────────────────────
const app = express();
 app.use(express.json({ limit: '2mb' }));
// app.use(cors({
//   origin: ALLOWED_ORIGIN,
//   methods: ['POST', 'OPTIONS'],
//   allowedHeaders: ['Content-Type']
// }));

app.use(cors({
  origin: function (origin, callback) {
    // allow requests with no origin (like mobile apps or curl requests)
    if (!origin) return callback(null, true);
    
    if (ALLOWED_ORIGIN.indexOf(origin) === -1) {
      var msg = 'The CORS policy for this site does not allow access from the specified Origin.';
      return callback(new Error(msg), false);
    }
    return callback(null, true);
  },
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
      return res.status(400).json({ error: `question must be ${MAX_QUESTION_LEN} characters or fewer.` });
    }
    if (!Array.isArray(tickets)) {
      return res.status(400).json({ error: 'tickets must be an array.' });
    }
    if (tickets.length === 0) {
      return res.status(400).json({ error: 'No ticket data provided. Upload an Excel file and apply your filters first.' });
    }
    if (tickets.length > MAX_TICKETS) {
      return res.status(400).json({ error: `Too many tickets in payload (max ${MAX_TICKETS}).` });
    }

    // ── Call Gemini AI ────────────────────────────────────────────────────
    const answer = await askGemini(question.trim(), tickets);

    return res.json({ answer });

  } catch (err) {
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
  console.log(`[server] Model: Gemini-2.5-Flash`);
});