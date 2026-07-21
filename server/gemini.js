'use strict';

const { GoogleGenerativeAI } = require('@google/generative-ai');

/**
 * Builds the prompt that is sent to the model.
 */
function buildPrompt(userQuestion, ticketData) {
  const dataJson = JSON.stringify(ticketData);
  return (
    `You are a strict ticket analyst assistant. ` +
    `You MUST answer ONLY based on the ticket data provided below. ` +
    `Do NOT use any external knowledge, do NOT invent ticket numbers, assignees, or dates. ` +
    `If the answer cannot be determined from the provided data, say so clearly.\n\n` +
    `TICKET DATA (JSON):\n${dataJson}\n\n` +
    `USER QUESTION: ${userQuestion}\n\n` +
    `ANSWER:`
  );
}

/**
 * Sends a question + ticket data to Gemini 1.5 Flash and returns the text response.
 */
async function askGemini(userQuestion, ticketData) {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    throw new Error('GEMINI_API_KEY is not set in environment variables.');
  }

  // Initialize Google Gemini SDK
  const genAI = new GoogleGenerativeAI(apiKey);
  const model = genAI.getGenerativeModel({ model: "gemini-2.5-flash" });

  const prompt = buildPrompt(userQuestion, ticketData);

  // Generate response
  const result = await model.generateContent(prompt);
  const response = await result.response;
  const generatedText = response.text();

  if (!generatedText) {
    throw new Error('Gemini returned an empty response.');
  }

  return generatedText.trim();
}

module.exports = { askGemini };