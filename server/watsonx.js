/**
 * watsonx.js
 *
 * Calls the IBM watsonx.ai text/generation REST endpoint.
 * All configuration is read from environment variables — no values
 * are hardcoded here.
 *
 * Prompt design
 * ─────────────
 * The model receives a strict system instruction that confines it to
 * the ticket data supplied in the request.  It is explicitly told NOT
 * to invent information or answer questions outside that scope.
 *
 * Security notes
 * ──────────────
 * • Bearer token is obtained via the IAM module (never hardcoded).
 * • TLS certificate validation is NOT disabled.
 * • No sensitive data is written to logs.
 */

'use strict';

const fetch = require('node-fetch');
const { getIamToken } = require('./iamToken');

/**
 * Builds the prompt that is sent to the model.
 *
 * @param {string}  userQuestion  - Raw question from the user
 * @param {Array}   ticketData    - Array of filtered incident objects from the frontend
 * @returns {string}
 */
function buildPrompt(userQuestion, ticketData) {
  // Serialise ticket data as compact JSON (no pretty-printing to save tokens)
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
 * Sends a question + ticket data to IBM watsonx.ai and returns the
 * model's text response.
 *
 * @param {string}  userQuestion  - Question typed by the user
 * @param {Array}   ticketData    - Filtered incidents array from the frontend
 * @returns {Promise<string>}     - The model's answer text
 */
async function askWatsonx(userQuestion, ticketData) {
  const region    = process.env.WATSONX_REGION    || 'us-south';
  const projectId = process.env.WATSONX_PROJECT_ID;
  const modelId   = process.env.WATSONX_MODEL_ID  || 'ibm/granite-13b-instruct-v2';

  if (!projectId) {
    throw new Error('WATSONX_PROJECT_ID is not set in environment variables.');
  }

  const endpoint =
    `https://${region}.ml.cloud.ibm.com/ml/v1/text/generation?version=2023-05-29`;

  const iamToken = await getIamToken();

  const requestBody = {
    model_id: modelId,
    project_id: projectId,
    input: buildPrompt(userQuestion, ticketData),
    parameters: {
      decoding_method: 'greedy',
      max_new_tokens: 600,
      min_new_tokens: 1,
      stop_sequences: [],
      repetition_penalty: 1.1
    }
  };

  const response = await fetch(endpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'Authorization': `Bearer ${iamToken}`
    },
    body: JSON.stringify(requestBody)
  });

  if (!response.ok) {
    // Return HTTP status only — never expose internal details to the client
    throw new Error(`watsonx.ai request failed — HTTP ${response.status}`);
  }

  const data = await response.json();

  const generatedText =
    data?.results?.[0]?.generated_text?.trim();

  if (!generatedText) {
    throw new Error('watsonx.ai returned an empty response.');
  }

  return generatedText;
}

module.exports = { askWatsonx };
