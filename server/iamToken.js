/**
 * iamToken.js
 *
 * Fetches a short-lived IBM Cloud IAM bearer token using the API key
 * stored in IBM_CLOUD_API_KEY.  The token is cached in memory and
 * refreshed automatically 5 minutes before it expires, so every
 * request to watsonx.ai always gets a valid token without an extra
 * round-trip.
 *
 * Security notes
 * ──────────────
 * • The API key is read exclusively from process.env — never hardcoded.
 * • The token is kept only in process memory and is never logged.
 * • TLS certificate validation is NOT disabled (node-fetch default).
 */

'use strict';

const fetch = require('node-fetch');

const IAM_TOKEN_URL = 'https://iam.cloud.ibm.com/identity/token';

// In-memory cache: { accessToken, expiresAtMs }
let tokenCache = null;

/**
 * Returns a valid IAM access token, fetching a fresh one when needed.
 * @returns {Promise<string>} Bearer token string
 */
async function getIamToken() {
  const nowMs = Date.now();
  const refreshBufferMs = 5 * 60 * 1000; // refresh 5 min before expiry

  if (tokenCache && nowMs < tokenCache.expiresAtMs - refreshBufferMs) {
    return tokenCache.accessToken;
  }

  const apiKey = process.env.IBM_CLOUD_API_KEY;
  if (!apiKey) {
    throw new Error('IBM_CLOUD_API_KEY is not set in environment variables.');
  }

  const body = new URLSearchParams({
    grant_type: 'urn:ibm:params:oauth:grant-type:apikey',
    apikey: apiKey
  });

  const response = await fetch(IAM_TOKEN_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: body.toString()
  });

  if (!response.ok) {
    // Surface HTTP status but never expose the API key in the error message
    throw new Error(`IAM token request failed — HTTP ${response.status}`);
  }

  const data = await response.json();

  if (!data.access_token || !data.expires_in) {
    throw new Error('IAM token response is missing expected fields.');
  }

  tokenCache = {
    accessToken: data.access_token,
    // expires_in is in seconds; store absolute expiry as ms
    expiresAtMs: nowMs + data.expires_in * 1000
  };

  return tokenCache.accessToken;
}

module.exports = { getIamToken };
