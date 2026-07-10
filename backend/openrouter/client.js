// Minimal OpenRouter chat client for Meta Llama 3.3 70B Instruct.
// Uses the OpenAI-compatible Chat Completions API via native fetch.

const OPENROUTER_API_URL = 'https://openrouter.ai/api/v1/chat/completions';
const OPENROUTER_MODEL = 'meta-llama/llama-3.3-70b-instruct';

function getApiKey() {
  return process.env.OPENROUTER_API_KEY || '';
}

/**
 * Send a chat completion request to OpenRouter.
 * @param {{ messages: Array<{ role: string, content: string }> }} options
 * @returns {Promise<object>} Parsed OpenRouter response JSON
 */
async function chatCompletion({ messages }) {
  const apiKey = getApiKey();
  if (!apiKey) {
    const err = new Error(
      'OPENROUTER_API_KEY is not set. Add it to .env.backfill and restart the server.'
    );
    err.code = 'MISSING_API_KEY';
    throw err;
  }

  if (!Array.isArray(messages) || messages.length === 0) {
    throw new Error('messages must be a non-empty array.');
  }

  const response = await fetch(OPENROUTER_API_URL, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${apiKey}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      model: OPENROUTER_MODEL,
      messages,
    }),
  });

  let data = null;
  try {
    data = await response.json();
  } catch (_) {
    data = null;
  }

  if (!response.ok) {
    const detail =
      (data && (data.error?.message || data.error || data.message)) ||
      `OpenRouter request failed with status ${response.status}`;
    const err = new Error(typeof detail === 'string' ? detail : JSON.stringify(detail));
    err.code = 'OPENROUTER_ERROR';
    err.status = response.status;
    err.body = data;
    throw err;
  }

  return data;
}

module.exports = {
  OPENROUTER_MODEL,
  chatCompletion,
  getApiKey,
};
