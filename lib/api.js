// Claude API wrapper. Calls Anthropic Messages API directly from the browser.
// Requires the user's API key (stored in localStorage) and the
// `anthropic-dangerous-direct-browser-access: true` header.

const ENDPOINT = "https://api.anthropic.com/v1/messages";
const VERSION = "2023-06-01";

export const MODELS = {
  "claude-haiku-4-5-20251001": {
    label: "Haiku 4.5 — cheapest (≈$0.05–0.10 / student)",
    inputPer1M: 1,
    outputPer1M: 5,
  },
  "claude-sonnet-4-6": {
    label: "Sonnet 4.6 — balanced (≈$0.15–0.30 / student)",
    inputPer1M: 3,
    outputPer1M: 15,
  },
  "claude-opus-4-7": {
    label: "Opus 4.7 — best quality (≈$0.80–1.50 / student)",
    inputPer1M: 15,
    outputPer1M: 75,
  },
};

export const DEFAULT_MODEL = "claude-sonnet-4-6";

export async function callClaude({
  apiKey,
  model = DEFAULT_MODEL,
  system,
  messages,
  maxTokens = 4096,
  thinking = false,
  signal,
}) {
  if (!apiKey) throw new Error("Missing API key");
  const body = {
    model,
    max_tokens: maxTokens,
    messages,
  };
  if (system) body.system = system;
  if (thinking) body.thinking = { type: "enabled", budget_tokens: 2048 };

  let lastErr;
  for (let attempt = 0; attempt < 4; attempt++) {
    try {
      const resp = await fetch(ENDPOINT, {
        method: "POST",
        signal,
        headers: {
          "Content-Type": "application/json",
          "x-api-key": apiKey,
          "anthropic-version": VERSION,
          "anthropic-dangerous-direct-browser-access": "true",
        },
        body: JSON.stringify(body),
      });
      if (resp.status === 429 || resp.status === 529 || resp.status >= 500) {
        const wait = 800 * Math.pow(2, attempt);
        await sleep(wait);
        continue;
      }
      if (!resp.ok) {
        const text = await resp.text();
        throw new Error(`API ${resp.status}: ${text.slice(0, 500)}`);
      }
      return await resp.json();
    } catch (err) {
      lastErr = err;
      if (err.name === "AbortError") throw err;
      await sleep(500 * (attempt + 1));
    }
  }
  throw lastErr ?? new Error("Claude API request failed after retries");
}

// Extracts the first text block from a response.
export function textOf(response) {
  const block = response.content?.find((b) => b.type === "text");
  return block?.text ?? "";
}

// Strict JSON extraction — Claude sometimes wraps JSON in markdown fences.
export function parseJsonResponse(text) {
  if (!text) throw new Error("empty response");
  // try fenced ```json ... ```
  const fenced = /```(?:json)?\s*([\s\S]*?)```/i.exec(text);
  const candidate = fenced ? fenced[1] : text;
  // grab first balanced { ... } block
  const start = candidate.indexOf("{");
  if (start < 0) throw new Error("no JSON object in response: " + text.slice(0, 200));
  let depth = 0;
  let end = -1;
  for (let i = start; i < candidate.length; i++) {
    const ch = candidate[i];
    if (ch === "{") depth++;
    else if (ch === "}") {
      depth--;
      if (depth === 0) { end = i; break; }
    }
  }
  if (end < 0) throw new Error("unterminated JSON in response");
  const json = candidate.slice(start, end + 1);
  return JSON.parse(json);
}

function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}
