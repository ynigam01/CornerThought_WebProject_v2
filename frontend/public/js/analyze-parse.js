// Analyze and Parse — sends source text + system prompt to Meta Llama via OpenRouter.

/**
 * ---------------------------------------------------------------------------
 * SYSTEM PROMPT
 * Paste / edit your system prompt between the backticks below.
 * ---------------------------------------------------------------------------
 */
export const ANALYZE_PARSE_SYSTEM_PROMPT = `
You are an assistant that parses reports into the following fields:
1.	What happened (high level summary)
2.	What happened (detailed)
3.	Cause(s)
4.	Impact(s)
5.	Action(s) taken
6.	Recommended Action(s)
7.	Lessons Learned

If any of the fields above do not exist in the text of the report, don't include them, but there should always be a what happened (high level summary and detailed) at a minimum. 
What happened is the core of the report (i.e. what went right or what went wrong) without including the causes or consequences. The high level summary is a single statement that should be no more than 5 words (think of it as a title). What happened (detailed) should be one contiguous statement that fleshes out what happened in more detail, again without including the causes or consequences.
Every action taken or recommended, as well as every lesson learned should be linked to a cause or impact (i.e. actions taken to mitigate the impact, recommended actions to reduce further impact or prevent incident from occurring in the future, lessons learned, i.e. knowledge items, that teams on future projects can use to be aware of to handle similar situations).
Do not infer anything. Only extract information that is directly mentioned in the report. For example do not infer what lessons may have been learned by the team or writer of the report. Do not infer what actions are going to be taken in the future, only pull the actions that have been explicitly mentioned. Do not infer what caused any incident or what the impact was.
The output should be in JSON format only (do not include markdown or extra commentary). Here is an example of what a JSON output could look like.
{
"what happened (high level summary)": ["example text"],
"what happened (detailed)": ["example text"], 
 "causes": [
    { "temp_id": "cause_1", "text": " example text " },
    { "temp_id": "cause_2", "text": " example text " }
  ],
  "impacts": [
    { "temp_id": "impact_1", "text": "example text" },
    { "temp_id": "impact_2", "text": "example text " }
  ],
  "actions taken": [
    {"temp_id": "action_1", "text": "example text", "linked_to": "cause_1"
    },
    {"temp_id": "action_2", "text": "example text", "linked_to": "impact_2"
    }
  ],
  "recommended actions": [
    {"temp_id": "action_3", "text": "example text", "linked_to": "cause_1"
    },
    {"temp_id": "action_4", "text": "example text", "linked_to": "impact_2"
    }
  ],
  "lessons": [
    {"temp_id": "lesson_1", "text": "example text",
      "linked_to": "cause_1"
    },
    {"temp_id": "lesson_2", "text": "example text",
      "linked_to": "impact_2"
    }
  ]
}

`;

function stripMarkdownFences(raw) {
  let text = String(raw || '').trim();
  const fenced = text.match(/^```(?:json)?\s*([\s\S]*?)\s*```$/i);
  if (fenced) {
    text = fenced[1].trim();
  }
  return text;
}

function firstTextValue(value) {
  if (Array.isArray(value)) {
    for (const item of value) {
      const text = String(item ?? '').trim();
      if (text) return text;
    }
    return '';
  }
  return String(value ?? '').trim();
}

function normalizeLinkedItems(values) {
  if (!Array.isArray(values)) return [];
  return values
    .map((item) => {
      if (!item || typeof item !== 'object') return null;
      const text = String(item.text || '').trim();
      if (!text) return null;
      return {
        tempId: String(item.temp_id || '').trim(),
        text,
        linkedTo: String(item.linked_to || '').trim(),
      };
    })
    .filter(Boolean);
}

function normalizeCauseImpactItems(values) {
  if (!Array.isArray(values)) return [];
  return values
    .map((item) => {
      if (!item || typeof item !== 'object') return null;
      const text = String(item.text || '').trim();
      if (!text) return null;
      return {
        tempId: String(item.temp_id || '').trim(),
        text,
      };
    })
    .filter(Boolean);
}

/**
 * Parse model content into a normalized Add Data payload.
 * @param {string} content
 * @param {string} originalText
 */
export function parseAnalyzeParseJson(content, originalText) {
  const stripped = stripMarkdownFences(content);
  let parsed;
  try {
    parsed = JSON.parse(stripped);
  } catch (_) {
    throw new Error('Model response was not valid JSON.');
  }

  if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) {
    throw new Error('Model response was not valid JSON.');
  }

  const highLevelTitle = firstTextValue(parsed['what happened (high level summary)']);
  const description = firstTextValue(parsed['what happened (detailed)']);

  if (!highLevelTitle || !description) {
    throw new Error(
      'Model JSON is missing required fields: what happened (high level summary) and what happened (detailed).'
    );
  }

  return {
    highLevelTitle,
    description,
    causes: normalizeCauseImpactItems(parsed.causes),
    impacts: normalizeCauseImpactItems(parsed.impacts),
    actionsTaken: normalizeLinkedItems(parsed['actions taken']),
    recommendedActions: normalizeLinkedItems(parsed['recommended actions']),
    lessons: normalizeLinkedItems(parsed.lessons),
    originalText: String(originalText || '').trim(),
  };
}

/**
 * Call Meta Llama with the system prompt and the user's report text.
 * @param {string} reportText - Text from the Analyze and Parse popup
 * @returns {Promise<object>} Normalized parsed lesson data for Add Data UI
 */
export async function analyzeAndParse(reportText) {
  const trimmed = String(reportText || '').trim();
  if (!trimmed) {
    throw new Error('Please enter text to analyze before continuing.');
  }

  const systemPrompt = String(ANALYZE_PARSE_SYSTEM_PROMPT || '').trim();
  const messages = [];

  if (systemPrompt) {
    messages.push({
      role: 'system',
      content: systemPrompt,
    });
  }

  messages.push({
    role: 'user',
    content: `Report: ${trimmed}`,
  });

  const response = await fetch('/api/openrouter/chat', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ messages }),
  });

  let data = null;
  try {
    data = await response.json();
  } catch (_) {
    data = null;
  }

  if (!response.ok) {
    throw new Error(
      (data && data.error) || `Analyze and Parse failed (${response.status}).`
    );
  }

  const content =
    data && typeof data.content === 'string'
      ? data.content
      : data &&
          data.choices &&
          data.choices[0] &&
          data.choices[0].message
        ? data.choices[0].message.content
        : null;

  if (content == null || String(content).trim() === '') {
    throw new Error('The model returned an empty response.');
  }

  return parseAnalyzeParseJson(content, trimmed);
}
