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

/**
 * Call Meta Llama with the system prompt and the user's report text.
 * @param {string} reportText - Text from the Analyze and Parse popup
 * @returns {Promise<string>} Model reply content
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

  return String(content);
}
