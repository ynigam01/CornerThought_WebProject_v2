// Find Relevant Lessons Learned — sends combined lesson + project details text
// and a system prompt to Meta Llama via OpenRouter.

/**
 * ---------------------------------------------------------------------------
 * SYSTEM PROMPT
 * Paste / edit your system prompt between the backticks below.
 * ---------------------------------------------------------------------------
 */
export const FIND_RELEVANT_SYSTEM_PROMPT = `
You are an analysis tool that takes an incident and determines which project parameters
are most relevant to the incident. You will receive the incident which describes the issue,
causes, impacts, actions taken and recommended, and lessons learned (not all fields will be
present). The incident may also contain notes, some of which may just be a retelling of the
incident or a description of the issue. Next you will receive a list of project details (i.e.
project parameters), one per line, in the format "N. Parameter Name: Parameter Entry", where
each numbered item appears on its own line.

You will need to determine which project parameters are most relevant to the incident. Only
select parameters that are genuinely and substantively relevant — do not select a parameter
just to have something to return, and do not infer or invent any project parameter that is
not explicitly present in the list you were given.

Return ONLY a JSON object with a single key "matches", whose value is an array of objects.
Each object in the array must have exactly these three fields:
- "number": the integer that precedes the parameter on its line
- "name": the exact Parameter Name text as it appears after that number, before the colon
- "entry": the exact Parameter Entry text as it appears after that colon

Copy "name" and "entry" verbatim from the line at "number" — do not paraphrase,
summarize, or reword them. These fields exist so the number can be verified against the
correct source line, so any mismatch between "number" and the "name"/"entry" you copied
will cause your answer to be rejected.

If no parameters are genuinely relevant, return {"matches": []}.

Return raw JSON only. Do not include any explanation, commentary, markdown formatting,
or code fences before or after the JSON object.

Example: given project details
1. Location: New York
2. Owner: Manhattan Transit Authority
3. Contractor: XYZ Construction
4. Budget: $4.2M
and an incident that only concerns a dispute with the contractor, you would return exactly:
{"matches":[{"number":3,"name":"Contractor","entry":"XYZ Construction"}]}
`;

/**
 * Call Meta Llama with the system prompt and the Find Relevant combined text.
 * @param {string} combinedText - Lesson + project details text from the popup
 * @returns {Promise<string>} Raw model response content
 */
export async function findRelevantLessons(combinedText) {
  const trimmed = String(combinedText || '').trim();
  if (!trimmed) {
    throw new Error('No lesson text available to find relevant lessons.');
  }

  const systemPrompt = String(FIND_RELEVANT_SYSTEM_PROMPT || '').trim();
  const messages = [];

  if (systemPrompt) {
    messages.push({
      role: 'system',
      content: systemPrompt,
    });
  }

  messages.push({
    role: 'user',
    content: trimmed,
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
      (data && data.error) || `Find Relevant Lessons Learned failed (${response.status}).`
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
