// check-cosine-batch.js
// Compare one query sentence against many others with all-MiniLM-L6-v2
// Usage: node check-cosine-batch.js   (or npm run check:cosine)

require('dotenv').config({ path: '.env.backfill' });
const { InferenceClient } = require('@huggingface/inference');

const HF_API_TOKEN = process.env.HF_API_TOKEN;
if (!HF_API_TOKEN) {
  console.error('Missing HF_API_TOKEN in .env.backfill');
  process.exit(1);
}

const hf = new InferenceClient(HF_API_TOKEN);
const MODEL = 'sentence-transformers/all-MiniLM-L6-v2';

// -------- edit these --------
const QUERY = 'Toronto';

const CANDIDATES = [
  'Dubai',
  'Lethbridge',
  'Vancouver',
  'New York',
  'Mississauga',
  'Ontario',
  'New Delhi',
  'Ottawa',
  'Pacific Ocean',
  'North America',
  'Canada',
  'Germany',
  'France'
];
// ----------------------------

async function getEmbedding(text) {
  const raw = await hf.featureExtraction({
    model: MODEL,
    inputs: String(text || '').trim(),
  });
  const embedding = Array.isArray(raw[0]) ? raw[0] : raw;
  if (!Array.isArray(embedding)) {
    throw new Error(`Unexpected embedding format: ${JSON.stringify(raw)}`);
  }
  return embedding;
}

function cosineSimilarity(a, b) {
  let dot = 0;
  let normA = 0;
  let normB = 0;
  for (let i = 0; i < a.length; i++) {
    const x = Number(a[i]) || 0;
    const y = Number(b[i]) || 0;
    dot += x * y;
    normA += x * x;
    normB += y * y;
  }
  if (normA <= 0 || normB <= 0) return 0;
  return dot / (Math.sqrt(normA) * Math.sqrt(normB));
}

async function main() {
  console.log(`\nQuery: "${QUERY}"\n`);
  console.log('Embedding query...');
  const queryEmb = await getEmbedding(QUERY);

  const rows = [];
  for (const text of CANDIDATES) {
    process.stdout.write(`Embedding: ${text.slice(0, 60)}...\n`);
    const emb = await getEmbedding(text);
    rows.push({ text, score: cosineSimilarity(queryEmb, emb) });
  }

  rows.sort((a, b) => b.score - a.score);

  console.log('\n=== Cosine similarity (highest first) ===\n');
  for (const row of rows) {
    console.log(`${row.score.toFixed(4)}  |  ${row.text}`);
  }

  for (const t of [0.7, 0.5, 0.35, 0.25]) {
    const n = rows.filter((r) => r.score >= t).length;
    console.log(`\nWould pass threshold ${t}: ${n}/${rows.length}`);
  }
  console.log('');
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
