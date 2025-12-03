// test-project-search.js
// Quick script to test semantic project search from Node

const dotenv = require('dotenv');
const { createClient } = require('@supabase/supabase-js');
const { InferenceClient } = require('@huggingface/inference');

// Load the same env file as the backfill script
dotenv.config({ path: '.env.backfill' });

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
const HF_API_TOKEN = process.env.HF_API_TOKEN;

if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY || !HF_API_TOKEN) {
  console.error('Missing SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, or HF_API_TOKEN');
  process.exit(1);
}

const supabase = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY);
const hfClient = new InferenceClient(HF_API_TOKEN);

// Same embedding helper as the backfill script
async function getEmbedding(text) {
  const raw = await hfClient.featureExtraction({
    model: 'sentence-transformers/all-MiniLM-L6-v2',
    inputs: text,
  });

  const embedding = Array.isArray(raw[0]) ? raw[0] : raw;

  if (!Array.isArray(embedding)) {
    throw new Error(`Unexpected embedding format: ${JSON.stringify(raw)}`);
  }

  return embedding;
}

async function runTestSearch(queryText) {
  console.log(`\n=== Testing project search for: "${queryText}" ===`);

  // 1) Embed the query
  const queryEmbedding = await getEmbedding(queryText);

  // 2) Call the match_project_search function in Postgres
  const { data, error } = await supabase.rpc('match_project_search', {
    query_embedding: queryEmbedding,
    query_text: queryText,
  });

  if (error) {
    console.error('RPC error:', error);
    process.exit(1);
  }

  if (!data || data.length === 0) {
    console.log('No matches found.');
    return;
  }

  // 3) Print results
  data.forEach((row, index) => {
    console.log(`\n#${index + 1} project_id=${row.project_id}`);
    console.log(`  score:      ${row.score.toFixed(3)} (similarity: ${row.similarity.toFixed(3)})`);
    console.log(`  parameter match: ${row.has_parameter_match}, text match: ${row.has_text_match}`);
    console.log('  snippet:');
    console.log('  ' + row.search_text.slice(0, 200).replace(/\s+/g, ' ') + '...');
  });
}

// Join all command-line arguments after the script name into one query string
const query = process.argv.slice(2).join(' ') || 'pipeline';

runTestSearch(query).catch((err) => {
  console.error('Fatal error:', err);
  process.exit(1);
});


