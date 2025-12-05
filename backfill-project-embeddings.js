// backfill-project-embeddings.js
// Fills projects.search_embedding using Hugging Face embeddings
// Model: sentence-transformers/all-MiniLM-L6-v2 (dimension 384)

const dotenv = require('dotenv');
const { createClient } = require('@supabase/supabase-js');
const { InferenceClient } = require('@huggingface/inference');

// Load environment variables from .env.backfill in the project root
dotenv.config({ path: '.env.backfill' });

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
const HF_API_TOKEN = process.env.HF_API_TOKEN;

if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY || !HF_API_TOKEN) {
  console.error('Missing SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, or HF_API_TOKEN in .env.backfill');
  process.exit(1);
}

const supabase = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY);
const hfClient = new InferenceClient(HF_API_TOKEN);

// Hugging Face embedding helper using Inference Providers featureExtraction
async function getEmbedding(text) {
  const raw = await hfClient.featureExtraction({
    model: 'sentence-transformers/all-MiniLM-L6-v2',
    inputs: text,
  });

  // HF typically returns [ [v1, v2, ...] ] for a single sentence
  const embedding = Array.isArray(raw[0]) ? raw[0] : raw;

  if (!Array.isArray(embedding)) {
    throw new Error(`Unexpected embedding format from HF: ${JSON.stringify(raw)}`);
  }

  return embedding;
}

async function main() {
  console.log('Fetching projects and existing embeddings...');

  // 1) Get all project search documents (public projects only, via the view)
  const { data: docs, error: docsError } = await supabase
    .from('project_search_documents')
    .select('project_id, search_text');

  if (docsError) {
    console.error('Error fetching project_search_documents:', docsError);
    process.exit(1);
  }

  // 2) Get current embeddings so we can skip already-filled projects
  const { data: projects, error: projectsError } = await supabase
    .from('projects')
    .select('project_id, search_embedding');

  if (projectsError) {
    console.error('Error fetching projects:', projectsError);
    process.exit(1);
  }

  const embeddingMap = new Map();
  for (const row of projects) {
    embeddingMap.set(row.project_id, row.search_embedding);
  }

  const toProcess = docs.filter((doc) => !embeddingMap.get(doc.project_id));

  console.log(`Total projects in view: ${docs.length}`);
  console.log(`Projects needing embeddings: ${toProcess.length}`);

  const batchDelayMs = 1100; // simple rate limit for HF free tier

  let processed = 0;
  for (const doc of toProcess) {
    try {
      console.log(`Embedding project_id=${doc.project_id} (${++processed}/${toProcess.length})`);

      const embedding = await getEmbedding(doc.search_text);

      const { error: updateError } = await supabase
        .from('projects')
        .update({ search_embedding: embedding })
        .eq('project_id', doc.project_id);

      if (updateError) {
        console.error(`Error updating project_id=${doc.project_id}:`, updateError);
      }
    } catch (err) {
      console.error(`Error processing project_id=${doc.project_id}:`, err.message);
    }

    // Small pause to avoid hitting HF rate limits
    await new Promise((resolve) => setTimeout(resolve, batchDelayMs));
  }

  console.log('Project backfill complete.');
}

main().catch((err) => {
  console.error('Fatal error:', err);
  process.exit(1);
});


