// server.js
// Minimal Express server that exposes /api/search-projects using HF embeddings
// and the match_project_search Postgres function.

const express = require('express');
const dotenv = require('dotenv');
const path = require('path');
const { createClient } = require('@supabase/supabase-js');
const { InferenceClient } = require('@huggingface/inference');

// Load env vars (reuse .env.backfill for now)
dotenv.config({ path: '.env.backfill' });

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
const HF_API_TOKEN = process.env.HF_API_TOKEN;

if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY || !HF_API_TOKEN) {
  console.error('Missing SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, or HF_API_TOKEN for server.js');
  process.exit(1);
}

const supabase = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY);
const hfClient = new InferenceClient(HF_API_TOKEN);

async function getEmbedding(text) {
  const trimmed = (text || '').trim();
  if (!trimmed) {
    // Return a zero vector of 384 dims if empty
    return Array(384).fill(0);
  }

  const raw = await hfClient.featureExtraction({
    model: 'sentence-transformers/all-MiniLM-L6-v2',
    inputs: trimmed,
  });

  const embedding = Array.isArray(raw[0]) ? raw[0] : raw;

  if (!Array.isArray(embedding)) {
    throw new Error(`Unexpected embedding format from HF: ${JSON.stringify(raw)}`);
  }

  return embedding;
}

const app = express();
app.use(express.json());

// Serve static frontend files from frontend/public
app.use(express.static(path.join(__dirname, 'frontend', 'public')));

// Simple health check
app.get('/api/health', (_req, res) => {
  res.json({ ok: true });
});

// POST /api/search-projects
app.post('/api/search-projects', async (req, res) => {
  try {
    const { queryText } = req.body || {};
    if (!queryText || typeof queryText !== 'string' || !queryText.trim()) {
      return res.status(400).json({ error: 'queryText is required' });
    }

    // 1) Embed the query
    const queryEmbedding = await getEmbedding(queryText);

    // 2) Call the match_project_search function in Postgres
    const { data, error } = await supabase.rpc('match_project_search', {
      query_embedding: queryEmbedding,
      query_text: queryText,
    });

    if (error) {
      console.error('match_project_search RPC error:', error);
      return res.status(500).json({ error: 'Search failed' });
    }

    const rows = data || [];

    // Look up project_name and project_description for the matched project_ids
    let projectMetaById = new Map();
    if (rows.length > 0) {
      const ids = Array.from(
        new Set(rows.map((r) => r.project_id).filter((id) => id != null))
      );

      if (ids.length > 0) {
        const { data: projects, error: projError } = await supabase
          .from('projects')
          .select('project_id, project_name, project_description')
          .in('project_id', ids);

        if (projError) {
          console.error('Error loading project metadata for search results:', projError);
        } else if (projects) {
          projectMetaById = new Map(
            projects.map((p) => [p.project_id, p])
          );
        }
      }
    }

    const results = rows.map((row) => {
      const meta = projectMetaById.get(row.project_id) || {};
      return {
        project_id: row.project_id,
        project_name: meta.project_name || null,
        project_description: meta.project_description || null,
        score: row.score,
      };
    });

    res.json({ query: queryText, results });
  } catch (err) {
    console.error('Unexpected error in /api/search-projects:', err);
    res.status(500).json({ error: 'Internal server error' });
  }
});

const PORT = process.env.PORT || 4000;
app.listen(PORT, () => {
  console.log(`Server listening on http://localhost:${PORT}`);
});


