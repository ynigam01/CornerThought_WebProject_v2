// server.js
// Minimal Express server that exposes /api/search-projects using HF embeddings
// and the match_project_search Postgres function.

const express = require('express');
const dotenv = require('dotenv');
const path = require('path');
const { createClient } = require('@supabase/supabase-js');
const { InferenceClient } = require('@huggingface/inference');
const {
  getTopMetadataTermsForProjectType,
  getMatchingLessonIdsForProjectTypeMetadataTerm,
} = require('./same-metadata-tracker');
const { saveLessons } = require('./dist/lessons/saveLessons');
const { registerDraftLessonRoutes } = require('./dist/lessons/draftRoutes');
const { rankRelevantLessons } = require('./dist/lessons/rankRelevantLessons');
const { rankUpcomingTaskLessons } = require('./dist/lessons/rankUpcomingTaskLessons');
const {
  OPENROUTER_MODEL,
  chatCompletion,
  getApiKey: getOpenRouterApiKey,
} = require('./backend/openrouter/client');

// Load env vars (reuse .env.backfill for now)
dotenv.config({ path: '.env.backfill' });

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
const HF_API_TOKEN = process.env.HF_API_TOKEN;
// Optional until Analyze/Parse is wired; required only by /api/openrouter/test.
const OPENROUTER_API_KEY = process.env.OPENROUTER_API_KEY;

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

/** Format metadata list text for embedding: "[metadata_type]: [metadata]". */
function formatMetadataListEmbedText(metadataType, metadata) {
  const type = metadataType != null ? String(metadataType).trim() : '';
  let metaText = '';
  if (metadata != null && typeof metadata === 'object') {
    try {
      metaText = JSON.stringify(metadata);
    } catch {
      metaText = String(metadata);
    }
  } else if (metadata != null) {
    metaText = String(metadata).trim();
  }

  if (type && metaText) return `${type}: ${metaText}`;
  return type || metaText;
}

const app = express();
// Allow larger bodies because Add Data attachments are sent as base64.
app.use(express.json({ limit: '25mb' }));
const SEARCH_VIEW = 'public_lessons_search';

// Serve static frontend files from frontend/public
app.use(express.static(path.join(__dirname, 'frontend', 'public')));

// Simple health check
app.get('/api/health', (_req, res) => {
  res.json({ ok: true });
});

// GET /api/project-type-metadata-top?projectTypeId=...
app.get('/api/project-type-metadata-top', async (req, res) => {
  try {
    const projectTypeId = String(req.query?.projectTypeId || '').trim();
    // #region agent log
    if (typeof fetch === 'function') {
      fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'debug-run-1',hypothesisId:'H12',location:'server.js:/api/project-type-metadata-top:start',message:'metadata API called',data:{projectTypeId,hasProjectTypeId:!!projectTypeId},timestamp:Date.now()})}).catch(()=>{});
    }
    // #endregion
    if (!projectTypeId) {
      return res.status(400).json({ error: 'projectTypeId is required' });
    }

    const items = await getTopMetadataTermsForProjectType({
      supabase,
      projectTypeId,
      limit: 10,
    });
    // #region agent log
    if (typeof fetch === 'function') {
      fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'debug-run-1',hypothesisId:'H13',location:'server.js:/api/project-type-metadata-top:success',message:'metadata API returning success',data:{projectTypeId,itemCount:Array.isArray(items)?items.length:null},timestamp:Date.now()})}).catch(()=>{});
    }
    // #endregion

    return res.json({ projectTypeId, items });
  } catch (err) {
    // #region agent log
    if (typeof fetch === 'function') {
      fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'debug-run-1',hypothesisId:'H15',location:'server.js:/api/project-type-metadata-top:error',message:'metadata API threw error',data:{errorMessage:err?.message||'unknown',errorCode:err?.code||null},timestamp:Date.now()})}).catch(()=>{});
    }
    // #endregion
    console.error('Unexpected error in /api/project-type-metadata-top:', err);
    return res.status(500).json({ error: 'Unable to load metadata terms' });
  }
});

// GET /api/project-type-metadata-search?projectTypeId=...&term=...
app.get('/api/project-type-metadata-search', async (req, res) => {
  try {
    const projectTypeId = String(req.query?.projectTypeId || '').trim();
    const term = String(req.query?.term || '').trim();
    // #region agent log
    if (typeof fetch === 'function') {
      fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'debug-metadata-click-1',hypothesisId:'H3',location:'server.js:/api/project-type-metadata-search:start',message:'metadata term search API called',data:{projectTypeId,termLength:term.length,hasProjectTypeId:!!projectTypeId},timestamp:Date.now()})}).catch(()=>{});
    }
    // #endregion
    if (!projectTypeId || !term) {
      return res.status(400).json({ error: 'projectTypeId and term are required' });
    }

    const lessonIds = await getMatchingLessonIdsForProjectTypeMetadataTerm({
      supabase,
      projectTypeId,
      term,
    });
    // #region agent log
    if (typeof fetch === 'function') {
      fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'debug-metadata-click-1',hypothesisId:'H4',location:'server.js:/api/project-type-metadata-search:lessonIds',message:'metadata term helper returned lesson ids',data:{projectTypeId,lessonIdsCount:Array.isArray(lessonIds)?lessonIds.length:null},timestamp:Date.now()})}).catch(()=>{});
    }
    // #endregion

    if (!lessonIds.length) {
      return res.json({ projectTypeId, term, results: [] });
    }

    const extendedSelect = `
      metadata_text,
      lesson_title,
      lesson_category,
      lessons_learned_id,
      fpc_id,
      future_project_consideration,
      lessons_learned_cause_id,
      lessons_learned_impact_id,
      project_name,
      project_type,
      industry
    `;
    const legacySelect = `
      metadata_text,
      lesson_title,
      lesson_category,
      future_project_consideration,
      project_name,
      project_type,
      industry
    `;

    let { data, error } = await supabase
      .from(SEARCH_VIEW)
      .select(extendedSelect)
      .in('lessons_learned_id', lessonIds);

    if (error) {
      const message = String(error?.message || '');
      const isMissingColumn =
        message.includes('does not exist') &&
        (message.includes('fpc_id') ||
          message.includes('lessons_learned_cause_id') ||
          message.includes('lessons_learned_impact_id') ||
          message.includes('lessons_learned_id'));
      if (isMissingColumn) {
        const fallback = await supabase
          .from(SEARCH_VIEW)
          .select(legacySelect)
          .in('lessons_learned_id', lessonIds);
        data = fallback.data;
        error = fallback.error;
      }
    }

    if (error) {
      console.error('Error loading metadata term search rows:', error);
      // #region agent log
      if (typeof fetch === 'function') {
        fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'debug-metadata-click-1',hypothesisId:'H5',location:'server.js:/api/project-type-metadata-search:viewError',message:'metadata term search view query failed',data:{errorMessage:error?.message||'unknown',errorCode:error?.code||null},timestamp:Date.now()})}).catch(()=>{});
      }
      // #endregion
      return res.status(500).json({ error: 'Unable to search metadata term results' });
    }

    return res.json({
      projectTypeId,
      term,
      results: Array.isArray(data) ? data : [],
    });
  } catch (err) {
    console.error('Unexpected error in /api/project-type-metadata-search:', err);
    // #region agent log
    if (typeof fetch === 'function') {
      fetch('http://127.0.0.1:7242/ingest/3f684587-b61e-4851-8662-761311dbc082',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({runId:'debug-metadata-click-1',hypothesisId:'H5',location:'server.js:/api/project-type-metadata-search:catch',message:'metadata term search API threw',data:{errorMessage:err?.message||'unknown',errorCode:err?.code||null},timestamp:Date.now()})}).catch(()=>{});
    }
    // #endregion
    return res.status(500).json({ error: 'Unable to search metadata term results' });
  }
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

// POST /api/backfill-project-type-embeddings
// Fills missing project_type.search_embedding for one organization.
app.post('/api/backfill-project-type-embeddings', async (req, res) => {
  try {
    const organizationId = req.body?.organizationId;
    if (organizationId == null || organizationId === '') {
      return res.status(400).json({ error: 'organizationId is required' });
    }

    const { data: rows, error: fetchError } = await supabase
      .from('project_type')
      .select('id, project_type, search_embedding')
      .eq('organization_id', organizationId);

    if (fetchError) {
      console.error('Error fetching project_type rows for embedding backfill:', fetchError);
      return res.status(500).json({ error: 'Failed to load project types' });
    }

    const allRows = rows || [];
    const total = allRows.length;
    let processed = 0;
    let skipped = 0;
    let failed = 0;
    const batchDelayMs = 1100;

    for (const row of allRows) {
      const typeText = String(row.project_type || '').trim();
      if (row.search_embedding != null || !typeText) {
        skipped += 1;
        continue;
      }

      try {
        const embedding = await getEmbedding(typeText);
        const { error: updateError } = await supabase
          .from('project_type')
          .update({ search_embedding: embedding })
          .eq('id', row.id);

        if (updateError) {
          console.error(`Error updating project_type id=${row.id}:`, updateError);
          failed += 1;
        } else {
          processed += 1;
        }
      } catch (err) {
        console.error(`Error embedding project_type id=${row.id}:`, err?.message || err);
        failed += 1;
      }

      await new Promise((resolve) => setTimeout(resolve, batchDelayMs));
    }

    return res.json({ processed, skipped, failed, total });
  } catch (err) {
    console.error('Unexpected error in /api/backfill-project-type-embeddings:', err);
    return res.status(500).json({ error: 'Internal server error' });
  }
});

// POST /api/backfill-project-and-details-embeddings
// Streams NDJSON progress while filling missing embeddings for one project
// (project_description) and its project_details ("parameter_name: parameter_entry").
app.post('/api/backfill-project-and-details-embeddings', async (req, res) => {
  const organizationId = req.body?.organizationId;
  const projectId = req.body?.projectId;
  const batchDelayMs = 1100;

  const writeEvent = (event) => {
    res.write(`${JSON.stringify(event)}\n`);
  };

  try {
    if (organizationId == null || organizationId === '') {
      return res.status(400).json({ error: 'organizationId is required' });
    }
    if (projectId == null || projectId === '') {
      return res.status(400).json({ error: 'projectId is required' });
    }

    const { data: project, error: projectError } = await supabase
      .from('projects')
      .select('project_id, project_description, search_embedding')
      .eq('organization_id', organizationId)
      .eq('project_id', projectId)
      .maybeSingle();

    if (projectError) {
      console.error('Error loading project for embedding backfill:', projectError);
      return res.status(500).json({ error: 'Failed to load project' });
    }
    if (!project) {
      return res.status(404).json({ error: 'Project not found for this organization' });
    }

    const { data: detailRows, error: detailsError } = await supabase
      .from('project_details')
      .select('id, parameter_name, parameter_entry, search_embedding')
      .eq('organization_id', organizationId)
      .eq('project_id', projectId);

    if (detailsError) {
      console.error('Error loading project_details for embedding backfill:', detailsError);
      return res.status(500).json({ error: 'Failed to load project details' });
    }

    const allDetails = detailRows || [];
    const detailsNeedingEmbed = allDetails.filter((row) => {
      if (row.search_embedding != null) return false;
      const name = String(row.parameter_name || '').trim();
      const entry = String(row.parameter_entry || '').trim();
      return Boolean(name || entry);
    });
    const detailTotal = detailsNeedingEmbed.length;

    res.status(200);
    res.setHeader('Content-Type', 'application/x-ndjson; charset=utf-8');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('X-Accel-Buffering', 'no');
    if (typeof res.flushHeaders === 'function') {
      res.flushHeaders();
    }

    writeEvent({ type: 'start', detailTotal });

    let projectStatus = 'skipped';
    const description = String(project.project_description || '').trim();
    if (project.search_embedding != null || !description) {
      projectStatus = 'skipped';
      writeEvent({ type: 'project', status: projectStatus });
    } else {
      try {
        const embedding = await getEmbedding(description);
        const { error: updateError } = await supabase
          .from('projects')
          .update({ search_embedding: embedding })
          .eq('project_id', project.project_id)
          .eq('organization_id', organizationId);

        if (updateError) {
          console.error(`Error updating project_id=${project.project_id}:`, updateError);
          projectStatus = 'failed';
        } else {
          projectStatus = 'updated';
        }
      } catch (err) {
        console.error(`Error embedding project_id=${project.project_id}:`, err?.message || err);
        projectStatus = 'failed';
      }
      writeEvent({ type: 'project', status: projectStatus });
      await new Promise((resolve) => setTimeout(resolve, batchDelayMs));
    }

    let processed = 0;
    let skipped = allDetails.length - detailTotal;
    let failed = 0;

    for (let i = 0; i < detailsNeedingEmbed.length; i += 1) {
      const row = detailsNeedingEmbed[i];
      const name = String(row.parameter_name || '').trim();
      const entry = String(row.parameter_entry || '').trim();
      const text = `${name}: ${entry}`.trim();
      const index = i + 1;
      let status = 'failed';

      try {
        const embedding = await getEmbedding(text);
        const { error: updateError } = await supabase
          .from('project_details')
          .update({ search_embedding: embedding })
          .eq('id', row.id);

        if (updateError) {
          console.error(`Error updating project_details id=${row.id}:`, updateError);
          failed += 1;
          status = 'failed';
        } else {
          processed += 1;
          status = 'updated';
        }
      } catch (err) {
        console.error(`Error embedding project_details id=${row.id}:`, err?.message || err);
        failed += 1;
        status = 'failed';
      }

      writeEvent({ type: 'detail', index, total: detailTotal, status });
      await new Promise((resolve) => setTimeout(resolve, batchDelayMs));
    }

    writeEvent({
      type: 'done',
      projectStatus,
      details: { processed, skipped, failed, total: allDetails.length },
    });
    return res.end();
  } catch (err) {
    console.error('Unexpected error in /api/backfill-project-and-details-embeddings:', err);
    if (res.headersSent) {
      try {
        writeEvent({ type: 'error', message: err?.message || 'Internal server error' });
        res.end();
      } catch (_) {
        // ignore write failures after stream errors
      }
      return;
    }
    return res.status(500).json({ error: 'Internal server error' });
  }
});

// POST /api/backfill-metadata-list-embeddings
// Streams NDJSON progress while filling missing search_embedding for
// lessons_learned_metadata_list rows on one project ("[metadata_type]: [metadata]").
app.post('/api/backfill-metadata-list-embeddings', async (req, res) => {
  const organizationId = req.body?.organizationId;
  const projectId = req.body?.projectId;
  const batchDelayMs = 1100;

  const writeEvent = (event) => {
    res.write(`${JSON.stringify(event)}\n`);
  };

  try {
    if (organizationId == null || organizationId === '') {
      return res.status(400).json({ error: 'organizationId is required' });
    }
    if (projectId == null || projectId === '') {
      return res.status(400).json({ error: 'projectId is required' });
    }

    const { data: rows, error: fetchError } = await supabase
      .from('lessons_learned_metadata_list')
      .select('id, metadata_type, metadata, search_embedding')
      .eq('organization_id', organizationId)
      .eq('project_id', projectId);

    if (fetchError) {
      console.error(
        'Error loading lessons_learned_metadata_list for embedding backfill:',
        fetchError
      );
      return res.status(500).json({ error: 'Failed to load metadata list' });
    }

    const allRows = rows || [];
    const needingEmbed = allRows.filter((row) => {
      if (row.search_embedding != null) return false;
      const text = formatMetadataListEmbedText(row.metadata_type, row.metadata);
      return Boolean(text);
    });
    const embedTotal = needingEmbed.length;

    res.status(200);
    res.setHeader('Content-Type', 'application/x-ndjson; charset=utf-8');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('X-Accel-Buffering', 'no');
    if (typeof res.flushHeaders === 'function') {
      res.flushHeaders();
    }

    writeEvent({ type: 'start', total: embedTotal });

    let processed = 0;
    let skipped = allRows.length - embedTotal;
    let failed = 0;

    for (let i = 0; i < needingEmbed.length; i += 1) {
      const row = needingEmbed[i];
      const text = formatMetadataListEmbedText(row.metadata_type, row.metadata);
      const index = i + 1;
      let status = 'failed';

      try {
        const embedding = await getEmbedding(text);
        const { error: updateError } = await supabase
          .from('lessons_learned_metadata_list')
          .update({ search_embedding: embedding })
          .eq('id', row.id);

        if (updateError) {
          console.error(
            `Error updating lessons_learned_metadata_list id=${row.id}:`,
            updateError
          );
          failed += 1;
          status = 'failed';
        } else {
          processed += 1;
          status = 'updated';
        }
      } catch (err) {
        console.error(
          `Error embedding lessons_learned_metadata_list id=${row.id}:`,
          err?.message || err
        );
        failed += 1;
        status = 'failed';
      }

      writeEvent({ type: 'row', index, total: embedTotal, status });
      await new Promise((resolve) => setTimeout(resolve, batchDelayMs));
    }

    writeEvent({
      type: 'done',
      processed,
      skipped,
      failed,
      total: allRows.length,
    });
    return res.end();
  } catch (err) {
    console.error('Unexpected error in /api/backfill-metadata-list-embeddings:', err);
    if (res.headersSent) {
      try {
        writeEvent({ type: 'error', message: err?.message || 'Internal server error' });
        res.end();
      } catch (_) {
        // ignore write failures after stream errors
      }
      return;
    }
    return res.status(500).json({ error: 'Internal server error' });
  }
});

// Shared handler for saving Add Data lessons learned (draft or for-review).
async function handleSaveLessons(req, res, review) {
  try {
    const body = req.body || {};
    const request = {
      userId: body.userId,
      organizationId: body.organizationId,
      projectId: body.projectId,
      projectTypeId: body.projectTypeId != null ? body.projectTypeId : null,
      review,
      entries: Array.isArray(body.entries) ? body.entries : [],
    };

    if (request.userId == null || request.organizationId == null) {
      return res.status(400).json({ error: 'Missing user information.' });
    }
    if (request.projectId == null) {
      return res.status(400).json({ error: 'Select a project before saving.' });
    }
    if (!request.entries.length) {
      return res.status(400).json({ error: 'No entries to save.' });
    }

    const result = await saveLessons(supabase, request);
    return res.json(result);
  } catch (err) {
    console.error(`Error saving lessons (${review}):`, err);
    return res.status(500).json({ error: err?.message || 'Failed to save lessons.' });
  }
}

// POST /api/lessons/draft - save Add Data entries as a draft
app.post('/api/lessons/draft', (req, res) => handleSaveLessons(req, res, 'draft'));

// POST /api/lessons/submit - save Add Data entries and mark them for review
app.post('/api/lessons/submit', (req, res) => handleSaveLessons(req, res, 'for review'));

// POST /api/openrouter/test - verify OpenRouter + Llama 3.3 70B connectivity
app.post('/api/openrouter/test', async (req, res) => {
  if (!getOpenRouterApiKey()) {
    return res.status(400).json({
      error: 'OPENROUTER_API_KEY is not set. Add it to .env.backfill and restart the server.',
    });
  }

  try {
    const data = await chatCompletion({
      messages: [{ role: 'user', content: 'Reply with the word: ok' }],
    });
    const content =
      data &&
      data.choices &&
      data.choices[0] &&
      data.choices[0].message
        ? data.choices[0].message.content
        : null;
    return res.json({
      ok: true,
      model: OPENROUTER_MODEL,
      content,
    });
  } catch (err) {
    if (err && err.code === 'MISSING_API_KEY') {
      return res.status(400).json({ error: err.message });
    }
    console.error('OpenRouter test failed:', err);
    return res.status(502).json({
      error: err?.message || 'OpenRouter request failed.',
    });
  }
});

// POST /api/openrouter/chat - send system + user messages to Llama via OpenRouter
app.post('/api/openrouter/chat', async (req, res) => {
  if (!getOpenRouterApiKey()) {
    return res.status(400).json({
      error: 'OPENROUTER_API_KEY is not set. Add it to .env.backfill and restart the server.',
    });
  }

  const messages = req.body && Array.isArray(req.body.messages) ? req.body.messages : null;
  if (!messages || messages.length === 0) {
    return res.status(400).json({ error: 'messages must be a non-empty array.' });
  }

  try {
    const data = await chatCompletion({ messages });
    const content =
      data &&
      data.choices &&
      data.choices[0] &&
      data.choices[0].message
        ? data.choices[0].message.content
        : null;
    return res.json({
      ok: true,
      model: OPENROUTER_MODEL,
      content,
    });
  } catch (err) {
    if (err && err.code === 'MISSING_API_KEY') {
      return res.status(400).json({ error: err.message });
    }
    console.error('OpenRouter chat failed:', err);
    return res.status(502).json({
      error: err?.message || 'OpenRouter request failed.',
    });
  }
});

// POST /api/find-relevant-lessons/rank
// Rank completed org lessons using HLT shortlist + weighted project type / metadata / params.
app.post('/api/find-relevant-lessons/rank', async (req, res) => {
  try {
    const body = req.body || {};
    const lessonId = body.lessonId;
    const organizationId = body.organizationId;
    const projectId = body.projectId;
    if (lessonId == null || organizationId == null || projectId == null) {
      return res.status(400).json({
        error: 'lessonId, organizationId, and projectId are required.',
      });
    }

    const result = await rankRelevantLessons(supabase, {
      lessonId,
      organizationId,
      projectId,
      matches: Array.isArray(body.matches) ? body.matches : [],
      projectDetailsNumberToIdMap:
        body.projectDetailsNumberToIdMap && typeof body.projectDetailsNumberToIdMap === 'object'
          ? body.projectDetailsNumberToIdMap
          : {},
    });

    return res.json(result);
  } catch (err) {
    console.error('find-relevant-lessons/rank failed:', err);
    return res.status(500).json({
      error: err?.message || 'Failed to rank relevant lessons.',
    });
  }
});

// POST /api/upcoming-task-lessons/rank
// Rank completed lessons for an upcoming assigned task via metadata-list embedding match.
app.post('/api/upcoming-task-lessons/rank', async (req, res) => {
  try {
    const body = req.body || {};
    const organizationId = body.organizationId;
    const userId = body.userId;
    const metadataListId = body.metadataListId;
    if (organizationId == null || userId == null || metadataListId == null) {
      return res.status(400).json({
        error: 'organizationId, userId, and metadataListId are required.',
      });
    }

    const result = await rankUpcomingTaskLessons(supabase, {
      organizationId,
      userId,
      metadataListId,
    });

    return res.json(result);
  } catch (err) {
    console.error('upcoming-task-lessons/rank failed:', err);
    return res.status(500).json({
      error: err?.message || 'Failed to rank upcoming task lessons.',
    });
  }
});

// Draft lesson editor write endpoints (/api/draft-lessons/...)
registerDraftLessonRoutes(app, supabase);

const PORT = process.env.PORT || 4000;
app.listen(PORT, () => {
  console.log(`Server listening on http://localhost:${PORT}`);
});


