import express from 'express';
import dotenv from 'dotenv';
import path from 'path';
import { createClient } from '@supabase/supabase-js';
import { InferenceClient } from '@huggingface/inference';

const {
  getTopMetadataTermsForProjectType,
  getMatchingLessonIdsForProjectTypeMetadataTerm,
} = require('../same-metadata-tracker');

dotenv.config({ path: path.join(__dirname, '..', '.env.backfill') });

const SUPABASE_URL = process.env.SUPABASE_URL;
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY;
const HF_API_TOKEN = process.env.HF_API_TOKEN;

if (!SUPABASE_URL || !SUPABASE_SERVICE_ROLE_KEY || !HF_API_TOKEN) {
  console.error('Missing SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, or HF_API_TOKEN for server.ts');
  process.exit(1);
}

const supabase = createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY);
const hfClient = new InferenceClient(HF_API_TOKEN);
const app = express();
const SEARCH_VIEW = 'public_lessons_search';
const CREATE_PROJECT_USER_TYPES = new Set(['Administrator', 'Company Administrator']);

app.use(express.json());
app.use(express.static(path.join(__dirname, '..', 'frontend', 'public')));

function normalizeRequiredString(value: unknown): string {
  return typeof value === 'string' ? value.trim() : '';
}

function normalizeOptionalString(value: unknown): string | null {
  const trimmed = normalizeRequiredString(value);
  return trimmed || null;
}

function normalizeAsset(value: unknown): string | null {
  const trimmed = normalizeRequiredString(value);
  return !trimmed || trimmed === 'N/A' ? null : trimmed;
}

function parseUserId(value: unknown): number | string | null {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value;
  }
  const raw = normalizeRequiredString(value);
  if (!raw) return null;
  const numeric = Number(raw);
  return Number.isNaN(numeric) ? raw : numeric;
}

async function getUserContext(userIdValue: unknown) {
  const userId = parseUserId(userIdValue);
  if (userId == null) {
    return { error: 'userId is required', status: 400 };
  }

  const { data: user, error } = await supabase
    .from('users')
    .select('id, organizationid, usertype')
    .eq('id', userId)
    .maybeSingle();

  if (error) {
    console.error('Error loading user context:', error);
    return { error: 'Unable to verify user', status: 500 };
  }

  if (!user || user.organizationid == null) {
    return { error: 'User organization could not be found', status: 403 };
  }

  return { user, status: 200 };
}

async function getEmbedding(text: string) {
  const trimmed = (text || '').trim();
  if (!trimmed) {
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

app.get('/api/health', (_req, res) => {
  res.json({ ok: true });
});

app.get('/api/project-types', async (req, res) => {
  try {
    const context = await getUserContext(req.query?.userId);
    if ('error' in context) {
      return res.status(context.status).json({ error: context.error });
    }

    const query = normalizeRequiredString(req.query?.q);
    let request = supabase
      .from('project_type')
      .select('id, project_type')
      .eq('organization_id', context.user.organizationid)
      .order('project_type', { ascending: true });

    if (query) {
      request = request.ilike('project_type', `%${query}%`);
    }

    const { data, error } = await request;

    if (error) {
      console.error('Error loading organization project types:', error);
      return res.status(500).json({ error: 'Unable to load project types' });
    }

    const seen = new Set<string>();
    const projectTypes = (data || [])
      .map((row) => ({
        id: row.id,
        project_type: row.project_type,
      }))
      .filter((row) => {
        const normalized = String(row.project_type || '').trim().toLowerCase();
        if (!normalized || seen.has(normalized)) return false;
        seen.add(normalized);
        return true;
      });

    return res.json({ projectTypes });
  } catch (err) {
    console.error('Unexpected error in /api/project-types:', err);
    return res.status(500).json({ error: 'Unable to load project types' });
  }
});

app.post('/api/projects', async (req, res) => {
  try {
    const context = await getUserContext(req.body?.userId);
    if ('error' in context) {
      return res.status(context.status).json({ error: context.error });
    }

    if (!CREATE_PROJECT_USER_TYPES.has(String(context.user.usertype || '').trim())) {
      return res.status(403).json({ error: 'You do not have permission to create projects' });
    }

    const name = normalizeRequiredString(req.body?.name);
    const type = normalizeRequiredString(req.body?.type);

    if (!name || !type) {
      return res.status(400).json({ error: 'Project name and type are required' });
    }

    const { data: typeRows, error: typeErr } = await supabase
      .from('project_type')
      .select('id')
      .eq('organization_id', context.user.organizationid)
      .ilike('project_type', type)
      .limit(1);

    if (typeErr) {
      console.error('Error looking up organization project_type:', typeErr);
      return res.status(500).json({ error: 'Unable to look up project type' });
    }

    if (!typeRows || typeRows.length === 0) {
      return res.status(400).json({ error: 'Please choose a valid Project Type from the list for your organization' });
    }

    const projectPayload = {
      project_name: name,
      project_type_id: typeRows[0].id,
      project_description: normalizeOptionalString(req.body?.descRaw),
      asset_new_existing: normalizeAsset(req.body?.assetRaw),
      start_date: normalizeOptionalString(req.body?.startRaw),
      end_date: normalizeOptionalString(req.body?.endRaw),
      organization_id: context.user.organizationid,
      search_embedding: null,
    };

    const { data: projectRow, error: projectErr } = await supabase
      .from('projects')
      .insert(projectPayload)
      .select()
      .single();

    if (projectErr) {
      console.error('Error inserting organization project:', projectErr);
      return res.status(500).json({ error: 'Failed to create project' });
    }

    return res.status(201).json({ project: projectRow });
  } catch (err) {
    console.error('Unexpected error in /api/projects:', err);
    return res.status(500).json({ error: 'An unexpected error occurred while creating the project' });
  }
});

app.get('/api/project-type-metadata-top', async (req, res) => {
  try {
    const projectTypeId = String(req.query?.projectTypeId || '').trim();
    if (!projectTypeId) {
      return res.status(400).json({ error: 'projectTypeId is required' });
    }

    const items = await getTopMetadataTermsForProjectType({
      supabase,
      projectTypeId,
      limit: 10,
    });

    return res.json({ projectTypeId, items });
  } catch (err) {
    console.error('Unexpected error in /api/project-type-metadata-top:', err);
    return res.status(500).json({ error: 'Unable to load metadata terms' });
  }
});

app.get('/api/project-type-metadata-search', async (req, res) => {
  try {
    const projectTypeId = String(req.query?.projectTypeId || '').trim();
    const term = String(req.query?.term || '').trim();
    if (!projectTypeId || !term) {
      return res.status(400).json({ error: 'projectTypeId and term are required' });
    }

    const lessonIds = await getMatchingLessonIdsForProjectTypeMetadataTerm({
      supabase,
      projectTypeId,
      term,
    });

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

    let { data, error }: { data: any; error: any } = await supabase
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
      return res.status(500).json({ error: 'Unable to search metadata term results' });
    }

    return res.json({
      projectTypeId,
      term,
      results: Array.isArray(data) ? data : [],
    });
  } catch (err) {
    console.error('Unexpected error in /api/project-type-metadata-search:', err);
    return res.status(500).json({ error: 'Unable to search metadata term results' });
  }
});

app.post('/api/search-projects', async (req, res) => {
  try {
    const { queryText } = req.body || {};
    if (!queryText || typeof queryText !== 'string' || !queryText.trim()) {
      return res.status(400).json({ error: 'queryText is required' });
    }

    const queryEmbedding = await getEmbedding(queryText);
    const { data, error } = await supabase.rpc('match_project_search', {
      query_embedding: queryEmbedding,
      query_text: queryText,
    });

    if (error) {
      console.error('match_project_search RPC error:', error);
      return res.status(500).json({ error: 'Search failed' });
    }

    const rows = data || [];
    let projectMetaById = new Map();

    if (rows.length > 0) {
      const ids = Array.from(
        new Set(rows.map((row) => row.project_id).filter((id) => id != null))
      );

      if (ids.length > 0) {
        const { data: projects, error: projError } = await supabase
          .from('projects')
          .select('project_id, project_name, project_description')
          .in('project_id', ids);

        if (projError) {
          console.error('Error loading project metadata for search results:', projError);
        } else if (projects) {
          projectMetaById = new Map(projects.map((project) => [project.project_id, project]));
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

    return res.json({ query: queryText, results });
  } catch (err) {
    console.error('Unexpected error in /api/search-projects:', err);
    return res.status(500).json({ error: 'Internal server error' });
  }
});

const PORT = process.env.PORT || 4000;
const server = app.listen(PORT, () => {
  console.log(`Server listening on http://localhost:${PORT}`);
});

server.on('error', (err) => {
  console.error('Server failed to start:', err);
});
