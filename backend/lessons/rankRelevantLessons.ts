import type { SupabaseClient } from '@supabase/supabase-js';
import { cosineSimilarity, parseEmbedding } from '../embeddings/cosineSimilarity';
import type { Id } from './types';

const HLT_THRESHOLD = 0.35;
const WEIGHT_PROJECT_TYPE = 0.5;
const WEIGHT_METADATA = 0.3;
const WEIGHT_PROJECT_PARAMS = 0.2;
const MAX_RESULTS = 50;

export interface RankMatchInput {
  number?: number | string;
  name?: string;
  entry?: string;
}

export interface RankRelevantLessonsRequest {
  lessonId: Id;
  organizationId: Id;
  projectId: Id;
  matches?: RankMatchInput[];
  projectDetailsNumberToIdMap?: Record<string, string | number>;
}

export interface RankRelevantLessonResult {
  lessonId: Id;
  projectId: Id | null;
  title: string | null;
  highLevelTitle: string | null;
  category: string | null;
  highLevelSimilarity: number;
  score: number;
  components: {
    projectType: number;
    metadata: number;
    projectParams: number;
  };
}

export interface RankRelevantLessonsResponse {
  matches: RankMatchInput[];
  results: RankRelevantLessonResult[];
}

function averageMaxCosine(
  queryEmbeddings: number[][],
  candidateEmbeddings: number[][],
): number {
  if (!queryEmbeddings.length || !candidateEmbeddings.length) return 0;
  let sum = 0;
  for (const q of queryEmbeddings) {
    let best = 0;
    for (const c of candidateEmbeddings) {
      const sim = cosineSimilarity(q, c);
      if (sim > best) best = sim;
    }
    sum += best;
  }
  return sum / queryEmbeddings.length;
}

function chunkIds<T>(ids: T[], size: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < ids.length; i += size) {
    out.push(ids.slice(i, i + size));
  }
  return out;
}

/**
 * Rank completed org lessons relevant to the current lesson using
 * high_level shortlist (>= 0.7) and weighted project-type / metadata / params boosts.
 */
export async function rankRelevantLessons(
  supabase: SupabaseClient,
  req: RankRelevantLessonsRequest,
): Promise<RankRelevantLessonsResponse> {
  const { lessonId, organizationId, projectId } = req;
  const matches = Array.isArray(req.matches) ? req.matches : [];
  const numberToIdMap = req.projectDetailsNumberToIdMap || {};

  if (lessonId == null || organizationId == null || projectId == null) {
    throw new Error('lessonId, organizationId, and projectId are required.');
  }

  const { data: currentLesson, error: lessonError } = await supabase
    .from('lessons_learned')
    .select('id, title, high_level_title, category, project_id, high_level_search_embedding')
    .eq('id', lessonId)
    .eq('organization_id', organizationId)
    .maybeSingle();

  if (lessonError) {
    throw new Error(lessonError.message || 'Failed to load current lesson.');
  }
  if (!currentLesson) {
    throw new Error('Lesson not found.');
  }

  const currentHlt = parseEmbedding(currentLesson.high_level_search_embedding);
  if (!currentHlt) {
    throw new Error(
      'Current lesson is missing a high_level_search_embedding. Embed the lesson first.',
    );
  }

  const currentProjectId =
    currentLesson.project_id != null ? currentLesson.project_id : projectId;

  const { data: currentProject, error: projectError } = await supabase
    .from('projects')
    .select('project_id, project_type_id')
    .eq('project_id', currentProjectId)
    .eq('organization_id', organizationId)
    .maybeSingle();

  if (projectError) {
    throw new Error(projectError.message || 'Failed to load current project.');
  }

  let currentProjectTypeEmbedding: number[] | null = null;
  if (currentProject && currentProject.project_type_id != null) {
    const { data: ptRow, error: ptError } = await supabase
      .from('project_type')
      .select('id, search_embedding')
      .eq('id', currentProject.project_type_id)
      .maybeSingle();
    if (ptError) {
      console.error('rankRelevantLessons: project_type load failed:', ptError.message);
    } else {
      currentProjectTypeEmbedding = parseEmbedding(ptRow?.search_embedding);
    }
  }

  const { data: currentMetaRows, error: metaError } = await supabase
    .from('lessons_learned_metadata')
    .select('id, search_embedding')
    .eq('lessons_learned_id', lessonId)
    .eq('organization_id', organizationId)
    .not('search_embedding', 'is', null);

  if (metaError) {
    console.error('rankRelevantLessons: metadata load failed:', metaError.message);
  }

  const currentMetaEmbeddings = (currentMetaRows || [])
    .map((row) => parseEmbedding(row.search_embedding))
    .filter((v): v is number[] => v != null);

  const matchedDetailIds = new Set<string>();
  for (const match of matches) {
    const numKey = String(match?.number ?? '').trim();
    if (!numKey) continue;
    const mapped = numberToIdMap[numKey] ?? numberToIdMap[Number(numKey) as unknown as string];
    if (mapped != null) matchedDetailIds.add(String(mapped));
  }

  let matchedParamEmbeddings: number[][] = [];
  if (matchedDetailIds.size > 0) {
    const ids = Array.from(matchedDetailIds);
    const { data: detailRows, error: detailError } = await supabase
      .from('project_details')
      .select('id, search_embedding')
      .eq('organization_id', organizationId)
      .in('id', ids);

    if (detailError) {
      console.error('rankRelevantLessons: matched details load failed:', detailError.message);
    } else {
      matchedParamEmbeddings = (detailRows || [])
        .map((row) => parseEmbedding(row.search_embedding))
        .filter((v): v is number[] => v != null);
    }
  }

  const { data: candidates, error: candidatesError } = await supabase
    .from('lessons_learned')
    .select(
      'id, title, high_level_title, category, project_id, high_level_search_embedding',
    )
    .eq('organization_id', organizationId)
    .eq('review', 'complete')
    .neq('id', lessonId)
    .not('high_level_search_embedding', 'is', null)
    .limit(2000);

  if (candidatesError) {
    throw new Error(candidatesError.message || 'Failed to load candidate lessons.');
  }

  type ShortItem = {
    id: Id;
    title: string | null;
    highLevelTitle: string | null;
    category: string | null;
    projectId: Id | null;
    highLevelSimilarity: number;
  };

  const shortlisted: ShortItem[] = [];
  for (const row of candidates || []) {
    const emb = parseEmbedding(row.high_level_search_embedding);
    if (!emb) continue;
    const sim = cosineSimilarity(currentHlt, emb);
    if (sim < HLT_THRESHOLD) continue;
    shortlisted.push({
      id: row.id,
      title: row.title != null ? String(row.title) : null,
      highLevelTitle:
        row.high_level_title != null ? String(row.high_level_title) : null,
      category: row.category != null ? String(row.category) : null,
      projectId: row.project_id != null ? row.project_id : null,
      highLevelSimilarity: sim,
    });
  }

  if (shortlisted.length === 0) {
    return { matches, results: [] };
  }

  const projectIds = Array.from(
    new Set(
      shortlisted
        .map((s) => s.projectId)
        .filter((id): id is Id => id != null),
    ),
  );
  const lessonIds = shortlisted.map((s) => s.id);

  const projectTypeIdByProjectId = new Map<string, Id>();
  for (const chunk of chunkIds(projectIds, 200)) {
    const { data: projects, error } = await supabase
      .from('projects')
      .select('project_id, project_type_id')
      .eq('organization_id', organizationId)
      .in('project_id', chunk);
    if (error) {
      console.error('rankRelevantLessons: projects batch failed:', error.message);
      continue;
    }
    for (const p of projects || []) {
      if (p.project_id != null && p.project_type_id != null) {
        projectTypeIdByProjectId.set(String(p.project_id), p.project_type_id);
      }
    }
  }

  const projectTypeIds = Array.from(new Set(projectTypeIdByProjectId.values()));
  const projectTypeEmbeddingById = new Map<string, number[]>();
  for (const chunk of chunkIds(projectTypeIds, 200)) {
    const { data: types, error } = await supabase
      .from('project_type')
      .select('id, search_embedding')
      .in('id', chunk);
    if (error) {
      console.error('rankRelevantLessons: project_type batch failed:', error.message);
      continue;
    }
    for (const t of types || []) {
      const emb = parseEmbedding(t.search_embedding);
      if (t.id != null && emb) projectTypeEmbeddingById.set(String(t.id), emb);
    }
  }

  const metadataEmbeddingsByLessonId = new Map<string, number[][]>();
  for (const chunk of chunkIds(lessonIds, 100)) {
    const { data: metaRows, error } = await supabase
      .from('lessons_learned_metadata')
      .select('lessons_learned_id, search_embedding')
      .eq('organization_id', organizationId)
      .in('lessons_learned_id', chunk)
      .not('search_embedding', 'is', null);
    if (error) {
      console.error('rankRelevantLessons: candidate metadata batch failed:', error.message);
      continue;
    }
    for (const row of metaRows || []) {
      const emb = parseEmbedding(row.search_embedding);
      if (!emb || row.lessons_learned_id == null) continue;
      const key = String(row.lessons_learned_id);
      const list = metadataEmbeddingsByLessonId.get(key) || [];
      list.push(emb);
      metadataEmbeddingsByLessonId.set(key, list);
    }
  }

  const detailEmbeddingsByProjectId = new Map<string, number[][]>();
  if (matchedParamEmbeddings.length > 0 && projectIds.length > 0) {
    for (const chunk of chunkIds(projectIds, 50)) {
      const { data: details, error } = await supabase
        .from('project_details')
        .select('project_id, search_embedding')
        .eq('organization_id', organizationId)
        .in('project_id', chunk)
        .not('search_embedding', 'is', null);
      if (error) {
        console.error(
          'rankRelevantLessons: candidate project_details batch failed:',
          error.message,
        );
        continue;
      }
      for (const row of details || []) {
        const emb = parseEmbedding(row.search_embedding);
        if (!emb || row.project_id == null) continue;
        const key = String(row.project_id);
        const list = detailEmbeddingsByProjectId.get(key) || [];
        list.push(emb);
        detailEmbeddingsByProjectId.set(key, list);
      }
    }
  }

  const results: RankRelevantLessonResult[] = shortlisted.map((item) => {
    let projectTypeScore = 0;
    if (currentProjectTypeEmbedding && item.projectId != null) {
      const typeId = projectTypeIdByProjectId.get(String(item.projectId));
      if (typeId != null) {
        const candTypeEmb = projectTypeEmbeddingById.get(String(typeId));
        if (candTypeEmb) {
          projectTypeScore = cosineSimilarity(currentProjectTypeEmbedding, candTypeEmb);
        }
      }
    }

    const candMeta = metadataEmbeddingsByLessonId.get(String(item.id)) || [];
    const metadataScore = averageMaxCosine(currentMetaEmbeddings, candMeta);

    const candDetails =
      item.projectId != null
        ? detailEmbeddingsByProjectId.get(String(item.projectId)) || []
        : [];
    const projectParamsScore = averageMaxCosine(matchedParamEmbeddings, candDetails);

    const score =
      WEIGHT_PROJECT_TYPE * projectTypeScore +
      WEIGHT_METADATA * metadataScore +
      WEIGHT_PROJECT_PARAMS * projectParamsScore;

    return {
      lessonId: item.id,
      projectId: item.projectId,
      title: item.title,
      highLevelTitle: item.highLevelTitle,
      category: item.category,
      highLevelSimilarity: item.highLevelSimilarity,
      score,
      components: {
        projectType: projectTypeScore,
        metadata: metadataScore,
        projectParams: projectParamsScore,
      },
    };
  });

  results.sort((a, b) => b.score - a.score || b.highLevelSimilarity - a.highLevelSimilarity);

  return {
    matches,
    results: results.slice(0, MAX_RESULTS),
  };
}
