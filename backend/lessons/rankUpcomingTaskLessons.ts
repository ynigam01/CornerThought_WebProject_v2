import type { SupabaseClient } from '@supabase/supabase-js';
import { cosineSimilarity, parseEmbedding } from '../embeddings/cosineSimilarity';
import type { Id } from './types';

const METADATA_THRESHOLD = 0.65;
const WEIGHT_PROJECT_TYPE = 0.5;
const WEIGHT_METADATA = 0.3;
const WEIGHT_PROJECT_PARAMS = 0.2;
const MAX_RESULTS = 50;
const META_PAGE_SIZE = 1000;

export interface RankUpcomingTaskLessonsRequest {
  organizationId: Id;
  userId: Id;
  metadataListId: Id;
}

export interface RankUpcomingTaskLessonResult {
  lessonId: Id;
  projectId: Id | null;
  title: string | null;
  highLevelTitle: string | null;
  category: string | null;
  metadataSimilarity: number;
  score: number;
  components: {
    projectType: number;
    metadata: number;
    projectParams: number;
  };
}

export interface RankUpcomingTaskLessonsResponse {
  results: RankUpcomingTaskLessonResult[];
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
 * Rank completed org lessons relevant to an upcoming assigned task
 * using metadata-list ↔ lesson-metadata cosine (>= 0.65) and
 * weighted project-type / metadata / project-params scores.
 */
export async function rankUpcomingTaskLessons(
  supabase: SupabaseClient,
  req: RankUpcomingTaskLessonsRequest,
): Promise<RankUpcomingTaskLessonsResponse> {
  const { organizationId, userId, metadataListId } = req;

  if (organizationId == null || userId == null || metadataListId == null) {
    throw new Error('organizationId, userId, and metadataListId are required.');
  }

  const { data: assignments, error: assignError } = await supabase
    .from('project_team_member_assignments')
    .select('id, lessons_learned_metadata_list_id, assignment_type')
    .eq('organization_id', organizationId)
    .eq('user_id', userId)
    .eq('lessons_learned_metadata_list_id', metadataListId)
    .eq('assignment_type', 'task')
    .limit(1);

  if (assignError) {
    throw new Error(assignError.message || 'Failed to verify task assignment.');
  }
  if (!assignments || assignments.length === 0) {
    throw new Error('Task is not assigned to this user.');
  }

  const { data: listRow, error: listError } = await supabase
    .from('lessons_learned_metadata_list')
    .select('id, project_id, search_embedding')
    .eq('id', metadataListId)
    .eq('organization_id', organizationId)
    .maybeSingle();

  if (listError) {
    throw new Error(listError.message || 'Failed to load task metadata list.');
  }
  if (!listRow) {
    throw new Error('Task metadata list not found.');
  }

  const taskEmbedding = parseEmbedding(listRow.search_embedding);
  if (!taskEmbedding) {
    throw new Error('Task is missing a search_embedding.');
  }

  const sourceProjectId = listRow.project_id;
  if (sourceProjectId == null) {
    throw new Error('Task metadata list has no project_id.');
  }

  const { data: sourceProject, error: projectError } = await supabase
    .from('projects')
    .select('project_id, project_type_id')
    .eq('project_id', sourceProjectId)
    .eq('organization_id', organizationId)
    .maybeSingle();

  if (projectError) {
    throw new Error(projectError.message || 'Failed to load source project.');
  }

  let sourceProjectTypeEmbedding: number[] | null = null;
  if (sourceProject && sourceProject.project_type_id != null) {
    const { data: ptRow, error: ptError } = await supabase
      .from('project_type')
      .select('id, search_embedding')
      .eq('id', sourceProject.project_type_id)
      .maybeSingle();
    if (ptError) {
      console.error('rankUpcomingTaskLessons: project_type load failed:', ptError.message);
    } else {
      sourceProjectTypeEmbedding = parseEmbedding(ptRow?.search_embedding);
    }
  }

  const { data: sourceDetails, error: sourceDetailsError } = await supabase
    .from('project_details')
    .select('id, search_embedding')
    .eq('organization_id', organizationId)
    .eq('project_id', sourceProjectId)
    .not('search_embedding', 'is', null);

  if (sourceDetailsError) {
    console.error(
      'rankUpcomingTaskLessons: source project_details load failed:',
      sourceDetailsError.message,
    );
  }

  const sourceDetailEmbeddings = (sourceDetails || [])
    .map((row) => parseEmbedding(row.search_embedding))
    .filter((v): v is number[] => v != null);

  // Scan org lesson metadata embeddings; keep best cosine per lesson_id.
  const bestMetaByLessonId = new Map<string, number>();
  let from = 0;
  while (true) {
    const { data: metaPage, error: metaError } = await supabase
      .from('lessons_learned_metadata')
      .select('lessons_learned_id, search_embedding')
      .eq('organization_id', organizationId)
      .not('search_embedding', 'is', null)
      .range(from, from + META_PAGE_SIZE - 1);

    if (metaError) {
      throw new Error(metaError.message || 'Failed to load lesson metadata embeddings.');
    }

    const page = metaPage || [];
    for (const row of page) {
      if (row.lessons_learned_id == null) continue;
      const emb = parseEmbedding(row.search_embedding);
      if (!emb) continue;
      const sim = cosineSimilarity(taskEmbedding, emb);
      if (sim < METADATA_THRESHOLD) continue;
      const key = String(row.lessons_learned_id);
      const prev = bestMetaByLessonId.get(key);
      if (prev == null || sim > prev) bestMetaByLessonId.set(key, sim);
    }

    if (page.length < META_PAGE_SIZE) break;
    from += META_PAGE_SIZE;
  }

  if (bestMetaByLessonId.size === 0) {
    return { results: [] };
  }

  const lessonIds = Array.from(bestMetaByLessonId.keys());
  const lessonsById = new Map<
    string,
    {
      id: Id;
      title: string | null;
      highLevelTitle: string | null;
      category: string | null;
      projectId: Id | null;
    }
  >();

  for (const chunk of chunkIds(lessonIds, 100)) {
    const { data: lessons, error: lessonsError } = await supabase
      .from('lessons_learned')
      .select('id, title, high_level_title, category, project_id, review')
      .eq('organization_id', organizationId)
      .eq('review', 'complete')
      .in('id', chunk);

    if (lessonsError) {
      throw new Error(lessonsError.message || 'Failed to load lessons learned.');
    }

    for (const row of lessons || []) {
      if (row.id == null) continue;
      lessonsById.set(String(row.id), {
        id: row.id,
        title: row.title != null ? String(row.title) : null,
        highLevelTitle:
          row.high_level_title != null ? String(row.high_level_title) : null,
        category: row.category != null ? String(row.category) : null,
        projectId: row.project_id != null ? row.project_id : null,
      });
    }
  }

  const projectIds = Array.from(
    new Set(
      Array.from(lessonsById.values())
        .map((l) => l.projectId)
        .filter((id): id is Id => id != null),
    ),
  );

  const projectTypeIdByProjectId = new Map<string, Id>();
  for (const chunk of chunkIds(projectIds, 200)) {
    const { data: projects, error } = await supabase
      .from('projects')
      .select('project_id, project_type_id')
      .eq('organization_id', organizationId)
      .in('project_id', chunk);
    if (error) {
      console.error('rankUpcomingTaskLessons: projects batch failed:', error.message);
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
      console.error('rankUpcomingTaskLessons: project_type batch failed:', error.message);
      continue;
    }
    for (const t of types || []) {
      const emb = parseEmbedding(t.search_embedding);
      if (t.id != null && emb) projectTypeEmbeddingById.set(String(t.id), emb);
    }
  }

  const detailEmbeddingsByProjectId = new Map<string, number[][]>();
  if (sourceDetailEmbeddings.length > 0 && projectIds.length > 0) {
    for (const chunk of chunkIds(projectIds, 50)) {
      const { data: details, error } = await supabase
        .from('project_details')
        .select('project_id, search_embedding')
        .eq('organization_id', organizationId)
        .in('project_id', chunk)
        .not('search_embedding', 'is', null);
      if (error) {
        console.error(
          'rankUpcomingTaskLessons: candidate project_details batch failed:',
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

  const results: RankUpcomingTaskLessonResult[] = [];

  for (const [lessonKey, metadataSimilarity] of bestMetaByLessonId.entries()) {
    const lesson = lessonsById.get(lessonKey);
    if (!lesson) continue;

    let projectTypeScore = 0;
    if (sourceProjectTypeEmbedding && lesson.projectId != null) {
      const typeId = projectTypeIdByProjectId.get(String(lesson.projectId));
      if (typeId != null) {
        const candTypeEmb = projectTypeEmbeddingById.get(String(typeId));
        if (candTypeEmb) {
          projectTypeScore = cosineSimilarity(sourceProjectTypeEmbedding, candTypeEmb);
        }
      }
    }

    const candDetails =
      lesson.projectId != null
        ? detailEmbeddingsByProjectId.get(String(lesson.projectId)) || []
        : [];
    const projectParamsScore = averageMaxCosine(sourceDetailEmbeddings, candDetails);

    const score =
      WEIGHT_PROJECT_TYPE * projectTypeScore +
      WEIGHT_METADATA * metadataSimilarity +
      WEIGHT_PROJECT_PARAMS * projectParamsScore;

    results.push({
      lessonId: lesson.id,
      projectId: lesson.projectId,
      title: lesson.title,
      highLevelTitle: lesson.highLevelTitle,
      category: lesson.category,
      metadataSimilarity,
      score,
      components: {
        projectType: projectTypeScore,
        metadata: metadataSimilarity,
        projectParams: projectParamsScore,
      },
    });
  }

  results.sort(
    (a, b) => b.score - a.score || b.metadataSimilarity - a.metadataSimilarity,
  );

  return { results: results.slice(0, MAX_RESULTS) };
}
