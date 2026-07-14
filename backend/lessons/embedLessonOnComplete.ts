import type { SupabaseClient } from '@supabase/supabase-js';
import { getEmbedding } from '../embeddings/getEmbedding';
import type { Id } from './types';

const BATCH_DELAY_MS = 1100;

function sleep(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function formatMetadataEmbedText(
  metadataType: unknown,
  metadata: unknown,
): string {
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

async function embedAndUpdate(
  supabase: SupabaseClient,
  table: string,
  id: Id,
  column: string,
  text: string,
  label: string,
): Promise<void> {
  const trimmed = String(text || '').trim();
  if (!trimmed) return;

  try {
    const embedding = await getEmbedding(trimmed);
    const { error } = await supabase
      .from(table)
      .update({ [column]: embedding })
      .eq('id', id);

    if (error) {
      console.error(`embedLessonOnComplete: failed updating ${label}:`, error.message);
    }
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    console.error(`embedLessonOnComplete: failed embedding ${label}:`, message);
  }

  await sleep(BATCH_DELAY_MS);
}

/**
 * Embeds all searchable text fields for a completed lesson into their
 * respective vector columns. Safe to fire-and-forget; never throws to callers
 * that catch — individual row failures are logged and skipped.
 */
export async function embedLessonOnComplete(
  supabase: SupabaseClient,
  lessonId: Id,
  organizationId: Id,
): Promise<void> {
  if (lessonId == null || organizationId == null) {
    throw new Error('Missing lesson or organization.');
  }

  const { data: lesson, error: lessonError } = await supabase
    .from('lessons_learned')
    .select('id, title, high_level_title')
    .eq('id', lessonId)
    .eq('organization_id', organizationId)
    .maybeSingle();

  if (lessonError) {
    throw new Error(lessonError.message || 'Failed to load lesson for embedding.');
  }
  if (!lesson) {
    throw new Error('Lesson not found for embedding.');
  }

  const title = String(lesson.title || '').trim();
  if (title) {
    try {
      const embedding = await getEmbedding(title);
      const { error } = await supabase
        .from('lessons_learned')
        .update({ title_search_embedding: embedding })
        .eq('id', lessonId)
        .eq('organization_id', organizationId);
      if (error) {
        console.error('embedLessonOnComplete: title update failed:', error.message);
      }
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : String(err);
      console.error('embedLessonOnComplete: title embed failed:', message);
    }
    await sleep(BATCH_DELAY_MS);
  }

  const highLevelTitle = String(lesson.high_level_title || '').trim();
  if (highLevelTitle) {
    try {
      const embedding = await getEmbedding(highLevelTitle);
      const { error } = await supabase
        .from('lessons_learned')
        .update({ high_level_search_embedding: embedding })
        .eq('id', lessonId)
        .eq('organization_id', organizationId);
      if (error) {
        console.error(
          'embedLessonOnComplete: high_level_title update failed:',
          error.message,
        );
      }
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : String(err);
      console.error('embedLessonOnComplete: high_level_title embed failed:', message);
    }
    await sleep(BATCH_DELAY_MS);
  }

  const scope = {
    lessons_learned_id: lessonId,
    organization_id: organizationId,
  };

  const [causesRes, impactsRes, actionsRes, fpcsRes, metadataRes] =
    await Promise.all([
      supabase
        .from('lessons_learned_causes')
        .select('id, cause')
        .eq('lessons_learned_id', scope.lessons_learned_id)
        .eq('organization_id', scope.organization_id),
      supabase
        .from('lessons_learned_impacts')
        .select('id, impact')
        .eq('lessons_learned_id', scope.lessons_learned_id)
        .eq('organization_id', scope.organization_id),
      supabase
        .from('action_items')
        .select('id, action_item')
        .eq('lessons_learned_id', scope.lessons_learned_id)
        .eq('organization_id', scope.organization_id),
      supabase
        .from('future_project_considerations')
        .select('id, fpc')
        .eq('lessons_learned_id', scope.lessons_learned_id)
        .eq('organization_id', scope.organization_id),
      supabase
        .from('lessons_learned_metadata')
        .select('id, metadata_type, metadata')
        .eq('lessons_learned_id', scope.lessons_learned_id)
        .eq('organization_id', scope.organization_id),
    ]);

  for (const res of [causesRes, impactsRes, actionsRes, fpcsRes, metadataRes]) {
    if (res.error) {
      console.error(
        'embedLessonOnComplete: failed loading secondary rows:',
        res.error.message,
      );
    }
  }

  for (const row of causesRes.data || []) {
    await embedAndUpdate(
      supabase,
      'lessons_learned_causes',
      row.id,
      'search_embedding',
      row.cause,
      `cause id=${row.id}`,
    );
  }

  for (const row of impactsRes.data || []) {
    await embedAndUpdate(
      supabase,
      'lessons_learned_impacts',
      row.id,
      'search_embedding',
      row.impact,
      `impact id=${row.id}`,
    );
  }

  for (const row of actionsRes.data || []) {
    await embedAndUpdate(
      supabase,
      'action_items',
      row.id,
      'search_embedding',
      row.action_item,
      `action_item id=${row.id}`,
    );
  }

  for (const row of fpcsRes.data || []) {
    await embedAndUpdate(
      supabase,
      'future_project_considerations',
      row.id,
      'search_embedding',
      row.fpc,
      `fpc id=${row.id}`,
    );
  }

  for (const row of metadataRes.data || []) {
    const text = formatMetadataEmbedText(row.metadata_type, row.metadata);
    await embedAndUpdate(
      supabase,
      'lessons_learned_metadata',
      row.id,
      'search_embedding',
      text,
      `metadata id=${row.id}`,
    );
  }
}
