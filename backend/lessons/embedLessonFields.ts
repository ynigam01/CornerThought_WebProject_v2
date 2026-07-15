import type { SupabaseClient } from '@supabase/supabase-js';
import { getEmbedding } from '../embeddings/getEmbedding';
import type { Id } from './types';

const BATCH_DELAY_MS = 1100;

export interface EmbedLessonFieldsOptions {
  /** When true, skip fields that already have a non-null embedding. */
  onlyMissing?: boolean;
}

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
  existingEmbedding: unknown,
  onlyMissing: boolean,
): Promise<void> {
  const trimmed = String(text || '').trim();
  if (!trimmed) return;
  if (onlyMissing && existingEmbedding != null) return;

  try {
    const embedding = await getEmbedding(trimmed);
    const { error } = await supabase
      .from(table)
      .update({ [column]: embedding })
      .eq('id', id);

    if (error) {
      console.error(`embedLessonFields: failed updating ${label}:`, error.message);
    }
  } catch (err: unknown) {
    const message = err instanceof Error ? err.message : String(err);
    console.error(`embedLessonFields: failed embedding ${label}:`, message);
  }

  await sleep(BATCH_DELAY_MS);
}

/**
 * Embeds searchable text fields for a lesson into their vector columns.
 * By default rewrites present text. With onlyMissing, skips fields that already have embeddings.
 */
export async function embedLessonFields(
  supabase: SupabaseClient,
  lessonId: Id,
  organizationId: Id,
  options: EmbedLessonFieldsOptions = {},
): Promise<void> {
  const onlyMissing = options.onlyMissing === true;

  if (lessonId == null || organizationId == null) {
    throw new Error('Missing lesson or organization.');
  }

  const { data: lesson, error: lessonError } = await supabase
    .from('lessons_learned')
    .select(
      'id, title, high_level_title, title_search_embedding, high_level_search_embedding',
    )
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
  if (
    title &&
    !(onlyMissing && lesson.title_search_embedding != null)
  ) {
    try {
      const embedding = await getEmbedding(title);
      const { error } = await supabase
        .from('lessons_learned')
        .update({ title_search_embedding: embedding })
        .eq('id', lessonId)
        .eq('organization_id', organizationId);
      if (error) {
        console.error('embedLessonFields: title update failed:', error.message);
      }
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : String(err);
      console.error('embedLessonFields: title embed failed:', message);
    }
    await sleep(BATCH_DELAY_MS);
  }

  const highLevelTitle = String(lesson.high_level_title || '').trim();
  if (
    highLevelTitle &&
    !(onlyMissing && lesson.high_level_search_embedding != null)
  ) {
    try {
      const embedding = await getEmbedding(highLevelTitle);
      const { error } = await supabase
        .from('lessons_learned')
        .update({ high_level_search_embedding: embedding })
        .eq('id', lessonId)
        .eq('organization_id', organizationId);
      if (error) {
        console.error(
          'embedLessonFields: high_level_title update failed:',
          error.message,
        );
      }
    } catch (err: unknown) {
      const message = err instanceof Error ? err.message : String(err);
      console.error('embedLessonFields: high_level_title embed failed:', message);
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
        .select('id, cause, search_embedding')
        .eq('lessons_learned_id', scope.lessons_learned_id)
        .eq('organization_id', scope.organization_id),
      supabase
        .from('lessons_learned_impacts')
        .select('id, impact, search_embedding')
        .eq('lessons_learned_id', scope.lessons_learned_id)
        .eq('organization_id', scope.organization_id),
      supabase
        .from('action_items')
        .select('id, action_item, search_embedding')
        .eq('lessons_learned_id', scope.lessons_learned_id)
        .eq('organization_id', scope.organization_id),
      supabase
        .from('future_project_considerations')
        .select('id, fpc, search_embedding')
        .eq('lessons_learned_id', scope.lessons_learned_id)
        .eq('organization_id', scope.organization_id),
      supabase
        .from('lessons_learned_metadata')
        .select('id, metadata_type, metadata, search_embedding')
        .eq('lessons_learned_id', scope.lessons_learned_id)
        .eq('organization_id', scope.organization_id),
    ]);

  for (const res of [causesRes, impactsRes, actionsRes, fpcsRes, metadataRes]) {
    if (res.error) {
      console.error(
        'embedLessonFields: failed loading secondary rows:',
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
      row.search_embedding,
      onlyMissing,
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
      row.search_embedding,
      onlyMissing,
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
      row.search_embedding,
      onlyMissing,
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
      row.search_embedding,
      onlyMissing,
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
      row.search_embedding,
      onlyMissing,
    );
  }
}

/**
 * Returns true if this completed lesson has any embeddable text field with a null embedding.
 */
export async function lessonNeedsMissingEmbeddings(
  supabase: SupabaseClient,
  lessonId: Id,
  organizationId: Id,
  lessonRow?: {
    title?: unknown;
    high_level_title?: unknown;
    title_search_embedding?: unknown;
    high_level_search_embedding?: unknown;
  } | null,
): Promise<boolean> {
  let lesson = lessonRow || null;
  if (!lesson) {
    const { data, error } = await supabase
      .from('lessons_learned')
      .select(
        'title, high_level_title, title_search_embedding, high_level_search_embedding',
      )
      .eq('id', lessonId)
      .eq('organization_id', organizationId)
      .maybeSingle();
    if (error || !data) return false;
    lesson = data;
  }

  const title = String(lesson.title || '').trim();
  if (title && lesson.title_search_embedding == null) return true;

  const highLevelTitle = String(lesson.high_level_title || '').trim();
  if (highLevelTitle && lesson.high_level_search_embedding == null) return true;

  const scope = {
    lessons_learned_id: lessonId,
    organization_id: organizationId,
  };

  const checks: Array<{
    table: string;
    textCol: string;
  }> = [
    { table: 'lessons_learned_causes', textCol: 'cause' },
    { table: 'lessons_learned_impacts', textCol: 'impact' },
    { table: 'action_items', textCol: 'action_item' },
    { table: 'future_project_considerations', textCol: 'fpc' },
    { table: 'lessons_learned_metadata', textCol: 'metadata' },
  ];

  for (const check of checks) {
    const selectCols =
      check.table === 'lessons_learned_metadata'
        ? 'id, metadata_type, metadata, search_embedding'
        : `id, ${check.textCol}, search_embedding`;

    const { data, error } = await supabase
      .from(check.table)
      .select(selectCols)
      .eq('lessons_learned_id', scope.lessons_learned_id)
      .eq('organization_id', scope.organization_id)
      .is('search_embedding', null);

    if (error) {
      console.error(
        `lessonNeedsMissingEmbeddings: ${check.table} query failed:`,
        error.message,
      );
      continue;
    }

    for (const row of data || []) {
      const record = row as unknown as Record<string, unknown>;
      if (check.table === 'lessons_learned_metadata') {
        const text = formatMetadataEmbedText(
          record.metadata_type,
          record.metadata,
        );
        if (text) return true;
      } else {
        const text = String(record[check.textCol] || '').trim();
        if (text) return true;
      }
    }
  }

  return false;
}
