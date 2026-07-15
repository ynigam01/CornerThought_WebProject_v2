import type { SupabaseClient } from '@supabase/supabase-js';
import type { Id } from './types';
import { embedLessonFields } from './embedLessonFields';

export interface EmbedLessonForFindRequest {
  lessonId: Id;
  organizationId: Id;
}

export interface EmbedLessonForFindResult {
  ok: true;
}

/**
 * Awaits full lesson field embeddings before Find Relevant can run.
 * Separate from complete so Complete stays fire-and-forget / non-blocking.
 */
export async function embedLessonForFind(
  supabase: SupabaseClient,
  req: EmbedLessonForFindRequest,
): Promise<EmbedLessonForFindResult> {
  const { lessonId, organizationId } = req;
  if (lessonId == null || organizationId == null) {
    throw new Error('Missing lesson or organization.');
  }

  await embedLessonFields(supabase, lessonId, organizationId);
  return { ok: true };
}
